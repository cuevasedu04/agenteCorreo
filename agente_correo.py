"""
Agente de correo ANAM -- envio de constancias por Outlook COM.

Corre LOCAL en la computadora de cada usuario, NO en el servidor central.
El sistema web (pantalla "Generador masivo") le habla directo por
localhost para pedirle que envie un correo; este agente es el UNICO lugar
que de verdad toca Outlook. El backend central (Django) nunca ve el PDF ni
intenta enviar nada el mismo -- solo recibe, al final, un aviso de si
salio bien o mal, para la bitacora de auditoria (ver
credencializacionBack/credencializacion/views.py, CorreoConstanciaEnviarView).

Por que un programa aparte y no parte de Django: Outlook COM (pywin32) solo
puede controlar el Outlook que esta abierto en la MISMA maquina donde corre
el proceso. Si el backend central viviera en una sola maquina, solo esa
maquina podria enviar -- sin importar desde que computadora entrara cada
usuario al navegador. Con este agente, cada quien envia con su propio
Outlook, en su propia maquina, con los buzones que YA tiene agregados a su
perfil (nadie mas necesita saber cuales son).

Requiere:
  - Windows, con Outlook instalado y abierto (sesion iniciada) durante todo
    el envio.
  - pip install -r requirements.txt

Uso:
  python agente_correo.py
  (deja la ventana abierta mientras se hacen envios desde el sistema web)

Empaquetado como .exe (opcional, para no pedirle a cada usuario que instale
Python): ver README.md.
"""
import os
import re
import sys
import tempfile
import threading

from flask import Flask, jsonify, request
from flask_cors import CORS

PUERTO = 8877

# Version que reporta /salud. El sistema web la compara contra
# VERSION_AGENTE_SIN_ADJUNTO (agente-correo.service.ts) para saber si este
# agente puede mandar correos SIN adjunto.
#   1 -- solo constancias (PDF obligatorio). Los agentes viejos no reportaban
#        version; el sistema los trata como 1.
#   2 -- el PDF es opcional: enlaces al cuestionario de un curso y avisos a
#        quienes no aprobaron van como correo de puro texto.
#   3 -- estabilidad: COM inicializado por hilo, envios serializados, CORS de
#        red privada (Chrome/Edge) y busqueda de buzon exacta antes que parcial.
#   4 -- imagenes en linea (CID) dentro del cuerpo, p. ej. el QR del cuestionario.
#   5 -- todo correo va dentro del fondo institucional (fondo_correo/), con el
#        texto en el recuadro de la imagen.
VERSION = 5

app = Flask(__name__)
# Sin restriccion de origen ni de headers: el sistema web puede estar
# servido desde un dominio distinto a localhost (el servidor central), y su
# HttpClient manda un header 'authorization' en TODAS sus peticiones (ver
# TokenInterceptor del frontend) -- sin allow_headers='*' el navegador
# bloquearia el preflight antes de que la peticion llegara aqui. No es un
# riesgo de seguridad relevante: el agente solo escucha en 127.0.0.1, asi
# que nadie fuera de ESTA maquina puede llegar a el sin importar el CORS.
# allow_private_network: Chrome/Edge (Private Network Access) exigen
# 'Access-Control-Allow-Private-Network: true' en el preflight cuando una
# pagina servida desde otro dominio llama a localhost.
CORS(app, origins='*', allow_headers='*', allow_private_network=True)

# Outlook COM no tolera bien llamadas concurrentes desde varios hilos; el
# servidor de Flask atiende cada peticion en un hilo distinto, asi que los
# envios se serializan.
_candado_outlook = threading.Lock()



# Piezas de fondo_correo/original.png (las genera generar_fondo.py). En el
# .exe de PyInstaller los datos viven en sys._MEIPASS, no junto al script.
_CARPETA_FONDO = os.path.join(
    getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__))), 'fondo_correo')
_PIEZAS_FONDO = ('fondo_arriba', 'fondo_izquierda', 'fondo_abajo')


def _envolver_en_fondo(cuerpo_html):
    """
    Pone el cuerpo del correo dentro del recuadro de la imagen institucional.
    Regresa (html, [(ruta, cid), ...]) con las piezas a adjuntar en linea; si
    faltan las piezas, regresa el cuerpo tal cual para no dejar de enviar.

    Va con tablas e <img> en vez de background-image porque Outlook de
    escritorio no muestra fondos CSS. El recuadro es una celda que crece con
    el texto; la foto de la izquierda se alinea abajo para seguir pegada a
    las ondas de la pieza inferior aunque el texto sea largo.
    """
    rutas = [(os.path.join(_CARPETA_FONDO, f'{cid}.jpg'), cid) for cid in _PIEZAS_FONDO]
    if not all(os.path.isfile(ruta) for ruta, _ in rutas):
        return cuerpo_html, []

    # Una plantilla escrita como documento completo trae su propio <body>.
    cuerpo = re.search(r'<body[^>]*>(.*)</body>', cuerpo_html, re.S | re.I)
    contenido = cuerpo.group(1) if cuerpo else cuerpo_html

    html = (
        '<html><head><meta charset="utf-8"></head>'
        '<body style="margin:0;padding:0;background:#ffffff;">'
        '<table role="presentation" width="800" cellpadding="0" cellspacing="0" border="0" '
        'style="width:800px;border-collapse:collapse;background:#eaf5f3;">'
        '<tr><td colspan="3" style="padding:0;line-height:0;font-size:0;">'
        '<img src="cid:fondo_arriba" width="800" height="105" alt="Sistema de Constancias de Evaluaciones" '
        'style="display:block;border:0;width:800px;height:105px;"></td></tr>'
        '<tr>'
        '<td width="348" valign="bottom" bgcolor="#e8f5f4" '
        'style="width:348px;padding:0;line-height:0;font-size:0;vertical-align:bottom;background:#e8f5f4;">'
        '<img src="cid:fondo_izquierda" width="348" height="338" alt="" '
        'style="display:block;border:0;width:348px;height:338px;"></td>'
        '<td width="427" height="338" valign="top" bgcolor="#f5fbfa" '
        'style="width:427px;height:338px;vertical-align:top;background:#f5fbfa;'
        'border:1px solid #3b8e80;border-radius:14px;padding:20px 24px;'
        'font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#222222;line-height:1.5;">'
        f'{contenido}</td>'
        '<td width="25" bgcolor="#eef7f5" style="width:25px;padding:0;background:#eef7f5;">&nbsp;</td>'
        '</tr>'
        '<tr><td colspan="3" style="padding:0;line-height:0;font-size:0;">'
        '<img src="cid:fondo_abajo" width="800" height="90" alt="" '
        'style="display:block;border:0;width:800px;height:90px;"></td></tr>'
        '</table></body></html>'
    )
    return html, rutas


class ErrorEnvioCorreo(Exception):
    """Outlook no disponible, buzon no encontrado, o Exchange rechazo el envio."""


@app.route('/salud', methods=['GET'])
def salud():
    """El frontend lo consulta antes de enviar, para avisar con un mensaje claro si el agente no esta corriendo en vez de fallar con un error de red generico."""
    return jsonify({'status': 'ok', 'version': VERSION})


@app.route('/enviar', methods=['POST'])
def enviar():
    destinatario = (request.form.get('destinatario') or '').strip()
    asunto = request.form.get('asunto') or ''
    cuerpo_html = request.form.get('cuerpo_html') or ''
    cuenta_remitente = (request.form.get('cuenta_remitente') or '').strip()
    archivo_pdf = request.files.get('pdf')
    # Version 4: imagenes que el cuerpo muestra con <img src="cid:NOMBRE">.
    # Cada archivo se llama NOMBRE.png; el nombre sin extension es su CID.
    imagenes_inline = request.files.getlist('inline')

    if not destinatario:
        return jsonify({'status': 'error', 'mensaje': 'Falta el destinatario.'}), 400

    # El PDF es opcional (version 2): las constancias lo traen, los enlaces
    # al cuestionario y los avisos a reprobados no.
    with tempfile.TemporaryDirectory() as carpeta_tmp:
        ruta_pdf = None
        if archivo_pdf:
            # basename: el nombre lo manda el navegador y no debe poder
            # escribir fuera de la carpeta temporal.
            nombre = os.path.basename(archivo_pdf.filename or '') or 'Constancia.pdf'
            ruta_pdf = os.path.join(carpeta_tmp, nombre)
            archivo_pdf.save(ruta_pdf)

        rutas_inline = []
        for imagen in imagenes_inline:
            nombre = os.path.basename(imagen.filename or '')
            if not nombre:
                continue
            ruta = os.path.join(carpeta_tmp, nombre)
            imagen.save(ruta)
            rutas_inline.append((ruta, os.path.splitext(nombre)[0]))

        try:
            with _candado_outlook:
                _enviar_por_outlook(destinatario, asunto, cuerpo_html, cuenta_remitente, ruta_pdf, rutas_inline)
        except ErrorEnvioCorreo as exc:
            return jsonify({'status': 'error', 'mensaje': str(exc)}), 502

    return jsonify({'status': 'success'})


def _enviar_por_outlook(destinatario, asunto, cuerpo_html, cuenta_remitente, ruta_pdf=None, rutas_inline=()):
    """
    Mismo patron que el script existente de RH para recibos de nomina
    (recibos_nomina.py): buzon compartido via SentOnBehalfOfName, adjunto
    por ruta de archivo en disco, Outlook/Exchange gestionan la entrega de
    principio a fin. Este agente nunca guarda ni usa contrasenas de correo.
    """
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise ErrorEnvioCorreo(
            'Este agente debe correr en Windows con pywin32 instalado '
            '(pip install pywin32).'
        ) from exc

    # Cada peticion corre en un hilo del servidor que no es el principal; COM
    # debe inicializarse en ESE hilo o Dispatch falla con "CoInitialize has
    # not been called".
    pythoncom.CoInitialize()
    try:
        _enviar_con_com(win32com.client, destinatario, asunto, cuerpo_html, cuenta_remitente, ruta_pdf, rutas_inline)
    finally:
        pythoncom.CoUninitialize()


def _enviar_con_com(win32com_client, destinatario, asunto, cuerpo_html, cuenta_remitente, ruta_pdf, rutas_inline=()):
    try:
        outlook = win32com_client.Dispatch('Outlook.Application')
    except Exception as exc:
        raise ErrorEnvioCorreo(
            'No se pudo conectar con Outlook. Verifica que esté instalado y '
            'abierto en esta máquina, con la sesión iniciada, durante todo '
            'el envío.'
        ) from exc

    carpeta_enviados = None
    if cuenta_remitente:
        carpeta_enviados = _buscar_carpeta_buzon(outlook, cuenta_remitente)
        if not carpeta_enviados:
            raise ErrorEnvioCorreo(
                f"No se encontró la cuenta «{cuenta_remitente}» en Outlook. "
                'Verifica que el buzón esté agregado a tu perfil de Outlook y '
                'que el nombre o correo coincida exactamente.'
            )

    try:
        mail = outlook.CreateItem(0)  # olMailItem
        mail.To = destinatario
        mail.Subject = asunto
        if cuenta_remitente:
            mail.SentOnBehalfOfName = cuenta_remitente
        cuerpo_html, rutas_fondo = _envolver_en_fondo(cuerpo_html)
        mail.HTMLBody = cuerpo_html
        if ruta_pdf:
            mail.Attachments.Add(Source=ruta_pdf)
        for ruta, cid in [*rutas_inline, *rutas_fondo]:
            # PR_ATTACH_CONTENT_ID: con el, Outlook muestra el adjunto dentro
            # del cuerpo (donde va <img src="cid:...">) en vez de como archivo.
            adjunto = mail.Attachments.Add(Source=ruta)
            adjunto.PropertyAccessor.SetProperty(
                'http://schemas.microsoft.com/mapi/proptag/0x3712001F', cid)
        if carpeta_enviados:
            mail.SaveSentMessageFolder = carpeta_enviados
        mail.Send()
    except Exception as exc:
        raise ErrorEnvioCorreo(f'Outlook rechazó el envío: {exc}') from exc


def _buscar_carpeta_buzon(outlook, cuenta: str):
    """
    Busca `cuenta` (DisplayName o correo del buzón) entre los Stores ya
    agregados al perfil de Outlook abierto y regresa su carpeta de
    Elementos enviados. None si no está agregado.
    """
    termino = cuenta.strip().lower()
    stores = list(outlook.Session.Stores)
    # Primero coincidencia exacta; solo si no hay, coincidencia parcial. Asi
    # "rh" no se queda con el primer buzon cuyo nombre contenga "rh" cuando
    # existe uno llamado exactamente "rh".
    exactos = [s for s in stores if (s.DisplayName or '').strip().lower() == termino]
    parciales = [s for s in stores if termino in (s.DisplayName or '').strip().lower()]
    for store in exactos or parciales:
        try:
            return store.GetDefaultFolder(5)  # olFolderSentMail
        except Exception:
            continue
    return None


if __name__ == '__main__':
    print(f'Agente de correo ANAM v{VERSION} escuchando en http://127.0.0.1:{PUERTO}')
    print('Deja esta ventana abierta mientras envías constancias desde el sistema web.')
    print('Ciérrala cuando termines de enviar.')
    app.run(host='127.0.0.1', port=PUERTO, threaded=True)
