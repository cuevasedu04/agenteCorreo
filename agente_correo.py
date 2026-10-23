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
VERSION = 3

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

        try:
            with _candado_outlook:
                _enviar_por_outlook(destinatario, asunto, cuerpo_html, cuenta_remitente, ruta_pdf)
        except ErrorEnvioCorreo as exc:
            return jsonify({'status': 'error', 'mensaje': str(exc)}), 502

    return jsonify({'status': 'success'})


def _enviar_por_outlook(destinatario, asunto, cuerpo_html, cuenta_remitente, ruta_pdf=None):
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
        _enviar_con_com(win32com.client, destinatario, asunto, cuerpo_html, cuenta_remitente, ruta_pdf)
    finally:
        pythoncom.CoUninitialize()


def _enviar_con_com(win32com_client, destinatario, asunto, cuerpo_html, cuenta_remitente, ruta_pdf):
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
        mail.HTMLBody = cuerpo_html
        if ruta_pdf:
            mail.Attachments.Add(Source=ruta_pdf)
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
