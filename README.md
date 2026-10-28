# Agente de correo ANAM

Programa pequeño que corre en la computadora de **cada usuario** que vaya a
enviar constancias por correo desde el sistema web. Es el único lugar que de
verdad toca Outlook (por COM/`win32com`); el sistema web y su backend central
nunca tocan Outlook directamente — solo le piden a este agente que envíe, y
él hace el trabajo con el Outlook que ya está abierto en esa misma máquina.

## Por qué existe

Outlook COM solo puede controlar el Outlook abierto en la **misma
computadora** donde corre el proceso que lo invoca. Si el envío se hiciera
desde el backend central (una sola máquina), todos los correos saldrían
siempre desde ahí, sin importar quién los mande desde el navegador. Con este
agente, cada persona envía con **su propio Outlook**, con los buzones
compartidos que **ya tiene agregados a su perfil** — nadie más necesita saber
cuáles son ni tener acceso a ellos.

## Requisitos

- Windows, con **Microsoft Outlook instalado**.
- Python 3.9 o superior (o usar el `.exe` empaquetado, ver abajo).

## Instalación y uso (con Python)

```
pip install -r requirements.txt
python agente_correo.py
```

Deja la ventana abierta. Va a mostrar:

```
Agente de correo ANAM escuchando en http://127.0.0.1:8877
Deja esta ventana abierta mientras envías constancias desde el sistema web.
```

**Antes de enviar constancias desde "Generador masivo"**, asegúrate de que:

1. Outlook esté abierto, con la sesión iniciada.
2. Este agente esté corriendo (la ventana de la consola abierta).
3. El buzón compartido desde el que quieres enviar ya esté agregado a tu
   perfil de Outlook (mismo nombre o correo que pongas en "Correo
   electrónico" como cuenta remitente).

El sistema web revisa automáticamente si el agente está corriendo antes de
enviar; si no lo encuentra, avisa con un mensaje claro en vez de fallar en
silencio.

## Empaquetarlo como `.exe` (para no pedirle Python a cada usuario)

En una máquina Windows, con las dependencias ya instaladas:

```
pip install pyinstaller
pyinstaller AgenteCorreoANAM.spec
```

El `.spec` ya incluye las piezas del fondo de los correos
(`fondo_correo\fondo_*.jpg`). Si se cambia `fondo_correo\original.png`, hay
que volver a cortarla con `python generar_fondo.py` (requiere `pillow`).

El ejecutable queda en `dist\AgenteCorreoANAM.exe`. Se puede copiar a
cualquier otra computadora Windows y correr directo, sin instalar Python.
Para que se abra solo al iniciar sesión, se puede agregar un acceso directo
a la carpeta de inicio de Windows (`shell:startup`).

## Seguridad

- El agente **solo escucha en `127.0.0.1`** (loopback): nada fuera de esta
  misma computadora puede llegar a él, sin importar la red.
- Nunca guarda ni pide contraseñas de correo: Outlook ya está autenticado
  (fue el usuario quien inició sesión ahí), el agente solo le dice qué
  correo armar y a quién.
