"""
Corta fondo_correo/original.png (1536x1024) en las piezas que usa
_envolver_en_fondo() de agente_correo.py. Solo hace falta correrlo si se
cambia la imagen original:

    pip install pillow
    python generar_fondo.py

Por que en piezas y no la imagen completa de fondo: el texto de cada correo
mide distinto. La franja del medio es una celda que crece con el texto, con
fondo_medio (las ondas) detras de toda ella y la foto encima. Con texto de
tamano normal las piezas encajan como la imagen original; con texto largo las
ondas se estiran y sobre la foto se ven ondas en vez de un bloque liso.
"""
import os

from PIL import Image, ImageOps

CARPETA = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'fondo_correo')

# Al doble del tamano en que se muestra (800 px de ancho) para que se vea
# nitido en pantallas de alta densidad.
ANCHO, ALTO = 1600, 1066
ALTO_ARRIBA = 210      # entre el titulo (termina en y=207) y la foto (y=227)
ALTO_MEDIO = 676       # alto de la franja del medio
ANCHO_IZQUIERDA = 720  # justo despues del arco dorado (termina en x=686)
FILAS_UNION = 4        # filas de fondo_medio que repiten el pie de la foto


def _ondas_del_medio(imagen, abajo):
    """La franja sin la foto: a la izquierda, las ondas de la derecha en espejo."""
    medio = imagen.crop((0, ALTO_ARRIBA, ANCHO, abajo))
    espejo = imagen.crop((ANCHO_IZQUIERDA, ALTO_ARRIBA, ANCHO_IZQUIERDA * 2, abajo))
    medio.paste(ImageOps.mirror(espejo), (0, 0))
    # Word deja ver 1 px del fondo bajo la foto (el parrafo que exige tras la
    # tabla): las ultimas filas son las de la foto para que no se note esa linea.
    union = imagen.crop((0, abajo - FILAS_UNION, ANCHO_IZQUIERDA, abajo))
    medio.paste(union, (0, ALTO_MEDIO - FILAS_UNION))
    return medio


def main():
    imagen = Image.open(os.path.join(CARPETA, 'original.png')).convert('RGB').resize((ANCHO, ALTO), Image.LANCZOS)
    abajo = ALTO_ARRIBA + ALTO_MEDIO
    piezas = {
        'fondo_arriba': imagen.crop((0, 0, ANCHO, ALTO_ARRIBA)),
        'fondo_izquierda': imagen.crop((0, ALTO_ARRIBA, ANCHO_IZQUIERDA, abajo)),
        'fondo_medio': _ondas_del_medio(imagen, abajo),
        'fondo_abajo': imagen.crop((0, abajo, ANCHO, ALTO)),
    }
    for nombre, pieza in piezas.items():
        pieza.save(os.path.join(CARPETA, f'{nombre}.jpg'), quality=85, optimize=True)


if __name__ == '__main__':
    main()
