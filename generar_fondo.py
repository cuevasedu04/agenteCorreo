"""
Corta fondo_correo/original.png (1536x1024) en las piezas que usa
_envolver_en_fondo() de agente_correo.py. Solo hace falta correrlo si se
cambia la imagen original:

    pip install pillow
    python generar_fondo.py

Por que en piezas y no la imagen completa de fondo: Outlook de escritorio
(motor de Word) no muestra background-image, y el texto de cada correo mide
distinto. Con piezas, el recuadro es una celda HTML que crece con el texto.
"""
import os

from PIL import Image

CARPETA = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'fondo_correo')

# Al doble del tamano en que se muestra (800 px de ancho) para que se vea
# nitido en pantallas de alta densidad.
ANCHO, ALTO = 1600, 1066
ALTO_ARRIBA = 210      # hasta el borde superior del recuadro
ALTO_MEDIO = 676       # alto del recuadro
ANCHO_IZQUIERDA = 696  # hasta el borde izquierdo del recuadro
FILAS_TITULO = 6       # filas de la pieza izquierda que alcanzan al titulo


def main():
    imagen = Image.open(os.path.join(CARPETA, 'original.png')).convert('RGB').resize((ANCHO, ALTO), Image.LANCZOS)
    abajo = ALTO_ARRIBA + ALTO_MEDIO
    piezas = {
        'fondo_arriba': (0, 0, ANCHO, ALTO_ARRIBA),
        'fondo_izquierda': (0, ALTO_ARRIBA, ANCHO_IZQUIERDA, abajo),
        'fondo_abajo': (0, abajo, ANCHO, ALTO),
    }
    for nombre, caja in piezas.items():
        pieza = imagen.crop(caja)
        if nombre == 'fondo_izquierda':
            # Las primeras filas traen la orilla inferior de "EVALUACIONES".
            # Con texto largo esta pieza baja y esa orilla quedaria suelta a
            # media altura: se tapa repitiendo una fila limpia de abajo.
            limpia = pieza.crop((0, FILAS_TITULO, pieza.width, FILAS_TITULO + 1))
            for y in range(FILAS_TITULO):
                pieza.paste(limpia, (0, y))
        pieza.save(os.path.join(CARPETA, f'{nombre}.jpg'), quality=85, optimize=True)


if __name__ == '__main__':
    main()
