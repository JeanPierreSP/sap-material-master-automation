"""Tema visual: paleta "Tinta y violeta", tipografía y hoja de estilo de Qt.

Es el único archivo con colores. Las pantallas y piezas piden un color por su nombre
(``tema.MARCA``) o un rol de la hoja (``setProperty("rol", "titulo")``), nunca un código.

Contraste (WCAG 2.1): todo texto sobre su fondo pasa AA (4,5:1 o más) y los bordes de
controles y las marcas de la barra de avance pasan 3:1. ``tests/test_ui.py`` lo mide,
así que un cambio de color que lo rompa no pasa las pruebas.
"""

from __future__ import annotations

from PySide6.QtGui import QColor, QFont, QFontDatabase, QPalette
from PySide6.QtWidgets import QApplication, QWidget

# --- Paleta -----------------------------------------------------------------
FONDO = "#F8F7F4"  # papel claro, detrás de todo
SUPERFICIE = "#FFFFFF"  # tarjetas, barra superior
SUPERFICIE_2 = "#F1EFEA"  # filas al pasar el mouse, encabezados de tabla
BORDE = "#E3E0D9"  # bordes decorativos (tarjetas, separadores)
BORDE_CONTROL = "#8C8894"  # bordes de lo que se puede tocar (3:1 sobre blanco)

TINTA = "#1D1B22"  # texto principal
TINTA_2 = "#55525E"  # texto secundario
TINTA_3 = "#6C6975"  # rótulos, horas, ayudas

MARCA = "#5A3FC0"  # el único acento: botón principal, pestaña elegida, foco
MARCA_FUERTE = "#46309C"  # botón principal al pasar el mouse; texto sobre MARCA_SUAVE
MARCA_SUAVE = "#EEEAFB"  # fondo de lo elegido

OK = "#1D7A46"
OK_SUAVE = "#E5F3EA"
AVISO = "#875A00"
AVISO_SUAVE = "#FAF0DA"
ERROR = "#B3242F"
ERROR_SUAVE = "#FBE9EA"
PISTA = "#E6E3EE"  # parte vacía de la barra de avance
VELO = (29, 27, 34, 96)  # RGBA detrás del buscador: la tinta, semitransparente

#: color de texto y de fondo para cada tono que usan las piezas
TONOS: dict[str, tuple[str, str]] = {
    "ok": (OK, OK_SUAVE),
    "aviso": (AVISO, AVISO_SUAVE),
    "error": (ERROR, ERROR_SUAVE),
    "marca": (MARCA_FUERTE, MARCA_SUAVE),
    "neutro": (TINTA_2, SUPERFICIE_2),
}

#: los estados del núcleo, traducidos a tonos
TONO_DE_ESTADO = {
    "ok": "ok",
    "aviso": "aviso",
    "bloquea": "error",
    "hecho": "ok",
    "fallo": "error",
    "pendiente": "neutro",
    "omitida": "neutro",
    "error": "error",
}

RADIO = 10
RADIO_CONTROL = 8
ANCHO_CONTENIDO = 1040  # ancho máximo de la columna central

FAMILIAS = ["Segoe UI", "Inter", "Noto Sans", "DejaVu Sans"]
FAMILIAS_MONO = ["Cascadia Mono", "Consolas", "DejaVu Sans Mono", "monospace"]


def familia_mono() -> str:
    """La primera monoespaciada instalada: la hoja de estilo de Qt acepta una sola familia."""
    instaladas = set(QFontDatabase.families())
    for familia in FAMILIAS_MONO[:-1]:
        if familia in instaladas:
            return familia
    return QFontDatabase.systemFont(QFontDatabase.SystemFont.FixedFont).family()


def fuente(tam: float = 10, peso: QFont.Weight = QFont.Weight.Normal, mono: bool = False) -> QFont:
    f = QFont()
    f.setFamilies(FAMILIAS_MONO if mono else FAMILIAS)
    f.setPointSizeF(tam)
    f.setWeight(peso)
    return f


def aplicar(app: QApplication) -> None:
    app.setFont(fuente())
    paleta = app.palette()
    # El texto de ejemplo de los campos no se puede fijar en la hoja: va por la paleta
    paleta.setColor(QPalette.ColorRole.PlaceholderText, QColor(TINTA_3))
    app.setPalette(paleta)
    app.setStyleSheet(hoja())


def repintar(widget: QWidget) -> None:
    """Vuelve a aplicar la hoja después de cambiar una propiedad usada en un selector."""
    widget.style().unpolish(widget)
    widget.style().polish(widget)
    widget.update()


def hoja() -> str:
    return f"""
QMainWindow, QWidget#raiz {{ background: {FONDO}; }}
QWidget {{ color: {TINTA}; }}
QScrollArea, QScrollArea > QWidget > QWidget#pagina {{ background: {FONDO}; border: none; }}
QToolTip {{ background: {TINTA}; color: {SUPERFICIE}; border: none; padding: 4px 8px; }}

/* Barra superior */
QFrame#barra {{ background: {SUPERFICIE}; border-bottom: 1px solid {BORDE}; }}
QLabel#marca_nombre {{ font-size: 15px; font-weight: 600; }}
QLabel#marca_sub {{ font-size: 11px; color: {TINTA_3}; }}
QPushButton#campo_buscar {{
    background: {SUPERFICIE}; color: {TINTA_3}; border: 1px solid {BORDE_CONTROL};
    border-radius: {RADIO_CONTROL}px; padding: 6px 10px; text-align: left; font-size: 13px;
}}
QPushButton#campo_buscar:hover {{ border-color: {MARCA}; }}
QLabel#tecla {{
    color: {TINTA_3}; background: {SUPERFICIE_2}; border: none; border-radius: 4px; padding: 0 6px; font-size: 11px;
}}

/* Textos */
QLabel[rol="titulo"] {{ font-size: 26px; font-weight: 600; }}
QLabel[rol="subtitulo"] {{ font-size: 14px; color: {TINTA_2}; }}
QLabel[rol="seccion"] {{ font-size: 15px; font-weight: 600; }}
QLabel[rol="fuerte"] {{ font-size: 14px; font-weight: 600; }}
QLabel[rol="cuerpo"] {{ font-size: 13px; color: {TINTA_2}; }}
QLabel[rol="rotulo"] {{ font-size: 12px; color: {TINTA_3}; }}
QLabel[rol="cifra"] {{ font-size: 26px; font-weight: 600; }}
QLabel[rol="avance"] {{ font-size: 30px; font-weight: 600; }}

/* Tarjetas */
QFrame[rol="tarjeta"] {{ background: {SUPERFICIE}; border: 1px solid {BORDE}; border-radius: {RADIO}px; }}
QFrame[rol="tarjeta"][foco="true"] {{ border: 2px solid {MARCA}; }}
QFrame[rol="tarjeta_apagada"] {{ background: transparent; border: 1px dashed {BORDE}; border-radius: {RADIO}px; }}
QFrame[rol="fila"] {{ background: transparent; border: none; border-radius: {RADIO_CONTROL}px; }}
QFrame[rol="fila"]:hover {{ background: {SUPERFICIE_2}; }}
QFrame[rol="separador"] {{ background: {BORDE}; border: none; max-height: 1px; min-height: 1px; }}

/* Botones */
QPushButton {{
    background: {SUPERFICIE}; color: {TINTA}; border: 1px solid {BORDE_CONTROL};
    border-radius: {RADIO_CONTROL}px; padding: 7px 14px; font-size: 13px;
}}
QPushButton:hover {{ background: {SUPERFICIE_2}; }}
QPushButton:focus {{ border: 2px solid {MARCA}; padding: 6px 13px; }}
QPushButton:disabled {{ color: {TINTA_3}; border-color: {BORDE}; background: {SUPERFICIE_2}; }}
QPushButton[variante="principal"] {{
    background: {MARCA}; color: {SUPERFICIE}; border: none; font-size: 15px; font-weight: 600;
    padding: 11px 28px;
}}
QPushButton[variante="principal"]:hover {{ background: {MARCA_FUERTE}; }}
QPushButton[variante="principal"]:focus {{ border: 2px solid {MARCA_FUERTE}; padding: 9px 26px; }}
QPushButton[variante="principal"]:disabled {{ background: {PISTA}; color: {TINTA_3}; }}
QPushButton[variante="detener"] {{
    background: {SUPERFICIE}; color: {ERROR}; border: 1px solid {ERROR}; font-size: 15px; font-weight: 600;
    padding: 10px 27px;
}}
QPushButton[variante="detener"]:hover {{ background: {ERROR_SUAVE}; }}
QPushButton[variante="texto"] {{
    background: transparent; border: none; color: {MARCA_FUERTE}; padding: 4px 6px; font-size: 13px;
}}
QPushButton[variante="texto"]:hover {{ background: {MARCA_SUAVE}; }}
QPushButton[variante="texto"]:focus {{ border: 1px solid {MARCA}; padding: 3px 5px; }}
QPushButton[variante="icono"] {{ background: transparent; border: none; padding: 6px; border-radius: {RADIO_CONTROL}px; }}
QPushButton[variante="icono"]:hover {{ background: {SUPERFICIE_2}; }}
QPushButton[variante="segmento"] {{
    background: transparent; border: none; color: {TINTA_2}; padding: 6px 12px; border-radius: 6px;
    font-weight: 600;
}}
QPushButton[variante="segmento"]:hover {{ background: {SUPERFICIE_2}; }}
QPushButton[variante="segmento"]:checked {{ background: {MARCA_SUAVE}; color: {MARCA_FUERTE}; }}
QFrame[rol="segmentos"] {{ background: {SUPERFICIE}; border: 1px solid {BORDE_CONTROL}; border-radius: {RADIO_CONTROL}px; }}

/* Pestañas de procesos */
QPushButton[variante="pestana"] {{
    background: transparent; border: none; border-bottom: 3px solid transparent; border-radius: 0;
    color: {TINTA_2}; padding: 8px 4px; margin-right: 18px; font-size: 14px; font-weight: 600;
}}
QPushButton[variante="pestana"]:hover {{ color: {TINTA}; }}
QPushButton[variante="pestana"]:checked {{ color: {TINTA}; border-bottom-color: {MARCA}; }}

/* Campos */
QLineEdit {{
    background: {SUPERFICIE}; border: 1px solid {BORDE_CONTROL}; border-radius: {RADIO_CONTROL}px;
    padding: 7px 10px; font-size: 13px; selection-background-color: {MARCA_SUAVE}; selection-color: {TINTA};
}}
QLineEdit:focus {{ border: 2px solid {MARCA}; padding: 6px 9px; }}
QLineEdit#buscador_campo {{ border: none; font-size: 16px; padding: 12px 4px; }}

/* Listas y log */
QListWidget {{ background: {SUPERFICIE}; border: none; outline: 0; font-size: 13px; }}
QListWidget::item {{ padding: 0; border-radius: {RADIO_CONTROL}px; }}
QListWidget::item:selected {{ background: {MARCA_SUAVE}; }}
QListWidget::item:hover:!selected {{ background: {SUPERFICIE_2}; }}
QPlainTextEdit#log {{
    background: {SUPERFICIE}; border: 1px solid {BORDE}; border-radius: {RADIO}px; padding: 10px;
    color: {TINTA}; selection-background-color: {MARCA_SUAVE}; selection-color: {TINTA};
    font-family: "{familia_mono()}"; font-size: 12px;
}}

/* Tablas */
QTableWidget {{
    background: {SUPERFICIE}; border: none; gridline-color: {BORDE}; font-size: 13px;
    selection-background-color: {MARCA_SUAVE}; selection-color: {TINTA};
}}
QHeaderView::section {{
    background: {SUPERFICIE_2}; color: {TINTA_2}; border: none; border-bottom: 1px solid {BORDE};
    padding: 6px 10px; font-size: 12px; font-weight: 600;
}}

/* Barras de desplazamiento finas */
QScrollBar:vertical {{ background: transparent; width: 10px; margin: 2px; }}
QScrollBar::handle:vertical {{ background: {BORDE}; border-radius: 4px; min-height: 30px; }}
QScrollBar::handle:vertical:hover {{ background: {BORDE_CONTROL}; }}
QScrollBar::add-line, QScrollBar::sub-line {{ height: 0; width: 0; }}
QScrollBar:horizontal {{ background: transparent; height: 10px; margin: 2px; }}
QScrollBar::handle:horizontal {{ background: {BORDE}; border-radius: 4px; min-width: 30px; }}

/* Buscador (Ctrl+K) */
QFrame#buscador_panel {{ background: {SUPERFICIE}; border: 1px solid {BORDE_CONTROL}; border-radius: 12px; }}
"""
