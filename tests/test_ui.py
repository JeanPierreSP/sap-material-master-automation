"""Ventana: tema, barra superior y navegación (sin pantalla)."""

import pytest

pytest.importorskip("PySide6.QtWidgets")

from PySide6.QtCore import Qt  # noqa: E402
from PySide6.QtTest import QTest  # noqa: E402

from sap_mm_auto.ui import tema  # noqa: E402


def _luminancia(color: str) -> float:
    canales = [int(color.lstrip("#")[i : i + 2], 16) / 255 for i in (0, 2, 4)]
    lineal = [c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4 for c in canales]
    return 0.2126 * lineal[0] + 0.7152 * lineal[1] + 0.0722 * lineal[2]


def contraste(a: str, b: str) -> float:
    claro, oscuro = sorted((_luminancia(a), _luminancia(b)), reverse=True)
    return (claro + 0.05) / (oscuro + 0.05)


@pytest.mark.parametrize(
    "texto, fondo",
    [
        (tema.TINTA, tema.FONDO),
        (tema.TINTA_2, tema.FONDO),
        (tema.TINTA_2, tema.SUPERFICIE_2),
        (tema.TINTA_3, tema.SUPERFICIE),
        (tema.TINTA_3, tema.SUPERFICIE_2),
        (tema.SUPERFICIE, tema.MARCA),
        (tema.MARCA, tema.FONDO),
        (tema.MARCA_FUERTE, tema.MARCA_SUAVE),
        *[(color, fondo) for color, fondo in tema.TONOS.values()],
        *[(color, tema.SUPERFICIE) for color, _ in tema.TONOS.values()],
    ],
)
def test_textos_pasan_aa(texto, fondo):
    assert contraste(texto, fondo) >= 4.5


@pytest.mark.parametrize("color", [tema.BORDE_CONTROL, tema.MARCA, tema.OK, tema.ERROR, tema.TINTA_3])
def test_controles_y_barra_pasan_3_a_1(color):
    fondo = tema.SUPERFICIE if color == tema.BORDE_CONTROL else tema.PISTA
    assert contraste(color, fondo) >= 3.0


def test_abre_en_el_inicio_con_sap_simulado(ventana):
    assert ventana.actual == "inicio"
    assert ventana.barra.sap_texto.text() == "SAP simulado"
    assert "demostración" in ventana.windowTitle()


def test_navegar_y_volver_con_esc(ventana, app):
    ventana.abrir_maestro("materiales")
    ventana.abrir_tarea("ampliacion_almacen")
    assert ventana.actual == "tarea:ampliacion_almacen"
    QTest.keyClick(ventana, Qt.Key.Key_Escape)
    assert ventana.actual == "maestro:materiales"
    QTest.keyClick(ventana, Qt.Key.Key_Escape)
    assert ventana.actual == "inicio"


def test_ctrl_k_abre_el_buscador(ventana, app):
    QTest.keyClick(ventana, Qt.Key.Key_K, Qt.KeyboardModifier.ControlModifier)
    assert ventana.buscador.isVisible()
    QTest.keyClick(ventana.buscador.campo, Qt.Key.Key_Escape)
    assert not ventana.buscador.isVisible()


def test_historial_desde_la_barra(ventana):
    ventana.barra.historial.click()
    assert ventana.actual == "historial"


def test_texto_de_ejemplo_de_los_campos_usa_la_paleta(app):
    from PySide6.QtGui import QPalette

    assert app.palette().color(QPalette.ColorRole.PlaceholderText).name().upper() == tema.TINTA_3
