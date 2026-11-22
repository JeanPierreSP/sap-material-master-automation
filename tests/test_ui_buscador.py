"""Buscador Ctrl+K: filtra al escribir, se maneja con el teclado y abre lo elegido."""

import pytest

pytest.importorskip("PySide6.QtWidgets")

from PySide6.QtCore import QPoint, Qt  # noqa: E402
from PySide6.QtTest import QTest  # noqa: E402


def test_filtra_por_transaccion(ventana):
    b = ventana.buscador
    b.abrir()
    b.campo.setText("mm01")
    assert [r.titulo for r in b.resultados] == ["Ampliación de almacén"]


def test_flechas_y_enter_abren_la_tarea(ventana):
    b = ventana.buscador
    b.abrir()
    b.campo.setText("mm02")
    QTest.keyClick(b.campo, Qt.Key.Key_Down)
    elegido = b.resultados[b.lista.currentRow()]
    QTest.keyClick(b.campo, Qt.Key.Key_Return)
    assert not b.isVisible()
    assert ventana.actual == f"tarea:{elegido.tarea.id}"


def test_sin_coincidencias_muestra_ayuda(ventana):
    b = ventana.buscador
    b.abrir()
    b.campo.setText("zzz")
    assert b.resultados == [] and not b.vacio.isHidden()


def test_elegir_un_maestro_abre_su_pantalla(ventana):
    b = ventana.buscador
    b.abrir()
    b.campo.setText("materiales")
    assert b.resultados[0].tarea is None
    QTest.keyClick(b.campo, Qt.Key.Key_Return)
    assert ventana.actual == "maestro:materiales"


def test_clic_afuera_cierra(ventana):
    b = ventana.buscador
    b.abrir()
    QTest.mouseClick(b, Qt.MouseButton.LeftButton, pos=QPoint(5, b.height() - 5))
    assert not b.isVisible()
