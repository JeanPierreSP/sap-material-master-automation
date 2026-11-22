"""Inicio: tarjetas de maestros, estado del flujo de hoy, SAP y resumen del día."""

from datetime import datetime

import pytest

pytest.importorskip("PySide6.QtWidgets")

from PySide6.QtCore import Qt  # noqa: E402
from PySide6.QtTest import QTest  # noqa: E402

from sap_mm_auto.tareas import TAREAS  # noqa: E402
from sap_mm_auto.ui.vistas.inicio import saludo  # noqa: E402


def test_tres_maestros_y_solo_materiales_disponible(ventana):
    inicio = ventana.vista("inicio")
    assert list(inicio.tarjetas) == ["materiales", "proveedores", "clientes"]
    assert inicio.tarjetas["materiales"].estado.text() == "Mantenimiento diario: 0 de 2 pasos hoy"
    assert inicio.tarjetas["proveedores"].estado.text() == "Sin tareas por ahora"


def test_estado_del_flujo_sigue_a_las_corridas_de_hoy(ventana):
    entorno = ventana.entorno
    descripcion = TAREAS["cambio_descripcion"]
    entorno.ejecutar(descripcion, entorno.archivo_de_ejemplo(descripcion))
    ventana.ir_al_inicio()
    tarjeta = ventana.vista("inicio").tarjetas["materiales"]
    assert tarjeta.estado.text() == "Mantenimiento diario: 1 de 2 pasos hoy"

    ampliacion = TAREAS["ampliacion_almacen"]
    entorno.ejecutar(ampliacion, entorno.archivo_de_ejemplo(ampliacion))  # trae errores a propósito
    ventana.ir_al_inicio()
    assert "con error" in tarjeta.estado.text() and tarjeta.punto.tono == "error"
    assert ventana.vista("inicio").hoy_resumen.text() == "2 corridas"


def test_clic_y_teclado_abren_el_maestro(ventana, app):
    tarjeta = ventana.vista("inicio").tarjetas["materiales"]
    QTest.mouseClick(tarjeta, Qt.MouseButton.LeftButton)
    assert ventana.actual == "maestro:materiales"
    ventana.ir_al_inicio()
    tarjeta.setFocus()
    QTest.keyClick(tarjeta, Qt.Key.Key_Return)
    assert ventana.actual == "maestro:materiales"


def test_maestro_sin_tareas_no_navega(ventana):
    QTest.mouseClick(ventana.vista("inicio").tarjetas["clientes"], Qt.MouseButton.LeftButton)
    assert ventana.actual == "inicio"


def test_tarjeta_de_sap_en_demo(ventana):
    inicio = ventana.vista("inicio")
    assert inicio.sap_estado.text() == "SAP simulado · listo"
    assert "demostración" in inicio.sap_detalle.text().lower()


def test_saludo_segun_la_hora():
    assert saludo(datetime(2026, 1, 1, 8)) == "Buenos días"
    assert saludo(datetime(2026, 1, 1, 15)) == "Buenas tardes"
    assert saludo(datetime(2026, 1, 1, 21)) == "Buenas noches"
