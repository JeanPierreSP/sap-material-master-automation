"""Pantalla del maestro: pestañas por proceso, flujo del día y corridas recientes."""

import pytest

pytest.importorskip("PySide6.QtWidgets")

from PySide6.QtCore import Qt  # noqa: E402
from PySide6.QtTest import QTest  # noqa: E402

from sap_mm_auto.historial import FALLO, HECHO, PENDIENTE  # noqa: E402
from sap_mm_auto.tareas import TAREAS  # noqa: E402


def correr(ventana, tarea_id):
    tarea = TAREAS[tarea_id]
    return ventana.entorno.ejecutar(tarea, ventana.entorno.archivo_de_ejemplo(tarea))


def test_una_pestana_por_proceso_y_la_primera_elegida(ventana):
    m = ventana.abrir_maestro("materiales")
    assert [b.text() for b in m.pestanas.values()] == ["Mantenimiento diario", "Bloqueos"]
    assert m.proceso_actual == "mantenimiento"
    m.pestanas["bloqueos"].click()
    assert m.proceso_actual == "bloqueos"
    assert [f.tarea.id for f in m.contenidos["bloqueos"].filas] == ["bloqueo_material"]


def test_flujo_del_dia_muestra_hecho_fallo_y_pendiente(ventana):
    m = ventana.abrir_maestro("materiales")
    flujo = m.contenidos["mantenimiento"]
    assert [f.paso.estado for f in flujo.filas] == [PENDIENTE, PENDIENTE]
    assert flujo.avance.text() == "0 de 2 pasos hechos hoy"

    correr(ventana, "cambio_descripcion")
    correr(ventana, "ampliacion_almacen")  # el archivo de ejemplo trae errores a propósito
    ventana.ir_al_inicio()
    ventana.abrir_maestro("materiales")
    assert [f.paso.estado for f in flujo.filas] == [HECHO, FALLO]
    assert [f.pastilla.text() for f in flujo.filas] == ["Hecho", "Falló"]
    assert "5 OK · 3 con error" in flujo.filas[1].detalle.text()
    assert flujo.avance.text() == "1 de 2 pasos hechos hoy"


def test_clic_en_un_paso_abre_la_tarea(ventana):
    m = ventana.abrir_maestro("materiales")
    QTest.mouseClick(m.contenidos["mantenimiento"].filas[1], Qt.MouseButton.LeftButton)
    assert ventana.actual == "tarea:ampliacion_almacen"


def test_corridas_recientes_del_maestro(ventana):
    m = ventana.abrir_maestro("materiales")
    assert m.filas_recientes == []
    correr(ventana, "bloqueo_material")
    m = ventana.abrir_maestro("materiales")
    assert [f.corrida.tarea for f in m.filas_recientes] == ["bloqueo_material"]
    QTest.mouseClick(m.filas_recientes[0], Qt.MouseButton.LeftButton)
    historial = ventana.vista("historial")
    assert ventana.actual == "historial" and historial.filas[0].desplegada


def test_abrir_un_proceso_directo(ventana):
    m = ventana.abrir_maestro("materiales", "bloqueos")
    assert m.proceso_actual == "bloqueos"


def test_maestro_sin_tareas_explica_que_no_hay(ventana):
    m = ventana.abrir_maestro("proveedores")
    assert m.pestanas == {} and m.proceso_actual == ""
