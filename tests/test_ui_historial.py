"""Historial: filtros Hoy / 7 días / Todo, Solo errores, búsqueda y detalle con log."""

from datetime import date, datetime

import pytest

pytest.importorskip("PySide6.QtWidgets")

from sap_mm_auto.entorno import sembrar_ejemplos  # noqa: E402
from sap_mm_auto.tareas import TAREAS  # noqa: E402
from sap_mm_auto.ui.vistas.historial import dia_corto, dia_legible  # noqa: E402


@pytest.fixture
def con_corridas(ventana):
    """Cinco corridas de ejemplo de días anteriores y dos de hoy (una con errores)."""
    e = ventana.entorno
    sembrar_ejemplos(e.historial, e.carpeta, datetime.now())
    for tarea_id in ("cambio_descripcion", "ampliacion_almacen"):
        tarea = TAREAS[tarea_id]
        e.ejecutar(tarea, e.archivo_de_ejemplo(tarea))
    return ventana


def test_filtros_de_periodo(con_corridas):
    h = con_corridas.abrir_historial(periodo="hoy")
    assert len(h.filas) == 2 and h.resumen.text() == "Hoy · 2 corridas · 1 con error"
    h.periodo.botones["7dias"].click()
    assert len(h.filas) == 6  # la de hace 9 días queda fuera
    h.periodo.botones["todo"].click()
    assert len(h.filas) == 7


def test_solo_errores(con_corridas):
    h = con_corridas.abrir_historial(solo_errores=True, periodo="todo")
    assert len(h.filas) == 3 and all(f.corrida.con_problemas for f in h.filas)
    assert {f.estado.text() for f in h.filas} == {"Con error", "Detenida"}


def test_buscar_un_material_dentro_de_los_logs(con_corridas):
    h = con_corridas.abrir_historial(periodo="todo")
    h.buscar.setText("demo-0099")  # solo aparece en el log de la ampliación de hoy
    assert [f.corrida.tarea for f in h.filas] == ["ampliacion_almacen"]
    h.buscar.setText("nada-que-ver")
    assert h.filas == [] and h.vacio.text() == "No hay corridas con estos filtros"


def test_detalle_con_cifras_y_log(con_corridas):
    h = con_corridas.abrir_historial(periodo="hoy")
    h.abrir_detalle(0)
    fila = h.filas[0]
    assert fila.desplegada
    assert "Inicio:" in fila.detalle.log.toPlainText()
    h.abrir_detalle(1)  # abrir otra cierra la anterior
    assert not fila.desplegada and h.filas[1].desplegada


def test_sin_corridas_invita_a_ejecutar(ventana):
    h = ventana.abrir_historial(periodo="todo")
    assert h.filas == [] and h.resumen.text() == "Todo el historial · 0 corridas"


def test_nombres_de_dias():
    hoy = date(2026, 11, 20)
    assert dia_legible(hoy, hoy) == "Hoy"
    assert dia_legible(date(2026, 11, 19), hoy) == "Ayer"
    assert dia_legible(date(2026, 11, 16), hoy) == "Lunes 16 de noviembre"
    assert dia_corto(date(2026, 11, 16), hoy) == "16 nov"
