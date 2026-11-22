"""Maestros, procesos y búsqueda."""

from sap_mm_auto import catalogo
from sap_mm_auto.tareas import TAREAS


def test_cada_tarea_esta_una_sola_vez_en_el_catalogo():
    en_catalogo = [t for m in catalogo.MAESTROS for t in m.tareas]
    assert sorted(en_catalogo) == sorted(TAREAS)


def test_solo_materiales_tiene_tareas():
    disponibles = [m.id for m in catalogo.MAESTROS if m.disponible]
    assert disponibles == ["materiales"]
    assert [m.nombre for m in catalogo.MAESTROS] == ["Materiales", "Proveedores", "Clientes"]


def test_flujo_encadenado_sabe_cual_sigue():
    maestro, proceso = catalogo.ubicar("cambio_descripcion")
    assert maestro.id == "materiales" and proceso.encadenado
    assert proceso.siguiente("cambio_descripcion") == "ampliacion_almacen"
    assert proceso.siguiente("ampliacion_almacen") is None
    assert proceso.paso("ampliacion_almacen") == 2
    _, bloqueos = catalogo.ubicar("bloqueo_material")
    assert bloqueos.siguiente("bloqueo_material") is None


def test_buscar_sin_tildes_ni_mayusculas():
    assert catalogo.buscar("DESCRIPCION")[0].tarea.id == "cambio_descripcion"
    assert catalogo.buscar("almacén")[0].tarea.id == "ampliacion_almacen"


def test_buscar_por_transaccion_y_por_maestro():
    por_transaccion = {r.tarea.id for r in catalogo.buscar("mm02") if r.tarea}
    assert por_transaccion == {"cambio_descripcion", "bloqueo_material"}
    proveedores = catalogo.buscar("prov")[0]
    assert proveedores.tarea is None and proveedores.detalle == "Próximamente"


def test_buscar_sin_texto_nunca_abre_vacio():
    resultados = catalogo.buscar("")
    assert resultados and resultados[0].tarea is not None
