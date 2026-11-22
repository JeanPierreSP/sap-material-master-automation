"""Cada tarea contra el SAP simulado: camino feliz y errores típicos."""

import pytest

from sap_mm_auto.sesion import ErrorSAP, PantallaInesperada
from sap_mm_auto.tareas import TAREAS

ampliacion = TAREAS["ampliacion_almacen"]
descripcion = TAREAS["cambio_descripcion"]
bloqueo = TAREAS["bloqueo_material"]


def test_ampliacion_graba_el_nuevo_almacen(sap, pantalla):
    msg = ampliacion.ejecutar(pantalla, {"MATERIAL": "DEMO-0001", "CENTRO": "1000", "ALMACEN": "0002"})
    assert "ampliado" in msg
    assert ("1000", "0002") in sap.materiales["DEMO-0001"].almacenes
    assert sap.grabaciones == [("MM01", "DEMO-0001")]


def test_ampliacion_no_graba_si_ya_existia(sap, pantalla):
    msg = ampliacion.ejecutar(pantalla, {"MATERIAL": "DEMO-0001", "CENTRO": "1000", "ALMACEN": "0001"})
    assert msg.startswith("Ya existía")
    assert sap.grabaciones == []
    assert pantalla.esta_en_inicio()


@pytest.mark.parametrize(
    "fila, texto",
    [
        ({"MATERIAL": "NO-EXISTE", "CENTRO": "1000", "ALMACEN": "0002"}, "no existe"),
        ({"MATERIAL": "DEMO-0001", "CENTRO": "9999", "ALMACEN": "0002"}, "centro 9999"),
        ({"MATERIAL": "DEMO-0001", "CENTRO": "1000", "ALMACEN": "0009"}, "almacén 0009"),
    ],
)
def test_ampliacion_errores_de_sap(sap, pantalla, fila, texto):
    with pytest.raises(ErrorSAP, match=texto):
        ampliacion.ejecutar(pantalla, fila)
    assert sap.grabaciones == []


def test_descripcion_cambia_y_no_graba_si_es_igual(sap, pantalla):
    msg = descripcion.ejecutar(pantalla, {"MATERIAL": "DEMO-0002", "DESCRIPCION": "ARANDELA 8 MM INOX"})
    assert "→" in msg and sap.materiales["DEMO-0002"].descripcion == "ARANDELA 8 MM INOX"
    msg = descripcion.ejecutar(pantalla, {"MATERIAL": "DEMO-0002", "DESCRIPCION": "ARANDELA 8 MM INOX"})
    assert msg.startswith("Sin cambios")
    assert len(sap.grabaciones) == 1


def test_descripcion_rechaza_mas_de_40_caracteres():
    assert "40" in descripcion.validar({"MATERIAL": "DEMO-0002", "DESCRIPCION": "X" * 41})


def test_material_bloqueado_por_otro_usuario():
    from sap_mm_auto.sesion import Pantalla
    from sap_mm_auto.simulador import SAPSimulado

    p = Pantalla(SAPSimulado(bloqueados_por_otro={"DEMO-0003"}))
    with pytest.raises(ErrorSAP, match="bloqueado"):
        descripcion.ejecutar(p, {"MATERIAL": "DEMO-0003", "DESCRIPCION": "NUEVA"})


def test_bloqueo_valida_el_status(sap, pantalla):
    bloqueo.ejecutar(pantalla, {"MATERIAL": "DEMO-0004", "STATUS": "01"})
    assert sap.materiales["DEMO-0004"].status == "01"
    with pytest.raises(ErrorSAP, match="no está definido"):
        bloqueo.ejecutar(pantalla, {"MATERIAL": "DEMO-0005", "STATUS": "99"})


def test_pantalla_desconocida_lanza_pantalla_inesperada():
    from sap_mm_auto.sesion import Pantalla
    from sap_mm_auto.simulador import SAPSimulado

    p = Pantalla(SAPSimulado(pantalla_inesperada={"DEMO-0001"}))
    with pytest.raises(PantallaInesperada):
        ampliacion.ejecutar(p, {"MATERIAL": "DEMO-0001", "CENTRO": "1000", "ALMACEN": "0002"})


def test_consulta_de_status_no_graba_y_vuelve_al_menu(sap, pantalla):
    from sap_mm_auto.tareas.consultas import status_material

    assert status_material(pantalla, "DEMO-0006") == "01"
    assert status_material(pantalla, "DEMO-0001") == ""
    assert sap.grabaciones == [] and pantalla.esta_en_inicio()
    with pytest.raises(ErrorSAP, match="no existe"):
        status_material(pantalla, "NO-EXISTE")
    assert pantalla.esta_en_inicio()


def test_mm03_no_permite_grabar(sap, pantalla):
    from sap_mm_auto.tareas import campos

    pantalla.ir_a("MM03")
    pantalla.escribir(campos.MATERIAL, "DEMO-0001")
    pantalla.enter()
    pantalla.guardar()
    assert "visualización" in pantalla.estado().texto and sap.grabaciones == []


def test_ampliacion_rechaza_material_con_status_de_bloqueo(sap, pantalla):
    with pytest.raises(ErrorSAP, match="status 01"):
        ampliacion.ejecutar(pantalla, {"MATERIAL": "DEMO-0006", "CENTRO": "1000", "ALMACEN": "0002"})
    assert sap.grabaciones == []
