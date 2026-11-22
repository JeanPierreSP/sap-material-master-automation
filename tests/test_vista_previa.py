"""Revisión antes de ejecutar: chequeos, filas y avisos."""

from pathlib import Path

from sap_mm_auto.sesion import Pantalla
from sap_mm_auto.simulador import SAPSimulado
from sap_mm_auto.tareas import TAREAS
from sap_mm_auto.tareas.consultas import status_material
from sap_mm_auto.vista_previa import AVISO, BLOQUEA, OK, revisar_antes

EJEMPLOS = Path(__file__).resolve().parents[1] / "datos_ejemplo"
ampliacion = TAREAS["ampliacion_almacen"]


def revisar_con_sap(tarea, archivo, sap=None):
    pantalla = Pantalla(sap or SAPSimulado())
    return revisar_antes(tarea, archivo, pantalla.esta_en_inicio, lambda m: status_material(pantalla, m))


def test_archivo_de_ejemplo_avisa_fila_vacia_material_bloqueado_e_inexistente():
    sap = SAPSimulado()
    r = revisar_con_sap(ampliacion, EJEMPLOS / "ampliacion_almacen.csv", sap)

    estados = {c.nombre: c.estado for c in r.chequeos}
    assert r.puede_ejecutar and r.con_avisos
    assert estados["Filas completas"] == AVISO
    assert estados["Materiales en SAP"] == AVISO
    assert estados["Sin filas repetidas"] == OK
    assert r.chequeos[-1].nombre.startswith("SAP") and r.chequeos[-1].estado == OK

    avisos = {(a.fila, a.clave): a for a in r.avisos}
    assert avisos[(5, "DEMO-0004")].rechaza and "ALMACEN" in avisos[(5, "DEMO-0004")].texto
    assert "no existe" in avisos[(6, "DEMO-0099")].texto.lower()
    assert "status 01" in avisos[(8, "DEMO-0006")].texto.lower()
    assert r.total == 8 and r.filas[0] == (2, {"MATERIAL": "DEMO-0001", "CENTRO": "1000", "ALMACEN": "0002"})

    # Revisar no graba nada y deja SAP donde estaba
    assert sap.grabaciones == [] and Pantalla(sap).esta_en_inicio()


def test_sin_consulta_a_sap_igual_revisa_el_archivo():
    r = revisar_antes(ampliacion, EJEMPLOS / "ampliacion_almacen.csv", lambda: True)
    assert "Materiales en SAP" not in {c.nombre for c in r.chequeos}
    assert [a.clave for a in r.avisos] == ["DEMO-0004"]


def test_filas_repetidas_bloquean(tmp_path):
    archivo = tmp_path / "e.csv"
    archivo.write_text("MATERIAL,CENTRO,ALMACEN\nDEMO-0001,1000,0002\ndemo-0001,1000,0002\n", encoding="utf-8")
    r = revisar_con_sap(ampliacion, archivo)
    repetidas = next(c for c in r.chequeos if c.nombre == "Sin filas repetidas")
    assert repetidas.estado == BLOQUEA and not r.puede_ejecutar
    assert r.motivo_para_frenar().startswith("Sin filas repetidas")


def test_archivo_inexistente_bloquea_y_no_consulta_sap(tmp_path):
    consultas = []
    r = revisar_antes(ampliacion, tmp_path / "no.csv", lambda: True, consultas.append)
    assert not r.puede_ejecutar and r.filas == [] and consultas == []


def test_sap_fuera_del_menu_bloquea():
    r = revisar_antes(ampliacion, EJEMPLOS / "ampliacion_almacen.csv", lambda: False, lambda m: "")
    assert r.chequeos[-1].estado == BLOQUEA and not r.puede_ejecutar


def test_bloqueo_no_avisa_status_pero_si_inexistentes(tmp_path):
    archivo = tmp_path / "b.csv"
    archivo.write_text("MATERIAL,STATUS\nDEMO-0006,02\nDEMO-0099,01\n", encoding="utf-8")
    r = revisar_con_sap(TAREAS["bloqueo_material"], archivo)
    assert [a.clave for a in r.avisos] == ["DEMO-0099"]
