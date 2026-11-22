"""Memoria diaria, freno y precondiciones."""

from datetime import date

from sap_mm_auto.controles import Freno, MemoriaDiaria, revisar
from sap_mm_auto.entrada import leer_columnas


def test_memoria_sobrevive_a_un_reinicio(tmp_path):
    MemoriaDiaria(tmp_path).registrar("tarea", "A")
    assert MemoriaDiaria(tmp_path).ya_hecho("tarea", "A")
    assert not MemoriaDiaria(tmp_path).ya_hecho("otra_tarea", "A")


def test_memoria_empieza_vacia_cada_dia(tmp_path):
    MemoriaDiaria(tmp_path, hoy=date(2026, 1, 1)).registrar("tarea", "A")
    assert not MemoriaDiaria(tmp_path, hoy=date(2026, 1, 2)).ya_hecho("tarea", "A")


def test_freno_tras_el_mismo_error_tres_veces():
    f = Freno(max_repetidos=3)
    f.error("El material DEMO-0001 no existe")
    f.error("El material DEMO-0002 no existe")  # mismo error, otro material
    assert not f.detenido
    f.error("El material DEMO-0003 no existe")
    assert f.detenido and "3 veces" in f.motivo


def test_freno_se_reinicia_con_un_exito():
    f = Freno(max_repetidos=2)
    f.error("El material DEMO-0001 no existe")
    f.exito()
    f.error("El material DEMO-0002 no existe")
    assert not f.detenido


def test_precondiciones_detectan_columnas_faltantes_y_excel_abierto(tmp_path):
    archivo = tmp_path / "entrada.csv"
    archivo.write_text("MATERIAL,CENTRO\nDEMO-0001,1000\n", encoding="utf-8")
    (tmp_path / "~$entrada.csv").write_text("", encoding="utf-8")  # así marca Excel un archivo abierto

    pre = revisar(archivo, ("MATERIAL", "CENTRO", "ALMACEN"), leer_columnas, lambda: True)
    estado = {r.nombre: r for r in pre.revisiones}
    assert not pre.todas_ok
    assert not estado["Archivo cerrado en Excel"].ok
    assert "ALMACEN" in estado["Columnas requeridas"].detalle
