"""Lectura de Excel y CSV."""

from openpyxl import Workbook

from sap_mm_auto.entrada import duplicados, leer_columnas, leer_filas


def test_lee_xlsx_y_normaliza_columnas(tmp_path):
    wb = Workbook()
    hoja = wb.active
    hoja.append(["material", " Centro ", "ALMACEN"])
    hoja.append(["DEMO-0001", 1000, "0002"])
    hoja.append([None, None, None])  # fila vacía: se ignora
    ruta = tmp_path / "entrada.xlsx"
    wb.save(ruta)

    assert leer_columnas(ruta) == ["MATERIAL", "CENTRO", "ALMACEN"]
    assert leer_filas(ruta) == [{"MATERIAL": "DEMO-0001", "CENTRO": "1000", "ALMACEN": "0002"}]


def test_detecta_duplicados(tmp_path):
    ruta = tmp_path / "e.csv"
    ruta.write_text("MATERIAL,DESCRIPCION\nA,X\nA,X\nB,Y\n", encoding="utf-8")
    datos = leer_filas(ruta)
    assert duplicados(datos, lambda f: f["MATERIAL"]) == ["A"]
