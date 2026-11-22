"""Lectura del archivo de entrada (Excel o CSV)."""

from __future__ import annotations

import csv
from pathlib import Path


def leer_filas(archivo: Path) -> list[dict[str, str]]:
    """Lee un .xlsx o .csv y devuelve una lista de filas con columnas en mayúsculas."""
    return [fila for _, fila in leer_filas_numeradas(archivo)]


def leer_filas_numeradas(archivo: Path) -> list[tuple[int, dict[str, str]]]:
    """Como ``leer_filas``, pero con el número de fila tal como se ve en Excel.

    El encabezado es la fila 1. Las filas vacías se saltan sin correr la numeración,
    para que un aviso diga la misma fila que la persona ve al abrir el archivo.
    """
    archivo = Path(archivo)
    if archivo.suffix.lower() == ".xlsx":
        filas = _leer_xlsx(archivo)
    elif archivo.suffix.lower() == ".csv":
        filas = _leer_csv(archivo)
    else:
        raise ValueError(f"Formato no soportado: {archivo.suffix} (usa .xlsx o .csv)")
    return [(n, f) for n, f in enumerate(filas, start=2) if any(v.strip() for v in f.values())]


def leer_columnas(archivo: Path) -> list[str]:
    archivo = Path(archivo)
    if archivo.suffix.lower() == ".xlsx":
        from openpyxl import load_workbook

        hoja = load_workbook(archivo, read_only=True).active
        primera = next(hoja.iter_rows(max_row=1, values_only=True), ())
        return [str(c).strip().upper() for c in primera if c is not None]
    with archivo.open(encoding="utf-8-sig", newline="") as f:
        return [c.strip().upper() for c in next(csv.reader(f), [])]


def duplicados(filas: list[dict[str, str]], clave) -> list[str]:
    """Claves que aparecen más de una vez en el archivo."""
    vistas: set[str] = set()
    repetidas: list[str] = []
    for fila in filas:
        k = clave(fila)
        if k in vistas and k not in repetidas:
            repetidas.append(k)
        vistas.add(k)
    return repetidas


def _leer_csv(archivo: Path) -> list[dict[str, str]]:
    with archivo.open(encoding="utf-8-sig", newline="") as f:
        lector = csv.DictReader(f)
        return [{(k or "").strip().upper(): (v or "").strip() for k, v in fila.items()} for fila in lector]


def _leer_xlsx(archivo: Path) -> list[dict[str, str]]:
    from openpyxl import load_workbook

    hoja = load_workbook(archivo, read_only=True, data_only=True).active
    filas = hoja.iter_rows(values_only=True)
    encabezado = [str(c).strip().upper() if c is not None else "" for c in next(filas, ())]
    resultado = []
    for valores in filas:
        resultado.append(
            {encabezado[i]: ("" if v is None else str(v).strip()) for i, v in enumerate(valores) if i < len(encabezado)}
        )
    return resultado
