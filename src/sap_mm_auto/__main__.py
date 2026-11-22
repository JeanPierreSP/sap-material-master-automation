"""Línea de comandos.

    python -m sap_mm_auto tareas
    python -m sap_mm_auto ejecutar ampliacion_almacen datos_ejemplo/ampliacion_almacen.csv --demo
    python -m sap_mm_auto ejecutar ampliacion_almacen mi_archivo.xlsx        # SAP real (Windows)
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from . import entrada, motor
from .controles import MemoriaDiaria, revisar
from .registro import Registro
from .sesion import ErrorSAP, Pantalla, conectar_sesion_activa
from .simulador import SAPSimulado
from .tareas import TAREAS


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(prog="sap_mm_auto", description="Automatización supervisada del maestro de materiales")
    sub = p.add_subparsers(dest="comando", required=True)
    sub.add_parser("tareas", help="Lista las tareas disponibles")
    e = sub.add_parser("ejecutar", help="Ejecuta una tarea sobre un archivo")
    e.add_argument("tarea", choices=sorted(TAREAS))
    e.add_argument("archivo", type=Path)
    e.add_argument("--demo", action="store_true", help="Usa el SAP simulado en lugar de SAP real")
    e.add_argument("--carpeta", type=Path, default=Path("registro"), help="Dónde guardar log y memoria")
    e.add_argument("--si", action="store_true", help="No pedir confirmación (solo con --demo)")
    args = p.parse_args(argv)

    if args.comando == "tareas":
        for t in TAREAS.values():
            print(f"{t.id:22} {', '.join(t.transacciones):6} {t.descripcion}")
        return 0

    tarea = TAREAS[args.tarea]
    try:
        sesion = SAPSimulado() if args.demo else conectar_sesion_activa()
    except ErrorSAP as exc:
        print(f"No se pudo conectar a SAP: {exc}", file=sys.stderr)
        return 2
    pantalla = Pantalla(sesion)

    pre = revisar(args.archivo, tarea.columnas, entrada.leer_columnas, pantalla.esta_en_inicio)
    for r in pre.revisiones:
        print(f"  [{'OK' if r.ok else 'X '}] {r.nombre}{' · ' + r.detalle if r.detalle else ''}")
    if not pre.todas_ok:
        print("Corrige lo marcado con X antes de ejecutar.", file=sys.stderr)
        return 1

    filas = entrada.leer_filas(args.archivo)
    repetidas = entrada.duplicados(filas, tarea.clave)
    if repetidas:
        print(f"El archivo tiene filas repetidas: {', '.join(repetidas)}", file=sys.stderr)
        return 1

    if not args.demo:
        print(f"\nEsta corrida MODIFICARÁ SAP: {tarea.nombre}, {len(filas)} filas ({', '.join(tarea.transacciones)}).")
        if input("Escribe SI para continuar: ").strip() != "SI":
            print("Cancelado.")
            return 1

    registro = Registro(args.carpeta, tarea.id, al_escribir=print)
    resumen = motor.ejecutar(tarea, filas, pantalla, MemoriaDiaria(args.carpeta), registro)
    return 0 if resumen.errores == 0 and not resumen.detenido else 3


if __name__ == "__main__":
    sys.exit(main())
