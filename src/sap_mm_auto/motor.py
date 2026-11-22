"""Bucle común: procesa las filas de una tarea con todos los controles."""

from __future__ import annotations

from typing import Callable

from .controles import Freno, MemoriaDiaria
from .registro import Registro, Resumen
from .sesion import ErrorSAP, Pantalla, PantallaInesperada
from .tareas import Fila, Tarea

Progreso = Callable[[int, int, str], None]


def ejecutar(
    tarea: Tarea,
    filas: list[Fila],
    pantalla: Pantalla,
    memoria: MemoriaDiaria,
    registro: Registro,
    freno: Freno | None = None,
    al_avanzar: Progreso | None = None,
    detener_solicitado: Callable[[], bool] = lambda: False,
) -> Resumen:
    """Procesa ``filas`` una por una.

    - Omite las filas que la memoria diaria ya tiene (corrida repetida).
    - Una fila con error no detiene la corrida: queda en el archivo de rechazadas.
    - Se detiene si SAP sale de la pantalla esperada, si el mismo error se repite
      seguido o si el usuario lo pide.
    """
    freno = freno or Freno()
    resumen = Resumen(total=len(filas))
    registro.info(f"Inicio: {tarea.nombre} · {len(filas)} filas")

    for i, fila in enumerate(filas, start=1):
        if detener_solicitado():
            freno.detener("Detenido por el usuario")
        if freno.detenido:
            break

        clave = tarea.clave(fila)
        if al_avanzar:
            al_avanzar(i, len(filas), clave)

        motivo = tarea.validar(fila)
        if motivo:
            resumen.errores += 1
            registro.rechazada(dict(fila), motivo)
            registro.info(f"[{i}/{len(filas)}] {clave}: rechazada · {motivo}")
            continue

        if memoria.ya_hecho(tarea.id, clave):
            resumen.omitidas += 1
            registro.info(f"[{i}/{len(filas)}] {clave}: omitida, ya se procesó hoy")
            continue

        try:
            if not pantalla.esta_en_inicio():
                pantalla.volver_al_inicio()
            mensaje = tarea.ejecutar(pantalla, fila)
        except PantallaInesperada as exc:
            resumen.errores += 1
            registro.rechazada(dict(fila), str(exc))
            freno.detener(f"Pantalla inesperada en {clave}: {exc}")
            registro.info(f"[{i}/{len(filas)}] {clave}: ERROR · {exc}")
            break
        except ErrorSAP as exc:
            resumen.errores += 1
            registro.rechazada(dict(fila), str(exc))
            freno.error(str(exc))
            registro.info(f"[{i}/{len(filas)}] {clave}: ERROR · {exc}")
            _intentar_volver(pantalla)
            continue

        memoria.registrar(tarea.id, clave)
        freno.exito()
        resumen.ok += 1
        registro.info(f"[{i}/{len(filas)}] {clave}: OK · {mensaje}")

    if al_avanzar:
        al_avanzar(resumen.procesadas, len(filas), "")
    resumen.detenido, resumen.motivo = freno.detenido, freno.motivo
    if freno.detenido:
        registro.info(f"Corrida detenida: {freno.motivo}")
    archivo = registro.cerrar()
    registro.info(
        f"Fin: {resumen.ok} OK · {resumen.omitidas} omitidas · {resumen.errores} con error"
        + (f" · rechazadas en {archivo.name}" if archivo else "")
    )
    return resumen


def _intentar_volver(pantalla: Pantalla) -> None:
    try:
        pantalla.volver_al_inicio()
    except ErrorSAP:
        pass
