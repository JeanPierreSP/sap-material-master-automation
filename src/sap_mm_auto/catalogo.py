"""Qué ofrece la aplicación: maestros, sus procesos y las tareas de cada proceso.

Es solo una descripción, sin lógica de SAP. La interfaz arma el inicio, las pestañas del
maestro y el buscador a partir de aquí: sumar un proceso o un maestro no toca pantallas.

Un proceso ``encadenado`` es un flujo del día: sus tareas son pasos que se hacen en orden,
y la interfaz puede lanzar el siguiente paso sola cuando el anterior termina bien.
"""

from __future__ import annotations

import unicodedata
from dataclasses import dataclass

from .tareas import TAREAS, Tarea


@dataclass(frozen=True)
class Proceso:
    id: str
    nombre: str
    tareas: tuple[str, ...]
    encadenado: bool = False
    descripcion: str = ""

    def paso(self, tarea_id: str) -> int:
        """Número de paso (desde 1) de la tarea dentro del proceso."""
        return self.tareas.index(tarea_id) + 1

    def siguiente(self, tarea_id: str) -> str | None:
        """La tarea que sigue en un flujo encadenado, o None si es la última o no hay flujo."""
        if not self.encadenado or tarea_id not in self.tareas:
            return None
        i = self.tareas.index(tarea_id)
        return self.tareas[i + 1] if i + 1 < len(self.tareas) else None


@dataclass(frozen=True)
class Maestro:
    id: str
    nombre: str
    descripcion: str
    procesos: tuple[Proceso, ...] = ()

    @property
    def disponible(self) -> bool:
        return bool(self.procesos)

    @property
    def tareas(self) -> tuple[str, ...]:
        return tuple(t for p in self.procesos for t in p.tareas)


MAESTROS: tuple[Maestro, ...] = (
    Maestro(
        "materiales",
        "Materiales",
        "Descripciones, ampliaciones a almacén y bloqueos.",
        procesos=(
            Proceso(
                "mantenimiento",
                "Mantenimiento diario",
                ("cambio_descripcion", "ampliacion_almacen"),
                encadenado=True,
                descripcion="Dos pasos en orden: primero las descripciones, después las ampliaciones.",
            ),
            Proceso(
                "bloqueos",
                "Bloqueos",
                ("bloqueo_material",),
                descripcion="Status de bloqueo por material. Se revierte quitando el status.",
            ),
        ),
    ),
    Maestro("proveedores", "Proveedores", "Altas y cambios de datos bancarios."),
    Maestro("clientes", "Clientes", "Altas y condiciones de pago."),
)


def maestro(maestro_id: str) -> Maestro:
    for m in MAESTROS:
        if m.id == maestro_id:
            return m
    raise KeyError(maestro_id)


def ubicar(tarea_id: str) -> tuple[Maestro, Proceso]:
    """Maestro y proceso al que pertenece una tarea."""
    for m in MAESTROS:
        for p in m.procesos:
            if tarea_id in p.tareas:
                return m, p
    raise KeyError(tarea_id)


# --------------------------------------------------------------------------
# Búsqueda (Ctrl+K)
# --------------------------------------------------------------------------
@dataclass(frozen=True)
class Resultado:
    maestro: Maestro
    tarea: Tarea | None = None
    proceso: Proceso | None = None

    @property
    def titulo(self) -> str:
        return self.tarea.nombre if self.tarea else self.maestro.nombre

    @property
    def detalle(self) -> str:
        if self.tarea and self.proceso:
            return f"{self.maestro.nombre} · {self.proceso.nombre} · {', '.join(self.tarea.transacciones)}"
        if not self.maestro.disponible:
            return "Próximamente"
        n = len(self.maestro.tareas)
        return f"{n} tarea{'s' if n != 1 else ''}"


def _normal(texto: str) -> str:
    sin_tildes = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode()
    return " ".join(sin_tildes.lower().split())


def buscar(texto: str, limite: int = 8) -> list[Resultado]:
    """Maestros y tareas que coinciden con ``texto`` (sin distinguir tildes ni mayúsculas).

    Primero lo que empieza con el texto, después lo que lo contiene en otra parte. Sin
    texto devuelve las tareas y luego los maestros, para que el buscador nunca abra vacío.
    """
    consulta = _normal(texto)
    candidatos: list[tuple[Resultado, str, str]] = []
    for m in MAESTROS:
        for p in m.procesos:
            for t in p.tareas:
                tarea = TAREAS[t]
                extra = " ".join((m.nombre, p.nombre, tarea.descripcion, *tarea.transacciones))
                candidatos.append((Resultado(m, tarea, p), _normal(tarea.nombre), _normal(extra)))
    for m in MAESTROS:
        candidatos.append((Resultado(m), _normal(m.nombre), _normal(m.descripcion)))

    if not consulta:
        return [r for r, _, _ in candidatos][:limite]

    primero, despues = [], []
    for resultado, titulo, extra in candidatos:
        if titulo.startswith(consulta) or any(p.startswith(consulta) for p in titulo.split()):
            primero.append(resultado)
        elif consulta in titulo or consulta in extra:
            despues.append(resultado)
    return (primero + despues)[:limite]
