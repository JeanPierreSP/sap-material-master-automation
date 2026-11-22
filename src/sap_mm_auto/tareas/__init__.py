"""Registro de tareas disponibles."""

from __future__ import annotations

from .ampliacion_almacen import AmpliacionAlmacen
from .base import Fila, Tarea
from .bloqueo_material import BloqueoMaterial
from .cambio_descripcion import CambioDescripcion

TAREAS: dict[str, Tarea] = {t.id: t for t in (AmpliacionAlmacen(), CambioDescripcion(), BloqueoMaterial())}

__all__ = ["TAREAS", "Fila", "Tarea"]
