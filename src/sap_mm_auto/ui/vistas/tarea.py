"""Pantalla de una tarea (provisional: la completa la etapa 4)."""

from __future__ import annotations

from . import Provisional


class VistaTarea(Provisional):
    def __init__(self, ventana, tarea) -> None:
        self.tarea = tarea
        super().__init__(ventana, tarea.nombre)

    def al_linea(self, texto): pass
    def al_avance(self, actual, total, clave): pass
    def al_fila(self, numero, clave, resultado): pass
    def al_terminar(self, corrida): pass
