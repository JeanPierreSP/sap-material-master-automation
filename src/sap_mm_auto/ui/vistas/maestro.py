"""Pantalla de un maestro (provisional: la completa la etapa 3)."""

from __future__ import annotations

from ... import catalogo
from . import Provisional


class VistaMaestro(Provisional):
    def __init__(self, ventana, maestro_id: str) -> None:
        self.maestro = catalogo.maestro(maestro_id)
        super().__init__(ventana, self.maestro.nombre)

    def elegir(self, proceso_id: str) -> None:
        pass
