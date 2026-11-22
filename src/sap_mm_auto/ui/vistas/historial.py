"""Historial de corridas (provisional: lo completa la etapa 3)."""

from __future__ import annotations

from . import Provisional


class VistaHistorial(Provisional):
    def __init__(self, ventana) -> None:
        super().__init__(ventana, "Historial")

    def filtrar(self, solo_errores=None, periodo=None) -> None:
        pass
