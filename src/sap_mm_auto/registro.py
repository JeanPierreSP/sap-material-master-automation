"""Registro de cada corrida: log de texto y CSV de filas rechazadas para reprocesar."""

from __future__ import annotations

import csv
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Callable


@dataclass
class Resumen:
    total: int = 0
    ok: int = 0
    omitidas: int = 0
    errores: int = 0
    detenido: bool = False
    motivo: str = ""

    @property
    def procesadas(self) -> int:
        return self.ok + self.omitidas + self.errores


@dataclass
class Registro:
    carpeta: Path
    tarea: str
    al_escribir: Callable[[str], None] | None = None
    _rechazadas: list[dict[str, str]] = field(default_factory=list)

    def __post_init__(self) -> None:
        self.carpeta = Path(self.carpeta)
        self.carpeta.mkdir(parents=True, exist_ok=True)
        sello = datetime.now().strftime("%Y-%m-%d_%H%M%S")
        self.ruta_log = self.carpeta / f"{self.tarea}_{sello}.log"
        self.ruta_rechazadas = self.carpeta / f"{self.tarea}_{sello}_rechazadas.csv"

    def info(self, texto: str) -> None:
        linea = f"{datetime.now():%H:%M:%S}  {texto}"
        with self.ruta_log.open("a", encoding="utf-8") as f:
            f.write(linea + "\n")
        if self.al_escribir:
            self.al_escribir(linea)

    def rechazada(self, fila: dict[str, str], motivo: str) -> None:
        self._rechazadas.append({**fila, "MOTIVO": motivo})

    def cerrar(self) -> Path | None:
        """Escribe las filas rechazadas, con las mismas columnas, listas para reprocesar."""
        if not self._rechazadas:
            return None
        columnas = list(dict.fromkeys(k for fila in self._rechazadas for k in fila))
        with self.ruta_rechazadas.open("w", encoding="utf-8-sig", newline="") as f:
            escritor = csv.DictWriter(f, fieldnames=columnas)
            escritor.writeheader()
            escritor.writerows(self._rechazadas)
        return self.ruta_rechazadas
