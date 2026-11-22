"""Contrato común de las tareas.

Cada tarea describe qué columnas necesita su archivo de entrada, cómo identificar una
fila (para no repetirla) y los pasos de SAP GUI para procesarla.
"""

from __future__ import annotations

from abc import ABC, abstractmethod
from typing import Mapping

from ..sesion import Pantalla

Fila = Mapping[str, str]


class Tarea(ABC):
    #: identificador corto, usado en la línea de comandos y en la memoria diaria
    id: str = ""
    #: nombre para la interfaz
    nombre: str = ""
    #: qué hace, en una línea
    descripcion: str = ""
    #: transacciones que usa (solo informativo)
    transacciones: tuple[str, ...] = ()
    #: columnas obligatorias del archivo de entrada
    columnas: tuple[str, ...] = ()
    #: si la revisión previa avisa de materiales con status de bloqueo
    avisar_status: bool = True

    def clave(self, fila: Fila) -> str:
        """Identifica la fila. Dos filas con la misma clave son el mismo trabajo."""
        return "|".join(fila[c].strip().upper() for c in self.columnas)

    def validar(self, fila: Fila) -> str | None:
        """Devuelve un motivo si la fila no se puede procesar, o None si está bien."""
        for columna in self.columnas:
            if not str(fila.get(columna, "")).strip():
                return f"Falta {columna}"
        return None

    @abstractmethod
    def ejecutar(self, pantalla: Pantalla, fila: Fila) -> str:
        """Procesa una fila en SAP y devuelve el mensaje de éxito.

        Debe lanzar ``ErrorSAP`` si algo falla. ``PantallaInesperada`` detiene la corrida.
        """
