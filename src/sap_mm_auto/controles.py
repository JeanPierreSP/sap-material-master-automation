"""Controles que hacen seguro correr un script contra un sistema productivo."""

from __future__ import annotations

import json
import os
import tempfile
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Callable, Iterable


# --------------------------------------------------------------------------
# Memoria diaria: evita repetir trabajo si una corrida se corta y se relanza
# --------------------------------------------------------------------------
class MemoriaDiaria:
    """Guarda en un JSON por día las filas ya procesadas de cada tarea.

    Se escribe después de cada fila (no al final), de modo que un corte de luz o un
    cierre de SAP a mitad de la corrida no deja trabajo hecho sin registrar.
    """

    def __init__(self, carpeta: Path, hoy: date | None = None) -> None:
        self.ruta = Path(carpeta) / f"memoria_{(hoy or date.today()).isoformat()}.json"
        self._datos: dict[str, list[str]] = {}
        if self.ruta.exists():
            self._datos = json.loads(self.ruta.read_text(encoding="utf-8"))

    def ya_hecho(self, tarea: str, clave: str) -> bool:
        return clave in self._datos.get(tarea, [])

    def registrar(self, tarea: str, clave: str) -> None:
        self._datos.setdefault(tarea, []).append(clave)
        self._escribir()

    def _escribir(self) -> None:
        self.ruta.parent.mkdir(parents=True, exist_ok=True)
        # Escritura atómica: primero un temporal y luego se reemplaza
        fd, tmp = tempfile.mkstemp(dir=self.ruta.parent, suffix=".tmp")
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(self._datos, f, ensure_ascii=False, indent=1)
        os.replace(tmp, self.ruta)


# --------------------------------------------------------------------------
# Freno: detiene la corrida antes de que un problema se repita en cadena
# --------------------------------------------------------------------------
@dataclass
class Freno:
    max_repetidos: int = 3
    detenido: bool = False
    motivo: str = ""
    _ultimo: str = ""
    _repetidos: int = 0

    def error(self, mensaje: str) -> None:
        """Registra un error de fila. Frena si el mismo error se repite seguido."""
        normalizado = _sin_datos_de_fila(mensaje)
        self._repetidos = self._repetidos + 1 if normalizado == self._ultimo else 1
        self._ultimo = normalizado
        if self._repetidos >= self.max_repetidos:
            self.detener(f"El mismo error se repitió {self._repetidos} veces seguidas: {mensaje}")

    def exito(self) -> None:
        self._ultimo, self._repetidos = "", 0

    def detener(self, motivo: str) -> None:
        self.detenido, self.motivo = True, motivo


def _sin_datos_de_fila(mensaje: str) -> str:
    """Quita números y códigos para comparar errores de distintas filas."""
    return "".join("#" if c.isdigit() else c for c in mensaje).lower()


# --------------------------------------------------------------------------
# Precondiciones: lo que se revisa antes de tocar SAP
# --------------------------------------------------------------------------
@dataclass
class Revision:
    nombre: str
    ok: bool
    detalle: str = ""


@dataclass
class Precondiciones:
    revisiones: list[Revision] = field(default_factory=list)

    @property
    def todas_ok(self) -> bool:
        return all(r.ok for r in self.revisiones)

    def agregar(self, nombre: str, ok: bool, detalle: str = "") -> None:
        self.revisiones.append(Revision(nombre, ok, detalle))


def revisar(
    archivo: Path,
    columnas: Iterable[str],
    leer_columnas: Callable[[Path], list[str]],
    sap_disponible: Callable[[], bool],
) -> Precondiciones:
    """Revisa archivo, columnas y SAP. No modifica nada."""
    pre = Precondiciones()
    archivo = Path(archivo)

    existe = archivo.is_file()  # una carpeta o una ruta vacía no sirven como entrada
    pre.agregar("Archivo de entrada", existe, str(archivo) if existe else "No existe")

    abierto = bool(archivo.name) and archivo.with_name("~$" + archivo.name).exists()
    pre.agregar("Archivo cerrado en Excel", not abierto, "Ciérralo antes de ejecutar" if abierto else "")

    if existe:
        presentes = {c.strip().upper() for c in leer_columnas(archivo)}
        faltan = [c for c in columnas if c not in presentes]
        pre.agregar("Columnas requeridas", not faltan, "Faltan: " + ", ".join(faltan) if faltan else "")

    sap_ok = sap_disponible()
    pre.agregar("SAP abierto en la pantalla inicial", sap_ok, "" if sap_ok else "Abre SAP y vuelve al menú")
    return pre
