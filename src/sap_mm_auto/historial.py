"""Historial de corridas: una línea JSON por corrida en ``historial.jsonl``.

Agregar es solo sumar una línea al final, así que una corrida cortada a la mitad no
estropea las anteriores. Al leer, una línea que no se entiende se salta.

También están aquí las preguntas que la interfaz le hace al historial: qué pasó hoy con
cada tarea (hecho, falló, pendiente) y qué corridas entran en cada filtro.
"""

from __future__ import annotations

import json
import unicodedata
from dataclasses import asdict, dataclass, fields
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Callable, Iterable

HECHO, FALLO, PENDIENTE = "hecho", "fallo", "pendiente"

#: días hacia atrás que cubre cada periodo (None = todo)
PERIODOS: dict[str, int | None] = {"hoy": 0, "7dias": 6, "todo": None}


@dataclass
class Corrida:
    tarea: str  # id de la tarea
    nombre: str  # nombre de la tarea al momento de correr
    maestro: str  # nombre del maestro
    inicio: str  # ISO 8601, hora local
    duracion: float  # segundos
    total: int = 0
    ok: int = 0
    omitidas: int = 0
    errores: int = 0
    detenido: bool = False
    motivo: str = ""
    archivo: str = ""  # archivo de entrada
    log: str = ""  # ruta del log de la corrida
    rechazadas: str = ""  # ruta del CSV de rechazadas, si hubo
    demo: bool = False

    @property
    def momento(self) -> datetime:
        return datetime.fromisoformat(self.inicio)

    @property
    def con_problemas(self) -> bool:
        return self.errores > 0 or self.detenido

    @property
    def estado(self) -> str:
        return FALLO if self.con_problemas else HECHO

    @property
    def resumen(self) -> str:
        partes = [f"{self.ok} OK"]
        if self.omitidas:
            partes.append(f"{self.omitidas} omitidas")
        if self.errores:
            partes.append(f"{self.errores} con error")
        if self.detenido:
            partes.append("detenida")
        return " · ".join(partes)

    def a_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def de_dict(cls, datos: dict) -> "Corrida":
        conocidos = {f.name for f in fields(cls)}
        return cls(**{k: v for k, v in datos.items() if k in conocidos})


class Historial:
    def __init__(self, ruta: Path) -> None:
        self.ruta = Path(ruta)

    def agregar(self, corrida: Corrida) -> None:
        self.ruta.parent.mkdir(parents=True, exist_ok=True)
        with self.ruta.open("a", encoding="utf-8") as f:
            f.write(json.dumps(corrida.a_dict(), ensure_ascii=False) + "\n")

    def leer(self) -> list[Corrida]:
        """Todas las corridas, de la más nueva a la más vieja."""
        if not self.ruta.exists():
            return []
        corridas = []
        for linea in self.ruta.read_text(encoding="utf-8").splitlines():
            try:
                corridas.append(Corrida.de_dict(json.loads(linea)))
            except (ValueError, TypeError):
                continue  # línea cortada o escrita a mano: no tumba la pantalla
        corridas.sort(key=lambda c: c.inicio, reverse=True)
        return corridas


# --------------------------------------------------------------------------
# Preguntas al historial
# --------------------------------------------------------------------------
def estado_del_dia(corridas: Iterable[Corrida], tarea_id: str, hoy: date | None = None) -> tuple[str, Corrida | None]:
    """Cómo quedó hoy una tarea según su última corrida del día."""
    hoy = hoy or date.today()
    for c in sorted(corridas, key=lambda c: c.inicio, reverse=True):
        if c.tarea == tarea_id and c.momento.date() == hoy:
            return c.estado, c
    return PENDIENTE, None


def filtrar(
    corridas: Iterable[Corrida],
    periodo: str = "todo",
    solo_errores: bool = False,
    texto: str = "",
    maestro: str = "",
    hoy: date | None = None,
    leer_log: Callable[[Corrida], str] | None = None,
) -> list[Corrida]:
    """Corridas que entran en el filtro, en el mismo orden en que llegan.

    ``texto`` busca en la tarea, el maestro y el archivo; con ``leer_log`` busca también
    dentro del log, para encontrar en qué corrida se tocó un material.
    """
    hoy = hoy or date.today()
    dias = PERIODOS[periodo]
    desde = hoy - timedelta(days=dias) if dias is not None else None
    consulta = _normal(texto)
    elegidas = []
    for c in corridas:
        if desde and c.momento.date() < desde:
            continue
        if solo_errores and not c.con_problemas:
            continue
        if maestro and c.maestro != maestro:
            continue
        if consulta:
            campos = _normal(" ".join((c.nombre, c.tarea, c.maestro, Path(c.archivo).name)))
            if consulta not in campos and not (leer_log and consulta in _normal(leer_log(c))):
                continue
        elegidas.append(c)
    return elegidas


def leer_log(corrida: Corrida, max_lineas: int | None = None) -> str:
    """Texto del log de una corrida ("" si ya no está)."""
    try:
        texto = Path(corrida.log).read_text(encoding="utf-8")
    except (OSError, ValueError):
        return ""
    if max_lineas is not None:
        texto = "\n".join(texto.splitlines()[-max_lineas:])
    return texto


def texto_duracion(segundos: float) -> str:
    """Duración corta para la interfaz: "<1 s", "45 s", "2 min 05 s", "1 h 03 min"."""
    s = int(round(segundos))
    if segundos < 1:
        return "<1 s"
    if s < 60:
        return f"{s} s"
    if s < 3600:
        return f"{s // 60} min {s % 60:02d} s"
    return f"{s // 3600} h {(s % 3600) // 60:02d} min"


def _normal(texto: str) -> str:
    return unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode().lower()
