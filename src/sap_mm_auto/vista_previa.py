"""Lo que se revisa antes de ejecutar: chequeos, filas del archivo y avisos por fila.

Cada chequeo termina en uno de tres estados:

- ``ok``: todo bien.
- ``aviso``: se puede ejecutar, pero conviene mirarlo (una fila que se rechazará, un
  material bloqueado). El flujo automático se frena ante un aviso.
- ``bloquea``: no se puede ejecutar hasta corregirlo (falta el archivo, SAP no está listo).

Nada de esto modifica SAP. La consulta de status es de solo lectura (MM03) y es opcional:
sin ella la revisión funciona igual, solo que sin avisos de materiales.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable

from . import entrada
from .controles import revisar
from .sesion import ErrorSAP, PantallaInesperada
from .tareas import Tarea

OK, AVISO, BLOQUEA = "ok", "aviso", "bloquea"

#: tope de consultas a MM03 por revisión: con un SAP real cada una toma alrededor de un segundo
MAX_CONSULTAS = 50


@dataclass
class Chequeo:
    nombre: str
    estado: str
    detalle: str = ""


@dataclass
class Aviso:
    fila: int  # número de fila como en Excel (la 1 es el encabezado)
    clave: str  # el material, o la clave de la fila si no hay material
    texto: str
    rechaza: bool = False  # True: el motor rechazará la fila; False: solo pide atención


@dataclass
class AntesDeEjecutar:
    chequeos: list[Chequeo] = field(default_factory=list)
    columnas: list[str] = field(default_factory=list)
    filas: list[tuple[int, dict[str, str]]] = field(default_factory=list)
    avisos: list[Aviso] = field(default_factory=list)

    @property
    def puede_ejecutar(self) -> bool:
        return all(c.estado != BLOQUEA for c in self.chequeos)

    @property
    def con_avisos(self) -> bool:
        return any(c.estado == AVISO for c in self.chequeos)

    @property
    def total(self) -> int:
        return len(self.filas)

    def avisos_de_fila(self, numero: int) -> list[Aviso]:
        return [a for a in self.avisos if a.fila == numero]

    def motivo_para_frenar(self) -> str:
        """Por qué un flujo automático no debería lanzar esta tarea sola ("" si puede)."""
        for estado in (BLOQUEA, AVISO):
            for c in self.chequeos:
                if c.estado == estado:
                    return f"{c.nombre}: {c.detalle}" if c.detalle else c.nombre
        return ""


def revisar_antes(
    tarea: Tarea,
    archivo: Path,
    sap_disponible: Callable[[], bool],
    consultar_status: Callable[[str], str] | None = None,
) -> AntesDeEjecutar:
    """Revisa archivo, filas y SAP. Devuelve los chequeos en el orden en que se muestran."""
    archivo = Path(archivo)
    pre = revisar(archivo, tarea.columnas, entrada.leer_columnas, sap_disponible)
    # controles.revisar deja la revisión de SAP al final: los chequeos de filas van antes
    *de_archivo, de_sap = pre.revisiones
    resultado = AntesDeEjecutar(chequeos=[Chequeo(r.nombre, OK if r.ok else BLOQUEA, r.detalle) for r in de_archivo])
    sap = Chequeo(de_sap.nombre, OK if de_sap.ok else BLOQUEA, de_sap.detalle)

    if all(r.ok for r in de_archivo):
        try:
            resultado.columnas = entrada.leer_columnas(archivo)
            resultado.filas = entrada.leer_filas_numeradas(archivo)
        except (OSError, ValueError) as exc:
            resultado.chequeos.append(Chequeo("Lectura del archivo", BLOQUEA, str(exc)))
        else:
            resultado.chequeos[0].detalle = f"{archivo.name} · {resultado.total} filas"
            _revisar_filas(tarea, resultado)
            if sap.estado == OK and consultar_status and "MATERIAL" in tarea.columnas:
                _consultar_materiales(tarea, resultado, consultar_status)

    resultado.chequeos.append(sap)
    return resultado


def _revisar_filas(tarea: Tarea, r: AntesDeEjecutar) -> None:
    if not r.filas:
        r.chequeos.append(Chequeo("Filas con datos", BLOQUEA, "El archivo no tiene filas"))
        return

    repetidas = entrada.duplicados([f for _, f in r.filas], tarea.clave)
    r.chequeos.append(
        Chequeo("Sin filas repetidas", BLOQUEA, "Repetidas: " + ", ".join(repetidas))
        if repetidas
        else Chequeo("Sin filas repetidas", OK)
    )

    incompletas = []
    for numero, fila in r.filas:
        motivo = tarea.validar(fila)
        if motivo:
            incompletas.append(numero)
            r.avisos.append(Aviso(numero, _clave_corta(tarea, fila), motivo, rechaza=True))
    if incompletas:
        n = len(incompletas)
        filas = ", ".join(str(x) for x in incompletas[:5]) + ("…" if n > 5 else "")
        texto = f"{n} fila{'s' if n > 1 else ''} se rechazará{'n' if n > 1 else ''} (fila {filas})"
        r.chequeos.append(Chequeo("Filas completas", AVISO, texto))
    else:
        r.chequeos.append(Chequeo("Filas completas", OK))


def _consultar_materiales(tarea: Tarea, r: AntesDeEjecutar, consultar: Callable[[str], str]) -> None:
    primera_fila: dict[str, int] = {}
    for numero, fila in r.filas:
        material = fila.get("MATERIAL", "").strip().upper()
        if material and material not in primera_fila:
            primera_fila[material] = numero

    revisados = list(primera_fila)[:MAX_CONSULTAS]
    bloqueados = inexistentes = 0
    for material in revisados:
        try:
            status = consultar(material)
        except PantallaInesperada as exc:
            r.chequeos.append(Chequeo("Materiales en SAP", AVISO, f"No se pudo consultar: {exc}"))
            return
        except ErrorSAP as exc:
            inexistentes += 1
            texto = "No existe en SAP" if "no existe" in str(exc).lower() else str(exc)
            r.avisos.append(Aviso(primera_fila[material], material, texto))
            continue
        if status and tarea.avisar_status:
            bloqueados += 1
            r.avisos.append(Aviso(primera_fila[material], material, f"Status {status}: bloqueado, SAP puede rechazarlo"))

    partes = []
    if bloqueados:
        partes.append(f"{bloqueados} bloqueado{'s' if bloqueados > 1 else ''}")
    if inexistentes:
        partes.append(f"{inexistentes} no existe{'n' if inexistentes > 1 else ''}")
    tope = f" (revisados los primeros {MAX_CONSULTAS})" if len(primera_fila) > MAX_CONSULTAS else ""
    if partes:
        r.chequeos.append(Chequeo("Materiales en SAP", AVISO, " · ".join(partes) + tope))
    else:
        n = len(revisados)
        r.chequeos.append(Chequeo("Materiales en SAP", OK, f"{n} revisado{'s' if n != 1 else ''} en MM03{tope}"))


def _clave_corta(tarea: Tarea, fila: dict[str, str]) -> str:
    material = fila.get("MATERIAL", "").strip()
    return material or tarea.clave(fila)
