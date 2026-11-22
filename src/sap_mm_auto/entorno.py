"""La puerta de la interfaz hacia todo lo demás.

La interfaz no conoce SAP, el motor ni los archivos de registro: le pide a ``Entorno``
que revise, que ejecute y que lea el historial. ``Entorno`` elige el SAP (el real o el
simulado), arma la ``Pantalla``, corre el motor y guarda cada corrida en el historial.

En modo demostración todo vive en una carpeta temporal y el SAP simulado dura lo que
dura la sesión: si se ejecuta dos veces la misma tarea, la segunda ve lo que hizo la primera.
"""

from __future__ import annotations

import tempfile
import time
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Callable

from . import catalogo, entrada, motor
from .controles import Freno, MemoriaDiaria
from .historial import Corrida, Historial
from .registro import Registro, Resumen
from .sesion import ErrorSAP, Pantalla, SesionSAP, conectar_sesion_activa
from .simulador import SAPSimulado
from .tareas import TAREAS, Tarea
from .tareas.consultas import status_material
from .vista_previa import AntesDeEjecutar, revisar_antes

EJEMPLOS = Path(__file__).resolve().parents[2] / "datos_ejemplo"


@dataclass(frozen=True)
class EstadoSAP:
    conectado: bool
    texto: str


class Entorno:
    def __init__(
        self,
        demo: bool = True,
        carpeta: Path | None = None,
        sap: SesionSAP | None = None,
        retardo: float = 0.0,
        con_ejemplos: bool = False,
    ) -> None:
        self.demo = demo
        if carpeta is None:
            carpeta = Path(tempfile.mkdtemp(prefix="sap_mm_demo_")) if demo else Path.home() / "sap_mm_auto" / "registro"
        self.carpeta = Path(carpeta)
        self.carpeta.mkdir(parents=True, exist_ok=True)
        self.historial = Historial(self.carpeta / "historial.jsonl")
        self._sap = sap if sap is not None else (SAPSimulado(retardo=retardo) if demo else None)
        if con_ejemplos and demo and not self.historial.ruta.exists():
            sembrar_ejemplos(self.historial, self.carpeta, datetime.now())

    # --- SAP --------------------------------------------------------------------
    def pantalla(self) -> Pantalla:
        if self._sap is not None:
            return Pantalla(self._sap)
        _iniciar_com()
        return Pantalla(conectar_sesion_activa())

    def estado_sap(self) -> EstadoSAP:
        if self.demo:
            return EstadoSAP(True, "SAP simulado")
        try:
            en_inicio = self.pantalla().esta_en_inicio()
        except ErrorSAP:
            return EstadoSAP(False, "SAP no está abierto")
        return EstadoSAP(True, "SAP abierto" if en_inicio else "SAP abierto · fuera del menú")

    # --- archivos -----------------------------------------------------------------
    @staticmethod
    def archivo_de_ejemplo(tarea: Tarea) -> Path | None:
        ruta = EJEMPLOS / f"{tarea.id}.csv"
        return ruta if ruta.exists() else None

    # --- revisar y ejecutar -------------------------------------------------------
    def revisar(self, tarea: Tarea, archivo: Path) -> AntesDeEjecutar:
        """Chequeos y vista previa. Solo lee: la consulta a SAP es por MM03."""

        def sap_listo() -> bool:
            try:
                return self.pantalla().esta_en_inicio()
            except ErrorSAP:
                return False

        def consultar(material: str) -> str:
            return status_material(self.pantalla(), material)

        return revisar_antes(tarea, archivo, sap_listo, consultar)

    def ejecutar(
        self,
        tarea: Tarea,
        archivo: Path,
        al_linea: Callable[[str], None] | None = None,
        al_avanzar: motor.Progreso | None = None,
        al_terminar_fila: motor.ResultadoFila | None = None,
        detener: Callable[[], bool] = lambda: False,
    ) -> Corrida:
        """Corre la tarea con todos los controles y la deja en el historial."""
        inicio, reloj = datetime.now(), time.monotonic()
        registro: Registro | None = None
        try:
            filas = entrada.leer_filas(archivo)
            registro = Registro(self.carpeta, tarea.id, al_escribir=al_linea)
            resumen = motor.ejecutar(
                tarea,
                filas,
                self.pantalla(),
                MemoriaDiaria(self.carpeta),
                registro,
                Freno(),
                al_avanzar=al_avanzar,
                detener_solicitado=detener,
                al_terminar_fila=al_terminar_fila,
            )
        except (ErrorSAP, OSError, ValueError) as exc:
            if al_linea:
                al_linea(f"No se pudo ejecutar: {exc}")
            resumen = Resumen(detenido=True, motivo=f"No se pudo ejecutar: {exc}")

        try:
            maestro = catalogo.ubicar(tarea.id)[0].nombre
        except KeyError:
            maestro = ""
        rechazadas = registro.ruta_rechazadas if registro and registro.ruta_rechazadas.exists() else None
        corrida = Corrida(
            tarea=tarea.id,
            nombre=tarea.nombre,
            maestro=maestro,
            inicio=inicio.isoformat(timespec="seconds"),
            duracion=round(time.monotonic() - reloj, 2),
            total=resumen.total,
            ok=resumen.ok,
            omitidas=resumen.omitidas,
            errores=resumen.errores,
            detenido=resumen.detenido,
            motivo=resumen.motivo,
            archivo=str(archivo),
            log=str(registro.ruta_log) if registro else "",
            rechazadas=str(rechazadas) if rechazadas else "",
            demo=self.demo,
        )
        self.historial.agregar(corrida)
        return corrida


def estimar_restante(transcurrido: float, hechas: int, total: int) -> float | None:
    """Segundos que faltan al ritmo actual, o None si todavía no hay con qué estimar."""
    if hechas < 1:
        return None
    if hechas >= total:
        return 0.0
    return transcurrido / hechas * (total - hechas)


def _iniciar_com() -> None:  # pragma: no cover - solo Windows con SAP real
    """SAP GUI Scripting es COM: cada hilo que lo usa necesita su CoInitialize."""
    try:
        import pythoncom  # type: ignore[import-not-found]
    except ImportError:
        return
    pythoncom.CoInitialize()


# --------------------------------------------------------------------------
# Historial de ejemplo para la demostración
# --------------------------------------------------------------------------
_EJEMPLOS = [
    # (días atrás, hora, tarea, ok, omitidas, errores, duración, detenida, materiales con resultado)
    (1, "09:12", "cambio_descripcion", 5, 0, 0, 7.4, False, [("DEMO-0003", "OK"), ("DEMO-0004", "OK")]),
    (1, "09:20", "ampliacion_almacen", 6, 1, 1, 11.8, False, [("DEMO-0008", "ERROR"), ("DEMO-0009", "OK")]),
    (2, "16:45", "bloqueo_material", 3, 0, 0, 4.1, False, [("DEMO-0002", "OK"), ("DEMO-0007", "OK")]),
    (4, "10:03", "ampliacion_almacen", 2, 0, 3, 6.0, True, [("DEMO-0005", "ERROR"), ("DEMO-0010", "ERROR")]),
    (9, "11:30", "cambio_descripcion", 12, 2, 0, 15.2, False, [("DEMO-0001", "OK"), ("DEMO-0006", "OK")]),
]


def sembrar_ejemplos(historial: Historial, carpeta: Path, ahora: datetime) -> None:
    """Corridas ficticias de días anteriores, para que el historial y sus filtros tengan qué mostrar."""
    for dias, hora, tarea_id, ok, omitidas, errores, duracion, detenida, muestras in _EJEMPLOS:
        h, m = (int(x) for x in hora.split(":"))
        momento = (ahora - timedelta(days=dias)).replace(hour=h, minute=m, second=0, microsecond=0)
        tarea = TAREAS[tarea_id]
        total = ok + omitidas + errores
        motivo = "El mismo error se repitió 3 veces seguidas" if detenida else ""
        log = carpeta / f"{tarea_id}_{momento:%Y-%m-%d_%H%M%S}.log"
        lineas = [f"{momento:%H:%M:%S}  Inicio: {tarea.nombre} · {total} filas"]
        for i, (material, resultado) in enumerate(muestras, start=1):
            detalle = f"Se ha modificado el material {material}" if resultado == "OK" else f"El material {material} no existe"
            lineas.append(f"{momento:%H:%M:%S}  [{i}/{total}] {material}: {resultado} · {detalle}")
        if detenida:
            lineas.append(f"{momento:%H:%M:%S}  Corrida detenida: {motivo}")
        lineas.append(f"{momento:%H:%M:%S}  Fin: {ok} OK · {omitidas} omitidas · {errores} con error")
        log.write_text("\n".join(lineas) + "\n", encoding="utf-8")
        historial.agregar(
            Corrida(
                tarea=tarea_id,
                nombre=tarea.nombre,
                maestro=catalogo.ubicar(tarea_id)[0].nombre,
                inicio=momento.isoformat(timespec="seconds"),
                duracion=duracion,
                total=total,
                ok=ok,
                omitidas=omitidas,
                errores=errores,
                detenido=detenida,
                motivo=motivo,
                archivo=f"datos_ejemplo/{tarea_id}.csv",
                log=str(log),
                demo=True,
            )
        )
