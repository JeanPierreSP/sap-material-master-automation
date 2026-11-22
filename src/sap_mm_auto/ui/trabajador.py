"""Trabajo en segundo plano: la ventana nunca se congela mientras SAP responde.

- ``en_segundo_plano``: corre una función (por ejemplo, la revisión previa) y entrega el
  resultado en el hilo de la interfaz.
- ``Corredor``: ejecuta una tarea con el ``Entorno`` y va avisando línea a línea, fila a
  fila, hasta entregar la ``Corrida`` terminada.

Las señales llegan a métodos de objetos que viven en el hilo de la interfaz, así que Qt
las encola solo: ninguna pantalla se toca desde otro hilo.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Callable

from PySide6.QtCore import QObject, QThread, Signal, Slot

from ..entorno import Entorno
from ..tareas import Tarea


class _Trabajo(QObject):
    listo = Signal(object)

    def __init__(self, funcion: Callable[[], Any]) -> None:
        super().__init__()
        self._funcion = funcion

    @Slot()
    def correr(self) -> None:
        try:
            resultado = self._funcion()
        except Exception as exc:  # noqa: BLE001 - el error se entrega a la pantalla, no se pierde
            resultado = exc
        self.listo.emit(resultado)
        QThread.currentThread().quit()


class _Entrega(QObject):
    """Vive en el hilo de la interfaz y entrega ahí el resultado."""

    def __init__(self, al_terminar: Callable[[Any], None], padre: QObject) -> None:
        super().__init__(padre)
        self._al_terminar = al_terminar

    @Slot(object)
    def entregar(self, resultado: Any) -> None:
        self._al_terminar(resultado)


def en_segundo_plano(padre: QObject, funcion: Callable[[], Any], al_terminar: Callable[[Any], None]) -> QThread:
    """Corre ``funcion`` en otro hilo. ``al_terminar`` recibe el resultado o la excepción."""
    hilo = QThread(padre)
    trabajo = _Trabajo(funcion)
    trabajo.moveToThread(hilo)
    entrega = _Entrega(al_terminar, padre)
    hilo.started.connect(trabajo.correr)
    trabajo.listo.connect(entrega.entregar)
    hilo.finished.connect(trabajo.deleteLater)
    hilo.finished.connect(entrega.deleteLater)
    hilo.finished.connect(hilo.deleteLater)
    hilo._trabajo = trabajo  # sin esta referencia Python lo recolecta antes de que corra
    hilo.start()
    return hilo


class Corredor(QObject):
    linea = Signal(str)
    avance = Signal(int, int, str)
    fila = Signal(int, str, str)
    terminado = Signal(object)

    def __init__(self, entorno: Entorno, tarea: Tarea, archivo: Path) -> None:
        super().__init__()
        self.entorno, self.tarea, self.archivo = entorno, tarea, archivo
        self._detener = False
        self.hilo: QThread | None = None

    def detener(self) -> None:
        self._detener = True

    @property
    def deteniendo(self) -> bool:
        return self._detener

    @Slot()
    def correr(self) -> None:
        corrida = self.entorno.ejecutar(
            self.tarea,
            self.archivo,
            al_linea=self.linea.emit,
            al_avanzar=self.avance.emit,
            al_terminar_fila=self.fila.emit,
            detener=lambda: self._detener,
        )
        self.terminado.emit(corrida)
        # quit() desde aquí y no por señal: la ventana espera este hilo apenas recibe el fin
        QThread.currentThread().quit()

    def lanzar(self, padre: QObject) -> None:
        self.hilo = QThread(padre)
        self.moveToThread(self.hilo)
        self.hilo.started.connect(self.correr)
        self.hilo.start()

    def esperar(self, ms: int = 30000) -> bool:
        return self.hilo.wait(ms) if self.hilo else True
