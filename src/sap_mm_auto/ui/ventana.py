"""Interfaz de escritorio (PySide6).

La persona elige una tarea, revisa las precondiciones y ejecuta. El avance y el log se
ven en vivo, y la corrida se puede detener en cualquier momento. En modo demostración
todo corre contra el SAP simulado.
"""

from __future__ import annotations

import sys
import tempfile
from pathlib import Path

from PySide6.QtCore import QObject, Qt, QThread, Signal
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from .. import entrada, motor
from ..controles import Freno, MemoriaDiaria, revisar
from ..registro import Registro, Resumen
from ..sesion import ErrorSAP, Pantalla, conectar_sesion_activa
from ..simulador import SAPSimulado
from ..tareas import TAREAS, Tarea

RAIZ = Path(__file__).resolve().parents[3]
EJEMPLOS = RAIZ / "datos_ejemplo"

ESTILO = """
QMainWindow, QWidget#panel { background: #f6f7f9; }
QWidget#lateral { background: #1f2a37; }
QLabel#marca { color: #ffffff; font-size: 15px; font-weight: 600; padding: 18px 16px 4px 16px; }
QLabel#submarca { color: #9aa5b1; font-size: 11px; padding: 0 16px 14px 16px; }
QListWidget#tareas { background: transparent; border: none; color: #d9dee4; font-size: 13px; outline: 0; }
QListWidget#tareas::item { padding: 10px 16px; }
QListWidget#tareas::item:selected { background: #2f3e4e; color: #ffffff; border-left: 3px solid #3aa0ff; }
QLabel#titulo { font-size: 20px; font-weight: 600; color: #1f2a37; }
QLabel#descripcion { color: #52606d; font-size: 13px; }
QLabel#chip { background: #e3eefc; color: #1d5fa8; border-radius: 9px; padding: 2px 10px; font-size: 11px; }
QFrame#tarjeta { background: #ffffff; border: 1px solid #e1e5ea; border-radius: 8px; }
QLabel#cifra { font-size: 22px; font-weight: 600; color: #1f2a37; }
QLabel#etiqueta { color: #7b8794; font-size: 11px; }
QPushButton { border: 1px solid #cbd2d9; border-radius: 6px; padding: 7px 14px; background: #ffffff; }
QPushButton#primario { background: #1d6fd1; color: #ffffff; border: none; font-weight: 600; }
QPushButton#primario:disabled { background: #9fbfe6; }
QPushButton#peligro { color: #b42318; }
QProgressBar { border: none; background: #e1e5ea; border-radius: 4px; height: 8px; text-align: center; }
QProgressBar::chunk { background: #1d6fd1; border-radius: 4px; }
QPlainTextEdit { background: #111827; color: #d1d5db; border-radius: 8px; padding: 8px; }
QLineEdit { border: 1px solid #cbd2d9; border-radius: 6px; padding: 6px 8px; background: #ffffff; }
"""


class _Trabajador(QObject):
    linea = Signal(str)
    avance = Signal(int, int, str)
    terminado = Signal(object)

    def __init__(self, tarea: Tarea, archivo: Path, demo: bool) -> None:
        super().__init__()
        self.tarea, self.archivo, self.demo = tarea, archivo, demo
        self._detener = False

    def detener(self) -> None:
        self._detener = True

    def correr(self) -> None:
        try:
            sesion = SAPSimulado(retardo=0.06) if self.demo else conectar_sesion_activa()
            carpeta = Path(tempfile.mkdtemp(prefix="demo_")) if self.demo else Path.home() / "sap_mm_auto" / "registro"
            resumen = motor.ejecutar(
                self.tarea,
                entrada.leer_filas(self.archivo),
                Pantalla(sesion),
                MemoriaDiaria(carpeta),
                Registro(carpeta, self.tarea.id, al_escribir=self.linea.emit),
                Freno(),
                al_avanzar=self.avance.emit,
                detener_solicitado=lambda: self._detener,
            )
        except (ErrorSAP, OSError, ValueError) as exc:
            self.linea.emit(f"No se pudo ejecutar: {exc}")
            resumen = Resumen(detenido=True, motivo=str(exc))
        self.terminado.emit(resumen)


class Ventana(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("SAP MM Automation · demostración")
        self.resize(1040, 660)
        self._hilo: QThread | None = None
        self._trabajador: _Trabajador | None = None

        raiz = QWidget()
        fila = QHBoxLayout(raiz)
        fila.setContentsMargins(0, 0, 0, 0)
        fila.setSpacing(0)
        fila.addWidget(self._lateral())
        fila.addWidget(self._panel(), 1)
        self.setCentralWidget(raiz)
        self.lista.setCurrentRow(0)

    # --- construcción --------------------------------------------------------
    def _lateral(self) -> QWidget:
        lateral = QWidget(objectName="lateral")
        lateral.setFixedWidth(240)
        col = QVBoxLayout(lateral)
        col.setContentsMargins(0, 0, 0, 0)
        col.addWidget(QLabel("SAP MM Automation", objectName="marca"))
        col.addWidget(QLabel("Maestro de materiales · RPA supervisado", objectName="submarca"))
        self.lista = QListWidget(objectName="tareas")
        for t in TAREAS.values():
            item = QListWidgetItem(t.nombre)
            item.setData(Qt.UserRole, t.id)
            self.lista.addItem(item)
        self.lista.currentItemChanged.connect(self._al_elegir)
        col.addWidget(self.lista, 1)
        return lateral

    def _panel(self) -> QWidget:
        panel = QWidget(objectName="panel")
        col = QVBoxLayout(panel)
        col.setContentsMargins(28, 24, 28, 24)
        col.setSpacing(12)

        cabecera = QHBoxLayout()
        self.titulo = QLabel(objectName="titulo")
        self.chip = QLabel(objectName="chip")
        cabecera.addWidget(self.titulo)
        cabecera.addWidget(self.chip)
        cabecera.addStretch()
        col.addLayout(cabecera)
        self.descripcion = QLabel(objectName="descripcion")
        col.addWidget(self.descripcion)

        archivo = QHBoxLayout()
        self.ruta = QLineEdit()
        examinar = QPushButton("Examinar…")
        examinar.clicked.connect(self._examinar)
        archivo.addWidget(QLabel("Archivo:"))
        archivo.addWidget(self.ruta, 1)
        archivo.addWidget(examinar)
        col.addLayout(archivo)

        self.demo = QCheckBox("Modo demostración (SAP simulado, no toca SAP)")
        self.demo.setChecked(True)
        self.demo.setEnabled(sys.platform == "win32")
        col.addWidget(self.demo)

        self.revisiones = QLabel()
        self.revisiones.setTextFormat(Qt.RichText)
        col.addWidget(self.revisiones)

        botones = QHBoxLayout()
        self.btn_revisar = QPushButton("Revisar")
        self.btn_ejecutar = QPushButton("Ejecutar", objectName="primario")
        self.btn_detener = QPushButton("Detener", objectName="peligro")
        self.btn_detener.setEnabled(False)
        self.btn_revisar.clicked.connect(self._revisar)
        self.btn_ejecutar.clicked.connect(self._ejecutar)
        self.btn_detener.clicked.connect(self._detener)
        for b in (self.btn_revisar, self.btn_ejecutar, self.btn_detener):
            botones.addWidget(b)
        botones.addStretch()
        col.addLayout(botones)

        self.progreso = QProgressBar()
        self.progreso.setTextVisible(False)
        self.texto_progreso = QLabel("Listo para ejecutar", objectName="etiqueta")
        col.addWidget(self.progreso)
        col.addWidget(self.texto_progreso)

        tarjetas = QHBoxLayout()
        self.cifras = {}
        for clave, nombre in (("ok", "Procesadas OK"), ("omitidas", "Omitidas (ya hechas hoy)"), ("errores", "Con error")):
            tarjeta = QFrame(objectName="tarjeta")
            t = QVBoxLayout(tarjeta)
            cifra = QLabel("0", objectName="cifra")
            t.addWidget(cifra)
            t.addWidget(QLabel(nombre, objectName="etiqueta"))
            self.cifras[clave] = cifra
            tarjetas.addWidget(tarjeta)
        col.addLayout(tarjetas)

        self.log = QPlainTextEdit()
        self.log.setReadOnly(True)
        self.log.setFont(QFont("DejaVu Sans Mono", 9))
        col.addWidget(self.log, 1)
        return panel

    # --- acciones --------------------------------------------------------------
    @property
    def tarea(self) -> Tarea:
        return TAREAS[self.lista.currentItem().data(Qt.UserRole)]

    def _al_elegir(self, *_: object) -> None:
        t = self.tarea
        self.titulo.setText(t.nombre)
        self.chip.setText(" · ".join(t.transacciones))
        self.descripcion.setText(f"{t.descripcion}  Columnas: {', '.join(t.columnas)}.")
        ejemplo = EJEMPLOS / f"{t.id}.csv"
        self.ruta.setText(str(ejemplo) if ejemplo.exists() else "")
        self.revisiones.setText("")
        self._reiniciar_cifras()

    def _examinar(self) -> None:
        ruta, _ = QFileDialog.getOpenFileName(self, "Archivo de entrada", str(EJEMPLOS), "Datos (*.xlsx *.csv)")
        if ruta:
            self.ruta.setText(ruta)

    def _revisar(self) -> bool:
        demo = self.demo.isChecked()

        def sap_ok() -> bool:
            if demo:
                return True
            try:
                return Pantalla(conectar_sesion_activa()).esta_en_inicio()
            except ErrorSAP:
                return False

        pre = revisar(Path(self.ruta.text()), self.tarea.columnas, entrada.leer_columnas, sap_ok)
        filas = []
        for r in pre.revisiones:
            color, marca = ("#1a7f37", "✓") if r.ok else ("#b42318", "✗")
            detalle = f" <span style='color:#7b8794'>· {r.detalle}</span>" if r.detalle and not r.ok else ""
            filas.append(f"<span style='color:{color}'>{marca}</span> {r.nombre}{detalle}")
        self.revisiones.setText("&nbsp;&nbsp;&nbsp;".join(filas))
        return pre.todas_ok

    def _ejecutar(self) -> None:
        if not self._revisar():
            return
        demo = self.demo.isChecked()
        if not demo:
            r = QMessageBox.warning(
                self,
                "Confirmar",
                f"Esta corrida modificará SAP ({', '.join(self.tarea.transacciones)}). ¿Continuar?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if r != QMessageBox.Yes:
                return
        self._reiniciar_cifras()
        self.log.clear()
        self._ocupado(True)
        self._hilo = QThread(self)
        self._trabajador = _Trabajador(self.tarea, Path(self.ruta.text()), demo)
        self._trabajador.moveToThread(self._hilo)
        self._hilo.started.connect(self._trabajador.correr)
        self._trabajador.linea.connect(self._al_linea)
        self._trabajador.avance.connect(self._al_avance)
        self._trabajador.terminado.connect(self._al_terminar)
        self._hilo.start()

    def _detener(self) -> None:
        if self._trabajador:
            self._trabajador.detener()
            self.texto_progreso.setText("Deteniendo después de la fila actual…")

    # --- señales ---------------------------------------------------------------
    def _al_linea(self, texto: str) -> None:
        self.log.appendPlainText(texto)
        for clave, marca in (("ok", ": OK"), ("omitidas", ": omitida"), ("errores", ": ERROR")):
            if marca in texto:
                self.cifras[clave].setText(str(int(self.cifras[clave].text()) + 1))
        if ": rechazada" in texto:
            self.cifras["errores"].setText(str(int(self.cifras["errores"].text()) + 1))

    def _al_avance(self, actual: int, total: int, clave: str) -> None:
        self.progreso.setMaximum(max(total, 1))
        self.progreso.setValue(actual if not clave else actual - 1)
        if clave:
            self.texto_progreso.setText(f"{actual} de {total} · {clave}")

    def _al_terminar(self, resumen: Resumen) -> None:
        self._ocupado(False)
        self.progreso.setValue(self.progreso.maximum())
        if resumen.detenido:
            self.texto_progreso.setText(f"Detenida: {resumen.motivo}")
        else:
            self.texto_progreso.setText(f"Terminada: {resumen.ok} OK · {resumen.omitidas} omitidas · {resumen.errores} con error")
        if self._hilo:
            self._hilo.quit()
            self._hilo.wait()

    def _ocupado(self, si: bool) -> None:
        self.btn_ejecutar.setEnabled(not si)
        self.btn_revisar.setEnabled(not si)
        self.btn_detener.setEnabled(si)
        self.lista.setEnabled(not si)

    def _reiniciar_cifras(self) -> None:
        for cifra in self.cifras.values():
            cifra.setText("0")
        self.progreso.setValue(0)
        self.texto_progreso.setText("Listo para ejecutar")


def main() -> int:
    app = QApplication(sys.argv)
    app.setStyleSheet(ESTILO)
    ventana = Ventana()
    ventana.show()
    return app.exec()
