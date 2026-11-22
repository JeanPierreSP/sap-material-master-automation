"""Ventana principal: barra superior, navegación entre pantallas y la corrida en curso.

Pantallas (``ui/vistas``):

    inicio  ->  maestro (pestañas por proceso)  ->  tarea (revisar, ejecutar, resultado)
    historial, desde la barra superior o desde cualquier pantalla

La ventana es dueña de la corrida en curso (a lo más una): la tarea la pide con
``ejecutar()`` y la ventana la sigue aunque la persona navegue a otra pantalla. La barra
superior muestra entonces su avance, y un clic ahí vuelve a la tarea.

Teclado: Ctrl+K abre el buscador y Esc vuelve a la pantalla anterior.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QCloseEvent, QKeySequence, QShortcut
from PySide6.QtWidgets import QApplication, QFrame, QLabel, QMainWindow, QMessageBox, QStackedWidget, QWidget

from ..entorno import Entorno, EstadoSAP
from ..historial import Corrida
from ..tareas import TAREAS, Tarea
from . import glifos, piezas, tema
from .buscador import Buscador
from .trabajador import Corredor
from .vistas.historial import VistaHistorial
from .vistas.inicio import VistaInicio
from .vistas.maestro import VistaMaestro
from .vistas.tarea import VistaTarea

NOMBRE = "MM Auto"


class BarraSuperior(QFrame):
    def __init__(self, ventana: "Ventana") -> None:
        super().__init__()
        self.setObjectName("barra")
        self.setFixedHeight(64)
        self.ventana = ventana
        caja = piezas.fila(espacio=12, margen=0)
        caja.setContentsMargins(20, 0, 20, 0)
        self.setLayout(caja)

        marca = QWidget()
        marca.setCursor(Qt.CursorShape.PointingHandCursor)
        marca.setToolTip("Ir al inicio")
        marca.mousePressEvent = lambda _e: ventana.ir_al_inicio()
        logo = QLabel()
        logo.setPixmap(glifos.logo(30, tema.MARCA, tema.SUPERFICIE))
        textos = piezas.columna(espacio=0)
        nombre = QLabel(NOMBRE)
        nombre.setObjectName("marca_nombre")
        modo = "demostración" if ventana.entorno.demo else "SAP real"
        sub = QLabel(f"Maestro de materiales · {modo}")
        sub.setObjectName("marca_sub")
        textos.addStretch(1)
        textos.addWidget(nombre)
        textos.addWidget(sub)
        textos.addStretch(1)
        caja_marca = piezas.fila(logo, espacio=10)
        caja_marca.addLayout(textos)
        marca.setLayout(caja_marca)
        caja.addWidget(marca)
        caja.addSpacing(24)

        self.campo_buscar = piezas.boton("", None)
        self.campo_buscar.setObjectName("campo_buscar")
        self.campo_buscar.setAccessibleName("Buscar tarea o maestro (Ctrl+K)")
        self.campo_buscar.setMinimumWidth(300)
        self.campo_buscar.setMaximumWidth(460)
        self.campo_buscar.setFixedHeight(38)
        icono = QLabel()
        icono.setPixmap(glifos.pixmap("buscar", tema.TINTA_3, 16))
        tecla = QLabel("Ctrl K")
        tecla.setObjectName("tecla")
        tecla.setFixedHeight(20)
        dentro = piezas.fila(icono, piezas.etiqueta("Buscar tarea o maestro", "rotulo"), None, espacio=8)
        dentro.addWidget(tecla, 0, Qt.AlignmentFlag.AlignVCenter)
        dentro.setContentsMargins(12, 0, 10, 0)
        self.campo_buscar.setLayout(dentro)
        self.campo_buscar.clicked.connect(ventana.abrir_buscador)
        caja.addWidget(self.campo_buscar, 1)
        caja.addStretch(1)

        self.corrida = piezas.boton("", "texto", "play", tema.MARCA)
        self.corrida.setToolTip("Volver a la tarea en curso")
        self.corrida.hide()
        caja.addWidget(self.corrida)

        self.historial = piezas.boton("Historial", "texto", "historial", tema.MARCA_FUERTE)
        self.historial.clicked.connect(lambda: ventana.abrir_historial())
        caja.addWidget(self.historial)

        self.sap_punto = piezas.Punto("neutro")
        self.sap_texto = piezas.etiqueta("SAP…", "rotulo")
        sap = QWidget()
        sap.setLayout(piezas.fila(self.sap_punto, self.sap_texto, espacio=6))
        caja.addWidget(sap)

    def poner_sap(self, conectado: bool, texto: str) -> None:
        self.sap_punto.poner("ok" if conectado else "error")
        self.sap_texto.setText(texto)
        self.sap_texto.setAccessibleName(f"Estado de SAP: {texto}")

    def poner_corrida(self, texto: str | None) -> None:
        self.corrida.setVisible(bool(texto))
        if texto:
            self.corrida.setText(texto)


class Ventana(QMainWindow):
    def __init__(self, entorno: Entorno | None = None) -> None:
        super().__init__()
        self.entorno = entorno or Entorno(demo=True, con_ejemplos=True, retardo=0.25)
        self.setWindowTitle(f"{NOMBRE} · demostración" if self.entorno.demo else NOMBRE)
        self.setWindowIcon(glifos.logo(64, tema.MARCA, tema.SUPERFICIE))
        self.resize(1280, 820)

        raiz = QWidget()
        raiz.setObjectName("raiz")
        caja = piezas.columna(raiz, espacio=0)
        self.barra = BarraSuperior(self)
        caja.addWidget(self.barra)
        self.pila = QStackedWidget()
        caja.addWidget(self.pila, 1)
        self.setCentralWidget(raiz)
        self.buscador = Buscador(self, raiz)

        self._vistas: dict[str, QWidget] = {}
        self._atras: list[str] = []
        self.actual = ""
        self.corredor: Corredor | None = None
        self._tarea_en_curso: Tarea | None = None
        # Procesos con el flujo automático prendido. Vive solo en memoria: se prende a
        # conciencia en cada sesión, nunca queda guardado de una vez para otra.
        self.flujos_automaticos: set[str] = set()
        self.barra.corrida.clicked.connect(self._volver_a_la_corrida)

        QShortcut(QKeySequence("Ctrl+K"), self, self.abrir_buscador)
        QShortcut(QKeySequence(Qt.Key.Key_Escape), self, self._al_escape)

        self.refrescar_sap()
        self.ir_al_inicio()

    # --- navegación ---------------------------------------------------------------
    def vista(self, clave: str) -> QWidget | None:
        return self._vistas.get(clave)

    def _mostrar(self, clave: str, crear, recordar: bool = True) -> QWidget:
        vista = self._vistas.get(clave)
        if vista is None:
            vista = crear()
            self._vistas[clave] = vista
            self.pila.addWidget(vista)
        if recordar and self.actual and self.actual != clave:
            self._atras.append(self.actual)
        self.actual = clave
        self.pila.setCurrentWidget(vista)
        if hasattr(vista, "al_mostrar"):
            vista.al_mostrar()
        return vista

    def ir_al_inicio(self) -> None:
        self._atras.clear()
        self._mostrar("inicio", lambda: VistaInicio(self), recordar=False)

    def abrir_maestro(self, maestro_id: str, proceso_id: str | None = None) -> VistaMaestro:
        vista = self._mostrar(f"maestro:{maestro_id}", lambda: VistaMaestro(self, maestro_id))
        if proceso_id:
            vista.elegir(proceso_id)
        return vista

    def abrir_tarea(self, tarea_id: str) -> VistaTarea:
        return self._mostrar(f"tarea:{tarea_id}", lambda: VistaTarea(self, TAREAS[tarea_id]))

    def abrir_historial(self, solo_errores: bool | None = None, periodo: str | None = None) -> VistaHistorial:
        vista = self._mostrar("historial", lambda: VistaHistorial(self))
        vista.filtrar(solo_errores=solo_errores, periodo=periodo)
        return vista

    def volver(self) -> None:
        while self._atras:
            clave = self._atras.pop()
            if clave in self._vistas:
                self.actual = clave
                vista = self._vistas[clave]
                self.pila.setCurrentWidget(vista)
                if hasattr(vista, "al_mostrar"):
                    vista.al_mostrar()
                return
        self.ir_al_inicio()

    def abrir_buscador(self) -> None:
        self.buscador.abrir()

    def _al_escape(self) -> None:
        if self.buscador.isVisible():
            self.buscador.cerrar()
        elif self.actual != "inicio":
            self.volver()

    def resizeEvent(self, evento) -> None:
        super().resizeEvent(evento)
        self.buscador.ajustar()

    # --- SAP ------------------------------------------------------------------------
    def refrescar_sap(self) -> EstadoSAP:
        estado = self.entorno.estado_sap()
        self.barra.poner_sap(estado.conectado, estado.texto)
        return estado

    # --- la corrida en curso ----------------------------------------------------------
    @property
    def ocupada(self) -> bool:
        return self.corredor is not None

    def ejecutar(self, vista: VistaTarea, tarea: Tarea, archivo: Path, confirmar: bool = True) -> Corredor | None:
        """Lanza la tarea. Con SAP real pregunta antes: esto modifica datos maestros."""
        if self.corredor is not None:
            return None
        if confirmar and not self.entorno.demo:
            r = QMessageBox.warning(
                self,
                "Confirmar",
                f"Esta corrida va a modificar SAP ({', '.join(tarea.transacciones)}).\n¿Ejecutar {tarea.nombre}?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if r != QMessageBox.StandardButton.Yes:
                return None
        corredor = Corredor(self.entorno, tarea, archivo)
        # La ventana se entera primero: así, cuando la tarea recibe el fin, ya no hay corrida
        # en curso y el flujo automático puede lanzar el paso siguiente.
        corredor.avance.connect(self._al_avance)
        corredor.terminado.connect(self._al_terminar)
        corredor.linea.connect(vista.al_linea)
        corredor.avance.connect(vista.al_avance)
        corredor.fila.connect(vista.al_fila)
        corredor.terminado.connect(vista.al_terminar)
        self.corredor, self._tarea_en_curso = corredor, tarea
        self.barra.poner_corrida(f"{tarea.nombre} · empezando")
        corredor.lanzar(self)
        return corredor

    def _al_avance(self, actual: int, total: int, clave: str) -> None:
        if self._tarea_en_curso and clave:
            self.barra.poner_corrida(f"{self._tarea_en_curso.nombre} · {actual} de {total}")

    def _al_terminar(self, _corrida: Corrida) -> None:
        if self.corredor:
            self.corredor.esperar()
        self.corredor, self._tarea_en_curso = None, None
        self.barra.poner_corrida(None)

    def _volver_a_la_corrida(self) -> None:
        if self._tarea_en_curso:
            self.abrir_tarea(self._tarea_en_curso.id)

    def closeEvent(self, evento: QCloseEvent) -> None:
        if self.corredor is not None:
            r = QMessageBox.question(
                self,
                "Corrida en curso",
                "Hay una corrida en curso. ¿Detenerla después de la fila actual y salir?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if r != QMessageBox.StandardButton.Yes:
                evento.ignore()
                return
            self.corredor.detener()
            self.corredor.esperar()
        evento.accept()


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(prog="python -m sap_mm_auto.ui", description="Interfaz de escritorio")
    p.add_argument("--sap", action="store_true", help="Usa la sesión abierta de SAP GUI (Windows) en vez del simulado")
    p.add_argument("--carpeta", type=Path, help="Dónde guardar logs, memoria e historial")
    args = p.parse_args(argv)

    if args.sap and sys.platform != "win32":
        print("--sap solo funciona en Windows con SAP GUI", file=sys.stderr)
        return 2
    app = QApplication.instance() or QApplication(sys.argv[:1])
    tema.aplicar(app)
    entorno = (
        Entorno(demo=False, carpeta=args.carpeta)
        if args.sap
        else Entorno(demo=True, carpeta=args.carpeta, con_ejemplos=True, retardo=0.25)
    )
    ventana = Ventana(entorno)
    ventana.show()
    return app.exec()
