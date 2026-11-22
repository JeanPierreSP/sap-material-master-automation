"""Historial de corridas: filtros, una fila por corrida agrupada por día y su detalle.

Filtros: Todas / Solo errores, Hoy / 7 días / Todo y una búsqueda que también mira dentro
de los logs (sirve para encontrar en qué corrida se tocó un material). Un clic en una fila
la despliega con sus cifras, el motivo si se detuvo y el final de su log.
"""

from __future__ import annotations

from datetime import date, timedelta
from pathlib import Path
from typing import TYPE_CHECKING

from PySide6.QtCore import Qt, QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import QFrame, QLineEdit, QPlainTextEdit, QWidget

from ...historial import Corrida, filtrar, leer_log, omitidas, texto_duracion
from .. import piezas, tema

if TYPE_CHECKING:
    from ..ventana import Ventana

POR_PAGINA = 40
LINEAS_DE_LOG = 200
DIAS = ["lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo"]
MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre",
         "noviembre", "diciembre"]


def dia_legible(dia: date, hoy: date | None = None) -> str:
    hoy = hoy or date.today()
    if dia == hoy:
        return "Hoy"
    if dia == hoy - timedelta(days=1):
        return "Ayer"
    return f"{DIAS[dia.weekday()].capitalize()} {dia.day} de {MESES[dia.month - 1]}"


def dia_corto(dia: date, hoy: date | None = None) -> str:
    """Como ``dia_legible`` pero corto, para columnas angostas: "Hoy", "Ayer", "20 nov"."""
    legible = dia_legible(dia, hoy)
    return legible if legible in ("Hoy", "Ayer") else f"{dia.day} {MESES[dia.month - 1][:3]}"


def abrir_archivo(ruta: str) -> None:
    if ruta and Path(ruta).exists():
        QDesktopServices.openUrl(QUrl.fromLocalFile(ruta))


def texto_estado(c: Corrida) -> tuple[str, str]:
    if c.detenido:
        return "Detenida", "error"
    if c.errores:
        return "Con error", "error"
    return "OK", "ok"


class FilaCorrida(QFrame):
    """Una corrida en una línea; un clic (o Enter) avisa a quien la contiene."""

    def __init__(self, corrida: Corrida, al_clic, mostrar_fecha: bool = False) -> None:
        super().__init__()
        self.corrida = corrida
        self._al_clic = al_clic
        self.setProperty("rol", "fila")
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self.setAccessibleName(f"{corrida.nombre}, {corrida.resumen}")
        caja = piezas.columna(self, espacio=0, margenes=(12, 9, 12, 9))
        momento = corrida.momento
        cuando = f"{dia_corto(momento.date())} · {momento:%H:%M}" if mostrar_fecha else f"{momento:%H:%M}"
        hora = piezas.etiqueta(cuando, "rotulo")
        hora.setFixedWidth(104 if mostrar_fecha else 52)
        textos = piezas.columna(espacio=0)
        textos.addWidget(piezas.etiqueta(corrida.nombre, "fuerte"))
        textos.addWidget(piezas.etiqueta(corrida.maestro, "rotulo"))
        resumen = piezas.etiqueta(corrida.resumen, "cuerpo")
        resumen.setMinimumWidth(230)
        duracion = piezas.etiqueta(texto_duracion(corrida.duracion), "rotulo")
        duracion.setFixedWidth(80)
        duracion.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        texto, tono = texto_estado(corrida)
        self.estado = piezas.Pastilla(texto, tono)
        self.estado.setFixedWidth(84)
        self.estado.setAlignment(Qt.AlignmentFlag.AlignCenter)
        linea = piezas.fila(hora, espacio=12)
        linea.addLayout(textos, 1)
        linea.addWidget(resumen)
        linea.addWidget(duracion)
        linea.addWidget(self.estado)
        self.flecha = piezas.etiqueta("›", "rotulo")
        self.flecha.setFixedWidth(12)
        linea.addWidget(self.flecha)
        caja.addLayout(linea)
        self.detalle: QWidget | None = None

    def mouseReleaseEvent(self, evento) -> None:
        if evento.button() == Qt.MouseButton.LeftButton:
            self._al_clic(self)

    def keyPressEvent(self, evento) -> None:
        if evento.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter, Qt.Key.Key_Space):
            self._al_clic(self)
        else:
            super().keyPressEvent(evento)

    # --- detalle ----------------------------------------------------------------------
    def desplegar(self, abierto: bool) -> None:
        if abierto and self.detalle is None:
            self.detalle = DetalleCorrida(self.corrida)
            self.layout().addWidget(self.detalle)
        if self.detalle is not None:
            self.detalle.setVisible(abierto)
        self.flecha.setText("⌄" if abierto else "›")

    @property
    def desplegada(self) -> bool:
        return self.detalle is not None and not self.detalle.isHidden()


class DetalleCorrida(QWidget):
    def __init__(self, c: Corrida) -> None:
        super().__init__()
        caja = piezas.columna(self, espacio=10, margenes=(64, 12, 0, 6))
        cifras = piezas.fila(espacio=8)
        cifras.addWidget(piezas.Pastilla(f"{c.ok} OK", "ok"))
        cifras.addWidget(piezas.Pastilla(omitidas(c.omitidas), "neutro"))
        cifras.addWidget(piezas.Pastilla(f"{c.errores} con error", "error" if c.errores else "neutro"))
        cifras.addWidget(piezas.etiqueta(f"de {c.total} filas · {texto_duracion(c.duracion)}", "rotulo"))
        cifras.addStretch(1)
        caja.addLayout(cifras)
        if c.detenido and c.motivo:
            caja.addWidget(piezas.etiqueta(f"Se detuvo: {c.motivo}", "cuerpo", envolver=True))
        if c.archivo:
            caja.addWidget(piezas.etiqueta(f"Archivo: {Path(c.archivo).name}", "rotulo"))
        self.log = QPlainTextEdit()
        self.log.setObjectName("log")
        self.log.setReadOnly(True)
        self.log.setFont(tema.fuente(9, mono=True))
        texto = leer_log(c, LINEAS_DE_LOG)
        self.log.setPlainText(texto or "El log de esta corrida ya no está en la carpeta de registro.")
        lineas = max(texto.count("\n") + 1, 2)
        self.log.setFixedHeight(min(lineas * 19 + 28, 220))
        caja.addWidget(self.log)
        botones = piezas.fila(espacio=8)
        abrir_log = piezas.boton("Abrir log", "texto", "externa", tema.MARCA_FUERTE)
        abrir_log.setEnabled(bool(texto))
        abrir_log.clicked.connect(lambda: abrir_archivo(c.log))
        botones.addWidget(abrir_log)
        if c.rechazadas:
            rechazadas = piezas.boton("Filas rechazadas (CSV)", "texto", "archivo", tema.MARCA_FUERTE)
            rechazadas.clicked.connect(lambda: abrir_archivo(c.rechazadas))
            botones.addWidget(rechazadas)
        botones.addStretch(1)
        caja.addLayout(botones)


class VistaHistorial(piezas.Pagina):
    def __init__(self, ventana: "Ventana") -> None:
        super().__init__()
        self.ventana = ventana
        self._limite = POR_PAGINA
        col = self.columna
        volver = piezas.boton("Volver", "texto", "atras", tema.MARCA_FUERTE)
        volver.clicked.connect(ventana.volver)
        col.addLayout(piezas.fila(volver, None))
        col.addWidget(piezas.etiqueta("Historial", "titulo"))
        self.resumen = piezas.etiqueta("", "subtitulo")
        col.addWidget(self.resumen)
        col.addSpacing(6)

        self.alcance = piezas.Segmentos([("todas", "Todas"), ("errores", "Solo errores")], "todas")
        self.periodo = piezas.Segmentos([("hoy", "Hoy"), ("7dias", "7 días"), ("todo", "Todo")], "7dias")
        self.buscar = QLineEdit()
        self.buscar.setPlaceholderText("Buscar tarea o material (DEMO-0004)")
        self.buscar.setAccessibleName("Buscar en el historial")
        self.buscar.setClearButtonEnabled(True)
        self.buscar.setMinimumWidth(280)
        self.buscar.setMaximumWidth(340)
        for b in (*self.alcance.botones.values(), *self.periodo.botones.values()):
            b.toggled.connect(lambda marcado: marcado and self._pintar())
        self.buscar.textChanged.connect(lambda _: self._pintar())
        col.addLayout(piezas.fila(self.alcance, self.periodo, None, self.buscar, espacio=10))
        col.addSpacing(4)

        self.tarjeta = piezas.tarjeta()
        self.lista = piezas.columna(self.tarjeta, espacio=2, margenes=(8, 8, 8, 8))
        col.addWidget(self.tarjeta)
        col.addStretch(1)
        self.filas: list[FilaCorrida] = []

    # --- lo que usan la ventana y las pruebas -------------------------------------------
    def filtrar(self, solo_errores: bool | None = None, periodo: str | None = None) -> None:
        if solo_errores is not None:
            self.alcance.botones["errores" if solo_errores else "todas"].setChecked(True)
        if periodo is not None:
            self.periodo.botones[periodo].setChecked(True)
        self._pintar()

    def abrir_detalle(self, indice: int) -> None:
        if 0 <= indice < len(self.filas):
            self._alternar(self.filas[indice])

    def mostrar_corrida(self, corrida: Corrida) -> None:
        """Deja visible y desplegada una corrida (la usa "Corridas recientes" del maestro)."""
        self.buscar.clear()
        self.alcance.botones["todas"].setChecked(True)
        self.periodo.botones["todo"].setChecked(True)
        self._pintar()
        for f in self.filas:
            if f.corrida.inicio == corrida.inicio and f.corrida.tarea == corrida.tarea:
                self._alternar(f)
                self.ensureWidgetVisible(f)
                break

    def al_mostrar(self) -> None:
        self._pintar()

    # --- pintar -----------------------------------------------------------------------
    def _pintar(self) -> None:
        todas = self.ventana.entorno.historial.leer()
        elegidas = filtrar(
            todas,
            periodo=self.periodo.elegida,
            solo_errores=self.alcance.elegida == "errores",
            texto=self.buscar.text().strip(),
            leer_log=leer_log,
        )
        con_error = sum(1 for c in elegidas if c.con_problemas)
        nombre_periodo = {"hoy": "Hoy", "7dias": "Últimos 7 días", "todo": "Todo el historial"}[self.periodo.elegida]
        n = len(elegidas)
        self.resumen.setText(
            f"{nombre_periodo} · {n} corrida{'s' if n != 1 else ''}" + (f" · {con_error} con error" if con_error else "")
        )

        piezas.vaciar(self.lista)
        self.filas = []
        self.vacio: QWidget | None = None
        if not elegidas:
            self._pintar_vacio(bool(todas))
            return
        dia_anterior = None
        for c in elegidas[: self._limite]:
            dia = c.momento.date()
            if dia != dia_anterior:
                cabecera = piezas.etiqueta(dia_legible(dia), "rotulo")
                cabecera.setContentsMargins(12, 10 if dia_anterior else 4, 0, 2)
                self.lista.addWidget(cabecera)
                dia_anterior = dia
            f = FilaCorrida(c, self._alternar)
            self.filas.append(f)
            self.lista.addWidget(f)
        if len(elegidas) > self._limite:
            mas = piezas.boton(f"Ver {min(POR_PAGINA, len(elegidas) - self._limite)} más", "texto")
            mas.clicked.connect(self._ver_mas)
            self.lista.addWidget(mas, 0, Qt.AlignmentFlag.AlignHCenter)

    def _pintar_vacio(self, hay_otras: bool) -> None:
        caja = piezas.columna(espacio=6, margenes=(16, 24, 16, 24))
        self.vacio = piezas.etiqueta("No hay corridas con estos filtros", "fuerte")
        caja.addWidget(self.vacio, 0, Qt.AlignmentFlag.AlignHCenter)
        texto = "Prueba con otro periodo o quita la búsqueda." if hay_otras else "Las corridas que ejecutes aparecen aquí."
        caja.addWidget(piezas.etiqueta(texto, "cuerpo"), 0, Qt.AlignmentFlag.AlignHCenter)
        if hay_otras:
            todo = piezas.boton("Ver todo el historial", "texto")
            todo.clicked.connect(self._ver_todo)
            caja.addWidget(todo, 0, Qt.AlignmentFlag.AlignHCenter)
        self.lista.addLayout(caja)

    def _ver_todo(self) -> None:
        self.buscar.clear()
        self.filtrar(solo_errores=False, periodo="todo")

    def _ver_mas(self) -> None:
        self._limite += POR_PAGINA
        self._pintar()

    def _alternar(self, fila: FilaCorrida) -> None:
        abrir = not fila.desplegada
        for f in self.filas:
            if f is not fila and f.desplegada:
                f.desplegar(False)
        fila.desplegar(abrir)
