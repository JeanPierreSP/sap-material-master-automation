"""Pantalla de un maestro: una pestaña por proceso y las corridas recientes.

En un proceso encadenado (flujo del día) cada tarea es un paso numerado, unido al
siguiente, con su estado de hoy: hecho, falló o pendiente. En un proceso suelto, las
tareas van en lista con el mismo estado. Debajo, las últimas corridas del maestro.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from PySide6.QtCore import QRectF, Qt
from PySide6.QtGui import QColor, QFont, QPainter, QPen
from PySide6.QtWidgets import QFrame, QPushButton, QWidget

from ... import catalogo
from ...historial import FALLO, HECHO, PENDIENTE, Corrida, estado_del_dia
from ...tareas import TAREAS, Tarea
from .. import glifos, piezas, tema
from .historial import FilaCorrida
from .inicio import ICONOS

if TYPE_CHECKING:
    from ..ventana import Ventana

RECIENTES = 5
TEXTO_ESTADO = {HECHO: ("Hecho", "ok"), FALLO: ("Falló", "error"), PENDIENTE: ("Pendiente", "neutro")}


class MarcaPaso(QWidget):
    """Círculo del paso (número, ✓ o !) con la línea que lo une al anterior y al siguiente."""

    def __init__(self, numero: int, primero: bool, ultimo: bool) -> None:
        super().__init__()
        self.numero, self.primero, self.ultimo = numero, primero, ultimo
        self.estado = PENDIENTE
        self.setFixedWidth(36)

    def poner(self, estado: str) -> None:
        self.estado = estado
        self.update()

    def paintEvent(self, _evento) -> None:
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        cx, cy, r = self.width() / 2, self.height() / 2, 14
        p.setPen(QPen(QColor(tema.BORDE_CONTROL), 1.5))
        if not self.primero:
            p.drawLine(int(cx), 0, int(cx), int(cy - r))
        if not self.ultimo:
            p.drawLine(int(cx), int(cy + r), int(cx), self.height())
        circulo = QRectF(cx - r, cy - r, 2 * r, 2 * r)
        if self.estado == HECHO:
            p.setPen(Qt.PenStyle.NoPen)
            p.setBrush(QColor(tema.OK))
            p.drawEllipse(circulo)
            p.drawPixmap(int(cx - 8), int(cy - 8), glifos.pixmap("check", tema.SUPERFICIE, 16, 2.2))
            return
        color = tema.ERROR if self.estado == FALLO else tema.BORDE_CONTROL
        p.setPen(QPen(QColor(color), 2))
        p.setBrush(QColor(tema.ERROR_SUAVE if self.estado == FALLO else tema.SUPERFICIE))
        p.drawEllipse(circulo)
        fuente = tema.fuente(10, QFont.Weight.DemiBold)
        p.setFont(fuente)
        p.setPen(QColor(tema.ERROR if self.estado == FALLO else tema.TINTA_2))
        p.drawText(circulo, Qt.AlignmentFlag.AlignCenter, "!" if self.estado == FALLO else str(self.numero))


class FilaTarea(QFrame):
    def __init__(self, ventana: "Ventana", tarea: Tarea, paso: MarcaPaso | None) -> None:
        super().__init__()
        self.ventana, self.tarea, self.paso = ventana, tarea, paso
        self.setProperty("rol", "fila")
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self.setAccessibleName(f"Abrir {tarea.nombre}")
        self.setMinimumHeight(72)
        caja = piezas.fila(espacio=14)
        caja.setContentsMargins(12, 10, 12, 10)
        self.setLayout(caja)
        caja.addWidget(paso if paso else piezas.IconoEnCaja("archivo", 36))
        textos = piezas.columna(espacio=2)
        textos.addWidget(piezas.etiqueta(tarea.nombre, "fuerte"))
        textos.addWidget(piezas.etiqueta(tarea.descripcion, "cuerpo", envolver=True))
        caja.addLayout(textos, 1)
        transaccion = piezas.Pastilla(" · ".join(tarea.transacciones), "marca")
        caja.addWidget(transaccion)
        caja.addSpacing(10)
        estado = piezas.columna(espacio=2)
        self.pastilla = piezas.Pastilla("Pendiente", "neutro")
        self.detalle = piezas.etiqueta("", "rotulo")
        estado.addWidget(self.pastilla, 0, Qt.AlignmentFlag.AlignRight)
        estado.addWidget(self.detalle, 0, Qt.AlignmentFlag.AlignRight)
        contenedor = QWidget()
        contenedor.setLayout(estado)
        contenedor.setFixedWidth(170)
        caja.addWidget(contenedor)
        caja.addWidget(piezas.etiqueta("›", "seccion"))

    def actualizar(self, corridas: list[Corrida]) -> str:
        estado, ultima = estado_del_dia(corridas, self.tarea.id)
        texto, tono = TEXTO_ESTADO[estado]
        self.pastilla.poner(texto, tono)
        self.detalle.setText(f"{ultima.momento:%H:%M} · {ultima.resumen}" if ultima else "Sin corridas hoy")
        if self.paso:
            self.paso.poner(estado)
        return estado

    def mouseReleaseEvent(self, evento) -> None:
        if evento.button() == Qt.MouseButton.LeftButton:
            self.ventana.abrir_tarea(self.tarea.id)

    def keyPressEvent(self, evento) -> None:
        if evento.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter, Qt.Key.Key_Space):
            self.ventana.abrir_tarea(self.tarea.id)
        else:
            super().keyPressEvent(evento)


class PestanaProceso(QWidget):
    def __init__(self, ventana: "Ventana", proceso: catalogo.Proceso) -> None:
        super().__init__()
        self.proceso = proceso
        caja = piezas.columna(self, espacio=10)
        self.titulo = piezas.etiqueta("Flujo del día" if proceso.encadenado else "Tareas", "seccion")
        self.avance = piezas.etiqueta("", "rotulo")
        caja.addLayout(piezas.fila(self.titulo, None, self.avance))
        if proceso.descripcion:
            caja.addWidget(piezas.etiqueta(proceso.descripcion, "cuerpo", envolver=True))
        tarjeta = piezas.tarjeta()
        lista = piezas.columna(tarjeta, espacio=0, margenes=(8, 8, 8, 8))
        self.filas: list[FilaTarea] = []
        n = len(proceso.tareas)
        for i, tarea_id in enumerate(proceso.tareas):
            paso = MarcaPaso(i + 1, i == 0, i == n - 1) if proceso.encadenado else None
            fila = FilaTarea(ventana, TAREAS[tarea_id], paso)
            self.filas.append(fila)
            lista.addWidget(fila)
        caja.addWidget(tarjeta)
        if proceso.encadenado:
            nota = piezas.etiqueta(
                "Con el flujo automático (en la pantalla de cada paso), al terminar bien un paso se lanza el "
                "siguiente. Se frena ante cualquier error o aviso.",
                "rotulo",
                envolver=True,
            )
            caja.addWidget(nota)

    def actualizar(self, corridas: list[Corrida]) -> None:
        estados = [f.actualizar(corridas) for f in self.filas]
        if self.proceso.encadenado:
            hechos = estados.count(HECHO)
            self.avance.setText(f"{hechos} de {len(estados)} pasos hechos hoy")
        else:
            self.avance.setText("")


class VistaMaestro(piezas.Pagina):
    def __init__(self, ventana: "Ventana", maestro_id: str) -> None:
        super().__init__()
        self.ventana = ventana
        self.maestro = catalogo.maestro(maestro_id)
        col = self.columna
        volver = piezas.boton("Datos maestros", "texto", "atras", tema.MARCA_FUERTE)
        volver.clicked.connect(ventana.ir_al_inicio)
        col.addLayout(piezas.fila(volver, None))

        cabecera = piezas.fila(piezas.IconoEnCaja(ICONOS.get(self.maestro.id, "cubo"), 52), 6, espacio=12)
        textos = piezas.columna(espacio=2)
        textos.addWidget(piezas.etiqueta(self.maestro.nombre, "titulo"))
        textos.addWidget(piezas.etiqueta(self.maestro.descripcion, "subtitulo"))
        cabecera.addLayout(textos, 1)
        col.addLayout(cabecera)
        col.addSpacing(4)

        self.pestanas: dict[str, QPushButton] = {}
        self.contenidos: dict[str, PestanaProceso] = {}
        barra = piezas.fila(espacio=0)
        for p in self.maestro.procesos:
            b = piezas.boton(p.nombre, "pestana")
            b.setCheckable(True)
            b.setAutoExclusive(True)
            b.clicked.connect(lambda _=False, pid=p.id: self.elegir(pid))
            self.pestanas[p.id] = b
            barra.addWidget(b)
        barra.addStretch(1)
        col.addLayout(barra)
        col.addWidget(piezas.separador())
        col.addSpacing(6)

        if not self.maestro.disponible:
            col.addWidget(piezas.etiqueta("Este maestro todavía no tiene tareas", "seccion"))
            col.addWidget(piezas.etiqueta("Cuando se sumen procesos al catálogo, aparecerán aquí.", "cuerpo"))
        for p in self.maestro.procesos:
            contenido = PestanaProceso(ventana, p)
            contenido.hide()
            self.contenidos[p.id] = contenido
            col.addWidget(contenido)

        col.addSpacing(14)
        self.titulo_recientes = piezas.etiqueta("Corridas recientes", "seccion")
        todo = piezas.boton("Ver historial", "texto", "derecha", tema.MARCA_FUERTE)
        todo.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        todo.clicked.connect(lambda: ventana.abrir_historial(periodo="7dias"))
        col.addLayout(piezas.fila(self.titulo_recientes, None, todo))
        self.tarjeta_recientes = piezas.tarjeta()
        self.recientes = piezas.columna(self.tarjeta_recientes, espacio=2, margenes=(8, 8, 8, 8))
        col.addWidget(self.tarjeta_recientes)
        col.addStretch(1)
        self.filas_recientes: list[FilaCorrida] = []

        if self.maestro.procesos:
            self.elegir(self.maestro.procesos[0].id)

    @property
    def proceso_actual(self) -> str:
        return next((pid for pid, c in self.contenidos.items() if not c.isHidden()), "")

    def elegir(self, proceso_id: str) -> None:
        for pid, contenido in self.contenidos.items():
            contenido.setVisible(pid == proceso_id)
            self.pestanas[pid].setChecked(pid == proceso_id)

    def al_mostrar(self) -> None:
        corridas = self.ventana.entorno.historial.leer()
        for contenido in self.contenidos.values():
            contenido.actualizar(corridas)
        piezas.vaciar(self.recientes)
        propias = [c for c in corridas if c.maestro == self.maestro.nombre][:RECIENTES]
        self.filas_recientes = []
        for c in propias:
            fila = FilaCorrida(c, self._abrir_corrida, mostrar_fecha=True)
            self.filas_recientes.append(fila)
            self.recientes.addWidget(fila)
        if not propias:
            vacio = piezas.etiqueta("Todavía no hay corridas de este maestro.", "cuerpo")
            vacio.setContentsMargins(12, 12, 12, 12)
            self.recientes.addWidget(vacio)

    def _abrir_corrida(self, fila: FilaCorrida) -> None:
        self.ventana.abrir_historial().mostrar_corrida(fila.corrida)
