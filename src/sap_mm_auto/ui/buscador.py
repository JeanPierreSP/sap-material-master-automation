"""Buscador de tareas y maestros (Ctrl+K).

Se abre encima de la pantalla actual, con un velo detrás. Se escribe y la lista se filtra
al instante (sin tildes ni mayúsculas, también por transacción: "mm02"). Flechas para
moverse, Enter para abrir, Esc o un clic afuera para cerrar.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from PySide6.QtCore import QEvent, QObject, QSize, Qt
from PySide6.QtGui import QColor, QPainter
from PySide6.QtWidgets import QFrame, QLabel, QLineEdit, QListWidget, QListWidgetItem, QWidget

from .. import catalogo
from . import glifos, piezas, tema

if TYPE_CHECKING:
    from .ventana import Ventana

ANCHO = 640


class Buscador(QWidget):
    def __init__(self, ventana: "Ventana", padre: QWidget) -> None:
        super().__init__(padre)
        self.ventana = ventana
        self.hide()
        self.panel = QFrame(self)
        self.panel.setObjectName("buscador_panel")
        self.panel.setFixedWidth(ANCHO)
        caja = piezas.columna(self.panel, espacio=0, margenes=(16, 4, 16, 12))

        icono = QLabel()
        icono.setPixmap(glifos.pixmap("buscar", tema.TINTA_3, 18))
        self.campo = QLineEdit()
        self.campo.setObjectName("buscador_campo")
        self.campo.setPlaceholderText("Buscar tarea, transacción o maestro")
        self.campo.setAccessibleName("Buscar")
        self.campo.textChanged.connect(self._llenar)
        self.campo.installEventFilter(self)
        caja.addLayout(piezas.fila(icono, self.campo, espacio=8))
        caja.addWidget(piezas.separador())
        caja.addSpacing(8)

        self.lista = QListWidget()
        self.lista.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.lista.itemActivated.connect(self._abrir)
        self.lista.itemClicked.connect(self._abrir)
        caja.addWidget(self.lista)
        self.vacio = piezas.etiqueta("Nada coincide. Prueba con el nombre de la tarea o su transacción.", "cuerpo")
        self.vacio.setContentsMargins(8, 12, 8, 12)
        caja.addWidget(self.vacio)
        caja.addSpacing(6)
        caja.addWidget(piezas.etiqueta("↑ ↓ para moverte · Enter para abrir · Esc para cerrar", "rotulo"))

    # --- abrir y cerrar --------------------------------------------------------------
    def abrir(self) -> None:
        self.ajustar()
        self.campo.clear()
        self._llenar("")
        self.show()
        self.raise_()
        self.campo.setFocus()

    def cerrar(self) -> None:
        self.hide()

    def ajustar(self) -> None:
        padre = self.parentWidget()
        self.setGeometry(0, 0, padre.width(), padre.height())
        self.panel.move(max((self.width() - ANCHO) // 2, 16), 72)

    # --- resultados -----------------------------------------------------------------
    @property
    def resultados(self) -> list[catalogo.Resultado]:
        return [self.lista.item(i).data(Qt.ItemDataRole.UserRole) for i in range(self.lista.count())]

    def _llenar(self, texto: str) -> None:
        self.lista.clear()
        for r in catalogo.buscar(texto, limite=8):
            item = QListWidgetItem()
            item.setData(Qt.ItemDataRole.UserRole, r)
            item.setSizeHint(QSize(ANCHO - 40, 52))
            self.lista.addItem(item)
            self.lista.setItemWidget(item, _Renglon(r))
        hay = self.lista.count() > 0
        self.lista.setVisible(hay)
        self.vacio.setVisible(not hay)
        if hay:
            self.lista.setCurrentRow(0)
        self.lista.setFixedHeight(min(self.lista.count(), 6) * 54 + 4)
        self.panel.adjustSize()

    def _abrir(self, item: QListWidgetItem | None = None) -> None:
        item = item or self.lista.currentItem()
        if item is None:
            return
        r: catalogo.Resultado = item.data(Qt.ItemDataRole.UserRole)
        self.cerrar()
        if r.tarea:
            self.ventana.abrir_tarea(r.tarea.id)
        else:
            self.ventana.abrir_maestro(r.maestro.id)

    # --- teclado y mouse ---------------------------------------------------------------
    def eventFilter(self, objeto: QObject, evento: QEvent) -> bool:
        if objeto is self.campo and evento.type() == QEvent.Type.KeyPress:
            tecla = evento.key()
            if tecla in (Qt.Key.Key_Down, Qt.Key.Key_Up) and self.lista.count():
                paso = 1 if tecla == Qt.Key.Key_Down else -1
                self.lista.setCurrentRow((self.lista.currentRow() + paso) % self.lista.count())
                return True
            if tecla in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
                self._abrir()
                return True
            if tecla == Qt.Key.Key_Escape:
                self.cerrar()
                return True
        return super().eventFilter(objeto, evento)

    def mousePressEvent(self, evento) -> None:
        if not self.panel.geometry().contains(evento.position().toPoint()):
            self.cerrar()

    def paintEvent(self, _evento) -> None:
        QPainter(self).fillRect(self.rect(), QColor(*tema.VELO))


class _Renglon(QWidget):
    def __init__(self, r: catalogo.Resultado) -> None:
        super().__init__()
        self.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
        apagado = r.tarea is None and not r.maestro.disponible
        if r.tarea:
            nombre_icono = "archivo"
        else:
            nombre_icono = {"materiales": "cubo", "proveedores": "camion", "clientes": "persona"}.get(r.maestro.id, "cubo")
        icono = piezas.IconoEnCaja(nombre_icono, 32, apagado=apagado)
        titulo = piezas.etiqueta(r.titulo, "fuerte")
        detalle = piezas.etiqueta(r.detalle, "rotulo")
        textos = piezas.columna(espacio=0)
        textos.addWidget(titulo)
        textos.addWidget(detalle)
        caja = piezas.fila(icono, 4, espacio=8, margen=0)
        caja.setContentsMargins(8, 6, 8, 6)
        caja.addLayout(textos, 1)
        tipo = piezas.etiqueta("Tarea" if r.tarea else "Maestro", "rotulo")
        caja.addWidget(tipo)
        self.setLayout(caja)
