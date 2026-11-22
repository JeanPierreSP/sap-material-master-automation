"""Piezas reutilizables de la interfaz: textos, tarjetas, botones, pastillas, barra de avance.

Ninguna sabe de SAP ni de tareas: reciben textos, números y tonos ("ok", "aviso",
"error", "marca", "neutro"). Los colores salen de ``tema``.
"""

from __future__ import annotations

from PySide6.QtCore import QRectF, QSize, Qt
from PySide6.QtGui import QColor, QPainter, QPen
from PySide6.QtWidgets import (
    QAbstractButton,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLayout,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from . import glifos, tema


# --------------------------------------------------------------------------
# Básicos
# --------------------------------------------------------------------------
def etiqueta(texto: str = "", rol: str = "cuerpo", envolver: bool = False) -> QLabel:
    e = QLabel(texto)
    e.setProperty("rol", rol)
    e.setWordWrap(envolver)
    e.setTextInteractionFlags(Qt.TextInteractionFlag.NoTextInteraction)
    return e


def tarjeta(rol: str = "tarjeta") -> QFrame:
    t = QFrame()
    t.setProperty("rol", rol)
    return t


def separador() -> QFrame:
    s = QFrame()
    s.setProperty("rol", "separador")
    return s


def boton(texto: str, variante: str | None = None, icono: str | None = None, color: str | None = None) -> QPushButton:
    b = QPushButton(texto)
    if variante:
        b.setProperty("variante", variante)
    if icono:
        b.setIcon(glifos.icono(icono, color or tema.TINTA_2, 16))
        b.setIconSize(QSize(16, 16))
    b.setCursor(Qt.CursorShape.PointingHandCursor)
    return b


def fila(*widgets: QWidget | int | None, espacio: int = 8, margen: int = 0) -> QHBoxLayout:
    """Layout horizontal. Un entero agrega ese espacio fijo; None, un estiramiento."""
    caja = QHBoxLayout()
    caja.setContentsMargins(margen, margen, margen, margen)
    caja.setSpacing(espacio)
    for w in widgets:
        if w is None:
            caja.addStretch(1)
        elif isinstance(w, int):
            caja.addSpacing(w)
        else:
            caja.addWidget(w)
    return caja


def columna(padre: QWidget | None = None, espacio: int = 8, margenes: tuple[int, int, int, int] = (0, 0, 0, 0)) -> QVBoxLayout:
    caja = QVBoxLayout(padre) if padre is not None else QVBoxLayout()
    caja.setContentsMargins(*margenes)
    caja.setSpacing(espacio)
    return caja


def vaciar(layout: QLayout) -> None:
    """Saca y destruye todo lo que tiene un layout (también los layouts anidados)."""
    while layout.count():
        item = layout.takeAt(0)
        if item.widget():
            item.widget().hide()
            item.widget().deleteLater()
        elif item.layout():
            vaciar(item.layout())


# --------------------------------------------------------------------------
# Estados
# --------------------------------------------------------------------------
class Pastilla(QLabel):
    """Texto corto con fondo suave del color de su tono."""

    def __init__(self, texto: str = "", tono: str = "neutro") -> None:
        super().__init__()
        self.setSizePolicy(QSizePolicy.Policy.Maximum, QSizePolicy.Policy.Fixed)
        self.poner(texto, tono)

    def poner(self, texto: str, tono: str) -> None:
        color, fondo = tema.TONOS[tono]
        self.tono = tono
        self.setText(texto)
        self.setStyleSheet(
            f"QLabel {{ background: {fondo}; color: {color}; border-radius: 10px; padding: 2px 10px; font-size: 12px; }}"
        )


class Punto(QWidget):
    """Círculo de estado. Va siempre junto a un texto: el color nunca es la única pista."""

    def __init__(self, tono: str = "neutro", lado: int = 9) -> None:
        super().__init__()
        self._lado = lado
        self.setFixedSize(lado, lado)
        self.poner(tono)

    def poner(self, tono: str) -> None:
        self.tono = tono
        self.update()

    def paintEvent(self, _evento) -> None:
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        color = tema.TONOS[self.tono][0] if self.tono != "neutro" else tema.BORDE_CONTROL
        p.setPen(Qt.PenStyle.NoPen)
        p.setBrush(QColor(color))
        p.drawEllipse(QRectF(0.5, 0.5, self._lado - 1, self._lado - 1))


class IconoEnCaja(QLabel):
    """Ícono sobre un cuadrado redondeado de color suave (cabecera de maestros y tareas)."""

    def __init__(self, nombre: str, lado: int = 44, apagado: bool = False) -> None:
        super().__init__()
        self.setFixedSize(lado, lado)
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        color, fondo = (tema.TINTA_3, tema.SUPERFICIE_2) if apagado else (tema.MARCA, tema.MARCA_SUAVE)
        self.setPixmap(glifos.pixmap(nombre, color, int(lado * 0.55)))
        self.setStyleSheet(f"QLabel {{ background: {fondo}; border-radius: {lado // 4}px; }}")


# --------------------------------------------------------------------------
# Controles
# --------------------------------------------------------------------------
class Interruptor(QAbstractButton):
    """Interruptor de encendido. Se maneja con el mouse, con Espacio y con el lector de pantalla."""

    def __init__(self, nombre_accesible: str) -> None:
        super().__init__()
        self.setCheckable(True)
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setAccessibleName(nombre_accesible)
        self.setFixedSize(44, 26)
        self.toggled.connect(lambda _: self.update())

    def sizeHint(self) -> QSize:
        return QSize(44, 26)

    def paintEvent(self, _evento) -> None:
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        pista = QRectF(2, 3, 40, 20)
        if self.hasFocus():
            p.setPen(QPen(QColor(tema.MARCA), 2))
            p.setBrush(Qt.BrushStyle.NoBrush)
            p.drawRoundedRect(QRectF(1, 1, 42, 24), 12, 12)
        p.setPen(Qt.PenStyle.NoPen if self.isChecked() else QPen(QColor(tema.BORDE_CONTROL), 1))
        p.setBrush(QColor(tema.MARCA if self.isChecked() else tema.SUPERFICIE_2))
        p.drawRoundedRect(pista, 10, 10)
        x = 24 if self.isChecked() else 6
        p.setPen(Qt.PenStyle.NoPen)
        p.setBrush(QColor(tema.SUPERFICIE if self.isChecked() else tema.TINTA_3))
        p.drawEllipse(QRectF(x, 6, 14, 14))


class Segmentos(QFrame):
    """Botones excluyentes entre sí (Hoy / 7 días / Todo)."""

    def __init__(self, opciones: list[tuple[str, str]], elegida: str) -> None:
        super().__init__()
        self.setProperty("rol", "segmentos")
        caja = fila(espacio=2, margen=3)
        self.setLayout(caja)
        self.botones: dict[str, QPushButton] = {}
        for clave, texto in opciones:
            b = boton(texto, "segmento")
            b.setCheckable(True)
            b.setAutoExclusive(True)
            b.setChecked(clave == elegida)
            caja.addWidget(b)
            self.botones[clave] = b

    @property
    def elegida(self) -> str:
        return next(k for k, b in self.botones.items() if b.isChecked())


# --------------------------------------------------------------------------
# Avance y resultados
# --------------------------------------------------------------------------
class BarraPorFila(QWidget):
    """Barra de avance con un tramo por fila, pintado según cómo terminó esa fila.

    Con muchas filas los tramos se juntan sin separación: sigue leyéndose como proporción.
    """

    COLORES = {"ok": tema.OK, "omitida": tema.TINTA_3, "error": tema.ERROR, "actual": tema.MARCA}

    def __init__(self) -> None:
        super().__init__()
        self.setMinimumHeight(10)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.preparar(0)

    def sizeHint(self) -> QSize:
        return QSize(400, 10)

    def preparar(self, total: int) -> None:
        self.estados: list[str] = [""] * total
        self.update()

    def marcar(self, numero: int, resultado: str) -> None:
        if 1 <= numero <= len(self.estados):
            self.estados[numero - 1] = resultado
            self.update()

    def en_curso(self, numero: int) -> None:
        if 1 <= numero <= len(self.estados) and not self.estados[numero - 1]:
            self.estados[numero - 1] = "actual"
            self.update()

    def cuenta(self, resultado: str) -> int:
        return self.estados.count(resultado)

    def paintEvent(self, _evento) -> None:
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        p.setPen(Qt.PenStyle.NoPen)
        n = max(len(self.estados), 1)
        hueco = 3.0 if n <= 60 else 0.0
        ancho = (self.width() - hueco * (n - 1)) / n
        alto = self.height()
        for i in range(n):
            estado = self.estados[i] if i < len(self.estados) else ""
            p.setBrush(QColor(self.COLORES.get(estado, tema.PISTA)))
            p.drawRoundedRect(QRectF(i * (ancho + hueco), 0, ancho, alto), 3, 3)


class Cifra(QFrame):
    """Tarjeta de resultado: un número grande con su rótulo y una marca del color del tono."""

    def __init__(self, rotulo: str, tono: str) -> None:
        super().__init__()
        self.setProperty("rol", "tarjeta")
        color = tema.TONOS[tono][0]
        caja = columna(self, espacio=2, margenes=(16, 12, 16, 12))
        marca = QFrame()
        marca.setFixedSize(24, 4)
        marca.setStyleSheet(f"background: {color}; border-radius: 2px; border: none;")
        self.numero = etiqueta("0", "cifra")
        self.numero.setStyleSheet(f"color: {color};")
        caja.addWidget(marca)
        caja.addSpacing(6)
        caja.addWidget(self.numero)
        caja.addWidget(etiqueta(rotulo, "rotulo"))

    def poner(self, valor: int) -> None:
        self.numero.setText(str(valor))


# --------------------------------------------------------------------------
# Página con desplazamiento y columna centrada
# --------------------------------------------------------------------------
class Pagina(QScrollArea):
    """Área con desplazamiento y una columna de ancho máximo, centrada en la ventana."""

    def __init__(self, margenes: tuple[int, int, int, int] = (32, 28, 32, 36)) -> None:
        super().__init__()
        self.setWidgetResizable(True)
        self.setFrameShape(QFrame.Shape.NoFrame)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        interior = QWidget()
        interior.setObjectName("pagina")
        externo = fila(espacio=0)
        interior.setLayout(externo)
        self.contenido = QWidget()
        self.contenido.setMaximumWidth(tema.ANCHO_CONTENIDO)
        self.columna = columna(self.contenido, espacio=14, margenes=margenes)
        externo.addStretch(1)
        externo.addWidget(self.contenido, 100)
        externo.addStretch(1)
        self.setWidget(interior)
