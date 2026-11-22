"""Íconos de línea dibujados con QPainter.

Así la aplicación no depende de archivos de imagen ni de fuentes de íconos, y cada ícono
toma el color que le pide la pantalla. Todos se dibujan en una grilla de 24 × 24.
"""

from __future__ import annotations

from typing import Callable

from PySide6.QtCore import QPointF, QRectF, Qt
from PySide6.QtGui import QBrush, QColor, QIcon, QPainter, QPainterPath, QPen, QPixmap


def _linea(p: QPainter, *puntos: tuple[float, float]) -> None:
    camino = QPainterPath(QPointF(*puntos[0]))
    for x, y in puntos[1:]:
        camino.lineTo(x, y)
    p.drawPath(camino)


def _cerrado(p: QPainter, *puntos: tuple[float, float], relleno: bool = False) -> None:
    camino = QPainterPath(QPointF(*puntos[0]))
    for x, y in puntos[1:]:
        camino.lineTo(x, y)
    camino.closeSubpath()
    if relleno:
        p.fillPath(camino, p.pen().color())
    p.drawPath(camino)


def _buscar(p):
    p.drawEllipse(QPointF(10.5, 10.5), 6.5, 6.5)
    _linea(p, (15.5, 15.5), (20.5, 20.5))


def _historial(p):
    p.drawArc(QRectF(3.5, 3.5, 17, 17), 160 * 16, -320 * 16)
    _linea(p, (3.2, 5.5), (3.6, 9.6), (7.6, 9.0))
    _linea(p, (12, 7.5), (12, 12), (15, 14))


def _atras(p):
    _linea(p, (19, 12), (5, 12))
    _linea(p, (11, 6), (5, 12), (11, 18))


def _derecha(p):
    _linea(p, (9, 6), (15, 12), (9, 18))


def _abajo(p):
    _linea(p, (6, 9), (12, 15), (18, 9))


def _check(p):
    _linea(p, (5, 12.5), (10, 17.5), (19, 7))


def _cruz(p):
    _linea(p, (6.5, 6.5), (17.5, 17.5))
    _linea(p, (17.5, 6.5), (6.5, 17.5))


def _alerta(p):
    _cerrado(p, (12, 3.5), (21, 19.5), (3, 19.5))
    _linea(p, (12, 9.5), (12, 13.5))
    p.setBrush(QBrush(p.pen().color()))
    p.drawEllipse(QPointF(12, 16.6), 0.6, 0.6)


def _play(p):
    _cerrado(p, (8, 5.5), (18.5, 12), (8, 18.5), relleno=True)


def _stop(p):
    camino = QPainterPath()
    camino.addRoundedRect(QRectF(6.5, 6.5, 11, 11), 2, 2)
    p.fillPath(camino, p.pen().color())


def _carpeta(p):
    _cerrado(p, (3, 6.5), (9.5, 6.5), (11.5, 8.5), (21, 8.5), (21, 18.5), (3, 18.5))


def _archivo(p):
    _cerrado(p, (6, 3), (14, 3), (19, 8), (19, 21), (6, 21))
    _linea(p, (14, 3), (14, 8), (19, 8))


def _refrescar(p):
    p.drawArc(QRectF(4.5, 4.5, 15, 15), 60 * 16, 290 * 16)
    _linea(p, (19.5, 3.8), (19.2, 8.4), (14.7, 8.0))


def _rayo(p):
    _cerrado(p, (13, 2.5), (5, 13.5), (11.5, 13.5), (10.5, 21.5), (19, 10), (12.5, 10))


def _cubo(p):
    _cerrado(p, (12, 2.8), (20.5, 7.4), (20.5, 16.6), (12, 21.2), (3.5, 16.6), (3.5, 7.4))
    _linea(p, (3.5, 7.4), (12, 12), (20.5, 7.4))
    _linea(p, (12, 12), (12, 21.2))


def _camion(p):
    p.drawRect(QRectF(2.5, 6.5, 11, 9.5))
    _linea(p, (13.5, 9.5), (17.5, 9.5), (21, 13), (21, 16), (13.5, 16))
    p.drawEllipse(QPointF(7, 17.8), 1.9, 1.9)
    p.drawEllipse(QPointF(17, 17.8), 1.9, 1.9)


def _persona(p):
    p.drawEllipse(QPointF(12, 8), 3.6, 3.6)
    camino = QPainterPath(QPointF(4.5, 20.5))
    camino.cubicTo(5.5, 14.5, 18.5, 14.5, 19.5, 20.5)
    p.drawPath(camino)


def _monitor(p):
    p.drawRoundedRect(QRectF(3, 4.5, 18, 12), 1.5, 1.5)
    _linea(p, (12, 16.5), (12, 20))
    _linea(p, (8.5, 20), (15.5, 20))


def _log(p):
    _linea(p, (5, 7), (19, 7))
    _linea(p, (5, 12), (19, 12))
    _linea(p, (5, 17), (14, 17))


def _externa(p):
    _linea(p, (14, 4), (20, 4), (20, 10))
    _linea(p, (20, 4), (11, 13))
    _linea(p, (17, 14), (17, 20), (4, 20), (4, 7), (10, 7))


DIBUJOS: dict[str, Callable[[QPainter], None]] = {
    "buscar": _buscar,
    "historial": _historial,
    "atras": _atras,
    "derecha": _derecha,
    "abajo": _abajo,
    "check": _check,
    "cruz": _cruz,
    "alerta": _alerta,
    "play": _play,
    "stop": _stop,
    "carpeta": _carpeta,
    "archivo": _archivo,
    "refrescar": _refrescar,
    "rayo": _rayo,
    "cubo": _cubo,
    "camion": _camion,
    "persona": _persona,
    "monitor": _monitor,
    "log": _log,
    "externa": _externa,
}


def pixmap(nombre: str, color: str, tam: int = 18, grosor: float = 1.6) -> QPixmap:
    escala = 2  # nítido también en pantallas al 150 % o 200 %
    pm = QPixmap(tam * escala, tam * escala)
    pm.fill(Qt.GlobalColor.transparent)
    pm.setDevicePixelRatio(escala)
    p = QPainter(pm)
    p.setRenderHint(QPainter.RenderHint.Antialiasing)
    p.scale(tam / 24, tam / 24)
    pluma = QPen(QColor(color), grosor * 24 / tam)
    pluma.setCapStyle(Qt.PenCapStyle.RoundCap)
    pluma.setJoinStyle(Qt.PenJoinStyle.RoundJoin)
    p.setPen(pluma)
    p.setBrush(Qt.BrushStyle.NoBrush)
    DIBUJOS[nombre](p)
    p.end()
    return pm


def icono(nombre: str, color: str, tam: int = 18) -> QIcon:
    return QIcon(pixmap(nombre, color, tam))


def logo(lado: int, fondo: str, trazo: str) -> QPixmap:
    """La marca: un cuadrado redondeado con tres filas, la última ya procesada (✓)."""
    escala = 2
    pm = QPixmap(lado * escala, lado * escala)
    pm.fill(Qt.GlobalColor.transparent)
    pm.setDevicePixelRatio(escala)
    p = QPainter(pm)
    p.setRenderHint(QPainter.RenderHint.Antialiasing)
    p.scale(lado / 24, lado / 24)
    p.setPen(Qt.PenStyle.NoPen)
    p.setBrush(QColor(fondo))
    p.drawRoundedRect(QRectF(0, 0, 24, 24), 6, 6)
    p.setBrush(QColor(trazo))
    for y, ancho in ((6, 12), (10.75, 9), (15.5, 6)):
        p.drawRoundedRect(QRectF(5, y, ancho, 2.5), 1.25, 1.25)
    pluma = QPen(QColor(trazo), 2.2)
    pluma.setCapStyle(Qt.PenCapStyle.RoundCap)
    pluma.setJoinStyle(Qt.PenJoinStyle.RoundJoin)
    p.setPen(pluma)
    _linea(p, (13.5, 16.5), (15.5, 18.5), (19.5, 13.5))
    p.end()
    return pm
