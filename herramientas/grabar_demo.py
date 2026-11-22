"""Graba docs/demo.gif y docs/captura.png recorriendo la interfaz en modo demostración.

    python herramientas/grabar_demo.py

Recorrido: inicio -> Materiales -> Ampliación de almacén -> revisar -> ejecutar -> resultado.
Todo corre contra el SAP simulado. La ventana no aparece en pantalla (WA_DontShowOnScreen):
se fotografía con ``grab()``. En un servidor sin pantalla, con QT_QPA_PLATFORM=offscreen.

Cada clic deja un anillo por un instante, para que quien mire el GIF siga la navegación.
Los cuadros iguales seguidos se juntan en uno más largo: el GIF pesa mucho menos.
"""

from __future__ import annotations

import sys
import time
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ / "src"))

from PIL import Image, ImageChops, ImageDraw  # noqa: E402
from PySide6.QtCore import QPoint, Qt  # noqa: E402
from PySide6.QtWidgets import QApplication, QWidget  # noqa: E402

from sap_mm_auto.entorno import Entorno  # noqa: E402
from sap_mm_auto.ui import tema  # noqa: E402
from sap_mm_auto.ui.ventana import Ventana  # noqa: E402

ANCHO, ALTO = 1280, 800  # tamaño lógico de la ventana
ANCHO_GIF = 1040
CADA = 0.12  # segundos entre cuadros
ANILLO = (90, 63, 192)  # tema.MARCA en RGB


class Grabadora:
    def __init__(self, app: QApplication, ventana: Ventana) -> None:
        self.app, self.ventana = app, ventana
        self.cuadros: list[tuple[Image.Image, int]] = []  # (imagen, duración en ms)
        self._clic: tuple[QPoint, int] | None = None

    def foto(self) -> Image.Image:
        qimg = self.ventana.grab().toImage().convertToFormat(self.ventana.grab().toImage().Format.Format_RGB888)
        ancho, alto = qimg.width(), qimg.height()
        datos = bytes(qimg.constBits())[: qimg.bytesPerLine() * alto]
        img = Image.frombuffer("RGB", (ancho, alto), datos, "raw", "RGB", qimg.bytesPerLine(), 1).copy()
        if self._clic:
            punto, quedan = self._clic
            escala = ancho / self.ventana.width()
            x, y, r = punto.x() * escala, punto.y() * escala, (26 - quedan * 3) * escala
            capa = Image.new("RGBA", img.size)
            ImageDraw.Draw(capa).ellipse((x - r, y - r, x + r, y + r), outline=(*ANILLO, 220), width=int(4 * escala))
            img = Image.alpha_composite(img.convert("RGBA"), capa).convert("RGB")
            self._clic = (punto, quedan - 1) if quedan > 1 else None
        return img

    def esperar(self, segundos: float) -> None:
        fin = time.time() + segundos
        while time.time() < fin:
            self.app.processEvents()
            img = self.foto()
            ms = int(CADA * 1000)
            if self.cuadros and ImageChops.difference(img, self.cuadros[-1][0]).getbbox() is None:
                anterior, duracion = self.cuadros[-1]
                self.cuadros[-1] = (anterior, duracion + ms)
            else:
                self.cuadros.append((img, ms))
            time.sleep(CADA)

    def hasta(self, condicion, limite: float = 30) -> None:
        fin = time.time() + limite
        while not condicion() and time.time() < fin:
            self.esperar(CADA)

    def clic(self, widget: QWidget, accion) -> None:
        """Marca el clic sobre ``widget`` y ejecuta ``accion`` (lo que haría el clic real)."""
        self._clic = (widget.mapTo(self.ventana, widget.rect().center()), 5)
        self.esperar(0.4)
        accion()


def main() -> None:
    app = QApplication.instance() or QApplication(sys.argv[:1])
    tema.aplicar(app)
    ventana = Ventana(Entorno(demo=True, con_ejemplos=True, retardo=0.12))
    ventana.setAttribute(Qt.WidgetAttribute.WA_DontShowOnScreen)
    ventana.resize(ANCHO, ALTO)
    ventana.show()
    g = Grabadora(app, ventana)

    g.esperar(1.6)  # inicio
    inicio = ventana.vista("inicio")
    g.clic(inicio.tarjetas["materiales"], lambda: ventana.abrir_maestro("materiales"))
    g.esperar(1.6)  # maestro con el flujo del día

    maestro = ventana.vista("maestro:materiales")
    paso2 = maestro.contenidos["mantenimiento"].filas[1]
    g.clic(paso2, lambda: ventana.abrir_tarea("ampliacion_almacen"))
    tarea = ventana.vista("tarea:ampliacion_almacen")
    g.hasta(lambda: tarea.revision is not None and not tarea._revisando)
    g.esperar(1.6)  # Antes de ejecutar

    barra = tarea.pagina.verticalScrollBar()
    for _ in range(12):  # bajar suave hasta la vista previa con los avisos
        barra.setValue(barra.value() + 40)
        g.esperar(0.05)
    g.esperar(2.0)
    barra.setValue(0)
    g.esperar(0.4)

    revisar = tarea.findChildren(type(tarea.principal), options=Qt.FindChildOption.FindChildrenRecursively)
    volver_a_revisar = next(b for b in revisar if b.text() == "Volver a revisar" and b.isVisibleTo(tarea))
    g.clic(volver_a_revisar, tarea.revisar)
    g.hasta(lambda: not tarea._revisando)
    g.esperar(0.8)

    g.clic(tarea.principal, tarea.ejecutar)
    g.hasta(lambda: tarea.estado == "terminado" and not ventana.ocupada)
    g.esperar(1.8)  # resultado
    for _ in range(6):
        barra.setValue(barra.value() + 40)
        g.esperar(0.05)
    g.esperar(2.4)

    docs = RAIZ / "docs"
    docs.mkdir(exist_ok=True)
    g.cuadros[-1][0].save(docs / "captura.png", optimize=True)
    alto_gif = round(ANCHO_GIF * g.cuadros[0][0].height / g.cuadros[0][0].width)
    chicos = [img.resize((ANCHO_GIF, alto_gif), Image.LANCZOS) for img, _ in g.cuadros]
    duraciones = [d for _, d in g.cuadros]
    paleta = [c.convert("P", palette=Image.ADAPTIVE, colors=96) for c in chicos]
    paleta[0].save(docs / "demo.gif", save_all=True, append_images=paleta[1:], duration=duraciones, loop=0,
                   optimize=True, disposal=1)
    ventana.close()
    print(f"{len(g.cuadros)} cuadros, {sum(duraciones) / 1000:.1f} s -> docs/demo.gif "
          f"({(docs / 'demo.gif').stat().st_size // 1024} KB) y docs/captura.png")


if __name__ == "__main__":
    main()
