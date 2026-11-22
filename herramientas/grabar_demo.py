"""Graba docs/demo.gif y docs/captura.png ejecutando la interfaz en modo demostración.

    QT_QPA_PLATFORM=offscreen python herramientas/grabar_demo.py
"""

from __future__ import annotations

import os
import sys
import time
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ / "src"))

from PIL import Image  # noqa: E402
from PySide6.QtWidgets import QApplication  # noqa: E402

from sap_mm_auto.ui.ventana import ESTILO, Ventana  # noqa: E402


def a_imagen(ventana: Ventana) -> Image.Image:
    qimg = ventana.grab().toImage().convertToFormat(ventana.grab().toImage().Format.Format_RGB888)
    ancho, alto = qimg.width(), qimg.height()
    datos = bytes(qimg.constBits())[: qimg.bytesPerLine() * alto]
    img = Image.frombuffer("RGB", (ancho, alto), datos, "raw", "RGB", qimg.bytesPerLine(), 1)
    return img.resize((ancho * 3 // 4, alto * 3 // 4), Image.LANCZOS)


def esperar(app: QApplication, segundos: float, cuadros: list, ventana: Ventana, cada: float = 0.12) -> None:
    fin = time.time() + segundos
    while time.time() < fin:
        app.processEvents()
        cuadros.append(a_imagen(ventana))
        time.sleep(cada)


def main() -> None:
    app = QApplication(sys.argv)
    app.setStyleSheet(ESTILO)
    os.chdir(RAIZ)  # ruta relativa en pantalla, sin rutas de la máquina que graba
    v = Ventana()
    v.show()
    v.ruta.setText("datos_ejemplo/ampliacion_almacen.csv")
    v.ruta.deselect()
    v.btn_revisar.setFocus()
    cuadros: list[Image.Image] = []

    esperar(app, 0.8, cuadros, v)
    v._revisar()
    esperar(app, 1.0, cuadros, v)
    v._ejecutar()
    limite = time.time() + 30
    while not v.btn_ejecutar.isEnabled() and time.time() < limite:
        esperar(app, 0.12, cuadros, v)
    esperar(app, 2.5, cuadros, v)

    docs = RAIZ / "docs"
    docs.mkdir(exist_ok=True)
    cuadros[-1].save(docs / "captura.png", optimize=True)
    paleta = [c.convert("P", palette=Image.ADAPTIVE, colors=128) for c in cuadros]
    paleta[0].save(docs / "demo.gif", save_all=True, append_images=paleta[1:], duration=120, loop=0, optimize=True)
    print(f"{len(cuadros)} cuadros → {docs / 'demo.gif'} ({(docs / 'demo.gif').stat().st_size // 1024} KB)")


if __name__ == "__main__":
    main()
