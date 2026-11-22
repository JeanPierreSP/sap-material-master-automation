"""La interfaz completa una corrida en modo demostración (sin pantalla)."""

import os
import time

import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
QtWidgets = pytest.importorskip("PySide6.QtWidgets")


def test_corrida_demo_desde_la_interfaz():
    from sap_mm_auto.ui.ventana import Ventana

    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    v = Ventana()
    v.lista.setCurrentRow(0)  # Ampliación de almacén, con su archivo de ejemplo
    v._ejecutar()

    limite = time.time() + 30
    while not v.btn_ejecutar.isEnabled() and time.time() < limite:
        app.processEvents()
        time.sleep(0.02)

    assert v.btn_ejecutar.isEnabled(), "la corrida no terminó"
    assert v.cifras["ok"].text() == "5"
    assert v.cifras["errores"].text() == "3"
    assert "Terminada" in v.texto_progreso.text()
