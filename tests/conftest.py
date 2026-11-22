import os
import sys
import time
from pathlib import Path

import pytest

# Las pruebas de la interfaz corren sin pantalla (también en CI)
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from sap_mm_auto.controles import MemoriaDiaria  # noqa: E402
from sap_mm_auto.registro import Registro  # noqa: E402
from sap_mm_auto.sesion import Pantalla  # noqa: E402
from sap_mm_auto.simulador import SAPSimulado  # noqa: E402


@pytest.fixture
def sap():
    return SAPSimulado()


@pytest.fixture
def pantalla(sap):
    return Pantalla(sap)


@pytest.fixture
def memoria(tmp_path):
    return MemoriaDiaria(tmp_path)


@pytest.fixture
def registro(tmp_path):
    return Registro(tmp_path, "prueba")


# --------------------------------------------------------------------------
# Interfaz
# --------------------------------------------------------------------------
@pytest.fixture
def app():
    QtWidgets = pytest.importorskip("PySide6.QtWidgets")
    from sap_mm_auto.ui import tema

    aplicacion = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    tema.aplicar(aplicacion)
    return aplicacion


@pytest.fixture
def ventana(app, tmp_path):
    """Ventana en modo demostración, con su propia carpeta temporal y sin historial de ejemplo."""
    from sap_mm_auto.entorno import Entorno
    from sap_mm_auto.ui.ventana import Ventana

    v = Ventana(Entorno(demo=True, carpeta=tmp_path / "registro"))
    v.show()
    app.processEvents()
    yield v
    if v.corredor:
        v.corredor.detener()
        v.corredor.esperar()
    v.hide()
    v.deleteLater()
    app.processEvents()


def esperar(app, condicion, segundos: float = 15.0) -> bool:
    """Procesa eventos hasta que ``condicion()`` se cumpla o se acabe el tiempo."""
    limite = time.time() + segundos
    while time.time() < limite:
        app.processEvents()
        if condicion():
            return True
        time.sleep(0.01)
    app.processEvents()
    return condicion()
