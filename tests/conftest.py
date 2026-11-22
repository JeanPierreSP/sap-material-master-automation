import sys
from pathlib import Path

import pytest

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
