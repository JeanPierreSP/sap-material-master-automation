"""Acceso a SAP GUI Scripting.

La conexión real toma la sesión que el usuario ya abrió: no inicia sesión ni guarda
credenciales, y todo queda registrado con su usuario y sus permisos.

El resto del proyecto trabaja contra la clase ``Pantalla``. Así, las pruebas y el
modo demostración pueden usar el SAP simulado (``simulador.py``) sin cambiar nada.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Protocol


class ErrorSAP(Exception):
    """SAP respondió con un error o no se pudo completar un paso."""


class PantallaInesperada(ErrorSAP):
    """SAP no está en la pantalla que el script esperaba. Detiene la corrida."""


class SesionSAP(Protocol):
    """Lo mínimo que se usa de un objeto GuiSession de SAP GUI Scripting."""

    Info: Any

    def findById(self, id: str) -> Any: ...  # noqa: N802 (nombre de la API de SAP)


def conectar_sesion_activa(conexion: int = 0, sesion: int = 0) -> SesionSAP:
    """Devuelve una sesión ya abierta de SAP GUI (solo Windows)."""
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:  # pragma: no cover - depende del sistema operativo
        raise ErrorSAP("pywin32 no está instalado; la conexión real solo funciona en Windows") from exc

    try:  # pragma: no cover - requiere SAP GUI
        sapgui = win32com.client.GetObject("SAPGUI")
        motor = sapgui.GetScriptingEngine
    except Exception as exc:  # pragma: no cover
        raise ErrorSAP("SAP GUI no está abierto o el scripting está deshabilitado") from exc

    if motor.Children.Count <= conexion:  # pragma: no cover
        raise ErrorSAP("No hay una conexión abierta en SAP GUI")
    con = motor.Children(conexion)
    if con.Children.Count <= sesion:  # pragma: no cover
        raise ErrorSAP("La conexión no tiene sesiones abiertas")
    return con.Children(sesion)


@dataclass(frozen=True)
class Estado:
    """Mensaje de la barra de estado de SAP."""

    tipo: str  # S = éxito, W = aviso, I = información, E = error, A = cancelación, "" = sin mensaje
    texto: str

    @property
    def es_error(self) -> bool:
        return self.tipo in {"E", "A"}


class Pantalla:
    """Operaciones de alto nivel sobre una sesión de SAP GUI."""

    INICIO = "SESSION_MANAGER"

    def __init__(self, sesion: SesionSAP) -> None:
        self._s = sesion

    # --- navegación ---------------------------------------------------------
    def ir_a(self, transaccion: str) -> None:
        self._s.findById("wnd[0]/tbar[0]/okcd").text = f"/n{transaccion}"
        self.enter()
        if self.transaccion_actual() != transaccion.upper():
            raise PantallaInesperada(f"No se pudo entrar a {transaccion}: {self.estado().texto}")

    def volver_al_inicio(self) -> None:
        self._s.findById("wnd[0]/tbar[0]/okcd").text = "/n"
        self.enter()

    def transaccion_actual(self) -> str:
        return str(self._s.Info.Transaction)

    def esta_en_inicio(self) -> bool:
        return self.transaccion_actual() == self.INICIO

    # --- campos -------------------------------------------------------------
    def escribir(self, campo: str, valor: object) -> None:
        self._elemento(campo).text = str(valor)

    def leer(self, campo: str) -> str:
        return str(self._elemento(campo).text)

    def marcar(self, campo: str, valor: bool = True) -> None:
        self._elemento(campo).selected = valor

    def existe(self, campo: str) -> bool:
        try:
            self._s.findById(campo)
        except Exception:  # SAP lanza com_error; el simulador, LookupError
            return False
        return True

    # --- teclas -------------------------------------------------------------
    def enter(self, ventana: int = 0) -> None:
        self._elemento(f"wnd[{ventana}]").sendVKey(0)

    def guardar(self) -> None:
        self._elemento("wnd[0]").sendVKey(11)

    # --- resultado ----------------------------------------------------------
    def estado(self) -> Estado:
        barra = self._elemento("wnd[0]/sbar")
        return Estado(str(barra.messageType), str(barra.text))

    def exigir_sin_error(self, paso: str) -> Estado:
        """Lanza ErrorSAP si la barra de estado trae un error después de ``paso``."""
        estado = self.estado()
        if estado.es_error:
            raise ErrorSAP(f"{paso}: {estado.texto}")
        return estado

    def _elemento(self, campo: str) -> Any:
        try:
            return self._s.findById(campo)
        except Exception as exc:
            raise PantallaInesperada(f"No se encontró el campo {campo}") from exc
