"""SAP simulado.

Imita la parte de SAP GUI Scripting que usan las tareas (``findById``, ``sendVKey``,
barra de estado y ``Info.Transaction``) sobre un maestro de materiales en memoria.
Sirve para dos cosas:

- probar cada cambio sin tocar un SAP real (``tests/``);
- mostrar la aplicación en modo demostración.

No pretende reproducir SAP: solo las pantallas y mensajes que necesitan las tareas,
incluidos los errores típicos (material inexistente, bloqueado por otro usuario,
almacén no definido) y una pantalla inesperada para probar el freno.
"""

from __future__ import annotations

import time
from dataclasses import dataclass, field
from typing import Any

from .tareas import campos

CENTROS = {"1000": {"0001", "0002", "0003"}, "1100": {"0001", "0002"}}
STATUS_VALIDOS = {"01", "02", "03"}

INICIO, INICIAL, VISTAS, ORG, DATOS, AVISO = "inicio", "inicial", "vistas", "org", "datos", "aviso"


@dataclass
class Material:
    descripcion: str
    almacenes: set[tuple[str, str]] = field(default_factory=set)
    status: str = ""


def materiales_de_ejemplo() -> dict[str, Material]:
    return {
        "DEMO-0001": Material("TORNILLO HEXAGONAL M8 X 30", {("1000", "0001")}),
        "DEMO-0002": Material("ARANDELA PLANA 8 MM", {("1000", "0001")}),
        "DEMO-0003": Material("CABLE THW 2.5 MM2 NEGRO", {("1000", "0001"), ("1000", "0002")}),
        "DEMO-0004": Material("TUERCA HEXAGONAL M8", {("1000", "0001")}),
        "DEMO-0005": Material("GUANTE DE NITRILO TALLA M", {("1000", "0001")}),
        "DEMO-0006": Material("CINTA AISLANTE 19 MM", {("1000", "0001")}),
        "DEMO-0007": Material("BROCA HSS 6 MM", {("1100", "0001")}),
        "DEMO-0008": Material("DISCO DE CORTE 4.5 PULG", {("1000", "0001")}),
        "DEMO-0009": Material("LIJA AL AGUA N 120", {("1000", "0001")}),
        "DEMO-0010": Material("SILICONA TRANSPARENTE 280 ML", {("1000", "0001")}),
    }


class _Info:
    def __init__(self, sim: "SAPSimulado") -> None:
        self._sim = sim

    @property
    def Transaction(self) -> str:  # noqa: N802 (nombre de la API de SAP)
        return self._sim.transaccion


class _Elemento:
    """Un control de pantalla. Lee y escribe en los campos del simulador."""

    def __init__(self, sim: "SAPSimulado", id_: str) -> None:
        self._sim, self._id = sim, id_

    @property
    def text(self) -> str:
        if self._id == "wnd[0]/sbar":
            return self._sim.mensaje
        return self._sim.campos.get(self._id, "")

    @text.setter
    def text(self, valor: str) -> None:
        self._sim.campos[self._id] = valor

    @property
    def selected(self) -> bool:
        return self._id in self._sim.marcados

    @selected.setter
    def selected(self, valor: bool) -> None:
        (self._sim.marcados.add if valor else self._sim.marcados.discard)(self._id)

    @property
    def messageType(self) -> str:  # noqa: N802
        return self._sim.tipo_mensaje

    def sendVKey(self, tecla: int) -> None:  # noqa: N802
        self._sim.tecla(self._id, tecla)


class SAPSimulado:
    """Sesión de SAP falsa. Se usa igual que la real: ``Pantalla(SAPSimulado())``."""

    def __init__(
        self,
        materiales: dict[str, Material] | None = None,
        bloqueados_por_otro: set[str] | None = None,
        pantalla_inesperada: set[str] | None = None,
        retardo: float = 0.0,
    ) -> None:
        self.materiales = materiales if materiales is not None else materiales_de_ejemplo()
        self.bloqueados_por_otro = bloqueados_por_otro or set()
        self.pantalla_inesperada = pantalla_inesperada or set()
        self.retardo = retardo
        self.Info = _Info(self)
        self.transaccion = "SESSION_MANAGER"
        self.paso = INICIO
        self.campos: dict[str, str] = {}
        self.marcados: set[str] = set()
        self.tipo_mensaje, self.mensaje = "", ""
        self.grabaciones: list[tuple[str, str]] = []  # (transacción, material) grabados

    # --- API de SAP GUI Scripting ------------------------------------------
    def findById(self, id_: str) -> Any:  # noqa: N802
        if id_ not in self._controles_visibles():
            raise LookupError(f"El control {id_} no existe en la pantalla actual")
        return _Elemento(self, id_)

    # --- lógica de pantallas ---------------------------------------------
    def _controles_visibles(self) -> set[str]:
        base = {"wnd[0]", "wnd[0]/tbar[0]/okcd", "wnd[0]/sbar"}
        if self.paso == INICIAL:
            return base | {campos.MATERIAL}
        if self.paso == VISTAS:
            return base | {"wnd[1]", campos.VISTAS_TABLA, campos.VISTA_DATOS_BASE_1, campos.VISTA_ALMACEN_1}
        if self.paso == ORG:
            return base | {"wnd[1]", campos.CENTRO, campos.ALMACEN}
        if self.paso == DATOS and self.transaccion == "MM02":
            return base | {campos.DESCRIPCION, campos.STATUS_MATERIAL}
        if self.paso == AVISO:
            return base | {"wnd[1]"}
        return base

    def tecla(self, ventana: str, tecla: int) -> None:
        if self.retardo:
            time.sleep(self.retardo)
        if tecla == 11:
            self._grabar()
            return
        comando = self.campos.pop("wnd[0]/tbar[0]/okcd", "").strip()
        if comando:
            self._comando(comando)
        elif ventana == "wnd[1]":
            self._enter_ventana()
        else:
            self._enter()

    def _msg(self, tipo: str = "", texto: str = "") -> None:
        self.tipo_mensaje, self.mensaje = tipo, texto

    def _comando(self, comando: str) -> None:
        self._msg()
        destino = comando[2:].upper() if comando.lower().startswith("/n") else comando.upper()
        if destino == "":
            self.transaccion, self.paso = "SESSION_MANAGER", INICIO
        elif destino in {"MM01", "MM02"}:
            self.transaccion, self.paso = destino, INICIAL
        else:
            self._msg("E", f"La transacción {destino} no existe")
            return
        self.campos.clear()
        self.marcados.clear()

    def _material(self) -> str:
        return self.campos.get(campos.MATERIAL, "").strip().upper()

    def _enter(self) -> None:
        self._msg()
        m = self._material()
        if self.paso == INICIAL:
            if m not in self.materiales:
                self._msg("E", f"El material {m} no existe")
            elif m in self.pantalla_inesperada:
                self.paso = AVISO  # aparece una ventana que el script no conoce
            elif self.transaccion == "MM02" and m in self.bloqueados_por_otro:
                self._msg("E", f"El material {m} está bloqueado por el usuario OTRO_USUARIO")
            else:
                self.paso = VISTAS
        elif self.paso == DATOS and self.transaccion == "MM02":
            status = self.campos.get(campos.STATUS_MATERIAL, "").strip()
            if status and status not in STATUS_VALIDOS:
                self._msg("E", f"El status de material {status} no está definido")

    def _enter_ventana(self) -> None:
        self._msg()
        m = self._material()
        if self.paso == VISTAS:
            if self.transaccion == "MM01" and campos.VISTA_ALMACEN_1 in self.marcados:
                self.paso = ORG
            elif self.transaccion == "MM02" and campos.VISTA_DATOS_BASE_1 in self.marcados:
                mat = self.materiales[m]
                self.campos[campos.DESCRIPCION] = mat.descripcion
                self.campos[campos.STATUS_MATERIAL] = mat.status
                self.paso = DATOS
            else:
                self._msg("E", "Seleccione al menos una vista válida")
        elif self.paso == ORG:
            centro = self.campos.get(campos.CENTRO, "").strip()
            almacen = self.campos.get(campos.ALMACEN, "").strip()
            if centro not in CENTROS:
                self._msg("E", f"El centro {centro} no existe")
            elif almacen not in CENTROS[centro]:
                self._msg("E", f"El almacén {almacen} no está definido para el centro {centro}")
            elif (centro, almacen) in self.materiales[m].almacenes:
                self._msg("I", f"El material {m} ya existe en el centro {centro}, almacén {almacen}")
            else:
                self.paso = DATOS

    def _grabar(self) -> None:
        m = self._material()
        if self.paso != DATOS:
            self._msg("E", "No hay datos para grabar")
            return
        if self.transaccion == "MM01":
            self.materiales[m].almacenes.add((self.campos[campos.CENTRO].strip(), self.campos[campos.ALMACEN].strip()))
            texto = f"Se ha ampliado el material {m}"
        else:
            mat = self.materiales[m]
            mat.descripcion = self.campos.get(campos.DESCRIPCION, mat.descripcion)
            mat.status = self.campos.get(campos.STATUS_MATERIAL, mat.status)
            texto = f"Se ha modificado el material {m}"
        self.grabaciones.append((self.transaccion, m))
        # Como SAP, vuelve a la pantalla inicial de la transacción
        self.paso = INICIAL
        self.campos = {}
        self.marcados.clear()
        self._msg("S", texto)
