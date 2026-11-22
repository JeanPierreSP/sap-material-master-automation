"""Cambia la descripción corta de un material (MM02, Datos base 1)."""

from __future__ import annotations

from ..sesion import ErrorSAP, Pantalla
from . import campos
from .base import Fila, Tarea

LARGO_MAXIMO = 40  # largo del campo MAKT-MAKTX en SAP


class CambioDescripcion(Tarea):
    id = "cambio_descripcion"
    nombre = "Cambio de descripción"
    descripcion = "Reemplaza la descripción corta de cada material del archivo."
    transacciones = ("MM02",)
    columnas = ("MATERIAL", "DESCRIPCION")

    def clave(self, fila: Fila) -> str:
        # La descripción es parte del trabajo: cambiarla de nuevo a otro texto sí cuenta
        return f"{fila['MATERIAL'].strip().upper()}|{fila['DESCRIPCION'].strip()}"

    def validar(self, fila: Fila) -> str | None:
        motivo = super().validar(fila)
        if motivo:
            return motivo
        if len(fila["DESCRIPCION"].strip()) > LARGO_MAXIMO:
            return f"La descripción supera {LARGO_MAXIMO} caracteres"
        return None

    def ejecutar(self, pantalla: Pantalla, fila: Fila) -> str:
        material = fila["MATERIAL"].strip()
        nueva = fila["DESCRIPCION"].strip()

        pantalla.ir_a("MM02")
        pantalla.escribir(campos.MATERIAL, material)
        pantalla.enter()
        pantalla.exigir_sin_error("Material")

        pantalla.marcar(campos.VISTA_DATOS_BASE_1)
        pantalla.enter(ventana=1)
        pantalla.exigir_sin_error("Vistas")

        anterior = pantalla.leer(campos.DESCRIPCION)
        if anterior == nueva:
            pantalla.volver_al_inicio()
            return "Sin cambios: ya tenía esa descripción"

        pantalla.escribir(campos.DESCRIPCION, nueva)
        pantalla.guardar()
        estado = pantalla.estado()
        if estado.es_error:
            raise ErrorSAP(f"Grabar: {estado.texto}")
        return f"{anterior} → {nueva}"
