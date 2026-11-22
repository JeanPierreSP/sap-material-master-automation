"""Bloquea materiales con un status de material para todos los centros (MM02).

No borra nada: el bloqueo se revierte quitando el status.
"""

from __future__ import annotations

from ..sesion import ErrorSAP, Pantalla
from . import campos
from .base import Fila, Tarea


class BloqueoMaterial(Tarea):
    id = "bloqueo_material"
    nombre = "Bloqueo de material"
    descripcion = "Asigna el status de bloqueo indicado a cada material del archivo."
    transacciones = ("MM02",)
    columnas = ("MATERIAL", "STATUS")
    avisar_status = False  # poner o cambiar el status es justamente el trabajo de esta tarea

    def ejecutar(self, pantalla: Pantalla, fila: Fila) -> str:
        material = fila["MATERIAL"].strip()
        status = fila["STATUS"].strip()

        pantalla.ir_a("MM02")
        pantalla.escribir(campos.MATERIAL, material)
        pantalla.enter()
        pantalla.exigir_sin_error("Material")

        pantalla.marcar(campos.VISTA_DATOS_BASE_1)
        pantalla.enter(ventana=1)
        pantalla.exigir_sin_error("Vistas")

        if pantalla.leer(campos.STATUS_MATERIAL) == status:
            pantalla.volver_al_inicio()
            return f"Ya tenía el status {status}"

        pantalla.escribir(campos.STATUS_MATERIAL, status)
        pantalla.enter()
        estado = pantalla.estado()
        if estado.es_error:
            raise ErrorSAP(f"Status {status}: {estado.texto}")

        pantalla.guardar()
        return pantalla.exigir_sin_error("Grabar").texto
