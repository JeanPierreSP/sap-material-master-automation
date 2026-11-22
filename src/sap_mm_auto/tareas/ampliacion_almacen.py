"""Amplía un material existente a un almacén adicional (MM01)."""

from __future__ import annotations

from ..sesion import ErrorSAP, Pantalla
from . import campos
from .base import Fila, Tarea


class AmpliacionAlmacen(Tarea):
    id = "ampliacion_almacen"
    nombre = "Ampliación de almacén"
    descripcion = "Amplía cada material del archivo al centro y almacén indicados."
    transacciones = ("MM01",)
    columnas = ("MATERIAL", "CENTRO", "ALMACEN")

    def ejecutar(self, pantalla: Pantalla, fila: Fila) -> str:
        material = fila["MATERIAL"].strip()
        centro = fila["CENTRO"].strip()
        almacen = fila["ALMACEN"].strip()

        pantalla.ir_a("MM01")
        pantalla.escribir(campos.MATERIAL, material)
        pantalla.enter()
        pantalla.exigir_sin_error("Material")

        # Vistas: solo Almacén 1, que es la que pide el nivel almacén
        pantalla.marcar(campos.VISTA_ALMACEN_1)
        pantalla.enter(ventana=1)

        pantalla.escribir(campos.CENTRO, centro)
        pantalla.escribir(campos.ALMACEN, almacen)
        pantalla.enter(ventana=1)
        estado = pantalla.estado()
        if estado.es_error:
            raise ErrorSAP(f"Niveles organizativos: {estado.texto}")
        if "ya existe" in estado.texto.lower():
            # El material ya estaba ampliado: no es un error, pero no hay nada que grabar
            pantalla.volver_al_inicio()
            return f"Ya existía en {centro}/{almacen}"

        pantalla.guardar()
        return pantalla.exigir_sin_error("Grabar").texto
