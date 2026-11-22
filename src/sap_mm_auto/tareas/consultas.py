"""Consultas de solo lectura en MM03.

Las usa la revisión previa para avisar, antes de ejecutar, de materiales que no existen
o que tienen un status de bloqueo. MM03 es de visualización: no hay forma de grabar.
"""

from __future__ import annotations

from ..sesion import ErrorSAP, Pantalla
from . import campos


def status_material(pantalla: Pantalla, material: str) -> str:
    """Devuelve el status del material ("" si no tiene). ``ErrorSAP`` si no existe.

    Siempre deja SAP en el menú inicial, también cuando la consulta falla.
    """
    pantalla.ir_a("MM03")
    try:
        pantalla.escribir(campos.MATERIAL, material.strip())
        pantalla.enter()
        pantalla.exigir_sin_error("Material")
        pantalla.marcar(campos.VISTA_DATOS_BASE_1)
        pantalla.enter(ventana=1)
        pantalla.exigir_sin_error("Vistas")
        return pantalla.leer(campos.STATUS_MATERIAL).strip()
    finally:
        try:
            pantalla.volver_al_inicio()
        except ErrorSAP:
            pass  # el error que importa es el de la consulta, no el de volver
