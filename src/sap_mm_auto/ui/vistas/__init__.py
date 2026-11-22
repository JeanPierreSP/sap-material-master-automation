"""Pantallas de la ventana: inicio, maestro, tarea e historial."""

from __future__ import annotations

from .. import piezas


class Provisional(piezas.Pagina):
    """Pantalla de paso mientras la interfaz se arma por etapas."""

    def __init__(self, ventana, titulo: str) -> None:
        super().__init__()
        self.ventana = ventana
        volver = piezas.boton("Volver", "texto", "atras")
        volver.clicked.connect(ventana.volver)
        self.columna.addLayout(piezas.fila(volver, None))
        self.columna.addWidget(piezas.etiqueta(titulo, "titulo"))
        self.columna.addWidget(piezas.etiqueta("Esta pantalla llega en la próxima etapa.", "subtitulo"))
        self.columna.addStretch(1)
