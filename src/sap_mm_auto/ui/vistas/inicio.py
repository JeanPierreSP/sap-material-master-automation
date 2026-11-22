"""Inicio: los maestros con el estado de su flujo de hoy, la conexión a SAP y el resumen del día.

Cada tarjeta dice lo que hace falta para decidir dónde entrar: qué cubre el maestro y,
si tiene un flujo del día, cuántos pasos van hechos. Los maestros sin tareas aparecen
apagados y con "Próximamente": están en el mapa, pero no compiten con el que se usa.
"""

from __future__ import annotations

from datetime import date, datetime
from typing import TYPE_CHECKING

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QFrame, QGridLayout, QWidget

from ... import catalogo
from ...historial import FALLO, HECHO, estado_del_dia
from .. import piezas, tema

if TYPE_CHECKING:
    from ..ventana import Ventana

ICONOS = {"materiales": "cubo", "proveedores": "camion", "clientes": "persona"}
COLUMNAS = 3


def saludo(ahora: datetime | None = None) -> str:
    hora = (ahora or datetime.now()).hour
    return "Buenos días" if hora < 12 else "Buenas tardes" if hora < 19 else "Buenas noches"


def estado_de_flujos(maestro: catalogo.Maestro, corridas, hoy: date | None = None) -> tuple[str, str]:
    """Texto y tono de la línea de estado de una tarjeta de maestro."""
    flujos = [p for p in maestro.procesos if p.encadenado]
    if not flujos:
        n = len(maestro.tareas)
        return f"{n} tarea{'s' if n != 1 else ''}", "neutro"
    p = flujos[0]
    estados = [estado_del_dia(corridas, t, hoy)[0] for t in p.tareas]
    hechos, fallos = estados.count(HECHO), estados.count(FALLO)
    if fallos:
        return f"{p.nombre}: {fallos} paso{'s' if fallos > 1 else ''} con error hoy", "error"
    if hechos == len(estados):
        return f"{p.nombre}: completo hoy", "ok"
    return f"{p.nombre}: {hechos} de {len(estados)} pasos hoy", "neutro"


class TarjetaMaestro(QFrame):
    def __init__(self, ventana: "Ventana", maestro: catalogo.Maestro) -> None:
        super().__init__()
        self.ventana, self.maestro = ventana, maestro
        disponible = maestro.disponible
        self.setProperty("rol", "tarjeta" if disponible else "tarjeta_apagada")
        self.setMinimumHeight(172)
        caja = piezas.columna(self, espacio=6, margenes=(20, 18, 20, 16))
        cabeza = piezas.fila(piezas.IconoEnCaja(ICONOS.get(maestro.id, "cubo"), 44, apagado=not disponible), None)
        if not disponible:
            cabeza.addWidget(piezas.Pastilla("Próximamente", "neutro"), 0, Qt.AlignmentFlag.AlignTop)
        caja.addLayout(cabeza)
        caja.addSpacing(6)
        nombre = piezas.etiqueta(maestro.nombre, "seccion")
        if not disponible:
            nombre.setStyleSheet(f"color: {tema.TINTA_2};")
        caja.addWidget(nombre)
        caja.addWidget(piezas.etiqueta(maestro.descripcion, "cuerpo", envolver=True))
        caja.addStretch(1)
        self.punto = piezas.Punto("neutro")
        self.estado = piezas.etiqueta("", "rotulo")
        pie = piezas.fila(self.punto, self.estado, None, espacio=6)
        if disponible:
            self.setCursor(Qt.CursorShape.PointingHandCursor)
            self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
            self.setAccessibleName(f"Abrir {maestro.nombre}")
            pie.addWidget(piezas.etiqueta("Abrir  ›", "rotulo"))
        else:
            self.punto.hide()
            self.estado.setText("Sin tareas por ahora")
        caja.addLayout(pie)

    def actualizar(self, corridas) -> None:
        if not self.maestro.disponible:
            return
        texto, tono = estado_de_flujos(self.maestro, corridas)
        self.estado.setText(texto)
        self.punto.poner(tono)

    def mouseReleaseEvent(self, evento) -> None:
        if self.maestro.disponible and evento.button() == Qt.MouseButton.LeftButton:
            self.ventana.abrir_maestro(self.maestro.id)

    def keyPressEvent(self, evento) -> None:
        if self.maestro.disponible and evento.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter, Qt.Key.Key_Space):
            self.ventana.abrir_maestro(self.maestro.id)
        else:
            super().keyPressEvent(evento)

    def focusInEvent(self, evento) -> None:
        self.setProperty("foco", True)
        tema.repintar(self)
        super().focusInEvent(evento)

    def focusOutEvent(self, evento) -> None:
        self.setProperty("foco", False)
        tema.repintar(self)
        super().focusOutEvent(evento)


class VistaInicio(piezas.Pagina):
    def __init__(self, ventana: "Ventana") -> None:
        super().__init__()
        self.ventana = ventana
        col = self.columna
        self.titulo = piezas.etiqueta("Datos maestros", "titulo")
        self.subtitulo = piezas.etiqueta("", "subtitulo")
        col.addWidget(self.titulo)
        col.addWidget(self.subtitulo)
        col.addSpacing(10)

        grilla = QGridLayout()
        grilla.setSpacing(16)
        self.tarjetas: dict[str, TarjetaMaestro] = {}
        for i, m in enumerate(catalogo.MAESTROS):
            t = TarjetaMaestro(ventana, m)
            self.tarjetas[m.id] = t
            grilla.addWidget(t, i // COLUMNAS, i % COLUMNAS)
        for c in range(COLUMNAS):
            grilla.setColumnStretch(c, 1)
        col.addLayout(grilla)
        col.addSpacing(10)

        abajo = piezas.fila(espacio=16)
        abajo.addWidget(self._tarjeta_sap(), 1)
        abajo.addWidget(self._tarjeta_hoy(), 1)
        col.addLayout(abajo)
        col.addStretch(1)

    def _tarjeta_sap(self) -> QWidget:
        t = piezas.tarjeta()
        caja = piezas.fila(piezas.IconoEnCaja("monitor", 40), 6, espacio=10)
        caja.setContentsMargins(18, 16, 18, 16)
        textos = piezas.columna(espacio=2)
        textos.addWidget(piezas.etiqueta("Conexión a SAP", "rotulo"))
        self.sap_punto = piezas.Punto("neutro")
        self.sap_estado = piezas.etiqueta("", "fuerte")
        textos.addLayout(piezas.fila(self.sap_punto, self.sap_estado, None, espacio=6))
        self.sap_detalle = piezas.etiqueta("", "cuerpo", envolver=True)
        textos.addWidget(self.sap_detalle)
        caja.addLayout(textos, 1)
        comprobar = piezas.boton("Comprobar", "texto", "refrescar", tema.MARCA_FUERTE)
        comprobar.clicked.connect(self._pintar_sap)
        caja.addWidget(comprobar, 0, Qt.AlignmentFlag.AlignTop)
        t.setLayout(caja)
        return t

    def _tarjeta_hoy(self) -> QWidget:
        t = piezas.tarjeta()
        caja = piezas.fila(piezas.IconoEnCaja("historial", 40), 6, espacio=10)
        caja.setContentsMargins(18, 16, 18, 16)
        textos = piezas.columna(espacio=2)
        textos.addWidget(piezas.etiqueta("Hoy", "rotulo"))
        self.hoy_resumen = piezas.etiqueta("", "fuerte")
        textos.addWidget(self.hoy_resumen)
        self.hoy_detalle = piezas.etiqueta("", "cuerpo", envolver=True)
        textos.addWidget(self.hoy_detalle)
        caja.addLayout(textos, 1)
        ver = piezas.boton("Ver historial", "texto", "derecha", tema.MARCA_FUERTE)
        ver.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        ver.clicked.connect(lambda: self.ventana.abrir_historial(periodo="hoy"))
        caja.addWidget(ver, 0, Qt.AlignmentFlag.AlignTop)
        t.setLayout(caja)
        return t

    # --- datos -------------------------------------------------------------------
    def al_mostrar(self) -> None:
        self.subtitulo.setText(f"{saludo()}. Elige un maestro para ver sus procesos y lo que falta hoy.")
        corridas = self.ventana.entorno.historial.leer()
        for t in self.tarjetas.values():
            t.actualizar(corridas)
        self._pintar_sap()
        hoy = [c for c in corridas if c.momento.date() == date.today()]
        con_error = sum(1 for c in hoy if c.con_problemas)
        if hoy:
            self.hoy_resumen.setText(f"{len(hoy)} corrida{'s' if len(hoy) != 1 else ''}")
            self.hoy_detalle.setText(f"{con_error} con error" if con_error else "Todas terminaron sin errores.")
        else:
            self.hoy_resumen.setText("Sin corridas todavía")
            self.hoy_detalle.setText("Lo que ejecutes hoy aparece aquí y en el historial.")

    def _pintar_sap(self) -> None:
        estado = self.ventana.refrescar_sap()
        self.sap_punto.poner("ok" if estado.conectado else "error")
        self.sap_estado.setText(estado.texto + (" · listo" if estado.conectado else ""))
        if self.ventana.entorno.demo:
            self.sap_detalle.setText("Modo demostración: las tareas corren contra un SAP en memoria, sin tocar ningún sistema.")
        elif estado.conectado:
            self.sap_detalle.setText("Se usa la sesión que ya abriste, con tu usuario y tus permisos.")
        else:
            self.sap_detalle.setText("Abre SAP GUI, inicia sesión y deja la pantalla en el menú inicial.")
