"""Pantalla de una tarea: revisar, ejecutar y ver el resultado, en el mismo lugar.

Tres estados en la zona central:

- listo: "Antes de ejecutar" (cada chequeo con su resultado) y la vista previa del archivo,
  con los avisos marcados en su fila.
- corriendo: "3 de 8 · quedan 5", tiempo estimado, cronómetro, barra por fila y log en vivo.
- terminado: las cifras (OK, omitidas, con error), accesos al log y a las rechazadas, y el log.

Abajo, fijo, el pie con EL botón de la pantalla (Ejecutar, Detener, o el paso siguiente) y,
en los pasos de un flujo encadenado, el interruptor de flujo automático: prendido, al
terminar bien un paso se abre el siguiente, se revisa y, si no hay avisos, se ejecuta solo.
"""

from __future__ import annotations

import time
from pathlib import Path
from typing import TYPE_CHECKING

from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QColor, QFont, QIcon
from PySide6.QtWidgets import (
    QAbstractItemView,
    QFileDialog,
    QFrame,
    QHeaderView,
    QLabel,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QWidget,
)

from ... import catalogo
from ...entorno import estimar_restante
from ...historial import HECHO, Corrida, estado_del_dia, texto_duracion
from ...tareas import TAREAS, Tarea
from ...vista_previa import AVISO, BLOQUEA, OK, AntesDeEjecutar, Chequeo
from .. import glifos, piezas, tema
from ..trabajador import en_segundo_plano
from .historial import abrir_archivo

if TYPE_CHECKING:
    from ..ventana import Ventana

FILAS_VISTA_PREVIA = 10
PAUSA_FLUJO_MS = 1200  # tiempo para leer qué paso sigue antes de que arranque solo
ICONO_CHEQUEO = {OK: ("check", tema.OK), AVISO: ("alerta", tema.AVISO), BLOQUEA: ("cruz", tema.ERROR)}


def reloj(segundos: float) -> str:
    s = int(segundos)
    return f"{s // 3600}:{s % 3600 // 60:02d}:{s % 60:02d}" if s >= 3600 else f"{s // 60:02d}:{s % 60:02d}"


class PasosFlujo(QWidget):
    """Los pasos del flujo en una línea; el actual resaltado y los hechos hoy con ✓."""

    def __init__(self, ventana: "Ventana", proceso: catalogo.Proceso, actual: str) -> None:
        super().__init__()
        self.ventana, self.proceso, self.actual = ventana, proceso, actual
        self.caja = piezas.fila(espacio=6)
        self.setLayout(self.caja)
        self.botones: dict[str, QPushButton] = {}
        for i, tarea_id in enumerate(proceso.tareas):
            if i:
                self.caja.addWidget(piezas.etiqueta("›", "rotulo"))
            b = piezas.boton(f"{i + 1}  {TAREAS[tarea_id].nombre}", "segmento")
            b.setCheckable(True)
            b.setChecked(tarea_id == actual)
            b.clicked.connect(lambda _=False, t=tarea_id: self._ir(t))
            self.botones[tarea_id] = b
            self.caja.addWidget(b)
        self.caja.addStretch(1)

    def _ir(self, tarea_id: str) -> None:
        self.botones[tarea_id].setChecked(tarea_id == self.actual)
        if tarea_id != self.actual:
            self.ventana.abrir_tarea(tarea_id)

    def actualizar(self, corridas: list[Corrida]) -> None:
        for tarea_id, b in self.botones.items():
            hecho = estado_del_dia(corridas, tarea_id)[0] == HECHO
            b.setIcon(glifos.icono("check", tema.OK, 14) if hecho else QIcon())
            b.setToolTip("Hecho hoy" if hecho else "")
            b.setChecked(tarea_id == self.actual)


class FichaChequeo(QWidget):
    def __init__(self, c: Chequeo) -> None:
        super().__init__()
        self.chequeo = c
        nombre_icono, color = ICONO_CHEQUEO[c.estado]
        icono = QLabel()
        icono.setPixmap(glifos.pixmap(nombre_icono, color, 16, 2.0))
        icono.setFixedWidth(20)
        nombre = piezas.etiqueta(c.nombre, "cuerpo")
        nombre.setStyleSheet(f"color: {tema.TINTA};")
        detalle = piezas.etiqueta(c.detalle, "rotulo", envolver=True)
        if c.estado != OK:
            detalle.setStyleSheet(f"color: {color};")
        caja = piezas.fila(icono, espacio=8)
        caja.setContentsMargins(0, 3, 0, 3)
        nombre.setFixedWidth(230)
        caja.addWidget(nombre)
        caja.addWidget(detalle, 1)
        self.setLayout(caja)


class VistaTarea(QWidget):
    def __init__(self, ventana: "Ventana", tarea: Tarea) -> None:
        super().__init__()
        self.ventana, self.tarea = ventana, tarea
        self.maestro, self.proceso = catalogo.ubicar(tarea.id)
        self.archivo: Path | None = ventana.entorno.archivo_de_ejemplo(tarea) if ventana.entorno.demo else None
        self.revision: AntesDeEjecutar | None = None
        self.estado = "revisando"
        self.corrida: Corrida | None = None
        self._inicio = 0.0
        self._actual = self._total = 0
        self._auto_pendiente = False
        self._revisando = False
        self._otra_vez = False
        self._archivo_revisado: Path | None = None

        caja = piezas.columna(self, espacio=0)
        self.pagina = piezas.Pagina(margenes=(32, 24, 32, 24))
        caja.addWidget(self.pagina, 1)
        caja.addWidget(self._armar_pie())
        col = self.pagina.columna

        volver = piezas.boton(self.maestro.nombre, "texto", "atras", tema.MARCA_FUERTE)
        volver.clicked.connect(lambda: ventana.abrir_maestro(self.maestro.id, self.proceso.id))
        col.addLayout(piezas.fila(volver, None))
        col.addWidget(piezas.etiqueta(f"{self.maestro.nombre} · {self.proceso.nombre}", "rotulo"))
        titulo = piezas.fila(piezas.etiqueta(tarea.nombre, "titulo"), 6, espacio=10)
        titulo.addWidget(piezas.Pastilla(" · ".join(tarea.transacciones), "marca"), 0, Qt.AlignmentFlag.AlignVCenter)
        titulo.addStretch(1)
        col.addLayout(titulo)
        col.addWidget(piezas.etiqueta(tarea.descripcion, "subtitulo", envolver=True))
        self.pasos = PasosFlujo(ventana, self.proceso, tarea.id) if self.proceso.encadenado else None
        if self.pasos:
            col.addWidget(self.pasos)
        col.addSpacing(4)

        self.aviso_flujo = piezas.etiqueta("", "cuerpo", envolver=True)
        self.aviso_flujo.hide()
        col.addWidget(self.aviso_flujo)

        # Tres bloques que se muestran de a uno (un QStackedWidget tomaría el alto del más alto)
        self.bloques = {"listo": self._armar_listo(), "corriendo": self._armar_corriendo(),
                        "terminado": self._armar_terminado()}
        for bloque in self.bloques.values():
            col.addWidget(bloque)

        self.log = QPlainTextEdit()
        self.log.setObjectName("log")
        self.log.setReadOnly(True)
        self.log.setFont(tema.fuente(9, mono=True))
        self.log.setMinimumHeight(240)
        self.log.setMaximumBlockCount(5000)
        self.log.setAccessibleName("Log de la corrida")
        self.titulo_log = piezas.etiqueta("Log de la corrida", "seccion")
        col.addWidget(self.titulo_log)
        col.addWidget(self.log)
        col.addStretch(1)

        self._reloj = QTimer(self)
        self._reloj.setInterval(1000)
        self._reloj.timeout.connect(self._pintar_avance)
        self._mostrar_zona("listo")

    # ======================================================================
    # Armado
    # ======================================================================
    def _armar_pie(self) -> QWidget:
        pie = QFrame()
        pie.setObjectName("barra")
        pie.setStyleSheet(f"QFrame#barra {{ border-bottom: none; border-top: 1px solid {tema.BORDE}; }}")
        externo = piezas.fila(None, espacio=0)
        externo.setContentsMargins(32, 12, 32, 12)
        pie.setLayout(externo)
        interior = QWidget()
        interior.setMaximumWidth(tema.ANCHO_CONTENIDO - 64)
        caja = piezas.fila(espacio=12)
        interior.setLayout(caja)
        externo.addWidget(interior, 100)
        externo.addStretch(1)

        self.automatico = piezas.Interruptor("Flujo automático")
        self.texto_automatico = piezas.etiqueta("", "rotulo", envolver=True)
        if self.proceso.encadenado:
            textos = piezas.columna(espacio=0)
            textos.addWidget(piezas.etiqueta("Flujo automático", "fuerte"))
            textos.addWidget(self.texto_automatico)
            caja.addWidget(self.automatico)
            caja.addLayout(textos, 1)
            siguiente = self.proceso.siguiente(self.tarea.id)
            self.texto_automatico.setText(
                f"Al terminar bien, sigue solo con {TAREAS[siguiente].nombre}." if siguiente
                else "Último paso del flujo: al terminar no sigue nada."
            )
            self.automatico.toggled.connect(self._al_cambiar_automatico)
        else:
            self.automatico.hide()
            caja.addStretch(1)

        self.secundario = piezas.boton("Volver a revisar")
        self.secundario.clicked.connect(self.volver_a_revisar)
        self.principal = piezas.boton("Ejecutar", "principal", "play", tema.SUPERFICIE)
        self.principal.setMinimumWidth(200)
        self.principal.clicked.connect(self._al_principal)
        caja.addWidget(self.secundario)
        caja.addWidget(self.principal)
        return pie

    def _armar_listo(self) -> QWidget:
        w = QWidget()
        caja = piezas.columna(w, espacio=14)

        # Archivo de entrada
        archivo = piezas.tarjeta()
        fila_archivo = piezas.fila(piezas.IconoEnCaja("archivo", 36), 4, espacio=10)
        fila_archivo.setContentsMargins(16, 12, 16, 12)
        textos = piezas.columna(espacio=0)
        textos.addWidget(piezas.etiqueta("Archivo de entrada", "rotulo"))
        self.nombre_archivo = piezas.etiqueta("", "fuerte")
        textos.addWidget(self.nombre_archivo)
        textos.addWidget(piezas.etiqueta("Columnas: " + ", ".join(self.tarea.columnas), "rotulo"))
        fila_archivo.addLayout(textos, 1)
        self.boton_ver_archivo = piezas.boton("Abrir", "texto", "externa", tema.MARCA_FUERTE)
        self.boton_ver_archivo.clicked.connect(lambda: abrir_archivo(str(self.archivo or "")))
        cambiar = piezas.boton("Elegir otro…", "texto", "carpeta", tema.MARCA_FUERTE)
        cambiar.clicked.connect(self._elegir_archivo)
        fila_archivo.addWidget(self.boton_ver_archivo)
        fila_archivo.addWidget(cambiar)
        archivo.setLayout(fila_archivo)
        caja.addWidget(archivo)

        # Antes de ejecutar
        antes = piezas.tarjeta()
        col = piezas.columna(antes, espacio=6, margenes=(18, 14, 18, 14))
        self.resumen_revision = piezas.Pastilla("Revisando…", "neutro")
        revisar = piezas.boton("Volver a revisar", "texto", "refrescar", tema.MARCA_FUERTE)
        revisar.clicked.connect(self.revisar)
        col.addLayout(piezas.fila(piezas.etiqueta("Antes de ejecutar", "seccion"), self.resumen_revision, None, revisar))
        col.addWidget(piezas.separador())
        self.lista_chequeos = piezas.columna(espacio=0)
        col.addLayout(self.lista_chequeos)
        caja.addWidget(antes)

        # Vista previa
        previa = piezas.tarjeta()
        col = piezas.columna(previa, espacio=8, margenes=(18, 14, 18, 14))
        self.titulo_previa = piezas.etiqueta("Vista previa", "seccion")
        self.detalle_previa = piezas.etiqueta("", "rotulo")
        col.addLayout(piezas.fila(self.titulo_previa, self.detalle_previa, None))
        self.tabla = QTableWidget()
        self.tabla.verticalHeader().hide()
        self.tabla.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.tabla.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
        self.tabla.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.tabla.setShowGrid(False)
        self.tabla.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.tabla.setFont(tema.fuente(9.5))
        col.addWidget(self.tabla)
        self.mas_filas = piezas.etiqueta("", "rotulo", envolver=True)
        col.addWidget(self.mas_filas)
        caja.addWidget(previa)
        return w

    def _armar_corriendo(self) -> QWidget:
        t = piezas.tarjeta()
        col = piezas.columna(t, espacio=8, margenes=(20, 16, 20, 18))
        self.rotulo_corriendo = piezas.etiqueta(f"Ejecutando · {self.tarea.nombre}", "rotulo")
        self.cronometro = piezas.etiqueta("00:00", "fuerte")
        self.cronometro.setFont(tema.fuente(13, QFont.Weight.DemiBold, mono=True))
        col.addLayout(piezas.fila(self.rotulo_corriendo, None, self.cronometro))
        self.contador = piezas.etiqueta("", "avance")
        self.quedan = piezas.etiqueta("", "subtitulo")
        fila = piezas.fila(self.contador, 8, espacio=4)
        fila.addWidget(self.quedan, 0, Qt.AlignmentFlag.AlignBottom)
        fila.addStretch(1)
        col.addLayout(fila)
        self.barra = piezas.BarraPorFila()
        col.addWidget(self.barra)
        self.ahora = piezas.etiqueta("", "rotulo")
        self.cuenta = piezas.etiqueta("", "rotulo")
        col.addLayout(piezas.fila(self.ahora, None, self.cuenta))
        return t

    def _armar_terminado(self) -> QWidget:
        w = QWidget()
        col = piezas.columna(w, espacio=12)
        cabeza = piezas.tarjeta()
        c = piezas.columna(cabeza, espacio=6, margenes=(20, 16, 20, 16))
        self.titulo_resultado = piezas.etiqueta("", "seccion")
        self.titulo_resultado.setFont(tema.fuente(14, QFont.Weight.DemiBold))
        self.detalle_resultado = piezas.etiqueta("", "cuerpo", envolver=True)
        c.addWidget(self.titulo_resultado)
        c.addWidget(self.detalle_resultado)
        botones = piezas.fila(espacio=6)
        self.boton_log = piezas.boton("Ver log", "texto", "log", tema.MARCA_FUERTE)
        self.boton_log.clicked.connect(lambda: self.corrida and abrir_archivo(self.corrida.log))
        self.boton_rechazadas = piezas.boton("Filas rechazadas (CSV)", "texto", "archivo", tema.MARCA_FUERTE)
        self.boton_rechazadas.clicked.connect(lambda: self.corrida and abrir_archivo(self.corrida.rechazadas))
        en_historial = piezas.boton("Ver en el historial", "texto", "historial", tema.MARCA_FUERTE)
        en_historial.clicked.connect(lambda: self.corrida and self.ventana.abrir_historial().mostrar_corrida(self.corrida))
        botones.addWidget(self.boton_log)
        botones.addWidget(self.boton_rechazadas)
        botones.addWidget(en_historial)
        botones.addStretch(1)
        c.addLayout(botones)
        col.addWidget(cabeza)
        cifras = piezas.fila(espacio=12)
        self.cifras = {
            "ok": piezas.Cifra("Procesadas OK", "ok"),
            "omitidas": piezas.Cifra("Omitidas (ya hechas hoy)", "neutro"),
            "errores": piezas.Cifra("Con error", "error"),
        }
        for cifra in self.cifras.values():
            cifras.addWidget(cifra, 1)
        col.addLayout(cifras)
        return w

    # ======================================================================
    # Estados
    # ======================================================================
    def _mostrar_zona(self, estado: str) -> None:
        self.estado = estado
        visible = "listo" if estado == "revisando" else estado
        for nombre, bloque in self.bloques.items():
            bloque.setVisible(nombre == visible)
        con_log = estado in ("corriendo", "terminado")
        self.log.setVisible(con_log)
        self.titulo_log.setVisible(con_log)
        self._pintar_pie()

    def _pintar_pie(self) -> None:
        ocupada_por_otra = self.ventana.ocupada and self.estado != "corriendo"
        siguiente = self.proceso.siguiente(self.tarea.id)
        self.secundario.setVisible(self.estado == "terminado")
        if self.estado == "corriendo":
            self._estilo_principal("Detener", "detener", "stop", tema.ERROR)
            self.principal.setEnabled(not self.ventana.corredor or not self.ventana.corredor.deteniendo)
            if self.ventana.corredor and self.ventana.corredor.deteniendo:
                self.principal.setText("Deteniendo…")
        elif self.estado == "terminado" and siguiente:
            self._estilo_principal(f"Ir a {TAREAS[siguiente].nombre}", "principal", "derecha", tema.SUPERFICIE)
            self.principal.setEnabled(True)
        elif self.estado == "terminado":
            self._estilo_principal("Volver a revisar", "principal", "refrescar", tema.SUPERFICIE)
            self.principal.setEnabled(True)
            self.secundario.hide()
        else:
            self._estilo_principal("Ejecutar", "principal", "play", tema.SUPERFICIE)
            self.principal.setEnabled(self.lista_para_ejecutar and not ocupada_por_otra)
            if ocupada_por_otra:
                self.principal.setToolTip("Hay otra corrida en curso: espera a que termine.")
            elif self.revision and not self.revision.puede_ejecutar:
                self.principal.setToolTip("Corrige lo marcado en Antes de ejecutar.")
            else:
                self.principal.setToolTip("")

    def _estilo_principal(self, texto: str, variante: str, icono: str, color: str) -> None:
        self.principal.setText(texto)
        self.principal.setIcon(glifos.icono(icono, color, 16))
        if self.principal.property("variante") != variante:
            self.principal.setProperty("variante", variante)
            tema.repintar(self.principal)

    def _al_principal(self) -> None:
        siguiente = self.proceso.siguiente(self.tarea.id)
        if self.estado == "corriendo":
            self.detener()
        elif self.estado == "terminado" and siguiente:
            self.ventana.abrir_tarea(siguiente)
        elif self.estado == "terminado":
            self.volver_a_revisar()
        else:
            self.ejecutar()

    # ======================================================================
    # Revisar
    # ======================================================================
    def al_mostrar(self) -> None:
        self._sincronizar_automatico()
        if self.pasos:
            self.pasos.actualizar(self.ventana.entorno.historial.leer())
        if self.estado in ("revisando", "listo"):
            self.revisar()
        else:
            self._pintar_pie()

    def volver_a_revisar(self) -> None:
        self.aviso_flujo.hide()
        self._mostrar_zona("listo")
        self.revisar()

    def revisar(self) -> None:
        if self.estado == "corriendo":
            return
        if self._revisando:
            # Cambió algo (el archivo, por ejemplo) mientras se revisaba: al terminar, otra vez
            self._otra_vez = True
            return
        self._revisando = True
        self.resumen_revision.poner("Revisando…", "neutro")
        if self.revision is None:
            self.mas_filas.setText("Leyendo el archivo y consultando SAP en MM03 (solo lectura)…")
            self.mas_filas.show()
            self.tabla.hide()
        self.nombre_archivo.setText(self.archivo.name if self.archivo else "Sin archivo elegido")
        self.boton_ver_archivo.setEnabled(bool(self.archivo and self.archivo.exists()))
        self._pintar_pie()
        archivo = self.archivo
        en_segundo_plano(
            self,
            lambda: self.ventana.entorno.revisar(self.tarea, archivo or Path("")),
            lambda resultado: self._al_revisar(resultado, archivo),
        )

    def _al_revisar(self, resultado, archivo: Path | None) -> None:
        self._revisando = False
        if self._otra_vez:
            self._otra_vez = False
            self.revisar()
            return
        if isinstance(resultado, Exception):
            resultado = AntesDeEjecutar(chequeos=[Chequeo("Revisión", BLOQUEA, f"No se pudo revisar: {resultado}")])
        self.revision, self._archivo_revisado = resultado, archivo
        self._pintar_chequeos()
        self._pintar_previa()
        if self.estado == "revisando":
            self._mostrar_zona("listo")
        self._pintar_pie()
        if self._auto_pendiente:
            self._auto_pendiente = False
            self._seguir_solo()

    def _pintar_chequeos(self) -> None:
        r = self.revision
        piezas.vaciar(self.lista_chequeos)
        for c in r.chequeos:
            self.lista_chequeos.addWidget(FichaChequeo(c))
        bloqueos = sum(1 for c in r.chequeos if c.estado == BLOQUEA)
        avisos = sum(1 for c in r.chequeos if c.estado == AVISO)
        if bloqueos:
            self.resumen_revision.poner(f"{bloqueos} por corregir", "error")
        elif avisos:
            self.resumen_revision.poner(f"{avisos} aviso{'s' if avisos > 1 else ''}", "aviso")
        else:
            self.resumen_revision.poner("Todo listo", "ok")

    def _pintar_previa(self) -> None:
        r = self.revision
        self.detalle_previa.setText(f"{self.archivo.name} · {r.total} filas" if self.archivo and r.total else "")
        columnas = ["Fila", *self.tarea.columnas, "Revisión"]
        visibles = r.filas[:FILAS_VISTA_PREVIA]
        self.tabla.clear()
        self.tabla.setColumnCount(len(columnas))
        self.tabla.setRowCount(len(visibles))
        self.tabla.setHorizontalHeaderLabels(columnas)
        cabecera = self.tabla.horizontalHeader()
        for i in range(len(columnas) - 1):
            cabecera.setSectionResizeMode(i, QHeaderView.ResizeMode.ResizeToContents)
        cabecera.setSectionResizeMode(len(columnas) - 1, QHeaderView.ResizeMode.Stretch)
        cabecera.setDefaultAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        for fila_tabla, (numero, datos) in enumerate(visibles):
            avisos = r.avisos_de_fila(numero)
            rechaza = any(a.rechaza for a in avisos)
            fondo = QColor(tema.ERROR_SUAVE if rechaza else tema.AVISO_SUAVE) if avisos else None
            valores = [str(numero), *[datos.get(c, "") for c in self.tarea.columnas], " · ".join(a.texto for a in avisos)]
            for col, valor in enumerate(valores):
                vacio = 0 < col < len(columnas) - 1 and not valor
                item = QTableWidgetItem("(vacío)" if vacio else valor)
                if vacio or col == 0:
                    item.setForeground(QColor(tema.TINTA_3))
                if col == len(columnas) - 1 and avisos:
                    item.setForeground(QColor(tema.ERROR if rechaza else tema.AVISO))
                if fondo:
                    item.setBackground(fondo)
                self.tabla.setItem(fila_tabla, col, item)
        for i in range(len(visibles)):
            self.tabla.setRowHeight(i, 30)
        self.tabla.setFixedHeight(len(visibles) * 30 + self.tabla.horizontalHeader().sizeHint().height() + 4)
        self.tabla.setVisible(bool(visibles))

        ocultas = [a for a in r.avisos if a.fila not in {n for n, _ in visibles}]
        textos = []
        if r.total > len(visibles):
            textos.append(f"Se muestran las primeras {len(visibles)} de {r.total} filas.")
        if ocultas:
            textos.append("Avisos más abajo: " + "; ".join(f"fila {a.fila} ({a.clave}): {a.texto}" for a in ocultas[:5]))
        if not visibles:
            textos.append("Sin filas que mostrar: revisa lo marcado arriba.")
        self.mas_filas.setText(" ".join(textos))
        self.mas_filas.setVisible(bool(textos))

    def _elegir_archivo(self) -> None:
        inicio = str(self.archivo.parent) if self.archivo else str(Path.home())
        ruta, _ = QFileDialog.getOpenFileName(self, "Archivo de entrada", inicio, "Excel o CSV (*.xlsx *.csv)")
        if ruta:
            self.poner_archivo(Path(ruta))

    def poner_archivo(self, ruta: Path) -> None:
        self.archivo = ruta
        self.volver_a_revisar()

    # ======================================================================
    # Ejecutar
    # ======================================================================
    @property
    def lista_para_ejecutar(self) -> bool:
        """Revisada, sin bloqueos y la revisión es del archivo que está elegido ahora."""
        return (
            bool(self.archivo and self.revision and self.revision.puede_ejecutar)
            and not self._revisando
            and self._archivo_revisado == self.archivo
        )

    def ejecutar(self, confirmar: bool = True) -> bool:
        if self.estado == "corriendo" or not self.lista_para_ejecutar:
            return False
        corredor = self.ventana.ejecutar(self, self.tarea, self.archivo, confirmar=confirmar)
        if corredor is None:
            return False
        self.corrida = None
        self.aviso_flujo.hide()
        self._total = self.revision.total
        self._actual = 0
        self._inicio = time.monotonic()
        self.barra.preparar(self._total)
        self.log.clear()
        self._mostrar_zona("corriendo")
        self._pintar_avance()
        self._reloj.start()
        return True

    def detener(self) -> None:
        if self.ventana.corredor:
            self.ventana.corredor.detener()
            self.ahora.setText("Deteniendo después de la fila actual…")
            self._pintar_pie()

    # --- señales de la corrida (llegan en el hilo de la interfaz) ------------------------
    def al_linea(self, texto: str) -> None:
        self.log.appendPlainText(texto)

    def al_avance(self, actual: int, total: int, clave: str) -> None:
        self._total = total
        if clave:
            self._actual = actual
            self.barra.en_curso(actual)
            if not (self.ventana.corredor and self.ventana.corredor.deteniendo):
                self.ahora.setText(f"Ahora: {clave.split('|')[0]}")
        self._pintar_avance()

    def al_fila(self, numero: int, clave: str, resultado: str) -> None:
        self.barra.marcar(numero, resultado)
        ok, om, err = self.barra.cuenta("ok"), self.barra.cuenta("omitida"), self.barra.cuenta("error")
        self.cuenta.setText(f"{ok} OK · {om} omitida{'s' if om != 1 else ''} · {err} con error")

    def _pintar_avance(self) -> None:
        transcurrido = time.monotonic() - self._inicio
        self.cronometro.setText(reloj(transcurrido))
        if not self._actual:
            self.contador.setText(f"0 de {self._total}")
            self.quedan.setText("empezando…")
            return
        hechas = self._actual - 1
        quedan = self._total - self._actual
        restante = estimar_restante(transcurrido, hechas, self._total)
        estimado = "calculando…" if restante is None else f"~{texto_duracion(restante)}"
        self.contador.setText(f"{self._actual} de {self._total}")
        self.quedan.setText(f"quedan {quedan} · {estimado}")

    def al_terminar(self, corrida: Corrida) -> None:
        self._reloj.stop()
        self.corrida = corrida
        self.cifras["ok"].poner(corrida.ok)
        self.cifras["omitidas"].poner(corrida.omitidas)
        self.cifras["errores"].poner(corrida.errores)
        if corrida.detenido:
            titulo, color = "Se detuvo", tema.ERROR
            detalle = corrida.motivo
        elif corrida.errores:
            n = corrida.errores
            titulo, color = f"Terminó con {n} fila{'s' if n > 1 else ''} con error", tema.ERROR
            detalle = "Las filas con error quedaron en el CSV de rechazadas, listas para corregir y volver a ejecutar."
        else:
            titulo, color = "Terminó sin errores", tema.OK
            detalle = "Todo lo del archivo quedó hecho."
        self.titulo_resultado.setText(titulo)
        self.titulo_resultado.setStyleSheet(f"color: {color};")
        self.detalle_resultado.setText(f"{detalle}  ·  {corrida.total} filas en {texto_duracion(corrida.duracion)}")
        self.boton_rechazadas.setVisible(bool(corrida.rechazadas))
        self.boton_log.setEnabled(bool(corrida.log))
        self._mostrar_zona("terminado")
        if self.pasos:
            self.pasos.actualizar(self.ventana.entorno.historial.leer())
        self._decidir_flujo(corrida)

    # ======================================================================
    # Flujo automático
    # ======================================================================
    def _sincronizar_automatico(self) -> None:
        prendido = self.proceso.id in self.ventana.flujos_automaticos
        if self.automatico.isChecked() != prendido:
            self.automatico.blockSignals(True)
            self.automatico.setChecked(prendido)
            self.automatico.blockSignals(False)
            self.automatico.update()

    def _al_cambiar_automatico(self, prendido: bool) -> None:
        if prendido and not self.ventana.entorno.demo:
            r = QMessageBox.question(
                self,
                "Flujo automático",
                "Con el flujo automático, los pasos siguientes se ejecutan en SAP sin volver a preguntar "
                "(solo si terminan bien y su revisión no trae avisos). ¿Prenderlo?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if r != QMessageBox.StandardButton.Yes:
                self.automatico.blockSignals(True)
                self.automatico.setChecked(False)
                self.automatico.blockSignals(False)
                return
        if prendido:
            self.ventana.flujos_automaticos.add(self.proceso.id)
        else:
            self.ventana.flujos_automaticos.discard(self.proceso.id)

    def _decidir_flujo(self, corrida: Corrida) -> None:
        siguiente = self.proceso.siguiente(self.tarea.id)
        if not (siguiente and self.automatico.isChecked()):
            return
        if corrida.con_problemas:
            n = corrida.errores
            motivo = "se detuvo" if corrida.detenido else f"terminó con {n} fila{'s' if n > 1 else ''} con error"
            self._avisar_flujo(f"Flujo automático detenido: este paso {motivo}. Revisa el resultado antes de seguir.")
            return
        vista = self.ventana.abrir_tarea(siguiente)
        vista.recibir_flujo(self.tarea)

    def recibir_flujo(self, anterior: Tarea) -> None:
        """La abre el paso anterior al terminar bien: revisar y, si todo está en orden, ejecutar."""
        self._avisar_flujo(f"Viene de {anterior.nombre}, que terminó bien. Revisando antes de seguir…", tono="marca")
        if self.estado == "terminado":
            self._mostrar_zona("listo")
        # Si ya hay una revisión en marcha (la abrió al_mostrar), se sigue cuando termine
        self._auto_pendiente = True
        if not self._revisando:
            self.revisar()

    def _seguir_solo(self) -> None:
        r = self.revision
        motivo = r.motivo_para_frenar() if r else "no se pudo revisar"
        if motivo:
            self._avisar_flujo(f"Flujo automático detenido antes de ejecutar: {motivo}. Revisa y ejecuta a mano.")
            return
        self._avisar_flujo("Revisión sin avisos: el paso arranca solo en un momento.", tono="marca")
        QTimer.singleShot(PAUSA_FLUJO_MS, self._arrancar_solo)

    def _arrancar_solo(self) -> None:
        if self.automatico.isChecked() and self.estado == "listo" and not self.ventana.ocupada:
            self.ejecutar(confirmar=False)

    def _avisar_flujo(self, texto: str, tono: str = "aviso") -> None:
        color, fondo = tema.TONOS[tono]
        self.aviso_flujo.setText(texto)
        self.aviso_flujo.setStyleSheet(
            f"QLabel {{ background: {fondo}; color: {color}; border-radius: {tema.RADIO_CONTROL}px; padding: 10px 14px; }}"
        )
        self.aviso_flujo.show()
