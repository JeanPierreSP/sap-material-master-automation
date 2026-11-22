"""Pantalla de la tarea: revisión, vista previa, corrida, resultado y flujo automático."""

import pytest

pytest.importorskip("PySide6.QtWidgets")

from conftest import esperar  # noqa: E402
from sap_mm_auto.entorno import Entorno  # noqa: E402
from sap_mm_auto.ui.ventana import Ventana  # noqa: E402
from sap_mm_auto.ui.vistas.tarea import FichaChequeo, reloj  # noqa: E402


def abrir_revisada(app, ventana, tarea_id):
    vista = ventana.abrir_tarea(tarea_id)
    assert esperar(app, lambda: vista.revision is not None and not vista._revisando), "la revisión no terminó"
    app.processEvents()  # Qt muestra en el ciclo siguiente lo que se agregó a un layout visible
    return vista


def esperar_fin(app, ventana, vista, segundos=20):
    assert esperar(app, lambda: vista.estado == "terminado" and not ventana.ocupada, segundos), "la corrida no terminó"


def test_revision_muestra_chequeos_y_avisos_en_la_vista_previa(app, ventana):
    vista = abrir_revisada(app, ventana, "ampliacion_almacen")
    assert vista.resumen_revision.text() == "2 avisos"
    fichas = [f for f in vista.findChildren(FichaChequeo) if not f.isHidden()]
    assert [f.chequeo.nombre for f in fichas] == [c.nombre for c in vista.revision.chequeos]
    assert [f.chequeo.estado for f in fichas].count("aviso") == 2
    assert fichas[-1].chequeo.nombre == "SAP abierto en la pantalla inicial"
    revision_col = vista.tabla.columnCount() - 1
    textos = {vista.tabla.item(f, 1).text(): vista.tabla.item(f, revision_col).text() for f in range(vista.tabla.rowCount())}
    assert textos["DEMO-0004"] == "Falta ALMACEN"
    assert textos["DEMO-0099"] == "No existe en SAP"
    assert textos["DEMO-0006"].startswith("Status 01")
    assert vista.tabla.item(3, 3).text() == "(vacío)"
    assert vista.principal.isEnabled() and vista.principal.text() == "Ejecutar"


def test_corrida_completa_desde_la_interfaz(app, ventana):
    vista = abrir_revisada(app, ventana, "ampliacion_almacen")
    vista.principal.click()
    assert vista.estado == "corriendo" and vista.principal.text() == "Detener"
    assert ventana.barra.corrida.isVisibleTo(ventana)
    esperar_fin(app, ventana, vista)

    assert [vista.cifras[k].numero.text() for k in ("ok", "omitidas", "errores")] == ["5", "0", "3"]
    assert (vista.barra.cuenta("ok"), vista.barra.cuenta("error")) == (5, 3)
    assert vista.titulo_resultado.text() == "Terminó con 3 filas con error"
    assert "Fin: 5 OK" in vista.log.toPlainText()
    assert vista.boton_rechazadas.isVisibleTo(vista)
    assert vista.principal.text() == "Volver a revisar"  # último paso del flujo
    assert not ventana.barra.corrida.isVisibleTo(ventana)
    assert ventana.entorno.historial.leer()[0].tarea == "ampliacion_almacen"


def test_avance_dice_cuantas_van_y_cuantas_quedan(app, ventana):
    vista = abrir_revisada(app, ventana, "ampliacion_almacen")
    vista.ejecutar()
    vista.al_avance(3, 8, "DEMO-0003|1000|0002")  # como lo manda el motor al empezar la fila 3
    assert vista.contador.text() == "3 de 8"
    assert vista.quedan.text().startswith("quedan 5 · ")
    assert vista.ahora.text() == "Ahora: DEMO-0003"
    esperar_fin(app, ventana, vista)


def test_detener_a_mitad_de_corrida(app, tmp_path):
    v = Ventana(Entorno(demo=True, carpeta=tmp_path, retardo=0.05))
    v.show()
    vista = abrir_revisada(app, v, "ampliacion_almacen")
    vista.principal.click()
    assert esperar(app, lambda: vista._actual >= 2)
    vista.principal.click()  # Detener
    assert vista.principal.text() == "Deteniendo…" and not vista.principal.isEnabled()
    esperar_fin(app, v, vista)
    corrida = v.entorno.historial.leer()[0]
    assert corrida.detenido and "usuario" in corrida.motivo
    assert vista.titulo_resultado.text() == "Se detuvo"
    v.close()


def test_filas_repetidas_bloquean_ejecutar(app, ventana, tmp_path):
    archivo = tmp_path / "repetidas.csv"
    archivo.write_text("MATERIAL,CENTRO,ALMACEN\nDEMO-0001,1000,0002\nDEMO-0001,1000,0002\n", encoding="utf-8")
    vista = abrir_revisada(app, ventana, "ampliacion_almacen")
    vista.poner_archivo(archivo)
    assert esperar(app, lambda: not vista._revisando)
    assert vista.resumen_revision.text() == "1 por corregir"
    assert not vista.principal.isEnabled() and "Corrige" in vista.principal.toolTip()


def test_sin_archivo_no_se_puede_ejecutar(app, tmp_path):
    v = Ventana(Entorno(demo=True, carpeta=tmp_path))
    vista = v.abrir_tarea("bloqueo_material")
    vista.archivo = None
    vista.volver_a_revisar()
    assert esperar(app, lambda: not vista._revisando)
    assert not vista.principal.isEnabled() and vista.nombre_archivo.text() == "Sin archivo elegido"


def test_otra_corrida_en_curso_apaga_ejecutar(app, tmp_path):
    v = Ventana(Entorno(demo=True, carpeta=tmp_path, retardo=0.05))
    v.show()
    vista = abrir_revisada(app, v, "ampliacion_almacen")
    vista.ejecutar()
    otra = abrir_revisada(app, v, "bloqueo_material")
    assert not otra.principal.isEnabled() and "otra corrida" in otra.principal.toolTip()
    vista.detener()
    esperar_fin(app, v, vista)
    v.close()


def test_flujo_automatico_frena_ante_avisos_del_paso_siguiente(app, ventana):
    paso1 = abrir_revisada(app, ventana, "cambio_descripcion")
    assert paso1.texto_automatico.text() == "Al terminar bien, sigue solo con Ampliación de almacén."
    paso1.automatico.setChecked(True)
    paso1.principal.click()
    paso2 = ventana.vista("tarea:ampliacion_almacen")
    assert esperar(app, lambda: ventana.actual == "tarea:ampliacion_almacen" and not ventana.vista(ventana.actual)._revisando)
    paso2 = ventana.vista("tarea:ampliacion_almacen")
    assert esperar(app, lambda: paso2.aviso_flujo.isVisibleTo(paso2) and "detenido" in paso2.aviso_flujo.text())
    assert "Filas completas" in paso2.aviso_flujo.text()
    assert paso2.estado == "listo" and len(ventana.entorno.historial.leer()) == 1  # el paso 2 no corrió
    assert paso2.automatico.isChecked()  # el interruptor es del proceso, no de la pantalla


def test_flujo_automatico_sigue_si_el_paso_siguiente_esta_limpio(app, ventana, tmp_path):
    limpio = tmp_path / "ampliacion_limpia.csv"
    limpio.write_text("MATERIAL,CENTRO,ALMACEN\nDEMO-0008,1000,0002\nDEMO-0009,1000,0003\n", encoding="utf-8")
    paso2 = abrir_revisada(app, ventana, "ampliacion_almacen")
    paso2.poner_archivo(limpio)
    paso1 = abrir_revisada(app, ventana, "cambio_descripcion")
    paso1.automatico.setChecked(True)
    paso1.principal.click()
    assert esperar(app, lambda: len(ventana.entorno.historial.leer()) == 2 and not ventana.ocupada, 20)
    assert [c.tarea for c in ventana.entorno.historial.leer()] == ["ampliacion_almacen", "cambio_descripcion"]
    assert esperar(app, lambda: paso2.estado == "terminado")
    assert paso2.titulo_resultado.text() == "Terminó sin errores"


def test_flujo_automatico_frena_si_el_paso_termina_con_errores(app, ventana, tmp_path):
    con_error = tmp_path / "descripcion.csv"
    con_error.write_text("MATERIAL,DESCRIPCION\nDEMO-0001,UNA DESCRIPCION DE MAS DE CUARENTA CARACTERES XX\n",
                         encoding="utf-8")
    paso1 = abrir_revisada(app, ventana, "cambio_descripcion")
    paso1.poner_archivo(con_error)
    assert esperar(app, lambda: not paso1._revisando)
    paso1.automatico.setChecked(True)
    paso1.principal.click()
    esperar_fin(app, ventana, paso1)
    assert ventana.actual == "tarea:cambio_descripcion"
    assert "terminó con 1 fila con error" in paso1.aviso_flujo.text()
    assert paso1.principal.text() == "Ir a Ampliación de almacén"


def test_reloj():
    assert reloj(5) == "00:05"
    assert reloj(125) == "02:05"
    assert reloj(3725) == "1:02:05"
