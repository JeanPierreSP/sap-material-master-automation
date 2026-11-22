"""Entorno: revisar, ejecutar y dejar cada corrida en el historial."""

from datetime import datetime

from sap_mm_auto.entorno import Entorno, estimar_restante, sembrar_ejemplos
from sap_mm_auto.historial import FALLO, HECHO, Historial, estado_del_dia, leer_log
from sap_mm_auto.tareas import TAREAS

ampliacion = TAREAS["ampliacion_almacen"]
descripcion = TAREAS["cambio_descripcion"]


def test_ejecutar_deja_la_corrida_en_el_historial(tmp_path):
    e = Entorno(demo=True, carpeta=tmp_path)
    resultados = []
    archivo = e.archivo_de_ejemplo(ampliacion)
    corrida = e.ejecutar(ampliacion, archivo, al_terminar_fila=lambda i, clave, r: resultados.append(r))

    assert (corrida.total, corrida.ok, corrida.errores, corrida.detenido) == (8, 5, 3, False)
    assert resultados == ["ok", "ok", "ok", "error", "error", "ok", "error", "ok"]
    assert corrida.maestro == "Materiales" and corrida.demo
    assert "Fin: 5 OK" in leer_log(corrida)
    assert corrida.rechazadas.endswith("_rechazadas.csv")
    assert e.historial.leer()[0] == corrida
    assert estado_del_dia(e.historial.leer(), "ampliacion_almacen")[0] == FALLO


def test_el_sap_simulado_dura_toda_la_sesion(tmp_path):
    e = Entorno(demo=True, carpeta=tmp_path)
    archivo = e.archivo_de_ejemplo(descripcion)
    primera = e.ejecutar(descripcion, archivo)
    segunda = e.ejecutar(descripcion, archivo)
    assert primera.ok == 5 and primera.estado == HECHO
    assert (segunda.ok, segunda.omitidas) == (0, 5)  # memoria diaria: no repite


def test_revisar_usa_la_consulta_de_solo_lectura(tmp_path):
    e = Entorno(demo=True, carpeta=tmp_path)
    r = e.revisar(ampliacion, e.archivo_de_ejemplo(ampliacion))
    assert {a.clave for a in r.avisos} == {"DEMO-0004", "DEMO-0099", "DEMO-0006"}
    assert e._sap.grabaciones == []


def test_archivo_ilegible_queda_como_corrida_detenida(tmp_path):
    e = Entorno(demo=True, carpeta=tmp_path)
    lineas = []
    corrida = e.ejecutar(ampliacion, tmp_path / "entrada.txt", al_linea=lineas.append)
    assert corrida.detenido and "No se pudo ejecutar" in corrida.motivo
    assert lineas and e.historial.leer()[0].detenido


def test_estado_sap_en_demo(tmp_path):
    estado = Entorno(demo=True, carpeta=tmp_path).estado_sap()
    assert estado.conectado and estado.texto == "SAP simulado"


def test_estimar_restante():
    assert estimar_restante(10, 0, 8) is None
    assert estimar_restante(6, 3, 8) == 10
    assert estimar_restante(6, 8, 8) == 0


def test_historial_de_ejemplo_trae_logs(tmp_path):
    h = Historial(tmp_path / "historial.jsonl")
    sembrar_ejemplos(h, tmp_path, datetime(2026, 3, 10, 8, 0))
    corridas = h.leer()
    assert len(corridas) == 5 and all(c.demo for c in corridas)
    assert any(c.detenido for c in corridas)
    assert all("Inicio:" in leer_log(c) for c in corridas)


def test_con_ejemplos_solo_en_demo(tmp_path):
    assert len(Entorno(demo=True, carpeta=tmp_path / "a", con_ejemplos=True).historial.leer()) == 5
