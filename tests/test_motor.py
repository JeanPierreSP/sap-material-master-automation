"""El bucle completo: errores por fila, corridas repetidas y detenciones."""

from sap_mm_auto import motor
from sap_mm_auto.controles import MemoriaDiaria
from sap_mm_auto.entrada import duplicados, leer_filas
from sap_mm_auto.registro import Registro
from sap_mm_auto.sesion import Pantalla
from sap_mm_auto.simulador import SAPSimulado
from sap_mm_auto.tareas import TAREAS

tarea = TAREAS["ampliacion_almacen"]


def filas(*items):
    return [{"MATERIAL": m, "CENTRO": c, "ALMACEN": a} for m, c, a in items]


def test_una_fila_con_error_no_detiene_la_corrida(pantalla, memoria, registro):
    entrada = filas(("DEMO-0001", "1000", "0002"), ("NO-EXISTE", "1000", "0002"), ("DEMO-0002", "1000", "0002"))
    r = motor.ejecutar(tarea, entrada, pantalla, memoria, registro)
    assert (r.ok, r.errores, r.detenido) == (2, 1, False)
    rechazadas = registro.ruta_rechazadas.read_text(encoding="utf-8-sig")
    assert "NO-EXISTE" in rechazadas and "MOTIVO" in rechazadas


def test_relanzar_la_corrida_no_repite_lo_hecho(tmp_path):
    sap = SAPSimulado()
    entrada = filas(("DEMO-0001", "1000", "0002"), ("DEMO-0002", "1000", "0002"))
    motor.ejecutar(tarea, entrada, Pantalla(sap), MemoriaDiaria(tmp_path), Registro(tmp_path, "t"))
    r = motor.ejecutar(tarea, entrada, Pantalla(sap), MemoriaDiaria(tmp_path), Registro(tmp_path, "t"))
    assert (r.ok, r.omitidas) == (0, 2)
    assert len(sap.grabaciones) == 2  # SAP solo grabó la primera vez


def test_frena_si_el_mismo_error_se_repite(pantalla, memoria, registro):
    entrada = filas(*[(f"NO-{i}", "1000", "0002") for i in range(5)], ("DEMO-0001", "1000", "0002"))
    r = motor.ejecutar(tarea, entrada, pantalla, memoria, registro)
    assert r.detenido and r.errores == 3 and r.ok == 0


def test_frena_ante_una_pantalla_inesperada(tmp_path):
    sap = SAPSimulado(pantalla_inesperada={"DEMO-0002"})
    entrada = filas(("DEMO-0001", "1000", "0002"), ("DEMO-0002", "1000", "0002"), ("DEMO-0003", "1000", "0003"))
    r = motor.ejecutar(tarea, entrada, Pantalla(sap), MemoriaDiaria(tmp_path), Registro(tmp_path, "t"))
    assert r.detenido and "inesperada" in r.motivo.lower()
    assert sap.grabaciones == [("MM01", "DEMO-0001")]  # la tercera fila no se tocó


def test_el_usuario_puede_detener(pantalla, memoria, registro):
    llamadas = {"n": 0}

    def detener():
        llamadas["n"] += 1
        return llamadas["n"] > 1

    entrada = filas(("DEMO-0001", "1000", "0002"), ("DEMO-0002", "1000", "0002"))
    r = motor.ejecutar(tarea, entrada, pantalla, memoria, registro, detener_solicitado=detener)
    assert r.ok == 1 and r.detenido and "usuario" in r.motivo


def test_archivos_de_ejemplo_sin_duplicados():
    from pathlib import Path

    carpeta = Path(__file__).resolve().parents[1] / "datos_ejemplo"
    for t in TAREAS.values():
        datos = leer_filas(carpeta / f"{t.id}.csv")
        assert datos and not duplicados(datos, t.clave)


def test_avisa_el_resultado_de_cada_fila(pantalla, memoria, registro):
    entrada = filas(("DEMO-0001", "1000", "0002"), ("NO-EXISTE", "1000", "0002"), ("DEMO-0002", "1000", ""))
    vistos = []
    motor.ejecutar(tarea, entrada, pantalla, memoria, registro, al_terminar_fila=lambda *a: vistos.append(a))
    assert vistos == [(1, "DEMO-0001|1000|0002", "ok"), (2, "NO-EXISTE|1000|0002", "error"), (3, "DEMO-0002|1000|", "error")]
