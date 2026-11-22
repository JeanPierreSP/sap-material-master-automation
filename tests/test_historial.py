"""Historial de corridas: guardar, leer y filtrar."""

from datetime import date, datetime

from sap_mm_auto.historial import (
    FALLO,
    HECHO,
    PENDIENTE,
    Corrida,
    Historial,
    estado_del_dia,
    filtrar,
    texto_duracion,
)

HOY = date(2026, 3, 10)


def corrida(tarea="ampliacion_almacen", dia=HOY, hora="10:00", errores=0, detenido=False, **extra):
    momento = datetime.fromisoformat(f"{dia.isoformat()}T{hora}:00")
    return Corrida(tarea, extra.pop("nombre", "Ampliación de almacén"), "Materiales", momento.isoformat(), 5.0,
                   total=8, ok=8 - errores, errores=errores, detenido=detenido, **extra)


def test_guardar_y_leer_de_la_mas_nueva_a_la_mas_vieja(tmp_path):
    h = Historial(tmp_path / "historial.jsonl")
    h.agregar(corrida(hora="09:00"))
    h.agregar(corrida(hora="11:00", errores=2))
    leidas = h.leer()
    assert [c.inicio[-8:-3] for c in leidas] == ["11:00", "09:00"]
    assert leidas[0].estado == FALLO and leidas[1].estado == HECHO


def test_una_linea_rota_no_impide_leer_las_demas(tmp_path):
    h = Historial(tmp_path / "historial.jsonl")
    h.agregar(corrida())
    with h.ruta.open("a", encoding="utf-8") as f:
        f.write('{"tarea": "cortada a la mit\n')
    h.agregar(corrida(hora="12:00"))
    assert len(h.leer()) == 2


def test_estado_del_dia_usa_la_ultima_corrida_de_hoy():
    corridas = [corrida(hora="09:00", errores=1), corrida(hora="10:30"), corrida(dia=date(2026, 3, 9))]
    estado, ultima = estado_del_dia(corridas, "ampliacion_almacen", hoy=HOY)
    assert estado == HECHO and ultima.inicio.endswith("10:30:00")
    assert estado_del_dia(corridas, "bloqueo_material", hoy=HOY) == (PENDIENTE, None)


def test_filtros_de_periodo_y_errores():
    corridas = [
        corrida(),
        corrida(dia=date(2026, 3, 8), errores=1),
        corrida(dia=date(2026, 3, 4)),  # justo 6 días atrás: entra en 7 días
        corrida(dia=date(2026, 2, 20), detenido=True),
    ]
    assert len(filtrar(corridas, "hoy", hoy=HOY)) == 1
    assert len(filtrar(corridas, "7dias", hoy=HOY)) == 3
    assert len(filtrar(corridas, "todo", hoy=HOY)) == 4
    assert len(filtrar(corridas, "todo", solo_errores=True, hoy=HOY)) == 2


def test_buscar_por_texto_y_dentro_del_log(tmp_path):
    log = tmp_path / "corrida.log"
    log.write_text("[1/1] DEMO-0042: OK · Se ha ampliado el material DEMO-0042\n", encoding="utf-8")
    corridas = [corrida(log=str(log)), corrida(tarea="bloqueo_material", nombre="Bloqueo de material")]
    assert len(filtrar(corridas, texto="bloqueo")) == 1
    from sap_mm_auto.historial import leer_log

    assert len(filtrar(corridas, texto="demo-0042", leer_log=leer_log)) == 1
    assert filtrar(corridas, texto="demo-0042") == []


def test_texto_duracion():
    assert texto_duracion(0.4) == "<1 s"
    assert texto_duracion(42) == "42 s"
    assert texto_duracion(125) == "2 min 05 s"
    assert texto_duracion(3780) == "1 h 03 min"
