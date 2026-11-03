"""Pruebas unitarias de las validaciones de dominio puras (sin Tkinter ni dependencias GUI)."""

from datetime import datetime
import pytest

from src.domain.enums import EstadoCobertura
from src.domain.validators import (
    calcular_descansos_y_horas_ciclo,
    calcular_estado_cobertura_puro,
    dentro_de_ventana_horaria,
    es_turno_ciclo,
    parsear_fecha,
    parsear_registrado_en,
    validar_datos_cambio,
    validar_datos_novedad,
    validar_horario_hm,
    validar_tramos_reemplazo,
)


def test_parsear_fecha():
    assert parsear_fecha("15/06/2026").strftime("%d/%m/%Y") == "15/06/2026"
    assert parsear_fecha("2026-06-15").strftime("%d/%m/%Y") == "15/06/2026"
    assert parsear_fecha("", obligatorio=False) is None
    with pytest.raises(ValueError):
        parsear_fecha("", obligatorio=True)
    with pytest.raises(ValueError):
        parsear_fecha("fecha-invalida")


def test_es_turno_ciclo():
    assert es_turno_ciclo("CICLO 1") is True
    assert es_turno_ciclo("ciclo especial") is True
    assert es_turno_ciclo("T1") is False
    assert es_turno_ciclo(None) is False


def test_validar_horario_hm():
    assert validar_horario_hm("06:30") == (6, 30)
    assert validar_horario_hm("23:59") == (23, 59)
    assert validar_horario_hm("") is None
    with pytest.raises(ValueError):
        validar_horario_hm("25:00")
    with pytest.raises(ValueError):
        validar_horario_hm("06:65")


def test_dentro_de_ventana_horaria():
    ahora = datetime(2026, 6, 15, 12, 0, 0)
    reciente = "2026-06-15 08:00:00"
    antiguo = "2026-06-10 08:00:00"
    assert dentro_de_ventana_horaria(reciente, 24, ahora=ahora) is True
    assert dentro_de_ventana_horaria(antiguo, 24, ahora=ahora) is False


def test_validar_datos_novedad():
    data_valida = {
        "legajo": 1234,
        "apellidos_nombres": "Pérez, Juan",
        "novedad": "Licencia",
        "fecha_inicio": "10/06/2026",
        "fecha_fin": "12/06/2026",
        "referencia_estacion": "Retiro",
        "supervisor": "González",
    }
    valido, err = validar_datos_novedad(data_valida, tipo_novedades=["Licencia"], personal_estacion=["Retiro"])
    assert valido is True
    assert err is None

    # Fecha fin menor que inicio
    data_invalida = dict(data_valida, fecha_fin="08/06/2026")
    valido, err = validar_datos_novedad(data_invalida)
    assert valido is False
    assert "no puede ser anterior" in err


def test_validar_tramos_reemplazo():
    tramos = [
        {"legajo_reemplazo": 999, "turnos_reemplazo": "CICLO 1", "fecha_inicio": "10/06/2026", "fecha_fin": "11/06/2026"},
        {"legajo_reemplazo": 998, "turnos_reemplazo": "CICLO 2", "fecha_inicio": "12/06/2026", "fecha_fin": "15/06/2026"},
    ]
    valido, err = validar_tramos_reemplazo("10/06/2026", "15/06/2026", "T1", tramos)
    assert valido is True
    assert err is None

    # Solapamiento de tramos
    tramos_solapados = [
        {"legajo_reemplazo": 999, "turnos_reemplazo": "CICLO 1", "fecha_inicio": "10/06/2026", "fecha_fin": "12/06/2026"},
        {"legajo_reemplazo": 998, "turnos_reemplazo": "CICLO 2", "fecha_inicio": "11/06/2026", "fecha_fin": "15/06/2026"},
    ]
    valido, err = validar_tramos_reemplazo("10/06/2026", "15/06/2026", "T1", tramos_solapados)
    assert valido is False
    assert "solapamiento" in err.lower()


def test_calcular_estado_cobertura_puro():
    # Cobertura total
    tramos_totales = [
        {"legajo_reemplazo": 999, "apellidos_nombres_reemplazo": "A", "fecha_inicio": "10/06/2026", "fecha_fin": "12/06/2026"},
        {"legajo_reemplazo": 998, "apellidos_nombres_reemplazo": "B", "fecha_inicio": "13/06/2026", "fecha_fin": "15/06/2026"},
    ]
    res = calcular_estado_cobertura_puro("10/06/2026", "15/06/2026", "T1", tramos_totales)
    assert res["estado"] == EstadoCobertura.SI.value

    # Cobertura parcial (hueco el día 12)
    tramos_parciales = [
        {"legajo_reemplazo": 999, "apellidos_nombres_reemplazo": "A", "fecha_inicio": "10/06/2026", "fecha_fin": "11/06/2026"},
        {"legajo_reemplazo": 998, "apellidos_nombres_reemplazo": "B", "fecha_inicio": "13/06/2026", "fecha_fin": "15/06/2026"},
    ]
    res_parcial = calcular_estado_cobertura_puro("10/06/2026", "15/06/2026", "T1", tramos_parciales)
    assert res_parcial["estado"] == EstadoCobertura.PARCIAL.value

    # Titular CICLO
    res_ciclo = calcular_estado_cobertura_puro("10/06/2026", "15/06/2026", "CICLO 1", [])
    assert res_ciclo["estado"] == EstadoCobertura.NA_CICLO.value


def test_calcular_descansos_y_horas_ciclo():
    empleado = {
        "legajo": 5001,
        "jornada_horas": 6,
        "franco": "Domingo",
        "franco_horario": "14:00",
    }
    jornadas = {
        "2026-06-01": {"toma": "06:00", "deja": "12:00", "franco_anticipado": 0},
        "2026-06-02": {"toma": "06:00", "deja": "12:00", "franco_anticipado": 0},
    }
    serie = calcular_descansos_y_horas_ciclo(5001, empleado, [], jornadas, 2026, 6)
    assert len(serie) == 30
    # El día 2 debe tener 6 hs trabajadas y hora disponible calculada a partir de las 12:00 + 18hs = 06:00
    dia2 = serie[1]
    assert dia2["dia"] == 2
    assert dia2["horas_trabajadas"] == "6.0 hs"
    assert dia2["hora_disponible"] == "06:00"
