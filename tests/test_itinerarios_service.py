"""Pruebas automatizadas del módulo de itinerarios, resolución de diagramas de 6 días y buscador por tren."""

from pathlib import Path
from src.data.database import DatabaseManager
from src.data.migrations import init_database_schema
from src.services.records_service import RecordsService
from src.domain.validators import (
    resolver_subindice_diagrama_6d,
    resolver_jornada_itinerario,
)


def test_resolver_subindice_diagrama_6d():
    # Caso 1: Franco Domingo habitual
    # Lunes -> 1, Martes -> 2, Sábado -> 6, Domingo -> Franco
    sub, es_f = resolver_subindice_diagrama_6d("Lunes", "Domingo")
    assert not es_f and sub == 1

    sub, es_f = resolver_subindice_diagrama_6d("Martes", "Domingo")
    assert not es_f and sub == 2

    sub, es_f = resolver_subindice_diagrama_6d("Sábado", "Domingo")
    assert not es_f and sub == 6

    sub, es_f = resolver_subindice_diagrama_6d("Domingo", "Domingo")
    assert es_f and sub is None

    # Caso 2: Franco Martes
    # Miércoles -> 1, Jueves -> 2, Martes -> Franco
    sub, es_f = resolver_subindice_diagrama_6d("Miércoles", "Martes")
    assert not es_f and sub == 1

    sub, es_f = resolver_subindice_diagrama_6d("Martes", "Martes")
    assert es_f and sub is None


def test_importar_itinerario_y_buscador_tren(tmp_path: Path):
    db_file = tmp_path / "test_itinerarios.sqlite"
    db = DatabaseManager(db_file, journal_mode="WAL")
    init_database_schema(db)
    svc = RecordsService(db)

    # Importar archivo real Indice_45.xlsx
    excel_path = Path("Indice_45.xlsx")
    assert excel_path.exists()

    turnos, vueltas = svc.importar_itinerario_excel(str(excel_path))
    assert turnos == 218
    assert vueltas >= 900

    # Consulta de tren 2303
    res_2303 = svc.buscar_turnos_por_tren("2303")
    assert len(res_2303) == 1
    assert res_2303[0]["codigo_turno"] == "401.1"
    assert res_2303[0]["origen"] == "PC"
    assert res_2303[0]["destino"] == "BQ"

    # Consulta de tren 4356 (intervienen múltiples turnos)
    res_4356 = svc.buscar_turnos_por_tren("4356")
    assert len(res_4356) == 3
    codigos_4356 = [r["codigo_turno"] for r in res_4356]
    assert "401.1" in codigos_4356
    assert "401.2" in codigos_4356
    assert "401.6" in codigos_4356

    # Test resolución dinámica de jornada
    turnos_map = svc.listar_turnos_itinerario_map()

    # Fecha: 2026-11-02 es un Lunes
    # Agente con franco Domingo -> 401 debe resolver a 401.1 (toma: 18:55, deje: 00:20)
    res_lunes = resolver_jornada_itinerario("401", "2026-11-02", "Domingo", turnos_map)
    assert not res_lunes["es_franco"]
    assert res_lunes["codigo_turno"] == "401.1"
    assert res_lunes["toma"] == "18:55"
    assert res_lunes["deje"] == "00:20"

    # Fecha: 2026-11-08 es un Domingo -> 401 debe resolver a FRANCO
    res_domingo = resolver_jornada_itinerario("401", "2026-11-08", "Domingo", turnos_map)
    assert res_domingo["es_franco"]
    assert "FRANCO" in res_domingo["codigo_turno"]
