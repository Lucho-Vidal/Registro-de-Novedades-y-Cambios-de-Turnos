"""Pruebas unitarias e integración para los 8 nuevos requerimientos de RENO v2."""

import os
import tempfile
from pathlib import Path
import pytest
from PySide6.QtCore import QTime
from PySide6.QtWidgets import QApplication

from src.core.config import (
    obtener_directorio_aplicacion,
    guardar_ruta_base_datos,
    obtener_ruta_base_datos,
    PATH_BASE_FILE,
    CONFIG_JSON_PATH,
)
from src.data.database import DatabaseManager
from src.data.migrations import init_database_schema
from src.services.records_service import RecordsService
from src.services.excel_service import ExcelService
from src.ui.views.dialogs.calendario_ciclo_dialog import AsignarJornadaDialog
from src.ui.views.novedades_cambios_view import NovedadesCambiosView


@pytest.fixture(scope="session")
def qapp():
    app = QApplication.instance()
    if not app:
        app = QApplication([])
    return app


@pytest.fixture
def fresh_db(tmp_path):
    db_file = tmp_path / "test_reno.sqlite"
    mgr = DatabaseManager(db_file, journal_mode="DELETE")
    init_database_schema(mgr)
    return mgr


def test_persistencia_ruta_base_datos(tmp_path, monkeypatch):
    """1. Persistencia y configuración de ruta de base de datos."""
    app_dir = obtener_directorio_aplicacion()
    assert app_dir.exists()

    custom_db = tmp_path / "custom" / "produccion.sqlite"
    custom_db.parent.mkdir(parents=True, exist_ok=True)

    import json
    test_path_base = tmp_path / "config.json"
    test_config_json = tmp_path / "reno_config.json"

    import src.core.config as cfg_mod
    monkeypatch.setattr(cfg_mod, "PATH_BASE_FILE", test_path_base)
    monkeypatch.setattr(cfg_mod, "CONFIG_JSON_PATH", test_config_json)

    exito = guardar_ruta_base_datos(str(custom_db))
    assert exito is True
    assert test_path_base.exists()
    data = json.loads(test_path_base.read_text(encoding="utf-8"))
    assert data["database_path"] == str(custom_db)

    sqlite_path, _ = obtener_ruta_base_datos()
    assert sqlite_path == custom_db


def test_export_auditoria_excel(fresh_db, tmp_path):
    """2. Exportación a Excel de trazabilidad y auditoría."""
    excel_svc = ExcelService(fresh_db)
    records_svc = RecordsService(fresh_db)

    with fresh_db.write_transaction() as conn:
        conn.execute(
            """INSERT INTO auditoria (creado_en, usuario_windows, accion, entidad, entidad_id, datos_anteriores, datos_nuevos)
               VALUES (datetime('now', 'localtime'), 'test_user', 'crear', 'novedad', 101, NULL, '{"motivo": "Licencia Médica"}')"""
        )
        conn.execute(
            """INSERT INTO auditoria (creado_en, usuario_windows, accion, entidad, entidad_id, datos_anteriores, datos_nuevos)
               VALUES (datetime('now', 'localtime'), 'admin', 'modificar', 'usuario', 2, '{"rol": "Operador"}', '{"rol": "Supervisor"}')"""
        )

    logs = records_svc.listar_auditoria()
    assert len(logs) >= 2

    out_file = tmp_path / "auditoria_test.xlsx"
    res_path = excel_svc.export_auditoria(out_file, logs, filtro_aplicado="usuario")
    assert Path(res_path).exists()
    assert Path(res_path).stat().st_size > 1000


def test_turnos_sin_cubrir_orden_y_turno_dia_ciclo(fresh_db):
    """3 & 7. Turnos sin cubrir ordenados por hora toma, turno día para ciclo y filtro especialidad."""
    records_svc = RecordsService(fresh_db)

    with fresh_db.write_transaction() as conn:
        conn.execute("DELETE FROM novedades")
        conn.execute("DELETE FROM novedad_reemplazos")
        conn.execute("DELETE FROM tipos_novedad WHERE nombre='Enfermedad'")
        conn.execute("INSERT INTO tipos_novedad (nombre, exportar_descubierto) VALUES ('Enfermedad', 1)")

        # Itinerario con turno base 407 y sub-turnos
        conn.execute(
            """INSERT OR REPLACE INTO turnos_itinerario
               (itinerario_vigente, codigo_turno, turno_base, tipo_turno, deposito_base, toma_servicio, deje_servicio)
               VALUES ('H45', '407.1', '407', 'Día Hábil', 'PC', '06:15', '12:15')"""
        )
        conn.execute(
            """INSERT OR REPLACE INTO turnos_itinerario
               (itinerario_vigente, codigo_turno, turno_base, tipo_turno, deposito_base, toma_servicio, deje_servicio)
               VALUES ('H45', '407.3', '407', 'Día Hábil', 'PC', '06:15', '12:15')"""
        )
        conn.execute(
            """INSERT OR REPLACE INTO turnos_itinerario
               (itinerario_vigente, codigo_turno, turno_base, tipo_turno, deposito_base, toma_servicio, deje_servicio)
               VALUES ('H45', '101.1', '101', 'Día Hábil', 'LLV', '04:30', '10:30')"""
        )

        # Novedad titular con turno 407 (06:15)
        conn.execute(
            """INSERT INTO novedades (registrado_en, legajo, apellidos_nombres, especialidad, dotacion, turnos, franco, novedad, fecha_inicio, fecha_fin, referencia_estacion, supervisor, activo)
               VALUES (datetime('now', 'localtime'), 1001, 'Titular Perez', 'Conductor Eléctrico', 'PC', '407', 'Domingo', 'Enfermedad', '2026-11-04', '2026-11-10', 'PC', 'Supervisor 1', 1)"""
        )
        nov_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]

        # Reemplazo asignado a un agente de ciclo (Legajo 2002)
        conn.execute(
            """INSERT INTO novedad_reemplazos (creado_en, novedad_id, legajo_reemplazo, apellidos_nombres_reemplazo, turnos_reemplazo, franco_reemplazo, fecha_inicio, fecha_fin, activo)
               VALUES (datetime('now', 'localtime'), ?, 2002, 'Ciclo Relevo', 'CICLO-1', 'Lunes', '2026-11-04', '2026-11-10', 1)""",
            (nov_id,),
        )

        # Pero el agente de ciclo TAMBIÉN se enfermó y tiene novedad propia el 2026-11-04
        conn.execute(
            """INSERT INTO novedades (registrado_en, legajo, apellidos_nombres, especialidad, dotacion, turnos, franco, novedad, fecha_inicio, fecha_fin, referencia_estacion, supervisor, activo)
               VALUES (datetime('now', 'localtime'), 2002, 'Ciclo Relevo', 'Conductor Eléctrico', 'PC', 'CICLO-1', 'Lunes', 'Enfermedad', '2026-11-04', '2026-11-06', 'PC', 'Supervisor 1', 1)"""
        )

        # Otra novedad con toma más temprana (04:30) y distinta especialidad
        conn.execute(
            """INSERT INTO novedades (registrado_en, legajo, apellidos_nombres, especialidad, dotacion, turnos, franco, novedad, fecha_inicio, fecha_fin, referencia_estacion, supervisor, activo)
               VALUES (datetime('now', 'localtime'), 3003, 'Gomez Guarda', 'Guarda Eléctrico', 'LLV', '101', 'Domingo', 'Enfermedad', '2026-11-04', '2026-11-04', 'LLV', 'Supervisor 1', 1)"""
        )

    # Consulta turnos sin cubrir para el 04/11/2026
    descubiertos = records_svc.obtener_turnos_sin_cubrir("04/11/2026")
    assert len(descubiertos) >= 3

    # Ordenamiento por hora toma: primero debe estar Gomez (04:30) antes que Perez / Ciclo (06:15)
    horas = [d.get("hora_toma") for d in descubiertos if d.get("hora_toma") != "-"]
    assert horas == sorted(horas)

    # El agente de ciclo con novedad debe tener resuelto el turno día que le tocaba cubrir (407.3 para miércoles)
    item_ciclo = next(d for d in descubiertos if int(d["legajo"]) == 2002)
    assert item_ciclo["turno_dia"] == "407.3"
    assert item_ciclo["hora_toma"] == "06:15"
    assert item_ciclo["cubierto"] == "No"
    assert "Titular Perez" in item_ciclo["turno_reemplazado"]

    # Filtro por especialidad
    desc_guardas = records_svc.obtener_turnos_sin_cubrir("04/11/2026", especialidad="Guarda Eléctrico")
    assert all(d["especialidad"] == "Guarda Eléctrico" for d in desc_guardas)
    assert any(int(d["legajo"]) == 3003 for d in desc_guardas)


def test_asignar_jornada_textarea(qapp):
    """4. Textarea para observaciones en AsignarJornadaDialog."""
    dlg = AsignarJornadaDialog(
        legajo=999,
        nombre="Test Agente",
        fecha_iso="2026-11-05",
        dia_semana="Jueves",
    )
    # Debe ser QPlainTextEdit (multilínea)
    from PySide6.QtWidgets import QPlainTextEdit
    assert isinstance(dlg.txt_obs, QPlainTextEdit)

    texto_multilinea = "Línea 1: Relevo autorizado\nLínea 2: Diagrama especial\nLínea 3: Sin franco compensatorio"
    dlg.txt_obs.setPlainText(texto_multilinea)
    assert dlg.txt_obs.toPlainText() == texto_multilinea
    # Compatibilidad con .text() y .setText()
    assert dlg.txt_obs.text() == texto_multilinea


def test_calendario_mensual_ciclo_novedad_propia(fresh_db):
    """5. Calendario mensual de ciclo reflejando novedad propia activa."""
    records_svc = RecordsService(fresh_db)

    with fresh_db.write_transaction() as conn:
        conn.execute("INSERT OR REPLACE INTO empleados (legajo, apellidos_nombres, especialidad, dotacion, turnos, franco, activo, actualizado_en) VALUES (7777, 'Ciclo Propio', 'Conductor Eléctrico', 'PC', 'CICLO-2', 'Martes', 1, datetime('now', 'localtime'))")
        conn.execute(
            """INSERT INTO novedades (registrado_en, legajo, apellidos_nombres, especialidad, dotacion, turnos, franco, novedad, fecha_inicio, fecha_fin, observaciones, referencia_estacion, supervisor, activo)
               VALUES (datetime('now', 'localtime'), 7777, 'Ciclo Propio', 'Conductor Eléctrico', 'PC', 'CICLO-2', 'Martes', 'Licencia Médica', '2026-11-10', '2026-11-12', 'Reposo 72 hs por traumatismo', 'PC', 'Sup 1', 1)"""
        )

    calendario = records_svc.obtener_calendario_mensual_ciclo(7777, 2026, 11)
    dias_novedad = [d for d in calendario if d["dia"] in (10, 11, 12)]
    assert len(dias_novedad) == 3

    for d in dias_novedad:
        assert "Novedad" in d["estado"]
        assert "Licencia Médica" in d["estado"]
        assert d["toma"] == "-"
        assert d["deja"] == "-"
        assert d["horas_trabajadas"] == "-"
        assert "Reposo 72 hs" in d["observaciones"]


def test_permisos_granulares_subpestanas(qapp, fresh_db):
    """6. Permisos y control granular de pestañas y sub-pestañas."""
    records_svc = RecordsService(fresh_db)
    excel_svc = ExcelService(fresh_db)

    # Caso A: Usuario solo con permiso 'novedades.ver'
    view_a = NovedadesCambiosView(
        records_svc,
        excel_svc,
        None,
        None,
        tiene_permiso_fn=lambda p: p == "novedades.ver",
    )
    assert view_a.subtabs.count() == 1
    assert "Novedades" in view_a.subtabs.tabText(0)

    # Caso B: Usuario con 'cambios_turno.ver' y 'ordenamientos.ver' pero sin 'novedades.ver'
    view_b = NovedadesCambiosView(
        records_svc,
        excel_svc,
        None,
        None,
        tiene_permiso_fn=lambda p: p in ("cambios_turno.ver", "ordenamientos.ver"),
    )
    assert view_b.subtabs.count() == 2
    assert "Cambios de Turno" in view_b.subtabs.tabText(0)
    assert "Ordenamientos" in view_b.subtabs.tabText(1)


def test_importar_paradas_tren_excel_multitren(fresh_db, tmp_path):
    """8. Importación de paradas con múltiples trenes y hojas desde un único archivo Excel."""
    import openpyxl
    records_svc = RecordsService(fresh_db)

    wb = openpyxl.Workbook()
    # Hoja 1: Formato matricial de cuadro horario
    ws1 = wb.active
    ws1.title = "Matriz_Trenes"
    ws1.append(["Estación", "1201", "1203", "1205"])
    ws1.append(["Plaza Constitución", "05:00", "05:20", "05:40"])
    ws1.append(["Temperley", "05:25", "05:45", "06:05"])
    ws1.append(["Llavallol", "05:35", "05:55", "06:15"])

    # Hoja 2: Formato columnar tradicional con otro tren
    ws2 = wb.create_sheet(title="Columnar_Tren")
    ws2.append(["TREN", "ESTACION", "HORA", "ORDEN"])
    ws2.append(["2401", "Plaza Constitución", "06:00", 1])
    ws2.append(["2401", "Lanús", "06:12", 2])
    ws2.append(["2401", "Banfield", "06:20", 3])

    xlsx_file = tmp_path / "multi_trenes.xlsx"
    wb.save(xlsx_file)

    insertados = records_svc.importar_paradas_tren_excel(str(xlsx_file), itinerario_default="H45")
    # 3 paradas * 3 trenes en hoja 1 = 9 paradas + 3 paradas en hoja 2 = 12 paradas totales
    assert insertados == 12

    paradas_1201 = records_svc.obtener_estaciones_y_horarios_por_tren("1201")
    assert len(paradas_1201) == 3
    assert paradas_1201[0]["estacion"] == "Plaza Constitución"
    assert paradas_1201[0]["hora"] == "05:00"

    paradas_2401 = records_svc.obtener_estaciones_y_horarios_por_tren("2401")
    assert len(paradas_2401) == 3
    assert paradas_2401[1]["estacion"] == "Lanús"
    assert paradas_2401[1]["hora"] == "06:12"
