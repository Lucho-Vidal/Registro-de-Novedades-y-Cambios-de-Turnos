"""Pruebas automatizadas de la nueva pestaña Buscador de Turnos e Itinerarios, tema oscuro y cambio de clave."""

from unittest.mock import patch
from pathlib import Path
from PySide6.QtWidgets import QApplication, QMessageBox

from src.data.database import DatabaseManager
from src.data.migrations import init_database_schema
from src.services.auth_service import AuthService
from src.services.excel_service import ExcelService
from src.services.records_service import RecordsService
from src.domain.enums import PERMISOS_BASE
from src.ui.views.buscador_turnos_view import BuscadorTurnosView
from src.ui.views.dialogs.cambiar_password_dialog import CambiarMiPasswordDialog


def test_buscador_turnos_tab_and_coverage(tmp_path: Path):
    app = QApplication.instance() or QApplication([])

    db_file = tmp_path / "test_buscador_tab.sqlite"
    db = DatabaseManager(db_file, journal_mode="WAL")
    init_database_schema(db)

    svc = RecordsService(db)
    auth = AuthService(db)
    excel_svc = ExcelService(db)
    auth.inicializar_permisos(PERMISOS_BASE)

    # 1. Importar itinerario Indice_45.xlsx
    excel_path = Path("Indice_45.xlsx")
    assert excel_path.exists()
    t_count, v_count = svc.importar_itinerario_excel(str(excel_path))
    assert t_count == 218
    assert v_count >= 900

    # 2. Agregar conductor titular para 401
    with db.write_transaction() as conn:
        conn.execute(
            """INSERT INTO empleados (legajo, apellidos_nombres, especialidad, dotacion, turnos, franco, activo, actualizado_en)
               VALUES (14215, 'FETTER Marcos', 'Conductor Electrico', 'PC', '401', 'Domingo', 1, '2026-11-01 00:00:00');"""
        )

    # 3. Registrar novedad 5753 abierta para titular 20189 con reemplazo 12479
    with db.write_transaction() as conn:
        conn.execute(
            """INSERT INTO empleados (legajo, apellidos_nombres, especialidad, dotacion, turnos, franco, activo, actualizado_en)
               VALUES (20189, 'FERNANDEZ Juan Ariel', 'Conductor Electrico', 'PC', '401', 'Lunes', 1, '2026-08-01 00:00:00'),
                      (12479, 'ALEGRE Maximiliano Nicolas', 'Conductor Electrico', 'PC', 'CICLO 11', 'Domingo', 1, '2026-08-01 00:00:00');"""
        )
        conn.execute(
            """INSERT INTO novedades (id, registrado_en, legajo, apellidos_nombres, especialidad, dotacion, turnos, franco, novedad, fecha_inicio, fecha_fin, referencia_estacion, supervisor, observaciones, activo)
               VALUES (5753, '15/08/2026 10:00', 20189, 'FERNANDEZ Juan Ariel', 'Conductor Electrico', 'PC', '401', 'Lunes', 'Enfermo', '15/08/2026', '-', 'CARDOZO', 'TORRES', '', 1);"""
        )
        conn.execute(
            """INSERT INTO novedad_reemplazos (id, novedad_id, legajo_reemplazo, apellidos_nombres_reemplazo, turnos_reemplazo, franco_reemplazo, fecha_inicio, fecha_fin, creado_en, activo)
               VALUES (3, 5753, 12479, 'ALEGRE Maximiliano Nicolas', 'CICLO 11', 'Domingo', '15/08/2026', NULL, '15/08/2026 10:05', 1);"""
        )

    # 4. Consultar búsqueda dual por tren 2303 en 03/11/2026 (Martes)
    turnos_tren = svc.buscar_turnos_con_cobertura(
        criterio="2303",
        fecha_target_iso="2026-11-03",
        tipo_dia="Día Hábil (Lunes a Viernes)",
    )
    assert len(turnos_tren) >= 1
    t0 = turnos_tren[0]
    assert t0["codigo_turno"] == "401.1"
    # Debe figurar el reemplazo 12479 ALEGRE en lugar del titular enfermo FERNANDEZ
    assert "ALEGRE Maximiliano" in t0["titular"]
    assert "FERNANDEZ Juan Ariel" in t0["cobertura"]

    # 5. Obtener vueltas del turno 401.1 y múltiples turnos
    vueltas = svc.obtener_vueltas_de_turno(t0["id"])
    assert len(vueltas) == 4
    trenes = [v["nro_tren"] for v in vueltas]
    assert "2303" in trenes
    assert "1310" in trenes

    vueltas_todas = svc.obtener_vueltas_de_turnos([t0["id"]])
    assert len(vueltas_todas) == 4
    assert vueltas_todas[0]["codigo_turno"] == "401.1"

    # 6. Instanciar BuscadorTurnosView y verificar controles y vueltas apiladas
    view = BuscadorTurnosView(svc, excel_svc)
    assert view.combo_servicio.count() == 3
    assert view.combo_itinerarios.isEditable() is True

    # Búsqueda dual por turno 401
    view.txt_criterio.setText("401")
    view.consultar()
    assert len(view._turnos_resultados) > 0
    # La tabla de vueltas contiene todas las vueltas apiladas de los turnos encontrados
    assert view.model_vueltas.rowCount() == len(view._vueltas_todas)
    assert view.model_vueltas.rowCount() >= 4

    # 7. Verificar autocalculo de toma, deja, horas trabajadas y descanso en Calendario de Ciclo del reemplazo
    cal_12479 = svc.obtener_calendario_mensual_ciclo(12479, 2026, 11)
    # Martes 03/11/2026: el reemplazo hace el turno 401.1
    d_mar = next(d for d in cal_12479 if d["fecha"] == "03/11/2026")
    assert d_mar["estado"] == "Reemplazando"
    assert d_mar["toma"] == "18:55"
    assert d_mar["deja"] == "00:20"
    assert float(d_mar["horas_trabajadas"].replace(" hs", "")) > 5.0
    assert d_mar["hora_disponible"] != "-"

    # Lunes 02/11/2026: es franco del titular (Fernandez) -> Franco por Relevo
    d_lun = next(d for d in cal_12479 if d["fecha"] == "02/11/2026")
    assert "Franco" in d_lun["estado"]

    view.close()


def test_cambiar_password_dialog(tmp_path: Path):
    app = QApplication.instance() or QApplication([])
    db_file = tmp_path / "test_auth_pwd.sqlite"
    db = DatabaseManager(db_file, journal_mode="WAL")
    init_database_schema(db)
    auth = AuthService(db)

    user_id = auth.crear_administrador_inicial("operador", "ClaveVieja1!", legajo=1234)
    user_actual = auth.autenticar("operador", "ClaveVieja1!")
    assert user_actual is not None

    dlg = CambiarMiPasswordDialog(auth, user_actual)
    dlg.txt_actual.setText("ClaveVieja1!")
    dlg.txt_nueva.setText("ClaveNueva2026$")
    dlg.txt_confirm.setText("ClaveNueva2026$")
    with patch.object(QMessageBox, "information"):
        dlg._on_guardar()

    # Verificar que ahora autentique con la nueva contraseña
    nuevo_login = auth.autenticar("operador", "ClaveNueva2026$")
    assert nuevo_login is not None
    dlg.close()
