"""Pruebas automatizadas de arranque e integración de la aplicación principal (main.py + MainWindow)."""

from __future__ import annotations

from pathlib import Path
from PySide6.QtWidgets import QApplication

import main
from src.core.config import (
    APP_NAME,
    APP_VERSION,
    cargar_configuracion_json,
    obtener_ruta_base_datos,
)
from src.core.settings import AppSettings
from src.data.database import DatabaseManager
from src.data.migrations import init_database_schema
from src.data.sync.supabase_sync import SupabaseSyncClient
from src.domain.enums import PERMISOS_BASE
from src.services.auth_service import AuthService
from src.services.backup_service import BackupService
from src.services.excel_service import ExcelService
from src.services.pdf_service import PDFService
from src.services.records_service import RecordsService
from src.services.word_service import WordService
from src.ui.main_window import MainWindow
from src.ui.views.dialogs.login_dialog import LoginDialog


def test_app_startup_and_main_window_instantiation(tmp_path: Path):
    """Verifica el contrato completo de inicio de main.py, LoginDialog y MainWindow."""
    app = QApplication.instance() or QApplication([])

    assert APP_NAME == "RENO"
    assert APP_VERSION == "2.0.0"

    cfg = cargar_configuracion_json()
    assert isinstance(cfg, dict)

    sqlite_path, _xlsx_path = obtener_ruta_base_datos()
    assert isinstance(sqlite_path, Path)

    # Inicializar BD temporal aislada
    db_file = tmp_path / "test_startup.sqlite"
    db = DatabaseManager(db_file, journal_mode="WAL")
    init_database_schema(db)

    records_service = RecordsService(db)
    auth_service = AuthService(db)
    excel_service = ExcelService(db)
    pdf_service = PDFService()
    word_service = WordService()
    backup_service = BackupService(db)

    auth_service.inicializar_permisos(PERMISOS_BASE)

    # Verificar LoginDialog en modo primer inicio
    login_dlg = LoginDialog(auth_service)
    assert login_dlg._es_primer_inicio is True

    # Crear administrador inicial y autenticar
    admin_id = auth_service.crear_administrador_inicial("admin", "Admin123!", legajo=1000)
    assert admin_id > 0
    usuario_actual = auth_service.autenticar("admin", "Admin123!")
    assert usuario_actual is not None
    assert usuario_actual["username"] == "admin"

    # Verificar cliente SupabaseSyncClient
    sync_client = SupabaseSyncClient(
        db,
        outbox=None,
        supabase_url="https://example.supabase.co",
        supabase_key="dummy_key",
    )
    assert hasattr(sync_client, "sincronizar_completo")

    # Verificar AppSettings
    settings = AppSettings()
    assert isinstance(settings.get_bool("notificaciones_activo", True), bool)

    # Instanciar MainWindow completa
    win = MainWindow(
        db=db,
        records_service=records_service,
        auth_service=auth_service,
        excel_service=excel_service,
        pdf_service=pdf_service,
        word_service=word_service,
        backup_service=backup_service,
        usuario_actual=usuario_actual,
        supabase_sync=None,
    )

    assert win.tabs.count() == 7
    assert hasattr(win, "tab_dashboard")
    assert hasattr(win, "tab_buscador_turnos")
    assert hasattr(win, "tab_novedades_cambios")
    assert hasattr(win, "tab_turnos_sin_cubrir")
    assert hasattr(win, "tab_empleados")
    assert hasattr(win, "tab_ciclo")
    assert hasattr(win, "tab_administracion")
    assert win.tab_novedades is not None
    assert win.tab_cambios is not None

    # Detener timers y cerrar limpiamente
    win.timer_refresh.stop()
    win.timer_sync.stop()
    win.timer_inactividad.stop()
    win.close()


def test_inactivity_handling(tmp_path: Path):
    """Verifica que la inactividad dispare el diálogo de login sin cerrar abruptamente la app."""
    import time
    app = QApplication.instance() or QApplication([])

    db_file = tmp_path / "test_inactivity.sqlite"
    db = DatabaseManager(db_file, journal_mode="DELETE")
    init_database_schema(db)

    auth = AuthService(db)
    auth.inicializar_permisos(PERMISOS_BASE)
    auth.crear_administrador_inicial("admin", "Admin123!", legajo=1000)
    user = auth.autenticar("admin", "Admin123!")

    win = MainWindow(
        db=db,
        records_service=RecordsService(db),
        auth_service=auth,
        excel_service=ExcelService(db),
        pdf_service=PDFService(),
        word_service=WordService(),
        backup_service=BackupService(db),
        usuario_actual=user,
        supabase_sync=None,
    )

    login_invoked = []
    def mock_mostrar_login():
        login_invoked.append(True)

    win._mostrar_login_por_inactividad = mock_mostrar_login

    # Simular tiempo transcurrido superior al límite de inactividad
    win._ultimo_movimiento = time.time() - (win._sesion_minutos * 60 + 10)
    win._verificar_inactividad()

    assert len(login_invoked) == 1
    assert win.timer_inactividad.isActive() is False

    win.timer_refresh.stop()
    win.timer_sync.stop()
    win.close()


def test_database_migration_from_legacy_schema(tmp_path: Path):
    """Verifica que la migración sobre un esquema heredado de v1 no falle al añadir columnas."""
    import sqlite3

    db_legacy_file = tmp_path / "legacy_reno.sqlite"
    # Crear tablas como estaban en v1 (sin sync_status, created_at, updated_at)
    conn = sqlite3.connect(db_legacy_file)
    conn.execute(
        """CREATE TABLE empleados (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            legajo INTEGER UNIQUE NOT NULL,
            apellidos_nombres TEXT NOT NULL,
            especialidad TEXT NOT NULL,
            dotacion TEXT NOT NULL,
            turnos TEXT NOT NULL,
            franco TEXT NOT NULL,
            activo INTEGER NOT NULL DEFAULT 1
        );"""
    )
    conn.execute(
        """CREATE TABLE tipos_novedad (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT UNIQUE NOT NULL,
            activo INTEGER NOT NULL DEFAULT 1
        );"""
    )
    conn.execute(
        """CREATE TABLE novedades (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            registrado_en TEXT NOT NULL,
            legajo INTEGER NOT NULL,
            apellidos_nombres TEXT NOT NULL,
            especialidad TEXT NOT NULL,
            dotacion TEXT NOT NULL,
            turnos TEXT NOT NULL,
            franco TEXT NOT NULL,
            novedad TEXT NOT NULL,
            fecha_inicio TEXT NOT NULL,
            fecha_fin TEXT,
            referencia_estacion TEXT,
            supervisor TEXT,
            observaciones TEXT,
            usuario_windows TEXT,
            usuario_id INTEGER,
            activo INTEGER NOT NULL DEFAULT 1
        );"""
    )
    conn.execute("INSERT INTO empleados (legajo, apellidos_nombres, especialidad, dotacion, turnos, franco) VALUES (1234, 'Perez Juan', 'Conductor', 'PC', 'T1', 'DOM');")
    conn.commit()
    conn.close()

    # Ejecutar la migración sobre el esquema heredado
    db = DatabaseManager(db_legacy_file, journal_mode="WAL")
    init_database_schema(db)

    # Verificar que las columnas se hayan añadido y los registros anteriores tengan timestamp poblado
    with db.read_connection() as c:
        emp_cols = {row[1] for row in c.execute("PRAGMA table_info(empleados);").fetchall()}
        assert "created_at" in emp_cols
        assert "updated_at" in emp_cols
        assert "sync_status" in emp_cols
        assert "franco_horario" in emp_cols
        assert "jornada_horas" in emp_cols

        row = c.execute("SELECT created_at, updated_at, sync_status FROM empleados WHERE legajo = 1234;").fetchone()
        assert row["created_at"] is not None
        assert row["updated_at"] is not None
        assert row["sync_status"] == "pending"

