"""Punto de entrada principal para RENO v2.0.0 (Registro de Novedades y Cambios de Turnos).

Arquitectura desacoplada basada en PySide6-Essentials, persistencia SQLite híbrida con Supabase,
estándar visual CRONO-PCT y captura global de excepciones con logging rotativo.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication

from src.core.config import (
    APP_NAME,
    APP_VERSION,
    cargar_configuracion_json,
    obtener_ruta_base_datos,
)
from src.core.logger import get_logger, inicializar_logging
from src.core.proxy import aplicar_configuracion_proxy
from src.core.theme import aplicar_icono_aplicacion, aplicar_tema_visual
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

logger = get_logger("App")


def _configurar_app_user_model_id() -> None:
    """Asigna AppUserModelID en Windows para mostrar el icono propio en barra de tareas."""
    if sys.platform != "win32":
        return
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("TrenesArgentinos.RENO.Operaciones.v2")
    except Exception:
        pass


def main() -> int:
    # 1. Inicializar sistema de logging rotativo y captura de excepciones
    log_file = inicializar_logging()
    logger.info("Iniciando %s v%s...", APP_NAME, APP_VERSION)
    _configurar_app_user_model_id()

    # 2. Inicializar QApplication
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setApplicationVersion(APP_VERSION)
    app.setOrganizationName("Trenes Argentinos Operaciones")
    if not aplicar_icono_aplicacion(app):
        logger.warning("No se pudo cargar el icono de aplicación RENO.")

    # 3. Cargar configuración y determinar ruta de base de datos
    config_data = cargar_configuracion_json()
    aplicar_configuracion_proxy(config_data)
    journal_mode = config_data.get("sqlite_journal_mode", "WAL")
    db_path, xlsx_path = obtener_ruta_base_datos()
    logger.info("Ruta de base de datos activa: %s (Modo: %s)", db_path, journal_mode)

    # 4. Inicializar capa de datos y migraciones SQLite
    db = DatabaseManager(db_path, journal_mode=journal_mode)
    init_database_schema(db)

    # 5. Inicializar servicios de negocio
    records_service = RecordsService(db)
    auth_service = AuthService(db)
    excel_service = ExcelService(db)
    pdf_service = PDFService()
    word_service = WordService()
    backup_service = BackupService(db)

    # Inicializar catálogo de permisos del sistema
    auth_service.inicializar_permisos(PERMISOS_BASE)

    # Auto-importar plantilla BASE si la tabla empleados está vacía y existe un XLSX asociado
    if xlsx_path and xlsx_path.exists():
        try:
            if len(records_service.listar_empleados()) == 0:
                importados = excel_service.import_base_empleados(str(xlsx_path))
                logger.info("Se importaron %d empleados iniciales desde %s", importados, xlsx_path)
        except Exception as e:
            logger.warning("No se pudo auto-importar plantilla BASE desde %s: %s", xlsx_path, e)

    # 6. Inicializar cliente de sincronización Supabase si está configurado
    supabase_url = os.environ.get("SUPABASE_URL") or config_data.get("supabase_url")
    supabase_key = (
        os.environ.get("SUPABASE_KEY")
        or config_data.get("supabase_anon_key")
        or config_data.get("supabase_key")
    )
    supabase_sync = None

    if supabase_url and supabase_key:
        try:
            supabase_sync = SupabaseSyncClient(
                db,
                outbox=None,
                supabase_url=supabase_url,
                supabase_key=supabase_key,
            )
            logger.info("Cliente de sincronización Supabase conectado a %s", supabase_url)
        except Exception as e:
            logger.warning("No se pudo iniciar el cliente Supabase: %s", e)

    # 7. Diálogo de Login / Asistente de primer inicio
    # Aplicar tema visual persistido antes de abrir el diálogo de login
    try:
        from src.core.theme import aplicar_tema_visual
        aplicar_tema_visual(db=db)
    except Exception as e:
        logger.warning("No se pudo pre-aplicar el tema visual inicial: %s", e)

    login_dialog = LoginDialog(auth_service)
    if not login_dialog.exec():
        logger.info("Inicio de sesión cancelado por el usuario. Saliendo.")
        return 0

    usuario_actual = login_dialog.usuario_autenticado
    if not usuario_actual:
        logger.error("No se obtuvo información del usuario autenticado.")
        return 1

    logger.info("Usuario '%s' inició sesión correctamente.", usuario_actual.get("username"))

    # 8. Construir y lanzar la ventana principal
    main_window = MainWindow(
        db=db,
        records_service=records_service,
        auth_service=auth_service,
        excel_service=excel_service,
        pdf_service=pdf_service,
        word_service=word_service,
        backup_service=backup_service,
        usuario_actual=usuario_actual,
        supabase_sync=supabase_sync,
    )
    main_window.show()

    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
