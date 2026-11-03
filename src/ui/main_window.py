"""Ventana principal de RENO v2.0.0 (Estándar visual y arquitectónico CRONO-PCT).

Integra navegación por pestañas de alto rendimiento, barra de estado con indicador
híbrido de sincronización Supabase, timers de actividad, notificaciones toast y control RBAC.
"""

import sys
import time
from typing import Any, Dict, Optional
from datetime import datetime

from PySide6.QtCore import QEvent, QPoint, QSize, Qt, QTimer
from PySide6.QtGui import QAction, QColor, QFont, QIcon
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMenu,
    QMenuBar,
    QMessageBox,
    QPushButton,
    QStatusBar,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from src.core.config import APP_NAME, APP_VERSION
from src.core.logger import get_logger
from src.core.settings import AppSettings
from src.data.database import DatabaseManager
from src.data.sync.outbox_queue import OutboxQueue
from src.data.sync.supabase_sync import SupabaseSyncClient
from src.services.auth_service import AuthService
from src.services.backup_service import BackupService
from src.services.excel_service import ExcelService
from src.services.pdf_service import PDFService
from src.services.records_service import RecordsService
from src.services.word_service import WordService
from src.ui.views.cambios_turno_view import CambiosTurnoView
from src.ui.views.dashboard_view import DashboardView
from src.ui.views.dialogs.admin_dialogs import (
    AdminAuditoriaDialog,
    AdminBackupsDialog,
    AdminCatalogosDialog,
    AdminConfigDialog,
    AdminEliminadosDialog,
    AdminEmpleadosDialog,
    AdminRolesDialog,
    AdminUsuariosDialog,
)
from src.ui.views.dialogs.toast_widget import ToastWidget
from src.ui.views.novedades_view import NovedadesView
from src.ui.views.personal_ciclo_view import PersonalCicloView
from src.ui.workers.task_runner import run_async

logger = get_logger("MainWindow")


class MainWindow(QMainWindow):
    """Ventana principal de la aplicación RENO v2.0.0."""

    def __init__(
        self,
        db: DatabaseManager,
        records_service: RecordsService,
        auth_service: AuthService,
        excel_service: ExcelService,
        pdf_service: PDFService,
        word_service: WordService,
        backup_service: BackupService,
        usuario_actual: Dict[str, Any],
        supabase_sync: Optional[SupabaseSyncClient] = None,
        parent=None,
    ):
        super().__init__(parent)
        self.db = db
        self.records_service = records_service
        self.auth_service = auth_service
        self.excel_service = excel_service
        self.pdf_service = pdf_service
        self.word_service = word_service
        self.backup_service = backup_service
        self.usuario_actual = usuario_actual
        self.supabase_sync = supabase_sync

        self.settings = AppSettings()
        self.outbox_queue = OutboxQueue(self.db)
        self.toast = ToastWidget()

        self._ultimo_movimiento = time.time()
        self._sesion_minutos = 30
        self._sin_revisar_count = 0

        self.setWindowTitle(f"{APP_NAME} v{APP_VERSION} - Operaciones")
        self.resize(1280, 800)

        # Cargar hoja de estilo corporativa
        self._aplicar_estilos()

        # Configurar UI
        self._init_menu_bar()
        self._init_tabs()
        self._init_status_bar()

        # Restaurar geometría
        self._restaurar_geometria()

        # Instalar filtro de eventos global para detectar inactividad
        QApplication.instance().installEventFilter(self)

        # Configurar temporizadores
        self._init_timers()

        # Primer refresco
        self.recargar_todo()
        self._verificar_backup_diario()

    def _aplicar_estilos(self) -> None:
        try:
            from pathlib import Path
            qss_path = Path(__file__).resolve().parent / "theme" / "crono_style.qss"
            if qss_path.exists():
                with open(qss_path, "r", encoding="utf-8") as f:
                    self.setStyleSheet(f.read())
        except Exception as e:
            logger.warning(f"No se pudo cargar la hoja de estilos QSS: {e}")

    def _tiene_permiso(self, codigo: str) -> bool:
        user_id = self.usuario_actual.get("id")
        if not user_id:
            return False
        return self.auth_service.tiene_permiso(user_id, codigo)

    def _init_menu_bar(self) -> None:
        menubar = self.menuBar()

        # Menú Archivo
        menu_archivo = menubar.addMenu("Archivo")
        if self._tiene_permiso("novedades.exportar") or self._tiene_permiso("cambios_turno.exportar"):
            act_export_all = menu_archivo.addAction("📊 Exportar Todo a Excel...")
            act_export_all.triggered.connect(self._exportar_todo_excel)

        if self._tiene_permiso("backup.gestionar"):
            act_backup_quick = menu_archivo.addAction("💾 Copia de Seguridad Rápida")
            act_backup_quick.triggered.connect(self._backup_rapido)

        menu_archivo.addSeparator()
        act_logout = menu_archivo.addAction("🔒 Cerrar Sesión")
        act_logout.triggered.connect(self._cerrar_sesion)
        act_salir = menu_archivo.addAction("✕ Salir")
        act_salir.triggered.connect(self.close)

        # Menú Administración
        menu_admin = menubar.addMenu("Administración")

        if self._tiene_permiso("usuarios.administrar"):
            act_usuarios = menu_admin.addAction("👤 Usuarios")
            act_usuarios.triggered.connect(lambda: AdminUsuariosDialog(self.auth_service, parent=self).exec())

        if self._tiene_permiso("roles.administrar"):
            act_roles = menu_admin.addAction("🛡️ Roles y Permisos")
            act_roles.triggered.connect(lambda: AdminRolesDialog(self.auth_service, parent=self).exec())

        if self._tiene_permiso("empleados.administrar"):
            act_emp = menu_admin.addAction("👷 Empleados (BASE)")
            act_emp.triggered.connect(lambda: AdminEmpleadosDialog(self.records_service, self.excel_service, parent=self).exec())

        if self._tiene_permiso("tiposNovedad.administrar") or self._tiene_permiso("dotaciones.administrar"):
            act_cat = menu_admin.addAction("📋 Catálogos Operativos")
            act_cat.triggered.connect(lambda: AdminCatalogosDialog(self.records_service, self.excel_service, parent=self).exec())

        if self._tiene_permiso("sesion.configurar"):
            act_cfg = menu_admin.addAction("⚙️ Configuración del Sistema")
            act_cfg.triggered.connect(lambda: AdminConfigDialog(self.db, parent=self).exec())

        if self._tiene_permiso("backup.gestionar"):
            act_bak = menu_admin.addAction("💾 Copias de Seguridad")
            act_bak.triggered.connect(lambda: AdminBackupsDialog(self.backup_service, parent=self).exec())

        if self._tiene_permiso("registros.recuperar"):
            act_del = menu_admin.addAction("🗑️ Registros Eliminados")
            act_del.triggered.connect(lambda: AdminEliminadosDialog(self.records_service, self.usuario_actual, parent=self).exec())

        if self._tiene_permiso("auditoria.ver"):
            act_aud = menu_admin.addAction("📜 Trazabilidad y Auditoría")
            act_aud.triggered.connect(lambda: AdminAuditoriaDialog(self.records_service, parent=self).exec())

        # Menú de Nuevos Registros
        self.menu_nuevos = menubar.addMenu("🔔 Registros Nuevos (0)")
        self.menu_nuevos.aboutToShow.connect(self._on_abrir_menu_nuevos)

        # Menú Ayuda
        menu_ayuda = menubar.addMenu("Ayuda")
        act_about = menu_ayuda.addAction("ℹ️ Acerca de RENO...")
        act_about.triggered.connect(self._acerca_de)

    def _init_tabs(self) -> None:
        self.tabs = QTabWidget()
        self.tabs.setDocumentMode(True)

        # 1. Panel de Control
        self.tab_dashboard = DashboardView(
            self.records_service,
            self.excel_service,
            self.pdf_service,
            parent=self,
        )
        self.tabs.addTab(self.tab_dashboard, "📈 Panel de Control")

        # 2. Novedades Operativas
        if self._tiene_permiso("novedades.ver"):
            self.tab_novedades = NovedadesView(
                self.records_service,
                self.excel_service,
                self.pdf_service,
                usuario_actual=self.usuario_actual,
                tiene_permiso_fn=self._tiene_permiso,
                parent=self,
            )
            self.tab_novedades.registro_modificado.connect(self._on_registro_modificado)
            self.tabs.addTab(self.tab_novedades, "📋 Novedades Operativas")

        # 3. Cambios de Turno
        if self._tiene_permiso("cambios_turno.ver"):
            self.tab_cambios = CambiosTurnoView(
                self.records_service,
                self.excel_service,
                self.pdf_service,
                self.word_service,
                usuario_actual=self.usuario_actual,
                tiene_permiso_fn=self._tiene_permiso,
                parent=self,
            )
            self.tab_cambios.registro_modificado.connect(self._on_registro_modificado)
            self.tabs.addTab(self.tab_cambios, "🔄 Cambios de Turno")

        # 4. Personal de Ciclo
        if self._tiene_permiso("reemplazos.ver"):
            self.tab_ciclo = PersonalCicloView(
                self.records_service,
                self.pdf_service,
                self.excel_service,
                usuario_actual=self.usuario_actual,
                parent=self,
            )
            self.tabs.addTab(self.tab_ciclo, "⏱️ Personal de Ciclo")

        self.setCentralWidget(self.tabs)

    def _init_status_bar(self) -> None:
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)

        # Usuario actual
        u_nom = self.usuario_actual.get("nombre") or self.usuario_actual.get("username") or "Usuario"
        self.lbl_user = QLabel(f"👤 {u_nom} ({self.usuario_actual.get('username')})")
        self.lbl_user.setStyleSheet("font-weight: 600; padding: 0 8px;")
        self.status_bar.addWidget(self.lbl_user)

        # Ruta base de datos
        db_name = self.db.db_path.name
        self.lbl_db = QLabel(f"📁 BD: {db_name} [{self.db.journal_mode}]")
        self.lbl_db.setStyleSheet("color: #6B7280; padding: 0 8px;")
        self.status_bar.addWidget(self.lbl_db)

        # Indicador de sincronización híbrida
        self.lbl_sync = QLabel("⚪ Offline")
        self.lbl_sync.setStyleSheet("padding: 2px 8px; font-weight: bold; border-radius: 4px; background-color: #E5E7EB; color: #4B5563;")
        self.status_bar.addPermanentWidget(self.lbl_sync)

        # Botón sincronizar ahora
        btn_sync = QPushButton("🔄 Sincronizar")
        btn_sync.setFixedHeight(22)
        btn_sync.setObjectName("btnSecondary")
        btn_sync.clicked.connect(self.sincronizar_ahora)
        self.status_bar.addPermanentWidget(btn_sync)

    def _init_timers(self) -> None:
        # Timer de auto-refresh cada 60s
        self.timer_refresh = QTimer(self)
        self.timer_refresh.setInterval(60 * 1000)
        self.timer_refresh.timeout.connect(self.recargar_todo)
        self.timer_refresh.start()

        # Timer de sincronización en segundo plano cada 120s
        self.timer_sync = QTimer(self)
        self.timer_sync.setInterval(120 * 1000)
        self.timer_sync.timeout.connect(self.sincronizar_ahora)
        self.timer_sync.start()

        # Timer de verificación de inactividad cada 30s
        self.timer_inactividad = QTimer(self)
        self.timer_inactividad.setInterval(30 * 1000)
        self.timer_inactividad.timeout.connect(self._verificar_inactividad)
        self.timer_inactividad.start()

    def eventFilter(self, watched, event) -> bool:
        if event.type() in (QEvent.MouseMove, QEvent.MouseButtonPress, QEvent.KeyPress):
            self._ultimo_movimiento = time.time()
        return super().eventFilter(watched, event)

    def _verificar_inactividad(self) -> None:
        inactivo_seg = time.time() - self._ultimo_movimiento
        limite_seg = self._sesion_minutos * 60
        if inactivo_seg >= limite_seg:
            logger.info("Sesión expirada por inactividad del usuario.")
            QMessageBox.warning(self, "Sesión Caducada", "Su sesión ha caducado por inactividad.")
            self._cerrar_sesion()

    def recargar_todo(self) -> None:
        """Refresca todas las vistas activas y recomputa notificaciones."""
        try:
            self._actualizar_notificaciones_nuevos()
            self.tab_dashboard.actualizar_dashboard(no_revisados_count=self._sin_revisar_count)

            if hasattr(self, "tab_novedades"):
                self.tab_novedades.recargar_datos()
            if hasattr(self, "tab_cambios"):
                self.tab_cambios.recargar_datos()
            if hasattr(self, "tab_ciclo"):
                self.tab_ciclo.consultar()

            self._actualizar_badge_sync()
        except Exception as e:
            logger.error(f"Error en recarga periódica: {e}")

    def _on_registro_modificado(self) -> None:
        self.tab_dashboard.actualizar_dashboard(no_revisados_count=self._sin_revisar_count)
        self.sincronizar_ahora()

    def _actualizar_notificaciones_nuevos(self) -> None:
        ultimo_ingreso = self.usuario_actual.get("ultimo_ingreso")
        if not ultimo_ingreso:
            return

        u_id = self.usuario_actual.get("id")
        win_u = self.usuario_actual.get("username")

        with self.db.read_connection() as conn:
            # Novedades cargadas por otros usuarios desde el último ingreso
            novs = conn.execute(
                """SELECT id, novedad, apellidos_nombres, registrado_en FROM novedades
                   WHERE activo=1 AND registrado_en >= ?
                     AND (usuario_id != ? OR (usuario_id IS NULL AND usuario_windows != ?))
                   ORDER BY id DESC LIMIT 10""",
                (ultimo_ingreso, u_id, win_u),
            ).fetchall()

            # Cambios cargados por otros
            cams = conn.execute(
                """SELECT id, legajo_1, legajo_2, registrado_en FROM cambios_turno
                   WHERE activo=1 AND registrado_en >= ?
                     AND (usuario_id != ? OR (usuario_id IS NULL AND usuario_windows != ?))
                   ORDER BY id DESC LIMIT 10""",
                (ultimo_ingreso, u_id, win_u),
            ).fetchall()

        total = len(novs) + len(cams)
        self._sin_revisar_count = total
        self.menu_nuevos.setTitle(f"🔔 Registros Nuevos ({total})")

        # Toast si hay nuevos y las notificaciones están activadas
        if total > 0 and self.settings.get_bool("notificaciones_activo", True):
            self.toast.show_message(
                "Nuevos Registros Operativos",
                f"Hay {total} novedades o cambios cargados desde su último acceso.",
            )

    def _on_abrir_menu_nuevos(self) -> None:
        self.menu_nuevos.clear()
        ultimo_ingreso = self.usuario_actual.get("ultimo_ingreso")
        u_id = self.usuario_actual.get("id")
        win_u = self.usuario_actual.get("username")

        with self.db.read_connection() as conn:
            novs = conn.execute(
                """SELECT id, novedad, apellidos_nombres, registrado_en FROM novedades
                   WHERE activo=1 AND registrado_en >= ?
                     AND (usuario_id != ? OR (usuario_id IS NULL AND usuario_windows != ?))
                   ORDER BY id DESC LIMIT 15""",
                (ultimo_ingreso, u_id, win_u),
            ).fetchall()

        if not novs:
            act_vacia = self.menu_nuevos.addAction("Sin nuevos registros pendientes")
            act_vacia.setEnabled(False)
        else:
            for n in novs:
                txt = f"#{n['id']} - {n['novedad']} ({n['apellidos_nombres']}) [{n['registrado_en']}]"
                self.menu_nuevos.addAction(txt)

    def sincronizar_ahora(self) -> None:
        if not self.supabase_sync:
            self._actualizar_badge_sync()
            return

        def _do_sync():
            return self.supabase_sync.sincronizar_completo()

        def _on_success(res):
            logger.info("Sincronización Supabase completada con éxito.")
            self._actualizar_badge_sync(sincronizado=True)

        def _on_error(err, tb):
            logger.warning(f"Sincronización fallida (modo offline activo): {err}")
            self._actualizar_badge_sync(sincronizado=False, error=True)

        self.lbl_sync.setText("🟡 Sincronizando...")
        self.lbl_sync.setStyleSheet("padding: 2px 8px; font-weight: bold; border-radius: 4px; background-color: #FEF3C7; color: #92400E;")
        run_async(_do_sync, on_success=_on_success, on_error=_on_error)

    def _actualizar_badge_sync(self, sincronizado: bool = False, error: bool = False) -> None:
        pendientes = self.outbox_queue.contar_pendientes()
        if not self.supabase_sync:
            self.lbl_sync.setText(f"⚪ Local (Pendientes: {pendientes})")
            self.lbl_sync.setStyleSheet("padding: 2px 8px; font-weight: bold; border-radius: 4px; background-color: #E5E7EB; color: #4B5563;")
            return

        if error:
            self.lbl_sync.setText(f"🔴 Offline ({pendientes} en cola)")
            self.lbl_sync.setStyleSheet("padding: 2px 8px; font-weight: bold; border-radius: 4px; background-color: #FEE2E2; color: #991B1B;")
        elif pendientes == 0:
            self.lbl_sync.setText("🟢 Sincronizado")
            self.lbl_sync.setStyleSheet("padding: 2px 8px; font-weight: bold; border-radius: 4px; background-color: #DCFCE7; color: #166534;")
        else:
            self.lbl_sync.setText(f"🟡 Pendientes: {pendientes}")
            self.lbl_sync.setStyleSheet("padding: 2px 8px; font-weight: bold; border-radius: 4px; background-color: #FEF3C7; color: #92400E;")

    def _exportar_todo_excel(self) -> None:
        default_name = f"RENO_Completo_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        file_path, _ = QFileDialog.getSaveFileName(self, "Exportar Base Completa", default_name, "Excel (*.xlsx)")
        if not file_path:
            return
        try:
            self.excel_service.export_database(file_path)
            QMessageBox.information(self, "Exportación", f"Base de datos exportada con éxito en:\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Fallo al exportar:\n{e}")

    def _backup_rapido(self) -> None:
        try:
            ruta = self.backup_service.crear_backup()
            QMessageBox.information(self, "Backup", f"Copia de seguridad consistente creada en:\n{ruta}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo crear el backup:\n{e}")

    def _verificar_backup_diario(self) -> None:
        try:
            creado = self.backup_service.ejecutar_backup_diario_automatico()
            if creado:
                logger.info(f"Backup automático diario creado al iniciar: {creado}")
        except Exception as e:
            logger.warning(f"No se pudo ejecutar backup automático diario: {e}")

    def _cerrar_sesion(self) -> None:
        self.close()

    def _acerca_de(self) -> None:
        QMessageBox.about(
            self,
            f"Acerca de {APP_NAME}",
            f"<b>{APP_NAME}</b><br/>"
            f"Versión: <b>{APP_VERSION}</b><br/>"
            f"Estándar visual y arquitectónico: <b>CRONO-PCT</b><br/>"
            f"Persistencia: SQLite3 (WAL/DELETE) + Supabase Sync<br/>"
            f"Desarrollado para Control Operativo y Registro de Novedades.",
        )

    def closeEvent(self, event) -> None:
        self.settings.set_window_geometry("MainWindow", self.saveGeometry())
        super().closeEvent(event)

    def _restaurar_geometria(self) -> None:
        geom = self.settings.get_window_geometry("MainWindow")
        if geom:
            self.restoreGeometry(geom)
        else:
            self.showMaximized()
