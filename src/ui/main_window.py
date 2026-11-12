"""Ventana principal de RENO v2.0.0 (Estándar visual y arquitectónico CRONO-PCT).

Integra navegación por pestañas de alto rendimiento, barra de estado con indicador
híbrido de sincronización Supabase, timers de actividad, notificaciones toast y control RBAC.
"""

import sys
import time
from typing import Any, Dict, List, Optional
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
from src.core.theme import (
    TEMA_CLARO,
    TEMA_OSCURO,
    aplicar_tema_visual,
    guardar_tema_visual,
    obtener_tema_visual,
)
from src.data.database import DatabaseManager
from src.data.sync.outbox_queue import OutboxQueue
from src.data.sync.supabase_sync import SupabaseSyncClient
from src.services.auth_service import AuthService
from src.services.backup_service import BackupService
from src.services.excel_service import ExcelService
from src.services.pdf_service import PDFService
from src.services.records_service import RecordsService
from src.services.word_service import WordService
from src.ui.views.administracion_view import AdministracionView
from src.ui.views.buscador_turnos_view import BuscadorTurnosView
from src.ui.views.cambios_turno_view import CambiosTurnoView
from src.ui.views.dashboard_view import DashboardView
from src.ui.views.empleados_view import EmpleadosView
from src.ui.views.novedades_cambios_view import NovedadesCambiosView
from src.ui.views.novedades_view import NovedadesView
from src.ui.views.personal_ciclo_view import PersonalCicloView
from src.ui.views.turnos_sin_cubrir_view import TurnosSinCubrirView
from src.ui.views.dialogs.acerca_dialog import AcercaDialog
from src.ui.views.dialogs.actualizacion_dialog import ActualizacionDialog
from src.services.updater_service import (
    InfoActualizacion,
    VerificarActualizacionWorker,
)
from src.ui.views.dialogs.admin_dialogs import (
    AdminAuditoriaDialog,
    AdminBackupsDialog,
    AdminCatalogosDialog,
    AdminConfigDialog,
    AdminDestinatariosDialog,
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
        try:
            with self.db.read_connection() as conn:
                r_sm = conn.execute("SELECT valor FROM configuracion WHERE clave = 'sesion_minutos'").fetchone()
                if r_sm and r_sm[0]:
                    self._sesion_minutos = max(1, int(r_sm[0]))
        except Exception:
            pass
        self._sin_revisar_count = 0

        self.setWindowTitle(f"{APP_NAME} v{APP_VERSION} - Operaciones")
        self.resize(1280, 800)

        # Configurar UI
        self._init_menu_bar()
        self._init_tabs()
        self._init_status_bar()

        # Cargar hoja de estilo corporativa aplicada a toda la aplicación
        self._aplicar_estilos()

        # Restaurar geometría
        self._restaurar_geometria()

        # Instalar filtro de eventos global para detectar inactividad
        QApplication.instance().installEventFilter(self)

        # Configurar temporizadores
        self._init_timers()

        # Primer refresco
        self.recargar_todo()
        self._verificar_backup_diario()

    def _aplicar_estilos(self, tema: str | None = None) -> None:
        try:
            tema_activo = aplicar_tema_visual(tema=tema, db=self.db)
            if hasattr(self, "tab_dashboard"):
                self.tab_dashboard.actualizar_tema(tema_activo)
            if hasattr(self, "tab_buscador_turnos") and hasattr(self.tab_buscador_turnos, "actualizar_tema"):
                self.tab_buscador_turnos.actualizar_tema()
            if hasattr(self, "tab_turnos_sin_cubrir") and hasattr(self.tab_turnos_sin_cubrir, "table"):
                self.tab_turnos_sin_cubrir.table.viewport().update()
            if hasattr(self, "tab_administracion") and hasattr(self.tab_administracion, "actualizar_tema"):
                self.tab_administracion.actualizar_tema(tema_activo)
        except Exception as e:
            logger.warning(f"No se pudo cargar la hoja de estilos QSS: {e}")

    def cambiar_tema(self, tema: str) -> None:
        guardar_tema_visual(tema, self.db)
        QTimer.singleShot(0, lambda: self._ejecutar_cambiar_tema(tema))

    def _ejecutar_cambiar_tema(self, tema: str) -> None:
        self._aplicar_estilos(tema)
        self.update()
        self.repaint()

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

        # Submenú Tema Visual
        menu_tema = menu_archivo.addMenu("🎨 Tema Visual")
        act_tema_claro = menu_tema.addAction("☀️ Modo Claro")
        act_tema_claro.triggered.connect(lambda: self.cambiar_tema("claro"))
        act_tema_oscuro = menu_tema.addAction("🌙 Modo Oscuro Moderno")
        act_tema_oscuro.triggered.connect(lambda: self.cambiar_tema("oscuro"))

        # Cambio de contraseña del usuario activo
        act_pass = menu_archivo.addAction("🔑 Cambiar mi Contraseña...")
        act_pass.triggered.connect(self._cambiar_mi_password)

        act_actualizaciones = menu_archivo.addAction("🚀 Buscar Actualizaciones...")
        act_actualizaciones.triggered.connect(lambda: self.buscar_actualizaciones_manual())

        menu_archivo.addSeparator()
        act_cambiar_usuario = menu_archivo.addAction("👥 Cambiar Usuario...")
        act_cambiar_usuario.triggered.connect(self._cambiar_usuario)
        act_logout = menu_archivo.addAction("🔒 Cerrar Sesión")
        act_logout.triggered.connect(self._cerrar_sesion)
        act_salir = menu_archivo.addAction("✕ Salir")
        act_salir.triggered.connect(self.close)

        # Menú Administración
        tiene_menu_admin = (
            self._tiene_permiso("usuarios.administrar")
            or self._tiene_permiso("roles.administrar")
            or self._tiene_permiso("empleados.administrar")
            or self._tiene_permiso("tiposNovedad.administrar")
            or self._tiene_permiso("dotaciones.administrar")
            or self._tiene_permiso("sesion.configurar")
            or self._tiene_permiso("destinatarios_informe.administrar")
            or self._tiene_permiso("destinatarios_ordenamiento.administrar")
            or self._tiene_permiso("backup.gestionar")
            or self._tiene_permiso("registros.recuperar")
            or self._tiene_permiso("auditoria.ver")
        )

        if tiene_menu_admin:
            menu_admin = menubar.addMenu("Administración")

            if self._tiene_permiso("usuarios.administrar"):
                act_usuarios = menu_admin.addAction("👤 Usuarios")
                act_usuarios.triggered.connect(lambda: AdminUsuariosDialog(self.auth_service, parent=self).exec())

            if self._tiene_permiso("roles.administrar"):
                act_roles = menu_admin.addAction("🛡️ Roles y Permisos")
                act_roles.triggered.connect(lambda: AdminRolesDialog(self.auth_service, parent=self).exec())

            if self._tiene_permiso("empleados.administrar"):
                act_emp = menu_admin.addAction("👷 Empleados (BASE)")
                act_emp.triggered.connect(self._mostrar_empleados)

            if self._tiene_permiso("tiposNovedad.administrar") or self._tiene_permiso("dotaciones.administrar"):
                act_cat = menu_admin.addAction("📋 Catálogos Operativos")
                act_cat.triggered.connect(lambda: AdminCatalogosDialog(self.records_service, self.excel_service, parent=self).exec())

            if self._tiene_permiso("sesion.configurar"):
                act_cfg = menu_admin.addAction("⚙️ Configuración del Sistema")
                act_cfg.triggered.connect(
                    lambda: AdminConfigDialog(
                        self.db,
                        records_service=self.records_service,
                        usuario_actual=self.usuario_actual,
                        tiene_permiso_fn=self._tiene_permiso,
                        parent=self,
                    ).exec()
                )

            if self._tiene_permiso("destinatarios_informe.administrar") or self._tiene_permiso("destinatarios_ordenamiento.administrar"):
                act_dest = menu_admin.addAction("📧 Destinatarios de Correo")
                act_dest.triggered.connect(
                    lambda: AdminDestinatariosDialog(
                        self.records_service,
                        usuario_actual=self.usuario_actual,
                        tiene_permiso_fn=self._tiene_permiso,
                        parent=self,
                    ).exec()
                )

            if self._tiene_permiso("backup.gestionar"):
                act_bak = menu_admin.addAction("💾 Copias de Seguridad")
                act_bak.triggered.connect(lambda: AdminBackupsDialog(self.backup_service, parent=self).exec())

            if self._tiene_permiso("registros.recuperar"):
                act_del = menu_admin.addAction("🗑️ Registros Eliminados")
                act_del.triggered.connect(lambda: AdminEliminadosDialog(self.records_service, self.usuario_actual, parent=self).exec())

            if self._tiene_permiso("auditoria.ver"):
                act_aud = menu_admin.addAction("📜 Trazabilidad y Auditoría")
                act_aud.triggered.connect(lambda: AdminAuditoriaDialog(self.records_service, excel_service=self.excel_service, parent=self).exec())

        # Menú de Nuevos Registros
        self.menu_nuevos = menubar.addMenu("🔔 Registros Nuevos (0)")
        self.menu_nuevos.aboutToShow.connect(self._on_abrir_menu_nuevos)

    def _init_tabs(self) -> None:
        if hasattr(self, "tabs") and self.tabs is not None:
            try:
                self.tabs.currentChanged.disconnect(self._on_tab_changed)
            except Exception:
                pass
            while self.tabs.count() > 0:
                w = self.tabs.widget(0)
                self.tabs.removeTab(0)
                if w:
                    w.deleteLater()
            for attr in ("tab_dashboard", "tab_buscador_turnos", "tab_novedades_cambios",
                         "tab_turnos_sin_cubrir", "tab_empleados", "tab_ciclo", "tab_administracion"):
                if hasattr(self, attr):
                    delattr(self, attr)
        else:
            self.tabs = QTabWidget()
            self.tabs.setDocumentMode(True)
            self.setCentralWidget(self.tabs)

        # Botones de versión y auto-updater en la esquina superior derecha (Estándar CRONO-PCT)
        corner_widget = QWidget()
        corner_lay = QHBoxLayout(corner_widget)
        corner_lay.setContentsMargins(0, 0, 0, 0)
        corner_lay.setSpacing(6)

        self.btn_actualizar = QPushButton("")
        self.btn_actualizar.setObjectName("btnActualizarHeader")
        self.btn_actualizar.setCursor(Qt.PointingHandCursor)
        self.btn_actualizar.setVisible(False)
        self.btn_actualizar.clicked.connect(self._on_clic_btn_actualizar)
        corner_lay.addWidget(self.btn_actualizar)

        self.btn_version = QPushButton(f"RENO v{APP_VERSION}")
        self.btn_version.setObjectName("btnVersion")
        self.btn_version.setCursor(Qt.PointingHandCursor)
        self.btn_version.setToolTip("Ver arquitectura, información del sistema y changelog")
        self.btn_version.clicked.connect(self._acerca_de)
        corner_lay.addWidget(self.btn_version)

        self.tabs.setCornerWidget(corner_widget, Qt.TopRightCorner)

        # 1. Panel de Control
        self.tab_dashboard = DashboardView(
            self.records_service,
            self.excel_service,
            self.pdf_service,
            parent=self,
        )
        self.tab_dashboard.clic_sin_revisar.connect(self._on_clic_kpi_sin_revisar)
        self.tabs.addTab(self.tab_dashboard, "📈 Panel de Control")

        # 2. Buscador de Turnos e Itinerarios
        if self._tiene_permiso("turnos.ver"):
            self.tab_buscador_turnos = BuscadorTurnosView(
                self.records_service,
                self.excel_service,
                usuario_actual=self.usuario_actual,
                tiene_permiso_fn=self._tiene_permiso,
                parent=self,
            )
            self.tabs.addTab(self.tab_buscador_turnos, "🔍 Buscador de Turnos")

        # 3. Novedades y Cambios de Turno (Sub-pestañas unificadas)
        if (
            self._tiene_permiso("novedades.ver")
            or self._tiene_permiso("cambios_turno.ver")
            or self._tiene_permiso("ordenamientos.ver")
        ):
            self.tab_novedades_cambios = NovedadesCambiosView(
                self.records_service,
                self.excel_service,
                self.pdf_service,
                self.word_service,
                usuario_actual=self.usuario_actual,
                tiene_permiso_fn=self._tiene_permiso,
                parent=self,
            )
            self.tab_novedades_cambios.registro_modificado.connect(self._on_registro_modificado)
            self.tabs.addTab(self.tab_novedades_cambios, "📝 Novedades y Cambios")

        # 4. Turnos Operativos Sin Cubrir
        if self._tiene_permiso("novedades.ver"):
            self.tab_turnos_sin_cubrir = TurnosSinCubrirView(
                self.records_service,
                self.excel_service,
                usuario_actual=self.usuario_actual,
                parent=self,
            )
            self.tabs.addTab(self.tab_turnos_sin_cubrir, "🚨 Turnos Sin Cubrir")

        # 5. Empleados (BASE)
        if self._tiene_permiso("empleados.administrar") or self._tiene_permiso("novedades.ver"):
            self.tab_empleados = EmpleadosView(
                self.records_service,
                self.excel_service,
                usuario_actual=self.usuario_actual,
                parent=self,
            )
            self.tabs.addTab(self.tab_empleados, "👷 Empleados (BASE)")

        # 6. Personal de Ciclo
        if self._tiene_permiso("reemplazos.ver"):
            self.tab_ciclo = PersonalCicloView(
                self.records_service,
                self.pdf_service,
                self.excel_service,
                usuario_actual=self.usuario_actual,
                tiene_permiso_fn=self._tiene_permiso,
                parent=self,
            )
            self.tabs.addTab(self.tab_ciclo, "⏱️ Personal de Ciclo")

        # 7. Administración Integral por Pestañas
        tiene_admin = (
            self._tiene_permiso("usuarios.administrar")
            or self._tiene_permiso("roles.administrar")
            or self._tiene_permiso("tiposNovedad.administrar")
            or self._tiene_permiso("dotaciones.administrar")
            or self._tiene_permiso("backup.gestionar")
            or self._tiene_permiso("auditoria.ver")
            or self._tiene_permiso("registros.recuperar")
            or self._tiene_permiso("sesion.configurar")
            or self._tiene_permiso("destinatarios_informe.administrar")
            or self._tiene_permiso("destinatarios_ordenamiento.administrar")
        )
        if tiene_admin:
            self.tab_administracion = AdministracionView(
                self.auth_service,
                self.records_service,
                self.backup_service,
                self.excel_service,
                self.db,
                usuario_actual=self.usuario_actual,
                tiene_permiso_fn=self._tiene_permiso,
                parent=self,
            )
            self.tabs.addTab(self.tab_administracion, "⚙️ Administración")

        self.tabs.currentChanged.connect(self._on_tab_changed)

    def _on_tab_changed(self, index: int) -> None:
        widget = self.tabs.widget(index)
        if widget is getattr(self, "tab_dashboard", None):
            self.tab_dashboard.actualizar_tema()
        elif widget is getattr(self, "tab_buscador_turnos", None) and hasattr(self.tab_buscador_turnos, "actualizar_tema"):
            self.tab_buscador_turnos.actualizar_tema()
        elif widget is getattr(self, "tab_administracion", None) and hasattr(self.tab_administracion, "actualizar_tema"):
            self.tab_administracion.actualizar_tema()

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

    def _reconstruir_ui_sesion(self) -> None:
        """Reconstruye menús, pestañas, status bar y estilos según el usuario logueado."""
        self.menuBar().clear()
        self._init_menu_bar()
        self._init_tabs()
        self._init_status_bar()
        self._aplicar_estilos()
        self.recargar_todo()

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

        # Verificación no bloqueante de actualizaciones a los 5 segundos del inicio (Estándar CRONO-PCT)
        QTimer.singleShot(5000, self.verificar_actualizaciones_silenciosa)

    def eventFilter(self, watched, event) -> bool:
        if event is not None and event.type() in (QEvent.MouseMove, QEvent.MouseButtonPress, QEvent.KeyPress):
            self._ultimo_movimiento = time.time()
        return super().eventFilter(watched, event)

    def _verificar_inactividad(self) -> None:
        inactivo_seg = time.time() - self._ultimo_movimiento
        limite_seg = self._sesion_minutos * 60
        if inactivo_seg >= limite_seg:
            logger.info("Sesión expirada por inactividad del usuario.")
            self.timer_inactividad.stop()
            self._mostrar_login_por_inactividad()

    def _mostrar_login_por_inactividad(self) -> None:
        """Presenta el diálogo de inicio de sesión tras vencer el tiempo de inactividad."""
        from src.ui.views.dialogs.login_dialog import LoginDialog

        # Cerrar posibles modales abiertos por seguridad
        active_modal = QApplication.activeModalWidget()
        if active_modal and active_modal is not self:
            try:
                active_modal.reject()
            except Exception:
                pass

        QMessageBox.warning(
            self,
            "Sesión Caducada",
            "Su sesión ha caducado por inactividad.\nPor favor, vuelva a iniciar sesión para continuar trabajando.",
        )

        login_dlg = LoginDialog(self.auth_service, parent=self)
        if login_dlg.exec():
            nuevo_user = login_dlg.usuario_autenticado
            if nuevo_user:
                self.usuario_actual = nuevo_user
                logger.info("Sesión restablecida con éxito tras inactividad para '%s'", nuevo_user.get("username"))
                self._ultimo_movimiento = time.time()
                self._reconstruir_ui_sesion()
                self.timer_inactividad.start()
                return

        # Si el usuario cierra o cancela el diálogo de login, se procede al cierre
        logger.info("El usuario canceló el inicio de sesión tras caducar por inactividad. Saliendo.")
        self.close()

    @property
    def tab_novedades(self):
        """Compatibilidad para acceso directo a NovedadesView."""
        if hasattr(self, "tab_novedades_cambios"):
            return getattr(self.tab_novedades_cambios, "tab_novedades", None)
        return None

    @property
    def tab_cambios(self):
        """Compatibilidad para acceso directo a CambiosTurnoView."""
        if hasattr(self, "tab_novedades_cambios"):
            return getattr(self.tab_novedades_cambios, "tab_cambios", None)
        return None

    def recargar_todo(self) -> None:
        """Refresca todas las vistas activas y recomputa notificaciones."""
        try:
            self._actualizar_notificaciones_nuevos()
            self.tab_dashboard.actualizar_dashboard(no_revisados_count=self._sin_revisar_count)

            if hasattr(self, "tab_novedades_cambios"):
                self.tab_novedades_cambios.recargar_datos()
            if hasattr(self, "tab_turnos_sin_cubrir"):
                self.tab_turnos_sin_cubrir.consultar()
            if hasattr(self, "tab_empleados"):
                self.tab_empleados.recargar()
            if hasattr(self, "tab_ciclo"):
                self.tab_ciclo.consultar()
            if hasattr(self, "tab_administracion"):
                self.tab_administracion.recargar_todo()

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
                """SELECT id, legajo, novedad, apellidos_nombres, registrado_en, usuario_windows, usuario_id FROM novedades
                   WHERE activo=1 AND registrado_en >= ?
                     AND (usuario_id != ? OR (usuario_id IS NULL AND usuario_windows != ?))
                   ORDER BY id DESC LIMIT 20""",
                (ultimo_ingreso, u_id, win_u),
            ).fetchall()

            # Cambios cargados por otros
            cams = conn.execute(
                """SELECT id, legajo_1, legajo_2, fecha_cambio, registrado_en, usuario_windows, usuario_id FROM cambios_turno
                   WHERE activo=1 AND registrado_en >= ?
                     AND (usuario_id != ? OR (usuario_id IS NULL AND usuario_windows != ?))
                   ORDER BY id DESC LIMIT 20""",
                (ultimo_ingreso, u_id, win_u),
            ).fetchall()

        novs_list = [dict(n) for n in novs]
        cams_list = [dict(c) for c in cams]

        total = len(novs_list) + len(cams_list)
        self._sin_revisar_count = total
        self.menu_nuevos.setTitle(f"🔔 Registros Nuevos ({total})")

        ids_actuales = {("nov", n["id"]) for n in novs_list} | {("cam", c["id"]) for c in cams_list}

        # Toast si hay nuevos y las notificaciones están activadas
        if total > 0 and self.settings.get_bool("notificaciones_activo", True):
            if not hasattr(self, "_ultimos_ids_notificados") or ids_actuales != self._ultimos_ids_notificados:
                self._ultimos_ids_notificados = ids_actuales

                if total == 1:
                    if novs_list:
                        n = novs_list[0]
                        nid = n["id"]
                        nom = n.get("apellidos_nombres") or f"Leg. {n.get('legajo', '')}"
                        msg = f"Novedad #{nid}: {n.get('novedad')} ({nom})"
                        on_click = lambda: self._abrir_formulario_novedad_por_id(nid)
                    else:
                        c = cams_list[0]
                        cid = c["id"]
                        msg = f"Cambio de Turno #{cid}: Leg. {c.get('legajo_1')} ↔ {c.get('legajo_2')}"
                        on_click = lambda: self._abrir_formulario_cambio_por_id(cid)
                else:
                    msg = f"Hay {len(novs_list)} novedades y {len(cams_list)} cambios cargados desde su último acceso."
                    on_click = lambda: self._mostrar_nuevos_registros(novs_list, cams_list)

                duracion = self.settings.get_int("toast_duracion", 6)
                self.toast.show_message(
                    "Nuevos Registros Operativos",
                    msg,
                    duration_sec=duracion,
                    action_text="👉 Clic para abrir el formulario",
                    on_click=on_click,
                )

    def _on_abrir_menu_nuevos(self) -> None:
        self.menu_nuevos.clear()
        ultimo_ingreso = self.usuario_actual.get("ultimo_ingreso")
        u_id = self.usuario_actual.get("id")
        win_u = self.usuario_actual.get("username")

        with self.db.read_connection() as conn:
            novs = conn.execute(
                """SELECT id, legajo, novedad, apellidos_nombres, registrado_en, usuario_windows, usuario_id FROM novedades
                   WHERE activo=1 AND registrado_en >= ?
                     AND (usuario_id != ? OR (usuario_id IS NULL AND usuario_windows != ?))
                   ORDER BY id DESC LIMIT 20""",
                (ultimo_ingreso, u_id, win_u),
            ).fetchall()

            cams = conn.execute(
                """SELECT id, legajo_1, legajo_2, fecha_cambio, registrado_en, usuario_windows, usuario_id FROM cambios_turno
                   WHERE activo=1 AND registrado_en >= ?
                     AND (usuario_id != ? OR (usuario_id IS NULL AND usuario_windows != ?))
                   ORDER BY id DESC LIMIT 20""",
                (ultimo_ingreso, u_id, win_u),
            ).fetchall()

        novs_list = [dict(n) for n in novs]
        cams_list = [dict(c) for c in cams]
        total = len(novs_list) + len(cams_list)

        if total == 0:
            act_vacia = self.menu_nuevos.addAction("Sin nuevos registros pendientes")
            act_vacia.setEnabled(False)
            return

        # Acción destacada para ver el listado modal
        act_ver_todos = self.menu_nuevos.addAction(f"📋 Ver listado de registros nuevos ({total})...")
        act_ver_todos.triggered.connect(lambda: self._mostrar_nuevos_registros(novs_list, cams_list))
        self.menu_nuevos.addSeparator()

        for n in novs_list:
            nom = n.get("apellidos_nombres") or f"Leg. {n.get('legajo', '')}"
            txt = f"📝 #{n['id']} - {n.get('novedad')} ({nom}) [{n.get('registrado_en', '')}]"
            act = self.menu_nuevos.addAction(txt)
            act.setToolTip("Haga clic para abrir el formulario de esta novedad")
            act.triggered.connect(lambda _, nid=n["id"]: self._abrir_formulario_novedad_por_id(nid))

        if cams_list:
            if novs_list:
                self.menu_nuevos.addSeparator()
            for c in cams_list:
                fc = f" ({c.get('fecha_cambio')})" if c.get("fecha_cambio") else ""
                txt = f"🔄 #{c['id']} - Permuta Leg. {c.get('legajo_1')} ↔ {c.get('legajo_2')}{fc} [{c.get('registrado_en', '')}]"
                act = self.menu_nuevos.addAction(txt)
                act.setToolTip("Haga clic para abrir el formulario de este cambio de turno")
                act.triggered.connect(lambda _, cid=c["id"]: self._abrir_formulario_cambio_por_id(cid))

    def _abrir_formulario_novedad_por_id(self, novedad_id: int) -> None:
        """Abre directamente el formulario modal de la novedad para visualización o edición."""
        fila = self.records_service.obtener_novedad(novedad_id)
        if not fila:
            QMessageBox.warning(self, "Aviso", f"No se encontró la novedad #{novedad_id}.")
            return

        puede_modificar_libre = self._tiene_permiso("registros.modificar")
        dentro_ventana = self.records_service.dentro_de_ventana(fila["registrado_en"], "editar_horas")
        falta_fecha_fin = not fila.get("fecha_fin") or fila.get("fecha_fin") == "-"
        puede_editar = self._tiene_permiso("novedades.editar") or puede_modificar_libre

        solo_lectura = False
        modo_restringido = False

        if not puede_editar:
            solo_lectura = True
        elif not puede_modificar_libre and not dentro_ventana:
            if falta_fecha_fin:
                modo_restringido = True
            else:
                solo_lectura = True

        tramos = self.records_service.listar_reemplazos_por_novedad(novedad_id)
        from src.ui.views.dialogs.novedad_dialog import NovedadDialog
        dlg = NovedadDialog(
            self.records_service,
            novedad_existente=fila,
            tramos_existentes=tramos,
            modo_restringido=modo_restringido,
            solo_lectura=solo_lectura,
            usuario_actual=self.usuario_actual,
            parent=self,
        )
        if hasattr(self, "tab_novedades_cambios"):
            self.tabs.setCurrentWidget(self.tab_novedades_cambios)
            if hasattr(self.tab_novedades_cambios, "tabs"):
                self.tab_novedades_cambios.tabs.setCurrentIndex(0)

        if dlg.exec():
            self.recargar_todo()

    def _abrir_formulario_cambio_por_id(self, cambio_id: int) -> None:
        """Abre directamente el formulario modal del cambio de turno para visualización o edición."""
        fila = self.records_service.obtener_cambio(cambio_id)
        if not fila:
            QMessageBox.warning(self, "Aviso", f"No se encontró el cambio de turno #{cambio_id}.")
            return

        puede_modificar_libre = self._tiene_permiso("registros.modificar")
        dentro_ventana = self.records_service.dentro_de_ventana(fila["registrado_en"], "editar_horas")
        puede_editar = self._tiene_permiso("cambios_turno.editar") or puede_modificar_libre

        solo_lectura = not (puede_editar and (puede_modificar_libre or dentro_ventana))

        from src.ui.views.dialogs.cambio_turno_dialog import CambioTurnoDialog
        dlg = CambioTurnoDialog(
            self.records_service,
            cambio_existente=fila,
            usuario_actual=self.usuario_actual,
            solo_lectura=solo_lectura,
            parent=self,
        )
        if hasattr(self, "tab_novedades_cambios"):
            self.tabs.setCurrentWidget(self.tab_novedades_cambios)
            if hasattr(self.tab_novedades_cambios, "tabs"):
                self.tab_novedades_cambios.tabs.setCurrentIndex(1)

        if dlg.exec():
            self.recargar_todo()

    def _mostrar_nuevos_registros(
        self,
        novs: Optional[List[Dict[str, Any]]] = None,
        cams: Optional[List[Dict[str, Any]]] = None,
    ) -> None:
        """Muestra el diálogo modal con todos los registros nuevos y selector de formulario."""
        if novs is None or cams is None:
            ultimo_ingreso = self.usuario_actual.get("ultimo_ingreso")
            u_id = self.usuario_actual.get("id")
            win_u = self.usuario_actual.get("username")
            with self.db.read_connection() as conn:
                novs_rows = conn.execute(
                    """SELECT id, legajo, novedad, apellidos_nombres, registrado_en, usuario_windows, usuario_id FROM novedades
                       WHERE activo=1 AND registrado_en >= ?
                         AND (usuario_id != ? OR (usuario_id IS NULL AND usuario_windows != ?))
                       ORDER BY id DESC LIMIT 50""",
                    (ultimo_ingreso, u_id, win_u),
                ).fetchall()
                cams_rows = conn.execute(
                    """SELECT id, legajo_1, legajo_2, fecha_cambio, registrado_en, usuario_windows, usuario_id FROM cambios_turno
                       WHERE activo=1 AND registrado_en >= ?
                         AND (usuario_id != ? OR (usuario_id IS NULL AND usuario_windows != ?))
                       ORDER BY id DESC LIMIT 50""",
                    (ultimo_ingreso, u_id, win_u),
                ).fetchall()
            novs = [dict(n) for n in novs_rows]
            cams = [dict(c) for c in cams_rows]

        def _ir_a_grilla():
            if hasattr(self, "tab_novedades_cambios"):
                self.tabs.setCurrentWidget(self.tab_novedades_cambios)

        from src.ui.views.dialogs.nuevos_registros_dialog import NuevosRegistrosDialog
        dlg = NuevosRegistrosDialog(
            novedades=novs,
            cambios=cams,
            on_abrir_novedad=self._abrir_formulario_novedad_por_id,
            on_abrir_cambio=self._abrir_formulario_cambio_por_id,
            on_ir_a_grilla=_ir_a_grilla,
            parent=self,
        )
        dlg.exec()

    def _on_clic_kpi_sin_revisar(self) -> None:
        """Manejador para el clic en la tarjeta KPI 'SIN REVISAR' del dashboard."""
        self._mostrar_nuevos_registros()

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
        """Cierra la sesión activa y presenta el diálogo de inicio de sesión."""
        self.timer_inactividad.stop()
        from src.ui.views.dialogs.login_dialog import LoginDialog

        login_dlg = LoginDialog(self.auth_service, parent=self)
        if login_dlg.exec():
            nuevo_user = login_dlg.usuario_autenticado
            if nuevo_user:
                self.usuario_actual = nuevo_user
                logger.info("Nueva sesión iniciada para '%s'", nuevo_user.get("username"))
                self._ultimo_movimiento = time.time()
                self._reconstruir_ui_sesion()
                self.timer_inactividad.start()
                return
        self.close()

    def _acerca_de(self) -> None:
        dlg = AcercaDialog(self)
        dlg.btn_buscar_updates.clicked.connect(lambda: self.buscar_actualizaciones_manual(parent_dialog=dlg))
        dlg.exec()

    def _on_clic_btn_actualizar(self) -> None:
        """Acción al hacer clic en el botón de actualización en la cabecera."""
        if hasattr(self, "_ultima_info_actualizacion") and self._ultima_info_actualizacion:
            es_obligatoria = getattr(self, "_ultima_actualizacion_obligatoria", False)
            dlg = ActualizacionDialog(self._ultima_info_actualizacion, es_obligatoria=es_obligatoria, parent=self)
            dlg.exec()
        else:
            self.buscar_actualizaciones_manual()

    def verificar_actualizaciones_silenciosa(self) -> None:
        """Verificación asíncrona no bloqueante de actualizaciones al iniciar la aplicación (Estándar CRONO-PCT)."""
        self._worker_updater = VerificarActualizacionWorker(parent=self)
        self._worker_updater.actualizacion_disponible.connect(self._on_actualizacion_disponible_silenciosa)
        self._worker_updater.start()

    def _on_actualizacion_disponible_silenciosa(self, info: InfoActualizacion, es_obligatoria: bool) -> None:
        """Callback cuando el chequeo en segundo plano detecta una versión remota superior."""
        self._ultima_info_actualizacion = info
        self._ultima_actualizacion_obligatoria = es_obligatoria

        self.btn_actualizar.setText(f"🚀 ¡Actualizar a v{info.version}!")
        if es_obligatoria:
            self.btn_actualizar.setStyleSheet(
                "background-color: #D97706; color: #FFFFFF; font-weight: bold; border-radius: 4px; padding: 4px 10px; border: none;"
            )
            self.btn_actualizar.setToolTip(f"Actualización obligatoria a v{info.version} requerida")
        else:
            self.btn_actualizar.setStyleSheet(
                "background-color: #0284C7; color: #FFFFFF; font-weight: bold; border-radius: 4px; padding: 4px 10px; border: none;"
            )
            self.btn_actualizar.setToolTip(f"Nueva versión v{info.version} disponible. Clic para ver novedades.")

        self.btn_actualizar.setVisible(True)

        if es_obligatoria:
            # Si es obligatoria, abre el diálogo de actualización sin posponer
            QTimer.singleShot(600, lambda: ActualizacionDialog(info, es_obligatoria=True, parent=self).exec())
        else:
            # Aviso no invasivo toast
            try:
                toast = ToastWidget(
                    self,
                    f"🚀 Nueva versión v{info.version} disponible. Clic en el botón superior para actualizar.",
                    duracion_seg=8,
                )
                toast.mostrar()
            except Exception:
                pass

    def buscar_actualizaciones_manual(self, parent_dialog=None) -> None:
        """Verificación manual de actualizaciones invocada por el usuario con retroalimentación explícita."""
        self._worker_manual_updater = VerificarActualizacionWorker(parent=self)

        def _on_disponible(info: InfoActualizacion, es_obligatoria: bool):
            self._ultima_info_actualizacion = info
            self._ultima_actualizacion_obligatoria = es_obligatoria
            self.btn_actualizar.setText(f"🚀 ¡Actualizar a v{info.version}!")
            self.btn_actualizar.setVisible(True)

            if parent_dialog and hasattr(parent_dialog, "accept"):
                parent_dialog.accept()

            dlg = ActualizacionDialog(info, es_obligatoria=es_obligatoria, parent=self)
            dlg.exec()

        def _on_al_dia():
            QMessageBox.information(
                parent_dialog or self,
                "Sistema al Día",
                f"¡Tenés instalada la última versión disponible de RENO (v{APP_VERSION})!\n\nNo se requieren actualizaciones en este momento.",
            )

        def _on_error(err: str):
            QMessageBox.warning(
                parent_dialog or self,
                "Buscar Actualizaciones",
                f"No se pudo consultar el servidor de actualizaciones:\n\n{err}\n\nVerifique la conexión de red o la configuración de proxy.",
            )

        self._worker_manual_updater.actualizacion_disponible.connect(_on_disponible)
        self._worker_manual_updater.sistema_actualizado.connect(_on_al_dia)
        self._worker_manual_updater.error_verificacion.connect(_on_error)
        self._worker_manual_updater.start()


    def _cambiar_usuario(self) -> None:
        """Permite iniciar sesión con otro usuario sin reiniciar la aplicación."""
        from src.ui.views.dialogs.login_dialog import LoginDialog

        login_dlg = LoginDialog(self.auth_service, parent=self)
        if login_dlg.exec():
            nuevo_user = login_dlg.usuario_autenticado
            if nuevo_user:
                self.usuario_actual = nuevo_user
                logger.info("Cambio de usuario exitoso a '%s'", nuevo_user.get("username"))
                self._ultimo_movimiento = time.time()
                self._reconstruir_ui_sesion()
                QMessageBox.information(
                    self,
                    "Sesión Iniciada",
                    f"Ha iniciado sesión como <b>{nuevo_user.get('username')}</b>.",
                )

    def _mostrar_empleados(self) -> None:
        if hasattr(self, "tab_empleados"):
            self.tabs.setCurrentWidget(self.tab_empleados)
        else:
            from src.ui.views.dialogs.admin_dialogs import AdminEmpleadosDialog
            AdminEmpleadosDialog(self.records_service, self.excel_service, parent=self).exec()

    def _cambiar_mi_password(self) -> None:
        from src.ui.views.dialogs.cambiar_password_dialog import CambiarMiPasswordDialog
        dlg = CambiarMiPasswordDialog(self.auth_service, self.usuario_actual, parent=self)
        if dlg.exec():
            uid = self.usuario_actual.get("id")
            if uid:
                actualizado = self.auth_service.obtener_usuario(int(uid))
                if actualizado:
                    self.usuario_actual = actualizado
            self.sincronizar_ahora()

    def _abrir_buscador_tren(self) -> None:
        from src.ui.views.dialogs.buscador_tren_dialog import BuscadorTrenDialog
        dlg = BuscadorTrenDialog(self.records_service, self.excel_service, parent=self)
        dlg.exec()

    def closeEvent(self, event) -> None:
        self.settings.set_window_geometry("MainWindow", self.saveGeometry())
        super().closeEvent(event)

    def _restaurar_geometria(self) -> None:
        geom = self.settings.get_window_geometry("MainWindow")
        if geom:
            self.restoreGeometry(geom)
        else:
            self.showMaximized()
