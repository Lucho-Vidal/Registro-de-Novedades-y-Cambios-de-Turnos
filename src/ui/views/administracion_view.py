"""Vista integral de Administración con pestañas operativas y de control (CRONO-PCT Standard)."""

from typing import Any, Dict, Optional
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QHBoxLayout,
    QLabel,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from src.data.database import DatabaseManager
from src.services.auth_service import AuthService
from src.services.backup_service import BackupService
from src.services.excel_service import ExcelService
from src.services.records_service import RecordsService
from src.ui.views.dialogs.admin_dialogs import (
    AdminAuditoriaWidget,
    AdminBackupsWidget,
    AdminCatalogosWidget,
    AdminConfigWidget,
    AdminEliminadosWidget,
    AdminRolesWidget,
    AdminUsuariosWidget,
)


class AdministracionView(QWidget):
    """Pestaña principal de Administración con sub-pestañas:
    1. 👥 Usuarios y Roles
    2. 📋 Catálogos Operativos
    3. 💾 Seguridad y Auditoría
    4. ⚙️ Configuración del Sistema
    """

    def __init__(
        self,
        auth_service: AuthService,
        records_service: RecordsService,
        backup_service: BackupService,
        excel_service: ExcelService,
        db: DatabaseManager,
        usuario_actual: Optional[Dict[str, Any]] = None,
        tiene_permiso_fn=None,
        parent=None,
    ):
        super().__init__(parent)
        self.auth_service = auth_service
        self.records_service = records_service
        self.backup_service = backup_service
        self.excel_service = excel_service
        self.db = db
        self.usuario_actual = usuario_actual or {}
        self.tiene_permiso_fn = tiene_permiso_fn

        self._init_ui()

    def _can(self, permission_code: str) -> bool:
        if self.tiene_permiso_fn is None:
            return True
        return bool(self.tiene_permiso_fn(permission_code))

    def _init_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(6, 6, 6, 6)
        layout.setSpacing(6)

        self.tabs = QTabWidget()
        self.tabs.setDocumentMode(True)

        # 1. Sub-pestaña: Usuarios y Roles
        if self._can("usuarios.administrar") or self._can("roles.administrar"):
            tab_usr_roles = QWidget()
            lay_ur = QVBoxLayout(tab_usr_roles)
            lay_ur.setContentsMargins(4, 4, 4, 4)
            subtabs_ur = QTabWidget()
            subtabs_ur.setDocumentMode(True)

            if self._can("usuarios.administrar"):
                self.widget_usuarios = AdminUsuariosWidget(self.auth_service, parent=self)
                subtabs_ur.addTab(self.widget_usuarios, "👤 Gestión de Usuarios")

            if self._can("roles.administrar"):
                self.widget_roles = AdminRolesWidget(self.auth_service, parent=self)
                subtabs_ur.addTab(self.widget_roles, "🛡️ Roles y Permisos")

            lay_ur.addWidget(subtabs_ur)
            self.tabs.addTab(tab_usr_roles, "👥 Usuarios y Roles")

        # 2. Sub-pestaña: Catálogos Operativos
        if (
            self._can("tiposNovedad.administrar")
            or self._can("dotaciones.administrar")
            or self._can("personalEstacion.ver")
        ):
            self.widget_catalogos = AdminCatalogosWidget(
                self.records_service,
                self.excel_service,
                usuario_actual=self.usuario_actual,
                tiene_permiso_fn=self.tiene_permiso_fn,
                parent=self,
            )
            self.tabs.addTab(self.widget_catalogos, "📋 Catálogos Operativos")

        # 3. Sub-pestaña: Seguridad y Auditoría (Backups, Auditoría, Eliminados)
        if (
            self._can("backup.gestionar")
            or self._can("auditoria.ver")
            or self._can("registros.recuperar")
        ):
            tab_seg_aud = QWidget()
            lay_sa = QVBoxLayout(tab_seg_aud)
            lay_sa.setContentsMargins(4, 4, 4, 4)
            subtabs_sa = QTabWidget()
            subtabs_sa.setDocumentMode(True)

            if self._can("backup.gestionar"):
                self.widget_backups = AdminBackupsWidget(self.backup_service, parent=self)
                subtabs_sa.addTab(self.widget_backups, "💾 Copias de Seguridad (Backups)")

            if self._can("auditoria.ver"):
                self.widget_auditoria = AdminAuditoriaWidget(self.records_service, parent=self)
                subtabs_sa.addTab(self.widget_auditoria, "📜 Trazabilidad y Auditoría")

            if self._can("registros.recuperar"):
                self.widget_eliminados = AdminEliminadosWidget(
                    self.records_service,
                    usuario_actual=self.usuario_actual,
                    parent=self,
                )
                subtabs_sa.addTab(self.widget_eliminados, "♻️ Registros Eliminados (Papelera)")

            lay_sa.addWidget(subtabs_sa)
            self.tabs.addTab(tab_seg_aud, "💾 Seguridad y Auditoría")

        # 4. Sub-pestaña: Configuración del Sistema
        if self._can("sesion.configurar"):
            self.widget_config = AdminConfigWidget(
                self.db,
                records_service=self.records_service,
                usuario_actual=self.usuario_actual,
                tiene_permiso_fn=self.tiene_permiso_fn,
                parent=self,
            )
            self.tabs.addTab(self.widget_config, "⚙️ Configuración del Sistema")

        layout.addWidget(self.tabs)

    def recargar_todo(self) -> None:
        """Refresca todos los paneles visibles de administración."""
        if hasattr(self, "widget_usuarios"):
            self.widget_usuarios.recargar()
        if hasattr(self, "widget_roles"):
            self.widget_roles.recargar()
        if hasattr(self, "widget_catalogos"):
            self.widget_catalogos.recargar_todo()
        if hasattr(self, "widget_backups"):
            self.widget_backups.recargar()
        if hasattr(self, "widget_auditoria"):
            self.widget_auditoria.recargar()
        if hasattr(self, "widget_eliminados"):
            self.widget_eliminados.recargar()
        if hasattr(self, "widget_config"):
            self.widget_config.recargar()
