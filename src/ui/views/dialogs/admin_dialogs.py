"""Módulo de ventanas modales de Administración (abiertas maximizadas según reglas del sistema).

Incluye:
- AdminUsuariosDialog: Gestión de usuarios y reseteo de claves.
- AdminRolesDialog: Matriz de roles y permisos agrupados por módulo.
- AdminEmpleadosDialog: CRUD de empleados con soft-delete y filtros.
- AdminCatalogosDialog: Tipos de novedad, dotaciones y personal de estación.
- AdminConfigDialog: Tiempos de edición, caducidad de claves y notificaciones.
- AdminBackupsDialog: Copias de seguridad en caliente y restauración.
- AdminEliminadosDialog: Papelera de recuperación y borrado definitivo.
- AdminAuditoriaDialog: Registro histórico de trazabilidad de acciones.
"""

from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QFileDialog,
    QFormLayout,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QScrollArea,
    QSpinBox,
    QSplitter,
    QTabWidget,
    QTableView,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from src.core.logger import (
    abrir_carpeta_en_explorador,
    abrir_en_bloc_de_notas,
    get_logger,
    leer_ultimas_lineas_log,
    obtener_ruta_archivo_log,
)
from src.core.proxy import (
    AUTH_MANUAL_FIJO,
    MODO_DIRECTO,
    MODO_MANUAL,
    MODO_WINDOWS_SSO,
    OPCIONES_AUTH_PROXY,
    OPCIONES_MODO,
    aplicar_configuracion_proxy,
    detectar_proxy_sistema,
    probar_conexion_supabase,
)
from src.core.theme import TEMA_CLARO, TEMA_OSCURO, aplicar_icono_aplicacion, obtener_tema_visual
from src.domain.enums import (
    PERMISOS_BASE,
    PERMISOS_METADATA,
    ESTRUCTURA_PERMISOS_PESTANAS,
    DEFINICIONES_ACCIONES,
)
from src.services.auth_service import AuthService
from src.services.backup_service import BackupService
from src.services.excel_service import ExcelService
from src.services.records_service import RecordsService
from src.ui.models.base_table_model import BaseTableModel
from src.ui.models.multi_filter_proxy import MultiFilterProxyModel

logger = get_logger("AdminDialogs")


class BaseMaximizedDialog(QDialog):
    """Clase base para garantizar que todos los modales principales de administración se abran maximizados y respondan al cierre."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowFlags(Qt.Window | Qt.WindowTitleHint | Qt.WindowMinMaxButtonsHint | Qt.WindowCloseButtonHint)
        aplicar_icono_aplicacion(self)

    def closeEvent(self, event) -> None:
        self.reject()
        event.accept()

    def showEvent(self, event) -> None:
        super().showEvent(event)
        self.showMaximized()


# ---------------------------------------------------------------------------
# 1. Administración de Usuarios
# ---------------------------------------------------------------------------

class AdminUsuariosWidget(QWidget):
    """Panel de gestión integral de cuentas de usuario, roles y claves."""

    def __init__(self, auth_service: AuthService, parent=None):
        super().__init__(parent)
        self.auth_service = auth_service

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        # Barra de acciones
        bar = QHBoxLayout()
        self.btn_nuevo = QPushButton("➕ Nuevo Usuario")
        self.btn_nuevo.clicked.connect(self._nuevo_usuario)
        bar.addWidget(self.btn_nuevo)

        self.btn_editar = QPushButton("✏️ Modificar Usuario")
        self.btn_editar.setObjectName("btnSecondary")
        self.btn_editar.clicked.connect(self._modificar_usuario)
        bar.addWidget(self.btn_editar)

        self.btn_toggle_activo = QPushButton("🔄 Activar / Desactivar")
        self.btn_toggle_activo.setObjectName("btnSecondary")
        self.btn_toggle_activo.clicked.connect(self._toggle_activo)
        bar.addWidget(self.btn_toggle_activo)

        self.btn_cambiar_clave = QPushButton("🔑 Restablecer Clave")
        self.btn_cambiar_clave.setObjectName("btnSecondary")
        self.btn_cambiar_clave.clicked.connect(self._cambiar_clave)
        bar.addWidget(self.btn_cambiar_clave)

        bar.addStretch()
        layout.addLayout(bar)

        # Tabla
        headers = ["ID", "USUARIO", "NOMBRE", "LEGAJO", "ROLES ASIGNADOS", "ESTADO"]
        keys = ["id", "username", "nombre", "legajo", "roles", "activo_str"]

        self.model = BaseTableModel(headers, keys, parent=self)
        self.table = QTableView()
        self.table.setModel(self.model)
        self.table.setSelectionBehavior(QTableView.SelectRows)
        self.table.setSelectionMode(QTableView.SingleSelection)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table.setColumnWidth(0, 60)
        self.table.setColumnWidth(1, 140)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.table.setColumnWidth(3, 100)
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        self.table.setColumnWidth(5, 100)
        self.table.doubleClicked.connect(lambda idx: self._modificar_usuario())
        layout.addWidget(self.table)

        self.recargar()

    def recargar(self) -> None:
        sel = self._obtener_seleccionado()
        sel_id = sel.get("id") if sel else None

        usuarios = self.auth_service.listar_usuarios()
        for u in usuarios:
            u["activo_str"] = "Activo" if u.get("activo") == 1 else "Inactivo"
        self.model.set_data(usuarios)

        if sel_id is not None:
            for row_idx, u in enumerate(usuarios):
                if u.get("id") == sel_id:
                    self.table.selectRow(row_idx)
                    break

    def _obtener_seleccionado(self) -> Optional[Dict[str, Any]]:
        if self.table.selectionModel():
            indexes = self.table.selectionModel().selectedRows()
            if indexes:
                return self.model.get_row(indexes[0].row())
        curr = self.table.currentIndex()
        if curr.isValid():
            return self.model.get_row(curr.row())
        return None

    def _nuevo_usuario(self) -> None:
        roles_disponibles = [r["nombre"] for r in self.auth_service.listar_roles()]
        dlg = QDialog(self)
        dlg.setWindowTitle("Nuevo Usuario")
        dlg.setMinimumWidth(400)
        form = QFormLayout(dlg)

        txt_user = QLineEdit()
        txt_pass = QLineEdit()
        txt_pass.setEchoMode(QLineEdit.Password)
        txt_nom = QLineEdit()
        txt_leg = QLineEdit()

        combo_rol = QComboBox()
        combo_rol.addItems(roles_disponibles)

        form.addRow("Usuario:", txt_user)
        form.addRow("Contraseña:", txt_pass)
        form.addRow("Nombre:", txt_nom)
        form.addRow("Legajo:", txt_leg)
        form.addRow("Rol principal:", combo_rol)

        btn_box = QHBoxLayout()
        btn_box.addStretch()
        btn_cancel = QPushButton("Cancelar")
        btn_cancel.setObjectName("btnSecondary")
        btn_cancel.setAutoDefault(False)
        btn_cancel.setDefault(False)
        btn_cancel.clicked.connect(dlg.reject)
        btn_ok = QPushButton("Crear")
        btn_ok.setAutoDefault(True)
        btn_ok.setDefault(True)
        btn_ok.clicked.connect(dlg.accept)
        btn_box.addWidget(btn_cancel)
        btn_box.addWidget(btn_ok)
        form.addRow(btn_box)
        dlg.adjustSize()

        if dlg.exec():
            u = txt_user.text().strip()
            p = txt_pass.text()
            n = txt_nom.text().strip()
            l_str = txt_leg.text().strip()
            r = combo_rol.currentText()
            if not u or not p or not l_str:
                QMessageBox.warning(self, "Validación", "Usuario, contraseña y legajo son obligatorios.")
                return
            try:
                self.auth_service.crear_usuario(u, p, nombre=n, legajo=int(l_str), roles=(r,) if r else ())
                self.recargar()
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

    def _modificar_usuario(self) -> None:
        u = self._obtener_seleccionado()
        if not u:
            QMessageBox.warning(self, "Usuarios", "Seleccione un usuario en la tabla para modificar.")
            return

        roles_disponibles = [r["nombre"] for r in self.auth_service.listar_roles()]
        roles_actuales = [r.strip() for r in (u.get("roles") or "").split(",") if r.strip()]
        rol_principal = roles_actuales[0] if roles_actuales else ""

        dlg = QDialog(self)
        dlg.setWindowTitle(f"Modificar Usuario - {u['username']}")
        dlg.setMinimumWidth(420)
        form = QFormLayout(dlg)
        form.setSpacing(10)

        txt_user = QLineEdit(str(u.get("username") or ""))
        txt_nom = QLineEdit(str(u.get("nombre") or ""))
        txt_leg = QLineEdit(str(u.get("legajo") or ""))

        combo_rol = QComboBox()
        combo_rol.addItem("")
        for r_nom in roles_disponibles:
            combo_rol.addItem(r_nom)
        idx_rol = combo_rol.findText(rol_principal)
        if idx_rol >= 0:
            combo_rol.setCurrentIndex(idx_rol)

        chk_activo = QCheckBox("Usuario activo (puede iniciar sesión)")
        chk_activo.setChecked(bool(u.get("activo") == 1))

        form.addRow("Usuario:", txt_user)
        form.addRow("Nombre:", txt_nom)
        form.addRow("Legajo:", txt_leg)
        form.addRow("Rol principal:", combo_rol)
        form.addRow("", chk_activo)

        btn_box = QHBoxLayout()
        btn_box.addStretch()
        btn_cancel = QPushButton("Cancelar")
        btn_cancel.setObjectName("btnSecondary")
        btn_cancel.clicked.connect(dlg.reject)
        btn_ok = QPushButton("Guardar Cambios")
        btn_ok.clicked.connect(dlg.accept)
        btn_box.addWidget(btn_cancel)
        btn_box.addWidget(btn_ok)
        form.addRow(btn_box)

        if dlg.exec():
            usr = txt_user.text().strip()
            nom = txt_nom.text().strip()
            leg_str = txt_leg.text().strip()
            rol_sel = combo_rol.currentText().strip()

            if not usr or not leg_str:
                QMessageBox.warning(self, "Validación", "Usuario y legajo son obligatorios.")
                return

            try:
                leg_val = int(leg_str)
            except ValueError:
                QMessageBox.warning(self, "Validación", "El legajo debe ser un número entero.")
                return

            try:
                roles_nuevos = [rol_sel] if rol_sel else []
                self.auth_service.actualizar_usuario(
                    user_id=int(u["id"]),
                    username=usr,
                    nombre=nom,
                    legajo=leg_val,
                    activo=1 if chk_activo.isChecked() else 0,
                    roles=roles_nuevos,
                )
                self.recargar()
                QMessageBox.information(self, "Éxito", "Usuario modificado correctamente.")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"No se pudo modificar el usuario:\n{e}")

    def _toggle_activo(self) -> None:
        u = self._obtener_seleccionado()
        if not u:
            QMessageBox.warning(self, "Usuarios", "Seleccione un usuario en la tabla primero.")
            return
        self.auth_service.cambiar_estado_usuario(u["id"])
        self.recargar()

    def _cambiar_clave(self) -> None:
        u = self._obtener_seleccionado()
        if not u:
            QMessageBox.warning(self, "Usuarios", "Seleccione un usuario en la tabla para cambiar o restablecer su contraseña.")
            return

        dlg = QDialog(self)
        dlg.setWindowTitle(f"Cambiar / Restablecer Clave - {u['username']}")
        dlg.setMinimumWidth(400)
        dlg.setWindowFlags((dlg.windowFlags() & ~Qt.WindowContextHelpButtonHint) | Qt.WindowCloseButtonHint)

        lay = QVBoxLayout(dlg)
        lay.setContentsMargins(20, 16, 20, 16)
        lay.setSpacing(12)

        lbl_info = QLabel(f"Usuario seleccionado: <b>{u['username']}</b> ({u.get('nombre') or ''})")
        lay.addWidget(lbl_info)

        form = QFormLayout()
        form.setSpacing(10)

        txt_nueva = QLineEdit()
        txt_nueva.setEchoMode(QLineEdit.Password)
        txt_nueva.setPlaceholderText("Nueva contraseña")

        txt_conf = QLineEdit()
        txt_conf.setEchoMode(QLineEdit.Password)
        txt_conf.setPlaceholderText("Confirmar nueva contraseña")

        def _fila_pwd(le: QLineEdit) -> QWidget:
            cont = QWidget()
            h = QHBoxLayout(cont)
            h.setContentsMargins(0, 0, 0, 0)
            h.setSpacing(6)
            h.addWidget(le, stretch=1)
            b_ver = QPushButton("👁️")
            b_ver.setObjectName("btnSecondary")
            b_ver.setFixedWidth(36)
            b_ver.setAutoDefault(False)
            b_ver.setDefault(False)
            b_ver.setFocusPolicy(Qt.NoFocus)

            def _toggle() -> None:
                if le.echoMode() == QLineEdit.Password:
                    le.setEchoMode(QLineEdit.Normal)
                    b_ver.setText("🔒")
                else:
                    le.setEchoMode(QLineEdit.Password)
                    b_ver.setText("👁️")

            b_ver.clicked.connect(_toggle)
            h.addWidget(b_ver)
            return cont

        form.addRow("Nueva clave:", _fila_pwd(txt_nueva))
        form.addRow("Confirmar clave:", _fila_pwd(txt_conf))
        txt_nueva.returnPressed.connect(txt_conf.setFocus)
        lay.addLayout(form)

        btn_box = QHBoxLayout()
        btn_box.addStretch()
        btn_cancel = QPushButton("Cancelar")
        btn_cancel.setObjectName("btnSecondary")
        btn_cancel.setAutoDefault(False)
        btn_cancel.setDefault(False)
        btn_cancel.clicked.connect(dlg.reject)

        btn_save = QPushButton("Guardar Contraseña")
        btn_save.setAutoDefault(True)
        btn_save.setDefault(True)

        def _guardar_nueva() -> None:
            v1 = txt_nueva.text()
            v2 = txt_conf.text()
            if not v1:
                QMessageBox.warning(dlg, "Validación", "La contraseña no puede estar vacía.")
                return
            if v1 != v2:
                QMessageBox.warning(dlg, "Validación", "Las contraseñas no coinciden.")
                return
            try:
                self.auth_service.cambiar_password(int(u["id"]), v1)
                QMessageBox.information(dlg, "Éxito", f"Contraseña actualizada exitosamente para '{u['username']}'.")
                dlg.accept()
            except Exception as ex:
                QMessageBox.critical(dlg, "Error", f"No se pudo actualizar la contraseña:\n{ex}")

        btn_save.clicked.connect(_guardar_nueva)
        btn_box.addWidget(btn_cancel)
        btn_box.addWidget(btn_save)
        lay.addLayout(btn_box)
        dlg.adjustSize()

        dlg.exec()


class AdminUsuariosDialog(BaseMaximizedDialog):
    """Gestión integral de cuentas de usuario, roles y claves."""

    def __init__(self, auth_service: AuthService, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Administración de Usuarios")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)

        self.widget = AdminUsuariosWidget(auth_service, parent=self)
        layout.addWidget(self.widget)

        btn_cerrar = QPushButton("Cerrar")
        btn_cerrar.setObjectName("btnSecondary")
        btn_cerrar.clicked.connect(self.accept)
        layout.addWidget(btn_cerrar, alignment=Qt.AlignRight)

    def recargar(self) -> None:
        self.widget.recargar()


# ---------------------------------------------------------------------------
# 2. Roles y Permisos
# ---------------------------------------------------------------------------

class AdminRolesWidget(QWidget):
    """Panel de gestión de roles y asignación granular de permisos agrupados por módulo."""

    def __init__(self, auth_service: AuthService, parent=None):
        super().__init__(parent)
        self.auth_service = auth_service
        self._checkboxes_permisos: Dict[str, QCheckBox] = {}
        self._group_boxes: List[QGroupBox] = []
        self._subgrupo_items: List[Tuple[QLabel, str, str, QPushButton, QPushButton]] = []
        self.rol_seleccionado_id: Optional[int] = None

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        splitter = QSplitter(Qt.Horizontal)

        # Panel izquierdo: Lista de roles
        left_widget = QWidget()
        lay_left = QVBoxLayout(left_widget)
        lay_left.setContentsMargins(0, 0, 0, 0)

        lay_left.addWidget(QLabel("<b>Roles Registrados:</b>"))
        self.tabla_roles = QTableWidget(0, 2)
        self.tabla_roles.setHorizontalHeaderLabels(["ID", "NOMBRE DEL ROL"])
        self.tabla_roles.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.tabla_roles.setColumnWidth(0, 60)
        self.tabla_roles.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.tabla_roles.setSelectionBehavior(QTableWidget.SelectRows)
        self.tabla_roles.setSelectionMode(QTableWidget.SingleSelection)
        self.tabla_roles.itemSelectionChanged.connect(self._on_rol_seleccionado)
        lay_left.addWidget(self.tabla_roles)

        h_rol_btn = QHBoxLayout()
        self.txt_nuevo_rol = QLineEdit()
        self.txt_nuevo_rol.setPlaceholderText("Nombre del nuevo rol...")
        self.btn_crear_rol = QPushButton("➕ Crear Rol")
        self.btn_crear_rol.clicked.connect(self._crear_rol)
        h_rol_btn.addWidget(self.txt_nuevo_rol)
        h_rol_btn.addWidget(self.btn_crear_rol)
        lay_left.addLayout(h_rol_btn)

        splitter.addWidget(left_widget)

        # Panel derecho: Matriz de permisos
        right_widget = QWidget()
        lay_right = QVBoxLayout(right_widget)
        lay_right.setContentsMargins(4, 0, 0, 0)
        lay_right.setSpacing(6)

        self.lbl_rol_activo = QLabel("<b>Seleccione un rol para editar sus permisos</b>")
        self.lbl_rol_activo.setObjectName("lblRolActivo")
        lay_right.addWidget(self.lbl_rol_activo)

        # Banner explicativo de acciones RBAC
        self.card_guia = QFrame()
        self.card_guia.setObjectName("cardGuia")
        lay_guia = QVBoxLayout(self.card_guia)
        lay_guia.setContentsMargins(10, 8, 10, 8)
        lay_guia.setSpacing(4)
        self.lbl_guia_title = QLabel("<b>💡 Guía de Acciones y Niveles de Acceso (RBAC):</b>")
        self.lbl_guia_title.setObjectName("lblGuiaTitle")
        lay_guia.addWidget(self.lbl_guia_title)

        self.lbl_guia_desc = QLabel()
        self.lbl_guia_desc.setObjectName("lblGuiaDesc")
        self.lbl_guia_desc.setWordWrap(True)
        lay_guia.addWidget(self.lbl_guia_desc)
        lay_right.addWidget(self.card_guia)

        # Barra de selección rápida de permisos
        bar_quick = QHBoxLayout()
        self.btn_marcar_todos = QPushButton("☑️ Marcar Todos")
        self.btn_marcar_todos.setObjectName("btnSecondary")
        self.btn_marcar_todos.setEnabled(False)
        self.btn_marcar_todos.clicked.connect(lambda: self._marcar_todos(True))
        bar_quick.addWidget(self.btn_marcar_todos)

        self.btn_desmarcar_todos = QPushButton("⬜ Desmarcar Todos")
        self.btn_desmarcar_todos.setObjectName("btnSecondary")
        self.btn_desmarcar_todos.setEnabled(False)
        self.btn_desmarcar_todos.clicked.connect(lambda: self._marcar_todos(False))
        bar_quick.addWidget(self.btn_desmarcar_todos)

        self.btn_solo_lectura = QPushButton("👁️ Solo Lectura (Ver y Exportar)")
        self.btn_solo_lectura.setObjectName("btnSecondary")
        self.btn_solo_lectura.setEnabled(False)
        self.btn_solo_lectura.clicked.connect(self._marcar_solo_lectura)
        bar_quick.addWidget(self.btn_solo_lectura)

        bar_quick.addStretch()
        lay_right.addLayout(bar_quick)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        self.permisos_container = QWidget()
        self.permisos_layout = QVBoxLayout(self.permisos_container)
        self.permisos_layout.setContentsMargins(4, 4, 4, 4)
        self.permisos_layout.setSpacing(10)
        scroll.setWidget(self.permisos_container)
        lay_right.addWidget(scroll)

        self.btn_guardar_permisos = QPushButton("💾 Guardar Permisos del Rol")
        self.btn_guardar_permisos.setObjectName("btnSuccess")
        self.btn_guardar_permisos.setEnabled(False)
        self.btn_guardar_permisos.clicked.connect(self._guardar_permisos)
        lay_right.addWidget(self.btn_guardar_permisos)

        splitter.addWidget(right_widget)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 2)
        layout.addWidget(splitter)

        self._cargar_permisos_ui()
        self.actualizar_tema()
        self.recargar_roles()

    def recargar(self) -> None:
        self.actualizar_tema()
        self.recargar_roles()

    def showEvent(self, event) -> None:
        super().showEvent(event)
        self.actualizar_tema()

    def actualizar_tema(self, tema: str | None = None) -> None:
        """Adapta colores, contrastes y estilos de la guía, botones y casillas según el tema visual activo."""
        tema_activo = tema or obtener_tema_visual()
        es_oscuro = (tema_activo == TEMA_OSCURO)

        # 1. Tarjeta Guía RBAC
        if hasattr(self, "card_guia"):
            if es_oscuro:
                self.card_guia.setStyleSheet(
                    "QFrame#cardGuia { background-color: #242436; border: 1px solid #3E3E55; border-radius: 6px; }"
                )
            else:
                self.card_guia.setStyleSheet(
                    "QFrame#cardGuia { background-color: #F8FAFC; border: 1px solid #CBD5E1; border-radius: 6px; }"
                )

        if hasattr(self, "lbl_guia_title"):
            color_titulo = "#60A5FA" if es_oscuro else "#1E3A8A"
            self.lbl_guia_title.setStyleSheet(f"color: {color_titulo}; font-size: 9.5pt; font-weight: bold;")

        if hasattr(self, "lbl_guia_desc"):
            color_txt = "#E2E8F0" if es_oscuro else "#334155"
            color_bold = "#FFFFFF" if es_oscuro else "#0F172A"
            color_code = "#38BDF8" if es_oscuro else "#1E3A8A"
            color_admin = "#34D399" if es_oscuro else "#047857"
            color_super = "#F87171" if es_oscuro else "#B91C1C"

            txt = (
                f"<div style='font-size: 8.5pt; line-height: 1.45; color: {color_txt};'>"
                f"• <b style='color: {color_bold};'>👁️ Ver / Leer:</b> Acceso de solo lectura para consultar datos, diagramas y filtros sin modificar.<br>"
                f"• <b style='color: {color_bold};'>➕ Crear:</b> Habilita los botones de alta ('+ Nuevo') y formularios para ingresar nuevos registros.<br>"
                f"• <b style='color: {color_bold};'>✏️ Editar:</b> Permite modificar datos existentes dentro de las ventanas temporales autorizadas (24h/72h).<br>"
                f"• <b style='color: {color_bold};'>🗑️ Eliminar:</b> Permite anular o enviar registros a la papelera dentro del límite temporal permitido.<br>"
                f"• <b style='color: {color_bold};'>📊 Exportar / 📥 Importar:</b> Generación y descarga de archivos (Excel, PDF, Word) o carga masiva.<br>"
                f"• <b style='color: {color_admin};'>⚙️ Administrar:</b> <u>Control integral y estructural del módulo</u>. A diferencia de Ver/Crear/Editar, quien "
                f"administra gestiona configuraciones maestras, catálogos (dotaciones, tipos), asigna roles o reactiva registros.<br>"
                f"• <b style='color: {color_super};'>⚡ Supervisión Irrestricta (<code style='color: {color_code};'>registros.modificar</code>):</b> Permiso exclusivo que ignora los límites "
                f"de 24h / 72h para editar o anular registros históricos."
                f"</div>"
            )
            self.lbl_guia_desc.setText(txt)

        # 2. Etiqueta de Rol Activo
        if hasattr(self, "lbl_rol_activo"):
            color_rol = "#60A5FA" if es_oscuro else "#1F497D"
            self.lbl_rol_activo.setStyleSheet(f"color: {color_rol}; font-size: 11pt; font-weight: bold;")

        # 3. Botones de acción rápida con alto contraste
        style_quick = (
            "QPushButton { background-color: #2A2A3C; color: #F1F5F9; border: 1px solid #4B5563; border-radius: 4px; padding: 5px 12px; font-weight: 600; font-size: 8.5pt; }"
            "QPushButton:hover { background-color: #374151; border-color: #60A5FA; color: #60A5FA; }"
            "QPushButton:disabled { background-color: #1E1E2E; color: #64748B; border-color: #2E2E3E; }"
            if es_oscuro else
            "QPushButton { background-color: #FFFFFF; color: #0F172A; border: 1px solid #94A3B8; border-radius: 4px; padding: 5px 12px; font-weight: 600; font-size: 8.5pt; }"
            "QPushButton:hover { background-color: #F1F5F9; border-color: #1E3A8A; color: #1E3A8A; }"
            "QPushButton:disabled { background-color: #F8FAFC; color: #94A3B8; border-color: #E2E8F0; }"
        )
        for btn in (getattr(self, "btn_marcar_todos", None), getattr(self, "btn_desmarcar_todos", None), getattr(self, "btn_solo_lectura", None)):
            if btn:
                btn.setStyleSheet(style_quick)

        # 4. Agrupadores de pestañas
        style_grp = (
            "QGroupBox { font-weight: bold; font-size: 10pt; color: #60A5FA; border: 1px solid #3E3E55; border-radius: 6px; margin-top: 10px; padding-top: 12px; }"
            "QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 4px; background-color: #1E1E2E; }"
            if es_oscuro else
            "QGroupBox { font-weight: bold; font-size: 10pt; color: #1E3A8A; border: 1px solid #CBD5E1; border-radius: 6px; margin-top: 10px; padding-top: 12px; }"
            "QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 4px; background-color: #FFFFFF; }"
        )
        for grp in getattr(self, "_group_boxes", []):
            grp.setStyleSheet(style_grp)

        # 5. Subgrupos (título, descripción y botones de enlace)
        color_sub_title = "#F1F5F9" if es_oscuro else "#0F172A"
        color_sub_desc = "#94A3B8" if es_oscuro else "#64748B"
        style_link_on = (
            "font-size: 8pt; font-weight: bold; padding: 1px 6px; color: #60A5FA; background: transparent; border: 1px solid #3B82F6; border-radius: 3px;"
            if es_oscuro else
            "font-size: 8pt; font-weight: bold; padding: 1px 6px; color: #1D4ED8; background: transparent; border: 1px solid #93C5FD; border-radius: 3px;"
        )
        style_link_off = (
            "font-size: 8pt; padding: 1px 6px; color: #94A3B8; background: transparent; border: 1px solid #4B5563; border-radius: 3px;"
            if es_oscuro else
            "font-size: 8pt; padding: 1px 6px; color: #64748B; background: transparent; border: 1px solid #CBD5E1; border-radius: 3px;"
        )

        for item in getattr(self, "_subgrupo_items", []):
            lbl, subpestana, desc, b_on, b_off = item
            lbl.setText(
                f"<span style='color: {color_sub_title}; font-weight: bold; font-size: 9.5pt;'>{subpestana}</span>"
                f" &nbsp;<span style='color: {color_sub_desc}; font-size: 8pt;'>({desc})</span>"
            )
            b_on.setStyleSheet(style_link_on)
            b_off.setStyleSheet(style_link_off)

        # 6. Casillas de permisos
        for code, chk in getattr(self, "_checkboxes_permisos", {}).items():
            if code == "registros.modificar":
                color_chk = "#F87171" if es_oscuro else "#B91C1C"
                chk.setStyleSheet(f"QCheckBox {{ font-weight: bold; color: {color_chk}; font-size: 9pt; }}")
            elif "administrar" in code:
                color_chk = "#34D399" if es_oscuro else "#047857"
                chk.setStyleSheet(f"QCheckBox {{ font-weight: bold; color: {color_chk}; font-size: 9pt; }}")
            else:
                color_chk = "#F1F5F9" if es_oscuro else "#1E293B"
                chk.setStyleSheet(f"QCheckBox {{ color: {color_chk}; font-size: 9pt; }}")

    def _cargar_permisos_ui(self) -> None:
        self._group_boxes.clear()
        self._subgrupo_items.clear()

        for tab_info in ESTRUCTURA_PERMISOS_PESTANAS:
            grp_pestana = QGroupBox(tab_info["pestana"])
            self._group_boxes.append(grp_pestana)

            lay_tab = QVBoxLayout(grp_pestana)
            lay_tab.setContentsMargins(10, 10, 10, 10)
            lay_tab.setSpacing(10)

            for subgrupo in tab_info["subgrupos"]:
                sub_widget = QWidget()
                lay_sub = QVBoxLayout(sub_widget)
                lay_sub.setContentsMargins(0, 0, 0, 0)
                lay_sub.setSpacing(4)

                # Encabezado del subgrupo con descripción
                header_box = QHBoxLayout()
                lbl_sub = QLabel()
                header_box.addWidget(lbl_sub, 1)

                # Acciones rápidas del subgrupo
                perm_codes = list(subgrupo["permisos"])
                btn_grp_on = QPushButton("Marcar grupo")
                btn_grp_on.setObjectName("btnLink")
                btn_grp_on.clicked.connect(lambda _, codes=perm_codes: self._set_grupo_checked(codes, True))
                header_box.addWidget(btn_grp_on)

                btn_grp_off = QPushButton("Desmarcar")
                btn_grp_off.setObjectName("btnLink")
                btn_grp_off.clicked.connect(lambda _, codes=perm_codes: self._set_grupo_checked(codes, False))
                header_box.addWidget(btn_grp_off)

                self._subgrupo_items.append((lbl_sub, subgrupo["subpestana"], subgrupo["descripcion"], btn_grp_on, btn_grp_off))
                lay_sub.addLayout(header_box)

                # Grilla de checkboxes para este subgrupo
                grid = QGridLayout()
                grid.setContentsMargins(4, 2, 4, 6)
                grid.setHorizontalSpacing(14)
                grid.setVerticalSpacing(4)

                col = 0
                row = 0
                for code in perm_codes:
                    meta = PERMISOS_METADATA.get(code, {})
                    accion = meta.get("accion", "Acción")
                    titulo = meta.get("titulo", code)
                    chk_text = f"[{accion}] {titulo}  ({code})"
                    chk = QCheckBox(chk_text)

                    # Tooltip informativo y pedagógico
                    tooltip = (
                        f"<b>Permiso:</b> <code>{code}</code><br>"
                        f"<b>Nivel de Acción:</b> {accion}<br>"
                        f"<b>Pestaña:</b> {meta.get('pestana', '')}<br>"
                        f"<b>Subpestaña:</b> {meta.get('subpestana', '')}<br><br>"
                        f"<b>Descripción funcional:</b><br>{meta.get('descripcion', '')}"
                    )
                    chk.setToolTip(tooltip)

                    self._checkboxes_permisos[code] = chk
                    grid.addWidget(chk, row, col)
                    col += 1
                    if col > 1:
                        col = 0
                        row += 1

                lay_sub.addLayout(grid)
                lay_tab.addWidget(sub_widget)

            self.permisos_layout.addWidget(grp_pestana)

        self.permisos_layout.addStretch()

    def recargar_roles(self) -> None:
        roles = self.auth_service.listar_roles()
        self.tabla_roles.setRowCount(len(roles))
        for r, rol in enumerate(roles):
            self.tabla_roles.setItem(r, 0, QTableWidgetItem(str(rol["id"])))
            self.tabla_roles.setItem(r, 1, QTableWidgetItem(str(rol["nombre"])))

    def _on_rol_seleccionado(self) -> None:
        sel = self.tabla_roles.selectedItems()
        if not sel:
            self.btn_guardar_permisos.setEnabled(False)
            self.btn_marcar_todos.setEnabled(False)
            self.btn_desmarcar_todos.setEnabled(False)
            self.btn_solo_lectura.setEnabled(False)
            return

        row = sel[0].row()
        r_id = int(self.tabla_roles.item(row, 0).text())
        r_nom = self.tabla_roles.item(row, 1).text()

        self.rol_seleccionado_id = r_id
        self.lbl_rol_activo.setText(f"<b>Permisos asignados al rol: {r_nom}</b>")
        self.btn_guardar_permisos.setEnabled(True)
        self.btn_marcar_todos.setEnabled(True)
        self.btn_desmarcar_todos.setEnabled(True)
        self.btn_solo_lectura.setEnabled(True)

        permisos_actuales = set(self.auth_service.permisos_de_rol(r_id))
        for code, chk in self._checkboxes_permisos.items():
            chk.setChecked(code in permisos_actuales)

    def _marcar_todos(self, checked: bool) -> None:
        for chk in self._checkboxes_permisos.values():
            chk.setChecked(checked)

    def _marcar_solo_lectura(self) -> None:
        for code, chk in self._checkboxes_permisos.items():
            es_lectura = code.endswith(".ver") or code.endswith(".exportar")
            chk.setChecked(es_lectura)

    def _set_grupo_checked(self, codes: List[str], checked: bool) -> None:
        for code in codes:
            chk = self._checkboxes_permisos.get(code)
            if chk:
                chk.setChecked(checked)

    def _crear_rol(self) -> None:
        nom = self.txt_nuevo_rol.text().strip()
        if not nom:
            return
        try:
            self.auth_service.crear_rol(nom)
            self.txt_nuevo_rol.clear()
            self.recargar_roles()
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

    def _guardar_permisos(self) -> None:
        if not self.rol_seleccionado_id:
            return
        seleccionados = [code for code, chk in self._checkboxes_permisos.items() if chk.isChecked()]
        try:
            self.auth_service.establecer_permisos_rol(self.rol_seleccionado_id, seleccionados)
            QMessageBox.information(self, "Éxito", "Permisos del rol actualizados correctamente.")
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))


class AdminRolesDialog(BaseMaximizedDialog):
    """Gestión de roles y asignación granular de permisos agrupados por módulo."""

    def __init__(self, auth_service: AuthService, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Administración de Roles y Permisos")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)

        self.widget = AdminRolesWidget(auth_service, parent=self)
        layout.addWidget(self.widget)

        btn_cerrar = QPushButton("Cerrar")
        btn_cerrar.setObjectName("btnSecondary")
        btn_cerrar.clicked.connect(self.accept)
        layout.addWidget(btn_cerrar, alignment=Qt.AlignRight)

    def showEvent(self, event) -> None:
        super().showEvent(event)
        if hasattr(self, "widget") and hasattr(self.widget, "actualizar_tema"):
            self.widget.actualizar_tema()

    def recargar_roles(self) -> None:
        self.widget.recargar_roles()


# ---------------------------------------------------------------------------
# 3. Administración de Empleados (BASE)
# ---------------------------------------------------------------------------

class AdminEmpleadosDialog(BaseMaximizedDialog):
    """CRUD completo de empleados con filtros en memoria y soft-delete."""

    def __init__(self, records_service: RecordsService, excel_service: ExcelService, parent=None):
        super().__init__(parent)
        self.records_service = records_service
        self.excel_service = excel_service
        self.setWindowTitle("Administración de Empleados (BASE)")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)

        # Toolbar
        bar = QHBoxLayout()
        self.btn_nuevo = QPushButton("➕ Nuevo Empleado")
        self.btn_nuevo.clicked.connect(self._nuevo_empleado)
        bar.addWidget(self.btn_nuevo)

        self.btn_editar = QPushButton("✏️ Modificar")
        self.btn_editar.setObjectName("btnSecondary")
        self.btn_editar.clicked.connect(self._modificar_empleado)
        bar.addWidget(self.btn_editar)

        self.btn_toggle = QPushButton("🔄 Activar / Desactivar")
        self.btn_toggle.setObjectName("btnSecondary")
        self.btn_toggle.clicked.connect(self._toggle_estado)
        bar.addWidget(self.btn_toggle)

        sep = QFrame()
        sep.setFrameShape(QFrame.VLine)
        sep.setFrameShadow(QFrame.Sunken)
        bar.addWidget(sep)

        self.btn_importar = QPushButton("📥 Importar desde Excel")
        self.btn_importar.setObjectName("btnSecondary")
        self.btn_importar.clicked.connect(self._importar_excel)
        bar.addWidget(self.btn_importar)

        bar.addStretch()

        # Filtros
        bar.addWidget(QLabel("Buscar:"))
        self.txt_search = QLineEdit()
        self.txt_search.setPlaceholderText("Nombre o legajo...")
        self.txt_search.textChanged.connect(self._on_search)
        bar.addWidget(self.txt_search)

        layout.addLayout(bar)

        # Tabla
        headers = ["ID", "LEGAJO", "APELLIDOS Y NOMBRES", "ESPECIALIDAD", "DOTACIÓN", "TURNOS", "FRANCO", "HORARIO FRANCO", "JORNADA HS", "ESTADO"]
        keys = ["id", "legajo", "apellidos_nombres", "especialidad", "dotacion", "turnos", "franco", "franco_horario", "jornada_horas", "estado_str"]

        self.model = BaseTableModel(headers, keys, parent=self)
        self.proxy = MultiFilterProxyModel(self)
        self.proxy.setSourceModel(self.model)

        self.table = QTableView()
        self.table.setModel(self.proxy)
        self.table.setSelectionBehavior(QTableView.SelectRows)
        self.table.setSelectionMode(QTableView.SingleSelection)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        layout.addWidget(self.table)

        btn_cerrar = QPushButton("Cerrar")
        btn_cerrar.setObjectName("btnSecondary")
        btn_cerrar.clicked.connect(self.accept)
        layout.addWidget(btn_cerrar, alignment=Qt.AlignRight)

        self.recargar()

    def recargar(self) -> None:
        empleados = self.records_service.listar_empleados(incluir_inactivos=True)
        for e in empleados:
            e["estado_str"] = "Activo" if e.get("activo") == 1 else "Inactivo"
        self.model.set_data(empleados)

    def _on_search(self, text: str) -> None:
        self.proxy.set_global_search(text)

    def _obtener_seleccionado(self) -> Optional[Dict[str, Any]]:
        sel = self.table.selectionModel().selectedRows()
        if not sel:
            return None
        source_idx = self.proxy.mapToSource(sel[0])
        return self.model.get_row(source_idx.row())

    def _nuevo_empleado(self) -> None:
        self._formulario_empleado(None)

    def _modificar_empleado(self) -> None:
        emp = self._obtener_seleccionado()
        if not emp:
            QMessageBox.information(self, "Modificar", "Seleccione un empleado en la tabla.")
            return
        self._formulario_empleado(emp)

    def _formulario_empleado(self, emp_existente: Optional[Dict[str, Any]]) -> None:
        dlg = QDialog(self)
        dlg.setWindowTitle("Modificar Empleado" if emp_existente else "Nuevo Empleado")
        dlg.setFixedSize(450, 400)
        lay = QFormLayout(dlg)

        txt_leg = QLineEdit(str(emp_existente["legajo"]) if emp_existente else "")
        txt_nom = QLineEdit(str(emp_existente["apellidos_nombres"]) if emp_existente else "")
        txt_esp = QComboBox()
        txt_esp.addItems(["CONDUCTOR", "GUARDATREN", "OTRO"])
        if emp_existente and emp_existente.get("especialidad"):
            txt_esp.setEditText(emp_existente["especialidad"])

        combo_dot = QComboBox()
        dots = [d["nombre"] for d in self.records_service.listar_dotaciones(incluir_inactivos=False)]
        combo_dot.addItems(dots)
        if emp_existente and emp_existente.get("dotacion"):
            combo_dot.setEditText(emp_existente["dotacion"])

        txt_turno = QLineEdit(str(emp_existente["turnos"]) if emp_existente else "")
        txt_franco = QLineEdit(str(emp_existente["franco"]) if emp_existente else "")
        txt_franco_h = QLineEdit(str(emp_existente["franco_horario"]) if emp_existente else "")
        spin_jornada = QSpinBox()
        spin_jornada.setRange(6, 7)
        spin_jornada.setValue(int(emp_existente.get("jornada_horas", 6)) if emp_existente else 6)

        lay.addRow("Legajo:", txt_leg)
        lay.addRow("Apellidos y Nombres:", txt_nom)
        lay.addRow("Especialidad:", txt_esp)
        lay.addRow("Dotación:", combo_dot)
        lay.addRow("Turnos:", txt_turno)
        lay.addRow("Franco:", txt_franco)
        lay.addRow("Horario Franco:", txt_franco_h)
        lay.addRow("Jornada CCT (hs):", spin_jornada)

        btn = QPushButton("Guardar")
        btn.clicked.connect(dlg.accept)
        lay.addRow(btn)

        if dlg.exec():
            data = {
                "legajo": txt_leg.text().strip(),
                "apellidos_nombres": txt_nom.text().strip(),
                "especialidad": txt_esp.currentText().strip(),
                "dotacion": combo_dot.currentText().strip(),
                "turnos": txt_turno.text().strip(),
                "franco": txt_franco.text().strip(),
                "franco_horario": txt_franco_h.text().strip(),
                "jornada_horas": spin_jornada.value(),
            }
            try:
                if emp_existente:
                    self.records_service.actualizar_empleado(emp_existente["id"], data)
                else:
                    self.records_service.crear_empleado(data)
                self.recargar()
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

    def _toggle_estado(self) -> None:
        emp = self._obtener_seleccionado()
        if not emp:
            return
        self.records_service.cambiar_estado_empleado(emp["id"])
        self.recargar()

    def _importar_excel(self) -> None:
        f_path, _ = QFileDialog.getOpenFileName(self, "Seleccionar Libro BASE Excel", "", "Excel (*.xlsx)")
        if not f_path:
            return
        try:
            ins, act = self.excel_service.import_base_empleados(f_path)
            QMessageBox.information(self, "Importación Exitosa", f"Empleados insertados: {ins}\nEmpleados actualizados: {act}")
            self.recargar()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo importar el archivo:\n{e}")


# ---------------------------------------------------------------------------
# 4. Catálogos Operativos (Tipos de Novedad, Dotaciones, Estación)
# ---------------------------------------------------------------------------

class AdminCatalogosWidget(QWidget):
    """Panel de gestión de Tipos de Novedad, Dotaciones y Personal de Estación."""

    def __init__(
        self,
        records_service: RecordsService,
        excel_service: ExcelService,
        usuario_actual: Optional[Dict[str, Any]] = None,
        tiene_permiso_fn=None,
        parent=None,
    ):
        super().__init__(parent)
        self.records_service = records_service
        self.excel_service = excel_service
        self.usuario_actual = usuario_actual or {}
        self.tiene_permiso_fn = tiene_permiso_fn

        self._tipos: List[Dict[str, Any]] = []
        self._dots: List[Dict[str, Any]] = []
        self._pes: List[Dict[str, Any]] = []

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        tabs = QTabWidget()

        # Tab 1: Tipos de novedad
        tab_tipos = QWidget()
        lay_t = QVBoxLayout(tab_tipos)
        h_t = QHBoxLayout()
        self.btn_nt = QPushButton("➕ Nuevo Tipo")
        self.btn_nt.clicked.connect(self._nuevo_tipo)
        h_t.addWidget(self.btn_nt)

        self.btn_et = QPushButton("✏️ Modificar")
        self.btn_et.setObjectName("btnSecondary")
        self.btn_et.clicked.connect(self._modificar_tipo)
        h_t.addWidget(self.btn_et)

        self.btn_tt = QPushButton("🔄 Activar / Desactivar")
        self.btn_tt.setObjectName("btnSecondary")
        self.btn_tt.clicked.connect(self._toggle_tipo)
        h_t.addWidget(self.btn_tt)

        h_t.addStretch()
        lay_t.addLayout(h_t)

        self.table_tipos = QTableWidget(0, 5)
        self.table_tipos.setHorizontalHeaderLabels(["ID", "NOMBRE", "DÍAS AUTORIZADOS", "EXPORTA DESCUBIERTO", "ESTADO"])
        self.table_tipos.setSelectionBehavior(QTableView.SelectRows)
        self.table_tipos.setSelectionMode(QTableView.SingleSelection)
        self.table_tipos.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table_tipos.setMaximumWidth(820)
        self.table_tipos.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table_tipos.setColumnWidth(0, 60)
        self.table_tipos.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table_tipos.setColumnWidth(2, 140)
        self.table_tipos.setColumnWidth(3, 160)
        self.table_tipos.setColumnWidth(4, 100)
        self.table_tipos.doubleClicked.connect(lambda _: self._modificar_tipo())
        lay_t.addWidget(self.table_tipos)
        tabs.addTab(tab_tipos, "Tipos de Novedad")

        # Tab 2: Dotaciones
        tab_dots = QWidget()
        lay_d = QVBoxLayout(tab_dots)
        h_d = QHBoxLayout()
        self.btn_nd = QPushButton("➕ Nueva Dotación")
        self.btn_nd.clicked.connect(self._nueva_dotacion)
        h_d.addWidget(self.btn_nd)

        self.btn_ed = QPushButton("✏️ Modificar")
        self.btn_ed.setObjectName("btnSecondary")
        self.btn_ed.clicked.connect(self._modificar_dotacion)
        h_d.addWidget(self.btn_ed)

        self.btn_td = QPushButton("🔄 Activar / Desactivar")
        self.btn_td.setObjectName("btnSecondary")
        self.btn_td.clicked.connect(self._toggle_dotacion)
        h_d.addWidget(self.btn_td)

        h_d.addStretch()
        lay_d.addLayout(h_d)

        self.table_dots = QTableWidget(0, 3)
        self.table_dots.setHorizontalHeaderLabels(["ID", "DOTACIÓN", "ESTADO"])
        self.table_dots.setSelectionBehavior(QTableView.SelectRows)
        self.table_dots.setSelectionMode(QTableView.SingleSelection)
        self.table_dots.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table_dots.setMaximumWidth(550)
        self.table_dots.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table_dots.setColumnWidth(0, 60)
        self.table_dots.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table_dots.setColumnWidth(2, 100)
        self.table_dots.doubleClicked.connect(lambda _: self._modificar_dotacion())
        lay_d.addWidget(self.table_dots)
        tabs.addTab(tab_dots, "Dotaciones")

        # Tab 3: Personal de Estación
        tab_pe = QWidget()
        lay_pe = QVBoxLayout(tab_pe)
        h_pe = QHBoxLayout()
        self.btn_npe = QPushButton("➕ Nuevo Personal")
        self.btn_npe.clicked.connect(self._nuevo_personal_estacion)
        h_pe.addWidget(self.btn_npe)

        self.btn_epe = QPushButton("✏️ Modificar")
        self.btn_epe.setObjectName("btnSecondary")
        self.btn_epe.clicked.connect(self._modificar_personal_estacion)
        h_pe.addWidget(self.btn_epe)

        self.btn_tpe = QPushButton("🔄 Activar / Desactivar")
        self.btn_tpe.setObjectName("btnSecondary")
        self.btn_tpe.clicked.connect(self._toggle_personal_estacion)
        h_pe.addWidget(self.btn_tpe)

        self.btn_imp_pe = QPushButton("📥 Importar Excel")
        self.btn_imp_pe.setObjectName("btnSecondary")
        self.btn_imp_pe.clicked.connect(self._importar_personal_estacion)
        h_pe.addWidget(self.btn_imp_pe)

        h_pe.addStretch()
        lay_pe.addLayout(h_pe)

        self.table_pe = QTableWidget(0, 4)
        self.table_pe.setHorizontalHeaderLabels(["ID", "NOMBRE Y APELLIDO", "DOTACIÓN", "ESTADO"])
        self.table_pe.setSelectionBehavior(QTableView.SelectRows)
        self.table_pe.setSelectionMode(QTableView.SingleSelection)
        self.table_pe.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table_pe.setMaximumWidth(820)
        self.table_pe.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table_pe.setColumnWidth(0, 60)
        self.table_pe.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table_pe.setColumnWidth(2, 180)
        self.table_pe.setColumnWidth(3, 100)
        self.table_pe.doubleClicked.connect(lambda _: self._modificar_personal_estacion())
        lay_pe.addWidget(self.table_pe)
        tabs.addTab(tab_pe, "Personal de Estación")

        # Tab 4: Turnos de Itinerario
        from src.ui.views.dialogs.admin_turnos_dialog import AdminTurnosWidget
        self.widget_turnos = AdminTurnosWidget(
            self.records_service,
            usuario_actual=self.usuario_actual,
            tiene_permiso_fn=self.tiene_permiso_fn,
            parent=self,
        )
        tabs.addTab(self.widget_turnos, "Turnos de Itinerario")

        layout.addWidget(tabs)
        self.recargar_todo()

    def _can(self, permission_code: str) -> bool:
        if self.tiene_permiso_fn is None:
            return True
        return bool(self.tiene_permiso_fn(permission_code))

    def _u_info(self):
        u_id = self.usuario_actual.get("id")
        win_u = self.usuario_actual.get("username")
        return u_id, win_u

    def recargar_todo(self) -> None:
        # Tipos
        self._tipos = self.records_service.listar_tipos(incluir_inactivos=True)
        self.table_tipos.setRowCount(len(self._tipos))
        for r, t in enumerate(self._tipos):
            self.table_tipos.setItem(r, 0, QTableWidgetItem(str(t["id"])))
            self.table_tipos.setItem(r, 1, QTableWidgetItem(str(t["nombre"])))
            self.table_tipos.setItem(r, 2, QTableWidgetItem(str(t.get("dias", 0))))
            self.table_tipos.setItem(r, 3, QTableWidgetItem("Sí" if t.get("exportar_descubierto") else "No"))
            self.table_tipos.setItem(r, 4, QTableWidgetItem("Activo" if t.get("activo") else "Inactivo"))

        # Dotaciones
        self._dots = self.records_service.listar_dotaciones(incluir_inactivos=True)
        self.table_dots.setRowCount(len(self._dots))
        for r, d in enumerate(self._dots):
            self.table_dots.setItem(r, 0, QTableWidgetItem(str(d["id"])))
            self.table_dots.setItem(r, 1, QTableWidgetItem(str(d["nombre"])))
            self.table_dots.setItem(r, 2, QTableWidgetItem("Activo" if d.get("activo") else "Inactivo"))

        # Personal estación
        self._pes = self.records_service.listar_personal_estacion(incluir_inactivos=True)
        self.table_pe.setRowCount(len(self._pes))
        for r, p in enumerate(self._pes):
            self.table_pe.setItem(r, 0, QTableWidgetItem(str(p["id"])))
            self.table_pe.setItem(r, 1, QTableWidgetItem(str(p["nombre"])))
            self.table_pe.setItem(r, 2, QTableWidgetItem(str(p.get("dotacion") or "-")))
            self.table_pe.setItem(r, 3, QTableWidgetItem("Activo" if p.get("activo") else "Inactivo"))

        # Turnos
        if hasattr(self, "widget_turnos"):
            self.widget_turnos.recargar()

    # --- 1. Tipos de Novedad CRUD ---

    def _nuevo_tipo(self) -> None:
        if not self._can("tiposNovedad.administrar"):
            QMessageBox.warning(self, "Permiso Denegado", "No posee permiso para administrar tipos de novedad.")
            return
        dlg = QDialog(self)
        dlg.setWindowTitle("Nuevo Tipo de Novedad")
        dlg.setFixedSize(380, 210)
        lay = QFormLayout(dlg)
        txt = QLineEdit()
        spin = QSpinBox()
        spin.setRange(0, 365)
        chk = QCheckBox("Exportar en reporte de turnos sin cubrir")
        chk.setChecked(True)
        lay.addRow("Nombre:", txt)
        lay.addRow("Días duración:", spin)
        lay.addRow(chk)
        btn = QPushButton("Guardar")
        btn.clicked.connect(dlg.accept)
        lay.addRow(btn)

        if dlg.exec():
            n = txt.text().strip()
            if n:
                u_id, win_u = self._u_info()
                try:
                    self.records_service.crear_tipo(
                        n,
                        dias=spin.value(),
                        exportar_descubierto=1 if chk.isChecked() else 0,
                        user_id=u_id,
                        windows_user=win_u,
                    )
                    self.recargar_todo()
                except Exception as e:
                    QMessageBox.critical(self, "Error", f"No se pudo crear el tipo:\n{e}")

    def _modificar_tipo(self) -> None:
        if not self._can("tiposNovedad.administrar"):
            QMessageBox.warning(self, "Permiso Denegado", "No posee permiso para modificar tipos de novedad.")
            return
        row = self.table_tipos.currentRow()
        if row < 0 or row >= len(self._tipos):
            QMessageBox.information(self, "Selección", "Seleccione un tipo de novedad de la tabla.")
            return
        t = self._tipos[row]

        dlg = QDialog(self)
        dlg.setWindowTitle("Modificar Tipo de Novedad")
        dlg.setFixedSize(380, 240)
        lay = QFormLayout(dlg)
        txt = QLineEdit(str(t["nombre"]))
        spin = QSpinBox()
        spin.setRange(0, 365)
        spin.setValue(int(t.get("dias") or 0))
        chk_exp = QCheckBox("Exportar en reporte de turnos sin cubrir")
        chk_exp.setChecked(bool(t.get("exportar_descubierto")))
        chk_act = QCheckBox("Activo")
        chk_act.setChecked(bool(t.get("activo")))

        lay.addRow("Nombre:", txt)
        lay.addRow("Días duración:", spin)
        lay.addRow(chk_exp)
        lay.addRow(chk_act)
        btn = QPushButton("Guardar Cambios")
        btn.clicked.connect(dlg.accept)
        lay.addRow(btn)

        if dlg.exec():
            n = txt.text().strip()
            if n:
                u_id, win_u = self._u_info()
                try:
                    self.records_service.actualizar_tipo(
                        t["id"],
                        nombre=n,
                        dias=spin.value(),
                        exportar_descubierto=1 if chk_exp.isChecked() else 0,
                        activo=1 if chk_act.isChecked() else 0,
                        user_id=u_id,
                        windows_user=win_u,
                    )
                    self.recargar_todo()
                except Exception as e:
                    QMessageBox.critical(self, "Error", f"No se pudo actualizar el tipo:\n{e}")

    def _toggle_tipo(self) -> None:
        if not self._can("tiposNovedad.administrar"):
            QMessageBox.warning(self, "Permiso Denegado", "No posee permiso para modificar tipos de novedad.")
            return
        row = self.table_tipos.currentRow()
        if row < 0 or row >= len(self._tipos):
            QMessageBox.information(self, "Selección", "Seleccione un tipo de novedad de la tabla.")
            return
        t = self._tipos[row]
        nuevo_act = 0 if t.get("activo") else 1
        u_id, win_u = self._u_info()
        try:
            self.records_service.actualizar_tipo(
                t["id"],
                nombre=t["nombre"],
                dias=int(t.get("dias") or 0),
                exportar_descubierto=int(t.get("exportar_descubierto") or 1),
                activo=nuevo_act,
                user_id=u_id,
                windows_user=win_u,
            )
            self.recargar_todo()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo cambiar el estado:\n{e}")

    # --- 2. Dotaciones CRUD ---

    def _nueva_dotacion(self) -> None:
        if not self._can("dotaciones.administrar"):
            QMessageBox.warning(self, "Permiso Denegado", "No posee permiso para administrar dotaciones.")
            return
        dlg = QDialog(self)
        dlg.setWindowTitle("Nueva Dotación")
        dlg.setFixedSize(340, 140)
        lay = QFormLayout(dlg)
        txt = QLineEdit()
        lay.addRow("Nombre:", txt)
        btn = QPushButton("Guardar")
        btn.clicked.connect(dlg.accept)
        lay.addRow(btn)

        if dlg.exec():
            n = txt.text().strip()
            if n:
                u_id, win_u = self._u_info()
                try:
                    self.records_service.crear_dotacion(n, user_id=u_id, windows_user=win_u)
                    self.recargar_todo()
                except Exception as e:
                    QMessageBox.critical(self, "Error", f"No se pudo crear la dotación:\n{e}")

    def _modificar_dotacion(self) -> None:
        if not self._can("dotaciones.administrar"):
            QMessageBox.warning(self, "Permiso Denegado", "No posee permiso para modificar dotaciones.")
            return
        row = self.table_dots.currentRow()
        if row < 0 or row >= len(self._dots):
            QMessageBox.information(self, "Selección", "Seleccione una dotación de la tabla.")
            return
        d = self._dots[row]

        dlg = QDialog(self)
        dlg.setWindowTitle("Modificar Dotación")
        dlg.setFixedSize(340, 170)
        lay = QFormLayout(dlg)
        txt = QLineEdit(str(d["nombre"]))
        chk_act = QCheckBox("Activo")
        chk_act.setChecked(bool(d.get("activo")))
        lay.addRow("Nombre:", txt)
        lay.addRow(chk_act)
        btn = QPushButton("Guardar Cambios")
        btn.clicked.connect(dlg.accept)
        lay.addRow(btn)

        if dlg.exec():
            n = txt.text().strip()
            if n:
                u_id, win_u = self._u_info()
                try:
                    self.records_service.actualizar_dotacion(
                        d["id"],
                        nombre=n,
                        activo=1 if chk_act.isChecked() else 0,
                        user_id=u_id,
                        windows_user=win_u,
                    )
                    self.recargar_todo()
                except Exception as e:
                    QMessageBox.critical(self, "Error", f"No se pudo actualizar la dotación:\n{e}")

    def _toggle_dotacion(self) -> None:
        if not self._can("dotaciones.administrar"):
            QMessageBox.warning(self, "Permiso Denegado", "No posee permiso para modificar dotaciones.")
            return
        row = self.table_dots.currentRow()
        if row < 0 or row >= len(self._dots):
            QMessageBox.information(self, "Selección", "Seleccione una dotación de la tabla.")
            return
        d = self._dots[row]
        nuevo_act = 0 if d.get("activo") else 1
        u_id, win_u = self._u_info()
        try:
            self.records_service.actualizar_dotacion(
                d["id"],
                nombre=d["nombre"],
                activo=nuevo_act,
                user_id=u_id,
                windows_user=win_u,
            )
            self.recargar_todo()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo cambiar el estado:\n{e}")

    # --- 3. Personal de Estación CRUD ---

    def _nuevo_personal_estacion(self) -> None:
        if not self._can("personalEstacion.crear"):
            QMessageBox.warning(self, "Permiso Denegado", "No posee permiso para crear personal de estación.")
            return
        dots = [d["nombre"] for d in self.records_service.listar_dotaciones(incluir_inactivos=False)]
        dlg = QDialog(self)
        dlg.setWindowTitle("Nuevo Personal de Estación")
        dlg.setFixedSize(380, 170)
        lay = QFormLayout(dlg)
        txt = QLineEdit()
        combo = QComboBox()
        combo.addItem("")
        combo.addItems(dots)
        lay.addRow("Nombre:", txt)
        lay.addRow("Dotación:", combo)
        btn = QPushButton("Guardar")
        btn.clicked.connect(dlg.accept)
        lay.addRow(btn)

        if dlg.exec():
            n = txt.text().strip()
            dot = combo.currentText().strip() or None
            if n:
                u_id, win_u = self._u_info()
                try:
                    self.records_service.crear_personal_estacion(
                        n,
                        dotacion=dot,
                        user_id=u_id,
                        windows_user=win_u,
                    )
                    self.recargar_todo()
                except Exception as e:
                    QMessageBox.critical(self, "Error", f"No se pudo crear el personal de estación:\n{e}")

    def _modificar_personal_estacion(self) -> None:
        if not self._can("personalEstacion.editar"):
            QMessageBox.warning(self, "Permiso Denegado", "No posee permiso para modificar personal de estación.")
            return
        row = self.table_pe.currentRow()
        if row < 0 or row >= len(self._pes):
            QMessageBox.information(self, "Selección", "Seleccione un personal de estación de la tabla.")
            return
        p = self._pes[row]

        dots = [d["nombre"] for d in self.records_service.listar_dotaciones(incluir_inactivos=False)]
        dlg = QDialog(self)
        dlg.setWindowTitle("Modificar Personal de Estación")
        dlg.setFixedSize(380, 200)
        lay = QFormLayout(dlg)
        txt = QLineEdit(str(p["nombre"]))
        combo = QComboBox()
        combo.addItem("")
        combo.addItems(dots)
        if p.get("dotacion"):
            idx = combo.findText(p["dotacion"])
            if idx >= 0:
                combo.setCurrentIndex(idx)
            else:
                combo.addItem(p["dotacion"])
                combo.setCurrentText(p["dotacion"])

        chk_act = QCheckBox("Activo")
        chk_act.setChecked(bool(p.get("activo")))

        lay.addRow("Nombre:", txt)
        lay.addRow("Dotación:", combo)
        lay.addRow(chk_act)
        btn = QPushButton("Guardar Cambios")
        btn.clicked.connect(dlg.accept)
        lay.addRow(btn)

        if dlg.exec():
            n = txt.text().strip()
            dot = combo.currentText().strip() or None
            if n:
                u_id, win_u = self._u_info()
                try:
                    self.records_service.actualizar_personal_estacion(
                        p["id"],
                        nombre=n,
                        activo=1 if chk_act.isChecked() else 0,
                        dotacion=dot,
                        cambiar_dotacion=True,
                        user_id=u_id,
                        windows_user=win_u,
                    )
                    self.recargar_todo()
                except Exception as e:
                    QMessageBox.critical(self, "Error", f"No se pudo modificar el personal de estación:\n{e}")

    def _toggle_personal_estacion(self) -> None:
        if not self._can("personalEstacion.editar"):
            QMessageBox.warning(self, "Permiso Denegado", "No posee permiso para modificar personal de estación.")
            return
        row = self.table_pe.currentRow()
        if row < 0 or row >= len(self._pes):
            QMessageBox.information(self, "Selección", "Seleccione un personal de estación de la tabla.")
            return
        p = self._pes[row]
        nuevo_act = 0 if p.get("activo") else 1
        u_id, win_u = self._u_info()
        try:
            self.records_service.actualizar_personal_estacion(
                p["id"],
                nombre=p["nombre"],
                activo=nuevo_act,
                dotacion=p.get("dotacion"),
                cambiar_dotacion=False,
                user_id=u_id,
                windows_user=win_u,
            )
            self.recargar_todo()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo cambiar el estado:\n{e}")

    def _importar_personal_estacion(self) -> None:
        if not self._can("personalEstacion.importar"):
            QMessageBox.warning(self, "Permiso Denegado", "No posee permiso para importar personal de estación.")
            return
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Importar Personal de Estación",
            "",
            "Archivos Excel (*.xlsx *.xls)",
        )
        if not path:
            return
        try:
            cant = self.excel_service.import_personal_estacion_sheet(path, clear_existing=False)
            QMessageBox.information(self, "Importación Exitosa", f"Se procesaron {cant} registros correctamente.")
            self.recargar_todo()
        except Exception as e:
            QMessageBox.critical(self, "Error de Importación", f"Error al importar el archivo Excel:\n{e}")


class AdminCatalogosDialog(BaseMaximizedDialog):
    """Gestión de Tipos de Novedad, Dotaciones y Personal de Estación."""

    def __init__(
        self,
        records_service: RecordsService,
        excel_service: ExcelService,
        usuario_actual: Optional[Dict[str, Any]] = None,
        tiene_permiso_fn=None,
        parent=None,
    ):
        super().__init__(parent)
        self.setWindowTitle("Catálogos Operativos del Sistema")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)

        self.widget = AdminCatalogosWidget(
            records_service,
            excel_service,
            usuario_actual=usuario_actual,
            tiene_permiso_fn=tiene_permiso_fn,
            parent=self,
        )
        layout.addWidget(self.widget)

        btn_cerrar = QPushButton("Cerrar")
        btn_cerrar.setObjectName("btnSecondary")
        btn_cerrar.clicked.connect(self.accept)
        layout.addWidget(btn_cerrar, alignment=Qt.AlignRight)

    def recargar_todo(self) -> None:
        self.widget.recargar_todo()


# ---------------------------------------------------------------------------
# 5. Configuración General y Destinatarios de Correo
# ---------------------------------------------------------------------------

class AdminDestinatariosWidget(QWidget):
    """Gestión de destinatarios de correo para Novedades (Informe) y Ordenamientos Operativos."""

    def __init__(
        self,
        records_service: RecordsService,
        usuario_actual: Optional[Dict[str, Any]] = None,
        tiene_permiso_fn=None,
        parent=None,
    ):
        super().__init__(parent)
        self.records_service = records_service
        self.usuario_actual = usuario_actual or {}
        self.tiene_permiso_fn = tiene_permiso_fn

        self._datos_informe: List[Dict[str, Any]] = []
        self._datos_ordenamiento: List[Dict[str, Any]] = []

        layout = QVBoxLayout(self)
        layout.setContentsMargins(6, 6, 6, 6)
        layout.setSpacing(8)

        self.tabs = QTabWidget()

        # Tab Informe
        self.tab_inf, self.table_inf, self.txt_search_inf = self._crear_panel_tipo("informe")
        self.tabs.addTab(self.tab_inf, "📋 Novedades (Informe Diario)")

        # Tab Ordenamiento
        self.tab_ord, self.table_ord, self.txt_search_ord = self._crear_panel_tipo("ordenamiento")
        self.tabs.addTab(self.tab_ord, "🚂 Ordenamientos Operativos")

        layout.addWidget(self.tabs)
        self.recargar()

    def _can(self, perm: str) -> bool:
        if self.tiene_permiso_fn:
            return bool(self.tiene_permiso_fn(perm))
        return True

    def _u_info(self):
        u_id = self.usuario_actual.get("id") if self.usuario_actual else None
        win_u = self.usuario_actual.get("windows_user") if self.usuario_actual else None
        if not win_u:
            from src.core.config import obtener_usuario_windows
            win_u = obtener_usuario_windows()
        return u_id, win_u

    def _crear_panel_tipo(self, tipo: str):
        widget = QWidget()
        lay = QVBoxLayout(widget)
        lay.setContentsMargins(8, 8, 8, 8)
        lay.setSpacing(8)

        if tipo == "informe":
            lbl_desc = QLabel(
                "<b>Lista de distribución para el Informe Diario de Novedades:</b> "
                "Los contactos activos en esta lista recibirán el consolidado diario de novedades operativas por correo."
            )
        else:
            lbl_desc = QLabel(
                "<b>Lista de distribución para Ordenamientos Operativos:</b> "
                "Los contactos activos en esta lista recibirán los ordenamientos de trenes transmitidos desde el módulo operativo."
            )
        lbl_desc.setWordWrap(True)
        lbl_desc.setStyleSheet("color: #475569; margin-bottom: 4px;")
        lay.addWidget(lbl_desc)

        # Toolbar
        bar = QHBoxLayout()
        btn_nuevo = QPushButton("➕ Nuevo Destinatario")
        btn_nuevo.clicked.connect(lambda: self._nuevo_destinatario(tipo))
        bar.addWidget(btn_nuevo)

        btn_editar = QPushButton("✏️ Modificar")
        btn_editar.setObjectName("btnSecondary")
        btn_editar.clicked.connect(lambda: self._modificar_destinatario(tipo))
        bar.addWidget(btn_editar)

        btn_toggle = QPushButton("🔄 Activar / Desactivar")
        btn_toggle.setObjectName("btnSecondary")
        btn_toggle.clicked.connect(lambda: self._toggle_destinatario(tipo))
        bar.addWidget(btn_toggle)

        bar.addStretch()

        bar.addWidget(QLabel("Buscar:"))
        txt_search = QLineEdit()
        txt_search.setPlaceholderText("Nombre o correo...")
        txt_search.setClearButtonEnabled(True)
        txt_search.setMaximumWidth(250)
        bar.addWidget(txt_search)

        lay.addLayout(bar)

        # Tabla
        table = QTableWidget()
        table.setColumnCount(4)
        table.setHorizontalHeaderLabels(["ID", "NOMBRE / ÁREA", "CORREO ELECTRÓNICO", "ESTADO"])
        table.setSelectionBehavior(QTableWidget.SelectRows)
        table.setSelectionMode(QTableWidget.SingleSelection)
        table.setAlternatingRowColors(True)
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        table.verticalHeader().setVisible(False)
        table.doubleClicked.connect(lambda: self._modificar_destinatario(tipo))

        txt_search.textChanged.connect(lambda text: self._filtrar_tabla(table, text))

        lay.addWidget(table)
        return widget, table, txt_search

    def _filtrar_tabla(self, table: QTableWidget, text: str) -> None:
        term = text.strip().lower()
        for r in range(table.rowCount()):
            coincide = False
            for c in range(table.columnCount()):
                item = table.item(r, c)
                if item and term in item.text().lower():
                    coincide = True
                    break
            table.setRowHidden(r, not coincide)

    def recargar(self) -> None:
        self.recargar_informe()
        self.recargar_ordenamiento()

    def recargar_informe(self) -> None:
        try:
            self._datos_informe = self.records_service.listar_destinatarios_informe(incluir_inactivos=True)
        except Exception:
            self._datos_informe = []
        self._popular_tabla(self.table_inf, self._datos_informe)
        self._filtrar_tabla(self.table_inf, self.txt_search_inf.text())

    def recargar_ordenamiento(self) -> None:
        try:
            self._datos_ordenamiento = self.records_service.listar_destinatarios_ordenamiento(incluir_inactivos=True)
        except Exception:
            self._datos_ordenamiento = []
        self._popular_tabla(self.table_ord, self._datos_ordenamiento)
        self._filtrar_tabla(self.table_ord, self.txt_search_ord.text())

    def _popular_tabla(self, table: QTableWidget, datos: List[Dict[str, Any]]) -> None:
        table.setRowCount(len(datos))
        for r, d in enumerate(datos):
            item_id = QTableWidgetItem(str(d["id"]))
            item_id.setTextAlignment(Qt.AlignCenter)
            item_nom = QTableWidgetItem(str(d["nombre"]))
            item_email = QTableWidgetItem(str(d["email"]))

            activo = bool(d.get("activo"))
            item_est = QTableWidgetItem("Activo" if activo else "Inactivo")
            item_est.setTextAlignment(Qt.AlignCenter)
            if not activo:
                item_est.setForeground(Qt.gray)

            table.setItem(r, 0, item_id)
            table.setItem(r, 1, item_nom)
            table.setItem(r, 2, item_email)
            table.setItem(r, 3, item_est)

    def _get_permiso(self, tipo: str) -> str:
        return "destinatarios.administrar" if tipo == "informe" else "destinatarios_ordenamiento.administrar"

    def _nuevo_destinatario(self, tipo: str) -> None:
        perm = self._get_permiso(tipo)
        if not self._can(perm):
            QMessageBox.warning(self, "Permiso Denegado", f"No posee permiso para administrar destinatarios de {tipo}.")
            return

        titulo = "Nuevo Destinatario - Novedades (Informe)" if tipo == "informe" else "Nuevo Destinatario - Ordenamientos Operativos"
        dlg = QDialog(self)
        dlg.setWindowTitle(titulo)
        dlg.setFixedSize(420, 180)
        lay = QFormLayout(dlg)
        lay.setSpacing(10)

        txt_nom = QLineEdit()
        txt_nom.setPlaceholderText("ej. Control de Trenes / Operaciones")
        txt_email = QLineEdit()
        txt_email.setPlaceholderText("ej. control@trenesargentinos.gob.ar")

        lay.addRow("Nombre / Área:", txt_nom)
        lay.addRow("Correo Electrónico:", txt_email)

        btn_box = QHBoxLayout()
        btn_box.addStretch()
        btn_cancel = QPushButton("Cancelar")
        btn_cancel.setObjectName("btnSecondary")
        btn_cancel.clicked.connect(dlg.reject)
        btn_guardar = QPushButton("Guardar")
        btn_guardar.clicked.connect(dlg.accept)
        btn_box.addWidget(btn_cancel)
        btn_box.addWidget(btn_guardar)
        lay.addRow(btn_box)

        if dlg.exec():
            nom = txt_nom.text().strip()
            email = txt_email.text().strip().lower()
            if not nom:
                QMessageBox.warning(self, "Validación", "El nombre o área del destinatario es obligatorio.")
                return
            if not email or "@" not in email or "." not in email:
                QMessageBox.warning(self, "Validación", "Ingrese una dirección de correo electrónico válida.")
                return

            u_id, win_u = self._u_info()
            try:
                if tipo == "informe":
                    self.records_service.crear_destinatario_informe(nom, email, user_id=u_id, windows_user=win_u)
                    self.recargar_informe()
                else:
                    self.records_service.crear_destinatario_ordenamiento(nom, email, user_id=u_id, windows_user=win_u)
                    self.recargar_ordenamiento()
                QMessageBox.information(self, "Éxito", "Destinatario agregado correctamente.")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"No se pudo registrar el destinatario:\n{e}")

    def _modificar_destinatario(self, tipo: str) -> None:
        perm = self._get_permiso(tipo)
        if not self._can(perm):
            QMessageBox.warning(self, "Permiso Denegado", f"No posee permiso para modificar destinatarios de {tipo}.")
            return

        table = self.table_inf if tipo == "informe" else self.table_ord
        datos = self._datos_informe if tipo == "informe" else self._datos_ordenamiento

        row = table.currentRow()
        if row < 0 or row >= len(datos):
            QMessageBox.information(self, "Selección", "Seleccione un destinatario de la tabla.")
            return

        item_id = table.item(row, 0)
        if not item_id:
            return
        try:
            d_id = int(item_id.text())
        except ValueError:
            return
        dest = next((d for d in datos if d["id"] == d_id), None)
        if not dest:
            return

        titulo = "Modificar Destinatario - Informe" if tipo == "informe" else "Modificar Destinatario - Ordenamientos"
        dlg = QDialog(self)
        dlg.setWindowTitle(titulo)
        dlg.setFixedSize(420, 210)
        lay = QFormLayout(dlg)
        lay.setSpacing(10)

        txt_nom = QLineEdit(str(dest["nombre"]))
        txt_email = QLineEdit(str(dest["email"]))
        chk_act = QCheckBox("Activo (recibe correos)")
        chk_act.setChecked(bool(dest.get("activo")))

        lay.addRow("Nombre / Área:", txt_nom)
        lay.addRow("Correo Electrónico:", txt_email)
        lay.addRow(chk_act)

        btn_box = QHBoxLayout()
        btn_box.addStretch()
        btn_cancel = QPushButton("Cancelar")
        btn_cancel.setObjectName("btnSecondary")
        btn_cancel.clicked.connect(dlg.reject)
        btn_guardar = QPushButton("Guardar Cambios")
        btn_guardar.clicked.connect(dlg.accept)
        btn_box.addWidget(btn_cancel)
        btn_box.addWidget(btn_guardar)
        lay.addRow(btn_box)

        if dlg.exec():
            nom = txt_nom.text().strip()
            email = txt_email.text().strip().lower()
            if not nom:
                QMessageBox.warning(self, "Validación", "El nombre o área del destinatario es obligatorio.")
                return
            if not email or "@" not in email or "." not in email:
                QMessageBox.warning(self, "Validación", "Ingrese una dirección de correo electrónico válida.")
                return

            u_id, win_u = self._u_info()
            try:
                if tipo == "informe":
                    self.records_service.actualizar_destinatario_informe(
                        dest["id"],
                        nombre=nom,
                        email=email,
                        activo=1 if chk_act.isChecked() else 0,
                        user_id=u_id,
                        windows_user=win_u,
                    )
                    self.recargar_informe()
                else:
                    self.records_service.actualizar_destinatario_ordenamiento(
                        dest["id"],
                        nombre=nom,
                        email=email,
                        activo=1 if chk_act.isChecked() else 0,
                        user_id=u_id,
                        windows_user=win_u,
                    )
                    self.recargar_ordenamiento()
                QMessageBox.information(self, "Éxito", "Destinatario modificado correctamente.")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"No se pudo modificar el destinatario:\n{e}")

    def _toggle_destinatario(self, tipo: str) -> None:
        perm = self._get_permiso(tipo)
        if not self._can(perm):
            QMessageBox.warning(self, "Permiso Denegado", f"No posee permiso para modificar destinatarios de {tipo}.")
            return

        table = self.table_inf if tipo == "informe" else self.table_ord
        datos = self._datos_informe if tipo == "informe" else self._datos_ordenamiento

        row = table.currentRow()
        if row < 0 or row >= len(datos):
            QMessageBox.information(self, "Selección", "Seleccione un destinatario de la tabla.")
            return

        item_id = table.item(row, 0)
        if not item_id:
            return
        try:
            d_id = int(item_id.text())
        except ValueError:
            return
        dest = next((d for d in datos if d["id"] == d_id), None)
        if not dest:
            return

        nuevo_act = 0 if dest.get("activo") else 1
        u_id, win_u = self._u_info()
        try:
            if tipo == "informe":
                self.records_service.actualizar_destinatario_informe(
                    dest["id"],
                    nombre=dest["nombre"],
                    email=dest["email"],
                    activo=nuevo_act,
                    user_id=u_id,
                    windows_user=win_u,
                )
                self.recargar_informe()
            else:
                self.records_service.actualizar_destinatario_ordenamiento(
                    dest["id"],
                    nombre=dest["nombre"],
                    email=dest["email"],
                    activo=nuevo_act,
                    user_id=u_id,
                    windows_user=win_u,
                )
                self.recargar_ordenamiento()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo cambiar el estado del destinatario:\n{e}")


class AdminDestinatariosDialog(BaseMaximizedDialog):
    """Diálogo maximizado de administración de destinatarios de correo."""

    def __init__(
        self,
        records_service: RecordsService,
        usuario_actual: Optional[Dict[str, Any]] = None,
        tiene_permiso_fn=None,
        parent=None,
    ):
        super().__init__(parent)
        self.setWindowTitle("Administración de Destinatarios de Correo")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)

        self.widget = AdminDestinatariosWidget(
            records_service,
            usuario_actual=usuario_actual,
            tiene_permiso_fn=tiene_permiso_fn,
            parent=self,
        )
        layout.addWidget(self.widget)

        btn_cerrar = QPushButton("Cerrar")
        btn_cerrar.setObjectName("btnSecondary")
        btn_cerrar.clicked.connect(self.accept)
        layout.addWidget(btn_cerrar, alignment=Qt.AlignRight)

    def recargar(self) -> None:
        self.widget.recargar()


class AdminConfigWidget(QWidget):
    """Panel de configuración de parámetros operativos, seguridad, sincronización Supabase y destinatarios."""

    def __init__(
        self,
        db,
        records_service: Optional[RecordsService] = None,
        usuario_actual: Optional[Dict[str, Any]] = None,
        tiene_permiso_fn=None,
        parent=None,
    ):
        super().__init__(parent)
        self.db = db
        if records_service is None:
            self.records_service = RecordsService(db)
        else:
            self.records_service = records_service
        self.usuario_actual = usuario_actual or {}
        self.tiene_permiso_fn = tiene_permiso_fn

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        # Usar pestañas para separar parámetros generales, Supabase y destinatarios
        self.tabs = QTabWidget()

        # Pestaña 1: Parámetros Generales
        tab_gen = QWidget()
        lay_gen_root = QHBoxLayout(tab_gen)
        lay_gen_root.setContentsMargins(15, 15, 15, 15)

        card_gen = QFrame()
        card_gen.setObjectName("cardFrame")
        card_gen.setMaximumWidth(620)
        lay_card_gen = QVBoxLayout(card_gen)
        lay_card_gen.setContentsMargins(15, 15, 15, 15)
        lay_card_gen.setSpacing(14)

        # 1. Grupo Ubicación persistente de la Base de Datos SQLite
        from src.core.config import obtener_ruta_base_datos, guardar_ruta_base_datos
        sqlite_actual, _ = obtener_ruta_base_datos()

        grp_db = QGroupBox("Base de Datos SQLite (Ubicación Persistente)")
        lay_db = QVBoxLayout(grp_db)
        lay_db.setSpacing(6)

        lbl_info_db = QLabel("Ruta del archivo de base de datos (.sqlite):")
        lbl_info_db.setStyleSheet("color: #64748B; font-size: 8.5pt;")
        lay_db.addWidget(lbl_info_db)

        bar_db = QHBoxLayout()
        self.txt_db_path = QLineEdit(str(sqlite_actual))
        self.txt_db_path.setReadOnly(True)
        self.txt_db_path.setPlaceholderText("Ruta del archivo RENO.sqlite...")
        self.btn_cambiar_db = QPushButton("📂 Examinar...")
        self.btn_cambiar_db.setObjectName("btnSecondary")
        self.btn_cambiar_db.setFixedWidth(110)
        self.btn_cambiar_db.clicked.connect(self._examinar_base_datos)
        bar_db.addWidget(self.txt_db_path, 1)
        bar_db.addWidget(self.btn_cambiar_db)
        lay_db.addLayout(bar_db)

        lay_card_gen.addWidget(grp_db)

        # 2. Parámetros Operativos
        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignLeft)
        form.setFieldGrowthPolicy(QFormLayout.FieldsStayAtSizeHint)

        # Ventanas
        self.spin_edit = QSpinBox()
        self.spin_edit.setRange(1, 720)
        self.spin_edit.setFixedWidth(120)
        form.addRow("Ventana para editar registros (horas):", self.spin_edit)

        self.spin_elim = QSpinBox()
        self.spin_elim.setRange(1, 720)
        self.spin_elim.setFixedWidth(120)
        form.addRow("Ventana para eliminar registros (horas):", self.spin_elim)

        # Claves
        self.spin_exp = QSpinBox()
        self.spin_exp.setRange(0, 365)
        self.spin_exp.setFixedWidth(120)
        form.addRow("Expiración de contraseñas (días, 0=desactivada):", self.spin_exp)

        self.txt_dom = QLineEdit()
        self.txt_dom.setMaximumWidth(360)
        form.addRow("Dominio de correo institucional:", self.txt_dom)

        # Notificaciones
        self.chk_notif = QCheckBox("Activar notificaciones toast de nuevos registros")
        form.addRow(self.chk_notif)

        self.spin_toast = QSpinBox()
        self.spin_toast.setRange(2, 30)
        self.spin_toast.setFixedWidth(120)
        form.addRow("Duración de notificaciones en pantalla (segundos):", self.spin_toast)

        lay_card_gen.addLayout(form)
        lay_card_gen.addStretch()

        lay_gen_root.addWidget(card_gen)
        lay_gen_root.addStretch(1)
        self.tabs.addTab(tab_gen, "⚙️ Parámetros Generales")

        # Pestaña 2: Sincronización Supabase (Nube)
        tab_sync = QWidget()
        lay_sync_root = QVBoxLayout(tab_sync)
        lay_sync_root.setContentsMargins(10, 10, 10, 10)

        scroll_sync = QScrollArea()
        scroll_sync.setWidgetResizable(True)
        scroll_sync.setFrameShape(QFrame.NoFrame)

        scroll_content = QWidget()
        lay_scroll = QHBoxLayout(scroll_content)
        lay_scroll.setContentsMargins(5, 5, 5, 5)

        card_sync = QFrame()
        card_sync.setObjectName("cardFrame")
        card_sync.setMaximumWidth(600)
        lay_card_sync = QVBoxLayout(card_sync)
        lay_card_sync.setContentsMargins(15, 15, 15, 15)
        lay_card_sync.setSpacing(14)

        form_sync = QFormLayout()
        form_sync.setLabelAlignment(Qt.AlignLeft)
        form_sync.setFieldGrowthPolicy(QFormLayout.FieldsStayAtSizeHint)

        self.chk_sync_enabled = QCheckBox("Habilitar sincronización híbrida con Supabase")
        form_sync.addRow(self.chk_sync_enabled)

        self.txt_supa_url = QLineEdit()
        self.txt_supa_url.setPlaceholderText("https://xxxx.supabase.co")
        self.txt_supa_url.setMaximumWidth(380)
        form_sync.addRow("Supabase URL:", self.txt_supa_url)

        self.txt_supa_key = QLineEdit()
        self.txt_supa_key.setEchoMode(QLineEdit.Password)
        self.txt_supa_key.setPlaceholderText("sb_publishable_...")
        self.txt_supa_key.setMaximumWidth(380)
        form_sync.addRow("Supabase Publishable Key:", self.txt_supa_key)

        self.spin_sync_interval = QSpinBox()
        self.spin_sync_interval.setRange(1, 60)
        self.spin_sync_interval.setValue(5)
        self.spin_sync_interval.setSuffix(" min")
        self.spin_sync_interval.setFixedWidth(120)
        form_sync.addRow("Intervalo de auto-sincronización:", self.spin_sync_interval)

        lay_card_sync.addLayout(form_sync)

        # 🌐 Red y Proxy Corporativo (Solución HTTP 407)
        self.box_proxy = QGroupBox("🌐 Red y Proxy Corporativo (Solución HTTP 407)")
        lay_proxy = QVBoxLayout(self.box_proxy)
        lay_proxy.setSpacing(10)

        self.cmb_proxy_mode = QComboBox()
        for modo_id, modo_desc in OPCIONES_MODO:
            self.cmb_proxy_mode.addItem(modo_desc, modo_id)

        lay_mode = QFormLayout()
        lay_mode.addRow("Modo de Conexión / Proxy:", self.cmb_proxy_mode)
        lay_proxy.addLayout(lay_mode)

        # Sub-sección SSO Windows
        self.widget_sso = QWidget()
        lay_sso = QFormLayout(self.widget_sso)
        lay_sso.setContentsMargins(0, 4, 0, 4)
        self.txt_proxy_sso = QLineEdit()
        self.txt_proxy_sso.setPlaceholderText("Vacío = Autodetectar desde Windows (Internet Settings / WPAD)")
        self.txt_proxy_sso.setMaximumWidth(380)
        self.lbl_proxy_sso_info = QLabel("<b>Proxy detectado en Windows:</b> ...")
        self.lbl_proxy_sso_info.setStyleSheet("color:#1e40af; font-size:11px;")
        lay_sso.addRow("Servidor:Puerto Proxy (opcional):", self.txt_proxy_sso)
        lay_sso.addRow("", self.lbl_proxy_sso_info)
        lay_proxy.addWidget(self.widget_sso)

        # Sub-sección Proxy Manual
        self.widget_manual = QWidget()
        lay_manual = QFormLayout(self.widget_manual)
        lay_manual.setContentsMargins(0, 4, 0, 4)

        self.cmb_proxy_auth = QComboBox()
        for auth_id, auth_desc in OPCIONES_AUTH_PROXY:
            self.cmb_proxy_auth.addItem(auth_desc, auth_id)

        self.txt_proxy_host = QLineEdit()
        self.txt_proxy_host.setPlaceholderText("ej. proxy.empresa.com.ar o 10.0.0.1")
        self.txt_proxy_host.setMaximumWidth(260)

        self.spin_proxy_port = QSpinBox()
        self.spin_proxy_port.setRange(1, 65535)
        self.spin_proxy_port.setValue(8080)
        self.spin_proxy_port.setFixedWidth(90)

        box_host_port = QHBoxLayout()
        box_host_port.addWidget(self.txt_proxy_host, 1)
        box_host_port.addWidget(QLabel("Puerto:"))
        box_host_port.addWidget(self.spin_proxy_port)

        self.txt_proxy_user = QLineEdit()
        self.txt_proxy_user.setPlaceholderText("ej. DOMINIO\\usuario o usuario")
        self.txt_proxy_user.setMaximumWidth(380)

        self.txt_proxy_pass = QLineEdit()
        self.txt_proxy_pass.setEchoMode(QLineEdit.Password)
        self.txt_proxy_pass.setMaximumWidth(330)

        self.btn_toggle_pass = QPushButton("👁️")
        self.btn_toggle_pass.setFixedWidth(36)
        self.btn_toggle_pass.setToolTip("Mostrar / Ocultar Contraseña")
        self.btn_toggle_pass.clicked.connect(self._toggle_ver_pass_proxy)

        box_pass = QHBoxLayout()
        box_pass.addWidget(self.txt_proxy_pass, 1)
        box_pass.addWidget(self.btn_toggle_pass)

        lay_manual.addRow("Autenticación del Proxy:", self.cmb_proxy_auth)
        lay_manual.addRow("Servidor / IP Proxy:", box_host_port)
        lay_manual.addRow("Usuario Proxy:", self.txt_proxy_user)
        lay_manual.addRow("Contraseña Proxy:", box_pass)
        lay_proxy.addWidget(self.widget_manual)

        self.cmb_proxy_mode.currentIndexChanged.connect(self._actualizar_visibilidad_proxy)
        lay_card_sync.addWidget(self.box_proxy)

        # Botones de prueba y sincronización total
        btn_test_box = QHBoxLayout()
        self.btn_probar_supa = QPushButton("🔌 Probar Conexión Ahora (Testear Modo Seleccionado)")
        self.btn_probar_supa.setObjectName("btnSecondary")
        self.btn_probar_supa.clicked.connect(self._probar_conexion_supa)
        btn_test_box.addWidget(self.btn_probar_supa)

        self.btn_sync_total = QPushButton("⬆️ Forzar Sincronización Total (Subir Toda la Base a Supabase)")
        self.btn_sync_total.setObjectName("btnPrimary")
        self.btn_sync_total.clicked.connect(self._forzar_sincronizacion_total)
        btn_test_box.addWidget(self.btn_sync_total)

        self.btn_vaciar_y_subir = QPushButton("🗑️ Vaciar Supabase y Re-subir Todo")
        self.btn_vaciar_y_subir.setObjectName("btnDanger")
        self.btn_vaciar_y_subir.setToolTip("Elimina todos los datos en Supabase y vuelve a subir toda la base local limpia (ideal para pase a producción)")
        self.btn_vaciar_y_subir.clicked.connect(self._vaciar_y_resubir_supabase)
        btn_test_box.addWidget(self.btn_vaciar_y_subir)

        btn_test_box.addStretch()
        lay_card_sync.addLayout(btn_test_box)
        lay_card_sync.addStretch()

        lay_scroll.addWidget(card_sync)
        lay_scroll.addStretch(1)
        scroll_sync.setWidget(scroll_content)
        lay_sync_root.addWidget(scroll_sync)
        self.tabs.addTab(tab_sync, "☁️ Sincronización Supabase")

        # Pestaña 3: Destinatarios de Correo
        self.widget_destinatarios = AdminDestinatariosWidget(
            self.records_service,
            usuario_actual=self.usuario_actual,
            tiene_permiso_fn=self.tiene_permiso_fn,
            parent=self,
        )
        self.tabs.addTab(self.widget_destinatarios, "📧 Destinatarios de Correo")

        # Tab 4: Actualizaciones Automáticas (OTA)
        self.tab_updater = self._crear_tab_updater()
        self.tabs.addTab(self.tab_updater, "🚀 Actualizaciones (OTA)")

        # Tab 5: Diagnóstico y Logs TXT (Réplica Horos)
        self.tab_logs = self._crear_tab_logs()
        self.tabs.addTab(self.tab_logs, "📄 Diagnóstico y Logs TXT")

        layout.addWidget(self.tabs)

        # Botón Guardar
        btn_box = QHBoxLayout()
        btn_box.addStretch()
        self.btn_save = QPushButton("💾 Guardar Configuración")
        self.btn_save.clicked.connect(self._guardar)
        btn_box.addWidget(self.btn_save)
        layout.addLayout(btn_box)

        self.tabs.currentChanged.connect(self._al_cambiar_pestana)

        self._cargar()

    def recargar(self) -> None:
        self._cargar()
        if hasattr(self, "widget_destinatarios"):
            self.widget_destinatarios.recargar()

    def _al_cambiar_pestana(self, idx: int) -> None:
        self.btn_save.setVisible(idx not in (2, 4))
        if idx == 4:
            self._actualizar_visor_logs()

    def _crear_tab_updater(self) -> QWidget:
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(14)

        from src.core.config import APP_NAME, APP_VERSION

        box_info = QGroupBox("Estado de Versión del Sistema")
        b_info = QFormLayout(box_info)
        lbl_app = QLabel(f"<b>{APP_NAME}</b> — Registro de Novedades y Cambios de Turnos")
        lbl_ver = QLabel(f"<b style='color: #0284C7; font-size: 13pt;'>v{APP_VERSION}</b>")
        b_info.addRow("Aplicación:", lbl_app)
        b_info.addRow("Versión Instalada:", lbl_ver)
        layout.addWidget(box_info)

        box_feed = QGroupBox("Configuración de URL y Canal de Actualización (version.json)")
        b_feed = QVBoxLayout(box_feed)
        b_feed.setSpacing(10)

        self.chk_updater_auto = QCheckBox("Comprobar actualizaciones automáticamente al iniciar la aplicación (a los 5 segundos)")
        self.chk_updater_auto.setStyleSheet("font-weight: bold;")
        b_feed.addWidget(self.chk_updater_auto)

        f_feed = QFormLayout()
        feed_row = QHBoxLayout()
        self.txt_updater_url = QLineEdit()
        self.txt_updater_url.setPlaceholderText("https://tu-proyecto.supabase.co/storage/v1/object/public/updates/version.json")
        btn_sugerir = QPushButton("↺ Desde Supabase")
        btn_sugerir.setToolTip("Construir URL automáticamente usando la URL de Supabase configurada")
        btn_sugerir.clicked.connect(self._sugerir_url_updates)
        feed_row.addWidget(self.txt_updater_url, 1)
        feed_row.addWidget(btn_sugerir)
        f_feed.addRow("URL del Feed de Updates:", feed_row)
        b_feed.addLayout(f_feed)

        lbl_hint = QLabel(
            "Si se deja vacío, RENO intentará resolver automáticamente desde "
            "<i>&lt;Supabase URL&gt;/storage/v1/object/public/updates/version.json</i>."
        )
        lbl_hint.setWordWrap(True)
        lbl_hint.setStyleSheet("color: #64748B; font-size: 11px;")
        b_feed.addWidget(lbl_hint)

        layout.addWidget(box_feed)

        box_actions = QGroupBox("Comprobación Manual de Actualizaciones")
        b_act = QHBoxLayout(box_actions)
        self.btn_buscar_updates_cfg = QPushButton("🔍 Buscar Actualizaciones Ahora")
        self.btn_buscar_updates_cfg.setObjectName("btnPrimary")
        self.btn_buscar_updates_cfg.clicked.connect(self._comprobar_actualizaciones_manual)
        b_act.addWidget(self.btn_buscar_updates_cfg)
        self.lbl_updater_status = QLabel("Listo para comprobar.")
        self.lbl_updater_status.setStyleSheet("color: #64748B; font-size: 11px; margin-left: 10px;")
        b_act.addWidget(self.lbl_updater_status, 1)
        layout.addWidget(box_actions)

        layout.addStretch()
        return widget

    def _sugerir_url_updates(self) -> None:
        sb_url = self.txt_supa_url.text().strip().rstrip("/")
        if not sb_url:
            QMessageBox.information(
                self,
                "URL de Updates",
                "Ingrese primero la 'Supabase URL' en la pestaña Sincronización y Red.",
            )
            return
        from src.core.config import cargar_configuracion_json

        bucket = str(cargar_configuracion_json().get("supabase_storage_bucket") or "updates").strip().strip("/") or "updates"
        self.txt_updater_url.setText(f"{sb_url}/storage/v1/object/public/{bucket}/version.json")

    def _comprobar_actualizaciones_manual(self) -> None:
        from src.services.updater_service import VerificarActualizacionWorker

        self.btn_buscar_updates_cfg.setEnabled(False)
        self.lbl_updater_status.setText("Conectando con el servidor de actualizaciones...")
        cfg_temp = self._recopilar_config_pantalla()

        self._worker_cfg_updater = VerificarActualizacionWorker(config=cfg_temp, parent=self)

        def _on_disponible(info, es_obligatoria: bool) -> None:
            self.btn_buscar_updates_cfg.setEnabled(True)
            self.lbl_updater_status.setText(f"🚀 Nueva versión disponible: v{info.version}")
            from src.ui.views.dialogs.actualizacion_dialog import ActualizacionDialog

            dlg = ActualizacionDialog(info, es_obligatoria=es_obligatoria, parent=self, config=cfg_temp)
            dlg.exec()

        def _on_actualizado() -> None:
            self.btn_buscar_updates_cfg.setEnabled(True)
            self.lbl_updater_status.setText("✅ Ya tienes la última versión instalada.")
            QMessageBox.information(self, "Actualizaciones", "La aplicación está actualizada a la última versión.")

        def _on_error(err: str) -> None:
            self.btn_buscar_updates_cfg.setEnabled(True)
            self.lbl_updater_status.setText(f"❌ Error: {err}")
            QMessageBox.warning(self, "Actualizaciones", f"No se pudo comprobar actualizaciones:\n{err}")

        self._worker_cfg_updater.actualizacion_disponible.connect(_on_disponible)
        self._worker_cfg_updater.sistema_actualizado.connect(_on_actualizado)
        self._worker_cfg_updater.error_verificacion.connect(_on_error)
        self._worker_cfg_updater.start()

    def _crear_tab_logs(self) -> QWidget:
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        box = QGroupBox("Registro de Depuración y Logs de Consola Persistidos")
        b_layout = QVBoxLayout(box)
        b_layout.setSpacing(10)

        desc = QLabel(
            "La salida de consola (prints, logs y excepciones) se registra automáticamente en un archivo de texto .txt. "
            "Esto permite diagnosticar fallos y tracebacks en terminales de usuario empaquetadas sin ventana de consola."
        )
        desc.setStyleSheet("color: #64748B; font-size: 11px;")
        desc.setWordWrap(True)
        b_layout.addWidget(desc)

        btn_bar = QHBoxLayout()
        btn_notepad = QPushButton("📄 Abrir Archivo de Log (TXT en Bloc de Notas)")
        btn_notepad.setObjectName("btnPrimary")
        btn_notepad.clicked.connect(self._abrir_log_notepad)
        btn_bar.addWidget(btn_notepad)

        btn_folder = QPushButton("📂 Abrir Carpeta de Logs")
        btn_folder.clicked.connect(self._abrir_carpeta_logs)
        btn_bar.addWidget(btn_folder)

        btn_bar.addStretch()

        btn_refresh_logs = QPushButton("🔄 Actualizar Visor")
        btn_refresh_logs.clicked.connect(self._actualizar_visor_logs)
        btn_bar.addWidget(btn_refresh_logs)

        btn_copy = QPushButton("📋 Copiar Registro")
        btn_copy.clicked.connect(self._copiar_logs)
        btn_bar.addWidget(btn_copy)

        b_layout.addLayout(btn_bar)

        self.txt_logs_viewer = QPlainTextEdit()
        self.txt_logs_viewer.setReadOnly(True)
        self.txt_logs_viewer.setStyleSheet(
            "background-color: #0F172A; color: #E2E8F0; font-family: monospace; font-size: 11px; "
            "border: 1px solid #334155; border-radius: 4px;"
        )
        b_layout.addWidget(self.txt_logs_viewer, 1)

        layout.addWidget(box, 1)
        return widget

    def _abrir_log_notepad(self) -> None:
        if not abrir_en_bloc_de_notas():
            QMessageBox.warning(
                self,
                "Logs",
                f"No se pudo invocar el editor automático. Archivo en:\n{obtener_ruta_archivo_log()}",
            )

    def _abrir_carpeta_logs(self) -> None:
        if not abrir_carpeta_en_explorador():
            QMessageBox.warning(self, "Logs", "No se pudo abrir el explorador de archivos.")

    def _actualizar_visor_logs(self) -> None:
        texto = leer_ultimas_lineas_log(max_lineas=300)
        self.txt_logs_viewer.setPlainText(texto)
        sb = self.txt_logs_viewer.verticalScrollBar()
        if sb:
            sb.setValue(sb.maximum())

    def _copiar_logs(self) -> None:
        texto = self.txt_logs_viewer.toPlainText()
        if texto:
            QApplication.clipboard().setText(texto)
            QMessageBox.information(self, "Logs", "Registro copiado al portapapeles.")

    def _examinar_base_datos(self) -> None:
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Seleccionar o Definir Ubicación de Base de Datos SQLite",
            self.txt_db_path.text(),
            "Base de Datos SQLite (*.sqlite *.db)",
        )
        if file_path:
            self.txt_db_path.setText(file_path)

    def _actualizar_visibilidad_proxy(self) -> None:
        modo = self.cmb_proxy_mode.currentData() or MODO_DIRECTO
        self.widget_sso.setVisible(modo == MODO_WINDOWS_SSO)
        self.widget_manual.setVisible(modo == MODO_MANUAL)

    def _toggle_ver_pass_proxy(self) -> None:
        if self.txt_proxy_pass.echoMode() == QLineEdit.Password:
            self.txt_proxy_pass.setEchoMode(QLineEdit.Normal)
            self.btn_toggle_pass.setText("🔒")
        else:
            self.txt_proxy_pass.setEchoMode(QLineEdit.Password)
            self.btn_toggle_pass.setText("👁️")

    def _recopilar_config_pantalla(self) -> dict[str, Any]:
        from src.core.config import cargar_configuracion_json

        base_cfg = cargar_configuracion_json()
        return {
            "supabase_url": self.txt_supa_url.text().strip(),
            "supabase_anon_key": self.txt_supa_key.text().strip(),
            "supabase_storage_bucket": str(base_cfg.get("supabase_storage_bucket") or "updates").strip() or "updates",
            "updater_url": self.txt_updater_url.text().strip(),
            "updater_auto_check": self.chk_updater_auto.isChecked(),
            "proxy_mode": self.cmb_proxy_mode.currentData() or MODO_DIRECTO,
            "proxy_sso_server": self.txt_proxy_sso.text().strip(),
            "proxy_auth": self.cmb_proxy_auth.currentData() or AUTH_MANUAL_FIJO,
            "proxy_host": self.txt_proxy_host.text().strip(),
            "proxy_port": self.spin_proxy_port.value(),
            "proxy_user": self.txt_proxy_user.text().strip(),
            "proxy_pass": self.txt_proxy_pass.text(),
        }

    def _cargar(self) -> None:
        from src.core.config import obtener_ruta_base_datos
        sqlite_actual, _ = obtener_ruta_base_datos()
        self.txt_db_path.setText(str(sqlite_actual))

        with self.db.read_connection() as conn:
            rows = dict(conn.execute("SELECT clave, valor FROM configuracion").fetchall())
        self.spin_edit.setValue(int(rows.get("editar_horas", 24)))
        self.spin_elim.setValue(int(rows.get("eliminar_horas", 72)))
        self.spin_exp.setValue(int(rows.get("clave_expiracion_dias", 90)))
        self.txt_dom.setText(str(rows.get("correo_dominio", "@trenesargentinos.gob.ar")))
        self.chk_notif.setChecked(bool(int(rows.get("notificaciones_activo", 1))))
        self.spin_toast.setValue(int(rows.get("toast_duracion", 6)))

        from src.core.config import cargar_configuracion_json
        app_cfg = cargar_configuracion_json()
        self.chk_sync_enabled.setChecked(bool(app_cfg.get("autosync_habilitado", False)))
        self.txt_supa_url.setText(str(app_cfg.get("supabase_url", "")))
        self.txt_supa_key.setText(str(app_cfg.get("supabase_anon_key", "")))
        self.spin_sync_interval.setValue(int(app_cfg.get("autosync_intervalo_min", 5)))

        modo_guardado = app_cfg.get("proxy_mode", MODO_DIRECTO)
        idx_modo = self.cmb_proxy_mode.findData(modo_guardado)
        if idx_modo >= 0:
            self.cmb_proxy_mode.setCurrentIndex(idx_modo)

        self.txt_proxy_sso.setText(str(app_cfg.get("proxy_sso_server", "")))
        self.lbl_proxy_sso_info.setText(f"<b>Proxy detectado en Windows:</b> {detectar_proxy_sistema(app_cfg)}")

        auth_guardado = app_cfg.get("proxy_auth", AUTH_MANUAL_FIJO)
        idx_auth = self.cmb_proxy_auth.findData(auth_guardado)
        if idx_auth >= 0:
            self.cmb_proxy_auth.setCurrentIndex(idx_auth)

        self.txt_proxy_host.setText(str(app_cfg.get("proxy_host", "")))
        self.spin_proxy_port.setValue(int(app_cfg.get("proxy_port", 8080) or 8080))
        self.txt_proxy_user.setText(str(app_cfg.get("proxy_user", "")))
        self.txt_proxy_pass.setText(str(app_cfg.get("proxy_pass", "")))

        self.txt_updater_url.setText(str(app_cfg.get("updater_url", "")))
        self.chk_updater_auto.setChecked(bool(app_cfg.get("updater_auto_check", True)))

        self._actualizar_visibilidad_proxy()
        self._actualizar_visor_logs()

    def _probar_conexion_supa(self) -> None:
        cfg_test = self._recopilar_config_pantalla()
        url = cfg_test.get("supabase_url", "")
        key = cfg_test.get("supabase_anon_key", "")
        if not url or not key:
            QMessageBox.warning(self, "Supabase", "Ingrese tanto la URL como la Anon Key para realizar la prueba.")
            return

        self.btn_probar_supa.setEnabled(False)
        self.btn_probar_supa.setText("⏳ Probando conexión...")
        from PySide6.QtWidgets import QApplication
        QApplication.processEvents()

        try:
            res = probar_conexion_supabase(cfg_test)
            ok = res.get("ok", False)
            modo = res.get("modo_usado", "—")
            proxy_det = res.get("proxy_detectado", "—")
            code = res.get("status_code", 0)
            det = res.get("detalle", "")

            msg = (
                f"<b>Resultado de la Prueba de Conexión:</b><br><br>"
                f"• <b>Modo Utilizado:</b> {modo}<br>"
                f"• <b>Proxy Detectado / Usado:</b> {proxy_det}<br>"
                f"• <b>Código HTTP:</b> {code or 'Sin respuesta'}<br><br>"
                f"<b>Detalle:</b><br>{det}"
            )

            if ok:
                QMessageBox.information(self, "Conexión Exitosa", msg)
            else:
                QMessageBox.warning(self, "Fallo de Conexión", msg)
        finally:
            self.btn_probar_supa.setEnabled(True)
            self.btn_probar_supa.setText("🔌 Probar Conexión Ahora (Testear Modo Seleccionado)")

    def _forzar_sincronizacion_total(self) -> None:
        resp = QMessageBox.question(
            self,
            "Sincronización Total a Supabase",
            "¿Desea marcar todas las tablas locales como pendientes y subirlas completamente a Supabase?\n\n"
            "Esta operación asegura que toda la base de datos histórica quede respaldada en la nube.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes,
        )
        if resp != QMessageBox.Yes:
            return

        from src.data.sync.outbox_queue import OutboxQueue
        from src.data.sync.supabase_sync import SupabaseSyncEngine
        from PySide6.QtWidgets import QApplication

        self.btn_sync_total.setEnabled(False)
        self.btn_sync_total.setText("⏳ Encolando registros locales...")
        QApplication.processEvents()

        try:
            outbox = OutboxQueue(self.db)
            total_encolados = outbox.forzar_resincronizacion_total()
            self.btn_sync_total.setText(f"⏳ Subiendo {total_encolados} registros a Supabase...")
            QApplication.processEvents()

            engine = SupabaseSyncEngine(self.db)
            enviados, recibidos = engine._ejecutar_sync(solo_push=True)
            pendientes = outbox.contar_pendientes()

            QMessageBox.information(
                self,
                "Sincronización Total Finalizada",
                f"<b>Sincronización Total Completada:</b><br><br>"
                f"• Registros encolados: {total_encolados}<br>"
                f"• Registros procesados y enviados: {enviados}<br>"
                f"• Pendientes restantes en cola: {pendientes}<br><br>"
                f"{'🟢 La base de datos está completamente sincronizada con Supabase.' if pendientes == 0 else '🟡 Quedan registros pendientes en cola.'}",
            )
        except Exception as ex:
            QMessageBox.critical(self, "Error de Sincronización", f"Ocurrió un error al sincronizar:\n\n{ex}")
        finally:
            self.btn_sync_total.setEnabled(True)
            self.btn_sync_total.setText("⬆️ Forzar Sincronización Total (Subir Toda la Base a Supabase)")

    def _vaciar_y_resubir_supabase(self) -> None:
        resp = QMessageBox.warning(
            self,
            "⚠️ Confirmación de Vaciado Remoto",
            "<b>¿Está seguro de que desea VACIAR completamente Supabase y volver a subir todo?</b><br><br>"
            "Esta operación:<br>"
            "1. Eliminará todos los registros actuales en Supabase (todas las tablas <code>reno_*</code>).<br>"
            "2. Encolará y subirá el 100% de la base de datos local SQLite de forma limpia e íntegra.<br><br>"
            "<i>Recomendado para el paso a Producción o para purgar datos de pruebas.</i>",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if resp != QMessageBox.Yes:
            return

        from src.data.sync.outbox_queue import OutboxQueue
        from src.data.sync.supabase_sync import SupabaseSyncEngine
        from PySide6.QtWidgets import QApplication

        self.btn_vaciar_y_subir.setEnabled(False)
        self.btn_sync_total.setEnabled(False)
        self.btn_vaciar_y_subir.setText("⏳ Vaciando tablas en Supabase...")
        QApplication.processEvents()

        try:
            engine = SupabaseSyncEngine(self.db)
            vaciadas = engine.vaciar_tablas_remotas()

            self.btn_vaciar_y_subir.setText("⏳ Encolando registros locales...")
            QApplication.processEvents()

            outbox = OutboxQueue(self.db)
            total_encolados = outbox.forzar_resincronizacion_total()

            self.btn_vaciar_y_subir.setText(f"⏳ Subiendo {total_encolados} registros a Supabase...")
            QApplication.processEvents()

            enviados, recibidos = engine._ejecutar_sync(solo_push=True)
            pendientes = outbox.contar_pendientes()

            QMessageBox.information(
                self,
                "Vaciado y Carga Completada",
                f"<b>Proceso Finalizado con Éxito:</b><br><br>"
                f"• Tablas remotas vaciadas: {vaciadas}<br>"
                f"• Registros locales encolados: {total_encolados}<br>"
                f"• Registros subidos a Supabase: {enviados}<br>"
                f"• Pendientes en cola: {pendientes}<br><br>"
                f"{'🟢 La base remota en Supabase ha quedado 100% sincronizada y limpia.' if pendientes == 0 else '🟡 Quedan registros pendientes en cola.'}",
            )
        except Exception as ex:
            QMessageBox.critical(self, "Error", f"Ocurrió un error durante el vaciado y carga:\n\n{ex}")
        finally:
            self.btn_vaciar_y_subir.setEnabled(True)
            self.btn_sync_total.setEnabled(True)
            self.btn_vaciar_y_subir.setText("🗑️ Vaciar Supabase y Re-subir Todo")

    def _guardar(self) -> None:
        with self.db.write_transaction() as conn:
            cfg = {
                "editar_horas": str(self.spin_edit.value()),
                "eliminar_horas": str(self.spin_elim.value()),
                "clave_expiracion_dias": str(self.spin_exp.value()),
                "correo_dominio": self.txt_dom.text().strip(),
                "notificaciones_activo": "1" if self.chk_notif.isChecked() else "0",
                "toast_duracion": str(self.spin_toast.value()),
            }
            for k, v in cfg.items():
                conn.execute(
                    "INSERT INTO configuracion (clave, valor) VALUES (?, ?) ON CONFLICT(clave) DO UPDATE SET valor=excluded.valor",
                    (k, v),
                )

        from src.core.config import cargar_configuracion_json, guardar_configuracion_json
        from src.core.theme import obtener_tema_visual
        app_cfg = cargar_configuracion_json()
        app_cfg["tema_visual"] = obtener_tema_visual(self.db)
        app_cfg["autosync_habilitado"] = self.chk_sync_enabled.isChecked()
        app_cfg["supabase_url"] = self.txt_supa_url.text().strip()
        app_cfg["supabase_anon_key"] = self.txt_supa_key.text().strip()
        app_cfg["autosync_intervalo_min"] = self.spin_sync_interval.value()

        app_cfg["proxy_mode"] = self.cmb_proxy_mode.currentData() or MODO_DIRECTO
        app_cfg["proxy_sso_server"] = self.txt_proxy_sso.text().strip()
        app_cfg["proxy_auth"] = self.cmb_proxy_auth.currentData() or AUTH_MANUAL_FIJO
        app_cfg["proxy_host"] = self.txt_proxy_host.text().strip()
        app_cfg["proxy_port"] = self.spin_proxy_port.value()
        app_cfg["proxy_user"] = self.txt_proxy_user.text().strip()
        app_cfg["proxy_pass"] = self.txt_proxy_pass.text()

        app_cfg["updater_url"] = self.txt_updater_url.text().strip()
        app_cfg["updater_auto_check"] = self.chk_updater_auto.isChecked()

        guardar_configuracion_json(app_cfg)
        aplicar_configuracion_proxy(app_cfg)

        from src.core.config import guardar_ruta_base_datos, obtener_ruta_base_datos
        sqlite_actual, _ = obtener_ruta_base_datos()
        nuevo_path = self.txt_db_path.text().strip()
        aviso_db = ""
        if nuevo_path and Path(nuevo_path) != Path(sqlite_actual):
            if guardar_ruta_base_datos(nuevo_path):
                aviso_db = (
                    f"\n\n⚠️ Ruta de Base de Datos Modificada:\n{nuevo_path}\n"
                    f"Para que el cambio surta efecto y comenzar a operar con la nueva base, reinicie la aplicación."
                )

        QMessageBox.information(self, "Éxito", f"Configuración guardada exitosamente.{aviso_db}")


class AdminConfigDialog(BaseMaximizedDialog):
    """Diálogo maximizado de configuración del sistema."""

    def __init__(
        self,
        db,
        records_service: Optional[RecordsService] = None,
        usuario_actual: Optional[Dict[str, Any]] = None,
        tiene_permiso_fn=None,
        parent=None,
    ):
        super().__init__(parent)
        self.setWindowTitle("Configuración del Sistema")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)

        self.widget = AdminConfigWidget(
            db,
            records_service=records_service,
            usuario_actual=usuario_actual,
            tiene_permiso_fn=tiene_permiso_fn,
            parent=self,
        )
        layout.addWidget(self.widget)

        btn_cerrar = QPushButton("Cerrar")
        btn_cerrar.setObjectName("btnSecondary")
        btn_cerrar.clicked.connect(self.accept)
        layout.addWidget(btn_cerrar, alignment=Qt.AlignRight)

    def recargar(self) -> None:
        self.widget.recargar()


# ---------------------------------------------------------------------------
# 6. Copias de Seguridad (Backups)
# ---------------------------------------------------------------------------

class AdminBackupsWidget(QWidget):
    """Panel gestor de copias de seguridad de la base SQLite."""

    def __init__(self, backup_service: BackupService, parent=None):
        super().__init__(parent)
        self.backup_service = backup_service

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        # 1. Barra de ruta de almacenamiento
        dir_bar = QHBoxLayout()
        dir_bar.addWidget(QLabel("<b>Carpeta de Almacenamiento:</b>"))
        self.txt_ruta_dir = QLineEdit(self.backup_service.ruta_directorio_backups())
        self.txt_ruta_dir.setReadOnly(True)
        dir_bar.addWidget(self.txt_ruta_dir, stretch=1)

        self.btn_cambiar_dir = QPushButton("📁 Cambiar Carpeta...")
        self.btn_cambiar_dir.setObjectName("btnSecondary")
        self.btn_cambiar_dir.clicked.connect(self._cambiar_directorio)
        dir_bar.addWidget(self.btn_cambiar_dir)
        layout.addLayout(dir_bar)

        # 2. Botones de acción
        bar = QHBoxLayout()
        self.btn_crear = QPushButton("💾 Crear Backup Ahora")
        self.btn_crear.clicked.connect(self._crear_backup)
        bar.addWidget(self.btn_crear)

        self.btn_guardar_en = QPushButton("💾 Guardar Copia en...")
        self.btn_guardar_en.setObjectName("btnSecondary")
        self.btn_guardar_en.clicked.connect(self._guardar_copia_en)
        bar.addWidget(self.btn_guardar_en)

        self.btn_restaurar = QPushButton("🔄 Restaurar Copia Seleccionada")
        self.btn_restaurar.setObjectName("btnDanger")
        self.btn_restaurar.clicked.connect(self._restaurar_backup)
        bar.addWidget(self.btn_restaurar)

        bar.addStretch()
        layout.addLayout(bar)

        headers = ["ARCHIVO DE BACKUP", "TAMAÑO", "FECHA Y HORA"]
        keys = ["nombre", "tamano_str", "fecha"]

        self.model = BaseTableModel(headers, keys, parent=self)
        self.table = QTableView()
        self.table.setModel(self.model)
        self.table.setSelectionBehavior(QTableView.SelectRows)
        self.table.setSelectionMode(QTableView.SingleSelection)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table.setColumnWidth(0, 360)
        self.table.setColumnWidth(1, 140)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        layout.addWidget(self.table)

        self.recargar()

    def recargar(self) -> None:
        if hasattr(self, "txt_ruta_dir"):
            self.txt_ruta_dir.setText(self.backup_service.ruta_directorio_backups())
        items = self.backup_service.listar_backups()
        datos = [
            {"nombre": n, "tamano_str": f"{sz / 1024:.1f} KB", "fecha": f}
            for n, sz, f in items
        ]
        self.model.set_data(datos)

    def _cambiar_directorio(self) -> None:
        curr = self.backup_service.ruta_directorio_backups()
        nueva_ruta = QFileDialog.getExistingDirectory(self, "Seleccionar Carpeta para Copias de Seguridad", curr)
        if nueva_ruta:
            self.backup_service.set_directorio_backups(nueva_ruta)
            self.txt_ruta_dir.setText(nueva_ruta)
            self.recargar()
            QMessageBox.information(self, "Carpeta Actualizada", f"La carpeta predeterminada de backups ahora es:\n{nueva_ruta}")

    def _guardar_copia_en(self) -> None:
        dest_dir = QFileDialog.getExistingDirectory(self, "Seleccionar Carpeta Destino para la Copia")
        if not dest_dir:
            return
        try:
            ruta = self.backup_service.crear_backup(directorio_destino=dest_dir)
            QMessageBox.information(self, "Backup Guardado", f"Copia de seguridad guardada con éxito en:\n{ruta}")
            self.recargar()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo guardar la copia:\n{e}")

    def _crear_backup(self) -> None:
        try:
            ruta = self.backup_service.crear_backup()
            QMessageBox.information(self, "Backup Creado", f"Copia creada exitosamente en:\n{ruta}")
            self.recargar()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo crear la copia:\n{e}")

    def _restaurar_backup(self) -> None:
        sel = self.table.selectionModel().selectedRows()
        if not sel:
            QMessageBox.information(self, "Restaurar", "Seleccione un archivo de backup en la tabla.")
            return

        item = self.model.get_row(sel[0].row())
        if not item:
            return

        resp = QMessageBox.warning(
            self,
            "Confirmación Crítica",
            f"¿Está completamente seguro de restaurar la base de datos a la versión {item['nombre']}?\n\n"
            f"Se creará una copia de seguridad preventiva automática antes de sobreescribir.",
            QMessageBox.Yes | QMessageBox.No,
        )
        if resp != QMessageBox.Yes:
            return

        try:
            self.backup_service.restaurar_backup(item["nombre"])
            QMessageBox.information(self, "Restauración Exitosa", "La base de datos fue restaurada con éxito.")
            self.recargar()
        except Exception as e:
            QMessageBox.critical(self, "Error de Restauración", f"Fallo al restaurar:\n{e}")


class AdminBackupsDialog(BaseMaximizedDialog):
    """Diálogo gestor de copias de seguridad de la base SQLite."""

    def __init__(self, backup_service: BackupService, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Copias de Seguridad (Backups)")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)

        self.widget = AdminBackupsWidget(backup_service, parent=self)
        layout.addWidget(self.widget)

        btn_cerrar = QPushButton("Cerrar")
        btn_cerrar.setObjectName("btnSecondary")
        btn_cerrar.clicked.connect(self.accept)
        layout.addWidget(btn_cerrar, alignment=Qt.AlignRight)

    def recargar(self) -> None:
        self.widget.recargar()


# ---------------------------------------------------------------------------
# 7. Registros Eliminados (Papelera)
# ---------------------------------------------------------------------------

class AdminEliminadosWidget(QWidget):
    """Panel de papelera de recuperación y borrado permanente de registros."""

    def __init__(self, records_service: RecordsService, usuario_actual: Optional[Dict[str, Any]] = None, parent=None):
        super().__init__(parent)
        self.records_service = records_service
        self.usuario_actual = usuario_actual or {}

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        bar = QHBoxLayout()
        self.btn_recuperar = QPushButton("♻️ Recuperar Registro")
        self.btn_recuperar.clicked.connect(self._recuperar)
        bar.addWidget(self.btn_recuperar)

        self.btn_borrar_def = QPushButton("❌ Borrado Definitivo")
        self.btn_borrar_def.setObjectName("btnDanger")
        self.btn_borrar_def.clicked.connect(self._borrar_definitivo)
        bar.addWidget(self.btn_borrar_def)

        bar.addStretch()
        layout.addLayout(bar)

        headers = ["TIPO", "ID", "REGISTRADO EN", "LEGAJO", "APELLIDOS Y NOMBRES", "DOTACIÓN", "DETALLE", "OBSERVACIONES"]
        keys = ["tipo", "id", "registrado_en", "legajo", "apellidos_nombres", "dotacion", "detalle", "observaciones"]

        self.model = BaseTableModel(headers, keys, parent=self)
        self.table = QTableView()
        self.table.setModel(self.model)
        self.table.setSelectionBehavior(QTableView.SelectRows)
        self.table.setSelectionMode(QTableView.SingleSelection)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table.setColumnWidth(0, 90)
        self.table.setColumnWidth(1, 60)
        self.table.setColumnWidth(2, 140)
        self.table.setColumnWidth(3, 80)
        self.table.setColumnWidth(4, 180)
        self.table.setColumnWidth(5, 110)
        self.table.horizontalHeader().setSectionResizeMode(6, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(7, QHeaderView.Stretch)
        layout.addWidget(self.table)

        self.recargar()

    def recargar(self) -> None:
        items = self.records_service.listar_eliminados()
        self.model.set_data(items)

    def _obtener_seleccionado(self) -> Optional[Dict[str, Any]]:
        sel = self.table.selectionModel().selectedRows()
        if not sel:
            return None
        return self.model.get_row(sel[0].row())

    def _recuperar(self) -> None:
        item = self._obtener_seleccionado()
        if not item:
            return
        try:
            self.records_service.recuperar_registro(
                item["tipo"], item["id"],
                user_id=self.usuario_actual.get("id"),
                windows_user=self.usuario_actual.get("username"),
            )
            QMessageBox.information(self, "Recuperado", f"El registro #{item['id']} ha sido restaurado exitosamente.")
            self.recargar()
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

    def _borrar_definitivo(self) -> None:
        item = self._obtener_seleccionado()
        if not item:
            return
        resp = QMessageBox.warning(
            self,
            "Borrado Definitivo",
            f"¿Está seguro de eliminar permanentemente el registro #{item['id']} ({item['tipo']})?\nEsta acción no se puede deshacer.",
            QMessageBox.Yes | QMessageBox.No,
        )
        if resp != QMessageBox.Yes:
            return
        try:
            self.records_service.borrar_definitivo(
                item["tipo"], item["id"],
                user_id=self.usuario_actual.get("id"),
                windows_user=self.usuario_actual.get("username"),
            )
            QMessageBox.information(self, "Eliminado", "Registro eliminado definitivamente.")
            self.recargar()
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))


class AdminEliminadosDialog(BaseMaximizedDialog):
    """Diálogo de papelera de recuperación y borrado permanente de registros."""

    def __init__(self, records_service: RecordsService, usuario_actual: Optional[Dict[str, Any]] = None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Registros Eliminados (Papelera)")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)

        self.widget = AdminEliminadosWidget(records_service, usuario_actual=usuario_actual, parent=self)
        layout.addWidget(self.widget)

        btn_cerrar = QPushButton("Cerrar")
        btn_cerrar.setObjectName("btnSecondary")
        btn_cerrar.clicked.connect(self.accept)
        layout.addWidget(btn_cerrar, alignment=Qt.AlignRight)

    def recargar(self) -> None:
        self.widget.recargar()


# ---------------------------------------------------------------------------
# 8. Auditoría del Sistema
# ---------------------------------------------------------------------------

class AdminAuditoriaWidget(QWidget):
    """Panel visor de logs de auditoría de modificaciones y accesos."""

    def __init__(self, records_service: RecordsService, excel_service=None, parent=None):
        super().__init__(parent)
        self.records_service = records_service
        if excel_service is None:
            from src.services.excel_service import ExcelService
            self.excel_service = ExcelService(records_service.db)
        else:
            self.excel_service = excel_service

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        bar = QHBoxLayout()
        bar.addWidget(QLabel("Filtrar auditoría:"))
        self.txt_filter = QLineEdit()
        self.txt_filter.setPlaceholderText("Usuario, acción, entidad...")
        self.txt_filter.textChanged.connect(self._on_search)
        bar.addWidget(self.txt_filter)

        self.btn_exportar = QPushButton("📊 Exportar a Excel")
        self.btn_exportar.setObjectName("btnSecondary")
        self.btn_exportar.clicked.connect(self._exportar_excel)
        bar.addWidget(self.btn_exportar)

        bar.addStretch()
        layout.addLayout(bar)

        headers = ["ID", "FECHA Y HORA", "USUARIO", "ACCIÓN", "ENTIDAD", "ID REGISTRO", "DATOS ANTERIORES", "DATOS NUEVOS"]
        keys = ["id", "creado_en", "usuario", "accion", "entidad", "entidad_id", "datos_anteriores", "datos_nuevos"]

        self.model = BaseTableModel(headers, keys, parent=self)
        self.table = QTableView()
        self.table.setModel(self.model)
        self.table.setSelectionBehavior(QTableView.SelectRows)
        self.table.setSelectionMode(QTableView.SingleSelection)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table.setColumnWidth(0, 60)
        self.table.setColumnWidth(1, 140)
        self.table.setColumnWidth(2, 110)
        self.table.setColumnWidth(3, 100)
        self.table.setColumnWidth(4, 120)
        self.table.setColumnWidth(5, 90)
        self.table.horizontalHeader().setSectionResizeMode(6, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(7, QHeaderView.Stretch)
        layout.addWidget(self.table)

        self.recargar()

    def recargar(self, text: str = "") -> None:
        logs = self.records_service.listar_auditoria(text)
        self.model.set_data(logs)

    def _on_search(self, text: str) -> None:
        self.recargar(text)

    def _exportar_excel(self) -> None:
        filas = self.model.get_all_data()
        if not filas:
            QMessageBox.information(self, "Exportar", "No hay registros de auditoría para exportar.")
            return

        from datetime import datetime
        now_str = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_name = f"Auditoria_RENO_{now_str}.xlsx"
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Guardar Reporte de Auditoría",
            default_name,
            "Libro de Excel (*.xlsx)",
        )
        if not file_path:
            return

        try:
            filtro = self.txt_filter.text().strip()
            self.excel_service.export_auditoria(file_path, filas, filtro_aplicado=filtro)
            QMessageBox.information(self, "Éxito", f"Reporte de auditoría exportado exitosamente a:\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al exportar reporte de auditoría:\n{e}")


class AdminAuditoriaDialog(BaseMaximizedDialog):
    """Diálogo visor de logs de auditoría de modificaciones y accesos."""

    def __init__(self, records_service: RecordsService, excel_service=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Trazabilidad y Auditoría del Sistema")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)

        self.widget = AdminAuditoriaWidget(records_service, excel_service=excel_service, parent=self)
        layout.addWidget(self.widget)

        btn_cerrar = QPushButton("Cerrar")
        btn_cerrar.setObjectName("btnSecondary")
        btn_cerrar.clicked.connect(self.accept)
        layout.addWidget(btn_cerrar, alignment=Qt.AlignRight)

    def recargar(self, text: str = "") -> None:
        self.widget.recargar(text)
