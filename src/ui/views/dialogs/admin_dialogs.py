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

from typing import Any, Dict, List, Optional
from datetime import datetime

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
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

from src.core.logger import get_logger
from src.domain.enums import PERMISOS_BASE
from src.services.auth_service import AuthService
from src.services.backup_service import BackupService
from src.services.excel_service import ExcelService
from src.services.records_service import RecordsService
from src.ui.models.base_table_model import BaseTableModel
from src.ui.models.multi_filter_proxy import MultiFilterProxyModel

logger = get_logger("AdminDialogs")


class BaseMaximizedDialog(QDialog):
    """Clase base para garantizar que todos los modales principales de administración se abran maximizados."""

    def showEvent(self, event) -> None:
        super().showEvent(event)
        self.showMaximized()


# ---------------------------------------------------------------------------
# 1. Administración de Usuarios
# ---------------------------------------------------------------------------

class AdminUsuariosDialog(BaseMaximizedDialog):
    """Gestión integral de cuentas de usuario, roles y claves."""

    def __init__(self, auth_service: AuthService, parent=None):
        super().__init__(parent)
        self.auth_service = auth_service
        self.setWindowTitle("Administración de Usuarios")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)

        # Barra de acciones
        bar = QHBoxLayout()
        self.btn_nuevo = QPushButton("➕ Nuevo Usuario")
        self.btn_nuevo.clicked.connect(self._nuevo_usuario)
        bar.addWidget(self.btn_nuevo)

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
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        layout.addWidget(self.table)

        btn_cerrar = QPushButton("Cerrar")
        btn_cerrar.setObjectName("btnSecondary")
        btn_cerrar.clicked.connect(self.accept)
        layout.addWidget(btn_cerrar, alignment=Qt.AlignRight)

        self.recargar()

    def recargar(self) -> None:
        usuarios = self.auth_service.listar_usuarios()
        for u in usuarios:
            u["activo_str"] = "Activo" if u.get("activo") == 1 else "Inactivo"
        self.model.set_data(usuarios)

    def _obtener_seleccionado(self) -> Optional[Dict[str, Any]]:
        indexes = self.table.selectionModel().selectedRows()
        if not indexes:
            return None
        return self.model.get_row(indexes[0].row())

    def _nuevo_usuario(self) -> None:
        roles_disponibles = [r["nombre"] for r in self.auth_service.listar_roles()]
        dlg = QDialog(self)
        dlg.setWindowTitle("Nuevo Usuario")
        dlg.setFixedSize(380, 320)
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
        btn_ok = QPushButton("Crear")
        btn_ok.clicked.connect(dlg.accept)
        btn_box.addWidget(btn_ok)
        form.addRow(btn_box)

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

    def _toggle_activo(self) -> None:
        u = self._obtener_seleccionado()
        if not u:
            return
        self.auth_service.cambiar_estado_usuario(u["id"])
        self.recargar()

    def _cambiar_clave(self) -> None:
        u = self._obtener_seleccionado()
        if not u:
            return
        dlg = QDialog(self)
        dlg.setWindowTitle(f"Restablecer Clave - {u['username']}")
        dlg.setFixedSize(320, 150)
        lay = QFormLayout(dlg)
        txt = QLineEdit()
        txt.setEchoMode(QLineEdit.Password)
        lay.addRow("Nueva clave:", txt)
        btn = QPushButton("Guardar")
        btn.clicked.connect(dlg.accept)
        lay.addRow(btn)

        if dlg.exec():
            val = txt.text()
            if val:
                self.auth_service.cambiar_password(u["id"], val)
                QMessageBox.information(self, "Éxito", "Contraseña actualizada.")


# ---------------------------------------------------------------------------
# 2. Roles y Permisos
# ---------------------------------------------------------------------------

class AdminRolesDialog(BaseMaximizedDialog):
    """Gestión de roles y asignación granular de permisos agrupados por módulo."""

    def __init__(self, auth_service: AuthService, parent=None):
        super().__init__(parent)
        self.auth_service = auth_service
        self.setWindowTitle("Administración de Roles y Permisos")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)

        splitter = QSplitter(Qt.Horizontal)

        # Panel izquierdo: Lista de roles
        left_widget = QWidget()
        lay_left = QVBoxLayout(left_widget)
        lay_left.setContentsMargins(0, 0, 0, 0)

        lay_left.addWidget(QLabel("<b>Roles Registrados:</b>"))
        self.tabla_roles = QTableWidget(0, 2)
        self.tabla_roles.setHorizontalHeaderLabels(["ID", "NOMBRE DEL ROL"])
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
        lay_right.setContentsMargins(0, 0, 0, 0)

        self.lbl_rol_activo = QLabel("<b>Seleccione un rol para editar sus permisos</b>")
        self.lbl_rol_activo.setStyleSheet("color: #1F497D; font-size: 11pt;")
        lay_right.addWidget(self.lbl_rol_activo)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        self.permisos_container = QWidget()
        self.permisos_layout = QVBoxLayout(self.permisos_container)
        scroll.setWidget(self.permisos_container)
        lay_right.addWidget(scroll)

        self.btn_guardar_permisos = QPushButton("💾 Guardar Permisos del Rol")
        self.btn_guardar_permisos.setEnabled(False)
        self.btn_guardar_permisos.clicked.connect(self._guardar_permisos)
        lay_right.addWidget(self.btn_guardar_permisos)

        splitter.addWidget(right_widget)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 2)
        layout.addWidget(splitter)

        btn_cerrar = QPushButton("Cerrar")
        btn_cerrar.setObjectName("btnSecondary")
        btn_cerrar.clicked.connect(self.accept)
        layout.addWidget(btn_cerrar, alignment=Qt.AlignRight)

        self._checkboxes_permisos: Dict[str, QCheckBox] = {}
        self.rol_seleccionado_id: Optional[int] = None

        self._cargar_permisos_ui()
        self.recargar_roles()

    def _cargar_permisos_ui(self) -> None:
        grupos: Dict[str, List[str]] = {
            "Novedades Operativas": [
                "novedades.ver", "novedades.crear", "novedades.editar", "novedades.eliminar",
                "novedades.importar", "novedades.exportar",
            ],
            "Cambios de Turno (Permutas)": [
                "cambios_turno.ver", "cambios_turno.crear", "cambios_turno.editar", "cambios_turno.eliminar",
                "cambios_turno.importar", "cambios_turno.exportar",
            ],
            "Personal de Ciclo y Reemplazos": [
                "reemplazos.ver", "reemplazos.asignar",
            ],
            "Catálogos y Empleados": [
                "empleados.administrar", "empleados.importar", "dotaciones.administrar",
                "tiposNovedad.administrar", "personalEstacion.ver", "personalEstacion.crear",
                "personalEstacion.editar", "personalEstacion.importar", "personalEstacion.exportar",
            ],
            "Administración General y Seguridad": [
                "usuarios.administrar", "roles.administrar", "auditoria.ver", "sesion.configurar",
                "registros.recuperar", "registros.modificar", "backup.gestionar",
                "destinatarios_informe.administrar",
            ],
        }

        for grupo_nombre, perm_list in grupos.items():
            grp = QGroupBox(grupo_nombre)
            grid = QGridLayout(grp)
            col = 0
            row = 0
            for code in perm_list:
                chk = QCheckBox(code)
                self._checkboxes_permisos[code] = chk
                grid.addWidget(chk, row, col)
                col += 1
                if col > 1:
                    col = 0
                    row += 1
            self.permisos_layout.addWidget(grp)

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
            return

        row = sel[0].row()
        r_id = int(self.tabla_roles.item(row, 0).text())
        r_nom = self.tabla_roles.item(row, 1).text()

        self.rol_seleccionado_id = r_id
        self.lbl_rol_activo.setText(f"<b>Permisos asignados al rol: {r_nom}</b>")
        self.btn_guardar_permisos.setEnabled(True)

        permisos_actuales = set(self.auth_service.permisos_de_rol(r_id))
        for code, chk in self._checkboxes_permisos.items():
            chk.setChecked(code in permisos_actuales)

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

        bar.addSeparator()

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

class AdminCatalogosDialog(BaseMaximizedDialog):
    """Gestión de Tipos de Novedad, Dotaciones y Personal de Estación."""

    def __init__(self, records_service: RecordsService, excel_service: ExcelService, parent=None):
        super().__init__(parent)
        self.records_service = records_service
        self.excel_service = excel_service
        self.setWindowTitle("Catálogos Operativos del Sistema")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)

        tabs = QTabWidget()

        # Tab 1: Tipos de novedad
        tab_tipos = QWidget()
        lay_t = QVBoxLayout(tab_tipos)
        h_t = QHBoxLayout()
        btn_nt = QPushButton("➕ Nuevo Tipo")
        btn_nt.clicked.connect(self._nuevo_tipo)
        h_t.addWidget(btn_nt)
        h_t.addStretch()
        lay_t.addLayout(h_t)

        self.table_tipos = QTableWidget(0, 5)
        self.table_tipos.setHorizontalHeaderLabels(["ID", "NOMBRE", "DÍAS AUTORIZADOS", "EXPORTA DESCUBIERTO", "ACTIVO"])
        self.table_tipos.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        lay_t.addWidget(self.table_tipos)
        tabs.addTab(tab_tipos, "Tipos de Novedad")

        # Tab 2: Dotaciones
        tab_dots = QWidget()
        lay_d = QVBoxLayout(tab_dots)
        h_d = QHBoxLayout()
        btn_nd = QPushButton("➕ Nueva Dotación")
        btn_nd.clicked.connect(self._nueva_dotacion)
        h_d.addWidget(btn_nd)
        h_d.addStretch()
        lay_d.addLayout(h_d)

        self.table_dots = QTableWidget(0, 3)
        self.table_dots.setHorizontalHeaderLabels(["ID", "DOTACIÓN", "ACTIVO"])
        self.table_dots.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        lay_d.addWidget(self.table_dots)
        tabs.addTab(tab_dots, "Dotaciones")

        # Tab 3: Personal de Estación
        tab_pe = QWidget()
        lay_pe = QVBoxLayout(tab_pe)
        h_pe = QHBoxLayout()
        btn_npe = QPushButton("➕ Nuevo Personal de Estación")
        btn_npe.clicked.connect(self._nuevo_personal_estacion)
        h_pe.addWidget(btn_npe)
        h_pe.addStretch()
        lay_pe.addLayout(h_pe)

        self.table_pe = QTableWidget(0, 4)
        self.table_pe.setHorizontalHeaderLabels(["ID", "NOMBRE Y APELLIDO", "DOTACIÓN", "ACTIVO"])
        self.table_pe.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        lay_pe.addWidget(self.table_pe)
        tabs.addTab(tab_pe, "Personal de Estación")

        layout.addWidget(tabs)

        btn_cerrar = QPushButton("Cerrar")
        btn_cerrar.setObjectName("btnSecondary")
        btn_cerrar.clicked.connect(self.accept)
        layout.addWidget(btn_cerrar, alignment=Qt.AlignRight)

        self.recargar_todo()

    def recargar_todo(self) -> None:
        # Tipos
        tipos = self.records_service.listar_tipos(incluir_inactivos=True)
        self.table_tipos.setRowCount(len(tipos))
        for r, t in enumerate(tipos):
            self.table_tipos.setItem(r, 0, QTableWidgetItem(str(t["id"])))
            self.table_tipos.setItem(r, 1, QTableWidgetItem(str(t["nombre"])))
            self.table_tipos.setItem(r, 2, QTableWidgetItem(str(t.get("dias", 0))))
            self.table_tipos.setItem(r, 3, QTableWidgetItem("Sí" if t.get("exportar_descubierto") else "No"))
            self.table_tipos.setItem(r, 4, QTableWidgetItem("Activo" if t.get("activo") else "Inactivo"))

        # Dotaciones
        dots = self.records_service.listar_dotaciones(incluir_inactivos=True)
        self.table_dots.setRowCount(len(dots))
        for r, d in enumerate(dots):
            self.table_dots.setItem(r, 0, QTableWidgetItem(str(d["id"])))
            self.table_dots.setItem(r, 1, QTableWidgetItem(str(d["nombre"])))
            self.table_dots.setItem(r, 2, QTableWidgetItem("Activo" if d.get("activo") else "Inactivo"))

        # Personal estación
        pes = self.records_service.listar_personal_estacion(incluir_inactivos=True)
        self.table_pe.setRowCount(len(pes))
        for r, p in enumerate(pes):
            self.table_pe.setItem(r, 0, QTableWidgetItem(str(p["id"])))
            self.table_pe.setItem(r, 1, QTableWidgetItem(str(p["nombre"])))
            self.table_pe.setItem(r, 2, QTableWidgetItem(str(p.get("dotacion") or "-")))
            self.table_pe.setItem(r, 3, QTableWidgetItem("Activo" if p.get("activo") else "Inactivo"))

    def _nuevo_tipo(self) -> None:
        dlg = QDialog(self)
        dlg.setWindowTitle("Nuevo Tipo de Novedad")
        dlg.setFixedSize(360, 200)
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
                self.records_service.crear_tipo(n, dias=spin.value(), exportar_descubierto=1 if chk.isChecked() else 0)
                self.recargar_todo()

    def _nueva_dotacion(self) -> None:
        dlg = QDialog(self)
        dlg.setWindowTitle("Nueva Dotación")
        dlg.setFixedSize(320, 130)
        lay = QFormLayout(dlg)
        txt = QLineEdit()
        lay.addRow("Nombre:", txt)
        btn = QPushButton("Guardar")
        btn.clicked.connect(dlg.accept)
        lay.addRow(btn)

        if dlg.exec():
            n = txt.text().strip()
            if n:
                self.records_service.crear_dotacion(n)
                self.recargar_todo()

    def _nuevo_personal_estacion(self) -> None:
        dots = [d["nombre"] for d in self.records_service.listar_dotaciones(incluir_inactivos=False)]
        dlg = QDialog(self)
        dlg.setWindowTitle("Nuevo Personal de Estación")
        dlg.setFixedSize(360, 160)
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
            d = combo.currentText().strip() or None
            if n:
                self.records_service.crear_personal_estacion(n, dotacion=d)
                self.recargar_todo()


# ---------------------------------------------------------------------------
# 5. Configuración General (Tiempos, Contraseñas, Toasts)
# ---------------------------------------------------------------------------

class AdminConfigDialog(BaseMaximizedDialog):
    """Configuración de parámetros operativos y de seguridad."""

    def __init__(self, db, parent=None):
        super().__init__(parent)
        self.db = db
        self.setWindowTitle("Configuración del Sistema")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)

        form = QFormLayout()

        # Ventanas
        self.spin_edit = QSpinBox()
        self.spin_edit.setRange(1, 720)
        form.addRow("Ventana para editar registros (horas):", self.spin_edit)

        self.spin_elim = QSpinBox()
        self.spin_elim.setRange(1, 720)
        form.addRow("Ventana para eliminar registros (horas):", self.spin_elim)

        # Claves
        self.spin_exp = QSpinBox()
        self.spin_exp.setRange(0, 365)
        form.addRow("Expiración obligatoria de contraseñas (días, 0=desactivada):", self.spin_exp)

        self.txt_dom = QLineEdit()
        form.addRow("Dominio de correo institucional para recuperación:", self.txt_dom)

        # Notificaciones
        self.chk_notif = QCheckBox("Activar notificaciones toast de nuevos registros")
        form.addRow(self.chk_notif)

        self.spin_toast = QSpinBox()
        self.spin_toast.setRange(2, 30)
        form.addRow("Duración de notificaciones en pantalla (segundos):", self.spin_toast)

        layout.addLayout(form)

        # Guardar
        btn_box = QHBoxLayout()
        btn_box.addStretch()
        btn_cancel = QPushButton("Cancelar")
        btn_cancel.setObjectName("btnSecondary")
        btn_cancel.clicked.connect(self.reject)
        btn_save = QPushButton("Guardar Configuración")
        btn_save.clicked.connect(self._guardar)
        btn_box.addWidget(btn_cancel)
        btn_box.addWidget(btn_save)
        layout.addLayout(btn_box)

        self._cargar()

    def _cargar(self) -> None:
        with self.db.read_connection() as conn:
            rows = dict(conn.execute("SELECT clave, valor FROM configuracion").fetchall())
        self.spin_edit.setValue(int(rows.get("editar_horas", 24)))
        self.spin_elim.setValue(int(rows.get("eliminar_horas", 72)))
        self.spin_exp.setValue(int(rows.get("clave_expiracion_dias", 90)))
        self.txt_dom.setText(str(rows.get("correo_dominio", "@trenesargentinos.gob.ar")))
        self.chk_notif.setChecked(bool(int(rows.get("notificaciones_activo", 1))))
        self.spin_toast.setValue(int(rows.get("toast_duracion", 6)))

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
        QMessageBox.information(self, "Éxito", "Configuración guardada exitosamente.")
        self.accept()


# ---------------------------------------------------------------------------
# 6. Copias de Seguridad (Backups)
# ---------------------------------------------------------------------------

class AdminBackupsDialog(BaseMaximizedDialog):
    """Gestor de copias de seguridad de la base SQLite."""

    def __init__(self, backup_service: BackupService, parent=None):
        super().__init__(parent)
        self.backup_service = backup_service
        self.setWindowTitle("Copias de Seguridad (Backups)")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)

        bar = QHBoxLayout()
        self.btn_crear = QPushButton("💾 Crear Backup Ahora")
        self.btn_crear.clicked.connect(self._crear_backup)
        bar.addWidget(self.btn_crear)

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
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        layout.addWidget(self.table)

        btn_cerrar = QPushButton("Cerrar")
        btn_cerrar.setObjectName("btnSecondary")
        btn_cerrar.clicked.connect(self.accept)
        layout.addWidget(btn_cerrar, alignment=Qt.AlignRight)

        self.recargar()

    def recargar(self) -> None:
        items = self.backup_service.listar_backups()
        datos = [
            {"nombre": n, "tamano_str": f"{sz / 1024:.1f} KB", "fecha": f}
            for n, sz, f in items
        ]
        self.model.set_data(datos)

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


# ---------------------------------------------------------------------------
# 7. Registros Eliminados (Papelera)
# ---------------------------------------------------------------------------

class AdminEliminadosDialog(BaseMaximizedDialog):
    """Papelera de recuperación y borrado permanente de registros."""

    def __init__(self, records_service: RecordsService, usuario_actual: Optional[Dict[str, Any]] = None, parent=None):
        super().__init__(parent)
        self.records_service = records_service
        self.usuario_actual = usuario_actual or {}
        self.setWindowTitle("Registros Eliminados (Papelera)")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)

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
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(6, QHeaderView.Stretch)
        layout.addWidget(self.table)

        btn_cerrar = QPushButton("Cerrar")
        btn_cerrar.setObjectName("btnSecondary")
        btn_cerrar.clicked.connect(self.accept)
        layout.addWidget(btn_cerrar, alignment=Qt.AlignRight)

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


# ---------------------------------------------------------------------------
# 8. Auditoría del Sistema
# ---------------------------------------------------------------------------

class AdminAuditoriaDialog(BaseMaximizedDialog):
    """Visor de logs de auditoría de modificaciones y accesos."""

    def __init__(self, records_service: RecordsService, parent=None):
        super().__init__(parent)
        self.records_service = records_service
        self.setWindowTitle("Trazabilidad y Auditoría del Sistema")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)

        bar = QHBoxLayout()
        bar.addWidget(QLabel("Filtrar auditoría:"))
        self.txt_filter = QLineEdit()
        self.txt_filter.setPlaceholderText("Usuario, acción, entidad...")
        self.txt_filter.textChanged.connect(self._on_search)
        bar.addWidget(self.txt_filter)
        bar.addStretch()
        layout.addLayout(bar)

        headers = ["ID", "FECHA Y HORA", "USUARIO", "ACCIÓN", "ENTIDAD", "ID REGISTRO", "DATOS ANTERIORES", "DATOS NUEVOS"]
        keys = ["id", "creado_en", "usuario", "accion", "entidad", "entidad_id", "datos_anteriores", "datos_nuevos"]

        self.model = BaseTableModel(headers, keys, parent=self)
        self.table = QTableView()
        self.table.setModel(self.model)
        self.table.setSelectionBehavior(QTableView.SelectRows)
        self.table.setSelectionMode(QTableView.SingleSelection)
        self.table.horizontalHeader().setSectionResizeMode(6, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(7, QHeaderView.Stretch)
        layout.addWidget(self.table)

        btn_cerrar = QPushButton("Cerrar")
        btn_cerrar.setObjectName("btnSecondary")
        btn_cerrar.clicked.connect(self.accept)
        layout.addWidget(btn_cerrar, alignment=Qt.AlignRight)

        self.recargar()

    def recargar(self, text: str = "") -> None:
        logs = self.records_service.listar_auditoria(text)
        self.model.set_data(logs)

    def _on_search(self, text: str) -> None:
        self.recargar(text)
