"""Vista principal para la Gestión y Consulta de Empleados (BASE).

Permite visualizar la base de personal, filtrar por legajo, nombre, dotación y especialidad,
dar de alta nuevos agentes, editar datos de turno/franco/jornada CCT, importar desde Excel
y alternar el estado activo/inactivo (soft-delete).
"""

from typing import Any, Dict, List, Optional

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDialog,
    QFileDialog,
    QFormLayout,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QTableView,
    QVBoxLayout,
    QWidget,
)

from src.core.logger import get_logger
from src.services.excel_service import ExcelService
from src.services.records_service import RecordsService
from src.ui.models.base_table_model import BaseTableModel
from src.ui.models.multi_filter_proxy import MultiFilterProxyModel

logger = get_logger("EmpleadosView")


class EmpleadosView(QWidget):
    """Pestaña principal de visualización y administración de la base de empleados."""

    def __init__(
        self,
        records_service: RecordsService,
        excel_service: ExcelService,
        usuario_actual: Optional[Dict[str, Any]] = None,
        parent=None,
    ):
        super().__init__(parent)
        self.records_service = records_service
        self.excel_service = excel_service
        self.usuario_actual = usuario_actual or {}

        self._init_ui()
        self.recargar()

    def _init_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        # 1. Barra de herramientas superior
        bar_frame = QFrame()
        bar_frame.setObjectName("filterBar")
        bar = QHBoxLayout(bar_frame)
        bar.setContentsMargins(8, 6, 8, 6)
        bar.setSpacing(10)

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

        sep1 = QFrame()
        sep1.setFrameShape(QFrame.VLine)
        sep1.setFrameShadow(QFrame.Sunken)
        bar.addWidget(sep1)

        self.btn_importar = QPushButton("📥 Importar Excel")
        self.btn_importar.setObjectName("btnSecondary")
        self.btn_importar.clicked.connect(self._importar_excel)
        bar.addWidget(self.btn_importar)

        self.btn_exportar = QPushButton("📊 Exportar Excel")
        self.btn_exportar.setObjectName("btnSecondary")
        self.btn_exportar.clicked.connect(self._exportar_excel)
        bar.addWidget(self.btn_exportar)

        bar.addStretch()

        # Filtro de Dotación
        bar.addWidget(QLabel("Dotación:"))
        self.combo_dotacion = QComboBox()
        self.combo_dotacion.addItem("Todas")
        dots = self.records_service.listar_dotaciones(incluir_inactivos=False)
        for d in dots:
            self.combo_dotacion.addItem(d["nombre"])
        self.combo_dotacion.currentIndexChanged.connect(self._on_filtro_changed)
        bar.addWidget(self.combo_dotacion)

        # Filtro de Especialidad
        bar.addWidget(QLabel("Especialidad:"))
        self.combo_esp = QComboBox()
        self.combo_esp.addItem("Todas")
        for esp in self.records_service.listar_especialidades_empleados():
            self.combo_esp.addItem(esp)
        self.combo_esp.currentIndexChanged.connect(self._on_filtro_changed)
        bar.addWidget(self.combo_esp)

        # Filtro de Búsqueda Rápida
        bar.addWidget(QLabel("Buscar:"))
        self.txt_search = QLineEdit()
        self.txt_search.setPlaceholderText("Legajo o apellido...")
        self.txt_search.textChanged.connect(self._on_search)
        self.txt_search.setFixedWidth(160)
        bar.addWidget(self.txt_search)

        # Checkbox Inactivos
        self.chk_inactivos = QCheckBox("Ver Inactivos")
        self.chk_inactivos.setChecked(True)
        self.chk_inactivos.toggled.connect(self.recargar)
        bar.addWidget(self.chk_inactivos)

        layout.addWidget(bar_frame)

        # 2. Tabla de Datos
        headers = [
            "ID", "LEGAJO", "APELLIDOS Y NOMBRES", "ESPECIALIDAD", "DOTACIÓN",
            "TURNOS", "FRANCO", "HORARIO FRANCO", "JORNADA HS", "ESTADO",
        ]
        keys = [
            "id", "legajo", "apellidos_nombres", "especialidad", "dotacion",
            "turnos", "franco", "franco_horario", "jornada_horas", "estado_str",
        ]

        self.model = BaseTableModel(headers, keys, parent=self)
        self.proxy = MultiFilterProxyModel(self)
        self.proxy.setSourceModel(self.model)

        self.table = QTableView()
        self.table.setModel(self.proxy)
        self.table.setSelectionBehavior(QTableView.SelectRows)
        self.table.setSelectionMode(QTableView.SingleSelection)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table.setColumnWidth(0, 50)   # ID
        self.table.setColumnWidth(1, 80)   # LEGAJO
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)  # APELLIDOS Y NOMBRES
        self.table.setColumnWidth(3, 180)  # ESPECIALIDAD (espacio completo sin truncamiento)
        self.table.setColumnWidth(4, 85)   # DOTACIÓN
        self.table.setColumnWidth(5, 95)   # TURNOS
        self.table.setColumnWidth(6, 95)   # FRANCO
        self.table.setColumnWidth(7, 125)  # HORARIO FRANCO
        self.table.setColumnWidth(8, 85)   # JORNADA HS
        self.table.setColumnWidth(9, 90)   # ESTADO
        self.table.doubleClicked.connect(self._on_double_click)
        layout.addWidget(self.table)

        # Pie con conteo
        self.lbl_conteo = QLabel("Total: 0 empleados")
        self.lbl_conteo.setStyleSheet("color: #6B7280; font-size: 8.5pt;")
        layout.addWidget(self.lbl_conteo)

    def recargar(self) -> None:
        incluir_inactivos = self.chk_inactivos.isChecked()
        empleados = self.records_service.listar_empleados(incluir_inactivos=incluir_inactivos)
        for e in empleados:
            e["estado_str"] = "Activo" if e.get("activo") == 1 else "Inactivo"
        self.model.set_data(empleados)
        self._on_filtro_changed()

    def _on_search(self, text: str) -> None:
        self.proxy.set_global_search(text)
        self._actualizar_conteo()

    def _on_filtro_changed(self) -> None:
        dot = self.combo_dotacion.currentText()
        esp = self.combo_esp.currentText()
        self.proxy.set_column_filter(4, "" if dot == "Todas" else dot)
        self.proxy.set_column_filter(3, "" if esp == "Todas" else esp)
        self._actualizar_conteo()

    def _actualizar_conteo(self) -> None:
        total = self.proxy.rowCount()
        self.lbl_conteo.setText(f"Total mostrados: {total} agentes")

    def _obtener_seleccionado(self) -> Optional[Dict[str, Any]]:
        sel = self.table.selectionModel().selectedRows()
        if not sel:
            return None
        source_idx = self.proxy.mapToSource(sel[0])
        return self.model.get_row(source_idx.row())

    def _on_double_click(self, index) -> None:
        self._modificar_empleado()

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
        dlg.resize(460, 420)
        lay = QFormLayout(dlg)

        txt_leg = QLineEdit(str(emp_existente["legajo"]) if emp_existente else "")
        txt_nom = QLineEdit(str(emp_existente["apellidos_nombres"]) if emp_existente else "")

        txt_esp = QComboBox()
        txt_esp.setEditable(True)
        txt_esp.addItems(self.records_service.listar_especialidades_empleados())
        if emp_existente and emp_existente.get("especialidad"):
            txt_esp.setCurrentText(str(emp_existente["especialidad"]))

        combo_dot = QComboBox()
        dots = [d["nombre"] for d in self.records_service.listar_dotaciones(incluir_inactivos=False)]
        combo_dot.addItems(dots)
        if emp_existente and emp_existente.get("dotacion"):
            combo_dot.setCurrentText(str(emp_existente["dotacion"]))

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

        btn_box = QHBoxLayout()
        btn_cancel = QPushButton("Cancelar")
        btn_cancel.setObjectName("btnSecondary")
        btn_cancel.clicked.connect(dlg.reject)
        btn_ok = QPushButton("Guardar")
        btn_ok.clicked.connect(dlg.accept)
        btn_box.addWidget(btn_cancel)
        btn_box.addWidget(btn_ok)
        lay.addRow(btn_box)

        if dlg.exec():
            leg_val = txt_leg.text().strip()
            nom_val = txt_nom.text().strip()
            if not leg_val or not nom_val:
                QMessageBox.warning(self, "Validación", "Legajo y Apellidos/Nombres son obligatorios.")
                return

            data = {
                "legajo": leg_val,
                "apellidos_nombres": nom_val,
                "especialidad": txt_esp.currentText().strip(),
                "dotacion": combo_dot.currentText().strip(),
                "turnos": txt_turno.text().strip(),
                "franco": txt_franco.text().strip(),
                "franco_horario": txt_franco_h.text().strip(),
                "jornada_horas": spin_jornada.value(),
            }
            try:
                u_id = self.usuario_actual.get("id")
                u_win = self.usuario_actual.get("username")
                if emp_existente:
                    self.records_service.actualizar_empleado(
                        emp_existente["id"],
                        data,
                        user_id=u_id,
                        windows_user=u_win,
                    )
                else:
                    self.records_service.crear_empleado(
                        data,
                        user_id=u_id,
                        windows_user=u_win,
                    )
                self.recargar()
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Fallo al guardar empleado:\n{e}")

    def _toggle_estado(self) -> None:
        emp = self._obtener_seleccionado()
        if not emp:
            QMessageBox.information(self, "Activar / Desactivar", "Seleccione un empleado.")
            return

        nuevo_activo = 0 if emp.get("activo") == 1 else 1
        accion = "activar" if nuevo_activo == 1 else "desactivar"
        try:
            self.records_service.toggle_empleado_activo(
                emp["id"],
                nuevo_activo,
                user_id=self.usuario_actual.get("id"),
                windows_user=self.usuario_actual.get("username"),
            )
            self.recargar()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo {accion} el empleado:\n{e}")

    def _importar_excel(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Importar Empleados desde Excel", "", "Excel (*.xlsx *.xlsm)")
        if not path:
            return
        try:
            cnt = self.excel_service.import_base_empleados(path)
            self.recargar()
            QMessageBox.information(self, "Importación", f"Se importaron / actualizaron {cnt} empleados con éxito.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Fallo al importar empleados:\n{e}")

    def _exportar_excel(self) -> None:
        path, _ = QFileDialog.getSaveFileName(self, "Exportar Base de Empleados", "Empleados_BASE.xlsx", "Excel (*.xlsx)")
        if not path:
            return
        try:
            self.excel_service.export_database(path, tables=("BASE",))
            QMessageBox.information(self, "Exportación", f"Base de empleados exportada exitosamente en:\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Fallo al exportar:\n{e}")
