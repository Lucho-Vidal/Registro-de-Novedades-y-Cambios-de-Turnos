"""Vista especializada para el seguimiento y gestión de Personal de Ciclo."""

from typing import Any, Dict, List, Optional
from datetime import datetime

from PySide6.QtCore import QDate, Qt
from PySide6.QtWidgets import (
    QComboBox,
    QDateEdit,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMessageBox,
    QPushButton,
    QSplitter,
    QTableView,
    QVBoxLayout,
    QWidget,
)

from src.core.logger import get_logger
from src.services.excel_service import ExcelService
from src.services.pdf_service import PDFService
from src.services.records_service import RecordsService
from src.ui.models.base_table_model import BaseTableModel
from src.ui.views.dialogs.calendario_ciclo_dialog import CalendarioCicloDialog

logger = get_logger("PersonalCicloView")


class PersonalCicloView(QWidget):
    """Vista dividida en dos tablas: Personal Libre (arriba) y Personal Asignado (abajo)."""

    def __init__(
        self,
        records_service: RecordsService,
        pdf_service: PDFService,
        excel_service: ExcelService,
        usuario_actual: Optional[Dict[str, Any]] = None,
        parent=None,
    ):
        super().__init__(parent)
        self.records_service = records_service
        self.pdf_service = pdf_service
        self.excel_service = excel_service
        self.usuario_actual = usuario_actual or {}

        self._init_ui()
        self.consultar()

    def _init_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        # 1. Barra superior de filtros
        bar_frame = QFrame()
        bar_frame.setStyleSheet("background-color: #FFFFFF; border: 1px solid #D1D5DB; border-radius: 4px; padding: 4px;")
        bar_layout = QHBoxLayout(bar_frame)
        bar_layout.setContentsMargins(8, 4, 8, 4)
        bar_layout.setSpacing(10)

        bar_layout.addWidget(QLabel("Fecha de Consulta:"))
        self.date_picker = QDateEdit()
        self.date_picker.setCalendarPopup(True)
        self.date_picker.setDisplayFormat("dd/MM/yyyy")
        self.date_picker.setDate(QDate.currentDate())
        bar_layout.addWidget(self.date_picker)

        bar_layout.addWidget(QLabel("Dotación:"))
        self.combo_dotacion = QComboBox()
        self.combo_dotacion.addItem("Todas")
        dots = self.records_service.listar_dotaciones(incluir_inactivos=False)
        for d in dots:
            self.combo_dotacion.addItem(d["nombre"])
        bar_layout.addWidget(self.combo_dotacion)

        bar_layout.addWidget(QLabel("Especialidad:"))
        self.combo_esp = QComboBox()
        self.combo_esp.addItems(["Todas", "CONDUCTOR", "GUARDATREN"])
        bar_layout.addWidget(self.combo_esp)

        self.btn_consultar = QPushButton("🔍 Actualizar")
        self.btn_consultar.clicked.connect(self.consultar)
        bar_layout.addWidget(self.btn_consultar)

        bar_layout.addStretch()

        self.btn_ver_calendario = QPushButton("📅 Ver Calendario Mensual")
        self.btn_ver_calendario.setObjectName("btnSecondary")
        self.btn_ver_calendario.clicked.connect(self._abrir_calendario_seleccionado)
        bar_layout.addWidget(self.btn_ver_calendario)

        layout.addWidget(bar_frame)

        # 2. Splitter vertical con las dos tablas
        splitter = QSplitter(Qt.Vertical)

        # Tabla Superior: Personal Libre / Sin Relevo
        grp_libres = QGroupBox("Personal de Ciclo DISPONIBLE (Sin relevo asignado en la fecha)")
        lay_libres = QVBoxLayout(grp_libres)
        headers_libres = ["LEGAJO", "APELLIDOS Y NOMBRES", "ESPECIALIDAD", "DOTACIÓN", "TURNO HABITUAL", "FRANCO", "HORARIO FRANCO"]
        keys_libres = ["legajo", "apellidos_nombres", "especialidad", "dotacion", "turnos", "franco", "franco_horario"]

        self.model_libres = BaseTableModel(headers_libres, keys_libres, parent=self)
        self.table_libres = QTableView()
        self.table_libres.setModel(self.model_libres)
        self.table_libres.setSelectionBehavior(QTableView.SelectRows)
        self.table_libres.setSelectionMode(QTableView.SingleSelection)
        self.table_libres.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table_libres.doubleClicked.connect(lambda idx: self._on_double_click(self.model_libres, idx))
        lay_libres.addWidget(self.table_libres)
        splitter.addWidget(grp_libres)

        # Tabla Inferior: Personal Ocupado / Con Relevo
        grp_ocupados = QGroupBox("Personal de Ciclo ASIGNADO (Realizando relevo en la fecha)")
        lay_ocupados = QVBoxLayout(grp_ocupados)
        headers_ocupados = [
            "LEGAJO", "APELLIDOS Y NOMBRES", "ESPECIALIDAD", "DOTACIÓN", "REEMPLAZA A (TITULAR)",
            "TURNO CUBIERTO", "MOTIVO / NOVEDAD", "DESDE", "HASTA",
        ]
        keys_ocupados = [
            "legajo", "apellidos_nombres", "especialidad", "dotacion", "reemplaza_a",
            "turno_titular", "tipo_novedad", "fecha_inicio_relevo", "fecha_fin_relevo",
        ]

        self.model_ocupados = BaseTableModel(headers_ocupados, keys_ocupados, parent=self)
        self.table_ocupados = QTableView()
        self.table_ocupados.setModel(self.model_ocupados)
        self.table_ocupados.setSelectionBehavior(QTableView.SelectRows)
        self.table_ocupados.setSelectionMode(QTableView.SingleSelection)
        self.table_ocupados.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table_ocupados.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        self.table_ocupados.doubleClicked.connect(lambda idx: self._on_double_click(self.model_ocupados, idx))
        lay_ocupados.addWidget(self.table_ocupados)
        splitter.addWidget(grp_ocupados)

        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 1)
        layout.addWidget(splitter)

        # 3. Pie con instrucciones
        lbl_hint = QLabel("💡 Haga doble clic sobre cualquier agente de ciclo para abrir su Calendario Mensual y consultar descansos CCT.")
        lbl_hint.setStyleSheet("color: #6B7280; font-size: 8.5pt;")
        layout.addWidget(lbl_hint)

    def consultar(self) -> None:
        f_str = self.date_picker.date().toString("dd/MM/yyyy")
        dot = self.combo_dotacion.currentText()
        esp = self.combo_esp.currentText()

        try:
            res = self.records_service.obtener_estado_personal_ciclo(f_str, dotacion=dot, especialidad=esp)
            self.model_libres.set_data(res.get("libres", []))
            self.model_ocupados.set_data(res.get("ocupados", []))
        except Exception as e:
            logger.error(f"Error al consultar personal de ciclo: {e}")

    def _on_double_click(self, model: BaseTableModel, index) -> None:
        emp = model.get_row(index.row())
        if emp:
            self._abrir_calendario(emp)

    def _abrir_calendario_seleccionado(self) -> None:
        # Chequear selección en tabla libres
        sel_l = self.table_libres.selectionModel().selectedRows()
        if sel_l:
            emp = self.model_libres.get_row(sel_l[0].row())
            if emp:
                self._abrir_calendario(emp)
                return

        # Chequear selección en tabla ocupados
        sel_o = self.table_ocupados.selectionModel().selectedRows()
        if sel_o:
            emp = self.model_ocupados.get_row(sel_o[0].row())
            if emp:
                self._abrir_calendario(emp)
                return

        QMessageBox.information(self, "Selección", "Seleccione un empleado en cualquiera de las dos tablas para ver su calendario.")

    def _abrir_calendario(self, emp: Dict[str, Any]) -> None:
        qdate = self.date_picker.date()
        dlg = CalendarioCicloDialog(
            empleado=emp,
            records_service=self.records_service,
            pdf_service=self.pdf_service,
            excel_service=self.excel_service,
            mes=qdate.month(),
            anio=qdate.year(),
            usuario_actual=self.usuario_actual,
            parent=self,
        )
        dlg.exec()
        self.consultar()
