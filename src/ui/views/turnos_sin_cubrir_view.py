"""Vista Principal: Turnos Operativos Sin Cubrir y Novedades Descubiertas.

Presenta la grilla detallada de turnos descubiertos y relevos para la fecha consultada,
con desglose del turno diario (ej: 401.1), horarios de toma y dejada calculados según
el itinerario y el detalle de qué turno estaba reemplazando si es personal de ciclo.
"""

from typing import Any, Dict, List, Optional
from datetime import datetime

from PySide6.QtCore import QDate, Qt
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QComboBox,
    QDateEdit,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QTableView,
    QVBoxLayout,
    QWidget,
)

from src.core.logger import get_logger
from src.core.settings import AppSettings
from src.services.excel_service import ExcelService
from src.services.records_service import RecordsService
from src.ui.models.base_table_model import BaseTableModel
from src.ui.models.multi_filter_proxy import MultiFilterProxyModel

logger = get_logger("TurnosSinCubrirView")


class TurnosSinCubrirView(QWidget):
    """Pestaña dedicada para el control y exportación de turnos sin cubrir y relevos."""

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
        self.settings = AppSettings()

        self._init_ui()
        self.consultar()

    def _init_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        # 1. Barra de Herramientas y Filtros
        bar_frame = QFrame()
        bar_frame.setObjectName("filterBar")
        bar = QHBoxLayout(bar_frame)
        bar.setContentsMargins(8, 6, 8, 6)
        bar.setSpacing(10)

        bar.addWidget(QLabel("Fecha de Consulta:"))
        self.date_picker = QDateEdit()
        self.date_picker.setCalendarPopup(True)
        self.date_picker.setDisplayFormat("dd/MM/yyyy")
        self.date_picker.setDate(QDate.currentDate())
        self.date_picker.dateChanged.connect(self.consultar)
        bar.addWidget(self.date_picker)

        bar.addWidget(QLabel("Dotación:"))
        self.combo_dot = QComboBox()
        self.combo_dot.addItem("Todas")
        dots = self.records_service.listar_dotaciones(incluir_inactivos=False)
        for d in dots:
            self.combo_dot.addItem(d["nombre"])
        self.combo_dot.currentIndexChanged.connect(self.consultar)
        bar.addWidget(self.combo_dot)

        self.btn_consultar = QPushButton("🔍 Actualizar")
        self.btn_consultar.clicked.connect(self.consultar)
        bar.addWidget(self.btn_consultar)

        self.btn_export = QPushButton("📊 Exportar Excel")
        self.btn_export.setObjectName("btnSecondary")
        self.btn_export.clicked.connect(self._exportar_excel)
        bar.addWidget(self.btn_export)

        bar.addStretch()

        bar.addWidget(QLabel("Buscar:"))
        self.txt_search = QLineEdit()
        self.txt_search.setPlaceholderText("Filtrar por agente o turno...")
        self.txt_search.textChanged.connect(self._on_search)
        self.txt_search.setFixedWidth(180)
        bar.addWidget(self.txt_search)

        layout.addWidget(bar_frame)

        # 2. Tabla Principal
        headers = [
            "LEGAJO", "APELLIDOS Y NOMBRES", "ESPECIALIDAD", "DOTACIÓN", "TURNO BASE",
            "FRANCO", "TURNO DÍA", "HORA TOMA", "HORA DEJA", "NOVEDAD",
            "DESDE", "HASTA", "CUBIERTO", "REEMPLAZO ASIGNADO", "TURNO QUE REEMPLAZA",
            "OBSERVACIONES",
        ]
        keys = [
            "legajo", "apellidos_nombres", "especialidad", "dotacion", "turnos",
            "franco", "turno_dia", "hora_toma", "hora_deja", "novedad",
            "fecha_inicio", "fecha_fin", "cubierto", "reemplazo", "turno_reemplazado",
            "observaciones",
        ]

        def conditional_colors(row_item: Dict[str, Any], col_idx: int):
            from src.core.theme import obtener_tema_visual
            cubierto = str(row_item.get("cubierto", "")).strip()
            turno_dia = str(row_item.get("turno_dia", "")).strip()
            es_oscuro = obtener_tema_visual() == "oscuro"

            if cubierto == "No":
                # Novedad descubierta / sin cubrir
                return (QColor("#3D1B1B"), QColor("#FCA5A5")) if es_oscuro else (QColor("#FFEBEE"), QColor("#B71C1C"))
            elif cubierto == "Sí":
                # Cubierto con reemplazo
                return (QColor("#132E1F"), QColor("#6EE7B7")) if es_oscuro else (QColor("#E8F5E9"), QColor("#1B5E20"))
            elif turno_dia == "FRANCO":
                return (QColor("#3D3214"), QColor("#FCD34D")) if es_oscuro else (QColor("#FFF9C4"), QColor("#B78103"))
            return None

        self.model = BaseTableModel(headers, keys, conditional_colors=conditional_colors, parent=self)
        self.proxy = MultiFilterProxyModel(self)
        self.proxy.setSourceModel(self.model)

        self.table = QTableView()
        self.table.setModel(self.proxy)
        self.table.setSelectionBehavior(QTableView.SelectRows)
        self.table.setSelectionMode(QTableView.SingleSelection)
        self.table.clicked.connect(self._on_table_clicked)
        self.table.doubleClicked.connect(self._on_table_double_clicked)
        self.table.setToolTip("Haga clic en la columna 'TURNO DÍA' o doble clic en la fila para ver el diagrama y vueltas completas.")

        # Configuración de anchos limpios e interactivos
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table.setColumnWidth(0, 75)   # Legajo
        self.table.setColumnWidth(1, 190)  # Apellidos y Nombres
        self.table.setColumnWidth(2, 140)  # Especialidad
        self.table.setColumnWidth(3, 95)   # Dotación
        self.table.setColumnWidth(4, 85)   # Turno Base
        self.table.setColumnWidth(5, 85)   # Franco
        self.table.setColumnWidth(6, 85)   # Turno Día
        self.table.setColumnWidth(7, 75)   # Hora Toma
        self.table.setColumnWidth(8, 75)   # Hora Deja
        self.table.setColumnWidth(9, 140)  # Novedad
        self.table.setColumnWidth(10, 90)  # Desde
        self.table.setColumnWidth(11, 90)  # Hasta
        self.table.setColumnWidth(12, 85)  # Cubierto
        self.table.setColumnWidth(13, 200) # Reemplazo
        self.table.setColumnWidth(14, 180) # Turno que reemplaza
        self.table.horizontalHeader().setSectionResizeMode(15, QHeaderView.Stretch) # Observaciones

        layout.addWidget(self.table)

        # 3. Pie Informativo
        self.lbl_conteo = QLabel("Mostrando 0 registros")
        self.lbl_conteo.setStyleSheet("color: #6B7280; font-size: 8.5pt;")
        layout.addWidget(self.lbl_conteo)

    def consultar(self) -> None:
        f_str = self.date_picker.date().toString("dd/MM/yyyy")
        dot = self.combo_dot.currentText()
        try:
            descubiertos = self.records_service.obtener_turnos_sin_cubrir(f_str, dotacion=dot)
            self.model.set_data(descubiertos)
            self._actualizar_conteo()
        except Exception as e:
            logger.error(f"Error al consultar turnos sin cubrir: {e}")

    def _on_search(self, text: str) -> None:
        self.proxy.set_global_search(text)
        self._actualizar_conteo()

    def _on_table_clicked(self, index) -> None:
        if not index.isValid():
            return
        # Clic específico sobre la columna 6 ("TURNO DÍA")
        if index.column() == 6:
            self._abrir_detalle_turno_index(index)

    def _on_table_double_clicked(self, index) -> None:
        if not index.isValid():
            return
        self._abrir_detalle_turno_index(index)

    def _abrir_detalle_turno_index(self, index) -> None:
        source_idx = self.proxy.mapToSource(index)
        row_data = self.model.get_row(source_idx.row())
        if not row_data:
            return

        turno_dia = str(row_data.get("turno_dia", "")).strip()
        turno_base = str(row_data.get("turnos", "")).strip()
        cod_turno = turno_dia if turno_dia and turno_dia not in ("FRANCO", "-", "None") else turno_base

        if not cod_turno or cod_turno in ("FRANCO", "-", "None"):
            QMessageBox.information(
                self,
                "Turno Día",
                "El agente se encuentra de Franco o no tiene un turno asignado para esta fecha.",
            )
            return

        from src.ui.views.dialogs.detalle_turno_dialog import DetalleTurnoDialog
        dlg = DetalleTurnoDialog(self.records_service, codigo_turno=cod_turno, parent=self)
        dlg.exec()

    def _actualizar_conteo(self) -> None:
        total = self.proxy.rowCount()
        self.lbl_conteo.setText(f"Total registrados en fecha: {total} novedades operativas")

    def _exportar_excel(self) -> None:
        f_str = self.date_picker.date().toString("dd/MM/yyyy")
        dot = self.combo_dot.currentText()
        filas = self.model.get_all_data()
        if not filas:
            QMessageBox.information(self, "Exportar", "No hay turnos para exportar en la fecha seleccionada.")
            return

        default_name = f"Turnos_Sin_Cubrir_{f_str.replace('/', '-')}.xlsx"
        file_path, _ = QFileDialog.getSaveFileName(self, "Guardar Reporte Excel", default_name, "Libro de Excel (*.xlsx)")
        if not file_path:
            return

        try:
            self.excel_service.export_turnos_sin_cubrir(file_path, filas, f_str, dotacion=dot)
            QMessageBox.information(self, "Éxito", f"Reporte exportado exitosamente a:\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al exportar:\n{e}")
