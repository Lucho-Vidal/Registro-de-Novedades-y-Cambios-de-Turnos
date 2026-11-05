"""Diálogo modal para búsqueda inversa de turnos e itinerarios por número de tren."""

from typing import Any, Dict, List, Optional

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QCompleter,
    QDialog,
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
)

from src.core.logger import get_logger
from src.services.excel_service import ExcelService
from src.services.records_service import RecordsService
from src.ui.models.base_table_model import BaseTableModel

logger = get_logger("BuscadorTrenDialog")


class BuscadorTrenDialog(QDialog):
    """Herramienta de consulta rápida que lista todos los turnos que operan sobre un tren."""

    def __init__(
        self,
        records_service: RecordsService,
        excel_service: Optional[ExcelService] = None,
        parent=None,
    ):
        super().__init__(parent)
        self.records_service = records_service
        self.excel_service = excel_service
        self._resultados: List[Dict[str, Any]] = []
        self._ordenamientos: List[Dict[str, Any]] = []

        self.setWindowTitle("🔍 Buscador de Turnos por Número de Tren")
        self.resize(1050, 600)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)

        self._init_ui()

    def showEvent(self, event) -> None:
        super().showEvent(event)
        self.txt_tren.setFocus()

    def _init_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)

        # 1. Cabecera y Barra de búsqueda
        top_frame = QFrame()
        top_frame.setStyleSheet("background-color: #FFFFFF; border: 1px solid #D1D5DB; border-radius: 6px; padding: 10px;")
        top_layout = QHBoxLayout(top_frame)

        top_layout.addWidget(QLabel("<b>Número de Tren:</b>"))
        self.txt_tren = QLineEdit()
        self.txt_tren.setPlaceholderText("Ej. 2303, 1310, 4356...")
        self.txt_tren.setFixedWidth(200)
        self.txt_tren.returnPressed.connect(self._buscar)
        top_layout.addWidget(self.txt_tren)

        btn_buscar = QPushButton("🔍 Consultar Turnos")
        btn_buscar.clicked.connect(self._buscar)
        top_layout.addWidget(btn_buscar)

        top_layout.addStretch()

        btn_importar = QPushButton("📥 Cargar / Actualizar Itinerario (Excel)")
        btn_importar.setObjectName("btnSecondary")
        btn_importar.clicked.connect(self._importar_excel)
        top_layout.addWidget(btn_importar)

        layout.addWidget(top_frame)

        # 2. Resumen y tabla de resultados
        self.lbl_conteo = QLabel("Ingrese un número de tren y presione 'Consultar Turnos'.")
        self.lbl_conteo.setStyleSheet("color: #4B5563; font-size: 9.5pt;")
        layout.addWidget(self.lbl_conteo)

        headers = [
            "TREN", "TURNO", "DEPÓSITO", "VUELTA", "REF", "ORIGEN", "SALIDA",
            "DESTINO", "LLEGADA", "OBS. VUELTA", "TOMA TURNO", "DEJE TURNO", "ITINERARIO",
        ]
        keys = [
            "nro_tren", "codigo_turno", "deposito_base", "nro_vuelta", "ref", "origen",
            "hora_salida", "destino", "hora_llegada", "obs_vuelta", "toma_servicio",
            "deje_servicio", "itinerario_vigente",
        ]

        self.model = BaseTableModel(headers, keys, parent=self)
        self.table_view = QTableView()
        self.table_view.setModel(self.model)
        self.table_view.setSelectionBehavior(QTableView.SelectRows)
        self.table_view.setSelectionMode(QTableView.SingleSelection)
        self.table_view.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.table_view.horizontalHeader().setSectionResizeMode(9, QHeaderView.Stretch)
        layout.addWidget(self.table_view, 2)

        self.lbl_ords = QLabel("<b>📄 Ordenamientos Operativos vinculados al tren:</b>")
        layout.addWidget(self.lbl_ords)

        headers_ords = [
            "N° ORDEN", "FECHA", "HORA", "LEGAJO", "PERSONAL ASIGNADO",
            "DOTACIÓN", "TURNO / CICLO", "TRENES", "MOTIVO DE LA ORDEN", "REF ESTACIÓN", "EMISOR",
        ]
        keys_ords = [
            "nro_orden", "fecha_emision", "hora_emision", "legajo_personal", "nombre_personal",
            "deposito", "ciclo_o_turno", "trenes_involucrados", "motivo_orden", "ref_pct", "emisor",
        ]
        self.model_ords = BaseTableModel(headers_ords, keys_ords, parent=self)
        self.table_ords = QTableView()
        self.table_ords.setModel(self.model_ords)
        self.table_ords.setSelectionBehavior(QTableView.SelectRows)
        self.table_ords.setSelectionMode(QTableView.SingleSelection)
        self.table_ords.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.table_ords.horizontalHeader().setSectionResizeMode(8, QHeaderView.Stretch)
        layout.addWidget(self.table_ords, 1)

        # 3. Barra inferior de botones
        bottom_layout = QHBoxLayout()
        self.btn_exportar = QPushButton("📊 Exportar Resultados a Excel")
        self.btn_exportar.setObjectName("btnSecondary")
        self.btn_exportar.clicked.connect(self._exportar_excel)
        bottom_layout.addWidget(self.btn_exportar)

        bottom_layout.addStretch()

        btn_cerrar = QPushButton("Cerrar")
        btn_cerrar.setObjectName("btnSecondary")
        btn_cerrar.clicked.connect(self.accept)
        bottom_layout.addWidget(btn_cerrar)

        layout.addLayout(bottom_layout)

    def _buscar(self) -> None:
        tren = self.txt_tren.text().strip()
        if not tren:
            QMessageBox.information(self, "Búsqueda", "Ingrese un número de tren para buscar.")
            return

        self._resultados = self.records_service.buscar_turnos_por_tren(tren)
        self.model.set_data(self._resultados)

        self._ordenamientos = self.records_service.buscar_ordenamientos_activos(tren)
        self.model_ords.set_data(self._ordenamientos)
        self.lbl_ords.setText(f"<b>📄 Ordenamientos Operativos vinculados al tren {tren} ({len(self._ordenamientos)}):</b>")

        if not self._resultados and not self._ordenamientos:
            self.lbl_conteo.setText(f"No se encontraron turnos ni ordenamientos vinculados al tren <b>{tren}</b>.")
        else:
            self.lbl_conteo.setText(
                f"Se encontraron <b>{len(self._resultados)}</b> intervenciones de turnos y <b>{len(self._ordenamientos)}</b> ordenamientos para el tren <b>{tren}</b>:"
            )

    def _importar_excel(self) -> None:
        f_path, _ = QFileDialog.getOpenFileName(
            self,
            "Seleccionar Archivo de Itinerario",
            "",
            "Archivos Excel (*.xlsx *.xlsm)",
        )
        if not f_path:
            return

        try:
            turnos, vueltas = self.records_service.importar_itinerario_excel(f_path)
            QMessageBox.information(
                self,
                "Itinerario Importado",
                f"Se importaron exitosamente {turnos} turnos y {vueltas} vueltas desde:\n{f_path}",
            )
            if self.txt_tren.text().strip():
                self._buscar()
        except Exception as e:
            QMessageBox.critical(self, "Error de Importación", f"No se pudo importar el archivo:\n{e}")

    def _exportar_excel(self) -> None:
        if not self._resultados:
            QMessageBox.information(self, "Exportación", "No hay resultados para exportar.")
            return

        tren = self.txt_tren.text().strip() or "General"
        def_name = f"Turnos_Tren_{tren}.xlsx"
        f_path, _ = QFileDialog.getSaveFileName(self, "Exportar Resultados", def_name, "Excel (*.xlsx)")
        if not f_path:
            return

        try:
            import openpyxl
            from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = f"Tren {tren}"

            headers = [
                "Tren", "Código Turno", "Depósito", "Vuelta N°", "Ref", "Origen",
                "Hora Salida", "Destino", "Hora Llegada", "Obs. Vuelta", "Toma", "Deje", "Itinerario"
            ]
            ws.append(headers)

            header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
            header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            thin_border = Border(
                left=Side(style="thin", color="D1D5DB"),
                right=Side(style="thin", color="D1D5DB"),
                top=Side(style="thin", color="D1D5DB"),
                bottom=Side(style="thin", color="D1D5DB"),
            )

            for col_idx in range(1, len(headers) + 1):
                cell = ws.cell(1, col_idx)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")

            for r in self._resultados:
                ws.append([
                    r.get("nro_tren"),
                    r.get("codigo_turno"),
                    r.get("deposito_base"),
                    r.get("nro_vuelta"),
                    r.get("ref"),
                    r.get("origen"),
                    r.get("hora_salida"),
                    r.get("destino"),
                    r.get("hora_llegada"),
                    r.get("obs_vuelta"),
                    r.get("toma_servicio"),
                    r.get("deje_servicio"),
                    r.get("itinerario_vigente"),
                ])

            for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=len(headers)):
                for cell in row:
                    cell.border = thin_border
                    cell.alignment = Alignment(vertical="center")

            for col in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = openpyxl.utils.get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = max(max_len + 3, 11)

            wb.save(f_path)
            QMessageBox.information(self, "Exportación Exitosa", f"Archivo generado en:\n{f_path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Fallo al exportar archivo Excel:\n{e}")
