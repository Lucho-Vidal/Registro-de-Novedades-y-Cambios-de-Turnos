"""Diálogo modal para visualizar el turno completo y sus vueltas asociadas con filtro de itinerario."""

from datetime import datetime
from typing import Any, Dict, List, Optional

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMessageBox,
    QPushButton,
    QTableView,
    QVBoxLayout,
)

from src.core.logger import get_logger
from src.services.records_service import RecordsService
from src.ui.models.base_table_model import BaseTableModel

logger = get_logger("DetalleTurnoDialog")


class DetalleTurnoDialog(QDialog):
    """Visualizador detallado del diagrama completo de un turno con selección de itinerario."""

    def __init__(
        self,
        records_service: RecordsService,
        codigo_turno: str,
        itinerario_inicial: Optional[str] = None,
        parent=None,
    ):
        super().__init__(parent)
        self.records_service = records_service
        self.codigo_turno = str(codigo_turno or "").strip()
        self.itinerario_inicial = itinerario_inicial

        self._turno_actual: Optional[Dict[str, Any]] = None
        self._vueltas: List[Dict[str, Any]] = []

        self.setWindowTitle(f"Detalle Completo del Turno: {self.codigo_turno}")
        self.resize(920, 620)
        self.setWindowFlags((self.windowFlags() & ~Qt.WindowContextHelpButtonHint) | Qt.WindowCloseButtonHint)

        self._init_ui()
        self._cargar_itinerarios()
        self._cargar_turno()

    def _init_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(10)

        # 1. Tarjeta de resumen de datos de cabecera y selector de itinerario
        card = QFrame()
        card.setObjectName("cardFrame")
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(12, 10, 12, 10)
        card_layout.setSpacing(8)

        row1 = QHBoxLayout()
        self.lbl_codigo = QLabel(f"<h2>Turno: <span style='color: #2563EB;'>{self.codigo_turno}</span></h2>")
        row1.addWidget(self.lbl_codigo)
        row1.addStretch()

        row1.addWidget(QLabel("<b>Itinerario:</b>"))
        self.combo_itinerario = QComboBox()
        self.combo_itinerario.setMinimumWidth(160)
        self.combo_itinerario.currentIndexChanged.connect(self._on_itinerario_changed)
        row1.addWidget(self.combo_itinerario)
        card_layout.addLayout(row1)

        row2 = QHBoxLayout()
        self.lbl_tipo = QLabel("<b>Tipo:</b> -")
        self.lbl_deposito = QLabel("<b>Depósito:</b> -")
        self.lbl_toma = QLabel("<b>Toma de servicio:</b> -")
        self.lbl_deje = QLabel("<b>Deje de servicio:</b> -")
        self.lbl_duracion = QLabel("<b>Duración:</b> -")

        row2.addWidget(self.lbl_tipo)
        row2.addWidget(self.lbl_deposito)
        row2.addWidget(self.lbl_toma)
        row2.addWidget(self.lbl_deje)
        row2.addWidget(self.lbl_duracion)
        row2.addStretch()
        card_layout.addLayout(row2)

        layout.addWidget(card)

        # 2. Grilla de Vueltas
        headers = [
            "VUELTA N°", "REF.", "N° TREN", "ORIGEN", "SALIDA",
            "DESTINO", "LLEGADA", "OBSERVACIONES DEL TRAMO",
        ]
        keys = [
            "nro_vuelta", "ref", "nro_tren", "origen", "hora_salida",
            "destino", "hora_llegada", "observaciones_tramo",
        ]

        self.model = BaseTableModel(headers, keys, parent=self)
        self.table = QTableView()
        self.table.setModel(self.model)
        self.table.setSelectionBehavior(QTableView.SelectRows)
        self.table.setSelectionMode(QTableView.SingleSelection)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table.setColumnWidth(0, 85)
        self.table.setColumnWidth(1, 65)
        self.table.setColumnWidth(2, 85)
        self.table.setColumnWidth(3, 130)
        self.table.setColumnWidth(4, 85)
        self.table.setColumnWidth(5, 130)
        self.table.setColumnWidth(6, 85)
        self.table.horizontalHeader().setSectionResizeMode(7, QHeaderView.Stretch)
        layout.addWidget(self.table)

        # 3. Barra inferior de control
        bottom_bar = QHBoxLayout()
        self.lbl_conteo = QLabel("Total: 0 tramos")
        self.lbl_conteo.setStyleSheet("color: #6B7280; font-size: 9pt;")
        bottom_bar.addWidget(self.lbl_conteo)

        bottom_bar.addStretch()

        self.btn_exportar = QPushButton("📊 Exportar a Excel")
        self.btn_exportar.setObjectName("btnSecondary")
        self.btn_exportar.clicked.connect(self._exportar_excel)
        bottom_bar.addWidget(self.btn_exportar)

        self.btn_cerrar = QPushButton("Cerrar")
        self.btn_cerrar.clicked.connect(self.accept)
        bottom_bar.addWidget(self.btn_cerrar)

        layout.addLayout(bottom_bar)

    def _cargar_itinerarios(self) -> None:
        self.combo_itinerario.blockSignals(True)
        self.combo_itinerario.clear()
        itinerarios = self.records_service.listar_itinerarios_disponibles()
        for iti in itinerarios:
            self.combo_itinerario.addItem(f"Itinerario {iti}", userData=iti)

        if self.itinerario_inicial:
            for idx in range(self.combo_itinerario.count()):
                if self.combo_itinerario.itemData(idx) == self.itinerario_inicial:
                    self.combo_itinerario.setCurrentIndex(idx)
                    break
        self.combo_itinerario.blockSignals(False)

    def _on_itinerario_changed(self) -> None:
        self._cargar_turno()

    def _cargar_turno(self) -> None:
        iti_sel = self.combo_itinerario.currentData()
        t = self.records_service.obtener_turno_por_codigo(self.codigo_turno, itinerario=iti_sel)
        self._turno_actual = t
        if t:
            self.lbl_codigo.setText(f"<h2>Turno: <span style='color: #2563EB;'>{t.get('codigo_turno', self.codigo_turno)}</span></h2>")
            self.lbl_tipo.setText(f"<b>Tipo:</b> {t.get('tipo_turno', '-')}")
            self.lbl_deposito.setText(f"<b>Depósito:</b> {t.get('deposito_base', '-')}")
            self.lbl_toma.setText(f"<b>Toma:</b> {t.get('toma_servicio', '-')}")
            self.lbl_deje.setText(f"<b>Deje:</b> {t.get('deje_servicio', '-')}")
            self.lbl_duracion.setText(f"<b>Duración:</b> {t.get('duracion_jornada', '-')}")

            self._vueltas = t.get("vueltas", [])
            self.model.set_data(self._vueltas)
            self.lbl_conteo.setText(f"Total: {len(self._vueltas)} tramos en el turno")
        else:
            self.lbl_tipo.setText("<b>Tipo:</b> -")
            self.lbl_deposito.setText("<b>Depósito:</b> -")
            self.lbl_toma.setText("<b>Toma:</b> -")
            self.lbl_deje.setText("<b>Deje:</b> -")
            self.lbl_duracion.setText("<b>Duración:</b> -")
            self._vueltas = []
            self.model.set_data([])
            self.lbl_conteo.setText("No se encontró diagrama para este turno en el itinerario seleccionado.")

    def _exportar_excel(self) -> None:
        if not self._vueltas:
            QMessageBox.information(self, "Exportar", "No hay tramos para exportar.")
            return

        def_name = f"Turno_{self.codigo_turno}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        path, _ = QFileDialog.getSaveFileName(self, "Exportar Vueltas del Turno", def_name, "Excel (*.xlsx)")
        if not path:
            return

        try:
            import openpyxl

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = f"Turno {self.codigo_turno}"

            ws.append(["Código Turno", "Turno Base", "Tipo", "Depósito", "Toma", "Deje", "Duración", "Itinerario"])
            t = self._turno_actual or {}
            ws.append([
                t.get("codigo_turno", self.codigo_turno),
                t.get("turno_base", ""),
                t.get("tipo_turno", ""),
                t.get("deposito_base", ""),
                t.get("toma_servicio", ""),
                t.get("deje_servicio", ""),
                t.get("duracion_jornada", ""),
                t.get("itinerario_vigente", ""),
            ])
            ws.append([])  # Línea en blanco

            ws.append(["Vuelta N°", "Ref.", "N° Tren", "Origen", "Salida", "Destino", "Llegada", "Observaciones"])
            for v in self._vueltas:
                ws.append([
                    v.get("nro_vuelta"),
                    v.get("ref"),
                    v.get("nro_tren"),
                    v.get("origen"),
                    v.get("hora_salida"),
                    v.get("destino"),
                    v.get("hora_llegada"),
                    v.get("observaciones_tramo"),
                ])

            wb.save(path)
            QMessageBox.information(self, "Exportación Exitosa", f"Archivo generado con éxito en:\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo exportar a Excel:\n{e}")
