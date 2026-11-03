"""Diálogo de búsqueda y selección rápida de empleados por legajo o apellido."""

from typing import Any, Dict, List, Optional
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QPushButton,
    QTableView,
    QVBoxLayout,
)

from src.ui.models.base_table_model import BaseTableModel
from src.ui.models.multi_filter_proxy import MultiFilterProxyModel


class LegajoSearchDialog(QDialog):
    """Buscador modal interactivo de empleados para formularios."""

    def __init__(self, empleados: List[Dict[str, Any]], parent=None):
        super().__init__(parent)
        self.setWindowTitle("Búsqueda de Empleados")
        self.resize(780, 500)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)

        self._empleados = empleados
        self.selected_empleado: Optional[Dict[str, Any]] = None

        self._init_ui()

    def _init_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setSpacing(10)

        # Buscador
        search_layout = QHBoxLayout()
        lbl = QLabel("Buscar (Legajo, Apellido, Especialidad, Dotación):")
        lbl.setStyleSheet("font-weight: bold;")
        self.txt_search = QLineEdit()
        self.txt_search.setPlaceholderText("Escriba para filtrar en tiempo real...")
        self.txt_search.textChanged.connect(self._on_search_changed)
        search_layout.addWidget(lbl)
        search_layout.addWidget(self.txt_search)
        layout.addLayout(search_layout)

        # Tabla
        headers = ["LEGAJO", "APELLIDOS Y NOMBRES", "ESPECIALIDAD", "DOTACIÓN", "TURNOS", "FRANCO"]
        keys = ["legajo", "apellidos_nombres", "especialidad", "dotacion", "turnos", "franco"]

        self.table_model = BaseTableModel(headers, keys, self._empleados, parent=self)
        self.proxy_model = MultiFilterProxyModel(self)
        self.proxy_model.setSourceModel(self.table_model)

        self.table_view = QTableView()
        self.table_view.setModel(self.proxy_model)
        self.table_view.setSelectionBehavior(QTableView.SelectRows)
        self.table_view.setSelectionMode(QTableView.SingleSelection)
        self.table_view.setSortingEnabled(True)
        self.table_view.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table_view.doubleClicked.connect(self._on_double_click)
        layout.addWidget(self.table_view)

        # Botones
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        self.btn_cancel = QPushButton("Cancelar")
        self.btn_cancel.setObjectName("btnSecondary")
        self.btn_cancel.clicked.connect(self.reject)
        btn_layout.addWidget(self.btn_cancel)

        self.btn_select = QPushButton("Seleccionar")
        self.btn_select.clicked.connect(self._on_select)
        btn_layout.addWidget(self.btn_select)

        layout.addLayout(btn_layout)

    def _on_search_changed(self, text: str) -> None:
        self.proxy_model.set_global_search(text)

    def _on_double_click(self, index) -> None:
        self._on_select()

    def _on_select(self) -> None:
        selected_indexes = self.table_view.selectionModel().selectedRows()
        if not selected_indexes:
            return
        proxy_idx = selected_indexes[0]
        source_idx = self.proxy_model.mapToSource(proxy_idx)
        self.selected_empleado = self.table_model.get_row(source_idx.row())
        self.accept()
