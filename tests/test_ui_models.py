"""Pruebas unitarias para los modelos tabulares y proxy de filtrado en PySide6."""

import os
import shutil
import tempfile
import pytest
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication

from src.ui.models.base_table_model import BaseTableModel
from src.ui.models.multi_filter_proxy import MultiFilterProxyModel


@pytest.fixture(scope="session")
def qapp():
    app = QApplication.instance()
    if not app:
        app = QApplication([])
    return app


def test_base_table_model_and_proxy(qapp):
    headers = ["ID", "LEGAJO", "NOMBRE", "DOTACIÓN", "CUBIERTO"]
    keys = ["id", "legajo", "nombre", "dotacion", "cubierto"]

    data = [
        {"id": 1, "legajo": 101, "nombre": "PEREZ JUAN", "dotacion": "PLAZA", "cubierto": "Sí"},
        {"id": 2, "legajo": 102, "nombre": "GARCIA CARLOS", "dotacion": "TEMPERLEY", "cubierto": "No"},
        {"id": 3, "legajo": 103, "nombre": "LOPEZ ANA", "dotacion": "PLAZA", "cubierto": "Parcial"},
    ]

    model = BaseTableModel(headers, keys, data)
    assert model.rowCount() == 3
    assert model.columnCount() == 5

    # Proxy model
    proxy = MultiFilterProxyModel()
    proxy.setSourceModel(model)

    assert proxy.rowCount() == 3

    # Filtrar por texto global "ana"
    proxy.set_global_search("ana")
    assert proxy.rowCount() == 1
    idx = proxy.index(0, 2)
    assert proxy.data(idx, Qt.DisplayRole) == "LOPEZ ANA"

    # Limpiar filtro
    proxy.clear_all_filters()
    assert proxy.rowCount() == 3

    # Filtrar por columna dotación (col 3)
    proxy.set_column_filter(3, "PLAZA")
    assert proxy.rowCount() == 2
