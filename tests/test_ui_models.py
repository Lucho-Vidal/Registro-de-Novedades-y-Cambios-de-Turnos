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


def test_dashboard_charts_theme_switch(qapp, tmp_path):
    from src.core.theme import aplicar_tema_visual, guardar_tema_visual, obtener_tema_visual
    from src.data.database import DatabaseManager
    from src.data.migrations import init_database_schema
    from src.services.excel_service import ExcelService
    from src.services.pdf_service import PDFService
    from src.services.records_service import RecordsService
    from src.ui.views.dashboard_view import DashboardView

    tema_previo = obtener_tema_visual()
    db = DatabaseManager(tmp_path / "test_dash.sqlite")
    init_database_schema(db)
    rs = RecordsService(db)
    es = ExcelService(db)
    ps = PDFService()

    dash = DashboardView(rs, es, ps)
    try:
        dash.show()
        qapp.processEvents()

        aplicar_tema_visual("oscuro", db=db)
        dash.actualizar_tema("oscuro")
        qapp.processEvents()
        assert obtener_tema_visual() == "oscuro"
        assert dash.chart_tipo._es_oscuro is True
        assert dash.chart_dotacion._es_oscuro is True
        assert dash.chart_tendencia._es_oscuro is True
        assert dash.chart_dia_semana._es_oscuro is True

        aplicar_tema_visual("claro", db=db)
        dash.actualizar_tema("claro")
        qapp.processEvents()
        assert obtener_tema_visual() == "claro"
        assert dash.chart_tipo._es_oscuro is False
        assert dash.chart_dotacion._es_oscuro is False
        assert dash.chart_tendencia._es_oscuro is False
        assert dash.chart_dia_semana._es_oscuro is False
    finally:
        dash.close()
        guardar_tema_visual(tema_previo)


def test_cambiar_password_dialog_and_service(qapp, tmp_path, monkeypatch):
    from PySide6.QtWidgets import QMessageBox
    from src.data.database import DatabaseManager
    from src.data.migrations import init_database_schema
    from src.services.auth_service import AuthService
    from src.ui.views.dialogs.cambiar_password_dialog import CambiarMiPasswordDialog

    monkeypatch.setattr(QMessageBox, "information", lambda *args, **kwargs: QMessageBox.Ok)

    db = DatabaseManager(tmp_path / "test_pwd.sqlite")
    init_database_schema(db)
    auth = AuthService(db)
    uid = auth.crear_administrador_inicial("operador.test", "clave1", legajo=555)
    user = auth.autenticar("OPERADOR.TEST", "clave1")
    assert user is not None

    dlg = CambiarMiPasswordDialog(auth, user)
    assert dlg.btn_save.isDefault() is True
    assert dlg.btn_cancel.isDefault() is False

    dlg.txt_actual.setText("clave1")
    dlg.txt_nueva.setText("123")
    dlg.txt_confirm.setText("123")
    dlg._on_guardar()

    assert dlg.result() == CambiarMiPasswordDialog.Accepted
    assert auth.autenticar("operador.test", "123") is not None
    assert auth.autenticar("operador.test", "clave1") is None

