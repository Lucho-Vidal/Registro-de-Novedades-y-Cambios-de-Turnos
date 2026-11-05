"""Tests para la gestión separada de destinatarios de correo (Informe vs Ordenamiento) y UI."""

from __future__ import annotations

from pathlib import Path
import pytest
from PySide6.QtWidgets import QApplication

from src.data.database import DatabaseManager
from src.data.migrations import init_database_schema
from src.services.records_service import RecordsService
from src.ui.views.dialogs.admin_dialogs import AdminConfigWidget, AdminDestinatariosWidget


@pytest.fixture(scope="module")
def qapp():
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


def test_destinatarios_separacion_y_crud_records_service(tmp_path: Path):
    """Verifica que destinatarios_informe y destinatarios_ordenamiento sean listas independientes con CRUD completo."""
    db = DatabaseManager(tmp_path / "test_dests.sqlite")
    init_database_schema(db)
    svc = RecordsService(db)

    # 1. Crear destinatarios de informe
    id_inf1 = svc.crear_destinatario_informe("Control Trenes", "control@empresa.com", windows_user="Operador1")
    id_inf2 = svc.crear_destinatario_informe("Jefatura Personal", "jefatura@empresa.com", windows_user="Operador1")

    # 2. Crear destinatarios de ordenamiento
    id_ord1 = svc.crear_destinatario_ordenamiento("Banderillero PC", "banderillero@empresa.com", windows_user="Operador2")

    # 3. Comprobar aislamiento entre listas
    infs = svc.listar_destinatarios_informe(incluir_inactivos=True)
    assert len(infs) == 2
    assert {d["email"] for d in infs} == {"control@empresa.com", "jefatura@empresa.com"}

    ords = svc.listar_destinatarios_ordenamiento(incluir_inactivos=True)
    assert len(ords) == 1
    assert ords[0]["email"] == "banderillero@empresa.com"
    assert ords[0]["nombre"] == "Banderillero PC"

    # 4. Modificar y desactivar en cada tabla
    svc.actualizar_destinatario_informe(id_inf1, nombre="Control Centralizado", email="control.central@empresa.com", activo=0)
    infs_activos = svc.listar_destinatarios_informe(incluir_inactivos=False)
    assert len(infs_activos) == 1
    assert infs_activos[0]["email"] == "jefatura@empresa.com"

    infs_todos = svc.listar_destinatarios_informe(incluir_inactivos=True)
    assert len(infs_todos) == 2
    mod_inf = next(d for d in infs_todos if d["id"] == id_inf1)
    assert mod_inf["nombre"] == "Control Centralizado"
    assert mod_inf["activo"] == 0

    # 5. Modificar y desactivar ordenamiento
    svc.actualizar_destinatario_ordenamiento(id_ord1, nombre="Cabina Temperley", email="temperley@empresa.com", activo=1)
    ords_activos = svc.listar_destinatarios_ordenamiento(incluir_inactivos=False)
    assert len(ords_activos) == 1
    assert ords_activos[0]["nombre"] == "Cabina Temperley"
    assert ords_activos[0]["email"] == "temperley@empresa.com"


def test_admin_destinatarios_widget_ui(qapp, tmp_path: Path):
    """Verifica la visualización de ambas listas de destinatarios en AdminDestinatariosWidget y AdminConfigWidget."""
    db = DatabaseManager(tmp_path / "test_widget_dests.sqlite")
    init_database_schema(db)
    svc = RecordsService(db)

    svc.crear_destinatario_informe("Informe Dest 1", "inf1@test.com")
    svc.crear_destinatario_ordenamiento("Orden Dest 1", "ord1@test.com")

    widget = AdminDestinatariosWidget(svc)
    assert widget.tabs.count() == 2
    assert "Informe" in widget.tabs.tabText(0)
    assert "Ordenamientos" in widget.tabs.tabText(1)

    assert widget.table_inf.rowCount() == 1
    assert widget.table_inf.item(0, 1).text() == "Informe Dest 1"
    assert widget.table_inf.item(0, 2).text() == "inf1@test.com"

    assert widget.table_ord.rowCount() == 1
    assert widget.table_ord.item(0, 1).text() == "Orden Dest 1"
    assert widget.table_ord.item(0, 2).text() == "ord1@test.com"

    # Filtrar en tabla
    widget.txt_search_inf.setText("inexistente")
    assert widget.table_inf.isRowHidden(0)
    widget.txt_search_inf.clear()
    assert not widget.table_inf.isRowHidden(0)

    # Verificar integración dentro de AdminConfigWidget
    config_widget = AdminConfigWidget(db, records_service=svc)
    assert config_widget.tabs.count() == 3
    assert "Destinatarios" in config_widget.tabs.tabText(2)
