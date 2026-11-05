"""Tests para validación de RBAC, permisos estrictos, cambio de roles y dimensiones de calendario de ciclo."""

from __future__ import annotations

from pathlib import Path
import pytest
from PySide6.QtCore import Qt, QModelIndex
from PySide6.QtWidgets import QApplication, QHeaderView

from src.data.database import DatabaseManager
from src.data.migrations import init_database_schema
from src.services.auth_service import AuthService
from src.services.excel_service import ExcelService
from src.services.pdf_service import PDFService
from src.services.records_service import RecordsService
from src.ui.models.base_table_model import BaseTableModel
from src.ui.views.dialogs.calendario_ciclo_dialog import CalendarioCicloDialog
from src.ui.views.personal_ciclo_view import PersonalCicloView


@pytest.fixture(scope="module")
def qapp():
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


def test_auth_service_permisos_estrictos_sin_fuga_registros_modificar(tmp_path: Path):
    """Verifica que el permiso registros.modificar NO confiera permisos administrativos ni de ciclo indebidamente."""
    db = DatabaseManager(tmp_path / "test_rbac_strict.sqlite")
    init_database_schema(db)
    auth_srv = AuthService(db)

    # 1. Crear rol Supervisor con solo 'registros.modificar' y 'novedades.ver'
    with db.write_transaction() as conn:
        r_id = conn.execute("INSERT INTO roles(nombre) VALUES ('Supervisor')").lastrowid
        for cod in ("registros.modificar", "novedades.ver"):
            conn.execute("INSERT OR IGNORE INTO permisos(codigo) VALUES (?)", (cod,))
            p_id = conn.execute("SELECT id FROM permisos WHERE codigo=?", (cod,)).fetchone()[0]
            conn.execute("INSERT INTO rol_permisos(rol_id, permiso_id) VALUES (?, ?)", (r_id, p_id))

    # 2. Crear usuario supervisor
    sup_id = auth_srv.crear_usuario(
        username="super1",
        password="Password123!",
        nombre="Supervisor Juan",
        legajo=5001,
        roles=("Supervisor",),
    )

    # 3. Comprobar permisos asignados
    assert auth_srv.tiene_permiso(sup_id, "registros.modificar") is True
    assert auth_srv.tiene_permiso(sup_id, "novedades.ver") is True

    # 4. Comprobar que NO tiene permisos administrativos ni de asignación de ciclo
    assert auth_srv.tiene_permiso(sup_id, "usuarios.administrar") is False
    assert auth_srv.tiene_permiso(sup_id, "roles.administrar") is False
    assert auth_srv.tiene_permiso(sup_id, "sesion.configurar") is False
    assert auth_srv.tiene_permiso(sup_id, "backup.gestionar") is False
    assert auth_srv.tiene_permiso(sup_id, "reemplazos.asignar") is False
    assert auth_srv.tiene_permiso(sup_id, "empleados.administrar") is False


def test_auth_service_actualizar_usuario_con_cambio_de_rol(tmp_path: Path):
    """Verifica que actualizar_usuario modifique datos y persista el cambio de rol en la base de datos."""
    db = DatabaseManager(tmp_path / "test_user_role_change.sqlite")
    init_database_schema(db)
    auth_srv = AuthService(db)

    # Crear usuario con rol Operador
    u_id = auth_srv.crear_usuario("testuser", "Pass1234!", nombre="Test User", legajo=1234, roles=("Operador",))
    assert auth_srv.roles_de_usuario(u_id) == ["Operador"]

    # Cambiar rol a Supervisor y actualizar legajo
    auth_srv.actualizar_usuario(
        user_id=u_id,
        username="testuser_mod",
        nombre="Test User Modificado",
        legajo=9999,
        activo=1,
        roles=["Supervisor"],
    )

    # Verificar persistencia en base de datos
    u_mod = auth_srv.obtener_usuario(u_id)
    assert u_mod["username"] == "testuser_mod"
    assert u_mod["nombre"] == "Test User Modificado"
    assert u_mod["legajo"] == 9999
    assert auth_srv.roles_de_usuario(u_id) == ["Supervisor"]

    # Verificar que aparece en listar_usuarios con el nuevo rol
    lista = auth_srv.listar_usuarios()
    u_item = next(x for x in lista if x["id"] == u_id)
    assert u_item["roles"] == "Supervisor"


def test_base_table_model_tooltip_role():
    """Verifica que BaseTableModel retorne ToolTipRole con el texto de la celda."""
    headers = ["ID", "OBSERVACIONES"]
    keys = ["id", "observaciones"]
    data = [{"id": 1, "observaciones": "Texto extenso de prueba para tooltip"}]

    model = BaseTableModel(headers, keys)
    model.set_data(data)

    idx_obs = model.index(0, 1)
    tooltip = model.data(idx_obs, Qt.ToolTipRole)
    assert tooltip == "Texto extenso de prueba para tooltip"


def test_calendario_ciclo_dialog_columnas_y_permisos(qapp, tmp_path: Path):
    """Verifica el dimensionamiento de columnas y la protección de permisos en CalendarioCicloDialog."""
    db = DatabaseManager(tmp_path / "test_cal_dialog.sqlite")
    init_database_schema(db)
    rec_srv = RecordsService(db)
    pdf_srv = PDFService()
    excel_srv = ExcelService(db)

    emp = {
        "legajo": 8001,
        "apellidos_nombres": "Ciclo Test",
        "dotacion": "PC",
        "especialidad": "Conductor Diésel",
        "turnos": "CICLO1",
        "franco": "Lunes",
        "franco_horario": "00:00 - 24:00",
    }

    # 1. Dialog con usuario SIN permiso reemplazos.asignar
    dlg_sin_permiso = CalendarioCicloDialog(
        empleado=emp,
        records_service=rec_srv,
        pdf_service=pdf_srv,
        excel_service=excel_srv,
        mes=11,
        anio=2026,
        tiene_permiso_fn=lambda p: False,
    )

    # Verificar que el botón de asignar está deshabilitado
    assert dlg_sin_permiso.btn_asignar.isEnabled() is False

    # Verificar dimensionamiento de columnas
    header = dlg_sin_permiso.table_view.horizontalHeader()
    # Col 9 (REEMPLAZA A / MOTIVO) debe tener ancho acotado (~190 px) y no modo Stretch
    assert dlg_sin_permiso.table_view.columnWidth(9) == 190
    assert header.sectionResizeMode(9) == QHeaderView.Interactive

    # Col 10 (OBSERVACIONES) debe estar en modo Stretch para tomar el ancho completo
    assert header.sectionResizeMode(10) == QHeaderView.Stretch

    # 2. Dialog con usuario CON permiso reemplazos.asignar
    dlg_con_permiso = CalendarioCicloDialog(
        empleado=emp,
        records_service=rec_srv,
        pdf_service=pdf_srv,
        excel_service=excel_srv,
        mes=11,
        anio=2026,
        tiene_permiso_fn=lambda p: p == "reemplazos.asignar",
    )
    assert dlg_con_permiso.btn_asignar.isEnabled() is True


def test_personal_ciclo_view_proteccion_drag_drop(qapp, tmp_path: Path):
    """Verifica que PersonalCicloView bloquee la asignación y liberación cuando no se tiene reemplazos.asignar."""
    db = DatabaseManager(tmp_path / "test_ciclo_view.sqlite")
    init_database_schema(db)
    rec_srv = RecordsService(db)
    pdf_srv = PDFService()
    excel_srv = ExcelService(db)

    # Vista con permisos de solo lectura (sin reemplazos.asignar)
    view = PersonalCicloView(
        rec_srv,
        pdf_srv,
        excel_srv,
        tiene_permiso_fn=lambda p: p == "reemplazos.ver",
    )

    assert view._can("reemplazos.asignar") is False
    assert view._can("reemplazos.ver") is True
