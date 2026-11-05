"""Tests unitarios y de integración para notificaciones interactivas toast y apertura de formularios."""

import os
os.environ["QT_QPA_PLATFORM"] = "offscreen"

import pytest
from PySide6.QtCore import Qt, QPoint
from PySide6.QtGui import QMouseEvent
from PySide6.QtWidgets import QApplication

from src.data.database import DatabaseManager
from src.data.migrations import init_database_schema
from src.services.auth_service import AuthService
from src.services.records_service import RecordsService
from src.services.excel_service import ExcelService
from src.services.pdf_service import PDFService
from src.services.word_service import WordService
from src.services.backup_service import BackupService
from src.ui.views.dialogs.toast_widget import ToastWidget
from src.ui.views.dialogs.novedad_dialog import NovedadDialog
from src.ui.views.dialogs.cambio_turno_dialog import CambioTurnoDialog
from src.ui.views.dialogs.nuevos_registros_dialog import NuevosRegistrosDialog
from src.ui.main_window import MainWindow


@pytest.fixture(scope="session")
def qapp():
    return QApplication.instance() or QApplication([])


@pytest.fixture
def test_db(tmp_path):
    db_file = tmp_path / "test_notifications.sqlite"
    db = DatabaseManager(db_file)
    init_database_schema(db)

    with db.write_transaction() as conn:
        # Insertar usuario administrador y otro operador
        conn.execute("""
            INSERT INTO usuarios(id, username, nombre, legajo, password_hash, creado_en, ultimo_ingreso)
            VALUES (1, 'admin.test', 'Admin Test', 1001, 'dummy_hash', '2026-01-01', '2026-11-05 00:00:00');
        """)
        conn.execute("""
            INSERT INTO usuarios(id, username, nombre, legajo, password_hash, creado_en)
            VALUES (2, 'otro.operador', 'Otro Operador', 1002, 'dummy_hash', '2026-01-01');
        """)
        conn.execute("INSERT INTO roles(id, nombre) VALUES (1, 'Administrador');")
        conn.execute("INSERT INTO usuario_roles(usuario_id, rol_id) VALUES (1, 1);")

        # Empleados y tipos
        conn.execute("""
            INSERT INTO empleados(legajo, apellidos_nombres, dotacion, especialidad, turnos, franco, actualizado_en)
            VALUES (1234, 'GARCIA JUAN', 'PC', 'CONDUCTOR', '101', 'SAB/DOM', datetime('now','localtime'));
        """)
        conn.execute("""
            INSERT INTO empleados(legajo, apellidos_nombres, dotacion, especialidad, turnos, franco, actualizado_en)
            VALUES (5678, 'PEREZ CARLOS', 'PC', 'GUARDATREN', '202', 'LUN/MAR', datetime('now','localtime'));
        """)
        conn.execute("""
            INSERT INTO tipos_novedad(nombre, dias, exportar_descubierto)
            VALUES ('LICENCIA MEDICA', 5, 1);
        """)

    return db


def test_toast_widget_interactive_click(qapp):
    toast = ToastWidget()
    callbacks = []
    toast.show_message("Alerta", "Hay 1 nuevo registro", on_click=lambda: callbacks.append("opened"))

    assert toast.lbl_action.isVisible()
    assert "abrir el formulario" in toast.lbl_action.text()

    # Simular clic sobre el toast
    ev = QMouseEvent(QMouseEvent.MouseButtonPress, QPoint(15, 15), Qt.LeftButton, Qt.LeftButton, Qt.NoModifier)
    toast.mousePressEvent(ev)

    assert callbacks == ["opened"], "El callback asignado no se ejecutó al hacer clic sobre el ToastWidget."


def test_novedad_dialog_solo_lectura(qapp, test_db):
    rs = RecordsService(test_db)
    nov_id = rs.crear_novedad({
        "legajo": 1234,
        "novedad": "LICENCIA MEDICA",
        "fecha_inicio": "05/11/2026",
        "fecha_fin": "10/11/2026",
    })
    nov = rs.obtener_novedad(nov_id)
    assert nov is not None

    dlg = NovedadDialog(rs, novedad_existente=nov, solo_lectura=True)
    assert not dlg.btn_guardar.isVisible()
    assert not dlg.txt_legajo.isEnabled()
    assert dlg.btn_cancel.text() == "Cerrar"


def test_cambio_turno_dialog_solo_lectura(qapp, test_db):
    rs = RecordsService(test_db)
    cam_id = rs.crear_cambio({
        "legajo_1": 1234,
        "legajo_2": 5678,
        "fecha_cambio": "06/11/2026",
    })
    cam = rs.obtener_cambio(cam_id)
    assert cam is not None

    dlg = CambioTurnoDialog(rs, cambio_existente=cam, solo_lectura=True)
    assert not dlg.btn_guardar.isVisible()
    assert not dlg.txt_legajo_1.isEnabled()
    assert dlg.btn_cancel.text() == "Cerrar"


def test_nuevos_registros_dialog_open_actions(qapp):
    opened = []
    novs = [{"id": 10, "legajo": 1234, "apellidos_nombres": "GARCIA", "novedad": "LICENCIA", "registrado_en": "2026-11-05 12:00:00"}]
    cams = [{"id": 20, "legajo_1": 1234, "legajo_2": 5678, "fecha_cambio": "07/11/2026", "registrado_en": "2026-11-05 12:01:00"}]

    dlg = NuevosRegistrosDialog(
        novedades=novs,
        cambios=cams,
        on_abrir_novedad=lambda nid: opened.append(("nov", nid)),
        on_abrir_cambio=lambda cid: opened.append(("cam", cid)),
    )
    assert dlg.table.rowCount() == 2

    # Abrir la fila seleccionada (primera fila por defecto)
    dlg.table.selectRow(0)
    dlg._abrir_seleccionado()
    assert len(opened) == 1


def test_main_window_notificaciones_y_menu(qapp, test_db):
    rs = RecordsService(test_db)
    auth = AuthService(test_db)
    admin_user = auth.obtener_usuario(1)

    # Insertar una novedad de otro usuario con fecha posterior a ultimo_ingreso
    nov_id = rs.crear_novedad(
        {
            "legajo": 1234,
            "novedad": "LICENCIA MEDICA",
            "fecha_inicio": "05/11/2026",
            "registrado_en": "2026-11-05 08:00:00",
        },
        user_id=2,
        windows_user="otro.operador",
    )

    mw = MainWindow(
        db=test_db,
        records_service=rs,
        auth_service=auth,
        excel_service=ExcelService(test_db),
        pdf_service=PDFService(),
        word_service=WordService(),
        backup_service=BackupService(test_db),
        usuario_actual=admin_user,
    )

    mw._actualizar_notificaciones_nuevos()
    assert mw._sin_revisar_count >= 1
    assert "1" in mw.menu_nuevos.title()

    # Verificar que el menú contiene acciones interactivas
    mw._on_abrir_menu_nuevos()
    acciones = mw.menu_nuevos.actions()
    assert len(acciones) > 0
    # La acción de novedad debe mencionar #500
    nov_action = next((a for a in acciones if f"#{nov_id}" in a.text()), None)
    assert nov_action is not None, f"La novedad nueva #{nov_id} debe figurar como opción interactiva en el menú."
