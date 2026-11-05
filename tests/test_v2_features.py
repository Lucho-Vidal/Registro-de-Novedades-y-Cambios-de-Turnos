"""Pruebas unitarias y de integración para las mejoras y correcciones de RENO v2.0.0:

1. Diálogos NovedadDialog y CambioTurnoDialog: inicialización correcta de elementos.
2. AsignarJornadaDialog: verificación de entradas QTimeEdit.
3. Exportación de Calendario de Ciclo a Excel (Individual y Consolidado).
4. Métricas de Dashboard y soporte de Modo Oscuro en gráficos y tarjetas KPI.
5. Especialidades dinámicas en RecordsService y filtrado en PersonalCicloView.
6. Pestaña EmpleadosView: integración y filtros.
"""

from pathlib import Path
from datetime import date
import pytest
from PySide6.QtCore import QTime
from PySide6.QtWidgets import QApplication

from src.core.settings import AppSettings
from src.data.database import DatabaseManager
from src.data.migrations import init_database_schema
from src.domain.enums import PERMISOS_BASE
from src.services.auth_service import AuthService
from src.services.excel_service import ExcelService
from src.services.pdf_service import PDFService
from src.services.records_service import RecordsService
from src.ui.views.dialogs.calendario_ciclo_dialog import AsignarJornadaDialog, CalendarioCicloDialog
from src.ui.views.dialogs.cambio_turno_dialog import CambioTurnoDialog
from src.ui.views.dialogs.novedad_dialog import NovedadDialog
from src.ui.views.dashboard_view import DashboardView
from src.ui.views.empleados_view import EmpleadosView
from src.ui.views.personal_ciclo_view import PersonalCicloView


@pytest.fixture
def test_env(tmp_path: Path):
    app = QApplication.instance() or QApplication([])
    db_file = tmp_path / "test_v2.sqlite"
    db = DatabaseManager(db_file, journal_mode="DELETE")
    init_database_schema(db)

    auth = AuthService(db)
    auth.inicializar_permisos(PERMISOS_BASE)
    admin_id = auth.crear_administrador_inicial("test_admin", "TestPass123!", legajo=9999)
    user = auth.autenticar("test_admin", "TestPass123!")

    records = RecordsService(db)
    excel = ExcelService(db)
    pdf = PDFService()

    # Insertar un empleado de ciclo para tests usando records.crear_empleado
    records.crear_empleado(
        {
            "legajo": 12345,
            "apellidos_nombres": "Perez Juan",
            "especialidad": "Conductor Electrico",
            "dotacion": "PC",
            "turnos": "CICLO 1",
            "franco": "Domingo",
            "franco_horario": "00:00 - 24:00",
            "jornada_horas": 6,
        },
        user_id=admin_id,
        windows_user="test_admin",
    )

    return {
        "db": db,
        "auth": auth,
        "records": records,
        "excel": excel,
        "pdf": pdf,
        "user": user,
        "tmp_path": tmp_path,
    }


def test_dialogs_initialization_not_empty(test_env):
    """Verifica que NovedadDialog y CambioTurnoDialog construyan sus widgets sin quedar vacíos."""
    records = test_env["records"]
    user = test_env["user"]

    # NovedadDialog
    nov_dlg = NovedadDialog(records_service=records, usuario_actual=user)
    assert hasattr(nov_dlg, "txt_legajo")
    assert hasattr(nov_dlg, "combo_tipo")
    assert nov_dlg.txt_legajo is not None

    # CambioTurnoDialog
    cam_dlg = CambioTurnoDialog(records_service=records, usuario_actual=user)
    assert hasattr(cam_dlg, "txt_legajo_1")
    assert hasattr(cam_dlg, "txt_legajo_2")
    assert cam_dlg.txt_legajo_1 is not None


def test_asignar_jornada_dialog_time_inputs(test_env):
    """Verifica que AsignarJornadaDialog use QTimeEdit y capture los horarios correctamente."""
    records = test_env["records"]

    dlg = AsignarJornadaDialog(
        legajo=12345,
        nombre="Perez Juan",
        fecha_iso="2026-11-04",
        dia_semana="Miércoles",
        records_service=records,
    )
    assert hasattr(dlg, "time_toma")
    assert hasattr(dlg, "time_deja")

    dlg.time_toma.setTime(QTime(7, 30))
    dlg.time_deja.setTime(QTime(13, 30))
    dlg.txt_obs.setText("Jornada especial")

    dlg._on_accept()
    res = dlg.datos_resultado
    assert res is not None
    assert res["toma"] == "07:30"
    assert res["deja"] == "13:30"
    assert res["franco_anticipado"] == 0
    assert res["observaciones"] == "Jornada especial"


def test_excel_export_calendario_ciclo_individual_y_todos(test_env):
    """Verifica las exportaciones Excel del cronograma mensual individual y consolidado general."""
    records = test_env["records"]
    excel = test_env["excel"]
    tmp_path = test_env["tmp_path"]

    emp = records.listar_empleados_ciclo()[0]
    dias = records.obtener_calendario_mensual_ciclo(emp["legajo"], 2026, 11)

    # 1. Exportación Individual
    ind_path = tmp_path / "ciclo_ind.xlsx"
    res_ind = excel.export_calendario_ciclo_individual(str(ind_path), emp, dias, 11, 2026)
    assert Path(res_ind).exists()
    assert Path(res_ind).stat().st_size > 0

    # 2. Exportación Consolidada de Todos
    all_path = tmp_path / "ciclo_todos.xlsx"
    res_all = excel.export_calendario_ciclo_todos(str(all_path), records, 11, 2026)
    assert Path(res_all).exists()
    assert Path(all_path).stat().st_size > 0


def test_dashboard_dark_mode_and_theme_update(test_env):
    """Verifica la respuesta del DashboardView ante el cambio de tema claro/oscuro."""
    from src.core.theme import obtener_tema_visual, guardar_tema_visual
    records = test_env["records"]
    excel = test_env["excel"]
    pdf = test_env["pdf"]

    dash = DashboardView(records, excel, pdf)
    tema_orig = obtener_tema_visual()

    try:
        # Modo Claro
        guardar_tema_visual("claro")
        dash.actualizar_tema()
        assert "1F497D" in dash.lbl_titulo.styleSheet()

        # Modo Oscuro
        guardar_tema_visual("oscuro")
        dash.actualizar_tema()
        assert "60A5FA" in dash.lbl_titulo.styleSheet()
    finally:
        guardar_tema_visual(tema_orig)


def test_especialidades_dinamicas_y_empleados_view(test_env):
    """Verifica la carga dinámica de especialidades y la vista EmpleadosView."""
    records = test_env["records"]
    excel = test_env["excel"]
    user = test_env["user"]

    esps = records.listar_especialidades_empleados()
    assert "Conductor Electrico" in esps

    # PersonalCicloView
    ciclo_view = PersonalCicloView(records, test_env["pdf"], excel, user)
    # El combo debe incluir "Todas" y "Conductor Electrico"
    items = [ciclo_view.combo_esp.itemText(i) for i in range(ciclo_view.combo_esp.count())]
    assert "Todas" in items
    assert "Conductor Electrico" in items

    # EmpleadosView
    emp_view = EmpleadosView(records, excel, user)
    assert emp_view.model.rowCount() >= 1
    # Filtrar por legajo
    emp_view.txt_search.setText("12345")
    assert emp_view.proxy.rowCount() == 1


def test_validar_novedades_superpuestas(test_env):
    """Verifica que el sistema rechace cargar dos novedades activas superpuestas para el mismo empleado."""
    records = test_env["records"]
    user = test_env["user"]

    # Crear empleado 2001
    records.crear_empleado({
        "legajo": 2001,
        "apellidos_nombres": "TEST OVERLAP",
        "especialidad": "CONDUCTOR",
        "dotacion": "PLAZA",
        "turnos": "401",
        "franco": "Domingo",
        "franco_horario": "00:00 - 24:00",
        "jornada_horas": 6,
    }, user_id=user["id"])

    # 1. Crear novedad del 10/11/2026 al 15/11/2026
    nov_id = records.crear_novedad({
        "legajo": 2001,
        "apellidos_nombres": "TEST OVERLAP",
        "especialidad": "CONDUCTOR",
        "dotacion": "PLAZA",
        "turnos": "401",
        "franco": "Domingo",
        "novedad": "ENFERMEDAD",
        "fecha_inicio": "10/11/2026",
        "fecha_fin": "15/11/2026",
        "referencia_estacion": "Plaza",
        "supervisor": "ADMIN",
    }, user_id=user["id"])
    assert nov_id > 0

    # 2. Intentar crear novedad superpuesta (12/11/2026 al 18/11/2026) -> Debe lanzar ValueError
    with pytest.raises(ValueError, match="ya posee una novedad activa"):
        records.crear_novedad({
            "legajo": 2001,
            "apellidos_nombres": "TEST OVERLAP",
            "especialidad": "CONDUCTOR",
            "dotacion": "PLAZA",
            "turnos": "401",
            "franco": "Domingo",
            "novedad": "LICENCIA",
            "fecha_inicio": "12/11/2026",
            "fecha_fin": "18/11/2026",
            "referencia_estacion": "Plaza",
            "supervisor": "ADMIN",
        }, user_id=user["id"])

    # 3. Crear novedad en fechas no superpuestas (20/11/2026 al 25/11/2026) -> Debe ser exitosa
    nov_id_2 = records.crear_novedad({
        "legajo": 2001,
        "apellidos_nombres": "TEST OVERLAP",
        "especialidad": "CONDUCTOR",
        "dotacion": "PLAZA",
        "turnos": "401",
        "franco": "Domingo",
        "novedad": "LICENCIA",
        "fecha_inicio": "20/11/2026",
        "fecha_fin": "25/11/2026",
        "referencia_estacion": "Plaza",
        "supervisor": "ADMIN",
    }, user_id=user["id"])
    assert nov_id_2 > 0


def test_validar_reemplazo_disponible(test_env):
    """Verifica que no se pueda asignar a un personal que tenga novedad activa o que ya esté relevando otra novedad."""
    records = test_env["records"]
    user = test_env["user"]

    # Empleado titular A (3001), Titular B (3002), Reemplazo R (3003)
    records.crear_empleado({"legajo": 3001, "apellidos_nombres": "TITULAR A", "especialidad": "COND", "dotacion": "PC", "turnos": "101", "franco": "Dom"}, user_id=user["id"])
    records.crear_empleado({"legajo": 3002, "apellidos_nombres": "TITULAR B", "especialidad": "COND", "dotacion": "PC", "turnos": "102", "franco": "Dom"}, user_id=user["id"])
    records.crear_empleado({"legajo": 3003, "apellidos_nombres": "RELEVO R", "especialidad": "COND", "dotacion": "PC", "turnos": "CICLO A", "franco": "Dom"}, user_id=user["id"])

    # Novedad de Titular A (01/12/2026 al 10/12/2026)
    nov_a = records.crear_novedad({
        "legajo": 3001, "apellidos_nombres": "TITULAR A", "especialidad": "COND", "dotacion": "PC",
        "turnos": "101", "franco": "Dom", "novedad": "ENFERMEDAD",
        "fecha_inicio": "01/12/2026", "fecha_fin": "10/12/2026", "referencia_estacion": "PC",
    }, user_id=user["id"])

    # Asignar Relevo R a Titular A del 02/12 al 05/12
    records.asignar_reemplazo_rapido(nov_a, 3003, "02/12/2026", "05/12/2026", user_id=user["id"])

    # Novedad de Titular B (01/12/2026 al 10/12/2026)
    nov_b = records.crear_novedad({
        "legajo": 3002, "apellidos_nombres": "TITULAR B", "especialidad": "COND", "dotacion": "PC",
        "turnos": "102", "franco": "Dom", "novedad": "ENFERMEDAD",
        "fecha_inicio": "01/12/2026", "fecha_fin": "10/12/2026", "referencia_estacion": "PC",
    }, user_id=user["id"])

    # Intentar asignar Relevo R a Titular B en fechas superpuestas (04/12 al 08/12) -> Debe fallar
    with pytest.raises(ValueError, match="ya se encuentra cubriendo la novedad"):
        records.asignar_reemplazo_rapido(nov_b, 3003, "04/12/2026", "08/12/2026", user_id=user["id"])

    # Asignar Relevo R a Titular B en fechas no superpuestas (06/12 al 09/12) -> Debe funcionar
    records.asignar_reemplazo_rapido(nov_b, 3003, "06/12/2026", "09/12/2026", user_id=user["id"])


def test_catalogos_operativos_crud(test_env):
    """Verifica el CRUD completo (crear, modificar y alternar activo) de los catálogos operativos."""
    records = test_env["records"]
    user = test_env["user"]

    # 1. Tipos de Novedad
    t_id = records.crear_tipo("TEST TIPO", dias=5, exportar_descubierto=1, user_id=user["id"])
    assert t_id > 0
    records.actualizar_tipo(t_id, "TEST TIPO MOD", dias=7, exportar_descubierto=0, activo=0, user_id=user["id"])
    t_mod = next(t for t in records.listar_tipos(incluir_inactivos=True) if t["id"] == t_id)
    assert t_mod["nombre"] == "TEST TIPO MOD"
    assert t_mod["dias"] == 7
    assert t_mod["exportar_descubierto"] == 0
    assert t_mod["activo"] == 0

    # 2. Dotaciones
    d_id = records.crear_dotacion("TEST DOT", user_id=user["id"])
    assert d_id > 0
    records.actualizar_dotacion(d_id, "TEST DOT MOD", activo=1, user_id=user["id"])
    d_mod = next(d for d in records.listar_dotaciones(incluir_inactivos=True) if d["id"] == d_id)
    assert d_mod["nombre"] == "TEST DOT MOD"

    # 3. Personal de Estación
    pe_id = records.crear_personal_estacion("TEST PE", dotacion="TEST DOT MOD", user_id=user["id"])
    assert pe_id > 0
    records.actualizar_personal_estacion(pe_id, "TEST PE MOD", activo=0, dotacion="OTRA", cambiar_dotacion=True, user_id=user["id"])
    pe_mod = next(p for p in records.listar_personal_estacion(incluir_inactivos=True) if p["id"] == pe_id)
    assert pe_mod["nombre"] == "TEST PE MOD"
    assert pe_mod["dotacion"] == "OTRA"
    assert pe_mod["activo"] == 0


def test_buscador_vacio_y_cambios_turno(test_env):
    """Verifica que el buscador retorne listas vacías inmediatamente ante criterio vacío y detecte permutas."""
    records = test_env["records"]
    user = test_env["user"]

    # 1. Búsqueda con criterio vacío -> Debe retornar lista vacía
    res_vacio = records.buscar_turnos_con_cobertura("", "04/11/2026")
    assert res_vacio == []

    # 2. Importar itinerario Indice_45.xlsx para contar con turnos reales
    excel_path = Path("Indice_45.xlsx")
    if excel_path.exists():
        records.importar_itinerario_excel(str(excel_path))

    # Registrar titular 4001 con turno 401 en dotación PC
    records.crear_empleado({"legajo": 4001, "apellidos_nombres": "AGENTE 1", "especialidad": "Conductor Electrico", "dotacion": "PC", "turnos": "401", "franco": "Dom"}, user_id=user["id"])
    records.crear_empleado({"legajo": 4002, "apellidos_nombres": "AGENTE 2", "especialidad": "Conductor Electrico", "dotacion": "PC", "turnos": "402", "franco": "Dom"}, user_id=user["id"])

    records.crear_cambio({
        "legajo_1": 4001, "apellidos_nombres_1": "AGENTE 1", "especialidad_1": "Conductor Electrico", "dotacion_1": "PC", "turnos_1": "401", "franco_1": "Dom",
        "legajo_2": 4002, "apellidos_nombres_2": "AGENTE 2", "especialidad_2": "Conductor Electrico", "dotacion_2": "PC", "turnos_2": "402", "franco_2": "Dom",
        "fecha_cambio": "04/11/2026", "referencia_estacion": "PC", "supervisor": "ADMIN",
    }, user_id=user["id"])

    res = records.buscar_turnos_con_cobertura("401", "04/11/2026")
    assert len(res) >= 1
    item = res[0]
    assert item["estado_alerta"] == "CAMBIO"
    assert "AGENTE 2" in item["titular"]


def test_turnos_crud_operations(test_env):
    """Verifica el ciclo completo de CRUD para turnos de itinerario y sus vueltas."""
    records = test_env["records"]
    user = test_env["user"]

    # 1. Crear Turno con 2 vueltas
    vueltas_nuevas = [
        {"nro_vuelta": 1, "ref": "A", "nro_tren": "1001", "origen": "PC", "hora_salida": "06:00", "destino": "TY", "hora_llegada": "07:00", "observaciones_tramo": "Ida"},
        {"nro_vuelta": 2, "ref": "B", "nro_tren": "1002", "origen": "TY", "hora_salida": "07:15", "destino": "PC", "hora_llegada": "08:15", "observaciones_tramo": "Vuelta"},
    ]
    payload_nuevo = {
        "itinerario_vigente": "TEST_ITI_1",
        "codigo_turno": "999.1",
        "turno_base": "999",
        "subindice": "1",
        "tipo_turno": "DIESEL",
        "deposito_base": "TY",
        "toma_servicio": "05:45",
        "deje_servicio": "09:00",
        "duracion_jornada": "03:15",
        "activo": 1,
    }
    turno_id = records.crear_turno_itinerario(
        payload=payload_nuevo,
        vueltas=vueltas_nuevas,
        user_id=user["id"],
        windows_user="test_admin",
    )
    assert turno_id > 0

    # 2. Obtener turno completo y verificar vueltas
    turno_comp = records.obtener_turno_completo(turno_id)
    assert turno_comp is not None
    assert turno_comp["codigo_turno"] == "999.1"
    assert turno_comp["turno_base"] == "999"
    assert int(turno_comp["subindice"]) == 1
    assert turno_comp["tipo_turno"] == "DIESEL"
    vueltas = turno_comp.get("vueltas", [])
    assert len(vueltas) == 2
    assert vueltas[0]["nro_tren"] == "1001"
    assert vueltas[1]["nro_tren"] == "1002"

    # 3. Actualizar Turno y reemplazar con 1 vuelta
    vueltas_actualizadas = [
        {"nro_vuelta": 1, "ref": "C", "nro_tren": "2001", "origen": "PC", "hora_salida": "10:00", "destino": "LP", "hora_llegada": "11:15", "observaciones_tramo": "Directo"},
    ]
    payload_mod = {
        "itinerario_vigente": "TEST_ITI_1_MOD",
        "codigo_turno": "999.2",
        "turno_base": "999",
        "subindice": "2",
        "tipo_turno": "ELECTRICO",
        "deposito_base": "PC",
        "toma_servicio": "09:30",
        "deje_servicio": "12:00",
        "duracion_jornada": "02:30",
        "activo": 1,
    }
    records.actualizar_turno_itinerario(
        turno_id=turno_id,
        payload=payload_mod,
        vueltas=vueltas_actualizadas,
        user_id=user["id"],
        windows_user="test_admin",
    )

    turno_mod = records.obtener_turno_completo(turno_id)
    assert turno_mod["codigo_turno"] == "999.2"
    assert turno_mod["tipo_turno"] == "ELECTRICO"
    vueltas_mod = turno_mod.get("vueltas", [])
    assert len(vueltas_mod) == 1
    assert vueltas_mod[0]["nro_tren"] == "2001"

    # 4. Cambiar estado activo (toggle)
    records.cambiar_estado_turno(turno_id, activo=False, user_id=user["id"], windows_user="test_admin")
    turno_inactivo = records.obtener_turno_completo(turno_id)
    assert turno_inactivo["activo"] == 0

    # 5. Listar turnos admin
    lista_inactivos = records.listar_turnos_admin(incluir_inactivos=True)
    assert any(t["id"] == turno_id for t in lista_inactivos)

    lista_activos = records.listar_turnos_admin(incluir_inactivos=False)
    assert not any(t["id"] == turno_id for t in lista_activos)

    # 6. Eliminar definitivo y verificar borrado en cascada de vueltas
    records.eliminar_turno_definitivo(turno_id, user_id=user["id"], windows_user="test_admin")
    t_del = records.obtener_turno_completo(turno_id)
    assert t_del is None


def test_turno_separator_delegate(test_env):
    """Verifica que TurnoSeparatorDelegate pinte sin errores y detecte cambios de turno entre filas."""
    from PySide6.QtGui import QStandardItemModel, QStandardItem, QPainter, QPixmap
    from PySide6.QtWidgets import QStyleOptionViewItem
    from PySide6.QtCore import QRect
    from src.ui.views.buscador_turnos_view import TurnoSeparatorDelegate

    model = QStandardItemModel(3, 2)
    # Fila 0: Turno 401
    model.setItem(0, 0, QStandardItem("401.1"))
    model.setItem(0, 1, QStandardItem("1001"))
    # Fila 1: Turno 401
    model.setItem(1, 0, QStandardItem("401.1"))
    model.setItem(1, 1, QStandardItem("1003"))
    # Fila 2: Turno 402 (diferente turno)
    model.setItem(2, 0, QStandardItem("402.1"))
    model.setItem(2, 1, QStandardItem("1005"))

    delegate = TurnoSeparatorDelegate(get_tema_oscuro_fn=lambda: True)

    pixmap = QPixmap(300, 100)
    painter = QPainter(pixmap)
    opt = QStyleOptionViewItem()
    opt.rect = QRect(0, 0, 100, 30)

    # Fila 1 a Fila 2 es cambio de turno ("401.1" != "402.1") -> debe dibujar línea separadora sin lanzar excepción
    idx_row1 = model.index(1, 0)
    delegate.paint(painter, opt, idx_row1)

    # Fila 0 a Fila 1 no es cambio de turno ("401.1" == "401.1")
    idx_row0 = model.index(0, 0)
    delegate.paint(painter, opt, idx_row0)

    painter.end()


def test_forms_combobox_and_supervisor_disabled(test_env):
    """Verifica que combo_estacion sea no editable y supervisor esté deshabilitado en diálogos."""
    records = test_env["records"]
    user = test_env["user"]

    nov_dlg = NovedadDialog(records_service=records, usuario_actual=user)
    assert nov_dlg.combo_estacion.isEditable() is False
    assert nov_dlg.txt_supervisor.isReadOnly() is True
    assert nov_dlg.txt_supervisor.isEnabled() is False

    cam_dlg = CambioTurnoDialog(records_service=records, usuario_actual=user)
    assert cam_dlg.combo_estacion.isEditable() is False
    assert cam_dlg.txt_supervisor.isReadOnly() is True
    assert cam_dlg.txt_supervisor.isEnabled() is False


def test_backup_service_custom_directory_and_destination(test_env):
    """Verifica que BackupService soporte rutas personalizadas de almacenamiento y destino."""
    db = test_env["db"]
    tmp_path = test_env["tmp_path"]
    from src.services.backup_service import BackupService

    bs = BackupService(db)
    custom_dir = tmp_path / "custom_backups"
    bs.set_directorio_backups(str(custom_dir))
    assert Path(bs.ruta_directorio_backups()).resolve() == custom_dir.resolve()

    # Crear backup en carpeta específica
    external_dir = tmp_path / "external_backups"
    external_dir.mkdir(parents=True, exist_ok=True)
    ruta = bs.crear_backup(directorio_destino=str(external_dir))
    assert Path(ruta).exists()
    assert Path(ruta).parent.resolve() == external_dir.resolve()


def test_detalle_turno_dialog_and_lookup(test_env):
    """Verifica la consulta de turno completo por código y la apertura del DetalleTurnoDialog."""
    records = test_env["records"]
    user = test_env["user"]

    # Crear turno de prueba con vueltas
    vueltas = [
        {"nro_vuelta": 1, "ref": "A", "nro_tren": "3301", "origen": "PC", "hora_salida": "08:00", "destino": "TY", "hora_llegada": "08:50", "observaciones_tramo": "Tramo 1"},
    ]
    t_id = records.crear_turno_itinerario(
        payload={"codigo_turno": "888.1", "turno_base": "888", "subindice": "1", "tipo_turno": "ELECTRICO", "deposito_base": "PC", "toma_servicio": "07:30", "deje_servicio": "11:00", "duracion_jornada": "03:30", "itinerario_vigente": "45"},
        vueltas=vueltas,
        user_id=user["id"],
    )
    assert t_id > 0

    # Consultar por código
    t_data = records.obtener_turno_por_codigo("888.1", itinerario="45")
    assert t_data is not None
    assert t_data["codigo_turno"] == "888.1"
    assert len(t_data["vueltas"]) == 1
    assert t_data["vueltas"][0]["nro_tren"] == "3301"

    # Inicializar DetalleTurnoDialog
    from src.ui.views.dialogs.detalle_turno_dialog import DetalleTurnoDialog
    dlg = DetalleTurnoDialog(records, codigo_turno="888.1", itinerario_inicial="45")
    assert dlg.model.rowCount() == 1
    assert "888.1" in dlg.lbl_codigo.text()


