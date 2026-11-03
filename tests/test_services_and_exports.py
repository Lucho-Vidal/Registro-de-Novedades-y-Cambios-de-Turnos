"""Pruebas unitarias e integrales para la capa de servicios y generadores documentales."""

import os
import shutil
import tempfile
from datetime import date, datetime, timedelta
from pathlib import Path
import pytest

from src.data.database import DatabaseManager
from src.data.migrations import init_database_schema
from src.domain.enums import PERMISOS_BASE
from src.services.auth_service import AuthService, CuentaBloqueadaError
from src.services.backup_service import BackupService
from src.services.excel_service import ExcelService
from src.services.pdf_service import PDFService
from src.services.records_service import RecordsService
from src.services.word_service import WordService


@pytest.fixture
def temp_dir():
    d = tempfile.mkdtemp(prefix="reno_test_services_")
    yield d
    shutil.rmtree(d, ignore_errors=True)


@pytest.fixture
def test_db(temp_dir):
    db_path = os.path.join(temp_dir, "test_reno.sqlite")
    db = DatabaseManager(db_path, journal_mode="DELETE")
    init_database_schema(db)
    return db


def test_records_and_auth_services(test_db):
    auth_srv = AuthService(test_db)
    rec_srv = RecordsService(test_db)

    # 1. Admin inicial y autenticación
    admin_id = auth_srv.crear_administrador_inicial("admin", "Admin123*", legajo=1001)
    assert admin_id > 0
    assert auth_srv.tiene_permiso(admin_id, "novedades.crear")
    assert auth_srv.tiene_permiso(admin_id, "registros.modificar")

    user_info = auth_srv.autenticar("admin", "Admin123*")
    assert user_info is not None
    assert user_info["username"] == "admin"

    # Intento fallido
    assert auth_srv.autenticar("admin", "ClaveErronea") is None

    # 2. Empleados
    emp1_id = rec_srv.crear_empleado({
        "legajo": 5001,
        "apellidos_nombres": "PEREZ JUAN",
        "especialidad": "CONDUCTOR",
        "dotacion": "PLAZA",
        "turnos": "T1 06:00-12:00",
        "franco": "Domingo",
        "jornada_horas": 6,
    }, user_id=admin_id)
    assert emp1_id > 0

    # Empleado Ciclo
    emp_ciclo_id = rec_srv.crear_empleado({
        "legajo": 5002,
        "apellidos_nombres": "GARCIA CARLOS",
        "especialidad": "CONDUCTOR",
        "dotacion": "PLAZA",
        "turnos": "CICLO D",
        "franco": "Lunes",
        "franco_horario": "14:00",
        "jornada_horas": 6,
    }, user_id=admin_id)
    assert emp_ciclo_id > 0

    # 3. Tipos de novedad
    tipo_id = rec_srv.crear_tipo("LICENCIA MEDICA", dias=5, exportar_descubierto=1, user_id=admin_id)
    assert tipo_id > 0

    # 4. Novedad con tramo de cobertura
    hoy_str = datetime.now().strftime("%d/%m/%Y")
    fin_str = (datetime.now() + timedelta(days=4)).strftime("%d/%m/%Y")
    nov_id = rec_srv.crear_novedad({
        "legajo": 5001,
        "apellidos_nombres": "PEREZ JUAN",
        "especialidad": "CONDUCTOR",
        "dotacion": "PLAZA",
        "turnos": "T1 06:00-12:00",
        "franco": "Domingo",
        "novedad": "LICENCIA MEDICA",
        "fecha_inicio": hoy_str,
        "fecha_fin": fin_str,
        "referencia_estacion": "Plaza C",
        "supervisor": "GONZALEZ J",
        "observaciones": "Reposo 5 días",
    }, user_id=admin_id)
    assert nov_id > 0

    # Asignar tramo de reemplazo con personal de ciclo
    tramos = [{
        "legajo_reemplazo": 5002,
        "fecha_inicio": hoy_str,
        "fecha_fin": fin_str,
    }]
    rec_srv.guardar_reemplazos_novedad(nov_id, tramos, user_id=admin_id)
    tramos_guardados = rec_srv.listar_reemplazos_por_novedad(nov_id)
    assert len(tramos_guardados) == 1
    assert tramos_guardados[0]["legajo_reemplazo"] == 5002

    # Verificar cálculo de cobertura
    nov_dict = rec_srv.obtener_novedad(nov_id)
    info_cob = rec_srv.calcular_estado_cobertura(nov_dict, tramos_guardados)
    assert info_cob["estado"] == "Sí"

    # 5. Cambio de turno
    cambio_id = rec_srv.crear_cambio({
        "legajo_1": 5001,
        "apellidos_nombres_1": "PEREZ JUAN",
        "especialidad_1": "CONDUCTOR",
        "dotacion_1": "PLAZA",
        "turnos_1": "T1 06:00-12:00",
        "franco_1": "Domingo",
        "legajo_2": 5002,
        "apellidos_nombres_2": "GARCIA CARLOS",
        "especialidad_2": "CONDUCTOR",
        "dotacion_2": "PLAZA",
        "turnos_2": "CICLO D",
        "franco_2": "Lunes",
        "fecha_cambio": hoy_str,
        "referencia_estacion": "Plaza C",
        "supervisor": "GONZALEZ J",
        "observaciones": "Permuta autorizada",
    }, user_id=admin_id)
    assert cambio_id > 0

    # 6. Turnos sin cubrir
    descubiertos = rec_srv.obtener_turnos_sin_cubrir(hoy_str, dotacion="PLAZA")
    assert len(descubiertos) == 1
    assert descubiertos[0]["cubierto"] == "Sí"
    assert "GARCIA CARLOS" in descubiertos[0]["reemplazo"]

    # 7. Calendario mensual de ciclo
    dias_ciclo = rec_srv.obtener_calendario_mensual_ciclo(5002, datetime.now().year, datetime.now().month)
    assert len(dias_ciclo) >= 28
    dia_hoy_info = next(d for d in dias_ciclo if d["dia"] == datetime.now().day)
    assert dia_hoy_info["estado"] == "Reemplazando"
    assert "PEREZ JUAN" in dia_hoy_info["reemplaza_a"]

    # 8. Métricas del Dashboard
    metricas = rec_srv.obtener_metricas_dashboard(30)
    assert metricas["resumen"]["novedades_hoy"] >= 1
    assert metricas["resumen"]["cambios_hoy"] >= 1
    assert len(metricas["por_tipo"]) >= 1


def test_document_exports(test_db, temp_dir):
    rec_srv = RecordsService(test_db)
    excel_srv = ExcelService(test_db)
    pdf_srv = PDFService()
    word_srv = WordService()
    backup_srv = BackupService(test_db)

    # Crear datos base
    hoy_str = datetime.now().strftime("%d/%m/%Y")
    nov_id = rec_srv.crear_novedad({
        "legajo": 7001,
        "apellidos_nombres": "LOPEZ MARIA",
        "especialidad": "GUARDATREN",
        "dotacion": "TEMPERLEY",
        "turnos": "G1 08:00-15:00",
        "franco": "Sábado",
        "novedad": "FERIADO TRABAJADO",
        "fecha_inicio": hoy_str,
        "fecha_fin": hoy_str,
        "referencia_estacion": "Temperley",
        "supervisor": "SUAREZ R",
        "observaciones": "Servicio de guardia",
    })
    cambio_id = rec_srv.crear_cambio({
        "legajo_1": 7001,
        "apellidos_nombres_1": "LOPEZ MARIA",
        "especialidad_1": "GUARDATREN",
        "dotacion_1": "TEMPERLEY",
        "turnos_1": "G1 08:00-15:00",
        "franco_1": "Sábado",
        "legajo_2": 7002,
        "apellidos_nombres_2": "DIAZ PEDRO",
        "especialidad_2": "GUARDATREN",
        "dotacion_2": "TEMPERLEY",
        "turnos_2": "G2 15:00-22:00",
        "franco_2": "Domingo",
        "fecha_cambio": hoy_str,
        "referencia_estacion": "Temperley",
        "supervisor": "SUAREZ R",
        "observaciones": "Permuta",
    })

    nov = rec_srv.obtener_novedad(nov_id)
    cambio = rec_srv.obtener_cambio(cambio_id)
    metricas = rec_srv.obtener_metricas_dashboard(30)

    # 1. Excel export database
    xlsx_db = os.path.join(temp_dir, "export_db.xlsx")
    excel_srv.export_database(xlsx_db)
    assert os.path.exists(xlsx_db) and os.path.getsize(xlsx_db) > 1000

    # 2. Excel export dashboard detailed (revisión de requerimiento de usuario)
    xlsx_dash = os.path.join(temp_dir, "export_dashboard.xlsx")
    excel_srv.export_dashboard_detailed(metricas, xlsx_dash, periodo_dias=30)
    assert os.path.exists(xlsx_dash) and os.path.getsize(xlsx_dash) > 1000

    # 3. PDF Constancia Novedad
    pdf_nov = os.path.join(temp_dir, "constancia_novedad.pdf")
    pdf_srv.generar_constancia_novedad(nov, [], pdf_nov)
    assert os.path.exists(pdf_nov) and os.path.getsize(pdf_nov) > 1000

    # 4. PDF Constancia Cambio
    pdf_cam = os.path.join(temp_dir, "constancia_cambio.pdf")
    pdf_srv.generar_constancia_cambio(cambio, pdf_cam)
    assert os.path.exists(pdf_cam) and os.path.getsize(pdf_cam) > 1000

    # 5. PDF Informe Ejecutivo Dashboard
    pdf_dash = os.path.join(temp_dir, "informe_dashboard.pdf")
    pdf_srv.generar_informe_dashboard_pdf(metricas, pdf_dash, periodo_dias=30)
    assert os.path.exists(pdf_dash) and os.path.getsize(pdf_dash) > 1000

    # 6. Word Documentos (.docx)
    docx_cam = os.path.join(temp_dir, "solicitud_cambio.docx")
    word_srv.generar_solicitud_cambio_turno_docx(cambio, docx_cam)
    assert os.path.exists(docx_cam) and os.path.getsize(docx_cam) > 1000

    docx_lic = os.path.join(temp_dir, "solicitud_licencia.docx")
    word_srv.generar_solicitud_licencia_docx(nov, docx_lic)
    assert os.path.exists(docx_lic) and os.path.getsize(docx_lic) > 1000

    # 7. Backup Service
    backup_file = backup_srv.crear_backup()
    assert os.path.exists(backup_file)
    backups = backup_srv.listar_backups()
    assert len(backups) >= 1
