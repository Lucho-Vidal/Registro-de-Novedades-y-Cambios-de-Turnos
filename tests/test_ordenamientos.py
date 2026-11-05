"""Tests integrales para el Módulo de Ordenamientos Operativos, Paradas por Tren y Usuario de Windows."""

from __future__ import annotations

from datetime import datetime, timedelta
from pathlib import Path

import openpyxl
import pytest
from PySide6.QtWidgets import QApplication

from src.core.config import obtener_usuario_windows
from src.data.database import DatabaseManager
from src.data.migrations import init_database_schema
from src.services.excel_service import ExcelService
from src.services.pdf_service import PDFService
from src.services.records_service import RecordsService
from src.ui.views.buscador_turnos_view import BuscadorTurnosView
from src.ui.views.dialogs.ordenamiento_dialog import OrdenamientoDialog


@pytest.fixture(scope="module")
def qapp():
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


def test_usuario_windows_persistencia_en_todos_los_modulos(tmp_path: Path):
    """Verifica que novedades, cambios de turno y ordenamientos guarden automáticamente el usuario de Windows."""
    db = DatabaseManager(tmp_path / "test_win_user.sqlite")
    init_database_schema(db)
    svc = RecordsService(db)

    win_actual = obtener_usuario_windows()
    assert win_actual and win_actual != "Desconocido"

    # 1. Crear novedad sin pasar windows_user explícito
    nov_id = svc.crear_novedad({
        "legajo": 1001,
        "apellidos_nombres": "Juan Perez",
        "especialidad": "Conductor Eléctrico",
        "dotacion": "PC",
        "turnos": "401",
        "franco": "Domingo",
        "novedad": "Informe",
        "fecha_inicio": "05/11/2026",
        "fecha_fin": "05/11/2026",
        "referencia_estacion": "PC",
        "supervisor": "admin",
        "observaciones": "Test control usuario windows",
    })
    nov = svc.obtener_novedad(nov_id)
    assert nov is not None
    assert nov["usuario_windows"] == win_actual

    # 2. Crear cambio de turno sin pasar windows_user explícito
    cam_id = svc.crear_cambio({
        "legajo_1": 1001,
        "apellidos_nombres_1": "Juan Perez",
        "especialidad_1": "Conductor Eléctrico",
        "dotacion_1": "PC",
        "turnos_1": "401",
        "franco_1": "Domingo",
        "legajo_2": 1002,
        "apellidos_nombres_2": "Carlos Gomez",
        "especialidad_2": "Conductor Eléctrico",
        "dotacion_2": "PC",
        "turnos_2": "402",
        "franco_2": "Lunes",
        "fecha_cambio": "06/11/2026",
        "referencia_estacion": "PC",
        "supervisor": "admin",
        "observaciones": "Permuta test",
    })
    cam = svc.obtener_cambio(cam_id)
    assert cam is not None
    assert cam["usuario_windows"] == win_actual


def test_crud_e_importacion_excel_paradas_tren(tmp_path: Path):
    """Verifica el CRUD y la importación desde Excel de estaciones intermedias y horarios por tren."""
    db = DatabaseManager(tmp_path / "test_paradas.sqlite")
    init_database_schema(db)
    svc = RecordsService(db)

    # 1. Crear paradas manuales (CRUD)
    p1 = svc.crear_parada_tren({
        "nro_tren": "1293",
        "itinerario": "H45",
        "regimen": "LV",
        "orden_parada": 1,
        "estacion": "P. Constitución",
        "hora_paso": "08:00",
    })
    p2 = svc.crear_parada_tren({
        "nro_tren": "1293",
        "itinerario": "H45",
        "regimen": "LV",
        "orden_parada": 2,
        "estacion": "Temperley",
        "hora_paso": "08:25",
    })

    # 2. Actualizar parada
    svc.actualizar_parada_tren(p2, {
        "nro_tren": "1293",
        "itinerario": "H45",
        "regimen": "LV",
        "orden_parada": 2,
        "estacion": "Temperley",
        "hora_paso": "08:27",
    })

    paradas = svc.obtener_estaciones_y_horarios_por_tren("1293", "H45")
    assert len(paradas) == 2
    assert paradas[0] == {"estacion": "P. Constitución", "hora": "08:00"}
    assert paradas[1] == {"estacion": "Temperley", "hora": "08:27"}

    # 3. Importar desde Excel
    xlsx_path = tmp_path / "paradas_tren.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["NRO_TREN", "ORDEN", "ESTACION", "HORA_PASO", "ITINERARIO", "REGIMEN"])
    ws.append(["1293", 3, "Llavallol", "08:38", "H45", "LV"])
    ws.append(["1293", 4, "Ezeiza", "08:55", "H45", "LV"])
    wb.save(xlsx_path)

    insertados = svc.importar_paradas_tren_excel(str(xlsx_path))
    assert insertados == 2

    paradas_post = svc.obtener_estaciones_y_horarios_por_tren("1293", "H45")
    assert len(paradas_post) == 4
    assert paradas_post[-1] == {"estacion": "Ezeiza", "hora": "08:55"}

    # 4. Eliminar parada
    svc.eliminar_parada_tren(p1)
    assert len(svc.listar_paradas_tren(nro_tren="1293")) == 3


def test_ordenamientos_dotaciones_dinamicas_ventana_y_papelera(tmp_path: Path):
    """Verifica numeración por dotación (PC, LLV, TY...), ventana de edición, papelera y PDF/Excel."""
    db = DatabaseManager(tmp_path / "test_ordenamientos.sqlite")
    init_database_schema(db)
    svc = RecordsService(db)
    pdf_svc = PDFService()
    excel_svc = ExcelService(db)

    # Numeración correlativa dinámica según dotación
    assert svc.generar_siguiente_nro_orden("PC") == "PC 0000-0001"
    assert svc.generar_siguiente_nro_orden("LLV") == "LLV 0000-0001"
    assert svc.generar_siguiente_nro_orden("TY") == "TY 0000-0001"

    ord_llv_id = svc.crear_ordenamiento(
        datos={
            "fecha_emision": "05/11/2026",
            "hora_emision": "09:15",
            "legajo_personal": 5432,
            "nombre_personal": "Roberto Sanchez",
            "deposito": "LLV",
            "itinerario": "H45",
            "ciclo_o_turno": "CICLO 2",
            "regimen": "LV",
            "motivo_orden": "Relevo de urgencia por diagramación Control Trenes",
            "observacion_pie": "Presentarse en plataforma 2",
            "ref_pct": "Gomez (LLV)",
            "emisor": "sup_llv",
        },
        tramos=[
            {
                "nro_secuencia": 1,
                "ref": "A1",
                "nro_tren": "2303",
                "origen": "P. Constitución",
                "hora_salida": "10:05",
                "destino": "Llavallol",
                "hora_llegada": "10:42",
                "observaciones": "Servicio normal",
            },
            {
                "nro_secuencia": 2,
                "ref": "A2",
                "nro_tren": "2310",
                "origen": "Llavallol",
                "hora_salida": "11:00",
                "destino": "P. Constitución",
                "hora_llegada": "11:38",
                "observaciones": "A las órdenes",
            },
        ],
    )

    ord_completo = svc.obtener_ordenamiento_completo(ord_llv_id)
    assert ord_completo is not None
    assert ord_completo["nro_orden"] == "LLV 0000-0001"
    assert ord_completo["deposito"] == "LLV"
    assert ord_completo["usuario_windows"] == obtener_usuario_windows()
    assert len(ord_completo["tramos"]) == 2
    assert svc.generar_siguiente_nro_orden("LLV") == "LLV 0000-0002"

    # Exportar PDF duplicado A4 y Excel
    pdf_out = tmp_path / "orden_duplicado.pdf"
    pdf_svc.generar_orden_servicio_duplicado(ord_completo, ord_completo["tramos"], str(pdf_out))
    assert pdf_out.exists() and pdf_out.stat().st_size > 1000

    xlsx_out = tmp_path / "ordenamientos.xlsx"
    lista_ords = svc.listar_ordenamientos()
    excel_svc.export_ordenamientos(lista_ords, str(xlsx_out))
    assert xlsx_out.exists()

    # Simular antigüedad > 48 hs para probar bloqueo de ventana de edición sin permiso especial
    hace_48h = (datetime.now() - timedelta(hours=48)).strftime("%Y-%m-%d %H:%M:%S")
    with db.write_transaction() as conn:
        conn.execute("UPDATE ordenamientos SET created_at = ? WHERE id = ?", (hace_48h, ord_llv_id))

    with pytest.raises(ValueError, match="superó el tiempo permitido"):
        svc.actualizar_ordenamiento(
            ord_llv_id,
            datos=ord_completo,
            tramos=ord_completo["tramos"],
            bypass_window=False,
        )

    # Con bypass_window=True (permiso registros.modificar / Administrador) sí permite editar
    ord_completo["motivo_orden"] = "Motivo actualizado por Administrador"
    svc.actualizar_ordenamiento(
        ord_llv_id,
        datos=ord_completo,
        tramos=ord_completo["tramos"],
        bypass_window=True,
    )

    # Eliminación lógica a Papelera y Recuperación
    svc.eliminar_ordenamiento(ord_llv_id, bypass_window=True)
    assert len(svc.listar_ordenamientos()) == 0
    eliminados = svc.listar_eliminados()
    assert any(e["tipo"] == "Ordenamiento" and int(e["id"]) == ord_llv_id for e in eliminados)

    svc.recuperar_registro("Ordenamiento", ord_llv_id)
    assert len(svc.listar_ordenamientos()) == 1


def test_ordenamiento_visible_al_buscar_tren_en_buscador(qapp, tmp_path: Path):
    """Verifica que al buscar un tren en el Buscador se visualicen los ordenamientos operativos emitidos."""
    db = DatabaseManager(tmp_path / "test_buscador_ord.sqlite")
    init_database_schema(db)
    svc = RecordsService(db)
    excel_svc = ExcelService(db)

    # Crear un turno con el tren 2303
    svc.crear_turno_itinerario(
        payload={
            "codigo_turno": "401.1",
            "turno_base": "401",
            "subindice": 1,
            "tipo_turno": "DIAGRAMA_6D",
            "deposito_base": "PC",
            "toma_servicio": "09:30",
            "deje_servicio": "15:30",
            "duracion_jornada": "06:00",
            "itinerario_vigente": "H45",
        },
        vueltas=[
            {
                "nro_vuelta": 1,
                "ref": "R1",
                "nro_tren": "2303",
                "origen": "PC",
                "hora_salida": "10:05",
                "destino": "LLV",
                "hora_llegada": "10:42",
                "observaciones_tramo": "",
            }
        ],
    )

    hoy_ddmmyyyy = datetime.now().strftime("%d/%m/%Y")
    hoy_iso = datetime.now().strftime("%Y-%m-%d")

    svc.crear_ordenamiento(
        datos={
            "fecha_emision": hoy_ddmmyyyy,
            "hora_emision": "09:00",
            "legajo_personal": 7777,
            "nombre_personal": "Conductor Ordenado",
            "deposito": "PC",
            "itinerario": "H45",
            "ciclo_o_turno": "401.1",
            "regimen": "LV",
            "motivo_orden": "Ordenamiento especial sobre tren 2303",
            "observacion_pie": "",
            "ref_pct": "PC",
            "emisor": "admin",
        },
        tramos=[
            {
                "nro_secuencia": 1,
                "ref": "R1",
                "nro_tren": "2303",
                "origen": "PC",
                "hora_salida": "10:05",
                "destino": "LLV",
                "hora_llegada": "10:42",
                "observaciones": "Tramo ordenado",
            }
        ],
    )

    # 1. Verificar en servicio buscar_turnos_con_cobertura y buscar_ordenamientos_activos
    turnos_res = svc.buscar_turnos_con_cobertura("2303", hoy_iso)
    assert len(turnos_res) == 1
    assert turnos_res[0]["estado_alerta"] == "ORDENAMIENTO"
    assert "PC 0000-0001" in turnos_res[0]["cobertura"]

    ords_res = svc.buscar_ordenamientos_activos("2303", hoy_iso)
    assert len(ords_res) == 1
    assert ords_res[0]["nro_orden"] == "PC 0000-0001"

    # 2. Verificar en la vista BuscadorTurnosView
    view = BuscadorTurnosView(svc, excel_svc)
    view.txt_criterio.setText("2303")
    view.consultar()

    assert view.model_turnos.rowCount() == 1
    assert view.model_ordenamientos.rowCount() == 1
    fila_ord = view.model_ordenamientos.get_row(0)
    assert fila_ord["nro_orden"] == "PC 0000-0001"
    assert "2303" in fila_ord["trenes_involucrados"]
