"""Servicio unificado de importación y exportación de hojas de cálculo con OpenPyXL.

Soporta exportación con filtros, cobertura de ciclo, turnos sin cubrir,
reporte detallado del panel de control (Dashboard) e importación de catálogos y base operativa.
"""

from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple
from datetime import datetime, date

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from src.core.logger import get_logger
from src.data.database import DatabaseManager
from src.domain.validators import parsear_fecha

logger = get_logger("ExcelService")


def _write_sheet(
    workbook: openpyxl.Workbook,
    title: str,
    headers: List[str],
    rows: List[Dict[str, Any]],
    keys: Optional[List[str]] = None,
) -> openpyxl.worksheet.worksheet.Worksheet:
    ws = workbook.create_sheet(title=title)
    ws.append(headers)

    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    center_align = Alignment(horizontal="center", vertical="center")

    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align

    for row_data in rows:
        if keys:
            ws.append([row_data.get(k, "") for k in keys])
        elif hasattr(row_data, "keys"):
            ws.append([row_data[k] for k in row_data.keys()])
        else:
            ws.append(list(row_data))

    ws.freeze_panes = "A2"
    if ws.max_row > 1 and ws.max_column >= 1:
        ws.auto_filter.ref = ws.dimensions

    for index, header in enumerate(headers, start=1):
        ws.column_dimensions[get_column_letter(index)].width = max(12, min(40, len(str(header)) + 4))

    return ws


def _clave_fecha(value: Any) -> str:
    val = (str(value) if value is not None else "").strip()
    if len(val) < 10:
        return ""
    if val[4] == "-" and val[7] == "-":
        return val[0:4] + val[5:7] + val[8:10]
    if val[2] == "/" and val[5] == "/":
        return val[6:10] + val[3:5] + val[0:2]
    return ""


class ExcelService:
    """Gestor de importación y exportación de libros de trabajo Excel."""

    def __init__(self, db: DatabaseManager):
        self.db = db

    # ---------------------------------------------------------
    # Exportaciones de Base y Registros
    # ---------------------------------------------------------

    def export_database(
        self,
        output_path: str,
        fecha_desde: Optional[str] = None,
        fecha_hasta: Optional[str] = None,
        id_desde: Optional[int] = None,
        id_hasta: Optional[int] = None,
        tipo_novedad: Optional[str] = None,
        tables: Optional[Tuple[str, ...]] = None,
    ) -> str:
        out_file = Path(output_path)
        out_file.parent.mkdir(parents=True, exist_ok=True)
        selected_tables: Set[str] = set(tables or ("BASE", "TipoNovedad", "NOVEDADES", "Cambio de Turnos"))

        wb = openpyxl.Workbook()
        wb.remove(wb.active)  # Quita la hoja por defecto Sheet

        with self.db.read_connection() as conn:
            if "BASE" in selected_tables:
                base_rows = conn.execute(
                    """SELECT legajo, apellidos_nombres, especialidad, dotacion, turnos, franco, franco_horario, jornada_horas
                       FROM empleados WHERE activo=1 ORDER BY apellidos_nombres"""
                ).fetchall()
                _write_sheet(
                    wb,
                    "BASE",
                    ["LEGAJO", "APELLIDOS Y NOMBRES", "ESPECIALIDAD", "DOTACION", "TURNOS", "FRANCO", "HORARIO FRANCO", "JORNADA HS"],
                    [dict(r) for r in base_rows],
                    ["legajo", "apellidos_nombres", "especialidad", "dotacion", "turnos", "franco", "franco_horario", "jornada_horas"],
                )

            if "TipoNovedad" in selected_tables:
                tipos = conn.execute(
                    "SELECT nombre, dias, exportar_descubierto FROM tipos_novedad WHERE activo=1 ORDER BY nombre"
                ).fetchall()
                _write_sheet(
                    wb,
                    "TipoNovedad",
                    ["TIPO DE NOVEDAD", "DURACION DIAS", "EXPORTAR DESCUBIERTO"],
                    [dict(r) for r in tipos],
                    ["nombre", "dias", "exportar_descubierto"],
                )

            if "PersonalEstacion" in selected_tables:
                pe_rows = conn.execute(
                    "SELECT nombre, dotacion, activo FROM personal_estacion WHERE activo=1 ORDER BY nombre"
                ).fetchall()
                _write_sheet(
                    wb,
                    "PersonalEstacion",
                    ["NOMBRE", "DOTACION", "ACTIVO"],
                    [dict(r) for r in pe_rows],
                    ["nombre", "dotacion", "activo"],
                )

            # Filtros comunes para tablas operativas
            common_conds = ["activo = 1"]
            common_params: List[Any] = []
            if id_desde is not None:
                common_conds.append("id >= ?")
                common_params.append(id_desde)
            if id_hasta is not None:
                common_conds.append("id <= ?")
                common_params.append(id_hasta)

            if "NOVEDADES" in selected_tables:
                nov_conds = list(common_conds)
                nov_params = list(common_params)
                if tipo_novedad:
                    nov_conds.append("novedad = ?")
                    nov_params.append(tipo_novedad)

                novedades = conn.execute(
                    f"""SELECT id, registrado_en, legajo, apellidos_nombres, especialidad, dotacion,
                               turnos, franco, novedad, fecha_inicio, fecha_fin, referencia_estacion,
                               supervisor, observaciones, usuario_windows
                        FROM novedades WHERE {' AND '.join(nov_conds)} ORDER BY id DESC""",
                    nov_params,
                ).fetchall()

                # Reemplazos y cobertura
                reemplazos = conn.execute(
                    """SELECT id, novedad_id, legajo_reemplazo, apellidos_nombres_reemplazo,
                              turnos_reemplazo, franco_reemplazo, fecha_inicio, fecha_fin
                       FROM novedad_reemplazos WHERE activo = 1 ORDER BY id ASC"""
                ).fetchall()
                por_novedad: Dict[int, List[Dict[str, Any]]] = {}
                for r in reemplazos:
                    por_novedad.setdefault(r["novedad_id"], []).append(dict(r))

                nov_dict_list = []
                for n in novedades:
                    item = dict(n)
                    # Filtrado por fecha registrado_en si fue solicitado
                    if fecha_desde or fecha_hasta:
                        c_fecha = _clave_fecha(item["registrado_en"])
                        if fecha_desde and c_fecha < _clave_fecha(fecha_desde):
                            continue
                        if fecha_hasta and c_fecha > _clave_fecha(fecha_hasta):
                            continue

                    tramos = por_novedad.get(item["id"], [])
                    if str(item.get("turnos") or "").upper().startswith("CICLO"):
                        item["cubierto"] = "N/A (Ciclo)"
                        item["detalle_reemplazos"] = ""
                    elif not tramos:
                        item["cubierto"] = "No"
                        item["detalle_reemplazos"] = ""
                    else:
                        item["cubierto"] = "Sí" if len(tramos) == 1 else "Parcial"
                        item["detalle_reemplazos"] = " | ".join(
                            f"{t['legajo_reemplazo']} - {t['apellidos_nombres_reemplazo']} [{t['fecha_inicio']} a {t['fecha_fin'] or 'abierto'}]"
                            for t in tramos
                        )
                    nov_dict_list.append(item)

                _write_sheet(
                    wb,
                    "NOVEDADES",
                    [
                        "ID", "REGISTRADO EN", "LEGAJO", "APELLIDOS Y NOMBRES", "ESPECIALIDAD", "DOTACION",
                        "TURNOS", "FRANCO", "NOVEDAD", "FECHA INICIO", "FECHA FIN", "CUBIERTO", "REEMPLAZOS",
                        "REFERENCIA ESTACION", "SUPERVISOR", "OBSERVACIONES", "USUARIO",
                    ],
                    nov_dict_list,
                    [
                        "id", "registrado_en", "legajo", "apellidos_nombres", "especialidad", "dotacion",
                        "turnos", "franco", "novedad", "fecha_inicio", "fecha_fin", "cubierto", "detalle_reemplazos",
                        "referencia_estacion", "supervisor", "observaciones", "usuario_windows",
                    ],
                )

            if "Cambio de Turnos" in selected_tables:
                cambios = conn.execute(
                    f"""SELECT id, registrado_en, legajo_1, apellidos_nombres_1, especialidad_1,
                               dotacion_1, turnos_1, franco_1, legajo_2, apellidos_nombres_2,
                               especialidad_2, dotacion_2, turnos_2, franco_2, fecha_cambio,
                               referencia_estacion, supervisor, observaciones, usuario_windows
                        FROM cambios_turno WHERE {' AND '.join(common_conds)} ORDER BY id DESC""",
                    common_params,
                ).fetchall()

                cambios_filtered = []
                for c in cambios:
                    item = dict(c)
                    if fecha_desde or fecha_hasta:
                        c_fecha = _clave_fecha(item["registrado_en"])
                        if fecha_desde and c_fecha < _clave_fecha(fecha_desde):
                            continue
                        if fecha_hasta and c_fecha > _clave_fecha(fecha_hasta):
                            continue
                    cambios_filtered.append(item)

                _write_sheet(
                    wb,
                    "Cambio de Turnos",
                    [
                        "ID", "REGISTRADO EN", "LEGAJO 1", "APELLIDOS Y NOMBRES 1", "ESPECIALIDAD 1", "DOTACION 1",
                        "TURNOS 1", "FRANCO 1", "LEGAJO 2", "APELLIDOS Y NOMBRES 2", "ESPECIALIDAD 2", "DOTACION 2",
                        "TURNOS 2", "FRANCO 2", "FECHA CAMBIO", "REFERENCIA ESTACION", "SUPERVISOR",
                        "OBSERVACIONES", "USUARIO",
                    ],
                    cambios_filtered,
                    [
                        "id", "registrado_en", "legajo_1", "apellidos_nombres_1", "especialidad_1", "dotacion_1",
                        "turnos_1", "franco_1", "legajo_2", "apellidos_nombres_2", "especialidad_2", "dotacion_2",
                        "turnos_2", "franco_2", "fecha_cambio", "referencia_estacion", "supervisor",
                        "observaciones", "usuario_windows",
                    ],
                )

        wb.save(out_file)
        return str(out_file)

    def export_turnos_sin_cubrir(
        self,
        output_path: str,
        filas_descubiertos: List[Dict[str, Any]],
        fecha_consulta: str,
        dotacion: Optional[str] = None,
    ) -> str:
        out_file = Path(output_path)
        out_file.parent.mkdir(parents=True, exist_ok=True)

        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        title = f"Descubiertos {fecha_consulta.replace('/', '-')}"[:31]
        _write_sheet(
            wb,
            title,
            [
                "REGISTRADO EN", "LEGAJO", "APELLIDOS Y NOMBRES", "ESPECIALIDAD", "DOTACION",
                "TURNOS", "FRANCO", "NOVEDAD", "FECHA INICIO", "FECHA FIN", "CUBIERTO",
                "REEMPLAZO HOY", "ESTADO COBERTURA", "DETALLE REEMPLAZOS", "OBSERVACIONES",
            ],
            filas_descubiertos,
            [
                "registrado_en", "legajo", "apellidos_nombres", "especialidad", "dotacion",
                "turnos", "franco", "novedad", "fecha_inicio", "fecha_fin", "cubierto",
                "reemplazo", "cobertura_resumen", "reemplazos_detalle", "observaciones",
            ],
        )
        wb.save(out_file)
        return str(out_file)

    def export_ordenamientos(
        self,
        ordenamientos: List[Dict[str, Any]],
        output_path: str,
    ) -> str:
        """Exporta el listado de Ordenamientos Operativos y sus tramos a un libro Excel."""
        out_file = Path(output_path)
        out_file.parent.mkdir(parents=True, exist_ok=True)

        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        _write_sheet(
            wb,
            "Ordenamientos",
            [
                "ID", "N° ORDEN", "FECHA EMISIÓN", "HORA EMISIÓN", "LEGAJO",
                "NOMBRE Y APELLIDO", "DOTACIÓN", "ITINERARIO", "TURNO / CICLO",
                "RÉGIMEN", "TRENES", "MOTIVO DE LA ORDEN", "REF ESTACIÓN",
                "EMISOR", "USUARIO WINDOWS", "OBSERVACIONES",
            ],
            ordenamientos,
            [
                "id", "nro_orden", "fecha_emision", "hora_emision", "legajo_personal",
                "nombre_personal", "deposito", "itinerario", "ciclo_o_turno",
                "regimen", "trenes_involucrados", "motivo_orden", "ref_pct",
                "emisor", "usuario_windows", "observacion_pie",
            ],
        )

        # Hoja de detalle de tramos si están disponibles en la BD
        with self.db.read_connection() as conn:
            ids = [int(o["id"]) for o in ordenamientos if o.get("id")]
            if ids:
                placeholders = ",".join("?" for _ in ids)
                tramos_rows = conn.execute(
                    f"""
                    SELECT o.nro_orden, o.fecha_emision, o.legajo_personal, o.nombre_personal,
                           o.deposito, ot.nro_secuencia, ot.ref, ot.nro_tren,
                           ot.origen, ot.hora_salida, ot.destino, ot.hora_llegada, ot.observaciones
                    FROM ordenamiento_tramos ot
                    JOIN ordenamientos o ON o.id = ot.ordenamiento_id
                    WHERE o.id IN ({placeholders})
                    ORDER BY o.id ASC, ot.nro_secuencia ASC
                    """,
                    ids,
                ).fetchall()
                _write_sheet(
                    wb,
                    "Detalle de Tramos",
                    [
                        "N° ORDEN", "FECHA", "LEGAJO", "AGENTE", "DOTACIÓN",
                        "SECUENCIA", "REF", "N° TREN", "ORIGEN", "HORA SALIDA",
                        "DESTINO", "HORA LLEGADA", "OBSERVACIONES",
                    ],
                    [dict(r) for r in tramos_rows],
                    [
                        "nro_orden", "fecha_emision", "legajo_personal", "nombre_personal", "deposito",
                        "nro_secuencia", "ref", "nro_tren", "origen", "hora_salida",
                        "destino", "hora_llegada", "observaciones",
                    ],
                )

        wb.save(out_file)
        return str(out_file)

    # ---------------------------------------------------------
    # Exportación Detallada del Dashboard (Requerimiento Usuario)
    # ---------------------------------------------------------

    def export_dashboard_detailed(
        self,
        metricas: Dict[str, Any],
        output_path: str,
        periodo_dias: int = 30,
    ) -> str:
        """Exporta métricas completas del dashboard: KPIs, desglose por tipo, dotación, día y serie temporal."""
        out_file = Path(output_path)
        out_file.parent.mkdir(parents=True, exist_ok=True)

        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        # 1. Resumen Ejecutivo
        resumen = metricas.get("resumen", {})
        ws_resumen = wb.create_sheet(title="Resumen Ejecutivo")
        ws_resumen.append(["MÉTRICA", "VALOR"])
        ws_resumen.append(["Novedades Registradas Hoy", resumen.get("novedades_hoy", 0)])
        ws_resumen.append(["Cambios de Turno Registrados Hoy", resumen.get("cambios_hoy", 0)])
        ws_resumen.append(["Total de Novedades del Mes", resumen.get("novedades_mes", 0)])
        ws_resumen.append(["Total de Cambios del Mes", resumen.get("cambios_mes", 0)])
        ws_resumen.append(["Volumen Total Operativo del Mes", resumen.get("total_mes", 0)])
        ws_resumen.append(["Período Analizado de Tendencia", f"Últimos {periodo_dias} días"])
        ws_resumen.append(["Fecha y Hora de Generación", datetime.now().strftime("%d/%m/%Y %H:%M:%S")])
        ws_resumen.column_dimensions["A"].width = 36
        ws_resumen.column_dimensions["B"].width = 20

        # Estilo encabezado
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
        for col in (1, 2):
            cell = ws_resumen.cell(row=1, column=col)
            cell.font = header_font
            cell.fill = header_fill

        # 2. Desglose por Tipo de Novedad
        por_tipo = metricas.get("por_tipo", [])
        total_tipo = sum(v for _, v in por_tipo) or 1
        filas_tipo = [
            {"tipo": t, "cantidad": c, "porcentaje": f"{(c / total_tipo) * 100:.1f}%"}
            for t, c in por_tipo
        ]
        _write_sheet(
            wb,
            "Novedades por Tipo",
            ["TIPO DE NOVEDAD", "CANTIDAD", "PARTICIPACION"],
            filas_tipo,
            ["tipo", "cantidad", "porcentaje"],
        )

        # 3. Desglose por Dotación
        por_dotacion = metricas.get("por_dotacion", [])
        total_dot = sum(v for _, v in por_dotacion) or 1
        filas_dot = [
            {"dotacion": d, "cantidad": c, "porcentaje": f"{(c / total_dot) * 100:.1f}%"}
            for d, c in por_dotacion
        ]
        _write_sheet(
            wb,
            "Registros por Dotacion",
            ["DOTACION", "CANTIDAD", "PARTICIPACION"],
            filas_dot,
            ["dotacion", "cantidad", "porcentaje"],
        )

        # 4. Desglose por Día de Semana
        por_dia = metricas.get("por_dia_semana", [])
        filas_dia = [{"dia": d, "cantidad": c} for d, c in por_dia]
        _write_sheet(
            wb,
            "Por Dia de Semana",
            ["DIA DE LA SEMANA", "CANTIDAD REGISTROS"],
            filas_dia,
            ["dia", "cantidad"],
        )

        # 5. Tendencia Diaria (Serie Temporal)
        tendencia = metricas.get("tendencia", [])
        filas_tendencia = [
            {
                "fecha": f,
                "novedades": n,
                "cambios": c,
                "total": n + c,
            }
            for f, n, c in tendencia
        ]
        _write_sheet(
            wb,
            "Tendencia Diaria",
            ["FECHA", "NOVEDADES", "CAMBIOS DE TURNO", "TOTAL DIARIO"],
            filas_tendencia,
            ["fecha", "novedades", "cambios", "total"],
        )

        wb.save(out_file)
        return str(out_file)

    def export_ordenamientos(
        self,
        ordenamientos: List[Dict[str, Any]],
        output_path: str,
    ) -> str:
        """Exporta el listado de Ordenamientos Operativos y sus tramos a un libro Excel."""
        out_file = Path(output_path)
        out_file.parent.mkdir(parents=True, exist_ok=True)

        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        _write_sheet(
            wb,
            "Ordenamientos",
            [
                "ID", "N° ORDEN", "FECHA EMISIÓN", "HORA EMISIÓN", "LEGAJO",
                "NOMBRE Y APELLIDO", "DOTACIÓN", "ITINERARIO", "TURNO / CICLO",
                "RÉGIMEN", "TRENES", "MOTIVO DE LA ORDEN", "REF ESTACIÓN",
                "EMISOR", "USUARIO WINDOWS", "OBSERVACIONES",
            ],
            ordenamientos,
            [
                "id", "nro_orden", "fecha_emision", "hora_emision", "legajo_personal",
                "nombre_personal", "deposito", "itinerario", "ciclo_o_turno",
                "regimen", "trenes_involucrados", "motivo_orden", "ref_pct",
                "emisor", "usuario_windows", "observacion_pie",
            ],
        )

        # Hoja de detalle de tramos si están disponibles en la BD
        with self.db.read_connection() as conn:
            ids = [int(o["id"]) for o in ordenamientos if o.get("id")]
            if ids:
                placeholders = ",".join("?" for _ in ids)
                tramos_rows = conn.execute(
                    f"""
                    SELECT o.nro_orden, o.fecha_emision, o.legajo_personal, o.nombre_personal,
                           o.deposito, ot.nro_secuencia, ot.ref, ot.nro_tren,
                           ot.origen, ot.hora_salida, ot.destino, ot.hora_llegada, ot.observaciones
                    FROM ordenamiento_tramos ot
                    JOIN ordenamientos o ON o.id = ot.ordenamiento_id
                    WHERE o.id IN ({placeholders})
                    ORDER BY o.id ASC, ot.nro_secuencia ASC
                    """,
                    ids,
                ).fetchall()
                _write_sheet(
                    wb,
                    "Detalle de Tramos",
                    [
                        "N° ORDEN", "FECHA", "LEGAJO", "AGENTE", "DOTACIÓN",
                        "SECUENCIA", "REF", "N° TREN", "ORIGEN", "HORA SALIDA",
                        "DESTINO", "HORA LLEGADA", "OBSERVACIONES",
                    ],
                    [dict(r) for r in tramos_rows],
                    [
                        "nro_orden", "fecha_emision", "legajo_personal", "nombre_personal", "deposito",
                        "nro_secuencia", "ref", "nro_tren", "origen", "hora_salida",
                        "destino", "hora_llegada", "observaciones",
                    ],
                )

        wb.save(out_file)
        return str(out_file)

    # ---------------------------------------------------------
    # Importaciones de Excel
    # ---------------------------------------------------------

    def import_base_empleados(self, file_path: str) -> Tuple[int, int]:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        sheet_name = "BASE" if "BASE" in wb.sheetnames else wb.sheetnames[0]
        ws = wb[sheet_name]

        headers = [str(cell.value).strip().upper() if cell.value is not None else "" for cell in ws[1]]
        header_map = {name: idx for idx, name in enumerate(headers) if name}

        legajo_idx = header_map.get("LEGAJO", 0)
        nombre_idx = header_map.get("APELLIDOS Y NOMBRES", 1)
        esp_idx = header_map.get("ESPECIALIDAD", 2)
        dot_idx = header_map.get("DOTACION", 3)
        turnos_idx = header_map.get("TURNOS", 4)
        franco_idx = header_map.get("FRANCO", 5)
        franco_h_idx = header_map.get("HORARIO FRANCO", header_map.get("FRANCO_HORARIO"))
        jornada_idx = header_map.get("JORNADA HS", header_map.get("JORNADA_HORAS"))

        insertados = 0
        actualizados = 0
        now = self.db.now()

        with self.db.write_transaction() as conn:
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row or row[legajo_idx] is None:
                    continue
                try:
                    legajo = int(str(row[legajo_idx]).strip())
                except (ValueError, TypeError):
                    continue

                nombre = str(row[nombre_idx] or "").strip()
                if not nombre:
                    continue

                esp = str(row[esp_idx] or "").strip() if esp_idx < len(row) else ""
                dot = str(row[dot_idx] or "").strip() if dot_idx < len(row) else ""
                turnos = str(row[turnos_idx] or "").strip() if turnos_idx < len(row) else ""
                franco = str(row[franco_idx] or "").strip() if franco_idx < len(row) else ""
                franco_h = str(row[franco_h_idx] or "").strip() if (franco_h_idx is not None and franco_h_idx < len(row)) else ""
                jornada_val = 6
                if jornada_idx is not None and jornada_idx < len(row):
                    try:
                        jornada_val = int(row[jornada_idx] or 6)
                    except (ValueError, TypeError):
                        jornada_val = 6

                cur = conn.execute("SELECT id FROM empleados WHERE legajo=?", (legajo,))
                existente = cur.fetchone()

                if existente:
                    conn.execute(
                        """UPDATE empleados SET apellidos_nombres=?, especialidad=?, dotacion=?,
                                                turnos=?, franco=?, franco_horario=?, jornada_horas=?,
                                                actualizado_en=?, activo=1
                           WHERE id=?""",
                        (nombre, esp, dot, turnos, franco, franco_h, jornada_val, now, existente[0]),
                    )
                    actualizados += 1
                else:
                    conn.execute(
                        """INSERT INTO empleados (legajo, apellidos_nombres, especialidad, dotacion,
                                                  turnos, franco, franco_horario, jornada_horas, actualizado_en, activo)
                           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 1)""",
                        (legajo, nombre, esp, dot, turnos, franco, franco_h, jornada_val, now),
                    )
                    insertados += 1

        return insertados, actualizados

    def import_catalog(
        self,
        table_name: str,
        column_name: str,
        file_path: str,
        clear_existing: bool = False,
    ) -> int:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        ws = wb.active
        items = set()
        for row in ws.iter_rows(values_only=True):
            if row and row[0] is not None:
                val = str(row[0]).strip()
                if val and val.lower() not in ("nombre", "tipo", "dotacion", "tipo de novedad"):
                    items.add(val)

        with self.db.write_transaction() as conn:
            if clear_existing:
                conn.execute(f"UPDATE {table_name} SET activo = 0")
            for item in items:
                conn.execute(
                    f"INSERT INTO {table_name} ({column_name}, activo) VALUES (?, 1) "
                    f"ON CONFLICT({column_name}) DO UPDATE SET activo = 1",
                    (item,),
                )
        return len(items)

    def import_personal_estacion_sheet(self, file_path: str, clear_existing: bool = False) -> int:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        ws = wb.active
        items = []
        for row in ws.iter_rows(values_only=True):
            if not row or row[0] is None:
                continue
            nombre = str(row[0]).strip()
            if not nombre or nombre.lower() in ("nombre", "personal", "personal de estación", "personal estacion"):
                continue
            dotacion = str(row[1]).strip() if len(row) > 1 and row[1] is not None else None
            items.append((nombre, dotacion))

        with self.db.write_transaction() as conn:
            if clear_existing:
                conn.execute("UPDATE personal_estacion SET activo = 0")
            for nombre, dotacion in items:
                conn.execute(
                    """INSERT INTO personal_estacion (nombre, dotacion, activo) VALUES (?, ?, 1)
                       ON CONFLICT(nombre) DO UPDATE SET
                           dotacion = COALESCE(excluded.dotacion, personal_estacion.dotacion),
                           activo = 1""",
                    (nombre, dotacion),
                )
        return len(items)

    # ---------------------------------------------------------
    # Exportaciones de Calendario de Ciclo (Individual y Masivo)
    # ---------------------------------------------------------

    def export_calendario_ciclo_individual(
        self,
        output_path: str,
        empleado: Dict[str, Any],
        dias_datos: List[Dict[str, Any]],
        mes: int,
        anio: int,
    ) -> str:
        out_file = Path(output_path)
        out_file.parent.mkdir(parents=True, exist_ok=True)

        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        legajo = empleado.get("legajo", "")
        nombre = empleado.get("apellidos_nombres", "")
        ws = wb.create_sheet(title=f"Ciclo_{legajo}"[:31])

        # Encabezado informativo del empleado
        ws.append(["CRONOGRAMA MENSUAL DE CICLO", f"Período: {mes:02d}/{anio}"])
        ws.append(["Agente:", f"{nombre} (Legajo: {legajo})"])
        ws.append(["Dotación / Especialidad:", f"{empleado.get('dotacion', '')} - {empleado.get('especialidad', '')}"])
        ws.append(["Turno / Franco CCT:", f"{empleado.get('turnos', '')} | Franco: {empleado.get('franco', '')} ({empleado.get('franco_horario', '')}) | Jornada: {empleado.get('jornada_horas', 6)} hs"])
        ws.append([])  # Fila en blanco

        # Título principal
        title_font = Font(name="Calibri", size=13, bold=True, color="1F497D")
        ws.cell(row=1, column=1).font = title_font

        headers = [
            "DÍA", "FECHA", "DÍA SEM.", "ESTADO", "HORA DISPONIBLE",
            "TOMA", "DEJA", "HS TRABAJADAS", "HS SEMANA", "REEMPLAZA A / MOTIVO", "OBSERVACIONES",
        ]
        keys = [
            "dia", "fecha", "dia_semana", "estado", "hora_disponible",
            "toma", "deja", "horas_trabajadas", "hs_semana", "reemplaza_a", "observaciones",
        ]

        ws.append(headers)
        header_row_idx = 6

        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
        center_align = Alignment(horizontal="center", vertical="center")

        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=header_row_idx, column=col_idx)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center_align

        fill_franco = PatternFill(start_color="FFF9C4", end_color="FFF9C4", fill_type="solid")
        fill_relevo = PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid")
        fill_disponible = PatternFill(start_color="F0F7FF", end_color="F0F7FF", fill_type="solid")

        for row_data in dias_datos:
            row_vals = [row_data.get(k, "") for k in keys]
            ws.append(row_vals)
            curr_row = ws.max_row
            estado = str(row_data.get("estado", ""))
            row_fill = None
            if "Franco" in estado:
                row_fill = fill_franco
            elif "Reemplazando" in estado or "Relevo" in estado:
                row_fill = fill_relevo
            elif estado == "Disponible":
                row_fill = fill_disponible

            if row_fill:
                for c in range(1, len(headers) + 1):
                    ws.cell(row=curr_row, column=c).fill = row_fill

        ws.freeze_panes = f"A{header_row_idx + 1}"
        for index, header in enumerate(headers, start=1):
            ws.column_dimensions[get_column_letter(index)].width = max(12, min(36, len(str(header)) + 4))

        wb.save(out_file)
        return str(out_file)

    def export_ordenamientos(
        self,
        ordenamientos: List[Dict[str, Any]],
        output_path: str,
    ) -> str:
        """Exporta el listado de Ordenamientos Operativos y sus tramos a un libro Excel."""
        out_file = Path(output_path)
        out_file.parent.mkdir(parents=True, exist_ok=True)

        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        _write_sheet(
            wb,
            "Ordenamientos",
            [
                "ID", "N° ORDEN", "FECHA EMISIÓN", "HORA EMISIÓN", "LEGAJO",
                "NOMBRE Y APELLIDO", "DOTACIÓN", "ITINERARIO", "TURNO / CICLO",
                "RÉGIMEN", "TRENES", "MOTIVO DE LA ORDEN", "REF ESTACIÓN",
                "EMISOR", "USUARIO WINDOWS", "OBSERVACIONES",
            ],
            ordenamientos,
            [
                "id", "nro_orden", "fecha_emision", "hora_emision", "legajo_personal",
                "nombre_personal", "deposito", "itinerario", "ciclo_o_turno",
                "regimen", "trenes_involucrados", "motivo_orden", "ref_pct",
                "emisor", "usuario_windows", "observacion_pie",
            ],
        )

        # Hoja de detalle de tramos si están disponibles en la BD
        with self.db.read_connection() as conn:
            ids = [int(o["id"]) for o in ordenamientos if o.get("id")]
            if ids:
                placeholders = ",".join("?" for _ in ids)
                tramos_rows = conn.execute(
                    f"""
                    SELECT o.nro_orden, o.fecha_emision, o.legajo_personal, o.nombre_personal,
                           o.deposito, ot.nro_secuencia, ot.ref, ot.nro_tren,
                           ot.origen, ot.hora_salida, ot.destino, ot.hora_llegada, ot.observaciones
                    FROM ordenamiento_tramos ot
                    JOIN ordenamientos o ON o.id = ot.ordenamiento_id
                    WHERE o.id IN ({placeholders})
                    ORDER BY o.id ASC, ot.nro_secuencia ASC
                    """,
                    ids,
                ).fetchall()
                _write_sheet(
                    wb,
                    "Detalle de Tramos",
                    [
                        "N° ORDEN", "FECHA", "LEGAJO", "AGENTE", "DOTACIÓN",
                        "SECUENCIA", "REF", "N° TREN", "ORIGEN", "HORA SALIDA",
                        "DESTINO", "HORA LLEGADA", "OBSERVACIONES",
                    ],
                    [dict(r) for r in tramos_rows],
                    [
                        "nro_orden", "fecha_emision", "legajo_personal", "nombre_personal", "deposito",
                        "nro_secuencia", "ref", "nro_tren", "origen", "hora_salida",
                        "destino", "hora_llegada", "observaciones",
                    ],
                )

        wb.save(out_file)
        return str(out_file)

    def export_calendario_ciclo_todos(
        self,
        output_path: str,
        records_service: Any,
        mes: int,
        anio: int,
    ) -> str:
        out_file = Path(output_path)
        out_file.parent.mkdir(parents=True, exist_ok=True)

        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        empleados = records_service.listar_empleados_ciclo()
        ws = wb.create_sheet(title=f"Consolidado_Ciclo_{mes:02d}_{anio}"[:31])

        headers = [
            "LEGAJO", "APELLIDOS Y NOMBRES", "DOTACIÓN", "ESPECIALIDAD", "TURNO BASE", "FRANCO CCT",
            "DÍA", "FECHA", "DÍA SEM.", "ESTADO", "HORA DISPONIBLE",
            "TOMA", "DEJA", "HS TRABAJADAS", "HS SEMANA", "REEMPLAZA A / MOTIVO", "OBSERVACIONES",
        ]
        ws.append(headers)

        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
        center_align = Alignment(horizontal="center", vertical="center")

        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center_align

        fill_franco = PatternFill(start_color="FFF9C4", end_color="FFF9C4", fill_type="solid")
        fill_relevo = PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid")

        keys_dia = [
            "dia", "fecha", "dia_semana", "estado", "hora_disponible",
            "toma", "deja", "horas_trabajadas", "hs_semana", "reemplaza_a", "observaciones",
        ]

        for emp in empleados:
            leg = int(emp["legajo"])
            dias_datos = records_service.obtener_calendario_mensual_ciclo(leg, anio, mes)
            emp_prefix = [
                emp.get("legajo", ""),
                emp.get("apellidos_nombres", ""),
                emp.get("dotacion", ""),
                emp.get("especialidad", ""),
                emp.get("turnos", ""),
                emp.get("franco", ""),
            ]
            for row_data in dias_datos:
                row_vals = emp_prefix + [row_data.get(k, "") for k in keys_dia]
                ws.append(row_vals)
                curr_row = ws.max_row
                estado = str(row_data.get("estado", ""))
                if "Franco" in estado:
                    for c in range(1, len(headers) + 1):
                        ws.cell(row=curr_row, column=c).fill = fill_franco
                elif "Reemplazando" in estado or "Relevo" in estado:
                    for c in range(1, len(headers) + 1):
                        ws.cell(row=curr_row, column=c).fill = fill_relevo

        ws.freeze_panes = "A2"
        if ws.max_row > 1:
            ws.auto_filter.ref = ws.dimensions

        for index, header in enumerate(headers, start=1):
            ws.column_dimensions[get_column_letter(index)].width = max(12, min(36, len(str(header)) + 4))

        wb.save(out_file)
        return str(out_file)

    def export_ordenamientos(
        self,
        ordenamientos: List[Dict[str, Any]],
        output_path: str,
    ) -> str:
        """Exporta el listado de Ordenamientos Operativos y sus tramos a un libro Excel."""
        out_file = Path(output_path)
        out_file.parent.mkdir(parents=True, exist_ok=True)

        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        _write_sheet(
            wb,
            "Ordenamientos",
            [
                "ID", "N° ORDEN", "FECHA EMISIÓN", "HORA EMISIÓN", "LEGAJO",
                "NOMBRE Y APELLIDO", "DOTACIÓN", "ITINERARIO", "TURNO / CICLO",
                "RÉGIMEN", "TRENES", "MOTIVO DE LA ORDEN", "REF ESTACIÓN",
                "EMISOR", "USUARIO WINDOWS", "OBSERVACIONES",
            ],
            ordenamientos,
            [
                "id", "nro_orden", "fecha_emision", "hora_emision", "legajo_personal",
                "nombre_personal", "deposito", "itinerario", "ciclo_o_turno",
                "regimen", "trenes_involucrados", "motivo_orden", "ref_pct",
                "emisor", "usuario_windows", "observacion_pie",
            ],
        )

        # Hoja de detalle de tramos si están disponibles en la BD
        with self.db.read_connection() as conn:
            ids = [int(o["id"]) for o in ordenamientos if o.get("id")]
            if ids:
                placeholders = ",".join("?" for _ in ids)
                tramos_rows = conn.execute(
                    f"""
                    SELECT o.nro_orden, o.fecha_emision, o.legajo_personal, o.nombre_personal,
                           o.deposito, ot.nro_secuencia, ot.ref, ot.nro_tren,
                           ot.origen, ot.hora_salida, ot.destino, ot.hora_llegada, ot.observaciones
                    FROM ordenamiento_tramos ot
                    JOIN ordenamientos o ON o.id = ot.ordenamiento_id
                    WHERE o.id IN ({placeholders})
                    ORDER BY o.id ASC, ot.nro_secuencia ASC
                    """,
                    ids,
                ).fetchall()
                _write_sheet(
                    wb,
                    "Detalle de Tramos",
                    [
                        "N° ORDEN", "FECHA", "LEGAJO", "AGENTE", "DOTACIÓN",
                        "SECUENCIA", "REF", "N° TREN", "ORIGEN", "HORA SALIDA",
                        "DESTINO", "HORA LLEGADA", "OBSERVACIONES",
                    ],
                    [dict(r) for r in tramos_rows],
                    [
                        "nro_orden", "fecha_emision", "legajo_personal", "nombre_personal", "deposito",
                        "nro_secuencia", "ref", "nro_tren", "origen", "hora_salida",
                        "destino", "hora_llegada", "observaciones",
                    ],
                )

        wb.save(out_file)
        return str(out_file)


    def export_auditoria(
        self,
        file_path: str | Path,
        logs: List[Dict[str, Any]],
        filtro_aplicado: str = "",
    ) -> str:
        """Exporta los logs de trazabilidad y auditoría a una planilla Excel (.xlsx)."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Auditoría del Sistema"
        ws.views.sheetView[0].showGridLines = True

        # Título
        ws.merge_cells("A1:H1")
        title_cell = ws["A1"]
        title_cell.value = "REGISTRO DE TRAZABILIDAD Y AUDITORÍA DEL SISTEMA - RENO"
        title_cell.font = Font(name="Segoe UI", size=14, bold=True, color="FFFFFF")
        title_cell.fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        title_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 36

        # Subtítulo con metadata
        now_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        sub_text = f"Generado el: {now_str} | Total registros: {len(logs)}"
        if filtro_aplicado:
            sub_text += f" | Filtro: '{filtro_aplicado}'"
        ws.merge_cells("A2:H2")
        sub_cell = ws["A2"]
        sub_cell.value = sub_text
        sub_cell.font = Font(name="Segoe UI", size=9, italic=True, color="475569")
        sub_cell.fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
        sub_cell.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[2].height = 20

        headers = [
            "ID", "FECHA Y HORA", "USUARIO", "ACCIÓN",
            "ENTIDAD", "ID REGISTRO", "DATOS ANTERIORES", "DATOS NUEVOS",
        ]
        keys = [
            "id", "creado_en", "usuario", "accion",
            "entidad", "entidad_id", "datos_anteriores", "datos_nuevos",
        ]

        # Encabezados
        for col_idx, h in enumerate(headers, 1):
            c = ws.cell(row=3, column=col_idx, value=h)
            c.font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
            c.fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = Border(
                top=Side(style="thin", color="1E3A8A"),
                bottom=Side(style="medium", color="1E3A8A"),
                left=Side(style="thin", color="CBD5E1"),
                right=Side(style="thin", color="CBD5E1"),
            )
        ws.row_dimensions[3].height = 24

        thin_border = Border(
            top=Side(style="thin", color="E2E8F0"),
            bottom=Side(style="thin", color="E2E8F0"),
            left=Side(style="thin", color="E2E8F0"),
            right=Side(style="thin", color="E2E8F0"),
        )
        zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

        for r_idx, log in enumerate(logs, start=4):
            fill = zebra_fill if r_idx % 2 == 0 else PatternFill(fill_type=None)
            for c_idx, k in enumerate(keys, start=1):
                val = log.get(k)
                if val is None:
                    val = ""
                cell = ws.cell(row=r_idx, column=c_idx, value=str(val))
                cell.font = Font(name="Segoe UI", size=9)
                cell.border = thin_border
                cell.fill = fill
                if k in ("id", "entidad_id"):
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif k in ("creado_en", "usuario", "accion", "entidad"):
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.row in (1, 2):
                    continue
                v_str = str(cell.value or "")
                if "\n" in v_str:
                    lines = v_str.split("\n")
                    max_len = max(max_len, max(len(l) for l in lines))
                else:
                    max_len = max(max_len, len(v_str))
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 10), 60)

        ws.column_dimensions["A"].width = 8
        ws.column_dimensions["B"].width = 19
        ws.column_dimensions["C"].width = 16
        ws.column_dimensions["D"].width = 14
        ws.column_dimensions["E"].width = 18
        ws.column_dimensions["F"].width = 12

        wb.save(file_path)
        return str(file_path)
