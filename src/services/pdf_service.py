"""Servicio de generación de documentos PDF con ReportLab.

Genera constancias formales de Novedades, Cambios de Turno,
Calendarios Mensuales de Personal de Ciclo e Informes Ejecutivos del Dashboard.
"""

from pathlib import Path
from typing import Any, Dict, List, Optional
from datetime import datetime

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch, cm
from reportlab.platypus import (
    HRFlowable,
    Image,
    KeepTogether,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

from src.core.logger import get_logger

logger = get_logger("PDFService")

# Paleta corporativa unificada con CRONO-PCT
PRIMARY = colors.HexColor("#1F497D")
SECONDARY = colors.HexColor("#4682B4")
ACCENT = colors.HexColor("#3CB371")
BG_LIGHT = colors.HexColor("#F4F6F9")
TEXT_DARK = colors.HexColor("#2C3E50")
BORDER_COLOR = colors.HexColor("#BDC3C7")


class NumberedCanvas:
    """Canvas de dos pasadas para numeración de páginas 'Página X de Y'."""
    pass


class PDFService:
    """Generador de documentos PDF para RENO v2.0.0."""

    @staticmethod
    def _crear_estilos() -> Dict[str, ParagraphStyle]:
        styles = getSampleStyleSheet()
        custom = {
            "Titulo": ParagraphStyle(
                "DocTitulo",
                parent=styles["Normal"],
                fontName="Helvetica-Bold",
                fontSize=16,
                leading=20,
                textColor=PRIMARY,
                alignment=1,  # Centrado
            ),
            "Subtitulo": ParagraphStyle(
                "DocSubtitulo",
                parent=styles["Normal"],
                fontName="Helvetica",
                fontSize=10,
                leading=14,
                textColor=TEXT_DARK,
                alignment=1,
            ),
            "Seccion": ParagraphStyle(
                "DocSeccion",
                parent=styles["Normal"],
                fontName="Helvetica-Bold",
                fontSize=12,
                leading=16,
                textColor=PRIMARY,
                spaceBefore=10,
                spaceAfter=4,
            ),
            "NormalCustom": ParagraphStyle(
                "DocNormal",
                parent=styles["Normal"],
                fontName="Helvetica",
                fontSize=9,
                leading=12,
                textColor=TEXT_DARK,
            ),
            "TablaHeader": ParagraphStyle(
                "DocTablaHeader",
                parent=styles["Normal"],
                fontName="Helvetica-Bold",
                fontSize=9,
                leading=11,
                textColor=colors.white,
                alignment=1,
            ),
            "TablaCell": ParagraphStyle(
                "DocTablaCell",
                parent=styles["Normal"],
                fontName="Helvetica",
                fontSize=8,
                leading=10,
                textColor=TEXT_DARK,
            ),
            "TablaCellCenter": ParagraphStyle(
                "DocTablaCellCenter",
                parent=styles["Normal"],
                fontName="Helvetica",
                fontSize=8,
                leading=10,
                textColor=TEXT_DARK,
                alignment=1,
            ),
        }
        return custom

    # ---------------------------------------------------------
    # Constancia de Novedad Operativa
    # ---------------------------------------------------------

    def generar_constancia_novedad(
        self,
        novedad: Dict[str, Any],
        tramos: List[Dict[str, Any]],
        output_path: str,
    ) -> str:
        out_file = Path(output_path)
        out_file.parent.mkdir(parents=True, exist_ok=True)
        doc = SimpleDocTemplate(
            str(out_file),
            pagesize=A4,
            leftMargin=1.5 * cm,
            rightMargin=1.5 * cm,
            topMargin=1.5 * cm,
            bottomMargin=1.5 * cm,
        )

        estilos = self._crear_estilos()
        story = []

        # Cabecera
        story.append(Paragraph("TRENES ARGENTINOS OPERACIONES", estilos["Subtitulo"]))
        story.append(Paragraph("CONSTANCIA DE NOVEDAD OPERATIVA", estilos["Titulo"]))
        story.append(Paragraph(f"Registro N°: {novedad.get('id', '-')} | Fecha: {novedad.get('registrado_en', '-')}", estilos["Subtitulo"]))
        story.append(Spacer(1, 10))
        story.append(HRFlowable(width="100%", thickness=1.5, color=PRIMARY, spaceBefore=0, spaceAfter=15))

        # Datos del Titular
        story.append(Paragraph("1. DATOS DEL TITULAR Y NOVEDAD", estilos["Seccion"]))
        datos_titular = [
            [
                Paragraph("<b>Legajo:</b>", estilos["NormalCustom"]),
                Paragraph(str(novedad.get("legajo", "-")), estilos["NormalCustom"]),
                Paragraph("<b>Apellidos y Nombres:</b>", estilos["NormalCustom"]),
                Paragraph(str(novedad.get("apellidos_nombres", "-")), estilos["NormalCustom"]),
            ],
            [
                Paragraph("<b>Especialidad:</b>", estilos["NormalCustom"]),
                Paragraph(str(novedad.get("especialidad", "-")), estilos["NormalCustom"]),
                Paragraph("<b>Dotación:</b>", estilos["NormalCustom"]),
                Paragraph(str(novedad.get("dotacion", "-")), estilos["NormalCustom"]),
            ],
            [
                Paragraph("<b>Turno Habitual:</b>", estilos["NormalCustom"]),
                Paragraph(str(novedad.get("turnos", "-")), estilos["NormalCustom"]),
                Paragraph("<b>Franco:</b>", estilos["NormalCustom"]),
                Paragraph(str(novedad.get("franco", "-")), estilos["NormalCustom"]),
            ],
            [
                Paragraph("<b>Tipo de Novedad:</b>", estilos["NormalCustom"]),
                Paragraph(str(novedad.get("novedad", "-")), estilos["NormalCustom"]),
                Paragraph("<b>Período:</b>", estilos["NormalCustom"]),
                Paragraph(f"{novedad.get('fecha_inicio', '-')} al {novedad.get('fecha_fin') or 'Abierto'}", estilos["NormalCustom"]),
            ],
            [
                Paragraph("<b>Ref. Estación:</b>", estilos["NormalCustom"]),
                Paragraph(str(novedad.get("referencia_estacion", "-")), estilos["NormalCustom"]),
                Paragraph("<b>Supervisor:</b>", estilos["NormalCustom"]),
                Paragraph(str(novedad.get("supervisor", "-")), estilos["NormalCustom"]),
            ],
        ]
        t_titular = Table(datos_titular, colWidths=[3.2 * cm, 5.5 * cm, 3.8 * cm, 5.5 * cm])
        t_titular.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), BG_LIGHT),
            ("GRID", (0, 0), (-1, -1), 0.5, BORDER_COLOR),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        story.append(t_titular)
        story.append(Spacer(1, 15))

        # Tramos de Cobertura / Relevos
        story.append(Paragraph("2. COBERTURA ASIGNADA (PERSONAL DE CICLO)", estilos["Seccion"]))
        if tramos:
            encabezados = [
                Paragraph("Legajo", estilos["TablaHeader"]),
                Paragraph("Personal de Reemplazo", estilos["TablaHeader"]),
                Paragraph("Turno Asignado", estilos["TablaHeader"]),
                Paragraph("Desde", estilos["TablaHeader"]),
                Paragraph("Hasta", estilos["TablaHeader"]),
            ]
            filas_tramos = [encabezados]
            for t in tramos:
                filas_tramos.append([
                    Paragraph(str(t.get("legajo_reemplazo", "-")), estilos["TablaCellCenter"]),
                    Paragraph(str(t.get("apellidos_nombres_reemplazo", "-")), estilos["TablaCell"]),
                    Paragraph(str(t.get("turnos_reemplazo", "-")), estilos["TablaCellCenter"]),
                    Paragraph(str(t.get("fecha_inicio", "-")), estilos["TablaCellCenter"]),
                    Paragraph(str(t.get("fecha_fin") or "Abierto"), estilos["TablaCellCenter"]),
                ])
            t_reemplazos = Table(filas_tramos, colWidths=[2.2 * cm, 7.5 * cm, 3.5 * cm, 2.4 * cm, 2.4 * cm])
            t_reemplazos.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), PRIMARY),
                ("GRID", (0, 0), (-1, -1), 0.5, BORDER_COLOR),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]))
            story.append(t_reemplazos)
        else:
            story.append(Paragraph("<i>Sin personal de relevo asignado (Turno descubierto o no aplica).</i>", estilos["NormalCustom"]))

        story.append(Spacer(1, 15))

        # Observaciones
        story.append(Paragraph("3. OBSERVACIONES", estilos["Seccion"]))
        obs_text = str(novedad.get("observaciones") or "Sin observaciones registradas.").replace("\n", "<br/>")
        t_obs = Table([[Paragraph(obs_text, estilos["NormalCustom"])]], colWidths=[18 * cm])
        t_obs.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), BG_LIGHT),
            ("BOX", (0, 0), (-1, -1), 0.5, BORDER_COLOR),
            ("TOPPADDING", (0, 0), (-1, -1), 8),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ]))
        story.append(t_obs)
        story.append(Spacer(1, 40))

        # Firmas
        firmas = [
            [
                Paragraph("_______________________________<br/>Firma del Agente", estilos["TablaCellCenter"]),
                Paragraph("_______________________________<br/>Firma y Sello del Supervisor", estilos["TablaCellCenter"]),
            ]
        ]
        t_firmas = Table(firmas, colWidths=[9 * cm, 9 * cm])
        t_firmas.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
        story.append(KeepTogether(t_firmas))

        doc.build(story)
        return str(out_file)

    # ---------------------------------------------------------
    # Constancia de Cambio de Turno
    # ---------------------------------------------------------

    def generar_constancia_cambio(
        self,
        cambio: Dict[str, Any],
        output_path: str,
    ) -> str:
        out_file = Path(output_path)
        out_file.parent.mkdir(parents=True, exist_ok=True)
        doc = SimpleDocTemplate(
            str(out_file),
            pagesize=A4,
            leftMargin=1.5 * cm,
            rightMargin=1.5 * cm,
            topMargin=1.5 * cm,
            bottomMargin=1.5 * cm,
        )

        estilos = self._crear_estilos()
        story = []

        # Cabecera
        story.append(Paragraph("TRENES ARGENTINOS OPERACIONES", estilos["Subtitulo"]))
        story.append(Paragraph("SOLICITUD / CONSTANCIA DE CAMBIO DE TURNO", estilos["Titulo"]))
        story.append(Paragraph(f"Registro N°: {cambio.get('id', '-')} | Fecha de Registro: {cambio.get('registrado_en', '-')}", estilos["Subtitulo"]))
        story.append(Spacer(1, 10))
        story.append(HRFlowable(width="100%", thickness=1.5, color=PRIMARY, spaceBefore=0, spaceAfter=15))

        # Agente 1
        story.append(Paragraph("1. DATOS DEL PRIMER INTERVINIENTE (SOLICITANTE)", estilos["Seccion"]))
        datos_agente1 = [
            [
                Paragraph("<b>Legajo:</b>", estilos["NormalCustom"]),
                Paragraph(str(cambio.get("legajo_1", "-")), estilos["NormalCustom"]),
                Paragraph("<b>Apellidos y Nombres:</b>", estilos["NormalCustom"]),
                Paragraph(str(cambio.get("apellidos_nombres_1", "-")), estilos["NormalCustom"]),
            ],
            [
                Paragraph("<b>Especialidad:</b>", estilos["NormalCustom"]),
                Paragraph(str(cambio.get("especialidad_1", "-")), estilos["NormalCustom"]),
                Paragraph("<b>Dotación:</b>", estilos["NormalCustom"]),
                Paragraph(str(cambio.get("dotacion_1", "-")), estilos["NormalCustom"]),
            ],
            [
                Paragraph("<b>Turno Original:</b>", estilos["NormalCustom"]),
                Paragraph(str(cambio.get("turnos_1", "-")), estilos["NormalCustom"]),
                Paragraph("<b>Franco:</b>", estilos["NormalCustom"]),
                Paragraph(str(cambio.get("franco_1", "-")), estilos["NormalCustom"]),
            ],
        ]
        t1 = Table(datos_agente1, colWidths=[3.2 * cm, 5.5 * cm, 3.8 * cm, 5.5 * cm])
        t1.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), BG_LIGHT),
            ("GRID", (0, 0), (-1, -1), 0.5, BORDER_COLOR),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        story.append(t1)
        story.append(Spacer(1, 12))

        # Agente 2
        story.append(Paragraph("2. DATOS DEL SEGUNDO INTERVINIENTE (RECEPTOR)", estilos["Seccion"]))
        datos_agente2 = [
            [
                Paragraph("<b>Legajo:</b>", estilos["NormalCustom"]),
                Paragraph(str(cambio.get("legajo_2", "-")), estilos["NormalCustom"]),
                Paragraph("<b>Apellidos y Nombres:</b>", estilos["NormalCustom"]),
                Paragraph(str(cambio.get("apellidos_nombres_2", "-")), estilos["NormalCustom"]),
            ],
            [
                Paragraph("<b>Especialidad:</b>", estilos["NormalCustom"]),
                Paragraph(str(cambio.get("especialidad_2", "-")), estilos["NormalCustom"]),
                Paragraph("<b>Dotación:</b>", estilos["NormalCustom"]),
                Paragraph(str(cambio.get("dotacion_2", "-")), estilos["NormalCustom"]),
            ],
            [
                Paragraph("<b>Turno Original:</b>", estilos["NormalCustom"]),
                Paragraph(str(cambio.get("turnos_2", "-")), estilos["NormalCustom"]),
                Paragraph("<b>Franco:</b>", estilos["NormalCustom"]),
                Paragraph(str(cambio.get("franco_2", "-")), estilos["NormalCustom"]),
            ],
        ]
        t2 = Table(datos_agente2, colWidths=[3.2 * cm, 5.5 * cm, 3.8 * cm, 5.5 * cm])
        t2.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), BG_LIGHT),
            ("GRID", (0, 0), (-1, -1), 0.5, BORDER_COLOR),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        story.append(t2)
        story.append(Spacer(1, 12))

        # Fecha de Cambio y Autorizaciones
        story.append(Paragraph("3. DETALLE DE LA PERMUTA Y REFERENCIAS", estilos["Seccion"]))
        datos_permuta = [
            [
                Paragraph("<b>Fecha del Cambio:</b>", estilos["NormalCustom"]),
                Paragraph(f"<b>{cambio.get('fecha_cambio', '-')}</b>", estilos["NormalCustom"]),
                Paragraph("<b>Ref. Estación:</b>", estilos["NormalCustom"]),
                Paragraph(str(cambio.get("referencia_estacion", "-")), estilos["NormalCustom"]),
            ],
            [
                Paragraph("<b>Supervisor:</b>", estilos["NormalCustom"]),
                Paragraph(str(cambio.get("supervisor", "-")), estilos["NormalCustom"]),
                Paragraph("<b>Observaciones:</b>", estilos["NormalCustom"]),
                Paragraph(str(cambio.get("observaciones") or "Sin observaciones"), estilos["NormalCustom"]),
            ],
        ]
        t_permuta = Table(datos_permuta, colWidths=[3.2 * cm, 5.5 * cm, 3.8 * cm, 5.5 * cm])
        t_permuta.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), BG_LIGHT),
            ("GRID", (0, 0), (-1, -1), 0.5, BORDER_COLOR),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        story.append(t_permuta)
        story.append(Spacer(1, 40))

        # Firmas Tripartitas
        firmas = [
            [
                Paragraph("________________________<br/>Firma Agente 1", estilos["TablaCellCenter"]),
                Paragraph("________________________<br/>Firma Agente 2", estilos["TablaCellCenter"]),
                Paragraph("________________________<br/>Firma Supervisor", estilos["TablaCellCenter"]),
            ]
        ]
        t_firmas = Table(firmas, colWidths=[6 * cm, 6 * cm, 6 * cm])
        t_firmas.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
        story.append(KeepTogether(t_firmas))

        doc.build(story)
        return str(out_file)

    # ---------------------------------------------------------
    # Calendario Mensual de Personal de Ciclo
    # ---------------------------------------------------------

    def generar_calendario_ciclo_pdf(
        self,
        empleado: Dict[str, Any],
        dias_mes: List[Dict[str, Any]],
        mes: int,
        anio: int,
        output_path: str,
    ) -> str:
        out_file = Path(output_path)
        out_file.parent.mkdir(parents=True, exist_ok=True)
        doc = SimpleDocTemplate(
            str(out_file),
            pagesize=A4,
            leftMargin=1.0 * cm,
            rightMargin=1.0 * cm,
            topMargin=1.0 * cm,
            bottomMargin=1.0 * cm,
        )

        estilos = self._crear_estilos()
        story = []

        nombre_meses = [
            "", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
            "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
        ]
        mes_txt = f"{nombre_meses[mes]} {anio}"

        # Cabecera
        story.append(Paragraph(f"CRONOGRAMA MENSUAL - PERSONAL DE CICLO: {mes_txt.upper()}", estilos["Titulo"]))
        sub = f"Agente: <b>{empleado.get('apellidos_nombres')}</b> (Leg. {empleado.get('legajo')}) | Dotación: {empleado.get('dotacion')} | Esp: {empleado.get('especialidad')} | Jornada CCT: {empleado.get('jornada_horas', 6)} hs"
        story.append(Paragraph(sub, estilos["Subtitulo"]))
        story.append(Spacer(1, 8))

        # Tabla mensual
        headers = [
            Paragraph("Día", estilos["TablaHeader"]),
            Paragraph("Fecha", estilos["TablaHeader"]),
            Paragraph("Día Sem", estilos["TablaHeader"]),
            Paragraph("Estado", estilos["TablaHeader"]),
            Paragraph("Disp.", estilos["TablaHeader"]),
            Paragraph("Toma", estilos["TablaHeader"]),
            Paragraph("Deja", estilos["TablaHeader"]),
            Paragraph("Hs Trab", estilos["TablaHeader"]),
            Paragraph("Hs Sem", estilos["TablaHeader"]),
            Paragraph("Reemplaza a / Motivo", estilos["TablaHeader"]),
        ]
        filas = [headers]

        for d in dias_mes:
            estado = d.get("estado", "")
            est_style = estilos["TablaCellCenter"]
            filas.append([
                Paragraph(str(d.get("dia")), estilos["TablaCellCenter"]),
                Paragraph(str(d.get("fecha")), estilos["TablaCellCenter"]),
                Paragraph(str(d.get("dia_semana", ""))[:3], estilos["TablaCellCenter"]),
                Paragraph(estado, est_style),
                Paragraph(str(d.get("hora_disponible", "-")), estilos["TablaCellCenter"]),
                Paragraph(str(d.get("toma", "-")), estilos["TablaCellCenter"]),
                Paragraph(str(d.get("deja", "-")), estilos["TablaCellCenter"]),
                Paragraph(str(d.get("horas_trabajadas", "-")), estilos["TablaCellCenter"]),
                Paragraph(str(d.get("hs_semana", "-")), estilos["TablaCellCenter"]),
                Paragraph(str(d.get("reemplaza_a") if d.get("reemplaza_a") != "-" else d.get("observaciones", "-")), estilos["TablaCell"]),
            ])

        col_w = [1.0 * cm, 1.8 * cm, 1.4 * cm, 2.8 * cm, 1.4 * cm, 1.4 * cm, 1.4 * cm, 1.5 * cm, 1.5 * cm, 4.8 * cm]
        t = Table(filas, colWidths=col_w, repeatRows=1)
        t_style = [
            ("BACKGROUND", (0, 0), (-1, 0), PRIMARY),
            ("GRID", (0, 0), (-1, -1), 0.5, BORDER_COLOR),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 2.5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
        ]

        # Resaltado de fines de semana o francos
        for idx, d in enumerate(dias_mes, start=1):
            if "Franco" in str(d.get("estado")):
                t_style.append(("BACKGROUND", (0, idx), (-1, idx), colors.HexColor("#FFF2CC")))
            elif str(d.get("estado")) == "Reemplazando":
                t_style.append(("BACKGROUND", (0, idx), (-1, idx), colors.HexColor("#E2EFDA")))

        t.setStyle(TableStyle(t_style))
        story.append(t)

        doc.build(story)
        return str(out_file)

    # ---------------------------------------------------------
    # Informe Ejecutivo del Dashboard (Requerimiento Usuario)
    # ---------------------------------------------------------

    def generar_informe_dashboard_pdf(
        self,
        metricas: Dict[str, Any],
        output_path: str,
        periodo_dias: int = 30,
        chart_image_paths: Optional[List[str]] = None,
    ) -> str:
        """Genera un informe ejecutivo del Panel de Control con KPIs, tablas detalladas y gráficos."""
        out_file = Path(output_path)
        out_file.parent.mkdir(parents=True, exist_ok=True)
        doc = SimpleDocTemplate(
            str(out_file),
            pagesize=A4,
            leftMargin=1.5 * cm,
            rightMargin=1.5 * cm,
            topMargin=1.5 * cm,
            bottomMargin=1.5 * cm,
        )

        estilos = self._crear_estilos()
        story = []

        # Título
        story.append(Paragraph("TRENES ARGENTINOS OPERACIONES", estilos["Subtitulo"]))
        story.append(Paragraph("INFORME EJECUTIVO - PANEL DE CONTROL (RENO v2.0)", estilos["Titulo"]))
        story.append(Paragraph(
            f"Período de Análisis: Últimos {periodo_dias} días | Generado el {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",
            estilos["Subtitulo"]
        ))
        story.append(Spacer(1, 10))
        story.append(HRFlowable(width="100%", thickness=1.5, color=PRIMARY, spaceBefore=0, spaceAfter=15))

        # 1. Resumen de KPIs
        story.append(Paragraph("1. RESUMEN DE INDICADORES CLAVE (KPIS)", estilos["Seccion"]))
        resumen = metricas.get("resumen", {})
        kpis = [
            [
                Paragraph("<b>Novedades Hoy:</b>", estilos["NormalCustom"]),
                Paragraph(str(resumen.get("novedades_hoy", 0)), estilos["NormalCustom"]),
                Paragraph("<b>Cambios Hoy:</b>", estilos["NormalCustom"]),
                Paragraph(str(resumen.get("cambios_hoy", 0)), estilos["NormalCustom"]),
            ],
            [
                Paragraph("<b>Total Novedades Mes:</b>", estilos["NormalCustom"]),
                Paragraph(str(resumen.get("novedades_mes", 0)), estilos["NormalCustom"]),
                Paragraph("<b>Total Cambios Mes:</b>", estilos["NormalCustom"]),
                Paragraph(str(resumen.get("cambios_mes", 0)), estilos["NormalCustom"]),
            ],
            [
                Paragraph("<b>Volumen Total Mensual:</b>", estilos["NormalCustom"]),
                Paragraph(f"<b>{resumen.get('total_mes', 0)} registros</b>", estilos["NormalCustom"]),
                Paragraph("<b>Estado del Sistema:</b>", estilos["NormalCustom"]),
                Paragraph("Operativo / Sincronizado", estilos["NormalCustom"]),
            ],
        ]
        t_kpi = Table(kpis, colWidths=[4.5 * cm, 4.5 * cm, 4.5 * cm, 4.5 * cm])
        t_kpi.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), BG_LIGHT),
            ("GRID", (0, 0), (-1, -1), 0.5, BORDER_COLOR),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ]))
        story.append(t_kpi)
        story.append(Spacer(1, 15))

        # Si se suministraron capturas de gráficos PNG desde la UI, incrustarlas
        if chart_image_paths:
            story.append(Paragraph("GRÁFICOS OPERATIVOS DEL SISTEMA", estilos["Seccion"]))
            for img_p in chart_image_paths:
                if Path(img_p).exists():
                    story.append(Image(str(img_p), width=17 * cm, height=6.5 * cm))
                    story.append(Spacer(1, 10))

        # 2. Desglose de Novedades por Tipo
        story.append(Paragraph("2. NOVEDADES REGISTRADAS POR TIPO (TOP)", estilos["Seccion"]))
        por_tipo = metricas.get("por_tipo", [])
        total_tipo = sum(v for _, v in por_tipo) or 1
        filas_tipo = [
            [Paragraph("Tipo de Novedad", estilos["TablaHeader"]), Paragraph("Cantidad", estilos["TablaHeader"]), Paragraph("Participación", estilos["TablaHeader"])]
        ]
        for t, c in por_tipo[:12]:
            filas_tipo.append([
                Paragraph(str(t), estilos["TablaCell"]),
                Paragraph(str(c), estilos["TablaCellCenter"]),
                Paragraph(f"{(c / total_tipo) * 100:.1f}%", estilos["TablaCellCenter"]),
            ])
        t_tipo = Table(filas_tipo, colWidths=[10 * cm, 4 * cm, 4 * cm])
        t_tipo.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), PRIMARY),
            ("GRID", (0, 0), (-1, -1), 0.5, BORDER_COLOR),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ]))
        story.append(t_tipo)
        story.append(Spacer(1, 15))

        # 3. Registros por Dotación
        story.append(Paragraph("3. DISTRIBUCIÓN POR DOTACIÓN", estilos["Seccion"]))
        por_dot = metricas.get("por_dotacion", [])
        total_dot = sum(v for _, v in por_dot) or 1
        filas_dot = [
            [Paragraph("Dotación", estilos["TablaHeader"]), Paragraph("Cantidad", estilos["TablaHeader"]), Paragraph("Participación", estilos["TablaHeader"])]
        ]
        for d, c in por_dot[:10]:
            filas_dot.append([
                Paragraph(str(d), estilos["TablaCell"]),
                Paragraph(str(c), estilos["TablaCellCenter"]),
                Paragraph(f"{(c / total_dot) * 100:.1f}%", estilos["TablaCellCenter"]),
            ])
        t_dot = Table(filas_dot, colWidths=[10 * cm, 4 * cm, 4 * cm])
        t_dot.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), SECONDARY),
            ("GRID", (0, 0), (-1, -1), 0.5, BORDER_COLOR),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ]))
        story.append(t_dot)
        story.append(Spacer(1, 15))

        # 4. Registros por Día de Semana
        story.append(Paragraph("4. DISTRIBUCIÓN POR DÍA DE LA SEMANA", estilos["Seccion"]))
        por_dia = metricas.get("por_dia_semana", [])
        filas_dia = [
            [Paragraph("Día", estilos["TablaHeader"]), Paragraph("Cantidad de Registros", estilos["TablaHeader"])]
        ]
        for dia, cant in por_dia:
            filas_dia.append([
                Paragraph(str(dia), estilos["TablaCell"]),
                Paragraph(str(cant), estilos["TablaCellCenter"]),
            ])
        t_dia = Table(filas_dia, colWidths=[9 * cm, 9 * cm])
        t_dia.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), PRIMARY),
            ("GRID", (0, 0), (-1, -1), 0.5, BORDER_COLOR),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ]))
        story.append(t_dia)

        doc.build(story)
        return str(out_file)
