"""Servicio de generación de formularios y documentos Word (.docx) con python-docx.

Genera solicitudes editables de cambio de turno, licencias y notas oficiales.
"""

from pathlib import Path
from typing import Any, Dict, List, Optional
from datetime import datetime

import docx
from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Inches, Pt, RGBColor

from src.core.logger import get_logger

logger = get_logger("WordService")

COLOR_PRIMARY = RGBColor(31, 73, 125)  # #1F497D
COLOR_SECONDARY = RGBColor(70, 130, 180)  # #4682B4
COLOR_MUTED = RGBColor(100, 100, 100)


def _set_cell_background(cell, color_hex: str):
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)


def _set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in (('top', top), ('bottom', bottom), ('left', left), ('right', right)):
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)


class WordService:
    """Generador de documentos .docx para RENO v2.0.0."""

    @staticmethod
    def _crear_documento_base() -> Document:
        doc = Document()
        for section in doc.sections:
            section.top_margin = Inches(0.8)
            section.bottom_margin = Inches(0.8)
            section.left_margin = Inches(0.8)
            section.right_margin = Inches(0.8)
        return doc

    def generar_solicitud_cambio_turno_docx(
        self,
        cambio: Dict[str, Any],
        output_path: str,
    ) -> str:
        out_file = Path(output_path)
        out_file.parent.mkdir(parents=True, exist_ok=True)
        doc = self._crear_documento_base()

        # Encabezado
        p_sub = doc.add_paragraph()
        p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_sub = p_sub.add_run("TRENES ARGENTINOS OPERACIONES\nLÍNEA GENERAL ROCA - GERENCIA DE OPERACIONES")
        r_sub.font.name = "Arial"
        r_sub.font.size = Pt(9)
        r_sub.font.color.rgb = COLOR_MUTED

        p_tit = doc.add_paragraph()
        p_tit.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_tit = p_tit.add_run("SOLICITUD FORMAL DE CAMBIO DE TURNO")
        r_tit.bold = True
        r_tit.font.name = "Arial"
        r_tit.font.size = Pt(14)
        r_tit.font.color.rgb = COLOR_PRIMARY

        doc.add_paragraph(f"Registro N°: {cambio.get('id', '-')} | Fecha: {cambio.get('registrado_en', '-')}")

        # Agente Solicitante
        p_sec1 = doc.add_paragraph()
        r_sec1 = p_sec1.add_run("1. AGENTE SOLICITANTE")
        r_sec1.bold = True
        r_sec1.font.color.rgb = COLOR_PRIMARY

        t1 = doc.add_table(rows=2, cols=4)
        t1.alignment = WD_TABLE_ALIGNMENT.CENTER
        datos1 = [
            ("Legajo:", str(cambio.get("legajo_1", "-")), "Apellidos y Nombres:", str(cambio.get("apellidos_nombres_1", "-"))),
            ("Especialidad:", str(cambio.get("especialidad_1", "-")), "Dotación:", str(cambio.get("dotacion_1", "-"))),
        ]
        for row_idx, fila in enumerate(datos1):
            for col_idx, texto in enumerate(fila):
                cell = t1.cell(row_idx, col_idx)
                p = cell.paragraphs[0]
                run = p.add_run(texto)
                if col_idx % 2 == 0:
                    run.bold = True
                _set_cell_background(cell, "F4F6F9" if row_idx % 2 == 0 else "FFFFFF")
                _set_cell_margins(cell, 80, 80, 100, 100)

        doc.add_paragraph()

        # Agente Receptor
        p_sec2 = doc.add_paragraph()
        r_sec2 = p_sec2.add_run("2. AGENTE RECEPTOR")
        r_sec2.bold = True
        r_sec2.font.color.rgb = COLOR_PRIMARY

        t2 = doc.add_table(rows=2, cols=4)
        t2.alignment = WD_TABLE_ALIGNMENT.CENTER
        datos2 = [
            ("Legajo:", str(cambio.get("legajo_2", "-")), "Apellidos y Nombres:", str(cambio.get("apellidos_nombres_2", "-"))),
            ("Especialidad:", str(cambio.get("especialidad_2", "-")), "Dotación:", str(cambio.get("dotacion_2", "-"))),
        ]
        for row_idx, fila in enumerate(datos2):
            for col_idx, texto in enumerate(fila):
                cell = t2.cell(row_idx, col_idx)
                p = cell.paragraphs[0]
                run = p.add_run(texto)
                if col_idx % 2 == 0:
                    run.bold = True
                _set_cell_background(cell, "F4F6F9" if row_idx % 2 == 0 else "FFFFFF")
                _set_cell_margins(cell, 80, 80, 100, 100)

        doc.add_paragraph()

        # Detalle de la Permuta
        p_sec3 = doc.add_paragraph()
        r_sec3 = p_sec3.add_run("3. DETALLE DE LA PERMUTA Y REFERENCIA")
        r_sec3.bold = True
        r_sec3.font.color.rgb = COLOR_PRIMARY

        p_det = doc.add_paragraph()
        p_det.add_run(f"Fecha efectiva del cambio de turno: ").bold = True
        p_det.add_run(f"{cambio.get('fecha_cambio', '-')}\n")
        p_det.add_run("Referencia de Estación: ").bold = True
        p_det.add_run(f"{cambio.get('referencia_estacion', '-')}\n")
        p_det.add_run("Supervisor interviniente: ").bold = True
        p_det.add_run(f"{cambio.get('supervisor', '-')}\n")
        p_det.add_run("Observaciones: ").bold = True
        p_det.add_run(f"{cambio.get('observaciones') or 'Sin observaciones'}")

        doc.add_paragraph()
        doc.add_paragraph()

        # Firmas
        t_firmas = doc.add_table(rows=1, cols=3)
        t_firmas.alignment = WD_TABLE_ALIGNMENT.CENTER
        firmas = ["Firma Solicitante", "Firma Receptor", "Firma y Sello Supervisor"]
        for idx, f in enumerate(firmas):
            cell = t_firmas.cell(0, idx)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run(f"\n\n_________________________\n{f}")

        doc.save(out_file)
        return str(out_file)

    def generar_solicitud_licencia_docx(
        self,
        novedad: Dict[str, Any],
        output_path: str,
    ) -> str:
        out_file = Path(output_path)
        out_file.parent.mkdir(parents=True, exist_ok=True)
        doc = self._crear_documento_base()

        # Encabezado
        p_sub = doc.add_paragraph()
        p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_sub = p_sub.add_run("TRENES ARGENTINOS OPERACIONES\nRECURSOS HUMANOS - CONTROL OPERATIVO")
        r_sub.font.name = "Arial"
        r_sub.font.size = Pt(9)
        r_sub.font.color.rgb = COLOR_MUTED

        p_tit = doc.add_paragraph()
        p_tit.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_tit = p_tit.add_run("FORMULARIO DE NOTIFICACIÓN DE LICENCIA / NOVEDAD")
        r_tit.bold = True
        r_tit.font.name = "Arial"
        r_tit.font.size = Pt(14)
        r_tit.font.color.rgb = COLOR_PRIMARY

        doc.add_paragraph(f"Novedad N°: {novedad.get('id', '-')} | Fecha: {novedad.get('registrado_en', '-')}")

        p_body = doc.add_paragraph()
        p_body.add_run(f"Por medio de la presente se deja constancia de la novedad operativa correspondiente al agente ")
        p_body.add_run(f"{novedad.get('apellidos_nombres')}").bold = True
        p_body.add_run(f" (Legajo N° {novedad.get('legajo')}), perteneciente a la dotación ")
        p_body.add_run(f"{novedad.get('dotacion')}").bold = True
        p_body.add_run(f" y especialidad ")
        p_body.add_run(f"{novedad.get('especialidad')}").bold = True
        p_body.add_run(".\n")

        # Tabla de Detalles
        t = doc.add_table(rows=4, cols=2)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        filas = [
            ("Motivo / Tipo de Novedad:", str(novedad.get("novedad", "-"))),
            ("Fecha de Inicio:", str(novedad.get("fecha_inicio", "-"))),
            ("Fecha de Fin:", str(novedad.get("fecha_fin") or "Abierto")),
            ("Observaciones:", str(novedad.get("observaciones") or "Sin observaciones")),
        ]
        for row_idx, (campo, val) in enumerate(filas):
            c1, c2 = t.cell(row_idx, 0), t.cell(row_idx, 1)
            c1.paragraphs[0].add_run(campo).bold = True
            c2.paragraphs[0].add_run(val)
            _set_cell_background(c1, "F4F6F9")
            _set_cell_margins(c1, 80, 80, 100, 100)
            _set_cell_margins(c2, 80, 80, 100, 100)

        doc.add_paragraph()
        doc.add_paragraph()

        # Firmas
        t_firmas = doc.add_table(rows=1, cols=2)
        t_firmas.alignment = WD_TABLE_ALIGNMENT.CENTER
        c1, c2 = t_firmas.cell(0, 0), t_firmas.cell(0, 1)
        c1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        c1.paragraphs[0].add_run("\n\n_________________________\nFirma del Agente")
        c2.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        c2.paragraphs[0].add_run("\n\n_________________________\nFirma y Sello del Supervisor")

        doc.save(out_file)
        return str(out_file)
