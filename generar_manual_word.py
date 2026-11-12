#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script para generar el Manual de Usuario y Administración de RENO
(Registro de Novedades y Cambios de Turno) en formato Microsoft Word (.docx).

Incluye explicaciones paso a paso de todos los módulos, tablas estilizadas,
recuadros de advertencias/consejos y espacios acotados para insertar capturas
de pantalla de la aplicación.
"""

import os
import sys
import argparse
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# --- Constantes de Diseño y Paleta de Colores ---
COLOR_PRIMARY = RGBColor(0x1E, 0x3A, 0x8A)      # Azul Marino Institucional (#1E3A8A)
COLOR_SECONDARY = RGBColor(0x25, 0x63, 0xEB)    # Azul Real (#2563EB)
COLOR_TEXT = RGBColor(0x33, 0x41, 0x55)         # Gris Oscuro (#334155)
COLOR_MUTED = RGBColor(0x64, 0x74, 0x8B)        # Gris Pizarra (#64748B)

HEX_PRIMARY = "1E3A8A"
HEX_SECONDARY = "2563EB"
HEX_BG_LIGHT = "F8FAFC"
HEX_BG_ALT_ROW = "F1F5F9"
HEX_BORDER_LIGHT = "CBD5E1"
HEX_BORDER_DASHED = "3B82F6"
HEX_CALLOUT_INFO = "3B82F6"
HEX_CALLOUT_WARN = "EAB308"
HEX_CALLOUT_TIP = "10B981"


class ManualGenerator:
    def __init__(self, output_path="Manual/Manual_de_Usuario_RENO_Paso_a_Paso.docx"):
        self.output_path = output_path
        self.doc = docx.Document()
        self.figure_counter = 1
        self._setup_page_layout()
        self._setup_styles()

    def _setup_page_layout(self):
        """Configura márgenes y tamaño de página estándar A4."""
        for section in self.doc.sections:
            section.page_width = Inches(8.27)    # A4 ancho (21 cm)
            section.page_height = Inches(11.69)  # A4 alto (29.7 cm)
            section.top_margin = Inches(1.0)
            section.bottom_margin = Inches(1.0)
            section.left_margin = Inches(1.0)
            section.right_margin = Inches(1.0)

    def _setup_styles(self):
        """Configura las fuentes y estilos por defecto del documento."""
        normal_style = self.doc.styles['Normal']
        normal_style.font.name = 'Calibri'
        normal_style.font.size = Pt(11)
        normal_style.font.color.rgb = COLOR_TEXT
        normal_style.paragraph_format.line_spacing = 1.15
        normal_style.paragraph_format.space_after = Pt(4)

    # --- Helpers de Contenido y Estilo ---

    def add_title_cover(self, title, subtitle, app_name, version, author, date_str):
        """Genera una portada elegante para el manual."""
        p_pre = self.doc.add_paragraph()
        p_pre.paragraph_format.space_before = Pt(40)
        p_pre.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_pre = p_pre.add_run(f"SISTEMA {app_name.upper()}")
        r_pre.font.size = Pt(14)
        r_pre.font.bold = True
        r_pre.font.color.rgb = COLOR_SECONDARY

        p_title = self.doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_title.paragraph_format.space_before = Pt(12)
        p_title.paragraph_format.space_after = Pt(8)
        r_title = p_title.add_run(title)
        r_title.font.size = Pt(28)
        r_title.font.bold = True
        r_title.font.color.rgb = COLOR_PRIMARY

        p_sub = self.doc.add_paragraph()
        p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_sub.paragraph_format.space_after = Pt(40)
        r_sub = p_sub.add_run(subtitle)
        r_sub.font.size = Pt(14)
        r_sub.font.italic = True
        r_sub.font.color.rgb = COLOR_MUTED

        # Recuadro decorativo de portada
        tbl = self.doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = tbl.cell(0, 0)
        cell.width = Inches(5.8)
        self._style_cell_box(cell, bg_color="F1F5F9", border_color=HEX_PRIMARY, border_style="solid", border_size="16")
        
        p_box = cell.paragraphs[0]
        p_box.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_b1 = p_box.add_run("GUÍA DE OPERACIÓN PASO A PASO Y ADMINISTRACIÓN\n")
        r_b1.font.bold = True
        r_b1.font.size = Pt(11)
        r_b1.font.color.rgb = COLOR_PRIMARY
        r_b2 = p_box.add_run("Incluye espacios acotados para inserción de capturas de pantalla de la interfaz gráfica.")
        r_b2.font.italic = True
        r_b2.font.size = Pt(10)
        r_b2.font.color.rgb = COLOR_TEXT

        p_meta = self.doc.add_paragraph()
        p_meta.paragraph_format.space_before = Pt(120)
        p_meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        r_meta1 = p_meta.add_run(f"Versión de Documentación: {version}\n")
        r_meta1.font.size = Pt(10)
        r_meta1.font.bold = True
        r_meta1.font.color.rgb = COLOR_TEXT
        
        r_meta2 = p_meta.add_run(f"Desarrollado por: {author}\nFecha: {date_str}\nEntorno de Base de Datos Compartida SQLite en Red")
        r_meta2.font.size = Pt(9.5)
        r_meta2.font.color.rgb = COLOR_MUTED

        self.doc.add_page_break()

    def add_h1(self, text):
        """Agrega un Título de Nivel 1 estilizado."""
        p = self.doc.add_paragraph()
        p.paragraph_format.space_before = Pt(20)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.size = Pt(16)
        run.font.bold = True
        run.font.color.rgb = COLOR_PRIMARY
        return p

    def add_h2(self, text):
        """Agrega un Título de Nivel 2 estilizado."""
        p = self.doc.add_paragraph()
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.size = Pt(13)
        run.font.bold = True
        run.font.color.rgb = COLOR_SECONDARY
        return p

    def add_h3(self, text):
        """Agrega un Título de Nivel 3."""
        p = self.doc.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.size = Pt(11.5)
        run.font.bold = True
        run.font.color.rgb = COLOR_TEXT
        return p

    def add_p(self, text, bold_prefix=None, italic=False, bullet=False):
        """Agrega un párrafo regular o con viñeta."""
        if bullet:
            p = self.doc.add_paragraph(style='List Bullet')
            p.paragraph_format.space_after = Pt(3)
        else:
            p = self.doc.add_paragraph()
            p.paragraph_format.space_after = Pt(5)

        if bold_prefix:
            r_bold = p.add_run(bold_prefix + " ")
            r_bold.font.bold = True
            r_bold.font.color.rgb = COLOR_TEXT

        run = p.add_run(text)
        run.font.italic = italic
        run.font.color.rgb = COLOR_TEXT
        return p

    def add_step(self, step_number, title, description):
        """Agrega un paso numerado con formato destacado."""
        p = self.doc.add_paragraph()
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.left_indent = Inches(0.2)

        r_num = p.add_run(f"Paso {step_number}: ")
        r_num.font.bold = True
        r_num.font.color.rgb = COLOR_SECONDARY
        
        r_title = p.add_run(f"{title}\n")
        r_title.font.bold = True
        r_title.font.color.rgb = COLOR_TEXT

        r_desc = p.add_run(description)
        r_desc.font.color.rgb = COLOR_TEXT
        return p

    def add_callout(self, text, title="NOTA IMPORTANTE", kind="info"):
        """Agrega un recuadro de aviso, advertencia o consejo."""
        tbl = self.doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.autofit = False
        cell = tbl.cell(0, 0)
        cell.width = Inches(6.27)

        color_map = {
            "info": (HEX_CALLOUT_INFO, "Información:", RGBColor(0x1D, 0x4E, 0xD8), "EFF6FF"),
            "warn": (HEX_CALLOUT_WARN, "Advertencia:", RGBColor(0xB4, 0x53, 0x09), "FFFBEB"),
            "tip": (HEX_CALLOUT_TIP, "Recomendación:", RGBColor(0x04, 0x78, 0x57), "ECFDF5"),
        }
        b_color, default_title, t_rgb, bg_color = color_map.get(kind, color_map["info"])
        box_title = title or default_title

        # Establecer estilo del recuadro (borde izquierdo grueso y fondo suave)
        self._style_callout_box(cell, border_color=b_color, bg_color=bg_color)

        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(2)
        r_title = p.add_run(f"📌 {box_title} ")
        r_title.font.bold = True
        r_title.font.color.rgb = t_rgb
        r_title.font.size = Pt(10.5)

        r_text = p.add_run(text)
        r_text.font.size = Pt(10)
        r_text.font.color.rgb = COLOR_TEXT

        # Espacio posterior para separar de la siguiente sección
        p_post = self.doc.add_paragraph()
        p_post.paragraph_format.space_after = Pt(4)

    def add_image_placeholder(self, screen_name, instructions, figure_desc, suggested_height_pt=90):
        """
        Crea un recuadro diseñado especialmente con borde punteado/discontinuo
        para que el usuario pegue o inserte una captura de pantalla de la aplicación.
        """
        tbl = self.doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.autofit = False
        cell = tbl.cell(0, 0)
        cell.width = Inches(6.27)

        # Aplicar borde discontinuo azul y fondo claro
        self._style_cell_box(cell, bg_color=HEX_BG_LIGHT, border_color=HEX_BORDER_DASHED, border_style="dashed", border_size="12")

        p_header = cell.paragraphs[0]
        p_header.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_header.paragraph_format.space_before = Pt(8)
        p_header.paragraph_format.space_after = Pt(2)
        
        r_icon = p_header.add_run("📷 [ESPACIO PARA CAPTURA DE PANTALLA]")
        r_icon.font.bold = True
        r_icon.font.size = Pt(10.5)
        r_icon.font.color.rgb = COLOR_SECONDARY

        p_screen = cell.add_paragraph()
        p_screen.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_screen.paragraph_format.space_after = Pt(2)
        r_screen = p_screen.add_run(f"Pantalla / Ventana: {screen_name}")
        r_screen.font.bold = True
        r_screen.font.size = Pt(10)
        r_screen.font.color.rgb = COLOR_PRIMARY

        p_inst = cell.add_paragraph()
        p_inst.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_inst.paragraph_format.space_after = Pt(4)
        r_inst = p_inst.add_run(f"👉 Elementos a capturar: {instructions}\n(Haga clic dentro de este recuadro y presione Ctrl + V para pegar la imagen)")
        r_inst.font.italic = True
        r_inst.font.size = Pt(9)
        r_inst.font.color.rgb = COLOR_MUTED

        # Párrafo de separación vertical que simula la altura del espacio para pegar
        p_space = cell.add_paragraph()
        p_space.paragraph_format.space_before = Pt(suggested_height_pt // 2)
        p_space.paragraph_format.space_after = Pt(suggested_height_pt // 2)

        # Epígrafe / Caption debajo de la figura
        p_caption = self.doc.add_paragraph()
        p_caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_caption.paragraph_format.space_before = Pt(4)
        p_caption.paragraph_format.space_after = Pt(10)
        
        r_cap_num = p_caption.add_run(f"Figura {self.figure_counter}: ")
        r_cap_num.font.bold = True
        r_cap_num.font.size = Pt(9.5)
        r_cap_num.font.color.rgb = COLOR_TEXT

        r_cap_text = p_caption.add_run(figure_desc)
        r_cap_text.font.italic = True
        r_cap_text.font.size = Pt(9.5)
        r_cap_text.font.color.rgb = COLOR_MUTED

        self.figure_counter += 1

    def add_table(self, headers, rows_data, col_widths=None):
        """Agrega una tabla formateada profesionalmente con cabecera sólida."""
        tbl = self.doc.add_table(rows=len(rows_data) + 1, cols=len(headers))
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.autofit = False

        # Cabecera
        hdr_row = tbl.rows[0]
        for idx, header_text in enumerate(headers):
            cell = hdr_row.cells[idx]
            if col_widths and idx < len(col_widths):
                cell.width = Inches(col_widths[idx])
            self._style_header_cell(cell, header_text)

        # Filas de datos con alternancia de fondo
        for r_idx, row_values in enumerate(rows_data):
            row = tbl.rows[r_idx + 1]
            bg_color = HEX_BG_ALT_ROW if (r_idx % 2 == 1) else "FFFFFF"
            for c_idx, val in enumerate(row_values):
                cell = row.cells[c_idx]
                if col_widths and c_idx < len(col_widths):
                    cell.width = Inches(col_widths[c_idx])
                self._style_data_cell(cell, str(val), bg_color)

        p_post = self.doc.add_paragraph()
        p_post.paragraph_format.space_after = Pt(6)

    # --- Métodos de Manipulación XML para Word ---

    def _style_header_cell(self, cell, text):
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{HEX_PRIMARY}"/>')
        borders = parse_xml(f'''
            <w:tcBorders {nsdecls("w")}>
                <w:top w:val="single" w:sz="6" w:space="0" w:color="{HEX_PRIMARY}"/>
                <w:bottom w:val="single" w:sz="12" w:space="0" w:color="{HEX_PRIMARY}"/>
                <w:left w:val="none"/>
                <w:right w:val="none"/>
            </w:tcBorders>
        ''')
        mar = parse_xml(f'''
            <w:tcMar {nsdecls("w")}>
                <w:top w:w="140" w:type="dxa"/>
                <w:bottom w:w="140" w:type="dxa"/>
                <w:left w:w="160" w:type="dxa"/>
                <w:right w:w="160" w:type="dxa"/>
            </w:tcMar>
        ''')
        tcPr = cell._tc.get_or_add_tcPr()
        tcPr.append(shd)
        tcPr.append(borders)
        tcPr.append(mar)

        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(text)
        run.font.bold = True
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    def _style_data_cell(self, cell, text, bg_hex):
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{bg_hex}"/>')
        borders = parse_xml(f'''
            <w:tcBorders {nsdecls("w")}>
                <w:top w:val="single" w:sz="4" w:space="0" w:color="{HEX_BORDER_LIGHT}"/>
                <w:bottom w:val="single" w:sz="4" w:space="0" w:color="{HEX_BORDER_LIGHT}"/>
                <w:left w:val="none"/>
                <w:right w:val="none"/>
            </w:tcBorders>
        ''')
        mar = parse_xml(f'''
            <w:tcMar {nsdecls("w")}>
                <w:top w:w="120" w:type="dxa"/>
                <w:bottom w:w="120" w:type="dxa"/>
                <w:left w:w="160" w:type="dxa"/>
                <w:right w:w="160" w:type="dxa"/>
            </w:tcMar>
        ''')
        tcPr = cell._tc.get_or_add_tcPr()
        tcPr.append(shd)
        tcPr.append(borders)
        tcPr.append(mar)

        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(text)
        run.font.size = Pt(9)
        run.font.color.rgb = COLOR_TEXT

    def _style_cell_box(self, cell, bg_color, border_color, border_style="dashed", border_size="12"):
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{bg_color}"/>')
        borders = parse_xml(f'''
            <w:tcBorders {nsdecls("w")}>
                <w:top w:val="{border_style}" w:sz="{border_size}" w:space="0" w:color="{border_color}"/>
                <w:left w:val="{border_style}" w:sz="{border_size}" w:space="0" w:color="{border_color}"/>
                <w:bottom w:val="{border_style}" w:sz="{border_size}" w:space="0" w:color="{border_color}"/>
                <w:right w:val="{border_style}" w:sz="{border_size}" w:space="0" w:color="{border_color}"/>
            </w:tcBorders>
        ''')
        mar = parse_xml(f'''
            <w:tcMar {nsdecls("w")}>
                <w:top w:w="180" w:type="dxa"/>
                <w:bottom w:w="180" w:type="dxa"/>
                <w:left w:w="220" w:type="dxa"/>
                <w:right w:w="220" w:type="dxa"/>
            </w:tcMar>
        ''')
        tcPr = cell._tc.get_or_add_tcPr()
        tcPr.append(shd)
        tcPr.append(borders)
        tcPr.append(mar)

    def _style_callout_box(self, cell, border_color, bg_color):
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{bg_color}"/>')
        borders = parse_xml(f'''
            <w:tcBorders {nsdecls("w")}>
                <w:top w:val="none"/>
                <w:bottom w:val="none"/>
                <w:right w:val="none"/>
                <w:left w:val="single" w:sz="24" w:space="0" w:color="{border_color}"/>
            </w:tcBorders>
        ''')
        mar = parse_xml(f'''
            <w:tcMar {nsdecls("w")}>
                <w:top w:w="140" w:type="dxa"/>
                <w:bottom w:w="140" w:type="dxa"/>
                <w:left w:w="200" w:type="dxa"/>
                <w:right w:w="200" w:type="dxa"/>
            </w:tcMar>
        ''')
        tcPr = cell._tc.get_or_add_tcPr()
        tcPr.append(shd)
        tcPr.append(borders)
        tcPr.append(mar)

    # --- Generación de Contenido por Módulo ---

    def build_document(self):
        """Construye todos los capítulos del manual con pasos e imágenes."""
        print("Generando portada e introducción...")
        self.add_title_cover(
            title="Manual de Usuario y Operación",
            subtitle="Guía Paso a Paso para la Carga, Cobertura, Gestión y Administración",
            app_name="RENO (Registro de Novedades y Turnos)",
            version="2.0 (Edición Ampliada)",
            author="Luciano Vidal",
            date_str="Noviembre 2026"
        )

        self._build_table_of_contents()
        self._build_section_1_intro()
        self._build_section_2_installation()
        self._build_section_3_auth()
        self._build_section_4_ui_overview()
        self._build_section_5_dashboard()
        self._build_section_6_novedades()
        self._build_section_7_reemplazos()
        self._build_section_8_cambios_turno()
        self._build_section_9_tablas_filtros()
        self._build_section_10_detalle_edicion()
        self._build_section_11_personal_ciclo()
        self._build_section_12_export_import()
        self._build_section_13_notificaciones()
        self._build_section_14_administracion()
        self._build_section_15_faq()
        self._build_section_16_anexo_permisos()

        os.makedirs(os.path.dirname(os.path.abspath(self.output_path)), exist_ok=True)
        self.doc.save(self.output_path)
        print(f"¡Manual generado exitosamente en: {self.output_path}!")

    def _build_table_of_contents(self):
        self.add_h1("Índice de Contenidos")
        toc_items = [
            ("1. Introducción y Arquitectura del Sistema", "Propósito, persistencia compartida en red y sincronización"),
            ("2. Instalación y Primer Uso", "Archivos requeridos, configuración path_base y creación del Administrador"),
            ("3. Acceso, Seguridad y Control de Sesión", "Inicio de sesión, recuperación de contraseña por email y caducidad"),
            ("4. Interfaz General y Navegación", "Ventana principal maximizada, barra de menús y temas visuales"),
            ("5. Panel de Control (Dashboard)", "Tarjetas de estado, métricas en tiempo real y gráficos estadísticos"),
            ("6. Registro de Novedades (Paso a Paso)", "Carga de novedades, búsqueda de personal, autocompletado y envío"),
            ("7. Cobertura de Novedades y Reemplazos de Ciclo", "Asignación de tramos, validación de cruces y estados de cobertura"),
            ("8. Registro de Cambios de Turno (Paso a Paso)", "Intercambio de turnos entre empleados, fechas y supervisor"),
            ("9. Tablas Operativas y Filtrado en Vivo", "Listados cronológicos, filtros por nombre/dotación y refresco 60s"),
            ("10. Detalle, Modificación y Eliminación", "Modales de detalle, límites de 24h/72h y modo restringido"),
            ("11. Módulo de Personal de Ciclo y Calendario", "Personal disponible/asignado, descansos CCT (18h/42h) y órdenes diarias"),
            ("12. Reportes, Descubiertos y Exportación a Excel", "Exportación filtrada, reporte de turnos sin cubrir e importación"),
            ("13. Sistema de Notificaciones y Avisos Toasts", "Menú de registros nuevos con contador y avisos emergentes"),
            ("14. Módulo de Administración del Sistema", "Usuarios, roles, empleados, catálogos, backups y auditoría"),
            ("15. Preguntas Frecuentes y Solución de Problemas", "Base ocupada, problemas de red, restablecimiento de acceso"),
            ("16. Anexo: Matriz de Permisos por Rol", "Detalle de permisos por módulo y acción en el sistema"),
        ]
        for title, desc in toc_items:
            self.add_p(f"{title} — {desc}", bullet=True)
        self.doc.add_page_break()

    def _build_section_1_intro(self):
        self.add_h1("1. Introducción y Arquitectura del Sistema")
        self.add_p(
            "La aplicación RENO (Registro de Novedades y Cambios de Turno) ha sido diseñada "
            "para reemplazar las tradicionales planillas compartidas de Excel por un sistema "
            "de escritorio robusto, ágil y seguro basado en una base de datos centralizada SQLite."
        )
        self.add_p(
            "Al ejecutarse sobre una base de datos estructurada con validaciones automáticas, "
            "se eliminan los riesgos de sobreescritura accidental, borrado de fórmulas y "
            "lentitud generada por macros de Excel en carpetas compartidas."
        )
        self.add_h2("1.1 Funcionamiento en Red Local Compartida")
        self.add_p(
            "RENO permite que múltiples estaciones de trabajo (computadoras en distintas dependencias o estaciones) "
            "operen simultáneamente contra una única base de datos ubicada en una carpeta compartida de red (SMB/Windows):",
            bold_prefix="Concurrencia transparente:"
        )
        self.add_p("Las lecturas no bloquean el archivo, permitiendo que varios usuarios consulten las tablas a la vez.", bullet=True)
        self.add_p("Las operaciones de guardado, modificación o borrado utilizan transacciones breves con bloqueo externo seguro.", bullet=True)
        self.add_p("Cada 60 segundos la aplicación refresca automáticamente los datos desde el archivo central sin interrumpir la escritura del usuario.", bullet=True)
        
        self.add_callout(
            "Para garantizar la máxima compatibilidad en unidades de red compartidas de Windows (SMB), "
            "el sistema utiliza el modo de diario tradicional (journal_mode=DELETE) y no WAL (Write-Ahead Logging), "
            "evitando bloqueos de archivos de memoria compartida entre distintas computadoras.",
            title="ARQUITECTURA DE DATOS EN RED",
            kind="info"
        )

    def _build_section_2_installation(self):
        self.add_h1("2. Instalación y Primer Uso")
        self.add_p(
            "RENO no requiere una instalación compleja con servicios en segundo plano. "
            "Se distribuye como un ejecutable portable independiente (RENO.exe) preparado para Windows."
        )
        self.add_h2("2.1 Archivos del Programa")
        self.add_p(
            "Para un correcto funcionamiento, en la misma carpeta del ejecutable pueden convivir los siguientes archivos auxiliares:"
        )
        self.add_table(
            headers=["Archivo / Carpeta", "Tipo", "Descripción y Función"],
            rows_data=[
                ["RENO.exe (o main.exe)", "Ejecutable", "Aplicación principal del sistema."],
                ["path_base", "Texto (sin extensión)", "Contiene la ruta absoluta o de red hacia el archivo de base de datos .sqlite (o el Excel original para la migración inicial)."],
                ["theme", "Texto (sin extensión)", "Guarda la preferencia del tema visual seleccionado (ej. flatly, cosmo, darkly) para ese equipo."],
                ["backups/", "Carpeta", "Directorio donde se generan los respaldos automáticos y manuales de la base de datos."],
            ],
            col_widths=[1.8, 1.2, 3.2]
        )

        self.add_h2("2.2 Primer Inicio: Creación del Administrador Inicial")
        self.add_p(
            "Cuando se ejecuta la aplicación por primera vez sobre una base de datos nueva o vacía, "
            "el sistema detecta automáticamente que no existen usuarios registrados y habilita el modo de inicialización."
        )
        self.add_step(1, "Iniciar la aplicación", "Haga doble clic sobre RENO.exe.")
        self.add_step(2, "Completar los datos del Administrador", "Aparecerá la pantalla de inicio de sesión con un campo extra llamado 'Legajo' y el botón 'Crear administrador inicial'.")
        self.add_step(3, "Ingresar credenciales maestras", "Defina el Nombre de Usuario, Contraseña y Legajo numérico del administrador principal.")
        self.add_step(4, "Confirmar creación", "Haga clic en 'Crear administrador inicial'. El sistema creará la cuenta con todos los permisos del rol Administrador y desaparecerá el botón de inicialización para los futuros arranques.")

        self.add_image_placeholder(
            screen_name="Inicio de Sesión - Primer Uso",
            instructions="Capturar la pantalla de login mostrando el campo Legajo y el botón 'Crear administrador inicial'.",
            figure_desc="Pantalla de inicio de sesión en modo de primer uso con creación de cuenta administrativa."
        )

    def _build_section_3_auth(self):
        self.add_h1("3. Control de Acceso, Seguridad y Sesión")
        self.add_h2("3.1 Inicio de Sesión Ordinario")
        self.add_p("Cada usuario debe autenticarse con su usuario y contraseña asignados:")
        self.add_step(1, "Usuario y Contraseña", "Ingrese su nombre de usuario en minúsculas y su clave personal.")
        self.add_step(2, "Ingresar al sistema", "Haga clic en 'Ingresar' o presione la tecla Enter.")
        self.add_p(
            "Por razones de seguridad, el sistema cuenta con un contador de intentos fallidos. "
            "Tras varios intentos incorrectos consecutivos, la cuenta se bloquea preventivamente requiriendo "
            "que un administrador la reactive o se use el restablecimiento de clave.",
            bold_prefix="Protección contra fuerza bruta:"
        )

        self.add_image_placeholder(
            screen_name="Pantalla de Login Habitual",
            instructions="Capturar la ventana de inicio de sesión estándar con campos de Usuario, Contraseña y enlace de recuperación.",
            figure_desc="Ventana estándar de inicio de sesión de RENO."
        )

        self.add_h2("3.2 Recuperación de Contraseña Olvidada")
        self.add_p(
            "Si un usuario olvida su contraseña, puede solicitar una contraseña temporal automática sin necesidad de llamar al administrador:"
        )
        self.add_step(1, "Solicitar recuperación", "En la pantalla de inicio de sesión, haga clic en el enlace '¿Olvidó su contraseña?'.")
        self.add_step(2, "Ingresar usuario", "Escriba su nombre de usuario y presione 'Enviar contraseña temporal'.")
        self.add_step(3, "Envío automático por correo", "El sistema autocompleta el correo institucional concatenando el usuario con el dominio configurado (ej: usuario@trenesargentinos.gob.ar) y envía un correo desatendido mediante Outlook.")
        
        self.add_callout(
            "El correo se transmite mediante una rutina de envío seguro (enviar_sin_rastro) que elimina "
            "cualquier copia del mensaje temporal tanto de la bandeja de 'Elementos enviados' como de 'Elementos eliminados', "
            "garantizando que ninguna otra persona con acceso al buzón pueda ver la clave generada.",
            title="CONFIDENCIALIDAD DEL ENVÍO",
            kind="tip"
        )

        self.add_image_placeholder(
            screen_name="Diálogo de Recuperación de Contraseña",
            instructions="Capturar la ventana modal que solicita el nombre de usuario para enviar la clave temporal por correo.",
            figure_desc="Modal de recuperación de contraseña con envío por correo institucional."
        )

        self.add_h2("3.3 Cambio Obligatorio de Contraseña")
        self.add_p(
            "La aplicación exige el cambio inmediato de contraseña en dos situaciones clave:"
        )
        self.add_p("Al iniciar sesión por primera vez con una contraseña temporal generada por el sistema.", bullet=True)
        self.add_p("Por expiración periódica: cuando han transcurrido los días configurados (por defecto 90 días) desde el último cambio.", bullet=True)
        self.add_p("El diálogo solicitará la nueva contraseña y su confirmación. Una vez validada, el acceso se reanuda con normalidad.")

        self.add_image_placeholder(
            screen_name="Diálogo de Cambio Obligatorio de Clave",
            instructions="Capturar el cuadro de diálogo solicitando 'Nueva contraseña' y 'Confirmar nueva contraseña'.",
            figure_desc="Ventana de cambio forzado de contraseña por caducidad o clave temporal."
        )

        self.add_h2("3.4 Tiempo de Inactividad y Cierre de Sesión")
        self.add_p(
            "Si la aplicación permanece sin interacción durante el período configurado (por defecto 30 minutos), "
            "la sesión se cierra automáticamente mostrando el aviso 'Sesión expirada por inactividad' para proteger los datos."
        )
        self.add_p(
            "Para cambiar manualmente de operador sin salir del programa, utilice la opción del menú Archivo > Cambiar usuario."
        )

    def _build_section_4_ui_overview(self):
        self.add_h1("4. Interfaz General y Navegación")
        self.add_p(
            "La interfaz gráfica de RENO está diseñada con ttkbootstrap, ofreciendo un entorno visual moderno, "
            "limpio y con alta legibilidad en pantallas táctiles o monitores estándar de oficina."
        )
        self.add_h2("4.1 Estructura de la Pantalla Principal")
        self.add_p("La pantalla principal se divide en cuatro áreas fundamentales:")
        self.add_p("Barra de Menús Superior: Da acceso a las funciones globales: Archivo, Registros nuevos, Opciones y Administración.", bullet=True)
        self.add_p("Barra de Acceso y Estado: Muestra el usuario activo, rol, base conectada y botones de cambio rápido de vista.", bullet=True)
        self.add_p("Área Central de Trabajo: Despliega las tablas de registros (Novedades o Cambios de Turno) o el Panel de Control.", bullet=True)
        self.add_p("Modales Optimizados: Las ventanas de carga, detalle, búsqueda y administración se abren maximizadas para garantizar la visualización completa de datos y tablas sin necesidad de desplazamientos molestos.", bullet=True)

        self.add_image_placeholder(
            screen_name="Ventana Principal del Sistema",
            instructions="Capturar la pantalla principal maximizada mostrando la barra de menús, botones de cambio de vista y la tabla de datos.",
            figure_desc="Estructura general de la ventana de trabajo de RENO."
        )

        self.add_h2("4.2 Personalización de Temas Visuales")
        self.add_p(
            "Cada usuario puede personalizar el estilo gráfico según su preferencia o iluminación del puesto de trabajo. "
            "Diríjase al menú Opciones > Seleccionar Tema y elija entre temas claros (flatly, cosmo, litera, minty) o temas oscuros (darkly, superhero, cyborg). "
            "La selección se guarda automáticamente en el archivo local 'theme'."
        )

        self.add_image_placeholder(
            screen_name="Selector de Temas Visuales",
            instructions="Capturar el menú desplegable o ventana de selección de temas de ttkbootstrap.",
            figure_desc="Menú de personalización de temas visuales de la aplicación."
        )

    def _build_section_5_dashboard(self):
        self.add_h1("5. Panel de Control (Dashboard) y Métricas")
        self.add_p(
            "Al iniciar sesión, si el usuario cuenta con permisos de visualización, el sistema presenta un "
            "Panel de Control gráfico con estadísticas en tiempo real sobre la operación diaria y mensual."
        )
        self.add_h2("5.1 Tarjetas de Indicadores Clave")
        self.add_table(
            headers=["Tarjeta / Métrica", "Significado Operativo"],
            rows_data=[
                ["Novedades de Hoy", "Cantidad total de novedades registradas en la fecha actual."],
                ["Cambios de Turno de Hoy", "Número de permutas o cambios de turno activos para el día de la fecha."],
                ["Total del Mes", "Acumulado mensual de eventos registrados en la base de datos."],
                ["Registros sin Revisar", "Novedades y cambios cargados por otros puestos que aún no han sido abiertos."],
            ],
            col_widths=[2.2, 4.0]
        )

        self.add_h2("5.2 Gráficos Estadísticos Dinámicos")
        self.add_p("El dashboard presenta cuatro representaciones gráficas interactivas:")
        self.add_p("Distribución por Tipo de Novedad: Proporción de licencias, partes médicos, francos compensatorios, etc.", bullet=True)
        self.add_p("Distribución por Dotación: Concentración de eventos por base o cabecera operativa.", bullet=True)
        self.add_p("Distribución por Día de la Semana: Días con mayor concentración de novedades para planificar refuerzos.", bullet=True)
        self.add_p("Tendencia Histórica: Gráfico de líneas con selector de horizonte temporal a 7, 30 o 90 días.", bullet=True)
        self.add_p("Desde cualquier tabla o módulo, presione el botón 'Panel de control' para regresar a esta pantalla resumen.")

        self.add_image_placeholder(
            screen_name="Panel de Control (Dashboard)",
            instructions="Capturar el dashboard completo con las 4 tarjetas superiores de métricas y los gráficos estadísticos.",
            figure_desc="Dashboard de indicadores y gráficos de análisis de novedades."
        )

    def _build_section_6_novedades(self):
        self.add_h1("6. Registro de Novedades (Paso a Paso)")
        self.add_p(
            "El formulario de novedades permite asentar cualquier eventualidad del personal (licencias médicas, "
            "vacaciones, suspensiones, informes de servicio, etc.), asignando reemplazos si corresponde."
        )
        self.add_h2("6.1 Pasos para Registrar una Novedad")
        
        self.add_step(1, "Abrir el formulario", "Desde la vista de Novedades, haga clic en el botón 'Cargar Novedad' (o presione el atajo configurado).")
        self.add_step(2, "Buscar el empleado", "Haga clic en 'Buscar Personal' o escriba el número de Legajo en el campo correspondiente. Se abrirá el buscador modal maximizado donde puede buscar por apellido o número.")
        self.add_step(3, "Carga automática de datos", "Al seleccionar el agente, el sistema completa automáticamente: Apellido y Nombre, Especialidad, Dotación, Turno habitual y Franco.")
        self.add_step(4, "Seleccionar Tipo de Novedad", "Despliegue el combo de 'Tipo Novedad'. Al elegir un tipo que tenga configurada una duración fija en días (ej. Licencia anual), el sistema calcula y completa automáticamente la Fecha de Fin.")
        self.add_step(5, "Establecer Fechas de Inicio y Fin", "Utilice el calendario desplegable (DateEntry) para seleccionar la fecha de inicio y verifique o complete la fecha de finalización.")
        self.add_step(6, "Referencia Estación inteligente", "Comience a escribir el nombre de la estación o del personal de referencia; el desplegable filtrará las opciones en tiempo real mostrando 'Nombre (DOTACIÓN)' sin perder el foco del teclado.")
        self.add_step(7, "Supervisor y Observaciones", "Indique el supervisor de turno y escriba las observaciones pertinentes en el cuadro de texto exclusivo.")
        self.add_step(8, "Guardar la novedad", "Presione 'Guardar Novedad'. El sistema validará los datos y grabará el registro en la base de datos compartida.")

        self.add_image_placeholder(
            screen_name="Buscador Modal de Personal",
            instructions="Capturar el buscador modal maximizado de empleados con la lista y la barra de búsqueda superior.",
            figure_desc="Modal de búsqueda y selección de empleados por legajo o nombre."
        )

        self.add_image_placeholder(
            screen_name="Formulario de Carga de Novedades",
            instructions="Capturar el formulario completo de novedades con datos de un empleado, fechas y referencia estación seleccionada.",
            figure_desc="Formulario principal para el registro de novedades del personal."
        )

        self.add_h2("6.2 Envío Automático de Correo para Novedades de Tipo 'Informe'")
        self.add_p(
            "Cuando se registra una novedad cuyo tipo es 'Informe', el sistema activa la integración con Microsoft Outlook: "
            "genera y despacha automáticamente un correo formal con el detalle de la novedad a todos los "
            "destinatarios configurados en la administración. Si Outlook no está disponible en la PC, el sistema "
            "guardará la novedad con éxito y mostrará un aviso advirtiendo que el correo no pudo enviarse."
        )

    def _build_section_7_reemplazos(self):
        self.add_h1("7. Cobertura de Novedades y Reemplazos de Ciclo")
        self.add_p(
            "Una de las fortalezas clave de RENO es el módulo de Reemplazos con Personal de Ciclo. "
            "Permite gestionar qué agentes con turno 'CICLO' cubrirán los turnos descubiertos por una novedad."
        )
        self.add_h2("7.1 Reglas del Negocio de Reemplazos")
        self.add_p("El personal cuyo turno comience con 'CICLO' no admite reemplazos (ellos mismos son el personal de relevo).", bullet=True)
        self.add_p("Los candidatos a reemplazo se filtran automáticamente: deben tener la misma Dotación, la misma Especialidad y turno 'CICLO%'.", bullet=True)
        self.add_p("Una novedad puede tener múltiples tramos de reemplazo (por ejemplo, los primeros 3 días cubiertos por un agente y los siguientes 4 por otro).", bullet=True)
        self.add_p("Validación de no solapamiento: El sistema rechaza tramos que se superpongan entre sí dentro de la misma novedad o con otras coberturas asignadas previamente al mismo personal de ciclo.", bullet=True)

        self.add_h2("7.2 Gestión de Tramos en el Formulario")
        self.add_step(1, "Habilitar sección de reemplazo", "En el formulario de novedades, la sección inferior contiene la tabla de tramos y el selector de candidatos de ciclo.")
        self.add_step(2, "Seleccionar candidato y fechas", "Elija el personal de ciclo de la lista y defina las fechas de inicio y fin del tramo de cobertura.")
        self.add_step(3, "Agregar tramo", "Haga clic en 'Agregar Tramo'. Si no hay solapamiento, se incorporará a la tabla de tramos de la novedad.")
        self.add_step(4, "Modificar o eliminar tramos", "Para ajustar fechas o cambiar el agente de un tramo, selecciónelo en la lista y utilice los botones 'Modificar tramo' o 'Quitar tramo'.")

        self.add_h2("7.3 Estados de Cobertura en el Sistema")
        self.add_table(
            headers=["Estado", "Indicador Visual", "Criterio de Asignación"],
            rows_data=[
                ["Sí", "Verde", "La novedad tiene cubierto el 100% de los días comprendidos entre inicio y fin."],
                ["Parcial", "Amarillo / Naranja", "Existen tramos asignados pero quedan días sin cobertura dentro del rango."],
                ["No", "Rojo", "La novedad no tiene ningún reemplazo asignado (turno descubierto)."],
                ["N/A (Ciclo)", "Gris", "El titular de la novedad es personal de ciclo; no requiere reemplazo."],
            ],
            col_widths=[1.5, 1.5, 3.2]
        )

        self.add_image_placeholder(
            screen_name="Sección de Reemplazos en Formulario",
            instructions="Capturar el bloque inferior del formulario de novedades con la tabla de tramos de reemplazo y los botones de acción.",
            figure_desc="Administración de tramos de cobertura con personal de ciclo dentro del formulario de novedad."
        )

    def _build_section_8_cambios_turno(self):
        self.add_h1("8. Registro de Cambios de Turno (Paso a Paso)")
        self.add_p(
            "El módulo de Cambios de Turno gestiona las permutas de servicio acordadas entre dos trabajadores, "
            "garantizando la integridad de dotación, especialidad y turnos correspondientes."
        )
        self.add_h2("8.1 Pasos para Registrar un Cambio de Turno")
        self.add_step(1, "Acceder al formulario", "Desde la barra superior o vista de tabla, presione 'Cargar Cambio de Turno'.")
        self.add_step(2, "Seleccionar Empleado Titular (1)", "Presione 'Buscar Personal' para el primer empleado. El sistema cargará su Legajo, Nombre, Especialidad, Dotación, Turno y Franco.")
        self.add_step(3, "Seleccionar Empleado Reemplazante (2)", "Presione 'Buscar Personal' para el segundo empleado. El sistema completará sus datos en la columna derecha.")
        self.add_step(4, "Fecha de cambio", "Seleccione la fecha exacta en la que se efectivizará el cambio de turno mediante el calendario.")
        self.add_step(5, "Referencia y Supervisión", "Indique la estación o personal de referencia en el combobox predictivo y especifique el supervisor que autoriza.")
        self.add_step(6, "Observaciones y Confirmación", "Agregue notas aclaratorias si fuese necesario y presione 'Guardar cambio de turno'.")

        self.add_image_placeholder(
            screen_name="Formulario de Cambio de Turnos",
            instructions="Capturar el formulario de cambio de turno con ambas columnas de empleados completas y fecha seleccionada.",
            figure_desc="Formulario de registro de cambio de turno entre dos empleados."
        )

    def _build_section_9_tablas_filtros(self):
        self.add_h1("9. Tablas Operativas y Filtrado en Vivo")
        self.add_p(
            "Las vistas principales de RENO presentan toda la información organizada en tablas de alta velocidad. "
            "Los registros se muestran ordenados de forma cronológica inversa (los más recientes arriba de todo)."
        )
        self.add_h2("9.1 Barra de Filtros en Tiempo Real")
        self.add_p(
            "En la parte superior de cada tabla se encuentra la barra de filtrado interactivo. "
            "No es necesario presionar 'Buscar': la tabla se actualiza automáticamente al tipear o seleccionar opciones:"
        )
        self.add_p("Filtro por Nombre o Legajo: Permite buscar por cualquier fragmento del apellido, nombre o número de legajo. Es insensible a mayúsculas y acentos.", bullet=True)
        self.add_p("Filtro por Dotación: Desplegable que limita la visualización a una base operativa específica.", bullet=True)
        self.add_p("Filtro por Tipo de Novedad: Exclusivo de la tabla de novedades, permite aislar licencias, informes, etc.", bullet=True)
        self.add_p("Botón Limpiar Filtros: Restablece de forma inmediata todos los filtros mostrando el listado completo.", bullet=True)

        self.add_h2("9.2 Columnas de Control")
        self.add_p(
            "En la tabla de Novedades se destaca la columna 'CUBIERTO', que muestra el estado de cobertura en tiempo real "
            "(Sí, Parcial, No, N/A Ciclo). Las columnas extensas como 'Observaciones' o 'Usuario Windows' se reservan "
            "para el visor de detalle a fin de mantener la tabla limpia y ágil."
        )

        self.add_image_placeholder(
            screen_name="Tabla de Novedades con Filtros",
            instructions="Capturar la tabla de novedades con la barra superior de filtros activa y la columna CUBIERTO visible.",
            figure_desc="Tabla operativa de novedades con filtros dinámicos y columna de cobertura de turnos."
        )

        self.add_image_placeholder(
            screen_name="Tabla de Cambios de Turno",
            instructions="Capturar la vista de tabla de cambios de turno mostrando los datos cruzados de ambos agentes.",
            figure_desc="Tabla operativa de cambios de turno registrados en el sistema."
        )

    def _build_section_10_detalle_edicion(self):
        self.add_h1("10. Detalle, Modificación y Eliminación de Registros")
        self.add_h2("10.1 Modal de Detalle Completo")
        self.add_p(
            "Haga doble clic sobre cualquier fila de la tabla para abrir el modal de Detalle del Registro. "
            "Esta ventana maximizada exhibe la totalidad de los datos: identificación completa de los agentes, "
            "fechas, tramos de reemplazo asignados, texto íntegro de observaciones y el Usuario de Windows que grabó el registro."
        )

        self.add_image_placeholder(
            screen_name="Modal de Detalle del Registro",
            instructions="Capturar la ventana maximizada con el detalle íntegro de una novedad (incluyendo observaciones y tramos de ciclo).",
            figure_desc="Ventana de visualización detallada de un registro seleccionado."
        )

        self.add_h2("10.2 Ventanas de Tiempo y Políticas de Edición")
        self.add_p(
            "Para preservar la seguridad e inmutabilidad de los datos históricos, el sistema aplica reglas estrictas de tiempo:"
        )
        self.add_table(
            headers=["Operación", "Ventana Estándar", "Bypass con Permiso 'registros.modificar'"],
            rows_data=[
                ["Edición completa", "Hasta 24 horas desde la creación.", "Sin límite de tiempo (habilitado para Administrador y Jefatura)."],
                ["Eliminación lógica", "Hasta 72 horas desde la creación.", "Sin límite de tiempo."],
                ["Completar Fecha Fin / Notas", "Fuera de las 24 horas si falta fecha fin.", "Permite cerrar fecha fin y agregar notas acumulativas sin alterar el registro original."],
            ],
            col_widths=[1.8, 2.0, 2.5]
        )

        self.add_h2("10.3 Modo de Edición Restringida")
        self.add_p(
            "Si una novedad fue registrada con fecha de fin abierta o indefinida y ya superó la ventana de 24 horas, "
            "el sistema entra en Modo Restringido: bloquea los datos sensibles del empleado pero permite completar "
            "la Fecha de Fin y añadir nuevas notas que se anexan al final de las observaciones preexistentes."
        )

        self.add_h2("10.4 Eliminación Lógica y Recuperación")
        self.add_p(
            "Cuando se elimina un registro, no se borra físicamente de la base de datos (se marca como inactivo). "
            "Esto asegura que ningún dato se pierda por accidente. Los administradores pueden consultar y restaurar "
            "estos registros desde el menú Administración > Registros eliminados."
        )

    def _build_section_11_personal_ciclo(self):
        self.add_h1("11. Módulo Especial: Gestión y Calendario de Personal de Ciclo")
        self.add_p(
            "El módulo de Personal de Ciclo (accesible con el permiso 'reemplazos.ver') proporciona las herramientas "
            "necesarias para el seguimiento exhaustivo del personal de relevo y el cumplimiento de las normativas de CCT."
        )
        self.add_h2("11.1 Vista de Personal de Ciclo (Disponibles vs Asignados)")
        self.add_p("La vista se divide verticalmente en dos tablas dinámicas:")
        self.add_p("Tabla Superior (Personal Disponible): Agentes de ciclo que no tienen ningún relevo asignado para la fecha de consulta seleccionada.", bullet=True)
        self.add_p("Tabla Inferior (Personal con Relevo Asignado): Agentes que se encuentran actualmente cubriendo una novedad, indicando el ID de novedad y el titular reemplazado.", bullet=True)

        self.add_image_placeholder(
            screen_name="Vista de Personal de Ciclo",
            instructions="Capturar la vista con las dos tablas separadas: disponibles arriba y asignados abajo, con los filtros de fecha y dotación.",
            figure_desc="Monitor de personal de ciclo: disponibles y asignados a relevos."
        )

        self.add_h2("11.2 Calendario Mensual de Ciclo y Descansos CCT")
        self.add_p(
            "Al hacer doble clic sobre cualquier agente de ciclo, se abre el Calendario Mensual de Ciclo. "
            "Presenta una fila por cada día del mes con el cálculo automático de jornadas y descansos según Convenio Colectivo de Trabajo (CCT):"
        )
        self.add_table(
            headers=["Columna del Calendario", "Descripción y Regla de Cálculo"],
            rows_data=[
                ["ESTADO", "Indica si el día es Jornada Habitual, Relevo, Franco Regular o Franco Anticipado."],
                ["ID NOVEDAD", "Número identificador de la novedad que está cubriendo (si aplica)."],
                ["HORA DISPONIBLE", "Hora exacta en la que el agente queda reglamentariamente habilitado para tomar servicio tras el descanso."],
                ["TOMA / DEJA", "Horario de inicio y fin de la prestación diaria."],
                ["HS TRABAJADAS", "Duración de la jornada diaria efectivamente cumplida."],
                ["HS SEMANA", "Contador acumulado de horas trabajadas desde el último franco."],
            ],
            col_widths=[1.8, 4.4]
        )

        self.add_callout(
            "Cálculo de descansos reglamentarios según la jornada del empleado (columna jornada_horas en base):\n"
            "• Para jornada de 6 horas: 18 horas de descanso entre servicios y 42 horas de descanso post-franco.\n"
            "• Para jornada de 7 horas: 17 horas de descanso entre servicios y 41 horas de descanso post-franco.",
            title="REGLAS DE CONVENIO CCT",
            kind="tip"
        )

        self.add_step(1, "Asignar Jornada o Franco Anticipado", "En días sin novedad, seleccione la fecha deseada en el calendario y presione 'Asignar jornada / Franco' (o doble clic en la celda).")
        self.add_step(2, "Configurar horarios", "Ingrese los horarios de Toma y Dejada o marque la casilla 'Franco Anticipado'.")
        self.add_step(3, "Actualización automática", "Al guardar, el calendario recalcula inmediatamente la 'Hora Disponible' del día siguiente y el acumulador de horas de la semana.")

        self.add_image_placeholder(
            screen_name="Calendario Mensual de Personal de Ciclo",
            instructions="Capturar la ventana maximizada del calendario mensual con las columnas Estado, Hora Disponible, Toma, Deja y Hs Semana.",
            figure_desc="Calendario mensual de ciclo con cálculo de descansos reglamentarios y disponibilidad."
        )

    def _build_section_12_export_import(self):
        self.add_h1("12. Reportes, Descubiertos y Exportación a Excel")
        self.add_h2("12.1 Exportar a Excel con Filtros Avanzados")
        self.add_p(
            "Permite generar planillas Excel (.xlsx) totalmente formateadas a partir de la base de datos:"
        )
        self.add_step(1, "Acceder al menú", "Diríjase a Archivo > Exportar a Excel.")
        self.add_step(2, "Seleccionar tabla", "Elija si desea exportar NOVEDADES o Cambio de Turnos.")
        self.add_step(3, "Definir filtros", "Opcionalmente configure rangos de fechas (Desde / Hasta), rango de ID o Tipo de Novedad.")
        self.add_step(4, "Descargar archivo", "Presione 'Exportar' y elija la carpeta de destino donde guardar el archivo.")

        self.add_image_placeholder(
            screen_name="Diálogo de Exportación a Excel",
            instructions="Capturar el cuadro de diálogo de exportación con selector de tabla y filtros de fechas.",
            figure_desc="Diálogo de configuración de exportación de datos a planilla Excel."
        )

        self.add_h2("12.2 Reporte de Turnos Sin Cubrir (Descubiertos)")
        self.add_p(
            "Diseñado para Jefatura y Coordinación Operativa. Permite obtener un listado específico de todas aquellas "
            "novedades que no tienen cobertura o cuya cobertura es sólo parcial para un período y dotación determinados."
        )
        self.add_p("Se accede desde Archivo > Exportar turnos sin cubrir o mediante el botón directo en la cabecera de la tabla de novedades.", bullet=True)
        self.add_p("El archivo Excel resultante resalta las fechas desprotegidas y el personal faltante.", bullet=True)

        self.add_image_placeholder(
            screen_name="Reporte de Turnos Sin Cubrir",
            instructions="Capturar la ventana de filtros para la exportación del reporte de turnos descubiertos.",
            figure_desc="Generación del informe gerencial de turnos sin cubrir."
        )

        self.add_h2("12.3 Importación Operativa de Planillas")
        self.add_p(
            "RENO permite incorporar masivamente datos desde archivos Excel preexistentes mediante las opciones de Archivo > Importar. "
            "El importador es inteligente y tolera variantes comunes en los nombres de encabezados (ej. 'LEGAJO SAP' o 'LEGAJO')."
        )
        self.add_callout(
            "Antes de procesar el archivo, el sistema siempre preguntará si desea eliminar los registros actuales "
            "o anexar la información nueva conservando la existente.",
            title="PREVENCIÓN DE PÉRDIDA DE DATOS",
            kind="warn"
        )

    def _build_section_13_notificaciones(self):
        self.add_h1("13. Sistema de Notificaciones y Registros Nuevos")
        self.add_p(
            "En entornos donde varios puestos de trabajo operan en paralelo, RENO mantiene a los operadores comunicados "
            "sin necesidad de llamadas ni refrescos manuales."
        )
        self.add_h2("13.1 Menú 'Registros nuevos (N)'")
        self.add_p(
            "En la barra superior de menús, el contador (N) indica cuántos registros fueron cargados por otras computadoras "
            "desde el último inicio de sesión del usuario activo (excluyendo automáticamente las cargas propias)."
        )
        self.add_p("Al abrir el menú, se despliega la lista con los últimos registros. Al hacer clic en cualquiera de ellos se abre directamente su Detalle.", bullet=True)
        self.add_p("Al consultar el menú, el contador pasa a (0) indicando que fueron revisados, pero el listado permanece disponible durante toda la sesión.", bullet=True)

        self.add_h2("13.2 Notificaciones Flotantes (Toasts)")
        self.add_p(
            "Cuando otra computadora guarda una nueva novedad, en la esquina inferior derecha de la pantalla aparece un "
            "aviso flotante no bloqueante (toast) con el resumen del registro. Este aviso se cierra automáticamente "
            "tras transcurrir los segundos configurados o al hacer clic sobre él."
        )

        self.add_image_placeholder(
            screen_name="Aviso Toast Flotante y Menú de Novedades",
            instructions="Capturar la esquina inferior de la pantalla mostrando el cartel de notificación emergente y el menú superior con el contador (N).",
            figure_desc="Notificaciones flotantes en tiempo real y menú de registros nuevos pendientes de revisión."
        )

    def _build_section_14_administracion(self):
        self.add_h1("14. Módulo de Administración del Sistema")
        self.add_p(
            "El menú Administración centraliza todas las herramientas de control y configuración. "
            "Solo es visible si el usuario tiene asignados roles de gestión."
        )
        self.add_h2("14.1 Gestión de Usuarios")
        self.add_p("Permite dar de alta cuentas, asociarles un rol, cambiar estados (activo/inactivo) y blanquear claves de acceso.")

        self.add_image_placeholder(
            screen_name="Administración de Usuarios",
            instructions="Capturar la ventana de gestión de usuarios con la tabla de operadores y botones de alta/edición.",
            figure_desc="Ventana de administración y control de usuarios del sistema."
        )

        self.add_h2("14.2 Roles y Matriz Visual de Permisos")
        self.add_p(
            "El editor de roles agrupa los permisos por módulos temáticos (Novedades, Cambios, Empleados, Reemplazos, etc.). "
            "Cuenta con un buscador rápido y botones 'Todos' y 'Ninguno' para facilitar asignaciones masivas con total precisión."
        )

        self.add_image_placeholder(
            screen_name="Editor de Roles y Permisos",
            instructions="Capturar el editor de permisos agrupado por módulos con casillas de verificación.",
            figure_desc="Asignación granular de permisos agrupados por módulo funcional."
        )

        self.add_h2("14.3 Gestión Integral de Empleados (BASE)")
        self.add_p(
            "Permite administrar el padrón maestro del personal. La tabla incorpora filtros dinámicos por Dotación, "
            "Especialidad y búsqueda por Nombre/Legajo. Al dar de alta o editar un empleado se configuran: "
            "Legajo SAP, Apellidos y Nombres, Especialidad, Dotación, Franco, Horario de Franco (para personal de ciclo) "
            "y la Jornada habitual de CCT (6 o 7 horas)."
        )

        self.add_image_placeholder(
            screen_name="Administración de Empleados",
            instructions="Capturar la ventana de empleados mostrando la tabla filtrable y el formulario con datos de CCT.",
            figure_desc="Padrón maestro de empleados y configuración de jornadas de ciclo."
        )

        self.add_h2("14.4 Catálogos del Sistema")
        self.add_p(
            "Administra las listas desplegables utilizadas en los formularios:"
        )
        self.add_p("Tipos de Novedad: Define la descripción, días de duración predeterminada y si entra en el reporte de turnos sin cubrir.", bullet=True)
        self.add_p("Dotaciones: Catálogo de bases y cabeceras operativas.", bullet=True)
        self.add_p("Personal de Estación: Catálogo de jefes y referentes asignados a cada estación, con soporte para importar y exportar a Excel.", bullet=True)

        self.add_h2("14.5 Destinatarios de Informes y Alertas")
        self.add_p("Define las direcciones de correo electrónico que reciben automáticamente los informes de novedades vía Outlook.")

        self.add_h2("14.6 Tiempos de Edición, Sesión y Apariencia")
        self.add_p("Configura los minutos de inactividad de sesión, las horas límite para editar (24h) y eliminar (72h), y la altura en píxeles de las filas de las tablas.")

        self.add_h2("14.7 Registros Eliminados (Papelera)")
        self.add_p("Permite consultar el histórico de registros dados de baja y recuperarlos con un solo clic o eliminarlos de manera permanente.")

        self.add_image_placeholder(
            screen_name="Papelera de Registros Eliminados",
            instructions="Capturar la tabla de registros eliminados con los botones 'Recuperar registro' y 'Eliminar definitivo'.",
            figure_desc="Recuperación y depuración de registros en la papelera del sistema."
        )

        self.add_h2("14.8 Copias de Seguridad (Backups)")
        self.add_p(
            "RENO realiza un respaldo automático diario de la base de datos SQLite al iniciar la aplicación, "
            "conservando las últimas N copias configuradas (por defecto 10). Desde esta ventana también es posible "
            "crear respaldos manuales inmediatos o restaurar una copia previa en caso de contingencia."
        )

        self.add_image_placeholder(
            screen_name="Gestión de Copias de Seguridad",
            instructions="Capturar la ventana de copias de seguridad con el listado de archivos de backup y opciones de restauración.",
            figure_desc="Módulo de respaldo y restauración de bases de datos."
        )

        self.add_h2("14.9 Auditoría de Eventos")
        self.add_p(
            "Bitácora inalterable que registra cada acción realizada en el sistema: usuario, acción (alta, modificación, "
            "borrado, exportación, login), tabla afectada, fecha y hora exacta, y estación Windows interviniente."
        )

        self.add_image_placeholder(
            screen_name="Visor de Auditoría",
            instructions="Capturar la tabla de eventos de auditoría con las columnas Usuario, Acción, Tabla y Fecha.",
            figure_desc="Historial de auditoría para trazabilidad de operaciones."
        )

    def _build_section_15_faq(self):
        self.add_h1("15. Preguntas Frecuentes y Solución de Problemas")
        self.add_table(
            headers=["Mensaje / Situación", "Causa Probable", "Solución Recomendada"],
            rows_data=[
                ["'La base está siendo utilizada por otra PC. Intente nuevamente.'", "Otra computadora está grabando un registro simultáneamente.", "Aguarde unos 5 a 10 segundos y vuelva a presionar Guardar. Las transacciones son ultrarrápidas."],
                ["No veo el menú Administración en la barra superior.", "Su usuario actual no posee permisos administrativos.", "Solicite a un administrador que asigne el rol o permisos correspondientes a su cuenta."],
                ["No puedo editar o eliminar un registro.", "Superó las 24 horas (edición) o 72 horas (eliminación).", "Los usuarios con rol de Administrador o Jefatura (permiso registros.modificar) pueden modificarlo sin límite de tiempo."],
                ["No aparece un registro cargado en otra PC.", "El ciclo de refresco automático es de 60 segundos.", "Espere unos instantes o revise el menú 'Registros nuevos (N)' para ver las últimas incorporaciones."],
                ["Advertencia: 'No se pudo enviar el informe por Outlook'.", "El cliente Outlook no está configurado o está cerrado.", "El registro quedó guardado correctamente en la base. Abra Outlook para restablecer los envíos futuros."],
                ["La aplicación no inicia o da error de base de datos.", "Ruta en 'path_base' incorrecta o carpeta de red desconectada.", "Verifique que la unidad de red compartida esté conectada en Windows y accesible."],
            ],
            col_widths=[1.8, 2.0, 2.5]
        )

    def _build_section_16_anexo_permisos(self):
        self.add_h1("16. Anexo: Matriz de Permisos por Rol")
        self.add_p(
            "A continuación se detalla la nómina de permisos que pueden asignarse a los distintos perfiles de usuario:"
        )
        self.add_table(
            headers=["Módulo", "Acciones Disponibles", "Descripción Operativa"],
            rows_data=[
                ["Novedades", "ver, crear, editar, eliminar, importar, exportar", "Gestión operativa de las novedades del personal."],
                ["Cambios de Turno", "ver, crear, editar, eliminar, importar, exportar", "Gestión operativa de las permutas de servicio."],
                ["Registros (Global)", "modificar, recuperar", "Edición y borrado sin límite de tiempo; acceso a la papelera."],
                ["Reemplazos (Ciclo)", "ver, asignar", "Visualización del módulo de ciclo y asignación de tramos y descansos."],
                ["Empleados", "ver, crear, editar, eliminar, importar, administrar", "Padrón general del personal y datos de CCT."],
                ["Personal Estación", "ver, crear, editar, eliminar, importar, exportar", "Catálogo de personal de referencia por estación."],
                ["Tipos de Novedad", "administrar", "Configuración de catálogo, días predeterminados y descubiertos."],
                ["Dotaciones", "administrar", "Catálogo de bases y dependencias."],
                ["Usuarios y Roles", "administrar", "Creación de cuentas, blanqueo y asignación de permisos."],
                ["Sesión y Tiempos", "configurar", "Límites de inactividad, ventanas de tiempo y contraseñas."],
                ["Backups", "gestionar", "Creación y restauración de copias de seguridad de la base."],
                ["Auditoría", "ver", "Consulta del historial de acciones y seguridad."],
            ],
            col_widths=[1.5, 2.3, 2.5]
        )
        self.add_callout(
            "Fin del documento. Este manual puede ser impreso o exportado a formato PDF "
            "directamente desde Microsoft Word una vez incorporadas las capturas de pantalla deseadas.",
            title="CIERRE DE GUÍA",
            kind="info"
        )


def main():
    parser = argparse.ArgumentParser(description="Generador de Manual de Usuario en Word (.docx) para RENO")
    parser.add_argument(
        "-o", "--output",
        default="Manual/Manual_de_Usuario_RENO_Paso_a_Paso.docx",
        help="Ruta de destino del archivo Word generado (por defecto: Manual/Manual_de_Usuario_RENO_Paso_a_Paso.docx)"
    )
    args = parser.parse_args()

    generator = ManualGenerator(output_path=args.output)
    generator.build_document()


if __name__ == "__main__":
    main()
