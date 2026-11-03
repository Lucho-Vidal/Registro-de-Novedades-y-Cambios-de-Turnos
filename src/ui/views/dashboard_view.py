"""Panel de Control (Dashboard) con KPIs, gráficos interactivos con QPainter y exportación detallada.

Permite visualizar el volumen de novedades y cambios, desgloses por dotación, tipo y día,
y exportar gráficos en PNG, informe ejecutivo PDF con ReportLab y detalle analítico en Excel.
"""

from typing import Any, Dict, List, Optional, Tuple
from datetime import datetime

from PySide6.QtCore import QPointF, QRectF, QSize, Qt, Signal
from PySide6.QtGui import QBrush, QColor, QFont, QPainter, QPaintEvent, QPen
from PySide6.QtWidgets import (
    QComboBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from src.core.logger import get_logger
from src.services.excel_service import ExcelService
from src.services.pdf_service import PDFService
from src.services.records_service import RecordsService

logger = get_logger("DashboardView")


class BaseChartWidget(QWidget):
    """Widget base de dibujado con QPainter y soporte para exportar como imagen PNG."""

    def __init__(self, title: str, parent=None):
        super().__init__(parent)
        self.title = title
        self.setMinimumSize(320, 260)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

    def export_as_image(self, file_path: str) -> None:
        pixmap = self.grab()
        pixmap.save(file_path, "PNG")


class BarChartHWidget(BaseChartWidget):
    """Gráfico de barras horizontales optimizado para tipos de novedad, dotaciones y días."""

    def __init__(self, title: str, bar_color: QColor, parent=None):
        super().__init__(title, parent)
        self.bar_color = bar_color
        self._data: List[Tuple[str, int]] = []

    def set_data(self, data: List[Tuple[str, int]]) -> None:
        self._data = list(data)
        self.update()

    def paintEvent(self, event: QPaintEvent) -> None:
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        w = self.width()
        h = self.height()

        # Fondo de la tarjeta
        painter.fillRect(0, 0, w, h, QColor("#FFFFFF"))
        painter.setPen(QPen(QColor("#E5E7EB"), 1))
        painter.drawRoundedRect(0, 0, w - 1, h - 1, 6, 6)

        # Título
        painter.setPen(QColor("#1F497D"))
        font_title = QFont("Segoe UI", 10, QFont.Bold)
        painter.setFont(font_title)
        painter.drawText(14, 24, self.title)

        if not self._data:
            painter.setPen(QColor("#9CA3AF"))
            painter.setFont(QFont("Segoe UI", 9))
            painter.drawText(QRectF(0, 30, w, h - 30), Qt.AlignCenter, "Sin datos para el período")
            return

        # Dibujar barras
        top_margin = 38
        bottom_margin = 15
        left_margin = 100
        right_margin = 45

        available_h = h - top_margin - bottom_margin
        available_w = max(10, w - left_margin - right_margin)
        max_val = max(v for _, v in self._data) or 1
        num_items = len(self._data)
        row_height = available_h / num_items

        font_label = QFont("Segoe UI", 8)
        font_val = QFont("Segoe UI", 8, QFont.Bold)

        for i, (label, val) in enumerate(self._data):
            y = top_margin + i * row_height
            bar_h = max(4, row_height * 0.6)
            bar_y = y + (row_height - bar_h) / 2
            bar_w = (val / max_val) * available_w

            # Etiqueta
            painter.setPen(QColor("#374151"))
            painter.setFont(font_label)
            rect_label = QRectF(6, y, left_margin - 12, row_height)
            label_text = str(label)
            if len(label_text) > 16:
                label_text = label_text[:14] + "..."
            painter.drawText(rect_label, Qt.AlignRight | Qt.AlignVCenter, label_text)

            # Barra
            painter.setBrush(QBrush(self.bar_color))
            painter.setPen(Qt.NoPen)
            painter.drawRoundedRect(QRectF(left_margin, bar_y, bar_w, bar_h), 3, 3)

            # Valor numérico
            painter.setPen(QColor("#1F2937"))
            painter.setFont(font_val)
            painter.drawText(left_margin + bar_w + 6, bar_y + bar_h - 2, str(val))


class TrendChartVWidget(BaseChartWidget):
    """Gráfico de barras verticales apiladas / serie temporal para la tendencia diaria."""

    def __init__(self, title: str, parent=None):
        super().__init__(title, parent)
        self._data: List[Tuple[str, int, int]] = []  # (fecha, novedades, cambios)
        self.color_nov = QColor("#1F497D")
        self.color_cam = QColor("#4682B4")

    def set_data(self, data: List[Tuple[str, int, int]]) -> None:
        self._data = list(data)
        self.update()

    def paintEvent(self, event: QPaintEvent) -> None:
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        w = self.width()
        h = self.height()

        painter.fillRect(0, 0, w, h, QColor("#FFFFFF"))
        painter.setPen(QPen(QColor("#E5E7EB"), 1))
        painter.drawRoundedRect(0, 0, w - 1, h - 1, 6, 6)

        # Título
        painter.setPen(QColor("#1F497D"))
        painter.setFont(QFont("Segoe UI", 10, QFont.Bold))
        painter.drawText(14, 24, self.title)

        # Leyenda
        painter.setFont(QFont("Segoe UI", 8))
        # Novedades
        painter.setBrush(QBrush(self.color_nov))
        painter.setPen(Qt.NoPen)
        painter.drawRect(w - 180, 14, 10, 10)
        painter.setPen(QColor("#374151"))
        painter.drawText(w - 165, 23, "Novedades")

        # Cambios
        painter.setBrush(QBrush(self.color_cam))
        painter.setPen(Qt.NoPen)
        painter.drawRect(w - 95, 14, 10, 10)
        painter.setPen(QColor("#374151"))
        painter.drawText(w - 80, 23, "Cambios")

        if not self._data:
            painter.setPen(QColor("#9CA3AF"))
            painter.drawText(QRectF(0, 30, w, h - 30), Qt.AlignCenter, "Sin datos para el período")
            return

        top_margin = 42
        bottom_margin = 28
        left_margin = 35
        right_margin = 15

        area_h = h - top_margin - bottom_margin
        area_w = w - left_margin - right_margin
        total_items = len(self._data)
        max_total = max((n + c) for _, n, c in self._data) or 1

        step_w = area_w / total_items
        bar_w = max(2.0, step_w * 0.65)
        label_step = max(1, total_items // 6)

        font_date = QFont("Segoe UI", 7)

        for i, (fecha_str, nov_cnt, cam_cnt) in enumerate(self._data):
            x = left_margin + i * step_w + (step_w - bar_w) / 2
            y_base = top_margin + area_h

            h_nov = (nov_cnt / max_total) * area_h
            h_cam = (cam_cnt / max_total) * area_h

            # Barra Novedades
            if nov_cnt > 0:
                painter.setBrush(QBrush(self.color_nov))
                painter.setPen(Qt.NoPen)
                painter.drawRect(QRectF(x, y_base - h_nov, bar_w, h_nov))

            # Barra Cambios apilada arriba
            if cam_cnt > 0:
                painter.setBrush(QBrush(self.color_cam))
                painter.setPen(Qt.NoPen)
                painter.drawRect(QRectF(x, y_base - h_nov - h_cam, bar_w, h_cam))

            # Etiquetas de fecha cada X intervalos
            if i % label_step == 0 or i == total_items - 1:
                partes = fecha_str.split("-")
                etiqueta = f"{partes[2]}/{partes[1]}" if len(partes) == 3 else fecha_str
                painter.setPen(QColor("#6B7280"))
                painter.setFont(font_date)
                painter.drawText(QRectF(x - 12, y_base + 4, bar_w + 24, 16), Qt.AlignCenter, etiqueta)


class DashboardView(QWidget):
    """Vista principal del Panel de Control de RENO v2.0.0."""

    navegar_novedades = Signal()
    navegar_cambios = Signal()
    navegar_ciclo = Signal()

    def __init__(
        self,
        records_service: RecordsService,
        excel_service: ExcelService,
        pdf_service: PDFService,
        parent=None,
    ):
        super().__init__(parent)
        self.records_service = records_service
        self.excel_service = excel_service
        self.pdf_service = pdf_service
        self.tendencia_dias = 30
        self._metricas_actuales: Dict[str, Any] = {}

        self._init_ui()

    def _init_ui(self) -> None:
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(16, 16, 16, 16)
        main_layout.setSpacing(12)

        # 1. Cabecera y Barra de Acciones / Exportación
        header_layout = QHBoxLayout()
        lbl_titulo = QLabel("Panel de Control")
        lbl_titulo.setStyleSheet("font-size: 18pt; font-weight: bold; color: #1F497D;")
        header_layout.addWidget(lbl_titulo)

        header_layout.addStretch()

        # Selector de período
        lbl_periodo = QLabel("Período tendencia:")
        lbl_periodo.setStyleSheet("font-weight: 600; color: #4B5563;")
        header_layout.addWidget(lbl_periodo)

        self.combo_periodo = QComboBox()
        self.combo_periodo.addItems(["7 días", "30 días", "90 días"])
        self.combo_periodo.setCurrentIndex(1)
        self.combo_periodo.currentIndexChanged.connect(self._cambiar_periodo)
        self.combo_periodo.setFixedWidth(100)
        header_layout.addWidget(self.combo_periodo)

        # Botones de Exportación (Requerimiento de usuario)
        btn_export_excel = QPushButton("📊 Exportar Detalle Excel")
        btn_export_excel.setObjectName("btnSecondary")
        btn_export_excel.setToolTip("Exporta libro de cálculo completo con KPIs y series temporales")
        btn_export_excel.clicked.connect(self.exportar_excel_detallado)
        header_layout.addWidget(btn_export_excel)

        btn_export_pdf = QPushButton("📑 Informe Ejecutivo PDF")
        btn_export_pdf.setObjectName("btnSecondary")
        btn_export_pdf.setToolTip("Genera informe formal en PDF con tablas y gráficos")
        btn_export_pdf.clicked.connect(self.exportar_pdf_ejecutivo)
        header_layout.addWidget(btn_export_pdf)

        btn_export_img = QPushButton("🖼️ Exportar Gráficos PNG")
        btn_export_img.setObjectName("btnSecondary")
        btn_export_img.setToolTip("Exporta imágenes PNG individuales de los gráficos")
        btn_export_img.clicked.connect(self.exportar_graficos_png)
        header_layout.addWidget(btn_export_img)

        main_layout.addLayout(header_layout)

        # 2. Tarjetas KPI
        cards_layout = QHBoxLayout()
        cards_layout.setSpacing(12)

        self.card_novedades_hoy = self._crear_tarjeta_kpi("NOVEDADES HOY", "0", "#1F497D", "Hoy")
        self.card_cambios_hoy = self._crear_tarjeta_kpi("CAMBIOS HOY", "0", "#4682B4", "Hoy")
        self.card_total_mes = self._crear_tarjeta_kpi("TOTAL DEL MES", "0", "#2E7D32", "Mes actual")
        self.card_sin_revisar = self._crear_tarjeta_kpi("SIN REVISAR", "0", "#F57C00", "Pendientes")

        cards_layout.addWidget(self.card_novedades_hoy)
        cards_layout.addWidget(self.card_cambios_hoy)
        cards_layout.addWidget(self.card_total_mes)
        cards_layout.addWidget(self.card_sin_revisar)
        main_layout.addLayout(cards_layout)

        # 3. Gráficos en ScrollArea o Grid de 2x2
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)

        container = QWidget()
        self.grid_charts = QGridLayout(container)
        self.grid_charts.setContentsMargins(0, 8, 0, 0)
        self.grid_charts.setSpacing(12)

        # Gráfico 1: Novedades por tipo
        self.chart_tipo = BarChartHWidget("Novedades por Tipo", QColor("#1F497D"))
        # Gráfico 2: Registros por dotación
        self.chart_dotacion = BarChartHWidget("Registros por Dotación (Top 10)", QColor("#2E7D32"))
        # Gráfico 3: Tendencia diaria
        self.chart_tendencia = TrendChartVWidget("Tendencia de Cargas Diarias")
        # Gráfico 4: Registros por día de la semana
        self.chart_dia_semana = BarChartHWidget("Registros por Día de la Semana", QColor("#F57C00"))

        self.grid_charts.addWidget(self.chart_tipo, 0, 0)
        self.grid_charts.addWidget(self.chart_dotacion, 0, 1)
        self.grid_charts.addWidget(self.chart_tendencia, 1, 0)
        self.grid_charts.addWidget(self.chart_dia_semana, 1, 1)

        scroll.setWidget(container)
        main_layout.addWidget(scroll)

        # Timestamp de actualización
        self.lbl_timestamp = QLabel("")
        self.lbl_timestamp.setStyleSheet("color: #6B7280; font-size: 8.5pt;")
        self.lbl_timestamp.setAlignment(Qt.AlignRight)
        main_layout.addWidget(self.lbl_timestamp)

    def _crear_tarjeta_kpi(self, titulo: str, valor_inicial: str, color_borde: str, subtitulo: str) -> QFrame:
        card = QFrame()
        card.setObjectName("kpiCard")
        card.setStyleSheet(f"""
            QFrame#kpiCard {{
                background-color: #FFFFFF;
                border-left: 4px solid {color_borde};
                border-top: 1px solid #E5E7EB;
                border-right: 1px solid #E5E7EB;
                border-bottom: 1px solid #E5E7EB;
                border-radius: 6px;
                padding: 10px;
            }}
        """)
        layout = QVBoxLayout(card)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(4)

        lbl_tit = QLabel(titulo)
        lbl_tit.setStyleSheet("font-size: 8.5pt; font-weight: bold; color: #6B7280;")

        lbl_val = QLabel(valor_inicial)
        lbl_val.setStyleSheet(f"font-size: 18pt; font-weight: bold; color: {color_borde};")
        card.lbl_val = lbl_val  # type: ignore

        lbl_sub = QLabel(subtitulo)
        lbl_sub.setStyleSheet("font-size: 7.5pt; color: #9CA3AF;")

        layout.addWidget(lbl_tit)
        layout.addWidget(lbl_val)
        layout.addWidget(lbl_sub)
        return card

    def _cambiar_periodo(self, index: int) -> None:
        dias_map = {0: 7, 1: 30, 2: 90}
        self.tendencia_dias = dias_map.get(index, 30)
        self.actualizar_dashboard()

    def actualizar_dashboard(self, no_revisados_count: int = 0) -> None:
        """Carga las métricas calculadas por el RecordsService y refresca widgets."""
        try:
            metricas = self.records_service.obtener_metricas_dashboard(self.tendencia_dias)
            self._metricas_actuales = metricas

            resumen = metricas.get("resumen", {})
            self.card_novedades_hoy.lbl_val.setText(str(resumen.get("novedades_hoy", 0)))  # type: ignore
            self.card_cambios_hoy.lbl_val.setText(str(resumen.get("cambios_hoy", 0)))  # type: ignore
            self.card_total_mes.lbl_val.setText(str(resumen.get("total_mes", 0)))  # type: ignore
            self.card_sin_revisar.lbl_val.setText(str(no_revisados_count))  # type: ignore

            self.chart_tipo.set_data(metricas.get("por_tipo", [])[:12])
            self.chart_dotacion.set_data(metricas.get("por_dotacion", []))
            self.chart_tendencia.set_data(metricas.get("tendencia", []))
            self.chart_dia_semana.set_data(metricas.get("por_dia_semana", []))

            self.lbl_timestamp.setText(f"Última actualización: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
        except Exception as e:
            logger.error(f"Error al actualizar dashboard: {e}")

    # ---------------------------------------------------------
    # Acciones de Exportación Detallada (Requerimiento de Usuario)
    # ---------------------------------------------------------

    def exportar_excel_detallado(self) -> None:
        if not self._metricas_actuales:
            QMessageBox.warning(self, "Exportación", "No hay métricas disponibles para exportar.")
            return

        default_name = f"Dashboard_Detalle_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        file_path, _ = QFileDialog.getSaveFileName(self, "Guardar Detalle de Dashboard Excel", default_name, "Libro de Excel (*.xlsx)")
        if not file_path:
            return

        try:
            self.excel_service.export_dashboard_detailed(
                self._metricas_actuales, file_path, periodo_dias=self.tendencia_dias
            )
            QMessageBox.information(self, "Exportación Exitosa", f"Archivo Excel detallado generado con éxito en:\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "Error al Exportar", f"No se pudo generar el archivo Excel:\n{e}")

    def exportar_pdf_ejecutivo(self) -> None:
        if not self._metricas_actuales:
            QMessageBox.warning(self, "Exportación", "No hay métricas disponibles para exportar.")
            return

        default_name = f"Informe_Ejecutivo_Dashboard_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
        file_path, _ = QFileDialog.getSaveFileName(self, "Guardar Informe Ejecutivo PDF", default_name, "Archivo PDF (*.pdf)")
        if not file_path:
            return

        try:
            # Capturar imágenes temporales de los 4 gráficos para incrustar en el PDF
            import tempfile
            import os
            temp_dir = tempfile.gettempdir()
            img_paths = [
                os.path.join(temp_dir, "reno_chart_tipo.png"),
                os.path.join(temp_dir, "reno_chart_dot.png"),
                os.path.join(temp_dir, "reno_chart_trend.png"),
                os.path.join(temp_dir, "reno_chart_dia.png"),
            ]
            self.chart_tipo.export_as_image(img_paths[0])
            self.chart_dotacion.export_as_image(img_paths[1])
            self.chart_tendencia.export_as_image(img_paths[2])
            self.chart_dia_semana.export_as_image(img_paths[3])

            self.pdf_service.generar_informe_dashboard_pdf(
                self._metricas_actuales, file_path, periodo_dias=self.tendencia_dias, chart_image_paths=img_paths
            )
            QMessageBox.information(self, "Exportación Exitosa", f"Informe ejecutivo PDF generado con éxito en:\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "Error al Exportar", f"No se pudo generar el informe PDF:\n{e}")

    def exportar_graficos_png(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "Seleccionar Carpeta para Guardar Gráficos PNG")
        if not folder:
            return

        try:
            import os
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            p1 = os.path.join(folder, f"grafico_novedades_tipo_{ts}.png")
            p2 = os.path.join(folder, f"grafico_registros_dotacion_{ts}.png")
            p3 = os.path.join(folder, f"grafico_tendencia_{ts}.png")
            p4 = os.path.join(folder, f"grafico_dia_semana_{ts}.png")

            self.chart_tipo.export_as_image(p1)
            self.chart_dotacion.export_as_image(p2)
            self.chart_tendencia.export_as_image(p3)
            self.chart_dia_semana.export_as_image(p4)

            QMessageBox.information(self, "Exportación Exitosa", f"Se guardaron 4 gráficos en alta resolución en:\n{folder}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudieron exportar los gráficos:\n{e}")
