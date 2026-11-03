"""Vista principal de Cambios de Turno con grilla interactiva, filtros y generación documental."""

from typing import Any, Dict, List, Optional
from datetime import datetime

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QComboBox,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QTableView,
    QToolBar,
    QVBoxLayout,
    QWidget,
)

from src.core.logger import get_logger
from src.services.excel_service import ExcelService
from src.services.pdf_service import PDFService
from src.services.records_service import RecordsService
from src.services.word_service import WordService
from src.ui.models.base_table_model import BaseTableModel
from src.ui.models.multi_filter_proxy import MultiFilterProxyModel
from src.ui.views.dialogs.cambio_turno_dialog import CambioTurnoDialog

logger = get_logger("CambiosTurnoView")


class CambiosTurnoView(QWidget):
    """Vista principal con grilla de cambios de turno (permutas) y exportaciones."""

    registro_modificado = Signal()

    def __init__(
        self,
        records_service: RecordsService,
        excel_service: ExcelService,
        pdf_service: PDFService,
        word_service: WordService,
        usuario_actual: Optional[Dict[str, Any]] = None,
        tiene_permiso_fn=None,
        parent=None,
    ):
        super().__init__(parent)
        self.records_service = records_service
        self.excel_service = excel_service
        self.pdf_service = pdf_service
        self.word_service = word_service
        self.usuario_actual = usuario_actual or {}
        self.tiene_permiso = tiene_permiso_fn or (lambda p: True)

        self._init_ui()
        self.recargar_datos()

    def _init_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)

        # 1. Toolbar
        toolbar = QToolBar()
        toolbar.setStyleSheet("spacing: 6px;")

        self.btn_nuevo = QPushButton("➕ Nuevo Cambio")
        self.btn_nuevo.setEnabled(self.tiene_permiso("cambios_turno.crear"))
        self.btn_nuevo.clicked.connect(self._nuevo_cambio)
        toolbar.addWidget(self.btn_nuevo)

        self.btn_editar = QPushButton("✏️ Modificar")
        self.btn_editar.setObjectName("btnSecondary")
        self.btn_editar.clicked.connect(self._modificar_cambio)
        toolbar.addWidget(self.btn_editar)

        self.btn_eliminar = QPushButton("🗑️ Eliminar")
        self.btn_eliminar.setObjectName("btnDanger")
        self.btn_eliminar.clicked.connect(self._eliminar_cambio)
        toolbar.addWidget(self.btn_eliminar)

        toolbar.addSeparator()

        self.btn_export_excel = QPushButton("📊 Exportar Excel")
        self.btn_export_excel.setObjectName("btnSecondary")
        self.btn_export_excel.setEnabled(self.tiene_permiso("cambios_turno.exportar"))
        self.btn_export_excel.clicked.connect(self._exportar_excel)
        toolbar.addWidget(self.btn_export_excel)

        self.btn_export_pdf = QPushButton("📑 Constancia PDF")
        self.btn_export_pdf.setObjectName("btnSecondary")
        self.btn_export_pdf.clicked.connect(self._exportar_pdf)
        toolbar.addWidget(self.btn_export_pdf)

        self.btn_export_docx = QPushButton("📝 Solicitud Word (.docx)")
        self.btn_export_docx.setObjectName("btnSecondary")
        self.btn_export_docx.clicked.connect(self._exportar_docx)
        toolbar.addWidget(self.btn_export_docx)

        layout.addWidget(toolbar)

        # 2. Barra de Filtros
        filter_frame = QFrame()
        filter_frame.setStyleSheet("background-color: #FFFFFF; border: 1px solid #E5E7EB; border-radius: 4px; padding: 4px;")
        filter_layout = QHBoxLayout(filter_frame)
        filter_layout.setContentsMargins(6, 4, 6, 4)
        filter_layout.setSpacing(8)

        filter_layout.addWidget(QLabel("Buscar:"))
        self.txt_search = QLineEdit()
        self.txt_search.setPlaceholderText("Texto libre...")
        self.txt_search.textChanged.connect(self._on_search_changed)
        self.txt_search.setFixedWidth(180)
        filter_layout.addWidget(self.txt_search)

        filter_layout.addWidget(QLabel("Dotación:"))
        self.combo_dotacion = QComboBox()
        self.combo_dotacion.addItem("Todas")
        self.combo_dotacion.currentIndexChanged.connect(self._on_filter_changed)
        filter_layout.addWidget(self.combo_dotacion)

        self.btn_limpiar = QPushButton("Limpiar")
        self.btn_limpiar.setObjectName("btnSecondary")
        self.btn_limpiar.clicked.connect(self._limpiar_filtros)
        filter_layout.addWidget(self.btn_limpiar)

        filter_layout.addStretch()
        layout.addWidget(filter_frame)

        # 3. Tabla
        headers = [
            "ID", "REGISTRADO EN", "LEG. 1", "APELLIDOS Y NOMBRES 1", "ESP. 1", "DOT. 1", "TURNO 1",
            "LEG. 2", "APELLIDOS Y NOMBRES 2", "ESP. 2", "DOT. 2", "TURNO 2",
            "FECHA CAMBIO", "REF. ESTACIÓN", "SUPERVISOR", "OBSERVACIONES",
        ]
        keys = [
            "id", "registrado_en", "legajo_1", "apellidos_nombres_1", "especialidad_1", "dotacion_1", "turnos_1",
            "legajo_2", "apellidos_nombres_2", "especialidad_2", "dotacion_2", "turnos_2",
            "fecha_cambio", "referencia_estacion", "supervisor", "observaciones",
        ]

        self.table_model = BaseTableModel(headers, keys, parent=self)
        self.proxy_model = MultiFilterProxyModel(self)
        self.proxy_model.setSourceModel(self.table_model)

        self.table_view = QTableView()
        self.table_view.setModel(self.proxy_model)
        self.table_view.setSelectionBehavior(QTableView.SelectRows)
        self.table_view.setSelectionMode(QTableView.SingleSelection)
        self.table_view.setSortingEnabled(True)
        self.table_view.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        self.table_view.horizontalHeader().setSectionResizeMode(8, QHeaderView.Stretch)
        self.table_view.doubleClicked.connect(self._on_double_click)
        layout.addWidget(self.table_view)

        # Conteo
        self.lbl_conteo = QLabel("Mostrando 0 registros")
        self.lbl_conteo.setStyleSheet("color: #6B7280; font-size: 8.5pt;")
        layout.addWidget(self.lbl_conteo)

    def recargar_datos(self) -> None:
        cambios = self.records_service.listar_cambios()
        self.table_model.set_data(cambios)
        self._actualizar_combo_dotaciones()
        self._actualizar_conteo()

    def _actualizar_combo_dotaciones(self) -> None:
        dots = self.records_service.listar_dotaciones(incluir_inactivos=False)
        cur = self.combo_dotacion.currentText()
        self.combo_dotacion.blockSignals(True)
        self.combo_dotacion.clear()
        self.combo_dotacion.addItem("Todas")
        for d in dots:
            self.combo_dotacion.addItem(d["nombre"])
        idx = self.combo_dotacion.findText(cur)
        if idx >= 0:
            self.combo_dotacion.setCurrentIndex(idx)
        self.combo_dotacion.blockSignals(False)

    def _on_search_changed(self, text: str) -> None:
        self.proxy_model.set_global_search(text)
        self._actualizar_conteo()

    def _on_filter_changed(self) -> None:
        # Columna 5 = Dotación 1
        dot = self.combo_dotacion.currentText()
        self.proxy_model.set_column_filter(5, dot)
        self._actualizar_conteo()

    def _limpiar_filtros(self) -> None:
        self.txt_search.clear()
        self.combo_dotacion.setCurrentIndex(0)
        self.proxy_model.clear_all_filters()
        self._actualizar_conteo()

    def _actualizar_conteo(self) -> None:
        visibles = self.proxy_model.rowCount()
        totales = self.table_model.rowCount()
        self.lbl_conteo.setText(f"Mostrando {visibles} de {totales} cambios de turno registrados")

    def _obtener_fila_seleccionada(self) -> Optional[Dict[str, Any]]:
        indexes = self.table_view.selectionModel().selectedRows()
        if not indexes:
            return None
        proxy_idx = indexes[0]
        source_idx = self.proxy_model.mapToSource(proxy_idx)
        return self.table_model.get_row(source_idx.row())

    def _nuevo_cambio(self) -> None:
        dlg = CambioTurnoDialog(
            self.records_service,
            usuario_actual=self.usuario_actual,
            parent=self,
        )
        if dlg.exec():
            self.recargar_datos()
            self.registro_modificado.emit()

    def _modificar_cambio(self) -> None:
        fila = self._obtener_fila_seleccionada()
        if not fila:
            QMessageBox.information(self, "Modificar", "Seleccione un cambio de turno en la tabla.")
            return

        puede_modificar_libre = self.tiene_permiso("registros.modificar")
        dentro_ventana = self.records_service.dentro_de_ventana(fila["registrado_en"], "editar_horas")

        if not puede_modificar_libre and not dentro_ventana:
            QMessageBox.warning(
                self,
                "Tiempo Expirado",
                "El registro superó el tiempo límite permitido para ser modificado.",
            )
            return

        dlg = CambioTurnoDialog(
            self.records_service,
            cambio_existente=fila,
            usuario_actual=self.usuario_actual,
            parent=self,
        )
        if dlg.exec():
            self.recargar_datos()
            self.registro_modificado.emit()

    def _eliminar_cambio(self) -> None:
        fila = self._obtener_fila_seleccionada()
        if not fila:
            QMessageBox.information(self, "Eliminar", "Seleccione un cambio de turno en la tabla.")
            return

        puede_eliminar_libre = self.tiene_permiso("registros.modificar")
        dentro_ventana = self.records_service.dentro_de_ventana(fila["registrado_en"], "eliminar_horas")

        if not puede_eliminar_libre and not dentro_ventana:
            QMessageBox.warning(
                self,
                "Tiempo Expirado",
                "El registro superó el límite de tiempo permitido para ser eliminado.",
            )
            return

        resp = QMessageBox.question(
            self,
            "Confirmar Eliminación",
            f"¿Está seguro de que desea eliminar el cambio de turno #{fila['id']} "
            f"entre {fila['apellidos_nombres_1']} y {fila['apellidos_nombres_2']}?",
            QMessageBox.Yes | QMessageBox.No,
        )
        if resp != QMessageBox.Yes:
            return

        try:
            self.records_service.eliminar_cambio(
                fila["id"],
                user_id=self.usuario_actual.get("id"),
                windows_user=self.usuario_actual.get("username"),
                bypass_window=puede_eliminar_libre,
            )
            self.recargar_datos()
            self.registro_modificado.emit()
            QMessageBox.information(self, "Eliminado", "El cambio de turno fue enviado a registros eliminados.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo eliminar el cambio de turno:\n{e}")

    def _exportar_excel(self) -> None:
        default_name = f"Cambios_Turno_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        file_path, _ = QFileDialog.getSaveFileName(self, "Exportar Cambios de Turno a Excel", default_name, "Libro de Excel (*.xlsx)")
        if not file_path:
            return
        try:
            self.excel_service.export_database(file_path, tables=("Cambio de Turnos",))
            QMessageBox.information(self, "Exportación", f"Cambios exportados exitosamente a:\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo exportar a Excel:\n{e}")

    def _exportar_pdf(self) -> None:
        fila = self._obtener_fila_seleccionada()
        if not fila:
            QMessageBox.information(self, "Exportar PDF", "Seleccione un cambio en la tabla para generar su constancia.")
            return

        default_name = f"Constancia_Cambio_{fila['id']}_{datetime.now().strftime('%Y%m%d')}.pdf"
        file_path, _ = QFileDialog.getSaveFileName(self, "Guardar Constancia PDF", default_name, "Archivo PDF (*.pdf)")
        if not file_path:
            return

        try:
            self.pdf_service.generar_constancia_cambio(fila, file_path)
            QMessageBox.information(self, "Constancia PDF", f"Constancia PDF generada exitosamente en:\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo generar la constancia PDF:\n{e}")

    def _exportar_docx(self) -> None:
        fila = self._obtener_fila_seleccionada()
        if not fila:
            QMessageBox.information(self, "Generar Word", "Seleccione un cambio en la tabla para generar su formulario Word.")
            return

        default_name = f"Solicitud_Cambio_{fila['id']}_{datetime.now().strftime('%Y%m%d')}.docx"
        file_path, _ = QFileDialog.getSaveFileName(self, "Guardar Formulario Word", default_name, "Documento de Word (*.docx)")
        if not file_path:
            return

        try:
            self.word_service.generar_solicitud_cambio_turno_docx(fila, file_path)
            QMessageBox.information(self, "Documento Word", f"Formulario Word generado exitosamente en:\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo generar el documento Word:\n{e}")

    def _on_double_click(self, index) -> None:
        self._modificar_cambio()
