"""Vista principal de Novedades Operativas con tabla interactiva, filtros y acciones."""

from typing import Any, Dict, List, Optional
from datetime import date, datetime

from PySide6.QtCore import QDate, Qt, Signal
from PySide6.QtWidgets import (
    QComboBox,
    QDateEdit,
    QDialog,
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

from src.core.config import obtener_usuario_windows
from src.core.logger import get_logger
from src.domain.validators import parsear_fecha
from src.services.excel_service import ExcelService
from src.services.pdf_service import PDFService
from src.services.records_service import RecordsService
from src.ui.models.base_table_model import BaseTableModel
from src.ui.models.multi_filter_proxy import MultiFilterProxyModel
from src.ui.views.dialogs.novedad_dialog import NovedadDialog

logger = get_logger("NovedadesView")


class TurnosSinCubrirDialog(QDialog):
    """Diálogo modal para consultar y exportar turnos sin cubrir en una fecha dada."""

    def __init__(
        self,
        records_service: RecordsService,
        excel_service: ExcelService,
        dotaciones: List[str],
        parent=None,
    ):
        super().__init__(parent)
        self.records_service = records_service
        self.excel_service = excel_service
        self.dotaciones = dotaciones
        self.setWindowTitle("Turnos Operativos Sin Cubrir")
        self.resize(950, 550)

        layout = QVBoxLayout(self)

        # Barra de consulta
        bar = QHBoxLayout()
        bar.addWidget(QLabel("Fecha de consulta:"))
        self.date_picker = QDateEdit()
        self.date_picker.setCalendarPopup(True)
        self.date_picker.setDisplayFormat("dd/MM/yyyy")
        self.date_picker.setDate(QDate.currentDate())
        bar.addWidget(self.date_picker)

        bar.addWidget(QLabel("Dotación:"))
        self.combo_dot = QComboBox()
        self.combo_dot.addItem("Todas")
        self.combo_dot.addItems(dotaciones)
        bar.addWidget(self.combo_dot)

        self.btn_consultar = QPushButton("🔍 Consultar")
        self.btn_consultar.clicked.connect(self._consultar)
        bar.addWidget(self.btn_consultar)

        bar.addStretch()

        self.btn_export = QPushButton("📊 Exportar a Excel")
        self.btn_export.setObjectName("btnSecondary")
        self.btn_export.clicked.connect(self._exportar_excel)
        bar.addWidget(self.btn_export)

        layout.addLayout(bar)

        # Tabla
        headers = [
            "LEGAJO", "APELLIDOS Y NOMBRES", "ESPECIALIDAD", "DOTACIÓN", "TURNO",
            "FRANCO", "NOVEDAD", "DESDE", "HASTA", "CUBIERTO", "REEMPLAZO ASIGNADO",
        ]
        keys = [
            "legajo", "apellidos_nombres", "especialidad", "dotacion", "turnos",
            "franco", "novedad", "fecha_inicio", "fecha_fin", "cubierto", "reemplazo",
        ]
        self.model = BaseTableModel(headers, keys, [], parent=self)
        self.table = QTableView()
        self.table.setModel(self.model)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        layout.addWidget(self.table)

        self._consultar()

    def _consultar(self) -> None:
        f_str = self.date_picker.date().toString("dd/MM/yyyy")
        dot = self.combo_dot.currentText()
        try:
            descubiertos = self.records_service.obtener_turnos_sin_cubrir(f_str, dotacion=dot)
            self.model.set_data(descubiertos)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo consultar turnos sin cubrir:\n{e}")

    def _exportar_excel(self) -> None:
        f_str = self.date_picker.date().toString("dd/MM/yyyy")
        dot = self.combo_dot.currentText()
        filas = self.model.get_all_data()
        if not filas:
            QMessageBox.information(self, "Exportar", "No hay turnos para exportar en la fecha seleccionada.")
            return

        default_name = f"Turnos_Sin_Cubrir_{f_str.replace('/', '-')}.xlsx"
        file_path, _ = QFileDialog.getSaveFileName(self, "Guardar Reporte Excel", default_name, "Libro de Excel (*.xlsx)")
        if not file_path:
            return

        try:
            self.excel_service.export_turnos_sin_cubrir(file_path, filas, f_str, dotacion=dot)
            QMessageBox.information(self, "Éxito", f"Reporte exportado exitosamente a:\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al exportar:\n{e}")


class NovedadesView(QWidget):
    """Vista principal con grilla de novedades operativas y controles de filtrado."""

    registro_modificado = Signal()

    def __init__(
        self,
        records_service: RecordsService,
        excel_service: ExcelService,
        pdf_service: PDFService,
        usuario_actual: Optional[Dict[str, Any]] = None,
        tiene_permiso_fn=None,
        parent=None,
    ):
        super().__init__(parent)
        self.records_service = records_service
        self.excel_service = excel_service
        self.pdf_service = pdf_service
        self.usuario_actual = usuario_actual or {}
        self.tiene_permiso = tiene_permiso_fn or (lambda p: True)

        self._init_ui()
        self.recargar_datos()

    def _init_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)

        # 1. Barra de Herramientas
        toolbar = QToolBar()
        toolbar.setStyleSheet("spacing: 6px;")

        self.btn_nuevo = QPushButton("➕ Nueva Novedad")
        self.btn_nuevo.setEnabled(self.tiene_permiso("novedades.crear"))
        self.btn_nuevo.clicked.connect(self._nueva_novedad)
        toolbar.addWidget(self.btn_nuevo)

        self.btn_editar = QPushButton("✏️ Modificar")
        self.btn_editar.setObjectName("btnSecondary")
        self.btn_editar.clicked.connect(self._modificar_novedad)
        toolbar.addWidget(self.btn_editar)

        self.btn_eliminar = QPushButton("🗑️ Eliminar")
        self.btn_eliminar.setObjectName("btnDanger")
        self.btn_eliminar.clicked.connect(self._eliminar_novedad)
        toolbar.addWidget(self.btn_eliminar)

        toolbar.addSeparator()

        self.btn_descubiertos = QPushButton("⚠️ Turnos Sin Cubrir")
        self.btn_descubiertos.setObjectName("btnSecondary")
        self.btn_descubiertos.clicked.connect(self._abrir_turnos_sin_cubrir)
        toolbar.addWidget(self.btn_descubiertos)

        toolbar.addSeparator()

        self.btn_export_excel = QPushButton("📊 Exportar Excel")
        self.btn_export_excel.setObjectName("btnSecondary")
        self.btn_export_excel.setEnabled(self.tiene_permiso("novedades.exportar"))
        self.btn_export_excel.clicked.connect(self._exportar_excel)
        toolbar.addWidget(self.btn_export_excel)

        self.btn_export_pdf = QPushButton("📑 Constancia PDF")
        self.btn_export_pdf.setObjectName("btnSecondary")
        self.btn_export_pdf.clicked.connect(self._exportar_pdf)
        toolbar.addWidget(self.btn_export_pdf)

        self.btn_mail = QPushButton("✉️ Enviar por Mail")
        self.btn_mail.setObjectName("btnSecondary")
        self.btn_mail.clicked.connect(self._enviar_por_mail)
        toolbar.addWidget(self.btn_mail)

        layout.addWidget(toolbar)

        # 2. Barra de Filtros
        filter_frame = QFrame()
        filter_frame.setObjectName("filterBar")
        filter_layout = QHBoxLayout(filter_frame)
        filter_layout.setContentsMargins(6, 4, 6, 4)
        filter_layout.setSpacing(8)

        # Búsqueda global
        filter_layout.addWidget(QLabel("Buscar:"))
        self.txt_search = QLineEdit()
        self.txt_search.setPlaceholderText("Texto libre...")
        self.txt_search.textChanged.connect(self._on_search_changed)
        self.txt_search.setFixedWidth(180)
        filter_layout.addWidget(self.txt_search)

        # Filtro Dotación
        filter_layout.addWidget(QLabel("Dotación:"))
        self.combo_dotacion = QComboBox()
        self.combo_dotacion.addItem("Todas")
        self.combo_dotacion.currentIndexChanged.connect(self._on_filter_changed)
        filter_layout.addWidget(self.combo_dotacion)

        # Filtro Tipo de Novedad
        filter_layout.addWidget(QLabel("Tipo:"))
        self.combo_tipo = QComboBox()
        self.combo_tipo.addItem("Todos")
        self.combo_tipo.currentIndexChanged.connect(self._on_filter_changed)
        filter_layout.addWidget(self.combo_tipo)

        # Filtro Rango de Fechas
        filter_layout.addWidget(QLabel("Desde:"))
        self.date_desde = QDateEdit()
        self.date_desde.setCalendarPopup(True)
        self.date_desde.setDisplayFormat("dd/MM/yyyy")
        self.date_desde.setDate(QDate.currentDate().addDays(-30))
        self.chk_filtrar_fecha = QComboBox()
        self.chk_filtrar_fecha.addItems(["Todas las fechas", "Filtrar por fecha"])
        self.chk_filtrar_fecha.currentIndexChanged.connect(self._on_filter_changed)
        filter_layout.addWidget(self.chk_filtrar_fecha)

        self.btn_limpiar = QPushButton("Limpiar")
        self.btn_limpiar.setObjectName("btnSecondary")
        self.btn_limpiar.clicked.connect(self._limpiar_filtros)
        filter_layout.addWidget(self.btn_limpiar)

        filter_layout.addStretch()
        layout.addWidget(filter_frame)

        # 3. Tabla de Datos
        headers = [
            "ID", "REGISTRADO EN", "LEGAJO", "APELLIDOS Y NOMBRES", "ESPECIALIDAD", "DOTACIÓN",
            "TURNO", "FRANCO", "NOVEDAD", "DESDE", "HASTA", "CUBIERTO", "REEMPLAZOS", "REF. ESTACIÓN",
            "SUPERVISOR", "USUARIO WIN",
        ]
        keys = [
            "id", "registrado_en", "legajo", "apellidos_nombres", "especialidad", "dotacion",
            "turnos", "franco", "novedad", "fecha_inicio", "fecha_fin", "cubierto", "reemplazos",
            "referencia_estacion", "supervisor", "usuario_windows",
        ]

        self.table_model = BaseTableModel(headers, keys, parent=self)
        self.proxy_model = MultiFilterProxyModel(self)
        self.proxy_model.setSourceModel(self.table_model)

        self.table_view = QTableView()
        self.table_view.setModel(self.proxy_model)
        self.table_view.setSelectionBehavior(QTableView.SelectRows)
        self.table_view.setSelectionMode(QTableView.SingleSelection)
        self.table_view.setSortingEnabled(True)
        self.table_view.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table_view.setColumnWidth(0, 55)
        self.table_view.setColumnWidth(1, 130)
        self.table_view.setColumnWidth(2, 75)
        self.table_view.setColumnWidth(3, 180)
        self.table_view.setColumnWidth(4, 110)
        self.table_view.setColumnWidth(5, 100)
        self.table_view.setColumnWidth(6, 80)
        self.table_view.setColumnWidth(7, 85)
        self.table_view.setColumnWidth(8, 140)
        self.table_view.setColumnWidth(9, 95)
        self.table_view.setColumnWidth(10, 95)
        self.table_view.setColumnWidth(11, 95)
        self.table_view.setColumnWidth(12, 160)
        self.table_view.horizontalHeader().setSectionResizeMode(13, QHeaderView.Stretch)
        self.table_view.setColumnWidth(14, 120)
        self.table_view.doubleClicked.connect(self._on_double_click)
        layout.addWidget(self.table_view)

        # Pie con conteo
        self.lbl_conteo = QLabel("Mostrando 0 registros")
        self.lbl_conteo.setStyleSheet("color: #6B7280; font-size: 8.5pt;")
        layout.addWidget(self.lbl_conteo)

    def recargar_datos(self) -> None:
        novedades = self.records_service.listar_novedades()
        coberturas = self.records_service.obtener_mapa_cobertura_novedades()

        # Enriquecer con cobertura calculada
        for n in novedades:
            info_cob = coberturas.get(n["id"])
            if info_cob:
                n["cubierto"] = info_cob.get("estado", "No")
                n["reemplazos"] = info_cob.get("detalle", "")
            else:
                n["cubierto"] = "No"
                n["reemplazos"] = ""

        self.table_model.set_data(novedades)
        self._actualizar_combos_filtros()
        self._actualizar_conteo()

    def _actualizar_combos_filtros(self) -> None:
        dotaciones = self.records_service.listar_dotaciones(incluir_inactivos=False)
        cur_dot = self.combo_dotacion.currentText()
        self.combo_dotacion.blockSignals(True)
        self.combo_dotacion.clear()
        self.combo_dotacion.addItem("Todas")
        for d in dotaciones:
            self.combo_dotacion.addItem(d["nombre"])
        idx_d = self.combo_dotacion.findText(cur_dot)
        if idx_d >= 0:
            self.combo_dotacion.setCurrentIndex(idx_d)
        self.combo_dotacion.blockSignals(False)

        tipos = self.records_service.listar_tipos(incluir_inactivos=False)
        cur_tip = self.combo_tipo.currentText()
        self.combo_tipo.blockSignals(True)
        self.combo_tipo.clear()
        self.combo_tipo.addItem("Todos")
        for t in tipos:
            self.combo_tipo.addItem(t["nombre"])
        idx_t = self.combo_tipo.findText(cur_tip)
        if idx_t >= 0:
            self.combo_tipo.setCurrentIndex(idx_t)
        self.combo_tipo.blockSignals(False)

    def _on_search_changed(self, text: str) -> None:
        self.proxy_model.set_global_search(text)
        self._actualizar_conteo()

    def _on_filter_changed(self) -> None:
        # Columna 5 = Dotación, Columna 8 = Novedad
        dot = self.combo_dotacion.currentText()
        self.proxy_model.set_column_filter(5, dot)

        tipo = self.combo_tipo.currentText()
        self.proxy_model.set_column_filter(8, tipo)

        self._actualizar_conteo()

    def _limpiar_filtros(self) -> None:
        self.txt_search.clear()
        self.combo_dotacion.setCurrentIndex(0)
        self.combo_tipo.setCurrentIndex(0)
        self.chk_filtrar_fecha.setCurrentIndex(0)
        self.proxy_model.clear_all_filters()
        self._actualizar_conteo()

    def _actualizar_conteo(self) -> None:
        visibles = self.proxy_model.rowCount()
        totales = self.table_model.rowCount()
        self.lbl_conteo.setText(f"Mostrando {visibles} de {totales} novedades registradas")

    def _obtener_fila_seleccionada(self) -> Optional[Dict[str, Any]]:
        indexes = self.table_view.selectionModel().selectedRows()
        if not indexes:
            return None
        proxy_idx = indexes[0]
        source_idx = self.proxy_model.mapToSource(proxy_idx)
        return self.table_model.get_row(source_idx.row())

    def _nueva_novedad(self) -> None:
        dlg = NovedadDialog(
            self.records_service,
            usuario_actual=self.usuario_actual,
            parent=self,
        )
        if dlg.exec():
            self.recargar_datos()
            self.registro_modificado.emit()

    def _modificar_novedad(self) -> None:
        fila = self._obtener_fila_seleccionada()
        if not fila:
            QMessageBox.information(self, "Modificar", "Seleccione una novedad en la tabla.")
            return

        nov_id = fila["id"]
        # Chequeo de permisos y ventana de edición
        puede_modificar_libre = self.tiene_permiso("registros.modificar")
        dentro_ventana = self.records_service.dentro_de_ventana(fila["registrado_en"], "editar_horas")
        falta_fecha_fin = not fila.get("fecha_fin") or fila.get("fecha_fin") == "-"

        if not puede_modificar_libre and not dentro_ventana:
            if falta_fecha_fin:
                modo_restringido = True
            else:
                QMessageBox.warning(
                    self,
                    "Tiempo Expirado",
                    "El registro superó el tiempo límite para editar y ya cuenta con fecha de fin.",
                )
                return
        else:
            modo_restringido = False

        tramos = self.records_service.listar_reemplazos_por_novedad(nov_id)
        dlg = NovedadDialog(
            self.records_service,
            novedad_existente=fila,
            tramos_existentes=tramos,
            modo_restringido=modo_restringido,
            usuario_actual=self.usuario_actual,
            parent=self,
        )
        if dlg.exec():
            self.recargar_datos()
            self.registro_modificado.emit()

    def _eliminar_novedad(self) -> None:
        fila = self._obtener_fila_seleccionada()
        if not fila:
            QMessageBox.information(self, "Eliminar", "Seleccione una novedad en la tabla.")
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
            f"¿Está seguro de que desea eliminar la novedad #{fila['id']} de {fila['apellidos_nombres']}?\n"
            f"(El registro quedará en la papelera y podrá ser recuperado por un administrador).",
            QMessageBox.Yes | QMessageBox.No,
        )
        if resp != QMessageBox.Yes:
            return

        try:
            self.records_service.eliminar_novedad(
                fila["id"],
                user_id=self.usuario_actual.get("id"),
                windows_user=obtener_usuario_windows(),
                bypass_window=puede_eliminar_libre,
            )
            self.recargar_datos()
            self.registro_modificado.emit()
            QMessageBox.information(self, "Eliminado", "La novedad fue enviada a registros eliminados.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo eliminar la novedad:\n{e}")

    def _abrir_turnos_sin_cubrir(self) -> None:
        dots = [d["nombre"] for d in self.records_service.listar_dotaciones(incluir_inactivos=False)]
        dlg = TurnosSinCubrirDialog(self.records_service, self.excel_service, dots, parent=self)
        dlg.exec()

    def _exportar_excel(self) -> None:
        default_name = f"Novedades_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        file_path, _ = QFileDialog.getSaveFileName(self, "Exportar Novedades a Excel", default_name, "Libro de Excel (*.xlsx)")
        if not file_path:
            return
        try:
            self.excel_service.export_database(file_path, tables=("NOVEDADES",))
            QMessageBox.information(self, "Exportación", f"Novedades exportadas exitosamente a:\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo exportar a Excel:\n{e}")

    def _exportar_pdf(self) -> None:
        fila = self._obtener_fila_seleccionada()
        if not fila:
            QMessageBox.information(self, "Exportar PDF", "Seleccione una novedad en la tabla para generar su constancia.")
            return

        default_name = f"Constancia_Novedad_{fila['id']}_{datetime.now().strftime('%Y%m%d')}.pdf"
        file_path, _ = QFileDialog.getSaveFileName(self, "Guardar Constancia PDF", default_name, "Archivo PDF (*.pdf)")
        if not file_path:
            return

        try:
            tramos = self.records_service.listar_reemplazos_por_novedad(fila["id"])
            self.pdf_service.generar_constancia_novedad(fila, tramos, file_path)
            QMessageBox.information(self, "Constancia PDF", f"Constancia PDF generada exitosamente en:\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo generar la constancia PDF:\n{e}")

    def _enviar_por_mail(self) -> None:
        fila = self._obtener_fila_seleccionada()
        if not fila:
            QMessageBox.information(self, "Enviar por Mail", "Seleccione una novedad en la tabla.")
            return
        try:
            import os
            import tempfile
            from src.services.mail_service import MailService

            tmp_dir = tempfile.gettempdir()
            pdf_path = os.path.join(tmp_dir, f"Constancia_Novedad_{fila['id']}.pdf")
            tramos = self.records_service.listar_reemplazos_por_novedad(fila["id"])
            self.pdf_service.generar_constancia_novedad(fila, tramos, pdf_path)

            dests = self.records_service.listar_destinatarios_informe(incluir_inactivos=False)
            emails = [d["email"] for d in dests if d.get("email")]
            asunto = f"Informe de Novedad #{fila['id']} - {fila.get('apellidos_nombres')} (Leg. {fila.get('legajo')})"
            cuerpo = (
                f"<p>Se adjunta constancia de novedad operativa <b>#{fila['id']}</b> ({fila.get('novedad')}).</p>"
                f"<ul>"
                f"<li><b>Agente:</b> {fila.get('apellidos_nombres')} (Leg. {fila.get('legajo')})</li>"
                f"<li><b>Dotación:</b> {fila.get('dotacion')} | <b>Especialidad:</b> {fila.get('especialidad')}</li>"
                f"<li><b>Desde:</b> {fila.get('fecha_inicio')} | <b>Hasta:</b> {fila.get('fecha_fin') or '-'}</li>"
                f"<li><b>Ref. Estación:</b> {fila.get('referencia_estacion')} | <b>Supervisor:</b> {fila.get('supervisor')}</li>"
                f"<li><b>Observaciones:</b> {fila.get('observaciones') or '-'}</li>"
                f"</ul>"
            )
            ok, msg = MailService.crear_borrador_outlook(
                destinatarios=emails,
                asunto=asunto,
                cuerpo_html=cuerpo,
                adjuntos=[pdf_path],
            )
            if not ok:
                QMessageBox.warning(self, "Correo", msg)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo preparar el correo:\n{e}")

    def _on_double_click(self, index) -> None:
        self._modificar_novedad()
