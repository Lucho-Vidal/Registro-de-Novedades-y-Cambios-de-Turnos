"""Vista principal de Ordenamientos Operativos con grilla interactiva, filtros, PDF duplicado, Excel y Mail."""

from __future__ import annotations

import os
import tempfile
from datetime import datetime
from typing import Any, Dict, List, Optional

from PySide6.QtCore import QDate, Qt, Signal
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDateEdit,
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
from src.services.excel_service import ExcelService
from src.services.mail_service import MailService
from src.services.pdf_service import PDFService
from src.services.records_service import RecordsService
from src.ui.models.base_table_model import BaseTableModel
from src.ui.models.multi_filter_proxy import MultiFilterProxyModel
from src.ui.views.dialogs.admin_tren_paradas_dialog import AdminTrenParadasDialog
from src.ui.views.dialogs.ordenamiento_dialog import OrdenamientoDialog


class OrdenamientosView(QWidget):
    """Vista de gestión de Ordenamientos Operativos dentro del módulo Novedades y Cambios."""

    registro_modificado = Signal()

    def __init__(
        self,
        records_service: RecordsService,
        excel_service: ExcelService,
        pdf_service: PDFService,
        usuario_actual: Optional[Dict[str, Any]] = None,
        tiene_permiso_fn=None,
        parent: Optional[QWidget] = None,
    ) -> None:
        super().__init__(parent)
        self.records_service = records_service
        self.excel_service = excel_service
        self.pdf_service = pdf_service
        self.mail_service = MailService()
        self.usuario_actual = usuario_actual or {}
        self.tiene_permiso = tiene_permiso_fn or (lambda p: True)

        self._init_ui()
        self.recargar_datos()

    def _init_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)

        # 1. Barra de Acciones
        toolbar = QToolBar()
        toolbar.setStyleSheet("spacing: 6px;")

        self.btn_nuevo = QPushButton("➕ Nuevo Ordenamiento")
        self.btn_nuevo.setEnabled(self.tiene_permiso("ordenamientos.crear") or self.tiene_permiso("novedades.crear"))
        self.btn_nuevo.clicked.connect(self._nuevo_ordenamiento)
        toolbar.addWidget(self.btn_nuevo)

        self.btn_editar = QPushButton("✏️ Modificar")
        self.btn_editar.setObjectName("btnSecondary")
        self.btn_editar.clicked.connect(self._modificar_ordenamiento)
        toolbar.addWidget(self.btn_editar)

        self.btn_eliminar = QPushButton("🗑️ Eliminar")
        self.btn_eliminar.setObjectName("btnDanger")
        self.btn_eliminar.clicked.connect(self._eliminar_ordenamiento)
        toolbar.addWidget(self.btn_eliminar)

        toolbar.addSeparator()

        self.btn_pdf = QPushButton("🖨️ Imprimir / PDF Duplicado")
        self.btn_pdf.setObjectName("btnSecondary")
        self.btn_pdf.clicked.connect(self._exportar_pdf)
        toolbar.addWidget(self.btn_pdf)

        self.btn_mail = QPushButton("✉️ Enviar por Mail")
        self.btn_mail.setObjectName("btnSecondary")
        self.btn_mail.clicked.connect(self._enviar_por_mail)
        toolbar.addWidget(self.btn_mail)

        self.btn_excel = QPushButton("📊 Exportar Excel")
        self.btn_excel.setObjectName("btnSecondary")
        self.btn_excel.clicked.connect(self._exportar_excel)
        toolbar.addWidget(self.btn_excel)

        toolbar.addSeparator()

        self.btn_paradas = QPushButton("🚉 Paradas por Tren (CRUD / Excel)")
        self.btn_paradas.setObjectName("btnSecondary")
        self.btn_paradas.clicked.connect(self._abrir_crud_paradas)
        toolbar.addWidget(self.btn_paradas)

        layout.addWidget(toolbar)

        # 2. Barra de Filtros
        filter_frame = QFrame()
        filter_frame.setObjectName("filterBar")
        filter_lay = QHBoxLayout(filter_frame)
        filter_lay.setContentsMargins(6, 4, 6, 4)
        filter_lay.setSpacing(8)

        filter_lay.addWidget(QLabel("Buscar:"))
        self.txt_search = QLineEdit()
        self.txt_search.setPlaceholderText("N° orden, legajo, nombre, tren, turno, motivo...")
        self.txt_search.setFixedWidth(260)
        self.txt_search.textChanged.connect(self.recargar_datos)
        filter_lay.addWidget(self.txt_search)

        filter_lay.addWidget(QLabel("Dotación:"))
        self.combo_dotacion = QComboBox()
        self.combo_dotacion.addItem("Todas")
        self.combo_dotacion.currentIndexChanged.connect(self.recargar_datos)
        filter_lay.addWidget(self.combo_dotacion)

        self.chk_fecha = QCheckBox("Fecha:")
        self.chk_fecha.toggled.connect(self._on_chk_fecha_toggled)
        filter_lay.addWidget(self.chk_fecha)

        self.date_filtro = QDateEdit()
        self.date_filtro.setCalendarPopup(True)
        self.date_filtro.setDisplayFormat("dd/MM/yyyy")
        self.date_filtro.setDate(QDate.currentDate())
        self.date_filtro.setEnabled(False)
        self.date_filtro.dateChanged.connect(self.recargar_datos)
        filter_lay.addWidget(self.date_filtro)

        filter_lay.addWidget(QLabel("Orden:"))
        self.combo_orden = QComboBox()
        self.combo_orden.addItems(["Ascendente (Por defecto)", "Descendente (Más recientes)"])
        self.combo_orden.currentIndexChanged.connect(self.recargar_datos)
        filter_lay.addWidget(self.combo_orden)

        self.btn_limpiar = QPushButton("Limpiar")
        self.btn_limpiar.setObjectName("btnSecondary")
        self.btn_limpiar.clicked.connect(self._limpiar_filtros)
        filter_lay.addWidget(self.btn_limpiar)

        filter_lay.addStretch()
        layout.addWidget(filter_frame)

        # 3. Tabla Principal
        headers = [
            "ID", "N° ORDEN", "FECHA", "HORA", "LEGAJO", "APELLIDO Y NOMBRE",
            "DOTACIÓN", "ITINERARIO", "TURNO / CICLO", "RÉGIMEN", "TRENES",
            "MOTIVO DE LA ORDEN", "REF ESTACIÓN", "EMISOR", "USUARIO WIN",
        ]
        keys = [
            "id", "nro_orden", "fecha_emision", "hora_emision", "legajo_personal", "nombre_personal",
            "deposito", "itinerario", "ciclo_o_turno", "regimen", "trenes_involucrados",
            "motivo_orden", "ref_pct", "emisor", "usuario_windows",
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
        self.table_view.setColumnWidth(0, 50)
        self.table_view.setColumnWidth(1, 115)
        self.table_view.setColumnWidth(2, 85)
        self.table_view.setColumnWidth(3, 60)
        self.table_view.setColumnWidth(4, 70)
        self.table_view.setColumnWidth(5, 175)
        self.table_view.setColumnWidth(6, 75)
        self.table_view.setColumnWidth(7, 75)
        self.table_view.setColumnWidth(8, 95)
        self.table_view.setColumnWidth(9, 65)
        self.table_view.setColumnWidth(10, 120)
        self.table_view.horizontalHeader().setSectionResizeMode(11, QHeaderView.Stretch)
        self.table_view.setColumnWidth(12, 130)
        self.table_view.setColumnWidth(13, 95)
        self.table_view.setColumnWidth(14, 95)
        self.table_view.doubleClicked.connect(lambda _: self._modificar_ordenamiento())
        layout.addWidget(self.table_view)

        self.lbl_conteo = QLabel("Mostrando 0 ordenamientos")
        self.lbl_conteo.setStyleSheet("color: #6B7280; font-size: 8.5pt;")
        layout.addWidget(self.lbl_conteo)

    def _on_chk_fecha_toggled(self, checked: bool) -> None:
        self.date_filtro.setEnabled(checked)
        self.recargar_datos()

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

    def recargar_datos(self) -> None:
        self._actualizar_combo_dotaciones()
        filtros: Dict[str, Any] = {}
        txt = self.txt_search.text().strip()
        if txt:
            filtros["texto"] = txt
        dot = self.combo_dotacion.currentText().strip()
        if dot and dot != "Todas":
            filtros["deposito"] = dot
        if self.chk_fecha.isChecked():
            filtros["fecha"] = self.date_filtro.date().toString("dd/MM/yyyy")

        orden_asc = self.combo_orden.currentIndex() == 0
        datos = self.records_service.listar_ordenamientos(
            filtros=filtros, incluir_inactivos=False, orden_asc=orden_asc
        )
        self.table_model.set_data(datos)
        self.lbl_conteo.setText(f"Mostrando {len(datos)} ordenamientos operativos")

    def _limpiar_filtros(self) -> None:
        self.txt_search.blockSignals(True)
        self.txt_search.clear()
        self.txt_search.blockSignals(False)
        self.combo_dotacion.setCurrentIndex(0)
        self.chk_fecha.setChecked(False)
        self.combo_orden.setCurrentIndex(0)
        self.recargar_datos()

    def _obtener_seleccionado(self) -> Optional[Dict[str, Any]]:
        indexes = self.table_view.selectionModel().selectedRows()
        if not indexes:
            return None
        source_idx = self.proxy_model.mapToSource(indexes[0])
        return self.table_model.get_row(source_idx.row())

    def _nuevo_ordenamiento(self) -> None:
        dlg = OrdenamientoDialog(
            self.records_service,
            pdf_service=self.pdf_service,
            usuario_actual=self.usuario_actual,
            parent=self,
        )
        if dlg.exec():
            self.recargar_datos()
            self.registro_modificado.emit()

    def _modificar_ordenamiento(self) -> None:
        fila = self._obtener_seleccionado()
        if not fila:
            QMessageBox.information(self, "Modificar", "Seleccione un ordenamiento en la tabla.")
            return

        completo = self.records_service.obtener_ordenamiento_completo(int(fila["id"]))
        if not completo:
            return

        puede_modificar_libre = self.tiene_permiso("registros.modificar")
        ref_time = completo.get("created_at") or f"{completo.get('fecha_emision')} {completo.get('hora_emision')}"
        dentro_ventana = self.records_service.dentro_de_ventana(ref_time, "editar_horas")

        if not puede_modificar_libre and not dentro_ventana:
            QMessageBox.warning(
                self,
                "Tiempo Expirado",
                "El ordenamiento superó el límite de tiempo permitido para ser modificado.",
            )
            return

        dlg = OrdenamientoDialog(
            self.records_service,
            pdf_service=self.pdf_service,
            ordenamiento_existente=completo,
            usuario_actual=self.usuario_actual,
            parent=self,
        )
        if dlg.exec():
            self.recargar_datos()
            self.registro_modificado.emit()

    def _eliminar_ordenamiento(self) -> None:
        fila = self._obtener_seleccionado()
        if not fila:
            QMessageBox.information(self, "Eliminar", "Seleccione un ordenamiento en la tabla.")
            return

        puede_eliminar_libre = self.tiene_permiso("registros.modificar")
        ref_time = fila.get("created_at") or f"{fila.get('fecha_emision')} {fila.get('hora_emision')}"
        dentro_ventana = self.records_service.dentro_de_ventana(ref_time, "eliminar_horas")

        if not puede_eliminar_libre and not dentro_ventana:
            QMessageBox.warning(
                self,
                "Tiempo Expirado",
                "El ordenamiento superó el límite de tiempo permitido para ser eliminado.",
            )
            return

        resp = QMessageBox.question(
            self,
            "Confirmar Eliminación",
            f"¿Está seguro que desea enviar a la papelera el Ordenamiento {fila['nro_orden']} "
            f"de {fila['nombre_personal']}?",
            QMessageBox.Yes | QMessageBox.No,
        )
        if resp != QMessageBox.Yes:
            return

        try:
            self.records_service.eliminar_ordenamiento(
                int(fila["id"]),
                user_id=self.usuario_actual.get("id"),
                windows_user=obtener_usuario_windows(),
                bypass_window=puede_eliminar_libre,
            )
            self.recargar_datos()
            self.registro_modificado.emit()
            QMessageBox.information(self, "Eliminado", "El ordenamiento fue enviado a la papelera.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo eliminar el ordenamiento:\n{e}")

    def _exportar_pdf(self) -> None:
        fila = self._obtener_seleccionado()
        if not fila:
            QMessageBox.information(self, "Imprimir / PDF", "Seleccione un ordenamiento en la tabla.")
            return
        completo = self.records_service.obtener_ordenamiento_completo(int(fila["id"]))
        if not completo:
            return

        nro_safe = str(completo.get("nro_orden") or completo["id"]).replace(" ", "_")
        default_name = f"Orden_Servicio_{nro_safe}_{datetime.now().strftime('%Y%m%d')}.pdf"
        out_path, _ = QFileDialog.getSaveFileName(
            self,
            "Guardar Orden de Servicio por Duplicado (PDF)",
            default_name,
            "Archivo PDF (*.pdf)",
        )
        if not out_path:
            return
        try:
            self.pdf_service.generar_orden_servicio_duplicado(
                completo, completo.get("tramos") or [], out_path
            )
            QMessageBox.information(self, "PDF Generado", f"Orden de servicio por duplicado generada en:\n{out_path}")
            try:
                os.startfile(out_path)
            except Exception:
                pass
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo generar el PDF:\n{e}")

    def _enviar_por_mail(self) -> None:
        fila = self._obtener_seleccionado()
        if not fila:
            QMessageBox.information(self, "Enviar por Mail", "Seleccione un ordenamiento en la tabla.")
            return
        completo = self.records_service.obtener_ordenamiento_completo(int(fila["id"]))
        if not completo:
            return

        try:
            nro_safe = str(completo.get("nro_orden") or completo["id"]).replace(" ", "_")
            tmp_dir = tempfile.gettempdir()
            pdf_path = os.path.join(tmp_dir, f"Orden_Servicio_{nro_safe}.pdf")
            self.pdf_service.generar_orden_servicio_duplicado(
                completo, completo.get("tramos") or [], pdf_path
            )

            dests = self.records_service.listar_destinatarios_informe(incluir_inactivos=False)
            emails = [d["email"] for d in dests if d.get("email")]
            asunto = (
                f"Orden de Servicio {completo.get('nro_orden')} - "
                f"{completo.get('nombre_personal')} (Leg. {completo.get('legajo_personal')})"
            )
            cuerpo = (
                f"<p>Se adjunta Orden de Servicio N° <b>{completo.get('nro_orden')}</b> emitida el "
                f"{completo.get('fecha_emision')} a las {completo.get('hora_emision')} hs.</p>"
                f"<ul>"
                f"<li><b>Agente:</b> {completo.get('nombre_personal')} (Leg. {completo.get('legajo_personal')})</li>"
                f"<li><b>Dotación:</b> {completo.get('deposito')} | <b>Turno/Ciclo:</b> {completo.get('ciclo_o_turno')}</li>"
                f"<li><b>Motivo:</b> {completo.get('motivo_orden')}</li>"
                f"<li><b>REF de {completo.get('deposito')}:</b> {completo.get('ref_pct')} | <b>Emisor:</b> {completo.get('emisor')}</li>"
                f"</ul>"
            )
            ok, msg = self.mail_service.crear_borrador_outlook(
                destinatarios=emails,
                asunto=asunto,
                cuerpo_html=cuerpo,
                adjuntos=[pdf_path],
            )
            if not ok:
                QMessageBox.warning(self, "Correo", msg)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo preparar el correo:\n{e}")

    def _exportar_excel(self) -> None:
        datos = [self.table_model.get_row(r) for r in range(self.table_model.rowCount())]
        if not datos:
            QMessageBox.information(self, "Exportar Excel", "No hay ordenamientos para exportar.")
            return
        default_name = f"Ordenamientos_Operativos_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        out_path, _ = QFileDialog.getSaveFileName(
            self, "Exportar Ordenamientos a Excel", default_name, "Libro de Excel (*.xlsx)"
        )
        if not out_path:
            return
        try:
            self.excel_service.export_ordenamientos(datos, out_path)
            QMessageBox.information(self, "Exportación", f"Ordenamientos exportados exitosamente a:\n{out_path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo exportar a Excel:\n{e}")

    def _abrir_crud_paradas(self) -> None:
        dlg = AdminTrenParadasDialog(
            self.records_service,
            usuario_actual=self.usuario_actual,
            parent=self,
        )
        dlg.exec()
