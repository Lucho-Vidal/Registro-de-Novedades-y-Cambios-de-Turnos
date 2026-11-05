"""Diálogo modal maximizado para el alta y edición de Ordenamientos Operativos."""

from __future__ import annotations

from datetime import datetime
from typing import Any, Dict, List, Optional

from PySide6.QtCore import QDate, QTime, Qt
from PySide6.QtWidgets import (
    QComboBox,
    QDateEdit,
    QDialog,
    QFileDialog,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QTimeEdit,
    QVBoxLayout,
    QWidget,
)

from src.core.config import DOTACIONES, obtener_usuario_windows
from src.domain.validators import parsear_fecha
from src.services.pdf_service import PDFService
from src.services.records_service import RecordsService
from src.ui.views.dialogs.admin_tren_paradas_dialog import AdminTrenParadasDialog
from src.ui.views.dialogs.legajo_search_dialog import LegajoSearchDialog


class OrdenamientoDialog(QDialog):
    """Formulario modal maximizado para crear o editar un Ordenamiento Operativo (Orden de Servicio)."""

    def __init__(
        self,
        records_service: RecordsService,
        pdf_service: Optional[PDFService] = None,
        ordenamiento_existente: Optional[Dict[str, Any]] = None,
        usuario_actual: Optional[Dict[str, Any]] = None,
        parent: Optional[QWidget] = None,
    ) -> None:
        super().__init__(parent)
        self.records_service = records_service
        self.pdf_service = pdf_service or PDFService()
        self.ordenamiento_existente = ordenamiento_existente
        self.usuario_actual = usuario_actual or {}
        self.tramos: List[Dict[str, Any]] = list(
            (ordenamiento_existente or {}).get("tramos") or []
        )
        self.saved_ordenamiento_id: Optional[int] = (
            ordenamiento_existente.get("id") if ordenamiento_existente else None
        )

        self.setWindowTitle(
            f"Modificar Ordenamiento Operativo - {ordenamiento_existente.get('nro_orden', '')}"
            if ordenamiento_existente
            else "Nuevo Ordenamiento Operativo (Orden de Servicio)"
        )
        self.resize(1120, 740)
        self.setWindowFlags(
            (self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
            | Qt.WindowMinMaxButtonsHint
            | Qt.WindowCloseButtonHint
        )

        self._empleados_cache = self.records_service.listar_empleados(incluir_inactivos=False)
        self._personal_estacion_cache = self.records_service.listar_personal_estacion(incluir_inactivos=False)
        self._paradas_tren_actual: Dict[str, str] = {}

        self._init_ui()
        if self.ordenamiento_existente:
            self._cargar_datos_existentes()
        else:
            self._actualizar_nro_orden_y_refs("PC")

    def showEvent(self, event) -> None:
        super().showEvent(event)
        self.showMaximized()

    def _init_ui(self) -> None:
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(10)

        # 1. Cabecera del Ordenamiento y Datos del Personal
        grp_cabecera = QGroupBox("1. Datos de Cabecera y Personal Asignado")
        lay_cab = QGridLayout(grp_cabecera)

        lay_cab.addWidget(QLabel("Dotación / Depósito (*):"), 0, 0)
        self.combo_deposito = QComboBox()
        self.combo_deposito.setEditable(True)
        dotaciones_db = [d["nombre"] for d in self.records_service.listar_dotaciones(incluir_inactivos=False)]
        if not dotaciones_db:
            dotaciones_db = [d for d in DOTACIONES if d != "Todas"]
        for d_nom in dotaciones_db:
            if self.combo_deposito.findText(d_nom) == -1:
                self.combo_deposito.addItem(d_nom)
        self.combo_deposito.currentTextChanged.connect(self._on_deposito_changed)
        lay_cab.addWidget(self.combo_deposito, 0, 1)

        lay_cab.addWidget(QLabel("N° de Orden (*):"), 0, 2)
        self.txt_nro_orden = QLineEdit()
        self.txt_nro_orden.setStyleSheet("font-weight: bold;")
        lay_cab.addWidget(self.txt_nro_orden, 0, 3)

        lay_cab.addWidget(QLabel("Fecha Emisión:"), 0, 4)
        self.date_emision = QDateEdit()
        self.date_emision.setCalendarPopup(True)
        self.date_emision.setDisplayFormat("dd/MM/yyyy")
        self.date_emision.setDate(QDate.currentDate())
        lay_cab.addWidget(self.date_emision, 0, 5)

        lay_cab.addWidget(QLabel("Hora Emisión:"), 0, 6)
        self.time_emision = QTimeEdit()
        self.time_emision.setDisplayFormat("HH:mm")
        self.time_emision.setTime(QTime.currentTime())
        lay_cab.addWidget(self.time_emision, 0, 7)

        # Fila 1: Legajo, Nombre, Turno/Ciclo, Régimen
        lay_cab.addWidget(QLabel("Legajo (*):"), 1, 0)
        h_leg = QHBoxLayout()
        self.txt_legajo = QLineEdit()
        self.txt_legajo.setMaximumWidth(100)
        self.txt_legajo.editingFinished.connect(self._on_legajo_changed)
        self.btn_buscar_legajo = QPushButton("🔍 Buscar")
        self.btn_buscar_legajo.setObjectName("btnSecondary")
        self.btn_buscar_legajo.clicked.connect(self._buscar_legajo_modal)
        h_leg.addWidget(self.txt_legajo)
        h_leg.addWidget(self.btn_buscar_legajo)
        lay_cab.addLayout(h_leg, 1, 1)

        lay_cab.addWidget(QLabel("Apellido y Nombre (*):"), 1, 2)
        self.txt_nombre = QLineEdit()
        lay_cab.addWidget(self.txt_nombre, 1, 3, 1, 3)

        lay_cab.addWidget(QLabel("Itinerario:"), 1, 6)
        self.combo_itinerario = QComboBox()
        self.combo_itinerario.setEditable(True)
        for it in self.records_service.listar_itinerarios_disponibles():
            self.combo_itinerario.addItem(it)
        if self.combo_itinerario.count() == 0:
            self.combo_itinerario.addItem("H45")
        lay_cab.addWidget(self.combo_itinerario, 1, 7)

        # Fila 2: Ciclo/Turno, Régimen, REF de Dotación, Emisor, Usuario Windows
        lay_cab.addWidget(QLabel("Ciclo o Turno:"), 2, 0)
        h_turno = QHBoxLayout()
        self.combo_turno = QComboBox()
        self.combo_turno.setEditable(True)
        self.combo_turno.addItem("")
        for tb in self.records_service.listar_codigos_turnos_base():
            self.combo_turno.addItem(tb)
        self.btn_cargar_vueltas = QPushButton("📋 Cargar Turno")
        self.btn_cargar_vueltas.setObjectName("btnSecondary")
        self.btn_cargar_vueltas.setToolTip("Carga todas las vueltas del turno seleccionado a la grilla de tramos")
        self.btn_cargar_vueltas.clicked.connect(self._cargar_vueltas_de_turno)
        h_turno.addWidget(self.combo_turno, 1)
        h_turno.addWidget(self.btn_cargar_vueltas)
        lay_cab.addLayout(h_turno, 2, 1, 1, 2)

        lay_cab.addWidget(QLabel("Régimen:"), 2, 3)
        self.combo_regimen = QComboBox()
        self.combo_regimen.addItems(["LV", "S", "D"])
        lay_cab.addWidget(self.combo_regimen, 2, 4)

        self.lbl_ref_dotacion = QLabel("REF de PC (*):")
        lay_cab.addWidget(self.lbl_ref_dotacion, 2, 5)
        self.combo_ref_pct = QComboBox()
        self.combo_ref_pct.setEditable(True)
        self.combo_ref_pct.currentTextChanged.connect(self._on_ref_estacion_changed)
        lay_cab.addWidget(self.combo_ref_pct, 2, 6, 1, 2)

        # Fila 3: Motivo de la orden, Emisor y Usuario Windows
        lay_cab.addWidget(QLabel("Motivo / Por orden de (*):"), 3, 0)
        self.txt_motivo = QLineEdit()
        self.txt_motivo.setPlaceholderText("Ej: A las órdenes de Control Trenes / Diagrama especial / Relevo...")
        lay_cab.addWidget(self.txt_motivo, 3, 1, 1, 3)

        lay_cab.addWidget(QLabel("Emisor (RENO):"), 3, 4)
        self.txt_emisor = QLineEdit()
        self.txt_emisor.setReadOnly(True)
        self.txt_emisor.setEnabled(False)
        emisor_def = self.usuario_actual.get("username") or self.usuario_actual.get("nombre") or "Sistema"
        self.txt_emisor.setText(str(emisor_def))
        lay_cab.addWidget(self.txt_emisor, 3, 5)

        lay_cab.addWidget(QLabel("Usuario Windows:"), 3, 6)
        self.txt_usuario_windows = QLineEdit()
        self.txt_usuario_windows.setReadOnly(True)
        self.txt_usuario_windows.setEnabled(False)
        self.txt_usuario_windows.setText(obtener_usuario_windows())
        lay_cab.addWidget(self.txt_usuario_windows, 3, 7)

        main_layout.addWidget(grp_cabecera)

        # 2. Constructor Dinámico de Tramos y Tabla de Diagrama
        grp_tramos = QGroupBox("2. Diagrama de Trenes y Tramos Operativos")
        lay_tramos = QVBoxLayout(grp_tramos)

        # Fila de carga rápida de tramo con autocompletado por N° de tren
        builder_grid = QGridLayout()
        builder_grid.addWidget(QLabel("Ref:"), 0, 0)
        self.txt_tramo_ref = QLineEdit()
        self.txt_tramo_ref.setPlaceholderText("Ref")
        self.txt_tramo_ref.setMaximumWidth(80)
        builder_grid.addWidget(self.txt_tramo_ref, 1, 0)

        builder_grid.addWidget(QLabel("N° de Tren:"), 0, 1)
        self.txt_tramo_tren = QLineEdit()
        self.txt_tramo_tren.setPlaceholderText("Ej: 1293")
        self.txt_tramo_tren.setMaximumWidth(100)
        self.txt_tramo_tren.textChanged.connect(self._on_tramo_tren_changed)
        builder_grid.addWidget(self.txt_tramo_tren, 1, 1)

        builder_grid.addWidget(QLabel("Sale de (Origen):"), 0, 2)
        self.combo_tramo_origen = QComboBox()
        self.combo_tramo_origen.setEditable(True)
        self.combo_tramo_origen.setMinimumWidth(150)
        self.combo_tramo_origen.currentTextChanged.connect(self._on_origen_selected)
        builder_grid.addWidget(self.combo_tramo_origen, 1, 2)

        builder_grid.addWidget(QLabel("Hora Salida:"), 0, 3)
        self.txt_tramo_h_sal = QLineEdit()
        self.txt_tramo_h_sal.setPlaceholderText("HH:MM")
        self.txt_tramo_h_sal.setMaximumWidth(80)
        builder_grid.addWidget(self.txt_tramo_h_sal, 1, 3)

        builder_grid.addWidget(QLabel("Llega a (Destino):"), 0, 4)
        self.combo_tramo_destino = QComboBox()
        self.combo_tramo_destino.setEditable(True)
        self.combo_tramo_destino.setMinimumWidth(150)
        self.combo_tramo_destino.currentTextChanged.connect(self._on_destino_selected)
        builder_grid.addWidget(self.combo_tramo_destino, 1, 4)

        builder_grid.addWidget(QLabel("Hora Llegada:"), 0, 5)
        self.txt_tramo_h_lleg = QLineEdit()
        self.txt_tramo_h_lleg.setPlaceholderText("HH:MM")
        self.txt_tramo_h_lleg.setMaximumWidth(80)
        builder_grid.addWidget(self.txt_tramo_h_lleg, 1, 5)

        builder_grid.addWidget(QLabel("Observaciones del Tramo:"), 0, 6)
        self.txt_tramo_obs = QLineEdit()
        self.txt_tramo_obs.setPlaceholderText("Ej: Condicional, Liviano, A las órdenes...")
        builder_grid.addWidget(self.txt_tramo_obs, 1, 6)

        lay_tramos.addLayout(builder_grid)

        # Botonera de manipulación de tramos
        bar_tramos = QHBoxLayout()
        self.btn_add_tramo = QPushButton("➕ Agregar Tramo a la Tabla")
        self.btn_add_tramo.clicked.connect(self._agregar_tramo_tabla)
        bar_tramos.addWidget(self.btn_add_tramo)

        self.btn_upd_tramo = QPushButton("✏️ Actualizar Fila Seleccionada")
        self.btn_upd_tramo.setObjectName("btnSecondary")
        self.btn_upd_tramo.clicked.connect(self._actualizar_tramo_seleccionado)
        bar_tramos.addWidget(self.btn_upd_tramo)

        self.btn_up_tramo = QPushButton("⬆️ Subir")
        self.btn_up_tramo.setObjectName("btnSecondary")
        self.btn_up_tramo.clicked.connect(lambda: self._mover_tramo(-1))
        bar_tramos.addWidget(self.btn_up_tramo)

        self.btn_down_tramo = QPushButton("⬇️ Bajar")
        self.btn_down_tramo.setObjectName("btnSecondary")
        self.btn_down_tramo.clicked.connect(lambda: self._mover_tramo(1))
        bar_tramos.addWidget(self.btn_down_tramo)

        self.btn_del_tramo = QPushButton("🗑️ Quitar Fila")
        self.btn_del_tramo.setObjectName("btnDanger")
        self.btn_del_tramo.clicked.connect(self._eliminar_tramo_seleccionado)
        bar_tramos.addWidget(self.btn_del_tramo)

        bar_tramos.addStretch()

        self.btn_admin_paradas = QPushButton("🚉 Paradas por Tren (CRUD / Importar Excel)")
        self.btn_admin_paradas.setObjectName("btnSecondary")
        self.btn_admin_paradas.clicked.connect(self._abrir_crud_paradas)
        bar_tramos.addWidget(self.btn_admin_paradas)

        lay_tramos.addLayout(bar_tramos)

        # Tabla de Tramos
        self.tabla_tramos = QTableWidget(0, 8)
        self.tabla_tramos.setHorizontalHeaderLabels([
            "#", "REF", "N° TREN", "SALE DE (ORIGEN)", "HORA SALIDA",
            "LLEGA A (DESTINO)", "HORA LLEGADA", "OBSERVACIONES DEL TRAMO"
        ])
        self.tabla_tramos.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.tabla_tramos.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeToContents)
        self.tabla_tramos.horizontalHeader().setSectionResizeMode(7, QHeaderView.Stretch)
        self.tabla_tramos.setSelectionBehavior(QTableWidget.SelectRows)
        self.tabla_tramos.setSelectionMode(QTableWidget.SingleSelection)
        self.tabla_tramos.itemSelectionChanged.connect(self._on_seleccion_tramo_tabla)
        lay_tramos.addWidget(self.tabla_tramos, 1)

        main_layout.addWidget(grp_tramos, 1)

        # 3. Observación al pie y Leyenda RITO
        grp_pie = QGroupBox("3. Observaciones de Pie de Página y Notificación RITO")
        lay_pie = QVBoxLayout(grp_pie)
        self.txt_obs_pie = QPlainTextEdit()
        self.txt_obs_pie.setPlaceholderText("Observaciones generales al pie del ordenamiento...")
        self.txt_obs_pie.setMaximumHeight(60)
        lay_pie.addWidget(self.txt_obs_pie)
        lbl_rito = QLabel(
            "Leyenda impresa al pie: 'Tomo conocimiento del siguiente diagrama según RITO Art. 7º "
            "(Acatamiento debido por el personal a las órdenes de sus superiores. "
            "El personal deberá acatar de inmediato las órdenes que reciba de sus superiores, "
            "pudiendo plantear el caso después por vía jerárquica si considera que dichas órdenes "
            "no se ajustan a la reglamentación.)'"
        )
        lbl_rito.setWordWrap(True)
        lbl_rito.setStyleSheet("font-style: italic; color: #64748B;")
        lay_pie.addWidget(lbl_rito)
        main_layout.addWidget(grp_pie)

        # 4. Botones de Acción Final
        btn_lay = QHBoxLayout()
        btn_lay.addStretch()

        btn_cancel = QPushButton("Cancelar")
        btn_cancel.setObjectName("btnSecondary")
        btn_cancel.clicked.connect(self.reject)
        btn_lay.addWidget(btn_cancel)

        self.btn_guardar = QPushButton("💾 Guardar Ordenamiento")
        self.btn_guardar.clicked.connect(lambda: self._guardar(imprimir_pdf=False))
        btn_lay.addWidget(self.btn_guardar)

        self.btn_guardar_pdf = QPushButton("🖨️ Guardar y Generar PDF Duplicado")
        self.btn_guardar_pdf.clicked.connect(lambda: self._guardar(imprimir_pdf=True))
        btn_lay.addWidget(self.btn_guardar_pdf)

        main_layout.addLayout(btn_lay)

    def _on_deposito_changed(self, nuevo_dep: str) -> None:
        dep_clean = (nuevo_dep or "PC").strip().upper() or "PC"
        self.lbl_ref_dotacion.setText(f"REF de {dep_clean} (*):")
        self._actualizar_nro_orden_y_refs(dep_clean)

    def _actualizar_nro_orden_y_refs(self, dep_clean: str) -> None:
        if not self.ordenamiento_existente:
            self.txt_nro_orden.setText(self.records_service.generar_siguiente_nro_orden(dep_clean))
        else:
            actual = self.txt_nro_orden.text().strip()
            if " " in actual:
                sufijo = actual.split(" ", 1)[1]
                self.txt_nro_orden.setText(f"{dep_clean} {sufijo}")

        ref_actual = self.combo_ref_pct.currentText()
        self.combo_ref_pct.blockSignals(True)
        self.combo_ref_pct.clear()
        self.combo_ref_pct.addItem("")

        # Primero agregar personal de estación de esa dotación, luego el resto
        for pe in self._personal_estacion_cache:
            if str(pe.get("dotacion") or "").strip().upper() == dep_clean:
                self.combo_ref_pct.addItem(f"{pe['nombre']} ({pe['dotacion']})", userData=pe.get("dotacion"))
        for pe in self._personal_estacion_cache:
            if str(pe.get("dotacion") or "").strip().upper() != dep_clean:
                dot_txt = f" ({pe['dotacion']})" if pe.get("dotacion") else ""
                self.combo_ref_pct.addItem(f"{pe['nombre']}{dot_txt}", userData=pe.get("dotacion"))

        if ref_actual:
            self.combo_ref_pct.setCurrentText(ref_actual)
        elif self.combo_ref_pct.count() > 1:
            self.combo_ref_pct.setCurrentIndex(1)
        self.combo_ref_pct.blockSignals(False)

    def _on_ref_estacion_changed(self, texto: str) -> None:
        """Si la referencia seleccionada tiene una dotación entre paréntesis (ej. 'Juan Perez (LLV)'), sincroniza la dotación."""
        if "(" in texto and texto.endswith(")"):
            posible_dot = texto.rsplit("(", 1)[-1][:-1].strip().upper()
            if posible_dot and posible_dot != self.combo_deposito.currentText().strip().upper():
                self.combo_deposito.setCurrentText(posible_dot)

    def _buscar_legajo_modal(self) -> None:
        dlg = LegajoSearchDialog(self._empleados_cache, parent=self)
        if dlg.exec():
            emp = dlg.selected_empleado
            if emp:
                self._cargar_empleado(emp)

    def _on_legajo_changed(self) -> None:
        val = self.txt_legajo.text().strip()
        if not val:
            return
        try:
            leg = int(val)
            emp = next((e for e in self._empleados_cache if int(e["legajo"]) == leg), None)
            if emp:
                self._cargar_empleado(emp)
        except ValueError:
            pass

    def _cargar_empleado(self, emp: Dict[str, Any]) -> None:
        self.txt_legajo.setText(str(emp["legajo"]))
        self.txt_nombre.setText(str(emp["apellidos_nombres"]))
        dot = str(emp.get("dotacion") or "").strip().upper()
        if dot:
            self.combo_deposito.setCurrentText(dot)
        turno = str(emp.get("turnos") or "").strip()
        if turno:
            self.combo_turno.setCurrentText(turno)

    def _cargar_vueltas_de_turno(self) -> None:
        cod_turno = self.combo_turno.currentText().strip()
        if not cod_turno:
            QMessageBox.information(self, "Cargar Turno", "Ingrese o seleccione un código de Turno primero.")
            return
        itin = self.combo_itinerario.currentText().strip()
        t_info = self.records_service.obtener_turno_por_codigo(cod_turno, itinerario=itin or None)
        if not t_info or not t_info.get("vueltas"):
            QMessageBox.information(
                self, "Sin Vueltas", f"No se encontraron vueltas cargadas en itinerario para el turno '{cod_turno}'."
            )
            return
        for v in t_info["vueltas"]:
            self.tramos.append({
                "nro_secuencia": len(self.tramos) + 1,
                "ref": str(v.get("ref") or ""),
                "nro_tren": str(v.get("nro_tren") or ""),
                "origen": str(v.get("origen") or ""),
                "hora_salida": str(v.get("hora_salida") or ""),
                "destino": str(v.get("destino") or ""),
                "hora_llegada": str(v.get("hora_llegada") or ""),
                "observaciones": str(v.get("observaciones_tramo") or ""),
            })
        self._render_tramos()

    def _on_tramo_tren_changed(self, nro_tren: str) -> None:
        nt = (nro_tren or "").strip()
        self._paradas_tren_actual.clear()
        if not nt:
            return
        itin = self.combo_itinerario.currentText().strip()
        paradas = self.records_service.obtener_estaciones_y_horarios_por_tren(nt, itinerario=itin or None)
        if not paradas:
            return

        self.combo_tramo_origen.blockSignals(True)
        self.combo_tramo_destino.blockSignals(True)
        self.combo_tramo_origen.clear()
        self.combo_tramo_destino.clear()

        for p in paradas:
            est = p["estacion"]
            hr = p["hora"]
            self._paradas_tren_actual[est.upper()] = hr
            self.combo_tramo_origen.addItem(est)
            self.combo_tramo_destino.addItem(est)

        if len(paradas) >= 1:
            self.combo_tramo_origen.setCurrentIndex(0)
            self.txt_tramo_h_sal.setText(paradas[0]["hora"])
        if len(paradas) >= 2:
            self.combo_tramo_destino.setCurrentIndex(len(paradas) - 1)
            self.txt_tramo_h_lleg.setText(paradas[-1]["hora"])

        self.combo_tramo_origen.blockSignals(False)
        self.combo_tramo_destino.blockSignals(False)

    def _on_origen_selected(self, estacion: str) -> None:
        hr = self._paradas_tren_actual.get((estacion or "").strip().upper())
        if hr:
            self.txt_tramo_h_sal.setText(hr)

    def _on_destino_selected(self, estacion: str) -> None:
        hr = self._paradas_tren_actual.get((estacion or "").strip().upper())
        if hr:
            self.txt_tramo_h_lleg.setText(hr)

    def _abrir_crud_paradas(self) -> None:
        tren_actual = self.txt_tramo_tren.text().strip()
        dlg = AdminTrenParadasDialog(
            self.records_service,
            usuario_actual=self.usuario_actual,
            tren_inicial=tren_actual,
            parent=self,
        )
        dlg.exec()
        if tren_actual:
            self._on_tramo_tren_changed(tren_actual)

    def _agregar_tramo_tabla(self) -> None:
        self._sincronizar_ediciones_celdas()
        nuevo = {
            "nro_secuencia": len(self.tramos) + 1,
            "ref": self.txt_tramo_ref.text().strip(),
            "nro_tren": self.txt_tramo_tren.text().strip(),
            "origen": self.combo_tramo_origen.currentText().strip(),
            "hora_salida": self.txt_tramo_h_sal.text().strip(),
            "destino": self.combo_tramo_destino.currentText().strip(),
            "hora_llegada": self.txt_tramo_h_lleg.text().strip(),
            "observaciones": self.txt_tramo_obs.text().strip(),
        }
        if not any([nuevo["nro_tren"], nuevo["origen"], nuevo["destino"], nuevo["observaciones"]]):
            QMessageBox.warning(
                self, "Tramo Vacío", "Ingrese al menos un N° de tren, estación o una observación para el tramo."
            )
            return
        self.tramos.append(nuevo)
        self._render_tramos()
        self.txt_tramo_ref.clear()
        self.txt_tramo_tren.clear()
        self.combo_tramo_origen.setCurrentText("")
        self.txt_tramo_h_sal.clear()
        self.combo_tramo_destino.setCurrentText("")
        self.txt_tramo_h_lleg.clear()
        self.txt_tramo_obs.clear()

    def _on_seleccion_tramo_tabla(self) -> None:
        sel = self.tabla_tramos.selectionModel().selectedRows()
        if not sel:
            return
        idx = sel[0].row()
        if 0 <= idx < len(self.tramos):
            t = self.tramos[idx]
            self.txt_tramo_ref.setText(str(t.get("ref") or ""))
            self.txt_tramo_tren.blockSignals(True)
            self.txt_tramo_tren.setText(str(t.get("nro_tren") or ""))
            self.txt_tramo_tren.blockSignals(False)
            self.combo_tramo_origen.setCurrentText(str(t.get("origen") or ""))
            self.txt_tramo_h_sal.setText(str(t.get("hora_salida") or ""))
            self.combo_tramo_destino.setCurrentText(str(t.get("destino") or ""))
            self.txt_tramo_h_lleg.setText(str(t.get("hora_llegada") or ""))
            self.txt_tramo_obs.setText(str(t.get("observaciones") or ""))

    def _actualizar_tramo_seleccionado(self) -> None:
        sel = self.tabla_tramos.selectionModel().selectedRows()
        if not sel:
            QMessageBox.information(self, "Actualizar Tramo", "Seleccione una fila en la tabla primero.")
            return
        idx = sel[0].row()
        self.tramos[idx] = {
            "nro_secuencia": idx + 1,
            "ref": self.txt_tramo_ref.text().strip(),
            "nro_tren": self.txt_tramo_tren.text().strip(),
            "origen": self.combo_tramo_origen.currentText().strip(),
            "hora_salida": self.txt_tramo_h_sal.text().strip(),
            "destino": self.combo_tramo_destino.currentText().strip(),
            "hora_llegada": self.txt_tramo_h_lleg.text().strip(),
            "observaciones": self.txt_tramo_obs.text().strip(),
        }
        self._render_tramos()

    def _mover_tramo(self, delta: int) -> None:
        self._sincronizar_ediciones_celdas()
        sel = self.tabla_tramos.selectionModel().selectedRows()
        if not sel:
            return
        idx = sel[0].row()
        nuevo_idx = idx + delta
        if 0 <= nuevo_idx < len(self.tramos):
            self.tramos[idx], self.tramos[nuevo_idx] = self.tramos[nuevo_idx], self.tramos[idx]
            self._render_tramos()
            self.tabla_tramos.selectRow(nuevo_idx)

    def _eliminar_tramo_seleccionado(self) -> None:
        self._sincronizar_ediciones_celdas()
        sel = self.tabla_tramos.selectionModel().selectedRows()
        if not sel:
            return
        idx = sel[0].row()
        if 0 <= idx < len(self.tramos):
            del self.tramos[idx]
            self._render_tramos()

    def _sincronizar_ediciones_celdas(self) -> None:
        """Sincroniza cualquier edición hecha directamente sobre las celdas de la tabla de tramos."""
        nuevos: List[Dict[str, Any]] = []
        for r in range(self.tabla_tramos.rowCount()):
            nuevos.append({
                "nro_secuencia": r + 1,
                "ref": (self.tabla_tramos.item(r, 1).text() if self.tabla_tramos.item(r, 1) else "").strip(),
                "nro_tren": (self.tabla_tramos.item(r, 2).text() if self.tabla_tramos.item(r, 2) else "").strip(),
                "origen": (self.tabla_tramos.item(r, 3).text() if self.tabla_tramos.item(r, 3) else "").strip(),
                "hora_salida": (self.tabla_tramos.item(r, 4).text() if self.tabla_tramos.item(r, 4) else "").strip(),
                "destino": (self.tabla_tramos.item(r, 5).text() if self.tabla_tramos.item(r, 5) else "").strip(),
                "hora_llegada": (self.tabla_tramos.item(r, 6).text() if self.tabla_tramos.item(r, 6) else "").strip(),
                "observaciones": (self.tabla_tramos.item(r, 7).text() if self.tabla_tramos.item(r, 7) else "").strip(),
            })
        if nuevos:
            self.tramos = nuevos

    def _render_tramos(self) -> None:
        self.tabla_tramos.setRowCount(len(self.tramos))
        for r, t in enumerate(self.tramos):
            t["nro_secuencia"] = r + 1
            it_sec = QTableWidgetItem(str(r + 1))
            it_sec.setFlags(it_sec.flags() & ~Qt.ItemIsEditable)
            self.tabla_tramos.setItem(r, 0, it_sec)
            self.tabla_tramos.setItem(r, 1, QTableWidgetItem(str(t.get("ref") or "")))
            self.tabla_tramos.setItem(r, 2, QTableWidgetItem(str(t.get("nro_tren") or "")))
            self.tabla_tramos.setItem(r, 3, QTableWidgetItem(str(t.get("origen") or "")))
            self.tabla_tramos.setItem(r, 4, QTableWidgetItem(str(t.get("hora_salida") or "")))
            self.tabla_tramos.setItem(r, 5, QTableWidgetItem(str(t.get("destino") or "")))
            self.tabla_tramos.setItem(r, 6, QTableWidgetItem(str(t.get("hora_llegada") or "")))
            self.tabla_tramos.setItem(r, 7, QTableWidgetItem(str(t.get("observaciones") or "")))

    def _cargar_datos_existentes(self) -> None:
        o = self.ordenamiento_existente or {}
        dep = str(o.get("deposito") or "PC").strip().upper()
        self.combo_deposito.setCurrentText(dep)
        self.txt_nro_orden.setText(str(o.get("nro_orden") or ""))

        if o.get("fecha_emision"):
            dt = parsear_fecha(o["fecha_emision"], obligatorio=False)
            if dt:
                self.date_emision.setDate(QDate(dt.year, dt.month, dt.day))

        if o.get("hora_emision"):
            partes = str(o["hora_emision"]).split(":")
            if len(partes) >= 2:
                try:
                    self.time_emision.setTime(QTime(int(partes[0]), int(partes[1])))
                except ValueError:
                    pass

        self.txt_legajo.setText(str(o.get("legajo_personal") or ""))
        self.txt_nombre.setText(str(o.get("nombre_personal") or ""))
        self.combo_itinerario.setCurrentText(str(o.get("itinerario") or ""))
        self.combo_turno.setCurrentText(str(o.get("ciclo_o_turno") or ""))
        reg = str(o.get("regimen") or "LV")
        idx_r = self.combo_regimen.findText(reg)
        if idx_r >= 0:
            self.combo_regimen.setCurrentIndex(idx_r)

        self.combo_ref_pct.setCurrentText(str(o.get("ref_pct") or ""))
        self.txt_motivo.setText(str(o.get("motivo_orden") or ""))
        if o.get("emisor"):
            self.txt_emisor.setText(str(o["emisor"]))
        if o.get("usuario_windows"):
            self.txt_usuario_windows.setText(str(o["usuario_windows"]))
        self.txt_obs_pie.setPlainText(str(o.get("observacion_pie") or ""))
        self._render_tramos()

    def _guardar(self, imprimir_pdf: bool = False) -> None:
        self._sincronizar_ediciones_celdas()
        win_usr = self.txt_usuario_windows.text().strip() or obtener_usuario_windows()

        datos = {
            "nro_orden": self.txt_nro_orden.text().strip(),
            "fecha_emision": self.date_emision.date().toString("dd/MM/yyyy"),
            "hora_emision": self.time_emision.time().toString("HH:mm"),
            "legajo_personal": self.txt_legajo.text().strip(),
            "nombre_personal": self.txt_nombre.text().strip(),
            "deposito": self.combo_deposito.currentText().strip().upper(),
            "itinerario": self.combo_itinerario.currentText().strip(),
            "ciclo_o_turno": self.combo_turno.currentText().strip(),
            "regimen": self.combo_regimen.currentText().strip(),
            "motivo_orden": self.txt_motivo.text().strip(),
            "observacion_pie": self.txt_obs_pie.toPlainText().strip(),
            "ref_pct": self.combo_ref_pct.currentText().strip(),
            "emisor": self.txt_emisor.text().strip(),
            "usuario_windows": win_usr,
        }

        permisos = set(self.usuario_actual.get("permisos") or [])
        es_admin = (self.usuario_actual.get("rol") or "").lower() == "administrador"
        bypass = es_admin or ("registros.modificar" in permisos)

        try:
            if self.ordenamiento_existente and self.ordenamiento_existente.get("id"):
                ord_id = int(self.ordenamiento_existente["id"])
                self.records_service.actualizar_ordenamiento(
                    ord_id,
                    datos,
                    self.tramos,
                    user_id=self.usuario_actual.get("id"),
                    windows_user=win_usr,
                    bypass_window=bypass,
                )
            else:
                ord_id = self.records_service.crear_ordenamiento(
                    datos,
                    self.tramos,
                    user_id=self.usuario_actual.get("id"),
                    windows_user=win_usr,
                )
            self.saved_ordenamiento_id = ord_id

            if imprimir_pdf:
                completo = self.records_service.obtener_ordenamiento_completo(ord_id)
                if completo:
                    nro_safe = str(completo.get("nro_orden") or ord_id).replace(" ", "_")
                    default_name = f"Orden_Servicio_{nro_safe}.pdf"
                    out_path, _ = QFileDialog.getSaveFileName(
                        self,
                        "Guardar Orden de Servicio por Duplicado (PDF)",
                        default_name,
                        "Documento PDF (*.pdf)",
                    )
                    if out_path:
                        self.pdf_service.generar_orden_servicio_duplicado(
                            completo, completo.get("tramos") or [], out_path
                        )
                        try:
                            import os
                            os.startfile(out_path)
                        except Exception:
                            pass

            self.accept()
        except Exception as e:
            QMessageBox.critical(self, "Error al Guardar", f"No se pudo guardar el ordenamiento:\n{e}")
