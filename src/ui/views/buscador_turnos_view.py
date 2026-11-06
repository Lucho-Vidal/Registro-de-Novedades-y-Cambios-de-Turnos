"""Pestaña Principal: Buscador de Turnos e Itinerarios Ferroviarios.

Integra búsqueda dual por tren o por código de turno, layout Master-Detail apilado
con todas las vueltas encontradas, filtro de itinerarios multiselección,
evaluación de titulares y reemplazos por novedad activa, y exportación a Excel.
"""

from typing import Any, Dict, List, Optional
from datetime import datetime

from PySide6.QtCore import QDate, QModelIndex, Qt, Signal
from PySide6.QtGui import QColor, QPainter, QPen, QStandardItem, QStandardItemModel
from PySide6.QtWidgets import (
    QComboBox,
    QDateEdit,
    QFileDialog,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QSplitter,
    QStyledItemDelegate,
    QStyleOptionViewItem,
    QTableView,
    QVBoxLayout,
    QWidget,
)

from src.core.logger import get_logger
from src.core.settings import AppSettings
from src.core.theme import obtener_tema_visual
from src.services.excel_service import ExcelService
from src.services.records_service import RecordsService
from src.ui.models.base_table_model import BaseTableModel

logger = get_logger("BuscadorTurnosView")


class TurnoSeparatorDelegate(QStyledItemDelegate):
    """Dibuja una línea de borde inferior más fuerte para diferenciar turnos diferentes."""

    def __init__(self, get_tema_oscuro_fn, parent=None):
        super().__init__(parent)
        self.get_tema_oscuro_fn = get_tema_oscuro_fn

    def paint(self, painter: QPainter, option: QStyleOptionViewItem, index: QModelIndex) -> None:
        super().paint(painter, option, index)

        model = index.model()
        if not model:
            return

        curr_row = index.row()
        total_rows = model.rowCount()

        # Comprobar si la fila siguiente pertenece a un turno distinto
        es_cambio_turno = False
        if curr_row < total_rows - 1:
            val_curr = model.index(curr_row, 0).data(Qt.DisplayRole)
            val_next = model.index(curr_row + 1, 0).data(Qt.DisplayRole)
            if val_curr != val_next:
                es_cambio_turno = True

        if es_cambio_turno:
            painter.save()
            es_oscuro = self.get_tema_oscuro_fn()
            # Borde visualmente nítido y destacado (cyan/azul cielo en oscuro, azul cobalto en claro)
            color_borde = QColor("#38BDF8") if es_oscuro else QColor("#1D4ED8")
            pen = QPen(color_borde, 2)
            painter.setPen(pen)
            r = option.rect
            y = r.bottom()
            painter.drawLine(r.left(), y, r.right(), y)
            painter.restore()


class CheckableComboBox(QComboBox):
    """ComboBox con soporte de selección múltiple mediante casillas de verificación."""

    selectionChanged = Signal(list)

    def __init__(self, placeholder: str = "Todos los itinerarios", parent=None):
        super().__init__(parent)
        self.setEditable(True)
        self.lineEdit().setReadOnly(True)
        self.closeOnLineEditClick = False
        self._placeholder = placeholder
        self.lineEdit().setText(self._placeholder)
        self._model = QStandardItemModel(self)
        self.setModel(self._model)
        self._model.dataChanged.connect(self._on_data_changed)

    def set_items(self, items: List[str]) -> None:
        self.blockSignals(True)
        self._model.clear()

        # Opción 0: Seleccionar Todos
        item_all = QStandardItem("Seleccionar Todos")
        item_all.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
        item_all.setData(Qt.Unchecked, Qt.CheckStateRole)
        self._model.appendRow(item_all)

        for text in items:
            item = QStandardItem(text)
            item.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
            item.setData(Qt.Unchecked, Qt.CheckStateRole)
            self._model.appendRow(item)

        self.blockSignals(False)
        self._update_text()

    def _on_data_changed(self, top_left, bottom_right) -> None:
        row = top_left.row()
        state = top_left.data(Qt.CheckStateRole)
        if row == 0:
            self.blockSignals(True)
            for r in range(1, self._model.rowCount()):
                it = self._model.item(r)
                if it:
                    it.setData(state, Qt.CheckStateRole)
            self.blockSignals(False)
        self._update_text()
        self.selectionChanged.emit(self.get_checked_items())

    def _update_text(self) -> None:
        checked = self.get_checked_items()
        total = max(0, self._model.rowCount() - 1)
        if not checked or len(checked) == total:
            self.lineEdit().setText(self._placeholder)
        elif len(checked) == 1:
            self.lineEdit().setText(checked[0])
        else:
            self.lineEdit().setText(f"{len(checked)} seleccionados ({', '.join(checked)})")

    def get_checked_items(self) -> List[str]:
        res = []
        for r in range(1, self._model.rowCount()):
            it = self._model.item(r)
            if it and it.data(Qt.CheckStateRole) == Qt.Checked:
                res.append(it.text())
        return res


class BuscadorTurnosView(QWidget):
    """Pestaña principal que integra búsqueda dual de turnos y trenes con vueltas apiladas."""

    def __init__(
        self,
        records_service: RecordsService,
        excel_service: ExcelService,
        usuario_actual: Optional[Dict[str, Any]] = None,
        tiene_permiso_fn=None,
        parent=None,
    ):
        super().__init__(parent)
        self.records_service = records_service
        self.excel_service = excel_service
        self.usuario_actual = usuario_actual or {}
        self.tiene_permiso_fn = tiene_permiso_fn
        self.settings = AppSettings()

        self._turnos_resultados: List[Dict[str, Any]] = []
        self._vueltas_todas: List[Dict[str, Any]] = []
        self._ordenamientos_resultados: List[Dict[str, Any]] = []
        self._ultimo_criterio_buscado: str = ""

        self._init_ui()
        self.cargar_itinerarios()
        # Las grillas permanecen vacías al inicio para apertura instantánea sin bloqueos

    def _can(self, permission_code: str) -> bool:
        if self.tiene_permiso_fn is None:
            return True
        return bool(self.tiene_permiso_fn(permission_code))

    def _init_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        # 1. Barra de Filtros y Búsqueda
        filter_card = QFrame()
        filter_card.setObjectName("cardFrame")
        filter_layout = QVBoxLayout(filter_card)
        filter_layout.setContentsMargins(10, 8, 10, 8)
        filter_layout.setSpacing(8)

        # Fila 1: Input dual y botones
        row1 = QHBoxLayout()
        row1.addWidget(QLabel("<b>Buscar:</b>"))
        self.txt_criterio = QLineEdit()
        self.txt_criterio.setPlaceholderText("Ingrese N° de Tren (ej: 2303, 1310) o Código de Turno (ej: 401, 410, PROG.1CK5)...")
        self.txt_criterio.setClearButtonEnabled(True)
        self.txt_criterio.returnPressed.connect(self.consultar)
        row1.addWidget(self.txt_criterio, stretch=3)

        self.btn_buscar = QPushButton("🔍 Consultar")
        self.btn_buscar.clicked.connect(self.consultar)
        row1.addWidget(self.btn_buscar)

        self.btn_limpiar = QPushButton("Limpiar")
        self.btn_limpiar.setObjectName("btnSecondary")
        self.btn_limpiar.clicked.connect(self._limpiar_filtros)
        row1.addWidget(self.btn_limpiar)

        row1.addStretch(1)

        if self._can("turnos.administrar") or self._can("turnos.crear"):
            self.btn_importar = QPushButton("📥 Importar Itinerario (Excel)")
            self.btn_importar.setObjectName("btnSecondary")
            self.btn_importar.clicked.connect(self._importar_itinerario_excel)
            row1.addWidget(self.btn_importar)

        if self._can("turnos.administrar") or self._can("turnos.editar"):
            self.btn_admin_turnos = QPushButton("⚙️ Administrar Turnos")
            self.btn_admin_turnos.setObjectName("btnSecondary")
            self.btn_admin_turnos.clicked.connect(self._abrir_admin_turnos)
            row1.addWidget(self.btn_admin_turnos)

        if self._can("turnos.ver") or self._can("turnos.administrar"):
            self.btn_paradas_tren = QPushButton("🚉 Paradas por Tren")
            self.btn_paradas_tren.setObjectName("btnSecondary")
            self.btn_paradas_tren.clicked.connect(self._abrir_paradas_tren)
            row1.addWidget(self.btn_paradas_tren)

        if self._can("turnos.exportar") or self._can("turnos.ver"):
            self.btn_exportar = QPushButton("📊 Exportar a Excel")
            self.btn_exportar.setObjectName("btnSecondary")
            self.btn_exportar.clicked.connect(self._exportar_grilla_excel)
            row1.addWidget(self.btn_exportar)

        filter_layout.addLayout(row1)

        # Fila 2: Fecha, Tipo de Servicio y Selección Múltiple de Itinerarios
        row2 = QHBoxLayout()
        row2.addWidget(QLabel("Fecha:"))
        self.date_picker = QDateEdit()
        self.date_picker.setCalendarPopup(True)
        self.date_picker.setDisplayFormat("dd/MM/yyyy")
        self.date_picker.setDate(QDate.currentDate())
        self.date_picker.dateChanged.connect(self._on_fecha_changed)
        row2.addWidget(self.date_picker)

        row2.addWidget(QLabel("Tipo de Servicio:"))
        self.combo_servicio = QComboBox()
        self.combo_servicio.addItems(["Día Hábil (Lunes a Viernes)", "Sábado", "Domingo / Feriado"])
        self.combo_servicio.currentIndexChanged.connect(self.consultar)
        row2.addWidget(self.combo_servicio)

        row2.addWidget(QLabel("Itinerarios:"))
        self.combo_itinerarios = CheckableComboBox(placeholder="Todos los itinerarios")
        self.combo_itinerarios.selectionChanged.connect(lambda _: self.consultar())
        self.combo_itinerarios.setMinimumWidth(220)
        row2.addWidget(self.combo_itinerarios)

        row2.addStretch()
        filter_layout.addLayout(row2)

        layout.addWidget(filter_card)

        # 2. Splitter Master-Detail
        splitter = QSplitter(Qt.Vertical)

        # Panel Master: Turnos Encontrados
        grp_turnos = QGroupBox("Turnos Encontrados (Haga clic en un turno para ubicar sus vueltas abajo)")
        lay_turnos = QVBoxLayout(grp_turnos)

        headers_turnos = [
            "CÓDIGO TURNO", "BASE", "TIPO", "DEPÓSITO", "TOMA", "DEJE", "DURACIÓN",
            "TITULAR / REEMPLAZO ASIGNADO", "ESTADO / COBERTURA", "ITINERARIO",
        ]
        keys_turnos = [
            "codigo_turno", "turno_base", "tipo_turno", "deposito_base", "toma_servicio",
            "deje_servicio", "duracion_jornada", "titular", "cobertura", "itinerario_vigente",
        ]

        def conditional_colors_turnos(row_item: Dict[str, Any], col_idx: int):
            alerta = row_item.get("estado_alerta", "")
            es_oscuro = obtener_tema_visual() == "oscuro"
            if alerta == "RELEVO":
                return (QColor("#132E1F"), QColor("#6EE7B7")) if es_oscuro else (QColor("#E8F5E9"), QColor("#1B5E20"))
            elif alerta == "DESCUBIERTO":
                return (QColor("#3D1B1B"), QColor("#FCA5A5")) if es_oscuro else (QColor("#FFEBEE"), QColor("#B71C1C"))
            elif alerta == "VACANTE":
                return (QColor("#3D3214"), QColor("#FCD34D")) if es_oscuro else (QColor("#FFF8E1"), QColor("#B78103"))
            elif alerta == "CAMBIO":
                return (QColor("#1E3A8A"), QColor("#93C5FD")) if es_oscuro else (QColor("#EFF6FF"), QColor("#1D4ED8"))
            elif alerta == "ORDENAMIENTO":
                return (QColor("#312E81"), QColor("#C7D2FE")) if es_oscuro else (QColor("#EEF2FF"), QColor("#3730A3"))
            return None

        self.model_turnos = BaseTableModel(headers_turnos, keys_turnos, conditional_colors=conditional_colors_turnos, parent=self)
        self.table_turnos = QTableView()
        self.table_turnos.setModel(self.model_turnos)
        self.table_turnos.setSelectionBehavior(QTableView.SelectRows)
        self.table_turnos.setSelectionMode(QTableView.SingleSelection)
        self.table_turnos.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table_turnos.setColumnWidth(0, 95)
        self.table_turnos.setColumnWidth(1, 85)
        self.table_turnos.setColumnWidth(2, 75)
        self.table_turnos.setColumnWidth(3, 75)
        self.table_turnos.setColumnWidth(4, 80)
        self.table_turnos.setColumnWidth(5, 110)
        self.table_turnos.setColumnWidth(6, 140)
        self.table_turnos.setColumnWidth(7, 210)
        self.table_turnos.horizontalHeader().setSectionResizeMode(8, QHeaderView.Stretch)
        self.table_turnos.selectionModel().selectionChanged.connect(self._on_turno_seleccionado)
        lay_turnos.addWidget(self.table_turnos)
        splitter.addWidget(grp_turnos)

        # Panel Detail: Vueltas de todos los turnos encontrados apiladas
        self.grp_vueltas = QGroupBox("Vueltas de Turnos Encontrados")
        lay_vueltas = QVBoxLayout(self.grp_vueltas)

        headers_vueltas = [
            "TURNO", "VUELTA N°", "REF.", "TREN", "ORIGEN", "SALIDA",
            "DESTINO", "LLEGADA", "OBSERVACIONES DEL TRAMO",
        ]
        keys_vueltas = [
            "codigo_turno", "nro_vuelta", "ref", "nro_tren", "origen", "hora_salida",
            "destino", "hora_llegada", "observaciones_tramo",
        ]

        def conditional_colors_vueltas(row_item: Dict[str, Any], col_idx: int):
            crit = str(self._ultimo_criterio_buscado or "").strip()
            es_oscuro = obtener_tema_visual() == "oscuro"
            if crit and str(row_item.get("nro_tren", "")).strip() == crit:
                # Resaltado en ámbar/dorado de la coincidencia
                return (QColor("#78350F"), QColor("#FDE68A")) if es_oscuro else (QColor("#FEF3C7"), QColor("#92400E"))
            return None

        self.model_vueltas = BaseTableModel(headers_vueltas, keys_vueltas, conditional_colors=conditional_colors_vueltas, parent=self)
        self.table_vueltas = QTableView()
        self.table_vueltas.setModel(self.model_vueltas)
        self.table_vueltas.setItemDelegate(
            TurnoSeparatorDelegate(
                lambda: obtener_tema_visual() == "oscuro",
                parent=self.table_vueltas,
            )
        )
        self.table_vueltas.setSelectionBehavior(QTableView.SelectRows)
        self.table_vueltas.setSelectionMode(QTableView.SingleSelection)
        self.table_vueltas.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table_vueltas.setColumnWidth(0, 95)
        self.table_vueltas.setColumnWidth(1, 80)
        self.table_vueltas.setColumnWidth(2, 65)
        self.table_vueltas.setColumnWidth(3, 75)
        self.table_vueltas.setColumnWidth(4, 110)
        self.table_vueltas.setColumnWidth(5, 75)
        self.table_vueltas.setColumnWidth(6, 110)
        self.table_vueltas.setColumnWidth(7, 75)
        self.table_vueltas.horizontalHeader().setSectionResizeMode(8, QHeaderView.Stretch)
        lay_vueltas.addWidget(self.table_vueltas)
        splitter.addWidget(self.grp_vueltas)

        # Panel 3: Ordenamientos Operativos emitidos para el Tren / Turno buscado
        self.grp_ordenamientos = QGroupBox("📄 Ordenamientos Operativos Emitidos para la Búsqueda (Doble clic para ver / imprimir PDF)")
        lay_ords = QVBoxLayout(self.grp_ordenamientos)

        headers_ords = [
            "N° ORDEN", "FECHA", "HORA", "LEGAJO", "PERSONAL ASIGNADO",
            "DOTACIÓN", "TURNO / CICLO", "TRENES INVOLUCRADOS", "MOTIVO DE LA ORDEN",
            "REF ESTACIÓN", "EMISOR",
        ]
        keys_ords = [
            "nro_orden", "fecha_emision", "hora_emision", "legajo_personal", "nombre_personal",
            "deposito", "ciclo_o_turno", "trenes_involucrados", "motivo_orden",
            "ref_pct", "emisor",
        ]

        self.model_ordenamientos = BaseTableModel(headers_ords, keys_ords, parent=self)
        self.table_ordenamientos = QTableView()
        self.table_ordenamientos.setModel(self.model_ordenamientos)
        self.table_ordenamientos.setSelectionBehavior(QTableView.SelectRows)
        self.table_ordenamientos.setSelectionMode(QTableView.SingleSelection)
        self.table_ordenamientos.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table_ordenamientos.setColumnWidth(0, 115)
        self.table_ordenamientos.setColumnWidth(1, 85)
        self.table_ordenamientos.setColumnWidth(2, 60)
        self.table_ordenamientos.setColumnWidth(3, 75)
        self.table_ordenamientos.setColumnWidth(4, 175)
        self.table_ordenamientos.setColumnWidth(5, 75)
        self.table_ordenamientos.setColumnWidth(6, 95)
        self.table_ordenamientos.setColumnWidth(7, 130)
        self.table_ordenamientos.horizontalHeader().setSectionResizeMode(8, QHeaderView.Stretch)
        self.table_ordenamientos.setColumnWidth(9, 120)
        self.table_ordenamientos.setColumnWidth(10, 95)
        self.table_ordenamientos.doubleClicked.connect(self._abrir_ordenamiento_detalle)
        lay_ords.addWidget(self.table_ordenamientos)
        splitter.addWidget(self.grp_ordenamientos)

        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 3)
        splitter.setStretchFactor(2, 2)
        layout.addWidget(splitter)

        # Inicializar selección según fecha actual
        self._on_fecha_changed(self.date_picker.date())

    def _on_fecha_changed(self, qdate: QDate) -> None:
        """Preselecciona inteligentemente el tipo de servicio según el día de la semana."""
        dia_semana_py = qdate.toPython().weekday()  # 0=Lunes, 5=Sábado, 6=Domingo
        if dia_semana_py < 5:
            self.combo_servicio.setCurrentIndex(0)  # Hábil
        elif dia_semana_py == 5:
            self.combo_servicio.setCurrentIndex(1)  # Sábado
        else:
            self.combo_servicio.setCurrentIndex(2)  # Domingo/Feriado

    def cargar_itinerarios(self) -> None:
        itis = self.records_service.listar_itinerarios_disponibles()
        self.combo_itinerarios.set_items(itis)

    def consultar(self) -> None:
        criterio = self.txt_criterio.text().strip()
        self._ultimo_criterio_buscado = criterio
        if not criterio:
            self._turnos_resultados = []
            self._vueltas_todas = []
            self._ordenamientos_resultados = []
            self.model_turnos.set_data([])
            self.model_vueltas.set_data([])
            self.model_ordenamientos.set_data([])
            self.grp_vueltas.setTitle("Vueltas de Turnos Encontrados (0)")
            self.grp_ordenamientos.setTitle("📄 Ordenamientos Operativos Emitidos para la Búsqueda (0)")
            return

        fecha_iso = self.date_picker.date().toString("yyyy-MM-dd")
        tipo_dia = self.combo_servicio.currentText()

        itis_sel = self.combo_itinerarios.get_checked_items()
        lista_itis = itis_sel if itis_sel else None

        self._turnos_resultados = self.records_service.buscar_turnos_con_cobertura(
            criterio=criterio,
            fecha_target_iso=fecha_iso,
            tipo_dia=tipo_dia,
            itinerarios=lista_itis,
        )
        self.model_turnos.set_data(self._turnos_resultados)

        # Cargar todas las vueltas apiladas de todos los turnos encontrados
        turno_ids = [t["id"] for t in self._turnos_resultados]
        self._vueltas_todas = self.records_service.obtener_vueltas_de_turnos(turno_ids)
        self.model_vueltas.set_data(self._vueltas_todas)

        # Buscar ordenamientos operativos activos para este tren / turno / criterio
        self._ordenamientos_resultados = self.records_service.buscar_ordenamientos_activos(
            criterio=criterio,
            fecha_target_iso=fecha_iso,
        )
        # Si no hay en la fecha exacta, mostrar también ordenamientos activos de cualquier fecha que incluyan el tren
        if not self._ordenamientos_resultados:
            self._ordenamientos_resultados = self.records_service.buscar_ordenamientos_activos(
                criterio=criterio,
                fecha_target_iso=None,
            )
        self.model_ordenamientos.set_data(self._ordenamientos_resultados)

        cant_turnos = len(self._turnos_resultados)
        cant_vueltas = len(self._vueltas_todas)
        cant_ords = len(self._ordenamientos_resultados)
        self.grp_vueltas.setTitle(
            f"Vueltas de Turnos Encontrados ({cant_vueltas} tramos en {cant_turnos} turnos)"
        )
        self.grp_ordenamientos.setTitle(
            f"📄 Ordenamientos Operativos Emitidos para '{criterio}' ({cant_ords} ordenamiento(s)) - Doble clic para ver / imprimir PDF"
        )

        if self._turnos_resultados:
            self.table_turnos.selectRow(0)

    def _on_turno_seleccionado(self) -> None:
        """Al seleccionar un turno en la grilla superior, desplazar la grilla inferior hasta sus vueltas."""
        indexes = self.table_turnos.selectionModel().selectedRows()
        if not indexes or not self._vueltas_todas:
            return

        row_idx = indexes[0].row()
        turno = self.model_turnos.get_row(row_idx)
        if not turno:
            return

        cod_t = turno.get("codigo_turno")
        for v_idx, v in enumerate(self._vueltas_todas):
            if v.get("codigo_turno") == cod_t:
                model_idx = self.model_vueltas.index(v_idx, 0)
                self.table_vueltas.scrollTo(model_idx)
                self.table_vueltas.selectRow(v_idx)
                break

    def _limpiar_filtros(self) -> None:
        self.txt_criterio.clear()
        self.date_picker.setDate(QDate.currentDate())
        self.combo_itinerarios.set_items(self.records_service.listar_itinerarios_disponibles())
        self.consultar()

    def _abrir_admin_turnos(self) -> None:
        """Abre la ventana modal de gestión completa de turnos y sus vueltas."""
        from src.ui.views.dialogs.admin_turnos_dialog import AdminTurnosDialog

        dlg = AdminTurnosDialog(
            self.records_service,
            usuario_actual=self.usuario_actual,
            tiene_permiso_fn=self.tiene_permiso_fn,
            parent=self,
        )
        if dlg.exec():
            self.cargar_itinerarios()
            self.consultar()

    def _abrir_paradas_tren(self) -> None:
        """Abre el diálogo CRUD e importador Excel de paradas y horarios de paso por tren."""
        from src.ui.views.dialogs.admin_tren_paradas_dialog import AdminTrenParadasDialog

        dlg = AdminTrenParadasDialog(
            self.records_service,
            usuario_actual=self.usuario_actual,
            tren_inicial=self.txt_criterio.text().strip(),
            parent=self,
        )
        dlg.exec()

    def _abrir_ordenamiento_detalle(self, index: QModelIndex) -> None:
        """Abre el diálogo de Ordenamiento al hacer doble clic desde los resultados del buscador."""
        from src.ui.views.dialogs.ordenamiento_dialog import OrdenamientoDialog

        row_data = self.model_ordenamientos.get_row(index.row())
        if not row_data or not row_data.get("id"):
            return
        completo = self.records_service.obtener_ordenamiento_completo(int(row_data["id"]))
        if not completo:
            return
        dlg = OrdenamientoDialog(
            self.records_service,
            ordenamiento_existente=completo,
            usuario_actual=self.usuario_actual,
            parent=self,
        )
        if dlg.exec():
            self.consultar()

    def _importar_itinerario_excel(self) -> None:
        f_path, _ = QFileDialog.getOpenFileName(
            self,
            "Seleccionar Archivo de Itinerario Normalizado",
            "",
            "Archivos Excel (*.xlsx *.xlsm)",
        )
        if not f_path:
            return

        try:
            turnos, vueltas = self.records_service.importar_itinerario_excel(f_path)
            QMessageBox.information(
                self,
                "Itinerario Importado",
                f"Se importaron con éxito {turnos} turnos y {vueltas} vueltas desde:\n{f_path}",
            )
            self.cargar_itinerarios()
            self.consultar()
        except Exception as e:
            QMessageBox.critical(self, "Error al Importar", f"No se pudo procesar el archivo:\n{e}")

    def _exportar_grilla_excel(self) -> None:
        if not self._turnos_resultados:
            QMessageBox.information(self, "Exportación", "No hay turnos para exportar.")
            return

        def_name = f"Turnos_Itinerario_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        f_path, _ = QFileDialog.getSaveFileName(self, "Exportar Consulta de Turnos", def_name, "Excel (*.xlsx)")
        if not f_path:
            return

        try:
            import openpyxl

            wb = openpyxl.Workbook()
            ws1 = wb.active
            ws1.title = "Turnos Encontrados"

            headers1 = [
                "Código Turno", "Turno Base", "Tipo", "Depósito", "Toma", "Deje", "Duración",
                "Titular / Reemplazo", "Estado Cobertura", "Itinerario Vigente"
            ]
            ws1.append(headers1)

            for t in self._turnos_resultados:
                ws1.append([
                    t.get("codigo_turno"),
                    t.get("turno_base"),
                    t.get("tipo_turno"),
                    t.get("deposito_base"),
                    t.get("toma_servicio"),
                    t.get("deje_servicio"),
                    t.get("duracion_jornada"),
                    t.get("titular"),
                    t.get("cobertura"),
                    t.get("itinerario_vigente"),
                ])

            # Pestaña 2: Todas las vueltas encontradas
            if self._vueltas_todas:
                ws2 = wb.create_sheet(title="Todas las Vueltas")
                headers2 = [
                    "Turno", "Vuelta N°", "Referencia", "Tren", "Origen",
                    "Salida", "Destino", "Llegada", "Observaciones",
                ]
                ws2.append(headers2)
                for v in self._vueltas_todas:
                    ws2.append([
                        v.get("codigo_turno"),
                        v.get("nro_vuelta"),
                        v.get("ref"),
                        v.get("nro_tren"),
                        v.get("origen"),
                        v.get("hora_salida"),
                        v.get("destino"),
                        v.get("llegada") or v.get("hora_llegada"),
                        v.get("observaciones_tramo"),
                    ])

            # Pestaña 3: Ordenamientos Operativos encontrados
            if self._ordenamientos_resultados:
                ws3 = wb.create_sheet(title="Ordenamientos Operativos")
                ws3.append([
                    "N° Orden", "Fecha", "Hora", "Legajo", "Personal Asignado",
                    "Dotación", "Turno / Ciclo", "Trenes Involucrados", "Motivo de la Orden",
                    "REF Estación", "Emisor",
                ])
                for o in self._ordenamientos_resultados:
                    ws3.append([
                        o.get("nro_orden"),
                        o.get("fecha_emision"),
                        o.get("hora_emision"),
                        o.get("legajo_personal"),
                        o.get("nombre_personal"),
                        o.get("deposito"),
                        o.get("ciclo_o_turno"),
                        o.get("trenes_involucrados"),
                        o.get("motivo_orden"),
                        o.get("ref_pct"),
                        o.get("emisor"),
                    ])

            wb.save(f_path)
            QMessageBox.information(self, "Exportación Exitosa", f"Archivo generado en:\n{f_path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Fallo al exportar:\n{e}")

    def actualizar_tema(self) -> None:
        """Refresca las tablas y delegados con el tema visual activo."""
        if hasattr(self, "table_turnos") and self.table_turnos.viewport():
            self.table_turnos.viewport().update()
        if hasattr(self, "table_vueltas") and self.table_vueltas.viewport():
            self.table_vueltas.viewport().update()
