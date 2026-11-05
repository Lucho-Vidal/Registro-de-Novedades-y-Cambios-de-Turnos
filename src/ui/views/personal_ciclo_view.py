"""Vista especializada para el seguimiento y gestión de Personal de Ciclo.

Permite consultar personal disponible y asignado en una fecha determinada,
con soporte interactivo de arrastrar y soltar (Drag and Drop) entre ambas tablas,
y acceso al Calendario Mensual de Ciclo.
"""

from typing import Any, Dict, List, Optional
import json

from PySide6.QtCore import QByteArray, QDate, QMimeData, QPoint, Qt, Signal
from PySide6.QtGui import QDrag
from PySide6.QtWidgets import (
    QComboBox,
    QDateEdit,
    QDialog,
    QDialogButtonBox,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QSplitter,
    QTableView,
    QVBoxLayout,
    QWidget,
)

from src.core.logger import get_logger
from src.services.excel_service import ExcelService
from src.services.pdf_service import PDFService
from src.services.records_service import RecordsService
from src.ui.models.base_table_model import BaseTableModel
from src.ui.views.dialogs.calendario_ciclo_dialog import CalendarioCicloDialog

logger = get_logger("PersonalCicloView")


class DragDropCicloTableView(QTableView):
    """Tabla con soporte Drag & Drop bidireccional para asignación y liberación de agentes de ciclo."""

    item_soltado = Signal(dict)

    def __init__(self, rol: str, parent=None):
        super().__init__(parent)
        self.rol = rol  # 'LIBRES' o 'OCUPADOS'
        self._drag_start_pos: Optional[QPoint] = None
        self.setAcceptDrops(True)
        self.viewport().setAcceptDrops(True)
        self.setDragDropMode(QTableView.DragDrop)
        self.setDropIndicatorShown(True)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self._drag_start_pos = event.pos()
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if not (event.buttons() & Qt.LeftButton) or not self._drag_start_pos:
            super().mouseMoveEvent(event)
            return

        if (event.pos() - self._drag_start_pos).manhattanLength() < 10:
            super().mouseMoveEvent(event)
            return

        indexes = self.selectionModel().selectedRows()
        if not indexes:
            return

        model = self.model()
        if hasattr(model, "get_row"):
            row_data = model.get_row(indexes[0].row())
        else:
            return

        drag = QDrag(self)
        mime = QMimeData()
        payload = json.dumps({"origen_rol": self.rol, "datos": row_data}).encode("utf-8")
        mime.setData("application/x-ciclo-row", QByteArray(payload))
        drag.setMimeData(mime)
        drag.exec(Qt.MoveAction)

    def dragEnterEvent(self, event):
        if event.mimeData().hasFormat("application/x-ciclo-row"):
            try:
                data_dict = json.loads(bytes(event.mimeData().data("application/x-ciclo-row")).decode("utf-8"))
                if data_dict.get("origen_rol") != self.rol:
                    event.acceptProposedAction()
                    return
            except Exception:
                pass
        event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasFormat("application/x-ciclo-row"):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event):
        if event.mimeData().hasFormat("application/x-ciclo-row"):
            try:
                data_dict = json.loads(bytes(event.mimeData().data("application/x-ciclo-row")).decode("utf-8"))
                if data_dict.get("origen_rol") != self.rol:
                    event.acceptProposedAction()
                    self.item_soltado.emit(data_dict.get("datos", {}))
                    return
            except Exception as e:
                logger.error("Error en dropEvent ciclo: %s", e)
        event.ignore()


class PersonalCicloView(QWidget):
    """Vista dividida en dos tablas: Personal Libre (arriba) y Personal Asignado (abajo) con Drag & Drop."""

    def __init__(
        self,
        records_service: RecordsService,
        pdf_service: PDFService,
        excel_service: ExcelService,
        usuario_actual: Optional[Dict[str, Any]] = None,
        parent=None,
    ):
        super().__init__(parent)
        self.records_service = records_service
        self.pdf_service = pdf_service
        self.excel_service = excel_service
        self.usuario_actual = usuario_actual or {}

        self._init_ui()
        self.consultar()

    def _init_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        # 1. Barra superior de filtros
        bar_frame = QFrame()
        bar_frame.setObjectName("filterBar")
        bar_layout = QHBoxLayout(bar_frame)
        bar_layout.setContentsMargins(8, 4, 8, 4)
        bar_layout.setSpacing(10)

        bar_layout.addWidget(QLabel("Fecha de Consulta:"))
        self.date_picker = QDateEdit()
        self.date_picker.setCalendarPopup(True)
        self.date_picker.setDisplayFormat("dd/MM/yyyy")
        self.date_picker.setDate(QDate.currentDate())
        bar_layout.addWidget(self.date_picker)

        bar_layout.addWidget(QLabel("Dotación:"))
        self.combo_dotacion = QComboBox()
        self.combo_dotacion.addItem("Todas")
        dots = self.records_service.listar_dotaciones(incluir_inactivos=False)
        for d in dots:
            self.combo_dotacion.addItem(d["nombre"])
        bar_layout.addWidget(self.combo_dotacion)

        bar_layout.addWidget(QLabel("Especialidad:"))
        self.combo_esp = QComboBox()
        self.combo_esp.addItem("Todas")
        for esp in self.records_service.listar_especialidades_empleados():
            self.combo_esp.addItem(esp)
        self.combo_esp.currentIndexChanged.connect(self.consultar)
        self.combo_dotacion.currentIndexChanged.connect(self.consultar)
        bar_layout.addWidget(self.combo_esp)

        self.btn_consultar = QPushButton("🔍 Actualizar")
        self.btn_consultar.clicked.connect(self.consultar)
        bar_layout.addWidget(self.btn_consultar)

        bar_layout.addStretch()

        self.btn_ver_calendario = QPushButton("📅 Ver Calendario Mensual")
        self.btn_ver_calendario.setObjectName("btnSecondary")
        self.btn_ver_calendario.clicked.connect(self._abrir_calendario_seleccionado)
        bar_layout.addWidget(self.btn_ver_calendario)

        layout.addWidget(bar_frame)

        # 2. Splitter vertical con las dos tablas
        splitter = QSplitter(Qt.Vertical)

        # Tabla Superior: Personal Libre / Sin Relevo
        grp_libres = QGroupBox("Personal de Ciclo DISPONIBLE (Arrastre hacia abajo para asignar a un turno sin cubrir)")
        lay_libres = QVBoxLayout(grp_libres)
        headers_libres = ["LEGAJO", "APELLIDOS Y NOMBRES", "ESPECIALIDAD", "DOTACIÓN", "TURNO HABITUAL", "FRANCO", "HORARIO FRANCO"]
        keys_libres = ["legajo", "apellidos_nombres", "especialidad", "dotacion", "turnos", "franco", "franco_horario"]

        self.model_libres = BaseTableModel(headers_libres, keys_libres, parent=self)
        self.table_libres = DragDropCicloTableView("LIBRES", parent=self)
        self.table_libres.setModel(self.model_libres)
        self.table_libres.setSelectionBehavior(QTableView.SelectRows)
        self.table_libres.setSelectionMode(QTableView.SingleSelection)
        self.table_libres.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table_libres.setColumnWidth(0, 80)   # LEGAJO
        self.table_libres.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)  # APELLIDOS Y NOMBRES
        self.table_libres.setColumnWidth(2, 180)  # ESPECIALIDAD (sin truncar)
        self.table_libres.setColumnWidth(3, 90)   # DOTACIÓN
        self.table_libres.setColumnWidth(4, 120)  # TURNO HABITUAL
        self.table_libres.setColumnWidth(5, 90)   # FRANCO
        self.table_libres.setColumnWidth(6, 130)  # HORARIO FRANCO
        self.table_libres.doubleClicked.connect(lambda idx: self._on_double_click(self.model_libres, idx))
        self.table_libres.item_soltado.connect(self._on_soltado_en_libres)
        lay_libres.addWidget(self.table_libres)
        splitter.addWidget(grp_libres)

        # Tabla Inferior: Personal Ocupado / Con Relevo
        grp_ocupados = QGroupBox("Personal de Ciclo ASIGNADO (Arrastre hacia arriba para liberar del relevo)")
        lay_ocupados = QVBoxLayout(grp_ocupados)
        headers_ocupados = [
            "LEGAJO", "APELLIDOS Y NOMBRES", "ESPECIALIDAD", "DOTACIÓN", "REEMPLAZA A (TITULAR)",
            "TURNO CUBIERTO", "MOTIVO / NOVEDAD", "DESDE", "HASTA",
        ]
        keys_ocupados = [
            "legajo", "apellidos_nombres", "especialidad", "dotacion", "reemplaza_a",
            "turno_titular", "tipo_novedad", "fecha_inicio_relevo", "fecha_fin_relevo",
        ]

        self.model_ocupados = BaseTableModel(headers_ocupados, keys_ocupados, parent=self)
        self.table_ocupados = DragDropCicloTableView("OCUPADOS", parent=self)
        self.table_ocupados.setModel(self.model_ocupados)
        self.table_ocupados.setSelectionBehavior(QTableView.SelectRows)
        self.table_ocupados.setSelectionMode(QTableView.SingleSelection)
        self.table_ocupados.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table_ocupados.setColumnWidth(0, 80)   # LEGAJO
        self.table_ocupados.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)  # APELLIDOS Y NOMBRES
        self.table_ocupados.setColumnWidth(2, 180)  # ESPECIALIDAD
        self.table_ocupados.setColumnWidth(3, 90)   # DOTACIÓN
        self.table_ocupados.setColumnWidth(4, 220)  # REEMPLAZA A
        self.table_ocupados.setColumnWidth(5, 110)  # TURNO CUBIERTO
        self.table_ocupados.setColumnWidth(6, 160)  # MOTIVO / NOVEDAD
        self.table_ocupados.setColumnWidth(7, 95)   # DESDE
        self.table_ocupados.setColumnWidth(8, 95)   # HASTA
        self.table_ocupados.doubleClicked.connect(lambda idx: self._on_double_click(self.model_ocupados, idx))
        self.table_ocupados.item_soltado.connect(self._on_soltado_en_ocupados)
        lay_ocupados.addWidget(self.table_ocupados)
        splitter.addWidget(grp_ocupados)

        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 1)
        layout.addWidget(splitter)

        # 3. Pie con instrucciones
        lbl_hint = QLabel("💡 Arrastre y suelte (Drag & Drop) entre ambas tablas para asignar o liberar relevos. Doble clic para abrir el Calendario Mensual.")
        lbl_hint.setStyleSheet("color: #6B7280; font-size: 8.5pt;")
        layout.addWidget(lbl_hint)

    def consultar(self) -> None:
        f_str = self.date_picker.date().toString("dd/MM/yyyy")
        dot = self.combo_dotacion.currentText()
        esp = self.combo_esp.currentText()

        try:
            res = self.records_service.obtener_estado_personal_ciclo(f_str, dotacion=dot, especialidad=esp)
            self.model_libres.set_data(res.get("libres", []))
            self.model_ocupados.set_data(res.get("ocupados", []))
        except Exception as e:
            logger.error(f"Error al consultar personal de ciclo: {e}")

    def _on_double_click(self, model: BaseTableModel, index) -> None:
        emp = model.get_row(index.row())
        if emp:
            self._abrir_calendario(emp)

    def _abrir_calendario_seleccionado(self) -> None:
        sel_l = self.table_libres.selectionModel().selectedRows()
        if sel_l:
            emp = self.model_libres.get_row(sel_l[0].row())
            if emp:
                self._abrir_calendario(emp)
                return

        sel_o = self.table_ocupados.selectionModel().selectedRows()
        if sel_o:
            emp = self.model_ocupados.get_row(sel_o[0].row())
            if emp:
                self._abrir_calendario(emp)
                return

        QMessageBox.information(self, "Selección", "Seleccione un empleado en cualquiera de las dos tablas para ver su calendario.")

    def _abrir_calendario(self, emp: Dict[str, Any]) -> None:
        qdate = self.date_picker.date()
        dlg = CalendarioCicloDialog(
            empleado=emp,
            records_service=self.records_service,
            pdf_service=self.pdf_service,
            excel_service=self.excel_service,
            mes=qdate.month(),
            anio=qdate.year(),
            usuario_actual=self.usuario_actual,
            parent=self,
        )
        dlg.exec()
        self.consultar()

    def _on_soltado_en_libres(self, emp: Dict[str, Any]) -> None:
        """Al arrastrar un agente ocupado hacia la grilla de libres: liberar del relevo."""
        reemplazo_id = emp.get("reemplazo_id")
        nombre = emp.get("apellidos_nombres", "")
        titular = emp.get("reemplaza_a", "")
        nov_id = emp.get("novedad_id", "")

        resp = QMessageBox.question(
            self,
            "Liberar Personal de Ciclo",
            f"¿Desea liberar al agente <b>{nombre}</b> del relevo activo de la novedad #{nov_id} (Titular: {titular})?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if resp == QMessageBox.Yes and reemplazo_id:
            try:
                self.records_service.desactivar_reemplazo(
                    int(reemplazo_id),
                    user_id=self.usuario_actual.get("id"),
                    windows_user=self.usuario_actual.get("username"),
                )
                self.consultar()
                QMessageBox.information(self, "Relevo Finalizado", f"El agente {nombre} ha sido liberado del relevo.")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"No se pudo liberar al agente:\n{e}")

    def _on_soltado_en_ocupados(self, emp: Dict[str, Any]) -> None:
        """Al arrastrar un agente libre hacia la grilla de ocupados: asignar como relevo a una novedad sin cubrir."""
        f_str = self.date_picker.date().toString("dd/MM/yyyy")
        dot = emp.get("dotacion") or self.combo_dotacion.currentText()
        nombre = emp.get("apellidos_nombres", "")
        legajo = emp.get("legajo")

        try:
            descubiertos = self.records_service.obtener_turnos_sin_cubrir(f_str, dotacion=dot if dot != "Todas" else None)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo consultar turnos descubiertos:\n{e}")
            return

        if not descubiertos:
            QMessageBox.information(
                self,
                "Sin Novedades Descubiertas",
                f"No se encontraron turnos descubiertos registrados en la fecha <b>{f_str}</b> (Dotación: {dot}).<br/><br/>"
                "Para cargar una nueva novedad, ingrese en la pestaña 'Novedades Operativas'.",
            )
            return

        if len(descubiertos) == 1:
            nov = descubiertos[0]
            resp = QMessageBox.question(
                self,
                "Asignar Relevo de Ciclo",
                f"¿Asignar a <b>{nombre}</b> (Leg. {legajo}) como relevo de la novedad #{nov['id']}?<br/><br/>"
                f"<b>Titular:</b> {nov.get('apellidos_nombres')} (Leg. {nov.get('legajo')})<br/>"
                f"<b>Turno:</b> {nov.get('turnos')} | <b>Motivo:</b> {nov.get('novedad')}",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes,
            )
            if resp == QMessageBox.Yes:
                try:
                    self.records_service.asignar_reemplazo_rapido(
                        novedad_id=nov["id"],
                        legajo_reemplazo=int(legajo),
                        fecha_inicio=f_str,
                        fecha_fin=nov.get("fecha_fin"),
                        user_id=self.usuario_actual.get("id"),
                        windows_user=self.usuario_actual.get("username"),
                    )
                    self.consultar()
                    QMessageBox.information(self, "Asignación Exitosa", f"Se asignó a {nombre} al relevo #{nov['id']}.")
                except Exception as e:
                    QMessageBox.critical(self, "Error", f"No se pudo asignar el relevo:\n{e}")
        else:
            dlg = QDialog(self)
            dlg.setWindowTitle("Seleccionar Novedad a Cubrir")
            dlg.resize(520, 320)
            lay = QVBoxLayout(dlg)
            lay.addWidget(QLabel(f"Seleccione la novedad a cubrir por <b>{nombre}</b>:"))
            lw = QListWidget()
            for nov in descubiertos:
                item_txt = f"#{nov['id']} | {nov.get('apellidos_nombres')} (Leg. {nov.get('legajo')}) - Turno {nov.get('turnos')} - {nov.get('novedad')}"
                l_item = QListWidgetItem(item_txt)
                l_item.setData(Qt.UserRole, nov)
                lw.addItem(l_item)
            lw.setCurrentRow(0)
            lay.addWidget(lw)
            bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
            bb.accepted.connect(dlg.accept)
            bb.rejected.connect(dlg.reject)
            lay.addWidget(bb)

            if dlg.exec():
                sel = lw.currentItem()
                if sel:
                    nov_sel = sel.data(Qt.UserRole)
                    try:
                        self.records_service.asignar_reemplazo_rapido(
                            novedad_id=nov_sel["id"],
                            legajo_reemplazo=int(legajo),
                            fecha_inicio=f_str,
                            fecha_fin=nov_sel.get("fecha_fin"),
                            user_id=self.usuario_actual.get("id"),
                            windows_user=self.usuario_actual.get("username"),
                        )
                        self.consultar()
                        QMessageBox.information(self, "Asignación Exitosa", f"Se asignó a {nombre} al relevo #{nov_sel['id']}.")
                    except Exception as e:
                        QMessageBox.critical(self, "Error", f"No se pudo asignar el relevo:\n{e}")
