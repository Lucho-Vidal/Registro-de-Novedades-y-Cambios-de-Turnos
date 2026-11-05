"""Panel y Diálogo de Administración de Turnos de Itinerario con roles y permisos (CRONO-PCT Standard)."""

from typing import Any, Dict, List, Optional
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QTableView,
    QVBoxLayout,
    QWidget,
)

from src.core.logger import get_logger
from src.services.records_service import RecordsService
from src.ui.models.base_table_model import BaseTableModel
from src.ui.views.dialogs.turno_dialog import TurnoDialog

logger = get_logger("AdminTurnos")


class BaseMaximizedDialog(QDialog):
    """Clase base para diálogos maximizados con botón de cierre funcional."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowFlags(Qt.Window | Qt.WindowTitleHint | Qt.WindowMinMaxButtonsHint | Qt.WindowCloseButtonHint)

    def closeEvent(self, event) -> None:
        self.reject()
        event.accept()

    def showEvent(self, event) -> None:
        super().showEvent(event)
        self.showMaximized()


class AdminTurnosWidget(QWidget):
    """Panel de administración y CRUD completo de turnos de itinerario con permisos y roles."""

    def __init__(
        self,
        records_service: RecordsService,
        usuario_actual: Optional[Dict[str, Any]] = None,
        tiene_permiso_fn=None,
        parent=None,
    ):
        super().__init__(parent)
        self.records_service = records_service
        self.usuario_actual = usuario_actual or {}
        self.tiene_permiso_fn = tiene_permiso_fn

        self._turnos: List[Dict[str, Any]] = []
        self._init_ui()
        self.recargar()

    def _can(self, permission_code: str) -> bool:
        if self.tiene_permiso_fn is None:
            return True
        return bool(self.tiene_permiso_fn(permission_code))

    def _init_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(6, 6, 6, 6)
        layout.setSpacing(8)

        # Barra de Filtros y Acciones
        bar_top = QHBoxLayout()

        self.btn_nuevo = QPushButton("➕ Nuevo Turno")
        self.btn_nuevo.setObjectName("btnPrimary")
        self.btn_nuevo.clicked.connect(self._nuevo_turno)
        bar_top.addWidget(self.btn_nuevo)

        self.btn_modificar = QPushButton("✏️ Modificar Turno")
        self.btn_modificar.setObjectName("btnSecondary")
        self.btn_modificar.clicked.connect(self._modificar_turno)
        bar_top.addWidget(self.btn_modificar)

        self.btn_toggle = QPushButton("🔄 Activar / Desactivar")
        self.btn_toggle.setObjectName("btnSecondary")
        self.btn_toggle.clicked.connect(self._toggle_estado)
        bar_top.addWidget(self.btn_toggle)

        self.btn_eliminar = QPushButton("🗑️ Eliminar Definitivo")
        self.btn_eliminar.setObjectName("btnDanger")
        self.btn_eliminar.clicked.connect(self._eliminar_definitivo)
        bar_top.addWidget(self.btn_eliminar)

        bar_top.addStretch()

        # Filtros
        bar_top.addWidget(QLabel("Depósito:"))
        self.combo_deposito = QComboBox()
        self.combo_deposito.addItem("Todos")
        dots = [d["nombre"] for d in self.records_service.listar_dotaciones(incluir_inactivos=False)]
        self.combo_deposito.addItems(dots or ["PC", "TY", "LL"])
        self.combo_deposito.currentIndexChanged.connect(lambda _: self.recargar())
        bar_top.addWidget(self.combo_deposito)

        bar_top.addWidget(QLabel("Buscar:"))
        self.txt_buscar = QLineEdit()
        self.txt_buscar.setPlaceholderText("Código, base, observaciones...")
        self.txt_buscar.textChanged.connect(lambda _: self.recargar())
        bar_top.addWidget(self.txt_buscar)

        self.chk_inactivos = QCheckBox("Incluir Inactivos")
        self.chk_inactivos.setChecked(True)
        self.chk_inactivos.stateChanged.connect(lambda _: self.recargar())
        bar_top.addWidget(self.chk_inactivos)

        layout.addLayout(bar_top)

        # Tabla de Turnos
        headers = [
            "ID", "CÓDIGO TURNO", "TURNO BASE", "SUB", "TIPO TURNO", "DEPÓSITO",
            "TOMA", "DEJE", "DURACIÓN", "VUELTAS", "ITIN.", "ESTADO",
        ]
        keys = [
            "id", "codigo_turno", "turno_base", "subindice", "tipo_turno", "deposito_base",
            "toma_servicio", "deje_servicio", "duracion_jornada", "cant_vueltas", "itinerario_vigente", "estado_str",
        ]

        self.model = BaseTableModel(headers, keys, parent=self)
        self.table = QTableView()
        self.table.setModel(self.model)
        self.table.setSelectionBehavior(QTableView.SelectRows)
        self.table.setSelectionMode(QTableView.SingleSelection)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table.setColumnWidth(0, 55)   # ID
        self.table.setColumnWidth(1, 110)  # CÓDIGO TURNO
        self.table.setColumnWidth(2, 95)   # TURNO BASE
        self.table.setColumnWidth(3, 50)   # SUB
        self.table.setColumnWidth(4, 120)  # TIPO TURNO
        self.table.setColumnWidth(5, 80)   # DEPÓSITO
        self.table.setColumnWidth(6, 70)   # TOMA
        self.table.setColumnWidth(7, 70)   # DEJE
        self.table.setColumnWidth(8, 80)   # DURACIÓN
        self.table.setColumnWidth(9, 75)   # VUELTAS
        self.table.setColumnWidth(10, 60)  # ITIN
        self.table.horizontalHeader().setSectionResizeMode(11, QHeaderView.Stretch) # ESTADO
        self.table.doubleClicked.connect(lambda _: self._modificar_turno())
        layout.addWidget(self.table)

        # Pie con conteo
        self.lbl_conteo = QLabel("Mostrando 0 turnos")
        self.lbl_conteo.setStyleSheet("color: #6B7280; font-size: 8.5pt;")
        layout.addWidget(self.lbl_conteo)

    def recargar(self) -> None:
        criterio = self.txt_buscar.text().strip()
        dep = self.combo_deposito.currentText()
        inactivos = self.chk_inactivos.isChecked()

        self._turnos = self.records_service.listar_turnos_admin(
            criterio=criterio,
            deposito=dep,
            incluir_inactivos=inactivos,
        )
        for t in self._turnos:
            t["estado_str"] = "Activo" if t.get("activo") else "Inactivo"
            if t.get("subindice") is None:
                t["subindice"] = "-"
        self.model.set_data(self._turnos)
        self.lbl_conteo.setText(f"Mostrando {len(self._turnos)} turnos de itinerario")

    def _obtener_seleccionado(self) -> Optional[Dict[str, Any]]:
        indexes = self.table.selectionModel().selectedRows()
        if not indexes:
            return None
        return self.model.get_row(indexes[0].row())

    def _nuevo_turno(self) -> None:
        if not (self._can("turnos.crear") or self._can("turnos.administrar")):
            QMessageBox.warning(self, "Permiso Denegado", "No posee permisos para dar de alta turnos de itinerario.")
            return

        dlg = TurnoDialog(
            self.records_service,
            usuario_actual=self.usuario_actual,
            parent=self,
        )
        if dlg.exec():
            self.recargar()

    def _modificar_turno(self) -> None:
        if not (self._can("turnos.editar") or self._can("turnos.administrar")):
            QMessageBox.warning(self, "Permiso Denegado", "No posee permisos para modificar turnos de itinerario.")
            return

        item = self._obtener_seleccionado()
        if not item:
            QMessageBox.information(self, "Selección", "Seleccione un turno de la tabla para modificar.")
            return

        dlg = TurnoDialog(
            self.records_service,
            turno_id=item["id"],
            usuario_actual=self.usuario_actual,
            parent=self,
        )
        if dlg.exec():
            self.recargar()

    def _toggle_estado(self) -> None:
        if not (self._can("turnos.editar") or self._can("turnos.administrar")):
            QMessageBox.warning(self, "Permiso Denegado", "No posee permisos para cambiar el estado del turno.")
            return

        item = self._obtener_seleccionado()
        if not item:
            QMessageBox.information(self, "Selección", "Seleccione un turno de la tabla.")
            return

        nuevo_act = 0 if item.get("activo") else 1
        u_id = self.usuario_actual.get("id")
        win_u = self.usuario_actual.get("username")
        try:
            self.records_service.cambiar_estado_turno(
                item["id"],
                activo=nuevo_act,
                user_id=u_id,
                windows_user=win_u,
            )
            self.recargar()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo cambiar el estado del turno:\n{e}")

    def _eliminar_definitivo(self) -> None:
        if not (self._can("turnos.eliminar") or self._can("turnos.administrar")):
            QMessageBox.warning(self, "Permiso Denegado", "No posee permisos para eliminar turnos de itinerario.")
            return

        item = self._obtener_seleccionado()
        if not item:
            QMessageBox.information(self, "Selección", "Seleccione un turno de la tabla para eliminar.")
            return

        resp = QMessageBox.warning(
            self,
            "Confirmación de Eliminación",
            f"¿Está seguro de eliminar permanentemente el turno '{item.get('codigo_turno')}' y todas sus vueltas asociadas?\n\nEsta acción no se puede deshacer.",
            QMessageBox.Yes | QMessageBox.No,
        )
        if resp != QMessageBox.Yes:
            return

        u_id = self.usuario_actual.get("id")
        win_u = self.usuario_actual.get("username")
        try:
            self.records_service.eliminar_turno_definitivo(
                item["id"],
                user_id=u_id,
                windows_user=win_u,
            )
            self.recargar()
        except Exception as e:
            QMessageBox.critical(self, "Error al Eliminar", f"No se pudo eliminar el turno:\n{e}")


class AdminTurnosDialog(BaseMaximizedDialog):
    """Diálogo modal maximizado para administración de turnos con permisos y roles."""

    def __init__(
        self,
        records_service: RecordsService,
        usuario_actual: Optional[Dict[str, Any]] = None,
        tiene_permiso_fn=None,
        parent=None,
    ):
        super().__init__(parent)
        self.setWindowTitle("Administración de Turnos de Itinerario")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)

        self.widget = AdminTurnosWidget(
            records_service,
            usuario_actual=usuario_actual,
            tiene_permiso_fn=tiene_permiso_fn,
            parent=self,
        )
        layout.addWidget(self.widget)

        btn_cerrar = QPushButton("Cerrar")
        btn_cerrar.setObjectName("btnSecondary")
        btn_cerrar.clicked.connect(self.accept)
        layout.addWidget(btn_cerrar, alignment=Qt.AlignRight)

    def recargar(self) -> None:
        self.widget.recargar()
