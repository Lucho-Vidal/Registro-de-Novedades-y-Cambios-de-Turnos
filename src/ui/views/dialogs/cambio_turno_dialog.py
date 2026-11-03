"""Diálogo modal para el alta y edición de Cambios de Turno (Permutas)."""

from typing import Any, Dict, List, Optional
from datetime import date, datetime

from PySide6.QtCore import QDate, Qt
from PySide6.QtWidgets import (
    QComboBox,
    QDateEdit,
    QDialog,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QVBoxLayout,
)

from src.core.logger import get_logger
from src.domain.validators import parsear_fecha
from src.services.records_service import RecordsService
from src.ui.views.dialogs.legajo_search_dialog import LegajoSearchDialog

logger = get_logger("CambioTurnoDialog")


class CambioTurnoDialog(QDialog):
    """Formulario para solicitar o modificar una permuta / cambio de turno."""

    def __init__(
        self,
        records_service: RecordsService,
        cambio_existente: Optional[Dict[str, Any]] = None,
        usuario_actual: Optional[Dict[str, Any]] = None,
        parent=None,
    ):
        super().__init__(parent)
        self.records_service = records_service
        self.cambio_existente = cambio_existente
        self.usuario_actual = usuario_actual or {}

        self.setWindowTitle("Modificar Cambio de Turno" if cambio_existente else "Nuevo Cambio de Turno (Permuta)")
        self.resize(840, 580)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)

        self._empleados_cache: List[Dict[str, Any]] = self.records_service.listar_empleados(incluir_inactivos=False)
        self._personal_estacion_cache: List[Dict[str, Any]] = self.records_service.listar_personal_estacion(incluir_inactivos=False)

        self._init_ui()
        if self.cambio_existente:
            self._cargar_datos_existentes()

    def _init_ui(self) -> None:
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(12)

        # 1. Agente 1 (Solicitante)
        grp_1 = QGroupBox("1. Agente Solicitante")
        lay_1 = QGridLayout(grp_1)

        lay_1.addWidget(QLabel("Legajo:"), 0, 0)
        h_leg1 = QHBoxLayout()
        self.txt_legajo_1 = QLineEdit()
        self.txt_legajo_1.setFixedWidth(100)
        self.txt_legajo_1.editingFinished.connect(lambda: self._on_legajo_changed(1))
        self.btn_buscar_1 = QPushButton("🔍 Buscar")
        self.btn_buscar_1.setObjectName("btnSecondary")
        self.btn_buscar_1.clicked.connect(lambda: self._buscar_legajo_modal(1))
        h_leg1.addWidget(self.txt_legajo_1)
        h_leg1.addWidget(self.btn_buscar_1)
        h_leg1.addStretch()
        lay_1.addLayout(h_leg1, 0, 1)

        lay_1.addWidget(QLabel("Apellidos y Nombres:"), 0, 2)
        self.txt_nombre_1 = QLineEdit()
        self.txt_nombre_1.setReadOnly(True)
        lay_1.addWidget(self.txt_nombre_1, 0, 3)

        lay_1.addWidget(QLabel("Especialidad:"), 1, 0)
        self.txt_especialidad_1 = QLineEdit()
        self.txt_especialidad_1.setReadOnly(True)
        lay_1.addWidget(self.txt_especialidad_1, 1, 1)

        lay_1.addWidget(QLabel("Dotación:"), 1, 2)
        self.txt_dotacion_1 = QLineEdit()
        self.txt_dotacion_1.setReadOnly(True)
        lay_1.addWidget(self.txt_dotacion_1, 1, 3)

        lay_1.addWidget(QLabel("Turno:"), 2, 0)
        self.txt_turnos_1 = QLineEdit()
        self.txt_turnos_1.setReadOnly(True)
        lay_1.addWidget(self.txt_turnos_1, 2, 1)

        lay_1.addWidget(QLabel("Franco:"), 2, 2)
        self.txt_franco_1 = QLineEdit()
        self.txt_franco_1.setReadOnly(True)
        lay_1.addWidget(self.txt_franco_1, 2, 3)

        main_layout.addWidget(grp_1)

        # 2. Agente 2 (Receptor)
        grp_2 = QGroupBox("2. Agente Receptor")
        lay_2 = QGridLayout(grp_2)

        lay_2.addWidget(QLabel("Legajo:"), 0, 0)
        h_leg2 = QHBoxLayout()
        self.txt_legajo_2 = QLineEdit()
        self.txt_legajo_2.setFixedWidth(100)
        self.txt_legajo_2.editingFinished.connect(lambda: self._on_legajo_changed(2))
        self.btn_buscar_2 = QPushButton("🔍 Buscar")
        self.btn_buscar_2.setObjectName("btnSecondary")
        self.btn_buscar_2.clicked.connect(lambda: self._buscar_legajo_modal(2))
        h_leg2.addWidget(self.txt_legajo_2)
        h_leg2.addWidget(self.btn_buscar_2)
        h_leg2.addStretch()
        lay_2.addLayout(h_leg2, 0, 1)

        lay_2.addWidget(QLabel("Apellidos y Nombres:"), 0, 2)
        self.txt_nombre_2 = QLineEdit()
        self.txt_nombre_2.setReadOnly(True)
        lay_2.addWidget(self.txt_nombre_2, 0, 3)

        lay_2.addWidget(QLabel("Especialidad:"), 1, 0)
        self.txt_especialidad_2 = QLineEdit()
        self.txt_especialidad_2.setReadOnly(True)
        lay_2.addWidget(self.txt_especialidad_2, 1, 1)

        lay_2.addWidget(QLabel("Dotación:"), 1, 2)
        self.txt_dotacion_2 = QLineEdit()
        self.txt_dotacion_2.setReadOnly(True)
        lay_2.addWidget(self.txt_dotacion_2, 1, 3)

        lay_2.addWidget(QLabel("Turno:"), 2, 0)
        self.txt_turnos_2 = QLineEdit()
        self.txt_turnos_2.setReadOnly(True)
        lay_2.addWidget(self.txt_turnos_2, 2, 1)

        lay_2.addWidget(QLabel("Franco:"), 2, 2)
        self.txt_franco_2 = QLineEdit()
        self.txt_franco_2.setReadOnly(True)
        lay_2.addWidget(self.txt_franco_2, 2, 3)

        main_layout.addWidget(grp_2)

        # 3. Datos del Cambio
        grp_det = QGroupBox("3. Datos de la Permuta")
        lay_det = QGridLayout(grp_det)

        lay_det.addWidget(QLabel("Fecha del Cambio:"), 0, 0)
        self.date_cambio = QDateEdit()
        self.date_cambio.setCalendarPopup(True)
        self.date_cambio.setDisplayFormat("dd/MM/yyyy")
        self.date_cambio.setDate(QDate.currentDate())
        lay_det.addWidget(self.date_cambio, 0, 1)

        lay_det.addWidget(QLabel("Referencia Estación:"), 0, 2)
        self.combo_estacion = QComboBox()
        self.combo_estacion.setEditable(True)
        self.combo_estacion.addItem("")
        for pe in self._personal_estacion_cache:
            dot_txt = f" ({pe['dotacion']})" if pe.get("dotacion") else ""
            self.combo_estacion.addItem(f"{pe['nombre']}{dot_txt}")
        lay_det.addWidget(self.combo_estacion, 0, 3)

        lay_det.addWidget(QLabel("Supervisor:"), 1, 0)
        self.txt_supervisor = QLineEdit()
        lay_det.addWidget(self.txt_supervisor, 1, 1)

        lay_det.addWidget(QLabel("Observaciones:"), 2, 0)
        self.txt_observaciones = QPlainTextEdit()
        self.txt_observaciones.setMaximumHeight(65)
        lay_det.addWidget(self.txt_observaciones, 2, 1, 1, 3)

        main_layout.addWidget(grp_det)

        # 4. Botones
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        self.btn_cancel = QPushButton("Cancelar")
        self.btn_cancel.setObjectName("btnSecondary")
        self.btn_cancel.clicked.connect(self.reject)
        btn_layout.addWidget(self.btn_cancel)

        self.btn_guardar = QPushButton("Guardar Cambio de Turno")
        self.btn_guardar.clicked.connect(self._guardar)
        btn_layout.addWidget(self.btn_guardar)

        main_layout.addLayout(btn_layout)

    def _buscar_legajo_modal(self, interviniente: int) -> None:
        dlg = LegajoSearchDialog(self._empleados_cache, parent=self)
        if dlg.exec():
            emp = dlg.selected_empleado
            if emp:
                self._cargar_empleado(interviniente, emp)

    def _on_legajo_changed(self, interviniente: int) -> None:
        txt = self.txt_legajo_1 if interviniente == 1 else self.txt_legajo_2
        val_str = txt.text().strip()
        if not val_str:
            return
        try:
            leg_int = int(val_str)
            emp = next((e for e in self._empleados_cache if int(e["legajo"]) == leg_int), None)
            if emp:
                self._cargar_empleado(interviniente, emp)
            else:
                QMessageBox.warning(self, "Búsqueda", f"No se encontró ningún empleado con legajo {leg_int}.")
        except ValueError:
            pass

    def _cargar_empleado(self, interviniente: int, emp: Dict[str, Any]) -> None:
        if interviniente == 1:
            self.txt_legajo_1.setText(str(emp["legajo"]))
            self.txt_nombre_1.setText(str(emp["apellidos_nombres"]))
            self.txt_especialidad_1.setText(str(emp["especialidad"]))
            self.txt_dotacion_1.setText(str(emp["dotacion"]))
            self.txt_turnos_1.setText(str(emp["turnos"]))
            self.txt_franco_1.setText(str(emp["franco"]))
        else:
            self.txt_legajo_2.setText(str(emp["legajo"]))
            self.txt_nombre_2.setText(str(emp["apellidos_nombres"]))
            self.txt_especialidad_2.setText(str(emp["especialidad"]))
            self.txt_dotacion_2.setText(str(emp["dotacion"]))
            self.txt_turnos_2.setText(str(emp["turnos"]))
            self.txt_franco_2.setText(str(emp["franco"]))

    def _cargar_datos_existentes(self) -> None:
        c = self.cambio_existente or {}
        if c.get("legajo_1"):
            self.txt_legajo_1.setText(str(c["legajo_1"]))
            self._on_legajo_changed(1)
        if c.get("legajo_2"):
            self.txt_legajo_2.setText(str(c["legajo_2"]))
            self._on_legajo_changed(2)

        if c.get("fecha_cambio"):
            dt = parsear_fecha(c["fecha_cambio"])
            if dt:
                self.date_cambio.setDate(QDate(dt.year, dt.month, dt.day))

        if c.get("referencia_estacion"):
            self.combo_estacion.setEditText(str(c["referencia_estacion"]))
        if c.get("supervisor"):
            self.txt_supervisor.setText(str(c["supervisor"]))
        if c.get("observaciones"):
            self.txt_observaciones.setPlainText(str(c["observaciones"]))

    def _guardar(self) -> None:
        leg1 = self.txt_legajo_1.text().strip()
        leg2 = self.txt_legajo_2.text().strip()

        if not leg1 or not leg2:
            QMessageBox.warning(self, "Validación", "Ambos legajos son obligatorios.")
            return

        if leg1 == leg2:
            QMessageBox.warning(self, "Validación", "No puede realizar un cambio de turno consigo mismo.")
            return

        payload = {
            "legajo_1": int(leg1),
            "apellidos_nombres_1": self.txt_nombre_1.text().strip(),
            "especialidad_1": self.txt_especialidad_1.text().strip(),
            "dotacion_1": self.txt_dotacion_1.text().strip(),
            "turnos_1": self.txt_turnos_1.text().strip(),
            "franco_1": self.txt_franco_1.text().strip(),
            "legajo_2": int(leg2),
            "apellidos_nombres_2": self.txt_nombre_2.text().strip(),
            "especialidad_2": self.txt_especialidad_2.text().strip(),
            "dotacion_2": self.txt_dotacion_2.text().strip(),
            "turnos_2": self.txt_turnos_2.text().strip(),
            "franco_2": self.txt_franco_2.text().strip(),
            "fecha_cambio": self.date_cambio.date().toString("dd/MM/yyyy"),
            "referencia_estacion": self.combo_estacion.currentText().strip(),
            "supervisor": self.txt_supervisor.text().strip(),
            "observaciones": self.txt_observaciones.toPlainText().strip(),
        }

        try:
            if self.cambio_existente:
                self.records_service.actualizar_cambio(
                    self.cambio_existente["id"],
                    payload,
                    user_id=self.usuario_actual.get("id"),
                    windows_user=self.usuario_actual.get("username"),
                    bypass_window=True,
                )
            else:
                self.records_service.crear_cambio(
                    payload,
                    user_id=self.usuario_actual.get("id"),
                    windows_user=self.usuario_actual.get("username"),
                )
            self.accept()
        except Exception as e:
            QMessageBox.critical(self, "Error al Guardar", f"No se pudo guardar el cambio de turno:\n{e}")
