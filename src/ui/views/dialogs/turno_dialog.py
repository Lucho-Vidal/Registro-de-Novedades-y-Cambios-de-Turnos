"""Diálogo modal para el alta y edición integral de un Turno de Itinerario y sus Vueltas."""

from typing import Any, Dict, List, Optional
from PySide6.QtCore import QTime, Qt
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDialog,
    QFormLayout,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QTableWidget,
    QTableWidgetItem,
    QTimeEdit,
    QVBoxLayout,
)

from src.core.logger import get_logger
from src.services.records_service import RecordsService

logger = get_logger("TurnoDialog")


class VueltaItemDialog(QDialog):
    """Mini modal para agregar o editar un tramo/vuelta de tren del turno."""

    def __init__(self, vuelta_existente: Optional[Dict[str, Any]] = None, parent=None):
        super().__init__(parent)
        self.vuelta_existente = vuelta_existente or {}
        self.setWindowTitle("Editar Vuelta de Tren" if vuelta_existente else "Agregar Vuelta de Tren")
        self.resize(380, 300)
        self.setWindowFlags((self.windowFlags() & ~Qt.WindowContextHelpButtonHint) | Qt.WindowCloseButtonHint)

        self._init_ui()

    def _init_ui(self) -> None:
        layout = QVBoxLayout(self)
        form = QFormLayout()

        self.txt_tren = QLineEdit(str(self.vuelta_existente.get("nro_tren") or ""))
        self.txt_tren.setPlaceholderText("Ej: 2301, 105, LIVIANO")
        form.addRow("N° de Tren:", self.txt_tren)

        self.txt_ref = QLineEdit(str(self.vuelta_existente.get("ref") or ""))
        self.txt_ref.setPlaceholderText("Ej: A, B, C...")
        form.addRow("Referencia (Ref):", self.txt_ref)

        self.txt_origen = QLineEdit(str(self.vuelta_existente.get("origen") or ""))
        self.txt_origen.setPlaceholderText("Ej: PLAZA, K.4, TEMPERLEY")
        form.addRow("Estación Origen:", self.txt_origen)

        self.time_salida = QTimeEdit()
        self.time_salida.setDisplayFormat("HH:mm")
        s_val = str(self.vuelta_existente.get("hora_salida") or "").strip()
        if s_val:
            qt_s = QTime.fromString(s_val, "HH:mm")
            if qt_s.isValid():
                self.time_salida.setTime(qt_s)
        form.addRow("Hora Salida:", self.time_salida)

        self.txt_destino = QLineEdit(str(self.vuelta_existente.get("destino") or ""))
        self.txt_destino.setPlaceholderText("Ej: LA PLATA, GLEW, EZEIZA")
        form.addRow("Estación Destino:", self.txt_destino)

        self.time_llegada = QTimeEdit()
        self.time_llegada.setDisplayFormat("HH:mm")
        l_val = str(self.vuelta_existente.get("hora_llegada") or "").strip()
        if l_val:
            qt_l = QTime.fromString(l_val, "HH:mm")
            if qt_l.isValid():
                self.time_llegada.setTime(qt_l)
        form.addRow("Hora Llegada:", self.time_llegada)

        self.txt_obs = QLineEdit(str(self.vuelta_existente.get("observaciones_tramo") or ""))
        form.addRow("Observaciones Tramo:", self.txt_obs)

        layout.addLayout(form)

        btn_box = QHBoxLayout()
        btn_box.addStretch()
        btn_cancel = QPushButton("Cancelar")
        btn_cancel.setObjectName("btnSecondary")
        btn_cancel.clicked.connect(self.reject)
        btn_save = QPushButton("Aceptar")
        btn_save.clicked.connect(self._on_aceptar)
        btn_box.addWidget(btn_cancel)
        btn_box.addWidget(btn_save)
        layout.addLayout(btn_box)

    def _on_aceptar(self) -> None:
        tren = self.txt_tren.text().strip()
        if not tren:
            QMessageBox.warning(self, "Validación", "El número o denominación del tren es obligatorio.")
            self.txt_tren.setFocus()
            return
        self.accept()

    def get_data(self) -> Dict[str, Any]:
        return {
            "nro_tren": self.txt_tren.text().strip(),
            "ref": self.txt_ref.text().strip(),
            "origen": self.txt_origen.text().strip(),
            "hora_salida": self.time_salida.time().toString("HH:mm"),
            "destino": self.txt_destino.text().strip(),
            "hora_llegada": self.time_llegada.time().toString("HH:mm"),
            "observaciones_tramo": self.txt_obs.text().strip(),
        }


class TurnoDialog(QDialog):
    """Formulario modal para crear o editar un turno y su diagrama de vueltas."""

    def __init__(
        self,
        records_service: RecordsService,
        turno_id: Optional[int] = None,
        usuario_actual: Optional[Dict[str, Any]] = None,
        parent=None,
    ):
        super().__init__(parent)
        self.records_service = records_service
        self.turno_id = turno_id
        self.usuario_actual = usuario_actual or {}

        self.setWindowTitle("Modificar Turno de Itinerario" if turno_id else "Nuevo Turno de Itinerario")
        self.resize(860, 640)
        self.setWindowFlags((self.windowFlags() & ~Qt.WindowContextHelpButtonHint) | Qt.WindowCloseButtonHint)

        self._vueltas: List[Dict[str, Any]] = []
        self._init_ui()

        if self.turno_id:
            self._cargar_datos_existentes()

    def _init_ui(self) -> None:
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(14, 14, 14, 14)
        main_layout.setSpacing(10)

        # 1. Datos Principales del Turno
        grp_datos = QGroupBox("1. Datos Generales del Turno")
        lay_datos = QGridLayout(grp_datos)

        lay_datos.addWidget(QLabel("Código Turno (*):"), 0, 0)
        self.txt_codigo = QLineEdit()
        self.txt_codigo.setPlaceholderText("Ej: 401.1, 401, PROG.1CK5")
        self.txt_codigo.textChanged.connect(self._auto_generar_base)
        lay_datos.addWidget(self.txt_codigo, 0, 1)

        lay_datos.addWidget(QLabel("Turno Base:"), 0, 2)
        self.txt_base = QLineEdit()
        self.txt_base.setPlaceholderText("Ej: 401")
        lay_datos.addWidget(self.txt_base, 0, 3)

        lay_datos.addWidget(QLabel("Subíndice:"), 1, 0)
        self.spin_subindice = QSpinBox()
        self.spin_subindice.setRange(0, 99)
        self.spin_subindice.setSpecialValueText("Ninguno")
        lay_datos.addWidget(self.spin_subindice, 1, 1)

        lay_datos.addWidget(QLabel("Tipo de Turno:"), 1, 2)
        self.combo_tipo = QComboBox()
        self.combo_tipo.addItems(["DIAGRAMA_6D", "CICLO", "ESPECIAL", "PROG_SEMANAL", "DISPONIBLE"])
        lay_datos.addWidget(self.combo_tipo, 1, 3)

        lay_datos.addWidget(QLabel("Depósito / Dotación:"), 2, 0)
        self.combo_deposito = QComboBox()
        self.combo_deposito.setEditable(True)
        dots = [d["nombre"] for d in self.records_service.listar_dotaciones(incluir_inactivos=False)]
        self.combo_deposito.addItems(dots or ["PC", "TY", "LL"])
        lay_datos.addWidget(self.combo_deposito, 2, 1)

        lay_datos.addWidget(QLabel("Itinerario Vigente:"), 2, 2)
        self.txt_itinerario = QLineEdit("45")
        lay_datos.addWidget(self.txt_itinerario, 2, 3)

        lay_datos.addWidget(QLabel("Toma de Servicio:"), 3, 0)
        self.time_toma = QTimeEdit()
        self.time_toma.setDisplayFormat("HH:mm")
        self.time_toma.setTime(QTime(4, 0))
        self.time_toma.timeChanged.connect(self._recalcular_duracion)
        lay_datos.addWidget(self.time_toma, 3, 1)

        lay_datos.addWidget(QLabel("Deje de Servicio:"), 3, 2)
        self.time_deje = QTimeEdit()
        self.time_deje.setDisplayFormat("HH:mm")
        self.time_deje.setTime(QTime(10, 0))
        self.time_deje.timeChanged.connect(self._recalcular_duracion)
        lay_datos.addWidget(self.time_deje, 3, 3)

        lay_datos.addWidget(QLabel("Duración Jornada:"), 4, 0)
        self.txt_duracion = QLineEdit("06:00")
        lay_datos.addWidget(self.txt_duracion, 4, 1)

        self.chk_activo = QCheckBox("Turno Activo")
        self.chk_activo.setChecked(True)
        lay_datos.addWidget(self.chk_activo, 4, 2)

        lay_datos.addWidget(QLabel("Observaciones:"), 5, 0)
        self.txt_observaciones = QLineEdit()
        self.txt_observaciones.setPlaceholderText("Observaciones operativas del turno")
        lay_datos.addWidget(self.txt_observaciones, 5, 1, 1, 3)

        main_layout.addWidget(grp_datos)

        # 2. Vueltas del Turno
        grp_vueltas = QGroupBox("2. Diagrama de Vueltas / Trenes del Turno")
        lay_vueltas = QVBoxLayout(grp_vueltas)

        bar_v = QHBoxLayout()
        self.btn_add_vuelta = QPushButton("➕ Agregar Vuelta")
        self.btn_add_vuelta.setObjectName("btnSecondary")
        self.btn_add_vuelta.clicked.connect(self._agregar_vuelta)
        bar_v.addWidget(self.btn_add_vuelta)

        self.btn_edit_vuelta = QPushButton("✏️ Modificar Vuelta")
        self.btn_edit_vuelta.setObjectName("btnSecondary")
        self.btn_edit_vuelta.clicked.connect(self._editar_vuelta)
        bar_v.addWidget(self.btn_edit_vuelta)

        self.btn_del_vuelta = QPushButton("🗑️ Quitar Vuelta")
        self.btn_del_vuelta.setObjectName("btnDanger")
        self.btn_del_vuelta.clicked.connect(self._quitar_vuelta)
        bar_v.addWidget(self.btn_del_vuelta)

        bar_v.addStretch()

        self.btn_up_vuelta = QPushButton("⬆️ Subir")
        self.btn_up_vuelta.setObjectName("btnSecondary")
        self.btn_up_vuelta.clicked.connect(lambda: self._mover_vuelta(-1))
        bar_v.addWidget(self.btn_up_vuelta)

        self.btn_down_vuelta = QPushButton("⬇️ Bajar")
        self.btn_down_vuelta.setObjectName("btnSecondary")
        self.btn_down_vuelta.clicked.connect(lambda: self._mover_vuelta(1))
        bar_v.addWidget(self.btn_down_vuelta)

        lay_vueltas.addLayout(bar_v)

        self.table_vueltas = QTableWidget(0, 8)
        self.table_vueltas.setHorizontalHeaderLabels([
            "VUELTA", "REF", "N° TREN", "ORIGEN", "SALIDA", "DESTINO", "LLEGADA", "OBSERVACIONES",
        ])
        self.table_vueltas.setSelectionBehavior(QTableWidget.SelectRows)
        self.table_vueltas.setSelectionMode(QTableWidget.SingleSelection)
        self.table_vueltas.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table_vueltas.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table_vueltas.setColumnWidth(0, 70)
        self.table_vueltas.setColumnWidth(1, 60)
        self.table_vueltas.setColumnWidth(2, 90)
        self.table_vueltas.setColumnWidth(3, 120)
        self.table_vueltas.setColumnWidth(4, 75)
        self.table_vueltas.setColumnWidth(5, 120)
        self.table_vueltas.setColumnWidth(6, 75)
        self.table_vueltas.horizontalHeader().setSectionResizeMode(7, QHeaderView.Stretch)
        self.table_vueltas.doubleClicked.connect(lambda _: self._editar_vuelta())
        lay_vueltas.addWidget(self.table_vueltas)

        main_layout.addWidget(grp_vueltas)

        # 3. Botones Guardar / Cancelar
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        btn_cancel = QPushButton("Cancelar")
        btn_cancel.setObjectName("btnSecondary")
        btn_cancel.clicked.connect(self.reject)
        btn_layout.addWidget(btn_cancel)

        btn_save = QPushButton("Guardar Turno e Itinerario")
        btn_save.setObjectName("btnPrimary")
        btn_save.clicked.connect(self._on_guardar)
        btn_layout.addWidget(btn_save)

        main_layout.addLayout(btn_layout)

    def _auto_generar_base(self, text: str) -> None:
        if not self.txt_base.text() or self.txt_base.text() == text.split(".")[0]:
            parts = text.strip().split(".")
            self.txt_base.setText(parts[0])
            if len(parts) > 1 and parts[1].isdigit():
                self.spin_subindice.setValue(int(parts[1]))

    def _recalcular_duracion(self) -> None:
        t_toma = self.time_toma.time()
        t_deje = self.time_deje.time()
        m_toma = t_toma.hour() * 60 + t_toma.minute()
        m_deje = t_deje.hour() * 60 + t_deje.minute()
        diff = m_deje - m_toma
        if diff < 0:
            diff += 24 * 60
        hh = diff // 60
        mm = diff % 60
        self.txt_duracion.setText(f"{hh:02d}:{mm:02d}")

    def _cargar_datos_existentes(self) -> None:
        t = self.records_service.obtener_turno_completo(self.turno_id)
        if not t:
            return

        self.txt_codigo.setText(str(t.get("codigo_turno") or ""))
        self.txt_base.setText(str(t.get("turno_base") or ""))
        if t.get("subindice") is not None:
            self.spin_subindice.setValue(int(t["subindice"]))
        else:
            self.spin_subindice.setValue(0)

        if t.get("tipo_turno"):
            idx = self.combo_tipo.findText(t["tipo_turno"])
            if idx >= 0:
                self.combo_tipo.setCurrentIndex(idx)

        if t.get("deposito_base"):
            self.combo_deposito.setCurrentText(str(t["deposito_base"]))

        if t.get("itinerario_vigente"):
            self.txt_itinerario.setText(str(t["itinerario_vigente"]))

        if t.get("toma_servicio"):
            qt = QTime.fromString(str(t["toma_servicio"]), "HH:mm")
            if qt.isValid():
                self.time_toma.setTime(qt)

        if t.get("deje_servicio"):
            qt = QTime.fromString(str(t["deje_servicio"]), "HH:mm")
            if qt.isValid():
                self.time_deje.setTime(qt)

        if t.get("duracion_jornada"):
            self.txt_duracion.setText(str(t["duracion_jornada"]))

        self.chk_activo.setChecked(bool(t.get("activo", 1)))
        self.txt_observaciones.setText(str(t.get("observaciones_turno") or ""))

        self._vueltas = list(t.get("vueltas") or [])
        self._refrescar_tabla_vueltas()

    def _refrescar_tabla_vueltas(self) -> None:
        self.table_vueltas.setRowCount(len(self._vueltas))
        for r, v in enumerate(self._vueltas):
            v["nro_vuelta"] = r + 1
            self.table_vueltas.setItem(r, 0, QTableWidgetItem(str(v["nro_vuelta"])))
            self.table_vueltas.setItem(r, 1, QTableWidgetItem(str(v.get("ref") or "-")))
            self.table_vueltas.setItem(r, 2, QTableWidgetItem(str(v.get("nro_tren") or "")))
            self.table_vueltas.setItem(r, 3, QTableWidgetItem(str(v.get("origen") or "")))
            self.table_vueltas.setItem(r, 4, QTableWidgetItem(str(v.get("hora_salida") or "")))
            self.table_vueltas.setItem(r, 5, QTableWidgetItem(str(v.get("destino") or "")))
            self.table_vueltas.setItem(r, 6, QTableWidgetItem(str(v.get("hora_llegada") or "")))
            self.table_vueltas.setItem(r, 7, QTableWidgetItem(str(v.get("observaciones_tramo") or "")))

    def _agregar_vuelta(self) -> None:
        dlg = VueltaItemDialog(parent=self)
        if dlg.exec():
            d = dlg.get_data()
            self._vueltas.append(d)
            self._refrescar_tabla_vueltas()

    def _editar_vuelta(self) -> None:
        r = self.table_vueltas.currentRow()
        if r < 0 or r >= len(self._vueltas):
            QMessageBox.information(self, "Selección", "Seleccione una vuelta de la tabla para editar.")
            return
        dlg = VueltaItemDialog(self._vueltas[r], parent=self)
        if dlg.exec():
            self._vueltas[r] = dlg.get_data()
            self._refrescar_tabla_vueltas()

    def _quitar_vuelta(self) -> None:
        r = self.table_vueltas.currentRow()
        if r < 0 or r >= len(self._vueltas):
            QMessageBox.information(self, "Selección", "Seleccione una vuelta de la tabla para quitar.")
            return
        self._vueltas.pop(r)
        self._refrescar_tabla_vueltas()

    def _mover_vuelta(self, delta: int) -> None:
        r = self.table_vueltas.currentRow()
        if r < 0 or r >= len(self._vueltas):
            return
        target = r + delta
        if 0 <= target < len(self._vueltas):
            self._vueltas[r], self._vueltas[target] = self._vueltas[target], self._vueltas[r]
            self._refrescar_tabla_vueltas()
            self.table_vueltas.selectRow(target)

    def _on_guardar(self) -> None:
        codigo = self.txt_codigo.text().strip()
        if not codigo:
            QMessageBox.warning(self, "Validación", "El código de turno es obligatorio.")
            self.txt_codigo.setFocus()
            return

        base = self.txt_base.text().strip() or codigo.split(".")[0]
        sub = self.spin_subindice.value() if self.spin_subindice.value() > 0 else None

        payload = {
            "codigo_turno": codigo,
            "turno_base": base,
            "subindice": sub,
            "tipo_turno": self.combo_tipo.currentText().strip(),
            "deposito_base": self.combo_deposito.currentText().strip() or "PC",
            "toma_servicio": self.time_toma.time().toString("HH:mm"),
            "deje_servicio": self.time_deje.time().toString("HH:mm"),
            "duracion_jornada": self.txt_duracion.text().strip(),
            "tiempo_toma": self.time_toma.time().toString("HH:mm"),
            "tiempo_deje": self.time_deje.time().toString("HH:mm"),
            "itinerario_vigente": self.txt_itinerario.text().strip() or "45",
            "observaciones_turno": self.txt_observaciones.text().strip(),
            "activo": 1 if self.chk_activo.isChecked() else 0,
        }

        u_id = self.usuario_actual.get("id")
        win_u = self.usuario_actual.get("username")

        try:
            if self.turno_id:
                self.records_service.actualizar_turno_itinerario(
                    self.turno_id,
                    payload,
                    vueltas=self._vueltas,
                    user_id=u_id,
                    windows_user=win_u,
                )
            else:
                self.records_service.crear_turno_itinerario(
                    payload,
                    vueltas=self._vueltas,
                    user_id=u_id,
                    windows_user=win_u,
                )
            self.accept()
        except Exception as e:
            QMessageBox.critical(self, "Error al Guardar", f"No se pudo guardar el turno:\n{e}")
