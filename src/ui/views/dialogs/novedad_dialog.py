"""Diálogo modal para el alta y edición integral de novedades operativas.

Soporta autocompletado por legajo, duración por tipo de novedad,
referencia de estación filtrable y gestión dinámica de tramos de reemplazo (Ciclo).
"""

from typing import Any, Dict, List, Optional
from datetime import date, datetime, timedelta

from PySide6.QtCore import QDate, Qt
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDateEdit,
    QDialog,
    QFrame,
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
    QVBoxLayout,
    QWidget,
)

from src.core.logger import get_logger
from src.core.config import obtener_usuario_windows
from src.domain.validators import es_turno_ciclo, parsear_fecha
from src.services.records_service import RecordsService
from src.ui.views.dialogs.legajo_search_dialog import LegajoSearchDialog

logger = get_logger("NovedadDialog")


class TramoEditDialog(QDialog):
    """Diálogo modal secundario para agregar o modificar un tramo de relevo de ciclo."""

    def __init__(
        self,
        candidatos: List[Dict[str, Any]],
        fecha_inicio_def: date,
        fecha_fin_def: Optional[date] = None,
        tramo_existente: Optional[Dict[str, Any]] = None,
        parent=None,
    ):
        super().__init__(parent)
        self.setWindowTitle("Asignar Tramo de Cobertura (Ciclo)")
        self.setMinimumWidth(480)
        self.candidatos = candidatos
        self.tramo_result: Optional[Dict[str, Any]] = None

        layout = QVBoxLayout(self)
        form_layout = QGridLayout()

        # Selector de candidato
        form_layout.addWidget(QLabel("Personal de Ciclo disponible:"), 0, 0)
        self.combo_candidato = QComboBox()
        for c in candidatos:
            self.combo_candidato.addItem(
                f"{c['legajo']} - {c['apellidos_nombres']} ({c['turnos']})",
                userData=c,
            )
        form_layout.addWidget(self.combo_candidato, 0, 1)

        # Fechas del tramo
        form_layout.addWidget(QLabel("Fecha de inicio:"), 1, 0)
        self.date_ini = QDateEdit()
        self.date_ini.setCalendarPopup(True)
        self.date_ini.setDisplayFormat("dd/MM/yyyy")
        self.date_ini.setDate(QDate(fecha_inicio_def.year, fecha_inicio_def.month, fecha_inicio_def.day))
        form_layout.addWidget(self.date_ini, 1, 1)

        form_layout.addWidget(QLabel("Fecha de fin:"), 2, 0)
        self.chk_fin = QCheckBox("Definir fecha de fin")
        self.date_fin = QDateEdit()
        self.date_fin.setCalendarPopup(True)
        self.date_fin.setDisplayFormat("dd/MM/yyyy")
        if fecha_fin_def:
            self.date_fin.setDate(QDate(fecha_fin_def.year, fecha_fin_def.month, fecha_fin_def.day))
            self.chk_fin.setChecked(True)
        else:
            self.date_fin.setDate(self.date_ini.date())
            self.date_fin.setEnabled(False)
            self.chk_fin.setChecked(False)

        self.chk_fin.toggled.connect(self.date_fin.setEnabled)
        h_fin = QHBoxLayout()
        h_fin.addWidget(self.chk_fin)
        h_fin.addWidget(self.date_fin)
        form_layout.addLayout(h_fin, 2, 1)

        layout.addLayout(form_layout)

        # Si se pasa tramo existente para edición, preseleccionarlo
        if tramo_existente:
            leg = tramo_existente.get("legajo_reemplazo") or tramo_existente.get("legajo")
            for i in range(self.combo_candidato.count()):
                c_data = self.combo_candidato.itemData(i)
                if c_data and int(c_data["legajo"]) == int(leg):
                    self.combo_candidato.setCurrentIndex(i)
                    break
            if tramo_existente.get("fecha_inicio"):
                dt_i = parsear_fecha(tramo_existente["fecha_inicio"])
                if dt_i:
                    self.date_ini.setDate(QDate(dt_i.year, dt_i.month, dt_i.day))
            if tramo_existente.get("fecha_fin") and tramo_existente["fecha_fin"] != "-":
                dt_f = parsear_fecha(tramo_existente["fecha_fin"])
                if dt_f:
                    self.date_fin.setDate(QDate(dt_f.year, dt_f.month, dt_f.day))
                    self.chk_fin.setChecked(True)
            else:
                self.chk_fin.setChecked(False)

        # Botones
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        btn_cancel = QPushButton("Cancelar")
        btn_cancel.setObjectName("btnSecondary")
        btn_cancel.clicked.connect(self.reject)
        btn_ok = QPushButton("Aceptar")
        btn_ok.clicked.connect(self._on_accept)
        btn_layout.addWidget(btn_cancel)
        btn_layout.addWidget(btn_ok)
        layout.addLayout(btn_layout)

    def _on_accept(self) -> None:
        cand = self.combo_candidato.currentData()
        if not cand:
            QMessageBox.warning(self, "Validación", "Debe seleccionar un candidato de ciclo.")
            return

        f_ini_py = self.date_ini.date().toPython()
        f_fin_py = self.date_fin.date().toPython() if self.chk_fin.isChecked() else None

        if f_fin_py and f_fin_py < f_ini_py:
            QMessageBox.warning(self, "Validación", "La fecha de fin no puede ser anterior a la de inicio.")
            return

        self.tramo_result = {
            "legajo_reemplazo": cand["legajo"],
            "apellidos_nombres_reemplazo": cand["apellidos_nombres"],
            "turnos_reemplazo": cand["turnos"],
            "franco_reemplazo": cand["franco"],
            "fecha_inicio": f_ini_py.strftime("%d/%m/%Y"),
            "fecha_fin": f_fin_py.strftime("%d/%m/%Y") if f_fin_py else None,
        }
        self.accept()


class NovedadDialog(QDialog):
    """Formulario modal para crear o editar una novedad operativa."""

    def __init__(
        self,
        records_service: RecordsService,
        novedad_existente: Optional[Dict[str, Any]] = None,
        tramos_existentes: Optional[List[Dict[str, Any]]] = None,
        modo_restringido: bool = False,
        solo_lectura: bool = False,
        usuario_actual: Optional[Dict[str, Any]] = None,
        parent=None,
    ):
        super().__init__(parent)
        self.records_service = records_service
        self.novedad_existente = novedad_existente
        self.tramos: List[Dict[str, Any]] = list(tramos_existentes or [])
        self.modo_restringido = modo_restringido
        self.solo_lectura = solo_lectura
        self.usuario_actual = usuario_actual or {}

        self.setWindowTitle("Modificar Novedad" if novedad_existente else "Nueva Novedad Operativa")
        self.resize(840, 680)
        self.setWindowFlags((self.windowFlags() & ~Qt.WindowContextHelpButtonHint) | Qt.WindowCloseButtonHint)

        self._empleados_cache: List[Dict[str, Any]] = self.records_service.listar_empleados(incluir_inactivos=False)
        self._tipos_cache: List[Dict[str, Any]] = self.records_service.listar_tipos(incluir_inactivos=False)
        self._personal_estacion_cache: List[Dict[str, Any]] = self.records_service.listar_personal_estacion(incluir_inactivos=False)

        self._init_ui()
        if self.novedad_existente:
            self._cargar_datos_existentes()

    def closeEvent(self, event) -> None:
        self.reject()
        event.accept()

    def _init_ui(self) -> None:
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(12)

        # 1. Grupo Titular
        grp_titular = QGroupBox("Datos del Personal Titular")
        lay_titular = QGridLayout(grp_titular)

        lay_titular.addWidget(QLabel("Legajo:"), 0, 0)
        h_leg = QHBoxLayout()
        self.txt_legajo = QLineEdit()
        self.txt_legajo.setFixedWidth(100)
        self.txt_legajo.editingFinished.connect(self._on_legajo_changed)
        self.btn_buscar_legajo = QPushButton("🔍 Buscar")
        self.btn_buscar_legajo.setObjectName("btnSecondary")
        self.btn_buscar_legajo.clicked.connect(self._buscar_legajo_modal)
        h_leg.addWidget(self.txt_legajo)
        h_leg.addWidget(self.btn_buscar_legajo)
        h_leg.addStretch()
        lay_titular.addLayout(h_leg, 0, 1)

        lay_titular.addWidget(QLabel("Apellidos y Nombres:"), 0, 2)
        self.txt_nombre = QLineEdit()
        self.txt_nombre.setReadOnly(True)
        lay_titular.addWidget(self.txt_nombre, 0, 3)

        lay_titular.addWidget(QLabel("Especialidad:"), 1, 0)
        self.txt_especialidad = QLineEdit()
        self.txt_especialidad.setReadOnly(True)
        lay_titular.addWidget(self.txt_especialidad, 1, 1)

        lay_titular.addWidget(QLabel("Dotación:"), 1, 2)
        self.txt_dotacion = QLineEdit()
        self.txt_dotacion.setReadOnly(True)
        lay_titular.addWidget(self.txt_dotacion, 1, 3)

        lay_titular.addWidget(QLabel("Turno habitual:"), 2, 0)
        self.txt_turnos = QLineEdit()
        self.txt_turnos.setReadOnly(True)
        lay_titular.addWidget(self.txt_turnos, 2, 1)

        lay_titular.addWidget(QLabel("Franco habitual:"), 2, 2)
        self.txt_franco = QLineEdit()
        self.txt_franco.setReadOnly(True)
        lay_titular.addWidget(self.txt_franco, 2, 3)

        main_layout.addWidget(grp_titular)

        # 2. Grupo Novedad
        grp_nov = QGroupBox("Detalle de la Novedad")
        lay_nov = QGridLayout(grp_nov)

        lay_nov.addWidget(QLabel("Tipo de Novedad:"), 0, 0)
        self.combo_tipo = QComboBox()
        self.combo_tipo.addItem("", userData={"dias": 0})
        for t in self._tipos_cache:
            self.combo_tipo.addItem(t["nombre"], userData=t)
        self.combo_tipo.currentIndexChanged.connect(self._on_tipo_novedad_changed)
        lay_nov.addWidget(self.combo_tipo, 0, 1)

        lay_nov.addWidget(QLabel("Fecha de Inicio:"), 0, 2)
        self.date_inicio = QDateEdit()
        self.date_inicio.setCalendarPopup(True)
        self.date_inicio.setDisplayFormat("dd/MM/yyyy")
        self.date_inicio.setDate(QDate.currentDate())
        self.date_inicio.dateChanged.connect(self._on_fecha_inicio_changed)
        lay_nov.addWidget(self.date_inicio, 0, 3)

        lay_nov.addWidget(QLabel("Fecha de Fin:"), 1, 0)
        h_fin = QHBoxLayout()
        self.chk_fecha_fin = QCheckBox("Tiene fecha de fin")
        self.date_fin = QDateEdit()
        self.date_fin.setCalendarPopup(True)
        self.date_fin.setDisplayFormat("dd/MM/yyyy")
        self.date_fin.setDate(QDate.currentDate())
        self.date_fin.setEnabled(False)
        self.chk_fecha_fin.toggled.connect(self.date_fin.setEnabled)
        h_fin.addWidget(self.chk_fecha_fin)
        h_fin.addWidget(self.date_fin)
        lay_nov.addLayout(h_fin, 1, 1)

        lay_nov.addWidget(QLabel("Referencia Estación (*):"), 1, 2)
        self.combo_estacion = QComboBox()
        self.combo_estacion.setEditable(False)
        self.combo_estacion.addItem("")
        for pe in self._personal_estacion_cache:
            dot_txt = f" ({pe['dotacion']})" if pe.get("dotacion") else ""
            self.combo_estacion.addItem(f"{pe['nombre']}{dot_txt}")
        # Agregar dotaciones si no figuran
        try:
            for d in self.records_service.listar_dotaciones(incluir_inactivos=False):
                d_nom = d.get("nombre")
                if d_nom and self.combo_estacion.findText(d_nom) == -1:
                    self.combo_estacion.addItem(d_nom)
        except Exception:
            pass
        lay_nov.addWidget(self.combo_estacion, 1, 3)

        lay_nov.addWidget(QLabel("Supervisor:"), 2, 0)
        self.txt_supervisor = QLineEdit()
        self.txt_supervisor.setReadOnly(True)
        self.txt_supervisor.setEnabled(False)
        u_sup = self.usuario_actual.get("username") or self.usuario_actual.get("nombre") or ""
        self.txt_supervisor.setText(u_sup)
        lay_nov.addWidget(self.txt_supervisor, 2, 1)

        lay_nov.addWidget(QLabel("Usuario Windows:"), 2, 2)
        self.txt_usuario_windows = QLineEdit()
        self.txt_usuario_windows.setReadOnly(True)
        self.txt_usuario_windows.setEnabled(False)
        self.txt_usuario_windows.setText(obtener_usuario_windows())
        lay_nov.addWidget(self.txt_usuario_windows, 2, 3)

        lay_nov.addWidget(QLabel("Observaciones:"), 3, 0)
        self.txt_observaciones = QPlainTextEdit()
        self.txt_observaciones.setMaximumHeight(70)
        lay_nov.addWidget(self.txt_observaciones, 3, 1, 1, 3)

        main_layout.addWidget(grp_nov)

        # 3. Grupo Cobertura de Ciclo
        self.grp_cobertura = QGroupBox("Cobertura con Personal de Ciclo (Reemplazos)")
        lay_cob = QVBoxLayout(self.grp_cobertura)

        bar_cob = QHBoxLayout()
        self.lbl_cobertura_status = QLabel("Cobertura: Sin reemplazo")
        self.lbl_cobertura_status.setStyleSheet("font-weight: bold; color: #C62828;")
        bar_cob.addWidget(self.lbl_cobertura_status)
        bar_cob.addStretch()

        self.btn_agregar_tramo = QPushButton("➕ Agregar Tramo")
        self.btn_agregar_tramo.setObjectName("btnSecondary")
        self.btn_agregar_tramo.clicked.connect(self._agregar_tramo)
        bar_cob.addWidget(self.btn_agregar_tramo)

        self.btn_modificar_tramo = QPushButton("✏️ Modificar Tramo")
        self.btn_modificar_tramo.setObjectName("btnSecondary")
        self.btn_modificar_tramo.clicked.connect(self._modificar_tramo)
        bar_cob.addWidget(self.btn_modificar_tramo)

        self.btn_eliminar_tramo = QPushButton("🗑️ Eliminar Tramo")
        self.btn_eliminar_tramo.setObjectName("btnSecondary")
        self.btn_eliminar_tramo.clicked.connect(self._eliminar_tramo)
        bar_cob.addWidget(self.btn_eliminar_tramo)

        lay_cob.addLayout(bar_cob)

        self.tabla_tramos = QTableWidget(0, 5)
        self.tabla_tramos.setHorizontalHeaderLabels(["LEGAJO", "PERSONAL DE CICLO", "TURNO", "DESDE", "HASTA"])
        self.tabla_tramos.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.tabla_tramos.setSelectionBehavior(QTableWidget.SelectRows)
        self.tabla_tramos.setSelectionMode(QTableWidget.SingleSelection)
        self.tabla_tramos.setMaximumHeight(130)
        lay_cob.addWidget(self.tabla_tramos)

        main_layout.addWidget(self.grp_cobertura)

        # 4. Botones Guardar / Cancelar
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        self.btn_cancel = QPushButton("Cancelar")
        self.btn_cancel.setObjectName("btnSecondary")
        self.btn_cancel.clicked.connect(self.reject)
        btn_layout.addWidget(self.btn_cancel)

        self.btn_guardar = QPushButton("Guardar Novedad")
        self.btn_guardar.clicked.connect(self._guardar)
        btn_layout.addWidget(self.btn_guardar)

        main_layout.addLayout(btn_layout)

        # Modo sólo lectura o restringido
        if self.solo_lectura:
            self._aplicar_solo_lectura()
        elif self.modo_restringido:
            self._aplicar_modo_restringido()

    def _aplicar_solo_lectura(self) -> None:
        if self.novedad_existente:
            self.setWindowTitle(f"Detalle de Novedad Operativa #{self.novedad_existente.get('id', '')}")
        self.txt_legajo.setEnabled(False)
        self.btn_buscar_legajo.setEnabled(False)
        self.combo_tipo.setEnabled(False)
        self.date_inicio.setEnabled(False)
        self.chk_fecha_fin.setEnabled(False)
        self.date_fin.setEnabled(False)
        self.combo_estacion.setEnabled(False)
        self.txt_supervisor.setEnabled(False)
        self.txt_usuario_windows.setEnabled(False)
        self.txt_observaciones.setReadOnly(True)
        self.grp_cobertura.setEnabled(False)
        self.btn_agregar_tramo.setEnabled(False)
        self.btn_modificar_tramo.setEnabled(False)
        self.btn_eliminar_tramo.setEnabled(False)
        self.btn_guardar.setVisible(False)
        self.btn_cancel.setText("Cerrar")

    def _aplicar_modo_restringido(self) -> None:
        self.txt_legajo.setEnabled(False)
        self.btn_buscar_legajo.setEnabled(False)
        self.combo_tipo.setEnabled(False)
        self.date_inicio.setEnabled(False)
        self.combo_estacion.setEnabled(False)
        self.txt_supervisor.setEnabled(False)
        self.grp_cobertura.setEnabled(False)
        self.chk_fecha_fin.setEnabled(True)
        self.date_fin.setEnabled(True)
        self.txt_observaciones.setEnabled(True)
        self.txt_observaciones.setPlaceholderText("Ingrese aquí las observaciones adicionales a concatenar...")

    def _buscar_legajo_modal(self) -> None:
        dlg = LegajoSearchDialog(self._empleados_cache, parent=self)
        if dlg.exec():
            emp = dlg.selected_empleado
            if emp:
                self._cargar_empleado_en_formulario(emp)

    def _on_legajo_changed(self) -> None:
        leg_str = self.txt_legajo.text().strip()
        if not leg_str:
            return
        try:
            leg_int = int(leg_str)
            emp = next((e for e in self._empleados_cache if int(e["legajo"]) == leg_int), None)
            if emp:
                self._cargar_empleado_en_formulario(emp)
            else:
                QMessageBox.warning(self, "Búsqueda", f"No se encontró ningún empleado con legajo {leg_int}.")
        except ValueError:
            pass

    def _cargar_empleado_en_formulario(self, emp: Dict[str, Any]) -> None:
        self.txt_legajo.setText(str(emp["legajo"]))
        self.txt_nombre.setText(str(emp["apellidos_nombres"]))
        self.txt_especialidad.setText(str(emp["especialidad"]))
        self.txt_dotacion.setText(str(emp["dotacion"]))
        self.txt_turnos.setText(str(emp["turnos"]))
        self.txt_franco.setText(str(emp["franco"]))

        if es_turno_ciclo(emp["turnos"]):
            self.grp_cobertura.setEnabled(False)
            self.lbl_cobertura_status.setText("Personal de CICLO (No requiere reemplazo)")
            self.lbl_cobertura_status.setStyleSheet("color: #616161; font-weight: bold;")
            self.tramos.clear()
            self._render_tramos()
        else:
            self.grp_cobertura.setEnabled(True)
            self._render_tramos()

    def _on_tipo_novedad_changed(self) -> None:
        data = self.combo_tipo.currentData()
        if not data:
            return
        dias = int(data.get("dias") or 0)
        if dias > 0:
            f_ini = self.date_inicio.date().toPython()
            f_fin = f_ini + timedelta(days=dias - 1)
            self.date_fin.setDate(QDate(f_fin.year, f_fin.month, f_fin.day))
            self.chk_fecha_fin.setChecked(True)
        else:
            self.chk_fecha_fin.setChecked(False)

    def _on_fecha_inicio_changed(self, new_qdate: QDate) -> None:
        data = self.combo_tipo.currentData()
        if data and int(data.get("dias") or 0) > 0:
            dias = int(data["dias"])
            f_ini = new_qdate.toPython()
            f_fin = f_ini + timedelta(days=dias - 1)
            self.date_fin.setDate(QDate(f_fin.year, f_fin.month, f_fin.day))
            self.chk_fecha_fin.setChecked(True)

    def _render_tramos(self) -> None:
        self.tabla_tramos.setRowCount(len(self.tramos))
        for r, t in enumerate(self.tramos):
            self.tabla_tramos.setItem(r, 0, QTableWidgetItem(str(t.get("legajo_reemplazo") or t.get("legajo"))))
            self.tabla_tramos.setItem(r, 1, QTableWidgetItem(str(t.get("apellidos_nombres_reemplazo") or t.get("apellidos_nombres"))))
            self.tabla_tramos.setItem(r, 2, QTableWidgetItem(str(t.get("turnos_reemplazo") or t.get("turnos"))))
            self.tabla_tramos.setItem(r, 3, QTableWidgetItem(str(t.get("fecha_inicio"))))
            self.tabla_tramos.setItem(r, 4, QTableWidgetItem(str(t.get("fecha_fin") or "Abierto")))

        if not self.tramos:
            self.lbl_cobertura_status.setText("Cobertura: No (Sin reemplazo)")
            self.lbl_cobertura_status.setStyleSheet("color: #C62828; font-weight: bold;")
        else:
            nombres = [str(t.get("apellidos_nombres_reemplazo") or t.get("apellidos_nombres")) for t in self.tramos]
            self.lbl_cobertura_status.setText(f"Cobertura: Asignada ({', '.join(nombres)})")
            self.lbl_cobertura_status.setStyleSheet("color: #2E7D32; font-weight: bold;")

    def _agregar_tramo(self) -> None:
        dot = self.txt_dotacion.text().strip()
        esp = self.txt_especialidad.text().strip()
        if not dot or not esp:
            QMessageBox.warning(self, "Validación", "Primero debe seleccionar un empleado titular.")
            return

        candidatos = self.records_service.obtener_candidatos_ciclo(dot, esp, excluir_legajo=self.txt_legajo.text().strip())
        if not candidatos:
            QMessageBox.information(self, "Personal de Ciclo", f"No hay personal de CICLO activo para la dotación {dot} y especialidad {esp}.")
            return

        f_ini = self.date_inicio.date().toPython()
        f_fin = self.date_fin.date().toPython() if self.chk_fecha_fin.isChecked() else None
        dlg = TramoEditDialog(candidatos, f_ini, f_fin, parent=self)
        if dlg.exec():
            res = dlg.tramo_result
            if res:
                leg_r = res.get("legajo_reemplazo") or res.get("legajo")
                try:
                    self.records_service.validar_reemplazo_disponible(
                        int(leg_r),
                        res.get("fecha_inicio"),
                        res.get("fecha_fin"),
                        excluir_novedad_id=self.novedad_existente.get("id") if self.novedad_existente else None,
                    )
                except ValueError as e:
                    QMessageBox.warning(self, "Relevo No Disponible", str(e))
                    return
                self.tramos.append(res)
                self._render_tramos()

    def _modificar_tramo(self) -> None:
        sel = self.tabla_tramos.selectionModel().selectedRows()
        if not sel:
            QMessageBox.information(self, "Modificar", "Seleccione una fila en la tabla de tramos.")
            return
        idx = sel[0].row()
        tramo = self.tramos[idx]

        dot = self.txt_dotacion.text().strip()
        esp = self.txt_especialidad.text().strip()
        candidatos = self.records_service.obtener_candidatos_ciclo(dot, esp, excluir_legajo=self.txt_legajo.text().strip())

        f_ini = self.date_inicio.date().toPython()
        f_fin = self.date_fin.date().toPython() if self.chk_fecha_fin.isChecked() else None
        dlg = TramoEditDialog(candidatos, f_ini, f_fin, tramo_existente=tramo, parent=self)
        if dlg.exec():
            res = dlg.tramo_result
            if res:
                leg_r = res.get("legajo_reemplazo") or res.get("legajo")
                try:
                    self.records_service.validar_reemplazo_disponible(
                        int(leg_r),
                        res.get("fecha_inicio"),
                        res.get("fecha_fin"),
                        excluir_novedad_id=self.novedad_existente.get("id") if self.novedad_existente else None,
                    )
                except ValueError as e:
                    QMessageBox.warning(self, "Relevo No Disponible", str(e))
                    return
                self.tramos[idx] = res
                self._render_tramos()

    def _eliminar_tramo(self) -> None:
        sel = self.tabla_tramos.selectionModel().selectedRows()
        if not sel:
            return
        idx = sel[0].row()
        del self.tramos[idx]
        self._render_tramos()

    def _cargar_datos_existentes(self) -> None:
        nov = self.novedad_existente or {}
        leg = nov.get("legajo")
        if leg:
            self.txt_legajo.setText(str(leg))
            self._on_legajo_changed()

        tipo = str(nov.get("novedad") or "")
        idx = self.combo_tipo.findText(tipo)
        if idx >= 0:
            self.combo_tipo.setCurrentIndex(idx)

        if nov.get("fecha_inicio"):
            dt_i = parsear_fecha(nov["fecha_inicio"])
            if dt_i:
                self.date_inicio.setDate(QDate(dt_i.year, dt_i.month, dt_i.day))

        if nov.get("fecha_fin") and nov["fecha_fin"] != "-":
            dt_f = parsear_fecha(nov["fecha_fin"])
            if dt_f:
                self.date_fin.setDate(QDate(dt_f.year, dt_f.month, dt_f.day))
                self.chk_fecha_fin.setChecked(True)
        else:
            self.chk_fecha_fin.setChecked(False)

        if nov.get("referencia_estacion"):
            ref_txt = str(nov["referencia_estacion"]).strip()
            idx = self.combo_estacion.findText(ref_txt)
            if idx >= 0:
                self.combo_estacion.setCurrentIndex(idx)
            else:
                self.combo_estacion.addItem(ref_txt)
                self.combo_estacion.setCurrentIndex(self.combo_estacion.count() - 1)
        if nov.get("supervisor"):
            self.txt_supervisor.setText(str(nov["supervisor"]))
        if nov.get("usuario_windows"):
            self.txt_usuario_windows.setText(str(nov["usuario_windows"]))
        if nov.get("observaciones"):
            self.txt_observaciones.setPlainText(str(nov["observaciones"]))

        self._render_tramos()

    def _guardar(self) -> None:
        win_usr = self.txt_usuario_windows.text().strip() or obtener_usuario_windows()
        if self.modo_restringido:
            # En modo restringido solo guardamos fecha de fin y/o nueva observación
            f_fin_str = self.date_fin.date().toString("dd/MM/yyyy") if self.chk_fecha_fin.isChecked() else None
            nueva_obs = self.txt_observaciones.toPlainText().strip()
            try:
                self.records_service.actualizar_fecha_fin_novedad(
                    self.novedad_existente["id"],  # type: ignore
                    fecha_fin=f_fin_str,
                    nueva_observacion=nueva_obs,
                    user_id=self.usuario_actual.get("id"),
                    windows_user=win_usr,
                )
                self.accept()
            except Exception as e:
                QMessageBox.critical(self, "Error", f"No se pudo actualizar el registro:\n{e}")
            return

        # Validación normal de campos
        leg_str = self.txt_legajo.text().strip()
        if not leg_str:
            QMessageBox.warning(self, "Validación", "El legajo es obligatorio.")
            return

        tipo_str = self.combo_tipo.currentText().strip()
        if not tipo_str:
            QMessageBox.warning(self, "Validación", "Debe seleccionar un tipo de novedad.")
            return

        f_ini_py = self.date_inicio.date().toPython()
        f_fin_py = self.date_fin.date().toPython() if self.chk_fecha_fin.isChecked() else None

        if f_fin_py and f_fin_py < f_ini_py:
            QMessageBox.warning(self, "Validación", "La fecha de fin no puede ser anterior a la de inicio.")
            return

        ref_est = self.combo_estacion.currentText().strip()
        if not ref_est:
            QMessageBox.warning(self, "Validación", "La Referencia de Estación es obligatoria.")
            self.combo_estacion.setFocus()
            return

        superv = self.txt_supervisor.text().strip()
        if not superv:
            superv = self.usuario_actual.get("username") or self.usuario_actual.get("nombre") or "Supervisor"
            self.txt_supervisor.setText(superv)

        data_payload = {
            "legajo": int(leg_str),
            "apellidos_nombres": self.txt_nombre.text().strip(),
            "especialidad": self.txt_especialidad.text().strip(),
            "dotacion": self.txt_dotacion.text().strip(),
            "turnos": self.txt_turnos.text().strip(),
            "franco": self.txt_franco.text().strip(),
            "novedad": tipo_str,
            "fecha_inicio": f_ini_py.strftime("%d/%m/%Y"),
            "fecha_fin": f_fin_py.strftime("%d/%m/%Y") if f_fin_py else "-",
            "referencia_estacion": ref_est,
            "supervisor": superv,
            "observaciones": self.txt_observaciones.toPlainText().strip(),
            "usuario_windows": win_usr,
        }

        try:
            if self.novedad_existente:
                nov_id = self.novedad_existente["id"]
                self.records_service.actualizar_novedad(
                    nov_id,
                    data_payload,
                    user_id=self.usuario_actual.get("id"),
                    windows_user=win_usr,
                    bypass_window=True,
                )
            else:
                nov_id = self.records_service.crear_novedad(
                    data_payload,
                    user_id=self.usuario_actual.get("id"),
                    windows_user=win_usr,
                )

            # Guardar tramos de relevo
            if self.grp_cobertura.isEnabled():
                self.records_service.guardar_reemplazos_novedad(
                    nov_id,
                    self.tramos,
                    user_id=self.usuario_actual.get("id"),
                    windows_user=win_usr,
                )

            self.accept()
        except Exception as e:
            QMessageBox.critical(self, "Error al Guardar", f"No se pudo guardar la novedad:\n{e}")
