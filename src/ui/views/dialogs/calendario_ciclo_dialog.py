"""Diálogo del Calendario Mensual de Personal de Ciclo (Maximizado).

Presenta el cronograma día por día del mes con cálculo de descansos reglamentarios CCT
(18/42 hs o 17/41 hs), horas trabajadas acumuladas de la semana y asignación de jornadas.
"""

from typing import Any, Dict, List, Optional
from datetime import date, datetime

from PySide6.QtCore import QDate, QTime, Qt
from PySide6.QtGui import QBrush, QColor, QFont
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDateEdit,
    QDialog,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QTableView,
    QTimeEdit,
    QVBoxLayout,
)

from src.core.logger import get_logger
from src.services.excel_service import ExcelService
from src.services.pdf_service import PDFService
from src.services.records_service import RecordsService
from src.ui.models.base_table_model import BaseTableModel

logger = get_logger("CalendarioCicloDialog")


class AsignarJornadaDialog(QDialog):
    """Diálogo modal para registrar horario de toma/deja o franco anticipado."""

    def __init__(
        self,
        legajo: int,
        nombre: str,
        fecha_iso: str,
        dia_semana: str,
        jornada_existente: Optional[Dict[str, Any]] = None,
        franco_agente: str = "Domingo",
        records_service: Optional[RecordsService] = None,
        parent=None,
    ):
        super().__init__(parent)
        self.setWindowTitle("Asignar Jornada / Franco")
        self.setMinimumWidth(440)

        self.legajo = legajo
        self.fecha_iso = fecha_iso
        self.dia_semana = dia_semana
        self.franco_agente = franco_agente
        self.records_service = records_service
        self.datos_resultado: Optional[Dict[str, Any]] = None

        layout = QVBoxLayout(self)
        form = QGridLayout()

        form.addWidget(QLabel("Personal de Ciclo:"), 0, 0)
        lbl_nom = QLabel(f"<b>{nombre}</b> (Leg. {legajo})")
        form.addWidget(lbl_nom, 0, 1)

        form.addWidget(QLabel("Fecha:"), 1, 0)
        partes = fecha_iso.split("-")
        f_fmt = f"{partes[2]}/{partes[1]}/{partes[0]}" if len(partes) == 3 else fecha_iso
        lbl_f = QLabel(f"<b>{f_fmt} ({dia_semana})</b> | Franco: <b>{franco_agente}</b>")
        form.addWidget(lbl_f, 1, 1)

        # Selector Asistido de Turno Itinerario
        form.addWidget(QLabel("Turno Itinerario (Opcional):"), 2, 0)
        h_turno = QHBoxLayout()
        self.txt_turno_base = QLineEdit()
        self.txt_turno_base.setPlaceholderText("Ej. 401, 410, PROG.1CK5...")
        self.btn_resolver = QPushButton("⚡ Autocompletar")
        self.btn_resolver.setObjectName("btnSecondary")
        self.btn_resolver.clicked.connect(self._autocompletar_por_turno)
        self.txt_turno_base.returnPressed.connect(self._autocompletar_por_turno)
        h_turno.addWidget(self.txt_turno_base)
        h_turno.addWidget(self.btn_resolver)
        form.addLayout(h_turno, 2, 1)

        # Franco anticipado checkbox
        self.chk_fa = QCheckBox("Marcar como Franco Anticipado")
        form.addWidget(self.chk_fa, 3, 0, 1, 2)

        # Horarios Toma / Deja con QTimeEdit
        form.addWidget(QLabel("Horario de Toma (HH:mm):"), 4, 0)
        self.time_toma = QTimeEdit()
        self.time_toma.setDisplayFormat("HH:mm")
        self.time_toma.setTime(QTime(6, 0))
        form.addWidget(self.time_toma, 4, 1)

        form.addWidget(QLabel("Horario de Deja (HH:mm):"), 5, 0)
        self.time_deja = QTimeEdit()
        self.time_deja.setDisplayFormat("HH:mm")
        self.time_deja.setTime(QTime(12, 0))
        form.addWidget(self.time_deja, 5, 1)

        form.addWidget(QLabel("Observaciones:"), 6, 0)
        self.txt_obs = QLineEdit()
        form.addWidget(self.txt_obs, 6, 1)

        self.chk_fa.toggled.connect(self._on_fa_toggled)
        layout.addLayout(form)

        if jornada_existente:
            if jornada_existente.get("franco_anticipado"):
                self.chk_fa.setChecked(True)
            if jornada_existente.get("toma") and jornada_existente["toma"] != "-":
                t = QTime.fromString(str(jornada_existente["toma"]).strip(), "HH:mm")
                if t.isValid():
                    self.time_toma.setTime(t)
            if jornada_existente.get("deja") and jornada_existente["deja"] != "-":
                d = QTime.fromString(str(jornada_existente["deja"]).strip(), "HH:mm")
                if d.isValid():
                    self.time_deja.setTime(d)
            if jornada_existente.get("observaciones"):
                self.txt_obs.setText(str(jornada_existente["observaciones"]))

        # Botones
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        btn_cancel = QPushButton("Cancelar")
        btn_cancel.setObjectName("btnSecondary")
        btn_cancel.clicked.connect(self.reject)
        btn_ok = QPushButton("Guardar")
        btn_ok.clicked.connect(self._on_accept)
        btn_layout.addWidget(btn_cancel)
        btn_layout.addWidget(btn_ok)
        layout.addLayout(btn_layout)

    def _autocompletar_por_turno(self) -> None:
        base = self.txt_turno_base.text().strip()
        if not base:
            QMessageBox.information(self, "Itinerario", "Ingrese un código de turno base (ej. 401, 410, PROG.1CK5).")
            return

        if not self.records_service:
            return

        turnos_map = self.records_service.listar_turnos_itinerario_map()
        if not turnos_map:
            QMessageBox.warning(
                self,
                "Itinerarios Vacíos",
                "No hay base de itinerarios cargada en el sistema. Cargue 'Indice_45.xlsx' desde el Buscador de Trenes.",
            )
            return

        from src.domain.validators import resolver_jornada_itinerario
        try:
            res = resolver_jornada_itinerario(
                turno_base=base,
                fecha_target=self.fecha_iso,
                franco_agente=self.franco_agente,
                turnos_dict=turnos_map,
            )

            if res.get("es_franco"):
                self.chk_fa.setChecked(True)
                self.txt_obs.setText(f"Turno {base} - FRANCO")
                QMessageBox.information(
                    self,
                    "Franco Detectado",
                    f"Para la fecha seleccionada ({self.dia_semana}), el turno {base} corresponde a FRANCO según el ciclo del agente.",
                )
            else:
                self.chk_fa.setChecked(False)
                if res.get("toma"):
                    t = QTime.fromString(res["toma"].strip(), "HH:mm")
                    if t.isValid():
                        self.time_toma.setTime(t)
                if res.get("deje"):
                    d = QTime.fromString(res["deje"].strip(), "HH:mm")
                    if d.isValid():
                        self.time_deja.setTime(d)

                detalles = f"Turno {res.get('codigo_turno')}"
                if res.get("deposito"):
                    detalles += f" ({res['deposito']})"
                self.txt_obs.setText(detalles)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Fallo al resolver turno itinerario:\n{e}")

    def _on_fa_toggled(self, checked: bool) -> None:
        self.time_toma.setEnabled(not checked)
        self.time_deja.setEnabled(not checked)

    def _on_accept(self) -> None:
        fa = 1 if self.chk_fa.isChecked() else 0
        toma = self.time_toma.time().toString("HH:mm") if not fa else ""
        deja = self.time_deja.time().toString("HH:mm") if not fa else ""
        obs = self.txt_obs.text().strip()

        self.datos_resultado = {
            "toma": toma,
            "deja": deja,
            "franco_anticipado": fa,
            "observaciones": obs,
        }
        self.accept()


class CalendarioCicloDialog(QDialog):
    """Ventana del Calendario Mensual de Personal de Ciclo (abierta maximizada)."""

    def __init__(
        self,
        empleado: Dict[str, Any],
        records_service: RecordsService,
        pdf_service: PDFService,
        excel_service: ExcelService,
        mes: Optional[int] = None,
        anio: Optional[int] = None,
        usuario_actual: Optional[Dict[str, Any]] = None,
        parent=None,
    ):
        super().__init__(parent)
        self.empleado = empleado
        self.records_service = records_service
        self.pdf_service = pdf_service
        self.excel_service = excel_service
        self.usuario_actual = usuario_actual or {}

        hoy = date.today()
        self.mes = mes or hoy.month
        self.anio = anio or hoy.year
        self._dias_datos: List[Dict[str, Any]] = []

        self.setWindowTitle(f"Calendario Mensual de Ciclo - {self.empleado.get('apellidos_nombres')} (Leg. {self.empleado.get('legajo')})")
        self.resize(1100, 700)
        self.setWindowFlags(Qt.Window | Qt.WindowTitleHint | Qt.WindowMinMaxButtonsHint | Qt.WindowCloseButtonHint)

        self._init_ui()
        self.cargar_mes()

    def closeEvent(self, event) -> None:
        self.reject()
        event.accept()

    def showEvent(self, event) -> None:
        super().showEvent(event)
        # Requerimiento: modales maximizados para garantizar visibilidad completa
        self.showMaximized()

    def _init_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)

        # 1. Cabecera con datos del agente y selector de Mes/Año
        header_frame = QFrame()
        header_frame.setObjectName("headerFrame")
        h_layout = QHBoxLayout(header_frame)

        # Info del empleado
        v_info = QVBoxLayout()
        v_info.setSpacing(3)
        lbl_nombre = QLabel(f"<b>{self.empleado.get('apellidos_nombres')}</b> (Legajo: {self.empleado.get('legajo')})")
        lbl_nombre.setObjectName("lblHeaderTitle")
        lbl_detalles = QLabel(
            f"Dotación: <b>{self.empleado.get('dotacion')}</b> | "
            f"Especialidad: <b>{self.empleado.get('especialidad')}</b> | "
            f"Turno: <b>{self.empleado.get('turnos')}</b> | "
            f"Franco CCT: <b>{self.empleado.get('franco')}</b> | "
            f"Jornada: <b>{self.empleado.get('jornada_horas', 6)} hs</b>"
        )
        lbl_detalles.setObjectName("lblHeaderSubtitle")
        v_info.addWidget(lbl_nombre)
        v_info.addWidget(lbl_detalles)
        h_layout.addLayout(v_info)

        h_layout.addStretch()

        # Selector de mes y año
        h_mes = QHBoxLayout()
        h_mes.addWidget(QLabel("Mes:"))
        self.combo_mes = QComboBox()
        meses = [
            "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
            "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
        ]
        self.combo_mes.addItems(meses)
        self.combo_mes.setCurrentIndex(self.mes - 1)
        self.combo_mes.currentIndexChanged.connect(self._on_mes_anio_changed)
        h_mes.addWidget(self.combo_mes)

        h_mes.addWidget(QLabel("Año:"))
        self.spin_anio = QSpinBox()
        self.spin_anio.setRange(2020, 2040)
        self.spin_anio.setValue(self.anio)
        self.spin_anio.valueChanged.connect(self._on_mes_anio_changed)
        h_mes.addWidget(self.spin_anio)

        h_layout.addLayout(h_mes)
        layout.addWidget(header_frame)

        # 2. Barra de acciones secundarias
        action_bar = QHBoxLayout()
        self.btn_asignar = QPushButton("⏱️ Asignar Jornada / Franco")
        self.btn_asignar.clicked.connect(self._asignar_jornada_seleccionada)
        action_bar.addWidget(self.btn_asignar)

        action_bar.addStretch()

        self.btn_excel = QPushButton("📊 Exportar Cronograma Excel")
        self.btn_excel.setObjectName("btnSecondary")
        self.btn_excel.clicked.connect(self._exportar_excel_individual)
        action_bar.addWidget(self.btn_excel)

        self.btn_excel_todos = QPushButton("📑 Exportar Todos los Ciclos (Excel)")
        self.btn_excel_todos.setObjectName("btnSecondary")
        self.btn_excel_todos.clicked.connect(self._exportar_excel_todos)
        action_bar.addWidget(self.btn_excel_todos)

        self.btn_pdf = QPushButton("📑 Exportar Cronograma PDF")
        self.btn_pdf.setObjectName("btnSecondary")
        self.btn_pdf.clicked.connect(self._exportar_pdf)
        action_bar.addWidget(self.btn_pdf)

        layout.addLayout(action_bar)

        # 3. Tabla de días
        headers = [
            "DÍA", "FECHA", "DÍA SEM.", "ESTADO", "HORA DISPONIBLE",
            "TOMA", "DEJA", "HS TRABAJADAS", "HS SEMANA", "REEMPLAZA A / MOTIVO", "OBSERVACIONES",
        ]
        keys = [
            "dia", "fecha", "dia_semana", "estado", "hora_disponible",
            "toma", "deja", "horas_trabajadas", "hs_semana", "reemplaza_a", "observaciones",
        ]

        def conditional_colors(row_item: Dict[str, Any], col_idx: int):
            estado = str(row_item.get("estado", ""))
            from src.core.theme import obtener_tema_visual
            es_oscuro = obtener_tema_visual() == "oscuro"
            if "Franco" in estado:
                return (QColor("#3D3214"), QColor("#FCD34D")) if es_oscuro else (QColor("#FFF9C4"), QColor("#B78103"))
            elif "Reemplazando" in estado or "Relevo" in estado:
                return (QColor("#132E1F"), QColor("#6EE7B7")) if es_oscuro else (QColor("#E8F5E9"), QColor("#2E7D32"))
            elif estado == "Disponible":
                return (QColor("#172554"), QColor("#93C5FD")) if es_oscuro else (QColor("#F0F7FF"), QColor("#1F497D"))
            return None

        self.table_model = BaseTableModel(headers, keys, conditional_colors=conditional_colors, parent=self)
        self.table_view = QTableView()
        self.table_view.setModel(self.table_model)
        self.table_view.setSelectionBehavior(QTableView.SelectRows)
        self.table_view.setSelectionMode(QTableView.SingleSelection)
        self.table_view.horizontalHeader().setSectionResizeMode(9, QHeaderView.Stretch)
        self.table_view.doubleClicked.connect(self._on_row_double_clicked)
        layout.addWidget(self.table_view)

        # 4. Pie de cierre
        btn_close = QPushButton("Cerrar")
        btn_close.setObjectName("btnSecondary")
        btn_close.clicked.connect(self.accept)
        btn_close_layout = QHBoxLayout()
        btn_close_layout.addStretch()
        btn_close_layout.addWidget(btn_close)
        layout.addLayout(btn_close_layout)

    def cargar_mes(self) -> None:
        leg = int(self.empleado["legajo"])
        self._dias_datos = self.records_service.obtener_calendario_mensual_ciclo(leg, self.anio, self.mes)
        self.table_model.set_data(self._dias_datos)

    def _on_mes_anio_changed(self) -> None:
        self.mes = self.combo_mes.currentIndex() + 1
        self.anio = self.spin_anio.value()
        self.cargar_mes()

    def _obtener_dia_seleccionado(self) -> Optional[Dict[str, Any]]:
        indexes = self.table_view.selectionModel().selectedRows()
        if not indexes:
            return None
        return self.table_model.get_row(indexes[0].row())

    def _asignar_jornada_seleccionada(self) -> None:
        dia_info = self._obtener_dia_seleccionado()
        if not dia_info:
            QMessageBox.information(self, "Asignar", "Seleccione un día en la tabla.")
            return

        dlg = AsignarJornadaDialog(
            legajo=int(self.empleado["legajo"]),
            nombre=str(self.empleado["apellidos_nombres"]),
            fecha_iso=dia_info["fecha_iso"],
            dia_semana=dia_info["dia_semana"],
            jornada_existente=dia_info,
            franco_agente=str(self.empleado.get("franco") or "Domingo"),
            records_service=self.records_service,
            parent=self,
        )
        if dlg.exec():
            res = dlg.datos_resultado
            if res:
                try:
                    self.records_service.guardar_jornada_ciclo(
                        legajo=int(self.empleado["legajo"]),
                        fecha=dia_info["fecha_iso"],
                        toma=res["toma"],
                        deja=res["deja"],
                        franco_anticipado=res["franco_anticipado"],
                        observaciones=res["observaciones"],
                        user_id=self.usuario_actual.get("id"),
                        windows_user=self.usuario_actual.get("username"),
                    )
                    self.cargar_mes()
                except Exception as e:
                    QMessageBox.critical(self, "Error", f"No se pudo guardar la jornada:\n{e}")

    def _on_row_double_clicked(self, index) -> None:
        dia_info = self.table_model.get_row(index.row())
        if not dia_info:
            return

        # Si el día tiene una novedad asociada, podemos informar
        if dia_info.get("novedad_id"):
            QMessageBox.information(
                self,
                f"Relevo Asignado - Novedad #{dia_info['novedad_id']}",
                f"Agente Titular: {dia_info.get('reemplaza_a')}\n"
                f"Turno cubierto: {dia_info.get('turno_cubierto')}\n"
                f"Período del relevo: {dia_info.get('periodo')}\n\n"
                f"Para modificar el turno u horas de toma/deja, utilice 'Asignar Jornada'.",
            )
        else:
            self._asignar_jornada_seleccionada()

    def _exportar_pdf(self) -> None:
        default_name = f"Calendario_Ciclo_{self.empleado.get('legajo')}_{self.anio}_{self.mes:02d}.pdf"
        file_path, _ = QFileDialog.getSaveFileName(self, "Guardar Cronograma Mensual PDF", default_name, "Archivo PDF (*.pdf)")
        if not file_path:
            return

        try:
            self.pdf_service.generar_calendario_ciclo_pdf(
                self.empleado,
                self._dias_datos,
                self.mes,
                self.anio,
                file_path,
            )
            QMessageBox.information(self, "Exportación", f"Cronograma PDF generado exitosamente en:\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo generar el PDF:\n{e}")

    def _exportar_excel_individual(self) -> None:
        default_name = f"Calendario_Ciclo_{self.empleado.get('legajo')}_{self.anio}_{self.mes:02d}.xlsx"
        file_path, _ = QFileDialog.getSaveFileName(self, "Guardar Cronograma Mensual Excel", default_name, "Excel (*.xlsx)")
        if not file_path:
            return
        try:
            self.excel_service.export_calendario_ciclo_individual(
                file_path,
                self.empleado,
                self._dias_datos,
                self.mes,
                self.anio,
            )
            QMessageBox.information(self, "Exportación", f"Cronograma Excel generado exitosamente en:\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo generar el archivo Excel:\n{e}")

    def _exportar_excel_todos(self) -> None:
        default_name = f"Calendario_Ciclo_Todos_{self.anio}_{self.mes:02d}.xlsx"
        file_path, _ = QFileDialog.getSaveFileName(self, "Guardar Consolidado de Todo el Personal de Ciclo", default_name, "Excel (*.xlsx)")
        if not file_path:
            return
        try:
            self.excel_service.export_calendario_ciclo_todos(
                file_path,
                self.records_service,
                self.mes,
                self.anio,
            )
            QMessageBox.information(self, "Exportación", f"Consolidado general de ciclo exportado exitosamente en:\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo exportar el consolidado general:\n{e}")
