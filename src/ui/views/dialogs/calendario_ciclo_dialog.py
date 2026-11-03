"""Diálogo del Calendario Mensual de Personal de Ciclo (Maximizado).

Presenta el cronograma día por día del mes con cálculo de descansos reglamentarios CCT
(18/42 hs o 17/41 hs), horas trabajadas acumuladas de la semana y asignación de jornadas.
"""

from typing import Any, Dict, List, Optional
from datetime import date, datetime

from PySide6.QtCore import QDate, Qt
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
        parent=None,
    ):
        super().__init__(parent)
        self.setWindowTitle("Asignar Jornada / Franco")
        self.setMinimumWidth(400)

        self.legajo = legajo
        self.fecha_iso = fecha_iso
        self.datos_resultado: Optional[Dict[str, Any]] = None

        layout = QVBoxLayout(self)
        form = QGridLayout()

        form.addWidget(QLabel("Personal de Ciclo:"), 0, 0)
        lbl_nom = QLabel(f"<b>{nombre}</b> (Leg. {legajo})")
        form.addWidget(lbl_nom, 0, 1)

        form.addWidget(QLabel("Fecha:"), 1, 0)
        partes = fecha_iso.split("-")
        f_fmt = f"{partes[2]}/{partes[1]}/{partes[0]}" if len(partes) == 3 else fecha_iso
        lbl_f = QLabel(f"<b>{f_fmt} ({dia_semana})</b>")
        form.addWidget(lbl_f, 1, 1)

        # Franco anticipado checkbox
        self.chk_fa = QCheckBox("Marcar como Franco Anticipado")
        form.addWidget(self.chk_fa, 2, 0, 1, 2)

        # Horarios Toma / Deja
        form.addWidget(QLabel("Horario de Toma (HH:MM):"), 3, 0)
        self.txt_toma = QLineEdit()
        self.txt_toma.setPlaceholderText("Ej. 06:00")
        form.addWidget(self.txt_toma, 3, 1)

        form.addWidget(QLabel("Horario de Deja (HH:MM):"), 4, 0)
        self.txt_deja = QLineEdit()
        self.txt_deja.setPlaceholderText("Ej. 12:00")
        form.addWidget(self.txt_deja, 4, 1)

        form.addWidget(QLabel("Observaciones:"), 5, 0)
        self.txt_obs = QLineEdit()
        form.addWidget(self.txt_obs, 5, 1)

        self.chk_fa.toggled.connect(self._on_fa_toggled)
        layout.addLayout(form)

        if jornada_existente:
            if jornada_existente.get("franco_anticipado"):
                self.chk_fa.setChecked(True)
            if jornada_existente.get("toma") and jornada_existente["toma"] != "-":
                self.txt_toma.setText(str(jornada_existente["toma"]))
            if jornada_existente.get("deja") and jornada_existente["deja"] != "-":
                self.txt_deja.setText(str(jornada_existente["deja"]))
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

    def _on_fa_toggled(self, checked: bool) -> None:
        if checked:
            self.txt_toma.clear()
            self.txt_deja.clear()
            self.txt_toma.setEnabled(False)
            self.txt_deja.setEnabled(False)
        else:
            self.txt_toma.setEnabled(True)
            self.txt_deja.setEnabled(True)

    def _on_accept(self) -> None:
        toma = self.txt_toma.text().strip()
        deja = self.txt_deja.text().strip()
        fa = 1 if self.chk_fa.isChecked() else 0
        obs = self.txt_obs.text().strip()

        if not fa:
            if (toma and not deja) or (deja and not toma):
                QMessageBox.warning(self, "Validación", "Debe ingresar tanto la hora de toma como la de dejada.")
                return

        self.datos_resultado = {
            "toma": toma if not fa else "",
            "deja": deja if not fa else "",
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
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)

        self._init_ui()
        self.cargar_mes()

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
        header_frame.setStyleSheet("background-color: #FFFFFF; border: 1px solid #D1D5DB; border-radius: 6px; padding: 10px;")
        h_layout = QHBoxLayout(header_frame)

        # Info del empleado
        v_info = QVBoxLayout()
        v_info.setSpacing(3)
        lbl_nombre = QLabel(f"<b>{self.empleado.get('apellidos_nombres')}</b> (Legajo: {self.empleado.get('legajo')})")
        lbl_nombre.setStyleSheet("font-size: 13pt; color: #1F497D;")
        lbl_detalles = QLabel(
            f"Dotación: <b>{self.empleado.get('dotacion')}</b> | "
            f"Especialidad: <b>{self.empleado.get('especialidad')}</b> | "
            f"Turno: <b>{self.empleado.get('turnos')}</b> | "
            f"Franco CCT: <b>{self.empleado.get('franco')}</b> | "
            f"Jornada: <b>{self.empleado.get('jornada_horas', 6)} hs</b>"
        )
        lbl_detalles.setStyleSheet("color: #4B5563; font-size: 9.5pt;")
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
            if "Franco" in estado:
                return (QColor("#FFF9C4"), QColor("#B78103"))  # Amarillo cálido
            elif estado == "Reemplazando":
                return (QColor("#E8F5E9"), QColor("#2E7D32"))  # Verde suave
            elif estado == "Disponible":
                return (QColor("#F0F7FF"), QColor("#1F497D"))  # Azul suave
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
