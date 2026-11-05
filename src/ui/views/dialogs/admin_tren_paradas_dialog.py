"""Diálogo CRUD e Importador Excel para Estaciones Intermedias y Horarios de Paso por Tren."""

from __future__ import annotations

from typing import Any, Dict, List, Optional

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QFileDialog,
    QFormLayout,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from src.core.config import obtener_usuario_windows
from src.services.records_service import RecordsService


class AdminTrenParadasDialog(QDialog):
    """Administración CRUD e importación desde Excel de estaciones intermedias por tren (`tren_paradas`)."""

    def __init__(
        self,
        records_service: RecordsService,
        usuario_actual: Optional[Dict[str, Any]] = None,
        tren_inicial: str = "",
        parent: Optional[QWidget] = None,
    ) -> None:
        super().__init__(parent)
        self.records_service = records_service
        self.usuario_actual = usuario_actual or {}
        self._parada_editando_id: Optional[int] = None
        self._datos_cache: List[Dict[str, Any]] = []

        self.setWindowTitle("Gestión de Estaciones Intermedias y Horarios de Paso por Tren")
        self.resize(1100, 680)
        self.setWindowFlags(
            (self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
            | Qt.WindowMinMaxButtonsHint
            | Qt.WindowCloseButtonHint
        )

        self._init_ui(tren_inicial)
        self.refrescar_tabla()

    def showEvent(self, event) -> None:
        super().showEvent(event)
        self.showMaximized()

    def _init_ui(self, tren_inicial: str) -> None:
        layout = QVBoxLayout(self)
        layout.setSpacing(10)

        # Barra de Filtros y Acciones Superiores
        grp_filtros = QGroupBox("Filtros de Búsqueda e Importación Masiva")
        lay_f = QHBoxLayout(grp_filtros)

        lay_f.addWidget(QLabel("N° Tren:"))
        self.txt_filtro_tren = QLineEdit(tren_inicial)
        self.txt_filtro_tren.setPlaceholderText("Ej: 1293, 4356...")
        self.txt_filtro_tren.setMaximumWidth(140)
        self.txt_filtro_tren.textChanged.connect(self.refrescar_tabla)
        lay_f.addWidget(self.txt_filtro_tren)

        lay_f.addWidget(QLabel("Estación:"))
        self.txt_filtro_estacion = QLineEdit()
        self.txt_filtro_estacion.setPlaceholderText("Ej: Temperley, Llavallol...")
        self.txt_filtro_estacion.setMaximumWidth(200)
        self.txt_filtro_estacion.textChanged.connect(self.refrescar_tabla)
        lay_f.addWidget(self.txt_filtro_estacion)

        lay_f.addWidget(QLabel("Itinerario:"))
        self.combo_filtro_itin = QComboBox()
        self.combo_filtro_itin.addItem("Todos")
        for it in self.records_service.listar_itinerarios_disponibles():
            self.combo_filtro_itin.addItem(it)
        self.combo_filtro_itin.currentIndexChanged.connect(self.refrescar_tabla)
        lay_f.addWidget(self.combo_filtro_itin)

        self.btn_limpiar_filtros = QPushButton("Limpiar Filtros")
        self.btn_limpiar_filtros.setObjectName("btnSecondary")
        self.btn_limpiar_filtros.clicked.connect(self._limpiar_filtros)
        lay_f.addWidget(self.btn_limpiar_filtros)

        lay_f.addStretch()

        self.btn_importar_excel = QPushButton("📥 Importar Paradas desde Excel")
        self.btn_importar_excel.clicked.connect(self._importar_excel)
        lay_f.addWidget(self.btn_importar_excel)

        layout.addWidget(grp_filtros)

        # Splitter Principal: Tabla a la izquierda + Formulario CRUD a la derecha
        splitter = QSplitter(Qt.Horizontal)

        # Panel Izquierdo: Tabla
        left_widget = QWidget()
        left_lay = QVBoxLayout(left_widget)
        left_lay.setContentsMargins(0, 0, 0, 0)

        self.lbl_contador = QLabel("0 paradas registradas")
        self.lbl_contador.setStyleSheet("font-weight: bold;")
        left_lay.addWidget(self.lbl_contador)

        self.tabla = QTableWidget(0, 7)
        self.tabla.setHorizontalHeaderLabels([
            "ID", "N° TREN", "ORDEN", "ESTACIÓN / PARADA", "HORA PASO", "ITINERARIO", "RÉGIMEN"
        ])
        self.tabla.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        self.tabla.setSelectionBehavior(QTableWidget.SelectRows)
        self.tabla.setSelectionMode(QTableWidget.SingleSelection)
        self.tabla.setEditTriggers(QTableWidget.NoEditTriggers)
        self.tabla.itemSelectionChanged.connect(self._on_seleccion_tabla)
        left_lay.addWidget(self.tabla)

        splitter.addWidget(left_widget)

        # Panel Derecho: Formulario CRUD
        grp_form = QGroupBox("Formulario de Parada / Horario de Paso")
        grp_form.setMaximumWidth(380)
        form_lay = QVBoxLayout(grp_form)

        f_grid = QFormLayout()
        self.txt_nro_tren = QLineEdit()
        self.txt_nro_tren.setPlaceholderText("Ej: 1293")
        f_grid.addRow("N° de Tren (*):", self.txt_nro_tren)

        self.spin_orden = QSpinBox()
        self.spin_orden.setRange(1, 100)
        self.spin_orden.setValue(1)
        f_grid.addRow("Orden de Parada:", self.spin_orden)

        self.txt_estacion = QLineEdit()
        self.txt_estacion.setPlaceholderText("Ej: P. Constitución, Temperley...")
        f_grid.addRow("Estación (*):", self.txt_estacion)

        self.txt_hora_paso = QLineEdit()
        self.txt_hora_paso.setPlaceholderText("HH:MM (Ej: 08:15)")
        f_grid.addRow("Hora de Paso (*):", self.txt_hora_paso)

        self.combo_itin = QComboBox()
        self.combo_itin.setEditable(True)
        for it in self.records_service.listar_itinerarios_disponibles():
            self.combo_itin.addItem(it)
        if self.combo_itin.count() == 0:
            self.combo_itin.addItem("H45")
        f_grid.addRow("Itinerario:", self.combo_itin)

        self.combo_regimen = QComboBox()
        self.combo_regimen.addItems(["LV", "S", "D", " DIAGRAMA_6D", "UNIPERSONAL"])
        f_grid.addRow("Régimen:", self.combo_regimen)

        form_lay.addLayout(f_grid)
        form_lay.addSpacing(10)

        self.lbl_modo_form = QLabel("Modo: Crear nueva parada")
        self.lbl_modo_form.setStyleSheet("font-weight: bold; color: #1F497D;")
        form_lay.addWidget(self.lbl_modo_form)

        btn_guardar = QPushButton("💾 Guardar Parada")
        btn_guardar.clicked.connect(self._guardar_parada)
        form_lay.addWidget(btn_guardar)

        btn_nuevo = QPushButton("➕ Limpiar / Nueva Parada")
        btn_nuevo.setObjectName("btnSecondary")
        btn_nuevo.clicked.connect(self._limpiar_formulario)
        form_lay.addWidget(btn_nuevo)

        self.btn_eliminar = QPushButton("🗑️ Eliminar Parada Seleccionada")
        self.btn_eliminar.setObjectName("btnDanger")
        self.btn_eliminar.setEnabled(False)
        self.btn_eliminar.clicked.connect(self._eliminar_parada)
        form_lay.addWidget(self.btn_eliminar)

        form_lay.addStretch()

        lbl_ayuda = QLabel(
            "Nota: Puede importar una planilla Excel con columnas\n"
            "TREN, ESTACION, HORA_PASO, ORDEN, ITINERARIO\n"
            "o una planilla de itinerarios estándar."
        )
        lbl_ayuda.setWordWrap(True)
        lbl_ayuda.setStyleSheet("color: #666; font-size: 11px;")
        form_lay.addWidget(lbl_ayuda)

        splitter.addWidget(grp_form)
        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 1)
        layout.addWidget(splitter, 1)

        # Pie de cierre
        bottom_lay = QHBoxLayout()
        bottom_lay.addStretch()
        btn_cerrar = QPushButton("Cerrar")
        btn_cerrar.setObjectName("btnSecondary")
        btn_cerrar.clicked.connect(self.accept)
        bottom_lay.addWidget(btn_cerrar)
        layout.addLayout(bottom_lay)

    def _limpiar_filtros(self) -> None:
        self.txt_filtro_tren.clear()
        self.txt_filtro_estacion.clear()
        self.combo_filtro_itin.setCurrentIndex(0)
        self.refrescar_tabla()

    def refrescar_tabla(self) -> None:
        nt = self.txt_filtro_tren.text().strip()
        est = self.txt_filtro_estacion.text().strip()
        itin = self.combo_filtro_itin.currentText().strip()

        self._datos_cache = self.records_service.listar_paradas_tren(
            nro_tren=nt or None,
            itinerario=itin if itin != "Todos" else None,
            estacion=est or None,
            limite=1500,
        )

        self.tabla.setRowCount(len(self._datos_cache))
        for r, item in enumerate(self._datos_cache):
            self.tabla.setItem(r, 0, QTableWidgetItem(str(item["id"])))
            self.tabla.setItem(r, 1, QTableWidgetItem(str(item["nro_tren"])))
            self.tabla.setItem(r, 2, QTableWidgetItem(str(item["orden_parada"])))
            self.tabla.setItem(r, 3, QTableWidgetItem(str(item["estacion"])))
            self.tabla.setItem(r, 4, QTableWidgetItem(str(item["hora_paso"])))
            self.tabla.setItem(r, 5, QTableWidgetItem(str(item.get("itinerario") or "")))
            self.tabla.setItem(r, 6, QTableWidgetItem(str(item.get("regimen") or "LV")))

        self.lbl_contador.setText(f"{len(self._datos_cache)} paradas mostradas")

    def _on_seleccion_tabla(self) -> None:
        sel = self.tabla.selectionModel().selectedRows()
        if not sel:
            self._limpiar_formulario()
            return
        idx = sel[0].row()
        if idx < 0 or idx >= len(self._datos_cache):
            return
        item = self._datos_cache[idx]
        self._parada_editando_id = int(item["id"])
        self.txt_nro_tren.setText(str(item["nro_tren"]))
        self.spin_orden.setValue(int(item.get("orden_parada") or 1))
        self.txt_estacion.setText(str(item["estacion"]))
        self.txt_hora_paso.setText(str(item["hora_paso"]))
        self.combo_itin.setCurrentText(str(item.get("itinerario") or ""))
        reg = str(item.get("regimen") or "LV")
        idx_r = self.combo_regimen.findText(reg)
        if idx_r >= 0:
            self.combo_regimen.setCurrentIndex(idx_r)
        self.lbl_modo_form.setText(f"Modo: Editando parada ID #{self._parada_editando_id}")
        self.btn_eliminar.setEnabled(True)

    def _limpiar_formulario(self) -> None:
        self._parada_editando_id = None
        if not self.txt_filtro_tren.text().strip():
            self.txt_nro_tren.clear()
        else:
            self.txt_nro_tren.setText(self.txt_filtro_tren.text().strip())
        self.spin_orden.setValue(1)
        self.txt_estacion.clear()
        self.txt_hora_paso.clear()
        self.lbl_modo_form.setText("Modo: Crear nueva parada")
        self.btn_eliminar.setEnabled(False)

    def _guardar_parada(self) -> None:
        datos = {
            "nro_tren": self.txt_nro_tren.text().strip(),
            "orden_parada": self.spin_orden.value(),
            "estacion": self.txt_estacion.text().strip(),
            "hora_paso": self.txt_hora_paso.text().strip(),
            "itinerario": self.combo_itin.currentText().strip(),
            "regimen": self.combo_regimen.currentText().strip(),
        }
        try:
            win_usr = obtener_usuario_windows()
            uid = self.usuario_actual.get("id")
            if self._parada_editando_id is not None:
                self.records_service.actualizar_parada_tren(
                    self._parada_editando_id, datos, user_id=uid, windows_user=win_usr
                )
            else:
                self.records_service.crear_parada_tren(
                    datos, user_id=uid, windows_user=win_usr
                )
            self._limpiar_formulario()
            self.refrescar_tabla()
        except Exception as e:
            QMessageBox.warning(self, "Validación", str(e))

    def _eliminar_parada(self) -> None:
        if self._parada_editando_id is None:
            return
        resp = QMessageBox.question(
            self,
            "Confirmar Eliminación",
            f"¿Está seguro que desea eliminar la parada ID #{self._parada_editando_id}?",
            QMessageBox.Yes | QMessageBox.No,
        )
        if resp != QMessageBox.Yes:
            return
        try:
            self.records_service.eliminar_parada_tren(
                self._parada_editando_id,
                user_id=self.usuario_actual.get("id"),
                windows_user=obtener_usuario_windows(),
            )
            self._limpiar_formulario()
            self.refrescar_tabla()
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

    def _importar_excel(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Seleccionar Planilla Excel de Paradas / Horarios por Tren",
            "",
            "Archivos Excel (*.xlsx *.xlsm)",
        )
        if not file_path:
            return
        try:
            itin_def = self.combo_itin.currentText().strip() or "H45"
            insertados = self.records_service.importar_paradas_tren_excel(
                file_path,
                itinerario_default=itin_def,
                user_id=self.usuario_actual.get("id"),
                windows_user=obtener_usuario_windows(),
            )
            self.refrescar_tabla()
            QMessageBox.information(
                self,
                "Importación Completada",
                f"Se importaron {insertados} nuevas paradas intermedias desde el archivo Excel.",
            )
        except Exception as e:
            QMessageBox.critical(self, "Error al Importar", f"No se pudo importar el archivo:\n{e}")
