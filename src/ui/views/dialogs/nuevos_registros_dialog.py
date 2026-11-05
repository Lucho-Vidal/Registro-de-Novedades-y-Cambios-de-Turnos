"""Diálogo modal para visualizar y abrir la lista de registros nuevos pendientes."""

from typing import Any, Callable, Dict, List, Optional
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
)


class NuevosRegistrosDialog(QDialog):
    """Muestra los nuevos registros operativos cargados desde el último acceso.
    
    Permite filtrar por texto, seleccionar cualquier novedad o cambio de turno y
    abrir directamente su formulario operativo mediante doble clic o el botón de acción.
    """

    def __init__(
        self,
        novedades: List[Dict[str, Any]],
        cambios: List[Dict[str, Any]],
        on_abrir_novedad: Callable[[int], None],
        on_abrir_cambio: Callable[[int], None],
        on_ir_a_grilla: Optional[Callable[[], None]] = None,
        parent=None,
    ):
        super().__init__(parent)
        self.novedades = novedades
        self.cambios = cambios
        self.on_abrir_novedad = on_abrir_novedad
        self.on_abrir_cambio = on_abrir_cambio
        self.on_ir_a_grilla = on_ir_a_grilla

        total = len(novedades) + len(cambios)
        self.setWindowTitle(f"🔔 Registros Nuevos Operativos ({total})")
        self.resize(880, 520)
        self.setWindowFlags((self.windowFlags() & ~Qt.WindowContextHelpButtonHint) | Qt.WindowCloseButtonHint)

        self._registros_planos: List[Dict[str, Any]] = []
        self._init_ui()
        self._cargar_registros()

    def _init_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(10)

        # Encabezado
        lbl_info = QLabel("<b>Registros ingresados por otros operadores desde su último acceso:</b>")
        lbl_info.setStyleSheet("font-size: 10pt;")
        layout.addWidget(lbl_info)

        lbl_sub = QLabel("Haga doble clic en una fila o presione <i>'Abrir Formulario'</i> para ver todos los detalles.")
        lbl_sub.setStyleSheet("font-size: 8.5pt; color: #64748B;")
        layout.addWidget(lbl_sub)

        # Buscador / Filtro rápido
        h_filtro = QHBoxLayout()
        h_filtro.addWidget(QLabel("🔍 Filtrar:"))
        self.txt_filtro = QLineEdit()
        self.txt_filtro.setPlaceholderText("Buscar por legajo, agente, novedad, fecha o usuario...")
        self.txt_filtro.textChanged.connect(self._filtrar_tabla)
        h_filtro.addWidget(self.txt_filtro)
        layout.addLayout(h_filtro)

        # Grilla de registros
        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(["TIPO", "ID", "DETALLE / PERSONAL", "FECHA REGISTRO", "OPERADOR"])
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setAlternatingRowColors(True)

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.Stretch)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)

        self.table.doubleClicked.connect(lambda _: self._abrir_seleccionado())
        layout.addWidget(self.table)

        # Botones de acción
        btn_box = QHBoxLayout()

        self.btn_grilla = QPushButton("📋 Ir a Novedades y Cambios")
        self.btn_grilla.setObjectName("btnSecondary")
        self.btn_grilla.clicked.connect(self._on_ir_grilla)
        btn_box.addWidget(self.btn_grilla)

        btn_box.addStretch()

        self.btn_cerrar = QPushButton("Cerrar")
        self.btn_cerrar.setObjectName("btnSecondary")
        self.btn_cerrar.clicked.connect(self.reject)
        btn_box.addWidget(self.btn_cerrar)

        self.btn_abrir = QPushButton("📝 Abrir Formulario")
        self.btn_abrir.setDefault(True)
        self.btn_abrir.clicked.connect(self._abrir_seleccionado)
        btn_box.addWidget(self.btn_abrir)

        layout.addLayout(btn_box)

    def _cargar_registros(self) -> None:
        self._registros_planos.clear()

        # Aplanar novedades
        for n in self.novedades:
            leg = n.get("legajo", "")
            nom = n.get("apellidos_nombres", "")
            nov_desc = n.get("novedad", "")
            detalle = f"{nov_desc} - {nom} (Leg. {leg})" if nom else nov_desc
            usr = n.get("usuario_windows") or (f"Usuario #{n['usuario_id']}" if n.get("usuario_id") else "-")
            self._registros_planos.append({
                "tipo": "novedad",
                "tipo_label": "📝 Novedad",
                "id": n["id"],
                "detalle": detalle,
                "registrado_en": str(n.get("registrado_en", "")),
                "operador": str(usr),
            })

        # Aplanar cambios de turno
        for c in self.cambios:
            l1 = c.get("legajo_1", "")
            l2 = c.get("legajo_2", "")
            fc = c.get("fecha_cambio", "")
            detalle = f"Permuta: Leg. {l1} ↔ Leg. {l2}"
            if fc:
                detalle += f" (Fecha: {fc})"
            usr = c.get("usuario_windows") or (f"Usuario #{c['usuario_id']}" if c.get("usuario_id") else "-")
            self._registros_planos.append({
                "tipo": "cambio",
                "tipo_label": "🔄 Cambio de Turno",
                "id": c["id"],
                "detalle": detalle,
                "registrado_en": str(c.get("registrado_en", "")),
                "operador": str(usr),
            })

        # Ordenar por id desc
        self._registros_planos.sort(key=lambda r: (r["registrado_en"], r["id"]), reverse=True)
        self._render_tabla(self._registros_planos)

    def _render_tabla(self, items: List[Dict[str, Any]]) -> None:
        self.table.setRowCount(len(items))
        for row_idx, item in enumerate(items):
            it_tipo = QTableWidgetItem(item["tipo_label"])
            it_id = QTableWidgetItem(f"#{item['id']}")
            it_id.setTextAlignment(Qt.AlignCenter)
            it_det = QTableWidgetItem(item["detalle"])
            it_fecha = QTableWidgetItem(item["registrado_en"])
            it_fecha.setTextAlignment(Qt.AlignCenter)
            it_op = QTableWidgetItem(item["operador"])

            # Guardar metadatos en el primer item
            it_tipo.setData(Qt.UserRole, item)

            self.table.setItem(row_idx, 0, it_tipo)
            self.table.setItem(row_idx, 1, it_id)
            self.table.setItem(row_idx, 2, it_det)
            self.table.setItem(row_idx, 3, it_fecha)
            self.table.setItem(row_idx, 4, it_op)

        if len(items) > 0:
            self.table.selectRow(0)

    def _filtrar_tabla(self, texto: str) -> None:
        txt = texto.strip().lower()
        if not txt:
            for r in range(self.table.rowCount()):
                self.table.setRowHidden(r, False)
            return

        for r in range(self.table.rowCount()):
            coincide = False
            for col in range(self.table.columnCount()):
                item = self.table.item(r, col)
                if item and txt in item.text().lower():
                    coincide = True
                    break
            self.table.setRowHidden(r, not coincide)

    def _abrir_seleccionado(self) -> None:
        row = self.table.currentRow()
        if row < 0:
            return
        item_tipo = self.table.item(row, 0)
        if not item_tipo:
            return
        meta = item_tipo.data(Qt.UserRole)
        if not meta:
            return

        tipo = meta.get("tipo")
        reg_id = meta.get("id")

        if tipo == "novedad":
            self.on_abrir_novedad(reg_id)
        elif tipo == "cambio":
            self.on_abrir_cambio(reg_id)

    def _on_ir_grilla(self) -> None:
        self.accept()
        if self.on_ir_a_grilla:
            self.on_ir_a_grilla()
