"""Modelo de tabla genérico de alto rendimiento con formateo y coloreado condicional.

Hereda de QAbstractTableModel y cumple con los estándares visuales de CRONO-PCT.
"""

from typing import Any, Callable, Dict, List, Optional, Tuple
from PySide6.QtCore import QAbstractTableModel, QModelIndex, Qt
from PySide6.QtGui import QBrush, QColor, QFont


class BaseTableModel(QAbstractTableModel):
    """Modelo tabular genérico para colecciones de diccionarios u objetos."""

    def __init__(
        self,
        headers: List[str],
        column_keys: List[str],
        data: Optional[List[Dict[str, Any]]] = None,
        alignments: Optional[Dict[int, Qt.AlignmentFlag]] = None,
        conditional_colors: Optional[Callable[[Dict[str, Any], int], Optional[Tuple[QColor, QColor]]]] = None,
        parent=None,
    ):
        super().__init__(parent)
        self._headers = headers
        self._column_keys = column_keys
        self._data: List[Dict[str, Any]] = data or []
        self._alignments = alignments or {}
        self._conditional_colors = conditional_colors

    def rowCount(self, parent: QModelIndex = QModelIndex()) -> int:
        if parent.isValid():
            return 0
        return len(self._data)

    def columnCount(self, parent: QModelIndex = QModelIndex()) -> int:
        if parent.isValid():
            return 0
        return len(self._headers)

    def data(self, index: QModelIndex, role: int = Qt.DisplayRole) -> Any:
        if not index.isValid() or not (0 <= index.row() < len(self._data)):
            return None

        row_item = self._data[index.row()]
        col = index.column()
        key = self._column_keys[col]
        raw_val = row_item.get(key, "")

        if role == Qt.DisplayRole:
            if raw_val is None:
                return "-"
            return str(raw_val)

        elif role == Qt.TextAlignmentRole:
            if col in self._alignments:
                return int(self._alignments[col] | Qt.AlignVCenter)
            # Por defecto centrado si es número o fecha corta, izquierda si es texto
            if key in ("id", "dia", "legajo", "legajo_1", "legajo_2", "fecha", "fecha_cambio", "fecha_inicio", "fecha_fin", "cubierto", "toma", "deja", "hora_disponible"):
                return int(Qt.AlignCenter)
            return int(Qt.AlignLeft | Qt.AlignVCenter)

        elif role == Qt.BackgroundRole:
            if self._conditional_colors:
                colors_res = self._conditional_colors(row_item, col)
                if colors_res and colors_res[0]:
                    return QBrush(colors_res[0])

            # Coloreado específico para la columna CUBIERTO
            if key == "cubierto":
                val_cob = str(raw_val).strip()
                if val_cob == "Sí":
                    return QBrush(QColor("#E8F5E9"))  # Verde suave
                elif val_cob == "Parcial":
                    return QBrush(QColor("#FFF8E1"))  # Amarillo / Ámbar suave
                elif val_cob == "No":
                    return QBrush(QColor("#FFEBEE"))  # Rojo suave
                elif "Ciclo" in val_cob:
                    return QBrush(QColor("#F5F5F5"))  # Gris

        elif role == Qt.ForegroundRole:
            if self._conditional_colors:
                colors_res = self._conditional_colors(row_item, col)
                if colors_res and len(colors_res) > 1 and colors_res[1]:
                    return QBrush(colors_res[1])

            if key == "cubierto":
                val_cob = str(raw_val).strip()
                if val_cob == "Sí":
                    return QBrush(QColor("#2E7D32"))
                elif val_cob == "Parcial":
                    return QBrush(QColor("#F57C00"))
                elif val_cob == "No":
                    return QBrush(QColor("#C62828"))
                elif "Ciclo" in val_cob:
                    return QBrush(QColor("#616161"))

        elif role == Qt.FontRole:
            if key in ("cubierto", "id", "hora_disponible"):
                f = QFont()
                f.setBold(True)
                return f

        return None

    def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.DisplayRole) -> Any:
        if role == Qt.DisplayRole:
            if orientation == Qt.Horizontal and 0 <= section < len(self._headers):
                return self._headers[section]
            elif orientation == Qt.Vertical:
                return str(section + 1)
        elif role == Qt.TextAlignmentRole and orientation == Qt.Horizontal:
            return int(Qt.AlignCenter)
        return None

    def set_data(self, data: List[Dict[str, Any]]) -> None:
        """Reemplaza los datos de forma atómica y notifica a las vistas vinculadas."""
        self.beginResetModel()
        self._data = list(data)
        self.endResetModel()

    def get_row(self, row_idx: int) -> Optional[Dict[str, Any]]:
        if 0 <= row_idx < len(self._data):
            return self._data[row_idx]
        return None

    def get_all_data(self) -> List[Dict[str, Any]]:
        return list(self._data)
