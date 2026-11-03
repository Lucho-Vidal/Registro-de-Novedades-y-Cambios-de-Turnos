"""Proxy model para filtrado y ordenamiento multicriterio en PySide6.

Permite combinar búsqueda rápida global, filtros por columnas específicas
y rangos temporales sin bloquear la interfaz de usuario.
"""

from typing import Any, Dict, Optional
from PySide6.QtCore import QModelIndex, QSortFilterProxyModel, Qt
from datetime import date, datetime

from src.domain.validators import parsear_fecha


class MultiFilterProxyModel(QSortFilterProxyModel):
    """QSortFilterProxyModel con filtros compuestos y búsqueda de texto global."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._global_search_text: str = ""
        self._column_filters: Dict[int, str] = {}  # col_idx -> exact/partial match
        self._date_col_idx: Optional[int] = None
        self._date_from: Optional[date] = None
        self._date_to: Optional[date] = None

        self.setFilterCaseSensitivity(Qt.CaseInsensitive)
        self.setSortCaseSensitivity(Qt.CaseInsensitive)
        self.setDynamicSortFilter(True)

    def _refrescar_filtro(self) -> None:
        self.invalidate()

    def set_global_search(self, text: str) -> None:
        self._global_search_text = (text or "").strip().lower()
        self._refrescar_filtro()

    def set_column_filter(self, col_idx: int, filter_value: str) -> None:
        val = (filter_value or "").strip().lower()
        if not val or val in ("todas", "todos", "todos los tipos"):
            self._column_filters.pop(col_idx, None)
        else:
            self._column_filters[col_idx] = val
        self._refrescar_filtro()

    def set_date_range(self, col_idx: int, d_from: Optional[date], d_to: Optional[date]) -> None:
        self._date_col_idx = col_idx
        self._date_from = d_from
        self._date_to = d_to
        self._refrescar_filtro()

    def clear_all_filters(self) -> None:
        self._global_search_text = ""
        self._column_filters.clear()
        self._date_from = None
        self._date_to = None
        self._refrescar_filtro()

    def filterAcceptsRow(self, source_row: int, source_parent: QModelIndex) -> bool:
        source_model = self.sourceModel()
        if not source_model:
            return True

        cols = source_model.columnCount(source_parent)

        # 1. Filtro global de texto
        if self._global_search_text:
            coincide_global = False
            for col in range(cols):
                idx = source_model.index(source_row, col, source_parent)
                val = str(source_model.data(idx, Qt.DisplayRole) or "").lower()
                if self._global_search_text in val:
                    coincide_global = True
                    break
            if not coincide_global:
                return False

        # 2. Filtros específicos por columna
        for col_idx, expected_val in self._column_filters.items():
            if col_idx < cols:
                idx = source_model.index(source_row, col_idx, source_parent)
                cell_val = str(source_model.data(idx, Qt.DisplayRole) or "").lower()
                if expected_val not in cell_val:
                    return False

        # 3. Filtro de rango de fecha
        if self._date_col_idx is not None and (self._date_from or self._date_to):
            if self._date_col_idx < cols:
                idx = source_model.index(source_row, self._date_col_idx, source_parent)
                raw_fecha = str(source_model.data(idx, Qt.DisplayRole) or "").strip()
                try:
                    dt = parsear_fecha(raw_fecha, obligatorio=False)
                    if dt:
                        d_val = dt.date()
                        if self._date_from and d_val < self._date_from:
                            return False
                        if self._date_to and d_val > self._date_to:
                            return False
                except Exception:
                    pass

        return True

    def lessThan(self, left: QModelIndex, right: QModelIndex) -> bool:
        source_model = self.sourceModel()
        left_val = source_model.data(left, Qt.DisplayRole)
        right_val = source_model.data(right, Qt.DisplayRole)

        if left_val is None:
            return True
        if right_val is None:
            return False

        # Comparación numérica si aplica
        try:
            num_left = float(left_val)
            num_right = float(right_val)
            return num_left < num_right
        except (ValueError, TypeError):
            pass

        # Comparación de fechas si aplica (DD/MM/YYYY o YYYY-MM-DD)
        try:
            dt_left = parsear_fecha(left_val)
            dt_right = parsear_fecha(right_val)
            if dt_left and dt_right:
                return dt_left < dt_right
        except Exception:
            pass

        return str(left_val).lower() < str(right_val).lower()
