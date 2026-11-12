"""Gestión de QSettings para geometrías, columnas y filtros de usuario."""

from __future__ import annotations

try:
    from PySide6.QtCore import QByteArray, QSettings
    from PySide6.QtWidgets import QHeaderView, QWidget
except ImportError:
    QByteArray = None
    QSettings = None
    QHeaderView = None
    QWidget = None


class AppSettings:
    """Envoltorio amigable para almacenar y recuperar estados de la UI."""

    def __init__(self) -> None:
        self._qs = QSettings("TrenesArgentinos", "RENO_v2")

    def guardar_geometria(self, widget: QWidget, nombre_clave: str) -> None:
        self._qs.setValue(f"geom_{nombre_clave}", widget.saveGeometry())

    def restaurar_geometria(self, widget: QWidget, nombre_clave: str) -> bool:
        geom = self._qs.value(f"geom_{nombre_clave}")
        if geom and isinstance(geom, QByteArray):
            return widget.restoreGeometry(geom)
        return False

    def guardar_columnas_header(self, header: QHeaderView, nombre_tabla: str) -> None:
        self._qs.setValue(f"header_{nombre_tabla}", header.saveState())

    def restaurar_columnas_header(self, header: QHeaderView, nombre_tabla: str) -> bool:
        state = self._qs.value(f"header_{nombre_tabla}")
        if state and isinstance(state, QByteArray):
            return header.restoreState(state)
        return False

    def set_valor(self, clave: str, valor: Any) -> None:
        self._qs.setValue(clave, valor)
        try:
            self._qs.sync()
        except Exception:
            pass

    def sync(self) -> None:
        try:
            self._qs.sync()
        except Exception:
            pass

    def get_valor(self, clave: str, default: Any = None) -> Any:
        return self._qs.value(clave, default)

    def get_bool(self, clave: str, default: bool = False) -> bool:
        val = self._qs.value(clave, default)
        if isinstance(val, bool):
            return val
        if isinstance(val, str):
            return val.strip().lower() in ("true", "1", "yes", "si", "sí")
        return bool(val)

    def set_window_geometry(self, nombre_clave: str, geometry: QByteArray) -> None:
        self._qs.setValue(f"geom_{nombre_clave}", geometry)

    def get_window_geometry(self, nombre_clave: str) -> QByteArray | None:
        geom = self._qs.value(f"geom_{nombre_clave}")
        return geom if isinstance(geom, QByteArray) else None

