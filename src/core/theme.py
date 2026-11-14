"""Módulo unificado para la gestión, persistencia y aplicación del tema visual (claro / oscuro)."""

from __future__ import annotations

import json
import logging
import os
from pathlib import Path
from typing import Any

from PySide6.QtCore import Qt
from PySide6.QtGui import QIcon, QPixmap
from PySide6.QtWidgets import QApplication

from .config import (
    CONFIG_JSON_PATH,
    cargar_configuracion_json,
    guardar_configuracion_json,
    obtener_rutas_icono_candidatas,
    resource_path,
)
from .settings import AppSettings

logger = logging.getLogger("reno.theme")

TEMA_CLARO = "claro"
TEMA_OSCURO = "oscuro"

_TEMA_ACTUAL: str | None = None


def obtener_tema_visual(db: Any = None) -> str:
    """Obtiene el tema visual persistido con caché en memoria y triple redundancia:

    0. Caché en memoria / variable de entorno de proceso (cuando db es None).
    1. Base de datos SQLite (tabla configuracion).
    2. Archivo de configuración local reno_config.json (lectura directa sin default forzado).
    3. QSettings de usuario (Windows Registry / AppData).
    """
    global _TEMA_ACTUAL

    # 0. Si no se pasa db (ej. durante paintEvent de gráficos), priorizar el estado activo en memoria
    if db is None:
        if _TEMA_ACTUAL in (TEMA_CLARO, TEMA_OSCURO):
            return _TEMA_ACTUAL
        tema_env = os.environ.get("RENO_TEMA_VISUAL")
        if tema_env in (TEMA_CLARO, TEMA_OSCURO):
            _TEMA_ACTUAL = tema_env
            return tema_env

    # 1. Consultar base de datos SQLite si está disponible
    if db is not None:
        try:
            with db.read_connection() as conn:
                row = conn.execute("SELECT valor FROM configuracion WHERE clave = 'tema_visual'").fetchone()
                if row and row[0] and str(row[0]).strip().lower() in (TEMA_CLARO, TEMA_OSCURO):
                    res = str(row[0]).strip().lower()
                    _TEMA_ACTUAL = res
                    os.environ["RENO_TEMA_VISUAL"] = res
                    return res
        except Exception as e:
            logger.debug("No se pudo leer tema de SQLite: %s", e)

    # 2. Consultar reno_config.json directamente (sin pisar con el default 'claro' si la clave no existe)
    try:
        if CONFIG_JSON_PATH.exists():
            with open(CONFIG_JSON_PATH, "r", encoding="utf-8") as f:
                raw_cfg = json.load(f)
            tema_json = raw_cfg.get("tema_visual")
            if tema_json and str(tema_json).strip().lower() in (TEMA_CLARO, TEMA_OSCURO):
                res = str(tema_json).strip().lower()
                _TEMA_ACTUAL = res
                os.environ["RENO_TEMA_VISUAL"] = res
                return res
    except Exception as e:
        logger.debug("No se pudo leer tema de config.json: %s", e)

    # 3. Consultar AppSettings (QSettings)
    try:
        settings = AppSettings()
        tema_qs = settings.get_valor("tema_visual", None)
        if tema_qs and str(tema_qs).strip().lower() in (TEMA_CLARO, TEMA_OSCURO):
            res = str(tema_qs).strip().lower()
            _TEMA_ACTUAL = res
            os.environ["RENO_TEMA_VISUAL"] = res
            return res
    except Exception as e:
        logger.debug("No se pudo leer tema de AppSettings: %s", e)

    # 4. Variable de entorno de proceso
    tema_env = os.environ.get("RENO_TEMA_VISUAL")
    if tema_env and tema_env in (TEMA_CLARO, TEMA_OSCURO):
        _TEMA_ACTUAL = tema_env
        return tema_env

    _TEMA_ACTUAL = TEMA_CLARO
    os.environ["RENO_TEMA_VISUAL"] = TEMA_CLARO
    return TEMA_CLARO


def guardar_tema_visual(tema: str, db: Any = None) -> None:
    """Guarda de forma atómica e inmediata el tema visual en todos los niveles de persistencia."""
    global _TEMA_ACTUAL
    tema_norm = TEMA_OSCURO if str(tema).strip().lower() == TEMA_OSCURO else TEMA_CLARO
    _TEMA_ACTUAL = tema_norm
    os.environ["RENO_TEMA_VISUAL"] = tema_norm

    # 1. Guardar en SQLite si hay conexión
    if db is not None:
        try:
            with db.write_transaction() as conn:
                conn.execute(
                    "INSERT INTO configuracion (clave, valor) VALUES ('tema_visual', ?) ON CONFLICT(clave) DO UPDATE SET valor = excluded.valor;",
                    (tema_norm,),
                )
        except Exception as e:
            logger.warning("Fallo al persistir tema en SQLite: %s", e)

    # 2. Guardar en reno_config.json
    try:
        cfg = cargar_configuracion_json()
        cfg["tema_visual"] = tema_norm
        guardar_configuracion_json(cfg)
    except Exception as e:
        logger.warning("Fallo al persistir tema en reno_config.json: %s", e)

    # 3. Guardar en QSettings con sync forzado
    try:
        settings = AppSettings()
        settings.set_valor("tema_visual", tema_norm)
        settings.sync()
    except Exception as e:
        logger.warning("Fallo al persistir tema en AppSettings: %s", e)


def resolver_archivo_qss(tema: str) -> Path:
    """Resuelve la ruta absoluta al archivo QSS correspondiente al tema."""
    qss_name = "crono_dark.qss" if tema == TEMA_OSCURO else "crono_style.qss"
    ruta_relativa = Path("src") / "ui" / "theme" / qss_name
    qss_path = resource_path(ruta_relativa)
    if not qss_path.exists():
        # Fallback relativo a este módulo
        qss_path = Path(__file__).resolve().parent.parent / "ui" / "theme" / qss_name
    return qss_path


def aplicar_tema_visual(tema: str | None = None, db: Any = None) -> str:
    """Carga y aplica el QSS a QApplication de manera inmediata."""
    global _TEMA_ACTUAL
    if tema is None:
        tema = obtener_tema_visual(db)
        _TEMA_ACTUAL = tema
        os.environ["RENO_TEMA_VISUAL"] = tema
    else:
        guardar_tema_visual(tema, db)
        tema = _TEMA_ACTUAL or TEMA_CLARO

    qss_path = resolver_archivo_qss(tema)
    if qss_path.exists():
        try:
            with open(qss_path, "r", encoding="utf-8") as f:
                qss_content = f.read()
                app = QApplication.instance()
                if app:
                    app.setStyleSheet(qss_content)
        except Exception as e:
            logger.warning("No se pudo cargar la hoja de estilos QSS '%s': %s", qss_path, e)

    return tema


def cargar_icono_aplicacion(prefer_ico: bool = False) -> QIcon:
    """Carga el QIcon de RENO verificando candidatos reales y pixmap no nulo."""
    for candidato in obtener_rutas_icono_candidatas(prefer_ico=prefer_ico):
        try:
            if not candidato.is_file():
                continue
            icono = QIcon(str(candidato))
            if not icono.isNull() and not icono.pixmap(32, 32).isNull():
                return icono
        except Exception as e:
            logger.debug("No se pudo cargar el icono '%s': %s", candidato, e)
    return QIcon()


def cargar_pixmap_icono(size: int = 48) -> QPixmap | None:
    """Carga el logotipo oficial de RENO escalado suavemente como QPixmap."""
    for candidato in obtener_rutas_icono_candidatas(prefer_ico=False):
        try:
            if not candidato.is_file():
                continue
            pix = QPixmap(str(candidato))
            if not pix.isNull():
                return pix.scaled(size, size, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        except Exception as e:
            logger.debug("No se pudo cargar el pixmap '%s': %s", candidato, e)
    return None


def aplicar_icono_aplicacion(target: Any = None, prefer_ico: bool = False) -> bool:
    """Asigna el icono oficial de RENO a QApplication y/o a una ventana o diálogo."""
    icono = cargar_icono_aplicacion(prefer_ico=prefer_ico)
    if icono.isNull():
        return False

    app = QApplication.instance()
    if target is None:
        if app is not None:
            app.setWindowIcon(icono)
            return True
        return False

    try:
        if hasattr(target, "setWindowIcon"):
            target.setWindowIcon(icono)
        if app is not None and app.windowIcon().isNull():
            app.setWindowIcon(icono)
        return True
    except Exception as e:
        logger.debug("No se pudo aplicar icono al objetivo '%s': %s", target, e)
        return False

