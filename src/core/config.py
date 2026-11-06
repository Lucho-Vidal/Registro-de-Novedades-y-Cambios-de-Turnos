"""Configuraciones transversales y resolución de rutas del sistema RENO."""

from __future__ import annotations

import getpass
import json
import os
import sys
from pathlib import Path
from typing import Any

APP_NAME = "RENO"
VERSION = "2.0.0"
APP_VERSION = VERSION

BASE_DIR = Path(__file__).resolve().parent.parent.parent


def obtener_directorio_aplicacion() -> Path:
    """Retorna el directorio base persistente de la aplicación.
    
    Si corre empaquetado (.exe de PyInstaller), retorna la carpeta donde reside el ejecutable.
    Si corre en entorno de desarrollo, retorna la raíz del proyecto.
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return BASE_DIR


APP_DATA_DIR = obtener_directorio_aplicacion()

DOTACIONES = ["Todas", "PC", "LLV", "TY", "LP", "OA", "K5", "RE", "CÑ", "AK"]


def obtener_usuario_windows() -> str:
    """Retorna el nombre de usuario activo de la sesión del sistema operativo (Windows/OS)."""
    try:
        user = os.environ.get("USERNAME") or os.environ.get("USER") or getpass.getuser()
        return str(user).strip() if user else "Desconocido"
    except Exception:
        return "Desconocido"


ESPECIALIDADES_EMPLEADO = (
    "Conductor Eléctrico",
    "Conductor Diesel",
    "Ayudante Habilitado",
    "Ayudante Conductor",
    "Guarda Eléctrico",
    "Guarda Diesel",
)

DIAS_SEMANA = ("Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo")

CONFIG_JSON_PATH = APP_DATA_DIR / "reno_config.json"
PATH_BASE_FILE = APP_DATA_DIR / "path_base"


def resource_path(rel_path: str | Path) -> Path:
    """Resuelve la ruta absoluta de un recurso, compatible con desarrollo y PyInstaller."""
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        base = Path(sys._MEIPASS)
    else:
        base = BASE_DIR
    return base / rel_path


def obtener_ruta_base_datos() -> tuple[Path, Path | None]:
    """Obtiene la ruta a la base de datos SQLite y opcionalmente el XLSX asociado.
    
    Lee desde 'path_base' o 'database_path' en reno_config.json.
    Si no existe o es relativa, se resuelve contra APP_DATA_DIR.
    """
    ruta_txt = ""
    if PATH_BASE_FILE.exists():
        try:
            with open(PATH_BASE_FILE, "r", encoding="utf-8") as f:
                contenido = f.read().strip()
                if contenido:
                    ruta_txt = contenido
        except Exception:
            pass

    if not ruta_txt and CONFIG_JSON_PATH.exists():
        try:
            with open(CONFIG_JSON_PATH, "r", encoding="utf-8") as f:
                cfg = json.load(f)
                ruta_txt = str(cfg.get("database_path", "")).strip()
        except Exception:
            pass

    if not ruta_txt:
        ruta_txt = "RENO.sqlite"

    path_obj = Path(ruta_txt)
    if not path_obj.is_absolute():
        path_obj = APP_DATA_DIR / path_obj

    if path_obj.suffix.lower() == ".xlsx":
        xlsx_path = path_obj
        sqlite_path = path_obj.with_suffix(".sqlite")
    else:
        sqlite_path = path_obj if path_obj.suffix.lower() == ".sqlite" else path_obj.with_suffix(".sqlite")
        xlsx_candidato = sqlite_path.with_suffix(".xlsx")
        xlsx_path = xlsx_candidato if xlsx_candidato.exists() else None

    return sqlite_path, xlsx_path


def guardar_ruta_base_datos(nueva_ruta: str | Path) -> bool:
    """Guarda la ruta de la base de datos tanto en path_base como en reno_config.json."""
    ruta_str = str(nueva_ruta).strip()
    exito = True
    try:
        with open(PATH_BASE_FILE, "w", encoding="utf-8") as f:
            f.write(ruta_str)
    except Exception:
        exito = False

    try:
        cfg = cargar_configuracion_app()
        cfg["database_path"] = ruta_str
        if not guardar_configuracion_app(cfg):
            exito = False
    except Exception:
        exito = False

    return exito


def cargar_configuracion_app() -> dict[str, Any]:
    """Carga reno_config.json con valores por defecto."""
    defaults: dict[str, Any] = {
        "supabase_url": "",
        "supabase_anon_key": "",
        "sqlite_journal_mode": "WAL",  # 'WAL' o 'DELETE'
        "autosync_habilitado": False,
        "autosync_intervalo_min": 5,
        "autorefresco_habilitado": True,
        "autorefresco_intervalo_seg": 30,
        "proxy_mode": "direct",  # 'direct', 'manual', 'windows_sso'
        "proxy_host": "",
        "proxy_port": 8080,
        "proxy_user": "",
        "proxy_pass": "",
        "proxy_sso_server": "",
        "proxy_auth": "manual_fijo",  # 'manual_fijo', 'sesion_reno', 'windows_sspi'
        "tema_visual": "claro",
    }
    if CONFIG_JSON_PATH.exists():
        try:
            with open(CONFIG_JSON_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
                defaults.update(data)
        except Exception:
            pass
    return defaults


def guardar_configuracion_app(config: dict[str, Any]) -> bool:
    """Guarda la configuración en reno_config.json."""
    try:
        with open(CONFIG_JSON_PATH, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
        return True
    except Exception:
        return False


cargar_configuracion_json = cargar_configuracion_app
guardar_configuracion_json = guardar_configuracion_app

