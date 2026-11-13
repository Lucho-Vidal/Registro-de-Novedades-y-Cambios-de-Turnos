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
PATH_BASE_FILE = APP_DATA_DIR / "config.json"


def resource_path(rel_path: str | Path) -> Path:
    """Resuelve la ruta absoluta de un recurso, compatible con desarrollo y PyInstaller."""
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        base = Path(sys._MEIPASS)
    else:
        base = BASE_DIR
    return base / rel_path


def obtener_ruta_base_datos() -> tuple[Path, Path | None]:
    """Obtiene la ruta a la base de datos SQLite y opcionalmente el XLSX asociado.
    
    Lee desde 'config.json' ('database_path') o 'reno_config.json'.
    Si no existe o es relativa, se resuelve contra APP_DATA_DIR.
    """
    ruta_txt = ""
    if PATH_BASE_FILE.exists():
        try:
            contenido = PATH_BASE_FILE.read_text(encoding="utf-8").strip()
            if contenido:
                if contenido.startswith("{"):
                    data = json.loads(contenido)
                    ruta_txt = str(data.get("database_path") or data.get("path_base") or "").strip()
                else:
                    ruta_txt = contenido
        except Exception:
            pass

    # Si no había config.json pero existe el archivo legado path_base, migrarlo y eliminarlo
    legacy_file = APP_DATA_DIR / "path_base"
    if not ruta_txt and legacy_file.exists():
        try:
            leg_val = legacy_file.read_text(encoding="utf-8").strip()
            if leg_val:
                ruta_txt = leg_val
                guardar_ruta_base_datos(leg_val)
            legacy_file.unlink(missing_ok=True)
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
    """Guarda la ruta de la base de datos en config.json y en reno_config.json."""
    ruta_str = str(nueva_ruta).strip()
    exito = True
    try:
        data: dict[str, Any] = {}
        if PATH_BASE_FILE.exists():
            try:
                contenido = PATH_BASE_FILE.read_text(encoding="utf-8").strip()
                if contenido.startswith("{"):
                    data = json.loads(contenido)
            except Exception:
                data = {}
        data["database_path"] = ruta_str
        with open(PATH_BASE_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        (APP_DATA_DIR / "path_base").unlink(missing_ok=True)
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


def cargar_variables_entorno(sobreescribir: bool = True) -> dict[str, str]:
    """Carga variables desde el archivo .env si existe en APP_DATA_DIR o BASE_DIR."""
    env_vars: dict[str, str] = {}
    candidatos: list[Path] = []
    for p in (APP_DATA_DIR / ".env", BASE_DIR / ".env"):
        if p not in candidatos:
            candidatos.append(p)

    for env_file in candidatos:
        if env_file.exists():
            try:
                for raw_line in env_file.read_text(encoding="utf-8-sig").splitlines():
                    line = raw_line.strip()
                    if not line or line.startswith("#") or "=" not in line:
                        continue
                    k, v = line.split("=", 1)
                    k = k.strip()
                    v = v.strip().strip('"').strip("'")
                    if k:
                        env_vars[k] = v
                        if sobreescribir or not os.environ.get(k):
                            os.environ[k] = v
                break
            except Exception:
                pass
    return env_vars


def cargar_configuracion_app() -> dict[str, Any]:
    """Carga reno_config.json con valores por defecto y soporte de .env."""
    env_vars = cargar_variables_entorno()
    sb_url_env = (env_vars.get("SUPABASE_URL") or os.environ.get("SUPABASE_URL") or "").strip().rstrip("/")
    sb_anon_env = (env_vars.get("SUPABASE_ANON_KEY") or os.environ.get("SUPABASE_ANON_KEY") or "").strip()
    sb_serv_env = (
        env_vars.get("SUPABASE_SERVICE_KEY")
        or env_vars.get("SUPABASE_SERVICE_ROLE_KEY")
        or os.environ.get("SUPABASE_SERVICE_KEY")
        or os.environ.get("SUPABASE_SERVICE_ROLE_KEY")
        or ""
    ).strip()
    sb_bucket_env = (
        env_vars.get("SUPABASE_STORAGE_BUCKET")
        or os.environ.get("SUPABASE_STORAGE_BUCKET")
        or "updates"
    ).strip() or "updates"
    updater_url_env = (env_vars.get("RENO_UPDATES_URL") or os.environ.get("RENO_UPDATES_URL") or "").strip()

    defaults: dict[str, Any] = {
        "supabase_url": sb_url_env,
        "supabase_anon_key": sb_anon_env,
        "supabase_service_role_key": sb_serv_env,
        "supabase_storage_bucket": sb_bucket_env,
        "updater_url": updater_url_env,
        "updater_auto_check": True,
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
                for k, v in data.items():
                    if v or k not in defaults:
                        defaults[k] = v
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

