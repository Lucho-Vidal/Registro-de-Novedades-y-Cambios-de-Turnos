"""Copias de seguridad de la base SQLite.

Las copias se guardan en la carpeta "backups" junto a la base de datos,
con el formato backup_YYYYMMDD_HHMMSS.sqlite. La retención (número de
copias a conservar) y la activación del backup automático se configuran
en la tabla "configuracion" (claves backup_activo y backup_retencion).
"""

import os
import shutil
import sqlite3
from datetime import datetime

DIRECTORIO_BACKUPS = "backups"


def _ruta_backups(store):
    return os.path.join(os.path.dirname(store.database_path) or ".", DIRECTORIO_BACKUPS)


def _aplicar_retencion(store):
    try:
        maximo = max(1, int(store.get_configuracion("backup_retencion", "10")))
    except (TypeError, ValueError):
        maximo = 10
    directorio = _ruta_backups(store)
    backups = listar_backups(store)
    for nombre, _tamano in backups[maximo:]:
        try:
            os.remove(os.path.join(directorio, nombre))
        except OSError:
            pass


def crear_backup(store):
    """Crea una copia consistente de la base usando la API de backup de SQLite.

    Returns:
        Ruta de la copia creada.
    """
    directorio = _ruta_backups(store)
    os.makedirs(directorio, exist_ok=True)
    base = datetime.now().strftime("%Y%m%d_%H%M%S")
    destino = os.path.join(directorio, f"backup_{base}.sqlite")
    contador = 1
    while os.path.exists(destino):
        destino = os.path.join(directorio, f"backup_{base}_{contador}.sqlite")
        contador += 1
    destino_tmp = destino + ".tmp"
    try:
        origen = sqlite3.connect(store.database_path)
        destino_con = sqlite3.connect(destino_tmp)
        try:
            origen.backup(destino_con)
        finally:
            destino_con.close()
            origen.close()
        os.replace(destino_tmp, destino)
    except Exception:
        if os.path.exists(destino_tmp):
            os.remove(destino_tmp)
        raise
    _aplicar_retencion(store)
    return destino


def listar_backups(store):
    """Lista las copias existentes (más recientes primero) como (nombre, tamaño)."""
    directorio = _ruta_backups(store)
    if not os.path.isdir(directorio):
        return []
    backups = []
    for nombre in os.listdir(directorio):
        ruta = os.path.join(directorio, nombre)
        if os.path.isfile(ruta) and nombre.startswith("backup_") and nombre.endswith(".sqlite"):
            backups.append((nombre, os.path.getsize(ruta)))
    backups.sort(key=lambda item: os.path.getmtime(os.path.join(directorio, item[0])), reverse=True)
    return backups


def hay_backup_hoy(store):
    hoy = datetime.now().strftime("%Y%m%d")
    for nombre, _tamano in listar_backups(store):
        if nombre.startswith(f"backup_{hoy}"):
            return True
    return False


def backup_automatico_si_corresponde(store):
    """Crea una copia diaria si está activa y aún no existe una para hoy."""
    if str(store.get_configuracion("backup_activo", "1")) != "1":
        return None
    if hay_backup_hoy(store):
        return None
    try:
        return crear_backup(store)
    except Exception:
        return None


def restaurar_backup(store, nombre):
    """Reemplaza la base activa por una copia. Devuelve True si se restauró."""
    origen = os.path.join(_ruta_backups(store), nombre)
    if not os.path.isfile(origen):
        raise FileNotFoundError("La copia de seguridad no existe.")
    temporal = store.database_path + ".restaurar"
    shutil.copy2(origen, temporal)
    os.replace(temporal, store.database_path)
    return True