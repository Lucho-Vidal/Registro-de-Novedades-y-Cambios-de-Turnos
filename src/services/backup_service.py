"""Servicio de copias de seguridad de la base de datos SQLite.

Utiliza la API de Backup de SQLite para copias consistentes en caliente,
gestiona políticas de retención y verificación previa de integridad en restauraciones.
"""

import os
import shutil
import sqlite3
from datetime import datetime
from pathlib import Path
from typing import List, Optional, Tuple

from src.core.logger import get_logger
from src.data.database import DatabaseManager

logger = get_logger("BackupService")

DIRECTORIO_BACKUPS = "backups"


class BackupService:
    """Gestor de copias de seguridad y restauración."""

    def __init__(self, db: DatabaseManager):
        self.db = db

    def ruta_directorio_backups(self) -> str:
        try:
            with self.db.read_connection() as conn:
                row = conn.execute("SELECT valor FROM configuracion WHERE clave='backup_directorio'").fetchone()
                if row and row[0] and os.path.isdir(str(row[0]).strip()):
                    return str(row[0]).strip()
        except Exception:
            pass
        base_dir = os.path.dirname(self.db.db_path) or "."
        ruta = os.path.join(base_dir, DIRECTORIO_BACKUPS)
        os.makedirs(ruta, exist_ok=True)
        return ruta

    def set_directorio_backups(self, ruta: str) -> None:
        """Configura una carpeta personalizada para almacenar los backups."""
        if not os.path.isdir(ruta):
            os.makedirs(ruta, exist_ok=True)
        with self.db.write_transaction() as conn:
            conn.execute(
                "INSERT INTO configuracion (clave, valor) VALUES ('backup_directorio', ?) ON CONFLICT(clave) DO UPDATE SET valor = excluded.valor;",
                (ruta,),
            )

    def listar_backups(self) -> List[Tuple[str, int, str]]:
        """Devuelve una lista de (nombre_archivo, tamaño_bytes, fecha_modificacion)."""
        directorio = self.ruta_directorio_backups()
        if not os.path.isdir(directorio):
            return []
        items = []
        for entrada in os.scandir(directorio):
            if entrada.is_file() and entrada.name.endswith(".sqlite"):
                stat = entrada.stat()
                items.append((
                    entrada.name,
                    stat.st_size,
                    datetime.fromtimestamp(stat.st_mtime).strftime("%d/%m/%Y %H:%M:%S"),
                ))
        items.sort(key=lambda x: x[0], reverse=True)
        return items

    def aplicar_retencion(self) -> int:
        with self.db.read_connection() as conn:
            row = conn.execute("SELECT valor FROM configuracion WHERE clave='backup_retencion'").fetchone()
        try:
            maximo = max(1, int(row[0])) if row else 10
        except (TypeError, ValueError):
            maximo = 10

        directorio = self.ruta_directorio_backups()
        backups = self.listar_backups()
        purgados = 0
        for nombre, _size, _fecha in backups[maximo:]:
            try:
                os.remove(os.path.join(directorio, nombre))
                purgados += 1
            except OSError as e:
                logger.warning(f"No se pudo eliminar backup antiguo {nombre}: {e}")
        return purgados

    def crear_backup(self, directorio_destino: Optional[str] = None) -> str:
        directorio = directorio_destino or self.ruta_directorio_backups()
        os.makedirs(directorio, exist_ok=True)
        base = datetime.now().strftime("%Y%m%d_%H%M%S")
        destino = os.path.join(directorio, f"backup_{base}.sqlite")
        contador = 1
        while os.path.exists(destino):
            destino = os.path.join(directorio, f"backup_{base}_{contador}.sqlite")
            contador += 1

        destino_tmp = destino + ".tmp"
        try:
            origen = sqlite3.connect(self.db.db_path)
            destino_con = sqlite3.connect(destino_tmp)
            try:
                origen.backup(destino_con)
            finally:
                destino_con.close()
                origen.close()
            os.replace(destino_tmp, destino)
            if not directorio_destino:
                self.aplicar_retencion()
            logger.info(f"Copia de seguridad creada exitosamente: {destino}")
            return destino
        except Exception as e:
            if os.path.exists(destino_tmp):
                try:
                    os.remove(destino_tmp)
                except OSError:
                    pass
            logger.error(f"Error al crear copia de seguridad: {e}")
            raise

    def restaurar_backup(self, backup_filename: str) -> None:
        directorio = self.ruta_directorio_backups()
        ruta_origen = os.path.join(directorio, backup_filename)
        if not os.path.isfile(ruta_origen):
            raise FileNotFoundError(f"El archivo de copia de seguridad {backup_filename} no existe.")

        # Verificar integridad del archivo a restaurar
        con_verif = sqlite3.connect(ruta_origen)
        try:
            check = con_verif.execute("PRAGMA integrity_check").fetchone()
            if not check or check[0].lower() != "ok":
                raise RuntimeError(f"El archivo de backup está corrupto: {check}")
        finally:
            con_verif.close()

        # Crear un backup preventivo de seguridad antes de sobreescribir
        self.crear_backup()

        # Restaurar
        origen = sqlite3.connect(ruta_origen)
        destino = sqlite3.connect(self.db.db_path)
        try:
            origen.backup(destino)
        finally:
            destino.close()
            origen.close()
        logger.info(f"Base de datos restaurada exitosamente desde {backup_filename}.")

    def eliminar_backup(self, backup_filename: str) -> None:
        directorio = self.ruta_directorio_backups()
        ruta = os.path.join(directorio, backup_filename)
        if os.path.isfile(ruta):
            os.remove(ruta)
            logger.info(f"Backup eliminado: {backup_filename}")

    def ejecutar_backup_diario_automatico(self) -> Optional[str]:
        with self.db.read_connection() as conn:
            row = conn.execute("SELECT valor FROM configuracion WHERE clave='backup_activo'").fetchone()
        activo = int(row[0]) if row and str(row[0]).isdigit() else 1
        if not activo:
            return None

        prefijo_hoy = f"backup_{datetime.now().strftime('%Y%m%d')}"
        backups = self.listar_backups()
        for nombre, _size, _fecha in backups:
            if nombre.startswith(prefijo_hoy):
                return None  # Ya existe backup de hoy

        return self.crear_backup()
