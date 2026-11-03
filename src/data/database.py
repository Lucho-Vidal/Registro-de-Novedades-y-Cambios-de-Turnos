"""Gestor SQLite con modos duales WAL / DELETE y locking seguro en red local."""

from __future__ import annotations

import logging
import os
import sqlite3
import time
from datetime import datetime
from contextlib import contextmanager
from pathlib import Path
from typing import Iterator

logger = logging.getLogger("reno.data.database")

try:
    import msvcrt
except ImportError:
    msvcrt = None

try:
    import fcntl
except ImportError:
    fcntl = None


class DatabaseBusyError(RuntimeError):
    """La base de datos se encuentra ocupada por otra terminal."""
    pass


def _try_lock(file_handle) -> bool:
    if msvcrt:
        file_handle.seek(0)
        try:
            msvcrt.locking(file_handle.fileno(), msvcrt.LK_NBLCK, 1)
            return True
        except OSError:
            return False
    if fcntl:
        try:
            fcntl.flock(file_handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
            return True
        except OSError:
            return False
    return True


def _unlock(file_handle) -> None:
    if msvcrt:
        file_handle.seek(0)
        try:
            msvcrt.locking(file_handle.fileno(), msvcrt.LK_UNLCK, 1)
        except OSError:
            pass
    elif fcntl:
        try:
            fcntl.flock(file_handle.fileno(), fcntl.LOCK_UN)
        except OSError:
            pass


class DatabaseManager:
    """Conexión y transacciones sobre SQLite con soporte de modos WAL y DELETE."""

    def __init__(self, db_path: Path | str, journal_mode: str = "WAL", lock_timeout: int = 15) -> None:
        self.db_path = Path(db_path).resolve()
        self.lock_path = Path(f"{self.db_path}.lock")
        self.journal_mode = journal_mode.upper() if journal_mode else "WAL"
        self.lock_timeout = lock_timeout
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._last_mtime: float = 0.0
        self._last_data_version: int = 0

    @staticmethod
    def now() -> str:
        """Devuelve la fecha y hora actual en formato ISO estándar 'YYYY-MM-DD HH:MM:SS'."""
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def connect(self) -> sqlite3.Connection:
        conn = sqlite3.connect(
            str(self.db_path),
            timeout=float(self.lock_timeout),
            isolation_level=None,
            check_same_thread=False,
        )
        conn.row_factory = sqlite3.Row

        conn.execute("PRAGMA foreign_keys = ON;")
        conn.execute("PRAGMA temp_store = MEMORY;")

        if self.journal_mode == "WAL":
            try:
                conn.execute("PRAGMA journal_mode = WAL;")
            except Exception:
                conn.execute("PRAGMA journal_mode = DELETE;")
            conn.execute("PRAGMA synchronous = NORMAL;")
            conn.execute("PRAGMA busy_timeout = 5000;")
        else:
            try:
                conn.execute("PRAGMA journal_mode = DELETE;")
            except Exception:
                pass
            conn.execute("PRAGMA synchronous = FULL;")
            conn.execute(f"PRAGMA busy_timeout = {self.lock_timeout * 1000};")

        return conn

    @contextmanager
    def read_connection(self) -> Iterator[sqlite3.Connection]:
        conn = self.connect()
        try:
            yield conn
        finally:
            conn.close()

    @contextmanager
    def write_transaction(self) -> Iterator[sqlite3.Connection]:
        lock_file = None
        # En modo DELETE usamos el archivo .lock para blindar accesos sobre carpetas SMB
        if self.journal_mode == "DELETE":
            started = time.monotonic()
            while True:
                try:
                    lock_file = open(self.lock_path, "a+b")
                    if os.path.getsize(self.lock_path) == 0:
                        lock_file.write(b"0")
                        lock_file.flush()
                    if _try_lock(lock_file):
                        break
                    raise OSError("Lock ocupado")
                except OSError:
                    if lock_file:
                        lock_file.close()
                        lock_file = None
                    if time.monotonic() - started >= self.lock_timeout:
                        raise DatabaseBusyError("La base de datos está siendo utilizada por otra terminal.")
                    time.sleep(0.12)

        conn = self.connect()
        try:
            conn.execute("BEGIN IMMEDIATE;")
            yield conn
            if conn.in_transaction:
                conn.execute("COMMIT;")
        except Exception:
            if conn.in_transaction:
                try:
                    conn.execute("ROLLBACK;")
                except Exception:
                    pass
            raise
        finally:
            conn.close()
            if lock_file is not None:
                _unlock(lock_file)
                lock_file.close()

    def verificar_cambios_externos(self) -> bool:
        """Indica si otra terminal o proceso escribió en la base SQLite."""
        if not self.db_path.exists():
            return False
        try:
            mtime_actual = self.db_path.stat().st_mtime
            with self.read_connection() as conn:
                row = conn.execute("PRAGMA data_version;").fetchone()
                dv_actual = int(row[0]) if row and row[0] is not None else 0

            hubo_cambio = (
                (self._last_mtime > 0 and mtime_actual != self._last_mtime)
                or (self._last_data_version > 0 and dv_actual != self._last_data_version)
            )
            self._last_mtime = mtime_actual
            self._last_data_version = dv_actual
            return hubo_cambio
        except Exception:
            return False

    def integrity_check(self) -> str:
        with self.read_connection() as conn:
            row = conn.execute("PRAGMA integrity_check;").fetchone()
            return str(row[0]) if row else "Error"
