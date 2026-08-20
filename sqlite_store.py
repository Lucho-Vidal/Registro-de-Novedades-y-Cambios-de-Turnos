"""Persistencia SQLite compartida para el registro de novedades.

La base está pensada para vivir en una carpeta SMB. No se utiliza WAL: cada
escritura toma un bloqueo externo breve y luego ejecuta una transacción SQLite.
"""

from contextlib import contextmanager
from datetime import datetime
import os
import sqlite3
import time

try:
    import msvcrt
except ImportError:  # pragma: no cover - solo se usa fuera de Windows
    msvcrt = None

try:
    import fcntl
except ImportError:  # pragma: no cover - solo se usa en POSIX
    fcntl = None


SCHEMA_VERSION = 1


class DatabaseBusyError(RuntimeError):
    """La base no pudo reservarse dentro del tiempo configurado."""


def _try_lock(file_handle):
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
    raise RuntimeError("No hay un mecanismo de bloqueo disponible en este sistema.")


def _unlock(file_handle):
    if msvcrt:
        file_handle.seek(0)
        try:
            msvcrt.locking(file_handle.fileno(), msvcrt.LK_UNLCK, 1)
        except OSError:
            pass
    elif fcntl:
        fcntl.flock(file_handle.fileno(), fcntl.LOCK_UN)


class SQLiteStore:
    def __init__(self, database_path, lock_timeout=15):
        self.database_path = os.path.abspath(database_path)
        self.lock_path = f"{self.database_path}.lock"
        self.lock_timeout = lock_timeout
        os.makedirs(os.path.dirname(self.database_path) or ".", exist_ok=True)

    def connect(self):
        connection = sqlite3.connect(
            self.database_path,
            timeout=self.lock_timeout,
            isolation_level=None,
        )
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA journal_mode=DELETE")
        connection.execute("PRAGMA synchronous=FULL")
        connection.execute(f"PRAGMA busy_timeout={self.lock_timeout * 1000}")
        connection.execute("PRAGMA foreign_keys=ON")
        connection.execute("PRAGMA temp_store=MEMORY")
        return connection

    @contextmanager
    def write_transaction(self):
        started = time.monotonic()
        lock_file = None
        while True:
            try:
                lock_file = open(self.lock_path, "a+b")
                if os.path.getsize(self.lock_path) == 0:
                    lock_file.write(b"0")
                    lock_file.flush()
                if _try_lock(lock_file):
                    break
                raise OSError("lock ocupado")
            except OSError:
                if lock_file:
                    lock_file.close()
                    lock_file = None
                if time.monotonic() - started >= self.lock_timeout:
                    raise DatabaseBusyError(
                        "La base está siendo utilizada por otra PC. Intente nuevamente."
                    )
                time.sleep(0.15)

        connection = self.connect()
        try:
            connection.execute("BEGIN IMMEDIATE")
            yield connection
            if connection.in_transaction:
                connection.execute("COMMIT")
        except Exception:
            if connection.in_transaction:
                connection.rollback()
            raise
        finally:
            connection.close()
            _unlock(lock_file)
            lock_file.close()

    @contextmanager
    def read_connection(self):
        connection = self.connect()
        try:
            yield connection
        finally:
            connection.close()

    def initialize(self):
        with self.write_transaction() as connection:
            connection.executescript(
                """
                CREATE TABLE IF NOT EXISTS schema_metadata (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS empleados (
                    id INTEGER PRIMARY KEY,
                    legajo INTEGER NOT NULL UNIQUE,
                    apellidos_nombres TEXT NOT NULL DEFAULT '',
                    especialidad TEXT NOT NULL DEFAULT '',
                    dotacion TEXT NOT NULL DEFAULT '',
                    turnos TEXT NOT NULL DEFAULT '',
                    franco TEXT NOT NULL DEFAULT '',
                    activo INTEGER NOT NULL DEFAULT 1 CHECK (activo IN (0, 1)),
                    actualizado_en TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS tipos_novedad (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    nombre TEXT NOT NULL COLLATE NOCASE UNIQUE,
                    activo INTEGER NOT NULL DEFAULT 1 CHECK (activo IN (0, 1))
                );

                CREATE TABLE IF NOT EXISTS dotaciones (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    nombre TEXT NOT NULL COLLATE NOCASE UNIQUE,
                    activo INTEGER NOT NULL DEFAULT 1 CHECK (activo IN (0, 1))
                );

                CREATE TABLE IF NOT EXISTS usuarios (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    username TEXT NOT NULL COLLATE NOCASE UNIQUE,
                    nombre TEXT NOT NULL DEFAULT '',
                    legajo INTEGER,
                    password_hash TEXT NOT NULL,
                    activo INTEGER NOT NULL DEFAULT 1 CHECK (activo IN (0, 1)),
                    creado_en TEXT NOT NULL,
                    ultimo_ingreso TEXT
                );

                CREATE TABLE IF NOT EXISTS roles (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    nombre TEXT NOT NULL COLLATE NOCASE UNIQUE
                );

                CREATE TABLE IF NOT EXISTS permisos (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    codigo TEXT NOT NULL COLLATE NOCASE UNIQUE
                );

                CREATE TABLE IF NOT EXISTS usuario_roles (
                    usuario_id INTEGER NOT NULL REFERENCES usuarios(id),
                    rol_id INTEGER NOT NULL REFERENCES roles(id),
                    PRIMARY KEY (usuario_id, rol_id)
                );

                CREATE TABLE IF NOT EXISTS rol_permisos (
                    rol_id INTEGER NOT NULL REFERENCES roles(id),
                    permiso_id INTEGER NOT NULL REFERENCES permisos(id),
                    PRIMARY KEY (rol_id, permiso_id)
                );

                CREATE TABLE IF NOT EXISTS novedades (
                    id INTEGER PRIMARY KEY,
                    registrado_en TEXT NOT NULL,
                    legajo INTEGER NOT NULL,
                    apellidos_nombres TEXT NOT NULL DEFAULT '',
                    especialidad TEXT NOT NULL DEFAULT '',
                    dotacion TEXT NOT NULL DEFAULT '',
                    turnos TEXT NOT NULL DEFAULT '',
                    franco TEXT NOT NULL DEFAULT '',
                    novedad TEXT NOT NULL,
                    fecha_inicio TEXT NOT NULL,
                    fecha_fin TEXT,
                    referencia_estacion TEXT NOT NULL,
                    supervisor TEXT NOT NULL,
                    observaciones TEXT,
                    usuario_windows TEXT,
                    usuario_id INTEGER REFERENCES usuarios(id),
                    activo INTEGER NOT NULL DEFAULT 1 CHECK (activo IN (0, 1))
                );

                CREATE TABLE IF NOT EXISTS cambios_turno (
                    id INTEGER PRIMARY KEY,
                    registrado_en TEXT NOT NULL,
                    legajo_1 INTEGER NOT NULL,
                    apellidos_nombres_1 TEXT NOT NULL DEFAULT '',
                    especialidad_1 TEXT NOT NULL DEFAULT '',
                    dotacion_1 TEXT NOT NULL DEFAULT '',
                    turnos_1 TEXT NOT NULL DEFAULT '',
                    franco_1 TEXT NOT NULL DEFAULT '',
                    legajo_2 INTEGER NOT NULL,
                    apellidos_nombres_2 TEXT NOT NULL DEFAULT '',
                    especialidad_2 TEXT NOT NULL DEFAULT '',
                    dotacion_2 TEXT NOT NULL DEFAULT '',
                    turnos_2 TEXT NOT NULL DEFAULT '',
                    franco_2 TEXT NOT NULL DEFAULT '',
                    fecha_cambio TEXT NOT NULL,
                    referencia_estacion TEXT NOT NULL,
                    supervisor TEXT NOT NULL,
                    observaciones TEXT,
                    usuario_windows TEXT,
                    usuario_id INTEGER REFERENCES usuarios(id),
                    activo INTEGER NOT NULL DEFAULT 1 CHECK (activo IN (0, 1))
                );

                CREATE TABLE IF NOT EXISTS auditoria (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    usuario_id INTEGER REFERENCES usuarios(id),
                    usuario_windows TEXT,
                    accion TEXT NOT NULL,
                    entidad TEXT NOT NULL,
                    entidad_id INTEGER,
                    datos_anteriores TEXT,
                    datos_nuevos TEXT,
                    creado_en TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS personal_estacion (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    nombre TEXT NOT NULL COLLATE NOCASE UNIQUE,
                    activo INTEGER NOT NULL DEFAULT 1 CHECK (activo IN (0, 1))
                );

                CREATE TABLE IF NOT EXISTS destinatarios_informe (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    nombre TEXT NOT NULL DEFAULT '',
                    email TEXT NOT NULL COLLATE NOCASE UNIQUE,
                    activo INTEGER NOT NULL DEFAULT 1 CHECK (activo IN (0, 1))
                );

                CREATE TABLE IF NOT EXISTS configuracion (
                    clave TEXT PRIMARY KEY,
                    valor TEXT NOT NULL
                );

                CREATE INDEX IF NOT EXISTS idx_novedades_legajo ON novedades(legajo);
                CREATE INDEX IF NOT EXISTS idx_novedades_registrado ON novedades(registrado_en DESC);
                CREATE INDEX IF NOT EXISTS idx_cambios_legajo_1 ON cambios_turno(legajo_1);
                CREATE INDEX IF NOT EXISTS idx_cambios_legajo_2 ON cambios_turno(legajo_2);
                CREATE INDEX IF NOT EXISTS idx_cambios_registrado ON cambios_turno(registrado_en DESC);
                """
            )
            connection.execute(
                "INSERT OR REPLACE INTO schema_metadata(key, value) VALUES (?, ?)",
                ("version", str(SCHEMA_VERSION)),
            )
            user_columns = {row[1] for row in connection.execute("PRAGMA table_info(usuarios)").fetchall()}
            if "legajo" not in user_columns:
                connection.execute("ALTER TABLE usuarios ADD COLUMN legajo INTEGER")
            if "password_cambiado_en" not in user_columns:
                connection.execute("ALTER TABLE usuarios ADD COLUMN password_cambiado_en TEXT")
            if "debe_cambiar_clave" not in user_columns:
                connection.execute("ALTER TABLE usuarios ADD COLUMN debe_cambiar_clave INTEGER NOT NULL DEFAULT 0")
            for table in ("novedades", "cambios_turno"):
                columns = {row[1] for row in connection.execute(f"PRAGMA table_info({table})").fetchall()}
                if "activo" not in columns:
                    connection.execute(f"ALTER TABLE {table} ADD COLUMN activo INTEGER NOT NULL DEFAULT 1")
            connection.execute(
                "INSERT OR IGNORE INTO configuracion(clave, valor) VALUES ('sesion_minutos', '30')"
            )
            connection.execute(
                "INSERT OR IGNORE INTO configuracion(clave, valor) VALUES ('editar_horas', '24')"
            )
            connection.execute(
                "INSERT OR IGNORE INTO configuracion(clave, valor) VALUES ('eliminar_horas', '72')"
            )
            connection.execute(
                "INSERT OR IGNORE INTO configuracion(clave, valor) VALUES ('backup_activo', '1')"
            )
            connection.execute(
                "INSERT OR IGNORE INTO configuracion(clave, valor) VALUES ('backup_retencion', '10')"
            )
            connection.execute(
                "INSERT OR IGNORE INTO configuracion(clave, valor) VALUES ('notificaciones_activo', '1')"
            )
            connection.execute(
                "INSERT OR IGNORE INTO configuracion(clave, valor) VALUES ('toast_duracion', '6')"
            )
            connection.execute(
                "INSERT OR IGNORE INTO configuracion(clave, valor) VALUES ('fila_alto', '30')"
            )
            connection.execute(
                "INSERT OR IGNORE INTO configuracion(clave, valor) VALUES ('clave_expiracion_dias', '90')"
            )
            connection.execute(
                "INSERT OR IGNORE INTO configuracion(clave, valor) VALUES ('correo_dominio', '@trenesargentinos.gob.ar')"
            )
            connection.execute("INSERT OR IGNORE INTO tipos_novedad(nombre) VALUES ('Informe')")
            for nombre in ("PC", "LLV", "TY", "LP", "OA", "K5", "RE", "CÑ", "AK"):
                connection.execute("INSERT OR IGNORE INTO dotaciones(nombre) VALUES (?)", (nombre,))
            connection.execute(
                """INSERT OR IGNORE INTO dotaciones(nombre)
                   SELECT DISTINCT TRIM(dotacion) FROM empleados
                   WHERE TRIM(dotacion) <> ''"""
            )

    def integrity_check(self):
        with self.read_connection() as connection:
            return connection.execute("PRAGMA integrity_check").fetchone()[0]

    def now(self):
        return datetime.now().isoformat(timespec="seconds")

    def next_id(self, table):
        if table not in {"novedades", "cambios_turno"}:
            raise ValueError("Tabla no permitida")
        with self.read_connection() as connection:
            return connection.execute(f"SELECT COALESCE(MAX(id), 0) + 1 FROM {table}").fetchone()[0]

    def listar_resumen_activos(self, table):
        if table not in {"novedades", "cambios_turno"}:
            raise ValueError("Tabla no permitida")
        name_col = "apellidos_nombres" if table == "novedades" else "apellidos_nombres_1"
        legajo_col = "legajo" if table == "novedades" else "legajo_1"
        with self.read_connection() as connection:
            return connection.execute(
                f"""SELECT id, registrado_en, usuario_windows, usuario_id,
                           {name_col} AS apellidos_nombres, {legajo_col} AS legajo
                    FROM {table} WHERE activo=1"""
            ).fetchall()

    def insert_novedad(self, values):
        with self.write_transaction() as connection:
            values = list(values)
            if values[0] is None:
                values[0] = connection.execute("SELECT COALESCE(MAX(id), 0) + 1 FROM novedades").fetchone()[0]
            cursor = connection.execute(
                """INSERT INTO novedades
                   (id, registrado_en, legajo, apellidos_nombres, especialidad, dotacion,
                    turnos, franco, novedad, fecha_inicio, fecha_fin, referencia_estacion,
                    supervisor, observaciones, usuario_windows, usuario_id)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                values,
            )
            return cursor.lastrowid

    def insert_cambio_turno(self, values):
        with self.write_transaction() as connection:
            values = list(values)
            if values[0] is None:
                values[0] = connection.execute("SELECT COALESCE(MAX(id), 0) + 1 FROM cambios_turno").fetchone()[0]
            cursor = connection.execute(
                """INSERT INTO cambios_turno
                   (id, registrado_en, legajo_1, apellidos_nombres_1, especialidad_1,
                    dotacion_1, turnos_1, franco_1, legajo_2, apellidos_nombres_2,
                    especialidad_2, dotacion_2, turnos_2, franco_2, fecha_cambio,
                    referencia_estacion, supervisor, observaciones, usuario_windows, usuario_id)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                values,
            )
            return cursor.lastrowid

    def get_base_rows(self):
        with self.read_connection() as connection:
            return connection.execute(
                """SELECT legajo, apellidos_nombres, especialidad, dotacion, turnos, franco
                   FROM empleados WHERE activo=1 ORDER BY apellidos_nombres"""
            ).fetchall()

    def get_tipo_novedades(self):
        with self.read_connection() as connection:
            return [row[0] for row in connection.execute(
                "SELECT nombre FROM tipos_novedad WHERE activo=1 ORDER BY nombre"
            ).fetchall()]

    def get_dotaciones(self):
        with self.read_connection() as connection:
            return [row[0] for row in connection.execute(
                "SELECT nombre FROM dotaciones WHERE activo=1 ORDER BY nombre"
            ).fetchall()]

    def get_personal_estacion(self, incluir_inactivos=False):
        with self.read_connection() as connection:
            query = "SELECT id, nombre, activo FROM personal_estacion"
            if not incluir_inactivos:
                query += " WHERE activo=1"
            return connection.execute(query + " ORDER BY nombre").fetchall()

    def get_destinatarios_informe(self, incluir_inactivos=False):
        with self.read_connection() as connection:
            query = "SELECT id, nombre, email, activo FROM destinatarios_informe"
            if not incluir_inactivos:
                query += " WHERE activo=1"
            return connection.execute(query + " ORDER BY nombre, email").fetchall()

    def get_configuracion(self, clave, default=None):
        with self.read_connection() as connection:
            row = connection.execute("SELECT valor FROM configuracion WHERE clave=?", (clave,)).fetchone()
            return row[0] if row else default

    def set_configuracion(self, clave, valor):
        with self.write_transaction() as connection:
            connection.execute(
                "INSERT INTO configuracion(clave, valor) VALUES (?, ?) "
                "ON CONFLICT(clave) DO UPDATE SET valor=excluded.valor",
                (clave, str(valor)),
            )

    def sincronizar_dotaciones(self):
        """Registra automáticamente dotaciones nuevas provenientes de empleados importados."""
        with self.write_transaction() as connection:
            connection.execute(
                """INSERT OR IGNORE INTO dotaciones(nombre)
                   SELECT DISTINCT TRIM(dotacion) FROM empleados
                   WHERE TRIM(dotacion) <> ''"""
            )
            return [row[0] for row in connection.execute(
                "SELECT nombre FROM dotaciones WHERE activo=1 ORDER BY nombre"
            ).fetchall()]


class SQLiteSheetAdapter:
    """Adaptador mínimo para que las tablas existentes lean desde SQLite."""

    def __init__(self, store, query):
        self.store = store
        self.query = query

    def iter_rows(self, min_row=2, values_only=True):
        with self.store.read_connection() as connection:
            rows = connection.execute(self.query).fetchall()
        for row in rows:
            yield tuple(row)
