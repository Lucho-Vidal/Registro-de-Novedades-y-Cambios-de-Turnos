"""Servicio de autenticación, gestión de credenciales y RBAC para RENO v2.0.0.

Soporta bcrypt, protección contra fuerza bruta, caducidad de contraseñas,
recuperación mediante correo electrónico institucional y control de permisos por rol.
"""

import json
import secrets
import time
from datetime import datetime
from typing import Any, Callable, Dict, List, Optional, Tuple

import bcrypt

from src.core.logger import get_logger
from src.data.database import DatabaseManager
from src.domain.enums import PERMISOS_BASE

logger = get_logger("AuthService")

MAX_INTENTOS_FALLIDOS = 5
BLOQUEO_SEGUNDOS = 30


class CuentaBloqueadaError(Exception):
    """Se levanta cuando una cuenta está bloqueada temporalmente por intentos fallidos."""

    def __init__(self, segundos_restantes: int):
        super().__init__(
            f"Cuenta temporalmente bloqueada por demasiados intentos fallidos. "
            f"Intente nuevamente en {segundos_restantes} segundos."
        )
        self.segundos_restantes = segundos_restantes


class AuthService:
    """Gestor de autenticación y permisos."""

    def __init__(self, db: DatabaseManager):
        self.db = db
        self._intentos_fallidos: Dict[str, Dict[str, Any]] = {}

    def crear_usuario(
        self,
        username: str,
        password: str,
        nombre: str = "",
        legajo: Optional[int] = None,
        roles: Tuple[str, ...] = (),
    ) -> int:
        if not password:
            raise ValueError("La contraseña no puede estar vacía.")
        if legajo is None:
            raise ValueError("El legajo es obligatorio para nuevos usuarios.")
        password_hash = bcrypt.hashpw(password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
        with self.db.write_transaction() as conn:
            cur = conn.execute(
                """INSERT INTO usuarios(username, nombre, legajo, password_hash, creado_en, password_cambiado_en)
                   VALUES (?, ?, ?, ?, ?, ?)""",
                (username.strip(), nombre.strip(), legajo, password_hash, self.db.now(), self.db.now()),
            )
            user_id = cur.lastrowid
            for role in roles:
                conn.execute("INSERT OR IGNORE INTO roles(nombre) VALUES (?)", (role,))
                role_id = conn.execute("SELECT id FROM roles WHERE nombre=?", (role,)).fetchone()[0]
                conn.execute(
                    "INSERT OR IGNORE INTO usuario_roles(usuario_id, rol_id) VALUES (?, ?)",
                    (user_id, role_id),
                )
            return user_id

    def autenticar(self, username: str, password: str) -> Optional[Dict[str, Any]]:
        clave = (username or "").strip().casefold()
        registro = self._intentos_fallidos.get(clave)
        ahora = time.time()
        if registro and registro.get("bloqueado_hasta"):
            restante = registro["bloqueado_hasta"] - ahora
            if restante > 0:
                raise CuentaBloqueadaError(int(restante) + 1)
            del self._intentos_fallidos[clave]

        with self.db.read_connection() as conn:
            user = conn.execute(
                "SELECT * FROM usuarios WHERE username=? COLLATE NOCASE AND activo=1",
                (username.strip(),),
            ).fetchone()

        if not user or not bcrypt.checkpw(password.encode("utf-8"), user["password_hash"].encode("utf-8")):
            intentos = self._intentos_fallidos.get(clave, {"count": 0})
            intentos["count"] += 1
            if intentos["count"] >= MAX_INTENTOS_FALLIDOS:
                intentos["count"] = 0
                intentos["bloqueado_hasta"] = ahora + BLOQUEO_SEGUNDOS
            self._intentos_fallidos[clave] = intentos
            return None

        self._intentos_fallidos.pop(clave, None)
        with self.db.write_transaction() as conn:
            conn.execute(
                "UPDATE usuarios SET ultimo_ingreso=? WHERE id=?",
                (datetime.now().isoformat(timespec="seconds"), user["id"]),
            )
        return dict(user)

    def tiene_permiso(self, user_id: int, permission_code: str) -> bool:
        with self.db.read_connection() as conn:
            row = conn.execute(
                """SELECT 1 FROM usuario_roles ur
                   JOIN rol_permisos rp ON rp.rol_id=ur.rol_id
                   JOIN permisos p ON p.id=rp.permiso_id
                   WHERE ur.usuario_id=? AND p.codigo=?
                   LIMIT 1""",
                (user_id, permission_code),
            ).fetchone()
            return row is not None

    def listar_usuarios(self) -> List[Dict[str, Any]]:
        with self.db.read_connection() as conn:
            rows = conn.execute(
                """SELECT u.id, u.username, u.nombre, u.legajo, u.activo,
                          COALESCE(GROUP_CONCAT(r.nombre, ', '), '') AS roles
                   FROM usuarios u LEFT JOIN usuario_roles ur ON ur.usuario_id=u.id
                   LEFT JOIN roles r ON r.id=ur.rol_id
                   GROUP BY u.id ORDER BY u.username"""
            ).fetchall()
            return [dict(r) for r in rows]

    def actualizar_usuario(
        self,
        user_id: int,
        username: str,
        nombre: str,
        legajo: Optional[int],
        activo: int,
        roles: Optional[List[str]] = None,
    ) -> None:
        now_ts = self.db.now()
        with self.db.write_transaction() as conn:
            conn.execute(
                """UPDATE usuarios
                   SET username=?, nombre=?, legajo=?, activo=?,
                       sync_status='pending', updated_at=?
                   WHERE id=?""",
                (username.strip(), nombre.strip(), legajo, int(activo), now_ts, user_id),
            )
            if roles is not None:
                conn.execute("DELETE FROM usuario_roles WHERE usuario_id=?", (user_id,))
                for role in roles:
                    r_clean = str(role).strip()
                    if not r_clean:
                        continue
                    conn.execute("INSERT OR IGNORE INTO roles(nombre) VALUES (?)", (r_clean,))
                    role_id = conn.execute("SELECT id FROM roles WHERE nombre=?", (r_clean,)).fetchone()[0]
                    conn.execute(
                        "INSERT INTO usuario_roles(usuario_id, rol_id) VALUES (?, ?)",
                        (user_id, role_id),
                    )

            user_row = conn.execute("SELECT * FROM usuarios WHERE id=?", (int(user_id),)).fetchone()
            if user_row:
                payload = json.dumps(dict(user_row), ensure_ascii=False)
                conn.execute(
                    """INSERT INTO sync_outbox(tabla, registro_pk, operacion, payload_json, creado_en, intentos)
                       VALUES ('usuarios', ?, 'UPSERT', ?, ?, 0)
                       ON CONFLICT(tabla, registro_pk) DO UPDATE SET
                           operacion='UPSERT', payload_json=excluded.payload_json,
                           creado_en=excluded.creado_en, intentos=0, ultimo_error=NULL""",
                    (str(user_row["username"]), payload, now_ts),
                )

    def obtener_usuario(self, user_id: int) -> Optional[Dict[str, Any]]:
        with self.db.read_connection() as conn:
            row = conn.execute("SELECT * FROM usuarios WHERE id=?", (user_id,)).fetchone()
            return dict(row) if row else None

    def cambiar_mi_password(self, user_id: int, actual: str, nueva: str, username: Optional[str] = None) -> None:
        if not nueva:
            raise ValueError("La contraseña nueva no puede estar vacía.")
        with self.db.read_connection() as conn:
            row = conn.execute("SELECT id, password_hash FROM usuarios WHERE id=?", (user_id,)).fetchone()
            if not row and username:
                row = conn.execute(
                    "SELECT id, password_hash FROM usuarios WHERE username=? COLLATE NOCASE",
                    (username.strip(),),
                ).fetchone()
        if not row or not bcrypt.checkpw(actual.encode("utf-8"), str(row["password_hash"]).encode("utf-8")):
            raise ValueError("La contraseña actual no es correcta.")
        self.cambiar_password(int(row["id"]), nueva)

    def listar_roles(self) -> List[Dict[str, Any]]:
        with self.db.read_connection() as conn:
            rows = conn.execute(
                """SELECT r.id, r.nombre,
                          COALESCE(GROUP_CONCAT(p.codigo, ', '), '') AS permisos
                   FROM roles r LEFT JOIN rol_permisos rp ON rp.rol_id=r.id
                   LEFT JOIN permisos p ON p.id=rp.permiso_id
                   GROUP BY r.id ORDER BY r.nombre"""
            ).fetchall()
            return [dict(r) for r in rows]

    def listar_permisos(self) -> List[str]:
        with self.db.read_connection() as conn:
            return [row[0] for row in conn.execute("SELECT codigo FROM permisos ORDER BY codigo").fetchall()]

    def crear_rol(self, nombre: str) -> int:
        with self.db.write_transaction() as conn:
            return conn.execute("INSERT INTO roles(nombre) VALUES (?)", (nombre.strip(),)).lastrowid

    def establecer_permisos_rol(self, role_id: int, permissions: List[str]) -> None:
        with self.db.write_transaction() as conn:
            permission_ids = []
            for permission in permissions:
                conn.execute("INSERT OR IGNORE INTO permisos(codigo) VALUES (?)", (permission,))
                permission_ids.append(conn.execute(
                    "SELECT id FROM permisos WHERE codigo=?", (permission,)
                ).fetchone()[0])
            conn.execute("DELETE FROM rol_permisos WHERE rol_id=?", (role_id,))
            conn.executemany(
                "INSERT INTO rol_permisos(rol_id, permiso_id) VALUES (?, ?)",
                [(role_id, pid) for pid in permission_ids],
            )

    def permisos_de_rol(self, role_id: int) -> List[str]:
        with self.db.read_connection() as conn:
            return [row[0] for row in conn.execute(
                """SELECT p.codigo FROM permisos p JOIN rol_permisos rp ON rp.permiso_id=p.id
                   WHERE rp.rol_id=? ORDER BY p.codigo""", (role_id,)
            ).fetchall()]

    def establecer_roles_usuario(self, user_id: int, roles: List[str]) -> None:
        with self.db.write_transaction() as conn:
            conn.execute("DELETE FROM usuario_roles WHERE usuario_id=?", (user_id,))
            for role in roles:
                conn.execute("INSERT OR IGNORE INTO roles(nombre) VALUES (?)", (role,))
                role_id = conn.execute("SELECT id FROM roles WHERE nombre=?", (role,)).fetchone()[0]
                conn.execute(
                    "INSERT INTO usuario_roles(usuario_id, rol_id) VALUES (?, ?)",
                    (user_id, role_id),
                )

    def roles_de_usuario(self, user_id: int) -> List[str]:
        with self.db.read_connection() as conn:
            return [row[0] for row in conn.execute(
                """SELECT r.nombre FROM roles r JOIN usuario_roles ur ON ur.rol_id=r.id
                   WHERE ur.usuario_id=? ORDER BY r.nombre""", (user_id,)
            ).fetchall()]

    def cambiar_estado_usuario(self, user_id: int) -> None:
        with self.db.write_transaction() as conn:
            conn.execute("UPDATE usuarios SET activo=1-activo WHERE id=?", (user_id,))

    def cambiar_password(self, user_id: int, password: str) -> None:
        if not password:
            raise ValueError("La contraseña no puede estar vacía.")
        password_hash = bcrypt.hashpw(password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
        now_ts = self.db.now()
        with self.db.write_transaction() as conn:
            conn.execute(
                """UPDATE usuarios
                   SET password_hash=?, password_cambiado_en=?, debe_cambiar_clave=0,
                       sync_status='pending', updated_at=?
                   WHERE id=?""",
                (password_hash, now_ts, now_ts, int(user_id)),
            )
            user_row = conn.execute("SELECT * FROM usuarios WHERE id=?", (int(user_id),)).fetchone()
            if user_row:
                payload = json.dumps(dict(user_row), ensure_ascii=False)
                conn.execute(
                    """INSERT INTO sync_outbox(tabla, registro_pk, operacion, payload_json, creado_en, intentos)
                       VALUES ('usuarios', ?, 'UPSERT', ?, ?, 0)
                       ON CONFLICT(tabla, registro_pk) DO UPDATE SET
                           operacion='UPSERT', payload_json=excluded.payload_json,
                           creado_en=excluded.creado_en, intentos=0, ultimo_error=NULL""",
                    (str(user_row["username"]), payload, now_ts),
                )

    def resetear_password(
        self,
        username: str,
        send_fn: Optional[Callable[[str, str], None]] = None,
    ) -> Tuple[str, str]:
        nombre = (username or "").strip()
        with self.db.read_connection() as conn:
            user = conn.execute(
                "SELECT * FROM usuarios WHERE username=? COLLATE NOCASE AND activo=1", (nombre,)
            ).fetchone()
        if not user:
            raise ValueError("Usuario no encontrado.")

        with self.db.read_connection() as conn:
            cfg = conn.execute("SELECT valor FROM configuracion WHERE clave='correo_dominio'").fetchone()
            dominio = str(cfg[0] if cfg else "@trenesargentinos.gob.ar")

        email = (nombre + dominio).strip()
        if not email or "@" not in email:
            raise RuntimeError("No se pudo armar el correo electrónico para ese usuario.")
        temporal = secrets.token_urlsafe(9)

        if send_fn is not None:
            send_fn(email, temporal)

        password_hash = bcrypt.hashpw(temporal.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
        now = self.db.now()
        with self.db.write_transaction() as conn:
            conn.execute(
                "UPDATE usuarios SET password_hash=?, password_cambiado_en=?, debe_cambiar_clave=1 WHERE id=?",
                (password_hash, now, user["id"]),
            )
            conn.execute(
                """INSERT INTO auditoria
                   (usuario_id, usuario_windows, accion, entidad, entidad_id,
                    datos_anteriores, datos_nuevos, creado_en)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                (
                    user["id"], None, "resetear_clave", "usuarios", user["id"],
                    None,
                    json.dumps({"metodo": "email", "correo": email}, ensure_ascii=False),
                    now,
                ),
            )
        return email, temporal

    def requiere_cambio_clave(self, user_id: int) -> bool:
        with self.db.read_connection() as conn:
            row = conn.execute(
                "SELECT password_cambiado_en, debe_cambiar_clave FROM usuarios WHERE id=?", (user_id,)
            ).fetchone()
            cfg = conn.execute("SELECT valor FROM configuracion WHERE clave='clave_expiracion_dias'").fetchone()
        if not row:
            return False
        if int(row["debe_cambiar_clave"] or 0) == 1:
            return True
        try:
            dias = max(0, int(cfg[0])) if cfg else 90
        except (TypeError, ValueError):
            dias = 0
        if dias <= 0:
            return False
        if not row["password_cambiado_en"]:
            return True
        try:
            cambiado = datetime.fromisoformat(str(row["password_cambiado_en"]))
        except ValueError:
            return True
        return (datetime.now() - cambiado).days >= dias

    def inicializar_permisos(self, permissions: Tuple[str, ...]) -> None:
        with self.db.write_transaction() as conn:
            conn.executemany("INSERT OR IGNORE INTO permisos(codigo) VALUES (?)", [(p,) for p in permissions])
            conn.execute(
                """INSERT OR IGNORE INTO rol_permisos(rol_id, permiso_id)
                   SELECT r.id, p.id
                   FROM roles r CROSS JOIN permisos p
                   WHERE r.nombre = 'Administrador' AND p.codigo IN ({})""".format(
                    ",".join("?" for _ in permissions)
                ),
                tuple(permissions),
            )

    def crear_administrador_inicial(self, username: str, password: str, legajo: Optional[int] = None) -> int:
        with self.db.read_connection() as conn:
            if conn.execute("SELECT 1 FROM usuarios LIMIT 1").fetchone():
                raise ValueError("Ya existe un usuario administrador.")
        permissions = PERMISOS_BASE
        with self.db.write_transaction() as conn:
            password_hash = bcrypt.hashpw(password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
            cur = conn.execute(
                """INSERT INTO usuarios(username, nombre, legajo, password_hash, creado_en, password_cambiado_en)
                   VALUES (?, ?, ?, ?, ?, ?)""",
                (username.strip(), username.strip(), legajo, password_hash, self.db.now(), self.db.now()),
            )
            user_id = cur.lastrowid
            role_id = conn.execute("INSERT INTO roles(nombre) VALUES ('Administrador')").lastrowid
            for permission in permissions:
                conn.execute("INSERT OR IGNORE INTO permisos(codigo) VALUES (?)", (permission,))
                permission_id = conn.execute("SELECT id FROM permisos WHERE codigo=?", (permission,)).fetchone()[0]
                conn.execute(
                    "INSERT INTO rol_permisos(rol_id, permiso_id) VALUES (?, ?)",
                    (role_id, permission_id),
                )
            conn.execute("INSERT INTO usuario_roles(usuario_id, rol_id) VALUES (?, ?)", (user_id, role_id))
            return user_id
