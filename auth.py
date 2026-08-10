"""Funciones de autenticación preparadas para la futura pantalla de login."""

from datetime import datetime

import bcrypt


class AuthService:
    def __init__(self, store):
        self.store = store

    def crear_usuario(self, username, password, nombre="", roles=()):
        if not password:
            raise ValueError("La contraseña no puede estar vacía.")
        password_hash = bcrypt.hashpw(password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
        with self.store.write_transaction() as connection:
            cursor = connection.execute(
                """INSERT INTO usuarios(username, nombre, password_hash, creado_en)
                   VALUES (?, ?, ?, ?)""",
                (username.strip(), nombre.strip(), password_hash, self.store.now()),
            )
            user_id = cursor.lastrowid
            for role in roles:
                connection.execute("INSERT OR IGNORE INTO roles(nombre) VALUES (?)", (role,))
                role_id = connection.execute(
                    "SELECT id FROM roles WHERE nombre=?", (role,)
                ).fetchone()[0]
                connection.execute(
                    "INSERT OR IGNORE INTO usuario_roles(usuario_id, rol_id) VALUES (?, ?)",
                    (user_id, role_id),
                )
            return user_id

    def autenticar(self, username, password):
        with self.store.read_connection() as connection:
            user = connection.execute(
                "SELECT * FROM usuarios WHERE username=? AND activo=1", (username.strip(),)
            ).fetchone()
        if not user or not bcrypt.checkpw(password.encode("utf-8"), user["password_hash"].encode("utf-8")):
            return None
        with self.store.write_transaction() as connection:
            connection.execute(
                "UPDATE usuarios SET ultimo_ingreso=? WHERE id=?",
                (datetime.now().isoformat(timespec="seconds"), user["id"]),
            )
        return dict(user)

    def tiene_permiso(self, user_id, permission_code):
        with self.store.read_connection() as connection:
            return connection.execute(
                """SELECT 1 FROM usuario_roles ur
                   JOIN rol_permisos rp ON rp.rol_id=ur.rol_id
                   JOIN permisos p ON p.id=rp.permiso_id
                   WHERE ur.usuario_id=? AND p.codigo=? LIMIT 1""",
                (user_id, permission_code),
            ).fetchone() is not None

    def listar_usuarios(self):
        with self.store.read_connection() as connection:
            return connection.execute(
                """SELECT u.id, u.username, u.nombre, u.activo,
                          COALESCE(GROUP_CONCAT(r.nombre, ', '), '') AS roles
                   FROM usuarios u LEFT JOIN usuario_roles ur ON ur.usuario_id=u.id
                   LEFT JOIN roles r ON r.id=ur.rol_id
                   GROUP BY u.id ORDER BY u.username"""
            ).fetchall()

    def listar_roles(self):
        with self.store.read_connection() as connection:
            return connection.execute(
                """SELECT r.id, r.nombre,
                          COALESCE(GROUP_CONCAT(p.codigo, ', '), '') AS permisos
                   FROM roles r LEFT JOIN rol_permisos rp ON rp.rol_id=r.id
                   LEFT JOIN permisos p ON p.id=rp.permiso_id
                   GROUP BY r.id ORDER BY r.nombre"""
            ).fetchall()

    def listar_permisos(self):
        with self.store.read_connection() as connection:
            return [row[0] for row in connection.execute(
                "SELECT codigo FROM permisos ORDER BY codigo"
            ).fetchall()]

    def crear_rol(self, nombre):
        with self.store.write_transaction() as connection:
            return connection.execute("INSERT INTO roles(nombre) VALUES (?)", (nombre.strip(),)).lastrowid

    def establecer_permisos_rol(self, role_id, permissions):
        with self.store.write_transaction() as connection:
            permission_ids = []
            for permission in permissions:
                connection.execute("INSERT OR IGNORE INTO permisos(codigo) VALUES (?)", (permission,))
                permission_ids.append(connection.execute(
                    "SELECT id FROM permisos WHERE codigo=?", (permission,)
                ).fetchone()[0])
            connection.execute("DELETE FROM rol_permisos WHERE rol_id=?", (role_id,))
            connection.executemany(
                "INSERT INTO rol_permisos(rol_id, permiso_id) VALUES (?, ?)",
                [(role_id, permission_id) for permission_id in permission_ids],
            )

    def permisos_de_rol(self, role_id):
        with self.store.read_connection() as connection:
            return [row[0] for row in connection.execute(
                """SELECT p.codigo FROM permisos p JOIN rol_permisos rp ON rp.permiso_id=p.id
                   WHERE rp.rol_id=? ORDER BY p.codigo""", (role_id,)
            ).fetchall()]

    def establecer_roles_usuario(self, user_id, roles):
        with self.store.write_transaction() as connection:
            connection.execute("DELETE FROM usuario_roles WHERE usuario_id=?", (user_id,))
            for role in roles:
                connection.execute("INSERT OR IGNORE INTO roles(nombre) VALUES (?)", (role,))
                role_id = connection.execute("SELECT id FROM roles WHERE nombre=?", (role,)).fetchone()[0]
                connection.execute(
                    "INSERT INTO usuario_roles(usuario_id, rol_id) VALUES (?, ?)",
                    (user_id, role_id),
                )

    def roles_de_usuario(self, user_id):
        with self.store.read_connection() as connection:
            return [row[0] for row in connection.execute(
                """SELECT r.nombre FROM roles r JOIN usuario_roles ur ON ur.rol_id=r.id
                   WHERE ur.usuario_id=? ORDER BY r.nombre""", (user_id,)
            ).fetchall()]

    def cambiar_estado_usuario(self, user_id):
        with self.store.write_transaction() as connection:
            connection.execute("UPDATE usuarios SET activo=1-activo WHERE id=?", (user_id,))

    def cambiar_password(self, user_id, password):
        if not password:
            raise ValueError("La contraseña no puede estar vacía.")
        password_hash = bcrypt.hashpw(password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
        with self.store.write_transaction() as connection:
            connection.execute("UPDATE usuarios SET password_hash=? WHERE id=?", (password_hash, user_id))

    def inicializar_permisos(self, permissions):
        with self.store.write_transaction() as connection:
            connection.executemany("INSERT OR IGNORE INTO permisos(codigo) VALUES (?)", [(p,) for p in permissions])
            connection.execute(
                """INSERT OR IGNORE INTO rol_permisos(rol_id, permiso_id)
                   SELECT r.id, p.id
                   FROM roles r CROSS JOIN permisos p
                   WHERE r.nombre = 'Administrador' AND p.codigo IN ({})""".format(
                       ",".join("?" for _ in permissions)
                   ),
                tuple(permissions),
            )

    def crear_administrador_inicial(self, username, password):
        """Crea el primer usuario y le asigna todos los permisos disponibles."""
        with self.store.read_connection() as connection:
            if connection.execute("SELECT 1 FROM usuarios LIMIT 1").fetchone():
                raise ValueError("Ya existe un usuario administrador.")
        permissions = (
            "novedades.ver", "novedades.crear", "novedades.editar", "novedades.eliminar",
            "cambios_turno.ver", "cambios_turno.crear", "cambios_turno.editar", "excel.exportar",
            "usuarios.administrar", "roles.administrar", "empleados.importar", "auditoria.ver",
            "dotaciones.administrar",
        )
        with self.store.write_transaction() as connection:
            password_hash = bcrypt.hashpw(password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
            user_id = connection.execute(
                "INSERT INTO usuarios(username, nombre, password_hash, creado_en) VALUES (?, ?, ?, ?)",
                (username.strip(), username.strip(), password_hash, self.store.now()),
            ).lastrowid
            role_id = connection.execute("INSERT INTO roles(nombre) VALUES ('Administrador')").lastrowid
            for permission in permissions:
                connection.execute("INSERT OR IGNORE INTO permisos(codigo) VALUES (?)", (permission,))
                permission_id = connection.execute(
                    "SELECT id FROM permisos WHERE codigo=?", (permission,)
                ).fetchone()[0]
                connection.execute(
                    "INSERT INTO rol_permisos(rol_id, permiso_id) VALUES (?, ?)",
                    (role_id, permission_id),
                )
            connection.execute(
                "INSERT INTO usuario_roles(usuario_id, rol_id) VALUES (?, ?)", (user_id, role_id)
            )
            return user_id
