"""Operaciones de negocio para registros operativos, tipos y auditoría."""

import json


class RecordsService:
    def __init__(self, store):
        self.store = store

    def _audit(self, connection, user_id, windows_user, action, entity, entity_id, before, after):
        connection.execute(
            """INSERT INTO auditoria
               (usuario_id, usuario_windows, accion, entidad, entidad_id,
                datos_anteriores, datos_nuevos, creado_en)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                user_id, windows_user, action, entity, entity_id,
                json.dumps(before, ensure_ascii=False, default=str) if before else None,
                json.dumps(after, ensure_ascii=False, default=str) if after else None,
                self.store.now(),
            ),
        )

    def listar_tipos(self, incluir_inactivos=True):
        with self.store.read_connection() as connection:
            query = "SELECT id, nombre, activo FROM tipos_novedad"
            if not incluir_inactivos:
                query += " WHERE activo=1"
            query += " ORDER BY nombre"
            return connection.execute(query).fetchall()

    def registrar_auditoria(self, action, entity, entity_id, user_id=None, windows_user=None, before=None, after=None):
        with self.store.write_transaction() as connection:
            self._audit(connection, user_id, windows_user, action, entity, entity_id, before, after)

    def crear_tipo(self, nombre, user_id=None, windows_user=None):
        with self.store.write_transaction() as connection:
            cursor = connection.execute("INSERT INTO tipos_novedad(nombre) VALUES (?)", (nombre.strip(),))
            self._audit(connection, user_id, windows_user, "crear", "tipo_novedad", cursor.lastrowid, None, {"nombre": nombre})
            return cursor.lastrowid

    def actualizar_tipo(self, type_id, nombre, activo, user_id=None, windows_user=None):
        with self.store.write_transaction() as connection:
            before = connection.execute("SELECT id, nombre, activo FROM tipos_novedad WHERE id=?", (type_id,)).fetchone()
            if not before:
                raise ValueError("El tipo de novedad no existe.")
            connection.execute("UPDATE tipos_novedad SET nombre=?, activo=? WHERE id=?", (nombre.strip(), int(activo), type_id))
            self._audit(connection, user_id, windows_user, "modificar", "tipo_novedad", type_id, dict(before), {"nombre": nombre, "activo": int(activo)})

    def listar_dotaciones(self, incluir_inactivos=True):
        with self.store.read_connection() as connection:
            query = "SELECT id, nombre, activo FROM dotaciones"
            if not incluir_inactivos:
                query += " WHERE activo=1"
            return connection.execute(query + " ORDER BY nombre").fetchall()

    def crear_dotacion(self, nombre, user_id=None, windows_user=None):
        nombre = nombre.strip()
        if not nombre:
            raise ValueError("La dotación no puede estar vacía.")
        with self.store.write_transaction() as connection:
            cursor = connection.execute("INSERT INTO dotaciones(nombre) VALUES (?)", (nombre,))
            self._audit(connection, user_id, windows_user, "crear", "dotacion", cursor.lastrowid, None, {"nombre": nombre})
            return cursor.lastrowid

    def actualizar_dotacion(self, dotacion_id, nombre, activo, user_id=None, windows_user=None):
        nombre = nombre.strip()
        if not nombre:
            raise ValueError("La dotación no puede estar vacía.")
        with self.store.write_transaction() as connection:
            before = connection.execute("SELECT id, nombre, activo FROM dotaciones WHERE id=?", (dotacion_id,)).fetchone()
            if not before:
                raise ValueError("La dotación no existe.")
            connection.execute("UPDATE dotaciones SET nombre=?, activo=? WHERE id=?", (nombre, int(activo), dotacion_id))
            self._audit(connection, user_id, windows_user, "modificar", "dotacion", dotacion_id, dict(before), {"nombre": nombre, "activo": int(activo)})

    def listar_personal_estacion(self, incluir_inactivos=True):
        return self.store.get_personal_estacion(incluir_inactivos)

    def crear_personal_estacion(self, nombre, user_id=None, windows_user=None):
        nombre = nombre.strip()
        if not nombre:
            raise ValueError("El nombre no puede estar vacío.")
        with self.store.write_transaction() as connection:
            cursor = connection.execute("INSERT INTO personal_estacion(nombre) VALUES (?)", (nombre,))
            self._audit(connection, user_id, windows_user, "crear", "personal_estacion", cursor.lastrowid, None, {"nombre": nombre})
            return cursor.lastrowid

    def actualizar_personal_estacion(self, item_id, nombre, activo, user_id=None, windows_user=None):
        nombre = nombre.strip()
        with self.store.write_transaction() as connection:
            before = connection.execute("SELECT id, nombre, activo FROM personal_estacion WHERE id=?", (item_id,)).fetchone()
            if not before:
                raise ValueError("El personal de estación no existe.")
            connection.execute("UPDATE personal_estacion SET nombre=?, activo=? WHERE id=?", (nombre, int(activo), item_id))
            self._audit(connection, user_id, windows_user, "modificar", "personal_estacion", item_id, dict(before), {"nombre": nombre, "activo": int(activo)})

    def listar_destinatarios_informe(self, incluir_inactivos=True):
        return self.store.get_destinatarios_informe(incluir_inactivos)

    def crear_destinatario_informe(self, nombre, email, user_id=None, windows_user=None):
        nombre, email = nombre.strip(), email.strip()
        if not email or "@" not in email:
            raise ValueError("Ingrese un correo válido.")
        with self.store.write_transaction() as connection:
            cursor = connection.execute("INSERT INTO destinatarios_informe(nombre, email) VALUES (?, ?)", (nombre, email))
            self._audit(connection, user_id, windows_user, "crear", "destinatario_informe", cursor.lastrowid, None, {"nombre": nombre, "email": email})
            return cursor.lastrowid

    def actualizar_destinatario_informe(self, item_id, nombre, email, activo, user_id=None, windows_user=None):
        with self.store.write_transaction() as connection:
            before = connection.execute("SELECT id, nombre, email, activo FROM destinatarios_informe WHERE id=?", (item_id,)).fetchone()
            if not before:
                raise ValueError("El destinatario no existe.")
            connection.execute("UPDATE destinatarios_informe SET nombre=?, email=?, activo=? WHERE id=?", (nombre.strip(), email.strip(), int(activo), item_id))
            self._audit(connection, user_id, windows_user, "modificar", "destinatario_informe", item_id, dict(before), {"nombre": nombre, "email": email, "activo": int(activo)})

    def obtener_novedad(self, record_id):
        with self.store.read_connection() as connection:
            return connection.execute("SELECT * FROM novedades WHERE id=?", (record_id,)).fetchone()

    def obtener_cambio(self, record_id):
        with self.store.read_connection() as connection:
            return connection.execute("SELECT * FROM cambios_turno WHERE id=?", (record_id,)).fetchone()

    def actualizar_novedad(self, record_id, data, user_id=None, windows_user=None):
        fields = ("legajo", "apellidos_nombres", "especialidad", "dotacion", "turnos", "franco", "novedad", "fecha_inicio", "fecha_fin", "referencia_estacion", "supervisor", "observaciones")
        with self.store.write_transaction() as connection:
            before = connection.execute("SELECT * FROM novedades WHERE id=?", (record_id,)).fetchone()
            if not before:
                raise ValueError("La novedad no existe.")
            assignments = ", ".join(f"{field}=?" for field in fields)
            connection.execute(f"UPDATE novedades SET {assignments} WHERE id=?", tuple(data[field] for field in fields) + (record_id,))
            after = connection.execute("SELECT * FROM novedades WHERE id=?", (record_id,)).fetchone()
            self._audit(connection, user_id, windows_user, "modificar", "novedad", record_id, dict(before), dict(after))

    def actualizar_cambio(self, record_id, data, user_id=None, windows_user=None):
        fields = ("legajo_1", "apellidos_nombres_1", "especialidad_1", "dotacion_1", "turnos_1", "franco_1", "legajo_2", "apellidos_nombres_2", "especialidad_2", "dotacion_2", "turnos_2", "franco_2", "fecha_cambio", "referencia_estacion", "supervisor", "observaciones")
        with self.store.write_transaction() as connection:
            before = connection.execute("SELECT * FROM cambios_turno WHERE id=?", (record_id,)).fetchone()
            if not before:
                raise ValueError("El cambio de turno no existe.")
            assignments = ", ".join(f"{field}=?" for field in fields)
            connection.execute(f"UPDATE cambios_turno SET {assignments} WHERE id=?", tuple(data[field] for field in fields) + (record_id,))
            after = connection.execute("SELECT * FROM cambios_turno WHERE id=?", (record_id,)).fetchone()
            self._audit(connection, user_id, windows_user, "modificar", "cambio_turno", record_id, dict(before), dict(after))

    def listar_auditoria(self, text_filter=""):
        with self.store.read_connection() as connection:
            like = f"%{text_filter.strip()}%"
            return connection.execute(
                """SELECT a.id, a.creado_en, COALESCE(u.username, a.usuario_windows, '') AS usuario,
                          a.accion, a.entidad, a.entidad_id, a.datos_anteriores, a.datos_nuevos
                   FROM auditoria a LEFT JOIN usuarios u ON u.id=a.usuario_id
                   WHERE ?='' OR COALESCE(u.username, a.usuario_windows, '') LIKE ?
                         OR a.accion LIKE ? OR a.entidad LIKE ?
                   ORDER BY a.id DESC""",
                (text_filter.strip(), like, like, like),
            ).fetchall()
