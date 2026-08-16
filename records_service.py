"""Operaciones de negocio para registros operativos, tipos y auditoría."""

import json
import sqlite3
from datetime import datetime, timedelta


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

    def _parsear_registrado_en(self, valor):
        if valor is None or str(valor).strip() == "":
            return None
        texto = str(valor).strip()
        try:
            return datetime.fromisoformat(texto)
        except ValueError:
            pass
        for formato in (
            "%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d/%m/%Y",
            "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d",
        ):
            try:
                return datetime.strptime(texto, formato)
            except ValueError:
                continue
        try:
            return datetime(1899, 12, 30) + timedelta(days=float(texto))
        except ValueError:
            return None

    def _horas_limite(self, connection, clave, default):
        row = connection.execute("SELECT valor FROM configuracion WHERE clave=?", (clave,)).fetchone()
        try:
            return max(1, int(row[0])) if row else default
        except (TypeError, ValueError):
            return default

    def _dentro_de_ventana(self, connection, registrado_en, clave, default):
        fecha = self._parsear_registrado_en(registrado_en)
        if fecha is None:
            return False
        horas = self._horas_limite(connection, clave, default)
        return (datetime.now() - fecha).total_seconds() <= horas * 3600

    def dentro_de_ventana(self, registrado_en, clave):
        default = 24 if clave == "editar_horas" else 72
        with self.store.read_connection() as connection:
            return self._dentro_de_ventana(connection, registrado_en, clave, default)

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

    def listar_empleados(self, incluir_inactivos=True):
        with self.store.read_connection() as connection:
            query = "SELECT id, legajo, apellidos_nombres, especialidad, dotacion, turnos, franco, activo FROM empleados"
            if not incluir_inactivos:
                query += " WHERE activo=1"
            return connection.execute(query + " ORDER BY apellidos_nombres").fetchall()

    def obtener_empleado(self, empleado_id):
        with self.store.read_connection() as connection:
            return connection.execute("SELECT * FROM empleados WHERE id=?", (empleado_id,)).fetchone()

    def _validar_empleado(self, data):
        try:
            legajo = int(str(data.get("legajo")).strip())
        except (TypeError, ValueError):
            raise ValueError("El legajo debe ser un número.")
        apellidos = (data.get("apellidos_nombres") or "").strip()
        if not apellidos:
            raise ValueError("Los apellidos y nombres son obligatorios.")
        return legajo, apellidos

    def crear_empleado(self, data, user_id=None, windows_user=None):
        legajo, apellidos = self._validar_empleado(data)
        valores = {
            "legajo": legajo,
            "apellidos_nombres": apellidos,
            "especialidad": (data.get("especialidad") or "").strip(),
            "dotacion": (data.get("dotacion") or "").strip(),
            "turnos": (data.get("turnos") or "").strip(),
            "franco": (data.get("franco") or "").strip(),
        }
        with self.store.write_transaction() as connection:
            try:
                cursor = connection.execute(
                    """INSERT INTO empleados(legajo, apellidos_nombres, especialidad, dotacion, turnos, franco, actualizado_en)
                       VALUES (?, ?, ?, ?, ?, ?, ?)""",
                    (valores["legajo"], valores["apellidos_nombres"], valores["especialidad"],
                     valores["dotacion"], valores["turnos"], valores["franco"], self.store.now()),
                )
            except sqlite3.IntegrityError:
                raise ValueError(f"Ya existe un empleado con el legajo {legajo}.")
            self._audit(connection, user_id, windows_user, "crear", "empleado", cursor.lastrowid, None, valores)
            return cursor.lastrowid

    def actualizar_empleado(self, empleado_id, data, user_id=None, windows_user=None):
        legajo, apellidos = self._validar_empleado(data)
        valores = {
            "legajo": legajo,
            "apellidos_nombres": apellidos,
            "especialidad": (data.get("especialidad") or "").strip(),
            "dotacion": (data.get("dotacion") or "").strip(),
            "turnos": (data.get("turnos") or "").strip(),
            "franco": (data.get("franco") or "").strip(),
        }
        with self.store.write_transaction() as connection:
            before = connection.execute("SELECT * FROM empleados WHERE id=?", (empleado_id,)).fetchone()
            if not before:
                raise ValueError("El empleado no existe.")
            try:
                connection.execute(
                    """UPDATE empleados SET legajo=?, apellidos_nombres=?, especialidad=?, dotacion=?,
                          turnos=?, franco=?, actualizado_en=? WHERE id=?""",
                    (valores["legajo"], valores["apellidos_nombres"], valores["especialidad"],
                     valores["dotacion"], valores["turnos"], valores["franco"], self.store.now(), empleado_id),
                )
            except sqlite3.IntegrityError:
                raise ValueError(f"Ya existe un empleado con el legajo {legajo}.")
            after = connection.execute("SELECT * FROM empleados WHERE id=?", (empleado_id,)).fetchone()
            self._audit(connection, user_id, windows_user, "modificar", "empleado", empleado_id, dict(before), dict(after))

    def cambiar_estado_empleado(self, empleado_id, user_id=None, windows_user=None):
        with self.store.write_transaction() as connection:
            before = connection.execute("SELECT * FROM empleados WHERE id=?", (empleado_id,)).fetchone()
            if not before:
                raise ValueError("El empleado no existe.")
            nuevo_estado = 1 - before["activo"]
            connection.execute("UPDATE empleados SET activo=? WHERE id=?", (nuevo_estado, empleado_id))
            accion = "desactivar" if nuevo_estado == 0 else "reactivar"
            self._audit(connection, user_id, windows_user, accion, "empleado", empleado_id, dict(before), {"activo": nuevo_estado})

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
            if not self._dentro_de_ventana(connection, before["registrado_en"], "editar_horas", 24):
                raise ValueError("El registro superó el tiempo permitido para editar.")
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
            if not self._dentro_de_ventana(connection, before["registrado_en"], "editar_horas", 24):
                raise ValueError("El registro superó el tiempo permitido para editar.")
            assignments = ", ".join(f"{field}=?" for field in fields)
            connection.execute(f"UPDATE cambios_turno SET {assignments} WHERE id=?", tuple(data[field] for field in fields) + (record_id,))
            after = connection.execute("SELECT * FROM cambios_turno WHERE id=?", (record_id,)).fetchone()
            self._audit(connection, user_id, windows_user, "modificar", "cambio_turno", record_id, dict(before), dict(after))

    def eliminar_novedad(self, record_id, user_id=None, windows_user=None):
        with self.store.write_transaction() as connection:
            before = connection.execute("SELECT * FROM novedades WHERE id=?", (record_id,)).fetchone()
            if not before:
                raise ValueError("La novedad no existe.")
            if not self._dentro_de_ventana(connection, before["registrado_en"], "eliminar_horas", 72):
                raise ValueError("El registro superó el tiempo permitido para eliminar.")
            connection.execute("UPDATE novedades SET activo=0 WHERE id=?", (record_id,))
            self._audit(connection, user_id, windows_user, "eliminar", "novedad", record_id, dict(before), {"activo": 0})

    def eliminar_cambio(self, record_id, user_id=None, windows_user=None):
        with self.store.write_transaction() as connection:
            before = connection.execute("SELECT * FROM cambios_turno WHERE id=?", (record_id,)).fetchone()
            if not before:
                raise ValueError("El cambio de turno no existe.")
            if not self._dentro_de_ventana(connection, before["registrado_en"], "eliminar_horas", 72):
                raise ValueError("El registro superó el tiempo permitido para eliminar.")
            connection.execute("UPDATE cambios_turno SET activo=0 WHERE id=?", (record_id,))
            self._audit(connection, user_id, windows_user, "eliminar", "cambio_turno", record_id, dict(before), {"activo": 0})

    def _tabla_entidad(self, tipo):
        if tipo == "Novedad":
            return "novedades", "novedad"
        if tipo == "Cambio":
            return "cambios_turno", "cambio_turno"
        raise ValueError("Tipo de registro inválido.")

    def recuperar_registro(self, tipo, record_id, user_id=None, windows_user=None):
        table, entidad = self._tabla_entidad(tipo)
        with self.store.write_transaction() as connection:
            before = connection.execute(f"SELECT * FROM {table} WHERE id=?", (record_id,)).fetchone()
            if not before:
                raise ValueError("El registro no existe.")
            connection.execute(f"UPDATE {table} SET activo=1 WHERE id=?", (record_id,))
            self._audit(connection, user_id, windows_user, "recuperar", entidad, record_id, dict(before), {"activo": 1})

    def borrar_definitivo(self, tipo, record_id, user_id=None, windows_user=None):
        table, entidad = self._tabla_entidad(tipo)
        with self.store.write_transaction() as connection:
            before = connection.execute(f"SELECT * FROM {table} WHERE id=?", (record_id,)).fetchone()
            if not before:
                raise ValueError("El registro no existe.")
            connection.execute(f"DELETE FROM {table} WHERE id=?", (record_id,))
            self._audit(connection, user_id, windows_user, "borrar_definitivo", entidad, record_id, dict(before), None)

    def listar_eliminados(self):
        with self.store.read_connection() as connection:
            return connection.execute(
                """SELECT 'Novedad' AS tipo, id, registrado_en, legajo, apellidos_nombres,
                          dotacion, novedad AS detalle, observaciones
                   FROM novedades WHERE activo=0
                   UNION ALL
                   SELECT 'Cambio' AS tipo, id, registrado_en, legajo_1, apellidos_nombres_1,
                          dotacion_1, fecha_cambio, observaciones
                   FROM cambios_turno WHERE activo=0
                   ORDER BY registrado_en DESC"""
            ).fetchall()

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
