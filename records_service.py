"""Operaciones de negocio para registros operativos, tipos y auditoría."""

import json
import sqlite3
import unicodedata
from datetime import datetime, timedelta

_DIAS_SEMANA = ("Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo")
_DIAS_SEMANA_NORM = {
    "lunes", "martes", "miercoles", "jueves", "viernes", "sabado", "domingo",
}


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

    def parsear_registrado_en(self, valor):
        if valor is None:
            return None
        texto = str(valor).strip()
        if not texto:
            return None
        # Vía rápida formato ISO: YYYY-MM-DD...
        if len(texto) >= 10 and texto[4] == "-" and texto[7] == "-":
            try:
                return datetime.fromisoformat(texto)
            except ValueError:
                pass
        # Vía rápida formato DD/MM/YYYY...
        elif len(texto) >= 10 and texto[2] == "/" and texto[5] == "/":
            if len(texto) == 10:
                try:
                    return datetime.strptime(texto, "%d/%m/%Y")
                except ValueError:
                    pass
            elif len(texto) == 16:
                try:
                    return datetime.strptime(texto, "%d/%m/%Y %H:%M")
                except ValueError:
                    pass
            elif len(texto) >= 19:
                try:
                    return datetime.strptime(texto[:19], "%d/%m/%Y %H:%M:%S")
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

    def _parsear_registrado_en(self, valor):
        return self.parsear_registrado_en(valor)

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
            query = "SELECT id, nombre, dias, exportar_descubierto, activo FROM tipos_novedad"
            if not incluir_inactivos:
                query += " WHERE activo=1"
            query += " ORDER BY nombre"
            return connection.execute(query).fetchall()

    def registrar_auditoria(self, action, entity, entity_id, user_id=None, windows_user=None, before=None, after=None):
        with self.store.write_transaction() as connection:
            self._audit(connection, user_id, windows_user, action, entity, entity_id, before, after)

    def crear_tipo(self, nombre, dias=0, exportar_descubierto=1, user_id=None, windows_user=None):
        nombre = (nombre or "").strip()
        if not nombre:
            raise ValueError("El nombre no puede estar vacío.")
        try:
            dias_val = max(0, int(dias or 0))
        except (ValueError, TypeError):
            dias_val = 0
        exportar_val = 1 if exportar_descubierto else 0
        with self.store.write_transaction() as connection:
            cursor = connection.execute(
                "INSERT INTO tipos_novedad(nombre, dias, exportar_descubierto) VALUES (?, ?, ?)",
                (nombre, dias_val, exportar_val),
            )
            self._audit(
                connection, user_id, windows_user, "crear", "tipo_novedad", cursor.lastrowid, None,
                {"nombre": nombre, "dias": dias_val, "exportar_descubierto": exportar_val},
            )
            return cursor.lastrowid

    def actualizar_tipo(self, type_id, nombre, dias=0, exportar_descubierto=1, activo=1, user_id=None, windows_user=None):
        nombre = (nombre or "").strip()
        if not nombre:
            raise ValueError("El nombre no puede estar vacío.")
        try:
            dias_val = max(0, int(dias or 0))
        except (ValueError, TypeError):
            dias_val = 0
        exportar_val = 1 if exportar_descubierto else 0
        with self.store.write_transaction() as connection:
            before = connection.execute("SELECT * FROM tipos_novedad WHERE id=?", (type_id,)).fetchone()
            if not before:
                raise ValueError("El tipo de novedad no existe.")
            connection.execute(
                "UPDATE tipos_novedad SET nombre=?, dias=?, exportar_descubierto=?, activo=? WHERE id=?",
                (nombre, dias_val, exportar_val, int(activo), type_id),
            )
            self._audit(
                connection, user_id, windows_user, "modificar", "tipo_novedad", type_id, dict(before),
                {"nombre": nombre, "dias": dias_val, "exportar_descubierto": exportar_val, "activo": int(activo)},
            )

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

    def crear_personal_estacion(self, nombre, dotacion=None, user_id=None, windows_user=None):
        nombre = nombre.strip()
        if not nombre:
            raise ValueError("El nombre no puede estar vacío.")
        dotacion = (dotacion or "").strip() or None
        with self.store.write_transaction() as connection:
            cursor = connection.execute(
                "INSERT INTO personal_estacion(nombre, dotacion) VALUES (?, ?)", (nombre, dotacion)
            )
            self._audit(connection, user_id, windows_user, "crear", "personal_estacion", cursor.lastrowid, None, {"nombre": nombre, "dotacion": dotacion})
            return cursor.lastrowid

    def actualizar_personal_estacion(self, item_id, nombre, activo, user_id=None, windows_user=None, *, dotacion=None, cambiar_dotacion=False):
        nombre = nombre.strip()
        if cambiar_dotacion:
            dotacion = (dotacion or "").strip() or None
        with self.store.write_transaction() as connection:
            before = connection.execute("SELECT id, nombre, dotacion, activo FROM personal_estacion WHERE id=?", (item_id,)).fetchone()
            if not before:
                raise ValueError("El personal de estación no existe.")
            despues_dotacion = dotacion if cambiar_dotacion else before[2]
            connection.execute(
                "UPDATE personal_estacion SET nombre=?, dotacion=?, activo=? WHERE id=?",
                (nombre, despues_dotacion, int(activo), item_id),
            )
            self._audit(connection, user_id, windows_user, "modificar", "personal_estacion", item_id, dict(before), {"nombre": nombre, "dotacion": despues_dotacion, "activo": int(activo)})

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

    def actualizar_novedad(self, record_id, data, user_id=None, windows_user=None, bypass_window=False):
        fields = ("legajo", "apellidos_nombres", "especialidad", "dotacion", "turnos", "franco", "novedad", "fecha_inicio", "fecha_fin", "referencia_estacion", "supervisor", "observaciones")
        with self.store.write_transaction() as connection:
            before = connection.execute("SELECT * FROM novedades WHERE id=?", (record_id,)).fetchone()
            if not before:
                raise ValueError("La novedad no existe.")
            if not bypass_window and not self._dentro_de_ventana(connection, before["registrado_en"], "editar_horas", 24):
                raise ValueError("El registro superó el tiempo permitido para editar.")
            assignments = ", ".join(f"{field}=?" for field in fields)
            connection.execute(f"UPDATE novedades SET {assignments} WHERE id=?", tuple(data[field] for field in fields) + (record_id,))
            after = connection.execute("SELECT * FROM novedades WHERE id=?", (record_id,)).fetchone()
            self._audit(connection, user_id, windows_user, "modificar", "novedad", record_id, dict(before), dict(after))

    def actualizar_fecha_fin_novedad(self, record_id, fecha_fin, nueva_observacion=None, user_id=None, windows_user=None):
        import validators
        fecha_fin_str = (fecha_fin or "").strip()
        if not fecha_fin_str:
            raise ValueError("Debe ingresar la fecha de fin.")
        val_fin = validators.parsear_fecha(fecha_fin_str, "Fecha de Fin", obligatorio=True)
        with self.store.write_transaction() as connection:
            before = connection.execute("SELECT * FROM novedades WHERE id=?", (record_id,)).fetchone()
            if not before:
                raise ValueError("La novedad no existe.")
            fecha_fin_anterior = (before["fecha_fin"] or "").strip()
            if fecha_fin_anterior and fecha_fin_anterior != "-":
                raise ValueError("El registro ya cuenta con una fecha de fin.")
            fecha_inicio_str = (before["fecha_inicio"] or "").strip()
            if fecha_inicio_str:
                val_ini = validators.parsear_fecha(fecha_inicio_str, "Fecha de Inicio", obligatorio=False)
                if val_ini and val_fin < val_ini:
                    raise ValueError("La fecha de fin no puede ser anterior a la fecha de inicio.")
            obs_existente = (before["observaciones"] or "").strip()
            nueva_obs = (nueva_observacion or "").strip()
            if nueva_obs:
                obs_final = f"{obs_existente}\n{nueva_obs}".strip() if obs_existente else nueva_obs
            else:
                obs_final = before["observaciones"]
            connection.execute(
                "UPDATE novedades SET fecha_fin=?, observaciones=? WHERE id=?",
                (fecha_fin_str, obs_final, record_id),
            )
            after = connection.execute("SELECT * FROM novedades WHERE id=?", (record_id,)).fetchone()
            self._audit(connection, user_id, windows_user, "modificar", "novedad", record_id, dict(before), dict(after))

    def actualizar_cambio(self, record_id, data, user_id=None, windows_user=None, bypass_window=False):
        fields = ("legajo_1", "apellidos_nombres_1", "especialidad_1", "dotacion_1", "turnos_1", "franco_1", "legajo_2", "apellidos_nombres_2", "especialidad_2", "dotacion_2", "turnos_2", "franco_2", "fecha_cambio", "referencia_estacion", "supervisor", "observaciones")
        with self.store.write_transaction() as connection:
            before = connection.execute("SELECT * FROM cambios_turno WHERE id=?", (record_id,)).fetchone()
            if not before:
                raise ValueError("El cambio de turno no existe.")
            if not bypass_window and not self._dentro_de_ventana(connection, before["registrado_en"], "editar_horas", 24):
                raise ValueError("El registro superó el tiempo permitido para editar.")
            assignments = ", ".join(f"{field}=?" for field in fields)
            connection.execute(f"UPDATE cambios_turno SET {assignments} WHERE id=?", tuple(data[field] for field in fields) + (record_id,))
            after = connection.execute("SELECT * FROM cambios_turno WHERE id=?", (record_id,)).fetchone()
            self._audit(connection, user_id, windows_user, "modificar", "cambio_turno", record_id, dict(before), dict(after))

    def eliminar_novedad(self, record_id, user_id=None, windows_user=None, bypass_window=False):
        with self.store.write_transaction() as connection:
            before = connection.execute("SELECT * FROM novedades WHERE id=?", (record_id,)).fetchone()
            if not before:
                raise ValueError("La novedad no existe.")
            if not bypass_window and not self._dentro_de_ventana(connection, before["registrado_en"], "eliminar_horas", 72):
                raise ValueError("El registro superó el tiempo permitido para eliminar.")
            connection.execute("UPDATE novedades SET activo=0 WHERE id=?", (record_id,))
            self._audit(connection, user_id, windows_user, "eliminar", "novedad", record_id, dict(before), {"activo": 0})

    def eliminar_cambio(self, record_id, user_id=None, windows_user=None, bypass_window=False):
        with self.store.write_transaction() as connection:
            before = connection.execute("SELECT * FROM cambios_turno WHERE id=?", (record_id,)).fetchone()
            if not before:
                raise ValueError("El cambio de turno no existe.")
            if not bypass_window and not self._dentro_de_ventana(connection, before["registrado_en"], "eliminar_horas", 72):
                raise ValueError("El registro superó el tiempo permitido para eliminar.")
            connection.execute("UPDATE cambios_turno SET activo=0 WHERE id=?", (record_id,))
            self._audit(connection, user_id, windows_user, "eliminar", "cambio_turno", record_id, dict(before), {"activo": 0})

    def obtener_turnos_sin_cubrir(self, fecha_consulta, dotacion=None):
        """Devuelve las novedades activas que corresponden a turnos sin cubrir en la fecha dada."""
        import validators
        f_sol = validators.parsear_fecha(fecha_consulta, "Fecha a consultar", obligatorio=True)
        with self.store.read_connection() as connection:
            query = """
                SELECT n.id, n.registrado_en, n.legajo, n.apellidos_nombres, n.especialidad,
                       n.dotacion, n.turnos, n.franco, n.novedad, n.fecha_inicio, n.fecha_fin,
                       n.referencia_estacion, n.supervisor, n.observaciones
                FROM novedades n
                JOIN tipos_novedad t ON LOWER(n.novedad) = LOWER(t.nombre)
                WHERE n.activo = 1 AND t.exportar_descubierto = 1
            """
            params = []
            if dotacion and dotacion != "Todas":
                query += " AND n.dotacion = ?"
                params.append(dotacion)
            query += " ORDER BY n.dotacion, n.turnos, n.apellidos_nombres"
            filas = connection.execute(query, params).fetchall()

        descubiertos = []
        for fila in filas:
            try:
                f_ini = validators.parsear_fecha(fila["fecha_inicio"], "Fecha de Inicio", obligatorio=False)
            except Exception:
                f_ini = None
            if not f_ini:
                continue

            fecha_fin_raw = (fila["fecha_fin"] or "").strip()
            f_fin = None
            if fecha_fin_raw and fecha_fin_raw != "-":
                try:
                    f_fin = validators.parsear_fecha(fecha_fin_raw, "Fecha de Fin", obligatorio=False)
                except Exception:
                    f_fin = None

            if f_ini.date() <= f_sol.date():
                if f_fin is None or f_sol.date() <= f_fin.date():
                    descubiertos.append(fila)

        return descubiertos

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

    def _desde_hace(self, dias):
        return datetime.now() - timedelta(days=max(1, int(dias)))

    def _es_dia_semana(self, valor):
        """Indica si el texto corresponde a un día de la semana (ignora mayúsculas/acentos)."""
        texto = unicodedata.normalize("NFKD", str(valor or ""))
        texto = "".join(c for c in texto if not unicodedata.combining(c)).lower().strip()
        return texto in _DIAS_SEMANA_NORM

    def obtener_metricas_dashboard(self, dias=30):
        """Calcula todas las métricas del panel en una sola lectura optimizada."""
        dias = max(1, int(dias))
        hoy = datetime.now()
        desde = hoy - timedelta(days=dias - 1)
        fechas = [(desde + timedelta(days=i)).strftime("%Y-%m-%d") for i in range(dias)]
        serie = {fecha: {"novedades": 0, "cambios": 0} for fecha in fechas}

        conteo_tipo = {}
        conteo_dotacion = {}
        conteo_dia_semana = {dia: 0 for dia in _DIAS_SEMANA}

        resumen = {
            "novedades_hoy": 0, "cambios_hoy": 0,
            "novedades_mes": 0, "cambios_mes": 0, "total_mes": 0,
        }

        def sumar_dotacion(dotacion):
            valor = (dotacion or "").strip()
            if not valor or self._es_dia_semana(valor):
                return
            conteo_dotacion[valor] = conteo_dotacion.get(valor, 0) + 1

        with self.store.read_connection() as connection:
            filas_nov = connection.execute(
                "SELECT registrado_en, novedad, dotacion FROM novedades WHERE activo=1"
            ).fetchall()
            filas_cam = connection.execute(
                "SELECT registrado_en, dotacion_1, dotacion_2 FROM cambios_turno WHERE activo=1"
            ).fetchall()

        for fila in filas_nov:
            fecha = self._parsear_registrado_en(fila["registrado_en"])
            if fecha is None:
                continue
            if fecha.date() == hoy.date():
                resumen["novedades_hoy"] += 1
            if fecha.year == hoy.year and fecha.month == hoy.month:
                resumen["novedades_mes"] += 1

            if fecha >= desde:
                conteo_dia_semana[_DIAS_SEMANA[fecha.weekday()]] += 1
                tipo = (fila["novedad"] or "").strip() or "Sin tipo"
                conteo_tipo[tipo] = conteo_tipo.get(tipo, 0) + 1
                sumar_dotacion(fila["dotacion"])
                clave = fecha.strftime("%Y-%m-%d")
                if clave in serie:
                    serie[clave]["novedades"] += 1

        for fila in filas_cam:
            fecha = self._parsear_registrado_en(fila["registrado_en"])
            if fecha is None:
                continue
            if fecha.date() == hoy.date():
                resumen["cambios_hoy"] += 1
            if fecha.year == hoy.year and fecha.month == hoy.month:
                resumen["cambios_mes"] += 1

            if fecha >= desde:
                conteo_dia_semana[_DIAS_SEMANA[fecha.weekday()]] += 1
                sumar_dotacion(fila["dotacion_1"])
                sumar_dotacion(fila["dotacion_2"])
                clave = fecha.strftime("%Y-%m-%d")
                if clave in serie:
                    serie[clave]["cambios"] += 1

        resumen["total_mes"] = resumen["novedades_mes"] + resumen["cambios_mes"]

        return {
            "resumen": resumen,
            "por_tipo": sorted(conteo_tipo.items(), key=lambda item: item[1], reverse=True),
            "por_dotacion": sorted(conteo_dotacion.items(), key=lambda item: item[1], reverse=True)[:10],
            "tendencia": [(f, serie[f]["novedades"], serie[f]["cambios"]) for f in fechas],
            "por_dia_semana": list(conteo_dia_semana.items()),
        }

    def dashboard_por_dia_semana(self, dias=30):
        return self.obtener_metricas_dashboard(dias)["por_dia_semana"]

    def dashboard_resumen(self):
        return self.obtener_metricas_dashboard(30)["resumen"]

    def dashboard_por_tipo(self, dias=30):
        return self.obtener_metricas_dashboard(dias)["por_tipo"]

    def dashboard_por_dotacion(self, dias=30, limite=10):
        return self.obtener_metricas_dashboard(dias)["por_dotacion"][:max(1, int(limite))]

    def dashboard_tendencia(self, dias=30):
        return self.obtener_metricas_dashboard(dias)["tendencia"]
