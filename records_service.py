"""Operaciones de negocio para registros operativos, tipos y auditoría."""

import calendar
import json
import sqlite3
import unicodedata
from datetime import date, datetime, time, timedelta

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
            query = "SELECT id, legajo, apellidos_nombres, especialidad, dotacion, turnos, franco, franco_horario, jornada_horas, activo FROM empleados"
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
            "franco_horario": (data.get("franco_horario") or "").strip(),
            "jornada_horas": 7 if str(data.get("jornada_horas")).strip() == "7" else 6,
        }
        with self.store.write_transaction() as connection:
            try:
                cursor = connection.execute(
                    """INSERT INTO empleados(legajo, apellidos_nombres, especialidad, dotacion, turnos, franco, franco_horario, jornada_horas, actualizado_en)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                    (valores["legajo"], valores["apellidos_nombres"], valores["especialidad"],
                     valores["dotacion"], valores["turnos"], valores["franco"], valores["franco_horario"], valores["jornada_horas"], self.store.now()),
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
            "franco_horario": (data.get("franco_horario") or "").strip(),
            "jornada_horas": 7 if str(data.get("jornada_horas")).strip() == "7" else 6,
        }
        with self.store.write_transaction() as connection:
            before = connection.execute("SELECT * FROM empleados WHERE id=?", (empleado_id,)).fetchone()
            if not before:
                raise ValueError("El empleado no existe.")
            try:
                connection.execute(
                    """UPDATE empleados SET legajo=?, apellidos_nombres=?, especialidad=?, dotacion=?,
                          turnos=?, franco=?, franco_horario=?, jornada_horas=?, actualizado_en=? WHERE id=?""",
                    (valores["legajo"], valores["apellidos_nombres"], valores["especialidad"],
                     valores["dotacion"], valores["turnos"], valores["franco"], valores["franco_horario"], valores["jornada_horas"], self.store.now(), empleado_id),
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
        nueva_obs = (nueva_observacion or "").strip()
        if not fecha_fin_str and not nueva_obs:
            raise ValueError("Debe ingresar la fecha de fin o una observación.")
        val_fin = None
        if fecha_fin_str:
            val_fin = validators.parsear_fecha(fecha_fin_str, "Fecha de Fin", obligatorio=True)
        with self.store.write_transaction() as connection:
            before = connection.execute("SELECT * FROM novedades WHERE id=?", (record_id,)).fetchone()
            if not before:
                raise ValueError("La novedad no existe.")
            fecha_fin_anterior = (before["fecha_fin"] or "").strip()
            if fecha_fin_str:
                if fecha_fin_anterior and fecha_fin_anterior != "-":
                    raise ValueError("El registro ya cuenta con una fecha de fin.")
                fecha_inicio_str = (before["fecha_inicio"] or "").strip()
                if fecha_inicio_str:
                    val_ini = validators.parsear_fecha(fecha_inicio_str, "Fecha de Inicio", obligatorio=False)
                    if val_ini and val_fin < val_ini:
                        raise ValueError("La fecha de fin no puede ser anterior a la fecha de inicio.")
                fecha_guardar = fecha_fin_str
            else:
                fecha_guardar = before["fecha_fin"]
            obs_existente = (before["observaciones"] or "").strip()
            if nueva_obs:
                obs_final = f"{obs_existente}\n{nueva_obs}".strip() if obs_existente else nueva_obs
            else:
                obs_final = before["observaciones"]
            connection.execute(
                "UPDATE novedades SET fecha_fin=?, observaciones=? WHERE id=?",
                (fecha_guardar, obs_final, record_id),
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

    def obtener_candidatos_ciclo(self, dotacion, especialidad, excluir_legajo=None):
        """Devuelve los empleados activos con igual dotación, igual especialidad y turno que comienza con CICLO."""
        dot_norm = (dotacion or "").strip()
        esp_norm = (especialidad or "").strip()
        if not dot_norm or not esp_norm:
            return []
        with self.store.read_connection() as connection:
            query = """
                SELECT legajo, apellidos_nombres, especialidad, dotacion, turnos, franco, franco_horario
                FROM empleados
                WHERE activo = 1
                  AND LOWER(TRIM(dotacion)) = LOWER(TRIM(?))
                  AND LOWER(TRIM(especialidad)) = LOWER(TRIM(?))
                  AND UPPER(TRIM(turnos)) LIKE 'CICLO%'
            """
            params = [dot_norm, esp_norm]
            if excluir_legajo is not None:
                try:
                    query += " AND legajo != ?"
                    params.append(int(excluir_legajo))
                except (ValueError, TypeError):
                    pass
            query += " ORDER BY apellidos_nombres"
            return [dict(r) for r in connection.execute(query, params).fetchall()]

    def listar_reemplazos_por_novedad(self, novedad_id):
        """Lista los tramos de reemplazo activos para una novedad dada."""
        with self.store.read_connection() as connection:
            rows = connection.execute(
                """SELECT id, novedad_id, legajo_reemplazo, apellidos_nombres_reemplazo,
                          turnos_reemplazo, franco_reemplazo, fecha_inicio, fecha_fin,
                          usuario_windows, usuario_id, creado_en
                   FROM novedad_reemplazos
                   WHERE novedad_id = ? AND activo = 1
                   ORDER BY id ASC""",
                (novedad_id,),
            ).fetchall()
            return [dict(r) for r in rows]

    def _rango_fechas_reemplazo(self, fecha_ini_str, fecha_fin_str, fecha_fin_nov_str=None):
        import validators
        try:
            dt_ini = validators.parsear_fecha(fecha_ini_str, "Fecha inicio reemplazo", obligatorio=False)
        except Exception:
            dt_ini = None
        if not dt_ini:
            return None, None
        f_ini = dt_ini.date()

        f_fin = None
        raw_fin = (fecha_fin_str or "").strip()
        if raw_fin and raw_fin != "-":
            try:
                dt_fin = validators.parsear_fecha(raw_fin, "Fecha fin reemplazo", obligatorio=False)
                if dt_fin:
                    f_fin = dt_fin.date()
            except Exception:
                f_fin = None
        elif fecha_fin_nov_str and str(fecha_fin_nov_str).strip() not in ("", "-"):
            try:
                dt_fin_nov = validators.parsear_fecha(fecha_fin_nov_str, "Fecha fin novedad", obligatorio=False)
                if dt_fin_nov:
                    f_fin = dt_fin_nov.date()
            except Exception:
                f_fin = None
        return f_ini, f_fin

    def _validar_disponibilidad_ciclo_conn(self, connection, legajo_reemplazo, fecha_inicio, fecha_fin, excluir_novedad_id=None):
        import validators
        dt_ini = validators.parsear_fecha(fecha_inicio, "Fecha inicio reemplazo", obligatorio=True)
        dt_fin = validators.parsear_fecha(fecha_fin, "Fecha fin reemplazo", obligatorio=False)
        t_ini = dt_ini.date()
        t_fin = dt_fin.date() if dt_fin else date.max

        query = """
            SELECT nr.*, n.legajo AS legajo_titular, n.apellidos_nombres AS apellidos_titular,
                   n.novedad AS tipo_novedad, n.fecha_fin AS fecha_fin_novedad
            FROM novedad_reemplazos nr
            JOIN novedades n ON n.id = nr.novedad_id
            WHERE nr.legajo_reemplazo = ? AND nr.activo = 1 AND n.activo = 1
        """
        params = [int(legajo_reemplazo)]
        if excluir_novedad_id is not None:
            query += " AND nr.novedad_id != ?"
            params.append(int(excluir_novedad_id))

        existentes = connection.execute(query, params).fetchall()
        for row in existentes:
            r_ini, r_fin_opt = self._rango_fechas_reemplazo(row["fecha_inicio"], row["fecha_fin"], row["fecha_fin_novedad"])
            if not r_ini:
                continue
            r_fin = r_fin_opt if r_fin_opt is not None else date.max
            if t_ini <= r_fin and r_ini <= t_fin:
                hasta_txt = row["fecha_fin"] or row["fecha_fin_novedad"] or "sin fecha fin"
                raise ValueError(
                    f"El personal de ciclo {row['apellidos_nombres_reemplazo']} (Leg. {legajo_reemplazo}) "
                    f"ya reemplaza a {row['apellidos_titular']} (Novedad #{row['novedad_id']}) "
                    f"del {row['fecha_inicio']} al {hasta_txt}."
                )

    def validar_disponibilidad_ciclo(self, legajo_reemplazo, fecha_inicio, fecha_fin, excluir_novedad_id=None):
        """Verifica que el empleado de CICLO no tenga otro reemplazo activo superpuesto en esas fechas."""
        with self.store.read_connection() as connection:
            self._validar_disponibilidad_ciclo_conn(
                connection, legajo_reemplazo, fecha_inicio, fecha_fin, excluir_novedad_id=excluir_novedad_id
            )

    def guardar_reemplazos_novedad(self, novedad_id, tramos, user_id=None, windows_user=None):
        """Valida y reemplaza los tramos de cobertura de una novedad."""
        import validators
        tramos_norm = []
        with self.store.write_transaction() as connection:
            nov = connection.execute("SELECT * FROM novedades WHERE id = ?", (novedad_id,)).fetchone()
            if not nov:
                raise ValueError("La novedad no existe.")

            fecha_fin_nov = (nov["fecha_fin"] or "").strip()
            if fecha_fin_nov == "-":
                fecha_fin_nov = ""

            for tramo in (tramos or []):
                leg = int(str(tramo.get("legajo_reemplazo") or tramo.get("legajo")).strip())
                f_ini = str(tramo.get("fecha_inicio") or nov["fecha_inicio"] or "").strip()
                f_fin = str(tramo.get("fecha_fin") or "").strip()
                if f_fin == "-":
                    f_fin = ""
                if not f_fin and fecha_fin_nov and len(tramos) == 1:
                    f_fin = fecha_fin_nov

                emp = connection.execute(
                    "SELECT legajo, apellidos_nombres, especialidad, dotacion, turnos, franco FROM empleados WHERE legajo = ?",
                    (leg,),
                ).fetchone()

                nombre = str(tramo.get("apellidos_nombres_reemplazo") or tramo.get("apellidos_nombres") or (emp["apellidos_nombres"] if emp else "")).strip()
                turno_r = str(tramo.get("turnos_reemplazo") or tramo.get("turnos") or (emp["turnos"] if emp else "")).strip()
                franco_r = str(tramo.get("franco_reemplazo") or tramo.get("franco") or (emp["franco"] if emp else "")).strip()

                if emp:
                    if (emp["dotacion"] or "").strip().casefold() != (nov["dotacion"] or "").strip().casefold():
                        raise ValueError(f"El reemplazo {nombre} (Leg. {leg}) no pertenece a la misma dotación ({nov['dotacion']}).")
                    if (emp["especialidad"] or "").strip().casefold() != (nov["especialidad"] or "").strip().casefold():
                        raise ValueError(f"El reemplazo {nombre} (Leg. {leg}) no tiene la misma especialidad ({nov['especialidad']}).")
                    if not validators.es_turno_ciclo(emp["turnos"]):
                        raise ValueError(f"El reemplazo {nombre} (Leg. {leg}) no tiene un turno que comience con CICLO.")
                elif not validators.es_turno_ciclo(turno_r):
                    raise ValueError(f"El reemplazo {nombre} (Leg. {leg}) debe pertenecer a un turno que comience con CICLO.")

                tramos_norm.append({
                    "legajo_reemplazo": leg,
                    "apellidos_nombres_reemplazo": nombre,
                    "turnos_reemplazo": turno_r,
                    "franco_reemplazo": franco_r,
                    "fecha_inicio": f_ini,
                    "fecha_fin": f_fin or None,
                })

            es_valido, err = validators.validar_tramos_reemplazo(
                nov["fecha_inicio"], fecha_fin_nov or None, nov["turnos"], tramos_norm
            )
            if not es_valido:
                raise ValueError(err)

            for t in tramos_norm:
                self._validar_disponibilidad_ciclo_conn(
                    connection, t["legajo_reemplazo"], t["fecha_inicio"], t["fecha_fin"], excluir_novedad_id=novedad_id
                )

            anteriores = [
                dict(r) for r in connection.execute(
                    "SELECT legajo_reemplazo, apellidos_nombres_reemplazo, fecha_inicio, fecha_fin FROM novedad_reemplazos WHERE novedad_id=? AND activo=1",
                    (novedad_id,),
                ).fetchall()
            ]
            connection.execute("UPDATE novedad_reemplazos SET activo = 0 WHERE novedad_id = ? AND activo = 1", (novedad_id,))
            ahora = self.store.now()
            for t in tramos_norm:
                connection.execute(
                    """INSERT INTO novedad_reemplazos
                       (novedad_id, legajo_reemplazo, apellidos_nombres_reemplazo, turnos_reemplazo,
                        franco_reemplazo, fecha_inicio, fecha_fin, usuario_windows, usuario_id, creado_en, activo)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1)""",
                    (
                        novedad_id, t["legajo_reemplazo"], t["apellidos_nombres_reemplazo"],
                        t["turnos_reemplazo"], t["franco_reemplazo"], t["fecha_inicio"],
                        t["fecha_fin"], windows_user, user_id, ahora,
                    ),
                )
            self._audit(
                connection, user_id, windows_user, "reemplazos", "novedad", novedad_id,
                {"reemplazos": anteriores}, {"reemplazos": tramos_norm},
            )

    def calcular_estado_cobertura(self, fila_novedad, tramos_o_fecha_ini=None, fecha_fin=None, turnos=None):
        """Calcula el estado de cobertura ('Sí', 'Parcial', 'No', 'N/A (Ciclo)') y textos descriptivos.

        Acepta dos formas de invocación:
        1) `calcular_estado_cobertura(fila_row_o_dict, lista_tramos)` -> devuelve dict `{"estado", "resumen", "detalle"}`
        2) `calcular_estado_cobertura(novedad_id, fecha_inicio, fecha_fin, turnos)` -> devuelve tupla `(estado, detalle)`
        """
        import validators
        modo_tupla = isinstance(fila_novedad, int) or (turnos is not None or fecha_fin is not None and not isinstance(tramos_o_fecha_ini, list))
        if modo_tupla:
            novedad_id = int(fila_novedad)
            tramos = self.listar_reemplazos_por_novedad(novedad_id)
            fila_dict = {
                "id": novedad_id,
                "fecha_inicio": tramos_o_fecha_ini,
                "fecha_fin": fecha_fin,
                "turnos": turnos or "",
            }
        else:
            fila_dict = dict(fila_novedad)
            tramos = tramos_o_fecha_ini or []

        turno_titular = fila_dict.get("turnos") or ""
        if validators.es_turno_ciclo(turno_titular):
            res = {
                "estado": "N/A (Ciclo)",
                "resumen": "N/A (Ciclo)",
                "detalle": "Personal de CICLO (sin reemplazo)",
            }
            return (res["estado"], res["detalle"]) if modo_tupla else res
        if not tramos:
            res = {
                "estado": "No",
                "resumen": "No",
                "detalle": "Sin reemplazo asignado",
            }
            return (res["estado"], res["detalle"]) if modo_tupla else res

        f_ini_nov_str = fila_dict.get("fecha_inicio") or ""
        f_fin_nov_str = (fila_dict.get("fecha_fin") or "").strip()
        if f_fin_nov_str == "-":
            f_fin_nov_str = ""

        try:
            dt_ini_nov = validators.parsear_fecha(f_ini_nov_str, "Inicio", obligatorio=False)
            f_ini_nov = dt_ini_nov.date() if dt_ini_nov else None
        except Exception:
            f_ini_nov = None

        try:
            dt_fin_nov = validators.parsear_fecha(f_fin_nov_str, "Fin", obligatorio=False) if f_fin_nov_str else None
            f_fin_nov = dt_fin_nov.date() if dt_fin_nov else None
        except Exception:
            f_fin_nov = None

        intervalos = []
        nombres_vistos = []
        partes_detalle = []
        for t in tramos:
            leg = t["legajo_reemplazo"]
            nom = t["apellidos_nombres_reemplazo"]
            etiqueta = f"{nom} ({leg})"
            if etiqueta not in nombres_vistos:
                nombres_vistos.append(etiqueta)
            hasta_txt = t["fecha_fin"] or "abierto"
            partes_detalle.append(f"{leg} - {nom} [{t['fecha_inicio']} a {hasta_txt}]")
            r_ini, r_fin = self._rango_fechas_reemplazo(t["fecha_inicio"], t["fecha_fin"], f_fin_nov_str or None)
            if r_ini:
                intervalos.append((r_ini, r_fin))

        intervalos.sort(key=lambda item: (item[0], item[1] or date.max))
        cubierto_total = False
        if f_ini_nov and intervalos:
            primero_ini, _ = intervalos[0]
            if primero_ini <= f_ini_nov:
                cursor = intervalos[0][1]
                sin_huecos = True
                for sig_ini, sig_fin in intervalos[1:]:
                    if cursor is None:
                        break
                    if sig_ini > cursor + timedelta(days=1):
                        sin_huecos = False
                        break
                    if sig_fin is None or sig_fin > cursor:
                        cursor = sig_fin
                if sin_huecos:
                    if f_fin_nov is not None:
                        if cursor is None or cursor >= f_fin_nov:
                            cubierto_total = True
                    else:
                        if cursor is None:
                            cubierto_total = True

        estado = "Sí" if cubierto_total else "Parcial"
        resumen = f"{estado}: {', '.join(nombres_vistos)}"
        detalle = " | ".join(partes_detalle)
        res = {
            "estado": estado,
            "resumen": resumen,
            "detalle": detalle,
        }
        return (res["estado"], res["detalle"]) if modo_tupla else res

    def obtener_mapa_cobertura_novedades(self):
        """Devuelve un diccionario {novedad_id: info_cobertura} para todas las novedades activas."""
        with self.store.read_connection() as connection:
            novedades = connection.execute(
                "SELECT id, turnos, fecha_inicio, fecha_fin FROM novedades WHERE activo = 1"
            ).fetchall()
            reemplazos = connection.execute(
                """SELECT id, novedad_id, legajo_reemplazo, apellidos_nombres_reemplazo,
                          turnos_reemplazo, franco_reemplazo, fecha_inicio, fecha_fin
                   FROM novedad_reemplazos WHERE activo = 1 ORDER BY id ASC"""
            ).fetchall()

        por_novedad = {}
        for r in reemplazos:
            por_novedad.setdefault(r["novedad_id"], []).append(dict(r))

        mapa = {}
        for nov in novedades:
            mapa[nov["id"]] = self.calcular_estado_cobertura(nov, por_novedad.get(nov["id"], []))
        return mapa

    def obtener_turnos_sin_cubrir(self, fecha_consulta, dotacion=None):
        """Devuelve las novedades activas que corresponden a turnos en la fecha dada junto con su estado de cobertura."""
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

            reemplazos_rows = connection.execute(
                """SELECT id, novedad_id, legajo_reemplazo, apellidos_nombres_reemplazo,
                          turnos_reemplazo, franco_reemplazo, fecha_inicio, fecha_fin
                   FROM novedad_reemplazos WHERE activo = 1 ORDER BY id ASC"""
            ).fetchall()

        reemplazos_por_nov = {}
        for r in reemplazos_rows:
            reemplazos_por_nov.setdefault(r["novedad_id"], []).append(dict(r))

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
                    item = dict(fila)
                    tramos = reemplazos_por_nov.get(fila["id"], [])
                    reemplazo_hoy = None
                    for t in tramos:
                        r_ini, r_fin_opt = self._rango_fechas_reemplazo(t["fecha_inicio"], t["fecha_fin"], fecha_fin_raw or None)
                        if r_ini and r_ini <= f_sol.date() and (r_fin_opt is None or f_sol.date() <= r_fin_opt):
                            reemplazo_hoy = t
                            break
                    if validators.es_turno_ciclo(fila["turnos"]):
                        item["cubierto"] = "N/A (Ciclo)"
                        item["reemplazo"] = ""
                    elif reemplazo_hoy:
                        item["cubierto"] = "Sí"
                        item["reemplazo"] = f"{reemplazo_hoy['legajo_reemplazo']} - {reemplazo_hoy['apellidos_nombres_reemplazo']} ({reemplazo_hoy['turnos_reemplazo']})"
                    else:
                        item["cubierto"] = "No"
                        item["reemplazo"] = ""
                    info_cob = self.calcular_estado_cobertura(fila, tramos)
                    item["cobertura_resumen"] = info_cob["resumen"]
                    item["reemplazos_detalle"] = info_cob["detalle"]
                    descubiertos.append(item)

        return descubiertos

    def obtener_estado_personal_ciclo(self, fecha_consulta, dotacion=None, especialidad=None):
        """Separa al personal de CICLO en libres (sin relevo) y ocupados (con relevo) en la fecha consultada."""
        import validators
        f_sol = validators.parsear_fecha(fecha_consulta, "Fecha de consulta", obligatorio=True).date()
        with self.store.read_connection() as connection:
            query_emp = """
                SELECT legajo, apellidos_nombres, especialidad, dotacion, turnos, franco, franco_horario
                FROM empleados
                WHERE activo = 1 AND UPPER(TRIM(turnos)) LIKE 'CICLO%'
            """
            params_emp = []
            if dotacion and dotacion != "Todas":
                query_emp += " AND LOWER(TRIM(dotacion)) = LOWER(TRIM(?))"
                params_emp.append(dotacion)
            if especialidad and especialidad != "Todas":
                query_emp += " AND LOWER(TRIM(especialidad)) = LOWER(TRIM(?))"
                params_emp.append(especialidad)
            query_emp += " ORDER BY dotacion, apellidos_nombres"
            empleados_ciclo = [dict(r) for r in connection.execute(query_emp, params_emp).fetchall()]

            reemplazos_activos = connection.execute(
                """SELECT nr.*, n.legajo AS legajo_titular, n.apellidos_nombres AS apellidos_titular,
                          n.turnos AS turno_titular, n.novedad AS tipo_novedad,
                          n.fecha_fin AS fecha_fin_novedad, n.dotacion AS dotacion_titular
                   FROM novedad_reemplazos nr
                   JOIN novedades n ON n.id = nr.novedad_id
                   WHERE nr.activo = 1 AND n.activo = 1"""
            ).fetchall()

        ocupados_por_legajo = {}
        for r in reemplazos_activos:
            r_ini, r_fin_opt = self._rango_fechas_reemplazo(r["fecha_inicio"], r["fecha_fin"], r["fecha_fin_novedad"])
            if r_ini and r_ini <= f_sol and (r_fin_opt is None or f_sol <= r_fin_opt):
                ocupados_por_legajo[int(r["legajo_reemplazo"])] = dict(r)

        libres = []
        ocupados = []
        for emp in empleados_ciclo:
            leg = int(emp["legajo"])
            asignacion = ocupados_por_legajo.get(leg)
            if asignacion:
                item = dict(emp)
                item["novedad_id"] = asignacion["novedad_id"]
                item["legajo_titular"] = asignacion["legajo_titular"]
                item["apellidos_titular"] = asignacion["apellidos_titular"]
                item["reemplazando_legajo"] = asignacion["legajo_titular"]
                item["reemplazando_nombre"] = asignacion["apellidos_titular"]
                item["reemplaza_a"] = f"{asignacion['legajo_titular']} - {asignacion['apellidos_titular']}"
                item["turno_titular"] = asignacion["turno_titular"]
                item["tipo_novedad"] = asignacion["tipo_novedad"]
                item["novedad_titular"] = asignacion["tipo_novedad"]
                item["fecha_inicio_relevo"] = asignacion["fecha_inicio"]
                item["fecha_fin_relevo"] = asignacion["fecha_fin"] or asignacion["fecha_fin_novedad"] or "-"
                item["reemplazo_desde"] = asignacion["fecha_inicio"]
                item["reemplazo_hasta"] = asignacion["fecha_fin"] or asignacion["fecha_fin_novedad"] or "Sin fecha fin"
                ocupados.append(item)
            else:
                libres.append(dict(emp))

        return {"libres": libres, "ocupados": ocupados}

    def guardar_jornada_ciclo(self, legajo, fecha, toma=None, deja=None, franco_anticipado=0, observaciones=None, user_id=None, windows_user=None):
        """Registra o actualiza la jornada/franco de un empleado de ciclo para una fecha determinada."""
        import validators
        fecha_obj = validators.parsear_fecha(fecha, "Fecha", obligatorio=True).date()
        fecha_iso = fecha_obj.strftime("%Y-%m-%d")
        toma_str = (toma or "").strip()
        deja_str = (deja or "").strip()
        obs_str = (observaciones or "").strip()
        fa_val = 1 if bool(franco_anticipado) else 0

        for val, campo in ((toma_str, "Toma"), (deja_str, "Deja")):
            if val:
                try:
                    partes = val.split(":")
                    if len(partes) != 2:
                        raise ValueError()
                    h, m = int(partes[0]), int(partes[1])
                    if not (0 <= h <= 23 and 0 <= m <= 59):
                        raise ValueError()
                except Exception:
                    raise ValueError(f"El horario de {campo} debe tener formato HH:MM (ej. 06:00).")

        with self.store.write_transaction() as connection:
            before = connection.execute(
                "SELECT * FROM ciclo_jornadas WHERE legajo=? AND fecha=?",
                (int(legajo), fecha_iso),
            ).fetchone()
            connection.execute(
                """INSERT INTO ciclo_jornadas
                   (legajo, fecha, toma, deja, franco_anticipado, observaciones, actualizado_en, usuario_windows, usuario_id)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                   ON CONFLICT(legajo, fecha) DO UPDATE SET
                       toma=excluded.toma,
                       deja=excluded.deja,
                       franco_anticipado=excluded.franco_anticipado,
                       observaciones=excluded.observaciones,
                       actualizado_en=excluded.actualizado_en,
                       usuario_windows=excluded.usuario_windows,
                       usuario_id=excluded.usuario_id""",
                (
                    int(legajo), fecha_iso, toma_str or None, deja_str or None, fa_val,
                    obs_str or None, self.store.now(), windows_user, user_id,
                ),
            )
            after = connection.execute(
                "SELECT * FROM ciclo_jornadas WHERE legajo=? AND fecha=?",
                (int(legajo), fecha_iso),
            ).fetchone()
            self._audit(
                connection, user_id, windows_user,
                "modificar" if before else "crear",
                "ciclo_jornada",
                after["id"] if after else None,
                dict(before) if before else None,
                dict(after) if after else None,
            )

    def obtener_jornadas_ciclo_rango(self, legajo, fecha_desde, fecha_hasta):
        """Devuelve un diccionario {YYYY-MM-DD: dict} con las jornadas del empleado en el rango."""
        with self.store.read_connection() as connection:
            rows = connection.execute(
                """SELECT id, legajo, fecha, toma, deja, franco_anticipado, observaciones
                   FROM ciclo_jornadas
                   WHERE legajo=? AND fecha BETWEEN ? AND ?""",
                (int(legajo), fecha_desde.strftime("%Y-%m-%d"), fecha_hasta.strftime("%Y-%m-%d")),
            ).fetchall()
            return {r["fecha"]: dict(r) for r in rows}

    def obtener_calendario_mensual_ciclo(self, legajo_ciclo, anio, mes):
        """Genera una fila por cada día del mes con cálculo de descansos reglamentarios y horas trabajadas."""
        anio = int(anio)
        mes = int(mes)
        _, num_dias = calendar.monthrange(anio, mes)

        with self.store.read_connection() as connection:
            emp = connection.execute(
                "SELECT legajo, apellidos_nombres, especialidad, dotacion, turnos, franco, franco_horario, jornada_horas FROM empleados WHERE legajo = ?",
                (int(legajo_ciclo),),
            ).fetchone()
            reemplazos = connection.execute(
                """SELECT nr.*, n.legajo AS legajo_titular, n.apellidos_nombres AS apellidos_titular,
                          n.turnos AS turno_titular, n.novedad AS tipo_novedad,
                          n.fecha_fin AS fecha_fin_novedad
                   FROM novedad_reemplazos nr
                   JOIN novedades n ON n.id = nr.novedad_id
                   WHERE nr.legajo_reemplazo = ? AND nr.activo = 1 AND n.activo = 1
                   ORDER BY nr.id ASC""",
                (int(legajo_ciclo),),
            ).fetchall()

        tramos_parsed = []
        for r in reemplazos:
            r_ini, r_fin_opt = self._rango_fechas_reemplazo(r["fecha_inicio"], r["fecha_fin"], r["fecha_fin_novedad"])
            if r_ini:
                tramos_parsed.append((r_ini, r_fin_opt, dict(r)))

        emp_jh = 6
        if emp:
            if "jornada_horas" in emp.keys() and emp["jornada_horas"] in (6, 7):
                emp_jh = int(emp["jornada_horas"])
            elif "guarda" in (emp["especialidad"] or "").lower():
                emp_jh = 7

        descanso_entre_jornada = 18 if emp_jh == 6 else 17
        descanso_post_franco = 42 if emp_jh == 6 else 41

        franco_emp = (emp["franco"] if emp else "") or ""
        franco_horario = ((emp["franco_horario"] if emp and "franco_horario" in emp.keys() else "") or "").strip()
        franco_norm = unicodedata.normalize("NFKD", franco_emp)
        franco_norm = "".join(c for c in franco_norm if not unicodedata.combining(c)).lower().strip()

        # Consultar jornadas cargadas desde 21 días antes del mes hasta el fin de mes
        fecha_primer_dia = date(anio, mes, 1)
        fecha_inicio_calc = fecha_primer_dia - timedelta(days=21)
        fecha_fin_calc = date(anio, mes, num_dias)
        jornadas_map = self.obtener_jornadas_ciclo_rango(legajo_ciclo, fecha_inicio_calc, fecha_fin_calc)

        def _parsear_hm(val_str):
            if not val_str:
                return None
            try:
                p = val_str.strip().split(":")
                return int(p[0]), int(p[1])
            except Exception:
                return None

        def _calc_hs_trabajadas(t_str, d_str):
            hm_t = _parsear_hm(t_str)
            hm_d = _parsear_hm(d_str)
            if not hm_t or not hm_d:
                return 0.0
            min_t = hm_t[0] * 60 + hm_t[1]
            min_d = hm_d[0] * 60 + hm_d[1]
            if min_d < min_t:
                min_d += 24 * 60
            return round((min_d - min_t) / 60.0, 2)

        def _calc_deja_dt(dia_date, t_str, d_str):
            hm_d = _parsear_hm(d_str)
            if not hm_d:
                return None
            hm_t = _parsear_hm(t_str)
            if hm_t:
                min_t = hm_t[0] * 60 + hm_t[1]
                min_d = hm_d[0] * 60 + hm_d[1]
                if min_d < min_t:
                    return datetime.combine(dia_date + timedelta(days=1), time(hm_d[0], hm_d[1]))
            return datetime.combine(dia_date, time(hm_d[0], hm_d[1]))

        dias_historia = {}
        acum_semana = 0.0

        d_step = fecha_inicio_calc
        while d_step <= fecha_fin_calc:
            f_iso = d_step.strftime("%Y-%m-%d")
            dia_semana_str = _DIAS_SEMANA[d_step.weekday()]
            dia_sem_norm = unicodedata.normalize("NFKD", dia_semana_str)
            dia_sem_norm = "".join(c for c in dia_sem_norm if not unicodedata.combining(c)).lower().strip()
            es_franco_reg = bool(franco_norm and franco_norm in dia_sem_norm)

            relevo_dia = None
            for r_ini, r_fin_opt, r_data in tramos_parsed:
                if r_ini <= d_step and (r_fin_opt is None or d_step <= r_fin_opt):
                    relevo_dia = r_data
                    break

            jornada_db = jornadas_map.get(f_iso)

            if relevo_dia:
                estado = "Reemplazando"
                es_franco = False
                novedad_id = relevo_dia["novedad_id"]
                reemplaza_a = f"{relevo_dia['legajo_titular']} - {relevo_dia['apellidos_titular']}"
                turno_cubierto = relevo_dia["turno_titular"] or "-"
                motivo = f"#{relevo_dia['novedad_id']} ({relevo_dia['tipo_novedad']})"
                periodo = f"{relevo_dia['fecha_inicio']} a {relevo_dia['fecha_fin'] or relevo_dia['fecha_fin_novedad'] or 'abierto'}"
                reemplazando_legajo = relevo_dia["legajo_titular"]
                reemplazando_nombre = relevo_dia["apellidos_titular"]
                novedad_titular = relevo_dia["tipo_novedad"]
                tramo_inicio = relevo_dia["fecha_inicio"]
                tramo_fin = relevo_dia["fecha_fin"] or relevo_dia["fecha_fin_novedad"]
                toma = jornada_db.get("toma") or "" if jornada_db else ""
                deja = jornada_db.get("deja") or "" if jornada_db else ""
            else:
                novedad_id = ""
                reemplaza_a = "-"
                turno_cubierto = "-"
                motivo = "-"
                periodo = "-"
                reemplazando_legajo = None
                reemplazando_nombre = None
                novedad_titular = None
                tramo_inicio = None
                tramo_fin = None
                if jornada_db and jornada_db.get("franco_anticipado"):
                    es_franco = True
                    estado = "Franco Anticipado"
                    toma = ""
                    deja = ""
                elif es_franco_reg:
                    es_franco = True
                    estado = f"Franco (desde {franco_horario})" if franco_horario else "Franco"
                    toma = ""
                    deja = ""
                else:
                    es_franco = False
                    estado = "Disponible"
                    toma = jornada_db.get("toma") or "" if jornada_db else ""
                    deja = jornada_db.get("deja") or "" if jornada_db else ""

            hs_trab = _calc_hs_trabajadas(toma, deja) if (toma and deja) else 0.0
            deja_dt = _calc_deja_dt(d_step, toma, deja) if deja else None

            # Acumulador semanal
            if es_franco:
                acum_semana = 0.0
                hs_semana_str = "-"
            else:
                acum_semana += hs_trab
                hs_semana_str = f"{acum_semana:.1f} hs" if acum_semana > 0 else "-"

            # Hora disponible para d_step
            if es_franco:
                hora_disp_str = "-"
            else:
                d_prev = d_step - timedelta(days=1)
                disp_dt = None
                if d_prev in dias_historia and not dias_historia[d_prev]["es_franco"]:
                    prev_deja_dt = dias_historia[d_prev]["deja_dt"]
                    if prev_deja_dt:
                        disp_dt = prev_deja_dt + timedelta(hours=descanso_entre_jornada)
                else:
                    d_busca = d_prev
                    while d_busca in dias_historia and dias_historia[d_busca]["es_franco"]:
                        d_busca -= timedelta(days=1)
                    if d_busca in dias_historia and dias_historia[d_busca]["deja_dt"]:
                        prev_deja_dt = dias_historia[d_busca]["deja_dt"]
                        disp_dt = prev_deja_dt + timedelta(hours=descanso_post_franco)

                if disp_dt:
                    if disp_dt.date() == d_step:
                        hora_disp_str = disp_dt.strftime("%H:%M")
                    elif disp_dt.date() < d_step:
                        hora_disp_str = "00:00"
                    else:
                        hora_disp_str = disp_dt.strftime("%d/%m %H:%M")
                else:
                    hora_disp_str = "-"

            dias_historia[d_step] = {
                "d": d_step,
                "es_franco": es_franco,
                "estado": estado,
                "toma": toma,
                "deja": deja,
                "deja_dt": deja_dt,
                "hs_trabajadas": hs_trab,
                "hs_trabajadas_str": f"{hs_trab:.1f} hs" if hs_trab > 0 else "-",
                "hs_semana_str": hs_semana_str,
                "hora_disponible": hora_disp_str,
                "franco_anticipado": bool(jornada_db and jornada_db.get("franco_anticipado")),
                "observaciones": jornada_db.get("observaciones") or "" if jornada_db else "",
                "novedad_id": novedad_id,
                "reemplaza_a": reemplaza_a,
                "turno_cubierto": turno_cubierto,
                "motivo": motivo,
                "periodo": periodo,
                "reemplazando_legajo": reemplazando_legajo,
                "reemplazando_nombre": reemplazando_nombre,
                "novedad_titular": novedad_titular,
                "tramo_inicio": tramo_inicio,
                "tramo_fin": tramo_fin,
            }

            d_step += timedelta(days=1)

        dias_mes = []
        for dia_num in range(1, num_dias + 1):
            f_actual = date(anio, mes, dia_num)
            info = dias_historia[f_actual]
            dias_mes.append({
                "dia": dia_num,
                "fecha": f_actual.strftime("%d/%m/%Y"),
                "fecha_iso": f_actual.strftime("%Y-%m-%d"),
                "dia_semana": _DIAS_SEMANA[f_actual.weekday()],
                "estado": info["estado"],
                "hora_disponible": info["hora_disponible"],
                "toma": info["toma"] or "-",
                "deja": info["deja"] or "-",
                "horas_trabajadas": info["hs_trabajadas_str"],
                "hs_semana": info["hs_semana_str"],
                "novedad_id": info["novedad_id"],
                "reemplaza_a": info["reemplaza_a"],
                "turno_cubierto": info["turno_cubierto"],
                "motivo": info["motivo"],
                "periodo": info["periodo"],
                "reemplazando_legajo": info["reemplazando_legajo"],
                "reemplazando_nombre": info["reemplazando_nombre"],
                "novedad_titular": info["novedad_titular"],
                "tramo_inicio": info["tramo_inicio"],
                "tramo_fin": info["tramo_fin"],
                "franco_anticipado": info["franco_anticipado"],
                "observaciones": info["observaciones"],
            })

        return dias_mes

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
