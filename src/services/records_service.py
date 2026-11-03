"""Casos de uso y operaciones de negocio para RENO v2.0.0.

Orquesta la persistencia local en SQLite, las reglas de dominio,
la auditoría y el encolamiento para sincronización Supabase.
"""

import calendar
import json
import sqlite3
import unicodedata
from datetime import date, datetime, time, timedelta
from typing import Any, Dict, List, Optional, Tuple, Union

from src.core.logger import get_logger
from src.data.database import DatabaseManager
from src.domain.enums import EstadoCobertura, EstadoDiaCiclo, SyncOperation
from src.domain.models import (
    CambioTurno,
    CicloJornada,
    DiaCalendarioCiclo,
    Empleado,
    Novedad,
    PersonalEstacion,
    TipoNovedad,
    TramoReemplazo,
)
from src.domain.validators import (
    DIAS_SEMANA,
    DIAS_SEMANA_NORM,
    calcular_descansos_ciclo,
    es_turno_ciclo,
    normalizar_dia_semana,
    parsear_fecha,
    validar_tramos_reemplazo,
)

logger = get_logger("RecordsService")


class RecordsService:
    """Servicio principal de negocio para novedades, cambios de turno, ciclo y dashboard."""

    def __init__(self, db: DatabaseManager):
        self.db = db

    # ---------------------------------------------------------
    # Auditoría y Cola de Sincronización
    # ---------------------------------------------------------

    def _audit(
        self,
        conn: sqlite3.Connection,
        user_id: Optional[int],
        windows_user: Optional[str],
        action: str,
        entity: str,
        entity_id: Optional[int],
        before: Optional[Dict[str, Any]],
        after: Optional[Dict[str, Any]],
    ) -> None:
        conn.execute(
            """INSERT INTO auditoria
               (usuario_id, usuario_windows, accion, entidad, entidad_id,
                datos_anteriores, datos_nuevos, creado_en)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                user_id,
                windows_user,
                action,
                entity,
                entity_id,
                json.dumps(before, ensure_ascii=False, default=str) if before else None,
                json.dumps(after, ensure_ascii=False, default=str) if after else None,
                self.db.now(),
            ),
        )

    def _enqueue_sync(
        self,
        conn: sqlite3.Connection,
        tabla: str,
        registro_id: Any,
        operacion: str,
        payload: Optional[Dict[str, Any]] = None,
    ) -> None:
        conn.execute(
            """INSERT INTO sync_outbox (tabla, registro_pk, operacion, payload_json, creado_en)
               VALUES (?, ?, ?, ?, ?)
               ON CONFLICT(tabla, registro_pk) DO UPDATE SET
                   operacion = excluded.operacion,
                   payload_json = excluded.payload_json,
                   creado_en = excluded.creado_en""",
            (
                tabla,
                str(registro_id),
                operacion,
                json.dumps(payload, ensure_ascii=False, default=str) if payload else None,
                self.db.now(),
            ),
        )

    # ---------------------------------------------------------
    # Fechas y Ventanas de Edición / Borrado
    # ---------------------------------------------------------

    def parsear_registrado_en(self, valor: Any) -> Optional[datetime]:
        """Parsea una fecha de registro en formatos ISO o DD/MM/YYYY con o sin hora."""
        if valor is None:
            return None
        texto = str(valor).strip()
        if not texto:
            return None

        if len(texto) >= 10 and texto[4] == "-" and texto[7] == "-":
            try:
                return datetime.fromisoformat(texto)
            except ValueError:
                pass
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
            "%d/%m/%Y %H:%M:%S",
            "%d/%m/%Y %H:%M",
            "%d/%m/%Y",
            "%Y-%m-%d %H:%M:%S",
            "%Y-%m-%d %H:%M",
            "%Y-%m-%d",
        ):
            try:
                return datetime.strptime(texto, formato)
            except ValueError:
                continue

        try:
            return datetime(1899, 12, 30) + timedelta(days=float(texto))
        except ValueError:
            return None

    def _horas_limite(self, conn: sqlite3.Connection, clave: str, default: int) -> int:
        row = conn.execute("SELECT valor FROM configuracion WHERE clave=?", (clave,)).fetchone()
        try:
            return max(1, int(row[0])) if row else default
        except (TypeError, ValueError):
            return default

    def dentro_de_ventana(self, registrado_en: Any, clave: str) -> bool:
        default = 24 if clave == "editar_horas" else 72
        dt = self.parsear_registrado_en(registrado_en)
        if dt is None:
            return False
        with self.db.read_connection() as conn:
            horas = self._horas_limite(conn, clave, default)
        return (datetime.now() - dt).total_seconds() <= horas * 3600

    # ---------------------------------------------------------
    # Tipos de Novedad CRUD
    # ---------------------------------------------------------

    def listar_tipos(self, incluir_inactivos: bool = True) -> List[Dict[str, Any]]:
        with self.db.read_connection() as conn:
            query = "SELECT id, nombre, dias, exportar_descubierto, activo FROM tipos_novedad"
            if not incluir_inactivos:
                query += " WHERE activo=1"
            query += " ORDER BY nombre"
            rows = conn.execute(query).fetchall()
            return [dict(r) for r in rows]

    def crear_tipo(
        self,
        nombre: str,
        dias: int = 0,
        exportar_descubierto: int = 1,
        user_id: Optional[int] = None,
        windows_user: Optional[str] = None,
    ) -> int:
        nombre = (nombre or "").strip()
        if not nombre:
            raise ValueError("El nombre no puede estar vacío.")
        try:
            dias_val = max(0, int(dias or 0))
        except (ValueError, TypeError):
            dias_val = 0
        exportar_val = 1 if exportar_descubierto else 0

        with self.db.write_transaction() as conn:
            cur = conn.execute(
                "INSERT INTO tipos_novedad(nombre, dias, exportar_descubierto) VALUES (?, ?, ?)",
                (nombre, dias_val, exportar_val),
            )
            item_id = cur.lastrowid
            payload = {"nombre": nombre, "dias": dias_val, "exportar_descubierto": exportar_val}
            self._audit(conn, user_id, windows_user, "crear", "tipo_novedad", item_id, None, payload)
            self._enqueue_sync(conn, "tipos_novedad", item_id, "INSERT", payload)
            return item_id

    def actualizar_tipo(
        self,
        type_id: int,
        nombre: str,
        dias: int = 0,
        exportar_descubierto: int = 1,
        activo: int = 1,
        user_id: Optional[int] = None,
        windows_user: Optional[str] = None,
    ) -> None:
        nombre = (nombre or "").strip()
        if not nombre:
            raise ValueError("El nombre no puede estar vacío.")
        try:
            dias_val = max(0, int(dias or 0))
        except (ValueError, TypeError):
            dias_val = 0
        exportar_val = 1 if exportar_descubierto else 0

        with self.db.write_transaction() as conn:
            before = conn.execute("SELECT * FROM tipos_novedad WHERE id=?", (type_id,)).fetchone()
            if not before:
                raise ValueError("El tipo de novedad no existe.")
            conn.execute(
                "UPDATE tipos_novedad SET nombre=?, dias=?, exportar_descubierto=?, activo=? WHERE id=?",
                (nombre, dias_val, exportar_val, int(activo), type_id),
            )
            after = {"nombre": nombre, "dias": dias_val, "exportar_descubierto": exportar_val, "activo": int(activo)}
            self._audit(conn, user_id, windows_user, "modificar", "tipo_novedad", type_id, dict(before), after)
            self._enqueue_sync(conn, "tipos_novedad", type_id, "UPDATE", after)

    # ---------------------------------------------------------
    # Dotaciones CRUD
    # ---------------------------------------------------------

    def listar_dotaciones(self, incluir_inactivos: bool = True) -> List[Dict[str, Any]]:
        with self.db.read_connection() as conn:
            query = "SELECT id, nombre, activo FROM dotaciones"
            if not incluir_inactivos:
                query += " WHERE activo=1"
            rows = conn.execute(query + " ORDER BY nombre").fetchall()
            return [dict(r) for r in rows]

    def crear_dotacion(
        self,
        nombre: str,
        user_id: Optional[int] = None,
        windows_user: Optional[str] = None,
    ) -> int:
        nombre = (nombre or "").strip()
        if not nombre:
            raise ValueError("La dotación no puede estar vacía.")
        with self.db.write_transaction() as conn:
            cur = conn.execute("INSERT INTO dotaciones(nombre) VALUES (?)", (nombre,))
            item_id = cur.lastrowid
            payload = {"nombre": nombre}
            self._audit(conn, user_id, windows_user, "crear", "dotacion", item_id, None, payload)
            self._enqueue_sync(conn, "dotaciones", item_id, "INSERT", payload)
            return item_id

    def actualizar_dotacion(
        self,
        dotacion_id: int,
        nombre: str,
        activo: int,
        user_id: Optional[int] = None,
        windows_user: Optional[str] = None,
    ) -> None:
        nombre = (nombre or "").strip()
        if not nombre:
            raise ValueError("La dotación no puede estar vacía.")
        with self.db.write_transaction() as conn:
            before = conn.execute("SELECT id, nombre, activo FROM dotaciones WHERE id=?", (dotacion_id,)).fetchone()
            if not before:
                raise ValueError("La dotación no existe.")
            conn.execute("UPDATE dotaciones SET nombre=?, activo=? WHERE id=?", (nombre, int(activo), dotacion_id))
            after = {"nombre": nombre, "activo": int(activo)}
            self._audit(conn, user_id, windows_user, "modificar", "dotacion", dotacion_id, dict(before), after)
            self._enqueue_sync(conn, "dotaciones", dotacion_id, "UPDATE", after)

    # ---------------------------------------------------------
    # Personal de Estación CRUD
    # ---------------------------------------------------------

    def listar_personal_estacion(self, incluir_inactivos: bool = True) -> List[Dict[str, Any]]:
        with self.db.read_connection() as conn:
            query = "SELECT id, nombre, dotacion, activo FROM personal_estacion"
            if not incluir_inactivos:
                query += " WHERE activo=1"
            rows = conn.execute(query + " ORDER BY nombre").fetchall()
            return [dict(r) for r in rows]

    def crear_personal_estacion(
        self,
        nombre: str,
        dotacion: Optional[str] = None,
        user_id: Optional[int] = None,
        windows_user: Optional[str] = None,
    ) -> int:
        nombre = (nombre or "").strip()
        if not nombre:
            raise ValueError("El nombre no puede estar vacío.")
        dotacion = (dotacion or "").strip() or None
        with self.db.write_transaction() as conn:
            cur = conn.execute(
                "INSERT INTO personal_estacion(nombre, dotacion) VALUES (?, ?)",
                (nombre, dotacion),
            )
            item_id = cur.lastrowid
            payload = {"nombre": nombre, "dotacion": dotacion}
            self._audit(conn, user_id, windows_user, "crear", "personal_estacion", item_id, None, payload)
            self._enqueue_sync(conn, "personal_estacion", item_id, "INSERT", payload)
            return item_id

    def actualizar_personal_estacion(
        self,
        item_id: int,
        nombre: str,
        activo: int,
        user_id: Optional[int] = None,
        windows_user: Optional[str] = None,
        *,
        dotacion: Optional[str] = None,
        cambiar_dotacion: bool = False,
    ) -> None:
        nombre = (nombre or "").strip()
        if cambiar_dotacion:
            dotacion = (dotacion or "").strip() or None
        with self.db.write_transaction() as conn:
            before = conn.execute("SELECT id, nombre, dotacion, activo FROM personal_estacion WHERE id=?", (item_id,)).fetchone()
            if not before:
                raise ValueError("El personal de estación no existe.")
            despues_dotacion = dotacion if cambiar_dotacion else before["dotacion"]
            conn.execute(
                "UPDATE personal_estacion SET nombre=?, dotacion=?, activo=? WHERE id=?",
                (nombre, despues_dotacion, int(activo), item_id),
            )
            after = {"nombre": nombre, "dotacion": despues_dotacion, "activo": int(activo)}
            self._audit(conn, user_id, windows_user, "modificar", "personal_estacion", item_id, dict(before), after)
            self._enqueue_sync(conn, "personal_estacion", item_id, "UPDATE", after)

    # ---------------------------------------------------------
    # Empleados (BASE) CRUD
    # ---------------------------------------------------------

    def listar_empleados(self, incluir_inactivos: bool = True) -> List[Dict[str, Any]]:
        with self.db.read_connection() as conn:
            query = """SELECT id, legajo, apellidos_nombres, especialidad, dotacion, turnos,
                              franco, franco_horario, jornada_horas, activo
                       FROM empleados"""
            if not incluir_inactivos:
                query += " WHERE activo=1"
            rows = conn.execute(query + " ORDER BY apellidos_nombres").fetchall()
            return [dict(r) for r in rows]

    def obtener_empleado(self, empleado_id: int) -> Optional[Dict[str, Any]]:
        with self.db.read_connection() as conn:
            row = conn.execute("SELECT * FROM empleados WHERE id=?", (empleado_id,)).fetchone()
            return dict(row) if row else None

    def obtener_empleado_por_legajo(self, legajo: int) -> Optional[Dict[str, Any]]:
        with self.db.read_connection() as conn:
            row = conn.execute("SELECT * FROM empleados WHERE legajo=?", (int(legajo),)).fetchone()
            return dict(row) if row else None

    def _validar_empleado(self, data: Dict[str, Any]) -> Tuple[int, str]:
        try:
            legajo = int(str(data.get("legajo")).strip())
        except (TypeError, ValueError):
            raise ValueError("El legajo debe ser un número entero.")
        apellidos = (data.get("apellidos_nombres") or "").strip()
        if not apellidos:
            raise ValueError("Los apellidos y nombres son obligatorios.")
        return legajo, apellidos

    def crear_empleado(
        self,
        data: Dict[str, Any],
        user_id: Optional[int] = None,
        windows_user: Optional[str] = None,
    ) -> int:
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
        with self.db.write_transaction() as conn:
            try:
                cur = conn.execute(
                    """INSERT INTO empleados
                       (legajo, apellidos_nombres, especialidad, dotacion, turnos, franco, franco_horario, jornada_horas, actualizado_en)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                    (
                        valores["legajo"],
                        valores["apellidos_nombres"],
                        valores["especialidad"],
                        valores["dotacion"],
                        valores["turnos"],
                        valores["franco"],
                        valores["franco_horario"],
                        valores["jornada_horas"],
                        self.db.now(),
                    ),
                )
                item_id = cur.lastrowid
            except sqlite3.IntegrityError:
                raise ValueError(f"Ya existe un empleado con el legajo {legajo}.")
            self._audit(conn, user_id, windows_user, "crear", "empleado", item_id, None, valores)
            self._enqueue_sync(conn, "empleados", item_id, "INSERT", valores)
            return item_id

    def actualizar_empleado(
        self,
        empleado_id: int,
        data: Dict[str, Any],
        user_id: Optional[int] = None,
        windows_user: Optional[str] = None,
    ) -> None:
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
        with self.db.write_transaction() as conn:
            before = conn.execute("SELECT * FROM empleados WHERE id=?", (empleado_id,)).fetchone()
            if not before:
                raise ValueError("El empleado no existe.")
            try:
                conn.execute(
                    """UPDATE empleados SET legajo=?, apellidos_nombres=?, especialidad=?, dotacion=?,
                                            turnos=?, franco=?, franco_horario=?, jornada_horas=?, actualizado_en=?
                       WHERE id=?""",
                    (
                        valores["legajo"],
                        valores["apellidos_nombres"],
                        valores["especialidad"],
                        valores["dotacion"],
                        valores["turnos"],
                        valores["franco"],
                        valores["franco_horario"],
                        valores["jornada_horas"],
                        self.db.now(),
                        empleado_id,
                    ),
                )
            except sqlite3.IntegrityError:
                raise ValueError(f"Ya existe un empleado con el legajo {legajo}.")
            after = dict(conn.execute("SELECT * FROM empleados WHERE id=?", (empleado_id,)).fetchone())
            self._audit(conn, user_id, windows_user, "modificar", "empleado", empleado_id, dict(before), after)
            self._enqueue_sync(conn, "empleados", empleado_id, "UPDATE", after)

    def cambiar_estado_empleado(
        self,
        empleado_id: int,
        user_id: Optional[int] = None,
        windows_user: Optional[str] = None,
    ) -> int:
        with self.db.write_transaction() as conn:
            before = conn.execute("SELECT * FROM empleados WHERE id=?", (empleado_id,)).fetchone()
            if not before:
                raise ValueError("El empleado no existe.")
            nuevo_estado = 1 - before["activo"]
            conn.execute("UPDATE empleados SET activo=? WHERE id=?", (nuevo_estado, empleado_id))
            accion = "desactivar" if nuevo_estado == 0 else "reactivar"
            self._audit(conn, user_id, windows_user, accion, "empleado", empleado_id, dict(before), {"activo": nuevo_estado})
            self._enqueue_sync(conn, "empleados", empleado_id, "UPDATE", {"activo": nuevo_estado})
            return nuevo_estado

    # ---------------------------------------------------------
    # Destinatarios de Informe CRUD
    # ---------------------------------------------------------

    def listar_destinatarios_informe(self, incluir_inactivos: bool = True) -> List[Dict[str, Any]]:
        with self.db.read_connection() as conn:
            query = "SELECT id, nombre, email, activo FROM destinatarios_informe"
            if not incluir_inactivos:
                query += " WHERE activo=1"
            rows = conn.execute(query + " ORDER BY nombre").fetchall()
            return [dict(r) for r in rows]

    def crear_destinatario_informe(
        self,
        nombre: str,
        email: str,
        user_id: Optional[int] = None,
        windows_user: Optional[str] = None,
    ) -> int:
        nombre, email = (nombre or "").strip(), (email or "").strip()
        if not email or "@" not in email:
            raise ValueError("Ingrese un correo electrónico válido.")
        with self.db.write_transaction() as conn:
            cur = conn.execute(
                "INSERT INTO destinatarios_informe(nombre, email) VALUES (?, ?)",
                (nombre, email),
            )
            item_id = cur.lastrowid
            payload = {"nombre": nombre, "email": email}
            self._audit(conn, user_id, windows_user, "crear", "destinatario_informe", item_id, None, payload)
            return item_id

    def actualizar_destinatario_informe(
        self,
        item_id: int,
        nombre: str,
        email: str,
        activo: int,
        user_id: Optional[int] = None,
        windows_user: Optional[str] = None,
    ) -> None:
        nombre, email = (nombre or "").strip(), (email or "").strip()
        if not email or "@" not in email:
            raise ValueError("Ingrese un correo electrónico válido.")
        with self.db.write_transaction() as conn:
            before = conn.execute("SELECT id, nombre, email, activo FROM destinatarios_informe WHERE id=?", (item_id,)).fetchone()
            if not before:
                raise ValueError("El destinatario no existe.")
            conn.execute(
                "UPDATE destinatarios_informe SET nombre=?, email=?, activo=? WHERE id=?",
                (nombre, email, int(activo), item_id),
            )
            after = {"nombre": nombre, "email": email, "activo": int(activo)}
            self._audit(conn, user_id, windows_user, "modificar", "destinatario_informe", item_id, dict(before), after)

    # ---------------------------------------------------------
    # Novedades CRUD
    # ---------------------------------------------------------

    def listar_novedades(
        self,
        filtros: Optional[Dict[str, Any]] = None,
        incluir_inactivos: bool = False,
    ) -> List[Dict[str, Any]]:
        with self.db.read_connection() as conn:
            query = """SELECT id, registrado_en, legajo, apellidos_nombres, especialidad,
                              dotacion, turnos, franco, novedad, fecha_inicio, fecha_fin,
                              referencia_estacion, supervisor, observaciones, usuario_windows,
                              usuario_id, activo
                       FROM novedades"""
            conds = []
            params = []
            if not incluir_inactivos:
                conds.append("activo = 1")

            if filtros:
                if filtros.get("legajo"):
                    conds.append("legajo = ?")
                    params.append(int(filtros["legajo"]))
                if filtros.get("dotacion"):
                    conds.append("LOWER(TRIM(dotacion)) = LOWER(TRIM(?))")
                    params.append(str(filtros["dotacion"]))
                if filtros.get("novedad"):
                    conds.append("LOWER(TRIM(novedad)) = LOWER(TRIM(?))")
                    params.append(str(filtros["novedad"]))

            if conds:
                query += " WHERE " + " AND ".join(conds)
            query += " ORDER BY id DESC"

            rows = conn.execute(query, params).fetchall()
            return [dict(r) for r in rows]

    def obtener_novedad(self, record_id: int) -> Optional[Dict[str, Any]]:
        with self.db.read_connection() as conn:
            row = conn.execute("SELECT * FROM novedades WHERE id=?", (record_id,)).fetchone()
            return dict(row) if row else None

    def crear_novedad(
        self,
        data: Dict[str, Any],
        user_id: Optional[int] = None,
        windows_user: Optional[str] = None,
    ) -> int:
        fields = (
            "legajo", "apellidos_nombres", "especialidad", "dotacion", "turnos",
            "franco", "novedad", "fecha_inicio", "fecha_fin", "referencia_estacion",
            "supervisor", "observaciones",
        )
        registrado_en = data.get("registrado_en") or self.db.now()
        with self.db.write_transaction() as conn:
            columns = ("registrado_en",) + fields + ("usuario_windows", "usuario_id", "activo")
            placeholders = ", ".join("?" for _ in columns)
            values = (
                registrado_en,
                int(str(data["legajo"]).strip()),
                str(data.get("apellidos_nombres") or "").strip(),
                str(data.get("especialidad") or "").strip(),
                str(data.get("dotacion") or "").strip(),
                str(data.get("turnos") or "").strip(),
                str(data.get("franco") or "").strip(),
                str(data.get("novedad") or "").strip(),
                str(data.get("fecha_inicio") or "").strip(),
                str(data.get("fecha_fin") or "").strip(),
                str(data.get("referencia_estacion") or "").strip(),
                str(data.get("supervisor") or "").strip(),
                str(data.get("observaciones") or "").strip(),
                windows_user,
                user_id,
                1,
            )
            cur = conn.execute(f"INSERT INTO novedades ({', '.join(columns)}) VALUES ({placeholders})", values)
            record_id = cur.lastrowid
            after = dict(conn.execute("SELECT * FROM novedades WHERE id=?", (record_id,)).fetchone())
            self._audit(conn, user_id, windows_user, "crear", "novedad", record_id, None, after)
            self._enqueue_sync(conn, "novedades", record_id, "INSERT", after)
            return record_id

    def actualizar_novedad(
        self,
        record_id: int,
        data: Dict[str, Any],
        user_id: Optional[int] = None,
        windows_user: Optional[str] = None,
        bypass_window: bool = False,
    ) -> None:
        fields = (
            "legajo", "apellidos_nombres", "especialidad", "dotacion", "turnos",
            "franco", "novedad", "fecha_inicio", "fecha_fin", "referencia_estacion",
            "supervisor", "observaciones",
        )
        with self.db.write_transaction() as conn:
            before = conn.execute("SELECT * FROM novedades WHERE id=?", (record_id,)).fetchone()
            if not before:
                raise ValueError("La novedad no existe.")
            if not bypass_window and not self.dentro_de_ventana(before["registrado_en"], "editar_horas"):
                raise ValueError("El registro superó el tiempo permitido para editar.")

            assignments = ", ".join(f"{field}=?" for field in fields)
            values = tuple(data[field] for field in fields) + (record_id,)
            conn.execute(f"UPDATE novedades SET {assignments} WHERE id=?", values)
            after = dict(conn.execute("SELECT * FROM novedades WHERE id=?", (record_id,)).fetchone())
            self._audit(conn, user_id, windows_user, "modificar", "novedad", record_id, dict(before), after)
            self._enqueue_sync(conn, "novedades", record_id, "UPDATE", after)

    def actualizar_fecha_fin_novedad(
        self,
        record_id: int,
        fecha_fin: Optional[str],
        nueva_observacion: Optional[str] = None,
        user_id: Optional[int] = None,
        windows_user: Optional[str] = None,
    ) -> None:
        fecha_fin_str = (fecha_fin or "").strip()
        nueva_obs = (nueva_observacion or "").strip()
        if not fecha_fin_str and not nueva_obs:
            raise ValueError("Debe ingresar la fecha de fin o una observación.")

        val_fin = None
        if fecha_fin_str:
            val_fin = parsear_fecha(fecha_fin_str, "Fecha de Fin", obligatorio=True)

        with self.db.write_transaction() as conn:
            before = conn.execute("SELECT * FROM novedades WHERE id=?", (record_id,)).fetchone()
            if not before:
                raise ValueError("La novedad no existe.")

            fecha_fin_anterior = (before["fecha_fin"] or "").strip()
            if fecha_fin_str:
                if fecha_fin_anterior and fecha_fin_anterior != "-":
                    raise ValueError("El registro ya cuenta con una fecha de fin.")
                fecha_inicio_str = (before["fecha_inicio"] or "").strip()
                if fecha_inicio_str:
                    val_ini = parsear_fecha(fecha_inicio_str, "Fecha de Inicio", obligatorio=False)
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

            conn.execute(
                "UPDATE novedades SET fecha_fin=?, observaciones=? WHERE id=?",
                (fecha_guardar, obs_final, record_id),
            )
            after = dict(conn.execute("SELECT * FROM novedades WHERE id=?", (record_id,)).fetchone())
            self._audit(conn, user_id, windows_user, "modificar", "novedad", record_id, dict(before), after)
            self._enqueue_sync(conn, "novedades", record_id, "UPDATE", after)

    def eliminar_novedad(
        self,
        record_id: int,
        user_id: Optional[int] = None,
        windows_user: Optional[str] = None,
        bypass_window: bool = False,
    ) -> None:
        with self.db.write_transaction() as conn:
            before = conn.execute("SELECT * FROM novedades WHERE id=?", (record_id,)).fetchone()
            if not before:
                raise ValueError("La novedad no existe.")
            if not bypass_window and not self.dentro_de_ventana(before["registrado_en"], "eliminar_horas"):
                raise ValueError("El registro superó el tiempo permitido para eliminar.")
            conn.execute("UPDATE novedades SET activo=0 WHERE id=?", (record_id,))
            after = {"activo": 0}
            self._audit(conn, user_id, windows_user, "eliminar", "novedad", record_id, dict(before), after)
            self._enqueue_sync(conn, "novedades", record_id, "UPDATE", after)

    # ---------------------------------------------------------
    # Cambios de Turno CRUD
    # ---------------------------------------------------------

    def listar_cambios(
        self,
        filtros: Optional[Dict[str, Any]] = None,
        incluir_inactivos: bool = False,
    ) -> List[Dict[str, Any]]:
        with self.db.read_connection() as conn:
            query = """SELECT id, registrado_en, legajo_1, apellidos_nombres_1, especialidad_1,
                              dotacion_1, turnos_1, franco_1, legajo_2, apellidos_nombres_2,
                              especialidad_2, dotacion_2, turnos_2, franco_2, fecha_cambio,
                              referencia_estacion, supervisor, observaciones, usuario_windows,
                              usuario_id, activo
                       FROM cambios_turno"""
            conds = []
            params = []
            if not incluir_inactivos:
                conds.append("activo = 1")

            if filtros:
                if filtros.get("legajo"):
                    conds.append("(legajo_1 = ? OR legajo_2 = ?)")
                    params.extend([int(filtros["legajo"]), int(filtros["legajo"])])
                if filtros.get("dotacion"):
                    conds.append("(LOWER(TRIM(dotacion_1)) = LOWER(TRIM(?)) OR LOWER(TRIM(dotacion_2)) = LOWER(TRIM(?)))")
                    params.extend([str(filtros["dotacion"]), str(filtros["dotacion"])])

            if conds:
                query += " WHERE " + " AND ".join(conds)
            query += " ORDER BY id DESC"

            rows = conn.execute(query, params).fetchall()
            return [dict(r) for r in rows]

    def obtener_cambio(self, record_id: int) -> Optional[Dict[str, Any]]:
        with self.db.read_connection() as conn:
            row = conn.execute("SELECT * FROM cambios_turno WHERE id=?", (record_id,)).fetchone()
            return dict(row) if row else None

    def crear_cambio(
        self,
        data: Dict[str, Any],
        user_id: Optional[int] = None,
        windows_user: Optional[str] = None,
    ) -> int:
        fields = (
            "legajo_1", "apellidos_nombres_1", "especialidad_1", "dotacion_1", "turnos_1", "franco_1",
            "legajo_2", "apellidos_nombres_2", "especialidad_2", "dotacion_2", "turnos_2", "franco_2",
            "fecha_cambio", "referencia_estacion", "supervisor", "observaciones",
        )
        registrado_en = data.get("registrado_en") or self.db.now()
        with self.db.write_transaction() as conn:
            columns = ("registrado_en",) + fields + ("usuario_windows", "usuario_id", "activo")
            placeholders = ", ".join("?" for _ in columns)
            values = (
                registrado_en,
                int(str(data["legajo_1"]).strip()),
                str(data.get("apellidos_nombres_1") or "").strip(),
                str(data.get("especialidad_1") or "").strip(),
                str(data.get("dotacion_1") or "").strip(),
                str(data.get("turnos_1") or "").strip(),
                str(data.get("franco_1") or "").strip(),
                int(str(data["legajo_2"]).strip()),
                str(data.get("apellidos_nombres_2") or "").strip(),
                str(data.get("especialidad_2") or "").strip(),
                str(data.get("dotacion_2") or "").strip(),
                str(data.get("turnos_2") or "").strip(),
                str(data.get("franco_2") or "").strip(),
                str(data.get("fecha_cambio") or "").strip(),
                str(data.get("referencia_estacion") or "").strip(),
                str(data.get("supervisor") or "").strip(),
                str(data.get("observaciones") or "").strip(),
                windows_user,
                user_id,
                1,
            )
            cur = conn.execute(f"INSERT INTO cambios_turno ({', '.join(columns)}) VALUES ({placeholders})", values)
            record_id = cur.lastrowid
            after = dict(conn.execute("SELECT * FROM cambios_turno WHERE id=?", (record_id,)).fetchone())
            self._audit(conn, user_id, windows_user, "crear", "cambio_turno", record_id, None, after)
            self._enqueue_sync(conn, "cambios_turno", record_id, "INSERT", after)
            return record_id

    def actualizar_cambio(
        self,
        record_id: int,
        data: Dict[str, Any],
        user_id: Optional[int] = None,
        windows_user: Optional[str] = None,
        bypass_window: bool = False,
    ) -> None:
        fields = (
            "legajo_1", "apellidos_nombres_1", "especialidad_1", "dotacion_1", "turnos_1", "franco_1",
            "legajo_2", "apellidos_nombres_2", "especialidad_2", "dotacion_2", "turnos_2", "franco_2",
            "fecha_cambio", "referencia_estacion", "supervisor", "observaciones",
        )
        with self.db.write_transaction() as conn:
            before = conn.execute("SELECT * FROM cambios_turno WHERE id=?", (record_id,)).fetchone()
            if not before:
                raise ValueError("El cambio de turno no existe.")
            if not bypass_window and not self.dentro_de_ventana(before["registrado_en"], "editar_horas"):
                raise ValueError("El registro superó el tiempo permitido para editar.")

            assignments = ", ".join(f"{field}=?" for field in fields)
            values = tuple(data[field] for field in fields) + (record_id,)
            conn.execute(f"UPDATE cambios_turno SET {assignments} WHERE id=?", values)
            after = dict(conn.execute("SELECT * FROM cambios_turno WHERE id=?", (record_id,)).fetchone())
            self._audit(conn, user_id, windows_user, "modificar", "cambio_turno", record_id, dict(before), after)
            self._enqueue_sync(conn, "cambios_turno", record_id, "UPDATE", after)

    def eliminar_cambio(
        self,
        record_id: int,
        user_id: Optional[int] = None,
        windows_user: Optional[str] = None,
        bypass_window: bool = False,
    ) -> None:
        with self.db.write_transaction() as conn:
            before = conn.execute("SELECT * FROM cambios_turno WHERE id=?", (record_id,)).fetchone()
            if not before:
                raise ValueError("La cambio de turno no existe.")
            if not bypass_window and not self.dentro_de_ventana(before["registrado_en"], "eliminar_horas"):
                raise ValueError("El registro superó el tiempo permitido para eliminar.")
            conn.execute("UPDATE cambios_turno SET activo=0 WHERE id=?", (record_id,))
            after = {"activo": 0}
            self._audit(conn, user_id, windows_user, "eliminar", "cambio_turno", record_id, dict(before), after)
            self._enqueue_sync(conn, "cambios_turno", record_id, "UPDATE", after)

    # ---------------------------------------------------------
    # Reemplazos y Personal de Ciclo
    # ---------------------------------------------------------

    def obtener_candidatos_ciclo(
        self,
        dotacion: str,
        especialidad: str,
        excluir_legajo: Optional[int] = None,
    ) -> List[Dict[str, Any]]:
        dot_norm = (dotacion or "").strip()
        esp_norm = (especialidad or "").strip()
        if not dot_norm or not esp_norm:
            return []
        with self.db.read_connection() as conn:
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
            rows = conn.execute(query, params).fetchall()
            return [dict(r) for r in rows]

    def listar_reemplazos_por_novedad(self, novedad_id: int) -> List[Dict[str, Any]]:
        with self.db.read_connection() as conn:
            rows = conn.execute(
                """SELECT id, novedad_id, legajo_reemplazo, apellidos_nombres_reemplazo,
                          turnos_reemplazo, franco_reemplazo, fecha_inicio, fecha_fin,
                          usuario_windows, usuario_id, creado_en
                   FROM novedad_reemplazos
                   WHERE novedad_id = ? AND activo = 1
                   ORDER BY id ASC""",
                (novedad_id,),
            ).fetchall()
            return [dict(r) for r in rows]

    def _rango_fechas_reemplazo(
        self,
        fecha_ini_str: str,
        fecha_fin_str: Optional[str],
        fecha_fin_nov_str: Optional[str] = None,
    ) -> Tuple[Optional[date], Optional[date]]:
        try:
            dt_ini = parsear_fecha(fecha_ini_str, "Fecha inicio reemplazo", obligatorio=False)
        except Exception:
            dt_ini = None
        if not dt_ini:
            return None, None
        f_ini = dt_ini.date()

        f_fin = None
        raw_fin = (fecha_fin_str or "").strip()
        if raw_fin and raw_fin != "-":
            try:
                dt_fin = parsear_fecha(raw_fin, "Fecha fin reemplazo", obligatorio=False)
                if dt_fin:
                    f_fin = dt_fin.date()
            except Exception:
                f_fin = None
        elif fecha_fin_nov_str and str(fecha_fin_nov_str).strip() not in ("", "-"):
            try:
                dt_fin_nov = parsear_fecha(fecha_fin_nov_str, "Fecha fin novedad", obligatorio=False)
                if dt_fin_nov:
                    f_fin = dt_fin_nov.date()
            except Exception:
                f_fin = None
        return f_ini, f_fin

    def _validar_disponibilidad_ciclo_conn(
        self,
        conn: sqlite3.Connection,
        legajo_reemplazo: int,
        fecha_inicio: str,
        fecha_fin: Optional[str],
        excluir_novedad_id: Optional[int] = None,
    ) -> None:
        dt_ini = parsear_fecha(fecha_inicio, "Fecha inicio reemplazo", obligatorio=True)
        dt_fin = parsear_fecha(fecha_fin, "Fecha fin reemplazo", obligatorio=False)
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

        existentes = conn.execute(query, params).fetchall()
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

    def validar_disponibilidad_ciclo(
        self,
        legajo_reemplazo: int,
        fecha_inicio: str,
        fecha_fin: Optional[str],
        excluir_novedad_id: Optional[int] = None,
    ) -> None:
        with self.db.read_connection() as conn:
            self._validar_disponibilidad_ciclo_conn(
                conn, legajo_reemplazo, fecha_inicio, fecha_fin, excluir_novedad_id=excluir_novedad_id
            )

    def guardar_reemplazos_novedad(
        self,
        novedad_id: int,
        tramos: List[Dict[str, Any]],
        user_id: Optional[int] = None,
        windows_user: Optional[str] = None,
    ) -> None:
        tramos_norm = []
        with self.db.write_transaction() as conn:
            nov = conn.execute("SELECT * FROM novedades WHERE id = ?", (novedad_id,)).fetchone()
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

                emp = conn.execute(
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
                    if not es_turno_ciclo(emp["turnos"]):
                        raise ValueError(f"El reemplazo {nombre} (Leg. {leg}) no tiene un turno que comience con CICLO.")
                elif not es_turno_ciclo(turno_r):
                    raise ValueError(f"El reemplazo {nombre} (Leg. {leg}) debe pertenecer a un turno que comience con CICLO.")

                tramos_norm.append({
                    "legajo_reemplazo": leg,
                    "apellidos_nombres_reemplazo": nombre,
                    "turnos_reemplazo": turno_r,
                    "franco_reemplazo": franco_r,
                    "fecha_inicio": f_ini,
                    "fecha_fin": f_fin or None,
                })

            es_valido, err = validar_tramos_reemplazo(
                nov["fecha_inicio"], fecha_fin_nov or None, nov["turnos"], tramos_norm
            )
            if not es_valido:
                raise ValueError(err)

            for t in tramos_norm:
                self._validar_disponibilidad_ciclo_conn(
                    conn, t["legajo_reemplazo"], t["fecha_inicio"], t["fecha_fin"], excluir_novedad_id=novedad_id
                )

            anteriores = [
                dict(r) for r in conn.execute(
                    "SELECT legajo_reemplazo, apellidos_nombres_reemplazo, fecha_inicio, fecha_fin FROM novedad_reemplazos WHERE novedad_id=? AND activo=1",
                    (novedad_id,),
                ).fetchall()
            ]
            conn.execute("UPDATE novedad_reemplazos SET activo = 0 WHERE novedad_id = ? AND activo = 1", (novedad_id,))
            ahora = self.db.now()
            for t in tramos_norm:
                conn.execute(
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
                conn, user_id, windows_user, "reemplazos", "novedad", novedad_id,
                {"reemplazos": anteriores}, {"reemplazos": tramos_norm},
            )
            self._enqueue_sync(conn, "novedad_reemplazos", novedad_id, "UPDATE", {"reemplazos": tramos_norm})

    def calcular_estado_cobertura(
        self,
        fila_novedad: Union[int, Dict[str, Any], sqlite3.Row],
        tramos_o_fecha_ini: Any = None,
        fecha_fin: Optional[str] = None,
        turnos: Optional[str] = None,
    ) -> Union[Dict[str, str], Tuple[str, str]]:
        """Calcula el estado de cobertura ('Sí', 'Parcial', 'No', 'N/A (Ciclo)') y textos descriptivos."""
        modo_tupla = isinstance(fila_novedad, int) or (turnos is not None or (fecha_fin is not None and not isinstance(tramos_o_fecha_ini, list)))
        if modo_tupla:
            novedad_id = int(fila_novedad)  # type: ignore
            tramos = self.listar_reemplazos_por_novedad(novedad_id)
            fila_dict = {
                "id": novedad_id,
                "fecha_inicio": tramos_o_fecha_ini,
                "fecha_fin": fecha_fin,
                "turnos": turnos or "",
            }
        else:
            fila_dict = dict(fila_novedad)  # type: ignore
            tramos = tramos_o_fecha_ini or []

        turno_titular = fila_dict.get("turnos") or ""
        if es_turno_ciclo(turno_titular):
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
            dt_ini_nov = parsear_fecha(f_ini_nov_str, "Inicio", obligatorio=False)
            f_ini_nov = dt_ini_nov.date() if dt_ini_nov else None
        except Exception:
            f_ini_nov = None

        try:
            dt_fin_nov = parsear_fecha(f_fin_nov_str, "Fin", obligatorio=False) if f_fin_nov_str else None
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

    def obtener_mapa_cobertura_novedades(self) -> Dict[int, Dict[str, str]]:
        """Devuelve un diccionario {novedad_id: info_cobertura} para todas las novedades activas."""
        with self.db.read_connection() as conn:
            novedades = conn.execute(
                "SELECT id, turnos, fecha_inicio, fecha_fin FROM novedades WHERE activo = 1"
            ).fetchall()
            reemplazos = conn.execute(
                """SELECT id, novedad_id, legajo_reemplazo, apellidos_nombres_reemplazo,
                          turnos_reemplazo, franco_reemplazo, fecha_inicio, fecha_fin
                   FROM novedad_reemplazos WHERE activo = 1 ORDER BY id ASC"""
            ).fetchall()

        por_novedad: Dict[int, List[Dict[str, Any]]] = {}
        for r in reemplazos:
            por_novedad.setdefault(r["novedad_id"], []).append(dict(r))

        mapa: Dict[int, Dict[str, str]] = {}
        for nov in novedades:
            res = self.calcular_estado_cobertura(nov, por_novedad.get(nov["id"], []))
            if isinstance(res, dict):
                mapa[nov["id"]] = res
        return mapa

    def obtener_turnos_sin_cubrir(
        self,
        fecha_consulta: str,
        dotacion: Optional[str] = None,
    ) -> List[Dict[str, Any]]:
        f_sol = parsear_fecha(fecha_consulta, "Fecha a consultar", obligatorio=True)
        with self.db.read_connection() as conn:
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
            filas = conn.execute(query, params).fetchall()

            reemplazos_rows = conn.execute(
                """SELECT id, novedad_id, legajo_reemplazo, apellidos_nombres_reemplazo,
                          turnos_reemplazo, franco_reemplazo, fecha_inicio, fecha_fin
                   FROM novedad_reemplazos WHERE activo = 1 ORDER BY id ASC"""
            ).fetchall()

        reemplazos_por_nov: Dict[int, List[Dict[str, Any]]] = {}
        for r in reemplazos_rows:
            reemplazos_por_nov.setdefault(r["novedad_id"], []).append(dict(r))

        descubiertos = []
        for fila in filas:
            try:
                f_ini = parsear_fecha(fila["fecha_inicio"], "Fecha de Inicio", obligatorio=False)
            except Exception:
                f_ini = None
            if not f_ini:
                continue

            fecha_fin_raw = (fila["fecha_fin"] or "").strip()
            f_fin = None
            if fecha_fin_raw and fecha_fin_raw != "-":
                try:
                    f_fin = parsear_fecha(fecha_fin_raw, "Fecha de Fin", obligatorio=False)
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

                    if es_turno_ciclo(fila["turnos"]):
                        item["cubierto"] = "N/A (Ciclo)"
                        item["reemplazo"] = ""
                    elif reemplazo_hoy:
                        item["cubierto"] = "Sí"
                        item["reemplazo"] = f"{reemplazo_hoy['legajo_reemplazo']} - {reemplazo_hoy['apellidos_nombres_reemplazo']} ({reemplazo_hoy['turnos_reemplazo']})"
                    else:
                        item["cubierto"] = "No"
                        item["reemplazo"] = ""

                    info_cob = self.calcular_estado_cobertura(fila, tramos)
                    if isinstance(info_cob, dict):
                        item["cobertura_resumen"] = info_cob["resumen"]
                        item["reemplazos_detalle"] = info_cob["detalle"]
                    descubiertos.append(item)

        return descubiertos

    def obtener_estado_personal_ciclo(
        self,
        fecha_consulta: str,
        dotacion: Optional[str] = None,
        especialidad: Optional[str] = None,
    ) -> Dict[str, List[Dict[str, Any]]]:
        f_sol = parsear_fecha(fecha_consulta, "Fecha de consulta", obligatorio=True).date()
        with self.db.read_connection() as conn:
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
            empleados_ciclo = [dict(r) for r in conn.execute(query_emp, params_emp).fetchall()]

            reemplazos_activos = conn.execute(
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

    def guardar_jornada_ciclo(
        self,
        legajo: int,
        fecha: str,
        toma: Optional[str] = None,
        deja: Optional[str] = None,
        franco_anticipado: int = 0,
        observaciones: Optional[str] = None,
        user_id: Optional[int] = None,
        windows_user: Optional[str] = None,
    ) -> None:
        fecha_obj = parsear_fecha(fecha, "Fecha", obligatorio=True).date()
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

        with self.db.write_transaction() as conn:
            before = conn.execute(
                "SELECT * FROM ciclo_jornadas WHERE legajo=? AND fecha=?",
                (int(legajo), fecha_iso),
            ).fetchone()
            conn.execute(
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
                    obs_str or None, self.db.now(), windows_user, user_id,
                ),
            )
            after = dict(conn.execute(
                "SELECT * FROM ciclo_jornadas WHERE legajo=? AND fecha=?",
                (int(legajo), fecha_iso),
            ).fetchone())
            self._audit(
                conn, user_id, windows_user,
                "modificar" if before else "crear",
                "ciclo_jornada",
                after["id"] if after else None,
                dict(before) if before else None,
                after,
            )
            self._enqueue_sync(conn, "ciclo_jornadas", after["id"], "UPSERT", after)

    def obtener_jornadas_ciclo_rango(
        self,
        legajo: int,
        fecha_desde: date,
        fecha_hasta: date,
    ) -> Dict[str, Dict[str, Any]]:
        with self.db.read_connection() as conn:
            rows = conn.execute(
                """SELECT id, legajo, fecha, toma, deja, franco_anticipado, observaciones
                   FROM ciclo_jornadas
                   WHERE legajo=? AND fecha BETWEEN ? AND ?""",
                (int(legajo), fecha_desde.strftime("%Y-%m-%d"), fecha_hasta.strftime("%Y-%m-%d")),
            ).fetchall()
            return {r["fecha"]: dict(r) for r in rows}

    def obtener_calendario_mensual_ciclo(
        self,
        legajo_ciclo: int,
        anio: int,
        mes: int,
    ) -> List[Dict[str, Any]]:
        anio = int(anio)
        mes = int(mes)
        _, num_dias = calendar.monthrange(anio, mes)

        with self.db.read_connection() as conn:
            emp = conn.execute(
                "SELECT legajo, apellidos_nombres, especialidad, dotacion, turnos, franco, franco_horario, jornada_horas FROM empleados WHERE legajo = ?",
                (int(legajo_ciclo),),
            ).fetchone()
            reemplazos = conn.execute(
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
        franco_norm = normalizar_dia_semana(franco_emp)

        fecha_primer_dia = date(anio, mes, 1)
        fecha_inicio_calc = fecha_primer_dia - timedelta(days=21)
        fecha_fin_calc = date(anio, mes, num_dias)
        jornadas_map = self.obtener_jornadas_ciclo_rango(legajo_ciclo, fecha_inicio_calc, fecha_fin_calc)

        def _parsear_hm(val_str: Optional[str]) -> Optional[Tuple[int, int]]:
            if not val_str:
                return None
            try:
                p = val_str.strip().split(":")
                return int(p[0]), int(p[1])
            except Exception:
                return None

        def _calc_hs_trabajadas(t_str: Optional[str], d_str: Optional[str]) -> float:
            hm_t = _parsear_hm(t_str)
            hm_d = _parsear_hm(d_str)
            if not hm_t or not hm_d:
                return 0.0
            min_t = hm_t[0] * 60 + hm_t[1]
            min_d = hm_d[0] * 60 + hm_d[1]
            if min_d < min_t:
                min_d += 24 * 60
            return round((min_d - min_t) / 60.0, 2)

        def _calc_deja_dt(dia_date: date, t_str: Optional[str], d_str: Optional[str]) -> Optional[datetime]:
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

        dias_historia: Dict[date, Dict[str, Any]] = {}
        acum_semana = 0.0

        d_step = fecha_inicio_calc
        while d_step <= fecha_fin_calc:
            f_iso = d_step.strftime("%Y-%m-%d")
            dia_semana_str = DIAS_SEMANA[d_step.weekday()]
            dia_sem_norm = normalizar_dia_semana(dia_semana_str)
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

            if es_franco:
                acum_semana = 0.0
                hs_semana_str = "-"
            else:
                acum_semana += hs_trab
                hs_semana_str = f"{acum_semana:.1f} hs" if acum_semana > 0 else "-"

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
                "dia_semana": DIAS_SEMANA[f_actual.weekday()],
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

    # ---------------------------------------------------------
    # Registros Eliminados y Recuperación
    # ---------------------------------------------------------

    def _tabla_entidad(self, tipo: str) -> Tuple[str, str]:
        if tipo.lower() in ("novedad", "novedades"):
            return "novedades", "novedad"
        if tipo.lower() in ("cambio", "cambios", "cambio_turno", "cambios_turno"):
            return "cambios_turno", "cambio_turno"
        raise ValueError("Tipo de registro inválido.")

    def recuperar_registro(
        self,
        tipo: str,
        record_id: int,
        user_id: Optional[int] = None,
        windows_user: Optional[str] = None,
    ) -> None:
        table, entidad = self._tabla_entidad(tipo)
        with self.db.write_transaction() as conn:
            before = conn.execute(f"SELECT * FROM {table} WHERE id=?", (record_id,)).fetchone()
            if not before:
                raise ValueError("El registro no existe.")
            conn.execute(f"UPDATE {table} SET activo=1 WHERE id=?", (record_id,))
            after = {"activo": 1}
            self._audit(conn, user_id, windows_user, "recuperar", entidad, record_id, dict(before), after)
            self._enqueue_sync(conn, table, record_id, "UPDATE", after)

    def borrar_definitivo(
        self,
        tipo: str,
        record_id: int,
        user_id: Optional[int] = None,
        windows_user: Optional[str] = None,
    ) -> None:
        table, entidad = self._tabla_entidad(tipo)
        with self.db.write_transaction() as conn:
            before = conn.execute(f"SELECT * FROM {table} WHERE id=?", (record_id,)).fetchone()
            if not before:
                raise ValueError("El registro no existe.")
            conn.execute(f"DELETE FROM {table} WHERE id=?", (record_id,))
            self._audit(conn, user_id, windows_user, "borrar_definitivo", entidad, record_id, dict(before), None)
            self._enqueue_sync(conn, table, record_id, "DELETE", None)

    def listar_eliminados(self) -> List[Dict[str, Any]]:
        with self.db.read_connection() as conn:
            rows = conn.execute(
                """SELECT 'Novedad' AS tipo, id, registrado_en, legajo, apellidos_nombres,
                          dotacion, novedad AS detalle, observaciones
                   FROM novedades WHERE activo=0
                   UNION ALL
                   SELECT 'Cambio' AS tipo, id, registrado_en, legajo_1, apellidos_nombres_1,
                          dotacion_1, fecha_cambio AS detalle, observaciones
                   FROM cambios_turno WHERE activo=0
                   ORDER BY registrado_en DESC"""
            ).fetchall()
            return [dict(r) for r in rows]

    # ---------------------------------------------------------
    # Auditoría
    # ---------------------------------------------------------

    def listar_auditoria(self, text_filter: str = "") -> List[Dict[str, Any]]:
        with self.db.read_connection() as conn:
            like = f"%{text_filter.strip()}%"
            rows = conn.execute(
                """SELECT a.id, a.creado_en, COALESCE(u.username, a.usuario_windows, '') AS usuario,
                          a.accion, a.entidad, a.entidad_id, a.datos_anteriores, a.datos_nuevos
                   FROM auditoria a LEFT JOIN usuarios u ON u.id=a.usuario_id
                   WHERE ?='' OR COALESCE(u.username, a.usuario_windows, '') LIKE ?
                         OR a.accion LIKE ? OR a.entidad LIKE ?
                   ORDER BY a.id DESC""",
                (text_filter.strip(), like, like, like),
            ).fetchall()
            return [dict(r) for r in rows]

    # ---------------------------------------------------------
    # Dashboard y Métricas
    # ---------------------------------------------------------

    def _es_dia_semana(self, valor: Any) -> bool:
        texto = normalizar_dia_semana(str(valor or ""))
        return texto in DIAS_SEMANA_NORM

    def obtener_metricas_dashboard(self, dias: int = 30) -> Dict[str, Any]:
        """Calcula todas las métricas del panel en una sola lectura optimizada."""
        dias = max(1, int(dias))
        hoy = datetime.now()
        desde = hoy - timedelta(days=dias - 1)
        fechas = [(desde + timedelta(days=i)).strftime("%Y-%m-%d") for i in range(dias)]
        serie = {fecha: {"novedades": 0, "cambios": 0} for fecha in fechas}

        conteo_tipo: Dict[str, int] = {}
        conteo_dotacion: Dict[str, int] = {}
        conteo_dia_semana: Dict[str, int] = {dia: 0 for dia in DIAS_SEMANA}

        resumen = {
            "novedades_hoy": 0, "cambios_hoy": 0,
            "novedades_mes": 0, "cambios_mes": 0, "total_mes": 0,
        }

        def sumar_dotacion(dotacion: Optional[str]) -> None:
            valor = (dotacion or "").strip()
            if not valor or self._es_dia_semana(valor):
                return
            conteo_dotacion[valor] = conteo_dotacion.get(valor, 0) + 1

        with self.db.read_connection() as conn:
            filas_nov = conn.execute(
                "SELECT registrado_en, novedad, dotacion FROM novedades WHERE activo=1"
            ).fetchall()
            filas_cam = conn.execute(
                "SELECT registrado_en, dotacion_1, dotacion_2 FROM cambios_turno WHERE activo=1"
            ).fetchall()

        for fila in filas_nov:
            fecha = self.parsear_registrado_en(fila["registrado_en"])
            if fecha is None:
                continue
            if fecha.date() == hoy.date():
                resumen["novedades_hoy"] += 1
            if fecha.year == hoy.year and fecha.month == hoy.month:
                resumen["novedades_mes"] += 1

            if fecha >= desde:
                conteo_dia_semana[DIAS_SEMANA[fecha.weekday()]] += 1
                tipo = (fila["novedad"] or "").strip() or "Sin tipo"
                conteo_tipo[tipo] = conteo_tipo.get(tipo, 0) + 1
                sumar_dotacion(fila["dotacion"])
                clave = fecha.strftime("%Y-%m-%d")
                if clave in serie:
                    serie[clave]["novedades"] += 1

        for fila in filas_cam:
            fecha = self.parsear_registrado_en(fila["registrado_en"])
            if fecha is None:
                continue
            if fecha.date() == hoy.date():
                resumen["cambios_hoy"] += 1
            if fecha.year == hoy.year and fecha.month == hoy.month:
                resumen["cambios_mes"] += 1

            if fecha >= desde:
                conteo_dia_semana[DIAS_SEMANA[fecha.weekday()]] += 1
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

    def dashboard_resumen(self) -> Dict[str, int]:
        return self.obtener_metricas_dashboard(30)["resumen"]

    def dashboard_por_tipo(self, dias: int = 30) -> List[Tuple[str, int]]:
        return self.obtener_metricas_dashboard(dias)["por_tipo"]

    def dashboard_por_dotacion(self, dias: int = 30, limite: int = 10) -> List[Tuple[str, int]]:
        return self.obtener_metricas_dashboard(dias)["por_dotacion"][:max(1, int(limite))]

    def dashboard_tendencia(self, dias: int = 30) -> List[Tuple[str, int, int]]:
        return self.obtener_metricas_dashboard(dias)["tendencia"]

    def dashboard_por_dia_semana(self, dias: int = 30) -> List[Tuple[str, int]]:
        return self.obtener_metricas_dashboard(dias)["por_dia_semana"]
