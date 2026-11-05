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

from src.core.config import obtener_usuario_windows
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
    validar_datos_ordenamiento,
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
        win_usr = (windows_user or "").strip() or obtener_usuario_windows()
        conn.execute(
            """INSERT INTO auditoria
               (usuario_id, usuario_windows, accion, entidad, entidad_id,
                datos_anteriores, datos_nuevos, creado_en)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                user_id,
                win_usr,
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

    def listar_especialidades_empleados(self) -> List[str]:
        """Devuelve el listado ordenado de especialidades únicas activas existentes en la base de datos."""
        with self.db.read_connection() as conn:
            rows = conn.execute(
                """SELECT DISTINCT especialidad FROM empleados
                   WHERE activo = 1 AND especialidad IS NOT NULL AND TRIM(especialidad) != ''
                   ORDER BY especialidad ASC"""
            ).fetchall()
            return [r["especialidad"] for r in rows]

    def listar_empleados_ciclo(self) -> List[Dict[str, Any]]:
        """Devuelve el personal activo correspondiente a régimen de ciclo mensual."""
        with self.db.read_connection() as conn:
            rows = conn.execute(
                """SELECT id, legajo, apellidos_nombres, especialidad, dotacion, turnos,
                          franco, franco_horario, jornada_horas, activo
                   FROM empleados
                   WHERE activo = 1 AND UPPER(TRIM(turnos)) LIKE 'CICLO%'
                   ORDER BY dotacion, apellidos_nombres"""
            ).fetchall()
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
            self._enqueue_sync(conn, "destinatarios_informe", item_id, "INSERT", payload)
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
            self._enqueue_sync(conn, "destinatarios_informe", item_id, "UPDATE", after)

    # ---------------------------------------------------------
    # Destinatarios de Ordenamiento CRUD
    # ---------------------------------------------------------

    def listar_destinatarios_ordenamiento(self, incluir_inactivos: bool = True) -> List[Dict[str, Any]]:
        with self.db.read_connection() as conn:
            query = "SELECT id, nombre, email, activo FROM destinatarios_ordenamiento"
            if not incluir_inactivos:
                query += " WHERE activo=1"
            rows = conn.execute(query + " ORDER BY nombre").fetchall()
            return [dict(r) for r in rows]

    def crear_destinatario_ordenamiento(
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
                "INSERT INTO destinatarios_ordenamiento(nombre, email) VALUES (?, ?)",
                (nombre, email),
            )
            item_id = cur.lastrowid
            payload = {"nombre": nombre, "email": email}
            self._audit(conn, user_id, windows_user, "crear", "destinatario_ordenamiento", item_id, None, payload)
            self._enqueue_sync(conn, "destinatarios_ordenamiento", item_id, "INSERT", payload)
            return item_id

    def actualizar_destinatario_ordenamiento(
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
            before = conn.execute("SELECT id, nombre, email, activo FROM destinatarios_ordenamiento WHERE id=?", (item_id,)).fetchone()
            if not before:
                raise ValueError("El destinatario no existe.")
            conn.execute(
                "UPDATE destinatarios_ordenamiento SET nombre=?, email=?, activo=? WHERE id=?",
                (nombre, email, int(activo), item_id),
            )
            after = {"nombre": nombre, "email": email, "activo": int(activo)}
            self._audit(conn, user_id, windows_user, "modificar", "destinatario_ordenamiento", item_id, dict(before), after)
            self._enqueue_sync(conn, "destinatarios_ordenamiento", item_id, "UPDATE", after)

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

    def validar_novedad_superpuesta(
        self,
        legajo: int,
        fecha_inicio_str: str,
        fecha_fin_str: Optional[str] = None,
        excluir_novedad_id: Optional[int] = None,
    ) -> None:
        """Verifica que el agente no posea una novedad activa que se solape en el período."""
        f_ini = parsear_fecha(fecha_inicio_str, "Fecha de inicio")
        if not f_ini:
            return
        f_fin = parsear_fecha(fecha_fin_str, "Fecha de fin", obligatorio=False) if (fecha_fin_str and fecha_fin_str != "-") else None

        with self.db.read_connection() as conn:
            query = "SELECT id, novedad, fecha_inicio, fecha_fin FROM novedades WHERE activo = 1 AND legajo = ?"
            params: List[Any] = [int(legajo)]
            if excluir_novedad_id:
                query += " AND id != ?"
                params.append(int(excluir_novedad_id))
            rows = conn.execute(query, params).fetchall()

        for r in rows:
            r_ini = parsear_fecha(r["fecha_inicio"])
            if not r_ini:
                continue
            r_fin_raw = r["fecha_fin"]
            r_fin = parsear_fecha(r_fin_raw) if (r_fin_raw and r_fin_raw != "-") else None

            # Solapamiento si: f_ini <= (r_fin or +inf) and (f_fin or +inf) >= r_ini
            if (r_fin is None or f_ini.date() <= r_fin.date()) and (f_fin is None or f_fin.date() >= r_ini.date()):
                fin_display = r_fin_raw if (r_fin_raw and r_fin_raw != "-") else "Abierta"
                raise ValueError(
                    f"El agente con legajo {legajo} ya posee una novedad activa (#{r['id']} - {r['novedad']}) "
                    f"que abarca el período {r['fecha_inicio']} a {fin_display}."
                )

    def validar_reemplazo_disponible(
        self,
        legajo_reemplazo: int,
        fecha_inicio_str: str,
        fecha_fin_str: Optional[str] = None,
        excluir_novedad_id: Optional[int] = None,
        excluir_reemplazo_id: Optional[int] = None,
    ) -> None:
        """Verifica que el personal de reemplazo no tenga novedad activa ni esté relevando otra novedad en esa fecha."""
        f_ini = parsear_fecha(fecha_inicio_str, "Fecha de inicio relevo")
        if not f_ini:
            return
        f_fin = parsear_fecha(fecha_fin_str, "Fecha de fin relevo", obligatorio=False) if (fecha_fin_str and fecha_fin_str != "-") else None

        with self.db.read_connection() as conn:
            # 1. Comprobar si el reemplazo tiene novedad activa propia
            novs = conn.execute(
                "SELECT id, novedad, fecha_inicio, fecha_fin FROM novedades WHERE activo = 1 AND legajo = ?",
                (int(legajo_reemplazo),),
            ).fetchall()
            for n in novs:
                n_ini = parsear_fecha(n["fecha_inicio"])
                if not n_ini:
                    continue
                n_fin = parsear_fecha(n["fecha_fin"]) if (n["fecha_fin"] and n["fecha_fin"] != "-") else None
                if (n_fin is None or f_ini.date() <= n_fin.date()) and (f_fin is None or f_fin.date() >= n_ini.date()):
                    fin_disp = n["fecha_fin"] if (n["fecha_fin"] and n["fecha_fin"] != "-") else "Abierta"
                    raise ValueError(
                        f"El agente de relevo legajo {legajo_reemplazo} no está disponible: "
                        f"posee una novedad activa propia (#{n['id']} - {n['novedad']}) del {n['fecha_inicio']} al {fin_disp}."
                    )

            # 2. Comprobar si ya está relevando otra novedad activa en el mismo tramo
            query_r = """SELECT nr.id, nr.novedad_id, nr.fecha_inicio, nr.fecha_fin, n.apellidos_nombres, n.turnos
                         FROM novedad_reemplazos nr
                         JOIN novedades n ON n.id = nr.novedad_id
                         WHERE nr.activo = 1 AND n.activo = 1 AND nr.legajo_reemplazo = ?"""
            params_r: List[Any] = [int(legajo_reemplazo)]
            if excluir_reemplazo_id:
                query_r += " AND nr.id != ?"
                params_r.append(int(excluir_reemplazo_id))
            if excluir_novedad_id:
                query_r += " AND nr.novedad_id != ?"
                params_r.append(int(excluir_novedad_id))

            relevos = conn.execute(query_r, params_r).fetchall()
            for r in relevos:
                r_ini = parsear_fecha(r["fecha_inicio"])
                if not r_ini:
                    continue
                r_fin = parsear_fecha(r["fecha_fin"]) if (r["fecha_fin"] and r["fecha_fin"] != "-") else None
                if (r_fin is None or f_ini.date() <= r_fin.date()) and (f_fin is None or f_fin.date() >= r_ini.date()):
                    fin_r = r["fecha_fin"] if (r["fecha_fin"] and r["fecha_fin"] != "-") else "Sin fecha fin"
                    raise ValueError(
                        f"El agente legajo {legajo_reemplazo} ya se encuentra cubriendo la novedad #{r['novedad_id']} "
                        f"(Titular: {r['apellidos_nombres']} - Turno: {r['turnos']}) en el período {r['fecha_inicio']} al {fin_r}."
                    )

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
        leg_val = int(str(data["legajo"]).strip())
        self.validar_novedad_superpuesta(leg_val, data.get("fecha_inicio"), data.get("fecha_fin"))

        win_usr = (windows_user or data.get("usuario_windows") or "").strip() or obtener_usuario_windows()
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
                win_usr,
                user_id,
                1,
            )
            cur = conn.execute(f"INSERT INTO novedades ({', '.join(columns)}) VALUES ({placeholders})", values)
            record_id = cur.lastrowid
            after = dict(conn.execute("SELECT * FROM novedades WHERE id=?", (record_id,)).fetchone())
            self._audit(conn, user_id, win_usr, "crear", "novedad", record_id, None, after)
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
        leg_val = int(str(data["legajo"]).strip())
        self.validar_novedad_superpuesta(
            leg_val,
            data.get("fecha_inicio"),
            data.get("fecha_fin"),
            excluir_novedad_id=record_id,
        )
        win_usr = (windows_user or data.get("usuario_windows") or "").strip() or obtener_usuario_windows()
        with self.db.write_transaction() as conn:
            before = conn.execute("SELECT * FROM novedades WHERE id=?", (record_id,)).fetchone()
            if not before:
                raise ValueError("La novedad no existe.")
            if not bypass_window and not self.dentro_de_ventana(before["registrado_en"], "editar_horas"):
                raise ValueError("El registro superó el tiempo permitido para editar.")

            assignments = ", ".join(f"{field}=?" for field in fields) + ", usuario_windows=?"
            values = tuple(data[field] for field in fields) + (win_usr, record_id)
            conn.execute(f"UPDATE novedades SET {assignments} WHERE id=?", values)
            after = dict(conn.execute("SELECT * FROM novedades WHERE id=?", (record_id,)).fetchone())
            self._audit(conn, user_id, win_usr, "modificar", "novedad", record_id, dict(before), after)
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
        win_usr = (windows_user or data.get("usuario_windows") or "").strip() or obtener_usuario_windows()
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
                win_usr,
                user_id,
                1,
            )
            cur = conn.execute(f"INSERT INTO cambios_turno ({', '.join(columns)}) VALUES ({placeholders})", values)
            record_id = cur.lastrowid
            after = dict(conn.execute("SELECT * FROM cambios_turno WHERE id=?", (record_id,)).fetchone())
            self._audit(conn, user_id, win_usr, "crear", "cambio_turno", record_id, None, after)
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
        win_usr = (windows_user or data.get("usuario_windows") or "").strip() or obtener_usuario_windows()
        with self.db.write_transaction() as conn:
            before = conn.execute("SELECT * FROM cambios_turno WHERE id=?", (record_id,)).fetchone()
            if not before:
                raise ValueError("El cambio de turno no existe.")
            if not bypass_window and not self.dentro_de_ventana(before["registrado_en"], "editar_horas"):
                raise ValueError("El registro superó el tiempo permitido para editar.")

            assignments = ", ".join(f"{field}=?" for field in fields) + ", usuario_windows=?"
            values = tuple(data[field] for field in fields) + (win_usr, record_id)
            conn.execute(f"UPDATE cambios_turno SET {assignments} WHERE id=?", values)
            after = dict(conn.execute("SELECT * FROM cambios_turno WHERE id=?", (record_id,)).fetchone())
            self._audit(conn, user_id, win_usr, "modificar", "cambio_turno", record_id, dict(before), after)
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
            win_usr = (windows_user or "").strip() or obtener_usuario_windows()
            for t in tramos_norm:
                conn.execute(
                    """INSERT INTO novedad_reemplazos
                       (novedad_id, legajo_reemplazo, apellidos_nombres_reemplazo, turnos_reemplazo,
                        franco_reemplazo, fecha_inicio, fecha_fin, usuario_windows, usuario_id, creado_en, activo)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1)""",
                    (
                        novedad_id, t["legajo_reemplazo"], t["apellidos_nombres_reemplazo"],
                        t["turnos_reemplazo"], t["franco_reemplazo"], t["fecha_inicio"],
                        t["fecha_fin"], win_usr, user_id, ahora,
                    ),
                )
            self._audit(
                conn, user_id, win_usr, "reemplazos", "novedad", novedad_id,
                {"reemplazos": anteriores}, {"reemplazos": tramos_norm},
            )
            self._enqueue_sync(conn, "novedad_reemplazos", novedad_id, "UPDATE", {"reemplazos": tramos_norm})

    def desactivar_reemplazo(
        self,
        reemplazo_id: int,
        user_id: Optional[int] = None,
        windows_user: Optional[str] = None,
    ) -> None:
        """Desactiva un tramo de relevo individual en novedad_reemplazos."""
        with self.db.write_transaction() as conn:
            before = conn.execute("SELECT * FROM novedad_reemplazos WHERE id = ?", (int(reemplazo_id),)).fetchone()
            if not before:
                return
            conn.execute("UPDATE novedad_reemplazos SET activo = 0 WHERE id = ?", (int(reemplazo_id),))
            self._audit(conn, user_id, windows_user, "eliminar", "novedad_reemplazo", int(reemplazo_id), dict(before), {"activo": 0})
            self._enqueue_sync(conn, "novedad_reemplazos", int(reemplazo_id), "UPDATE", {"activo": 0})

    def asignar_reemplazo_rapido(
        self,
        novedad_id: int,
        legajo_reemplazo: int,
        fecha_inicio: str,
        fecha_fin: Optional[str] = None,
        user_id: Optional[int] = None,
        windows_user: Optional[str] = None,
    ) -> int:
        """Asigna un agente de ciclo como relevo de una novedad para una fecha o rango."""
        self.validar_reemplazo_disponible(
            int(legajo_reemplazo),
            fecha_inicio,
            fecha_fin,
            excluir_novedad_id=novedad_id,
        )
        with self.db.write_transaction() as conn:
            emp = conn.execute(
                "SELECT legajo, apellidos_nombres, especialidad, dotacion, turnos, franco FROM empleados WHERE legajo = ?",
                (int(legajo_reemplazo),),
            ).fetchone()
            if not emp:
                raise ValueError(f"No existe el empleado con legajo {legajo_reemplazo}")

            f_fin_val = fecha_fin if fecha_fin and str(fecha_fin).strip() not in ("", "-", "None") else None
            ahora = self.db.now()
            cur = conn.execute(
                """INSERT INTO novedad_reemplazos
                   (novedad_id, legajo_reemplazo, apellidos_nombres_reemplazo, turnos_reemplazo,
                    franco_reemplazo, fecha_inicio, fecha_fin, usuario_windows, usuario_id, creado_en, activo)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1)""",
                (
                    int(novedad_id), int(emp["legajo"]), emp["apellidos_nombres"],
                    emp["turnos"], emp["franco"], str(fecha_inicio).strip(),
                    f_fin_val, windows_user, user_id, ahora,
                ),
            )
            r_id = cur.lastrowid
            after = dict(conn.execute("SELECT * FROM novedad_reemplazos WHERE id=?", (r_id,)).fetchone())
            self._audit(conn, user_id, windows_user, "crear", "novedad_reemplazo", r_id, None, after)
            self._enqueue_sync(conn, "novedad_reemplazos", r_id, "INSERT", after)
            return r_id

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

        turnos_map = self.listar_turnos_itinerario_map()
        from src.domain.validators import resolver_jornada_itinerario

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

                    turno_base_titular = item.get("turnos") or ""
                    franco_titular = item.get("franco") or "Domingo"
                    fecha_iso_sol = f_sol.strftime("%Y-%m-%d")

                    turno_reemplazado = "-"
                    turno_a_resolver = turno_base_titular
                    franco_a_resolver = franco_titular

                    if es_turno_ciclo(turno_base_titular):
                        item["cubierto"] = "N/A (Ciclo)"
                        item["reemplazo"] = ""
                        # Verificar si este personal de ciclo estaba reemplazando una novedad en esta fecha
                        tit_cubierto = None
                        for r in reemplazos_rows:
                            if int(r["legajo_reemplazo"]) == int(fila["legajo"]):
                                r_ini_c, r_fin_c = self._rango_fechas_reemplazo(r["fecha_inicio"], r["fecha_fin"], None)
                                if r_ini_c and r_ini_c <= f_sol.date() and (r_fin_c is None or f_sol.date() <= r_fin_c):
                                    tit_cubierto = r
                                    break
                        if tit_cubierto:
                            with self.db.read_connection() as c_nov:
                                row_tit = c_nov.execute("SELECT legajo, apellidos_nombres, turnos, franco FROM novedades WHERE id=?", (tit_cubierto["novedad_id"],)).fetchone()
                                if row_tit:
                                    turno_reemplazado = f"{row_tit['turnos']} (Titular: {row_tit['apellidos_nombres']})"
                                    turno_a_resolver = row_tit["turnos"]
                                    franco_a_resolver = row_tit["franco"] or "Domingo"
                                else:
                                    turno_reemplazado = f"Novedad #{tit_cubierto['novedad_id']}"
                        else:
                            turno_reemplazado = "Sin asignación"
                    elif reemplazo_hoy:
                        item["cubierto"] = "Sí"
                        item["reemplazo"] = f"{reemplazo_hoy['legajo_reemplazo']} - {reemplazo_hoy['apellidos_nombres_reemplazo']} ({reemplazo_hoy['turnos_reemplazo']})"
                        turno_reemplazado = f"Cubierto por Leg. {reemplazo_hoy['legajo_reemplazo']} ({reemplazo_hoy['turnos_reemplazo']})"
                    else:
                        item["cubierto"] = "No"
                        item["reemplazo"] = ""
                        turno_reemplazado = "-"

                    item["turno_reemplazado"] = turno_reemplazado

                    # Resolver turno del día, hora toma y hora deja
                    try:
                        res_jor = resolver_jornada_itinerario(
                            turno_base=turno_a_resolver,
                            fecha_target=fecha_iso_sol,
                            franco_agente=franco_a_resolver,
                            turnos_dict=turnos_map,
                        )
                        if res_jor.get("es_franco"):
                            item["turno_dia"] = "FRANCO"
                            item["hora_toma"] = "-"
                            item["hora_deja"] = "-"
                        else:
                            item["turno_dia"] = res_jor.get("codigo_turno") or turno_a_resolver
                            item["hora_toma"] = res_jor.get("toma") or "-"
                            item["hora_deja"] = res_jor.get("deje") or "-"
                    except Exception:
                        item["turno_dia"] = turno_a_resolver
                        item["hora_toma"] = "-"
                        item["hora_deja"] = "-"

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
                query_emp += " AND LOWER(TRIM(especialidad)) LIKE '%' || LOWER(TRIM(?)) || '%'"
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
                item["reemplazo_id"] = asignacion["id"]
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
                          n.turnos AS turno_titular, n.franco AS franco_titular, n.novedad AS tipo_novedad,
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
        turnos_itinerario_map = self.listar_turnos_itinerario_map()

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

                # Si no hay horario manual ingresado, resolver automáticamente desde turnos_itinerario
                if not toma and not deja and turno_cubierto and turno_cubierto != "-" and turnos_itinerario_map:
                    from src.domain.validators import resolver_jornada_itinerario
                    franco_titular = relevo_dia.get("franco_titular") or "Domingo"
                    res_jornada = resolver_jornada_itinerario(
                        turno_base=turno_cubierto,
                        fecha_target=f_iso,
                        franco_agente=franco_titular,
                        turnos_dict=turnos_itinerario_map,
                    )
                    if res_jornada.get("es_franco"):
                        es_franco = True
                        estado = "Franco por Relevo"
                        toma = ""
                        deja = ""
                    else:
                        toma = res_jornada.get("toma") or ""
                        deja = res_jornada.get("deje") or ""
                        cod_variante = res_jornada.get("codigo_turno") or turno_cubierto
                        reemplaza_a = f"{relevo_dia['legajo_titular']} - {relevo_dia['apellidos_titular']} (Turno {cod_variante})"
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
        if tipo.lower() in ("ordenamiento", "ordenamientos"):
            return "ordenamientos", "ordenamiento"
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
            if table == "ordenamientos":
                conn.execute(
                    "UPDATE ordenamientos SET activo=1, estado='activo', deleted_at=NULL, updated_at=datetime('now', 'localtime') WHERE id=?",
                    (record_id,),
                )
                after = {"activo": 1, "estado": "activo", "deleted_at": None}
            else:
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
            if table == "ordenamientos":
                conn.execute("DELETE FROM ordenamiento_tramos WHERE ordenamiento_id=?", (record_id,))
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
                   SELECT 'Cambio' AS tipo, id, registrado_en, legajo_1 AS legajo, apellidos_nombres_1 AS apellidos_nombres,
                          dotacion_1 AS dotacion, fecha_cambio AS detalle, observaciones
                   FROM cambios_turno WHERE activo=0
                   UNION ALL
                   SELECT 'Ordenamiento' AS tipo, id, COALESCE(deleted_at, created_at, fecha_emision) AS registrado_en,
                          legajo_personal AS legajo, nombre_personal AS apellidos_nombres,
                          deposito AS dotacion, (nro_orden || ' - ' || motivo_orden) AS detalle, observacion_pie AS observaciones
                   FROM ordenamientos WHERE activo=0 OR estado='eliminado'
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

    # ---------------------------------------------------------
    # Itinerarios de Conducción Eléctrica y Buscador por Tren
    # ---------------------------------------------------------

    def importar_itinerario_excel(self, file_path: str, itinerario_nombre: Optional[str] = None) -> Tuple[int, int]:
        """Importa el archivo normalizado de itinerarios (ej. Indice_45.xlsx) a SQLite."""
        import openpyxl
        import re

        wb = openpyxl.load_workbook(file_path, data_only=True)
        ws = wb.active if "Datos" not in wb.sheetnames else wb["Datos"]

        filas_por_turno: Dict[str, List[Dict[str, Any]]] = {}

        for row in range(2, ws.max_row + 1):
            t_raw = ws.cell(row, 1).value
            if not t_raw:
                continue
            codigo_turno = str(t_raw).strip().replace(",", ".")

            toma = str(ws.cell(row, 2).value or "").strip()
            deja = str(ws.cell(row, 3).value or "").strip()
            nro_vuelta_raw = ws.cell(row, 4).value
            try:
                nro_vuelta = int(nro_vuelta_raw) if nro_vuelta_raw is not None else len(filas_por_turno.get(codigo_turno, [])) + 1
            except (ValueError, TypeError):
                nro_vuelta = len(filas_por_turno.get(codigo_turno, [])) + 1

            ref = str(ws.cell(row, 5).value or "").strip()
            nro_tren = str(ws.cell(row, 6).value or "").strip()
            if nro_tren in ("None", "-", ""):
                nro_tren = None

            origen = str(ws.cell(row, 7).value or "").strip()
            h_salida = str(ws.cell(row, 8).value or "").strip()
            destino = str(ws.cell(row, 9).value or "").strip()
            h_llegada = str(ws.cell(row, 10).value or "").strip()
            obs_tramo = str(ws.cell(row, 11).value or "").strip()
            circ_iti = str(ws.cell(row, 12).value or "").strip()
            hoja = str(ws.cell(row, 13).value or "").strip()

            item = {
                "codigo_turno": codigo_turno,
                "toma": toma,
                "deja": deja,
                "nro_vuelta": nro_vuelta,
                "ref": ref,
                "nro_tren": nro_tren,
                "origen": origen,
                "h_salida": h_salida,
                "destino": destino,
                "h_llegada": h_llegada,
                "obs_tramo": obs_tramo,
                "circ_iti": circ_iti,
                "hoja": hoja,
            }
            filas_por_turno.setdefault(codigo_turno, []).append(item)

        turnos_guardados = 0
        vueltas_guardadas = 0

        with self.db.write_transaction() as conn:
            for cod, vueltas in filas_por_turno.items():
                v0 = vueltas[0]
                toma = v0["toma"]
                deja = v0["deja"]
                itinerario = itinerario_nombre or v0["circ_iti"] or "Dic25"

                # Deducir depósito
                hoja_upper = v0["hoja"].upper()
                deposito = "PC"
                if "LLV" in hoja_upper:
                    deposito = "LLV"
                elif "TEMPERLEY" in hoja_upper or "TY" in hoja_upper:
                    deposito = "TY"
                elif "KM5" in hoja_upper:
                    deposito = "KM5"
                elif "LP" in hoja_upper or "LA PLATA" in hoja_upper:
                    deposito = "LP"
                elif "TOLOSA" in hoja_upper or "OA" in hoja_upper:
                    deposito = "OA"

                # Clasificar tipo de turno y parsear base y subíndice
                m = re.match(r"^(\d+)\.(\d+)$", cod)
                if m:
                    base = m.group(1)
                    subindice = int(m.group(2))
                    tipo_turno = "DIAGRAMA_6D"
                elif cod.startswith("PROG") or "PROG" in cod:
                    base = cod
                    subindice = None
                    tipo_turno = "PROGRAMA"
                elif "ORD" in cod or "DH" in cod:
                    base = cod
                    subindice = None
                    tipo_turno = "FRANQUERO"
                else:
                    base = cod
                    subindice = None
                    tipo_turno = "UNIPERSONAL"

                # Calcular duración aproximada si toma y deja están en formato HH:MM
                duracion = ""
                try:
                    if len(toma) == 5 and len(deja) == 5 and ":" in toma and ":" in deja:
                        t_h, t_m = map(int, toma.split(":"))
                        d_h, d_m = map(int, deja.split(":"))
                        t_min = t_h * 60 + t_m
                        d_min = d_h * 60 + d_m
                        diff = (d_min - t_min) % (24 * 60)
                        duracion = f"{diff // 60:02d}:{diff % 60:02d}"
                except Exception:
                    duracion = ""

                cur = conn.execute(
                    """INSERT INTO turnos_itinerario
                       (codigo_turno, turno_base, subindice, tipo_turno, deposito_base,
                        toma_servicio, deje_servicio, duracion_jornada, itinerario_vigente,
                        observaciones_turno, activo, sync_status, created_at, updated_at)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, '', 1, 'pending', datetime('now', 'localtime'), datetime('now', 'localtime'))
                       ON CONFLICT(codigo_turno) DO UPDATE SET
                        turno_base = excluded.turno_base,
                        subindice = excluded.subindice,
                        tipo_turno = excluded.tipo_turno,
                        deposito_base = excluded.deposito_base,
                        toma_servicio = excluded.toma_servicio,
                        deje_servicio = excluded.deje_servicio,
                        duracion_jornada = excluded.duracion_jornada,
                        itinerario_vigente = excluded.itinerario_vigente,
                        updated_at = datetime('now', 'localtime');""",
                    (cod, base, subindice, tipo_turno, deposito, toma, deja, duracion, itinerario),
                )

                turno_row = conn.execute("SELECT id FROM turnos_itinerario WHERE codigo_turno = ?;", (cod,)).fetchone()
                turno_id = turno_row[0]
                turnos_guardados += 1

                # Reemplazar vueltas
                conn.execute("DELETE FROM turnos_vueltas WHERE turno_id = ?;", (turno_id,))
                for v in vueltas:
                    conn.execute(
                        """INSERT INTO turnos_vueltas
                           (turno_id, nro_vuelta, ref, nro_tren, origen, hora_salida,
                            destino, hora_llegada, observaciones_tramo, sync_status, created_at, updated_at)
                           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'pending', datetime('now', 'localtime'), datetime('now', 'localtime'));""",
                        (
                            turno_id,
                            v["nro_vuelta"],
                            v["ref"],
                            v["nro_tren"],
                            v["origen"],
                            v["h_salida"],
                            v["destino"],
                            v["h_llegada"],
                            v["obs_tramo"],
                        ),
                    )
                    vueltas_guardadas += 1

        logger.info("Importación de itinerario completada: %d turnos, %d vueltas.", turnos_guardados, vueltas_guardadas)
        return turnos_guardados, vueltas_guardadas

    def buscar_turnos_por_tren(self, nro_tren: str) -> List[Dict[str, Any]]:
        """Busca de forma instantánea todos los turnos que intervienen en un número de tren."""
        tren_clean = str(nro_tren or "").strip()
        if not tren_clean:
            return []

        with self.db.read_connection() as conn:
            rows = conn.execute(
                """SELECT
                    v.id AS vuelta_id,
                    v.nro_tren,
                    ti.codigo_turno,
                    ti.turno_base,
                    ti.tipo_turno,
                    ti.deposito_base,
                    ti.toma_servicio,
                    ti.deje_servicio,
                    v.nro_vuelta,
                    v.ref,
                    v.origen,
                    v.hora_salida,
                    v.destino,
                    v.hora_llegada,
                    v.observaciones_tramo AS obs_vuelta,
                    ti.observaciones_turno AS obs_turno,
                    ti.itinerario_vigente
                   FROM turnos_vueltas v
                   JOIN turnos_itinerario ti ON v.turno_id = ti.id
                   WHERE v.nro_tren LIKE ? OR v.nro_tren = ?
                   ORDER BY v.hora_salida ASC, ti.codigo_turno ASC;""",
                (f"%{tren_clean}%", tren_clean),
            ).fetchall()
            return [dict(r) for r in rows]

    def listar_turnos_itinerario_map(self) -> Dict[str, Dict[str, Any]]:
        """Devuelve un mapa indexado por codigo_turno para resolución ultra-rápida."""
        with self.db.read_connection() as conn:
            rows = conn.execute(
                """SELECT id, codigo_turno, turno_base, subindice, tipo_turno,
                          deposito_base, toma_servicio, deje_servicio, duracion_jornada,
                          itinerario_vigente, observaciones_turno
                   FROM turnos_itinerario WHERE activo = 1;"""
            ).fetchall()
            return {r["codigo_turno"]: dict(r) for r in rows}

    def listar_itinerarios_disponibles(self) -> List[str]:
        """Devuelve la lista de códigos/nombres de itinerarios únicos cargados."""
        with self.db.read_connection() as conn:
            rows = conn.execute(
                """SELECT DISTINCT itinerario_vigente FROM turnos_itinerario
                   WHERE activo = 1 ORDER BY itinerario_vigente ASC;"""
            ).fetchall()
            return [r[0] for r in rows if r[0]]

    def obtener_vueltas_de_turno(self, turno_id: int) -> List[Dict[str, Any]]:
        """Devuelve la totalidad de vueltas de un turno ordenadas cronológicamente."""
        with self.db.read_connection() as conn:
            rows = conn.execute(
                """SELECT id, turno_id, nro_vuelta, ref, nro_tren, origen,
                          hora_salida, destino, hora_llegada, observaciones_tramo
                   FROM turnos_vueltas
                   WHERE turno_id = ?
                   ORDER BY nro_vuelta ASC, hora_salida ASC;""",
                (int(turno_id),),
            ).fetchall()
            return [dict(r) for r in rows]

    def obtener_turno_por_codigo(
        self,
        codigo_turno: str,
        itinerario: Optional[str] = None,
    ) -> Optional[Dict[str, Any]]:
        """Busca un turno por su código (ej: '401.1' o '401') y opcionalmente por itinerario."""
        cod = str(codigo_turno or "").strip()
        if not cod:
            return None
        with self.db.read_connection() as conn:
            query = """
                SELECT id, codigo_turno, turno_base, subindice, tipo_turno, deposito_base,
                       toma_servicio, deje_servicio, duracion_jornada, tiempo_toma, tiempo_deje,
                       itinerario_vigente, observaciones_turno, activo
                FROM turnos_itinerario
                WHERE activo = 1 AND (codigo_turno = ? OR turno_base = ?)
            """
            params = [cod, cod]
            if itinerario and str(itinerario).strip() not in ("Todos", ""):
                query += " AND itinerario_vigente = ?"
                params.append(str(itinerario).strip())
            query += " ORDER BY itinerario_vigente DESC, id DESC LIMIT 1;"
            row = conn.execute(query, params).fetchone()
            if not row:
                return None
            res = dict(row)
            res["vueltas"] = self.obtener_vueltas_de_turno(res["id"])
            return res

    def buscar_turnos_con_cobertura(
        self,
        criterio: str,
        fecha_target_iso: str,
        tipo_dia: str = "HABIL",
        itinerarios: Optional[List[str]] = None,
    ) -> List[Dict[str, Any]]:
        """Búsqueda dual (por tren o por turno) con resolución de conductor titular y cobertura de novedades."""
        crit_clean = str(criterio or "").strip()
        if not crit_clean:
            return []

        f_obj = parsear_fecha(fecha_target_iso)
        if not f_obj:
            f_obj = date.today()
        dia_sem_nom = DIAS_SEMANA[f_obj.weekday()]

        with self.db.read_connection() as conn:
            # 1. Identificar turnos candidatos según criterio dual
            where_clauses = ["ti.activo = 1"]
            params: List[Any] = []

            if itinerarios:
                placeholders = ",".join("?" for _ in itinerarios)
                where_clauses.append(f"ti.itinerario_vigente IN ({placeholders})")
                params.extend(itinerarios)

            if crit_clean:
                # Criterio dual: coincide con turno_base, codigo_turno o nro_tren en vueltas
                crit_normalizado = crit_clean.replace(",", ".")
                where_clauses.append("""(
                    ti.codigo_turno LIKE ? OR
                    ti.turno_base = ? OR
                    EXISTS (
                        SELECT 1 FROM turnos_vueltas tv
                        WHERE tv.turno_id = ti.id AND (tv.nro_tren = ? OR tv.nro_tren LIKE ?)
                    )
                )""")
                params.extend([f"%{crit_normalizado}%", crit_clean, crit_clean, f"%{crit_clean}%"])

            query = f"""
                SELECT DISTINCT
                    ti.id, ti.codigo_turno, ti.turno_base, ti.subindice, ti.tipo_turno,
                    ti.deposito_base, ti.toma_servicio, ti.deje_servicio, ti.duracion_jornada,
                    ti.itinerario_vigente, ti.observaciones_turno
                FROM turnos_itinerario ti
                WHERE {" AND ".join(where_clauses)}
                ORDER BY ti.turno_base ASC, ti.subindice ASC, ti.codigo_turno ASC
                LIMIT 300;
            """
            turnos_candidatos = [dict(r) for r in conn.execute(query, params).fetchall()]

            # 2. Precargar conductores en dotación para asignar titular
            conductores = [
                dict(r) for r in conn.execute(
                    """SELECT legajo, apellidos_nombres, dotacion, turnos, franco
                       FROM empleados WHERE activo = 1 AND LOWER(especialidad) LIKE '%conductor%';"""
                ).fetchall()
            ]

            # 3. Precargar novedades y reemplazos activos en la fecha evaluando rangos en Python
            raw_novedades = [
                dict(r) for r in conn.execute(
                    """SELECT n.id, n.legajo, n.apellidos_nombres, n.novedad, n.fecha_inicio, n.fecha_fin,
                              nr.id AS reemplazo_id, nr.legajo_reemplazo, nr.apellidos_nombres_reemplazo,
                              nr.turnos_reemplazo, nr.fecha_inicio AS r_fecha_inicio, nr.fecha_fin AS r_fecha_fin
                       FROM novedades n
                       LEFT JOIN novedad_reemplazos nr ON nr.novedad_id = n.id AND nr.activo = 1
                       WHERE n.activo = 1;"""
                ).fetchall()
            ]

            nov_map_por_legajo: Dict[int, Dict[str, Any]] = {}
            for n in raw_novedades:
                try:
                    d_ini = parsear_fecha(n.get("fecha_inicio"))
                except Exception:
                    continue
                if not d_ini or f_obj < d_ini:
                    continue
                f_fin_str = n.get("fecha_fin")
                if f_fin_str and str(f_fin_str).strip() not in ("", "-", "None"):
                    try:
                        d_fin = parsear_fecha(f_fin_str)
                    except Exception:
                        d_fin = None
                    if d_fin and f_obj > d_fin:
                        continue

                # La novedad está activa para el titular en la fecha
                reemplazo_aplica = False
                leg_r = n.get("legajo_reemplazo")
                if leg_r:
                    try:
                        r_ini = parsear_fecha(n.get("r_fecha_inicio")) or d_ini
                    except Exception:
                        r_ini = d_ini
                    r_fin_str = n.get("r_fecha_fin")
                    if f_obj >= r_ini:
                        if not r_fin_str or str(r_fin_str).strip() in ("", "-", "None"):
                            reemplazo_aplica = True
                        else:
                            try:
                                r_fin = parsear_fecha(r_fin_str)
                            except Exception:
                                r_fin = None
                            if not r_fin or f_obj <= r_fin:
                                reemplazo_aplica = True

                leg_tit = int(n["legajo"])
                nov_entry = dict(n)
                if not reemplazo_aplica:
                    nov_entry["legajo_reemplazo"] = None
                    nov_entry["apellidos_nombres_reemplazo"] = None

                # Si ya había una novedad sin reemplazo y esta tiene reemplazo, priorizar
                if leg_tit not in nov_map_por_legajo or reemplazo_aplica:
                    nov_map_por_legajo[leg_tit] = nov_entry

            # 3.b Precargar cambios de turno (permutas) activos para la fecha
            raw_cambios = [
                dict(r) for r in conn.execute(
                    """SELECT legajo_1, apellidos_nombres_1, turnos_1,
                              legajo_2, apellidos_nombres_2, turnos_2, fecha_cambio
                       FROM cambios_turno WHERE activo = 1;"""
                ).fetchall()
            ]
            cambios_por_legajo: Dict[int, Dict[str, Any]] = {}
            for c in raw_cambios:
                try:
                    d_c = parsear_fecha(c.get("fecha_cambio"))
                except Exception:
                    continue
                if d_c and d_c.date() == f_obj.date():
                    leg1 = int(c["legajo_1"])
                    leg2 = int(c["legajo_2"])
                    cambios_por_legajo[leg1] = {
                        "reemplazante_legajo": leg2,
                        "reemplazante_nombre": c["apellidos_nombres_2"],
                        "titular_cedente_legajo": leg1,
                        "titular_cedente_nombre": c["apellidos_nombres_1"],
                    }
                    cambios_por_legajo[leg2] = {
                        "reemplazante_legajo": leg1,
                        "reemplazante_nombre": c["apellidos_nombres_1"],
                        "titular_cedente_legajo": leg2,
                        "titular_cedente_nombre": c["apellidos_nombres_2"],
                    }

            # 3.c Precargar ordenamientos operativos activos para la fecha
            raw_ords = [
                dict(r) for r in conn.execute(
                    """SELECT o.id, o.nro_orden, o.fecha_emision, o.legajo_personal, o.nombre_personal,
                              o.deposito, o.ciclo_o_turno, o.motivo_orden, o.ref_pct,
                              GROUP_CONCAT(ot.nro_tren, ',') AS trenes_str
                       FROM ordenamientos o
                       LEFT JOIN ordenamiento_tramos ot ON ot.ordenamiento_id = o.id
                       WHERE o.activo = 1 AND o.estado = 'activo'
                       GROUP BY o.id;"""
                ).fetchall()
            ]
            ords_por_legajo: Dict[int, Dict[str, Any]] = {}
            ords_por_turno: Dict[str, Dict[str, Any]] = {}
            ords_por_tren: Dict[str, List[Dict[str, Any]]] = {}
            for o in raw_ords:
                try:
                    d_o = parsear_fecha(o.get("fecha_emision"))
                except Exception:
                    continue
                if d_o and d_o.date() == f_obj.date():
                    try:
                        ords_por_legajo[int(o["legajo_personal"])] = o
                    except Exception:
                        pass
                    ct = str(o.get("ciclo_o_turno") or "").strip()
                    if ct:
                        ords_por_turno[ct.upper()] = o
                    for tr_num in str(o.get("trenes_str") or "").split(","):
                        tr_c = tr_num.strip()
                        if tr_c:
                            ords_por_tren.setdefault(tr_c, []).append(o)

            # 4. Resolver titular y cobertura para cada turno
            from src.domain.validators import resolver_subindice_diagrama_6d

            resultados = []
            for t in turnos_candidatos:
                base = str(t["turno_base"])
                sub = t["subindice"]
                tipo = t["tipo_turno"]

                titular_encontrado: Optional[Dict[str, Any]] = None

                # Buscar en la dotación de conductores
                if tipo == "DIAGRAMA_6D" and sub is not None:
                    # Para diagrama de 6 días, cruzar con rotación de francos
                    for c in conductores:
                        if c["turnos"] == base:
                            sub_calc, es_f = resolver_subindice_diagrama_6d(dia_sem_nom, c.get("franco") or "")
                            if not es_f and sub_calc == sub:
                                titular_encontrado = c
                                break
                    if not titular_encontrado:
                        for c in conductores:
                            if c["turnos"] == base:
                                titular_encontrado = c
                                break
                else:
                    for c in conductores:
                        if c["turnos"] == base or c["turnos"] == t["codigo_turno"]:
                            titular_encontrado = c
                            break

                titular_str = "-"
                reemplazo_str = "Normal (Titular en servicio)"
                estado_alerta = "OK"

                ord_match = (
                    ords_por_turno.get(str(t["codigo_turno"]).upper())
                    or ords_por_turno.get(base.upper())
                )
                if not ord_match and crit_clean in ords_por_tren:
                    ord_match = ords_por_tren[crit_clean][0]

                if titular_encontrado:
                    leg_tit = int(titular_encontrado["legajo"])
                    nom_tit = titular_encontrado["apellidos_nombres"]
                    if not ord_match and leg_tit in ords_por_legajo:
                        ord_match = ords_por_legajo[leg_tit]

                    if ord_match:
                        nro_ord = ord_match["nro_orden"]
                        nom_ord = ord_match["nombre_personal"]
                        leg_ord = ord_match["legajo_personal"]
                        mot_ord = ord_match["motivo_orden"]
                        titular_str = f"📄 {nom_ord} (Leg. {leg_ord})"
                        reemplazo_str = f"Ordenamiento {nro_ord}: {mot_ord}"
                        estado_alerta = "ORDENAMIENTO"
                    elif leg_tit in nov_map_por_legajo:
                        nov_info = nov_map_por_legajo[leg_tit]
                        motivo = nov_info["novedad"]
                        if nov_info.get("legajo_reemplazo"):
                            leg_r = nov_info["legajo_reemplazo"]
                            nom_r = nov_info["apellidos_nombres_reemplazo"]
                            titular_str = f"⚠️ {nom_r} (Relevo Leg. {leg_r})"
                            reemplazo_str = f"Reemplazo por {motivo}: {nom_tit} (Leg. {leg_tit})"
                            estado_alerta = "RELEVO"
                        else:
                            titular_str = f"{nom_tit} (Leg. {leg_tit})"
                            reemplazo_str = f"🔴 {motivo} (Turno DESCUBIERTO)"
                            estado_alerta = "DESCUBIERTO"
                    elif leg_tit in cambios_por_legajo:
                        cambio_info = cambios_por_legajo[leg_tit]
                        leg_c = cambio_info["reemplazante_legajo"]
                        nom_c = cambio_info["reemplazante_nombre"]
                        titular_str = f"🔄 {nom_c} (Leg. {leg_c})"
                        reemplazo_str = f"Cambio de Turno (Permuta con {nom_tit} Leg. {leg_tit})"
                        estado_alerta = "CAMBIO"
                    else:
                        titular_str = f"{nom_tit} (Leg. {leg_tit})"
                else:
                    if ord_match:
                        nro_ord = ord_match["nro_orden"]
                        nom_ord = ord_match["nombre_personal"]
                        leg_ord = ord_match["legajo_personal"]
                        mot_ord = ord_match["motivo_orden"]
                        titular_str = f"📄 {nom_ord} (Leg. {leg_ord})"
                        reemplazo_str = f"Ordenamiento {nro_ord}: {mot_ord}"
                        estado_alerta = "ORDENAMIENTO"
                    else:
                        titular_str = "Sin titular asignado"
                        reemplazo_str = "Vacante / A cubrir"
                        estado_alerta = "VACANTE"

                item_res = dict(t)
                item_res["titular"] = titular_str
                item_res["cobertura"] = reemplazo_str
                item_res["estado_alerta"] = estado_alerta
                item_res["fecha_consulta"] = fecha_target_iso
                item_res["dia_servicio"] = tipo_dia
                resultados.append(item_res)

            return resultados

    def obtener_vueltas_de_turnos(self, turno_ids: List[int]) -> List[Dict[str, Any]]:
        """Obtiene todas las vueltas de múltiples turnos ordenadas por turno y nro_vuelta."""
        if not turno_ids:
            return []
        placeholders = ",".join("?" for _ in turno_ids)
        with self.db.read_connection() as conn:
            rows = conn.execute(
                f"""
                SELECT
                    tv.id, tv.turno_id, tv.nro_vuelta, tv.ref, tv.nro_tren,
                    tv.origen, tv.hora_salida, tv.destino, tv.hora_llegada, tv.observaciones_tramo,
                    ti.codigo_turno, ti.turno_base
                FROM turnos_vueltas tv
                JOIN turnos_itinerario ti ON ti.id = tv.turno_id
                WHERE tv.turno_id IN ({placeholders})
                ORDER BY ti.turno_base ASC, ti.subindice ASC, ti.codigo_turno ASC, tv.nro_vuelta ASC;
                """,
                turno_ids,
            ).fetchall()
            return [dict(r) for r in rows]


    def listar_codigos_turnos_base(self) -> List[str]:
        """Devuelve los turnos base únicos (ej: 401, 410, PROG.1CK5) para autocompletar."""
        with self.db.read_connection() as conn:
            rows = conn.execute(
                """SELECT DISTINCT turno_base FROM turnos_itinerario
                   WHERE activo = 1 ORDER BY turno_base ASC;"""
            ).fetchall()
            return [r[0] for r in rows]

    # ---------------------------------------------------------
    # CRUD de Turnos de Itinerario y Vueltas
    # ---------------------------------------------------------

    def listar_turnos_admin(
        self,
        criterio: str = "",
        deposito: Optional[str] = None,
        incluir_inactivos: bool = True,
        limite: int = 1000,
    ) -> List[Dict[str, Any]]:
        """Lista los turnos de itinerario con conteo de vueltas asociadas para administración."""
        query = """
            SELECT ti.id, ti.codigo_turno, ti.turno_base, ti.subindice, ti.tipo_turno,
                   ti.deposito_base, ti.toma_servicio, ti.deje_servicio, ti.duracion_jornada,
                   ti.itinerario_vigente, ti.observaciones_turno, ti.activo,
                   (SELECT COUNT(*) FROM turnos_vueltas tv WHERE tv.turno_id = ti.id) AS cant_vueltas
            FROM turnos_itinerario ti
            WHERE 1=1
        """
        params: List[Any] = []
        if not incluir_inactivos:
            query += " AND ti.activo = 1"
        if deposito and deposito not in ("Todas", "Todos", ""):
            query += " AND ti.deposito_base = ?"
            params.append(deposito)
        if criterio:
            c_like = f"%{criterio.strip()}%"
            query += " AND (ti.codigo_turno LIKE ? OR ti.turno_base LIKE ? OR ti.observaciones_turno LIKE ?)"
            params.extend([c_like, c_like, c_like])
        query += " ORDER BY ti.turno_base ASC, ti.subindice ASC, ti.codigo_turno ASC LIMIT ?"
        params.append(limite)

        with self.db.read_connection() as conn:
            rows = conn.execute(query, params).fetchall()
            return [dict(r) for r in rows]

    def obtener_turno_completo(self, turno_id: int) -> Optional[Dict[str, Any]]:
        """Obtiene la ficha completa de un turno de itinerario y todas sus vueltas."""
        with self.db.read_connection() as conn:
            t = conn.execute("SELECT * FROM turnos_itinerario WHERE id = ?", (turno_id,)).fetchone()
            if not t:
                return None
            res = dict(t)
            vueltas = conn.execute(
                "SELECT * FROM turnos_vueltas WHERE turno_id = ? ORDER BY nro_vuelta ASC",
                (turno_id,),
            ).fetchall()
            res["vueltas"] = [dict(v) for v in vueltas]
            return res

    def crear_turno_itinerario(
        self,
        payload: Dict[str, Any],
        vueltas: Optional[List[Dict[str, Any]]] = None,
        user_id: Optional[int] = None,
        windows_user: Optional[str] = None,
    ) -> int:
        """Crea un nuevo turno de itinerario y sus vueltas asociadas."""
        codigo = str(payload.get("codigo_turno") or "").strip()
        base = str(payload.get("turno_base") or "").strip()
        if not codigo:
            raise ValueError("El código de turno es obligatorio.")
        if not base:
            base = codigo.split(".")[0]

        deposito = str(payload.get("deposito_base") or "PC").strip()
        tipo = str(payload.get("tipo_turno") or "DIAGRAMA_6D").strip()
        itinerario = str(payload.get("itinerario_vigente") or "45").strip()

        with self.db.write_transaction() as conn:
            exist = conn.execute("SELECT id FROM turnos_itinerario WHERE codigo_turno = ?", (codigo,)).fetchone()
            if exist:
                raise ValueError(f"Ya existe un turno con el código '{codigo}'.")

            cur = conn.execute(
                """INSERT INTO turnos_itinerario (
                    codigo_turno, turno_base, subindice, tipo_turno, deposito_base,
                    toma_servicio, deje_servicio, duracion_jornada, tiempo_toma, tiempo_deje,
                    itinerario_vigente, observaciones_turno, activo, sync_status, created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'pending', datetime('now', 'localtime'), datetime('now', 'localtime'))""",
                (
                    codigo, base, payload.get("subindice"), tipo, deposito,
                    payload.get("toma_servicio"), payload.get("deje_servicio"),
                    payload.get("duracion_jornada"), payload.get("tiempo_toma"), payload.get("tiempo_deje"),
                    itinerario, payload.get("observaciones_turno"), int(payload.get("activo", 1)),
                ),
            )
            turno_id = cur.lastrowid

            if vueltas:
                for idx, v in enumerate(vueltas, start=1):
                    conn.execute(
                        """INSERT INTO turnos_vueltas (
                            turno_id, nro_vuelta, ref, nro_tren, origen, hora_salida,
                            destino, hora_llegada, observaciones_tramo, sync_status, created_at, updated_at
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'pending', datetime('now', 'localtime'), datetime('now', 'localtime'))""",
                        (
                            turno_id, v.get("nro_vuelta", idx), v.get("ref"), str(v.get("nro_tren") or ""),
                            v.get("origen"), v.get("hora_salida"), v.get("destino"), v.get("hora_llegada"),
                            v.get("observaciones_tramo"),
                        ),
                    )

            audit_payload = dict(payload)
            audit_payload["vueltas_count"] = len(vueltas or [])
            self._audit(conn, user_id, windows_user, "crear", "turno_itinerario", turno_id, None, audit_payload)
            self._enqueue_sync(conn, "turnos_itinerario", turno_id, "INSERT", audit_payload)
            return turno_id

    def actualizar_turno_itinerario(
        self,
        turno_id: int,
        payload: Dict[str, Any],
        vueltas: Optional[List[Dict[str, Any]]] = None,
        user_id: Optional[int] = None,
        windows_user: Optional[str] = None,
    ) -> None:
        """Actualiza un turno de itinerario existente y opcionalmente sus vueltas."""
        codigo = str(payload.get("codigo_turno") or "").strip()
        base = str(payload.get("turno_base") or "").strip()
        if not codigo:
            raise ValueError("El código de turno es obligatorio.")
        if not base:
            base = codigo.split(".")[0]

        deposito = str(payload.get("deposito_base") or "PC").strip()
        tipo = str(payload.get("tipo_turno") or "DIAGRAMA_6D").strip()
        itinerario = str(payload.get("itinerario_vigente") or "45").strip()

        with self.db.write_transaction() as conn:
            before = conn.execute("SELECT * FROM turnos_itinerario WHERE id = ?", (turno_id,)).fetchone()
            if not before:
                raise ValueError("El turno no existe.")

            dup = conn.execute(
                "SELECT id FROM turnos_itinerario WHERE codigo_turno = ? AND id != ?",
                (codigo, turno_id),
            ).fetchone()
            if dup:
                raise ValueError(f"Ya existe otro turno con el código '{codigo}'.")

            conn.execute(
                """UPDATE turnos_itinerario SET
                    codigo_turno = ?, turno_base = ?, subindice = ?, tipo_turno = ?, deposito_base = ?,
                    toma_servicio = ?, deje_servicio = ?, duracion_jornada = ?, tiempo_toma = ?, tiempo_deje = ?,
                    itinerario_vigente = ?, observaciones_turno = ?, activo = ?,
                    sync_status = 'pending', updated_at = datetime('now', 'localtime')
                WHERE id = ?""",
                (
                    codigo, base, payload.get("subindice"), tipo, deposito,
                    payload.get("toma_servicio"), payload.get("deje_servicio"),
                    payload.get("duracion_jornada"), payload.get("tiempo_toma"), payload.get("tiempo_deje"),
                    itinerario, payload.get("observaciones_turno"), int(payload.get("activo", 1)),
                    turno_id,
                ),
            )

            if vueltas is not None:
                conn.execute("DELETE FROM turnos_vueltas WHERE turno_id = ?", (turno_id,))
                for idx, v in enumerate(vueltas, start=1):
                    conn.execute(
                        """INSERT INTO turnos_vueltas (
                            turno_id, nro_vuelta, ref, nro_tren, origen, hora_salida,
                            destino, hora_llegada, observaciones_tramo, sync_status, created_at, updated_at
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'pending', datetime('now', 'localtime'), datetime('now', 'localtime'))""",
                        (
                            turno_id, v.get("nro_vuelta", idx), v.get("ref"), str(v.get("nro_tren") or ""),
                            v.get("origen"), v.get("hora_salida"), v.get("destino"), v.get("hora_llegada"),
                            v.get("observaciones_tramo"),
                        ),
                    )

            after = dict(payload)
            self._audit(conn, user_id, windows_user, "modificar", "turno_itinerario", turno_id, dict(before), after)
            self._enqueue_sync(conn, "turnos_itinerario", turno_id, "UPDATE", after)

    def cambiar_estado_turno(
        self,
        turno_id: int,
        activo: int,
        user_id: Optional[int] = None,
        windows_user: Optional[str] = None,
    ) -> None:
        """Activa o desactiva (soft-delete) un turno de itinerario."""
        with self.db.write_transaction() as conn:
            before = conn.execute("SELECT * FROM turnos_itinerario WHERE id = ?", (turno_id,)).fetchone()
            if not before:
                raise ValueError("El turno no existe.")
            conn.execute(
                "UPDATE turnos_itinerario SET activo = ?, sync_status = 'pending', updated_at = datetime('now', 'localtime') WHERE id = ?",
                (int(activo), turno_id),
            )
            after = dict(before)
            after["activo"] = int(activo)
            self._audit(conn, user_id, windows_user, "cambiar_estado", "turno_itinerario", turno_id, dict(before), after)
            self._enqueue_sync(conn, "turnos_itinerario", turno_id, "UPDATE", after)

    def eliminar_turno_definitivo(
        self,
        turno_id: int,
        user_id: Optional[int] = None,
        windows_user: Optional[str] = None,
    ) -> None:
        """Elimina permanentemente un turno de itinerario y sus vueltas."""
        with self.db.write_transaction() as conn:
            before = conn.execute("SELECT * FROM turnos_itinerario WHERE id = ?", (turno_id,)).fetchone()
            if not before:
                raise ValueError("El turno no existe.")
            conn.execute("DELETE FROM turnos_vueltas WHERE turno_id = ?", (turno_id,))
            conn.execute("DELETE FROM turnos_itinerario WHERE id = ?", (turno_id,))
            self._audit(conn, user_id, windows_user, "eliminar", "turno_itinerario", turno_id, dict(before), None)
            self._enqueue_sync(conn, "turnos_itinerario", turno_id, "DELETE", {"id": turno_id})

    # ---------------------------------------------------------
    # Estaciones Intermedias y Horarios de Paso por Tren (tren_paradas)
    # ---------------------------------------------------------

    def listar_paradas_tren(
        self,
        nro_tren: Optional[str] = None,
        itinerario: Optional[str] = None,
        estacion: Optional[str] = None,
        limite: int = 1000,
    ) -> List[Dict[str, Any]]:
        """Lista las estaciones intermedias y horarios de paso por tren."""
        query = """
            SELECT id, nro_tren, itinerario, regimen, orden_parada, estacion, hora_paso, created_at, updated_at
            FROM tren_paradas
            WHERE 1=1
        """
        params: List[Any] = []
        if nro_tren and str(nro_tren).strip():
            nt = str(nro_tren).strip()
            query += " AND (nro_tren = ? OR nro_tren LIKE ?)"
            params.extend([nt, f"%{nt}%"])
        if itinerario and str(itinerario).strip() not in ("", "Todos", "Todas"):
            query += " AND itinerario = ?"
            params.append(str(itinerario).strip())
        if estacion and str(estacion).strip():
            query += " AND LOWER(estacion) LIKE ?"
            params.append(f"%{str(estacion).strip().lower()}%")
        query += " ORDER BY nro_tren ASC, orden_parada ASC, hora_paso ASC LIMIT ?"
        params.append(limite)

        with self.db.read_connection() as conn:
            rows = conn.execute(query, params).fetchall()
            return [dict(r) for r in rows]

    def crear_parada_tren(
        self,
        datos: Dict[str, Any],
        user_id: Optional[int] = None,
        windows_user: Optional[str] = None,
    ) -> int:
        """Crea una parada intermedia / horario de paso para un tren."""
        nro_tren = str(datos.get("nro_tren") or "").strip()
        estacion = str(datos.get("estacion") or "").strip()
        hora_paso = str(datos.get("hora_paso") or "").strip()
        if not nro_tren:
            raise ValueError("El número de tren es obligatorio.")
        if not estacion:
            raise ValueError("La estación de parada es obligatoria.")
        if not hora_paso:
            raise ValueError("La hora de paso es obligatoria.")

        itinerario = str(datos.get("itinerario") or "").strip()
        regimen = str(datos.get("regimen") or "LV").strip()
        try:
            orden_parada = int(datos.get("orden_parada") or 1)
        except (ValueError, TypeError):
            orden_parada = 1

        with self.db.write_transaction() as conn:
            cur = conn.execute(
                """INSERT INTO tren_paradas (nro_tren, itinerario, regimen, orden_parada, estacion, hora_paso, sync_status, created_at, updated_at)
                   VALUES (?, ?, ?, ?, ?, ?, 'pending', datetime('now', 'localtime'), datetime('now', 'localtime'))""",
                (nro_tren, itinerario, regimen, orden_parada, estacion, hora_paso),
            )
            pid = cur.lastrowid
            after = dict(conn.execute("SELECT * FROM tren_paradas WHERE id = ?", (pid,)).fetchone())
            self._audit(conn, user_id, windows_user, "crear", "tren_parada", pid, None, after)
            self._enqueue_sync(conn, "tren_paradas", pid, "INSERT", after)
            return pid

    def actualizar_parada_tren(
        self,
        parada_id: int,
        datos: Dict[str, Any],
        user_id: Optional[int] = None,
        windows_user: Optional[str] = None,
    ) -> None:
        """Actualiza una parada intermedia de tren."""
        nro_tren = str(datos.get("nro_tren") or "").strip()
        estacion = str(datos.get("estacion") or "").strip()
        hora_paso = str(datos.get("hora_paso") or "").strip()
        if not nro_tren or not estacion or not hora_paso:
            raise ValueError("N° de Tren, Estación y Hora de Paso son obligatorios.")

        itinerario = str(datos.get("itinerario") or "").strip()
        regimen = str(datos.get("regimen") or "LV").strip()
        try:
            orden_parada = int(datos.get("orden_parada") or 1)
        except (ValueError, TypeError):
            orden_parada = 1

        with self.db.write_transaction() as conn:
            before = conn.execute("SELECT * FROM tren_paradas WHERE id = ?", (parada_id,)).fetchone()
            if not before:
                raise ValueError("La parada de tren no existe.")
            conn.execute(
                """UPDATE tren_paradas SET
                   nro_tren=?, itinerario=?, regimen=?, orden_parada=?, estacion=?, hora_paso=?,
                   sync_status='pending', updated_at=datetime('now', 'localtime')
                   WHERE id=?""",
                (nro_tren, itinerario, regimen, orden_parada, estacion, hora_paso, parada_id),
            )
            after = dict(conn.execute("SELECT * FROM tren_paradas WHERE id = ?", (parada_id,)).fetchone())
            self._audit(conn, user_id, windows_user, "modificar", "tren_parada", parada_id, dict(before), after)
            self._enqueue_sync(conn, "tren_paradas", parada_id, "UPDATE", after)

    def eliminar_parada_tren(
        self,
        parada_id: int,
        user_id: Optional[int] = None,
        windows_user: Optional[str] = None,
    ) -> None:
        """Elimina una parada de tren."""
        with self.db.write_transaction() as conn:
            before = conn.execute("SELECT * FROM tren_paradas WHERE id = ?", (parada_id,)).fetchone()
            if not before:
                raise ValueError("La parada de tren no existe.")
            conn.execute("DELETE FROM tren_paradas WHERE id = ?", (parada_id,))
            self._audit(conn, user_id, windows_user, "eliminar", "tren_parada", parada_id, dict(before), None)
            self._enqueue_sync(conn, "tren_paradas", parada_id, "DELETE", {"id": parada_id})

    def importar_paradas_tren_excel(
        self,
        file_path: str,
        itinerario_default: Optional[str] = None,
        user_id: Optional[int] = None,
        windows_user: Optional[str] = None,
    ) -> int:
        """Importa estaciones intermedias y horarios de paso por tren desde un archivo Excel.

        Soporta tanto:
        1) Planillas dedicadas con columnas (TREN/NRO_TREN, ESTACION/PARADA, HORA/HORA_PASO, ORDEN, ITINERARIO, REGIMEN).
        2) Planillas de itinerarios estándar (con ORIGEN, HORA_SALIDA, DESTINO, HORA_LLEGADA por vuelta).
        """
        import openpyxl

        wb = openpyxl.load_workbook(file_path, data_only=True)
        ws = wb.active

        headers_raw = [str(ws.cell(1, c).value or "").strip().upper() for c in range(1, ws.max_column + 1)]
        col_map: Dict[str, int] = {h: idx + 1 for idx, h in enumerate(headers_raw) if h}

        def _find_col(*candidates: str) -> Optional[int]:
            for cand in candidates:
                for h_name, c_idx in col_map.items():
                    if cand in h_name:
                        return c_idx
            return None

        c_tren = _find_col("NRO_TREN", "TREN", "N° TREN")
        c_est = _find_col("ESTACION", "ESTACIÓN", "PARADA")
        c_hora = _find_col("HORA_PASO", "HORA", "PASO", "HORARIO")
        c_orden = _find_col("ORDEN", "SECUENCIA")
        c_itin = _find_col("ITINERARIO", "CIRC", "VIGENCIA")
        c_reg = _find_col("REGIMEN", "RÉGIMEN", "DIA", "TIPO")

        c_origen = _find_col("ORIGEN", "DESDE", "SALE DE")
        c_h_sal = _find_col("SALIDA", "H_SALIDA", "HORA_SALIDA")
        c_destino = _find_col("DESTINO", "HASTA", "LLEGA A")
        c_h_lleg = _find_col("LLEGADA", "H_LLEGADA", "HORA_LLEGADA")

        registros_a_insertar: List[Tuple[str, str, str, int, str, str]] = []
        orden_tracker: Dict[Tuple[str, str], int] = {}

        for r_idx in range(2, ws.max_row + 1):
            if c_tren and c_est and c_hora:
                nt = str(ws.cell(r_idx, c_tren).value or "").strip()
                est = str(ws.cell(r_idx, c_est).value or "").strip()
                hr_val = ws.cell(r_idx, c_hora).value
                hr = hr_val.strftime("%H:%M") if hasattr(hr_val, "strftime") else str(hr_val or "").strip()
                if not nt or not est or not hr or nt in ("None", "-"):
                    continue
                itin = str(ws.cell(r_idx, c_itin).value or itinerario_default or "").strip() if c_itin else (itinerario_default or "")
                reg = str(ws.cell(r_idx, c_reg).value or "LV").strip() if c_reg else "LV"
                ord_key = (nt, itin)
                if c_orden and ws.cell(r_idx, c_orden).value is not None:
                    try:
                        ord_num = int(ws.cell(r_idx, c_orden).value)
                    except (ValueError, TypeError):
                        orden_tracker[ord_key] = orden_tracker.get(ord_key, 0) + 1
                        ord_num = orden_tracker[ord_key]
                else:
                    orden_tracker[ord_key] = orden_tracker.get(ord_key, 0) + 1
                    ord_num = orden_tracker[ord_key]
                registros_a_insertar.append((nt, itin, reg, ord_num, est, hr))
            elif c_tren and c_origen and c_h_sal and c_destino and c_h_lleg:
                nt = str(ws.cell(r_idx, c_tren).value or "").strip()
                if not nt or nt in ("None", "-"):
                    continue
                itin = str(ws.cell(r_idx, c_itin).value or itinerario_default or "").strip() if c_itin else (itinerario_default or "")
                reg = str(ws.cell(r_idx, c_reg).value or "LV").strip() if c_reg else "LV"
                orig = str(ws.cell(r_idx, c_origen).value or "").strip()
                hs_val = ws.cell(r_idx, c_h_sal).value
                hs = hs_val.strftime("%H:%M") if hasattr(hs_val, "strftime") else str(hs_val or "").strip()
                dest = str(ws.cell(r_idx, c_destino).value or "").strip()
                hl_val = ws.cell(r_idx, c_h_lleg).value
                hl = hl_val.strftime("%H:%M") if hasattr(hl_val, "strftime") else str(hl_val or "").strip()
                ord_key = (nt, itin)
                for est, hr in ((orig, hs), (dest, hl)):
                    if est and hr:
                        orden_tracker[ord_key] = orden_tracker.get(ord_key, 0) + 1
                        registros_a_insertar.append((nt, itin, reg, orden_tracker[ord_key], est, hr))
            else:
                # Fallback posicional: col 1=Tren, col 2=Estacion, col 3=Hora, col 4=Orden, col 5=Itinerario
                nt = str(ws.cell(r_idx, 1).value or "").strip()
                est = str(ws.cell(r_idx, 2).value or "").strip()
                hr_val = ws.cell(r_idx, 3).value
                hr = hr_val.strftime("%H:%M") if hasattr(hr_val, "strftime") else str(hr_val or "").strip()
                if not nt or not est or not hr or nt in ("None", "-"):
                    continue
                itin = str(ws.cell(r_idx, 5).value or itinerario_default or "").strip()
                ord_key = (nt, itin)
                orden_tracker[ord_key] = orden_tracker.get(ord_key, 0) + 1
                registros_a_insertar.append((nt, itin, "LV", orden_tracker[ord_key], est, hr))

        insertados = 0
        with self.db.write_transaction() as conn:
            for nt, itin, reg, ord_num, est, hr in registros_a_insertar:
                existe = conn.execute(
                    "SELECT id FROM tren_paradas WHERE nro_tren=? AND itinerario=? AND UPPER(estacion)=UPPER(?) AND hora_paso=?",
                    (nt, itin, est, hr),
                ).fetchone()
                if not existe:
                    conn.execute(
                        """INSERT INTO tren_paradas (nro_tren, itinerario, regimen, orden_parada, estacion, hora_paso, sync_status, created_at, updated_at)
                           VALUES (?, ?, ?, ?, ?, ?, 'pending', datetime('now', 'localtime'), datetime('now', 'localtime'))""",
                        (nt, itin, reg, ord_num, est, hr),
                    )
                    insertados += 1
            self._audit(conn, user_id, windows_user, "importar", "tren_paradas", None, None, {"insertados": insertados})
        return insertados

    def obtener_estaciones_y_horarios_por_tren(
        self,
        nro_tren: str,
        itinerario: Optional[str] = None,
    ) -> List[Dict[str, str]]:
        """Obtiene la secuencia de estaciones y horarios de paso para un N° de tren dado.

        Combina registros de `tren_paradas` y tramos de `turnos_vueltas` para garantizar cobertura total.
        """
        nt = str(nro_tren or "").strip()
        if not nt:
            return []

        vistas: Dict[str, str] = {}
        resultado: List[Dict[str, str]] = []

        with self.db.read_connection() as conn:
            q_paradas = "SELECT estacion, hora_paso, orden_parada FROM tren_paradas WHERE nro_tren = ?"
            p_paradas: List[Any] = [nt]
            if itinerario and str(itinerario).strip() not in ("", "Todos", "Todas"):
                q_paradas += " AND (itinerario = ? OR itinerario = '')"
                p_paradas.append(str(itinerario).strip())
            q_paradas += " ORDER BY orden_parada ASC, hora_paso ASC"

            for r in conn.execute(q_paradas, p_paradas).fetchall():
                est = str(r["estacion"] or "").strip()
                hr = str(r["hora_paso"] or "").strip()
                if est and est.upper() not in vistas:
                    vistas[est.upper()] = hr
                    resultado.append({"estacion": est, "hora": hr})

            # Complementar con turnos_vueltas
            q_vueltas = """
                SELECT v.origen, v.hora_salida, v.destino, v.hora_llegada
                FROM turnos_vueltas v
                JOIN turnos_itinerario t ON t.id = v.turno_id
                WHERE v.nro_tren = ?
            """
            p_vueltas: List[Any] = [nt]
            if itinerario and str(itinerario).strip() not in ("", "Todos", "Todas"):
                q_vueltas += " AND t.itinerario_vigente = ?"
                p_vueltas.append(str(itinerario).strip())
            q_vueltas += " ORDER BY v.hora_salida ASC"

            for r in conn.execute(q_vueltas, p_vueltas).fetchall():
                for est_raw, hr_raw in ((r["origen"], r["hora_salida"]), (r["destino"], r["hora_llegada"])):
                    est = str(est_raw or "").strip()
                    hr = str(hr_raw or "").strip()
                    if est and est.upper() not in vistas:
                        vistas[est.upper()] = hr
                        resultado.append({"estacion": est, "hora": hr})

        resultado.sort(key=lambda x: x.get("hora") or "99:99")
        return resultado

    # ---------------------------------------------------------
    # Módulo de Ordenamientos Operativos
    # ---------------------------------------------------------

    def generar_siguiente_nro_orden(self, deposito: str = "PC") -> str:
        """Genera el siguiente número correlativo de orden con formato '<DOTACION> 0000-XXXX'."""
        prefijo = (deposito or "PC").strip().upper() or "PC"
        with self.db.read_connection() as conn:
            rows = conn.execute(
                "SELECT nro_orden FROM ordenamientos WHERE UPPER(nro_orden) LIKE ?",
                (f"{prefijo} %",),
            ).fetchall()
            max_num = 0
            for r in rows:
                txt = str(r[0] or "").strip()
                if "-" in txt:
                    try:
                        num_part = int(txt.split("-")[-1].strip())
                        if num_part > max_num:
                            max_num = num_part
                    except ValueError:
                        pass
            sig = max_num + 1
            return f"{prefijo} 0000-{sig:04d}"

    def crear_ordenamiento(
        self,
        datos: Dict[str, Any],
        tramos: List[Dict[str, Any]],
        user_id: Optional[int] = None,
        windows_user: Optional[str] = None,
    ) -> int:
        """Crea un nuevo ordenamiento operativo y sus tramos asociados."""
        es_valido, err = validar_datos_ordenamiento(datos, tramos)
        if not es_valido:
            raise ValueError(err)

        deposito = str(datos.get("deposito") or "PC").strip().upper()
        nro_orden = str(datos.get("nro_orden") or "").strip()
        if not nro_orden:
            nro_orden = self.generar_siguiente_nro_orden(deposito)

        win_usr = (windows_user or datos.get("usuario_windows") or "").strip() or obtener_usuario_windows()
        ahora_dt = datetime.now()
        fecha_emision = str(datos.get("fecha_emision") or ahora_dt.strftime("%d/%m/%Y")).strip()
        hora_emision = str(datos.get("hora_emision") or ahora_dt.strftime("%H:%M")).strip()

        with self.db.write_transaction() as conn:
            dup = conn.execute("SELECT id FROM ordenamientos WHERE nro_orden = ?", (nro_orden,)).fetchone()
            if dup:
                nro_orden = self.generar_siguiente_nro_orden(deposito)

            cur = conn.execute(
                """INSERT INTO ordenamientos (
                    nro_orden, fecha_emision, hora_emision, legajo_personal, nombre_personal,
                    deposito, itinerario, ciclo_o_turno, regimen, motivo_orden, observacion_pie,
                    ref_pct, emisor, usuario_id, usuario_windows, estado, activo, sync_status,
                    created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'activo', 1, 'pending', datetime('now', 'localtime'), datetime('now', 'localtime'))""",
                (
                    nro_orden,
                    fecha_emision,
                    hora_emision,
                    int(str(datos["legajo_personal"]).strip()),
                    str(datos.get("nombre_personal") or "").strip(),
                    deposito,
                    str(datos.get("itinerario") or "").strip(),
                    str(datos.get("ciclo_o_turno") or "").strip(),
                    str(datos.get("regimen") or "LV").strip(),
                    str(datos.get("motivo_orden") or "").strip(),
                    str(datos.get("observacion_pie") or "").strip(),
                    str(datos.get("ref_pct") or "").strip(),
                    str(datos.get("emisor") or "").strip(),
                    user_id,
                    win_usr,
                ),
            )
            ord_id = cur.lastrowid

            for idx, t in enumerate(tramos, start=1):
                conn.execute(
                    """INSERT INTO ordenamiento_tramos (
                        ordenamiento_id, nro_secuencia, ref, nro_tren, origen,
                        hora_salida, destino, hora_llegada, observaciones, sync_status,
                        created_at, updated_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'pending', datetime('now', 'localtime'), datetime('now', 'localtime'))""",
                    (
                        ord_id,
                        int(t.get("nro_secuencia") or idx),
                        str(t.get("ref") or "").strip(),
                        str(t.get("nro_tren") or "").strip(),
                        str(t.get("origen") or "").strip(),
                        str(t.get("hora_salida") or "").strip(),
                        str(t.get("destino") or "").strip(),
                        str(t.get("hora_llegada") or "").strip(),
                        str(t.get("observaciones") or "").strip(),
                    ),
                )

            after = dict(conn.execute("SELECT * FROM ordenamientos WHERE id = ?", (ord_id,)).fetchone())
            after["tramos_count"] = len(tramos)
            self._audit(conn, user_id, win_usr, "crear", "ordenamiento", ord_id, None, after)
            self._enqueue_sync(conn, "ordenamientos", ord_id, "INSERT", after)
            return ord_id

    def actualizar_ordenamiento(
        self,
        ord_id: int,
        datos: Dict[str, Any],
        tramos: List[Dict[str, Any]],
        user_id: Optional[int] = None,
        windows_user: Optional[str] = None,
        bypass_window: bool = False,
    ) -> None:
        """Actualiza un ordenamiento existente y reemplaza sus tramos."""
        es_valido, err = validar_datos_ordenamiento(datos, tramos)
        if not es_valido:
            raise ValueError(err)

        win_usr = (windows_user or datos.get("usuario_windows") or "").strip() or obtener_usuario_windows()
        deposito = str(datos.get("deposito") or "PC").strip().upper()

        with self.db.write_transaction() as conn:
            before = conn.execute("SELECT * FROM ordenamientos WHERE id = ?", (ord_id,)).fetchone()
            if not before:
                raise ValueError("El ordenamiento no existe.")

            ref_time = before["created_at"] or f"{before['fecha_emision']} {before['hora_emision']}"
            if not bypass_window and not self.dentro_de_ventana(ref_time, "editar_horas"):
                raise ValueError("El ordenamiento superó el tiempo permitido para editar.")

            nro_orden = str(datos.get("nro_orden") or before["nro_orden"]).strip()
            conn.execute(
                """UPDATE ordenamientos SET
                    nro_orden=?, fecha_emision=?, hora_emision=?, legajo_personal=?, nombre_personal=?,
                    deposito=?, itinerario=?, ciclo_o_turno=?, regimen=?, motivo_orden=?, observacion_pie=?,
                    ref_pct=?, emisor=?, usuario_windows=?, sync_status='pending', updated_at=datetime('now', 'localtime')
                   WHERE id=?""",
                (
                    nro_orden,
                    str(datos.get("fecha_emision") or before["fecha_emision"]).strip(),
                    str(datos.get("hora_emision") or before["hora_emision"]).strip(),
                    int(str(datos["legajo_personal"]).strip()),
                    str(datos.get("nombre_personal") or "").strip(),
                    deposito,
                    str(datos.get("itinerario") or "").strip(),
                    str(datos.get("ciclo_o_turno") or "").strip(),
                    str(datos.get("regimen") or "LV").strip(),
                    str(datos.get("motivo_orden") or "").strip(),
                    str(datos.get("observacion_pie") or "").strip(),
                    str(datos.get("ref_pct") or "").strip(),
                    str(datos.get("emisor") or before["emisor"] or "").strip(),
                    win_usr,
                    ord_id,
                ),
            )

            conn.execute("DELETE FROM ordenamiento_tramos WHERE ordenamiento_id = ?", (ord_id,))
            for idx, t in enumerate(tramos, start=1):
                conn.execute(
                    """INSERT INTO ordenamiento_tramos (
                        ordenamiento_id, nro_secuencia, ref, nro_tren, origen,
                        hora_salida, destino, hora_llegada, observaciones, sync_status,
                        created_at, updated_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'pending', datetime('now', 'localtime'), datetime('now', 'localtime'))""",
                    (
                        ord_id,
                        int(t.get("nro_secuencia") or idx),
                        str(t.get("ref") or "").strip(),
                        str(t.get("nro_tren") or "").strip(),
                        str(t.get("origen") or "").strip(),
                        str(t.get("hora_salida") or "").strip(),
                        str(t.get("destino") or "").strip(),
                        str(t.get("hora_llegada") or "").strip(),
                        str(t.get("observaciones") or "").strip(),
                    ),
                )

            after = dict(conn.execute("SELECT * FROM ordenamientos WHERE id = ?", (ord_id,)).fetchone())
            after["tramos_count"] = len(tramos)
            self._audit(conn, user_id, win_usr, "modificar", "ordenamiento", ord_id, dict(before), after)
            self._enqueue_sync(conn, "ordenamientos", ord_id, "UPDATE", after)

    def eliminar_ordenamiento(
        self,
        ord_id: int,
        user_id: Optional[int] = None,
        windows_user: Optional[str] = None,
        bypass_window: bool = False,
    ) -> None:
        """Eliminación lógica (soft-delete) de un ordenamiento enviándolo a la papelera."""
        win_usr = (windows_user or "").strip() or obtener_usuario_windows()
        with self.db.write_transaction() as conn:
            before = conn.execute("SELECT * FROM ordenamientos WHERE id = ?", (ord_id,)).fetchone()
            if not before:
                raise ValueError("El ordenamiento no existe.")
            ref_time = before["created_at"] or f"{before['fecha_emision']} {before['hora_emision']}"
            if not bypass_window and not self.dentro_de_ventana(ref_time, "eliminar_horas"):
                raise ValueError("El ordenamiento superó el tiempo permitido para eliminar.")

            ahora = self.db.now()
            conn.execute(
                "UPDATE ordenamientos SET activo=0, estado='eliminado', deleted_at=?, updated_at=?, sync_status='pending' WHERE id=?",
                (ahora, ahora, ord_id),
            )
            after = {"activo": 0, "estado": "eliminado", "deleted_at": ahora}
            self._audit(conn, user_id, win_usr, "eliminar", "ordenamiento", ord_id, dict(before), after)
            self._enqueue_sync(conn, "ordenamientos", ord_id, "UPDATE", after)

    def obtener_ordenamiento_completo(self, ord_id: int) -> Optional[Dict[str, Any]]:
        """Obtiene la cabecera del ordenamiento y todos sus tramos ordenados por secuencia."""
        with self.db.read_connection() as conn:
            row = conn.execute("SELECT * FROM ordenamientos WHERE id = ?", (ord_id,)).fetchone()
            if not row:
                return None
            res = dict(row)
            tramos = conn.execute(
                "SELECT * FROM ordenamiento_tramos WHERE ordenamiento_id = ? ORDER BY nro_secuencia ASC, id ASC",
                (ord_id,),
            ).fetchall()
            res["tramos"] = [dict(t) for t in tramos]
            return res

    def listar_ordenamientos(
        self,
        filtros: Optional[Dict[str, Any]] = None,
        incluir_inactivos: bool = False,
        orden_asc: bool = True,
    ) -> List[Dict[str, Any]]:
        """Lista los ordenamientos operativos con resumen de trenes involucrados."""
        with self.db.read_connection() as conn:
            query = """
                SELECT o.*,
                       COALESCE(GROUP_CONCAT(NULLIF(TRIM(ot.nro_tren), ''), ', '), '') AS trenes_involucrados,
                       COUNT(ot.id) AS cant_tramos
                FROM ordenamientos o
                LEFT JOIN ordenamiento_tramos ot ON ot.ordenamiento_id = o.id
            """
            conds = []
            params: List[Any] = []
            if not incluir_inactivos:
                conds.append("o.activo = 1 AND o.estado = 'activo'")

            if filtros:
                if filtros.get("deposito") and str(filtros["deposito"]).strip() not in ("", "Todas", "Todos"):
                    conds.append("UPPER(TRIM(o.deposito)) = UPPER(TRIM(?))")
                    params.append(str(filtros["deposito"]).strip())
                if filtros.get("legajo"):
                    conds.append("o.legajo_personal = ?")
                    params.append(int(filtros["legajo"]))
                if filtros.get("texto"):
                    like = f"%{str(filtros['texto']).strip()}%"
                    conds.append("""(
                        o.nro_orden LIKE ? OR
                        CAST(o.legajo_personal AS TEXT) LIKE ? OR
                        o.nombre_personal LIKE ? OR
                        o.ciclo_o_turno LIKE ? OR
                        o.motivo_orden LIKE ? OR
                        EXISTS (
                            SELECT 1 FROM ordenamiento_tramos ot2
                            WHERE ot2.ordenamiento_id = o.id AND ot2.nro_tren LIKE ?
                        )
                    )""")
                    params.extend([like, like, like, like, like, like])

            if conds:
                query += " WHERE " + " AND ".join(conds)
            query += " GROUP BY o.id"
            query += " ORDER BY o.id ASC" if orden_asc else " ORDER BY o.id DESC"

            rows = conn.execute(query, params).fetchall()
            resultado = [dict(r) for r in rows]

            if filtros and filtros.get("fecha"):
                f_filtro = parsear_fecha(filtros["fecha"], obligatorio=False)
                if f_filtro:
                    filtrados_fecha = []
                    for item in resultado:
                        d_em = parsear_fecha(item.get("fecha_emision"), obligatorio=False)
                        if d_em and d_em.date() == f_filtro.date():
                            filtrados_fecha.append(item)
                    return filtrados_fecha

            return resultado

    def buscar_ordenamientos_activos(
        self,
        criterio: str,
        fecha_target_iso: Optional[str] = None,
    ) -> List[Dict[str, Any]]:
        """Busca ordenamientos activos cuyos tramos, turno, legajo o N° de orden coincidan con el criterio (ej. N° de tren)."""
        crit = str(criterio or "").strip()
        if not crit:
            return []

        f_obj = parsear_fecha(fecha_target_iso, obligatorio=False) if fecha_target_iso else None
        like = f"%{crit}%"

        with self.db.read_connection() as conn:
            rows = conn.execute(
                """
                SELECT o.*,
                       COALESCE(GROUP_CONCAT(NULLIF(TRIM(ot.nro_tren), ''), ', '), '') AS trenes_involucrados,
                       COUNT(ot.id) AS cant_tramos
                FROM ordenamientos o
                LEFT JOIN ordenamiento_tramos ot ON ot.ordenamiento_id = o.id
                WHERE o.activo = 1 AND o.estado = 'activo'
                  AND (
                      o.nro_orden LIKE ? OR
                      o.ciclo_o_turno LIKE ? OR
                      CAST(o.legajo_personal AS TEXT) = ? OR
                      EXISTS (
                          SELECT 1 FROM ordenamiento_tramos ot2
                          WHERE ot2.ordenamiento_id = o.id AND (ot2.nro_tren = ? OR ot2.nro_tren LIKE ?)
                      )
                  )
                GROUP BY o.id
                ORDER BY o.id ASC
                """,
                (like, like, crit, crit, like),
            ).fetchall()

            resultados: List[Dict[str, Any]] = []
            for r in rows:
                d_item = dict(r)
                if f_obj:
                    d_em = parsear_fecha(d_item.get("fecha_emision"), obligatorio=False)
                    if d_em and d_em.date() != f_obj.date():
                        continue
                tramos = conn.execute(
                    "SELECT * FROM ordenamiento_tramos WHERE ordenamiento_id = ? ORDER BY nro_secuencia ASC, id ASC",
                    (d_item["id"],),
                ).fetchall()
                d_item["tramos"] = [dict(t) for t in tramos]
                resultados.append(d_item)
            return resultados


