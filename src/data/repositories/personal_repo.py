"""Repositorio de persistencia para Personal, Catálogos, RBAC y Auditoría."""

from __future__ import annotations

import json
from datetime import datetime
from typing import Any, Sequence

from ...domain.enums import SyncOperation
from ..database import DatabaseManager


class PersonalRepository:
    """CRUD de empleados, dotaciones, tipos de novedad, roles, usuarios y auditoría."""

    def __init__(self, db: DatabaseManager) -> None:
        self.db = db

    def _ahora_iso(self) -> str:
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def _encolar(self, conn, tabla: str, pk: str | int, operacion: str, payload: dict[str, Any]) -> None:
        conn.execute(
            """INSERT INTO sync_outbox (tabla, registro_pk, operacion, payload_json, creado_en)
               VALUES (?, ?, ?, ?, ?)
               ON CONFLICT(tabla, registro_pk) DO UPDATE SET
                   operacion = excluded.operacion,
                   payload_json = excluded.payload_json,
                   creado_en = excluded.creado_en;""",
            (tabla, str(pk), operacion, json.dumps(payload, ensure_ascii=False, default=str), self._ahora_iso()),
        )

    # --- Empleados ---

    def get_base_rows(self) -> list[tuple[Any, ...]]:
        """Devuelve tuplas (legajo, apellidos_nombres, especialidad, dotacion, turnos, franco, franco_horario, jornada_horas)."""
        with self.db.read_connection() as conn:
            rows = conn.execute(
                """SELECT legajo, apellidos_nombres, especialidad, dotacion, turnos, franco, franco_horario, jornada_horas
                   FROM empleados WHERE activo = 1 ORDER BY apellidos_nombres;"""
            ).fetchall()
            return [tuple(r) for r in rows]

    def listar_empleados(self, incluir_inactivos: bool = True) -> list[dict[str, Any]]:
        with self.db.read_connection() as conn:
            sql = "SELECT * FROM empleados"
            if not incluir_inactivos:
                sql += " WHERE activo = 1"
            sql += " ORDER BY apellidos_nombres;"
            return [dict(r) for r in conn.execute(sql).fetchall()]

    def obtener_empleado_por_id(self, empleado_id: int) -> dict[str, Any] | None:
        with self.db.read_connection() as conn:
            row = conn.execute("SELECT * FROM empleados WHERE id = ?;", (empleado_id,)).fetchone()
            return dict(row) if row else None

    def obtener_empleado_por_legajo(self, legajo: int) -> dict[str, Any] | None:
        with self.db.read_connection() as conn:
            row = conn.execute("SELECT * FROM empleados WHERE legajo = ?;", (legajo,)).fetchone()
            return dict(row) if row else None

    def guardar_empleado(self, data: dict[str, Any], empleado_id: int | None = None) -> int:
        ahora = self._ahora_iso()
        legajo = int(data["legajo"])
        jornada_hs = 7 if str(data.get("jornada_horas", "")).strip() == "7" else 6
        with self.db.write_transaction() as conn:
            if empleado_id is None:
                cur = conn.execute(
                    """INSERT INTO empleados
                       (legajo, apellidos_nombres, especialidad, dotacion, turnos, franco,
                        franco_horario, jornada_horas, activo, actualizado_en, sync_status, updated_at)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'pending', ?);""",
                    (
                        legajo,
                        str(data.get("apellidos_nombres", "")).strip(),
                        str(data.get("especialidad", "")).strip(),
                        str(data.get("dotacion", "")).strip(),
                        str(data.get("turnos", "")).strip(),
                        str(data.get("franco", "")).strip(),
                        str(data.get("franco_horario", "")).strip(),
                        jornada_hs,
                        1 if data.get("activo", True) else 0,
                        ahora,
                        ahora,
                    ),
                )
                empleado_id = cur.lastrowid
            else:
                conn.execute(
                    """UPDATE empleados SET
                           legajo = ?, apellidos_nombres = ?, especialidad = ?, dotacion = ?,
                           turnos = ?, franco = ?, franco_horario = ?, jornada_horas = ?,
                           activo = ?, actualizado_en = ?, sync_status = 'pending', updated_at = ?
                       WHERE id = ?;""",
                    (
                        legajo,
                        str(data.get("apellidos_nombres", "")).strip(),
                        str(data.get("especialidad", "")).strip(),
                        str(data.get("dotacion", "")).strip(),
                        str(data.get("turnos", "")).strip(),
                        str(data.get("franco", "")).strip(),
                        str(data.get("franco_horario", "")).strip(),
                        jornada_hs,
                        1 if data.get("activo", True) else 0,
                        ahora,
                        ahora,
                        empleado_id,
                    ),
                )

            row = dict(conn.execute("SELECT * FROM empleados WHERE id = ?;", (empleado_id,)).fetchone())
            self._encolar(conn, "empleados", legajo, SyncOperation.UPSERT.value, row)
            return empleado_id

    def cambiar_estado_empleado(self, empleado_id: int) -> int:
        ahora = self._ahora_iso()
        with self.db.write_transaction() as conn:
            conn.execute("UPDATE empleados SET activo = 1 - activo, sync_status = 'pending', updated_at = ? WHERE id = ?;", (ahora, empleado_id))
            row = dict(conn.execute("SELECT * FROM empleados WHERE id = ?;", (empleado_id,)).fetchone())
            self._encolar(conn, "empleados", row["legajo"], SyncOperation.UPSERT.value, row)
            return int(row["activo"])

    # --- Catálogos: Tipos de novedad, Dotaciones, Personal estación ---

    def listar_tipos_novedad(self, incluir_inactivos: bool = True) -> list[dict[str, Any]]:
        with self.db.read_connection() as conn:
            sql = "SELECT id, nombre, dias, exportar_descubierto, activo FROM tipos_novedad"
            if not incluir_inactivos:
                sql += " WHERE activo = 1"
            sql += " ORDER BY nombre;"
            return [dict(r) for r in conn.execute(sql).fetchall()]

    def guardar_tipo_novedad(self, nombre: str, dias: int = 0, exportar_descubierto: int = 1, activo: int = 1, tipo_id: int | None = None) -> int:
        ahora = self._ahora_iso()
        with self.db.write_transaction() as conn:
            if tipo_id is None:
                cur = conn.execute(
                    """INSERT INTO tipos_novedad (nombre, dias, exportar_descubierto, activo, sync_status, updated_at)
                       VALUES (?, ?, ?, ?, 'pending', ?);""",
                    (nombre.strip(), int(dias), int(exportar_descubierto), int(activo), ahora),
                )
                tipo_id = cur.lastrowid
            else:
                conn.execute(
                    """UPDATE tipos_novedad SET nombre = ?, dias = ?, exportar_descubierto = ?, activo = ?, sync_status = 'pending', updated_at = ?
                       WHERE id = ?;""",
                    (nombre.strip(), int(dias), int(exportar_descubierto), int(activo), ahora, tipo_id),
                )
            row = dict(conn.execute("SELECT * FROM tipos_novedad WHERE id = ?;", (tipo_id,)).fetchone())
            self._encolar(conn, "tipos_novedad", tipo_id, SyncOperation.UPSERT.value, row)
            return tipo_id

    def listar_dotaciones(self, incluir_inactivos: bool = True) -> list[dict[str, Any]]:
        with self.db.read_connection() as conn:
            sql = "SELECT id, nombre, activo FROM dotaciones"
            if not incluir_inactivos:
                sql += " WHERE activo = 1"
            sql += " ORDER BY nombre;"
            return [dict(r) for r in conn.execute(sql).fetchall()]

    def guardar_dotacion(self, nombre: str, activo: int = 1, dotacion_id: int | None = None) -> int:
        ahora = self._ahora_iso()
        with self.db.write_transaction() as conn:
            if dotacion_id is None:
                cur = conn.execute(
                    """INSERT INTO dotaciones (nombre, activo, sync_status, updated_at) VALUES (?, ?, 'pending', ?);""",
                    (nombre.strip(), int(activo), ahora),
                )
                dotacion_id = cur.lastrowid
            else:
                conn.execute(
                    """UPDATE dotaciones SET nombre = ?, activo = ?, sync_status = 'pending', updated_at = ? WHERE id = ?;""",
                    (nombre.strip(), int(activo), ahora, dotacion_id),
                )
            row = dict(conn.execute("SELECT * FROM dotaciones WHERE id = ?;", (dotacion_id,)).fetchone())
            self._encolar(conn, "dotaciones", dotacion_id, SyncOperation.UPSERT.value, row)
            return dotacion_id

    def listar_personal_estacion(self, incluir_inactivos: bool = True) -> list[dict[str, Any]]:
        with self.db.read_connection() as conn:
            sql = "SELECT id, nombre, dotacion, activo FROM personal_estacion"
            if not incluir_inactivos:
                sql += " WHERE activo = 1"
            sql += " ORDER BY nombre;"
            return [dict(r) for r in conn.execute(sql).fetchall()]

    def guardar_personal_estacion(self, nombre: str, dotacion: str | None = None, activo: int = 1, item_id: int | None = None) -> int:
        ahora = self._ahora_iso()
        with self.db.write_transaction() as conn:
            if item_id is None:
                cur = conn.execute(
                    """INSERT INTO personal_estacion (nombre, dotacion, activo, sync_status, updated_at) VALUES (?, ?, ?, 'pending', ?);""",
                    (nombre.strip(), (dotacion or "").strip() or None, int(activo), ahora),
                )
                item_id = cur.lastrowid
            else:
                conn.execute(
                    """UPDATE personal_estacion SET nombre = ?, dotacion = ?, activo = ?, sync_status = 'pending', updated_at = ? WHERE id = ?;""",
                    (nombre.strip(), (dotacion or "").strip() or None, int(activo), ahora, item_id),
                )
            row = dict(conn.execute("SELECT * FROM personal_estacion WHERE id = ?;", (item_id,)).fetchone())
            self._encolar(conn, "personal_estacion", item_id, SyncOperation.UPSERT.value, row)
            return item_id

    # --- Destinatarios Informe ---

    def listar_destinatarios_informe(self, incluir_inactivos: bool = True) -> list[dict[str, Any]]:
        with self.db.read_connection() as conn:
            sql = "SELECT id, nombre, email, activo FROM destinatarios_informe"
            if not incluir_inactivos:
                sql += " WHERE activo = 1"
            sql += " ORDER BY email;"
            return [dict(r) for r in conn.execute(sql).fetchall()]

    def guardar_destinatario_informe(self, nombre: str, email: str, activo: int = 1, item_id: int | None = None) -> int:
        ahora = self._ahora_iso()
        with self.db.write_transaction() as conn:
            if item_id is None:
                cur = conn.execute(
                    """INSERT INTO destinatarios_informe (nombre, email, activo, sync_status, updated_at) VALUES (?, ?, ?, 'pending', ?);""",
                    (nombre.strip(), email.strip(), int(activo), ahora),
                )
                item_id = cur.lastrowid
            else:
                conn.execute(
                    """UPDATE destinatarios_informe SET nombre = ?, email = ?, activo = ?, sync_status = 'pending', updated_at = ? WHERE id = ?;""",
                    (nombre.strip(), email.strip(), int(activo), ahora, item_id),
                )
            row = dict(conn.execute("SELECT * FROM destinatarios_informe WHERE id = ?;", (item_id,)).fetchone())
            self._encolar(conn, "destinatarios_informe", item_id, SyncOperation.UPSERT.value, row)
            return item_id

    # --- Configuración Clave-Valor ---

    def get_configuracion(self, clave: str, default: Any = None) -> Any:
        with self.db.read_connection() as conn:
            row = conn.execute("SELECT valor FROM configuracion WHERE clave = ?;", (clave,)).fetchone()
            return row[0] if row else default

    def set_configuracion(self, clave: str, valor: str) -> None:
        with self.db.write_transaction() as conn:
            conn.execute("INSERT OR REPLACE INTO configuracion (clave, valor) VALUES (?, ?);", (clave, str(valor)))

    # --- Auditoría Inmutable ---

    def registrar_auditoria(
        self,
        accion: str,
        entidad: str,
        entidad_id: int | str | None,
        usuario_id: int | None = None,
        usuario_windows: str | None = None,
        datos_anteriores: Any = None,
        datos_nuevos: Any = None,
    ) -> int:
        ahora = self._ahora_iso()
        str_ant = json.dumps(datos_anteriores, ensure_ascii=False, default=str) if datos_anteriores else None
        str_nue = json.dumps(datos_nuevos, ensure_ascii=False, default=str) if datos_nuevos else None
        ent_id_int = int(entidad_id) if (entidad_id is not None and str(entidad_id).isdigit()) else None

        with self.db.write_transaction() as conn:
            cur = conn.execute(
                """INSERT INTO auditoria
                   (usuario_id, usuario_windows, accion, entidad, entidad_id, datos_anteriores, datos_nuevos, creado_en, sync_status)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'pending');""",
                (usuario_id, usuario_windows, accion, entidad, ent_id_int, str_ant, str_nue, ahora),
            )
            aid = cur.lastrowid
            row = dict(conn.execute("SELECT * FROM auditoria WHERE id = ?;", (aid,)).fetchone())
            self._encolar(conn, "auditoria", aid, SyncOperation.UPSERT.value, row)
            return aid

    def listar_auditoria(self, text_filter: str = "") -> list[dict[str, Any]]:
        with self.db.read_connection() as conn:
            like = f"%{text_filter.strip()}%"
            rows = conn.execute(
                """SELECT a.id, a.creado_en, COALESCE(u.username, a.usuario_windows, '') AS usuario,
                          a.accion, a.entidad, a.entidad_id, a.datos_anteriores, a.datos_nuevos
                   FROM auditoria a LEFT JOIN usuarios u ON u.id = a.usuario_id
                   WHERE ? = '' OR COALESCE(u.username, a.usuario_windows, '') LIKE ?
                         OR a.accion LIKE ? OR a.entidad LIKE ?
                   ORDER BY a.id DESC;""",
                (text_filter.strip(), like, like, like),
            ).fetchall()
            return [dict(r) for r in rows]
