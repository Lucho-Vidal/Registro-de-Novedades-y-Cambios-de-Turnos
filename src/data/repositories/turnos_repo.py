"""Repositorio de persistencia para Cambios de Turno y Jornadas de Ciclo."""

from __future__ import annotations

import json
from datetime import date, datetime
from typing import Any

from ...domain.enums import SyncOperation
from ..database import DatabaseManager


class TurnosRepository:
    """CRUD de cambios de turno y ciclo_jornadas con outbox transaccional."""

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

    def listar_cambios_activos(self) -> list[dict[str, Any]]:
        with self.db.read_connection() as conn:
            filas = conn.execute(
                """SELECT id, registrado_en, legajo_1, apellidos_nombres_1, especialidad_1,
                          dotacion_1, turnos_1, franco_1, legajo_2, apellidos_nombres_2,
                          especialidad_2, dotacion_2, turnos_2, franco_2, fecha_cambio,
                          referencia_estacion, supervisor, observaciones, usuario_windows, usuario_id
                   FROM cambios_turno WHERE activo = 1 ORDER BY id DESC;"""
            ).fetchall()
            return [dict(f) for f in filas]

    def obtener_cambio_por_id(self, record_id: int) -> dict[str, Any] | None:
        with self.db.read_connection() as conn:
            row = conn.execute("SELECT * FROM cambios_turno WHERE id = ?;", (record_id,)).fetchone()
            return dict(row) if row else None

    def guardar_cambio_turno(
        self,
        data: dict[str, Any],
        record_id: int | None = None,
        user_id: int | None = None,
        windows_user: str | None = None,
    ) -> int:
        ahora = self._ahora_iso()
        fields = (
            "legajo_1", "apellidos_nombres_1", "especialidad_1", "dotacion_1", "turnos_1", "franco_1",
            "legajo_2", "apellidos_nombres_2", "especialidad_2", "dotacion_2", "turnos_2", "franco_2",
            "fecha_cambio", "referencia_estacion", "supervisor", "observaciones",
        )
        with self.db.write_transaction() as conn:
            if record_id is None:
                row_max = conn.execute("SELECT COALESCE(MAX(id), 0) + 1 FROM cambios_turno;").fetchone()
                record_id = int(row_max[0])
                conn.execute(
                    """INSERT INTO cambios_turno
                       (id, registrado_en, legajo_1, apellidos_nombres_1, especialidad_1,
                        dotacion_1, turnos_1, franco_1, legajo_2, apellidos_nombres_2,
                        especialidad_2, dotacion_2, turnos_2, franco_2, fecha_cambio,
                        referencia_estacion, supervisor, observaciones, usuario_windows, usuario_id,
                        activo, sync_status, updated_at)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, 'pending', ?);""",
                    (record_id, data.get("registrado_en") or ahora)
                    + tuple(data.get(f) for f in fields)
                    + (windows_user, user_id, ahora),
                )
            else:
                set_clause = ", ".join(f"{f} = ?" for f in fields)
                conn.execute(
                    f"""UPDATE cambios_turno SET {set_clause}, sync_status = 'pending', updated_at = ? WHERE id = ?;""",
                    tuple(data.get(f) for f in fields) + (ahora, record_id),
                )

            row = dict(conn.execute("SELECT * FROM cambios_turno WHERE id = ?;", (record_id,)).fetchone())
            self._encolar(conn, "cambios_turno", record_id, SyncOperation.UPSERT.value, row)
            return record_id

    def soft_delete_cambio(self, record_id: int) -> None:
        ahora = self._ahora_iso()
        with self.db.write_transaction() as conn:
            conn.execute("UPDATE cambios_turno SET activo = 0, sync_status = 'pending', updated_at = ? WHERE id = ?;", (ahora, record_id))
            row = dict(conn.execute("SELECT * FROM cambios_turno WHERE id = ?;", (record_id,)).fetchone())
            self._encolar(conn, "cambios_turno", record_id, SyncOperation.UPSERT.value, row)

    def restaurar_cambio(self, record_id: int) -> None:
        ahora = self._ahora_iso()
        with self.db.write_transaction() as conn:
            conn.execute("UPDATE cambios_turno SET activo = 1, sync_status = 'pending', updated_at = ? WHERE id = ?;", (ahora, record_id))
            row = dict(conn.execute("SELECT * FROM cambios_turno WHERE id = ?;", (record_id,)).fetchone())
            self._encolar(conn, "cambios_turno", record_id, SyncOperation.UPSERT.value, row)

    def borrar_definitivo_cambio(self, record_id: int) -> None:
        with self.db.write_transaction() as conn:
            conn.execute("DELETE FROM cambios_turno WHERE id = ?;", (record_id,))
            self._encolar(conn, "cambios_turno", record_id, SyncOperation.DELETE.value, {"id": record_id})

    def guardar_jornada_ciclo(
        self,
        legajo: int,
        fecha_iso: str,
        toma: str | None,
        deja: str | None,
        franco_anticipado: int,
        observaciones: str | None,
        user_id: int | None = None,
        windows_user: str | None = None,
    ) -> int:
        ahora = self._ahora_iso()
        with self.db.write_transaction() as conn:
            conn.execute(
                """INSERT INTO ciclo_jornadas
                   (legajo, fecha, toma, deja, franco_anticipado, observaciones, actualizado_en,
                    usuario_windows, usuario_id, sync_status, updated_at)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'pending', ?)
                   ON CONFLICT(legajo, fecha) DO UPDATE SET
                       toma = excluded.toma,
                       deja = excluded.deja,
                       franco_anticipado = excluded.franco_anticipado,
                       observaciones = excluded.observaciones,
                       actualizado_en = excluded.actualizado_en,
                       usuario_windows = excluded.usuario_windows,
                       usuario_id = excluded.usuario_id,
                       sync_status = 'pending',
                       updated_at = excluded.updated_at;""",
                (
                    legajo,
                    fecha_iso,
                    toma or None,
                    deja or None,
                    1 if franco_anticipado else 0,
                    observaciones or None,
                    ahora,
                    windows_user,
                    user_id,
                    ahora,
                ),
            )
            row = dict(
                conn.execute(
                    "SELECT * FROM ciclo_jornadas WHERE legajo = ? AND fecha = ?;", (legajo, fecha_iso)
                ).fetchone()
            )
            pk = f"{legajo}_{fecha_iso}"
            self._encolar(conn, "ciclo_jornadas", pk, SyncOperation.UPSERT.value, row)
            return int(row["id"])

    def obtener_jornadas_ciclo_rango(self, legajo: int, fecha_desde_iso: str, fecha_hasta_iso: str) -> dict[str, dict[str, Any]]:
        with self.db.read_connection() as conn:
            rows = conn.execute(
                """SELECT id, legajo, fecha, toma, deja, franco_anticipado, observaciones
                   FROM ciclo_jornadas
                   WHERE legajo = ? AND fecha BETWEEN ? AND ?;""",
                (legajo, fecha_desde_iso, fecha_hasta_iso),
            ).fetchall()
            return {r["fecha"]: dict(r) for r in rows}
