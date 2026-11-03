"""Repositorio de persistencia para Novedades y Reemplazos con soporte Outbox."""

from __future__ import annotations

import json
from datetime import datetime
from typing import Any, Sequence

from ...domain.enums import EstadoCobertura, SyncOperation, SyncStatus
from ...domain.models import Novedad, TramoReemplazo
from ...domain.validators import calcular_estado_cobertura_puro, es_turno_ciclo
from ..database import DatabaseManager


class NovedadesRepository:
    """CRUD de novedades y tramos de cobertura con encolado transaccional en sync_outbox."""

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

    def listar_novedades_activas(self) -> list[dict[str, Any]]:
        """Devuelve todas las novedades activas ordenadas por ID descendente."""
        with self.db.read_connection() as conn:
            filas = conn.execute(
                """SELECT id, registrado_en, legajo, apellidos_nombres, especialidad,
                          dotacion, turnos, franco, novedad, fecha_inicio, fecha_fin,
                          referencia_estacion, supervisor, observaciones, usuario_windows, usuario_id
                   FROM novedades WHERE activo = 1 ORDER BY id DESC;"""
            ).fetchall()
            return [dict(f) for f in filas]

    def listar_reemplazos_activos(self) -> list[dict[str, Any]]:
        """Devuelve todos los tramos de reemplazo activos."""
        with self.db.read_connection() as conn:
            filas = conn.execute(
                """SELECT id, novedad_id, legajo_reemplazo, apellidos_nombres_reemplazo,
                          turnos_reemplazo, franco_reemplazo, fecha_inicio, fecha_fin,
                          usuario_windows, usuario_id, creado_en
                   FROM novedad_reemplazos WHERE activo = 1 ORDER BY id ASC;"""
            ).fetchall()
            return [dict(f) for f in filas]

    def obtener_mapa_cobertura(self) -> dict[int, dict[str, str]]:
        """Devuelve {novedad_id: {'estado': ..., 'resumen': ..., 'detalle': ...}}."""
        novedades = self.listar_novedades_activas()
        reemplazos = self.listar_reemplazos_activos()
        reemp_por_nov: dict[int, list[dict[str, Any]]] = {}
        for r in reemplazos:
            reemp_por_nov.setdefault(r["novedad_id"], []).append(r)

        mapa: dict[int, dict[str, str]] = {}
        for n in novedades:
            nid = int(n["id"])
            tramos = reemp_por_nov.get(nid, [])
            mapa[nid] = calcular_estado_cobertura_puro(
                n["fecha_inicio"], n["fecha_fin"], n["turnos"], tramos
            )
        return mapa

    def obtener_novedad_por_id(self, record_id: int) -> dict[str, Any] | None:
        with self.db.read_connection() as conn:
            row = conn.execute("SELECT * FROM novedades WHERE id = ?;", (record_id,)).fetchone()
            return dict(row) if row else None

    def listar_tramos_por_novedad(self, novedad_id: int) -> list[dict[str, Any]]:
        with self.db.read_connection() as conn:
            rows = conn.execute(
                """SELECT * FROM novedad_reemplazos WHERE novedad_id = ? AND activo = 1 ORDER BY id ASC;""",
                (novedad_id,),
            ).fetchall()
            return [dict(r) for r in rows]

    def guardar_novedad(
        self,
        data: dict[str, Any],
        tramos: Sequence[dict[str, Any]] | None = None,
        record_id: int | None = None,
        user_id: int | None = None,
        windows_user: str | None = None,
    ) -> int:
        ahora = self._ahora_iso()
        with self.db.write_transaction() as conn:
            if record_id is None:
                row_max = conn.execute("SELECT COALESCE(MAX(id), 0) + 1 FROM novedades;").fetchone()
                record_id = int(row_max[0])
                conn.execute(
                    """INSERT INTO novedades
                       (id, registrado_en, legajo, apellidos_nombres, especialidad, dotacion,
                        turnos, franco, novedad, fecha_inicio, fecha_fin, referencia_estacion,
                        supervisor, observaciones, usuario_windows, usuario_id, activo, sync_status, updated_at)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, 'pending', ?);""",
                    (
                        record_id,
                        data.get("registrado_en") or ahora,
                        int(data["legajo"]),
                        str(data.get("apellidos_nombres", "")),
                        str(data.get("especialidad", "")),
                        str(data.get("dotacion", "")),
                        str(data.get("turnos", "")),
                        str(data.get("franco", "")),
                        str(data.get("novedad", "")),
                        str(data.get("fecha_inicio", "")),
                        data.get("fecha_fin") or None,
                        str(data.get("referencia_estacion", "")),
                        str(data.get("supervisor", "")),
                        data.get("observaciones") or None,
                        windows_user,
                        user_id,
                        ahora,
                    ),
                )
                op = SyncOperation.UPSERT.value
            else:
                conn.execute(
                    """UPDATE novedades SET
                           legajo = ?, apellidos_nombres = ?, especialidad = ?, dotacion = ?,
                           turnos = ?, franco = ?, novedad = ?, fecha_inicio = ?, fecha_fin = ?,
                           referencia_estacion = ?, supervisor = ?, observaciones = ?,
                           sync_status = 'pending', updated_at = ?
                       WHERE id = ?;""",
                    (
                        int(data["legajo"]),
                        str(data.get("apellidos_nombres", "")),
                        str(data.get("especialidad", "")),
                        str(data.get("dotacion", "")),
                        str(data.get("turnos", "")),
                        str(data.get("franco", "")),
                        str(data.get("novedad", "")),
                        str(data.get("fecha_inicio", "")),
                        data.get("fecha_fin") or None,
                        str(data.get("referencia_estacion", "")),
                        str(data.get("supervisor", "")),
                        data.get("observaciones") or None,
                        ahora,
                        record_id,
                    ),
                )
                op = SyncOperation.UPSERT.value

            # Cargar payload para outbox
            nov_actual = dict(conn.execute("SELECT * FROM novedades WHERE id = ?;", (record_id,)).fetchone())
            self._encolar(conn, "novedades", record_id, op, nov_actual)

            # Si se suministran tramos de reemplazo
            if tramos is not None:
                conn.execute("UPDATE novedad_reemplazos SET activo = 0, sync_status = 'pending' WHERE novedad_id = ?;", (record_id,))
                for t in tramos:
                    cur_r = conn.execute(
                        """INSERT INTO novedad_reemplazos
                           (novedad_id, legajo_reemplazo, apellidos_nombres_reemplazo, turnos_reemplazo,
                            franco_reemplazo, fecha_inicio, fecha_fin, usuario_windows, usuario_id,
                            creado_en, activo, sync_status, updated_at)
                           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, 'pending', ?);""",
                        (
                            record_id,
                            int(t["legajo_reemplazo"]),
                            str(t.get("apellidos_nombres_reemplazo", "")),
                            str(t.get("turnos_reemplazo", "")),
                            str(t.get("franco_reemplazo", "")),
                            str(t["fecha_inicio"]),
                            t.get("fecha_fin") or None,
                            windows_user,
                            user_id,
                            ahora,
                            ahora,
                        ),
                    )
                    r_id = cur_r.lastrowid
                    r_actual = dict(conn.execute("SELECT * FROM novedad_reemplazos WHERE id = ?;", (r_id,)).fetchone())
                    self._encolar(conn, "novedad_reemplazos", r_id, SyncOperation.UPSERT.value, r_actual)

            return record_id

    def soft_delete_novedad(self, record_id: int) -> None:
        ahora = self._ahora_iso()
        with self.db.write_transaction() as conn:
            conn.execute("UPDATE novedades SET activo = 0, sync_status = 'pending', updated_at = ? WHERE id = ?;", (ahora, record_id))
            row = dict(conn.execute("SELECT * FROM novedades WHERE id = ?;", (record_id,)).fetchone())
            self._encolar(conn, "novedades", record_id, SyncOperation.UPSERT.value, row)

    def restaurar_novedad(self, record_id: int) -> None:
        ahora = self._ahora_iso()
        with self.db.write_transaction() as conn:
            conn.execute("UPDATE novedades SET activo = 1, sync_status = 'pending', updated_at = ? WHERE id = ?;", (ahora, record_id))
            row = dict(conn.execute("SELECT * FROM novedades WHERE id = ?;", (record_id,)).fetchone())
            self._encolar(conn, "novedades", record_id, SyncOperation.UPSERT.value, row)

    def borrar_definitivo_novedad(self, record_id: int) -> None:
        with self.db.write_transaction() as conn:
            conn.execute("DELETE FROM novedad_reemplazos WHERE novedad_id = ?;", (record_id,))
            conn.execute("DELETE FROM novedades WHERE id = ?;", (record_id,))
            self._encolar(conn, "novedades", record_id, SyncOperation.DELETE.value, {"id": record_id})
