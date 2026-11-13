"""Gestión transaccional de la cola sync_outbox."""

from __future__ import annotations

import json
from typing import Any

from ...domain.models import SyncQueueItem
from ..database import DatabaseManager


class OutboxQueue:
    """Administra los elementos pendientes de sincronizar hacia Supabase."""

    def __init__(self, db: DatabaseManager) -> None:
        self.db = db

    def contar_pendientes(self) -> int:
        with self.db.read_connection() as conn:
            row = conn.execute("SELECT COUNT(*) FROM sync_outbox;").fetchone()
            return int(row[0]) if row else 0

    def obtener_pendientes(self, limite: int = 200) -> list[SyncQueueItem]:
        with self.db.read_connection() as conn:
            rows = conn.execute(
                """SELECT id, tabla, registro_pk, operacion, payload_json, creado_en, intentos, ultimo_error
                   FROM sync_outbox ORDER BY id ASC LIMIT ?;""",
                (limite,),
            ).fetchall()
            return [
                SyncQueueItem(
                    id=r["id"],
                    tabla=r["tabla"],
                    registro_pk=r["registro_pk"],
                    operacion=r["operacion"],
                    payload_json=r["payload_json"],
                    creado_en=r["creado_en"],
                    intentos=r["intentos"],
                    ultimo_error=r["ultimo_error"],
                )
                for r in rows
            ]

    def eliminar_procesados(self, ids: list[int]) -> None:
        if not ids:
            return
        with self.db.write_transaction() as conn:
            placeholders = ",".join("?" for _ in ids)
            conn.execute(f"DELETE FROM sync_outbox WHERE id IN ({placeholders});", ids)

    def registrar_fallo(self, item_id: int, mensaje_error: str) -> None:
        with self.db.write_transaction() as conn:
            conn.execute(
                """UPDATE sync_outbox
                   SET intentos = intentos + 1, ultimo_error = ?
                   WHERE id = ?;""",
                (mensaje_error[:500], item_id),
            )

    def forzar_resincronizacion_total(self) -> int:
        """Marca todas las filas de las tablas principales como pendientes en sync_outbox."""
        tablas = [
            ("dotaciones", "nombre"),
            ("tipos_novedad", "nombre"),
            ("personal_estacion", "nombre"),
            ("destinatarios_informe", "email"),
            ("destinatarios_ordenamiento", "email"),
            ("roles", "nombre"),
            ("usuarios", "username"),
            ("empleados", "legajo"),
            ("novedades", "id"),
            ("novedad_reemplazos", "id"),
            ("cambios_turno", "id"),
            ("ciclo_jornadas", "id"),
            ("tren_paradas", "id"),
            ("ordenamientos", "id"),
            ("ordenamiento_tramos", "id"),
            ("auditoria", "id"),
        ]
        total = 0
        with self.db.write_transaction() as conn:
            for t, pk in tablas:
                try:
                    filas = conn.execute(f"SELECT * FROM {t};").fetchall()
                except Exception:
                    continue
                for f in filas:
                    d = dict(f)
                    val_pk = str(d.get(pk, d.get("id", "")))
                    if not val_pk or val_pk == "None":
                        continue
                    conn.execute(
                        """INSERT INTO sync_outbox (tabla, registro_pk, operacion, payload_json, creado_en)
                           VALUES (?, ?, 'UPSERT', ?, datetime('now', 'localtime'))
                           ON CONFLICT(tabla, registro_pk) DO UPDATE SET
                               operacion = 'UPSERT',
                               payload_json = excluded.payload_json,
                               intentos = 0,
                               ultimo_error = NULL;""",
                        (t, val_pk, json.dumps(d, ensure_ascii=False, default=str)),
                    )
                    total += 1
                try:
                    conn.execute(f"UPDATE {t} SET sync_status = 'pending';")
                except Exception:
                    pass
        return total
