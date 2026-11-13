"""Motor de sincronización híbrida hacia Supabase bajo tablas reno_*."""

from __future__ import annotations

import base64
import json
import logging
import os
import re
import sys
import threading
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime
from typing import Any

from ...core.config import cargar_configuracion_app
from ...core.proxy import (
    OfflineError,
    ProxyAuthError,
    ejecutar_request_supabase,
)
from ..database import DatabaseManager
from .outbox_queue import OutboxQueue

logger = logging.getLogger("reno.data.sync")


def sanitizar_supabase_url(url: str | None) -> str:
    if not url:
        return ""
    u = str(url).strip().rstrip("/")
    m = re.search(r"supabase\.com/dashboard/project/([a-zA-Z0-9_\-]+)", u, re.IGNORECASE)
    if m:
        return f"https://{m.group(1)}.supabase.co"
    u = re.sub(r"/(rest|api)/v\d+.*$", "", u, flags=re.IGNORECASE).rstrip("/")
    if u.startswith("http://"):
        u = "https://" + u[len("http://"):]
    elif not u.startswith("https://"):
        u = "https://" + u
    return u.rstrip("/")


def ejecutar_request_http(
    config: dict[str, Any],
    url: str,
    method: str = "GET",
    data: bytes | None = None,
    headers: dict[str, str] | None = None,
    timeout_sec: int = 15,
) -> tuple[int, bytes, dict[str, str]]:
    """Ejecuta una petición HTTP respetando la configuración de proxy corporativo multi-tier."""
    return ejecutar_request_supabase(
        config=config,
        url=url,
        method=method,
        data=data,
        headers=headers,
        timeout_sec=timeout_sec,
    )


class SupabaseSyncEngine:
    """Motor de push (lotes de outbox) y pull (resolución de colisiones)."""

    def __init__(
        self,
        db: DatabaseManager,
        outbox: OutboxQueue | None = None,
        supabase_url: str | None = None,
        supabase_key: str | None = None,
    ) -> None:
        self.db = db
        self.outbox = outbox if isinstance(outbox, OutboxQueue) else OutboxQueue(db)
        self.supabase_url = supabase_url
        self.supabase_key = supabase_key
        self._lock = threading.Lock()

    def sincronizar_completo(self) -> tuple[int, int]:
        """Ejecuta ciclo completo de sincronización (push + pull)."""
        return self.sincronizar(solo_push=False, solo_pull=False)

    def sincronizar(self, solo_push: bool = False, solo_pull: bool = False) -> tuple[int, int]:
        if not self._lock.acquire(blocking=False):
            return 0, 0
        try:
            return self._ejecutar_sync(solo_push=solo_push, solo_pull=solo_pull)
        finally:
            self._lock.release()

    def _ejecutar_sync(self, solo_push: bool = False, solo_pull: bool = False) -> tuple[int, int]:
        cfg = cargar_configuracion_app()
        url_base = sanitizar_supabase_url(self.supabase_url or cfg.get("supabase_url"))
        anon_key = (self.supabase_key or cfg.get("supabase_anon_key") or "").strip()
        if not url_base or not anon_key:
            raise ValueError("Credenciales de Supabase no configuradas.")

        from ...core.proxy import construir_headers_supabase

        headers_base = construir_headers_supabase(
            anon_key,
            {
                "Content-Type": "application/json",
                "Prefer": "return=minimal",
            },
        )

        enviados = 0
        recibidos = 0

        # --- 1. PUSH DE OUTBOX ---
        if not solo_pull:
            pendientes = self.outbox.obtener_pendientes(limite=200)
            if pendientes:
                # Agrupar por tabla
                por_tabla: dict[str, list[tuple[int, dict[str, Any]]]] = {}
                for item in pendientes:
                    try:
                        d = json.loads(item.payload_json) if item.payload_json else {}
                        por_tabla.setdefault(item.tabla, []).append((item.id, d))
                    except Exception:
                        continue

                # Mapeo tabla local -> tabla remota prefijada y su primary key
                pks_map = {
                    "dotaciones": ("reno_dotaciones", "nombre"),
                    "tipos_novedad": ("reno_tipos_novedad", "nombre"),
                    "personal_estacion": ("reno_personal_estacion", "nombre"),
                    "destinatarios_informe": ("reno_destinatarios_informe", "email"),
                    "destinatarios_ordenamiento": ("reno_destinatarios_ordenamiento", "email"),
                    "roles": ("reno_roles", "nombre"),
                    "usuarios": ("reno_usuarios", "username"),
                    "empleados": ("reno_empleados", "legajo"),
                    "novedades": ("reno_novedades", "id"),
                    "novedad_reemplazos": ("reno_novedad_reemplazos", "id"),
                    "cambios_turno": ("reno_cambios_turno", "id"),
                    "ciclo_jornadas": ("reno_ciclo_jornadas", "legajo,fecha"),
                    "tren_paradas": ("reno_tren_paradas", "id"),
                    "ordenamientos": ("reno_ordenamientos", "id"),
                    "ordenamiento_tramos": ("reno_ordenamiento_tramos", "id"),
                    "auditoria": ("reno_auditoria", "id"),
                }

                orden_tablas = [
                    "dotaciones", "tipos_novedad", "personal_estacion", "destinatarios_informe",
                    "destinatarios_ordenamiento",
                    "roles", "usuarios", "empleados", "novedades", "novedad_reemplazos",
                    "cambios_turno", "ciclo_jornadas", "tren_paradas", "ordenamientos",
                    "ordenamiento_tramos", "auditoria"
                ]

                ids_a_eliminar = []
                for t in orden_tablas:
                    if t not in por_tabla:
                        continue
                    remota_info = pks_map.get(t)
                    if not remota_info:
                        continue
                    tabla_remota, pk_col = remota_info
                    items_lote = por_tabla[t]
                    lote_payload = []
                    for outbox_id, p in items_lote:
                        p_limpio = dict(p)
                        p_limpio.pop("sync_status", None)
                        lote_payload.append(p_limpio)

                    endpoint = f"{url_base}/rest/v1/{tabla_remota}?on_conflict={pk_col}"
                    data_bytes = json.dumps(lote_payload, ensure_ascii=False).encode("utf-8")
                    req_hdrs = dict(headers_base)
                    req_hdrs["Prefer"] = "resolution=merge-duplicates,return=minimal"

                    try:
                        status, body, _ = ejecutar_request_http(
                            config=cfg,
                            url=endpoint,
                            method="POST",
                            data=data_bytes,
                            headers=req_hdrs,
                            timeout_sec=15,
                        )
                        if 200 <= status < 300:
                            for outbox_id, _ in items_lote:
                                ids_a_eliminar.append(outbox_id)
                            enviados += len(items_lote)
                        else:
                            msg_err = body.decode("utf-8", "ignore")
                            for outbox_id, _ in items_lote:
                                self.outbox.registrar_fallo(outbox_id, msg_err)
                    except Exception as ex:
                        for outbox_id, _ in items_lote:
                            self.outbox.registrar_fallo(outbox_id, str(ex))
                        raise

                self.outbox.eliminar_procesados(ids_a_eliminar)

        # --- 2. PULL INCREMENTAL + ANTI-COLISIÓN ---
        if not solo_push:
            tablas_pull = [
                ("reno_dotaciones", "dotaciones", "nombre"),
                ("reno_tipos_novedad", "tipos_novedad", "nombre"),
                ("reno_personal_estacion", "personal_estacion", "nombre"),
                ("reno_destinatarios_informe", "destinatarios_informe", "email"),
                ("reno_destinatarios_ordenamiento", "destinatarios_ordenamiento", "email"),
                ("reno_empleados", "empleados", "legajo"),
                ("reno_novedades", "novedades", "id"),
                ("reno_novedad_reemplazos", "novedad_reemplazos", "id"),
                ("reno_cambios_turno", "cambios_turno", "id"),
                ("reno_ciclo_jornadas", "ciclo_jornadas", "legajo,fecha"),
            ]

            with self.db.write_transaction() as conn:
                for remota, local, pk_fld in tablas_pull:
                    endpoint_pull = f"{url_base}/rest/v1/{remota}?select=*"
                    try:
                        status, body, _ = ejecutar_request_http(
                            config=cfg,
                            url=endpoint_pull,
                            method="GET",
                            headers=headers_base,
                            timeout_sec=15,
                        )
                        if 200 <= status < 300:
                            remotos = json.loads(body.decode("utf-8")) if body else []
                            for r in remotos:
                                recibidos += 1
                                # Anti-colisión en novedades y cambios por ID
                                if local == "novedades":
                                    row_loc = conn.execute(
                                        "SELECT legajo, fecha_inicio, sync_status FROM novedades WHERE id = ?;",
                                        (r["id"],),
                                    ).fetchone()
                                    if row_loc and row_loc["sync_status"] == "pending":
                                        if str(row_loc["legajo"]) != str(r["legajo"]) or str(row_loc["fecha_inicio"]) != str(r["fecha_inicio"]):
                                            max_id = conn.execute("SELECT COALESCE(MAX(id), 0) FROM novedades;").fetchone()[0]
                                            nuevo_id = max_id + 1
                                            conn.execute("UPDATE novedades SET id = ? WHERE id = ?;", (nuevo_id, r["id"]))
                                            conn.execute("UPDATE novedad_reemplazos SET novedad_id = ? WHERE novedad_id = ?;", (nuevo_id, r["id"]))
                                            conn.execute("UPDATE sync_outbox SET registro_pk = ? WHERE tabla = 'novedades' AND registro_pk = ?;", (str(nuevo_id), str(r["id"])))
                                        else:
                                            continue

                                    conn.execute(
                                        """INSERT OR REPLACE INTO novedades
                                           (id, registrado_en, legajo, apellidos_nombres, especialidad, dotacion,
                                            turnos, franco, novedad, fecha_inicio, fecha_fin, referencia_estacion,
                                            supervisor, observaciones, usuario_windows, usuario_id, activo, sync_status, updated_at)
                                           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'synced', ?);""",
                                        (
                                            r["id"], r["registrado_en"], r["legajo"], r.get("apellidos_nombres", ""),
                                            r.get("especialidad", ""), r.get("dotacion", ""), r.get("turnos", ""),
                                            r.get("franco", ""), r["novedad"], r["fecha_inicio"], r.get("fecha_fin"),
                                            r.get("referencia_estacion", ""), r.get("supervisor", ""), r.get("observaciones"),
                                            r.get("usuario_windows"), r.get("usuario_id"), 1 if r.get("activo", True) else 0,
                                            r.get("updated_at") or datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                        ),
                                    )
                    except Exception as pull_ex:
                        logger.debug("Error en pull de %s: %s", remota, pull_ex)

        return enviados, recibidos


SupabaseSyncClient = SupabaseSyncEngine

