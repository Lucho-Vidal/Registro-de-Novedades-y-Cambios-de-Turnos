"""Inicialización y migración automática no destructiva de la base SQLite para RENO v2.0.0."""

from __future__ import annotations

import logging
from datetime import datetime

from .database import DatabaseManager

logger = logging.getLogger("reno.data.migrations")

SCHEMA_VERSION_V2 = 2


def init_database_schema(db: DatabaseManager) -> None:
    """Crea o migra todas las tablas de RENO preservando el 100% de los datos existentes."""
    with db.write_transaction() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS schema_metadata (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS empleados (
                id INTEGER PRIMARY KEY,
                legajo INTEGER NOT NULL UNIQUE,
                apellidos_nombres TEXT NOT NULL DEFAULT '',
                especialidad TEXT NOT NULL DEFAULT '',
                dotacion TEXT NOT NULL DEFAULT '',
                turnos TEXT NOT NULL DEFAULT '',
                franco TEXT NOT NULL DEFAULT '',
                franco_horario TEXT NOT NULL DEFAULT '',
                jornada_horas INTEGER NOT NULL DEFAULT 6 CHECK (jornada_horas IN (6, 7)),
                activo INTEGER NOT NULL DEFAULT 1 CHECK (activo IN (0, 1)),
                actualizado_en TEXT NOT NULL,
                sync_status TEXT NOT NULL DEFAULT 'pending',
                created_at TEXT DEFAULT (datetime('now', 'localtime')),
                updated_at TEXT DEFAULT (datetime('now', 'localtime'))
            );

            CREATE TABLE IF NOT EXISTS tipos_novedad (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre TEXT NOT NULL COLLATE NOCASE UNIQUE,
                dias INTEGER NOT NULL DEFAULT 0,
                exportar_descubierto INTEGER NOT NULL DEFAULT 1 CHECK (exportar_descubierto IN (0, 1)),
                activo INTEGER NOT NULL DEFAULT 1 CHECK (activo IN (0, 1)),
                sync_status TEXT NOT NULL DEFAULT 'pending',
                created_at TEXT DEFAULT (datetime('now', 'localtime')),
                updated_at TEXT DEFAULT (datetime('now', 'localtime'))
            );

            CREATE TABLE IF NOT EXISTS dotaciones (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre TEXT NOT NULL COLLATE NOCASE UNIQUE,
                activo INTEGER NOT NULL DEFAULT 1 CHECK (activo IN (0, 1)),
                sync_status TEXT NOT NULL DEFAULT 'pending',
                created_at TEXT DEFAULT (datetime('now', 'localtime')),
                updated_at TEXT DEFAULT (datetime('now', 'localtime'))
            );

            CREATE TABLE IF NOT EXISTS personal_estacion (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre TEXT NOT NULL COLLATE NOCASE UNIQUE,
                dotacion TEXT,
                activo INTEGER NOT NULL DEFAULT 1 CHECK (activo IN (0, 1)),
                sync_status TEXT NOT NULL DEFAULT 'pending',
                created_at TEXT DEFAULT (datetime('now', 'localtime')),
                updated_at TEXT DEFAULT (datetime('now', 'localtime'))
            );

            CREATE TABLE IF NOT EXISTS destinatarios_informe (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre TEXT NOT NULL DEFAULT '',
                email TEXT NOT NULL COLLATE NOCASE UNIQUE,
                activo INTEGER NOT NULL DEFAULT 1 CHECK (activo IN (0, 1)),
                sync_status TEXT NOT NULL DEFAULT 'pending',
                created_at TEXT DEFAULT (datetime('now', 'localtime')),
                updated_at TEXT DEFAULT (datetime('now', 'localtime'))
            );

            CREATE TABLE IF NOT EXISTS roles (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre TEXT NOT NULL COLLATE NOCASE UNIQUE,
                sync_status TEXT NOT NULL DEFAULT 'pending',
                created_at TEXT DEFAULT (datetime('now', 'localtime')),
                updated_at TEXT DEFAULT (datetime('now', 'localtime'))
            );

            CREATE TABLE IF NOT EXISTS permisos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                codigo TEXT NOT NULL COLLATE NOCASE UNIQUE
            );

            CREATE TABLE IF NOT EXISTS usuario_roles (
                usuario_id INTEGER NOT NULL REFERENCES usuarios(id) ON DELETE CASCADE,
                rol_id INTEGER NOT NULL REFERENCES roles(id) ON DELETE CASCADE,
                PRIMARY KEY (usuario_id, rol_id)
            );

            CREATE TABLE IF NOT EXISTS rol_permisos (
                rol_id INTEGER NOT NULL REFERENCES roles(id) ON DELETE CASCADE,
                permiso_id INTEGER NOT NULL REFERENCES permisos(id) ON DELETE CASCADE,
                PRIMARY KEY (rol_id, permiso_id)
            );

            CREATE TABLE IF NOT EXISTS usuarios (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT NOT NULL COLLATE NOCASE UNIQUE,
                nombre TEXT NOT NULL DEFAULT '',
                legajo INTEGER,
                password_hash TEXT NOT NULL,
                activo INTEGER NOT NULL DEFAULT 1 CHECK (activo IN (0, 1)),
                creado_en TEXT NOT NULL,
                ultimo_ingreso TEXT,
                password_cambiado_en TEXT,
                debe_cambiar_clave INTEGER NOT NULL DEFAULT 0,
                sync_status TEXT NOT NULL DEFAULT 'pending',
                created_at TEXT DEFAULT (datetime('now', 'localtime')),
                updated_at TEXT DEFAULT (datetime('now', 'localtime'))
            );

            CREATE TABLE IF NOT EXISTS novedades (
                id INTEGER PRIMARY KEY,
                registrado_en TEXT NOT NULL,
                legajo INTEGER NOT NULL,
                apellidos_nombres TEXT NOT NULL DEFAULT '',
                especialidad TEXT NOT NULL DEFAULT '',
                dotacion TEXT NOT NULL DEFAULT '',
                turnos TEXT NOT NULL DEFAULT '',
                franco TEXT NOT NULL DEFAULT '',
                novedad TEXT NOT NULL,
                fecha_inicio TEXT NOT NULL,
                fecha_fin TEXT,
                referencia_estacion TEXT NOT NULL,
                supervisor TEXT NOT NULL,
                observaciones TEXT,
                usuario_windows TEXT,
                usuario_id INTEGER REFERENCES usuarios(id),
                activo INTEGER NOT NULL DEFAULT 1 CHECK (activo IN (0, 1)),
                sync_status TEXT NOT NULL DEFAULT 'pending',
                created_at TEXT DEFAULT (datetime('now', 'localtime')),
                updated_at TEXT DEFAULT (datetime('now', 'localtime'))
            );

            CREATE TABLE IF NOT EXISTS novedad_reemplazos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                novedad_id INTEGER NOT NULL REFERENCES novedades(id) ON DELETE CASCADE,
                legajo_reemplazo INTEGER NOT NULL,
                apellidos_nombres_reemplazo TEXT NOT NULL DEFAULT '',
                turnos_reemplazo TEXT NOT NULL DEFAULT '',
                franco_reemplazo TEXT NOT NULL DEFAULT '',
                fecha_inicio TEXT NOT NULL,
                fecha_fin TEXT,
                usuario_windows TEXT,
                usuario_id INTEGER REFERENCES usuarios(id),
                creado_en TEXT NOT NULL,
                activo INTEGER NOT NULL DEFAULT 1 CHECK (activo IN (0, 1)),
                sync_status TEXT NOT NULL DEFAULT 'pending',
                created_at TEXT DEFAULT (datetime('now', 'localtime')),
                updated_at TEXT DEFAULT (datetime('now', 'localtime'))
            );

            CREATE TABLE IF NOT EXISTS cambios_turno (
                id INTEGER PRIMARY KEY,
                registrado_en TEXT NOT NULL,
                legajo_1 INTEGER NOT NULL,
                apellidos_nombres_1 TEXT NOT NULL DEFAULT '',
                especialidad_1 TEXT NOT NULL DEFAULT '',
                dotacion_1 TEXT NOT NULL DEFAULT '',
                turnos_1 TEXT NOT NULL DEFAULT '',
                franco_1 TEXT NOT NULL DEFAULT '',
                legajo_2 INTEGER NOT NULL,
                apellidos_nombres_2 TEXT NOT NULL DEFAULT '',
                especialidad_2 TEXT NOT NULL DEFAULT '',
                dotacion_2 TEXT NOT NULL DEFAULT '',
                turnos_2 TEXT NOT NULL DEFAULT '',
                franco_2 TEXT NOT NULL DEFAULT '',
                fecha_cambio TEXT NOT NULL,
                referencia_estacion TEXT NOT NULL,
                supervisor TEXT NOT NULL,
                observaciones TEXT,
                usuario_windows TEXT,
                usuario_id INTEGER REFERENCES usuarios(id),
                activo INTEGER NOT NULL DEFAULT 1 CHECK (activo IN (0, 1)),
                sync_status TEXT NOT NULL DEFAULT 'pending',
                created_at TEXT DEFAULT (datetime('now', 'localtime')),
                updated_at TEXT DEFAULT (datetime('now', 'localtime'))
            );

            CREATE TABLE IF NOT EXISTS ciclo_jornadas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                legajo INTEGER NOT NULL,
                fecha TEXT NOT NULL,
                toma TEXT,
                deja TEXT,
                franco_anticipado INTEGER NOT NULL DEFAULT 0 CHECK (franco_anticipado IN (0, 1)),
                observaciones TEXT,
                actualizado_en TEXT NOT NULL,
                usuario_windows TEXT,
                usuario_id INTEGER REFERENCES usuarios(id),
                sync_status TEXT NOT NULL DEFAULT 'pending',
                created_at TEXT DEFAULT (datetime('now', 'localtime')),
                updated_at TEXT DEFAULT (datetime('now', 'localtime')),
                UNIQUE(legajo, fecha)
            );

            CREATE TABLE IF NOT EXISTS configuracion (
                clave TEXT PRIMARY KEY,
                valor TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS auditoria (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                usuario_id INTEGER REFERENCES usuarios(id),
                usuario_windows TEXT,
                accion TEXT NOT NULL,
                entidad TEXT NOT NULL,
                entidad_id INTEGER,
                datos_anteriores TEXT,
                datos_nuevos TEXT,
                creado_en TEXT NOT NULL,
                sync_status TEXT NOT NULL DEFAULT 'pending'
            );

            -- Cola transaccional Outbox para sincronización asíncrona hacia Supabase
            CREATE TABLE IF NOT EXISTS sync_outbox (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                tabla TEXT NOT NULL,
                registro_pk TEXT NOT NULL,
                operacion TEXT NOT NULL DEFAULT 'UPSERT',
                payload_json TEXT,
                creado_en TEXT NOT NULL,
                intentos INTEGER NOT NULL DEFAULT 0,
                ultimo_error TEXT,
                UNIQUE(tabla, registro_pk)
            );

            CREATE INDEX IF NOT EXISTS idx_novedades_legajo ON novedades(legajo);
            CREATE INDEX IF NOT EXISTS idx_novedades_registrado ON novedades(registrado_en DESC);
            CREATE INDEX IF NOT EXISTS idx_reemplazos_novedad ON novedad_reemplazos(novedad_id, activo);
            CREATE INDEX IF NOT EXISTS idx_reemplazos_legajo ON novedad_reemplazos(legajo_reemplazo, activo);
            CREATE INDEX IF NOT EXISTS idx_ciclo_jornadas_legajo_fecha ON ciclo_jornadas(legajo, fecha);
            CREATE INDEX IF NOT EXISTS idx_cambios_legajo_1 ON cambios_turno(legajo_1);
            CREATE INDEX IF NOT EXISTS idx_cambios_legajo_2 ON cambios_turno(legajo_2);
            CREATE INDEX IF NOT EXISTS idx_cambios_registrado ON cambios_turno(registrado_en DESC);
            CREATE INDEX IF NOT EXISTS idx_sync_outbox_tabla ON sync_outbox(tabla);

            -- Base documental de Itinerarios de Conducción Eléctrica
            CREATE TABLE IF NOT EXISTS turnos_itinerario (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                codigo_turno TEXT UNIQUE NOT NULL,
                turno_base TEXT NOT NULL,
                subindice INTEGER,
                tipo_turno TEXT NOT NULL,
                deposito_base TEXT NOT NULL,
                toma_servicio TEXT,
                deje_servicio TEXT,
                duracion_jornada TEXT,
                tiempo_toma TEXT,
                tiempo_deje TEXT,
                itinerario_vigente TEXT NOT NULL,
                observaciones_turno TEXT,
                activo INTEGER NOT NULL DEFAULT 1,
                sync_status TEXT NOT NULL DEFAULT 'pending',
                created_at TEXT,
                updated_at TEXT
            );

            CREATE TABLE IF NOT EXISTS turnos_vueltas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                turno_id INTEGER NOT NULL REFERENCES turnos_itinerario(id) ON DELETE CASCADE,
                nro_vuelta INTEGER NOT NULL,
                ref TEXT,
                nro_tren TEXT,
                origen TEXT,
                hora_salida TEXT,
                destino TEXT,
                hora_llegada TEXT,
                observaciones_tramo TEXT,
                sync_status TEXT NOT NULL DEFAULT 'pending',
                created_at TEXT,
                updated_at TEXT
            );

            CREATE INDEX IF NOT EXISTS idx_turnos_itin_base ON turnos_itinerario(turno_base);
            CREATE INDEX IF NOT EXISTS idx_turnos_itin_deposito ON turnos_itinerario(deposito_base);
            CREATE INDEX IF NOT EXISTS idx_turnos_vueltas_tren ON turnos_vueltas(nro_tren);
            CREATE INDEX IF NOT EXISTS idx_turnos_vueltas_turno ON turnos_vueltas(turno_id);

            -- Estaciones intermedias y horarios de paso por tren
            CREATE TABLE IF NOT EXISTS tren_paradas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nro_tren TEXT NOT NULL,
                itinerario TEXT NOT NULL DEFAULT '',
                regimen TEXT NOT NULL DEFAULT 'LV',
                orden_parada INTEGER NOT NULL DEFAULT 1,
                estacion TEXT NOT NULL,
                hora_paso TEXT NOT NULL,
                sync_status TEXT NOT NULL DEFAULT 'pending',
                created_at TEXT DEFAULT (datetime('now', 'localtime')),
                updated_at TEXT DEFAULT (datetime('now', 'localtime'))
            );

            CREATE INDEX IF NOT EXISTS idx_tren_paradas_tren ON tren_paradas(nro_tren);
            CREATE INDEX IF NOT EXISTS idx_tren_paradas_itin ON tren_paradas(itinerario);

            -- Módulo de Ordenamientos Operativos
            CREATE TABLE IF NOT EXISTS ordenamientos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nro_orden TEXT UNIQUE NOT NULL,
                fecha_emision TEXT NOT NULL,
                hora_emision TEXT NOT NULL,
                legajo_personal INTEGER NOT NULL,
                nombre_personal TEXT NOT NULL DEFAULT '',
                deposito TEXT NOT NULL DEFAULT 'PC',
                itinerario TEXT NOT NULL DEFAULT '',
                ciclo_o_turno TEXT NOT NULL DEFAULT '',
                regimen TEXT NOT NULL DEFAULT 'LV',
                motivo_orden TEXT NOT NULL DEFAULT '',
                observacion_pie TEXT NOT NULL DEFAULT '',
                ref_pct TEXT NOT NULL DEFAULT '',
                emisor TEXT NOT NULL DEFAULT '',
                usuario_id INTEGER REFERENCES usuarios(id),
                usuario_windows TEXT,
                estado TEXT NOT NULL DEFAULT 'activo',
                activo INTEGER NOT NULL DEFAULT 1 CHECK (activo IN (0, 1)),
                sync_status TEXT NOT NULL DEFAULT 'pending',
                created_at TEXT DEFAULT (datetime('now', 'localtime')),
                updated_at TEXT DEFAULT (datetime('now', 'localtime')),
                deleted_at TEXT
            );

            CREATE TABLE IF NOT EXISTS ordenamiento_tramos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ordenamiento_id INTEGER NOT NULL REFERENCES ordenamientos(id) ON DELETE CASCADE,
                nro_secuencia INTEGER NOT NULL,
                ref TEXT NOT NULL DEFAULT '',
                nro_tren TEXT NOT NULL DEFAULT '',
                origen TEXT NOT NULL DEFAULT '',
                hora_salida TEXT NOT NULL DEFAULT '',
                destino TEXT NOT NULL DEFAULT '',
                hora_llegada TEXT NOT NULL DEFAULT '',
                observaciones TEXT NOT NULL DEFAULT '',
                sync_status TEXT NOT NULL DEFAULT 'pending',
                created_at TEXT DEFAULT (datetime('now', 'localtime')),
                updated_at TEXT DEFAULT (datetime('now', 'localtime'))
            );

            CREATE INDEX IF NOT EXISTS idx_ordenamientos_fecha ON ordenamientos(fecha_emision);
            CREATE INDEX IF NOT EXISTS idx_ordenamientos_legajo ON ordenamientos(legajo_personal);
            CREATE INDEX IF NOT EXISTS idx_ordenamientos_estado ON ordenamientos(estado, activo);
            CREATE INDEX IF NOT EXISTS idx_ord_tramos_ord_id ON ordenamiento_tramos(ordenamiento_id);
            CREATE INDEX IF NOT EXISTS idx_ord_tramos_tren ON ordenamiento_tramos(nro_tren);
            """
        )

        # Migraciones no destructivas de columnas existentes si provienen de v1
        tablas_sincronizables = [
            "empleados",
            "tipos_novedad",
            "dotaciones",
            "personal_estacion",
            "destinatarios_informe",
            "usuarios",
            "roles",
            "novedades",
            "novedad_reemplazos",
            "cambios_turno",
            "ciclo_jornadas",
            "auditoria",
            "turnos_itinerario",
            "turnos_vueltas",
            "tren_paradas",
            "ordenamientos",
            "ordenamiento_tramos",
        ]
        for t in tablas_sincronizables:
            cols = {row[1] for row in conn.execute(f"PRAGMA table_info({t});").fetchall()}
            if "sync_status" not in cols:
                conn.execute(f"ALTER TABLE {t} ADD COLUMN sync_status TEXT NOT NULL DEFAULT 'pending';")
            if "created_at" not in cols and t != "auditoria":
                conn.execute(f"ALTER TABLE {t} ADD COLUMN created_at TEXT;")
                conn.execute(f"UPDATE {t} SET created_at = datetime('now', 'localtime') WHERE created_at IS NULL;")
            if "updated_at" not in cols and t != "auditoria":
                conn.execute(f"ALTER TABLE {t} ADD COLUMN updated_at TEXT;")
                conn.execute(f"UPDATE {t} SET updated_at = datetime('now', 'localtime') WHERE updated_at IS NULL;")

        # Empleados: franco_horario y jornada_horas
        emp_cols = {row[1] for row in conn.execute("PRAGMA table_info(empleados);").fetchall()}
        if "franco_horario" not in emp_cols:
            conn.execute("ALTER TABLE empleados ADD COLUMN franco_horario TEXT NOT NULL DEFAULT '';")
        if "jornada_horas" not in emp_cols:
            conn.execute("ALTER TABLE empleados ADD COLUMN jornada_horas INTEGER NOT NULL DEFAULT 6;")
            conn.execute("UPDATE empleados SET jornada_horas = 7 WHERE LOWER(especialidad) LIKE '%guarda%';")

        # Tipos de novedad: dias y exportar_descubierto
        tn_cols = {row[1] for row in conn.execute("PRAGMA table_info(tipos_novedad);").fetchall()}
        if "dias" not in tn_cols:
            conn.execute("ALTER TABLE tipos_novedad ADD COLUMN dias INTEGER NOT NULL DEFAULT 0;")
        if "exportar_descubierto" not in tn_cols:
            conn.execute("ALTER TABLE tipos_novedad ADD COLUMN exportar_descubierto INTEGER NOT NULL DEFAULT 1;")

        # Parámetros de configuración por defecto
        config_defaults = [
            ("sesion_minutos", "30"),
            ("editar_horas", "24"),
            ("eliminar_horas", "72"),
            ("backup_activo", "1"),
            ("backup_retencion", "10"),
            ("notificaciones_activo", "1"),
            ("toast_duracion", "6"),
            ("fila_alto", "30"),
            ("clave_expiracion_dias", "90"),
            ("correo_dominio", "@trenesargentinos.gob.ar"),
        ]
        for k, v in config_defaults:
            conn.execute("INSERT OR IGNORE INTO configuracion(clave, valor) VALUES (?, ?);", (k, v))

        for nom in ("PC", "LLV", "TY", "LP", "OA", "K5", "RE", "CÑ", "AK"):
            conn.execute("INSERT OR IGNORE INTO dotaciones(nombre) VALUES (?);", (nom,))

        conn.execute("INSERT OR IGNORE INTO tipos_novedad(nombre, dias, exportar_descubierto) VALUES ('Informe', 1, 0);")
        conn.execute("INSERT OR IGNORE INTO tipos_novedad(nombre, dias, exportar_descubierto) VALUES ('Dona sangre', 1, 1);")

        # Poblar tren_paradas automáticamente desde turnos_vueltas si está vacía
        cnt_paradas = conn.execute("SELECT COUNT(*) FROM tren_paradas;").fetchone()[0]
        if cnt_paradas == 0:
            vueltas_rows = conn.execute(
                """
                SELECT v.nro_tren, t.itinerario_vigente, t.tipo_turno,
                       v.origen, v.hora_salida, v.destino, v.hora_llegada
                FROM turnos_vueltas v
                JOIN turnos_itinerario t ON t.id = v.turno_id
                WHERE v.nro_tren IS NOT NULL AND TRIM(v.nro_tren) != ''
                ORDER BY v.nro_tren, v.hora_salida
                """
            ).fetchall()
            vistas: set[tuple[str, str, str, str]] = set()
            orden_por_tren: dict[tuple[str, str], int] = {}
            for r in vueltas_rows:
                nt = str(r[0]).strip()
                itin = str(r[1] or "").strip()
                reg = str(r[2] or "LV").strip()
                orig = str(r[3] or "").strip()
                hs = str(r[4] or "").strip()
                dest = str(r[5] or "").strip()
                hl = str(r[6] or "").strip()
                for est, hr in ((orig, hs), (dest, hl)):
                    if est and hr:
                        key = (nt, itin, est, hr)
                        if key not in vistas:
                            vistas.add(key)
                            ord_k = (nt, itin)
                            orden_por_tren[ord_k] = orden_por_tren.get(ord_k, 0) + 1
                            conn.execute(
                                """
                                INSERT INTO tren_paradas(nro_tren, itinerario, regimen, orden_parada, estacion, hora_paso)
                                VALUES (?, ?, ?, ?, ?, ?);
                                """,
                                (nt, itin, reg, orden_por_tren[ord_k], est, hr),
                            )

        conn.execute(
            "INSERT OR REPLACE INTO schema_metadata(key, value) VALUES (?, ?);",
            ("version", str(SCHEMA_VERSION_V2)),
        )
        logger.info("Esquema SQLite verificado y actualizado a v%s.", SCHEMA_VERSION_V2)
