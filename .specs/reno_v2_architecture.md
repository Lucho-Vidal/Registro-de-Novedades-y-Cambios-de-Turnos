# SDD — RENO v2.0.0 Architecture & Technical Specification

## 1. Visión General
RENO v2.0.0 (Registro de Novedades y Cambios de Turnos) es un sistema híbrido de escritorio para Windows desarrollado con **PySide6-Essentials**, persistencia local **SQLite** (conmutables entre modo WAL de alto rendimiento y modo DELETE con cerrojo `.lock` para carpetas compartidas SMB) y sincronización asíncrona bidireccional hacia **Supabase (PostgreSQL)** mediante tablas dedicadas bajo el prefijo `reno_*`.

---

## 2. Estructura de Paquetes
```text
RENO/
├── .specs/
│   ├── reno_v2_architecture.md
│   └── schema_reno_supabase.sql
├── src/
│   ├── core/
│   │   ├── config.py
│   │   ├── settings.py
│   │   └── logger.py
│   ├── domain/
│   │   ├── models.py
│   │   ├── enums.py
│   │   └── validators.py
│   ├── data/
│   │   ├── database.py
│   │   ├── migrations.py
│   │   ├── repositories/
│   │   │   ├── novedades_repo.py
│   │   │   ├── turnos_repo.py
│   │   │   └── personal_repo.py
│   │   └── sync/
│   │       ├── outbox_queue.py
│   │       └── supabase_sync.py
│   ├── services/
│   │   ├── records_service.py
│   │   ├── auth_service.py
│   │   ├── excel_service.py
│   │   ├── pdf_service.py
│   │   ├── word_service.py
│   │   ├── mail_service.py
│   │   └── backup_service.py
│   └── ui/
│       ├── main_window.py
│       ├── theme/
│       │   └── crono_style.qss
│       ├── models/
│       │   ├── base_table_model.py
│       │   └── multi_filter_proxy.py
│       ├── workers/
│       │   └── task_runner.py
│       └── views/
│           ├── dashboard_view.py
│           ├── novedades_view.py
│           ├── cambios_turno_view.py
│           ├── personal_ciclo_view.py
│           └── dialogs/
├── assets/
├── tests/
├── main.py
├── main.spec
└── requirements.txt
```

---

## 3. Modelo de Datos y Esquema Relacional
Todas las tablas locales y remotas cuentan con columnas de sincronización `sync_status` (`pending` / `synced`) y `updated_at`.
La base local cuenta adicionalmente con `sync_outbox` como cola transaccional.

### Tablas Principales:
1. `empleados` / `reno_empleados`: Ficha de personal activo (legajo, apellidos_nombres, especialidad, dotacion, turnos, franco, franco_horario, jornada_horas, activo).
2. `tipos_novedad` / `reno_tipos_novedad`: Catálogo de novedades (nombre, dias de autocompletado, exportar_descubierto, activo).
3. `dotaciones` / `reno_dotaciones`: Catálogo de bases/dotaciones.
4. `personal_estacion` / `reno_personal_estacion`: Referencias de estaciones con su dotación asociada.
5. `destinatarios_informe` / `reno_destinatarios_informe`: Correos receptores automáticos ante novedades de tipo "Informe".
6. `usuarios` / `reno_usuarios`: Autenticación local/remota con bcrypt.
7. `roles` / `reno_roles`, `permisos` / `reno_permisos`, `usuario_roles`, `rol_permisos`: Matriz RBAC con 31 permisos granulares.
8. `novedades` / `reno_novedades`: Registro operativo de novedades, estado de cobertura, observaciones.
9. `novedad_reemplazos` / `reno_novedad_reemplazos`: Tramos de cobertura realizados por personal de CICLO.
10. `cambios_turno` / `reno_cambios_turno`: Permutas y rotaciones operativas entre dos agentes.
11. `ciclo_jornadas` / `reno_ciclo_jornadas`: Tomas, dejadas y francos anticipados asignados al personal de ciclo.
12. `configuracion` / `reno_configuracion`: Parámetros operativos (`sesion_minutos`, `editar_horas`, `eliminar_horas`, `backup_activo`, etc.).
13. `auditoria` / `reno_auditoria`: Registro inmutable de eventos antes/después en formato JSON.

---

## 4. Reglas de Negocio
- **CCT Descansos Ciclo:** Jornadas habituales de 6 hs exigen 18 hs entre jornadas y 42 hs post-franco. Jornadas habituales de 7 hs exigen 17 hs entre jornadas y 41 hs post-franco. Acumulador semanal reiniciado en cada franco.
- **Ventanas Horarias y Bypass:** Edición permitida dentro de `editar_horas` (24h) y eliminación dentro de `eliminar_horas` (72h). Usuarios con `registros.modificar` se saltan la restricción. Novedades sin `fecha_fin` pueden cerrarse fuera de ventana.
- **Exportación de Gráficos:** Los 4 gráficos del dashboard admiten exportación directa a PNG (alta resolución), reporte PDF ejecutivo con ReportLab y tabla detallada con serie temporal a Excel con OpenPyXL.
