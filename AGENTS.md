# AGENTS - RENO v2.0.0

## Scope and entrypoint
- Desktop app migrated to **PySide6-Essentials** adopting the visual and architectural standard of **CRONO-PCT**.
- Hybrid persistence: Local SQLite-first (Write-Ahead Logging / DELETE + `.lock`) + Outbox transactional queue synchronizing with Supabase (PostgreSQL 15+ REST, prefix `reno_*`).
- `main.py` is the sole runtime entrypoint.
- Data layer: `src/data/database.py` (`DatabaseManager`), `src/data/migrations.py`, `src/data/repositories/`, `src/data/sync/` (`OutboxQueue`, `SupabaseSyncClient`).
- Domain layer: `src/domain/models.py`, `src/domain/enums.py`, `src/domain/validators.py` (0% external dependencies).
- Services layer: `src/services/records_service.py`, `src/services/auth_service.py`, `src/services/excel_service.py`, `src/services/pdf_service.py` (ReportLab), `src/services/word_service.py` (python-docx), `src/services/backup_service.py`, `src/services/mail_service.py`.
- UI layer: `src/ui/main_window.py`, `src/ui/theme/crono_style.qss`, `src/ui/models/` (`BaseTableModel`, `MultiFilterProxyModel`), `src/ui/views/` (`DashboardView`, `NovedadesView`, `CambiosTurnoView`, `PersonalCicloView`), `src/ui/views/dialogs/`, `src/ui/workers/task_runner.py`.
- Automated test suite: `pytest tests/` (`tests/test_domain_validators.py`, `tests/test_services_and_exports.py`, `tests/test_ui_models.py`).

## Dependencies
- `requirements.txt`: `PySide6-Essentials>=6.5.0`, `openpyxl>=3.1.2`, `bcrypt>=4.0.0`, `pywin32>=306`, `reportlab>=4.0.0`, `python-docx>=1.0.0`, `httpx>=0.25.0`.

## Run, build, package
- Run locally: `python main.py`
- Run test suite: `pytest tests/`
- Build `.exe`: `pyinstaller main.spec` (configured with module exclusions for PySide6 to keep binary lightweight <50MB).

## Key data/workflow behaviors (100% Functional Parity)
- Store is a single SQLite file (`*.sqlite`); auto-migrates schema and creates `sync_outbox` transactional queue.
- New records are inserted and displayed latest-first (`ORDER BY id DESC`) in `NOVEDADES` and `Cambio de Turnos` views.
- Empleados CRUD lives in `Administración > Empleados` (`empleados.administrar`) with soft delete (`activo` toggle).
- Tipos de novedad cuenta con `dias` (duración en días para autocompletar `fecha_fin = fecha_inicio + (dias - 1)`) y `exportar_descubierto` (define si entra en el reporte de turnos sin cubrir).
- Reemplazos / Personal de Ciclo: la tabla `novedad_reemplazos` almacena tramos de cobertura sin solapamiento interno ni con otras novedades activas para el mismo personal de ciclo. Los candidatos se filtran por igual dotación, especialidad y turno `CICLO%`.
- Calendario mensual de ciclo y cálculo de descansos CCT: calcula hora disponible respetando el CCT según `jornada_horas` (para 6 hs: 18 hs entre jornadas y 42 hs post-franco; para 7 hs: 17 hs entre jornadas y 41 hs post-franco) y acumula `HS SEMANA`.
- Modales maximizados: Las ventanas modales principales de administración, detalle de registro, búsqueda de legajos y calendario de ciclo se abren automáticamente maximizadas (`showMaximized()`).
- Dashboard con exportación detallada: visualiza KPIs, 4 gráficos interactivos `QPainter` y permite exportación en imagen PNG, informe ejecutivo PDF (ReportLab) y desglose detallado Excel multihistorias (OpenPyXL).
- Sincronización híbrida: barra de estado con indicador visual (Verde=Sincronizado, Amarillo=Pendiente/Sincronizando, Gris/Rojo=Offline).
- Permission `registros.modificar`: otorga edición y eliminación sin importar la antigüedad de `registrado_en`, omitiendo las ventanas de tiempo `editar_horas`/`eliminar_horas`.
- Auto-backup diario al iniciar con retención configurable (`backup_retencion`, default 10 copias).
