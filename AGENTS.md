# AGENTS

## Scope and entrypoint
- Desktop app (Tkinter + `ttkbootstrap`) with SQLite-first persistence (via `sqlite3`) and Excel import/export via `openpyxl`.
- `main.py` is the sole runtime entrypoint.
- Data layer: `sqlite_store.py` (SQLite I/O + `SQLiteSheetAdapter`), `database_bootstrap.py` (database location + one-time migration from XLSX).
- Excel bridge: `excel_migration.py` (import XLSX → SQLite), `excel_exporter.py` (export SQLite → XLSX), `excel_store.py` (only `get_windows_user` is used by the app).
- Auth/roles: `auth.py` (`AuthService`, bcrypt), `login_view.py` (login + first-run admin creation).
- UI modules: `forms.py`, `tables.py`, `admin_views.py`, `records_service.py`, `outlook_mailer.py`, `backups.py`.
- `validators.py` is implemented (validation for novedades/cambios forms) — `pytest test_validators.py` is the test suite.
- Windows-specific: `ctypes.windll.user32` for DPI awareness, `root.state('zoomed')`, `.bat` updater, Inno Setup script.

## Dependencies
- `requirements.txt`: `openpyxl`, `ttkbootstrap==1.10.1`, `bcrypt`, `pywin32`, `pyinstaller` (not pinned upstream).

## Run, build, package
- Run locally: `python main.py`
- Build `.exe`: `pyinstaller --onefile --windowed main.py` (or `RENO.spec`; specs are gitignored — the CLI command is the source of truth).
- Built exe lands in `dist/` (gitignored; `dist/main.exe` remains tracked for historical reasons).

## Required sidecar files
- `path_base` (no extension) — contains the target path: either the `.xlsx` (initial migration source) or the `.sqlite` database; read/written in CWD.
- `theme` (no extension) — stores the selected `ttkbootstrap` theme name; defaults to `flatly` if missing.

## Key data/workflow behaviors
- Store is a single SQLite file (`*.sqlite`); if it doesn't exist and the configured `.xlsx` exists, it auto-migrates once (`migrate_workbook`).
- New records are inserted and displayed latest-first (`ORDER BY id DESC`) in `NOVEDADES` and `Cambio de Turnos` views.
- `BASE`/empleados rows cached in memory (`base_rows` + `base_index`) — keep legajo lookups aligned with that cache (`db_store.get_base_rows()`).
- Empleados CRUD lives in `Administración > Empleados` (`empleados.administrar`, solo administradores) with soft delete (`activo` toggle via `cambiar_estado_empleado`); after any change call `actualizar_cache_base()` + `cargarDotaciones()` to keep the legajo modal and dotación cache in sync. `empleados.importar` still covers Excel import via Archivo.
- Personal de estación, tipos de novedad y dotaciones can be imported from the 1st column of the 1st sheet of an XLSX via `migrate_personal_estacion_sheet` / `migrate_tipos_novedad_sheet` / `migrate_dotaciones_sheet` (all delegate to `_migrar_catalogo`: skips a header row matching known prefixes; reactivates existing names with `ON CONFLICT(nombre) DO UPDATE`; optional `clear_existing` soft-clears all first) — button "Importar desde Excel" in `Administración > Personal de estación` (permiso `personalEstacion.editar`), `Tipos de novedad` (permiso `novedades.editar`) and `Dotaciones` (permiso `dotaciones.administrar`). The buttons share `AdminViews._importar_desde_excel`.
- Employee form (`Administración > Empleados`): Legajo (número), Apellidos y nombres (obligatorio), Especialidad/Dotación/Franco are required readonly Comboboxes (`ESPECIALIDADES_EMPLEADO`, `self.app.dotaciones`, `DIAS_SEMANA`); legacy values not in a list are appended to the options so they stay editable.
- Periodic refresh every 60s (`root.after`); reloads views from SQLite.
- Auth: first run has no users → login shows a "Legajo" field and a "Crear administrador inicial" button (hidden/destroyed once a user exists). Permissions are checked per-action via `tiene_permiso`/`requerir_permiso`.
- Import/export are split per table: `novedades.importar`/`novedades.exportar` and `cambios_turno.importar`/`cambios_turno.exportar` (used in the Archivo menu, `importar_excel_operativo` and `exportar_excel`; the export dialog checks the permission of the selected table). `empleados.importar`/`usuarios.administrar` still gate the BASE import.
- Session control: 30s activity timeout checked against configured `sesion_minutos` (default 30 min).
- Edit/delete windows: `novedades`/`cambios_turno` use soft delete via `activo` (0 = deleted, not listed, recoverable). `editar_horas` (default 24) and `eliminar_horas` (default 72) limit editing/deleting by age of `registrado_en`; unparseable legacy dates are blocked. Recovery + permanent delete live in `Administración > Registros eliminados` (`registros.recuperar`); windows are set in `Administración > Tiempos de edición` (`sesion.configurar`).
- Observations text and filter variables are **separated per form** (`observaciones_novedades_text`, `observaciones_cambios_text`, etc.) to prevent cross-talk when switching views.
- `Readonly.TEntry` style with bound `<Key>` → `"break"` for programmatic-only Entry fields.
- Windows username stored in the `usuario_windows` column on save/audit.

## Installer / update quirks
- `RENO.iss` is the single Inno Setup script (output installer `Instalador\RENO.exe`); it has hardcoded absolute paths (`C:\Users\Lucia\...`) — update before building on another machine.
- `actulizar.bat` copies `main.exe` from `%~dp0` to `C:\Registro de novedades y cambios de turnos TK\main.exe`.

## Backups
- `backups.py`: copies of the SQLite DB stored in `backups\` next to the database as `backup_YYYYMMDD_HHMMSS.sqlite`, using the SQLite backup API.
- Auto-backup runs once per day at startup (`backup_activo`, default 1); retention = last N copies (`backup_retencion`, default 10). Manual create/restore/delete/retention live in `Administración > Copias de seguridad` (`backup.gestionar`).
- `registrado_en` is stored ISO `%Y-%m-%d %H:%M:%S` for new records; exports filter dates in Python accepting both ISO and legacy `DD/MM/YYYY`.

## Repo hygiene
- Verification: `pytest test_validators.py` + manual app run. No lint/typecheck configured.
- `.gitignore` excludes `*.xlsx`, `*.sqlite*`, `build/`, `dist/`, `main.spec`, `RENO.spec`, `theme`, `path_base` — avoid committing local data or build artifacts. `dist/main.exe` and `Instalador/Registro de novedades.exe` remain tracked by decision.