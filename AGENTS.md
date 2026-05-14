# AGENTS

## Scope and entrypoint
- Desktop app (Tkinter + `ttkbootstrap`) that reads/writes an Excel workbook via `openpyxl`.
- `main.py` is the sole runtime entrypoint.
- Supporting modules: `excel_store.py` (Excel I/O), `config.py` (constants), `validators.py` (**stubs only** — not yet implemented).
- Windows-specific: `ctypes.windll.user32` for DPI awareness, `root.state('zoomed')`, `.bat` updater, Inno Setup script.

## Dependencies
- Required but **not pinned** anywhere: `pip install openpyxl ttkbootstrap`.

## Run, build, package
- Run locally: `python main.py`
- Build `.exe`: `pyinstaller --onefile --windowed main.py`
- `main.spec` is **gitignored** — do not commit. The CLI command above is the source of truth.
- Built exe lands in `dist/` (gitignored).

## Required sidecar files
- `path_base` (no extension) — contains the target `.xlsx` path; read/written in CWD.
- `theme` (no extension) — stores the selected `ttkbootstrap` theme name; defaults to `flatly` if missing.
- If the Excel file doesn't exist, the app auto-creates it with sheets: `BASE`, `NOVEDADES`, `TipoNovedad`, `Cambio de Turnos`.

## Key data/workflow behaviors
- New records insert at **row 2** in `NOVEDADES` and `Cambio de Turnos` (latest-first display).
- On shared PCs, syncs from disk before save, then computes next ID from latest workbook state.
- Periodic refresh every 60s (`root.after`); skips reload when `mtime` unchanged.
- `BASE` rows cached in memory (`base_rows` + `base_index`) — keep legajo lookups aligned with that cache.
- `PermissionError` on save (workbook open elsewhere) is explicitly caught in UI.
- Observations text and filter variables are **separated per form** (`observaciones_novedades_text`, `observaciones_cambios_text`, etc.) to prevent cross-talk when switching views.
- `Readonly.TEntry` style with bound `<Key>` → `"break"` for programmatic-only Entry fields.
- Windows username saved in auto-created column `USUARIO WINDOWS`.

## Installer / update quirks
- `inno registro novedades.iss` has hardcoded absolute paths (`C:\Users\Luciano\...`) — update before building on another machine.
- `actulizar.bat` copies `main.exe` from `%~dp0` to `C:\Registro de novedades y cambios de turnos TK\main.exe`.

## Repo hygiene
- No test/lint/typecheck config — verification is manual app run + workbook interaction.
- `.gitignore` excludes `*.xlsx`, `build/`, `*.docx`, `main.spec` — avoid committing local data or build artifacts.
