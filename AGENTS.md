# AGENTS

## Scope and entrypoint
- This repo is a single-file desktop app: `main.py` is the only Python source and runtime entrypoint.
- The app is a Tkinter + `ttkbootstrap` GUI that reads/writes an Excel workbook via `openpyxl`.
- It is Windows-oriented (`ctypes.windll.user32`, `root.state('zoomed')`, `.bat` updater, Inno Setup script).

## Run and package (verified)
- Run locally from source: `python main.py`.
- Build the executable (per `README.md`): `pyinstaller --onefile --windowed main.py`.
- `main.spec` exists, but the documented build command above is the source of truth in this repo.

## Required runtime sidecar files
- `path_base` (no extension) stores the target `.xlsx` path; the app reads/writes this file in the current working directory.
- `theme` stores the selected `ttkbootstrap` theme name; if missing, app defaults to `flatly`.
- If the Excel file pointed to by `path_base` does not exist, the app auto-creates it with sheets: `BASE`, `NOVEDADES`, `TipoNovedad`, `Cambio de Turnos`.

## Data/workflow behaviors worth knowing
- New records are inserted at row 2 in `NOVEDADES` and `Cambio de Turnos` (latest-first display pattern).
- In shared-PC usage, the app syncs from disk before save and then computes next ID from the latest workbook state.
- Periodic refresh is Tkinter-safe (`root.after` every 60s), and reload is skipped when file `mtime` is unchanged.
- `BASE` rows are cached in memory (`base_rows` + `base_index`); keep legajo lookups/index-based flows aligned with that cache.
- Saving fails with `PermissionError` when the workbook is open elsewhere (explicitly handled in UI).

## Installer/update quirks
- `inno registro novedades.iss` uses hardcoded absolute `C:\Users\Luciano\workspace\...` paths for inputs/output; update these before building an installer on another machine.
- `actulizar.bat` copies `main.exe` to `C:\Registro de novedades y cambios de turnos TK\main.exe`.

## Repo-specific hygiene
- There is no test/lint/typecheck config in this repo; verification is manual app run + workbook interaction.
- `.gitignore` excludes `*.xlsx`, `build/`, and `*.docx`; avoid committing local workbook data.
