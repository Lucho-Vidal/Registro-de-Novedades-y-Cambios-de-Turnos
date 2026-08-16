"""Ubicación y creación de la base de datos configurada para la aplicación."""

from pathlib import Path

from excel_migration import migrate_workbook
from sqlite_store import SQLiteStore


def configured_paths():
    try:
        excel_file = Path(Path("path_base").read_text(encoding="utf-8").strip())
    except (FileNotFoundError, OSError):
        excel_file = Path("RENO.xlsx")
    if excel_file.suffix.lower() == ".sqlite":
        return excel_file.with_suffix(".xlsx"), excel_file
    return excel_file, excel_file.with_suffix(".sqlite")


def open_configured_store():
    excel_file, database_file = configured_paths()
    store = SQLiteStore(database_file)
    database_exists = database_file.exists()
    store.initialize()
    if not database_exists and excel_file.exists() and excel_file.suffix.lower() == ".xlsx":
        migrate_workbook(str(excel_file), store)
    return excel_file, store
