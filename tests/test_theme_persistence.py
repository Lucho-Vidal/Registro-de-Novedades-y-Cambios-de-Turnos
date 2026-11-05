"""Pruebas unitarias de persistencia inmediata del tema visual en RENO v2.0."""

import pytest
from PySide6.QtWidgets import QApplication

from src.core.config import cargar_configuracion_json
from src.core.settings import AppSettings
from src.core.theme import (
    TEMA_CLARO,
    TEMA_OSCURO,
    aplicar_tema_visual,
    guardar_tema_visual,
    obtener_tema_visual,
)
from src.data.database import DatabaseManager
from src.data.migrations import init_database_schema


@pytest.fixture(scope="session")
def qapp():
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


@pytest.fixture
def temp_db(tmp_path):
    db_file = tmp_path / "test_theme.sqlite"
    db = DatabaseManager(db_file)
    init_database_schema(db)
    return db


def test_persistencia_inmediata_tema_visual(qapp, temp_db, tmp_path, monkeypatch):
    test_cfg_path = tmp_path / "reno_config.json"
    monkeypatch.setattr("src.core.config.CONFIG_JSON_PATH", test_cfg_path)

    # 1. Guardar modo oscuro
    guardar_tema_visual(TEMA_OSCURO, db=temp_db)

    # Verificar en SQLite
    with temp_db.read_connection() as conn:
        val_db = conn.execute("SELECT valor FROM configuracion WHERE clave = 'tema_visual'").fetchone()[0]
        assert val_db == TEMA_OSCURO

    # Verificar en reno_config.json
    cfg = cargar_configuracion_json()
    assert cfg.get("tema_visual") == TEMA_OSCURO

    # Verificar en AppSettings
    assert AppSettings().get_valor("tema_visual") == TEMA_OSCURO

    # Verificar con obtener_tema_visual
    assert obtener_tema_visual(temp_db) == TEMA_OSCURO

    # 2. Aplicar tema oscuro
    aplicado = aplicar_tema_visual(db=temp_db)
    assert aplicado == TEMA_OSCURO
    assert len(qapp.styleSheet()) > 0
    assert "#1E1E2E" in qapp.styleSheet()

    # 3. Guardar modo claro
    guardar_tema_visual(TEMA_CLARO, db=temp_db)

    with temp_db.read_connection() as conn:
        val_db2 = conn.execute("SELECT valor FROM configuracion WHERE clave = 'tema_visual'").fetchone()[0]
        assert val_db2 == TEMA_CLARO

    assert cargar_configuracion_json().get("tema_visual") == TEMA_CLARO
    assert AppSettings().get_valor("tema_visual") == TEMA_CLARO
    assert obtener_tema_visual(temp_db) == TEMA_CLARO

    # 4. Aplicar tema claro
    aplicado_claro = aplicar_tema_visual(db=temp_db)
    assert aplicado_claro == TEMA_CLARO
    assert "#F4F6F9" in qapp.styleSheet() or "#FFFFFF" in qapp.styleSheet()
