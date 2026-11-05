"""Pruebas unitarias para el sistema de proxy corporativo y configuración de Supabase."""

import os
from unittest.mock import patch, MagicMock

import pytest
from PySide6.QtWidgets import QApplication

from src.core.config import cargar_configuracion_json, guardar_configuracion_json
from src.core.proxy import (
    AUTH_MANUAL_FIJO,
    AUTH_WINDOWS_SSPI,
    MODO_DIRECTO,
    MODO_MANUAL,
    MODO_WINDOWS_SSO,
    aplicar_configuracion_proxy,
    detectar_proxy_sistema,
    probar_conexion_supabase,
)
from src.data.database import DatabaseManager
from src.data.migrations import init_database_schema
from src.ui.views.dialogs.admin_dialogs import AdminConfigWidget


@pytest.fixture(scope="session")
def qapp():
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


@pytest.fixture
def temp_db(tmp_path):
    db_file = tmp_path / "test_proxy.sqlite"
    db = DatabaseManager(db_file)
    init_database_schema(db)
    return db


def test_detectar_proxy_sistema():
    det = detectar_proxy_sistema({"proxy_mode": MODO_WINDOWS_SSO, "proxy_sso_server": "10.0.0.1:8080"})
    assert "10.0.0.1:8080" in det

    det_direct = detectar_proxy_sistema({"proxy_mode": MODO_DIRECTO})
    assert isinstance(det_direct, str)
    assert len(det_direct) > 0


def test_aplicar_configuracion_proxy():
    # 1. Modo manual
    cfg_manual = {
        "proxy_mode": MODO_MANUAL,
        "proxy_host": "192.168.1.10",
        "proxy_port": 8080,
        "proxy_user": "test_user",
        "proxy_pass": "secret123",
    }
    aplicar_configuracion_proxy(cfg_manual)
    assert "HTTP_PROXY" in os.environ
    assert "192.168.1.10:8080" in os.environ["HTTP_PROXY"]

    # 2. Modo directo (limpia variables)
    cfg_directo = {"proxy_mode": MODO_DIRECTO}
    aplicar_configuracion_proxy(cfg_directo)
    assert "HTTP_PROXY" not in os.environ
    assert "HTTPS_PROXY" not in os.environ


def test_probar_conexion_supabase_sin_credenciales():
    res = probar_conexion_supabase({})
    assert res["ok"] is False
    assert res["status_code"] == 0
    assert "Falta configurar" in res["detalle"]


def test_admin_config_widget_proxy_controls(qapp, temp_db, tmp_path, monkeypatch):
    test_cfg_path = tmp_path / "reno_config.json"
    monkeypatch.setattr("src.core.config.CONFIG_JSON_PATH", test_cfg_path)

    # Configuración inicial guardada
    init_cfg = {
        "supabase_url": "https://testproj.supabase.co",
        "supabase_anon_key": "dummy_anon_key",
        "proxy_mode": MODO_MANUAL,
        "proxy_host": "proxy.corp.local",
        "proxy_port": 3128,
        "proxy_user": "corp_usr",
        "proxy_pass": "corp_pass",
        "proxy_auth": AUTH_MANUAL_FIJO,
    }
    guardar_configuracion_json(init_cfg)

    widget = AdminConfigWidget(temp_db)

    # Verificar que cargó los campos en la interfaz
    assert widget.txt_supa_url.text() == "https://testproj.supabase.co"
    assert widget.cmb_proxy_mode.currentData() == MODO_MANUAL
    assert widget.txt_proxy_host.text() == "proxy.corp.local"
    assert widget.spin_proxy_port.value() == 3128
    assert widget.txt_proxy_user.text() == "corp_usr"
    assert widget.txt_proxy_pass.text() == "corp_pass"
    assert not widget.widget_manual.isHidden()
    assert widget.widget_sso.isHidden()

    # Probar cambio a SSO
    idx_sso = widget.cmb_proxy_mode.findData(MODO_WINDOWS_SSO)
    widget.cmb_proxy_mode.setCurrentIndex(idx_sso)
    assert not widget.widget_sso.isHidden()
    assert widget.widget_manual.isHidden()

    # Probar cambio a Directo
    idx_dir = widget.cmb_proxy_mode.findData(MODO_DIRECTO)
    widget.cmb_proxy_mode.setCurrentIndex(idx_dir)
    assert widget.widget_sso.isHidden()
    assert widget.widget_manual.isHidden()

    # Probar toggle de contraseña
    widget.btn_toggle_pass.click()
    from PySide6.QtWidgets import QLineEdit
    assert widget.txt_proxy_pass.echoMode() == QLineEdit.Normal
    widget.btn_toggle_pass.click()
    assert widget.txt_proxy_pass.echoMode() == QLineEdit.Password
