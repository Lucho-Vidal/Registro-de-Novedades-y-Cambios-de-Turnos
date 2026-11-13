"""Pruebas unitarias para el sistema de proxy corporativo y configuración de Supabase."""

import io
import os
import urllib.error
from unittest.mock import patch, MagicMock

import pytest
from PySide6.QtWidgets import QApplication

from src.core import config as config_mod
from src.core.config import cargar_configuracion_app, cargar_configuracion_json, guardar_configuracion_json
from src.core.proxy import (
    AUTH_MANUAL_FIJO,
    AUTH_WINDOWS_SSPI,
    MODO_DIRECTO,
    MODO_MANUAL,
    MODO_WINDOWS_SSO,
    ProxyAuthError,
    _normalizar_proxy_manual,
    aplicar_configuracion_proxy,
    construir_proxy_url_manual,
    descargar_archivo_con_proxy,
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

    # 2. Cambio a Windows SSO sin servidor explícito limpia proxy manual inyectado previamente
    aplicar_configuracion_proxy({"proxy_mode": MODO_WINDOWS_SSO, "proxy_sso_server": ""})
    assert os.environ.get("HTTP_PROXY", "") != cfg_manual["proxy_host"]

    # 3. Modo directo (limpia variables)
    cfg_directo = {"proxy_mode": MODO_DIRECTO}
    aplicar_configuracion_proxy(cfg_directo)
    assert "HTTP_PROXY" not in os.environ
    assert "HTTPS_PROXY" not in os.environ


def test_normalizar_proxy_y_construir_url():
    host, port, user, pwd, auth_mode = _normalizar_proxy_manual(
        {
            "proxy_host": "http://proxy.corp.local:3128/",
            "proxy_user": "DOM\\usr",
            "proxy_pass": "p@ss:123",
            "proxy_auth": AUTH_MANUAL_FIJO,
        }
    )
    assert host == "proxy.corp.local"
    assert port == 3128
    assert auth_mode == AUTH_MANUAL_FIJO

    url_con_creds = construir_proxy_url_manual(
        {
            "proxy_host": "http://proxy.corp.local:3128",
            "proxy_user": "DOM\\usr",
            "proxy_pass": "p@ss:123",
        }
    )
    assert url_con_creds == "http://DOM%5Cusr:p%40ss%3A123@proxy.corp.local:3128"
    assert construir_proxy_url_manual({"proxy_host": ""}) == ""


def test_cargar_configuracion_app_preserva_booleanos_false(tmp_path, monkeypatch):
    cfg_path = tmp_path / "reno_config.json"
    monkeypatch.setattr(config_mod, "CONFIG_JSON_PATH", cfg_path)
    monkeypatch.setattr(config_mod, "APP_DATA_DIR", tmp_path)
    monkeypatch.setattr(config_mod, "BASE_DIR", tmp_path)
    guardar_configuracion_json({"updater_auto_check": False, "proxy_mode": MODO_MANUAL})

    cfg = cargar_configuracion_app()
    assert cfg["updater_auto_check"] is False
    assert cfg["proxy_mode"] == MODO_MANUAL


def test_probar_conexion_supabase_sin_credenciales():
    res = probar_conexion_supabase({})
    assert res["ok"] is False
    assert res["status_code"] == 0
    assert "Falta configurar" in res["detalle"]


def test_descargar_archivo_con_proxy_directo_manual_y_rewind(qapp, tmp_path):
    class _FakeResp:
        def __init__(self, payload: bytes, status: int = 200):
            self.status = status
            self.headers = {"Content-Length": str(len(payload))}
            self._bio = io.BytesIO(payload)

        def read(self, n: int = -1) -> bytes:
            return self._bio.read(n)

        def __enter__(self):
            return self

        def __exit__(self, *args):
            return False

    # 1. Modo Manual incluye Proxy-Authorization Basic en request
    fake_opener = MagicMock()
    captured = {}

    def _open_manual(req, timeout=60):
        captured["proxy_auth"] = req.get_header("Proxy-authorization")
        return _FakeResp(b"INSTALLER_BYTES_OK")

    fake_opener.open.side_effect = _open_manual
    out_manual = tmp_path / "manual.bin"
    progreso = []
    with patch("src.core.proxy.urllib.request.build_opener", return_value=fake_opener):
        with open(out_manual, "wb") as f:
            nbytes = descargar_archivo_con_proxy(
                {
                    "proxy_mode": MODO_MANUAL,
                    "proxy_host": "proxy.corp.local",
                    "proxy_port": 8080,
                    "proxy_user": "alice",
                    "proxy_pass": "wonder",
                },
                "https://ejemplo.supabase.co/storage/v1/object/public/updates/app.exe",
                f,
                on_chunk=lambda cur, tot: progreso.append((cur, tot)),
            )
    assert nbytes == len(b"INSTALLER_BYTES_OK")
    assert out_manual.read_bytes() == b"INSTALLER_BYTES_OK"
    assert captured["proxy_auth"] and captured["proxy_auth"].startswith("Basic ")
    assert progreso[-1] == (len(b"INSTALLER_BYTES_OK"), len(b"INSTALLER_BYTES_OK"))

    # 2. Modo SSO: si QNAM escribe bytes parciales y retorna 407, _reset_stream trunca antes de fallback
    out_sso = tmp_path / "sso.bin"

    def _qnam_407(url, archivo_out, headers, config, **kwargs):
        archivo_out.write(b"HTML_407_ERROR_PAGE")
        return 407, 19

    fake_opener_sso = MagicMock()
    fake_opener_sso.open.return_value = _FakeResp(b"CLEAN_SSO_BINARY")

    with patch("src.core.proxy._descargar_stream_via_qnam", side_effect=_qnam_407), \
         patch("src.core.proxy.sys.platform", "linux"), \
         patch("src.core.proxy.urllib.request.build_opener", return_value=fake_opener_sso):
        with open(out_sso, "wb") as f:
            n_sso = descargar_archivo_con_proxy(
                {"proxy_mode": MODO_WINDOWS_SSO, "proxy_sso_server": "10.1.2.3:8080"},
                "https://ejemplo.supabase.co/storage/v1/object/public/updates/app.exe",
                f,
            )
    assert n_sso == len(b"CLEAN_SSO_BINARY")
    assert out_sso.read_bytes() == b"CLEAN_SSO_BINARY"

    # 3. Modo Directo lanzando ProxyAuthError en HTTP 407
    fake_opener_407 = MagicMock()
    fake_opener_407.open.side_effect = urllib.error.HTTPError(
        "https://ejemplo.com", 407, "Proxy Auth Required", {}, io.BytesIO(b"")
    )
    out_407 = tmp_path / "fail407.bin"
    with patch("src.core.proxy.urllib.request.build_opener", return_value=fake_opener_407):
        with open(out_407, "wb") as f:
            with pytest.raises(ProxyAuthError):
                descargar_archivo_con_proxy({"proxy_mode": MODO_DIRECTO}, "https://ejemplo.com/app.exe", f)


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

