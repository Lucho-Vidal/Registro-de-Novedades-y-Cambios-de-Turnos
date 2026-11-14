"""Pruebas automatizadas para la resolución, carga visual y empaquetado del icono de RENO."""

from __future__ import annotations

import os
import shutil
import struct
import sys
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtGui import QIcon, QPixmap
from PySide6.QtWidgets import QApplication, QLabel

from src.core.config import BASE_DIR, obtener_ruta_icono, resource_path
from src.core.theme import (
    aplicar_icono_aplicacion,
    cargar_icono_aplicacion,
    cargar_pixmap_icono,
)
from src.data.database import DatabaseManager
from src.data.migrations import init_database_schema
from src.domain.enums import PERMISOS_BASE
from src.services.auth_service import AuthService
from src.services.backup_service import BackupService
from src.services.excel_service import ExcelService
from src.services.pdf_service import PDFService
from src.services.records_service import RecordsService
from src.services.updater_service import InfoActualizacion
from src.services.word_service import WordService
from src.ui.main_window import MainWindow
from src.ui.views.dialogs.acerca_dialog import AcercaDialog
from src.ui.views.dialogs.actualizacion_dialog import ActualizacionDialog
from src.ui.views.dialogs.login_dialog import (
    CambioPasswordObligatorioDialog,
    LoginDialog,
)


def _get_app() -> QApplication:
    return QApplication.instance() or QApplication([])


def test_icon_assets_exist_and_multi_resolution():
    """Verifica que assets/reno_icon.png y el ICO multi-resolución existan y sean válidos."""
    png_path = BASE_DIR / "assets" / "reno_icon.png"
    ico_path = BASE_DIR / "assets" / "reno_icon.ico"

    assert png_path.is_file()
    assert ico_path.is_file()
    assert png_path.stat().st_size > 1000
    assert ico_path.stat().st_size > 1000

    # Verificar encabezado binario Windows ICO (ICONDIR) con múltiples resoluciones
    ico_bytes = ico_path.read_bytes()
    reserved, img_type, count = struct.unpack("<HHH", ico_bytes[:6])
    assert reserved == 0
    assert img_type == 1
    assert count >= 6

    resolutions = set()
    for i in range(count):
        w, h = struct.unpack("<BB", ico_bytes[6 + i * 16 : 8 + i * 16])
        resolutions.add((w or 256, h or 256))
    for expected in ((16, 16), (32, 32), (48, 48), (256, 256)):
        assert expected in resolutions


def test_obtener_ruta_icono_dev_and_frozen(tmp_path: Path, monkeypatch):
    """Verifica resolución de icono en modo desarrollo, override externo y PyInstaller _MEIPASS."""
    # Desarrollo: prioriza PNG para Qt por defecto, ICO con prefer_ico=True
    assert obtener_ruta_icono(prefer_ico=False) == BASE_DIR / "assets" / "reno_icon.png"
    assert obtener_ruta_icono(prefer_ico=True) == BASE_DIR / "assets" / "reno_icon.ico"

    # Modo frozen (_MEIPASS): simular ejecutable aislado
    meipass_dir = tmp_path / "_meipass"
    (meipass_dir / "assets").mkdir(parents=True)
    shutil.copy2(BASE_DIR / "assets" / "reno_icon.png", meipass_dir / "assets" / "reno_icon.png")
    shutil.copy2(BASE_DIR / "assets" / "reno_icon.ico", meipass_dir / "assets" / "reno_icon.ico")

    exe_dir = tmp_path / "dist"
    exe_dir.mkdir(parents=True)
    fake_exe = exe_dir / "RENO.exe"
    fake_exe.write_bytes(b"")

    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "_MEIPASS", str(meipass_dir), raising=False)
    monkeypatch.setattr(sys, "executable", str(fake_exe))

    assert resource_path("assets/reno_icon.png") == meipass_dir / "assets" / "reno_icon.png"
    assert obtener_ruta_icono(prefer_ico=False) == meipass_dir / "assets" / "reno_icon.png"
    assert obtener_ruta_icono(prefer_ico=True) == meipass_dir / "assets" / "reno_icon.ico"

    # Si existe un icono personalizado junto al ejecutable, debe priorizarse
    custom_ico = exe_dir / "assets" / "reno_icon.ico"
    custom_ico.parent.mkdir(parents=True)
    shutil.copy2(BASE_DIR / "assets" / "reno_icon.ico", custom_ico)
    assert obtener_ruta_icono(prefer_ico=True) == custom_ico


def test_qt_icon_and_pixmap_loading():
    """Verifica que Qt cargue tanto PNG como ICO y genere pixmaps válidos."""
    _get_app()

    icono_png = cargar_icono_aplicacion(prefer_ico=False)
    assert not icono_png.isNull()
    assert not icono_png.pixmap(32, 32).isNull()

    icono_ico = cargar_icono_aplicacion(prefer_ico=True)
    assert not icono_ico.isNull()
    assert not icono_ico.pixmap(32, 32).isNull()

    pix = cargar_pixmap_icono(48)
    assert pix is not None
    assert not pix.isNull()
    assert pix.width() == 48 and pix.height() == 48

    # Directo desde QIcon
    direct = QIcon(str(BASE_DIR / "assets" / "reno_icon.png"))
    assert not direct.isNull()


def test_windows_and_dialogs_have_non_null_icon(tmp_path: Path):
    """Verifica que MainWindow, LoginDialog, AcercaDialog y ActualizacionDialog apliquen el icono."""
    app = _get_app()
    assert aplicar_icono_aplicacion(app) is True
    assert not app.windowIcon().isNull()

    db = DatabaseManager(tmp_path / "icon_ui.sqlite", journal_mode="WAL")
    init_database_schema(db)
    auth = AuthService(db)
    auth.inicializar_permisos(PERMISOS_BASE)

    login_dlg = LoginDialog(auth)
    assert not login_dlg.windowIcon().isNull()
    # Verificar que el logotipo gráfico de RENO se renderizó en el header del login
    pix_labels = [lbl for lbl in login_dlg.findChildren(QLabel) if lbl.pixmap() is not None and not lbl.pixmap().isNull()]
    assert len(pix_labels) >= 1
    login_dlg.close()
    login_dlg.deleteLater()

    pwd_dlg = CambioPasswordObligatorioDialog(1, auth)
    assert not pwd_dlg.windowIcon().isNull()
    pwd_dlg.close()
    pwd_dlg.deleteLater()

    acerca_dlg = AcercaDialog()
    assert not acerca_dlg.windowIcon().isNull()
    acerca_pix = [lbl for lbl in acerca_dlg.findChildren(QLabel) if lbl.pixmap() is not None and not lbl.pixmap().isNull()]
    assert len(acerca_pix) >= 1
    acerca_dlg.close()
    acerca_dlg.deleteLater()

    info = InfoActualizacion(
        version="2.0.1",
        fecha="2026-11-14",
        titulo="Actualización de prueba",
        novedades=["Prueba de icono"],
        instalador_nombre="RENO_Setup_v2.0.1.exe",
        instalador_url="https://example.com/setup.exe",
        sha256="0" * 64,
        tamano_bytes=1024,
        obligatoria=False,
    )
    upd_dlg = ActualizacionDialog(info)
    assert not upd_dlg.windowIcon().isNull()
    upd_dlg.close()
    upd_dlg.deleteLater()

    auth.crear_administrador_inicial("admin", "Admin123!", legajo=1001)
    usuario = auth.autenticar("admin", "Admin123!")
    win = MainWindow(
        db=db,
        records_service=RecordsService(db),
        auth_service=auth,
        excel_service=ExcelService(db),
        pdf_service=PDFService(),
        word_service=WordService(),
        backup_service=BackupService(db),
        usuario_actual=usuario,
        supabase_sync=None,
    )
    assert not win.windowIcon().isNull()
    win.close()
    win.deleteLater()
    app.processEvents()


def test_packaging_files_reference_reno_icon():
    """Verifica que main.spec y RENO.iss incluyan la carpeta assets y el archivo reno_icon.ico."""
    spec_text = (BASE_DIR / "main.spec").read_text(encoding="utf-8")
    assert "('assets', 'assets')" in spec_text
    assert "reno_icon.ico" in spec_text
    assert "icon=app_icon" in spec_text

    iss_text = (BASE_DIR / "RENO.iss").read_text(encoding="utf-8")
    assert r"SetupIconFile=assets\reno_icon.ico" in iss_text
    assert r"UninstallDisplayIcon={app}\{#MyAppExeName}" in iss_text
