# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_all

bcrypt_datas, bcrypt_binaries, bcrypt_hiddenimports = collect_all("bcrypt")

# Exclusiones de módulos no utilizados para optimizar peso del ejecutable
excludes = [
    "tkinter",
    "ttkbootstrap",
    "PySide6.QtWebEngineCore",
    "PySide6.QtWebEngineWidgets",
    "PySide6.QtWebEngineQuick",
    "PySide6.QtQml",
    "PySide6.QtQuick",
    "PySide6.QtQuickWidgets",
    "PySide6.Qt3DCore",
    "PySide6.Qt3DRender",
    "PySide6.Qt3DInput",
    "PySide6.Qt3DLogic",
    "PySide6.Qt3DAnimation",
    "PySide6.QtSensors",
    "PySide6.QtPositioning",
    "PySide6.QtBluetooth",
    "PySide6.QtNfc",
    "PySide6.QtCharts",
]

datas = bcrypt_datas + [
    ('src/ui/theme/crono_style.qss', 'src/ui/theme'),
    ('src/ui/theme/crono_dark.qss', 'src/ui/theme'),
]

hiddenimports = bcrypt_hiddenimports + [
    "shiboken6",
    "bcrypt",
    "bcrypt._bcrypt",
    "openpyxl",
    "openpyxl.cell._writer",
    "reportlab",
    "docx",
    "httpx",
    "win32com",
    "win32com.client",
    "pythoncom",
    "win32cred",
]

a = Analysis(
    ['main.py'],
    pathex=['.'],
    binaries=bcrypt_binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    noarchive=False,
    optimize=1,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='RENO',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
