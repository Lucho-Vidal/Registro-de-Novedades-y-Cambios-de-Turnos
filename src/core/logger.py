"""Sistema de logging rotativo y captura de excepciones globales para RENO."""

from __future__ import annotations

import logging
import os
import sys
import tempfile
import traceback
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Any

from PySide6.QtCore import QtMsgType, qInstallMessageHandler

logger = logging.getLogger("reno")

_LOGGING_INICIALIZADO = False


def get_logger(name: str = "reno") -> logging.Logger:
    """Devuelve un logger configurado con el prefijo 'reno.'."""
    if name == "reno" or name.startswith("reno."):
        return logging.getLogger(name)
    return logging.getLogger(f"reno.{name}")


def obtener_directorio_logs() -> Path:
    """Devuelve el directorio seguro y escribible para almacenar los archivos de registro."""
    candidatos = []
    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        candidatos.append(Path(local_app_data) / "RENO" / "logs")

    app_data = os.environ.get("APPDATA")
    if app_data:
        candidatos.append(Path(app_data) / "RENO" / "logs")

    base_dir = Path(__file__).resolve().parent.parent.parent / "logs"
    candidatos.append(base_dir)
    candidatos.append(Path(tempfile.gettempdir()) / "RENO" / "logs")

    for p in candidatos:
        try:
            p.mkdir(parents=True, exist_ok=True)
            test_file = p / ".write_test"
            test_file.write_text("ok", encoding="utf-8")
            test_file.unlink(missing_ok=True)
            return p
        except Exception:
            continue

    fallback = Path(tempfile.gettempdir()) / "RENO" / "logs"
    fallback.mkdir(parents=True, exist_ok=True)
    return fallback


def obtener_ruta_log() -> Path:
    """Devuelve la ruta absoluta al archivo reno.log."""
    return obtener_directorio_logs() / "reno.log"


class StreamToLogger:
    """Redirige las escrituras de un flujo al logger."""

    def __init__(self, target_logger: logging.Logger, log_level: int, original_stream: Any = None) -> None:
        self.logger = target_logger
        self.log_level = log_level
        self.original_stream = original_stream
        self._buffer: list[str] = []
        self._in_write: bool = False

    def write(self, buf: str) -> None:
        if self._in_write:
            return
        self._in_write = True
        try:
            if self.original_stream:
                try:
                    self.original_stream.write(buf)
                    self.original_stream.flush()
                except Exception:
                    pass

            if not buf:
                return

            lineas = buf.splitlines(keepends=True)
            for linea in lineas:
                if linea.endswith(("\r", "\n")):
                    texto_completo = "".join(self._buffer) + linea.rstrip("\r\n")
                    self._buffer.clear()
                    texto_limpio = texto_completo.strip()
                    if texto_limpio:
                        try:
                            self.logger.log(self.log_level, texto_limpio)
                        except Exception:
                            pass
                else:
                    self._buffer.append(linea)
        finally:
            self._in_write = False

    def flush(self) -> None:
        if self._in_write:
            return
        self._in_write = True
        try:
            if self._buffer:
                texto = "".join(self._buffer).strip()
                self._buffer.clear()
                if texto:
                    try:
                        self.logger.log(self.log_level, texto)
                    except Exception:
                        pass
            if self.original_stream:
                try:
                    self.original_stream.flush()
                except Exception:
                    pass
        finally:
            self._in_write = False


def _manejador_mensajes_qt(mode: QtMsgType, context: Any, message: str) -> None:
    msg = str(message or "").strip()
    if not msg:
        return
    origen = f" [{context.file}:{context.line}]" if context and getattr(context, "file", None) else ""
    qt_logger = logging.getLogger("reno.qt")
    if mode == QtMsgType.QtDebugMsg:
        qt_logger.debug("%s%s", msg, origen)
    elif mode == QtMsgType.QtInfoMsg:
        qt_logger.info("%s%s", msg, origen)
    elif mode == QtMsgType.QtWarningMsg:
        qt_logger.warning("%s%s", msg, origen)
    elif mode == QtMsgType.QtCriticalMsg:
        qt_logger.error("%s%s", msg, origen)
    elif mode == QtMsgType.QtFatalMsg:
        qt_logger.critical("%s%s", msg, origen)
    else:
        qt_logger.warning("%s%s", msg, origen)


def _manejador_excepciones_no_controladas(exc_type: type, exc_value: BaseException, exc_traceback: Any) -> None:
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return
    logger.critical(
        "CRITICAL ERROR - Excepción no controlada en la aplicación:\n%s",
        "".join(traceback.format_exception(exc_type, exc_value, exc_traceback)),
    )


def inicializar_logging(nivel: int = logging.INFO) -> Path:
    """Configura el sistema de logging rotativo y captura global de excepciones."""
    global _LOGGING_INICIALIZADO
    log_file = obtener_ruta_log()
    if _LOGGING_INICIALIZADO:
        return log_file

    root_logger = logging.getLogger()
    root_logger.setLevel(nivel)
    formato = logging.Formatter(
        "[%(asctime)s] [%(levelname)-7s] [%(name)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    try:
        file_handler = RotatingFileHandler(
            str(log_file),
            maxBytes=10 * 1024 * 1024,
            backupCount=5,
            encoding="utf-8",
        )
        file_handler.setLevel(nivel)
        file_handler.setFormatter(formato)
        root_logger.addHandler(file_handler)
    except Exception as e:
        sys.stderr.write(f"No se pudo crear el archivo de log: {e}\n")

    if not getattr(sys, "frozen", False):
        tiene_console = any(
            isinstance(h, logging.StreamHandler) and not isinstance(h, RotatingFileHandler)
            for h in root_logger.handlers
        )
        if not tiene_console and (sys.stdout is not None or sys.stderr is not None):
            try:
                ch = logging.StreamHandler(sys.stdout or sys.stderr)
                ch.setLevel(nivel)
                ch.setFormatter(formato)
                root_logger.addHandler(ch)
            except Exception:
                pass

    stdout_logger = logging.getLogger("reno.stdout")
    stderr_logger = logging.getLogger("reno.stderr")
    sys.stdout = StreamToLogger(stdout_logger, logging.INFO, sys.__stdout__ if sys.stdout is not None else None)
    sys.stderr = StreamToLogger(stderr_logger, logging.ERROR, sys.__stderr__ if sys.stderr is not None else None)

    sys.excepthook = _manejador_excepciones_no_controladas

    try:
        qInstallMessageHandler(_manejador_mensajes_qt)
    except Exception as e:
        logger.debug("No se pudo instalar manejador Qt: %s", e)

    _LOGGING_INICIALIZADO = True
    logger.info("=== INICIO DE SESIÓN RENO v2.0.0 - Logs en %s ===", log_file)
    return log_file
