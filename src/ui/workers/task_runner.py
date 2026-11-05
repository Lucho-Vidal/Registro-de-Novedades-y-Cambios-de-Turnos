"""Ejecución concurrente de tareas pesadas en segundo plano mediante QThreadPool y QRunnable.

Permite realizar sincronizaciones con Supabase, exportaciones Excel/PDF y cargas masivas
garantizando 60 FPS fluidos y sin congelar la ventana principal de la aplicación.
"""

import sys
import traceback
from typing import Any, Callable, Optional
from PySide6.QtCore import QObject, QRunnable, QThreadPool, Signal

from src.core.logger import get_logger

logger = get_logger("TaskRunner")


class WorkerSignals(QObject):
    """Señales emitidas por un BackgroundTask."""
    started = Signal()
    progress = Signal(int, str)  # porcentaje (0-100), mensaje descriptivo
    finished = Signal(object)   # resultado retornado por la función
    error = Signal(str, str)    # mensaje_error, traceback_completo


class BackgroundTask(QRunnable):
    """Ejecutor QRunnable que corre una función con argumentos en un hilo del pool."""

    def __init__(self, fn: Callable[..., Any], *args: Any, **kwargs: Any):
        super().__init__()
        self.fn = fn
        self.args = args
        self.kwargs = kwargs
        self.signals = WorkerSignals()
        self.setAutoDelete(True)

    def run(self) -> None:
        try:
            self.signals.started.emit()
        except Exception:
            pass

        try:
            # Si la función acepta un callback de progreso, se lo inyectamos
            if "progress_callback" in self.kwargs or "progress_callback" in self.fn.__code__.co_varnames:
                self.kwargs["progress_callback"] = lambda pct, msg="": self._emit_progress(int(pct), str(msg))
            res = self.fn(*self.args, **self.kwargs)
            try:
                self.signals.finished.emit(res)
            except Exception:
                pass
        except Exception as e:
            tb = traceback.format_exc()
            logger.error(f"Error en tarea asíncrona: {e}\n{tb}")
            try:
                self.signals.error.emit(str(e), tb)
            except Exception:
                pass

    def _emit_progress(self, pct: int, msg: str) -> None:
        try:
            self.signals.progress.emit(pct, msg)
        except Exception:
            pass


_ACTIVE_TASKS: set[BackgroundTask] = set()


def _cleanup_task(task: BackgroundTask) -> None:
    _ACTIVE_TASKS.discard(task)


def run_async(
    fn: Callable[..., Any],
    *args: Any,
    on_success: Optional[Callable[[Any], None]] = None,
    on_error: Optional[Callable[[str, str], None]] = None,
    on_progress: Optional[Callable[[int, str], None]] = None,
    on_started: Optional[Callable[[], None]] = None,
    pool: Optional[QThreadPool] = None,
    **kwargs: Any,
) -> BackgroundTask:
    """Envía una función para ejecutarse de inmediato en el thread pool con callbacks conectados."""
    task = BackgroundTask(fn, *args, **kwargs)
    _ACTIVE_TASKS.add(task)

    if on_started:
        task.signals.started.connect(on_started)
    if on_progress:
        task.signals.progress.connect(on_progress)

    def _wrapped_success(res: Any) -> None:
        try:
            if on_success:
                on_success(res)
        finally:
            _cleanup_task(task)

    def _wrapped_error(err: str, tb: str) -> None:
        try:
            if on_error:
                on_error(err, tb)
        finally:
            _cleanup_task(task)

    task.signals.finished.connect(_wrapped_success)
    task.signals.error.connect(_wrapped_error)

    target_pool = pool or QThreadPool.globalInstance()
    target_pool.start(task)
    return task
