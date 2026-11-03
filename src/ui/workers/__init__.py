"""Módulo de trabajadores en segundo plano para tareas asíncronas sin congelar la UI."""

from src.ui.workers.task_runner import BackgroundTask, WorkerSignals, run_async

__all__ = ["BackgroundTask", "WorkerSignals", "run_async"]
