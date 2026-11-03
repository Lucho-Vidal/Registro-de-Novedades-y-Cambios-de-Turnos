"""Modelos y dataclasses inmutables y tipadas para RENO v2.0.0."""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from datetime import date, datetime
from typing import Any

from .enums import EstadoCobertura, EstadoDiaCiclo, SyncStatus


@dataclass(frozen=True)
class Empleado:
    id: int | None
    legajo: int
    apellidos_nombres: str
    especialidad: str = ""
    dotacion: str = ""
    turnos: str = ""
    franco: str = ""
    franco_horario: str = ""
    jornada_horas: int = 6
    activo: bool = True
    actualizado_en: str = ""
    sync_status: SyncStatus = SyncStatus.PENDING

    def to_dict(self) -> dict[str, Any]:
        d = asdict(self)
        d["activo"] = 1 if self.activo else 0
        d["sync_status"] = self.sync_status.value
        return d


@dataclass(frozen=True)
class TramoReemplazo:
    id: int | None
    novedad_id: int
    legajo_reemplazo: int
    apellidos_nombres_reemplazo: str
    turnos_reemplazo: str = ""
    franco_reemplazo: str = ""
    fecha_inicio: str = ""
    fecha_fin: str | None = None
    usuario_windows: str | None = None
    usuario_id: int | None = None
    creado_en: str = ""
    activo: bool = True
    sync_status: SyncStatus = SyncStatus.PENDING

    def to_dict(self) -> dict[str, Any]:
        d = asdict(self)
        d["activo"] = 1 if self.activo else 0
        d["sync_status"] = self.sync_status.value
        return d


@dataclass(frozen=True)
class Novedad:
    id: int | None
    registrado_en: str
    legajo: int
    apellidos_nombres: str
    especialidad: str
    dotacion: str
    turnos: str
    franco: str
    novedad: str
    fecha_inicio: str
    fecha_fin: str | None
    referencia_estacion: str
    supervisor: str
    observaciones: str | None = None
    usuario_windows: str | None = None
    usuario_id: int | None = None
    activo: bool = True
    reemplazos: tuple[TramoReemplazo, ...] = field(default_factory=tuple)
    cobertura_estado: EstadoCobertura = EstadoCobertura.NO
    cobertura_resumen: str = "No"
    cobertura_detalle: str = ""
    sync_status: SyncStatus = SyncStatus.PENDING

    def to_dict(self) -> dict[str, Any]:
        d = asdict(self)
        d["activo"] = 1 if self.activo else 0
        d["cobertura_estado"] = self.cobertura_estado.value
        d["sync_status"] = self.sync_status.value
        d["reemplazos"] = [r.to_dict() for r in self.reemplazos]
        return d


@dataclass(frozen=True)
class CambioTurno:
    id: int | None
    registrado_en: str
    legajo_1: int
    apellidos_nombres_1: str
    especialidad_1: str
    dotacion_1: str
    turnos_1: str
    franco_1: str
    legajo_2: int
    apellidos_nombres_2: str
    especialidad_2: str
    dotacion_2: str
    turnos_2: str
    franco_2: str
    fecha_cambio: str
    referencia_estacion: str
    supervisor: str
    observaciones: str | None = None
    usuario_windows: str | None = None
    usuario_id: int | None = None
    activo: bool = True
    sync_status: SyncStatus = SyncStatus.PENDING

    def to_dict(self) -> dict[str, Any]:
        d = asdict(self)
        d["activo"] = 1 if self.activo else 0
        d["sync_status"] = self.sync_status.value
        return d


@dataclass(frozen=True)
class CicloJornada:
    id: int | None
    legajo: int
    fecha: str  # YYYY-MM-DD
    toma: str | None = None
    deja: str | None = None
    franco_anticipado: bool = False
    observaciones: str | None = None
    actualizado_en: str = ""
    usuario_windows: str | None = None
    usuario_id: int | None = None
    sync_status: SyncStatus = SyncStatus.PENDING

    def to_dict(self) -> dict[str, Any]:
        d = asdict(self)
        d["franco_anticipado"] = 1 if self.franco_anticipado else 0
        d["sync_status"] = self.sync_status.value
        return d


@dataclass(frozen=True)
class DiaCalendarioCiclo:
    dia: int
    fecha: str
    fecha_iso: str
    dia_semana: str
    estado: str
    hora_disponible: str
    toma: str
    deja: str
    horas_trabajadas: str
    hs_semana: str
    novedad_id: int | str
    reemplaza_a: str
    turno_cubierto: str
    motivo: str
    periodo: str
    reemplazando_legajo: int | None
    reemplazando_nombre: str | None
    novedad_titular: str | None
    tramo_inicio: str | None
    tramo_fin: str | None
    franco_anticipado: bool
    observaciones: str


@dataclass(frozen=True)
class TipoNovedad:
    id: int | None
    nombre: str
    dias: int = 0
    exportar_descubierto: bool = True
    activo: bool = True
    sync_status: SyncStatus = SyncStatus.PENDING


@dataclass(frozen=True)
class Dotacion:
    id: int | None
    nombre: str
    activo: bool = True
    sync_status: SyncStatus = SyncStatus.PENDING


@dataclass(frozen=True)
class PersonalEstacion:
    id: int | None
    nombre: str
    dotacion: str | None = None
    activo: bool = True
    sync_status: SyncStatus = SyncStatus.PENDING


@dataclass(frozen=True)
class DestinatarioInforme:
    id: int | None
    nombre: str
    email: str
    activo: bool = True
    sync_status: SyncStatus = SyncStatus.PENDING


@dataclass(frozen=True)
class Operador:
    id: int
    username: str
    nombre: str
    legajo: int | None
    roles: tuple[str, ...] = field(default_factory=tuple)
    permisos: tuple[str, ...] = field(default_factory=tuple)
    activo: bool = True
    ultimo_ingreso: str | None = None


@dataclass(frozen=True)
class AuditoriaItem:
    id: int | None
    creado_en: str
    usuario: str
    accion: str
    entidad: str
    entidad_id: int | None
    datos_anteriores: Any = None
    datos_nuevos: Any = None


@dataclass(frozen=True)
class SyncQueueItem:
    id: int | None
    tabla: str
    registro_pk: str
    operacion: str
    payload_json: str
    creado_en: str
    intentos: int = 0
    ultimo_error: str | None = None
