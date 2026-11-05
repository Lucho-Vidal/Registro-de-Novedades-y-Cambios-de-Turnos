"""Enums y constantes tipadas del dominio RENO."""

from __future__ import annotations

from enum import Enum


class SyncStatus(str, Enum):
    PENDING = "pending"
    SYNCED = "synced"
    ERROR = "error"


class SyncOperation(str, Enum):
    UPSERT = "UPSERT"
    DELETE = "DELETE"


class EstadoCobertura(str, Enum):
    SI = "Sí"
    PARCIAL = "Parcial"
    NO = "No"
    NA_CICLO = "N/A (Ciclo)"


class EstadoDiaCiclo(str, Enum):
    DISPONIBLE = "Disponible"
    REEMPLAZANDO = "Reemplazando"
    FRANCO = "Franco"
    FRANCO_ANTICIPADO = "Franco Anticipado"


class ModoEdicionNovedad(str, Enum):
    COMPLETA = "completa"
    FECHA_FIN = "fecha_fin"


PERMISOS_BASE: tuple[str, ...] = (
    "novedades.ver",
    "novedades.crear",
    "novedades.editar",
    "novedades.eliminar",
    "novedades.importar",
    "novedades.exportar",
    "cambios_turno.ver",
    "cambios_turno.crear",
    "cambios_turno.editar",
    "cambios_turno.eliminar",
    "cambios_turno.importar",
    "cambios_turno.exportar",
    "ordenamientos.ver",
    "ordenamientos.crear",
    "ordenamientos.editar",
    "ordenamientos.eliminar",
    "ordenamientos.exportar",
    "usuarios.administrar",
    "roles.administrar",
    "empleados.importar",
    "empleados.administrar",
    "auditoria.ver",
    "dotaciones.administrar",
    "tiposNovedad.administrar",
    "personalEstacion.ver",
    "personalEstacion.crear",
    "personalEstacion.editar",
    "personalEstacion.importar",
    "personalEstacion.exportar",
    "destinatarios_informe.administrar",
    "turnos.ver",
    "turnos.crear",
    "turnos.editar",
    "turnos.eliminar",
    "turnos.administrar",
    "sesion.configurar",
    "registros.recuperar",
    "registros.modificar",
    "backup.gestionar",
    "reemplazos.ver",
    "reemplazos.asignar",
)

PERMISOS_GRUPOS: tuple[tuple[str, str], ...] = (
    ("novedades", "Novedades"),
    ("cambios_turno", "Cambios de turno"),
    ("ordenamientos", "Ordenamientos operativos"),
    ("empleados", "Empleados"),
    ("personalEstacion", "Personal de estación"),
    ("dotaciones", "Dotaciones"),
    ("tiposNovedad", "Tipos de novedad"),
    ("turnos", "Turnos e Itinerarios"),
    ("destinatarios_informe", "Destinatarios de informes"),
    ("usuarios", "Usuarios"),
    ("roles", "Roles y permisos"),
    ("auditoria", "Auditoría"),
    ("sesion", "Sesión"),
    ("registros", "Registros"),
    ("reemplazos", "Reemplazos / Ciclo"),
    ("backup", "Copias de seguridad"),
)

ACCIONES_PERMISO: dict[str, str] = {
    "ver": "Ver",
    "crear": "Crear",
    "editar": "Editar",
    "eliminar": "Eliminar",
    "importar": "Importar",
    "exportar": "Exportar",
    "administrar": "Administrar",
    "configurar": "Configurar",
    "recuperar": "Recuperar",
    "gestionar": "Gestionar",
    "modificar": "Modificar",
    "asignar": "Asignar",
}
