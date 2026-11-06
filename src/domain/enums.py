"""Enums y constantes tipadas del dominio RENO."""

from __future__ import annotations

from enum import Enum
from typing import Any


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
    "turnos.ver",
    "turnos.crear",
    "turnos.editar",
    "turnos.eliminar",
    "turnos.exportar",
    "turnos.administrar",
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
    "reemplazos.ver",
    "reemplazos.asignar",
    "empleados.administrar",
    "empleados.importar",
    "usuarios.administrar",
    "roles.administrar",
    "dotaciones.administrar",
    "tiposNovedad.administrar",
    "personalEstacion.ver",
    "personalEstacion.crear",
    "personalEstacion.editar",
    "personalEstacion.importar",
    "personalEstacion.exportar",
    "destinatarios_informe.administrar",
    "destinatarios_ordenamiento.administrar",
    "sesion.configurar",
    "auditoria.ver",
    "backup.gestionar",
    "registros.recuperar",
    "registros.modificar",
)

# Definiciones pedagógicas de las acciones para la interfaz y configuración de roles
DEFINICIONES_ACCIONES: dict[str, str] = {
    "Ver (Lectura)": "Permite acceder a la pestaña o subpestaña y consultar datos, diagramas y filtros sin modificar información.",
    "Crear": "Habilita los formularios y botones de alta ('+ Nuevo') para ingresar nuevos registros en el sistema.",
    "Editar": "Permite modificar registros existentes dentro de las ventanas temporales autorizadas (24h en novedades / 72h en cambios).",
    "Eliminar": "Permite anular o enviar a la papelera registros existentes dentro del plazo temporal permitido.",
    "Exportar": "Habilita la descarga de información y reportes en formatos Excel (.xlsx), PDF o documentos oficiales Word (.docx).",
    "Importar": "Permite la carga e incorporación masiva de datos estructurados desde archivos externos de Excel.",
    "Administrar": "Otorga control integral y estructural sobre el catálogo o módulo (configuración de parámetros, usuarios, asignación de roles, dotaciones y anulación de restricciones operativas). A diferencia de 'Crear' o 'Editar', quien administra gestiona la arquitectura y reglas del sistema.",
    "Supervisión Irrestricta": "Permiso especial que anula los límites de tiempo de 24h/72h, permitiendo editar o eliminar cualquier novedad o cambio histórico.",
}

# Metadatos detallados de cada permiso: acción, título amigable, pestaña, subpestaña y descripción
PERMISOS_METADATA: dict[str, dict[str, str]] = {
    # 1. Buscador de Turnos e Itinerarios
    "turnos.ver": {
        "codigo": "turnos.ver",
        "accion": "Ver (Lectura)",
        "titulo": "Consultar Turnos e Itinerarios",
        "pestana": "🔍 Buscador de Turnos",
        "subpestana": "Consulta y Diagramas",
        "descripcion": "Permite acceder al buscador de turnos, visualizar diagramas y paradas en modo solo lectura.",
    },
    "turnos.crear": {
        "codigo": "turnos.crear",
        "accion": "Crear",
        "titulo": "Registrar Nuevos Turnos",
        "pestana": "🔍 Buscador de Turnos",
        "subpestana": "Consulta y Diagramas",
        "descripcion": "Habilita la creación manual de turnos en la malla horaria.",
    },
    "turnos.editar": {
        "codigo": "turnos.editar",
        "accion": "Editar",
        "titulo": "Modificar Turnos",
        "pestana": "🔍 Buscador de Turnos",
        "subpestana": "Consulta y Diagramas",
        "descripcion": "Permite modificar cabeceras, estaciones y horarios de turnos existentes.",
    },
    "turnos.eliminar": {
        "codigo": "turnos.eliminar",
        "accion": "Eliminar",
        "titulo": "Eliminar Turnos",
        "pestana": "🔍 Buscador de Turnos",
        "subpestana": "Consulta y Diagramas",
        "descripcion": "Permite eliminar turnos cargados en la base de datos.",
    },
    "turnos.exportar": {
        "codigo": "turnos.exportar",
        "accion": "Exportar",
        "titulo": "Exportar Diagramación a Excel",
        "pestana": "🔍 Buscador de Turnos",
        "subpestana": "Consulta y Diagramas",
        "descripcion": "Permite descargar la diagramación y búsqueda de turnos a planillas Excel (.xlsx).",
    },
    "turnos.administrar": {
        "codigo": "turnos.administrar",
        "accion": "Administrar",
        "titulo": "Administración e Importación Masiva de Malla",
        "pestana": "🔍 Buscador de Turnos",
        "subpestana": "Consulta y Diagramas",
        "descripcion": "Permite la importación masiva de itinerarios/paradas desde Excel y control general de diagramas.",
    },

    # 2. Novedades y Cambios > Novedades Operativas
    "novedades.ver": {
        "codigo": "novedades.ver",
        "accion": "Ver (Lectura)",
        "titulo": "Consultar Novedades y Turnos Descubiertos",
        "pestana": "📝 Novedades y Cambios",
        "subpestana": "📋 Subpestaña: Novedades Operativas",
        "descripcion": "Permite ver el listado de novedades, la pestaña de turnos sin cubrir y el visor de detalle.",
    },
    "novedades.crear": {
        "codigo": "novedades.crear",
        "accion": "Crear",
        "titulo": "Registrar Nueva Novedad",
        "pestana": "📝 Novedades y Cambios",
        "subpestana": "📋 Subpestaña: Novedades Operativas",
        "descripcion": "Habilita el botón '+ Nueva Novedad' y el formulario de alta para ausencias, licencias médicas, etc.",
    },
    "novedades.editar": {
        "codigo": "novedades.editar",
        "accion": "Editar",
        "titulo": "Modificar Novedades",
        "pestana": "📝 Novedades y Cambios",
        "subpestana": "📋 Subpestaña: Novedades Operativas",
        "descripcion": "Permite editar novedades vigentes dentro del límite temporal configurado (por defecto 24 hs).",
    },
    "novedades.eliminar": {
        "codigo": "novedades.eliminar",
        "accion": "Eliminar",
        "titulo": "Eliminar Novedades",
        "pestana": "📝 Novedades y Cambios",
        "subpestana": "📋 Subpestaña: Novedades Operativas",
        "descripcion": "Permite dar de baja o enviar a la papelera novedades dentro del límite temporal permitido.",
    },
    "novedades.importar": {
        "codigo": "novedades.importar",
        "accion": "Importar",
        "titulo": "Importación Masiva de Novedades",
        "pestana": "📝 Novedades y Cambios",
        "subpestana": "📋 Subpestaña: Novedades Operativas",
        "descripcion": "Permite cargar novedades de manera masiva desde archivos externos de Excel.",
    },
    "novedades.exportar": {
        "codigo": "novedades.exportar",
        "accion": "Exportar",
        "titulo": "Exportar Novedades (Excel / PDF)",
        "pestana": "📝 Novedades y Cambios",
        "subpestana": "📋 Subpestaña: Novedades Operativas",
        "descripcion": "Permite descargar listados en Excel y generar informes diarios ejecutivos en PDF (ReportLab).",
    },

    # 3. Novedades y Cambios > Cambios de Turno
    "cambios_turno.ver": {
        "codigo": "cambios_turno.ver",
        "accion": "Ver (Lectura)",
        "titulo": "Consultar Cambios de Turno",
        "pestana": "📝 Novedades y Cambios",
        "subpestana": "🔄 Subpestaña: Cambios de Turno",
        "descripcion": "Permite ver el registro de cambios de turno y permutas entre empleados.",
    },
    "cambios_turno.crear": {
        "codigo": "cambios_turno.crear",
        "accion": "Crear",
        "titulo": "Registrar Nuevo Cambio de Turno",
        "pestana": "📝 Novedades y Cambios",
        "subpestana": "🔄 Subpestaña: Cambios de Turno",
        "descripcion": "Habilita el botón '+ Nuevo Cambio' para cargar permutas entre empleados solicitantes y aceptantes.",
    },
    "cambios_turno.editar": {
        "codigo": "cambios_turno.editar",
        "accion": "Editar",
        "titulo": "Modificar Cambios de Turno",
        "pestana": "📝 Novedades y Cambios",
        "subpestana": "🔄 Subpestaña: Cambios de Turno",
        "descripcion": "Permite modificar datos de permutas dentro del plazo configurado (por defecto 72 hs).",
    },
    "cambios_turno.eliminar": {
        "codigo": "cambios_turno.eliminar",
        "accion": "Eliminar",
        "titulo": "Eliminar Cambios de Turno",
        "pestana": "📝 Novedades y Cambios",
        "subpestana": "🔄 Subpestaña: Cambios de Turno",
        "descripcion": "Permite cancelar o enviar cambios de turno a la papelera dentro del tiempo permitido.",
    },
    "cambios_turno.importar": {
        "codigo": "cambios_turno.importar",
        "accion": "Importar",
        "titulo": "Importación Masiva de Cambios",
        "pestana": "📝 Novedades y Cambios",
        "subpestana": "🔄 Subpestaña: Cambios de Turno",
        "descripcion": "Permite incorporar registros de cambios de turno desde archivos Excel.",
    },
    "cambios_turno.exportar": {
        "codigo": "cambios_turno.exportar",
        "accion": "Exportar",
        "titulo": "Exportar Cambios (Excel / PDF / Word)",
        "pestana": "📝 Novedades y Cambios",
        "subpestana": "🔄 Subpestaña: Cambios de Turno",
        "descripcion": "Habilita la descarga a Excel, informes PDF y generación de notas oficiales Word (.docx).",
    },

    # 4. Novedades y Cambios > Ordenamientos Operativos
    "ordenamientos.ver": {
        "codigo": "ordenamientos.ver",
        "accion": "Ver (Lectura)",
        "titulo": "Consultar Ordenamientos",
        "pestana": "📝 Novedades y Cambios",
        "subpestana": "📄 Subpestaña: Ordenamientos",
        "descripcion": "Permite acceder y consultar el historial de ordenamientos y afectaciones de personal.",
    },
    "ordenamientos.crear": {
        "codigo": "ordenamientos.crear",
        "accion": "Crear",
        "titulo": "Emitir Nuevo Ordenamiento",
        "pestana": "📝 Novedades y Cambios",
        "subpestana": "📄 Subpestaña: Ordenamientos",
        "descripcion": "Habilita la generación y emisión de nuevos ordenamientos de servicio.",
    },
    "ordenamientos.editar": {
        "codigo": "ordenamientos.editar",
        "accion": "Editar",
        "titulo": "Modificar Ordenamientos",
        "pestana": "📝 Novedades y Cambios",
        "subpestana": "📄 Subpestaña: Ordenamientos",
        "descripcion": "Permite modificar datos u observaciones en ordenamientos vigentes.",
    },
    "ordenamientos.eliminar": {
        "codigo": "ordenamientos.eliminar",
        "accion": "Eliminar",
        "titulo": "Anular / Eliminar Ordenamientos",
        "pestana": "📝 Novedades y Cambios",
        "subpestana": "📄 Subpestaña: Ordenamientos",
        "descripcion": "Permite cancelar o eliminar ordenamientos operativos.",
    },
    "ordenamientos.exportar": {
        "codigo": "ordenamientos.exportar",
        "accion": "Exportar",
        "titulo": "Exportar Ordenamientos (PDF / Excel)",
        "pestana": "📝 Novedades y Cambios",
        "subpestana": "📄 Subpestaña: Ordenamientos",
        "descripcion": "Permite descargar órdenes operativas en documentos PDF y planillas Excel.",
    },

    # 5. Personal de Ciclo y Reemplazos
    "reemplazos.ver": {
        "codigo": "reemplazos.ver",
        "accion": "Ver (Lectura)",
        "titulo": "Consultar Personal de Ciclo y Calendarios",
        "pestana": "⏱️ Personal de Ciclo",
        "subpestana": "Personal de Ciclo",
        "descripcion": "Permite acceder a la pestaña de Personal de Ciclo, ver diagramas mensuales, descansos y novedades propias.",
    },
    "reemplazos.asignar": {
        "codigo": "reemplazos.asignar",
        "accion": "Asignar / Editar",
        "titulo": "Asignar Coberturas, Jornadas y Francos",
        "pestana": "⏱️ Personal de Ciclo",
        "subpestana": "Personal de Ciclo",
        "descripcion": "Permite asignar tramos de cobertura a personal de ciclo respetando descansos CCT y horas semanales.",
    },

    # 6. Empleados (BASE)
    "empleados.administrar": {
        "codigo": "empleados.administrar",
        "accion": "Administrar",
        "titulo": "Gestión Integral del Padrón de Empleados",
        "pestana": "👷 Empleados (BASE)",
        "subpestana": "Padrón de Empleados",
        "descripcion": "Administración total de la nómina: altas, bajas lógicas, cambio de dotación, francos y especialidad.",
    },
    "empleados.importar": {
        "codigo": "empleados.importar",
        "accion": "Importar",
        "titulo": "Importación Masiva de Empleados",
        "pestana": "👷 Empleados (BASE)",
        "subpestana": "Padrón de Empleados",
        "descripcion": "Permite incorporar o actualizar la nómina masivamente desde planillas Excel oficiales.",
    },

    # 7. Administración > Usuarios y Roles
    "usuarios.administrar": {
        "codigo": "usuarios.administrar",
        "accion": "Administrar",
        "titulo": "Gestión de Cuentas de Usuario",
        "pestana": "⚙️ Administración",
        "subpestana": "👥 Usuarios y Roles",
        "descripcion": "Permite crear usuarios, restablecer contraseñas, activar/desactivar accesos y asignar roles.",
    },
    "roles.administrar": {
        "codigo": "roles.administrar",
        "accion": "Administrar",
        "titulo": "Gestión de Roles y Permisos",
        "pestana": "⚙️ Administración",
        "subpestana": "👥 Usuarios y Roles",
        "descripcion": "Permite crear nuevos roles y configurar granularmente su matriz de permisos por pestaña y subpestaña.",
    },

    # 8. Administración > Catálogos Operativos
    "dotaciones.administrar": {
        "codigo": "dotaciones.administrar",
        "accion": "Administrar",
        "titulo": "Gestión de Dotaciones Operativas",
        "pestana": "⚙️ Administración",
        "subpestana": "📋 Catálogos Operativos",
        "descripcion": "Permite crear, editar y dar de baja dotaciones operativas del sistema.",
    },
    "tiposNovedad.administrar": {
        "codigo": "tiposNovedad.administrar",
        "accion": "Administrar",
        "titulo": "Gestión de Tipos de Novedad",
        "pestana": "⚙️ Administración",
        "subpestana": "📋 Catálogos Operativos",
        "descripcion": "Permite configurar tipos de novedad, códigos breves, duración estimada en días y turnos sin cubrir.",
    },
    "personalEstacion.ver": {
        "codigo": "personalEstacion.ver",
        "accion": "Ver (Lectura)",
        "titulo": "Consultar Personal de Estación",
        "pestana": "⚙️ Administración",
        "subpestana": "📋 Catálogos Operativos",
        "descripcion": "Permite consultar el catálogo de personal de estación asignado a coberturas.",
    },
    "personalEstacion.crear": {
        "codigo": "personalEstacion.crear",
        "accion": "Crear",
        "titulo": "Registrar Personal de Estación",
        "pestana": "⚙️ Administración",
        "subpestana": "📋 Catálogos Operativos",
        "descripcion": "Permite dar de alta personal de estación en el catálogo operativo.",
    },
    "personalEstacion.editar": {
        "codigo": "personalEstacion.editar",
        "accion": "Editar",
        "titulo": "Modificar Personal de Estación",
        "pestana": "⚙️ Administración",
        "subpestana": "📋 Catálogos Operativos",
        "descripcion": "Permite editar datos de personal de estación existente.",
    },
    "personalEstacion.importar": {
        "codigo": "personalEstacion.importar",
        "accion": "Importar",
        "titulo": "Importar Personal de Estación",
        "pestana": "⚙️ Administración",
        "subpestana": "📋 Catálogos Operativos",
        "descripcion": "Permite importar personal de estación masivamente desde archivos Excel.",
    },
    "personalEstacion.exportar": {
        "codigo": "personalEstacion.exportar",
        "accion": "Exportar",
        "titulo": "Exportar Personal de Estación",
        "pestana": "⚙️ Administración",
        "subpestana": "📋 Catálogos Operativos",
        "descripcion": "Permite descargar el catálogo de personal de estación en formato Excel.",
    },

    # 9. Administración > Destinatarios de Correo
    "destinatarios_informe.administrar": {
        "codigo": "destinatarios_informe.administrar",
        "accion": "Administrar",
        "titulo": "Destinatarios de Informes Diarios",
        "pestana": "⚙️ Administración",
        "subpestana": "📧 Destinatarios de Correo",
        "descripcion": "Permite configurar las cuentas de correo electrónico para el envío automático del informe diario de novedades.",
    },
    "destinatarios_ordenamiento.administrar": {
        "codigo": "destinatarios_ordenamiento.administrar",
        "accion": "Administrar",
        "titulo": "Destinatarios de Ordenamientos",
        "pestana": "⚙️ Administración",
        "subpestana": "📧 Destinatarios de Correo",
        "descripcion": "Permite configurar la lista de correos para la distribución oficial de ordenamientos operativos.",
    },

    # 10. Administración > Seguridad, Auditoría y Sistema
    "sesion.configurar": {
        "codigo": "sesion.configurar",
        "accion": "Configurar",
        "titulo": "Configuración General y Base SQLite",
        "pestana": "⚙️ Administración",
        "subpestana": "💾 Seguridad, Auditoría y Sistema",
        "descripcion": "Permite ajustar tiempo de inactividad, ventanas horarias de edición (24h/72h) y ruta persistente del archivo SQLite.",
    },
    "auditoria.ver": {
        "codigo": "auditoria.ver",
        "accion": "Ver / Exportar",
        "titulo": "Trazabilidad y Auditoría",
        "pestana": "⚙️ Administración",
        "subpestana": "💾 Seguridad, Auditoría y Sistema",
        "descripcion": "Permite revisar el log histórico inmutable de acciones del sistema y exportarlo a Excel.",
    },
    "backup.gestionar": {
        "codigo": "backup.gestionar",
        "accion": "Gestionar",
        "titulo": "Copias de Seguridad (Backups)",
        "pestana": "⚙️ Administración",
        "subpestana": "💾 Seguridad, Auditoría y Sistema",
        "descripcion": "Permite generar backups manuales, configurar retención y restaurar copias de seguridad de la base de datos.",
    },
    "registros.recuperar": {
        "codigo": "registros.recuperar",
        "accion": "Recuperar",
        "titulo": "Papelera de Reciclaje y Restauración",
        "pestana": "⚙️ Administración",
        "subpestana": "💾 Seguridad, Auditoría y Sistema",
        "descripcion": "Permite acceder a la papelera de registros eliminados y recuperarlos hacia la base activa.",
    },
    "registros.modificar": {
        "codigo": "registros.modificar",
        "accion": "Supervisión Irrestricta",
        "titulo": "Modificación y Baja Irrestricta (Supervisores)",
        "pestana": "⚙️ Administración",
        "subpestana": "💾 Seguridad, Auditoría y Sistema",
        "descripcion": "Otorga permiso exclusivo a supervisores para editar o eliminar novedades y cambios históricos fuera de las ventanas de 24h/72h.",
    },
}

# Estructura jerárquica de permisos agrupada estrictamente por Pestaña y Subpestaña
ESTRUCTURA_PERMISOS_PESTANAS: tuple[dict[str, Any], ...] = (
    {
        "pestana": "🔍 Pestaña: Buscador de Turnos e Itinerarios",
        "icono": "🔍",
        "subgrupos": [
            {
                "subpestana": "Malla Horaria e Itinerarios",
                "descripcion": "Permisos para consulta, alta, edición, eliminación y exportación de diagramas y paradas.",
                "permisos": [
                    "turnos.ver",
                    "turnos.crear",
                    "turnos.editar",
                    "turnos.eliminar",
                    "turnos.exportar",
                    "turnos.administrar",
                ],
            }
        ],
    },
    {
        "pestana": "📝 Pestaña: Novedades y Cambios",
        "icono": "📝",
        "subgrupos": [
            {
                "subpestana": "📋 Subpestaña: Novedades Operativas",
                "descripcion": "Registro diario de novedades del personal, turnos descubiertos e informes.",
                "permisos": [
                    "novedades.ver",
                    "novedades.crear",
                    "novedades.editar",
                    "novedades.eliminar",
                    "novedades.importar",
                    "novedades.exportar",
                ],
            },
            {
                "subpestana": "🔄 Subpestaña: Cambios de Turno (Permutas)",
                "descripcion": "Permutas entre empleados, notas oficiales Word y exportaciones.",
                "permisos": [
                    "cambios_turno.ver",
                    "cambios_turno.crear",
                    "cambios_turno.editar",
                    "cambios_turno.eliminar",
                    "cambios_turno.importar",
                    "cambios_turno.exportar",
                ],
            },
            {
                "subpestana": "📄 Subpestaña: Ordenamientos Operativos",
                "descripcion": "Emisión, seguimiento, edición y exportación de órdenes operativas de servicio.",
                "permisos": [
                    "ordenamientos.ver",
                    "ordenamientos.crear",
                    "ordenamientos.editar",
                    "ordenamientos.eliminar",
                    "ordenamientos.exportar",
                ],
            },
        ],
    },
    {
        "pestana": "⏱️ Pestaña: Personal de Ciclo y Reemplazos",
        "icono": "⏱️",
        "subgrupos": [
            {
                "subpestana": "Personal de Ciclo",
                "descripcion": "Diagramas mensuales, cómputo de horas CCT y cobertura de turnos descubiertos.",
                "permisos": [
                    "reemplazos.ver",
                    "reemplazos.asignar",
                ],
            }
        ],
    },
    {
        "pestana": "👷 Pestaña: Empleados (BASE)",
        "icono": "👷",
        "subgrupos": [
            {
                "subpestana": "Padrón de Empleados",
                "descripcion": "Administración del padrón general de empleados e importación desde Excel.",
                "permisos": [
                    "empleados.administrar",
                    "empleados.importar",
                ],
            }
        ],
    },
    {
        "pestana": "⚙️ Pestaña: Administración",
        "icono": "⚙️",
        "subgrupos": [
            {
                "subpestana": "👥 Subpestaña: Gestión de Usuarios y Roles",
                "descripcion": "Cuentas de usuario, perfiles de acceso y matriz de permisos.",
                "permisos": [
                    "usuarios.administrar",
                    "roles.administrar",
                ],
            },
            {
                "subpestana": "📋 Subpestaña: Catálogos Operativos",
                "descripcion": "Tipos de novedad, dotaciones y relevos de personal de estación.",
                "permisos": [
                    "dotaciones.administrar",
                    "tiposNovedad.administrar",
                    "personalEstacion.ver",
                    "personalEstacion.crear",
                    "personalEstacion.editar",
                    "personalEstacion.importar",
                    "personalEstacion.exportar",
                ],
            },
            {
                "subpestana": "📧 Subpestaña: Destinatarios de Correo",
                "descripcion": "Listas de correos separadas para informe diario y para ordenamientos.",
                "permisos": [
                    "destinatarios_informe.administrar",
                    "destinatarios_ordenamiento.administrar",
                ],
            },
            {
                "subpestana": "💾 Subpestaña: Seguridad, Auditoría y Sistema",
                "descripcion": "Parámetros globales, copias de seguridad, papelera, auditoría y supervisión irrestricta.",
                "permisos": [
                    "sesion.configurar",
                    "auditoria.ver",
                    "backup.gestionar",
                    "registros.recuperar",
                    "registros.modificar",
                ],
            },
        ],
    },
)

PERMISOS_GRUPOS: tuple[tuple[str, str], ...] = (
    ("turnos", "Turnos e Itinerarios"),
    ("novedades", "Novedades"),
    ("cambios_turno", "Cambios de turno"),
    ("ordenamientos", "Ordenamientos operativos"),
    ("reemplazos", "Reemplazos / Ciclo"),
    ("empleados", "Empleados"),
    ("personalEstacion", "Personal de estación"),
    ("dotaciones", "Dotaciones"),
    ("tiposNovedad", "Tipos de novedad"),
    ("destinatarios_informe", "Destinatarios de informes"),
    ("destinatarios_ordenamiento", "Destinatarios de ordenamientos"),
    ("usuarios", "Usuarios"),
    ("roles", "Roles y permisos"),
    ("auditoria", "Auditoría"),
    ("sesion", "Sesión"),
    ("backup", "Copias de seguridad"),
    ("registros", "Registros"),
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
