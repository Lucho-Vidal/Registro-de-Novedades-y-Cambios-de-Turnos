# Core package
from .config import (
    APP_NAME,
    VERSION,
    DOTACIONES,
    ESPECIALIDADES_EMPLEADO,
    DIAS_SEMANA,
    resource_path,
    obtener_ruta_base_datos,
    cargar_configuracion_app,
    guardar_configuracion_app,
)
from .settings import AppSettings
from .logger import inicializar_logging, obtener_ruta_log
