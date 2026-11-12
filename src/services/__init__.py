"""Capa de Servicios de RENO v2.0.0.

Contiene los casos de uso, lógica de orquestación de negocio,
autenticación, generación de documentos (PDF/DOCX/Excel), correo y backups.
"""

try:
    from src.services.records_service import RecordsService
except ImportError:
    RecordsService = None

try:
    from src.services.auth_service import AuthService, CuentaBloqueadaError
except ImportError:
    AuthService = None
    CuentaBloqueadaError = None

try:
    from src.services.excel_service import ExcelService
except ImportError:
    ExcelService = None

try:
    from src.services.pdf_service import PDFService
except ImportError:
    PDFService = None

try:
    from src.services.word_service import WordService
except ImportError:
    WordService = None

try:
    from src.services.mail_service import MailService
except ImportError:
    MailService = None

try:
    from src.services.backup_service import BackupService
except ImportError:
    BackupService = None

from src.services.updater_service import (
    InfoActualizacion,
    VerificarActualizacionWorker,
    DescargarActualizacionWorker,
    ejecutar_instalador_y_reiniciar,
    comparar_versiones,
    parsear_version,
    es_actualizacion_obligatoria,
    calcular_sha256,
    dividir_archivo,
)

__all__ = [
    "RecordsService",
    "AuthService",
    "CuentaBloqueadaError",
    "ExcelService",
    "PDFService",
    "WordService",
    "MailService",
    "BackupService",
    "InfoActualizacion",
    "VerificarActualizacionWorker",
    "DescargarActualizacionWorker",
    "ejecutar_instalador_y_reiniciar",
    "comparar_versiones",
    "parsear_version",
    "es_actualizacion_obligatoria",
    "calcular_sha256",
    "dividir_archivo",
]
