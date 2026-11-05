"""Vista contenedora para Novedades Operativas y Cambios de Turno (Permutas) con pestañas internas."""

from typing import Any, Dict, Optional
from PySide6.QtCore import Signal
from PySide6.QtWidgets import QTabWidget, QVBoxLayout, QWidget

from src.services.excel_service import ExcelService
from src.services.pdf_service import PDFService
from src.services.records_service import RecordsService
from src.services.word_service import WordService
from src.ui.views.cambios_turno_view import CambiosTurnoView
from src.ui.views.novedades_view import NovedadesView
from src.ui.views.ordenamientos_view import OrdenamientosView


class NovedadesCambiosView(QWidget):
    """Contenedor de sub-pestañas para Novedades y Cambios de Turno."""

    registro_modificado = Signal()

    def __init__(
        self,
        records_service: RecordsService,
        excel_service: ExcelService,
        pdf_service: PDFService,
        word_service: WordService,
        usuario_actual: Optional[Dict[str, Any]] = None,
        tiene_permiso_fn=None,
        parent=None,
    ):
        super().__init__(parent)
        self.records_service = records_service
        self.excel_service = excel_service
        self.pdf_service = pdf_service
        self.word_service = word_service
        self.usuario_actual = usuario_actual or {}
        self.tiene_permiso_fn = tiene_permiso_fn

        self._init_ui()

    def _init_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(6, 6, 6, 6)
        layout.setSpacing(4)

        self.subtabs = QTabWidget()
        self.subtabs.setDocumentMode(True)

        # 1. Novedades Operativas
        self.tab_novedades = NovedadesView(
            self.records_service,
            self.excel_service,
            self.pdf_service,
            usuario_actual=self.usuario_actual,
            tiene_permiso_fn=self.tiene_permiso_fn,
            parent=self,
        )
        self.tab_novedades.registro_modificado.connect(self.registro_modificado.emit)
        self.subtabs.addTab(self.tab_novedades, "📋 Novedades Operativas")

        # 2. Cambios de Turno
        self.tab_cambios = CambiosTurnoView(
            self.records_service,
            self.excel_service,
            self.pdf_service,
            self.word_service,
            usuario_actual=self.usuario_actual,
            tiene_permiso_fn=self.tiene_permiso_fn,
            parent=self,
        )
        self.tab_cambios.registro_modificado.connect(self.registro_modificado.emit)
        self.subtabs.addTab(self.tab_cambios, "🔄 Cambios de Turno")

        # 3. Ordenamientos Operativos
        self.tab_ordenamientos = OrdenamientosView(
            self.records_service,
            self.excel_service,
            self.pdf_service,
            usuario_actual=self.usuario_actual,
            tiene_permiso_fn=self.tiene_permiso_fn,
            parent=self,
        )
        self.tab_ordenamientos.registro_modificado.connect(self.registro_modificado.emit)
        self.subtabs.addTab(self.tab_ordenamientos, "📄 Ordenamientos")

        layout.addWidget(self.subtabs)

    def recargar_datos(self) -> None:
        """Refresca los datos en todas las sub-pestañas."""
        if hasattr(self, "tab_novedades"):
            self.tab_novedades.recargar_datos()
        if hasattr(self, "tab_cambios"):
            self.tab_cambios.recargar_datos()
        if hasattr(self, "tab_ordenamientos"):
            self.tab_ordenamientos.recargar_datos()
