"""Diálogo modal interactivo para descarga e instalación de actualizaciones (Estándar CRONO-PCT)."""

from __future__ import annotations

import sys
from typing import Any, Optional

from PySide6.QtCore import Qt
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)

from src.core.config import APP_VERSION
from src.core.logger import get_logger
from src.core.theme import aplicar_icono_aplicacion
from src.services.updater_service import (
    DescargarActualizacionWorker,
    InfoActualizacion,
    ejecutar_instalador_y_reiniciar,
)

logger = get_logger("ActualizacionDialog")


class ActualizacionDialog(QDialog):
    """Ventana interactiva de actualización de RENO con changelog, barra de progreso y reinicio desatendido."""

    def __init__(
        self,
        info: InfoActualizacion,
        es_obligatoria: bool = False,
        parent: Optional[QWidget] = None,
        config: Optional[dict[str, Any]] = None,
    ):
        super().__init__(parent)
        self.info = info
        self.es_obligatoria = es_obligatoria
        self.config = config
        self.worker_descarga: Optional[DescargarActualizacionWorker] = None
        self.ruta_instalador_descargado: Optional[str] = None

        self.setWindowTitle(f"Actualización Disponible — RENO v{self.info.version}")
        self.resize(640, 520)
        self.setMinimumSize(560, 460)
        aplicar_icono_aplicacion(self)

        # Si es obligatoria, no permitimos cerrar con la cruz de la ventana sin actualizar
        if self.es_obligatoria:
            self.setWindowFlags(
                (self.windowFlags() & ~Qt.WindowContextHelpButtonHint & ~Qt.WindowCloseButtonHint)
                | Qt.CustomizeWindowHint
                | Qt.WindowTitleHint
            )
        else:
            self.setWindowFlags((self.windowFlags() & ~Qt.WindowContextHelpButtonHint) | Qt.WindowCloseButtonHint)

        self._init_ui()

    def _init_ui(self) -> None:
        layout_principal = QVBoxLayout(self)
        layout_principal.setContentsMargins(20, 20, 20, 20)
        layout_principal.setSpacing(14)

        # 1. Cabecera visual con estilo card
        header_card = QFrame()
        header_card.setObjectName("headerCard")
        header_card.setStyleSheet(
            """
            #headerCard {
                background-color: #F8FAFC;
                border: 1px solid #E2E8F0;
                border-radius: 8px;
            }
            """
        )
        header_layout = QHBoxLayout(header_card)
        header_layout.setContentsMargins(16, 14, 16, 14)

        lbl_icon = QLabel("🚀" if not self.es_obligatoria else "⚠️")
        lbl_icon.setStyleSheet("font-size: 32pt;")
        header_layout.addWidget(lbl_icon)

        info_header = QVBoxLayout()
        info_header.setSpacing(3)

        lbl_badge = QLabel(f"RENO v{self.info.version} Disponible")
        lbl_badge.setStyleSheet(
            "font-size: 13pt; font-weight: bold; color: #1E3A8A;"
            if not self.es_obligatoria
            else "font-size: 13pt; font-weight: bold; color: #D97706;"
        )

        lbl_sub = QLabel(f"Versión actual instalada: v{APP_VERSION}  •  Fecha de publicación: {self.info.fecha or 'Reciente'}")
        lbl_sub.setStyleSheet("font-size: 9.5pt; color: #64748B;")

        lbl_titulo = QLabel(self.info.titulo)
        lbl_titulo.setStyleSheet("font-size: 10pt; font-weight: 600; color: #334155;")
        lbl_titulo.setWordWrap(True)

        info_header.addWidget(lbl_badge)
        info_header.addWidget(lbl_sub)
        info_header.addWidget(lbl_titulo)
        header_layout.addLayout(info_header)
        header_layout.addStretch()

        layout_principal.addWidget(header_card)

        # 2. Banner de advertencia si es obligatoria
        if self.es_obligatoria:
            banner_warn = QFrame()
            banner_warn.setStyleSheet(
                """
                background-color: #FFFBEB;
                border: 1px solid #FCD34D;
                border-left: 4px solid #D97706;
                border-radius: 4px;
                padding: 8px 12px;
                """
            )
            warn_lay = QHBoxLayout(banner_warn)
            warn_lay.setContentsMargins(4, 4, 4, 4)
            lbl_warn_txt = QLabel(
                "<b>Actualización Requerida:</b> Esta versión introduce cambios estructurales de base de datos o correcciones "
                "esenciales. Es obligatorio actualizar para continuar operando en red con seguridad."
            )
            lbl_warn_txt.setStyleSheet("color: #92400E; font-size: 9.5pt;")
            lbl_warn_txt.setWordWrap(True)
            warn_lay.addWidget(lbl_warn_txt)
            layout_principal.addWidget(banner_warn)

        # 3. Visor de novedades (Changelog)
        lbl_seccion = QLabel("Novedades y mejoras incluidas:")
        lbl_seccion.setStyleSheet("font-size: 10pt; font-weight: 600; color: #1E3A8A;")
        layout_principal.addWidget(lbl_seccion)

        self.browser_novedades = QTextBrowser()
        self.browser_novedades.setStyleSheet(
            """
            QTextBrowser {
                background-color: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 6px;
                padding: 8px;
            }
            """
        )
        self._cargar_novedades_html()
        layout_principal.addWidget(self.browser_novedades, stretch=1)

        # 4. Sección de progreso de descarga (inicialmente oculta o en reposo)
        self.frame_progreso = QFrame()
        self.frame_progreso.setStyleSheet(
            """
            QFrame {
                background-color: #F1F5F9;
                border: 1px solid #E2E8F0;
                border-radius: 6px;
                padding: 6px 10px;
            }
            """
        )
        lay_progreso = QVBoxLayout(self.frame_progreso)
        lay_progreso.setSpacing(6)

        header_prog = QHBoxLayout()
        self.lbl_estado_descarga = QLabel("Listo para descargar.")
        self.lbl_estado_descarga.setStyleSheet("font-weight: 600; font-size: 9pt; color: #334155;")
        self.lbl_velocidad = QLabel("")
        self.lbl_velocidad.setStyleSheet("font-size: 9pt; color: #64748B;")
        header_prog.addWidget(self.lbl_estado_descarga)
        header_prog.addStretch()
        header_prog.addWidget(self.lbl_velocidad)
        lay_progreso.addLayout(header_prog)

        self.barra_progreso = QProgressBar()
        self.barra_progreso.setRange(0, 100)
        self.barra_progreso.setValue(0)
        self.barra_progreso.setTextVisible(True)
        self.barra_progreso.setFixedHeight(18)
        self.barra_progreso.setStyleSheet(
            """
            QProgressBar {
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                text-align: center;
                background-color: #FFFFFF;
                font-size: 9pt;
            }
            QProgressBar::chunk {
                background-color: #2563EB;
                border-radius: 3px;
            }
            """
        )
        lay_progreso.addWidget(self.barra_progreso)

        layout_principal.addWidget(self.frame_progreso)

        # 5. Barra de botones inferior
        layout_botones = QHBoxLayout()
        layout_botones.setSpacing(10)

        # Botón posponer (solo si es opcional)
        if not self.es_obligatoria:
            self.btn_posponer = QPushButton("Recordarme más tarde")
            self.btn_posponer.setStyleSheet(
                """
                QPushButton {
                    padding: 8px 16px;
                    background-color: transparent;
                    color: #64748B;
                    border: 1px solid #CBD5E1;
                    border-radius: 6px;
                    font-weight: 500;
                }
                QPushButton:hover {
                    background-color: #F1F5F9;
                    color: #334155;
                }
                """
            )
            self.btn_posponer.clicked.connect(self.reject)
            layout_botones.addWidget(self.btn_posponer)

        layout_botones.addStretch()

        self.btn_accion = QPushButton("Descargar e Instalar Ahora")
        self.btn_accion.setStyleSheet(
            """
            QPushButton {
                padding: 8px 22px;
                background-color: #2563EB;
                color: #FFFFFF;
                border: none;
                border-radius: 6px;
                font-weight: bold;
                font-size: 10pt;
            }
            QPushButton:hover {
                background-color: #1D4ED8;
            }
            """
        )
        self.btn_accion.clicked.connect(self._on_clic_accion_principal)
        layout_botones.addWidget(self.btn_accion)

        layout_principal.addLayout(layout_botones)

    def _cargar_novedades_html(self) -> None:
        """Formatea las novedades en HTML limpio y legible."""
        items_html = ""
        for nov in self.info.novedades:
            # Resaltar etiquetas como [Operación], [Seguridad], etc.
            linea = nov.strip()
            if linea.startswith("[") and "]" in linea:
                tag, resto = linea.split("]", 1)
                tag_clean = tag.lstrip("[").strip()
                linea_html = f"<b><span style='color: #0369A1;'>[{tag_clean}]</span></b> {resto.strip()}"
            else:
                linea_html = linea
            items_html += f"<li style='margin-bottom: 6px; line-height: 1.3;'>{linea_html}</li>"

        html = f"""
        <div style="font-family: 'Segoe UI', Calibri, sans-serif; font-size: 9.5pt; color: #334155;">
            <ul style="margin: 4px 0 4px 18px; padding: 0;">
                {items_html or '<li>Mejoras generales de rendimiento, estabilidad y seguridad.</li>'}
            </ul>
        </div>
        """
        self.browser_novedades.setHtml(html)

    def _on_clic_accion_principal(self) -> None:
        """Controla el flujo de la acción según el estado actual."""
        if self.ruta_instalador_descargado:
            # Estado: Ya descargado -> Ejecutar e instalar
            self._ejecutar_instalador()
            return

        if self.worker_descarga and self.worker_descarga.isRunning():
            # Estado: Descargando -> Cancelar
            self.worker_descarga.cancelar()
            self.lbl_estado_descarga.setText("Cancelando descarga...")
            self.btn_accion.setEnabled(False)
            return

        # Estado inicial: Iniciar descarga
        self._iniciar_descarga()

    def _iniciar_descarga(self) -> None:
        self.btn_accion.setText("Cancelar")
        self.btn_accion.setStyleSheet(
            """
            QPushButton {
                padding: 8px 18px;
                background-color: #EF4444;
                color: #FFFFFF;
                border: none;
                border-radius: 6px;
                font-weight: 500;
            }
            QPushButton:hover {
                background-color: #DC2626;
            }
            """
        )
        if hasattr(self, "btn_posponer"):
            self.btn_posponer.setEnabled(False)

        self.lbl_estado_descarga.setText("Iniciando conexión con el servidor...")
        self.barra_progreso.setValue(0)

        self.worker_descarga = DescargarActualizacionWorker(self.info, config=self.config, parent=self)
        self.worker_descarga.progreso.connect(self._on_progreso_descarga)
        self.worker_descarga.descarga_completada.connect(self._on_descarga_completada)
        self.worker_descarga.error_descarga.connect(self._on_error_descarga)
        self.worker_descarga.start()

    def _on_progreso_descarga(self, bytes_act: int, bytes_tot: int, vel_bps: float, texto: str) -> None:
        if bytes_tot > 0:
            pct = int((bytes_act / bytes_tot) * 100)
            self.barra_progreso.setValue(min(100, max(0, pct)))
        self.lbl_estado_descarga.setText(f"Descargando: {texto}")

        if vel_bps > 0:
            if vel_bps >= 1024 * 1024:
                self.lbl_velocidad.setText(f"Velocidad: {vel_bps / (1024*1024):.1f} MB/s")
            else:
                self.lbl_velocidad.setText(f"Velocidad: {vel_bps / 1024:.0f} KB/s")
        else:
            self.lbl_velocidad.setText("")

    def _on_descarga_completada(self, ruta_instalador: str) -> None:
        self.ruta_instalador_descargado = ruta_instalador
        self.barra_progreso.setValue(100)
        self.lbl_estado_descarga.setText("✅ Descarga completa e integridad SHA-256 verificada.")
        self.lbl_velocidad.setText("")

        self.btn_accion.setText("Instalar y Reiniciar RENO")
        self.btn_accion.setEnabled(True)
        self.btn_accion.setStyleSheet(
            """
            QPushButton {
                padding: 10px 24px;
                background-color: #16A34A;
                color: #FFFFFF;
                border: none;
                border-radius: 6px;
                font-weight: bold;
                font-size: 10.5pt;
            }
            QPushButton:hover {
                background-color: #15803D;
            }
            """
        )

    def _on_error_descarga(self, error_msg: str) -> None:
        self.lbl_estado_descarga.setText("❌ Error en la descarga.")
        self.lbl_velocidad.setText("")
        self.btn_accion.setText("Reintentar Descarga")
        self.btn_accion.setEnabled(True)
        self.btn_accion.setStyleSheet(
            """
            QPushButton {
                padding: 8px 20px;
                background-color: #2563EB;
                color: #FFFFFF;
                border: none;
                border-radius: 6px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #1D4ED8;
            }
            """
        )
        if hasattr(self, "btn_posponer"):
            self.btn_posponer.setEnabled(True)

        QMessageBox.warning(
            self,
            "Fallo de Actualización",
            f"No se pudo completar la descarga del instalador:\n\n{error_msg}\n\nVerifique su conexión de red o intente nuevamente.",
        )

    def _ejecutar_instalador(self) -> None:
        if not self.ruta_instalador_descargado:
            return

        exito = ejecutar_instalador_y_reiniciar(self.ruta_instalador_descargado)
        if not exito:
            QMessageBox.critical(
                self,
                "Error al Iniciar Instalador",
                f"No se pudo iniciar automáticamente el instalador en:\n{self.ruta_instalador_descargado}\n\n"
                "Por favor, cierre la aplicación y ejecute dicho archivo manualmente.",
            )

    def closeEvent(self, event) -> None:
        if self.worker_descarga and self.worker_descarga.isRunning():
            self.worker_descarga.cancelar()
            self.worker_descarga.wait(2000)

        if self.es_obligatoria and not self.ruta_instalador_descargado:
            # Si es obligatoria y el usuario intenta cerrar sin actualizar, rechazamos
            event.ignore()
            QMessageBox.warning(
                self,
                "Actualización Requerida",
                "Esta versión es de actualización obligatoria para garantizar la integridad operativa de los datos. "
                "Debe presionar 'Descargar e Instalar Ahora'.",
            )
            return

        super().closeEvent(event)
