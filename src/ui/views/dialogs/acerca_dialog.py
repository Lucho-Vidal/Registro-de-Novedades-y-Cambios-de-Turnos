"""Diálogo modal de información del sistema y changelog (CRONO-PCT Standard)."""

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog,
    QFormLayout,
    QFrame,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QScrollArea,
    QTabWidget,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)


class AcercaDialog(QDialog):
    """Ventana modal informativa con detalles técnicos de arquitectura y changelog."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Acerca de RENO")
        self.resize(680, 520)
        self.setWindowFlags((self.windowFlags() & ~Qt.WindowContextHelpButtonHint) | Qt.WindowCloseButtonHint)

        self._init_ui()

    def _init_ui(self) -> None:
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(18, 18, 18, 18)
        main_layout.setSpacing(14)

        # Encabezado
        header_frame = QFrame()
        header_frame.setObjectName("cardFrame")
        header_layout = QHBoxLayout(header_frame)
        header_layout.setContentsMargins(16, 12, 16, 12)

        lbl_icon = QLabel("🚆")
        lbl_icon.setStyleSheet("font-size: 34pt;")
        header_layout.addWidget(lbl_icon)

        title_layout = QVBoxLayout()
        lbl_app = QLabel("<b>RENO</b> — Registro de Novedades y Cambios de Turnos")
        lbl_app.setStyleSheet("font-size: 13pt; font-weight: 700; color: #1E3A8A;")
        lbl_sub = QLabel("Sistema de Gestión Operativa Ferroviaria y Control de Dotación")
        lbl_sub.setStyleSheet("font-size: 9.5pt; color: #64748B;")
        lbl_ver = QLabel("<b>Versión 2.0.0</b> • <i>Estándar Arquitectónico CRONO-PCT</i>")
        lbl_ver.setStyleSheet("font-size: 9.5pt; color: #0284C7;")

        title_layout.addWidget(lbl_app)
        title_layout.addWidget(lbl_sub)
        title_layout.addWidget(lbl_ver)
        header_layout.addLayout(title_layout)
        header_layout.addStretch()

        main_layout.addWidget(header_frame)

        # Pestañas
        tabs = QTabWidget()

        # Tab 1: Arquitectura y Sistema
        tab_info = QWidget()
        lay_info = QVBoxLayout(tab_info)
        lay_info.setContentsMargins(14, 14, 14, 14)

        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignRight)
        form.setSpacing(10)

        form.addRow("<b>Versión Runtime:</b>", QLabel("RENO v2.0.0 Enterprise"))
        form.addRow("<b>Framework GUI:</b>", QLabel("PySide6-Essentials (Qt 6.7+ Python Native)"))
        form.addRow("<b>Persistencia Híbrida:</b>", QLabel("SQLite WAL / DELETE Local-First + Outbox Queue"))
        form.addRow("<b>Sincronización Nube:</b>", QLabel("PostgreSQL 15+ REST (Supabase Engine, esquema reno_*)"))
        form.addRow("<b>Motor de Reportes:</b>", QLabel("ReportLab 4.x (PDF Ejecutivo) + OpenPyXL 3.1+ (Excel)"))
        form.addRow("<b>Seguridad & RBAC:</b>", QLabel("Bcrypt Hashing + Permisos Granulares por Módulo"))
        form.addRow("<b>Reglas CCT:</b>", QLabel("Convenio Colectivo Ferroviario (Descansos 42hs/41hs y 18hs/17hs)"))

        lay_info.addLayout(form)
        lay_info.addStretch()
        tabs.addTab(tab_info, "ℹ️ Información del Sistema")

        # Tab 2: Historial de Versiones (Changelog)
        tab_change = QWidget()
        lay_change = QVBoxLayout(tab_change)
        lay_change.setContentsMargins(10, 10, 10, 10)

        browser = QTextBrowser()
        browser.setOpenExternalLinks(True)
        browser.setHtml("""
        <style>
            h3 { color: #1E3A8A; margin-top: 12px; margin-bottom: 4px; font-size: 11pt; }
            h4 { color: #0284C7; margin-top: 8px; margin-bottom: 2px; font-size: 10pt; }
            ul { margin-top: 2px; margin-bottom: 8px; padding-left: 20px; }
            li { margin-bottom: 3px; font-size: 9pt; }
            .badge { background-color: #E0F2FE; color: #0369A1; padding: 2px 6px; border-radius: 4px; font-weight: bold; }
        </style>

        <h3>Versión 2.0.0 <span class="badge">ACTUAL</span></h3>
        <h4>Reorganización y Estándar CRONO-PCT:</h4>
        <ul>
            <li><b>Nueva Navegación por Pestañas Integradas:</b> Panel de Control, Buscador de Turnos, Novedades y Cambios, Turnos Sin Cubrir, Empleados (BASE), Personal de Ciclo y Administración.</li>
            <li><b>Sub-pestañas operativas:</b> Novedades y Cambios agrupados en contenedor unificado; Administración organizada en Usuarios/Roles, Catálogos Operativos, Seguridad/Auditoría y Configuración.</li>
            <li><b>Turnos Operativos Sin Cubrir:</b> Nueva vista independiente con resolución automática de turno del día (ej. 401.1), hora toma/deja y seguimiento de reemplazos para personal de ciclo.</li>
            <li><b>Buscador de Turnos e Itinerarios:</b> Consulta en tiempo real de trenes, turnos titulares, relevos activos y cambios de turnos (permutas) sin demoras.</li>
            <li><b>Catálogos Operativos CRUD:</b> Gestión completa (crear, modificar y activación/desactivación soft delete) con control de permisos RBAC para Tipos de Novedad, Dotaciones y Personal de Estación.</li>
            <li><b>Validaciones de Integridad:</b> Detección y bloqueo automático de solapamiento de novedades activas y validación de disponibilidad de reemplazos.</li>
            <li><b>Tema Oscuro Moderno:</b> Eliminación total de bloqueos al cambiar de tema mediante columnas interactivas (optimización de rendimiento 70x).</li>
            <li><b>Calendarios y Selectores:</b> Días claramente visibles y legibles en todos los selectores de fecha en modo claro y oscuro.</li>
            <li><b>Exportación de Calendario de Ciclo:</b> Nuevas exportaciones a formato Excel XLSX individual y masivo para personal de ciclo.</li>
        </ul>

        <h3>Versión 1.9.0</h3>
        <ul>
            <li>Sincronización transaccional fuera de línea (Outbox Queue) con Supabase PostgreSQL.</li>
            <li>Gestión de reemplazos y diagramas para personal de ciclo con descansos CCT.</li>
            <li>Exportaciones de novedades y cambios en formatos PDF, Excel y Word.</li>
            <li>Auditoría y trazabilidad completa de modificaciones y papelera de reciclaje.</li>
        </ul>
        """)
        lay_change.addWidget(browser)
        tabs.addTab(tab_change, "📜 Historial de Cambios (Changelog)")

        main_layout.addWidget(tabs)

        # Botón Cerrar
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        btn_cerrar = QPushButton("Cerrar")
        btn_cerrar.setObjectName("btnPrimary")
        btn_cerrar.clicked.connect(self.accept)
        btn_layout.addWidget(btn_cerrar)

        main_layout.addLayout(btn_layout)
