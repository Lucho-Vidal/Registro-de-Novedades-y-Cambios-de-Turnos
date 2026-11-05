"""Diálogo de autenticación, recuperación de contraseñas y asistente de primer inicio."""

from typing import Any, Dict, Optional
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog,
    QFormLayout,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from src.core.logger import get_logger
from src.core.theme import obtener_tema_visual
from src.services.auth_service import AuthService, CuentaBloqueadaError
from src.services.mail_service import MailService

logger = get_logger("LoginDialog")


class CambioPasswordObligatorioDialog(QDialog):
    """Diálogo modal forzoso para actualizar contraseñas vencidas o temporales."""

    def __init__(self, user_id: int, auth_service: AuthService, parent=None):
        super().__init__(parent)
        self.user_id = user_id
        self.auth_service = auth_service

        self.setWindowTitle("Actualización Obligatoria de Contraseña")
        self.setMinimumWidth(420)
        self.setWindowFlags((self.windowFlags() & ~Qt.WindowContextHelpButtonHint) | Qt.WindowCloseButtonHint)

        es_oscuro = obtener_tema_visual() == "oscuro"
        color_info = "#CBD5E1" if es_oscuro else "#4B5563"

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 16, 20, 16)
        layout.setSpacing(12)

        lbl_info = QLabel("Por políticas de seguridad, debe definir una nueva contraseña antes de ingresar al sistema.")
        lbl_info.setWordWrap(True)
        lbl_info.setStyleSheet(f"color: {color_info}; font-size: 9.5pt;")
        layout.addWidget(lbl_info)

        form = QFormLayout()
        form.setSpacing(10)
        self.txt_pass1 = QLineEdit()
        self.txt_pass1.setEchoMode(QLineEdit.Password)
        self.txt_pass1.setPlaceholderText("Nueva contraseña")
        form.addRow("Nueva contraseña:", self._crear_fila_password(self.txt_pass1))

        self.txt_pass2 = QLineEdit()
        self.txt_pass2.setEchoMode(QLineEdit.Password)
        self.txt_pass2.setPlaceholderText("Confirmar nueva contraseña")
        form.addRow("Confirmar contraseña:", self._crear_fila_password(self.txt_pass2))

        self.txt_pass1.returnPressed.connect(self.txt_pass2.setFocus)
        layout.addLayout(form)

        btn_box = QHBoxLayout()
        btn_box.addStretch()
        self.btn_cancel = QPushButton("Cancelar")
        self.btn_cancel.setObjectName("btnSecondary")
        self.btn_cancel.setAutoDefault(False)
        self.btn_cancel.setDefault(False)
        self.btn_cancel.clicked.connect(self.reject)

        self.btn_ok = QPushButton("Cambiar Contraseña")
        self.btn_ok.setAutoDefault(True)
        self.btn_ok.setDefault(True)
        self.btn_ok.clicked.connect(self._on_guardar)

        btn_box.addWidget(self.btn_cancel)
        btn_box.addWidget(self.btn_ok)
        layout.addLayout(btn_box)
        self.adjustSize()

    @staticmethod
    def _crear_fila_password(line_edit: QLineEdit) -> QWidget:
        cont = QWidget()
        lay = QHBoxLayout(cont)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(6)
        lay.addWidget(line_edit, stretch=1)

        btn_ver = QPushButton("👁️")
        btn_ver.setObjectName("btnSecondary")
        btn_ver.setFixedWidth(36)
        btn_ver.setToolTip("Mostrar / ocultar contraseña")
        btn_ver.setAutoDefault(False)
        btn_ver.setDefault(False)
        btn_ver.setFocusPolicy(Qt.NoFocus)

        def _toggle() -> None:
            if line_edit.echoMode() == QLineEdit.Password:
                line_edit.setEchoMode(QLineEdit.Normal)
                btn_ver.setText("🔒")
            else:
                line_edit.setEchoMode(QLineEdit.Password)
                btn_ver.setText("👁️")

        btn_ver.clicked.connect(_toggle)
        lay.addWidget(btn_ver)
        return cont

    def _on_guardar(self) -> None:
        p1 = self.txt_pass1.text()
        p2 = self.txt_pass2.text()
        if not p1:
            QMessageBox.warning(self, "Validación", "La contraseña no puede estar vacía.")
            return
        if p1 != p2:
            QMessageBox.warning(self, "Validación", "Las contraseñas ingresadas no coinciden.")
            return

        try:
            self.auth_service.cambiar_password(self.user_id, p1)
            QMessageBox.information(self, "Éxito", "Contraseña actualizada exitosamente.")
            self.accept()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo cambiar la contraseña:\n{e}")


class LoginDialog(QDialog):
    """Pantalla de inicio de sesión con soporte para primer inicio y reseteo de clave."""

    def __init__(self, auth_service: AuthService, parent=None):
        super().__init__(parent)
        self.auth_service = auth_service
        self.usuario_autenticado: Optional[Dict[str, Any]] = None

        self.setWindowTitle("Acceso a RENO v2.0")
        self.setFixedSize(380, 360)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)

        self._es_primer_inicio = len(self.auth_service.listar_usuarios()) == 0
        self._init_ui()

    def _init_ui(self) -> None:
        es_oscuro = obtener_tema_visual() == "oscuro"
        color_logo = "#60A5FA" if es_oscuro else "#1F497D"
        color_sub = "#94A3B8" if es_oscuro else "#6B7280"
        color_olvido = "#60A5FA" if es_oscuro else "#1F497D"

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(12)

        # Logotipo / Título
        lbl_logo = QLabel("RENO")
        lbl_logo.setAlignment(Qt.AlignCenter)
        lbl_logo.setStyleSheet(f"font-size: 26pt; font-weight: bold; color: {color_logo}; letter-spacing: 2px;")
        layout.addWidget(lbl_logo)

        lbl_sub = QLabel("Registro de Novedades y Cambios de Turnos")
        lbl_sub.setAlignment(Qt.AlignCenter)
        lbl_sub.setStyleSheet(f"font-size: 8.5pt; color: {color_sub}; margin-bottom: 8px;")
        layout.addWidget(lbl_sub)

        # Campos
        self.txt_username = QLineEdit()
        self.txt_username.setPlaceholderText("Nombre de usuario")
        layout.addWidget(self.txt_username)

        self.txt_password = QLineEdit()
        self.txt_password.setPlaceholderText("Contraseña")
        self.txt_password.setEchoMode(QLineEdit.Password)
        self.txt_username.returnPressed.connect(self.txt_password.setFocus)
        layout.addWidget(self.txt_password)

        if self._es_primer_inicio:
            self.lbl_aviso_primer = QLabel("⚠️ No existen usuarios. Se creará el Administrador Inicial.")
            self.lbl_aviso_primer.setStyleSheet("color: #D97706; font-size: 8pt; font-weight: bold;")
            layout.addWidget(self.lbl_aviso_primer)

            self.txt_legajo = QLineEdit()
            self.txt_legajo.setPlaceholderText("Número de legajo personal")
            layout.addWidget(self.txt_legajo)

            self.btn_accion = QPushButton("Crear Administrador Inicial")
            self.btn_accion.setAutoDefault(True)
            self.btn_accion.setDefault(True)
            self.btn_accion.clicked.connect(self._crear_primer_admin)
            layout.addWidget(self.btn_accion)
        else:
            self.btn_accion = QPushButton("Ingresar")
            self.btn_accion.setAutoDefault(True)
            self.btn_accion.setDefault(True)
            self.btn_accion.clicked.connect(self._iniciar_sesion)
            layout.addWidget(self.btn_accion)

            # Enlace de olvido de clave
            btn_olvido = QPushButton("¿Olvidó su contraseña?")
            btn_olvido.setAutoDefault(False)
            btn_olvido.setDefault(False)
            btn_olvido.setStyleSheet(f"""
                QPushButton {{
                    background: transparent;
                    border: none;
                    color: {color_olvido};
                    text-decoration: underline;
                    font-size: 8.5pt;
                }}
                QPushButton:hover {{
                    color: #38BDF8;
                }}
            """)
            btn_olvido.clicked.connect(self._olvido_password)
            layout.addWidget(btn_olvido, alignment=Qt.AlignCenter)

    def _iniciar_sesion(self) -> None:
        user = self.txt_username.text().strip()
        pwd = self.txt_password.text()

        if not user or not pwd:
            QMessageBox.warning(self, "Acceso", "Ingrese usuario y contraseña.")
            return

        try:
            res = self.auth_service.autenticar(user, pwd)
            if not res:
                QMessageBox.critical(self, "Acceso Denegado", "Usuario o contraseña incorrectos.")
                return

            # Verificar si requiere cambio de contraseña
            if self.auth_service.requiere_cambio_clave(res["id"]):
                dlg = CambioPasswordObligatorioDialog(res["id"], self.auth_service, parent=self)
                if not dlg.exec():
                    return  # No se permite el acceso hasta cambiar la contraseña

            self.usuario_autenticado = res
            self.accept()
        except CuentaBloqueadaError as e:
            QMessageBox.critical(self, "Cuenta Bloqueada", str(e))
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Fallo al autenticar:\n{e}")

    def _crear_primer_admin(self) -> None:
        user = self.txt_username.text().strip()
        pwd = self.txt_password.text()
        leg_str = self.txt_legajo.text().strip()

        if not user or not pwd or not leg_str:
            QMessageBox.warning(self, "Validación", "Todos los campos son obligatorios.")
            return

        try:
            leg_int = int(leg_str)
        except ValueError:
            QMessageBox.warning(self, "Validación", "El legajo debe ser un número entero.")
            return

        try:
            self.auth_service.crear_administrador_inicial(user, pwd, legajo=leg_int)
            QMessageBox.information(self, "Éxito", "Administrador inicial creado exitosamente. Ahora puede iniciar sesión.")
            self.usuario_autenticado = self.auth_service.autenticar(user, pwd)
            self.accept()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo crear el administrador inicial:\n{e}")

    def _olvido_password(self) -> None:
        user = self.txt_username.text().strip()
        if not user:
            QMessageBox.information(self, "Recuperación", "Ingrese su nombre de usuario en el campo superior y presione nuevamente este enlace.")
            return

        resp = QMessageBox.question(
            self,
            "Recuperar Contraseña",
            f"Se enviará una contraseña temporal al correo institucional vinculado a '{user}'.\n¿Desea continuar?",
            QMessageBox.Yes | QMessageBox.No,
        )
        if resp != QMessageBox.Yes:
            return

        try:
            def enviar_outlook(dest, temp_pwd):
                asunto = "RENO - Recuperación de Contraseña"
                cuerpo = (
                    f"Estimado/a,\n\n"
                    f"Se ha solicitado el restablecimiento de su clave de acceso para el sistema RENO.\n\n"
                    f"Su contraseña temporal es: {temp_pwd}\n\n"
                    f"Al ingresar al sistema se le solicitará que defina una nueva contraseña definitiva.\n\n"
                    f"Atentamente,\nControl Operativo - RENO"
                )
                MailService.enviar_sin_rastro([dest], asunto, cuerpo)

            email, _ = self.auth_service.resetear_password(user, send_fn=enviar_outlook)
            QMessageBox.information(
                self,
                "Correo Enviado",
                f"Se ha enviado la contraseña temporal al correo {email} mediante Outlook.\n"
                f"Revise su bandeja de entrada e ingrese con la clave recibida.",
            )
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo restablecer la contraseña:\n{e}")
