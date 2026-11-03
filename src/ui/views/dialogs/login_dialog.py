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
)

from src.core.logger import get_logger
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
        self.setFixedSize(380, 240)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)

        layout = QVBoxLayout(self)
        layout.setSpacing(10)

        lbl_info = QLabel("Por políticas de seguridad, debe definir una nueva contraseña antes de ingresar al sistema.")
        lbl_info.setWordWrap(True)
        lbl_info.setStyleSheet("color: #4B5563; font-size: 9pt;")
        layout.addWidget(lbl_info)

        form = QFormLayout()
        self.txt_pass1 = QLineEdit()
        self.txt_pass1.setEchoMode(QLineEdit.Password)
        form.addRow("Nueva contraseña:", self.txt_pass1)

        self.txt_pass2 = QLineEdit()
        self.txt_pass2.setEchoMode(QLineEdit.Password)
        form.addRow("Confirmar contraseña:", self.txt_pass2)
        layout.addLayout(form)

        btn_box = QHBoxLayout()
        btn_box.addStretch()
        btn_ok = QPushButton("Cambiar Contraseña")
        btn_ok.clicked.connect(self._on_guardar)
        btn_box.addWidget(btn_ok)
        layout.addLayout(btn_box)

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
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(12)

        # Logotipo / Título
        lbl_logo = QLabel("RENO")
        lbl_logo.setAlignment(Qt.AlignCenter)
        lbl_logo.setStyleSheet("font-size: 26pt; font-weight: bold; color: #1F497D; letter-spacing: 2px;")
        layout.addWidget(lbl_logo)

        lbl_sub = QLabel("Registro de Novedades y Cambios de Turnos")
        lbl_sub.setAlignment(Qt.AlignCenter)
        lbl_sub.setStyleSheet("font-size: 8.5pt; color: #6B7280; margin-bottom: 8px;")
        layout.addWidget(lbl_sub)

        # Campos
        self.txt_username = QLineEdit()
        self.txt_username.setPlaceholderText("Nombre de usuario")
        layout.addWidget(self.txt_username)

        self.txt_password = QLineEdit()
        self.txt_password.setPlaceholderText("Contraseña")
        self.txt_password.setEchoMode(QLineEdit.Password)
        self.txt_password.returnPressed.connect(self._iniciar_sesion)
        layout.addWidget(self.txt_password)

        if self._es_primer_inicio:
            self.lbl_aviso_primer = QLabel("⚠️ No existen usuarios. Se creará el Administrador Inicial.")
            self.lbl_aviso_primer.setStyleSheet("color: #D97706; font-size: 8pt; font-weight: bold;")
            layout.addWidget(self.lbl_aviso_primer)

            self.txt_legajo = QLineEdit()
            self.txt_legajo.setPlaceholderText("Número de legajo personal")
            layout.addWidget(self.txt_legajo)

            self.btn_accion = QPushButton("Crear Administrador Inicial")
            self.btn_accion.clicked.connect(self._crear_primer_admin)
            layout.addWidget(self.btn_accion)
        else:
            self.btn_accion = QPushButton("Ingresar")
            self.btn_accion.clicked.connect(self._iniciar_sesion)
            layout.addWidget(self.btn_accion)

            # Enlace de olvido de clave
            btn_olvido = QPushButton("¿Olvidó su contraseña?")
            btn_olvido.setStyleSheet("""
                QPushButton {
                    background: transparent;
                    border: none;
                    color: #1F497D;
                    text-decoration: underline;
                    font-size: 8.5pt;
                }
                QPushButton:hover {
                    color: #2563EB;
                }
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
