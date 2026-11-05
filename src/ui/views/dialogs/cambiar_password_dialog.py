"""Diálogo modal para el cambio voluntario de contraseña del usuario en sesión."""

from typing import Any, Dict
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from src.core.theme import obtener_tema_visual
from src.services.auth_service import AuthService


class CambiarMiPasswordDialog(QDialog):
    """Permite al usuario cambiar su propia contraseña ingresando la actual y la nueva."""

    def __init__(self, auth_service: AuthService, usuario_actual: Dict[str, Any], parent=None):
        super().__init__(parent)
        self.auth_service = auth_service
        self.usuario_actual = usuario_actual

        self.setWindowTitle("Cambiar mi Contraseña")
        self.setMinimumWidth(420)
        self.setWindowFlags((self.windowFlags() & ~Qt.WindowContextHelpButtonHint) | Qt.WindowCloseButtonHint)
        self._init_ui()
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

    def _init_ui(self) -> None:
        es_oscuro = obtener_tema_visual() == "oscuro"
        color_user = "#60A5FA" if es_oscuro else "#1F497D"

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 16, 20, 16)
        layout.setSpacing(12)

        lbl_user = QLabel(f"Usuario: <b>{self.usuario_actual.get('username')}</b>")
        lbl_user.setStyleSheet(f"font-size: 10.5pt; color: {color_user};")
        layout.addWidget(lbl_user)

        form = QFormLayout()
        form.setSpacing(10)

        self.txt_actual = QLineEdit()
        self.txt_actual.setEchoMode(QLineEdit.Password)
        self.txt_actual.setPlaceholderText("Contraseña actual")
        form.addRow("Contraseña Actual:", self._crear_fila_password(self.txt_actual))

        self.txt_nueva = QLineEdit()
        self.txt_nueva.setEchoMode(QLineEdit.Password)
        self.txt_nueva.setPlaceholderText("Nueva contraseña")
        form.addRow("Nueva Contraseña:", self._crear_fila_password(self.txt_nueva))

        self.txt_confirm = QLineEdit()
        self.txt_confirm.setEchoMode(QLineEdit.Password)
        self.txt_confirm.setPlaceholderText("Confirmar nueva contraseña")
        form.addRow("Confirmar Nueva:", self._crear_fila_password(self.txt_confirm))

        self.txt_actual.returnPressed.connect(self.txt_nueva.setFocus)
        self.txt_nueva.returnPressed.connect(self.txt_confirm.setFocus)

        layout.addLayout(form)

        btn_box = QHBoxLayout()
        btn_box.addStretch()

        self.btn_cancel = QPushButton("Cancelar")
        self.btn_cancel.setObjectName("btnSecondary")
        self.btn_cancel.setAutoDefault(False)
        self.btn_cancel.setDefault(False)
        self.btn_cancel.clicked.connect(self.reject)

        self.btn_save = QPushButton("Guardar Contraseña")
        self.btn_save.setAutoDefault(True)
        self.btn_save.setDefault(True)
        self.btn_save.clicked.connect(self._on_guardar)

        btn_box.addWidget(self.btn_cancel)
        btn_box.addWidget(self.btn_save)
        layout.addLayout(btn_box)

    def _on_guardar(self) -> None:
        act = self.txt_actual.text()
        nueva = self.txt_nueva.text()
        conf = self.txt_confirm.text()

        if not act or not nueva or not conf:
            QMessageBox.warning(self, "Validación", "Complete todos los campos.")
            return

        if nueva != conf:
            QMessageBox.warning(self, "Validación", "Las contraseñas nuevas no coinciden.")
            return

        try:
            user_id = int(self.usuario_actual.get("id") or 0)
            username = str(self.usuario_actual.get("username") or "")
            self.auth_service.cambiar_mi_password(user_id, act, nueva, username=username)
            QMessageBox.information(self, "Éxito", "Su contraseña ha sido modificada exitosamente.")
            self.accept()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo cambiar la contraseña:\n{e}")
