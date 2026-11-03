"""Widget de notificación flotante no bloqueante (Toast) en la esquina inferior derecha."""

from PySide6.QtCore import QPoint, Qt, QTimer
from PySide6.QtWidgets import (
    QApplication,
    QFrame,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QVBoxLayout,
    QWidget,
)


class ToastWidget(QFrame):
    """Toast emergente no invasivo para notificaciones de nuevos registros o sincronización."""

    def __init__(self, parent=None):
        super().__init__(parent, Qt.ToolTip | Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setAttribute(Qt.WA_ShowWithoutActivating)
        self.setAttribute(Qt.WA_TranslucentBackground, False)

        self.setStyleSheet("""
            ToastWidget {
                background-color: #1F497D;
                color: #FFFFFF;
                border: 1px solid #16365C;
                border-radius: 8px;
            }
            QLabel {
                color: #FFFFFF;
            }
        """)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(14, 10, 14, 10)
        layout.setSpacing(10)

        self.lbl_icon = QLabel("🔔")
        self.lbl_icon.setStyleSheet("font-size: 14pt;")
        layout.addWidget(self.lbl_icon)

        v_text = QVBoxLayout()
        v_text.setSpacing(2)
        self.lbl_title = QLabel("RENO - Notificación")
        self.lbl_title.setStyleSheet("font-weight: bold; font-size: 9.5pt;")
        self.lbl_message = QLabel("")
        self.lbl_message.setStyleSheet("font-size: 8.5pt; color: #E0E7FF;")
        v_text.addWidget(self.lbl_title)
        v_text.addWidget(self.lbl_message)
        layout.addLayout(v_text)

        btn_close = QPushButton("✕")
        btn_close.setFixedSize(20, 20)
        btn_close.setStyleSheet("""
            QPushButton {
                background: transparent;
                border: none;
                color: #FFFFFF;
                font-size: 10pt;
                font-weight: bold;
            }
            QPushButton:hover {
                color: #EF4444;
            }
        """)
        btn_close.clicked.connect(self.hide)
        layout.addWidget(btn_close)

        self._timer = QTimer(self)
        self._timer.setSingleShot(True)
        self._timer.timeout.connect(self.hide)

    def show_message(self, title: str, message: str, duration_sec: int = 6) -> None:
        self.lbl_title.setText(title)
        self.lbl_message.setText(message)
        self.adjustSize()

        # Posicionar en la esquina inferior derecha de la pantalla activa
        screen = QApplication.primaryScreen().geometry()
        margin = 25
        x = screen.right() - self.width() - margin
        y = screen.bottom() - self.height() - margin
        self.move(QPoint(x, y))

        self.show()
        self._timer.start(duration_sec * 1000)
