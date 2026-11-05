"""Widget de notificación flotante no bloqueante (Toast) en la esquina inferior derecha."""

from typing import Optional
from PySide6.QtCore import QPoint, Qt, QTimer, Signal
from PySide6.QtWidgets import (
    QApplication,
    QFrame,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QVBoxLayout,
)


class ToastWidget(QFrame):
    """Toast emergente no invasivo para notificaciones de nuevos registros o sincronización.
    
    Al hacer clic sobre el cuerpo de la notificación se emite la señal `clicked`,
    permitiendo abrir el formulario o detalle del registro notificado.
    """

    clicked = Signal()

    def __init__(self, parent=None):
        super().__init__(parent, Qt.ToolTip | Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setAttribute(Qt.WA_ShowWithoutActivating)
        self.setAttribute(Qt.WA_TranslucentBackground, False)
        self.setCursor(Qt.PointingHandCursor)

        self.setStyleSheet("""
            ToastWidget {
                background-color: #1F497D;
                color: #FFFFFF;
                border: 1.5px solid #3B82F6;
                border-radius: 8px;
            }
            ToastWidget:hover {
                background-color: #1A3F6D;
                border-color: #60A5FA;
            }
            QLabel {
                color: #FFFFFF;
                background: transparent;
            }
        """)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(14, 10, 14, 10)
        layout.setSpacing(10)

        self.lbl_icon = QLabel("🔔")
        self.lbl_icon.setStyleSheet("font-size: 14pt;")
        layout.addWidget(self.lbl_icon)

        v_text = QVBoxLayout()
        v_text.setSpacing(3)
        self.lbl_title = QLabel("RENO - Notificación")
        self.lbl_title.setStyleSheet("font-weight: bold; font-size: 9.5pt;")
        self.lbl_message = QLabel("")
        self.lbl_message.setStyleSheet("font-size: 8.5pt; color: #E0E7FF;")
        self.lbl_action = QLabel("👉 Clic para abrir el formulario")
        self.lbl_action.setStyleSheet("font-size: 8pt; color: #93C5FD; font-style: italic;")

        v_text.addWidget(self.lbl_title)
        v_text.addWidget(self.lbl_message)
        v_text.addWidget(self.lbl_action)
        layout.addLayout(v_text)

        self.btn_close = QPushButton("✕")
        self.btn_close.setFixedSize(22, 22)
        self.btn_close.setCursor(Qt.ArrowCursor)
        self.btn_close.setStyleSheet("""
            QPushButton {
                background: transparent;
                border: none;
                color: #FFFFFF;
                font-size: 10pt;
                font-weight: bold;
                border-radius: 4px;
            }
            QPushButton:hover {
                color: #FFFFFF;
                background-color: #EF4444;
            }
        """)
        self.btn_close.clicked.connect(self._on_close_clicked)
        layout.addWidget(self.btn_close, 0, Qt.AlignTop)

        self._on_click_callback = None
        self._timer = QTimer(self)
        self._timer.setSingleShot(True)
        self._timer.timeout.connect(self.hide)

    def _on_close_clicked(self) -> None:
        self._timer.stop()
        self.hide()

    def mousePressEvent(self, event) -> None:
        if event.button() == Qt.LeftButton:
            self._timer.stop()
            self.hide()
            if self._on_click_callback:
                self._on_click_callback()
            self.clicked.emit()
            event.accept()
            return
        super().mousePressEvent(event)

    def show_message(
        self,
        title: str,
        message: str,
        duration_sec: int = 6,
        action_text: Optional[str] = "👉 Clic para abrir el formulario",
        on_click=None,
    ) -> None:
        self.lbl_title.setText(title)
        self.lbl_message.setText(message)

        if action_text:
            self.lbl_action.setText(action_text)
            self.lbl_action.setVisible(True)
        else:
            self.lbl_action.setVisible(False)

        self._on_click_callback = on_click
        self.adjustSize()

        # Posicionar en la esquina inferior derecha de la pantalla activa
        screen = QApplication.primaryScreen().geometry()
        margin = 25
        x = screen.right() - self.width() - margin
        y = screen.bottom() - self.height() - margin
        self.move(QPoint(x, y))

        self.show()
        self._timer.start(duration_sec * 1000)
