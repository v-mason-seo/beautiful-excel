from PySide6.QtWidgets import QWidget, QHBoxLayout, QComboBox, QLabel
from PySide6.QtCore import Signal
from constants import CLR_SIDEBAR, CATEGORIES


class MenuBar(QWidget):
    category_changed = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedHeight(48)
        self.setStyleSheet(f"background-color: {CLR_SIDEBAR};")

        layout = QHBoxLayout(self)
        layout.setContentsMargins(16, 0, 16, 0)

        label = QLabel("카테고리:")
        label.setStyleSheet("color: rgba(255,255,255,0.7); font-size: 13px;")
        layout.addWidget(label)

        self.combo = QComboBox()
        self.combo.addItems(CATEGORIES)
        self.combo.setFixedWidth(160)
        self.combo.currentTextChanged.connect(self.category_changed)
        layout.addWidget(self.combo)
        layout.addStretch()

    def current_category(self) -> str:
        return self.combo.currentText()
