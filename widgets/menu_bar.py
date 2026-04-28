from PySide6.QtCore import Signal
from PySide6.QtWidgets import QComboBox, QHBoxLayout, QLabel, QPushButton, QWidget

from constants import CATEGORIES, CLR_ACCENT, CLR_SIDEBAR

_BTN_STYLE = f"""
QPushButton {{
    background-color: {CLR_ACCENT};
    color: #FFFFFF;
    border: none;
    border-radius: 4px;
    padding: 0 14px;
    font-size: 13px;
    font-weight: bold;
}}
QPushButton:hover {{ background-color: #5BA3F0; }}
QPushButton:pressed {{ background-color: #3A7BD5; }}
"""

_GHOST_BTN_STYLE = """
QPushButton {
    background-color: rgba(255,255,255,0.10);
    color: rgba(255,255,255,0.85);
    border: 1px solid rgba(255,255,255,0.25);
    border-radius: 4px;
    padding: 0 12px;
    font-size: 13px;
}
QPushButton:hover { background-color: rgba(255,255,255,0.18); }
QPushButton:pressed { background-color: rgba(255,255,255,0.08); }
"""


class MenuBar(QWidget):
    category_changed = Signal(str)
    save_requested = Signal()       # 저장 버튼 클릭
    settings_requested = Signal()   # 설정 버튼 클릭

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedHeight(48)
        self.setStyleSheet(f"background-color: {CLR_SIDEBAR};")

        layout = QHBoxLayout(self)
        layout.setContentsMargins(16, 0, 16, 0)
        layout.setSpacing(10)

        # ── 카테고리 콤보박스 ────────────────────────────────────────
        label = QLabel("카테고리:")
        label.setStyleSheet("color: rgba(255,255,255,0.7); font-size: 13px;")
        layout.addWidget(label)

        self.combo = QComboBox()
        self.combo.addItems(CATEGORIES)
        self.combo.setFixedWidth(160)
        self.combo.setFixedHeight(32)
        self.combo.currentTextChanged.connect(self.category_changed)
        layout.addWidget(self.combo)

        layout.addStretch()

        # ── 저장 버튼 ────────────────────────────────────────────────
        self.btn_save = QPushButton("💾  저장")
        self.btn_save.setFixedHeight(32)
        self.btn_save.setStyleSheet(_BTN_STYLE)
        self.btn_save.setToolTip("현재 데이터를 Excel 파일로 저장 (서식 유지)")
        self.btn_save.clicked.connect(self.save_requested)
        layout.addWidget(self.btn_save)

        # ── 설정 버튼 ────────────────────────────────────────────────
        self.btn_settings = QPushButton("⚙  설정")
        self.btn_settings.setFixedHeight(32)
        self.btn_settings.setStyleSheet(_GHOST_BTN_STYLE)
        self.btn_settings.setToolTip("매핑 / 저장 경로 / 옵션 설정")
        self.btn_settings.clicked.connect(self.settings_requested)
        layout.addWidget(self.btn_settings)

    def current_category(self) -> str:
        return self.combo.currentText()
