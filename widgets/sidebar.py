from PySide6.QtWidgets import QWidget, QVBoxLayout, QPushButton, QLabel
from constants import CLR_SIDEBAR, CLR_ACCENT, APP_VERSION


class SidebarButton(QPushButton):
    """선택 상태에 따라 우측 accent 바를 표시하는 사이드바 버튼."""

    def __init__(self, icon: str, text: str, parent=None):
        super().__init__(parent)
        self.setText(f"  {icon}  {text}")
        self.setFixedHeight(48)
        self._set_selected(False)

    def set_selected(self, selected: bool):
        self._set_selected(selected)

    def _set_selected(self, selected: bool):
        if selected:
            self.setStyleSheet(f"""
                QPushButton {{
                    background-color: rgba(255, 255, 255, 0.12);
                    color: #FFFFFF;
                    border: none;
                    border-right: 3px solid {CLR_ACCENT};
                    text-align: left;
                    padding-left: 16px;
                    font-size: 14px;
                    font-weight: bold;
                }}
            """)
        else:
            self.setStyleSheet("""
                QPushButton {
                    background-color: transparent;
                    color: rgba(255, 255, 255, 0.7);
                    border: none;
                    text-align: left;
                    padding-left: 16px;
                    font-size: 14px;
                }
                QPushButton:hover {
                    background-color: rgba(255, 255, 255, 0.07);
                    color: #FFFFFF;
                }
            """)


class Sidebar(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedWidth(180)
        self.setStyleSheet(f"background-color: {CLR_SIDEBAR};")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        logo = QLabel("◈ ExcelFlow")
        logo.setStyleSheet(
            "color: #FFFFFF; font-size: 18px; font-weight: bold;"
            " padding: 24px 16px 20px 16px;"
        )
        layout.addWidget(logo)

        self.btn_sr = SidebarButton("📋", "SR")
        self.btn_history = SidebarButton("🕓", "History")
        layout.addWidget(self.btn_sr)
        layout.addWidget(self.btn_history)

        layout.addStretch()

        version_label = QLabel(f"v{APP_VERSION}")
        version_label.setStyleSheet(
            "color: rgba(255,255,255,0.4); font-size: 11px; padding: 12px 16px;"
        )
        layout.addWidget(version_label)

        self._buttons = [self.btn_sr, self.btn_history]
        self.select(0)

    def select(self, index: int):
        for i, btn in enumerate(self._buttons):
            btn.set_selected(i == index)
