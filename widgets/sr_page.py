from PySide6.QtWidgets import QWidget, QVBoxLayout, QTabWidget, QLabel
from PySide6.QtCore import Qt
from widgets.menu_bar import MenuBar
from widgets.common import PasteableTable
from constants import EXCEL_MAP


class SRPage(QWidget):
    def __init__(self, excel_loader, parent=None):
        super().__init__(parent)
        self._loader = excel_loader

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        self._menu_bar = MenuBar()
        layout.addWidget(self._menu_bar)

        self._content = QWidget()
        self._content_layout = QVBoxLayout(self._content)
        self._content_layout.setContentsMargins(16, 16, 16, 16)
        layout.addWidget(self._content)

        self._menu_bar.category_changed.connect(self._on_combo_changed)
        self._on_combo_changed(self._menu_bar.current_category())

    def _on_combo_changed(self, text: str):
        filepath = EXCEL_MAP.get(text)
        if not filepath:
            return
        try:
            sheet_data = self._loader.load(filepath)
            self._build_content(sheet_data)
        except Exception as e:
            self._show_error(str(e))

    def _build_content(self, sheet_data: dict):
        self._clear_content()

        if not sheet_data:
            return

        if len(sheet_data) == 1:
            name, rows = next(iter(sheet_data.items()))
            self._content_layout.addWidget(self._make_table(rows))
        else:
            tabs = QTabWidget()
            for sheet_name, rows in sheet_data.items():
                tabs.addTab(self._make_table(rows), sheet_name)
            self._content_layout.addWidget(tabs)

    def _make_table(self, rows: list) -> PasteableTable:
        table = PasteableTable()
        if rows:
            table.load_data(rows[0], rows[1:])
        return table

    def _show_error(self, message: str):
        self._clear_content()
        label = QLabel(f"파일을 불러올 수 없습니다:\n{message}")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setStyleSheet("color: #888; font-size: 13px; padding: 32px;")
        self._content_layout.addWidget(label)

    def _clear_content(self):
        while self._content_layout.count():
            child = self._content_layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()
