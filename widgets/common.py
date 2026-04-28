from PySide6.QtWidgets import (
    QTableWidget, QTableWidgetItem, QAbstractItemView, QHeaderView, QApplication,
)
from PySide6.QtGui import QKeySequence
from constants import CLR_CARD, CLR_SIDEBAR, CLR_ROW_ALT


class PasteableTable(QTableWidget):
    """Ctrl+V 붙여넣기를 지원하는 QTableWidget 서브클래스."""

    def __init__(self, parent=None):
        super().__init__(0, 0, parent)
        self._apply_style()

    def _apply_style(self):
        self.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.setAlternatingRowColors(True)
        self.verticalHeader().setVisible(False)
        self.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.horizontalHeader().setStretchLastSection(True)
        self.setStyleSheet(f"""
            QTableWidget {{
                background-color: {CLR_CARD};
                gridline-color: #E0E4EF;
                border: none;
            }}
            QTableWidget::item:selected {{
                background-color: #D0E4FF;
                color: #1E2235;
            }}
            QHeaderView::section {{
                background-color: {CLR_SIDEBAR};
                color: #FFFFFF;
                padding: 6px 8px;
                border: none;
                font-weight: bold;
            }}
            QTableWidget::item:alternate {{
                background-color: {CLR_ROW_ALT};
            }}
        """)

    def keyPressEvent(self, event):
        if event.matches(QKeySequence.StandardKey.Paste):
            self._paste_from_clipboard()
        else:
            super().keyPressEvent(event)

    def _paste_from_clipboard(self):
        text = QApplication.clipboard().text()
        if not text:
            return

        # Alt+Enter 줄바꿈(\r)을 공백으로 치환한 뒤 행/셀 분리
        rows = text.rstrip("\n").split("\n")
        parsed = [row.replace("\r", " ").split("\t") for row in rows]

        num_rows = len(parsed)
        num_cols = max(len(r) for r in parsed)

        start_row = max(self.currentRow(), 0)
        start_col = max(self.currentColumn(), 0)

        needed_rows = start_row + num_rows
        if needed_rows > self.rowCount():
            self.setRowCount(needed_rows)

        needed_cols = start_col + num_cols
        if needed_cols > self.columnCount():
            self.setColumnCount(needed_cols)

        for r, row_data in enumerate(parsed):
            for c, cell in enumerate(row_data):
                self.setItem(start_row + r, start_col + c, QTableWidgetItem(cell))

    def load_data(self, headers: list, rows: list):
        """헤더와 행 데이터로 테이블을 채운다."""
        self.clear()
        self.setColumnCount(len(headers))
        self.setHorizontalHeaderLabels([str(h) if h is not None else "" for h in headers])
        self.setRowCount(len(rows))
        for r, row in enumerate(rows):
            for c, val in enumerate(row):
                text = "" if val is None else str(val)
                self.setItem(r, c, QTableWidgetItem(text))
