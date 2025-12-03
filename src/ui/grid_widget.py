"""
그리드 위젯 - 엑셀 데이터 표시 및 편집을 위한 QTableWidget
"""

from PySide6.QtWidgets import QTableWidget, QTableWidgetItem, QHeaderView
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont


class GridWidget(QTableWidget):
    """
    엑셀 데이터를 표시하고 편집하는 그리드 위젯
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_ui()

    def setup_ui(self):
        """
        그리드 UI 초기 설정
        """
        # 기본 설정
        self.setRowCount(20)
        self.setColumnCount(10)

        # 헤더 설정
        self.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.verticalHeader().setSectionResizeMode(QHeaderView.Interactive)

        # 편집 가능 설정
        self.setEditTriggers(QTableWidget.DoubleClicked | QTableWidget.EditKeyPressed)

        # 선택 모드
        self.setSelectionMode(QTableWidget.ExtendedSelection)
        self.setSelectionBehavior(QTableWidget.SelectItems)

        # 그리드 라인 표시
        self.setShowGrid(True)

        # 기본 폰트
        default_font = QFont("맑은 고딕", 10)
        self.setFont(default_font)

    def set_data(self, data, headers=None):
        """
        그리드에 데이터 설정

        Args:
            data: 2차원 리스트 형태의 데이터
            headers: 헤더 리스트 (선택사항)
        """
        if not data:
            return

        rows = len(data)
        cols = len(data[0]) if data else 0

        self.setRowCount(rows)
        self.setColumnCount(cols)

        # 헤더 설정
        if headers:
            self.setHorizontalHeaderLabels(headers)

        # 데이터 삽입
        for row_idx, row_data in enumerate(data):
            for col_idx, cell_value in enumerate(row_data):
                item = QTableWidgetItem(str(cell_value) if cell_value is not None else "")
                self.setItem(row_idx, col_idx, item)

        # 컬럼 너비 자동 조정
        self.resizeColumnsToContents()

    def get_data(self):
        """
        그리드에서 데이터 가져오기

        Returns:
            2차원 리스트 형태의 데이터
        """
        data = []
        for row in range(self.rowCount()):
            row_data = []
            for col in range(self.columnCount()):
                item = self.item(row, col)
                row_data.append(item.text() if item else "")
            data.append(row_data)
        return data

    def get_headers(self):
        """
        헤더 데이터 가져오기

        Returns:
            헤더 리스트
        """
        headers = []
        for col in range(self.columnCount()):
            header_item = self.horizontalHeaderItem(col)
            headers.append(header_item.text() if header_item else f"Column {col + 1}")
        return headers

    def clear_all(self):
        """
        그리드의 모든 데이터 삭제
        """
        self.clearContents()
        self.setRowCount(20)
        self.setColumnCount(10)

    def apply_font_size(self, font_size):
        """
        모든 셀에 폰트 크기 적용

        Args:
            font_size: 폰트 크기 (pt)
        """
        font = QFont("맑은 고딕", font_size)

        for row in range(self.rowCount()):
            for col in range(self.columnCount()):
                item = self.item(row, col)
                if item:
                    item.setFont(font)

    def set_cell_bold(self, row, col, bold=True):
        """
        특정 셀을 Bold 처리

        Args:
            row: 행 인덱스
            col: 열 인덱스
            bold: Bold 여부
        """
        item = self.item(row, col)
        if item:
            font = item.font()
            font.setBold(bold)
            item.setFont(font)

    def set_column_bold(self, col, bold=True):
        """
        특정 컬럼의 모든 셀을 Bold 처리

        Args:
            col: 열 인덱스
            bold: Bold 여부
        """
        for row in range(self.rowCount()):
            self.set_cell_bold(row, col, bold)

    def set_header_font_size(self, col, font_size):
        """
        특정 컬럼의 헤더 폰트 크기 변경

        Args:
            col: 열 인덱스
            font_size: 폰트 크기 (pt)
        """
        header_item = self.horizontalHeaderItem(col)
        if header_item:
            font = header_item.font()
            font.setPointSize(font_size)
            header_item.setFont(font)

    def set_column_width(self, col, width):
        """
        특정 컬럼의 너비 설정

        Args:
            col: 열 인덱스
            width: 너비 (픽셀)
        """
        self.setColumnWidth(col, width)

    def paste_from_clipboard(self, clipboard_text):
        """
        클립보드 데이터를 그리드에 붙여넣기

        Args:
            clipboard_text: 클립보드 텍스트 (탭/줄바꿈 구분)
        """
        if not clipboard_text:
            return

        # 현재 선택된 셀 위치
        current_row = self.currentRow()
        current_col = self.currentColumn()

        if current_row < 0:
            current_row = 0
        if current_col < 0:
            current_col = 0

        # 탭과 줄바꿈으로 데이터 파싱
        lines = clipboard_text.split('\n')
        rows_data = []
        max_cols = 0

        for line in lines:
            if line.strip():
                row_data = line.split('\t')
                rows_data.append(row_data)
                max_cols = max(max_cols, len(row_data))

        # 그리드 크기 확장 (필요시)
        required_rows = current_row + len(rows_data)
        required_cols = current_col + max_cols

        if required_rows > self.rowCount():
            self.setRowCount(required_rows)
        if required_cols > self.columnCount():
            self.setColumnCount(required_cols)

        # 데이터 삽입
        for row_offset, row_data in enumerate(rows_data):
            for col_offset, cell_value in enumerate(row_data):
                row_idx = current_row + row_offset
                col_idx = current_col + col_offset
                item = QTableWidgetItem(cell_value.strip())
                self.setItem(row_idx, col_idx, item)

    def keyPressEvent(self, event):
        """
        키보드 이벤트 처리 (Ctrl+V 붙여넣기)
        """
        if event.key() == Qt.Key_V and event.modifiers() == Qt.ControlModifier:
            from PySide6.QtWidgets import QApplication
            clipboard = QApplication.clipboard()
            clipboard_text = clipboard.text()
            self.paste_from_clipboard(clipboard_text)
        else:
            super().keyPressEvent(event)
