"""
그리드 위젯 - 엑셀 데이터 표시 및 편집을 위한 QTableWidget
"""

from PySide6.QtWidgets import QTableWidget, QTableWidgetItem, QHeaderView
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont, QBrush, QColor

# 헤더 스타일 상수
HEADER_BG_COLOR = QColor(200, 220, 240)  # 연한 파란색 배경
HEADER_FONT_BOLD = True


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

        # 첫 번째 행(헤더)에 스타일 적용
        self.apply_header_row_style()

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

    def get_formatting(self):
        """
        그리드의 서식 정보 가져오기

        Returns:
            dict: {
                'fonts': {(row, col): {'bold': bool, 'size': int, 'name': str}},
                'colors': {(row, col): {'bg_color': str}}
            }
        """
        formatting = {
            'fonts': {},
            'colors': {}
        }

        for row in range(self.rowCount()):
            for col in range(self.columnCount()):
                item = self.item(row, col)
                if not item:
                    continue

                # 폰트 정보 추출
                font = item.font()
                if font.bold():
                    formatting['fonts'][(row, col)] = {
                        'bold': True,
                        'size': font.pointSize() if font.pointSize() > 0 else 10,
                        'name': font.family() if font.family() else '맑은 고딕'
                    }

                # 배경색 정보 추출
                bg_brush = item.background()
                if bg_brush.style() != 0:  # 0 = NoBrush (투명)
                    color = bg_brush.color()
                    # RGB를 HEX로 변환
                    hex_color = f"{color.red():02X}{color.green():02X}{color.blue():02X}"
                    formatting['colors'][(row, col)] = {
                        'bg_color': hex_color
                    }

        return formatting

    def clear_all(self):
        """
        그리드의 모든 데이터 삭제
        """
        self.clearContents()
        self.setRowCount(20)
        self.setColumnCount(10)

    def apply_header_row_style(self):
        """
        첫 번째 행(헤더)에 스타일 적용
        - 배경색: 연한 파란색
        - 폰트: Bold
        """
        if self.rowCount() == 0:
            return

        header_brush = QBrush(HEADER_BG_COLOR)

        for col in range(self.columnCount()):
            item = self.item(0, col)
            if item:
                # 배경색 적용
                item.setBackground(header_brush)
                # Bold 폰트 적용
                if HEADER_FONT_BOLD:
                    font = item.font()
                    font.setBold(True)
                    item.setFont(font)

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

        # 데이터 삽입 (Consolas 폰트 적용)
        paste_font = QFont("Consolas", 10)
        for row_offset, row_data in enumerate(rows_data):
            for col_offset, cell_value in enumerate(row_data):
                row_idx = current_row + row_offset
                col_idx = current_col + col_offset
                item = QTableWidgetItem(cell_value.strip())
                item.setFont(paste_font)
                self.setItem(row_idx, col_idx, item)

        # 첫 번째 행(헤더)에 스타일 적용
        if current_row == 0:
            self.apply_header_row_style()

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

    # === 최적화 적용 메서드 ===

    def apply_optimization(self, optimization_result):
        """
        최적화 결과를 그리드에 적용

        Args:
            optimization_result: optimizer에서 반환한 최적화 결과
        """
        # 1. 폰트 최적화 적용
        if 'font_optimization' in optimization_result:
            self._apply_font_optimization(optimization_result['font_optimization'])

        # 2. 빈 셀 최적화 적용
        if 'empty_cell_optimization' in optimization_result:
            self._apply_empty_cell_optimization(optimization_result['empty_cell_optimization'])

        # 3. Bold 최적화 적용
        if 'bold_optimization' in optimization_result:
            self._apply_bold_optimization(optimization_result['bold_optimization'])

        # 4. 헤더 줄바꿈 최적화 적용
        if 'header_wrap_optimization' in optimization_result:
            self._apply_header_wrap_optimization(optimization_result['header_wrap_optimization'])

        # 5. 레이아웃 최적화 적용
        if 'layout_optimization' in optimization_result:
            self._apply_layout_optimization(optimization_result['layout_optimization'])

    def _apply_font_optimization(self, font_opt):
        """
        폰트 최적화 적용
        """
        font_size = font_opt.get('default_font_size', 10)
        font_name = font_opt.get('default_font_name', '맑은 고딕')

        if font_opt.get('apply_to_all_cells'):
            self.apply_font_size(font_size)

    def _apply_empty_cell_optimization(self, empty_cell_opt):
        """
        빈 컬럼 최적화 적용 (데이터가 없는 컬럼)
        - 헤더를 여러 줄로 표시
        - 행 높이 자동 조정
        - 컬럼 너비 최소화
        - 헤더 제외 셀 배경색 제거
        """
        empty_columns = empty_cell_opt.get('empty_columns', {})
        max_lines = 1  # 최대 줄 수 추적

        for col_idx, col_info in empty_columns.items():
            if col_info.get('is_empty'):
                # 줄바꿈된 헤더 텍스트 가져오기
                wrapped_header = col_info.get('header_wrap', '')

                # 첫 번째 행(헤더)의 셀에 줄바꿈된 텍스트 적용
                header_item = self.item(0, col_idx)
                if header_item:
                    header_item.setText(wrapped_header)
                    # 텍스트 정렬 설정 (중앙 정렬)
                    header_item.setTextAlignment(Qt.AlignCenter)

                # 줄 수 계산
                line_count = wrapped_header.count('\n') + 1
                max_lines = max(max_lines, line_count)

                # 헤더 제외한 셀 배경색 제거 (채우기 없음)
                for row_idx in range(1, self.rowCount()):
                    cell_item = self.item(row_idx, col_idx)
                    if cell_item:
                        cell_item.setBackground(QBrush())  # 채우기 없음

        # 컬럼 너비 조정
        column_widths = empty_cell_opt.get('column_widths', {})
        for col_idx, width in column_widths.items():
            self.set_column_width(col_idx, width)

        # 첫 번째 행(헤더) 높이 조정 - 줄바꿈된 텍스트에 맞게
        if max_lines > 1:
            base_height = 25  # 기본 행 높이 (픽셀)
            new_height = base_height * max_lines
            self.setRowHeight(0, new_height)

    def _apply_bold_optimization(self, bold_opt):
        """
        Bold 최적화 적용
        """
        for col_idx, col_info in bold_opt.items():
            common_prefix = col_info.get('common_prefix', '')
            bold_length = col_info.get('bold_length', 0)
            affected_rows = col_info.get('affected_rows', [])

            # 각 셀의 공통 접두사 부분을 Bold 처리
            for row_idx in affected_rows:
                self._set_cell_partial_bold(row_idx, col_idx, bold_length)

    def _apply_header_wrap_optimization(self, header_wrap_opt):
        """
        헤더 줄바꿈 최적화 적용
        """
        for col_idx, col_info in header_wrap_opt.items():
            if col_info.get('wrap_text'):
                self._set_header_wrap(col_idx, True)

    def _set_cell_partial_bold(self, row, col, bold_length):
        """
        셀의 일부분만 Bold 처리 (앞부분)

        Args:
            row: 행 인덱스
            col: 열 인덱스
            bold_length: Bold 처리할 문자 수
        """
        item = self.item(row, col)
        if not item:
            return

        text = item.text()
        if len(text) < bold_length:
            return

        # 전체를 Bold로 설정하는 간단한 방법
        # (Qt의 Rich Text 제한으로 인해 부분 Bold는 복잡함)
        # 공통 접두사가 있는 셀은 Bold 처리
        font = item.font()
        font.setBold(True)
        item.setFont(font)

    def _set_header_wrap(self, col, wrap=True):
        """
        헤더 텍스트 줄바꿈 설정

        Args:
            col: 열 인덱스
            wrap: 줄바꿈 여부
        """
        header_item = self.horizontalHeaderItem(col)
        if header_item:
            # Qt에서는 헤더 아이템의 텍스트 줄바꿈이 자동으로 처리됨
            # 필요시 헤더 높이 조정
            if wrap:
                self.horizontalHeader().setDefaultSectionSize(60)  # 높이 증가

    def _apply_layout_optimization(self, layout_opt):
        """
        레이아웃 최적화 적용

        Args:
            layout_opt: 레이아웃 최적화 정보
        """
        # 1. 컬럼 너비 적용 (mm to pixel 변환)
        column_widths = layout_opt.get('column_widths', {})
        for col_idx, width_mm in column_widths.items():
            # mm를 픽셀로 변환 (96 DPI 기준: 1mm ≈ 3.78 pixels)
            width_px = int(width_mm * 3.78)
            self.set_column_width(col_idx, width_px)

        # 2. 행 높이 적용 (mm to pixel 변환)
        row_heights = layout_opt.get('row_heights', {})
        for row_idx, height_mm in row_heights.items():
            # mm를 픽셀로 변환
            height_px = int(height_mm * 3.78)

            if row_idx == -1:
                # 헤더 행 높이
                self.verticalHeader().setDefaultSectionSize(height_px)
            else:
                # 데이터 행 높이
                self.setRowHeight(row_idx, height_px)

        # 3. 페이지 분할 정보 표시 (선택사항)
        page_breaks = layout_opt.get('page_breaks', {})
        total_pages = page_breaks.get('total_pages', 1)
        break_points = page_breaks.get('break_points', [])

        # 페이지 분할 정보를 시각적으로 표시 (추후 확장 가능)
        # 현재는 로직만 구현하고 실제 표시는 생략
