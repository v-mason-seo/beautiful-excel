from typing import Dict, List, Optional

from PySide6.QtCore import Qt
from PySide6.QtGui import QFontMetrics, QKeySequence
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QHeaderView,
    QTableWidget,
    QTableWidgetItem,
)


class PasteableTable(QTableWidget):
    """
    Ctrl+V 붙여넣기를 지원하는 QTableWidget 서브클래스.

    개선 사항:
    - 스마트 붙여넣기: 헤더 매핑 기반 자동 컬럼 매핑 (set_smart_paste_config 주입)
    - 헤더 word-wrap 및 높이 동적 조정
    - 매핑 실패 시 기존 위치 기반 붙여넣기로 fallback
    """

    def __init__(self, parent=None):
        super().__init__(0, 0, parent)
        self.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.setAlternatingRowColors(True)
        self.verticalHeader().setVisible(False)

        hh = self.horizontalHeader()
        hh.setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        hh.setStretchLastSection(True)
        hh.setDefaultAlignment(Qt.AlignmentFlag.AlignCenter | Qt.TextWordWrap)
        hh.setMinimumHeight(36)
        try:
            hh.setWordWrap(True)
        except AttributeError:
            pass

        # 스마트 붙여넣기 설정 (SRPage 에서 카테고리 변경 시 주입)
        self._smart_paste_enabled: bool = False
        self._smart_paste_mapping: Dict[str, List[str]] = {}

    # ------------------------------------------------------------------
    # 스마트 붙여넣기 설정 주입
    # ------------------------------------------------------------------

    def set_smart_paste_config(self, enabled: bool, mapping: Dict[str, List[str]]):
        self._smart_paste_enabled = enabled
        self._smart_paste_mapping = mapping

    # ------------------------------------------------------------------
    # 키보드 이벤트
    # ------------------------------------------------------------------

    def keyPressEvent(self, event):
        if event.matches(QKeySequence.StandardKey.Paste):
            self._handle_paste()
        else:
            super().keyPressEvent(event)

    # ------------------------------------------------------------------
    # 붙여넣기 처리
    # ------------------------------------------------------------------

    def _handle_paste(self):
        text = QApplication.clipboard().text()
        if not text:
            return

        if self._smart_paste_enabled and self._smart_paste_mapping:
            used = self._smart_paste(text)
            if used:
                return
        self._positional_paste(text)

    def _smart_paste(self, text: str) -> bool:
        """
        헤더 매핑 기반 스마트 붙여넣기.
        그리드 헤더(horizontalHeader 레이블)를 대상으로 매핑.

        Returns:
            True  = 매핑 적용 성공
            False = 매핑 없음, caller 가 positional paste 로 fallback
        """
        from core.smart_paste import SmartPasteEngine

        target_headers = [
            (self.horizontalHeaderItem(c).text() if self.horizontalHeaderItem(c) else "")
            for c in range(self.columnCount())
        ]

        start_row = max(self.currentRow(), 0)
        cells, used_mapping = SmartPasteEngine.apply_smart_paste(
            text, target_headers, self._smart_paste_mapping, grid_start_row=start_row
        )

        if not cells or not used_mapping:
            return False

        max_row = max(r for r, c, v in cells) + 1
        max_col = max(c for r, c, v in cells) + 1
        if max_row > self.rowCount():
            self.setRowCount(max_row)
        if max_col > self.columnCount():
            self.setColumnCount(max_col)

        for row, col, val in cells:
            self.setItem(row, col, QTableWidgetItem(val))

        return True

    def _positional_paste(self, text: str):
        """기존 위치 기반 붙여넣기."""
        rows = text.rstrip("\n").split("\n")
        parsed = [row.replace("\r", " ").split("\t") for row in rows if row.strip() or "\t" in row]
        if not parsed:
            return

        num_rows = len(parsed)
        num_cols = max(len(r) for r in parsed)
        start_row = max(self.currentRow(), 0)
        start_col = max(self.currentColumn(), 0)

        if start_row + num_rows > self.rowCount():
            self.setRowCount(start_row + num_rows)
        if start_col + num_cols > self.columnCount():
            self.setColumnCount(start_col + num_cols)

        for r, row_data in enumerate(parsed):
            for c, cell in enumerate(row_data):
                self.setItem(start_row + r, start_col + c, QTableWidgetItem(cell))

    # ------------------------------------------------------------------
    # 데이터 로드 / 추출
    # ------------------------------------------------------------------

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
        self._adjust_header_height()

    def get_all_data(self) -> List[List[str]]:
        """현재 테이블의 전체 데이터를 2D 리스트로 반환."""
        return [
            [
                (self.item(r, c).text() if self.item(r, c) else "")
                for c in range(self.columnCount())
            ]
            for r in range(self.rowCount())
        ]

    # ------------------------------------------------------------------
    # 헤더 높이 동적 조정
    # ------------------------------------------------------------------

    def _adjust_header_height(self):
        hh = self.horizontalHeader()
        fm = QFontMetrics(hh.font())
        max_lines = 1
        for col in range(self.columnCount()):
            item = self.horizontalHeaderItem(col)
            if not item:
                continue
            text = item.text()
            col_width = self.columnWidth(col)
            if col_width > 0:
                text_width = fm.horizontalAdvance(text)
                if text_width > col_width - 8:
                    lines = max(1, -(-text_width // (col_width - 8)))
                else:
                    lines = text.count("\n") + 1
                max_lines = max(max_lines, lines)
        hh.setMinimumHeight(max(36, max_lines * 18 + 8))
