"""
인쇄 엔진 - 엑셀 데이터를 프린터로 출력하는 기능
"""

from PySide6.QtPrintSupport import QPrinter, QPrintPreviewDialog, QPrintDialog
from PySide6.QtGui import QPainter, QFont, QColor, QPen, QPageSize, QPageLayout
from PySide6.QtCore import QRectF, Qt, QMarginsF
from PySide6.QtWidgets import QMessageBox


class PrintEngine:
    """
    엑셀 데이터 인쇄 엔진
    """

    def __init__(self, settings):
        """
        Args:
            settings: 용지 설정 딕셔너리 (paper_size, orientation, font_size)
        """
        self.settings = settings

    def print_preview(self, parent, data, headers):
        """
        인쇄 미리보기 다이얼로그 표시

        Args:
            parent: 부모 위젯
            data: 2차원 리스트 형태의 데이터
            headers: 헤더 리스트
        """
        try:
            # QPrinter 생성
            printer = self._create_printer()

            # 미리보기 다이얼로그
            preview_dialog = QPrintPreviewDialog(printer, parent)
            preview_dialog.setWindowTitle("인쇄 미리보기 - Beautiful Excel")

            # 미리보기 렌더링 연결
            preview_dialog.paintRequested.connect(
                lambda p: self._render_page(p, data, headers)
            )

            # 다이얼로그 표시
            preview_dialog.exec()

        except Exception as e:
            QMessageBox.critical(
                parent,
                "오류",
                f"인쇄 미리보기 실패:\n{str(e)}"
            )

    def print_document(self, parent, data, headers):
        """
        인쇄 다이얼로그 표시 및 인쇄 실행

        Args:
            parent: 부모 위젯
            data: 2차원 리스트 형태의 데이터
            headers: 헤더 리스트

        Returns:
            bool: 인쇄 성공 여부
        """
        try:
            # QPrinter 생성
            printer = self._create_printer()

            # 인쇄 다이얼로그
            print_dialog = QPrintDialog(printer, parent)
            print_dialog.setWindowTitle("인쇄 - Beautiful Excel")

            # 사용자가 확인을 누른 경우
            if print_dialog.exec() == QPrintDialog.Accepted:
                # 실제 인쇄 수행
                self._render_page(printer, data, headers)
                return True
            else:
                return False

        except Exception as e:
            QMessageBox.critical(
                parent,
                "오류",
                f"인쇄 실패:\n{str(e)}"
            )
            return False

    def _create_printer(self):
        """
        설정에 따라 QPrinter 객체 생성

        Returns:
            QPrinter: 설정이 적용된 프린터 객체
        """
        printer = QPrinter(QPrinter.HighResolution)

        # 용지 크기 설정
        paper_size = self.settings.get('paper_size', 'A4')
        if paper_size == 'A4':
            page_size = QPageSize(QPageSize.A4)
        elif paper_size == 'A3':
            page_size = QPageSize(QPageSize.A3)
        else:
            page_size = QPageSize(QPageSize.A4)

        # 용지 방향 설정
        orientation = self.settings.get('orientation', 'landscape')
        if orientation == 'landscape':
            page_orientation = QPageLayout.Landscape
        else:
            page_orientation = QPageLayout.Portrait

        # 페이지 레이아웃 설정 (여백 10mm)
        margins = QMarginsF(10, 10, 10, 10)
        page_layout = QPageLayout(page_size, page_orientation, margins, QPageLayout.Millimeter)
        printer.setPageLayout(page_layout)

        return printer

    def _render_page(self, printer, data, headers):
        """
        페이지에 데이터 렌더링

        Args:
            printer: QPrinter 객체
            data: 2차원 리스트 형태의 데이터
            headers: 헤더 리스트
        """
        if not data or not headers:
            return

        painter = QPainter()
        painter.begin(printer)

        try:
            # 폰트 설정
            font_size = self.settings.get('font_size', 10)
            font = QFont("맑은 고딕", font_size)
            painter.setFont(font)

            # 페이지 영역 계산
            page_rect = printer.pageRect(QPrinter.DevicePixel)
            content_width = page_rect.width()
            content_height = page_rect.height()

            # 컬럼 너비 계산
            num_columns = len(headers)
            column_width = content_width / num_columns

            # 행 높이 계산
            font_metrics = painter.fontMetrics()
            row_height = font_metrics.height() * 1.5

            # 현재 Y 위치
            current_y = 0

            # 헤더 그리기
            header_font = QFont("맑은 고딕", font_size, QFont.Bold)
            painter.setFont(header_font)
            painter.setPen(QPen(QColor(0, 0, 0)))

            for col_idx, header in enumerate(headers):
                x = col_idx * column_width
                rect = QRectF(x, current_y, column_width, row_height)
                painter.drawRect(rect)
                painter.drawText(
                    rect,
                    Qt.AlignCenter | Qt.TextWordWrap,
                    str(header)
                )

            current_y += row_height

            # 데이터 그리기
            data_font = QFont("맑은 고딕", font_size)
            painter.setFont(data_font)

            for row_idx, row_data in enumerate(data):
                # 페이지 넘침 확인
                if current_y + row_height > content_height:
                    printer.newPage()
                    current_y = 0

                    # 새 페이지에 헤더 다시 그리기
                    painter.setFont(header_font)
                    for col_idx, header in enumerate(headers):
                        x = col_idx * column_width
                        rect = QRectF(x, current_y, column_width, row_height)
                        painter.drawRect(rect)
                        painter.drawText(
                            rect,
                            Qt.AlignCenter | Qt.TextWordWrap,
                            str(header)
                        )
                    current_y += row_height
                    painter.setFont(data_font)

                # 데이터 셀 그리기
                for col_idx, cell_value in enumerate(row_data):
                    x = col_idx * column_width
                    rect = QRectF(x, current_y, column_width, row_height)
                    painter.drawRect(rect)
                    painter.drawText(
                        rect,
                        Qt.AlignLeft | Qt.AlignVCenter | Qt.TextWordWrap,
                        str(cell_value) if cell_value else ""
                    )

                current_y += row_height

        finally:
            painter.end()

    def _calculate_pages(self, data, headers):
        """
        데이터를 기준으로 페이지 수 계산

        Args:
            data: 2차원 리스트 형태의 데이터
            headers: 헤더 리스트

        Returns:
            int: 예상 페이지 수
        """
        # 임시 프린터로 페이지 크기 계산
        temp_printer = self._create_printer()
        page_rect = temp_printer.pageRect(QPrinter.DevicePixel)

        # 행 높이 계산
        font_size = self.settings.get('font_size', 10)
        row_height = font_size * 1.5 * 3.78  # pt to px

        # 페이지당 행 수
        rows_per_page = int(page_rect.height() / row_height) - 1  # 헤더 제외

        if rows_per_page <= 0:
            rows_per_page = 1

        # 전체 페이지 수
        total_rows = len(data)
        total_pages = (total_rows + rows_per_page - 1) // rows_per_page

        return total_pages
