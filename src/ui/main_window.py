"""
메인 윈도우 - Beautiful Excel 프로그램의 메인 GUI
"""

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QMenuBar,
    QMenu, QFileDialog, QMessageBox, QStatusBar, QToolBar
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QAction

from ui.grid_widget import GridWidget
from ui.settings_panel import SettingsPanel
from core.excel_loader import ExcelLoader
from core.exporter import ExcelExporter
from core.optimizer import ExcelOptimizer
from core.print_engine import PrintEngine


class MainWindow(QMainWindow):
    """
    Beautiful Excel 메인 윈도우
    """

    def __init__(self):
        super().__init__()
        self.settings = {
            'paper_size': 'A4',
            'orientation': 'landscape',
            'font_size': 10
        }
        self.current_file = None
        self.setup_ui()

    def setup_ui(self):
        """
        UI 구성
        """
        self.setWindowTitle("Beautiful Excel - 엑셀 출력 최적화")
        self.setGeometry(100, 100, 1200, 800)

        # 중앙 위젯 설정
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # 레이아웃
        layout = QVBoxLayout()
        central_widget.setLayout(layout)

        # 설정 패널 (그리드 위에 배치)
        self.settings_panel = SettingsPanel(self, self.settings)
        self.settings_panel.settings_changed.connect(self.apply_settings)
        self.settings_panel.optimization_requested.connect(self.apply_optimization)
        layout.addWidget(self.settings_panel)

        # 그리드 위젯
        self.grid_widget = GridWidget()
        layout.addWidget(self.grid_widget)

        # 메뉴바 생성
        self.create_menu_bar()

        # 툴바 생성
        self.create_toolbar()

        # 상태바
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.update_status("준비")

    def create_menu_bar(self):
        """
        메뉴바 생성
        """
        menubar = self.menuBar()

        # 파일 메뉴
        file_menu = menubar.addMenu("파일(&F)")

        # 열기
        open_action = QAction("열기(&O)", self)
        open_action.setShortcut("Ctrl+O")
        open_action.triggered.connect(self.open_file)
        file_menu.addAction(open_action)

        # 저장
        save_action = QAction("저장(&S)", self)
        save_action.setShortcut("Ctrl+S")
        save_action.triggered.connect(self.save_file)
        file_menu.addAction(save_action)

        # 다른 이름으로 저장
        save_as_action = QAction("다른 이름으로 저장(&A)", self)
        save_as_action.setShortcut("Ctrl+Shift+S")
        save_as_action.triggered.connect(self.save_file_as)
        file_menu.addAction(save_as_action)

        file_menu.addSeparator()

        # 종료
        exit_action = QAction("종료(&X)", self)
        exit_action.setShortcut("Ctrl+Q")
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        # 편집 메뉴
        edit_menu = menubar.addMenu("편집(&E)")

        # 붙여넣기
        paste_action = QAction("붙여넣기(&V)", self)
        paste_action.setShortcut("Ctrl+V")
        paste_action.triggered.connect(self.paste_data)
        edit_menu.addAction(paste_action)

        # 모두 지우기
        clear_action = QAction("모두 지우기(&C)", self)
        clear_action.triggered.connect(self.clear_all)
        edit_menu.addAction(clear_action)

        # 도움말 메뉴
        help_menu = menubar.addMenu("도움말(&H)")

        # 프로그램 정보
        about_action = QAction("프로그램 정보(&A)", self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)

    def create_toolbar(self):
        """
        툴바 생성 - 주요 기능 버튼 배치
        """
        toolbar = QToolBar("메인 툴바")
        toolbar.setMovable(False)
        self.addToolBar(toolbar)

        # 파일 열기
        open_action = QAction("열기", self)
        open_action.setShortcut("Ctrl+O")
        open_action.setToolTip("엑셀 파일 열기 (Ctrl+O)")
        open_action.triggered.connect(self.open_file)
        toolbar.addAction(open_action)

        # 저장
        save_action = QAction("저장", self)
        save_action.setShortcut("Ctrl+S")
        save_action.setToolTip("파일 저장 (Ctrl+S)")
        save_action.triggered.connect(self.save_file)
        toolbar.addAction(save_action)

        # 다른 이름으로 저장
        save_as_action = QAction("다른 이름으로 저장", self)
        save_as_action.setShortcut("Ctrl+Shift+S")
        save_as_action.setToolTip("다른 이름으로 저장 (Ctrl+Shift+S)")
        save_as_action.triggered.connect(self.save_file_as)
        toolbar.addAction(save_as_action)

        toolbar.addSeparator()

        # 붙여넣기
        paste_action = QAction("붙여넣기", self)
        paste_action.setShortcut("Ctrl+V")
        paste_action.setToolTip("클립보드 데이터 붙여넣기 (Ctrl+V)")
        paste_action.triggered.connect(self.paste_data)
        toolbar.addAction(paste_action)

        # 모두 지우기
        clear_action = QAction("모두 지우기", self)
        clear_action.setToolTip("그리드 데이터 모두 삭제")
        clear_action.triggered.connect(self.clear_all)
        toolbar.addAction(clear_action)

        toolbar.addSeparator()

        # 인쇄 미리보기
        preview_action = QAction("인쇄 미리보기", self)
        preview_action.setShortcut("Ctrl+Shift+P")
        preview_action.setToolTip("인쇄 미리보기 (Ctrl+Shift+P)")
        preview_action.triggered.connect(self.print_preview)
        toolbar.addAction(preview_action)

        # 인쇄
        print_action = QAction("인쇄", self)
        print_action.setShortcut("Ctrl+P")
        print_action.setToolTip("인쇄 (Ctrl+P)")
        print_action.triggered.connect(self.print_document)
        toolbar.addAction(print_action)

    def open_file(self):
        """
        엑셀 파일 열기
        """
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "엑셀 파일 열기",
            "",
            "Excel Files (*.xlsx *.xls);;All Files (*.*)"
        )

        if file_path:
            self.load_excel_file(file_path)

    def load_excel_file(self, file_path):
        """
        엑셀 파일 로드

        Args:
            file_path: 엑셀 파일 경로
        """
        try:
            self.update_status(f"파일 로드 중: {file_path}")

            # 엑셀 파일 로드
            result = ExcelLoader.load_file(file_path)

            data = result['data']
            headers = result['headers']
            formatting = result['formatting']

            # 그리드에 데이터 설정
            if data and len(data) > 1:
                # 첫 행은 헤더로 사용
                self.grid_widget.set_data(data[1:], headers)

                # 서식 정보 적용
                self._apply_formatting_to_grid(formatting)

                self.current_file = file_path
                self.update_status(f"파일 로드 완료: {file_path} ({len(data)-1}행, {len(headers)}열)")

                QMessageBox.information(
                    self,
                    "완료",
                    f"엑셀 파일을 성공적으로 불러왔습니다.\n\n"
                    f"파일: {file_path}\n"
                    f"행 수: {len(data)-1}\n"
                    f"열 수: {len(headers)}"
                )
            else:
                QMessageBox.warning(
                    self,
                    "경고",
                    "엑셀 파일에 데이터가 없습니다."
                )
                self.update_status("준비")

        except Exception as e:
            QMessageBox.critical(
                self,
                "오류",
                f"엑셀 파일 로드 실패:\n{str(e)}"
            )
            self.update_status("파일 로드 실패")

    def _apply_formatting_to_grid(self, formatting):
        """
        로드된 서식 정보를 그리드에 적용

        Args:
            formatting: 서식 정보 딕셔너리
        """
        # 폰트 서식 적용
        if 'fonts' in formatting:
            for (row, col), font_info in formatting['fonts'].items():
                if font_info.get('bold'):
                    self.grid_widget.set_cell_bold(row, col, True)

    def save_file(self):
        """
        현재 파일 저장
        """
        if self.current_file:
            self.save_excel_file(self.current_file)
        else:
            self.save_file_as()

    def save_file_as(self):
        """
        다른 이름으로 저장
        """
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "엑셀 파일 저장",
            "",
            "Excel Files (*.xlsx);;All Files (*.*)"
        )

        if file_path:
            self.save_excel_file(file_path)

    def save_excel_file(self, file_path):
        """
        엑셀 파일 저장

        Args:
            file_path: 저장할 파일 경로
        """
        try:
            self.update_status(f"파일 저장 중: {file_path}")

            # 그리드에서 데이터 가져오기
            data = self.grid_widget.get_data()
            formatting = self.grid_widget.get_formatting()

            # 빈 데이터 체크
            has_data = False
            for row in data:
                if any(cell.strip() for cell in row):
                    has_data = True
                    break

            if not has_data:
                QMessageBox.warning(
                    self,
                    "경고",
                    "저장할 데이터가 없습니다."
                )
                self.update_status("준비")
                return

            # 엑셀 파일로 저장 (첫 번째 행이 헤더로 처리됨)
            ExcelExporter.save_to_excel(
                file_path=file_path,
                data=data,
                headers=None,  # 데이터의 첫 번째 행을 헤더로 사용
                formatting=formatting,  # 서식 정보 전달
                settings=self.settings
            )

            self.current_file = file_path
            self.update_status(f"파일 저장 완료: {file_path}")

            QMessageBox.information(
                self,
                "완료",
                f"엑셀 파일을 성공적으로 저장했습니다.\n\n"
                f"저장 경로: {file_path}"
            )

        except Exception as e:
            QMessageBox.critical(
                self,
                "오류",
                f"엑셀 파일 저장 실패:\n{str(e)}"
            )
            self.update_status("파일 저장 실패")

    def paste_data(self):
        """
        클립보드에서 데이터 붙여넣기
        """
        from PySide6.QtWidgets import QApplication
        clipboard = QApplication.clipboard()
        clipboard_text = clipboard.text()

        if clipboard_text:
            self.grid_widget.paste_from_clipboard(clipboard_text)
            self.update_status("클립보드 데이터 붙여넣기 완료")
        else:
            QMessageBox.warning(self, "경고", "클립보드에 데이터가 없습니다.")

    def clear_all(self):
        """
        그리드의 모든 데이터 삭제
        """
        reply = QMessageBox.question(
            self,
            "확인",
            "모든 데이터를 삭제하시겠습니까?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            self.grid_widget.clear_all()
            self.current_file = None
            self.update_status("데이터 삭제 완료")

    def apply_settings(self, settings):
        """
        설정 적용

        Args:
            settings: 설정 딕셔너리
        """
        self.settings = settings
        self.update_status(
            f"설정 적용: {settings['paper_size']}, "
            f"{settings['orientation']}, "
            f"{settings['font_size']}pt"
        )

        # 글자 크기 즉시 적용
        self.grid_widget.apply_font_size(settings['font_size'])

    def apply_optimization(self):
        """
        최적화 로직 적용
        """
        try:
            self.update_status("최적화 적용 중...")

            # 그리드에서 데이터 가져오기
            data = self.grid_widget.get_data()

            # 데이터가 없으면 중단
            has_data = False
            for row in data:
                if any(cell.strip() for cell in row):
                    has_data = True
                    break

            if not has_data:
                QMessageBox.warning(
                    self,
                    "경고",
                    "최적화할 데이터가 없습니다."
                )
                self.update_status("준비")
                return

            # 데이터의 첫 번째 행을 헤더로 사용
            headers = data[0] if data else []

            # 최적화 실행
            optimizer = ExcelOptimizer(self.settings)
            optimization_result = optimizer.optimize(data, headers)

            # 최적화 결과를 그리드에 적용
            self.grid_widget.apply_optimization(optimization_result)

            # 최적화 결과 요약
            summary = self._create_optimization_summary(optimization_result)

            self.update_status("최적화 완료")

            QMessageBox.information(
                self,
                "최적화 완료",
                f"데이터 최적화가 완료되었습니다.\n\n{summary}"
            )

        except Exception as e:
            QMessageBox.critical(
                self,
                "오류",
                f"최적화 실패:\n{str(e)}"
            )
            self.update_status("최적화 실패")

    def _create_optimization_summary(self, optimization_result):
        """
        최적화 결과 요약 생성

        Args:
            optimization_result: 최적화 결과

        Returns:
            str: 요약 메시지
        """
        summary_parts = []

        # 폰트 최적화
        font_opt = optimization_result.get('font_optimization', {})
        if font_opt.get('apply_to_all_cells'):
            font_size = font_opt.get('default_font_size', 10)
            summary_parts.append(f"✓ 폰트 크기: {font_size}pt 일괄 적용")

        # 빈 셀 최적화
        empty_opt = optimization_result.get('empty_cell_optimization', {})
        empty_columns = empty_opt.get('empty_columns', {})
        if empty_columns:
            summary_parts.append(f"✓ 빈 셀 최적화: {len(empty_columns)}개 컬럼")

        # Bold 최적화
        bold_opt = optimization_result.get('bold_optimization', {})
        if bold_opt:
            total_cells = sum(len(info['affected_rows']) for info in bold_opt.values())
            summary_parts.append(f"✓ 공통 텍스트 Bold: {total_cells}개 셀")

        # 헤더 줄바꿈
        header_opt = optimization_result.get('header_wrap_optimization', {})
        if header_opt:
            summary_parts.append(f"✓ 헤더 줄바꿈: {len(header_opt)}개 컬럼")

        # 레이아웃 최적화
        layout_opt = optimization_result.get('layout_optimization', {})
        if layout_opt:
            page_breaks = layout_opt.get('page_breaks', {})
            total_pages = page_breaks.get('total_pages', 1)

            paper_size = self.settings.get('paper_size', 'A4')
            orientation = '가로' if self.settings.get('orientation') == 'landscape' else '세로'

            summary_parts.append(f"✓ 용지 설정: {paper_size} {orientation}")
            summary_parts.append(f"✓ 예상 페이지 수: {total_pages}페이지")

        if not summary_parts:
            return "적용된 최적화가 없습니다."

        return "\n".join(summary_parts)

    def print_preview(self):
        """
        인쇄 미리보기
        """
        # 그리드에서 데이터 가져오기
        data = self.grid_widget.get_data()
        headers = self.grid_widget.get_headers()

        # 데이터가 없으면 중단
        has_data = False
        for row in data:
            if any(cell.strip() for cell in row):
                has_data = True
                break

        if not has_data:
            QMessageBox.warning(
                self,
                "경고",
                "인쇄할 데이터가 없습니다."
            )
            return

        # 인쇄 엔진 생성 및 미리보기 표시
        print_engine = PrintEngine(self.settings)
        print_engine.print_preview(self, data, headers)
        self.update_status("인쇄 미리보기 완료")

    def print_document(self):
        """
        인쇄 다이얼로그 표시 및 인쇄
        """
        # 그리드에서 데이터 가져오기
        data = self.grid_widget.get_data()
        headers = self.grid_widget.get_headers()

        # 데이터가 없으면 중단
        has_data = False
        for row in data:
            if any(cell.strip() for cell in row):
                has_data = True
                break

        if not has_data:
            QMessageBox.warning(
                self,
                "경고",
                "인쇄할 데이터가 없습니다."
            )
            return

        # 인쇄 엔진 생성 및 인쇄 실행
        print_engine = PrintEngine(self.settings)
        success = print_engine.print_document(self, data, headers)

        if success:
            self.update_status("인쇄 완료")
            QMessageBox.information(
                self,
                "완료",
                "문서를 성공적으로 인쇄했습니다."
            )
        else:
            self.update_status("인쇄 취소")

    def show_about(self):
        """
        프로그램 정보 표시
        """
        QMessageBox.about(
            self,
            "프로그램 정보",
            "<h2>Beautiful Excel</h2>"
            "<p>버전: 0.1.0</p>"
            "<p>엑셀 데이터를 A4/A3 용지 규격에 맞춰<br>"
            "자동으로 정리 및 최적화하여 출력하는 프로그램</p>"
            "<p><b>개발 환경:</b> Python, PySide6</p>"
            "<p><b>작성:</b> Claude Code</p>"
        )

    def update_status(self, message):
        """
        상태바 메시지 업데이트

        Args:
            message: 표시할 메시지
        """
        self.status_bar.showMessage(message)

    def keyPressEvent(self, event):
        """
        키보드 이벤트 처리 (F5 단축키)
        """
        if event.key() == Qt.Key_F5:
            self.apply_optimization()
        else:
            super().keyPressEvent(event)
