"""
메인 윈도우 - Beautiful Excel 프로그램의 메인 GUI
"""

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QMenuBar,
    QMenu, QFileDialog, QMessageBox, QStatusBar
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QAction

from .grid_widget import GridWidget
from .settings_panel import SettingsPanel


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
        엑셀 파일 로드 (Phase 3에서 구현)

        Args:
            file_path: 엑셀 파일 경로
        """
        self.current_file = file_path
        self.update_status(f"파일 로드: {file_path}")
        QMessageBox.information(
            self,
            "안내",
            "엑셀 파일 로드 기능은 Phase 3에서 구현됩니다.\n"
            f"선택된 파일: {file_path}"
        )

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
        엑셀 파일 저장 (Phase 3에서 구현)

        Args:
            file_path: 저장할 파일 경로
        """
        self.current_file = file_path
        self.update_status(f"파일 저장: {file_path}")
        QMessageBox.information(
            self,
            "안내",
            "엑셀 파일 저장 기능은 Phase 3에서 구현됩니다.\n"
            f"저장 경로: {file_path}"
        )

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
        최적화 로직 적용 (Phase 4에서 구현)
        """
        self.update_status("최적화 적용 중...")
        QMessageBox.information(
            self,
            "안내",
            "최적화 기능은 Phase 4에서 구현됩니다.\n\n"
            "구현 예정 기능:\n"
            "- 폰트/크기 일괄 변환\n"
            "- 빈 셀 최적화\n"
            "- 공통 텍스트 Bold 처리\n"
            "- 헤더 자동 줄바꿈"
        )
        self.update_status("준비")

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
