"""
설정 패널 - 그리드 위에 배치되는 인라인 설정 컨트롤
"""

from PySide6.QtWidgets import (
    QWidget, QHBoxLayout, QGroupBox, QRadioButton, QSpinBox,
    QLabel, QPushButton, QButtonGroup
)
from PySide6.QtCore import Signal


class SettingsPanel(QWidget):
    """
    인라인 설정 패널 위젯
    """

    settings_changed = Signal(dict)
    optimization_requested = Signal()

    def __init__(self, parent=None, current_settings=None):
        super().__init__(parent)
        self.current_settings = current_settings or {
            'paper_size': 'A4',
            'orientation': 'landscape',
            'font_size': 10
        }
        self.setup_ui()
        self.load_settings()

    def setup_ui(self):
        """
        UI 구성
        """
        layout = QHBoxLayout()
        layout.setContentsMargins(5, 5, 5, 5)

        # 용지 크기 그룹
        paper_group = QGroupBox("용지 크기")
        paper_layout = QHBoxLayout()

        self.paper_button_group = QButtonGroup(self)
        self.radio_a4 = QRadioButton("A4")
        self.radio_a3 = QRadioButton("A3")

        self.paper_button_group.addButton(self.radio_a4, 1)
        self.paper_button_group.addButton(self.radio_a3, 2)

        paper_layout.addWidget(self.radio_a4)
        paper_layout.addWidget(self.radio_a3)
        paper_group.setLayout(paper_layout)

        # 용지 방향 그룹
        orientation_group = QGroupBox("용지 방향")
        orientation_layout = QHBoxLayout()

        self.orientation_button_group = QButtonGroup(self)
        self.radio_landscape = QRadioButton("가로")
        self.radio_portrait = QRadioButton("세로")

        self.orientation_button_group.addButton(self.radio_landscape, 1)
        self.orientation_button_group.addButton(self.radio_portrait, 2)

        orientation_layout.addWidget(self.radio_landscape)
        orientation_layout.addWidget(self.radio_portrait)
        orientation_group.setLayout(orientation_layout)

        # 글자 크기 그룹
        font_group = QGroupBox("글자 크기")
        font_layout = QHBoxLayout()

        font_label = QLabel("폰트 크기:")
        self.font_size_spinbox = QSpinBox()
        self.font_size_spinbox.setMinimum(8)
        self.font_size_spinbox.setMaximum(14)
        self.font_size_spinbox.setSuffix(" pt")
        self.font_size_spinbox.setValue(10)

        font_layout.addWidget(font_label)
        font_layout.addWidget(self.font_size_spinbox)
        font_group.setLayout(font_layout)

        # 적용 버튼
        self.btn_apply = QPushButton("설정 적용")
        self.btn_apply.setFixedWidth(100)
        self.btn_apply.clicked.connect(self.apply_settings)

        # 최적화 버튼
        self.btn_optimize = QPushButton("최적화 적용 (F5)")
        self.btn_optimize.setFixedWidth(150)
        self.btn_optimize.clicked.connect(self.request_optimization)

        # 레이아웃 구성
        layout.addWidget(paper_group)
        layout.addWidget(orientation_group)
        layout.addWidget(font_group)
        layout.addWidget(self.btn_apply)
        layout.addWidget(self.btn_optimize)
        layout.addStretch()

        self.setLayout(layout)

    def load_settings(self):
        """
        현재 설정 값을 UI에 로드
        """
        # 용지 크기
        if self.current_settings['paper_size'] == 'A4':
            self.radio_a4.setChecked(True)
        else:
            self.radio_a3.setChecked(True)

        # 용지 방향
        if self.current_settings['orientation'] == 'landscape':
            self.radio_landscape.setChecked(True)
        else:
            self.radio_portrait.setChecked(True)

        # 글자 크기
        self.font_size_spinbox.setValue(self.current_settings['font_size'])

    def get_settings(self):
        """
        현재 설정 값 가져오기

        Returns:
            dict: 설정 값 딕셔너리
        """
        settings = {
            'paper_size': 'A4' if self.radio_a4.isChecked() else 'A3',
            'orientation': 'landscape' if self.radio_landscape.isChecked() else 'portrait',
            'font_size': self.font_size_spinbox.value()
        }
        return settings

    def apply_settings(self):
        """
        설정 적용
        """
        settings = self.get_settings()
        self.current_settings = settings
        self.settings_changed.emit(settings)

    def request_optimization(self):
        """
        최적화 요청
        """
        self.optimization_requested.emit()
