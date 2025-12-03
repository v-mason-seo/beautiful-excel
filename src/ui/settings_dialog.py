"""
설정 대화상자 - 용지, 방향, 글자 크기 설정
"""

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QGroupBox,
    QRadioButton, QSpinBox, QLabel, QPushButton,
    QButtonGroup
)
from PySide6.QtCore import Qt, Signal


class SettingsDialog(QDialog):
    """
    프로그램 설정을 위한 대화상자
    """

    settings_changed = Signal(dict)

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
        self.setWindowTitle("설정")
        self.setModal(True)
        self.setMinimumWidth(400)

        layout = QVBoxLayout()

        # 용지 크기 선택
        paper_group = QGroupBox("용지 크기")
        paper_layout = QVBoxLayout()

        self.paper_button_group = QButtonGroup(self)
        self.radio_a4 = QRadioButton("A4 (210mm × 297mm)")
        self.radio_a3 = QRadioButton("A3 (297mm × 420mm)")

        self.paper_button_group.addButton(self.radio_a4, 1)
        self.paper_button_group.addButton(self.radio_a3, 2)

        paper_layout.addWidget(self.radio_a4)
        paper_layout.addWidget(self.radio_a3)
        paper_group.setLayout(paper_layout)

        # 용지 방향 선택
        orientation_group = QGroupBox("용지 방향")
        orientation_layout = QVBoxLayout()

        self.orientation_button_group = QButtonGroup(self)
        self.radio_landscape = QRadioButton("가로 (Landscape)")
        self.radio_portrait = QRadioButton("세로 (Portrait)")

        self.orientation_button_group.addButton(self.radio_landscape, 1)
        self.orientation_button_group.addButton(self.radio_portrait, 2)

        orientation_layout.addWidget(self.radio_landscape)
        orientation_layout.addWidget(self.radio_portrait)
        orientation_group.setLayout(orientation_layout)

        # 글자 크기 선택
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
        font_layout.addStretch()
        font_group.setLayout(font_layout)

        # 버튼
        button_layout = QHBoxLayout()
        self.btn_ok = QPushButton("확인")
        self.btn_cancel = QPushButton("취소")

        self.btn_ok.clicked.connect(self.accept_settings)
        self.btn_cancel.clicked.connect(self.reject)

        button_layout.addStretch()
        button_layout.addWidget(self.btn_ok)
        button_layout.addWidget(self.btn_cancel)

        # 레이아웃 구성
        layout.addWidget(paper_group)
        layout.addWidget(orientation_group)
        layout.addWidget(font_group)
        layout.addLayout(button_layout)

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

    def accept_settings(self):
        """
        설정 적용
        """
        settings = self.get_settings()
        self.settings_changed.emit(settings)
        self.accept()
