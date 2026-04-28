"""
설정 대화상자 - 매핑 JSON 편집 / 저장 경로 / 스마트 붙여넣기 옵션
"""

import json
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QCheckBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QTabWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)


class SettingsDialog(QDialog):
    """
    3탭 설정 대화상자.
      탭1: 헤더 매핑 JSON 편집
      탭2: 저장 경로
      탭3: 스마트 붙여넣기 옵션
    """

    def __init__(self, parent=None, app_config=None):
        super().__init__(parent)
        self._config = app_config
        self._setup_ui()
        self._load()

    # ------------------------------------------------------------------
    # UI
    # ------------------------------------------------------------------

    def _setup_ui(self):
        self.setWindowTitle("ExcelFlow 설정")
        self.setMinimumSize(680, 520)
        self.setModal(True)

        root = QVBoxLayout(self)

        self.tabs = QTabWidget()
        root.addWidget(self.tabs)
        self.tabs.addTab(self._tab_mapping(), "헤더 매핑")
        self.tabs.addTab(self._tab_path(), "저장 경로")
        self.tabs.addTab(self._tab_options(), "붙여넣기 옵션")

        btn_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btn_box.accepted.connect(self._on_ok)
        btn_box.rejected.connect(self.reject)
        root.addWidget(btn_box)

    # ── 탭1: 헤더 매핑 ───────────────────────────────────────────────

    def _tab_mapping(self) -> QWidget:
        w = QWidget()
        v = QVBoxLayout(w)

        desc = QLabel(
            "SR 타입별 헤더 매핑을 JSON 형식으로 편집합니다.\n"
            "형식: { \"SR타입\": { \"그리드헤더명\": [\"별칭1\", \"별칭2\"] } }"
        )
        desc.setWordWrap(True)
        v.addWidget(desc)

        self._mapping_edit = QTextEdit()
        self._mapping_edit.setFont(QFont("Courier New", 10))
        self._mapping_edit.setAcceptRichText(False)
        v.addWidget(self._mapping_edit)

        btn_row = QHBoxLayout()
        btn_val = QPushButton("JSON 유효성 검사")
        btn_val.clicked.connect(self._validate_mapping)
        btn_rst = QPushButton("기본값 복원")
        btn_rst.clicked.connect(self._reset_mapping)
        btn_row.addWidget(btn_val)
        btn_row.addWidget(btn_rst)
        btn_row.addStretch()
        v.addLayout(btn_row)

        self._mapping_status = QLabel("")
        v.addWidget(self._mapping_status)
        return w

    # ── 탭2: 저장 경로 ───────────────────────────────────────────────

    def _tab_path(self) -> QWidget:
        w = QWidget()
        v = QVBoxLayout(w)
        v.setSpacing(16)

        v.addWidget(QLabel("SR 결과 저장 기본 경로:"))
        row = QHBoxLayout()
        self._save_path_edit = QLineEdit()
        self._save_path_edit.setPlaceholderText(r"예: C:\sr_result")
        btn = QPushButton("폴더 선택...")
        btn.setFixedWidth(100)
        btn.clicked.connect(lambda: self._browse(self._save_path_edit))
        row.addWidget(self._save_path_edit)
        row.addWidget(btn)
        v.addLayout(row)

        v.addStretch()
        return w

    # ── 탭3: 붙여넣기 옵션 ──────────────────────────────────────────

    def _tab_options(self) -> QWidget:
        w = QWidget()
        v = QVBoxLayout(w)
        v.setSpacing(12)

        self._chk_smart = QCheckBox(
            "헤더 매핑 기반 스마트 붙여넣기 활성화\n"
            "(비활성화 시 위치 기반 붙여넣기 사용)"
        )
        v.addWidget(self._chk_smart)

        note = QLabel(
            "<small>복사한 표에 헤더 행이 있을 때 그리드 컬럼에 자동 매핑합니다.<br>"
            "헤더를 감지하지 못하면 자동으로 위치 기반으로 전환됩니다.</small>"
        )
        note.setWordWrap(True)
        v.addWidget(note)
        v.addStretch()
        return w

    # ------------------------------------------------------------------
    # 로드 / 저장
    # ------------------------------------------------------------------

    def _load(self):
        if not self._config:
            return
        self._mapping_edit.setPlainText(self._config.get_mapping_json())
        self._save_path_edit.setText(self._config.get("save_path", r"C:\sr_result"))
        self._chk_smart.setChecked(self._config.get("smart_paste_enabled", True))

    def _save(self) -> bool:
        if not self._config:
            return True

        json_text = self._mapping_edit.toPlainText().strip()
        if json_text:
            try:
                self._config.set_mapping_from_json(json_text)
            except (json.JSONDecodeError, ValueError) as e:
                QMessageBox.critical(self, "JSON 오류", f"매핑 JSON 형식 오류:\n{e}")
                self.tabs.setCurrentIndex(0)
                return False

        self._config.set("save_path", self._save_path_edit.text().strip())
        self._config.set("smart_paste_enabled", self._chk_smart.isChecked())
        self._config.save()
        return True

    # ------------------------------------------------------------------
    # 슬롯
    # ------------------------------------------------------------------

    def _on_ok(self):
        if self._save():
            self.accept()

    def _validate_mapping(self):
        text = self._mapping_edit.toPlainText().strip()
        try:
            parsed = json.loads(text)
            if not isinstance(parsed, dict):
                raise ValueError("최상위는 JSON 객체여야 합니다.")
            self._set_status(f"유효  |  SR 타입: {', '.join(parsed.keys())}", ok=True)
        except Exception as e:
            self._set_status(f"오류: {e}", ok=False)

    def _reset_mapping(self):
        reply = QMessageBox.question(
            self, "기본값 복원", "매핑을 기본값으로 초기화하시겠습니까?",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No,
        )
        if reply != QMessageBox.Yes:
            return
        from config.app_config import _DEFAULT_MAPPING_FILE
        if _DEFAULT_MAPPING_FILE.exists():
            self._mapping_edit.setPlainText(_DEFAULT_MAPPING_FILE.read_text(encoding="utf-8"))
            self._set_status("기본값으로 복원되었습니다.", ok=True)
        else:
            QMessageBox.warning(self, "경고", "기본 매핑 파일을 찾을 수 없습니다.")

    def _browse(self, edit: QLineEdit):
        folder = QFileDialog.getExistingDirectory(self, "폴더 선택", edit.text() or str(Path.home()))
        if folder:
            edit.setText(folder)

    def _set_status(self, msg: str, ok: bool):
        color = "green" if ok else "red"
        self._mapping_status.setStyleSheet(f"color: {color};")
        self._mapping_status.setText(msg)
