"""
Beautiful Excel - 메인 진입점

엑셀 데이터를 A4/A3 용지 규격에 맞춰 자동으로 정리 및 최적화하는 프로그램
"""

import sys
from PySide6.QtWidgets import QApplication

from ui.main_window import MainWindow


def main():
    """
    프로그램 메인 함수
    """
    app = QApplication(sys.argv)
    app.setApplicationName("Beautiful Excel")
    app.setApplicationVersion("0.1.0")

    # 화이트 모드 스타일 적용
    app.setStyle("Fusion")

    # 라이트 팔레트 설정
    from PySide6.QtGui import QPalette, QColor
    palette = QPalette()
    palette.setColor(QPalette.Window, QColor(255, 255, 255))
    palette.setColor(QPalette.WindowText, QColor(0, 0, 0))
    palette.setColor(QPalette.Base, QColor(255, 255, 255))
    palette.setColor(QPalette.AlternateBase, QColor(245, 245, 245))
    palette.setColor(QPalette.ToolTipBase, QColor(255, 255, 255))
    palette.setColor(QPalette.ToolTipText, QColor(0, 0, 0))
    palette.setColor(QPalette.Text, QColor(0, 0, 0))
    palette.setColor(QPalette.Button, QColor(240, 240, 240))
    palette.setColor(QPalette.ButtonText, QColor(0, 0, 0))
    palette.setColor(QPalette.BrightText, QColor(255, 0, 0))
    palette.setColor(QPalette.Highlight, QColor(42, 130, 218))
    palette.setColor(QPalette.HighlightedText, QColor(255, 255, 255))
    app.setPalette(palette)

    # 메인 윈도우 생성 및 표시
    window = MainWindow()
    window.show()

    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
