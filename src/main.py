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

    # 메인 윈도우 생성 및 표시
    window = MainWindow()
    window.show()

    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
