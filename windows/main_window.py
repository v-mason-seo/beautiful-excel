from PySide6.QtWidgets import QMainWindow, QWidget, QHBoxLayout, QStackedWidget, QFrame
from widgets.sidebar import Sidebar
from widgets.sr_page import SRPage
from widgets.history_page import HistoryPage
from excel_loader import ExcelLoader
from config.app_config import AppConfig


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ExcelFlow")
        self.resize(1000, 700)
        self._center()

        self._loader = ExcelLoader()
        self._app_config = AppConfig()

        central = QWidget()
        self.setCentralWidget(central)
        layout = QHBoxLayout(central)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        self._sidebar = Sidebar()
        layout.addWidget(self._sidebar)

        divider = QFrame()
        divider.setFrameShape(QFrame.Shape.VLine)
        divider.setStyleSheet("color: #2E3450;")
        layout.addWidget(divider)

        self._stack = QStackedWidget()
        layout.addWidget(self._stack)

        self._sr_page = SRPage(self._loader, self._app_config)
        self._history_page = HistoryPage()
        self._stack.addWidget(self._sr_page)      # index 0
        self._stack.addWidget(self._history_page)  # index 1

        self._sidebar.btn_sr.clicked.connect(lambda: self._switch(0))
        self._sidebar.btn_history.clicked.connect(lambda: self._switch(1))

    def _switch(self, index: int):
        self._stack.setCurrentIndex(index)
        self._sidebar.select(index)

    def _center(self):
        screen = self.screen().availableGeometry()
        self.move(
            (screen.width() - self.width()) // 2,
            (screen.height() - self.height()) // 2,
        )

    def closeEvent(self, event):
        self._loader.quit()
        super().closeEvent(event)
