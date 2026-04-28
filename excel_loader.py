try:
    import win32com.client as win32
    _WIN32_AVAILABLE = True
except ImportError:
    _WIN32_AVAILABLE = False


class ExcelLoader:
    """win32com으로 DRM Excel을 열고 시트/데이터를 읽는 헬퍼."""

    def __init__(self):
        self._excel = None

    def _ensure_excel(self):
        if not _WIN32_AVAILABLE:
            raise RuntimeError(
                "pywin32이 설치되어 있지 않습니다.\n"
                "CMD에서: pip install pywin32"
            )
        if self._excel is None:
            self._excel = win32.Dispatch("Excel.Application")
            self._excel.Visible = False
            self._excel.DisplayAlerts = False

    def load(self, filepath: str) -> dict:
        """
        반환 형태: {시트명: [[행데이터], ...]}
        헤더 포함 전체 UsedRange를 읽음.
        """
        self._ensure_excel()
        wb = self._excel.Workbooks.Open(filepath)
        result = {}
        try:
            for sheet in wb.Sheets:
                used = sheet.UsedRange
                if used is None:
                    result[sheet.Name] = []
                    continue
                raw = used.Value2
                if raw is None:
                    result[sheet.Name] = []
                elif not isinstance(raw[0], tuple):
                    result[sheet.Name] = [list(raw)]
                else:
                    result[sheet.Name] = [list(row) for row in raw]
        finally:
            wb.Close(SaveChanges=False)
        return result

    def quit(self):
        if self._excel:
            self._excel.Quit()
            self._excel = None
