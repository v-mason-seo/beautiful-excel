# ExcelFlow

Windows 데스크탑 앱. Excel/HTML 표 형태의 작업 데이터를 A4 인쇄용 Excel 템플릿에 자동으로 채워 넣는다.

## 기술 스택

- **Python 3.11+** (3.13 권장, 3.14 불가)
- **PySide6** — GUI 프레임워크
- **win32com** (`pywin32`) — DRM 보호 Excel 파일 읽기/쓰기 (Microsoft Excel 설치 필요)
- **BeautifulSoup4** — 클립보드 HTML 파싱 (헤더 행 감지)

## 폴더 구조

```
/
├── main.py              # 진입점
├── constants.py         # 색상, EXCEL_MAP, 버전
├── excel_loader.py      # win32com DRM Excel 로더
├── widgets/
│   ├── common.py        # PasteableTable 등 공용 위젯
│   ├── sidebar.py       # SidebarButton, Sidebar
│   ├── menu_bar.py      # MenuBar (카테고리 콤보박스)
│   ├── sr_page.py       # SR 페이지
│   └── history_page.py  # History 페이지 (stub)
├── windows/
│   └── main_window.py   # QMainWindow (1000×700)
└── templates/           # Excel 템플릿 파일 (.gitignore 적용)
```

## 환경 설정

```cmd
# CMD 사용 (PowerShell 은 UnauthorizedAccess 문제)
python -m venv venv
venv\Scripts\activate.bat
pip install -r requirements.txt
python main.py
```

## 주요 설계 결정

- `win32com` import 실패 시 앱은 기동되지만 Excel 로드 시 에러 메시지 표시
- `ExcelLoader` 인스턴스는 `MainWindow`에서 생성·소유, `closeEvent`에서 `quit()` 호출
- `PasteableTable`은 Ctrl+V 붙여넣기 지원, 행 자동 확장
- `EXCEL_MAP`의 템플릿 경로는 `constants.py`에서 관리

## TODO (미구현)

- 클립보드 HTML 파싱으로 헤더 행 자동 감지 (BeautifulSoup)
- 다중 행 헤더 병합
- 플레이스홀더 토큰 시스템 (`{{SR_NO}}`, `{{TITLE}}`, `{{DATE}}`)
- 메타데이터 추출 (정규식)
- 멀티 시트 자동 매칭
- 내보내기 (Excel/PDF, win32com)
- History 페이지
