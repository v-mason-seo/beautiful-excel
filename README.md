# Beautiful Excel 📄

엑셀 데이터를 A4/A3 용지 규격에 맞춰 자동으로 정리 및 최적화하여 출력하는 프로그램

## 프로젝트 개요

**목적**: 폐쇄망 환경에서 엑셀 파일을 용지 규격에 맞춰 자동으로 최적화하여 출력
**개발 환경**: Python 3.10+, PySide6, Windows OS

## 주요 기능

### 데이터 입/출력
- ✅ 엑셀 파일 불러오기 (XLSX/XLS)
- ✅ 클립보드 붙여넣기
- ✅ 최적화된 데이터 엑셀 파일 저장

### 설정 메뉴
- **용지 선택**: A4 (기본값), A3
- **방향 선택**: 가로 (기본값), 세로
- **글자 크기**: 8pt ~ 14pt (기본값: 10pt)

### 자동 최적화 기능
- 📝 폰트/크기 일괄 변환
- 🔍 빈 셀 최적화 (헤더 축소, 너비 조정)
- 💪 컬럼별 공통 텍스트 Bold 처리
- 📐 헤더 자동 줄바꿈
- 📏 용지 규격 맞춤 (컬럼 너비 및 행 높이 자동 조정)

### 인쇄 기능
- ✅ 인쇄 미리보기 (Ctrl+Shift+P)
- ✅ 인쇄 다이얼로그 (Ctrl+P)
- ✅ 용지 설정 반영 (A4/A3, 가로/세로, 여백 10mm)
- ✅ 페이지 분할 자동 처리

## 설치 방법

### 1. Python 가상환경 생성
```bash
python -m venv venv
```

### 2. 가상환경 활성화
**Windows:**
```bash
venv\Scripts\activate
```

**macOS/Linux:**
```bash
source venv/bin/activate
```

### 3. 필수 라이브러리 설치
```bash
pip install -r requirements.txt
```

## 실행 방법

### 개발 모드
```bash
python run.py
```

### 또는 모듈로 실행
```bash
python -m src.main
```

### Windows 실행 파일(.exe) 빌드
```bash
pyinstaller --onefile --windowed --name BeautifulExcel run.py
```

## 프로젝트 구조

```
beautiful-excel/
├── run.py                   # 실행 진입점 (권장)
├── src/
│   ├── main.py              # 메인 애플리케이션
│   ├── ui/                  # GUI 관련 모듈
│   │   ├── __init__.py
│   │   ├── main_window.py   # 메인 윈도우
│   │   ├── settings_panel.py # 설정 패널
│   │   └── grid_widget.py   # 그리드 위젯
│   ├── core/                # 핵심 로직 모듈
│   │   ├── __init__.py
│   │   ├── excel_loader.py  # 엑셀 파일 로드 ✅
│   │   ├── exporter.py      # 엑셀 파일 저장 ✅
│   │   ├── optimizer.py     # 최적화 엔진 ✅
│   │   ├── layout_optimizer.py # 레이아웃 최적화 ✅
│   │   └── print_engine.py  # 인쇄 엔진 ✅
│   └── utils/               # 유틸리티 모듈
│       └── __init__.py
├── tests/                   # 테스트 코드
│   ├── __init__.py
│   ├── test_excel_io.py    # 입출력 테스트 ✅
│   ├── test_optimizer.py   # 최적화 테스트 ✅
│   ├── test_layout_optimizer.py # 레이아웃 테스트 ✅
│   └── test_print_manual.py # 인쇄 테스트 ✅
├── resources/               # 리소스 파일
├── claudedocs/              # 개발 문서
│   └── development_plan.md
├── requirements.txt         # 의존성 패키지 목록
├── .gitignore              # Git 제외 목록
├── CLAUDE.md               # 개발 요구사항
└── README.md               # 프로젝트 문서
```

## 개발 일정

- **Phase 1**: 프로젝트 초기 설정 ✅
- **Phase 2**: GUI 기본 구조 구현 ✅
- **Phase 3**: 데이터 입출력 기능 ✅
- **Phase 4**: 핵심 최적화 로직 ✅
- **Phase 5**: 용지 규격 맞춤 최적화 ✅
- **Phase 6**: 출력 및 미리보기 ✅
- **Phase 7**: 테스트 및 안정화 (진행 예정)
- **Phase 8**: 배포 준비 (진행 예정)

## 기술 스택

| 구분 | 기술 | 버전 |
|------|------|------|
| 언어 | Python | 3.10+ |
| GUI | PySide6 | 6.6+ |
| Excel | openpyxl | 3.1+ |
| Excel (Legacy) | xlrd | 2.0+ |
| 데이터 처리 | pandas | 2.1+ |
| 빌드 | PyInstaller | 6.0+ |

## 라이선스

Internal Use Only

## 작성자

개발: Claude Code
작성일: 2025-12-03
