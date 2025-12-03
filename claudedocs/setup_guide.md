# Beautiful Excel 설치 가이드

## Phase 1 완료 체크리스트

- ✅ 프로젝트 디렉토리 구조 생성
- ✅ requirements.txt 작성
- ✅ .gitignore 작성
- ✅ README.md 작성
- ✅ 기본 Python 패키지 구조 생성
- ✅ src/main.py 진입점 생성

---

## 다음 단계: 개발 환경 설정

### 1. Python 가상환경 생성

프로젝트 루트 디렉토리에서 실행:

```bash
python -m venv venv
```

또는 Python 3.10+ 명시적 사용:

```bash
python3.10 -m venv venv
```

### 2. 가상환경 활성화

**Windows (PowerShell):**
```powershell
.\venv\Scripts\Activate.ps1
```

**Windows (CMD):**
```cmd
venv\Scripts\activate.bat
```

**macOS/Linux:**
```bash
source venv/bin/activate
```

### 3. pip 업그레이드

```bash
python -m pip install --upgrade pip
```

### 4. 필수 라이브러리 설치

```bash
pip install -r requirements.txt
```

### 5. 설치 확인

```bash
python src/main.py
```

예상 출력:
```
Beautiful Excel v0.1.0
프로젝트 초기 설정이 완료되었습니다.

다음 단계: Phase 2 - GUI 기본 구조 구현
```

---

## 설치 예상 라이브러리 목록

### 핵심 라이브러리
- **PySide6** (6.6.0+): Qt 기반 GUI 프레임워크
- **openpyxl** (3.1.0+): XLSX 파일 읽기/쓰기
- **xlrd** (2.0.0+): XLS 파일 읽기 (레거시)
- **xlwt** (1.3.0+): XLS 파일 쓰기 (레거시)
- **pandas** (2.1.0+): 데이터 처리 및 분석
- **numpy** (1.24.0+): 수치 계산

### 빌드 도구
- **pyinstaller** (6.0.0+): Windows .exe 빌드

### 개발 도구 (선택사항)
- **pytest** (7.4.0+): 테스트 프레임워크
- **black** (23.0.0+): 코드 포맷터
- **flake8** (6.1.0+): 린터

---

## 문제 해결

### Windows에서 PySide6 설치 오류

**증상**: `pip install PySide6` 시 빌드 에러

**해결**:
1. Visual C++ 재배포 가능 패키지 설치
2. Microsoft C++ Build Tools 설치
3. 또는 미리 빌드된 wheel 파일 사용:
   ```bash
   pip install PySide6 --only-binary :all:
   ```

### xlrd 버전 경고

**증상**: `.xlsx` 파일을 xlrd로 읽을 때 경고

**해결**:
- XLSX 파일은 `openpyxl` 사용 (이미 구현됨)
- XLS 파일만 `xlrd` 사용

### 가상환경 활성화 오류 (Windows)

**증상**: PowerShell에서 `Activate.ps1` 실행 권한 오류

**해결**:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

---

## 프로젝트 구조 설명

```
beautiful-excel/
├── src/                     # 소스 코드
│   ├── main.py             # 진입점
│   ├── ui/                 # GUI 컴포넌트
│   ├── core/               # 핵심 로직
│   └── utils/              # 유틸리티
├── tests/                  # 테스트 코드
├── resources/              # 리소스 파일
├── claudedocs/             # 개발 문서
│   ├── development_plan.md # 개발 계획표
│   └── setup_guide.md      # 이 파일
├── requirements.txt        # 의존성 목록
├── .gitignore             # Git 제외 목록
├── CLAUDE.md              # 개발 요구사항
└── README.md              # 프로젝트 소개
```

---

## 다음 개발 단계

### Phase 2: GUI 기본 구조 구현 (예정)

1. PySide6 기반 메인 윈도우 생성
2. 상단 메뉴바 구현 (파일, 설정)
3. 설정 패널 구현 (용지, 방향, 글자 크기)
4. QTableWidget 기반 그리드 위젯 구현
5. 기본 레이아웃 구성

**예상 소요 시간**: 2-3일

---

## 참고 자료

- [PySide6 공식 문서](https://doc.qt.io/qtforpython-6/)
- [openpyxl 문서](https://openpyxl.readthedocs.io/)
- [pandas 문서](https://pandas.pydata.org/docs/)
- [PyInstaller 가이드](https://pyinstaller.org/en/stable/)

---

**Phase 1 완료일**: 2025-12-03
**다음 단계**: Phase 2 - GUI 기본 구조 구현
