# 엑셀 출력 최적화 프로그램 개발 계획표

## 프로젝트 개요

**프로젝트명**: Beautiful Excel - 엑셀 출력 최적화 프로그램
**목적**: 엑셀 데이터를 A4/A3 용지 규격에 맞춰 자동으로 정리 및 최적화하여 출력
**개발 환경**: Python, PySide6, Windows OS
**예상 개발 기간**: 단계별 순차 진행

---

## 개발 단계별 계획

### Phase 1: 프로젝트 초기 설정 (Foundation)

#### 1.1 프로젝트 구조 설정
- [ ] Python 가상환경 구성
- [ ] 필수 라이브러리 설치 (PySide6, openpyxl, pandas 등)
- [ ] 프로젝트 디렉토리 구조 설계
  ```
  beautiful-excel/
  ├── src/
  │   ├── main.py              # 진입점
  │   ├── ui/                  # GUI 관련
  │   │   ├── main_window.py
  │   │   ├── menu_bar.py
  │   │   └── grid_widget.py
  │   ├── core/                # 핵심 로직
  │   │   ├── excel_loader.py
  │   │   ├── optimizer.py
  │   │   └── exporter.py
  │   └── utils/               # 유틸리티
  ├── tests/                   # 테스트 코드
  ├── resources/               # 리소스 파일
  └── requirements.txt
  ```
- [ ] requirements.txt 작성

**예상 소요 시간**: 1일

---

### Phase 2: GUI 기본 구조 구현 (UI Framework) ✅

#### 2.1 메인 윈도우 구현
- [x] PySide6 기반 메인 윈도우 클래스 생성
- [x] 윈도우 크기, 타이틀 설정
- [x] 기본 레이아웃 구성 (중앙: 그리드, 상단: 메뉴바)

#### 2.2 상단 메뉴바 구현
- [x] 메뉴바 UI 구성
  - 파일 메뉴 (열기, 저장, 종료)
  - 편집 메뉴 (붙여넣기, 모두 지우기)
  - 설정 메뉴 (용지 설정, 최적화 적용)
  - 도움말 메뉴 (프로그램 정보)
- [x] 설정 패널 구현
  - **용지 선택**: A4 / A3 (기본값: A4)
  - **방향 선택**: 가로 / 세로 (기본값: 가로)
  - **글자 크기**: 숫자 입력 필드 (기본값: 10pt, 범위: 8-14pt)

#### 2.3 그리드/테이블 위젯 구현
- [x] QTableWidget 선택 및 구현
- [x] 헤더 설정 기능
- [x] 셀 편집 기능
- [x] 스타일 적용 기능 (폰트, Bold, 색상 등)
- [x] 클립보드 붙여넣기 기능 (Ctrl+V)
- [x] 데이터 입출력 메서드 구현

**완료일**: 2025-12-03

---

### Phase 3: 데이터 입출력 기능 구현 (I/O Operations) ✅

#### 3.1 엑셀 파일 불러오기
- [x] openpyxl을 이용한 XLSX 파일 읽기
- [x] xlrd를 이용한 XLS 파일 읽기 (레거시 지원)
- [x] 파일 대화상자 구현
- [x] 엑셀 데이터를 그리드에 표시
- [x] 서식 정보 유지 (폰트, 색상, Bold 등)

#### 3.2 클립보드 붙여넣기
- [x] QClipboard를 이용한 클립보드 데이터 읽기
- [x] 탭/줄바꿈 기반 데이터 파싱
- [x] 그리드에 데이터 삽입
- [x] Ctrl+V 단축키 지원

#### 3.3 엑셀 파일 저장
- [x] openpyxl을 이용한 XLSX 파일 쓰기
- [x] 최적화된 서식 적용 후 저장
- [x] 파일 저장 대화상자 구현

**완료일**: 2025-12-03

---

### Phase 4: 핵심 최적화 로직 구현 (Optimization Engine) ✅

#### 4.1 폰트/크기 일괄 변환
- [x] 설정 메뉴의 글자 크기 값 가져오기
- [x] 모든 데이터 셀에 폰트 크기 일괄 적용
- [x] 기본 글꼴 설정 (맑은 고딕)

#### 4.2 빈 컬럼 최적화
- [x] 헤더를 제외한 데이터가 모두 비어있는 컬럼 식별
- [x] 해당 컬럼의 헤더를 여러 줄로 표시 (3글자마다 줄바꿈)
- [x] 헤더 행 높이 자동 조정 (줄 수에 맞게)
- [x] 컬럼 너비 최소화 (40px)
- [x] 헤더 제외 셀 배경색 제거 (채우기 없음)

#### 4.3 컬럼별 공통 텍스트 Bold 처리
- [x] 각 컬럼별 데이터 수집
- [x] 앞에서부터 공통 문자열 탐지 알고리즘 구현
  ```python
  # 예시: ["서울-강남-대치", "서울-강남-역삼", "서울-서초-방배"]
  # → "서울-" 부분을 Bold 처리
  ```
- [x] 공통 영역에 Bold 서식 적용 (2글자 이상)
- [x] Qt QFont를 이용한 Bold 처리

#### 4.4 헤더 자동 줄바꿈
- [x] 각 컬럼의 데이터 최대 길이 계산
- [x] 헤더 길이와 비교
- [x] 헤더가 더 긴 경우 자동 줄바꿈 적용 (1.5배 이상, 10자 이상)
- [x] 줄바꿈된 헤더의 행 높이 자동 조정

**완료일**: 2025-12-04

---

### Phase 5: 용지 규격 맞춤 최적화 (Layout Optimization) ✅

#### 5.1 용지 크기 계산
- [x] A4 규격: 210mm × 297mm (가로) / 297mm × 210mm (세로)
- [x] A3 규격: 297mm × 420mm (가로) / 420mm × 297mm (세로)
- [x] 여백 설정 (상하좌우 기본 10mm)

#### 5.2 컬럼 너비 자동 조정
- [x] 용지 가용 너비 계산
- [x] 전체 컬럼 수 기반 초기 너비 배분
- [x] 데이터 길이 기반 가중치 계산
- [x] 최적 너비 재조정 알고리즘 구현

#### 5.3 행 높이 자동 조정
- [x] 글자 크기 기반 최소 행 높이 계산
- [x] 줄바꿈된 셀의 행 높이 자동 확장
- [x] 용지 높이 내 최대 행 수 계산

#### 5.4 페이지 분할 처리
- [x] 한 페이지에 들어갈 수 없는 경우 자동 분할
- [x] 헤더 반복 출력 설정

**완료일**: 2025-12-04

---

### Phase 6: 출력 및 미리보기 (Print & Preview) ✅

#### 6.1 인쇄 미리보기
- [x] QPrintPreviewDialog 구현
- [x] 최적화된 레이아웃 미리보기
- [x] 페이지별 분할 확인

#### 6.2 인쇄 기능
- [x] QPrinter를 이용한 인쇄 구현
- [x] 용지 설정 반영 (A4/A3, 가로/세로)
- [x] 인쇄 대화상자 구현

**완료일**: 2025-12-04

---

### Phase 7: 테스트 및 안정화 (Testing & Stabilization)

#### 7.1 단위 테스트
- [ ] 각 모듈별 단위 테스트 작성
- [ ] 최적화 로직 검증

#### 7.2 통합 테스트
- [ ] 다양한 엑셀 파일 형식 테스트
- [ ] 다양한 데이터 패턴 테스트
  - 빈 셀이 많은 경우
  - 긴 헤더가 있는 경우
  - 공통 텍스트가 있는 경우
- [ ] 용지 크기별 레이아웃 테스트

#### 7.3 버그 수정 및 최적화
- [ ] 발견된 버그 수정
- [ ] 성능 최적화
- [ ] 사용자 경험 개선

**예상 소요 시간**: 3-4일

---

### Phase 8: 배포 준비 (Deployment)

#### 8.1 Windows 실행 파일 생성
- [ ] PyInstaller 설정
- [ ] .exe 파일 빌드
- [ ] 의존성 패키징

#### 8.2 문서화
- [ ] 사용자 매뉴얼 작성
- [ ] 설치 가이드 작성
- [ ] README 작성

#### 8.3 최종 테스트
- [ ] Windows 환경에서 .exe 실행 테스트
- [ ] 폐쇄망 환경 시뮬레이션 테스트

**예상 소요 시간**: 2일

---

## 기술 스택 상세

### 핵심 라이브러리

| 라이브러리 | 용도 | 버전 |
|-----------|------|------|
| **PySide6** | GUI 프레임워크 | 6.6+ |
| **openpyxl** | XLSX 파일 읽기/쓰기 | 3.1+ |
| **xlrd** | XLS 파일 읽기 (레거시) | 2.0+ |
| **pandas** | 데이터 처리 | 2.1+ |
| **PyInstaller** | 실행 파일 빌드 | 6.0+ |

### 개발 도구
- Python 3.10+
- Visual Studio Code 또는 PyCharm
- Git (버전 관리)

---

## 핵심 알고리즘 설계

### 1. 공통 텍스트 탐지 알고리즘
```python
def find_common_prefix(strings):
    """
    컬럼 내 모든 문자열의 공통 접두사 찾기
    """
    if not strings:
        return ""

    strings = [s for s in strings if s]  # 빈 문자열 제외
    if not strings:
        return ""

    min_length = min(len(s) for s in strings)
    common = ""

    for i in range(min_length):
        char = strings[0][i]
        if all(s[i] == char for s in strings):
            common += char
        else:
            break

    return common
```

### 2. 빈 컬럼 판별 및 헤더 줄바꿈
```python
def is_column_data_empty(data, col_index):
    """
    컬럼의 데이터가 모두 비어있는지 확인 (첫 번째 행 = 헤더 제외)
    """
    if not data or len(data) <= 1:
        return True
    
    for row in data[1:]:
        if col_index < len(row):
            cell_value = row[col_index]
            if cell_value and str(cell_value).strip() != "":
                return False
    return True

def wrap_header_text(header_text, max_chars_per_line=3):
    """
    헤더 텍스트를 여러 줄로 줄바꿈
    """
    if not header_text or len(header_text) <= max_chars_per_line:
        return header_text
    
    lines = []
    for i in range(0, len(header_text), max_chars_per_line):
        lines.append(header_text[i:i + max_chars_per_line])
    
    return '\n'.join(lines)
```

### 3. 컬럼 너비 최적화
```python
def optimize_column_widths(data, paper_width, font_size):
    """
    용지 너비에 맞춰 컬럼 너비 최적화
    """
    num_columns = len(data[0])
    char_width = font_size * 0.6  # 평균 문자 너비 (폰트 크기 기준)

    # 각 컬럼별 최대 문자 수 계산
    max_chars = []
    for col_idx in range(num_columns):
        col_data = [row[col_idx] for row in data]
        max_len = max(len(str(cell)) for cell in col_data)
        max_chars.append(max_len)

    # 비율 기반 너비 배분
    total_chars = sum(max_chars)
    widths = []
    for max_char in max_chars:
        width = (max_char / total_chars) * paper_width
        widths.append(max(width, 20))  # 최소 너비 20mm

    return widths
```

---

## 위험 요소 및 대응 방안

### 위험 요소

1. **다양한 엑셀 서식 호환성**
   - 대응: openpyxl과 xlrd를 함께 사용하여 XLSX/XLS 모두 지원
   - 테스트: 다양한 버전의 엑셀 파일로 테스트

2. **한글 폰트 처리**
   - 대응: Windows 기본 폰트 사용 (맑은 고딕, 굴림 등)
   - 폴백 폰트 설정

3. **대용량 데이터 처리**
   - 대응: pandas를 이용한 효율적인 데이터 처리
   - 필요시 페이징 처리

4. **인쇄 미리보기 정확도**
   - 대응: 실제 인쇄와 미리보기 간 차이 최소화
   - 여러 프린터에서 테스트

---

## 성공 기준

- [ ] XLSX/XLS 파일을 정상적으로 불러올 수 있음
- [ ] 클립보드 데이터를 정상적으로 붙여넣을 수 있음
- [ ] 설정된 용지 크기(A4/A3)와 방향에 맞춰 레이아웃이 최적화됨
- [ ] 글자 크기가 설정값대로 일괄 변경됨
- [ ] 빈 컬럼(데이터 없는 컬럼)의 헤더가 여러 줄로 표시되고 너비가 최소화됨
- [ ] 컬럼별 공통 텍스트가 Bold 처리됨
- [ ] 긴 헤더가 자동으로 줄바꿈됨
- [ ] 최적화된 데이터를 엑셀 파일로 저장할 수 있음
- [ ] 인쇄 미리보기가 정상 작동함
- [ ] Windows에서 .exe 파일로 실행 가능함

---

## 다음 단계

1. ✅ 개발 계획표 작성 완료
2. ✅ Phase 1 완료: 프로젝트 구조 설정
3. ✅ Phase 2 완료: GUI 기본 구조 구현
4. ✅ Phase 3 완료: 데이터 입출력 기능 구현
5. ✅ Phase 4 완료: 핵심 최적화 로직 구현
6. ✅ Phase 5 완료: 용지 규격 맞춤 최적화
7. ✅ Phase 6 완료: 출력 및 미리보기
8. ⏳ Phase 7 시작: 테스트 및 안정화

---

## 개발 진행 상황

| Phase | 단계 | 상태 | 완료일 |
|-------|------|------|--------|
| Phase 1 | 프로젝트 초기 설정 | ✅ 완료 | 2025-12-03 |
| Phase 2 | GUI 기본 구조 구현 | ✅ 완료 | 2025-12-03 |
| Phase 3 | 데이터 입출력 기능 | ✅ 완료 | 2025-12-03 |
| Phase 4 | 핵심 최적화 로직 | ✅ 완료 | 2025-12-04 |
| Phase 5 | 용지 규격 맞춤 최적화 | ✅ 완료 | 2025-12-04 |
| Phase 6 | 출력 및 미리보기 | ✅ 완료 | 2025-12-04 |
| Phase 7 | 테스트 및 안정화 | ⏳ 대기 | - |
| Phase 8 | 배포 준비 | ⏳ 대기 | - |

---

**작성일**: 2025-12-03
**최종 업데이트**: 2025-12-05 (빈 컬럼 최적화 개선)
**작성자**: Claude Code
