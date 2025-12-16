# Rich Text 부분 스타일링 기능 상세 작업 계획서

## 1. 개요

### 1.1 기능 목적
- GridWidget 셀 내 문자열에서 **컬럼별로 다른 부분**만 빨간색으로 강조 표시
- 폐쇄망 환경에서 프린트 출력물 확인 시 데이터 차이를 쉽게 파악

### 1.2 예시
| col1 | col2 | col3 |
|------|------|------|
| vs-dev-**01** | sbnet-**dev**-01 | lb-my-dev-**01**-south1 |
| vs-dev-**02** | sbnet-**prd**-01 | lb-my-dev-**02**-south1 |
| vs-dev-**03** | sbnet-**stg**-01 | lb-my-dev-**03**-south1 |

→ 빨간색(Bold) 부분이 컬럼별로 **다른 부분**

---

## 2. 아키텍처 설계

### 2.1 전체 구조
```
┌─────────────────────────────────────────────────────────────┐
│  RichTextCell (데이터 모델)                                  │
│  ├── text: str                    # 원본 텍스트             │
│  └── styles: List[TextStyle]      # 스타일 정보 리스트       │
│       ├── start: int              # 시작 인덱스             │
│       ├── end: int                # 종료 인덱스             │
│       ├── color: str              # 색상 (예: "#FF0000")    │
│       └── bold: bool              # Bold 여부               │
└─────────────────────────────────────────────────────────────┘
           │                           │
           ▼                           ▼
    ┌──────────────┐           ┌──────────────┐
    │ GridWidget   │           │ ExcelExporter│
    │ + Delegate   │           │ + RichText   │
    │ (HTML 렌더링)│           │ (CellRichText│
    └──────────────┘           └──────────────┘
```

### 2.2 핵심 컴포넌트

| 컴포넌트 | 파일 경로 | 역할 |
|----------|-----------|------|
| `RichTextCell` | `src/core/rich_text.py` | 텍스트 + 스타일 데이터 모델 |
| `TextDiffAnalyzer` | `src/core/diff_analyzer.py` | 컬럼별 차이점 감지 |
| `RichTextDelegate` | `src/ui/rich_text_delegate.py` | Qt HTML 렌더링 delegate |
| `ExcelExporter` (확장) | `src/core/exporter.py` | CellRichText 지원 추가 |

---

## 3. 단계별 구현 계획

### Phase 1: 데이터 모델 구현

#### 3.1.1 RichTextCell 클래스
**파일**: `src/core/rich_text.py`

```python
from dataclasses import dataclass, field
from typing import List, Optional

@dataclass
class TextStyle:
    """텍스트 스타일 정보"""
    start: int          # 시작 인덱스 (inclusive)
    end: int            # 종료 인덱스 (exclusive)
    color: Optional[str] = None      # HEX 색상 (예: "#FF0000")
    bold: bool = False

@dataclass
class RichTextCell:
    """Rich Text 셀 데이터"""
    text: str                               # 원본 텍스트
    styles: List[TextStyle] = field(default_factory=list)

    def to_html(self) -> str:
        """HTML 문자열 변환 (Qt 렌더링용)"""
        pass

    def to_openpyxl_rich_text(self):
        """openpyxl CellRichText 변환"""
        pass
```

#### 3.1.2 예상 작업 시간
- 구현: 2시간
- 테스트: 1시간

---

### Phase 2: 차이점 감지 알고리즘

#### 3.2.1 TextDiffAnalyzer 클래스
**파일**: `src/core/diff_analyzer.py`

```python
class TextDiffAnalyzer:
    """컬럼별 텍스트 차이점 분석"""

    @staticmethod
    def find_common_prefix(strings: List[str]) -> str:
        """공통 접두사 찾기"""
        pass

    @staticmethod
    def find_common_suffix(strings: List[str]) -> str:
        """공통 접미사 찾기"""
        pass

    @staticmethod
    def analyze_column(values: List[str]) -> List[RichTextCell]:
        """
        컬럼 값들을 분석하여 차이점 하이라이트

        예시:
        입력: ["vs-dev-01", "vs-dev-02", "vs-dev-03"]
        출력: [
            RichTextCell("vs-dev-01", [TextStyle(7, 9, "#FF0000", True)]),
            RichTextCell("vs-dev-02", [TextStyle(7, 9, "#FF0000", True)]),
            RichTextCell("vs-dev-03", [TextStyle(7, 9, "#FF0000", True)])
        ]
        """
        pass
```

#### 3.2.2 알고리즘 상세

**접두사/접미사 비교 방식** (권장)
```
입력: ["sbnet-dev-01", "sbnet-prd-01", "sbnet-stg-01"]

1. 공통 접두사: "sbnet-"
2. 공통 접미사: "-01"
3. 다른 부분: "dev", "prd", "stg" (인덱스 6~9)

결과: 인덱스 6~9를 빨간색으로 표시
```

#### 3.2.3 예상 작업 시간
- 구현: 3시간
- 테스트: 2시간

---

### Phase 3: Qt Rich Text Delegate

#### 3.3.1 RichTextDelegate 클래스
**파일**: `src/ui/rich_text_delegate.py`

```python
from PySide6.QtWidgets import QStyledItemDelegate
from PySide6.QtGui import QTextDocument, QPainter
from PySide6.QtCore import Qt, QSize

class RichTextDelegate(QStyledItemDelegate):
    """HTML Rich Text 렌더링 Delegate"""

    def paint(self, painter: QPainter, option, index):
        """HTML을 사용하여 셀 렌더링"""
        # QTextDocument로 HTML 렌더링
        pass

    def sizeHint(self, option, index) -> QSize:
        """셀 크기 계산"""
        pass
```

#### 3.3.2 GridWidget 통합
```python
# GridWidget에 delegate 적용
delegate = RichTextDelegate(self)
self.setItemDelegate(delegate)
```

#### 3.3.3 예상 작업 시간
- 구현: 4시간
- 테스트: 2시간

---

### Phase 4: Excel Export 확장

#### 3.4.1 ExcelExporter 수정
**파일**: `src/core/exporter.py`

```python
from openpyxl.cell.rich_text import TextBlock, CellRichText
from openpyxl.cell.text import InlineFont

def _apply_rich_text(self, ws, rich_text_data: Dict):
    """Rich Text 데이터를 엑셀 셀에 적용"""
    for (row, col), rich_cell in rich_text_data.items():
        cell = ws.cell(row=row+1, column=col+1)

        # CellRichText 생성
        rich = CellRichText()
        current_pos = 0

        for style in rich_cell.styles:
            # 스타일 적용 전 일반 텍스트
            if current_pos < style.start:
                rich.append(rich_cell.text[current_pos:style.start])

            # 스타일이 적용된 텍스트
            font = InlineFont(
                b=style.bold,
                color=style.color.replace("#", "") if style.color else None
            )
            rich.append(TextBlock(font, rich_cell.text[style.start:style.end]))
            current_pos = style.end

        # 나머지 텍스트
        if current_pos < len(rich_cell.text):
            rich.append(rich_cell.text[current_pos:])

        cell.value = rich
```

#### 3.4.2 예상 작업 시간
- 구현: 3시간
- 테스트: 2시간

---

### Phase 5: UI 통합 및 사용자 인터페이스

#### 3.5.1 메뉴/버튼 추가
- 메뉴: `최적화` → `차이점 하이라이트` (단축키: F6)
- 동작: 선택된 컬럼 또는 전체 컬럼에 대해 차이점 분석 실행

#### 3.5.2 MainWindow 통합
```python
def highlight_differences(self):
    """컬럼별 차이점 하이라이트"""
    data = self.grid_widget.get_data()
    analyzer = TextDiffAnalyzer()

    for col_idx in range(len(data[0])):
        col_values = [row[col_idx] for row in data[1:]]  # 헤더 제외
        rich_cells = analyzer.analyze_column(col_values)

        # GridWidget에 적용
        for row_idx, rich_cell in enumerate(rich_cells):
            self.grid_widget.set_rich_text(row_idx + 1, col_idx, rich_cell)
```

#### 3.5.3 예상 작업 시간
- 구현: 2시간
- 테스트: 1시간

---

## 4. 파일 구조

```
src/
├── core/
│   ├── rich_text.py          # [신규] RichTextCell, TextStyle
│   ├── diff_analyzer.py      # [신규] TextDiffAnalyzer
│   └── exporter.py           # [수정] CellRichText 지원 추가
├── ui/
│   ├── rich_text_delegate.py # [신규] RichTextDelegate
│   ├── grid_widget.py        # [수정] delegate 적용, rich text 메서드
│   └── main_window.py        # [수정] 메뉴 추가
└── tests/
    ├── test_rich_text.py     # [신규] 데이터 모델 테스트
    ├── test_diff_analyzer.py # [신규] 차이점 분석 테스트
    └── test_export_rich.py   # [신규] Rich Text 엑셀 저장 테스트
```

---

## 5. 의존성

### 5.1 기존 의존성
- PySide6 (이미 설치됨)
- openpyxl (이미 설치됨)

### 5.2 추가 의존성
- 없음 (기존 라이브러리로 구현 가능)

### 5.3 openpyxl 버전 확인
```bash
# CellRichText는 openpyxl 2.6+ 필요
pip show openpyxl
```

---

## 6. 테스트 계획

### 6.1 단위 테스트

| 테스트 | 검증 내용 |
|--------|-----------|
| `test_rich_text_to_html` | RichTextCell → HTML 변환 정확성 |
| `test_rich_text_to_openpyxl` | RichTextCell → CellRichText 변환 |
| `test_diff_common_prefix` | 공통 접두사 감지 |
| `test_diff_common_suffix` | 공통 접미사 감지 |
| `test_diff_analyze_column` | 전체 컬럼 차이점 분석 |
| `test_export_rich_text` | 엑셀 저장 후 스타일 유지 확인 |

### 6.2 통합 테스트

| 시나리오 | 검증 내용 |
|----------|-----------|
| 데이터 붙여넣기 → 차이점 하이라이트 → 엑셀 저장 | E2E 워크플로우 |
| 1000행 데이터 처리 | 성능 테스트 (목표: < 3초) |
| 한글/특수문자 포함 데이터 | 유니코드 처리 |

---

## 7. 일정 요약

| Phase | 작업 | 예상 시간 |
|-------|------|-----------|
| 1 | 데이터 모델 (RichTextCell) | 3시간 |
| 2 | 차이점 분석 알고리즘 | 5시간 |
| 3 | Qt Rich Text Delegate | 6시간 |
| 4 | Excel Export 확장 | 5시간 |
| 5 | UI 통합 | 3시간 |
| - | 통합 테스트 및 버그 수정 | 3시간 |
| **합계** | | **25시간** |

---

## 8. 위험 요소 및 대응

| 위험 | 영향 | 대응 방안 |
|------|------|-----------|
| openpyxl CellRichText 호환성 | 엑셀 저장 실패 | 버전 확인, 대안 방식 준비 |
| Qt delegate 성능 저하 | 대량 데이터 시 느림 | 캐싱, 가상화 적용 |
| 복잡한 차이점 패턴 | 분석 정확도 저하 | 알고리즘 개선, 수동 설정 옵션 |

---

## 9. 향후 확장 가능성

1. **사용자 정의 색상**: 빨간색 외 다른 색상 선택 가능
2. **스타일 옵션**: Bold, Italic, 밑줄 등 선택
3. **차이점 감지 모드**: 접두사/접미사 외 LCS 알고리즘 추가
4. **수동 스타일링**: 사용자가 직접 텍스트 범위 선택하여 스타일 적용

---

## 10. 승인 및 시작

- [ ] 계획서 검토 완료
- [ ] 구현 시작 승인
- [ ] Phase 1 시작

**작성일**: 2025-12-16
**작성자**: Claude Code
