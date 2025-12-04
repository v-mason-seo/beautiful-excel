"""
엑셀 데이터 최적화 엔진
용지 규격에 맞춰 데이터를 최적화하는 핵심 로직
"""

from typing import List, Dict, Tuple, Optional, Any


# 용지 규격 정의 (단위: mm)
PAPER_SIZES = {
    'A4': {
        'landscape': {'width': 297, 'height': 210},  # 가로
        'portrait': {'width': 210, 'height': 297}     # 세로
    },
    'A3': {
        'landscape': {'width': 420, 'height': 297},  # 가로
        'portrait': {'width': 297, 'height': 420}     # 세로
    }
}

# 기본 여백 (단위: mm)
DEFAULT_MARGINS = {
    'top': 10,
    'bottom': 10,
    'left': 10,
    'right': 10
}


class ExcelOptimizer:
    """
    엑셀 데이터 최적화 클래스
    """

    def __init__(self, settings: Optional[Dict[str, Any]] = None):
        """
        초기화

        Args:
            settings: 최적화 설정 (용지 크기, 방향, 글자 크기 등)
        """
        self.settings = settings or {
            'paper_size': 'A4',
            'orientation': 'landscape',
            'font_size': 10
        }
        self.margins = DEFAULT_MARGINS.copy()

    def optimize(
        self,
        data: List[List[str]],
        headers: List[str]
    ) -> Dict[str, Any]:
        """
        데이터 최적화 수행

        Args:
            data: 2차원 리스트 형태의 데이터 (헤더 제외)
            headers: 헤더 리스트

        Returns:
            dict: {
                'font_optimization': 폰트 최적화 정보,
                'empty_cell_optimization': 빈 셀 최적화 정보,
                'bold_optimization': Bold 처리 정보,
                'header_wrap_optimization': 헤더 줄바꿈 정보,
                'layout_optimization': 레이아웃 최적화 정보
            }
        """
        optimization_result = {
            'font_optimization': self._optimize_font(data, headers),
            'empty_cell_optimization': self._optimize_empty_cells(data, headers),
            'bold_optimization': self._optimize_bold_text(data, headers),
            'header_wrap_optimization': self._optimize_header_wrap(data, headers),
            'layout_optimization': self._optimize_layout(data, headers)
        }

        return optimization_result

    def _optimize_font(
        self,
        data: List[List[str]],
        headers: List[str]
    ) -> Dict[str, Any]:
        """
        폰트/크기 일괄 변환

        Args:
            data: 데이터 리스트
            headers: 헤더 리스트

        Returns:
            dict: 폰트 최적화 정보
        """
        font_size = self.settings.get('font_size', 10)
        font_name = '맑은 고딕'

        return {
            'default_font_size': font_size,
            'default_font_name': font_name,
            'apply_to_all_cells': True
        }

    def _optimize_empty_cells(
        self,
        data: List[List[str]],
        headers: List[str]
    ) -> Dict[str, Any]:
        """
        빈 셀 최적화

        Args:
            data: 데이터 리스트
            headers: 헤더 리스트

        Returns:
            dict: {
                'empty_columns': {컬럼_인덱스: {'empty_ratio': 비율, 'header_font_size': 크기}},
                'column_widths': {컬럼_인덱스: 너비}
            }
        """
        if not data or len(data) == 0:
            return {'empty_columns': {}, 'column_widths': {}}

        num_columns = len(headers)
        empty_columns = {}
        column_widths = {}

        for col_idx in range(num_columns):
            # 빈 셀 비율 계산
            empty_ratio = self._calculate_empty_ratio(data, col_idx)

            # 빈 셀이 50% 이상인 경우 최적화 적용
            if empty_ratio >= 0.5:
                # 헤더 폰트 크기 축소 (기본 크기에서 2pt 감소, 최소 8pt)
                default_font_size = self.settings.get('font_size', 10)
                reduced_font_size = max(default_font_size - 2, 8)

                empty_columns[col_idx] = {
                    'empty_ratio': empty_ratio,
                    'header_font_size': reduced_font_size,
                    'reason': f'빈 셀 {empty_ratio*100:.1f}%'
                }

                # 컬럼 너비 최소화 (픽셀 단위)
                column_widths[col_idx] = 60  # 최소 너비
            else:
                # 데이터 길이 기반 너비 계산
                max_length = self._get_column_max_length(data, col_idx)
                header_length = len(headers[col_idx]) if col_idx < len(headers) else 0
                max_length = max(max_length, header_length)

                # 한글은 더 넓은 공간 필요 (픽셀 단위)
                column_widths[col_idx] = min(max(max_length * 10, 80), 300)

        return {
            'empty_columns': empty_columns,
            'column_widths': column_widths
        }

    def _optimize_bold_text(
        self,
        data: List[List[str]],
        headers: List[str]
    ) -> Dict[str, Any]:
        """
        컬럼별 공통 텍스트 Bold 처리

        Args:
            data: 데이터 리스트
            headers: 헤더 리스트

        Returns:
            dict: {
                컬럼_인덱스: {
                    'common_prefix': 공통 접두사,
                    'bold_length': Bold 처리할 문자 길이,
                    'affected_rows': 영향받는 행 리스트
                }
            }
        """
        if not data or len(data) == 0:
            return {}

        num_columns = len(headers)
        bold_optimization = {}

        for col_idx in range(num_columns):
            # 컬럼 데이터 추출
            column_data = self._extract_column_data(data, col_idx)

            # 공통 접두사 찾기
            common_prefix = self._find_common_prefix(column_data)

            # 공통 접두사가 2글자 이상인 경우만 적용
            if common_prefix and len(common_prefix) >= 2:
                affected_rows = []
                for row_idx, cell_value in enumerate(column_data):
                    if cell_value.startswith(common_prefix):
                        affected_rows.append(row_idx)

                bold_optimization[col_idx] = {
                    'common_prefix': common_prefix,
                    'bold_length': len(common_prefix),
                    'affected_rows': affected_rows,
                    'reason': f'공통 접두사 "{common_prefix}"'
                }

        return bold_optimization

    def _optimize_header_wrap(
        self,
        data: List[List[str]],
        headers: List[str]
    ) -> Dict[str, Any]:
        """
        헤더 자동 줄바꿈

        Args:
            data: 데이터 리스트
            headers: 헤더 리스트

        Returns:
            dict: {
                컬럼_인덱스: {
                    'wrap_text': True/False,
                    'reason': 이유
                }
            }
        """
        if not data or len(data) == 0:
            return {}

        num_columns = len(headers)
        header_wrap = {}

        for col_idx in range(num_columns):
            if col_idx >= len(headers):
                continue

            header_length = len(headers[col_idx])
            max_data_length = self._get_column_max_length(data, col_idx)

            # 헤더가 데이터보다 1.5배 이상 긴 경우 줄바꿈 적용
            if header_length > max_data_length * 1.5 and header_length > 10:
                header_wrap[col_idx] = {
                    'wrap_text': True,
                    'reason': f'헤더({header_length}자) > 데이터({max_data_length}자)'
                }

        return header_wrap

    # === 유틸리티 메서드 ===

    @staticmethod
    def _calculate_empty_ratio(data: List[List[str]], col_index: int) -> float:
        """
        컬럼의 빈 셀 비율 계산

        Args:
            data: 데이터 리스트
            col_index: 컬럼 인덱스

        Returns:
            float: 빈 셀 비율 (0.0 ~ 1.0)
        """
        if not data:
            return 0.0

        total = len(data)
        empty = 0

        for row in data:
            if col_index < len(row):
                cell_value = row[col_index]
                if not cell_value or str(cell_value).strip() == "":
                    empty += 1
            else:
                empty += 1

        return empty / total if total > 0 else 0.0

    @staticmethod
    def _get_column_max_length(data: List[List[str]], col_index: int) -> int:
        """
        컬럼 데이터의 최대 문자 길이

        Args:
            data: 데이터 리스트
            col_index: 컬럼 인덱스

        Returns:
            int: 최대 문자 길이
        """
        max_length = 0

        for row in data:
            if col_index < len(row):
                cell_value = row[col_index]
                if cell_value:
                    max_length = max(max_length, len(str(cell_value)))

        return max_length

    @staticmethod
    def _extract_column_data(data: List[List[str]], col_index: int) -> List[str]:
        """
        특정 컬럼의 데이터 추출

        Args:
            data: 데이터 리스트
            col_index: 컬럼 인덱스

        Returns:
            List[str]: 컬럼 데이터 리스트
        """
        column_data = []

        for row in data:
            if col_index < len(row):
                cell_value = row[col_index]
                if cell_value and str(cell_value).strip():
                    column_data.append(str(cell_value).strip())

        return column_data

    @staticmethod
    def _find_common_prefix(strings: List[str]) -> str:
        """
        문자열 리스트의 공통 접두사 찾기

        Args:
            strings: 문자열 리스트

        Returns:
            str: 공통 접두사
        """
        if not strings or len(strings) == 0:
            return ""

        # 빈 문자열 제외
        strings = [s for s in strings if s]
        if not strings:
            return ""

        # 문자열이 1개만 있으면 공통 접두사 없음
        if len(strings) == 1:
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

    # === 레이아웃 최적화 메서드 ===

    def _optimize_layout(
        self,
        data: List[List[str]],
        headers: List[str]
    ) -> Dict[str, Any]:
        """
        용지 규격에 맞춘 레이아웃 최적화

        Args:
            data: 데이터 리스트
            headers: 헤더 리스트

        Returns:
            dict: {
                'paper_dimensions': 용지 크기 정보,
                'available_dimensions': 가용 영역 크기,
                'column_widths': 최적화된 컬럼 너비 (mm),
                'row_heights': 최적화된 행 높이 (mm),
                'page_breaks': 페이지 분할 정보
            }
        """
        # 용지 크기 정보 가져오기
        paper_dims = self._get_paper_dimensions()

        # 가용 영역 계산
        available_dims = self._calculate_available_dimensions(paper_dims)

        # 컬럼 너비 최적화
        column_widths = self._optimize_column_widths_for_paper(
            data, headers, available_dims['width']
        )

        # 행 높이 최적화
        row_heights = self._optimize_row_heights_for_paper(
            data, headers, available_dims['height']
        )

        # 페이지 분할 계산
        page_breaks = self._calculate_page_breaks(
            data, headers, row_heights, available_dims['height']
        )

        return {
            'paper_dimensions': paper_dims,
            'available_dimensions': available_dims,
            'column_widths': column_widths,
            'row_heights': row_heights,
            'page_breaks': page_breaks
        }

    def _get_paper_dimensions(self) -> Dict[str, float]:
        """
        용지 크기 정보 가져오기

        Returns:
            dict: {'width': 너비(mm), 'height': 높이(mm)}
        """
        paper_size = self.settings.get('paper_size', 'A4')
        orientation = self.settings.get('orientation', 'landscape')

        if paper_size not in PAPER_SIZES:
            paper_size = 'A4'

        if orientation not in ['landscape', 'portrait']:
            orientation = 'landscape'

        return PAPER_SIZES[paper_size][orientation].copy()

    def _calculate_available_dimensions(
        self,
        paper_dims: Dict[str, float]
    ) -> Dict[str, float]:
        """
        여백을 제외한 가용 영역 계산

        Args:
            paper_dims: 용지 크기 정보

        Returns:
            dict: {'width': 가용 너비(mm), 'height': 가용 높이(mm)}
        """
        available_width = (
            paper_dims['width']
            - self.margins['left']
            - self.margins['right']
        )

        available_height = (
            paper_dims['height']
            - self.margins['top']
            - self.margins['bottom']
        )

        return {
            'width': available_width,
            'height': available_height
        }

    def _optimize_column_widths_for_paper(
        self,
        data: List[List[str]],
        headers: List[str],
        available_width: float
    ) -> Dict[int, float]:
        """
        용지 너비에 맞춰 컬럼 너비 최적화

        Args:
            data: 데이터 리스트
            headers: 헤더 리스트
            available_width: 가용 너비 (mm)

        Returns:
            dict: {컬럼_인덱스: 너비(mm)}
        """
        if not data or len(data) == 0:
            return {}

        num_columns = len(headers)
        font_size = self.settings.get('font_size', 10)

        # 문자당 평균 너비 (mm) - 폰트 크기 기반 추정
        # 한글은 영문보다 약간 넓음
        char_width_mm = font_size * 0.5  # 대략적인 추정치

        # 각 컬럼별 최대 문자 수 계산
        max_chars_per_column = []
        for col_idx in range(num_columns):
            # 헤더 길이 고려
            header_length = len(headers[col_idx]) if col_idx < len(headers) else 0
            data_max_length = self._get_column_max_length(data, col_idx)
            max_length = max(header_length, data_max_length)
            max_chars_per_column.append(max_length)

        # 전체 문자 수 합계
        total_chars = sum(max_chars_per_column)

        # 비율 기반 너비 배분
        column_widths = {}
        min_width_mm = 20  # 최소 컬럼 너비 (mm)

        for col_idx, max_chars in enumerate(max_chars_per_column):
            if total_chars > 0:
                # 문자 비율에 따라 너비 배분
                width = (max_chars / total_chars) * available_width
                # 최소 너비 보장
                width = max(width, min_width_mm)
                column_widths[col_idx] = width
            else:
                column_widths[col_idx] = min_width_mm

        # 전체 너비가 가용 너비를 초과하는 경우 비율로 축소
        total_width = sum(column_widths.values())
        if total_width > available_width:
            scale_factor = available_width / total_width
            for col_idx in column_widths:
                column_widths[col_idx] *= scale_factor
                # 최소 너비 재보장
                column_widths[col_idx] = max(column_widths[col_idx], min_width_mm)

        return column_widths

    def _optimize_row_heights_for_paper(
        self,
        data: List[List[str]],
        headers: List[str],
        available_height: float
    ) -> Dict[int, float]:
        """
        용지 높이에 맞춰 행 높이 최적화

        Args:
            data: 데이터 리스트
            headers: 헤더 리스트
            available_height: 가용 높이 (mm)

        Returns:
            dict: {행_인덱스: 높이(mm)}
        """
        font_size = self.settings.get('font_size', 10)

        # 기본 행 높이 (mm) - 폰트 크기 기반
        # 1pt = 0.3527mm, 행간 여유 1.2배
        base_row_height_mm = font_size * 0.3527 * 1.2

        row_heights = {}

        # 헤더 행 (-1 인덱스로 표현)
        row_heights[-1] = base_row_height_mm * 1.5  # 헤더는 조금 더 높게

        # 데이터 행
        for row_idx in range(len(data)):
            # 기본적으로 모든 행은 동일한 높이
            row_heights[row_idx] = base_row_height_mm

            # 줄바꿈이 필요한 셀이 있는지 확인 (추후 확장 가능)
            # 현재는 기본 높이 사용

        return row_heights

    def _calculate_page_breaks(
        self,
        data: List[List[str]],
        headers: List[str],
        row_heights: Dict[int, float],
        available_height: float
    ) -> Dict[str, Any]:
        """
        페이지 분할 계산

        Args:
            data: 데이터 리스트
            headers: 헤더 리스트
            row_heights: 행 높이 정보
            available_height: 가용 높이 (mm)

        Returns:
            dict: {
                'total_pages': 총 페이지 수,
                'rows_per_page': 페이지당 행 수,
                'break_points': [페이지 분할 지점 행 인덱스]
            }
        """
        if not data or len(data) == 0:
            return {
                'total_pages': 1,
                'rows_per_page': [],
                'break_points': []
            }

        # 헤더 높이
        header_height = row_heights.get(-1, 0)

        # 데이터 행 높이 (모두 동일하다고 가정)
        data_row_height = row_heights.get(0, 0)

        # 한 페이지에 들어갈 수 있는 행 수 계산
        available_for_data = available_height - header_height
        rows_per_page_count = int(available_for_data / data_row_height)

        # 최소 1행은 보장
        rows_per_page_count = max(rows_per_page_count, 1)

        # 페이지 분할 지점 계산
        total_rows = len(data)
        break_points = []
        rows_per_page = []

        current_row = 0
        while current_row < total_rows:
            next_break = min(current_row + rows_per_page_count, total_rows)
            rows_in_this_page = next_break - current_row
            rows_per_page.append(rows_in_this_page)

            if next_break < total_rows:
                break_points.append(next_break)

            current_row = next_break

        total_pages = len(rows_per_page)

        return {
            'total_pages': total_pages,
            'rows_per_page': rows_per_page,
            'break_points': break_points
        }
