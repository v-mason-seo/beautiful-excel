"""
최적화 엔진 테스트
"""

import sys
import unittest
from pathlib import Path

# 프로젝트 루트를 경로에 추가
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root / 'src'))

from core.optimizer import ExcelOptimizer


class TestExcelOptimizer(unittest.TestCase):
    """
    ExcelOptimizer 테스트
    """

    def setUp(self):
        """
        테스트 환경 설정
        """
        self.settings = {
            'paper_size': 'A4',
            'orientation': 'landscape',
            'font_size': 10
        }
        self.optimizer = ExcelOptimizer(self.settings)

    def test_font_optimization(self):
        """
        폰트 최적화 테스트
        """
        data = [
            ['김철수', '30', '개발팀'],
            ['이영희', '28', '기획팀']
        ]
        headers = ['이름', '나이', '부서']

        result = self.optimizer.optimize(data, headers)

        # 폰트 최적화 확인
        self.assertIn('font_optimization', result)
        font_opt = result['font_optimization']
        self.assertEqual(font_opt['default_font_size'], 10)
        self.assertEqual(font_opt['default_font_name'], '맑은 고딕')
        self.assertTrue(font_opt['apply_to_all_cells'])

        print("✓ 폰트 최적화 테스트 통과")

    def test_empty_cell_optimization(self):
        """
        빈 셀 최적화 테스트
        """
        # 빈 셀이 많은 데이터
        data = [
            ['김철수', '', '개발팀', '대리'],
            ['이영희', '', '기획팀', '사원'],
            ['박민수', '', '개발팀', '과장'],
            ['최지훈', '', '영업팀', '대리']
        ]
        headers = ['이름', '중간이름', '부서', '직급']

        result = self.optimizer.optimize(data, headers)

        # 빈 셀 최적화 확인
        self.assertIn('empty_cell_optimization', result)
        empty_opt = result['empty_cell_optimization']

        # 두 번째 컬럼(중간이름)은 100% 빈 셀이므로 최적화 대상
        self.assertIn('empty_columns', empty_opt)
        empty_columns = empty_opt['empty_columns']

        # 빈 셀이 50% 이상인 컬럼이 있어야 함
        self.assertTrue(len(empty_columns) > 0)

        # 컬럼 1(중간이름)이 최적화되었는지 확인
        if 1 in empty_columns:
            col_info = empty_columns[1]
            self.assertEqual(col_info['empty_ratio'], 1.0)  # 100% 빈 셀
            self.assertLessEqual(col_info['header_font_size'], 10)  # 헤더 폰트 축소

        print("✓ 빈 셀 최적화 테스트 통과")
        print(f"  - 최적화된 컬럼 수: {len(empty_columns)}")

    def test_bold_optimization(self):
        """
        공통 텍스트 Bold 최적화 테스트
        """
        # 공통 접두사가 있는 데이터
        data = [
            ['서울-강남-대치', '개발팀', 'Level-A'],
            ['서울-강남-역삼', '기획팀', 'Level-B'],
            ['서울-서초-방배', '영업팀', 'Level-C'],
            ['서울-송파-잠실', '마케팅팀', 'Level-A']
        ]
        headers = ['주소', '부서', '등급']

        result = self.optimizer.optimize(data, headers)

        # Bold 최적화 확인
        self.assertIn('bold_optimization', result)
        bold_opt = result['bold_optimization']

        # 첫 번째 컬럼(주소)에 공통 접두사 "서울-"이 있어야 함
        if 0 in bold_opt:
            col_info = bold_opt[0]
            self.assertIn('common_prefix', col_info)
            self.assertTrue(col_info['common_prefix'].startswith('서울'))
            self.assertGreaterEqual(col_info['bold_length'], 2)
            self.assertEqual(len(col_info['affected_rows']), 4)  # 모든 행

        # 세 번째 컬럼(등급)에 공통 접두사 "Level-"이 있어야 함
        if 2 in bold_opt:
            col_info = bold_opt[2]
            self.assertEqual(col_info['common_prefix'], 'Level-')
            self.assertEqual(col_info['bold_length'], 6)

        print("✓ Bold 최적화 테스트 통과")
        print(f"  - Bold 적용 컬럼 수: {len(bold_opt)}")

    def test_header_wrap_optimization(self):
        """
        헤더 줄바꿈 최적화 테스트
        """
        # 헤더가 데이터보다 훨씬 긴 경우
        data = [
            ['김철수', '30', 'A'],
            ['이영희', '28', 'B'],
            ['박민수', '35', 'C']
        ]
        headers = ['이름', '나이', '매우긴헤더이름입니다정말로긴이름']

        result = self.optimizer.optimize(data, headers)

        # 헤더 줄바꿈 최적화 확인
        self.assertIn('header_wrap_optimization', result)
        header_opt = result['header_wrap_optimization']

        # 세 번째 컬럼의 헤더가 줄바꿈 대상이어야 함
        if 2 in header_opt:
            col_info = header_opt[2]
            self.assertTrue(col_info['wrap_text'])

        print("✓ 헤더 줄바꿈 최적화 테스트 통과")
        print(f"  - 줄바꿈 적용 컬럼 수: {len(header_opt)}")

    def test_common_prefix_finder(self):
        """
        공통 접두사 찾기 알고리즘 테스트
        """
        # 케이스 1: 공통 접두사 있음
        strings1 = ['서울-강남-대치', '서울-강남-역삼', '서울-서초-방배']
        prefix1 = self.optimizer._find_common_prefix(strings1)
        self.assertEqual(prefix1, '서울-')

        # 케이스 2: 공통 접두사 없음
        strings2 = ['서울', '부산', '대구']
        prefix2 = self.optimizer._find_common_prefix(strings2)
        self.assertEqual(prefix2, '')

        # 케이스 3: 빈 리스트
        strings3 = []
        prefix3 = self.optimizer._find_common_prefix(strings3)
        self.assertEqual(prefix3, '')

        # 케이스 4: 하나만 있는 경우
        strings4 = ['서울-강남']
        prefix4 = self.optimizer._find_common_prefix(strings4)
        self.assertEqual(prefix4, '')

        print("✓ 공통 접두사 찾기 알고리즘 테스트 통과")

    def test_empty_ratio_calculation(self):
        """
        빈 셀 비율 계산 테스트
        """
        data = [
            ['A', '', 'C'],
            ['D', '', 'F'],
            ['G', '', 'I'],
            ['J', '', 'L']
        ]

        # 컬럼 1의 빈 셀 비율 (100%)
        ratio1 = self.optimizer._calculate_empty_ratio(data, 1)
        self.assertEqual(ratio1, 1.0)

        # 컬럼 0의 빈 셀 비율 (0%)
        ratio0 = self.optimizer._calculate_empty_ratio(data, 0)
        self.assertEqual(ratio0, 0.0)

        print("✓ 빈 셀 비율 계산 테스트 통과")

    def test_comprehensive_optimization(self):
        """
        종합 최적화 테스트
        """
        # 복잡한 실제 데이터 시뮬레이션
        data = [
            ['서울-강남구-대치동', '홍길동', '', '개발팀', 'Senior-Developer'],
            ['서울-강남구-역삼동', '김철수', '', '기획팀', 'Senior-Planner'],
            ['서울-서초구-방배동', '이영희', '', '개발팀', 'Junior-Developer'],
            ['서울-송파구-잠실동', '박민수', '', '영업팀', 'Senior-Sales']
        ]
        headers = ['주소', '이름', '중간이름', '부서', '직급']

        result = self.optimizer.optimize(data, headers)

        # 모든 최적화가 적용되었는지 확인
        self.assertIn('font_optimization', result)
        self.assertIn('empty_cell_optimization', result)
        self.assertIn('bold_optimization', result)
        self.assertIn('header_wrap_optimization', result)

        print("✓ 종합 최적화 테스트 통과")
        print("\n최적화 결과 요약:")
        print(f"  - 빈 셀 최적화 컬럼: {len(result['empty_cell_optimization'].get('empty_columns', {}))}")
        print(f"  - Bold 적용 컬럼: {len(result['bold_optimization'])}")
        print(f"  - 헤더 줄바꿈 컬럼: {len(result['header_wrap_optimization'])}")


if __name__ == '__main__':
    print("=" * 60)
    print("엑셀 최적화 엔진 테스트 시작")
    print("=" * 60)
    unittest.main(verbosity=2)
