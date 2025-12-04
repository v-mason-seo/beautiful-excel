"""
레이아웃 최적화 테스트
"""

import sys
import os

# 프로젝트 루트를 sys.path에 추가
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from src.core.optimizer import ExcelOptimizer, PAPER_SIZES, DEFAULT_MARGINS


def test_paper_dimensions():
    """용지 크기 정보 테스트"""
    print("=== 용지 크기 정보 테스트 ===")

    # A4 가로
    settings = {'paper_size': 'A4', 'orientation': 'landscape', 'font_size': 10}
    optimizer = ExcelOptimizer(settings)
    paper_dims = optimizer._get_paper_dimensions()
    print(f"A4 가로: {paper_dims}")
    assert paper_dims['width'] == 297
    assert paper_dims['height'] == 210

    # A4 세로
    settings = {'paper_size': 'A4', 'orientation': 'portrait', 'font_size': 10}
    optimizer = ExcelOptimizer(settings)
    paper_dims = optimizer._get_paper_dimensions()
    print(f"A4 세로: {paper_dims}")
    assert paper_dims['width'] == 210
    assert paper_dims['height'] == 297

    # A3 가로
    settings = {'paper_size': 'A3', 'orientation': 'landscape', 'font_size': 10}
    optimizer = ExcelOptimizer(settings)
    paper_dims = optimizer._get_paper_dimensions()
    print(f"A3 가로: {paper_dims}")
    assert paper_dims['width'] == 420
    assert paper_dims['height'] == 297

    # A3 세로
    settings = {'paper_size': 'A3', 'orientation': 'portrait', 'font_size': 10}
    optimizer = ExcelOptimizer(settings)
    paper_dims = optimizer._get_paper_dimensions()
    print(f"A3 세로: {paper_dims}")
    assert paper_dims['width'] == 297
    assert paper_dims['height'] == 420

    print("✓ 용지 크기 정보 테스트 통과\n")


def test_available_dimensions():
    """가용 영역 계산 테스트"""
    print("=== 가용 영역 계산 테스트 ===")

    settings = {'paper_size': 'A4', 'orientation': 'landscape', 'font_size': 10}
    optimizer = ExcelOptimizer(settings)

    paper_dims = optimizer._get_paper_dimensions()
    available_dims = optimizer._calculate_available_dimensions(paper_dims)

    print(f"용지 크기: {paper_dims}")
    print(f"가용 영역: {available_dims}")

    # 여백 제외 계산 확인
    expected_width = 297 - 10 - 10  # 297 - left - right
    expected_height = 210 - 10 - 10  # 210 - top - bottom

    assert available_dims['width'] == expected_width
    assert available_dims['height'] == expected_height

    print(f"예상 가용 너비: {expected_width}mm, 실제: {available_dims['width']}mm")
    print(f"예상 가용 높이: {expected_height}mm, 실제: {available_dims['height']}mm")
    print("✓ 가용 영역 계산 테스트 통과\n")


def test_column_width_optimization():
    """컬럼 너비 최적화 테스트"""
    print("=== 컬럼 너비 최적화 테스트 ===")

    settings = {'paper_size': 'A4', 'orientation': 'landscape', 'font_size': 10}
    optimizer = ExcelOptimizer(settings)

    # 테스트 데이터
    headers = ['짧은헤더', '중간길이헤더', '매우긴헤더텍스트입니다']
    data = [
        ['데이터1', '중간길이데이터', '짧음'],
        ['데이터2', '중간길이데이터', '짧음'],
        ['데이터3', '중간길이데이터', '짧음']
    ]

    available_width = 277  # A4 가로 가용 너비 (297 - 20)

    column_widths = optimizer._optimize_column_widths_for_paper(
        data, headers, available_width
    )

    print(f"컬럼 너비 최적화 결과:")
    total_width = 0
    for col_idx, width in column_widths.items():
        print(f"  컬럼 {col_idx} ('{headers[col_idx]}'): {width:.2f}mm")
        total_width += width

    print(f"전체 너비: {total_width:.2f}mm (가용 너비: {available_width}mm)")

    # 전체 너비가 가용 너비를 초과하지 않는지 확인
    assert total_width <= available_width + 1  # 약간의 오차 허용

    print("✓ 컬럼 너비 최적화 테스트 통과\n")


def test_row_height_optimization():
    """행 높이 최적화 테스트"""
    print("=== 행 높이 최적화 테스트 ===")

    settings = {'paper_size': 'A4', 'orientation': 'landscape', 'font_size': 10}
    optimizer = ExcelOptimizer(settings)

    headers = ['컬럼1', '컬럼2', '컬럼3']
    data = [
        ['데이터1', '데이터2', '데이터3'],
        ['데이터4', '데이터5', '데이터6']
    ]

    available_height = 190  # A4 가로 가용 높이

    row_heights = optimizer._optimize_row_heights_for_paper(
        data, headers, available_height
    )

    print(f"행 높이 최적화 결과:")
    print(f"  헤더 행: {row_heights[-1]:.2f}mm")
    for row_idx in range(len(data)):
        print(f"  데이터 행 {row_idx}: {row_heights[row_idx]:.2f}mm")

    # 헤더 행이 데이터 행보다 높은지 확인
    assert row_heights[-1] > row_heights[0]

    print("✓ 행 높이 최적화 테스트 통과\n")


def test_page_breaks_calculation():
    """페이지 분할 계산 테스트"""
    print("=== 페이지 분할 계산 테스트 ===")

    settings = {'paper_size': 'A4', 'orientation': 'landscape', 'font_size': 10}
    optimizer = ExcelOptimizer(settings)

    headers = ['컬럼1', '컬럼2', '컬럼3']

    # 50행 데이터 생성
    data = [[f'데이터{i}-1', f'데이터{i}-2', f'데이터{i}-3'] for i in range(50)]

    available_height = 190  # A4 가로 가용 높이

    row_heights = optimizer._optimize_row_heights_for_paper(
        data, headers, available_height
    )

    page_breaks = optimizer._calculate_page_breaks(
        data, headers, row_heights, available_height
    )

    print(f"페이지 분할 계산 결과:")
    print(f"  총 페이지 수: {page_breaks['total_pages']}")
    print(f"  페이지당 행 수: {page_breaks['rows_per_page']}")
    print(f"  분할 지점: {page_breaks['break_points']}")

    # 모든 행이 포함되었는지 확인
    total_rows_in_pages = sum(page_breaks['rows_per_page'])
    assert total_rows_in_pages == len(data)

    print("✓ 페이지 분할 계산 테스트 통과\n")


def test_full_layout_optimization():
    """전체 레이아웃 최적화 통합 테스트"""
    print("=== 전체 레이아웃 최적화 통합 테스트 ===")

    settings = {'paper_size': 'A4', 'orientation': 'landscape', 'font_size': 10}
    optimizer = ExcelOptimizer(settings)

    headers = ['이름', '부서', '직급', '이메일']
    data = [
        ['홍길동', '개발팀', '팀장', 'hong@example.com'],
        ['김철수', '개발팀', '대리', 'kim@example.com'],
        ['이영희', '기획팀', '과장', 'lee@example.com'],
        ['박민수', '기획팀', '사원', 'park@example.com'],
        ['정수진', '마케팅팀', '차장', 'jung@example.com']
    ]

    # 전체 최적화 실행
    optimization_result = optimizer.optimize(data, headers)

    print("최적화 결과:")
    print(f"  폰트 최적화: {optimization_result.get('font_optimization', {}).get('default_font_size')}pt")
    print(f"  빈 셀 최적화: {len(optimization_result.get('empty_cell_optimization', {}).get('empty_columns', {}))}개 컬럼")
    print(f"  Bold 최적화: {len(optimization_result.get('bold_optimization', {}))}개 컬럼")
    print(f"  헤더 줄바꿈: {len(optimization_result.get('header_wrap_optimization', {}))}개 컬럼")

    # 레이아웃 최적화 결과
    layout_opt = optimization_result.get('layout_optimization', {})
    print(f"\n레이아웃 최적화:")
    print(f"  용지 크기: {layout_opt.get('paper_dimensions')}")
    print(f"  가용 영역: {layout_opt.get('available_dimensions')}")
    print(f"  컬럼 너비: {len(layout_opt.get('column_widths', {}))}개")
    print(f"  행 높이: {len(layout_opt.get('row_heights', {}))}개")

    page_breaks = layout_opt.get('page_breaks', {})
    print(f"  페이지 분할:")
    print(f"    총 페이지: {page_breaks.get('total_pages')}")
    print(f"    페이지당 행 수: {page_breaks.get('rows_per_page')}")

    # 레이아웃 최적화가 포함되었는지 확인
    assert 'layout_optimization' in optimization_result
    assert layout_opt.get('paper_dimensions') is not None
    assert layout_opt.get('column_widths') is not None
    assert layout_opt.get('page_breaks') is not None

    print("\n✓ 전체 레이아웃 최적화 통합 테스트 통과\n")


def run_all_tests():
    """모든 테스트 실행"""
    print("\n" + "=" * 60)
    print("레이아웃 최적화 테스트 시작")
    print("=" * 60 + "\n")

    try:
        test_paper_dimensions()
        test_available_dimensions()
        test_column_width_optimization()
        test_row_height_optimization()
        test_page_breaks_calculation()
        test_full_layout_optimization()

        print("=" * 60)
        print("✅ 모든 테스트 통과!")
        print("=" * 60)
        return True

    except AssertionError as e:
        print(f"\n❌ 테스트 실패: {e}")
        import traceback
        traceback.print_exc()
        return False

    except Exception as e:
        print(f"\n❌ 예외 발생: {e}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == '__main__':
    success = run_all_tests()
    sys.exit(0 if success else 1)
