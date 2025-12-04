"""
PrintEngine 수동 테스트 (GUI 필요)

실행 방법: PYTHONPATH=src python tests/test_print_manual.py
"""

import sys
sys.path.insert(0, 'src')

from PySide6.QtWidgets import QApplication
from core.print_engine import PrintEngine


def test_print_engine_basic():
    """
    PrintEngine 기본 기능 테스트
    """
    print("=== PrintEngine 기본 테스트 ===\n")

    # 1. PrintEngine 객체 생성
    settings = {
        'paper_size': 'A4',
        'orientation': 'landscape',
        'font_size': 10
    }

    engine = PrintEngine(settings)
    print("✓ PrintEngine 객체 생성 성공")
    print(f"  - 용지 크기: {settings['paper_size']}")
    print(f"  - 방향: {settings['orientation']}")
    print(f"  - 폰트 크기: {settings['font_size']}pt\n")

    # 2. 프린터 생성 테스트
    try:
        printer = engine._create_printer()
        print("✓ QPrinter 객체 생성 성공")
        print(f"  - Page Size: {printer.pageSize()}")
        print(f"  - Orientation: {printer.orientation()}\n")
    except Exception as e:
        print(f"✗ QPrinter 생성 실패: {e}\n")

    # 3. 페이지 계산 테스트
    data = [
        ['서울-강남-대치', '홍길동', '010-1234-5678'],
        ['서울-강남-역삼', '김철수', '010-2345-6789'],
        ['서울-서초-방배', '이영희', '010-3456-7890'],
    ]
    headers = ['주소', '이름', '전화번호']

    try:
        pages = engine._calculate_pages(data, headers)
        print("✓ 페이지 계산 성공")
        print(f"  - 데이터 행 수: {len(data)}")
        print(f"  - 예상 페이지 수: {pages}\n")
    except Exception as e:
        print(f"✗ 페이지 계산 실패: {e}\n")

    # 4. 다양한 설정 테스트
    print("=== 다양한 설정 테스트 ===\n")

    test_configs = [
        {'paper_size': 'A4', 'orientation': 'landscape', 'font_size': 10},
        {'paper_size': 'A4', 'orientation': 'portrait', 'font_size': 12},
        {'paper_size': 'A3', 'orientation': 'landscape', 'font_size': 8},
        {'paper_size': 'A3', 'orientation': 'portrait', 'font_size': 14},
    ]

    for i, config in enumerate(test_configs, 1):
        print(f"테스트 {i}: {config['paper_size']} {config['orientation']} {config['font_size']}pt")
        try:
            engine = PrintEngine(config)
            printer = engine._create_printer()
            pages = engine._calculate_pages(data, headers)
            print(f"  ✓ 성공 - 예상 페이지: {pages}\n")
        except Exception as e:
            print(f"  ✗ 실패: {e}\n")

    print("=== 테스트 완료 ===")


if __name__ == '__main__':
    # QApplication 생성 (Qt GUI 컴포넌트 사용을 위해 필요)
    app = QApplication(sys.argv)

    test_print_engine_basic()

    print("\n인쇄 기능의 전체 테스트는 GUI를 통해 수동으로 확인해야 합니다:")
    print("  1. 프로그램 실행: python run.py")
    print("  2. 데이터 로드 또는 붙여넣기")
    print("  3. 메뉴 > 인쇄 > 인쇄 미리보기 (Ctrl+Shift+P)")
    print("  4. 메뉴 > 인쇄 > 인쇄 (Ctrl+P)")
