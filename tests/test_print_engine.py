"""
PrintEngine 테스트
"""

import sys
import pytest
from PySide6.QtWidgets import QApplication
from PySide6.QtPrintSupport import QPrinter

# 경로 추가
sys.path.insert(0, '../src')

from core.print_engine import PrintEngine


@pytest.fixture
def app():
    """QApplication 픽스처 (GUI 테스트용)"""
    app = QApplication.instance()
    if app is None:
        app = QApplication(sys.argv)
    yield app


def test_print_engine_creation():
    """
    PrintEngine 객체 생성 테스트
    """
    settings = {
        'paper_size': 'A4',
        'orientation': 'landscape',
        'font_size': 10
    }

    engine = PrintEngine(settings)
    assert engine.settings == settings


def test_create_printer_a4_landscape():
    """
    A4 가로 용지 프린터 생성 테스트
    """
    settings = {
        'paper_size': 'A4',
        'orientation': 'landscape',
        'font_size': 10
    }

    engine = PrintEngine(settings)
    printer = engine._create_printer()

    assert printer.pageSize() == QPrinter.A4
    assert printer.orientation() == QPrinter.Landscape


def test_create_printer_a3_portrait():
    """
    A3 세로 용지 프린터 생성 테스트
    """
    settings = {
        'paper_size': 'A3',
        'orientation': 'portrait',
        'font_size': 12
    }

    engine = PrintEngine(settings)
    printer = engine._create_printer()

    assert printer.pageSize() == QPrinter.A3
    assert printer.orientation() == QPrinter.Portrait


def test_calculate_pages_single_page():
    """
    단일 페이지 계산 테스트
    """
    settings = {
        'paper_size': 'A4',
        'orientation': 'landscape',
        'font_size': 10
    }

    engine = PrintEngine(settings)

    # 작은 데이터셋
    data = [
        ['Data1', 'Data2', 'Data3'],
        ['Data4', 'Data5', 'Data6'],
    ]
    headers = ['Header1', 'Header2', 'Header3']

    pages = engine._calculate_pages(data, headers)
    assert pages >= 1


def test_calculate_pages_multiple_pages():
    """
    다중 페이지 계산 테스트
    """
    settings = {
        'paper_size': 'A4',
        'orientation': 'portrait',
        'font_size': 10
    }

    engine = PrintEngine(settings)

    # 큰 데이터셋 (100 행)
    data = [[f'Data{i}_{j}' for j in range(5)] for i in range(100)]
    headers = ['Header1', 'Header2', 'Header3', 'Header4', 'Header5']

    pages = engine._calculate_pages(data, headers)
    assert pages > 1


def test_printer_margins():
    """
    프린터 여백 설정 테스트
    """
    settings = {
        'paper_size': 'A4',
        'orientation': 'landscape',
        'font_size': 10
    }

    engine = PrintEngine(settings)
    printer = engine._create_printer()

    # 여백이 설정되어 있는지 확인 (정확한 값 확인은 어려움)
    assert printer.pageRect().width() < printer.paperRect().width()
    assert printer.pageRect().height() < printer.paperRect().height()


def test_different_font_sizes():
    """
    다양한 폰트 크기 설정 테스트
    """
    for font_size in [8, 10, 12, 14]:
        settings = {
            'paper_size': 'A4',
            'orientation': 'landscape',
            'font_size': font_size
        }

        engine = PrintEngine(settings)
        assert engine.settings['font_size'] == font_size


if __name__ == '__main__':
    pytest.main([__file__, '-v'])
