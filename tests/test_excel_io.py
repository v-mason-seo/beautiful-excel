"""
엑셀 입출력 기능 테스트
"""

import os
import sys
import unittest
from pathlib import Path

# 프로젝트 루트를 경로에 추가
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root / 'src'))

from core.excel_loader import ExcelLoader
from core.exporter import ExcelExporter


class TestExcelIO(unittest.TestCase):
    """
    엑셀 입출력 기능 테스트
    """

    def setUp(self):
        """
        테스트 환경 설정
        """
        self.test_dir = Path(__file__).parent
        self.sample_data = [
            ['이름', '나이', '부서', '직급'],
            ['김철수', '30', '개발팀', '대리'],
            ['이영희', '28', '기획팀', '사원'],
            ['박민수', '35', '개발팀', '과장']
        ]

    def test_create_sample_excel(self):
        """
        샘플 엑셀 파일 생성 테스트
        """
        output_path = self.test_dir / 'sample_output.xlsx'

        ExcelExporter.save_to_excel(
            file_path=str(output_path),
            data=self.sample_data[1:],  # 데이터 (헤더 제외)
            headers=self.sample_data[0],  # 헤더
            settings={
                'paper_size': 'A4',
                'orientation': 'landscape',
                'font_size': 10
            }
        )

        # 파일 생성 확인
        self.assertTrue(output_path.exists())
        print(f"✓ 샘플 엑셀 파일 생성 성공: {output_path}")

    def test_load_excel(self):
        """
        엑셀 파일 로드 테스트
        """
        # 먼저 샘플 파일 생성
        sample_path = self.test_dir / 'sample_load_test.xlsx'

        ExcelExporter.save_to_excel(
            file_path=str(sample_path),
            data=self.sample_data[1:],
            headers=self.sample_data[0]
        )

        # 파일 로드
        result = ExcelLoader.load_file(str(sample_path))

        # 검증
        self.assertIsNotNone(result)
        self.assertIn('data', result)
        self.assertIn('headers', result)
        self.assertIn('formatting', result)

        # 헤더 검증
        self.assertEqual(result['headers'], self.sample_data[0])

        # 데이터 행 수 검증 (헤더 포함)
        self.assertEqual(len(result['data']), len(self.sample_data))

        print(f"✓ 엑셀 파일 로드 성공")
        print(f"  - 헤더: {result['headers']}")
        print(f"  - 데이터 행 수: {len(result['data']) - 1}")

        # 정리
        if sample_path.exists():
            sample_path.unlink()

    def test_round_trip(self):
        """
        저장 후 로드 왕복 테스트
        """
        test_path = self.test_dir / 'sample_round_trip.xlsx'

        # 1. 저장
        ExcelExporter.save_to_excel(
            file_path=str(test_path),
            data=self.sample_data[1:],
            headers=self.sample_data[0],
            settings={'font_size': 12}
        )

        # 2. 로드
        result = ExcelLoader.load_file(str(test_path))

        # 3. 검증
        loaded_headers = result['headers']
        loaded_data = result['data'][1:]  # 헤더 제외

        self.assertEqual(loaded_headers, self.sample_data[0])
        self.assertEqual(len(loaded_data), len(self.sample_data) - 1)

        print(f"✓ 왕복 테스트 성공 (저장 → 로드)")

        # 정리
        if test_path.exists():
            test_path.unlink()

    def test_empty_data(self):
        """
        빈 데이터 처리 테스트
        """
        test_path = self.test_dir / 'sample_empty.xlsx'

        # 빈 데이터로 저장
        ExcelExporter.save_to_excel(
            file_path=str(test_path),
            data=[],
            headers=[]
        )

        # 파일이 생성되었는지 확인
        self.assertTrue(test_path.exists())

        # 로드
        result = ExcelLoader.load_file(str(test_path))

        # 빈 데이터 검증
        self.assertEqual(result['data'], [])
        self.assertEqual(result['headers'], [])

        print(f"✓ 빈 데이터 처리 테스트 성공")

        # 정리
        if test_path.exists():
            test_path.unlink()

    def tearDown(self):
        """
        테스트 정리
        """
        # 테스트 파일 정리
        test_files = [
            'sample_output.xlsx',
            'sample_load_test.xlsx',
            'sample_round_trip.xlsx',
            'sample_empty.xlsx'
        ]

        for filename in test_files:
            file_path = self.test_dir / filename
            if file_path.exists():
                try:
                    file_path.unlink()
                except:
                    pass


if __name__ == '__main__':
    print("=" * 60)
    print("엑셀 입출력 기능 테스트 시작")
    print("=" * 60)
    unittest.main(verbosity=2)
