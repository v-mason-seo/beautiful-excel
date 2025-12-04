#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Beautiful Excel - 엑셀 출력 최적화 프로그램
실행 진입점
"""

import sys
from pathlib import Path

# 프로젝트 루트를 Python 경로에 추가
project_root = Path(__file__).parent
src_path = project_root / 'src'
sys.path.insert(0, str(src_path))

# 메인 애플리케이션 실행
if __name__ == '__main__':
    from main import main
    sys.exit(main())
