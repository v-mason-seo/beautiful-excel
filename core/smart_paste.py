"""
스마트 붙여넣기 엔진 - 헤더명 기반 자동 컬럼 매핑
"""

from typing import Dict, List, Optional, Tuple


class SmartPasteEngine:
    """
    클립보드 텍스트를 분석하여 헤더명 기반으로 그리드 컬럼에 자동 매핑.

    처리 흐름:
      1. parse_clipboard()      - 텍스트 → 2D 리스트
      2. detect_header_rows()   - 상위 몇 행이 헤더인지 판단
      3. flatten_headers()      - 멀티 행 헤더 → 단일 키 문자열
      4. build_column_mapping() - 소스 헤더 → 그리드 컬럼 인덱스 매핑
      5. apply_smart_paste()    - 최종 (row, col, value) 셀 목록 반환
    """

    @staticmethod
    def parse_clipboard(text: str) -> List[List[str]]:
        """탭/줄바꿈으로 구분된 클립보드 텍스트를 2D 리스트로 변환."""
        if not text:
            return []
        rows = text.rstrip("\n").split("\n")
        return [row.replace("\r", " ").split("\t") for row in rows if row.strip() or "\t" in row]

    @staticmethod
    def detect_header_rows(data: List[List[str]], max_header_rows: int = 3) -> int:
        """
        상위 몇 행이 헤더인지 휴리스틱 판단.

        판단 기준: 순수 숫자 비율이 30% 이하면 헤더 행으로 간주.
        """
        if not data:
            return 0

        header_count = 0
        for row in data[:max_header_rows]:
            non_empty = [c.strip() for c in row if c.strip()]
            if not non_empty:
                header_count += 1
                continue

            numeric_count = sum(
                1 for c in non_empty
                if c.replace(",", "").replace(".", "", 1).lstrip("-").isdigit()
            )
            if numeric_count / len(non_empty) > 0.3:
                break
            header_count += 1

        return header_count

    @staticmethod
    def flatten_headers(header_rows: List[List[str]]) -> List[str]:
        """
        여러 행 헤더를 단일 문자열 키로 변환.
        빈 셀(병합 셀)은 왼쪽 위 값을 이어받아 채움.

        예:
          ["서버 정보", "",    "네트워크"]
          ["서버명",   "IP", "VLAN"    ]
          → ["서버 정보/서버명", "서버 정보/IP", "네트워크/VLAN"]
        """
        if not header_rows:
            return []
        if len(header_rows) == 1:
            return [c.strip() for c in header_rows[0]]

        num_cols = max(len(row) for row in header_rows)
        filled_rows: List[List[str]] = []
        for row in header_rows:
            padded = list(row) + [""] * (num_cols - len(row))
            filled, last = [], ""
            for cell in padded:
                val = cell.strip()
                if val:
                    last = val
                    filled.append(val)
                else:
                    filled.append(last)
            filled_rows.append(filled)

        flat: List[str] = []
        for col_idx in range(num_cols):
            parts: List[str] = []
            for row in filled_rows:
                val = row[col_idx] if col_idx < len(row) else ""
                if val and val not in parts:
                    parts.append(val)
            flat.append("/".join(parts))
        return flat

    @staticmethod
    def build_column_mapping(
        source_headers: List[str],
        target_headers: List[str],
        mapping_config: Dict[str, List[str]],
    ) -> Dict[int, int]:
        """
        소스 헤더 인덱스 → 대상(그리드) 컬럼 인덱스 매핑 반환.

        우선순위:
          1. mapping_config 별칭 일치 (대소문자 무시)
          2. 그리드 헤더와 직접 이름 일치 (대소문자 무시)
        """
        alias_to_target: Dict[str, int] = {}
        for grid_col_name, aliases in mapping_config.items():
            target_idx: Optional[int] = None
            for t_idx, t_hdr in enumerate(target_headers):
                if t_hdr.strip().lower() == grid_col_name.strip().lower():
                    target_idx = t_idx
                    break
            if target_idx is None:
                continue
            for alias in aliases:
                alias_to_target[alias.strip().lower()] = target_idx
            alias_to_target[grid_col_name.strip().lower()] = target_idx

        col_map: Dict[int, int] = {}
        for src_idx, src_hdr in enumerate(source_headers):
            src_lower = src_hdr.strip().lower()
            if src_lower in alias_to_target:
                col_map[src_idx] = alias_to_target[src_lower]
            else:
                for t_idx, t_hdr in enumerate(target_headers):
                    if src_lower == t_hdr.strip().lower():
                        col_map[src_idx] = t_idx
                        break
        return col_map

    @staticmethod
    def apply_smart_paste(
        clipboard_text: str,
        target_headers: List[str],
        mapping_config: Dict[str, List[str]],
        grid_start_row: int = 0,
    ) -> Tuple[List[Tuple[int, int, str]], bool]:
        """
        클립보드 텍스트를 분석해 그리드에 쓸 셀 목록을 반환.

        Returns:
            (cells, used_mapping)
            - cells: [(row, col, value), ...]
            - used_mapping: True = 헤더 매핑 적용, False = 위치 기반 fallback
        """
        data = SmartPasteEngine.parse_clipboard(clipboard_text)
        if not data:
            return [], False

        header_count = SmartPasteEngine.detect_header_rows(data)

        if header_count == 0 or not mapping_config or not target_headers:
            cells = [
                (grid_start_row + r, c, cell.strip())
                for r, row in enumerate(data)
                for c, cell in enumerate(row)
            ]
            return cells, False

        flat_headers = SmartPasteEngine.flatten_headers(data[:header_count])
        data_rows = data[header_count:]
        col_map = SmartPasteEngine.build_column_mapping(flat_headers, target_headers, mapping_config)

        if not col_map:
            cells = [
                (grid_start_row + r, c, cell.strip())
                for r, row in enumerate(data_rows)
                for c, cell in enumerate(row)
            ]
            return cells, False

        cells = [
            (grid_start_row + r, col_map[src_col], val.strip())
            for r, row in enumerate(data_rows)
            for src_col, val in enumerate(row)
            if src_col in col_map
        ]
        return cells, True
