"""
앱 전역 설정 - JSON 기반 영속 저장 (~/.excelflow/config.json)
"""

import json
from pathlib import Path
from typing import Any, Dict, Optional

_PACKAGE_DIR = Path(__file__).parent
_DEFAULT_MAPPING_FILE = _PACKAGE_DIR / "mapping_config.json"

DEFAULT_CONFIG: Dict[str, Any] = {
    "save_path": r"C:\sr_result",
    "smart_paste_enabled": True,
}


class AppConfig:
    CONFIG_DIR = Path.home() / ".excelflow"
    CONFIG_FILE = CONFIG_DIR / "config.json"

    def __init__(self):
        self._config: Dict[str, Any] = {}
        self._mapping: Dict[str, Any] = {}
        self.load()

    def load(self):
        if _DEFAULT_MAPPING_FILE.exists():
            with open(_DEFAULT_MAPPING_FILE, "r", encoding="utf-8") as f:
                self._mapping = json.load(f)
        else:
            self._mapping = {}

        if self.CONFIG_FILE.exists():
            try:
                with open(self.CONFIG_FILE, "r", encoding="utf-8") as f:
                    saved = json.load(f)
                self._config = {**DEFAULT_CONFIG, **{k: v for k, v in saved.items() if k != "mapping"}}
                for sr_type, col_map in saved.get("mapping", {}).items():
                    if sr_type in self._mapping:
                        self._mapping[sr_type].update(col_map)
                    else:
                        self._mapping[sr_type] = col_map
            except Exception:
                self._config = dict(DEFAULT_CONFIG)
        else:
            self._config = dict(DEFAULT_CONFIG)

    def save(self):
        self.CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        payload = dict(self._config)
        payload["mapping"] = self._mapping
        with open(self.CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)

    def get(self, key: str, default: Any = None) -> Any:
        return self._config.get(key, default)

    def set(self, key: str, value: Any):
        self._config[key] = value

    def get_mapping(self, sr_type: Optional[str] = None) -> Dict:
        if sr_type:
            return self._mapping.get(sr_type, {})
        return self._mapping

    def get_mapping_json(self) -> str:
        return json.dumps(self._mapping, ensure_ascii=False, indent=2)

    def set_mapping_from_json(self, json_str: str):
        """JSON 문자열 파싱 후 매핑 갱신. 형식 오류 시 ValueError."""
        parsed = json.loads(json_str)
        if not isinstance(parsed, dict):
            raise ValueError("최상위 값은 JSON 객체(dict) 이어야 합니다.")
        self._mapping = parsed
