import json
import os


class FwsConfigManager:
    """共通 JSON 保存・読み込みクラス"""

    def load_json(self, file_path: str, default_value=None):
        if not os.path.exists(file_path):
            return default_value if default_value is not None else {}

        try:
            with open(file_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return default_value if default_value is not None else {}

    def save_json(self, file_path: str, data) -> bool:
        try:
            with open(file_path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            return True
        except Exception:
            return False