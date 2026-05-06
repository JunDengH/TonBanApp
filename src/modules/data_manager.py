# src/modules/data_manager.py
"""
数据存储模块：负责总名单库的本地持久化（JSON）
"""
import json
import os
from pathlib import Path


class DataManager:
    def __init__(self, config_path=None):
        if config_path is None:
            # 默认存储到项目根目录下的 data/config.json
            root = Path(__file__).resolve().parent.parent.parent
            config_path = root / "data" / "config.json"
        self.config_path = Path(config_path)
        self.config_path.parent.mkdir(parents=True, exist_ok=True)
        self._data = self._load()

    def _load(self):
        """从JSON加载配置，文件不存在则返回默认结构"""
        if not self.config_path.exists():
            return {"name_list": [], "total_count": 0}
        try:
            with open(self.config_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            return {"name_list": [], "total_count": 0}

    def _save(self):
        """持久化到JSON"""
        with open(self.config_path, "w", encoding="utf-8") as f:
            json.dump(self._data, f, ensure_ascii=False, indent=2)

    # ---------- 公共接口 ----------
    def update_name_list(self, names: list):
        """更新名单库"""
        # 去重并保持顺序
        seen = set()
        unique = []
        for n in names:
            if n and n not in seen:
                seen.add(n)
                unique.append(n)
        self._data["name_list"] = unique
        self._data["total_count"] = len(unique)
        self._save()

    def get_name_list(self) -> list:
        return list(self._data.get("name_list", []))

    def get_total_count(self) -> int:
        return self._data.get("total_count", 0)

    def has_name_list(self) -> bool:
        return self.get_total_count() > 0