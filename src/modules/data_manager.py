# src/modules/data_manager.py
"""
数据存储模块：负责总名单库的本地持久化（JSON）
"""
import json
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
            return self._default_data()
        try:
            with open(self.config_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                if "senior_assistants" not in data:
                    data["senior_assistants"] = []
                if "senior_should_fixed_enabled" not in data:
                    data["senior_should_fixed_enabled"] = False
                if "name_grades" not in data:
                    data["name_grades"] = {}
                return data
        except (json.JSONDecodeError, OSError):
            return self._default_data()

    def _default_data(self):
        return {
            "name_list": [],
            "total_count": 0,
            "senior_assistants": [],
            "senior_should_fixed_enabled": False,
            "name_grades": {},
        }

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
        # 名单变化时，自动清理已不存在的“大四助理”配置
        senior_list = self._data.get("senior_assistants", [])
        self._data["senior_assistants"] = [n for n in senior_list if n in unique]
        grades = self._data.get("name_grades", {})
        self._data["name_grades"] = {
            n: grades.get(n, "")
            for n in unique
            if grades.get(n, "")
        }
        self._save()

    def add_name(self, name: str, grade: str = "") -> bool:
        """手动添加单个姓名；返回是否成功新增。"""
        name = (name or "").strip()
        grade = (grade or "").strip()
        if not name:
            return False
        names = self.get_name_list()
        if name in names:
            return False
        names.append(name)
        self._data["name_list"] = names
        self._data["total_count"] = len(names)
        if grade:
            self._data.setdefault("name_grades", {})[name] = grade
        if grade == "大四":
            seniors = self.get_senior_assistants()
            if name not in seniors:
                seniors.append(name)
            self._data["senior_assistants"] = seniors
        self._save()
        return True

    def get_name_list(self) -> list:
        return list(self._data.get("name_list", []))

    def get_total_count(self) -> int:
        return self._data.get("total_count", 0)

    def get_name_grades(self) -> dict:
        return dict(self._data.get("name_grades", {}))

    def get_name_grade(self, name: str) -> str:
        return self._data.get("name_grades", {}).get(name, "")

    def has_name_list(self) -> bool:
        return self.get_total_count() > 0

    def get_senior_assistants(self) -> list:
        return list(self._data.get("senior_assistants", []))

    def set_senior_assistants(self, names: list):
        total_set = set(self.get_name_list())
        seen = set()
        valid = []
        for n in names or []:
            if n in total_set and n not in seen:
                seen.add(n)
                valid.append(n)
        self._data["senior_assistants"] = valid
        self._save()

    def is_senior_should_fixed_enabled(self) -> bool:
        return bool(self._data.get("senior_should_fixed_enabled", False))

    def set_senior_should_fixed_enabled(self, enabled: bool):
        self._data["senior_should_fixed_enabled"] = bool(enabled)
        self._save()
