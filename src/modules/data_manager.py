# src/modules/data_manager.py
"""
数据存储模块：负责总名单库的本地持久化（JSON）
"""
import json
from pathlib import Path


SENIOR_MODE_NORMAL = "normal"
SENIOR_MODE_REDUCED = "reduced"
SENIOR_MODE_NONE = "none"
SENIOR_SHOULD_MODES = {SENIOR_MODE_NORMAL, SENIOR_MODE_REDUCED, SENIOR_MODE_NONE}
APP_THEME_DARK = "dark"
APP_THEME_LIGHT = "light"
APP_THEMES = {APP_THEME_DARK, APP_THEME_LIGHT}


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
                mode = data.get("senior_should_mode")
                if mode not in SENIOR_SHOULD_MODES:
                    mode = (
                        SENIOR_MODE_REDUCED
                        if data.get("senior_should_fixed_enabled", False)
                        else SENIOR_MODE_NORMAL
                    )
                data["senior_should_mode"] = mode
                data["senior_should_fixed_enabled"] = mode == SENIOR_MODE_REDUCED
                if "name_grades" not in data:
                    data["name_grades"] = {}
                if "name_majors" not in data:
                    data["name_majors"] = {}
                if "name_departments" not in data:
                    data["name_departments"] = {}
                if "long_term_leave_assistants" not in data:
                    data["long_term_leave_assistants"] = []
                if "final_week_enabled" not in data:
                    data["final_week_enabled"] = False
                else:
                    data["final_week_enabled"] = bool(data["final_week_enabled"])
                if data.get("app_theme") not in APP_THEMES:
                    data["app_theme"] = APP_THEME_DARK
                if not data.get("output_dir"):
                    data["output_dir"] = self._default_output_dir()
                return data
        except (json.JSONDecodeError, OSError):
            return self._default_data()

    def _default_output_dir(self):
        root = Path(__file__).resolve().parent.parent.parent
        return str(root / "output")

    def _default_data(self):
        return {
            "name_list": [],
            "total_count": 0,
            "senior_assistants": [],
            "senior_should_fixed_enabled": False,
            "senior_should_mode": SENIOR_MODE_NORMAL,
            "long_term_leave_assistants": [],
            "final_week_enabled": False,
            "name_grades": {},
            "name_majors": {},
            "name_departments": {},
            "app_theme": APP_THEME_DARK,
            "output_dir": self._default_output_dir(),
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
        # 名单变化时，自动清理已不存在的毕业季特殊逻辑名单配置。
        senior_list = self._data.get("senior_assistants", [])
        self._data["senior_assistants"] = [n for n in senior_list if n in unique]
        leave_list = self._data.get("long_term_leave_assistants", [])
        self._data["long_term_leave_assistants"] = [n for n in leave_list if n in unique]
        grades = self._data.get("name_grades", {})
        self._data["name_grades"] = {
            n: grades.get(n, "")
            for n in unique
            if grades.get(n, "")
        }
        majors = self._data.get("name_majors", {})
        self._data["name_majors"] = {
            n: majors.get(n, "")
            for n in unique
            if majors.get(n, "")
        }
        departments = self._data.get("name_departments", {})
        self._data["name_departments"] = {
            n: departments.get(n, "")
            for n in unique
            if departments.get(n, "")
        }
        self._save()

    def add_name(self, name: str, grade: str = "", major: str = "", department: str = "") -> bool:
        """手动添加单个姓名；返回是否成功新增。毕业季/长期请假等规则由调用方显式设置。"""
        name = (name or "").strip()
        grade = (grade or "").strip()
        major = (major or "").strip()
        department = (department or "").strip()
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
        if major:
            self._data.setdefault("name_majors", {})[name] = major
        if department:
            self._data.setdefault("name_departments", {})[name] = department
        self._save()
        return True

    def apply_contact_roster(self, records: list, seniors: list):
        """用通讯录解析结果原子地重建总名单。

        - name_list/total_count 按 records 顺序、姓名去重重建；
        - name_grades/name_majors/name_departments 按 records 重建（仅写入非空值）；
        - senior_assistants 覆盖为 seniors（仅保留在新名单中的，按名单顺序）；
        - long_term_leave_assistants 按姓名保留（剔除不在新名单的人）。
        """
        names = []
        seen = set()
        grades = {}
        majors = {}
        departments = {}
        for record in records or []:
            name = (record.get("name") or "").strip()
            if not name or name in seen:
                continue
            seen.add(name)
            names.append(name)
            if record.get("grade"):
                grades[name] = str(record["grade"]).strip()
            if record.get("major"):
                majors[name] = str(record["major"]).strip()
            if record.get("department"):
                departments[name] = str(record["department"]).strip()

        name_set = set(names)
        senior_set = set(seniors or [])
        self._data["name_list"] = names
        self._data["total_count"] = len(names)
        self._data["name_grades"] = grades
        self._data["name_majors"] = majors
        self._data["name_departments"] = departments
        self._data["senior_assistants"] = [n for n in names if n in senior_set]
        self._data["long_term_leave_assistants"] = [
            n for n in self.get_long_term_leave_assistants() if n in name_set
        ]
        self._save()

    def name_exists(self, name: str) -> bool:
        return (name or "").strip() in self._data.get("name_list", [])

    def set_name_grade(self, name: str, grade: str):
        """设置/清除某助理的年级。姓名不在总名单时不做任何处理。"""
        name = (name or "").strip()
        grade = (grade or "").strip()
        if name not in self._data.get("name_list", []):
            return
        grades = self._data.setdefault("name_grades", {})
        if grade:
            grades[name] = grade
        else:
            grades.pop(name, None)
        self._save()

    def set_name_major(self, name: str, major: str):
        """设置/清除某助理的专业。姓名不在总名单时不做任何处理。"""
        name = (name or "").strip()
        major = (major or "").strip()
        if name not in self._data.get("name_list", []):
            return
        majors = self._data.setdefault("name_majors", {})
        if major:
            majors[name] = major
        else:
            majors.pop(name, None)
        self._save()

    def set_name_department(self, name: str, department: str):
        """设置/清除某助理的部门。姓名不在总名单时不做任何处理。"""
        name = (name or "").strip()
        department = (department or "").strip()
        if name not in self._data.get("name_list", []):
            return
        departments = self._data.setdefault("name_departments", {})
        if department:
            departments[name] = department
        else:
            departments.pop(name, None)
        self._save()

    def set_name_senior(self, name: str, enabled: bool):
        """加入/移出毕业季特殊逻辑名单。"""
        name = (name or "").strip()
        if name not in self._data.get("name_list", []):
            return
        seniors = self.get_senior_assistants()
        if enabled and name not in seniors:
            seniors.append(name)
        elif not enabled:
            seniors = [n for n in seniors if n != name]
        self._data["senior_assistants"] = seniors
        self._save()

    def set_name_long_term_leave(self, name: str, enabled: bool):
        """加入/移出长期请假名单。"""
        name = (name or "").strip()
        if name not in self._data.get("name_list", []):
            return
        leave = self.get_long_term_leave_assistants()
        if enabled and name not in leave:
            leave.append(name)
        elif not enabled:
            leave = [n for n in leave if n != name]
        self._data["long_term_leave_assistants"] = leave
        self._save()

    def remove_name(self, name: str) -> bool:
        """从总名单及所有关联结构中删除一个助理；返回是否删除成功。"""
        name = (name or "").strip()
        names = self._data.get("name_list", [])
        if name not in names:
            return False
        self._data["name_list"] = [n for n in names if n != name]
        self._data["total_count"] = len(self._data["name_list"])
        self._data["senior_assistants"] = [n for n in self.get_senior_assistants() if n != name]
        self._data["long_term_leave_assistants"] = [
            n for n in self.get_long_term_leave_assistants() if n != name
        ]
        self._data.get("name_grades", {}).pop(name, None)
        self._data.get("name_majors", {}).pop(name, None)
        self._data.get("name_departments", {}).pop(name, None)
        self._save()
        return True

    def rename_name(self, old_name: str, new_name: str) -> bool:
        """重命名助理，保持名单顺序并迁移年级、专业、毕业季、长期请假等所有引用。

        返回 False 的情况：旧名不存在、新名为空、或新名与其他人重名。
        """
        old_name = (old_name or "").strip()
        new_name = (new_name or "").strip()
        names = self._data.get("name_list", [])
        if old_name not in names or not new_name:
            return False
        if new_name != old_name and new_name in names:
            return False
        self._data["name_list"] = [new_name if n == old_name else n for n in names]
        for key in ("senior_assistants", "long_term_leave_assistants"):
            self._data[key] = [new_name if n == old_name else n for n in self._data.get(key, [])]
        for key in ("name_grades", "name_majors", "name_departments"):
            mapping = self._data.get(key, {})
            if old_name in mapping:
                mapping[new_name] = mapping.pop(old_name)
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

    def get_name_majors(self) -> dict:
        return dict(self._data.get("name_majors", {}))

    def get_name_major(self, name: str) -> str:
        return self._data.get("name_majors", {}).get(name, "")

    def get_name_departments(self) -> dict:
        return dict(self._data.get("name_departments", {}))

    def get_name_department(self, name: str) -> str:
        return self._data.get("name_departments", {}).get(name, "")

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
        return self.get_senior_should_mode() == SENIOR_MODE_REDUCED

    def set_senior_should_fixed_enabled(self, enabled: bool):
        self.set_senior_should_mode(SENIOR_MODE_REDUCED if enabled else SENIOR_MODE_NORMAL)

    def get_senior_should_mode(self) -> str:
        mode = self._data.get("senior_should_mode", SENIOR_MODE_NORMAL)
        if mode not in SENIOR_SHOULD_MODES:
            return SENIOR_MODE_NORMAL
        return mode

    def set_senior_should_mode(self, mode: str):
        if mode not in SENIOR_SHOULD_MODES:
            raise ValueError("invalid senior should mode")
        self._data["senior_should_mode"] = mode
        self._data["senior_should_fixed_enabled"] = mode == SENIOR_MODE_REDUCED
        self._save()

    def get_long_term_leave_assistants(self) -> list:
        return list(self._data.get("long_term_leave_assistants", []))

    def set_long_term_leave_assistants(self, names: list):
        total_set = set(self.get_name_list())
        seen = set()
        valid = []
        for n in names or []:
            if n in total_set and n not in seen:
                seen.add(n)
                valid.append(n)
        self._data["long_term_leave_assistants"] = valid
        self._save()

    def get_final_week_enabled(self) -> bool:
        return bool(self._data.get("final_week_enabled", False))

    def set_final_week_enabled(self, enabled: bool):
        self._data["final_week_enabled"] = bool(enabled)
        self._save()

    def get_app_theme(self) -> str:
        theme = self._data.get("app_theme", APP_THEME_DARK)
        return theme if theme in APP_THEMES else APP_THEME_DARK

    def set_app_theme(self, theme_name: str):
        if theme_name not in APP_THEMES:
            raise ValueError("invalid app theme")
        self._data["app_theme"] = theme_name
        self._save()

    def get_output_dir(self) -> str:
        output_dir = (self._data.get("output_dir") or "").strip()
        return output_dir or self._default_output_dir()

    def set_output_dir(self, path: str):
        clean_path = (path or "").strip()
        if not clean_path:
            clean_path = self._default_output_dir()
        self._data["output_dir"] = str(Path(clean_path))
        self._save()
