# src/ui/main_ui.py
"""应用主窗口组装。"""
import customtkinter as ctk

from src.modules.data_manager import DataManager
from src.ui.home_tab import HomeTab
from src.ui.monthly_tab import MonthlyTab
from src.ui.name_list_tab import NameListTab
from src.ui.senior_rule_tab import SeniorRuleTab
from src.ui.ui_helpers import FONT_SMALL
from src.ui.weekly_tab import WeeklyTab


ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


class TongBanApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("资助服务中心 - 统班应用")
        self._center_window(1280, 920)
        self.minsize(1180, 860)

        self.data_mgr = DataManager()
        self.pages = {}
        self.nav_buttons = {}
        self.refresh_dirty = {}
        self.current_page = None
        self._build_shell()
        self.show_page("home")

    def _center_window(self, width, height):
        self.update_idletasks()
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = max((screen_width - width) // 2, 0)
        y = max((screen_height - height) // 2, 0)
        self.geometry(f"{width}x{height}+{x}+{y}")

    def _build_shell(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        sidebar = ctk.CTkFrame(self, width=220, corner_radius=0)
        sidebar.grid(row=0, column=0, sticky="nsw")
        sidebar.grid_propagate(False)
        sidebar.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            sidebar,
            text="资助服务中心统班",
            font=("Microsoft YaHei", 20, "bold"),
            anchor="w",
            wraplength=170,
            justify="left",
        ).grid(row=0, column=0, sticky="ew", padx=18, pady=(24, 4))
        ctk.CTkLabel(
            sidebar,
            text="名单维护、周统计、月度报表",
            font=FONT_SMALL,
            text_color=("gray35", "gray70"),
            anchor="w",
            wraplength=170,
            justify="left",
        ).grid(row=1, column=0, sticky="ew", padx=18, pady=(0, 18))

        nav_items = [
            ("home", "首页"),
            ("names", "总名单管理"),
            ("senior", "大四助理规则"),
            ("weekly", "周统计生成"),
            ("monthly", "月统计Word生成"),
        ]
        for idx, (key, text) in enumerate(nav_items, start=2):
            btn = ctk.CTkButton(
                sidebar,
                text=text,
                height=38,
                anchor="w",
                command=lambda k=key: self.show_page(k),
            )
            btn.grid(row=idx, column=0, sticky="ew", padx=14, pady=4)
            self.nav_buttons[key] = btn

        ctk.CTkLabel(
            sidebar,
            text="输出目录：output",
            font=FONT_SMALL,
            text_color=("gray35", "gray70"),
            anchor="w",
        ).grid(row=20, column=0, sticky="sew", padx=18, pady=(0, 18))
        sidebar.grid_rowconfigure(19, weight=1)

        content = ctk.CTkFrame(self, fg_color="transparent")
        content.grid(row=0, column=1, sticky="nsew", padx=12, pady=12)
        content.grid_columnconfigure(0, weight=1)
        content.grid_rowconfigure(0, weight=1)
        self.content = content

        self.pages["home"] = HomeTab(content, self.data_mgr, navigate=self.show_page)
        self.pages["names"] = NameListTab(
            content,
            self.data_mgr,
            on_name_list_changed=self._on_name_list_changed,
        )
        self.pages["senior"] = SeniorRuleTab(
            content,
            self.data_mgr,
            navigate=self.show_page,
            on_rule_changed=self._on_rule_changed,
        )
        self.pages["weekly"] = WeeklyTab(content, self.data_mgr)
        self.pages["monthly"] = MonthlyTab(content, self.data_mgr)
        self.refresh_dirty = {
            "names": False,
            "senior": False,
            "monthly": False,
        }

        for page in self.pages.values():
            page.grid_remove()

    def show_page(self, key):
        if key not in self.pages:
            return
        if self.current_page == key:
            return
        if self.current_page in self.pages:
            self.pages[self.current_page].grid_remove()
        self.pages[key].grid()
        self.current_page = key
        self._refresh_nav_buttons()
        if key == "home":
            self.pages["home"].refresh()
        else:
            self._refresh_page_if_dirty(key)

    def _refresh_page_if_dirty(self, key):
        if not self.refresh_dirty.get(key, False):
            return
        if key == "names":
            self.pages["names"].refresh()
        elif key == "senior":
            self.pages["senior"].refresh()
        elif key == "monthly":
            self.pages["monthly"].refresh_names()
        self.refresh_dirty[key] = False

    def _refresh_nav_buttons(self):
        for key, btn in self.nav_buttons.items():
            if key == self.current_page:
                btn.configure(fg_color=("#1f6aa5", "#144870"), text_color="white")
            else:
                btn.configure(
                    fg_color=("gray78", "gray25"),
                    hover_color=("gray70", "gray32"),
                    text_color=("gray10", "gray95"),
                )

    def _on_name_list_changed(self):
        self.refresh_dirty["senior"] = True
        self.refresh_dirty["monthly"] = True

    def _on_rule_changed(self):
        pass


def run_app():
    app = TongBanApp()
    app.mainloop()
