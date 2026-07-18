# src/ui/main_ui.py
"""应用主窗口组装。"""
import datetime
import tkinter.filedialog as fd
import tkinter as tk
from pathlib import Path

import customtkinter as ctk

try:
    # Private API: lets us flip the global appearance flag without triggering
    # CustomTkinter's full O(all-widgets) repaint. Guarded so a library upgrade
    # that moves this falls back to the (correct but slower) public path.
    from customtkinter.windows.widgets.appearance_mode.appearance_mode_tracker import (
        AppearanceModeTracker,
    )
except Exception:  # pragma: no cover - defensive import
    AppearanceModeTracker = None

try:
    # Base class that owns each widget's per-mode index. We call its
    # _set_appearance_mode directly on window objects (CTk / CTkToplevel) to
    # update the index WITHOUT the window's own title-bar withdraw/deiconify
    # dance, which would corrupt an active modal grab.
    from customtkinter.windows.widgets.appearance_mode.appearance_mode_base_class import (
        CTkAppearanceModeBaseClass,
    )
except Exception:  # pragma: no cover - defensive import
    CTkAppearanceModeBaseClass = None

from src.modules.data_manager import DataManager
from src.ui import ui_helpers as theme_tokens
from src.ui.final_weeks_tab import FinalWeeksTab
from src.ui.home_tab import HomeTab
from src.ui.monthly_tab import MonthlyTab
from src.ui.name_list_tab import NameListTab
from src.ui.rules_tab import RulesTab
from src.ui.ui_helpers import (
    ACTIVE_BORDER,
    APP_BG,
    BORDER_COLOR,
    CARD_BG,
    CONTENT_BG,
    FONT_BODY,
    FONT_BODY_BOLD,
    FONT_SMALL,
    PRIMARY,
    SIDEBAR_ACTIVE,
    SIDEBAR_BG,
    SIDEBAR_HOVER,
    SUBTLE_BUTTON_BG,
    SUBTLE_BUTTON_HOVER,
    TEXT_MUTED,
    TEXT_ON_DARK,
    TEXT_PRIMARY,
    input_style,
    muted_button_style,
    option_menu_style,
    primary_button_style,
    resolve_color,
    subtle_button_style,
)
from src.ui.weekly_tab import WeeklyTab


ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")


DEFAULT_PAGE = "names"
PAGE_ALIASES = {
    "senior": "rules",
    "leave": "rules",
}
NAV_SECTIONS = [
    {
        "title": "核心流程",
        "items": [
            {"key": "names", "icon": "names", "label": "总名单"},
            {"key": "rules", "icon": "rules", "label": "规则配置"},
            {"key": "weekly", "icon": "weekly", "label": "周统计"},
            {"key": "final_weeks", "icon": "final_weeks", "label": "期末周统计"},
            {"key": "monthly", "icon": "monthly", "label": "月统计"},
        ],
    },
    {
        "title": "辅助",
        "items": [
            {"key": "home", "icon": "overview", "label": "概览"},
        ],
    },
]


def normalize_page_key(key):
    return PAGE_ALIASES.get(key, key)


class LineIcon(tk.Canvas):
    """Small dependency-free line icons for the shell chrome."""

    def __init__(self, parent, icon, size, canvas_bg, badge_fill, stroke, clickable=False):
        super().__init__(
            parent,
            width=size,
            height=size,
            bg=resolve_color(canvas_bg),
            bd=0,
            highlightthickness=0,
            cursor="hand2" if clickable else "arrow",
        )
        self.icon = icon
        self.size = size
        # Keep the raw (light, dark) tokens so the icon can be re-resolved when
        # the appearance mode changes; Canvas itself needs single color strings.
        self._canvas_bg_token = canvas_bg
        self._badge_fill_token = badge_fill
        self._stroke_token = stroke
        self.canvas_bg = resolve_color(canvas_bg)
        self.badge_fill = resolve_color(badge_fill)
        self.stroke = resolve_color(stroke)
        self._draw()

    def set_colors(self, badge_fill, stroke, canvas_bg=None):
        self._badge_fill_token = badge_fill
        self._stroke_token = stroke
        self.badge_fill = resolve_color(badge_fill)
        self.stroke = resolve_color(stroke)
        if canvas_bg is not None:
            self._canvas_bg_token = canvas_bg
            self.canvas_bg = resolve_color(canvas_bg)
            self.configure(bg=self.canvas_bg)
        self._draw()

    def refresh_appearance(self):
        """Re-resolve stored color tokens after an appearance-mode switch."""
        self.canvas_bg = resolve_color(self._canvas_bg_token)
        self.badge_fill = resolve_color(self._badge_fill_token)
        self.stroke = resolve_color(self._stroke_token)
        self.configure(bg=self.canvas_bg)
        self._draw()

    def _draw(self):
        self.delete("all")
        s = self.size
        scale = s / 28
        self._rounded_rect(1, 1, s - 1, s - 1, 7 * scale, fill=self.badge_fill, outline="")
        getattr(self, f"_draw_{self.icon}", self._draw_overview)(scale)

    def _stroke_width(self):
        return max(1.7, self.size / 14.5)

    def _rounded_rect(self, x1, y1, x2, y2, radius, **kwargs):
        points = [
            x1 + radius, y1,
            x2 - radius, y1,
            x2, y1,
            x2, y1 + radius,
            x2, y2 - radius,
            x2, y2,
            x2 - radius, y2,
            x1 + radius, y2,
            x1, y2,
            x1, y2 - radius,
            x1, y1 + radius,
            x1, y1,
        ]
        self.create_polygon(points, smooth=True, splinesteps=12, **kwargs)

    def _line(self, *coords, **kwargs):
        width = kwargs.pop("width", self._stroke_width())
        self.create_line(*coords, fill=self.stroke, width=width, capstyle="round", joinstyle="round", **kwargs)

    def _rect(self, x1, y1, x2, y2, **kwargs):
        width = kwargs.pop("width", self._stroke_width())
        self.create_rectangle(x1, y1, x2, y2, outline=self.stroke, width=width, **kwargs)

    def _draw_brand(self, scale):
        self._line(8 * scale, 9 * scale, 20 * scale, 9 * scale, 20 * scale, 19 * scale, 8 * scale, 19 * scale, 8 * scale, 9 * scale)
        self._line(11 * scale, 13 * scale, 17 * scale, 13 * scale)
        self._line(11 * scale, 16 * scale, 15 * scale, 16 * scale)
        self._line(17 * scale, 6 * scale, 21 * scale, 6 * scale, 21 * scale, 10 * scale)

    def _draw_names(self, scale):
        self._rect(8 * scale, 6 * scale, 20 * scale, 22 * scale)
        self._line(11 * scale, 11 * scale, 17 * scale, 11 * scale)
        self._line(11 * scale, 15 * scale, 17 * scale, 15 * scale)
        self._line(11 * scale, 19 * scale, 15 * scale, 19 * scale)
        self._line(11 * scale, 6 * scale, 12 * scale, 4 * scale, 16 * scale, 4 * scale, 17 * scale, 6 * scale)

    def _draw_rules(self, scale):
        self._line(8 * scale, 9 * scale, 20 * scale, 9 * scale)
        self._line(8 * scale, 14 * scale, 20 * scale, 14 * scale)
        self._line(8 * scale, 19 * scale, 20 * scale, 19 * scale)
        self.create_oval(11 * scale, 7 * scale, 15 * scale, 11 * scale, outline=self.stroke, width=self._stroke_width())
        self.create_oval(16 * scale, 12 * scale, 20 * scale, 16 * scale, outline=self.stroke, width=self._stroke_width())
        self.create_oval(9 * scale, 17 * scale, 13 * scale, 21 * scale, outline=self.stroke, width=self._stroke_width())

    def _draw_weekly(self, scale):
        self._line(8 * scale, 8 * scale, 20 * scale, 8 * scale)
        for y in (12, 16, 20):
            self.create_oval(8 * scale, (y - 1) * scale, 10 * scale, (y + 1) * scale, fill=self.stroke, outline=self.stroke)
            self._line(13 * scale, y * scale, 20 * scale, y * scale)

    def _draw_final_weeks(self, scale):
        self._line(7 * scale, 8 * scale, 21 * scale, 8 * scale)
        for offset, y in enumerate((12, 17)):
            self.create_oval(8 * scale, (y - 1) * scale, 10 * scale, (y + 1) * scale, fill=self.stroke, outline=self.stroke)
            self._line(13 * scale, y * scale, 21 * scale, y * scale)
            self._line(13 * scale, (y + 3) * scale, 18 * scale, (y + 3) * scale)

    def _draw_monthly(self, scale):
        self._rect(7 * scale, 8 * scale, 21 * scale, 21 * scale)
        self._line(7 * scale, 12 * scale, 21 * scale, 12 * scale)
        self._line(10 * scale, 5 * scale, 10 * scale, 9 * scale)
        self._line(18 * scale, 5 * scale, 18 * scale, 9 * scale)
        for x in (11, 15, 19):
            self.create_oval((x - 0.8) * scale, 15 * scale, (x + 0.8) * scale, 16.6 * scale, fill=self.stroke, outline=self.stroke)
        for x in (11, 15):
            self.create_oval((x - 0.8) * scale, 18 * scale, (x + 0.8) * scale, 19.6 * scale, fill=self.stroke, outline=self.stroke)

    def _draw_overview(self, scale):
        for x1, y1 in ((8, 8), (16, 8), (8, 16), (16, 16)):
            self._rect(x1 * scale, y1 * scale, (x1 + 4) * scale, (y1 + 4) * scale)

    def _draw_total(self, scale):
        self.create_oval(11 * scale, 7 * scale, 17 * scale, 13 * scale, outline=self.stroke, width=self._stroke_width())
        self._line(8 * scale, 21 * scale, 9.5 * scale, 17 * scale, 14 * scale, 15 * scale, 18.5 * scale, 17 * scale, 20 * scale, 21 * scale)

    def _draw_senior(self, scale):
        self.create_polygon(
            7 * scale, 11 * scale,
            14 * scale, 7 * scale,
            21 * scale, 11 * scale,
            14 * scale, 15 * scale,
            fill="",
            outline=self.stroke,
            width=self._stroke_width(),
        )
        self._line(10 * scale, 14 * scale, 10 * scale, 18 * scale, 14 * scale, 20 * scale, 18 * scale, 18 * scale, 18 * scale, 14 * scale)
        self._line(21 * scale, 11 * scale, 21 * scale, 17 * scale)

    def _draw_leave(self, scale):
        self._rect(8 * scale, 7 * scale, 20 * scale, 21 * scale)
        self._line(8 * scale, 11 * scale, 20 * scale, 11 * scale)
        self._line(11 * scale, 16 * scale, 13 * scale, 18 * scale, 18 * scale, 14 * scale)


class TongBanApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.data_mgr = DataManager()
        self.theme_mode = theme_tokens.apply_theme(self.data_mgr.get_app_theme())
        ctk.set_appearance_mode("Dark" if self.theme_mode == "dark" else "Light")
        self.configure(fg_color=APP_BG)
        self.title("资助服务中心 - 统班应用")
        self._center_window(1280, 920)
        self.minsize(1180, 860)

        self.pages = {}
        self.nav_buttons = {}
        self.refresh_dirty = {}
        # Pages whose widgets still show the previous theme; repainted lazily
        # when navigated to so a theme switch only repaints the visible page.
        self._pages_need_repaint = set()
        self.current_page = None
        self._build_shell()
        self.show_page(DEFAULT_PAGE)

    def _center_window(self, width, height):
        self.update_idletasks()
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = max((screen_width - width) // 2, 0)
        y = max((screen_height - height) // 2, 0)
        self.geometry(f"{width}x{height}+{x}+{y}")

    def _build_shell(self):
        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=0)
        self.grid_rowconfigure(1, weight=1)

        self._build_top_bar()

        sidebar = ctk.CTkFrame(self, width=224, corner_radius=0, fg_color=SIDEBAR_BG)
        self.sidebar = sidebar
        sidebar.grid(row=1, column=0, sticky="nsw")
        sidebar.grid_propagate(False)
        sidebar.grid_columnconfigure(0, weight=1)

        row_index = 0
        for section in NAV_SECTIONS:
            ctk.CTkLabel(
                sidebar,
                text=section["title"],
                font=FONT_SMALL,
                text_color=TEXT_MUTED,
                anchor="w",
            ).grid(row=row_index, column=0, sticky="ew", padx=18, pady=(24 if row_index == 0 else 14, 6))
            row_index += 1
            for item in section["items"]:
                self._add_nav_item(sidebar, row_index, item)
                row_index += 1

        sidebar.grid_rowconfigure(19, weight=1)
        self._add_settings_button(sidebar)

        content = ctk.CTkFrame(self, fg_color=CONTENT_BG, corner_radius=0)
        content.grid(row=1, column=1, sticky="nsew", padx=0, pady=0)
        content.grid_columnconfigure(0, weight=1)
        content.grid_rowconfigure(0, weight=1)
        self.content = content

        self.pages["home"] = HomeTab(content, self.data_mgr, navigate=self.show_page)
        self.pages["names"] = NameListTab(
            content,
            self.data_mgr,
            on_name_list_changed=self._on_name_list_changed,
        )
        self.pages["rules"] = RulesTab(
            content,
            self.data_mgr,
            navigate=self.show_page,
            on_rule_changed=self._on_rule_changed,
        )
        self.pages["weekly"] = WeeklyTab(content, self.data_mgr)
        self.pages["final_weeks"] = FinalWeeksTab(content, self.data_mgr)
        self.pages["monthly"] = MonthlyTab(content, self.data_mgr)
        self.refresh_dirty = {
            "names": False,
            "rules": False,
            "monthly": False,
        }

        for page in self.pages.values():
            page.grid_remove()

        self._refresh_top_bar()

    def _build_top_bar(self):
        top_bar = ctk.CTkFrame(
            self,
            height=80,
            fg_color=CARD_BG,
            corner_radius=0,
            border_width=0,
        )
        self.top_bar = top_bar
        top_bar.grid(row=0, column=0, columnspan=2, sticky="new")
        top_bar.grid_propagate(False)
        top_bar.grid_rowconfigure(0, weight=1)
        top_bar.grid_columnconfigure(0, weight=0)
        top_bar.grid_columnconfigure(1, weight=1)
        top_bar.grid_columnconfigure(2, weight=0)

        brand = ctk.CTkFrame(top_bar, width=224, fg_color="transparent")
        brand.grid(row=0, column=0, sticky="nsew")
        brand.grid_propagate(False)
        brand.grid_rowconfigure(0, weight=1)
        brand.grid_columnconfigure(1, weight=1)

        LineIcon(
            brand,
            icon="brand",
            size=52,
            canvas_bg=CARD_BG,
            badge_fill=PRIMARY,
            stroke=TEXT_ON_DARK,
        ).grid(row=0, column=0, sticky="w", padx=(20, 10))

        title_cluster = ctk.CTkFrame(brand, fg_color="transparent")
        title_cluster.grid(row=0, column=1, sticky="w")

        name_row = ctk.CTkFrame(title_cluster, fg_color="transparent")
        name_row.grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(
            name_row,
            text="TonBan",
            font=("Microsoft YaHei UI", 21, "bold"),
            text_color=TEXT_PRIMARY,
            anchor="w",
        ).grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(
            name_row,
            text=" v1.8 ",
            font=("Microsoft YaHei UI", 11, "bold"),
            text_color=TEXT_MUTED,
            fg_color=SIDEBAR_ACTIVE,
            corner_radius=6,
            height=20,
        ).grid(row=0, column=1, sticky="w", padx=(9, 0))
        ctk.CTkLabel(
            title_cluster,
            text="值班统计工作台",
            font=("Microsoft YaHei UI", 11),
            text_color=TEXT_MUTED,
            anchor="w",
        ).grid(row=1, column=0, sticky="w", pady=(2, 0))

        status_bar = ctk.CTkFrame(top_bar, fg_color="transparent")
        status_bar.grid(row=0, column=1, sticky="w", padx=(30, 18), pady=(12, 12))

        now = datetime.datetime.now()
        self.data_cycle_label = ctk.CTkLabel(
            status_bar,
            text=f"当前数据周期  {now:%Y年%m月}",
            font=("Microsoft YaHei UI", 14, "bold"),
            text_color=TEXT_PRIMARY,
            anchor="w",
        )
        self.data_cycle_label.grid(row=0, column=0, sticky="w", padx=(0, 16))

        ctk.CTkFrame(status_bar, width=1, height=38, fg_color=BORDER_COLOR).grid(row=0, column=1, padx=(0, 18))

        metrics = ctk.CTkFrame(status_bar, fg_color="transparent")
        metrics.grid(row=0, column=2, sticky="w")
        self.top_total_metric = self._make_top_metric(metrics, 0, "total", "总名单", "-")
        self.top_senior_metric = self._make_top_metric(metrics, 1, "senior", "毕业季", "-")
        self.top_leave_metric = self._make_top_metric(metrics, 2, "leave", "长期请假", "-")

        actions = ctk.CTkFrame(top_bar, fg_color="transparent")
        actions.grid(row=0, column=2, sticky="e", padx=(0, 26), pady=(12, 12))
        ctk.CTkButton(
            actions,
            text="刷新",
            width=66,
            height=32,
            command=self._refresh_current_context,
            fg_color=SUBTLE_BUTTON_BG,
            hover_color=SUBTLE_BUTTON_HOVER,
            text_color=TEXT_PRIMARY,
            border_width=1,
            border_color=BORDER_COLOR,
            corner_radius=8,
        ).grid(row=0, column=0)

        self.top_divider = ctk.CTkFrame(self, height=1, fg_color=BORDER_COLOR)
        self.top_divider.grid(
            row=0,
            column=0,
            columnspan=2,
            sticky="sew",
        )

    def _make_top_metric(self, parent, column, icon, label, value):
        item = ctk.CTkFrame(
            parent,
            fg_color="transparent",
        )
        item.grid(row=0, column=column, sticky="w", padx=(0, 28))
        LineIcon(
            item,
            icon=icon,
            size=52,
            canvas_bg=CARD_BG,
            badge_fill=SIDEBAR_ACTIVE,
            stroke=PRIMARY,
        ).grid(row=0, column=0, sticky="w", padx=(0, 12))
        ctk.CTkLabel(
            item,
            text=label,
            font=("Microsoft YaHei UI", 14),
            text_color=TEXT_MUTED,
            anchor="w",
        ).grid(row=0, column=1, sticky="w")
        value_label = ctk.CTkLabel(
            item,
            text=value,
            font=("Microsoft YaHei UI", 16, "bold"),
            text_color=TEXT_PRIMARY,
            anchor="w",
        )
        value_label.grid(row=0, column=2, sticky="w", padx=(9, 0))
        return value_label

    def _refresh_current_context(self):
        self._refresh_top_bar()
        page = self.pages.get(self.current_page)
        refresh = getattr(page, "refresh", None)
        if callable(refresh):
            refresh()

    def _add_settings_button(self, sidebar):
        settings_box = ctk.CTkFrame(sidebar, fg_color="transparent")
        settings_box.grid(row=20, column=0, sticky="sew", padx=18, pady=(0, 18))
        settings_box.grid_columnconfigure(0, weight=1)
        self.settings_button = ctk.CTkButton(
            settings_box,
            text="设置",
            height=36,
            command=self._open_settings_dialog,
            **subtle_button_style(),
        )
        self.settings_button.grid(row=0, column=0, sticky="ew")

    def _open_settings_dialog(self):
        dialog = ctk.CTkToplevel(self)
        dialog.title("设置")
        dialog.configure(fg_color=CONTENT_BG)
        dialog.geometry("560x360")
        dialog.minsize(520, 330)
        dialog.transient(self)
        dialog.grab_set()
        dialog.grid_columnconfigure(0, weight=1)

        header = ctk.CTkFrame(dialog, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=18, pady=(18, 8))
        header.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            header,
            text="应用设置",
            font=("Microsoft YaHei UI", 18, "bold"),
            text_color=TEXT_PRIMARY,
            anchor="w",
        ).grid(row=0, column=0, sticky="ew")
        ctk.CTkLabel(
            header,
            text="设置会立即保存，并用于后续周统计和月统计输出。",
            font=FONT_SMALL,
            text_color=TEXT_MUTED,
            anchor="w",
        ).grid(row=1, column=0, sticky="ew", pady=(2, 0))

        theme_panel = ctk.CTkFrame(dialog, fg_color=CARD_BG, corner_radius=8, border_width=1, border_color=BORDER_COLOR)
        theme_panel.grid(row=1, column=0, sticky="ew", padx=18, pady=(0, 10))
        theme_panel.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(theme_panel, text="外观设置", font=FONT_BODY_BOLD, text_color=TEXT_PRIMARY).grid(
            row=0, column=0, sticky="w", padx=14, pady=14
        )
        theme_var = ctk.StringVar(value="深色" if self.theme_mode == "dark" else "浅色")
        ctk.CTkOptionMenu(
            theme_panel,
            values=["深色", "浅色"],
            variable=theme_var,
            width=120,
            command=lambda value: self._change_theme("dark" if value == "深色" else "light"),
            **option_menu_style(),
        ).grid(row=0, column=1, sticky="e", padx=14, pady=14)

        output_panel = ctk.CTkFrame(dialog, fg_color=CARD_BG, corner_radius=8, border_width=1, border_color=BORDER_COLOR)
        output_panel.grid(row=2, column=0, sticky="ew", padx=18, pady=(0, 10))
        output_panel.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(output_panel, text="输出路径", font=FONT_BODY_BOLD, text_color=TEXT_PRIMARY, anchor="w").grid(
            row=0, column=0, sticky="ew", padx=14, pady=(14, 2)
        )
        ctk.CTkLabel(
            output_panel,
            text="生成文件会直接保存到该文件夹；未设置时使用项目 output 目录。",
            font=FONT_SMALL,
            text_color=TEXT_MUTED,
            anchor="w",
        ).grid(row=1, column=0, sticky="ew", padx=14, pady=(0, 8))

        output_row = ctk.CTkFrame(output_panel, fg_color="transparent")
        output_row.grid(row=2, column=0, sticky="ew", padx=14, pady=(0, 14))
        output_row.grid_columnconfigure(0, weight=1)
        output_var = ctk.StringVar(value=self.data_mgr.get_output_dir())
        output_entry = ctk.CTkEntry(output_row, textvariable=output_var, height=34, **input_style())
        output_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ctk.CTkButton(
            output_row,
            text="选择文件夹",
            width=104,
            height=34,
            command=lambda: self._choose_output_dir(output_var),
            **subtle_button_style(),
        ).grid(row=0, column=1)

        actions = ctk.CTkFrame(dialog, fg_color="transparent")
        actions.grid(row=3, column=0, sticky="ew", padx=18, pady=(0, 18))
        actions.grid_columnconfigure(0, weight=1)
        ctk.CTkButton(
            actions,
            text="恢复默认输出",
            width=126,
            command=lambda: self._reset_output_dir(output_var),
            **muted_button_style(),
        ).grid(row=0, column=0, sticky="w")
        ctk.CTkButton(
            actions,
            text="关闭",
            width=88,
            command=dialog.destroy,
            **primary_button_style(),
        ).grid(row=0, column=1, sticky="e")
        dialog.bind("<Escape>", lambda _event: dialog.destroy())

    def _choose_output_dir(self, output_var):
        initial_dir = Path(self.data_mgr.get_output_dir())
        try:
            initial_dir.mkdir(parents=True, exist_ok=True)
        except OSError:
            initial_dir = Path.cwd()
        selected = fd.askdirectory(
            parent=self,
            title="选择默认输出文件夹",
            initialdir=str(initial_dir),
        )
        if selected:
            self.data_mgr.set_output_dir(selected)
            output_var.set(self.data_mgr.get_output_dir())

    def _reset_output_dir(self, output_var):
        self.data_mgr.set_output_dir("")
        output_var.set(self.data_mgr.get_output_dir())

    def _change_theme(self, theme_name):
        if theme_name == self.theme_mode:
            return
        self.theme_mode = theme_tokens.apply_theme(theme_name)
        self.data_mgr.set_app_theme(self.theme_mode)
        # This runs from the theme OptionMenu's native-menu callback. Repainting
        # every widget synchronously inside that modal callback wedges input
        # focus on Windows (the menu has not finished unposting), which blocks
        # opening or switching to other windows. Defer the repaint until the
        # callback has returned and the event loop is idle.
        self.after_idle(self._apply_theme_mode)

    def _apply_theme_mode(self):
        mode_string = "Dark" if self.theme_mode == "dark" else "Light"

        if AppearanceModeTracker is None:
            # Fallback: correct but repaints every widget (slow). Only used if
            # the private tracker API is unavailable.
            ctk.set_appearance_mode(mode_string)
            self._refresh_line_icons()
            self._refresh_top_bar()
            return

        # Flip the global flag directly so newly created widgets (dialogs,
        # dropdowns) inherit the mode, but WITHOUT update_callbacks() — that is
        # the all-widgets repaint that froze the UI for ~2s. Color tokens are
        # (light, dark) tuples, so CustomTkinter resolves each widget correctly
        # when we repaint it via its own _set_appearance_mode.
        AppearanceModeTracker.appearance_mode = 1 if self.theme_mode == "dark" else 0
        AppearanceModeTracker.appearance_mode_set_by = "user"

        # Repaint only what is on screen: the root window, the persistent chrome,
        # and the current page. The content frame and divider are repainted as
        # single widgets (not their subtrees) because every page lives under
        # content — walking it would repaint all pages and reintroduce the lag.
        self._set_widget_appearance(self, mode_string)
        self._set_widget_appearance(getattr(self, "content", None), mode_string)
        self._set_widget_appearance(getattr(self, "top_divider", None), mode_string)
        self._repaint_subtree(getattr(self, "top_bar", None), mode_string)
        self._repaint_subtree(getattr(self, "sidebar", None), mode_string)

        self._pages_need_repaint = set(self.pages)
        if self.current_page in self.pages:
            self._repaint_subtree(self.pages[self.current_page], mode_string)
            self._pages_need_repaint.discard(self.current_page)

        # Any dialog open right now (e.g. the Settings window this came from)
        # must follow immediately; later dialogs inherit the flag on creation.
        for top in self._open_toplevels():
            self._repaint_subtree(top, mode_string)

        # Plain Tk canvases (the line icons) are outside CustomTkinter's
        # appearance tracker, so re-resolve their colors explicitly. The nav
        # buttons/indicators are NOT re-toggled here: their active-state colors
        # are (light, dark) tuples already re-resolved by the sidebar subtree
        # repaint above, so a second _refresh_nav_buttons pass would only
        # re-draw them redundantly and add lag.
        self._refresh_line_icons()
        self._refresh_top_bar()

    def _set_widget_appearance(self, widget, mode_string):
        # Match by capability, not by CTkBaseClass: CTkScrollableFrame,
        # CTkToplevel and the CTk root are appearance-aware (they own a
        # per-widget appearance mode and resolve (light, dark) tuples from it)
        # but are NOT CTkBaseClass. Skipping them left their background stuck on
        # the previous theme while their children repainted — i.e. the色串 bug.
        setter = getattr(widget, "_set_appearance_mode", None)
        if not callable(setter):
            return

        # Window objects (CTk root / CTkToplevel) must NOT go through their own
        # _set_appearance_mode here: on Windows it retints the title bar via a
        # withdraw()/deiconify() + focus dance. Triggered while the Settings
        # dialog holds a modal grab_set(), that dance corrupts the grab and
        # leaves every button unclickable. Instead update only the per-mode
        # index and re-tint the background (configure(fg_color=...)).
        if isinstance(widget, (ctk.CTk, ctk.CTkToplevel)):
            if CTkAppearanceModeBaseClass is not None:
                try:
                    CTkAppearanceModeBaseClass._set_appearance_mode(widget, mode_string)
                    widget.configure(fg_color=widget._fg_color)
                except Exception:
                    pass
            return

        try:
            setter(mode_string)
        except Exception:
            pass

    def _repaint_subtree(self, root, mode_string):
        if root is None:
            return
        self._set_widget_appearance(root, mode_string)
        for widget in self._walk_widgets(root):
            self._set_widget_appearance(widget, mode_string)

    def _open_toplevels(self):
        return [
            widget
            for widget in self._walk_widgets(self)
            if isinstance(widget, ctk.CTkToplevel) and widget.winfo_exists()
        ]

    def _repaint_page_if_stale(self, key):
        if key not in self._pages_need_repaint:
            return
        mode_string = "Dark" if self.theme_mode == "dark" else "Light"
        self._repaint_subtree(self.pages[key], mode_string)
        self._pages_need_repaint.discard(key)

    def _walk_widgets(self, root):
        for child in root.winfo_children():
            yield child
            yield from self._walk_widgets(child)

    def _refresh_line_icons(self):
        # Every LineIcon lives in the persistent chrome (top bar + sidebar), so
        # only walk those subtrees instead of the whole widget tree on each
        # theme switch / nav refresh.
        for root in (getattr(self, "top_bar", None), getattr(self, "sidebar", None)):
            if root is None:
                continue
            for widget in self._walk_widgets(root):
                if isinstance(widget, LineIcon):
                    widget.refresh_appearance()

    def _refresh_top_bar(self):
        if not hasattr(self, "top_total_metric"):
            return
        total_count = self.data_mgr.get_total_count()
        senior_count = len(self.data_mgr.get_senior_assistants())
        leave_count = len(self.data_mgr.get_long_term_leave_assistants())
        senior_mode = {
            "normal": "正常值班",
            "reduced": "少值班",
            "none": "无需值班",
        }.get(self.data_mgr.get_senior_should_mode(), "正常值班")
        self.top_total_metric.configure(text=f"{total_count} 人")
        self.top_senior_metric.configure(text=f"{senior_mode} · {senior_count} 人")
        self.top_leave_metric.configure(text=f"{leave_count} 人")

    def _add_nav_item(self, sidebar, row, item):
        key = item["key"]
        nav_row = ctk.CTkFrame(sidebar, height=44, fg_color="transparent")
        nav_row.grid(row=row, column=0, sticky="ew", padx=12, pady=3)
        nav_row.grid_propagate(False)
        nav_row.grid_columnconfigure(2, weight=1)
        nav_row.grid_rowconfigure(0, weight=1)

        indicator = ctk.CTkFrame(nav_row, width=3, height=24, corner_radius=2, fg_color="transparent")
        indicator.grid(row=0, column=0, sticky="nsw", padx=(0, 7), pady=10)

        icon = LineIcon(
            nav_row,
            icon=item["icon"],
            size=28,
            canvas_bg=SIDEBAR_BG,
            badge_fill=CARD_BG,
            stroke=TEXT_MUTED,
            clickable=True,
        )
        icon.grid(row=0, column=1, sticky="w", padx=(0, 8), pady=8)

        btn = ctk.CTkButton(
            nav_row,
            text=item["label"],
            height=38,
            corner_radius=8,
            anchor="w",
            font=FONT_BODY,
            fg_color="transparent",
            hover_color=SIDEBAR_HOVER,
            text_color=TEXT_PRIMARY,
            border_width=1,
            border_color=SIDEBAR_BG,
            command=lambda k=key: self.show_page(k),
        )
        btn.grid(row=0, column=2, sticky="nsew")
        icon.bind("<Button-1>", lambda _event, k=key: self.show_page(k))
        self.nav_buttons[key] = {"button": btn, "indicator": indicator, "icon": icon}

    def show_page(self, key):
        requested_key = key
        key = normalize_page_key(key)
        if key not in self.pages:
            return
        page_changed = self.current_page != key
        if page_changed and self.current_page in self.pages:
            self.pages[self.current_page].grid_remove()
        if page_changed:
            self._repaint_page_if_stale(key)
            self.pages[key].grid()
            self.current_page = key
        if key == "rules" and requested_key in ("senior", "leave"):
            self.pages["rules"].select_rule(requested_key)
        self._refresh_nav_buttons()
        if not page_changed:
            return
        if key == "home":
            self.pages["home"].refresh()
        else:
            self._refresh_page_if_dirty(key)

    def _refresh_page_if_dirty(self, key):
        if not self.refresh_dirty.get(key, False):
            return
        refresh = getattr(self.pages.get(key), "refresh", None)
        if callable(refresh):
            refresh()
        self.refresh_dirty[key] = False

    def _refresh_nav_buttons(self):
        for key, entry in self.nav_buttons.items():
            btn = entry["button"]
            indicator = entry["indicator"]
            icon = entry["icon"]
            if key == self.current_page:
                indicator.configure(fg_color=PRIMARY)
                # TEXT_ON_DARK (not literal white): in dark mode PRIMARY is a
                # near-white badge, so a white stroke would be invisible.
                icon.set_colors(PRIMARY, TEXT_ON_DARK)
                btn.configure(
                    fg_color=SIDEBAR_ACTIVE,
                    hover_color=SIDEBAR_ACTIVE,
                    text_color=PRIMARY,
                    border_color=ACTIVE_BORDER,
                    font=FONT_BODY_BOLD,
                )
            else:
                indicator.configure(fg_color="transparent")
                icon.set_colors(CARD_BG, TEXT_MUTED)
                btn.configure(
                    fg_color="transparent",
                    hover_color=SIDEBAR_HOVER,
                    text_color=TEXT_PRIMARY,
                    border_color=SIDEBAR_BG,
                    font=FONT_BODY,
                )

    def _on_name_list_changed(self):
        self.refresh_dirty["rules"] = True
        self.refresh_dirty["monthly"] = True
        self._refresh_top_bar()

    def _on_rule_changed(self):
        self._refresh_top_bar()


def run_app():
    app = TongBanApp()
    app.mainloop()
