"""首页工作台。"""
import datetime

import customtkinter as ctk

from src.ui.ui_helpers import (
    CHIP_FG,
    CHIP_MUTED,
    CHIP_SELECTED,
    CHIP_TEXT,
    FONT_BODY,
    FONT_SECTION,
    FONT_SMALL,
    OUTPUT_DIR,
    clear_children,
    make_tab_page,
)


class HomeTab(ctk.CTkFrame):
    def __init__(self, parent, data_mgr, navigate):
        super().__init__(parent, fg_color="transparent")
        self.data_mgr = data_mgr
        self.navigate = navigate
        self.status_values = {}

        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)
        self.grid(row=0, column=0, sticky="nsew")
        self._build()

    def _build(self):
        page = make_tab_page(self, "首页")
        page.grid_columnconfigure(0, weight=7, uniform="home_main")
        page.grid_columnconfigure(1, weight=4, uniform="home_main")
        page.grid_rowconfigure(1, weight=1)

        self._build_status_band(page)
        self._build_workflow_panel(page)
        self._build_recent_panel(page)
        self.refresh()

    def _build_status_band(self, page):
        band = ctk.CTkFrame(page, corner_radius=10)
        band.grid(row=0, column=0, columnspan=2, sticky="ew", padx=4, pady=(2, 10))
        band.grid_columnconfigure(0, weight=1)

        title_box = ctk.CTkFrame(band, fg_color="transparent")
        title_box.grid(row=0, column=0, sticky="ew", padx=18, pady=16)
        title_box.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            title_box,
            text="工作台",
            font=("Microsoft YaHei", 22, "bold"),
            anchor="w",
        ).grid(row=0, column=0, sticky="w")
        self.home_hint = ctk.CTkLabel(
            title_box,
            text="",
            font=FONT_BODY,
            text_color=("gray35", "gray70"),
            anchor="w",
        )
        self.home_hint.grid(row=1, column=0, sticky="w", pady=(3, 0))

        metrics = ctk.CTkFrame(band, fg_color="transparent")
        metrics.grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 14))
        metrics.grid_columnconfigure((0, 1, 2), weight=1, uniform="status")
        for idx, (key, label) in enumerate(
            [
                ("名单人数", "总名单"),
                ("大四规则", "大四规则"),
                ("大四人数", "已选大四"),
            ]
        ):
            item = ctk.CTkFrame(metrics, fg_color=("gray92", "gray18"), corner_radius=8)
            item.grid(row=0, column=idx, sticky="ew", padx=6)
            item.grid_columnconfigure(0, weight=1)
            ctk.CTkLabel(
                item,
                text=label,
                font=FONT_SMALL,
                text_color=("gray35", "gray70"),
                anchor="w",
            ).grid(row=0, column=0, sticky="w", padx=14, pady=(10, 1))
            value = ctk.CTkLabel(
                item,
                text="-",
                font=("Microsoft YaHei", 20, "bold"),
                anchor="w",
            )
            value.grid(row=1, column=0, sticky="w", padx=14, pady=(0, 11))
            self.status_values[key] = value

    def _build_workflow_panel(self, page):
        panel = ctk.CTkFrame(page, corner_radius=10)
        panel.grid(row=1, column=0, sticky="nsew", padx=(4, 7), pady=(0, 4))
        panel.grid_columnconfigure(0, weight=1)
        panel.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(panel, text="常用流程", font=FONT_SECTION).grid(
            row=0, column=0, sticky="w", padx=16, pady=(14, 8)
        )

        steps = ctk.CTkFrame(panel, fg_color="transparent")
        steps.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 12))
        steps.grid_columnconfigure(0, weight=1)
        for idx, (title, desc, target, button_text) in enumerate(
            [
                ("总名单管理", "导入或更新当前助理名单", "names", "维护名单"),
                ("大四助理规则", "选择大四助理并控制规则开关", "senior", "设置规则"),
                ("周统计生成", "选择排班名单和实际名单，导出 Excel", "weekly", "生成周统计"),
                ("月统计Word生成", "汇总周 Excel，结转上月数据，导出 Word", "monthly", "生成月统计"),
            ]
        ):
            self._add_workflow_row(steps, idx, title, desc, target, button_text)

    def _add_workflow_row(self, parent, index, title, desc, target, button_text):
        row = ctk.CTkFrame(parent, fg_color=("gray92", "gray18"), corner_radius=8)
        row.grid(row=index, column=0, sticky="ew", padx=4, pady=5)
        row.grid_columnconfigure(1, weight=1)

        step = ctk.CTkLabel(
            row,
            text=str(index + 1),
            width=34,
            height=34,
            corner_radius=17,
            fg_color=CHIP_SELECTED,
            text_color="white",
            font=("Microsoft YaHei", 15, "bold"),
        )
        step.grid(row=0, column=0, rowspan=2, padx=(12, 10), pady=12)

        ctk.CTkLabel(
            row,
            text=title,
            font=("Microsoft YaHei", 14, "bold"),
            anchor="w",
        ).grid(row=0, column=1, sticky="ew", pady=(11, 0))
        ctk.CTkLabel(
            row,
            text=desc,
            font=FONT_SMALL,
            text_color=("gray35", "gray70"),
            anchor="w",
        ).grid(row=1, column=1, sticky="ew", pady=(0, 11))

        ctk.CTkButton(
            row,
            text=button_text,
            width=104,
            height=32,
            command=lambda t=target: self.navigate(t),
        ).grid(row=0, column=2, rowspan=2, sticky="e", padx=12, pady=12)

    def _build_recent_panel(self, page):
        panel = ctk.CTkFrame(page, corner_radius=10)
        panel.grid(row=1, column=1, sticky="nsew", padx=(7, 4), pady=(0, 4))
        panel.grid_columnconfigure(0, weight=1)
        panel.grid_rowconfigure(2, weight=1)

        ctk.CTkLabel(panel, text="最近输出", font=FONT_SECTION).grid(
            row=0, column=0, sticky="w", padx=16, pady=(14, 2)
        )
        ctk.CTkLabel(
            panel,
            text="output",
            font=FONT_SMALL,
            text_color=("gray35", "gray70"),
            anchor="w",
        ).grid(row=1, column=0, sticky="w", padx=16, pady=(0, 8))

        self.recent_frame = ctk.CTkScrollableFrame(panel, fg_color="transparent")
        self.recent_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 10))

    def refresh(self):
        total_count = self.data_mgr.get_total_count()
        senior_count = len(self.data_mgr.get_senior_assistants())
        senior_enabled = self.data_mgr.is_senior_should_fixed_enabled()
        senior_status = "开启" if senior_enabled else "关闭"

        self.status_values["名单人数"].configure(text=f"{total_count} 人")
        self.status_values["大四规则"].configure(text=senior_status)
        self.status_values["大四人数"].configure(text=f"{senior_count} 人")

        if total_count <= 0:
            self.home_hint.configure(text="请先导入总名单")
        else:
            self.home_hint.configure(text="可以开始生成周统计或月统计")

        self._refresh_recent_outputs()

    def _refresh_recent_outputs(self):
        clear_children(self.recent_frame)
        files = []
        if OUTPUT_DIR.exists():
            files = sorted(
                [p for p in OUTPUT_DIR.iterdir() if p.is_file()],
                key=lambda p: p.stat().st_mtime,
                reverse=True,
            )[:8]

        if not files:
            empty = ctk.CTkFrame(self.recent_frame, fg_color=CHIP_FG, corner_radius=8)
            empty.grid(row=0, column=0, sticky="ew", padx=4, pady=4)
            empty.grid_columnconfigure(0, weight=1)
            ctk.CTkLabel(
                empty,
                text="当前暂无输出文件",
                font=FONT_BODY,
                text_color=CHIP_MUTED,
                anchor="w",
            ).grid(row=0, column=0, sticky="w", padx=12, pady=12)
            return

        for idx, path in enumerate(files):
            stat = path.stat()
            stamp = datetime.datetime.fromtimestamp(stat.st_mtime).strftime("%m-%d %H:%M")
            size = f"{stat.st_size // 1024 + 1} KB"
            self._add_recent_output_row(idx, path.name, f"{size} · {stamp}")

    def _add_recent_output_row(self, row_index, filename, meta):
        row = ctk.CTkFrame(self.recent_frame, fg_color=CHIP_FG, corner_radius=8)
        row.grid(row=row_index, column=0, sticky="ew", padx=4, pady=4)
        row.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            row,
            text=filename,
            font=FONT_BODY,
            text_color=CHIP_TEXT,
            anchor="w",
            wraplength=320,
        ).grid(row=0, column=0, sticky="ew", padx=12, pady=(9, 0))
        ctk.CTkLabel(
            row,
            text=meta,
            font=FONT_SMALL,
            text_color=CHIP_MUTED,
            anchor="w",
        ).grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 9))
