"""首页工作台。"""
import datetime
from pathlib import Path

import customtkinter as ctk

from src.ui.ui_helpers import (
    BORDER_COLOR,
    CARD_BG,
    CARD_SUBTLE_BG,
    FONT_BODY,
    FONT_BODY_BOLD,
    FONT_SECTION,
    FONT_SMALL,
    PRIMARY,
    SUCCESS,
    TEXT_MUTED,
    TEXT_PRIMARY,
    WARNING,
    clear_children,
    make_bottom_action_bar,
    make_check_group,
    make_check_panel,
    make_tab_page,
    set_check_item,
    set_status,
    subtle_button_style,
)


WORKFLOW_STEPS = [
    {
        "title": "总名单",
        "desc": "导入或更新当前助理名单",
        "target": "names",
        "button_text": "维护名单",
    },
    {
        "title": "规则配置",
        "desc": "统一配置毕业季特殊逻辑和长期请假规则",
        "target": "rules",
        "button_text": "配置规则",
    },
    {
        "title": "周统计",
        "desc": "选择排班名单和实际名单，导出 Excel",
        "target": "weekly",
        "button_text": "生成周统计",
    },
    {
        "title": "期末周统计",
        "desc": "拆分双周排班 PDF，匹配两份实际名单",
        "target": "final_weeks",
        "button_text": "生成期末周统计",
    },
    {
        "title": "月统计",
        "desc": "汇总周 Excel，结转上月数据，导出 Word",
        "target": "monthly",
        "button_text": "生成月统计",
    },
]


class HomeTab(ctk.CTkFrame):
    def __init__(self, parent, data_mgr, navigate):
        super().__init__(parent, fg_color="transparent")
        self.data_mgr = data_mgr
        self.navigate = navigate
        self.workflow_status_labels = {}

        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)
        self.grid(row=0, column=0, sticky="nsew")
        self._build()

    def _build(self):
        page = make_tab_page(self, "概览", "查看当前流程状态、最近输出和下一步建议。")
        page.grid_rowconfigure(0, weight=1)
        page.grid_columnconfigure(0, weight=1)
        page.grid_columnconfigure(1, weight=0, minsize=330)

        self._build_flow_list(page)
        self._build_status_panel(page)
        self.home_status = make_bottom_action_bar(
            page,
            row=1,
            primary_text="刷新概览",
            primary_command=self.refresh,
            status_text="等待刷新",
            secondary_buttons=[
                {"text": "去总名单", "width": 104, "command": lambda: self.navigate("names")},
                {"text": "去周统计", "width": 104, "command": lambda: self.navigate("weekly")},
            ],
        )
        self.refresh()

    def _build_flow_list(self, page):
        panel = ctk.CTkFrame(page, fg_color=CARD_BG, corner_radius=8, border_width=1, border_color=BORDER_COLOR)
        panel.grid(row=0, column=0, sticky="nsew", padx=(4, 10), pady=4)
        panel.grid_columnconfigure(0, weight=1)
        panel.grid_rowconfigure(2, weight=1)

        ctk.CTkLabel(panel, text="流程总览", font=FONT_SECTION, text_color=TEXT_PRIMARY, anchor="w").grid(
            row=0, column=0, sticky="ew", padx=16, pady=(14, 2)
        )
        self.home_hint = ctk.CTkLabel(
            panel,
            text="",
            font=FONT_SMALL,
            text_color=TEXT_MUTED,
            anchor="w",
            justify="left",
            wraplength=860,
        )
        self.home_hint.grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 10))

        self.flow_rows = ctk.CTkFrame(panel, fg_color=CARD_SUBTLE_BG, corner_radius=8, border_width=1, border_color=BORDER_COLOR)
        self.flow_rows.grid(row=2, column=0, sticky="nsew", padx=16, pady=(0, 16))
        self.flow_rows.grid_columnconfigure(0, weight=1)
        for index, step in enumerate(WORKFLOW_STEPS):
            self._add_flow_row(index, step)

    def _add_flow_row(self, index, step):
        row = ctk.CTkFrame(self.flow_rows, fg_color=CARD_BG if index % 2 == 0 else CARD_SUBTLE_BG, corner_radius=0)
        row.grid(row=index, column=0, sticky="ew", padx=0, pady=(0, 1))
        row.grid_columnconfigure(2, weight=1)

        marker = ctk.CTkLabel(
            row,
            text=str(index + 1),
            width=30,
            height=30,
            corner_radius=15,
            fg_color=PRIMARY,
            text_color="white",
            font=("Microsoft YaHei UI", 13, "bold"),
        )
        marker.grid(row=0, column=0, rowspan=2, padx=(12, 12), pady=10)
        ctk.CTkLabel(row, text=step["title"], font=FONT_BODY_BOLD, text_color=TEXT_PRIMARY, anchor="w", width=92).grid(
            row=0, column=1, rowspan=2, sticky="w", padx=(0, 10), pady=10
        )
        ctk.CTkLabel(row, text=step["desc"], font=FONT_BODY, text_color=TEXT_PRIMARY, anchor="w").grid(
            row=0, column=2, sticky="ew", pady=(10, 0)
        )
        status = ctk.CTkLabel(row, text="", font=FONT_SMALL, text_color=TEXT_MUTED, anchor="w")
        status.grid(row=1, column=2, sticky="ew", pady=(0, 10))
        self.workflow_status_labels[step["target"]] = status
        ctk.CTkButton(
            row,
            text=step["button_text"],
            width=104,
            height=32,
            command=lambda target=step["target"]: self.navigate(target),
            **subtle_button_style(),
        ).grid(row=0, column=3, rowspan=2, sticky="e", padx=12, pady=10)

    def _build_status_panel(self, page):
        panel = make_check_panel(page, title="当前状态")
        panel.grid(row=0, column=1, sticky="nsew", padx=(0, 4), pady=4)
        self.home_check_items = {}
        make_check_group(
            panel,
            row=1,
            title="数据状态",
            items=[
                ("names", "总名单"),
                ("senior", "毕业季规则"),
                ("leave", "长期请假"),
                ("next", "下一步建议"),
            ],
            registry=self.home_check_items,
        )

        recent = ctk.CTkFrame(panel, fg_color=CARD_SUBTLE_BG, corner_radius=8, border_width=1, border_color=BORDER_COLOR)
        recent.grid(row=2, column=0, sticky="ew", padx=12, pady=(0, 14))
        recent.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(recent, text="最近输出", font=FONT_BODY_BOLD, text_color=TEXT_PRIMARY, anchor="w").grid(
            row=0, column=0, sticky="ew", padx=12, pady=(10, 2)
        )
        ctk.CTkLabel(
            recent,
            text="当前输出目录，按修改时间排序",
            font=FONT_SMALL,
            text_color=TEXT_MUTED,
            anchor="w",
        ).grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 8))
        self.recent_frame = ctk.CTkFrame(recent, fg_color="transparent")
        self.recent_frame.grid(row=2, column=0, sticky="ew", padx=8, pady=(0, 10))
        self.recent_frame.grid_columnconfigure(0, weight=1)

    def refresh(self):
        total_count = self.data_mgr.get_total_count()
        senior_count = len(self.data_mgr.get_senior_assistants())
        leave_count = len(self.data_mgr.get_long_term_leave_assistants())
        senior_mode = {
            "normal": "正常值班",
            "reduced": "少值班",
            "none": "无需值班",
        }.get(self.data_mgr.get_senior_should_mode(), "正常值班")
        recent_files = self._get_recent_output_files(limit=4)

        if total_count <= 0:
            self.home_hint.configure(text="当前流程停在总名单前置步骤：请先导入或添加助理名单。", text_color=WARNING)
        else:
            self.home_hint.configure(text="基础名单已就绪。建议先核对规则配置，再生成周统计和月统计。", text_color=SUCCESS)

        set_check_item(self.home_check_items, "names", f"总名单 {total_count} 人" if total_count else "尚未导入总名单", "success" if total_count else "error")
        set_check_item(self.home_check_items, "senior", f"{senior_mode}，已保存 {senior_count} 人", "success")
        set_check_item(self.home_check_items, "leave", f"长期请假 {leave_count} 人", "success")
        next_text = "继续核对规则，或生成周统计。" if total_count else "先维护总名单。"
        set_check_item(self.home_check_items, "next", next_text, "success" if total_count else "warning")

        self._refresh_workflow_status(total_count, senior_count, leave_count, len(recent_files))
        self._refresh_recent_outputs(recent_files)
        if hasattr(self, "home_status"):
            set_status(self.home_status, "概览已刷新", "success")

    def _refresh_workflow_status(self, total_count, senior_count, leave_count, recent_count):
        self._set_workflow_status(
            "names",
            f"已保存 {total_count} 人" if total_count else "阻塞：尚未导入总名单",
            "success" if total_count else "warning",
        )
        if not total_count:
            rule_text = "等待总名单后再配置规则"
            rule_kind = "warning"
        elif senior_count or leave_count:
            rule_text = f"已配置毕业季 {senior_count} 人，长期请假 {leave_count} 人"
            rule_kind = "success"
        else:
            rule_text = "可选：当前使用默认规则"
            rule_kind = "info"
        self._set_workflow_status("rules", rule_text, rule_kind)
        self._set_workflow_status(
            "weekly",
            "可选择本周 Word 并生成 Excel" if total_count else "等待总名单",
            "success" if total_count else "warning",
        )
        self._set_workflow_status(
            "monthly",
            f"最近输出 {recent_count} 个文件，可继续汇总" if recent_count else "等待周统计 Excel 输出",
            "success" if recent_count else "info",
        )

    def _set_workflow_status(self, target, text, kind):
        colors = {
            "success": SUCCESS,
            "warning": WARNING,
            "info": TEXT_MUTED,
        }
        label = self.workflow_status_labels.get(target)
        if label is not None:
            label.configure(text=text, text_color=colors.get(kind, TEXT_MUTED))

    def _get_recent_output_files(self, limit=5):
        output_dir = Path(self.data_mgr.get_output_dir())
        if not output_dir.exists():
            return []
        return sorted(
            [p for p in output_dir.iterdir() if p.is_file()],
            key=lambda p: p.stat().st_mtime,
            reverse=True,
        )[:limit]

    def _refresh_recent_outputs(self, files):
        clear_children(self.recent_frame)
        if not files:
            ctk.CTkLabel(
                self.recent_frame,
                text="暂无输出文件。完成周统计或月统计后会出现在这里。",
                font=FONT_SMALL,
                text_color=TEXT_MUTED,
                anchor="w",
                justify="left",
                wraplength=270,
            ).grid(row=0, column=0, sticky="ew", padx=4, pady=4)
            return

        for idx, path in enumerate(files):
            stat = path.stat()
            stamp = datetime.datetime.fromtimestamp(stat.st_mtime).strftime("%m-%d %H:%M")
            size = f"{stat.st_size // 1024 + 1} KB"
            row = ctk.CTkFrame(self.recent_frame, fg_color=CARD_BG, corner_radius=6)
            row.grid(row=idx, column=0, sticky="ew", padx=4, pady=3)
            row.grid_columnconfigure(0, weight=1)
            ctk.CTkLabel(
                row,
                text=path.name,
                font=FONT_BODY,
                text_color=TEXT_PRIMARY,
                anchor="w",
                justify="left",
                wraplength=260,
            ).grid(row=0, column=0, sticky="ew", padx=10, pady=(7, 0))
            ctk.CTkLabel(
                row,
                text=f"{size} · {stamp}",
                font=FONT_SMALL,
                text_color=TEXT_MUTED,
                anchor="w",
            ).grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 7))
