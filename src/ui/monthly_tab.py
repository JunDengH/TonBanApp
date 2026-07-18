"""月统计 Word 生成 Tab。"""
import datetime
import os
import tkinter.filedialog as fd
import tkinter.messagebox as mb
from pathlib import Path

import customtkinter as ctk

from src.modules.overtime_word import (
    export_overtime_word,
    export_penalty_word,
    import_overtime_word,
    import_penalty_word,
)
from src.modules.word_generator import (
    build_monthly_rows,
    collect_monthly_warnings,
    generate_monthly_word_from_rows,
)
from src.ui.name_picker import NamePickerPanel
from src.ui.ui_helpers import (
    BORDER_COLOR,
    CARD_BG,
    CARD_SUBTLE_BG,
    CONTENT_BG,
    FONT_BODY,
    FONT_BODY_BOLD,
    FONT_SMALL,
    SUCCESS,
    TEXT_MUTED,
    TEXT_PRIMARY,
    add_record_row,
    bind_focus_border,
    clear_children,
    input_style,
    make_check_group,
    make_dialog_header,
    make_section,
    make_tab_page,
    muted_button_style,
    option_menu_style,
    primary_button_style,
    set_check_item,
    set_status,
    subtle_button_style,
    WARNING,
)
from src.utils.output_paths import build_unique_output_path
from src.utils.report_titles import build_monthly_report_title


def summarize_adjustment_records(records, unit, empty_text):
    records = {
        name: times
        for name, times in (records or {}).items()
        if isinstance(times, int) and times > 0
    }
    if not records:
        return empty_text
    return f"已记录 {len(records)} 人，共 {sum(records.values())} {unit}"


MONTHLY_LAYOUT_SECTIONS = ["周统计文件", "报表设置", "月度调整", "生成Word"]
ADJUSTMENT_PANELS = ["加班补录", "罚班记录"]


class MonthlyTab(ctk.CTkFrame):
    def __init__(self, parent, data_mgr):
        super().__init__(parent, fg_color="transparent")
        self.data_mgr = data_mgr
        self.month_overtime_records = {}
        self.month_penalty_records = {}
        self.month_overtime_records_frame = None
        self.month_penalty_records_frame = None
        self.month_overtime_picker = None
        self.month_penalty_picker = None

        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)
        self.grid(row=0, column=0, sticky="nsew")
        self._build()

    def _build(self):
        page = make_tab_page(
            self,
            "月统计",
            subtitle="汇总周统计 Excel，结转上月数据，直接维护加班补录和罚班记录，最后生成 Word 月报。",
        )
        page.grid_rowconfigure(0, weight=1)
        page.grid_columnconfigure(0, weight=1)
        page.grid_columnconfigure(1, weight=0, minsize=330)

        workarea = ctk.CTkScrollableFrame(page, fg_color="transparent")
        workarea.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        workarea.grid_columnconfigure(0, weight=1)

        self._build_file_section(workarea)
        self._build_settings_section(workarea)
        self._build_adjustments_section(workarea)
        self._build_check_panel(page)
        self._build_actions(page)

        self.refresh_names()
        self._refresh_month_overtime_textbox()
        self._refresh_month_penalty_textbox()
        self._refresh_monthly_check_panel()

    def _build_file_section(self, page):
        files, r = make_section(
            page,
            "周统计文件",
            "至少选择 1 份、最多 4 份周统计 Excel；上月 Word 用于继承多值班次和缺班数量。",
        )
        files.grid(row=0, column=0, sticky="ew", padx=4, pady=(0, 6))
        files.grid_columnconfigure(0, weight=1)

        self.week_slots = []
        self.week_file_display_vars = []
        self.week_file_status_labels = []

        table = ctk.CTkFrame(
            files,
            fg_color=CARD_SUBTLE_BG,
            corner_radius=8,
            border_width=1,
            border_color=BORDER_COLOR,
        )
        table.grid(row=r, column=0, sticky="ew", padx=16, pady=(4, 10))
        table.grid_columnconfigure(1, weight=1)

        headers = [("周次", 76), ("Excel 文件", 0), ("状态", 84), ("操作", 132)]
        for col, (title, width) in enumerate(headers):
            ctk.CTkLabel(
                table,
                text=title,
                width=width or 1,
                font=FONT_SMALL,
                text_color=TEXT_MUTED,
                anchor="w",
            ).grid(row=0, column=col, sticky="ew", padx=(14 if col == 0 else 6, 8), pady=(10, 6))

        for i in range(4):
            var = ctk.StringVar(value="")
            self.week_slots.append(var)
            display_var = ctk.StringVar(value=f"选择第{i + 1}周周统计 Excel")
            self.week_file_display_vars.append(display_var)
            self._add_week_file_row(table, i, var, display_var)
            var.trace_add("write", lambda *_args, idx=i: self._sync_week_file_display(idx))

        self.prev_word_path = ctk.StringVar(value="")
        self.prev_word_display_var = ctk.StringVar(value="首次运行或无上月数据时可留空")
        self.prev_word_path.trace_add("write", lambda *_: self._sync_prev_word_display())
        prev_line = ctk.CTkFrame(files, fg_color="transparent")
        prev_line.grid(row=r + 1, column=0, sticky="ew", padx=16, pady=(2, 12))
        prev_line.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(
            prev_line,
            text="上月Word",
            width=80,
            anchor="w",
            font=FONT_BODY,
            text_color=TEXT_PRIMARY,
        ).grid(
            row=0, column=0, sticky="w", padx=(0, 8)
        )
        prev_display = ctk.CTkFrame(
            prev_line,
            height=36,
            corner_radius=8,
            fg_color=CARD_SUBTLE_BG,
            border_width=1,
            border_color=BORDER_COLOR,
        )
        prev_display.grid(row=0, column=1, sticky="ew", padx=(0, 8))
        prev_display.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            prev_display,
            textvariable=self.prev_word_display_var,
            font=FONT_BODY,
            text_color=TEXT_MUTED,
            anchor="w",
        ).grid(row=0, column=0, sticky="ew", padx=12, pady=7)
        ctk.CTkButton(
            prev_line,
            text="清空",
            width=62,
            height=36,
            command=lambda: self.prev_word_path.set(""),
            **subtle_button_style(),
        ).grid(
            row=0, column=2, padx=(0, 6)
        )
        ctk.CTkButton(
            prev_line,
            text="选择",
            width=84,
            height=36,
            command=self._on_pick_prev_word,
            **primary_button_style(),
        ).grid(
            row=0, column=3
        )

    def _sync_prev_word_display(self):
        raw_value = self.prev_word_path.get().strip()
        self.prev_word_display_var.set(Path(raw_value).name if raw_value else "首次运行或无上月数据时可留空")
        self._refresh_monthly_check_panel()

    def _add_week_file_row(self, parent, index, variable, display_var):
        row = index + 1
        ctk.CTkLabel(
            parent,
            text=f"第{index + 1}周",
            width=76,
            font=FONT_BODY,
            text_color=TEXT_PRIMARY,
            anchor="w",
        ).grid(row=row, column=0, sticky="ew", padx=(14, 6), pady=6)
        ctk.CTkLabel(
            parent,
            textvariable=display_var,
            font=FONT_BODY,
            text_color=TEXT_MUTED,
            anchor="w",
        ).grid(row=row, column=1, sticky="ew", padx=6, pady=6)
        status = ctk.CTkLabel(
            parent,
            text="未选择",
            width=84,
            font=FONT_SMALL,
            text_color=WARNING,
            anchor="w",
        )
        status.grid(row=row, column=2, sticky="ew", padx=6, pady=6)
        self.week_file_status_labels.append(status)

        actions = ctk.CTkFrame(parent, fg_color="transparent")
        actions.grid(row=row, column=3, sticky="e", padx=(0, 10), pady=5)
        ctk.CTkButton(
            actions,
            text="选择",
            width=62,
            height=28,
            command=lambda idx=index: self._on_pick_month_file(idx),
            **subtle_button_style(),
        ).grid(row=0, column=0, padx=(0, 6))
        ctk.CTkButton(
            actions,
            text="清空",
            width=58,
            height=28,
            command=lambda var=variable: var.set(""),
            **muted_button_style(),
        ).grid(row=0, column=1)

    def _sync_week_file_display(self, idx):
        raw_value = self.week_slots[idx].get().strip()
        if raw_value:
            self.week_file_display_vars[idx].set(Path(raw_value).name)
            self.week_file_status_labels[idx].configure(text="已选择", text_color=SUCCESS)
        else:
            self.week_file_display_vars[idx].set(f"选择第{idx + 1}周周统计 Excel")
            self.week_file_status_labels[idx].configure(text="未选择", text_color=WARNING)
        self._refresh_monthly_check_panel()

    def _build_settings_section(self, page):
        settings, sr = make_section(page, "报表设置", "确认报表标题；加班补录和罚班记录在下方直接维护。")
        settings.grid(row=1, column=0, sticky="ew", padx=4, pady=(0, 6))

        title_row = ctk.CTkFrame(settings, fg_color="transparent")
        title_row.grid(row=sr, column=0, sticky="ew", padx=16, pady=(0, 8))
        title_row.grid_columnconfigure(7, weight=1)
        now = datetime.datetime.now()
        default_season = "春季" if 2 <= now.month <= 7 else "秋季"
        self.title_year_var = ctk.StringVar(value=str(now.year))
        self.title_season_var = ctk.StringVar(value=default_season)
        self.title_week_start_var = ctk.StringVar(value="1")
        self.title_week_end_var = ctk.StringVar(value="4")
        for title_var in (
            self.title_year_var,
            self.title_season_var,
            self.title_week_start_var,
            self.title_week_end_var,
        ):
            title_var.trace_add("write", lambda *_: self._refresh_monthly_check_panel())
        ctk.CTkLabel(
            title_row,
            text="报表标题",
            width=72,
            anchor="w",
            font=FONT_BODY,
            text_color=TEXT_PRIMARY,
        ).grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        year_entry = ctk.CTkEntry(title_row, textvariable=self.title_year_var, width=76, **input_style())
        bind_focus_border(year_entry)
        year_entry.grid(row=0, column=1, padx=(0, 6), pady=4)
        ctk.CTkOptionMenu(
            title_row,
            values=["春季", "秋季"],
            variable=self.title_season_var,
            width=82,
            **option_menu_style(),
        ).grid(
            row=0, column=2, padx=(0, 6), pady=4
        )
        week_start_entry = ctk.CTkEntry(title_row, textvariable=self.title_week_start_var, width=54, **input_style())
        bind_focus_border(week_start_entry)
        week_start_entry.grid(row=0, column=3, padx=(0, 3), pady=4)
        ctk.CTkLabel(title_row, text="-", font=FONT_BODY, text_color=TEXT_MUTED).grid(
            row=0, column=4, padx=(0, 3), pady=4
        )
        week_end_entry = ctk.CTkEntry(title_row, textvariable=self.title_week_end_var, width=54, **input_style())
        bind_focus_border(week_end_entry)
        week_end_entry.grid(row=0, column=5, padx=(0, 6), pady=4)
        ctk.CTkLabel(title_row, text="周助理值班统计", font=FONT_BODY, text_color=TEXT_MUTED).grid(
            row=0, column=6, sticky="w", pady=4
        )

        self.month_overtime_name_var = ctk.StringVar(value="")
        self.month_overtime_shifts_var = ctk.StringVar(value="1")
        self.month_overtime_search_var = ctk.StringVar(value="")
        self.month_overtime_search_var.trace_add("write", lambda *_: self.refresh_names())
        self.month_penalty_name_var = ctk.StringVar(value="")
        self.month_penalty_shifts_var = ctk.StringVar(value="1")
        self.month_penalty_search_var = ctk.StringVar(value="")
        self.month_penalty_search_var.trace_add("write", lambda *_: self.refresh_names())

    def _build_adjustments_section(self, page):
        adjustments, ar = make_section(
            page,
            "月度调整",
            "加班补录计入总计班次；罚班记录计入应值班次。导入会覆盖当前对应记录。",
        )
        adjustments.grid(row=2, column=0, sticky="ew", padx=4, pady=(0, 6))
        adjustments.grid_columnconfigure((0, 1), weight=1, uniform="monthly_adjustments")

        overtime_panel = self._make_adjustment_panel(adjustments, ar, 0, ADJUSTMENT_PANELS[0])
        self._build_adjustment_controls(
            overtime_panel,
            name_var=self.month_overtime_name_var,
            count_var=self.month_overtime_shifts_var,
            count_placeholder="班次",
            add_command=self._on_add_month_overtime,
            import_command=self._on_import_month_overtime,
            export_command=self._on_export_month_overtime,
            clear_command=self._on_clear_month_overtime,
            picker_attr="month_overtime_picker",
        )
        self.month_overtime_summary = self._make_adjustment_summary(overtime_panel)
        self.month_overtime_records_frame = self._make_adjustment_record_list(overtime_panel)

        penalty_panel = self._make_adjustment_panel(adjustments, ar, 1, ADJUSTMENT_PANELS[1])
        self._build_adjustment_controls(
            penalty_panel,
            name_var=self.month_penalty_name_var,
            count_var=self.month_penalty_shifts_var,
            count_placeholder="次数",
            add_command=self._on_add_month_penalty,
            import_command=self._on_import_month_penalty,
            export_command=self._on_export_month_penalty,
            clear_command=self._on_clear_month_penalty,
            picker_attr="month_penalty_picker",
        )
        self.month_penalty_summary = self._make_adjustment_summary(penalty_panel)
        self.month_penalty_records_frame = self._make_adjustment_record_list(penalty_panel)

    def _make_adjustment_panel(self, parent, row, column, title):
        panel = ctk.CTkFrame(
            parent,
            fg_color=CARD_SUBTLE_BG,
            corner_radius=8,
            border_width=1,
            border_color=BORDER_COLOR,
        )
        panel.grid(row=row, column=column, sticky="nsew", padx=(16 if column == 0 else 8, 16), pady=(2, 14))
        panel.grid_columnconfigure(0, weight=1)
        panel.grid_rowconfigure(4, weight=1)
        ctk.CTkLabel(
            panel,
            text=title,
            font=FONT_BODY_BOLD,
            text_color=TEXT_PRIMARY,
            anchor="w",
        ).grid(row=0, column=0, sticky="w", padx=12, pady=(12, 4))
        return panel

    def _build_adjustment_controls(
        self,
        panel,
        name_var,
        count_var,
        count_placeholder,
        add_command,
        import_command,
        export_command,
        clear_command,
        picker_attr,
    ):
        controls = ctk.CTkFrame(panel, fg_color="transparent")
        controls.grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 8))
        controls.grid_columnconfigure(0, weight=1)

        picker = NamePickerPanel(
            controls,
            mode="single",
            height=118,
            placeholder="搜索姓名",
            on_selection_change=lambda selected, var=name_var: self._set_adjustment_name(var, selected),
        )
        picker.grid(row=0, column=0, sticky="ew")
        setattr(self, picker_attr, picker)

        add_row = ctk.CTkFrame(controls, fg_color="transparent")
        add_row.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        add_row.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            add_row,
            textvariable=name_var,
            font=FONT_SMALL,
            text_color=TEXT_MUTED,
            anchor="w",
        ).grid(row=0, column=0, sticky="ew", padx=(0, 8))
        count_entry = ctk.CTkEntry(
            add_row,
            textvariable=count_var,
            placeholder_text=count_placeholder,
            width=70,
            height=34,
            **input_style(),
        )
        bind_focus_border(count_entry)
        count_entry.grid(row=0, column=1, padx=(0, 8))
        ctk.CTkButton(
            add_row,
            text="添加",
            width=72,
            height=34,
            command=add_command,
            **primary_button_style(),
        ).grid(row=0, column=2)

        tools = ctk.CTkFrame(panel, fg_color="transparent")
        tools.grid(row=2, column=0, sticky="ew", padx=12, pady=(0, 8))
        tools.grid_columnconfigure(3, weight=1)
        ctk.CTkButton(tools, text="导入", width=70, height=30, command=import_command, **subtle_button_style()).grid(
            row=0, column=0, padx=(0, 8)
        )
        ctk.CTkButton(tools, text="导出", width=70, height=30, command=export_command, **subtle_button_style()).grid(
            row=0, column=1, padx=(0, 8)
        )
        ctk.CTkButton(tools, text="清空", width=70, height=30, command=clear_command, **muted_button_style()).grid(
            row=0, column=2
        )

    def _make_adjustment_summary(self, panel):
        label = ctk.CTkLabel(
            panel,
            text="",
            font=FONT_SMALL,
            text_color=TEXT_MUTED,
            anchor="w",
        )
        label.grid(row=3, column=0, sticky="ew", padx=12, pady=(0, 6))
        return label

    def _make_adjustment_record_list(self, panel):
        frame = ctk.CTkScrollableFrame(
            panel,
            height=126,
            fg_color="transparent",
        )
        frame.grid(row=4, column=0, sticky="nsew", padx=8, pady=(0, 10))
        return frame

    def _build_check_panel(self, page):
        panel = ctk.CTkScrollableFrame(
            page,
            width=318,
            fg_color=CARD_BG,
            corner_radius=8,
            border_width=1,
            border_color=BORDER_COLOR,
        )
        panel.grid(row=0, column=1, sticky="nsew", padx=(0, 4), pady=4)
        panel.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            panel,
            text="生成前检查",
            font=("Microsoft YaHei UI", 20, "bold"),
            text_color=TEXT_PRIMARY,
            anchor="w",
        ).grid(row=0, column=0, sticky="ew", padx=18, pady=(18, 12))

        self.monthly_check_items = {}
        self._add_check_group(
            panel,
            row=1,
            title="数据完整性检查",
            items=[
                ("names", "总名单状态"),
                ("weekly_count", "周统计 Excel 数量"),
                ("weekly_paths", "周统计文件读取"),
                ("previous_word", "上月月报状态"),
                ("title", "报表标题字段"),
            ],
        )
        self._add_check_group(
            panel,
            row=2,
            title="规则与优先级检查",
            items=[
                ("senior_rule", "毕业季特殊逻辑"),
                ("leave_rule", "长期请假规则"),
                ("adjustments", "月度调整记录"),
            ],
        )
        self._add_check_group(
            panel,
            row=3,
            title="生成后内容预览要点",
            items=[
                ("preview_rows", "各周数据汇总与人员出勤统计"),
                ("preview_adjustments", "加班补录 / 罚班记录明细汇总"),
                ("preview_sort", "最终值班人次汇总与姓名排序"),
            ],
            static_kind="info",
        )

        next_box = ctk.CTkFrame(
            panel,
            fg_color=CARD_SUBTLE_BG,
            corner_radius=8,
            border_width=1,
            border_color=BORDER_COLOR,
        )
        next_box.grid(row=4, column=0, sticky="ew", padx=12, pady=(4, 14))
        next_box.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            next_box,
            text="下一步操作",
            font=FONT_BODY_BOLD,
            text_color=TEXT_PRIMARY,
            anchor="w",
        ).grid(row=0, column=0, sticky="ew", padx=12, pady=(10, 2))
        self.monthly_next_action_label = ctk.CTkLabel(
            next_box,
            text="请先补齐必要输入。",
            font=FONT_SMALL,
            text_color=TEXT_MUTED,
            anchor="w",
            justify="left",
            wraplength=270,
        )
        self.monthly_next_action_label.grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 10))

    def _add_check_group(self, parent, row, title, items, static_kind=None):
        make_check_group(
            parent,
            row=row,
            title=title,
            items=items,
            registry=self.monthly_check_items,
            static_kind=static_kind,
        )

    def _set_check_item(self, key, text, kind):
        set_check_item(getattr(self, "monthly_check_items", None), key, text, kind)

    def _refresh_monthly_check_panel(self):
        if not hasattr(self, "monthly_check_items"):
            return
        total_count = self.data_mgr.get_total_count()
        self._set_check_item(
            "names",
            f"总名单 {total_count} 人" if total_count else "尚未导入总名单",
            "success" if total_count else "error",
        )

        selected_paths = [slot.get().strip() for slot in getattr(self, "week_slots", []) if slot.get().strip()]
        selected_count = len(selected_paths)
        self._set_check_item(
            "weekly_count",
            f"已选择 {selected_count} / 4 份周统计 Excel",
            "success" if 1 <= selected_count <= 4 else "error",
        )
        missing_paths = [path for path in selected_paths if not os.path.exists(path)]
        if not selected_paths:
            self._set_check_item("weekly_paths", "等待选择周统计 Excel 文件", "warning")
        elif missing_paths:
            self._set_check_item("weekly_paths", f"{len(missing_paths)} 个文件路径不可读取", "error")
        else:
            self._set_check_item("weekly_paths", "已选择的周统计文件路径可读取", "success")

        prev_path = self.prev_word_path.get().strip() if hasattr(self, "prev_word_path") else ""
        if not prev_path:
            self._set_check_item("previous_word", "未选择上月 Word，将按首次运行处理", "warning")
        elif os.path.exists(prev_path):
            self._set_check_item("previous_word", "上月 Word 已选择，可用于结转", "success")
        else:
            self._set_check_item("previous_word", "上月 Word 路径不可读取", "error")

        title_ready = True
        title_message = "标题字段合法"
        try:
            self._build_forced_monthly_title_from_inputs()
        except Exception as exc:
            title_ready = False
            title_message = str(exc)
        self._set_check_item("title", title_message, "success" if title_ready else "error")

        senior_mode = {
            "normal": "正常值班",
            "reduced": "少值班",
            "none": "无需值班",
        }.get(self.data_mgr.get_senior_should_mode(), "正常值班")
        senior_count = len(self.data_mgr.get_senior_assistants())
        leave_count = len(self.data_mgr.get_long_term_leave_assistants())
        self._set_check_item("senior_rule", f"{senior_mode}，已保存 {senior_count} 人", "success")
        self._set_check_item("leave_rule", f"长期请假 {leave_count} 人；若重叠，毕业季规则优先", "success")

        overtime_count = sum(self.month_overtime_records.values()) if self.month_overtime_records else 0
        penalty_count = sum(self.month_penalty_records.values()) if self.month_penalty_records else 0
        adjustment_text = f"加班补录 {overtime_count} 班，罚班记录 {penalty_count} 次"
        self._set_check_item("adjustments", adjustment_text, "success" if overtime_count or penalty_count else "warning")

        ready = (
            total_count > 0
            and 1 <= selected_count <= 4
            and not missing_paths
            and title_ready
            and (not prev_path or os.path.exists(prev_path))
        )
        if hasattr(self, "monthly_next_action_label"):
            if ready:
                self.monthly_next_action_label.configure(
                    text="检查通过。建议先打开预览，确认人员和班次后再导出正式 Word。",
                    text_color=SUCCESS,
                )
            else:
                self.monthly_next_action_label.configure(
                    text="请补齐总名单、至少 1 份周统计 Excel 或修正不可读取的文件路径。",
                    text_color=WARNING,
                )

    def _build_actions(self, page):
        actions = ctk.CTkFrame(
            page,
            fg_color=CARD_BG,
            corner_radius=8,
            border_width=1,
            border_color=BORDER_COLOR,
        )
        actions.grid(row=1, column=0, columnspan=2, sticky="ew", padx=4, pady=(4, 2))
        actions.grid_columnconfigure(1, weight=1)
        ctk.CTkButton(
            actions,
            text="清空选择",
            width=104,
            height=38,
            command=self._clear_monthly_file_inputs,
            **subtle_button_style(),
        ).grid(row=0, column=0, sticky="w", padx=12, pady=12)
        self.month_status = ctk.CTkLabel(
            actions,
            text="等待生成",
            font=FONT_SMALL,
            text_color=TEXT_MUTED,
            anchor="w",
        )
        self.month_status.grid(row=0, column=1, sticky="ew", padx=(0, 12))
        ctk.CTkButton(
            actions,
            text="预览并生成 Word",
            width=172,
            height=38,
            command=self._on_generate_monthly_word,
            **primary_button_style(),
        ).grid(row=0, column=2, sticky="e", padx=(0, 12), pady=12)

    def _clear_monthly_file_inputs(self):
        for slot in self.week_slots:
            slot.set("")
        self.prev_word_path.set("")
        self._refresh_monthly_check_panel()
        set_status(self.month_status, "已清空文件选择", "info")

    def _refresh_adjustment_summary(self):
        if hasattr(self, "month_overtime_summary"):
            self.month_overtime_summary.configure(
                text=summarize_adjustment_records(
                    self.month_overtime_records,
                    "班",
                    "当前未设置加班补录",
                )
            )
        if hasattr(self, "month_penalty_summary"):
            self.month_penalty_summary.configure(
                text=summarize_adjustment_records(
                    self.month_penalty_records,
                    "次",
                    "当前未设置罚班记录",
                )
            )
        self._refresh_monthly_check_panel()

    def _frame_is_alive(self, frame):
        return frame is not None and frame.winfo_exists()

    def _on_pick_month_file(self, idx):
        path = fd.askopenfilename(
            parent=self,
            title=f"选择第{idx + 1}周Excel",
            filetypes=[("Excel文件", "*.xlsx")],
        )
        if path:
            self.week_slots[idx].set(path)

    def _on_pick_prev_word(self):
        path = fd.askopenfilename(
            parent=self,
            title="选择上月月度Word",
            filetypes=[("Word文件", "*.docx")],
        )
        if path:
            self.prev_word_path.set(path)

    def refresh(self):
        """供顶部“刷新”按钮和跨页脏标记统一调用：重读名单并刷新所有面板。"""
        self.refresh_names()
        self._refresh_month_overtime_textbox()
        self._refresh_month_penalty_textbox()
        self._refresh_monthly_check_panel()

    def refresh_names(self):
        names = self.data_mgr.get_name_list()
        self._refresh_adjustment_picker(
            picker=self.month_overtime_picker,
            selected_var=self.month_overtime_name_var,
            names=names,
        )
        self._refresh_adjustment_picker(
            picker=self.month_penalty_picker,
            selected_var=self.month_penalty_name_var,
            names=names,
        )

    def _set_adjustment_name(self, selected_var, selected_names):
        selected_var.set(selected_names[0] if selected_names else "")

    def _refresh_adjustment_picker(self, picker, selected_var, names):
        if not self._frame_is_alive(picker):
            return
        current = selected_var.get().strip()
        if current not in names:
            selected_var.set("")
            current = ""
        picker.set_names(names, selected_names=[current] if current else [])

    def _refresh_month_overtime_textbox(self, scroll_to_bottom=False):
        if not self._frame_is_alive(self.month_overtime_records_frame):
            self._refresh_adjustment_summary()
            return
        clear_children(self.month_overtime_records_frame)
        self.month_overtime_records_frame.grid_columnconfigure(0, weight=1)
        self.month_overtime_records_frame.grid_columnconfigure(1, weight=0)
        if not self.month_overtime_records:
            ctk.CTkLabel(
                self.month_overtime_records_frame,
                text="当前未设置月统计加班补录",
                font=FONT_SMALL,
                text_color=TEXT_MUTED,
            ).grid(row=0, column=0, sticky="w", padx=8, pady=8)
            self._refresh_adjustment_summary()
            return
        for idx, (name, times) in enumerate(self.month_overtime_records.items()):
            add_record_row(
                self.month_overtime_records_frame,
                idx,
                idx + 1,
                name,
                f"+{times} 班",
                lambda n=name: self._remove_month_overtime(n),
            )
        if scroll_to_bottom:
            self._scroll_month_overtime_records_to_bottom()
        self._refresh_adjustment_summary()

    def _scroll_month_overtime_records_to_bottom(self):
        def _scroll():
            canvas = getattr(self.month_overtime_records_frame, "_parent_canvas", None)
            if canvas is not None:
                canvas.yview_moveto(1.0)

        self.after(60, _scroll)

    def _remove_month_overtime(self, name):
        self.month_overtime_records.pop(name, None)
        self._refresh_month_overtime_textbox()

    def _on_add_month_overtime(self):
        names = self.data_mgr.get_name_list()
        if not names:
            mb.showwarning("提示", "请先在【总名单管理】中导入总名单。", parent=self)
            return
        name = self.month_overtime_name_var.get().strip()
        if name not in names:
            mb.showwarning("提示", "请选择有效的加班助理。", parent=self)
            return
        shifts = self.month_overtime_shifts_var.get().strip()
        if not shifts.isdigit() or int(shifts) <= 0:
            mb.showwarning("提示", "加班班次需填写正整数。", parent=self)
            return
        current = self.month_overtime_records.pop(name, 0)
        self.month_overtime_records[name] = current + int(shifts)
        self._refresh_month_overtime_textbox(scroll_to_bottom=True)

    def _on_clear_month_overtime(self):
        self.month_overtime_records = {}
        self._refresh_month_overtime_textbox()

    def _refresh_month_penalty_textbox(self):
        if not self._frame_is_alive(self.month_penalty_records_frame):
            self._refresh_adjustment_summary()
            return
        clear_children(self.month_penalty_records_frame)
        self.month_penalty_records_frame.grid_columnconfigure(0, weight=1)
        if not self.month_penalty_records:
            ctk.CTkLabel(
                self.month_penalty_records_frame,
                text="当前未设置罚班记录",
                font=FONT_SMALL,
                text_color=TEXT_MUTED,
            ).grid(row=0, column=0, sticky="w", padx=8, pady=8)
            self._refresh_adjustment_summary()
            return
        for idx, (name, times) in enumerate(self.month_penalty_records.items()):
            add_record_row(
                self.month_penalty_records_frame,
                idx,
                idx + 1,
                name,
                f"+{times} 罚",
                lambda n=name: self._remove_month_penalty(n),
            )
        self._refresh_adjustment_summary()

    def _remove_month_penalty(self, name):
        self.month_penalty_records.pop(name, None)
        self._refresh_month_penalty_textbox()

    def _on_add_month_penalty(self):
        names = self.data_mgr.get_name_list()
        if not names:
            mb.showwarning("提示", "请先在【总名单管理】中导入总名单。", parent=self)
            return
        name = self.month_penalty_name_var.get().strip()
        if name not in names:
            mb.showwarning("提示", "请选择有效的罚班助理。", parent=self)
            return
        shifts = self.month_penalty_shifts_var.get().strip()
        if not shifts.isdigit() or int(shifts) <= 0:
            mb.showwarning("提示", "罚班次数需填写正整数。", parent=self)
            return
        current = self.month_penalty_records.pop(name, 0)
        self.month_penalty_records[name] = current + int(shifts)
        self._refresh_month_penalty_textbox()

    def _on_clear_month_penalty(self):
        self.month_penalty_records = {}
        self._refresh_month_penalty_textbox()

    def _on_import_month_penalty(self):
        if not self.data_mgr.has_name_list():
            mb.showwarning("提示", "请先在【总名单管理】中导入总名单。", parent=self)
            return
        path = fd.askopenfilename(
            parent=self,
            title="导入罚班记录Word",
            filetypes=[("Word文件", "*.docx")],
        )
        if not path:
            return
        if self.month_penalty_records:
            ok = mb.askyesno(
                "确认导入",
                "导入后将覆盖当前罚班记录，是否继续？",
                parent=self,
            )
            if not ok:
                return
        try:
            records = import_penalty_word(path, self.data_mgr.get_name_list())
            self.month_penalty_records = records
            self._refresh_month_penalty_textbox()
            mb.showinfo("成功", f"已导入 {len(records)} 条罚班记录。", parent=self)
        except (FileNotFoundError, ValueError) as e:
            mb.showerror("导入失败", str(e), parent=self)
        except Exception as e:
            mb.showerror("导入失败", str(e), parent=self)

    def _on_export_month_penalty(self):
        try:
            out = export_penalty_word(self.month_penalty_records, self.data_mgr.get_output_dir())
            mb.showinfo("成功", f"罚班记录Word已导出：\n{out}", parent=self)
        except ValueError as e:
            mb.showwarning("提示", str(e), parent=self)
        except PermissionError as e:
            mb.showerror("文件被占用", str(e), parent=self)
        except Exception as e:
            mb.showerror("导出失败", str(e), parent=self)

    def _on_import_month_overtime(self):
        if not self.data_mgr.has_name_list():
            mb.showwarning("提示", "请先在【总名单管理】中导入总名单。", parent=self)
            return
        path = fd.askopenfilename(
            parent=self,
            title="导入加班补录Word",
            filetypes=[("Word文件", "*.docx")],
        )
        if not path:
            return
        if self.month_overtime_records:
            ok = mb.askyesno(
                "确认导入",
                "导入后将覆盖当前加班补录记录，是否继续？",
                parent=self,
            )
            if not ok:
                return
        try:
            records = import_overtime_word(path, self.data_mgr.get_name_list())
            self.month_overtime_records = records
            self._refresh_month_overtime_textbox(scroll_to_bottom=True)
            mb.showinfo("成功", f"已导入 {len(records)} 条加班补录记录。", parent=self)
        except (FileNotFoundError, ValueError) as e:
            mb.showerror("导入失败", str(e), parent=self)
        except Exception as e:
            mb.showerror("导入失败", str(e), parent=self)

    def _on_export_month_overtime(self):
        try:
            out = export_overtime_word(self.month_overtime_records, self.data_mgr.get_output_dir())
            mb.showinfo("成功", f"加班补录Word已导出：\n{out}", parent=self)
        except ValueError as e:
            mb.showwarning("提示", str(e), parent=self)
        except PermissionError as e:
            mb.showerror("文件被占用", str(e), parent=self)
        except Exception as e:
            mb.showerror("导出失败", str(e), parent=self)

    def _on_generate_monthly_word(self):
        if not self.data_mgr.has_name_list():
            mb.showwarning("提示", "请先在【总名单管理】中导入总名单。", parent=self)
            return
        paths = [v.get().strip() for v in self.week_slots if v.get().strip()]
        if not (1 <= len(paths) <= 4):
            mb.showwarning("提示", "请至少选择 1 份、最多 4 份周统计Excel文件。", parent=self)
            return
        for path in paths:
            if not os.path.exists(path):
                mb.showwarning("提示", f"文件不存在: {path}", parent=self)
                return
        prev_path = self.prev_word_path.get().strip()
        if prev_path and not os.path.exists(prev_path):
            mb.showwarning("提示", f"上月Word文件不存在: {prev_path}", parent=self)
            return

        try:
            set_status(self.month_status, "正在生成预览...", "info")
            self.update_idletasks()
            forced_title = self._build_forced_monthly_title_from_inputs()
            rows = build_monthly_rows(
                weekly_excel_paths=paths,
                total_names=self.data_mgr.get_name_list(),
                prev_word_path=prev_path or None,
                overtime_shifts=self.month_overtime_records,
                penalty_shifts=self.month_penalty_records,
            )
            warnings = collect_monthly_warnings(
                weekly_excel_paths=paths,
                total_names=self.data_mgr.get_name_list(),
                senior_assistants=self.data_mgr.get_senior_assistants(),
            )
            if warnings["messages"]:
                summary = "生成前检测到以下情况：\n\n" + "\n".join(
                    f"· {msg}" for msg in warnings["messages"]
                ) + "\n\n是否仍要继续生成？"
                if not mb.askyesno("生成前提醒", summary, parent=self):
                    set_status(self.month_status, "已取消生成", "info")
                    return
            self._open_monthly_preview(rows, forced_title, warnings["highlight_names"])
            set_status(self.month_status, "请确认预览数据", "info")
        except ValueError as e:
            set_status(self.month_status, "输入或解析失败", "error")
            mb.showerror("输入或解析失败", str(e), parent=self)
        except Exception as e:
            set_status(self.month_status, "预览失败", "error")
            mb.showerror("预览失败", str(e), parent=self)

    def _open_monthly_preview(self, rows, title_text, highlight_names=frozenset()):
        dialog = ctk.CTkToplevel(self)
        dialog.configure(fg_color=CONTENT_BG)
        dialog.title("月统计预览")
        dialog.geometry("980x640")
        dialog.minsize(880, 560)
        dialog.transient(self.winfo_toplevel())
        dialog.grab_set()
        dialog.grid_columnconfigure(0, weight=1)
        dialog.grid_rowconfigure(1, weight=1)

        make_dialog_header(
            dialog,
            "月统计预览",
            "可直接编辑表格数据；确认生成时会按编辑后的多值班次、缺班数量重新排序。",
        )

        table = ctk.CTkScrollableFrame(
            dialog,
            fg_color=CARD_SUBTLE_BG,
            border_width=1,
            border_color=BORDER_COLOR,
        )
        table.grid(row=1, column=0, sticky="nsew", padx=18, pady=(0, 10))
        widths = [52, 128, 98, 98, 98, 98]
        columns = ["序号", "姓名", "总计班次", "应值班次", "多值班次", "缺班数量"]
        for col, width in enumerate(widths):
            table.grid_columnconfigure(col, weight=0, minsize=width)
            ctk.CTkLabel(
                table,
                text=columns[col],
                width=width,
                font=FONT_SMALL,
                text_color=TEXT_MUTED,
            ).grid(row=0, column=col, sticky="ew", padx=3, pady=(6, 4))

        row_vars = []
        for row_idx, row_data in enumerate(rows, start=1):
            display_row = self._make_preview_display_row(row_data)
            is_flagged = str(row_data["姓名"]) in highlight_names
            ctk.CTkLabel(
                table,
                text=str(row_idx),
                width=widths[0],
                font=FONT_BODY,
                text_color=TEXT_MUTED,
            ).grid(row=row_idx, column=0, sticky="ew", padx=3, pady=3)
            vars_for_row = {}
            for col_idx, key in enumerate(columns[1:], start=1):
                var = ctk.StringVar(value=str(display_row[key]))
                vars_for_row[key] = var
                if key == "姓名":
                    ctk.CTkLabel(
                        table,
                        textvariable=var,
                        width=widths[col_idx],
                        height=30,
                        font=FONT_BODY,
                        text_color=WARNING if is_flagged else TEXT_PRIMARY,
                    ).grid(row=row_idx, column=col_idx, sticky="ew", padx=3, pady=3)
                else:
                    entry = ctk.CTkEntry(
                        table,
                        textvariable=var,
                        width=widths[col_idx],
                        height=30,
                        **input_style(),
                    )
                    if key in ("多值班次", "缺班数量"):
                        entry.configure(state="disabled", text_color=TEXT_MUTED)
                    else:
                        bind_focus_border(entry)
                    entry.grid(row=row_idx, column=col_idx, sticky="ew", padx=3, pady=3)
            # 与周统计预览一致：编辑总计/应值时，实时刷新只读的多值/缺班显示。
            for editable_key in ("总计班次", "应值班次"):
                vars_for_row[editable_key].trace_add(
                    "write",
                    lambda *_a, item=vars_for_row: self._sync_monthly_derived_vars(item),
                )
            row_vars.append(vars_for_row)

        actions = ctk.CTkFrame(dialog, fg_color="transparent")
        actions.grid(row=2, column=0, sticky="ew", padx=18, pady=(0, 14))
        actions.grid_columnconfigure(0, weight=1)
        ctk.CTkButton(
            actions,
            text="取消",
            width=88,
            command=dialog.destroy,
            **muted_button_style(),
        ).grid(row=0, column=1, padx=(0, 8))
        ctk.CTkButton(
            actions,
            text="生成Word",
            width=116,
            command=lambda: self._on_confirm_monthly_preview(dialog, row_vars, title_text),
            **primary_button_style(),
        ).grid(row=0, column=2)

        dialog.bind("<Escape>", lambda _event: dialog.destroy())

    def _make_preview_display_row(self, row_data):
        return {
            "姓名": row_data["姓名"],
            "总计班次": int(row_data["总计班次"]),
            "应值班次": int(row_data["应值班次"]),
            "多值班次": max(0, int(row_data["多值班次"])),
            "缺班数量": max(0, int(row_data["缺班数量"])),
        }

    def _sync_monthly_derived_vars(self, vars_for_row):
        """编辑总计/应值后，按口径实时回填只读的多值班次、缺班数量显示。"""
        total_text = vars_for_row["总计班次"].get().strip()
        should_text = vars_for_row["应值班次"].get().strip()
        if total_text.isdigit() and should_text.isdigit():
            total = int(total_text)
            should = int(should_text)
            vars_for_row["多值班次"].set(str(max(0, total - should)))
            vars_for_row["缺班数量"].set(str(max(0, should - total)))

    def _on_confirm_monthly_preview(self, dialog, row_vars, title_text):
        try:
            rows = self._collect_preview_rows(row_vars)
        except ValueError as e:
            mb.showwarning("提示", str(e), parent=dialog)
            return

        try:
            output_path = build_unique_output_path(self.data_mgr.get_output_dir(), title_text, ".docx")
            set_status(self.month_status, "正在生成...", "info")
            self.update_idletasks()
            out = generate_monthly_word_from_rows(
                rows=rows,
                output_path=str(output_path),
                title_text=title_text,
            )
            dialog.destroy()
            set_status(self.month_status, f"已生成：{Path(out).name}", "success")
            mb.showinfo("成功", f"月统计Word已生成：\n{out}", parent=self)
        except PermissionError as e:
            set_status(self.month_status, "文件被占用", "error")
            mb.showerror("文件被占用", str(e), parent=dialog)
        except ValueError as e:
            set_status(self.month_status, "输入或解析失败", "error")
            mb.showerror("输入或解析失败", str(e), parent=dialog)
        except Exception as e:
            set_status(self.month_status, "生成失败", "error")
            mb.showerror("生成失败", str(e), parent=dialog)

    def _collect_preview_rows(self, row_vars):
        rows = []
        seen_names = set()
        numeric_fields = ("总计班次", "应值班次")
        for idx, vars_for_row in enumerate(row_vars, start=1):
            name = vars_for_row["姓名"].get().strip()
            if not name:
                raise ValueError(f"第 {idx} 行姓名不能为空。")
            if name in seen_names:
                raise ValueError(f"姓名重复：{name}")
            seen_names.add(name)

            row = {"姓名": name}
            for field in numeric_fields:
                text = vars_for_row[field].get().strip()
                if not text.isdigit():
                    raise ValueError(f"第 {idx} 行【{field}】必须为非负整数。")
                row[field] = int(text)
            row["多值班次"] = row["总计班次"] - row["应值班次"]
            row["缺班数量"] = row["应值班次"] - row["总计班次"]
            rows.append(row)
        return rows

    def _build_forced_monthly_title_from_inputs(self) -> str:
        return build_monthly_report_title(
            self.title_year_var.get(),
            self.title_season_var.get(),
            self.title_week_start_var.get(),
            self.title_week_end_var.get(),
        )
