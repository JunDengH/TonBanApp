"""Independent two-slot final-weeks statistics workspace."""
from __future__ import annotations

import datetime
import os
import tkinter.filedialog as fd
import tkinter.messagebox as mb
from pathlib import Path

import customtkinter as ctk

from src.modules.final_weeks_parser import FinalWeekBlock, parse_final_weeks_pdf
from src.modules.final_weeks_service import (
    CombinedFinalPeriodResult,
    FinalWeekResult,
    build_combined_final_period_rows,
    build_final_weeks_rows,
    write_combined_final_period_excel,
    write_final_weeks_excels,
)
from src.modules.word_parser import (
    find_typo_suspects,
    parse_weekly_schedule,
    scan_weekly_word,
)
from src.ui.ui_helpers import (
    BORDER_COLOR,
    CARD_BG,
    CARD_SUBTLE_BG,
    CONTENT_BG,
    ERROR,
    FONT_BODY,
    FONT_BODY_BOLD,
    FONT_SECTION,
    FONT_SMALL,
    PANEL_BG,
    SUCCESS,
    TEXT_MUTED,
    TEXT_PRIMARY,
    WARNING,
    add_file_row,
    bind_focus_border,
    checkbox_style,
    input_style,
    make_bottom_action_bar,
    make_check_group,
    make_check_panel,
    make_dialog_header,
    make_tab_page,
    muted_button_style,
    option_menu_style,
    primary_button_style,
    set_check_item,
    set_status,
    subtle_button_style,
)
from src.ui.weekly_tab import WeeklyTab


SUPPORTED_DAYS = ("周一", "周二", "周三", "周四", "周五")
FINAL_WEEKS_LAYOUT_SECTIONS = ["期末排班 PDF", "期末周1", "期末周2", "预览并生成"]
FINAL_WEEKS_SPLIT_EXPORT_ENABLED = False


def _validate_week_numbers(first: str, second: str) -> tuple[int, int]:
    values = []
    for slot, raw in enumerate((first, second), start=1):
        text = str(raw or "").strip()
        if not text.isdigit() or int(text) <= 0:
            raise ValueError(f"期末周{slot}的周次必须是正整数。")
        values.append(int(text))
    if values[0] == values[1]:
        raise ValueError("期末周1和期末周2的实际周次不能相同。")
    return values[0], values[1]


class FinalWeeksTab(ctk.CTkFrame):
    """Generate one combined final-period workbook, retaining split export."""

    def __init__(self, parent, data_mgr):
        super().__init__(parent, fg_color="transparent")
        self.data_mgr = data_mgr
        self.parsed_blocks: list[FinalWeekBlock] = []

        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)
        self.grid(row=0, column=0, sticky="nsew")
        self._build()

    def _build(self):
        subtitle = (
            "一份双周期末排班 PDF 对应两份实际值班 Word，确认后分别生成两份周统计 Excel。"
            if FINAL_WEEKS_SPLIT_EXPORT_ENABLED
            else "一份双周期末排班 PDF 对应两份实际值班 Word，两周期末周合并为一个统计周期。"
        )
        page = make_tab_page(
            self,
            "期末周统计",
            subtitle=subtitle,
        )
        page.grid_rowconfigure(0, weight=1)
        page.grid_columnconfigure(0, weight=1)
        page.grid_columnconfigure(1, weight=0, minsize=330)

        now = datetime.datetime.now()
        self.year_var = ctk.StringVar(value=str(now.year))
        self.season_var = ctk.StringVar(value="春季" if 2 <= now.month <= 7 else "秋季")
        self.pdf_path = ctk.StringVar(value="")
        self.week_number_vars = {1: ctk.StringVar(value=""), 2: ctk.StringVar(value="")}
        self.actual_word_paths = {1: ctk.StringVar(value=""), 2: ctk.StringVar(value="")}
        self.holiday_vars = (
            {
                slot: {day: ctk.BooleanVar(value=False) for day in SUPPORTED_DAYS}
                for slot in (1, 2)
            }
            if FINAL_WEEKS_SPLIT_EXPORT_ENABLED
            else {}
        )
        for variable in [
            self.year_var,
            self.season_var,
            self.pdf_path,
            *self.week_number_vars.values(),
            *self.actual_word_paths.values(),
        ]:
            variable.trace_add("write", lambda *_args: self._refresh_checks())

        workarea = ctk.CTkFrame(
            page,
            fg_color=CARD_BG,
            corner_radius=8,
            border_width=1,
            border_color=BORDER_COLOR,
        )
        workarea.grid(row=0, column=0, sticky="nsew", padx=(4, 10), pady=4)
        workarea.grid_columnconfigure(0, weight=1)
        workarea.grid_rowconfigure(3, weight=1)
        self._build_workarea(workarea)
        self._build_check_panel(page)

        self.status_label = make_bottom_action_bar(
            page,
            row=1,
            primary_text=(
                "预览两周统计"
                if FINAL_WEEKS_SPLIT_EXPORT_ENABLED
                else "预览期末周统计"
            ),
            primary_command=self._on_preview,
            status_text="等待选择期末排班 PDF",
            secondary_buttons=[{"text": "清空选择", "width": 104, "command": self._clear_inputs}],
        )
        self._refresh_checks()

    def _build_workarea(self, parent):
        ctk.CTkLabel(
            parent,
            text="期末周统计工作台",
            font=FONT_SECTION,
            text_color=TEXT_PRIMARY,
            anchor="w",
        ).grid(row=0, column=0, sticky="ew", padx=16, pady=(14, 2))
        ctk.CTkLabel(
            parent,
            text="系统按 PDF 版面顺序识别期末周1和期末周2；实际周数自动填入，也可在生成前修正。",
            font=FONT_SMALL,
            text_color=TEXT_MUTED,
            anchor="w",
            justify="left",
            wraplength=850,
        ).grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 10))

        title_row = ctk.CTkFrame(parent, fg_color=CARD_SUBTLE_BG, corner_radius=8, border_width=1, border_color=BORDER_COLOR)
        title_row.grid(row=2, column=0, sticky="ew", padx=16, pady=(0, 10))
        title_row.grid_columnconfigure(4, weight=1)
        ctk.CTkLabel(title_row, text="年份", font=FONT_BODY, text_color=TEXT_PRIMARY).grid(
            row=0, column=0, padx=(14, 8), pady=12
        )
        year_entry = ctk.CTkEntry(title_row, textvariable=self.year_var, width=96, height=34, **input_style())
        bind_focus_border(year_entry)
        year_entry.grid(row=0, column=1, padx=(0, 18), pady=12)
        ctk.CTkLabel(title_row, text="学期", font=FONT_BODY, text_color=TEXT_PRIMARY).grid(
            row=0, column=2, padx=(0, 8), pady=12
        )
        ctk.CTkOptionMenu(
            title_row,
            values=["春季", "秋季"],
            variable=self.season_var,
            width=110,
            **option_menu_style(),
        ).grid(row=0, column=3, padx=(0, 18), pady=12)
        ctk.CTkLabel(
            title_row,
            text=(
                "输出文件名使用下方两个实际周次"
                if FINAL_WEEKS_SPLIT_EXPORT_ENABLED
                else "合并文件名不含具体周数"
            ),
            font=FONT_SMALL,
            text_color=TEXT_MUTED,
            anchor="e",
        ).grid(row=0, column=4, sticky="e", padx=(0, 14), pady=12)

        scroll = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        scroll.grid(row=3, column=0, sticky="nsew", padx=10, pady=(0, 10))
        scroll.grid_columnconfigure(0, weight=1)

        pdf_panel = ctk.CTkFrame(scroll, fg_color=CARD_SUBTLE_BG, corner_radius=8, border_width=1, border_color=BORDER_COLOR)
        pdf_panel.grid(row=0, column=0, sticky="ew", padx=6, pady=(0, 10))
        pdf_panel.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(pdf_panel, text="期末排班 PDF", font=FONT_BODY_BOLD, text_color=TEXT_PRIMARY, anchor="w").grid(
            row=0, column=0, sticky="ew", padx=16, pady=(12, 2)
        )
        ctk.CTkLabel(
            pdf_panel,
            text="PDF 必须包含两个有周次标题的排班区块，周数不要求连续。",
            font=FONT_SMALL,
            text_color=TEXT_MUTED,
            anchor="w",
        ).grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 6))
        add_file_row(pdf_panel, 2, "排班 PDF", self.pdf_path, "选择包含两个期末周的 PDF", self._pick_pdf)

        for slot in (1, 2):
            self._build_slot_panel(scroll, row=slot, slot=slot)

    def _build_slot_panel(self, parent, row, slot):
        panel = ctk.CTkFrame(parent, fg_color=CARD_SUBTLE_BG, corner_radius=8, border_width=1, border_color=BORDER_COLOR)
        panel.grid(row=row, column=0, sticky="ew", padx=6, pady=(0, 10))
        panel.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            panel,
            text=f"期末周{slot}",
            font=FONT_BODY_BOLD,
            text_color=TEXT_PRIMARY,
            anchor="w",
        ).grid(row=0, column=0, sticky="ew", padx=16, pady=(12, 2))
        ctk.CTkLabel(
            panel,
            text=f"选择与 PDF 中第 {slot} 个周区块对应的实际值班 Word。",
            font=FONT_SMALL,
            text_color=TEXT_MUTED,
            anchor="w",
        ).grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 6))

        week_row = ctk.CTkFrame(panel, fg_color="transparent")
        week_row.grid(row=2, column=0, sticky="ew", padx=16, pady=4)
        week_row.grid_columnconfigure(2, weight=1)
        ctk.CTkLabel(week_row, text="实际周次", width=88, font=FONT_BODY, text_color=TEXT_PRIMARY, anchor="w").grid(
            row=0, column=0, sticky="w", padx=(0, 10)
        )
        entry = ctk.CTkEntry(
            week_row,
            textvariable=self.week_number_vars[slot],
            width=96,
            height=34,
            placeholder_text="自动识别",
            **input_style(),
        )
        bind_focus_border(entry)
        entry.grid(row=0, column=1, sticky="w")
        ctk.CTkLabel(
            week_row,
            text=(
                "仅用于标题和文件名，不改变两个 Word 的对应顺序"
                if FINAL_WEEKS_SPLIT_EXPORT_ENABLED
                else "仅用于核对两个 Word 的对应顺序，合并文件名不含具体周数"
            ),
            font=FONT_SMALL,
            text_color=TEXT_MUTED,
            anchor="w",
        ).grid(row=0, column=2, sticky="w", padx=(12, 0))

        add_file_row(
            panel,
            3,
            "实际值班 Word",
            self.actual_word_paths[slot],
            f"选择期末周{slot}实际值班 Word",
            lambda target=slot: self._pick_actual_word(target),
        )

        if FINAL_WEEKS_SPLIT_EXPORT_ENABLED:
            holiday = ctk.CTkFrame(
                panel,
                fg_color=PANEL_BG,
                corner_radius=8,
                border_width=1,
                border_color=BORDER_COLOR,
            )
            holiday.grid(row=4, column=0, sticky="ew", padx=16, pady=(8, 12))
            holiday.grid_columnconfigure(
                (0, 1, 2, 3, 4),
                weight=1,
                uniform=f"final_week_{slot}_days",
            )
            for index, day in enumerate(SUPPORTED_DAYS):
                ctk.CTkCheckBox(
                    holiday,
                    text=day,
                    variable=self.holiday_vars[slot][day],
                    command=self._refresh_checks,
                    font=FONT_BODY,
                    **checkbox_style(),
                ).grid(row=0, column=index, sticky="w", padx=10, pady=10)

    def _build_check_panel(self, page):
        panel = make_check_panel(page, title="生成前检查", scrollable=True)
        panel.grid(row=0, column=1, sticky="nsew", padx=(0, 4), pady=4)
        self.check_items = {}
        make_check_group(
            panel,
            row=1,
            title="输入完整性",
            items=[
                ("names", "总名单状态"),
                ("pdf", "期末排班 PDF"),
                ("split", "双周区块识别"),
                ("slot1", "期末周1实际值班 Word"),
                ("slot2", "期末周2实际值班 Word"),
                ("week_numbers", "实际周次"),
                ("typo", "姓名拼写检查"),
            ],
            registry=self.check_items,
        )
        make_check_group(
            panel,
            row=2,
            title="规则与输出",
            items=[
                ("senior", "毕业季规则"),
                ("leave", "长期请假规则"),
                ("multiplier", "期末周双倍计数"),
                ("output", "输出文件"),
            ],
            registry=self.check_items,
        )
        self.next_action = ctk.CTkLabel(
            panel,
            text="",
            font=FONT_SMALL,
            text_color=TEXT_MUTED,
            anchor="w",
            justify="left",
            wraplength=280,
        )
        self.next_action.grid(row=3, column=0, sticky="ew", padx=16, pady=(0, 16))

    def _pick_pdf(self):
        path = fd.askopenfilename(
            parent=self,
            title="选择期末排班 PDF",
            filetypes=[("PDF 排班表", "*.pdf")],
        )
        if not path:
            return
        self.pdf_path.set(path)
        self._load_pdf_metadata(show_error=True)

    def _pick_actual_word(self, slot):
        path = fd.askopenfilename(
            parent=self,
            title=f"选择期末周{slot}实际值班 Word",
            filetypes=[("Word 文件", "*.docx")],
        )
        if path:
            self.actual_word_paths[slot].set(path)

    def _load_pdf_metadata(self, show_error=False):
        self.parsed_blocks = []
        path = self.pdf_path.get().strip()
        if not path or not os.path.exists(path) or not self.data_mgr.has_name_list():
            self._refresh_checks()
            return False
        try:
            self.parsed_blocks = parse_final_weeks_pdf(path, self.data_mgr.get_name_list())
            for block in self.parsed_blocks:
                self.week_number_vars[block.slot].set(str(block.week_number))
            set_status(self.status_label, "已识别两个期末周，请继续选择两份实际 Word", "success")
            self._refresh_checks()
            return True
        except Exception as exc:
            self._refresh_checks()
            set_status(self.status_label, "PDF 双周识别失败", "error")
            if show_error:
                mb.showerror("PDF 解析失败", str(exc), parent=self)
            return False

    def _refresh_checks(self):
        if not hasattr(self, "check_items"):
            return
        names = self.data_mgr.get_name_list()
        set_check_item(
            self.check_items,
            "names",
            f"总名单 {len(names)} 人" if names else "尚未导入总名单",
            "success" if names else "error",
        )
        pdf_path = self.pdf_path.get().strip()
        pdf_ready = bool(pdf_path) and os.path.isfile(pdf_path) and Path(pdf_path).suffix.lower() == ".pdf"
        set_check_item(self.check_items, "pdf", Path(pdf_path).name if pdf_ready else "请选择可读取的 PDF", "success" if pdf_ready else "warning")
        if len(self.parsed_blocks) == 2:
            recognized = "、".join(f"期末周{b.slot}=第{b.week_number}周" for b in self.parsed_blocks)
            set_check_item(self.check_items, "split", recognized, "success")
        else:
            set_check_item(self.check_items, "split", "等待识别两个周次区块", "warning")

        words_ready = True
        for slot in (1, 2):
            path = self.actual_word_paths[slot].get().strip()
            ready = bool(path) and os.path.isfile(path) and Path(path).suffix.lower() == ".docx"
            words_ready = words_ready and ready
            set_check_item(
                self.check_items,
                f"slot{slot}",
                Path(path).name if ready else f"请选择期末周{slot}实际值班 Word",
                "success" if ready else "warning",
            )
        try:
            first, second = _validate_week_numbers(
                self.week_number_vars[1].get(), self.week_number_vars[2].get()
            )
            week_ready = True
            set_check_item(self.check_items, "week_numbers", f"第{first}周、第{second}周", "success")
        except ValueError as exc:
            week_ready = False
            set_check_item(self.check_items, "week_numbers", str(exc), "warning")

        suspect_count = 0
        if len(self.parsed_blocks) == 2:
            for block in self.parsed_blocks:
                suspect_count += len(find_typo_suspects(block.unknown_names, names))
        set_check_item(
            self.check_items,
            "typo",
            f"排班 PDF 中发现 {suspect_count} 个疑似姓名，预览前确认" if suspect_count else "排班 PDF 暂未发现疑似姓名",
            "warning" if suspect_count else "success",
        )

        senior_mode = self.data_mgr.get_senior_should_mode()
        seniors = self.data_mgr.get_senior_assistants()
        conflict = senior_mode == "reduced" and bool(seniors)
        set_check_item(
            self.check_items,
            "senior",
            "毕业季少值班与期末周双倍计数冲突" if conflict else f"当前模式：{senior_mode}，已保存 {len(seniors)} 人",
            "error" if conflict else "success",
        )
        leaves = self.data_mgr.get_long_term_leave_assistants()
        set_check_item(self.check_items, "leave", f"长期请假 {len(leaves)} 人；重叠时毕业季优先", "success")
        set_check_item(
            self.check_items,
            "multiplier",
            (
                "两个期末周分别计算，每次出现固定计为两次"
                if FINAL_WEEKS_SPLIT_EXPORT_ENABLED
                else "两个期末周合并计算；每次实际出现计 2 班，不做放假核减"
            ),
            "success",
        )
        if week_ready:
            output_text = (
                "将分别生成两份以实际周数命名的 Excel"
                if FINAL_WEEKS_SPLIT_EXPORT_ENABLED
                else f"将生成 {self.year_var.get().strip()}{self.season_var.get()}学期期末周助理值班统计.xlsx"
            )
            set_check_item(self.check_items, "output", output_text, "info")
        else:
            set_check_item(self.check_items, "output", "修正周次后显示输出文件", "warning")

        ready = bool(names) and pdf_ready and len(self.parsed_blocks) == 2 and words_ready and week_ready and not conflict
        self.next_action.configure(
            text=(
                (
                    "检查通过，可以预览两周统计。"
                    if FINAL_WEEKS_SPLIT_EXPORT_ENABLED
                    else "检查通过，可以预览合并期末周统计。"
                )
                if ready
                else "请补齐 PDF、两个周次和两份实际 Word，并处理规则冲突。"
            ),
            text_color=SUCCESS if ready else TEXT_MUTED,
        )

    def _clear_inputs(self):
        self.pdf_path.set("")
        self.parsed_blocks = []
        for slot in (1, 2):
            self.week_number_vars[slot].set("")
            self.actual_word_paths[slot].set("")
            for variable in self.holiday_vars.get(slot, {}).values():
                variable.set(False)
        set_status(self.status_label, "已清空选择", "info")
        self._refresh_checks()

    def _validate_inputs(self) -> tuple[int, int]:
        if not self.data_mgr.has_name_list():
            raise ValueError("请先在总名单页面导入总名单。")
        pdf_path = self.pdf_path.get().strip()
        if not pdf_path or not os.path.isfile(pdf_path) or Path(pdf_path).suffix.lower() != ".pdf":
            raise ValueError("请选择有效的期末排班 PDF。")
        if len(self.parsed_blocks) != 2 and not self._load_pdf_metadata(show_error=False):
            raise ValueError("排班 PDF 尚未成功识别出两个期末周区块。")
        for slot in (1, 2):
            path = self.actual_word_paths[slot].get().strip()
            if not path or not os.path.isfile(path) or Path(path).suffix.lower() != ".docx":
                raise ValueError(f"请选择有效的期末周{slot}实际值班 Word。")
        if self.data_mgr.get_senior_should_mode() == "reduced" and self.data_mgr.get_senior_assistants():
            raise ValueError("期末周双倍计数与毕业季『少值班』模式冲突，请改为正常值班或无需值班。")
        return _validate_week_numbers(self.week_number_vars[1].get(), self.week_number_vars[2].get())

    def _collect_typo_suspects(self, blocks):
        total_names = self.data_mgr.get_name_list()
        suspects = []
        for block in blocks:
            for suspect in find_typo_suspects(block.unknown_names, total_names):
                suspects.append({**suspect, "source": f"期末周{block.slot}排班 PDF"})
        for slot in (1, 2):
            scan = scan_weekly_word(self.actual_word_paths[slot].get(), total_names)
            for suspect in find_typo_suspects(scan["unknown_names"], total_names):
                suspects.append({**suspect, "source": f"期末周{slot}实际 Word"})
        return suspects

    def _on_preview(self):
        if FINAL_WEEKS_SPLIT_EXPORT_ENABLED:
            return self._on_preview_split()
        return self._on_preview_combined()

    def _on_preview_combined(self):
        try:
            week_numbers = self._validate_inputs()
            total_names = self.data_mgr.get_name_list()
            set_status(self.status_label, "正在检查姓名拼写...", "info")
            self.update_idletasks()
            typo_suspects = self._collect_typo_suspects(self.parsed_blocks)
            corrections = {}
            if typo_suspects:
                corrections = WeeklyTab._show_typo_correction_dialog(
                    self,
                    typo_suspects,
                )
                if corrections is None:
                    set_status(self.status_label, "已取消预览", "info")
                    return

            blocks = parse_final_weeks_pdf(
                self.pdf_path.get(),
                total_names,
                corrections,
            )
            actual_schedules = {}
            warning_messages = []
            for slot in (1, 2):
                actual_path = self.actual_word_paths[slot].get()
                actual_schedules[slot] = parse_weekly_schedule(
                    actual_path,
                    total_names,
                    corrections,
                )
                scan = scan_weekly_word(actual_path, total_names, corrections)
                if scan["matched_count"] == 0:
                    warning_messages.append(
                        f"期末周{slot}实际 Word 未识别到任何总名单内姓名。"
                    )
                if scan["unknown_names"]:
                    warning_messages.append(
                        f"期末周{slot}实际 Word 中存在未识别内容："
                        + "、".join(scan["unknown_names"][:10])
                    )

            result = build_combined_final_period_rows(
                blocks=blocks,
                actual_schedules=actual_schedules,
                total_names=total_names,
                senior_assistants=self.data_mgr.get_senior_assistants(),
                senior_should_mode=self.data_mgr.get_senior_should_mode(),
                long_term_leave_assistants=(
                    self.data_mgr.get_long_term_leave_assistants()
                ),
            )
            result.week_numbers = week_numbers
            warning_messages.extend(result.warnings.get("messages", []))

            if warning_messages:
                summary = "生成前检测到以下情况：\n\n" + "\n".join(
                    f"· {message}" for message in warning_messages
                )
                if not mb.askyesno(
                    "生成前提醒",
                    summary + "\n\n是否继续预览？",
                    parent=self,
                ):
                    set_status(self.status_label, "已取消预览", "info")
                    return
            self._open_combined_preview(result)
            set_status(self.status_label, "请确认合并统计数据", "info")
        except PermissionError as exc:
            set_status(self.status_label, "文件被占用", "error")
            mb.showerror("文件被占用", str(exc), parent=self)
        except Exception as exc:
            set_status(self.status_label, "预览失败", "error")
            mb.showerror("期末周统计失败", str(exc), parent=self)

    def _on_preview_split(self):
        try:
            week_numbers = self._validate_inputs()
            total_names = self.data_mgr.get_name_list()
            set_status(self.status_label, "正在检查姓名拼写...", "info")
            self.update_idletasks()
            typo_suspects = self._collect_typo_suspects(self.parsed_blocks)
            corrections = {}
            if typo_suspects:
                corrections = WeeklyTab._show_typo_correction_dialog(self, typo_suspects)
                if corrections is None:
                    set_status(self.status_label, "已取消预览", "info")
                    return

            blocks = parse_final_weeks_pdf(self.pdf_path.get(), total_names, corrections)
            actual_schedules = {}
            warning_messages = []
            for slot in (1, 2):
                actual_path = self.actual_word_paths[slot].get()
                actual_schedules[slot] = parse_weekly_schedule(actual_path, total_names, corrections)
                scan = scan_weekly_word(actual_path, total_names, corrections)
                if scan["matched_count"] == 0:
                    warning_messages.append(f"期末周{slot}实际 Word 未识别到任何总名单内姓名。")
                if scan["unknown_names"]:
                    warning_messages.append(
                        f"期末周{slot}实际 Word 中存在未识别内容：" + "、".join(scan["unknown_names"][:10])
                    )

            results = build_final_weeks_rows(
                blocks=blocks,
                actual_schedules=actual_schedules,
                total_names=total_names,
                holidays_by_slot={
                    slot: [day for day, var in self.holiday_vars[slot].items() if var.get()]
                    for slot in (1, 2)
                },
                senior_assistants=self.data_mgr.get_senior_assistants(),
                senior_should_mode=self.data_mgr.get_senior_should_mode(),
                long_term_leave_assistants=self.data_mgr.get_long_term_leave_assistants(),
            )
            for result, week_number in zip(results, week_numbers):
                result.week_number = week_number
                warning_messages.extend(result.warnings.get("messages", []))

            if warning_messages:
                summary = "生成前检测到以下情况：\n\n" + "\n".join(f"· {message}" for message in warning_messages)
                if not mb.askyesno("生成前提醒", summary + "\n\n是否继续预览？", parent=self):
                    set_status(self.status_label, "已取消预览", "info")
                    return
            self._open_split_preview(results)
            set_status(self.status_label, "请在两个页签中确认数据", "info")
        except PermissionError as exc:
            set_status(self.status_label, "文件被占用", "error")
            mb.showerror("文件被占用", str(exc), parent=self)
        except Exception as exc:
            set_status(self.status_label, "预览失败", "error")
            mb.showerror("期末周统计失败", str(exc), parent=self)

    def _open_combined_preview(self, result: CombinedFinalPeriodResult):
        dialog = ctk.CTkToplevel(self)
        dialog.configure(fg_color=CONTENT_BG)
        dialog.title("期末周统计预览")
        dialog.geometry("930x680")
        dialog.minsize(820, 560)
        dialog.transient(self.winfo_toplevel())
        dialog.grab_set()
        dialog.grid_columnconfigure(0, weight=1)
        dialog.grid_rowconfigure(1, weight=1)
        make_dialog_header(
            dialog,
            "期末周统计预览",
            (
                f"第{result.week_numbers[0]}周与第{result.week_numbers[1]}周合并统计；"
                "应值和实际可修正，缺班会实时重算。"
            ),
        )
        preview = ctk.CTkFrame(dialog, fg_color=CARD_BG)
        preview.grid(row=1, column=0, sticky="nsew", padx=18, pady=(0, 10))
        preview.grid_columnconfigure(0, weight=1)
        preview.grid_rowconfigure(0, weight=1)
        ctk.CTkLabel(
            preview,
            text="正在载入预览数据...",
            font=FONT_BODY,
            text_color=TEXT_MUTED,
        ).grid(row=0, column=0, pady=30)
        dialog.update_idletasks()
        for child in preview.winfo_children():
            child.destroy()
        row_vars = self._build_preview_table(preview, result)

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
            text="生成期末周 Excel",
            width=160,
            command=lambda: self._confirm_combined_preview(
                dialog,
                result,
                row_vars,
            ),
            **primary_button_style(),
        ).grid(row=0, column=2)
        dialog.bind("<Escape>", lambda _event: dialog.destroy())

    def _open_split_preview(self, results: list[FinalWeekResult]):
        dialog = ctk.CTkToplevel(self)
        dialog.configure(fg_color=CONTENT_BG)
        dialog.title("期末周统计预览")
        dialog.geometry("930x680")
        dialog.minsize(820, 560)
        dialog.transient(self.winfo_toplevel())
        dialog.grab_set()
        dialog.grid_columnconfigure(0, weight=1)
        dialog.grid_rowconfigure(1, weight=1)
        make_dialog_header(
            dialog,
            "期末周统计预览",
            "分别确认两个页签；应值和实际可修正，缺班会实时重算。",
        )
        tabs = ctk.CTkTabview(
            dialog,
            fg_color=CARD_BG,
            segmented_button_fg_color=CARD_SUBTLE_BG,
            segmented_button_selected_color=PANEL_BG,
            segmented_button_selected_hover_color=PANEL_BG,
            segmented_button_unselected_color=CARD_SUBTLE_BG,
            segmented_button_unselected_hover_color=PANEL_BG,
            text_color=TEXT_PRIMARY,
        )
        tabs.grid(row=1, column=0, sticky="nsew", padx=18, pady=(0, 10))
        row_vars_by_slot = {}
        preview_tabs = {}
        for result in results:
            tab_name = f"期末周{result.slot}（第{result.week_number}周）"
            tab = tabs.add(tab_name)
            tab.grid_columnconfigure(0, weight=1)
            tab.grid_rowconfigure(0, weight=1)
            preview_tabs[result.slot] = tab
            ctk.CTkLabel(
                tab,
                text="正在载入预览数据...",
                font=FONT_BODY,
                text_color=TEXT_MUTED,
            ).grid(row=0, column=0, pady=30)
        dialog.update_idletasks()
        for result in results:
            tab = preview_tabs[result.slot]
            for child in tab.winfo_children():
                child.destroy()
            row_vars_by_slot[result.slot] = self._build_preview_table(tab, result)

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
            text="生成两周 Excel",
            width=146,
            command=lambda: self._confirm_split_preview(dialog, results, row_vars_by_slot),
            **primary_button_style(),
        ).grid(row=0, column=2)
        dialog.bind("<Escape>", lambda _event: dialog.destroy())

    def _build_preview_table(self, parent, result):
        table = ctk.CTkScrollableFrame(parent, fg_color=CARD_SUBTLE_BG, border_width=1, border_color=BORDER_COLOR)
        table.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        columns = ["序号", "姓名", "应值班次", "实际班次", "缺班", "备注"]
        widths = [48, 126, 92, 92, 76, 240]
        for column, width in enumerate(widths):
            table.grid_columnconfigure(column, weight=1 if column == 5 else 0, minsize=width)
            ctk.CTkLabel(
                table,
                text=columns[column],
                width=width,
                font=FONT_SMALL,
                text_color=TEXT_MUTED,
            ).grid(row=0, column=column, sticky="ew", padx=3, pady=(6, 4))
        row_vars = []
        highlights = result.warnings.get("highlight_names", set())
        for row_index, row_data in enumerate(result.rows, start=1):
            ctk.CTkLabel(table, text=str(row_index), width=widths[0], font=FONT_BODY, text_color=TEXT_MUTED).grid(
                row=row_index, column=0, sticky="ew", padx=3, pady=3
            )
            variables = {
                "姓名": ctk.StringVar(value=str(row_data["姓名"])),
                "应值班次": ctk.StringVar(value=str(row_data["应值班次"])),
                "实际班次": ctk.StringVar(value=str(row_data["实际班次"])),
                "缺班": ctk.StringVar(value=str(row_data["缺班"])),
                "备注": ctk.StringVar(value=str(row_data.get("备注", ""))),
            }
            name_color = WARNING if row_data["姓名"] in highlights else TEXT_PRIMARY
            ctk.CTkLabel(
                table,
                textvariable=variables["姓名"],
                width=widths[1],
                font=FONT_BODY,
                text_color=name_color,
            ).grid(row=row_index, column=1, sticky="ew", padx=3, pady=3)
            for column, key in enumerate(("应值班次", "实际班次"), start=2):
                entry = ctk.CTkEntry(table, textvariable=variables[key], width=widths[column], height=30, **input_style())
                bind_focus_border(entry)
                entry.grid(row=row_index, column=column, sticky="ew", padx=3, pady=3)
                variables[key].trace_add("write", lambda *_args, item=variables: self._sync_absence(item))
            absence = ctk.CTkEntry(table, textvariable=variables["缺班"], width=widths[4], height=30, **input_style())
            absence.configure(state="disabled", text_color=TEXT_MUTED)
            absence.grid(row=row_index, column=4, sticky="ew", padx=3, pady=3)
            remark = ctk.CTkEntry(table, textvariable=variables["备注"], width=widths[5], height=30, **input_style())
            bind_focus_border(remark)
            remark.grid(row=row_index, column=5, sticky="ew", padx=3, pady=3)
            row_vars.append(variables)
        return row_vars

    @staticmethod
    def _sync_absence(variables):
        should = variables["应值班次"].get().strip()
        actual = variables["实际班次"].get().strip()
        if should.isdigit() and actual.isdigit():
            variables["缺班"].set(str(max(0, int(should) - int(actual))))

    @staticmethod
    def _collect_preview_rows(row_vars):
        rows = []
        for index, variables in enumerate(row_vars, start=1):
            should = variables["应值班次"].get().strip()
            actual = variables["实际班次"].get().strip()
            if not should.isdigit():
                raise ValueError(f"第 {index} 行【应值班次】必须为非负整数。")
            if not actual.isdigit():
                raise ValueError(f"第 {index} 行【实际班次】必须为非负整数。")
            rows.append({
                "姓名": variables["姓名"].get().strip(),
                "应值班次": int(should),
                "实际班次": int(actual),
                "缺班": max(0, int(should) - int(actual)),
                "备注": variables["备注"].get().strip(),
            })
        return rows

    def _confirm_combined_preview(self, dialog, result, row_vars):
        try:
            result.rows = self._collect_preview_rows(row_vars)
            set_status(self.status_label, "正在生成期末周 Excel...", "info")
            self.update_idletasks()
            output = write_combined_final_period_excel(
                result,
                self.data_mgr.get_output_dir(),
                self.year_var.get().strip(),
                self.season_var.get(),
            )
            dialog.destroy()
            set_status(self.status_label, "期末周统计已生成", "success")
            self._refresh_checks()
            mb.showinfo(
                "生成成功",
                "已生成期末周统计：\n\n" + output,
                parent=self,
            )
        except PermissionError as exc:
            set_status(self.status_label, "输出文件被占用", "error")
            mb.showerror("文件被占用", str(exc), parent=dialog)
        except Exception as exc:
            set_status(self.status_label, "生成失败", "error")
            mb.showerror("生成失败", str(exc), parent=dialog)

    def _confirm_split_preview(self, dialog, results, row_vars_by_slot):
        try:
            for result in results:
                result.rows = self._collect_preview_rows(row_vars_by_slot[result.slot])
            set_status(self.status_label, "正在生成两份 Excel...", "info")
            self.update_idletasks()
            outputs = write_final_weeks_excels(
                results,
                self.data_mgr.get_output_dir(),
                self.year_var.get().strip(),
                self.season_var.get(),
            )
            dialog.destroy()
            set_status(self.status_label, "两份期末周统计已生成", "success")
            self._refresh_checks()
            mb.showinfo("生成成功", "已生成两份周统计：\n\n" + "\n".join(outputs), parent=self)
        except PermissionError as exc:
            set_status(self.status_label, "输出文件被占用", "error")
            mb.showerror("文件被占用", str(exc), parent=dialog)
        except Exception as exc:
            set_status(self.status_label, "生成失败", "error")
            mb.showerror("生成失败", str(exc), parent=dialog)

    def refresh(self):
        self._refresh_checks()
