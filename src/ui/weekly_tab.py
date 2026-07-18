"""周统计生成 Tab。"""
import datetime
import os
import tkinter.filedialog as fd
import tkinter.messagebox as mb
from pathlib import Path

import customtkinter as ctk

from src.modules.excel_generator import (
    build_weekly_rows,
    collect_typo_suspects,
    collect_weekly_warnings,
    generate_weekly_excel_from_rows,
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
    SUCCESS,
    TEXT_MUTED,
    TEXT_PRIMARY,
    WARNING,
    add_file_row,
    checkbox_style,
    make_bottom_action_bar,
    make_check_group,
    make_check_panel,
    make_dialog_header,
    make_tab_page,
    input_style,
    bind_focus_border,
    option_menu_style,
    muted_button_style,
    primary_button_style,
    set_check_item,
    set_status,
    subtle_button_style,
)
from src.utils.output_paths import build_unique_output_path
from src.utils.report_titles import build_weekly_report_title


WEEKLY_FLOW_STEPS = ["选择文件", "放假核减", "生成Excel"]
WEEKLY_LAYOUT_SECTIONS = ["选择排班文件 / 实际 Word", "放假核减", "生成 Excel"]


class WeeklyTab(ctk.CTkFrame):
    def __init__(self, parent, data_mgr):
        super().__init__(parent, fg_color="transparent")
        self.data_mgr = data_mgr

        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)
        self.grid(row=0, column=0, sticky="nsew")
        self._build()

    def _build(self):
        page = make_tab_page(
            self,
            "周统计",
            subtitle="按排班 PDF/Word、实际 Word、放假核减三步生成周统计 Excel。",
        )
        page.grid_rowconfigure(0, weight=1)
        page.grid_columnconfigure(0, weight=1)
        page.grid_columnconfigure(1, weight=0, minsize=330)

        self.week_schedule_word_path = ctk.StringVar(value="")
        self.week_actual_word_path = ctk.StringVar(value="")
        self.week_schedule_word_path.trace_add("write", lambda *_: self._refresh_weekly_check_panel())
        self.week_actual_word_path.trace_add("write", lambda *_: self._refresh_weekly_check_panel())
        now = datetime.datetime.now()
        default_season = "春季" if 2 <= now.month <= 7 else "秋季"
        self.week_title_year_var = ctk.StringVar(value=str(now.year))
        self.week_title_season_var = ctk.StringVar(value=default_season)
        self.week_title_number_var = ctk.StringVar(value="1")
        for title_var in (
            self.week_title_year_var,
            self.week_title_season_var,
            self.week_title_number_var,
        ):
            title_var.trace_add("write", lambda *_: self._refresh_weekly_check_panel())

        workarea = ctk.CTkFrame(
            page,
            fg_color=CARD_BG,
            corner_radius=8,
            border_width=1,
            border_color=BORDER_COLOR,
        )
        workarea.grid(row=0, column=0, sticky="nsew", padx=(4, 10), pady=4)
        workarea.grid_columnconfigure(0, weight=1)
        workarea.grid_rowconfigure(4, weight=1)

        self._build_workarea(workarea)
        self._build_check_panel(page)
        self.week_status = make_bottom_action_bar(
            page,
            row=1,
            primary_text="生成 Excel",
            primary_command=self._on_generate_weekly,
            status_text="等待生成",
            secondary_buttons=[
                {"text": "清空选择", "width": 104, "command": self._clear_weekly_inputs},
            ],
        )
        self._refresh_weekly_check_panel()

    def _build_workarea(self, parent):
        ctk.CTkLabel(parent, text="周统计工作台", font=FONT_SECTION, text_color=TEXT_PRIMARY, anchor="w").grid(
            row=0, column=0, sticky="ew", padx=16, pady=(14, 2)
        )
        ctk.CTkLabel(
            parent,
            text="排班文件（PDF 优先，兼容 Word）；实际名单仍使用 Word。生成不会修改原文件。",
            font=FONT_SMALL,
            text_color=TEXT_MUTED,
            anchor="w",
            justify="left",
            wraplength=860,
        ).grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 10))

        step_strip = ctk.CTkFrame(parent, fg_color=CARD_SUBTLE_BG, corner_radius=8, border_width=1, border_color=BORDER_COLOR)
        step_strip.grid(row=2, column=0, sticky="ew", padx=16, pady=(0, 10))
        step_strip.grid_columnconfigure((0, 1, 2), weight=1, uniform="weekly_steps")
        for idx, (title, desc) in enumerate(
            [
                ("1. 选择文件", "排班 PDF/Word + 实际 Word"),
                ("2. 放假核减", "周一到周五"),
                ("3. 生成 Excel", "预览确认后保存到设置输出目录"),
            ]
        ):
            item = ctk.CTkFrame(step_strip, fg_color="transparent")
            item.grid(row=0, column=idx, sticky="ew", padx=12, pady=10)
            item.grid_columnconfigure(0, weight=1)
            ctk.CTkLabel(item, text=title, font=FONT_BODY_BOLD, text_color=TEXT_PRIMARY, anchor="w").grid(
                row=0, column=0, sticky="ew"
            )
            ctk.CTkLabel(item, text=desc, font=FONT_SMALL, text_color=TEXT_MUTED, anchor="w").grid(
                row=1, column=0, sticky="ew", pady=(2, 0)
            )

        self._build_title_fields(parent, row=3)

        form = ctk.CTkFrame(parent, fg_color=CARD_SUBTLE_BG, corner_radius=8, border_width=1, border_color=BORDER_COLOR)
        form.grid(row=4, column=0, sticky="nsew", padx=16, pady=(0, 16))
        form.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(form, text="选择排班文件 / 实际 Word", font=FONT_BODY_BOLD, text_color=TEXT_PRIMARY, anchor="w").grid(
            row=0, column=0, sticky="ew", padx=14, pady=(12, 2)
        )
        ctk.CTkLabel(
            form,
            text="文件名过长时只显示文件名，完整路径保留在输入变量中。",
            font=FONT_SMALL,
            text_color=TEXT_MUTED,
            anchor="w",
        ).grid(row=1, column=0, sticky="ew", padx=14, pady=(0, 8))
        add_file_row(
            form,
            2,
            "排班文件",
            self.week_schedule_word_path,
            "优先选择 PDF，也兼容 Word",
            self._on_pick_schedule_file,
        )
        add_file_row(
            form,
            3,
            "实际名单",
            self.week_actual_word_path,
            "选择实际周统计名单 Word",
            self._on_pick_actual_file,
        )

        holiday = ctk.CTkFrame(form, fg_color=CARD_BG, corner_radius=8, border_width=1, border_color=BORDER_COLOR)
        holiday.grid(row=4, column=0, sticky="ew", padx=16, pady=(12, 10))
        holiday.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(holiday, text="放假核减", font=FONT_BODY_BOLD, text_color=TEXT_PRIMARY, anchor="w").grid(
            row=0, column=0, sticky="w", padx=12, pady=(10, 2)
        )
        ctk.CTkLabel(
            holiday,
            text="只勾选实际放假的工作日；长期请假不受放假核减影响。期末周请使用侧边栏独立入口。",
            font=FONT_SMALL,
            text_color=TEXT_MUTED,
            anchor="w",
        ).grid(row=0, column=1, sticky="ew", padx=12, pady=(10, 2))

        days = ctk.CTkFrame(holiday, fg_color="transparent")
        days.grid(row=1, column=0, columnspan=2, sticky="ew", padx=8, pady=(4, 10))
        days.grid_columnconfigure((0, 1, 2, 3, 4), weight=1, uniform="holiday_days")
        self.holiday_vars = {}
        for i, day in enumerate(["周一", "周二", "周三", "周四", "周五"]):
            var = ctk.BooleanVar(value=False)
            self.holiday_vars[day] = var
            ctk.CTkCheckBox(
                days,
                text=day,
                variable=var,
                command=self._refresh_weekly_check_panel,
                font=FONT_BODY,
                **checkbox_style(),
            ).grid(row=0, column=i, sticky="w", padx=8, pady=4)

        ctk.CTkLabel(
            form,
            text="生成 Excel 前请看右侧检查区：总名单、两份 Word、规则状态和放假核减都会在此确认。",
            font=FONT_SMALL,
            text_color=TEXT_MUTED,
            anchor="w",
            justify="left",
            wraplength=820,
        ).grid(row=5, column=0, sticky="ew", padx=16, pady=(0, 12))

    def _build_check_panel(self, page):
        panel = make_check_panel(page, title="生成前检查")
        panel.grid(row=0, column=1, sticky="nsew", padx=(0, 4), pady=4)
        self.weekly_check_items = {}
        make_check_group(
            panel,
            row=1,
            title="输入完整性",
            items=[
                ("names", "总名单状态"),
                ("schedule", "排班文件"),
                ("actual", "实际 Word"),
                ("typo", "姓名拼写检查"),
                ("title", "报表标题"),
                ("holidays", "放假核减"),
            ],
            registry=self.weekly_check_items,
        )
        make_check_group(
            panel,
            row=2,
            title="统计规则",
            items=[
                ("senior_rule", "毕业季特殊逻辑"),
                ("leave_rule", "长期请假规则"),
                ("output", "生成周统计 Excel"),
            ],
            registry=self.weekly_check_items,
        )
        next_box = ctk.CTkFrame(panel, fg_color=CARD_SUBTLE_BG, corner_radius=8, border_width=1, border_color=BORDER_COLOR)
        next_box.grid(row=3, column=0, sticky="ew", padx=12, pady=(0, 14))
        next_box.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(next_box, text="下一步操作", font=FONT_BODY_BOLD, text_color=TEXT_PRIMARY, anchor="w").grid(
            row=0, column=0, sticky="ew", padx=12, pady=(10, 2)
        )
        self.weekly_next_action_label = ctk.CTkLabel(
            next_box,
            text="请先补齐必要输入。",
            font=FONT_SMALL,
            text_color=TEXT_MUTED,
            anchor="w",
            justify="left",
            wraplength=270,
        )
        self.weekly_next_action_label.grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 10))

    def _build_title_fields(self, parent, row):
        title_panel = ctk.CTkFrame(parent, fg_color=CARD_SUBTLE_BG, corner_radius=8, border_width=1, border_color=BORDER_COLOR)
        title_panel.grid(row=row, column=0, sticky="ew", padx=16, pady=(0, 10))
        title_panel.grid_columnconfigure(7, weight=1)
        ctk.CTkLabel(
            title_panel,
            text="报表标题",
            width=72,
            anchor="w",
            font=FONT_BODY_BOLD,
            text_color=TEXT_PRIMARY,
        ).grid(row=0, column=0, sticky="w", padx=(14, 8), pady=12)
        year_entry = ctk.CTkEntry(title_panel, textvariable=self.week_title_year_var, width=76, **input_style())
        bind_focus_border(year_entry)
        year_entry.grid(row=0, column=1, padx=(0, 6), pady=12)
        ctk.CTkOptionMenu(
            title_panel,
            values=["春季", "秋季"],
            variable=self.week_title_season_var,
            width=82,
            **option_menu_style(),
        ).grid(row=0, column=2, padx=(0, 6), pady=12)
        ctk.CTkLabel(title_panel, text="学期第", font=FONT_BODY, text_color=TEXT_MUTED).grid(
            row=0, column=3, sticky="w", pady=12
        )
        week_entry = ctk.CTkEntry(title_panel, textvariable=self.week_title_number_var, width=54, **input_style())
        bind_focus_border(week_entry)
        week_entry.grid(row=0, column=4, padx=(6, 6), pady=12)
        ctk.CTkLabel(title_panel, text="周助理值班统计", font=FONT_BODY, text_color=TEXT_MUTED).grid(
            row=0, column=5, sticky="w", pady=12
        )

    def refresh(self):
        """供顶部“刷新”按钮和跨页脏标记统一调用。"""
        self._refresh_weekly_check_panel()

    def _refresh_weekly_check_panel(self):
        if not hasattr(self, "weekly_check_items"):
            return
        total_count = self.data_mgr.get_total_count()
        set_check_item(
            self.weekly_check_items,
            "names",
            f"总名单 {total_count} 人" if total_count else "尚未导入总名单",
            "success" if total_count else "error",
        )

        schedule_path = self.week_schedule_word_path.get().strip() if hasattr(self, "week_schedule_word_path") else ""
        actual_path = self.week_actual_word_path.get().strip() if hasattr(self, "week_actual_word_path") else ""
        self._set_file_check("schedule", schedule_path, "排班文件")
        self._set_file_check("actual", actual_path, "实际 Word")
        self._set_typo_check(schedule_path, actual_path, total_count)

        title_ready = True
        title_message = "标题字段合法"
        try:
            self._build_weekly_output_basename()
        except Exception as exc:
            title_ready = False
            title_message = str(exc)
        set_check_item(self.weekly_check_items, "title", title_message, "success" if title_ready else "error")

        holidays = [day for day, var in getattr(self, "holiday_vars", {}).items() if var.get()]
        if holidays:
            set_check_item(self.weekly_check_items, "holidays", f"已勾选 {len(holidays)} 天：{'、'.join(holidays)}", "warning")
        else:
            set_check_item(self.weekly_check_items, "holidays", "未勾选放假日，将按完整工作周统计", "info")

        senior_mode = {
            "normal": "正常值班",
            "reduced": "少值班",
            "none": "无需值班",
        }.get(self.data_mgr.get_senior_should_mode(), "正常值班")
        senior_count = len(self.data_mgr.get_senior_assistants())
        leave_count = len(self.data_mgr.get_long_term_leave_assistants())
        set_check_item(self.weekly_check_items, "senior_rule", f"{senior_mode}，已保存 {senior_count} 人", "success")
        set_check_item(self.weekly_check_items, "leave_rule", f"长期请假 {leave_count} 人；重叠时毕业季优先", "success")

        set_check_item(self.weekly_check_items, "output", "将预览后直接保存到设置中的输出文件夹", "info")

        ready = (
            total_count > 0
            and bool(schedule_path)
            and os.path.exists(schedule_path)
            and bool(actual_path)
            and os.path.exists(actual_path)
            and title_ready
        )
        if ready:
            self.weekly_next_action_label.configure(text="检查通过。可以预览并生成周统计 Excel。", text_color=SUCCESS)
        else:
            self.weekly_next_action_label.configure(
                text="请补齐总名单、标题、排班文件和实际 Word，并确认文件路径可读取。",
                text_color=WARNING,
            )

    def _set_file_check(self, key, path, label):
        if not path:
            set_check_item(self.weekly_check_items, key, f"等待选择{label}", "error")
        elif os.path.exists(path):
            set_check_item(self.weekly_check_items, key, f"{label} 已选择：{Path(path).name}", "success")
        else:
            set_check_item(self.weekly_check_items, key, f"{label} 路径不可读取", "error")

    def _set_typo_check(self, schedule_path, actual_path, total_count):
        """实时显示疑似打错字数量；按 (路径, 修改时间, 名单内容) 缓存，避免每次刷新重扫 Word。"""
        readable = [p for p in (schedule_path, actual_path) if p and os.path.exists(p)]
        if not total_count or not readable:
            set_check_item(self.weekly_check_items, "typo", "选好总名单与 Word 后自动检查", "info")
            return

        names = self.data_mgr.get_name_list()
        try:
            cache_key = tuple((p, os.path.getmtime(p)) for p in readable) + (tuple(names),)
        except OSError:
            cache_key = None
        if cache_key is not None and getattr(self, "_typo_cache_key", None) == cache_key:
            suspects = self._typo_cache_value
        else:
            try:
                suspects = collect_typo_suspects(schedule_path, actual_path, names)
            except Exception:
                set_check_item(self.weekly_check_items, "typo", "姓名拼写检查暂不可用（生成时仍会校验）", "info")
                return
            self._typo_cache_key = cache_key
            self._typo_cache_value = suspects

        if not suspects:
            set_check_item(self.weekly_check_items, "typo", "未发现疑似打错字", "success")
            return
        strong = sum(1 for s in suspects if s["level"] != "weak")
        weak = sum(1 for s in suspects if s["level"] == "weak")
        parts = []
        if strong:
            parts.append(f"疑似打错字 {strong}")
        if weak:
            parts.append(f"形近待确认 {weak}")
        set_check_item(self.weekly_check_items, "typo", "、".join(parts) + "，生成时可查看明细", "warning")

    def _clear_weekly_inputs(self):
        self.week_schedule_word_path.set("")
        self.week_actual_word_path.set("")
        for var in getattr(self, "holiday_vars", {}).values():
            var.set(False)
        self._refresh_weekly_check_panel()
        set_status(self.week_status, "已清空文件选择", "info")

    def _on_pick_schedule_file(self):
        path = fd.askopenfilename(
            parent=self,
            title="选择排班文件（PDF 优先，兼容 Word）",
        filetypes=[
            ("支持的排班文件", "*.pdf *.docx"),
            ("PDF 排班表", "*.pdf"),
            ("Word 排班表", "*.docx"),
        ],
        )
        if path:
            self.week_schedule_word_path.set(path)

    def _on_pick_actual_file(self):
        path = fd.askopenfilename(
            parent=self,
            title="选择实际周统计名单Word",
            filetypes=[("Word文件", "*.docx")],
        )
        if path:
            self.week_actual_word_path.set(path)

    def _on_generate_weekly(self):
        if not self.data_mgr.has_name_list():
            self._refresh_weekly_check_panel()
            mb.showwarning("提示", "请先在【总名单管理】中导入总名单。", parent=self)
            return

        schedule_word_path = self.week_schedule_word_path.get().strip()
        actual_word_path = self.week_actual_word_path.get().strip()
        if not schedule_word_path or not os.path.exists(schedule_word_path):
            self._refresh_weekly_check_panel()
            mb.showwarning("提示", "请选择有效的排班 PDF 或 Word 文件。", parent=self)
            return
        if not actual_word_path or not os.path.exists(actual_word_path):
            self._refresh_weekly_check_panel()
            mb.showwarning("提示", "请选择有效的实际周统计名单Word文件。", parent=self)
            return

        try:
            output_basename = self._build_weekly_output_basename()
            total_names = self.data_mgr.get_name_list()

            # Step 1: 获取疑似错字（不使用纠错映射，用于发现错字）
            set_status(self.week_status, "正在检查姓名拼写...", "info")
            self.update_idletasks()
            typo_suspects = collect_typo_suspects(schedule_word_path, actual_word_path, total_names)

            # Step 2: 弹纠错对话框（用户选择确认的错字）
            corrections = {}
            if typo_suspects:
                corrections = self._show_typo_correction_dialog(typo_suspects)
                if corrections is None:
                    set_status(self.week_status, "已取消生成", "info")
                    return

            # Step 3: 用纠错映射构建周统计行
            set_status(self.week_status, "正在生成预览...", "info")
            self.update_idletasks()
            rows = build_weekly_rows(
                word_path=schedule_word_path,
                total_names=total_names,
                holidays=[d for d, v in self.holiday_vars.items() if v.get()],
                actual_word_path=actual_word_path,
                senior_assistants=self.data_mgr.get_senior_assistants(),
                senior_should_fixed_enabled=self.data_mgr.get_senior_should_mode(),
                long_term_leave_assistants=self.data_mgr.get_long_term_leave_assistants(),
                final_week_enabled=False,
                corrections=corrections,
            )

            # Step 4: 收集剩余警告（未纠错的 typo 仍会出现在这里）
            warnings = collect_weekly_warnings(
                rows,
                schedule_word_path,
                actual_word_path,
                total_names,
                corrections=corrections,
            )

            # Step 5: 如果有剩余提醒（未纠错 typo + 其他），弹窗确认
            remaining_strong = warnings.get("typo_messages_strong", [])
            remaining_weak = warnings.get("typo_messages_weak", [])
            if remaining_strong or remaining_weak or warnings["messages"]:
                sections = []
                if remaining_strong:
                    sections.append(
                        "【疑似打错字（未纠错）】\n"
                        + "\n".join(f"· {msg}" for msg in remaining_strong)
                    )
                if remaining_weak:
                    sections.append(
                        "【形近待确认（未纠错）】\n"
                        + "\n".join(f"· {msg}" for msg in remaining_weak)
                    )
                if warnings["messages"]:
                    sections.append("\n".join(f"· {msg}" for msg in warnings["messages"]))
                summary = "生成前检测到以下情况：\n\n" + "\n\n".join(sections) + "\n\n是否仍要继续生成？"
                if not mb.askyesno("生成前提醒", summary, parent=self):
                    set_status(self.week_status, "已取消生成", "info")
                    return

            # Step 6: 预览
            self._open_weekly_preview(rows, output_basename, warnings["highlight_names"])
            set_status(self.week_status, "请确认预览数据", "info")
        except PermissionError as e:
            set_status(self.week_status, "文件被占用", "error")
            mb.showerror("文件被占用", str(e), parent=self)
        except Exception as e:
            set_status(self.week_status, "生成失败", "error")
            mb.showerror("生成失败", str(e), parent=self)

    def _show_typo_correction_dialog(self, typo_suspects: list) -> dict | None:
        """弹出纠错对话框，让用户勾选确认的错字。

        Returns:
            None: 用户取消生成
            dict: 纠错映射 {"张灏琛": "张颢琛"}，可能为空 {}
        """
        dialog = ctk.CTkToplevel(self)
        dialog.configure(fg_color=CONTENT_BG)
        dialog.title("疑似打错字提醒")
        dialog.geometry("720x440")
        dialog.minsize(640, 360)
        dialog.transient(self.winfo_toplevel())
        dialog.grab_set()
        dialog.grid_columnconfigure(0, weight=1)
        dialog.grid_rowconfigure(2, weight=1)

        make_dialog_header(
            dialog,
            "疑似打错字提醒",
            "勾选确认的错字，系统会在本次生成中自动纠错（不修改原文件）。多候选时可用下拉框选择。",
        )

        # 工具栏：全选 / 清空
        toolbar = ctk.CTkFrame(dialog, fg_color="transparent")
        toolbar.grid(row=1, column=0, sticky="ew", padx=18, pady=(0, 4))
        toolbar.grid_columnconfigure(0, weight=1)

        check_vars: list[ctk.BooleanVar] = []
        candidate_vars: list = []

        list_frame = ctk.CTkScrollableFrame(
            dialog,
            fg_color=CARD_SUBTLE_BG,
            border_width=1,
            border_color=BORDER_COLOR,
        )
        list_frame.grid(row=2, column=0, sticky="nsew", padx=18, pady=(0, 8))
        list_frame.grid_columnconfigure(0, weight=1)

        level_labels = {"strong": "高度疑似", "medium": "疑似", "weak": "形近待确认"}
        level_colors = {"strong": ERROR, "medium": WARNING, "weak": TEXT_MUTED}

        for idx, suspect in enumerate(typo_suspects):
            row = ctk.CTkFrame(list_frame, fg_color="transparent")
            row.grid(row=idx, column=0, sticky="ew", padx=4, pady=3)
            row.grid_columnconfigure(4, weight=1)

            check_var = ctk.BooleanVar(value=False)
            check_vars.append(check_var)
            ctk.CTkCheckBox(
                row, text="", variable=check_var, width=24, **checkbox_style()
            ).grid(row=0, column=0, sticky="w", padx=(8, 4), pady=6)

            ctk.CTkLabel(
                row, text=suspect["name"], width=84,
                font=FONT_BODY_BOLD, text_color=TEXT_PRIMARY, anchor="w",
            ).grid(row=0, column=1, sticky="w", padx=(4, 8), pady=6)

            ctk.CTkLabel(
                row, text="→", font=FONT_BODY, text_color=TEXT_MUTED,
            ).grid(row=0, column=2, sticky="w", padx=(0, 8), pady=6)

            candidates = suspect["candidates"]
            if len(candidates) == 1:
                ctk.CTkLabel(
                    row, text=candidates[0], width=88,
                    font=FONT_BODY, text_color=SUCCESS, anchor="w",
                ).grid(row=0, column=3, sticky="w", pady=6)
                candidate_vars.append(candidates[0])
            else:
                cand_var = ctk.StringVar(value=candidates[0])
                candidate_vars.append(cand_var)
                ctk.CTkOptionMenu(
                    row, values=candidates, variable=cand_var, width=104,
                    **option_menu_style(),
                ).grid(row=0, column=3, sticky="w", pady=6)

            level = suspect["level"]
            info_text = f"{suspect['source']}  ·  {level_labels.get(level, level)}"
            ctk.CTkLabel(
                row, text=info_text, font=FONT_SMALL,
                text_color=level_colors.get(level, TEXT_MUTED), anchor="e",
            ).grid(row=0, column=4, sticky="e", padx=(8, 12), pady=6)

        def _select_all():
            for v in check_vars:
                v.set(True)

        def _clear_all():
            for v in check_vars:
                v.set(False)

        ctk.CTkButton(
            toolbar, text="全选", width=70, height=30,
            command=_select_all, **subtle_button_style(),
        ).grid(row=0, column=1, sticky="e", padx=(0, 8))
        ctk.CTkButton(
            toolbar, text="清空", width=70, height=30,
            command=_clear_all, **muted_button_style(),
        ).grid(row=0, column=2, sticky="e")

        actions = ctk.CTkFrame(dialog, fg_color="transparent")
        actions.grid(row=3, column=0, sticky="ew", padx=18, pady=(0, 14))
        actions.grid_columnconfigure(0, weight=1)

        result = {"corrections": None}

        def _confirm():
            corrections = {}
            for i, suspect in enumerate(typo_suspects):
                if check_vars[i].get():
                    cand = candidate_vars[i]
                    if isinstance(cand, str):
                        corrections[suspect["name"]] = cand
                    else:
                        corrections[suspect["name"]] = cand.get()
            result["corrections"] = corrections
            dialog.destroy()

        def _cancel():
            result["corrections"] = None
            dialog.destroy()

        ctk.CTkButton(
            actions, text="取消", width=88,
            command=_cancel, **muted_button_style(),
        ).grid(row=0, column=1, sticky="e", padx=(0, 8))
        ctk.CTkButton(
            actions, text="确定纠错并生成", width=140,
            command=_confirm, **primary_button_style(),
        ).grid(row=0, column=2, sticky="e")

        dialog.bind("<Escape>", lambda _e: _cancel())
        dialog.protocol("WM_DELETE_WINDOW", _cancel)
        dialog.wait_window()
        return result["corrections"]

    def _build_weekly_output_basename(self) -> str:
        return build_weekly_report_title(
            self.week_title_year_var.get(),
            self.week_title_season_var.get(),
            self.week_title_number_var.get(),
        )

    def _open_weekly_preview(self, rows, output_basename, highlight_names=frozenset()):
        dialog = ctk.CTkToplevel(self)
        dialog.configure(fg_color=CONTENT_BG)
        dialog.title("周统计预览")
        dialog.geometry("860x620")
        dialog.minsize(780, 520)
        dialog.transient(self.winfo_toplevel())
        dialog.grab_set()
        dialog.grid_columnconfigure(0, weight=1)
        dialog.grid_rowconfigure(1, weight=1)

        make_dialog_header(
            dialog,
            "周统计预览",
            "姓名不可编辑；可修正应值班次、实际班次和备注，缺班会在生成时重新计算。",
        )

        table = ctk.CTkScrollableFrame(
            dialog,
            fg_color=CARD_SUBTLE_BG,
            border_width=1,
            border_color=BORDER_COLOR,
        )
        table.grid(row=1, column=0, sticky="nsew", padx=18, pady=(0, 10))
        columns = ["序号", "姓名", "应值班次", "实际班次", "缺班", "备注"]
        widths = [48, 126, 92, 92, 76, 220]
        for col, width in enumerate(widths):
            table.grid_columnconfigure(col, weight=1 if col == 5 else 0, minsize=width)
            ctk.CTkLabel(
                table,
                text=columns[col],
                width=width,
                font=FONT_SMALL,
                text_color=TEXT_MUTED,
            ).grid(row=0, column=col, sticky="ew", padx=3, pady=(6, 4))

        row_vars = []
        for row_idx, row_data in enumerate(rows, start=1):
            ctk.CTkLabel(table, text=str(row_idx), width=widths[0], font=FONT_BODY, text_color=TEXT_MUTED).grid(
                row=row_idx, column=0, sticky="ew", padx=3, pady=3
            )
            vars_for_row = {
                "姓名": ctk.StringVar(value=str(row_data["姓名"])),
                "应值班次": ctk.StringVar(value=str(row_data["应值班次"])),
                "实际班次": ctk.StringVar(value=str(row_data["实际班次"])),
                "缺班": ctk.StringVar(value=str(row_data["缺班"])),
                "备注": ctk.StringVar(value=str(row_data.get("备注", ""))),
            }
            is_flagged = str(row_data["姓名"]) in highlight_names
            name_color = WARNING if is_flagged else TEXT_PRIMARY
            ctk.CTkLabel(table, textvariable=vars_for_row["姓名"], width=widths[1], font=FONT_BODY, text_color=name_color).grid(
                row=row_idx, column=1, sticky="ew", padx=3, pady=3
            )
            for col_idx, key in enumerate(["应值班次", "实际班次"], start=2):
                entry = ctk.CTkEntry(table, textvariable=vars_for_row[key], width=widths[col_idx], height=30, **input_style())
                if is_flagged and key == "应值班次":
                    entry.configure(text_color=WARNING)
                bind_focus_border(entry)
                entry.grid(row=row_idx, column=col_idx, sticky="ew", padx=3, pady=3)
                vars_for_row[key].trace_add("write", lambda *_args, item=vars_for_row: self._sync_weekly_absence_var(item))
            absence_entry = ctk.CTkEntry(table, textvariable=vars_for_row["缺班"], width=widths[4], height=30, **input_style())
            absence_entry.configure(state="disabled", text_color=TEXT_MUTED)
            absence_entry.grid(row=row_idx, column=4, sticky="ew", padx=3, pady=3)
            remark_entry = ctk.CTkEntry(table, textvariable=vars_for_row["备注"], width=widths[5], height=30, **input_style())
            bind_focus_border(remark_entry)
            remark_entry.grid(row=row_idx, column=5, sticky="ew", padx=3, pady=3)
            row_vars.append(vars_for_row)

        actions = ctk.CTkFrame(dialog, fg_color="transparent")
        actions.grid(row=2, column=0, sticky="ew", padx=18, pady=(0, 14))
        actions.grid_columnconfigure(0, weight=1)
        ctk.CTkButton(actions, text="取消", width=88, command=dialog.destroy, **muted_button_style()).grid(
            row=0, column=1, padx=(0, 8)
        )
        ctk.CTkButton(
            actions,
            text="生成Excel",
            width=116,
            command=lambda: self._on_confirm_weekly_preview(dialog, row_vars, output_basename),
            **primary_button_style(),
        ).grid(row=0, column=2)
        dialog.bind("<Escape>", lambda _event: dialog.destroy())

    def _sync_weekly_absence_var(self, row_vars):
        should_text = row_vars["应值班次"].get().strip()
        actual_text = row_vars["实际班次"].get().strip()
        if should_text.isdigit() and actual_text.isdigit():
            row_vars["缺班"].set(str(max(0, int(should_text) - int(actual_text))))

    def _on_confirm_weekly_preview(self, dialog, row_vars, output_basename):
        try:
            rows = self._collect_weekly_preview_rows(row_vars)
            output_path = build_unique_output_path(self.data_mgr.get_output_dir(), output_basename, ".xlsx")
            set_status(self.week_status, "正在生成...", "info")
            self.update_idletasks()
            out = generate_weekly_excel_from_rows(rows, str(output_path))
            dialog.destroy()
            set_status(self.week_status, f"已生成：{Path(out).name}", "success")
            self._refresh_weekly_check_panel()
            mb.showinfo("成功", f"周统计已生成：\n{out}", parent=self)
        except PermissionError as e:
            set_status(self.week_status, "文件被占用", "error")
            mb.showerror("文件被占用", str(e), parent=dialog)
        except ValueError as e:
            set_status(self.week_status, "输入或解析失败", "error")
            mb.showwarning("提示", str(e), parent=dialog)
        except Exception as e:
            set_status(self.week_status, "生成失败", "error")
            mb.showerror("生成失败", str(e), parent=dialog)

    def _collect_weekly_preview_rows(self, row_vars):
        rows = []
        for idx, vars_for_row in enumerate(row_vars, start=1):
            name = vars_for_row["姓名"].get().strip()
            should_text = vars_for_row["应值班次"].get().strip()
            actual_text = vars_for_row["实际班次"].get().strip()
            if not should_text.isdigit():
                raise ValueError(f"第 {idx} 行【应值班次】必须为非负整数。")
            if not actual_text.isdigit():
                raise ValueError(f"第 {idx} 行【实际班次】必须为非负整数。")
            should = int(should_text)
            actual = int(actual_text)
            rows.append({
                "姓名": name,
                "应值班次": should,
                "实际班次": actual,
                "缺班": max(0, should - actual),
                "备注": vars_for_row["备注"].get().strip(),
            })
        return rows
