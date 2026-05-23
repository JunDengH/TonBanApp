"""月统计 Word 生成 Tab。"""
import datetime
import os
import tkinter.filedialog as fd
import tkinter.messagebox as mb
from pathlib import Path

import customtkinter as ctk

from src.modules.word_generator import generate_monthly_word
from src.ui.ui_helpers import (
    CHIP_FG,
    CHIP_HOVER,
    CHIP_SELECTED,
    CHIP_TEXT,
    FONT_BODY,
    FONT_SMALL,
    OUTPUT_DIR,
    add_file_row,
    clear_children,
    make_section,
    make_tab_page,
    muted_button_style,
    set_status,
)


class MonthlyTab(ctk.CTkFrame):
    def __init__(self, parent, data_mgr):
        super().__init__(parent, fg_color="transparent")
        self.data_mgr = data_mgr
        self.month_overtime_records = {}
        self.month_overtime_name_buttons = {}

        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)
        self.grid(row=0, column=0, sticky="nsew")
        self._build()

    def _build(self):
        page = make_tab_page(self, "月统计Word生成")

        files, r = make_section(
            page,
            "周统计文件",
            "可选择 1 到 4 份周统计 Excel；上月 Word 用于继承多值班次和缺班数量。",
        )
        files.grid(row=0, column=0, sticky="ew", padx=4, pady=(0, 6))

        self.week_slots = []
        for i in range(4):
            var = ctk.StringVar(value="")
            self.week_slots.append(var)
            add_file_row(
                files,
                r + i,
                f"第{i + 1}周",
                var,
                f"选择第{i + 1}周周统计Excel",
                lambda idx=i: self._on_pick_month_file(idx),
            )

        self.prev_word_path = ctk.StringVar(value="")
        prev_line = ctk.CTkFrame(files, fg_color="transparent")
        prev_line.grid(row=r + 4, column=0, sticky="ew", padx=16, pady=(8, 10))
        prev_line.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(prev_line, text="上月Word", width=80, anchor="w", font=FONT_BODY).grid(
            row=0, column=0, sticky="w", padx=(0, 8)
        )
        ctk.CTkEntry(
            prev_line,
            textvariable=self.prev_word_path,
            placeholder_text="首次运行或无上月数据时可留空",
        ).grid(row=0, column=1, sticky="ew", padx=(0, 8))
        ctk.CTkButton(prev_line, text="清空", width=62, command=lambda: self.prev_word_path.set("")).grid(
            row=0, column=2, padx=(0, 6)
        )
        ctk.CTkButton(prev_line, text="选择", width=84, command=self._on_pick_prev_word).grid(
            row=0, column=3
        )

        settings, sr = make_section(page, "月统计设置", "标题和月统计加班补录仅影响最终 Word 报表。")
        settings.grid(row=1, column=0, sticky="ew", padx=4, pady=(0, 6))

        overtime = ctk.CTkFrame(settings, fg_color=("gray92", "gray18"), corner_radius=8)
        overtime.grid(row=sr, column=0, sticky="ew", padx=16, pady=(2, 8))
        overtime.grid_columnconfigure(0, weight=6, uniform="overtime")
        overtime.grid_columnconfigure(1, weight=1)
        overtime.grid_columnconfigure(2, weight=5, uniform="overtime")
        overtime.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(overtime, text="加班补录", font=FONT_BODY).grid(
            row=0, column=0, columnspan=3, sticky="w", padx=12, pady=(8, 3)
        )

        self.month_overtime_name_var = ctk.StringVar(value="")
        self.month_overtime_shifts_var = ctk.StringVar(value="1")
        self.month_overtime_search_var = ctk.StringVar(value="")
        self.month_overtime_search_var.trace_add("write", lambda *_: self.refresh_names())
        self.month_overtime_name_frame = ctk.CTkFrame(
            overtime,
            fg_color=("gray88", "gray22"),
            corner_radius=6,
        )
        self.month_overtime_name_frame.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 10))
        self.month_overtime_name_frame.grid_columnconfigure(0, weight=1)
        self.month_overtime_name_frame.grid_rowconfigure(2, weight=1)
        ctk.CTkLabel(
            self.month_overtime_name_frame,
            text="选择加班助理",
            font=FONT_SMALL,
            text_color=("gray35", "gray70"),
        ).grid(row=0, column=0, sticky="w", padx=8, pady=(6, 2))
        ctk.CTkEntry(
            self.month_overtime_name_frame,
            textvariable=self.month_overtime_search_var,
            placeholder_text="搜索加班助理",
            height=32,
        ).grid(row=1, column=0, sticky="ew", padx=8, pady=(0, 4))
        self.month_overtime_name_list = ctk.CTkScrollableFrame(
            self.month_overtime_name_frame,
            height=122,
            fg_color="transparent",
        )
        self.month_overtime_name_list.grid(row=2, column=0, sticky="nsew", padx=8, pady=(0, 6))

        right_panel = ctk.CTkFrame(overtime, fg_color=("gray88", "gray22"), corner_radius=6)
        right_panel.grid(row=1, column=2, sticky="nsew", padx=(0, 12), pady=(0, 10))
        right_panel.grid_columnconfigure(0, weight=1)
        right_panel.grid_rowconfigure(2, weight=1)
        ctk.CTkLabel(
            right_panel,
            text="补录记录",
            font=FONT_SMALL,
            text_color=("gray35", "gray70"),
        ).grid(row=0, column=0, sticky="w", padx=8, pady=(6, 2))

        overtime_controls = ctk.CTkFrame(right_panel, fg_color="transparent")
        overtime_controls.grid(row=1, column=0, sticky="ew", padx=8, pady=(0, 6))
        overtime_controls.grid_columnconfigure(0, weight=1)
        self.month_overtime_selected_label = ctk.CTkLabel(
            overtime_controls,
            text="当前选择：-",
            font=FONT_SMALL,
            text_color=("gray35", "gray70"),
            anchor="w",
        )
        self.month_overtime_selected_label.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 4))
        ctk.CTkEntry(
            overtime_controls,
            textvariable=self.month_overtime_shifts_var,
            placeholder_text="班次",
            width=90,
        ).grid(row=1, column=0, sticky="w", padx=(0, 8), pady=(0, 6))
        ctk.CTkButton(overtime_controls, text="添加", width=70, command=self._on_add_month_overtime).grid(
            row=1, column=1, padx=(0, 6), pady=(0, 6)
        )
        ctk.CTkButton(
            overtime_controls,
            text="清空",
            width=70,
            command=self._on_clear_month_overtime,
            **muted_button_style(),
        ).grid(row=1, column=2, pady=(0, 6))
        self.month_overtime_records_frame = ctk.CTkScrollableFrame(
            right_panel,
            height=108,
            fg_color="transparent",
        )
        self.month_overtime_records_frame.grid(row=2, column=0, sticky="nsew", padx=8, pady=(0, 6))

        title_row = ctk.CTkFrame(settings, fg_color="transparent")
        title_row.grid(row=sr + 1, column=0, sticky="ew", padx=16, pady=(0, 10))
        now = datetime.datetime.now()
        default_season = "春季" if 2 <= now.month <= 7 else "秋季"
        self.title_year_var = ctk.StringVar(value=str(now.year))
        self.title_season_var = ctk.StringVar(value=default_season)
        self.title_week_start_var = ctk.StringVar(value="1")
        self.title_week_end_var = ctk.StringVar(value="4")
        ctk.CTkLabel(title_row, text="标题", width=48, anchor="w", font=FONT_BODY).grid(
            row=0, column=0, sticky="w", padx=(0, 8)
        )
        ctk.CTkEntry(title_row, textvariable=self.title_year_var, width=76).grid(
            row=0, column=1, padx=(0, 6)
        )
        ctk.CTkOptionMenu(title_row, values=["春季", "秋季"], variable=self.title_season_var, width=82).grid(
            row=0, column=2, padx=(0, 6)
        )
        ctk.CTkEntry(title_row, textvariable=self.title_week_start_var, width=54).grid(
            row=0, column=3, padx=(0, 3)
        )
        ctk.CTkLabel(title_row, text="-", font=FONT_BODY).grid(row=0, column=4, padx=(0, 3))
        ctk.CTkEntry(title_row, textvariable=self.title_week_end_var, width=54).grid(
            row=0, column=5, padx=(0, 6)
        )
        ctk.CTkLabel(title_row, text="周助理值班统计", font=FONT_BODY, text_color=("gray35", "gray70")).grid(
            row=0, column=6, sticky="w"
        )

        actions = ctk.CTkFrame(page, fg_color="transparent")
        actions.grid(row=2, column=0, sticky="ew", padx=4, pady=(4, 2))
        actions.grid_columnconfigure(1, weight=1)
        ctk.CTkButton(
            actions,
            text="生成月统计Word",
            command=self._on_generate_monthly_word,
            width=180,
            height=36,
        ).grid(row=0, column=0, padx=(0, 12))
        self.month_status = ctk.CTkLabel(actions, text="等待生成", font=FONT_SMALL, anchor="w")
        self.month_status.grid(row=0, column=1, sticky="ew")

        self.refresh_names()
        self._refresh_month_overtime_textbox()

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

    def refresh_names(self):
        if not hasattr(self, "month_overtime_name_list"):
            return
        names = self.data_mgr.get_name_list()
        current = self.month_overtime_name_var.get().strip()
        keyword = self.month_overtime_search_var.get().strip() if hasattr(self, "month_overtime_search_var") else ""
        visible = [name for name in names if not keyword or keyword in name]
        clear_children(self.month_overtime_name_list)
        self.month_overtime_name_buttons = {}
        if not names:
            self.month_overtime_name_var.set("请先导入名单")
            self.month_overtime_selected_label.configure(text="当前选择：-")
            ctk.CTkLabel(
                self.month_overtime_name_list,
                text="请先导入名单",
                font=FONT_SMALL,
                text_color=("gray35", "gray70"),
            ).grid(row=0, column=0, sticky="w", padx=8, pady=8)
            return
        if not visible:
            ctk.CTkLabel(
                self.month_overtime_name_list,
                text="没有匹配的姓名",
                font=FONT_SMALL,
                text_color=("gray35", "gray70"),
            ).grid(row=0, column=0, sticky="w", padx=8, pady=8)
            return
        selected = current if current in visible else visible[0]
        self.month_overtime_name_var.set(selected)
        self.month_overtime_selected_label.configure(text=f"当前选择：{selected}")
        for idx, name in enumerate(visible):
            self.month_overtime_name_list.grid_columnconfigure(idx % 3, weight=1)
            btn = ctk.CTkButton(
                self.month_overtime_name_list,
                text=name,
                height=30,
                fg_color=CHIP_SELECTED if name == selected else CHIP_FG,
                hover_color=CHIP_HOVER,
                text_color=CHIP_TEXT,
                command=lambda n=name: self._select_month_overtime_name(n),
            )
            btn.grid(row=idx // 3, column=idx % 3, sticky="ew", padx=5, pady=5)
            self.month_overtime_name_buttons[name] = btn

    def _select_month_overtime_name(self, name):
        self.month_overtime_name_var.set(name)
        self.month_overtime_selected_label.configure(text=f"当前选择：{name}")
        for item, btn in self.month_overtime_name_buttons.items():
            btn.configure(fg_color=CHIP_SELECTED if item == name else CHIP_FG)

    def _refresh_month_overtime_textbox(self):
        clear_children(self.month_overtime_records_frame)
        if not self.month_overtime_records:
            ctk.CTkLabel(
                self.month_overtime_records_frame,
                text="当前未设置月统计加班补录",
                font=FONT_SMALL,
                text_color=("gray35", "gray70"),
            ).grid(row=0, column=0, sticky="w", padx=8, pady=8)
            return
        for idx, (name, times) in enumerate(self.month_overtime_records.items()):
            row = idx // 2
            col = idx % 2
            self.month_overtime_records_frame.grid_columnconfigure(col, weight=1)
            card = ctk.CTkFrame(self.month_overtime_records_frame, fg_color=CHIP_FG, corner_radius=8)
            card.grid(row=row, column=col, sticky="ew", padx=5, pady=5)
            card.grid_columnconfigure(0, weight=1)
            ctk.CTkLabel(
                card,
                text=f"{name}  +{times} 班",
                font=FONT_BODY,
                text_color=CHIP_TEXT,
                anchor="w",
            ).grid(row=0, column=0, sticky="ew", padx=10, pady=6)
            ctk.CTkButton(
                card,
                text="移除",
                width=52,
                height=24,
                command=lambda n=name: self._remove_month_overtime(n),
                **muted_button_style(),
            ).grid(row=0, column=1, sticky="e", padx=(0, 8), pady=6)

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
        self.month_overtime_records[name] = self.month_overtime_records.get(name, 0) + int(shifts)
        self._refresh_month_overtime_textbox()

    def _on_clear_month_overtime(self):
        self.month_overtime_records = {}
        self._refresh_month_overtime_textbox()

    def _on_generate_monthly_word(self):
        if not self.data_mgr.has_name_list():
            mb.showwarning("提示", "请先在【总名单管理】中导入总名单。", parent=self)
            return
        paths = [v.get().strip() for v in self.week_slots if v.get().strip()]
        if len(paths) < 1:
            mb.showwarning("提示", "请至少选择一个周统计Excel文件。", parent=self)
            return
        for path in paths:
            if not os.path.exists(path):
                mb.showwarning("提示", f"文件不存在: {path}", parent=self)
                return
        prev_path = self.prev_word_path.get().strip()
        if prev_path and not os.path.exists(prev_path):
            mb.showwarning("提示", f"上月Word文件不存在: {prev_path}", parent=self)
            return

        default_name = f"月度统计_{datetime.datetime.now():%Y%m%d_%H%M%S}.docx"
        save_path = fd.asksaveasfilename(
            parent=self,
            title="保存月统计Word",
            defaultextension=".docx",
            initialfile=default_name,
            initialdir=str(OUTPUT_DIR),
            filetypes=[("Word文件", "*.docx")],
        )
        if not save_path:
            return

        try:
            set_status(self.month_status, "正在生成...", "info")
            self.update_idletasks()
            forced_title = self._build_forced_monthly_title_from_inputs()
            out = generate_monthly_word(
                weekly_excel_paths=paths,
                total_names=self.data_mgr.get_name_list(),
                prev_word_path=prev_path or None,
                output_path=save_path,
                title_text=forced_title,
                overtime_shifts=self.month_overtime_records,
            )
            set_status(self.month_status, f"已生成：{Path(out).name}", "success")
            mb.showinfo("成功", f"月统计Word已生成：\n{out}", parent=self)
        except PermissionError as e:
            set_status(self.month_status, "文件被占用", "error")
            mb.showerror("文件被占用", str(e), parent=self)
        except ValueError as e:
            set_status(self.month_status, "输入或解析失败", "error")
            mb.showerror("输入或解析失败", str(e), parent=self)
        except Exception as e:
            set_status(self.month_status, "生成失败", "error")
            mb.showerror("生成失败", str(e), parent=self)

    def _build_forced_monthly_title_from_inputs(self) -> str:
        year = self.title_year_var.get().strip()
        season = self.title_season_var.get().strip()
        start_week = self.title_week_start_var.get().strip()
        end_week = self.title_week_end_var.get().strip()

        if not (year.isdigit() and len(year) == 4):
            raise ValueError("标题年份请填写4位数字，如 2026。")
        if season not in ("春季", "秋季"):
            raise ValueError("标题学期只能选择 春季 或 秋季。")
        if not (start_week.isdigit() and end_week.isdigit()):
            raise ValueError("标题周数请填写正整数。")
        start = int(start_week)
        end = int(end_week)
        if start <= 0 or end <= 0:
            raise ValueError("标题周数必须大于 0。")
        if start > end:
            raise ValueError("标题起始周不能大于结束周。")
        return f"{year}{season}学期{start}-{end}周助理值班统计"
