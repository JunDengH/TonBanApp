"""周统计生成 Tab。"""
import datetime
import os
import tkinter.filedialog as fd
import tkinter.messagebox as mb
from pathlib import Path

import customtkinter as ctk

from src.modules.excel_generator import generate_weekly_excel
from src.ui.ui_helpers import (
    FONT_BODY,
    FONT_SMALL,
    OUTPUT_DIR,
    add_file_row,
    make_section,
    make_tab_page,
    set_status,
)


class WeeklyTab(ctk.CTkFrame):
    def __init__(self, parent, data_mgr):
        super().__init__(parent, fg_color="transparent")
        self.data_mgr = data_mgr

        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)
        self.grid(row=0, column=0, sticky="nsew")
        self._build()

    def _build(self):
        page = make_tab_page(self, "周统计生成")

        section, r = make_section(
            page,
            "周统计Excel",
            "排班名单用于应值班次，实际名单用于实际班次；放假日期会核减当天排班人员的应值班次。",
        )
        section.grid(row=0, column=0, sticky="ew", padx=4, pady=4)

        self.week_schedule_word_path = ctk.StringVar(value="")
        self.week_actual_word_path = ctk.StringVar(value="")

        add_file_row(
            section,
            r,
            "排班名单",
            self.week_schedule_word_path,
            "选择排班名单Word",
            self._on_pick_schedule_file,
        )
        add_file_row(
            section,
            r + 1,
            "实际名单",
            self.week_actual_word_path,
            "选择实际周统计名单Word",
            self._on_pick_actual_file,
        )

        holiday_box = ctk.CTkFrame(section, fg_color=("gray92", "gray18"), corner_radius=8)
        holiday_box.grid(row=r + 2, column=0, sticky="ew", padx=16, pady=(12, 4))
        ctk.CTkLabel(holiday_box, text="本周放假", font=FONT_BODY).grid(
            row=0, column=0, sticky="w", padx=12, pady=(10, 4)
        )
        self.holiday_vars = {}
        for i, day in enumerate(["周一", "周二", "周三", "周四", "周五"]):
            var = ctk.BooleanVar(value=False)
            self.holiday_vars[day] = var
            ctk.CTkCheckBox(holiday_box, text=day, variable=var, font=FONT_BODY).grid(
                row=1, column=i, sticky="w", padx=12, pady=(0, 12)
            )

        actions = ctk.CTkFrame(section, fg_color="transparent")
        actions.grid(row=r + 3, column=0, sticky="ew", padx=16, pady=(12, 16))
        actions.grid_columnconfigure(1, weight=1)
        ctk.CTkButton(
            actions,
            text="生成周统计Excel",
            command=self._on_generate_weekly,
            width=180,
            height=38,
        ).grid(row=0, column=0, sticky="w", padx=(0, 12))
        self.week_status = ctk.CTkLabel(actions, text="等待生成", font=FONT_SMALL, anchor="w")
        self.week_status.grid(row=0, column=1, sticky="ew")

    def _on_pick_schedule_file(self):
        path = fd.askopenfilename(
            parent=self,
            title="选择排班名单Word",
            filetypes=[("Word文件", "*.docx")],
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
            mb.showwarning("提示", "请先在【总名单管理】中导入总名单。", parent=self)
            return

        schedule_word_path = self.week_schedule_word_path.get().strip()
        actual_word_path = self.week_actual_word_path.get().strip()
        if not schedule_word_path or not os.path.exists(schedule_word_path):
            mb.showwarning("提示", "请选择有效的排班名单Word文件。", parent=self)
            return
        if not actual_word_path or not os.path.exists(actual_word_path):
            mb.showwarning("提示", "请选择有效的实际周统计名单Word文件。", parent=self)
            return

        default_name = f"周统计_{datetime.datetime.now():%Y%m%d_%H%M%S}.xlsx"
        save_path = fd.asksaveasfilename(
            parent=self,
            title="保存周统计Excel",
            defaultextension=".xlsx",
            initialfile=default_name,
            initialdir=str(OUTPUT_DIR),
            filetypes=[("Excel文件", "*.xlsx")],
        )
        if not save_path:
            return

        try:
            set_status(self.week_status, "正在生成...", "info")
            self.update_idletasks()
            out = generate_weekly_excel(
                word_path=schedule_word_path,
                total_names=self.data_mgr.get_name_list(),
                holidays=[d for d, v in self.holiday_vars.items() if v.get()],
                output_path=save_path,
                actual_word_path=actual_word_path,
                senior_assistants=self.data_mgr.get_senior_assistants(),
                senior_should_fixed_enabled=self.data_mgr.is_senior_should_fixed_enabled(),
            )
            set_status(self.week_status, f"已生成：{Path(out).name}", "success")
            mb.showinfo("成功", f"周统计已生成：\n{out}", parent=self)
        except PermissionError as e:
            set_status(self.week_status, "文件被占用", "error")
            mb.showerror("文件被占用", str(e), parent=self)
        except Exception as e:
            set_status(self.week_status, "生成失败", "error")
            mb.showerror("生成失败", str(e), parent=self)
