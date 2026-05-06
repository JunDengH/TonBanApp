# src/ui/main_ui.py
"""
主界面：三个Tab（基础设置 / 周统计 / 月统计）
"""
import os
import datetime
import tkinter.filedialog as fd
import tkinter.messagebox as mb
from tkinter import ttk
from pathlib import Path

import customtkinter as ctk

from src.modules.data_manager import DataManager
from src.modules.word_parser import parse_total_name_list
from src.modules.excel_generator import (
    generate_weekly_excel,
    generate_monthly_excel,  # 保留：回退/兼容用
)
from src.modules.word_generator import generate_monthly_word

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

OUTPUT_DIR = Path(__file__).resolve().parent.parent.parent / "output"
OUTPUT_DIR.mkdir(exist_ok=True)


class TongBanApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("资助服务中心 - 统班应用")
        self.geometry("720x560")

        self.data_mgr = DataManager()

        # Tab视图
        self.tabs = ctk.CTkTabview(self, width=700, height=520)
        self.tabs.pack(padx=10, pady=10, fill="both", expand=True)

        self.tab_base = self.tabs.add("基础设置")
        self.tab_week = self.tabs.add("周统计生成")
        self.tab_month = self.tabs.add("月统计Word生成")

        self._build_base_tab()
        self._build_week_tab()
        self._build_month_tab()

    # -------------------- Tab 1：基础设置 --------------------
    def _build_base_tab(self):
        frame = self.tab_base
        ctk.CTkLabel(frame, text="总名单管理", font=("Microsoft YaHei", 18, "bold"))\
            .pack(pady=(20, 10))

        self.lbl_name_count = ctk.CTkLabel(
            frame, text=self._get_name_count_text(),
            font=("Microsoft YaHei", 14),
        )
        self.lbl_name_count.pack(pady=10)

        ctk.CTkButton(frame, text="导入 / 更新总名单 (Word)",
                      command=self._on_import_name_list, width=240, height=40)\
            .pack(pady=10)

        ctk.CTkButton(frame, text="查看当前名单",
                      command=self._on_view_name_list, width=240, height=40)\
            .pack(pady=10)

        # 名单展示
        self.txt_names = ctk.CTkTextbox(frame, width=600, height=260)
        self.txt_names.pack(pady=10)
        self._refresh_name_textbox()

    def _get_name_count_text(self):
        return f"当前名单人数：{self.data_mgr.get_total_count()} 人"

    def _refresh_name_textbox(self):
        self.txt_names.delete("1.0", "end")
        names = self.data_mgr.get_name_list()
        if names:
            self.txt_names.insert("1.0", "  ".join(names))

    def _on_import_name_list(self):
        path = fd.askopenfilename(
            title="选择总名单Word文件",
            filetypes=[("Word文件", "*.docx")],
        )
        if not path:
            return
        try:
            names = parse_total_name_list(path)
            if not names:
                mb.showwarning("提示", "未解析到任何姓名，请检查文档内容。")
                return
            self.data_mgr.update_name_list(names)
            self.lbl_name_count.configure(text=self._get_name_count_text())
            self._refresh_name_textbox()
            if hasattr(self, "overtime_name_combo"):
                self._refresh_overtime_name_menu()
            mb.showinfo("成功", f"名单已更新，共 {len(names)} 人。")
        except Exception as e:
            mb.showerror("解析失败", str(e))

    def _on_view_name_list(self):
        self._refresh_name_textbox()

    # -------------------- Tab 2：周统计 --------------------
    def _build_week_tab(self):
        frame = self.tab_week
        ctk.CTkLabel(frame, text="周统计Excel生成",
                     font=("Microsoft YaHei", 18, "bold")).pack(pady=(20, 10))

        # 双Word输入：排班名单 + 实际统计名单
        self.week_schedule_word_path = ctk.StringVar(value="")
        self.week_actual_word_path = ctk.StringVar(value="")

        schedule_row = ctk.CTkFrame(frame)
        schedule_row.pack(pady=6, fill="x", padx=20)
        ctk.CTkLabel(schedule_row, text="排班名单:", width=70).pack(side="left")
        ctk.CTkEntry(schedule_row, textvariable=self.week_schedule_word_path,
                     placeholder_text="选择排班名单Word（用于应值班次）").pack(
            side="left", fill="x", expand=True, padx=(0, 8))
        ctk.CTkButton(schedule_row, text="选择文件",
                      command=self._on_pick_schedule_file, width=100).pack(side="right")

        actual_row = ctk.CTkFrame(frame)
        actual_row.pack(pady=6, fill="x", padx=20)
        ctk.CTkLabel(actual_row, text="实际名单:", width=70).pack(side="left")
        ctk.CTkEntry(actual_row, textvariable=self.week_actual_word_path,
                     placeholder_text="选择实际周统计名单Word（用于实际班次）").pack(
            side="left", fill="x", expand=True, padx=(0, 8))
        ctk.CTkButton(actual_row, text="选择文件",
                      command=self._on_pick_actual_file, width=100).pack(side="right")

        # 放假选项
        ctk.CTkLabel(frame, text="勾选本周放假的日期：",
                     font=("Microsoft YaHei", 13)).pack(pady=(15, 5))
        self.holiday_vars = {}
        cb_frame = ctk.CTkFrame(frame)
        cb_frame.pack(pady=5)
        for i, day in enumerate(["周一", "周二", "周三", "周四", "周五"]):
            var = ctk.BooleanVar(value=False)
            self.holiday_vars[day] = var
            ctk.CTkCheckBox(cb_frame, text=day, variable=var)\
                .grid(row=0, column=i, padx=6, pady=6)

        # 加班补录（可选）
        ctk.CTkLabel(frame, text="加班补录：",
                     font=("Microsoft YaHei", 13)).pack(pady=(10, 4))
        overtime_row = ctk.CTkFrame(frame)
        overtime_row.pack(pady=4, fill="x", padx=20)

        self.overtime_records = {}
        self.overtime_shifts_var = ctk.StringVar(value="1")

        # 采用下拉三角样式（展开后可用鼠标滚轮浏览）
        self.overtime_name_var = ctk.StringVar(value="")

        self.overtime_name_combo = ttk.Combobox(
            overtime_row,
            textvariable=self.overtime_name_var,
            state="readonly",
            width=18,
            height=12,
            font=("Microsoft YaHei", 15),
            values=("(请先导入名单)",),
            style="Overtime.TCombobox",
        )
        self.overtime_name_combo.pack(side="left", padx=(0, 8))
        self._init_overtime_combobox_style()
        self.overtime_name_combo.bind("<<ComboboxSelected>>", self._on_overtime_combo_selected)
        self.overtime_name_combo.bind("<FocusOut>", self._on_overtime_combo_focus_out)
        self.overtime_name_combo.bind("<ButtonRelease-1>", self._on_overtime_combo_mouse_release)
        self.overtime_name_combo.bind("<KeyRelease>", self._on_overtime_combo_key_release)
        # 下拉弹出列表字体同步放大
        self.option_add("*TCombobox*Listbox.font", ("Microsoft YaHei", 15))

        ctk.CTkEntry(
            overtime_row,
            textvariable=self.overtime_shifts_var,
            width=90,
            placeholder_text="班次",
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(overtime_row, text="添加", width=70,
                      command=self._on_add_overtime).pack(side="left", padx=(0, 6))
        ctk.CTkButton(overtime_row, text="清空", width=70,
                      command=self._on_clear_overtime).pack(side="left")

        self.overtime_text = ctk.CTkTextbox(frame, width=660, height=68)
        self.overtime_text.pack(pady=(6, 10), padx=20)
        self._refresh_overtime_name_menu()
        self._refresh_overtime_textbox()

        ctk.CTkButton(frame, text="生成周统计Excel",
                      command=self._on_generate_weekly,
                      width=240, height=40).pack(pady=20)

        self.week_status = ctk.CTkLabel(frame, text="", text_color="gray")
        self.week_status.pack(pady=5)

    def _on_pick_schedule_file(self):
        p = fd.askopenfilename(title="选择排班名单Word",
                               filetypes=[("Word文件", "*.docx")])
        if p:
            self.week_schedule_word_path.set(p)

    def _on_pick_actual_file(self):
        p = fd.askopenfilename(title="选择实际周统计名单Word",
                               filetypes=[("Word文件", "*.docx")])
        if p:
            self.week_actual_word_path.set(p)

    def _refresh_overtime_name_menu(self):
        names = self.data_mgr.get_name_list()
        if not names:
            self.overtime_name_combo["values"] = ("(请先导入名单)",)
            self.overtime_name_var.set("(请先导入名单)")
            return
        self.overtime_name_combo["values"] = tuple(names)
        cur = self.overtime_name_var.get().strip()
        if cur not in names:
            self.overtime_name_var.set(names[0])

    def _refresh_overtime_textbox(self):
        self.overtime_text.delete("1.0", "end")
        if not self.overtime_records:
            self.overtime_text.insert("1.0", "当前未设置加班补录")
            return
        lines = [f"{name}：+{times} 班" for name, times in self.overtime_records.items()]
        self.overtime_text.insert("1.0", "\n".join(lines))

    def _on_add_overtime(self):
        names = self.data_mgr.get_name_list()
        if not names:
            mb.showwarning("提示", "请先在【基础设置】中导入总名单。")
            return

        name = self.overtime_name_var.get().strip()
        if name not in names:
            mb.showwarning("提示", "请选择有效的加班助理。")
            return

        shifts = self.overtime_shifts_var.get().strip()
        if not shifts.isdigit() or int(shifts) <= 0:
            mb.showwarning("提示", "加班班次需填写正整数。")
            return

        self.overtime_records[name] = self.overtime_records.get(name, 0) + int(shifts)
        self._refresh_overtime_textbox()

    def _init_overtime_combobox_style(self):
        """名字选择下拉框：系统灰色风格 + 大字体。"""
        style = ttk.Style()
        # 使用 clam 主题以保证 fieldbackground 在 Windows 下可生效（非白色）
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure(
            "Overtime.TCombobox",
            fieldbackground="#E0E0E0",
            background="#E0E0E0",
            foreground="#000000",
            arrowsize=16,
            padding=3,
        )
        style.map(
            "Overtime.TCombobox",
            fieldbackground=[("readonly", "#E0E0E0")],
            foreground=[("readonly", "#000000")],
            selectbackground=[("readonly", "#E0E0E0")],
            selectforeground=[("readonly", "#000000")],
        )

        # 下拉列表改为系统灰色，避免白底
        self.option_add("*TCombobox*Listbox.background", "#E0E0E0")
        self.option_add("*TCombobox*Listbox.foreground", "#000000")
        self.option_add("*TCombobox*Listbox.selectBackground", "#D0D0D0")
        self.option_add("*TCombobox*Listbox.selectForeground", "#000000")

    def _on_overtime_combo_selected(self, _event=None):
        """选中后清除蓝色文本选区。"""
        self.after(30, lambda: self.overtime_name_combo.selection_clear())

    def _on_overtime_combo_focus_out(self, _event=None):
        self.overtime_name_combo.selection_clear()

    def _on_overtime_combo_mouse_release(self, _event=None):
        self.after(30, lambda: self.overtime_name_combo.selection_clear())

    def _on_overtime_combo_key_release(self, _event=None):
        self.after(30, lambda: self.overtime_name_combo.selection_clear())

    def _on_clear_overtime(self):
        self.overtime_records = {}
        self._refresh_overtime_textbox()

    def _on_generate_weekly(self):
        if not self.data_mgr.has_name_list():
            mb.showwarning("提示", "请先在【基础设置】中导入总名单。")
            return
        schedule_word_path = self.week_schedule_word_path.get().strip()
        actual_word_path = self.week_actual_word_path.get().strip()

        if not schedule_word_path or not os.path.exists(schedule_word_path):
            mb.showwarning("提示", "请选择有效的排班名单Word文件。")
            return
        if not actual_word_path or not os.path.exists(actual_word_path):
            mb.showwarning("提示", "请选择有效的实际周统计名单Word文件。")
            return
        holidays = [d for d, v in self.holiday_vars.items() if v.get()]

        # 让用户选择保存路径
        default_name = f"周统计_{datetime.datetime.now():%Y%m%d_%H%M%S}.xlsx"
        save_path = fd.asksaveasfilename(
            title="保存周统计Excel",
            defaultextension=".xlsx",
            initialfile=default_name,
            initialdir=str(OUTPUT_DIR),
            filetypes=[("Excel文件", "*.xlsx")],
        )
        if not save_path:
            return

        try:
            out = generate_weekly_excel(
                word_path=schedule_word_path,
                total_names=self.data_mgr.get_name_list(),
                holidays=holidays,
                overtime_shifts=self.overtime_records,
                output_path=save_path,
                actual_word_path=actual_word_path,
            )
            self.week_status.configure(text=f"✔ 已生成：{out}", text_color="green")
            mb.showinfo("成功", f"周统计已生成：\n{out}")
        except PermissionError as e:
            mb.showerror("文件被占用", str(e))
        except Exception as e:
            mb.showerror("生成失败", str(e))

    # -------------------- Tab 3：月统计 Word --------------------
    def _build_month_tab(self):
        frame = self.tab_month
        ctk.CTkLabel(frame, text="月统计Word生成",
                     font=("Microsoft YaHei", 18, "bold")).pack(pady=(15, 5))
        ctk.CTkLabel(
            frame,
            text="基于 4 份本月周Excel + 上月月度Word（可留空）生成月度Word报表",
            font=("Microsoft YaHei", 11), text_color="gray",
        ).pack(pady=(0, 10))

        # 4 个周 Excel 输入
        self.week_slots = []
        for i in range(4):
            row = ctk.CTkFrame(frame)
            row.pack(pady=4, fill="x", padx=20)
            var = ctk.StringVar(value="")
            self.week_slots.append(var)
            ctk.CTkLabel(row, text=f"第{i+1}周:", width=60).pack(side="left")
            ctk.CTkEntry(row, textvariable=var,
                         placeholder_text=f"选择第{i+1}周的周统计Excel文件")\
                .pack(side="left", fill="x", expand=True, padx=8)
            ctk.CTkButton(row, text="选择", width=80,
                          command=lambda idx=i: self._on_pick_month_file(idx))\
                .pack(side="right")

        # 上月 Word 输入（可选）
        prev_row = ctk.CTkFrame(frame)
        prev_row.pack(pady=(15, 4), fill="x", padx=20)
        self.prev_word_path = ctk.StringVar(value="")
        ctk.CTkLabel(prev_row, text="上月Word:", width=60).pack(side="left")
        ctk.CTkEntry(prev_row, textvariable=self.prev_word_path,
                     placeholder_text="可留空：首次运行或无上月数据时不需要")\
            .pack(side="left", fill="x", expand=True, padx=8)
        ctk.CTkButton(prev_row, text="选择", width=80,
                      command=self._on_pick_prev_word).pack(side="right")
        ctk.CTkButton(prev_row, text="清空", width=60,
                      command=lambda: self.prev_word_path.set(""))\
            .pack(side="right", padx=(0, 4))

        ctk.CTkLabel(
            frame,
            text="提示：勿手动修改上月Word的表格结构；空格或斜线单元格会按 0 参与本月计算。",
            font=("Microsoft YaHei", 10), text_color="gray",
        ).pack(pady=(2, 6))

        # 标题手动配置：年份 / 春秋季 / 周范围
        title_row = ctk.CTkFrame(frame)
        title_row.pack(pady=(4, 8), fill="x", padx=20)

        now = datetime.datetime.now()
        default_season = "春季" if 2 <= now.month <= 7 else "秋季"

        self.title_year_var = ctk.StringVar(value=str(now.year))
        self.title_season_var = ctk.StringVar(value=default_season)
        self.title_week_start_var = ctk.StringVar(value="1")
        self.title_week_end_var = ctk.StringVar(value="4")

        ctk.CTkLabel(title_row, text="标题:", width=42).pack(side="left")
        ctk.CTkEntry(title_row, textvariable=self.title_year_var, width=72,
                     placeholder_text="年份")\
            .pack(side="left", padx=(0, 4))
        ctk.CTkOptionMenu(title_row, values=["春季", "秋季"],
                          variable=self.title_season_var, width=80)\
            .pack(side="left", padx=(0, 4))
        ctk.CTkEntry(title_row, textvariable=self.title_week_start_var, width=52,
                     placeholder_text="起周")\
            .pack(side="left", padx=(0, 2))
        ctk.CTkLabel(title_row, text="-").pack(side="left")
        ctk.CTkEntry(title_row, textvariable=self.title_week_end_var, width=52,
                     placeholder_text="止周")\
            .pack(side="left", padx=(2, 4))
        ctk.CTkLabel(title_row, text="周助理值班统计", text_color="gray")\
            .pack(side="left")

        ctk.CTkButton(frame, text="生成月统计Word",
                      command=self._on_generate_monthly_word,
                      width=240, height=40).pack(pady=15)

        self.month_status = ctk.CTkLabel(frame, text="", text_color="gray")
        self.month_status.pack(pady=5)

    def _on_pick_month_file(self, idx):
        p = fd.askopenfilename(title=f"选择第{idx+1}周Excel",
                               filetypes=[("Excel文件", "*.xlsx")])
        if p:
            self.week_slots[idx].set(p)

    def _on_pick_prev_word(self):
        p = fd.askopenfilename(title="选择上月月度Word",
                               filetypes=[("Word文件", "*.docx")])
        if p:
            self.prev_word_path.set(p)

    def _on_generate_monthly_word(self):
        if not self.data_mgr.has_name_list():
            mb.showwarning("提示", "请先在【基础设置】中导入总名单。")
            return
        paths = [v.get().strip() for v in self.week_slots if v.get().strip()]
        if len(paths) < 1:
            mb.showwarning("提示", "请至少选择一个周统计Excel文件。")
            return
        for p in paths:
            if not os.path.exists(p):
                mb.showwarning("提示", f"文件不存在: {p}")
                return

        prev_path = self.prev_word_path.get().strip()
        if prev_path and not os.path.exists(prev_path):
            mb.showwarning("提示", f"上月Word文件不存在: {prev_path}")
            return

        default_name = f"月度统计_{datetime.datetime.now():%Y%m%d_%H%M%S}.docx"
        save_path = fd.asksaveasfilename(
            title="保存月统计Word",
            defaultextension=".docx",
            initialfile=default_name,
            initialdir=str(OUTPUT_DIR),
            filetypes=[("Word文件", "*.docx")],
        )
        if not save_path:
            return

        try:
            # 强制标题格式：202x春/秋季学期x-x周助理值班统计（全部手动可配）
            forced_title = self._build_forced_monthly_title_from_inputs()
            out = generate_monthly_word(
                weekly_excel_paths=paths,
                total_names=self.data_mgr.get_name_list(),
                prev_word_path=prev_path or None,
                output_path=save_path,
                title_text=forced_title,
            )
            self.month_status.configure(text=f"✔ 已生成：{out}", text_color="green")
            mb.showinfo("成功", f"月统计Word已生成：\n{out}")
        except FileNotFoundError as e:
            mb.showerror("模板缺失", str(e))
        except PermissionError as e:
            mb.showerror("文件被占用", str(e))
        except ValueError as e:
            mb.showerror("输入或解析失败", str(e))
        except Exception as e:
            mb.showerror("生成失败", str(e))

    def _build_forced_monthly_title_from_inputs(self) -> str:
        """
        从用户手填输入拼装标题：
        202x春/秋季学期x-x周助理值班统计
        """
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


def run_app():
    app = TongBanApp()
    app.mainloop()