"""总名单管理页面。"""
import re
import tkinter.filedialog as fd
import tkinter.messagebox as mb

import customtkinter as ctk

from src.modules.word_parser import parse_total_name_list
from src.ui.ui_helpers import (
    CHIP_FG,
    CHIP_HOVER,
    CHIP_MUTED,
    CHIP_TEXT,
    FONT_BODY,
    FONT_SMALL,
    add_name_chip,
    clear_children,
    make_section,
    make_tab_page,
    muted_button_style,
)


NAME_FULLMATCH = re.compile(r"[\u4e00-\u9fa5]{2,4}")
GRADE_OPTIONS = ["大一", "大二", "大三", "大四"]


class NameListTab(ctk.CTkFrame):
    def __init__(self, parent, data_mgr, on_name_list_changed=None):
        super().__init__(parent, fg_color="transparent")
        self.data_mgr = data_mgr
        self.on_name_list_changed = on_name_list_changed

        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)
        self.grid(row=0, column=0, sticky="nsew")
        self._build()

    def _build(self):
        page = make_tab_page(self, "总名单管理")

        section, r = make_section(page, "总名单", "从 Word 导入当前助理总名单，名单会保存在本地配置中。")
        section.grid(row=0, column=0, sticky="nsew", padx=4, pady=4)
        section.grid_rowconfigure(r + 2, weight=1)

        summary = ctk.CTkFrame(section, fg_color=("gray92", "gray18"), corner_radius=8)
        summary.grid(row=r, column=0, sticky="ew", padx=16, pady=(4, 10))
        summary.grid_columnconfigure(0, weight=1)
        self.lbl_name_count = ctk.CTkLabel(
            summary,
            text=self._get_name_count_text(),
            font=("Microsoft YaHei", 22, "bold"),
            anchor="w",
        )
        self.lbl_name_count.grid(row=0, column=0, sticky="w", padx=14, pady=(10, 0))
        ctk.CTkLabel(
            summary,
            text="可在名单末尾点击 + 手动添加；年级为大四时会自动计入大四助理规则。",
            font=FONT_SMALL,
            text_color=("gray35", "gray70"),
        ).grid(row=1, column=0, sticky="w", padx=14, pady=(0, 10))

        actions = ctk.CTkFrame(section, fg_color="transparent")
        actions.grid(row=r + 1, column=0, sticky="ew", padx=16, pady=(0, 10))
        actions.grid_columnconfigure(0, weight=1)
        self.search_var = ctk.StringVar(value="")
        self.search_var.trace_add("write", lambda *_: self._refresh_name_chips())
        ctk.CTkEntry(
            actions,
            textvariable=self.search_var,
            placeholder_text="搜索姓名",
        ).grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ctk.CTkButton(
            actions,
            text="导入 / 更新总名单",
            command=self._on_import_name_list,
            height=38,
            width=150,
        ).grid(row=0, column=1, sticky="e", padx=(0, 6))
        ctk.CTkButton(
            actions,
            text="刷新",
            command=self.refresh,
            height=38,
            width=82,
            **muted_button_style(),
        ).grid(row=0, column=2, sticky="e")

        self.names_frame = ctk.CTkScrollableFrame(section, height=520)
        self.names_frame.grid(row=r + 2, column=0, sticky="nsew", padx=16, pady=(0, 16))

        self.refresh()

    def refresh(self):
        self.lbl_name_count.configure(text=self._get_name_count_text())
        self._refresh_name_chips()

    def _get_name_count_text(self):
        return f"{self.data_mgr.get_total_count()} 人"

    def _refresh_name_chips(self):
        names = self.data_mgr.get_name_list()
        keyword = self.search_var.get().strip() if hasattr(self, "search_var") else ""
        if keyword:
            names = [name for name in names if keyword in name]
        self._render_name_chips_with_add(names, keyword=keyword)

    def _render_name_chips_with_add(self, names, keyword=""):
        clear_children(self.names_frame)
        columns = 7
        for col in range(columns):
            self.names_frame.grid_columnconfigure(col, weight=1, uniform="name_chips")
        if not names and keyword:
            ctk.CTkLabel(
                self.names_frame,
                text="没有匹配的姓名",
                font=FONT_BODY,
                text_color=CHIP_MUTED,
            ).grid(row=0, column=0, sticky="w", padx=8, pady=8)
            add_index = columns
        elif not names:
            add_index = 0
        else:
            for idx, name in enumerate(names):
                add_name_chip(self.names_frame, name, idx // columns, idx % columns, width=96)
            add_index = len(names)

        self._add_manual_name_chip(add_index // columns, add_index % columns)

    def _add_manual_name_chip(self, row, column):
        chip = ctk.CTkFrame(self.names_frame, fg_color=CHIP_FG, corner_radius=7)
        chip.grid(row=row, column=column, sticky="ew", padx=5, pady=5)
        chip.grid_columnconfigure(0, weight=1, minsize=96)
        btn = ctk.CTkButton(
            chip,
            text="+",
            height=34,
            fg_color=CHIP_FG,
            hover_color=CHIP_HOVER,
            text_color=CHIP_TEXT,
            font=("Microsoft YaHei", 18, "bold"),
            command=self._open_add_name_dialog,
        )
        btn.grid(row=0, column=0, sticky="ew", padx=0, pady=0)

    def _open_add_name_dialog(self):
        dialog = ctk.CTkToplevel(self)
        dialog.title("手动添加姓名")
        dialog.resizable(False, False)
        dialog.transient(self.winfo_toplevel())
        dialog.grab_set()

        dialog.grid_columnconfigure(0, weight=1)
        content = ctk.CTkFrame(dialog, fg_color="transparent")
        content.grid(row=0, column=0, sticky="nsew", padx=18, pady=18)
        content.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            content,
            text="添加助理",
            font=("Microsoft YaHei", 18, "bold"),
            anchor="w",
        ).grid(row=0, column=0, sticky="w", pady=(0, 10))

        name_var = ctk.StringVar(value="")
        grade_var = ctk.StringVar(value=GRADE_OPTIONS[0])

        ctk.CTkLabel(content, text="姓名", font=FONT_BODY, anchor="w").grid(
            row=1, column=0, sticky="w", pady=(2, 4)
        )
        name_entry = ctk.CTkEntry(content, textvariable=name_var, width=280, placeholder_text="请输入 2-4 个中文字符")
        name_entry.grid(row=2, column=0, sticky="ew")

        ctk.CTkLabel(content, text="年级", font=FONT_BODY, anchor="w").grid(
            row=3, column=0, sticky="w", pady=(12, 4)
        )
        ctk.CTkOptionMenu(content, variable=grade_var, values=GRADE_OPTIONS, width=280).grid(
            row=4, column=0, sticky="ew"
        )

        actions = ctk.CTkFrame(content, fg_color="transparent")
        actions.grid(row=5, column=0, sticky="ew", pady=(18, 0))
        actions.grid_columnconfigure(0, weight=1)
        ctk.CTkButton(
            actions,
            text="取消",
            width=82,
            command=dialog.destroy,
            **muted_button_style(),
        ).grid(row=0, column=1, padx=(0, 8))
        ctk.CTkButton(
            actions,
            text="保存",
            width=92,
            command=lambda: self._save_manual_name(dialog, name_var.get(), grade_var.get()),
        ).grid(row=0, column=2)

        dialog.bind("<Return>", lambda _event: self._save_manual_name(dialog, name_var.get(), grade_var.get()))
        dialog.bind("<Escape>", lambda _event: dialog.destroy())
        dialog.update_idletasks()
        x = self.winfo_rootx() + max((self.winfo_width() - dialog.winfo_width()) // 2, 0)
        y = self.winfo_rooty() + max((self.winfo_height() - dialog.winfo_height()) // 2, 0)
        dialog.geometry(f"+{x}+{y}")
        name_entry.focus_set()

    def _save_manual_name(self, dialog, name, grade):
        name = (name or "").strip()
        if not NAME_FULLMATCH.fullmatch(name):
            mb.showwarning("提示", "姓名需为 2-4 个中文字符。", parent=dialog)
            return
        if name in self.data_mgr.get_name_list():
            mb.showwarning("提示", f"{name} 已在总名单中。", parent=dialog)
            return
        if grade not in GRADE_OPTIONS:
            mb.showwarning("提示", "请选择有效年级。", parent=dialog)
            return
        if not self.data_mgr.add_name(name, grade):
            mb.showwarning("提示", "添加失败，请检查姓名是否重复。", parent=dialog)
            return
        dialog.destroy()
        self.search_var.set("")
        self.refresh()
        if self.on_name_list_changed is not None:
            self.on_name_list_changed()
        extra = "，已自动计入大四助理规则" if grade == "大四" else ""
        mb.showinfo("成功", f"已添加 {name}{extra}。", parent=self)

    def _on_import_name_list(self):
        path = fd.askopenfilename(
            parent=self,
            title="选择总名单Word文件",
            filetypes=[("Word文件", "*.docx")],
        )
        if not path:
            return
        try:
            names = parse_total_name_list(path)
            if not names:
                mb.showwarning("提示", "未解析到任何姓名，请检查文档内容。", parent=self)
                return
            self.data_mgr.update_name_list(names)
            self.refresh()
            if self.on_name_list_changed is not None:
                self.on_name_list_changed()
            mb.showinfo("成功", f"名单已更新，共 {len(names)} 人。", parent=self)
        except Exception as e:
            mb.showerror("解析失败", str(e), parent=self)
