"""总名单管理页面。"""
import re
import tkinter.filedialog as fd
import tkinter.messagebox as mb

import customtkinter as ctk

from src.modules.contact_parser import compute_graduation_seniors, parse_contact_xlsx
from src.modules.word_parser import parse_total_name_list
from src.ui.name_picker import NamePickerPanel
from src.ui.ui_helpers import (
    BORDER_COLOR,
    CARD_BG,
    CARD_SUBTLE_BG,
    CONTENT_BG,
    FONT_BODY,
    FONT_BODY_BOLD,
    FONT_SECTION,
    FONT_SMALL,
    PANEL_BG,
    SUCCESS,
    TEXT_MUTED,
    TEXT_PRIMARY,
    WARNING,
    add_compact_table_row,
    bind_focus_border,
    checkbox_style,
    clear_children,
    input_style,
    make_bottom_action_bar,
    make_check_group,
    make_check_panel,
    make_dialog_header,
    make_tab_page,
    muted_button_style,
    primary_button_style,
    set_check_item,
    set_status,
    subtle_button_style,
    table_cell_padx,
)


NAME_FULLMATCH = re.compile(r"[\u4e00-\u9fa5]{2,4}")
# \u603b\u540d\u5355\u8868\u683c\u5217\u89c4\u683c\uff08\u8868\u5934\u4e0e\u6570\u636e\u884c\u5171\u7528\uff0c\u4fdd\u8bc1\u9010\u5217\u5bf9\u9f50\uff09\u3002
NAME_TABLE_COLUMNS = [
    ("\u5e8f\u53f7", 48),
    ("\u59d3\u540d", 96),
    ("\u90e8\u95e8", 132),
    ("\u4e13\u4e1a", 120),
    ("\u5e74\u7ea7", 68),
    ("\u89c4\u5219\u6807\u8bb0", 132),
]


class NameListTab(ctk.CTkFrame):
    def __init__(self, parent, data_mgr, on_name_list_changed=None):
        super().__init__(parent, fg_color="transparent")
        self.data_mgr = data_mgr
        self.on_name_list_changed = on_name_list_changed
        self.last_import_text = "尚未在本轮导入文件"
        self.last_duplicate_count = 0
        self._editor_dialog = None

        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)
        self.grid(row=0, column=0, sticky="nsew")
        self._build()

    def _build(self):
        page = make_tab_page(self, "总名单", "导入 Excel 通讯录可自动识别 部门/专业/年级 并把最靠前年级标记为毕业季；“编辑”可增删助理并修改姓名、部门、专业、年级和规则。")
        page.grid_rowconfigure(0, weight=1)
        page.grid_columnconfigure(0, weight=1)
        page.grid_columnconfigure(1, weight=0, minsize=330)

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
        self.name_status = make_bottom_action_bar(
            page,
            row=1,
            primary_text="导入名单",
            primary_command=self._on_import_name_list,
            status_text="等待名单操作",
            secondary_buttons=[
                {"text": "编辑", "width": 104, "command": self._open_editor_dialog},
            ],
        )
        self.refresh()

    def _build_workarea(self, parent):
        ctk.CTkLabel(
            parent,
            text="名单工作区",
            font=FONT_SECTION,
            text_color=TEXT_PRIMARY,
            anchor="w",
        ).grid(row=0, column=0, sticky="ew", padx=16, pady=(14, 2))
        ctk.CTkLabel(
            parent,
            text="名单来源是导入的 Excel 通讯录（或 Word 名单）；搜索只影响当前显示，不会修改保存名单。导入通讯录会替换总名单并按届号重设毕业季。",
            font=FONT_SMALL,
            text_color=TEXT_MUTED,
            anchor="w",
            justify="left",
            wraplength=860,
        ).grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 10))

        summary = ctk.CTkFrame(parent, fg_color=CARD_SUBTLE_BG, corner_radius=8, border_width=1, border_color=BORDER_COLOR)
        summary.grid(row=2, column=0, sticky="ew", padx=16, pady=(0, 10))
        summary.grid_columnconfigure((0, 1, 2), weight=1, uniform="name_summary")
        self.total_value = self._make_inline_metric(summary, 0, "总名单", "-")
        self.visible_value = self._make_inline_metric(summary, 1, "当前显示", "-")
        self.rule_value = self._make_inline_metric(summary, 2, "规则关联", "-")

        tools = ctk.CTkFrame(parent, fg_color="transparent")
        tools.grid(row=3, column=0, sticky="ew", padx=16, pady=(0, 8))
        # 搜索栏拉伸占满整行剩余空间，两个按钮固定在右侧。
        tools.grid_columnconfigure(0, weight=1)
        self.search_var = ctk.StringVar(value="")
        self.search_var.trace_add("write", lambda *_: self._refresh_name_rows())
        search = ctk.CTkEntry(
            tools,
            textvariable=self.search_var,
            placeholder_text="搜索姓名",
            height=36,
            **input_style(),
        )
        bind_focus_border(search)
        search.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        ctk.CTkButton(
            tools,
            text="清空搜索",
            width=92,
            height=36,
            command=lambda: self.search_var.set(""),
            **subtle_button_style(),
        ).grid(row=0, column=1, sticky="e", padx=(0, 10))
        ctk.CTkButton(
            tools,
            text="特殊识别",
            width=92,
            height=36,
            command=self._on_special_recognize,
            **subtle_button_style(),
        ).grid(row=0, column=2, sticky="e")

        table_shell = ctk.CTkFrame(parent, fg_color=PANEL_BG, corner_radius=8, border_width=1, border_color=BORDER_COLOR)
        table_shell.grid(row=4, column=0, sticky="nsew", padx=16, pady=(0, 16))
        table_shell.grid_columnconfigure(0, weight=1)
        table_shell.grid_rowconfigure(1, weight=1)

        header = ctk.CTkFrame(table_shell, fg_color=CARD_SUBTLE_BG, corner_radius=0)
        header.grid(row=0, column=0, sticky="ew")
        for col, (text, width) in enumerate(NAME_TABLE_COLUMNS):
            header.grid_columnconfigure(col, minsize=width)
            ctk.CTkLabel(
                header,
                text=text,
                width=width,
                font=FONT_SMALL,
                text_color=TEXT_MUTED,
                anchor="w",
                height=34,
            ).grid(row=0, column=col, sticky="w", padx=table_cell_padx(col))
        # 与数据行一致的末尾弹性占位列（吸收滚动条宽度），保证逐列对齐。
        header.grid_columnconfigure(len(NAME_TABLE_COLUMNS), weight=1)

        self.names_frame = ctk.CTkScrollableFrame(table_shell, fg_color=PANEL_BG, corner_radius=0)
        self.names_frame.grid(row=1, column=0, sticky="nsew", padx=0, pady=0)
        self.names_frame.grid_columnconfigure(0, weight=1)

    def _make_inline_metric(self, parent, column, label, value):
        cell = ctk.CTkFrame(parent, fg_color="transparent")
        cell.grid(row=0, column=column, sticky="ew", padx=12, pady=10)
        cell.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(cell, text=label, font=FONT_SMALL, text_color=TEXT_MUTED, anchor="w").grid(
            row=0, column=0, sticky="ew"
        )
        value_label = ctk.CTkLabel(cell, text=value, font=FONT_BODY_BOLD, text_color=TEXT_PRIMARY, anchor="w")
        value_label.grid(row=1, column=0, sticky="ew", pady=(2, 0))
        return value_label

    def _build_check_panel(self, page):
        panel = make_check_panel(page, title="名单校验")
        panel.grid(row=0, column=1, sticky="nsew", padx=(0, 4), pady=4)
        self.name_check_items = {}
        make_check_group(
            panel,
            row=1,
            title="名单来源",
            items=[
                ("total", "总名单状态"),
                ("visible", "当前名单表"),
                ("duplicates", "重复姓名检查"),
            ],
            registry=self.name_check_items,
        )
        make_check_group(
            panel,
            row=2,
            title="规则关联",
            items=[
                ("senior", "毕业季规则关联"),
                ("leave", "长期请假规则关联"),
                ("manual", "手动添加状态"),
            ],
            registry=self.name_check_items,
        )
        next_box = ctk.CTkFrame(panel, fg_color=CARD_SUBTLE_BG, corner_radius=8, border_width=1, border_color=BORDER_COLOR)
        next_box.grid(row=3, column=0, sticky="ew", padx=12, pady=(0, 14))
        next_box.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(next_box, text="下一步操作", font=FONT_BODY_BOLD, text_color=TEXT_PRIMARY, anchor="w").grid(
            row=0, column=0, sticky="ew", padx=12, pady=(10, 2)
        )
        self.name_next_action = ctk.CTkLabel(
            next_box,
            text="请先导入或添加名单。",
            font=FONT_SMALL,
            text_color=TEXT_MUTED,
            anchor="w",
            justify="left",
            wraplength=270,
        )
        self.name_next_action.grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 10))

    def refresh(self):
        self._refresh_name_rows()

    def _refresh_name_rows(self):
        all_names = self.data_mgr.get_name_list()
        keyword = self.search_var.get().strip() if hasattr(self, "search_var") else ""
        names = [name for name in all_names if not keyword or keyword in name]
        senior_count = len(self.data_mgr.get_senior_assistants())
        leave_count = len(self.data_mgr.get_long_term_leave_assistants())

        self.total_value.configure(text=f"{len(all_names)} 人")
        self.visible_value.configure(text=f"{len(names)} / {len(all_names)} 人")
        self.rule_value.configure(text=f"毕业季 {senior_count} 人 · 请假 {leave_count} 人")
        self._render_name_rows(names, keyword, bool(all_names))
        self._refresh_checks(names, all_names, keyword)

    def _render_name_rows(self, names, keyword, has_total):
        clear_children(self.names_frame)
        if not names:
            text = "没有匹配的姓名，请调整搜索关键词。" if keyword else "暂无总名单，请先导入 Word 或手动添加姓名。"
            empty = ctk.CTkFrame(self.names_frame, fg_color=CARD_BG, corner_radius=0)
            empty.grid(row=0, column=0, sticky="ew")
            empty.grid_columnconfigure(0, weight=1)
            ctk.CTkLabel(
                empty,
                text=text,
                font=FONT_BODY,
                text_color=TEXT_MUTED,
                anchor="w",
                justify="left",
            ).grid(row=0, column=0, sticky="ew", padx=14, pady=14)
            if not has_total:
                ctk.CTkButton(
                    empty,
                    text="导入总名单",
                    width=118,
                    command=self._on_import_name_list,
                    **primary_button_style(),
                ).grid(row=0, column=1, sticky="e", padx=14, pady=12)
            return

        senior_names = set(self.data_mgr.get_senior_assistants())
        leave_names = set(self.data_mgr.get_long_term_leave_assistants())
        for index, name in enumerate(names, start=1):
            department = self.data_mgr.get_name_department(name) or "-"
            major = self.data_mgr.get_name_major(name) or "-"
            grade = self.data_mgr.get_name_grade(name) or "-"
            tags = []
            if name in senior_names:
                tags.append("毕业季")
            if name in leave_names:
                tags.append("长期请假")
            tag_text = "、".join(tags) if tags else "普通名单"
            tag_color = SUCCESS if tags else TEXT_MUTED
            add_compact_table_row(
                self.names_frame,
                index - 1,
                [str(index), name, department, major, grade, tag_text],
                colors=[TEXT_MUTED, TEXT_PRIMARY, TEXT_MUTED, TEXT_PRIMARY, TEXT_MUTED, tag_color],
                bold_columns={1},
                height=40,
                widths=[w for _, w in NAME_TABLE_COLUMNS],
            )

    def _refresh_checks(self, names, all_names, keyword):
        total_count = len(all_names)
        set_check_item(
            self.name_check_items,
            "total",
            f"已保存 {total_count} 人" if total_count else "尚未导入总名单",
            "success" if total_count else "error",
        )
        visible_text = f"当前名单表显示 {len(names)} / {total_count} 人"
        if keyword and not names:
            visible_text = "搜索后无匹配姓名"
        set_check_item(self.name_check_items, "visible", visible_text, "warning" if keyword and not names else "success")
        duplicate_text = (
            f"本轮导入发现 {self.last_duplicate_count} 个重复姓名，已去重"
            if self.last_duplicate_count
            else "未发现需要处理的重复姓名"
        )
        set_check_item(self.name_check_items, "duplicates", duplicate_text, "warning" if self.last_duplicate_count else "success")

        senior_count = len(self.data_mgr.get_senior_assistants())
        leave_count = len(self.data_mgr.get_long_term_leave_assistants())
        set_check_item(self.name_check_items, "senior", f"毕业季规则已关联 {senior_count} 人", "success")
        set_check_item(self.name_check_items, "leave", f"长期请假规则已关联 {leave_count} 人", "success")
        set_check_item(self.name_check_items, "manual", self.last_import_text, "info")
        if total_count:
            self.name_next_action.configure(text="名单已就绪。建议继续核对规则配置。", text_color=SUCCESS)
        else:
            self.name_next_action.configure(text="请先导入 Word 总名单，或用手动添加补录人员。", text_color=WARNING)

    # ---------- 特殊识别：一键删除毕业季（大四）助理 ----------
    def _on_special_recognize(self):
        seniors = self.data_mgr.get_senior_assistants()
        if not seniors:
            mb.showinfo("特殊识别", "当前总名单中没有毕业季（大四）助理。", parent=self)
            return
        preview = "、".join(seniors[:12]) + ("…" if len(seniors) > 12 else "")
        confirmed = mb.askyesno(
            "特殊识别 · 删除毕业季助理",
            f"将从总名单中永久删除 {len(seniors)} 名毕业季（大四）助理：\n\n{preview}\n\n"
            "该操作不可逆：删除后这些助理及其专业、年级、毕业季/长期请假规则记录都会被一并移除。\n"
            "是否继续？",
            icon="warning",
            parent=self,
        )
        if not confirmed:
            return
        removed = len(seniors)
        remaining = [n for n in self.data_mgr.get_name_list() if n not in set(seniors)]
        self.data_mgr.update_name_list(remaining)
        self._after_name_data_changed(f"特殊识别：已删除 {removed} 名毕业季助理")
        mb.showinfo("特殊识别", f"已删除 {removed} 名毕业季（大四）助理。", parent=self)

    def _after_name_data_changed(self, status_text):
        """名单数据变更后的统一收尾：清空搜索、刷新本页、通知其它页。"""
        self.search_var.set("")
        self.last_import_text = status_text
        self.last_duplicate_count = 0
        self.refresh()
        if self.on_name_list_changed is not None:
            self.on_name_list_changed()
        set_status(self.name_status, status_text, "success")

    # ---------- 编辑：增删助理、修改姓名/专业/年级/规则 ----------
    def _open_editor_dialog(self):
        dialog = ctk.CTkToplevel(self)
        dialog.title("编辑助理名单")
        dialog.configure(fg_color=CONTENT_BG)
        dialog.geometry("860x560")
        dialog.minsize(780, 500)
        dialog.transient(self.winfo_toplevel())
        dialog.grab_set()
        dialog.grid_columnconfigure(0, weight=4, uniform="editor")
        dialog.grid_columnconfigure(1, weight=5, uniform="editor")
        dialog.grid_rowconfigure(1, weight=1)
        self._editor_dialog = dialog
        self._editor_original = None

        make_dialog_header(
            dialog,
            "编辑助理名单",
            "左侧选择助理后在右侧修改并保存；也可新增或删除助理。删除不可逆。",
            columnspan=2,
        )

        # 左：助理列表 + 新增
        left = ctk.CTkFrame(dialog, fg_color=CARD_BG, corner_radius=8, border_width=1, border_color=BORDER_COLOR)
        left.grid(row=1, column=0, sticky="nsew", padx=(18, 8), pady=(0, 18))
        left.grid_columnconfigure(0, weight=1)
        left.grid_rowconfigure(1, weight=1)
        self._editor_count_label = ctk.CTkLabel(
            left, text="", font=FONT_SMALL, text_color=TEXT_MUTED, anchor="w"
        )
        self._editor_count_label.grid(row=0, column=0, sticky="ew", padx=14, pady=(12, 4))
        self._editor_picker = NamePickerPanel(
            left,
            mode="single",
            height=360,
            placeholder="搜索助理",
            on_selection_change=self._editor_on_pick,
        )
        self._editor_picker.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 8))
        ctk.CTkButton(
            left,
            text="＋ 新增助理",
            height=34,
            command=self._editor_set_add_mode,
            **subtle_button_style(),
        ).grid(row=2, column=0, sticky="ew", padx=14, pady=(0, 14))

        # 右：助理信息表单
        form = ctk.CTkFrame(dialog, fg_color=CARD_BG, corner_radius=8, border_width=1, border_color=BORDER_COLOR)
        form.grid(row=1, column=1, sticky="nsew", padx=(8, 18), pady=(0, 18))
        form.grid_columnconfigure(0, weight=1)
        self._editor_title_label = ctk.CTkLabel(
            form, text="新增助理", font=FONT_BODY_BOLD, text_color=TEXT_PRIMARY, anchor="w"
        )
        self._editor_title_label.grid(row=0, column=0, sticky="ew", padx=16, pady=(14, 8))

        self._editor_name_var = ctk.StringVar(value="")
        self._editor_department_var = ctk.StringVar(value="")
        self._editor_major_var = ctk.StringVar(value="")
        self._editor_grade_var = ctk.StringVar(value="")
        self._editor_senior_var = ctk.BooleanVar(value=False)
        self._editor_leave_var = ctk.BooleanVar(value=False)

        self._editor_name_entry = self._add_editor_field(form, 1, "姓名", self._editor_name_var, "2-4 个中文字符")
        self._add_editor_field(form, 3, "部门", self._editor_department_var, "如 阳光俱乐部（可留空）")
        self._add_editor_field(form, 5, "专业", self._editor_major_var, "如 新闻、计科（可留空）")
        self._add_editor_field(form, 7, "年级", self._editor_grade_var, "如 22级、大四（可留空）")

        rules = ctk.CTkFrame(form, fg_color=CARD_SUBTLE_BG, corner_radius=8, border_width=1, border_color=BORDER_COLOR)
        rules.grid(row=9, column=0, sticky="ew", padx=16, pady=(10, 6))
        rules.grid_columnconfigure((0, 1), weight=1)
        ctk.CTkLabel(rules, text="规则标记", font=FONT_SMALL, text_color=TEXT_MUTED, anchor="w").grid(
            row=0, column=0, columnspan=2, sticky="ew", padx=12, pady=(10, 2)
        )
        ctk.CTkCheckBox(
            rules, text="毕业季（大四）", variable=self._editor_senior_var, font=FONT_BODY, **checkbox_style()
        ).grid(row=1, column=0, sticky="w", padx=12, pady=(2, 12))
        ctk.CTkCheckBox(
            rules, text="长期请假", variable=self._editor_leave_var, font=FONT_BODY, **checkbox_style()
        ).grid(row=1, column=1, sticky="w", padx=12, pady=(2, 12))

        actions = ctk.CTkFrame(form, fg_color="transparent")
        actions.grid(row=10, column=0, sticky="ew", padx=16, pady=(8, 14))
        actions.grid_columnconfigure(1, weight=1)
        self._editor_delete_btn = ctk.CTkButton(
            actions, text="删除该助理", width=104, command=self._editor_delete, **muted_button_style()
        )
        self._editor_delete_btn.grid(row=0, column=0, sticky="w")
        ctk.CTkButton(
            actions, text="保存", width=104, command=self._editor_save, **primary_button_style()
        ).grid(row=0, column=2, sticky="e")

        close_bar = ctk.CTkFrame(dialog, fg_color="transparent")
        close_bar.grid(row=2, column=0, columnspan=2, sticky="ew", padx=18, pady=(0, 14))
        close_bar.grid_columnconfigure(0, weight=1)
        ctk.CTkButton(
            close_bar, text="关闭", width=88, command=self._close_editor, **muted_button_style()
        ).grid(row=0, column=1, sticky="e")

        def _close():
            self._close_editor()

        dialog.protocol("WM_DELETE_WINDOW", _close)
        dialog.bind("<Escape>", lambda _event: _close())
        self._editor_refresh(select=None)
        self._editor_set_add_mode()

    def _add_editor_field(self, parent, row, label, variable, placeholder):
        ctk.CTkLabel(parent, text=label, font=FONT_BODY, text_color=TEXT_PRIMARY, anchor="w").grid(
            row=row, column=0, sticky="w", padx=16, pady=(8, 2)
        )
        entry = ctk.CTkEntry(parent, textvariable=variable, placeholder_text=placeholder, height=34, **input_style())
        bind_focus_border(entry)
        entry.grid(row=row + 1, column=0, sticky="ew", padx=16)
        return entry

    def _editor_alive(self):
        dialog = getattr(self, "_editor_dialog", None)
        return dialog is not None and dialog.winfo_exists()

    def _close_editor(self):
        if self._editor_alive():
            self._editor_dialog.destroy()
        self._editor_dialog = None

    def _editor_on_pick(self, selected_names):
        if selected_names:
            self._editor_load(selected_names[0])

    def _editor_load(self, name):
        self._editor_original = name
        self._editor_name_var.set(name)
        self._editor_department_var.set(self.data_mgr.get_name_department(name))
        self._editor_major_var.set(self.data_mgr.get_name_major(name))
        self._editor_grade_var.set(self.data_mgr.get_name_grade(name))
        self._editor_senior_var.set(name in set(self.data_mgr.get_senior_assistants()))
        self._editor_leave_var.set(name in set(self.data_mgr.get_long_term_leave_assistants()))
        self._editor_title_label.configure(text=f"编辑助理：{name}")
        self._editor_delete_btn.configure(state="normal")

    def _editor_set_add_mode(self):
        self._editor_original = None
        self._editor_name_var.set("")
        self._editor_department_var.set("")
        self._editor_major_var.set("")
        self._editor_grade_var.set("")
        self._editor_senior_var.set(False)
        self._editor_leave_var.set(False)
        self._editor_title_label.configure(text="新增助理")
        self._editor_delete_btn.configure(state="disabled")
        picker = getattr(self, "_editor_picker", None)
        if picker is not None:
            picker.selected_name = ""
            picker._draw_cards()
        if self._editor_alive():
            self._editor_name_entry.focus_set()

    def _editor_refresh(self, select):
        names = self.data_mgr.get_name_list()
        self._editor_count_label.configure(text=f"共 {len(names)} 名助理")
        self._editor_picker.set_names(names, selected_names=[select] if select else [])
        if select:
            self._editor_load(select)

    def _editor_save(self):
        name = self._editor_name_var.get().strip()
        department = self._editor_department_var.get().strip()
        major = self._editor_major_var.get().strip()
        grade = self._editor_grade_var.get().strip()
        is_senior = bool(self._editor_senior_var.get())
        is_leave = bool(self._editor_leave_var.get())
        if not NAME_FULLMATCH.fullmatch(name):
            mb.showwarning("提示", "姓名需为 2-4 个中文字符。", parent=self._editor_dialog)
            return
        original = self._editor_original
        if original is None:
            if self.data_mgr.name_exists(name):
                mb.showwarning("提示", f"{name} 已在总名单中。", parent=self._editor_dialog)
                return
            self.data_mgr.add_name(name, grade, major, department)
            status = f"已新增助理 {name}"
        else:
            if name != original:
                if not self.data_mgr.rename_name(original, name):
                    mb.showwarning("提示", f"重命名失败：{name} 可能已存在。", parent=self._editor_dialog)
                    return
            self.data_mgr.set_name_grade(name, grade)
            self.data_mgr.set_name_major(name, major)
            self.data_mgr.set_name_department(name, department)
            status = f"已更新助理 {name}"
        self.data_mgr.set_name_senior(name, is_senior)
        self.data_mgr.set_name_long_term_leave(name, is_leave)
        self._editor_refresh(select=name)
        self._after_name_data_changed(status)

    def _editor_delete(self):
        original = self._editor_original
        if not original:
            return
        confirmed = mb.askyesno(
            "确认删除",
            f"确定从总名单中删除助理 {original}？\n该操作不可逆，其专业、年级、规则记录会一并移除。",
            icon="warning",
            parent=self._editor_dialog,
        )
        if not confirmed:
            return
        self.data_mgr.remove_name(original)
        self._editor_refresh(select=None)
        self._editor_set_add_mode()
        self._after_name_data_changed(f"已删除助理 {original}")

    def _on_import_name_list(self):
        path = fd.askopenfilename(
            parent=self,
            title="选择通讯录(.xlsx) 或 名单(.docx)",
            filetypes=[
                ("通讯录 / 名单", "*.xlsx *.docx"),
                ("Excel 通讯录", "*.xlsx"),
                ("Word 名单", "*.docx"),
            ],
        )
        if not path:
            return
        if path.lower().endswith(".xlsx"):
            self._import_contact_xlsx(path)
        else:
            self._import_word_name_list(path)

    def _import_contact_xlsx(self, path):
        """导入 Excel 通讯录：自动识别 部门/姓名/专业/年级，并把最靠前年级标记为毕业季。"""
        try:
            records = parse_contact_xlsx(path)
            if not records:
                self.last_import_text = "导入失败：未识别到任何助理"
                set_status(self.name_status, self.last_import_text, "warning")
                self.refresh()
                mb.showwarning("提示", "未识别到任何助理，请检查通讯录内容。", parent=self)
                return
            seniors = compute_graduation_seniors(records)
            self.data_mgr.apply_contact_roster(records, seniors)
            self.last_duplicate_count = 0
            total = self.data_mgr.get_total_count()
            dept_count = len(set(self.data_mgr.get_name_departments().values()))
            self.last_import_text = (
                f"通讯录导入成功：共 {total} 人、{dept_count} 个部门，"
                f"已自动标记毕业季 {len(seniors)} 人"
            )
            self.search_var.set("")
            self.refresh()
            if self.on_name_list_changed is not None:
                self.on_name_list_changed()
            set_status(self.name_status, self.last_import_text, "success")
            mb.showinfo(
                "成功",
                f"通讯录已导入：共 {total} 人、{dept_count} 个部门。\n"
                f"已自动识别部门/专业/年级，并把最靠前年级（毕业季）{len(seniors)} 人标记为毕业季助理。",
                parent=self,
            )
        except Exception as e:
            self.last_import_text = "解析失败，请检查通讯录格式"
            set_status(self.name_status, self.last_import_text, "error")
            self.refresh()
            mb.showerror("解析失败", str(e), parent=self)

    def _import_word_name_list(self, path):
        """兼容旧流程：从 Word 名单仅导入姓名（不含部门/专业/年级）。"""
        try:
            names = parse_total_name_list(path)
            if not names:
                self.last_import_text = "导入失败：未解析到任何姓名"
                set_status(self.name_status, self.last_import_text, "warning")
                self.refresh()
                mb.showwarning("提示", "未解析到任何姓名，请检查文档内容。", parent=self)
                return
            self.last_duplicate_count = len(names) - len(dict.fromkeys(names))
            self.data_mgr.update_name_list(names)
            self.last_import_text = f"导入成功：共 {self.data_mgr.get_total_count()} 人"
            if self.last_duplicate_count:
                self.last_import_text += f"，重复姓名 {self.last_duplicate_count} 个已去重"
            self.search_var.set("")
            self.refresh()
            if self.on_name_list_changed is not None:
                self.on_name_list_changed()
            set_status(self.name_status, self.last_import_text, "success")
            mb.showinfo("成功", f"名单已更新，共 {self.data_mgr.get_total_count()} 人。", parent=self)
        except Exception as e:
            self.last_import_text = "解析失败，请检查 Word 文件格式"
            set_status(self.name_status, self.last_import_text, "error")
            self.refresh()
            mb.showerror("解析失败", str(e), parent=self)
