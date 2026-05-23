"""大四助理规则页面。"""
import customtkinter as ctk

from src.ui.ui_helpers import (
    CHIP_FG,
    CHIP_SELECTED,
    CHIP_TEXT,
    FONT_BODY,
    FONT_SMALL,
    clear_children,
    make_section,
    make_tab_page,
    muted_button_style,
    render_name_chips,
)


class SeniorRuleTab(ctk.CTkFrame):
    def __init__(self, parent, data_mgr, navigate=None, on_rule_changed=None):
        super().__init__(parent, fg_color="transparent")
        self.data_mgr = data_mgr
        self.navigate = navigate
        self.on_rule_changed = on_rule_changed
        self.senior_vars = {}
        self.senior_tiles = {}

        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)
        self.grid(row=0, column=0, sticky="nsew")
        self._build()

    def _build(self):
        page = make_tab_page(self, "大四助理规则")
        page.grid_columnconfigure(0, weight=1)

        section, r = make_section(
            page,
            "规则设置",
            "开启后，所选大四助理每周应值班次固定为 1 次；放假核减仍优先生效。",
        )
        section.grid(row=0, column=0, sticky="nsew", padx=4, pady=4)
        section.grid_rowconfigure(r + 3, weight=1)

        top = ctk.CTkFrame(section, fg_color=("gray92", "gray18"), corner_radius=8)
        top.grid(row=r, column=0, sticky="ew", padx=16, pady=(4, 10))
        top.grid_columnconfigure(1, weight=1)
        self.senior_rule_enabled_var = ctk.BooleanVar(
            value=self.data_mgr.is_senior_should_fixed_enabled()
        )
        ctk.CTkSwitch(
            top,
            text="大四助理规则",
            variable=self.senior_rule_enabled_var,
            command=self._on_toggle_senior_rule,
            font=FONT_BODY,
        ).grid(row=0, column=0, sticky="w", padx=14, pady=12)
        self.senior_status_label = ctk.CTkLabel(top, text="", font=FONT_SMALL, anchor="e")
        self.senior_status_label.grid(row=0, column=1, sticky="e", padx=14, pady=12)

        self.empty_hint = ctk.CTkFrame(section, fg_color=("gray92", "gray18"), corner_radius=8)
        self.empty_hint.grid(row=r + 1, column=0, sticky="ew", padx=16, pady=(0, 10))
        self.empty_hint.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            self.empty_hint,
            text="还没有总名单。请先导入总名单，再选择大四助理。",
            font=FONT_BODY,
            text_color=("gray35", "gray70"),
        ).grid(row=0, column=0, sticky="w", padx=14, pady=12)
        if self.navigate is not None:
            ctk.CTkButton(
                self.empty_hint,
                text="去总名单管理",
                width=120,
                command=lambda: self.navigate("names"),
            ).grid(row=0, column=1, sticky="e", padx=14, pady=12)

        tools = ctk.CTkFrame(section, fg_color="transparent")
        tools.grid(row=r + 2, column=0, sticky="ew", padx=16, pady=(0, 8))
        tools.grid_columnconfigure(0, weight=1)
        self.senior_search_var = ctk.StringVar(value="")
        self.senior_search_var.trace_add("write", lambda *_: self._refresh_senior_selector())
        ctk.CTkEntry(
            tools,
            textvariable=self.senior_search_var,
            placeholder_text="搜索姓名",
        ).grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ctk.CTkButton(tools, text="全选", width=70, command=self._select_all_seniors).grid(
            row=0, column=1, padx=(0, 6)
        )
        ctk.CTkButton(
            tools,
            text="清空",
            width=70,
            command=self._clear_senior_selection,
            **muted_button_style(),
        ).grid(row=0, column=2)

        self.senior_picker = ctk.CTkScrollableFrame(section, height=390)
        self.senior_picker.grid(row=r + 3, column=0, sticky="nsew", padx=16, pady=(0, 10))

        bottom = ctk.CTkFrame(section, fg_color="transparent")
        bottom.grid(row=r + 4, column=0, sticky="ew", padx=16, pady=(0, 10))
        bottom.grid_columnconfigure(0, weight=1)
        self.pending_status_label = ctk.CTkLabel(bottom, text="", font=FONT_SMALL, anchor="w")
        self.pending_status_label.grid(row=0, column=0, sticky="w", padx=(0, 8))
        ctk.CTkButton(
            bottom,
            text="保存选择",
            width=110,
            command=self._on_save_senior_selection,
        ).grid(row=0, column=1, sticky="e")

        saved_frame = ctk.CTkFrame(section, fg_color=("gray92", "gray18"), corner_radius=8)
        saved_frame.grid(row=r + 5, column=0, sticky="ew", padx=16, pady=(0, 16))
        saved_frame.grid_columnconfigure(0, weight=1)
        saved_frame.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(
            saved_frame,
            text="已保存的大四助理",
            font=FONT_SMALL,
            text_color=("gray35", "gray70"),
        ).grid(row=0, column=0, sticky="w", padx=12, pady=(10, 0))
        self.senior_confirmed_frame = ctk.CTkScrollableFrame(
            saved_frame,
            height=112,
            fg_color="transparent",
        )
        self.senior_confirmed_frame.grid(row=1, column=0, sticky="ew", padx=8, pady=(4, 10))

        self.refresh()

    def refresh(self):
        self.senior_vars = {}
        self.senior_rule_enabled_var.set(self.data_mgr.is_senior_should_fixed_enabled())
        self._refresh_senior_selector()
        self._refresh_senior_status()
        self._refresh_confirmed_senior_textbox()

    def _refresh_senior_selector(self):
        if not hasattr(self, "senior_picker"):
            return
        clear_children(self.senior_picker)
        self.senior_tiles = {}

        names = self.data_mgr.get_name_list()
        has_names = bool(names)
        if has_names:
            self.empty_hint.grid_remove()
        else:
            self.empty_hint.grid()

        saved = set(self.data_mgr.get_senior_assistants())
        for name in names:
            if name not in self.senior_vars:
                self.senior_vars[name] = ctk.BooleanVar(value=name in saved)

        valid_names = set(names)
        for name in list(self.senior_vars):
            if name not in valid_names:
                del self.senior_vars[name]

        keyword = self.senior_search_var.get().strip() if hasattr(self, "senior_search_var") else ""
        visible = [n for n in names if not keyword or keyword in n]
        if not visible:
            text = "没有匹配的姓名" if has_names else "请先导入总名单"
            ctk.CTkLabel(
                self.senior_picker,
                text=text,
                font=FONT_SMALL,
                text_color=("gray35", "gray70"),
            ).grid(row=0, column=0, sticky="w", padx=8, pady=8)
            return

        for idx, name in enumerate(visible):
            row = idx // 4
            col = idx % 4
            self.senior_picker.grid_columnconfigure(col, weight=1)
            tile = ctk.CTkFrame(
                self.senior_picker,
                fg_color=CHIP_SELECTED if self.senior_vars[name].get() else CHIP_FG,
                corner_radius=8,
            )
            tile.grid(row=row, column=col, sticky="ew", padx=6, pady=6)
            tile.grid_columnconfigure(0, weight=1)
            ctk.CTkCheckBox(
                tile,
                text=name,
                variable=self.senior_vars[name],
                font=FONT_BODY,
                text_color=CHIP_TEXT,
                command=lambda n=name: self._on_senior_checked(n),
            ).grid(row=0, column=0, sticky="w", padx=10, pady=7)
            self.senior_tiles[name] = tile

    def _on_senior_checked(self, name):
        tile = self.senior_tiles.get(name)
        if tile is not None:
            tile.configure(fg_color=CHIP_SELECTED if self.senior_vars[name].get() else CHIP_FG)
        self._refresh_senior_status()

    def _refresh_senior_status(self):
        enabled = self.data_mgr.is_senior_should_fixed_enabled()
        chosen_count = len([v for v in self.senior_vars.values() if v.get()])
        saved_count = len(self.data_mgr.get_senior_assistants())
        status = "已开启" if enabled else "已关闭"
        self.senior_status_label.configure(text=f"{status} · 已保存 {saved_count} 人")
        self.pending_status_label.configure(text=f"当前勾选 {chosen_count} 人")

    def _refresh_confirmed_senior_textbox(self):
        seniors = self.data_mgr.get_senior_assistants()
        render_name_chips(
            self.senior_confirmed_frame,
            seniors,
            columns=6,
            empty_text="当前未保存大四助理",
            width=96,
        )

    def _select_all_seniors(self):
        for name, var in self.senior_vars.items():
            var.set(True)
            if name in self.senior_tiles:
                self.senior_tiles[name].configure(fg_color=CHIP_SELECTED)
        self._refresh_senior_status()

    def _clear_senior_selection(self):
        for name, var in self.senior_vars.items():
            var.set(False)
            if name in self.senior_tiles:
                self.senior_tiles[name].configure(fg_color=CHIP_FG)
        self._refresh_senior_status()

    def _on_toggle_senior_rule(self):
        self.data_mgr.set_senior_should_fixed_enabled(self.senior_rule_enabled_var.get())
        self._refresh_senior_status()
        if self.on_rule_changed is not None:
            self.on_rule_changed()

    def _on_save_senior_selection(self):
        names = self.data_mgr.get_name_list()
        selected = [name for name in names if self.senior_vars.get(name) and self.senior_vars[name].get()]
        self.data_mgr.set_senior_assistants(selected)
        self._refresh_senior_status()
        self._refresh_confirmed_senior_textbox()
        if self.on_rule_changed is not None:
            self.on_rule_changed()
