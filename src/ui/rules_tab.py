"""规则配置聚合页面。"""
import customtkinter as ctk

from src.modules.data_manager import (
    SENIOR_MODE_NONE,
    SENIOR_MODE_NORMAL,
    SENIOR_MODE_REDUCED,
)
from src.ui.name_picker import NamePickerPanel
from src.ui.ui_helpers import (
    BORDER_COLOR,
    CARD_BG,
    CARD_SUBTLE_BG,
    FONT_BODY,
    FONT_BODY_BOLD,
    FONT_SECTION,
    FONT_SMALL,
    SUCCESS,
    TEXT_MUTED,
    TEXT_PRIMARY,
    WARNING,
    checkbox_style,
    make_bottom_action_bar,
    make_check_group,
    make_check_panel,
    make_tab_page,
    muted_button_style,
    option_menu_style,
    set_check_item,
    set_status,
    subtle_button_style,
)


SENIOR_RULE_TAB = "毕业季规则"
LEAVE_RULE_TAB = "长期请假"
SENIOR_MODE_OPTIONS = {
    "正常值班": SENIOR_MODE_NORMAL,
    "少值班": SENIOR_MODE_REDUCED,
    "无需值班": SENIOR_MODE_NONE,
}
SENIOR_MODE_LABELS = {value: label for label, value in SENIOR_MODE_OPTIONS.items()}
RULE_PRIORITY_NOTE = "规则优先级：同一姓名同时出现在两个账本时，毕业季规则优先于长期请假规则。"


class RulesTab(ctk.CTkFrame):
    def __init__(self, parent, data_mgr, navigate=None, on_rule_changed=None):
        super().__init__(parent, fg_color="transparent")
        self.data_mgr = data_mgr
        self.navigate = navigate
        self.on_rule_changed = on_rule_changed
        self.selected_rule_key = "senior"

        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)
        self.grid(row=0, column=0, sticky="nsew")
        self._build()

    def _build(self):
        page = make_tab_page(
            self,
            "规则配置",
            "在同一张校验台里维护毕业季特殊逻辑和长期请假规则。",
        )
        page.grid_rowconfigure(0, weight=1)
        page.grid_columnconfigure(0, weight=1)
        page.grid_columnconfigure(1, weight=0, minsize=330)

        workarea = ctk.CTkFrame(page, fg_color="transparent")
        workarea.grid(row=0, column=0, sticky="nsew", padx=(4, 10), pady=4)
        workarea.grid_columnconfigure((0, 1), weight=1, uniform="rule_ledgers")
        workarea.grid_rowconfigure(2, weight=1)

        priority = ctk.CTkFrame(
            workarea,
            fg_color=CARD_BG,
            corner_radius=8,
            border_width=1,
            border_color=BORDER_COLOR,
        )
        priority.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        priority.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            priority,
            text=RULE_PRIORITY_NOTE,
            font=FONT_BODY_BOLD,
            text_color=TEXT_PRIMARY,
            anchor="w",
            justify="left",
            wraplength=900,
        ).grid(row=0, column=0, sticky="ew", padx=16, pady=(12, 2))
        ctk.CTkLabel(
            priority,
            text="保存前请先核对右侧检查区；这里不改变统计公式，只维护统计前规则名单。",
            font=FONT_SMALL,
            text_color=TEXT_MUTED,
            anchor="w",
            justify="left",
            wraplength=900,
        ).grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 12))

        self._build_senior_ledger(workarea)
        self._build_leave_ledger(workarea)
        self._build_check_panel(page)
        self.rules_status = make_bottom_action_bar(
            page,
            row=1,
            primary_text="保存毕业季选择",
            primary_command=self._on_save_senior_selection,
            status_text="等待保存规则",
            secondary_buttons=[
                {"text": "保存请假选择", "width": 148, "command": self._on_save_leave_selection, "style": "subtle"},
            ],
            secondary_on_right=True,
        )
        self.refresh()

    def _build_senior_ledger(self, parent):
        ledger = self._make_ledger(parent, 0, "毕业季规则账本", "搜索并勾选毕业季助理，设置应值班次模式。")
        self.senior_rule_mode_var = ctk.StringVar(
            value=SENIOR_MODE_LABELS.get(self.data_mgr.get_senior_should_mode(), "正常值班")
        )
        mode = ctk.CTkFrame(ledger, fg_color=CARD_SUBTLE_BG, corner_radius=8, border_width=1, border_color=BORDER_COLOR)
        mode.grid(row=2, column=0, sticky="ew", padx=14, pady=(0, 10))
        mode.grid_columnconfigure(2, weight=1)
        ctk.CTkLabel(mode, text="应值班次模式", font=FONT_BODY, text_color=TEXT_PRIMARY).grid(
            row=0, column=0, sticky="w", padx=12, pady=10
        )
        ctk.CTkOptionMenu(
            mode,
            values=list(SENIOR_MODE_OPTIONS.keys()),
            variable=self.senior_rule_mode_var,
            command=self._on_change_senior_mode,
            width=118,
            **option_menu_style(),
        ).grid(row=0, column=1, sticky="w", padx=(0, 12), pady=10)
        self.senior_mode_status = ctk.CTkLabel(mode, text="", font=FONT_SMALL, text_color=TEXT_MUTED, anchor="e")
        self.senior_mode_status.grid(row=0, column=2, sticky="e", padx=12, pady=10)

        self.senior_picker = NamePickerPanel(
            ledger,
            mode="multiple",
            height=350,
            placeholder="搜索毕业季助理",
            on_selection_change=lambda _selected: self._refresh_status(),
            on_filter_change=self._refresh_status,
        )
        self.senior_picker.grid(row=3, column=0, sticky="nsew", padx=14, pady=(0, 10))

        tools = self._make_ledger_tools(ledger, 4, self._select_all_seniors, self._clear_senior_selection)
        self.senior_selected_status = ctk.CTkLabel(tools, text="", font=FONT_SMALL, text_color=TEXT_MUTED, anchor="e")
        self.senior_selected_status.grid(row=0, column=2, sticky="e")

    def _build_leave_ledger(self, parent):
        ledger = self._make_ledger(parent, 1, "长期请假账本", "搜索并勾选长期请假助理；重叠时毕业季规则优先。")
        note = ctk.CTkFrame(ledger, fg_color=CARD_SUBTLE_BG, corner_radius=8, border_width=1, border_color=BORDER_COLOR)
        note.grid(row=2, column=0, sticky="ew", padx=14, pady=(0, 10))
        note.grid_columnconfigure(1, weight=1)
        self.leave_rule_enabled_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            note,
            text="长期请假规则启用",
            variable=self.leave_rule_enabled_var,
            state="disabled",
            font=FONT_BODY,
            **checkbox_style(),
        ).grid(row=0, column=0, sticky="w", padx=12, pady=10)
        ctk.CTkLabel(
            note,
            text="已选人员每周应值班次按 2 次处理",
            font=FONT_SMALL,
            text_color=TEXT_MUTED,
            anchor="e",
        ).grid(row=0, column=1, sticky="e", padx=12, pady=10)

        self.leave_picker = NamePickerPanel(
            ledger,
            mode="multiple",
            height=350,
            placeholder="搜索长期请假助理",
            on_selection_change=lambda _selected: self._refresh_status(),
            on_filter_change=self._refresh_status,
        )
        self.leave_picker.grid(row=3, column=0, sticky="nsew", padx=14, pady=(0, 10))

        tools = self._make_ledger_tools(ledger, 4, self._select_all_leave, self._clear_leave_selection)
        self.leave_selected_status = ctk.CTkLabel(tools, text="", font=FONT_SMALL, text_color=TEXT_MUTED, anchor="e")
        self.leave_selected_status.grid(row=0, column=2, sticky="e")

    def _make_ledger(self, parent, column, title, subtitle):
        ledger = ctk.CTkFrame(parent, fg_color=CARD_BG, corner_radius=8, border_width=1, border_color=BORDER_COLOR)
        ledger.grid(row=2, column=column, sticky="nsew", padx=(0, 6) if column == 0 else (6, 0))
        ledger.grid_columnconfigure(0, weight=1)
        ledger.grid_rowconfigure(3, weight=1)
        ctk.CTkLabel(ledger, text=title, font=FONT_SECTION, text_color=TEXT_PRIMARY, anchor="w").grid(
            row=0, column=0, sticky="ew", padx=16, pady=(14, 2)
        )
        ctk.CTkLabel(
            ledger,
            text=subtitle,
            font=FONT_SMALL,
            text_color=TEXT_MUTED,
            anchor="w",
            justify="left",
            wraplength=420,
        ).grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 12))
        return ledger

    def _make_ledger_tools(self, parent, row, select_command, clear_command):
        tools = ctk.CTkFrame(parent, fg_color="transparent")
        tools.grid(row=row, column=0, sticky="ew", padx=14, pady=(0, 14))
        tools.grid_columnconfigure(2, weight=1)
        ctk.CTkButton(
            tools,
            text="全选显示",
            width=84,
            height=30,
            command=select_command,
            **subtle_button_style(),
        ).grid(row=0, column=0, sticky="w", padx=(0, 8))
        ctk.CTkButton(
            tools,
            text="清空显示",
            width=84,
            height=30,
            command=clear_command,
            **muted_button_style(),
        ).grid(row=0, column=1, sticky="w")
        return tools

    def _build_check_panel(self, page):
        panel = make_check_panel(page, title="规则校验")
        panel.grid(row=0, column=1, sticky="nsew", padx=(0, 4), pady=4)
        self.rule_check_items = {}
        make_check_group(
            panel,
            row=1,
            title="账本状态",
            items=[
                ("names", "总名单状态"),
                ("senior", "毕业季账本"),
                ("leave", "长期请假账本"),
            ],
            registry=self.rule_check_items,
        )
        make_check_group(
            panel,
            row=2,
            title="优先级检查",
            items=[
                ("priority", "毕业季优先规则"),
                ("overlap", "重叠人员"),
                ("mode", "毕业季模式"),
            ],
            registry=self.rule_check_items,
        )
        next_box = ctk.CTkFrame(panel, fg_color=CARD_SUBTLE_BG, corner_radius=8, border_width=1, border_color=BORDER_COLOR)
        next_box.grid(row=3, column=0, sticky="ew", padx=12, pady=(0, 14))
        next_box.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(next_box, text="下一步操作", font=FONT_BODY_BOLD, text_color=TEXT_PRIMARY, anchor="w").grid(
            row=0, column=0, sticky="ew", padx=12, pady=(10, 2)
        )
        self.rule_next_action = ctk.CTkLabel(
            next_box,
            text="请先导入总名单。",
            font=FONT_SMALL,
            text_color=TEXT_MUTED,
            anchor="w",
            justify="left",
            wraplength=270,
        )
        self.rule_next_action.grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 10))

    def select_rule(self, rule_key):
        self.selected_rule_key = "leave" if rule_key == "leave" else "senior"
        self._refresh_status()

    def refresh(self):
        names = self.data_mgr.get_name_list()
        self.senior_rule_mode_var.set(
            SENIOR_MODE_LABELS.get(self.data_mgr.get_senior_should_mode(), "正常值班")
        )
        self.senior_picker.set_names(names, selected_names=self.data_mgr.get_senior_assistants())
        self.leave_picker.set_names(names, selected_names=self.data_mgr.get_long_term_leave_assistants())
        self._refresh_status()

    def _refresh_status(self):
        if not hasattr(self, "rule_check_items"):
            return
        total_count = self.data_mgr.get_total_count()
        seniors = self.data_mgr.get_senior_assistants()
        leaves = self.data_mgr.get_long_term_leave_assistants()
        overlap = [name for name in seniors if name in set(leaves)]
        senior_selected = self.senior_picker.get_selected_names() if hasattr(self, "senior_picker") else []
        leave_selected = self.leave_picker.get_selected_names() if hasattr(self, "leave_picker") else []
        senior_visible = len(self.senior_picker.get_visible_names()) if hasattr(self, "senior_picker") else 0
        leave_visible = len(self.leave_picker.get_visible_names()) if hasattr(self, "leave_picker") else 0
        senior_mode = SENIOR_MODE_LABELS.get(self.data_mgr.get_senior_should_mode(), "正常值班")

        self.senior_mode_status.configure(text=f"已保存 {len(seniors)} 人")
        self.senior_selected_status.configure(text=f"显示 {senior_visible} 人 · 勾选 {len(senior_selected)} 人")
        self.leave_selected_status.configure(text=f"显示 {leave_visible} 人 · 勾选 {len(leave_selected)} 人")

        set_check_item(
            self.rule_check_items,
            "names",
            f"总名单 {total_count} 人" if total_count else "尚未导入总名单",
            "success" if total_count else "error",
        )
        set_check_item(self.rule_check_items, "senior", f"已保存毕业季 {len(seniors)} 人", "success")
        set_check_item(self.rule_check_items, "leave", f"已保存长期请假 {len(leaves)} 人", "success")
        set_check_item(self.rule_check_items, "priority", "重叠人员按毕业季规则处理", "success")
        overlap_text = "无重叠人员" if not overlap else f"重叠 {len(overlap)} 人：{'、'.join(overlap[:5])}"
        if len(overlap) > 5:
            overlap_text += "..."
        set_check_item(self.rule_check_items, "overlap", overlap_text, "warning" if overlap else "success")
        set_check_item(self.rule_check_items, "mode", f"当前毕业季模式：{senior_mode}", "success")

        if total_count:
            self.rule_next_action.configure(text="规则可保存。下一步可生成周统计或月统计。", text_color=SUCCESS)
        else:
            self.rule_next_action.configure(text="请先在总名单页导入人员，再配置两个规则账本。", text_color=WARNING)

    def _select_all_seniors(self):
        self.senior_picker.select_all_visible()

    def _clear_senior_selection(self):
        self.senior_picker.clear_visible()

    def _select_all_leave(self):
        self.leave_picker.select_all_visible()

    def _clear_leave_selection(self):
        self.leave_picker.clear_visible()

    def _on_change_senior_mode(self, _choice=None):
        mode = SENIOR_MODE_OPTIONS.get(self.senior_rule_mode_var.get(), SENIOR_MODE_NORMAL)
        self.data_mgr.set_senior_should_mode(mode)
        self._refresh_status()
        if self.on_rule_changed is not None:
            self.on_rule_changed()

    def _on_save_senior_selection(self):
        self.data_mgr.set_senior_assistants(self.senior_picker.get_selected_names())
        self._refresh_status()
        set_status(self.rules_status, "已保存毕业季规则选择", "success")
        if self.on_rule_changed is not None:
            self.on_rule_changed()

    def _on_save_leave_selection(self):
        self.data_mgr.set_long_term_leave_assistants(self.leave_picker.get_selected_names())
        self._refresh_status()
        set_status(self.rules_status, "已保存长期请假规则选择", "success")
        if self.on_rule_changed is not None:
            self.on_rule_changed()
