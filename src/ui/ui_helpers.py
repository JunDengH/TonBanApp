"""UI shared constants and small layout helpers."""
from pathlib import Path

import customtkinter as ctk


ROOT_DIR = Path(__file__).resolve().parent.parent.parent
OUTPUT_DIR = ROOT_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

SPACE_XS = 4
SPACE_SM = 8
SPACE_MD = 12
SPACE_LG = 16
SPACE_XL = 20

FONT_FAMILY = "Microsoft YaHei UI"
FONT_TITLE = (FONT_FAMILY, 22, "bold")
FONT_PAGE_TITLE = (FONT_FAMILY, 24, "bold")
FONT_SECTION = (FONT_FAMILY, 15, "bold")
FONT_BODY = (FONT_FAMILY, 13)
FONT_BODY_BOLD = (FONT_FAMILY, 13, "bold")
FONT_SMALL = (FONT_FAMILY, 11)
FONT_METRIC = (FONT_FAMILY, 22, "bold")

DEFAULT_THEME = "dark"

THEME_PALETTES = {
    "light": {
        "APP_BG": "#EFF3F7",
        "CONTENT_BG": "#F7F9FC",
        "SIDEBAR_BG": "#F3F6FA",
        "SIDEBAR_ACTIVE": "#E8F0FF",
        "SIDEBAR_HOVER": "#EDEFF4",
        "CARD_BG": "#FFFFFF",
        "CARD_SUBTLE_BG": "#F8FAFC",
        "PANEL_BG": "#FFFFFF",
        "BORDER_COLOR": "#D8DEE8",
        "TEXT_PRIMARY": "#111827",
        "TEXT_MUTED": "#5B6472",
        "TEXT_ON_DARK": "#FFFFFF",
        "PRIMARY": "#1D4ED8",
        "PRIMARY_HOVER": "#1E40AF",
        "ACCENT": "#D97706",
        "ACCENT_HOVER": "#B45309",
        "SUCCESS": "#15803D",
        "ERROR": "#B91C1C",
        "WARNING": "#C2410C",
        "MUTED_BUTTON_BG": "#EEF2F7",
        "MUTED_BUTTON_HOVER": "#E5EAF1",
        "SUBTLE_BUTTON_BG": "#FFFFFF",
        "SUBTLE_BUTTON_HOVER": "#F1F5F9",
        "OPTION_BUTTON": "#E5EAF1",
        "OPTION_BUTTON_HOVER": "#D8DEE8",
        "ACTIVE_BORDER": "#C8D7F5",
        "CHIP_FG": "#FFFFFF",
        "CHIP_HOVER": "#F1F5F9",
        "CHIP_SELECTED": "#E8F0FF",
        "CHIP_TEXT": "#1F2937",
        "CHIP_MUTED": "#6B7280",
        "CHIP_BORDER": "#D8DEE8",
        "CHIP_SELECTED_BORDER": "#1D4ED8",
    },
    "dark": {
        "APP_BG": "#141414",
        "CONTENT_BG": "#18181B",
        "SIDEBAR_BG": "#111111",
        "SIDEBAR_ACTIVE": "#2A2A2E",
        "SIDEBAR_HOVER": "#202124",
        "CARD_BG": "#1F1F22",
        "CARD_SUBTLE_BG": "#18181B",
        "PANEL_BG": "#171719",
        "BORDER_COLOR": "#343437",
        "TEXT_PRIMARY": "#F4F4F5",
        "TEXT_MUTED": "#A1A1AA",
        "TEXT_ON_DARK": "#111111",
        "PRIMARY": "#E5E5E5",
        "PRIMARY_HOVER": "#FAFAFA",
        "ACCENT": "#A3A3A3",
        "ACCENT_HOVER": "#D4D4D8",
        "SUCCESS": "#15803D",
        "ERROR": "#DC2626",
        "WARNING": "#D97706",
        "MUTED_BUTTON_BG": "#2A2A2D",
        "MUTED_BUTTON_HOVER": "#333336",
        "SUBTLE_BUTTON_BG": "#222225",
        "SUBTLE_BUTTON_HOVER": "#2B2B2E",
        "OPTION_BUTTON": "#2A2A2D",
        "OPTION_BUTTON_HOVER": "#343437",
        "ACTIVE_BORDER": "#4B4B50",
        "CHIP_FG": "#1F1F22",
        "CHIP_HOVER": "#2B2B2E",
        "CHIP_SELECTED": "#2E2E32",
        "CHIP_TEXT": "#F4F4F5",
        "CHIP_MUTED": "#A1A1AA",
        "CHIP_BORDER": "#343437",
        "CHIP_SELECTED_BORDER": "#C4C4C6",
    },
}

def resolve_color(value):
    """Pick the concrete color for the current appearance mode.

    Theme tokens are stored as ``(light, dark)`` tuples so CustomTkinter can
    switch appearance natively. Plain Tk widgets (``tk.Canvas``) cannot consume
    tuples, so canvas-drawing code resolves the single color through here.
    """
    if isinstance(value, (tuple, list)) and len(value) == 2:
        return value[0] if ctk.get_appearance_mode() == "Light" else value[1]
    return value


def _install_theme_tokens():
    """Expose every color token as a ``(light, dark)`` tuple module global.

    CustomTkinter resolves such tuples against the active appearance mode and
    redraws all widgets in a single pass on ``set_appearance_mode``, so no
    manual per-widget retinting is needed.
    """
    light = THEME_PALETTES["light"]
    dark = THEME_PALETTES["dark"]
    globals().update({name: (light[name], dark[name]) for name in dark})


_install_theme_tokens()


def apply_theme(theme_name=DEFAULT_THEME):
    """Normalize a theme name.

    Kept for backward compatibility: tokens are now constant ``(light, dark)``
    tuples, so switching appearance is handled by ``ctk.set_appearance_mode``
    and this no longer rebinds any globals.
    """
    return theme_name if theme_name in THEME_PALETTES else DEFAULT_THEME


STATUS_MARKS = {
    "success": "[通过]",
    "warning": "[待确认]",
    "error": "[阻塞]",
    "info": "[说明]",
}
STATUS_COLORS = {
    "success": SUCCESS,
    "warning": WARNING,
    "error": ERROR,
    "info": TEXT_MUTED,
}


def primary_button_style():
    return {
        "fg_color": PRIMARY,
        "hover_color": PRIMARY_HOVER,
        "text_color": TEXT_ON_DARK,
        "corner_radius": 8,
    }


def accent_button_style():
    return {
        "fg_color": ACCENT,
        "hover_color": ACCENT_HOVER,
        "text_color": TEXT_ON_DARK,
        "corner_radius": 8,
    }


def muted_button_style():
    return {
        "fg_color": MUTED_BUTTON_BG,
        "hover_color": MUTED_BUTTON_HOVER,
        "text_color": TEXT_PRIMARY,
        "corner_radius": 8,
    }


def subtle_button_style():
    return {
        "fg_color": SUBTLE_BUTTON_BG,
        "hover_color": SUBTLE_BUTTON_HOVER,
        "text_color": TEXT_PRIMARY,
        "corner_radius": 8,
        "border_width": 1,
        "border_color": BORDER_COLOR,
    }


def input_style():
    return {
        "fg_color": PANEL_BG,
        "border_color": BORDER_COLOR,
        "text_color": TEXT_PRIMARY,
        "placeholder_text_color": TEXT_MUTED,
        "corner_radius": 8,
    }


def bind_focus_border(entry):
    """Bind focus events to toggle border color between default and PRIMARY."""
    def _on_focus_in(_event):
        entry.configure(border_color=PRIMARY)

    def _on_focus_out(_event):
        entry.configure(border_color=BORDER_COLOR)

    entry.bind("<FocusIn>", _on_focus_in)
    entry.bind("<FocusOut>", _on_focus_out)


def option_menu_style():
    return {
        "fg_color": PANEL_BG,
        "button_color": OPTION_BUTTON,
        "button_hover_color": OPTION_BUTTON_HOVER,
        "dropdown_fg_color": CARD_BG,
        "dropdown_hover_color": SIDEBAR_ACTIVE,
        "dropdown_text_color": TEXT_PRIMARY,
        "text_color": TEXT_PRIMARY,
        "corner_radius": 8,
    }


def checkbox_style():
    return {
        "fg_color": PRIMARY,
        "hover_color": PRIMARY_HOVER,
        "border_color": BORDER_COLOR,
        "checkmark_color": TEXT_ON_DARK,
        "text_color": TEXT_PRIMARY,
    }


def make_tab_page(tab, title, subtitle=None):
    tab.grid_columnconfigure(0, weight=1)
    tab.grid_rowconfigure(1, weight=1)
    header = ctk.CTkFrame(tab, fg_color="transparent")
    header.grid(row=0, column=0, sticky="ew", padx=18, pady=(14, 8))
    header.grid_columnconfigure(0, weight=1)
    ctk.CTkLabel(
        header,
        text=title,
        font=FONT_PAGE_TITLE,
        text_color=TEXT_PRIMARY,
        anchor="w",
    ).grid(row=0, column=0, sticky="w")
    if subtitle:
        ctk.CTkLabel(
            header,
            text=subtitle,
            font=FONT_SMALL,
            text_color=TEXT_MUTED,
            anchor="w",
            justify="left",
        ).grid(row=1, column=0, sticky="w", pady=(3, 0))
    page = ctk.CTkFrame(tab, fg_color="transparent")
    page.grid(row=1, column=0, sticky="nsew", padx=14, pady=(0, 12))
    page.grid_columnconfigure(0, weight=1)
    return page


def make_section(parent, title, subtitle=None):
    frame = ctk.CTkFrame(
        parent,
        fg_color=CARD_BG,
        corner_radius=8,
        border_width=1,
        border_color=BORDER_COLOR,
    )
    frame.grid_columnconfigure(0, weight=1)
    ctk.CTkLabel(frame, text=title, font=FONT_SECTION, text_color=TEXT_PRIMARY).grid(
        row=0, column=0, sticky="w", padx=16, pady=(12, 2)
    )
    start_row = 1
    if subtitle:
        ctk.CTkLabel(
            frame,
            text=subtitle,
            font=FONT_SMALL,
            text_color=TEXT_MUTED,
            wraplength=860,
            justify="left",
        ).grid(row=1, column=0, sticky="w", padx=16, pady=(0, 10))
        start_row = 2
    return frame, start_row


def make_inner_card(parent, fg_color=CARD_SUBTLE_BG):
    card = ctk.CTkFrame(parent, fg_color=fg_color, corner_radius=8, border_width=1, border_color=BORDER_COLOR)
    card.grid_columnconfigure(0, weight=1)
    return card


def make_subpanel(parent, title, subtitle=None):
    panel = ctk.CTkFrame(
        parent,
        fg_color=CARD_SUBTLE_BG,
        corner_radius=8,
        border_width=1,
        border_color=BORDER_COLOR,
    )
    panel.grid_columnconfigure(0, weight=1)
    ctk.CTkLabel(
        panel,
        text=title,
        font=FONT_BODY_BOLD,
        text_color=TEXT_PRIMARY,
        anchor="w",
    ).grid(row=0, column=0, sticky="w", padx=12, pady=(11, 2 if subtitle else 8))
    start_row = 1
    if subtitle:
        ctk.CTkLabel(
            panel,
            text=subtitle,
            font=FONT_SMALL,
            text_color=TEXT_MUTED,
            anchor="w",
            justify="left",
            wraplength=520,
        ).grid(row=1, column=0, sticky="w", padx=12, pady=(0, 8))
        start_row = 2
    return panel, start_row


def make_dialog_header(parent, title, subtitle=None, columnspan=1):
    header = ctk.CTkFrame(parent, fg_color="transparent")
    header.grid(row=0, column=0, columnspan=columnspan, sticky="ew", padx=18, pady=(16, 8))
    header.grid_columnconfigure(0, weight=1)
    ctk.CTkLabel(
        header,
        text=title,
        font=("Microsoft YaHei UI", 18, "bold"),
        text_color=TEXT_PRIMARY,
        anchor="w",
    ).grid(row=0, column=0, sticky="w")
    if subtitle:
        ctk.CTkLabel(
            header,
            text=subtitle,
            font=FONT_SMALL,
            text_color=TEXT_MUTED,
            anchor="w",
            justify="left",
        ).grid(row=1, column=0, sticky="w", pady=(2, 0))
    return header


def make_action_strip(
    parent,
    row,
    primary_text,
    command,
    status_text="等待操作",
    style="accent",
    padx=16,
    pady=(12, 16),
):
    actions = ctk.CTkFrame(
        parent,
        fg_color=CARD_SUBTLE_BG,
        corner_radius=8,
        border_width=1,
        border_color=BORDER_COLOR,
    )
    actions.grid(row=row, column=0, sticky="ew", padx=padx, pady=pady)
    actions.grid_columnconfigure(1, weight=1)
    button_style = primary_button_style() if style == "primary" else accent_button_style()
    ctk.CTkButton(
        actions,
        text=primary_text,
        command=command,
        width=160,
        height=38,
        **button_style,
    ).grid(row=0, column=0, sticky="w", padx=12, pady=12)
    status = ctk.CTkLabel(
        actions,
        text=status_text,
        font=FONT_SMALL,
        text_color=TEXT_MUTED,
        anchor="w",
    )
    status.grid(row=0, column=1, sticky="ew", padx=(0, 12))
    return status


def _bottom_button_style(style_name):
    if style_name == "muted":
        return muted_button_style()
    if style_name == "primary":
        return primary_button_style()
    if style_name == "accent":
        return accent_button_style()
    return subtle_button_style()


def make_bottom_action_bar(
    parent,
    row,
    primary_text,
    primary_command,
    status_text="等待操作",
    secondary_buttons=None,
    primary_style="primary",
    columnspan=2,
    secondary_on_right=False,
):
    actions = ctk.CTkFrame(
        parent,
        fg_color=CARD_BG,
        corner_radius=8,
        border_width=1,
        border_color=BORDER_COLOR,
    )
    actions.grid(row=row, column=0, columnspan=columnspan, sticky="ew", padx=4, pady=(4, 2))
    primary_kwargs = primary_button_style() if primary_style == "primary" else accent_button_style()

    if secondary_on_right:
        # 状态文字占左侧，所有按钮成组靠右——适合多个对等操作（如两个保存按钮）。
        actions.grid_columnconfigure(0, weight=1)
        status = ctk.CTkLabel(
            actions,
            text=status_text,
            font=FONT_SMALL,
            text_color=TEXT_MUTED,
            anchor="w",
        )
        status.grid(row=0, column=0, sticky="ew", padx=(14, 12), pady=12)
        col = 1
        for button in secondary_buttons or []:
            ctk.CTkButton(
                actions,
                text=button["text"],
                width=button.get("width", 148),
                height=38,
                command=button["command"],
                **_bottom_button_style(button.get("style", "subtle")),
            ).grid(row=0, column=col, sticky="e", padx=(0, 10), pady=12)
            col += 1
        ctk.CTkButton(
            actions,
            text=primary_text,
            width=148,
            height=38,
            command=primary_command,
            **primary_kwargs,
        ).grid(row=0, column=col, sticky="e", padx=(0, 12), pady=12)
        return status

    actions.grid_columnconfigure(8, weight=1)
    col = 0
    for button in secondary_buttons or []:
        ctk.CTkButton(
            actions,
            text=button["text"],
            width=button.get("width", 104),
            height=38,
            command=button["command"],
            **_bottom_button_style(button.get("style", "subtle")),
        ).grid(row=0, column=col, sticky="w", padx=(12 if col == 0 else 0, 8), pady=12)
        col += 1

    status = ctk.CTkLabel(
        actions,
        text=status_text,
        font=FONT_SMALL,
        text_color=TEXT_MUTED,
        anchor="w",
    )
    status.grid(row=0, column=8, sticky="ew", padx=(4 if col else 12, 12))

    ctk.CTkButton(
        actions,
        text=primary_text,
        width=148,
        height=38,
        command=primary_command,
        **primary_kwargs,
    ).grid(row=0, column=9, sticky="e", padx=(0, 12), pady=12)
    return status


def make_empty_state(parent, text, action_text=None, command=None):
    frame = ctk.CTkFrame(parent, fg_color=CARD_SUBTLE_BG, corner_radius=8, border_width=1, border_color=BORDER_COLOR)
    frame.grid_columnconfigure(0, weight=1)
    ctk.CTkLabel(
        frame,
        text=text,
        font=FONT_BODY,
        text_color=TEXT_MUTED,
        anchor="w",
        justify="left",
    ).grid(row=0, column=0, sticky="w", padx=14, pady=12)
    if action_text and command:
        ctk.CTkButton(
            frame,
            text=action_text,
            width=124,
            command=command,
            **primary_button_style(),
        ).grid(row=0, column=1, sticky="e", padx=14, pady=12)
    return frame


def add_file_row(parent, row, label, variable, placeholder, command, button_text="选择"):
    line = ctk.CTkFrame(parent, fg_color="transparent")
    line.grid(row=row, column=0, sticky="ew", padx=16, pady=4)
    line.grid_columnconfigure(1, weight=1)
    ctk.CTkLabel(line, text=label, width=88, anchor="w", font=FONT_BODY, text_color=TEXT_PRIMARY).grid(
        row=0, column=0, sticky="w", padx=(0, 10)
    )
    display_var = ctk.StringVar(value=Path(variable.get()).name if variable.get() else placeholder)

    def _sync_display(*_args):
        raw_value = variable.get().strip()
        display_var.set(Path(raw_value).name if raw_value else placeholder)

    variable.trace_add("write", _sync_display)
    display = ctk.CTkFrame(
        line,
        fg_color=PANEL_BG,
        height=36,
        corner_radius=8,
        border_width=1,
        border_color=BORDER_COLOR,
    )
    display.grid(
        row=0,
        column=1,
        sticky="ew",
        padx=(0, 8),
    )
    display.grid_columnconfigure(0, weight=1)
    ctk.CTkLabel(
        display,
        textvariable=display_var,
        font=FONT_BODY,
        text_color=TEXT_MUTED,
        anchor="w",
    ).grid(row=0, column=0, sticky="ew", padx=12, pady=7)
    ctk.CTkButton(line, text=button_text, width=86, height=36, command=command, **subtle_button_style()).grid(
        row=0, column=2, sticky="e"
    )


def make_list_toolbar(parent, row, search_var, buttons=None, placeholder="搜索姓名", padx=16, pady=(0, 10)):
    toolbar = ctk.CTkFrame(parent, fg_color="transparent")
    toolbar.grid(row=row, column=0, sticky="ew", padx=padx, pady=pady)
    toolbar.grid_columnconfigure(0, weight=1)
    search_entry = ctk.CTkEntry(
        toolbar,
        textvariable=search_var,
        placeholder_text=placeholder,
        height=36,
        **input_style(),
    )
    bind_focus_border(search_entry)
    search_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))
    for idx, button in enumerate(buttons or [], start=1):
        style = button.get("style", "subtle")
        if style == "primary":
            button_style = primary_button_style()
        elif style == "accent":
            button_style = accent_button_style()
        elif style == "muted":
            button_style = muted_button_style()
        else:
            button_style = subtle_button_style()
        ctk.CTkButton(
            toolbar,
            text=button["text"],
            width=button.get("width", 82),
            height=36,
            command=button["command"],
            **button_style,
        ).grid(row=0, column=idx, sticky="e", padx=(0, 6) if idx < len(buttons or []) else 0)
    return toolbar


def set_status(label, text, kind="info"):
    label.configure(text=text, text_color=STATUS_COLORS.get(kind, STATUS_COLORS["info"]))


def set_check_item(registry, key, text, kind="info"):
    if not registry or key not in registry:
        return
    registry[key].configure(
        text=f"{STATUS_MARKS.get(kind, STATUS_MARKS['info'])} {text}",
        text_color=STATUS_COLORS.get(kind, STATUS_COLORS["info"]),
    )


def make_check_group(parent, row, title, items, registry, static_kind=None, wraplength=270):
    group = ctk.CTkFrame(
        parent,
        fg_color=CARD_SUBTLE_BG,
        corner_radius=8,
        border_width=1,
        border_color=BORDER_COLOR,
    )
    group.grid(row=row, column=0, sticky="ew", padx=12, pady=(0, 12))
    group.grid_columnconfigure(0, weight=1)
    ctk.CTkLabel(
        group,
        text=title,
        font=FONT_BODY_BOLD,
        text_color=TEXT_PRIMARY,
        anchor="w",
    ).grid(row=0, column=0, sticky="ew", padx=12, pady=(11, 5))
    for idx, (key, text) in enumerate(items, start=1):
        label = ctk.CTkLabel(
            group,
            text=text,
            font=FONT_SMALL,
            text_color=TEXT_MUTED,
            anchor="w",
            justify="left",
            wraplength=wraplength,
        )
        label.grid(row=idx, column=0, sticky="ew", padx=12, pady=(0, 8))
        registry[key] = label
        if static_kind is not None:
            set_check_item(registry, key, text, static_kind)
    return group


def make_check_panel(parent, title="生成前检查", width=318, scrollable=False):
    panel_class = ctk.CTkScrollableFrame if scrollable else ctk.CTkFrame
    panel = panel_class(
        parent,
        width=width,
        fg_color=CARD_BG,
        corner_radius=8,
        border_width=1,
        border_color=BORDER_COLOR,
    )
    panel.grid_columnconfigure(0, weight=1)
    ctk.CTkLabel(
        panel,
        text=title,
        font=("Microsoft YaHei UI", 20, "bold"),
        text_color=TEXT_PRIMARY,
        anchor="w",
    ).grid(row=0, column=0, sticky="ew", padx=18, pady=(18, 12))
    return panel


def table_cell_padx(col):
    """统一的表格列内边距：首列留左缩进，其余仅留右侧列间距，便于表头与数据行对齐。"""
    return (12 if col == 0 else 0, 10)


def add_compact_table_row(parent, row, values, weights=None, colors=None, bold_columns=None, height=38, widths=None):
    line = ctk.CTkFrame(parent, fg_color=CARD_BG if row % 2 == 0 else CARD_SUBTLE_BG, corner_radius=0)
    line.grid(row=row, column=0, sticky="ew", padx=0, pady=(0, 1))
    weights = weights or [0] * len(values)
    colors = colors or [TEXT_PRIMARY] * len(values)
    bold_columns = set(bold_columns or [])
    for col in range(len(values)):
        weight = weights[col] if col < len(weights) else 0
        if widths is not None:
            line.grid_columnconfigure(col, weight=weight, minsize=widths[col])
        else:
            line.grid_columnconfigure(col, weight=weight)
    if widths is not None:
        # 末尾留一个弹性占位列，吸收剩余宽度和滚动条偏移，使各定宽列与表头逐列对齐。
        line.grid_columnconfigure(len(values), weight=1)
    for col, value in enumerate(values):
        cell_kwargs = {}
        if widths is not None:
            cell_kwargs["width"] = widths[col]
            wrap = max(widths[col] - 4, 1)
        else:
            wrap = 360 if col == 1 else 160
        ctk.CTkLabel(
            line,
            text=value,
            font=FONT_BODY_BOLD if col in bold_columns else FONT_BODY,
            text_color=colors[col] if col < len(colors) else TEXT_PRIMARY,
            anchor="w",
            justify="left",
            wraplength=wrap,
            height=height,
            **cell_kwargs,
        ).grid(row=0, column=col, sticky="w", padx=table_cell_padx(col), pady=2)
    return line


def clear_children(frame):
    for child in frame.winfo_children():
        child.destroy()


def get_name_chip_style(selected=False):
    return {
        "fg_color": CHIP_SELECTED if selected else CHIP_FG,
        "corner_radius": 7,
        "border_width": 1,
        "border_color": CHIP_SELECTED_BORDER if selected else CHIP_BORDER,
    }


def get_name_button_style(selected=False):
    return {
        "fg_color": CHIP_SELECTED if selected else CHIP_FG,
        "hover_color": CHIP_HOVER,
        "text_color": CHIP_TEXT,
        "border_width": 1,
        "border_color": CHIP_SELECTED_BORDER if selected else CHIP_BORDER,
        "corner_radius": 7,
    }


def configure_name_tile(tile, selected=False):
    tile.configure(
        fg_color=CHIP_SELECTED if selected else CHIP_FG,
        border_color=CHIP_SELECTED_BORDER if selected else CHIP_BORDER,
    )


def add_name_chip(parent, text, row, column, width=112, selected=False):
    chip = ctk.CTkFrame(parent, **get_name_chip_style(selected=selected))
    chip.grid(row=row, column=column, sticky="ew", padx=5, pady=5)
    ctk.CTkLabel(
        chip,
        text=text,
        font=FONT_BODY,
        text_color=CHIP_TEXT,
        height=30,
        anchor="center",
    ).grid(row=0, column=0, sticky="ew", padx=10, pady=2)
    chip.grid_columnconfigure(0, weight=1, minsize=width)
    return chip


def add_record_row(parent, row, index, name, value_text, remove_command):
    card = ctk.CTkFrame(parent, **get_name_chip_style())
    card.grid(row=row, column=0, sticky="ew", padx=4, pady=3)
    card.grid_columnconfigure(1, weight=1)
    ctk.CTkLabel(
        card,
        text=str(index),
        width=28,
        font=FONT_SMALL,
        text_color=TEXT_MUTED,
        anchor="center",
    ).grid(row=0, column=0, sticky="w", padx=(8, 4), pady=5)
    ctk.CTkLabel(
        card,
        text=name,
        font=FONT_BODY,
        text_color=CHIP_TEXT,
        anchor="w",
    ).grid(row=0, column=1, sticky="ew", padx=(0, 8), pady=5)
    ctk.CTkLabel(
        card,
        text=value_text,
        font=FONT_SMALL,
        text_color=CHIP_TEXT,
        anchor="e",
    ).grid(row=0, column=2, sticky="e", padx=(0, 8), pady=5)
    ctk.CTkButton(
        card,
        text="移除",
        width=52,
        height=24,
        command=remove_command,
        **muted_button_style(),
    ).grid(row=0, column=3, sticky="e", padx=(0, 8), pady=5)
    return card


def render_name_chips(parent, names, columns=6, empty_text="暂无数据", width=112):
    clear_children(parent)
    for col in range(columns):
        parent.grid_columnconfigure(col, weight=1, uniform="name_chips")
    if not names:
        ctk.CTkLabel(
            parent,
            text=empty_text,
            font=FONT_BODY,
            text_color=CHIP_MUTED,
        ).grid(row=0, column=0, sticky="w", padx=8, pady=8)
        return
    for idx, name in enumerate(names):
        add_name_chip(parent, name, idx // columns, idx % columns, width=width)


def make_stat_cell(parent, column, label, value_text="-", min_width=140):
    cell = ctk.CTkFrame(parent, fg_color=CARD_SUBTLE_BG, corner_radius=8, border_width=1, border_color=BORDER_COLOR)
    cell.grid(row=0, column=column, sticky="ew", padx=5)
    cell.grid_columnconfigure(0, weight=1, minsize=min_width)
    ctk.CTkLabel(
        cell,
        text=label,
        font=FONT_SMALL,
        text_color=TEXT_MUTED,
        anchor="w",
    ).grid(row=0, column=0, sticky="w", padx=12, pady=(9, 0))
    value = ctk.CTkLabel(
        cell,
        text=value_text,
        font=FONT_BODY_BOLD,
        text_color=TEXT_PRIMARY,
        anchor="w",
    )
    value.grid(row=1, column=0, sticky="w", padx=12, pady=(1, 10))
    return value
