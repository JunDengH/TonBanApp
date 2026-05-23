"""UI shared constants and small layout helpers."""
from pathlib import Path

import customtkinter as ctk


ROOT_DIR = Path(__file__).resolve().parent.parent.parent
OUTPUT_DIR = ROOT_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

FONT_TITLE = ("Microsoft YaHei", 20, "bold")
FONT_SECTION = ("Microsoft YaHei", 15, "bold")
FONT_BODY = ("Microsoft YaHei", 13)
FONT_SMALL = ("Microsoft YaHei", 11)

CHIP_FG = ("#eef3f8", "#24313d")
CHIP_HOVER = ("#dfeaf5", "#2d3c4b")
CHIP_SELECTED = ("#d8ebff", "#17476f")
CHIP_TEXT = ("#1d2a35", "#f3f7fb")
CHIP_MUTED = ("#56616b", "#aab6c2")


def make_tab_page(tab, title):
    tab.grid_columnconfigure(0, weight=1)
    tab.grid_rowconfigure(1, weight=1)
    ctk.CTkLabel(tab, text=title, font=FONT_TITLE).grid(
        row=0, column=0, sticky="w", padx=18, pady=(10, 6)
    )
    page = ctk.CTkFrame(tab, fg_color="transparent")
    page.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 8))
    page.grid_columnconfigure(0, weight=1)
    return page


def make_section(parent, title, subtitle=None):
    frame = ctk.CTkFrame(parent, corner_radius=10)
    frame.grid_columnconfigure(0, weight=1)
    ctk.CTkLabel(frame, text=title, font=FONT_SECTION).grid(
        row=0, column=0, sticky="w", padx=16, pady=(10, 2)
    )
    start_row = 1
    if subtitle:
        ctk.CTkLabel(
            frame,
            text=subtitle,
            font=FONT_SMALL,
            text_color=("gray35", "gray70"),
            wraplength=820,
            justify="left",
        ).grid(row=1, column=0, sticky="w", padx=16, pady=(0, 8))
        start_row = 2
    return frame, start_row


def add_file_row(parent, row, label, variable, placeholder, command, button_text="选择"):
    line = ctk.CTkFrame(parent, fg_color="transparent")
    line.grid(row=row, column=0, sticky="ew", padx=16, pady=3)
    line.grid_columnconfigure(1, weight=1)
    ctk.CTkLabel(line, text=label, width=80, anchor="w", font=FONT_BODY).grid(
        row=0, column=0, sticky="w", padx=(0, 8)
    )
    ctk.CTkEntry(line, textvariable=variable, placeholder_text=placeholder).grid(
        row=0, column=1, sticky="ew", padx=(0, 8)
    )
    ctk.CTkButton(line, text=button_text, width=84, command=command).grid(
        row=0, column=2, sticky="e"
    )


def set_status(label, text, kind="info"):
    colors = {
        "info": ("gray35", "gray70"),
        "success": ("#16784f", "#3bc985"),
        "error": ("#9f1d1d", "#ff8a8a"),
    }
    label.configure(text=text, text_color=colors.get(kind, colors["info"]))


def muted_button_style():
    return {
        "fg_color": ("gray75", "gray28"),
        "hover_color": ("gray68", "gray34"),
        "text_color": ("gray10", "gray95"),
    }


def clear_children(frame):
    for child in frame.winfo_children():
        child.destroy()


def add_name_chip(parent, text, row, column, width=112):
    chip = ctk.CTkFrame(parent, fg_color=CHIP_FG, corner_radius=7)
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
