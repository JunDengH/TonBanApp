"""Reusable card-style name picker for large assistant lists."""
import tkinter as tk

import customtkinter as ctk

from src.ui.ui_helpers import (
    BORDER_COLOR,
    CARD_SUBTLE_BG,
    CHIP_BORDER,
    CHIP_FG,
    CHIP_HOVER,
    CHIP_SELECTED,
    CHIP_SELECTED_BORDER,
    CHIP_TEXT,
    FONT_BODY,
    FONT_SMALL,
    PANEL_BG,
    PRIMARY,
    TEXT_MUTED,
    TEXT_ON_DARK,
    bind_focus_border,
    input_style,
    resolve_color,
)


class NamePickerPanel(ctk.CTkFrame):
    def __init__(
        self,
        parent,
        mode="multiple",
        height=220,
        placeholder="搜索姓名",
        on_selection_change=None,
        on_filter_change=None,
    ):
        super().__init__(
            parent,
            fg_color=CARD_SUBTLE_BG,
            corner_radius=8,
            border_width=1,
            border_color=BORDER_COLOR,
        )
        self.mode = mode
        self.on_selection_change = on_selection_change
        self.on_filter_change = on_filter_change
        self.all_names = []
        self.visible_names = []
        self.selected_names = set()
        self.selected_name = ""
        self.hovered_name = ""
        self.card_boxes = []
        self.columns = 1 if mode == "single" else 2
        self.card_height = 36
        self.card_gap = 8
        self.card_padding = 8

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        self.search_var = ctk.StringVar(value="")
        self.search_var.trace_add("write", lambda *_: self._render())
        search_entry = ctk.CTkEntry(
            self,
            textvariable=self.search_var,
            placeholder_text=placeholder,
            height=34,
            **input_style(),
        )
        bind_focus_border(search_entry)
        search_entry.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 8))

        list_shell = ctk.CTkFrame(
            self,
            fg_color=PANEL_BG,
            corner_radius=8,
            border_width=1,
            border_color=BORDER_COLOR,
        )
        list_shell.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        list_shell.grid_columnconfigure(0, weight=1)
        list_shell.grid_rowconfigure(0, weight=1)

        self.canvas = tk.Canvas(
            list_shell,
            height=height,
            bg=resolve_color(PANEL_BG),
            bd=0,
            highlightthickness=0,
            cursor="hand2",
        )
        self.canvas.grid(row=0, column=0, sticky="nsew", padx=(6, 2), pady=6)
        self.canvas.bind("<Button-1>", self._on_canvas_click)
        self.canvas.bind("<Configure>", lambda _event: self._draw_cards())
        self.canvas.bind("<Leave>", self._on_canvas_leave)
        self.canvas.bind("<Motion>", self._on_canvas_motion)
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)

        self.scrollbar = ctk.CTkScrollbar(
            list_shell,
            command=self.canvas.yview,
            button_color=BORDER_COLOR,
            button_hover_color=PRIMARY,
        )
        self.scrollbar.grid(row=0, column=1, sticky="ns", padx=(0, 6), pady=6)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.hint_label = ctk.CTkLabel(self, text="", font=FONT_SMALL, text_color=TEXT_MUTED, anchor="w")
        self.hint_label.grid(row=2, column=0, sticky="ew", padx=10, pady=(0, 9))

    def set_names(self, names, selected_names=None):
        self.all_names = list(names or [])
        if selected_names is not None:
            selected = [name for name in selected_names if name in self.all_names]
            if self.mode == "single":
                self.selected_name = selected[0] if selected else ""
            else:
                self.selected_names = set(selected)
        self._render()

    def get_selected_names(self):
        if self.mode == "single":
            return [self.selected_name] if self.selected_name else []
        return [name for name in self.all_names if name in self.selected_names]

    def get_visible_names(self):
        return list(self.visible_names)

    def select_all_visible(self):
        if self.mode == "single":
            return
        self.selected_names.update(self.visible_names)
        self._draw_cards()
        self._notify_selection_change()

    def clear_visible(self):
        if self.mode == "single":
            self.selected_name = ""
        else:
            self.selected_names.difference_update(self.visible_names)
        self._draw_cards()
        self._notify_selection_change()

    def _render(self):
        if not hasattr(self, "canvas"):
            return
        keyword = self.search_var.get().strip()
        self.visible_names = [name for name in self.all_names if not keyword or keyword in name]
        self.hovered_name = ""
        self._draw_cards()
        self.hint_label.configure(text=f"显示 {len(self.visible_names)} / {len(self.all_names)} 人")
        if self.on_filter_change is not None:
            self.on_filter_change()

    def _draw_cards(self):
        if not hasattr(self, "canvas"):
            return
        self.canvas.delete("all")
        self.card_boxes = []
        width = max(self.canvas.winfo_width(), 220)

        if not self.all_names:
            self._draw_empty_text("请先导入名单", width)
            return
        if not self.visible_names:
            self._draw_empty_text("没有匹配的姓名", width)
            return

        card_width = self._card_width(width)
        for index, name in enumerate(self.visible_names):
            col = index % self.columns
            row = index // self.columns
            x1 = self.card_padding + col * (card_width + self.card_gap)
            y1 = self.card_padding + row * (self.card_height + self.card_gap)
            x2 = x1 + card_width
            y2 = y1 + self.card_height
            self._draw_name_card(name, x1, y1, x2, y2)
            self.card_boxes.append((name, x1, y1, x2, y2))

        row_count = (len(self.visible_names) + self.columns - 1) // self.columns
        content_height = self.card_padding * 2 + row_count * self.card_height + max(0, row_count - 1) * self.card_gap
        self.canvas.configure(scrollregion=(0, 0, width, max(content_height, self.canvas.winfo_height())))

    def _set_appearance_mode(self, mode_string):
        super()._set_appearance_mode(mode_string)
        if hasattr(self, "canvas"):
            self.canvas.configure(bg=resolve_color(PANEL_BG))
            self._draw_cards()

    def _draw_empty_text(self, text, width):
        self.canvas.create_text(
            12,
            14,
            text=text,
            fill=resolve_color(TEXT_MUTED),
            font=FONT_SMALL,
            anchor="nw",
        )
        self.canvas.configure(scrollregion=(0, 0, width, max(self.canvas.winfo_height(), 1)))

    def _card_width(self, width):
        total_gap = self.card_gap * (self.columns - 1)
        return max(80, (width - self.card_padding * 2 - total_gap) / self.columns)

    def _draw_name_card(self, name, x1, y1, x2, y2):
        selected = self._is_selected(name)
        hovered = name == self.hovered_name
        fill = CHIP_SELECTED if selected else (CHIP_HOVER if hovered else CHIP_FG)
        outline = CHIP_SELECTED_BORDER if selected else CHIP_BORDER
        self._rounded_rect(
            x1, y1, x2, y2, 8,
            fill=resolve_color(fill),
            outline=resolve_color(outline),
        )

        box_x = x1 + 12
        box_y = y1 + 10
        box_size = 16
        self._rounded_rect(
            box_x,
            box_y,
            box_x + box_size,
            box_y + box_size,
            4,
            fill=resolve_color(PRIMARY if selected else PANEL_BG),
            outline=resolve_color(PRIMARY if selected else BORDER_COLOR),
        )
        if selected:
            self.canvas.create_text(
                box_x + box_size / 2,
                box_y + box_size / 2 - 1,
                text="✓",
                # Theme-aware: PRIMARY box is near-white in dark mode, so a
                # literal white check would vanish. TEXT_ON_DARK flips per theme.
                fill=resolve_color(TEXT_ON_DARK),
                font=(FONT_BODY[0], 10, "bold"),
            )
        self.canvas.create_text(
            box_x + box_size + 10,
            y1 + self.card_height / 2,
            text=name,
            fill=resolve_color(CHIP_TEXT),
            font=FONT_BODY,
            anchor="w",
        )

    def _rounded_rect(self, x1, y1, x2, y2, radius, **kwargs):
        points = [
            x1 + radius,
            y1,
            x2 - radius,
            y1,
            x2,
            y1,
            x2,
            y1 + radius,
            x2,
            y2 - radius,
            x2,
            y2,
            x2 - radius,
            y2,
            x1 + radius,
            y2,
            x1,
            y2,
            x1,
            y2 - radius,
            x1,
            y1 + radius,
            x1,
            y1,
        ]
        return self.canvas.create_polygon(points, smooth=True, splinesteps=12, **kwargs)

    def _is_selected(self, name):
        if self.mode == "single":
            return name == self.selected_name
        return name in self.selected_names

    def _on_canvas_click(self, event):
        name = self._name_at(event.x, event.y)
        if not name:
            return
        if self.mode == "single":
            self.selected_name = name
        elif name in self.selected_names:
            self.selected_names.remove(name)
        else:
            self.selected_names.add(name)
        self._draw_cards()
        self._notify_selection_change()

    def _on_canvas_motion(self, event):
        name = self._name_at(event.x, event.y)
        if name == self.hovered_name:
            return
        self.hovered_name = name
        self._draw_cards()

    def _on_canvas_leave(self, _event):
        if self.hovered_name:
            self.hovered_name = ""
            self._draw_cards()

    def _name_at(self, x, y):
        canvas_x = self.canvas.canvasx(x)
        canvas_y = self.canvas.canvasy(y)
        for name, x1, y1, x2, y2 in self.card_boxes:
            if x1 <= canvas_x <= x2 and y1 <= canvas_y <= y2:
                return name
        return ""

    def _on_mousewheel(self, event):
        steps = -1 * int(event.delta / 120) if event.delta else 0
        if steps:
            self.canvas.yview_scroll(steps, "units")
        return "break"

    def _notify_selection_change(self):
        if self.on_selection_change is not None:
            self.on_selection_change(self.get_selected_names())
