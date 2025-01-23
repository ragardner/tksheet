from __future__ import annotations

import tkinter as tk
from collections.abc import (
    Callable,
)
from typing import Literal

from .other_classes import (
    DotDict,
)
from .constants import (
    ctrl_key,
    rc_binding,
)


class FindWindowTkText(tk.Text):
    def __init__(
        self,
        parent: tk.Misc,
    ) -> None:
        super().__init__(
            parent,
            spacing1=0,
            spacing2=1,
            spacing3=0,
            bd=0,
            highlightthickness=0,
            undo=True,
            maxundo=30,
        )
        self.parent = parent
        self.rc_popup_menu = tk.Menu(self, tearoff=0)
        self.bind("<1>", lambda event: self.focus_set())
        self.bind(rc_binding, self.rc)
        self.bind(f"<{ctrl_key}-a>", self.select_all)
        self.bind(f"<{ctrl_key}-A>", self.select_all)
        self.bind("<Delete>", self.delete_key)

    def reset(
        self,
        menu_kwargs: dict,
        sheet_ops: dict,
        font: tuple,
        bg: str,
        fg: str,
        select_bg: str,
        select_fg: str,
    ) -> None:
        self.config(
            font=font,
            background=bg,
            foreground=fg,
            insertbackground=fg,
            selectbackground=select_bg,
            selectforeground=select_fg,
        )
        self.editor_del_key = sheet_ops.editor_del_key
        self.rc_popup_menu.delete(0, "end")
        self.rc_popup_menu.add_command(
            label=sheet_ops.select_all_label,
            accelerator=sheet_ops.select_all_accelerator,
            command=self.select_all,
            **menu_kwargs,
        )
        self.rc_popup_menu.add_command(
            label=sheet_ops.cut_label,
            accelerator=sheet_ops.cut_accelerator,
            command=self.cut,
            **menu_kwargs,
        )
        self.rc_popup_menu.add_command(
            label=sheet_ops.copy_label,
            accelerator=sheet_ops.copy_accelerator,
            command=self.copy,
            **menu_kwargs,
        )
        self.rc_popup_menu.add_command(
            label=sheet_ops.paste_label,
            accelerator=sheet_ops.paste_accelerator,
            command=self.paste,
            **menu_kwargs,
        )
        self.rc_popup_menu.add_command(
            label=sheet_ops.undo_label,
            accelerator=sheet_ops.undo_accelerator,
            command=self.undo,
            **menu_kwargs,
        )

    def rc(self, event: object) -> None:
        self.focus_set()
        self.rc_popup_menu.tk_popup(event.x_root, event.y_root)

    def delete_key(self, event: object = None) -> None:
        if self.editor_del_key == "forward":
            return
        elif not self.editor_del_key:
            return "break"
        elif self.editor_del_key == "backward":
            if self.tag_ranges("sel"):
                return
            if self.index("insert") == "1.0":
                return "break"
            self.delete("insert-1c")
            return "break"

    def select_all(self, event: object = None) -> Literal["break"]:
        self.tag_add(tk.SEL, "1.0", tk.END)
        self.mark_set(tk.INSERT, tk.END)
        # self.see(tk.INSERT)
        return "break"

    def cut(self, event: object = None) -> Literal["break"]:
        self.event_generate(f"<{ctrl_key}-x>")
        self.event_generate("<KeyRelease>")
        return "break"

    def copy(self, event: object = None) -> Literal["break"]:
        self.event_generate(f"<{ctrl_key}-c>")
        return "break"

    def paste(self, event: object = None) -> Literal["break"]:
        self.event_generate(f"<{ctrl_key}-v>")
        self.event_generate("<KeyRelease>")
        return "break"

    def undo(self, event: object = None) -> Literal["break"]:
        self.event_generate(f"<{ctrl_key}-z>")
        self.event_generate("<KeyRelease>")
        return "break"


class FindWindow(tk.Frame):
    def __init__(
        self,
        parent: tk.Misc,
        find_next_func: Callable,
        find_prev_func: Callable,
        close_func: Callable,
    ) -> None:
        super().__init__(
            parent,
            width=0,
            height=0,
            bd=0,
        )
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.grid_propagate(False)
        self.parent = parent
        self.tktext = FindWindowTkText(self)
        self.tktext.grid(row=0, column=0, sticky="nswe")
        self.bg = None
        self.fg = None

        self.find_previous_arrow = tk.Label(self, text="▲", cursor="hand2", highlightthickness=1)
        self.find_previous_arrow.bind("<Button-1>", find_prev_func)
        self.find_previous_arrow.grid(row=0, column=1)

        self.find_next_arrow = tk.Label(self, text="▼", cursor="hand2", highlightthickness=1)
        self.find_next_arrow.bind("<Button-1>", find_next_func)
        self.find_next_arrow.grid(row=0, column=2)

        self.find_in_selection = False
        self.in_selection = tk.Label(self, text="🔎", cursor="hand2", highlightthickness=1)
        self.in_selection.bind("<Button-1>", self.toggle_in_selection)
        self.in_selection.grid(row=0, column=3)

        self.close = tk.Label(self, text="✕", cursor="hand2", highlightthickness=1)
        self.close.bind("<Button-1>", close_func)
        self.close.grid(row=0, column=4)

        for widget in (self.find_previous_arrow, self.find_next_arrow, self.in_selection, self.close):
            widget.bind("<Enter>", lambda w, widget=widget: self.enter_label(widget=widget))
            widget.bind("<Leave>", lambda w, widget=widget: self.leave_label(widget=widget))

    def enter_label(self, widget: tk.Misc) -> None:
        widget.config(
            highlightbackground=self.fg,
            highlightcolor=self.fg,
        )

    def leave_label(self, widget: tk.Misc) -> None:
        if widget == self.in_selection and self.find_in_selection:
            return
        widget.config(
            highlightbackground=self.bg,
            highlightcolor=self.fg,
        )

    def toggle_in_selection(self, event: tk.Misc) -> None:
        self.find_in_selection = not self.find_in_selection
        self.enter_label(self.in_selection)

    def get(self) -> str:
        return self.tktext.get("1.0", "end-1c")

    def get_num_lines(self) -> int:
        return int(self.tktext.index("end-1c").split(".")[0])

    def set_text(self, text: str = "") -> None:
        self.tktext.delete(1.0, "end")
        self.tktext.insert(1.0, text)

    def reset(
        self,
        border_color: str,
        menu_kwargs: DotDict,
        sheet_ops: DotDict,
        bg: str,
        fg: str,
        select_bg: str,
        select_fg: str,
    ) -> None:
        self.bg = bg
        self.fg = fg
        self.tktext.reset(
            menu_kwargs=menu_kwargs,
            sheet_ops=sheet_ops,
            font=menu_kwargs.font,
            bg=bg,
            fg=fg,
            select_bg=select_bg,
            select_fg=select_fg,
        )
        for widget in (self.find_previous_arrow, self.find_next_arrow, self.in_selection, self.close):
            widget.config(
                font=menu_kwargs.font,
                bg=bg,
                fg=fg,
                highlightbackground=bg,
                highlightcolor=fg,
            )
        if self.find_in_selection:
            self.enter_label(self.in_selection)
        self.config(
            background=bg,
            highlightbackground=border_color,
            highlightcolor=border_color,
            highlightthickness=1,
        )
