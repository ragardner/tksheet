from __future__ import annotations

import tkinter as tk
from collections.abc import Callable
from typing import Literal

from .functions import (
    convert_align,
)
from .other_classes import (
    DotDict,
)
from .vars import (
    ctrl_key,
    rc_binding,
)


class TextEditorTkText(tk.Text):
    def __init__(
        self,
        parent: tk.Misc,
        newline_binding: None | Callable = None,
    ) -> None:
        tk.Text.__init__(
            self,
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
        self.newline_bindng = newline_binding
        self.rc_popup_menu = tk.Menu(self, tearoff=0)
        self.bind("<1>", lambda event: self.focus_set())
        self.bind(rc_binding, self.rc)
        self.bind(f"<{ctrl_key}-a>", self.select_all)
        self.bind(f"<{ctrl_key}-A>", self.select_all)
        self._orig = self._w + "_orig"
        self.tk.call("rename", self._w, self._orig)
        self.tk.createcommand(self._w, self._proxy)

    def reset(
        self,
        menu_kwargs: dict,
        sheet_ops: dict,
        align: str,
        font: tuple,
        bg: str,
        fg: str,
        state: str,
        text: str = "",
    ) -> None:
        self.config(
            font=font,
            background=bg,
            foreground=fg,
            insertbackground=fg,
            state=state,
        )
        self.align = align
        self.rc_popup_menu.delete(0, None)
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
        align = convert_align(align)
        if align == "w":
            self.align = "left"
        elif align == "e":
            self.align = "right"
        self.delete(1.0, "end")
        self.insert(1.0, text)
        self.yview_moveto(1)
        self.tag_configure("align", justify=self.align)
        self.tag_add("align", 1.0, "end")

    def _proxy(self, command: object, *args) -> object:
        try:
            result = self.tk.call((self._orig, command) + args)
        except Exception:
            return
        if command in (
            "insert",
            "delete",
            "replace",
        ):
            self.tag_add("align", 1.0, "end")
            self.event_generate("<<TextModified>>")
            if args and len(args) > 1 and args[1] != "\n" and args != ("1.0", "end"):
                if self.yview() != (0.0, 1.0) and self.newline_bindng is not None:
                    self.newline_bindng(check_lines=False)
        return result

    def rc(self, event: object) -> None:
        self.focus_set()
        self.rc_popup_menu.tk_popup(event.x_root, event.y_root)

    def select_all(self, event: object = None) -> Literal["break"]:
        self.tag_add(tk.SEL, "1.0", tk.END)
        self.mark_set(tk.INSERT, tk.END)
        # self.see(tk.INSERT)
        return "break"

    def cut(self, event: object = None) -> Literal["break"]:
        self.event_generate(f"<{ctrl_key}-x>")
        return "break"

    def copy(self, event: object = None) -> Literal["break"]:
        self.event_generate(f"<{ctrl_key}-c>")
        return "break"

    def paste(self, event: object = None) -> Literal["break"]:
        self.event_generate(f"<{ctrl_key}-v>")
        return "break"

    def undo(self, event: object = None) -> Literal["break"]:
        self.event_generate(f"<{ctrl_key}-z>")
        return "break"


class TextEditor(tk.Frame):
    def __init__(
        self,
        parent: tk.Misc,
        newline_binding: None | Callable = None,
    ) -> None:
        tk.Frame.__init__(
            self,
            parent,
            width=0,
            height=0,
            bd=0,
        )
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.grid_propagate(False)
        self.parent = parent
        self.r = 0
        self.c = 0
        self.tktext = TextEditorTkText(self, newline_binding=newline_binding)
        self.tktext.grid(row=0, column=0, sticky="nswe")

    def get(self) -> str:
        return self.tktext.get("1.0", "end-1c")

    def get_num_lines(self) -> int:
        return int(self.tktext.index("end-1c").split(".")[0])

    def set_text(self, text: str = "") -> None:
        self.tktext.delete(1.0, "end")
        self.tktext.insert(1.0, text)

    def scroll_to_bottom(self) -> None:
        self.tktext.yview_moveto(1)

    def reset(
        self,
        width: int,
        height: int,
        border_color: str,
        show_border: bool,
        menu_kwargs: DotDict,
        sheet_ops: DotDict,
        bg: str,
        fg: str,
        align: str,
        state: str,
        r: int = 0,
        c: int = 0,
        text: str = "",
    ) -> None:
        self.r = r
        self.c = c
        self.tktext.reset(
            menu_kwargs=menu_kwargs,
            sheet_ops=sheet_ops,
            align=align,
            font=menu_kwargs.font,
            bg=bg,
            fg=fg,
            state=state,
            text=text,
        )
        self.config(
            width=width,
            height=height,
            background=bg,
            highlightbackground=border_color,
            highlightcolor=border_color,
            highlightthickness=2 if show_border else 0,
        )
