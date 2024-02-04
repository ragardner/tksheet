from __future__ import annotations

import tkinter as tk
from collections.abc import Callable
from typing import Literal

from .other_classes import (
    DotDict,
)
from .vars import (
    ctrl_key,
    rc_binding,
)


class TextEditor_(tk.Text):
    def __init__(
        self,
        parent: tk.Misc,
        menu_kwargs: DotDict,
        sheet_ops: DotDict,
        text: None | str = None,
        state: str = "normal",
        bg: str = "white",
        fg: str = "black",
        align: str = "w",
        newline_binding: None | Callable = None,
    ) -> None:
        tk.Text.__init__(
            self,
            parent,
            font=menu_kwargs.font,
            state=state,
            spacing1=0,
            spacing2=0,
            spacing3=0,
            bd=0,
            highlightthickness=0,
            undo=True,
            maxundo=30,
            background=bg,
            foreground=fg,
            insertbackground=fg,
        )
        self.parent = parent
        self.newline_bindng = newline_binding
        if align == "w":
            self.align = "left"
        elif align == "center":
            self.align = "center"
        elif align == "e":
            self.align = "right"
        self.tag_configure("align", justify=self.align)
        if text:
            self.insert(1.0, text)
            self.yview_moveto(1)
        self.tag_add("align", 1.0, "end")
        self.rc_popup_menu = tk.Menu(self, tearoff=0)
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
        self.bind("<1>", lambda event: self.focus_set())
        self.bind(rc_binding, self.rc)
        self.bind(f"<{ctrl_key}-a>", self.select_all)
        self.bind(f"<{ctrl_key}-A>", self.select_all)
        self._orig = self._w + "_orig"
        self.tk.call("rename", self._w, self._orig)
        self.tk.createcommand(self._w, self._proxy)

    def _proxy(self, command: object, *args) -> object:
        cmd = (self._orig, command) + args
        try:
            result = self.tk.call(cmd)
        except Exception:
            return
        if command in (
            "insert",
            "delete",
            "replace",
        ):
            self.tag_add("align", 1.0, "end")
            self.event_generate("<<TextModified>>")
            if args and len(args) > 1 and args[1] != "\n":
                out_of_bounds = self.yview()
                if out_of_bounds != (0.0, 1.0) and self.newline_bindng is not None:
                    self.newline_bindng(
                        r=self.parent.r,
                        c=self.parent.c,
                        check_lines=False,
                    )
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
        menu_kwargs: DotDict,
        sheet_ops: DotDict,
        border_color: str,
        text: None | str = None,
        state: str = "normal",
        width: None | int = None,
        height: None | int = None,
        show_border: bool = True,
        bg: str = "white",
        fg: str = "black",
        align: str = "w",
        r: int = 0,
        c: int = 0,
        newline_binding: None | Callable = None,
    ) -> None:
        tk.Frame.__init__(
            self,
            parent,
            height=height,
            width=width,
            highlightbackground=border_color,
            highlightcolor=border_color,
            highlightthickness=2 if show_border else 0,
            bd=0,
        )
        self.parent = parent
        self.r = r
        self.c = c
        self.textedit = TextEditor_(
            self,
            menu_kwargs=menu_kwargs,
            sheet_ops=sheet_ops,
            text=text,
            state=state,
            bg=bg,
            fg=fg,
            align=align,
            newline_binding=newline_binding,
        )
        self.textedit.grid(row=0, column=0, sticky="nswe")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.grid_propagate(False)
        self.w_ = width
        self.h_ = height
        self.textedit.focus_set()

    def get(self) -> str:
        return self.textedit.get("1.0", "end-1c")

    def get_num_lines(self) -> int:
        return int(self.textedit.index("end-1c").split(".")[0])

    def set_text(self, text) -> None:
        self.textedit.delete(1.0, "end")
        self.textedit.insert(1.0, text)

    def scroll_to_bottom(self) -> None:
        self.textedit.yview_moveto(1)
