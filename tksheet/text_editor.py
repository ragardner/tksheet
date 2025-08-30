from __future__ import annotations

import tkinter as tk
from collections.abc import Callable
from typing import Any, Literal

from .constants import align_helper, ctrl_key
from .functions import convert_align
from .other_classes import DotDict, FontTuple


class TextEditorTkText(tk.Text):
    def __init__(
        self,
        parent: tk.Misc,
        newline_binding: None | Callable = None,
        rc_bindings: list[str] = "<3>",
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
        self.newline_bindng = newline_binding
        self.rc_popup_menu = tk.Menu(self, tearoff=0)
        self.bind("<1>", lambda event: self.focus_set())
        for b in rc_bindings:
            self.bind(b, self.rc)
        self.bind(f"<{ctrl_key}-a>", self.select_all)
        self.bind(f"<{ctrl_key}-A>", self.select_all)
        self.bind("<Delete>", self.delete_key)
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
        select_bg: str,
        select_fg: str,
        state: str,
        text: str = "",
    ) -> None:
        self.config(
            font=font,
            background=bg,
            foreground=fg,
            insertbackground=fg,
            state=state,
            selectbackground=select_bg,
            selectforeground=select_fg,
        )
        self.editor_del_key = sheet_ops.editor_del_key
        self.align = align
        self.rc_popup_menu.delete(0, "end")
        self.rc_popup_menu.add_command(
            label=sheet_ops.select_all_label,
            accelerator=sheet_ops.select_all_accelerator,
            image=sheet_ops.select_all_image,
            compound=sheet_ops.select_all_compound,
            command=self.select_all,
            **menu_kwargs,
        )
        self.rc_popup_menu.add_command(
            label=sheet_ops.cut_label,
            accelerator=sheet_ops.cut_accelerator,
            image=sheet_ops.cut_image,
            compound=sheet_ops.cut_compound,
            command=self.cut,
            **menu_kwargs,
        )
        self.rc_popup_menu.add_command(
            label=sheet_ops.copy_label,
            accelerator=sheet_ops.copy_accelerator,
            image=sheet_ops.copy_image,
            compound=sheet_ops.copy_compound,
            command=self.copy,
            **menu_kwargs,
        )
        self.rc_popup_menu.add_command(
            label=sheet_ops.paste_label,
            accelerator=sheet_ops.paste_accelerator,
            image=sheet_ops.paste_image,
            compound=sheet_ops.paste_compound,
            command=self.paste,
            **menu_kwargs,
        )
        self.rc_popup_menu.add_command(
            label=sheet_ops.undo_label,
            accelerator=sheet_ops.undo_accelerator,
            image=sheet_ops.undo_image,
            compound=sheet_ops.undo_compound,
            command=self.undo,
            **menu_kwargs,
        )
        self.rc_popup_menu.add_command(
            label=sheet_ops.redo_label,
            accelerator=sheet_ops.redo_accelerator,
            image=sheet_ops.redo_image,
            compound=sheet_ops.redo_compound,
            command=self.redo,
            **menu_kwargs,
        )
        self.align = align_helper[convert_align(align)]
        self.delete(1.0, "end")
        self.insert(1.0, text)
        self.yview_moveto(1)
        self.tag_configure("align", justify=self.align)
        self.tag_add("align", 1.0, "end")
        self.edit_reset()

    def _proxy(self, command: Any, *args) -> Any:
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
            if (
                args
                and len(args) > 1
                and args[1] != "\n"
                and args != ("1.0", "end")
                and self.yview() != (0.0, 1.0)
                and self.newline_bindng is not None
            ):
                self.newline_bindng(check_lines=False)
        return result

    def rc(self, event: Any) -> None:
        self.focus_set()
        self.rc_popup_menu.tk_popup(event.x_root, event.y_root)

    def delete_key(self, event: Any = None) -> None:
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

    def select_all(self, event: Any = None) -> Literal["break"]:
        self.tag_add(tk.SEL, "1.0", "end-1c")
        self.mark_set(tk.INSERT, "end-1c")
        # self.see(tk.INSERT)
        return "break"

    def cut(self, event: Any = None) -> Literal["break"]:
        self.event_generate(f"<{ctrl_key}-x>")
        self.event_generate("<KeyRelease>")
        return "break"

    def copy(self, event: Any = None) -> Literal["break"]:
        self.event_generate(f"<{ctrl_key}-c>")
        return "break"

    def paste(self, event: Any = None) -> Literal["break"]:
        self.event_generate(f"<{ctrl_key}-v>")
        self.event_generate("<KeyRelease>")
        return "break"

    def undo(self, event: Any = None) -> Literal["break"]:
        self.event_generate(f"<{ctrl_key}-z>")
        self.event_generate("<KeyRelease>")
        return "break"

    def redo(self, event: Any = None) -> Literal["break"]:
        self.event_generate(f"<{ctrl_key}-Shift-z>")
        self.event_generate("<KeyRelease>")
        return "break"


class TextEditor(tk.Frame):
    def __init__(
        self,
        parent: tk.Misc,
        newline_binding: None | Callable = None,
        rc_bindings: list[str] = "<3>",
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
        self.r = 0
        self.c = 0
        self.tktext = TextEditorTkText(self, newline_binding=newline_binding, rc_bindings=rc_bindings)
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
        font: FontTuple,
        bg: str,
        fg: str,
        select_bg: str,
        select_fg: str,
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
            font=font,
            bg=bg,
            fg=fg,
            select_bg=select_bg,
            select_fg=select_fg,
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
