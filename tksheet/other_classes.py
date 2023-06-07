from __future__ import annotations

import tkinter as tk
from collections import namedtuple
from collections.abc import Iterator

from .vars import (
    ctrl_key,
    get_font,
    rc_binding,
)

CurrentlySelectedClass = namedtuple(
    "CurrentlySelectedClass",
    "row column type_ tags",
)
DrawnItem = namedtuple("DrawnItem", "iid showing")
TextCfg = namedtuple("TextCfg", "txt tf font align")
DraggedRowColumn = namedtuple("DraggedRowColumn", "dragged to_move")
ProgressBar = namedtuple("ProgressBar", "bg fg pc name")


class CanUseKeys:
    def __init__(self, **kwargs) -> None:
        self.__dict__.update(kwargs)

    def __bool__(self) -> bool:
        if any(value for value in self.__dict__.values()):
            return True
        return False

    def __getitem__(self, key: str) -> object:
        if isinstance(key, str):
            return getattr(self, key)
        else:
            raise ValueError(f"Key must be type 'str' not '{type(key)}'.")


class SheetEvent(CanUseKeys):
    def __init__(
        self,
        name: str = None,
        sheet: object = None,
        boxes: None | dict | tuple = None,
        cells_table: None | dict = None,
        cells_header: None | dict = None,
        cells_index: None | dict = None,
        selected: None | tuple = None,
        data: object = None,
        key: None | str = None,
        value: object = None,
        location: None | int | tuple[int] = None,
        resized_rows: None | dict = None,
        resized_columns: None | dict = None,
        # resized_index: None, dict] = None,
        # resized_header: None, dict] = None,
        being_selected: None | tuple = None,
        named_spans: None | dict = None,
        **kwargs,
    ):
        self.eventname = "" if name is None else name
        self.sheetname = "!sheet" if sheet is None else sheet
        self.cells = CanUseKeys(
            table={} if cells_table is None else cells_table,
            header={} if cells_header is None else cells_header,
            index={} if cells_index is None else cells_index,
        )
        self.moved = CanUseKeys(
            rows=CanUseKeys(
                data={},
                displayed={},
            ),
            columns=CanUseKeys(
                data={},
                displayed={},
            ),
        )
        self.added = CanUseKeys(
            rows=CanUseKeys(table={},
                            header={},
                            row_heights={},
                            displayed_rows={},),
            columns=CanUseKeys(table={},
                               header={},
                               column_widths={},
                               displayed_columns={},),
        )
        self.deleted = CanUseKeys(
            rows={},
            columns={},
            header={},
            index={},
            column_widths={},
            row_heights={},
            options={},
            displayed_rows=None,
            displayed_columns=None,
        )
        self.named_spans = {} if named_spans is None else named_spans
        self.selection_boxes = {} if boxes is None else {boxes[:-1]: boxes[-1]} if isinstance(boxes, tuple) else boxes
        self.selected = tuple() if selected is None else selected
        self.being_selected = tuple() if being_selected is None else being_selected
        self.data = [] if data is None else data
        self.key = "" if key is None else key
        self.value = None if value is None else value
        self.location = tuple() if location is None else location
        self.resized = CanUseKeys(
            rows={} if resized_rows is None else resized_rows,
            columns={} if resized_columns is None else resized_columns,
        )


class Span(CanUseKeys):
    def __init__(
        self,
        cells: tuple[int, int] | Iterator[tuple[int, int]] | str | None = None,
        rows: Iterator[int, ...] | None = None,
        cols: Iterator[int, ...] | None = None,
        indexes: Iterator[int, ...] | None = None,
        headers: Iterator[int, ...] | None = None,
        index: bool = False,
        header: bool = False,
        name: str | int | None = None,
    ) -> None:
        __slots__ = (  # noqa: F841
            "cells",
            "rows",
            "cols",
            "indexes" "headers" "index" "header" "name",
        )
        self.cells = cells
        self.rows = rows
        self.cols = cols
        self.indexes = indexes
        self.headers = headers
        self.index = index
        self.header = header
        self.name = name if name is None else f"{name}"


class TextEditor_(tk.Text):
    def __init__(
        self,
        parent,
        font=get_font(),
        text: None | str = None,
        state="normal",
        bg="white",
        fg="black",
        popup_menu_font=("Arial", 11, "normal"),
        popup_menu_bg="white",
        popup_menu_fg="black",
        popup_menu_highlight_bg="blue",
        popup_menu_highlight_fg="white",
        align="w",
        newline_binding=None,
    ):
        tk.Text.__init__(
            self,
            parent,
            font=font,
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
            label="Select all",
            accelerator="Ctrl+A",
            font=popup_menu_font,
            foreground=popup_menu_fg,
            background=popup_menu_bg,
            activebackground=popup_menu_highlight_bg,
            activeforeground=popup_menu_highlight_fg,
            command=self.select_all,
        )
        self.rc_popup_menu.add_command(
            label="Cut",
            accelerator="Ctrl+X",
            font=popup_menu_font,
            foreground=popup_menu_fg,
            background=popup_menu_bg,
            activebackground=popup_menu_highlight_bg,
            activeforeground=popup_menu_highlight_fg,
            command=self.cut,
        )
        self.rc_popup_menu.add_command(
            label="Copy",
            accelerator="Ctrl+C",
            font=popup_menu_font,
            foreground=popup_menu_fg,
            background=popup_menu_bg,
            activebackground=popup_menu_highlight_bg,
            activeforeground=popup_menu_highlight_fg,
            command=self.copy,
        )
        self.rc_popup_menu.add_command(
            label="Paste",
            accelerator="Ctrl+V",
            font=popup_menu_font,
            foreground=popup_menu_fg,
            background=popup_menu_bg,
            activebackground=popup_menu_highlight_bg,
            activeforeground=popup_menu_highlight_fg,
            command=self.paste,
        )
        self.rc_popup_menu.add_command(
            label="Undo",
            accelerator="Ctrl+Z",
            font=popup_menu_font,
            foreground=popup_menu_fg,
            background=popup_menu_bg,
            activebackground=popup_menu_highlight_bg,
            activeforeground=popup_menu_highlight_fg,
            command=self.undo,
        )
        self.bind("<1>", lambda event: self.focus_set())
        self.bind(rc_binding, self.rc)
        self._orig = self._w + "_orig"
        self.tk.call("rename", self._w, self._orig)
        self.tk.createcommand(self._w, self._proxy)

    def _proxy(self, command, *args):
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

    def rc(self, event):
        self.focus_set()
        self.rc_popup_menu.tk_popup(event.x_root, event.y_root)

    def select_all(self, event=None):
        self.event_generate(f"<{ctrl_key}-a>")
        return "break"

    def cut(self, event=None):
        self.event_generate(f"<{ctrl_key}-x>")
        return "break"

    def copy(self, event=None):
        self.event_generate(f"<{ctrl_key}-c>")
        return "break"

    def paste(self, event=None):
        self.event_generate(f"<{ctrl_key}-v>")
        return "break"

    def undo(self, event=None):
        self.event_generate(f"<{ctrl_key}-z>")
        return "break"


class TextEditor(tk.Frame):
    def __init__(
        self,
        parent,
        font=get_font(),
        text=None,
        state="normal",
        width=None,
        height=None,
        border_color="black",
        show_border=True,
        bg="white",
        fg="black",
        popup_menu_font=("Arial", 11, "normal"),
        popup_menu_bg="white",
        popup_menu_fg="black",
        popup_menu_highlight_bg="blue",
        popup_menu_highlight_fg="white",
        binding=None,
        align="w",
        r=0,
        c=0,
        newline_binding=None,
    ):
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
            font=font,
            text=text,
            state=state,
            bg=bg,
            fg=fg,
            popup_menu_font=popup_menu_font,
            popup_menu_bg=popup_menu_bg,
            popup_menu_fg=popup_menu_fg,
            popup_menu_highlight_bg=popup_menu_highlight_bg,
            popup_menu_highlight_fg=popup_menu_highlight_fg,
            align=align,
            newline_binding=newline_binding,
        )
        self.textedit.grid(row=0, column=0, sticky="nswe")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.grid_propagate(False)
        self.w_ = width
        self.h_ = height
        self.binding = binding
        self.textedit.focus_set()

    def get(self):
        return self.textedit.get("1.0", "end-1c")

    def get_num_lines(self):
        return int(self.textedit.index("end-1c").split(".")[0])

    def set_text(self, text):
        self.textedit.delete(1.0, "end")
        self.textedit.insert(1.0, text)

    def scroll_to_bottom(self):
        self.textedit.yview_moveto(1)


class GeneratedMouseEvent:
    def __init__(self):
        self.keycode = "??"
        self.num = 1
