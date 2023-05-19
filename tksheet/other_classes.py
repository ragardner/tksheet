import bisect
import tkinter as tk
from collections import namedtuple
from itertools import islice
from typing import Any, Union

from .vars import ctrl_key, get_font, rc_binding

CurrentlySelectedClass = namedtuple("CurrentlySelectedClass", "row column type_ tags")
DrawnItem = namedtuple("DrawnItem", "iid showing")
TextCfg = namedtuple("TextCfg", "txt tf font align")
DraggedRowColumn = namedtuple("DraggedRowColumn", "dragged to_move")
_ProgBar = namedtuple("_ProgBar", "bg fg pc name")


class TextEditor_(tk.Text):
    def __init__(
        self,
        parent,
        font=get_font(),
        text=None,
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
        if command in ("insert", "delete", "replace"):
            self.tag_add("align", 1.0, "end")
            self.event_generate("<<TextModified>>")
            if args and len(args) > 1 and args[1] != "\n":
                out_of_bounds = self.yview()
                if out_of_bounds != (0.0, 1.0) and self.newline_bindng is not None:
                    self.newline_bindng(r=self.parent.r, c=self.parent.c, check_lines=False)
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


def tksheet_type_error(kwarg, valid_types, not_type):
    valid_types = ", ".join(f"{type_}" for type_ in valid_types)
    return f"Argument '{kwarg}' must be one of the following types: {valid_types}, " f"not {type(not_type)}."


def dropdown_search_function(search_for, data):
    search_len = len(search_for)
    best_match = {"rn": float("inf"), "st": float("inf"), "len_diff": float("inf")}
    for rn, row in enumerate(data):
        dd_val = rf"{row[0]}".lower()
        st = dd_val.find(search_for)
        if st > -1:
            # priority is start index
            # if there's already a matching start
            # then compare the len difference
            len_diff = len(dd_val) - search_len
            if st < best_match["st"] or (st == best_match["st"] and len_diff < best_match["len_diff"]):
                best_match["rn"] = rn
                best_match["st"] = st
                best_match["len_diff"] = len_diff
    if best_match["rn"] != float("inf"):
        return best_match["rn"]
    return None


def event_dict(
    name: str = None,
    sheet: Any = None,
    boxes: Union[None, dict, tuple] = None,
    cells_table: Union[None, dict] = None,
    cells_header: Union[None, dict] = None,
    cells_index: Union[None, dict] = None,
    selected: Union[None, tuple] = None,
    data: Any = None,
    key: Union[None, str] = None,
    value: Any = None,
    location: Union[None, int, tuple[int]] = None,
    resized_rows: Union[None, dict] = None,
    resized_columns: Union[None, dict] = None,
    #resized_index: Union[None, dict] = None,
    #resized_header: Union[None, dict] = None,
    **kwargs,
) -> dict:
    return {
        "eventname": "" if name is None else name,
        "sheetname": "!sheet" if sheet is None else sheet,
        "cells": {
            "table": {} if cells_table is None else cells_table,
            "header": {} if cells_header is None else cells_header,
            "index": {} if cells_index is None else cells_index,
        },
        "moved": {
            "rows": {},
            "columns": {},
        },
        "added": {
            "rows": {},
            "columns": {},
        },
        "deleted": {
            "rows": {},
            "columns": {},
            "header": {},
            "index": {},
            "column_widths": {},
            "row_heights": {},
            "options": {},
            "old_displayed_columns": None,
            "old_displayed_rows": None,
        },
        "selection_boxes": {} if boxes is None else {boxes[:-1]: boxes[-1]} if isinstance(boxes, tuple) else boxes,
        "selected": tuple() if selected is None else selected,
        "data": [] if data is None else data,
        "key": "" if key is None else key,
        "value": None if value is None else value,
        "location": tuple() if location is None else location,
        "resized": {
            "rows": {} if resized_rows is None else resized_rows,
            "columns": {} if resized_columns is None else resized_columns,
            #"header": {} if resized_header is None else resized_header,
            #"index": {} if resized_index is None else resized_index,
        }
    }


def get_dropdown_kwargs(
    values=[],
    set_value=None,
    state="normal",
    redraw=True,
    selection_function=None,
    modified_function=None,
    search_function=dropdown_search_function,
    validate_input=True,
    text=None,
    **kwargs,
):
    return {
        "values": values,
        "set_value": set_value,
        "state": state,
        "redraw": redraw,
        "selection_function": selection_function,
        "modified_function": modified_function,
        "search_function": search_function,
        "validate_input": validate_input,
        "text": text,
    }


def get_dropdown_dict(**kwargs):
    return {
        "values": kwargs["values"],
        "window": "no dropdown open",
        "canvas_id": "no dropdown open",
        "select_function": kwargs["selection_function"],
        "modified_function": kwargs["modified_function"],
        "search_function": kwargs["search_function"],
        "validate_input": kwargs["validate_input"],
        "text": kwargs["text"],
        "state": kwargs["state"],
    }


def get_checkbox_kwargs(checked=False, state="normal", redraw=True, check_function=None, text="", **kwargs):
    return {
        "checked": checked,
        "state": state,
        "redraw": redraw,
        "check_function": check_function,
        "text": text,
    }


def get_checkbox_dict(**kwargs):
    return {
        "check_function": kwargs["check_function"],
        "state": kwargs["state"],
        "text": kwargs["text"],
    }


def is_iterable(o):
    if isinstance(o, str):
        return False
    try:
        iter(o)
        return True
    except Exception:
        return False


def num2alpha(n):
    s = ""
    n += 1
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def get_n2a(n=0, _type="numbers"):
    if _type == "letters":
        return num2alpha(n)
    elif _type == "numbers":
        return f"{n + 1}"
    else:
        return f"{num2alpha(n)} {n + 1}"


def get_index_of_gap_in_sorted_integer_seq_forward(seq, start=0):
    prevn = seq[start]
    for idx, n in enumerate(islice(seq, start + 1, None), start + 1):
        if n != prevn + 1:
            return idx
        prevn = n
    return None


def get_index_of_gap_in_sorted_integer_seq_reverse(seq, start=0):
    prevn = seq[start]
    for idx, n in zip(range(start, -1, -1), reversed(seq[:start])):
        if n != prevn - 1:
            return idx
        prevn = n
    return None


def get_seq_without_gaps_at_index(seq, position, get_st_end=False):
    start_idx = bisect.bisect_left(seq, position)
    forward_gap = get_index_of_gap_in_sorted_integer_seq_forward(seq, start_idx)
    reverse_gap = get_index_of_gap_in_sorted_integer_seq_reverse(seq, start_idx)
    if forward_gap is not None:
        seq[:] = seq[:forward_gap]
    if reverse_gap is not None:
        seq[:] = seq[reverse_gap:]
    if get_st_end:
        return seq[0], seq[-1]
    return seq


def get_checkbox_points(x1, y1, x2, y2, radius=8):
    return [
        x1 + radius,
        y1,
        x1 + radius,
        y1,
        x2 - radius,
        y1,
        x2 - radius,
        y1,
        x2,
        y1,
        x2,
        y1 + radius,
        x2,
        y1 + radius,
        x2,
        y2 - radius,
        x2,
        y2 - radius,
        x2,
        y2,
        x2 - radius,
        y2,
        x2 - radius,
        y2,
        x1 + radius,
        y2,
        x1 + radius,
        y2,
        x1,
        y2,
        x1,
        y2 - radius,
        x1,
        y2 - radius,
        x1,
        y1 + radius,
        x1,
        y1 + radius,
        x1,
        y1,
    ]