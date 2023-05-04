from __future__ import annotations

import bisect
import tkinter as tk
from itertools import islice
from typing import TYPE_CHECKING, NamedTuple

from ._tksheet_vars import *

if TYPE_CHECKING:
    from collections.abc import MutableSequence, Sequence
    from typing import Any

    from typing_extensions import Self

    from tksheet.types import Align, Font, ScreenUnits, State


class CurrentlySelectedClass(NamedTuple):
    row: int
    column: int
    type_: ...


class CtrlKeyEvent(NamedTuple):
    eventname: str
    selectionboxes: ...
    currentlyselected: ...
    rows: ...


class PasteEvent(NamedTuple):
    eventname: str
    currentlyselected: ...
    rows: ...


class UndoEvent(NamedTuple):
    eventname: str
    type: ...
    storeddata: ...


class SelectCellEvent(NamedTuple):
    eventname: str
    row: int
    column: int


class SelectColumnEvent(NamedTuple):
    eventname: str
    column: int


class SelectRowEvent(NamedTuple):
    eventname: str
    row: int


class DeselectionEvent(NamedTuple):
    eventname: str
    selectionboxes: ...


class SelectionEvent(NamedTuple):
    eventname: str
    selectionboxes: ...


class InsertEvent(NamedTuple):
    eventname: str
    dataindex: ...
    displayindex: ...
    quantity: ...


class DeleteRowColumnEvent(NamedTuple):
    eventname: str
    deleteindexes: ...


class EditCellEvent(NamedTuple):
    row: int
    column: int
    key: ...
    text: str
    eventname: str


class EditHeaderEvent(NamedTuple):
    column: int
    key: ...
    text: str
    eventname: str


class EditIndexEvent(NamedTuple):
    row: int
    key: ...
    text: str
    eventname: str


class BeginDragDropEvent(NamedTuple):
    eventname: str
    columnstomove: ...
    movedto: ...


class EndDragDropEvent(NamedTuple):
    eventname: str
    oldindexes: ...
    newindexes: ...
    movedto: ...


class ResizeEvent(NamedTuple):
    eventname: str
    index: ...
    oldsize: ...
    newsize: ...


class DropDownModifiedEvent(NamedTuple):
    eventname: str
    row: int
    column: int
    value: ...


class DrawnItem(NamedTuple):
    iid: ...
    showing: ...


class TextCfg(NamedTuple):
    txt: str
    tf: ...
    font: Font
    align: Align


class DraggedRowColumn(NamedTuple):
    dragged: ...
    to_move: ...


class _ProgBar(NamedTuple):
    bg: ...
    fg: ...
    pc: ...
    name: str


class TextEditor_(tk.Text):
    def __init__(
        self,
        parent: tk.Misc,
        font: Font = get_font(),
        text: str | None = None,
        state: State = "normal",
        bg: str = "white",
        fg: str = "black",
        popup_menu_font: Font = ("Arial", 11, "normal"),
        popup_menu_bg: str = "white",
        popup_menu_fg: str = "black",
        popup_menu_highlight_bg: str = "blue",
        popup_menu_highlight_fg: str = "white",
        align: Align = "w",
        newline_binding=None,
    ):
        super().__init__(
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
        self.bind("<1>", lambda _: self.focus_set())
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

    def rc(self, event: tk.Event[Self][Self]) -> None:
        self.focus_set()
        self.rc_popup_menu.tk_popup(event.x_root, event.y_root)

    def select_all(self):
        self.event_generate(f"<{ctrl_key}-a>")
        return "break"

    def cut(self):
        self.event_generate(f"<{ctrl_key}-x>")
        return "break"

    def copy(self):
        self.event_generate(f"<{ctrl_key}-c>")
        return "break"

    def paste(self):
        self.event_generate(f"<{ctrl_key}-v>")
        return "break"

    def undo(self):
        self.event_generate(f"<{ctrl_key}-z>")
        return "break"


class TextEditor(tk.Frame):
    def __init__(
        self,
        parent: tk.Misc,
        font: Font = get_font(),
        text: str | None = None,
        state: State = "normal",
        width: ScreenUnits | None = None,
        height: ScreenUnits | None = None,
        border_color: str = "black",
        show_border: bool = True,
        bg: str = "white",
        fg: str = "black",
        popup_menu_font: Font = ("Arial", 11, "normal"),
        popup_menu_bg: str = "white",
        popup_menu_fg: str = "black",
        popup_menu_highlight_bg: str = "blue",
        popup_menu_highlight_fg: str = "white",
        binding=None,
        align: Align = "w",
        r: int = 0,
        c: int = 0,
        newline_binding=None,
    ):
        super().__init__(
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

    def set_text(self, text: str) -> None:
        self.textedit.delete(1.0, "end")
        self.textedit.insert(1.0, text)

    def scroll_to_bottom(self):
        self.textedit.yview_moveto(1)


class GeneratedMouseEvent:
    def __init__(self):
        self.keycode = "??"
        self.num = 1


def dropdown_search_function(search_for: str, data) -> str | None:
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


def get_dropdown_kwargs(
    values=[],
    set_value=None,
    state: State = "normal",
    redraw: bool = True,
    selection_function=None,
    modified_function=None,
    search_function=dropdown_search_function,
    validate_input: bool = True,
    text: str | None = None,
    **_,
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


def get_checkbox_kwargs(
    checked: bool = False, state: State = "normal", redraw: bool = True, check_function=None, text: str = "", **_
):
    return {"checked": checked, "state": state, "redraw": redraw, "check_function": check_function, "text": text}


def get_checkbox_dict(**kwargs):
    return {"check_function": kwargs["check_function"], "state": kwargs["state"], "text": kwargs["text"]}


def is_iterable(o: Any):
    if isinstance(o, str):
        return False
    try:
        iter(o)
        return True
    except Exception:
        return False


def num2alpha(n: int) -> str:
    s = ""
    n += 1
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def get_n2a(n: int = 0, _type: str = "numbers") -> str:
    if _type == "letters":
        return num2alpha(n)
    elif _type == "numbers":
        return f"{n + 1}"
    else:
        return f"{num2alpha(n)} {n + 1}"


def get_index_of_gap_in_sorted_integer_seq_forward(seq: Sequence[int], start: int = 0):
    prevn = seq[start]
    for idx, n in enumerate(islice(seq, start + 1, None), start + 1):
        if n != prevn + 1:
            return idx
        prevn = n
    return None


def get_index_of_gap_in_sorted_integer_seq_reverse(seq: Sequence[int], start: int = 0):
    prevn = seq[start]
    for idx, n in zip(range(start, -1, -1), reversed(seq[:start])):
        if n != prevn - 1:
            return idx
        prevn = n
    return None


def get_seq_without_gaps_at_index(seq: MutableSequence[int], position: int) -> Sequence[int]:
    start_idx = bisect.bisect_left(seq, position)
    forward_gap = get_index_of_gap_in_sorted_integer_seq_forward(seq, start_idx)
    reverse_gap = get_index_of_gap_in_sorted_integer_seq_reverse(seq, start_idx)
    if forward_gap is not None:
        seq[:] = seq[:forward_gap]
    if reverse_gap is not None:
        seq[:] = seq[reverse_gap:]
    return seq
