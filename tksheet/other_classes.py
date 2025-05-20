from __future__ import annotations

import copy
import tkinter as tk
from collections import namedtuple
from collections.abc import Callable, Hashable, Iterator
from typing import Any, Literal

FontTuple = namedtuple("FontTuple", "family size style")
Box_nt = namedtuple(
    "Box_nt",
    "from_r from_c upto_r upto_c",
)
Box_t = namedtuple(
    "Box_t",
    "from_r from_c upto_r upto_c type_",
)
Box_st = namedtuple("Box_st", "coords type_")
Loc = namedtuple("Loc", "row column", defaults=(None, None))

Highlight = namedtuple(
    "Highlight",
    (
        "bg",
        "fg",
        "end",  # only used for row options highlights
    ),
    defaults=(
        None,
        None,
        False,
    ),
)
DrawnItem = namedtuple("DrawnItem", "iid showing")
TextCfg = namedtuple("TextCfg", "txt tf font align")
DraggedRowColumn = namedtuple("DraggedRowColumn", "dragged to_move")


class SelectionBox:
    __slots__ = ("fill_iid", "bd_iid", "index", "header", "coords", "type_", "state")

    def __init__(
        self,
        fill_iid: int | None = None,
        bd_iid: int | None = None,
        index: int | None = None,
        header: int | None = None,
        coords: tuple[int, int, int, int] = None,
        type_: Literal["cells", "rows", "columns"] = "cells",
        state: Literal["normal", "hidden"] = "normal",
    ) -> None:
        self.fill_iid = fill_iid
        self.bd_iid = bd_iid
        self.index = index
        self.header = header
        self.coords = coords
        self.type_ = type_
        self.state = state


Selected = namedtuple(
    "Selected",
    (
        "row",
        "column",
        "type_",
        "box",
        "iid",
        "fill_iid",
    ),
    defaults=(
        None,
        None,
        None,
        None,
        None,
        None,
    ),
)


def num2alpha(n: int) -> str | None:
    try:
        s = ""
        n += 1
        while n > 0:
            n, r = divmod(n - 1, 26)
            s = chr(65 + r) + s
        return s
    except Exception:
        return None


class SpanRange:
    def __init__(self, from_: int, upto_: int) -> None:
        __slots__ = ("from_", "upto_")  # noqa: F841
        self.from_ = from_
        self.upto_ = upto_

    def __iter__(self) -> Iterator:
        return iter(range(self.from_, self.upto_))

    def __reversed__(self) -> Iterator:
        return reversed(range(self.from_, self.upto_))

    def __contains__(self, n: int) -> bool:
        return n >= self.from_ and n < self.upto_

    def __eq__(self, v: SpanRange) -> bool:
        return self.from_ == v.from_ and self.upto_ == v.upto_

    def __ne__(self, v: SpanRange) -> bool:
        return self.from_ != v.from_ or self.upto_ != v.upto_

    def __len__(self) -> int:
        return self.upto_ - self.from_


class DotDict(dict):
    def __init__(self, *args, **kwargs) -> None:
        super().__init__(*args, **kwargs)
        # Recursively turn nested dicts into DotDicts
        for key, value in self.items():
            if type(value) is dict:  # noqa: E721
                self[key] = DotDict(value)

    def __getstate__(self) -> DotDict:
        return self

    def __setstate__(self, state: DotDict) -> None:
        self.update(state)

    def __setitem__(self, key: Hashable, item: Any) -> None:
        if type(item) is dict:  # noqa: E721
            super().__setitem__(key, DotDict(item))
        else:
            super().__setitem__(key, item)

    __setattr__ = __setitem__
    __getattr__ = dict.__getitem__
    __delattr__ = dict.__delitem__


class EventDataDict(DotDict):
    """
    A subclass of DotDict with no changes
    For better clarity in type hinting
    """


class Span(dict):
    def __init__(self, *args, **kwargs) -> None:
        super().__init__(*args, **kwargs)
        # Recursively turn nested dicts into DotDicts
        for key, item in self.items():
            if key == "data" or key == "value":
                self["widget"].set_data(self, data=item)
            elif type(item) is dict:  # noqa: E721
                self[key] = DotDict(item)

    def __getstate__(self) -> Span:
        return self

    def __setstate__(self, state: Span) -> None:
        self.update(state)

    def __getitem__(self, key: Hashable) -> Any:
        if key == "data" or key == "value":
            return self["widget"].get_data(self)
        else:
            return super().__getitem__(key)

    def __setitem__(self, key: Hashable, item: Any) -> None:
        if key == "data" or key == "value":
            self["widget"].set_data(self, data=item)
        elif key == "bg":
            self["widget"].highlight(self, bg=item)
        elif key == "fg":
            self["widget"].highlight(self, fg=item)
        elif key == "align":
            self["widget"].align(self, align=item)
        elif type(item) is dict:  # noqa: E721
            super().__setitem__(key, DotDict(item))
        else:
            super().__setitem__(key, item)

    def format(
        self,
        formatter_options: dict | None = None,
        formatter_class: Any = None,
        redraw: bool = True,
        **kwargs,
    ) -> Span:
        if formatter_options is None:
            formatter_options = {}
        return self["widget"].format(
            self,
            formatter_options={"formatter": formatter_class, **formatter_options, **kwargs},
            formatter_class=formatter_class,
            redraw=redraw,
            **kwargs,
        )

    def del_format(self) -> Span:
        return self["widget"].del_format(self)

    def note(self, note: str | None = None, readonly: bool = True) -> Span:
        return self["widget"].note(self, note=note, readonly=readonly)

    def highlight(
        self,
        bg: bool | None | str = False,
        fg: bool | None | str = False,
        end: bool | None = None,
        overwrite: bool = False,
        redraw: bool = True,
    ) -> Span:
        return self["widget"].highlight(
            self,
            bg=bg,
            fg=fg,
            end=end,
            overwrite=overwrite,
            redraw=redraw,
        )

    def dehighlight(self, redraw: bool = True) -> Span:
        return self["widget"].dehighlight(self, redraw=redraw)

    del_highlight = dehighlight

    def readonly(self, readonly: bool = True) -> Span:
        return self["widget"].readonly(self, readonly=readonly)

    def dropdown(
        self,
        values: list[Any] | None = None,
        edit_data: bool = True,
        set_values: dict[tuple[int, int], Any] | None = None,
        set_value: Any = None,
        state: str = "normal",
        redraw: bool = True,
        selection_function: Callable | None = None,
        modified_function: Callable | None = None,
        search_function: Callable | None = None,
        validate_input: bool = True,
        text: None | str = None,
    ) -> Span:
        return self["widget"].dropdown(
            self,
            values=[] if values is None else values,
            edit_data=edit_data,
            set_values={} if set_values is None else set_values,
            set_value=set_value,
            state=state,
            redraw=redraw,
            selection_function=selection_function,
            modified_function=modified_function,
            search_function=search_function,
            validate_input=validate_input,
            text=text,
        )

    def del_dropdown(self) -> Span:
        return self["widget"].del_dropdown(self)

    def checkbox(
        self,
        edit_data: bool = True,
        checked: bool | None = None,
        state: str = "normal",
        redraw: bool = True,
        check_function: Callable | None = None,
        text: str = "",
    ) -> Span:
        return self["widget"].checkbox(
            self,
            edit_data=edit_data,
            checked=checked,
            state=state,
            redraw=redraw,
            check_function=check_function,
            text=text,
        )

    def del_checkbox(self) -> Span:
        return self["widget"].del_checkbox(self)

    def align(self, align: str | None, redraw: bool = True) -> Span:
        return self["widget"].align(self, align=align, redraw=redraw)

    def del_align(self, redraw: bool = True) -> Span:
        return self["widget"].del_align(self, redraw=redraw)

    def clear(self, undo: bool | None = None, redraw: bool = True) -> Span:
        if undo is not None:
            self["widget"].clear(self, undo=undo, redraw=redraw)
        else:
            self["widget"].clear(self, redraw=redraw)
        return self

    def tag(self, *tags) -> Span:
        self["widget"].tag(self, tags=tags)
        return self

    def untag(self) -> Span:
        if self.kind == "cell":
            for r in self.rows:
                for c in self.columns:
                    self["widget"].untag(cell=(r, c))
        elif self.kind == "row":
            self["widget"].untag(rows=self.rows)
        elif self.kind == "column":
            self["widget"].untag(columns=self.columns)
        return self

    def options(
        self,
        type_: str | None = None,
        name: str | None = None,
        table: bool | None = None,
        index: bool | None = None,
        header: bool | None = None,
        tdisp: bool | None = None,
        idisp: bool | None = None,
        hdisp: bool | None = None,
        transposed: bool | None = None,
        ndim: int | None = None,
        convert: Callable | None = None,
        undo: bool | None = None,
        emit_event: bool | None = None,
        widget: Any = None,
        expand: str | None = None,
        formatter_options: dict | None = None,
        **kwargs,
    ) -> Span:
        if isinstance(expand, str) and expand.lower() in ("down", "right", "both", "table"):
            self.expand(expand)

        if isinstance(convert, Callable):
            self["convert"] = convert

        if isinstance(type_, str):
            self["type_"] = type_.lower()

        if isinstance(name, str):
            if isinstance(name, str) and not name:
                name = f"{num2alpha(self['widget'].named_span_id)}"
                self["widget"].named_span_id += 1
            self["name"] = name

        if isinstance(table, bool):
            self["table"] = table
        if isinstance(index, bool):
            self["index"] = index
        if isinstance(header, bool):
            self["header"] = header
        if isinstance(transposed, bool):
            self["transposed"] = transposed
        if isinstance(tdisp, bool):
            self["tdisp"] = tdisp
        if isinstance(idisp, bool):
            self["idisp"] = idisp
        if isinstance(hdisp, bool):
            self["hdisp"] = hdisp
        if isinstance(undo, bool):
            self["undo"] = undo
        if isinstance(emit_event, bool):
            self["emit_event"] = emit_event

        if isinstance(ndim, int) and ndim in (0, 1, 2):
            self["ndim"] = ndim

        if isinstance(formatter_options, dict):
            self["type_"] = "format"
            self["kwargs"] = {"formatter": None, **formatter_options}

        elif kwargs:
            self["kwargs"] = kwargs

        if widget is not None:
            self["widget"] = widget
        return self

    def transpose(self) -> Span:
        self["transposed"] = not self["transposed"]
        return self

    def expand(self, direction: Literal["both", "table", "down", "right"] = "both") -> Span:
        if direction == "both" or direction == "table":
            self["upto_r"], self["upto_c"] = None, None
        elif direction == "down":
            self["upto_r"] = None
        elif direction == "right":
            self["upto_c"] = None
        else:
            raise ValueError(f"Expand argument must be either 'both', 'table', 'down' or 'right'. Not {direction}")
        return self

    @property
    def kind(self) -> str:
        if self["from_r"] is None:
            return "column"
        if self["from_c"] is None:
            return "row"
        return "cell"

    @property
    def rows(self) -> SpanRange:
        rng_from_r = 0 if self["from_r"] is None else self["from_r"]
        rng_upto_r = self["widget"].total_rows() if self["upto_r"] is None else self["upto_r"]
        return SpanRange(rng_from_r, rng_upto_r)

    @property
    def columns(self) -> SpanRange:
        rng_from_c = 0 if self["from_c"] is None else self["from_c"]
        rng_upto_c = self["widget"].total_columns() if self["upto_c"] is None else self["upto_c"]
        return SpanRange(rng_from_c, rng_upto_c)

    @property
    def coords(self) -> tuple[int, int, int, int]:
        rows = self.rows
        cols = self.columns
        return Box_nt(rows.from_, cols.from_, rows.upto_, cols.upto_)

    def copy_self(self) -> "Span":
        # Create a new Span instance
        span = Span()

        # Iterate over all dictionary items to capture all attributes
        for key, value in self.items():
            if key == "widget":
                # Tkinter widget: retain reference (do not copy)
                span[key] = value
            elif key == "kwargs":
                # Handle kwargs, which may contain functions or lambdas
                span[key] = {}
                for k, v in value.items():
                    if callable(v):
                        # Functions/lambdas: shallow copy (immutable, but check for safety)
                        span[key][k] = v
                    else:
                        try:
                            # Deep copy non-callable items in kwargs
                            span[key][k] = copy.deepcopy(v)
                        except Exception:
                            # Handle non-copyable objects (e.g., complex closures or objects)
                            span[key][k] = v  # Fallback to shallow copy
            elif key == "convert" and callable(value):
                # Convert is a callable: shallow copy (immutable)
                span[key] = value
            else:
                # Deep copy other values to handle nested objects like DotDict
                try:
                    span[key] = copy.deepcopy(value)
                except Exception:
                    # Fallback for non-copyable objects
                    span[key] = value  # Shallow copy as fallback

        # Ensure widget is set if not already present (edge case)
        if "widget" not in span and hasattr(self, "widget"):
            span["widget"] = self["widget"]

        return span

    __setattr__ = __setitem__
    __getattr__ = __getitem__
    __delattr__ = dict.__delitem__


class GeneratedMouseEvent:
    def __init__(self):
        self.keycode = "??"
        self.num = 1


class Node:
    __slots__ = ("text", "iid", "parent", "children")

    def __init__(
        self,
        text: str,
        iid: str,
        parent: str = "",
        children: list[str] | None = None,
    ) -> None:
        self.text: str = text
        self.iid: str = iid
        self.parent: str = parent
        self.children: list[str] = children if children else []


class StorageBase:
    __slots__ = ("canvas_id", "window", "open")

    def __init__(self) -> None:
        self.canvas_id = None
        self.window = None
        self.open = False


class DropdownStorage(StorageBase):
    def get_coords(self) -> int | tuple[int, int] | None:
        """
        Returns None if not open or window is None
        """
        if self.open and self.window is not None:
            return self.window.get_coords()
        return None


class EditorStorageBase(StorageBase):
    def focus(self) -> None:
        if self.window:
            self.window.tktext.focus_set()

    def get(self) -> str:
        if self.window:
            return self.window.get()
        return ""

    def set(self, value: str) -> None:
        if not self.window:
            return
        self.window.set_text(value)

    def highlight_from(self, index: tk.Misc, r: int | str, c: int | str) -> None:
        self.window.tktext.tag_add("sel", index, "end")
        self.window.tktext.mark_set("insert", f"{r}.{c}")

    def autocomplete(self, value: str | None) -> None:
        current_val = self.get()
        if not value or len(current_val) >= len(value) or current_val != value[: len(current_val)]:
            return
        cursor_pos = self.tktext.index("insert")
        line, column = cursor_pos.split(".")
        index = self.window.tktext.index(f"{line}.{column}")
        self.tktext.insert(index, value[len(current_val) :])
        self.highlight_from(index, line, column)

    @property
    def tktext(self) -> Any:
        if self.window:
            return self.window.tktext
        return self.window


class TextEditorStorage(EditorStorageBase):
    @property
    def coords(self) -> tuple[int, int]:
        return self.window.r, self.window.c

    @property
    def row(self) -> int:
        return self.window.r

    @property
    def column(self) -> int:
        return self.window.c


class ProgressBar:
    __slots__ = ("bg", "fg", "name", "percent", "del_when_done")

    def __init__(self, bg: str, fg: str, name: Hashable, percent: int, del_when_done: bool) -> None:
        self.bg = bg
        self.fg = fg
        self.name = name
        self.percent = percent
        self.del_when_done = del_when_done

    def __len__(self):
        return 2

    def __getitem__(self, key: Hashable) -> Any:
        if key == 0:
            return self.bg
        elif key == 1:
            return self.fg
        elif key == 2:
            return self.name
        elif key == 3:
            return self.percent
        elif key == 4:
            return self.del_when_done
        else:
            return self.__getattribute__(key)
