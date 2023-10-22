from __future__ import annotations

import pickle
import tkinter as tk
from collections import namedtuple
from collections.abc import Generator, Hashable, Iterator
from functools import partial

from .vars import (
    ctrl_key,
    get_font,
    rc_binding,
)

pickle_obj = partial(pickle.dumps, protocol=pickle.HIGHEST_PROTOCOL)

CurrentlySelectedClass = namedtuple(
    "CurrentlySelectedClass",
    "row column type_ tags",
)
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

    def __setitem__(self, key: str, value: object) -> None:
        if isinstance(key, str):
            setattr(self, key, value)
        else:
            raise ValueError(f"Key must be type 'str' not '{type(key)}'.")


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
        if n >= self.from_ and n < self.upto_:
            return True
        return False

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

    def __getstate__(self) -> SpanDict:
        return self

    def __setstate__(self, state: DotDict) -> None:
        self.update(state)

    def __setitem__(self, key: Hashable, item: object) -> None:
        if type(item) is dict:  # noqa: E721
            super().__setitem__(key, DotDict(item))
        else:
            super().__setitem__(key, item)

    __setattr__ = __setitem__
    __getattr__ = dict.__getitem__
    __delattr__ = dict.__delitem__


class SpanDict(dict):
    def __init__(self, *args, **kwargs) -> None:
        super().__init__(*args, **kwargs)
        # Recursively turn nested dicts into DotDicts
        for key, item in self.items():
            if key == "data" or key == "value":
                self["widget"].set_data(self, item)
            elif type(item) is dict: # noqa: E721
                self[key] = DotDict(item)

    def __getstate__(self) -> SpanDict:
        return self

    def __setstate__(self, state: SpanDict) -> None:
        self.update(state)

    def __getitem__(self, key: Hashable) -> object:
        if key == "data" or key == "value":
            return self["widget"].get_data(self)
        else:
            return super().__getitem__(key)

    def __setitem__(self, key: Hashable, item: object) -> None:
        if key == "data" or key == "value":
            self["widget"].set_data(self, item)
        elif key == "bg":
            self["widget"].highlight(self, bg=item)
        elif key == "fg":
            self["widget"].highlight(self, fg=item)
        elif type(item) is dict: # noqa: E721
            super().__setitem__(key, DotDict(item))
        else:
            super().__setitem__(key, item)

    def format(
        self,
        formatter_options={},
        formatter_class=None,
        redraw: bool = True,
        **kwargs,
    ) -> SpanDict:
        self["widget"].format(
            self,
            formatter_options={"formatter": formatter_class, **formatter_options, **kwargs},
            formatter_class=formatter_class,
            redraw=redraw,
            **kwargs,
        )
        return self

    def del_format(self) -> SpanDict:
        self["widget"].del_format(self)
        return self

    def highlight(self, **kwargs) -> SpanDict:
        self["widget"].highlight(self, **kwargs)
        return self

    def dehighlight(self, redraw: bool = True) -> SpanDict:
        self["widget"].dehighlight(self, redraw=redraw)

    def readonly(self, readonly: bool = True) -> SpanDict:
        self["widget"].readonly(self, readonly=readonly)
        return self

    def dropdown(self, *args, **kwargs) -> SpanDict:
        self["widget"].dropdown(self, *args, **kwargs)

    def del_dropdown(self) -> SpanDict:
        self["widget"].del_dropdown(self)

    def checkbox(self, *args, **kwargs) -> SpanDict:
        self["widget"].dropdown(self, *args, **kwargs)

    def del_checkbox(self) -> SpanDict:
        self["widget"].del_checkbox(self)

    def align(self, align: str | None, redraw: bool = True) -> SpanDict:
        self["widget"].align(self, align=align, redraw=redraw)

    def del_align(self, redraw: bool = True) -> SpanDict:
        self["widget"].del_align(self, redraw=redraw)

    def clear(self, undo: bool | None = None, redraw: bool = True) -> SpanDict:
        if undo is not None:
            self["widget"].clear(self, undo=undo, redraw=redraw)
        else:
            self["widget"].clear(self, redraw=redraw)
        return self

    def options(self, convert: object = None, **kwargs) -> SpanDict:
        if "expand" in kwargs:
            self.expand(kwargs["expand"])
        for k in (
            "name",
            "index",
            "header",
            "table",
            "transpose",
            "ndim",
            "displayed",
            "undo",
        ):
            if k in kwargs:
                self[k] = kwargs[k]
        if "invert" in kwargs:
            self["transpose"] = kwargs["invert"]
        if (k := "formatter") in kwargs or (k := "format") in kwargs:
            self["type_"] = "format"
            self["kwargs"] = {"formatter": None, **kwargs[k]}
        if convert != self["convert"]:
            self["convert"] = convert
        return self

    def invert(self, invert: bool = True) -> SpanDict:
        self["transpose"] = invert
        return self

    def expand(self, direction: str = "both") -> SpanDict:
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
    def rows(self) -> Generator[int]:
        rng_from_r = 0 if self["from_r"] is None else self["from_r"]
        if self["upto_r"] is None:
            rng_upto_r = self["widget"].total_rows()
        else:
            rng_upto_r = self["upto_r"]
        return SpanRange(rng_from_r, rng_upto_r)

    @property
    def columns(self) -> Generator[int]:
        rng_from_c = 0 if self["from_c"] is None else self["from_c"]
        if self["upto_c"] is None:
            rng_upto_c = self["widget"].total_columns()
        else:
            rng_upto_c = self["upto_c"]
        return SpanRange(rng_from_c, rng_upto_c)

    @property
    def ranges(self) -> tuple[Generator[int], Generator[int]]:
        rng_from_r = 0 if self["from_r"] is None else self["from_r"]
        rng_from_c = 0 if self["from_c"] is None else self["from_c"]
        if self["upto_r"] is None:
            rng_upto_r = self["widget"].total_rows()
        else:
            rng_upto_r = self["upto_r"]
        if self["upto_c"] is None:
            rng_upto_c = self["widget"].total_columns()
        else:
            rng_upto_c = self["upto_c"]
        return SpanRange(rng_from_r, rng_upto_r), SpanRange(rng_from_c, rng_upto_c)

    def pickle_self(self) -> bytes:
        x = self["widget"]
        self["widget"] = None
        p = pickle_obj(self)
        self["widget"] = x
        return p

    __setattr__ = __setitem__
    __getattr__ = __getitem__
    __delattr__ = dict.__delitem__


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