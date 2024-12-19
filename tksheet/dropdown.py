from __future__ import annotations

import tkinter as tk
from collections.abc import Sequence
from typing import Literal

from .sheet import Sheet


class Dropdown(Sheet):
    def __init__(
        self,
        parent: tk.Misc,
        state: Literal["normal", "readonly", "disabled"] = "readonly",
        values: Sequence[str] = [],
        justify: Literal["left", "center", "right"] = "left",
        font: tuple[str, int, str] | None = None,
        **kwargs,
    ) -> None:
        super().__init__(
            parent,
            show_row_index=False,
            show_header=False,
            show_top_left=False,
            show_x_scrollbar=False,
            show_y_scrollbar=False,
            show_horizontal_grid=False,
            show_vertical_grid=False,
            outline_thickness=1,
            toplevel_dropdowns=True,
            align=justify,
            **kwargs,
        )
        if font is not None:
            self.font(newfont=font)
        self.dropdown(0, 0, state=state, values=values)
        self.set_row_heights()
        self.set_column_widths()
        if state != "disabled":
            self._bind_b1()
        self.height_and_width(height=self.MT.row_positions[1], width=self.MT.col_positions[1])

    def _bind_b1(self) -> None:
        self.bind("<Button-1>", self.b1)

    def b1(self, event: object = None) -> None:
        if self.MT.dropdown.open:
            self.close_dropdown()
        else:
            self.open_dropdown(0, 0)

    def open(self, event: object = None) -> None:
        self.open_dropdown(0, 0)

    def close(self, event: object = None) -> None:
        self.close_dropdown(0, 0)

    def __getitem__(self, key: int | Literal["state", "values", "justify", "font"]) -> object:
        if key == "state":
            return self.props(0, 0, key="dropdown")["state"]
        elif key == "values":
            return self.values()
        elif key == "justify":
            return self.table_align()
        elif key == "font":
            return self.font()

    def __setitem__(self, key: str, value: object) -> None:
        if key == "state":
            self.props(0, 0, key="dropdown")["state"] = value
            if value == "disabled":
                self.unbind("<Button-1>")
            else:
                self._bind_b1()
        elif key == "values":
            self.values(value)
        elif key == "justify":
            self.table_align(value)
        elif key == "font":
            self.font(value)

    def set(self, value: str) -> None:
        self.MT.data[0][0] = value

    def get(self) -> str:
        return self.MT.data[0][0]

    def current(self, newindex: int | None = None) -> int:
        values = self.values()
        if isinstance(newindex, int):
            if len(values) > newindex:
                self.set(values[newindex])
        else:
            try:
                e = self.get()
                index = values.index(e)
                return index
            except Exception:
                return -1

    def values(self, values: Sequence[str] | None = None) -> Sequence[str]:
        if values:
            self.set_dropdown_values(values=values)
            self.MT.data[0][0] = values[0]
        return self.props(0, 0, key="dropdown")["values"]

    def config_(
        self,
        state: Literal["normal", "readonly", "disabled"] | None = None,
        values: Sequence[str] | None = None,
        justify: Literal["left", "center", "right"] | None = None,
        font: tuple[str, int, str] | None = None,
        **kwargs,
    ) -> None:
        if state is not None:
            self.__setitem__("state", state)
        if values is not None:
            self.__setitem__("values", values)
        if justify is not None:
            self.__setitem__("justify", justify)
        if font is not None:
            self.__setitem__("font", font)
        if kwargs:
            self.set_options(**kwargs)
