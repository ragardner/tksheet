from __future__ import annotations

import csv
import tkinter as tk
from bisect import bisect_left
from collections import deque
from collections.abc import Callable, Generator, Hashable, Iterator, Sequence
from contextlib import suppress
from functools import partial
from itertools import accumulate, chain, filterfalse, islice, product, repeat
from operator import attrgetter
from re import IGNORECASE, escape, sub
from timeit import default_timer
from tkinter import ttk
from typing import Any, Literal

from .column_headers import ColumnHeaders
from .constants import (
    ALL_BINDINGS,
    BINDING_TO_ATTR,
    MODIFIED_BINDINGS,
    SELECT_BINDINGS,
    USER_OS,
    align_value_error,
    backwards_compatibility_keys,
    emitted_events,
    named_span_types,
    scrollbar_options_keys,
)
from .find_window import replacer
from .functions import (
    add_highlight,
    add_to_options,
    alpha2idx,
    bisect_in,
    box_gen_coords,
    consecutive_ranges,
    convert_align,
    del_from_options,
    del_named_span_options,
    del_named_span_options_nested,
    dropdown_search_function,
    event_dict,
    fix_format_kwargs,
    force_bool,
    get_checkbox_dict,
    get_checkbox_kwargs,
    get_dropdown_dict,
    get_dropdown_kwargs,
    idx_param_to_int,
    is_iterable,
    key_to_span,
    mod_event_val,
    mod_note,
    new_tk_event,
    num2alpha,
    pop_positions,
    set_align,
    set_readonly,
    span_froms,
    span_ranges,
    stored_event_dict,
    tksheet_type_error,
    try_binding,
    unpack,
)
from .main_table import MainTable
from .other_classes import (
    Box_nt,
    DotDict,
    EventDataDict,
    FontTuple,
    GeneratedMouseEvent,
    Node,
    ProgressBar,
    Selected,
    SelectionBox,
    Span,
)
from .row_index import RowIndex
from .sheet_options import new_sheet_options
from .sorting import fast_sort_key, natural_sort_key, version_sort_key  # noqa: F401
from .themes import (
    theme_black,
    theme_dark,
    theme_dark_blue,
    theme_dark_green,
    theme_light_blue,
    theme_light_green,
)
from .tksheet_types import Binding, CellPropertyKey, CreateSpanTypes, ExtraBinding
from .top_left_rectangle import TopLeftRectangle


class Sheet(tk.Frame):
    def __init__(
        self,
        parent: tk.Misc,
        name: str = "!sheet",
        show_table: bool = True,
        show_top_left: bool | None = None,
        show_row_index: bool = True,
        show_header: bool = True,
        show_x_scrollbar: bool = True,
        show_y_scrollbar: bool = True,
        width: int | None = None,
        height: int | None = None,
        headers: None | list[Any] = None,
        header: None | list[Any] = None,
        row_index: None | list[Any] = None,
        index: None | list[Any] = None,
        default_header: Literal["letters", "numbers", "both"] | None = "letters",
        default_row_index: Literal["letters", "numbers", "both"] | None = "numbers",
        data_reference: None | Sequence[Sequence[Any]] = None,
        data: None | Sequence[Sequence[Any]] = None,
        # either (start row, end row, "rows"), (start column, end column, "rows") or
        # (cells start row, cells start column, cells end row, cells end column, "cells")  # noqa: E501
        startup_select: tuple[int, int, str] | tuple[int, int, int, int, str] = None,
        startup_focus: bool = True,
        total_columns: int | None = None,
        total_rows: int | None = None,
        default_column_width: int = 120,
        default_header_height: str | int = "1",
        default_row_index_width: int = 70,
        default_row_height: str | int = "1",
        min_column_width: int = 1,
        max_column_width: float = float("inf"),
        max_row_height: float = float("inf"),
        max_header_height: float = float("inf"),
        max_index_width: float = float("inf"),
        after_redraw_time_ms: int = 16,
        set_all_heights_and_widths: bool = False,
        zoom: int = 100,
        align: str = "nw",
        header_align: str = "n",
        row_index_align: str | None = None,
        index_align: str = "n",
        displayed_columns: list[int] | None = None,
        all_columns_displayed: bool = True,
        displayed_rows: list[int] | None = None,
        all_rows_displayed: bool = True,
        to_clipboard_dialect: csv.Dialect = csv.excel_tab,
        to_clipboard_delimiter: str = "\t",
        to_clipboard_quotechar: str = '"',
        to_clipboard_lineterminator: str = "\n",
        from_clipboard_delimiters: list[str] | str = "\t",
        show_default_header_for_empty: bool = True,
        show_default_index_for_empty: bool = True,
        page_up_down_select_row: bool = True,
        paste_can_expand_x: bool = False,
        paste_can_expand_y: bool = False,
        paste_insert_column_limit: int | None = None,
        paste_insert_row_limit: int | None = None,
        show_dropdown_borders: bool = False,
        arrow_key_down_right_scroll_page: bool = False,
        cell_auto_resize_enabled: bool = True,
        auto_resize_row_index: bool | Literal["empty"] = "empty",
        auto_resize_columns: int | None = None,
        auto_resize_rows: int | None = None,
        set_cell_sizes_on_zoom: bool = False,
        font: tuple[str, int, str] = FontTuple(
            "Calibri",
            13 if USER_OS == "darwin" else 11,
            "normal",
        ),
        header_font: tuple[str, int, str] = FontTuple(
            "Calibri",
            13 if USER_OS == "darwin" else 11,
            "normal",
        ),
        index_font: tuple[str, int, str] = FontTuple(
            "Calibri",
            13 if USER_OS == "darwin" else 11,
            "normal",
        ),  # currently has no effect
        popup_menu_font: tuple[str, int, str] = FontTuple(
            "Calibri",
            13 if USER_OS == "darwin" else 11,
            "normal",
        ),
        max_undos: int = 30,
        column_drag_and_drop_perform: bool = True,
        row_drag_and_drop_perform: bool = True,
        empty_horizontal: int = 50,
        empty_vertical: int = 50,
        horizontal_grid_to_end_of_window: bool = False,
        vertical_grid_to_end_of_window: bool = False,
        show_vertical_grid: bool = True,
        show_horizontal_grid: bool = True,
        display_selected_fg_over_highlights: bool = False,
        show_selected_cells_border: bool = True,
        edit_cell_tab: Literal["right", "down", ""] = "right",
        edit_cell_return: Literal["right", "down", ""] = "down",
        editor_del_key: Literal["forward", "backward", ""] = "forward",
        treeview: bool = False,
        treeview_indent: str | int = "2",
        rounded_boxes: bool = True,
        alternate_color: str = "",
        allow_cell_overflow: bool = False,
        # "" no wrap, "w" word wrap, "c" char wrap
        table_wrap: Literal["", "w", "c"] = "c",
        index_wrap: Literal["", "w", "c"] = "c",
        header_wrap: Literal["", "w", "c"] = "c",
        sort_key: Callable = natural_sort_key,
        tooltips: bool = False,
        user_can_create_notes: bool = False,
        note_corners: bool = False,
        tooltip_width: int = 210,
        tooltip_height: int = 210,
        tooltip_hover_delay: int = 1200,
        # colors
        outline_thickness: int = 0,
        theme: str = "light blue",
        outline_color: str = theme_light_blue["outline_color"],
        frame_bg: str = theme_light_blue["table_bg"],
        popup_menu_fg: str = theme_light_blue["popup_menu_fg"],
        popup_menu_bg: str = theme_light_blue["popup_menu_bg"],
        popup_menu_highlight_bg: str = theme_light_blue["popup_menu_highlight_bg"],
        popup_menu_highlight_fg: str = theme_light_blue["popup_menu_highlight_fg"],
        table_grid_fg: str = theme_light_blue["table_grid_fg"],
        table_bg: str = theme_light_blue["table_bg"],
        table_fg: str = theme_light_blue["table_fg"],
        table_editor_bg: str = theme_light_blue["table_editor_bg"],
        table_editor_fg: str = theme_light_blue["table_editor_fg"],
        table_editor_select_bg: str = theme_light_blue["table_editor_select_bg"],
        table_editor_select_fg: str = theme_light_blue["table_editor_select_fg"],
        table_selected_box_cells_fg: str = theme_light_blue["table_selected_box_cells_fg"],
        table_selected_box_rows_fg: str = theme_light_blue["table_selected_box_rows_fg"],
        table_selected_box_columns_fg: str = theme_light_blue["table_selected_box_columns_fg"],
        table_selected_cells_border_fg: str = theme_light_blue["table_selected_cells_border_fg"],
        table_selected_cells_bg: str = theme_light_blue["table_selected_cells_bg"],
        table_selected_cells_fg: str = theme_light_blue["table_selected_cells_fg"],
        table_selected_rows_border_fg: str = theme_light_blue["table_selected_rows_border_fg"],
        table_selected_rows_bg: str = theme_light_blue["table_selected_rows_bg"],
        table_selected_rows_fg: str = theme_light_blue["table_selected_rows_fg"],
        table_selected_columns_border_fg: str = theme_light_blue["table_selected_columns_border_fg"],
        table_selected_columns_bg: str = theme_light_blue["table_selected_columns_bg"],
        table_selected_columns_fg: str = theme_light_blue["table_selected_columns_fg"],
        resizing_line_fg: str = theme_light_blue["resizing_line_fg"],
        drag_and_drop_bg: str = theme_light_blue["drag_and_drop_bg"],
        index_bg: str = theme_light_blue["index_bg"],
        index_border_fg: str = theme_light_blue["index_border_fg"],
        index_grid_fg: str = theme_light_blue["index_grid_fg"],
        index_fg: str = theme_light_blue["index_fg"],
        index_editor_bg: str = theme_light_blue["index_editor_bg"],
        index_editor_fg: str = theme_light_blue["index_editor_fg"],
        index_editor_select_bg: str = theme_light_blue["index_editor_select_bg"],
        index_editor_select_fg: str = theme_light_blue["index_editor_select_fg"],
        index_selected_cells_bg: str = theme_light_blue["index_selected_cells_bg"],
        index_selected_cells_fg: str = theme_light_blue["index_selected_cells_fg"],
        index_selected_rows_bg: str = theme_light_blue["index_selected_rows_bg"],
        index_selected_rows_fg: str = theme_light_blue["index_selected_rows_fg"],
        index_hidden_rows_expander_bg: str = theme_light_blue["index_hidden_rows_expander_bg"],
        header_bg: str = theme_light_blue["header_bg"],
        header_border_fg: str = theme_light_blue["header_border_fg"],
        header_grid_fg: str = theme_light_blue["header_grid_fg"],
        header_fg: str = theme_light_blue["header_fg"],
        header_editor_bg: str = theme_light_blue["header_editor_bg"],
        header_editor_fg: str = theme_light_blue["header_editor_fg"],
        header_editor_select_bg: str = theme_light_blue["header_editor_select_bg"],
        header_editor_select_fg: str = theme_light_blue["header_editor_select_fg"],
        header_selected_cells_bg: str = theme_light_blue["header_selected_cells_bg"],
        header_selected_cells_fg: str = theme_light_blue["header_selected_cells_fg"],
        header_selected_columns_bg: str = theme_light_blue["header_selected_columns_bg"],
        header_selected_columns_fg: str = theme_light_blue["header_selected_columns_fg"],
        header_hidden_columns_expander_bg: str = theme_light_blue["header_hidden_columns_expander_bg"],
        top_left_bg: str = theme_light_blue["top_left_bg"],
        top_left_fg: str = theme_light_blue["top_left_fg"],
        top_left_fg_highlight: str = theme_light_blue["top_left_fg_highlight"],
        vertical_scroll_background: str = theme_light_blue["vertical_scroll_background"],
        horizontal_scroll_background: str = theme_light_blue["horizontal_scroll_background"],
        vertical_scroll_troughcolor: str = theme_light_blue["vertical_scroll_troughcolor"],
        horizontal_scroll_troughcolor: str = theme_light_blue["horizontal_scroll_troughcolor"],
        vertical_scroll_lightcolor: str = theme_light_blue["vertical_scroll_lightcolor"],
        horizontal_scroll_lightcolor: str = theme_light_blue["horizontal_scroll_lightcolor"],
        vertical_scroll_darkcolor: str = theme_light_blue["vertical_scroll_darkcolor"],
        horizontal_scroll_darkcolor: str = theme_light_blue["horizontal_scroll_darkcolor"],
        vertical_scroll_relief: str = theme_light_blue["vertical_scroll_relief"],
        horizontal_scroll_relief: str = theme_light_blue["horizontal_scroll_relief"],
        vertical_scroll_troughrelief: str = theme_light_blue["vertical_scroll_troughrelief"],
        horizontal_scroll_troughrelief: str = theme_light_blue["horizontal_scroll_troughrelief"],
        vertical_scroll_bordercolor: str = theme_light_blue["vertical_scroll_bordercolor"],
        horizontal_scroll_bordercolor: str = theme_light_blue["horizontal_scroll_bordercolor"],
        vertical_scroll_active_bg: str = theme_light_blue["vertical_scroll_active_bg"],
        horizontal_scroll_active_bg: str = theme_light_blue["horizontal_scroll_active_bg"],
        vertical_scroll_not_active_bg: str = theme_light_blue["vertical_scroll_not_active_bg"],
        horizontal_scroll_not_active_bg: str = theme_light_blue["horizontal_scroll_not_active_bg"],
        vertical_scroll_pressed_bg: str = theme_light_blue["vertical_scroll_pressed_bg"],
        horizontal_scroll_pressed_bg: str = theme_light_blue["horizontal_scroll_pressed_bg"],
        vertical_scroll_active_fg: str = theme_light_blue["vertical_scroll_active_fg"],
        horizontal_scroll_active_fg: str = theme_light_blue["horizontal_scroll_active_fg"],
        vertical_scroll_not_active_fg: str = theme_light_blue["vertical_scroll_not_active_fg"],
        horizontal_scroll_not_active_fg: str = theme_light_blue["horizontal_scroll_not_active_fg"],
        vertical_scroll_pressed_fg: str = theme_light_blue["vertical_scroll_pressed_fg"],
        horizontal_scroll_pressed_fg: str = theme_light_blue["horizontal_scroll_pressed_fg"],
        vertical_scroll_borderwidth: int = 1,
        horizontal_scroll_borderwidth: int = 1,
        vertical_scroll_gripcount: int = 0,
        horizontal_scroll_gripcount: int = 0,
        scrollbar_theme_inheritance: str = "default",
        scrollbar_show_arrows: bool = True,
        # changing the arrowsize (width) of the scrollbars
        # is not working with 'default' theme
        # use 'clam' theme instead if you want to change the width
        vertical_scroll_arrowsize: str | int = "",
        horizontal_scroll_arrowsize: str | int = "",
        # backwards compatibility
        column_width: int | None = None,
        header_height: str | int | None = None,
        row_height: str | int | None = None,
        row_index_width: int | None = None,
        expand_sheet_if_paste_too_big: bool | None = None,
        **kwargs,
    ) -> None:
        super().__init__(
            parent,
            background=frame_bg,
            highlightthickness=outline_thickness,
            highlightbackground=outline_color,
            highlightcolor=outline_color,
        )
        self.unique_id = f"{default_timer()}{self.winfo_id()}".replace(".", "")
        self._startup_complete = False
        self.ops = new_sheet_options()
        if column_width is not None:
            default_column_width = column_width
        if header_height is not None:
            default_header_height = header_height
        if row_height is not None:
            default_row_height = row_height
        if row_index_width is not None:
            default_row_index_width = row_index_width
        if expand_sheet_if_paste_too_big is not None:
            paste_can_expand_x = expand_sheet_if_paste_too_big
            paste_can_expand_y = expand_sheet_if_paste_too_big
        if treeview:
            index_align = "w"
            auto_resize_row_index = True
            paste_can_expand_y = False
        for k, v in chain(locals().items(), kwargs.items()):
            if (xk := backwards_compatibility_keys.get(k, k)) in self.ops and v != self.ops[xk]:
                self.ops[xk] = v
        self.ops.from_clipboard_delimiters = (
            from_clipboard_delimiters
            if isinstance(from_clipboard_delimiters, str)
            else "".join(from_clipboard_delimiters)
        )
        self.PAR = parent
        self.name = name
        self.last_event_data = EventDataDict()
        self.bound_events = DotDict({k: [] for k in emitted_events})
        self._dropdown_cls = Dropdown
        self.after_redraw_id = None
        self.after_redraw_time_ms = after_redraw_time_ms
        self.named_span_id = 0
        if width is not None or height is not None:
            self.grid_propagate(0)
        if width is not None:
            self.config(width=width)
        if height is not None:
            self.config(height=height)
        if width is not None and height is None:
            self.config(height=300)
        if height is not None and width is None:
            self.config(width=350)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(1, weight=1)
        self.RI: RowIndex = RowIndex(
            parent=self,
            row_index_align=(
                convert_align(row_index_align) if row_index_align is not None else convert_align(index_align)
            ),
        )
        self.CH: ColumnHeaders = ColumnHeaders(
            parent=self,
            header_align=convert_align(header_align),
        )
        self.MT: MainTable = MainTable(
            parent=self,
            column_headers_canvas=self.CH,
            row_index_canvas=self.RI,
            show_index=show_row_index,
            show_header=show_header,
            headers=headers,
            header=header,
            data_reference=data if data_reference is None else data_reference,
            total_cols=total_columns,
            total_rows=total_rows,
            row_index=row_index,
            index=index,
            zoom=zoom,
            align=convert_align(align),
            displayed_columns=displayed_columns,
            all_columns_displayed=all_columns_displayed,
            displayed_rows=displayed_rows,
            all_rows_displayed=all_rows_displayed,
        )
        self.TL: TopLeftRectangle = TopLeftRectangle(
            parent=self,
            main_canvas=self.MT,
            row_index_canvas=self.RI,
            header_canvas=self.CH,
        )
        style = ttk.Style()
        for orientation in ("Vertical", "Horizontal"):
            style.element_create(
                f"Sheet{self.unique_id}.{orientation}.TScrollbar.trough",
                "from",
                scrollbar_theme_inheritance,
            )
            style.element_create(
                f"Sheet{self.unique_id}.{orientation}.TScrollbar.thumb",
                "from",
                scrollbar_theme_inheritance,
            )
            style.element_create(
                f"Sheet{self.unique_id}.{orientation}.TScrollbar.grip",
                "from",
                scrollbar_theme_inheritance,
            )
            if not scrollbar_show_arrows:
                style.layout(
                    f"Sheet{self.unique_id}.{orientation}.TScrollbar",
                    [
                        (
                            f"Sheet{self.unique_id}.{orientation}.TScrollbar.trough",
                            {
                                "children": [
                                    (
                                        f"Sheet{self.unique_id}.{orientation}.TScrollbar.thumb",
                                        {
                                            "unit": "1",
                                            "children": [
                                                (
                                                    f"Sheet{self.unique_id}.{orientation}.TScrollbar.grip",
                                                    {"sticky": ""},
                                                )
                                            ],
                                            "sticky": "nswe",
                                        },
                                    )
                                ],
                                "sticky": "ns" if orientation == "Vertical" else "ew",
                            },
                        )
                    ],
                )
        self.set_scrollbar_options()
        self.yscroll = ttk.Scrollbar(
            self,
            command=self.MT._yscrollbar,
            orient="vertical",
            style=f"Sheet{self.unique_id}.Vertical.TScrollbar",
        )
        self.MT["yscrollcommand"] = self.yscroll.set
        self.xscroll = ttk.Scrollbar(
            self,
            command=self.MT._xscrollbar,
            orient="horizontal",
            style=f"Sheet{self.unique_id}.Horizontal.TScrollbar",
        )
        self.MT["xscrollcommand"] = self.xscroll.set
        self.show()
        if show_top_left is False or (show_top_left is None and (not show_row_index or not show_header)):
            self.hide("top_left")
        if not show_row_index:
            self.hide("row_index")
        if not show_header:
            self.hide("header")
        if not show_table:
            self.hide("table")
        if not show_x_scrollbar:
            self.hide("x_scrollbar")
        if not show_y_scrollbar:
            self.hide("y_scrollbar")
        self.update_idletasks()
        self.MT.update_idletasks()
        self.RI.update_idletasks()
        self.CH.update_idletasks()
        if theme != "light blue":
            self.change_theme(theme)
            for k, v in locals().items():
                if k in theme_light_blue and v != theme_light_blue[k]:
                    self.set_options(**{k: v})
        if set_all_heights_and_widths:
            self.set_all_cell_sizes_to_text()
        if startup_select is not None:
            try:
                if startup_select[-1] == "cells":
                    self.MT.create_selection_box(*startup_select)
                    self.see(startup_select[0], startup_select[1])
                elif startup_select[-1] == "rows":
                    self.MT.create_selection_box(
                        startup_select[0],
                        0,
                        startup_select[1],
                        len(self.MT.col_positions) - 1,
                        "rows",
                    )
                    self.see(startup_select[0], 0)
                elif startup_select[-1] in ("cols", "columns"):
                    self.MT.create_selection_box(
                        0,
                        startup_select[0],
                        len(self.MT.row_positions) - 1,
                        startup_select[1],
                        "columns",
                    )
                    self.see(0, startup_select[0])
            except Exception:
                pass
        self.refresh()
        if startup_focus:
            self.MT.focus_set()
        self.MT.create_rc_menus()
        self.after_idle(self.startup_complete)

    def startup_complete(self, _mod: bool = True) -> bool:
        self._startup_complete = _mod
        return self._startup_complete

    # Sheet Colors

    def change_theme(self, theme: str = "light blue", redraw: bool = True) -> Sheet:
        if theme.lower() in ("light blue", "light_blue"):
            self.set_options(**theme_light_blue, redraw=False)
            self.config(bg=theme_light_blue["table_bg"])
        elif theme.lower() == "dark":
            self.set_options(**theme_dark, redraw=False)
            self.config(bg=theme_dark["table_bg"])
        elif theme.lower() in ("light green", "light_green"):
            self.set_options(**theme_light_green, redraw=False)
            self.config(bg=theme_light_green["table_bg"])
        elif theme.lower() in ("dark blue", "dark_blue"):
            self.set_options(**theme_dark_blue, redraw=False)
            self.config(bg=theme_dark_blue["table_bg"])
        elif theme.lower() in ("dark green", "dark_green"):
            self.set_options(**theme_dark_green, redraw=False)
            self.config(bg=theme_dark_green["table_bg"])
        elif theme.lower() == "black":
            self.set_options(**theme_black, redraw=False)
            self.config(bg=theme_black["table_bg"])
        self.MT.recreate_all_selection_boxes()
        self.set_refresh_timer(redraw)
        self.TL.redraw()
        return self

    # Header and Index

    def set_header_data(
        self,
        value: Any,
        c: int | None | Iterator[int] = None,
        redraw: bool = True,
    ) -> Sheet:
        if c is None:
            if not isinstance(value, int) and not is_iterable(value):
                raise ValueError(f"Argument 'value' must be non-string iterable or int, not {type(value)} type.")
            if not isinstance(value, list) and not isinstance(value, int):
                value = list(value)
            self.MT._headers = value
        elif isinstance(c, int):
            self.CH.set_cell_data(c, value)
        elif is_iterable(value) and is_iterable(c):
            for c_, v in zip(c, value):
                self.CH.set_cell_data(c_, v)
        return self.set_refresh_timer(redraw)

    def headers(
        self,
        newheaders: Any = None,
        index: None | int = None,
        reset_col_positions: bool = False,
        show_headers_if_not_sheet: bool = True,
        redraw: bool = True,
    ) -> Any:
        self.set_refresh_timer(redraw)
        return self.MT.headers(
            newheaders,
            index,
            reset_col_positions=reset_col_positions,
            show_headers_if_not_sheet=show_headers_if_not_sheet,
            redraw=False,
        )

    def set_index_data(
        self,
        value: Any,
        r: int | None | Iterator[int] = None,
        redraw: bool = True,
    ) -> Sheet:
        if r is None:
            if not isinstance(value, int) and not is_iterable(value):
                raise ValueError(f"Argument 'value' must be non-string iterable or int, not {type(value)} type.")
            if not isinstance(value, list) and not isinstance(value, int):
                value = list(value)
            self.MT._row_index = value
        elif isinstance(r, int):
            self.RI.set_cell_data(r, value)
        elif is_iterable(value) and is_iterable(r):
            for r_, v in zip(r, value):
                self.RI.set_cell_data(r_, v)
        return self.set_refresh_timer(redraw)

    def row_index(
        self,
        newindex: Any = None,
        index: None | int = None,
        reset_row_positions: bool = False,
        show_index_if_not_sheet: bool = True,
        redraw: bool = True,
    ) -> Any:
        self.set_refresh_timer(redraw)
        return self.MT.row_index(
            newindex,
            index,
            reset_row_positions=reset_row_positions,
            show_index_if_not_sheet=show_index_if_not_sheet,
            redraw=False,
        )

    # Bindings and Functionality

    def enable_bindings(self, *bindings: Binding, menu: bool = True) -> Sheet:
        self.MT.enable_bindings(bindings, menu=menu)
        return self

    def disable_bindings(self, *bindings: Binding) -> Sheet:
        self.MT.disable_bindings(bindings)
        return self

    def extra_bindings(
        self,
        bindings: ExtraBinding | Sequence[ExtraBinding] | None = None,
        func: Callable | None = None,
    ) -> Sheet:
        # bindings is None, unbind all
        if bindings is None:
            bindings = "all"
        # bindings is str, func arg is None or Callable
        if isinstance(bindings, str):
            iterable = ((bindings, func),)
        # bindings is list or tuple of strings, func arg is None or Callable
        elif is_iterable(bindings) and isinstance(bindings[0], str):
            iterable = ((b, func) for b in bindings)
        # bindings is a list or tuple of two tuples or lists
        # in this case the func arg is ignored
        # e.g. [(binding, function), (binding, function), ...]
        elif is_iterable(bindings):
            iterable = bindings

        for b, f in iterable:
            b = b.lower()
            if f is not None and b in emitted_events:
                self.bind(b, f)
            # Handle group bindings
            if b in ("all", "bind_all", "unbind_all"):
                for component, attr in ALL_BINDINGS:
                    setattr(getattr(self, component), attr, f)
            elif b in ("all_select_events", "select", "selectevents", "select_events"):
                for component, attr in SELECT_BINDINGS:
                    setattr(getattr(self, component), attr, f)
            elif b in ("all_modified_events", "sheetmodified", "sheet_modified", "modified_events", "modified"):
                for component, attr in MODIFIED_BINDINGS:
                    setattr(getattr(self, component), attr, f)
            # Handle individual bindings
            elif b in BINDING_TO_ATTR:
                component, attr = BINDING_TO_ATTR[b]
                setattr(getattr(self, component), attr, f)
        return self

    def bind(
        self,
        binding: str,
        func: Callable,
        add: str | None = None,
    ) -> Sheet:
        """
        In addition to normal use, the following special tksheet bindings can be used:
        - "<<SheetModified>>"
        - "<<SheetRedrawn>>"
        - "<<SheetSelect>>"
        - "<<Copy>>"
        - "<<Cut>>"
        - "<<Paste>>"
        - "<<Delete>>"
        - "<<Undo>>"
        - "<<Redo>>"
        - "<<SelectAll>>"

        Use `add` parameter to bind to multiple functions
        """
        if binding == "<ButtonPress-1>":
            self.MT.extra_b1_press_func = func
            self.CH.extra_b1_press_func = func
            self.RI.extra_b1_press_func = func
            self.TL.extra_b1_press_func = func
        elif binding == "<ButtonMotion-1>":
            self.MT.extra_b1_motion_func = func
            self.CH.extra_b1_motion_func = func
            self.RI.extra_b1_motion_func = func
            self.TL.extra_b1_motion_func = func
        elif binding == "<ButtonRelease-1>":
            self.MT.extra_b1_release_func = func
            self.CH.extra_b1_release_func = func
            self.RI.extra_b1_release_func = func
            self.TL.extra_b1_release_func = func
        elif binding == "<Double-Button-1>":
            self.MT.extra_double_b1_func = func
            self.CH.extra_double_b1_func = func
            self.RI.extra_double_b1_func = func
            self.TL.extra_double_b1_func = func
        elif binding == "<Motion>":
            self.MT.bind("<Motion>", func, add=add)
            self.CH.extra_motion_func = func
            self.RI.extra_motion_func = func
            self.TL.extra_motion_func = func
        elif binding in self.ops.rc_bindings:
            self.MT.extra_rc_func = func
            self.CH.extra_rc_func = func
            self.RI.extra_rc_func = func
            self.TL.extra_rc_func = func
        elif binding in emitted_events:
            if add:
                self.bound_events[binding].append(func)
            else:
                self.bound_events[binding] = [func]
        else:
            if binding == "<Enter>":
                self.MT.extra_enter_func = func
            else:
                self.MT.bind(binding, func, add=add)
            self.CH.bind(binding, func, add=add)
            self.RI.bind(binding, func, add=add)
            self.TL.bind(binding, func, add=add)
        return self

    def unbind(self, binding: str) -> Sheet:
        if binding in emitted_events:
            self.bound_events[binding] = []
        elif binding == "<ButtonPress-1>":
            self.MT.extra_b1_press_func = None
            self.CH.extra_b1_press_func = None
            self.RI.extra_b1_press_func = None
            self.TL.extra_b1_press_func = None
        elif binding == "<ButtonMotion-1>":
            self.MT.extra_b1_motion_func = None
            self.CH.extra_b1_motion_func = None
            self.RI.extra_b1_motion_func = None
            self.TL.extra_b1_motion_func = None
        elif binding == "<ButtonRelease-1>":
            self.MT.extra_b1_release_func = None
            self.CH.extra_b1_release_func = None
            self.RI.extra_b1_release_func = None
            self.TL.extra_b1_release_func = None
        elif binding == "<Double-Button-1>":
            self.MT.extra_double_b1_func = None
            self.CH.extra_double_b1_func = None
            self.RI.extra_double_b1_func = None
            self.TL.extra_double_b1_func = None
        elif binding == "<Motion>":
            self.MT.extra_motion_func = None
            self.CH.extra_motion_func = None
            self.RI.extra_motion_func = None
            self.TL.extra_motion_func = None
        elif binding in self.ops.rc_bindings:
            self.MT.extra_rc_func = None
            self.CH.extra_rc_func = None
            self.RI.extra_rc_func = None
            self.TL.extra_rc_func = None
        else:
            self.MT.unbind(binding)
            self.CH.unbind(binding)
            self.RI.unbind(binding)
            self.TL.unbind(binding)
        return self

    def edit_validation(self, func: Callable | None = None) -> Sheet:
        if not isinstance(func, (Callable, None)):
            raise ValueError("Argument must be either Callable or None.")
        self.MT.edit_validation_func = func
        return self

    def bulk_table_edit_validation(self, func: Callable | None = None) -> Sheet:
        if not isinstance(func, (Callable, None)):
            raise ValueError("Argument must be either Callable or None.")
        self.MT.bulk_table_edit_validation_func = func
        return self

    def popup_menu_add_command(
        self,
        label: str,
        func: Callable,
        table_menu: bool = True,
        index_menu: bool = True,
        header_menu: bool = True,
        empty_space_menu: bool = True,
        image: tk.PhotoImage | Literal[""] = "",
        compound: Literal["top", "bottom", "left", "right", "none"] | None = None,
        accelerator: str | None = None,
    ) -> Sheet:
        dct = {"command": func, "image": image, "accelerator": accelerator, "compound": compound}
        if table_menu:
            self.MT.extra_table_rc_menu_funcs[label] = dct
        if index_menu:
            self.MT.extra_index_rc_menu_funcs[label] = dct
        if header_menu:
            self.MT.extra_header_rc_menu_funcs[label] = dct
        if empty_space_menu:
            self.MT.extra_empty_space_rc_menu_funcs[label] = dct
        return self

    def popup_menu_del_command(self, label: str | None = None) -> Sheet:
        if label is None:
            self.MT.extra_table_rc_menu_funcs = {}
            self.MT.extra_index_rc_menu_funcs = {}
            self.MT.extra_header_rc_menu_funcs = {}
            self.MT.extra_empty_space_rc_menu_funcs = {}
        else:
            if label in self.MT.extra_table_rc_menu_funcs:
                del self.MT.extra_table_rc_menu_funcs[label]
            if label in self.MT.extra_index_rc_menu_funcs:
                del self.MT.extra_index_rc_menu_funcs[label]
            if label in self.MT.extra_header_rc_menu_funcs:
                del self.MT.extra_header_rc_menu_funcs[label]
            if label in self.MT.extra_empty_space_rc_menu_funcs:
                del self.MT.extra_empty_space_rc_menu_funcs[label]
        return self

    def basic_bindings(self, enable: bool = False) -> Sheet:
        for canvas in (self.MT, self.CH, self.RI, self.TL):
            canvas.basic_bindings(enable)
        return self

    def cut(self, event: Any = None, validation: bool = True) -> None | EventDataDict:
        return self.MT.ctrl_x(event, validation)

    def copy(self, event: Any = None) -> None | EventDataDict:
        return self.MT.ctrl_c(event)

    def paste(self, event: Any = None, validation: bool = True) -> None | EventDataDict:
        return self.MT.ctrl_v(event, validation)

    def delete(self, event: Any = None, validation: bool = True) -> None | EventDataDict:
        return self.MT.delete_key(event, validation)

    def undo(self, event: Any = None) -> None | EventDataDict:
        return self.MT.undo(event)

    def redo(self, event: Any = None) -> None | EventDataDict:
        return self.MT.redo(event)

    def has_focus(self) -> bool:
        """
        Check if any Sheet widgets have focus
        Includes child widgets such as scroll bars
        Returns bool
        """
        try:
            widget = self.focus_get()
            return widget == self or any(widget == c for c in self.children.values())
        except Exception:
            return False

    def focus_set(
        self,
        canvas: Literal[
            "table",
            "header",
            "row_index",
            "index",
            "topleft",
            "top_left",
        ] = "table",
    ) -> Sheet:
        if canvas == "table":
            self.MT.focus_set()
        elif canvas == "header":
            self.CH.focus_set()
        elif canvas in ("row_index", "index"):
            self.RI.focus_set()
        elif canvas in ("topleft", "top_left"):
            self.TL.focus_set()
        return self

    def zoom_in(self) -> Sheet:
        self.MT.zoom_in()
        return self

    def zoom_out(self) -> Sheet:
        self.MT.zoom_out()
        return self

    @property
    def event(self) -> EventDataDict:
        return self.last_event_data

    # Span objects

    def span(
        self,
        *key: CreateSpanTypes,
        type_: str = "",
        name: str = "",
        table: bool = True,
        index: bool = False,
        header: bool = False,
        tdisp: bool = False,
        idisp: bool = True,
        hdisp: bool = True,
        transposed: bool = False,
        ndim: int = 0,
        convert: Callable | None = None,
        undo: bool = True,
        emit_event: bool = False,
        widget: Any = None,
        expand: None | str = None,
        formatter_options: dict | None = None,
        **kwargs,
    ) -> Span:
        """
        Create a span / get an existing span by name
        Returns the created span
        """
        if name and name in self.MT.named_spans:
            return self.MT.named_spans[name]
        elif not name:
            name = num2alpha(self.named_span_id)
            self.named_span_id += 1
            while name in self.MT.named_spans:
                name = num2alpha(self.named_span_id)
                self.named_span_id += 1
        span = self.span_from_key(*key)
        span.name = name
        if expand is not None:
            span.expand(expand)
        if isinstance(formatter_options, dict):
            span.type_ = "format"
            span.kwargs = {"formatter": None, **formatter_options}
        else:
            span.type_ = type_.lower()
            span.kwargs = kwargs
        span.table = table
        span.header = header
        span.index = index
        span.tdisp = tdisp
        span.idisp = idisp
        span.hdisp = hdisp
        span.transposed = transposed
        span.ndim = ndim
        span.convert = convert
        span.undo = undo
        span.emit_event = emit_event
        span.widget = self if widget is None else widget
        return span

    # Named Spans

    def named_span(self, span: Span) -> Span:
        if span.name in self.MT.named_spans:
            raise ValueError(f"Span '{span.name}' already exists.")
        if not span.name:
            raise ValueError("Span must have a name.")
        if span.type_ not in named_span_types:
            raise ValueError(f"Span 'type_' must be one of the following: {', '.join(named_span_types)}.")
        self.MT.named_spans[span.name] = span
        self.create_options_from_span(span)
        return span

    def create_options_from_span(self, span: Span, set_data: bool = True) -> Sheet:
        if span.type_ == "format":
            self.format(span, set_data=set_data, **span.kwargs)
        else:
            getattr(self, span.type_)(span, **span.kwargs)
        return self

    def del_named_span(self, name: str) -> Sheet:
        if name not in self.MT.named_spans:
            raise ValueError(f"Span '{name}' does not exist.")
        span = self.MT.named_spans[name]
        type_ = span.type_
        from_r, from_c, upto_r, upto_c = self.MT.named_span_coords(name)
        totalrows = self.MT.get_max_row_idx() + 1
        totalcols = self.MT.get_max_column_idx() + 1
        rng_from_r = 0 if from_r is None else from_r
        rng_from_c = 0 if from_c is None else from_c
        rng_upto_r = totalrows if upto_r is None else upto_r
        rng_upto_c = totalcols if upto_c is None else upto_c
        if span["header"]:
            del_named_span_options(
                self.CH.cell_options,
                range(rng_from_c, rng_upto_c),
                type_,
            )
        if span["index"]:
            del_named_span_options(
                self.RI.cell_options,
                range(rng_from_r, rng_upto_r),
                type_,
            )
        # col options
        if from_r is None:
            del_named_span_options(
                self.MT.col_options,
                range(rng_from_c, rng_upto_c),
                type_,
            )
        # row options
        elif from_c is None:
            del_named_span_options(
                self.MT.row_options,
                range(rng_from_r, rng_upto_r),
                type_,
            )
        # cell options
        elif isinstance(from_r, int) and isinstance(from_c, int):
            del_named_span_options_nested(
                self.MT.cell_options,
                range(rng_from_r, rng_upto_r),
                range(rng_from_c, rng_upto_c),
                type_,
            )
        del self.MT.named_spans[name]
        return self

    def set_named_spans(self, named_spans: None | dict[str, Span] = None) -> Sheet:
        if named_spans is None:
            for name in self.MT.named_spans:
                self.del_named_span(name)
            named_spans = {}
        self.MT.named_spans = named_spans
        return self

    def get_named_span(self, name: str) -> dict[str, Span]:
        return self.MT.named_spans[name]

    def get_named_spans(self) -> dict[str, Span]:
        return self.MT.named_spans

    # Getting Sheet Data

    def __getitem__(
        self,
        *key: CreateSpanTypes,
    ) -> Span:
        return self.span_from_key(*key)

    def span_from_key(self, *key: CreateSpanTypes) -> None | Span:
        if not key:
            key = (None, None, None, None)
        span = key_to_span(key if len(key) != 1 else key[0], self.MT.named_spans, self)
        if isinstance(span, str):
            raise ValueError(span)
        return span

    def ranges_from_span(self, span: Span) -> tuple[Generator, Generator]:
        return span_ranges(
            span,
            totalrows=self.MT.total_data_rows,
            totalcols=self.MT.total_data_cols,
        )

    def get_data(self, *key: CreateSpanTypes) -> Any:
        """
        e.g. retrieves entire table as pandas dataframe
        sheet["A1"].expand().options(convert=pandas.DataFrame).data

        must deal with
        - format
        - transpose
        - ndim
        - index
        - header
        - tdisp
        - idisp
        - hdisp
        - convert

        format
        - if span.type_ == "format" and span.kwargs
        - does not format the sheets internal data, only data being returned
        - formats the data before retrieval with specified formatter function e.g.
        - format=int_formatter() -> returns kwargs, span["kwargs"] now has formatter kwargs
        - uses format_data(value=cell, **span["kwargs"]) on every table cell returned

        tranpose
        - make sublists become columns rather than rows
        - no effect on single cell

        ndim
        - defaults to None
        - when None it will follow rules:
            - if single cell in a list of lists return single cell
            - if single sublist in a list of lists then return single list
        - when ndim == 1 force single list
        - when ndim == 2 force list of lists

        table
        - gets table data, if false table data will not be included

        index
        - gets index data, is at the beginning of each row normally
          becomes its own column if transposed

        header
        - get header data, is its own row normally
          goes at the top of each column if transposed

        tdisp
        - if True gets displayed sheet values instead of actual data

        idisp
        - if True gets displayed index values instead of actual data

        hdisp
        - if True gets displayed header values instead of actual data

        convert
        - instead of returning data normally it returns
          convert(data) - sends the data to an optional convert function

        """
        span = self.span_from_key(*key)
        rows, cols = self.ranges_from_span(span)
        table, index, header = span.table, span.index, span.header
        fmt_kw = span.kwargs if span.type_ == "format" and span.kwargs else None
        t_data = partial(self.MT.get_cell_data, get_displayed=True) if span.tdisp else self.MT.get_cell_data
        i_data = self.RI.cell_str if span.idisp else self.RI.get_cell_data
        h_data = self.CH.cell_str if span.hdisp else self.CH.get_cell_data
        res = []
        if span.transposed:
            # Index row (first row when transposed)
            if index:
                index_row = [""] if header and table else []
                index_row.extend(i_data(r) for r in rows)
                res.append(index_row)
            # Header and/or table data as columns
            if header:
                for c in cols:
                    col = [h_data(c)]
                    if table:
                        col.extend(t_data(r, c, fmt_kw=fmt_kw) for r in rows)
                    res.append(col)
            elif table:
                res.extend([t_data(r, c, fmt_kw=fmt_kw) for r in rows] for c in cols)
        else:
            # Header row
            if header:
                header_row = [""] if index and table else []
                header_row.extend(h_data(c) for c in cols)
                res.append(header_row)
            # Index and/or table data as rows
            if index:
                for r in rows:
                    row = [i_data(r)]
                    if table:
                        row.extend(t_data(r, c, fmt_kw=fmt_kw) for c in cols)
                    res.append(row)
            elif table:
                res.extend([t_data(r, c, fmt_kw=fmt_kw) for c in cols] for r in rows)

        if not span.ndim:
            # it's a cell
            if len(res) == 1 and len(res[0]) == 1:
                res = res[0][0]
            # it's a single sublist
            elif len(res) == 1:
                res = res[0]
            # retrieving a list of index cells or
            elif (index and not span.transposed and not table and not header) or (
                # it's a column that's spread across sublists
                table and res and not span.transposed and len(res[0]) == 1 and len(res[-1]) == 1
            ):
                res = list(chain.from_iterable(res))
        elif span.ndim == 1:
            # flatten sublists
            res = res[0] if len(res) == 1 and len(res[0]) == 1 else list(chain.from_iterable(res))
        # if span.ndim == 2 res keeps its current dimensions as a list of lists

        if span.convert is None:
            return res
        else:
            return span.convert(res)

    def get_total_rows(self, include_index: bool = False) -> int:
        return self.MT.total_data_rows(include_index=include_index)

    def get_total_columns(self, include_header: bool = False) -> int:
        return self.MT.total_data_cols(include_header=include_header)

    def get_value_for_empty_cell(
        self,
        r: int,
        c: int,
        r_ops: bool = True,
        c_ops: bool = True,
    ) -> Any:
        return self.MT.get_value_for_empty_cell(r, c, r_ops, c_ops)

    @property
    def data(self) -> Sequence[Sequence[Any]]:
        return self.MT.data

    def __iter__(self) -> Iterator[list[Any] | tuple[Any]]:
        return self.MT.data.__iter__()

    def __reversed__(self) -> Iterator[list[Any] | tuple[Any]]:
        return reversed(self.MT.data)

    def __contains__(self, key: Any) -> bool:
        if isinstance(key, (list, tuple)):
            return key in self.MT.data
        return any(key in row for row in self.MT.data)

    # Setting Sheet Data

    def reset(
        self,
        table: bool = True,
        header: bool = True,
        index: bool = True,
        row_heights: bool = True,
        column_widths: bool = True,
        cell_options: bool = True,
        tags: bool = True,
        undo_stack: bool = True,
        selections: bool = True,
        sheet_options: bool = False,
        displayed_rows: bool = True,
        displayed_columns: bool = True,
        tree: bool = True,
        redraw: bool = True,
    ) -> Sheet:
        if table:
            self.MT.hide_text_editor_and_dropdown(redraw=False)
            self.MT.data = []
        if header:
            self.CH.hide_text_editor_and_dropdown(redraw=False)
            self.MT._headers = []
        if index:
            self.RI.hide_text_editor_and_dropdown(redraw=False)
            self.MT._row_index = []
        if displayed_columns:
            self.MT.displayed_columns = []
            self.MT.all_columns_displayed = True
        if displayed_rows:
            self.MT.displayed_rows = []
            self.MT.all_rows_displayed = True
        if selections:
            self.MT.deselect(redraw=False)
        if row_heights:
            self.MT.saved_row_heights = {}
            self.MT.set_row_positions([])
        if column_widths:
            self.MT.saved_column_widths = {}
            self.MT.set_col_positions([])
        if cell_options:
            self.reset_all_options()
        if tags:
            self.MT.reset_tags()
        if undo_stack:
            self.reset_undos()
        if sheet_options:
            self.ops = new_sheet_options()
            self.change_theme(redraw=False)
        if tree:
            self.RI.tree_reset()
        if tree or row_heights or column_widths:
            self.MT.hide_dropdown_editor_all_canvases()
        return self.set_refresh_timer(redraw)

    def set_sheet_data(
        self,
        data: list | tuple | None = None,
        reset_col_positions: bool = True,
        reset_row_positions: bool = True,
        redraw: bool = True,
        verify: bool = False,
        reset_highlights: bool = False,
        keep_formatting: bool = True,
        delete_options: bool = False,
    ) -> Any:
        if data is None:
            data = []
        if verify and (not isinstance(data, list) or not all(isinstance(row, list) for row in data)):
            raise ValueError("Data argument must be a list of lists, sublists being rows")
        if delete_options:
            self.reset_all_options()
        elif reset_highlights:
            self.dehighlight_all()
        return self.MT.data_reference(
            data,
            reset_col_positions,
            reset_row_positions,
            redraw,
            return_id=False,
            keep_formatting=keep_formatting,
        )

    @data.setter
    def data(self, value: list[list[Any]]) -> None:
        self.data_reference(value)

    def new_tksheet_event(self) -> EventDataDict:
        return event_dict(
            name="",
            sheet=self.name,
            widget=self,
            selected=self.MT.selected,
        )

    def set_data(
        self,
        *key: CreateSpanTypes,
        data: Any = None,
        undo: bool | None = None,
        emit_event: bool | None = None,
        redraw: bool = True,
        event_data: EventDataDict | None = None,
    ) -> EventDataDict:
        """
        e.g.
        df = pandas.DataFrame([[1, 2, 3], [4, 5, 6]])
        sheet["A1"] = df.values.tolist()

        just uses the spans from_r, from_c values

        'data' parameter could be:
        - list of lists
        - list of values
        - not a list or tuple
            - can have all three of the below set at the same time
            or just one or two:
                - if span.table then sets table cell
                - if span.index then sets index cell
                - if span.header than sets header cell

        expands sheet if required but can only undo any added
        displayed row/column positions, not expanded MT.data list

        format
        - if span.type_ == "format" and span.kwargs
        - formats the data before setting with specified formatter function e.g.
        - format=int_formatter() -> returns kwargs, span["kwargs"] now has formatter kwargs
        - uses format_data(value=cell, **span["kwargs"]) on every set cell

        transposed
        - switches range orientation
        - a single list will go to row without transposed
        - multi lists will go to rows without transposed
        - with transposed a single list will go to column
        - with transposed multi lists will go to columns

        header
        - assumes there's a header, sets header cell data as well

        index
        - assumes there's an index, sets index cell data as well

        undo
        - if True adds to undo stack and if undo is enabled
          for end user they can undo/redo the change

        in the table comments below -
        - t stands for table
        - i stands for index
        - h stands for header

        """
        span = self.span_from_key(*key)
        startr, startc = span_froms(span)
        table, index, header = span.table, span.index, span.header
        fmt_kw = span.kwargs if span.type_ == "format" and span.kwargs else None
        transposed = span.transposed
        maxr, maxc = startr, startc
        if event_data is None:
            event_data = event_dict(
                name="edit_table",
                sheet=self.name,
                widget=self,
                selected=self.MT.selected,
            )
        set_t = self.event_data_set_table_cell
        set_i, set_h = self.event_data_set_index_cell, self.event_data_set_header_cell
        istart = 1 if index else 0
        hstart = 1 if header else 0
        # data is list
        if isinstance(data, (list, tuple)):
            if not data:
                return
            # data is a list of lists
            if isinstance(data[0], (list, tuple)):
                if transposed:
                    if table:
                        """
                        - sublists are columns
                            1  2  3
                        A [[-, i, i],
                        B  [h, t, t],
                        C  [h, t, t],]
                        """
                        if index:
                            for r, v in enumerate(islice(data[0], hstart, None)):
                                maxr = r
                                event_data = set_i(r, v, event_data)
                        if header:
                            for c, sl in enumerate(islice(data, istart, None), start=startc):
                                maxc = c
                                for v in islice(sl, 0, 1):
                                    event_data = set_h(c, v, event_data)
                        for c, sl in enumerate(islice(data, istart, None), start=startc):
                            maxc = c
                            for r, v in enumerate(islice(sl, hstart, None), start=startr):
                                event_data = set_t(r, c, v, event_data, fmt_kw)
                                maxr = r
                    elif not table:
                        if index and header:
                            """
                            - first sublist is index, rest are headers
                            [['1', '2', '3'], ['A'], ['B'], ['C']]
                            """
                            for r, v in enumerate(data[0], start=startr):
                                maxr = r
                                event_data = set_i(r, v, event_data)
                            for c, sl in enumerate(islice(data, 1, None), start=startc):
                                maxc = c
                                if sl:
                                    event_data = set_h(c, sl[0], event_data)
                        elif index:
                            """
                            ['1', '2', '3']
                            """
                            for r, v in enumerate(data, start=startr):
                                maxr = r
                                event_data = set_i(r, v, event_data)
                        elif header:
                            """
                            [['A'], ['B'], ['C']]
                            """
                            for c, sl in enumerate(data, start=startc):
                                maxc = c
                                if sl:
                                    event_data = set_h(c, sl[0], event_data)
                elif not transposed:
                    if table:
                        """
                        - sublists are rows
                            A  B  C
                        1 [[-, h, h],
                        2  [i, t, t],
                        3  [i, t, t],]
                        """
                        if index:
                            for r, sl in enumerate(islice(data, hstart, None), start=startr):
                                maxr = r
                                for v in islice(sl, 0, 1):
                                    event_data = set_i(r, v, event_data)
                        if header:
                            for c, v in enumerate(islice(data[0], istart, None)):
                                maxc = c
                                event_data = set_h(c, v, event_data)
                        for r, sl in enumerate(islice(data, hstart, None), start=startr):
                            maxr = r
                            for c, v in enumerate(islice(sl, istart, None), start=startc):
                                maxc = c
                                event_data = set_t(r, c, v, event_data, fmt_kw)
                    elif not table:
                        if index and header:
                            """
                            - first sublist is headers, rest are index values
                            [['A', 'B', 'C'], ['1'], ['2'], ['3']]
                            """
                            for c, v in enumerate(data[0], start=startc):
                                maxc = c
                                event_data = set_h(c, v, event_data)
                            for r, sl in enumerate(islice(data, 1, None), start=startr):
                                maxr = r
                                if sl:
                                    event_data = set_i(r, sl[0], event_data)
                        elif index:
                            """
                            [['1'], ['2'], ['3']]
                            """
                            for r, sl in enumerate(data, start=startr):
                                maxr = r
                                if sl:
                                    event_data = set_i(r, sl[0], event_data)
                        elif header:
                            """
                            ['A', 'B', 'C']
                            """
                            for c, v in enumerate(data, start=startc):
                                maxc = c
                                event_data = set_h(c, v, event_data)
            # data is list of values
            else:
                # setting a list of index values, ignore transposed
                if index and not table:
                    """
                     1  2  3
                    [i, i, i]
                    """
                    for r, v in enumerate(data, start=startr):
                        maxr = r
                        event_data = set_i(r, v, event_data)
                # setting a list of header values, ignore transposed
                elif header and not table:
                    """
                     A  B  C
                    [h, h, h]
                    """
                    for c, v in enumerate(data, start=startc):
                        maxc = c
                        event_data = set_h(c, v, event_data)
                # includes table values, transposed taken into account
                elif table:
                    if transposed:
                        """
                        - single list is column, span.index ignored
                            1  2  3
                        A  [h, t, t]
                        """
                        if header:
                            event_data = set_h(startc, data[0], event_data)
                        for r, v in enumerate(islice(data, hstart, None), start=startr):
                            maxr = r
                            event_data = set_t(r, startc, v, event_data, fmt_kw)
                    elif not transposed:
                        """
                        - single list is row, span.header ignored
                            A  B  C
                        1  [i, t, t]
                        """
                        if index:
                            event_data = set_i(startr, data[0], event_data)
                        for c, v in enumerate(islice(data, istart, None), start=startc):
                            maxc = c
                            event_data = set_t(startr, c, v, event_data, fmt_kw)
        # data is a value
        else:
            """
                  A
            1   t/i/h
            """
            if table:
                event_data = set_t(startr, startc, data, event_data, fmt_kw)
            if index:
                event_data = set_i(startr, data, event_data)
            if header:
                event_data = set_h(startc, data, event_data)
        # add row/column lines (positions) if required
        if self.MT.all_columns_displayed and maxc >= (ncols := len(self.MT.col_positions) - 1):
            disp_ins_col, widths_n_cols = len(self.MT.col_positions) - 1, maxc + 1 - ncols
            w = self.ops.default_column_width
            event_data = self.MT.add_columns(
                columns={},
                header={},
                column_widths=dict(zip(range(disp_ins_col, disp_ins_col + widths_n_cols), repeat(w))),
                event_data=event_data,
                create_selections=False,
            )
        if self.MT.all_rows_displayed and maxr >= (nrows := len(self.MT.row_positions) - 1):
            disp_ins_row, heights_n_rows = len(self.MT.row_positions) - 1, maxr + 1 - nrows
            h = self.MT.get_default_row_height()
            event_data = self.MT.add_rows(
                rows={},
                index={},
                row_heights=dict(zip(range(disp_ins_row, disp_ins_row + heights_n_rows), repeat(h))),
                event_data=event_data,
                create_selections=False,
            )
        if (
            event_data["cells"]["table"]
            or event_data["cells"]["index"]
            or event_data["cells"]["header"]
            or event_data["added"]["columns"]
            or event_data["added"]["rows"]
        ):
            if undo is True or (undo is None and span.undo):
                self.MT.undo_stack.append(stored_event_dict(event_data))
            self.MT.sheet_modified(
                event_data, emit_event=emit_event is True or (emit_event is None and span.emit_event)
            )
        self.set_refresh_timer(redraw)
        return event_data

    def clear(
        self,
        *key: CreateSpanTypes,
        undo: bool | None = None,
        emit_event: bool | None = None,
        redraw: bool = True,
    ) -> EventDataDict:
        span = self.span_from_key(*key)
        rows, cols = self.ranges_from_span(span)
        clear_t = self.event_data_set_table_cell
        clear_i = self.event_data_set_index_cell
        clear_h = self.event_data_set_header_cell
        quick_tval = self.MT.get_value_for_empty_cell
        quick_ival = self.RI.get_value_for_empty_cell
        quick_hval = self.CH.get_value_for_empty_cell
        table, index, header = span.table, span.index, span.header
        event_data = event_dict(
            name="edit_table",
            sheet=self.name,
            widget=self,
            selected=self.MT.selected,
        )
        if index:
            for r in rows:
                event_data = clear_i(r, quick_ival(r), event_data)
        if header:
            for c in cols:
                event_data = clear_h(c, quick_hval(c), event_data)
        if table:
            for r in rows:
                for c in cols:
                    event_data = clear_t(r, c, quick_tval(r, c), event_data)
        if event_data["cells"]["table"] or event_data["cells"]["header"] or event_data["cells"]["index"]:
            if undo is True or (undo is None and span.undo):
                self.MT.undo_stack.append(stored_event_dict(event_data))
            self.MT.sheet_modified(
                event_data, emit_event=emit_event is True or (emit_event is None and span.emit_event)
            )
        self.set_refresh_timer(redraw)
        return event_data

    def event_data_set_table_cell(
        self,
        datarn: int,
        datacn: int,
        value: Any,
        event_data: EventDataDict,
        fmt_kw: dict | None = None,
        check_readonly: bool = False,
    ) -> EventDataDict:
        if fmt_kw is not None or self.MT.input_valid_for_cell(datarn, datacn, value, check_readonly=check_readonly):
            event_data["cells"]["table"][(datarn, datacn)] = self.MT.get_cell_data(datarn, datacn)
            self.MT.set_cell_data(datarn, datacn, value, kwargs=fmt_kw)
        return event_data

    def event_data_set_index_cell(
        self,
        datarn: int,
        value: Any,
        event_data: EventDataDict,
        check_readonly: bool = False,
    ) -> EventDataDict:
        if self.RI.input_valid_for_cell(datarn, value, check_readonly=check_readonly):
            event_data["cells"]["index"][datarn] = self.RI.get_cell_data(datarn)
            self.RI.set_cell_data(datarn, value)
        return event_data

    def event_data_set_header_cell(
        self,
        datacn: int,
        value: Any,
        event_data: EventDataDict,
        check_readonly: bool = False,
    ) -> EventDataDict:
        if self.CH.input_valid_for_cell(datacn, value, check_readonly=check_readonly):
            event_data["cells"]["header"][datacn] = self.CH.get_cell_data(datacn)
            self.CH.set_cell_data(datacn, value)
        return event_data

    def insert_row(
        self,
        row: list[Any] | tuple[Any] | None = None,
        idx: str | int | None = None,
        height: int | None = None,
        row_index: bool = False,
        fill: bool = True,
        undo: bool = True,
        emit_event: bool = False,
        redraw: bool = True,
    ) -> EventDataDict:
        return self.insert_rows(
            rows=1 if row is None else [row] if isinstance(row, (list, tuple)) else row,
            idx=idx,
            heights=[height] if isinstance(height, int) else height,
            row_index=row_index,
            fill=fill,
            undo=undo,
            emit_event=emit_event,
            redraw=redraw,
        )

    def insert_column(
        self,
        column: list[Any] | tuple[Any] | None = None,
        idx: str | int | None = None,
        width: int | None = None,
        header: bool = False,
        fill: bool = True,
        undo: bool = True,
        emit_event: bool = False,
        redraw: bool = True,
    ) -> EventDataDict:
        return self.insert_columns(
            columns=1 if column is None else [column] if isinstance(column, (list, tuple)) else column,
            idx=idx,
            widths=[width] if isinstance(width, int) else width,
            headers=header,
            fill=fill,
            undo=undo,
            emit_event=emit_event,
            redraw=redraw,
        )

    def insert_rows(
        self,
        rows: Iterator[list[Any]] | int = 1,
        idx: str | int | None = None,
        heights: list[int] | tuple[int] | None = None,
        row_index: bool = False,
        fill: bool = True,
        undo: bool = True,
        emit_event: bool = False,
        create_selections: bool = True,
        add_column_widths: bool = True,
        push_ops: bool = True,
        tree: bool = True,
        redraw: bool = True,
    ) -> EventDataDict:
        total_cols = None
        if (idx := idx_param_to_int(idx)) is None:
            idx = len(self.MT.data)
        if isinstance(rows, int):
            if rows < 1:
                raise ValueError(f"rows arg must be greater than 0, not {rows}")
        elif fill:
            total_cols = self.MT.total_data_cols() if total_cols is None else total_cols
            len_check = (total_cols + 1) if row_index else total_cols
            for rn, r in enumerate(rows):
                if len_check > (lnr := len(r)):
                    r.extend(
                        self.MT.gen_empty_row_seq(
                            rn,
                            end=total_cols,
                            start=(lnr - 1) if row_index else lnr,
                            r_ops=False,
                            c_ops=False,
                        )
                    )
        event_data = self.MT.add_rows(
            *self.MT.get_args_for_add_rows(
                data_ins_row=idx,
                displayed_ins_row=idx if self.MT.all_rows_displayed else bisect_left(self.MT.displayed_rows, idx),
                rows=rows,
                heights=heights,
                row_index=row_index,
            ),
            add_col_positions=add_column_widths,
            event_data=self.MT.new_event_dict("add_rows", state=True),
            create_selections=create_selections,
            push_ops=push_ops,
            tree=tree,
        )
        if undo:
            self.MT.undo_stack.append(stored_event_dict(event_data))
        self.MT.sheet_modified(event_data, emit_event=emit_event)
        self.set_refresh_timer(redraw)
        return event_data

    def insert_columns(
        self,
        columns: Iterator[list[Any]] | int = 1,
        idx: str | int | None = None,
        widths: list[int] | tuple[int] | None = None,
        headers: bool = False,
        fill: bool = True,
        undo: bool = True,
        emit_event: bool = False,
        create_selections: bool = True,
        add_row_heights: bool = True,
        push_ops: bool = True,
        redraw: bool = True,
    ) -> EventDataDict:
        old_total = self.MT.equalize_data_row_lengths()
        total_rows = self.MT.total_data_rows()
        if (idx := idx_param_to_int(idx)) is None:
            idx = old_total
        if isinstance(columns, int):
            if columns < 1:
                raise ValueError(f"columns arg must be greater than 0, not {columns}")
        elif fill:
            len_check = (total_rows + 1) if headers else total_rows
            for i, column in enumerate(columns):
                if (col_len := len(column)) < len_check:
                    column += [
                        self.MT.get_value_for_empty_cell(
                            r,
                            idx + i,
                            r_ops=False,
                            c_ops=False,
                        )
                        for r in range((col_len - 1) if headers else col_len, total_rows)
                    ]
        event_data = self.MT.add_columns(
            *self.MT.get_args_for_add_columns(
                data_ins_col=idx,
                displayed_ins_col=idx if self.MT.all_columns_displayed else bisect_left(self.MT.displayed_columns, idx),
                columns=columns,
                widths=widths,
                headers=headers,
            ),
            add_row_positions=add_row_heights,
            event_data=self.MT.new_event_dict("add_columns", state=True),
            create_selections=create_selections,
            push_ops=push_ops,
        )
        if undo:
            self.MT.undo_stack.append(stored_event_dict(event_data))
        self.MT.sheet_modified(event_data, emit_event=emit_event)
        self.set_refresh_timer(redraw)
        return event_data

    def del_rows(
        self,
        rows: int | Iterator[int],
        data_indexes: bool = True,
        undo: bool = True,
        emit_event: bool = False,
        redraw: bool = True,
    ) -> EventDataDict:
        event_data = self.MT.delete_rows(
            rows=[rows] if isinstance(rows, int) else sorted(set(rows)),
            data_indexes=data_indexes,
            undo=undo,
            emit_event=emit_event,
        )
        self.set_refresh_timer(redraw)
        return event_data

    del_row = del_rows
    delete_row = del_rows
    delete_rows = del_rows

    def del_columns(
        self,
        columns: int | Iterator[int],
        data_indexes: bool = True,
        undo: bool = True,
        emit_event: bool = False,
        redraw: bool = True,
    ) -> EventDataDict:
        event_data = self.MT.delete_columns(
            columns=[columns] if isinstance(columns, int) else sorted(set(columns)),
            data_indexes=data_indexes,
            undo=undo,
            emit_event=emit_event,
        )
        self.set_refresh_timer(redraw)
        return event_data

    del_column = del_columns
    delete_column = del_columns
    delete_columns = del_columns

    def sheet_data_dimensions(
        self,
        total_rows: int | None = None,
        total_columns: int | None = None,
    ) -> Sheet:
        self.MT.data_dimensions(total_rows, total_columns)
        return self

    def set_sheet_data_and_display_dimensions(
        self,
        total_rows: int | None = None,
        total_columns: int | None = None,
    ) -> Sheet:
        self.sheet_display_dimensions(total_rows=total_rows, total_columns=total_columns)
        self.MT.data_dimensions(total_rows=total_rows, total_columns=total_columns)
        return self

    def total_rows(
        self,
        number: int | None = None,
        mod_positions: bool = True,
        mod_data: bool = True,
    ) -> int | Sheet:
        total_rows = self.MT.total_data_rows()
        if number is None:
            return total_rows
        if not isinstance(number, int) or number < 0:
            raise ValueError("number argument must be integer and > 0")
        if number > total_rows and mod_positions:
            self.MT.insert_row_positions(heights=number - total_rows)
        elif number < total_rows:
            if not self.MT.all_rows_displayed:
                self.MT.display_rows(enable=False, reset_row_positions=False, deselect_all=True)
            self.MT.row_positions[number + 1 :] = []
        if mod_data:
            self.MT.data_dimensions(total_rows=number)
        return self

    def total_columns(
        self,
        number: int | None = None,
        mod_positions: bool = True,
        mod_data: bool = True,
    ) -> int | Sheet:
        total_cols = self.MT.total_data_cols()
        if number is None:
            return total_cols
        if not isinstance(number, int) or number < 0:
            raise ValueError("number argument must be integer and > 0")
        if number > total_cols and mod_positions:
            self.MT.insert_col_positions(widths=number - total_cols)
        elif number < total_cols:
            if not self.MT.all_columns_displayed:
                self.MT.display_columns(enable=False, reset_col_positions=False, deselect_all=True)
            self.MT.col_positions[number + 1 :] = []
        if mod_data:
            self.MT.data_dimensions(total_columns=number)
        return self

    def move_row(self, row: int, moveto: int) -> tuple[dict[int, int], dict[int, int], EventDataDict]:
        return self.move_rows(moveto, [row])

    def move_column(self, column: int, moveto: int) -> tuple[dict[int, int], dict[int, int], EventDataDict]:
        return self.move_columns(moveto, [column])

    def move_rows(
        self,
        move_to: int | None = None,
        to_move: list[int] | None = None,
        move_data: bool = True,
        data_indexes: bool = False,
        create_selections: bool = True,
        undo: bool = True,
        emit_event: bool = False,
        move_heights: bool = True,
        event_data: EventDataDict | None = None,
        redraw: bool = True,
    ) -> tuple[dict[int, int], dict[int, int], EventDataDict]:
        if not event_data:
            event_data = self.MT.new_event_dict("move_rows", state=True)
        if not data_indexes:
            event_data.value = move_to
        data_idxs, disp_idxs, event_data = self.MT.move_rows_adjust_options_dict(
            *self.MT.get_args_for_move_rows(
                move_to=move_to,
                to_move=to_move,
                data_indexes=data_indexes,
            ),
            move_data=move_data,
            move_heights=move_heights,
            create_selections=create_selections,
            event_data=event_data,
        )
        if undo:
            self.MT.undo_stack.append(stored_event_dict(event_data))
        self.MT.sheet_modified(event_data, emit_event=emit_event)
        self.set_refresh_timer(redraw)
        return data_idxs, disp_idxs, event_data

    def move_columns(
        self,
        move_to: int | None = None,
        to_move: list[int] | None = None,
        move_data: bool = True,
        data_indexes: bool = False,
        create_selections: bool = True,
        undo: bool = True,
        emit_event: bool = False,
        move_widths: bool = True,
        event_data: EventDataDict | None = None,
        redraw: bool = True,
    ) -> tuple[dict[int, int], dict[int, int], EventDataDict]:
        if not event_data:
            event_data = self.MT.new_event_dict("move_columns", state=True)
        if not data_indexes:
            event_data.value = move_to
        data_idxs, disp_idxs, event_data = self.MT.move_columns_adjust_options_dict(
            *self.MT.get_args_for_move_columns(
                move_to=move_to,
                to_move=to_move,
                data_indexes=data_indexes,
            ),
            move_data=move_data,
            move_widths=move_widths,
            create_selections=create_selections,
            event_data=event_data,
        )
        if undo:
            self.MT.undo_stack.append(stored_event_dict(event_data))
        self.MT.sheet_modified(event_data, emit_event=emit_event)
        self.set_refresh_timer(redraw)
        return data_idxs, disp_idxs, event_data

    def mapping_move_columns(
        self,
        data_new_idxs: dict[int, int],
        disp_new_idxs: None | dict[int, int] = None,
        move_data: bool = True,
        create_selections: bool = True,
        undo: bool = True,
        emit_event: bool = False,
        redraw: bool = True,
    ) -> tuple[dict[int, int], dict[int, int], EventDataDict]:
        data_idxs, disp_idxs, event_data = self.MT.move_columns_adjust_options_dict(
            data_new_idxs=data_new_idxs,
            data_old_idxs=dict(zip(data_new_idxs.values(), data_new_idxs)),
            totalcols=None,
            disp_new_idxs=disp_new_idxs,
            move_data=move_data,
            create_selections=create_selections,
        )
        if undo:
            self.MT.undo_stack.append(stored_event_dict(event_data))
        self.MT.sheet_modified(event_data, emit_event=emit_event)
        self.set_refresh_timer(redraw)
        return data_idxs, disp_idxs, event_data

    def mapping_move_rows(
        self,
        data_new_idxs: dict[int, int],
        disp_new_idxs: None | dict[int, int] = None,
        move_data: bool = True,
        create_selections: bool = True,
        undo: bool = True,
        emit_event: bool = False,
        event_data: EventDataDict | None = None,
        redraw: bool = True,
    ) -> tuple[dict[int, int], dict[int, int], EventDataDict]:
        data_idxs, disp_idxs, event_data = self.MT.move_rows_adjust_options_dict(
            data_new_idxs=data_new_idxs,
            data_old_idxs=dict(zip(data_new_idxs.values(), data_new_idxs)),
            totalrows=None,
            disp_new_idxs=disp_new_idxs,
            move_data=move_data,
            create_selections=create_selections,
            event_data=event_data,
        )
        if undo:
            self.MT.undo_stack.append(stored_event_dict(event_data))
        self.MT.sheet_modified(event_data, emit_event=emit_event)
        self.set_refresh_timer(redraw)
        return data_idxs, disp_idxs, event_data

    def equalize_data_row_lengths(
        self,
        include_header: bool = True,
    ) -> int:
        return self.MT.equalize_data_row_lengths(
            include_header=include_header,
        )

    def full_move_rows_idxs(self, data_idxs: dict[int, int], max_idx: int | None = None) -> dict[int, int]:
        """
        Converts the dict provided by moving rows event data
        Under the keys ['moved']['rows']['data']
        Into a dict of {old index: new index} for every row
        Includes row numbers in cell options, spans, etc.
        """
        return self.MT.get_full_new_idxs(
            self.MT.get_max_row_idx() if max_idx is None else max_idx,
            data_idxs,
        )

    def full_move_columns_idxs(self, data_idxs: dict[int, int], max_idx: int | None = None) -> dict[int, int]:
        """
        Converts the dict provided by moving columns event data
        Under the keys ['moved']['columns']['data']
        Into a dict of {old index: new index} for every column
        Includes column numbers in cell options, spans, etc.
        """
        return self.MT.get_full_new_idxs(
            self.MT.get_max_column_idx() if max_idx is None else max_idx,
            data_idxs,
        )

    # Sorting

    def sort(
        self,
        *box: CreateSpanTypes,
        reverse: bool = False,
        row_wise: bool = False,
        validation: bool = True,
        key: Callable | None = None,
        undo: bool = True,
    ) -> EventDataDict:
        """
        Sort the data within specified box regions in the table.

        This method sorts the data within one or multiple box regions defined by their coordinates.
        Each box's columns are sorted independently.

        Args:
            boxes (CreateSpanTypes): A type that can create a Span.

            reverse (bool): If True, sorts in descending order. Default is False (ascending).

            row_wise (bool): if True sorts cells row-wise. Default is column-wise.

            validation (bool): If True, checks if the new cell values are valid according to any restrictions
                (e.g., dropdown validations) before applying the sort. Default is True.

            key (Callable | None): A function to extract a comparison key from each element in the columns.
                If None, a natural sorting key will be used, which can handle most Python objects, including:
                - None, booleans, numbers (int, float), datetime objects, strings, and unknown types.
                Note: Performance might be slow with the natural sort for very large datasets.

        Returns:
            EventDataDict: A dictionary containing information about the sorting event,
            including changes made to the table.

        Raises:
            ValueError: If the input boxes are not in the expected format or if the data cannot be sorted.

        Note:
            - If validation is enabled, any cell content that fails validation will not be updated.
            - Any cell options attached to cells will not be moved. Event data can be used to re-locate them.
        """
        return self.MT.sort_boxes(
            boxes={Box_nt(*self.span_from_key(*box).coords): "cells"},
            reverse=reverse,
            row_wise=row_wise,
            validation=validation,
            key=key,
            undo=undo,
        )

    def sort_rows(
        self,
        rows: Iterator[int] | Span | int | None = None,
        reverse: bool = False,
        validation: bool = True,
        key: Callable | None = None,
        undo: bool = True,
    ) -> EventDataDict:
        if isinstance(rows, Span):
            rows = rows.rows
        elif isinstance(rows, int):
            rows = (rows,)
        return self.RI._sort_rows(rows=rows, reverse=reverse, validation=validation, key=key, undo=undo)

    def sort_columns(
        self,
        columns: Iterator[int] | Span | int | None = None,
        reverse: bool = False,
        validation: bool = True,
        key: Callable | None = None,
        undo: bool = True,
    ) -> EventDataDict:
        if isinstance(columns, Span):
            columns = columns.columns
        elif isinstance(columns, int):
            columns = (columns,)
        return self.CH._sort_columns(columns=columns, reverse=reverse, validation=validation, key=key, undo=undo)

    def sort_rows_by_column(
        self,
        column: int | None = None,
        reverse: bool = False,
        key: Callable | None = None,
        undo: bool = True,
    ) -> EventDataDict:
        return self.CH._sort_rows_by_column(column=column, reverse=reverse, key=key, undo=undo)

    def sort_columns_by_row(
        self,
        row: int | None = None,
        reverse: bool = False,
        key: Callable | None = None,
        undo: bool = True,
    ) -> EventDataDict:
        return self.RI._sort_columns_by_row(row=row, reverse=reverse, key=key, undo=undo)

    # Find and Replace

    @property
    def find_open(self) -> bool:
        return self.MT.find_window.open

    def open_find(self, focus: bool = False) -> Sheet:
        self.MT.open_find_window(focus=focus)
        return self

    def close_find(self) -> Sheet:
        self.MT.close_find_window()
        return self

    def next_match(self, within: bool | None = None, find: str | None = None) -> Sheet:
        self.MT.find_next(within=within, find=find)
        return self

    def prev_match(self, within: bool | None = None, find: str | None = None) -> Sheet:
        self.MT.find_previous(within=within, find=find)
        return self

    def replace_all(self, mapping: dict[str, str], within: bool = False) -> EventDataDict:
        event_data = self.MT.new_event_dict("edit_table", boxes=self.MT.get_boxes())
        if within:
            iterable = chain.from_iterable(
                (
                    box_gen_coords(
                        *box.coords,
                        start_r=box.coords.from_r,
                        start_c=box.coords.from_c,
                        reverse=False,
                        all_rows_displayed=self.MT.all_rows_displayed,
                        all_cols_displayed=self.MT.all_columns_displayed,
                        displayed_rows=self.MT.displayed_rows,
                        displayed_cols=self.MT.displayed_columns,
                    )
                    for box in self.MT.selection_boxes.values()
                )
            )
        else:
            iterable = box_gen_coords(
                from_r=0,
                from_c=0,
                upto_r=self.MT.total_data_rows(include_index=False),
                upto_c=self.MT.total_data_cols(include_header=False),
                start_r=0,
                start_c=0,
                reverse=False,
            )
        for r, c in iterable:
            for find, replace in mapping.items():
                m = self.MT.find_match(find, r, c)
                if (
                    m
                    and not within
                    or (
                        within
                        and (self.MT.all_rows_displayed or bisect_in(self.MT.displayed_rows, r))
                        and (self.MT.all_columns_displayed or bisect_in(self.MT.displayed_columns, c))
                    )
                ):
                    current = f"{self.MT.get_cell_data(r, c, True)}"
                    new = sub(escape(find), replacer(find, replace, current), current, flags=IGNORECASE)
                    if not self.MT.edit_validation_func or (
                        self.MT.edit_validation_func
                        and (new := self.MT.edit_validation_func(mod_event_val(event_data, new, (r, c)))) is not None
                    ):
                        event_data = self.MT.event_data_set_cell(
                            r,
                            c,
                            new,
                            event_data,
                        )
        event_data = self.MT.bulk_edit_validation(event_data)
        if event_data["cells"]["table"]:
            self.MT.refresh()
            if self.MT.undo_enabled:
                self.MT.undo_stack.append(stored_event_dict(event_data))
            try_binding(self.MT.extra_end_replace_all_func, event_data)
            self.MT.sheet_modified(event_data)
        return event_data

    # Notes

    def note(self, *key: CreateSpanTypes, note: str | None = None, readonly: bool = True) -> Span:
        """
        note=None to delete notes for the span area.
        Or use a str to set notes for the span area.
        """
        span = self.span_from_key(*key)
        rows, cols = self.ranges_from_span(span)
        table, index, header = span.table, span.index, span.header
        if span.kind == "cell":
            if header:
                for c in cols:
                    mod_note(self.CH.cell_options, c, note, readonly)
            for r in rows:
                if index:
                    mod_note(self.RI.cell_options, r, note, readonly)
                if table:
                    for c in cols:
                        mod_note(self.MT.cell_options, (r, c), note, readonly)
        elif span.kind == "row":
            for r in rows:
                if index:
                    mod_note(self.RI.cell_options, r, note, readonly)
                if table:
                    mod_note(self.MT.row_options, r, note, readonly)
        elif span.kind == "column":
            for c in cols:
                if header:
                    mod_note(self.CH.cell_options, c, note, readonly)
                if table:
                    mod_note(self.MT.col_options, c, note, readonly)
        return span

    # Highlighting Cells

    def highlight(
        self,
        *key: CreateSpanTypes,
        bg: bool | None | str = False,
        fg: bool | None | str = False,
        end: bool | None = None,
        overwrite: bool = False,
        redraw: bool = True,
    ) -> Span:
        span = self.span_from_key(*key)
        if bg is False and fg is False and end is None:
            return span
        rows, cols = self.ranges_from_span(span)
        table, index, header = span.table, span.index, span.header
        if span.kind == "cell":
            if header:
                for c in cols:
                    add_highlight(self.CH.cell_options, c, bg, fg, end, overwrite)
            for r in rows:
                if index:
                    add_highlight(self.RI.cell_options, r, bg, fg, end, overwrite)
                if table:
                    for c in cols:
                        add_highlight(self.MT.cell_options, (r, c), bg, fg, end, overwrite)
        elif span.kind == "row":
            for r in rows:
                if index:
                    add_highlight(self.RI.cell_options, r, bg, fg, end, overwrite)
                if table:
                    add_highlight(self.MT.row_options, r, bg, fg, end, overwrite)
        elif span.kind == "column":
            for c in cols:
                if header:
                    add_highlight(self.CH.cell_options, c, bg, fg, end, overwrite)
                if table:
                    add_highlight(self.MT.col_options, c, bg, fg, end, overwrite)
        self.set_refresh_timer(redraw)
        return span

    def dehighlight(
        self,
        *key: CreateSpanTypes,
        redraw: bool = True,
    ) -> Span:
        span = self.span_from_key(*key)
        self.del_options_using_span(span, "highlight")
        self.set_refresh_timer(redraw)
        return span

    def dehighlight_all(
        self,
        cells: bool = True,
        rows: bool = True,
        columns: bool = True,
        header: bool = True,
        index: bool = True,
        redraw: bool = True,
    ) -> Sheet:
        if cells:
            del_from_options(self.MT.cell_options, "highlight")
        if rows:
            del_from_options(self.MT.row_options, "highlight")
        if columns:
            del_from_options(self.MT.col_options, "highlight")
        if index:
            del_from_options(self.RI.cell_options, "highlight")
        if header:
            del_from_options(self.CH.cell_options, "highlight")
        return self.set_refresh_timer(redraw)

    # Dropdown Boxes

    def dropdown(
        self,
        *key: CreateSpanTypes,
        values: list[Any] | None = None,
        edit_data: bool = True,
        set_values: dict[tuple[int, int], Any] | None = None,
        set_value: Any = None,
        state: Literal["normal", "readonly", "disabled"] = "normal",
        redraw: bool = True,
        selection_function: Callable | None = None,
        modified_function: Callable | None = None,
        search_function: Callable | None = None,
        validate_input: bool = True,
        text: None | str = None,
    ) -> Span:
        if values is None:
            values = []
        if set_values is None:
            set_values = {}
        if not search_function:
            search_function = dropdown_search_function
        v = set_value if set_value is not None else values[0] if values else ""
        kwargs = {
            "values": values,
            "state": state,
            "selection_function": selection_function,
            "modified_function": modified_function,
            "search_function": search_function,
            "validate_input": validate_input,
            "text": text,
            "default_value": set_value,
        }
        d = get_dropdown_dict(**kwargs)
        span = self.span_from_key(*key)
        rows, cols = self.ranges_from_span(span)
        table, index, header = span.table, span.index, span.header
        set_tdata = self.MT.set_cell_data
        set_idata = self.RI.set_cell_data
        set_hdata = self.CH.set_cell_data
        if index:
            for r in rows:
                self.del_index_cell_options_dropdown_and_checkbox(r)
                add_to_options(self.RI.cell_options, r, "dropdown", d)
                if edit_data:
                    set_idata(r, value=set_values.get(r, v))
        if header:
            for c in cols:
                self.del_header_cell_options_dropdown_and_checkbox(c)
                add_to_options(self.CH.cell_options, c, "dropdown", d)
                if edit_data:
                    set_hdata(c, value=set_values.get(c, v))
        if table:
            if span.kind == "cell":
                for r in rows:
                    for c in cols:
                        self.del_cell_options_dropdown_and_checkbox(r, c)
                        add_to_options(self.MT.cell_options, (r, c), "dropdown", d)
                        if edit_data:
                            set_tdata(r, c, value=set_values.get((r, c), v))
            elif span.kind == "row":
                for r in rows:
                    self.del_row_options_dropdown_and_checkbox(r)
                    add_to_options(self.MT.row_options, r, "dropdown", d)
                    if edit_data:
                        for c in cols:
                            set_tdata(r, c, value=set_values.get((r, c), v))
            elif span.kind == "column":
                for c in cols:
                    self.del_column_options_dropdown_and_checkbox(c)
                    add_to_options(self.MT.col_options, c, "dropdown", d)
                    if edit_data:
                        for r in rows:
                            set_tdata(r, c, value=set_values.get((r, c), v))
        self.set_refresh_timer(redraw)
        return span

    def del_dropdown(
        self,
        *key: CreateSpanTypes,
        redraw: bool = True,
    ) -> Span:
        span = self.span_from_key(*key)
        if span.table:
            self.MT.hide_dropdown_window()
        if span.index:
            self.RI.hide_dropdown_window()
        if span.header:
            self.CH.hide_dropdown_window()
        self.del_options_using_span(span, "dropdown")
        self.set_refresh_timer(redraw)
        return span

    def open_dropdown(self, r: int, c: int) -> Sheet:
        self.MT.open_dropdown_window(r, c)
        return self

    def close_dropdown(self, r: int | None = None, c: int | None = None) -> Sheet:
        self.MT.close_dropdown_window(r, c)
        return self

    def open_header_dropdown(self, c: int) -> Sheet:
        self.CH.open_dropdown_window(c)
        return self

    def close_header_dropdown(self, c: int | None = None) -> Sheet:
        self.CH.close_dropdown_window(c)
        return self

    def open_index_dropdown(self, r: int) -> Sheet:
        self.RI.open_dropdown_window(r)
        return self

    def close_index_dropdown(self, r: int | None = None) -> Sheet:
        self.RI.close_dropdown_window(r)
        return self

    # Check Boxes

    def checkbox(
        self,
        *key: CreateSpanTypes,
        edit_data: bool = True,
        checked: bool | None = None,
        state: Literal["normal", "disabled"] = "normal",
        redraw: bool = True,
        check_function: Callable | None = None,
        text: str = "",
    ) -> Span:
        kwargs = {
            "state": state,
            "check_function": check_function,
            "text": text,
        }
        d = get_checkbox_dict(**kwargs)
        span = self.span_from_key(*key)
        rows, cols = self.ranges_from_span(span)
        table, index, header = span.table, span.index, span.header
        set_tdata = self.MT.set_cell_data
        set_idata = self.RI.set_cell_data
        set_hdata = self.CH.set_cell_data
        if index:
            for r in rows:
                self.del_index_cell_options_dropdown_and_checkbox(r)
                add_to_options(self.RI.cell_options, r, "checkbox", d)
                if edit_data:
                    set_idata(r, checked if isinstance(checked, bool) else force_bool(self.RI.get_cell_data(r)))
        if header:
            for c in cols:
                self.del_header_cell_options_dropdown_and_checkbox(c)
                add_to_options(self.CH.cell_options, c, "checkbox", d)
                if edit_data:
                    set_hdata(c, checked if isinstance(checked, bool) else force_bool(self.CH.get_cell_data(c)))
        if table:
            if span.kind == "cell":
                for r in rows:
                    for c in cols:
                        self.MT.delete_cell_format(r, c, clear_values=False)
                        self.del_cell_options_dropdown_and_checkbox(r, c)
                        add_to_options(self.MT.cell_options, (r, c), "checkbox", d)
                        if edit_data:
                            set_tdata(
                                r, c, checked if isinstance(checked, bool) else force_bool(self.get_cell_data(r, c))
                            )
            elif span.kind == "row":
                for r in rows:
                    self.MT.delete_row_format(r, clear_values=False)
                    self.del_row_options_dropdown_and_checkbox(r)
                    add_to_options(self.MT.row_options, r, "checkbox", d)
                    if edit_data:
                        for c in cols:
                            set_tdata(
                                r, c, checked if isinstance(checked, bool) else force_bool(self.get_cell_data(r, c))
                            )
            elif span.kind == "column":
                for c in cols:
                    self.MT.delete_column_format(c, clear_values=False)
                    self.del_column_options_dropdown_and_checkbox(c)
                    add_to_options(self.MT.col_options, c, "checkbox", d)
                    if edit_data:
                        for r in rows:
                            set_tdata(
                                r, c, checked if isinstance(checked, bool) else force_bool(self.get_cell_data(r, c))
                            )
        self.set_refresh_timer(redraw)
        return span

    def del_checkbox(
        self,
        *key: CreateSpanTypes,
        redraw: bool = True,
    ) -> Span:
        span = self.span_from_key(*key)
        self.del_options_using_span(span, "checkbox")
        self.set_refresh_timer(redraw)
        return span

    def click_checkbox(
        self,
        *key: CreateSpanTypes,
        checked: bool | None = None,
        redraw: bool = True,
    ) -> Span:
        span = self.span_from_key(*key)
        rows, cols = self.ranges_from_span(span)
        table, index, header = span.table, span.index, span.header
        set_tdata = self.MT.set_cell_data
        set_idata = self.RI.set_cell_data
        set_hdata = self.CH.set_cell_data
        get_tdata = self.MT.get_cell_data
        get_idata = self.RI.get_cell_data
        get_hdata = self.CH.get_cell_data
        get_tkw = self.MT.get_cell_kwargs
        get_ikw = self.RI.get_cell_kwargs
        get_hkw = self.CH.get_cell_kwargs
        if span.kind == "cell":
            if header:
                for c in cols:
                    if get_hkw(c, key="checkbox"):
                        set_hdata(c, not get_hdata(c) if checked is None else bool(checked))
            for r in rows:
                if index and get_ikw(r, key="checkbox"):
                    set_idata(r, not get_idata(r) if checked is None else bool(checked))
                if table:
                    for c in cols:
                        if get_tkw(r, c, key="checkbox"):
                            set_tdata(r, c, not get_tdata(r, c) if checked is None else bool(checked))
        elif span.kind == "row":
            totalcols = range(self.MT.total_data_cols())
            for r in rows:
                if index and get_ikw(r, key="checkbox"):
                    set_idata(r, not get_idata(r) if checked is None else bool(checked))
                if table:
                    for c in totalcols:
                        if get_tkw(r, c, key="checkbox"):
                            set_tdata(r, c, not get_tdata(r, c) if checked is None else bool(checked))
        elif span.kind == "column":
            totalrows = range(self.MT.total_data_rows())
            for c in cols:
                if header and get_hkw(c, key="checkbox"):
                    set_hdata(c, not get_hdata(c) if checked is None else bool(checked))
                if table:
                    for r in totalrows:
                        if get_tkw(r, c, key="checkbox"):
                            set_tdata(r, c, not get_tdata(r, c) if checked is None else bool(checked))
        self.set_refresh_timer(redraw)
        return span

    # Data Formatting

    def format(
        self,
        *key: CreateSpanTypes,
        formatter_options: dict | None = None,
        formatter_class: Any = None,
        redraw: bool = True,
        set_data: bool = True,
        **kwargs,
    ) -> Span:
        if formatter_options is None:
            formatter_options = {}
        span = self.span_from_key(*key)
        rows, cols = self.ranges_from_span(span)
        kwargs = fix_format_kwargs({"formatter": formatter_class, **formatter_options, **kwargs})
        if span.kind == "cell" and span.table:
            for r in rows:
                for c in cols:
                    self.del_cell_options_checkbox(r, c)
                    add_to_options(self.MT.cell_options, (r, c), "format", kwargs)
                    if set_data:
                        self.MT.set_cell_data(
                            r,
                            c,
                            value=kwargs["value"] if "value" in kwargs else self.MT.get_cell_data(r, c),
                            kwargs=kwargs,
                        )
        elif span.kind == "row":
            for r in rows:
                self.del_row_options_checkbox(r)
                kwargs = fix_format_kwargs(kwargs)
                add_to_options(self.MT.row_options, r, "format", kwargs)
                if set_data:
                    for c in cols:
                        self.MT.set_cell_data(
                            r,
                            c,
                            value=kwargs["value"] if "value" in kwargs else self.MT.get_cell_data(r, c),
                            kwargs=kwargs,
                        )
        elif span.kind == "column":
            for c in cols:
                self.del_column_options_checkbox(c)
                kwargs = fix_format_kwargs(kwargs)
                add_to_options(self.MT.col_options, c, "format", kwargs)
                if set_data:
                    for r in rows:
                        self.MT.set_cell_data(
                            r,
                            c,
                            value=kwargs["value"] if "value" in kwargs else self.MT.get_cell_data(r, c),
                            kwargs=kwargs,
                        )
        self.set_refresh_timer(redraw)
        return span

    def del_format(
        self,
        *key: CreateSpanTypes,
        clear_values: bool = False,
        redraw: bool = True,
    ) -> Span:
        span = self.span_from_key(*key)
        rows, cols = self.ranges_from_span(span)
        if span.table:
            if span.kind == "cell":
                for r in rows:
                    for c in cols:
                        self.MT.delete_cell_format(r, c, clear_values=clear_values)
            elif span.kind == "row":
                for r in rows:
                    self.MT.delete_row_format(r, clear_values=clear_values)
            elif span.kind == "column":
                for c in cols:
                    self.MT.delete_column_format(c, clear_values=clear_values)
            self.set_refresh_timer(redraw)
        return span

    def reapply_formatting(self) -> Sheet:
        self.MT.reapply_formatting()
        return self

    def del_all_formatting(self, clear_values: bool = False) -> Sheet:
        self.MT.delete_all_formatting(clear_values=clear_values)
        return self

    delete_all_formatting = del_all_formatting

    def formatted(self, r: int, c: int) -> dict:
        return self.MT.get_cell_kwargs(r, c, key="format")

    # Readonly Cells

    def readonly(
        self,
        *key: CreateSpanTypes,
        readonly: bool = True,
    ) -> Span:
        span = self.span_from_key(*key)
        rows, cols = self.ranges_from_span(span)
        table, index, header = span.table, span.index, span.header
        if span.kind == "cell":
            if header:
                for c in cols:
                    set_readonly(self.CH.cell_options, c, readonly)
            for r in rows:
                if index:
                    set_readonly(self.RI.cell_options, r, readonly)
                if table:
                    for c in cols:
                        set_readonly(self.MT.cell_options, (r, c), readonly)
        elif span.kind == "row":
            for r in rows:
                if index:
                    set_readonly(self.RI.cell_options, r, readonly)
                if table:
                    set_readonly(self.MT.row_options, r, readonly)
        elif span.kind == "column":
            for c in cols:
                if header:
                    set_readonly(self.CH.cell_options, c, readonly)
                if table:
                    set_readonly(self.MT.col_options, c, readonly)
        return span

    # Text Font and Alignment

    def font(self, newfont: tuple[str, int, str] | None = None, **_) -> tuple[str, int, str]:
        return self.MT.set_table_font(newfont)

    table_font = font

    def header_font(self, newfont: tuple[str, int, str] | None = None) -> tuple[str, int, str]:
        return self.MT.set_header_font(newfont)

    def index_font(self, newfont: tuple[str, int, str] | None = None) -> tuple[str, int, str]:
        return self.MT.set_index_font(newfont)

    def popup_menu_font(self, newfont: tuple[str, int, str] | None = None) -> tuple[str, int, str]:
        if newfont:
            self.ops.popup_menu_font = FontTuple(*(newfont[0], int(round(newfont[1])), newfont[2]))
        return self.ops.popup_menu_font

    def table_align(
        self,
        align: str | None = None,
        redraw: bool = True,
    ) -> str | Sheet:
        if align is None:
            return self.MT.align
        elif convert_align(align):
            self.MT.align = convert_align(align)
        else:
            raise ValueError()
        return self.set_refresh_timer(redraw)

    def header_align(
        self,
        align: str | None = None,
        redraw: bool = True,
    ) -> str | Sheet:
        if align is None:
            return self.CH.align
        elif convert_align(align):
            self.CH.align = convert_align(align)
        else:
            raise ValueError(align_value_error)
        return self.set_refresh_timer(redraw)

    def row_index_align(
        self,
        align: str | None = None,
        redraw: bool = True,
    ) -> str | Sheet:
        if align is None:
            return self.RI.align
        elif convert_align(align):
            self.RI.align = convert_align(align)
        else:
            raise ValueError(align_value_error)
        return self.set_refresh_timer(redraw)

    index_align = row_index_align

    def align(
        self,
        *key: CreateSpanTypes,
        align: str | None = None,
        redraw: bool = True,
    ) -> Span:
        span = self.span_from_key(*key)
        rows, cols = self.ranges_from_span(span)
        table, index, header = span.table, span.index, span.header
        align = convert_align(align)
        if span.kind == "cell":
            if header:
                for c in cols:
                    set_align(self.CH.cell_options, c, align)
            for r in rows:
                if index:
                    set_align(self.RI.cell_options, r, align)
                if table:
                    for c in cols:
                        set_align(self.MT.cell_options, (r, c), align)
        elif span.kind == "row":
            for r in rows:
                if index:
                    set_align(self.RI.cell_options, r, align)
                if table:
                    set_align(self.MT.row_options, r, align)
        elif span.kind == "column":
            for c in cols:
                if header:
                    set_align(self.CH.cell_options, c, align)
                if table:
                    set_align(self.MT.col_options, c, align)
        self.set_refresh_timer(redraw)
        return span

    def del_align(
        self,
        *key: CreateSpanTypes,
        redraw: bool = True,
    ) -> Span:
        span = self.span_from_key(*key)
        self.del_options_using_span(span, "align")
        self.set_refresh_timer(redraw)
        return span

    # Getting Selected Cells

    def get_currently_selected(self) -> tuple[()] | Selected:
        # if self.MT.selected:
        #     return self.MT.selected._replace(type_=self.MT.selected.type_[:-1])
        return self.MT.selected

    @property
    def selected(self) -> tuple[()] | Selected:
        return self.MT.selected

    @selected.setter
    def selected(self, selected: tuple[()] | Selected) -> Sheet:
        if selected:
            self.MT.set_currently_selected(
                r=selected[0],
                c=selected[1],
                item=selected[4],
                box=selected[3],
            )
        else:
            self.MT.deselect(redraw=False)
        return self.set_refresh_timer()

    def get_selected_rows(
        self,
        get_cells: bool = False,
        get_cells_as_rows: bool = False,
        return_tuple: bool = False,
    ) -> tuple[int] | tuple[tuple[int, int]] | set[int] | set[tuple[int, int]]:
        if return_tuple:
            return tuple(self.MT.get_selected_rows(get_cells=get_cells, get_cells_as_rows=get_cells_as_rows))
        return self.MT.get_selected_rows(get_cells=get_cells, get_cells_as_rows=get_cells_as_rows)

    def get_selected_columns(
        self,
        get_cells: bool = False,
        get_cells_as_columns: bool = False,
        return_tuple: bool = False,
    ) -> tuple[int] | tuple[tuple[int, int]] | set[int] | set[tuple[int, int]]:
        if return_tuple:
            return tuple(self.MT.get_selected_cols(get_cells=get_cells, get_cells_as_cols=get_cells_as_columns))
        return self.MT.get_selected_cols(get_cells=get_cells, get_cells_as_cols=get_cells_as_columns)

    def get_selected_cells(
        self,
        get_rows: bool = False,
        get_columns: bool = False,
        sort_by_row: bool = False,
        sort_by_column: bool = False,
        reverse: bool = False,
    ) -> list[tuple[int, int]] | set[tuple[int, int]]:
        if sort_by_row and sort_by_column:
            return sorted(
                self.MT.get_selected_cells(get_rows=get_rows, get_cols=get_columns),
                key=lambda t: t,
                reverse=reverse,
            )
        elif sort_by_row:
            return sorted(
                self.MT.get_selected_cells(get_rows=get_rows, get_cols=get_columns),
                key=lambda t: t[0],
                reverse=reverse,
            )
        elif sort_by_column:
            return sorted(
                self.MT.get_selected_cells(get_rows=get_rows, get_cols=get_columns),
                key=lambda t: t[1],
                reverse=reverse,
            )
        return self.MT.get_selected_cells(get_rows=get_rows, get_cols=get_columns)

    def gen_selected_cells(
        self,
        get_rows: bool = False,
        get_columns: bool = False,
    ) -> Generator[tuple[int, int]]:
        yield from self.MT.gen_selected_cells(get_rows=get_rows, get_cols=get_columns)

    def get_all_selection_boxes(self) -> tuple[tuple[int, int, int, int]]:
        return self.MT.get_all_selection_boxes()

    def get_all_selection_boxes_with_types(self) -> list[tuple[tuple[int, int, int, int], str]]:
        return self.MT.get_all_selection_boxes_with_types()

    @property
    def boxes(self) -> list[tuple[tuple[int, int, int, int], str]]:
        return self.MT.get_all_selection_boxes_with_types()

    @boxes.setter
    def boxes(self, boxes: Sequence[tuple[tuple[int, int, int, int], str]]) -> Sheet:
        self.MT.deselect(redraw=False)
        self.MT.reselect_from_get_boxes(
            boxes={box[0] if isinstance(box[0], tuple) else tuple(box[0]): box[1] for box in boxes}
        )
        return self.set_refresh_timer()

    @property
    def canvas_boxes(self) -> dict[int, SelectionBox]:
        return self.MT.selection_boxes

    def cell_selected(
        self,
        r: int,
        c: int,
        rows: bool = False,
        columns: bool = False,
    ) -> bool:
        return self.MT.cell_selected(r, c, inc_cols=columns, inc_rows=rows)

    def row_selected(self, r: int, cells: bool = False) -> bool:
        return self.MT.row_selected(r, cells=cells)

    def column_selected(self, c: int, cells: bool = False) -> bool:
        return self.MT.col_selected(c, cells=cells)

    def anything_selected(
        self,
        exclude_columns: bool = False,
        exclude_rows: bool = False,
        exclude_cells: bool = False,
    ) -> bool:
        return bool(
            self.MT.anything_selected(
                exclude_columns=exclude_columns,
                exclude_rows=exclude_rows,
                exclude_cells=exclude_cells,
            )
        )

    def all_selected(self) -> bool:
        return self.MT.all_selected()

    def get_ctrl_x_c_boxes(
        self,
        nrows: bool = True,
    ) -> tuple[dict[tuple[int, int, int, int], str], int] | dict[tuple[int, int, int, int], str]:
        if nrows:
            return self.MT.get_ctrl_x_c_boxes()
        return self.MT.get_ctrl_x_c_boxes()[0]

    @property
    def ctrl_boxes(self) -> dict[tuple[int, int, int, int], str]:
        return self.MT.get_ctrl_x_c_boxes()[0]

    def get_selected_min_max(
        self,
    ) -> tuple[int, int, int, int] | tuple[None, None, None, None]:
        """
        Returns (min_y, min_x, max_y, max_x) of all selection boxes
        """
        return self.MT.get_selected_min_max()

    # Modifying Selected Cells

    def set_currently_selected(self, row: int | None = None, column: int | None = None, **kwargs) -> Sheet:
        self.MT.set_currently_selected(
            r=row,
            c=column,
            **kwargs,
        )
        return self

    def select_row(self, row: int, redraw: bool = True, run_binding_func: bool = True) -> Sheet:
        self.RI.select_row(
            row if isinstance(row, int) else int(row),
            redraw=False,
            run_binding_func=run_binding_func,
            ext=True,
        )
        return self.set_refresh_timer(redraw)

    def select_column(self, column: int, redraw: bool = True, run_binding_func: bool = True) -> Sheet:
        self.CH.select_col(
            column if isinstance(column, int) else int(column),
            redraw=False,
            run_binding_func=run_binding_func,
            ext=True,
        )
        return self.set_refresh_timer(redraw)

    def select_cell(self, row: int, column: int, redraw: bool = True, run_binding_func: bool = True) -> Sheet:
        self.MT.select_cell(
            row if isinstance(row, int) else int(row),
            column if isinstance(column, int) else int(column),
            redraw=False,
            run_binding_func=run_binding_func,
            ext=True,
        )
        return self.set_refresh_timer(redraw)

    def select_all(self, redraw: bool = True, run_binding_func: bool = True) -> Sheet:
        self.MT.select_all(redraw=False, run_binding_func=run_binding_func)
        return self.set_refresh_timer(redraw)

    def add_cell_selection(
        self,
        row: int,
        column: int,
        redraw: bool = True,
        run_binding_func: bool = True,
        set_as_current: bool = True,
    ) -> Sheet:
        self.MT.add_selection(
            r=row,
            c=column,
            redraw=False,
            run_binding_func=run_binding_func,
            set_as_current=set_as_current,
            ext=True,
        )
        return self.set_refresh_timer(redraw)

    def add_row_selection(
        self,
        row: int,
        redraw: bool = True,
        run_binding_func: bool = True,
        set_as_current: bool = True,
    ) -> Sheet:
        self.RI.add_selection(
            r=row,
            redraw=False,
            run_binding_func=run_binding_func,
            set_as_current=set_as_current,
            ext=True,
        )
        return self.set_refresh_timer(redraw)

    def add_column_selection(
        self,
        column: int,
        redraw: bool = True,
        run_binding_func: bool = True,
        set_as_current: bool = True,
    ) -> Sheet:
        self.CH.add_selection(
            c=column,
            redraw=False,
            run_binding_func=run_binding_func,
            set_as_current=set_as_current,
            ext=True,
        )
        return self.set_refresh_timer(redraw)

    def toggle_select_cell(
        self,
        row: int,
        column: int,
        add_selection: bool = True,
        redraw: bool = True,
        run_binding_func: bool = True,
        set_as_current: bool = True,
    ) -> Sheet:
        self.MT.toggle_select_cell(
            row=row,
            column=column,
            add_selection=add_selection,
            redraw=False,
            run_binding_func=run_binding_func,
            set_as_current=set_as_current,
            ext=True,
        )
        return self.set_refresh_timer(redraw)

    def toggle_select_row(
        self,
        row: int,
        add_selection: bool = True,
        redraw: bool = True,
        run_binding_func: bool = True,
        set_as_current: bool = True,
    ) -> Sheet:
        self.RI.toggle_select_row(
            row=row,
            add_selection=add_selection,
            redraw=False,
            run_binding_func=run_binding_func,
            set_as_current=set_as_current,
            ext=True,
        )
        return self.set_refresh_timer(redraw)

    def toggle_select_column(
        self,
        column: int,
        add_selection: bool = True,
        redraw: bool = True,
        run_binding_func: bool = True,
        set_as_current: bool = True,
    ) -> Sheet:
        self.CH.toggle_select_col(
            column=column,
            add_selection=add_selection,
            redraw=False,
            run_binding_func=run_binding_func,
            set_as_current=set_as_current,
            ext=True,
        )
        return self.set_refresh_timer(redraw)

    def create_selection_box(
        self,
        r1: int,
        c1: int,
        r2: int,
        c2: int,
        type_: Literal["cells", "rows", "columns", "cols"] = "cells",
    ) -> int:
        return self.MT.create_selection_box(
            r1=r1,
            c1=c1,
            r2=r2,
            c2=c2,
            type_="columns" if type_ == "cols" else type_,
            ext=True,
        )

    def recreate_all_selection_boxes(self) -> Sheet:
        self.MT.recreate_all_selection_boxes()
        return self

    def deselect(
        self,
        row: int | None | Literal["all"] = None,
        column: int | None = None,
        cell: tuple[int, int] | None = None,
        redraw: bool = True,
    ) -> Sheet:
        self.MT.deselect(r=row, c=column, cell=cell, redraw=False)
        return self.set_refresh_timer(redraw)

    def deselect_any(
        self,
        rows: Iterator[int] | int | None,
        columns: Iterator[int] | int | None,
        redraw: bool = True,
    ) -> Sheet:
        self.MT.deselect_any(rows=rows, columns=columns, redraw=False)
        return self.set_refresh_timer(redraw)

    # Row Heights and Column Widths

    def default_column_width(self, width: int | None = None) -> int:
        if isinstance(width, int):
            self.ops.default_column_width = width
        return self.ops.default_column_width

    def default_row_height(self, height: int | str | None = None) -> int:
        if isinstance(height, (int, str)):
            self.ops.default_row_height = height
        return self.ops.default_row_height

    def default_header_height(self, height: int | str | None = None) -> int:
        if isinstance(height, (int, str)):
            self.ops.default_header_height = height
        return self.ops.default_header_height

    def set_cell_size_to_text(
        self,
        row: int,
        column: int,
        only_set_if_too_small: bool = False,
        redraw: bool = True,
    ) -> Sheet:
        self.MT.set_cell_size_to_text(r=row, c=column, only_if_too_small=only_set_if_too_small)
        return self.set_refresh_timer(redraw)

    def set_all_cell_sizes_to_text(
        self,
        redraw: bool = True,
        width: int | None = None,
        slim: bool = False,
    ) -> tuple[list[float], list[float]]:
        self.MT.set_all_cell_sizes_to_text(width=width, slim=slim)
        self.set_refresh_timer(redraw)
        return self.MT.row_positions, self.MT.col_positions

    def set_all_column_widths(
        self,
        width: int | None = None,
        only_set_if_too_small: bool = False,
        redraw: bool = True,
        recreate_selection_boxes: bool = True,
    ) -> Sheet:
        self.CH.set_width_of_all_cols(
            width=width,
            only_if_too_small=only_set_if_too_small,
            recreate=recreate_selection_boxes,
        )
        return self.set_refresh_timer(redraw)

    def set_all_row_heights(
        self,
        height: int | None = None,
        only_set_if_too_small: bool = False,
        redraw: bool = True,
        recreate_selection_boxes: bool = True,
    ) -> Sheet:
        self.RI.set_height_of_all_rows(
            height=height,
            only_if_too_small=only_set_if_too_small,
            recreate=recreate_selection_boxes,
        )
        return self.set_refresh_timer(redraw)

    def column_width(
        self,
        column: int | Literal["all", "displayed"] | None = None,
        width: int | Literal["default", "text"] | None = None,
        only_set_if_too_small: bool = False,
        redraw: bool = True,
    ) -> Sheet | int:
        if column == "all":
            if width == "default":
                self.MT.reset_col_positions()
            elif width == "text":
                self.set_all_column_widths(only_set_if_too_small=only_set_if_too_small)
            elif isinstance(width, int):
                self.MT.reset_col_positions(width=width)

        elif column == "displayed":
            if width == "default":
                for c in range(*self.MT.visible_text_columns):
                    self.CH.set_col_width(
                        c, width=self.ops.default_column_width, only_if_too_small=only_set_if_too_small
                    )
            elif width == "text":
                for c in range(*self.MT.visible_text_columns):
                    self.CH.set_col_width(c, only_if_too_small=only_set_if_too_small)
            elif isinstance(width, int):
                for c in range(*self.MT.visible_text_columns):
                    self.CH.set_col_width(c, width=width, only_if_too_small=only_set_if_too_small)

        elif isinstance(column, int):
            if width == "default":
                self.CH.set_col_width(
                    col=column, width=self.ops.default_column_width, only_if_too_small=only_set_if_too_small
                )
            elif width == "text":
                self.CH.set_col_width(col=column, only_if_too_small=only_set_if_too_small)
            elif isinstance(width, int):
                self.CH.set_col_width(col=column, width=width, only_if_too_small=only_set_if_too_small)
            elif width is None:
                return int(self.MT.col_positions[column + 1] - self.MT.col_positions[column])

        else:
            return self.set_refresh_timer(redraw)

    def row_height(
        self,
        row: int | Literal["all", "displayed"] | None = None,
        height: int | Literal["default", "text"] | None = None,
        only_set_if_too_small: bool = False,
        redraw: bool = True,
    ) -> Sheet | int:
        if row == "all":
            if height == "default":
                self.MT.reset_row_positions()
            elif height == "text":
                self.set_all_row_heights()
            elif isinstance(height, int):
                self.MT.reset_row_positions(height=height)

        elif row == "displayed":
            if height == "default":
                height = self.MT.get_default_row_height()
                for r in range(*self.MT.visible_text_rows):
                    self.RI.set_row_height(r, height=height, only_if_too_small=only_set_if_too_small)
            elif height == "text":
                for r in range(*self.MT.visible_text_rows):
                    self.RI.set_row_height(r, only_if_too_small=only_set_if_too_small)
            elif isinstance(height, int):
                for r in range(*self.MT.visible_text_rows):
                    self.RI.set_row_height(r, height=height, only_if_too_small=only_set_if_too_small)

        elif isinstance(row, int):
            if height == "default":
                height = self.MT.get_default_row_height()
                self.RI.set_row_height(row=row, height=height, only_if_too_small=only_set_if_too_small)
            elif height == "text":
                self.RI.set_row_height(row=row, only_if_too_small=only_set_if_too_small)
            elif isinstance(height, int):
                self.RI.set_row_height(row=row, height=height, only_if_too_small=only_set_if_too_small)
            elif height is None:
                return int(self.MT.row_positions[row + 1] - self.MT.row_positions[row])

        else:
            return self.set_refresh_timer(redraw)

    def get_column_widths(self, canvas_positions: bool = False) -> list[float]:
        if canvas_positions:
            return self.MT.col_positions
        return self.MT.get_column_widths()

    def get_row_heights(self, canvas_positions: bool = False) -> list[float]:
        if canvas_positions:
            return self.MT.row_positions
        return self.MT.get_row_heights()

    def get_safe_row_heights(self) -> list[int]:
        default_h = self.MT.get_default_row_height()
        return [0 if e == default_h else e for e in self.MT.gen_row_heights()]

    def set_safe_row_heights(self, heights: list[int]) -> Sheet:
        default_h = self.MT.get_default_row_height()
        self.MT.row_positions = list(
            accumulate(chain([0], (self.valid_row_height(e) if e else default_h for e in heights)))
        )
        return self

    def get_row_text_height(
        self,
        row: int,
        visible_only: bool = False,
        only_if_too_small: bool = False,
    ) -> int:
        return self.RI.get_row_text_height(
            row=row,
            visible_only=visible_only,
            only_if_too_small=only_if_too_small,
        )

    def get_column_text_width(
        self,
        column: int,
        visible_only: bool = False,
        only_if_too_small: bool = False,
    ) -> int:
        return self.CH.get_col_text_width(
            col=column,
            visible_only=visible_only,
            only_if_too_small=only_if_too_small,
        )

    def set_column_widths(
        self,
        column_widths: Iterator[float] | None = None,
        canvas_positions: bool = False,
        reset: bool = False,
    ) -> Sheet:
        if reset or column_widths is None:
            self.MT.reset_col_positions()
        elif is_iterable(column_widths):
            if canvas_positions and isinstance(column_widths, list):
                self.MT.col_positions = column_widths
            else:
                self.MT.col_positions = list(accumulate(chain([0], column_widths)))
        return self

    def set_row_heights(
        self,
        row_heights: Iterator[float] | None = None,
        canvas_positions: bool = False,
        reset: bool = False,
    ) -> Sheet:
        if reset or row_heights is None:
            self.MT.reset_row_positions()
        elif is_iterable(row_heights):
            if canvas_positions and isinstance(row_heights, list):
                self.MT.row_positions = row_heights
            else:
                self.MT.row_positions = list(accumulate(chain([0], row_heights)))
        return self

    def set_width_of_index_to_text(self, text: None | str = None, *args, **kwargs) -> Sheet:
        self.RI.set_width_of_index_to_text(text=text)
        return self

    def set_index_width(self, pixels: int, redraw: bool = True) -> Sheet:
        if self.ops.auto_resize_row_index:
            self.ops.auto_resize_row_index = False
        self.RI.set_width(pixels, set_TL=True)
        return self.set_refresh_timer(redraw)

    def set_height_of_header_to_text(self, text: None | str = None) -> Sheet:
        self.CH.set_height_of_header_to_text(text=text)
        return self

    def set_header_height_pixels(self, pixels: int, redraw: bool = True) -> Sheet:
        self.CH.set_height(pixels, set_TL=True)
        return self.set_refresh_timer(redraw)

    def set_header_height_lines(self, nlines: int, redraw: bool = True) -> Sheet:
        self.CH.set_height(
            self.MT.get_lines_cell_height(
                nlines,
                font=self.ops.header_font,
            ),
            set_TL=True,
        )
        return self.set_refresh_timer(redraw)

    def del_row_position(self, idx: int, deselect_all: bool = False) -> Sheet:
        self.MT.del_row_position(idx=idx, deselect_all=deselect_all)
        return self

    delete_row_position = del_row_position

    def del_row_positions(self, idxs: Iterator[int] | None = None) -> Sheet:
        self.MT.del_row_positions(idxs=idxs)
        self.set_refresh_timer()
        return self

    def del_column_position(self, idx: int, deselect_all: bool = False) -> Sheet:
        self.MT.del_col_position(idx, deselect_all=deselect_all)
        return self

    delete_column_position = del_column_position

    def del_column_positions(self, idxs: Iterator[int] | None = None) -> Sheet:
        self.MT.del_col_positions(idxs=idxs)
        self.set_refresh_timer()
        return self

    def insert_column_position(
        self,
        idx: Literal["end"] | int = "end",
        width: int | None = None,
        deselect_all: bool = False,
        redraw: bool = False,
    ) -> Sheet:
        self.MT.insert_col_positions(idx=idx, widths=width, deselect_all=deselect_all)
        return self.set_refresh_timer(redraw)

    def insert_row_position(
        self,
        idx: Literal["end"] | int = "end",
        height: int | None = None,
        deselect_all: bool = False,
        redraw: bool = False,
    ) -> Sheet:
        self.MT.insert_row_positions(idx=idx, heights=height, deselect_all=deselect_all)
        return self.set_refresh_timer(redraw)

    def insert_column_positions(
        self,
        idx: Literal["end"] | int = "end",
        widths: Sequence[float] | int | None = None,
        deselect_all: bool = False,
        redraw: bool = False,
    ) -> Sheet:
        self.MT.insert_col_positions(idx=idx, widths=widths, deselect_all=deselect_all)
        return self.set_refresh_timer(redraw)

    def insert_row_positions(
        self,
        idx: Literal["end"] | int = "end",
        heights: Sequence[float] | int | None = None,
        deselect_all: bool = False,
        redraw: bool = False,
    ) -> Sheet:
        self.MT.insert_row_positions(idx=idx, heights=heights, deselect_all=deselect_all)
        return self.set_refresh_timer(redraw)

    def sheet_display_dimensions(
        self,
        total_rows: int | None = None,
        total_columns: int | None = None,
    ) -> tuple[int, int] | Sheet:
        if total_rows is None and total_columns is None:
            return len(self.MT.row_positions) - 1, len(self.MT.col_positions) - 1
        if isinstance(total_rows, int):
            height = self.MT.get_default_row_height()
            self.MT.row_positions = list(accumulate(chain([0], repeat(height, total_rows))))
        if isinstance(total_columns, int):
            width = self.ops.default_column_width
            self.MT.col_positions = list(accumulate(chain([0], repeat(width, total_columns))))
        return self

    def move_row_position(self, row: int, moveto: int) -> Sheet:
        self.MT.move_row_position(row, moveto)
        return self

    def move_column_position(self, column: int, moveto: int) -> Sheet:
        self.MT.move_col_position(column, moveto)
        return self

    def get_example_canvas_column_widths(self, total_cols: int | None = None) -> list[float]:
        colpos = int(self.ops.default_column_width)
        if isinstance(total_cols, int):
            return list(accumulate(chain([0], repeat(colpos, total_cols))))
        return list(accumulate(chain([0], repeat(colpos, len(self.MT.col_positions) - 1))))

    def get_example_canvas_row_heights(self, total_rows: int | None = None) -> list[float]:
        rowpos = self.MT.get_default_row_height()
        if isinstance(total_rows, int):
            return list(accumulate(chain([0], repeat(rowpos, total_rows))))
        return list(accumulate(chain([0], repeat(rowpos, len(self.MT.row_positions) - 1))))

    def verify_row_heights(self, row_heights: list[float], canvas_positions: bool = False) -> bool:
        if not isinstance(row_heights, list):
            return False
        if canvas_positions:
            if row_heights[0] != 0:
                return False
            return not any(
                x - z < self.MT.min_row_height or not isinstance(x, int) or isinstance(x, bool)
                for z, x in zip(row_heights, islice(row_heights, 1, None))
            )
        return not any(z < self.MT.min_row_height or not isinstance(z, int) or isinstance(z, bool) for z in row_heights)

    def verify_column_widths(self, column_widths: list[float], canvas_positions: bool = False) -> bool:
        if not isinstance(column_widths, list):
            return False
        if canvas_positions:
            if column_widths[0] != 0:
                return False
            return not any(
                x - z < self.ops.min_column_width or not isinstance(x, int) or isinstance(x, bool)
                for z, x in zip(column_widths, islice(column_widths, 1, None))
            )
        return not any(
            z < self.ops.min_column_width or not isinstance(z, int) or isinstance(z, bool) for z in column_widths
        )

    def valid_row_height(self, height: int) -> int:
        if height < self.MT.min_row_height:
            return self.MT.min_row_height
        elif height > self.ops.max_row_height:
            return self.ops.max_row_height
        return height

    def valid_column_width(self, width: int) -> int:
        if width < self.ops.min_column_width:
            return self.ops.min_column_width
        elif width > self.ops.max_column_width:
            return self.ops.max_column_width
        return width

    @property
    def visible_rows(self) -> tuple[int, int]:
        """
        returns: tuple[visible start row int, visible end row int]
        """
        return self.MT.visible_text_rows

    @property
    def visible_columns(self) -> tuple[int, int]:
        """
        returns: tuple[visible start column int, visible end column int]
        """
        return self.MT.visible_text_columns

    # Identifying Bound Event Mouse Position

    def identify_region(self, event: Any) -> Literal["table", "index", "header", "top left"]:
        if event.widget == self.MT:
            return "table"
        elif event.widget == self.RI:
            return "index"
        elif event.widget == self.CH:
            return "header"
        elif event.widget == self.TL:
            return "top left"

    def identify_row(
        self,
        event: Any,
        exclude_index: bool = False,
        allow_end: bool = True,
    ) -> int | None:
        ev_w = event.widget
        if ev_w == self.MT:
            return self.MT.identify_row(y=event.y, allow_end=allow_end)
        elif ev_w == self.RI:
            if exclude_index:
                return None
            else:
                return self.MT.identify_row(y=event.y, allow_end=allow_end)
        elif ev_w == self.CH or ev_w == self.TL:
            return None

    def identify_column(
        self,
        event: Any,
        exclude_header: bool = False,
        allow_end: bool = True,
    ) -> int | None:
        ev_w = event.widget
        if ev_w == self.MT:
            return self.MT.identify_col(x=event.x, allow_end=allow_end)
        elif ev_w == self.RI or ev_w == self.TL:
            return None
        elif ev_w == self.CH:
            if exclude_header:
                return None
            else:
                return self.MT.identify_col(x=event.x, allow_end=allow_end)

    # Scroll Positions and Cell Visibility

    def sync_scroll(self, widget: Any) -> Sheet:
        if widget is self:
            return self
        self.MT.synced_scrolls.add(widget)
        if isinstance(widget, Sheet):
            widget.MT.synced_scrolls.add(self)
        return self

    def unsync_scroll(self, widget: Any = None) -> Sheet:
        if widget is None:
            for widget in self.MT.synced_scrolls:
                if isinstance(widget, Sheet):
                    widget.MT.synced_scrolls.discard(self)
            self.MT.synced_scrolls = set()
        else:
            if isinstance(widget, Sheet) and self in widget.MT.synced_scrolls:
                widget.MT.synced_scrolls.discard(self)
            self.MT.synced_scrolls.discard(widget)
        return self

    def see(
        self,
        row: int = 0,
        column: int = 0,
        keep_yscroll: bool = False,
        keep_xscroll: bool = False,
        bottom_right_corner: bool | None = None,
        check_cell_visibility: bool = True,
        redraw: bool = True,
    ) -> Sheet:
        self.MT.see(
            row,
            column,
            keep_yscroll,
            keep_xscroll,
            bottom_right_corner,
            check_cell_visibility=check_cell_visibility,
            redraw=False,
        )
        return self.set_refresh_timer(redraw)

    def cell_visible(self, r: int, c: int) -> bool:
        return self.MT.cell_visible(r, c)

    def cell_completely_visible(self, r: int, c: int, seperate_axes: bool = False) -> bool:
        if seperate_axes:
            d = self.MT.cell_visibility_info(r, c)
            return d["yvis"], d["xvis"]
        else:
            return self.MT.cell_completely_visible(r, c)

    def set_xview(self, position: None | float = None, option: str = "moveto") -> Sheet | tuple[float, float]:
        if position is not None:
            self.MT.set_xviews(option, position)
            return self
        return self.MT.xview()

    xview = set_xview
    xview_moveto = set_xview

    def set_yview(self, position: None | float = None, option: str = "moveto") -> Sheet | tuple[float, float]:
        if position is not None:
            self.MT.set_yviews(option, position)
            return self
        return self.MT.yview()

    yview = set_yview
    yview_moveto = set_yview

    def set_view(self, x_args: list[str, float], y_args: list[str, float]) -> Sheet:
        self.MT.set_view(x_args, y_args)
        return self

    def get_xview(self) -> tuple[float, float]:
        return self.MT.xview()

    def get_yview(self) -> tuple[float, float]:
        return self.MT.yview()

    # Hiding Columns

    def displayed_column_to_data(self, c: int) -> int:
        return c if self.MT.all_columns_displayed else self.MT.displayed_columns[c]

    data_c = displayed_column_to_data
    datacn = displayed_column_to_data
    dcol = displayed_column_to_data

    def data_column_to_displayed(self, c: int) -> int:
        return self.MT.dispcn(c)

    dispcn = data_column_to_displayed

    def display_columns(
        self,
        columns: None | Literal["all"] | Iterator[int] = None,
        all_columns_displayed: None | bool = None,
        reset_col_positions: bool = True,
        redraw: bool = False,
        deselect_all: bool = True,
        **kwargs,
    ) -> list[int] | None:
        if "all_displayed" in kwargs:
            all_columns_displayed = kwargs["all_displayed"]
        res = self.MT.display_columns(
            columns=None if isinstance(columns, str) and columns.lower() == "all" else columns,
            all_columns_displayed=(
                True if isinstance(columns, str) and columns.lower() == "all" else all_columns_displayed
            ),
            reset_col_positions=reset_col_positions,
            deselect_all=deselect_all,
        )
        if "refresh" in kwargs:
            redraw = kwargs["refresh"]
        self.set_refresh_timer(redraw)
        return res

    def hide_columns(
        self,
        columns: int | set[int] | Iterator[int],
        redraw: bool = True,
        deselect_all: bool = True,
        data_indexes: bool = False,
    ) -> Sheet:
        if isinstance(columns, int):
            columns = {columns}
        elif not isinstance(columns, set):
            columns = set(columns)
            if not columns:
                return
        if self.MT.all_columns_displayed:
            self.MT.displayed_columns = list(filterfalse(columns.__contains__, range(self.MT.total_data_cols())))
            to_pop = {c: c for c in columns}
        else:
            to_pop = {}
            new_disp = []
            if data_indexes:
                for i, c in enumerate(self.MT.displayed_columns):
                    if c not in columns:
                        new_disp.append(c)
                    else:
                        to_pop[i] = c
            else:
                for i, c in enumerate(self.MT.displayed_columns):
                    if i not in columns:
                        new_disp.append(c)
                    else:
                        to_pop[i] = c
            self.MT.displayed_columns = new_disp
        self.MT.all_columns_displayed = False
        self.MT.set_col_positions(
            pop_positions(
                itr=self.MT.gen_column_widths,
                to_pop=to_pop,
                save_to=self.MT.saved_column_widths,
            ),
        )
        if deselect_all:
            self.MT.deselect(redraw=False)
        return self.set_refresh_timer(redraw)

    def show_columns(
        self,
        columns: int | Iterator[int],
        redraw: bool = True,
        deselect_all: bool = True,
    ) -> Sheet:
        """
        'columns' argument must be data indexes
        Function will return without action if Sheet.all_columns
        """
        if self.MT.all_columns_displayed:
            return
        if isinstance(columns, int):
            columns = [columns]
        cws = self.MT.get_column_widths()
        default_col_w = self.ops.default_column_width
        for column in columns:
            idx = bisect_left(self.MT.displayed_columns, column)
            if len(self.MT.displayed_columns) == idx or self.MT.displayed_columns[idx] != column:
                self.MT.displayed_columns.insert(idx, column)
                cws.insert(idx, self.MT.saved_column_widths.pop(column, default_col_w))
        self.MT.set_col_positions(cws)
        if deselect_all:
            self.MT.deselect(redraw=False)
        return self.set_refresh_timer(redraw)

    def all_columns_displayed(self, a: bool | None = None) -> bool:
        v = bool(self.MT.all_columns_displayed)
        if isinstance(a, bool):
            self.MT.all_columns_displayed = a
        return v

    @property
    def all_columns(self) -> bool:
        return self.MT.all_columns_displayed

    @all_columns.setter
    def all_columns(self, a: bool) -> Sheet:
        self.MT.display_columns(
            columns=None,
            all_columns_displayed=a,
        )
        return self

    @property
    def displayed_columns(self) -> list[int]:
        return self.MT.displayed_columns

    @displayed_columns.setter
    def displayed_columns(self, columns: list[int]) -> Sheet:
        self.display_columns(columns=columns, reset_col_positions=True, redraw=False)
        return self.set_refresh_timer()

    # Hiding Rows

    def displayed_row_to_data(self, r: int) -> int:
        return r if self.MT.all_rows_displayed else self.MT.displayed_rows[r]

    data_r = displayed_row_to_data
    datarn = displayed_row_to_data
    drow = displayed_row_to_data

    def data_row_to_displayed(self, r: int) -> int:
        return self.MT.disprn(r)

    disprn = data_row_to_displayed

    def display_rows(
        self,
        rows: None | Literal["all"] | Iterator[int] = None,
        all_rows_displayed: None | bool = None,
        reset_row_positions: bool = True,
        redraw: bool = False,
        deselect_all: bool = True,
        **kwargs,
    ) -> list[int] | None:
        if "all_displayed" in kwargs:
            all_rows_displayed = kwargs["all_displayed"]
        res = self.MT.display_rows(
            rows=None if isinstance(rows, str) and rows.lower() == "all" else rows,
            all_rows_displayed=True if isinstance(rows, str) and rows.lower() == "all" else all_rows_displayed,
            reset_row_positions=reset_row_positions,
            deselect_all=deselect_all,
        )
        if "refresh" in kwargs:
            redraw = kwargs["refresh"]
        self.set_refresh_timer(redraw)
        return res

    def hide_rows(
        self,
        rows: int | set[int] | Iterator[int],
        redraw: bool = True,
        deselect_all: bool = True,
        data_indexes: bool = False,
        row_heights: bool = True,
    ) -> Sheet:
        if isinstance(rows, int):
            rows = {rows}
        elif not isinstance(rows, set):
            rows = set(rows)
            if not rows:
                return
        if self.MT.all_rows_displayed:
            self.MT.displayed_rows = list(filterfalse(rows.__contains__, range(self.MT.total_data_rows())))
            to_pop = {r: r for r in rows}
        else:
            to_pop = {}
            new_disp = []
            if data_indexes:
                for i, r in enumerate(self.MT.displayed_rows):
                    if r not in rows:
                        new_disp.append(r)
                    else:
                        to_pop[i] = r
            else:
                for i, r in enumerate(self.MT.displayed_rows):
                    if i not in rows:
                        new_disp.append(r)
                    else:
                        to_pop[i] = r
            self.MT.displayed_rows = new_disp
        self.MT.all_rows_displayed = False
        if row_heights:
            self.MT.set_row_positions(
                pop_positions(
                    itr=self.MT.gen_row_heights,
                    to_pop=to_pop,
                    save_to=self.MT.saved_row_heights,
                ),
            )
        if deselect_all:
            self.MT.deselect(redraw=False)
        return self.set_refresh_timer(redraw)

    def show_rows(
        self,
        rows: int | Iterator[int],
        redraw: bool = True,
        deselect_all: bool = True,
    ) -> Sheet:
        """
        'rows' argument must be data indexes
        Function will return without action if Sheet.all_rows
        """
        if self.MT.all_rows_displayed:
            return
        if isinstance(rows, int):
            rows = [rows]
        default_row_h = self.MT.get_default_row_height()
        rhs = self.MT.get_row_heights()
        for row in rows:
            idx = bisect_left(self.MT.displayed_rows, row)
            if len(self.MT.displayed_rows) == idx or self.MT.displayed_rows[idx] != row:
                self.MT.displayed_rows.insert(idx, row)
                rhs.insert(idx, self.MT.saved_row_heights.pop(row, default_row_h))
        self.MT.set_row_positions(rhs)
        if deselect_all:
            self.MT.deselect(redraw=False)
        return self.set_refresh_timer(redraw)

    def all_rows_displayed(self, a: bool | None = None) -> bool:
        v = bool(self.MT.all_rows_displayed)
        if isinstance(a, bool):
            self.MT.all_rows_displayed = a
        return v

    @property
    def all_rows(self) -> bool:
        return self.MT.all_rows_displayed

    @all_rows.setter
    def all_rows(self, a: bool) -> Sheet:
        self.MT.display_rows(
            rows=None,
            all_rows_displayed=a,
        )
        return self

    @property
    def displayed_rows(self) -> list[int]:
        return self.MT.displayed_rows

    @displayed_rows.setter
    def displayed_rows(self, rows: list[int]) -> Sheet:
        self.display_rows(rows=rows, reset_row_positions=True, redraw=False)
        return self.set_refresh_timer()

    # Hiding Sheet Elements

    def show(
        self,
        canvas: Literal[
            "all",
            "row_index",
            "index",
            "header",
            "top_left",
            "x_scrollbar",
            "y_scrollbar",
            "table",
        ] = "all",
    ) -> Sheet:
        if canvas in ("all", "table"):
            self.MT.grid(row=1, column=1, sticky="nswe")
        if canvas in ("all", "row_index", "index"):
            self.RI.grid(row=1, column=0, sticky="nswe")
            self.MT.show_index = True
        if canvas in ("all", "header"):
            self.CH.grid(row=0, column=1, sticky="nswe")
            self.MT.show_header = True
        if canvas in ("all", "top_left") or (
            self.ops.show_top_left is not False and self.MT.show_header and self.MT.show_index
        ):
            self.TL.grid(row=0, column=0)
        if canvas in ("all", "x_scrollbar"):
            self.xscroll.grid(row=2, column=0, columnspan=2, sticky="nswe")
            self.xscroll_showing = True
            self.xscroll_disabled = False
        if canvas in ("all", "y_scrollbar"):
            self.yscroll.grid(row=0, column=2, rowspan=3, sticky="nswe")
            self.yscroll_showing = True
            self.yscroll_disabled = False
        if self._startup_complete:
            self.MT.update_idletasks()
        return self

    def hide(
        self,
        canvas: Literal[
            "all",
            "row_index",
            "header",
            "top_left",
            "x_scrollbar",
            "y_scrollbar",
            "table",
        ] = "all",
    ) -> Sheet:
        if canvas in ("all", "row_index"):
            self.RI.grid_remove()
            self.MT.show_index = False
        if canvas in ("all", "header"):
            self.CH.grid_remove()
            self.MT.show_header = False
        if canvas in ("all", "top_left") or (
            not self.ops.show_top_left and (not self.MT.show_index or not self.MT.show_header)
        ):
            self.TL.grid_remove()
        if canvas in ("all", "x_scrollbar"):
            self.xscroll.grid_remove()
            self.xscroll_showing = False
            self.xscroll_disabled = True
        if canvas in ("all", "y_scrollbar"):
            self.yscroll.grid_remove()
            self.yscroll_showing = False
            self.yscroll_disabled = True
        if canvas in ("all", "table"):
            self.MT.grid_remove()
        return self

    # Sheet Height and Width

    def height_and_width(
        self,
        height: int | None = None,
        width: int | None = None,
    ) -> Sheet:
        if width is not None or height is not None:
            self.grid_propagate(0)
        elif width is None and height is None:
            self.grid_propagate(1)
        if width is not None:
            self.config(width=width)
        if height is not None:
            self.config(height=height)
        return self

    def get_frame_y(self, y: int) -> int:
        return y + self.CH.current_height

    def get_frame_x(self, x: int) -> int:
        return x + self.RI.current_width

    # Cell Text Editor

    def next_cell(self, r: int, c: int, key: Literal["Return", "Tab", "??"]) -> Sheet:
        self.MT.go_to_next_cell(r, c, key)
        return self

    # works on currently selected box
    def open_cell(self, ignore_existing_editor: bool = True) -> Sheet:
        self.MT.open_cell(event=GeneratedMouseEvent(), ignore_existing_editor=ignore_existing_editor)
        return self

    def open_header_cell(self, ignore_existing_editor: bool = True) -> Sheet:
        self.CH.open_cell(event=GeneratedMouseEvent(), ignore_existing_editor=ignore_existing_editor)
        return self

    def open_index_cell(self, ignore_existing_editor: bool = True) -> Sheet:
        self.RI.open_cell(event=GeneratedMouseEvent(), ignore_existing_editor=ignore_existing_editor)
        return self

    def set_text_editor_value(
        self,
        text: str = "",
    ) -> Sheet:
        if self.MT.text_editor.open:
            self.MT.text_editor.window.set_text(text)
        return self

    def set_index_text_editor_value(
        self,
        text: str = "",
    ) -> Sheet:
        if self.RI.text_editor.open:
            self.RI.text_editor.window.set_text(text)
        return self

    def set_header_text_editor_value(
        self,
        text: str = "",
    ) -> Sheet:
        if self.CH.text_editor.open:
            self.CH.text_editor.window.set_text(text)
        return self

    def destroy_text_editor(self, event: Any = None) -> Sheet:
        self.MT.hide_text_editor(reason=event)
        self.RI.hide_text_editor(reason=event)
        self.CH.hide_text_editor(reason=event)
        return self

    def get_text_editor_widget(self, event: Any = None) -> tk.Text | None:
        try:
            return self.MT.text_editor.tktext
        except Exception:
            return None

    def bind_key_text_editor(self, key: str, function: Callable) -> Sheet:
        self.MT.text_editor_user_bound_keys[key] = function
        return self

    def unbind_key_text_editor(self, key: str) -> Sheet:
        if key == "all":
            for key in self.MT.text_editor_user_bound_keys:
                with suppress(Exception):
                    self.MT.text_editor.tktext.unbind(key)
            self.MT.text_editor_user_bound_keys = {}
        else:
            if key in self.MT.text_editor_user_bound_keys:
                del self.MT.text_editor_user_bound_keys[key]
            with suppress(Exception):
                self.MT.text_editor.tktext.unbind(key)
        return self

    def get_text_editor_value(self) -> str | None:
        return self.MT.text_editor.get()

    def close_text_editor(self, set_data: bool = True) -> Sheet:
        event = new_tk_event("ButtonPress-1" if set_data else "Escape")
        if self.MT.text_editor.open:
            self.MT.close_text_editor(event=event)
        if self.RI.text_editor.open:
            self.RI.close_text_editor(event=event)
        if self.CH.text_editor.open:
            self.CH.close_text_editor(event=event)
        return self

    # Sheet Options and Other Functions

    def set_options(self, redraw: bool = True, **kwargs) -> Sheet:
        enabled = tuple(self.MT.enabled_bindings)
        for k, v in kwargs.items():
            if k in self.ops and v != self.ops[k]:
                if k.endswith("bindings"):
                    for b in enabled:
                        self.MT._disable_binding(b)
                self.ops[k] = v
                if k.endswith("bindings"):
                    for b in enabled:
                        self.MT._enable_binding(b)
        if "name" in kwargs:
            self.name = kwargs["name"]
        if "min_column_width" in kwargs:
            self.MT.set_min_column_width(kwargs["min_column_width"])
        if "from_clipboard_delimiters" in kwargs:
            self.ops.from_clipboard_delimiters = (
                self.ops.from_clipboard_delimiters
                if isinstance(self.ops.from_clipboard_delimiters, str)
                else "".join(self.ops.from_clipboard_delimiters)
            )
        if "default_row_height" in kwargs:
            self.default_row_height(kwargs["default_row_height"])
        if "expand_sheet_if_paste_too_big" in kwargs:
            self.ops.paste_can_expand_x = kwargs["expand_sheet_if_paste_too_big"]
            self.ops.paste_can_expand_y = kwargs["expand_sheet_if_paste_too_big"]
        if "show_top_left" in kwargs:
            self.show("top_left")
        if "popup_menu_font" in kwargs:
            self.popup_menu_font(self.ops.popup_menu_font)
        if "font" in kwargs:
            self.MT.set_table_font(kwargs["font"])
        elif "table_font" in kwargs:
            self.MT.set_table_font(kwargs["table_font"])
        if "header_font" in kwargs:
            self.MT.set_header_font(kwargs["header_font"])
        if "index_font" in kwargs:
            self.MT.set_index_font(kwargs["index_font"])
        if "theme" in kwargs:
            self.change_theme(kwargs["theme"])
        if "header_bg" in kwargs:
            self.CH.config(background=kwargs["header_bg"])
        if "index_bg" in kwargs:
            self.RI.config(background=kwargs["index_bg"])
        if "top_left_bg" in kwargs:
            self.TL.config(background=kwargs["top_left_bg"])
            self.TL.itemconfig(self.TL.select_all_box, fill=kwargs["top_left_bg"])
        if "top_left_fg" in kwargs:
            self.TL.itemconfig("rw", fill=kwargs["top_left_fg"])
            self.TL.itemconfig("rh", fill=kwargs["top_left_fg"])
        if "frame_bg" in kwargs:
            self.config(background=kwargs["frame_bg"])
        if "table_bg" in kwargs:
            self.MT.config(background=kwargs["table_bg"])
        if "outline_thickness" in kwargs:
            self.config(highlightthickness=kwargs["outline_thickness"])
        if "outline_color" in kwargs:
            self.config(
                highlightbackground=kwargs["outline_color"],
                highlightcolor=kwargs["outline_color"],
            )
        if any(k in kwargs for k in scrollbar_options_keys):
            self.set_scrollbar_options()
        if "treeview" in kwargs:
            self.index_align("nw", redraw=False)
            self.ops.paste_can_expand_y = False
        return self.set_refresh_timer(redraw)

    def set_scrollbar_options(self) -> Sheet:
        style = ttk.Style()
        for orientation in ("vertical", "horizontal"):
            style.configure(
                f"Sheet{self.unique_id}.{orientation.capitalize()}.TScrollbar",
                background=self.ops[f"{orientation}_scroll_background"],
                troughcolor=self.ops[f"{orientation}_scroll_troughcolor"],
                lightcolor=self.ops[f"{orientation}_scroll_lightcolor"],
                darkcolor=self.ops[f"{orientation}_scroll_darkcolor"],
                relief=self.ops[f"{orientation}_scroll_relief"],
                troughrelief=self.ops[f"{orientation}_scroll_troughrelief"],
                bordercolor=self.ops[f"{orientation}_scroll_bordercolor"],
                borderwidth=self.ops[f"{orientation}_scroll_borderwidth"],
                gripcount=self.ops[f"{orientation}_scroll_gripcount"],
                arrowsize=self.ops[f"{orientation}_scroll_arrowsize"],
            )
            style.map(
                f"Sheet{self.unique_id}.{orientation.capitalize()}.TScrollbar",
                foreground=[
                    ("!active", self.ops[f"{orientation}_scroll_not_active_fg"]),
                    ("pressed", self.ops[f"{orientation}_scroll_pressed_fg"]),
                    ("active", self.ops[f"{orientation}_scroll_active_fg"]),
                ],
                background=[
                    ("!active", self.ops[f"{orientation}_scroll_not_active_bg"]),
                    ("pressed", self.ops[f"{orientation}_scroll_pressed_bg"]),
                    ("active", self.ops[f"{orientation}_scroll_active_bg"]),
                ],
            )
        return self

    def event_widget_is_sheet(
        self,
        event: Any,
        table: bool = True,
        index: bool = True,
        header: bool = True,
        top_left: bool = True,
    ) -> bool:
        return (
            table
            and event.widget == self.MT
            or index
            and event.widget == self.RI
            or header
            and event.widget == self.CH
            or top_left
            and event.widget == self.TL
        )

    def props(
        self,
        row: int,
        column: int | str,
        key: None | CellPropertyKey = None,
        cellops: bool = True,
        rowops: bool = True,
        columnops: bool = True,
    ) -> dict:
        """
        Retrieve options (properties - props)
        from a single cell in the main table

        Also retrieves any row or column options
        impacting that cell
        """
        return self.MT.get_cell_kwargs(
            datarn=row,
            datacn=alpha2idx(column) if isinstance(column, str) else column,
            key=key,
            cell=cellops,
            row=rowops,
            column=columnops,
        )

    def index_props(
        self,
        row: int,
        key: None | CellPropertyKey = None,
    ) -> dict:
        """
        Retrieve options (properties - props)
        from a cell in the index
        """
        return self.RI.get_cell_kwargs(
            datarn=row,
            key=key,
        )

    def header_props(
        self,
        column: int | str,
        key: None | CellPropertyKey = None,
    ) -> dict:
        """
        Retrieve options (properties - props)
        from a cell in the header
        """
        return self.CH.get_cell_kwargs(
            datacn=alpha2idx(column) if isinstance(column, str) else column,
            key=key,
        )

    def get_cell_options(
        self,
        key: None | CellPropertyKey = None,
        canvas: Literal["table", "row_index", "index", "header"] = "table",
    ) -> dict:
        if canvas == "table":
            target = self.MT.cell_options
        elif canvas in ("row_index", "index"):
            target = self.RI.cell_options
        elif canvas == "header":
            target = self.CH.cell_options
        if key is None:
            return target
        return {k: v[key] for k, v in target.items() if key in v}

    def get_row_options(self, key: None | CellPropertyKey = None) -> dict:
        if key is None:
            return self.MT.row_options
        return {k: v[key] for k, v in self.MT.row_options.items() if key in v}

    def get_column_options(self, key: None | CellPropertyKey = None) -> dict:
        if key is None:
            return self.MT.col_options
        return {k: v[key] for k, v in self.MT.col_options.items() if key in v}

    def get_index_options(self, key: None | CellPropertyKey = None) -> dict:
        if key is None:
            return self.RI.cell_options
        return {k: v[key] for k, v in self.RI.cell_options.items() if key in v}

    def get_header_options(self, key: None | CellPropertyKey = None) -> dict:
        if key is None:
            return self.CH.cell_options
        return {k: v[key] for k, v in self.CH.cell_options.items() if key in v}

    def del_out_of_bounds_options(self) -> Sheet:
        maxc = self.total_columns()
        maxr = self.total_rows()
        for name in tuple(self.MT.named_spans):
            span = self.MT.named_spans[name]
            if (
                (isinstance(span.upto_r, int) and span.upto_r > maxr)
                or (isinstance(span.upto_c, int) and span.upto_c > maxc)
                or span.from_r >= maxr
                or span.from_c >= maxc
            ):
                self.del_named_span(name)
        self.MT.cell_options = {k: v for k, v in self.MT.cell_options.items() if k[0] < maxr and k[1] < maxc}
        self.RI.cell_options = {k: v for k, v in self.RI.cell_options.items() if k < maxr}
        self.CH.cell_options = {k: v for k, v in self.CH.cell_options.items() if k < maxc}
        self.MT.col_options = {k: v for k, v in self.MT.col_options.items() if k < maxc}
        self.MT.row_options = {k: v for k, v in self.MT.row_options.items() if k < maxr}
        return self

    delete_out_of_bounds_options = del_out_of_bounds_options

    def reset_all_options(self) -> Sheet:
        self.MT.named_spans = {}
        self.MT.cell_options = {}
        self.RI.cell_options = {}
        self.CH.cell_options = {}
        self.MT.col_options = {}
        self.MT.row_options = {}
        return self

    def show_ctrl_outline(
        self,
        canvas: Literal["table"] = "table",
        start_cell: tuple[int, int] = (0, 0),
        end_cell: tuple[int, int] = (1, 1),
    ) -> Sheet:
        self.MT.show_ctrl_outline(canvas=canvas, start_cell=start_cell, end_cell=end_cell)
        return self

    def reset_undos(self) -> Sheet:
        self.MT.purge_undo_and_redo_stack()
        return self

    def get_undo_stack(self) -> deque:
        return self.MT.undo_stack

    def get_redo_stack(self) -> deque:
        return self.MT.redo_stack

    def set_undo_stack(self, stack: deque) -> Sheet:
        self.MT.undo_stack = stack
        return self

    def set_redo_stack(self, stack: deque) -> Sheet:
        self.MT.redo_stack = stack
        return self

    def redraw(self, redraw_header: bool = True, redraw_row_index: bool = True) -> Sheet:
        self.MT.main_table_redraw_grid_and_text(redraw_header=redraw_header, redraw_row_index=redraw_row_index)
        return self

    refresh = redraw

    # Progress Bars

    def create_progress_bar(
        self,
        row: int,
        column: int,
        bg: str,
        fg: str,
        name: Hashable,
        percent: int = 0,
        del_when_done: bool = False,
    ) -> Sheet:
        self.MT.progress_bars[(row, column)] = ProgressBar(
            bg=bg,
            fg=fg,
            name=name,
            percent=percent,
            del_when_done=del_when_done,
        )
        return self.set_refresh_timer()

    def progress_bar(
        self,
        name: Hashable | None = None,
        cell: tuple[int, int] | None = None,
        percent: int | None = None,
        bg: str | None = None,
        fg: str | None = None,
    ) -> Sheet:
        if name is not None:
            bars = (bar for bar in self.MT.progress_bars.values() if bar.name == name)
        elif cell is not None:
            bars = (self.MT.progress_bars[cell],)
        for bar in bars:
            if isinstance(percent, int):
                bar.percent = percent
            if isinstance(bg, str):
                bar.bg = bg
            if isinstance(fg, str):
                bar.fg = fg
        return self.set_refresh_timer()

    def del_progress_bar(
        self,
        name: Hashable | None = None,
        cell: tuple[int, int] | None = None,
    ) -> Sheet:
        if name is not None:
            for cell in tuple(cell for cell, bar in self.MT.progress_bars.items() if bar.name == name):
                del self.MT.progress_bars[cell]
        elif cell is not None:
            del self.MT.progress_bars[cell]
        return self.set_refresh_timer()

    # Tags

    def tag(
        self,
        *key: CreateSpanTypes,
        tags: Iterator[str] | str = "",
    ) -> Sheet:
        span = self.span_from_key(*key)
        rows, cols = self.ranges_from_span(span)
        if isinstance(tags, str):
            tags = (tags,)
        if span.kind == "cell":
            for tag in tags:
                if tag not in self.MT.tagged_cells:
                    self.MT.tagged_cells[tag] = set()
                for r in rows:
                    for c in cols:
                        self.MT.tagged_cells[tag].add((r, c))
        elif span.kind == "row":
            for tag in tags:
                if tag not in self.MT.tagged_rows:
                    self.MT.tagged_rows[tag] = set()
                self.MT.tagged_rows[tag].update(set(rows))
        elif span.kind == "column":
            for tag in tags:
                if tag not in self.MT.tagged_columns:
                    self.MT.tagged_columns[tag] = set()
                self.MT.tagged_columns[tag].update(set(cols))
        return self

    def tag_cell(
        self,
        cell: tuple[int, int],
        *tags,
    ) -> Sheet:
        if (
            not isinstance(cell, tuple)
            or not len(cell) == 2
            or not isinstance(cell[0], int)
            or not isinstance(cell[1], int)
        ):
            raise ValueError("'cell' argument must be tuple[int, int].")
        for tag in unpack(tags):
            if tag not in self.MT.tagged_cells:
                self.MT.tagged_cells[tag] = set()
            self.MT.tagged_cells[tag].add(cell)
        return self

    def tag_rows(
        self,
        rows: int | Iterator[int],
        *tags,
    ) -> Sheet:
        if isinstance(rows, int):
            rows = [rows]
        for tag in unpack(tags):
            if tag not in self.MT.tagged_rows:
                self.MT.tagged_rows[tag] = set()
            self.MT.tagged_rows[tag].update(rows)
        return self

    def tag_columns(
        self,
        columns: int | Iterator[int],
        *tags,
    ) -> Sheet:
        if isinstance(columns, int):
            columns = [columns]
        for tag in unpack(tags):
            if tag not in self.MT.tagged_columns:
                self.MT.tagged_columns[tag] = set()
            self.MT.tagged_columns[tag].update(columns)
        return self

    def untag(
        self,
        cell: tuple[int, int] | None = None,
        rows: int | Iterator[int] | None = None,
        columns: int | Iterator[int] | None = None,
    ) -> Sheet:
        if isinstance(cell, tuple):
            for tagged in self.MT.tagged_cells.values():
                tagged.discard(cell)
        if isinstance(rows, int):
            rows = (rows,)
        if is_iterable(rows):
            for tagged in self.MT.tagged_rows.values():
                for row in rows:
                    tagged.discard(row)
        if isinstance(columns, int):
            columns = (columns,)
        if is_iterable(columns):
            for tagged in self.MT.tagged_columns.values():
                for column in columns:
                    tagged.discard(column)
        return self

    def tag_del(
        self,
        *tags,
        cells: bool = True,
        rows: bool = True,
        columns: bool = True,
    ) -> Sheet:
        for tag in unpack(tags):
            if cells and tag in self.MT.tagged_cells:
                del self.MT.tagged_cells[tag]
            if rows and tag in self.MT.tagged_rows:
                del self.MT.tagged_rows[tag]
            if columns and tag in self.MT.tagged_columns:
                del self.MT.tagged_columns[tag]
        return self

    def tag_has(
        self,
        *tags,
    ) -> DotDict:
        res = DotDict(
            cells=set(),
            rows=set(),
            columns=set(),
        )
        for tag in unpack(tags):
            res.cells.update(self.MT.tagged_cells.get(tag, set()))
            res.rows.update(self.MT.tagged_rows.get(tag, set()))
            res.columns.update(self.MT.tagged_columns.get(tag, set()))
        return res

    # Treeview Mode

    def tree_build(
        self,
        data: list[list[Any]],
        iid_column: int,
        parent_column: int,
        text_column: None | int | list[str] = None,
        push_ops: bool = False,
        row_heights: Sequence[int] | None | False = None,
        open_ids: Iterator[str] | None = None,
        safety: bool = True,
        ncols: int | None = None,
        lower: bool = False,
        include_iid_column: bool = True,
        include_parent_column: bool = True,
        include_text_column: bool = True,
    ) -> Sheet:
        build = self.RI.safe_tree_build if safety else self.RI.tree_build
        self.reset(cell_options=False, column_widths=False, header=False, redraw=False)
        build(
            data=data,
            iid_column=iid_column,
            parent_column=parent_column,
            text_column=text_column,
            push_ops=push_ops,
            row_heights=row_heights,
            open_ids=open_ids,
            ncols=ncols,
            lower=lower,
            include_iid_column=include_iid_column,
            include_parent_column=include_parent_column,
            include_text_column=include_text_column,
        )
        return self

    def tree_reset(self) -> Sheet:
        self.deselect()
        self.RI.tree_reset()
        return self

    def tree_get_open(self) -> set[str]:
        """
        Returns the set[str] of iids that are open in the treeview
        """
        return self.RI.tree_open_ids

    def tree_set_open(self, open_ids: Iterator[str]) -> Sheet:
        """
        Accepts set[str] of iids that are open in the treeview
        Closes everything else
        """
        self.hide_rows(
            rows={rn for rn in self.MT.displayed_rows if self.MT._row_index[rn].parent},
            redraw=False,
            deselect_all=False,
            data_indexes=True,
        )
        self.RI.tree_open_ids = set()
        open_ids = filter(self.exists, open_ids)
        try:
            first_id = next(open_ids)
        except StopIteration:
            return self.set_refresh_timer()
        self.show_rows(
            rows=self._tree_open(chain((first_id,), open_ids)),
            redraw=False,
            deselect_all=False,
        )
        return self.set_refresh_timer()

    def _tree_open(self, items: Iterator[str]) -> Generator[int]:
        """
        Only meant for internal use
        """
        disp_set = set(self.MT.displayed_rows)
        index = self.MT._row_index
        rns = self.RI.rns
        open_ids = self.RI.tree_open_ids
        descendants = self.RI.get_iid_descendants
        for item in items:
            if item in rns and index[rns[item]].children:
                open_ids.add(item)
                if rns[item] in disp_set:
                    for did in descendants(item, check_open=True):
                        disp_set.add(rns[did])
                        yield rns[did]

    def tree_open(self, *items: str, redraw: bool = True) -> Sheet:
        """
        If used without args all items are opened
        """
        to_show = self._tree_open(items) if (items := set(unpack(items))) else self._tree_open(self.get_children())
        return self.show_rows(
            rows=to_show,
            redraw=redraw,
            deselect_all=False,
        )

    def _tree_close(self, items: Iterator[str]) -> set[int]:
        """
        Only meant for internal use
        """
        to_hide = set()
        disp_set = set(self.MT.displayed_rows)
        index = self.MT._row_index
        rns = self.RI.rns
        open_ids = self.RI.tree_open_ids
        descendants = self.RI.get_iid_descendants
        for item in items:
            if index[rns[item]].children:
                open_ids.discard(item)
                if rns[item] in disp_set:
                    for did in descendants(item, check_open=True):
                        to_hide.add(rns[did])
        return to_hide

    def tree_close(self, *items: str, redraw: bool = True) -> Sheet:
        """
        If used without args all items are closed
        """
        to_hide = self._tree_close(unpack(items)) if items else self._tree_close(self.get_children())
        return self.hide_rows(
            rows=to_hide,
            redraw=redraw,
            deselect_all=False,
            data_indexes=True,
        )

    def insert(
        self,
        parent: str = "",
        index: None | int | Literal["end"] = None,
        iid: None | str = None,
        text: None | str = None,
        values: None | list[Any] = None,
        create_selections: bool = False,
        undo: bool = True,
    ) -> str:
        """
        Insert an item into the treeview

        Returns:
            str: new item iid
        """
        if not iid:
            iid = self.RI.new_iid()
        if iid in self.RI.rns:
            raise ValueError(f"iid '{iid}' already exists.")
        if iid == parent:
            raise ValueError(f"iid '{iid}' cannot be equal to parent '{parent}'.")
        if parent and parent not in self.RI.rns:
            raise ValueError(f"parent '{parent}' does not exist.")
        if text is None:
            text = iid
        new_node = Node(text, iid, parent)
        datarn = self._get_id_insert_row(index=index, parent=parent)
        self.insert_rows(
            rows=[[new_node] + ([] if values is None else values)],
            idx=datarn,
            heights=[],
            row_index=True,
            create_selections=create_selections,
            fill=False,
            undo=undo,
            tree=True,
        )
        # only relevant here because in rc add rows they are visible and undo sets old open_ids
        if parent and (parent not in self.RI.tree_open_ids or not self.item_displayed(parent)):
            self.hide_rows(datarn, deselect_all=False, data_indexes=True)
        return iid

    def _get_id_insert_row(self, index: int | None, parent: str) -> int:
        if parent:
            chn = self.RI.iid_children(parent)
            if isinstance(index, int):
                index = min(index, len(chn))
                if index == 0:
                    return self.RI.rns[parent] + 1
                else:
                    prev_chld = chn[index - 1]
                    return self.RI.rns[prev_chld] + self.RI.num_descendants(prev_chld) + 1
            else:
                if chn:
                    last_chld = chn[-1]
                    last_chld_rn = self.RI.rns[last_chld]
                    return last_chld_rn + self.RI.num_descendants(last_chld) + 1
                else:
                    return self.RI.rns[parent] + 1
        else:
            if isinstance(index, int):
                if index == 0:
                    return 0
                datarn = self.top_index_row(index)
                return len(self.MT._row_index) if datarn is None else datarn
            else:
                return len(self.MT._row_index)

    def bulk_insert(
        self,
        data: list[list[Any]],
        parent: str = "",
        index: None | int | Literal["end"] = None,
        iid_column: int | None = None,
        text_column: int | None | str = None,
        create_selections: bool = False,
        include_iid_column: bool = True,
        include_text_column: bool = True,
        undo: bool = True,
    ) -> dict[str, int]:
        """
        Insert multiple items into the treeview at once, under the same parent.

        Returns:
            dict:
            - Keys (str): iid
            - Values (int): row numbers
        """
        to_insert = []
        if parent and parent not in self.RI.rns:
            raise ValueError(f"parent '{parent}' does not exist.")
        datarn = self._get_id_insert_row(index=index, parent=parent)
        rns_to_add = {}
        for rn, r in enumerate(data, start=datarn):
            iid = self.RI.new_iid() if iid_column is None else r[iid_column]
            if iid in self.RI.rns:
                raise ValueError(f"iid '{iid}' already exists.")
            new_node = Node(
                r[text_column] if isinstance(text_column, int) else text_column if isinstance(text_column, str) else "",
                iid,
                parent,
            )
            exclude = set()
            if isinstance(iid_column, int) and not include_iid_column:
                exclude.add(iid_column)
            if isinstance(text_column, int) and not include_text_column:
                exclude.add(text_column)
            if exclude:
                to_insert.append([new_node] + [e for c, e in enumerate(r) if c not in exclude])
            else:
                to_insert.append([new_node] + r)
            rns_to_add[iid] = rn
        self.insert_rows(
            rows=to_insert,
            idx=datarn,
            heights=[],
            row_index=True,
            create_selections=create_selections,
            fill=False,
            undo=undo,
            tree=True,
        )
        if parent and (parent not in self.RI.tree_open_ids or not self.item_displayed(parent)):
            self.hide_rows(range(datarn, datarn + len(to_insert)), deselect_all=False, data_indexes=True)
        return rns_to_add

    def item(
        self,
        item: str,
        iid: str | None = None,
        text: str | None = None,
        values: list | None = None,
        open_: bool | None = None,
        undo: bool = True,
        emit_event: bool = True,
        redraw: bool = True,
    ) -> DotDict | Sheet:
        """
        Modify options for item
        If no options are set then returns DotDict of options for item
        Else returns Sheet
        """
        if item not in self.RI.rns:
            raise ValueError(f"Item '{item}' does not exist.")
        get_only = all(param is None for param in (iid, text, values, open_))

        if isinstance(open_, bool):
            if self.RI.iid_children(item):
                if open_:
                    self.RI.tree_open_ids.add(item)
                    if self.item_displayed(item):
                        self.show_rows(
                            rows=map(self.RI.rns.__getitem__, self.RI.get_iid_descendants(item, check_open=True)),
                            redraw=False,
                            deselect_all=False,
                        )
                else:
                    self.RI.tree_open_ids.discard(item)
                    if self.item_displayed(item):
                        self.hide_rows(
                            rows=set(map(self.RI.rns.__getitem__, self.RI.get_iid_descendants(item, check_open=True))),
                            redraw=False,
                            deselect_all=False,
                            data_indexes=True,
                        )
            else:
                self.RI.tree_open_ids.discard(item)

        event_data = None if get_only else self.MT.new_event_dict("edit_table", state=True)

        if isinstance(text, str):
            event_data.cells.index = {self.RI.rns[item]: self.MT._row_index[self.RI.rns[item]].text}
            self.MT._row_index[self.RI.rns[item]].text = text

        if isinstance(values, list):
            span = self.span(self.RI.rns[item], undo=False, emit_event=False)
            event_data = self.set_data(span, data=values, event_data=event_data)

        if isinstance(iid, str):
            if iid in self.RI.rns:
                raise ValueError(f"Cannot rename '{iid}', it already exists.")
            event_data = self.RI.rename_iid(old=item, new=iid, event_data=event_data)

        if event_data:
            if undo and self.MT.undo_enabled:
                self.MT.undo_stack.append(stored_event_dict(event_data))
            self.MT.sheet_modified(event_data, emit_event=emit_event)

        self.set_refresh_timer(redraw=not get_only and redraw)
        if get_only:
            return DotDict(
                text=self.MT._row_index[self.RI.rns[item]].text,
                values=self[self.RI.rns[item]].options(ndim=1).data,
                open_=item in self.RI.tree_open_ids,
            )
        return self

    def itemrow(self, item: str) -> int:
        try:
            return self.RI.rns[item]
        except ValueError as error:
            raise ValueError(f"item '{item}' does not exist.") from error

    def rowitem(self, row: int, data_index: bool = False) -> str | None:
        try:
            if data_index:
                return self.MT._row_index[row].iid
            return self.MT._row_index[self.MT.displayed_rows[row]].iid
        except Exception:
            return None

    def get_children(self, item: None | str = None) -> Generator[str]:
        if item is None:
            yield from map(attrgetter("iid"), self.MT._row_index)
        elif item == "":
            yield from map(attrgetter("iid"), self.RI.gen_top_nodes())
        else:
            yield from self.RI.iid_children(item)

    def tree_traverse(self, item: None | str = None) -> Generator[str]:
        if item is None:
            for tiid in map(attrgetter("iid"), self.RI.gen_top_nodes()):
                yield tiid
                yield from self.RI.get_iid_descendants(tiid)
        elif item == "":
            yield from map(attrgetter("iid"), self.RI.gen_top_nodes())
        else:
            yield from self.RI.iid_children(item)

    get_iids = tree_traverse

    def del_items(self, *items, undo: bool = True) -> Sheet:
        """
        Also deletes all descendants of items
        """
        rows = list(map(self.RI.rns.__getitem__, filter(self.exists, unpack(items))))
        if rows:
            self.del_rows(rows, data_indexes=True, undo=undo)
        return self.set_refresh_timer(rows)

    def top_index_row(self, index: int) -> int:
        try:
            return next(self.RI.rns[n.iid] for i, n in enumerate(self.RI.gen_top_nodes()) if i == index)
        except Exception:
            return None

    def move(
        self,
        item: str,
        parent: str,
        index: int | None = None,
        select: bool = True,
        undo: bool = True,
        emit_event: bool = False,
    ) -> tuple[dict[int, int], dict[int, int], EventDataDict]:
        """
        Moves item to be under parent as child at index
        'parent' can be an empty str which will put the item at top level
        """
        if item not in self.RI.rns:
            raise ValueError(f"Item '{item}' does not exist.")
        elif parent and parent not in self.RI.rns:
            raise ValueError(f"Parent '{parent}' does not exist.")
        elif self.RI.move_pid_causes_recursive_loop(item, parent):
            raise ValueError(f"Item '{item}' causes recursive loop with parent '{parent}.")
        data_new_idxs, disp_new_idxs, event_data = self.MT.move_rows_adjust_options_dict(
            data_new_idxs={},
            data_old_idxs={},
            totalrows=None,
            disp_new_idxs={},
            move_data=True,
            move_heights=True,
            create_selections=select,
            event_data=None,
            node_change={"iids": {item}, "new_par": parent, "index": index},
        )
        if undo:
            self.MT.undo_stack.append(stored_event_dict(event_data))
        self.MT.sheet_modified(event_data, emit_event=emit_event)
        self.set_refresh_timer()
        return data_new_idxs, disp_new_idxs, event_data

    def set_children(
        self,
        parent: str,
        *newchildren: str,
        index: int | None = None,
        select: bool = True,
        undo: bool = True,
        emit_event: bool = False,
    ) -> Sheet:
        """
        Moves everything in '*newchildren' under 'parent'
        """
        if parent and parent not in self.RI.rns:
            raise ValueError(f"Parent '{parent}' does not exist.")
        iids = set()
        for iid in unpack(newchildren):
            if iid not in self.RI.rns:
                raise ValueError(f"Item '{iid}' does not exist.")
            elif self.RI.move_pid_causes_recursive_loop(iid, parent):
                raise ValueError(f"Item '{iid}' causes recursive loop with parent '{parent}.")
            iids.add(iid)
        data_new_idxs, disp_new_idxs, event_data = self.MT.move_rows_adjust_options_dict(
            data_new_idxs={},
            data_old_idxs={},
            totalrows=None,
            disp_new_idxs={},
            move_data=True,
            move_heights=True,
            create_selections=select,
            event_data=None,
            node_change={"iids": iids, "new_par": parent, "index": index},
        )
        if undo:
            self.MT.undo_stack.append(stored_event_dict(event_data))
        self.MT.sheet_modified(event_data, emit_event=emit_event)
        self.set_refresh_timer()
        return data_new_idxs, disp_new_idxs, event_data

    reattach = move

    def exists(self, item: str) -> bool:
        return item in self.RI.rns

    def parent(self, item: str) -> str:
        if item not in self.RI.rns:
            raise ValueError(f"Item '{item}' does not exist.")
        return self.RI.iid_parent(item)

    def index(self, item: str, safety: bool = False) -> int:
        """
        Finds the index of an item amongst it's siblings in the
        treeview.
        """
        if item not in self.RI.rns:
            raise ValueError(f"Item '{item}' does not exist.")
        elif self.RI.iid_parent(item):
            return self.RI.parent_node(item).children.index(item)
        else:
            return next(index for index, iid in enumerate(self.get_children("")) if iid == item)

    def item_displayed(self, item: str) -> bool:
        """
        Check if an item (row) is currently displayed on the sheet
        - Does not check if the item is visible to the user
        """
        if item not in self.RI.rns:
            raise ValueError(f"Item '{item}' does not exist.")
        return bisect_in(self.MT.displayed_rows, self.RI.rns[item])

    def display_item(self, item: str, redraw: bool = False) -> Sheet:
        """
        Ensure that item is displayed in the tree
        - Opens all of an item's ancestors
        - Unlike the ttk treeview 'see' function
          this function does **NOT** scroll to the item
        """
        if not self.item_displayed(item) and self.RI.iid_parent(item):
            self.show_rows(
                rows=self._tree_open(self.RI.get_iid_ancestors(item)),
                redraw=False,
                deselect_all=False,
            )
        return self.set_refresh_timer(redraw)

    def scroll_to_item(self, item: str, redraw: bool = False) -> Sheet:
        """
        Scrolls to an item and ensures that it is displayed
        """
        self.display_item(item, redraw=False)
        self.see(
            row=bisect_left(self.MT.displayed_rows, self.RI.rns[item]),
            keep_xscroll=True,
            redraw=False,
        )
        return self.set_refresh_timer(redraw)

    def selection(self, cells: bool = False) -> list[str]:
        """
        Get currently selected item ids
        """
        return [
            self.MT._row_index[self.MT.displayed_rows[rn]].iid for rn in self.get_selected_rows(get_cells_as_rows=cells)
        ]

    @property
    def tree_selected(self) -> str | None:
        """
        Get the iid at the currently selected box
        """
        if selected := self.selected:
            return self.rowitem(selected.row)
        return None

    def selection_set(self, *items, run_binding: bool = True, redraw: bool = True) -> Sheet:
        boxes_to_hide = tuple(self.MT.selection_boxes)
        self.selection_add(*items, run_binding=False, redraw=False)
        for iid in boxes_to_hide:
            self.MT.hide_selection_box(iid)
        if run_binding:
            self.MT.run_selection_binding("rows")
        return self.set_refresh_timer(redraw)

    def selection_add(self, *items, run_binding: bool = True, redraw: bool = True) -> Sheet:
        to_open = set()
        quick_displayed_check = set(self.MT.displayed_rows)
        for item in filter(self.RI.rns.__contains__, unpack(items)):
            if self.RI.rns[item] not in quick_displayed_check and self.RI.iid_parent(item):
                to_open.update(self.RI.get_iid_ancestors(item))
        if to_open:
            self.show_rows(
                rows=self._tree_open(to_open),
                redraw=False,
                deselect_all=False,
            )
        for startr, endr in consecutive_ranges(
            sorted(
                bisect_left(
                    self.MT.displayed_rows,
                    self.RI.rns[item],
                )
                for item in filter(self.RI.rns.__contains__, unpack(items))
            )
        ):
            self.MT.create_selection_box(
                startr,
                0,
                endr,
                len(self.MT.col_positions) - 1,
                "rows",
                set_current=True,
                ext=True,
            )
        if run_binding:
            self.MT.run_selection_binding("rows")
        return self.set_refresh_timer(redraw)

    def selection_remove(self, *items, redraw: bool = True) -> Sheet:
        for item in unpack(items):
            if item not in self.RI.rns:
                continue
            try:
                self.deselect(bisect_left(self.MT.displayed_rows, self.RI.rns[item]), redraw=False)
            except Exception:
                continue
        return self.set_refresh_timer(redraw)

    def selection_toggle(self, *items, redraw: bool = True) -> Sheet:
        selected = {self.MT._row_index[self.displayed_row_to_data(rn)].iid for rn in self.get_selected_rows()}
        add = []
        remove = []
        for item in unpack(items):
            if item in self.RI.rns:
                if item in selected:
                    remove.append(item)
                else:
                    add.append(item)
        self.selection_remove(*remove, redraw=False)
        self.selection_add(*add, redraw=False)
        return self.set_refresh_timer(redraw)

    def descendants(self, item: str, check_open: bool = False) -> Generator[str]:
        return self.RI.get_iid_descendants(item, check_open=check_open)

    # Functions not in docs

    def event_generate(self, *args, **kwargs) -> None:
        self.MT.event_generate(*args, **kwargs)

    def emit_event(
        self,
        event: str,
        data: EventDataDict | None = None,
    ) -> None:
        if data is None:
            data = EventDataDict()
        data["sheetname"] = self.name
        self.last_event_data = data
        for func in self.bound_events[event]:
            func(data)

    def set_refresh_timer(
        self,
        redraw: bool = True,
        index: bool = True,
        header: bool = True,
    ) -> Sheet:
        if redraw and self.after_redraw_id is None:
            self.after_redraw_id = self.after(
                self.after_redraw_time_ms, lambda: self.after_redraw(index=index, header=header)
            )
        return self

    def after_redraw(self, index: bool = True, header: bool = True) -> None:
        self.MT.main_table_redraw_grid_and_text(redraw_header=header, redraw_row_index=index)
        self.after_redraw_id = None

    def del_options_using_span(
        self,
        span: Span,
        key: str,
        redraw: bool = True,
    ) -> Span:
        rows, cols = self.ranges_from_span(span)
        table, index, header = span.table, span.index, span.header
        # index header
        if header and span.kind in ("cell", "column"):
            self.CH.hide_dropdown_window()
            del_from_options(self.CH.cell_options, key, cols)
        if index and span.kind in ("cell", "row"):
            self.RI.hide_dropdown_window()
            del_from_options(self.RI.cell_options, key, rows)
        # table
        if table and span.kind == "cell":
            del_from_options(self.MT.cell_options, key, product(rows, cols))
        elif table and span.kind == "row":
            del_from_options(self.MT.row_options, key, rows)
        elif table and span.kind == "column":
            del_from_options(self.MT.col_options, key, cols)
        self.set_refresh_timer(redraw)
        return span

    #  ##########       TABLE       ##########

    def del_cell_options_dropdown(self, datarn: int, datacn: int) -> None:
        self.MT.hide_dropdown_window()
        if (datarn, datacn) in self.MT.cell_options and "dropdown" in self.MT.cell_options[(datarn, datacn)]:
            del self.MT.cell_options[(datarn, datacn)]["dropdown"]

    def del_cell_options_checkbox(self, datarn: int, datacn: int) -> None:
        if (datarn, datacn) in self.MT.cell_options and "checkbox" in self.MT.cell_options[(datarn, datacn)]:
            del self.MT.cell_options[(datarn, datacn)]["checkbox"]

    def del_cell_options_dropdown_and_checkbox(self, datarn: int, datacn: int) -> None:
        self.del_cell_options_dropdown(datarn, datacn)
        self.del_cell_options_checkbox(datarn, datacn)

    def del_row_options_dropdown(self, datarn: int) -> None:
        self.MT.hide_dropdown_window()
        if datarn in self.MT.row_options and "dropdown" in self.MT.row_options[datarn]:
            del self.MT.row_options[datarn]["dropdown"]

    def del_row_options_checkbox(self, datarn: int) -> None:
        if datarn in self.MT.row_options and "checkbox" in self.MT.row_options[datarn]:
            del self.MT.row_options[datarn]["checkbox"]

    def del_row_options_dropdown_and_checkbox(self, datarn: int) -> None:
        self.del_row_options_dropdown(datarn)
        self.del_row_options_checkbox(datarn)

    def del_column_options_dropdown(self, datacn: int) -> None:
        self.MT.hide_dropdown_window()
        if datacn in self.MT.col_options and "dropdown" in self.MT.col_options[datacn]:
            del self.MT.col_options[datacn]["dropdown"]

    def del_column_options_checkbox(self, datacn: int) -> None:
        if datacn in self.MT.col_options and "checkbox" in self.MT.col_options[datacn]:
            del self.MT.col_options[datacn]["checkbox"]

    def del_column_options_dropdown_and_checkbox(self, datacn: int) -> None:
        self.del_column_options_dropdown(datacn)
        self.del_column_options_checkbox(datacn)

    #  ##########       INDEX       ##########

    def del_index_cell_options_dropdown(self, datarn: int) -> None:
        self.RI.hide_dropdown_window()
        if datarn in self.RI.cell_options and "dropdown" in self.RI.cell_options[datarn]:
            del self.RI.cell_options[datarn]["dropdown"]

    def del_index_cell_options_checkbox(self, datarn: int) -> None:
        if datarn in self.RI.cell_options and "checkbox" in self.RI.cell_options[datarn]:
            del self.RI.cell_options[datarn]["checkbox"]

    def del_index_cell_options_dropdown_and_checkbox(self, datarn: int) -> None:
        self.del_index_cell_options_dropdown(datarn)
        self.del_index_cell_options_checkbox(datarn)

    #  ##########       HEADER       ##########

    def del_header_cell_options_dropdown(self, datacn: int) -> None:
        self.CH.hide_dropdown_window()
        if datacn in self.CH.cell_options and "dropdown" in self.CH.cell_options[datacn]:
            del self.CH.cell_options[datacn]["dropdown"]

    def del_header_cell_options_checkbox(self, datacn: int) -> None:
        if datacn in self.CH.cell_options and "checkbox" in self.CH.cell_options[datacn]:
            del self.CH.cell_options[datacn]["checkbox"]

    def del_header_cell_options_dropdown_and_checkbox(self, datacn: int) -> None:
        self.del_header_cell_options_dropdown(datacn)
        self.del_header_cell_options_checkbox(datacn)

    # ##########       OLD FUNCTIONS       ##########

    def get_cell_data(self, r: int, c: int, get_displayed: bool = False) -> Any:
        if get_displayed:
            return self.MT.cell_str(r, c, get_displayed=True)
        else:
            return self.MT.get_cell_data(r, c)

    def get_row_data(
        self,
        r: int,
        get_displayed: bool = False,
        get_index: bool = False,
        get_index_displayed: bool = True,
        only_columns: int | Iterator[int] | None = None,
    ) -> list[Any]:
        if only_columns is not None:
            if isinstance(only_columns, int):
                only_columns = (only_columns,)
            elif not is_iterable(only_columns):
                raise ValueError(tksheet_type_error("only_columns", ["int", "iterable", "None"], only_columns))
        if r >= self.MT.total_data_rows():
            raise IndexError(f"Row #{r} is out of range.")

        if r >= len(self.MT.data):
            total_data_cols = self.MT.total_data_cols()
            self.MT.fix_data_len(r, total_data_cols - 1)

        iterable = only_columns if only_columns is not None else range(len(self.MT.data[r]))
        f = partial(self.MT.cell_str, get_displayed=True) if get_displayed else self.MT.get_cell_data
        row = [self.RI.get_cell_data(r, get_displayed=get_index_displayed)] if get_index else []
        row.extend(f(r, c) for c in iterable)
        return row

    def get_column_data(
        self,
        c: int,
        get_displayed: bool = False,
        get_header: bool = False,
        get_header_displayed: bool = True,
        only_rows: int | Iterator[int] | None = None,
    ) -> list[Any]:
        if only_rows is not None:
            if isinstance(only_rows, int):
                only_rows = (only_rows,)
            elif not is_iterable(only_rows):
                raise ValueError(tksheet_type_error("only_rows", ["int", "iterable", "None"], only_rows))
        iterable = only_rows if only_rows is not None else range(len(self.MT.data))
        f = partial(self.MT.cell_str, get_displayed=True) if get_displayed else self.MT.get_cell_data
        col = [self.CH.get_cell_data(c, get_displayed=get_header_displayed)] if get_header else []
        col.extend(f(r, c) for r in iterable)
        return col

    def get_sheet_data(
        self,
        get_displayed: bool = False,
        get_header: bool = False,
        get_index: bool = False,
        get_header_displayed: bool = True,
        get_index_displayed: bool = True,
        only_rows: Iterator[int] | int | None = None,
        only_columns: Iterator[int] | int | None = None,
    ) -> list[list[Any]]:
        if only_rows is not None:
            if isinstance(only_rows, int):
                only_rows = (only_rows,)
            elif not is_iterable(only_rows):
                raise ValueError(tksheet_type_error("only_rows", ["int", "iterable", "None"], only_rows))
        if only_columns is not None:
            if isinstance(only_columns, int):
                only_columns = (only_columns,)
            elif not is_iterable(only_columns):
                raise ValueError(tksheet_type_error("only_columns", ["int", "iterable", "None"], only_columns))

        if get_header:
            maxlen = len(self.MT._headers) if isinstance(self.MT._headers, (list, tuple)) else 0
            data = []
            for rn in only_rows if only_rows is not None else range(len(self.MT.data)):
                r = self.get_row_data(rn, get_displayed=get_displayed, only_columns=only_columns)
                if len(r) > maxlen:
                    maxlen = len(r)
                if get_index:
                    row = [self.RI.get_cell_data(rn, get_displayed=get_index_displayed)]
                    row.extend(r)
                    data.append(row)
                else:
                    data.append(r)
            iterable = only_columns if only_columns is not None else range(maxlen)
            header_row = [""] if get_index else []
            header_row.extend(self.CH.get_cell_data(cn, get_displayed=get_header_displayed) for cn in iterable)
            result = [header_row]
            result.extend(data)
            return result

        else:
            iterable = only_rows if only_rows is not None else range(len(self.MT.data))
            return [
                self.get_row_data(
                    rn,
                    get_displayed=get_displayed,
                    get_index=get_index,
                    get_index_displayed=get_index_displayed,
                    only_columns=only_columns,
                )
                for rn in iterable
            ]

    def yield_sheet_rows(
        self,
        get_displayed: bool = False,
        get_header: bool = False,
        get_index: bool = False,
        get_index_displayed: bool = True,
        get_header_displayed: bool = True,
        only_rows: int | Iterator[int] | None = None,
        only_columns: int | Iterator[int] | None = None,
    ) -> Generator[list[Any]]:
        if only_rows is not None:
            if isinstance(only_rows, int):
                only_rows = (only_rows,)
            elif not is_iterable(only_rows):
                raise ValueError(tksheet_type_error("only_rows", ["int", "iterable", "None"], only_rows))
        if only_columns is not None:
            if isinstance(only_columns, int):
                only_columns = (only_columns,)
            elif not is_iterable(only_columns):
                raise ValueError(tksheet_type_error("only_columns", ["int", "iterable", "None"], only_columns))

        if get_header:
            iterable = only_columns if only_columns is not None else range(self.MT.total_data_cols())
            header_row = [""] if get_index else []
            header_row.extend(self.CH.get_cell_data(c, get_displayed=get_header_displayed) for c in iterable)
            yield header_row

        iterable = only_rows if only_rows is not None else range(len(self.MT.data))
        yield from (
            self.get_row_data(
                r,
                get_displayed=get_displayed,
                get_index=get_index,
                get_index_displayed=get_index_displayed,
                only_columns=only_columns,
            )
            for r in iterable
        )

    def get_header_data(self, c: int, get_displayed: bool = False):
        return self.CH.get_cell_data(datacn=c, get_displayed=get_displayed)

    def get_index_data(self, r: int, get_displayed: bool = False):
        return self.RI.get_cell_data(datarn=r, get_displayed=get_displayed)

    def data_reference(
        self,
        newdataref=None,
        reset_col_positions: bool = True,
        reset_row_positions: bool = True,
        redraw: bool = True,
    ) -> Any:
        self.set_refresh_timer(redraw)
        return self.MT.data_reference(newdataref, reset_col_positions, reset_row_positions)

    def set_cell_data(
        self,
        r: int,
        c: int,
        value: Any = "",
        redraw: bool = True,
        keep_formatting: bool = True,
    ) -> Sheet:
        if not keep_formatting:
            self.MT.delete_cell_format(r, c, clear_values=False)
        self.MT.set_cell_data(r, c, value)
        return self.set_refresh_timer(redraw)

    def set_row_data(
        self,
        r: int,
        values: Sequence[Any] = [],
        add_columns: bool = True,
        redraw: bool = True,
        keep_formatting: bool = True,
    ) -> Sheet:
        if r >= len(self.MT.data):
            raise Exception("Row number is out of range")
        if not keep_formatting:
            self.MT.delete_row_format(r, clear_values=False)
        maxidx = len(self.MT.data[r]) - 1
        if not values:
            self.MT.data[r] = self.MT.get_empty_row_seq(r, len(self.MT.data[r]))
        else:
            if add_columns:
                for c, v in enumerate(values):
                    if c > maxidx:
                        self.MT.data[r].append(v)
                        if self.MT.all_columns_displayed:
                            self.MT.insert_col_positions()
                    else:
                        self.set_cell_data(r=r, c=c, value=v, redraw=False, keep_formatting=keep_formatting)
            else:
                for c, v in enumerate(values):
                    if c > maxidx:
                        self.MT.data[r].append(v)
                    else:
                        self.set_cell_data(r=r, c=c, value=v, redraw=False, keep_formatting=keep_formatting)
        return self.set_refresh_timer(redraw)

    def set_column_data(
        self,
        c: int,
        values: Sequence[Any] = [],
        add_rows: bool = True,
        redraw: bool = True,
        keep_formatting: bool = True,
    ) -> Sheet:
        if not keep_formatting:
            self.MT.delete_column_format(c, clear_values=False)
        if add_rows:
            maxidx = len(self.MT.data) - 1
            total_cols = None
            height = self.MT.get_default_row_height()
            for rn, v in enumerate(values):
                if rn > maxidx:
                    if total_cols is None:
                        total_cols = self.MT.total_data_cols()
                    self.MT.fix_data_len(rn, total_cols - 1)
                    if self.MT.all_rows_displayed:
                        self.MT.insert_row_positions(heights=height)
                    maxidx += 1
                if c >= len(self.MT.data[rn]):
                    self.MT.fix_row_len(rn, c)
                self.set_cell_data(r=rn, c=c, value=v, redraw=False, keep_formatting=keep_formatting)
        else:
            for rn, v in enumerate(values):
                if c >= len(self.MT.data[rn]):
                    self.MT.fix_row_len(rn, c)
                self.set_cell_data(r=rn, c=c, value=v, redraw=False, keep_formatting=keep_formatting)
        return self.set_refresh_timer(redraw)

    def readonly_rows(
        self,
        rows: list[int] | int,
        readonly: bool = True,
        redraw: bool = False,
    ) -> Sheet:
        for r in (rows,) if isinstance(rows, int) else rows:
            set_readonly(self.MT.row_options, r, readonly)
        return self.set_refresh_timer(redraw)

    def readonly_columns(
        self,
        columns: list[int] | int,
        readonly: bool = True,
        redraw: bool = False,
    ) -> Sheet:
        for c in (columns,) if isinstance(columns, int) else columns:
            set_readonly(self.MT.col_options, c, readonly)
        return self.set_refresh_timer(redraw)

    def readonly_cells(
        self,
        row: int = 0,
        column: int = 0,
        cells: list[tuple[int, int]] | None = None,
        readonly: bool = True,
        redraw: bool = False,
    ) -> Sheet:
        if cells:
            for r, c in cells:
                set_readonly(self.MT.cell_options, (r, c), readonly=readonly)
        else:
            set_readonly(self.MT.cell_options, (row, column), readonly=readonly)
        return self.set_refresh_timer(redraw)

    def readonly_header(
        self,
        columns: list[int] | int,
        readonly: bool = True,
        redraw: bool = False,
    ) -> Sheet:
        for c in (columns, int) if isinstance(columns, int) else columns:
            set_readonly(self.CH.cell_options, c, readonly)
        return self.set_refresh_timer(redraw)

    def readonly_index(
        self,
        rows: list[int] | int,
        readonly: bool = True,
        redraw: bool = False,
    ) -> Sheet:
        for r in (rows,) if isinstance(rows, int) else rows:
            set_readonly(self.RI.cell_options, r, readonly)
        return self.set_refresh_timer(redraw)

    def dehighlight_rows(
        self,
        rows: list[int] | Literal["all"],
        redraw: bool = True,
    ) -> Sheet:
        if isinstance(rows, int):
            rows = [rows]
        if not rows or rows == "all":
            for r in self.MT.row_options:
                if "highlight" in self.MT.row_options[r]:
                    del self.MT.row_options[r]["highlight"]

            for r in self.RI.cell_options:
                if "highlight" in self.RI.cell_options[r]:
                    del self.RI.cell_options[r]["highlight"]
        else:
            for r in rows:
                with suppress(Exception):
                    del self.MT.row_options[r]["highlight"]
                with suppress(Exception):
                    del self.RI.cell_options[r]["highlight"]
        return self.set_refresh_timer(redraw)

    def dehighlight_columns(
        self,
        columns: list[int] | Literal["all"],
        redraw: bool = True,
    ) -> Sheet:
        if isinstance(columns, int):
            columns = [columns]
        if not columns or columns == "all":
            for c in self.MT.col_options:
                if "highlight" in self.MT.col_options[c]:
                    del self.MT.col_options[c]["highlight"]

            for c in self.CH.cell_options:
                if "highlight" in self.CH.cell_options[c]:
                    del self.CH.cell_options[c]["highlight"]
        else:
            for c in columns:
                with suppress(Exception):
                    del self.MT.col_options[c]["highlight"]
                with suppress(Exception):
                    del self.CH.cell_options[c]["highlight"]
        return self.set_refresh_timer(redraw)

    def highlight_rows(
        self,
        rows: Iterator[int] | int,
        bg: None | str = None,
        fg: None | str = None,
        highlight_index: bool = True,
        redraw: bool = True,
        end_of_screen: bool = False,
        overwrite: bool = True,
    ) -> Sheet:
        if bg is None and fg is None:
            return
        for r in (rows,) if isinstance(rows, int) else rows:
            add_highlight(self.MT.row_options, r, bg, fg, end_of_screen, overwrite)
        if highlight_index:
            self.highlight_cells(cells=rows, canvas="index", bg=bg, fg=fg, redraw=False)
        return self.set_refresh_timer(redraw)

    def highlight_columns(
        self,
        columns: Iterator[int] | int,
        bg: bool | None | str = False,
        fg: bool | None | str = False,
        highlight_header: bool = True,
        redraw: bool = True,
        overwrite: bool = True,
    ) -> Sheet:
        if bg is False and fg is False:
            return
        for c in (columns,) if isinstance(columns, int) else columns:
            add_highlight(self.MT.col_options, c, bg, fg, None, overwrite)
        if highlight_header:
            self.highlight_cells(cells=columns, canvas="header", bg=bg, fg=fg, redraw=False)
        return self.set_refresh_timer(redraw)

    def highlight_cells(
        self,
        row: int | Literal["all"] = 0,
        column: int | Literal["all"] = 0,
        cells: list[tuple[int, int]] | None = None,
        canvas: Literal["table", "index", "header"] = "table",
        bg: bool | None | str = False,
        fg: bool | None | str = False,
        redraw: bool = True,
        overwrite: bool = True,
    ) -> Sheet:
        if bg is False and fg is False:
            return
        if canvas == "table":
            if cells:
                for r_, c_ in cells:
                    add_highlight(self.MT.cell_options, (r_, c_), bg, fg, None, overwrite)
            else:
                if (
                    isinstance(row, str)
                    and row.lower() == "all"
                    and isinstance(column, str)
                    and column.lower() == "all"
                ):
                    totalrows, totalcols = self.MT.total_data_rows(), self.MT.total_data_cols()
                    for r_ in range(totalrows):
                        for c_ in range(totalcols):
                            add_highlight(self.MT.cell_options, (r_, c_), bg, fg, None, overwrite)
                elif isinstance(row, str) and row.lower() == "all" and isinstance(column, int):
                    for r_ in range(self.MT.total_data_rows()):
                        add_highlight(self.MT.cell_options, (r_, column), bg, fg, None, overwrite)
                elif isinstance(column, str) and column.lower() == "all" and isinstance(row, int):
                    for c_ in range(self.MT.total_data_cols()):
                        add_highlight(self.MT.cell_options, (row, c_), bg, fg, None, overwrite)
                elif isinstance(row, int) and isinstance(column, int):
                    add_highlight(self.MT.cell_options, (row, column), bg, fg, None, overwrite)
        elif canvas in ("row_index", "index"):
            if isinstance(cells, int):
                add_highlight(self.RI.cell_options, cells, bg, fg, None, overwrite)
            elif cells:
                for r_ in cells:
                    add_highlight(self.RI.cell_options, r_, bg, fg, None, overwrite)
            else:
                add_highlight(self.RI.cell_options, row, bg, fg, None, overwrite)
        elif canvas == "header":
            if isinstance(cells, int):
                add_highlight(self.CH.cell_options, cells, bg, fg, None, overwrite)
            elif cells:
                for c_ in cells:
                    add_highlight(self.CH.cell_options, c_, bg, fg, None, overwrite)
            else:
                add_highlight(self.CH.cell_options, column, bg, fg, None, overwrite)
        return self.set_refresh_timer(redraw)

    def dehighlight_cells(
        self,
        row: int | Literal["all"] = 0,
        column: int = 0,
        cells: list[tuple[int, int]] | None = None,
        canvas: Literal["table", "row_index", "index", "header"] = "table",
        all_: bool = False,
        redraw: bool = True,
    ) -> Sheet:
        if row == "all" and canvas == "table":
            for k, v in self.MT.cell_options.items():
                if "highlight" in v:
                    del self.MT.cell_options[k]["highlight"]
        elif row == "all" and canvas == "row_index":
            for k, v in self.RI.cell_options.items():
                if "highlight" in v:
                    del self.RI.cell_options[k]["highlight"]
        elif row == "all" and canvas == "header":
            for k, v in self.CH.cell_options.items():
                if "highlight" in v:
                    del self.CH.cell_options[k]["highlight"]
        if canvas == "table":
            if cells and not all_:
                for t in cells:
                    try:
                        del self.MT.cell_options[t]["highlight"]
                    except Exception:
                        continue
            elif not all_:
                if (
                    row,
                    column,
                ) in self.MT.cell_options and "highlight" in self.MT.cell_options[(row, column)]:
                    del self.MT.cell_options[(row, column)]["highlight"]
            elif all_:
                for k in self.MT.cell_options:
                    if "highlight" in self.MT.cell_options[k]:
                        del self.MT.cell_options[k]["highlight"]
        elif canvas in ("row_index", "index"):
            if cells and not all_:
                for r in cells:
                    try:
                        del self.RI.cell_options[r]["highlight"]
                    except Exception:
                        continue
            elif not all_:
                if row in self.RI.cell_options and "highlight" in self.RI.cell_options[row]:
                    del self.RI.cell_options[row]["highlight"]
            elif all_:
                for r in self.RI.cell_options:
                    if "highlight" in self.RI.cell_options[r]:
                        del self.RI.cell_options[r]["highlight"]
        elif canvas == "header":
            if cells and not all_:
                for c in cells:
                    try:
                        del self.CH.cell_options[c]["highlight"]
                    except Exception:
                        continue
            elif not all_:
                if column in self.CH.cell_options and "highlight" in self.CH.cell_options[column]:
                    del self.CH.cell_options[column]["highlight"]
            elif all_:
                for c in self.CH.cell_options:
                    if "highlight" in self.CH.cell_options[c]:
                        del self.CH.cell_options[c]["highlight"]
        return self.set_refresh_timer(redraw)

    def get_highlighted_cells(
        self,
        canvas: Literal["table", "row_index", "index", "header"] = "table",
    ) -> dict | None:
        if canvas == "table":
            return {k: v["highlight"] for k, v in self.MT.cell_options.items() if "highlight" in v}
        elif canvas in ("row_index", "index"):
            return {k: v["highlight"] for k, v in self.RI.cell_options.items() if "highlight" in v}
        elif canvas == "header":
            return {k: v["highlight"] for k, v in self.CH.cell_options.items() if "highlight" in v}

    def align_cells(
        self,
        row: int = 0,
        column: int = 0,
        cells: list[tuple[int, int]] | dict[tuple[int, int], str] | None = None,
        align: str | None = "global",
        redraw: bool = True,
    ) -> Sheet:
        if isinstance(cells, dict):
            for k, v in cells.items():
                set_align(self.MT.cell_options, k, convert_align(v))
        elif cells:
            align = convert_align(align)
            for k in cells:
                set_align(self.MT.cell_options, k, align)
        else:
            set_align(self.MT.cell_options, (row, column), convert_align(align))
        return self.set_refresh_timer(redraw)

    def align_rows(
        self,
        rows: list[int] | dict[int, str] | int,
        align: str | None = "global",
        align_index: bool = False,
        redraw: bool = True,
    ) -> Sheet:
        if isinstance(rows, dict):
            for k, v in rows.items():
                align = convert_align(v)
                set_align(self.MT.row_options, k, align)
                if align_index:
                    set_align(self.RI.cell_options, k, align)
        elif is_iterable(rows):
            align = convert_align(align)
            for k in rows:
                set_align(self.MT.row_options, k, align)
                if align_index:
                    set_align(self.RI.cell_options, k, align)
        elif isinstance(rows, int):
            set_align(self.MT.row_options, rows, convert_align(align))
            if align_index:
                set_align(self.RI.cell_options, rows, align)
        return self.set_refresh_timer(redraw)

    def align_columns(
        self,
        columns: list[int] | dict[int, str] | int,
        align: str | None = "global",
        align_header: bool = False,
        redraw: bool = True,
    ) -> Sheet:
        if isinstance(columns, dict):
            for k, v in columns.items():
                align = convert_align(v)
                set_align(self.MT.col_options, k, align)
                if align_header:
                    set_align(self.CH.cell_options, k, align)
        elif is_iterable(columns):
            align = convert_align(align)
            for k in columns:
                set_align(self.MT.col_options, k, align)
                if align_header:
                    set_align(self.CH.cell_options, k, align)
        elif isinstance(columns, int):
            set_align(self.MT.col_options, columns, convert_align(align))
            if align_header:
                set_align(self.CH.cell_options, columns, align)
        return self.set_refresh_timer(redraw)

    def align_header(
        self,
        columns: list[int] | dict[int, str] | int,
        align: str | None = "global",
        redraw: bool = True,
    ) -> Sheet:
        if isinstance(columns, dict):
            for k, v in columns.items():
                set_align(self.CH.cell_options, k, convert_align(v))
        elif is_iterable(columns):
            align = convert_align(align)
            for k in columns:
                set_align(self.CH.cell_options, k, align)
        elif isinstance(columns, int):
            set_align(self.CH.cell_options, columns, convert_align(align))
        return self.set_refresh_timer(redraw)

    def align_index(
        self,
        rows: list[int] | dict[int, str] | int,
        align: str | None = "global",
        redraw: bool = True,
    ) -> Sheet:
        if isinstance(rows, dict):
            for k, v in rows.items():
                set_align(self.RI.cell_options, k, convert_align(v))
        elif is_iterable(rows):
            align = convert_align(align)
            for k in rows:
                set_align(self.RI.cell_options, k, align)
        elif isinstance(rows, int):
            set_align(self.RI.cell_options, rows, convert_align(align))
        return self.set_refresh_timer(redraw)

    def get_cell_alignments(self) -> dict:
        return {(r, c): v["align"] for (r, c), v in self.MT.cell_options.items() if "align" in v}

    def get_column_alignments(self) -> dict:
        return {c: v["align"] for c, v in self.MT.col_options.items() if "align" in v}

    def get_row_alignments(self) -> dict:
        return {r: v["align"] for r, v in self.MT.row_options.items() if "align" in v}

    def create_checkbox(
        self,
        r: int | Literal["all"] = 0,
        c: int | Literal["all"] = 0,
        *args,
        **kwargs,
    ) -> None:
        kwargs = get_checkbox_kwargs(*args, **kwargs)
        v = kwargs["checked"]
        d = get_checkbox_dict(**kwargs)
        if isinstance(r, str) and r.lower() == "all" and isinstance(c, int):
            for r_ in range(self.MT.total_data_rows()):
                self._create_checkbox(r_, c, v, d)
        elif isinstance(c, str) and c.lower() == "all" and isinstance(r, int):
            for c_ in range(self.MT.total_data_cols()):
                self._create_checkbox(r, c_, v, d)
        elif isinstance(r, str) and r.lower() == "all" and isinstance(c, str) and c.lower() == "all":
            totalcols = self.MT.total_data_cols()
            for r_ in range(self.MT.total_data_rows()):
                for c_ in range(totalcols):
                    self._create_checkbox(r_, c_, v, d)
        elif isinstance(r, int) and isinstance(c, int):
            self._create_checkbox(r, c, v, d)
        self.set_refresh_timer(kwargs["redraw"])

    def _create_checkbox(self, r: int, c: int, v: bool, d: dict) -> None:
        self.MT.delete_cell_format(r, c, clear_values=False)
        self.del_cell_options_dropdown_and_checkbox(r, c)
        add_to_options(self.MT.cell_options, (r, c), "checkbox", d)
        self.MT.set_cell_data(r, c, v)

    def checkbox_cell(
        self,
        r: int | Literal["all"] = 0,
        c: int | Literal["all"] = 0,
        *args,
        **kwargs,
    ) -> None:
        self.create_checkbox(r=r, c=c, **get_checkbox_kwargs(*args, **kwargs))

    def checkbox_row(
        self,
        r: Iterator[int] | int | Literal["all"] = 0,
        *args,
        **kwargs,
    ) -> None:
        kwargs = get_checkbox_kwargs(*args, **kwargs)
        d = get_checkbox_dict(**kwargs)
        if isinstance(r, str) and r.lower() == "all":
            for r_ in range(self.MT.total_data_rows()):
                self._checkbox_row(r_, kwargs["checked"], d)
        elif isinstance(r, int):
            self._checkbox_row(r, kwargs["checked"], d)
        elif is_iterable(r):
            for r_ in r:
                self._checkbox_row(r_, kwargs["checked"], d)
        self.set_refresh_timer(kwargs["redraw"])

    def _checkbox_row(self, r: int, v: bool, d: dict) -> None:
        self.MT.delete_row_format(r, clear_values=False)
        self.del_row_options_dropdown_and_checkbox(r)
        add_to_options(self.MT.row_options, r, "checkbox", d)
        for c in range(self.MT.total_data_cols()):
            self.MT.set_cell_data(r, c, v)

    def checkbox_column(
        self,
        c: Iterator[int] | int | Literal["all"] = 0,
        *args,
        **kwargs,
    ) -> None:
        kwargs = get_checkbox_kwargs(*args, **kwargs)
        d = get_checkbox_dict(**kwargs)
        if isinstance(c, str) and c.lower() == "all":
            for c in range(self.MT.total_data_cols()):
                self._checkbox_column(c, kwargs["checked"], d)
        elif isinstance(c, int):
            self._checkbox_column(c, kwargs["checked"], d)
        elif is_iterable(c):
            for c_ in c:
                self._checkbox_column(c_, kwargs["checked"], d)
        self.set_refresh_timer(kwargs["redraw"])

    def _checkbox_column(self, c: int, v: bool, d: dict) -> None:
        self.MT.delete_column_format(c, clear_values=False)
        self.del_column_options_dropdown_and_checkbox(c)
        add_to_options(self.MT.col_options, c, "checkbox", d)
        for r in range(self.MT.total_data_rows()):
            self.MT.set_cell_data(r, c, v)

    def create_header_checkbox(
        self,
        c: Iterator[int] | int | Literal["all"] = 0,
        *args,
        **kwargs,
    ) -> None:
        kwargs = get_checkbox_kwargs(*args, **kwargs)
        d = get_checkbox_dict(**kwargs)
        if isinstance(c, str) and c.lower() == "all":
            for c_ in range(self.MT.total_data_cols()):
                self._create_header_checkbox(c_, kwargs["checked"], d)
        elif isinstance(c, int):
            self._create_header_checkbox(c, kwargs["checked"], d)
        elif is_iterable(c):
            for c_ in c:
                self._create_header_checkbox(c_, kwargs["checked"], d)
        self.set_refresh_timer(kwargs["redraw"])

    checkbox_header = create_header_checkbox

    def _create_header_checkbox(self, c: int, v: bool, d: dict) -> None:
        self.del_header_cell_options_dropdown_and_checkbox(c)
        add_to_options(self.CH.cell_options, c, "checkbox", d)
        self.CH.set_cell_data(c, v)

    def create_index_checkbox(
        self,
        r: Iterator[int] | int | Literal["all"] = 0,
        *args,
        **kwargs,
    ) -> None:
        kwargs = get_checkbox_kwargs(*args, **kwargs)
        d = get_checkbox_dict(**kwargs)
        if isinstance(r, str) and r.lower() == "all":
            for r_ in range(self.MT.total_data_rows()):
                self._create_index_checkbox(r_, kwargs["checked"], d)
        elif isinstance(r, int):
            self._create_index_checkbox(r, kwargs["checked"], d)
        elif is_iterable(r):
            for r_ in r:
                self._create_index_checkbox(r_, kwargs["checked"], d)
        self.set_refresh_timer(kwargs["redraw"])

    checkbox_index = create_index_checkbox

    def _create_index_checkbox(self, r: int, v: bool, d: dict) -> None:
        self.del_index_cell_options_dropdown_and_checkbox(r)
        add_to_options(self.RI.cell_options, r, "checkbox", d)
        self.RI.set_cell_data(r, v)

    def delete_checkbox(
        self,
        r: int | Literal["all"] = 0,
        c: int | Literal["all"] = 0,
    ) -> None:
        if isinstance(r, str) and r.lower() == "all" and isinstance(c, int):
            for r_, _ in self.MT.cell_options:
                if "checkbox" in self.MT.cell_options[(r_, c)]:
                    self.del_cell_options_checkbox(r_, c)
        elif isinstance(c, str) and c.lower() == "all" and isinstance(r, int):
            for _, c_ in self.MT.cell_options:
                if "checkbox" in self.MT.cell_options[(r, c_)]:
                    self.del_cell_options_checkbox(r, c_)
        elif isinstance(r, str) and r.lower() == "all" and isinstance(c, str) and c.lower() == "all":
            for r_, c_ in self.MT.cell_options:
                if "checkbox" in self.MT.cell_options[(r_, c_)]:
                    self.del_cell_options_checkbox(r_, c_)
        elif isinstance(r, int) and isinstance(c, int):
            self.del_cell_options_checkbox(r, c)

    def delete_cell_checkbox(
        self,
        r: int | Literal["all"] = 0,
        c: int | Literal["all"] = 0,
    ) -> None:
        self.delete_checkbox(r, c)

    def delete_row_checkbox(
        self,
        r: Iterator[int] | int | Literal["all"] = 0,
    ) -> None:
        if isinstance(r, str) and r.lower() == "all":
            for r_ in self.MT.row_options:
                self.del_table_row_options_checkbox(r_)
        elif isinstance(r, int):
            self.del_table_row_options_checkbox(r)
        elif is_iterable(r):
            for r_ in r:
                self.del_table_row_options_checkbox(r_)

    def delete_column_checkbox(
        self,
        c: Iterator[int] | int | Literal["all"] = 0,
    ) -> None:
        if isinstance(c, str) and c.lower() == "all":
            for c_ in self.MT.col_options:
                self.del_table_column_options_checkbox(c_)
        elif isinstance(c, int):
            self.del_table_column_options_checkbox(c)
        elif is_iterable(c):
            for c_ in c:
                self.del_table_column_options_checkbox(c_)

    def delete_header_checkbox(
        self,
        c: Iterator[int] | int | Literal["all"] = 0,
    ) -> None:
        if isinstance(c, str) and c.lower() == "all":
            for c_ in self.CH.cell_options:
                if "checkbox" in self.CH.cell_options[c_]:
                    self.del_header_cell_options_checkbox(c_)
        if isinstance(c, int):
            self.del_header_cell_options_checkbox(c)
        elif is_iterable(c):
            for c_ in c:
                self.del_header_cell_options_checkbox(c_)

    def delete_index_checkbox(
        self,
        r: Iterator[int] | int | Literal["all"] = 0,
    ) -> None:
        if isinstance(r, str) and r.lower() == "all":
            for r_ in self.RI.cell_options:
                if "checkbox" in self.RI.cell_options[r_]:
                    self.del_index_cell_options_checkbox(r_)
        if isinstance(r, int):
            self.del_index_cell_options_checkbox(r)
        elif is_iterable(r):
            for r_ in r:
                self.del_index_cell_options_checkbox(r_)

    def click_header_checkbox(self, c: int, checked: bool | None = None) -> Sheet:
        kwargs = self.CH.get_cell_kwargs(c, key="checkbox")
        if kwargs:
            if not isinstance(self.MT._headers[c], bool):
                if checked is None:
                    self.MT._headers[c] = False
                else:
                    self.MT._headers[c] = bool(checked)
            else:
                self.MT._headers[c] = not self.MT._headers[c]
        return self

    def click_index_checkbox(self, r: int, checked: bool | None = None) -> Sheet:
        kwargs = self.RI.get_cell_kwargs(r, key="checkbox")
        if kwargs:
            if not isinstance(self.MT._row_index[r], bool):
                if checked is None:
                    self.MT._row_index[r] = False
                else:
                    self.MT._row_index[r] = bool(checked)
            else:
                self.MT._row_index[r] = not self.MT._row_index[r]
        return self

    def get_checkboxes(self) -> dict:
        return {
            **{k: v["checkbox"] for k, v in self.MT.cell_options.items() if "checkbox" in v},
            **{k: v["checkbox"] for k, v in self.MT.row_options.items() if "checkbox" in v},
            **{k: v["checkbox"] for k, v in self.MT.col_options.items() if "checkbox" in v},
        }

    def get_header_checkboxes(self) -> dict:
        return {k: v["checkbox"] for k, v in self.CH.cell_options.items() if "checkbox" in v}

    def get_index_checkboxes(self) -> dict:
        return {k: v["checkbox"] for k, v in self.RI.cell_options.items() if "checkbox" in v}

    def create_dropdown(
        self,
        r: int | Literal["all"] = 0,
        c: int | Literal["all"] = 0,
        *args,
        **kwargs,
    ) -> Sheet:
        kwargs = get_dropdown_kwargs(*args, **kwargs)
        d = get_dropdown_dict(**kwargs)
        v = kwargs["set_value"] if kwargs["set_value"] is not None else kwargs["values"][0] if kwargs["values"] else ""
        edit = kwargs["edit_data"]
        if isinstance(r, str) and r.lower() == "all" and isinstance(c, int):
            for r_ in range(self.MT.total_data_rows()):
                self._create_dropdown(r_, c, v, d, edit)
        elif isinstance(c, str) and c.lower() == "all" and isinstance(r, int):
            for c_ in range(self.MT.total_data_cols()):
                self._create_dropdown(r, c_, v, d, edit)
        elif isinstance(r, str) and r.lower() == "all" and isinstance(c, str) and c.lower() == "all":
            totalcols = self.MT.total_data_cols()
            for r_ in range(self.MT.total_data_rows()):
                for c_ in range(totalcols):
                    self._create_dropdown(r_, c_, v, d, edit)
        elif isinstance(r, int) and isinstance(c, int):
            self._create_dropdown(r, c, v, d, edit)
        return self.set_refresh_timer(kwargs["redraw"])

    def _create_dropdown(self, r: int, c: int, v: Any, d: dict, edit: bool = True) -> None:
        self.del_cell_options_dropdown_and_checkbox(r, c)
        add_to_options(self.MT.cell_options, (r, c), "dropdown", d)
        if edit:
            self.MT.set_cell_data(r, c, v)

    def dropdown_cell(
        self,
        r: int | Literal["all"] = 0,
        c: int | Literal["all"] = 0,
        *args,
        **kwargs,
    ) -> Sheet:
        return self.create_dropdown(r=r, c=c, **get_dropdown_kwargs(*args, **kwargs))

    def dropdown_row(
        self,
        r: Iterator[int] | int | Literal["all"] = 0,
        *args,
        **kwargs,
    ) -> Sheet:
        kwargs = get_dropdown_kwargs(*args, **kwargs)
        d = get_dropdown_dict(**kwargs)
        v = kwargs["set_value"] if kwargs["set_value"] is not None else kwargs["values"][0] if kwargs["values"] else ""
        edit = kwargs["edit_data"]
        if isinstance(r, str) and r.lower() == "all":
            for r_ in range(self.MT.total_data_rows()):
                self._dropdown_row(r_, v, d, edit)
        elif isinstance(r, int):
            self._dropdown_row(r, v, d, edit)
        elif is_iterable(r):
            for r_ in r:
                self._dropdown_row(r_, v, d, edit)
        return self.set_refresh_timer(kwargs["redraw"])

    def _dropdown_row(self, r: int, v: Any, d: dict, edit: bool = True) -> None:
        self.del_row_options_dropdown_and_checkbox(r)
        add_to_options(self.MT.row_options, r, "dropdown", d)
        if edit:
            for c in range(self.MT.total_data_cols()):
                self.MT.set_cell_data(r, c, v)

    def dropdown_column(
        self,
        c: Iterator[int] | int | Literal["all"] = 0,
        *args,
        **kwargs,
    ) -> Sheet:
        kwargs = get_dropdown_kwargs(*args, **kwargs)
        d = get_dropdown_dict(**kwargs)
        v = kwargs["set_value"] if kwargs["set_value"] is not None else kwargs["values"][0] if kwargs["values"] else ""
        edit = kwargs["edit_data"]
        if isinstance(c, str) and c.lower() == "all":
            for c_ in range(self.MT.total_data_cols()):
                self._dropdown_column(c_, v, d, edit)
        elif isinstance(c, int):
            self._dropdown_column(c, v, d, edit)
        elif is_iterable(c):
            for c_ in c:
                self._dropdown_column(c_, v, d, edit)
        return self.set_refresh_timer(kwargs["redraw"])

    def _dropdown_column(self, c: int, v: Any, d: dict, edit: bool = True) -> None:
        self.del_column_options_dropdown_and_checkbox(c)
        add_to_options(self.MT.col_options, c, "dropdown", d)
        if edit:
            for r in range(self.MT.total_data_rows()):
                self.MT.set_cell_data(r, c, v)

    def create_header_dropdown(
        self,
        c: Iterator[int] | int | Literal["all"] = 0,
        *args,
        **kwargs,
    ) -> Sheet:
        kwargs = get_dropdown_kwargs(*args, **kwargs)
        d = get_dropdown_dict(**kwargs)
        v = kwargs["set_value"] if kwargs["set_value"] is not None else kwargs["values"][0] if kwargs["values"] else ""
        edit = kwargs["edit_data"]
        if isinstance(c, str) and c.lower() == "all":
            for c_ in range(self.MT.total_data_cols()):
                self._create_header_dropdown(c_, v, d, edit)
        elif isinstance(c, int):
            self._create_header_dropdown(c, v, d, edit)
        elif is_iterable(c):
            for c_ in c:
                self._create_header_dropdown(c_, v, d, edit)
        return self.set_refresh_timer(kwargs["redraw"])

    dropdown_header = create_header_dropdown

    def _create_header_dropdown(self, c: int, v: Any, d: dict, edit: bool = True) -> None:
        self.del_header_cell_options_dropdown_and_checkbox(c)
        add_to_options(self.CH.cell_options, c, "dropdown", d)
        if edit:
            self.CH.set_cell_data(c, v)

    def create_index_dropdown(
        self,
        r: Iterator[int] | int | Literal["all"] = 0,
        *args,
        **kwargs,
    ) -> Sheet:
        kwargs = get_dropdown_kwargs(*args, **kwargs)
        d = get_dropdown_dict(**kwargs)
        v = kwargs["set_value"] if kwargs["set_value"] is not None else kwargs["values"][0] if kwargs["values"] else ""
        edit = kwargs["edit_data"]
        if isinstance(r, str) and r.lower() == "all":
            for r_ in range(self.MT.total_data_rows()):
                self._create_index_dropdown(r_, v, d, edit)
        elif isinstance(r, int):
            self._create_index_dropdown(r, v, d, edit)
        elif is_iterable(r):
            for r_ in r:
                self._create_index_dropdown(r_, v, d, edit)
        return self.set_refresh_timer(kwargs["redraw"])

    dropdown_index = create_index_dropdown

    def _create_index_dropdown(self, r: int, v: Any, d: dict, edit: bool = True) -> None:
        self.del_index_cell_options_dropdown_and_checkbox(r)
        add_to_options(self.RI.cell_options, r, "dropdown", d)
        if edit:
            self.RI.set_cell_data(r, v)

    def delete_dropdown(
        self,
        r: int | Literal["all"] = 0,
        c: int | Literal["all"] = 0,
    ) -> None:
        if isinstance(r, str) and r.lower() == "all" and isinstance(c, int):
            for r_, _ in self.MT.cell_options:
                if "dropdown" in self.MT.cell_options[(r_, c)]:
                    self.del_cell_options_dropdown(r_, c)
        elif isinstance(c, str) and c.lower() == "all" and isinstance(r, int):
            for _, c_ in self.MT.cell_options:
                if "dropdown" in self.MT.cell_options[(r, c_)]:
                    self.del_cell_options_dropdown(r, c_)
        elif isinstance(r, str) and r.lower() == "all" and isinstance(c, str) and c.lower() == "all":
            for r_, c_ in self.MT.cell_options:
                if "dropdown" in self.MT.cell_options[(r_, c_)]:
                    self.del_cell_options_dropdown(r_, c_)
        elif isinstance(r, int) and isinstance(c, int):
            self.del_cell_options_dropdown(r, c)

    def delete_cell_dropdown(
        self,
        r: int | Literal["all"] = 0,
        c: int | Literal["all"] = 0,
    ) -> None:
        self.delete_dropdown(r=r, c=c)

    def delete_row_dropdown(
        self,
        r: Iterator[int] | int | Literal["all"] = "all",
    ) -> None:
        if isinstance(r, str) and r.lower() == "all":
            for r_ in self.MT.row_options:
                if "dropdown" in self.MT.row_options[r_]:
                    self.del_table_row_options_dropdown(r_)
        elif isinstance(r, int):
            self.del_table_row_options_dropdown(r)
        elif is_iterable(r):
            for r_ in r:
                self.del_table_row_options_dropdown(r_)

    def delete_column_dropdown(
        self,
        c: Iterator[int] | int | Literal["all"] = "all",
    ) -> None:
        if isinstance(c, str) and c.lower() == "all":
            for c_ in self.MT.col_options:
                if "dropdown" in self.MT.col_options[c_]:
                    self.del_column_options_dropdown(datacn=c_)
        elif isinstance(c, int):
            self.del_column_options_dropdown(datacn=c)
        elif is_iterable(c):
            for c_ in c:
                self.del_column_options_dropdown(datacn=c_)

    def delete_header_dropdown(self, c: Iterator[int] | int | Literal["all"]) -> None:
        if isinstance(c, str) and c.lower() == "all":
            for c_ in self.CH.cell_options:
                if "dropdown" in self.CH.cell_options[c_]:
                    self.del_header_cell_options_dropdown(c_)
        elif isinstance(c, int):
            self.del_header_cell_options_dropdown(c)
        elif is_iterable(c):
            for c_ in c:
                self.del_header_cell_options_dropdown(c_)

    def delete_index_dropdown(self, r: Iterator[int] | int | Literal["all"]) -> None:
        if isinstance(r, str) and r.lower() == "all":
            for r_ in self.RI.cell_options:
                if "dropdown" in self.RI.cell_options[r_]:
                    self.del_index_cell_options_dropdown(r_)
        elif isinstance(r, int):
            self.del_index_cell_options_dropdown(r)
        elif is_iterable(r):
            for r_ in r:
                self.del_index_cell_options_dropdown(r_)

    def get_dropdowns(self) -> dict:
        d = {
            **{k: v["dropdown"] for k, v in self.MT.cell_options.items() if "dropdown" in v},
            **{k: v["dropdown"] for k, v in self.MT.row_options.items() if "dropdown" in v},
            **{k: v["dropdown"] for k, v in self.MT.col_options.items() if "dropdown" in v},
        }
        return d

    def get_header_dropdowns(self) -> dict:
        d = {k: v["dropdown"] for k, v in self.CH.cell_options.items() if "dropdown" in v}
        return d

    def get_index_dropdowns(self) -> dict:
        d = {k: v["dropdown"] for k, v in self.RI.cell_options.items() if "dropdown" in v}
        return d

    def set_dropdown_values(
        self,
        r: int = 0,
        c: int = 0,
        set_existing_dropdown: bool = False,
        values: list[Any] | None = None,
        set_value: Any = None,
    ) -> Sheet:
        if values is None:
            values = []
        if set_existing_dropdown:
            if self.MT.dropdown.open:
                r_, c_ = self.MT.dropdown.get_coords()
            else:
                raise Exception("No dropdown box is currently open")
        else:
            r_ = r
            c_ = c
        kwargs = self.MT.get_cell_kwargs(r, c, key="dropdown")
        kwargs["values"] = values
        if self.MT.dropdown.open:
            self.MT.dropdown.window.values(values)
        if set_value is not None:
            self.set_cell_data(r_, c_, set_value)
            if self.MT.dropdown.open:
                self.MT.text_editor.window.set_text(set_value)
        return self

    def set_header_dropdown_values(
        self,
        c: int = 0,
        set_existing_dropdown: bool = False,
        values: list[Any] | None = None,
        set_value: Any = None,
    ) -> Sheet:
        if values is None:
            values = []
        if set_existing_dropdown:
            if self.CH.dropdown.open:
                c_ = self.CH.dropdown.get_coords()
            else:
                raise Exception("No dropdown box is currently open")
        else:
            c_ = c
        kwargs = self.CH.get_cell_kwargs(c_, key="dropdown")
        if kwargs:
            kwargs["values"] = values
            if self.CH.dropdown.open:
                self.CH.dropdown.window.values(values)
            if set_value is not None:
                self.MT.headers(newheaders=set_value, index=c_)

        return self

    def set_index_dropdown_values(
        self,
        r: int = 0,
        set_existing_dropdown: bool = False,
        values: list[Any] | None = None,
        set_value: Any = None,
    ) -> Sheet:
        if values is None:
            values = []
        if set_existing_dropdown:
            if self.RI.current_dropdown_window is not None:
                r_ = self.RI.current_dropdown_window.r
            else:
                raise Exception("No dropdown box is currently open")
        else:
            r_ = r
        kwargs = self.RI.get_cell_kwargs(r_, key="dropdown")
        if kwargs:
            kwargs["values"] = values
            if self.RI.dropdown.open:
                self.RI.dropdown.window.values(values)
            if set_value is not None:
                self.MT.row_index(newindex=set_value, index=r_)
        return self

    def get_dropdown_values(self, r: int = 0, c: int = 0) -> None | list:
        kwargs = self.MT.get_cell_kwargs(r, c, key="dropdown")
        if kwargs:
            return kwargs["values"]

    def get_header_dropdown_values(self, c: int = 0) -> None | list:
        kwargs = self.CH.get_cell_kwargs(c, key="dropdown")
        if kwargs:
            return kwargs["values"]

    def get_index_dropdown_values(self, r: int = 0) -> None | list:
        kwargs = self.RI.get_cell_kwargs(r, key="dropdown")
        if kwargs:
            return kwargs["values"]

    def dropdown_functions(
        self,
        r: int,
        c: int,
        selection_function: Literal[""] | Callable | None = "",
        modified_function: Literal[""] | Callable | None = "",
    ) -> None | dict:
        kwargs = self.MT.get_cell_kwargs(r, c, key="dropdown")
        if kwargs:
            if selection_function != "":
                kwargs["select_function"] = selection_function
            if modified_function != "":
                kwargs["modified_function"] = modified_function
            return kwargs

    def header_dropdown_functions(
        self,
        c: int,
        selection_function: Literal[""] | Callable | None = "",
        modified_function: Literal[""] | Callable | None = "",
    ) -> None | dict:
        kwargs = self.CH.get_cell_kwargs(c, key="dropdown")
        if selection_function != "":
            kwargs["selection_function"] = selection_function
        if modified_function != "":
            kwargs["modified_function"] = modified_function
        return kwargs

    def index_dropdown_functions(
        self,
        r: int,
        selection_function: Literal[""] | Callable | None = "",
        modified_function: Literal[""] | Callable | None = "",
    ) -> None | dict:
        kwargs = self.RI.get_cell_kwargs(r, key="dropdown")
        if selection_function != "":
            kwargs["select_function"] = selection_function
        if modified_function != "":
            kwargs["modified_function"] = modified_function
        return kwargs

    def get_dropdown_value(self, r: int = 0, c: int = 0) -> Any:
        if self.MT.get_cell_kwargs(r, c, key="dropdown"):
            return self.get_cell_data(r, c)

    def get_header_dropdown_value(self, c: int = 0) -> Any:
        if self.CH.get_cell_kwargs(c, key="dropdown"):
            return self.MT._headers[c]

    def get_index_dropdown_value(self, r: int = 0) -> Any:
        if self.RI.get_cell_kwargs(r, key="dropdown"):
            return self.MT._row_index[r]

    def format_cell(
        self,
        r: int | Literal["all"],
        c: int | Literal["all"],
        formatter_options: dict | None = None,
        formatter_class: Any = None,
        redraw: bool = True,
        **kwargs,
    ) -> Sheet:
        if formatter_options is None:
            formatter_options = {}
        kwargs = fix_format_kwargs({"formatter": formatter_class, **formatter_options, **kwargs})
        if isinstance(r, str) and r.lower() == "all" and isinstance(c, int):
            for r_ in range(self.MT.total_data_rows()):
                self._format_cell(r_, c, kwargs)
        elif isinstance(c, str) and c.lower() == "all" and isinstance(r, int):
            for c_ in range(self.MT.total_data_cols()):
                self._format_cell(r, c_, kwargs)
        elif isinstance(r, str) and r.lower() == "all" and isinstance(c, str) and c.lower() == "all":
            for r_ in range(self.MT.total_data_rows()):
                for c_ in range(self.MT.total_data_cols()):
                    self._format_cell(r_, c_, kwargs)
        else:
            self._format_cell(r, c, kwargs)
        return self.set_refresh_timer(redraw)

    def _format_cell(self, r: int, c: int, d: dict) -> None:
        v = d["value"] if "value" in d else self.MT.get_cell_data(r, c)
        self.del_cell_options_checkbox(r, c)
        add_to_options(self.MT.cell_options, (r, c), "format", d)
        self.MT.set_cell_data(r, c, value=v, kwargs=d)

    def delete_cell_format(
        self,
        r: Literal["all"] | int = "all",
        c: Literal["all"] | int = "all",
        clear_values: bool = False,
    ) -> Sheet:
        if isinstance(r, str) and r.lower() == "all" and isinstance(c, int):
            for r_, _ in self.MT.cell_options:
                if "format" in self.MT.cell_options[(r_, c)]:
                    self.MT.delete_cell_format(r_, c, clear_values=clear_values)
        elif isinstance(c, str) and c.lower() == "all" and isinstance(r, int):
            for _, c_ in self.MT.cell_options:
                if "format" in self.MT.cell_options[(r, c_)]:
                    self.MT.delete_cell_format(r, c_, clear_values=clear_values)
        elif isinstance(r, str) and r.lower() == "all" and isinstance(c, str) and c.lower() == "all":
            for r_, c_ in self.MT.cell_options:
                if "format" in self.MT.cell_options[(r_, c_)]:
                    self.MT.delete_cell_format(r_, c_, clear_values=clear_values)
        else:
            self.MT.delete_cell_format(r, c, clear_values=clear_values)
        return self

    def format_row(
        self,
        r: Iterator[int] | int | Literal["all"],
        formatter_options: dict | None = None,
        formatter_class: Any = None,
        redraw: bool = True,
        **kwargs,
    ) -> Sheet:
        if formatter_options is None:
            formatter_options = {}
        kwargs = fix_format_kwargs({"formatter": formatter_class, **formatter_options, **kwargs})
        if isinstance(r, str) and r.lower() == "all":
            for r_ in range(len(self.MT.data)):
                self._format_row(r_, kwargs)
        elif is_iterable(r):
            for r_ in r:
                self._format_row(r_, kwargs)
        else:
            self._format_row(r, kwargs)
        return self.set_refresh_timer(redraw)

    def _format_row(self, r: int, d: dict) -> None:
        self.del_row_options_checkbox(r)
        add_to_options(self.MT.row_options, r, "format", d)
        for c in range(self.MT.total_data_cols()):
            self.MT.set_cell_data(
                r,
                c,
                value=d["value"] if "value" in d else self.MT.get_cell_data(r, c),
                kwargs=d,
            )

    def delete_row_format(
        self,
        r: Iterator[int] | int | Literal["all"] = "all",
        clear_values: bool = False,
    ) -> Sheet:
        if is_iterable(r):
            for r_ in r:
                self.MT.delete_row_format(r_, clear_values=clear_values)
        else:
            self.MT.delete_row_format(r, clear_values=clear_values)
        return self

    def format_column(
        self,
        c: Iterator[int] | int | Literal["all"],
        formatter_options: dict | None = None,
        formatter_class: Any = None,
        redraw: bool = True,
        **kwargs,
    ) -> Sheet:
        if formatter_options is None:
            formatter_options = {}
        kwargs = fix_format_kwargs({"formatter": formatter_class, **formatter_options, **kwargs})
        if isinstance(c, str) and c.lower() == "all":
            for c_ in range(self.MT.total_data_cols()):
                self._format_column(c_, kwargs)
        elif is_iterable(c):
            for c_ in c:
                self._format_column(c_, kwargs)
        else:
            self._format_column(c, kwargs)
        return self.set_refresh_timer(redraw)

    def _format_column(self, c: int, d: dict) -> None:
        self.del_column_options_checkbox(c)
        add_to_options(self.MT.col_options, c, "format", d)
        for r in range(self.MT.total_data_rows()):
            self.MT.set_cell_data(
                r,
                c,
                value=d["value"] if "value" in d else self.MT.get_cell_data(r, c),
                kwargs=d,
            )

    def delete_column_format(
        self,
        c: Iterator[int] | int | Literal["all"] = "all",
        clear_values: bool = False,
    ) -> Sheet:
        if is_iterable(c):
            for c_ in c:
                self.MT.delete_column_format(c_, clear_values=clear_values)
        else:
            self.MT.delete_column_format(c, clear_values=clear_values)
        return self


class Dropdown(Sheet):
    def __init__(
        self,
        parent: tk.Misc,
        r: int,
        c: int,
        bg: str,
        fg: str,
        select_bg: str,
        select_fg: str,
        ops: dict,
        outline_color: str,
        width: int | None = None,
        height: int | None = None,
        font: None | tuple[str, int, str] = None,
        outline_thickness: int = 2,
        values: list[Any] | None = None,
        close_dropdown_window: Callable | None = None,
        search_function: Callable = dropdown_search_function,
        modified_function: None | Callable = None,
        arrowkey_RIGHT: Callable | None = None,
        arrowkey_LEFT: Callable | None = None,
        align: str = "nw",
        # False for using r, c
        # "r" for r
        # "c" for c
        single_index: Literal["r", "c"] | bool = False,
    ) -> None:
        Sheet.__init__(
            self,
            parent=parent,
            name="!SheetDropdown",
            outline_thickness=outline_thickness,
            show_horizontal_grid=True,
            show_vertical_grid=False,
            show_header=False,
            show_row_index=False,
            show_top_left=False,
            empty_horizontal=0,
            empty_vertical=0,
            horizontal_grid_to_end_of_window=True,
            show_selected_cells_border=False,
            set_cell_sizes_on_zoom=True,
            scrollbar_show_arrows=False,
            rounded_boxes=False,
        )
        self.parent = parent
        self.close_dropdown_window = close_dropdown_window
        self.arrowkey_RIGHT = arrowkey_RIGHT
        self.arrowkey_LEFT = arrowkey_LEFT
        self.single_index = single_index
        self.bind("<Motion>", self.mouse_motion)
        self.bind("<ButtonPress-1>", self.b1)
        self.bind("<Up>", self.arrowkey_UP)
        self.bind("<Tab>", self.arrowkey_RIGHT)
        self.bind("<Right>", self.arrowkey_RIGHT)
        self.bind("<Down>", self.arrowkey_DOWN)
        self.bind("<Left>", self.arrowkey_LEFT)
        self.bind("<Prior>", self.arrowkey_UP)
        self.bind("<Next>", self.arrowkey_DOWN)
        self.bind("<Return>", self.b1)
        self.reset(
            r=r,
            c=c,
            bg=bg,
            fg=fg,
            select_bg=select_bg,
            select_fg=select_fg,
            width=width,
            height=height,
            font=font,
            ops=ops,
            outline_color=outline_color,
            align=align,
            values=[] if values is None else values,
            search_function=search_function,
            modified_function=modified_function,
        )

    def reset(
        self,
        r: int,
        c: int,
        bg: str,
        fg: str,
        select_bg: str,
        select_fg: str,
        width: int,
        height: int,
        font: tuple[str, int, str],
        ops: DotDict,
        outline_color: str,
        align: str,
        values: list[Any] | None = None,
        search_function: Callable = dropdown_search_function,
        modified_function: None | Callable = None,
    ) -> None:
        self.set_yview(0.0)
        self.deselect(redraw=False)
        self.r = r
        self.c = c
        self.row = -1
        self.search_function = search_function
        self.modified_function = modified_function
        self.height_and_width(height=height, width=width)
        self.table_align(align)
        self.set_options(
            outline_color=outline_color,
            table_grid_fg=fg,
            table_selected_cells_border_fg=fg,
            table_selected_cells_bg=select_bg,
            table_selected_rows_border_fg=fg,
            table_selected_rows_bg=select_bg,
            table_selected_rows_fg=select_fg,
            table_selected_box_cells_fg=select_bg,
            table_selected_box_rows_fg=select_bg,
            font=font,
            table_fg=fg,
            table_bg=bg,
            **{k: ops[k] for k in scrollbar_options_keys},
        )
        self.values(values, width=width)

    def arrowkey_UP(self, event: Any = None) -> None:
        if self.row > 0:
            self.row -= 1
        else:
            self.row = 0
        self.see(self.row, 0, redraw=False)
        self.select_row(self.row)

    def arrowkey_DOWN(self, event: Any = None) -> None:
        if len(self.MT.data) - 1 > self.row:
            self.row += 1
        self.see(self.row, 0, redraw=False)
        self.select_row(self.row)

    def search_and_see(self, event: Any = None) -> str:
        if self.search_function is not None:
            rn = self.search_function(search_for=rf"{event['value']}", data=(r[0] for r in self.MT.data))
            if isinstance(rn, int):
                self.row = rn
                self.see(self.row, 0, redraw=False)
                self.select_row(self.row)
                return self.MT.data[rn][0]

    def mouse_motion(self, event: Any) -> None:
        row = self.identify_row(event, exclude_index=True, allow_end=False)
        if row is not None and row != self.row:
            self.row = row
            self.select_row(self.row)

    def _reselect(self) -> None:
        rows = self.get_selected_rows()
        if rows:
            self.select_row(next(iter(rows)))

    def b1(self, event: Any = None) -> None:
        if event is None:
            row = None
        elif event.keysym == "Return":
            row = self.get_selected_rows()
            row = None if not row else next(iter(row))
        else:
            row = self.identify_row(event, exclude_index=True, allow_end=False)
        if self.single_index:
            if row is None:
                self.close_dropdown_window(self.r if self.single_index == "r" else self.c)
            else:
                self.close_dropdown_window(
                    self.r if self.single_index == "r" else self.c,
                    self.get_cell_data(row, 0),
                )
        else:
            if row is None:
                self.close_dropdown_window(self.r, self.c)
            else:
                self.close_dropdown_window(self.r, self.c, self.get_cell_data(row, 0))

    def get_coords(self) -> int | tuple[int, int]:
        if self.single_index == "r":
            return self.r
        elif self.single_index == "c":
            return self.c
        return self.r, self.c

    def values(
        self,
        values: list[Any] | None = None,
        redraw: bool = True,
        width: int | None = None,
    ) -> None:
        if values is None:
            values = []
        self.set_sheet_data(
            [[v] for v in values],
            reset_col_positions=False,
            reset_row_positions=False,
            redraw=True,
            verify=False,
        )
        self.MT.main_table_redraw_grid_and_text(True, True, True, True, True)
        if self.yscroll_showing:
            self.set_all_cell_sizes_to_text(redraw=redraw, width=width - self.yscroll.winfo_width() - 4, slim=True)
        else:
            self.set_all_cell_sizes_to_text(redraw=redraw, width=width - 4, slim=True)
