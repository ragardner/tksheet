from __future__ import annotations

import tkinter as tk
from bisect import bisect_left
from collections import defaultdict, deque
from collections.abc import (
    Callable,
    Generator,
    Hashable,
    Iterator,
    Sequence,
)
from itertools import accumulate, chain, islice, product, repeat
from timeit import default_timer
from tkinter import ttk
from typing import Literal

from .column_headers import ColumnHeaders
from .functions import (
    add_highlight,
    add_to_options,
    consecutive_ranges,
    convert_align,
    data_to_displayed_idxs,
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
    new_tk_event,
    num2alpha,
    pickled_event_dict,
    pop_positions,
    set_align,
    set_readonly,
    span_froms,
    span_ranges,
    tksheet_type_error,
    unpack,
)
from .main_table import MainTable
from .other_classes import (
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
from .sheet_options import (
    new_sheet_options,
)
from .themes import (
    theme_black,
    theme_dark,
    theme_dark_blue,
    theme_dark_green,
    theme_light_blue,
    theme_light_green,
)
from .top_left_rectangle import TopLeftRectangle
from .types import (
    CreateSpanTypes,
)
from .vars import (
    USER_OS,
    backwards_compatibility_keys,
    emitted_events,
    named_span_types,
    rc_binding,
    scrollbar_options_keys,
)


class Sheet(tk.Frame):
    def __init__(
        self,
        parent: tk.Misc,
        name: str = "!sheet",
        show_table: bool = True,
        show_top_left: bool = True,
        show_row_index: bool = True,
        show_header: bool = True,
        show_x_scrollbar: bool = True,
        show_y_scrollbar: bool = True,
        width: int | None = None,
        height: int | None = None,
        headers: None | list[object] = None,
        header: None | list[object] = None,
        row_index: None | list[object] = None,
        index: None | list[object] = None,
        default_header: Literal["letters", "numbers", "both"] = "letters",
        default_row_index: Literal["letters", "numbers", "both"] = "numbers",
        data_reference: None | Sequence[Sequence[object]] = None,
        data: None | Sequence[Sequence[object]] = None,
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
        max_column_width: Literal["inf"] | float = "inf",
        max_row_height: Literal["inf"] | float = "inf",
        max_header_height: Literal["inf"] | float = "inf",
        max_index_width: Literal["inf"] | float = "inf",
        after_redraw_time_ms: int = 20,
        set_all_heights_and_widths: bool = False,
        zoom: int = 100,
        align: str = "w",
        header_align: str = "center",
        row_index_align: str | None = None,
        index_align: str = "center",
        displayed_columns: list[int] = [],
        all_columns_displayed: bool = True,
        displayed_rows: list[int] = [],
        all_rows_displayed: bool = True,
        to_clipboard_delimiter: str = "\t",
        to_clipboard_quotechar: str = '"',
        to_clipboard_lineterminator: str = "\n",
        from_clipboard_delimiters: list[str] | str = ["\t"],
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
        selected_rows_to_end_of_window: bool = False,
        horizontal_grid_to_end_of_window: bool = False,
        vertical_grid_to_end_of_window: bool = False,
        show_vertical_grid: bool = True,
        show_horizontal_grid: bool = True,
        display_selected_fg_over_highlights: bool = False,
        show_selected_cells_border: bool = True,
        treeview: bool = False,
        treeview_indent: str | int = "6",
        rounded_boxes: bool = True,
        alternate_color: str = "",
        # colors
        outline_thickness: int = 0,
        outline_color: str = theme_light_blue["outline_color"],
        theme: str = "light blue",
        frame_bg: str = theme_light_blue["table_bg"],
        popup_menu_fg: str = theme_light_blue["popup_menu_fg"],
        popup_menu_bg: str = theme_light_blue["popup_menu_bg"],
        popup_menu_highlight_bg: str = theme_light_blue["popup_menu_highlight_bg"],
        popup_menu_highlight_fg: str = theme_light_blue["popup_menu_highlight_fg"],
        table_grid_fg: str = theme_light_blue["table_grid_fg"],
        table_bg: str = theme_light_blue["table_bg"],
        table_fg: str = theme_light_blue["table_fg"],
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
        index_selected_cells_bg: str = theme_light_blue["index_selected_cells_bg"],
        index_selected_cells_fg: str = theme_light_blue["index_selected_cells_fg"],
        index_selected_rows_bg: str = theme_light_blue["index_selected_rows_bg"],
        index_selected_rows_fg: str = theme_light_blue["index_selected_rows_fg"],
        index_hidden_rows_expander_bg: str = theme_light_blue["index_hidden_rows_expander_bg"],
        header_bg: str = theme_light_blue["header_bg"],
        header_border_fg: str = theme_light_blue["header_border_fg"],
        header_grid_fg: str = theme_light_blue["header_grid_fg"],
        header_fg: str = theme_light_blue["header_fg"],
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
        vertical_scroll_borderwidth: int = 1,
        horizontal_scroll_borderwidth: int = 1,
        vertical_scroll_gripcount: int = 0,
        horizontal_scroll_gripcount: int = 0,
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
    ) -> None:
        tk.Frame.__init__(
            self,
            parent,
            background=frame_bg,
            highlightthickness=outline_thickness,
            highlightbackground=outline_color,
            highlightcolor=outline_color,
        )
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
        for k, v in locals().items():
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
        self.dropdown_class = Dropdown
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
        self.RI = RowIndex(
            parent=self,
            row_index_align=(
                convert_align(row_index_align) if row_index_align is not None else convert_align(index_align)
            ),
            default_row_index=default_row_index,
        )
        self.CH = ColumnHeaders(
            parent=self,
            default_header=default_header,
            header_align=convert_align(header_align),
        )
        self.MT = MainTable(
            parent=self,
            max_column_width=max_column_width,
            max_header_height=max_header_height,
            max_row_height=max_row_height,
            max_index_width=max_index_width,
            show_index=show_row_index,
            show_header=show_header,
            column_headers_canvas=self.CH,
            row_index_canvas=self.RI,
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
        self.TL = TopLeftRectangle(
            parent=self,
            main_canvas=self.MT,
            row_index_canvas=self.RI,
            header_canvas=self.CH,
        )
        self.unique_id = f"{default_timer()}{self.winfo_id()}".replace(".", "")
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
        self.xscroll = ttk.Scrollbar(
            self,
            command=self.MT._xscrollbar,
            orient="horizontal",
            style=f"Sheet{self.unique_id}.Horizontal.TScrollbar",
        )
        if show_top_left:
            self.TL.grid(row=0, column=0)
        if show_table:
            self.MT.grid(row=1, column=1, sticky="nswe")
            self.MT["xscrollcommand"] = self.xscroll.set
            self.MT["yscrollcommand"] = self.yscroll.set
        if show_row_index:
            self.RI.grid(row=1, column=0, sticky="nswe")
            self.RI["yscrollcommand"] = self.yscroll.set
        if show_header:
            self.CH.grid(row=0, column=1, sticky="nswe")
            self.CH["xscrollcommand"] = self.xscroll.set
        if show_x_scrollbar:
            self.xscroll.grid(row=2, column=0, columnspan=2, sticky="nswe")
            self.xscroll_showing = True
            self.xscroll_disabled = False
        else:
            self.xscroll_showing = False
            self.xscroll_disabled = True
        if show_y_scrollbar:
            self.yscroll.grid(row=0, column=2, rowspan=3, sticky="nswe")
            self.yscroll_showing = True
            self.yscroll_disabled = False
        else:
            self.yscroll_showing = False
            self.yscroll_disabled = True
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
        value: object,
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
        newheaders: object = None,
        index: None | int = None,
        reset_col_positions: bool = False,
        show_headers_if_not_sheet: bool = True,
        redraw: bool = True,
    ) -> object:
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
        value: object,
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
        newindex: object = None,
        index: None | int = None,
        reset_row_positions: bool = False,
        show_index_if_not_sheet: bool = True,
        redraw: bool = True,
    ) -> object:
        self.set_refresh_timer(redraw)
        return self.MT.row_index(
            newindex,
            index,
            reset_row_positions=reset_row_positions,
            show_index_if_not_sheet=show_index_if_not_sheet,
            redraw=False,
        )

    # Bindings and Functionality

    def enable_bindings(self, *bindings: str) -> Sheet:
        self.MT.enable_bindings(bindings)
        return self

    def disable_bindings(self, *bindings: str) -> Sheet:
        self.MT.disable_bindings(bindings)
        return self

    def extra_bindings(
        self,
        bindings: str | list | tuple | None = None,
        func: Callable | None = None,
    ) -> Sheet:
        # bindings is None, unbind all
        if bindings is None:
            bindings = "all"
        # bindings is str, func arg is None or Callable
        if isinstance(bindings, str):
            iterable = [(bindings, func)]
        # bindings is list or tuple of strings, func arg is None or Callable
        elif is_iterable(bindings) and isinstance(bindings[0], str):
            iterable = [(b, func) for b in bindings]
        # bindings is a list or tuple of two tuples or lists
        # in this case the func arg is ignored
        # e.g. [(binding, function), (binding, function), ...]
        elif is_iterable(bindings):
            iterable = bindings

        for b, f in iterable:
            b = b.lower()

            if f is not None and b in emitted_events:
                self.bind(b, f)

            if b in (
                "all",
                "bind_all",
                "unbind_all",
            ):
                self.MT.extra_begin_ctrl_c_func = f
                self.MT.extra_begin_ctrl_x_func = f
                self.MT.extra_begin_ctrl_v_func = f
                self.MT.extra_begin_ctrl_z_func = f
                self.MT.extra_begin_delete_key_func = f
                self.RI.ri_extra_begin_drag_drop_func = f
                self.CH.ch_extra_begin_drag_drop_func = f
                self.MT.extra_begin_del_rows_rc_func = f
                self.MT.extra_begin_del_cols_rc_func = f
                self.MT.extra_begin_insert_cols_rc_func = f
                self.MT.extra_begin_insert_rows_rc_func = f
                self.MT.extra_begin_edit_cell_func = f
                self.CH.extra_begin_edit_cell_func = f
                self.RI.extra_begin_edit_cell_func = f
                self.CH.column_width_resize_func = f
                self.RI.row_height_resize_func = f

            if b in (
                "all",
                "bind_all",
                "unbind_all",
                "all_select_events",
                "select",
                "selectevents",
                "select_events",
            ):
                self.MT.selection_binding_func = f
                self.MT.select_all_binding_func = f
                self.RI.selection_binding_func = f
                self.CH.selection_binding_func = f
                self.MT.drag_selection_binding_func = f
                self.RI.drag_selection_binding_func = f
                self.CH.drag_selection_binding_func = f
                self.MT.shift_selection_binding_func = f
                self.RI.shift_selection_binding_func = f
                self.CH.shift_selection_binding_func = f
                self.MT.ctrl_selection_binding_func = f
                self.RI.ctrl_selection_binding_func = f
                self.CH.ctrl_selection_binding_func = f
                self.MT.deselection_binding_func = f

            if b in (
                "all",
                "bind_all",
                "unbind_all",
                "all_modified_events",
                "sheetmodified",
                "sheet_modified" "modified_events",
                "modified",
            ):
                self.MT.extra_end_ctrl_c_func = f
                self.MT.extra_end_ctrl_x_func = f
                self.MT.extra_end_ctrl_v_func = f
                self.MT.extra_end_ctrl_z_func = f
                self.MT.extra_end_delete_key_func = f
                self.RI.ri_extra_end_drag_drop_func = f
                self.CH.ch_extra_end_drag_drop_func = f
                self.MT.extra_end_del_rows_rc_func = f
                self.MT.extra_end_del_cols_rc_func = f
                self.MT.extra_end_insert_cols_rc_func = f
                self.MT.extra_end_insert_rows_rc_func = f
                self.MT.extra_end_edit_cell_func = f
                self.CH.extra_end_edit_cell_func = f
                self.RI.extra_end_edit_cell_func = f

            if b in (
                "begin_copy",
                "begin_ctrl_c",
            ):
                self.MT.extra_begin_ctrl_c_func = f
            if b in (
                "ctrl_c",
                "end_copy",
                "end_ctrl_c",
                "copy",
            ):
                self.MT.extra_end_ctrl_c_func = f

            if b in (
                "begin_cut",
                "begin_ctrl_x",
            ):
                self.MT.extra_begin_ctrl_x_func = f
            if b in (
                "ctrl_x",
                "end_cut",
                "end_ctrl_x",
                "cut",
            ):
                self.MT.extra_end_ctrl_x_func = f

            if b in (
                "begin_paste",
                "begin_ctrl_v",
            ):
                self.MT.extra_begin_ctrl_v_func = f
            if b in (
                "ctrl_v",
                "end_paste",
                "end_ctrl_v",
                "paste",
            ):
                self.MT.extra_end_ctrl_v_func = f

            if b in (
                "begin_undo",
                "begin_ctrl_z",
            ):
                self.MT.extra_begin_ctrl_z_func = f
            if b in (
                "ctrl_z",
                "end_undo",
                "end_ctrl_z",
                "undo",
            ):
                self.MT.extra_end_ctrl_z_func = f

            if b in (
                "begin_delete_key",
                "begin_delete",
            ):
                self.MT.extra_begin_delete_key_func = f
            if b in (
                "delete_key",
                "end_delete",
                "end_delete_key",
                "delete",
            ):
                self.MT.extra_end_delete_key_func = f

            if b in (
                "begin_edit_cell",
                "begin_edit_table",
            ):
                self.MT.extra_begin_edit_cell_func = f
            if b in (
                "end_edit_cell",
                "edit_cell",
                "edit_table",
            ):
                self.MT.extra_end_edit_cell_func = f

            if b == "begin_edit_header":
                self.CH.extra_begin_edit_cell_func = f
            if b in (
                "end_edit_header",
                "edit_header",
            ):
                self.CH.extra_end_edit_cell_func = f

            if b == "begin_edit_index":
                self.RI.extra_begin_edit_cell_func = f
            if b in (
                "end_edit_index",
                "edit_index",
            ):
                self.RI.extra_end_edit_cell_func = f

            if b in (
                "begin_row_index_drag_drop",
                "begin_move_rows",
            ):
                self.RI.ri_extra_begin_drag_drop_func = f
            if b in (
                "row_index_drag_drop",
                "move_rows",
                "end_move_rows",
                "end_row_index_drag_drop",
            ):
                self.RI.ri_extra_end_drag_drop_func = f

            if b in (
                "begin_column_header_drag_drop",
                "begin_move_columns",
            ):
                self.CH.ch_extra_begin_drag_drop_func = f
            if b in (
                "column_header_drag_drop",
                "move_columns",
                "end_move_columns",
                "end_column_header_drag_drop",
            ):
                self.CH.ch_extra_end_drag_drop_func = f

            if b in (
                "begin_rc_delete_row",
                "begin_delete_rows",
            ):
                self.MT.extra_begin_del_rows_rc_func = f
            if b in (
                "rc_delete_row",
                "end_rc_delete_row",
                "end_delete_rows",
                "delete_rows",
            ):
                self.MT.extra_end_del_rows_rc_func = f

            if b in (
                "begin_rc_delete_column",
                "begin_delete_columns",
            ):
                self.MT.extra_begin_del_cols_rc_func = f
            if b in (
                "rc_delete_column",
                "end_rc_delete_column",
                "end_delete_columns",
                "delete_columns",
            ):
                self.MT.extra_end_del_cols_rc_func = f

            if b in (
                "begin_rc_insert_column",
                "begin_insert_column",
                "begin_insert_columns",
                "begin_add_column",
                "begin_rc_add_column",
                "begin_add_columns",
            ):
                self.MT.extra_begin_insert_cols_rc_func = f
            if b in (
                "rc_insert_column",
                "end_rc_insert_column",
                "end_insert_column",
                "end_insert_columns",
                "rc_add_column",
                "end_rc_add_column",
                "end_add_column",
                "end_add_columns",
            ):
                self.MT.extra_end_insert_cols_rc_func = f

            if b in (
                "begin_rc_insert_row",
                "begin_insert_row",
                "begin_insert_rows",
                "begin_rc_add_row",
                "begin_add_row",
                "begin_add_rows",
            ):
                self.MT.extra_begin_insert_rows_rc_func = f
            if b in (
                "rc_insert_row",
                "end_rc_insert_row",
                "end_insert_row",
                "end_insert_rows",
                "rc_add_row",
                "end_rc_add_row",
                "end_add_row",
                "end_add_rows",
            ):
                self.MT.extra_end_insert_rows_rc_func = f

            if b == "column_width_resize":
                self.CH.column_width_resize_func = f
            if b == "row_height_resize":
                self.RI.row_height_resize_func = f

            if b == "cell_select":
                self.MT.selection_binding_func = f
            if b in (
                "select_all",
                "ctrl_a",
            ):
                self.MT.select_all_binding_func = f
            if b == "row_select":
                self.RI.selection_binding_func = f
            if b in (
                "col_select",
                "column_select",
            ):
                self.CH.selection_binding_func = f
            if b == "drag_select_cells":
                self.MT.drag_selection_binding_func = f
            if b == "drag_select_rows":
                self.RI.drag_selection_binding_func = f
            if b == "drag_select_columns":
                self.CH.drag_selection_binding_func = f
            if b == "shift_cell_select":
                self.MT.shift_selection_binding_func = f
            if b == "shift_row_select":
                self.RI.shift_selection_binding_func = f
            if b == "shift_column_select":
                self.CH.shift_selection_binding_func = f
            if b == "ctrl_cell_select":
                self.MT.ctrl_selection_binding_func = f
            if b == "ctrl_row_select":
                self.RI.ctrl_selection_binding_func = f
            if b == "ctrl_column_select":
                self.CH.ctrl_selection_binding_func = f
            if b == "deselect":
                self.MT.deselection_binding_func = f
        return self

    def bind(
        self,
        binding: str,
        func: Callable,
        add: str | None = None,
    ) -> Sheet:
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
            self.MT.extra_motion_func = func
            self.CH.extra_motion_func = func
            self.RI.extra_motion_func = func
            self.TL.extra_motion_func = func
        elif binding == rc_binding:
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
        elif binding == rc_binding:
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
        self.MT.edit_validation_func = func
        return self

    def popup_menu_add_command(
        self,
        label: str,
        func: Callable,
        table_menu: bool = True,
        index_menu: bool = True,
        header_menu: bool = True,
        empty_space_menu: bool = True,
    ) -> Sheet:
        if label not in self.MT.extra_table_rc_menu_funcs and table_menu:
            self.MT.extra_table_rc_menu_funcs[label] = func
        if label not in self.MT.extra_index_rc_menu_funcs and index_menu:
            self.MT.extra_index_rc_menu_funcs[label] = func
        if label not in self.MT.extra_header_rc_menu_funcs and header_menu:
            self.MT.extra_header_rc_menu_funcs[label] = func
        if label not in self.MT.extra_empty_space_rc_menu_funcs and empty_space_menu:
            self.MT.extra_empty_space_rc_menu_funcs[label] = func
        self.MT.create_rc_menus()
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
        self.MT.create_rc_menus()
        return self

    def basic_bindings(self, enable: bool = False) -> Sheet:
        for canvas in (self.MT, self.CH, self.RI, self.TL):
            canvas.basic_bindings(enable)
        return self

    def cut(self, event: object = None, validation: bool = True) -> None | EventDataDict:
        return self.MT.ctrl_x(event, validation)

    def copy(self, event: object = None) -> None | EventDataDict:
        return self.MT.ctrl_c(event)

    def paste(self, event: object = None, validation: bool = True) -> None | EventDataDict:
        return self.MT.ctrl_v(event, validation)

    def delete(self, event: object = None, validation: bool = True) -> None | EventDataDict:
        return self.MT.delete_key(event, validation)

    def undo(self, event: object = None) -> None | EventDataDict:
        return self.MT.undo(event)

    def redo(self, event: object = None) -> None | EventDataDict:
        return self.MT.redo(event)

    def has_focus(
        self,
    ) -> bool:
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
        convert: object = None,
        undo: bool = False,
        emit_event: bool = False,
        widget: object = None,
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

    def named_span(
        self,
        span: Span,
    ) -> Span:
        if span.name in self.MT.named_spans:
            raise ValueError(f"Span '{span.name}' already exists.")
        if not span.name:
            raise ValueError("Span must have a name.")
        if span.type_ not in named_span_types:
            raise ValueError(f"Span 'type_' must be one of the following: {', '.join(named_span_types)}.")
        self.MT.named_spans[span.name] = span
        self.create_options_from_span(span)
        return span

    def create_options_from_span(self, span: Span) -> Sheet:
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

    def set_named_spans(self, named_spans: None | dict = None) -> Sheet:
        if named_spans is None:
            for name in self.MT.named_spans:
                self.del_named_span(name)
            named_spans = {}
        self.MT.named_spans = named_spans
        return self

    def get_named_span(self, name: str) -> dict:
        return self.MT.named_spans[name]

    def get_named_spans(self) -> dict:
        return self.MT.named_spans

    # Getting Sheet Data

    def __getitem__(
        self,
        *key: CreateSpanTypes,
    ) -> Span:
        return self.span_from_key(*key)

    def span_from_key(
        self,
        *key: CreateSpanTypes,
    ) -> None | Span:
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

    def get_data(
        self,
        *key: CreateSpanTypes,
    ) -> object:
        """
        e.g. retrieves entire table as pandas dataframe
        sheet["A1"].expand().options(pandas.DataFrame).data

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
        tdisp, idisp, hdisp = span.tdisp, span.idisp, span.hdisp
        table, index, header = span.table, span.index, span.header
        fmt_kw = span.kwargs if span.type_ == "format" and span.kwargs else None
        quick_tdata, quick_idata, quick_hdata = self.MT.get_cell_data, self.RI.get_cell_data, self.CH.get_cell_data
        res = []
        if span.transposed:
            if index:
                if index and header:
                    if table:
                        res.append([""] + [quick_idata(r, get_displayed=idisp) for r in rows])
                    else:
                        res.append([quick_idata(r, get_displayed=idisp) for r in rows])
                else:
                    res.append([quick_idata(r, get_displayed=idisp) for r in rows])
            if header:
                if table:
                    res.extend(
                        [quick_hdata(c, get_displayed=hdisp)]
                        + [quick_tdata(r, c, get_displayed=tdisp, fmt_kw=fmt_kw) for r in rows]
                        for c in cols
                    )
                else:
                    res.extend([quick_hdata(c, get_displayed=hdisp)] for c in cols)
            elif table:
                res.extend([quick_tdata(r, c, get_displayed=tdisp, fmt_kw=fmt_kw) for r in rows] for c in cols)
        elif not span.transposed:
            if header:
                if header and index:
                    if table:
                        res.append([""] + [quick_hdata(c, get_displayed=hdisp) for c in cols])
                    else:
                        res.append([quick_hdata(c, get_displayed=hdisp) for c in cols])
                else:
                    res.append([quick_hdata(c, get_displayed=hdisp) for c in cols])
            if index:
                if table:
                    res.extend(
                        [quick_idata(r, get_displayed=idisp)]
                        + [quick_tdata(r, c, get_displayed=tdisp, fmt_kw=fmt_kw) for c in cols]
                        for r in rows
                    )
                else:
                    res.extend([quick_idata(r, get_displayed=idisp)] for r in rows)
            elif table:
                res.extend([quick_tdata(r, c, get_displayed=tdisp, fmt_kw=fmt_kw) for c in cols] for r in rows)
        if not span.ndim:
            # it's a cell
            if len(res) == 1 and len(res[0]) == 1:
                res = res[0][0]
            # it's a single list
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
            if len(res) == 1 and len(res[0]) == 1:
                res = res[0]
            else:
                res = list(chain.from_iterable(res))
        # if span.ndim == 2 res keeps its current
        # dimensions as a list of lists
        if span.convert is not None:
            return span.convert(res)
        return res

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
    ) -> object:
        return self.MT.get_value_for_empty_cell(r, c, r_ops, c_ops)

    @property
    def data(self):
        return self.MT.data

    def __iter__(self) -> Iterator[list[object] | tuple[object]]:
        return self.MT.data.__iter__()

    def __reversed__(self) -> Iterator[list[object] | tuple[object]]:
        return reversed(self.MT.data)

    def __contains__(self, key: object) -> bool:
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
        if selections:
            self.MT.deselect(redraw=False)
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
    ) -> object:
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
    def data(self, value: list[list[object]]) -> None:
        self.data_reference(value)

    def set_data(
        self,
        *key: CreateSpanTypes,
        data: object = None,
        undo: bool | None = None,
        emit_event: bool | None = None,
        redraw: bool = True,
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
        if data is None:
            data = []
        startr, startc = span_froms(span)
        table, index, header = span.table, span.index, span.header
        fmt_kw = span.kwargs if span.type_ == "format" and span.kwargs else None
        transposed = span.transposed
        maxr, maxc = startr, startc
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
            event_data = self.MT.add_columns(
                *self.MT.get_args_for_add_columns(
                    data_ins_col=len(self.MT.col_positions) - 1,
                    displayed_ins_col=len(self.MT.col_positions) - 1,
                    numcols=maxc + 1 - ncols,
                    columns=[],
                    widths=None,
                    headers=False,
                ),
                event_data=event_data,
                create_selections=False,
            )
        if self.MT.all_rows_displayed and maxr >= (nrows := len(self.MT.row_positions) - 1):
            event_data = self.MT.add_rows(
                *self.MT.get_args_for_add_rows(
                    data_ins_row=len(self.MT.row_positions) - 1,
                    displayed_ins_row=len(self.MT.row_positions) - 1,
                    numrows=maxr + 1 - nrows,
                    rows=[],
                    heights=None,
                    row_index=False,
                ),
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
                self.MT.undo_stack.append(pickled_event_dict(event_data))
            if emit_event is True or (emit_event is None and span.emit_event):
                self.emit_event("<<SheetModified>>", event_data)
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
                self.MT.undo_stack.append(pickled_event_dict(event_data))
            if emit_event is True or (emit_event is None and span.emit_event):
                self.emit_event("<<SheetModified>>", event_data)
        self.set_refresh_timer(redraw)
        return event_data

    def event_data_set_table_cell(
        self,
        datarn: int,
        datacn: int,
        value: object,
        event_data: dict,
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
        value: object,
        event_data: dict,
        check_readonly: bool = False,
    ) -> EventDataDict:
        if self.RI.input_valid_for_cell(datarn, value, check_readonly=check_readonly):
            event_data["cells"]["index"][datarn] = self.RI.get_cell_data(datarn)
            self.RI.set_cell_data(datarn, value)
        return event_data

    def event_data_set_header_cell(
        self,
        datacn: int,
        value: object,
        event_data: dict,
        check_readonly: bool = False,
    ) -> EventDataDict:
        if self.CH.input_valid_for_cell(datacn, value, check_readonly=check_readonly):
            event_data["cells"]["header"][datacn] = self.CH.get_cell_data(datacn)
            self.CH.set_cell_data(datacn, value)
        return event_data

    def insert_row(
        self,
        row: list[object] | tuple[object] | None = None,
        idx: str | int | None = None,
        height: int | None = None,
        row_index: bool = False,
        fill: bool = True,
        undo: bool = False,
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
        column: list[object] | tuple[object] | None = None,
        idx: str | int | None = None,
        width: int | None = None,
        header: bool = False,
        fill: bool = True,
        undo: bool = False,
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
        rows: list[tuple[object] | list[object]] | tuple[tuple[object] | list[object]] | int = 1,
        idx: str | int | None = None,
        heights: list[int] | tuple[int] | None = None,
        row_index: bool = False,
        fill: bool = True,
        undo: bool = False,
        emit_event: bool = False,
        create_selections: bool = True,
        add_column_widths: bool = True,
        push_ops: bool = True,
        redraw: bool = True,
    ) -> EventDataDict:
        total_cols = None
        if (idx := idx_param_to_int(idx)) is None:
            idx = len(self.MT.data)
        if isinstance(rows, int):
            if rows < 1:
                raise ValueError(f"rows arg must be greater than 0, not {rows}")
            total_cols = self.MT.total_data_cols()
            if row_index:
                data = [
                    [self.RI.get_value_for_empty_cell(idx + i, r_ops=False)]
                    + self.MT.get_empty_row_seq(
                        idx + i,
                        total_cols,
                        r_ops=False,
                        c_ops=False,
                    )
                    for i in range(rows)
                ]
            else:
                data = [
                    self.MT.get_empty_row_seq(
                        idx + i,
                        total_cols,
                        r_ops=False,
                        c_ops=False,
                    )
                    for i in range(rows)
                ]
        else:
            data = rows
        if not isinstance(rows, int) and fill:
            total_cols = self.MT.total_data_cols() if total_cols is None else total_cols
            len_check = (total_cols + 1) if row_index else total_cols
            for rn, r in enumerate(data):
                if len_check > (lnr := len(r)):
                    r += self.MT.get_empty_row_seq(
                        rn,
                        end=total_cols,
                        start=(lnr - 1) if row_index else lnr,
                        r_ops=False,
                        c_ops=False,
                    )
        numrows = len(data)
        if self.MT.all_rows_displayed:
            displayed_ins_idx = idx
        else:
            displayed_ins_idx = bisect_left(self.MT.displayed_rows, idx)
        event_data = self.MT.add_rows(
            *self.MT.get_args_for_add_rows(
                data_ins_row=idx,
                displayed_ins_row=displayed_ins_idx,
                numrows=numrows,
                rows=data,
                heights=heights,
                row_index=row_index,
            ),
            add_col_positions=add_column_widths,
            event_data=event_dict(
                name="add_rows",
                sheet=self.name,
                boxes=self.MT.get_boxes(),
                selected=self.MT.selected,
            ),
            create_selections=create_selections,
            push_ops=push_ops,
        )
        if undo:
            self.MT.undo_stack.append(pickled_event_dict(event_data))
        if emit_event:
            self.emit_event("<<SheetModified>>", event_data)
        self.set_refresh_timer(redraw)
        return event_data

    def insert_columns(
        self,
        columns: list[tuple[object] | list[object]] | tuple[tuple[object] | list[object]] | int = 1,
        idx: str | int | None = None,
        widths: list[int] | tuple[int] | None = None,
        headers: bool = False,
        fill: bool = True,
        undo: bool = False,
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
            if headers:
                data = [
                    [self.CH.get_value_for_empty_cell(datacn, c_ops=False)]
                    + [
                        self.MT.get_value_for_empty_cell(
                            datarn,
                            datacn,
                            r_ops=False,
                            c_ops=False,
                        )
                        for datarn in range(total_rows)
                    ]
                    for datacn in range(idx, idx + columns)
                ]
            else:
                data = [
                    [
                        self.MT.get_value_for_empty_cell(
                            datarn,
                            datacn,
                            r_ops=False,
                            c_ops=False,
                        )
                        for datarn in range(total_rows)
                    ]
                    for datacn in range(idx, idx + columns)
                ]
            numcols = columns
        else:
            data = columns
            numcols = len(columns)
            if fill:
                len_check = (total_rows + 1) if headers else total_rows
                for i, column in enumerate(data):
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
        if self.MT.all_columns_displayed:
            displayed_ins_idx = idx
        elif not self.MT.all_columns_displayed:
            displayed_ins_idx = bisect_left(self.MT.displayed_columns, idx)
        event_data = self.MT.add_columns(
            *self.MT.get_args_for_add_columns(
                data_ins_col=idx,
                displayed_ins_col=displayed_ins_idx,
                numcols=numcols,
                columns=data,
                widths=widths,
                headers=headers,
            ),
            add_row_positions=add_row_heights,
            event_data=event_dict(
                name="add_columns",
                sheet=self.name,
                boxes=self.MT.get_boxes(),
                selected=self.MT.selected,
            ),
            create_selections=create_selections,
            push_ops=push_ops,
        )
        if undo:
            self.MT.undo_stack.append(pickled_event_dict(event_data))
        if emit_event:
            self.emit_event("<<SheetModified>>", event_data)
        self.set_refresh_timer(redraw)
        return event_data

    def del_row(
        self,
        idx: int = 0,
        data_indexes: bool = True,
        undo: bool = False,
        emit_event: bool = False,
        redraw: bool = True,
    ) -> EventDataDict:
        return self.del_rows(
            rows=idx,
            data_indexes=data_indexes,
            undo=undo,
            emit_event=emit_event,
            redraw=redraw,
        )

    delete_row = del_row

    def del_column(
        self,
        idx: int = 0,
        data_indexes: bool = True,
        undo: bool = False,
        emit_event: bool = False,
        redraw: bool = True,
    ) -> EventDataDict:
        return self.del_columns(
            columns=idx,
            data_indexes=data_indexes,
            undo=undo,
            emit_event=emit_event,
            redraw=redraw,
        )

    delete_column = del_column

    def del_rows(
        self,
        rows: int | Iterator[int],
        data_indexes: bool = True,
        undo: bool = False,
        emit_event: bool = False,
        redraw: bool = True,
    ) -> EventDataDict:
        rows = [rows] if isinstance(rows, int) else sorted(rows)
        event_data = event_dict(
            name="delete_rows",
            sheet=self.name,
            widget=self,
            boxes=self.MT.get_boxes(),
            selected=self.MT.selected,
        )
        if not data_indexes:
            event_data = self.MT.delete_rows_displayed(rows, event_data)
            event_data = self.MT.delete_rows_data(
                rows if self.MT.all_rows_displayed else [self.MT.displayed_rows[r] for r in rows],
                event_data,
            )
        else:
            if self.MT.all_rows_displayed:
                rows = rows
            else:
                rows = data_to_displayed_idxs(rows, self.MT.displayed_rows)
            event_data = self.MT.delete_rows_data(rows, event_data)
            event_data = self.MT.delete_rows_displayed(
                rows,
                event_data,
            )
        if undo:
            self.MT.undo_stack.append(pickled_event_dict(event_data))
        if emit_event:
            self.emit_event("<<SheetModified>>", event_data)
        self.MT.deselect("all", redraw=False)
        self.set_refresh_timer(redraw)
        return event_data

    delete_rows = del_rows

    def del_columns(
        self,
        columns: int | Iterator[int],
        data_indexes: bool = True,
        undo: bool = False,
        emit_event: bool = False,
        redraw: bool = True,
    ) -> EventDataDict:
        columns = [columns] if isinstance(columns, int) else sorted(columns)
        event_data = event_dict(
            name="delete_columns",
            sheet=self.name,
            widget=self,
            boxes=self.MT.get_boxes(),
            selected=self.MT.selected,
        )
        if not data_indexes:
            event_data = self.MT.delete_columns_displayed(columns, event_data)
            event_data = self.MT.delete_columns_data(
                columns if self.MT.all_columns_displayed else [self.MT.displayed_columns[c] for c in columns],
                event_data,
            )
        else:
            if self.MT.all_columns_displayed:
                columns = columns
            else:
                columns = data_to_displayed_idxs(columns, self.MT.displayed_columns)
            event_data = self.MT.delete_columns_data(columns, event_data)
            event_data = self.MT.delete_columns_displayed(
                columns,
                event_data,
            )
        if undo:
            self.MT.undo_stack.append(pickled_event_dict(event_data))
        if emit_event:
            self.emit_event("<<SheetModified>>", event_data)
        self.MT.deselect("all", redraw=False)
        self.set_refresh_timer(redraw)
        return event_data

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

    def move_row(self, row: int, moveto: int) -> tuple[dict, dict, dict]:
        return self.move_rows(moveto, row)

    def move_column(self, column: int, moveto: int) -> tuple[dict, dict, dict]:
        return self.move_columns(moveto, column)

    def move_rows(
        self,
        move_to: int | None = None,
        to_move: list[int] | None = None,
        move_data: bool = True,
        data_indexes: bool = False,
        create_selections: bool = True,
        undo: bool = False,
        emit_event: bool = False,
        move_heights: bool = True,
        redraw: bool = True,
    ) -> tuple[dict, dict, dict]:
        data_idxs, disp_idxs, event_data = self.MT.move_rows_adjust_options_dict(
            *self.MT.get_args_for_move_rows(
                move_to=move_to,
                to_move=to_move,
                data_indexes=data_indexes,
            ),
            move_data=move_data,
            move_heights=move_heights,
            create_selections=create_selections,
            data_indexes=data_indexes,
        )
        if undo:
            self.MT.undo_stack.append(pickled_event_dict(event_data))
        if emit_event:
            self.emit_event("<<SheetModified>>", event_data)
        self.set_refresh_timer(redraw)
        return data_idxs, disp_idxs, event_data

    def move_columns(
        self,
        move_to: int | None = None,
        to_move: list[int] | None = None,
        move_data: bool = True,
        data_indexes: bool = False,
        create_selections: bool = True,
        undo: bool = False,
        emit_event: bool = False,
        move_widths: bool = True,
        redraw: bool = True,
    ) -> tuple[dict, dict, dict]:
        data_idxs, disp_idxs, event_data = self.MT.move_columns_adjust_options_dict(
            *self.MT.get_args_for_move_columns(
                move_to=move_to,
                to_move=to_move,
                data_indexes=data_indexes,
            ),
            move_data=move_data,
            move_widths=move_widths,
            create_selections=create_selections,
            data_indexes=data_indexes,
        )
        if undo:
            self.MT.undo_stack.append(pickled_event_dict(event_data))
        if emit_event:
            self.emit_event("<<SheetModified>>", event_data)
        self.set_refresh_timer(redraw)
        return data_idxs, disp_idxs, event_data

    def mapping_move_columns(
        self,
        data_new_idxs: dict[int, int],
        disp_new_idxs: None | dict[int, int] = None,
        move_data: bool = True,
        data_indexes: bool = False,
        create_selections: bool = True,
        undo: bool = False,
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
            data_indexes=data_indexes,
        )
        if undo:
            self.MT.undo_stack.append(pickled_event_dict(event_data))
        if emit_event:
            self.emit_event("<<SheetModified>>", event_data)
        self.set_refresh_timer(redraw)
        return data_idxs, disp_idxs, event_data

    def mapping_move_rows(
        self,
        data_new_idxs: dict[int, int],
        disp_new_idxs: None | dict[int, int] = None,
        move_data: bool = True,
        data_indexes: bool = False,
        create_selections: bool = True,
        undo: bool = False,
        emit_event: bool = False,
        redraw: bool = True,
    ) -> tuple[dict[int, int], dict[int, int], EventDataDict]:
        data_idxs, disp_idxs, event_data = self.MT.move_rows_adjust_options_dict(
            data_new_idxs=data_new_idxs,
            data_old_idxs=dict(zip(data_new_idxs.values(), data_new_idxs)),
            totalrows=None,
            disp_new_idxs=disp_new_idxs,
            move_data=move_data,
            create_selections=create_selections,
            data_indexes=data_indexes,
        )
        if undo:
            self.MT.undo_stack.append(pickled_event_dict(event_data))
        if emit_event:
            self.emit_event("<<SheetModified>>", event_data)
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
        values: list = [],
        edit_data: bool = True,
        set_values: dict[tuple[int, int], object] = {},
        set_value: object = None,
        state: str = "normal",
        redraw: bool = True,
        selection_function: Callable | None = None,
        modified_function: Callable | None = None,
        search_function: Callable = dropdown_search_function,
        validate_input: bool = True,
        text: None | str = None,
    ) -> Span:
        v = set_value if set_value is not None else values[0] if values else ""
        kwargs = {
            "values": values,
            "state": state,
            "selection_function": selection_function,
            "modified_function": modified_function,
            "search_function": search_function,
            "validate_input": validate_input,
            "text": text,
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
                    set_idata(r, value=set_values[r] if r in set_values else v)
        if header:
            for c in cols:
                self.del_header_cell_options_dropdown_and_checkbox(c)
                add_to_options(self.CH.cell_options, c, "dropdown", d)
                if edit_data:
                    set_hdata(c, value=set_values[c] if c in set_values else v)
        if table:
            if span.kind == "cell":
                for r in rows:
                    for c in cols:
                        self.del_cell_options_dropdown_and_checkbox(r, c)
                        add_to_options(self.MT.cell_options, (r, c), "dropdown", d)
                        if edit_data:
                            set_tdata(r, c, value=set_values[(r, c)] if (r, c) in set_values else v)
            elif span.kind == "row":
                for r in rows:
                    self.del_row_options_dropdown_and_checkbox(r)
                    add_to_options(self.MT.row_options, r, "dropdown", d)
                    if edit_data:
                        for c in cols:
                            set_tdata(r, c, value=set_values[(r, c)] if (r, c) in set_values else v)
            elif span.kind == "column":
                for c in cols:
                    self.del_column_options_dropdown_and_checkbox(c)
                    add_to_options(self.MT.col_options, c, "dropdown", d)
                    if edit_data:
                        for r in rows:
                            set_tdata(r, c, value=set_values[(r, c)] if (r, c) in set_values else v)
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
        state: str = "normal",
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
                    set_idata(r, checked if isinstance(checked, bool) else force_bool(self.get_index_data(r)))
        if header:
            for c in cols:
                self.del_header_cell_options_dropdown_and_checkbox(c)
                add_to_options(self.CH.cell_options, c, "checkbox", d)
                if edit_data:
                    set_hdata(c, checked if isinstance(checked, bool) else force_bool(self.get_header_data(c)))
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
        formatter_options: dict = {},
        formatter_class: object = None,
        redraw: bool = True,
        **kwargs,
    ) -> Span:
        span = self.span_from_key(*key)
        rows, cols = self.ranges_from_span(span)
        kwargs = fix_format_kwargs({"formatter": formatter_class, **formatter_options, **kwargs})
        if span.kind == "cell" and span.table:
            for r in rows:
                for c in cols:
                    self.del_cell_options_checkbox(r, c)
                    add_to_options(self.MT.cell_options, (r, c), "format", kwargs)
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
                set_readonly(self.MT.row_options, r, readonly)
        elif span.kind == "column":
            for c in cols:
                set_readonly(self.MT.col_options, c, readonly)
        return span

    # Text Font and Alignment

    def font(
        self,
        newfont: tuple[str, int, str] | None = None,
        reset_row_positions: bool = True,
    ) -> tuple[str, int, str]:
        return self.MT.set_table_font(newfont, reset_row_positions=reset_row_positions)

    def header_font(self, newfont: tuple[str, int, str] | None = None) -> tuple[str, int, str]:
        return self.MT.set_header_font(newfont)

    def table_align(
        self,
        align: str = None,
        redraw: bool = True,
    ) -> str | Sheet:
        if align is None:
            return self.MT.align
        elif convert_align(align):
            self.MT.align = convert_align(align)
        else:
            raise ValueError("Align must be one of the following values: c, center, w, west, e, east")
        return self.set_refresh_timer(redraw)

    def header_align(
        self,
        align: str = None,
        redraw: bool = True,
    ) -> str | Sheet:
        if align is None:
            return self.CH.align
        elif convert_align(align):
            self.CH.align = convert_align(align)
        else:
            raise ValueError("Align must be one of the following values: c, center, w, west, e, east")
        return self.set_refresh_timer(redraw)

    def row_index_align(
        self,
        align: str = None,
        redraw: bool = True,
    ) -> str | Sheet:
        if align is None:
            return self.RI.align
        elif convert_align(align):
            self.RI.align = convert_align(align)
        else:
            raise ValueError("Align must be one of the following values: c, center, w, west, e, east")
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
    ) -> list[tuple[int, int]] | set[tuple[int, int]]:
        if sort_by_row and sort_by_column:
            sels = sorted(
                self.MT.get_selected_cells(get_rows=get_rows, get_cols=get_columns),
                key=lambda t: t[1],
            )
            return sorted(sels, key=lambda t: t[0])
        elif sort_by_row:
            return sorted(
                self.MT.get_selected_cells(get_rows=get_rows, get_cols=get_columns),
                key=lambda t: t[0],
            )
        elif sort_by_column:
            return sorted(
                self.MT.get_selected_cells(get_rows=get_rows, get_cols=get_columns),
                key=lambda t: t[1],
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
        if self.MT.anything_selected(
            exclude_columns=exclude_columns,
            exclude_rows=exclude_rows,
            exclude_cells=exclude_cells,
        ):
            return True
        return False

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
        if column == "all" and width == "default":
            self.MT.reset_col_positions()
        elif column == "displayed" and width == "text":
            for c in range(*self.MT.visible_text_columns):
                self.CH.set_col_width(c)
        elif width == "text" and isinstance(column, int):
            self.CH.set_col_width(col=column, width=None, only_if_too_small=only_set_if_too_small)
        elif isinstance(width, int) and isinstance(column, int):
            self.CH.set_col_width(col=column, width=width, only_if_too_small=only_set_if_too_small)
        elif isinstance(column, int):
            return int(self.MT.col_positions[column + 1] - self.MT.col_positions[column])
        return self.set_refresh_timer(redraw)

    def row_height(
        self,
        row: int | Literal["all", "displayed"] | None = None,
        height: int | Literal["default", "text"] | None = None,
        only_set_if_too_small: bool = False,
        redraw: bool = True,
    ) -> Sheet | int:
        if row == "all" and height == "default":
            self.MT.reset_row_positions()
        elif row == "displayed" and height == "text":
            for r in range(*self.MT.visible_text_rows):
                self.RI.set_row_height(r)
        elif height == "text" and isinstance(row, int):
            self.RI.set_row_height(row=row, height=None, only_if_too_small=only_set_if_too_small)
        elif isinstance(height, int) and isinstance(row, int):
            self.RI.set_row_height(row=row, height=height, only_if_too_small=only_set_if_too_small)
        elif isinstance(row, int):
            return int(self.MT.row_positions[row + 1] - self.MT.row_positions[row])
        return self.set_refresh_timer(redraw)

    def get_column_widths(self, canvas_positions: bool = False) -> list[float]:
        if canvas_positions:
            return self.MT.col_positions
        return self.MT.get_column_widths()

    def get_row_heights(self, canvas_positions: bool = False) -> list[float]:
        if canvas_positions:
            return self.MT.row_positions
        return self.MT.get_row_heights()

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
        self.MT.insert_col_position(idx=idx, width=width, deselect_all=deselect_all)
        return self.set_refresh_timer(redraw)

    def insert_row_position(
        self,
        idx: Literal["end"] | int = "end",
        height: int | None = None,
        deselect_all: bool = False,
        redraw: bool = False,
    ) -> Sheet:
        self.MT.insert_row_position(idx=idx, height=height, deselect_all=deselect_all)
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
                x - z < self.MT.min_column_width or not isinstance(x, int) or isinstance(x, bool)
                for z, x in zip(column_widths, islice(column_widths, 1, None))
            )
        return not any(
            z < self.MT.min_column_width or not isinstance(z, int) or isinstance(z, bool) for z in column_widths
        )

    def valid_row_height(self, height: int) -> int:
        if height < self.MT.min_row_height:
            return self.MT.min_row_height
        elif height > self.MT.max_row_height:
            return self.MT.max_row_height
        return height

    def valid_column_width(self, width: int) -> int:
        if width < self.MT.min_column_width:
            return self.MT.min_column_width
        elif width > self.MT.max_column_width:
            return self.MT.max_column_width
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

    def identify_region(self, event: object) -> Literal["table", "index", "header", "top left"]:
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
        event: object,
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
        event: object,
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

    def sync_scroll(self, widget: object) -> Sheet:
        if widget is self:
            return self
        self.MT.synced_scrolls.add(widget)
        if isinstance(widget, Sheet):
            widget.MT.synced_scrolls.add(self)
        return self

    def unsync_scroll(self, widget: object = None) -> Sheet:
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
        bottom_right_corner: bool = False,
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
        return self.MT.cell_completely_visible(r, c, seperate_axes)

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
    dcol = displayed_column_to_data

    def display_columns(
        self,
        columns: None | Literal["all"] | Iterator[int] = None,
        all_columns_displayed: None | bool = None,
        reset_col_positions: bool = True,
        refresh: bool = False,
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
        if refresh or redraw:
            self.set_refresh_timer(redraw if redraw else refresh)
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
            self.MT.displayed_columns = [c for c in range(self.MT.total_data_cols()) if c not in columns]
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

    # uses data indexes
    def show_columns(
        self,
        columns: int | Iterator[int],
        redraw: bool = True,
        deselect_all: bool = True,
    ) -> Sheet:
        if isinstance(columns, int):
            columns = [columns]
        cws = self.MT.get_column_widths()
        for column in columns:
            idx = bisect_left(self.MT.displayed_columns, column)
            if len(self.MT.displayed_columns) == idx or self.MT.displayed_columns[idx] != column:
                self.MT.displayed_columns.insert(idx, column)
                cws.insert(idx, self.MT.saved_column_widths.pop(column, self.PAR.ops.default_column_width))
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
        self.MT.all_columns_displayed = a
        return self

    @property
    def displayed_columns(self) -> list[int]:
        return self.MT.displayed_columns

    @displayed_columns.setter
    def displayed_columns(self, columns: list[int]) -> Sheet:
        self.display_columns(columns=columns, reset_col_positions=False, redraw=True)
        return self

    # Hiding Rows

    def displayed_row_to_data(self, r: int) -> int:
        return r if self.MT.all_rows_displayed else self.MT.displayed_rows[r]

    data_r = displayed_row_to_data
    drow = displayed_row_to_data

    def display_rows(
        self,
        rows: None | Literal["all"] | Iterator[int] = None,
        all_rows_displayed: None | bool = None,
        reset_row_positions: bool = True,
        refresh: bool = False,
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
        if refresh or redraw:
            self.set_refresh_timer(redraw if redraw else refresh)
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
            self.MT.displayed_rows = [r for r in range(self.MT.total_data_rows()) if r not in rows]
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

    # uses data indexes
    def show_rows(
        self,
        rows: int | Iterator[int],
        redraw: bool = True,
        deselect_all: bool = True,
    ) -> Sheet:
        if isinstance(rows, int):
            rows = [rows]
        rhs = self.MT.get_row_heights()
        for row in rows:
            idx = bisect_left(self.MT.displayed_rows, row)
            if len(self.MT.displayed_rows) == idx or self.MT.displayed_rows[idx] != row:
                self.MT.displayed_rows.insert(idx, row)
                rhs.insert(idx, self.MT.saved_row_heights.pop(row, self.MT.get_default_row_height()))
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
        self.MT.all_rows_displayed = a
        return self

    @property
    def displayed_rows(self) -> list[int]:
        return self.MT.displayed_rows

    @displayed_rows.setter
    def displayed_rows(self, rows: list[int]) -> Sheet:
        self.display_rows(rows=rows, reset_row_positions=False, redraw=True)
        return self

    # Hiding Sheet Elements

    def hide(
        self,
        canvas: Literal[
            "all",
            "row_index",
            "header",
            "top_left",
            "x_scrollbar",
            "y_scrollbar",
        ] = "all",
    ) -> Sheet:
        if canvas.lower() == "all":
            self.TL.grid_forget()
            self.RI.grid_forget()
            self.RI["yscrollcommand"] = 0
            self.MT.show_index = False
            self.CH.grid_forget()
            self.CH["xscrollcommand"] = 0
            self.MT.show_header = False
            self.MT.grid_forget()
            self.yscroll.grid_forget()
            self.xscroll.grid_forget()
            self.xscroll_showing = False
            self.yscroll_showing = False
            self.xscroll_disabled = True
            self.yscroll_disabled = True
        elif canvas.lower() == "row_index":
            self.RI.grid_forget()
            self.RI["yscrollcommand"] = 0
            self.MT.show_index = False
        elif canvas.lower() == "header":
            self.CH.grid_forget()
            self.CH["xscrollcommand"] = 0
            self.MT.show_header = False
        elif canvas.lower() == "top_left":
            self.TL.grid_forget()
        elif canvas.lower() == "x_scrollbar":
            self.xscroll.grid_forget()
            self.xscroll_showing = False
            self.xscroll_disabled = True
        elif canvas.lower() == "y_scrollbar":
            self.yscroll.grid_forget()
            self.yscroll_showing = False
            self.yscroll_disabled = True
        return self

    def show(
        self,
        canvas: Literal[
            "all",
            "row_index",
            "header",
            "top_left",
            "x_scrollbar",
            "y_scrollbar",
        ] = "all",
    ) -> Sheet:
        if canvas == "all":
            self.hide()
            self.TL.grid(row=0, column=0)
            self.RI.grid(row=1, column=0, sticky="nswe")
            self.CH.grid(row=0, column=1, sticky="nswe")
            self.MT.grid(row=1, column=1, sticky="nswe")
            self.yscroll.grid(row=0, column=2, rowspan=3, sticky="nswe")
            self.xscroll.grid(row=2, column=0, columnspan=2, sticky="nswe")
            self.MT["xscrollcommand"] = self.xscroll.set
            self.CH["xscrollcommand"] = self.xscroll.set
            self.MT["yscrollcommand"] = self.yscroll.set
            self.RI["yscrollcommand"] = self.yscroll.set
            self.xscroll_showing = True
            self.yscroll_showing = True
            self.xscroll_disabled = False
            self.yscroll_disabled = False
        elif canvas == "row_index":
            self.RI.grid(row=1, column=0, sticky="nswe")
            self.MT["yscrollcommand"] = self.yscroll.set
            self.RI["yscrollcommand"] = self.yscroll.set
            self.MT.show_index = True
        elif canvas == "header":
            self.CH.grid(row=0, column=1, sticky="nswe")
            self.MT["xscrollcommand"] = self.xscroll.set
            self.CH["xscrollcommand"] = self.xscroll.set
            self.MT.show_header = True
        elif canvas == "top_left":
            self.TL.grid(row=0, column=0)
        elif canvas == "x_scrollbar":
            self.xscroll.grid(row=2, column=0, columnspan=2, sticky="nswe")
            self.xscroll_showing = True
            self.xscroll_disabled = False
        elif canvas == "y_scrollbar":
            self.yscroll.grid(row=0, column=2, rowspan=3, sticky="nswe")
            self.yscroll_showing = True
            self.yscroll_disabled = False
        self.MT.update_idletasks()
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

    def destroy_text_editor(self, event: object = None) -> Sheet:
        self.MT.hide_text_editor(reason=event)
        self.RI.hide_text_editor(reason=event)
        self.CH.hide_text_editor(reason=event)
        return self

    def get_text_editor_widget(self, event: object = None) -> tk.Text | None:
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
                try:
                    self.MT.text_editor.tktext.unbind(key)
                except Exception:
                    pass
            self.MT.text_editor_user_bound_keys = {}
        else:
            if key in self.MT.text_editor_user_bound_keys:
                del self.MT.text_editor_user_bound_keys[key]
            try:
                self.MT.text_editor.tktext.unbind(key)
            except Exception:
                pass
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
        for k, v in kwargs.items():
            if k in self.ops and v != self.ops[k]:
                if k.endswith("bindings"):
                    self.MT._disable_binding(k.split("_")[0])
                self.ops[k] = v
                if k.endswith("bindings"):
                    self.MT._enable_binding(k.split("_")[0])
        if "from_clipboard_delimiters" in kwargs:
            self.ops.from_clipboard_delimiters = (
                self.ops.from_clipboard_delimiters
                if isinstance(self.ops.from_clipboard_delimiters, str)
                else "".join(self.ops.from_clipboard_delimiters)
            )
        if "default_row_height" in kwargs:
            self.default_row_height(kwargs["default_row_height"])
        if "default_header" in kwargs:
            self.CH.default_header = kwargs["default_header"].lower()
        if "default_row_index" in kwargs:
            self.RI.default_index = kwargs["default_row_index"].lower()
        if "max_column_width" in kwargs:
            self.MT.max_column_width = float(kwargs["max_column_width"])
        if "max_row_height" in kwargs:
            self.MT.max_row_height = float(kwargs["max_row_height"])
        if "max_header_height" in kwargs:
            self.MT.max_header_height = float(kwargs["max_header_height"])
        if "max_index_width" in kwargs:
            self.MT.max_index_width = float(kwargs["max_index_width"])
        if "expand_sheet_if_paste_too_big" in kwargs:
            self.ops.paste_can_expand_x = kwargs["expand_sheet_if_paste_too_big"]
            self.ops.paste_can_expand_y = kwargs["expand_sheet_if_paste_too_big"]
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
        self.MT.create_rc_menus()
        if "treeview" in kwargs:
            self.set_options(auto_resize_row_index=True, redraw=False)
            self.index_align("w", redraw=False)
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
        event: object,
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

    def get_cell_options(
        self,
        key: None | str = None,
        canvas: Literal["table", "row_index", "header"] = "table",
    ) -> dict:
        if canvas == "table":
            target = self.MT.cell_options
        elif canvas == "row_index":
            target = self.RI.cell_options
        elif canvas == "header":
            target = self.CH.cell_options
        if key is None:
            return target
        return {k: v[key] for k, v in target.items() if key in v}

    def get_row_options(self, key: None | str = None) -> dict:
        if key is None:
            return self.MT.row_options
        return {k: v[key] for k, v in self.MT.row_options.items() if key in v}

    def get_column_options(self, key: None | str = None) -> dict:
        if key is None:
            return self.MT.col_options
        return {k: v[key] for k, v in self.MT.col_options.items() if key in v}

    def get_index_options(self, key: None | str = None) -> dict:
        if key is None:
            return self.RI.cell_options
        return {k: v[key] for k, v in self.RI.cell_options.items() if key in v}

    def get_header_options(self, key: None | str = None) -> dict:
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
            res.cells.update(self.MT.tagged_cells[tag] if tag in self.MT.tagged_cells else set())
            res.rows.update(self.MT.tagged_rows[tag] if tag in self.MT.tagged_rows else set())
            res.columns.update(self.MT.tagged_columns[tag] if tag in self.MT.tagged_columns else set())
        return res

    # Treeview Mode

    def tree_build(
        self,
        data: list[list[object]],
        iid_column: int,
        parent_column: int,
        text_column: None | int = None,
        push_ops: bool = False,
        row_heights: Sequence[int] | None | False = None,
        open_ids: Iterator[str] | None = None,
        safety: bool = True,
        ncols: int | None = None,
    ) -> Sheet:
        self.reset(cell_options=False, column_widths=False, header=False, redraw=False)
        if text_column is None:
            text_column = iid_column
        tally_of_ids = defaultdict(lambda: -1)
        if not isinstance(ncols, int):
            ncols = max(map(len, data), default=0)
        for rn, row in enumerate(data):
            if safety and ncols > (lnr := len(row)):
                row += self.MT.get_empty_row_seq(rn, end=ncols, start=lnr)
            iid, pid = row[iid_column].lower(), row[parent_column].lower()
            if safety:
                if not iid:
                    continue
                tally_of_ids[iid] += 1
                if tally_of_ids[iid] > 0:
                    x = 1
                    while iid in tally_of_ids:
                        new = f"{row[iid_column]}_DUPLICATED_{x}"
                        iid = new.lower()
                        x += 1
                    tally_of_ids[iid] += 1
                    row[iid_column] = new
            if iid in self.RI.tree:
                self.RI.tree[iid].text = row[text_column]
            else:
                self.RI.tree[iid] = Node(row[text_column], iid, "")
            if safety and (iid == pid or self.RI.pid_causes_recursive_loop(iid, pid)):
                row[parent_column] = ""
                pid = ""
            if pid:
                if pid not in self.RI.tree:
                    self.RI.tree[pid] = Node(row[text_column], pid)
                self.RI.tree[iid].parent = self.RI.tree[pid]
                self.RI.tree[pid].children.append(self.RI.tree[iid])
            else:
                self.RI.tree[iid].parent = ""
            self.RI.tree_rns[iid] = rn
        for n in self.RI.tree.values():
            if n.parent is None:
                n.parent = ""
                newrow = self.MT.get_empty_row_seq(len(data), ncols)
                newrow[iid_column] = n.iid
                self.RI.tree_rns[n.iid] = len(data)
                data.append(newrow)
        self.insert_rows(
            rows=[[self.RI.tree[iid]] + data[self.RI.tree_rns[iid]] for iid in self.get_children()],
            idx=0,
            heights={} if row_heights is False else row_heights,
            row_index=True,
            create_selections=False,
            fill=False,
            push_ops=push_ops,
            redraw=False,
        )
        self.MT.all_rows_displayed = False
        self.MT.displayed_rows = list(range(len(self.MT._row_index)))
        self.RI.tree_rns = {n.iid: i for i, n in enumerate(self.MT._row_index)}
        if open_ids:
            self.tree_set_open(open_ids=open_ids)
        else:
            self.hide_rows(
                {self.RI.tree_rns[iid] for iid in self.get_children() if self.RI.tree[iid].parent},
                deselect_all=False,
                data_indexes=True,
                row_heights=False if row_heights is False else True,
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
            set(rn for rn in self.MT.displayed_rows if self.MT._row_index[rn].parent),
            redraw=False,
            deselect_all=False,
            data_indexes=True,
        )
        open_ids = set(filter(self.exists, map(str.lower, open_ids)))
        self.RI.tree_open_ids = set()
        if open_ids:
            self.show_rows(
                rows=self._tree_open(open_ids),
                redraw=False,
                deselect_all=False,
            )
        return self.set_refresh_timer()

    def _tree_open(self, items: set[str]) -> list[int]:
        """
        Only meant for internal use
        """
        to_open = []
        disp_set = set(self.MT.displayed_rows)
        quick_rns = self.RI.tree_rns
        quick_open_ids = self.RI.tree_open_ids
        for item in filter(items.__contains__, self.get_children()):
            if self.RI.tree[item].children:
                quick_open_ids.add(item)
                if quick_rns[item] in disp_set:
                    to_disp = [quick_rns[did] for did in self.RI.get_iid_descendants(item, check_open=True)]
                    for i in to_disp:
                        disp_set.add(i)
                    to_open.extend(to_disp)
        return to_open

    def tree_open(self, *items, redraw: bool = True) -> Sheet:
        """
        If used without args all items are opened
        """
        if items := set(unpack(items)):
            to_open = self._tree_open(items)
        else:
            to_open = self._tree_open(set(self.get_children()))
        return self.show_rows(
            rows=to_open,
            redraw=redraw,
            deselect_all=False,
        )

    def _tree_close(self, items: Iterator[str]) -> list[int]:
        """
        Only meant for internal use
        """
        to_close = set()
        disp_set = set(self.MT.displayed_rows)
        quick_rns = self.RI.tree_rns
        quick_open_ids = self.RI.tree_open_ids
        for item in items:
            if self.RI.tree[item].children:
                quick_open_ids.discard(item)
                if quick_rns[item] in disp_set:
                    for did in self.RI.get_iid_descendants(item, check_open=True):
                        to_close.add(quick_rns[did])
        return to_close

    def tree_close(self, *items, redraw: bool = True) -> Sheet:
        """
        If used without args all items are closed
        """
        if items:
            to_close = self._tree_close(unpack(items))
        else:
            to_close = self._tree_close(self.get_children())
        return self.hide_rows(
            rows=to_close,
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
        values: None | list = None,
        create_selections: bool = False,
    ) -> str:
        if iid is None:
            i = 0
            while (iid := f"{i}") in self.RI.tree:
                i += 1
        iid, pid = iid.lower(), parent.lower()
        if not iid:
            raise ValueError("iid cannot be empty string.")
        if iid in self.RI.tree:
            raise ValueError(f"iid '{iid}' already exists.")
        if iid == pid:
            raise ValueError(f"iid '{iid}' cannot be equal to parent '{pid}'.")
        if pid and pid not in self.RI.tree:
            raise ValueError(f"parent '{parent}' does not exist.")
        if text is None:
            text = iid
        parent_node = self.RI.tree[pid] if parent else ""
        self.RI.tree[iid] = Node(text, iid, parent_node)
        if parent_node:
            if isinstance(index, int):
                idx = self.RI.tree_rns[pid] + index + 1
                for count, cid in enumerate(self.get_children(pid)):
                    if count >= index:
                        break
                    idx += sum(1 for _ in self.RI.get_iid_descendants(cid))
                self.RI.tree[pid].children.insert(index, self.RI.tree[iid])
            else:
                idx = self.RI.tree_rns[pid] + sum(1 for _ in self.RI.get_iid_descendants(pid)) + 1
                self.RI.tree[pid].children.append(self.RI.tree[iid])
        else:
            if isinstance(index, int):
                idx = index
                if index:
                    for count, cid in enumerate(self.get_children("")):
                        if count >= index:
                            break
                        idx += sum(1 for _ in self.RI.get_iid_descendants(cid)) + 1
            else:
                idx = len(self.MT._row_index)
        if values is None:
            values = []
        self.insert_rows(
            idx=idx,
            rows=[[self.RI.tree[iid]] + values],
            row_index=True,
            create_selections=create_selections,
            fill=False,
        )
        self.RI.tree_rns[iid] = idx
        if pid and (pid not in self.RI.tree_open_ids or not self.item_displayed(pid)):
            self.hide_rows(idx, deselect_all=False, data_indexes=True)
        return iid

    def item(
        self,
        item: str,
        iid: str | None = None,
        text: str | None = None,
        values: list | None = None,
        open_: bool | None = None,
        redraw: bool = True,
    ) -> DotDict | Sheet:
        """
        Modify options for item
        If no options are set then returns DotDict of options for item
        Else returns Sheet
        """
        if not (item := item.lower()) or item not in self.RI.tree:
            raise ValueError(f"Item '{item}' does not exist.")
        if isinstance(iid, str):
            if iid in self.RI.tree:
                raise ValueError(f"Cannot rename '{iid}', it already exists.")
            iid = iid.lower()
            self.RI.tree[item].iid = iid
            self.RI.tree[iid] = self.RI.tree.pop(item)
            self.RI.tree_rns[iid] = self.RI.tree_rns.pop(item)
            if iid in self.RI.tree_open_ids:
                self.RI.tree_open_ids[iid] = self.RI.tree_open_ids.pop(item)
        if isinstance(text, str):
            self.RI.tree[item].text = text
        if isinstance(values, list):
            self.set_data(self.RI.tree_rns[item], data=values)
        if isinstance(open_, bool):
            if self.RI.tree[item].children:
                if open_:
                    self.RI.tree_open_ids.add(item)
                    if self.item_displayed(item):
                        self.show_rows(
                            rows=(self.RI.tree_rns[did] for did in self.RI.get_iid_descendants(item, check_open=True)),
                            redraw=False,
                            deselect_all=False,
                        )
                else:
                    self.RI.tree_open_ids.discard(item)
                    if self.item_displayed(item):
                        self.hide_rows(
                            rows={self.RI.tree_rns[did] for did in self.RI.get_iid_descendants(item, check_open=True)},
                            redraw=False,
                            deselect_all=False,
                            data_indexes=True,
                        )
            else:
                self.RI.tree_open_ids.discard(item)
        get = not (isinstance(iid, str) or isinstance(text, str) or isinstance(values, list) or isinstance(open_, bool))
        self.set_refresh_timer(redraw=not get and redraw)
        if get:
            return DotDict(
                text=self.RI.tree[item].text,
                values=self[self.RI.tree_rns[item]].options(ndim=1).data,
                open_=item in self.RI.tree_open_ids,
            )
        return self

    def itemrow(self, item: str) -> int:
        try:
            return self.RI.tree_rns[item.lower()]
        except Exception:
            raise ValueError(f"item '{item.lower()}' does not exist.")

    def rowitem(self, row: int, data_index: bool = False) -> str | None:
        try:
            if data_index:
                return self.MT._row_index[row].iid
            return self.MT._row_index[self.MT.displayed_rows[row]].iid
        except Exception:
            return None

    def get_children(self, item: None | str = None) -> Generator[str]:
        if item is None:
            for n in self.RI.tree.values():
                if not n.parent:
                    yield n.iid
                    for iid in self.RI.get_iid_descendants(n.iid):
                        yield iid
        elif item == "":
            yield from (n.iid for n in self.RI.tree.values() if not n.parent)
        else:
            yield from (n.iid for n in self.RI.tree[item].children)

    def del_items(self, *items) -> Sheet:
        """
        Also deletes all descendants of items
        """
        rows_to_del = []
        iids_to_del = []
        for item in unpack(items):
            item = item.lower()
            if item not in self.RI.tree:
                continue
            rows_to_del.append(self.RI.tree_rns[item])
            iids_to_del.append(item)
            for did in self.RI.get_iid_descendants(item):
                rows_to_del.append(self.RI.tree_rns[did])
                iids_to_del.append(did)
        for item in unpack(items):
            self.RI.remove_node_from_parents_children(self.RI.tree[item])
        self.del_rows(rows_to_del)
        for iid in iids_to_del:
            self.RI.tree_open_ids.discard(iid)
            if self.RI.tree[iid].parent and len(self.RI.tree[iid].parent.children) == 1:
                self.RI.tree_open_ids.discard(self.RI.tree[iid].parent.iid)
            del self.RI.tree[iid]
        return self.set_refresh_timer()

    def set_children(self, parent: str, *newchildren) -> Sheet:
        """
        Moves everything in '*newchildren' under 'parent'
        """
        for iid in unpack(newchildren):
            self.move(iid, parent)
        return self

    def top_index_row(self, index: int) -> int:
        try:
            return next(self.RI.tree_rns[n.iid] for i, n in enumerate(self.RI.gen_top_nodes()) if i == index)
        except Exception:
            return None

    def move(self, item: str, parent: str, index: int | None = None) -> Sheet:
        """
        Moves item to be under parent as child at index
        'parent' can be empty string which will make item a top node
        """
        if (item := item.lower()) and item not in self.RI.tree:
            raise ValueError(f"Item '{item}' does not exist.")
        if (parent := parent.lower()) and parent not in self.RI.tree:
            raise ValueError(f"Parent '{parent}' does not exist.")
        mapping = {}
        to_show = []
        item_node = self.RI.tree[item]
        if parent:
            if self.RI.pid_causes_recursive_loop(item, parent):
                raise ValueError(f"iid '{item}' causes a recursive loop with parent '{parent}'.")
            parent_node = self.RI.tree[parent]
            if parent_node.children:
                if index is None or index >= len(parent_node.children):
                    index = len(parent_node.children) - 1
                item_r = self.RI.tree_rns[item]
                new_r = self.RI.tree_rns[parent_node.children[index].iid]
                new_r_desc = sum(1 for _ in self.RI.get_iid_descendants(parent_node.children[index].iid))
                item_desc = sum(1 for _ in self.RI.get_iid_descendants(item))
                if item_r < new_r:
                    r_ctr = new_r + new_r_desc - item_desc
                else:
                    r_ctr = new_r
            else:
                if index is None:
                    index = 0
                r_ctr = self.RI.tree_rns[parent_node.iid] + 1
            mapping[item_r] = r_ctr
            if parent in self.RI.tree_open_ids and self.item_displayed(parent):
                to_show.append(r_ctr)
            r_ctr += 1
            for did in self.RI.get_iid_descendants(item):
                mapping[self.RI.tree_rns[did]] = r_ctr
                if to_show and self.RI.ancestors_all_open(did, item_node.parent):
                    to_show.append(r_ctr)
                r_ctr += 1
            if parent == self.RI.items_parent(item):
                pop_index = parent_node.children.index(item_node)
                parent_node.children.insert(index, parent_node.children.pop(pop_index))
            else:
                self.RI.remove_node_from_parents_children(item_node)
                item_node.parent = parent_node
                parent_node.children.insert(index, item_node)
        else:
            if index is None:
                new_r = self.top_index_row((sum(1 for _ in self.RI.gen_top_nodes()) - 1))
            else:
                if (new_r := self.top_index_row(index)) is None:
                    new_r = self.top_index_row((sum(1 for _ in self.RI.gen_top_nodes()) - 1))
            item_r = self.RI.tree_rns[item]
            if item_r < new_r:
                par_desc = sum(1 for _ in self.RI.get_iid_descendants(self.rowitem(new_r, data_index=True)))
                item_desc = sum(1 for _ in self.RI.get_iid_descendants(item))
                r_ctr = new_r + par_desc - item_desc
            else:
                r_ctr = new_r
            mapping[item_r] = r_ctr
            to_show.append(r_ctr)
            r_ctr += 1
            for did in self.RI.get_iid_descendants(item):
                mapping[self.RI.tree_rns[did]] = r_ctr
                if to_show and self.RI.ancestors_all_open(did, item_node.parent):
                    to_show.append(r_ctr)
                r_ctr += 1
            self.RI.remove_node_from_parents_children(item_node)
            self.RI.tree[item].parent = ""
        self.mapping_move_rows(
            data_new_idxs=mapping,
            data_indexes=True,
            create_selections=False,
            redraw=False,
        )
        if parent and (parent not in self.RI.tree_open_ids or not self.item_displayed(parent)):
            self.hide_rows(set(mapping.values()), data_indexes=True)
        self.show_rows(to_show)
        return self.set_refresh_timer()

    reattach = move

    def exists(self, item: str) -> bool:
        return item.lower() in self.RI.tree

    def parent(self, item: str) -> str:
        if (item := item.lower()) not in self.RI.tree:
            raise ValueError(f"Item '{item}' does not exist.")
        return self.RI.tree[item].parent.iid if self.RI.tree[item].parent else self.RI.tree[item].parent

    def index(self, item: str) -> int:
        if (item := item.lower()) not in self.RI.tree:
            raise ValueError(f"Item '{item}' does not exist.")
        if not self.RI.tree[item].parent:
            return next(index for index, node in enumerate(self.RI.gen_top_nodes()) if node == self.RI.tree[item])
        return self.RI.tree[item].parent.children.index(self.RI.tree[item])

    def item_displayed(self, item: str) -> bool:
        """
        Check if an item (row) is currently displayed on the sheet
        - Does not check if the item is visible to the user
        """
        if (item := item.lower()) not in self.RI.tree:
            raise ValueError(f"Item '{item}' does not exist.")
        return self.RI.tree_rns[item] in self.MT.displayed_rows

    def display_item(self, item: str, redraw=False) -> Sheet:
        """
        Ensure that item is displayed in the tree
        - Opens all of an item's ancestors
        - Unlike the ttk treeview 'see' function
          this function does **NOT** scroll to the item
        """
        if (item := item.lower()) not in self.RI.tree:
            raise ValueError(f"Item '{item}' does not exist.")
        if self.RI.tree[item].parent:
            for iid in self.RI.get_iid_ancestors(item):
                self.item(iid, open_=True, redraw=False)
        return self.set_refresh_timer(redraw)

    def scroll_to_item(self, item: str, redraw=False) -> Sheet:
        """
        Scrolls to an item and ensures that it is displayed
        """
        if (item := item.lower()) not in self.RI.tree:
            raise ValueError(f"Item '{item}' does not exist.")
        self.display_item(item, redraw=False)
        self.see(
            row=bisect_left(self.MT.displayed_rows, self.RI.tree_rns[item]),
            keep_xscroll=True,
            redraw=False,
        )
        return self.set_refresh_timer(redraw)

    def selection(self, cells: bool = False) -> list[str]:
        """
        Get currently selected item ids
        """
        return [
            self.MT._row_index[self.displayed_row_to_data(rn)].iid
            for rn in self.get_selected_rows(get_cells_as_rows=cells)
        ]

    def selection_set(self, *items, redraw: bool = True) -> Sheet:
        if any(item.lower() in self.RI.tree for item in unpack(items)):
            self.deselect(redraw=False)
            self.selection_add(*items, redraw=False)
        return self.set_refresh_timer(redraw)

    def selection_add(self, *items, redraw: bool = True) -> Sheet:
        to_open = []
        quick_displayed_check = set(self.MT.displayed_rows)
        for item in unpack(items):
            if self.RI.tree_rns[(item := item.lower())] not in quick_displayed_check and self.RI.tree[item].parent:
                to_open.extend(list(self.RI.get_iid_ancestors(item)))
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
                    self.RI.tree_rns[item.lower()],
                )
                for item in unpack(items)
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
        self.MT.run_selection_binding("rows")
        return self.set_refresh_timer(redraw)

    def selection_remove(self, *items, redraw: bool = True) -> Sheet:
        for item in unpack(items):
            if (item := item.lower()) not in self.RI.tree:
                continue
            try:
                self.deselect(bisect_left(self.MT.displayed_rows, self.RI.tree_rns[item]), redraw=False)
            except Exception:
                continue
        return self.set_refresh_timer(redraw)

    def selection_toggle(self, *items, redraw: bool = True) -> Sheet:
        selected = set(self.MT._row_index[self.displayed_row_to_data(rn)].iid for rn in self.get_selected_rows())
        add = []
        remove = []
        for item in unpack(items):
            if (item := item.lower()) in self.RI.tree:
                if item in selected:
                    remove.append(item)
                else:
                    add.append(item)
        self.selection_remove(*remove, redraw=False)
        self.selection_add(*add, redraw=False)
        return self.set_refresh_timer(redraw)

    # Functions not in docs

    def event_generate(self, *args, **kwargs):
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

    def set_refresh_timer(self, redraw: bool = True) -> Sheet:
        if redraw and self.after_redraw_id is None:
            self.after_redraw_id = self.after(self.after_redraw_time_ms, self.after_redraw)
        return self

    def after_redraw(self) -> None:
        self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
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

    def get_cell_data(self, r: int, c: int, get_displayed: bool = False) -> object:
        return self.MT.get_cell_data(r, c, get_displayed)

    def get_row_data(
        self,
        r: int,
        get_displayed: bool = False,
        get_index: bool = False,
        get_index_displayed: bool = True,
        only_columns: int | Iterator[int] | None = None,
    ) -> list[object]:
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
        if get_index:
            return [self.get_index_data(r, get_displayed=get_index_displayed)] + [
                self.MT.get_cell_data(r, c, get_displayed=get_displayed) for c in iterable
            ]
        else:
            return [self.MT.get_cell_data(r, c, get_displayed=get_displayed) for c in iterable]

    def get_column_data(
        self,
        c: int,
        get_displayed: bool = False,
        get_header: bool = False,
        get_header_displayed: bool = True,
        only_rows: int | Iterator[int] | None = None,
    ) -> list[object]:
        if only_rows is not None:
            if isinstance(only_rows, int):
                only_rows = (only_rows,)
            elif not is_iterable(only_rows):
                raise ValueError(tksheet_type_error("only_rows", ["int", "iterable", "None"], only_rows))
        iterable = only_rows if only_rows is not None else range(len(self.MT.data))
        return ([self.get_header_data(c, get_displayed=get_header_displayed)] if get_header else []) + [
            self.MT.get_cell_data(r, c, get_displayed=get_displayed) for r in iterable
        ]

    def get_sheet_data(
        self,
        get_displayed: bool = False,
        get_header: bool = False,
        get_index: bool = False,
        get_header_displayed: bool = True,
        get_index_displayed: bool = True,
        only_rows: Iterator[int] | int | None = None,
        only_columns: Iterator[int] | int | None = None,
    ) -> list[object]:
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
                    data.append([self.get_index_data(rn, get_displayed=get_index_displayed)] + r)
                else:
                    data.append(r)
            iterable = only_columns if only_columns is not None else range(maxlen)
            if get_index:
                return [[""] + [self.get_header_data(cn, get_displayed=get_header_displayed) for cn in iterable]] + data
            else:
                return [[self.get_header_data(cn, get_displayed=get_header_displayed) for cn in iterable]] + data
        elif not get_header:
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
    ) -> Iterator[list[object]]:
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
            yield ([""] if get_index else []) + [
                self.get_header_data(c, get_displayed=get_header_displayed) for c in iterable
            ]
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
    ) -> object:
        self.set_refresh_timer(redraw)
        return self.MT.data_reference(newdataref, reset_col_positions, reset_row_positions)

    def set_cell_data(
        self,
        r: int,
        c: int,
        value: object = "",
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
        values=tuple(),
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
                            self.MT.insert_col_position("end")
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
        values=tuple(),
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
                        self.MT.insert_row_position("end", height=height)
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
        rows: list | int = [],
        readonly: bool = True,
        redraw: bool = False,
    ) -> Sheet:
        if isinstance(rows, int):
            rows = [rows]
        else:
            rows = rows
        if not readonly:
            for r in rows:
                if r in self.MT.row_options and "readonly" in self.MT.row_options[r]:
                    del self.MT.row_options[r]["readonly"]
        else:
            for r in rows:
                if r not in self.MT.row_options:
                    self.MT.row_options[r] = {}
                self.MT.row_options[r]["readonly"] = True
        return self.set_refresh_timer(redraw)

    def readonly_columns(
        self,
        columns: list | int = [],
        readonly: bool = True,
        redraw: bool = False,
    ) -> Sheet:
        if isinstance(columns, int):
            columns = [columns]
        else:
            columns = columns
        if not readonly:
            for c in columns:
                if c in self.MT.col_options and "readonly" in self.MT.col_options[c]:
                    del self.MT.col_options[c]["readonly"]
        else:
            for c in columns:
                if c not in self.MT.col_options:
                    self.MT.col_options[c] = {}
                self.MT.col_options[c]["readonly"] = True
        return self.set_refresh_timer(redraw)

    def readonly_cells(
        self,
        row: int = 0,
        column: int = 0,
        cells: list = [],
        readonly: bool = True,
        redraw: bool = False,
    ) -> Sheet:
        if not readonly:
            if cells:
                for r, c in cells:
                    if (r, c) in self.MT.cell_options and "readonly" in self.MT.cell_options[(r, c)]:
                        del self.MT.cell_options[(r, c)]["readonly"]
            else:
                if (
                    row,
                    column,
                ) in self.MT.cell_options and "readonly" in self.MT.cell_options[(row, column)]:
                    del self.MT.cell_options[(row, column)]["readonly"]
        else:
            if cells:
                for r, c in cells:
                    if (r, c) not in self.MT.cell_options:
                        self.MT.cell_options[(r, c)] = {}
                    self.MT.cell_options[(r, c)]["readonly"] = True
            else:
                if (row, column) not in self.MT.cell_options:
                    self.MT.cell_options[(row, column)] = {}
                self.MT.cell_options[(row, column)]["readonly"] = True
        return self.set_refresh_timer(redraw)

    def readonly_header(
        self,
        columns: list = [],
        readonly: bool = True,
        redraw: bool = False,
    ) -> Sheet:
        self.CH.readonly_header(columns=columns, readonly=readonly)
        return self.set_refresh_timer(redraw)

    def readonly_index(
        self,
        rows: list = [],
        readonly: bool = True,
        redraw: bool = False,
    ) -> Sheet:
        self.RI.readonly_index(rows=rows, readonly=readonly)
        return self.set_refresh_timer(redraw)

    def dehighlight_rows(
        self,
        rows: list[int] | Literal["all"] = [],
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
                try:
                    del self.MT.row_options[r]["highlight"]
                except Exception:
                    pass
                try:
                    del self.RI.cell_options[r]["highlight"]
                except Exception:
                    pass
        return self.set_refresh_timer(redraw)

    def dehighlight_columns(
        self,
        columns: list[int] | Literal["all"] = [],
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
                try:
                    del self.MT.col_options[c]["highlight"]
                except Exception:
                    pass
                try:
                    del self.CH.cell_options[c]["highlight"]
                except Exception:
                    pass
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
        cells: list[tuple[int, int]] = [],
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
        cells: list[tuple[int, int]] = [],
        canvas: Literal["table", "row_index", "header"] = "table",
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
        elif canvas == "row_index":
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

    def get_highlighted_cells(self, canvas: str = "table") -> dict | None:
        if canvas == "table":
            return {k: v["highlight"] for k, v in self.MT.cell_options.items() if "highlight" in v}
        elif canvas == "row_index":
            return {k: v["highlight"] for k, v in self.RI.cell_options.items() if "highlight" in v}
        elif canvas == "header":
            return {k: v["highlight"] for k, v in self.CH.cell_options.items() if "highlight" in v}

    def align_cells(
        self,
        row: int = 0,
        column: int = 0,
        cells: list = [],
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
        rows: list | dict | int = [],
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
        columns: list | dict | int = [],
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
        columns: list | dict | int = [],
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
        rows: list | dict | int = [],
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

    def checkbox_row(self, r: Iterator[int] | int | Literal["all"] = 0, *args, **kwargs) -> None:
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

    def create_header_checkbox(self, c: Iterator[int] | int | Literal["all"] = 0, *args, **kwargs) -> None:
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

    def _create_header_checkbox(self, c: int, v: bool, d: dict) -> None:
        self.del_header_cell_options_dropdown_and_checkbox(c)
        add_to_options(self.CH.cell_options, c, "checkbox", d)
        self.CH.set_cell_data(c, v)

    def create_index_checkbox(self, r: Iterator[int] | int | Literal["all"] = 0, *args, **kwargs) -> None:
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
            for r_, c_ in self.MT.cell_options:
                if "checkbox" in self.MT.cell_options[(r_, c)]:
                    self.del_cell_options_checkbox(r_, c)
        elif isinstance(c, str) and c.lower() == "all" and isinstance(r, int):
            for r_, c_ in self.MT.cell_options:
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

    def delete_row_checkbox(self, r: Iterator[int] | int | Literal["all"] = 0) -> None:
        if isinstance(r, str) and r.lower() == "all":
            for r_ in self.MT.row_options:
                self.del_table_row_options_checkbox(r_)
        elif isinstance(r, int):
            self.del_table_row_options_checkbox(r)
        elif is_iterable(r):
            for r_ in r:
                self.del_table_row_options_checkbox(r_)

    def delete_column_checkbox(self, c: Iterator[int] | int | Literal["all"] = 0) -> None:
        if isinstance(c, str) and c.lower() == "all":
            for c_ in self.MT.col_options:
                self.del_table_column_options_checkbox(c_)
        elif isinstance(c, int):
            self.del_table_column_options_checkbox(c)
        elif is_iterable(c):
            for c_ in c:
                self.del_table_column_options_checkbox(c_)

    def delete_header_checkbox(self, c: Iterator[int] | int | Literal["all"] = 0) -> None:
        if isinstance(c, str) and c.lower() == "all":
            for c_ in self.CH.cell_options:
                if "checkbox" in self.CH.cell_options[c_]:
                    self.del_header_cell_options_checkbox(c_)
        if isinstance(c, int):
            self.del_header_cell_options_checkbox(c)
        elif is_iterable(c):
            for c_ in c:
                self.del_header_cell_options_checkbox(c_)

    def delete_index_checkbox(self, r: Iterator[int] | int | Literal["all"] = 0) -> None:
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
        if kwargs["set_value"] is None:
            if kwargs["values"] and (v := self.MT.get_cell_data(r, c)) not in kwargs["values"]:
                v = kwargs["values"][0]
            else:
                v == ""
        else:
            v = kwargs["set_value"]
        if isinstance(r, str) and r.lower() == "all" and isinstance(c, int):
            for r_ in range(self.MT.total_data_rows()):
                self._create_dropdown(r_, c, v, d)
        elif isinstance(c, str) and c.lower() == "all" and isinstance(r, int):
            for c_ in range(self.MT.total_data_cols()):
                self._create_dropdown(r, c_, v, d)
        elif isinstance(r, str) and r.lower() == "all" and isinstance(c, str) and c.lower() == "all":
            totalcols = self.MT.total_data_cols()
            for r_ in range(self.MT.total_data_rows()):
                for c_ in range(totalcols):
                    self._create_dropdown(r_, c_, v, d)
        elif isinstance(r, int) and isinstance(c, int):
            self._create_dropdown(r, c, v, d)
        return self.set_refresh_timer(kwargs["redraw"])

    def _create_dropdown(self, r: int, c: int, v: object, d: dict) -> None:
        self.del_cell_options_dropdown_and_checkbox(r, c)
        add_to_options(self.MT.cell_options, (r, c), "dropdown", d)
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
        if isinstance(r, str) and r.lower() == "all":
            for r_ in range(self.MT.total_data_rows()):
                self._dropdown_row(r_, v, d)
        elif isinstance(r, int):
            self._dropdown_row(r, v, d)
        elif is_iterable(r):
            for r_ in r:
                self._dropdown_row(r_, v, d)
        return self.set_refresh_timer(kwargs["redraw"])

    def _dropdown_row(self, r: int, v: object, d: dict) -> None:
        self.del_row_options_dropdown_and_checkbox(r)
        add_to_options(self.MT.row_options, r, "dropdown", d)
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
        if isinstance(c, str) and c.lower() == "all":
            for c_ in range(self.MT.total_data_cols()):
                self._dropdown_column(c_, v, d)
        elif isinstance(c, int):
            self._dropdown_column(c, v, d)
        elif is_iterable(c):
            for c_ in c:
                self._dropdown_column(c_, v, d)
        return self.set_refresh_timer(kwargs["redraw"])

    def _dropdown_column(self, c: int, v: object, d: dict) -> None:
        self.del_column_options_dropdown_and_checkbox(c)
        add_to_options(self.MT.col_options, c, "dropdown", d)
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
        if isinstance(c, str) and c.lower() == "all":
            for c_ in range(self.MT.total_data_cols()):
                self._create_header_dropdown(c_, v, d)
        elif isinstance(c, int):
            self._create_header_dropdown(c, v, d)
        elif is_iterable(c):
            for c_ in c:
                self._create_header_dropdown(c_, v, d)
        return self.set_refresh_timer(kwargs["redraw"])

    def _create_header_dropdown(self, c: int, v: object, d: dict) -> None:
        self.del_header_cell_options_dropdown_and_checkbox(c)
        add_to_options(self.CH.cell_options, c, "dropdown", d)
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
        if isinstance(r, str) and r.lower() == "all":
            for r_ in range(self.MT.total_data_rows()):
                self._create_index_dropdown(r_, v, d)
        elif isinstance(r, int):
            self._create_index_dropdown(r, v, d)
        elif is_iterable(r):
            for r_ in r:
                self._create_index_dropdown(r_, v, d)
        return self.set_refresh_timer(kwargs["redraw"])

    def _create_index_dropdown(self, r: int, v: object, d: dict) -> None:
        self.del_index_cell_options_dropdown_and_checkbox(r)
        add_to_options(self.RI.cell_options, r, "dropdown", d)
        self.RI.set_cell_data(r, v)

    def delete_dropdown(
        self,
        r: int | Literal["all"] = 0,
        c: int | Literal["all"] = 0,
    ) -> None:
        if isinstance(r, str) and r.lower() == "all" and isinstance(c, int):
            for r_, c_ in self.MT.cell_options:
                if "dropdown" in self.MT.cell_options[(r_, c)]:
                    self.del_cell_options_dropdown(r_, c)
        elif isinstance(c, str) and c.lower() == "all" and isinstance(r, int):
            for r_, c_ in self.MT.cell_options:
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
        if "dropdown" in self.MT.options:
            return {**d, "dropdown": self.MT.options["dropdown"]}
        return d

    def get_header_dropdowns(self) -> dict:
        d = {k: v["dropdown"] for k, v in self.CH.cell_options.items() if "dropdown" in v}
        if "dropdown" in self.CH.options:
            return {**d, "dropdown": self.CH.options["dropdown"]}
        return d

    def get_index_dropdowns(self) -> dict:
        d = {k: v["dropdown"] for k, v in self.RI.cell_options.items() if "dropdown" in v}
        if "dropdown" in self.RI.options:
            return {**d, "dropdown": self.RI.options["dropdown"]}
        return d

    def set_dropdown_values(
        self,
        r: int = 0,
        c: int = 0,
        set_existing_dropdown: bool = False,
        values: list[object] = [],
        set_value: object = None,
    ) -> Sheet:
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
        values: list[object] = [],
        set_value: object = None,
    ) -> Sheet:
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
        values: list[object] = [],
        set_value: object = None,
    ) -> Sheet:
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
                # here
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

    def get_dropdown_value(self, r: int = 0, c: int = 0) -> object:
        if self.MT.get_cell_kwargs(r, c, key="dropdown"):
            return self.get_cell_data(r, c)

    def get_header_dropdown_value(self, c: int = 0) -> object:
        if self.CH.get_cell_kwargs(c, key="dropdown"):
            return self.MT._headers[c]

    def get_index_dropdown_value(self, r: int = 0) -> object:
        if self.RI.get_cell_kwargs(r, key="dropdown"):
            return self.MT._row_index[r]

    def delete_all_formatting(self, clear_values: bool = False) -> None:
        self.MT.delete_all_formatting(clear_values=clear_values)

    def format_cell(
        self,
        r: int | Literal["all"],
        c: int | Literal["all"],
        formatter_options: dict = {},
        formatter_class: object = None,
        redraw: bool = True,
        **kwargs,
    ) -> Sheet:
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
            for r_, c_ in self.MT.cell_options:
                if "format" in self.MT.cell_options[(r_, c)]:
                    self.MT.delete_cell_format(r_, c, clear_values=clear_values)
        elif isinstance(c, str) and c.lower() == "all" and isinstance(r, int):
            for r_, c_ in self.MT.cell_options:
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
        formatter_options: dict = {},
        formatter_class: object = None,
        redraw: bool = True,
        **kwargs,
    ) -> Sheet:
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
        formatter_options: dict = {},
        formatter_class: object = None,
        redraw: bool = True,
        **kwargs,
    ) -> Sheet:
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
        ops: dict,
        outline_color: str,
        width: int | None = None,
        height: int | None = None,
        font: None | tuple[str, int, str] = None,
        outline_thickness: int = 2,
        values: list[object] = [],
        close_dropdown_window: Callable | None = None,
        search_function: Callable = dropdown_search_function,
        arrowkey_RIGHT: Callable | None = None,
        arrowkey_LEFT: Callable | None = None,
        align: str = "w",
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
            selected_rows_to_end_of_window=True,
            horizontal_grid_to_end_of_window=True,
            set_cell_sizes_on_zoom=True,
            show_selected_cells_border=False,
            scrollbar_show_arrows=False,
        )
        self.parent = parent
        self.close_dropdown_window = close_dropdown_window
        self.search_function = search_function
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
        self.reset(r, c, width, height, font, ops, outline_color, align, values)

    def reset(
        self,
        r: int,
        c: int,
        width: int,
        height: int,
        font: tuple[str, int, str],
        ops: DotDict,
        outline_color: str,
        align: str,
        values: list[object] | None = None,
    ) -> None:
        self.deselect(redraw=False)
        self.r = r
        self.c = c
        self.row = -1
        self.height_and_width(height=height, width=width)
        self.table_align(align)
        self.set_options(
            outline_color=outline_color,
            table_grid_fg=ops.popup_menu_fg,
            table_selected_cells_border_fg=ops.popup_menu_fg,
            table_selected_cells_bg=ops.popup_menu_highlight_bg,
            table_selected_rows_border_fg=ops.popup_menu_fg,
            table_selected_rows_bg=ops.popup_menu_highlight_bg,
            table_selected_rows_fg=ops.popup_menu_highlight_fg,
            table_selected_box_cells_fg=ops.popup_menu_highlight_bg,
            table_selected_box_rows_fg=ops.popup_menu_highlight_bg,
            font=font,
            table_fg=ops.popup_menu_fg,
            table_bg=ops.popup_menu_bg,
            **{k: ops[k] for k in scrollbar_options_keys},
        )
        self.values(values, width=width - self.yscroll.winfo_width() - 4)

    def arrowkey_UP(self, event: object = None) -> None:
        if self.row > 0:
            self.row -= 1
        else:
            self.row = 0
        self.see(self.row, 0, redraw=False)
        self.select_row(self.row)

    def arrowkey_DOWN(self, event: object = None) -> None:
        if len(self.MT.data) - 1 > self.row:
            self.row += 1
        self.see(self.row, 0, redraw=False)
        self.select_row(self.row)

    def search_and_see(self, event: object = None) -> None:
        if self.search_function is not None:
            rn = self.search_function(search_for=rf"{event['value']}".lower(), data=self.MT.data)
            if rn is not None:
                self.row = rn
                self.see(self.row, 0, redraw=False)
                self.select_row(self.row)

    def mouse_motion(self, event: object) -> None:
        row = self.identify_row(event, exclude_index=True, allow_end=False)
        if row is not None and row != self.row:
            self.row = row
            self.select_row(self.row)

    def _reselect(self) -> None:
        rows = self.get_selected_rows()
        if rows:
            self.select_row(next(iter(rows)))

    def b1(self, event: object = None) -> None:
        if event is None:
            row = None
        elif event.keysym == "Return":
            row = self.get_selected_rows()
            if not row:
                row = None
            else:
                row = next(iter(row))
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
        values: list = [],
        redraw: bool = True,
        width: int | None = None,
    ) -> None:
        self.set_sheet_data(
            [[v] for v in values],
            reset_col_positions=False,
            reset_row_positions=False,
            redraw=False,
            verify=False,
        )
        self.set_all_cell_sizes_to_text(redraw=redraw, width=width, slim=True)
