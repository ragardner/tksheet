from __future__ import annotations

import tkinter as tk
from bisect import bisect_left
from collections import deque
from collections.abc import (
    Callable,
    Iterator,
    Sequence,
)
from itertools import accumulate, chain, islice
from tkinter import ttk

from .column_headers import ColumnHeaders
from .functions import (
    alpha2num,
    data_to_displayed_idxs,
    dropdown_search_function,
    ev_stack_dict,
    event_dict,
    get_checkbox_dict,
    get_checkbox_kwargs,
    get_dropdown_dict,
    get_dropdown_kwargs,
    is_iterable,
    named_span_dict,
    num2alpha,
    tksheet_type_error,
    str_to_int,
)
from .main_table import MainTable
from .other_classes import (
    GeneratedMouseEvent,
    Span,
)
from .row_index import RowIndex
from .top_left_rectangle import TopLeftRectangle
from .vars import (
    emitted_events,
    get_font,
    get_header_font,
    get_index_font,
    rc_binding,
    theme_black,
    theme_dark,
    theme_dark_blue,
    theme_dark_green,
    theme_light_blue,
    theme_light_green,
)


class Sheet(tk.Frame):
    def __init__(
        self,
        parent,
        name: str = "!sheet",
        show_table: bool = True,
        show_top_left: bool = True,
        show_row_index: bool = True,
        show_header: bool = True,
        show_x_scrollbar: bool = True,
        show_y_scrollbar: bool = True,
        width: int = None,
        height: int = None,
        headers: list = None,
        header: list = None,
        default_header: str = "letters",  # letters, numbers or both
        default_row_index: str = "numbers",  # letters, numbers or both
        to_clipboard_delimiter="\t",
        to_clipboard_quotechar='"',
        to_clipboard_lineterminator="\n",
        from_clipboard_delimiters=["\t"],
        show_default_header_for_empty: bool = True,
        show_default_index_for_empty: bool = True,
        page_up_down_select_row: bool = True,
        expand_sheet_if_paste_too_big: bool = False,
        paste_insert_column_limit: int = None,
        paste_insert_row_limit: int = None,
        show_dropdown_borders: bool = False,
        arrow_key_down_right_scroll_page: bool = False,
        enable_edit_cell_auto_resize: bool = True,
        edit_cell_validation: bool = True,
        data_reference: list = None,
        data: list = None,
        # either (start row, end row, "rows"), (start column, end column, "rows") or
        # (cells start row, cells start column, cells end row, cells end column, "cells")  # noqa: E501
        startup_select: tuple = None,
        startup_focus: bool = True,
        total_columns: int = None,
        total_rows: int = None,
        column_width: int = 120,
        header_height: str = "1",  # str or int
        max_column_width: str = "inf",  # str or int
        max_row_height: str = "inf",  # str or int
        max_header_height: str = "inf",  # str or int
        max_index_width: str = "inf",  # str or int
        row_index: list = None,
        index: list = None,
        after_redraw_time_ms: int = 20,
        row_index_width: int = None,
        auto_resize_default_row_index: bool = True,
        auto_resize_columns: int | None = None,
        auto_resize_rows: int | None = None,
        set_all_heights_and_widths: bool = False,
        row_height: str = "1",  # str or int
        font: tuple = get_font(),
        header_font: tuple = get_header_font(),
        index_font: tuple = get_index_font(),  # currently has no effect
        popup_menu_font: tuple = get_font(),
        align: str = "w",
        header_align: str = "center",
        row_index_align: str | None = None,
        index_align: str = "center",
        displayed_columns: list = [],
        all_columns_displayed: bool = True,
        displayed_rows: list = [],
        all_rows_displayed: bool = True,
        max_undos: int = 30,
        outline_thickness: int = 0,
        outline_color: str = theme_light_blue["outline_color"],
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
        theme: str = "light blue",
        popup_menu_fg: str = theme_light_blue["popup_menu_fg"],
        popup_menu_bg: str = theme_light_blue["popup_menu_bg"],
        popup_menu_highlight_bg: str = theme_light_blue["popup_menu_highlight_bg"],
        popup_menu_highlight_fg: str = theme_light_blue["popup_menu_highlight_fg"],
        frame_bg: str = theme_light_blue["table_bg"],
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
    ):
        tk.Frame.__init__(
            self,
            parent,
            background=frame_bg,
            highlightthickness=outline_thickness,
            highlightbackground=outline_color,
            highlightcolor=outline_color,
        )
        self.C = parent
        self.name = name
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
            parentframe=self,
            row_index_align=self.convert_align(row_index_align)
            if row_index_align is not None
            else self.convert_align(index_align),
            index_bg=index_bg,
            index_border_fg=index_border_fg,
            index_grid_fg=index_grid_fg,
            index_fg=index_fg,
            index_selected_cells_bg=index_selected_cells_bg,
            index_selected_cells_fg=index_selected_cells_fg,
            index_selected_rows_bg=index_selected_rows_bg,
            index_selected_rows_fg=index_selected_rows_fg,
            index_hidden_rows_expander_bg=index_hidden_rows_expander_bg,
            drag_and_drop_bg=drag_and_drop_bg,
            resizing_line_fg=resizing_line_fg,
            row_drag_and_drop_perform=row_drag_and_drop_perform,
            default_row_index=default_row_index,
            auto_resize_width=auto_resize_default_row_index,
            show_default_index_for_empty=show_default_index_for_empty,
        )
        self.CH = ColumnHeaders(
            parentframe=self,
            default_header=default_header,
            header_align=self.convert_align(header_align),
            header_bg=header_bg,
            header_border_fg=header_border_fg,
            header_grid_fg=header_grid_fg,
            header_fg=header_fg,
            header_selected_cells_bg=header_selected_cells_bg,
            header_selected_cells_fg=header_selected_cells_fg,
            header_selected_columns_bg=header_selected_columns_bg,
            header_selected_columns_fg=header_selected_columns_fg,
            header_hidden_columns_expander_bg=header_hidden_columns_expander_bg,
            drag_and_drop_bg=drag_and_drop_bg,
            column_drag_and_drop_perform=column_drag_and_drop_perform,
            resizing_line_fg=resizing_line_fg,
            show_default_header_for_empty=show_default_header_for_empty,
        )
        self.MT = MainTable(
            parentframe=self,
            max_column_width=max_column_width,
            max_header_height=max_header_height,
            max_row_height=max_row_height,
            max_index_width=max_index_width,
            row_index_width=row_index_width,
            header_height=header_height,
            column_width=column_width,
            row_height=row_height,
            show_index=show_row_index,
            show_header=show_header,
            enable_edit_cell_auto_resize=enable_edit_cell_auto_resize,
            edit_cell_validation=edit_cell_validation,
            page_up_down_select_row=page_up_down_select_row,
            expand_sheet_if_paste_too_big=expand_sheet_if_paste_too_big,
            paste_insert_column_limit=paste_insert_column_limit,
            paste_insert_row_limit=paste_insert_row_limit,
            show_dropdown_borders=show_dropdown_borders,
            arrow_key_down_right_scroll_page=arrow_key_down_right_scroll_page,
            display_selected_fg_over_highlights=display_selected_fg_over_highlights,
            show_vertical_grid=show_vertical_grid,
            show_horizontal_grid=show_horizontal_grid,
            to_clipboard_delimiter=to_clipboard_delimiter,
            to_clipboard_quotechar=to_clipboard_quotechar,
            to_clipboard_lineterminator=to_clipboard_lineterminator,
            from_clipboard_delimiters=from_clipboard_delimiters,
            column_headers_canvas=self.CH,
            row_index_canvas=self.RI,
            headers=headers,
            header=header,
            data_reference=data if data_reference is None else data_reference,
            auto_resize_columns=auto_resize_columns,
            auto_resize_rows=auto_resize_rows,
            total_cols=total_columns,
            total_rows=total_rows,
            row_index=row_index,
            index=index,
            font=font,
            header_font=header_font,
            index_font=index_font,
            popup_menu_font=popup_menu_font,
            popup_menu_fg=popup_menu_fg,
            popup_menu_bg=popup_menu_bg,
            popup_menu_highlight_bg=popup_menu_highlight_bg,
            popup_menu_highlight_fg=popup_menu_highlight_fg,
            align=self.convert_align(align),
            table_bg=table_bg,
            table_grid_fg=table_grid_fg,
            table_fg=table_fg,
            show_selected_cells_border=show_selected_cells_border,
            table_selected_box_cells_fg=table_selected_box_cells_fg,
            table_selected_box_rows_fg=table_selected_box_rows_fg,
            table_selected_box_columns_fg=table_selected_box_columns_fg,
            table_selected_cells_border_fg=table_selected_cells_border_fg,
            table_selected_cells_bg=table_selected_cells_bg,
            table_selected_cells_fg=table_selected_cells_fg,
            table_selected_rows_border_fg=table_selected_rows_border_fg,
            table_selected_rows_bg=table_selected_rows_bg,
            table_selected_rows_fg=table_selected_rows_fg,
            table_selected_columns_border_fg=table_selected_columns_border_fg,
            table_selected_columns_bg=table_selected_columns_bg,
            table_selected_columns_fg=table_selected_columns_fg,
            displayed_columns=displayed_columns,
            all_columns_displayed=all_columns_displayed,
            displayed_rows=displayed_rows,
            all_rows_displayed=all_rows_displayed,
            selected_rows_to_end_of_window=selected_rows_to_end_of_window,
            horizontal_grid_to_end_of_window=horizontal_grid_to_end_of_window,
            vertical_grid_to_end_of_window=vertical_grid_to_end_of_window,
            empty_horizontal=empty_horizontal,
            empty_vertical=empty_vertical,
            max_undos=max_undos,
        )
        self.TL = TopLeftRectangle(
            parentframe=self,
            main_canvas=self.MT,
            row_index_canvas=self.RI,
            header_canvas=self.CH,
            top_left_bg=top_left_bg,
            top_left_fg=top_left_fg,
            top_left_fg_highlight=top_left_fg_highlight,
        )
        self.yscroll = ttk.Scrollbar(self, command=self.MT.set_yviews, orient="vertical")
        self.xscroll = ttk.Scrollbar(self, command=self.MT.set_xviews, orient="horizontal")
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

    def set_refresh_timer(self, redraw: bool = True) -> None:
        if redraw and self.after_redraw_id is None:
            self.after_redraw_id = self.after(self.after_redraw_time_ms, self.after_redraw)

    def after_redraw(self, redraw_header: bool = True, redraw_row_index: bool = True) -> None:
        self.MT.main_table_redraw_grid_and_text(redraw_header=redraw_header, redraw_row_index=redraw_row_index)
        self.after_redraw_id = None

    def show(self, canvas: str = "all") -> None:
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

    def hide(self, canvas: str = "all") -> None:
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

    def height_and_width(self, height: int | None = None, width: int | None = None) -> None:
        if width is not None or height is not None:
            self.grid_propagate(0)
        elif width is None and height is None:
            self.grid_propagate(1)
        if width is not None:
            self.config(width=width)
        if height is not None:
            self.config(height=height)

    def focus_set(self, canvas: str = "table") -> None:
        if canvas == "table":
            self.MT.focus_set()
        elif canvas == "header":
            self.CH.focus_set()
        elif canvas == "index":
            self.RI.focus_set()
        elif canvas == "topleft":
            self.TL.focus_set()

    def displayed_column_to_data(self, c: int) -> int:
        return c if self.MT.all_columns_displayed else self.MT.displayed_columns[c]

    def displayed_row_to_data(self, r: int) -> int:
        return r if self.MT.all_rows_displayed else self.MT.displayed_rows[r]

    def popup_menu_add_command(
        self,
        label: str,
        func: Callable,
        table_menu: bool = True,
        index_menu: bool = True,
        header_menu: bool = True,
        empty_space_menu: bool = True,
    ) -> None:
        if label not in self.MT.extra_table_rc_menu_funcs and table_menu:
            self.MT.extra_table_rc_menu_funcs[label] = func
        if label not in self.MT.extra_index_rc_menu_funcs and index_menu:
            self.MT.extra_index_rc_menu_funcs[label] = func
        if label not in self.MT.extra_header_rc_menu_funcs and header_menu:
            self.MT.extra_header_rc_menu_funcs[label] = func
        if label not in self.MT.extra_empty_space_rc_menu_funcs and empty_space_menu:
            self.MT.extra_empty_space_rc_menu_funcs[label] = func
        self.MT.create_rc_menus()

    def popup_menu_del_command(self, label: str | None = None) -> None:
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

    def extra_bindings(self, bindings: str | list | tuple, func: Callable | None = None) -> None:
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

            if func is not None and b in emitted_events:
                self.bind_event(b, f)

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

    def emit_event(self, event: str, data: None | dict = None) -> None:
        if data is None:
            data = {}
        data["sheetname"] = self.name
        self.event_generate(event, data=data)

    def bind_event(self, sequence: str, func: Callable, add: str | None = None) -> None:
        widget = self

        def _substitute(*args) -> tuple[None]:
            def e() -> None:
                return None

            e.data = args[0]
            e.widget = widget
            return (e,)

        funcid = widget._register(func, _substitute, needcleanup=1)
        cmd = '{0}if {{"[{1} %d]" == "break"}} break\n'.format("+" if add else "", funcid)
        widget.tk.call("bind", widget._w, sequence, cmd)

    def bind(self, binding: str, func: Callable, add: str | None = None) -> None:
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
        else:
            self.MT.bind(binding, func, add=add)
            self.CH.bind(binding, func, add=add)
            self.RI.bind(binding, func, add=add)
            self.TL.bind(binding, func, add=add)

    def unbind(self, binding: str) -> None:
        if binding == "<ButtonPress-1>":
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

    def enable_bindings(self, *bindings) -> None:
        self.MT.enable_bindings(bindings)

    def disable_bindings(self, *bindings) -> None:
        self.MT.disable_bindings(bindings)

    def basic_bindings(self, enable: bool = False) -> None:
        for canvas in (self.MT, self.CH, self.RI, self.TL):
            canvas.basic_bindings(enable)

    def edit_bindings(self, enable: bool = False) -> None:
        if enable:
            self.MT.edit_bindings(True)
        elif not enable:
            self.MT.edit_bindings(False)

    def cell_edit_binding(self, enable: bool = False, keys: list = []) -> None:
        self.MT.bind_cell_edit(enable, keys=keys)

    def identify_region(self, event) -> str:
        if event.widget == self.MT:
            return "table"
        elif event.widget == self.RI:
            return "index"
        elif event.widget == self.CH:
            return "header"
        elif event.widget == self.TL:
            return "top left"

    def identify_row(self, event, exclude_index: bool = False, allow_end: bool = True) -> int | None:
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

    def identify_column(self, event, exclude_header: bool = False, allow_end: bool = True) -> int | None:
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

    def get_example_canvas_column_widths(self, total_cols: int | None = None) -> list:
        colpos = int(self.MT.default_column_width)
        if total_cols is not None:
            return list(accumulate(chain([0], (colpos for c in range(total_cols)))))
        return list(accumulate(chain([0], (colpos for c in range(len(self.MT.col_positions) - 1)))))

    def get_example_canvas_row_heights(self, total_rows=None):
        rowpos = self.MT.default_row_height[1]
        if total_rows is not None:
            return list(accumulate(chain([0], (rowpos for c in range(total_rows)))))
        return list(accumulate(chain([0], (rowpos for c in range(len(self.MT.row_positions) - 1)))))

    def get_column_widths(self, canvas_positions: bool = False):
        if canvas_positions:
            return [int(n) for n in self.MT.col_positions]
        return self.MT.get_column_widths()

    def get_row_heights(self, canvas_positions: bool = False):
        if canvas_positions:
            return [int(n) for n in self.MT.row_positions]
        return self.MT.get_row_heights()

    def set_all_cell_sizes_to_text(self, redraw: bool = True):
        self.MT.set_all_cell_sizes_to_text()
        self.set_refresh_timer(redraw)
        return self.MT.row_positions, self.MT.col_positions

    def set_all_column_widths(
        self,
        width=None,
        only_set_if_too_small: bool = False,
        redraw: bool = True,
        recreate_selection_boxes: bool = True,
    ):
        self.CH.set_width_of_all_cols(
            width=width,
            only_set_if_too_small=only_set_if_too_small,
            recreate=recreate_selection_boxes,
        )
        self.set_refresh_timer(redraw)

    def column_width(
        self,
        column=None,
        width=None,
        only_set_if_too_small: bool = False,
        redraw: bool = True,
    ):
        if column == "all":
            if width == "default":
                self.MT.reset_col_positions()
        elif column == "displayed":
            if width == "text":
                sc, ec = self.MT.get_visible_columns(self.MT.canvasx(0), self.MT.canvasx(self.winfo_width()))
                for c in range(sc, ec - 1):
                    self.CH.set_col_width(c)
        elif width == "text" and column is not None:
            self.CH.set_col_width(col=column, width=None, only_set_if_too_small=only_set_if_too_small)
        elif width is not None and column is not None:
            self.CH.set_col_width(col=column, width=width, only_set_if_too_small=only_set_if_too_small)
        elif column is not None:
            return int(self.MT.col_positions[column + 1] - self.MT.col_positions[column])
        self.set_refresh_timer(redraw)

    def set_column_widths(
        self,
        column_widths: Iterator[int, float] | None = None,
        canvas_positions: bool = False,
        reset: bool = False,
        verify: bool = False,
    ):
        cwx = None
        if reset:
            self.MT.reset_col_positions()
            return
        if verify:
            cwx = self.verify_column_widths(column_widths, canvas_positions)
        if is_iterable(column_widths):
            if canvas_positions and isinstance(column_widths, list):
                self.MT.col_positions = column_widths
            else:
                self.MT.col_positions = list(accumulate(chain([0], (width for width in column_widths))))
        return cwx

    def set_all_row_heights(
        self,
        height: None | int = None,
        only_set_if_too_small: bool = False,
        redraw: bool = True,
        recreate_selection_boxes: bool = True,
    ) -> None:
        self.RI.set_height_of_all_rows(
            height=height,
            only_set_if_too_small=only_set_if_too_small,
            recreate=recreate_selection_boxes,
        )
        self.set_refresh_timer(redraw)

    def set_cell_size_to_text(
        self,
        row,
        column,
        only_set_if_too_small: bool = False,
        redraw: bool = True,
    ):
        self.MT.set_cell_size_to_text(r=row, c=column, only_set_if_too_small=only_set_if_too_small)
        self.set_refresh_timer(redraw)

    def set_width_of_index_to_text(self, text=None, *args, **kwargs):
        self.RI.set_width_of_index_to_text(text=text)

    def set_height_of_header_to_text(self, text=None):
        self.CH.set_height_of_header_to_text(text=text)

    def row_height(
        self,
        row=None,
        height=None,
        only_set_if_too_small: bool = False,
        redraw: bool = True,
    ):
        if row == "all":
            if height == "default":
                self.MT.reset_row_positions()
        elif row == "displayed":
            if height == "text":
                sr, er = self.MT.get_visible_rows(self.MT.canvasy(0), self.MT.canvasy(self.winfo_width()))
                for r in range(sr, er - 1):
                    self.RI.set_row_height(r)
        elif height == "text" and row is not None:
            self.RI.set_row_height(row=row, height=None, only_set_if_too_small=only_set_if_too_small)
        elif height is not None and row is not None:
            self.RI.set_row_height(row=row, height=height, only_set_if_too_small=only_set_if_too_small)
        elif row is not None:
            return int(self.MT.row_positions[row + 1] - self.MT.row_positions[row])
        self.set_refresh_timer(redraw)

    def set_row_heights(
        self,
        row_heights: Iterator[int, float] | None = None,
        canvas_positions: bool = False,
        reset: bool = False,
        verify: bool = False,
    ) -> None:
        if reset:
            self.MT.reset_row_positions()
            return
        if is_iterable(row_heights):
            qmin = self.MT.min_row_height
            if canvas_positions and isinstance(row_heights, list):
                if verify:
                    self.MT.row_positions = list(
                        accumulate(
                            chain(
                                [0],
                                (
                                    height if qmin < height else qmin
                                    for height in [
                                        x - z
                                        for z, x in zip(
                                            islice(row_heights, 0, None),
                                            islice(row_heights, 1, None),
                                        )
                                    ]
                                ),
                            )
                        )
                    )
                else:
                    self.MT.row_positions = row_heights
            else:
                if verify:
                    self.MT.row_positions = [
                        qmin if z < qmin or not isinstance(z, int) or isinstance(z, bool) else z for z in row_heights
                    ]
                else:
                    self.MT.row_positions = list(accumulate(chain([0], (height for height in row_heights))))

    def verify_row_heights(self, row_heights: list, canvas_positions: bool = False):
        if row_heights[0] != 0 or isinstance(row_heights[0], bool):
            return False
        if not isinstance(row_heights, list):
            return False
        if canvas_positions:
            if any(
                x - z < self.MT.min_row_height or not isinstance(x, int) or isinstance(x, bool)
                for z, x in zip(islice(row_heights, 0, None), islice(row_heights, 1, None))
            ):
                return False
        elif not canvas_positions:
            if any(z < self.MT.min_row_height or not isinstance(z, int) or isinstance(z, bool) for z in row_heights):
                return False
        return True

    def verify_column_widths(self, column_widths: list, canvas_positions: bool = False):
        if column_widths[0] != 0 or isinstance(column_widths[0], bool):
            return False
        if not isinstance(column_widths, list):
            return False
        if canvas_positions:
            if any(
                x - z < self.MT.min_column_width or not isinstance(x, int) or isinstance(x, bool)
                for z, x in zip(islice(column_widths, 0, None), islice(column_widths, 1, None))
            ):
                return False
        elif not canvas_positions:
            if any(
                z < self.MT.min_column_width or not isinstance(z, int) or isinstance(z, bool) for z in column_widths
            ):
                return False
        return True

    def default_row_height(self, height=None):
        if height is not None:
            self.MT.default_row_height = (
                height if isinstance(height, str) else "pixels",
                height if isinstance(height, int) else self.MT.get_lines_cell_height(int(height)),
            )
        return self.MT.default_row_height[1]

    def default_header_height(self, height=None):
        if height is not None:
            self.MT.default_header_height = (
                height if isinstance(height, str) else "pixels",
                height
                if isinstance(height, int)
                else self.MT.get_lines_cell_height(int(height), font=self.MT.header_font),
            )
        return self.MT.default_header_height[1]

    def default_column_width(self, width=None):
        if width is not None:
            if width < self.MT.min_column_width:
                self.MT.default_column_width = self.MT.min_column_width + 20
            else:
                self.MT.default_column_width = int(width)
        return self.MT.default_column_width

    def cut(self, event=None):
        self.MT.ctrl_x(event)

    def copy(self, event=None):
        self.MT.ctrl_c(event)

    def paste(self, event=None):
        self.MT.ctrl_v(event)

    def delete(self, event=None):
        self.MT.delete_key(event)

    def undo(self, event=None):
        self.MT.undo(event)

    def redo(self, event=None):
        self.MT.redo(event)

    def delete_row_position(self, idx: int, deselect_all: bool = False):
        self.MT.del_row_position(idx=idx, deselect_all=deselect_all)

    def delete_column_position(self, idx: int, deselect_all: bool = False):
        self.MT.del_col_position(idx, deselect_all=deselect_all)

    def total_rows(
        self,
        number=None,
        mod_positions: bool = True,
        mod_data: bool = True,
    ):
        if number is None:
            return self.MT.total_data_rows()
        if not isinstance(number, int) or number < 0:
            raise ValueError("number argument must be integer and > 0")
        if number > len(self.MT.data):
            if mod_positions:
                height = self.MT.get_lines_cell_height(int(self.MT.default_row_height[0]))
                for r in range(number - len(self.MT.data)):
                    self.MT.insert_row_position("end", height)
        elif number < len(self.MT.data):
            if not self.MT.all_rows_displayed:
                self.MT.display_rows(enable=False, reset_row_positions=False, deselect_all=True)
            self.MT.row_positions[number + 1 :] = []
        if mod_data:
            self.MT.data_dimensions(total_rows=number)

    def total_columns(
        self,
        number=None,
        mod_positions: bool = True,
        mod_data: bool = True,
    ):
        total_cols = self.MT.total_data_cols()
        if number is None:
            return total_cols
        if not isinstance(number, int) or number < 0:
            raise ValueError("number argument must be integer and > 0")
        if number > total_cols:
            if mod_positions:
                width = self.MT.default_column_width
                for c in range(number - total_cols):
                    self.MT.insert_col_position("end", width)
        elif number < total_cols:
            if not self.MT.all_columns_displayed:
                self.MT.display_columns(enable=False, reset_col_positions=False, deselect_all=True)
            self.MT.col_positions[number + 1 :] = []
        if mod_data:
            self.MT.data_dimensions(total_columns=number)

    def sheet_display_dimensions(
        self,
        total_rows=None,
        total_columns=None,
    ):
        if total_rows is None and total_columns is None:
            return len(self.MT.row_positions) - 1, len(self.MT.col_positions) - 1
        if total_rows is not None:
            height = self.MT.get_lines_cell_height(int(self.MT.default_row_height[0]))
            self.MT.row_positions = list(accumulate(chain([0], (height for row in range(total_rows)))))
        if total_columns is not None:
            width = self.MT.default_column_width
            self.MT.col_positions = list(accumulate(chain([0], (width for column in range(total_columns)))))

    def set_sheet_data_and_display_dimensions(
        self,
        total_rows=None,
        total_columns=None,
    ):
        self.sheet_display_dimensions(total_rows=total_rows, total_columns=total_columns)
        self.MT.data_dimensions(total_rows=total_rows, total_columns=total_columns)

    def delete_row(
        self,
        idx: int = 0,
        index_type: str = "displayed",
        undo=undo,
        redraw: bool = True,
    ):
        self.delete_rows(
            rows=idx,
            index_type=index_type,
            undo=undo,
            redraw=redraw,
        )

    def delete_rows(
        self,
        rows: int | Iterator,
        index_type: str = "displayed",
        undo: bool = False,
        redraw: bool = True,
    ):
        rows = [rows] if isinstance(rows, int) else sorted(rows)
        event_data = event_dict(
            name="delete_rows",
            sheet=self.name,
            boxes=self.MT.get_boxes(),
            selected=self.MT.currently_selected(),
        )
        if "displayed" in index_type.lower():
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
            self.MT.undo_stack.append(ev_stack_dict(event_data))
        self.MT.deselect("all", redraw=False)
        self.set_refresh_timer(redraw)
        return event_data

    def delete_column(
        self,
        idx: int = 0,
        index_type: str = "displayed",
        undo: bool = False,
        redraw: bool = True,
    ):
        self.delete_columns(
            columns=idx,
            index_type=index_type,
            undo=undo,
            redraw=redraw,
        )

    def delete_columns(
        self,
        columns: int | Iterator,
        index_type: str = "displayed",
        undo: bool = False,
        redraw: bool = True,
    ):
        columns = [columns] if isinstance(columns, int) else sorted(columns)
        event_data = event_dict(
            name="delete_columns",
            sheet=self.name,
            boxes=self.MT.get_boxes(),
            selected=self.MT.currently_selected(),
        )
        if "displayed" in index_type.lower():
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
            self.MT.undo_stack.append(ev_stack_dict(event_data))
        self.MT.deselect("all", redraw=False)
        self.set_refresh_timer(redraw)
        return event_data

    def insert_column_position(
        self,
        idx="end",
        width=None,
        deselect_all: bool = False,
        redraw: bool = False,
    ):
        self.MT.insert_col_position(idx=idx, width=width, deselect_all=deselect_all)
        self.set_refresh_timer(redraw)

    def insert_column_positions(
        self,
        idx="end",
        widths=None,
        deselect_all: bool = False,
        redraw: bool = False,
    ):
        self.MT.insert_col_positions(idx=idx, widths=widths, deselect_all=deselect_all)
        self.set_refresh_timer(redraw)

    def insert_row_position(
        self,
        idx="end",
        height=None,
        deselect_all: bool = False,
        redraw: bool = False,
    ):
        self.MT.insert_row_position(idx=idx, height=height, deselect_all=deselect_all)
        self.set_refresh_timer(redraw)

    def insert_row_positions(
        self,
        idx="end",
        heights=None,
        deselect_all: bool = False,
        redraw: bool = False,
    ):
        self.MT.insert_row_positions(idx=idx, heights=heights, deselect_all=deselect_all)
        self.set_refresh_timer(redraw)

    def move_row_position(self, row: int, moveto: int):
        self.MT.move_row_position(row, moveto)

    def move_column_position(self, column: int, moveto: int):
        self.MT.move_col_position(column, moveto)

    def move_columns_using_mapping(
        self,
        data_new_idxs: dict,
        disp_new_idxs: None | dict = None,
        move_data: bool = True,
        create_selections: bool = True,
        index_type: str = "displayed",
        undo: bool = False,
        redraw: bool = True,
    ):
        data_idxs, disp_idxs, event_data = self.MT.move_columns_adjust_options_dict(
            data_new_idxs=data_new_idxs,
            data_old_idxs=dict(zip(data_new_idxs.values(), data_new_idxs)),
            totalcols=None,
            disp_new_idxs=disp_new_idxs,
            move_data=move_data,
            create_selections=create_selections,
            index_type=index_type,
        )
        if undo:
            self.MT.undo_stack.append(ev_stack_dict(event_data))
        self.set_refresh_timer(redraw)
        return data_idxs, disp_idxs, event_data

    def move_columns(
        self,
        move_to: int | None = None,
        to_move: list[int] | None = None,
        move_data: bool = True,
        index_type: str = "displayed",
        create_selections: bool = True,
        undo: bool = False,
        redraw: bool = True,
    ):
        data_idxs, disp_idxs, event_data = self.MT.move_columns_adjust_options_dict(
            *self.MT.get_args_for_move_columns(
                move_to=move_to,
                to_move=to_move,
                index_type=index_type,
            ),
            move_data=move_data,
            create_selections=create_selections,
            index_type=index_type,
        )
        if undo:
            self.MT.undo_stack.append(ev_stack_dict(event_data))
        self.set_refresh_timer(redraw)
        return data_idxs, disp_idxs, event_data

    def move_column(self, column: int, moveto: int):
        self.move_columns(moveto, column)

    def move_rows_using_mapping(
        self,
        data_new_idxs: dict,
        disp_new_idxs: None | dict = None,
        move_data: bool = True,
        create_selections: bool = True,
        index_type: str = "displayed",
        undo: bool = False,
        redraw: bool = True,
    ):
        data_idxs, disp_idxs, event_data = self.MT.move_rows_adjust_options_dict(
            data_new_idxs=data_new_idxs,
            data_old_idxs=dict(zip(data_new_idxs.values(), data_new_idxs)),
            totalcols=None,
            disp_new_idxs=disp_new_idxs,
            move_data=move_data,
            create_selections=create_selections,
            index_type=index_type,
        )
        if undo:
            self.MT.undo_stack.append(ev_stack_dict(event_data))
        self.set_refresh_timer(redraw)
        return data_idxs, disp_idxs, event_data

    def move_rows(
        self,
        move_to: int | None = None,
        to_move: list[int] | None = None,
        move_data: bool = True,
        index_type: str = "displayed",
        create_selections: bool = True,
        undo: bool = False,
        redraw: bool = True,
    ):
        data_idxs, disp_idxs, event_data = self.MT.move_rows_adjust_options_dict(
            *self.MT.get_args_for_move_rows(
                move_to=move_to,
                to_move=to_move,
                index_type=index_type,
            ),
            move_data=move_data,
            create_selections=create_selections,
            index_type=index_type,
        )
        if undo:
            self.MT.undo_stack.append(ev_stack_dict(event_data))
        self.set_refresh_timer(redraw)
        return data_idxs, disp_idxs, event_data

    def move_row(self, row: int, moveto: int):
        self.move_rows(moveto, row)

    # works on currently selected box
    def open_cell(self, ignore_existing_editor: bool = True):
        self.MT.open_cell(event=GeneratedMouseEvent(), ignore_existing_editor=ignore_existing_editor)

    def open_header_cell(self, ignore_existing_editor: bool = True):
        self.CH.open_cell(event=GeneratedMouseEvent(), ignore_existing_editor=ignore_existing_editor)

    def open_index_cell(self, ignore_existing_editor: bool = True):
        self.RI.open_cell(event=GeneratedMouseEvent(), ignore_existing_editor=ignore_existing_editor)

    def set_text_editor_value(self, text="", r=None, c=None):
        if self.MT.text_editor is not None and r is None and c is None:
            self.MT.text_editor.set_text(text)
        elif self.MT.text_editor is not None and self.MT.text_editor_loc == (r, c):
            self.MT.text_editor.set_text(text)

    def bind_text_editor_set(self, func, row, column):
        self.MT.bind_text_editor_destroy(func, row, column)

    def destroy_text_editor(self, event=None):
        self.MT.destroy_text_editor(event=event)

    def get_text_editor_widget(self, event=None):
        try:
            return self.MT.text_editor.textedit
        except Exception:
            return None

    def bind_key_text_editor(self, key: str, function):
        self.MT.text_editor_user_bound_keys[key] = function

    def unbind_key_text_editor(self, key: str):
        if key == "all":
            for key in self.MT.text_editor_user_bound_keys:
                try:
                    self.MT.text_editor.textedit.unbind(key)
                except Exception:
                    pass
            self.MT.text_editor_user_bound_keys = {}
        else:
            if key in self.MT.text_editor_user_bound_keys:
                del self.MT.text_editor_user_bound_keys[key]
            try:
                self.MT.text_editor.textedit.unbind(key)
            except Exception:
                pass

    def get_xview(self):
        return self.MT.xview()

    def get_yview(self):
        return self.MT.yview()

    def set_xview(self, position, option="moveto"):
        self.MT.set_xviews(option, position)

    def set_yview(self, position, option="moveto"):
        self.MT.set_yviews(option, position)

    def set_view(self, x_args, y_args):
        self.MT.set_view(x_args, y_args)

    def see(
        self,
        row=0,
        column=0,
        keep_yscroll: bool = False,
        keep_xscroll: bool = False,
        bottom_right_corner: bool = False,
        check_cell_visibility: bool = True,
        redraw: bool = True,
    ):
        self.MT.see(
            row,
            column,
            keep_yscroll,
            keep_xscroll,
            bottom_right_corner,
            check_cell_visibility=check_cell_visibility,
            redraw=False,
        )
        self.set_refresh_timer(redraw)

    def select_row(self, row, redraw: bool = True, run_binding_func: bool = True):
        self.RI.select_row(
            int(row) if not isinstance(row, int) else row,
            redraw=False,
            run_binding_func=run_binding_func,
        )
        self.set_refresh_timer(redraw)

    def select_column(self, column, redraw: bool = True, run_binding_func: bool = True):
        self.CH.select_col(
            int(column) if not isinstance(column, int) else column,
            redraw=False,
            run_binding_func=run_binding_func,
        )
        self.set_refresh_timer(redraw)

    def select_cell(self, row, column, redraw: bool = True, run_binding_func: bool = True):
        self.MT.select_cell(
            int(row) if not isinstance(row, int) else row,
            int(column) if not isinstance(column, int) else column,
            redraw=False,
            run_binding_func=run_binding_func,
        )
        self.set_refresh_timer(redraw)

    def select_all(self, redraw: bool = True, run_binding_func: bool = True):
        self.MT.select_all(redraw=False, run_binding_func=run_binding_func)
        self.set_refresh_timer(redraw)

    def move_down(self):
        self.MT.move_down()

    def add_cell_selection(
        self,
        row,
        column,
        redraw: bool = True,
        run_binding_func: bool = True,
        set_as_current: bool = True,
    ):
        self.MT.add_selection(
            r=row,
            c=column,
            redraw=False,
            run_binding_func=run_binding_func,
            set_as_current=set_as_current,
        )
        self.set_refresh_timer(redraw)

    def add_row_selection(
        self,
        row,
        redraw: bool = True,
        run_binding_func: bool = True,
        set_as_current: bool = True,
    ):
        self.RI.add_selection(
            r=row,
            redraw=False,
            run_binding_func=run_binding_func,
            set_as_current=set_as_current,
        )
        self.set_refresh_timer(redraw)

    def add_column_selection(
        self,
        column,
        redraw: bool = True,
        run_binding_func: bool = True,
        set_as_current: bool = True,
    ):
        self.CH.add_selection(
            c=column,
            redraw=False,
            run_binding_func=run_binding_func,
            set_as_current=set_as_current,
        )
        self.set_refresh_timer(redraw)

    def toggle_select_cell(
        self,
        row,
        column,
        add_selection: bool = True,
        redraw: bool = True,
        run_binding_func: bool = True,
        set_as_current: bool = True,
    ):
        self.MT.toggle_select_cell(
            row=row,
            column=column,
            add_selection=add_selection,
            redraw=False,
            run_binding_func=run_binding_func,
            set_as_current=set_as_current,
        )
        self.set_refresh_timer(redraw)

    def toggle_select_row(
        self,
        row,
        add_selection: bool = True,
        redraw: bool = True,
        run_binding_func: bool = True,
        set_as_current: bool = True,
    ):
        self.RI.toggle_select_row(
            row=row,
            add_selection=add_selection,
            redraw=False,
            run_binding_func=run_binding_func,
            set_as_current=set_as_current,
        )
        self.set_refresh_timer(redraw)

    def toggle_select_column(
        self,
        column,
        add_selection: bool = True,
        redraw: bool = True,
        run_binding_func: bool = True,
        set_as_current: bool = True,
    ):
        self.CH.toggle_select_col(
            column=column,
            add_selection=add_selection,
            redraw=False,
            run_binding_func=run_binding_func,
            set_as_current=set_as_current,
        )
        self.set_refresh_timer(redraw)

    def deselect(
        self,
        row: int | None | str = None,
        column: int | None = None,
        cell: tuple | None = None,
        redraw: bool = True,
    ):
        self.MT.deselect(r=row, c=column, cell=cell, redraw=False)
        self.set_refresh_timer(redraw)

    # (row, column, type_) e.g. (0, 0, "column") as a named tuple
    def get_currently_selected(self):
        return self.MT.currently_selected()

    def set_currently_selected(self, row=None, column=None, **kwargs):
        self.MT.set_currently_selected(
            r=row,
            c=column,
            **kwargs,
        )

    def get_selected_rows(
        self,
        get_cells: bool = False,
        get_cells_as_rows: bool = False,
        return_tuple: bool = False,
    ):
        if return_tuple:
            return tuple(self.MT.get_selected_rows(get_cells=get_cells, get_cells_as_rows=get_cells_as_rows))
        else:
            return self.MT.get_selected_rows(get_cells=get_cells, get_cells_as_rows=get_cells_as_rows)

    def get_selected_columns(
        self,
        get_cells: bool = False,
        get_cells_as_columns: bool = False,
        return_tuple: bool = False,
    ):
        if return_tuple:
            return tuple(self.MT.get_selected_cols(get_cells=get_cells, get_cells_as_cols=get_cells_as_columns))
        else:
            return self.MT.get_selected_cols(get_cells=get_cells, get_cells_as_cols=get_cells_as_columns)

    def get_selected_cells(
        self,
        get_rows: bool = False,
        get_columns: bool = False,
        sort_by_row: bool = False,
        sort_by_column: bool = False,
    ):
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
        else:
            return self.MT.get_selected_cells(get_rows=get_rows, get_cols=get_columns)

    def get_all_selection_boxes(self):
        return self.MT.get_all_selection_boxes()

    def get_all_selection_boxes_with_types(self):
        return self.MT.get_all_selection_boxes_with_types()

    def create_selection_box(
        self,
        r1,
        c1,
        r2,
        c2,
        type_="cells",
    ):
        return self.MT.create_selection_box(r1=r1, c1=c1, r2=r2, c2=c2, type_="columns" if type_ == "cols" else type_)

    def recreate_all_selection_boxes(self):
        self.MT.recreate_all_selection_boxes()

    def cell_visible(self, r, c):
        return self.MT.cell_visible(r, c)

    def cell_completely_visible(self, r, c, seperate_axes: bool = False):
        return self.MT.cell_completely_visible(r, c, seperate_axes)

    def cell_selected(self, r, c):
        return self.MT.cell_selected(r, c)

    def row_selected(self, r):
        return self.MT.row_selected(r)

    def column_selected(self, c):
        return self.MT.col_selected(c)

    def anything_selected(
        self,
        exclude_columns: bool = False,
        exclude_rows: bool = False,
        exclude_cells: bool = False,
    ):
        if self.MT.anything_selected(
            exclude_columns=exclude_columns,
            exclude_rows=exclude_rows,
            exclude_cells=exclude_cells,
        ):
            return True
        return False

    def all_selected(self):
        return self.MT.all_selected()

    def readonly_rows(
        self,
        rows: int | Iterator = [],
        readonly: bool = True,
        redraw: bool = False,
    ) -> None:
        if isinstance(rows, int):
            rows_ = [rows]
        else:
            rows_ = rows
        if not readonly:
            for r in rows_:
                if r in self.MT.row_options and "readonly" in self.MT.row_options[r]:
                    del self.MT.row_options[r]["readonly"]
        else:
            for r in rows_:
                if r not in self.MT.row_options:
                    self.MT.row_options[r] = {}
                self.MT.row_options[r]["readonly"] = True
        self.set_refresh_timer(redraw)

    def readonly_columns(
        self,
        columns: int | Iterator = [],
        readonly: bool = True,
        redraw: bool = False,
    ):
        if isinstance(columns, int):
            cols_ = [columns]
        else:
            cols_ = columns
        if not readonly:
            for c in cols_:
                if c in self.MT.col_options and "readonly" in self.MT.col_options[c]:
                    del self.MT.col_options[c]["readonly"]
        else:
            for c in cols_:
                if c not in self.MT.col_options:
                    self.MT.col_options[c] = {}
                self.MT.col_options[c]["readonly"] = True
        self.set_refresh_timer(redraw)

    def readonly_cells(
        self,
        row=0,
        column=0,
        cells: Iterator[Sequence[int, int]] = [],
        readonly: bool = True,
        redraw: bool = False,
    ) -> None:
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
        self.set_refresh_timer(redraw)

    def readonly_header(
        self,
        columns=[],
        readonly: bool = True,
        redraw: bool = False,
    ) -> None:
        if isinstance(columns, int):
            columns_ = [columns]
        else:
            columns_ = columns
        if not readonly:
            for c in columns_:
                if c in self.CH.cell_options and "readonly" in self.CH.cell_options[c]:
                    del self.CH.cell_options[c]["readonly"]
        else:
            for c in columns_:
                if c not in self.CH.cell_options:
                    self.CH.cell_options[c] = {}
                self.CH.cell_options[c]["readonly"] = True
        self.set_refresh_timer(redraw)

    def readonly_index(
        self,
        rows=[],
        readonly: bool = True,
        redraw: bool = False,
    ) -> None:
        if isinstance(rows, int):
            rows_ = [rows]
        else:
            rows_ = rows
        if not readonly:
            for r in rows_:
                if r in self.RI.cell_options and "readonly" in self.RI.cell_options[r]:
                    del self.RI.cell_options[r]["readonly"]
        else:
            for r in rows_:
                if r not in self.RI.cell_options:
                    self.RI.cell_options[r] = {}
                self.RI.cell_options[r]["readonly"] = True
        self.set_refresh_timer(redraw)

    def dehighlight_all(self, redraw: bool = True):
        for k in self.MT.cell_options:
            if "highlight" in self.MT.cell_options[k]:
                del self.MT.cell_options[k]["highlight"]
        for k in self.MT.row_options:
            if "highlight" in self.MT.row_options[k]:
                del self.MT.row_options[k]["highlight"]
        for k in self.MT.col_options:
            if "highlight" in self.MT.col_options[k]:
                del self.MT.col_options[k]["highlight"]
        for k in self.RI.cell_options:
            if "highlight" in self.RI.cell_options[k]:
                del self.RI.cell_options[k]["highlight"]
        for k in self.CH.cell_options:
            if "highlight" in self.CH.cell_options[k]:
                del self.CH.cell_options[k]["highlight"]
        self.set_refresh_timer(redraw)

    def dehighlight_rows(self, rows=[], redraw: bool = True):
        if isinstance(rows, int):
            rows_ = [rows]
        else:
            rows_ = rows
        if not rows_ or rows_ == "all":
            for r in self.MT.row_options:
                if "highlight" in self.MT.row_options[r]:
                    del self.MT.row_options[r]["highlight"]

            for r in self.RI.cell_options:
                if "highlight" in self.RI.cell_options[r]:
                    del self.RI.cell_options[r]["highlight"]
        else:
            for r in rows_:
                try:
                    del self.MT.row_options[r]["highlight"]
                except Exception:
                    pass
                try:
                    del self.RI.cell_options[r]["highlight"]
                except Exception:
                    pass
        self.set_refresh_timer(redraw)

    def dehighlight_columns(self, columns=[], redraw: bool = True):
        if isinstance(columns, int):
            columns_ = [columns]
        else:
            columns_ = columns
        if not columns_ or columns_ == "all":
            for c in self.MT.col_options:
                if "highlight" in self.MT.col_options[c]:
                    del self.MT.col_options[c]["highlight"]

            for c in self.CH.cell_options:
                if "highlight" in self.CH.cell_options[c]:
                    del self.CH.cell_options[c]["highlight"]
        else:
            for c in columns_:
                try:
                    del self.MT.col_options[c]["highlight"]
                except Exception:
                    pass
                try:
                    del self.CH.cell_options[c]["highlight"]
                except Exception:
                    pass
        self.set_refresh_timer(redraw)

    def highlight_rows(
        self,
        rows=[],
        bg=None,
        fg=None,
        highlight_index: bool = True,
        redraw: bool = True,
        end_of_screen: bool = False,
        overwrite: bool = True,
    ):
        if bg is None and fg is None:
            return
        for r in (rows,) if isinstance(rows, int) else rows:
            if r not in self.MT.row_options:
                self.MT.row_options[r] = {}
            if "highlight" in self.MT.row_options[r] and not overwrite:
                self.MT.row_options[r]["highlight"] = (
                    self.MT.row_options[r]["highlight"][0] if bg is None else bg,
                    self.MT.row_options[r]["highlight"][1] if fg is None else fg,
                    self.MT.row_options[r]["highlight"][2]
                    if self.MT.row_options[r]["highlight"][2] != end_of_screen
                    else end_of_screen,
                )
            else:
                self.MT.row_options[r]["highlight"] = (bg, fg, end_of_screen)
        if highlight_index:
            self.highlight_cells(cells=rows, canvas="index", bg=bg, fg=fg, redraw=False)
        self.set_refresh_timer(redraw)

    def highlight_columns(
        self,
        columns=[],
        bg=None,
        fg=None,
        highlight_header: bool = True,
        redraw: bool = True,
        overwrite: bool = True,
    ):
        if bg is None and fg is None:
            return
        for c in (columns,) if isinstance(columns, int) else columns:
            if c not in self.MT.col_options:
                self.MT.col_options[c] = {}
            if "highlight" in self.MT.col_options[c] and not overwrite:
                self.MT.col_options[c]["highlight"] = (
                    self.MT.col_options[c]["highlight"][0] if bg is None else bg,
                    self.MT.col_options[c]["highlight"][1] if fg is None else fg,
                )
            else:
                self.MT.col_options[c]["highlight"] = (bg, fg)
        if highlight_header:
            self.highlight_cells(cells=columns, canvas="header", bg=bg, fg=fg, redraw=False)
        self.set_refresh_timer(redraw)

    def highlight_cells(
        self,
        row=0,
        column=0,
        cells=[],
        canvas="table",
        bg=None,
        fg=None,
        redraw: bool = True,
        overwrite: bool = True,
    ):
        if bg is None and fg is None:
            return
        if canvas == "table":
            if cells:
                for r_, c_ in cells:
                    if (r_, c_) not in self.MT.cell_options:
                        self.MT.cell_options[(r_, c_)] = {}
                    if "highlight" in self.MT.cell_options[(r_, c_)] and not overwrite:
                        self.MT.cell_options[(r_, c_)]["highlight"] = (
                            self.MT.cell_options[(r_, c_)]["highlight"][0] if bg is None else bg,
                            self.MT.cell_options[(r_, c_)]["highlight"][1] if fg is None else fg,
                        )
                    else:
                        self.MT.cell_options[(r_, c_)]["highlight"] = (bg, fg)
            else:
                if isinstance(row, str) and row.lower() == "all" and isinstance(column, int):
                    riter = range(self.MT.total_data_rows())
                    citer = (column,)
                elif isinstance(column, str) and column.lower() == "all" and isinstance(row, int):
                    riter = (row,)
                    citer = range(self.MT.total_data_cols())
                elif isinstance(row, int) and isinstance(column, int):
                    riter = (row,)
                    citer = (column,)
                for r_ in riter:
                    for c_ in citer:
                        if (r_, c_) not in self.MT.cell_options:
                            self.MT.cell_options[(r_, c_)] = {}
                        if "highlight" in self.MT.cell_options[(r_, c_)] and not overwrite:
                            self.MT.cell_options[(r_, c_)]["highlight"] = (
                                self.MT.cell_options[(r_, c_)]["highlight"][0] if bg is None else bg,
                                self.MT.cell_options[(r_, c_)]["highlight"][1] if fg is None else fg,
                            )
                        else:
                            self.MT.cell_options[(r_, c_)]["highlight"] = (bg, fg)
        elif canvas in ("row_index", "index"):
            if bg is None and fg is None:
                return
            iterable = (
                cells if (cells and not isinstance(cells, int)) else (cells,) if isinstance(cells, int) else (row,)
            )
            for r_ in iterable:
                if r_ not in self.RI.cell_options:
                    self.RI.cell_options[r_] = {}
                if "highlight" in self.RI.cell_options[r_] and not overwrite:
                    self.RI.cell_options[r_]["highlight"] = (
                        self.RI.cell_options[r_]["highlight"][0] if bg is None else bg,
                        self.RI.cell_options[r_]["highlight"][1] if fg is None else fg,
                    )
                else:
                    self.RI.cell_options[r_]["highlight"] = (bg, fg)
        elif canvas == "header":
            if bg is None and fg is None:
                return
            iterable = (
                cells if (cells and not isinstance(cells, int)) else (cells,) if isinstance(cells, int) else (column,)
            )
            for c_ in iterable:
                if c_ not in self.CH.cell_options:
                    self.CH.cell_options[c_] = {}
                if "highlight" in self.CH.cell_options[c_] and not overwrite:
                    self.CH.cell_options[c_]["highlight"] = (
                        self.CH.cell_options[c_]["highlight"][0] if bg is None else bg,
                        self.CH.cell_options[c_]["highlight"][1] if fg is None else fg,
                    )
                else:
                    self.CH.cell_options[c_]["highlight"] = (bg, fg)
        self.set_refresh_timer(redraw)

    def dehighlight_cells(
        self, row=0, column=0, cells=[], canvas: str = "table", all_: bool = False, redraw: bool = True
    ) -> None:
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
                        pass
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
                        pass
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
                        pass
            elif not all_:
                if column in self.CH.cell_options and "highlight" in self.CH.cell_options[column]:
                    del self.CH.cell_options[column]["highlight"]
            elif all_:
                for c in self.CH.cell_options:
                    if "highlight" in self.CH.cell_options[c]:
                        del self.CH.cell_options[c]["highlight"]
        self.set_refresh_timer(redraw)

    def delete_out_of_bounds_options(self) -> None:
        maxc = self.total_columns()
        maxr = self.total_rows()
        self.MT.cell_options = {k: v for k, v in self.MT.cell_options.items() if k[0] < maxr and k[1] < maxc}
        self.RI.cell_options = {k: v for k, v in self.RI.cell_options.items() if k < maxr}
        self.CH.cell_options = {k: v for k, v in self.CH.cell_options.items() if k < maxc}
        self.MT.col_options = {k: v for k, v in self.MT.col_options.items() if k < maxc}
        self.MT.row_options = {k: v for k, v in self.MT.row_options.items() if k < maxr}
        # needs to handle named ranges that are out of bounds

    def reset_all_options(self) -> None:
        self.MT.named_spans = {}
        self.MT.cell_options = {}
        self.RI.cell_options = {}
        self.CH.cell_options = {}
        self.MT.col_options = {}
        self.MT.row_options = {}
        self.MT.options = {}
        self.RI.options = {}
        self.CH.options = {}

    def get_cell_options(self, canvas: str = "table") -> dict:
        if canvas == "table":
            return self.MT.cell_options
        elif canvas == "row_index":
            return self.RI.cell_options
        elif canvas == "header":
            return self.CH.cell_options

    def get_highlighted_cells(self, canvas: str = "table") -> dict:
        if canvas == "table":
            return {k: v["highlight"] for k, v in self.MT.cell_options.items() if "highlight" in v}
        elif canvas == "row_index":
            return {k: v["highlight"] for k, v in self.RI.cell_options.items() if "highlight" in v}
        elif canvas == "header":
            return {k: v["highlight"] for k, v in self.CH.cell_options.items() if "highlight" in v}

    def get_frame_y(self, y: int) -> int:
        return y + self.CH.current_height

    def get_frame_x(self, x: int) -> int:
        return x + self.RI.current_width

    def convert_align(self, align: str) -> str:
        a = align.lower()
        if a in ("c", "center", "centre"):
            return "center"
        elif a in ("w", "west", "left"):
            return "w"
        elif a in ("e", "east", "right"):
            return "e"
        raise ValueError(("Align must be one of the following values: c, " "center, w, west, left, e, east, right"))

    def get_cell_alignments(self) -> dict:
        return {(r, c): v["align"] for (r, c), v in self.MT.cell_options.items() if "align" in v}

    def get_column_alignments(self) -> dict:
        return {c: v["align"] for c, v in self.MT.col_options.items() if "align" in v}

    def get_row_alignments(self) -> dict:
        return {r: v["align"] for r, v in self.MT.row_options.items() if "align" in v}

    def align_rows(
        self,
        rows=[],
        align: str = "global",
        align_index: bool = False,
        redraw: bool = True,
    ) -> None:  # "center", "w", "e" or "global"
        if align == "global" or self.convert_align(align):
            if isinstance(rows, dict):
                for k, v in rows.items():
                    self.MT.align_rows_(rows=k, align=v, align_index=align_index)
            else:
                self.MT.align_rows_(
                    rows=rows,
                    align=align if align == "global" else self.convert_align(align),
                    align_index=align_index,
                )
        self.set_refresh_timer(redraw)

    def align_rows_(
        self,
        rows=[],
        align: str = "global",
        align_index: bool = False,
    ) -> None:  # "center", "w", "e" or "global"
        if isinstance(rows, str) and rows.lower() == "all" and align == "global":
            for r in self.MT.row_options:
                if "align" in self.MT.row_options[r]:
                    del self.MT.row_options[r]["align"]
            if align_index:
                for r in self.RI.cell_options:
                    if r in self.RI.cell_options and "align" in self.RI.cell_options[r]:
                        del self.RI.cell_options[r]["align"]
            return
        if isinstance(rows, int):
            rows_ = [rows]
        elif isinstance(rows, str) and rows.lower() == "all":
            rows_ = (r for r in range(self.MT.total_data_rows()))
        else:
            rows_ = rows
        if align == "global":
            for r in rows_:
                if r in self.MT.row_options and "align" in self.MT.row_options[r]:
                    del self.MT.row_options[r]["align"]
                if align_index and r in self.RI.cell_options and "align" in self.RI.cell_options[r]:
                    del self.RI.cell_options[r]["align"]
        else:
            for r in rows_:
                if r not in self.MT.row_options:
                    self.MT.row_options[r] = {}
                self.MT.row_options[r]["align"] = align
                if align_index:
                    if r not in self.RI.cell_options:
                        self.RI.cell_options[r] = {}
                    self.RI.cell_options[r]["align"] = align

    def align_columns(
        self,
        columns=[],
        align: str = "global",
        align_header: bool = False,
        redraw: bool = True,
    ) -> None:  # "center", "w", "e" or "global"
        if align == "global" or self.convert_align(align):
            if isinstance(columns, dict):
                for k, v in columns.items():
                    self.align_columns_(columns=k, align=v, align_header=align_header)
            else:
                self.align_columns_(
                    columns=columns,
                    align=align if align == "global" else self.convert_align(align),
                    align_header=align_header,
                )
        self.set_refresh_timer(redraw)

    def align_columns_(
        self,
        columns=[],
        align: str = "global",
        align_header: bool = False,
    ) -> None:  # "center", "w", "e" or "global"
        if isinstance(columns, str) and columns.lower() == "all" and align == "global":
            for c in self.MT.col_options:
                if "align" in self.MT.col_options[c]:
                    del self.MT.col_options[c]["align"]
            if align_header:
                for c in self.CH.cell_options:
                    if c in self.CH.cell_options and "align" in self.CH.cell_options[c]:
                        del self.CH.cell_options[c]["align"]
            return
        if isinstance(columns, int):
            cols_ = [columns]
        elif isinstance(columns, str) and columns.lower() == "all":
            cols_ = (c for c in range(self.MT.total_data_cols()))
        else:
            cols_ = columns
        if align == "global":
            for c in cols_:
                if c in self.MT.col_options and "align" in self.MT.col_options[c]:
                    del self.MT.col_options[c]["align"]
                if align_header and c in self.CH.cell_options and "align" in self.CH.cell_options[c]:
                    del self.CH.cell_options[c]["align"]
        else:
            for c in cols_:
                if c not in self.MT.col_options:
                    self.MT.col_options[c] = {}
                self.MT.col_options[c]["align"] = align
                if align_header:
                    if c not in self.CH.cell_options:
                        self.CH.cell_options[c] = {}
                    self.CH.cell_options[c]["align"] = align

    def align_cells(
        self,
        row=0,
        column=0,
        cells=[],
        align: str = "global",
        redraw: bool = True,
    ) -> None:  # "center", "w", "e" or "global"
        if align == "global" or self.convert_align(align):
            if isinstance(cells, dict):
                for (r, c), v in cells.items():
                    self.align_cells_(row=r, column=c, cells=[], align=v)
            else:
                self.align_cells_(
                    row=row,
                    column=column,
                    cells=cells,
                    align=align if align == "global" else self.convert_align(align),
                )
        self.set_refresh_timer(redraw)

    def align_cells_(
        self,
        row=0,
        column=0,
        cells=[],
        align: str = "global",
    ) -> None:  # "center", "w", "e" or "global"
        if isinstance(row, str) and row.lower() == "all" and align == "global":
            for r, c in self.MT.cell_options:
                if "align" in self.MT.cell_options[(r, c)]:
                    del self.MT.cell_options[(r, c)]["align"]
            return
        if align == "global":
            if cells:
                for r, c in cells:
                    if (r, c) in self.MT.cell_options and "align" in self.MT.cell_options[(r, c)]:
                        del self.MT.cell_options[(r, c)]["align"]
            else:
                if (row, column) in self.MT.cell_options and "align" in self.MT.cell_options[(row, column)]:
                    del self.MT.cell_options[(row, column)]["align"]
        else:
            if cells:
                for r, c in cells:
                    if (r, c) not in self.MT.cell_options:
                        self.MT.cell_options[(r, c)] = {}
                    self.MT.cell_options[(r, c)]["align"] = align
            else:
                if (row, column) not in self.MT.cell_options:
                    self.MT.cell_options[(row, column)] = {}
                self.MT.cell_options[(row, column)]["align"] = align

    def align_header(
        self,
        columns=[],
        align: str = "global",
        redraw: bool = True,
    ) -> None:
        if align == "global" or self.convert_align(align):
            if isinstance(columns, dict):
                for k, v in columns.items():
                    self.align_header_(columns=k, align=v)
            else:
                self.align_header_(
                    columns=columns,
                    align=align if align == "global" else self.convert_align(align),
                )
        self.set_refresh_timer(redraw)

    def align_header_(
        self,
        columns=[],
        align: str = "global",
    ) -> None:
        if isinstance(columns, int):
            cols = [columns]
        else:
            cols = columns
        if align == "global":
            for c in cols:
                if c in self.CH.cell_options and "align" in self.CH.cell_options[c]:
                    del self.CH.cell_options[c]["align"]
        else:
            for c in cols:
                if c not in self.CH.cell_options:
                    self.CH.cell_options[c] = {}
                self.CH.cell_options[c]["align"] = align

    def align_index(
        self,
        rows=[],
        align: str = "global",
        redraw: bool = True,
    ) -> None:
        if align == "global" or self.convert_align(align):
            if isinstance(rows, dict):
                for k, v in rows.items():
                    self.align_index_(rows=rows, align=v)
            else:
                self.align_index_(
                    rows=rows,
                    align=align if align == "global" else self.convert_align(align),
                )
        self.set_refresh_timer(redraw)

    def align_index_(
        self,
        rows: int | list = [],
        align: str = "global",
    ) -> None:
        if isinstance(rows, int):
            rows = [rows]
        else:
            rows = rows
        if align == "global":
            for r in rows:
                if r in self.RI.cell_options and "align" in self.RI.cell_options[r]:
                    del self.RI.cell_options[r]["align"]
        else:
            for r in rows:
                if r not in self.RI.cell_options:
                    self.RI.cell_options[r] = {}
                self.RI.cell_options[r]["align"] = align

    def align(
        self,
        align: str = None,
        redraw: bool = True,
    ) -> str | None:
        if align is None:
            return self.MT.align
        elif self.convert_align(align):
            self.MT.align = self.convert_align(align)
        else:
            raise ValueError("Align must be one of the following values: c, center, w, west, e, east")
        self.set_refresh_timer(redraw)

    def header_align(
        self,
        align: str = None,
        redraw: bool = True,
    ) -> str | None:
        if align is None:
            return self.CH.align
        elif self.convert_align(align):
            self.CH.align = self.convert_align(align)
        else:
            raise ValueError("Align must be one of the following values: c, center, w, west, e, east")
        self.set_refresh_timer(redraw)

    def row_index_align(
        self,
        align: str = None,
        redraw: bool = True,
    ) -> str | None:
        if align is None:
            return self.RI.align
        elif self.convert_align(align):
            self.RI.align = self.convert_align(align)
        else:
            raise ValueError("Align must be one of the following values: c, center, w, west, e, east")
        self.set_refresh_timer(redraw)

    def index_align(
        self,
        align: str = None,
        redraw: bool = True,
    ) -> str | None:
        self.row_index_align(align, redraw)

    def font(self, newfont: tuple | None = None, reset_row_positions: bool = True) -> tuple:
        return self.MT.set_table_font(newfont, reset_row_positions=reset_row_positions)

    def header_font(self, newfont: tuple | None = None) -> tuple:
        return self.MT.set_header_font(newfont)

    def set_options(self, redraw: bool = True, **kwargs) -> None:
        if "auto_resize_columns" in kwargs:
            self.MT.auto_resize_columns = kwargs["auto_resize_columns"]
        if "auto_resize_rows" in kwargs:
            self.MT.auto_resize_rows = kwargs["auto_resize_rows"]
        if "to_clipboard_delimiter" in kwargs:
            self.MT.to_clipboard_delimiter = kwargs["to_clipboard_delimiter"]
        if "to_clipboard_quotechar" in kwargs:
            self.MT.to_clipboard_quotechar = kwargs["to_clipboard_quotechar"]
        if "to_clipboard_lineterminator" in kwargs:
            self.MT.to_clipboard_lineterminator = kwargs["to_clipboard_lineterminator"]
        if "from_clipboard_delimiters" in kwargs:
            self.MT.from_clipboard_delimiters = kwargs["from_clipboard_delimiters"]
        if "show_dropdown_borders" in kwargs:
            self.MT.show_dropdown_borders = kwargs["show_dropdown_borders"]
        if "edit_cell_validation" in kwargs:
            self.MT.edit_cell_validation = kwargs["edit_cell_validation"]
        if "show_default_header_for_empty" in kwargs:
            self.CH.show_default_header_for_empty = kwargs["show_default_header_for_empty"]
        if "show_default_index_for_empty" in kwargs:
            self.RI.show_default_index_for_empty = kwargs["show_default_index_for_empty"]
        if "selected_rows_to_end_of_window" in kwargs:
            self.MT.selected_rows_to_end_of_window = kwargs["selected_rows_to_end_of_window"]
        if "horizontal_grid_to_end_of_window" in kwargs:
            self.MT.horizontal_grid_to_end_of_window = kwargs["horizontal_grid_to_end_of_window"]
        if "vertical_grid_to_end_of_window" in kwargs:
            self.MT.vertical_grid_to_end_of_window = kwargs["vertical_grid_to_end_of_window"]
        if "paste_insert_column_limit" in kwargs:
            self.MT.paste_insert_column_limit = kwargs["paste_insert_column_limit"]
        if "paste_insert_row_limit" in kwargs:
            self.MT.paste_insert_row_limit = kwargs["paste_insert_row_limit"]
        if "expand_sheet_if_paste_too_big" in kwargs:
            self.MT.expand_sheet_if_paste_too_big = kwargs["expand_sheet_if_paste_too_big"]
        if "arrow_key_down_right_scroll_page" in kwargs:
            self.MT.arrow_key_down_right_scroll_page = kwargs["arrow_key_down_right_scroll_page"]
        if "enable_edit_cell_auto_resize" in kwargs:
            self.MT.cell_auto_resize_enabled = kwargs["enable_edit_cell_auto_resize"]
        if "header_hidden_columns_expander_bg" in kwargs:
            self.CH.header_hidden_columns_expander_bg = kwargs["header_hidden_columns_expander_bg"]
        if "index_hidden_rows_expander_bg" in kwargs:
            self.RI.index_hidden_rows_expander_bg = kwargs["index_hidden_rows_expander_bg"]
        if "page_up_down_select_row" in kwargs:
            self.MT.page_up_down_select_row = kwargs["page_up_down_select_row"]
        if "display_selected_fg_over_highlights" in kwargs:
            self.MT.display_selected_fg_over_highlights = kwargs["display_selected_fg_over_highlights"]
        if "show_horizontal_grid" in kwargs:
            self.MT.show_horizontal_grid = kwargs["show_horizontal_grid"]
        if "show_vertical_grid" in kwargs:
            self.MT.show_vertical_grid = kwargs["show_vertical_grid"]
        if "empty_horizontal" in kwargs:
            self.MT.empty_horizontal = kwargs["empty_horizontal"]
        if "empty_vertical" in kwargs:
            self.MT.empty_vertical = kwargs["empty_vertical"]
        if "row_height" in kwargs:
            self.MT.default_row_height = (
                kwargs["row_height"] if isinstance(kwargs["row_height"], str) else "pixels",
                kwargs["row_height"]
                if isinstance(kwargs["row_height"], int)
                else self.MT.get_lines_cell_height(int(kwargs["row_height"])),
            )
        if "column_width" in kwargs:
            self.MT.default_column_width = (
                self.MT.min_column_width + 20
                if kwargs["column_width"] < self.MT.min_column_width
                else int(kwargs["column_width"])
            )
        if "header_height" in kwargs:
            self.MT.default_header_height = (
                kwargs["header_height"] if isinstance(kwargs["header_height"], str) else "pixels",
                kwargs["header_height"]
                if isinstance(kwargs["header_height"], int)
                else self.MT.get_lines_cell_height(int(kwargs["header_height"]), font=self.MT.header_font),
            )
        if "row_drag_and_drop_perform" in kwargs:
            self.RI.row_drag_and_drop_perform = kwargs["row_drag_and_drop_perform"]
        if "column_drag_and_drop_perform" in kwargs:
            self.CH.column_drag_and_drop_perform = kwargs["column_drag_and_drop_perform"]
        if "popup_menu_font" in kwargs:
            self.MT.popup_menu_font = kwargs["popup_menu_font"]
        if "popup_menu_fg" in kwargs:
            self.MT.popup_menu_fg = kwargs["popup_menu_fg"]
        if "popup_menu_bg" in kwargs:
            self.MT.popup_menu_bg = kwargs["popup_menu_bg"]
        if "popup_menu_highlight_bg" in kwargs:
            self.MT.popup_menu_highlight_bg = kwargs["popup_menu_highlight_bg"]
        if "popup_menu_highlight_fg" in kwargs:
            self.MT.popup_menu_highlight_fg = kwargs["popup_menu_highlight_fg"]
        if "top_left_fg_highlight" in kwargs:
            self.TL.top_left_fg_highlight = kwargs["top_left_fg_highlight"]
        if "auto_resize_default_row_index" in kwargs:
            self.RI.auto_resize_width = kwargs["auto_resize_default_row_index"]
        if "header_selected_columns_bg" in kwargs:
            self.CH.header_selected_columns_bg = kwargs["header_selected_columns_bg"]
        if "header_selected_columns_fg" in kwargs:
            self.CH.header_selected_columns_fg = kwargs["header_selected_columns_fg"]
        if "index_selected_rows_bg" in kwargs:
            self.RI.index_selected_rows_bg = kwargs["index_selected_rows_bg"]
        if "index_selected_rows_fg" in kwargs:
            self.RI.index_selected_rows_fg = kwargs["index_selected_rows_fg"]
        if "table_selected_rows_border_fg" in kwargs:
            self.MT.table_selected_rows_border_fg = kwargs["table_selected_rows_border_fg"]
        if "table_selected_rows_bg" in kwargs:
            self.MT.table_selected_rows_bg = kwargs["table_selected_rows_bg"]
        if "table_selected_rows_fg" in kwargs:
            self.MT.table_selected_rows_fg = kwargs["table_selected_rows_fg"]
        if "table_selected_columns_border_fg" in kwargs:
            self.MT.table_selected_columns_border_fg = kwargs["table_selected_columns_border_fg"]
        if "table_selected_columns_bg" in kwargs:
            self.MT.table_selected_columns_bg = kwargs["table_selected_columns_bg"]
        if "table_selected_columns_fg" in kwargs:
            self.MT.table_selected_columns_fg = kwargs["table_selected_columns_fg"]
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
        if "font" in kwargs:
            self.MT.set_table_font(kwargs["font"])
        if "header_font" in kwargs:
            self.MT.set_header_font(kwargs["header_font"])
        if "index_font" in kwargs:
            self.MT.set_index_font(kwargs["index_font"])
        if "theme" in kwargs:
            self.change_theme(kwargs["theme"])
        if "show_selected_cells_border" in kwargs:
            self.MT.show_selected_cells_border = kwargs["show_selected_cells_border"]
        if "header_bg" in kwargs:
            self.CH.config(background=kwargs["header_bg"])
            self.CH.header_bg = kwargs["header_bg"]
        if "header_border_fg" in kwargs:
            self.CH.header_border_fg = kwargs["header_border_fg"]
        if "header_grid_fg" in kwargs:
            self.CH.header_grid_fg = kwargs["header_grid_fg"]
        if "header_fg" in kwargs:
            self.CH.header_fg = kwargs["header_fg"]
        if "header_selected_cells_bg" in kwargs:
            self.CH.header_selected_cells_bg = kwargs["header_selected_cells_bg"]
        if "header_selected_cells_fg" in kwargs:
            self.CH.header_selected_cells_fg = kwargs["header_selected_cells_fg"]
        if "index_bg" in kwargs:
            self.RI.config(background=kwargs["index_bg"])
            self.RI.index_bg = kwargs["index_bg"]
        if "index_border_fg" in kwargs:
            self.RI.index_border_fg = kwargs["index_border_fg"]
        if "index_grid_fg" in kwargs:
            self.RI.index_grid_fg = kwargs["index_grid_fg"]
        if "index_fg" in kwargs:
            self.RI.index_fg = kwargs["index_fg"]
        if "index_selected_cells_bg" in kwargs:
            self.RI.index_selected_cells_bg = kwargs["index_selected_cells_bg"]
        if "index_selected_cells_fg" in kwargs:
            self.RI.index_selected_cells_fg = kwargs["index_selected_cells_fg"]
        if "top_left_bg" in kwargs:
            self.TL.config(background=kwargs["top_left_bg"])
            self.TL.top_left_bg = kwargs["top_left_bg"]
            self.TL.itemconfig(self.TL.select_all_box, fill=kwargs["top_left_bg"])
        if "top_left_fg" in kwargs:
            self.TL.top_left_fg = kwargs["top_left_fg"]
            self.TL.itemconfig("rw", fill=kwargs["top_left_fg"])
            self.TL.itemconfig("rh", fill=kwargs["top_left_fg"])
        if "frame_bg" in kwargs:
            self.config(background=kwargs["frame_bg"])
        if "table_bg" in kwargs:
            self.MT.config(background=kwargs["table_bg"])
            self.MT.table_bg = kwargs["table_bg"]
        if "table_grid_fg" in kwargs:
            self.MT.table_grid_fg = kwargs["table_grid_fg"]
        if "table_fg" in kwargs:
            self.MT.table_fg = kwargs["table_fg"]
        if "table_selected_box_cells_fg" in kwargs:
            self.MT.table_selected_box_cells_fg = kwargs["table_selected_box_cells_fg"]
        if "table_selected_box_rows_fg" in kwargs:
            self.MT.table_selected_box_rows_fg = kwargs["table_selected_box_rows_fg"]
        if "table_selected_box_columns_fg" in kwargs:
            self.MT.table_selected_box_columns_fg = kwargs["table_selected_box_columns_fg"]
        if "table_selected_cells_border_fg" in kwargs:
            self.MT.table_selected_cells_border_fg = kwargs["table_selected_cells_border_fg"]
        if "table_selected_cells_bg" in kwargs:
            self.MT.table_selected_cells_bg = kwargs["table_selected_cells_bg"]
        if "table_selected_cells_fg" in kwargs:
            self.MT.table_selected_cells_fg = kwargs["table_selected_cells_fg"]
        if "resizing_line_fg" in kwargs:
            self.CH.resizing_line_fg = kwargs["resizing_line_fg"]
            self.RI.resizing_line_fg = kwargs["resizing_line_fg"]
        if "drag_and_drop_bg" in kwargs:
            self.CH.drag_and_drop_bg = kwargs["drag_and_drop_bg"]
            self.RI.drag_and_drop_bg = kwargs["drag_and_drop_bg"]
        if "outline_thickness" in kwargs:
            self.config(highlightthickness=kwargs["outline_thickness"])
        if "outline_color" in kwargs:
            self.config(
                highlightbackground=kwargs["outline_color"],
                highlightcolor=kwargs["outline_color"],
            )
        self.MT.create_rc_menus()
        self.set_refresh_timer(redraw)

    def change_theme(self, theme: str = "light blue", redraw: bool = True):
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

    def get_header_data(self, c: int, get_displayed: bool = False):
        return self.CH.get_cell_data(datacn=c, get_displayed=get_displayed)

    def get_index_data(self, r: int, get_displayed: bool = False):
        return self.RI.get_cell_data(datarn=r, get_displayed=get_displayed)

    def get_sheet_data(
        self,
        get_displayed: bool = False,
        get_header: bool = False,
        get_index: bool = False,
        get_header_displayed: bool = True,
        get_index_displayed: bool = True,
        only_rows: Iterator | int | None = None,
        only_columns: Iterator | int | None = None,
        **kwargs,
    ) -> list:
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

    def get_cell_data(self, r, c, get_displayed: bool = False, **kwargs):
        return self.MT.get_cell_data(r, c, get_displayed)

    def get_row_data(
        self,
        r,
        get_displayed: bool = False,
        get_index: bool = False,
        get_index_displayed: bool = True,
        only_columns=None,
        **kwargs,
    ):
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
        c,
        get_displayed: bool = False,
        get_header: bool = False,
        get_header_displayed: bool = True,
        only_rows=None,
        **kwargs,
    ):
        if only_rows is not None:
            if isinstance(only_rows, int):
                only_rows = (only_rows,)
            elif not is_iterable(only_rows):
                raise ValueError(tksheet_type_error("only_rows", ["int", "iterable", "None"], only_rows))
        iterable = only_rows if only_rows is not None else range(len(self.MT.data))
        return ([self.get_header_data(c, get_displayed=get_header_displayed)] if get_header else []) + [
            self.MT.get_cell_data(r, c, get_displayed=get_displayed) for r in iterable
        ]

    def yield_sheet_rows(
        self,
        get_displayed: bool = False,
        get_header: bool = False,
        get_index: bool = False,
        get_index_displayed: bool = True,
        get_header_displayed: bool = True,
        only_rows=None,
        only_columns=None,
        **kwargs,
    ):
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
            maxlen = self.MT.total_data_cols()
            iterable = only_columns if only_columns is not None else range(maxlen)
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

    @property
    def data(self):
        return self.MT.data

    def formatted(self, r, c):
        if (r, c) in self.MT.cell_options and "format" in self.MT.cell_options[(r, c)]:
            return True
        return False

    def data_reference(
        self,
        newdataref=None,
        reset_col_positions: bool = True,
        reset_row_positions: bool = True,
        redraw: bool = False,
    ):
        return self.MT.data_reference(newdataref, reset_col_positions, reset_row_positions, redraw)

    def set_header_data(self, value, c=None, redraw: bool = True):
        if c is None:
            if not isinstance(value, int) and not is_iterable(value):
                raise ValueError(("Argument 'value' must be non-string iterable or int, " f"not {type(value)} type."))
            if not isinstance(value, list) and not isinstance(value, int):
                value = list(value)
            self.MT._headers = value
        elif isinstance(c, int):
            self.MT._headers[c] = value
        elif is_iterable(value) and is_iterable(c):
            for c_, v in zip(c, value):
                self.MT._headers[c_] = v
        self.set_refresh_timer(redraw)

    def set_index_data(self, value, r=None, redraw: bool = True):
        if r is None:
            if not isinstance(value, int) and not is_iterable(value):
                raise ValueError(("Argument 'value' must be non-string iterable or int, " f"not {type(value)} type."))
            if not isinstance(value, list) and not isinstance(value, int):
                value = list(value)
            self.MT._row_index = value
        elif isinstance(r, int):
            self.MT._row_index[r] = value
        elif is_iterable(value) and is_iterable(r):
            for r_, v in zip(r, value):
                self.MT._row_index[r_] = v
        self.set_refresh_timer(redraw)

    def __getitem__(self, key: str | int | slice) -> object:
        """
        Get data from sheet

        If getting a single cell:
            Returns that cells value

        If getting more than one cell:
            Returns a list of lists where each sublist is from a single row

        Key can be either:
            [int] - Get whole row at that index
            [str] - Get whole column at that index - "A" is converted to 0
            [int:int] - Get a range of rows up to and including stop index
            [str:str] - Either e.g.
                ["0":"2"] - Rows 0, 1, 2
                ["A":"C"] - Columns 0, 1, 2
                ["A:C"] - Columns 0, 1, 2
                ["A1":"C1"] - Cells (0, 0), (0, 1), (0, 2)
                ["A1:C1"] - Cells (0, 0), (0, 1), (0, 2)
        """
        if isinstance(key, int):
            return self.MT.data[key]

        elif isinstance(key, str):
            # cell
            if ":" in key:
                ...
            # columns
            elif key.isalpha() and (c := alpha2num(key)) is not None:
                return [self.MT.get_cell_data(r, c, get_displayed=False) for r in range(len(self.MT.data))]
            # rows
            elif (r := str_to_int(key)) is not None:
                return [self.MT.get_cell_data(r, c, get_displayed=False) for c in range(len(self.MT.data[r]))]


        elif isinstance(key, slice):
            ...

    def set_data(self, span: tuple, ) -> Sheet:
        ...

    def set_sheet_data(
        self,
        data=None,
        reset_col_positions: bool = True,
        reset_row_positions: bool = True,
        redraw: bool = True,
        verify: bool = False,
        reset_highlights: bool = False,
        keep_formatting: bool = True,
        delete_options: bool = False,
    ):
        if data is None:
            data = [[]]
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

    def set_cell_data(
        self,
        r,
        c,
        value="",
        redraw: bool = True,
        keep_formatting: bool = True,
    ):
        if not keep_formatting:
            self.MT.delete_cell_format(r, c, clear_values=False)
        self.MT.set_cell_data(r, c, value)
        self.set_refresh_timer(redraw)

    def set_row_data(
        self,
        r,
        values=tuple(),
        add_columns: bool = True,
        redraw: bool = True,
        keep_formatting: bool = True,
    ):
        if r >= len(self.MT.data):
            raise Exception("Row number is out of range")
        if not keep_formatting:
            self.MT.delete_row_format(r, clear_values=False)
        maxidx = len(self.MT.data[r]) - 1
        if not values:
            self.MT.data[r][:] = self.MT.get_empty_row_seq(r, len(self.MT.data[r]))
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
        self.set_refresh_timer(redraw)

    def set_column_data(
        self,
        c,
        values=tuple(),
        add_rows: bool = True,
        redraw: bool = True,
        keep_formatting: bool = True,
    ):
        if not keep_formatting:
            self.MT.delete_column_format(c, clear_values=False)
        if add_rows:
            maxidx = len(self.MT.data) - 1
            total_cols = None
            height = self.MT.default_row_height[1]
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
        self.set_refresh_timer(redraw)

    def insert_column(
        self,
        column: list | tuple | None = None,
        idx: str | int = "end",
        width: int | None = None,
        headers: bool = False,
        create_selections: bool = True,
        undo: bool = False,
        redraw: bool = True,
    ):
        self.insert_columns(
            columns=1 if column is None else [column] if isinstance(column, (list, tuple)) else column,
            idx=idx,
            widths=[width] if isinstance(width, int) else width,
            headers=headers,
            create_selections=create_selections,
            undo=undo,
            redraw=redraw,
        )

    def insert_columns(
        self,
        columns: list | tuple | int = 1,
        idx: str | int = "end",
        widths: list | None = None,
        headers: bool = False,
        create_selections: bool = True,
        undo: bool = False,
        redraw: bool = True,
    ):
        old_total = self.MT.equalize_data_row_lengths()
        if isinstance(columns, int):
            if columns < 1:
                raise ValueError(f"columns arg must be greater than 0, not {columns}")
            total_rows = self.MT.total_data_rows()
            start = old_total if idx == "end" else idx
            # should be a list of lists where each list is a column
            data = [
                [self.MT.get_value_for_empty_cell(datarn, datacn, c_ops=idx == "end") for datarn in range(total_rows)]
                for datacn in range(start, start + columns)
            ]
            numcols = columns
        else:
            data = columns
            numcols = len(columns)
        idx = old_total if idx == "end" else idx
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
            event_data=event_dict(
                name="add_columns",
                sheet=self.name,
                boxes=self.MT.get_boxes(),
                selected=self.MT.currently_selected(),
            ),
        )
        if undo:
            self.MT.undo_stack.append(ev_stack_dict(event_data))
        self.set_refresh_timer(redraw)
        return event_data

    def insert_row(
        self,
        row: list | tuple | None = None,
        idx: str | int = "end",
        height: int | None = None,
        row_index: bool = False,
        create_selections: bool = True,
        undo: bool = False,
        redraw: bool = True,
    ):
        self.insert_rows(
            rows=1 if row is None else [row] if isinstance(row, (list, tuple)) else row,
            idx=idx,
            heights=[height] if isinstance(height, int) else height,
            row_index=row_index,
            create_selections=create_selections,
            undo=undo,
            redraw=redraw,
        )

    def insert_rows(
        self,
        rows: list | tuple | int = 1,
        idx: str | int = "end",
        heights: list | tuple | None = None,
        row_index: bool = False,
        create_selections: bool = True,
        undo: bool = False,
        redraw: bool = True,
    ):
        total_cols = None
        idx = len(self.MT.data) if idx == "end" else idx
        if isinstance(rows, int):
            if rows < 1:
                raise ValueError(f"rows arg must be greater than 0, not {rows}")
            total_cols = self.MT.total_data_cols()
            data = [self.MT.get_empty_row_seq(idx + i, total_cols, r_ops=False) for i in range(rows)]
        elif not isinstance(rows, list):
            data = list(rows)
        else:
            data = rows
        try:
            data = [r if isinstance(r, list) else list(r) for r in data]
        except Exception as msg:
            raise ValueError(f"'rows' arg must be int or list of lists. {msg}")
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
            event_data=event_dict(
                name="add_rows",
                sheet=self.name,
                boxes=self.MT.get_boxes(),
                selected=self.MT.currently_selected(),
            ),
        )
        if undo:
            self.MT.undo_stack.append(ev_stack_dict(event_data))
        self.set_refresh_timer(redraw)
        return event_data

    def sheet_data_dimensions(self, total_rows=None, total_columns=None):
        self.MT.data_dimensions(total_rows, total_columns)

    def get_total_rows(self, include_index: bool = False):
        return self.MT.total_data_rows(include_index=include_index)

    def get_total_columns(self, include_header: bool = False):
        return self.MT.total_data_cols(include_header=include_header)

    def equalize_data_row_lengths(self, include_header: bool = False):
        return self.MT.equalize_data_row_lengths(include_header=include_header)

    def display_rows(
        self,
        rows=None,
        all_rows_displayed=None,
        reset_row_positions: bool = True,
        refresh: bool = False,
        redraw: bool = False,
        deselect_all: bool = True,
        **kwargs,
    ):
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

    def display_columns(
        self,
        columns=None,
        all_columns_displayed=None,
        reset_col_positions: bool = True,
        refresh: bool = False,
        redraw: bool = False,
        deselect_all: bool = True,
        **kwargs,
    ):
        if "all_displayed" in kwargs:
            all_columns_displayed = kwargs["all_displayed"]
        res = self.MT.display_columns(
            columns=None if isinstance(columns, str) and columns.lower() == "all" else columns,
            all_columns_displayed=True
            if isinstance(columns, str) and columns.lower() == "all"
            else all_columns_displayed,
            reset_col_positions=reset_col_positions,
            deselect_all=deselect_all,
        )
        if refresh or redraw:
            self.set_refresh_timer(redraw if redraw else refresh)
        return res

    def all_rows_displayed(self, a=None):
        v = bool(self.MT.all_rows_displayed)
        if type(a) == bool:
            self.MT.all_rows_displayed = a
        return v

    def all_columns_displayed(self, a=None):
        v = bool(self.MT.all_columns_displayed)
        if type(a) == bool:
            self.MT.all_columns_displayed = a
        return v

    # uses displayed indexes
    def hide_rows(self, rows=set(), redraw: bool = True, deselect_all: bool = True):
        if isinstance(rows, int):
            _rows = {rows}
        elif isinstance(rows, set):
            _rows = rows
        else:
            _rows = set(rows)
        if not _rows:
            return
        if self.MT.all_rows_displayed:
            _rows = [r for r in range(self.MT.total_data_rows()) if r not in _rows]
        else:
            _rows = [e for r, e in enumerate(self.MT.displayed_rows) if r not in _rows]
        self.display_rows(
            rows=_rows,
            all_rows_displayed=False,
            redraw=redraw,
            deselect_all=deselect_all,
        )

    # uses displayed indexes
    def hide_columns(self, columns=set(), redraw: bool = True, deselect_all: bool = True):
        if isinstance(columns, int):
            _columns = {columns}
        elif isinstance(columns, set):
            _columns = columns
        else:
            _columns = set(columns)
        if not _columns:
            return
        if self.MT.all_columns_displayed:
            _columns = [c for c in range(self.MT.total_data_cols()) if c not in _columns]
        else:
            _columns = [e for c, e in enumerate(self.MT.displayed_columns) if c not in _columns]
        self.display_columns(
            columns=_columns,
            all_columns_displayed=False,
            redraw=redraw,
            deselect_all=deselect_all,
        )

    def show_ctrl_outline(self, canvas="table", start_cell=(0, 0), end_cell=(1, 1)):
        self.MT.show_ctrl_outline(canvas=canvas, start_cell=start_cell, end_cell=end_cell)

    def get_ctrl_x_c_boxes(self):
        return self.MT.get_ctrl_x_c_boxes()

    def get_selected_min_max(
        self,
    ):  # returns (min_y, min_x, max_y, max_x) of any selections including rows/columns
        return self.MT.get_selected_min_max()

    def headers(
        self,
        newheaders=None,
        index=None,
        reset_col_positions: bool = False,
        show_headers_if_not_sheet: bool = True,
        redraw: bool = False,
    ):
        self.set_refresh_timer(redraw)
        return self.MT.headers(
            newheaders,
            index,
            reset_col_positions=reset_col_positions,
            show_headers_if_not_sheet=show_headers_if_not_sheet,
            redraw=False,
        )

    def row_index(
        self,
        newindex=None,
        index=None,
        reset_row_positions: bool = False,
        show_index_if_not_sheet: bool = True,
        redraw: bool = False,
    ):
        self.set_refresh_timer(redraw)
        return self.MT.row_index(
            newindex,
            index,
            reset_row_positions=reset_row_positions,
            show_index_if_not_sheet=show_index_if_not_sheet,
            redraw=False,
        )

    def reset_undos(self):
        self.MT.purge_undo_and_redo_stack()

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

    def redraw(self, redraw_header: bool = True, redraw_row_index: bool = True):
        self.MT.main_table_redraw_grid_and_text(redraw_header=redraw_header, redraw_row_index=redraw_row_index)

    def refresh(self, redraw_header: bool = True, redraw_row_index: bool = True):
        self.MT.main_table_redraw_grid_and_text(redraw_header=redraw_header, redraw_row_index=redraw_row_index)

    def span(
        self,
        cells: Iterator[tuple[int, int]] | str | None = None,
        rows: Iterator[int] | None = None,
        cols: Iterator[int] | None = None,
        name: str | None = None,
    ):
        return Span(
            cells,
            rows,
            cols,
            name,
        )

    def set_named_spans(self, named_spans: None | dict = None) -> Sheet:
        if named_spans is None:
            named_spans = {}
        self.MT.named_spans = named_spans
        return self

    def set_named_span(self, name: str, named_span: dict) -> Sheet:
        self.MT.named_spans[name] = named_span
        return self

    def get_named_spans(self) -> dict:
        return self.MT.named_spans

    def get_named_span(self, name: str) -> dict:
        return self.MT.named_spans[name]

    def create_named_span(
        self,
        from_r: None | int = None,
        from_c: None | int = None,
        upto_r: None | int = None,
        upto_c: None | int = None,
        type_: str = "",
        name: None | str = None,
        table: bool = True,
        header: bool = False,
        index: bool = False,
        **kwargs,
    ) -> str:
        if name is None:
            name = f"{num2alpha(self.named_span_id)}"
            self.named_span_id += 1
        if name in self.MT.named_spans:
            raise ValueError(f"Name: '{name}' already exists, delete before re-use.")
        type_ = type_.lower()
        if from_c is None or upto_c is None:
            for r in range(from_r, upto_r):
                if table:
                    self.create_table_kwargs(r, c=None, type_=type_, **kwargs)
                if index:
                    self.create_index_kwargs(r, type_, **kwargs)
        elif from_r is None or upto_r is None:
            for c in range(from_c, upto_c):
                if table:
                    self.create_table_kwargs(r=None, c=c, type_=type_, **kwargs)
                if header:
                    self.create_header_kwargs(c, type_, **kwargs)
        else:
            for r in range(from_r, upto_r):
                if index:
                    self.create_index_kwargs(r, type_, **kwargs)
                for c in range(from_c, upto_c):
                    if table:
                        self.create_table_kwargs(r, c, type_, **kwargs)
            if header:
                for c in range(from_c, upto_c):
                    self.create_header_kwargs(c, type_, **kwargs)
        self.MT.named_spans[name] = named_span_dict(
            from_r=from_r,
            from_c=from_c,
            upto_r=upto_r,
            upto_c=upto_c,
            type_=type_,
            name=name,
            table=table,
            header=header,
            index=index,
            kwargs=kwargs,
        )
        return name

    def delete_named_span(self, name: str) -> None:
        if name not in self.MT.named_spans:
            return
        r1, c1, r2, c2 = self.MT.named_span_coords(name)
        # check header
        if isinstance(c1, int) and isinstance(c2, int):
            for type_, span in self.MT.named_spans[name].items():
                if span["header"]:
                    for c in range(c1, c2):
                        if c in self.CH.cell_options:
                            if type_ in self.CH.cell_options[c]:
                                del self.CH.cell_options[c][type_]
        if isinstance(r1, int) and isinstance(r2, int):
            # check index
            for type_, span in self.MT.named_spans[name].items():
                if span["index"]:
                    for r in range(r1, r2):
                        if r in self.RI.cell_options:
                            if type_ in self.RI.cell_options[r]:
                                del self.RI.cell_options[r][type_]
            if isinstance(c1, int) and isinstance(c2, int):
                # check table cell options
                for type_, span in self.MT.named_spans[name].items():
                    for r in range(r1, r2):
                        for c in range(c1, c2):
                            if (r, c) in self.MT.cell_options:
                                if type_ in self.MT.cell_options[(r, c)]:
                                    del self.MT.cell_options[(r, c)][type_]
            else:
                # check row options
                for type_, span in self.MT.named_spans[name].items():
                    for r in range(r1, r2):
                        if r in self.MT.row_options:
                            if type_ in self.MT.row_options[r]:
                                del self.MT.row_options[r][type_]
        else:
            # check col options
            for type_, span in self.MT.named_spans[name].items():
                for c in range(c1, c2):
                    if c in self.MT.col_options:
                        if type_ in self.MT.col_options[c]:
                            del self.MT.col_options[c][type_]
        del self.MT.named_spans[name]

    def create_table_kwargs(self, r: None | int, c: None | int, type_: str, **kwargs):
        if type_ in ("format",):
            if c is None:
                self.format_row(r, **kwargs)
            elif r is None:
                self.format_column(c, **kwargs)
            else:
                self.format_cell(r, c, **kwargs)
        elif type_ in ("dropdown",):
            if c is None:
                self.dropdown_row(r, **kwargs)
            elif r is None:
                self.dropdown_column(c, **kwargs)
            else:
                self.create_dropdown(r, c, **kwargs)
        elif type_ in ("checkbox",):
            if c is None:
                self.checkbox_row(r, **kwargs)
            elif r is None:
                self.checkbox_column(c, **kwargs)
            else:
                self.create_checkbox(r, c, **kwargs)
        elif type_ in ("highlight",):
            if c is None:
                self.highlight_rows(r, **kwargs)
            elif r is None:
                self.highlight_columns(c, **kwargs)
            else:
                self.highlight_cells(r, c, **kwargs)
        elif type_ in ("align",):
            if c is None:
                self.align_rows(r, **kwargs)
            elif r is None:
                self.align_columns(c, **kwargs)
            else:
                self.align_cells(r, c, **kwargs)
        elif type_ in ("readonly",):
            if c is None:
                self.readonly_rows(r, **kwargs)
            elif r is None:
                self.readonly_columns(c, **kwargs)
            else:
                self.readonly_cells(r, c, **kwargs)

    def create_header_kwargs(self, c: int, type_: str, **kwargs):
        if type_ in ("format",):
            # cannot format header/index currently
            pass

        elif type_ in ("dropdown",):
            pass

        elif type_ in ("checkbox",):
            pass

        elif type_ in ("highlight",):
            pass

        elif type_ in ("align",):
            pass

        elif type_ in ("readonly",):
            pass

    def create_index_kwargs(self, r: int, type_: str, **kwargs):
        if type_ in ("format",):
            # cannot format header/index currently
            pass

        elif type_ in ("dropdown",):
            pass

        elif type_ in ("checkbox",):
            pass

        elif type_ in ("highlight",):
            pass

        elif type_ in ("align",):
            pass

        elif type_ in ("readonly",):
            pass

    def create_checkbox(self, r=0, c=0, *args, **kwargs):
        _kwargs = get_checkbox_kwargs(*args, **kwargs)
        if isinstance(r, str) and r.lower() == "all" and isinstance(c, int):
            for r_ in range(self.MT.total_data_rows()):
                self.create_checkbox_(datarn=r_, datacn=c, **_kwargs)
        elif isinstance(c, str) and c.lower() == "all" and isinstance(r, int):
            for c_ in range(self.MT.total_data_cols()):
                self.create_checkbox_(datarn=r, datacn=c_, **_kwargs)
        elif isinstance(r, str) and r.lower() == "all" and isinstance(c, str) and c.lower() == "all":
            totalcols = self.MT.total_data_cols()
            for r_ in range(self.MT.total_data_rows()):
                for c_ in range(totalcols):
                    self.create_checkbox_(datarn=r_, datacn=c_, **_kwargs)
        elif isinstance(r, int) and isinstance(c, int):
            self.create_checkbox_(datarn=r, datacn=c, **_kwargs)
        self.set_refresh_timer(_kwargs["redraw"])

    def create_checkbox_(self, datarn: int = 0, datacn: int = 0, **kwargs):
        self.MT.delete_cell_format(datarn, datacn, clear_values=False)
        if (datarn, datacn) in self.MT.cell_options and (
            "dropdown" in self.MT.cell_options[(datarn, datacn)] or "checkbox" in self.MT.cell_options[(datarn, datacn)]
        ):
            self.delete_table_cell_options_dropdown_and_checkbox(datarn, datacn)
        if (datarn, datacn) not in self.MT.cell_options:
            self.MT.cell_options[(datarn, datacn)] = {}
        self.MT.cell_options[(datarn, datacn)]["checkbox"] = get_checkbox_dict(**kwargs)
        self.MT.set_cell_data(datarn, datacn, kwargs["checked"])

    def checkbox_cell(self, r=0, c=0, *args, **kwargs):
        self.create_checkbox(r=r, c=c, **get_checkbox_kwargs(*args, **kwargs))

    def create_header_checkbox(self, c=0, *args, **kwargs):
        _kwargs = get_checkbox_kwargs(*args, **kwargs)
        if isinstance(c, str) and c.lower() == "all":
            for c_ in range(self.MT.total_data_cols()):
                self.create_header_checkbox_(datacn=c_, **_kwargs)
        elif isinstance(c, int):
            self.create_header_checkbox_(datacn=c, **_kwargs)
        elif is_iterable(c):
            for c_ in c:
                self.create_header_checkbox_(datacn=c_, **_kwargs)
        else:
            self.checkbox_header(**_kwargs)
        self.set_refresh_timer(_kwargs["redraw"])

    def create_header_checkbox_(self, datacn=0, **kwargs):
        if datacn in self.CH.cell_options and (
            "dropdown" in self.CH.cell_options[datacn] or "checkbox" in self.CH.cell_options[datacn]
        ):
            self.delete_header_cell_options_dropdown_and_checkbox(datacn)
        if datacn not in self.CH.cell_options:
            self.CH.cell_options[datacn] = {}
        self.CH.cell_options[datacn]["checkbox"] = get_checkbox_dict(**kwargs)
        self.CH.set_cell_data(datacn=datacn, value=kwargs["checked"])

    def checkbox_header(self, **kwargs):
        self.CH.destroy_opened_dropdown_window()
        if "dropdown" in self.CH.options or "checkbox" in self.CH.options:
            self.delete_header_options_dropdown_and_checkbox()
        if "checkbox" not in self.CH.options:
            self.CH.options["checkbox"] = {}
        self.CH.options["checkbox"] = get_checkbox_dict(**kwargs)
        total_cols = self.MT.total_data_cols()
        if isinstance(self.MT._headers, int):
            for datacn in range(total_cols):
                self.MT.set_cell_data(datarn=self.MT._headers, datacn=datacn, value=kwargs["checked"])
        else:
            for datacn in range(total_cols):
                self.CH.set_cell_data(datacn=datacn, value=kwargs["checked"])

    def create_index_checkbox(self, r=0, *args, **kwargs):
        _kwargs = get_checkbox_kwargs(*args, **kwargs)
        if isinstance(r, str) and r.lower() == "all":
            for r_ in range(self.MT.total_data_rows()):
                self.create_index_checkbox_(datarn=r_, **_kwargs)
        elif isinstance(r, int):
            self.create_index_checkbox_(datarn=r, **_kwargs)
        elif is_iterable(r):
            for r_ in r:
                self.create_index_checkbox_(datarn=r_, **_kwargs)
        else:
            self.checkbox_index(**_kwargs)
        self.set_refresh_timer(_kwargs["redraw"])

    def create_index_checkbox_(self, datarn=0, **kwargs):
        if datarn in self.RI.cell_options and (
            "dropdown" in self.RI.cell_options[datarn] or "checkbox" in self.RI.cell_options[datarn]
        ):
            self.delete_index_cell_options_dropdown_and_checkbox(datarn)
        if datarn not in self.RI.cell_options:
            self.RI.cell_options[datarn] = {}
        self.RI.cell_options[datarn]["checkbox"] = get_checkbox_dict(**kwargs)
        self.RI.set_cell_data(datarn=datarn, value=kwargs["checked"])

    def checkbox_index(self, **kwargs):
        self.RI.destroy_opened_dropdown_window()
        if "dropdown" in self.RI.options or "checkbox" in self.RI.options:
            self.delete_index_options_dropdown_and_checkbox()
        if "checkbox" not in self.RI.options:
            self.RI.options["checkbox"] = {}
        self.RI.options["checkbox"] = get_checkbox_dict(**kwargs)
        total_rows = self.MT.total_data_rows()
        if isinstance(self.MT._row_index, int):
            for datarn in range(total_rows):
                self.MT.set_cell_data(datarn=datarn, datacn=self.MT._row_index, value=kwargs["checked"])
        else:
            for datarn in range(total_rows):
                self.RI.set_cell_data(datarn=datarn, value=kwargs["checked"])

    def checkbox_row(self, r=0, *args, **kwargs):
        _kwargs = get_checkbox_kwargs(*args, **kwargs)
        if isinstance(r, str) and r.lower() == "all":
            for r_ in range(self.MT.total_data_rows()):
                self.checkbox_row_(datarn=r_, **_kwargs)
        elif isinstance(r, int):
            self.checkbox_row_(datarn=r, **_kwargs)
        elif is_iterable(r):
            for r_ in r:
                self.checkbox_row_(datarn=r_, **_kwargs)
        self.set_refresh_timer(_kwargs["redraw"])

    def checkbox_row_(self, datarn: int = 0, **kwargs):
        self.MT.delete_row_format(datarn, clear_values=False)
        if datarn in self.MT.row_options and (
            "dropdown" in self.MT.row_options[datarn] or "checkbox" in self.MT.row_options[datarn]
        ):
            self.delete_row_options_dropdown_and_checkbox(datarn)
        if datarn not in self.MT.row_options:
            self.MT.row_options[datarn] = {}
        self.MT.row_options[datarn]["checkbox"] = get_checkbox_dict(**kwargs)
        for datacn in range(self.MT.total_data_cols()):
            self.MT.set_cell_data(datarn, datacn, kwargs["checked"])

    def checkbox_column(self, c=0, *args, **kwargs):
        _kwargs = get_checkbox_kwargs(*args, **kwargs)
        if isinstance(c, str) and c.lower() == "all":
            for c in range(self.MT.total_data_cols()):
                self.checkbox_column_(datacn=c, **_kwargs)
        elif isinstance(c, int):
            self.checkbox_column_(datacn=c, **_kwargs)
        elif is_iterable(c):
            for c_ in c:
                self.checkbox_column_(datacn=c_, **_kwargs)
        self.set_refresh_timer(_kwargs["redraw"])

    def checkbox_column_(self, datacn: int = 0, **kwargs):
        self.MT.delete_column_format(datacn, clear_values=False)
        if datacn in self.MT.col_options and (
            "dropdown" in self.MT.col_options[datacn] or "checkbox" in self.MT.col_options[datacn]
        ):
            self.delete_column_options_dropdown_and_checkbox(datacn)
        if datacn not in self.MT.col_options:
            self.MT.col_options[datacn] = {}
        self.MT.col_options[datacn]["checkbox"] = get_checkbox_dict(**kwargs)
        for datarn in range(self.MT.total_data_rows()):
            self.MT.set_cell_data(datarn, datacn, kwargs["checked"])

    def checkbox_sheet(self, *args, **kwargs):
        self.checkbox_sheet_(**get_checkbox_kwargs(*args, **kwargs))

    def checkbox_sheet_(self, **kwargs):
        self.MT.delete_sheet_format(clear_values=False)
        if "dropdown" in self.MT.options or "checkbox" in self.MT.options:
            self.delete_table_options_dropdown_and_checkbox()
        self.MT.options["checkbox"] = get_checkbox_dict(**kwargs)
        total_cols = self.MT.total_data_cols()
        for datarn in range(self.MT.total_data_rows()):
            for datacn in range(total_cols):
                self.MT.set_cell_data(datarn, datacn, kwargs["checked"])

    def delete_checkbox(self, r=0, c=0):
        if isinstance(r, str) and r.lower() == "all" and isinstance(c, int):
            for r_, c_ in self.MT.cell_options:
                if "checkbox" in self.MT.cell_options[(r_, c)]:
                    self.delete_table_cell_options_checkbox(r_, c)
        elif isinstance(c, str) and c.lower() == "all" and isinstance(r, int):
            for r_, c_ in self.MT.cell_options:
                if "checkbox" in self.MT.cell_options[(r, c_)]:
                    self.delete_table_cell_options_checkbox(r, c_)
        elif isinstance(r, str) and r.lower() == "all" and isinstance(c, str) and c.lower() == "all":
            for r_, c_ in self.MT.cell_options:
                if "checkbox" in self.MT.cell_options[(r_, c_)]:
                    self.delete_table_cell_options_checkbox(r_, c_)
        elif isinstance(r, int) and isinstance(c, int):
            self.delete_table_cell_options_checkbox(r, c)

    def delete_cell_checkbox(self, r=0, c=0):
        self.delete_checkbox(r, c)

    def delete_row_checkbox(self, r=0):
        if isinstance(r, str) and r.lower() == "all":
            for r_ in self.MT.row_options:
                self.delete_row_options_checkbox(r_)
        elif isinstance(r, int):
            self.delete_row_options_checkbox(r)
        elif is_iterable(r):
            for r_ in r:
                self.delete_row_options_checkbox(r_)

    def delete_column_checkbox(self, c=0):
        if isinstance(c, str) and c.lower() == "all":
            for c_ in self.MT.col_options:
                self.delete_column_options_checkbox(c_)
        elif isinstance(c, int):
            self.delete_column_options_checkbox(c)
        elif is_iterable(c):
            for c_ in c:
                self.delete_column_options_checkbox(c_)

    def delete_sheet_checkbox(self):
        self.delete_table_options_checkbox()

    def delete_header_checkbox(self, c=0):
        if isinstance(c, str) and c.lower() == "all":
            for c_ in self.CH.cell_options:
                if "checkbox" in self.CH.cell_options[c_]:
                    self.delete_header_cell_options_checkbox(c_)
        if isinstance(c, int):
            self.delete_header_cell_options_checkbox(c)
        else:
            self.delete_header_options_checkbox()

    def delete_index_checkbox(self, r=0):
        if isinstance(r, str) and r.lower() == "all":
            for r_ in self.RI.cell_options:
                if "checkbox" in self.RI.cell_options[r_]:
                    self.delete_index_cell_options_checkbox(r_)
        if isinstance(r, int):
            self.delete_index_cell_options_checkbox(r)
        else:
            self.delete_index_options_checkbox()

    def click_checkbox(self, r, c, checked=None):
        kwargs = self.MT.get_cell_kwargs(r, c, key="checkbox")
        if kwargs:
            if not type(self.MT.data[r][c]) == bool:
                if checked is None:
                    self.MT.data[r][c] = False
                else:
                    self.MT.data[r][c] = bool(checked)
            else:
                self.MT.data[r][c] = not self.MT.data[r][c]

    def click_header_checkbox(self, c, checked=None):
        kwargs = self.CH.get_cell_kwargs(c, key="checkbox")
        if kwargs:
            if not type(self.MT._headers[c]) == bool:
                if checked is None:
                    self.MT._headers[c] = False
                else:
                    self.MT._headers[c] = bool(checked)
            else:
                self.MT._headers[c] = not self.MT._headers[c]

    def click_index_checkbox(self, r, checked=None):
        kwargs = self.RI.get_cell_kwargs(r, key="checkbox")
        if kwargs:
            if not type(self.MT._row_index[r]) == bool:
                if checked is None:
                    self.MT._row_index[r] = False
                else:
                    self.MT._row_index[r] = bool(checked)
            else:
                self.MT._row_index[r] = not self.MT._row_index[r]

    def get_checkboxes(self):
        d = {
            **{k: v["checkbox"] for k, v in self.MT.cell_options.items() if "checkbox" in v},
            **{k: v["checkbox"] for k, v in self.MT.row_options.items() if "checkbox" in v},
            **{k: v["checkbox"] for k, v in self.MT.col_options.items() if "checkbox" in v},
        }
        if "checkbox" in self.MT.options:
            return {**d, "checkbox": self.MT.options["checkbox"]}
        return d

    def get_header_checkboxes(self):
        d = {k: v["checkbox"] for k, v in self.CH.cell_options.items() if "checkbox" in v}
        if "checkbox" in self.CH.options:
            return {**d, "checkbox": self.CH.options["checkbox"]}
        return d

    def get_index_checkboxes(self):
        d = {k: v["checkbox"] for k, v in self.RI.cell_options.items() if "checkbox" in v}
        if "checkbox" in self.RI.options:
            return {**d, "checkbox": self.RI.options["checkbox"]}
        return d

    def checkbox(self, r, c, checked=None, state=None, check_function="", text=None):
        if type(checked) == bool:
            self.set_cell_data(r, c, checked)
        kwargs = self.MT.get_cell_kwargs(r, c, key="checkbox")
        if check_function != "":
            kwargs["check_function"] = check_function
        if state and state.lower() in ("normal", "disabled"):
            kwargs["state"] = state
        if text is not None:
            kwargs["text"] = text
        return {**kwargs, "checked": self.MT.data[r][c]}

    def header_checkbox(self, c, checked=None, state=None, check_function="", text=None):
        if type(checked) == bool:
            self.headers(newheaders=checked, index=c)
        kwargs = self.CH.get_cell_kwargs(c, key="checkbox")
        if kwargs:
            if check_function != "":
                kwargs["check_function"] = check_function
            if state and state.lower() in ("normal", "disabled"):
                kwargs["state"] = state
            if text is not None:
                kwargs["text"] = text
            return {**kwargs, "checked": self.MT._headers[c]}

    def index_checkbox(self, r, checked=None, state=None, check_function="", text=None):
        if type(checked) == bool:
            self.row_index(newindex=checked, index=r)
        kwargs = self.RI.get_cell_kwargs(r, key="checkbox")
        if kwargs:
            if check_function != "":
                kwargs["check_function"] = check_function
            if state and state.lower() in ("normal", "disabled"):
                kwargs["state"] = state
            if text is not None:
                kwargs["text"] = text
            return {**kwargs, "checked": self.MT._row_index[r]}

    def create_dropdown(self, r=0, c=0, *args, **kwargs):
        _kwargs = get_dropdown_kwargs(*args, **kwargs)
        if isinstance(r, str) and r.lower() == "all" and isinstance(c, int):
            for r_ in range(self.MT.total_data_rows()):
                self.create_dropdown_(datarn=r_, datacn=c, **_kwargs)
        elif isinstance(c, str) and c.lower() == "all" and isinstance(r, int):
            for c_ in range(self.MT.total_data_cols()):
                self.create_dropdown_(datarn=r, datacn=c_, **_kwargs)
        elif isinstance(r, str) and r.lower() == "all" and isinstance(c, str) and c.lower() == "all":
            totalcols = self.MT.total_data_cols()
            for r_ in range(self.MT.total_data_rows()):
                for c_ in range(totalcols):
                    self.create_dropdown_(datarn=r_, datacn=c_, **_kwargs)
        elif isinstance(r, int) and isinstance(c, int):
            self.create_dropdown_(datarn=r, datacn=c, **_kwargs)
        self.set_refresh_timer(_kwargs["redraw"])

    def create_dropdown_(self, datarn: int = 0, datacn: int = 0, **kwargs):
        if (datarn, datacn) in self.MT.cell_options and (
            "dropdown" in self.MT.cell_options[(datarn, datacn)] or "checkbox" in self.MT.cell_options[(datarn, datacn)]
        ):
            self.delete_table_cell_options_dropdown_and_checkbox(datarn, datacn)
        if (datarn, datacn) not in self.MT.cell_options:
            self.MT.cell_options[(datarn, datacn)] = {}
        self.MT.cell_options[(datarn, datacn)]["dropdown"] = get_dropdown_dict(**kwargs)
        self.MT.set_cell_data(
            datarn,
            datacn,
            kwargs["set_value"] if kwargs["set_value"] is not None else kwargs["values"][0] if kwargs["values"] else "",
        )

    def dropdown_cell(self, r=0, c=0, *args, **kwargs):
        self.create_dropdown(r=r, c=c, **get_dropdown_kwargs(*args, **kwargs))

    def dropdown_row(self, r=0, *args, **kwargs):
        _kwargs = get_dropdown_kwargs(*args, **kwargs)
        if isinstance(r, str) and r.lower() == "all":
            for r_ in range(self.MT.total_data_rows()):
                self.dropdown_row_(datarn=r_, **_kwargs)
        elif isinstance(r, int):
            self.dropdown_row_(datarn=r, **_kwargs)
        elif is_iterable(r):
            for r_ in r:
                self.dropdown_row_(datarn=r_, **_kwargs)
        self.set_refresh_timer(_kwargs["redraw"])

    def dropdown_row_(self, datarn: int = 0, **kwargs):
        if datarn in self.MT.row_options and (
            "dropdown" in self.MT.row_options[datarn] or "checkbox" in self.MT.row_options[datarn]
        ):
            self.delete_row_options_dropdown_and_checkbox(datarn)
        if datarn not in self.MT.row_options:
            self.MT.row_options[datarn] = {}
        self.MT.row_options[datarn]["dropdown"] = get_dropdown_dict(**kwargs)
        value = (
            kwargs["set_value"] if kwargs["set_value"] is not None else kwargs["values"][0] if kwargs["values"] else ""
        )
        for datacn in range(self.MT.total_data_cols()):
            self.MT.set_cell_data(datarn, datacn, value)

    def dropdown_column(self, c=0, *args, **kwargs):
        _kwargs = get_dropdown_kwargs(*args, **kwargs)
        if isinstance(c, str) and c.lower() == "all":
            for c_ in range(self.MT.total_data_cols()):
                self.dropdown_column_(datacn=c_, **_kwargs)
        elif isinstance(c, int):
            self.dropdown_column_(datacn=c, **_kwargs)
        elif is_iterable(c):
            for c_ in c:
                self.dropdown_column_(datacn=c_, **_kwargs)
        self.set_refresh_timer(_kwargs["redraw"])

    def dropdown_column_(self, datacn: int = 0, **kwargs):
        if datacn in self.MT.col_options and (
            "dropdown" in self.MT.col_options[datacn] or "checkbox" in self.MT.col_options[datacn]
        ):
            self.delete_column_options_dropdown_and_checkbox(datacn)
        if datacn not in self.MT.col_options:
            self.MT.col_options[datacn] = {}
        self.MT.col_options[datacn]["dropdown"] = get_dropdown_dict(**kwargs)
        value = (
            kwargs["set_value"] if kwargs["set_value"] is not None else kwargs["values"][0] if kwargs["values"] else ""
        )
        for datarn in range(self.MT.total_data_rows()):
            self.MT.set_cell_data(datarn, datacn, value)

    def dropdown_sheet(self, *args, **kwargs):
        _kwargs = get_dropdown_kwargs(*args, **kwargs)
        self.dropdown_sheet_(**_kwargs)
        self.set_refresh_timer(_kwargs["redraw"])

    def dropdown_sheet_(self, **kwargs):
        if "dropdown" in self.MT.options or "checkbox" in self.MT.options:
            self.delete_table_options_dropdown_and_checkbox()
        self.MT.options["dropdown"] = get_dropdown_dict(**kwargs)
        value = (
            kwargs["set_value"] if kwargs["set_value"] is not None else kwargs["values"][0] if kwargs["values"] else ""
        )
        total_cols = self.MT.total_data_cols()
        for datarn in range(self.MT.total_data_rows()):
            for datacn in range(total_cols):
                self.MT.set_cell_data(datarn, datacn, value)

    def create_header_dropdown(self, c=0, *args, **kwargs):
        _kwargs = get_dropdown_kwargs(*args, **kwargs)
        if isinstance(c, str) and c.lower() == "all":
            for c_ in range(self.MT.total_data_cols()):
                self.create_header_dropdown_(datacn=c_, **_kwargs)
        elif isinstance(c, int):
            self.create_header_dropdown_(datacn=c, **_kwargs)
        elif is_iterable(c):
            for c_ in c:
                self.create_header_dropdown_(datacn=c_, **_kwargs)
        elif c is None:
            self.dropdown_header_(**_kwargs)
        self.set_refresh_timer(_kwargs["redraw"])

    def dropdown_header(self, **kwargs):
        self.CH.destroy_opened_dropdown_window()
        if "dropdown" in self.CH.options or "checkbox" in self.CH.options:
            self.delete_header_options_dropdown_and_checkbox()
        if "dropdown" not in self.CH.options:
            self.CH.options["dropdown"] = {}
        self.CH.options["dropdown"] = get_dropdown_dict(**kwargs)
        total_cols = self.MT.total_data_cols()
        value = (
            kwargs["set_value"] if kwargs["set_value"] is not None else kwargs["values"][0] if kwargs["values"] else ""
        )
        if isinstance(self.MT._headers, int):
            for datacn in range(total_cols):
                self.MT.set_cell_data(datarn=self.MT._headers, datacn=datacn, value=value)
        else:
            for datacn in range(total_cols):
                self.CH.set_cell_data(datacn=datacn, value=value)

    def create_header_dropdown_(self, datacn=0, **kwargs):
        if datacn in self.CH.cell_options and (
            "dropdown" in self.CH.cell_options[datacn] or "checkbox" in self.CH.cell_options[datacn]
        ):
            self.delete_header_cell_options_dropdown_and_checkbox(datacn)
        if datacn not in self.CH.cell_options:
            self.CH.cell_options[datacn] = {}
        self.CH.cell_options[datacn]["dropdown"] = get_dropdown_dict(**kwargs)
        self.CH.set_cell_data(
            datacn=datacn,
            value=kwargs["set_value"]
            if kwargs["set_value"] is not None
            else kwargs["values"][0]
            if kwargs["values"]
            else "",
        )

    def create_index_dropdown(self, r=0, *args, **kwargs):
        _kwargs = get_dropdown_kwargs(*args, **kwargs)
        if isinstance(r, str) and r.lower() == "all":
            for r_ in range(self.MT.total_data_rows()):
                self.create_index_dropdown_(datarn=r_, **_kwargs)
        elif isinstance(r, int):
            self.create_index_dropdown_(datarn=r, **_kwargs)
        elif is_iterable(r):
            for r_ in r:
                self.create_index_dropdown_(datarn=r_, **_kwargs)
        elif r is None:
            self.dropdown_index(**_kwargs)
        self.set_refresh_timer(_kwargs["redraw"])

    def dropdown_index(self, **kwargs):
        self.RI.destroy_opened_dropdown_window()
        if "dropdown" in self.RI.options or "checkbox" in self.RI.options:
            self.delete_index_options_dropdown_and_checkbox()
        if "dropdown" not in self.RI.options:
            self.RI.options["dropdown"] = {}
        self.RI.options["dropdown"] = get_dropdown_dict(**kwargs)
        total_rows = self.MT.total_data_rows()
        value = (
            kwargs["set_value"] if kwargs["set_value"] is not None else kwargs["values"][0] if kwargs["values"] else ""
        )
        if isinstance(self.MT._row_index, int):
            for datarn in range(total_rows):
                self.MT.set_cell_data(datarn=datarn, datacn=self.MT._row_index, value=value)
        else:
            for datarn in range(total_rows):
                self.RI.set_cell_data(datarn=datarn, value=value)

    def create_index_dropdown_(self, datarn, **kwargs):
        if datarn in self.RI.cell_options and (
            "dropdown" in self.RI.cell_options[datarn] or "checkbox" in self.RI.cell_options[datarn]
        ):
            self.delete_index_cell_options_dropdown_and_checkbox(datarn)
        if datarn not in self.RI.cell_options:
            self.RI.cell_options[datarn] = {}
        self.RI.cell_options[datarn]["dropdown"] = get_dropdown_dict(**kwargs)
        self.RI.set_cell_data(
            datarn=datarn,
            value=kwargs["set_value"]
            if kwargs["set_value"] is not None
            else kwargs["values"][0]
            if kwargs["values"]
            else "",
        )

    def delete_dropdown(self, r=0, c=0):
        if isinstance(r, str) and r.lower() == "all" and isinstance(c, int):
            for r_, c_ in self.MT.cell_options:
                if "dropdown" in self.MT.cell_options[(r_, c)]:
                    self.delete_table_cell_options_dropdown(r_, c)
        elif isinstance(c, str) and c.lower() == "all" and isinstance(r, int):
            for r_, c_ in self.MT.cell_options:
                if "dropdown" in self.MT.cell_options[(r, c_)]:
                    self.delete_table_cell_options_dropdown(r, c_)
        elif isinstance(r, str) and r.lower() == "all" and isinstance(c, str) and c.lower() == "all":
            for r_, c_ in self.MT.cell_options:
                if "dropdown" in self.MT.cell_options[(r_, c_)]:
                    self.delete_table_cell_options_dropdown(r_, c_)
        elif isinstance(r, int) and isinstance(c, int):
            self.delete_table_cell_options_dropdown(r, c)

    def delete_cell_dropdown(self, r=0, c=0):
        self.delete_dropdown(r=r, c=c)

    def delete_row_dropdown(self, r="all"):
        if isinstance(r, str) and r.lower() == "all":
            for r_ in self.MT.row_options:
                if "dropdown" in self.MT.row_options[r_]:
                    self.delete_row_options_dropdown(datarn=r_)
        elif isinstance(r, int):
            self.delete_row_options_dropdown(datarn=r)
        elif is_iterable(r):
            for r_ in r:
                self.delete_row_options_dropdown(datarn=r_)

    def delete_column_dropdown(self, c="all"):
        if isinstance(c, str) and c.lower() == "all":
            for c_ in self.MT.col_options:
                if "dropdown" in self.MT.col_options[c_]:
                    self.delete_column_options_dropdown(datacn=c_)
        elif isinstance(c, int):
            self.delete_column_options_dropdown(datacn=c)
        elif is_iterable(c):
            for c_ in c:
                self.delete_column_options_dropdown(datacn=c_)

    def delete_sheet_dropdown(self):
        self.delete_table_options_dropdown()

    def delete_header_dropdown(self, c=None):
        if isinstance(c, str) and c.lower() == "all":
            for c_ in self.CH.cell_options:
                if "dropdown" in self.CH.cell_options[c_]:
                    self.delete_header_cell_options_dropdown(c_)
        elif isinstance(c, int):
            self.delete_header_cell_options_dropdown(c)
        elif is_iterable(c):
            for c_ in c:
                self.delete_header_cell_options_dropdown(c_)
        elif c is None:
            self.delete_header_options_dropdown(c)

    def delete_index_dropdown(self, r=0):
        if isinstance(r, str) and r.lower() == "all":
            for r_ in self.RI.cell_options:
                if "dropdown" in self.RI.cell_options[r_]:
                    self.delete_index_cell_options_dropdown(r_)
        elif isinstance(r, int):
            self.delete_index_cell_options_dropdown(r)
        elif is_iterable(r):
            for r_ in r:
                self.delete_index_cell_options_dropdown(r_)
        elif r is None:
            self.delete_index_options_dropdown()

    def get_dropdowns(self):
        d = {
            **{k: v["dropdown"] for k, v in self.MT.cell_options.items() if "dropdown" in v},
            **{k: v["dropdown"] for k, v in self.MT.row_options.items() if "dropdown" in v},
            **{k: v["dropdown"] for k, v in self.MT.col_options.items() if "dropdown" in v},
        }
        if "dropdown" in self.MT.options:
            return {**d, "dropdown": self.MT.options["dropdown"]}
        return d

    def get_header_dropdowns(self):
        d = {k: v["dropdown"] for k, v in self.CH.cell_options.items() if "dropdown" in v}
        if "dropdown" in self.CH.options:
            return {**d, "dropdown": self.CH.options["dropdown"]}
        return d

    def get_index_dropdowns(self):
        d = {k: v["dropdown"] for k, v in self.RI.cell_options.items() if "dropdown" in v}
        if "dropdown" in self.RI.options:
            return {**d, "dropdown": self.RI.options["dropdown"]}
        return d

    def set_dropdown_values(self, r=0, c=0, set_existing_dropdown: bool = False, values=[], set_value=None):
        if set_existing_dropdown:
            if self.MT.existing_dropdown_window is not None:
                r_ = self.MT.existing_dropdown_window.r
                c_ = self.MT.existing_dropdown_window.c
            else:
                raise Exception("No dropdown box is currently open")
        else:
            r_ = r
            c_ = c
        kwargs = self.MT.get_cell_kwargs(r, c, key="dropdown")
        kwargs["values"] = values
        if kwargs["window"] != "no dropdown open":
            kwargs["window"].values(values)
        if set_value is not None:
            self.set_cell_data(r_, c_, set_value)
            if (
                kwargs["window"] != "no dropdown open"
                and self.MT.text_editor_loc is not None
                and self.MT.text_editor is not None
            ):
                self.MT.text_editor.set_text(set_value)

    def set_header_dropdown_values(self, c=0, set_existing_dropdown: bool = False, values=[], set_value=None):
        if set_existing_dropdown:
            if self.CH.existing_dropdown_window is not None:
                c_ = self.CH.existing_dropdown_window.c
            else:
                raise Exception("No dropdown box is currently open")
        else:
            c_ = c
        kwargs = self.CH.get_cell_kwargs(c_, key="dropdown")
        if kwargs:
            kwargs["values"] = values
            if kwargs["window"] != "no dropdown open":
                kwargs["window"].values(values)
            if set_value is not None:
                self.MT.headers(newheaders=set_value, index=c_)

    def set_index_dropdown_values(self, r, set_existing_dropdown: bool = False, values=[], set_value=None):
        if set_existing_dropdown:
            if self.RI.existing_dropdown_window is not None:
                r_ = self.RI.existing_dropdown_window.r
            else:
                raise Exception("No dropdown box is currently open")
        else:
            r_ = r
        kwargs = self.RI.get_cell_kwargs(r_, key="dropdown")
        if kwargs:
            kwargs["values"] = values
            if kwargs["window"] != "no dropdown open":
                kwargs["window"].values(values)
            if set_value is not None:
                self.MT.row_index(newindex=set_value, index=r_)

    def get_dropdown_values(self, r=0, c=0):
        kwargs = self.MT.get_cell_kwargs(r, c, key="dropdown")
        if kwargs:
            return kwargs["values"]

    def get_header_dropdown_values(self, c=0):
        kwargs = self.CH.get_cell_kwargs(c, key="dropdown")
        if kwargs:
            return kwargs["values"]

    def get_index_dropdown_values(self, r=0):
        kwargs = self.RI.get_cell_kwargs(r, key="dropdown")
        if kwargs:
            kwargs["values"]

    def dropdown_functions(self, r, c, selection_function="", modified_function=""):
        kwargs = self.MT.get_cell_kwargs(r, c, key="dropdown")
        if kwargs:
            if selection_function != "":
                kwargs["select_function"] = selection_function
            if modified_function != "":
                kwargs["modified_function"] = modified_function
            return kwargs

    def header_dropdown_functions(self, c, selection_function="", modified_function=""):
        kwargs = self.CH.get_cell_kwargs(c, key="dropdown")
        if selection_function != "":
            kwargs["selection_function"] = selection_function
        if modified_function != "":
            kwargs["modified_function"] = modified_function
        return kwargs

    def index_dropdown_functions(self, r, selection_function="", modified_function=""):
        kwargs = self.RI.get_cell_kwargs(r, key="dropdown")
        if selection_function != "":
            kwargs["select_function"] = selection_function
        if modified_function != "":
            kwargs["modified_function"] = modified_function
        return kwargs

    def get_dropdown_value(self, r=0, c=0):
        if self.MT.get_cell_kwargs(r, c, key="dropdown"):
            return self.get_cell_data(r, c)

    def get_header_dropdown_value(self, c=0):
        if self.CH.get_cell_kwargs(c, key="dropdown"):
            return self.MT._headers[c]

    def get_index_dropdown_value(self, r=0):
        if self.RI.get_cell_kwargs(r, key="dropdown"):
            return self.MT._row_index[r]

    def open_dropdown(self, r, c):
        self.MT.open_dropdown_window(r, c)

    def close_dropdown(self, r, c):
        self.MT.close_dropdown_window(r, c)

    def open_header_dropdown(self, c):
        self.CH.open_dropdown_window(c)

    def close_header_dropdown(self, c):
        self.CH.close_dropdown_window(c)

    def open_index_dropdown(self, r):
        self.RI.open_dropdown_window(r)

    def close_index_dropdown(self, r):
        self.RI.close_dropdown_window(r)

    def reapply_formatting(self):
        self.MT.reapply_formatting()

    def delete_all_formatting(self, clear_values: bool = False):
        self.MT.delete_all_formatting(clear_values=clear_values)

    def format_cell(
        self,
        r,
        c,
        formatter_options={},
        formatter_class=None,
        redraw: bool = True,
        **kwargs,
    ):
        if isinstance(r, str) and r.lower() == "all" and isinstance(c, int):
            for r_ in range(self.MT.total_data_rows()):
                self.format_cell_(
                    datarn=r_,
                    datacn=c,
                    **{"formatter": formatter_class, **formatter_options, **kwargs},
                )
        elif isinstance(c, str) and c.lower() == "all" and isinstance(r, int):
            for c_ in range(self.MT.total_data_cols()):
                self.format_cell_(
                    datarn=r,
                    datacn=c_,
                    **{"formatter": formatter_class, **formatter_options, **kwargs},
                )
        elif isinstance(r, str) and r.lower() == "all" and isinstance(c, str) and c.lower() == "all":
            for r_ in range(self.MT.total_data_rows()):
                for c_ in range(self.MT.total_data_cols()):
                    self.format_cell_(
                        datarn=r_,
                        datacn=c_,
                        **{"formatter": formatter_class, **formatter_options, **kwargs},
                    )
        else:
            self.format_cell_(
                datarn=r,
                datacn=c,
                **{"formatter": formatter_class, **formatter_options, **kwargs},
            )
        self.set_refresh_timer(redraw)

    def format_cell_(self, datarn: int, datacn: int, **kwargs) -> None:
        if (datarn, datacn) in self.MT.cell_options and "checkbox" in self.MT.cell_options[(datarn, datacn)]:
            self.delete_table_cell_options_checkbox(datarn, datacn)
        kwargs = self.format_fix_kwargs(kwargs)
        if (datarn, datacn) not in self.MT.cell_options:
            self.MT.cell_options[(datarn, datacn)] = {}
        self.MT.cell_options[(datarn, datacn)]["format"] = kwargs
        self.MT.set_cell_data(
            datarn,
            datacn,
            value=kwargs["value"] if "value" in kwargs else self.MT.get_cell_data(datarn, datacn),
            kwargs=kwargs,
        )

    def delete_cell_format(
        self,
        r="all",
        c="all",
        clear_values: bool = False,
    ) -> None:
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

    def format_row(
        self,
        r,
        formatter_options={},
        formatter_class=None,
        redraw: bool = True,
        **kwargs,
    ):
        if isinstance(r, str) and r.lower() == "all":
            for r_ in range(len(self.MT.data)):
                self.format_row_(r_, **{"formatter": formatter_class, **formatter_options, **kwargs})
        elif is_iterable(r):
            for r_ in r:
                self.format_row_(r_, **{"formatter": formatter_class, **formatter_options, **kwargs})
        else:
            self.format_row_(r, **{"formatter": formatter_class, **formatter_options, **kwargs})
        self.set_refresh_timer(redraw)

    def format_row_(self, datarn: int, **kwargs):
        if datarn in self.MT.row_options and "checkbox" in self.MT.row_options[datarn]:
            self.delete_row_options_checkbox(datarn)
        kwargs = self.format_fix_kwargs(kwargs)
        if datarn not in self.MT.row_options:
            self.MT.row_options[datarn] = {}
        self.MT.row_options[datarn]["format"] = kwargs
        for datacn in range(self.MT.total_data_cols()):
            self.MT.set_cell_data(
                datarn,
                datacn,
                value=kwargs["value"] if "value" in kwargs else self.MT.get_cell_data(datarn, datacn),
                kwargs=kwargs,
            )

    def delete_row_format(self, r="all", clear_values: bool = False):
        if is_iterable(r):
            for r_ in r:
                self.MT.delete_row_format(r_, clear_values=clear_values)
        else:
            self.MT.delete_row_format(r, clear_values=clear_values)

    def format_column(
        self,
        c,
        formatter_options={},
        formatter_class=None,
        redraw: bool = True,
        **kwargs,
    ):
        if isinstance(c, str) and c.lower() == "all":
            for c_ in range(self.MT.total_data_cols()):
                self.format_column_(c_, **{"formatter": formatter_class, **formatter_options, **kwargs})
        elif is_iterable(c):
            for c_ in c:
                self.format_column_(c_, **{"formatter": formatter_class, **formatter_options, **kwargs})
        else:
            self.format_column_(c, **{"formatter": formatter_class, **formatter_options, **kwargs})
        self.set_refresh_timer(redraw)

    def format_column_(self, datacn: int, **kwargs):
        if datacn in self.MT.col_options and "checkbox" in self.MT.col_options[datacn]:
            self.delete_column_options_checkbox(datacn)
        kwargs = self.format_fix_kwargs(kwargs)
        if datacn not in self.MT.col_options:
            self.MT.col_options[datacn] = {}
        self.MT.col_options[datacn]["format"] = kwargs
        for datarn in range(self.MT.total_data_rows()):
            self.MT.set_cell_data(
                datarn,
                datacn,
                value=kwargs["value"] if "value" in kwargs else self.MT.get_cell_data(datarn, datacn),
                kwargs=kwargs,
            )

    def delete_column_format(self, c="all", clear_values: bool = False):
        if is_iterable(c):
            for c_ in c:
                self.MT.delete_column_format(c_, clear_values=clear_values)
        else:
            self.MT.delete_column_format(c, clear_values=clear_values)

    def format_sheet(self, formatter_options={}, formatter_class=None, redraw: bool = True, **kwargs):
        self.format_sheet_(**{"formatter": formatter_class, **formatter_options, **kwargs})
        self.set_refresh_timer(redraw)

    def format_sheet_(self, **kwargs):
        kwargs = self.format_fix_kwargs(kwargs)
        self.MT.options["format"] = kwargs
        for datarn in range(self.MT.total_data_rows()):
            for datacn in range(self.MT.total_data_cols()):
                self.MT.set_cell_data(
                    datarn,
                    datacn,
                    value=kwargs["value"] if "value" in kwargs else self.MT.get_cell_data(datarn, datacn),
                    kwargs=kwargs,
                )

    def delete_sheet_format(self, clear_values: bool = False):
        self.MT.delete_sheet_format(clear_values=clear_values)

    def format_fix_kwargs(self, kwargs: dict) -> dict:
        if kwargs["formatter"] is None:
            if kwargs["nullable"]:
                if isinstance(kwargs["datatypes"], (list, tuple)):
                    kwargs["datatypes"] = tuple(kwargs["datatypes"]) + (type(None),)
                else:
                    kwargs["datatypes"] = (kwargs["datatypes"], type(None))
            elif (isinstance(kwargs["datatypes"], (list, tuple)) and type(None) in kwargs["datatypes"]) or kwargs[
                "datatypes"
            ] is type(None):
                raise TypeError("Non-nullable cells cannot have NoneType as a datatype.")
        if not isinstance(kwargs["invalid_value"], str):
            kwargs["invalid_value"] = f"{kwargs['invalid_value']}"
        return kwargs

    #  ##########       TABLE       ##########

    def delete_table_cell_options_dropdown(self, datarn: int, datacn: int):
        self.MT.destroy_opened_dropdown_window()
        if (datarn, datacn) in self.MT.cell_options and "dropdown" in self.MT.cell_options[(datarn, datacn)]:
            del self.MT.cell_options[(datarn, datacn)]["dropdown"]

    def delete_table_cell_options_checkbox(self, datarn: int, datacn: int):
        if (datarn, datacn) in self.MT.cell_options and "checkbox" in self.MT.cell_options[(datarn, datacn)]:
            del self.MT.cell_options[(datarn, datacn)]["checkbox"]

    def delete_table_cell_options_dropdown_and_checkbox(self, datarn: int, datacn: int):
        self.delete_table_cell_options_dropdown(datarn, datacn)
        self.delete_table_cell_options_checkbox(datarn, datacn)

    def delete_row_options_dropdown(self, datarn: int):
        self.MT.destroy_opened_dropdown_window()
        if datarn in self.MT.row_options and "dropdown" in self.MT.row_options[datarn]:
            del self.MT.row_options[datarn]["dropdown"]

    def delete_row_options_checkbox(self, datarn: int):
        if datarn in self.MT.row_options and "checkbox" in self.MT.row_options[datarn]:
            del self.MT.row_options[datarn]["checkbox"]

    def delete_row_options_dropdown_and_checkbox(self, datarn: int):
        self.delete_row_options_dropdown(datarn)
        self.delete_row_options_checkbox(datarn)

    def delete_column_options_dropdown(self, datacn: int):
        self.MT.destroy_opened_dropdown_window()
        if datacn in self.MT.col_options and "dropdown" in self.MT.col_options[datacn]:
            del self.MT.col_options[datacn]["dropdown"]

    def delete_column_options_checkbox(self, datacn: int):
        if datacn in self.MT.col_options and "checkbox" in self.MT.col_options[datacn]:
            del self.MT.col_options[datacn]["checkbox"]

    def delete_column_options_dropdown_and_checkbox(self, datacn):
        self.delete_column_options_dropdown(datacn)
        self.delete_column_options_checkbox(datacn)

    def delete_table_options_dropdown(self):
        self.MT.destroy_opened_dropdown_window()
        if "dropdown" in self.MT.options:
            del self.MT.options["dropdown"]

    def delete_table_options_checkbox(self):
        if "checkbox" in self.MT.options:
            del self.MT.options["checkbox"]

    def delete_table_options_dropdown_and_checkbox(self):
        self.delete_table_options_dropdown()
        self.delete_table_options_checkbox()

    #  ##########       INDEX       ##########

    def delete_index_options_dropdown(self):
        self.RI.destroy_opened_dropdown_window()
        if "dropdown" in self.RI.options:
            del self.RI.options["dropdown"]

    def delete_index_options_checkbox(self):
        if "checkbox" in self.RI.options:
            del self.RI.options["checkbox"]

    def delete_index_options_dropdown_and_checkbox(self):
        self.delete_index_options_dropdown()
        self.delete_index_options_checkbox()

    def delete_index_cell_options_dropdown(self, datarn):
        self.RI.destroy_opened_dropdown_window(datarn=datarn)
        if datarn in self.RI.cell_options and "dropdown" in self.RI.cell_options[datarn]:
            del self.RI.cell_options[datarn]["dropdown"]

    def delete_index_cell_options_checkbox(self, datarn):
        if datarn in self.RI.cell_options and "checkbox" in self.RI.cell_options[datarn]:
            del self.RI.cell_options[datarn]["checkbox"]

    def delete_index_cell_options_dropdown_and_checkbox(self, datarn):
        self.delete_index_cell_options_dropdown(datarn)
        self.delete_index_cell_options_checkbox(datarn)

    #  ##########       HEADER       ##########

    def delete_header_options_dropdown(self):
        self.destroy_opened_dropdown_window()
        if "dropdown" in self.CH.options:
            del self.CH.options["dropdown"]

    def delete_header_options_checkbox(self):
        if "checkbox" in self.CH.options:
            del self.CH.options["checkbox"]

    def delete_header_options_dropdown_and_checkbox(self):
        self.delete_header_options_dropdown()
        self.delete_header_options_checkbox()

    def delete_header_cell_options_dropdown(self, datacn):
        self.destroy_opened_dropdown_window(datacn=datacn)
        if datacn in self.CH.cell_options and "dropdown" in self.CH.cell_options[datacn]:
            del self.CH.cell_options[datacn]["dropdown"]

    def delete_header_cell_options_checkbox(self, datacn):
        if datacn in self.CH.cell_options and "checkbox" in self.CH.cell_options[datacn]:
            del self.CH.cell_options[datacn]["checkbox"]

    def delete_header_cell_options_dropdown_and_checkbox(self, datacn):
        self.delete_header_cell_options_dropdown(datacn)
        self.delete_header_cell_options_checkbox(datacn)


class Dropdown(Sheet):
    def __init__(
        self,
        parent,
        r,
        c,
        width=None,
        height=None,
        font=None,
        colors={
            "bg": theme_light_blue["popup_menu_bg"],
            "fg": theme_light_blue["popup_menu_fg"],
            "highlight_bg": theme_light_blue["popup_menu_highlight_bg"],
            "highlight_fg": theme_light_blue["popup_menu_highlight_fg"],
        },
        outline_color=theme_light_blue["table_fg"],
        outline_thickness=2,
        values=[],
        close_dropdown_window=None,
        modified_function=None,
        search_function=dropdown_search_function,
        arrowkey_RIGHT=None,
        arrowkey_LEFT=None,
        align="w",
        # False for using r, c "r" for r "c" for c
        single_index: str | bool = False,
    ):
        Sheet.__init__(
            self,
            parent=parent,
            outline_thickness=outline_thickness,
            outline_color=outline_color,
            table_grid_fg=colors["fg"],
            show_horizontal_grid=True,
            show_vertical_grid=False,
            show_header=False,
            show_row_index=False,
            show_top_left=False,
            # alignments other than w for dropdown boxes are broken at the moment
            align="w",
            empty_horizontal=0,
            empty_vertical=0,
            selected_rows_to_end_of_window=True,
            horizontal_grid_to_end_of_window=True,
            show_selected_cells_border=False,
            table_selected_cells_border_fg=colors["fg"],
            table_selected_cells_bg=colors["highlight_bg"],
            table_selected_rows_border_fg=colors["fg"],
            table_selected_rows_bg=colors["highlight_bg"],
            table_selected_rows_fg=colors["highlight_fg"],
            width=width,
            height=height,
            font=font if font else get_font(),
            table_fg=colors["fg"],
            table_bg=colors["bg"],
        )
        self.parent = parent
        self.close_dropdown_window = close_dropdown_window
        self.modified_function = modified_function
        self.search_function = search_function
        self.arrowkey_RIGHT = arrowkey_RIGHT
        self.arrowkey_LEFT = arrowkey_LEFT
        self.h_ = height
        self.w_ = width
        self.r = r
        self.c = c
        self.row = -1
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
        if values:
            self.values(values)

    def arrowkey_UP(self, event=None):
        self.deselect("all")
        if self.row > 0:
            self.row -= 1
        else:
            self.row = 0
        self.see(self.row, 0, redraw=False)
        self.select_row(self.row)

    def arrowkey_DOWN(self, event=None):
        self.deselect("all")
        if len(self.MT.data) - 1 > self.row:
            self.row += 1
        self.see(self.row, 0, redraw=False)
        self.select_row(self.row)

    def search_and_see(self, event=None):
        if self.modified_function is not None:
            self.modified_function(event)
        if self.search_function is not None:
            rn = self.search_function(search_for=rf"{event['value']}".lower(), data=self.MT.data)
            if rn is not None:
                self.row = rn
                self.see(self.row, 0, redraw=False)
                self.select_row(self.row)

    def mouse_motion(self, event=None):
        row = self.identify_row(event, exclude_index=True, allow_end=False)
        if row is not None and row != self.row:
            self.row = row
            self.deselect("all", redraw=False)
            self.select_row(self.row)

    def _reselect(self):
        rows = self.get_selected_rows()
        if rows:
            self.select_row(next(iter(rows)))

    def b1(self, event=None):
        if event is None:
            row = None
        elif event.keycode == 13:
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

    def values(self, values=[], redraw=True):
        self.set_sheet_data(
            [[v] for v in values],
            reset_col_positions=False,
            reset_row_positions=False,
            redraw=False,
            verify=False,
        )
        self.set_all_cell_sizes_to_text(redraw=redraw)
