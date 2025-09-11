from __future__ import annotations

import csv as csv
import io
import tkinter as tk
from bisect import bisect_left, bisect_right
from collections import defaultdict, deque
from collections.abc import Callable, Generator, Hashable, Iterator, Sequence
from functools import partial
from itertools import accumulate, chain, filterfalse, islice, repeat
from operator import itemgetter
from re import IGNORECASE, escape, sub
from tkinter import TclError
from typing import Any, Literal

from .colors import color_map
from .column_headers import ColumnHeaders
from .constants import (
    USER_OS,
    _test_str,
    bind_add_columns,
    bind_add_rows,
    bind_del_columns,
    bind_del_rows,
    ctrl_key,
    font_value_error,
    text_editor_close_bindings,
    text_editor_newline_bindings,
    text_editor_to_unbind,
    val_modifying_options,
)
from .find_window import FindWindow, replacer
from .formatters import (
    data_to_str,
    format_data,
    get_clipboard_data,
    get_data_with_valid_check,
    is_bool_like,
    try_to_bool,
)
from .functions import (
    add_to_displayed,
    any_editor_or_dropdown_open,
    b_index,
    bisect_in,
    box_gen_coords,
    box_is_single_cell,
    cell_down_within_box,
    cell_right_within_box,
    color_tup,
    consecutive_ranges,
    data_to_displayed_idxs,
    diff_gen,
    estimate_max_visible_cells,
    event_dict,
    event_has_char_key,
    event_opens_dropdown_or_checkbox,
    float_to_int,
    gen_coords,
    gen_formatted,
    get_bg_fg,
    get_data_from_clipboard,
    get_horizontal_gridline_points,
    get_menu_kwargs,
    get_new_indexes,
    get_seq_without_gaps_at_index,
    get_vertical_gridline_points,
    index_exists,
    insert_items,
    int_x_iter,
    is_iterable,
    is_last_cell,
    is_type_int,
    len_to_idx,
    mod_event_val,
    mod_span,
    move_elements_by_mapping_gen,
    move_fast,
    new_tk_event,
    next_cell,
    push_n,
    recursive_bind,
    rounded_box_coords,
    safe_copy,
    span_idxs_post_move,
    stored_event_dict,
    try_binding,
    widget_descendants,
    wrap_text,
)
from .menus import build_empty_rc_menu, build_header_rc_menu, build_index_rc_menu, build_table_rc_menu
from .other_classes import (
    Box_nt,
    Box_st,
    Box_t,
    DropdownStorage,
    EditorStorageBase,
    EventDataDict,
    FontTuple,
    Highlight,
    Loc,
    ProgressBar,
    Selected,
    SelectionBox,
    TextEditorStorage,
)
from .row_index import RowIndex
from .sorting import sort_selection
from .text_editor import TextEditor
from .tksheet_types import Binding
from .tooltip import Tooltip


class MainTable(tk.Canvas):
    def __init__(
        self,
        parent: tk.Misc,
        row_index_canvas: RowIndex,
        column_headers_canvas: ColumnHeaders,
        **kwargs,
    ) -> None:
        super().__init__(
            parent,
            background=parent.ops.table_bg,
            highlightthickness=0,
        )
        self.PAR = parent
        self.PAR_width = 0
        self.PAR_height = 0
        self.cells_cache = None
        self.table_txt_height, self.index_txt_height, self.header_txt_height = 0, 0, 0
        self.scrollregion = ()
        self.current_cursor = ""
        self.ctrl_b1_pressed = False
        self.b1_pressed_loc = None
        self.closed_dropdown = None
        self.allow_auto_resize_columns = True
        self.allow_auto_resize_rows = True
        self.span = self.PAR.span
        self.synced_scrolls = set()
        self.tooltip = Tooltip(
            **{
                "parent": self,
                "sheet_ops": self.PAR.ops,
                "menu_kwargs": get_menu_kwargs(self.PAR.ops),
                **get_bg_fg(self.PAR.ops),
                "scrollbar_style": f"Sheet{self.PAR.unique_id}.Vertical.TScrollbar",
                "rc_bindings": self.PAR.ops.rc_bindings,
            }
        )
        self.tooltip_widgets = widget_descendants(self.tooltip)
        self.tooltip_coords, self.tooltip_after_id, self.tooltip_showing = None, None, False
        self.tooltip_cell_content = ""
        recursive_bind(self.tooltip, "<Leave>", self.close_tooltip_save)
        self.dropdown = DropdownStorage()
        self.text_editor = TextEditorStorage()
        self.find_window = EditorStorageBase()
        self.find_window_left_x_pc = 1
        self.event_linker = {
            "<<Copy>>": self.ctrl_c,
            "<<Cut>>": self.ctrl_x,
            "<<Paste>>": self.ctrl_v,
            "<<Delete>>": self.delete_key,
            "<<Undo>>": self.undo,
            "<<Redo>>": self.redo,
            "<<SelectAll>>": self.select_all,
        }
        self.enabled_bindings = set()
        self.enabled_bindings_menu_entries = set()
        self.selection_box_ctr = 0
        self.disp_selection_fills = set()
        self.hidd_selection_fills = set()

        self.disp_ctrl_outline = {}
        self.disp_text = {}
        self.disp_high = {}
        self.disp_grid = {}
        self.disp_borders = set()
        self.disp_resize_lines = {}
        self.disp_dropdown = {}
        self.disp_checkbox = {}
        self.disp_corners = set()
        self.hidd_ctrl_outline = {}
        self.hidd_text = {}
        self.hidd_high = {}
        self.hidd_grid = {}
        self.hidd_borders = set()
        self.hidd_resize_lines = {}
        self.hidd_dropdown = {}
        self.hidd_checkbox = {}
        self.hidd_corners = set()

        self.selection_boxes: dict[int, SelectionBox] = {}
        self.selected = ()
        self.named_spans = {}
        self.reset_tags()
        self.cell_options = {}
        self.col_options = {}
        self.row_options = {}
        self.purge_undo_and_redo_stack()
        self.progress_bars = {}

        self.extra_table_rc_menu_funcs = {}
        self.extra_index_rc_menu_funcs = {}
        self.extra_header_rc_menu_funcs = {}
        self.extra_empty_space_rc_menu_funcs = {}

        self.show_index = kwargs["show_index"]
        self.show_header = kwargs["show_header"]
        self.min_row_height = 0
        self.min_header_height = 0
        self.being_drawn_item = None
        self.extra_enter_func = None
        self.extra_b1_press_func = None
        self.extra_b1_motion_func = None
        self.extra_b1_release_func = None
        self.extra_double_b1_func = None
        self.extra_rc_func = None

        self.extra_end_replace_all_func = None

        self.edit_validation_func = None
        self.bulk_table_edit_validation_func = None

        self.extra_begin_sort_cells_func = None
        self.extra_end_sort_cells_func = None

        self.extra_begin_ctrl_c_func = None
        self.extra_end_ctrl_c_func = None

        self.extra_begin_ctrl_x_func = None
        self.extra_end_ctrl_x_func = None

        self.extra_begin_ctrl_v_func = None
        self.extra_end_ctrl_v_func = None

        self.extra_begin_ctrl_z_func = None
        self.extra_end_ctrl_z_func = None

        self.extra_begin_delete_key_func = None
        self.extra_end_delete_key_func = None

        self.extra_begin_edit_cell_func = None
        self.extra_end_edit_cell_func = None

        self.extra_begin_del_rows_rc_func = None
        self.extra_end_del_rows_rc_func = None

        self.extra_begin_del_cols_rc_func = None
        self.extra_end_del_cols_rc_func = None

        self.extra_begin_insert_cols_rc_func = None
        self.extra_end_insert_cols_rc_func = None

        self.extra_begin_insert_rows_rc_func = None
        self.extra_end_insert_rows_rc_func = None

        self.text_editor_user_bound_keys = {}

        self.selection_binding_func = None
        self.deselection_binding_func = None
        self.drag_selection_binding_func = None
        self.shift_selection_binding_func = None
        self.ctrl_selection_binding_func = None
        self.select_all_binding_func = None

        self.tab_enabled = False
        self.up_enabled = False
        self.right_enabled = False
        self.down_enabled = False
        self.left_enabled = False
        self.prior_enabled = False
        self.next_enabled = False
        self.single_selection_enabled = False
        # with this mode every left click adds the cell to selected cells
        self.toggle_selection_enabled = False
        self.drag_selection_enabled = False
        self.select_all_enabled = False
        self.undo_enabled = False
        self.find_enabled = False
        self.cut_enabled = False
        self.copy_enabled = False
        self.paste_enabled = False
        self.delete_key_enabled = False
        self.rc_select_enabled = False
        self.ctrl_select_enabled = False
        self.rc_delete_column_enabled = False
        self.rc_insert_column_enabled = False
        self.rc_delete_row_enabled = False
        self.rc_insert_row_enabled = False
        self.rc_sort_cells_enabled = False
        self.rc_sort_row_enabled = False
        self.rc_sort_column_enabled = False
        self.rc_sort_rows_enabled = False
        self.rc_sort_columns_enabled = False
        self.rc_popup_menus_enabled = False
        self.edit_cell_enabled = False
        self.CH = column_headers_canvas
        self.CH.MT = self
        self.CH.RI = row_index_canvas
        self.RI = row_index_canvas
        self.RI.MT = self
        self.RI.CH = column_headers_canvas
        self.TL = None  # is set from within TopLeftRectangle() __init__
        self.all_columns_displayed = True
        self.all_rows_displayed = True
        self.align = kwargs["align"]
        self.PAR.ops.table_font = FontTuple(
            self.PAR.ops.table_font[0],
            max(1, int(round(self.PAR.ops.table_font[1]) * kwargs["zoom"] / 100)),
            self.PAR.ops.table_font[2],
        )
        self.PAR.ops.index_font = FontTuple(
            self.PAR.ops.index_font[0],
            max(1, int(round(self.PAR.ops.index_font[1]) * kwargs["zoom"] / 100)),
            self.PAR.ops.index_font[2],
        )
        self.PAR.ops.header_font = FontTuple(
            self.PAR.ops.header_font[0],
            max(1, int(round(self.PAR.ops.header_font[1]) * kwargs["zoom"] / 100)),
            self.PAR.ops.header_font[2],
        )
        self.PAR.ops.popup_menu_font = FontTuple(
            self.PAR.ops.popup_menu_font[0],
            max(1, int(round(self.PAR.ops.popup_menu_font[1]) * kwargs["zoom"] / 100)),
            self.PAR.ops.popup_menu_font[2],
        )
        self.txt_measure_canvas = tk.Canvas(self)
        self.txt_measure_canvas_text = self.txt_measure_canvas.create_text(0, 0, text="", font=self.PAR.ops.table_font)

        self.RI.set_width(self.PAR.ops.default_row_index_width)

        self.char_widths = {}
        self.set_table_font_help()
        self.set_header_font_help()
        self.set_index_font_help()

        self.data = kwargs["data_reference"]
        if isinstance(self.data, (list, tuple)):
            self.data = kwargs["data_reference"]
        else:
            self.data = []
        if not self.data and (
            isinstance(kwargs["total_rows"], int)
            and isinstance(kwargs["total_cols"], int)
            and kwargs["total_rows"] > 0
            and kwargs["total_cols"] > 0
        ):
            self.data = [list(repeat("", kwargs["total_cols"])) for _ in range(kwargs["total_rows"])]
        _header = kwargs["header"] if kwargs["header"] is not None else kwargs["headers"]
        if isinstance(_header, int):
            self._headers = _header
        else:
            if _header:
                self._headers = _header
            else:
                self._headers = []
        _row_index = kwargs["index"] if kwargs["index"] is not None else kwargs["row_index"]
        if isinstance(_row_index, int):
            self._row_index = _row_index
        else:
            if _row_index:
                self._row_index = _row_index
            else:
                self._row_index = []
        self.saved_row_heights = {}
        self.saved_column_widths = {}
        self.displayed_columns = []
        self.displayed_rows = []
        self.set_col_positions(itr=[])
        self.set_row_positions(itr=[])
        self.display_rows(
            rows=kwargs["displayed_rows"] if isinstance(kwargs["displayed_rows"], list) else [],
            all_rows_displayed=kwargs["all_rows_displayed"],
            reset_row_positions=False,
            deselect_all=False,
        )
        self.reset_row_positions()
        self.display_columns(
            columns=kwargs["displayed_columns"] if isinstance(kwargs["displayed_columns"], list) else [],
            all_columns_displayed=kwargs["all_columns_displayed"],
            reset_col_positions=False,
            deselect_all=False,
        )
        self.rc_popup_menu, self.empty_rc_popup_menu = None, None
        self.reset_col_positions()
        self.basic_bindings()

    def event_generate(self, *args, **kwargs) -> None:
        for arg in args:
            if arg in self.event_linker:
                self.event_linker[arg]()
            else:
                super().event_generate(*args, **kwargs)

    def refresh(self, event: Any = None) -> None:
        self.PAR.set_refresh_timer()

    def window_configured(self, event: Any) -> None:
        w = self.PAR.winfo_width()
        if w != self.PAR_width:
            self.PAR_width = w
            self.allow_auto_resize_columns = True
        h = self.PAR.winfo_height()
        if h != self.PAR_height:
            self.PAR_height = h
            self.allow_auto_resize_rows = True
        self.main_table_redraw_grid_and_text(True, True)

    def basic_bindings(self, enable: bool = True) -> None:
        bindings = (
            ("<Enter>", self, self.enter),
            ("<Configure>", self, self.window_configured),
            ("<ButtonPress-1>", self, self.b1_press),
            ("<B1-Motion>", self, self.b1_motion),
            ("<ButtonRelease-1>", self, self.b1_release),
            ("<Double-Button-1>", self, self.double_b1),
            ("<MouseWheel>", self, self.mousewheel),
            ("<MouseWheel>", self.RI, self.mousewheel),
            ("<Shift-ButtonPress-1>", self, self.shift_b1_press),
            ("<Shift-ButtonPress-1>", self.CH, self.CH.shift_b1_press),
            ("<Shift-ButtonPress-1>", self.RI, self.RI.shift_b1_press),
            (f"<{ctrl_key}-ButtonPress-1>", self, self.ctrl_b1_press),
            (f"<{ctrl_key}-ButtonPress-1>", self.CH, self.CH.ctrl_b1_press),
            (f"<{ctrl_key}-ButtonPress-1>", self.RI, self.RI.ctrl_b1_press),
            (f"<{ctrl_key}-Shift-ButtonPress-1>", self, self.ctrl_shift_b1_press),
            (f"<{ctrl_key}-Shift-ButtonPress-1>", self.CH, self.CH.ctrl_shift_b1_press),
            (f"<{ctrl_key}-Shift-ButtonPress-1>", self.RI, self.RI.ctrl_shift_b1_press),
            (f"<{ctrl_key}-B1-Motion>", self, self.ctrl_b1_motion),
            (f"<{ctrl_key}-B1-Motion>", self.CH, self.CH.ctrl_b1_motion),
            (f"<{ctrl_key}-B1-Motion>", self.RI, self.RI.ctrl_b1_motion),
        )
        for b in self.PAR.ops.rc_bindings:
            bindings += ((b, self, self.rc),)
        all_canvas_bindings = (
            ("<Shift-MouseWheel>", self.shift_mousewheel),
            ("<Control-MouseWheel>", self.ctrl_mousewheel),
            ("<Control-plus>", self.zoom_in),
            ("<Control-equal>", self.zoom_in),
            ("<Meta-Command-equal>", self.zoom_in),
            ("<Control-minus>", self.zoom_out),
            ("<Meta-Command-minus>", self.zoom_out),
        )
        mt_ri_canvas_linux_bindings = {
            ("<Button-4>", self.mousewheel),
            ("<Button-5>", self.mousewheel),
        }
        all_canvas_linux_bindings = {
            ("<Shift-Button-4>", self.shift_mousewheel),
            ("<Shift-Button-5>", self.shift_mousewheel),
            ("<Control-Button-4>", self.ctrl_mousewheel),
            ("<Control-Button-5>", self.ctrl_mousewheel),
        }
        if enable:
            for b in bindings:
                b[1].bind(b[0], b[2])
            for b in all_canvas_bindings:
                for canvas in (self, self.RI, self.CH):
                    canvas.bind(b[0], b[1])
            if USER_OS == "linux":
                for b in mt_ri_canvas_linux_bindings:
                    for canvas in (self, self.RI):
                        canvas.bind(b[0], b[1])
                for b in all_canvas_linux_bindings:
                    for canvas in (self, self.RI, self.CH):
                        canvas.bind(b[0], b[1])
        else:
            for b in bindings:
                b[1].unbind(b[0])
            for b in all_canvas_bindings:
                for canvas in (self, self.RI, self.CH):
                    canvas.unbind(b[0])
            if USER_OS == "linux":
                for b in mt_ri_canvas_linux_bindings:
                    for canvas in (self, self.RI):
                        canvas.unbind(b[0])
                for b in all_canvas_linux_bindings:
                    for canvas in (self, self.RI, self.CH):
                        canvas.unbind(b[0])

    def reset_tags(self) -> None:
        self.tagged_cells = {}
        self.tagged_rows = {}
        self.tagged_columns = {}

    def create_ctrl_outline(
        self,
        x1: int,
        y1: int,
        x2: int,
        y2: int,
        fill: str,
        dash: tuple[int, int],
        width: int,
        outline: str,
        tags: str | tuple[str, ...],
    ) -> int:
        if self.hidd_ctrl_outline:
            t, sh = self.hidd_ctrl_outline.popitem()
            self.coords(t, x1, y1, x2, y2)
            if sh:
                self.itemconfig(t, fill=fill, dash=dash, width=width, outline=outline, tags=tags)
            else:
                self.itemconfig(
                    t,
                    fill=fill,
                    dash=dash,
                    width=width,
                    outline=outline,
                    tags=tags,
                    state="normal",
                )
        else:
            t = self.create_rectangle(
                x1,
                y1,
                x2,
                y2,
                fill=fill,
                dash=dash,
                width=width,
                outline=outline,
                tags=tags,
            )
        self.lift(t)
        self.disp_ctrl_outline[t] = True
        return t

    def delete_ctrl_outlines(self, iid: int | None = None) -> None:
        if isinstance(iid, int) and iid in self.disp_ctrl_outline:
            self.hidd_ctrl_outline[iid] = self.disp_ctrl_outline.pop(iid)
            self.itemconfig(iid, state="hidden")
            self.hidd_ctrl_outline[iid] = False
        else:
            self.hidd_ctrl_outline.update(self.disp_ctrl_outline)
            self.disp_ctrl_outline = {}
            for t, sh in self.hidd_ctrl_outline.items():
                if sh:
                    self.itemconfig(t, state="hidden")
                    self.hidd_ctrl_outline[t] = False

    def show_ctrl_outline(
        self,
        canvas: Literal["table"] = "table",
        start_cell: tuple[int, int] = (0, 0),
        end_cell: tuple[int, int] = (0, 0),
        dash: tuple[int, int] = (15, 15),
        outline: str | None = None,
        delete_on_timer: bool = True,
    ) -> None:
        iid = self.create_ctrl_outline(
            self.col_positions[start_cell[0]] + 1,
            self.row_positions[start_cell[1]] + 1,
            self.col_positions[end_cell[0]],
            self.row_positions[end_cell[1]],
            fill="",
            dash=dash,
            width=2,
            outline=self.PAR.ops.resizing_line_fg if outline is None else outline,
            tags="lift",
        )
        if delete_on_timer:
            self.after(1500, lambda: self.delete_ctrl_outlines(iid))

    def escape(self, event: tk.Misc | None) -> None:
        if self.find_window.open:
            self.close_find_window()
        else:
            self.deselect()

    def get_find_window_dimensions_coords(self, w_width: int | None) -> tuple[int, int, int, int]:
        if w_width is None:
            w_width = self.winfo_width()
        width = min(self.char_width_fn("X") * 23, w_width - 7)
        height = self.min_row_height
        if self.find_window.window and self.find_window.window.replace_visible:
            height *= 2
        # Position from left based on percentage
        xpos = w_width * self.find_window_left_x_pc
        # Clamp to stay within canvas bounds
        xpos = min(xpos, w_width - width - 7)  # Don’t exceed right edge
        xpos = max(0, xpos)  # Don’t go left of 0
        return width, height, self.canvasx(xpos), self.canvasy(7)

    def reposition_find_window(self, w_width: int | None = None) -> None:
        if w_width is None:
            w_width = self.winfo_width()
        w, h, x, y = self.get_find_window_dimensions_coords(w_width=w_width)
        self.coords(self.find_window.canvas_id, x, y)
        self.itemconfig(
            self.find_window.canvas_id,
            width=w,
            height=h,
            state="normal",
        )

    def drag_find_window(self, event: tk.Event) -> None:
        """Receives a tkinter b1-motion event, is bound to a label on the find window"""
        # Convert screen coordinates to canvas window coordinates
        window_x = event.x_root - self.winfo_rootx()
        # Get the visible canvas width
        visible_width = self.winfo_width()
        if visible_width > 0:
            # Calculate the new percentage using widget-relative coordinates
            new_pc = window_x / visible_width
            # Clamp the percentage between 0 and 1
            self.find_window_left_x_pc = min(max(new_pc, 0), 1)
        # Reposition the find window based on the updated percentage
        self.reposition_find_window()

    def open_find_window(
        self,
        event: tk.Misc | None = None,
        focus: bool = True,
    ) -> Literal["break"]:
        if self.find_window.open:
            self.find_window.window.tktext.focus_set()
            return "break"
        width, height, x, y = self.get_find_window_dimensions_coords(w_width=self.winfo_width())
        if not self.find_window.window:
            self.find_window.window = FindWindow(
                self,
                find_prev_func=self.find_previous,
                find_next_func=self.find_next,
                close_func=self.close_find_window,
                replace_func=self.replace_next,
                replace_all_func=self.replace_all,
                toggle_replace_func=self.reposition_find_window,
                drag_func=self.drag_find_window,
                rc_bindings=self.PAR.ops.rc_bindings,
            )
            self.find_window.canvas_id = self.create_window((x, y), window=self.find_window.window, anchor="nw")
        else:
            self.coords(self.find_window.canvas_id, x, y)
            if not self.find_window.open:
                self.itemconfig(self.find_window.canvas_id, state="normal")
        self.find_window.open = True
        self.find_window.window.reset(
            **{
                "menu_kwargs": get_menu_kwargs(self.PAR.ops),
                "sheet_ops": self.PAR.ops,
                "border_color": self.PAR.ops.table_selected_box_cells_fg,
                "grid_color": self.PAR.ops.table_grid_fg,
                **get_bg_fg(self.PAR.ops),
                "replace_enabled": "replace" in self.enabled_bindings or "all" in self.enabled_bindings,
            }
        )
        self.itemconfig(self.find_window.canvas_id, width=width, height=height)
        if focus:
            self.find_window.tktext.focus_set()
        return "break"

    def replace_next(self, event: tk.Misc | None = None) -> None:
        find = self.find_window.get().lower()
        replace = self.find_window.window.get_replace()
        sel = self.selected
        if sel:
            datarn, datacn = self.datarn(sel.row), self.datacn(sel.column)
            m = self.find_match(find, datarn, datacn)
            if m:
                current = f"{self.get_cell_data(datarn, datacn, True)}"
                new = sub(escape(find), replacer(find, replace, current), current, flags=IGNORECASE)
                event_data = self.new_single_edit_event(
                    sel.row, sel.column, datarn, datacn, "replace_next", self.get_cell_data(datarn, datacn), new
                )
                value, event_data = self.single_edit_run_validation(datarn, datacn, event_data)
                if value is not None and (
                    self.set_cell_data_undo(
                        r=datarn,
                        c=datacn,
                        datarn=datarn,
                        datacn=datacn,
                        value=value,
                        redraw=False,
                    )
                ):
                    try_binding(self.extra_end_edit_cell_func, event_data)
        if self.find_window.window.find_in_selection:
            found_next = self.find_see_and_set(self.find_within(find))
        else:
            found_next = self.find_see_and_set(self.find_all_cells(find))
        if not found_next and not self.find_window.window.find_in_selection:
            self.deselect()

    def replace_all(self, event: tk.Misc | None = None) -> None:
        find = self.find_window.get().lower()
        replace = self.find_window.window.get_replace()
        tree = self.PAR.ops.treeview
        event_data = self.new_event_dict("edit_table")
        boxes = self.get_boxes()
        event_data["selection_boxes"] = boxes
        if self.find_window.window.find_in_selection:
            iterable = chain.from_iterable(
                (
                    box_gen_coords(
                        *box.coords,
                        start_r=box.coords.from_r,
                        start_c=box.coords.from_c,
                        reverse=False,
                        all_rows_displayed=self.all_rows_displayed,
                        all_cols_displayed=self.all_columns_displayed,
                        displayed_rows=self.displayed_rows,
                        displayed_cols=self.displayed_columns,
                    )
                    for box in self.selection_boxes.values()
                )
            )
        else:
            iterable = box_gen_coords(
                from_r=0,
                from_c=0,
                upto_r=self.total_data_rows(include_index=False),
                upto_c=self.total_data_cols(include_header=False),
                start_r=0,
                start_c=0,
                reverse=False,
            )
        for r, c in iterable:
            m = self.find_match(find, r, c)
            if m and (
                (tree or self.all_rows_displayed or bisect_in(self.displayed_rows, r))
                and (self.all_columns_displayed or bisect_in(self.displayed_columns, c))
            ):
                current = f"{self.get_cell_data(r, c, True)}"
                new = sub(escape(find), replacer(find, replace, current), current, flags=IGNORECASE)
                if not self.edit_validation_func or (
                    self.edit_validation_func
                    and (new := self.edit_validation_func(mod_event_val(event_data, new, (r, c)))) is not None
                ):
                    event_data = self.event_data_set_cell(
                        r,
                        c,
                        new,
                        event_data,
                    )
        event_data = self.bulk_edit_validation(event_data)
        if event_data["cells"]["table"]:
            self.refresh()
            if self.undo_enabled:
                self.undo_stack.append(stored_event_dict(event_data))
            try_binding(self.extra_end_replace_all_func, event_data, "end_edit_table")
            self.sheet_modified(event_data)
            self.PAR.emit_event("<<SheetModified>>", event_data)

    def find_see_and_set(
        self, coords: tuple[int, int, int | None] | None, within: bool | None = None
    ) -> tuple[int, int]:
        if coords:
            row, column, item = coords
            if self.PAR.ops.treeview:
                self.PAR.scroll_to_item(self.PAR.rowitem(row, data_index=True))
            disp_row = self.disprn(row) if not self.all_rows_displayed else row
            disp_col = self.dispcn(column) if not self.all_columns_displayed else column
            if within or (self.find_window.window and self.find_window.window.find_in_selection):
                self.set_currently_selected(disp_row, disp_col, item=item)
            else:
                self.select_cell(disp_row, disp_col, redraw=False)
            if not self.see(disp_row, disp_col):
                self.refresh()
        return coords

    def find_match(self, find: str, r: int, c: int) -> bool:
        try:
            value = self.data[r][c]
        except Exception:
            value = ""
        kwargs = self.get_cell_kwargs(r, c, key="format")
        if kwargs:
            # assumed given formatter class has __str__() or value attribute
            value = data_to_str(value, **kwargs) if kwargs["formatter"] is None else str(value)
        if value is None:
            return find == ""
        elif not find:
            return str(value) == ""
        else:
            return find in str(value).lower()

    def find_within_current_box(
        self,
        current_box: SelectionBox,
        find: str,
        reverse: bool,
        stop: None | tuple[int, int] = None,
    ) -> None | tuple[int, int, int]:
        if stop:
            start_r, start_c = current_box.coords[0], current_box.coords[1]
        else:
            start_r, start_c = next_cell(
                *current_box.coords,
                self.selected.row,
                self.selected.column,
                reverse=reverse,
            )
        iterable = box_gen_coords(
            *current_box.coords,
            start_r,
            start_c,
            reverse=reverse,
            all_rows_displayed=self.all_rows_displayed,
            all_cols_displayed=self.all_columns_displayed,
            displayed_rows=self.displayed_rows,
            displayed_cols=self.displayed_columns,
            no_wrap=True,
        )
        if stop:
            for r, c in iterable:
                if (r, c) == stop:
                    return None
                elif (
                    self.find_match(find, r, c)  # will not show hidden rows
                    and (self.all_rows_displayed or bisect_in(self.displayed_rows, r))
                    and (self.all_columns_displayed or bisect_in(self.displayed_columns, c))
                ):
                    return (r, c, current_box.fill_iid)
        else:
            return next(
                (
                    (r, c, current_box.fill_iid)
                    for r, c in iterable
                    if (
                        self.find_match(find, r, c)  # will not show hidden rows
                        and (self.all_rows_displayed or bisect_in(self.displayed_rows, r))
                        and (self.all_columns_displayed or bisect_in(self.displayed_columns, c))
                    )
                ),
                None,
            )

    def find_within_non_current_boxes(self, current_id: int, find: str, reverse: bool) -> None | tuple[int, int, int]:
        fn = partial(
            box_gen_coords,
            reverse=reverse,
            all_rows_displayed=self.all_rows_displayed,
            all_cols_displayed=self.all_columns_displayed,
            displayed_rows=self.displayed_rows,
            displayed_cols=self.displayed_columns,
        )
        if reverse:
            # iterate backwards through selection boxes from the box before current
            idx = next(i for i, k in enumerate(reversed(self.selection_boxes)) if k == current_id)
            return next(
                (
                    (r, c, item)
                    for item, box in chain(
                        islice(reversed(self.selection_boxes.items()), idx + 1, None),
                        islice(reversed(self.selection_boxes.items()), 0, idx),
                    )
                    for r, c in fn(*box.coords, box.coords.upto_r - 1, box.coords.upto_c - 1)
                    if (
                        self.find_match(find, r, c)  # will not show hidden rows
                        and (self.all_rows_displayed or bisect_in(self.displayed_rows, r))
                        and (self.all_columns_displayed or bisect_in(self.displayed_columns, c))
                    )
                ),
                None,
            )
        else:
            # iterate forwards through selection boxes from the box after current
            idx = next(i for i, k in enumerate(self.selection_boxes) if k == current_id)
            return next(
                (
                    (r, c, item)
                    for item, box in chain(
                        islice(self.selection_boxes.items(), idx + 1, None),
                        islice(self.selection_boxes.items(), 0, idx),
                    )
                    for r, c in fn(*box.coords, box.coords.from_r, box.coords.from_c)
                    if (
                        self.find_match(find, r, c)
                        and (self.all_rows_displayed or bisect_in(self.displayed_rows, r))
                        and (self.all_columns_displayed or bisect_in(self.displayed_columns, c))
                    )
                ),
                None,
            )

    def find_within(self, find: str, reverse: bool = False) -> tuple[int, int, int] | None:
        if not self.selected:
            return None
        current_box = self.selection_boxes[self.selected.fill_iid]
        current_id = self.selected.fill_iid
        if is_last_cell(*current_box.coords, self.selected.row, self.selected.column, reverse=reverse):
            if (coord := self.find_within_non_current_boxes(current_id=current_id, find=find, reverse=reverse)) or (
                coord := self.find_within_current_box(current_box=current_box, find=find, reverse=reverse)
            ):
                return coord
        else:
            if (coord := self.find_within_current_box(current_box=current_box, find=find, reverse=reverse)) or (
                coord := self.find_within_non_current_boxes(current_id=current_id, find=find, reverse=reverse)
            ):
                return coord
            elif self.selected.row != current_box.coords[0] or self.selected.column != current_box.coords[1]:
                stop = (self.datarn(self.selected.row), self.datacn(self.selected.column))
                coord = self.find_within_current_box(current_box=current_box, find=find, reverse=reverse, stop=stop)
                if coord:
                    return coord
        return None

    def find_all_cells(self, find: str, reverse: bool = False) -> tuple[int, int, None] | None:
        tree = self.PAR.ops.treeview
        totalrows = self.total_data_rows(include_index=False)
        totalcols = self.total_data_cols(include_header=False)
        if self.selected:
            start_r, start_c = next_cell(
                0,
                0,
                totalrows,
                totalcols,
                self.datarn(self.selected.row),
                self.datacn(self.selected.column),
                reverse=reverse,
            )
        else:
            start_r, start_c = 0, 0
        return next(
            (
                (r, c, None)
                for r, c in box_gen_coords(
                    from_r=0,
                    from_c=0,
                    upto_r=totalrows,
                    upto_c=totalcols,
                    start_r=start_r,
                    start_c=start_c,
                    reverse=reverse,
                )
                if (
                    self.find_match(find, r, c)
                    and (tree or self.all_rows_displayed or bisect_in(self.displayed_rows, r))
                    and (self.all_columns_displayed or bisect_in(self.displayed_columns, c))
                )
            ),
            None,
        )

    def replace_toggle(self, event: tk.Event | None) -> None:
        if not self.find_window.open:
            self.open_find_window(focus=False)
        if not self.find_window.window.replace_visible:
            self.find_window.window.toggle_replace_window()
        self.find_window.window.replace_tktext.focus_set()

    def find_next(
        self,
        event: tk.Misc | None = None,
        within: bool | None = None,
        find: str | None = None,
        reverse: bool = False,
    ) -> Literal["break"]:
        if find is None:
            find = self.find_window.get().lower()
        if find is None and not self.find_window.open:
            self.open_find_window(focus=False)
        if within or (self.find_window.window and self.find_window.window.find_in_selection):
            self.find_see_and_set(self.find_within(find, reverse=reverse), within=within)
        else:
            self.find_see_and_set(self.find_all_cells(find, reverse=reverse), within=within)
        return "break"

    def find_previous(
        self, event: tk.Misc | None = None, within: bool | None = None, find: str | None = None
    ) -> Literal["break"]:
        return self.find_next(find=find, within=within, reverse=True)

    def close_find_window(self, event: tk.Misc | None = None) -> None:
        if self.find_window.open:
            self.itemconfig(self.find_window.canvas_id, state="hidden")
            self.find_window.open = False
            self.focus_set()

    def get_ctrl_x_c_boxes(self) -> tuple[dict[tuple[int, int, int, int], str], int]:
        maxrows = 0
        if not self.selected:
            return {}, maxrows
        if self.selected.type_ in ("cells", "columns"):
            curr_box = self.selection_boxes[self.selected.fill_iid].coords
            maxrows = curr_box[2] - curr_box[0]
            boxes = {
                box.coords: box.type_
                for _, box in self.get_selection_items(rows=False)
                if maxrows >= box.coords[2] - box.coords[0]
            }
        else:
            boxes = {box.coords: "rows" for _, box in self.get_selection_items(columns=False, cells=False)}
        return boxes, maxrows

    def io_csv_writer(self) -> tuple[io.StringIO, csv.writer]:
        s = io.StringIO()
        writer = csv.writer(
            s,
            dialect=self.PAR.ops.to_clipboard_dialect,
            delimiter=self.PAR.ops.to_clipboard_delimiter,
            quotechar=self.PAR.ops.to_clipboard_quotechar,
            lineterminator=self.PAR.ops.to_clipboard_lineterminator,
        )
        return s, writer

    def new_event_dict(self, name: str, boxes: dict | None = None, state: bool = False) -> EventDataDict:
        return event_dict(
            name=name,
            sheet=self.PAR.name,
            widget=self,
            boxes=self.get_boxes() if boxes is None else boxes,
            selected=self.selected,
            sheet_state=self.copy_sheet_state() if state else None,
        )

    def new_single_edit_event(
        self, r: int, c: int, datarn: int, datacn: int, k: str, before_val: Any, after_val: Any
    ) -> EventDataDict:
        return event_dict(
            name="end_edit_table",
            sheet=self.PAR.name,
            widget=self,
            cells_table={(datarn, datacn): before_val},
            key=k,
            value=after_val,
            loc=Loc(r, c),
            row=r,
            column=c,
            boxes=self.get_boxes(),
            selected=self.selected,
            data={(datarn, datacn): after_val},
        )

    def ctrl_c(self, event=None) -> None | EventDataDict:
        if not self.selected:
            return
        event_data = self.new_event_dict("begin_ctrl_c")
        boxes, maxrows = self.get_ctrl_x_c_boxes()
        event_data["selection_boxes"] = boxes
        s, writer = self.io_csv_writer()
        if not try_binding(self.extra_begin_ctrl_c_func, event_data):
            return
        if self.selected.type_ in ("cells", "columns"):
            for rn in range(maxrows):
                row = []
                for r1, c1, _, c2 in boxes:
                    datarn = (r1 + rn) if self.all_rows_displayed else self.displayed_rows[r1 + rn]
                    for c in range(c1, c2):
                        datacn = self.datacn(c)
                        v = self.get_cell_clipboard(datarn, datacn)
                        event_data["cells"]["table"][(datarn, datacn)] = v
                        row.append(v)
                writer.writerow(row)
        else:
            for r1, c1, r2, c2 in boxes:
                for rn in range(r2 - r1):
                    row = []
                    datarn = (r1 + rn) if self.all_rows_displayed else self.displayed_rows[r1 + rn]
                    for c in range(c1, c2):
                        datacn = self.datacn(c)
                        v = self.get_cell_clipboard(datarn, datacn)
                        event_data["cells"]["table"][(datarn, datacn)] = v
                        row.append(v)
                    writer.writerow(row)
        for r1, c1, r2, c2 in boxes:
            self.show_ctrl_outline(canvas="table", start_cell=(c1, r1), end_cell=(c2, r2))
        self.clipboard_clear()
        if len(event_data["cells"]["table"]) == 1 and self.PAR.ops.to_clipboard_lineterminator not in next(
            iter(event_data["cells"]["table"].values())
        ):
            self.clipboard_append(next(iter(event_data["cells"]["table"].values())))
        else:
            self.clipboard_append(s.getvalue())
        self.update_idletasks()
        try_binding(self.extra_end_ctrl_c_func, event_data, new_name="end_ctrl_c")
        self.PAR.emit_event("<<Copy>>", EventDataDict({**event_data, **{"eventname": "copy"}}))
        return event_data

    def ctrl_c_plain(self, event=None) -> None | EventDataDict:
        if not self.selected:
            return
        event_data = self.new_event_dict("begin_ctrl_c")
        boxes, maxrows = self.get_ctrl_x_c_boxes()
        event_data["selection_boxes"] = boxes
        s = io.StringIO()
        delim, line_term = self.PAR.ops.to_clipboard_delimiter, self.PAR.ops.to_clipboard_lineterminator
        if not try_binding(self.extra_begin_ctrl_c_func, event_data):
            return
        if self.selected.type_ in ("cells", "columns"):
            for rn in range(maxrows):
                row = []
                for r1, c1, _, c2 in boxes:
                    datarn = (r1 + rn) if self.all_rows_displayed else self.displayed_rows[r1 + rn]
                    for c in range(c1, c2):
                        datacn = self.datacn(c)
                        v = self.get_cell_clipboard(datarn, datacn)
                        event_data["cells"]["table"][(datarn, datacn)] = v
                        row.append(v)
                s.write(delim.join(row) + line_term)
        else:
            for r1, c1, r2, c2 in boxes:
                for rn in range(r2 - r1):
                    row = []
                    datarn = (r1 + rn) if self.all_rows_displayed else self.displayed_rows[r1 + rn]
                    for c in range(c1, c2):
                        datacn = self.datacn(c)
                        v = self.get_cell_clipboard(datarn, datacn)
                        event_data["cells"]["table"][(datarn, datacn)] = v
                        row.append(v)
                    s.write(delim.join(row) + line_term)
        for r1, c1, r2, c2 in boxes:
            self.show_ctrl_outline(canvas="table", start_cell=(c1, r1), end_cell=(c2, r2))
        self.clipboard_clear()
        if len(event_data["cells"]["table"]) == 1 and self.PAR.ops.to_clipboard_lineterminator not in next(
            iter(event_data["cells"]["table"].values())
        ):
            self.clipboard_append(next(iter(event_data["cells"]["table"].values())))
        else:
            self.clipboard_append(s.getvalue())
        self.update_idletasks()
        try_binding(self.extra_end_ctrl_c_func, event_data, new_name="end_ctrl_c")
        self.PAR.emit_event("<<Copy>>", EventDataDict({**event_data, **{"eventname": "copy"}}))
        return event_data

    def ctrl_x(self, event=None, validation: bool = True) -> None | EventDataDict:
        if not self.selected:
            return
        event_data = self.new_event_dict("edit_table")
        boxes, maxrows = self.get_ctrl_x_c_boxes()
        event_data["selection_boxes"] = boxes
        s, writer = self.io_csv_writer()
        if not try_binding(self.extra_begin_ctrl_x_func, event_data, "begin_ctrl_x"):
            return
        if self.selected.type_ in ("cells", "columns"):
            for rn in range(maxrows):
                row = []
                for r1, c1, r2, c2 in boxes:
                    if r2 - r1 < maxrows:
                        continue
                    datarn = (r1 + rn) if self.all_rows_displayed else self.displayed_rows[r1 + rn]
                    for c in range(c1, c2):
                        datacn = self.datacn(c)
                        row.append(self.get_cell_clipboard(datarn, datacn))
                        val = self.get_value_for_empty_cell(datarn, datacn)
                        if (
                            not self.edit_validation_func
                            or not validation
                            or (
                                self.edit_validation_func
                                and (val := self.edit_validation_func(mod_event_val(event_data, val, (r1 + rn, c))))
                                is not None
                            )
                        ):
                            event_data = self.event_data_set_cell(
                                datarn,
                                datacn,
                                val,
                                event_data,
                            )
                writer.writerow(row)
        else:
            for r1, c1, r2, c2 in boxes:
                for rn in range(r2 - r1):
                    row = []
                    datarn = (r1 + rn) if self.all_rows_displayed else self.displayed_rows[r1 + rn]
                    for c in range(c1, c2):
                        datacn = self.datacn(c)
                        row.append(self.get_cell_clipboard(datarn, datacn))
                        val = self.get_value_for_empty_cell(datarn, datacn)
                        if (
                            not self.edit_validation_func
                            or not validation
                            or (
                                self.edit_validation_func
                                and (val := self.edit_validation_func(mod_event_val(event_data, val, (r1 + rn, c))))
                                is not None
                            )
                        ):
                            event_data = self.event_data_set_cell(
                                datarn,
                                datacn,
                                val,
                                event_data,
                            )
                    writer.writerow(row)
        self.clipboard_clear()
        if len(event_data["cells"]["table"]) == 1 and self.PAR.ops.to_clipboard_lineterminator not in next(
            iter(event_data["cells"]["table"].values())
        ):
            self.clipboard_append(next(iter(event_data["cells"]["table"].values())))
        else:
            self.clipboard_append(s.getvalue())
        self.update_idletasks()
        event_data = self.bulk_edit_validation(event_data)
        if event_data["cells"]["table"]:
            self.refresh()
            for r1, c1, r2, c2 in boxes:
                self.show_ctrl_outline(canvas="table", start_cell=(c1, r1), end_cell=(c2, r2))
            if self.undo_enabled:
                self.undo_stack.append(stored_event_dict(event_data))
            try_binding(self.extra_end_ctrl_x_func, event_data, "end_ctrl_x")
            self.sheet_modified(event_data)
            self.PAR.emit_event("<<Cut>>", event_data)
        return event_data

    def sort_boxes(
        self,
        event: tk.Event | None = None,
        boxes: dict[tuple[int, int, int, int], Literal["cells", "rows", "columns"]] | None = None,
        reverse: bool = False,
        row_wise: bool = False,
        validation: bool = True,
        key: Callable | None = None,
        undo: bool = True,
    ) -> EventDataDict:
        if not boxes:
            boxes = self.get_boxes()
        if not boxes:
            boxes = {Box_nt(0, 0, len(self.row_positions) - 1, len(self.col_positions) - 1): "cells"}
        event_data = self.new_event_dict("edit_table", boxes=boxes)
        try_binding(self.extra_begin_sort_cells_func, event_data)
        if key is None:
            key = self.PAR.ops.sort_key
        for r1, c1, r2, c2 in boxes:
            data = sort_selection(
                [[self.get_cell_data(self.datarn(r), self.datacn(c)) for c in range(c1, c2)] for r in range(r1, r2)],
                reverse=reverse,
                key=key,
                row_wise=row_wise,
            )
            for ir, r in enumerate(range(r1, r2)):
                data_r = self.datarn(r)
                for ic, c in enumerate(range(c1, c2)):
                    data_c = self.datacn(c)
                    val = data[ir][ic]
                    if (
                        not self.edit_validation_func
                        or not validation
                        or (
                            self.edit_validation_func
                            and (val := self.edit_validation_func(mod_event_val(event_data, val, (data_r, data_c))))
                            is not None
                        )
                    ):
                        event_data = self.event_data_set_cell(
                            datarn=data_r,
                            datacn=data_c,
                            value=val,
                            event_data=event_data,
                        )
        event_data = self.bulk_edit_validation(event_data)
        if event_data["cells"]["table"]:
            if undo and self.undo_enabled:
                self.undo_stack.append(stored_event_dict(event_data))
            try_binding(self.extra_end_sort_cells_func, event_data, "end_edit_table")
            self.sheet_modified(event_data)
            self.PAR.emit_event("<<SheetModified>>", event_data)
            self.refresh()
        return event_data

    def ctrl_v(self, event: Any = None, validation: bool = True) -> None | EventDataDict:
        if not self.PAR.ops.paste_can_expand_x and len(self.col_positions) == 1:
            return
        if not self.PAR.ops.paste_can_expand_y and len(self.row_positions) == 1:
            return
        event_data = self.new_event_dict("edit_table", state=True)
        if self.selected:
            selected_r = self.selected.box.from_r
            selected_c = self.selected.box.from_c
            curr_coords = (self.selected.row, self.selected.column)
        elif not self.selected and not self.PAR.ops.paste_can_expand_x and not self.PAR.ops.paste_can_expand_y:
            return
        else:
            if not self.data:
                selected_c, selected_r = 0, 0
            else:
                if len(self.col_positions) == 1 and len(self.row_positions) > 1:
                    selected_c, selected_r = 0, len(self.row_positions) - 1
                elif len(self.row_positions) == 1 and len(self.col_positions) > 1:
                    selected_c, selected_r = len(self.col_positions) - 1, 0
                elif len(self.row_positions) > 1 and len(self.col_positions) > 1:
                    selected_c, selected_r = 0, len(self.row_positions) - 1
            curr_coords = (selected_r, selected_c)
        try:
            data = get_data_from_clipboard(
                widget=self,
                delimiters=self.PAR.ops.from_clipboard_delimiters,
                lineterminator=self.PAR.ops.to_clipboard_lineterminator,
            )
        except Exception:
            return
        new_data_numcols = max(map(len, data))
        new_data_numrows = len(data)
        for rn, r in enumerate(data):
            if len(r) < new_data_numcols:
                data[rn] += list(repeat("", new_data_numcols - len(r)))
        if self.selected:
            (
                lastbox_r1,
                lastbox_c1,
                lastbox_r2,
                lastbox_c2,
            ) = self.selection_boxes[self.selected.fill_iid].coords
            lastbox_numrows = lastbox_r2 - lastbox_r1
            lastbox_numcols = lastbox_c2 - lastbox_c1
            if lastbox_numrows > new_data_numrows and not lastbox_numrows % new_data_numrows:
                repeat_num = int(lastbox_numrows / new_data_numrows)
                data.extend(list(chain.from_iterable([r.copy() for r in data] for _ in range(repeat_num - 1))))
                new_data_numrows *= repeat_num

            if lastbox_numcols > new_data_numcols and not lastbox_numcols % new_data_numcols:
                repeat_num = int(lastbox_numcols / new_data_numcols)
                for rn, row in enumerate(data):
                    copies = [row.copy() for _ in range(repeat_num - 1)]
                    data[rn].extend(chain.from_iterable(copies))
                new_data_numcols *= repeat_num
        added_rows = 0
        added_cols = 0
        total_data_cols = None
        if self.PAR.ops.paste_can_expand_x and selected_c + new_data_numcols > len(self.col_positions) - 1:
            total_data_cols = self.equalize_data_row_lengths()
            added_cols = selected_c + new_data_numcols - len(self.col_positions) + 1
            if (
                isinstance(self.PAR.ops.paste_insert_column_limit, int)
                and self.PAR.ops.paste_insert_column_limit < len(self.col_positions) - 1 + added_cols
            ):
                added_cols = self.PAR.ops.paste_insert_column_limit - len(self.col_positions) - 1
        if self.PAR.ops.paste_can_expand_y and selected_r + new_data_numrows > len(self.row_positions) - 1:
            added_rows = selected_r + new_data_numrows - len(self.row_positions) + 1
            if (
                isinstance(self.PAR.ops.paste_insert_row_limit, int)
                and self.PAR.ops.paste_insert_row_limit < len(self.row_positions) - 1 + added_rows
            ):
                added_rows = self.PAR.ops.paste_insert_row_limit - len(self.row_positions) - 1
        if selected_c + new_data_numcols > len(self.col_positions) - 1:
            adjusted_new_data_numcols = len(self.col_positions) - 1 - selected_c
        else:
            adjusted_new_data_numcols = new_data_numcols
        if selected_r + new_data_numrows > len(self.row_positions) - 1:
            adjusted_new_data_numrows = len(self.row_positions) - 1 - selected_r
        else:
            adjusted_new_data_numrows = new_data_numrows
        selected_r_adjusted_new_data_numrows = selected_r + adjusted_new_data_numrows
        selected_c_adjusted_new_data_numcols = selected_c + adjusted_new_data_numcols
        endrow = selected_r_adjusted_new_data_numrows
        boxes = {
            (
                selected_r,
                selected_c,
                selected_r_adjusted_new_data_numrows,
                selected_c_adjusted_new_data_numcols,
            ): "cells"
        }
        event_data["selection_boxes"] = boxes
        for ndr, r in enumerate(range(selected_r, selected_r_adjusted_new_data_numrows)):
            datarn = self.datarn(r)
            for ndc, c in enumerate(range(selected_c, selected_c_adjusted_new_data_numcols)):
                event_data["data"][(datarn, self.datacn(c))] = data[ndr][ndc]

        if not try_binding(self.extra_begin_ctrl_v_func, event_data, "begin_ctrl_v"):
            return
        # the order of actions here is important:
        # edit existing sheet (not including any added rows/columns)

        # then if there are any added rows/columns:
        # create empty rows/columns dicts for any added rows/columns
        # edit those dicts with so far unused cells of data from clipboard
        # instead of editing table using set cell data, add any new rows then columns with pasted data
        for ndr, r in enumerate(range(selected_r, selected_r_adjusted_new_data_numrows)):
            for ndc, c in enumerate(range(selected_c, selected_c_adjusted_new_data_numcols)):
                val = data[ndr][ndc]
                if (
                    not self.edit_validation_func
                    or not validation
                    or (
                        self.edit_validation_func
                        and (val := self.edit_validation_func(mod_event_val(event_data, val, (r, c)))) is not None
                    )
                ):
                    event_data = self.event_data_set_cell(
                        datarn=self.datarn(r),
                        datacn=self.datacn(c),
                        value=val,
                        event_data=event_data,
                    )
        if added_rows:
            ctr = 0
            data_ins_row = len(self.data)
            displayed_ins_row = len(self.row_positions) - 1
            if total_data_cols is None:
                total_data_cols = self.total_data_cols()
            rows, index, row_heights = self.get_args_for_add_rows(
                data_ins_row=data_ins_row,
                displayed_ins_row=displayed_ins_row,
                rows=added_rows,
                total_data_cols=total_data_cols,
            )
            for ndr, r in zip(
                range(
                    adjusted_new_data_numrows,
                    new_data_numrows,
                ),
                rows,
            ):
                for ndc, c in enumerate(
                    range(
                        selected_c,
                        selected_c_adjusted_new_data_numcols,
                    )
                ):
                    val = data[ndr][ndc]
                    datacn = self.datacn(c)
                    if (
                        not self.edit_validation_func
                        or not validation
                        or (
                            self.edit_validation_func
                            and (val := self.edit_validation_func(mod_event_val(event_data, val, (r, c)))) is not None
                            and self.input_valid_for_cell(r, datacn, val, ignore_empty=True)
                        )
                    ):
                        rows[r][datacn] = self.format_value(r, datacn, val)
                        ctr += 1
            if ctr:
                event_data = self.add_rows(
                    rows=rows,
                    index=index if isinstance(self._row_index, list) and self._row_index else {},
                    row_heights=row_heights,
                    event_data=event_data,
                    mod_event_boxes=False,
                )
        if added_cols:
            ctr = 0
            if total_data_cols is None:
                total_data_cols = self.total_data_cols()
            data_ins_col = total_data_cols
            displayed_ins_col = len(self.col_positions) - 1
            columns, headers, column_widths = self.get_args_for_add_columns(
                data_ins_col=data_ins_col,
                displayed_ins_col=displayed_ins_col,
                columns=added_cols,
            )
            # only add the extra rows if expand_y is allowed
            if self.PAR.ops.paste_can_expand_x and self.PAR.ops.paste_can_expand_y:
                endrow = selected_r + new_data_numrows
            else:
                endrow = selected_r + adjusted_new_data_numrows
            for ndr, r in enumerate(
                range(
                    selected_r,
                    endrow,
                )
            ):
                for ndc, c in zip(
                    range(
                        adjusted_new_data_numcols,
                        new_data_numcols,
                    ),
                    columns,
                ):
                    val = data[ndr][ndc]
                    datarn = self.datarn(r)
                    if (
                        not self.edit_validation_func
                        or not validation
                        or (
                            self.edit_validation_func
                            and (val := self.edit_validation_func(mod_event_val(event_data, val, (r, c)))) is not None
                            and self.input_valid_for_cell(datarn, c, val, ignore_empty=True)
                        )
                    ):
                        columns[c][datarn] = self.format_value(datarn, c, val)
                        ctr += 1
            if ctr:
                event_data = self.add_columns(
                    columns=columns,
                    header=headers if isinstance(self._headers, list) and self._headers else {},
                    column_widths=column_widths,
                    event_data=event_data,
                    mod_event_boxes=False,
                )
        selboxr = selected_r + new_data_numrows if added_rows else selected_r_adjusted_new_data_numrows
        selboxc = selected_c + new_data_numcols if added_cols else selected_c_adjusted_new_data_numcols
        event_data = self.bulk_edit_validation(event_data)
        if event_data["cells"]["table"] or event_data["added"]["rows"] or event_data["added"]["columns"]:
            self.deselect("all", redraw=False)
            self.set_currently_selected(
                *curr_coords,
                item=self.create_selection_box(
                    selected_r,
                    selected_c,
                    selboxr,
                    selboxc,
                    type_="cells",
                    set_current=False,
                    run_binding=True,
                ),
            )
            event_data["selection_boxes"] = self.get_boxes()
            event_data["selected"] = self.selected
            self.see(selected_r, selected_c, redraw=False)
            self.refresh()
            if self.undo_enabled:
                self.undo_stack.append(stored_event_dict(event_data))
            try_binding(self.extra_end_ctrl_v_func, event_data, "end_ctrl_v")
            self.sheet_modified(event_data)
            self.PAR.emit_event("<<Paste>>", event_data)
            self.show_ctrl_outline("table", (selected_c, selected_r), (selboxc, selboxr), dash=())
        return event_data

    def delete_key(self, event: Any = None, validation: bool = True) -> None | EventDataDict:
        if not self.selected:
            return
        event_data = self.new_event_dict("edit_table")
        boxes = self.get_boxes()
        event_data["selection_boxes"] = boxes
        if not try_binding(self.extra_begin_delete_key_func, event_data, "begin_delete"):
            return
        for box in boxes:
            for r, c in gen_coords(*box):
                datarn, datacn = self.datarn(r), self.datacn(c)
                val = self.get_value_for_empty_cell(datarn, datacn)
                if (
                    not self.edit_validation_func
                    or not validation
                    or (
                        self.edit_validation_func
                        and (val := self.edit_validation_func(mod_event_val(event_data, val, (r, c)))) is not None
                    )
                ):
                    event_data = self.event_data_set_cell(
                        datarn,
                        datacn,
                        val,
                        event_data,
                    )
        event_data = self.bulk_edit_validation(event_data)
        if event_data["cells"]["table"]:
            self.refresh()
            if self.undo_enabled:
                self.undo_stack.append(stored_event_dict(event_data))
            try_binding(self.extra_end_delete_key_func, event_data, "end_delete")
            self.sheet_modified(event_data)
            self.PAR.emit_event("<<Delete>>", event_data)
            for r1, c1, r2, c2 in boxes:
                self.show_ctrl_outline(canvas="table", start_cell=(c1, r1), end_cell=(c2, r2))
        return event_data

    def event_data_set_cell(self, datarn: int, datacn: int, value: Any, event_data: EventDataDict) -> EventDataDict:
        """If bulk_table_edit_validation_func -> only updates event_data.data"""
        if self.input_valid_for_cell(datarn, datacn, value):
            if self.bulk_table_edit_validation_func:
                event_data["data"][(datarn, datacn)] = value
            else:
                event_data["cells"]["table"][(datarn, datacn)] = self.get_cell_data(datarn, datacn)
                self.set_cell_data(datarn, datacn, value)
        return event_data

    def bulk_edit_validation(self, event_data: EventDataDict) -> EventDataDict:
        if self.bulk_table_edit_validation_func:
            self.bulk_table_edit_validation_func(event_data)
            for (datarn, datacn), value in event_data["data"].items():
                if self.input_valid_for_cell(datarn, datacn, value):
                    event_data["cells"]["table"][(datarn, datacn)] = self.get_cell_data(datarn, datacn)
                    self.set_cell_data(datarn, datacn, value)
        return event_data

    def get_args_for_move_columns(
        self,
        move_to: int,
        to_move: list[int],
        data_indexes: bool = False,
    ) -> tuple[dict[int, int], dict[int, int], int, dict[int, int]]:
        if not data_indexes or self.all_columns_displayed:
            disp_new_idxs = get_new_indexes(
                move_to=move_to,
                to_move=to_move,
            )
            data_new_idxs = dict(disp_new_idxs)
        else:
            disp_new_idxs = {}
            data_new_idxs = get_new_indexes(
                move_to=move_to,
                to_move=to_move,
            )
        # at_least_cols should not be len in this case as move_to can be len
        fix_len = (move_to - 1) if move_to else move_to
        if not self.all_columns_displayed and not data_indexes:
            fix_len = self.datacn(fix_len)
        totalcols = self.equalize_data_row_lengths(at_least_cols=fix_len)
        if not self.all_columns_displayed and not data_indexes:
            keep = set(map(self.datacn, to_move))
            data_new_idxs = {
                k: v
                for k, v in zip(
                    move_elements_by_mapping_gen(
                        self.displayed_columns,
                        data_new_idxs,
                        dict(zip(data_new_idxs.values(), data_new_idxs)),
                    ),
                    self.displayed_columns,
                )
                if k in keep
            }
        return data_new_idxs, dict(zip(data_new_idxs.values(), data_new_idxs)), totalcols, disp_new_idxs

    def move_columns_adjust_options_dict(
        self,
        data_new_idxs: dict[int, int],
        data_old_idxs: dict[int, int],
        totalcols: int | None,
        disp_new_idxs: None | dict[int, int] = None,
        move_data: bool = True,
        move_widths: bool = True,
        create_selections: bool = True,
        event_data: EventDataDict | None = None,
    ) -> tuple[dict[int, int], dict[int, int], EventDataDict]:
        self.saved_column_widths = {}
        if not event_data:
            event_data = self.new_event_dict("move_columns", state=True)
        if not event_data["moved"]["columns"]:
            event_data["moved"]["columns"] = {
                "data": data_new_idxs,
                "displayed": {} if disp_new_idxs is None else disp_new_idxs,
            }
        event_data["options"] = self.copy_options()
        event_data["named_spans"] = {k: span.copy_self() for k, span in self.named_spans.items()}

        if move_widths and disp_new_idxs:
            self.set_col_positions(
                itr=move_elements_by_mapping_gen(
                    self.get_column_widths(),
                    disp_new_idxs,
                    dict(
                        zip(
                            disp_new_idxs.values(),
                            disp_new_idxs,
                        )
                    ),
                )
            )
            if create_selections:
                self.deselect("all", run_binding=False, redraw=False)
                for boxst, boxend in consecutive_ranges(sorted(disp_new_idxs.values())):
                    self.create_selection_box(
                        0,
                        boxst,
                        len(self.row_positions) - 1,
                        boxend,
                        "columns",
                        run_binding=True,
                    )
            else:
                self.recreate_all_selection_boxes()

        if move_data:
            if not isinstance(totalcols, int):
                totalcols = max(data_new_idxs.values(), default=0)
                if totalcols:
                    totalcols += 1
                totalcols = self.equalize_data_row_lengths(at_least_cols=totalcols)

            self.data = [move_fast(k, data_new_idxs, data_old_idxs) for k in self.data]
            maxidx = len_to_idx(totalcols)
            self.CH.fix_header(maxidx)
            if isinstance(self._headers, list) and self._headers:
                self._headers = move_fast(self._headers, data_new_idxs, data_old_idxs)
            maxidx = self.get_max_column_idx(maxidx)
            full_new_idxs, full_old_idxs = self.get_full_new_idxs(
                max_idx=maxidx,
                new_idxs=data_new_idxs,
                old_idxs=data_old_idxs,
                get_inverse=True,
            )
            self.tagged_cells = {
                tags: {(k[0], full_new_idxs[k[1]]) for k in tagged} for tags, tagged in self.tagged_cells.items()
            }
            self.cell_options = {(k[0], full_new_idxs[k[1]]): v for k, v in self.cell_options.items()}
            self.progress_bars = {(k[0], full_new_idxs[k[1]]): v for k, v in self.progress_bars.items()}
            self.col_options = {full_new_idxs[k]: v for k, v in self.col_options.items()}
            self.tagged_columns = {
                tags: {full_new_idxs[k] for k in tagged} for tags, tagged in self.tagged_columns.items()
            }
            self.CH.cell_options = {full_new_idxs[k]: v for k, v in self.CH.cell_options.items()}
            self.displayed_columns = sorted(full_new_idxs[k] for k in self.displayed_columns)
            if self.named_spans:
                totalrows = self.total_data_rows()
                new_ops = self.PAR.create_options_from_span
                qkspan = self.span()
                for span in self.named_spans.values():
                    # span is neither a cell options nor col options span, continue
                    if not isinstance(span["from_c"], int):
                        continue
                    oldupto_colrange, newupto_colrange, newfrom, newupto = span_idxs_post_move(
                        data_new_idxs,
                        full_new_idxs,
                        totalcols,
                        span,
                        "c",
                    )
                    # add cell/col kwargs for columns that are new to the span
                    old_span_idxs = {full_new_idxs[k] for k in range(span["from_c"], oldupto_colrange)}
                    for k in range(newfrom, newupto_colrange):
                        if k not in old_span_idxs:
                            oldidx = full_old_idxs[k]
                            # event_data is used to preserve old cell value
                            # in case cells are modified by
                            # formatting, checkboxes, dropdown boxes
                            if (
                                span["type_"] in val_modifying_options
                                and span["header"]
                                and oldidx not in event_data["cells"]["header"]
                            ):
                                event_data["cells"]["header"][oldidx] = self.CH.get_cell_data(k)
                            # the span targets columns
                            if span["from_r"] is None:
                                if span["type_"] in val_modifying_options:
                                    for datarn in range(len(self.data)):
                                        if (datarn, oldidx) not in event_data["cells"]["table"]:
                                            event_data["cells"]["table"][(datarn, oldidx)] = self.get_cell_data(
                                                datarn, k
                                            )
                                # create new col options
                                new_ops(
                                    mod_span(
                                        qkspan,
                                        span,
                                        from_c=k,
                                        upto_c=k + 1,
                                    )
                                )
                            # the span targets cells
                            else:
                                rng_upto_r = totalrows if span["upto_r"] is None else span["upto_r"]
                                for datarn in range(span["from_r"], rng_upto_r):
                                    if (
                                        span["type_"] in val_modifying_options
                                        and (datarn, oldidx) not in event_data["cells"]["table"]
                                    ):
                                        event_data["cells"]["table"][(datarn, oldidx)] = self.get_cell_data(datarn, k)
                                    # create new cell options
                                    new_ops(
                                        mod_span(
                                            qkspan,
                                            span,
                                            from_r=datarn,
                                            upto_r=datarn + 1,
                                            from_c=k,
                                            upto_c=k + 1,
                                        )
                                    )
                    # remove span specific kwargs from cells/columns
                    # that are no longer in the span,
                    # cell options/col options keys are new idxs
                    for k in range(span["from_c"], oldupto_colrange):
                        # has it moved outside of new span coords
                        if (
                            isinstance(newupto, int) and (full_new_idxs[k] < newfrom or full_new_idxs[k] >= newupto)
                        ) or (newupto is None and full_new_idxs[k] < newfrom):
                            # span includes header
                            if (
                                span["header"]
                                and full_new_idxs[k] in self.CH.cell_options
                                and span["type_"] in self.CH.cell_options[full_new_idxs[k]]
                            ):
                                del self.CH.cell_options[full_new_idxs[k]][span["type_"]]
                            # span is for col options
                            if span["from_r"] is None:
                                if (
                                    full_new_idxs[k] in self.col_options
                                    and span["type_"] in self.col_options[full_new_idxs[k]]
                                ):
                                    del self.col_options[full_new_idxs[k]][span["type_"]]
                            # span is for cell options
                            else:
                                rng_upto_r = totalrows if span["upto_r"] is None else span["upto_r"]
                                for r in range(span["from_r"], rng_upto_r):
                                    if (r, full_new_idxs[k]) in self.cell_options and span[
                                        "type_"
                                    ] in self.cell_options[(r, full_new_idxs[k])]:
                                        del self.cell_options[(r, full_new_idxs[k])][span["type_"]]
                    # finally, change the span coords
                    span["from_c"], span["upto_c"] = newfrom, newupto
        return data_new_idxs, disp_new_idxs, event_data

    def get_max_column_idx(self, maxidx: int | None = None) -> int:
        if maxidx is None:
            maxidx = len_to_idx(self.total_data_cols())
        maxiget = partial(max, key=itemgetter(1))
        return max(
            max(self.cell_options, key=itemgetter(1), default=(0, maxidx))[1],
            max(self.col_options, default=maxidx),
            max(self.CH.cell_options, default=maxidx),
            maxiget(map(maxiget, self.tagged_cells.values()), default=(0, maxidx))[1],
            max(map(max, self.tagged_columns.values()), default=maxidx),
            max((d.from_c for d in self.named_spans.values() if isinstance(d.from_c, int)), default=maxidx),
            max((d.upto_c for d in self.named_spans.values() if isinstance(d.upto_c, int)), default=maxidx),
            self.displayed_columns[-1] if self.displayed_columns else maxidx,
        )

    def get_args_for_move_rows(
        self,
        move_to: int,
        to_move: list[int],
        data_indexes: bool = False,
    ) -> tuple[dict[int, int], dict[int, int], int, dict[int, int]]:
        if not data_indexes or self.all_rows_displayed:
            disp_new_idxs = get_new_indexes(
                move_to=move_to,
                to_move=to_move,
            )
            data_new_idxs = dict(disp_new_idxs)
        else:
            disp_new_idxs = {}
            data_new_idxs = get_new_indexes(
                move_to=move_to,
                to_move=to_move,
            )
        # move_to can be len and fix_data_len() takes index so - 1
        fix_len = (move_to - 1) if move_to else move_to
        if not self.all_rows_displayed and not data_indexes:
            fix_len = self.datarn(fix_len)
        self.fix_data_len(fix_len)
        totalrows = max(self.total_data_rows(), len(self.row_positions) - 1)
        if not self.all_rows_displayed and not data_indexes:
            keep = set(map(self.datarn, to_move))
            data_new_idxs = {
                k: v
                for k, v in zip(
                    move_elements_by_mapping_gen(
                        self.displayed_rows,
                        data_new_idxs,
                        dict(zip(data_new_idxs.values(), data_new_idxs)),
                    ),
                    self.displayed_rows,
                )
                if k in keep
            }
        return data_new_idxs, dict(zip(data_new_idxs.values(), data_new_idxs)), totalrows, disp_new_idxs

    def move_rows_data(
        self,
        data_new_idxs: dict[int, int],
        data_old_idxs: dict[int, int],
        maxidx: int,
    ) -> None:
        self.data = move_fast(
            self.data,
            data_new_idxs,
            data_old_idxs,
        )
        self.RI.fix_index(maxidx)
        if isinstance(self._row_index, list) and self._row_index:
            self._row_index = move_fast(self._row_index, data_new_idxs, data_old_idxs)

    def move_rows_adjust_options_dict(
        self,
        data_new_idxs: dict[int, int],
        data_old_idxs: dict[int, int],
        totalrows: int | None,
        disp_new_idxs: None | dict[int, int] = None,
        move_data: bool = True,
        move_heights: bool = True,
        create_selections: bool = True,
        event_data: EventDataDict | None = None,
        undo_modification: EventDataDict | None = None,
        node_change: None | dict = None,
        manage_tree: bool = True,
    ) -> tuple[dict[int, int], dict[int, int], EventDataDict]:
        self.saved_row_heights = {}
        if not isinstance(totalrows, int):
            totalrows = max(
                self.total_data_rows(),
                len(self.row_positions) - 1,
                max(data_new_idxs.values(), default=0),
            )
            self.fix_data_len(totalrows - 1)
        if not event_data:
            event_data = self.new_event_dict("move_rows", state=True)
        if not event_data["moved"]["rows"]:
            event_data["moved"]["rows"] = {
                "data": data_new_idxs,
                "displayed": {} if disp_new_idxs is None else disp_new_idxs,
            }
        event_data["options"] = self.copy_options()
        event_data["named_spans"] = {k: span.copy_self() for k, span in self.named_spans.items()}

        if move_data:
            maxidx = len_to_idx(totalrows)
            if manage_tree and self.PAR.ops.treeview:
                two_step_move = self.RI.move_rows_mod_nodes(
                    data_new_idxs=data_new_idxs,
                    data_old_idxs=data_old_idxs,
                    disp_new_idxs=disp_new_idxs,
                    maxidx=maxidx,
                    event_data=event_data,
                    # undo_modification is the old saved stuff
                    # event_data is the new event
                    # node change only comes from Sheet.move()
                    undo_modification=undo_modification,
                    node_change=node_change,
                )
                data_new_idxs, data_old_idxs, disp_new_idxs, event_data = next(two_step_move)
                if not data_new_idxs and not disp_new_idxs:
                    return data_new_idxs, disp_new_idxs, event_data
            else:
                self.move_rows_data(data_new_idxs, data_old_idxs, maxidx)
            maxidx = self.get_max_row_idx(maxidx)
            full_new_idxs, full_old_idxs = self.get_full_new_idxs(
                max_idx=maxidx,
                new_idxs=data_new_idxs,
                old_idxs=data_old_idxs,
                get_inverse=True,
            )
            self.tagged_cells = {
                tags: {(full_new_idxs[k[0]], k[1]) for k in tagged} for tags, tagged in self.tagged_cells.items()
            }
            self.cell_options = {(full_new_idxs[k[0]], k[1]): v for k, v in self.cell_options.items()}
            self.progress_bars = {(full_new_idxs[k[0]], k[1]): v for k, v in self.progress_bars.items()}
            self.tagged_rows = {tags: {full_new_idxs[k] for k in tagged} for tags, tagged in self.tagged_rows.items()}
            self.row_options = {full_new_idxs[k]: v for k, v in self.row_options.items()}
            self.RI.cell_options = {full_new_idxs[k]: v for k, v in self.RI.cell_options.items()}
            self.RI.rns = {v: full_new_idxs[k] for v, k in self.RI.rns.items()}
            self.displayed_rows = sorted(full_new_idxs[k] for k in self.displayed_rows)
            if manage_tree and self.PAR.ops.treeview:
                next(two_step_move)

            if self.named_spans:
                totalcols = self.total_data_cols()
                new_ops = self.PAR.create_options_from_span
                qkspan = self.span()
                for span in self.named_spans.values():
                    # span is neither a cell options nor row options span, continue
                    if not isinstance(span["from_r"], int):
                        continue
                    oldupto_rowrange, newupto_rowrange, newfrom, newupto = span_idxs_post_move(
                        data_new_idxs,
                        full_new_idxs,
                        totalrows,
                        span,
                        "r",
                    )
                    # add cell/row kwargs for rows that are new to the span
                    old_span_idxs = {full_new_idxs[k] for k in range(span["from_r"], oldupto_rowrange)}
                    for k in range(newfrom, newupto_rowrange):
                        if k not in old_span_idxs:
                            oldidx = full_old_idxs[k]
                            # event_data is used to preserve old cell value
                            # in case cells are modified by
                            # formatting, checkboxes, dropdown boxes
                            if (
                                span["type_"] in val_modifying_options
                                and span["index"]
                                and oldidx not in event_data["cells"]["index"]
                            ):
                                event_data["cells"]["index"][oldidx] = self.RI.get_cell_data(k)
                            # the span targets rows
                            if span["from_c"] is None:
                                if span["type_"] in val_modifying_options:
                                    for datacn in range(len(self.data[k])):
                                        if (oldidx, datacn) not in event_data["cells"]["table"]:
                                            event_data["cells"]["table"][(oldidx, datacn)] = self.get_cell_data(
                                                k, datacn
                                            )
                                # create new row options
                                new_ops(
                                    mod_span(
                                        qkspan,
                                        span,
                                        from_r=k,
                                        upto_r=k + 1,
                                    )
                                )
                            # the span targets cells
                            else:
                                rng_upto_c = totalcols if span["upto_c"] is None else span["upto_c"]
                                for datacn in range(span["from_c"], rng_upto_c):
                                    if (
                                        span["type_"] in val_modifying_options
                                        and (oldidx, datacn) not in event_data["cells"]["table"]
                                    ):
                                        event_data["cells"]["table"][(oldidx, datacn)] = self.get_cell_data(k, datacn)
                                    # create new cell options
                                    new_ops(
                                        mod_span(
                                            qkspan,
                                            span,
                                            from_r=k,
                                            upto_r=k + 1,
                                            from_c=datacn,
                                            upto_c=datacn + 1,
                                        )
                                    )
                    # remove span specific kwargs from cells/rows
                    # that are no longer in the span,
                    # cell options/row options keys are new idxs
                    for k in range(span["from_r"], oldupto_rowrange):
                        # has it moved outside of new span coords
                        if (
                            isinstance(newupto, int) and (full_new_idxs[k] < newfrom or full_new_idxs[k] >= newupto)
                        ) or (newupto is None and full_new_idxs[k] < newfrom):
                            # span includes index
                            if (
                                span["index"]
                                and full_new_idxs[k] in self.RI.cell_options
                                and span["type_"] in self.RI.cell_options[full_new_idxs[k]]
                            ):
                                del self.RI.cell_options[full_new_idxs[k]][span["type_"]]
                            # span is for row options
                            if span["from_c"] is None:
                                if (
                                    full_new_idxs[k] in self.row_options
                                    and span["type_"] in self.row_options[full_new_idxs[k]]
                                ):
                                    del self.row_options[full_new_idxs[k]][span["type_"]]
                            # span is for cell options
                            else:
                                rng_upto_c = totalcols if span["upto_c"] is None else span["upto_c"]
                                for c in range(span["from_c"], rng_upto_c):
                                    if (full_new_idxs[k], c) in self.cell_options and span[
                                        "type_"
                                    ] in self.cell_options[(full_new_idxs[k], c)]:
                                        del self.cell_options[(full_new_idxs[k], c)][span["type_"]]
                    # finally, change the span coords
                    span["from_r"], span["upto_r"] = newfrom, newupto

        if not undo_modification and create_selections and self.PAR.ops.treeview:
            self.PAR.selection_set(*[self._row_index[k].iid for k in data_new_idxs.values()])

        elif not undo_modification and move_heights and disp_new_idxs:
            self.set_row_positions(
                itr=move_elements_by_mapping_gen(
                    self.get_row_heights(),
                    disp_new_idxs,
                    dict(
                        zip(
                            disp_new_idxs.values(),
                            disp_new_idxs,
                        )
                    ),
                )
            )
            if create_selections:
                self.deselect("all", run_binding=False, redraw=False)
                for boxst, boxend in consecutive_ranges(sorted(disp_new_idxs.values())):
                    self.create_selection_box(
                        boxst,
                        0,
                        boxend,
                        len(self.col_positions) - 1,
                        "rows",
                        run_binding=True,
                    )
            else:
                self.recreate_all_selection_boxes()

        return data_new_idxs, disp_new_idxs, event_data

    def get_max_row_idx(self, maxidx: int | None = None) -> int:
        if maxidx is None:
            maxidx = len_to_idx(self.total_data_rows())
        maxiget = partial(max, key=itemgetter(0))
        return max(
            max(self.cell_options, key=itemgetter(0), default=(maxidx, 0))[0],
            max(self.row_options, default=maxidx),
            max(self.RI.cell_options, default=maxidx),
            maxiget(map(maxiget, self.tagged_cells.values()), default=(maxidx, 0))[0],
            max(map(max, self.tagged_rows.values()), default=maxidx),
            max((d.from_r for d in self.named_spans.values() if isinstance(d.from_r, int)), default=maxidx),
            max((d.upto_r for d in self.named_spans.values() if isinstance(d.upto_r, int)), default=maxidx),
            self.displayed_rows[-1] if self.displayed_rows else maxidx,
        )

    def get_full_new_idxs(
        self,
        max_idx: int,
        new_idxs: dict[int, int],
        old_idxs: None | dict[int, int] = None,
        get_inverse: bool = False,
    ) -> dict[int, int]:
        # return a dict of all row or column indexes
        # old indexes and new indexes, not just the
        # ones that were moved e.g.
        # {old index: new index, ...}
        # all the way from 0 to max_idx
        if old_idxs is None:
            old_idxs = {v: k for k, v in new_idxs.items()}
        if get_inverse:
            d = {}
            d_in = {}
            for v, k in enumerate(move_elements_by_mapping_gen(tuple(range(max_idx + 1)), new_idxs, old_idxs)):
                d[k] = v
                d_in[v] = k
            return d, d_in
        else:
            return {
                k: v for v, k in enumerate(move_elements_by_mapping_gen(tuple(range(max_idx + 1)), new_idxs, old_idxs))
            }

    def undo(self, event: Any = None) -> None | EventDataDict:
        if not self.undo_stack:
            return
        modification = self.undo_stack[-1]["data"]
        if not try_binding(self.extra_begin_ctrl_z_func, modification, "begin_undo"):
            return
        event_data = self.undo_modification_invert_event(modification)
        self.redo_stack.append(stored_event_dict(event_data))
        self.undo_stack.pop()
        self.sheet_modified(event_data, purge_redo=False)
        try_binding(self.extra_end_ctrl_z_func, event_data, "end_undo")
        self.PAR.emit_event("<<Undo>>", event_data)
        return event_data

    def redo(self, event: Any = None) -> None | EventDataDict:
        if not self.redo_stack:
            return
        modification = self.redo_stack[-1]["data"]
        if not try_binding(self.extra_begin_ctrl_z_func, modification, "begin_redo"):
            return
        event_data = self.undo_modification_invert_event(modification, name="redo")
        self.undo_stack.append(stored_event_dict(event_data))
        self.redo_stack.pop()
        self.sheet_modified(event_data, purge_redo=False)
        try_binding(self.extra_end_ctrl_z_func, event_data, "end_redo")
        self.PAR.emit_event("<<Redo>>", event_data)
        return event_data

    def sheet_modified(self, event_data: EventDataDict, purge_redo: bool = True, emit_event: bool = True) -> None:
        if emit_event:
            self.PAR.emit_event("<<SheetModified>>", event_data)
        if purge_redo:
            self.purge_redo_stack()

    def edit_cells_using_modification(self, modification: dict, event_data: dict) -> EventDataDict:
        treeview = self.PAR.ops.treeview
        for datarn, v in modification["cells"]["index"].items():
            if not self.edit_validation_func or (
                self.edit_validation_func
                and (v := self.edit_validation_func(mod_event_val(event_data, v, row=datarn))) is not None
            ):
                if treeview:
                    self._row_index[datarn].text = v
                else:
                    self._row_index[datarn] = v
        for datacn, v in modification["cells"]["header"].items():
            if not self.edit_validation_func or (
                self.edit_validation_func
                and (v := self.edit_validation_func(mod_event_val(event_data, v, column=datacn))) is not None
            ):
                self._headers[datacn] = v
        for k, v in modification["cells"]["table"].items():
            if not self.edit_validation_func or (
                self.edit_validation_func
                and (v := self.edit_validation_func(mod_event_val(event_data, v, loc=k))) is not None
            ):
                event_data = self.event_data_set_cell(k[0], k[1], v, event_data)
        return event_data

    def save_cells_using_modification(self, modification: EventDataDict, event_data: EventDataDict) -> EventDataDict:
        for datarn in modification["cells"]["index"]:
            event_data["cells"]["index"][datarn] = self.RI.get_cell_data(datarn)

        for datacn in modification["cells"]["header"]:
            event_data["cells"]["header"][datacn] = self.CH.get_cell_data(datacn)

        for k in modification["cells"]["table"]:
            event_data["cells"]["table"][k] = self.get_cell_data(k[0], k[1])
        return event_data

    def restore_sheet_state(self, modification: EventDataDict) -> None:
        if "cell_options" in modification["options"]:
            self.cell_options = modification["options"]["cell_options"]
        if "column_options" in modification["options"]:
            self.col_options = modification["options"]["column_options"]
        if "row_options" in modification["options"]:
            self.row_options = modification["options"]["row_options"]
        if "CH_cell_options" in modification["options"]:
            self.CH.cell_options = modification["options"]["CH_cell_options"]
        if "RI_cell_options" in modification["options"]:
            self.RI.cell_options = modification["options"]["RI_cell_options"]
        if "tagged_cells" in modification["options"]:
            self.tagged_cells = modification["options"]["tagged_cells"]
        if "tagged_rows" in modification["options"]:
            self.tagged_rows = modification["options"]["tagged_rows"]
        if "tagged_columns" in modification["options"]:
            self.tagged_columns = modification["options"]["tagged_columns"]
        if modification["named_spans"]:
            self.named_spans = modification["named_spans"]
        if modification["sheet_state"]:
            self.RI.tree_open_ids = modification["sheet_state"]["tree_open_ids"]
            self.row_positions = modification["sheet_state"]["row_positions"]
            self.col_positions = modification["sheet_state"]["col_positions"]
            self.displayed_rows = modification["sheet_state"]["displayed_rows"]
            self.displayed_columns = modification["sheet_state"]["displayed_columns"]
            self.all_rows_displayed = modification["sheet_state"]["all_rows_displayed"]
            self.all_columns_displayed = modification["sheet_state"]["all_columns_displayed"]
            self.saved_row_heights = modification["sheet_state"]["saved_row_heights"]
            self.saved_column_widths = modification["sheet_state"]["saved_column_widths"]
            self.recreate_all_selection_boxes()

    def undo_modification_invert_event(self, modification: EventDataDict, name: str = "undo") -> EventDataDict:
        self.deselect("all", redraw=False)
        event_data = self.new_event_dict(modification["eventname"], state=True)
        event_data["selection_boxes"] = modification["selection_boxes"]
        event_data["selected"] = modification["selected"]
        saved_cells = False

        if modification["added"]["rows"] or modification["added"]["columns"]:
            event_data = self.save_cells_using_modification(modification, event_data)
            saved_cells = True

        if modification["treeview"]["renamed"]:
            for old, new in modification["treeview"]["renamed"].items():
                event_data = self.RI.rename_iid(new, old, event_data=event_data)

        if modification["moved"]["columns"]:
            totalcols = max(self.equalize_data_row_lengths(), max(modification["moved"]["columns"]["data"].values()))
            data_new_idxs, disp_new_idxs, event_data = self.move_columns_adjust_options_dict(
                data_new_idxs=dict(
                    zip(
                        modification["moved"]["columns"]["data"].values(),
                        modification["moved"]["columns"]["data"],
                    )
                ),
                data_old_idxs=modification["moved"]["columns"]["data"],
                totalcols=totalcols,
                disp_new_idxs=dict(
                    zip(
                        modification["moved"]["columns"]["displayed"].values(),
                        modification["moved"]["columns"]["displayed"],
                    )
                ),
                event_data=event_data,
            )
            event_data["moved"]["columns"] = {
                "data": data_new_idxs,
                "displayed": disp_new_idxs,
            }

        if modification["moved"]["rows"]:
            totalrows = max(self.total_data_rows(), max(modification["moved"]["rows"]["data"].values()))
            data_new_idxs, disp_new_idxs, event_data = self.move_rows_adjust_options_dict(
                data_new_idxs=dict(
                    zip(
                        modification["moved"]["rows"]["data"].values(),
                        modification["moved"]["rows"]["data"],
                    )
                ),
                data_old_idxs=modification["moved"]["rows"]["data"],
                totalrows=totalrows,
                disp_new_idxs=dict(
                    zip(
                        modification["moved"]["rows"]["displayed"].values(),
                        modification["moved"]["rows"]["displayed"],
                    )
                ),
                event_data=event_data,
                undo_modification=modification,
            )
            event_data["moved"]["rows"] = {
                "data": data_new_idxs,
                "displayed": disp_new_idxs,
            }

        if modification["added"]["rows"]:
            self.deselect("all", run_binding=False, redraw=False)
            event_data = self.delete_rows_data(
                rows=tuple(modification["added"]["rows"]["table"]),
                event_data=event_data,
            )
            event_data = self.delete_rows_displayed(
                rows=tuple(modification["added"]["rows"]["row_heights"]),
                event_data=event_data,
                from_undo=True,
            )

        if modification["added"]["columns"]:
            self.deselect("all", run_binding=False, redraw=False)
            event_data = self.delete_columns_data(
                cols=tuple(modification["added"]["columns"]["table"]),
                event_data=event_data,
            )
            event_data = self.delete_columns_displayed(
                cols=tuple(modification["added"]["columns"]["column_widths"]),
                event_data=event_data,
                from_undo=True,
            )

        if modification["deleted"]["rows"] or modification["deleted"]["row_heights"]:
            event_data["treeview"] = modification["treeview"]
            self.add_rows(
                rows=modification["deleted"]["rows"],
                index=modification["deleted"]["index"],
                row_heights=modification["deleted"]["row_heights"],
                event_data=event_data,
                create_ops=False,
                create_selections=False,
                add_col_positions=False,
                from_undo=True,
            )

        if modification["deleted"]["columns"] or modification["deleted"]["column_widths"]:
            self.add_columns(
                columns=modification["deleted"]["columns"],
                header=modification["deleted"]["header"],
                column_widths=modification["deleted"]["column_widths"],
                event_data=event_data,
                create_ops=False,
                create_selections=False,
                add_row_positions=False,
                from_undo=True,
            )

        if modification["eventname"].startswith(("edit", "move")):
            if not saved_cells:
                event_data = self.save_cells_using_modification(modification, event_data)
            event_data = self.edit_cells_using_modification(modification, event_data)
            event_data = self.bulk_edit_validation(event_data)

        elif modification["eventname"].startswith("add"):
            event_data["eventname"] = modification["eventname"].replace("add", "delete")

        elif modification["eventname"].startswith("delete"):
            event_data["eventname"] = modification["eventname"].replace("delete", "add")

        self.restore_sheet_state(modification)

        if not modification["eventname"].startswith("move") and (
            len(self.row_positions) > 1 or len(self.col_positions) > 1
        ):
            self.deselect("all", redraw=False)
            self.reselect_from_get_boxes(
                modification["selection_boxes"],
                modification["selected"],
            )

        if self.selected:
            self.see(self.selected.row, self.selected.column, redraw=False)

        self.refresh()
        return event_data

    def see(
        self,
        r: int | None = None,
        c: int | None = None,
        keep_yscroll: bool = False,
        keep_xscroll: bool = False,
        bottom_right_corner: bool | None = None,
        check_cell_visibility: bool = True,
        redraw: bool = True,
        r_pc: float = 0.0,
        c_pc: float = 0.0,
        index: bool = True,
    ) -> bool:
        need_y_redraw = False
        need_x_redraw = False
        vis_info = self.cell_visibility_info(r, c)
        yvis, xvis = vis_info["yvis"], vis_info["xvis"]
        top_left_x, top_left_y, bottom_right_x, bottom_right_y = vis_info["visible_region"]
        x1, y1, x2, y2 = vis_info["cell_coords"]
        # y scroll
        if not check_cell_visibility or (check_cell_visibility and not yvis) and len(self.row_positions) > 1:
            if bottom_right_corner is None:
                both_above = y1 < top_left_y and y2 < top_left_y
                y1_above_y2_below = y1 < top_left_y and y2 > bottom_right_y
                y1_above_y2_visible = y1 < top_left_y and top_left_y <= y2 <= bottom_right_y
                brc = not (both_above or y1_above_y2_below or y1_above_y2_visible)
            else:
                brc = bottom_right_corner
            if brc:
                if r is not None and not keep_yscroll:
                    winfo_height = self.winfo_height()
                    if self.row_positions[r + 1] - self.row_positions[r] > winfo_height:
                        y = self.row_positions[r]
                    else:
                        y = self.row_positions[r + 1] + 1 - winfo_height
                    y = y / (self.row_positions[-1] + self.PAR.ops.empty_vertical)
                    args = [
                        "moveto",
                        y - 1 if y > 1 else y,
                    ]
                    self.set_yviews(*args, redraw=False)
                    need_y_redraw = True
            else:
                if r is not None and not keep_yscroll:
                    y = max(
                        0, int(self.row_positions[r] + ((self.row_positions[r + 1] - self.row_positions[r]) * r_pc)) - 2
                    )
                    y = y / (self.row_positions[-1] + self.PAR.ops.empty_vertical)
                    args = [
                        "moveto",
                        y - 1 if y > 1 else y,
                    ]
                    self.set_yviews(*args, redraw=False)
                    need_y_redraw = True
        # x scroll
        if not check_cell_visibility or (check_cell_visibility and not xvis) and len(self.col_positions) > 1:
            if bottom_right_corner is None:
                both_left = x1 < top_left_x and x2 < top_left_x
                x1_left_x2_right = x1 < top_left_x and x2 > bottom_right_x
                x1_left_x2_visible = x1 < top_left_x and top_left_x <= x2 <= bottom_right_x
                brc = not (both_left or x1_left_x2_right or x1_left_x2_visible)
            else:
                brc = bottom_right_corner
            if brc:
                if c is not None and not keep_xscroll:
                    winfo_width = self.winfo_width()
                    if self.col_positions[c + 1] - self.col_positions[c] > winfo_width:
                        x = self.col_positions[c]
                    else:
                        x = self.col_positions[c + 1] + 1 - winfo_width
                    x = x / (self.col_positions[-1] + self.PAR.ops.empty_horizontal)
                    args = [
                        "moveto",
                        x - 1 if x > 1 else x,
                    ]
                    self.set_xviews(*args, redraw=False)
                    need_x_redraw = True
            else:
                if c is not None and not keep_xscroll:
                    x = max(
                        0, int(self.col_positions[c] + ((self.col_positions[c + 1] - self.col_positions[c]) * c_pc)) - 2
                    )
                    x = x / (self.col_positions[-1] + self.PAR.ops.empty_horizontal)
                    args = [
                        "moveto",
                        x - 1 if x > 1 else x,
                    ]
                    self.set_xviews(*args, redraw=False)
                    need_x_redraw = True
            # the index may have resized after scrolling making x calculation wrong
            if need_x_redraw and index and self.PAR.ops.auto_resize_row_index and self.show_index:
                self.main_table_redraw_grid_and_text(redraw_header=False, redraw_row_index=False, redraw_table=False)
                self.see(
                    r=r,
                    c=c,
                    keep_yscroll=keep_yscroll,
                    keep_xscroll=keep_xscroll,
                    check_cell_visibility=check_cell_visibility,
                    redraw=redraw,
                    r_pc=r_pc,
                    c_pc=c_pc,
                    index=False,
                )
        if (need_y_redraw or need_x_redraw) and self.find_window.open:
            self.reposition_find_window()  # prevent it from appearing to move around
        if redraw and (need_y_redraw or need_x_redraw):
            self.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
            return True
        return False

    def get_cell_coords(self, r: int | None = None, c: int | None = None) -> tuple[float, float, float, float]:
        return (
            0 if not c else self.col_positions[c] + 1,
            0 if not r else self.row_positions[r] + 1,
            0 if not c else self.col_positions[c + 1],
            0 if not r else self.row_positions[r + 1],
        )

    def cell_completely_visible(self, r: int | None = 0, c: int | None = 0) -> bool:
        cx1, cy1, cx2, cy2 = self.get_canvas_visible_area()
        x1, y1, x2, y2 = self.get_cell_coords(r, c)
        return cy1 <= y1 and cy2 >= y2 and cx1 <= x1 and cx2 >= x2

    def cell_visibility_info(
        self,
        r: int,
        c: int,
    ) -> dict[Literal["yvis", "xvis", "visible_region", "cell_coords"], bool | tuple[int, int, int, int]]:
        cx1, cy1, cx2, cy2 = self.get_canvas_visible_area()
        x1, y1, x2, y2 = self.get_cell_coords(r, c)
        return {
            "yvis": cy1 <= y1 and cy2 >= y2,
            "xvis": cx1 <= x1 and cx2 >= x2,
            "visible_region": (cx1, cy1, cx2, cy2),
            "cell_coords": (x1, y1, x2, y2),
        }

    def cell_visible(self, r: int = 0, c: int = 0) -> bool:
        cx1, cy1, cx2, cy2 = self.get_canvas_visible_area()
        x1, y1, x2, y2 = self.get_cell_coords(r, c)
        return x1 <= cx2 or y1 <= cy2 or x2 >= cx1 or y2 >= cy1

    def select_all(self, redraw: bool = True, run_binding_func: bool = True) -> None:
        selected = self.selected
        self.deselect("all", redraw=False)
        if len(self.row_positions) > 1 and len(self.col_positions) > 1:
            item = self.create_selection_box(
                0,
                0,
                len(self.row_positions) - 1,
                len(self.col_positions) - 1,
                set_current=False,
            )
            if selected:
                self.set_currently_selected(selected.row, selected.column, item=item)
            else:
                self.set_currently_selected(0, 0, item=item)
            if redraw:
                self.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
            if run_binding_func:
                event_data = self.get_select_event(self.being_drawn_item)
                try_binding(self.select_all_binding_func, event_data)
                self.PAR.emit_event("<<SheetSelect>>", data=event_data)
                self.PAR.emit_event("<<SelectAll>>", data=event_data)

    def select_columns(self, event: Any) -> None:
        if not self.selected:
            return
        r1, c1, r2, c2 = self.selection_boxes[self.selected.fill_iid].coords
        r, c = self.selected.row, self.selected.column
        self.set_currently_selected(
            r=r,
            c=c,
            item=self.CH.select_col(range(c1, c2), redraw=True),
        )

    def select_rows(self, event: Any) -> None:
        if not self.selected:
            return
        r1, c1, r2, c2 = self.selection_boxes[self.selected.fill_iid].coords
        r, c = self.selected.row, self.selected.column
        self.set_currently_selected(
            r=r,
            c=c,
            item=self.RI.select_row(range(r1, r2), redraw=True),
        )

    def select_row_start(self, event: Any) -> None:
        if self.selected:
            self.see(self.selected.row, 0)
            self.select_cell(self.selected.row, 0, redraw=True)

    def select_a1(self, event: Any) -> None:
        if len(self.row_positions) > 1 and len(self.col_positions) > 1:
            self.see(0, 0)
            self.select_cell(0, 0, redraw=True)

    def select_cell(
        self,
        r: int,
        c: int,
        redraw: bool = False,
        run_binding_func: bool = True,
        ext: bool = False,
    ) -> int:
        boxes_to_hide = tuple(self.selection_boxes)
        fill_iid = self.create_selection_box(r, c, r + 1, c + 1, state="hidden", ext=ext)
        for iid in boxes_to_hide:
            self.hide_selection_box(iid)
        if redraw:
            self.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
        if run_binding_func:
            self.run_selection_binding("cells")
        return fill_iid

    def add_selection(
        self,
        r: int,
        c: int,
        redraw: bool = False,
        run_binding_func: bool = True,
        set_as_current: bool = False,
        ext: bool = False,
    ) -> int:
        fill_iid = self.create_selection_box(r, c, r + 1, c + 1, state="hidden", set_current=set_as_current, ext=ext)
        if redraw:
            self.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
        if run_binding_func:
            self.run_selection_binding("cells")
        return fill_iid

    def toggle_select_cell(
        self,
        row: int,
        column: int,
        add_selection: bool = True,
        redraw: bool = True,
        run_binding_func: bool = True,
        set_as_current: bool = True,
        ext: bool = False,
    ) -> int | None:
        if add_selection:
            if self.cell_selected(row, column, inc_rows=True, inc_cols=True):
                fill_iid = self.deselect(r=row, c=column, redraw=redraw)
            else:
                fill_iid = self.add_selection(
                    r=row,
                    c=column,
                    redraw=redraw,
                    run_binding_func=run_binding_func,
                    set_as_current=set_as_current,
                    ext=ext,
                )
        else:
            if self.cell_selected(row, column, inc_rows=True, inc_cols=True):
                fill_iid = self.deselect(r=row, c=column, redraw=redraw)
            else:
                fill_iid = self.select_cell(row, column, redraw=redraw, ext=ext)
        return fill_iid

    def get_select_event(self, being_drawn_item: None | int = None) -> EventDataDict:
        return event_dict(
            name="select",
            sheet=self.PAR.name,
            selected=self.selected,
            being_selected=self.coords_and_type(being_drawn_item),
            boxes=self.get_boxes(),
        )

    def deselect(
        self,
        r: int | None | str = None,
        c: int | None = None,
        cell: tuple[int, int] | None = None,
        redraw: bool = True,
        run_binding: bool = True,
    ) -> None:
        if not self.selected:
            return
        curr_box = self.selected.fill_iid
        if r == "all" or (r is None and c is None and cell is None):
            self.hide_dropdown_editor_all_canvases()
            for item, _ in self.get_selection_items():
                self.hide_selection_box(item)
        elif r in ("allrows", "allcols"):
            for item, _ in self.get_selection_items(
                columns=r == "allcols",
                rows=r == "allrows",
                cells=False,
            ):
                self.hide_selection_box(item)
                if item == curr_box:
                    self.hide_dropdown_editor_all_canvases()
        elif isinstance(r, int) and c is None and cell is None:
            for item, box in self.get_selection_items(columns=False, cells=False):
                r1, c1, r2, c2 = box.coords
                if r >= r1 and r < r2:
                    resel = self.selected.fill_iid == item
                    to_sel = self.selected.row
                    self.hide_selection_box(item)
                    if item == curr_box:
                        self.hide_dropdown_editor_all_canvases()
                    if r2 - r1 != 1:
                        if r == r1:
                            self.create_selection_box(
                                r1 + 1,
                                0,
                                r2,
                                len(self.col_positions) - 1,
                                "rows",
                                set_current=resel,
                            )
                        elif r == r2 - 1:
                            self.create_selection_box(
                                r1,
                                0,
                                r2 - 1,
                                len(self.col_positions) - 1,
                                "rows",
                                set_current=resel,
                            )
                        else:
                            self.create_selection_box(
                                r1,
                                0,
                                r,
                                len(self.col_positions) - 1,
                                "rows",
                                set_current=resel and to_sel >= r1 and to_sel < r,
                            )
                            self.create_selection_box(
                                r + 1,
                                0,
                                r2,
                                len(self.col_positions) - 1,
                                "rows",
                                set_current=resel and to_sel >= r + 1 and to_sel < r2,
                            )
        elif isinstance(c, int) and r is None and cell is None:
            for item, box in self.get_selection_items(rows=False, cells=False):
                r1, c1, r2, c2 = box.coords
                if c >= c1 and c < c2:
                    resel = self.selected.fill_iid == item
                    to_sel = self.selected.column
                    self.hide_selection_box(item)
                    if item == curr_box:
                        self.hide_dropdown_editor_all_canvases()
                    if c2 - c1 != 1:
                        if c == c1:
                            self.create_selection_box(
                                0,
                                c1 + 1,
                                len(self.row_positions) - 1,
                                c2,
                                "columns",
                                set_current=resel,
                            )
                        elif c == c2 - 1:
                            self.create_selection_box(
                                0,
                                c1,
                                len(self.row_positions) - 1,
                                c2 - 1,
                                "columns",
                                set_current=resel,
                            )
                        else:
                            self.create_selection_box(
                                0,
                                c1,
                                len(self.row_positions) - 1,
                                c,
                                "columns",
                                set_current=resel and to_sel >= c1 and to_sel < c,
                            )
                            self.create_selection_box(
                                0,
                                c + 1,
                                len(self.row_positions) - 1,
                                c2,
                                "columns",
                                set_current=resel and to_sel >= c + 1 and to_sel < c2,
                            )
        elif (isinstance(r, int) and isinstance(c, int) and cell is None) or cell is not None:
            if cell is not None:
                r, c = cell[0], cell[1]
            for item, box in self.get_selection_items(reverse=True):
                r1, c1, r2, c2 = box.coords
                if r >= r1 and c >= c1 and r < r2 and c < c2:
                    self.hide_selection_box(item)
                    if item == curr_box:
                        self.hide_text_editor_and_dropdown(redraw=False)
        if redraw:
            self.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
        sel_event = self.get_select_event(being_drawn_item=self.being_drawn_item)
        if run_binding:
            try_binding(self.deselection_binding_func, sel_event)
        self.PAR.emit_event("<<SheetSelect>>", data=sel_event)

    def deselect_any(
        self,
        rows: Iterator[int] | int | None = None,
        columns: Iterator[int] | int | None = None,
        redraw: bool = True,
    ) -> None:
        if not self.selected:
            return
        rows = int_x_iter(rows)
        columns = int_x_iter(columns)
        curr_box = self.selected.fill_iid
        if is_iterable(rows) and is_iterable(columns):
            rows = tuple(consecutive_ranges(sorted(rows)))
            columns = tuple(consecutive_ranges(sorted(columns)))
            for item, box in self.get_selection_items(reverse=True):
                r1, c1, r2, c2 = box.coords
                hidden = False
                for rows_st, rows_end in rows:
                    if hidden:
                        break
                    for cols_st, cols_end in columns:
                        if ((rows_end >= r1 and rows_end <= r2) or (rows_st >= r1 and rows_st < r2)) and (
                            (cols_end >= c1 and cols_end <= c2) or (cols_st >= c1 and cols_st < c2)
                        ):
                            hidden = self.hide_selection_box(item)
                            if item == curr_box:
                                self.hide_dropdown_editor_all_canvases()
                            break
        elif is_iterable(rows):
            rows = tuple(consecutive_ranges(sorted(rows)))
            for item, box in self.get_selection_items(reverse=True):
                r1, c1, r2, c2 = box.coords
                for rows_st, rows_end in rows:
                    if (rows_end >= r1 and rows_end <= r2) or (rows_st >= r1 and rows_st < r2):
                        self.hide_selection_box(item)
                        if item == curr_box:
                            self.hide_dropdown_editor_all_canvases()
                        break
        elif is_iterable(columns):
            columns = tuple(consecutive_ranges(sorted(columns)))
            for item, box in self.get_selection_items(reverse=True):
                r1, c1, r2, c2 = box.coords
                for cols_st, cols_end in columns:
                    if (cols_end >= c1 and cols_end <= c2) or (cols_st >= c1 and cols_st < c2):
                        self.hide_selection_box(item)
                        if item == curr_box:
                            self.hide_dropdown_editor_all_canvases()
                        break
        else:
            self.deselect()
        if redraw:
            self.refresh()

    def page_UP(self, event: Any = None) -> None:
        r, end = self.visible_text_rows
        if r == end - 1:
            r = max(0, r - 1)
        if self.PAR.ops.page_up_down_select_row:
            if self.RI.row_selection_enabled and (
                self.selected and self.selected.type_ == "rows" or not self.anything_selected()
            ):
                self.RI.select_row(r, redraw=False)
            elif (self.single_selection_enabled or self.toggle_selection_enabled) and self.anything_selected(
                exclude_columns=True, exclude_rows=True
            ):
                self.select_cell(r, self.selected.column, redraw=False)
        if not self.see(r, keep_xscroll=True, bottom_right_corner=True):
            self.main_table_redraw_grid_and_text(redraw_row_index=True)

    def page_DOWN(self, event: Any = None) -> None:
        st, r = self.visible_text_rows
        r -= 1
        if st == r:
            r = min(len(self.row_positions) - 2, r + 1)
        if self.PAR.ops.page_up_down_select_row:
            if self.RI.row_selection_enabled and (
                self.selected and self.selected.type_ == "rows" or not self.anything_selected()
            ):
                self.RI.select_row(r, redraw=False)

            elif (self.single_selection_enabled or self.toggle_selection_enabled) and self.anything_selected(
                exclude_columns=True, exclude_rows=True
            ):
                self.select_cell(r, self.selected.column, redraw=False)
        if not self.see(r, keep_xscroll=True, bottom_right_corner=False):
            self.main_table_redraw_grid_and_text(redraw_row_index=True)

    def arrowkey_UP(self, event: Any = None) -> None:
        if not self.selected:
            return
        r = self.selected.row
        c = self.selected.column
        if not r:
            self.see(r, c)
        else:
            self.see(r - 1, c, redraw=False)
            if self.selected.type_ == "rows" and self.RI.row_selection_enabled:
                self.RI.select_row(r - 1, redraw=True)

            elif self.selected.type_ in ("cells", "columns") and (
                self.single_selection_enabled or self.toggle_selection_enabled
            ):
                self.select_cell(r - 1, c, redraw=True)

    def arrowkey_LEFT(self, event: Any = None) -> Literal["break"]:
        if not self.selected:
            return
        r = self.selected.row
        c = self.selected.column
        if not c:
            self.see(r, c)
        else:
            self.see(r, c - 1, redraw=False)
            if self.selected.type_ == "columns" and self.CH.col_selection_enabled:
                self.CH.select_col(c - 1, redraw=True)

            elif self.selected.type_ in ("cells", "rows") and (
                self.single_selection_enabled or self.toggle_selection_enabled
            ):
                self.select_cell(r, c - 1, redraw=True)
        return "break"

    def arrowkey_DOWN(self, event: Any = None) -> None:
        if not self.selected:
            return
        r = self.selected.row
        c = self.selected.column
        if self.selected.row == len(self.row_positions) - 2:
            self.see(r, c)
        else:
            brc = False if self.PAR.ops.arrow_key_down_right_scroll_page else None
            self.see(
                r + 1,
                c,
                bottom_right_corner=brc,
                redraw=False,
            )
            if self.selected.type_ == "rows" and self.RI.row_selection_enabled:
                self.RI.select_row(r + 1, redraw=True)

            elif self.selected.type_ in ("cells", "columns") and (
                self.single_selection_enabled or self.toggle_selection_enabled
            ):
                self.select_cell(r + 1, c, redraw=True)

    def arrowkey_RIGHT(self, event: Any = None) -> None:
        if not self.selected:
            return
        r = self.selected.row
        c = self.selected.column
        if self.selected.column == len(self.col_positions) - 2:
            self.see(r, c)
        else:
            brc = False if self.PAR.ops.arrow_key_down_right_scroll_page else None
            self.see(
                r,
                c + 1,
                bottom_right_corner=brc,
                redraw=False,
            )
            if self.selected.type_ == "columns" and self.CH.col_selection_enabled:
                self.CH.select_col(c + 1, redraw=True)

            elif self.selected.type_ in ("cells", "rows") and (
                self.single_selection_enabled or self.toggle_selection_enabled
            ):
                self.select_cell(r, c + 1, redraw=True)

    def shift_arrowkey_select_box(
        self,
        r1: int,
        c1: int,
        r2: int,
        c2: int,
        current_r: int,
        current_c: int,
        see_r: int,
        see_c: int,
        keep_xscroll: bool = False,
        keep_yscroll: bool = False,
    ) -> None:
        box_to_hide = self.selected.fill_iid
        self.set_currently_selected(
            current_r,
            current_c,
            self.create_selection_box(
                r1,
                c1,
                r2,
                c2,
                self.selected.type_,
                set_current=False,
                run_binding=True,
            ),
            run_binding=False,
        )
        self.hide_selection_box(box_to_hide)
        self.see(see_r, see_c, keep_xscroll=keep_xscroll, keep_yscroll=keep_yscroll, redraw=False)
        self.refresh()

    def shift_arrowkey_UP(self, event: Any = None) -> None:
        if not self.selected:
            return
        r1, c1, r2, c2 = self.selected.box
        current_r, current_c = self.selected.row, self.selected.column
        if not r1 and self.selected.row == r2 - 1:
            return
        if (r2 - r1 != 1 and self.selected.row != r1 and self.selected.row != r2 - 1) or (
            r2 - r1 != 1 and self.selected.row == r1
        ):
            r2 -= 1
            see_r = r2 - 1
        else:
            r1 -= 1
            see_r = r1
        self.shift_arrowkey_select_box(r1, c1, r2, c2, current_r, current_c, see_r, c1, keep_xscroll=True)

    def shift_arrowkey_DOWN(self, event: Any = None) -> None:
        if not self.selected:
            return
        r1, c1, r2, c2 = self.selected.box
        current_r, current_c = self.selected.row, self.selected.column
        if r2 == len(self.row_positions) - 1 and self.selected.row == r1:
            return
        if r2 - r1 == 1 or self.selected.row == r1:
            r2 += 1
            see_r = r2 - 1
        else:
            r1 += 1
            see_r = r1
        self.shift_arrowkey_select_box(r1, c1, r2, c2, current_r, current_c, see_r, c1, keep_xscroll=True)

    def shift_arrowkey_LEFT(self, event: Any = None) -> None:
        if not self.selected:
            return
        r1, c1, r2, c2 = self.selected.box
        current_r, current_c = self.selected.row, self.selected.column
        if not c1 and self.selected.column == c2 - 1:
            return
        if (c2 - c1 != 1 and self.selected.column != c1 and self.selected.column != c2 - 1) or (
            c2 - c1 != 1 and self.selected.column == c1
        ):
            c2 -= 1
            see_c = c2 - 1
        else:
            c1 -= 1
            see_c = c1
        self.shift_arrowkey_select_box(r1, c1, r2, c2, current_r, current_c, r1, see_c, keep_yscroll=True)

    def shift_arrowkey_RIGHT(self, event: Any = None) -> None:
        if not self.selected:
            return
        r1, c1, r2, c2 = self.selected.box
        current_r, current_c = self.selected.row, self.selected.column
        if c2 == len(self.col_positions) - 1 and self.selected.column == c1:
            return
        if c2 - c1 == 1 or self.selected.column == c1:
            c2 += 1
            see_c = c2 - 1
        else:
            c1 += 1
            see_c = c1
        self.shift_arrowkey_select_box(r1, c1, r2, c2, current_r, current_c, r1, see_c, keep_yscroll=True)

    def menu_add_command(self, menu: tk.Menu, **kwargs) -> None:
        if "label" not in kwargs:
            return
        try:
            index = menu.index(kwargs["label"])
            menu.delete(index)
        except TclError:
            pass
        menu.add_command(**kwargs)

    def table_edit_cell_enabled(self) -> bool:
        return (
            self.rc_popup_menus_enabled
            and self.edit_cell_enabled
            and any(x in self.enabled_bindings_menu_entries for x in ("all", "edit_cell", "edit_bindings", "edit"))
        )

    def index_edit_cell_enabled(self) -> bool:
        return (
            self.rc_popup_menus_enabled
            and self.RI.edit_cell_enabled
            and "edit_index" in self.enabled_bindings_menu_entries
        )

    def header_edit_cell_enabled(self) -> bool:
        return (
            self.rc_popup_menus_enabled
            and self.CH.edit_cell_enabled
            and "edit_header" in self.enabled_bindings_menu_entries
        )

    def create_rc_menus(self) -> None:
        if not self.rc_popup_menu:
            self.rc_popup_menu = tk.Menu(self, tearoff=0, background=self.PAR.ops.popup_menu_bg)
        if not self.CH.ch_rc_popup_menu:
            self.CH.ch_rc_popup_menu = tk.Menu(self.CH, tearoff=0, background=self.PAR.ops.popup_menu_bg)
        if not self.RI.ri_rc_popup_menu:
            self.RI.ri_rc_popup_menu = tk.Menu(self.RI, tearoff=0, background=self.PAR.ops.popup_menu_bg)
        if not self.empty_rc_popup_menu:
            self.empty_rc_popup_menu = tk.Menu(self, tearoff=0, background=self.PAR.ops.popup_menu_bg)

    def enable_bindings(self, bindings: Any, menu: bool = True) -> None:
        if not bindings:
            self._enable_binding("all", menu)
        elif isinstance(bindings, (list, tuple)):
            for binding in bindings:
                if isinstance(binding, (list, tuple)):
                    for bind in binding:
                        self._enable_binding(bind.lower(), menu)
                elif isinstance(binding, str):
                    self._enable_binding(binding.lower(), menu)
        elif isinstance(bindings, str):
            self._enable_binding(bindings.lower(), menu)

    def disable_bindings(self, bindings: Any) -> None:
        if not bindings:
            self._disable_binding("all")
        elif isinstance(bindings, (list, tuple)):
            for binding in bindings:
                if isinstance(binding, (list, tuple)):
                    for bind in binding:
                        self._disable_binding(bind.lower())
                elif isinstance(binding, str):
                    self._disable_binding(binding.lower())
        elif isinstance(bindings, str):
            self._disable_binding(bindings)

    def _enable_binding(self, binding: Binding, menu: bool = True) -> None:
        if binding == "enable_all":
            binding = "all"
        if binding in (
            "all",
            "single",
            "single_selection_mode",
            "single_select",
            "toggle",
            "toggle_selection_mode",
            "toggle_select",
        ):
            self._tksheet_bind("row_start_bindings", self.select_row_start)
            self._tksheet_bind("table_start_bindings", self.select_a1)
            self._tksheet_bind("escape_bindings", self.escape)
            if binding in ("all", "single", "single_selection_mode", "single_select"):
                self.single_selection_enabled = True
                self.toggle_selection_enabled = False
            elif binding in ("toggle", "toggle_selection_mode", "toggle_select"):
                self.toggle_selection_enabled = True
                self.single_selection_enabled = False
        if binding in ("all", "drag_select"):
            self.drag_selection_enabled = True
        if binding in ("all", "column_width_resize"):
            self.CH.width_resizing_enabled = True
        if binding in ("all", "column_select"):
            self.CH.col_selection_enabled = True
            self._tksheet_bind("select_columns_bindings", self.select_columns)
        if binding in ("all", "column_height_resize"):
            self.CH.height_resizing_enabled = True
            self.TL.rh_state()
        if binding in ("all", "column_drag_and_drop", "move_columns"):
            self.CH.drag_and_drop_enabled = True
        if binding in ("all", "double_click_column_resize"):
            self.CH.double_click_resizing_enabled = True
        if binding in ("all", "row_height_resize"):
            self.RI.height_resizing_enabled = True
        if binding in ("all", "double_click_row_resize"):
            self.RI.double_click_resizing_enabled = True
        if binding in ("all", "row_width_resize"):
            self.RI.width_resizing_enabled = True
            self.TL.rw_state()
        if binding in ("all", "row_select"):
            self.RI.row_selection_enabled = True
            self._tksheet_bind("select_rows_bindings", self.select_rows)
        if binding in ("all", "row_drag_and_drop", "move_rows"):
            self.RI.drag_and_drop_enabled = True
        if binding in ("all", "select_all"):
            self.select_all_enabled = True
            self.TL.sa_state()
            self._tksheet_bind("select_all_bindings", self.select_all)
        if binding in ("all", "arrowkeys", "tab"):
            self.tab_enabled = True
            self._tksheet_bind("tab_bindings", self.tab_key)
        if binding in ("all", "arrowkeys", "up"):
            self.up_enabled = True
            self._tksheet_bind("up_bindings", self.arrowkey_UP)
            self._tksheet_bind("shift_up_bindings", self.shift_arrowkey_UP)
        if binding in ("all", "arrowkeys", "right"):
            self.right_enabled = True
            self._tksheet_bind("right_bindings", self.arrowkey_RIGHT)
            self._tksheet_bind("shift_right_bindings", self.shift_arrowkey_RIGHT)
        if binding in ("all", "arrowkeys", "down"):
            self.down_enabled = True
            self._tksheet_bind("down_bindings", self.arrowkey_DOWN)
            self._tksheet_bind("shift_down_bindings", self.shift_arrowkey_DOWN)
        if binding in ("all", "arrowkeys", "left"):
            self.left_enabled = True
            self._tksheet_bind("left_bindings", self.arrowkey_LEFT)
            self._tksheet_bind("shift_left_bindings", self.shift_arrowkey_LEFT)
        if binding in ("all", "arrowkeys", "prior"):
            self.prior_enabled = True
            self._tksheet_bind("prior_bindings", self.page_UP)
        if binding in ("all", "arrowkeys", "next"):
            self.next_enabled = True
            self._tksheet_bind("next_bindings", self.page_DOWN)
        if binding in ("all", "copy", "edit_bindings", "edit"):
            self.copy_enabled = True
            self._tksheet_bind("copy_bindings", self.ctrl_c)
            self._tksheet_bind("copy_plain_bindings", self.ctrl_c_plain)
        if binding in ("all", "cut", "edit_bindings", "edit"):
            self.cut_enabled = True
            self._tksheet_bind("cut_bindings", self.ctrl_x)
        if binding in ("all", "paste", "edit_bindings", "edit"):
            self.paste_enabled = True
            self._tksheet_bind("paste_bindings", self.ctrl_v)
        if binding in ("all", "delete", "edit_bindings", "edit"):
            self.delete_key_enabled = True
            self._tksheet_bind("delete_bindings", self.delete_key)
        if binding in ("all", "undo", "redo", "edit_bindings", "edit"):
            self.undo_enabled = True
            self._tksheet_bind("undo_bindings", self.undo)
            self._tksheet_bind("redo_bindings", self.redo)
        if binding in ("all", "find"):
            self.find_enabled = True
            self._tksheet_bind("find_bindings", self.open_find_window)
            self._tksheet_bind("find_next_bindings", self.find_next)
            self._tksheet_bind("find_previous_bindings", self.find_previous)
        if binding in ("all", "replace"):
            self._tksheet_bind("toggle_replace_bindings", self.replace_toggle)
        if binding in bind_del_columns:
            self.rc_delete_column_enabled = True
            self.rc_popup_menus_enabled = True
            self.rc_select_enabled = True
        if binding in bind_del_rows:
            self.rc_delete_row_enabled = True
            self.rc_popup_menus_enabled = True
            self.rc_select_enabled = True
        if binding in bind_add_columns:
            self.rc_insert_column_enabled = True
            self.rc_popup_menus_enabled = True
            self.rc_select_enabled = True
        if binding in bind_add_rows:
            self.rc_insert_row_enabled = True
            self.rc_popup_menus_enabled = True
            self.rc_select_enabled = True
        if binding in ("all", "sort_cells"):
            self.rc_sort_cells_enabled = True
        if binding in ("all", "sort_row"):
            self.rc_sort_row_enabled = True
        if binding in ("all", "sort_column", "sort_col"):
            self.rc_sort_column_enabled = True
        if binding in ("all", "sort_columns", "sort_cols"):
            self.rc_sort_columns_enabled = True
        if binding in ("all", "sort_rows"):
            self.rc_sort_rows_enabled = True
        if binding in ("all", "right_click_popup_menu", "rc_popup_menu", "rc_menu"):
            self.rc_popup_menus_enabled = True
            self.rc_select_enabled = True
        if binding in ("all", "right_click_select", "rc_select"):
            self.rc_select_enabled = True
        if binding in ("all", "edit_cell", "edit_bindings", "edit"):
            self.edit_cell_enabled = True
            for w in (self, self.RI, self.CH):
                w.bind("<Key>", self.open_cell)
        if binding in ("edit_header"):
            self.CH.edit_cell_enabled = True
        if binding in ("edit_index"):
            self.RI.edit_cell_enabled = True
        # has to be specifically enabled
        if binding in ("ctrl_click_select", "ctrl_select"):
            self.ctrl_select_enabled = True
        self.enabled_bindings.add(binding)
        if menu:
            self.enabled_bindings_menu_entries.add(binding)

    def _tksheet_bind(self, bindings_key: str, func: Callable) -> None:
        for widget in (self, self.RI, self.CH, self.TL):
            for binding in self.PAR.ops[bindings_key]:
                widget.bind(binding, func)

    def _disable_binding(self, binding: Binding) -> None:
        if binding == "disable_all":
            binding = "all"
        if binding == "all":
            self.enabled_bindings = set()
            self.enabled_bindings_menu_entries = set()
        else:
            self.enabled_bindings.discard(binding)
            self.enabled_bindings_menu_entries.discard(binding)
        if binding in (
            "all",
            "single",
            "single_selection_mode",
            "single_select",
            "toggle",
            "toggle_selection_mode",
            "toggle_select",
        ):
            self._tksheet_unbind("row_start_bindings")
            self._tksheet_unbind("table_start_bindings")
            self._tksheet_unbind("escape_bindings")
            self.single_selection_enabled = False
            self.toggle_selection_enabled = False
        if binding in ("all", "drag_select"):
            self.drag_selection_enabled = False
        if binding in ("all", "column_width_resize"):
            self.CH.width_resizing_enabled = False
        if binding in ("all", "column_select"):
            self.CH.col_selection_enabled = False
            self._tksheet_unbind("select_columns_bindings")
        if binding in ("all", "column_height_resize"):
            self.CH.height_resizing_enabled = False
            self.TL.rh_state("hidden")
        if binding in ("all", "column_drag_and_drop", "move_columns"):
            self.CH.drag_and_drop_enabled = False
        if binding in ("all", "double_click_column_resize"):
            self.CH.double_click_resizing_enabled = False
        if binding in ("all", "row_height_resize"):
            self.RI.height_resizing_enabled = False
        if binding in ("all", "double_click_row_resize"):
            self.RI.double_click_resizing_enabled = False
        if binding in ("all", "row_width_resize"):
            self.RI.width_resizing_enabled = False
            self.TL.rw_state("hidden")
        if binding in ("all", "row_select"):
            self.RI.row_selection_enabled = False
            self._tksheet_unbind("select_rows_bindings")
        if binding in ("all", "row_drag_and_drop", "move_rows"):
            self.RI.drag_and_drop_enabled = False
        if binding in bind_del_columns:
            self.rc_delete_column_enabled = False
        if binding in bind_del_rows:
            self.rc_delete_row_enabled = False
        if binding in bind_add_columns:
            self.rc_insert_column_enabled = False
        if binding in bind_add_rows:
            self.rc_insert_row_enabled = False
        if binding in ("all", "sort_cells"):
            self.rc_sort_cells_enabled = False
        if binding in ("all", "sort_row"):
            self.rc_sort_row_enabled = False
        if binding in ("all", "sort_column", "sort_col"):
            self.rc_sort_column_enabled = False
        if binding in ("all", "sort_columns", "sort_cols"):
            self.rc_sort_columns_enabled = False
        if binding in ("all", "sort_rows"):
            self.rc_sort_rows_enabled = False
        if binding in ("all", "right_click_popup_menu", "rc_popup_menu", "rc_menu"):
            self.rc_popup_menus_enabled = False
        if binding in ("all", "right_click_select", "rc_select"):
            self.rc_select_enabled = False
        if binding in ("all", "edit_cell", "edit_bindings", "edit"):
            self.edit_cell_enabled = False
            for w in (self, self.RI, self.CH):
                w.unbind("<Key>")
        if binding in ("all", "edit_header", "edit_bindings", "edit"):
            self.CH.edit_cell_enabled = False
        if binding in ("all", "edit_index", "edit_bindings", "edit"):
            self.RI.edit_cell_enabled = False
        if binding in ("all", "ctrl_click_select", "ctrl_select"):
            self.ctrl_select_enabled = False
        if binding in ("all", "select_all"):
            self.select_all_enabled = False
            self.TL.sa_state("hidden")
            self._tksheet_unbind("select_all_bindings")
        if binding in ("all", "copy", "edit_bindings", "edit"):
            self.copy_enabled = False
            self._tksheet_unbind("copy_bindings")
            self._tksheet_unbind("copy_plain_bindings")
        if binding in ("all", "cut", "edit_bindings", "edit"):
            self.cut_enabled = False
            self._tksheet_unbind("cut_bindings")
        if binding in ("all", "paste", "edit_bindings", "edit"):
            self.paste_enabled = False
            self._tksheet_unbind("paste_bindings")
        if binding in ("all", "delete", "edit_bindings", "edit"):
            self.delete_key_enabled = False
            self._tksheet_unbind("delete_bindings")
        if binding in ("all", "arrowkeys", "tab"):
            self.tab_enabled = False
            self._tksheet_unbind("tab_bindings")
        if binding in ("all", "arrowkeys", "up"):
            self.up_enabled = False
            self._tksheet_unbind("up_bindings")
        if binding in ("all", "arrowkeys", "right"):
            self.right_enabled = False
            self._tksheet_unbind("right_bindings")
        if binding in ("all", "arrowkeys", "down"):
            self.down_enabled = False
            self._tksheet_unbind("down_bindings")
        if binding in ("all", "arrowkeys", "left"):
            self.left_enabled = False
            self._tksheet_unbind("left_bindings")
        if binding in ("all", "arrowkeys", "prior"):
            self.prior_enabled = False
            self._tksheet_unbind("prior_bindings")
        if binding in ("all", "arrowkeys", "next"):
            self.next_enabled = False
            self._tksheet_unbind("next_bindings")
        if binding in ("all", "undo", "redo", "edit_bindings", "edit"):
            self.undo_enabled = False
            self._tksheet_unbind("undo_bindings", "redo_bindings")
        if binding in ("all", "find"):
            self.find_enabled = False
            self._tksheet_unbind("find_bindings")
            self._tksheet_unbind("find_next_bindings")
            self._tksheet_unbind("find_previous_bindings")
            self.close_find_window()
        if binding in ("all", "replace"):
            self._tksheet_unbind("toggle_replace_bindings")

    def _tksheet_unbind(self, *keys) -> None:
        for widget in (self, self.RI, self.CH, self.TL):
            for bindings_key in keys:
                for binding in self.PAR.ops[bindings_key]:
                    widget.unbind(binding)

    def reset_resize_vars(self) -> None:
        self.RI.rsz_w = None
        self.RI.rsz_h = None
        self.CH.rsz_w = None
        self.CH.rsz_h = None

    def enter(self, event: Any) -> None:
        self.config(cursor="")
        self.current_cursor = ""
        self.reset_resize_vars()
        try_binding(self.extra_enter_func, event)

    def not_currently_resizing(self) -> bool:
        return all(v is None for v in (self.RI.rsz_h, self.RI.rsz_w, self.CH.rsz_h, self.CH.rsz_w))

    def is_readonly(self, datarn: int, datacn: int) -> bool:
        return (
            ((datarn, datacn) in self.cell_options and "readonly" in self.cell_options[(datarn, datacn)])
            or (datarn in self.row_options and "readonly" in self.row_options[datarn])
            or (datacn in self.col_options and "readonly" in self.col_options[datacn])
        )

    def single_cell_selected(self) -> bool:
        return len(self.selection_boxes) == 1 and box_is_single_cell(*next(iter(self.selection_boxes.values())).coords)

    def estimate_selections_readonly(self) -> bool:
        max_cells = estimate_max_visible_cells(self)
        total_cells = 0
        for box in self.selection_boxes.values():
            r1, c1, r2, c2 = box.coords
            total_cells += (r2 - r1) * (c2 - c1)
        if total_cells > max_cells:
            return False
        for box in self.selection_boxes.values():
            r1, c1, r2, c2 = box.coords
            for r in range(r1, r2):
                for c in range(c1, c2):
                    datarn, datacn = self.datarn(r), self.datacn(c)
                    if not (
                        ((datarn, datacn) in self.cell_options and "readonly" in self.cell_options[(datarn, datacn)])
                        or (datarn in self.row_options and "readonly" in self.row_options[datarn])
                        or (datacn in self.col_options and "readonly" in self.col_options[datacn])
                    ):
                        return False
        return True

    def rc(self, event: Any = None) -> None:
        self.mouseclick_outside_editor_or_dropdown_all_canvases()
        self.focus_set()
        popup_menu = None
        if (self.single_selection_enabled or self.toggle_selection_enabled) and self.not_currently_resizing():
            r = self.identify_row(y=event.y)
            c = self.identify_col(x=event.x)
            if r < len(self.row_positions) - 1 and c < len(self.col_positions) - 1:
                if self.col_selected(c):
                    if self.rc_popup_menus_enabled:
                        popup_menu = self.CH.ch_rc_popup_menu
                        build_header_rc_menu(self, popup_menu, self.selected)
                elif self.row_selected(r):
                    if self.rc_popup_menus_enabled:
                        popup_menu = self.RI.ri_rc_popup_menu
                        build_index_rc_menu(self, popup_menu, self.selected)
                elif self.cell_selected(r, c):
                    if self.rc_popup_menus_enabled:
                        popup_menu = self.rc_popup_menu
                        build_table_rc_menu(self, popup_menu, self.selected)
                else:
                    if self.rc_select_enabled:
                        if self.single_selection_enabled:
                            self.select_cell(r, c, redraw=True)
                        elif self.toggle_selection_enabled:
                            self.toggle_select_cell(r, c, redraw=True)
                    if self.rc_popup_menus_enabled:
                        popup_menu = self.rc_popup_menu
                        build_table_rc_menu(self, popup_menu, self.selected)
            else:
                self.deselect("all")
                if self.rc_popup_menus_enabled:
                    popup_menu = self.empty_rc_popup_menu
                    build_empty_rc_menu(self, popup_menu)
        try_binding(self.extra_rc_func, event)
        if popup_menu:
            popup_menu.tk_popup(event.x_root, event.y_root)

    def b1_press(self, event: Any = None) -> None:
        self.closed_dropdown = self.mouseclick_outside_editor_or_dropdown_all_canvases()
        self.focus_set()
        if (
            self.identify_col(x=event.x, allow_end=False) is None
            or self.identify_row(y=event.y, allow_end=False) is None
        ):
            self.deselect("all")
        r = self.identify_row(y=event.y)
        c = self.identify_col(x=event.x)
        if self.single_selection_enabled and self.not_currently_resizing():
            if r < len(self.row_positions) - 1 and c < len(self.col_positions) - 1:
                self.being_drawn_item = True
                self.being_drawn_item = self.select_cell(r, c, redraw=True)
        elif self.toggle_selection_enabled and self.not_currently_resizing():
            r = self.identify_row(y=event.y)
            c = self.identify_col(x=event.x)
            if r < len(self.row_positions) - 1 and c < len(self.col_positions) - 1:
                self.toggle_select_cell(r, c, redraw=True)
        self.b1_pressed_loc = (r, c)
        try_binding(self.extra_b1_press_func, event)

    def create_resize_line(
        self, x1: int, y1: int, x2: int, y2: int, width: int, fill: str, tag: str | tuple[str]
    ) -> None:
        if self.hidd_resize_lines:
            t, sh = self.hidd_resize_lines.popitem()
            self.coords(t, x1, y1, x2, y2)
            if sh:
                self.itemconfig(t, width=width, fill=fill, tags=tag)
            else:
                self.itemconfig(t, width=width, fill=fill, tags=tag, state="normal")
            self.lift(t)
        else:
            t = self.create_line(x1, y1, x2, y2, width=width, fill=fill, tags=tag)
        self.disp_resize_lines[t] = True

    def delete_resize_lines(self) -> None:
        self.hidd_resize_lines.update(self.disp_resize_lines)
        self.disp_resize_lines = {}
        for t, sh in self.hidd_resize_lines.items():
            if sh:
                self.itemconfig(t, state="hidden")
                self.hidd_resize_lines[t] = False

    def ctrl_b1_press(self, event: Any = None) -> None:
        self.mouseclick_outside_editor_or_dropdown_all_canvases()
        self.focus_set()
        self.ctrl_b1_pressed = True
        if self.ctrl_select_enabled and self.not_currently_resizing():
            self.b1_pressed_loc = None
            rowsel = int(self.identify_row(y=event.y))
            colsel = int(self.identify_col(x=event.x))
            if rowsel < len(self.row_positions) - 1 and colsel < len(self.col_positions) - 1:
                self.being_drawn_item = True
                self.being_drawn_item = self.add_selection(rowsel, colsel, set_as_current=True, run_binding_func=False)
                sel_event = self.get_select_event(being_drawn_item=self.being_drawn_item)
                try_binding(self.ctrl_selection_binding_func, sel_event)
                self.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True, redraw_table=True)
                self.PAR.emit_event("<<SheetSelect>>", data=sel_event)
        elif not self.ctrl_select_enabled:
            self.b1_press(event)

    def ctrl_shift_b1_press(self, event: Any = None) -> None:
        self.mouseclick_outside_editor_or_dropdown_all_canvases()
        self.focus_set()
        self.ctrl_b1_pressed = True
        if self.ctrl_select_enabled and self.drag_selection_enabled and self.not_currently_resizing():
            self.b1_pressed_loc = None
            rowsel = int(self.identify_row(y=event.y))
            colsel = int(self.identify_col(x=event.x))
            if rowsel < len(self.row_positions) - 1 and colsel < len(self.col_positions) - 1:
                if self.selected and self.selected.type_ == "cells":
                    self.being_drawn_item = self.recreate_selection_box(
                        *self.get_shift_select_box(self.selected.row, rowsel, self.selected.column, colsel),
                        fill_iid=self.selected.fill_iid,
                    )
                else:
                    self.being_drawn_item = self.add_selection(
                        rowsel,
                        colsel,
                        set_as_current=True,
                        run_binding_func=False,
                    )
                self.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True, redraw_table=True)
                sel_event = self.get_select_event(being_drawn_item=self.being_drawn_item)
                try_binding(self.shift_selection_binding_func, sel_event)
                self.PAR.emit_event("<<SheetSelect>>", data=sel_event)
        elif not self.ctrl_select_enabled:
            self.shift_b1_press(event)

    def shift_b1_press(self, event: Any = None) -> None:
        self.mouseclick_outside_editor_or_dropdown_all_canvases()
        self.focus_set()
        if self.drag_selection_enabled and self.not_currently_resizing():
            self.b1_pressed_loc = None
            rowsel = int(self.identify_row(y=event.y))
            colsel = int(self.identify_col(x=event.x))
            if rowsel < len(self.row_positions) - 1 and colsel < len(self.col_positions) - 1:
                if self.selected and self.selected.type_ == "cells":
                    r_to_sel, c_to_sel = self.selected.row, self.selected.column
                    self.deselect("all", redraw=False)
                    self.being_drawn_item = self.create_selection_box(
                        *self.get_shift_select_box(r_to_sel, rowsel, c_to_sel, colsel),
                    )
                    self.set_currently_selected(r_to_sel, c_to_sel, self.being_drawn_item)
                else:
                    self.being_drawn_item = self.select_cell(
                        rowsel,
                        colsel,
                        redraw=False,
                        run_binding_func=False,
                    )
                self.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True, redraw_table=True)
                sel_event = self.get_select_event(being_drawn_item=self.being_drawn_item)
                try_binding(self.shift_selection_binding_func, sel_event)
                self.PAR.emit_event("<<SheetSelect>>", data=sel_event)

    def get_shift_select_box(self, min_r: int, rowsel: int, min_c: int, colsel: int) -> tuple[int, int, int, int]:
        if rowsel >= min_r and colsel >= min_c:
            return min_r, min_c, rowsel + 1, colsel + 1
        elif rowsel >= min_r and min_c >= colsel:
            return min_r, colsel, rowsel + 1, min_c + 1
        elif min_r >= rowsel and colsel >= min_c:
            return rowsel, min_c, min_r + 1, colsel + 1
        elif min_r >= rowsel and min_c >= colsel:
            return rowsel, colsel, min_r + 1, min_c + 1

    def get_b1_motion_box(
        self,
        start_row: int,
        start_col: int,
        end_row: int,
        end_col: int,
    ) -> tuple[int, int, int, int, Literal["cells"]]:
        if end_row >= start_row and end_col >= start_col and (end_row - start_row or end_col - start_col):
            return start_row, start_col, end_row + 1, end_col + 1, "cells"
        elif end_row >= start_row and end_col < start_col and (end_row - start_row or start_col - end_col):
            return start_row, end_col, end_row + 1, start_col + 1, "cells"
        elif end_row < start_row and end_col >= start_col and (start_row - end_row or end_col - start_col):
            return end_row, start_col, start_row + 1, end_col + 1, "cells"
        elif end_row < start_row and end_col < start_col and (start_row - end_row or start_col - end_col):
            return end_row, end_col, start_row + 1, start_col + 1, "cells"
        else:
            return start_row, start_col, start_row + 1, start_col + 1, "cells"

    def b1_motion(self, event: Any) -> None:
        if self.drag_selection_enabled and all(
            v is None
            for v in (
                self.RI.rsz_h,
                self.RI.rsz_w,
                self.CH.rsz_h,
                self.CH.rsz_w,
            )
        ):
            need_redraw = False
            end_row = self.identify_row(y=event.y)
            end_col = self.identify_col(x=event.x)
            if (
                end_row < len(self.row_positions) - 1
                and end_col < len(self.col_positions) - 1
                and self.selected
                and self.selected.type_ == "cells"
            ):
                box = self.get_b1_motion_box(
                    self.selected.row,
                    self.selected.column,
                    end_row,
                    end_col,
                )
                if (
                    box is not None
                    and self.being_drawn_item is not None
                    and self.coords_and_type(self.being_drawn_item) != box
                ):
                    if box[2] - box[0] != 1 or box[3] - box[1] != 1:
                        self.being_drawn_item = self.recreate_selection_box(*box[:-1], fill_iid=self.selected.fill_iid)
                    else:
                        self.being_drawn_item = self.select_cell(
                            box[0],
                            box[1],
                            run_binding_func=False,
                        )
                    need_redraw = True
                    sel_event = self.get_select_event(being_drawn_item=self.being_drawn_item)
                    try_binding(self.drag_selection_binding_func, sel_event)
                    self.PAR.emit_event("<<SheetSelect>>", data=sel_event)
            if self.scroll_if_event_offscreen(event):
                need_redraw = True
            if need_redraw:
                self.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True, redraw_table=True)
        try_binding(self.extra_b1_motion_func, event)

    def ctrl_b1_motion(self, event: Any) -> None:
        if self.ctrl_select_enabled and self.drag_selection_enabled and self.not_currently_resizing():
            need_redraw = False
            end_row = self.identify_row(y=event.y)
            end_col = self.identify_col(x=event.x)
            if (
                end_row < len(self.row_positions) - 1
                and end_col < len(self.col_positions) - 1
                and self.selected
                and self.selected.type_ == "cells"
            ):
                box = self.get_b1_motion_box(
                    self.selected.row,
                    self.selected.column,
                    end_row,
                    end_col,
                )
                if (
                    box is not None
                    and self.being_drawn_item is not None
                    and self.coords_and_type(self.being_drawn_item) != box
                ):
                    if box[2] - box[0] != 1 or box[3] - box[1] != 1:
                        self.being_drawn_item = self.recreate_selection_box(*box[:-1], self.selected.fill_iid)
                    else:
                        self.hide_selection_box(self.selected.fill_iid)
                        self.being_drawn_item = self.add_selection(
                            box[0],
                            box[1],
                            run_binding_func=False,
                            set_as_current=True,
                        )
                    need_redraw = True
                    sel_event = self.get_select_event(being_drawn_item=self.being_drawn_item)
                    try_binding(self.drag_selection_binding_func, sel_event)
                    self.PAR.emit_event("<<SheetSelect>>", data=sel_event)
            if self.scroll_if_event_offscreen(event):
                need_redraw = True
            if need_redraw:
                self.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True, redraw_table=True)
        elif not self.ctrl_select_enabled:
            self.b1_motion(event)

    def b1_release(self, event: Any = None) -> None:
        to_hide = self.being_drawn_item
        if self.being_drawn_item is not None and (to_sel := self.coords_and_type(self.being_drawn_item)):
            r_to_sel, c_to_sel = self.selected.row, self.selected.column
            self.being_drawn_item = None
            self.set_currently_selected(
                r_to_sel,
                c_to_sel,
                item=self.create_selection_box(
                    *to_sel,
                    state=(
                        "hidden"
                        if (to_sel.upto_r - to_sel.from_r == 1 and to_sel.upto_c - to_sel.from_c == 1)
                        else "normal"
                    ),
                    set_current=False,
                ),
                run_binding=False,
            )
        self.hide_selection_box(to_hide)
        self.RI.rsz_w = None
        self.CH.rsz_h = None
        if self.b1_pressed_loc is not None:
            r = self.identify_row(y=event.y, allow_end=False)
            c = self.identify_col(x=event.x, allow_end=False)
            if r is not None and c is not None and (r, c) == self.b1_pressed_loc:
                datarn = self.datarn(r)
                datacn = self.datacn(c)
                if self.get_cell_kwargs(datarn, datacn, key="dropdown") or self.get_cell_kwargs(
                    datarn, datacn, key="checkbox"
                ):
                    canvasx = self.canvasx(event.x)
                    if (
                        self.closed_dropdown != self.b1_pressed_loc
                        and self.get_cell_kwargs(datarn, datacn, key="dropdown")
                        and canvasx > self.col_positions[c + 1] - self.table_txt_height - 4
                        and canvasx < self.col_positions[c + 1] - 1
                    ) or (
                        self.get_cell_kwargs(datarn, datacn, key="checkbox")
                        and canvasx < self.col_positions[c] + self.table_txt_height + 4
                        and self.canvasy(event.y) < self.row_positions[r] + self.table_txt_height + 4
                    ):
                        self.open_cell(event)
            else:
                self.mouseclick_outside_editor_or_dropdown_all_canvases()
        self.ctrl_b1_pressed = False
        self.b1_pressed_loc = None
        self.closed_dropdown = None
        try_binding(self.extra_b1_release_func, event)

    def double_b1(self, event: Any = None) -> None:
        self.mouseclick_outside_editor_or_dropdown_all_canvases()
        self.focus_set()
        if (
            self.identify_col(x=event.x, allow_end=False) is None
            or self.identify_row(y=event.y, allow_end=False) is None
        ):
            self.deselect("all")
        elif self.single_selection_enabled and self.not_currently_resizing():
            r = self.identify_row(y=event.y)
            c = self.identify_col(x=event.x)
            if r < len(self.row_positions) - 1 and c < len(self.col_positions) - 1:
                self.select_cell(r, c, redraw=True)
                if self.edit_cell_enabled:
                    self.open_cell(event)
        elif self.toggle_selection_enabled and self.not_currently_resizing():
            r = self.identify_row(y=event.y)
            c = self.identify_col(x=event.x)
            if r < len(self.row_positions) - 1 and c < len(self.col_positions) - 1:
                self.toggle_select_cell(r, c, redraw=True)
                if self.edit_cell_enabled:
                    self.open_cell(event)
        try_binding(self.extra_double_b1_func, event)

    def identify_row(self, event: Any = None, y: int | None = None, allow_end: bool = True) -> int | None:
        if event is None:
            y2 = self.canvasy(y)
        elif y is None:
            y2 = self.canvasy(event.y)
        r = bisect_left(self.row_positions, y2)
        if r != 0:
            r -= 1
        if not allow_end and r >= len(self.row_positions) - 1:
            return None
        return r

    def identify_col(self, event: Any = None, x: int | None = None, allow_end: bool = True) -> int | None:
        if event is None:
            x2 = self.canvasx(x)
        elif x is None:
            x2 = self.canvasx(event.x)
        c = bisect_left(self.col_positions, x2)
        if c != 0:
            c -= 1
        if not allow_end and c >= len(self.col_positions) - 1:
            return None
        return c

    def fix_views(self) -> None:
        xcheck = self.xview()
        ycheck = self.yview()
        if xcheck and xcheck[0] <= 0:
            self.xview("moveto", 0)
            if self.show_header:
                self.CH.xview("moveto", 0)
        elif len(xcheck) > 1 and xcheck[1] >= 1:
            self.xview("moveto", 1)
            if self.show_header:
                self.CH.xview("moveto", 1)
        if ycheck and ycheck[0] <= 0:
            self.yview("moveto", 0)
            if self.show_index:
                self.RI.yview("moveto", 0)
        elif len(ycheck) > 1 and ycheck[1] >= 1:
            self.yview("moveto", 1)
            if self.show_index:
                self.RI.yview("moveto", 1)

    def scroll_if_event_offscreen(self, event: Any) -> bool:
        need_redraw = False
        if self.data:
            xcheck = self.xview()
            ycheck = self.yview()
            if len(xcheck) > 1 and xcheck[0] > 0 and event.x < 0:
                try:
                    self.xview_scroll(-1, "units")
                    self.CH.xview_scroll(-1, "units")
                except Exception:
                    pass
                need_redraw = True
            if len(ycheck) > 1 and ycheck[0] > 0 and event.y < 0:
                try:
                    self.yview_scroll(-1, "units")
                    self.RI.yview_scroll(-1, "units")
                except Exception:
                    pass
                need_redraw = True
            if len(xcheck) > 1 and xcheck[1] < 1 and event.x > self.winfo_width():
                try:
                    self.xview_scroll(1, "units")
                    self.CH.xview_scroll(1, "units")
                except Exception:
                    pass
                need_redraw = True
            if len(ycheck) > 1 and ycheck[1] < 1 and event.y > self.winfo_height():
                try:
                    self.yview_scroll(1, "units")
                    self.RI.yview_scroll(1, "units")
                except Exception:
                    pass
                need_redraw = True
        if need_redraw:
            self.fix_views()
            self.x_move_synced_scrolls("moveto", self.xview()[0])
            self.y_move_synced_scrolls("moveto", self.yview()[0])
        return need_redraw

    def x_move_synced_scrolls(self, *args: Any, redraw: bool = True, use_scrollbar: bool = False) -> None:
        for widget in self.synced_scrolls:
            # try:
            if hasattr(widget, "MT"):
                if use_scrollbar:
                    widget.MT._xscrollbar(*args, move_synced=False)
                else:
                    widget.MT.set_xviews(*args, move_synced=False, redraw=redraw)
            else:
                widget.xview(*args)
            # except Exception:
            #     continue

    def y_move_synced_scrolls(self, *args: Any, redraw: bool = True, use_scrollbar: bool = False) -> None:
        for widget in self.synced_scrolls:
            # try:
            if hasattr(widget, "MT"):
                if use_scrollbar:
                    widget.MT._yscrollbar(*args, move_synced=False)
                else:
                    widget.MT.set_yviews(*args, move_synced=False, redraw=redraw)
            else:
                widget.yview(*args)
            # except Exception:
            #     continue

    def _xscrollbar(self, *args: Any, move_synced: bool = True) -> None:
        self.xview(*args)
        if self.show_header:
            self.CH.xview(*args)
        self.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=False)
        if move_synced:
            self.x_move_synced_scrolls(*args, use_scrollbar=True)

    def _yscrollbar(self, *args: Any, move_synced: bool = True) -> None:
        self.yview(*args)
        if self.show_index:
            self.RI.yview(*args)
        self.main_table_redraw_grid_and_text(redraw_header=False, redraw_row_index=True)
        if move_synced:
            self.y_move_synced_scrolls(*args, use_scrollbar=True)

    def set_xviews(
        self,
        *args: Any,
        move_synced: bool = True,
        redraw: bool = True,
    ) -> None:
        self.main_table_redraw_grid_and_text(setting_views=True)
        if not self.PAR._startup_complete:
            self.update_idletasks()
        self.xview(*args)
        if self.show_header:
            if not self.PAR._startup_complete:
                self.CH.update_idletasks()
            self.CH.xview(*args)
        if redraw:
            self.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=False)
        if move_synced:
            self.x_move_synced_scrolls(*args)
        self.fix_views()

    def set_yviews(
        self,
        *args: Any,
        move_synced: bool = True,
        redraw: bool = True,
    ) -> None:
        self.main_table_redraw_grid_and_text(setting_views=True)
        if not self.PAR._startup_complete:
            self.update_idletasks()
        self.yview(*args)
        if self.show_index:
            if not self.PAR._startup_complete:
                self.RI.update_idletasks()
            self.RI.yview(*args)
        if redraw:
            self.main_table_redraw_grid_and_text(redraw_header=False, redraw_row_index=True)
        if move_synced:
            self.y_move_synced_scrolls(*args)
        self.fix_views()

    def set_view(self, x_args: list[str, float], y_args: list[str, float]) -> None:
        self.set_xviews(*x_args)
        self.set_yviews(*y_args)

    def mousewheel(self, event: Any) -> None:
        if event.delta < 0 or event.num == 5:
            self.yview_scroll(1, "units")
            self.RI.yview_scroll(1, "units")
            self.y_move_synced_scrolls("moveto", self.yview()[0])
        elif event.delta >= 0 or event.num == 4:
            if self.canvasy(0) <= 0:
                return
            self.yview_scroll(-1, "units")
            self.RI.yview_scroll(-1, "units")
            self.y_move_synced_scrolls("moveto", self.yview()[0])
        self.main_table_redraw_grid_and_text(redraw_header=False, redraw_row_index=True)

    def shift_mousewheel(self, event: Any) -> None:
        if event.delta < 0 or event.num == 5:
            self.xview_scroll(1, "units")
            self.CH.xview_scroll(1, "units")
            self.x_move_synced_scrolls("moveto", self.xview()[0])
        elif event.delta >= 0 or event.num == 4:
            if self.canvasx(0) <= 0:
                return
            self.xview_scroll(-1, "units")
            self.CH.xview_scroll(-1, "units")
            self.x_move_synced_scrolls("moveto", self.xview()[0])
        self.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=False)

    def ctrl_mousewheel(self, event: Any) -> None:
        if event.delta < 0 or event.num == 5:
            self.zoom_out()
        elif event.delta >= 0 or event.num == 4:
            self.zoom_in()

    def zoom_in(self, event: Any = None) -> None:
        self.zoom_font(
            (self.PAR.ops.table_font[0], self.PAR.ops.table_font[1] + 1, self.PAR.ops.table_font[2]),
            (self.PAR.ops.index_font[0], self.PAR.ops.index_font[1] + 1, self.PAR.ops.index_font[2]),
            (self.PAR.ops.header_font[0], self.PAR.ops.header_font[1] + 1, self.PAR.ops.header_font[2]),
            (self.PAR.ops.popup_menu_font[0], self.PAR.ops.popup_menu_font[1] + 1, self.PAR.ops.popup_menu_font[2]),
            "in",
        )

    def zoom_out(self, event: Any = None) -> None:
        if (
            self.PAR.ops.table_font[1] < 2
            or self.PAR.ops.index_font[1] < 2
            or self.PAR.ops.header_font[1] < 2
            or self.PAR.ops.popup_menu_font[1] < 2
        ):
            return
        self.zoom_font(
            (self.PAR.ops.table_font[0], self.PAR.ops.table_font[1] - 1, self.PAR.ops.table_font[2]),
            (self.PAR.ops.index_font[0], self.PAR.ops.index_font[1] - 1, self.PAR.ops.index_font[2]),
            (self.PAR.ops.header_font[0], self.PAR.ops.header_font[1] - 1, self.PAR.ops.header_font[2]),
            (self.PAR.ops.popup_menu_font[0], self.PAR.ops.popup_menu_font[1] - 1, self.PAR.ops.popup_menu_font[2]),
            "out",
        )

    def zoom_font(
        self,
        table_font: FontTuple,
        index_font: FontTuple,
        header_font: FontTuple,
        popup_font: FontTuple,
        zoom: Literal["in", "out"],
    ) -> None:
        self.saved_column_widths = {}
        self.saved_row_heights = {}
        # should record position prior to change and then see after change
        y = self.canvasy(0)
        x = self.canvasx(0)
        r = self.identify_row(y=0)
        c = self.identify_col(x=0)
        try:
            r_pc = (y - self.row_positions[r]) / (self.row_positions[r + 1] - self.row_positions[r])
        except Exception:
            r_pc = 0.0
        try:
            c_pc = (x - self.col_positions[c]) / (self.col_positions[c + 1] - self.col_positions[c])
        except Exception:
            c_pc = 0.0
        old_min_row_height = int(self.min_row_height)
        old_default_row_height = int(self.get_default_row_height())
        self.set_table_font(table_font, row_heights=False)
        self.set_index_font(index_font, row_heights=False)
        self.set_header_font(header_font)
        self.PAR.ops.popup_menu_font = FontTuple(*popup_font)
        if self.PAR.ops.set_cell_sizes_on_zoom:
            self.set_all_cell_sizes_to_text()
            self.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
        elif not self.PAR.ops.set_cell_sizes_on_zoom:
            self.change_font_manage_row_heights(old_min_row_height, old_default_row_height)
            self.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
            self.recreate_all_selection_boxes()
        self.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
        self.refresh_open_window_positions(zoom=zoom)
        self.RI.refresh_open_window_positions(zoom=zoom)
        self.CH.refresh_open_window_positions(zoom=zoom)
        self.see(
            r,
            c,
            check_cell_visibility=False,
            redraw=True,
            r_pc=r_pc,
            c_pc=c_pc,
        )

    def get_txt_w(self, txt: str, font: None | FontTuple = None) -> int:
        self.txt_measure_canvas.itemconfig(
            self.txt_measure_canvas_text,
            text=txt,
            font=self.PAR.ops.table_font if font is None else font,
        )
        b = self.txt_measure_canvas.bbox(self.txt_measure_canvas_text)
        return b[2] - b[0]

    def get_txt_h(self, txt: str, font: None | FontTuple = None) -> int:
        self.txt_measure_canvas.itemconfig(
            self.txt_measure_canvas_text,
            text=txt,
            font=self.PAR.ops.table_font if font is None else font,
        )
        b = self.txt_measure_canvas.bbox(self.txt_measure_canvas_text)
        return b[3] - b[1]

    def get_txt_dimensions(self, txt: str, font: None | FontTuple = None) -> tuple[int, int]:
        self.txt_measure_canvas.itemconfig(
            self.txt_measure_canvas_text,
            text=txt,
            font=self.PAR.ops.table_font if font is None else font,
        )
        b = self.txt_measure_canvas.bbox(self.txt_measure_canvas_text)
        return b[2] - b[0], b[3] - b[1]

    def get_lines_cell_height(self, n: int, font: None | FontTuple = None) -> int:
        if font == self.PAR.ops.table_font:
            return 3 + (n * self.table_txt_height)

        elif font == self.PAR.ops.index_font:
            return 3 + (n * self.index_txt_height)

        elif font == self.PAR.ops.header_font:
            return 3 + (n * self.header_txt_height)

        else:
            return 3 + self.get_txt_h(
                txt="\n".join("|" for _ in range(n)) if n > 1 else "|",
                font=font,
            )

    def set_min_column_width(self, width: int) -> None:
        if width:
            self.PAR.ops.min_column_width = width
        if self.PAR.ops.min_column_width > self.PAR.ops.max_column_width:
            self.PAR.ops.max_column_width = self.PAR.ops.min_column_width + 20
        if (
            isinstance(self.PAR.ops.auto_resize_columns, (int, float))
            and self.PAR.ops.auto_resize_columns < self.PAR.ops.min_column_width
        ):
            self.PAR.ops.auto_resize_columns = self.PAR.ops.min_column_width

    def get_default_row_height(self) -> int:
        if isinstance(self.PAR.ops.default_row_height, str):
            if int(self.PAR.ops.default_row_height) == 1:
                return self.min_row_height
            else:
                return self.min_row_height + self.get_lines_cell_height(int(self.PAR.ops.default_row_height) - 1)
        return self.PAR.ops.default_row_height

    def get_default_header_height(self) -> int:
        if isinstance(self.PAR.ops.default_header_height, str):
            if int(self.PAR.ops.default_header_height) == 1:
                return self.min_header_height
            else:
                return self.min_header_height + self.get_lines_cell_height(
                    int(self.PAR.ops.default_header_height) - 1, font=self.PAR.ops.header_font
                )
        return self.PAR.ops.default_header_height

    def check_font(self, newfont: tuple) -> None:
        if (
            not isinstance(newfont, tuple)
            or len(newfont) != 3
            or not isinstance(newfont[0], str)
            or not isinstance(newfont[1], (int, float))
            or not isinstance(newfont[2], str)
        ):
            raise ValueError(font_value_error)

    def set_table_font(self, newfont: tuple | None = None, row_heights: bool = True) -> tuple[str, int, str]:
        if newfont:
            self.check_font(newfont)
            self.PAR.ops.table_font = FontTuple(*(newfont[0], int(round(newfont[1])), newfont[2]))
            old_min_row_height = int(self.min_row_height)
            old_default_row_height = int(self.get_default_row_height())
            self.set_table_font_help()
            if row_heights:
                self.change_font_manage_row_heights(old_min_row_height, old_default_row_height)
                self.recreate_all_selection_boxes()
        return self.PAR.ops.table_font

    def set_table_font_help(self):
        self.table_font = self.PAR.ops.table_font
        if self.PAR.ops.table_font not in self.char_widths:
            self.char_widths[self.PAR.ops.table_font] = {}
        self.table_test_str_w = self.get_txt_w(_test_str)
        self.table_txt_width, self.table_txt_height = self.get_txt_dimensions("|", self.PAR.ops.table_font)
        self.min_row_height = max(6, self.table_txt_height, self.index_txt_height) + 6

        # dropdown stuff
        mod = (self.table_txt_height - 1) if self.table_txt_height % 2 else self.table_txt_height
        small_mod = int(mod / 5)
        mid_y = int(self.min_row_height / 2)
        self.dd_up_arrow = (
            4 + 4 * small_mod,
            mid_y + small_mod,
            4 + 2 * small_mod,
            mid_y - small_mod,
            4,
            mid_y + small_mod,
        )
        self.dd_down_arrow = (
            4 + 4 * small_mod,
            mid_y - small_mod,
            4 + 2 * small_mod,
            mid_y + small_mod,
            4,
            mid_y - small_mod,
        )

    def set_index_font(self, newfont: tuple | None = None, row_heights: bool = True) -> tuple[str, int, str]:
        if newfont:
            self.check_font(newfont)
            self.PAR.ops.index_font = FontTuple(*(newfont[0], int(round(newfont[1])), newfont[2]))
            old_min_row_height = int(self.min_row_height)
            old_default_row_height = int(self.get_default_row_height())
            self.set_index_font_help()
            if row_heights:
                self.change_font_manage_row_heights(old_min_row_height, old_default_row_height)
                self.recreate_all_selection_boxes()
        return self.PAR.ops.index_font

    def set_index_font_help(self) -> None:
        self.RI.index_font = self.PAR.ops.index_font
        if self.PAR.ops.index_font not in self.char_widths:
            self.char_widths[self.PAR.ops.index_font] = {}
        self.RI.index_test_str_w = self.get_txt_w(_test_str, self.PAR.ops.index_font)
        self.index_txt_width, self.index_txt_height = self.get_txt_dimensions("|", self.PAR.ops.index_font)
        self.min_row_height = max(6, self.table_txt_height, self.index_txt_height) + 6

    def change_font_manage_row_heights(self, old_min_row_height: int, old_default_row_height: int) -> None:
        default_row_height = self.get_default_row_height()
        self.row_positions = list(
            accumulate(
                chain(
                    [0],
                    (
                        (
                            self.min_row_height
                            if h == old_min_row_height
                            else (
                                default_row_height
                                if h == old_default_row_height
                                else self.min_row_height
                                if h < self.min_row_height
                                else h
                            )
                        )
                        for h in self.gen_row_heights()
                    ),
                )
            )
        )

    def set_header_font(self, newfont: tuple | None = None) -> tuple[str, int, str]:
        if newfont:
            self.check_font(newfont)
            self.PAR.ops.header_font = FontTuple(*(newfont[0], int(round(newfont[1])), newfont[2]))
            self.set_header_font_help()
            self.recreate_all_selection_boxes()
        return self.PAR.ops.header_font

    def set_header_font_help(self) -> None:
        self.CH.header_font = self.PAR.ops.header_font
        if self.PAR.ops.header_font not in self.char_widths:
            self.char_widths[self.PAR.ops.header_font] = {}
        self.CH.header_test_str_w = self.get_txt_w(_test_str, self.PAR.ops.header_font)
        self.header_txt_width, self.header_txt_height = self.get_txt_dimensions("|", self.PAR.ops.header_font)
        self.min_header_height = self.header_txt_height + 6
        if (
            isinstance(self.PAR.ops.default_header_height, int)
            and self.PAR.ops.default_header_height < self.min_header_height
        ):
            self.PAR.ops.default_header_height = int(self.min_header_height)
        self.CH.set_height(self.get_default_header_height(), set_TL=True)

    def purge_undo_and_redo_stack(self) -> None:
        self.undo_stack = deque(maxlen=self.PAR.ops.max_undos)
        self.redo_stack = deque(maxlen=self.PAR.ops.max_undos)

    def purge_redo_stack(self) -> None:
        self.redo_stack = deque(maxlen=self.PAR.ops.max_undos)

    def data_reference(
        self,
        newdataref: list | tuple | None = None,
        reset_col_positions: bool = True,
        reset_row_positions: bool = True,
        redraw: bool = False,
        return_id: bool = True,
        keep_formatting: bool = True,
    ) -> Any:
        if isinstance(newdataref, (list, tuple)):
            self.hide_dropdown_editor_all_canvases()
            self.data = newdataref
            if keep_formatting:
                self.reapply_formatting()
            else:
                self.delete_all_formatting(clear_values=False)
            self.purge_undo_and_redo_stack()
            if reset_col_positions:
                self.reset_col_positions()
            if reset_row_positions:
                self.reset_row_positions()
            if redraw:
                self.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
        if return_id:
            return id(self.data)
        else:
            return self.data

    def get_cell_dimensions(self, datarn: int, datacn: int) -> tuple[int, int]:
        txt = self.cell_str(datarn, datacn, get_displayed=True)
        if txt:
            self.txt_measure_canvas.itemconfig(self.txt_measure_canvas_text, text=txt, font=self.PAR.ops.table_font)
            b = self.txt_measure_canvas.bbox(self.txt_measure_canvas_text)
            w = b[2] - b[0] + 7
            h = b[3] - b[1] + 5
        else:
            w = self.PAR.ops.min_column_width
            h = self.min_row_height
        if self.get_cell_kwargs(datarn, datacn, key="dropdown") or self.get_cell_kwargs(datarn, datacn, key="checkbox"):
            return w + self.table_txt_height, h
        return w, h

    def set_cell_size_to_text(
        self, r: int, c: int, only_if_too_small: bool = False, redraw: bool = True, run_binding: bool = False
    ) -> bool:
        min_column_width = self.PAR.ops.min_column_width
        min_rh = self.min_row_height
        w = min_column_width
        h = min_rh
        datacn = self.datacn(c)
        datarn = self.datarn(r)
        tw, h = self.get_cell_dimensions(datarn, datacn)
        if tw > w:
            w = tw
        if h < min_rh:
            h = min_rh
        elif h > self.PAR.ops.max_row_height:
            h = int(self.PAR.ops.max_row_height)
        if w < min_column_width:
            w = min_column_width
        elif w > self.PAR.ops.max_column_width:
            w = int(self.PAR.ops.max_column_width)
        cell_needs_resize_w = False
        cell_needs_resize_h = False
        if only_if_too_small:
            if w > self.col_positions[c + 1] - self.col_positions[c]:
                cell_needs_resize_w = True
            if h > self.row_positions[r + 1] - self.row_positions[r]:
                cell_needs_resize_h = True
        else:
            if w != self.col_positions[c + 1] - self.col_positions[c]:
                cell_needs_resize_w = True
            if h != self.row_positions[r + 1] - self.row_positions[r]:
                cell_needs_resize_h = True
        if cell_needs_resize_w:
            old_width = self.col_positions[c + 1] - self.col_positions[c]
            new_col_pos = self.col_positions[c] + w
            increment = new_col_pos - self.col_positions[c + 1]
            self.col_positions[c + 2 :] = [
                e + increment for e in islice(self.col_positions, c + 2, len(self.col_positions))
            ]
            self.col_positions[c + 1] = new_col_pos
            new_width = self.col_positions[c + 1] - self.col_positions[c]
            if run_binding and self.CH.column_width_resize_func and old_width != new_width:
                self.CH.column_width_resize_func(
                    event_dict(
                        name="resize",
                        sheet=self.PAR.name,
                        resized_columns={c: {"old_size": old_width, "new_size": new_width}},
                    )
                )
        if cell_needs_resize_h:
            old_height = self.row_positions[r + 1] - self.row_positions[r]
            new_row_pos = self.row_positions[r] + h
            increment = new_row_pos - self.row_positions[r + 1]
            self.row_positions[r + 2 :] = [
                e + increment for e in islice(self.row_positions, r + 2, len(self.row_positions))
            ]
            self.row_positions[r + 1] = new_row_pos
            new_height = self.row_positions[r + 1] - self.row_positions[r]
            if run_binding and self.RI.row_height_resize_func and old_height != new_height:
                self.RI.row_height_resize_func(
                    event_dict(
                        name="resize",
                        sheet=self.PAR.name,
                        resized_rows={r: {"old_size": old_height, "new_size": new_height}},
                    )
                )
        if cell_needs_resize_w or cell_needs_resize_h:
            self.recreate_all_selection_boxes()
            if redraw:
                self.refresh()
                return True
            else:
                return False
        return False

    def get_cell_max_width(self, datarn: int, dispcn: int) -> int:
        datacn = self.datacn(dispcn)
        col_width = self.col_positions[dispcn + 1] - self.col_positions[dispcn]
        if kwargs := self.get_cell_kwargs(datarn, datacn, "dropdown"):
            max_width = col_width - self.table_txt_height - 5
        else:
            max_width = col_width - 2
            if (kwargs := self.get_cell_kwargs(datarn, datacn, "dropdown")) and max_width > self.table_txt_height + 1:
                box_w = self.table_txt_height + 1
                max_width -= box_w + 4
        if self.PAR.ops.allow_cell_overflow and not kwargs:
            if self.cells_cache is None:
                disprn = self.disprn(datarn)
                self.cells_cache = self._redraw_precache_cells(disprn, disprn + 1, 0, len(self.col_positions) - 1)
            if not (align := self.get_cell_kwargs(datarn, datacn, key="align")):
                align = self.align
            if align[-1] == "w":
                max_width += sum(
                    self._overflow(
                        self.cells_cache,
                        range(dispcn + 1, len(self.col_positions) - 1),
                        datarn,
                    )
                )
            elif align[-1] == "e":
                max_width += sum(
                    self._overflow(
                        self.cells_cache,
                        reversed(range(0, dispcn)),
                        datarn,
                    )
                )
        return max_width

    def get_wrapped_cell_height(self, datarn: int, datacn: int) -> int:
        dispcn = self.dispcn(datacn)
        n_lines = max(
            1,
            sum(
                1
                for _ in wrap_text(
                    text=self.cell_str(datarn, datacn, get_displayed=True),
                    max_width=self.get_cell_max_width(datarn, dispcn),
                    max_lines=float("inf"),
                    char_width_fn=self.char_width_fn,
                    widths=self.char_widths[self.table_font],
                    wrap=self.PAR.ops.table_wrap,
                )
            ),
        )
        return 3 + (n_lines * self.table_txt_height)

    def set_all_cell_sizes_to_text(
        self,
        width: int | None = None,
        slim: bool = False,
    ) -> tuple[list[float], list[float]]:
        min_column_width = self.PAR.ops.min_column_width
        max_column_width = float_to_int(self.PAR.ops.max_column_width)
        max_row_height = float_to_int(self.PAR.ops.max_row_height)
        min_rh = self.min_row_height
        h = min_rh
        rhs = defaultdict(lambda: int(min_rh))
        cws = []
        qconf = self.txt_measure_canvas.itemconfig
        qbbox = self.txt_measure_canvas.bbox
        qtxtm = self.txt_measure_canvas_text
        qtxth = self.table_txt_height
        qfont = self.PAR.ops.table_font
        numrows = self.total_data_rows()
        numcols = self.total_data_cols()
        itercols = range(numcols) if self.all_columns_displayed else self.displayed_columns
        iterrows = range(numrows) if self.all_rows_displayed else self.displayed_rows
        if is_iterable(self._row_index):
            for datarn in iterrows:
                w_, h = self.RI.get_cell_dimensions(datarn)
                if h < min_rh:
                    h = min_rh
                elif h > max_row_height:
                    h = max_row_height
                if h > rhs[datarn]:
                    rhs[datarn] = h
        added_w_space = 1 if slim else 7
        for datacn in itercols:
            w = min_column_width if width is None else width
            hw = self.CH.get_cell_dimensions(datacn)[0]
            if hw > w:
                w = hw
            for datarn in iterrows:
                if txt := self.cell_str(datarn, datacn, get_displayed=True):
                    qconf(qtxtm, text=txt, font=qfont)
                    b = qbbox(qtxtm)
                    tw = b[2] - b[0] + added_w_space
                    h = b[3] - b[1] + 5
                else:
                    tw = min_column_width
                    h = min_rh
                # self.get_cell_kwargs not used here to boost performance
                if (
                    (datarn, datacn) in self.cell_options
                    and "dropdown" in self.cell_options[(datarn, datacn)]
                    or datarn in self.row_options
                    and "dropdown" in self.row_options[datarn]
                    or datacn in self.col_options
                    and "dropdown" in self.col_options[datacn]
                ) or (
                    (datarn, datacn) in self.cell_options
                    and "checkbox" in self.cell_options[(datarn, datacn)]
                    or datarn in self.row_options
                    and "checkbox" in self.row_options[datarn]
                    or datacn in self.col_options
                    and "checkbox" in self.col_options[datacn]
                ):
                    tw += qtxth
                if tw > w:
                    w = tw
                if h < min_rh:
                    h = min_rh
                elif h > max_row_height:
                    h = max_row_height
                if h > rhs[datarn]:
                    rhs[datarn] = h
            if w < min_column_width:
                w = min_column_width
            elif w > max_column_width:
                w = max_column_width
            cws.append(w)
        self.set_row_positions(itr=rhs.values())
        self.set_col_positions(itr=cws)
        self.recreate_all_selection_boxes()
        return self.row_positions, self.col_positions

    def set_col_positions(self, itr: Iterator[float]) -> None:
        self.col_positions = list(accumulate(chain([0], itr)))

    def reset_col_positions(self, ncols: int | None = None, width: int | None = None) -> None:
        if width is None:
            width = self.PAR.ops.default_column_width
        if isinstance(ncols, int):
            self.set_col_positions(itr=repeat(width, ncols))
        elif self.all_columns_displayed:
            self.set_col_positions(itr=repeat(width, self.total_data_cols()))
        else:
            self.set_col_positions(itr=repeat(width, len(self.displayed_columns)))

    def set_row_positions(self, itr: Iterator[float]) -> None:
        self.row_positions = list(accumulate(chain([0], itr)))

    def reset_row_positions(self, nrows: int | None = None, height: int | None = None) -> None:
        if height is None:
            height = self.get_default_row_height()
        if isinstance(nrows, int):
            self.set_row_positions(itr=repeat(height, nrows))
        elif self.all_rows_displayed:
            self.set_row_positions(itr=repeat(height, self.total_data_rows()))
        else:
            self.set_row_positions(itr=repeat(height, len(self.displayed_rows)))

    def del_col_position(self, idx: int, deselect_all: bool = False) -> None:
        if deselect_all:
            self.deselect("all", redraw=False)
        if idx == "end" or len(self.col_positions) <= idx + 1:
            del self.col_positions[-1]
        else:
            w = self.col_positions[idx + 1] - self.col_positions[idx]
            idx += 1
            del self.col_positions[idx]
            self.col_positions[idx:] = [e - w for e in islice(self.col_positions, idx, len(self.col_positions))]

    def del_row_position(self, idx: int, deselect_all: bool = False) -> None:
        if deselect_all:
            self.deselect("all", redraw=False)
        if idx == "end" or len(self.row_positions) <= idx + 1:
            del self.row_positions[-1]
        else:
            w = self.row_positions[idx + 1] - self.row_positions[idx]
            idx += 1
            del self.row_positions[idx]
            self.row_positions[idx:] = [e - w for e in islice(self.row_positions, idx, len(self.row_positions))]

    def del_col_positions(self, idxs: Iterator[int] | None = None) -> None:
        if idxs is None:
            del self.col_positions[-1]
        else:
            if not isinstance(idxs, set):
                idxs = set(idxs)
            self.set_col_positions(itr=(w for i, w in enumerate(self.gen_column_widths()) if i not in idxs))

    def del_row_positions(self, idxs: Iterator[int] | None = None) -> None:
        if idxs is None:
            del self.row_positions[-1]
        else:
            if not isinstance(idxs, set):
                idxs = set(idxs)
            self.set_row_positions(itr=(h for i, h in enumerate(self.gen_row_heights()) if i not in idxs))

    def get_column_widths(self) -> list[int]:
        return list(diff_gen(self.col_positions))

    def get_row_heights(self) -> list[int]:
        return list(diff_gen(self.row_positions))

    def gen_column_widths(self) -> Generator[int]:
        return diff_gen(self.col_positions)

    def gen_row_heights(self) -> Generator[int]:
        return diff_gen(self.row_positions)

    def insert_col_positions(
        self,
        idx: Literal["end"] | int = "end",
        widths: Sequence[float] | int | None = None,
        deselect_all: bool = False,
    ) -> None:
        if deselect_all:
            self.deselect("all", redraw=False)
        if widths is None:
            w = [self.PAR.ops.default_column_width]
        elif isinstance(widths, int):
            w = list(repeat(self.PAR.ops.default_column_width, widths))
        else:
            w = widths
        if idx == "end" or len(self.col_positions) == idx + 1:
            if len(w) > 1:
                self.col_positions += list(accumulate(chain([self.col_positions[-1] + w[0]], islice(w, 1, None))))
            else:
                self.col_positions.append(self.col_positions[-1] + w[0])
        else:
            if len(w) > 1:
                idx += 1
                self.col_positions[idx:idx] = list(
                    accumulate(chain([self.col_positions[idx - 1] + w[0]], islice(w, 1, None)))
                )
                idx += len(w)
                sumw = sum(w)
                self.col_positions[idx:] = [e + sumw for e in islice(self.col_positions, idx, len(self.col_positions))]
            else:
                w = w[0]
                idx += 1
                self.col_positions.insert(idx, self.col_positions[idx - 1] + w)
                idx += 1
                self.col_positions[idx:] = [e + w for e in islice(self.col_positions, idx, len(self.col_positions))]

    def insert_row_positions(
        self,
        idx: Literal["end"] | int = "end",
        heights: Sequence[float] | int | None = None,
        deselect_all: bool = False,
    ) -> None:
        if deselect_all:
            self.deselect("all", redraw=False)
        default_row_height = self.get_default_row_height()
        if heights is None:
            h = [default_row_height]
        elif isinstance(heights, int):
            h = list(repeat(default_row_height, heights))
        else:
            h = heights
        if idx == "end" or len(self.row_positions) == idx + 1:
            if len(h) > 1:
                self.row_positions += list(accumulate(chain([self.row_positions[-1] + h[0]], islice(h, 1, None))))
            else:
                self.row_positions.append(self.row_positions[-1] + h[0])
        else:
            if len(h) > 1:
                idx += 1
                self.row_positions[idx:idx] = list(
                    accumulate(chain([self.row_positions[idx - 1] + h[0]], islice(h, 1, None)))
                )
                idx += len(h)
                sumh = sum(h)
                self.row_positions[idx:] = [e + sumh for e in islice(self.row_positions, idx, len(self.row_positions))]
            else:
                h = h[0]
                idx += 1
                self.row_positions.insert(idx, self.row_positions[idx - 1] + h)
                idx += 1
                self.row_positions[idx:] = [e + h for e in islice(self.row_positions, idx, len(self.row_positions))]

    def named_span_coords(self, name: str | dict) -> tuple[int, int, int | None, int | None]:
        dct = self.named_spans[name] if isinstance(name, str) else name
        return (
            dct["from_r"],
            dct["from_c"],
            dct["upto_r"],
            dct["upto_c"],
        )

    def adjust_options_post_add_columns(
        self,
        cols: list | tuple,
        create_ops: bool = True,
    ) -> None:
        self.tagged_cells = {
            tags: {(r, push_n(c, cols)) for (r, c) in tagged} for tags, tagged in self.tagged_cells.items()
        }
        self.cell_options = {(r, push_n(c, cols)): v for (r, c), v in self.cell_options.items()}
        self.progress_bars = {(r, push_n(c, cols)): v for (r, c), v in self.progress_bars.items()}
        self.tagged_columns = {tags: {push_n(c, cols) for c in tagged} for tags, tagged in self.tagged_columns.items()}
        self.col_options = {push_n(c, cols): v for c, v in self.col_options.items()}
        self.CH.cell_options = {push_n(c, cols): v for c, v in self.CH.cell_options.items()}
        # if there are named spans where columns were added
        # add options to gap which was created by adding columns
        totalrows = None
        new_ops = partial(self.PAR.create_options_from_span, set_data=False)
        qkspan = self.span()
        for span in self.named_spans.values():
            if isinstance(span["from_c"], int):
                for datacn in cols:
                    if span["from_c"] > datacn:
                        span["from_c"] += 1
                        if isinstance(span["upto_c"], int):
                            span["upto_c"] += 1
                    elif span["from_c"] <= datacn and (
                        (isinstance(span["upto_c"], int) and span["upto_c"] > datacn) or span["upto_c"] is None
                    ):
                        if isinstance(span["upto_c"], int):
                            span["upto_c"] += 1
                        # if to_add then it's an undo/redo and don't
                        # need to create fresh options
                        if create_ops:
                            # if rows are none it's a column options span
                            if span["from_r"] is None:
                                new_ops(
                                    mod_span(
                                        qkspan,
                                        span,
                                        from_c=datacn,
                                        upto_c=datacn + 1,
                                    )
                                )
                            # cells
                            else:
                                if totalrows is None:
                                    totalrows = self.total_data_rows()
                                rng_upto_r = totalrows if span["upto_r"] is None else span["upto_r"]
                                for rn in range(span["from_r"], rng_upto_r):
                                    new_ops(
                                        mod_span(
                                            qkspan,
                                            span,
                                            from_r=rn,
                                            from_c=datacn,
                                            upto_r=rn + 1,
                                            upto_c=datacn + 1,
                                        )
                                    )

    def adjust_options_post_add_rows(
        self,
        rows: list[int] | tuple[int],
        create_ops: bool = True,
        tree: bool = True,
    ) -> None:
        self.tagged_cells = {
            tags: {(push_n(r, rows), c) for (r, c) in tagged} for tags, tagged in self.tagged_cells.items()
        }
        self.cell_options = {(push_n(r, rows), c): v for (r, c), v in self.cell_options.items()}
        self.progress_bars = {(push_n(r, rows), c): v for (r, c), v in self.progress_bars.items()}
        self.tagged_rows = {tags: {push_n(r, rows) for r in tagged} for tags, tagged in self.tagged_rows.items()}
        self.row_options = {push_n(r, rows): v for r, v in self.row_options.items()}
        self.RI.cell_options = {push_n(r, rows): v for r, v in self.RI.cell_options.items()}
        if tree:
            self.RI.rns = {k: push_n(r, rows) for k, r in self.RI.rns.items()}
        # if there are named spans where rows were added
        # add options to gap which was created by adding rows
        totalcols = None
        new_ops = partial(self.PAR.create_options_from_span, set_data=False)
        qkspan = self.span()
        for span in self.named_spans.values():
            if isinstance(span["from_r"], int):
                for datarn in rows:
                    if span["from_r"] > datarn:
                        span["from_r"] += 1
                        if isinstance(span["upto_r"], int):
                            span["upto_r"] += 1
                    elif span["from_r"] <= datarn and (
                        (isinstance(span["upto_r"], int) and span["upto_r"] > datarn) or span["upto_r"] is None
                    ):
                        if isinstance(span["upto_r"], int):
                            span["upto_r"] += 1
                        # if to_add then it's an undo/redo and don't
                        # need to create fresh options
                        if create_ops:
                            # if cols are none it's a row options span
                            if span["from_c"] is None:
                                new_ops(
                                    mod_span(
                                        qkspan,
                                        span,
                                        from_r=datarn,
                                        upto_r=datarn + 1,
                                    )
                                )
                            # cells
                            else:
                                if totalcols is None:
                                    totalcols = self.total_data_cols()
                                rng_upto_c = totalcols if span["upto_c"] is None else span["upto_c"]
                                for cn in range(span["from_c"], rng_upto_c):
                                    new_ops(
                                        mod_span(
                                            qkspan,
                                            span,
                                            from_r=datarn,
                                            from_c=cn,
                                            upto_r=datarn + 1,
                                            upto_c=cn + 1,
                                        )
                                    )

    def adjust_options_post_delete_columns(
        self,
        to_del: None | set[int] = None,
        to_bis: None | list[int] = None,
        named_spans: None | set[str] = None,
    ) -> list[int]:
        if to_del is None:
            to_del = set()
        if not to_bis:
            to_bis = sorted(to_del)
        self.tagged_cells = {
            tags: {
                (
                    r,
                    c if not (num := bisect_left(to_bis, c)) else c - num,
                )
                for (r, c) in tagged
                if c not in to_del
            }
            for tags, tagged in self.tagged_cells.items()
        }
        self.cell_options = {
            (
                r,
                c if not (num := bisect_left(to_bis, c)) else c - num,
            ): v
            for (r, c), v in self.cell_options.items()
            if c not in to_del
        }
        self.progress_bars = {
            (
                r,
                c if not (num := bisect_left(to_bis, c)) else c - num,
            ): v
            for (r, c), v in self.progress_bars.items()
            if c not in to_del
        }
        self.tagged_columns = {
            tags: {c if not (num := bisect_left(to_bis, c)) else c - num for c in tagged if c not in to_del}
            for tags, tagged in self.tagged_columns.items()
        }
        self.col_options = {
            c if not (num := bisect_left(to_bis, c)) else c - num: v
            for c, v in self.col_options.items()
            if c not in to_del
        }
        self.CH.cell_options = {
            c if not (num := bisect_left(to_bis, c)) else c - num: v
            for c, v in self.CH.cell_options.items()
            if c not in to_del
        }
        self.del_columns_from_named_spans(
            to_del=to_del,
            to_bis=to_bis,
            named_spans=named_spans,
        )
        return to_bis

    def del_columns_from_named_spans(
        self,
        to_del: set[int],
        to_bis: list,
        named_spans: None | set = None,
    ) -> None:
        if named_spans is None:
            named_spans = self.get_spans_to_del_from_cols(cols=to_del)
        for name in named_spans:
            del self.named_spans[name]
        for span in self.named_spans.values():
            if isinstance(span["from_c"], int):
                for c in to_bis:
                    if span["from_c"] > c:
                        span["from_c"] -= 1
                    if isinstance(span["upto_c"], int) and span["upto_c"] > c:
                        span["upto_c"] -= 1

    def get_spans_to_del_from_cols(self, cols: set[int]) -> set[str]:
        total = self.total_data_cols()
        return {
            nm
            for nm, sp in self.named_spans.items()
            if isinstance(sp["from_c"], int)
            and all(c in cols for c in range(sp["from_c"], total if sp["upto_c"] is None else sp["upto_c"]))
        }

    def adjust_options_post_delete_rows(
        self,
        to_del: None | set[int] = None,
        to_bis: None | list[int] = None,
        named_spans: None | set[str] = None,
    ) -> list[int]:
        if to_del is None:
            to_del = set()
        if not to_bis:
            to_bis = sorted(to_del)
        self.tagged_cells = {
            tags: {
                (
                    r if not (num := bisect_left(to_bis, r)) else r - num,
                    c,
                )
                for (r, c) in tagged
                if r not in to_del
            }
            for tags, tagged in self.tagged_cells.items()
        }
        self.cell_options = {
            (
                r if not (num := bisect_left(to_bis, r)) else r - num,
                c,
            ): v
            for (r, c), v in self.cell_options.items()
            if r not in to_del
        }
        self.progress_bars = {
            (
                r if not (num := bisect_left(to_bis, r)) else r - num,
                c,
            ): v
            for (r, c), v in self.progress_bars.items()
            if r not in to_del
        }
        self.tagged_rows = {
            tags: {r if not (num := bisect_left(to_bis, r)) else r - num for r in tagged if r not in to_del}
            for tags, tagged in self.tagged_rows.items()
        }
        self.row_options = {
            r if not (num := bisect_left(to_bis, r)) else r - num: v
            for r, v in self.row_options.items()
            if r not in to_del
        }
        self.RI.cell_options = {
            r if not (num := bisect_left(to_bis, r)) else r - num: v
            for r, v in self.RI.cell_options.items()
            if r not in to_del
        }
        self.RI.rns = {
            v: r if not (num := bisect_left(to_bis, r)) else r - num for v, r in self.RI.rns.items() if r not in to_del
        }
        self.del_rows_from_named_spans(
            to_del=to_del,
            to_bis=to_bis,
            named_spans=named_spans,
        )
        return to_bis

    def del_rows_from_named_spans(
        self,
        to_del: set,
        to_bis: list,
        named_spans: None | set[str] = None,
    ) -> None:
        if named_spans is None:
            named_spans = self.get_spans_to_del_from_rows(rows=to_del)
        for name in named_spans:
            del self.named_spans[name]
        for span in self.named_spans.values():
            if isinstance(span["from_r"], int):
                for r in to_bis:
                    if span["from_r"] > r:
                        span["from_r"] -= 1
                    if isinstance(span["upto_r"], int) and span["upto_r"] > r:
                        span["upto_r"] -= 1

    def get_spans_to_del_from_rows(self, rows: set[int]) -> set[str]:
        total = self.total_data_rows()
        return {
            nm
            for nm, sp in self.named_spans.items()
            if isinstance(sp["from_r"], int)
            and all(r in rows for r in range(sp["from_r"], total if sp["upto_r"] is None else sp["upto_r"]))
        }

    def add_columns(
        self,
        columns: dict,
        header: dict,
        column_widths: dict,
        event_data: dict,
        create_ops: bool = True,
        create_selections: bool = True,
        add_row_positions: bool = True,
        push_ops: bool = True,
        mod_event_boxes: bool = True,
        from_undo: bool = False,
    ) -> EventDataDict | None:
        if not from_undo and not try_binding(self.extra_begin_insert_cols_rc_func, event_data, "begin_add_columns"):
            return
        self.saved_column_widths = {}
        if not from_undo and not self.all_columns_displayed:
            self.displayed_columns = add_to_displayed(self.displayed_columns, columns)
        cws = self.get_column_widths()
        if column_widths and next(iter(column_widths)) > len(cws):
            default_width = self.PAR.ops.default_column_width
            to_add = dict.fromkeys(range(len(cws), next(iter(column_widths))), default_width)
            to_add.update(column_widths)
            column_widths = to_add
        if not from_undo:
            self.set_col_positions(insert_items(cws, column_widths))
        # rn needed for indexing but cn insert
        maxrn = 0
        for cn, rowdict in columns.items():
            for rn, v in rowdict.items():
                if rn < len(self.data) and cn > len(self.data[rn]):
                    self.fix_row_len(rn, cn - 1)
                elif rn >= len(self.data):
                    self.fix_data_len(rn, cn - 1)
                if rn > maxrn:
                    maxrn = rn
                self.data[rn].insert(cn, v)
        # if not hiding rows then we can extend row positions if necessary
        if add_row_positions and self.all_rows_displayed and maxrn >= len(self.row_positions) - 1:
            default_height = self.get_default_row_height()
            event_data["added"]["rows"] = {
                "table": {},
                "index": {},
                "row_heights": dict.fromkeys(range(len(self.row_positions) - 1, maxrn + 1), default_height),
            }
            if not from_undo:
                self.set_row_positions(
                    chain(self.gen_row_heights(), repeat(default_height, maxrn + 1 - (len(self.row_positions) - 1)))
                )
        if isinstance(self._headers, list) and header:
            self._headers = insert_items(self._headers, header, self.CH.fix_header)
        if push_ops and columns:
            self.adjust_options_post_add_columns(
                cols=tuple(columns),
                create_ops=create_ops,
            )
        if create_selections:
            self.deselect("all", redraw=False)
            for boxst, boxend in consecutive_ranges(tuple(column_widths)):
                self.create_selection_box(
                    0,
                    boxst,
                    len(self.row_positions) - 1,
                    boxend,
                    "columns",
                    run_binding=True,
                )
            if mod_event_boxes:
                event_data["selection_boxes"] = self.get_boxes()
                event_data["selected"] = self.selected
        event_data["added"]["columns"] = {
            "table": columns,
            "header": header,
            "column_widths": column_widths,
        }
        if not from_undo:
            try_binding(self.extra_end_insert_cols_rc_func, event_data, "end_add_columns")
        return event_data

    def rc_add_columns(self, event: Any = None) -> None:
        rowlen = self.equalize_data_row_lengths()
        selcols = sorted(self.get_selected_cols())
        if (
            selcols
            and isinstance(self.CH.popup_menu_loc, int)
            and selcols[0] <= self.CH.popup_menu_loc
            and selcols[-1] >= self.CH.popup_menu_loc
        ):
            selcols = get_seq_without_gaps_at_index(selcols, self.CH.popup_menu_loc)
            numcols = len(selcols)
            displayed_ins_col = selcols[0] if event == "left" else selcols[-1] + 1
            if self.all_columns_displayed:
                data_ins_col = int(displayed_ins_col)
            else:
                if displayed_ins_col == len(self.col_positions) - 1:
                    data_ins_col = rowlen
                else:
                    if len(self.displayed_columns) > displayed_ins_col:
                        data_ins_col = int(self.displayed_columns[displayed_ins_col])
                    else:
                        data_ins_col = int(self.displayed_columns[-1])
        else:
            numcols = 1
            if event == "left":
                displayed_ins_col = 0
                data_ins_col = 0
            else:
                displayed_ins_col = len(self.col_positions) - 1
                data_ins_col = self.total_data_cols()
        if (
            isinstance(self.PAR.ops.paste_insert_column_limit, int)
            and self.PAR.ops.paste_insert_column_limit < displayed_ins_col + numcols
        ):
            numcols = self.PAR.ops.paste_insert_column_limit - len(self.col_positions) - 1
            if numcols < 1:
                return
        event_data = self.new_event_dict("add_columns", state=True)
        columns, headers, widths = self.get_args_for_add_columns(data_ins_col, displayed_ins_col, numcols)
        event_data = self.add_columns(
            columns=columns,
            header=headers if isinstance(self._headers, list) and self._headers else {},
            column_widths=widths,
            event_data=event_data,
        )
        if event_data:
            if self.undo_enabled:
                self.undo_stack.append(stored_event_dict(event_data))
            self.refresh()
            self.sheet_modified(event_data)

    def add_rows(
        self,
        rows: dict[int, list[Any]],
        index: dict[int, Any],
        row_heights: dict[int, float | int],
        event_data: EventDataDict,
        create_ops: bool = True,
        create_selections: bool = True,
        add_col_positions: bool = True,
        push_ops: bool = True,
        tree: bool = True,
        mod_event_boxes: bool = True,
        from_undo: bool = False,
    ) -> EventDataDict | None:
        if not from_undo and not try_binding(self.extra_begin_insert_rows_rc_func, event_data, "begin_add_rows"):
            return
        self.saved_row_heights = {}
        if not from_undo and not self.all_rows_displayed and not tree:
            # we have to handle this separately for the treeview
            # based on whether the rows will be displayed
            self.displayed_rows = add_to_displayed(self.displayed_rows, rows)
        maxcn = 0
        # rn needed for insert but cn indexing
        for rn, row in rows.items():
            cn = len(row) - 1
            if rn > len(self.data):
                self.fix_data_len(rn - 1, cn)
            self.data.insert(rn, row)
            if cn > maxcn:
                maxcn = cn
        if isinstance(self._row_index, list) and index:
            self._row_index = insert_items(self._row_index, index, self.RI.fix_index)
        # if not hiding columns then we can extend col positions if necessary
        if add_col_positions and self.all_columns_displayed and maxcn >= len(self.col_positions) - 1:
            default_width = self.PAR.ops.default_column_width
            event_data["added"]["columns"] = {
                "table": {},
                "header": {},
                "column_widths": dict.fromkeys(range(len(self.col_positions) - 1, maxcn + 1), default_width),
            }
            if not from_undo:
                self.set_col_positions(
                    itr=chain(
                        self.gen_column_widths(),
                        repeat(default_width, maxcn + 1 - (len(self.col_positions) - 1)),
                    )
                )
        if push_ops and rows:
            self.adjust_options_post_add_rows(
                rows=tuple(rows),
                create_ops=create_ops,
                tree=tree,
            )
        event_data["added"]["rows"] = {
            "table": rows,
            "index": index,
            "row_heights": row_heights,
        }
        if tree and self.PAR.ops.treeview:
            # creates + overwrites row heights based on whether rows are showing
            event_data = self.RI.tree_add_rows(event_data=event_data)

        row_heights = event_data["added"]["rows"]["row_heights"]
        rhs = self.get_row_heights()
        if row_heights and next(iter(row_heights)) > len(rhs):
            default_height = self.get_default_row_height()
            to_add = dict.fromkeys(range(len(rhs), next(iter(row_heights))), default_height)
            to_add.update(row_heights)
            row_heights = to_add
        if not from_undo:
            self.set_row_positions(insert_items(rhs, row_heights))
        if create_selections:
            self.deselect("all", redraw=False)
            for boxst, boxend in consecutive_ranges(tuple(row_heights)):
                self.create_selection_box(
                    boxst,
                    0,
                    boxend,
                    len(self.col_positions) - 1,
                    "rows",
                    run_binding=True,
                )
            if mod_event_boxes:
                event_data["selection_boxes"] = self.get_boxes()
                event_data["selected"] = self.selected
        if not from_undo:
            try_binding(self.extra_end_insert_rows_rc_func, event_data, "end_add_rows")
        return event_data

    def rc_add_rows(self, event: Any = None) -> None:
        total_data_rows = self.total_data_rows()
        selrows = sorted(self.get_selected_rows())
        if (
            selrows
            and isinstance(self.RI.popup_menu_loc, int)
            and selrows[0] <= self.RI.popup_menu_loc
            and selrows[-1] >= self.RI.popup_menu_loc
        ):
            selrows = get_seq_without_gaps_at_index(selrows, self.RI.popup_menu_loc)
            numrows = len(selrows)
            displayed_ins_row = selrows[0] if event == "above" else selrows[-1] + 1
            if self.all_rows_displayed:
                data_ins_row = int(displayed_ins_row)
            else:
                if displayed_ins_row == len(self.row_positions) - 1:
                    data_ins_row = total_data_rows
                else:
                    if len(self.displayed_rows) > displayed_ins_row:
                        data_ins_row = int(self.displayed_rows[displayed_ins_row])
                    else:
                        data_ins_row = int(self.displayed_rows[-1])
        else:
            numrows = 1
            if event == "above":
                displayed_ins_row = 0
                data_ins_row = 0
            else:
                displayed_ins_row = len(self.row_positions) - 1
                data_ins_row = self.total_data_rows()
        if (
            isinstance(self.PAR.ops.paste_insert_row_limit, int)
            and self.PAR.ops.paste_insert_row_limit < displayed_ins_row + numrows
        ):
            numrows = self.PAR.ops.paste_insert_row_limit - len(self.row_positions) - 1
            if numrows < 1:
                return
        event_data = self.new_event_dict("add_rows", state=True)
        if self.PAR.ops.treeview:
            heights = []
        else:
            h = self.get_default_row_height()
            heights = repeat(h, numrows)
        rows, index, heights = self.get_args_for_add_rows(
            data_ins_row=data_ins_row,
            displayed_ins_row=displayed_ins_row,
            rows=numrows,
            heights=heights,
        )
        event_data = self.add_rows(
            rows=rows,
            index=index if isinstance(self._row_index, list) and self._row_index else {},
            row_heights=heights,
            event_data=event_data,
        )
        if event_data:
            if self.undo_enabled:
                self.undo_stack.append(stored_event_dict(event_data))
            self.refresh()
            self.sheet_modified(event_data)

    def get_args_for_add_columns(
        self,
        data_ins_col: int,
        displayed_ins_col: int,
        columns: list[list[Any]] | int,
        widths: list[int] | tuple[int] | None = None,
        headers: bool = False,
    ) -> tuple[dict, dict, dict]:
        header_dict = {}
        columns_dict = {}
        if isinstance(columns, int):
            total_data_rows = self.total_data_rows()
            rng = total_data_rows if total_data_rows else 1

            if isinstance(self._headers, list):
                for datacn in range(data_ins_col, data_ins_col + columns):
                    columns_dict[datacn] = {}
                    header_dict[datacn] = self.CH.get_value_for_empty_cell(datacn, c_ops=False)
                    for datarn in range(rng):
                        columns_dict[datacn][datarn] = self.get_value_for_empty_cell(datarn, datacn, c_ops=False)
            else:
                for datacn in range(data_ins_col, data_ins_col + columns):
                    columns_dict[datacn] = {}
                    for datarn in range(rng):
                        columns_dict[datacn][datarn] = self.get_value_for_empty_cell(datarn, datacn, c_ops=False)
        else:
            if isinstance(self._headers, list) and headers:
                for datacn, column in enumerate(columns, data_ins_col):
                    if column:
                        header_dict[datacn] = column[0]
                        columns_dict[datacn] = {
                            datarn: self.format_value(datarn, datacn, v)
                            for datarn, v in enumerate(islice(column, 1, None))
                        }
            else:
                for datacn, column in enumerate(columns, data_ins_col):
                    if column:
                        columns_dict[datacn] = {
                            datarn: self.format_value(datarn, datacn, v) for datarn, v in enumerate(column)
                        }

        rng = range(displayed_ins_col, displayed_ins_col + len(columns_dict))
        if widths is None:
            w = self.PAR.ops.default_column_width
            widths_dict = dict(zip(rng, repeat(w)))
        else:
            widths_dict = dict(zip(rng, widths))
        return columns_dict, header_dict, widths_dict

    def get_args_for_add_rows(
        self,
        data_ins_row: int,
        displayed_ins_row: int,
        rows: list[list[Any]] | int,
        heights: Iterator[int] | None = None,
        row_index: bool = False,
        total_data_cols: int | None = None,
    ) -> tuple[dict, dict, dict]:
        index_dict = {}
        rows_dict = {}
        if isinstance(rows, int):
            if total_data_cols is None:
                total_data_cols = self.total_data_cols()

            rng = total_data_cols if total_data_cols else 1
            if isinstance(self._row_index, list):
                tree = self.PAR.ops.treeview
                for datarn in range(data_ins_row, data_ins_row + rows):
                    rows_dict[datarn] = [self.get_value_for_empty_cell(datarn, c, c_ops=True) for c in range(rng)]
                    index_dict[datarn] = self.RI.get_value_for_empty_cell(data_ins_row if tree else datarn, r_ops=False)
            else:
                for datarn in range(data_ins_row, data_ins_row + rows):
                    rows_dict[datarn] = [self.get_value_for_empty_cell(datarn, c, c_ops=True) for c in range(rng)]
        else:
            if isinstance(self._row_index, list) and row_index:
                for datarn, row in enumerate(rows, data_ins_row):
                    if row:
                        index_dict[datarn] = row[0]
                    if len(row) > 1:
                        rows_dict[datarn] = [self.format_value(datarn, datacn, v) for datacn, v in enumerate(row[1:])]
                    else:
                        rows_dict[datarn] = []
            else:
                for datarn, row in enumerate(rows, data_ins_row):
                    rows_dict[datarn] = [self.format_value(datarn, datacn, v) for datacn, v in enumerate(row)]

        rng = range(displayed_ins_row, displayed_ins_row + len(rows_dict))
        if heights is None:
            h = self.get_default_row_height()
            heights_dict = dict(zip(rng, repeat(h)))
        else:
            heights_dict = dict(zip(rng, heights))
        return rows_dict, index_dict, heights_dict

    def copy_options(self) -> dict:
        return {
            "cell_options": dict(self.cell_options),
            "column_options": dict(self.col_options),
            "row_options": dict(self.row_options),
            "CH_cell_options": dict(self.CH.cell_options),
            "RI_cell_options": dict(self.RI.cell_options),
            "tagged_cells": {f"{tag}": set(s) for tag, s in self.tagged_cells.items()},
            "tagged_rows": {f"{tag}": set(s) for tag, s in self.tagged_rows.items()},
            "tagged_columns": {f"{tag}": set(s) for tag, s in self.tagged_columns.items()},
        }

    def copy_sheet_state(self) -> dict:
        return {
            "row_positions": list(self.row_positions),
            "col_positions": list(self.col_positions),
            "displayed_rows": list(self.displayed_rows),
            "displayed_columns": list(self.displayed_columns),
            "all_rows_displayed": bool(self.all_rows_displayed),
            "saved_row_heights": dict(self.saved_row_heights),
            "saved_column_widths": dict(self.saved_column_widths),
            "all_columns_displayed": bool(self.all_columns_displayed),
            "tree_open_ids": set(self.RI.tree_open_ids),
        }

    def delete_columns_data(self, cols: list[int], event_data: EventDataDict | None = None) -> EventDataDict:
        self.mouseclick_outside_editor_or_dropdown_all_canvases()
        if not event_data:
            event_data = self.new_event_dict("delete_columns", state=True)
        event_data["options"] = self.copy_options()
        event_data["named_spans"] = {k: span.copy_self() for k, span in self.named_spans.items()}
        for i, datacn in enumerate(cols):
            for rn in range(len(self.data)):
                if datacn not in event_data["deleted"]["columns"]:
                    event_data["deleted"]["columns"][datacn] = {}
                try:
                    event_data["deleted"]["columns"][datacn][rn] = self.data[rn].pop(datacn - i)
                except Exception:
                    continue
            try:
                event_data["deleted"]["header"][datacn] = self._headers.pop(datacn - i)
            except Exception:
                continue
        cols_set = set(cols)
        self.adjust_options_post_delete_columns(
            to_del=cols_set,
            to_bis=cols,
            named_spans=self.get_spans_to_del_from_cols(cols=cols_set),
        )
        if not self.all_columns_displayed:
            self.displayed_columns = [
                c if not (num := bisect_left(cols, c)) else c - num
                for c in filterfalse(cols_set.__contains__, self.displayed_columns)
            ]
        return event_data

    def delete_columns_displayed(
        self,
        cols: list[int],
        event_data: EventDataDict | None = None,
        from_undo: bool = False,
    ) -> EventDataDict:
        if not event_data:
            event_data = self.new_event_dict("delete_columns", state=True)
        self.saved_column_widths = {}
        if cols:
            for c in cols:
                if len(self.col_positions) > c + 1:
                    event_data["deleted"]["column_widths"][c] = self.col_positions[c + 1] - self.col_positions[c]
            if not from_undo:
                cols_set = set(cols)
                self.set_col_positions(
                    itr=(width for c, width in enumerate(self.gen_column_widths()) if c not in cols_set)
                )
        return event_data

    def delete_columns(
        self,
        event: Any = None,
        columns: list[int] | None = None,
        data_indexes: bool = False,
        undo: bool = True,
        emit_event: bool = True,
    ) -> EventDataDict:
        event_data = self.new_event_dict("delete_columns", state=True)
        if not columns and not (columns := sorted(self.get_selected_cols())):
            return event_data
        if not try_binding(self.extra_begin_del_cols_rc_func, event_data, "begin_delete_columns"):
            return
        if self.all_columns_displayed:
            data_columns = columns
            disp_columns = columns
        else:
            if data_indexes:
                data_columns = columns
                disp_columns = data_to_displayed_idxs(data_columns, self.displayed_columns)
            else:
                data_columns = [self.displayed_columns[c] for c in columns]
                disp_columns = columns
        event_data = self.delete_columns_displayed(
            disp_columns,
            event_data,
        )
        event_data = self.delete_columns_data(
            data_columns,
            event_data,
        )
        if undo and self.undo_enabled:
            self.undo_stack.append(stored_event_dict(event_data))
        try_binding(self.extra_end_del_cols_rc_func, event_data, "end_delete_columns")
        if emit_event:
            self.sheet_modified(event_data)
        self.deselect("all")
        return event_data

    def delete_rows_data(self, rows: list[int], event_data: EventDataDict | None = None) -> EventDataDict:
        self.mouseclick_outside_editor_or_dropdown_all_canvases()
        if not event_data:
            event_data = self.new_event_dict("delete_rows", state=True)
        event_data["options"] = self.copy_options()
        event_data["named_spans"] = {k: span.copy_self() for k, span in self.named_spans.items()}

        for i, datarn in enumerate(rows):
            event_data["deleted"]["rows"][datarn] = self.data.pop(datarn - i)

        if self.PAR.ops.treeview:
            event_data["deleted"]["index"] = {datarn: self._row_index[datarn] for datarn in rows}
            event_data = self.RI.tree_del_rows(event_data=event_data)
        elif isinstance(self._row_index, list):
            for i, datarn in enumerate(rows):
                r = datarn - i
                if len(self._row_index) > r:
                    event_data["deleted"]["index"][datarn] = self._row_index.pop(r)

        rows_set = set(rows)
        self.adjust_options_post_delete_rows(
            to_del=rows_set,
            to_bis=rows,
            named_spans=self.get_spans_to_del_from_rows(rows=rows_set),
        )
        if not self.all_rows_displayed:
            self.displayed_rows = [
                r if not (num := bisect_left(rows, r)) else r - num
                for r in filterfalse(rows_set.__contains__, self.displayed_rows)
            ]
        return event_data

    def delete_rows_displayed(
        self,
        rows: list[int],
        event_data: EventDataDict | None = None,
        from_undo: bool = False,
    ) -> EventDataDict:
        if not event_data:
            event_data = self.new_event_dict("delete_rows", state=True)
        self.saved_row_heights = {}
        if rows:
            for r in rows:
                if len(self.row_positions) > r + 1:
                    event_data["deleted"]["row_heights"][r] = self.row_positions[r + 1] - self.row_positions[r]
            if not from_undo:
                rows_set = set(rows)
                self.set_row_positions(
                    itr=(height for r, height in enumerate(self.gen_row_heights()) if r not in rows_set)
                )
        return event_data

    def delete_rows(
        self,
        event: Any = None,
        rows: list[int] | None = None,
        data_indexes: bool = False,
        undo: bool = True,
        emit_event: bool = True,
    ) -> None | EventDataDict:
        event_data = self.new_event_dict("delete_rows", state=True)
        if not rows and not (rows := sorted(self.get_selected_rows())):
            return
        if not try_binding(self.extra_begin_del_rows_rc_func, event_data, "begin_delete_rows"):
            return
        if self.all_rows_displayed:
            data_rows = rows
            disp_rows = rows
        else:
            if data_indexes:
                data_rows = rows
                # dont create disp_rows twice when using treeview mode
                disp_rows = [] if self.PAR.ops.treeview else data_to_displayed_idxs(data_rows, self.displayed_rows)
            else:
                data_rows = [self.displayed_rows[r] for r in rows]
                disp_rows = rows
        if self.PAR.ops.treeview:
            # remove any included descendants &
            # add all item descendants back in for safety,
            # update disp rows to del afterwards
            iids = {self._row_index[r].iid for r in data_rows}
            all_iids = set()
            for iid in iids:
                if not any(ancestor in iids for ancestor in self.RI.get_iid_ancestors(iid)):
                    all_iids.add(iid)
                    all_iids.update(self.RI.get_iid_descendants(iid))
            data_rows = sorted(map(self.RI.rns.__getitem__, all_iids))
            disp_rows = data_to_displayed_idxs(data_rows, self.displayed_rows)
        event_data = self.delete_rows_displayed(
            disp_rows,
            event_data,
        )
        event_data = self.delete_rows_data(
            data_rows,
            event_data,
        )
        if undo and self.undo_enabled:
            self.undo_stack.append(stored_event_dict(event_data))
        try_binding(self.extra_end_del_rows_rc_func, event_data, "end_delete_rows")
        if emit_event:
            self.sheet_modified(event_data)
        self.deselect("all")
        return event_data

    def move_row_position(self, idx1: int, idx2: int):
        if not len(self.row_positions) <= 2:
            if idx1 < idx2:
                height = self.row_positions[idx1 + 1] - self.row_positions[idx1]
                self.row_positions.insert(idx2 + 1, self.row_positions.pop(idx1 + 1))
                for i in range(idx1 + 1, idx2 + 1):
                    self.row_positions[i] -= height
                self.row_positions[idx2 + 1] = self.row_positions[idx2] + height
            else:
                height = self.row_positions[idx1 + 1] - self.row_positions[idx1]
                self.row_positions.insert(idx2 + 1, self.row_positions.pop(idx1 + 1))
                for i in range(idx2 + 2, idx1 + 2):
                    self.row_positions[i] += height
                self.row_positions[idx2 + 1] = self.row_positions[idx2] + height

    def move_col_position(self, idx1: int, idx2: int):
        if not len(self.col_positions) <= 2:
            if idx1 < idx2:
                width = self.col_positions[idx1 + 1] - self.col_positions[idx1]
                self.col_positions.insert(idx2 + 1, self.col_positions.pop(idx1 + 1))
                for i in range(idx1 + 1, idx2 + 1):
                    self.col_positions[i] -= width
                self.col_positions[idx2 + 1] = self.col_positions[idx2] + width
            else:
                width = self.col_positions[idx1 + 1] - self.col_positions[idx1]
                self.col_positions.insert(idx2 + 1, self.col_positions.pop(idx1 + 1))
                for i in range(idx2 + 2, idx1 + 2):
                    self.col_positions[i] += width
                self.col_positions[idx2 + 1] = self.col_positions[idx2] + width

    def display_rows(
        self,
        rows: int | Iterator[int] | None = None,
        all_rows_displayed: bool | None = None,
        reset_row_positions: bool = True,
        deselect_all: bool = True,
    ) -> list[int] | None:
        if rows is None and all_rows_displayed is None:
            return list(range(self.total_data_rows())) if self.all_rows_displayed else self.displayed_rows
        if rows is not None and rows != self.displayed_rows:
            self.purge_undo_and_redo_stack()
            self.displayed_rows = sorted(rows)
        # setting all_rows_displayed
        if all_rows_displayed is not None:
            # setting it to True and it's currently False
            if all_rows_displayed and not self.all_rows_displayed:
                self.purge_undo_and_redo_stack()
                self.all_rows_displayed = True
            # setting it to False and it's currently True
            elif not all_rows_displayed and self.all_rows_displayed:
                # if rows is None then displayed_rows needs to be reset
                if rows is None:
                    self.displayed_rows = list(range(self.total_data_rows()))
                self.all_rows_displayed = False
        if reset_row_positions:
            self.reset_row_positions()
        if deselect_all:
            self.deselect("all", redraw=False)

    def display_columns(
        self,
        columns: int | Iterator[int] | None = None,
        all_columns_displayed: bool | None = None,
        reset_col_positions: bool = True,
        deselect_all: bool = True,
    ) -> list[int] | None:
        if columns is None and all_columns_displayed is None:
            return list(range(self.total_data_cols())) if self.all_columns_displayed else self.displayed_columns
        if columns is not None and columns != self.displayed_columns:
            self.purge_undo_and_redo_stack()
            self.displayed_columns = sorted(columns)
        # setting all_columns_displayed
        if all_columns_displayed is not None:
            # setting it to True and it's currently False
            if all_columns_displayed and not self.all_columns_displayed:
                self.purge_undo_and_redo_stack()
                self.all_columns_displayed = True
            # setting it to False and it's currently True
            elif not all_columns_displayed and self.all_columns_displayed:
                # if columns is None then displayed_columns needs to be reset
                if columns is None:
                    self.displayed_columns = list(range(self.total_data_cols()))
                self.all_columns_displayed = False
        if reset_col_positions:
            self.reset_col_positions()
        if deselect_all:
            self.deselect("all", redraw=False)

    def headers(
        self,
        newheaders: Any = None,
        index: int | None = None,
        reset_col_positions: bool = False,
        show_headers_if_not_sheet: bool = True,
        redraw: bool = False,
    ) -> Any:
        if newheaders is not None:
            if isinstance(newheaders, (list, tuple)):
                self._headers = list(newheaders) if isinstance(newheaders, tuple) else newheaders
            elif isinstance(newheaders, int):
                self._headers = int(newheaders)
            elif isinstance(self._headers, list) and isinstance(index, int):
                if index >= len(self._headers):
                    self.CH.fix_header(index)
                self._headers[index] = f"{newheaders}"
            elif not isinstance(newheaders, (list, tuple, int)) and index is None:
                try:
                    self._headers = list(newheaders)
                except ValueError as error:
                    raise ValueError(
                        """
                        New header must be iterable or int \
                        (use int to use a row as the header
                        """
                    ) from error
            if reset_col_positions:
                self.reset_col_positions()
            elif (
                show_headers_if_not_sheet
                and isinstance(self._headers, list)
                and (self.col_positions == [0] or not self.col_positions)
            ):
                colpos = int(self.PAR.ops.default_column_width)
                if self.all_columns_displayed:
                    self.set_col_positions(itr=repeat(colpos, len(self._headers)))
                else:
                    self.set_col_positions(itr=repeat(colpos, len(self.displayed_columns)))
            if redraw:
                self.refresh()
        else:
            if not isinstance(self._headers, int) and index is not None and isinstance(index, int):
                return self._headers[index]
            else:
                return self._headers

    def row_index(
        self,
        newindex: Any = None,
        index: int | None = None,
        reset_row_positions: bool = False,
        show_index_if_not_sheet: bool = True,
        redraw: bool = False,
    ) -> Any:
        if newindex is not None:
            if not self._row_index and not isinstance(self._row_index, int):
                self.RI.set_width(self.PAR.ops.default_row_index_width, set_TL=True)
            if isinstance(newindex, (list, tuple)):
                self._row_index = list(newindex) if isinstance(newindex, tuple) else newindex
            elif isinstance(newindex, int):
                self._row_index = int(newindex)
            elif isinstance(index, int):
                if index >= len(self._row_index):
                    self.RI.fix_index(index)
                self._row_index[index] = f"{newindex}"
            elif not isinstance(newindex, (list, tuple, int)) and index is None:
                try:
                    self._row_index = list(newindex)
                except ValueError as error:
                    raise ValueError(
                        """
                        New index must be iterable or int \
                        (use int to use a column as the index
                        """
                    ) from error
            if reset_row_positions:
                self.reset_row_positions()
            elif (
                show_index_if_not_sheet
                and isinstance(self._row_index, list)
                and (self.row_positions == [0] or not self.row_positions)
            ):
                rowpos = self.get_default_row_height()
                if self.all_rows_displayed:
                    self.set_row_positions(itr=repeat(rowpos, len(self._row_index)))
                else:
                    self.set_row_positions(itr=repeat(rowpos, len(self.displayed_rows)))
            if redraw:
                self.refresh()
        else:
            if not isinstance(self._row_index, int) and index is not None and isinstance(index, int):
                return self._row_index[index]
            else:
                return self._row_index

    def total_data_cols(self, include_header: bool = True) -> int:
        h_total = len(self._headers) if include_header and isinstance(self._headers, (list, tuple)) else 0
        d_total = max(map(len, self.data), default=0)  # max(map(len, )) is faster
        return max(h_total, d_total)

    def total_data_rows(self, include_index: bool = True) -> int:
        i_total = len(self._row_index) if include_index and isinstance(self._row_index, (list, tuple)) else 0
        d_total = len(self.data)
        return max(i_total, d_total)

    def data_dimensions(self, total_rows: int | None = None, total_columns: int | None = None):
        if total_rows is None and total_columns is None:
            return self.total_data_rows(), self.total_data_cols()
        if total_rows is not None:
            if len(self.data) < total_rows:
                ncols = self.total_data_cols() if total_columns is None else total_columns
                self.data.extend(self.get_empty_row_seq(r, ncols) for r in range(total_rows - len(self.data)))
            else:
                self.data[total_rows:] = []
        if total_columns is not None:
            for rn, r in enumerate(self.data):
                if (lnr := len(r)) > total_columns:
                    r = r[:total_columns]
                elif lnr < total_columns:
                    r.extend(self.gen_empty_row_seq(rn, end=total_columns, start=lnr))

    def equalize_data_row_lengths(
        self,
        include_header: bool = True,
        total_data_cols: int | None = None,
        at_least_cols: int | None = None,
    ) -> int:
        if not isinstance(total_data_cols, int):
            total_data_cols = self.total_data_cols(include_header=include_header)
        if isinstance(at_least_cols, int) and at_least_cols > total_data_cols:
            total_data_cols = at_least_cols
        total_data_cols = max(total_data_cols, len(self.col_positions) - 1)
        if not isinstance(self._headers, int) and include_header and total_data_cols > len(self._headers):
            self.CH.fix_header(total_data_cols - 1)
        empty_v = self.get_value_for_empty_cell
        for rn, r in enumerate(self.data):
            if total_data_cols > (lnr := len(r)):
                r.extend(empty_v(rn, c, r_ops=True, c_ops=True) for c in range(lnr, total_data_cols))
        return total_data_cols

    def get_canvas_visible_area(self) -> tuple[float, float, float, float]:
        return (
            self.canvasx(0),
            self.canvasy(0),
            self.canvasx(self.winfo_width()),
            self.canvasy(self.winfo_height()),
        )

    @property
    def visible_text_rows(self) -> tuple[int, int]:
        start = bisect_left(self.row_positions, self.canvasy(0))
        end = bisect_right(self.row_positions, self.canvasy(self.winfo_height()))
        start = start - 1 if start else start
        end = end - 1 if end == len(self.row_positions) else end
        return start, end

    @property
    def visible_text_columns(self) -> tuple[int, int]:
        start = bisect_left(self.col_positions, self.canvasx(0))
        end = bisect_right(self.col_positions, self.canvasx(self.winfo_width()))
        start = start - 1 if start else start
        end = end - 1 if end == len(self.col_positions) else end
        return start, end

    def redraw_highlight_get_text_fg(
        self,
        r: int,
        c: int,
        fc: int | float,
        fr: int | float,
        sc: int | float,
        sr: int | float,
        sel_cells_bg: tuple[int, int, int],
        sel_cols_bg: tuple[int, int, int],
        sel_rows_bg: tuple[int, int, int],
        selections: dict,
        datarn: int,
        datacn: int,
        can_width: int | None,
        dont_blend: bool,
        alternate_color: Highlight | None,
        has_dd: bool,
        tags: str | tuple[str],
    ) -> tuple[str, bool]:
        if (datarn, datacn) in self.progress_bars:
            kwargs = self.progress_bars[(datarn, datacn)]
        else:
            if (datarn, datacn) in self.cell_options and "highlight" in self.cell_options[(datarn, datacn)]:
                kwargs = self.cell_options[(datarn, datacn)]["highlight"]
            elif datarn in self.row_options and "highlight" in self.row_options[datarn]:
                kwargs = self.row_options[datarn]["highlight"]
            elif datacn in self.col_options and "highlight" in self.col_options[datacn]:
                kwargs = self.col_options[datacn]["highlight"]
            else:
                kwargs = {}
        if alt := bool(not kwargs and alternate_color and r % 2):
            kwargs = alternate_color

        if kwargs:
            high_bg = kwargs[0]
            if high_bg and not high_bg.startswith("#"):
                high_bg = color_map[high_bg]

            # cell is a single currently selected cell box
            #         not highlighted
            #         on an alternate row color
            if alt and dont_blend:
                txtfg = self.PAR.ops.table_fg

            # cell is highlighted and cell selected
            elif "cells" in selections and (r, c) in selections["cells"]:
                txtfg = (
                    self.PAR.ops.table_selected_cells_fg
                    if kwargs[1] is None or self.PAR.ops.display_selected_fg_over_highlights
                    else kwargs[1]
                )
                self.redraw_highlight(
                    x1=fc + 1,
                    y1=fr + 1,
                    x2=sc,
                    y2=sr,
                    fill=(
                        self.PAR.ops.table_selected_cells_bg
                        if high_bg is None
                        else (
                            f"#{int((int(high_bg[1:3], 16) + sel_cells_bg[0]) / 2):02X}"
                            + f"{int((int(high_bg[3:5], 16) + sel_cells_bg[1]) / 2):02X}"
                            + f"{int((int(high_bg[5:], 16) + sel_cells_bg[2]) / 2):02X}"
                        )
                    ),
                    outline=self.PAR.ops.table_fg if has_dd and self.PAR.ops.show_dropdown_borders else "",
                    tags=tags,
                    can_width=None,
                    pc=None,
                )

            # cell is highlighted and row selected
            elif "rows" in selections and r in selections["rows"]:
                txtfg = (
                    self.PAR.ops.table_selected_rows_fg
                    if kwargs[1] is None or self.PAR.ops.display_selected_fg_over_highlights
                    else kwargs[1]
                )
                self.redraw_highlight(
                    x1=fc + 1,
                    y1=fr + 1,
                    x2=sc,
                    y2=sr,
                    fill=(
                        self.PAR.ops.table_selected_rows_bg
                        if high_bg is None
                        else (
                            f"#{int((int(high_bg[1:3], 16) + sel_rows_bg[0]) / 2):02X}"
                            + f"{int((int(high_bg[3:5], 16) + sel_rows_bg[1]) / 2):02X}"
                            + f"{int((int(high_bg[5:], 16) + sel_rows_bg[2]) / 2):02X}"
                        )
                    ),
                    outline=self.PAR.ops.table_fg if has_dd and self.PAR.ops.show_dropdown_borders else "",
                    tags=tags,
                    can_width=None,
                    pc=None,
                )

            # cell is highlighted and column selected
            elif "columns" in selections and c in selections["columns"]:
                txtfg = (
                    self.PAR.ops.table_selected_columns_fg
                    if kwargs[1] is None or self.PAR.ops.display_selected_fg_over_highlights
                    else kwargs[1]
                )
                self.redraw_highlight(
                    x1=fc + 1,
                    y1=fr + 1,
                    x2=sc,
                    y2=sr,
                    fill=(
                        self.PAR.ops.table_selected_columns_bg
                        if high_bg is None
                        else (
                            f"#{int((int(high_bg[1:3], 16) + sel_cols_bg[0]) / 2):02X}"
                            + f"{int((int(high_bg[3:5], 16) + sel_cols_bg[1]) / 2):02X}"
                            + f"{int((int(high_bg[5:], 16) + sel_cols_bg[2]) / 2):02X}"
                        )
                    ),
                    outline=self.PAR.ops.table_fg if has_dd and self.PAR.ops.show_dropdown_borders else "",
                    tags=tags,
                    can_width=None,
                    pc=None,
                )

            # cell is just highlighted (no selection)
            else:
                txtfg = self.PAR.ops.table_fg if kwargs[1] is None else kwargs[1]
                if high_bg:  # Only draw if fill exists
                    if not isinstance(kwargs, ProgressBar):
                        self.redraw_highlight(
                            x1=fc + 1,
                            y1=fr + 1,
                            x2=sc,
                            y2=sr,
                            fill=high_bg,
                            outline=self.PAR.ops.table_fg if has_dd and self.PAR.ops.show_dropdown_borders else "",
                            tags=tags,
                            can_width=can_width if (len(kwargs) > 2 and kwargs[2]) else None,
                            pc=None,
                        )
                    else:
                        if kwargs.del_when_done and kwargs.percent >= 100:
                            del self.progress_bars[(datarn, datacn)]
                        else:
                            self.redraw_highlight(
                                x1=fc + 1,
                                y1=fr + 1,
                                x2=sc,
                                y2=sr,
                                fill=high_bg,
                                outline=self.PAR.ops.table_fg if has_dd and self.PAR.ops.show_dropdown_borders else "",
                                tags=tags,
                                can_width=None,
                                pc=kwargs.percent,
                            )

        elif not kwargs:
            if "cells" in selections and (r, c) in selections["cells"]:
                txtfg = self.PAR.ops.table_selected_cells_fg
                self.redraw_highlight(
                    x1=fc + 1,
                    y1=fr + 1,
                    x2=sc,
                    y2=sr,
                    fill=self.PAR.ops.table_selected_cells_bg,
                    outline=self.PAR.ops.table_fg if has_dd and self.PAR.ops.show_dropdown_borders else "",
                    tags=tags,
                    can_width=None,
                    pc=None,
                )
            elif "rows" in selections and r in selections["rows"]:
                txtfg = self.PAR.ops.table_selected_rows_fg
                self.redraw_highlight(
                    x1=fc + 1,
                    y1=fr + 1,
                    x2=sc,
                    y2=sr,
                    fill=self.PAR.ops.table_selected_rows_bg,
                    outline=self.PAR.ops.table_fg if has_dd and self.PAR.ops.show_dropdown_borders else "",
                    tags=tags,
                    can_width=None,
                    pc=None,
                )
            elif "columns" in selections and c in selections["columns"]:
                txtfg = self.PAR.ops.table_selected_columns_fg
                self.redraw_highlight(
                    x1=fc + 1,
                    y1=fr + 1,
                    x2=sc,
                    y2=sr,
                    fill=self.PAR.ops.table_selected_columns_bg,
                    outline=self.PAR.ops.table_fg if has_dd and self.PAR.ops.show_dropdown_borders else "",
                    tags=tags,
                    can_width=None,
                    pc=None,
                )
            else:
                txtfg = self.PAR.ops.table_fg
                self.redraw_highlight(
                    x1=fc + 1,
                    y1=fr + 1,
                    x2=sc,
                    y2=sr,
                    fill="",
                    outline=self.PAR.ops.table_fg if has_dd and self.PAR.ops.show_dropdown_borders else "",
                    tags=tags,
                    can_width=None,
                    pc=None,
                )
        return txtfg

    def redraw_highlight(
        self,
        x1: int | float,
        y1: int | float,
        x2: int | float,
        y2: int | float,
        fill: str,
        outline: str,
        tags: str | tuple[str],
        can_width: None | float = None,
        pc: None | float = None,
    ) -> bool:
        if not self.PAR.ops.show_horizontal_grid:
            y2 += 1
        if not self.PAR.ops.show_vertical_grid:
            x2 += 1
        if not is_type_int(pc) or pc >= 100:
            coords = (
                x1 - 1 if outline else x1,
                y1 - 1 if outline else y1,
                x2 if can_width is None else x2 + can_width,
                y2,
            )
        elif pc <= 0:
            coords = (x1, y1, x1, y2)
        else:
            coords = (x1, y1, (x2 - x1) * (pc / 100), y2)
        if self.hidd_high:
            iid, showing = self.hidd_high.popitem()
            self.coords(iid, coords)
            if showing:
                self.itemconfig(iid, fill=fill, outline=outline, tags=tags)
            else:
                self.itemconfig(iid, fill=fill, outline=outline, state="normal", tags=tags)
        else:
            iid = self.create_rectangle(coords, fill=fill, outline=outline, tags=tags)
        self.disp_high[iid] = True
        return True

    def redraw_gridline(self, points: list[float]) -> None:
        if points:
            if self.hidd_grid:
                iid, sh = self.hidd_grid.popitem()
                self.coords(iid, points)
                if sh:
                    self.itemconfig(iid, fill=self.PAR.ops.table_grid_fg, width=1, capstyle="butt", joinstyle="round")
                else:
                    self.itemconfig(
                        iid,
                        fill=self.PAR.ops.table_grid_fg,
                        width=1,
                        capstyle="butt",
                        joinstyle="round",
                        state="normal",
                    )
            else:
                iid = self.create_line(
                    points, fill=self.PAR.ops.table_grid_fg, width=1, capstyle="butt", joinstyle="round"
                )
            self.disp_grid[iid] = True

    def _auto_resize_columns(self, can_width: float, col_pos_exists: bool) -> bool:
        if self.PAR.ops.auto_resize_columns and self.allow_auto_resize_columns and col_pos_exists:
            max_w = can_width - self.PAR.ops.empty_horizontal
            if self.PAR.ops.auto_resize_columns < self.PAR.ops.min_column_width:
                min_column_width = self.PAR.ops.min_column_width
            else:
                min_column_width = self.PAR.ops.auto_resize_columns
            if (len(self.col_positions) - 1) * min_column_width < max_w:
                change = int((max_w - self.col_positions[-1]) / (len(self.col_positions) - 1))
                widths = [
                    int(b - a) + change - 1
                    for a, b in zip(self.col_positions, islice(self.col_positions, 1, len(self.col_positions)))
                ]
                diffs = {}
                for i, w in enumerate(widths):
                    if w < min_column_width:
                        diffs[i] = min_column_width - w
                        widths[i] = min_column_width
                if diffs and len(diffs) < len(widths):
                    change = sum(diffs.values()) / (len(widths) - len(diffs))
                    for i in range(len(widths)):
                        if i not in diffs:
                            widths[i] -= change
                self.col_positions = list(accumulate(chain([0], widths)))
                return True
        return False

    def _auto_resize_rows(self, can_height: float, row_pos_exists: bool) -> bool:
        if self.PAR.ops.auto_resize_rows and self.allow_auto_resize_rows and row_pos_exists:
            max_h = can_height - self.PAR.ops.empty_vertical
            if self.PAR.ops.auto_resize_rows < self.min_row_height:
                min_row_height = self.min_row_height
            else:
                min_row_height = self.PAR.ops.auto_resize_rows
            if (len(self.row_positions) - 1) * min_row_height < max_h:
                change = int((max_h - self.row_positions[-1]) / (len(self.row_positions) - 1))
                heights = [
                    int(b - a) + change - 1
                    for a, b in zip(self.row_positions, islice(self.row_positions, 1, len(self.row_positions)))
                ]
                diffs = {}
                for i, h in enumerate(heights):
                    if h < min_row_height:
                        diffs[i] = min_row_height - h
                        heights[i] = min_row_height
                if diffs and len(diffs) < len(heights):
                    change = sum(diffs.values()) / (len(heights) - len(diffs))
                    for i in range(len(heights)):
                        if i not in diffs:
                            heights[i] -= change
                self.row_positions = list(accumulate(chain([0], heights)))
                return True
        return False

    def _manage_scroll_bars(self, can_height: float, can_width: float) -> None:
        if (
            self.PAR.ops.auto_resize_row_index is not True
            and can_width >= self.col_positions[-1] + self.PAR.ops.empty_horizontal
            and self.PAR.xscroll_showing
        ):
            self.PAR.xscroll.grid_remove()
            self.PAR.xscroll_showing = False
        elif (
            can_width < self.col_positions[-1] + self.PAR.ops.empty_horizontal
            and not self.PAR.xscroll_showing
            and not self.PAR.xscroll_disabled
            and can_height > 40
        ):
            self.PAR.xscroll.grid()
            self.PAR.xscroll_showing = True
        if can_height >= self.row_positions[-1] + self.PAR.ops.empty_vertical and self.PAR.yscroll_showing:
            self.PAR.yscroll.grid_remove()
            self.PAR.yscroll_showing = False
        elif (
            can_height < self.row_positions[-1] + self.PAR.ops.empty_vertical
            and not self.PAR.yscroll_showing
            and not self.PAR.yscroll_disabled
            and can_width > 40
        ):
            self.PAR.yscroll.grid()
            self.PAR.yscroll_showing = True

    def _overflow(
        self,
        cells: dict,
        rnge: Generator[int],
        datarn: int,
    ) -> Generator[float]:
        for c_ in rnge:
            if (
                cells[(datarn, cells["datacn"][c_])]
                or (datarn, cells["datacn"][c_]) in cells["dropdown"]
                or (datarn, cells["datacn"][c_]) in cells["checkbox"]
            ):
                return
            else:
                yield self.col_positions[c_ + 1] - self.col_positions[c_]

    def _redraw_precache_cells(
        self,
        text_start_row: int,
        text_end_row: int,
        text_start_col: int,
        text_end_col: int,
    ) -> dict:
        cells = {"datarn": {}, "datacn": {}, "dropdown": {}, "checkbox": {}}
        for r in range(text_start_row, text_end_row):
            datarn = r if self.all_rows_displayed else self.displayed_rows[r]
            cells["datarn"][r] = datarn
            for c in range(text_start_col, text_end_col):
                if c in cells["datacn"]:
                    datacn = cells["datacn"][c]
                else:
                    datacn = c if self.all_columns_displayed else self.displayed_columns[c]
                    cells["datacn"][c] = datacn
                t = (datarn, datacn)

                # self.get_cell_kwargs not used here to boost performance
                if t in self.cell_options and "dropdown" in self.cell_options[t]:
                    cells["dropdown"][t] = self.cell_options[t]["dropdown"]

                elif datarn in self.row_options and "dropdown" in self.row_options[datarn]:
                    cells["dropdown"][t] = self.row_options[datarn]["dropdown"]

                elif datacn in self.col_options and "dropdown" in self.col_options[datacn]:
                    cells["dropdown"][t] = self.col_options[datacn]["dropdown"]

                else:
                    if t in self.cell_options and "checkbox" in self.cell_options[t]:
                        cells["checkbox"][t] = self.cell_options[t]["checkbox"]

                    elif datarn in self.row_options and "checkbox" in self.row_options[datarn]:
                        cells["checkbox"][t] = self.row_options[datarn]["checkbox"]

                    elif datacn in self.col_options and "checkbox" in self.col_options[datacn]:
                        cells["checkbox"][t] = self.col_options[datacn]["checkbox"]

                cells[t] = self.cell_str(datarn, datacn, get_displayed=True)

        return cells

    def char_width_fn(self, c: str) -> int:
        if c in self.char_widths[self.table_font]:
            return self.char_widths[self.table_font][c]
        else:
            if c == "\t":
                self.txt_measure_canvas.itemconfig(
                    self.txt_measure_canvas_text,
                    text="\t",
                    font=self.table_font,
                )
                b = self.txt_measure_canvas.bbox(self.txt_measure_canvas_text)
                wd = b[2] - b[0] - 2
            else:
                self.txt_measure_canvas.itemconfig(
                    self.txt_measure_canvas_text,
                    text=_test_str + c,
                    font=self.table_font,
                )
                b = self.txt_measure_canvas.bbox(self.txt_measure_canvas_text)
                wd = b[2] - b[0] - self.table_test_str_w
            self.char_widths[self.table_font][c] = wd
            return wd

    def redraw_corner(self, x: float, y: float, tags: str | tuple[str]) -> None:
        if self.hidd_corners:
            iid = self.hidd_corners.pop()
            self.coords(iid, x - 10, y, x, y, x, y + 10)
            self.itemconfig(iid, fill=self.PAR.ops.table_grid_fg, state="normal", tags=tags)
            self.disp_corners.add(iid)
        else:
            self.disp_corners.add(
                self.create_polygon(x - 10, y, x, y, x, y + 10, fill=self.PAR.ops.table_grid_fg, tags=tags)
            )

    def redraw_grid_and_text(
        self,
        last_row_line_pos: float,
        last_col_line_pos: float,
        scrollpos_top: int,
        scrollpos_bot: int,
        scrollpos_left: int,
        scrollpos_right: int,
        x_stop: int | float,
        y_stop: int | float,
        col_pos_exists: bool,
        row_pos_exists: bool,
        can_width: int,
        can_height: int,
        grid_start_row: int,
        grid_end_row: int,
        grid_start_col: int,
        grid_end_col: int,
        text_start_row: int,
        text_end_row: int,
        text_start_col: int,
        text_end_col: int,
    ) -> None:
        # reset canvas item storage
        self.hidd_text.update(self.disp_text)
        self.disp_text = {}
        self.hidd_high.update(self.disp_high)
        self.disp_high = {}
        self.hidd_grid.update(self.disp_grid)
        self.disp_grid = {}
        self.hidd_dropdown.update(self.disp_dropdown)
        self.disp_dropdown = {}
        self.hidd_checkbox.update(self.disp_checkbox)
        self.disp_checkbox = {}
        self.hidd_corners.update(self.disp_corners)
        self.disp_corners = set()
        points = []
        # manage horizontal grid lines
        if self.PAR.ops.show_horizontal_grid and row_pos_exists:
            if self.PAR.ops.horizontal_grid_to_end_of_window:
                x_grid_stop = scrollpos_right + can_width
            else:
                x_grid_stop = x_stop + 1 if last_col_line_pos > scrollpos_right else x_stop - 1
            points.extend(
                get_horizontal_gridline_points(
                    left=scrollpos_left,
                    stop=x_grid_stop,
                    positions=self.row_positions,
                    start=grid_start_row,
                    end=grid_end_row,
                )
            )
        # manage vertical grid lines
        if self.PAR.ops.show_vertical_grid and col_pos_exists:
            if self.PAR.ops.vertical_grid_to_end_of_window:
                y_grid_stop = scrollpos_bot + can_height
            else:
                y_grid_stop = y_stop + 1 if last_row_line_pos > scrollpos_bot else y_stop - 1
            if self.PAR.ops.show_horizontal_grid:
                points.extend([scrollpos_left - 2, grid_end_row, grid_start_col - 2, scrollpos_top - 2])
            points.extend(
                get_vertical_gridline_points(
                    top=scrollpos_top,
                    stop=y_grid_stop,
                    positions=self.col_positions,
                    start=grid_start_col,
                    end=grid_end_col,
                )
            )
        if points:
            self.redraw_gridline(points)

        font = self.PAR.ops.table_font
        dd_coords = self.dropdown.get_coords()
        selections = self.get_redraw_selections(text_start_row, grid_end_row, text_start_col, grid_end_col)
        sel_cells_bg = color_tup(self.PAR.ops.table_selected_cells_bg)
        sel_cols_bg = color_tup(self.PAR.ops.table_selected_columns_bg)
        sel_rows_bg = color_tup(self.PAR.ops.table_selected_rows_bg)
        current_loc = (self.selected.row, self.selected.column) if self.selected else ()
        if self.PAR.ops.alternate_color:
            alternate_color = Highlight(
                bg=self.PAR.ops.alternate_color,
                fg=None,
                end=False,
            )
            if self.selected and box_is_single_cell(*self.selected.box) and self.PAR.ops.show_selected_cells_border:
                dont_blend = current_loc
            else:
                dont_blend = ()
        else:
            alternate_color = None
            dont_blend = ()
        if self.PAR.ops.show_selected_cells_border:
            override = ()
        else:
            override = (
                color_tup(self.PAR.ops.table_selected_cells_fg),
                color_tup(self.PAR.ops.table_selected_columns_fg),
                color_tup(self.PAR.ops.table_selected_rows_fg),
            )
        allow_overflow = self.PAR.ops.allow_cell_overflow
        wrap = self.PAR.ops.table_wrap
        cells = self._redraw_precache_cells(
            text_start_row=text_start_row,
            text_end_row=text_end_row,
            text_start_col=text_start_col,
            text_end_col=text_end_col,
        )
        note_corners = self.PAR.ops.note_corners

        # This is a little messy but
        # we try to avoid any function use to maximise performance
        for r in range(text_start_row, text_end_row):
            rtopgridln = self.row_positions[r]
            rbotgridln = self.row_positions[r + 1]
            datarn = cells["datarn"][r]
            for c in range(text_start_col, text_end_col):
                cleftgridln = self.col_positions[c]
                crightgridln = self.col_positions[c + 1]
                datacn = cells["datacn"][c]
                disp_loc = (r, c)
                loc = (datarn, datacn)
                tag = f"{r}_{c}"
                fill = self.redraw_highlight_get_text_fg(
                    r=r,
                    c=c,
                    fc=cleftgridln,
                    fr=rtopgridln,
                    sc=crightgridln,
                    sr=rbotgridln,
                    sel_cells_bg=override[0] if override and disp_loc == current_loc else sel_cells_bg,
                    sel_cols_bg=override[1] if override and disp_loc == current_loc else sel_cols_bg,
                    sel_rows_bg=override[2] if override and disp_loc == current_loc else sel_rows_bg,
                    selections=selections,
                    datarn=datarn,
                    datacn=datacn,
                    can_width=can_width,
                    dont_blend=disp_loc == dont_blend,
                    alternate_color=alternate_color,
                    has_dd=loc in cells["dropdown"],
                    tags=("h", "c", tag),
                )
                if loc in self.cell_options and "align" in self.cell_options[(datarn, datacn)]:
                    align = self.cell_options[(datarn, datacn)]["align"]
                elif datarn in self.row_options and "align" in self.row_options[datarn]:
                    align = self.row_options[datarn]["align"]
                elif datacn in self.col_options and "align" in self.col_options[datacn]:
                    align = self.col_options[datacn]["align"]
                else:
                    align = self.align

                kws = cells["dropdown"][loc] if loc in cells["dropdown"] else None
                if kws:
                    max_width = crightgridln - cleftgridln - self.table_txt_height - 5
                    if align[-1] == "w":
                        draw_x = cleftgridln + 2
                    elif align[-1] == "e":
                        draw_x = crightgridln - 5 - self.table_txt_height
                    elif align[-1] == "n":
                        draw_x = cleftgridln + (crightgridln - cleftgridln - self.table_txt_height) / 2

                    # redraw dropdown

                    x1 = cleftgridln
                    y1 = rtopgridln
                    x2 = crightgridln
                    y2 = self.row_positions[r + 1]
                    # if not dd_drawn and self.PAR.ops.show_dropdown_borders:
                    #     self.redraw_highlight(x1 + 1, y1 + 1, x2, y2, fill="", outline=self.PAR.ops.table_fg)
                    if max_width >= 5:
                        if dd_coords == disp_loc:
                            # up arrow
                            points = (
                                x2 - self.dd_up_arrow[0],
                                y1 + self.dd_up_arrow[1],
                                x2 - self.dd_up_arrow[2],
                                y1 + self.dd_up_arrow[3],
                                x2 - self.dd_up_arrow[4],
                                y1 + self.dd_up_arrow[5],
                            )
                        else:
                            # down arrow
                            points = (
                                x2 - self.dd_down_arrow[0],
                                y1 + self.dd_down_arrow[1],
                                x2 - self.dd_down_arrow[2],
                                y1 + self.dd_down_arrow[3],
                                x2 - self.dd_down_arrow[4],
                                y1 + self.dd_down_arrow[5],
                            )
                        _fill = fill if kws["state"] != "disabled" else self.PAR.ops.table_grid_fg
                        if self.hidd_dropdown:
                            cid, sh = self.hidd_dropdown.popitem()
                            self.coords(cid, points)
                            if sh:
                                self.itemconfig(cid, fill=_fill)
                            else:
                                self.itemconfig(cid, fill=_fill, state="normal")
                        else:
                            cid = self.create_line(
                                points, fill=_fill, width=2, capstyle="round", joinstyle="bevel", tag="lift"
                            )
                        self.disp_dropdown[cid] = True

                elif loc in cells["checkbox"]:
                    kws = cells["checkbox"][loc]

                    if align[-1] == "w":
                        draw_x = cleftgridln + 2
                    elif align[-1] == "e":
                        draw_x = crightgridln - 2
                    elif align[-1] == "n":
                        draw_x = cleftgridln + (crightgridln - cleftgridln) / 2

                    max_width = crightgridln - cleftgridln - 2

                    if max_width > self.table_txt_height + 1:
                        box_w = self.table_txt_height + 1
                        if align[-1] == "w":
                            draw_x += box_w + 3
                        elif align[-1] == "n":
                            draw_x += box_w / 2 + 1
                        max_width -= box_w + 4
                        try:
                            draw_check = bool(self.data[datarn][datacn])
                        except Exception:
                            draw_check = False

                        # redraw checkbox

                        x1 = cleftgridln + 2
                        y1 = rtopgridln + 2
                        x2 = cleftgridln + self.table_txt_height + 3
                        y2 = rtopgridln + self.table_txt_height + 3
                        points = rounded_box_coords(x1, y1, x2, y2)
                        _fill = fill if kws["state"] == "normal" else self.PAR.ops.table_grid_fg
                        if self.hidd_checkbox:
                            cid, sh = self.hidd_checkbox.popitem()
                            self.coords(cid, points)
                            if sh:
                                self.itemconfig(cid, fill="", outline=_fill)
                            else:
                                self.itemconfig(cid, fill="", outline=_fill, state="normal")
                        else:
                            cid = self.create_polygon(points, fill="", outline=_fill, smooth=True, tag="lift")
                        self.disp_checkbox[cid] = True
                        if draw_check:
                            points = rounded_box_coords(x1 + 4, y1 + 4, x2 - 3, y2 - 3, radius=4)
                            if self.hidd_checkbox:
                                cid, sh = self.hidd_checkbox.popitem()
                                self.coords(cid, points)
                                if sh:
                                    self.itemconfig(cid, fill=_fill, outline="")
                                else:
                                    self.itemconfig(cid, fill=_fill, outline="", state="normal")
                            else:
                                cid = self.create_polygon(points, fill=_fill, outline="", smooth=True, tag="lift")
                            self.disp_checkbox[cid] = True

                else:
                    max_width = crightgridln - cleftgridln - 2
                    if align[-1] == "w":
                        draw_x = cleftgridln + 2
                    elif align[-1] == "e":
                        draw_x = crightgridln - 2
                    elif align[-1] == "n":
                        draw_x = cleftgridln + (crightgridln - cleftgridln) / 2

                tags = ("lift", "c", tag)

                if (
                    note_corners
                    and max_width > 5
                    and (
                        (loc in self.cell_options and "note" in self.cell_options[(datarn, datacn)])
                        or (datarn in self.row_options and "note" in self.row_options[datarn])
                        or (datacn in self.col_options and "note" in self.col_options[datacn])
                    )
                ):
                    self.redraw_corner(crightgridln, rtopgridln, tags)

                # redraw text
                if (
                    not cells[loc]
                    or (align[-1] == "w" and draw_x > scrollpos_right)
                    or cleftgridln + 5 > scrollpos_right
                ):
                    continue
                if allow_overflow and not kws:
                    if align[-1] == "w":
                        max_width += sum(self._overflow(cells, range(c + 1, text_end_col), datarn))
                    elif align[-1] == "e":
                        max_width += sum(self._overflow(cells, reversed(range(text_start_col, c)), datarn))
                if max_width <= 1:
                    continue
                start_line = max(0, int((scrollpos_top - rtopgridln) / self.table_txt_height))
                draw_y = rtopgridln + 3 + (start_line * self.table_txt_height)
                gen_lines = wrap_text(
                    text=cells[loc],
                    max_width=max_width,
                    max_lines=int((rbotgridln - rtopgridln - 2) / self.table_txt_height),
                    char_width_fn=self.char_width_fn,
                    widths=self.char_widths[font],
                    wrap=wrap,
                    start_line=start_line,
                )
                if align[-1] == "w" or align[-1] == "e":
                    if self.hidd_text:
                        iid, showing = self.hidd_text.popitem()
                        self.coords(iid, draw_x, draw_y)
                        if showing:
                            self.itemconfig(
                                iid,
                                text="\n".join(gen_lines),
                                fill=fill,
                                font=font,
                                anchor=align,
                                tags=tags,
                            )
                        else:
                            self.itemconfig(
                                iid,
                                text="\n".join(gen_lines),
                                fill=fill,
                                font=font,
                                anchor=align,
                                state="normal",
                                tags=tags,
                            )
                    else:
                        iid = self.create_text(
                            draw_x,
                            draw_y,
                            text="\n".join(gen_lines),
                            fill=fill,
                            font=font,
                            anchor=align,
                            tags=tags,
                        )
                    self.disp_text[iid] = True

                elif align[-1] == "n":
                    for t in gen_lines:
                        if self.hidd_text:
                            iid, showing = self.hidd_text.popitem()
                            self.coords(iid, draw_x, draw_y)
                            if showing:
                                self.itemconfig(
                                    iid,
                                    text=t,
                                    fill=fill,
                                    font=font,
                                    anchor=align,
                                    tags=tags,
                                )
                            else:
                                self.itemconfig(
                                    iid,
                                    text=t,
                                    fill=fill,
                                    font=font,
                                    anchor=align,
                                    state="normal",
                                    tags=tags,
                                )
                        else:
                            iid = self.create_text(
                                draw_x,
                                draw_y,
                                text=t,
                                fill=fill,
                                font=font,
                                anchor=align,
                                tags=tags,
                            )
                        self.disp_text[iid] = True
                        draw_y += self.table_txt_height
        for dct in (
            self.hidd_text,
            self.hidd_high,
            self.hidd_grid,
            self.hidd_dropdown,
            self.hidd_checkbox,
        ):
            for iid, showing in dct.items():
                if showing:
                    self.itemconfig(iid, state="hidden")
                    dct[iid] = False
        for iid in self.hidd_corners:
            self.itemconfig(iid, state="hidden")
        if self.PAR.ops.show_selected_cells_border:
            for _, box in self.selection_boxes.items():
                if box.bd_iid:
                    self.tag_raise(box.bd_iid)
            if self.selected:
                self.tag_raise(self.selected.iid)
        self.lift("lift")
        self.tag_bind("c", "<Enter>", self.enter_cell)
        self.tag_bind("c", "<Leave>", self.leave_cell)

    def enter_cell(self, event: tk.Event | None = None) -> None:
        if any_editor_or_dropdown_open(self):
            return
        can_x, can_y = self.canvasx(event.x), self.canvasy(event.y)
        for i in self.find_overlapping(can_x - 1, can_y - 1, can_x + 1, can_y + 1):
            try:
                if (coords := self.gettags(i)[2]) == self.tooltip_coords:
                    return
                self.tooltip_coords = coords
                self.tooltip_last_x, self.tooltip_last_y = self.winfo_pointerx(), self.winfo_pointery()
                self.start_tooltip_timer()
                return
            except Exception:
                continue

    def leave_cell(self, event: tk.Event | None = None) -> None:
        if self.tooltip_after_id is not None:
            self.after_cancel(self.tooltip_after_id)
            self.tooltip_after_id = None
        if self.tooltip_showing:
            if self.winfo_containing(self.winfo_pointerx(), self.winfo_pointery()) not in self.tooltip_widgets:
                self.close_tooltip_save()
        else:
            self.tooltip_coords = None

    def start_tooltip_timer(self) -> None:
        self.tooltip_after_id = self.after(self.PAR.ops.tooltip_hover_delay, self.check_and_show_tooltip)

    def check_and_show_tooltip(self, event: tk.Event | None = None) -> None:
        current_x, current_y = self.winfo_pointerx(), self.winfo_pointery()
        if current_x < 0 or current_y < 0:
            return
        if abs(current_x - self.tooltip_last_x) <= 1 and abs(current_y - self.tooltip_last_y) <= 1:
            self.show_tooltip()
        else:
            self.tooltip_last_x, self.tooltip_last_y = current_x, current_y
            self.tooltip_after_id = self.after(400, self.check_and_show_tooltip)

    def hide_tooltip(self) -> None:
        self.tooltip.withdraw()
        self.tooltip_showing, self.tooltip_coords = False, None

    def show_tooltip(self) -> None:
        if self.text_editor.open or self.dropdown.open:
            return
        coords = self.tooltip_coords.split("_")
        r, c = int(coords[0]), int(coords[1])
        datarn, datacn = self.datarn(r), self.datacn(c)
        kws = self.get_cell_kwargs(datarn, datacn, key="note")
        if not self.PAR.ops.tooltips and not kws and not self.PAR.ops.user_can_create_notes:
            return
        self.CH.hide_tooltip()
        self.RI.hide_tooltip()
        cell_readonly = self.get_cell_kwargs(datarn, datacn, "readonly") or not self.table_edit_cell_enabled()
        if kws:
            note = kws["note"]
            note_readonly = kws["readonly"]
        elif self.PAR.ops.user_can_create_notes:
            note = ""
            note_readonly = False
        else:
            note = None
            note_readonly = True
        note_only = not self.PAR.ops.tooltips and isinstance(note, str)
        self.tooltip_cell_content = f"{self.get_cell_data(datarn, datacn, none_to_empty_str=True)}"
        self.tooltip.reset(
            **{
                "text": self.tooltip_cell_content,
                "cell_readonly": cell_readonly,
                "note": note,
                "note_readonly": note_readonly,
                "row": r,
                "col": c,
                "menu_kwargs": get_menu_kwargs(self.PAR.ops),
                **get_bg_fg(self.PAR.ops),
                "user_can_create_notes": self.PAR.ops.user_can_create_notes,
                "note_only": note_only,
                "width": self.PAR.ops.tooltip_width,
                "height": self.PAR.ops.tooltip_height,
            }
        )
        self.tooltip.set_position(self.tooltip_last_x - 4, self.tooltip_last_y - 4)
        self.tooltip_showing = True

    def close_tooltip_save(self, event: tk.Event | None = None) -> None:
        widget = self.winfo_containing(self.winfo_pointerx(), self.winfo_pointery())
        if any(widget == tw for tw in self.tooltip_widgets):
            return
        if not self.tooltip.cell_readonly or not self.tooltip.note_readonly:
            r, c, cell, note = self.tooltip.get()
            datarn, datacn = self.datarn(r), self.datacn(c)
            if not self.tooltip.cell_readonly and cell != self.tooltip_cell_content:
                event_data = self.new_single_edit_event(
                    r, c, datarn, datacn, "??", self.get_cell_data(datarn, datacn), cell
                )
                value, event_data = self.single_edit_run_validation(datarn, datacn, event_data)
                if value is not None:
                    self.set_cell_data_undo(r=r, c=c, datarn=datarn, datacn=datacn, value=value, redraw=False)
                    try_binding(self.extra_end_edit_cell_func, event_data)
            if not self.tooltip.note_readonly:
                self.PAR.note(datarn, datacn, note=note if note else None, readonly=False)
            self.refresh()
        self.hide_tooltip()
        self.focus_set()

    def main_table_redraw_grid_and_text(
        self,
        redraw_header: bool = False,
        redraw_row_index: bool = False,
        redraw_table: bool = True,
        setting_views: bool = False,
        set_scrollregion: bool = True,
    ) -> bool:
        try:
            can_width = self.winfo_width()
            can_height = self.winfo_height()
        except Exception:
            return False
        row_pos_exists = self.row_positions != [0] and self.row_positions
        col_pos_exists = self.col_positions != [0] and self.col_positions

        resized_cols = self._auto_resize_columns(can_width=can_width, col_pos_exists=col_pos_exists)
        resized_rows = self._auto_resize_rows(can_height=can_height, row_pos_exists=row_pos_exists)
        self._manage_scroll_bars(can_height=can_height, can_width=can_width)

        last_col_line_pos = self.col_positions[-1] + 1
        last_row_line_pos = self.row_positions[-1] + 1
        scrollregion = (
            0,
            0,
            last_col_line_pos + self.PAR.ops.empty_horizontal + 2,
            last_row_line_pos + self.PAR.ops.empty_vertical + 2,
        )
        if set_scrollregion and (setting_views or scrollregion != self.scrollregion):
            try:
                self.configure(scrollregion=scrollregion)
            except Exception:
                return False
            self.scrollregion = scrollregion
            if (
                not self.CH.configure_scrollregion(last_col_line_pos)
                or not self.RI.configure_scrollregion(last_row_line_pos)
                or setting_views
            ):
                return False
        scrollpos_top = self.canvasy(0)
        scrollpos_bot = self.canvasy(can_height)
        scrollpos_left = self.canvasx(0)
        scrollpos_right = self.canvasx(can_width)
        grid_start_row = bisect_left(self.row_positions, scrollpos_top)
        grid_end_row = bisect_right(self.row_positions, scrollpos_bot)
        grid_start_col = bisect_left(self.col_positions, scrollpos_left)
        grid_end_col = bisect_right(self.col_positions, scrollpos_right)
        text_start_row = grid_start_row - 1 if grid_start_row else grid_start_row
        text_end_row = grid_end_row - 1 if grid_end_row == len(self.row_positions) else grid_end_row
        text_start_col = grid_start_col - 1 if grid_start_col else grid_start_col
        text_end_col = grid_end_col - 1 if grid_end_col == len(self.col_positions) else grid_end_col

        if self.find_window.open:
            self.reposition_find_window()
        # check if auto resizing row index
        changed_w = False
        if self.PAR.ops.auto_resize_row_index and self.show_index:
            changed_w = self.RI.auto_set_index_width(
                end_row=grid_end_row,
                only_rows=map(self.datarn, range(text_start_row, text_end_row)),
            )
        if resized_cols or resized_rows or changed_w:
            self.recreate_all_selection_boxes()
        x_stop = min(last_col_line_pos, scrollpos_right)
        y_stop = min(last_row_line_pos, scrollpos_bot)
        if redraw_header and self.show_header:
            self.CH.redraw_grid_and_text(
                last_col_line_pos=last_col_line_pos,
                scrollpos_left=scrollpos_left,
                x_stop=x_stop,
                grid_start_col=grid_start_col,
                grid_end_col=grid_end_col,
                text_start_col=text_start_col,
                text_end_col=text_end_col,
                scrollpos_right=scrollpos_right,
                col_pos_exists=col_pos_exists,
                set_scrollregion=set_scrollregion,
            )
        if redraw_row_index and self.show_index:
            self.RI.redraw_grid_and_text(
                last_row_line_pos=last_row_line_pos,
                scrollpos_top=scrollpos_top,
                y_stop=y_stop,
                grid_start_row=grid_start_row,
                grid_end_row=grid_end_row,
                text_start_row=text_start_row,
                text_end_row=text_end_row,
                scrollpos_bot=scrollpos_bot,
                row_pos_exists=row_pos_exists,
                set_scrollregion=set_scrollregion,
            )
        if redraw_table:
            self.redraw_grid_and_text(
                last_row_line_pos=last_row_line_pos,
                last_col_line_pos=last_col_line_pos,
                scrollpos_top=scrollpos_top,
                scrollpos_bot=scrollpos_bot,
                scrollpos_left=scrollpos_left,
                scrollpos_right=scrollpos_right,
                x_stop=x_stop,
                y_stop=y_stop,
                col_pos_exists=col_pos_exists,
                row_pos_exists=row_pos_exists,
                can_width=can_width,
                can_height=can_height,
                grid_start_row=grid_start_row,
                grid_end_row=grid_end_row,
                grid_start_col=grid_start_col,
                grid_end_col=grid_end_col,
                text_start_row=text_start_row,
                text_end_row=text_end_row,
                text_start_col=text_start_col,
                text_end_col=text_end_col,
            )
        event_data = {"sheetname": "", "header": redraw_header, "row_index": redraw_row_index, "table": redraw_table}
        self.PAR.emit_event("<<SheetRedrawn>>", data=event_data)
        return True

    def get_selection_items(
        self,
        cells: bool = True,
        rows: bool = True,
        columns: bool = True,
        reverse: bool = False,
    ) -> tuple[tuple[int, SelectionBox]]:
        """
        Most recent selection box should be last
        """
        itr = reversed(self.selection_boxes.items()) if reverse else self.selection_boxes.items()
        return tuple(
            (iid, box)
            for iid, box in itr
            if cells and box.type_ == "cells" or rows and box.type_ == "rows" or columns and box.type_ == "columns"
        )

    def get_boxes(self) -> dict[Box_nt, Literal["cells", "rows", "columns"]]:
        return {box.coords: box.type_ for box in self.selection_boxes.values()}

    def reselect_from_get_boxes(
        self,
        boxes: dict,
        selected: tuple = (),
    ) -> None:
        for (r1, c1, r2, c2), v in boxes.items():
            if r2 < len(self.row_positions) and c2 < len(self.col_positions):
                self.create_selection_box(r1, c1, r2, c2, v, run_binding=True)
        if selected:
            self.set_currently_selected(selected.row, selected.column, box=selected.box)

    def set_currently_selected(
        self,
        r: int | None = None,
        c: int | None = None,
        item: int | None = None,
        box: tuple[int, int, int, int] | None = None,
        run_binding: bool = True,
    ) -> None:
        def box_created(r: int, c: int, box: SelectionBox) -> bool:
            r1, c1, r2, c2 = box.coords
            if r1 <= r and c1 <= c and r2 >= r and c2 >= c:
                self.create_currently_selected_box(r, c, box.type_, box.fill_iid)
                if run_binding:
                    self.run_selection_binding(box.type_)
                return True
            return False

        # set current to a particular existing selection box
        if isinstance(item, int) and item in self.selection_boxes:
            selection_box = self.selection_boxes[item]
            r1, c1, _, _ = selection_box.coords
            if box_created(r1 if r is None else r, c1 if c is None else c, selection_box):
                return

        # set current to any existing selection box with coordinates: box
        if isinstance(box, tuple):
            for _, selection_box in self.get_selection_items(reverse=True):
                if box == selection_box.coords and box_created(
                    box[0] if r is None else r, box[1] if c is None else c, selection_box
                ):
                    return

        # set current to a coordinate, find the top most box there
        if isinstance(r, int) and isinstance(c, int):
            for _, selection_box in self.get_selection_items(reverse=True):
                if box_created(r, c, selection_box):
                    return

            # wasn't provided an item and couldn't find a box at coords so select cell
            if r < len(self.row_positions) - 1 and c < len(self.col_positions) - 1:
                self.select_cell(r, c, redraw=True)

    def set_current_to_last(self) -> None:
        if self.selection_boxes:
            box = next(iter(reversed(self.selection_boxes.values())))
            r1, c1, r2, c2 = box.coords
            if r2 - r1 == 1 and c2 - c1 == 1:
                box.state = "hidden"
            self.set_currently_selected(item=box.fill_iid)

    def coords_and_type(self, item: int) -> tuple:
        if item in self.selection_boxes:
            return Box_t(*(self.selection_boxes[item].coords + (self.selection_boxes[item].type_,)))
        return ()

    def get_selected_box_bg_fg(self, type_: str) -> tuple:
        type_map = {
            "cells": (self.PAR.ops.table_selected_cells_bg, self.PAR.ops.table_selected_box_cells_fg),
            "rows": (self.PAR.ops.table_selected_rows_bg, self.PAR.ops.table_selected_box_rows_fg),
            "columns": (self.PAR.ops.table_selected_columns_bg, self.PAR.ops.table_selected_box_columns_fg),
        }
        return type_map[type_]

    def create_currently_selected_box(
        self,
        r: int,
        c: int,
        type_: Literal["cells", "rows", "columns"],
        fill_iid: int,
    ) -> int:
        _, outline = self.get_selected_box_bg_fg(type_=type_)
        x1 = self.col_positions[c] + 1
        y1 = self.row_positions[r] + 1
        x2 = self.col_positions[c + 1] if index_exists(self.col_positions, c + 1) else self.col_positions[c] + 1
        y2 = self.row_positions[r + 1] if index_exists(self.row_positions, r + 1) else self.row_positions[r] + 1
        self.hide_selected()
        iid = self.display_border(
            x1,
            y1,
            x2,
            y2,
            fill=outline if self.PAR.ops.show_selected_cells_border else "",
            state="normal",
            tags=("selected", "lift"),
            width=2,
        )
        self.selected = Selected(
            row=r,
            column=c,
            type_=type_,
            box=self.selection_boxes[fill_iid].coords,
            iid=iid,
            fill_iid=fill_iid,
        )
        return iid

    def display_border(
        self,
        x1: int,
        y1: int,
        x2: int,
        y2: int,
        fill: str,
        state: str,
        tags: str | tuple[str],
        width: int,
        iid: None | int = None,
    ) -> int:
        radius = 0 if not self.PAR.ops.rounded_boxes or not x2 - x1 or not y2 - y1 else 5
        coords = rounded_box_coords(
            x1,
            y1,
            x2,
            y2,
            radius=radius,
        )
        if isinstance(iid, int):
            self.itemconfig(iid, fill=fill, state=state, tags=tags, width=width)
            self.coords(iid, *coords)
        else:
            if self.hidd_borders:
                iid = self.hidd_borders.pop()
                self.itemconfig(iid, fill=fill, state=state, tags=tags, width=width)
                self.coords(iid, *coords)
            else:
                iid = self.create_line(
                    *coords,
                    fill=fill,
                    state=state,
                    tags=tags,
                    width=width,
                    smooth=True,
                )
            self.disp_borders.add(iid)
        return iid

    def hide_border(self, item: int | None) -> None:
        if isinstance(item, int):
            self.disp_borders.discard(item)
            self.hidd_borders.add(item)
            self.itemconfig(item, state="hidden")

    def hide_box_fill(self, item: int | None) -> None:
        if isinstance(item, int):
            self.disp_selection_fills.discard(item)
            self.hidd_selection_fills.add(item)

    def hide_selection_box(self, item: int | None) -> bool:
        if item is None or item is True or item not in self.selection_boxes:
            return False
        box = self.selection_boxes.pop(item)
        self.hide_box_fill(box.fill_iid)
        self.hide_border(box.bd_iid)
        if self.selected.fill_iid == item:
            self.hide_selected()
            self.set_current_to_last()
        if item == self.being_drawn_item:
            self.being_drawn_item = None
        elif item == self.RI.being_drawn_item:
            self.RI.being_drawn_item = None
        elif item == self.CH.being_drawn_item:
            self.CH.being_drawn_item = None
        return True

    def hide_selected(self) -> None:
        if self.selected:
            self.hide_border(self.selected.iid)
            self.selected = ()

    def get_selection_fill(self) -> int:
        if self.hidd_selection_fills:
            iid = self.hidd_selection_fills.pop()
        else:
            self.selection_box_ctr += 1
            iid = self.selection_box_ctr
        self.disp_selection_fills.add(iid)
        return iid

    def create_selection_box(
        self,
        r1: int,
        c1: int,
        r2: int,
        c2: int,
        type_: str = "cells",
        state: str = "normal",
        set_current: bool | tuple[int, int] = True,
        run_binding: bool = False,
        ext: bool = False,
    ) -> int:
        if self.col_positions == [0]:
            c1 = 0
            c2 = 0
        if self.row_positions == [0]:
            r1 = 0
            r2 = 0
        if type_ == "cells":
            mt_border_col = self.PAR.ops.table_selected_cells_border_fg
        elif type_ == "rows":
            mt_border_col = self.PAR.ops.table_selected_rows_border_fg
        elif type_ == "columns":
            mt_border_col = self.PAR.ops.table_selected_columns_border_fg
        if self.selection_boxes:
            next(reversed(self.selection_boxes.values())).state = "normal"
        x1, y1, x2, y2 = self.box_coords_x_canvas_coords(r1, c1, r2, c2)
        fill_iid = self.get_selection_fill()
        bd_iid = None
        if self.PAR.ops.show_selected_cells_border and (
            ext
            or self.ctrl_b1_pressed
            or (self.being_drawn_item is None and self.RI.being_drawn_item is None and self.CH.being_drawn_item is None)
        ):
            bd_iid = self.display_border(
                x1,
                y1,
                self.canvasx(self.winfo_width()) if self.PAR.name == "!SheetDropdown" else x2,
                y2,
                fill=mt_border_col,
                state="normal",
                tags=f"{type_}bd",
                width=1,
            )
            self.tag_raise(bd_iid)
        self.selection_boxes[fill_iid] = SelectionBox(
            fill_iid=fill_iid,
            bd_iid=bd_iid,
            index=fill_iid,
            header=fill_iid,
            coords=Box_nt(r1, c1, r2, c2),
            type_=type_,
            state=state,
        )
        if set_current:
            if set_current is True:
                curr_r = r1
                curr_c = c1
            elif isinstance(set_current, tuple):
                curr_r = set_current[0]
                curr_c = set_current[1]
            self.create_currently_selected_box(curr_r, curr_c, type_, fill_iid)
        if run_binding:
            self.run_selection_binding(type_)
        return fill_iid

    def box_coords_x_canvas_coords(self, r1: int, c1: int, r2: int, c2: int) -> tuple[float, float, float, float]:
        return self.col_positions[c1], self.row_positions[r1], self.col_positions[c2], self.row_positions[r2]

    def recreate_selection_box(
        self,
        r1: int,
        c1: int,
        r2: int,
        c2: int,
        fill_iid: int,
        state: str = "",
        run_binding: bool = False,
    ) -> int:
        type_ = self.selection_boxes[fill_iid].type_
        self.selection_boxes[fill_iid].coords = Box_nt(r1, c1, r2, c2)
        if type_ == "cells":
            mt_border_col = self.PAR.ops.table_selected_cells_border_fg
        elif type_ == "rows":
            mt_border_col = self.PAR.ops.table_selected_rows_border_fg
        elif type_ == "columns":
            mt_border_col = self.PAR.ops.table_selected_columns_border_fg
        if not state:
            if r2 - r1 > 1 or c2 - c1 > 1:
                state = "normal"
            elif next(reversed(self.selection_boxes)) == fill_iid:
                state = "hidden"
            else:
                state = "normal"
        if self.selected.fill_iid == fill_iid:
            self.selected = self.selected._replace(box=Box_nt(r1, c1, r2, c2))
        x1, y1, x2, y2 = self.box_coords_x_canvas_coords(r1, c1, r2, c2)
        self.selection_boxes[fill_iid].state = state
        if bd_iid := self.selection_boxes[fill_iid].bd_iid:
            if self.PAR.ops.show_selected_cells_border:
                self.display_border(
                    x1,
                    y1,
                    x2,
                    y2,
                    fill=mt_border_col,
                    state="normal",
                    tags=f"{type_}bd",
                    width=1,
                    iid=bd_iid,
                )
                self.tag_raise(bd_iid)
            else:
                self.hide_border(bd_iid)
        if run_binding:
            self.run_selection_binding(type_)
        return fill_iid

    def run_selection_binding(self, type_: str) -> None:
        if type_ == "cells":
            sel_event = self.get_select_event(being_drawn_item=self.being_drawn_item)
            if self.selection_binding_func:
                self.selection_binding_func(sel_event)
        elif type_ == "rows":
            sel_event = self.get_select_event(being_drawn_item=self.RI.being_drawn_item)
            if self.RI.selection_binding_func:
                self.RI.selection_binding_func(sel_event)
        elif type_ == "columns":
            sel_event = self.get_select_event(being_drawn_item=self.CH.being_drawn_item)
            if self.CH.selection_binding_func:
                self.CH.selection_binding_func(sel_event)
        self.PAR.emit_event("<<SheetSelect>>", data=sel_event)

    def recreate_all_selection_boxes(self) -> None:
        if not self.selected:
            return

        modified = False
        row_limit = len(self.row_positions) - 1
        col_limit = len(self.col_positions) - 1

        for item, box in self.get_selection_items():
            r1, c1, r2, c2 = box.coords
            # check coordinates
            r1 = min(r1, row_limit - (1 if row_limit > 0 else 0))
            c1 = min(c1, col_limit - (1 if col_limit > 0 else 0))
            r2 = min(r2, row_limit)
            c2 = min(c2, col_limit)

            modified = modified or (r1 >= row_limit or c1 >= col_limit or r2 > row_limit or c2 > col_limit)
            self.recreate_selection_box(r1, c1, r2, c2, item, run_binding=False)

        if self.selected:
            r, c = self.selected.row, self.selected.column
            if r < row_limit and c < col_limit:
                self.set_currently_selected(r, c, item=self.selected.fill_iid, run_binding=False)
            else:
                box = self.selection_boxes[self.selected.fill_iid]
                self.set_currently_selected(box.coords[0], box.coords[1], item=box.fill_iid, run_binding=False)

        if modified:
            self.PAR.emit_event(
                "<<SheetSelect>>",
                data=self.get_select_event(self.being_drawn_item),
            )

    def get_redraw_selections(self, startr: int, endr: int, startc: int, endc: int) -> dict:
        d = defaultdict(set)
        ignore_hidd_current = not self.PAR.ops.show_selected_cells_border
        for _, box in self.get_selection_items():
            r1, c1, r2, c2 = box.coords
            if box.state == "normal" or ignore_hidd_current:
                if box.type_ == "cells":
                    for r in range(startr, endr):
                        for c in range(startc, endc):
                            if r1 <= r and c1 <= c and r2 > r and c2 > c:
                                d["cells"].add((r, c))
                elif box.type_ == "rows":
                    for r in range(startr, endr):
                        if r1 <= r and r2 > r:
                            d["rows"].add(r)
                elif box.type_ == "columns":
                    for c in range(startc, endc):
                        if c1 <= c and c2 > c:
                            d["columns"].add(c)
        return d

    def get_selected_min_max(self) -> tuple[int, int, int, int] | tuple[None, None, None, None]:
        if not self.get_selection_items():
            return None, None, None, None
        min_y, min_x = float("inf"), float("inf")
        max_y, max_x = 0, 0

        for _, box in self.get_selection_items():
            r1, c1, r2, c2 = box.coords
            min_y, min_x = min(min_y, r1), min(min_x, c1)
            max_y, max_x = max(max_y, r2), max(max_x, c2)

        return (min_y, min_x, max_y, max_x) if min_y != float("inf") else (None, None, None, None)

    def get_selected_rows(
        self,
        get_cells: bool = False,
        get_cells_as_rows: bool = False,
    ) -> set[int] | set[tuple[int, int]]:
        if get_cells:
            s = {
                (r, c)
                for _, box in self.get_selection_items(cells=False, columns=False)
                for r in range(box.coords.from_r, box.coords.upto_r)
                for c in range(0, len(self.col_positions) - 1)
            }
            if get_cells_as_rows:
                return s | self.get_selected_cells()
        else:
            s = {
                r
                for _, box in self.get_selection_items(cells=False, columns=False)
                for r in range(box.coords.from_r, box.coords.upto_r)
            }
            if get_cells_as_rows:
                return s | set(map(itemgetter(0), self.gen_selected_cells()))
        return s

    def get_selected_cols(
        self,
        get_cells: bool = False,
        get_cells_as_cols: bool = False,
    ) -> set[int] | set[tuple[int, int]]:
        if get_cells:
            s = {
                (r, c)
                for _, box in self.get_selection_items(cells=False, rows=False)
                for r in range(0, len(self.row_positions) - 1)
                for c in range(box.coords.from_c, box.coords.upto_c)
            }
            if get_cells_as_cols:
                return s | self.get_selected_cells()
        else:
            s = {
                c
                for _, box in self.get_selection_items(cells=False, rows=False)
                for c in range(box.coords.from_c, box.coords.upto_c)
            }
            if get_cells_as_cols:
                return s | set(map(itemgetter(1), self.gen_selected_cells()))
        return s

    def get_selected_cells(
        self,
        get_rows: bool = False,
        get_cols: bool = False,
    ) -> set[tuple[int, int]]:
        return {
            (r, c)
            for _, box in self.get_selection_items(rows=get_rows, columns=get_cols)
            for r in range(box.coords.from_r, box.coords.upto_r)
            for c in range(box.coords.from_c, box.coords.upto_c)
        }

    def gen_selected_cells(
        self,
        get_rows: bool = False,
        get_cols: bool = False,
    ) -> Generator[tuple[int, int]]:
        yield from (
            (r, c)
            for _, box in self.get_selection_items(rows=get_rows, columns=get_cols)
            for r in range(box.coords.from_r, box.coords.upto_r)
            for c in range(box.coords.from_c, box.coords.upto_c)
        )

    def get_all_selection_boxes(self) -> tuple[tuple[int, int, int, int]]:
        return tuple(box.coords for _, box in self.get_selection_items())

    def get_all_selection_boxes_with_types(self) -> list[tuple[tuple[int, int, int, int], str]]:
        return [Box_st(box.coords, box.type_) for _, box in self.get_selection_items()]

    def all_selected(self) -> bool:
        return any(
            not r1 and not c1 and r2 == len(self.row_positions) - 1 and c2 == len(self.col_positions) - 1
            for r1, c1, r2, c2 in self.get_all_selection_boxes()
        )

    def cell_selected(
        self,
        r: int,
        c: int,
        inc_cols: bool = False,
        inc_rows: bool = False,
    ) -> bool:
        return (
            isinstance(r, int)
            and isinstance(c, int)
            and any(
                box.coords.from_r <= r and box.coords.upto_r > r and box.coords.from_c <= c and box.coords.upto_c > c
                for _, box in self.get_selection_items(
                    rows=inc_rows,
                    columns=inc_cols,
                )
            )
        )

    def col_selected(self, c: int, cells: bool = False) -> bool:
        return isinstance(c, int) and any(
            box.coords.from_c <= c and box.coords.upto_c > c
            for _, box in self.get_selection_items(
                cells=cells,
                rows=False,
            )
        )

    def row_selected(self, r: int, cells: bool = False) -> bool:
        return isinstance(r, int) and any(
            box.coords.from_r <= r and box.coords.upto_r > r
            for _, box in self.get_selection_items(
                cells=cells,
                columns=False,
            )
        )

    def anything_selected(
        self,
        exclude_columns: bool = False,
        exclude_rows: bool = False,
        exclude_cells: bool = False,
    ) -> list[int]:
        return [
            item
            for item, _ in self.get_selection_items(
                columns=not exclude_columns,
                rows=not exclude_rows,
                cells=not exclude_cells,
            )
        ]

    def open_cell(
        self,
        event: Any = None,
        ignore_existing_editor: bool = False,
    ) -> None:
        if not self.anything_selected() or (not ignore_existing_editor and self.text_editor.open):
            return
        if not self.selected:
            return
        r, c = self.selected.row, self.selected.column
        datacn = self.datacn(c)
        datarn = self.datarn(r)
        if self.get_cell_kwargs(datarn, datacn, key="readonly"):
            return
        elif self.get_cell_kwargs(datarn, datacn, key="dropdown") or self.get_cell_kwargs(
            datarn, datacn, key="checkbox"
        ):
            if event_opens_dropdown_or_checkbox(event):
                if self.get_cell_kwargs(datarn, datacn, key="dropdown"):
                    self.open_dropdown_window(r, c, event=event)
                elif self.get_cell_kwargs(datarn, datacn, key="checkbox"):
                    self.click_checkbox(r=r, c=c, datarn=datarn, datacn=datacn)
        else:
            self.open_text_editor(event=event, r=r, c=c, dropdown=False)

    # displayed indexes
    def get_cell_align(self, r: int, c: int) -> str:
        datarn = self.datarn(r)
        datacn = self.datacn(c)
        cell_alignment = self.get_cell_kwargs(datarn, datacn, key="align")
        if cell_alignment:
            return cell_alignment
        return self.align

    # displayed indexes
    def open_text_editor(
        self,
        event: Any = None,
        r: int = 0,
        c: int = 0,
        text: str | None = None,
        state: str = "normal",
        dropdown: bool = False,
    ) -> bool:
        text = f"{self.get_cell_data(self.datarn(r), self.datacn(c), none_to_empty_str=True)}"
        extra_func_key = "??"
        if event_opens_dropdown_or_checkbox(event):
            if hasattr(event, "keysym") and event.keysym in ("Return", "F2", "BackSpace"):
                extra_func_key = event.keysym
                if event.keysym == "BackSpace":
                    text = ""
        elif event_has_char_key(event):
            extra_func_key = text = event.char
        elif event is not None:
            return False
        if self.extra_begin_edit_cell_func:
            try:
                text = self.extra_begin_edit_cell_func(
                    event_dict(
                        name="begin_edit_table",
                        sheet=self.PAR.name,
                        key=extra_func_key,
                        value=text,
                        loc=Loc(r, c),
                        row=r,
                        column=c,
                        boxes=self.get_boxes(),
                        selected=self.selected,
                    )
                )
            except Exception:
                return False
            if text is None:
                return False
            else:
                text = text if isinstance(text, str) else f"{text}"
        if self.PAR.ops.cell_auto_resize_enabled:
            self.set_cell_size_to_text(r, c, only_if_too_small=True, redraw=True, run_binding=True)
        if self.text_editor.open and (r, c) == self.text_editor.coords:
            self.text_editor.window.set_text(self.text_editor.get() + "" if not isinstance(text, str) else text)
            return False
        self.hide_text_editor()
        self.hide_tooltip()
        if not self.see(r, c):
            self.main_table_redraw_grid_and_text(True, True)
        x = self.col_positions[c]
        y = self.row_positions[r]
        w = self.col_positions[c + 1] - x + 1
        h = self.row_positions[r + 1] - y + 1
        kwargs = {
            "menu_kwargs": get_menu_kwargs(self.PAR.ops),
            "sheet_ops": self.PAR.ops,
            "font": self.PAR.ops.table_font,
            "border_color": self.PAR.ops.table_selected_box_cells_fg,
            "text": text,
            "state": state,
            "width": w,
            "height": h,
            "show_border": True,
            **get_bg_fg(self.PAR.ops),
            "align": self.get_cell_align(r, c),
            "r": r,
            "c": c,
        }
        if not self.text_editor.window:
            self.text_editor.window = TextEditor(
                self, newline_binding=self.text_editor_newline_binding, rc_bindings=self.PAR.ops.rc_bindings
            )
            self.text_editor.canvas_id = self.create_window((x, y), window=self.text_editor.window, anchor="nw")
        self.text_editor.window.reset(**kwargs)
        if not self.text_editor.open:
            self.itemconfig(self.text_editor.canvas_id, state="normal")
            self.text_editor.open = True
        self.coords(self.text_editor.canvas_id, x, y)
        for b in text_editor_newline_bindings:
            self.text_editor.tktext.bind(b, self.text_editor_newline_binding)
        for b in text_editor_close_bindings:
            self.text_editor.tktext.bind(b, self.close_text_editor)
        if not dropdown:
            self.text_editor.tktext.focus_set()
            self.text_editor.window.scroll_to_bottom()
            self.text_editor.tktext.bind("<FocusOut>", self.close_text_editor)
        for key, func in self.text_editor_user_bound_keys.items():
            self.text_editor.tktext.bind(key, func)
        return True

    # displayed indexes
    def text_editor_newline_binding(
        self,
        event: Any = None,
        check_lines: bool = True,
    ) -> None:
        r, c = self.text_editor.coords
        curr_height = self.text_editor.window.winfo_height()
        if curr_height < self.min_row_height:
            return
        if (
            not check_lines
            or self.get_lines_cell_height(
                self.text_editor.window.get_num_lines() + 1,
                font=self.text_editor.tktext.cget("font"),
            )
            > curr_height
        ):
            new_height = min(
                curr_height + self.table_txt_height,
                self.scrollregion[3] - self.scrollregion[1] - self.row_positions[r],
            )
            if new_height != curr_height:
                self.text_editor.window.config(height=new_height)
                if self.dropdown.open and self.dropdown.get_coords() == (r, c):
                    text_editor_h = self.text_editor.window.winfo_height()
                    win_h, anchor = self.get_dropdown_height_anchor(r, c, text_editor_h)
                    if anchor == "nw":
                        self.coords(
                            self.dropdown.canvas_id,
                            self.col_positions[c],
                            self.row_positions[r] + text_editor_h - 1,
                        )
                        self.itemconfig(self.dropdown.canvas_id, anchor=anchor, height=win_h)
                    elif anchor == "sw":
                        self.coords(
                            self.dropdown.canvas_id,
                            self.col_positions[c],
                            self.row_positions[r],
                        )
                        self.itemconfig(self.dropdown.canvas_id, anchor=anchor, height=win_h)

    def refresh_open_window_positions(self, zoom: Literal["in", "out"]):
        if self.text_editor.open:
            r, c = self.text_editor.coords
            self.text_editor.window.config(height=self.row_positions[r + 1] - self.row_positions[r])
            self.text_editor.tktext.config(font=self.PAR.ops.table_font)
            self.coords(
                self.text_editor.canvas_id,
                self.col_positions[c],
                self.row_positions[r],
            )
        if self.dropdown.open:
            if zoom == "in":
                self.dropdown.window.zoom_in()
            elif zoom == "out":
                self.dropdown.window.zoom_out()
            r, c = self.dropdown.get_coords()
            if self.text_editor.open:
                text_editor_h = self.text_editor.window.winfo_height()
                win_h, anchor = self.get_dropdown_height_anchor(r, c, text_editor_h)
            else:
                text_editor_h = self.row_positions[r + 1] - self.row_positions[r] + 1
                anchor = self.itemcget(self.dropdown.canvas_id, "anchor")
                # win_h = 0
            self.dropdown.window.config(width=self.col_positions[c + 1] - self.col_positions[c] + 1)
            if anchor == "nw":
                self.coords(
                    self.dropdown.canvas_id,
                    self.col_positions[c],
                    self.row_positions[r] + text_editor_h - 1,
                )
                # self.itemconfig(self.dropdown.canvas_id, anchor=anchor, height=win_h)
            elif anchor == "sw":
                self.coords(
                    self.dropdown.canvas_id,
                    self.col_positions[c],
                    self.row_positions[r],
                )
                # self.itemconfig(self.dropdown.canvas_id, anchor=anchor, height=win_h)

    def hide_text_editor(self) -> None:
        if self.text_editor.open:
            for binding in text_editor_to_unbind:
                self.text_editor.tktext.unbind(binding)
            self.itemconfig(self.text_editor.canvas_id, state="hidden")
            self.text_editor.open = False

    def close_text_editor(self, event: tk.Event) -> Literal["break"] | None:
        # checking if text editor should be closed or not
        # errors if __tk_filedialog is open
        try:
            focused = self.focus_get()
        except Exception:
            focused = None
        try:
            if focused == self.text_editor.tktext.rc_popup_menu:
                return "break"
        except Exception:
            pass
        if focused is None:
            return "break"
        if event.keysym == "Escape":
            self.hide_text_editor_and_dropdown()
            self.focus_set()
            return
        # setting cell data with text editor value
        value = self.text_editor.get()
        r, c = self.text_editor.coords
        datarn, datacn = self.datarn(r), self.datacn(c)
        event_data = self.new_single_edit_event(
            r, c, datarn, datacn, event.keysym, self.get_cell_data(datarn, datacn), value
        )
        value, event_data = self.single_edit_run_validation(datarn, datacn, event_data)
        edited = False
        if value is not None and (
            edited := self.set_cell_data_undo(
                r=r,
                c=c,
                datarn=datarn,
                datacn=datacn,
                value=value,
                redraw=False,
            )
        ):
            try_binding(self.extra_end_edit_cell_func, event_data)
        if (
            r is not None
            and c is not None
            and self.selected
            and r == self.selected.row
            and c == self.selected.column
            and (self.single_selection_enabled or self.toggle_selection_enabled)
            and (edited or self.cell_equal_to(datarn, datacn, value))
        ):
            self.go_to_next_cell(r, c, event.keysym)
        self.recreate_all_selection_boxes()
        self.hide_text_editor_and_dropdown()
        if event.keysym != "FocusOut":
            self.focus_set()
        return "break"

    def go_to_next_cell(self, r: int, c: int, key: Any = "Return") -> None:
        r1, c1, r2, c2 = self.selection_boxes[self.selected.fill_iid].coords
        numrows, numcols = r2 - r1, c2 - c1
        if key == "Return":
            direction = self.PAR.ops.edit_cell_return
        elif key == "Tab":
            direction = self.PAR.ops.edit_cell_tab
        else:
            return
        if numcols == 1 and numrows == 1:
            if direction == "right":
                new_r, new_c = cell_right_within_box(
                    r,
                    c,
                    0,
                    0,
                    len(self.row_positions) - 1,
                    len(self.col_positions) - 1,
                    len(self.row_positions) - 1,
                    len(self.col_positions) - 1,
                )
            elif direction == "down":
                new_r, new_c = cell_down_within_box(
                    r,
                    c,
                    0,
                    0,
                    len(self.row_positions) - 1,
                    len(self.col_positions) - 1,
                    len(self.row_positions) - 1,
                    len(self.col_positions) - 1,
                )
            if direction in ("right", "down"):
                self.select_cell(new_r, new_c)
                self.see(new_r, new_c)
        else:
            if direction == "right":
                new_r, new_c = cell_right_within_box(r, c, r1, c1, r2, c2, numrows, numcols)
            elif direction == "down":
                new_r, new_c = cell_down_within_box(r, c, r1, c1, r2, c2, numrows, numcols)
            if direction in ("right", "down"):
                self.set_currently_selected(new_r, new_c, item=self.selected.fill_iid)
                self.see(new_r, new_c)

    def single_edit_run_validation(
        self, datarn: int, datacn: int, event_data: EventDataDict
    ) -> tuple[Any, EventDataDict]:
        value = event_data.value
        if self.edit_validation_func and (value := self.edit_validation_func(event_data)) is not None:
            event_data["data"][(datarn, datacn)] = value
            event_data["value"] = value
        if self.bulk_table_edit_validation_func:
            self.bulk_table_edit_validation_func(event_data)
            if (datarn, datacn) in event_data["data"]:
                if event_data["data"][(datarn, datacn)] is not None:
                    value = event_data["data"][(datarn, datacn)]
            else:
                value = None
        return value, event_data

    def tab_key(self, event: Any = None) -> str:
        if self.selected:
            r, c = self.selected.row, self.selected.column
            r1, c1, r2, c2 = self.selection_boxes[self.selected.fill_iid].coords
            numcols = c2 - c1
            numrows = r2 - r1
        else:
            if (
                self.row_positions == [0]
                or self.col_positions == [0]
                or not self.row_positions
                or not self.col_positions
            ):
                return
            r, c = len(self.row_positions) - 2, len(self.col_positions) - 2
            r1, c1, r2, c2 = 0, 0, len(self.row_positions) - 1, len(self.col_positions) - 1
            numcols, numrows = 1, 1
        if numcols == 1 and numrows == 1:
            new_r, new_c = cell_right_within_box(
                r,
                c,
                0,
                0,
                len(self.row_positions) - 1,
                len(self.col_positions) - 1,
                len(self.row_positions) - 1,
                len(self.col_positions) - 1,
            )
            self.select_cell(new_r, new_c)
        else:
            new_r, new_c = cell_right_within_box(r, c, r1, c1, r2, c2, numrows, numcols)
            self.set_currently_selected(new_r, new_c, item=self.selected.fill_iid)
        self.see(new_r, new_c, bottom_right_corner=True)
        if not self.PAR.ops.show_selected_cells_border:
            self.refresh()
        return "break"

    def get_space_bot(self, r: int, text_editor_h: int | None = None) -> int:
        if len(self.row_positions) <= 1:
            if text_editor_h is None:
                win_h = int(self.winfo_height())
                sheet_h = int(1 + self.PAR.ops.empty_vertical)
            else:
                win_h = int(self.winfo_height() - text_editor_h)
                sheet_h = int(1 + self.PAR.ops.empty_vertical - text_editor_h)
        else:
            if text_editor_h is None:
                win_h = int(self.canvasy(0) + self.winfo_height() - self.row_positions[r + 1])
                sheet_h = int(self.row_positions[-1] + 1 + self.PAR.ops.empty_vertical - self.row_positions[r + 1])
            else:
                win_h = int(self.canvasy(0) + self.winfo_height() - (self.row_positions[r] + text_editor_h))
                sheet_h = int(
                    self.row_positions[-1] + 1 + self.PAR.ops.empty_vertical - (self.row_positions[r] + text_editor_h)
                )
        win_h = max(0, win_h - 1)
        sheet_h = max(0, sheet_h - 1)
        return win_h if win_h >= sheet_h else sheet_h

    def get_dropdown_height_anchor(self, r: int, c: int, text_editor_h: int | None = None) -> tuple:
        win_h = 5
        datarn, datacn = self.datarn(r), self.datacn(c)
        for i, v in enumerate(self.get_cell_kwargs(datarn, datacn, key="dropdown")["values"]):
            v_numlines = len(v.split("\n") if isinstance(v, str) else f"{v}".split("\n"))
            if v_numlines > 1:
                win_h += 8 + (v_numlines * self.table_txt_height)  # end of cell
            else:
                win_h += self.min_row_height
            if i == 5:
                break
        win_h = min(win_h, 500)
        space_bot = self.get_space_bot(r, text_editor_h)
        space_top = int(self.row_positions[r])
        anchor = "nw"
        win_h2 = int(win_h)
        if win_h > space_bot:
            if space_bot >= space_top:
                anchor = "nw"
                win_h = space_bot - 1
            elif space_top > space_bot:
                anchor = "sw"
                win_h = space_top - 1
        if win_h < self.table_txt_height + 5:
            win_h = self.table_txt_height + 5
        elif win_h > win_h2:
            win_h = win_h2
        return win_h, anchor

    def dropdown_text_editor_modified(
        self,
        event: tk.Misc,
    ) -> None:
        r, c = self.dropdown.get_coords()
        event_data = event_dict(
            name="table_dropdown_modified",
            sheet=self.PAR.name,
            value=self.text_editor.get(),
            loc=Loc(r, c),
            row=r,
            column=c,
            boxes=self.get_boxes(),
            selected=self.selected,
        )
        try_binding(self.dropdown.window.modified_function, event_data)
        # return to tk.Text action if control/command is held down
        # or keysym was not a character
        if (hasattr(event, "state") and event.state & (0x0004 | 0x00000010)) or (
            hasattr(event, "keysym") and len(event.keysym) > 2 and event.keysym != "space"
        ):
            return
        self.text_editor.autocomplete(self.dropdown.window.search_and_see(event_data))
        return "break"

    def open_dropdown_window(
        self,
        r: int,
        c: int,
        event: Any = None,
    ) -> None:
        self.hide_text_editor()
        datarn = self.datarn(r)
        datacn = self.datacn(c)
        kwargs = self.get_cell_kwargs(datarn, datacn, key="dropdown")
        if kwargs["state"] == "disabled":
            return
        if kwargs["state"] == "normal" and not self.open_text_editor(event=event, r=r, c=c, dropdown=True):
            return
        win_h, anchor = self.get_dropdown_height_anchor(r, c)
        win_w = self.col_positions[c + 1] - self.col_positions[c] + 1
        if anchor == "nw":
            if kwargs["state"] == "normal":
                self.text_editor.window.update_idletasks()
                ypos = self.row_positions[r] + self.text_editor.window.winfo_height() - 1
            else:
                ypos = self.row_positions[r + 1]
        else:
            ypos = self.row_positions[r]
        reset_kwargs = {
            "r": r,
            "c": c,
            "bg": self.PAR.ops.table_editor_bg,
            "fg": self.PAR.ops.table_editor_fg,
            "select_bg": self.PAR.ops.table_editor_select_bg,
            "select_fg": self.PAR.ops.table_editor_select_fg,
            "width": win_w,
            "height": win_h,
            "font": self.PAR.ops.table_font,
            "ops": self.PAR.ops,
            "outline_color": self.get_selected_box_bg_fg(type_="cells")[1],
            "align": self.get_cell_align(r, c),
            "values": kwargs["values"],
            "search_function": kwargs["search_function"],
            "modified_function": kwargs["modified_function"],
        }
        if self.dropdown.window:
            self.dropdown.window.reset(**reset_kwargs)
            self.coords(self.dropdown.canvas_id, self.col_positions[c], ypos)
            self.itemconfig(self.dropdown.canvas_id, state="normal", anchor=anchor)
            self.dropdown.window.tkraise()
        else:
            self.dropdown.window = self.PAR._dropdown_cls(
                self.winfo_toplevel(),
                **reset_kwargs,
                close_dropdown_window=self.close_dropdown_window,
                arrowkey_RIGHT=self.arrowkey_RIGHT,
                arrowkey_LEFT=self.arrowkey_LEFT,
            )
            self.dropdown.canvas_id = self.create_window(
                (self.col_positions[c], ypos),
                window=self.dropdown.window,
                anchor=anchor,
            )
        self.update_idletasks()
        if kwargs["state"] == "normal":
            self.text_editor.tktext.bind(
                "<KeyRelease>",
                self.dropdown_text_editor_modified,
            )
            try:
                self.after(1, lambda: self.text_editor.tktext.focus())
                self.after(2, self.text_editor.window.scroll_to_bottom())
            except Exception:
                return
            redraw = False
        else:
            self.dropdown.window.bind("<FocusOut>", lambda _: self.close_dropdown_window(r, c))
            self.dropdown.window.bind("<Escape>", self.close_dropdown_window)
            self.dropdown.window.focus_set()
            redraw = True
        self.dropdown.open = True
        if redraw:
            self.main_table_redraw_grid_and_text(redraw_header=False, redraw_row_index=False)

    # displayed indexes, not data
    def close_dropdown_window(
        self,
        r: int | None = None,
        c: int | None = None,
        selection: Any = None,
        redraw: bool = True,
    ) -> None:
        if r is not None and c is not None and selection is not None:
            datacn = self.datacn(c)
            datarn = self.datarn(r)
            kwargs = self.get_cell_kwargs(datarn, datacn, key="dropdown")
            event_data = self.new_single_edit_event(
                r, c, datarn, datacn, "??", self.get_cell_data(datarn, datacn), selection
            )
            try_binding(kwargs["select_function"], event_data)
            selection, event_data = self.single_edit_run_validation(datarn, datacn, event_data)
            if selection is not None and self.set_cell_data_undo(
                r,
                c,
                datarn=datarn,
                datacn=datacn,
                value=selection,
                redraw=not redraw,
            ):
                try_binding(self.extra_end_edit_cell_func, event_data)
            self.recreate_all_selection_boxes()
        self.focus_set()
        self.hide_text_editor_and_dropdown(redraw=redraw)

    def hide_text_editor_and_dropdown(self, redraw: bool = True) -> None:
        self.hide_text_editor()
        self.hide_dropdown_window()
        if redraw:
            self.refresh()

    def mouseclick_outside_editor_or_dropdown(self) -> tuple[int, int] | None:
        closed_dd_coords = self.dropdown.get_coords()
        if self.text_editor.open:
            self.close_text_editor(new_tk_event("ButtonPress-1"))
        self.hide_dropdown_window()
        self.focus_set()
        return closed_dd_coords

    def mouseclick_outside_editor_or_dropdown_all_canvases(self) -> tuple[int, int] | None:
        self.CH.mouseclick_outside_editor_or_dropdown()
        self.RI.mouseclick_outside_editor_or_dropdown()
        return self.mouseclick_outside_editor_or_dropdown()

    def hide_dropdown_editor_all_canvases(self) -> None:
        self.hide_text_editor_and_dropdown(redraw=False)
        self.RI.hide_text_editor_and_dropdown(redraw=False)
        self.CH.hide_text_editor_and_dropdown(redraw=False)

    def hide_dropdown_window(self) -> None:
        if self.dropdown.open:
            self.dropdown.window.unbind("<FocusOut>")
            self.itemconfig(self.dropdown.canvas_id, state="hidden")
            self.dropdown.open = False

    def click_checkbox(
        self,
        r: int,
        c: int,
        datarn: int | None = None,
        datacn: int | None = None,
        undo: bool = True,
        redraw: bool = True,
    ) -> None:
        if datarn is None:
            datarn = self.datarn(r)
        if datacn is None:
            datacn = self.datacn(c)
        kwargs = self.get_cell_kwargs(datarn, datacn, key="checkbox")
        if kwargs["state"] == "normal":
            pre_edit_value = self.get_cell_data(datarn, datacn)
            value = not self.data[datarn][datacn] if isinstance(self.data[datarn][datacn], bool) else False
            self.set_cell_data_undo(
                r,
                c,
                value=value,
                undo=undo,
                cell_resize=False,
                check_input_valid=False,
            )
            event_data = self.new_single_edit_event(r, c, datarn, datacn, "??", pre_edit_value, value)
            if kwargs["check_function"] is not None:
                kwargs["check_function"](event_data)
            try_binding(self.extra_end_edit_cell_func, event_data)
        if redraw:
            self.refresh()

    # internal event use
    def set_cell_data_undo(
        self,
        r: int = 0,
        c: int = 0,
        datarn: int | None = None,
        datacn: int | None = None,
        value: str | None = None,
        undo: bool = True,
        cell_resize: bool = True,
        redraw: bool = True,
        check_input_valid: bool = True,
    ) -> bool:
        if value is None:
            value = ""
        if datacn is None:
            datacn = self.datacn(c)
        if datarn is None:
            datarn = self.datarn(r)
        event_data = event_dict(
            name="edit_table",
            sheet=self.PAR.name,
            widget=self,
            cells_table={(datarn, datacn): self.get_cell_data(datarn, datacn)},
            boxes=self.get_boxes(),
            selected=self.selected,
        )
        if not check_input_valid or self.input_valid_for_cell(datarn, datacn, value):
            if self.undo_enabled and undo:
                self.undo_stack.append(stored_event_dict(event_data))
            self.set_cell_data(datarn, datacn, value)
            if cell_resize and self.PAR.ops.cell_auto_resize_enabled:
                self.set_cell_size_to_text(r, c, only_if_too_small=True, redraw=redraw, run_binding=True)
            self.sheet_modified(event_data)
            return True
        return False

    def set_cell_data(
        self,
        datarn: int,
        datacn: int,
        value: Any,
        kwargs: dict | None = None,
        expand_sheet: bool = True,
    ) -> None:
        if expand_sheet:
            if datarn >= len(self.data):
                self.fix_data_len(datarn, datacn)
            elif datacn >= len(self.data[datarn]):
                self.fix_row_len(datarn, datacn)
        if expand_sheet or (len(self.data) > datarn and len(self.data[datarn]) > datacn):
            if (
                datarn,
                datacn,
            ) in self.cell_options and "checkbox" in self.cell_options[(datarn, datacn)]:
                self.data[datarn][datacn] = try_to_bool(value)
            else:
                if not kwargs:
                    kwargs = self.get_cell_kwargs(datarn, datacn, key="format")
                if kwargs:
                    if kwargs["formatter"] is None:
                        self.data[datarn][datacn] = format_data(value=value, **kwargs)
                    else:
                        self.data[datarn][datacn] = kwargs["formatter"](value, **kwargs)
                else:
                    self.data[datarn][datacn] = value

    def format_value(self, datarn: int, datacn: int, value: Any) -> Any:
        if (datarn, datacn) in self.cell_options and "checkbox" in self.cell_options[(datarn, datacn)]:
            return try_to_bool(value)
        elif kwargs := self.get_cell_kwargs(datarn, datacn, key="format"):
            if kwargs["formatter"] is None:
                return format_data(value=value, **kwargs)
            else:
                return kwargs["formatter"](value, **kwargs)
        else:
            return value

    def get_value_for_empty_cell(self, datarn: int, datacn: int, r_ops: bool = True, c_ops: bool = True) -> Any:
        if self.get_cell_kwargs(
            datarn,
            datacn,
            key="checkbox",
            cell=r_ops and c_ops,
            row=r_ops,
            column=c_ops,
        ):
            return False
        elif (
            kwargs := self.get_cell_kwargs(
                datarn,
                datacn,
                key="dropdown",
                cell=r_ops and c_ops,
                row=r_ops,
                column=c_ops,
            )
        ) and kwargs["validate_input"]:
            if kwargs["default_value"] is None:
                if kwargs["values"]:
                    return safe_copy(kwargs["values"][0])
                else:
                    return self.format_value(datarn, datacn, "")
            else:
                return safe_copy(kwargs["default_value"])
        else:
            return self.format_value(datarn, datacn, "")

    def get_empty_row_seq(
        self, datarn: int, end: int, start: int = 0, r_ops: bool = True, c_ops: bool = True
    ) -> list[Any]:
        return [self.get_value_for_empty_cell(datarn, datacn, r_ops=r_ops, c_ops=c_ops) for datacn in range(start, end)]

    def gen_empty_row_seq(
        self, datarn: int, end: int, start: int = 0, r_ops: bool = True, c_ops: bool = True
    ) -> Generator[Any]:
        yield from (self.get_value_for_empty_cell(datarn, c, r_ops=r_ops, c_ops=c_ops) for c in range(start, end))

    def fix_row_len(self, datarn: int, datacn: int) -> None:
        self.data[datarn].extend(self.get_empty_row_seq(datarn, end=datacn + 1, start=len(self.data[datarn])))

    def fix_row_values(self, datarn: int, start: int | None = None, end: int | None = None):
        if datarn < len(self.data):
            for datacn, v in enumerate(islice(self.data[datarn], start, end)):
                if not self.input_valid_for_cell(datarn, datacn, v):
                    self.data[datarn][datacn] = self.get_value_for_empty_cell(datarn, datacn)

    def fix_data_len(self, datarn: int, datacn: int | None = None) -> int:
        ncols = self.total_data_cols() if datacn is None else datacn + 1
        self.data.extend(self.get_empty_row_seq(rn, end=ncols, start=0) for rn in range(len(self.data), datarn + 1))
        return len(self.data)

    def reapply_formatting(self) -> None:
        for c in gen_formatted(self.col_options):
            for r in range(len(self.data)):
                if not (
                    (r, c) in self.cell_options
                    and "format" in self.cell_options[(r, c)]
                    or r in self.row_options
                    and "format" in self.row_options[r]
                ):
                    self.set_cell_data(r, c, value=self.data[r][c])
        for r in gen_formatted(self.row_options):
            for c in range(len(self.data[r])):
                if not ((r, c) in self.cell_options and "format" in self.cell_options[(r, c)]):
                    self.set_cell_data(r, c, value=self.data[r][c])
        for r, c in gen_formatted(self.cell_options):
            if len(self.data) > r and len(self.data[r]) > c:
                self.set_cell_data(r, c, value=self.data[r][c])

    def delete_all_formatting(self, clear_values: bool = False) -> None:
        self.delete_cell_format("all", clear_values=clear_values)
        self.delete_row_format("all", clear_values=clear_values)
        self.delete_column_format("all", clear_values=clear_values)

    def delete_cell_format(
        self,
        datarn: Literal["all"] | int = "all",
        datacn: int = 0,
        clear_values: bool = False,
    ) -> None:
        if isinstance(datarn, str) and datarn.lower() == "all":
            itr = gen_formatted(self.cell_options)
        else:
            itr = ((datarn, datacn),)
        get_val = self.get_value_for_empty_cell
        for key in itr:
            try:
                del self.cell_options[key]["format"]
            except Exception:
                continue
            if clear_values:
                self.set_cell_data(*key, get_val(*key), expand_sheet=False)

    def delete_row_format(self, datarn: Literal["all"] | int = "all", clear_values: bool = False) -> None:
        itr = gen_formatted(self.row_options) if isinstance(datarn, str) and datarn.lower() == "all" else (datarn,)
        get_val = self.get_value_for_empty_cell
        for datarn in itr:
            try:
                del self.row_options[datarn]["format"]
            except Exception:
                continue
            if clear_values:
                for datacn in range(len(self.data[datarn])):
                    self.set_cell_data(datarn, datacn, get_val(datarn, datacn), expand_sheet=False)

    def delete_column_format(self, datacn: Literal["all"] | int = "all", clear_values: bool = False) -> None:
        itr = gen_formatted(self.col_options) if isinstance(datacn, str) and datacn.lower() == "all" else (datacn,)
        get_val = self.get_value_for_empty_cell
        for datacn in itr:
            try:
                del self.col_options[datacn]["format"]
            except Exception:
                continue
            if clear_values:
                for datarn in range(len(self.data)):
                    self.set_cell_data(datarn, datacn, get_val(datarn, datacn), expand_sheet=False)

    def cell_str(self, datarn: int, datacn: int, get_displayed: bool = False, **kwargs) -> str:
        """
        deals with possibility of formatter class being in self.data cell
        if cell is formatted - possibly returns invalid_value kwarg if
        cell value is not in datatypes kwarg
        if get displayed is true then Nones are replaced by
        """
        if get_displayed:
            # check for dropdown
            if (datarn, datacn) in self.cell_options and "dropdown" in self.cell_options[(datarn, datacn)]:
                kws = self.cell_options[(datarn, datacn)]["dropdown"]
            elif datarn in self.row_options and "dropdown" in self.row_options[datarn]:
                kws = self.row_options[datarn]["dropdown"]
            elif datacn in self.col_options and "dropdown" in self.col_options[datacn]:
                kws = self.col_options[datacn]["dropdown"]
            else:
                kws = None
            if kws and kws["text"] is not None:
                return f"{kws['text']}"

            # check for checkbox
            if (datarn, datacn) in self.cell_options and "checkbox" in self.cell_options[(datarn, datacn)]:
                kws = self.cell_options[(datarn, datacn)]["checkbox"]
            elif datarn in self.row_options and "checkbox" in self.row_options[datarn]:
                kws = self.row_options[datarn]["checkbox"]
            elif datacn in self.col_options and "checkbox" in self.col_options[datacn]:
                kws = self.col_options[datacn]["checkbox"]
            else:
                kws = None
            if kws:
                return f"{kws['text']}"
        try:
            value = self.data[datarn][datacn]
        except Exception:
            value = ""
        # check for format
        if (datarn, datacn) in self.cell_options and "format" in self.cell_options[(datarn, datacn)]:
            kws = self.cell_options[(datarn, datacn)]["format"]
        elif datarn in self.row_options and "format" in self.row_options[datarn]:
            kws = self.row_options[datarn]["format"]
        elif datacn in self.col_options and "format" in self.col_options[datacn]:
            kws = self.col_options[datacn]["format"]
        else:
            kws = None
        if kws:
            if kws["formatter"] is None:
                if get_displayed:
                    return data_to_str(value, **kws)
                else:
                    return f"{get_data_with_valid_check(value, **kws)}"
            else:
                if get_displayed:
                    # assumed given formatter class has __str__()
                    return f"{value}"
                else:
                    # assumed given formatter class has get_data_with_valid_check()
                    return f"{value.get_data_with_valid_check()}"
        else:
            return "" if value is None else str(value)

    def get_cell_data(
        self,
        datarn: int,
        datacn: int,
        get_displayed: bool = False,
        none_to_empty_str: bool = False,
        fmt_kw: dict | None = None,
    ) -> Any:
        if get_displayed:
            if fmt_kw:
                return format_data(value=self.cell_str(datarn, datacn, get_displayed=True), **fmt_kw)
            else:
                return self.cell_str(datarn, datacn, get_displayed=True)
        try:  # when successful try is more than twice as fast as len check
            value = self.data[datarn][datacn]
        except Exception:
            value = self.get_value_for_empty_cell(datarn, datacn)
        kwargs = self.get_cell_kwargs(datarn, datacn, key="format")
        if kwargs and kwargs["formatter"] is not None:
            value = value.value  # assumed given formatter class has value attribute
        if fmt_kw:
            value = format_data(value=value, **fmt_kw)
        return "" if (value is None and none_to_empty_str) else value

    def input_valid_for_cell(
        self,
        datarn: int,
        datacn: int,
        value: Any,
        check_readonly: bool = True,
        ignore_empty: bool = False,
    ) -> bool:
        if check_readonly and self.get_cell_kwargs(datarn, datacn, key="readonly"):
            return False
        elif self.get_cell_kwargs(datarn, datacn, key="format"):
            return True
        elif self.cell_equal_to(datarn, datacn, value, ignore_empty=ignore_empty, check_fmt=False) or (
            (kwargs := self.get_cell_kwargs(datarn, datacn, key="dropdown"))
            and kwargs["validate_input"]
            and value not in kwargs["values"]
        ):
            return False
        elif self.get_cell_kwargs(datarn, datacn, key="checkbox"):
            return is_bool_like(value)
        else:
            return True

    def cell_equal_to(
        self,
        datarn: int,
        datacn: int,
        new: Any,
        ignore_empty: bool = False,
        check_fmt: bool = True,
    ) -> bool:
        current = self.get_cell_data(datarn, datacn)
        if check_fmt:
            kws = self.get_cell_kwargs(datarn, datacn, key="format")
            if kws and kws["formatter"] is None:
                new = format_data(value=new, **kws)
        return current == new and not (ignore_empty and not current and not new)

    def get_cell_clipboard(self, datarn: int, datacn: int) -> str | int | float | bool:
        value = self.data[datarn][datacn] if len(self.data) > datarn and len(self.data[datarn]) > datacn else ""
        kwargs = self.get_cell_kwargs(datarn, datacn, key="format")
        if kwargs:
            if kwargs["formatter"] is None:
                return get_clipboard_data(value, **kwargs)
            else:
                # assumed given formatter class has get_clipboard_data() function
                # and it returns one of above type hints
                return value.get_clipboard_data()
        return f"{value}"

    def get_cell_kwargs(
        self,
        datarn: int,
        datacn: int,
        key: Hashable | None = "format",
        cell: bool = True,
        row: bool = True,
        column: bool = True,
    ) -> dict:
        if key is None:
            if cell and (datarn, datacn) in self.cell_options:
                return self.cell_options[(datarn, datacn)]

            elif row and datarn in self.row_options:
                return self.row_options[datarn]

            elif column and datacn in self.col_options:
                return self.col_options[datacn]

            else:
                return {}
        else:
            if cell and (datarn, datacn) in self.cell_options and key in self.cell_options[(datarn, datacn)]:
                return self.cell_options[(datarn, datacn)][key]

            elif row and datarn in self.row_options and key in self.row_options[datarn]:
                return self.row_options[datarn][key]

            elif column and datacn in self.col_options and key in self.col_options[datacn]:
                return self.col_options[datacn][key]

            else:
                return {}

    def datacn(self, c: int) -> int:
        return c if self.all_columns_displayed else self.displayed_columns[c]

    def datarn(self, r: int) -> int:
        return r if self.all_rows_displayed else self.displayed_rows[r]

    def dispcn(self, datacn: int) -> int:
        return datacn if self.all_columns_displayed else b_index(self.displayed_columns, datacn)

    def try_dispcn(self, datacn: int) -> int | None:
        try:
            return self.dispcn(datacn)
        except Exception:
            return None

    def disprn(self, datarn: int) -> int:
        return datarn if self.all_rows_displayed else b_index(self.displayed_rows, datarn)

    def try_disprn(self, datarn: int) -> int | None:
        try:
            return self.disprn(datarn)
        except Exception:
            return None
