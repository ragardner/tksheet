from __future__ import annotations

import tkinter as tk
from collections import defaultdict
from collections.abc import Callable, Hashable, Iterator, Sequence
from functools import partial
from itertools import islice, repeat
from math import ceil
from operator import itemgetter
from typing import Any, Literal

from .colors import color_map
from .constants import (
    USER_OS,
    _test_str,
    text_editor_close_bindings,
    text_editor_newline_bindings,
    text_editor_to_unbind,
)
from .formatters import is_bool_like, try_to_bool
from .functions import (
    any_editor_or_dropdown_open,
    consecutive_ranges,
    event_dict,
    event_has_char_key,
    event_opens_dropdown_or_checkbox,
    get_bg_fg,
    get_menu_kwargs,
    get_n2a,
    int_x_tuple,
    is_contiguous,
    mod_event_val,
    new_tk_event,
    recursive_bind,
    rounded_box_coords,
    safe_copy,
    stored_event_dict,
    try_b_index,
    try_binding,
    widget_descendants,
    wrap_text,
)
from .menus import build_empty_rc_menu, build_header_rc_menu
from .other_classes import DraggedRowColumn, DropdownStorage, EventDataDict, TextEditorStorage
from .row_index import RowIndex
from .sorting import sort_column, sort_rows_by_column, sort_tree_rows_by_column
from .text_editor import TextEditor
from .tooltip import Tooltip


class ColumnHeaders(tk.Canvas):
    def __init__(self, parent: tk.Misc, **kwargs) -> None:
        super().__init__(
            parent,
            background=parent.ops.header_bg,
            highlightthickness=0,
        )
        self.PAR = parent
        self.ops = self.PAR.ops
        self.current_height = None
        self.MT = None  # is set from within MainTable() __init__
        self.RI: RowIndex | None = None  # is set from within MainTable() __init__
        self.TL = None  # is set from within TopLeftRectangle() __init__
        self.tooltip = Tooltip(
            **{
                "parent": self,
                "sheet_ops": self.ops,
                "menu_kwargs": get_menu_kwargs(self.ops),
                **get_bg_fg(self.ops),
                "scrollbar_style": f"Sheet{self.PAR.unique_id}.Vertical.TScrollbar",
                "rc_bindings": self.ops.rc_bindings,
            }
        )
        self.tooltip_widgets = widget_descendants(self.tooltip)
        self.tooltip_coords, self.tooltip_after_id, self.tooltip_showing = None, None, False
        self.tooltip_cell_content = ""
        recursive_bind(self.tooltip, "<Leave>", self.close_tooltip_save)
        self.current_cursor = ""
        self.popup_menu_loc = None
        self.extra_begin_edit_cell_func = None
        self.extra_end_edit_cell_func = None
        self.b1_pressed_loc = None
        self.closed_dropdown = None
        self.being_drawn_item = None
        self.extra_motion_func = None
        self.extra_b1_press_func = None
        self.extra_b1_motion_func = None
        self.extra_b1_release_func = None
        self.extra_double_b1_func = None
        self.ch_extra_begin_drag_drop_func = None
        self.ch_extra_end_drag_drop_func = None
        self.ch_extra_begin_sort_rows_func = None
        self.ch_extra_end_sort_rows_func = None
        self.extra_rc_func = None
        self.selection_binding_func = None
        self.shift_selection_binding_func = None
        self.ctrl_selection_binding_func = None
        self.drag_selection_binding_func = None
        self.column_width_resize_func = None
        self.width_resizing_enabled = False
        self.height_resizing_enabled = False
        self.double_click_resizing_enabled = False
        self.col_selection_enabled = False
        self.drag_and_drop_enabled = False
        self.rc_delete_col_enabled = False
        self.rc_insert_col_enabled = False
        self.hide_columns_enabled = False
        self.edit_cell_enabled = False
        self.dragged_col = None
        self.visible_col_dividers = {}
        self.col_height_resize_bbox = ()
        self.cell_options = {}
        self.rsz_w = None
        self.rsz_h = None
        self.lines_start_at = 0
        self.currently_resizing_width = False
        self.currently_resizing_height = False
        self.ch_rc_popup_menu = None
        self.dropdown = DropdownStorage()
        self.text_editor = TextEditorStorage()

        self.disp_text = {}
        self.disp_high = {}
        self.disp_grid = {}
        self.disp_fill_sels = {}
        self.disp_resize_lines = {}
        self.disp_dropdown = {}
        self.disp_checkbox = {}
        self.disp_corners = set()
        self.hidd_text = {}
        self.hidd_high = {}
        self.hidd_grid = {}
        self.hidd_fill_sels = {}
        self.hidd_resize_lines = {}
        self.hidd_dropdown = {}
        self.hidd_checkbox = {}
        self.hidd_corners = set()

        self.align = kwargs["header_align"]
        self.basic_bindings()

    def event_generate(self, *args, **kwargs) -> None:
        for arg in args:
            if self.MT and arg in self.MT.event_linker:
                self.MT.event_linker[arg]()
            else:
                super().event_generate(*args, **kwargs)

    def basic_bindings(self, enable: bool = True) -> None:
        if enable:
            self.bind("<Motion>", self.mouse_motion)
            self.bind("<ButtonPress-1>", self.b1_press)
            self.bind("<B1-Motion>", self.b1_motion)
            self.bind("<ButtonRelease-1>", self.b1_release)
            self.bind("<Double-Button-1>", self.double_b1)
            for b in self.ops.rc_bindings:
                self.bind(b, self.rc)
            self.bind("<MouseWheel>", self.mousewheel)
            if USER_OS == "linux":
                self.bind("<Button-4>", self.mousewheel)
                self.bind("<Button-5>", self.mousewheel)
        else:
            self.unbind("<Motion>")
            self.unbind("<ButtonPress-1>")
            self.unbind("<B1-Motion>")
            self.unbind("<ButtonRelease-1>")
            self.unbind("<Double-Button-1>")
            for b in self.ops.rc_bindings:
                self.unbind(b)
            self.unbind("<MouseWheel>")
            if USER_OS == "linux":
                self.unbind("<Button-4>")
                self.unbind("<Button-5>")

    def mousewheel(self, event: tk.Event) -> None:
        self.MT.mousewheel(event)

    def set_height(self, new_height: int, set_TL: bool = False) -> bool:
        try:
            self.config(height=new_height)
        except Exception:
            return False
        if set_TL and self.TL is not None:
            self.TL.set_dimensions(new_h=new_height)
        if expanded := isinstance(self.current_height, int) and new_height > self.current_height:
            self.MT.recreate_all_selection_boxes()
        self.current_height = new_height
        return expanded

    def is_readonly(self, datacn: int) -> bool:
        return datacn in self.cell_options and "readonly" in self.cell_options[datacn]

    def rc(self, event: Any) -> None:
        self.mouseclick_outside_editor_or_dropdown_all_canvases(inside=True)
        self.focus_set()
        popup_menu = None
        if self.MT.identify_col(x=event.x, allow_end=False) is None:
            self.MT.deselect("all")
            c = len(self.MT.col_positions) - 1
            if self.MT.rc_popup_menus_enabled:
                popup_menu = self.MT.empty_rc_popup_menu
                build_empty_rc_menu(self.MT, popup_menu)
        elif self.col_selection_enabled and not self.currently_resizing_width and not self.currently_resizing_height:
            c = self.MT.identify_col(x=event.x)
            if c < len(self.MT.col_positions) - 1:
                if self.MT.col_selected(c):
                    if self.MT.rc_popup_menus_enabled:
                        popup_menu = self.ch_rc_popup_menu
                        build_header_rc_menu(self.MT, popup_menu, self.MT.selected)
                else:
                    if self.MT.single_selection_enabled and self.MT.rc_select_enabled:
                        self.select_col(c, redraw=True)
                    elif self.MT.toggle_selection_enabled and self.MT.rc_select_enabled:
                        self.toggle_select_col(c, redraw=True)
                    if self.MT.rc_popup_menus_enabled:
                        popup_menu = self.ch_rc_popup_menu
                        build_header_rc_menu(self.MT, popup_menu, self.MT.selected)
        try_binding(self.extra_rc_func, event)
        if popup_menu is not None:
            self.popup_menu_loc = c
            popup_menu.tk_popup(event.x_root, event.y_root)

    def ctrl_b1_press(self, event: Any) -> None:
        self.mouseclick_outside_editor_or_dropdown_all_canvases(inside=True)
        if (
            (self.drag_and_drop_enabled or self.col_selection_enabled)
            and self.MT.ctrl_select_enabled
            and self.rsz_h is None
            and self.rsz_w is None
        ):
            c = self.MT.identify_col(x=event.x)
            if c < len(self.MT.col_positions) - 1:
                c_selected = self.MT.col_selected(c)
                if not c_selected and self.col_selection_enabled:
                    self.being_drawn_item = True
                    self.being_drawn_item = self.add_selection(c, set_as_current=True, run_binding_func=False)
                    self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
                    sel_event = self.MT.get_select_event(being_drawn_item=self.being_drawn_item)
                    try_binding(self.ctrl_selection_binding_func, sel_event)
                    self.PAR.emit_event("<<SheetSelect>>", data=sel_event)
                elif c_selected:
                    self.MT.deselect(c=c)
        elif not self.MT.ctrl_select_enabled:
            self.b1_press(event)

    def ctrl_shift_b1_press(self, event: Any) -> None:
        self.mouseclick_outside_editor_or_dropdown_all_canvases(inside=True)
        x = event.x
        c = self.MT.identify_col(x=x)
        if (
            (self.drag_and_drop_enabled or self.col_selection_enabled)
            and self.MT.ctrl_select_enabled
            and self.rsz_h is None
            and self.rsz_w is None
        ):
            if c < len(self.MT.col_positions) - 1:
                c_selected = self.MT.col_selected(c)
                if not c_selected and self.col_selection_enabled:
                    if self.MT.selected and self.MT.selected.type_ == "columns":
                        self.being_drawn_item = self.MT.recreate_selection_box(
                            *self.get_shift_select_box(c, self.MT.selected.column),
                            fill_iid=self.MT.selected.fill_iid,
                        )
                    else:
                        self.being_drawn_item = self.add_selection(
                            c,
                            run_binding_func=False,
                            set_as_current=True,
                        )
                    self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
                    sel_event = self.MT.get_select_event(being_drawn_item=self.being_drawn_item)
                    try_binding(self.ctrl_selection_binding_func, sel_event)
                    self.PAR.emit_event("<<SheetSelect>>", data=sel_event)
                elif c_selected:
                    self.dragged_col = DraggedRowColumn(
                        dragged=c,
                        to_move=sorted(self.MT.get_selected_cols()),
                    )
        elif not self.MT.ctrl_select_enabled:
            self.shift_b1_press(event)

    def shift_b1_press(self, event: Any) -> None:
        self.mouseclick_outside_editor_or_dropdown_all_canvases(inside=True)
        c = self.MT.identify_col(x=event.x)
        if (
            (self.drag_and_drop_enabled or self.col_selection_enabled)
            and self.rsz_h is None
            and self.rsz_w is None
            and c < len(self.MT.col_positions) - 1
        ):
            c_selected = self.MT.col_selected(c)
            if not c_selected and self.col_selection_enabled:
                if self.MT.selected and self.MT.selected.type_ == "columns":
                    r_to_sel, c_to_sel = self.MT.selected.row, self.MT.selected.column
                    self.MT.deselect("all", redraw=False)
                    self.being_drawn_item = self.MT.create_selection_box(
                        *self.get_shift_select_box(c, c_to_sel), "columns"
                    )
                    self.MT.set_currently_selected(r_to_sel, c_to_sel, self.being_drawn_item)
                else:
                    self.being_drawn_item = self.select_col(c, run_binding_func=False)
                self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
                sel_event = self.MT.get_select_event(being_drawn_item=self.being_drawn_item)
                try_binding(self.shift_selection_binding_func, sel_event)
                self.PAR.emit_event("<<SheetSelect>>", data=sel_event)
            elif c_selected:
                self.dragged_col = DraggedRowColumn(
                    dragged=c,
                    to_move=sorted(self.MT.get_selected_cols()),
                )

    def get_shift_select_box(self, c: int, min_c: int) -> tuple[int, int, int, int, str]:
        if c >= min_c:
            return 0, min_c, len(self.MT.row_positions) - 1, c + 1
        elif c < min_c:
            return 0, c, len(self.MT.row_positions) - 1, min_c + 1

    def create_resize_line(
        self,
        x1: int,
        y1: int,
        x2: int,
        y2: int,
        width: int,
        fill: str,
        tag: str | tuple[str],
    ) -> None:
        if self.hidd_resize_lines:
            t, sh = self.hidd_resize_lines.popitem()
            self.coords(t, x1, y1, x2, y2)
            if sh:
                self.itemconfig(t, width=width, fill=fill, tag=tag)
            else:
                self.itemconfig(t, width=width, fill=fill, tag=tag, state="normal")
            self.lift(t)
        else:
            t = self.create_line(x1, y1, x2, y2, width=width, fill=fill, tag=tag)
        self.disp_resize_lines[t] = True

    def delete_resize_lines(self) -> None:
        self.hidd_resize_lines.update(self.disp_resize_lines)
        self.disp_resize_lines = {}
        for t, sh in self.hidd_resize_lines.items():
            if sh:
                self.itemconfig(t, tags=("",), state="hidden")
                self.hidd_resize_lines[t] = False

    def check_mouse_position_width_resizers(self, x: int, y: int) -> int | None:
        for c, (x1, y1, x2, y2) in self.visible_col_dividers.items():
            if x >= x1 and y >= y1 and x <= x2 and y <= y2:
                return c

    def mouse_motion(self, event: Any) -> None:
        if not self.currently_resizing_height and not self.currently_resizing_width:
            x = self.canvasx(event.x)
            y = self.canvasy(event.y)
            mouse_over_resize = False
            if self.width_resizing_enabled:
                c = self.check_mouse_position_width_resizers(x, y)
                if c is not None:
                    self.rsz_w, mouse_over_resize = c, True
                    if self.current_cursor != "sb_h_double_arrow":
                        self.config(cursor="sb_h_double_arrow")
                        self.current_cursor = "sb_h_double_arrow"
                else:
                    self.rsz_w = None
            if self.height_resizing_enabled and not mouse_over_resize:
                try:
                    x1, y1, x2, y2 = (
                        self.col_height_resize_bbox[0],
                        self.col_height_resize_bbox[1],
                        self.col_height_resize_bbox[2],
                        self.col_height_resize_bbox[3],
                    )
                    if x >= x1 and y >= y1 and x <= x2 and y <= y2:
                        self.rsz_h, mouse_over_resize = True, True
                        if self.current_cursor != "sb_v_double_arrow":
                            self.config(cursor="sb_v_double_arrow")
                            self.current_cursor = "sb_v_double_arrow"
                    else:
                        self.rsz_h = None
                except Exception:
                    self.rsz_h = None
            if not mouse_over_resize:
                if self.MT.col_selected(self.MT.identify_col(event, allow_end=False)):
                    if self.current_cursor != "hand2":
                        self.config(cursor="hand2")
                        self.current_cursor = "hand2"
                else:
                    if self.current_cursor != "":
                        self.config(cursor="")
                        self.current_cursor = ""
                    self.MT.reset_resize_vars()
        try_binding(self.extra_motion_func, event)

    def double_b1(self, event: Any) -> None:
        self.mouseclick_outside_editor_or_dropdown_all_canvases(inside=True)
        self.focus_set()
        if (
            self.double_click_resizing_enabled
            and self.width_resizing_enabled
            and self.rsz_w is not None
            and not self.currently_resizing_width
        ):
            col = self.rsz_w - 1
            old_width = self.MT.col_positions[self.rsz_w] - self.MT.col_positions[self.rsz_w - 1]
            new_width = self.set_col_width(col)
            self.MT.allow_auto_resize_columns = False
            self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
            if self.column_width_resize_func is not None and old_width != new_width:
                self.column_width_resize_func(
                    event_dict(
                        name="resize",
                        sheet=self.PAR.name,
                        resized_columns={col: {"old_size": old_width, "new_size": new_width}},
                    )
                )
        elif self.col_selection_enabled and self.rsz_h is None and self.rsz_w is None:
            c = self.MT.identify_col(x=event.x)
            if c < len(self.MT.col_positions) - 1:
                if self.MT.single_selection_enabled:
                    self.select_col(c, redraw=True)
                elif self.MT.toggle_selection_enabled:
                    self.toggle_select_col(c, redraw=True)
                datacn = c if self.MT.all_columns_displayed else self.MT.displayed_columns[c]
                if (
                    self.get_cell_kwargs(datacn, key="dropdown")
                    or self.get_cell_kwargs(datacn, key="checkbox")
                    or self.edit_cell_enabled
                ):
                    self.open_cell(event)
        self.rsz_w = None
        self.mouse_motion(event)
        try_binding(self.extra_double_b1_func, event)

    def b1_press(self, event: Any) -> None:
        self.MT.unbind("<MouseWheel>")
        self.focus_set()
        self.closed_dropdown = self.mouseclick_outside_editor_or_dropdown_all_canvases(inside=True)
        x = self.canvasx(event.x)
        y = self.canvasy(event.y)
        c = self.MT.identify_col(x=event.x)
        self.b1_pressed_loc = c
        if self.check_mouse_position_width_resizers(x, y) is None:
            self.rsz_w = None
        if self.width_resizing_enabled and self.rsz_w is not None:
            x1, y1, x2, y2 = self.MT.get_canvas_visible_area()
            self.currently_resizing_width = True
            x = self.MT.col_positions[self.rsz_w]
            line2x = self.MT.col_positions[self.rsz_w - 1]
            self.create_resize_line(
                x,
                0,
                x,
                self.current_height,
                width=1,
                fill=self.ops.resizing_line_fg,
                tag=("rw", "rwl"),
            )
            self.MT.create_resize_line(x, y1, x, y2, width=1, fill=self.ops.resizing_line_fg, tag=("rw", "rwl"))
            self.create_resize_line(
                line2x,
                0,
                line2x,
                self.current_height,
                width=1,
                fill=self.ops.resizing_line_fg,
                tag=("rw", "rwl2"),
            )
            self.MT.create_resize_line(
                line2x, y1, line2x, y2, width=1, fill=self.ops.resizing_line_fg, tag=("rw", "rwl2")
            )
        elif self.height_resizing_enabled and self.rsz_w is None and self.rsz_h is not None:
            self.currently_resizing_height = True
        elif self.MT.identify_col(x=event.x, allow_end=False) is None:
            self.MT.deselect("all")
        elif self.col_selection_enabled and self.rsz_w is None and self.rsz_h is None:
            if c < len(self.MT.col_positions) - 1:
                datacn = c if self.MT.all_columns_displayed else self.MT.displayed_columns[c]
                if (
                    self.MT.col_selected(c)
                    and not self.event_over_dropdown(c, datacn, event, x)
                    and not self.event_over_checkbox(c, datacn, event, x)
                ):
                    self.dragged_col = DraggedRowColumn(
                        dragged=c,
                        to_move=sorted(self.MT.get_selected_cols()),
                    )
                else:
                    if self.MT.single_selection_enabled:
                        self.being_drawn_item = True
                        self.being_drawn_item = self.select_col(c, redraw=True)
                    elif self.MT.toggle_selection_enabled:
                        self.toggle_select_col(c, redraw=True)
        try_binding(self.extra_b1_press_func, event)

    def b1_motion(self, event: Any) -> None:
        x1, y1, x2, y2 = self.MT.get_canvas_visible_area()
        if self.width_resizing_enabled and self.rsz_w is not None and self.currently_resizing_width:
            x = self.canvasx(event.x)
            size = x - self.MT.col_positions[self.rsz_w - 1]
            if size >= self.ops.min_column_width and size < self.ops.max_column_width:
                self.hide_resize_and_ctrl_lines(ctrl_lines=False)
                line2x = self.MT.col_positions[self.rsz_w - 1]
                self.create_resize_line(
                    x,
                    0,
                    x,
                    self.current_height,
                    width=1,
                    fill=self.ops.resizing_line_fg,
                    tag=("rw", "rwl"),
                )
                self.MT.create_resize_line(x, y1, x, y2, width=1, fill=self.ops.resizing_line_fg, tag=("rw", "rwl"))
                self.create_resize_line(
                    line2x,
                    0,
                    line2x,
                    self.current_height,
                    width=1,
                    fill=self.ops.resizing_line_fg,
                    tag=("rw", "rwl2"),
                )
                self.MT.create_resize_line(
                    line2x,
                    y1,
                    line2x,
                    y2,
                    width=1,
                    fill=self.ops.resizing_line_fg,
                    tag=("rw", "rwl2"),
                )
                self.drag_width_resize()
        elif self.height_resizing_enabled and self.rsz_h is not None and self.currently_resizing_height:
            evy = event.y
            if evy > self.current_height:
                if evy > self.ops.max_header_height:
                    evy = int(self.ops.max_header_height)
                self.drag_height_resize(evy)
            else:
                if evy < self.MT.min_header_height:
                    evy = int(self.MT.min_header_height)
                self.drag_height_resize(evy)
        elif (
            self.drag_and_drop_enabled
            and self.col_selection_enabled
            and self.MT.anything_selected(exclude_cells=True, exclude_rows=True)
            and self.rsz_h is None
            and self.rsz_w is None
            and self.dragged_col is not None
        ):
            x = self.canvasx(event.x)
            if x > 0:
                self.show_drag_and_drop_indicators(
                    self.drag_and_drop_motion(event),
                    y1,
                    y2,
                    self.dragged_col.to_move,
                )
        elif (
            self.MT.drag_selection_enabled and self.col_selection_enabled and self.rsz_h is None and self.rsz_w is None
        ):
            need_redraw = False
            end_col = self.MT.identify_col(x=event.x)
            if end_col < len(self.MT.col_positions) - 1 and self.MT.selected:
                if self.MT.selected.type_ == "columns":
                    box = self.get_b1_motion_box(self.MT.selected.column, end_col)
                    if (
                        box is not None
                        and self.being_drawn_item is not None
                        and self.MT.coords_and_type(self.being_drawn_item) != box
                    ):
                        if box[3] - box[1] != 1:
                            self.being_drawn_item = self.MT.recreate_selection_box(
                                *box[:-1],
                                fill_iid=self.MT.selected.fill_iid,
                            )
                        else:
                            self.being_drawn_item = self.select_col(self.MT.selected.column, run_binding_func=False)
                        need_redraw = True
                        sel_event = self.MT.get_select_event(being_drawn_item=self.being_drawn_item)
                        try_binding(self.drag_selection_binding_func, sel_event)
                        self.PAR.emit_event("<<SheetSelect>>", data=sel_event)
                if self.scroll_if_event_offscreen(event):
                    need_redraw = True
            if need_redraw:
                self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=False)
        try_binding(self.extra_b1_motion_func, event)

    def drag_height_resize(self, height: int) -> None:
        if self.set_height(height, set_TL=True):
            self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=False, redraw_table=False)

    def get_b1_motion_box(self, start_col: int, end_col: int) -> tuple[int, int, int, int, Literal["columns"]]:
        if end_col >= start_col:
            return 0, start_col, len(self.MT.row_positions) - 1, end_col + 1, "columns"
        elif end_col < start_col:
            return 0, end_col, len(self.MT.row_positions) - 1, start_col + 1, "columns"

    def ctrl_b1_motion(self, event: Any) -> None:
        x1, y1, x2, y2 = self.MT.get_canvas_visible_area()
        if (
            self.drag_and_drop_enabled
            and self.col_selection_enabled
            and self.MT.anything_selected(exclude_cells=True, exclude_rows=True)
            and self.rsz_h is None
            and self.rsz_w is None
            and self.dragged_col is not None
        ):
            x = self.canvasx(event.x)
            if x > 0:
                self.show_drag_and_drop_indicators(
                    self.drag_and_drop_motion(event),
                    y1,
                    y2,
                    self.dragged_col.to_move,
                )
        elif (
            self.MT.ctrl_select_enabled
            and self.MT.drag_selection_enabled
            and self.col_selection_enabled
            and self.rsz_h is None
            and self.rsz_w is None
        ):
            need_redraw = False
            end_col = self.MT.identify_col(x=event.x)
            if end_col < len(self.MT.col_positions) - 1 and self.MT.selected:
                if self.MT.selected.type_ == "columns":
                    box = self.get_b1_motion_box(self.MT.selected.column, end_col)
                    if (
                        box is not None
                        and self.being_drawn_item is not None
                        and self.MT.coords_and_type(self.being_drawn_item) != box
                    ):
                        if box[3] - box[1] != 1:
                            self.being_drawn_item = self.MT.recreate_selection_box(
                                *box[:-1],
                                self.MT.selected.fill_iid,
                            )
                        else:
                            self.MT.hide_selection_box(self.MT.selected.fill_iid)
                            self.being_drawn_item = self.add_selection(box[1], run_binding_func=False)
                        need_redraw = True
                        sel_event = self.MT.get_select_event(being_drawn_item=self.being_drawn_item)
                        try_binding(self.drag_selection_binding_func, sel_event)
                        self.PAR.emit_event("<<SheetSelect>>", data=sel_event)
                if self.scroll_if_event_offscreen(event):
                    need_redraw = True
            if need_redraw:
                self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=False)
        elif not self.MT.ctrl_select_enabled:
            self.b1_motion(event)

    def drag_and_drop_motion(self, event: Any) -> float:
        x = event.x
        wend = self.winfo_width()
        xcheck = self.xview()
        if x >= wend - 0 and len(xcheck) > 1 and xcheck[1] < 1:
            if x >= wend + 15:
                self.MT.xview_scroll(2, "units")
                self.xview_scroll(2, "units")
            else:
                self.MT.xview_scroll(1, "units")
                self.xview_scroll(1, "units")
            self.fix_xview()
            self.MT.x_move_synced_scrolls("moveto", self.MT.xview()[0])
            self.MT.main_table_redraw_grid_and_text(redraw_header=True)
        elif x <= 0 and len(xcheck) > 1 and xcheck[0] > 0:
            if x >= -15:
                self.MT.xview_scroll(-1, "units")
                self.xview_scroll(-1, "units")
            else:
                self.MT.xview_scroll(-2, "units")
                self.xview_scroll(-2, "units")
            self.fix_xview()
            self.MT.x_move_synced_scrolls("moveto", self.MT.xview()[0])
            self.MT.main_table_redraw_grid_and_text(redraw_header=True)
        col = self.MT.identify_col(x=x)
        if col == len(self.MT.col_positions) - 1:
            col -= 1
        if col >= self.dragged_col.to_move[0] and col <= self.dragged_col.to_move[-1]:
            if is_contiguous(self.dragged_col.to_move):
                return self.MT.col_positions[self.dragged_col.to_move[0]]
            return self.MT.col_positions[col]
        elif col > self.dragged_col.to_move[-1]:
            return self.MT.col_positions[col + 1]
        return self.MT.col_positions[col]

    def show_drag_and_drop_indicators(
        self,
        xpos: float,
        y1: float,
        y2: float,
        cols: Sequence[int],
    ) -> None:
        self.hide_resize_and_ctrl_lines()
        self.create_resize_line(
            xpos,
            0,
            xpos,
            self.current_height,
            width=3,
            fill=self.ops.drag_and_drop_bg,
            tag="move_columns",
        )
        self.MT.create_resize_line(xpos, y1, xpos, y2, width=3, fill=self.ops.drag_and_drop_bg, tag="move_columns")
        for boxst, boxend in consecutive_ranges(cols):
            self.MT.show_ctrl_outline(
                start_cell=(boxst, 0),
                end_cell=(boxend, len(self.MT.row_positions) - 1),
                dash=(),
                outline=self.ops.drag_and_drop_bg,
                delete_on_timer=False,
            )

    def hide_resize_and_ctrl_lines(self, ctrl_lines: bool = True) -> None:
        self.delete_resize_lines()
        self.MT.delete_resize_lines()
        if ctrl_lines:
            self.MT.delete_ctrl_outlines()

    def scroll_if_event_offscreen(self, event: Any) -> bool:
        xcheck = self.xview()
        need_redraw = False
        if event.x > self.winfo_width() and len(xcheck) > 1 and xcheck[1] < 1:
            try:
                self.MT.xview_scroll(1, "units")
                self.xview_scroll(1, "units")
            except Exception:
                pass
            self.fix_xview()
            self.MT.x_move_synced_scrolls("moveto", self.MT.xview()[0])
            need_redraw = True
        elif event.x < 0 and self.canvasx(self.winfo_width()) > 0 and xcheck and xcheck[0] > 0:
            try:
                self.xview_scroll(-1, "units")
                self.MT.xview_scroll(-1, "units")
            except Exception:
                pass
            self.fix_xview()
            self.MT.x_move_synced_scrolls("moveto", self.MT.xview()[0])
            need_redraw = True
        return need_redraw

    def fix_xview(self) -> None:
        xcheck = self.xview()
        if xcheck and xcheck[0] < 0:
            self.MT.set_xviews("moveto", 0)
        elif len(xcheck) > 1 and xcheck[1] > 1:
            self.MT.set_xviews("moveto", 1)

    def event_over_dropdown(self, c: int, datacn: int, event: Any, canvasx: float) -> bool:
        return (
            event.y < self.MT.header_txt_height + 5
            and self.get_cell_kwargs(datacn, key="dropdown")
            and canvasx < self.MT.col_positions[c + 1]
            and canvasx > self.MT.col_positions[c + 1] - self.MT.header_txt_height - 4
        )

    def event_over_checkbox(self, c: int, datacn: int, event: Any, canvasx: float) -> bool:
        return (
            event.y < self.MT.header_txt_height + 5
            and self.get_cell_kwargs(datacn, key="checkbox")
            and canvasx < self.MT.col_positions[c] + self.MT.header_txt_height + 4
        )

    def drag_width_resize(self) -> None:
        new_col_pos = int(self.coords("rwl")[0])
        old_width = self.MT.col_positions[self.rsz_w] - self.MT.col_positions[self.rsz_w - 1]
        size = new_col_pos - self.MT.col_positions[self.rsz_w - 1]
        if size < self.ops.min_column_width:
            new_col_pos = ceil(self.MT.col_positions[self.rsz_w - 1] + self.ops.min_column_width)
        elif size > self.ops.max_column_width:
            new_col_pos = int(self.MT.col_positions[self.rsz_w - 1] + self.ops.max_column_width)
        increment = new_col_pos - self.MT.col_positions[self.rsz_w]
        self.MT.col_positions[self.rsz_w + 1 :] = [
            e + increment for e in islice(self.MT.col_positions, self.rsz_w + 1, None)
        ]
        self.MT.col_positions[self.rsz_w] = new_col_pos
        new_width = self.MT.col_positions[self.rsz_w] - self.MT.col_positions[self.rsz_w - 1]
        self.MT.allow_auto_resize_columns = False
        self.MT.recreate_all_selection_boxes()
        self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True, set_scrollregion=False)
        if self.column_width_resize_func is not None and old_width != new_width:
            self.column_width_resize_func(
                event_dict(
                    name="resize",
                    sheet=self.PAR.name,
                    resized_columns={self.rsz_w - 1: {"old_size": old_width, "new_size": new_width}},
                )
            )

    def b1_release(self, event: Any) -> None:
        to_hide = self.being_drawn_item
        if self.being_drawn_item is not None and (to_sel := self.MT.coords_and_type(self.being_drawn_item)):
            r_to_sel, c_to_sel = self.MT.selected.row, self.MT.selected.column
            self.being_drawn_item = None
            self.MT.set_currently_selected(
                r_to_sel,
                c_to_sel,
                item=self.MT.create_selection_box(*to_sel, set_current=False),
                run_binding=False,
            )
        self.MT.hide_selection_box(to_hide)
        self.MT.bind("<MouseWheel>", self.MT.mousewheel)
        if self.width_resizing_enabled and self.rsz_w is not None and self.currently_resizing_width:
            self.drag_width_resize()
            self.hide_resize_and_ctrl_lines(ctrl_lines=False)
            self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
        elif (
            self.drag_and_drop_enabled
            and self.col_selection_enabled
            and self.MT.anything_selected(exclude_cells=True, exclude_rows=True)
            and self.rsz_h is None
            and self.rsz_w is None
            and self.dragged_col is not None
            and self.find_withtag("move_columns")
        ):
            self.hide_resize_and_ctrl_lines()
            c = self.MT.identify_col(x=event.x)
            totalcols = len(self.dragged_col.to_move)
            if (
                c is not None
                and totalcols != len(self.MT.col_positions) - 1
                and not (
                    c >= self.dragged_col.to_move[0]
                    and c <= self.dragged_col.to_move[-1]
                    and is_contiguous(self.dragged_col.to_move)
                )
            ):
                if c > self.dragged_col.to_move[-1]:
                    c += 1
                if c > len(self.MT.col_positions) - 1:
                    c = len(self.MT.col_positions) - 1
                event_data = self.MT.new_event_dict("move_columns", state=True)
                event_data["value"] = c
                if try_binding(self.ch_extra_begin_drag_drop_func, event_data, "begin_move_columns"):
                    data_new_idxs, disp_new_idxs, event_data = self.MT.move_columns_adjust_options_dict(
                        *self.MT.get_args_for_move_columns(
                            move_to=c,
                            to_move=self.dragged_col.to_move,
                        ),
                        move_data=self.ops.column_drag_and_drop_perform,
                        move_widths=self.ops.column_drag_and_drop_perform,
                        event_data=event_data,
                    )
                    event_data["moved"]["columns"] = {
                        "data": data_new_idxs,
                        "displayed": disp_new_idxs,
                    }
                    if self.MT.undo_enabled:
                        self.MT.undo_stack.append(stored_event_dict(event_data))
                    self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
                    try_binding(self.ch_extra_end_drag_drop_func, event_data, "end_move_columns")
                    self.MT.sheet_modified(event_data)
        elif self.b1_pressed_loc is not None and self.rsz_w is None and self.rsz_h is None:
            c = self.MT.identify_col(x=event.x)
            if (
                c is not None
                and c < len(self.MT.col_positions) - 1
                and c == self.b1_pressed_loc
                and self.b1_pressed_loc != self.closed_dropdown
            ):
                datacn = c if self.MT.all_columns_displayed else self.MT.displayed_columns[c]
                canvasx = self.canvasx(event.x)
                if self.event_over_dropdown(
                    c,
                    datacn,
                    event,
                    canvasx,
                ) or self.event_over_checkbox(
                    c,
                    datacn,
                    event,
                    canvasx,
                ):
                    self.open_cell(event)
            else:
                self.mouseclick_outside_editor_or_dropdown_all_canvases(inside=True)
            self.b1_pressed_loc = None
            self.closed_dropdown = None
        self.dragged_col = None
        self.currently_resizing_width = False
        self.currently_resizing_height = False
        self.rsz_w = None
        self.rsz_h = None
        self.mouse_motion(event)
        try_binding(self.extra_b1_release_func, event)

    def _sort_columns(
        self,
        event: tk.Event | None = None,
        columns: Iterator[int] | None = None,
        reverse: bool = False,
        validation: bool = True,
        key: Callable | None = None,
        undo: bool = True,
    ) -> EventDataDict:
        if columns is None:
            columns = self.MT.get_selected_cols()
        if not columns:
            columns = list(range(0, len(self.MT.col_positions) - 1))
        event_data = self.MT.new_event_dict("edit_table")
        try_binding(self.MT.extra_begin_sort_cells_func, event_data)
        if key is None:
            key = self.ops.sort_key
        for c in columns:
            datacn = self.MT.datacn(c)
            for r, val in enumerate(
                sort_column(
                    (self.MT.get_cell_data(row, datacn) for row in range(len(self.MT.data))),
                    reverse=reverse,
                    key=key,
                )
            ):
                if (
                    not self.MT.edit_validation_func
                    or not validation
                    or (
                        self.MT.edit_validation_func
                        and (val := self.MT.edit_validation_func(mod_event_val(event_data, val, (r, datacn))))
                        is not None
                    )
                ):
                    event_data = self.MT.event_data_set_cell(
                        datarn=r,
                        datacn=datacn,
                        value=val,
                        event_data=event_data,
                    )
        event_data = self.MT.bulk_edit_validation(event_data)
        if event_data["cells"]["table"]:
            if undo and self.MT.undo_enabled:
                self.MT.undo_stack.append(stored_event_dict(event_data))
            try_binding(self.MT.extra_end_sort_cells_func, event_data, "end_edit_table")
            self.MT.sheet_modified(event_data)
            self.PAR.emit_event("<<SheetModified>>", event_data)
            self.MT.refresh()
        return event_data

    def _sort_rows_by_column(
        self,
        event: tk.Event | None = None,
        column: int | None = None,
        reverse: bool = False,
        key: Callable | None = None,
        undo: bool = True,
    ) -> EventDataDict:
        event_data = self.MT.new_event_dict("move_rows", state=True)
        if not self.MT.data:
            return event_data
        if column is None:
            if not self.MT.selected:
                return event_data
            column = self.MT.datacn(self.MT.selected.column)
        if try_binding(self.ch_extra_begin_sort_rows_func, event_data, "begin_move_rows"):
            if key is None:
                key = self.ops.sort_key
            disp_new_idxs, disp_row_ctr = {}, 0
            if self.ops.treeview:
                new_nodes_order, data_new_idxs = sort_tree_rows_by_column(
                    data=self.MT.data,
                    column=column,
                    index=self.MT._row_index,
                    rns=self.RI.rns,
                    reverse=reverse,
                    key=key,
                )
                for node in new_nodes_order:
                    if (idx := try_b_index(self.MT.displayed_rows, self.RI.rns[node.iid])) is not None:
                        disp_new_idxs[idx] = disp_row_ctr
                        disp_row_ctr += 1
            else:
                new_rows_order, data_new_idxs = sort_rows_by_column(
                    self.MT.data,
                    column=column,
                    reverse=reverse,
                    key=key,
                )
                if self.MT.all_rows_displayed:
                    disp_new_idxs = data_new_idxs
                else:
                    for old_idx, _ in new_rows_order:
                        if (idx := try_b_index(self.MT.displayed_rows, old_idx)) is not None:
                            disp_new_idxs[idx] = disp_row_ctr
                            disp_row_ctr += 1
            if self.ops.treeview:
                data_new_idxs, disp_new_idxs, event_data = self.MT.move_rows_adjust_options_dict(
                    data_new_idxs=data_new_idxs,
                    data_old_idxs=dict(zip(data_new_idxs.values(), data_new_idxs)),
                    totalrows=None,
                    disp_new_idxs=disp_new_idxs,
                    move_data=True,
                    create_selections=False,
                    event_data=event_data,
                    manage_tree=False,
                )
            else:
                data_new_idxs, disp_new_idxs, _ = self.PAR.mapping_move_rows(
                    data_new_idxs=data_new_idxs,
                    disp_new_idxs=disp_new_idxs,
                    move_data=True,
                    create_selections=False,
                    undo=False,
                    emit_event=False,
                    redraw=True,
                )
            event_data["moved"]["rows"] = {
                "data": data_new_idxs,
                "displayed": disp_new_idxs,
            }
            if undo and self.MT.undo_enabled:
                self.MT.undo_stack.append(stored_event_dict(event_data))
            try_binding(self.ch_extra_end_sort_rows_func, event_data, "end_move_rows")
            self.MT.sheet_modified(event_data)
            self.PAR.emit_event("<<SheetModified>>", event_data)
            self.MT.refresh()
        return event_data

    def toggle_select_col(
        self,
        column: int,
        add_selection: bool = True,
        redraw: bool = True,
        run_binding_func: bool = True,
        set_as_current: bool = True,
        ext: bool = False,
    ) -> int:
        if add_selection:
            if self.MT.col_selected(column):
                fill_iid = self.MT.deselect(c=column, redraw=redraw)
            else:
                fill_iid = self.add_selection(
                    c=column,
                    redraw=redraw,
                    run_binding_func=run_binding_func,
                    set_as_current=set_as_current,
                    ext=ext,
                )
        else:
            if self.MT.col_selected(column):
                fill_iid = self.MT.deselect(c=column, redraw=redraw)
            else:
                fill_iid = self.select_col(column, redraw=redraw, ext=ext)
        return fill_iid

    def select_col(
        self,
        c: int | Iterator[int],
        redraw: bool = False,
        run_binding_func: bool = True,
        ext: bool = False,
    ) -> int | list[int]:
        boxes_to_hide = tuple(self.MT.selection_boxes)
        fill_iids = [
            self.MT.create_selection_box(
                0,
                start,
                len(self.MT.row_positions) - 1,
                end,
                "columns",
                set_current=True,
                ext=ext,
            )
            for start, end in consecutive_ranges(int_x_tuple(c))
        ]
        for iid in boxes_to_hide:
            self.MT.hide_selection_box(iid)
        if redraw:
            self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
        if run_binding_func:
            self.MT.run_selection_binding("columns")
        return fill_iids[0] if len(fill_iids) == 1 else fill_iids

    def add_selection(
        self,
        c: int,
        redraw: bool = False,
        run_binding_func: bool = True,
        set_as_current: bool = True,
        ext: bool = False,
    ) -> int:
        box = (0, c, len(self.MT.row_positions) - 1, c + 1, "columns")
        fill_iid = self.MT.create_selection_box(*box, set_current=set_as_current, ext=ext)
        if redraw:
            self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
        if run_binding_func:
            self.MT.run_selection_binding("columns")
        return fill_iid

    def get_cell_dimensions(self, datacn: int) -> tuple[int, int]:
        txt = self.cell_str(datacn, fix=False)
        if txt:
            self.MT.txt_measure_canvas.itemconfig(
                self.MT.txt_measure_canvas_text,
                text=txt,
                font=self.ops.header_font,
            )
            b = self.MT.txt_measure_canvas.bbox(self.MT.txt_measure_canvas_text)
            w = b[2] - b[0] + 7
            h = b[3] - b[1] + 5
        else:
            w = self.ops.min_column_width
            h = self.MT.min_header_height
        if datacn in self.cell_options and (
            self.get_cell_kwargs(datacn, key="dropdown") or self.get_cell_kwargs(datacn, key="checkbox")
        ):
            return w + self.MT.header_txt_height, h
        return w, h

    def set_height_of_header_to_text(
        self,
        text: None | str = None,
        only_if_too_small: bool = False,
    ) -> int:
        h = self.MT.min_header_height
        if (text is None and not self.MT._headers and isinstance(self.MT._headers, list)) or (
            isinstance(self.MT._headers, int) and self.MT._headers >= len(self.MT.data)
        ):
            return h
        self.fix_header()
        qconf = self.MT.txt_measure_canvas.itemconfig
        qbbox = self.MT.txt_measure_canvas.bbox
        qtxtm = self.MT.txt_measure_canvas_text
        qfont = self.ops.header_font
        default_header_height = self.MT.get_default_header_height()
        if text is not None and text:
            qconf(qtxtm, text=text, font=qfont)
            b = qbbox(qtxtm)
            if (th := b[3] - b[1] + 5) > h:
                h = th
        elif text is None:
            if self.MT.all_columns_displayed:
                if isinstance(self.MT._headers, list):
                    iterable = range(len(self.MT._headers))
                else:
                    iterable = range(len(self.MT.data[self.MT._headers]))
            else:
                iterable = self.MT.displayed_columns
            if (
                isinstance(self.MT._headers, list)
                and (th := max(map(itemgetter(0), map(self.get_cell_dimensions, iterable)), default=h)) > h
            ):
                h = th
            elif isinstance(self.MT._headers, int):
                datarn = self.MT._headers
                for datacn in iterable:
                    if txt := self.MT.cell_str(datarn, datacn, get_displayed=True):
                        qconf(qtxtm, text=txt, font=qfont)
                        b = qbbox(qtxtm)
                        th = b[3] - b[1] + 5
                    else:
                        th = default_header_height
                    if th > h:
                        h = th
        space_bot = self.MT.get_space_bot(0)
        if h > space_bot and space_bot > self.MT.min_header_height:
            h = space_bot
        if h < self.MT.min_header_height:
            h = int(self.MT.min_header_height)
        elif h > self.ops.max_header_height:
            h = int(self.ops.max_header_height)
        if not only_if_too_small or (only_if_too_small and h > self.current_height):
            self.set_height(h, set_TL=True)
            self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
        return h

    def get_col_text_width(
        self,
        col: int,
        visible_only: bool = False,
        only_if_too_small: bool = False,
    ) -> int:
        self.fix_header()
        w = self.ops.min_column_width
        datacn = col if self.MT.all_columns_displayed else self.MT.displayed_columns[col]
        # header
        hw, hh_ = self.get_cell_dimensions(datacn)
        # table
        if self.MT.data:
            if self.MT.all_rows_displayed:
                iterable = range(*self.MT.visible_text_rows) if visible_only else range(0, len(self.MT.data))
            else:
                if visible_only:
                    start_row, end_row = self.MT.visible_text_rows
                else:
                    start_row, end_row = 0, len(self.MT.displayed_rows)
                iterable = self.MT.displayed_rows[start_row:end_row]
            qconf = self.MT.txt_measure_canvas.itemconfig
            qbbox = self.MT.txt_measure_canvas.bbox
            qtxtm = self.MT.txt_measure_canvas_text
            qtxth = self.MT.table_txt_height
            qfont = self.ops.table_font
            for datarn in iterable:
                if txt := self.MT.cell_str(datarn, datacn, get_displayed=True):
                    qconf(qtxtm, text=txt, font=qfont)
                    b = qbbox(qtxtm)
                    if (
                        (
                            self.MT.get_cell_kwargs(datarn, datacn, key="dropdown")
                            or self.MT.get_cell_kwargs(datarn, datacn, key="checkbox")
                        )
                        and (tw := b[2] - b[0] + qtxth + 7) > w
                        or (tw := b[2] - b[0] + 7) > w
                    ):
                        w = tw
        if hw > w:
            w = hw
        if only_if_too_small and w < self.MT.col_positions[col + 1] - self.MT.col_positions[col]:
            w = self.MT.col_positions[col + 1] - self.MT.col_positions[col]
        if w <= self.ops.min_column_width:
            w = self.ops.min_column_width
        elif w > self.ops.max_column_width:
            w = int(self.ops.max_column_width)
        return w

    def set_col_width(
        self,
        col: int,
        width: None | int = None,
        only_if_too_small: bool = False,
        visible_only: bool = False,
        recreate: bool = True,
    ) -> int:
        if width is None:
            width = self.get_col_text_width(col=col, visible_only=visible_only)
        if width <= self.ops.min_column_width:
            width = self.ops.min_column_width
        elif width > self.ops.max_column_width:
            width = int(self.ops.max_column_width)
        if only_if_too_small and width <= self.MT.col_positions[col + 1] - self.MT.col_positions[col]:
            return self.MT.col_positions[col + 1] - self.MT.col_positions[col]
        new_col_pos = self.MT.col_positions[col] + width
        increment = new_col_pos - self.MT.col_positions[col + 1]
        self.MT.col_positions[col + 2 :] = [
            e + increment for e in islice(self.MT.col_positions, col + 2, len(self.MT.col_positions))
        ]
        self.MT.col_positions[col + 1] = new_col_pos
        if recreate:
            self.MT.recreate_all_selection_boxes()
        return width

    def set_width_of_all_cols(
        self,
        width: None | int = None,
        only_if_too_small: bool = False,
        recreate: bool = True,
    ) -> None:
        if width is None:
            if self.MT.all_columns_displayed:
                iterable = range(self.MT.total_data_cols())
            else:
                iterable = range(len(self.MT.displayed_columns))
            self.MT.set_col_positions(
                itr=(self.get_col_text_width(cn, only_if_too_small=only_if_too_small) for cn in iterable)
            )
        elif width is not None:
            if self.MT.all_columns_displayed:
                self.MT.set_col_positions(itr=repeat(width, self.MT.total_data_cols()))
            else:
                self.MT.set_col_positions(itr=repeat(width, len(self.MT.displayed_columns)))
        if recreate:
            self.MT.recreate_all_selection_boxes()

    def redraw_highlight_get_text_fg(
        self,
        fc: float,
        sc: float,
        c: int,
        sel_cells_bg: str,
        sel_cols_bg: str,
        selections: dict,
        datacn: int,
        has_dd: bool,
        tags: str | tuple[str],
    ) -> tuple[str, bool]:
        redrawn = False
        kwargs = self.get_cell_kwargs(datacn, key="highlight")
        if kwargs:
            high_bg = kwargs[0]
            if high_bg and not high_bg.startswith("#"):
                high_bg = color_map[high_bg]
            if "columns" in selections and c in selections["columns"]:
                txtfg = (
                    self.ops.header_selected_columns_fg
                    if kwargs[1] is None or self.ops.display_selected_fg_over_highlights
                    else kwargs[1]
                )
                redrawn = self.redraw_highlight(
                    fc + 1,
                    0,
                    sc,
                    self.current_height - 1,
                    fill=self.ops.header_selected_columns_bg
                    if high_bg is None
                    else (
                        f"#{int((int(high_bg[1:3], 16) + int(sel_cols_bg[1:3], 16)) / 2):02X}"
                        + f"{int((int(high_bg[3:5], 16) + int(sel_cols_bg[3:5], 16)) / 2):02X}"
                        + f"{int((int(high_bg[5:], 16) + int(sel_cols_bg[5:], 16)) / 2):02X}"
                    ),
                    outline=self.ops.header_fg if has_dd and self.ops.show_dropdown_borders else "",
                    tags=tags,
                )
            elif "cells" in selections and c in selections["cells"]:
                txtfg = (
                    self.ops.header_selected_cells_fg
                    if kwargs[1] is None or self.ops.display_selected_fg_over_highlights
                    else kwargs[1]
                )
                redrawn = self.redraw_highlight(
                    fc + 1,
                    0,
                    sc,
                    self.current_height - 1,
                    fill=self.ops.header_selected_cells_bg
                    if high_bg is None
                    else (
                        f"#{int((int(high_bg[1:3], 16) + int(sel_cells_bg[1:3], 16)) / 2):02X}"
                        + f"{int((int(high_bg[3:5], 16) + int(sel_cells_bg[3:5], 16)) / 2):02X}"
                        + f"{int((int(high_bg[5:], 16) + int(sel_cells_bg[5:], 16)) / 2):02X}"
                    ),
                    outline=self.ops.header_fg if has_dd and self.ops.show_dropdown_borders else "",
                    tags=tags,
                )
            else:
                txtfg = self.ops.header_fg if kwargs[1] is None else kwargs[1]
                if high_bg:
                    redrawn = self.redraw_highlight(
                        fc + 1,
                        0,
                        sc,
                        self.current_height - 1,
                        fill=high_bg,
                        outline=self.ops.header_fg if has_dd and self.ops.show_dropdown_borders else "",
                        tags=tags,
                    )
        elif not kwargs:
            if "columns" in selections and c in selections["columns"]:
                txtfg = self.ops.header_selected_columns_fg
                redrawn = self.redraw_highlight(
                    fc + 1,
                    0,
                    sc,
                    self.current_height - 1,
                    fill=self.ops.header_selected_columns_bg,
                    outline=self.ops.header_fg if has_dd and self.ops.show_dropdown_borders else "",
                    tags=tags,
                )
            elif "cells" in selections and c in selections["cells"]:
                txtfg = self.ops.header_selected_cells_fg
                redrawn = self.redraw_highlight(
                    fc + 1,
                    0,
                    sc,
                    self.current_height - 1,
                    fill=self.ops.header_selected_cells_bg,
                    outline=self.ops.header_fg if has_dd and self.ops.show_dropdown_borders else "",
                    tags=tags,
                )
            else:
                txtfg = self.ops.header_fg
                redrawn = self.redraw_highlight(
                    fc + 1,
                    0,
                    sc,
                    self.current_height - 1,
                    fill="",
                    outline=self.ops.header_fg if has_dd and self.ops.show_dropdown_borders else "",
                    tags=tags,
                )
        return txtfg, redrawn

    def redraw_highlight(
        self,
        x1: float,
        y1: float,
        x2: float,
        y2: float,
        fill: str,
        outline: str,
        tags: str | tuple[str],
    ) -> bool:
        if not self.ops.show_vertical_grid:
            x2 += 1
        if self.hidd_high:
            iid, showing = self.hidd_high.popitem()
            self.coords(iid, x1, y1, x2, y2)
            if showing:
                self.itemconfig(iid, fill=fill, outline=outline, tags=tags)
            else:
                self.itemconfig(iid, fill=fill, outline=outline, state="normal", tags=tags)
        else:
            iid = self.create_rectangle(x1, y1, x2, y2, fill=fill, outline=outline, tags=tags)
        self.disp_high[iid] = True
        return True

    def redraw_gridline(
        self,
        points: Sequence[float],
        fill: str,
        width: int,
    ) -> None:
        if self.hidd_grid:
            t, sh = self.hidd_grid.popitem()
            self.coords(t, points)
            if sh:
                self.itemconfig(t, fill=fill, width=width)
            else:
                self.itemconfig(t, fill=fill, width=width, state="normal")
            self.disp_grid[t] = True
        else:
            self.disp_grid[self.create_line(points, fill=fill, width=width)] = True

    def redraw_dropdown(
        self,
        x1: float,
        y1: float,
        x2: float,
        y2: float,
        fill: str,
        outline: str,
        draw_outline: bool = True,
        draw_arrow: bool = True,
        open_: bool = False,
    ) -> None:
        # if draw_outline and self.ops.show_dropdown_borders:
        #     self.redraw_highlight(x1 + 1, y1 + 1, x2, y2, fill="", outline=self.ops.header_fg)
        if draw_arrow:
            mod = (self.MT.header_txt_height - 1) if self.MT.header_txt_height % 2 else self.MT.header_txt_height
            small_mod = int(mod / 5)
            mid_y = int(self.MT.min_header_height / 2)
            if open_:
                # up arrow
                points = (
                    x2 - 4 - small_mod - small_mod - small_mod - small_mod,
                    y1 + mid_y + small_mod,
                    x2 - 4 - small_mod - small_mod,
                    y1 + mid_y - small_mod,
                    x2 - 4,
                    y1 + mid_y + small_mod,
                )
            else:
                # down arrow
                points = (
                    x2 - 4 - small_mod - small_mod - small_mod - small_mod,
                    y1 + mid_y - small_mod,
                    x2 - 4 - small_mod - small_mod,
                    y1 + mid_y + small_mod,
                    x2 - 4,
                    y1 + mid_y - small_mod,
                )
            if self.hidd_dropdown:
                t, sh = self.hidd_dropdown.popitem()
                self.coords(t, points)
                if sh:
                    self.itemconfig(t, fill=fill)
                else:
                    self.itemconfig(t, fill=fill, state="normal")
            else:
                t = self.create_line(points, fill=fill, width=2, capstyle=tk.ROUND, joinstyle=tk.BEVEL, tag="lift")
            self.disp_dropdown[t] = True

    def redraw_checkbox(
        self,
        x1: float,
        y1: float,
        x2: float,
        y2: float,
        fill: str,
        outline: str,
        draw_check: bool = False,
    ) -> None:
        points = rounded_box_coords(x1, y1, x2, y2)
        if self.hidd_checkbox:
            t, sh = self.hidd_checkbox.popitem()
            self.coords(t, points)
            if sh:
                self.itemconfig(t, fill=outline, outline=fill)
            else:
                self.itemconfig(t, fill=outline, outline=fill, state="normal")
        else:
            t = self.create_polygon(points, fill=outline, outline=fill, smooth=True, tag="lift")
        self.disp_checkbox[t] = True
        if draw_check:
            # draw filled box
            x1 = x1 + 4
            y1 = y1 + 4
            x2 = x2 - 3
            y2 = y2 - 3
            points = rounded_box_coords(x1, y1, x2, y2, radius=4)
            if self.hidd_checkbox:
                t, sh = self.hidd_checkbox.popitem()
                self.coords(t, points)
                if sh:
                    self.itemconfig(t, fill=fill, outline=outline)
                else:
                    self.itemconfig(t, fill=fill, outline=outline, state="normal")
            else:
                t = self.create_polygon(points, fill=fill, outline=outline, smooth=True, tag="lift")
            self.disp_checkbox[t] = True

    def configure_scrollregion(self, last_col_line_pos: float) -> bool:
        try:
            self.configure(
                scrollregion=(
                    0,
                    0,
                    last_col_line_pos + self.ops.empty_horizontal + 2,
                    self.current_height,
                )
            )
            return True
        except Exception:
            return False

    def char_width_fn(self, c: str) -> int:
        if c in self.MT.char_widths[self.header_font]:
            return self.MT.char_widths[self.header_font][c]
        else:
            self.MT.txt_measure_canvas.itemconfig(
                self.MT.txt_measure_canvas_text,
                text=_test_str + c,
                font=self.header_font,
            )
            b = self.MT.txt_measure_canvas.bbox(self.MT.txt_measure_canvas_text)
            wd = b[2] - b[0] - self.header_test_str_w
            self.MT.char_widths[self.header_font][c] = wd
            return wd

    def redraw_corner(self, x: float, y: float, tags: str | tuple[str]) -> None:
        if self.hidd_corners:
            iid = self.hidd_corners.pop()
            self.coords(iid, x - 10, y, x, y, x, y + 10)
            self.itemconfig(iid, fill=self.ops.header_grid_fg, state="normal", tags=tags)
            self.disp_corners.add(iid)
        else:
            self.disp_corners.add(
                self.create_polygon(x - 10, y, x, y, x, y + 10, fill=self.ops.header_grid_fg, tags=tags)
            )

    def redraw_grid_and_text(
        self,
        last_col_line_pos: float,
        scrollpos_left: float,
        x_stop: float,
        grid_start_col: int,
        grid_end_col: int,
        text_start_col: int,
        text_end_col: int,
        scrollpos_right: float,
        col_pos_exists: bool,
        set_scrollregion: bool,
    ) -> bool:
        if set_scrollregion and not self.configure_scrollregion(last_col_line_pos=last_col_line_pos):
            return False
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
        self.visible_col_dividers = {}
        self.col_height_resize_bbox = (
            scrollpos_left,
            self.current_height - 2,
            x_stop,
            self.current_height,
        )
        top = self.canvasy(0)
        if (self.ops.show_vertical_grid or self.width_resizing_enabled) and col_pos_exists:
            yend = self.current_height - 5
            points = [
                x_stop - 1,
                self.current_height - 1,
                scrollpos_left - 1,
                self.current_height - 1,
                scrollpos_left - 1,
                -1,
            ]
            for c in range(grid_start_col, grid_end_col):
                draw_x = self.MT.col_positions[c]
                if c and self.width_resizing_enabled:
                    self.visible_col_dividers[c] = (draw_x - 2, 1, draw_x + 2, yend)
                points.extend(
                    (
                        draw_x,
                        -1,
                        draw_x,
                        self.current_height,
                        draw_x,
                        -1,
                        self.MT.col_positions[c + 1] if len(self.MT.col_positions) - 1 > c else draw_x,
                        -1,
                    )
                )
            self.redraw_gridline(points=points, fill=self.ops.header_grid_fg, width=1)
        sel_cells_bg = (
            self.ops.header_selected_cells_bg
            if self.ops.header_selected_cells_bg.startswith("#")
            else color_map[self.ops.header_selected_cells_bg]
        )
        sel_cols_bg = (
            self.ops.header_selected_columns_bg
            if self.ops.header_selected_columns_bg.startswith("#")
            else color_map[self.ops.header_selected_columns_bg]
        )
        font = self.ops.header_font
        selections = self.get_redraw_selections(text_start_col, grid_end_col)
        dd_coords = self.dropdown.get_coords()
        wrap = self.ops.header_wrap
        txt_h = self.MT.header_txt_height
        note_corners = self.ops.note_corners
        for c in range(text_start_col, text_end_col):
            draw_y = 3
            cleftgridln = self.MT.col_positions[c]
            crightgridln = self.MT.col_positions[c + 1]
            datacn = c if self.MT.all_columns_displayed else self.MT.displayed_columns[c]
            kwargs = self.get_cell_kwargs(datacn, key="dropdown")
            tag = f"{c}"
            fill, dd_drawn = self.redraw_highlight_get_text_fg(
                fc=cleftgridln,
                sc=crightgridln,
                c=c,
                sel_cells_bg=sel_cells_bg,
                sel_cols_bg=sel_cols_bg,
                selections=selections,
                datacn=datacn,
                has_dd=bool(kwargs),
                tags=("h", "c", tag),
            )
            if datacn in self.cell_options and "align" in self.cell_options[datacn]:
                align = self.cell_options[datacn]["align"]
            else:
                align = self.align
            if kwargs:
                max_width = crightgridln - cleftgridln - txt_h - 5
                if align[-1] == "w":
                    draw_x = cleftgridln + 2
                elif align[-1] == "e":
                    draw_x = crightgridln - 5 - txt_h
                elif align[-1] == "n":
                    draw_x = cleftgridln + (crightgridln - cleftgridln - txt_h) / 2
                self.redraw_dropdown(
                    cleftgridln,
                    0,
                    crightgridln,
                    self.current_height - 1,
                    fill=fill if kwargs["state"] != "disabled" else self.ops.header_grid_fg,
                    outline=fill,
                    draw_outline=not dd_drawn,
                    draw_arrow=max_width >= 5,
                    open_=dd_coords == c,
                )
            else:
                max_width = crightgridln - cleftgridln - 2
                if align[-1] == "w":
                    draw_x = cleftgridln + 2
                elif align[-1] == "e":
                    draw_x = crightgridln - 2
                elif align[-1] == "n":
                    draw_x = cleftgridln + (crightgridln - cleftgridln) / 2
                if (kwargs := self.get_cell_kwargs(datacn, key="checkbox")) and max_width > txt_h + 1:
                    box_w = txt_h + 1
                    if align[-1] == "w":
                        draw_x += box_w + 3
                    elif align[-1] == "n":
                        draw_x += box_w / 2 + 1
                    max_width -= box_w + 4
                    try:
                        draw_check = (
                            self.MT._headers[datacn]
                            if isinstance(self.MT._headers, (list, tuple))
                            else self.MT.data[self.MT._headers][datacn]
                        )
                    except Exception:
                        draw_check = False
                    self.redraw_checkbox(
                        cleftgridln + 2,
                        2,
                        cleftgridln + txt_h + 3,
                        txt_h + 3,
                        fill=fill if kwargs["state"] == "normal" else self.ops.header_grid_fg,
                        outline="",
                        draw_check=draw_check,
                    )
            if (
                max_width < self.MT.header_txt_width
                or (align[-1] == "w" and draw_x > scrollpos_right)
                or (align[-1] == "e" and cleftgridln + 5 > scrollpos_right)
                or (align[-1] == "n" and cleftgridln + 5 > scrollpos_right)
            ):
                continue
            tags = ("lift", "c", tag)
            if note_corners and max_width > 5 and datacn in self.cell_options and "note" in self.cell_options[datacn]:
                self.redraw_corner(crightgridln, 0, tags)
            text = self.cell_str(datacn, fix=False)
            if not text:
                continue
            gen_lines = wrap_text(
                text=text,
                max_width=max_width,
                max_lines=int((self.current_height - top - 2) / txt_h),
                char_width_fn=self.char_width_fn,
                widths=self.MT.char_widths[font],
                wrap=wrap,
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
            else:
                for line in gen_lines:
                    if self.hidd_text:
                        iid, showing = self.hidd_text.popitem()
                        self.coords(iid, draw_x, draw_y)
                        if showing:
                            self.itemconfig(
                                iid,
                                text=line,
                                fill=fill,
                                font=font,
                                anchor=align,
                                tags=tags,
                            )
                        else:
                            self.itemconfig(
                                iid,
                                text=line,
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
                            text=line,
                            fill=fill,
                            font=font,
                            anchor=align,
                            tags=tags,
                        )
                    self.disp_text[iid] = True
                    draw_y += self.MT.header_txt_height

        for dct in (self.hidd_text, self.hidd_high, self.hidd_grid, self.hidd_dropdown, self.hidd_checkbox):
            for iid, showing in dct.items():
                if showing:
                    self.itemconfig(iid, state="hidden")
                    dct[iid] = False
        for iid in self.hidd_corners:
            self.itemconfig(iid, state="hidden")
        self.tag_raise("lift")
        if self.disp_resize_lines:
            self.tag_raise("rw")
        self.tag_bind("c", "<Enter>", self.enter_cell)
        self.tag_bind("c", "<Leave>", self.leave_cell)
        return True

    def enter_cell(self, event: tk.Event | None = None) -> None:
        if any_editor_or_dropdown_open(self.MT):
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
        self.tooltip_after_id = self.after(self.ops.tooltip_hover_delay, self.check_and_show_tooltip)

    def check_and_show_tooltip(self, event: tk.Event | None = None) -> None:
        current_x, current_y = self.winfo_pointerx(), self.winfo_pointery()
        if current_x < 0 or current_y < 0:
            return
        if (
            not self.cget("cursor")
            and abs(current_x - self.tooltip_last_x) <= 1
            and abs(current_y - self.tooltip_last_y) <= 1
        ):
            self.show_tooltip()
        else:
            self.tooltip_last_x, self.tooltip_last_y = current_x, current_y
            self.tooltip_after_id = self.after(400, self.check_and_show_tooltip)

    def hide_tooltip(self):
        self.tooltip.withdraw()
        self.tooltip_showing, self.tooltip_coords = False, None

    def show_tooltip(self) -> None:
        c = int(self.tooltip_coords)
        datacn = self.MT.datacn(c)
        kws = self.get_cell_kwargs(datacn, key="note")
        if not self.ops.tooltips and not kws and not self.ops.user_can_create_notes:
            return
        self.MT.hide_tooltip()
        self.RI.hide_tooltip()
        cell_readonly = self.get_cell_kwargs(datacn, "readonly") or not self.MT.index_edit_cell_enabled()
        if kws:
            note = kws["note"]
            note_readonly = kws["readonly"]
        elif self.ops.user_can_create_notes:
            note = ""
            note_readonly = False
        else:
            note = None
            note_readonly = True
        self.tooltip_cell_content = f"{self.get_cell_data(datacn, none_to_empty_str=True)}"
        self.tooltip.reset(
            **{
                "text": self.tooltip_cell_content,
                "cell_readonly": cell_readonly,
                "note": note,
                "note_readonly": note_readonly,
                "row": 0,
                "col": c,
                "menu_kwargs": get_menu_kwargs(self.ops),
                **get_bg_fg(self.ops),
                "user_can_create_notes": self.ops.user_can_create_notes,
                "note_only": not self.ops.tooltips and isinstance(note, str),
                "width": self.ops.tooltip_width,
                "height": self.ops.tooltip_height,
            }
        )
        self.tooltip.set_position(self.tooltip_last_x - 4, self.tooltip_last_y - 4)
        self.tooltip_showing = True

    def close_tooltip_save(self, event: tk.Event | None = None) -> None:
        widget = self.winfo_containing(self.winfo_pointerx(), self.winfo_pointery())
        if any(widget == tw for tw in self.tooltip_widgets):
            return
        if not self.tooltip.cell_readonly or not self.tooltip.note_readonly:
            _, c, cell, note = self.tooltip.get()
            datacn = self.MT.datacn(c)
            if not self.tooltip.cell_readonly and cell != self.tooltip_cell_content:
                event_data = self.new_single_edit_event(c, datacn, "??", self.get_cell_data(datacn), cell)
                self.do_single_edit(c, datacn, event_data, cell)
            if not self.tooltip.note_readonly:
                span = self.PAR.span(None, datacn, None, datacn + 1).options(table=False, header=True)
                self.PAR.note(span, note=note if note else None, readonly=False)
            self.MT.refresh()
        self.hide_tooltip()
        self.focus_set()

    def get_redraw_selections(self, startc: int, endc: int) -> dict[str, set[int]]:
        d = defaultdict(set)
        for _, box in self.MT.get_selection_items():
            _, c1, _, c2 = box.coords
            for c in range(startc, endc):
                if c1 <= c and c2 > c:
                    d[box.type_ if box.type_ != "rows" else "cells"].add(c)
        return d

    def open_cell(self, event: Any = None, ignore_existing_editor: bool = False) -> None:
        if not self.MT.anything_selected() or (not ignore_existing_editor and self.text_editor.open):
            return
        if not self.MT.selected:
            return
        c = self.MT.selected.column
        datacn = self.MT.datacn(c)
        if self.get_cell_kwargs(datacn, key="readonly"):
            return
        elif self.get_cell_kwargs(datacn, key="dropdown") or self.get_cell_kwargs(datacn, key="checkbox"):
            if event_opens_dropdown_or_checkbox(event):
                if self.get_cell_kwargs(datacn, key="dropdown"):
                    self.open_dropdown_window(c, event=event)
                elif self.get_cell_kwargs(datacn, key="checkbox"):
                    self.click_checkbox(c, datacn)
        elif self.edit_cell_enabled:
            self.open_text_editor(event=event, c=c, dropdown=False)

    # displayed indexes
    def get_cell_align(self, c: int) -> str:
        datacn = c if self.MT.all_columns_displayed else self.MT.displayed_columns[c]
        if datacn in self.cell_options and "align" in self.cell_options[datacn]:
            align = self.cell_options[datacn]["align"]
        else:
            align = self.align
        return align

    # c is displayed col
    def open_text_editor(
        self,
        event: Any = None,
        c: int = 0,
        text: Any = None,
        state: str = "normal",
        dropdown: bool = False,
    ) -> bool:
        text = f"{self.get_cell_data(self.MT.datacn(c), none_to_empty_str=True, redirect_int=True)}"
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
                        name="begin_edit_header",
                        sheet=self.PAR.name,
                        key=extra_func_key,
                        value=text,
                        loc=c,
                        column=c,
                        boxes=self.MT.get_boxes(),
                        selected=self.MT.selected,
                    )
                )
            except Exception:
                return False
            if text is None:
                return False
            else:
                text = text if isinstance(text, str) else f"{text}"
        if self.ops.cell_auto_resize_enabled:
            if self.height_resizing_enabled:
                self.set_height_of_header_to_text(text)
            self.set_col_width_run_binding(c)
        if self.text_editor.open and c == self.text_editor.column:
            self.text_editor.set_text(self.text_editor.get() + "" if not isinstance(text, str) else text)
            return False
        self.hide_text_editor()
        self.hide_tooltip()
        if not self.MT.see(0, c, keep_yscroll=True):
            self.MT.main_table_redraw_grid_and_text(True, True)
        x = self.MT.col_positions[c] + 1
        y = 0
        w = self.MT.col_positions[c + 1] - x
        h = self.current_height + 1
        kwargs = {
            "menu_kwargs": get_menu_kwargs(self.ops),
            "sheet_ops": self.ops,
            "font": self.ops.header_font,
            "border_color": self.ops.header_selected_columns_bg,
            "text": text,
            "state": state,
            "width": w,
            "height": h,
            "show_border": True,
            "bg": self.ops.header_editor_bg,
            "fg": self.ops.header_editor_fg,
            "select_bg": self.ops.header_editor_select_bg,
            "select_fg": self.ops.header_editor_select_fg,
            "align": self.get_cell_align(c),
            "c": c,
        }
        if not self.text_editor.window:
            self.text_editor.window = TextEditor(
                self, newline_binding=self.text_editor_newline_binding, rc_bindings=self.ops.rc_bindings
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
        for key, func in self.MT.text_editor_user_bound_keys.items():
            self.text_editor.tktext.bind(key, func)
        return True

    # displayed indexes
    def text_editor_has_wrapped(
        self,
        r: int = 0,
        c: int = 0,
        check_lines: None = None,  # just here to receive text editor arg
    ) -> None:
        if self.width_resizing_enabled:
            curr_width = self.text_editor.window.winfo_width()
            new_width = curr_width + (self.MT.header_txt_height * 2)
            if new_width != curr_width:
                self.text_editor.window.config(width=new_width)
                self.set_col_width_run_binding(c, width=new_width, only_if_too_small=False)
                if self.dropdown.open and self.dropdown.get_coords() == c:
                    self.itemconfig(self.dropdown.canvas_id, width=new_width)
                    self.dropdown.window.update_idletasks()
                    self.dropdown.window._reselect()
                self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=False, redraw_table=True)
                self.coords(self.text_editor.canvas_id, self.MT.col_positions[c] + 1, 0)

    # displayed indexes
    def text_editor_newline_binding(self, event: Any = None, check_lines: bool = True) -> None:
        if not self.height_resizing_enabled:
            return
        curr_height = self.text_editor.window.winfo_height()
        if curr_height < self.MT.min_header_height:
            return
        if (
            not check_lines
            or self.MT.get_lines_cell_height(
                self.text_editor.window.get_num_lines() + 1,
                font=self.text_editor.tktext.cget("font"),
            )
            > curr_height
        ):
            c = self.text_editor.column
            new_height = curr_height + self.MT.header_txt_height
            space_bot = self.MT.get_space_bot(0)
            if new_height > space_bot:
                new_height = space_bot
            if new_height != curr_height:
                self.text_editor.window.config(height=new_height)
                self.set_height(new_height, set_TL=True)
                if self.dropdown.open and self.dropdown.get_coords() == c:
                    win_h, anchor = self.get_dropdown_height_anchor(c, new_height)
                    self.coords(
                        self.dropdown.canvas_id,
                        self.MT.col_positions[c],
                        new_height - 1,
                    )
                    self.itemconfig(self.dropdown.canvas_id, anchor=anchor, height=win_h)

    def refresh_open_window_positions(self, zoom: Literal["in", "out"]) -> None:
        if self.text_editor.open:
            c = self.text_editor.column
            self.text_editor.window.config(
                height=self.current_height,
                width=self.MT.col_positions[c + 1] - self.MT.col_positions[c] + 1,
            )
            self.text_editor.tktext.config(font=self.ops.header_font)
            self.coords(
                self.text_editor.canvas_id,
                self.MT.col_positions[c],
                0,
            )
        if self.dropdown.open:
            if zoom == "in":
                self.dropdown.window.zoom_in()
            elif zoom == "out":
                self.dropdown.window.zoom_out()
            c = self.dropdown.get_coords()
            if self.text_editor.open:
                text_editor_h = self.text_editor.window.winfo_height()
                win_h, anchor = self.get_dropdown_height_anchor(c, text_editor_h)
            else:
                text_editor_h = self.current_height
                anchor = self.itemcget(self.dropdown.canvas_id, "anchor")
                # win_h = 0
            self.dropdown.window.config(width=self.MT.col_positions[c + 1] - self.MT.col_positions[c] + 1)
            if anchor == "nw":
                self.coords(
                    self.dropdown.canvas_id,
                    self.MT.col_positions[c],
                    text_editor_h - 1,
                )
                # self.itemconfig(self.dropdown.canvas_id, anchor=anchor, height=win_h)
            elif anchor == "sw":
                self.coords(
                    self.dropdown.canvas_id,
                    self.MT.col_positions[c],
                    0,
                )
                # self.itemconfig(self.dropdown.canvas_id, anchor=anchor, height=win_h)

    def hide_text_editor(self) -> None:
        if self.text_editor.open:
            for binding in text_editor_to_unbind:
                self.text_editor.tktext.unbind(binding)
            self.itemconfig(self.text_editor.canvas_id, state="hidden")
            self.text_editor.open = False

    def do_single_edit(self, c: int, datacn: int, event_data: EventDataDict, val):
        edited = False
        set_data = partial(self.set_cell_data_undo, c=c, datacn=datacn, check_input_valid=False)
        if self.MT.edit_validation_func:
            val = self.MT.edit_validation_func(event_data)
            if val is not None and self.input_valid_for_cell(datacn, val):
                edited = set_data(value=val)
        elif self.input_valid_for_cell(datacn, val):
            edited = set_data(value=val)
        if edited:
            try_binding(self.extra_end_edit_cell_func, event_data)

    # c is displayed col
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
        text_editor_value = self.text_editor.get()
        c = self.text_editor.column
        datacn = c if self.MT.all_columns_displayed else self.MT.displayed_columns[c]
        event_data = self.new_single_edit_event(c, datacn, event.keysym, self.get_cell_data(datacn), text_editor_value)
        self.do_single_edit(c, datacn, event_data, text_editor_value)
        self.MT.recreate_all_selection_boxes()
        self.hide_text_editor_and_dropdown()
        if event.keysym != "FocusOut":
            self.focus_set()
        return "break"

    def get_dropdown_height_anchor(
        self,
        c: int,
        text_editor_h: None | int = None,
    ) -> tuple[int, Literal["nw"]]:
        win_h = 5
        datacn = self.MT.datacn(c)
        for i, v in enumerate(self.get_cell_kwargs(datacn, key="dropdown")["values"]):
            v_numlines = len(v.split("\n") if isinstance(v, str) else f"{v}".split("\n"))
            if v_numlines > 1:
                win_h += self.MT.header_txt_height + (v_numlines * self.MT.header_txt_height) + 5  # end of cell
            else:
                win_h += self.MT.min_header_height
            if i == 5:
                break
        win_h = min(win_h, 500)
        space_bot = self.MT.get_space_bot(0, text_editor_h)
        win_h2 = int(win_h)
        if win_h > space_bot:
            win_h = space_bot - 1
        if win_h < self.MT.header_txt_height + 5:
            win_h = self.MT.header_txt_height + 5
        elif win_h > win_h2:
            win_h = win_h2
        return win_h, "nw"

    def dropdown_text_editor_modified(
        self,
        event: tk.Misc,
    ) -> None:
        c = self.dropdown.get_coords()
        event_data = event_dict(
            name="table_dropdown_modified",
            sheet=self.PAR.name,
            value=self.text_editor.get(),
            loc=c,
            row=0,
            column=c,
            boxes=self.MT.get_boxes(),
            selected=self.MT.selected,
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

    def open_dropdown_window(self, c: int, event: Any = None) -> None:
        self.hide_text_editor()
        kwargs = self.get_cell_kwargs(self.MT.datacn(c), key="dropdown")
        if kwargs["state"] == "disabled":
            return
        if kwargs["state"] == "normal" and not self.open_text_editor(event=event, c=c, dropdown=True):
            return
        win_h, anchor = self.get_dropdown_height_anchor(c)
        win_w = self.MT.col_positions[c + 1] - self.MT.col_positions[c] + 1
        ypos = self.current_height - 1
        reset_kwargs = {
            "r": 0,
            "c": c,
            "bg": self.ops.header_editor_bg,
            "fg": self.ops.header_editor_fg,
            "select_bg": self.ops.header_editor_select_bg,
            "select_fg": self.ops.header_editor_select_fg,
            "width": win_w,
            "height": win_h,
            "font": self.ops.header_font,
            "ops": self.ops,
            "outline_color": self.ops.header_selected_columns_bg,
            "align": self.get_cell_align(c),
            "values": kwargs["values"],
            "search_function": kwargs["search_function"],
            "modified_function": kwargs["modified_function"],
        }
        if self.dropdown.window:
            self.dropdown.window.reset(**reset_kwargs)
            self.itemconfig(self.dropdown.canvas_id, state="normal")
            self.coords(self.dropdown.canvas_id, self.MT.col_positions[c], ypos)
            self.dropdown.window.tkraise()
        else:
            self.dropdown.window = self.PAR._dropdown_cls(
                self.winfo_toplevel(),
                **reset_kwargs,
                single_index="c",
                close_dropdown_window=self.close_dropdown_window,
                arrowkey_RIGHT=self.MT.arrowkey_RIGHT,
                arrowkey_LEFT=self.MT.arrowkey_LEFT,
            )
            self.dropdown.canvas_id = self.create_window(
                (self.MT.col_positions[c], ypos),
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
            self.dropdown.window.bind("<FocusOut>", lambda _x: self.close_dropdown_window(c))
            self.dropdown.window.bind("<Escape>", self.close_dropdown_window)
            self.dropdown.window.focus_set()
            redraw = True
        self.dropdown.open = True
        if redraw:
            self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=False, redraw_table=False)

    def close_dropdown_window(
        self,
        c: None | int = None,
        selection: Any = None,
        redraw: bool = True,
    ) -> None:
        if c is not None and selection is not None:
            datacn = c if self.MT.all_columns_displayed else self.MT.displayed_columns[c]
            kwargs = self.get_cell_kwargs(datacn, key="dropdown")
            event_data = self.new_single_edit_event(c, datacn, "??", self.get_cell_data(datacn), selection)
            try_binding(kwargs["select_function"], event_data)
            selection = selection if not self.MT.edit_validation_func else self.MT.edit_validation_func(event_data)
            if selection is not None:
                edited = self.set_cell_data_undo(c, datacn=datacn, value=selection, redraw=not redraw)
                if edited:
                    try_binding(self.extra_end_edit_cell_func, event_data)
            self.MT.recreate_all_selection_boxes()
        self.focus_set()
        self.hide_text_editor_and_dropdown(redraw=redraw)

    def hide_text_editor_and_dropdown(self, redraw: bool = True) -> None:
        self.hide_text_editor()
        self.hide_dropdown_window()
        if redraw:
            self.MT.refresh()

    def mouseclick_outside_editor_or_dropdown(self, inside: bool = False) -> int | None:
        closed_dd_coords = self.dropdown.get_coords()
        if self.text_editor.open:
            self.close_text_editor(new_tk_event("ButtonPress-1"))
        if closed_dd_coords is not None:
            self.hide_dropdown_window()
            if inside:
                self.MT.main_table_redraw_grid_and_text(
                    redraw_header=True,
                    redraw_row_index=False,
                    redraw_table=False,
                )
        return closed_dd_coords

    def mouseclick_outside_editor_or_dropdown_all_canvases(self, inside: bool = False) -> int | None:
        self.RI.mouseclick_outside_editor_or_dropdown()
        self.MT.mouseclick_outside_editor_or_dropdown()
        return self.mouseclick_outside_editor_or_dropdown(inside)

    def hide_dropdown_window(self) -> None:
        if self.dropdown.open:
            self.dropdown.window.unbind("<FocusOut>")
            self.itemconfig(self.dropdown.canvas_id, state="hidden")
            self.dropdown.open = False

    # internal event use
    def set_cell_data_undo(
        self,
        c: int = 0,
        datacn: int | None = None,
        value: Any = "",
        cell_resize: bool = True,
        undo: bool = True,
        redraw: bool = True,
        check_input_valid: bool = True,
    ) -> bool:
        if datacn is None:
            datacn = c if self.MT.all_columns_displayed else self.MT.displayed_columns[c]
        event_data = event_dict(
            name="edit_header",
            sheet=self.PAR.name,
            widget=self,
            cells_header={datacn: self.get_cell_data(datacn)},
            boxes=self.MT.get_boxes(),
            selected=self.MT.selected,
        )
        edited = False
        if isinstance(self.MT._headers, int):
            disprn = self.MT.try_disprn(self.MT._headers)
            edited = self.MT.set_cell_data_undo(
                r=disprn if isinstance(disprn, int) else 0,
                c=c,
                datarn=self.MT._headers,
                datacn=datacn,
                value=value,
                undo=True,
                cell_resize=isinstance(disprn, int),
            )
        else:
            self.fix_header(datacn)
            if not check_input_valid or self.input_valid_for_cell(datacn, value):
                if self.MT.undo_enabled and undo:
                    self.MT.undo_stack.append(stored_event_dict(event_data))
                self.set_cell_data(datacn=datacn, value=value)
                edited = True
        if edited and cell_resize and self.ops.cell_auto_resize_enabled:
            if self.height_resizing_enabled:
                self.set_height_of_header_to_text(self.cell_str(datacn, fix=False))
            self.set_col_width_run_binding(c)
        if redraw:
            self.MT.refresh()
        if edited:
            self.MT.sheet_modified(event_data)
        return edited

    def set_cell_data(self, datacn: int | None = None, value: Any = "") -> None:
        if isinstance(self.MT._headers, int):
            self.MT.set_cell_data(datarn=self.MT._headers, datacn=datacn, value=value)
        else:
            self.fix_header(datacn)
            if self.get_cell_kwargs(datacn, key="checkbox"):
                self.MT._headers[datacn] = try_to_bool(value)
            else:
                self.MT._headers[datacn] = value

    def input_valid_for_cell(self, datacn: int, value: Any, check_readonly: bool = True) -> bool:
        if check_readonly and self.get_cell_kwargs(datacn, key="readonly"):
            return False
        elif self.get_cell_kwargs(datacn, key="checkbox"):
            return is_bool_like(value)
        else:
            return not (
                self.cell_equal_to(datacn, value)
                or (
                    (kwargs := self.get_cell_kwargs(datacn, key="dropdown"))
                    and kwargs["validate_input"]
                    and value not in kwargs["values"]
                )
            )

    def cell_equal_to(self, datacn: int, value: Any) -> bool:
        self.fix_header(datacn)
        if isinstance(self.MT._headers, list):
            return self.MT._headers[datacn] == value
        elif isinstance(self.MT._headers, int):
            return self.MT.cell_equal_to(self.MT._headers, datacn, value)

    def get_cell_data(
        self,
        datacn: int,
        get_displayed: bool = False,
        none_to_empty_str: bool = False,
        redirect_int: bool = False,
    ) -> Any:
        if get_displayed:
            return self.cell_str(datacn, fix=False)
        if redirect_int and isinstance(self.MT._headers, int):  # internal use
            return self.MT.get_cell_data(self.MT._headers, datacn, none_to_empty_str=True)
        if (
            isinstance(self.MT._headers, int)
            or not self.MT._headers
            or datacn >= len(self.MT._headers)
            or (self.MT._headers[datacn] is None and none_to_empty_str)
        ):
            return ""
        return self.MT._headers[datacn]

    def cell_str(self, datacn: int, fix: bool = True) -> str:
        kwargs = self.get_cell_kwargs(datacn, key="dropdown")
        if kwargs:
            if kwargs["text"] is not None:
                return f"{kwargs['text']}"
        else:
            kwargs = self.get_cell_kwargs(datacn, key="checkbox")
            if kwargs:
                return f"{kwargs['text']}"
        if isinstance(self.MT._headers, int):
            return self.MT.cell_str(self.MT._headers, datacn, get_displayed=True)
        if fix:
            self.fix_header(datacn)
        try:
            value = self.MT._headers[datacn]
            if value is None:
                value = ""
            elif not isinstance(value, str):
                value = f"{value}"
        except Exception:
            value = ""
        if not value and self.ops.show_default_header_for_empty:
            value = get_n2a(datacn, self.ops.default_header)
        return value

    def get_value_for_empty_cell(self, datacn: int, c_ops: bool = True) -> Any:
        if self.get_cell_kwargs(datacn, key="checkbox", cell=c_ops):
            return False
        elif (kwargs := self.get_cell_kwargs(datacn, key="dropdown", cell=c_ops)) and kwargs["validate_input"]:
            if kwargs["default_value"] is None:
                if kwargs["values"]:
                    return safe_copy(kwargs["values"][0])
                else:
                    return ""
            else:
                return safe_copy(kwargs["default_value"])
        else:
            return ""

    def get_empty_header_seq(self, end: int, start: int = 0, c_ops: bool = True) -> list[Any]:
        return [self.get_value_for_empty_cell(datacn, c_ops=c_ops) for datacn in range(start, end)]

    def fix_header(self, datacn: None | int = None) -> None:
        if isinstance(self.MT._headers, int):
            return
        if isinstance(self.MT._headers, float):
            self.MT._headers = int(self.MT._headers)
            return
        if not isinstance(self.MT._headers, list):
            try:
                self.MT._headers = list(self.MT._headers)
            except Exception:
                self.MT._headers = []
        if isinstance(datacn, int) and datacn >= len(self.MT._headers):
            self.MT._headers.extend(self.get_empty_header_seq(end=datacn + 1, start=len(self.MT._headers)))

    # displayed indexes
    def set_col_width_run_binding(self, c: int, width: int | None = None, only_if_too_small: bool = True) -> None:
        old_width = self.MT.col_positions[c + 1] - self.MT.col_positions[c]
        new_width = self.set_col_width(c, width=width, only_if_too_small=only_if_too_small)
        if self.column_width_resize_func is not None and old_width != new_width:
            self.column_width_resize_func(
                event_dict(
                    name="resize",
                    sheet=self.PAR.name,
                    resized_columns={c: {"old_size": old_width, "new_size": new_width}},
                )
            )

    # internal event use
    def click_checkbox(self, c: int, datacn: int | None = None, undo: bool = True, redraw: bool = True) -> None:
        if datacn is None:
            datacn = c if self.MT.all_columns_displayed else self.MT.displayed_columns[c]
        kwargs = self.get_cell_kwargs(datacn, key="checkbox")
        if kwargs["state"] == "normal":
            pre_edit_value = self.get_cell_data(datacn)
            if isinstance(self.MT._headers, list):
                value = not self.MT._headers[datacn] if isinstance(self.MT._headers[datacn], bool) else False
            elif isinstance(self.MT._headers, int):
                value = (
                    not self.MT.data[self.MT._headers][datacn]
                    if isinstance(self.MT.data[self.MT._headers][datacn], bool)
                    else False
                )
            else:
                value = False
            self.set_cell_data_undo(c, datacn=datacn, value=value, cell_resize=False)
            event_data = self.new_single_edit_event(c, datacn, "??", pre_edit_value, value)
            if kwargs["check_function"] is not None:
                kwargs["check_function"](event_data)
            try_binding(self.extra_end_edit_cell_func, event_data)
        if redraw:
            self.MT.refresh()

    def get_cell_kwargs(self, datacn: int, key: Hashable | None = "dropdown", cell: bool = True) -> dict:
        if cell and datacn in self.cell_options:
            return self.cell_options[datacn] if key is None else self.cell_options[datacn].get(key, {})
        else:
            return {}

    def new_single_edit_event(self, c: int, datacn: int, k: str, before_val: Any, after_val: Any) -> EventDataDict:
        return event_dict(
            name="end_edit_header",
            sheet=self.PAR.name,
            widget=self,
            cells_header={datacn: before_val},
            key=k,
            value=after_val,
            loc=c,
            column=c,
            boxes=self.MT.get_boxes(),
            selected=self.MT.selected,
            data={datacn: after_val},
        )
