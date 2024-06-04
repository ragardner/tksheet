from __future__ import annotations

import tkinter as tk
from collections import defaultdict
from collections.abc import (
    Callable,
    Hashable,
    Sequence,
)
from functools import (
    partial,
)
from itertools import (
    cycle,
    islice,
    repeat,
)
from math import ceil, floor
from operator import (
    itemgetter,
)
from typing import Literal

from .colors import (
    color_map,
)
from .formatters import is_bool_like, try_to_bool
from .functions import (
    consecutive_ranges,
    event_dict,
    get_n2a,
    is_contiguous,
    new_tk_event,
    pickled_event_dict,
    rounded_box_coords,
    try_binding,
)
from .other_classes import (
    DotDict,
    DraggedRowColumn,
    DropdownStorage,
    TextEditorStorage,
)
from .text_editor import (
    TextEditor,
)
from .vars import (
    USER_OS,
    rc_binding,
    symbols_set,
    text_editor_close_bindings,
    text_editor_newline_bindings,
    text_editor_to_unbind,
)


class ColumnHeaders(tk.Canvas):
    def __init__(self, *args, **kwargs):
        tk.Canvas.__init__(
            self,
            kwargs["parent"],
            background=kwargs["parent"].ops.header_bg,
            highlightthickness=0,
        )
        self.PAR = kwargs["parent"]
        self.current_height = None  # is set from within MainTable() __init__ or from Sheet parameters
        self.MT = None  # is set from within MainTable() __init__
        self.RI = None  # is set from within MainTable() __init__
        self.TL = None  # is set from within TopLeftRectangle() __init__
        self.popup_menu_loc = None
        self.extra_begin_edit_cell_func = None
        self.extra_end_edit_cell_func = None
        self.centre_alignment_text_mod_indexes = (slice(1, None), slice(None, -1))
        self.c_align_cyc = cycle(self.centre_alignment_text_mod_indexes)
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
        self.col_height_resize_bbox = tuple()
        self.cell_options = {}
        self.rsz_w = None
        self.rsz_h = None
        self.new_col_height = 0
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
        self.disp_boxes = set()
        self.hidd_text = {}
        self.hidd_high = {}
        self.hidd_grid = {}
        self.hidd_fill_sels = {}
        self.hidd_resize_lines = {}
        self.hidd_dropdown = {}
        self.hidd_checkbox = {}
        self.hidd_boxes = set()

        self.default_header = kwargs["default_header"].lower()
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
            self.bind(rc_binding, self.rc)
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
            self.unbind(rc_binding)
            self.unbind("<MouseWheel>")
            if USER_OS == "linux":
                self.unbind("<Button-4>")
                self.unbind("<Button-5>")

    def mousewheel(self, event: object) -> None:
        if isinstance(self.MT._headers, int):
            maxlines = max(
                (
                    len(
                        self.MT.get_valid_cell_data_as_str(self.MT._headers, datacn, get_displayed=True)
                        .rstrip()
                        .split("\n")
                    )
                    for datacn in range(len(self.MT.data[self.MT._headers]))
                ),
                default=0,
            )
        elif isinstance(self.MT._headers, (list, tuple)):
            maxlines = max(
                (
                    len(e.rstrip().split("\n")) if isinstance(e, str) else len(f"{e}".rstrip().split("\n"))
                    for e in self.MT._headers
                ),
                default=0,
            )
        if maxlines == 1:
            maxlines = 0
        if self.lines_start_at > maxlines:
            self.lines_start_at = maxlines
        if (event.delta < 0 or event.num == 5) and self.lines_start_at < maxlines:
            self.lines_start_at += 1
        elif (event.delta >= 0 or event.num == 4) and self.lines_start_at > 0:
            self.lines_start_at -= 1
        self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=False, redraw_table=False)

    def set_height(self, new_height: int, set_TL: bool = False) -> None:
        self.current_height = new_height
        try:
            self.config(height=new_height)
        except Exception:
            return
        if set_TL and self.TL is not None:
            self.TL.set_dimensions(new_h=new_height)

    def rc(self, event: object) -> None:
        self.mouseclick_outside_editor_or_dropdown_all_canvases(inside=True)
        self.focus_set()
        popup_menu = None
        if self.MT.identify_col(x=event.x, allow_end=False) is None:
            self.MT.deselect("all")
            c = len(self.MT.col_positions) - 1
            if self.MT.rc_popup_menus_enabled:
                popup_menu = self.MT.empty_rc_popup_menu
        elif self.col_selection_enabled and not self.currently_resizing_width and not self.currently_resizing_height:
            c = self.MT.identify_col(x=event.x)
            if c < len(self.MT.col_positions) - 1:
                if self.MT.col_selected(c):
                    if self.MT.rc_popup_menus_enabled:
                        popup_menu = self.ch_rc_popup_menu
                else:
                    if self.MT.single_selection_enabled and self.MT.rc_select_enabled:
                        self.select_col(c, redraw=True)
                    elif self.MT.toggle_selection_enabled and self.MT.rc_select_enabled:
                        self.toggle_select_col(c, redraw=True)
                    if self.MT.rc_popup_menus_enabled:
                        popup_menu = self.ch_rc_popup_menu
        try_binding(self.extra_rc_func, event)
        if popup_menu is not None:
            self.popup_menu_loc = c
            popup_menu.tk_popup(event.x_root, event.y_root)

    def ctrl_b1_press(self, event: object) -> None:
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

    def ctrl_shift_b1_press(self, event: object) -> None:
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

    def shift_b1_press(self, event: object) -> None:
        self.mouseclick_outside_editor_or_dropdown_all_canvases(inside=True)
        x = event.x
        c = self.MT.identify_col(x=x)
        if (self.drag_and_drop_enabled or self.col_selection_enabled) and self.rsz_h is None and self.rsz_w is None:
            if c < len(self.MT.col_positions) - 1:
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
        if c > min_c:
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

    def mouse_motion(self, event: object) -> None:
        if not self.currently_resizing_height and not self.currently_resizing_width:
            x = self.canvasx(event.x)
            y = self.canvasy(event.y)
            mouse_over_resize = False
            mouse_over_selected = False
            if self.width_resizing_enabled:
                c = self.check_mouse_position_width_resizers(x, y)
                if c is not None:
                    self.rsz_w, mouse_over_resize = c, True
                    if self.MT.current_cursor != "sb_h_double_arrow":
                        self.config(cursor="sb_h_double_arrow")
                        self.MT.current_cursor = "sb_h_double_arrow"
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
                        if self.MT.current_cursor != "sb_v_double_arrow":
                            self.config(cursor="sb_v_double_arrow")
                            self.MT.current_cursor = "sb_v_double_arrow"
                    else:
                        self.rsz_h = None
                except Exception:
                    self.rsz_h = None
            if not mouse_over_resize:
                if self.MT.col_selected(self.MT.identify_col(event, allow_end=False)):
                    mouse_over_selected = True
                    if self.MT.current_cursor != "hand2":
                        self.config(cursor="hand2")
                        self.MT.current_cursor = "hand2"
            if not mouse_over_resize and not mouse_over_selected:
                self.MT.reset_mouse_motion_creations()
        try_binding(self.extra_motion_func, event)

    def double_b1(self, event: object) -> None:
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

    def b1_press(self, event: object) -> None:
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
                fill=self.PAR.ops.resizing_line_fg,
                tag="rwl",
            )
            self.MT.create_resize_line(x, y1, x, y2, width=1, fill=self.PAR.ops.resizing_line_fg, tag="rwl")
            self.create_resize_line(
                line2x,
                0,
                line2x,
                self.current_height,
                width=1,
                fill=self.PAR.ops.resizing_line_fg,
                tag="rwl2",
            )
            self.MT.create_resize_line(line2x, y1, line2x, y2, width=1, fill=self.PAR.ops.resizing_line_fg, tag="rwl2")
        elif self.height_resizing_enabled and self.rsz_w is None and self.rsz_h is not None:
            x1, y1, x2, y2 = self.MT.get_canvas_visible_area()
            self.currently_resizing_height = True
            y = event.y
            if y < self.MT.min_header_height:
                y = int(self.MT.min_header_height)
            self.new_col_height = y
            self.create_resize_line(x1, y, x2, y, width=1, fill=self.PAR.ops.resizing_line_fg, tag="rhl")
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

    def b1_motion(self, event: object) -> None:
        x1, y1, x2, y2 = self.MT.get_canvas_visible_area()
        if self.width_resizing_enabled and self.rsz_w is not None and self.currently_resizing_width:
            x = self.canvasx(event.x)
            size = x - self.MT.col_positions[self.rsz_w - 1]
            if size >= self.MT.min_column_width and size < self.MT.max_column_width:
                self.hide_resize_and_ctrl_lines(ctrl_lines=False)
                line2x = self.MT.col_positions[self.rsz_w - 1]
                self.create_resize_line(
                    x,
                    0,
                    x,
                    self.current_height,
                    width=1,
                    fill=self.PAR.ops.resizing_line_fg,
                    tag="rwl",
                )
                self.MT.create_resize_line(x, y1, x, y2, width=1, fill=self.PAR.ops.resizing_line_fg, tag="rwl")
                self.create_resize_line(
                    line2x,
                    0,
                    line2x,
                    self.current_height,
                    width=1,
                    fill=self.PAR.ops.resizing_line_fg,
                    tag="rwl2",
                )
                self.MT.create_resize_line(
                    line2x,
                    y1,
                    line2x,
                    y2,
                    width=1,
                    fill=self.PAR.ops.resizing_line_fg,
                    tag="rwl2",
                )
        elif self.height_resizing_enabled and self.rsz_h is not None and self.currently_resizing_height:
            evy = event.y
            self.hide_resize_and_ctrl_lines(ctrl_lines=False)
            if evy > self.current_height:
                y = self.MT.canvasy(evy - self.current_height)
                if evy > self.MT.max_header_height:
                    evy = int(self.MT.max_header_height)
                    y = self.MT.canvasy(evy - self.current_height)
                self.new_col_height = evy
                self.MT.create_resize_line(x1, y, x2, y, width=1, fill=self.PAR.ops.resizing_line_fg, tag="rhl")
            else:
                y = evy
                if y < self.MT.min_header_height:
                    y = int(self.MT.min_header_height)
                self.new_col_height = y
                self.create_resize_line(x1, y, x2, y, width=1, fill=self.PAR.ops.resizing_line_fg, tag="rhl")
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

    def get_b1_motion_box(self, start_col: int, end_col: int) -> tuple[int, int, int, int, Literal["columns"]]:
        if end_col >= start_col:
            return 0, start_col, len(self.MT.row_positions) - 1, end_col + 1, "columns"
        elif end_col < start_col:
            return 0, end_col, len(self.MT.row_positions) - 1, start_col + 1, "columns"

    def ctrl_b1_motion(self, event: object) -> None:
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

    def drag_and_drop_motion(self, event: object) -> float:
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
            fill=self.PAR.ops.drag_and_drop_bg,
            tag="move_columns",
        )
        self.MT.create_resize_line(xpos, y1, xpos, y2, width=3, fill=self.PAR.ops.drag_and_drop_bg, tag="move_columns")
        for boxst, boxend in consecutive_ranges(cols):
            self.MT.show_ctrl_outline(
                start_cell=(boxst, 0),
                end_cell=(boxend, len(self.MT.row_positions) - 1),
                dash=(),
                outline=self.PAR.ops.drag_and_drop_bg,
                delete_on_timer=False,
            )

    def hide_resize_and_ctrl_lines(self, ctrl_lines: bool = True) -> None:
        self.delete_resize_lines()
        self.MT.delete_resize_lines()
        if ctrl_lines:
            self.MT.delete_ctrl_outlines()

    def scroll_if_event_offscreen(self, event: object) -> bool:
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

    def event_over_dropdown(self, c: int, datacn: int, event: object, canvasx: float) -> bool:
        if (
            event.y < self.MT.header_txt_height + 5
            and self.get_cell_kwargs(datacn, key="dropdown")
            and canvasx < self.MT.col_positions[c + 1]
            and canvasx > self.MT.col_positions[c + 1] - self.MT.header_txt_height - 4
        ):
            return True
        return False

    def event_over_checkbox(self, c: int, datacn: int, event: object, canvasx: float) -> bool:
        if (
            event.y < self.MT.header_txt_height + 5
            and self.get_cell_kwargs(datacn, key="checkbox")
            and canvasx < self.MT.col_positions[c] + self.MT.header_txt_height + 4
        ):
            return True
        return False

    def b1_release(self, event: object) -> None:
        if self.being_drawn_item is not None and (to_sel := self.MT.coords_and_type(self.being_drawn_item)):
            r_to_sel, c_to_sel = self.MT.selected.row, self.MT.selected.column
            self.MT.hide_selection_box(self.being_drawn_item)
            self.MT.set_currently_selected(
                r_to_sel,
                c_to_sel,
                item=self.MT.create_selection_box(*to_sel, set_current=False),
            )
            sel_event = self.MT.get_select_event(being_drawn_item=self.being_drawn_item)
            try_binding(self.drag_selection_binding_func, sel_event)
            self.PAR.emit_event("<<SheetSelect>>", data=sel_event)
        else:
            self.being_drawn_item = None
        self.MT.bind("<MouseWheel>", self.MT.mousewheel)
        if self.width_resizing_enabled and self.rsz_w is not None and self.currently_resizing_width:
            self.currently_resizing_width = False
            new_col_pos = int(self.coords("rwl")[0])
            self.hide_resize_and_ctrl_lines(ctrl_lines=False)
            old_width = self.MT.col_positions[self.rsz_w] - self.MT.col_positions[self.rsz_w - 1]
            size = new_col_pos - self.MT.col_positions[self.rsz_w - 1]
            if size < self.MT.min_column_width:
                new_col_pos = ceil(self.MT.col_positions[self.rsz_w - 1] + self.MT.min_column_width)
            elif size > self.MT.max_column_width:
                new_col_pos = floor(self.MT.col_positions[self.rsz_w - 1] + self.MT.max_column_width)
            increment = new_col_pos - self.MT.col_positions[self.rsz_w]
            self.MT.col_positions[self.rsz_w + 1 :] = [
                e + increment for e in islice(self.MT.col_positions, self.rsz_w + 1, len(self.MT.col_positions))
            ]
            self.MT.col_positions[self.rsz_w] = new_col_pos
            new_width = self.MT.col_positions[self.rsz_w] - self.MT.col_positions[self.rsz_w - 1]
            self.MT.allow_auto_resize_columns = False
            self.MT.recreate_all_selection_boxes()
            self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
            if self.column_width_resize_func is not None and old_width != new_width:
                self.column_width_resize_func(
                    event_dict(
                        name="resize",
                        sheet=self.PAR.name,
                        resized_columns={self.rsz_w - 1: {"old_size": old_width, "new_size": new_width}},
                    )
                )
        elif self.height_resizing_enabled and self.rsz_h is not None and self.currently_resizing_height:
            self.currently_resizing_height = False
            self.hide_resize_and_ctrl_lines(ctrl_lines=False)
            self.set_height(self.new_col_height, set_TL=True)
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
                event_data = event_dict(
                    name="move_columns",
                    sheet=self.PAR.name,
                    widget=self,
                    boxes=self.MT.get_boxes(),
                    selected=self.MT.selected,
                    value=c,
                )
                if try_binding(self.ch_extra_begin_drag_drop_func, event_data, "begin_move_columns"):
                    data_new_idxs, disp_new_idxs, event_data = self.MT.move_columns_adjust_options_dict(
                        *self.MT.get_args_for_move_columns(
                            move_to=c,
                            to_move=self.dragged_col.to_move,
                        ),
                        move_data=self.PAR.ops.column_drag_and_drop_perform,
                        move_widths=self.PAR.ops.column_drag_and_drop_perform,
                        event_data=event_data,
                    )
                    event_data["moved"]["columns"] = {
                        "data": data_new_idxs,
                        "displayed": disp_new_idxs,
                    }
                    if self.MT.undo_enabled:
                        self.MT.undo_stack.append(pickled_event_dict(event_data))
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
        c: int,
        redraw: bool = False,
        run_binding_func: bool = True,
        ext: bool = False,
    ) -> int:
        self.MT.deselect("all", redraw=False)
        fill_iid = self.MT.create_selection_box(0, c, len(self.MT.row_positions) - 1, c + 1, "columns", ext=ext)
        if redraw:
            self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
        if run_binding_func:
            self.MT.run_selection_binding("columns")
        return fill_iid

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

    def display_box(
        self,
        x1: int,
        y1: int,
        x2: int,
        y2: int,
        fill: str,
        outline: str,
        state: str,
        tags: str | tuple[str],
        iid: None | int = None,
    ) -> int:
        coords = rounded_box_coords(
            x1,
            y1,
            x2,
            y2,
            radius=8 if self.PAR.ops.rounded_boxes else 0,
        )
        if isinstance(iid, int):
            self.coords(iid, coords)
            self.itemconfig(iid, fill=fill, outline=outline, state=state, tags=tags)
        else:
            if self.hidd_boxes:
                iid = self.hidd_boxes.pop()
                self.coords(iid, coords)
                self.itemconfig(iid, fill=fill, outline=outline, state=state, tags=tags)
            else:
                iid = self.create_polygon(coords, fill=fill, outline=outline, state=state, tags=tags, smooth=True)
            self.disp_boxes.add(iid)
        return iid

    def hide_box(self, item: int) -> None:
        if isinstance(item, int):
            self.disp_boxes.discard(item)
            self.hidd_boxes.add(item)
            self.itemconfig(item, state="hidden")

    def get_cell_dimensions(self, datacn: int) -> tuple[int, int]:
        txt = self.get_valid_cell_data_as_str(datacn, fix=False)
        if txt:
            self.MT.txt_measure_canvas.itemconfig(
                self.MT.txt_measure_canvas_text,
                text=txt,
                font=self.PAR.ops.header_font,
            )
            b = self.MT.txt_measure_canvas.bbox(self.MT.txt_measure_canvas_text)
            w = b[2] - b[0] + 7
            h = b[3] - b[1] + 5
        else:
            w = self.MT.min_column_width
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
        qfont = self.PAR.ops.header_font
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
                    if txt := self.MT.get_valid_cell_data_as_str(datarn, datacn, get_displayed=True):
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
        elif h > self.MT.max_header_height:
            h = int(self.MT.max_header_height)
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
        w = self.MT.min_column_width
        datacn = col if self.MT.all_columns_displayed else self.MT.displayed_columns[col]
        # header
        hw, hh_ = self.get_cell_dimensions(datacn)
        # table
        if self.MT.data:
            if self.MT.all_rows_displayed:
                if visible_only:
                    iterable = range(*self.MT.visible_text_rows)
                else:
                    iterable = range(0, len(self.MT.data))
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
            qfont = self.PAR.ops.table_font
            for datarn in iterable:
                if txt := self.MT.get_valid_cell_data_as_str(datarn, datacn, get_displayed=True):
                    qconf(qtxtm, text=txt, font=qfont)
                    b = qbbox(qtxtm)
                    if (
                        self.MT.get_cell_kwargs(datarn, datacn, key="dropdown")
                        or self.MT.get_cell_kwargs(datarn, datacn, key="checkbox")
                    ) and (tw := b[2] - b[0] + qtxth + 7) > w:
                        w = tw
                    elif (tw := b[2] - b[0] + 7) > w:
                        w = tw
        if hw > w:
            w = hw
        if only_if_too_small and w < self.MT.col_positions[col + 1] - self.MT.col_positions[col]:
            w = self.MT.col_positions[col + 1] - self.MT.col_positions[col]
        if w <= self.MT.min_column_width:
            w = int(self.MT.min_column_width)
        elif w > self.MT.max_column_width:
            w = int(self.MT.max_column_width)
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
        if width <= self.MT.min_column_width:
            width = int(self.MT.min_column_width)
        elif width > self.MT.max_column_width:
            width = int(self.MT.max_column_width)
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
        c_2: str,
        c_3: str,
        selections: dict,
        datacn: int,
    ) -> tuple[str, bool]:
        redrawn = False
        kwargs = self.get_cell_kwargs(datacn, key="highlight")
        if kwargs:
            if kwargs[0] is not None:
                c_1 = kwargs[0] if kwargs[0].startswith("#") else color_map[kwargs[0]]
            if "columns" in selections and c in selections["columns"]:
                tf = (
                    self.PAR.ops.header_selected_columns_fg
                    if kwargs[1] is None or self.PAR.ops.display_selected_fg_over_highlights
                    else kwargs[1]
                )
                if kwargs[0] is not None:
                    fill = (
                        f"#{int((int(c_1[1:3], 16) + int(c_3[1:3], 16)) / 2):02X}"
                        + f"{int((int(c_1[3:5], 16) + int(c_3[3:5], 16)) / 2):02X}"
                        + f"{int((int(c_1[5:], 16) + int(c_3[5:], 16)) / 2):02X}"
                    )
            elif "cells" in selections and c in selections["cells"]:
                tf = (
                    self.PAR.ops.header_selected_cells_fg
                    if kwargs[1] is None or self.PAR.ops.display_selected_fg_over_highlights
                    else kwargs[1]
                )
                if kwargs[0] is not None:
                    fill = (
                        f"#{int((int(c_1[1:3], 16) + int(c_2[1:3], 16)) / 2):02X}"
                        + f"{int((int(c_1[3:5], 16) + int(c_2[3:5], 16)) / 2):02X}"
                        + f"{int((int(c_1[5:], 16) + int(c_2[5:], 16)) / 2):02X}"
                    )
            else:
                tf = self.PAR.ops.header_fg if kwargs[1] is None else kwargs[1]
                if kwargs[0] is not None:
                    fill = kwargs[0]
            if kwargs[0] is not None:
                redrawn = self.redraw_highlight(
                    fc + 1,
                    0,
                    sc,
                    self.current_height - 1,
                    fill=fill,
                    outline=(
                        self.PAR.ops.header_fg
                        if self.get_cell_kwargs(datacn, key="dropdown") and self.PAR.ops.show_dropdown_borders
                        else ""
                    ),
                    tag="hi",
                )
        elif not kwargs:
            if "columns" in selections and c in selections["columns"]:
                tf = self.PAR.ops.header_selected_columns_fg
            elif "cells" in selections and c in selections["cells"]:
                tf = self.PAR.ops.header_selected_cells_fg
            else:
                tf = self.PAR.ops.header_fg
        return tf, redrawn

    def redraw_highlight(
        self,
        x1: float,
        y1: float,
        x2: float,
        y2: float,
        fill: str,
        outline: str,
        tag: str | tuple[str],
    ) -> bool:
        coords = (x1, y1, x2, y2)
        if self.hidd_high:
            iid, showing = self.hidd_high.popitem()
            self.coords(iid, coords)
            if showing:
                self.itemconfig(iid, fill=fill, outline=outline)
            else:
                self.itemconfig(iid, fill=fill, outline=outline, tag=tag, state="normal")
        else:
            iid = self.create_rectangle(coords, fill=fill, outline=outline, tag=tag)
        self.disp_high[iid] = True
        return True

    def redraw_gridline(
        self,
        points: Sequence[float],
        fill: str,
        width: int,
        tag: str | tuple[str],
    ) -> None:
        if self.hidd_grid:
            t, sh = self.hidd_grid.popitem()
            self.coords(t, points)
            if sh:
                self.itemconfig(t, fill=fill, width=width, tag=tag)
            else:
                self.itemconfig(t, fill=fill, width=width, tag=tag, state="normal")
            self.disp_grid[t] = True
        else:
            self.disp_grid[self.create_line(points, fill=fill, width=width, tag=tag)] = True

    def redraw_dropdown(
        self,
        x1: float,
        y1: float,
        x2: float,
        y2: float,
        fill: str,
        outline: str,
        tag: str | tuple[str],
        draw_outline: bool = True,
        draw_arrow: bool = True,
        dd_is_open: bool = False,
    ) -> None:
        if draw_outline and self.PAR.ops.show_dropdown_borders:
            self.redraw_highlight(x1 + 1, y1 + 1, x2, y2, fill="", outline=self.PAR.ops.header_fg, tag=tag)
        if draw_arrow:
            mod = (self.MT.header_txt_height - 1) if self.MT.header_txt_height % 2 else self.MT.header_txt_height
            half_mod = mod / 2
            qtr_mod = mod / 4
            mid_y = (
                (self.MT.header_first_ln_ins - 1) if self.MT.header_first_ln_ins % 2 else self.MT.header_first_ln_ins
            )
            if dd_is_open:
                points = (
                    x2 - 3 - mod,
                    y1 + mid_y + qtr_mod,
                    x2 - 3 - half_mod,
                    y1 + mid_y - qtr_mod,
                    x2 - 3,
                    y1 + mid_y + qtr_mod,
                )
            else:
                points = (
                    x2 - 3 - mod,
                    y1 + mid_y - qtr_mod,
                    x2 - 3 - half_mod,
                    y1 + mid_y + qtr_mod,
                    x2 - 3,
                    y1 + mid_y - qtr_mod,
                )
            if self.hidd_dropdown:
                t, sh = self.hidd_dropdown.popitem()
                self.coords(t, points)
                if sh:
                    self.itemconfig(t, fill=fill)
                else:
                    self.itemconfig(t, fill=fill, tag=tag, state="normal")
                self.lift(t)
            else:
                t = self.create_polygon(
                    points,
                    fill=fill,
                    tag=tag,
                )
            self.disp_dropdown[t] = True

    def redraw_checkbox(
        self,
        x1: float,
        y1: float,
        x2: float,
        y2: float,
        fill: str,
        outline: str,
        tag: str | tuple[str],
        draw_check: bool = False,
    ) -> None:
        points = rounded_box_coords(x1, y1, x2, y2)
        if self.hidd_checkbox:
            t, sh = self.hidd_checkbox.popitem()
            self.coords(t, points)
            if sh:
                self.itemconfig(t, fill=outline, outline=fill)
            else:
                self.itemconfig(t, fill=outline, outline=fill, tag=tag, state="normal")
            self.lift(t)
        else:
            t = self.create_polygon(points, fill=outline, outline=fill, tag=tag, smooth=True)
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
                    self.itemconfig(t, fill=fill, outline=outline, tag=tag, state="normal")
                self.lift(t)
            else:
                t = self.create_polygon(points, fill=fill, outline=outline, tag=tag, smooth=True)
            self.disp_checkbox[t] = True

    def configure_scrollregion(self, last_col_line_pos: float) -> None:
        self.configure(
            scrollregion=(
                0,
                0,
                last_col_line_pos + self.PAR.ops.empty_horizontal + 2,
                self.current_height,
            )
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
    ) -> bool:
        try:
            self.configure_scrollregion(last_col_line_pos=last_col_line_pos)
        except Exception:
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
        self.visible_col_dividers = {}
        self.col_height_resize_bbox = (
            scrollpos_left,
            self.current_height - 2,
            x_stop,
            self.current_height,
        )
        draw_x = self.MT.col_positions[grid_start_col]
        yend = self.current_height - 5
        if (self.PAR.ops.show_vertical_grid or self.width_resizing_enabled) and col_pos_exists:
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
            self.redraw_gridline(points=points, fill=self.PAR.ops.header_grid_fg, width=1, tag="v")
        top = self.canvasy(0)
        c_2 = (
            self.PAR.ops.header_selected_cells_bg
            if self.PAR.ops.header_selected_cells_bg.startswith("#")
            else color_map[self.PAR.ops.header_selected_cells_bg]
        )
        c_3 = (
            self.PAR.ops.header_selected_columns_bg
            if self.PAR.ops.header_selected_columns_bg.startswith("#")
            else color_map[self.PAR.ops.header_selected_columns_bg]
        )
        font = self.PAR.ops.header_font
        selections = self.get_redraw_selections(text_start_col, grid_end_col)
        dd_coords = self.dropdown.get_coords()
        for c in range(text_start_col, text_end_col):
            draw_y = self.MT.header_first_ln_ins
            cleftgridln = self.MT.col_positions[c]
            crightgridln = self.MT.col_positions[c + 1]
            datacn = c if self.MT.all_columns_displayed else self.MT.displayed_columns[c]
            fill, dd_drawn = self.redraw_highlight_get_text_fg(
                cleftgridln, crightgridln, c, c_2, c_3, selections, datacn
            )

            if datacn in self.cell_options and "align" in self.cell_options[datacn]:
                align = self.cell_options[datacn]["align"]
            else:
                align = self.align

            kwargs = self.get_cell_kwargs(datacn, key="dropdown")
            if align == "w":
                draw_x = cleftgridln + 3
                if kwargs:
                    mw = crightgridln - cleftgridln - self.MT.header_txt_height - 2
                    self.redraw_dropdown(
                        cleftgridln,
                        0,
                        crightgridln,
                        self.current_height - 1,
                        fill=fill,
                        outline=fill,
                        tag="dd",
                        draw_outline=not dd_drawn,
                        draw_arrow=mw >= 5,
                        dd_is_open=dd_coords == c,
                    )
                else:
                    mw = crightgridln - cleftgridln - 1

            elif align == "e":
                if kwargs:
                    mw = crightgridln - cleftgridln - self.MT.header_txt_height - 2
                    draw_x = crightgridln - 5 - self.MT.header_txt_height
                    self.redraw_dropdown(
                        cleftgridln,
                        0,
                        crightgridln,
                        self.current_height - 1,
                        fill=fill,
                        outline=fill,
                        tag="dd",
                        draw_outline=not dd_drawn,
                        draw_arrow=mw >= 5,
                        dd_is_open=dd_coords == c,
                    )
                else:
                    mw = crightgridln - cleftgridln - 1
                    draw_x = crightgridln - 3

            elif align == "center":
                if kwargs:
                    mw = crightgridln - cleftgridln - self.MT.header_txt_height - 2
                    draw_x = cleftgridln + ceil((crightgridln - cleftgridln - self.MT.header_txt_height) / 2)
                    self.redraw_dropdown(
                        cleftgridln,
                        0,
                        crightgridln,
                        self.current_height - 1,
                        fill=fill,
                        outline=fill,
                        tag="dd",
                        draw_outline=not dd_drawn,
                        draw_arrow=mw >= 5,
                        dd_is_open=dd_coords == c,
                    )
                else:
                    mw = crightgridln - cleftgridln - 1
                    draw_x = cleftgridln + floor((crightgridln - cleftgridln) / 2)
            if not kwargs:
                kwargs = self.get_cell_kwargs(datacn, key="checkbox")
                if kwargs and mw > self.MT.header_txt_height + 1:
                    box_w = self.MT.header_txt_height + 1
                    if align == "w":
                        draw_x += box_w + 3
                    elif align == "center":
                        draw_x += ceil(box_w / 2) + 1
                    mw -= box_w + 3
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
                        cleftgridln + self.MT.header_txt_height + 3,
                        self.MT.header_txt_height + 3,
                        fill=fill if kwargs["state"] == "normal" else self.PAR.ops.header_grid_fg,
                        outline="",
                        tag="cb",
                        draw_check=draw_check,
                    )
            lns = self.get_valid_cell_data_as_str(datacn, fix=False)
            if not lns:
                continue
            lns = lns.split("\n")
            if mw > self.MT.header_txt_width and not (
                (align == "w" and draw_x > scrollpos_right)
                or (align == "e" and cleftgridln + 5 > scrollpos_right)
                or (align == "center" and cleftgridln + 5 > scrollpos_right)
            ):
                for txt in islice(
                    lns,
                    self.lines_start_at if self.lines_start_at < len(lns) else len(lns) - 1,
                    None,
                ):
                    if draw_y > top:
                        if self.hidd_text:
                            iid, showing = self.hidd_text.popitem()
                            self.coords(iid, draw_x, draw_y)
                            if showing:
                                self.itemconfig(
                                    iid,
                                    text=txt,
                                    fill=fill,
                                    font=font,
                                    anchor=align,
                                )
                            else:
                                self.itemconfig(
                                    iid,
                                    text=txt,
                                    fill=fill,
                                    font=font,
                                    anchor=align,
                                    state="normal",
                                )
                            self.tag_raise(iid)
                        else:
                            iid = self.create_text(
                                draw_x,
                                draw_y,
                                text=txt,
                                fill=fill,
                                font=font,
                                anchor=align,
                                tag="t",
                            )
                        self.disp_text[iid] = True
                        wd = self.bbox(iid)
                        wd = wd[2] - wd[0]
                        if wd > mw:
                            if align == "w":
                                txt = txt[: int(len(txt) * (mw / wd))]
                                self.itemconfig(iid, text=txt)
                                wd = self.bbox(iid)
                                while wd[2] - wd[0] > mw:
                                    txt = txt[:-1]
                                    self.itemconfig(iid, text=txt)
                                    wd = self.bbox(iid)
                            elif align == "e":
                                txt = txt[len(txt) - int(len(txt) * (mw / wd)) :]
                                self.itemconfig(iid, text=txt)
                                wd = self.bbox(iid)
                                while wd[2] - wd[0] > mw:
                                    txt = txt[1:]
                                    self.itemconfig(iid, text=txt)
                                    wd = self.bbox(iid)
                            elif align == "center":
                                self.c_align_cyc = cycle(self.centre_alignment_text_mod_indexes)
                                tmod = ceil((len(txt) - int(len(txt) * (mw / wd))) / 2)
                                txt = txt[tmod - 1 : -tmod]
                                self.itemconfig(iid, text=txt)
                                wd = self.bbox(iid)
                                while wd[2] - wd[0] > mw:
                                    txt = txt[next(self.c_align_cyc)]
                                    self.itemconfig(iid, text=txt)
                                    wd = self.bbox(iid)
                                self.coords(iid, draw_x, draw_y)
                    draw_y += self.MT.header_xtra_lines_increment
                    if draw_y - 1 > self.current_height:
                        break
        for dct in (self.hidd_text, self.hidd_high, self.hidd_grid, self.hidd_dropdown, self.hidd_checkbox):
            for iid, showing in dct.items():
                if showing:
                    self.itemconfig(iid, state="hidden")
                    dct[iid] = False
        return True

    def get_redraw_selections(self, startc: int, endc: int) -> dict[str, set[int]]:
        d = defaultdict(set)
        for item, box in self.MT.get_selection_items():
            r1, c1, r2, c2 = box.coords
            for c in range(startc, endc):
                if c1 <= c and c2 > c:
                    d[box.type_ if box.type_ != "rows" else "cells"].add(c)
        return d

    def open_cell(self, event: object = None, ignore_existing_editor: bool = False) -> None:
        if not self.MT.anything_selected() or (not ignore_existing_editor and self.text_editor.open):
            return
        if not self.MT.selected:
            return
        c = self.MT.selected.column
        datacn = self.MT.datacn(c)
        if self.get_cell_kwargs(datacn, key="readonly"):
            return
        elif self.get_cell_kwargs(datacn, key="dropdown") or self.get_cell_kwargs(datacn, key="checkbox"):
            if self.MT.event_opens_dropdown_or_checkbox(event):
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
        event: object = None,
        c: int = 0,
        text: object = None,
        state: str = "normal",
        dropdown: bool = False,
    ) -> bool:
        text = None
        extra_func_key = "??"
        if event is None or self.MT.event_opens_dropdown_or_checkbox(event):
            if event is not None:
                if hasattr(event, "keysym") and event.keysym == "Return":
                    extra_func_key = "Return"
                elif hasattr(event, "keysym") and event.keysym == "F2":
                    extra_func_key = "F2"
            text = self.get_cell_data(self.MT.datacn(c), none_to_empty_str=True, redirect_int=True)
        elif event is not None and (
            (hasattr(event, "keysym") and event.keysym == "BackSpace") or event.keycode in (8, 855638143)
        ):
            extra_func_key = "BackSpace"
            text = ""
        elif event is not None and (
            (hasattr(event, "char") and event.char.isalpha())
            or (hasattr(event, "char") and event.char.isdigit())
            or (hasattr(event, "char") and event.char in symbols_set)
        ):
            extra_func_key = event.char
            text = event.char
        else:
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
        text = "" if text is None else text
        if self.PAR.ops.cell_auto_resize_enabled:
            if self.height_resizing_enabled:
                self.set_height_of_header_to_text(text)
            self.set_col_width_run_binding(c)
        if self.text_editor.open and c == self.text_editor.column:
            self.text_editor.set_text(self.text_editor.get() + "" if not isinstance(text, str) else text)
            return
        if self.text_editor.open:
            self.hide_text_editor()
        if not self.MT.see(r=0, c=c, keep_yscroll=True, check_cell_visibility=True):
            self.MT.refresh()
        x = self.MT.col_positions[c] + 1
        y = 0
        w = self.MT.col_positions[c + 1] - x
        h = self.current_height + 1
        if text is None:
            text = self.get_cell_data(self.MT.datacn(c), none_to_empty_str=True, redirect_int=True)
        bg, fg = self.PAR.ops.header_bg, self.PAR.ops.header_fg
        kwargs = {
            "menu_kwargs": DotDict(
                {
                    "font": self.PAR.ops.table_font,
                    "foreground": self.PAR.ops.popup_menu_fg,
                    "background": self.PAR.ops.popup_menu_bg,
                    "activebackground": self.PAR.ops.popup_menu_highlight_bg,
                    "activeforeground": self.PAR.ops.popup_menu_highlight_fg,
                }
            ),
            "sheet_ops": self.PAR.ops,
            "border_color": self.PAR.ops.header_selected_columns_bg,
            "text": text,
            "state": state,
            "width": w,
            "height": h,
            "show_border": True,
            "bg": bg,
            "fg": fg,
            "align": self.get_cell_align(c),
            "c": c,
        }
        if not self.text_editor.window:
            self.text_editor.window = TextEditor(self, newline_binding=self.text_editor_newline_binding)
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
    def text_editor_newline_binding(self, event: object = None, check_lines: bool = True) -> None:
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
            new_height = curr_height + self.MT.header_xtra_lines_increment
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
            self.text_editor.tktext.config(font=self.PAR.ops.header_font)
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

    def hide_text_editor(self, reason: None | str = None) -> None:
        if self.text_editor.open:
            for binding in text_editor_to_unbind:
                self.text_editor.tktext.unbind(binding)
            self.itemconfig(self.text_editor.canvas_id, state="hidden")
            self.text_editor.open = False
        if reason == "Escape":
            self.focus_set()

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
            return
        # setting cell data with text editor value
        text_editor_value = self.text_editor.get()
        c = self.text_editor.column
        datacn = c if self.MT.all_columns_displayed else self.MT.displayed_columns[c]
        event_data = event_dict(
            name="end_edit_header",
            sheet=self.PAR.name,
            widget=self,
            cells_header={datacn: self.get_cell_data(datacn)},
            key=event.keysym,
            value=text_editor_value,
            loc=c,
            column=c,
            boxes=self.MT.get_boxes(),
            selected=self.MT.selected,
        )
        edited = False
        set_data = partial(
            self.set_cell_data_undo,
            c=c,
            datacn=datacn,
            check_input_valid=False,
        )
        if self.MT.edit_validation_func:
            text_editor_value = self.MT.edit_validation_func(event_data)
            if text_editor_value is not None and self.input_valid_for_cell(datacn, text_editor_value):
                edited = set_data(value=text_editor_value)
        elif self.input_valid_for_cell(datacn, text_editor_value):
            edited = set_data(value=text_editor_value)
        if edited:
            try_binding(self.extra_end_edit_cell_func, event_data)
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
                win_h += (
                    self.MT.header_first_ln_ins + (v_numlines * self.MT.header_xtra_lines_increment) + 5
                )  # end of cell
            else:
                win_h += self.MT.min_header_height
            if i == 5:
                break
        if win_h > 500:
            win_h = 500
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
        dd_window: object,
        event: dict,
        modified_func: Callable | None,
    ) -> None:
        if modified_func:
            modified_func(event)
        dd_window.search_and_see(event)

    def open_dropdown_window(self, c: int, event: object = None) -> None:
        self.hide_text_editor("Escape")
        kwargs = self.get_cell_kwargs(self.MT.datacn(c), key="dropdown")
        if kwargs["state"] == "normal":
            if not self.open_text_editor(event=event, c=c, dropdown=True):
                return
        win_h, anchor = self.get_dropdown_height_anchor(c)
        win_w = self.MT.col_positions[c + 1] - self.MT.col_positions[c] + 1
        ypos = self.current_height - 1
        reset_kwargs = {
            "r": 0,
            "c": c,
            "width": win_w,
            "height": win_h,
            "font": self.PAR.ops.header_font,
            "ops": self.PAR.ops,
            "outline_color": self.PAR.ops.popup_menu_fg,
            "align": self.get_cell_align(c),
            "values": kwargs["values"],
        }
        if self.dropdown.window:
            self.dropdown.window.reset(**reset_kwargs)
            self.itemconfig(self.dropdown.canvas_id, state="normal")
            self.coords(self.dropdown.canvas_id, self.MT.col_positions[c], ypos)
            self.dropdown.window.tkraise()
        else:
            self.dropdown.window = self.PAR.dropdown_class(
                self.winfo_toplevel(),
                **reset_kwargs,
                single_index="c",
                close_dropdown_window=self.close_dropdown_window,
                search_function=kwargs["search_function"],
                arrowkey_RIGHT=self.MT.arrowkey_RIGHT,
                arrowkey_LEFT=self.MT.arrowkey_LEFT,
            )
            self.dropdown.canvas_id = self.create_window(
                (self.MT.col_positions[c], ypos),
                window=self.dropdown.window,
                anchor=anchor,
            )
        if kwargs["state"] == "normal":
            self.text_editor.tktext.bind(
                "<<TextModified>>",
                lambda _x: self.dropdown_text_editor_modified(
                    self.dropdown.window,
                    event_dict(
                        name="header_dropdown_modified",
                        sheet=self.PAR.name,
                        value=self.text_editor.get(),
                        loc=c,
                        column=c,
                        boxes=self.MT.get_boxes(),
                        selected=self.MT.selected,
                    ),
                    kwargs["modified_function"],
                ),
            )
            self.update_idletasks()
            try:
                self.after(1, lambda: self.text_editor.tktext.focus())
                self.after(2, self.text_editor.window.scroll_to_bottom())
            except Exception:
                return
            redraw = False
        else:
            self.dropdown.window.bind("<FocusOut>", lambda _x: self.close_dropdown_window(c))
            self.update_idletasks()
            self.dropdown.window.focus_set()
            redraw = True
        self.dropdown.open = True
        if redraw:
            self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=False, redraw_table=False)

    def close_dropdown_window(
        self,
        c: None | int = None,
        selection: object = None,
        redraw: bool = True,
    ) -> None:
        if c is not None and selection is not None:
            datacn = c if self.MT.all_columns_displayed else self.MT.displayed_columns[c]
            kwargs = self.get_cell_kwargs(datacn, key="dropdown")
            pre_edit_value = self.get_cell_data(datacn)
            edited = False
            event_data = event_dict(
                name="end_edit_header",
                sheet=self.PAR.name,
                widget=self,
                cells_header={datacn: pre_edit_value},
                key="??",
                value=selection,
                loc=c,
                column=c,
                boxes=self.MT.get_boxes(),
                selected=self.MT.selected,
            )
            if kwargs["select_function"] is not None:
                kwargs["select_function"](event_data)
            if self.MT.edit_validation_func:
                selection = self.MT.edit_validation_func(event_data)
                if selection is not None:
                    edited = self.set_cell_data_undo(c, datacn=datacn, value=selection, redraw=not redraw)
            else:
                edited = self.set_cell_data_undo(c, datacn=datacn, value=selection, redraw=not redraw)
            if edited:
                try_binding(self.extra_end_edit_cell_func, event_data)
            self.focus_set()
            self.MT.recreate_all_selection_boxes()
        self.hide_text_editor_and_dropdown(redraw=redraw)

    def hide_text_editor_and_dropdown(self, redraw: bool = True) -> None:
        self.hide_text_editor("Escape")
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
        value: object = "",
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
            edited = self.MT.set_cell_data_undo(r=self.MT._headers, c=c, datacn=datacn, value=value, undo=True)
        else:
            self.fix_header(datacn)
            if not check_input_valid or self.input_valid_for_cell(datacn, value):
                if self.MT.undo_enabled and undo:
                    self.MT.undo_stack.append(pickled_event_dict(event_data))
                self.set_cell_data(datacn=datacn, value=value)
                edited = True
        if edited and cell_resize and self.PAR.ops.cell_auto_resize_enabled:
            if self.height_resizing_enabled:
                self.set_height_of_header_to_text(self.get_valid_cell_data_as_str(datacn, fix=False))
            self.set_col_width_run_binding(c)
        if redraw:
            self.MT.refresh()
        if edited:
            self.MT.sheet_modified(event_data)
        return edited

    def set_cell_data(self, datacn: int | None = None, value: object = "") -> None:
        if isinstance(self.MT._headers, int):
            self.MT.set_cell_data(datarn=self.MT._headers, datacn=datacn, value=value)
        else:
            self.fix_header(datacn)
            if self.get_cell_kwargs(datacn, key="checkbox"):
                self.MT._headers[datacn] = try_to_bool(value)
            else:
                self.MT._headers[datacn] = value

    def input_valid_for_cell(self, datacn: int, value: object, check_readonly: bool = True) -> bool:
        if check_readonly and self.get_cell_kwargs(datacn, key="readonly"):
            return False
        if self.get_cell_kwargs(datacn, key="checkbox"):
            return is_bool_like(value)
        if self.cell_equal_to(datacn, value):
            return False
        kwargs = self.get_cell_kwargs(datacn, key="dropdown")
        if kwargs and kwargs["validate_input"] and value not in kwargs["values"]:
            return False
        return True

    def cell_equal_to(self, datacn: int, value: object) -> bool:
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
    ) -> object:
        if get_displayed:
            return self.get_valid_cell_data_as_str(datacn, fix=False)
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

    def get_valid_cell_data_as_str(self, datacn: int, fix: bool = True) -> str:
        kwargs = self.get_cell_kwargs(datacn, key="dropdown")
        if kwargs:
            if kwargs["text"] is not None:
                return f"{kwargs['text']}"
        else:
            kwargs = self.get_cell_kwargs(datacn, key="checkbox")
            if kwargs:
                return f"{kwargs['text']}"
        if isinstance(self.MT._headers, int):
            return self.MT.get_valid_cell_data_as_str(self.MT._headers, datacn, get_displayed=True)
        if fix:
            self.fix_header(datacn)
        try:
            value = "" if self.MT._headers[datacn] is None else f"{self.MT._headers[datacn]}"
        except Exception:
            value = ""
        if not value and self.PAR.ops.show_default_header_for_empty:
            value = get_n2a(datacn, self.default_header)
        return value

    def get_value_for_empty_cell(self, datacn: int, c_ops: bool = True) -> object:
        if self.get_cell_kwargs(datacn, key="checkbox", cell=c_ops):
            return False
        kwargs = self.get_cell_kwargs(datacn, key="dropdown", cell=c_ops)
        if kwargs and kwargs["validate_input"] and kwargs["values"]:
            return kwargs["values"][0]
        return ""

    def get_empty_header_seq(self, end: int, start: int = 0, c_ops: bool = True) -> list[object]:
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
            event_data = event_dict(
                name="end_edit_header",
                sheet=self.PAR.name,
                widget=self,
                cells_header={datacn: pre_edit_value},
                key="??",
                value=value,
                loc=c,
                column=c,
                boxes=self.MT.get_boxes(),
                selected=self.MT.selected,
            )
            if kwargs["check_function"] is not None:
                kwargs["check_function"](event_data)
            try_binding(self.extra_end_edit_cell_func, event_data)
        if redraw:
            self.MT.refresh()

    def get_cell_kwargs(self, datacn: int, key: Hashable = "dropdown", cell: bool = True) -> dict:
        if cell and datacn in self.cell_options and key in self.cell_options[datacn]:
            return self.cell_options[datacn][key]
        return {}
