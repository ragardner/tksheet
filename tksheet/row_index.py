from __future__ import annotations

import tkinter as tk
from collections import defaultdict
from collections.abc import (
    Callable,
    Generator,
    Hashable,
    Iterator,
    Sequence,
)
from functools import (
    partial,
)
from itertools import (
    chain,
    cycle,
    islice,
    repeat,
)
from math import (
    ceil,
    floor,
)
from operator import (
    itemgetter,
)
from typing import Literal

from .colors import (
    color_map,
)
from .formatters import (
    is_bool_like,
    try_to_bool,
)
from .functions import (
    consecutive_chunks,
    event_dict,
    get_n2a,
    is_contiguous,
    new_tk_event,
    num2alpha,
    pickled_event_dict,
    rounded_box_coords,
    try_binding,
)
from .other_classes import (
    DotDict,
    DraggedRowColumn,
    DropdownStorage,
    Node,
    TextEditorStorage,
)
from .text_editor import (
    TextEditor,
)
from .vars import (
    rc_binding,
    symbols_set,
    text_editor_close_bindings,
    text_editor_newline_bindings,
    text_editor_to_unbind,
)


class RowIndex(tk.Canvas):
    def __init__(self, *args, **kwargs):
        tk.Canvas.__init__(
            self,
            kwargs["parent"],
            background=kwargs["parent"].ops.index_bg,
            highlightthickness=0,
        )
        self.PAR = kwargs["parent"]
        self.MT = None  # is set from within MainTable() __init__
        self.CH = None  # is set from within MainTable() __init__
        self.TL = None  # is set from within TopLeftRectangle() __init__
        self.popup_menu_loc = None
        self.extra_begin_edit_cell_func = None
        self.extra_end_edit_cell_func = None
        self.b1_pressed_loc = None
        self.closed_dropdown = None
        self.centre_alignment_text_mod_indexes = (slice(1, None), slice(None, -1))
        self.c_align_cyc = cycle(self.centre_alignment_text_mod_indexes)
        self.being_drawn_item = None
        self.extra_motion_func = None
        self.extra_b1_press_func = None
        self.extra_b1_motion_func = None
        self.extra_b1_release_func = None
        self.extra_rc_func = None
        self.selection_binding_func = None
        self.shift_selection_binding_func = None
        self.ctrl_selection_binding_func = None
        self.drag_selection_binding_func = None
        self.ri_extra_begin_drag_drop_func = None
        self.ri_extra_end_drag_drop_func = None
        self.extra_double_b1_func = None
        self.row_height_resize_func = None
        self.new_row_width = 0
        self.cell_options = {}
        self.drag_and_drop_enabled = False
        self.dragged_row = None
        self.width_resizing_enabled = False
        self.height_resizing_enabled = False
        self.double_click_resizing_enabled = False
        self.row_selection_enabled = False
        self.rc_insert_row_enabled = False
        self.rc_delete_row_enabled = False
        self.edit_cell_enabled = False
        self.visible_row_dividers = {}
        self.row_width_resize_bbox = tuple()
        self.rsz_w = None
        self.rsz_h = None
        self.currently_resizing_width = False
        self.currently_resizing_height = False
        self.ri_rc_popup_menu = None
        self.dropdown = DropdownStorage()
        self.text_editor = TextEditorStorage()

        self.disp_text = {}
        self.disp_high = {}
        self.disp_grid = {}
        self.disp_fill_sels = {}
        self.disp_bord_sels = {}
        self.disp_resize_lines = {}
        self.disp_dropdown = {}
        self.disp_checkbox = {}
        self.disp_tree_arrow = {}
        self.disp_boxes = set()
        self.hidd_text = {}
        self.hidd_high = {}
        self.hidd_grid = {}
        self.hidd_fill_sels = {}
        self.hidd_bord_sels = {}
        self.hidd_resize_lines = {}
        self.hidd_dropdown = {}
        self.hidd_checkbox = {}
        self.hidd_tree_arrow = {}
        self.hidd_boxes = set()

        self.align = kwargs["row_index_align"]
        self.default_index = kwargs["default_row_index"].lower()

        self.tree_reset()
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
        else:
            self.unbind("<Motion>")
            self.unbind("<ButtonPress-1>")
            self.unbind("<B1-Motion>")
            self.unbind("<ButtonRelease-1>")
            self.unbind("<Double-Button-1>")
            self.unbind(rc_binding)

    def tree_reset(self) -> None:
        # treeview mode
        self.tree = {}
        self.tree_open_ids = set()
        self.tree_rns = {}
        if self.MT:
            self.MT.displayed_rows = []
            self.MT._row_index = []
            self.MT.data = []
            self.MT.row_positions = [0]
            self.MT.saved_row_heights = {}

    def set_width(self, new_width: int, set_TL: bool = False, recreate_selection_boxes: bool = True) -> None:
        self.current_width = new_width
        try:
            self.config(width=new_width)
        except Exception:
            return
        if set_TL:
            self.TL.set_dimensions(new_w=new_width, recreate_selection_boxes=recreate_selection_boxes)

    def rc(self, event: object) -> None:
        self.mouseclick_outside_editor_or_dropdown_all_canvases(inside=True)
        self.focus_set()
        popup_menu = None
        if self.MT.identify_row(y=event.y, allow_end=False) is None:
            self.MT.deselect("all")
            r = len(self.MT.col_positions) - 1
            if self.MT.rc_popup_menus_enabled:
                popup_menu = self.MT.empty_rc_popup_menu
        elif self.row_selection_enabled and not self.currently_resizing_width and not self.currently_resizing_height:
            r = self.MT.identify_row(y=event.y)
            if r < len(self.MT.row_positions) - 1:
                if self.MT.row_selected(r):
                    if self.MT.rc_popup_menus_enabled:
                        popup_menu = self.ri_rc_popup_menu
                else:
                    if self.MT.single_selection_enabled and self.MT.rc_select_enabled:
                        self.select_row(r, redraw=True)
                    elif self.MT.toggle_selection_enabled and self.MT.rc_select_enabled:
                        self.toggle_select_row(r, redraw=True)
                    if self.MT.rc_popup_menus_enabled:
                        popup_menu = self.ri_rc_popup_menu
        try_binding(self.extra_rc_func, event)
        if popup_menu is not None:
            self.popup_menu_loc = r
            popup_menu.tk_popup(event.x_root, event.y_root)

    def ctrl_b1_press(self, event: object) -> None:
        self.mouseclick_outside_editor_or_dropdown_all_canvases(inside=True)
        if (
            (self.drag_and_drop_enabled or self.row_selection_enabled)
            and self.MT.ctrl_select_enabled
            and self.rsz_h is None
            and self.rsz_w is None
        ):
            r = self.MT.identify_row(y=event.y)
            if r < len(self.MT.row_positions) - 1:
                r_selected = self.MT.row_selected(r)
                if not r_selected and self.row_selection_enabled:
                    self.being_drawn_item = True
                    self.being_drawn_item = self.add_selection(r, set_as_current=True, run_binding_func=False)
                    self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
                    sel_event = self.MT.get_select_event(being_drawn_item=self.being_drawn_item)
                    try_binding(self.ctrl_selection_binding_func, sel_event)
                    self.PAR.emit_event("<<SheetSelect>>", data=sel_event)
                elif r_selected:
                    self.MT.deselect(r=r)
        elif not self.MT.ctrl_select_enabled:
            self.b1_press(event)

    def ctrl_shift_b1_press(self, event: object) -> None:
        self.mouseclick_outside_editor_or_dropdown_all_canvases(inside=True)
        y = event.y
        r = self.MT.identify_row(y=y)
        if (
            (self.drag_and_drop_enabled or self.row_selection_enabled)
            and self.MT.ctrl_select_enabled
            and self.rsz_h is None
            and self.rsz_w is None
        ):
            if r < len(self.MT.row_positions) - 1:
                r_selected = self.MT.row_selected(r)
                if not r_selected and self.row_selection_enabled:
                    if self.MT.selected and self.MT.selected.type_ == "rows":
                        self.being_drawn_item = self.MT.recreate_selection_box(
                            *self.get_shift_select_box(r, self.MT.selected.row),
                            fill_iid=self.MT.selected.fill_iid,
                        )
                    else:
                        self.being_drawn_item = self.add_selection(
                            r,
                            run_binding_func=False,
                            set_as_current=True,
                        )
                    self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
                    sel_event = self.MT.get_select_event(being_drawn_item=self.being_drawn_item)
                    try_binding(self.ctrl_selection_binding_func, sel_event)
                    self.PAR.emit_event("<<SheetSelect>>", data=sel_event)
                elif r_selected:
                    self.dragged_row = DraggedRowColumn(
                        dragged=r,
                        to_move=sorted(self.MT.get_selected_rows()),
                    )
        elif not self.MT.ctrl_select_enabled:
            self.shift_b1_press(event)

    def shift_b1_press(self, event: object) -> None:
        self.mouseclick_outside_editor_or_dropdown_all_canvases(inside=True)
        y = event.y
        r = self.MT.identify_row(y=y)
        if (self.drag_and_drop_enabled or self.row_selection_enabled) and self.rsz_h is None and self.rsz_w is None:
            if r < len(self.MT.row_positions) - 1:
                r_selected = self.MT.row_selected(r)
                if not r_selected and self.row_selection_enabled:
                    if self.MT.selected and self.MT.selected.type_ == "rows":
                        r_to_sel, c_to_sel = self.MT.selected.row, self.MT.selected.column
                        self.MT.deselect("all", redraw=False)
                        self.being_drawn_item = self.MT.create_selection_box(
                            *self.get_shift_select_box(r, r_to_sel), "rows"
                        )
                        self.MT.set_currently_selected(r_to_sel, c_to_sel, self.being_drawn_item)
                    else:
                        self.being_drawn_item = self.select_row(r, run_binding_func=False)
                    self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
                    sel_event = self.MT.get_select_event(being_drawn_item=self.being_drawn_item)
                    try_binding(self.shift_selection_binding_func, sel_event)
                    self.PAR.emit_event("<<SheetSelect>>", data=sel_event)
                elif r_selected:
                    self.dragged_row = DraggedRowColumn(
                        dragged=r,
                        to_move=sorted(self.MT.get_selected_rows()),
                    )

    def get_shift_select_box(self, r: int, min_r: int) -> tuple[int, int, int, int, str]:
        if r > min_r:
            return min_r, 0, r + 1, len(self.MT.col_positions) - 1
        elif r < min_r:
            return r, 0, min_r + 1, len(self.MT.col_positions) - 1

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

    def check_mouse_position_height_resizers(self, x: int, y: int) -> int | None:
        for r, (x1, y1, x2, y2) in self.visible_row_dividers.items():
            if x >= x1 and y >= y1 and x <= x2 and y <= y2:
                return r

    def mouse_motion(self, event: object) -> None:
        if not self.currently_resizing_height and not self.currently_resizing_width:
            x = self.canvasx(event.x)
            y = self.canvasy(event.y)
            mouse_over_resize = False
            mouse_over_selected = False
            if self.height_resizing_enabled and not mouse_over_resize:
                r = self.check_mouse_position_height_resizers(x, y)
                if r is not None:
                    self.rsz_h, mouse_over_resize = r, True
                    if self.MT.current_cursor != "sb_v_double_arrow":
                        self.config(cursor="sb_v_double_arrow")
                        self.MT.current_cursor = "sb_v_double_arrow"
                else:
                    self.rsz_h = None
            if self.width_resizing_enabled and not mouse_over_resize:
                try:
                    x1, y1, x2, y2 = (
                        self.row_width_resize_bbox[0],
                        self.row_width_resize_bbox[1],
                        self.row_width_resize_bbox[2],
                        self.row_width_resize_bbox[3],
                    )
                    if x >= x1 and y >= y1 and x <= x2 and y <= y2:
                        self.rsz_w, mouse_over_resize = True, True
                        if self.MT.current_cursor != "sb_h_double_arrow":
                            self.config(cursor="sb_h_double_arrow")
                            self.MT.current_cursor = "sb_h_double_arrow"
                    else:
                        self.rsz_w = None
                except Exception:
                    self.rsz_w = None
            if not mouse_over_resize:
                if self.MT.row_selected(self.MT.identify_row(event, allow_end=False)):
                    mouse_over_selected = True
                    if self.MT.current_cursor != "hand2":
                        self.config(cursor="hand2")
                        self.MT.current_cursor = "hand2"
            if not mouse_over_resize and not mouse_over_selected:
                self.MT.reset_mouse_motion_creations()
        try_binding(self.extra_motion_func, event)

    def double_b1(self, event: object):
        self.mouseclick_outside_editor_or_dropdown_all_canvases(inside=True)
        self.focus_set()
        if (
            self.double_click_resizing_enabled
            and self.height_resizing_enabled
            and self.rsz_h is not None
            and not self.currently_resizing_height
        ):
            row = self.rsz_h - 1
            old_height = self.MT.row_positions[self.rsz_h] - self.MT.row_positions[self.rsz_h - 1]
            new_height = self.set_row_height(row)
            self.MT.allow_auto_resize_rows = False
            self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
            if self.row_height_resize_func is not None and old_height != new_height:
                self.row_height_resize_func(
                    event_dict(
                        name="resize",
                        sheet=self.PAR.name,
                        resized_rows={row: {"old_size": old_height, "new_size": new_height}},
                    )
                )
        elif self.width_resizing_enabled and self.rsz_h is None and self.rsz_w is True:
            self.set_width_of_index_to_text()
        elif (self.row_selection_enabled or self.PAR.ops.treeview) and self.rsz_h is None and self.rsz_w is None:
            r = self.MT.identify_row(y=event.y)
            if r < len(self.MT.row_positions) - 1:
                iid = self.event_over_tree_arrow(r, self.canvasy(event.y), event.x)
                if self.row_selection_enabled:
                    if self.MT.single_selection_enabled:
                        self.select_row(r, redraw=iid is None)
                    elif self.MT.toggle_selection_enabled:
                        self.toggle_select_row(r, redraw=iid is None)
                datarn = r if self.MT.all_rows_displayed else self.MT.displayed_rows[r]
                if (
                    self.get_cell_kwargs(datarn, key="dropdown")
                    or self.get_cell_kwargs(datarn, key="checkbox")
                    or self.edit_cell_enabled
                ):
                    self.open_cell(event)
                elif iid is not None:
                    self.PAR.item(iid, open_=iid not in self.tree_open_ids)
        self.rsz_h = None
        self.mouse_motion(event)
        try_binding(self.extra_double_b1_func, event)

    def b1_press(self, event: object):
        self.MT.unbind("<MouseWheel>")
        self.focus_set()
        self.closed_dropdown = self.mouseclick_outside_editor_or_dropdown_all_canvases(inside=True)
        x = self.canvasx(event.x)
        y = self.canvasy(event.y)
        r = self.MT.identify_row(y=event.y)
        self.b1_pressed_loc = r
        if self.check_mouse_position_height_resizers(x, y) is None:
            self.rsz_h = None
        if (
            not x >= self.row_width_resize_bbox[0]
            and y >= self.row_width_resize_bbox[1]
            and x <= self.row_width_resize_bbox[2]
            and y <= self.row_width_resize_bbox[3]
        ):
            self.rsz_w = None
        if self.height_resizing_enabled and self.rsz_h is not None:
            self.currently_resizing_height = True
            y = self.MT.row_positions[self.rsz_h]
            line2y = self.MT.row_positions[self.rsz_h - 1]
            x1, y1, x2, y2 = self.MT.get_canvas_visible_area()
            self.create_resize_line(
                0,
                y,
                self.current_width,
                y,
                width=1,
                fill=self.PAR.ops.resizing_line_fg,
                tag="rhl",
            )
            self.MT.create_resize_line(x1, y, x2, y, width=1, fill=self.PAR.ops.resizing_line_fg, tag="rhl")
            self.create_resize_line(
                0,
                line2y,
                self.current_width,
                line2y,
                width=1,
                fill=self.PAR.ops.resizing_line_fg,
                tag="rhl2",
            )
            self.MT.create_resize_line(x1, line2y, x2, line2y, width=1, fill=self.PAR.ops.resizing_line_fg, tag="rhl2")
        elif self.width_resizing_enabled and self.rsz_h is None and self.rsz_w is True:
            self.currently_resizing_width = True
            x1, y1, x2, y2 = self.MT.get_canvas_visible_area()
            x = int(event.x)
            if x < self.MT.min_column_width:
                x = int(self.MT.min_column_width)
            self.new_row_width = x
            self.create_resize_line(x, y1, x, y2, width=1, fill=self.PAR.ops.resizing_line_fg, tag="rwl")
        elif self.MT.identify_row(y=event.y, allow_end=False) is None:
            self.MT.deselect("all")
        elif self.row_selection_enabled and self.rsz_h is None and self.rsz_w is None:
            r = self.MT.identify_row(y=event.y)
            if r < len(self.MT.row_positions) - 1:
                datarn = r if self.MT.all_rows_displayed else self.MT.displayed_rows[r]
                if (
                    self.MT.row_selected(r)
                    and not self.event_over_dropdown(r, datarn, event, y)
                    and not self.event_over_checkbox(r, datarn, event, y)
                ):
                    self.dragged_row = DraggedRowColumn(
                        dragged=r,
                        to_move=sorted(self.MT.get_selected_rows()),
                    )
                else:
                    if self.MT.single_selection_enabled:
                        self.being_drawn_item = True
                        self.being_drawn_item = self.select_row(r, redraw=True)
                    elif self.MT.toggle_selection_enabled:
                        self.toggle_select_row(r, redraw=True)
        try_binding(self.extra_b1_press_func, event)

    def b1_motion(self, event: object):
        x1, y1, x2, y2 = self.MT.get_canvas_visible_area()
        if self.height_resizing_enabled and self.rsz_h is not None and self.currently_resizing_height:
            y = self.canvasy(event.y)
            size = y - self.MT.row_positions[self.rsz_h - 1]
            if size >= self.MT.min_row_height and size < self.MT.max_row_height:
                self.hide_resize_and_ctrl_lines(ctrl_lines=False)
                line2y = self.MT.row_positions[self.rsz_h - 1]
                self.create_resize_line(
                    0,
                    y,
                    self.current_width,
                    y,
                    width=1,
                    fill=self.PAR.ops.resizing_line_fg,
                    tag="rhl",
                )
                self.MT.create_resize_line(x1, y, x2, y, width=1, fill=self.PAR.ops.resizing_line_fg, tag="rhl")
                self.create_resize_line(
                    0,
                    line2y,
                    self.current_width,
                    line2y,
                    width=1,
                    fill=self.PAR.ops.resizing_line_fg,
                    tag="rhl2",
                )
                self.MT.create_resize_line(
                    x1,
                    line2y,
                    x2,
                    line2y,
                    width=1,
                    fill=self.PAR.ops.resizing_line_fg,
                    tag="rhl2",
                )
        elif self.width_resizing_enabled and self.rsz_w is not None and self.currently_resizing_width:
            evx = event.x
            self.hide_resize_and_ctrl_lines(ctrl_lines=False)
            if evx > self.current_width:
                x = self.MT.canvasx(evx - self.current_width)
                if evx > self.MT.max_index_width:
                    evx = int(self.MT.max_index_width)
                    x = self.MT.canvasx(evx - self.current_width)
                self.new_row_width = evx
                self.MT.create_resize_line(x, y1, x, y2, width=1, fill=self.PAR.ops.resizing_line_fg, tag="rwl")
            else:
                x = evx
                if x < self.MT.min_column_width:
                    x = int(self.MT.min_column_width)
                self.new_row_width = x
                self.create_resize_line(x, y1, x, y2, width=1, fill=self.PAR.ops.resizing_line_fg, tag="rwl")
        if (
            self.drag_and_drop_enabled
            and self.row_selection_enabled
            and self.MT.anything_selected(exclude_cells=True, exclude_columns=True)
            and self.rsz_h is None
            and self.rsz_w is None
            and self.dragged_row is not None
        ):
            y = self.canvasy(event.y)
            if y > 0:
                self.show_drag_and_drop_indicators(
                    self.drag_and_drop_motion(event),
                    x1,
                    x2,
                    self.dragged_row.to_move,
                )
        elif (
            self.MT.drag_selection_enabled and self.row_selection_enabled and self.rsz_h is None and self.rsz_w is None
        ):
            need_redraw = False
            end_row = self.MT.identify_row(y=event.y)
            if end_row < len(self.MT.row_positions) - 1 and self.MT.selected:
                if self.MT.selected.type_ == "rows":
                    box = self.get_b1_motion_box(self.MT.selected.row, end_row)
                    if (
                        box is not None
                        and self.being_drawn_item is not None
                        and self.MT.coords_and_type(self.being_drawn_item) != box
                    ):
                        if box[2] - box[0] != 1:
                            self.being_drawn_item = self.MT.recreate_selection_box(
                                *box[:-1],
                                fill_iid=self.MT.selected.fill_iid,
                            )
                        else:
                            self.being_drawn_item = self.select_row(self.MT.selected.row, run_binding_func=False)
                        need_redraw = True
                        sel_event = self.MT.get_select_event(being_drawn_item=self.being_drawn_item)
                        try_binding(self.drag_selection_binding_func, sel_event)
                        self.PAR.emit_event("<<SheetSelect>>", data=sel_event)
                if self.scroll_if_event_offscreen(event):
                    need_redraw = True
            if need_redraw:
                self.MT.main_table_redraw_grid_and_text(redraw_header=False, redraw_row_index=True)
        try_binding(self.extra_b1_motion_func, event)

    def get_b1_motion_box(self, start_row: int, end_row: int) -> tuple[int, int, int, int, Literal["rows"]]:
        if end_row >= start_row:
            return start_row, 0, end_row + 1, len(self.MT.col_positions) - 1, "rows"
        elif end_row < start_row:
            return end_row, 0, start_row + 1, len(self.MT.col_positions) - 1, "rows"

    def ctrl_b1_motion(self, event: object) -> None:
        x1, y1, x2, y2 = self.MT.get_canvas_visible_area()
        if (
            self.drag_and_drop_enabled
            and self.row_selection_enabled
            and self.rsz_h is None
            and self.rsz_w is None
            and self.dragged_row is not None
            and self.MT.anything_selected(exclude_cells=True, exclude_columns=True)
        ):
            y = self.canvasy(event.y)
            if y > 0:
                self.show_drag_and_drop_indicators(
                    self.drag_and_drop_motion(event),
                    x1,
                    x2,
                    self.dragged_row.to_move,
                )
        elif (
            self.MT.ctrl_select_enabled
            and self.row_selection_enabled
            and self.MT.drag_selection_enabled
            and self.rsz_h is None
            and self.rsz_w is None
        ):
            need_redraw = False
            end_row = self.MT.identify_row(y=event.y)
            if end_row < len(self.MT.row_positions) - 1 and self.MT.selected:
                if self.MT.selected.type_ == "rows":
                    box = self.get_b1_motion_box(self.MT.selected.row, end_row)
                    if (
                        box is not None
                        and self.being_drawn_item is not None
                        and self.MT.coords_and_type(self.being_drawn_item) != box
                    ):
                        if box[2] - box[0] != 1:
                            self.being_drawn_item = self.MT.recreate_selection_box(
                                *box[:-1],
                                self.MT.selected.fill_iid,
                            )
                        else:
                            self.MT.hide_selection_box(self.MT.selected.fill_iid)
                            self.being_drawn_item = self.add_selection(box[0], run_binding_func=False)
                        need_redraw = True
                        sel_event = self.MT.get_select_event(being_drawn_item=self.being_drawn_item)
                        try_binding(self.drag_selection_binding_func, sel_event)
                        self.PAR.emit_event("<<SheetSelect>>", data=sel_event)
                if self.scroll_if_event_offscreen(event):
                    need_redraw = True
            if need_redraw:
                self.MT.main_table_redraw_grid_and_text(redraw_header=False, redraw_row_index=True)
        elif not self.MT.ctrl_select_enabled:
            self.b1_motion(event)

    def drag_and_drop_motion(self, event: object) -> float:
        y = event.y
        hend = self.winfo_height()
        ycheck = self.yview()
        if y >= hend - 0 and len(ycheck) > 1 and ycheck[1] < 1:
            if y >= hend + 15:
                self.MT.yview_scroll(2, "units")
                self.yview_scroll(2, "units")
            else:
                self.MT.yview_scroll(1, "units")
                self.yview_scroll(1, "units")
            self.fix_yview()
            self.MT.y_move_synced_scrolls("moveto", self.MT.yview()[0])
            self.MT.main_table_redraw_grid_and_text(redraw_row_index=True)
        elif y <= 0 and len(ycheck) > 1 and ycheck[0] > 0:
            if y >= -15:
                self.MT.yview_scroll(-1, "units")
                self.yview_scroll(-1, "units")
            else:
                self.MT.yview_scroll(-2, "units")
                self.yview_scroll(-2, "units")
            self.fix_yview()
            self.MT.y_move_synced_scrolls("moveto", self.MT.yview()[0])
            self.MT.main_table_redraw_grid_and_text(redraw_row_index=True)
        row = self.MT.identify_row(y=y)
        if row == len(self.MT.row_positions) - 1:
            row -= 1
        if row >= self.dragged_row.to_move[0] and row <= self.dragged_row.to_move[-1]:
            if is_contiguous(self.dragged_row.to_move):
                return self.MT.row_positions[self.dragged_row.to_move[0]]
            return self.MT.row_positions[row]
        elif row > self.dragged_row.to_move[-1]:
            return self.MT.row_positions[row + 1]
        return self.MT.row_positions[row]

    def show_drag_and_drop_indicators(
        self,
        ypos: float,
        x1: float,
        x2: float,
        rows: Sequence[int],
    ) -> None:
        self.hide_resize_and_ctrl_lines()
        self.create_resize_line(
            0,
            ypos,
            self.current_width,
            ypos,
            width=3,
            fill=self.PAR.ops.drag_and_drop_bg,
            tag="move_rows",
        )
        self.MT.create_resize_line(x1, ypos, x2, ypos, width=3, fill=self.PAR.ops.drag_and_drop_bg, tag="move_rows")
        for chunk in consecutive_chunks(rows):
            self.MT.show_ctrl_outline(
                start_cell=(0, chunk[0]),
                end_cell=(len(self.MT.col_positions) - 1, chunk[-1] + 1),
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
        ycheck = self.yview()
        need_redraw = False
        if event.y > self.winfo_height() and len(ycheck) > 1 and ycheck[1] < 1:
            try:
                self.MT.yview_scroll(1, "units")
                self.yview_scroll(1, "units")
            except Exception:
                pass
            self.fix_yview()
            self.MT.y_move_synced_scrolls("moveto", self.MT.yview()[0])
            need_redraw = True
        elif event.y < 0 and self.canvasy(self.winfo_height()) > 0 and ycheck and ycheck[0] > 0:
            try:
                self.yview_scroll(-1, "units")
                self.MT.yview_scroll(-1, "units")
            except Exception:
                pass
            self.fix_yview()
            self.MT.y_move_synced_scrolls("moveto", self.MT.yview()[0])
            need_redraw = True
        return need_redraw

    def fix_yview(self) -> None:
        ycheck = self.yview()
        if ycheck and ycheck[0] < 0:
            self.MT.set_yviews("moveto", 0)
        if len(ycheck) > 1 and ycheck[1] > 1:
            self.MT.set_yviews("moveto", 1)

    def event_over_dropdown(self, r: int, datarn: int, event: object, canvasy: float) -> bool:
        if (
            canvasy < self.MT.row_positions[r] + self.MT.index_txt_height
            and self.get_cell_kwargs(datarn, key="dropdown")
            and event.x > self.current_width - self.MT.index_txt_height - 4
        ):
            return True
        return False

    def event_over_checkbox(self, r: int, datarn: int, event: object, canvasy: float) -> bool:
        if (
            canvasy < self.MT.row_positions[r] + self.MT.index_txt_height
            and self.get_cell_kwargs(datarn, key="checkbox")
            and event.x < self.MT.index_txt_height + 4
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
        if self.height_resizing_enabled and self.rsz_h is not None and self.currently_resizing_height:
            self.currently_resizing_height = False
            new_row_pos = int(self.coords("rhl")[1])
            self.hide_resize_and_ctrl_lines(ctrl_lines=False)
            old_height = self.MT.row_positions[self.rsz_h] - self.MT.row_positions[self.rsz_h - 1]
            size = new_row_pos - self.MT.row_positions[self.rsz_h - 1]
            if size < self.MT.min_row_height:
                new_row_pos = ceil(self.MT.row_positions[self.rsz_h - 1] + self.MT.min_row_height)
            elif size > self.MT.max_row_height:
                new_row_pos = floor(self.MT.row_positions[self.rsz_h - 1] + self.MT.max_row_height)
            increment = new_row_pos - self.MT.row_positions[self.rsz_h]
            self.MT.row_positions[self.rsz_h + 1 :] = [
                e + increment for e in islice(self.MT.row_positions, self.rsz_h + 1, len(self.MT.row_positions))
            ]
            self.MT.row_positions[self.rsz_h] = new_row_pos
            new_height = self.MT.row_positions[self.rsz_h] - self.MT.row_positions[self.rsz_h - 1]
            self.MT.allow_auto_resize_rows = False
            self.MT.recreate_all_selection_boxes()
            self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
            if self.row_height_resize_func is not None and old_height != new_height:
                self.row_height_resize_func(
                    event_dict(
                        name="resize",
                        sheet=self.PAR.name,
                        resized_rows={self.rsz_h - 1: {"old_size": old_height, "new_size": new_height}},
                    )
                )
        elif self.width_resizing_enabled and self.rsz_w is not None and self.currently_resizing_width:
            self.currently_resizing_width = False
            self.hide_resize_and_ctrl_lines(ctrl_lines=False)
            self.set_width(self.new_row_width, set_TL=True)
            self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
        if (
            self.drag_and_drop_enabled
            and self.MT.anything_selected(exclude_cells=True, exclude_columns=True)
            and self.row_selection_enabled
            and self.rsz_h is None
            and self.rsz_w is None
            and self.dragged_row is not None
            and self.find_withtag("move_rows")
        ):
            self.hide_resize_and_ctrl_lines()
            r = self.MT.identify_row(y=event.y)
            totalrows = len(self.dragged_row.to_move)
            if (
                r is not None
                and totalrows != len(self.MT.row_positions) - 1
                and not (
                    r >= self.dragged_row.to_move[0]
                    and r <= self.dragged_row.to_move[-1]
                    and is_contiguous(self.dragged_row.to_move)
                )
            ):
                if r > self.dragged_row.to_move[-1]:
                    r += 1
                if r > len(self.MT.row_positions) - 1:
                    r = len(self.MT.row_positions) - 1
                event_data = event_dict(
                    name="move_rows",
                    sheet=self.PAR.name,
                    widget=self,
                    boxes=self.MT.get_boxes(),
                    selected=self.MT.selected,
                    value=r,
                )
                if try_binding(self.ri_extra_begin_drag_drop_func, event_data, "begin_move_rows"):
                    data_new_idxs, disp_new_idxs, event_data = self.MT.move_rows_adjust_options_dict(
                        *self.MT.get_args_for_move_rows(
                            move_to=r,
                            to_move=self.dragged_row.to_move,
                        ),
                        move_data=self.PAR.ops.row_drag_and_drop_perform,
                        move_heights=self.PAR.ops.row_drag_and_drop_perform,
                        event_data=event_data,
                    )
                    event_data["moved"]["rows"] = {
                        "data": data_new_idxs,
                        "displayed": disp_new_idxs,
                    }
                    if self.MT.undo_enabled:
                        self.MT.undo_stack.append(pickled_event_dict(event_data))
                    self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
                    try_binding(self.ri_extra_end_drag_drop_func, event_data, "end_move_rows")
                    self.MT.sheet_modified(event_data)
        elif self.b1_pressed_loc is not None and self.rsz_w is None and self.rsz_h is None:
            r = self.MT.identify_row(y=event.y)
            if (
                r is not None
                and r < len(self.MT.row_positions) - 1
                and r == self.b1_pressed_loc
                and self.b1_pressed_loc != self.closed_dropdown
            ):
                datarn = r if self.MT.all_rows_displayed else self.MT.displayed_rows[r]
                canvasy = self.canvasy(event.y)
                if self.event_over_dropdown(r, datarn, event, canvasy) or self.event_over_checkbox(
                    r, datarn, event, canvasy
                ):
                    self.open_cell(event)
                elif (iid := self.event_over_tree_arrow(r, canvasy, event.x)) is not None:
                    if self.MT.selection_boxes:
                        self.select_row(r, redraw=False)
                    self.PAR.item(iid, open_=iid not in self.tree_open_ids)
            else:
                self.mouseclick_outside_editor_or_dropdown_all_canvases(inside=True)
            self.b1_pressed_loc = None
            self.closed_dropdown = None
        self.dragged_row = None
        self.currently_resizing_width = False
        self.currently_resizing_height = False
        self.rsz_w = None
        self.rsz_h = None
        self.mouse_motion(event)
        try_binding(self.extra_b1_release_func, event)

    def event_over_tree_arrow(
        self,
        r: int,
        canvasy: float,
        eventx: int,
    ) -> bool:
        if self.PAR.ops.treeview and (
            canvasy < self.MT.row_positions[r] + self.MT.index_txt_height + 3
            and isinstance(self.MT._row_index, list)
            and (datarn := self.MT.datarn(r)) < len(self.MT._row_index)
            and eventx
            < self.get_treeview_indent((iid := self.MT._row_index[datarn].iid)) + self.MT.index_txt_height + 1
        ):
            return iid
        return None

    def toggle_select_row(
        self,
        row: int,
        add_selection: bool = True,
        redraw: bool = True,
        run_binding_func: bool = True,
        set_as_current: bool = True,
        ext: bool = False,
    ) -> int | None:
        if add_selection:
            if self.MT.row_selected(row):
                fill_iid = self.MT.deselect(r=row, redraw=redraw)
            else:
                fill_iid = self.add_selection(
                    r=row,
                    redraw=redraw,
                    run_binding_func=run_binding_func,
                    set_as_current=set_as_current,
                    ext=ext,
                )
        else:
            if self.MT.row_selected(row):
                fill_iid = self.MT.deselect(r=row, redraw=redraw)
            else:
                fill_iid = self.select_row(row, redraw=redraw, ext=ext)
        return fill_iid

    def select_row(
        self,
        r: int,
        redraw: bool = False,
        run_binding_func: bool = True,
        ext: bool = False,
    ) -> int:
        self.MT.deselect("all", redraw=False)
        box = (r, 0, r + 1, len(self.MT.col_positions) - 1, "rows")
        fill_iid = self.MT.create_selection_box(*box, ext=ext)
        if redraw:
            self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
        if run_binding_func:
            self.MT.run_selection_binding("rows")
        return fill_iid

    def add_selection(
        self,
        r: int,
        redraw: bool = False,
        run_binding_func: bool = True,
        set_as_current: bool = True,
        ext: bool = False,
    ) -> int:
        box = (r, 0, r + 1, len(self.MT.col_positions) - 1, "rows")
        fill_iid = self.MT.create_selection_box(*box, set_current=set_as_current, ext=ext)
        if redraw:
            self.MT.main_table_redraw_grid_and_text(redraw_header=False, redraw_row_index=True)
        if run_binding_func:
            self.MT.run_selection_binding("rows")
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
                iid = self.create_polygon(
                    coords,
                    fill=fill,
                    outline=outline,
                    state=state,
                    tags=tags,
                    smooth=True,
                )
            self.disp_boxes.add(iid)
        return iid

    def hide_box(self, item: int | None) -> None:
        if isinstance(item, int):
            self.disp_boxes.discard(item)
            self.hidd_boxes.add(item)
            self.itemconfig(item, state="hidden")

    def get_cell_dimensions(self, datarn: int) -> tuple[int, int]:
        txt = self.get_valid_cell_data_as_str(datarn, fix=False)
        if txt:
            self.MT.txt_measure_canvas.itemconfig(
                self.MT.txt_measure_canvas_text, text=txt, font=self.PAR.ops.index_font
            )
            b = self.MT.txt_measure_canvas.bbox(self.MT.txt_measure_canvas_text)
            w = b[2] - b[0] + 7
            h = b[3] - b[1] + 5
        else:
            w = self.PAR.ops.default_row_index_width
            h = self.MT.min_row_height
        if self.get_cell_kwargs(datarn, key="dropdown") or self.get_cell_kwargs(datarn, key="checkbox"):
            w += self.MT.index_txt_height + 2
        if self.PAR.ops.treeview:
            if datarn in self.cell_options and "align" in self.cell_options[datarn]:
                align = self.cell_options[datarn]["align"]
            else:
                align = self.align
            if align == "w":
                w += self.MT.index_txt_height
            w += self.get_treeview_indent(self.MT._row_index[datarn].iid) + 5
        return w, h

    def get_row_text_height(
        self,
        row: int,
        visible_only: bool = False,
        only_if_too_small: bool = False,
    ) -> int:
        h = self.MT.min_row_height
        datarn = row if self.MT.all_rows_displayed else self.MT.displayed_rows[row]
        # index
        _w, ih = self.get_cell_dimensions(datarn)
        # table
        if self.MT.data:
            if self.MT.all_columns_displayed:
                if visible_only:
                    iterable = range(*self.MT.visible_text_columns)
                else:
                    if not self.MT.data or datarn >= len(self.MT.data):
                        iterable = range(0, 0)
                    else:
                        iterable = range(0, len(self.MT.data[datarn]))
            else:
                if visible_only:
                    start_col, end_col = self.MT.visible_text_columns
                else:
                    start_col, end_col = 0, len(self.MT.displayed_columns)
                iterable = self.MT.displayed_columns[start_col:end_col]
            for datacn in iterable:
                if (txt := self.MT.get_valid_cell_data_as_str(datarn, datacn, get_displayed=True)) and (
                    th := self.MT.get_txt_h(txt) + 5
                ) > h:
                    h = th
        if ih > h:
            h = ih
        if only_if_too_small and h < self.MT.row_positions[row + 1] - self.MT.row_positions[row]:
            return self.MT.row_positions[row + 1] - self.MT.row_positions[row]
        if h < self.MT.min_row_height:
            h = int(self.MT.min_row_height)
        elif h > self.MT.max_row_height:
            h = int(self.MT.max_row_height)
        return h

    def set_row_height(
        self,
        row: int,
        height: None | int = None,
        only_if_too_small: bool = False,
        visible_only: bool = False,
        recreate: bool = True,
    ) -> int:
        if height is None:
            height = self.get_row_text_height(row=row, visible_only=visible_only)
        if height < self.MT.min_row_height:
            height = int(self.MT.min_row_height)
        elif height > self.MT.max_row_height:
            height = int(self.MT.max_row_height)
        if only_if_too_small and height <= self.MT.row_positions[row + 1] - self.MT.row_positions[row]:
            return self.MT.row_positions[row + 1] - self.MT.row_positions[row]
        new_row_pos = self.MT.row_positions[row] + height
        increment = new_row_pos - self.MT.row_positions[row + 1]
        self.MT.row_positions[row + 2 :] = [
            e + increment for e in islice(self.MT.row_positions, row + 2, len(self.MT.row_positions))
        ]
        self.MT.row_positions[row + 1] = new_row_pos
        if recreate:
            self.MT.recreate_all_selection_boxes()
        return height

    def get_index_text_width(
        self,
        only_rows: Iterator[int] | None = None,
    ) -> int:
        self.fix_index()
        w = self.PAR.ops.default_row_index_width
        if (not self.MT._row_index and isinstance(self.MT._row_index, list)) or (
            isinstance(self.MT._row_index, int) and self.MT._row_index >= len(self.MT.data)
        ):
            return w
        qconf = self.MT.txt_measure_canvas.itemconfig
        qbbox = self.MT.txt_measure_canvas.bbox
        qtxtm = self.MT.txt_measure_canvas_text
        if only_rows:
            iterable = only_rows
        elif self.MT.all_rows_displayed:
            if isinstance(self.MT._row_index, list):
                iterable = range(len(self.MT._row_index))
            else:
                iterable = range(len(self.MT.data))
        else:
            iterable = self.MT.displayed_rows
        if (
            isinstance(self.MT._row_index, list)
            and (tw := max(map(itemgetter(0), map(self.get_cell_dimensions, iterable)), default=w)) > w
        ):
            w = tw
        elif isinstance(self.MT._row_index, int):
            datacn = self.MT._row_index
            for datarn in iterable:
                if txt := self.MT.get_valid_cell_data_as_str(datarn, datacn, get_displayed=True):
                    qconf(qtxtm, text=txt)
                    b = qbbox(qtxtm)
                    if (tw := b[2] - b[0] + 10) > w:
                        w = tw
        if w > self.MT.max_index_width:
            w = int(self.MT.max_index_width)
        return w

    def set_width_of_index_to_text(
        self,
        text: None | str = None,
        only_rows: list = [],
    ) -> int:
        self.fix_index()
        w = self.PAR.ops.default_row_index_width
        if (text is None and isinstance(self.MT._row_index, list) and not self.MT._row_index) or (
            isinstance(self.MT._row_index, int) and self.MT._row_index >= len(self.MT.data)
        ):
            return w
        if text is not None and text:
            self.MT.txt_measure_canvas.itemconfig(self.MT.txt_measure_canvas_text, text=text)
            b = self.MT.txt_measure_canvas.bbox(self.MT.txt_measure_canvas_text)
            if (tw := b[2] - b[0] + 10) > w:
                w = tw
        elif text is None:
            w = self.get_index_text_width(only_rows=only_rows)
        if w > self.MT.max_index_width:
            w = int(self.MT.max_index_width)
        self.set_width(w, set_TL=True)
        self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
        return w

    def set_height_of_all_rows(
        self,
        height: int | None = None,
        only_if_too_small: bool = False,
        recreate: bool = True,
    ) -> None:
        if height is None:
            if self.MT.all_columns_displayed:
                iterable = range(self.MT.total_data_rows())
            else:
                iterable = range(len(self.MT.displayed_rows))
            self.MT.set_row_positions(
                itr=(self.get_row_text_height(rn, only_if_too_small=only_if_too_small) for rn in iterable)
            )
        elif height is not None:
            if self.MT.all_rows_displayed:
                self.MT.set_row_positions(itr=repeat(height, len(self.MT.data)))
            else:
                self.MT.set_row_positions(itr=repeat(height, len(self.MT.displayed_rows)))
        if recreate:
            self.MT.recreate_all_selection_boxes()

    def auto_set_index_width(self, end_row: int, only_rows: list) -> bool:
        if not isinstance(self.MT._row_index, int) and not self.MT._row_index:
            if self.default_index == "letters":
                new_w = self.MT.get_txt_w(f"{num2alpha(end_row)}") + 20
            elif self.default_index == "numbers":
                new_w = self.MT.get_txt_w(f"{end_row}") + 20
            elif self.default_index == "both":
                new_w = self.MT.get_txt_w(f"{end_row + 1} {num2alpha(end_row)}") + 20
        elif self.PAR.ops.auto_resize_row_index is True:
            new_w = self.get_index_text_width(only_rows=only_rows)
        else:
            new_w = None
        if new_w is not None and (sheet_w_x := floor(self.PAR.winfo_width() * 0.7)) < new_w:
            new_w = sheet_w_x
        if new_w and (self.current_width - new_w > 20 or new_w - self.current_width > 3):
            self.set_width(new_w, set_TL=True, recreate_selection_boxes=False)
            return True
        return False

    def redraw_highlight_get_text_fg(
        self,
        fr: float,
        sr: float,
        r: int,
        c_2: str,
        c_3: str,
        selections: dict,
        datarn: int,
    ) -> tuple[str, str, bool]:
        redrawn = False
        kwargs = self.get_cell_kwargs(datarn, key="highlight")
        if kwargs:
            if kwargs[0] is not None:
                c_1 = kwargs[0] if kwargs[0].startswith("#") else color_map[kwargs[0]]
            if "rows" in selections and r in selections["rows"]:
                txtfg = (
                    self.PAR.ops.index_selected_rows_fg
                    if kwargs[1] is None or self.PAR.ops.display_selected_fg_over_highlights
                    else kwargs[1]
                )
                if kwargs[0] is not None:
                    fill = (
                        f"#{int((int(c_1[1:3], 16) + int(c_3[1:3], 16)) / 2):02X}"
                        + f"{int((int(c_1[3:5], 16) + int(c_3[3:5], 16)) / 2):02X}"
                        + f"{int((int(c_1[5:], 16) + int(c_3[5:], 16)) / 2):02X}"
                    )
            elif "cells" in selections and r in selections["cells"]:
                txtfg = (
                    self.PAR.ops.index_selected_cells_fg
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
                txtfg = self.PAR.ops.index_fg if kwargs[1] is None else kwargs[1]
                if kwargs[0] is not None:
                    fill = kwargs[0]
            if kwargs[0] is not None:
                redrawn = self.redraw_highlight(
                    0,
                    fr + 1,
                    self.current_width - 1,
                    sr,
                    fill=fill,
                    outline=(
                        self.PAR.ops.index_fg
                        if self.get_cell_kwargs(datarn, key="dropdown") and self.PAR.ops.show_dropdown_borders
                        else ""
                    ),
                    tag="s",
                )
            tree_arrow_fg = txtfg
        elif not kwargs:
            if "rows" in selections and r in selections["rows"]:
                txtfg = self.PAR.ops.index_selected_rows_fg
                tree_arrow_fg = self.PAR.ops.selected_rows_tree_arrow_fg
            elif "cells" in selections and r in selections["cells"]:
                txtfg = self.PAR.ops.index_selected_cells_fg
                tree_arrow_fg = self.PAR.ops.selected_cells_tree_arrow_fg
            else:
                txtfg = self.PAR.ops.index_fg
                tree_arrow_fg = self.PAR.ops.tree_arrow_fg
        return txtfg, tree_arrow_fg, redrawn

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
        points: tuple[float],
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

    def redraw_tree_arrow(
        self,
        x1: float,
        y1: float,
        r: int,
        fill: str,
        tag: str | tuple[str],
        indent: float,
        open_: bool = False,
    ) -> None:
        mod = (self.MT.index_txt_height - 1) if self.MT.index_txt_height % 2 else self.MT.index_txt_height
        half_mod = mod / 2
        qtr_mod = mod / 4
        small_mod = int(half_mod / 4) - 1
        mid_y = int(self.MT.min_row_height / 2)
        # up arrow
        if open_:
            points = (
                x1 + 2 + indent,
                y1 + mid_y + qtr_mod,
                x1 + 2 + half_mod + indent,
                y1 + mid_y - qtr_mod,
                x1 + 2 + mod + indent,
                y1 + mid_y + qtr_mod,
            )
        # right pointing arrow
        else:
            points = (
                x1 + half_mod + indent,
                y1 + mid_y - half_mod + small_mod,
                x1 + mod + indent - small_mod,
                y1 + mid_y,
                x1 + half_mod + indent,
                y1 + mid_y + half_mod - small_mod,
            )
        if self.hidd_tree_arrow:
            t, sh = self.hidd_tree_arrow.popitem()
            self.coords(t, points)
            if sh:
                self.itemconfig(t, fill=fill)
            else:
                self.itemconfig(t, fill=fill, tag=tag, state="normal")
            self.lift(t)
        else:
            t = self.create_line(
                points,
                fill=fill,
                tag=tag,
                width=2,
                capstyle=tk.ROUND,
                joinstyle=tk.BEVEL,
            )
        self.disp_tree_arrow[t] = True

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
        open_: bool = False,
    ) -> None:
        if draw_outline and self.PAR.ops.show_dropdown_borders:
            self.redraw_highlight(x1 + 1, y1 + 1, x2, y2, fill="", outline=self.PAR.ops.index_fg, tag=tag)
        if draw_arrow:
            mod = (self.MT.index_txt_height - 1) if self.MT.index_txt_height % 2 else self.MT.index_txt_height
            half_mod = mod / 2
            qtr_mod = mod / 4
            mid_y = (self.MT.index_first_ln_ins - 1) if self.MT.index_first_ln_ins % 2 else self.MT.index_first_ln_ins
            if open_:
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

    def configure_scrollregion(self, last_row_line_pos: float) -> None:
        self.configure(
            scrollregion=(
                0,
                0,
                self.current_width,
                last_row_line_pos + self.PAR.ops.empty_vertical + 2,
            )
        )

    def redraw_grid_and_text(
        self,
        last_row_line_pos: float,
        scrollpos_top: int,
        y_stop: int,
        grid_start_row: int,
        grid_end_row: int,
        text_start_row: int,
        text_end_row: int,
        scrollpos_bot: int,
        row_pos_exists: bool,
    ) -> None:
        try:
            self.configure_scrollregion(last_row_line_pos=last_row_line_pos)
        except Exception:
            return
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
        self.hidd_tree_arrow.update(self.disp_tree_arrow)
        self.disp_tree_arrow = {}
        self.visible_row_dividers = {}
        draw_y = self.MT.row_positions[grid_start_row]
        xend = self.current_width - 6
        self.row_width_resize_bbox = (
            self.current_width - 2,
            scrollpos_top,
            self.current_width,
            scrollpos_bot,
        )
        if (self.PAR.ops.show_horizontal_grid or self.height_resizing_enabled) and row_pos_exists:
            points = [
                self.current_width - 1,
                y_stop - 1,
                self.current_width - 1,
                scrollpos_top - 1,
                -1,
                scrollpos_top - 1,
            ]
            for r in range(grid_start_row, grid_end_row):
                draw_y = self.MT.row_positions[r]
                if r and self.height_resizing_enabled:
                    self.visible_row_dividers[r] = (1, draw_y - 2, xend, draw_y + 2)
                points.extend(
                    (
                        -1,
                        draw_y,
                        self.current_width,
                        draw_y,
                        -1,
                        draw_y,
                        -1,
                        self.MT.row_positions[r + 1] if len(self.MT.row_positions) - 1 > r else draw_y,
                    )
                )
            self.redraw_gridline(points=points, fill=self.PAR.ops.index_grid_fg, width=1, tag="h")
        c_2 = (
            self.PAR.ops.index_selected_cells_bg
            if self.PAR.ops.index_selected_cells_bg.startswith("#")
            else color_map[self.PAR.ops.index_selected_cells_bg]
        )
        c_3 = (
            self.PAR.ops.index_selected_rows_bg
            if self.PAR.ops.index_selected_rows_bg.startswith("#")
            else color_map[self.PAR.ops.index_selected_rows_bg]
        )
        font = self.PAR.ops.index_font
        selections = self.get_redraw_selections(text_start_row, grid_end_row)
        dd_coords = self.dropdown.get_coords()
        treeview = self.PAR.ops.treeview

        for r in range(text_start_row, text_end_row):
            rtopgridln = self.MT.row_positions[r]
            rbotgridln = self.MT.row_positions[r + 1]
            if rbotgridln - rtopgridln < self.MT.index_txt_height:
                continue
            datarn = r if self.MT.all_rows_displayed else self.MT.displayed_rows[r]
            fill, tree_arrow_fg, dd_drawn = self.redraw_highlight_get_text_fg(
                rtopgridln,
                rbotgridln,
                r,
                c_2,
                c_3,
                selections,
                datarn,
            )

            if datarn in self.cell_options and "align" in self.cell_options[datarn]:
                align = self.cell_options[datarn]["align"]
            else:
                align = self.align
            dropdown_kwargs = self.get_cell_kwargs(datarn, key="dropdown")
            if align == "w":
                draw_x = 3
                if dropdown_kwargs:
                    mw = self.current_width - self.MT.index_txt_height - 2
                    self.redraw_dropdown(
                        0,
                        rtopgridln,
                        self.current_width - 1,
                        rbotgridln - 1,
                        fill=fill,
                        outline=fill,
                        tag="dd",
                        draw_outline=not dd_drawn,
                        draw_arrow=mw >= 5,
                        open_=dd_coords == r,
                    )
                else:
                    mw = self.current_width - 2

            elif align == "e":
                if dropdown_kwargs:
                    mw = self.current_width - self.MT.index_txt_height - 2
                    draw_x = self.current_width - 5 - self.MT.index_txt_height
                    self.redraw_dropdown(
                        0,
                        rtopgridln,
                        self.current_width - 1,
                        rbotgridln - 1,
                        fill=fill,
                        outline=fill,
                        tag="dd",
                        draw_outline=not dd_drawn,
                        draw_arrow=mw >= 5,
                        open_=dd_coords == r,
                    )
                else:
                    mw = self.current_width - 2
                    draw_x = self.current_width - 3

            elif align == "center":
                if dropdown_kwargs:
                    mw = self.current_width - self.MT.index_txt_height - 2
                    draw_x = ceil((self.current_width - self.MT.index_txt_height) / 2)
                    self.redraw_dropdown(
                        0,
                        rtopgridln,
                        self.current_width - 1,
                        rbotgridln - 1,
                        fill=fill,
                        outline=fill,
                        tag="dd",
                        draw_outline=not dd_drawn,
                        draw_arrow=mw >= 5,
                        open_=dd_coords == r,
                    )
                else:
                    mw = self.current_width - 1
                    draw_x = floor(self.current_width / 2)
            checkbox_kwargs = self.get_cell_kwargs(datarn, key="checkbox")
            if checkbox_kwargs and not dropdown_kwargs and mw > self.MT.index_txt_height + 1:
                box_w = self.MT.index_txt_height + 1
                if align == "w":
                    draw_x += box_w + 3
                    mw -= box_w + 3
                elif align == "center":
                    draw_x += ceil(box_w / 2) + 1
                    mw -= box_w + 2
                else:
                    mw -= box_w + 1
                try:
                    draw_check = (
                        self.MT._row_index[datarn]
                        if isinstance(self.MT._row_index, (list, tuple))
                        else self.MT.data[datarn][self.MT._row_index]
                    )
                except Exception:
                    draw_check = False
                self.redraw_checkbox(
                    2,
                    rtopgridln + 2,
                    self.MT.index_txt_height + 3,
                    rtopgridln + self.MT.index_txt_height + 3,
                    fill=fill if checkbox_kwargs["state"] == "normal" else self.PAR.ops.index_grid_fg,
                    outline="",
                    tag="cb",
                    draw_check=draw_check,
                )
            if treeview and isinstance(self.MT._row_index, list) and len(self.MT._row_index) > datarn:
                iid = self.MT._row_index[datarn].iid
                mw -= self.MT.index_txt_height
                if align == "w":
                    draw_x += self.MT.index_txt_height + 1
                indent = self.get_treeview_indent(iid)
                draw_x += indent + 5
                if self.tree[iid].children:
                    self.redraw_tree_arrow(
                        0,
                        rtopgridln,
                        r=r,
                        fill=tree_arrow_fg,
                        tag="ta",
                        indent=indent,
                        open_=self.MT._row_index[datarn].iid in self.tree_open_ids,
                    )
            lns = self.get_valid_cell_data_as_str(datarn, fix=False)
            if not lns:
                continue
            lns = lns.split("\n")
            draw_y = rtopgridln + self.MT.index_first_ln_ins
            if mw > 5:
                draw_y = rtopgridln + self.MT.index_first_ln_ins
                start_ln = int((scrollpos_top - rtopgridln) / self.MT.index_xtra_lines_increment)
                if start_ln < 0:
                    start_ln = 0
                draw_y += start_ln * self.MT.index_xtra_lines_increment
                if draw_y + self.MT.index_half_txt_height - 1 <= rbotgridln and len(lns) > start_ln:
                    for txt in islice(lns, start_ln, None):
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
                            if align == "w" and dropdown_kwargs:
                                txt = txt[: int(len(txt) * (mw / wd))]
                                self.itemconfig(iid, text=txt)
                                wd = self.bbox(iid)
                                while wd[2] - wd[0] > mw:
                                    txt = txt[:-1]
                                    self.itemconfig(iid, text=txt)
                                    wd = self.bbox(iid)
                            elif align == "e" and (dropdown_kwargs or checkbox_kwargs):
                                txt = txt[len(txt) - int(len(txt) * (mw / wd)) :]
                                self.itemconfig(iid, text=txt)
                                wd = self.bbox(iid)
                                while wd[2] - wd[0] > mw:
                                    txt = txt[1:]
                                    self.itemconfig(iid, text=txt)
                                    wd = self.bbox(iid)
                            elif align == "center" and (dropdown_kwargs or checkbox_kwargs):
                                tmod = ceil((len(txt) - int(len(txt) * (mw / wd))) / 2)
                                txt = txt[tmod - 1 : -tmod]
                                self.itemconfig(iid, text=txt)
                                wd = self.bbox(iid)
                                self.c_align_cyc = cycle(self.centre_alignment_text_mod_indexes)
                                while wd[2] - wd[0] > mw:
                                    txt = txt[next(self.c_align_cyc)]
                                    self.itemconfig(iid, text=txt)
                                    wd = self.bbox(iid)
                                self.coords(iid, draw_x, draw_y)
                        draw_y += self.MT.index_xtra_lines_increment
                        if draw_y + self.MT.index_half_txt_height - 1 > rbotgridln:
                            break
        for dct in (
            self.hidd_text,
            self.hidd_high,
            self.hidd_grid,
            self.hidd_dropdown,
            self.hidd_checkbox,
            self.hidd_tree_arrow,
        ):
            for iid, showing in dct.items():
                if showing:
                    self.itemconfig(iid, state="hidden")
                    dct[iid] = False
        return True

    def get_redraw_selections(self, startr: int, endr: int) -> dict[str, set[int]]:
        d = defaultdict(set)
        for item, box in self.MT.get_selection_items():
            r1, c1, r2, c2 = box.coords
            for r in range(startr, endr):
                if r1 <= r and r2 > r:
                    d[box.type_ if box.type_ != "columns" else "cells"].add(r)
        return d

    def open_cell(self, event: object = None, ignore_existing_editor: bool = False) -> None:
        if not self.MT.anything_selected() or (not ignore_existing_editor and self.text_editor.open):
            return
        if not self.MT.selected:
            return
        r = self.MT.selected.row
        datarn = r if self.MT.all_rows_displayed else self.MT.displayed_rows[r]
        if self.get_cell_kwargs(datarn, key="readonly"):
            return
        elif self.get_cell_kwargs(datarn, key="dropdown") or self.get_cell_kwargs(datarn, key="checkbox"):
            if self.MT.event_opens_dropdown_or_checkbox(event):
                if self.get_cell_kwargs(datarn, key="dropdown"):
                    self.open_dropdown_window(r, event=event)
                elif self.get_cell_kwargs(datarn, key="checkbox"):
                    self.click_checkbox(r, datarn)
        elif self.edit_cell_enabled:
            self.open_text_editor(event=event, r=r, dropdown=False)

    # displayed indexes
    def get_cell_align(self, r: int) -> str:
        datarn = r if self.MT.all_rows_displayed else self.MT.displayed_rows[r]
        if datarn in self.cell_options and "align" in self.cell_options[datarn]:
            align = self.cell_options[datarn]["align"]
        else:
            align = self.align
        return align

    # r is displayed row
    def open_text_editor(
        self,
        event: object = None,
        r: int = 0,
        text: None | str = None,
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
            text = self.get_cell_data(self.MT.datarn(r), none_to_empty_str=True, redirect_int=True)
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
                        name="begin_edit_index",
                        sheet=self.PAR.name,
                        key=extra_func_key,
                        value=text,
                        loc=r,
                        row=r,
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
            self.set_row_height_run_binding(r)
        if self.text_editor.open and r == self.text_editor.row:
            self.text_editor.set_text(self.text_editor.get() + "" if not isinstance(text, str) else text)
            return
        if self.text_editor.open:
            self.hide_text_editor()
        if not self.MT.see(r=r, c=0, keep_yscroll=True, check_cell_visibility=True):
            self.MT.refresh()
        x = 0
        y = self.MT.row_positions[r]
        w = self.current_width + 1
        h = self.MT.row_positions[r + 1] - y + 1
        if text is None:
            text = self.get_cell_data(self.MT.datarn(r), none_to_empty_str=True, redirect_int=True)
        bg, fg = self.PAR.ops.index_bg, self.PAR.ops.index_fg
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
            "border_color": self.PAR.ops.index_selected_rows_bg,
            "text": text,
            "state": state,
            "width": w,
            "height": h,
            "show_border": True,
            "bg": bg,
            "fg": fg,
            "align": self.get_cell_align(r),
            "r": r,
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

    def text_editor_newline_binding(self, event: object = None, check_lines: bool = True) -> None:
        if not self.height_resizing_enabled:
            return
        curr_height = self.text_editor.window.winfo_height()
        if curr_height < self.MT.min_row_height:
            return
        if (
            not check_lines
            or self.MT.get_lines_cell_height(
                self.text_editor.window.get_num_lines() + 1,
                font=self.text_editor.tktext.cget("font"),
            )
            > curr_height
        ):
            r = self.text_editor.row
            new_height = curr_height + self.MT.index_xtra_lines_increment
            space_bot = self.MT.get_space_bot(r)
            if new_height > space_bot:
                new_height = space_bot
            if new_height != curr_height:
                self.set_row_height(r, new_height)
                self.MT.refresh()
                self.text_editor.window.config(height=new_height)
                self.coords(self.text_editor.canvas_id, 0, self.MT.row_positions[r] + 1)
                if self.dropdown.open and self.dropdown.get_coords() == r:
                    text_editor_h = self.text_editor.window.winfo_height()
                    win_h, anchor = self.get_dropdown_height_anchor(r, text_editor_h)
                    if anchor == "nw":
                        self.coords(
                            self.dropdown.canvas_id,
                            0,
                            self.MT.row_positions[r] + text_editor_h - 1,
                        )
                        self.itemconfig(self.dropdown.canvas_id, anchor=anchor, height=win_h)
                    elif anchor == "sw":
                        self.coords(
                            self.dropdown.canvas_id,
                            0,
                            self.MT.row_positions[r],
                        )
                        self.itemconfig(self.dropdown.canvas_id, anchor=anchor, height=win_h)

    def refresh_open_window_positions(self, zoom: Literal["in", "out"]) -> None:
        if self.text_editor.open:
            r = self.text_editor.row
            self.text_editor.window.config(height=self.MT.row_positions[r + 1] - self.MT.row_positions[r])
            self.text_editor.tktext.config(font=self.PAR.ops.index_font)
            self.coords(
                self.text_editor.canvas_id,
                0,
                self.MT.row_positions[r],
            )
        if self.dropdown.open:
            if zoom == "in":
                self.dropdown.window.zoom_in()
            elif zoom == "out":
                self.dropdown.window.zoom_out()
            r = self.dropdown.get_coords()
            if self.text_editor.open:
                text_editor_h = self.text_editor.window.winfo_height()
                win_h, anchor = self.get_dropdown_height_anchor(r, text_editor_h)
            else:
                text_editor_h = self.MT.row_positions[r + 1] - self.MT.row_positions[r] + 1
                anchor = self.itemcget(self.dropdown.canvas_id, "anchor")
                # win_h = 0
            self.dropdown.window.config(width=self.current_width + 1)
            if anchor == "nw":
                self.coords(
                    self.dropdown.canvas_id,
                    0,
                    self.MT.row_positions[r] + text_editor_h - 1,
                )
                # self.itemconfig(self.dropdown.canvas_id, anchor=anchor, height=win_h)
            elif anchor == "sw":
                self.coords(
                    self.dropdown.canvas_id,
                    0,
                    self.MT.row_positions[r],
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

    # r is displayed row
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
        text_editor_value = self.text_editor.get()
        r = self.text_editor.row
        datarn = r if self.MT.all_rows_displayed else self.MT.displayed_rows[r]
        event_data = event_dict(
            name="end_edit_index",
            sheet=self.PAR.name,
            widget=self,
            cells_index={datarn: self.get_cell_data(datarn)},
            key=event.keysym,
            value=text_editor_value,
            loc=r,
            row=r,
            boxes=self.MT.get_boxes(),
            selected=self.MT.selected,
        )
        edited = False
        set_data = partial(
            self.set_cell_data_undo,
            r=r,
            datarn=datarn,
            check_input_valid=False,
        )
        if self.MT.edit_validation_func:
            text_editor_value = self.MT.edit_validation_func(event_data)
            if text_editor_value is not None and self.input_valid_for_cell(datarn, text_editor_value):
                edited = set_data(value=text_editor_value)
        elif self.input_valid_for_cell(datarn, text_editor_value):
            edited = set_data(value=text_editor_value)
        if edited:
            try_binding(self.extra_end_edit_cell_func, event_data)
        self.MT.recreate_all_selection_boxes()
        self.hide_text_editor_and_dropdown()
        if event.keysym != "FocusOut":
            self.focus_set()
        return "break"

    def get_dropdown_height_anchor(self, r: int, text_editor_h: None | int = None) -> tuple[int, str]:
        win_h = 5
        datarn = self.MT.datarn(r)
        for i, v in enumerate(self.get_cell_kwargs(datarn, key="dropdown")["values"]):
            v_numlines = len(v.split("\n") if isinstance(v, str) else f"{v}".split("\n"))
            if v_numlines > 1:
                win_h += (
                    self.MT.index_first_ln_ins + (v_numlines * self.MT.index_xtra_lines_increment) + 5
                )  # end of cell
            else:
                win_h += self.MT.min_row_height
            if i == 5:
                break
        if win_h > 500:
            win_h = 500
        space_bot = self.MT.get_space_bot(r, text_editor_h)
        space_top = int(self.MT.row_positions[r])
        anchor = "nw"
        win_h2 = int(win_h)
        if win_h > space_bot:
            if space_bot >= space_top:
                anchor = "nw"
                win_h = space_bot - 1
            elif space_top > space_bot:
                anchor = "sw"
                win_h = space_top - 1
        if win_h < self.MT.index_txt_height + 5:
            win_h = self.MT.index_txt_height + 5
        elif win_h > win_h2:
            win_h = win_h2
        return win_h, anchor

    def dropdown_text_editor_modified(
        self,
        dd_window: object,
        event: dict,
        modified_func: Callable | None,
    ) -> None:
        if modified_func:
            modified_func(event)
        dd_window.search_and_see(event)

    # r is displayed row
    def open_dropdown_window(self, r: int, event: object = None) -> None:
        self.hide_text_editor("Escape")
        kwargs = self.get_cell_kwargs(self.MT.datarn(r), key="dropdown")
        if kwargs["state"] == "normal":
            if not self.open_text_editor(event=event, r=r, dropdown=True):
                return
        win_h, anchor = self.get_dropdown_height_anchor(r)
        win_w = self.current_width + 1
        if anchor == "nw":
            if kwargs["state"] == "normal":
                self.text_editor.window.update_idletasks()
                ypos = self.MT.row_positions[r] + self.text_editor.window.winfo_height() - 1
            else:
                ypos = self.MT.row_positions[r + 1]
        else:
            ypos = self.MT.row_positions[r]
        reset_kwargs = {
            "r": r,
            "c": 0,
            "width": win_w,
            "height": win_h,
            "font": self.PAR.ops.index_font,
            "ops": self.PAR.ops,
            "outline_color": self.PAR.ops.popup_menu_fg,
            "align": self.get_cell_align(r),
            "values": kwargs["values"],
        }
        if self.dropdown.window:
            self.dropdown.window.reset(**reset_kwargs)
            self.itemconfig(self.dropdown.canvas_id, state="normal", anchor=anchor)
            self.coords(self.dropdown.canvas_id, 0, ypos)
            self.dropdown.window.tkraise()
        else:
            self.dropdown.window = self.PAR.dropdown_class(
                self.winfo_toplevel(),
                **reset_kwargs,
                single_index="r",
                close_dropdown_window=self.close_dropdown_window,
                search_function=kwargs["search_function"],
                arrowkey_RIGHT=self.MT.arrowkey_RIGHT,
                arrowkey_LEFT=self.MT.arrowkey_LEFT,
            )
            self.dropdown.canvas_id = self.create_window(
                (0, ypos),
                window=self.dropdown.window,
                anchor=anchor,
            )
        if kwargs["state"] == "normal":
            self.text_editor.tktext.bind(
                "<<TextModified>>",
                lambda _x: self.dropdown_text_editor_modified(
                    self.dropdown.window,
                    event_dict(
                        name="index_dropdown_modified",
                        sheet=self.PAR.name,
                        value=self.text_editor.get(),
                        loc=r,
                        row=r,
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
            self.dropdown.window.bind("<FocusOut>", lambda _x: self.close_dropdown_window(r))
            self.update_idletasks()
            self.dropdown.window.focus_set()
            redraw = True
        self.dropdown.open = True
        if redraw:
            self.MT.main_table_redraw_grid_and_text(redraw_header=False, redraw_row_index=True, redraw_table=False)

    # r is displayed row
    def close_dropdown_window(
        self,
        r: int | None = None,
        selection: object = None,
        redraw: bool = True,
    ) -> None:
        if r is not None and selection is not None:
            datarn = r if self.MT.all_rows_displayed else self.MT.displayed_rows[r]
            kwargs = self.get_cell_kwargs(datarn, key="dropdown")
            pre_edit_value = self.get_cell_data(datarn)
            edited = False
            event_data = event_dict(
                name="end_edit_index",
                sheet=self.PAR.name,
                widget=self,
                cells_header={datarn: pre_edit_value},
                key="??",
                value=selection,
                loc=r,
                row=r,
                boxes=self.MT.get_boxes(),
                selected=self.MT.selected,
            )
            if kwargs["select_function"] is not None:
                kwargs["select_function"](event_data)
            if self.MT.edit_validation_func:
                selection = self.MT.edit_validation_func(event_data)
                if selection is not None:
                    edited = self.set_cell_data_undo(r, datarn=datarn, value=selection, redraw=not redraw)
            else:
                edited = self.set_cell_data_undo(r, datarn=datarn, value=selection, redraw=not redraw)
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
                    redraw_header=False,
                    redraw_row_index=True,
                    redraw_table=False,
                )
        return closed_dd_coords

    def mouseclick_outside_editor_or_dropdown_all_canvases(self, inside: bool = False) -> int | None:
        self.CH.mouseclick_outside_editor_or_dropdown()
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
        r: int = 0,
        datarn: None | int = None,
        value: str = "",
        cell_resize: bool = True,
        undo: bool = True,
        redraw: bool = True,
        check_input_valid: bool = True,
    ) -> bool:
        if datarn is None:
            datarn = r if self.MT.all_rows_displayed else self.MT.displayed_rows[r]
        event_data = event_dict(
            name="edit_index",
            sheet=self.PAR.name,
            widget=self,
            cells_index={datarn: self.get_cell_data(datarn)},
            boxes=self.MT.get_boxes(),
            selected=self.MT.selected,
        )
        edited = False
        if isinstance(self.MT._row_index, int):
            edited = self.MT.set_cell_data_undo(r=r, c=self.MT._row_index, datarn=datarn, value=value, undo=True)
        else:
            self.fix_index(datarn)
            if not check_input_valid or self.input_valid_for_cell(datarn, value):
                if self.MT.undo_enabled and undo:
                    self.MT.undo_stack.append(pickled_event_dict(event_data))
                self.set_cell_data(datarn=datarn, value=value)
                edited = True
        if edited and cell_resize and self.PAR.ops.cell_auto_resize_enabled:
            self.set_row_height_run_binding(r, only_if_too_small=False)
        if redraw:
            self.MT.refresh()
        if edited:
            self.MT.sheet_modified(event_data)
        return edited

    def set_cell_data(self, datarn: int | None = None, value: object = "") -> None:
        if isinstance(self.MT._row_index, int):
            self.MT.set_cell_data(datarn=datarn, datacn=self.MT._row_index, value=value)
        else:
            self.fix_index(datarn)
            if self.get_cell_kwargs(datarn, key="checkbox"):
                self.MT._row_index[datarn] = try_to_bool(value)
            else:
                self.MT._row_index[datarn] = value

    def input_valid_for_cell(self, datarn: int, value: object, check_readonly: bool = True) -> bool:
        if check_readonly and self.get_cell_kwargs(datarn, key="readonly"):
            return False
        if self.get_cell_kwargs(datarn, key="checkbox"):
            return is_bool_like(value)
        if self.cell_equal_to(datarn, value):
            return False
        kwargs = self.get_cell_kwargs(datarn, key="dropdown")
        if kwargs and kwargs["validate_input"] and value not in kwargs["values"]:
            return False
        return True

    def cell_equal_to(self, datarn: int, value: object) -> bool:
        self.fix_index(datarn)
        if isinstance(self.MT._row_index, list):
            return self.MT._row_index[datarn] == value
        elif isinstance(self.MT._row_index, int):
            return self.MT.cell_equal_to(datarn, self.MT._row_index, value)

    def get_cell_data(
        self,
        datarn: int,
        get_displayed: bool = False,
        none_to_empty_str: bool = False,
        redirect_int: bool = False,
    ) -> object:
        if get_displayed:
            return self.get_valid_cell_data_as_str(datarn, fix=False)
        if redirect_int and isinstance(self.MT._row_index, int):  # internal use
            return self.MT.get_cell_data(datarn, self.MT._row_index, none_to_empty_str=True)
        if (
            isinstance(self.MT._row_index, int)
            or not self.MT._row_index
            or datarn >= len(self.MT._row_index)
            or (self.MT._row_index[datarn] is None and none_to_empty_str)
        ):
            return ""
        return self.MT._row_index[datarn]

    def get_valid_cell_data_as_str(self, datarn: int, fix: bool = True) -> str:
        kwargs = self.get_cell_kwargs(datarn, key="dropdown")
        if kwargs:
            if kwargs["text"] is not None:
                return f"{kwargs['text']}"
        else:
            kwargs = self.get_cell_kwargs(datarn, key="checkbox")
            if kwargs:
                return f"{kwargs['text']}"
        if isinstance(self.MT._row_index, int):
            return self.MT.get_valid_cell_data_as_str(datarn, self.MT._row_index, get_displayed=True)
        if fix:
            self.fix_index(datarn)
        try:
            value = "" if self.MT._row_index[datarn] is None else f"{self.MT._row_index[datarn]}"
        except Exception:
            value = ""
        if not value and self.PAR.ops.show_default_index_for_empty:
            value = get_n2a(datarn, self.default_index)
        return value

    def get_value_for_empty_cell(self, datarn: int, r_ops: bool = True) -> object:
        if self.get_cell_kwargs(datarn, key="checkbox", cell=r_ops):
            return False
        kwargs = self.get_cell_kwargs(datarn, key="dropdown", cell=r_ops)
        if kwargs and kwargs["validate_input"] and kwargs["values"]:
            return kwargs["values"][0]
        return ""

    def get_empty_index_seq(self, end: int, start: int = 0, r_ops: bool = True) -> list[object]:
        return [self.get_value_for_empty_cell(datarn, r_ops=r_ops) for datarn in range(start, end)]

    def fix_index(self, datarn: int | None = None) -> None:
        if isinstance(self.MT._row_index, int):
            return
        if isinstance(self.MT._row_index, float):
            self.MT._row_index = int(self.MT._row_index)
            return
        if not isinstance(self.MT._row_index, list):
            try:
                self.MT._row_index = list(self.MT._row_index)
            except Exception:
                self.MT._row_index = []
        if isinstance(datarn, int) and datarn >= len(self.MT._row_index):
            self.MT._row_index.extend(self.get_empty_index_seq(end=datarn + 1, start=len(self.MT._row_index)))

    def set_row_height_run_binding(self, r: int, only_if_too_small: bool = True) -> None:
        old_height = self.MT.row_positions[r + 1] - self.MT.row_positions[r]
        new_height = self.set_row_height(r, only_if_too_small=only_if_too_small)
        if self.row_height_resize_func is not None and old_height != new_height:
            self.row_height_resize_func(
                event_dict(
                    name="resize",
                    sheet=self.PAR.name,
                    resized_rows={r: {"old_size": old_height, "new_size": new_height}},
                )
            )

    # internal event use
    def click_checkbox(self, r: int, datarn: int | None = None, undo: bool = True, redraw: bool = True) -> None:
        if datarn is None:
            datarn = r if self.MT.all_rows_displayed else self.MT.displayed_rows[r]
        kwargs = self.get_cell_kwargs(datarn, key="checkbox")
        if kwargs["state"] == "normal":
            pre_edit_value = self.get_cell_data(datarn)
            if isinstance(self.MT._row_index, list):
                value = not self.MT._row_index[datarn] if isinstance(self.MT._row_index[datarn], bool) else False
            elif isinstance(self.MT._row_index, int):
                value = (
                    not self.MT.data[datarn][self.MT._row_index]
                    if type(self.MT.data[datarn][self.MT._row_index], bool)
                    else False
                )
            else:
                value = False
            self.set_cell_data_undo(r, datarn=datarn, value=value, cell_resize=False)
            event_data = event_dict(
                name="end_edit_index",
                sheet=self.PAR.name,
                widget=self,
                cells_index={datarn: pre_edit_value},
                key="??",
                value=value,
                loc=r,
                row=r,
                boxes=self.MT.get_boxes(),
                selected=self.MT.selected,
            )
            if kwargs["check_function"] is not None:
                kwargs["check_function"](event_data)
            try_binding(self.extra_end_edit_cell_func, event_data)
        if redraw:
            self.MT.refresh()

    def get_cell_kwargs(self, datarn: int, key: Hashable = "dropdown", cell: bool = True) -> dict:
        if cell and datarn in self.cell_options and key in self.cell_options[datarn]:
            return self.cell_options[datarn][key]
        return {}

    # Treeview Mode

    def get_node_level(self, node: Node, level: int = 0) -> Generator[int]:
        yield level
        if node.parent:
            yield from self.get_node_level(node.parent, level + 1)

    def ancestors_all_open(self, iid: str, stop_at: str | Node = "") -> bool:
        if stop_at:
            stop_at = stop_at.iid
            for iid in self.get_iid_ancestors(iid):
                if iid == stop_at:
                    return True
                if iid not in self.tree_open_ids:
                    return False
            return True
        else:
            return all(iid in self.tree_open_ids for iid in self.get_iid_ancestors(iid))

    def get_iid_ancestors(self, iid: str) -> Generator[str]:
        if self.tree[iid].parent:
            yield self.tree[iid].parent.iid
            yield from self.get_iid_ancestors(self.tree[iid].parent.iid)

    def get_iid_descendants(self, iid: str, check_open: bool = False) -> Generator[str]:
        for cnode in self.tree[iid].children:
            yield cnode.iid
            if (check_open and cnode.children and cnode.iid in self.tree_open_ids) or (
                not check_open and cnode.children
            ):
                yield from self.get_iid_descendants(cnode.iid, check_open)

    def items_parent(self, iid: str) -> str:
        if self.tree[iid].parent:
            return self.tree[iid].parent.iid
        return ""

    def gen_top_nodes(self) -> Generator[Node]:
        yield from (node for node in self.MT._row_index if node.parent == "")

    def get_treeview_indent(self, iid: str) -> int:
        if isinstance(self.PAR.ops.treeview_indent, str):
            indent = self.MT.index_txt_width * int(self.PAR.ops.treeview_indent)
        else:
            indent = self.PAR.ops.treeview_indent
        return indent * max(self.get_node_level(self.tree[iid]))

    def remove_node_from_parents_children(self, node: Node) -> None:
        if node.parent:
            node.parent.children.remove(node)
            if not node.parent.children:
                self.tree_open_ids.discard(node.parent)

    def pid_causes_recursive_loop(self, iid: str, pid: str) -> bool:
        return any(
            i == pid
            for i in chain(
                self.get_iid_descendants(iid),
                islice(self.get_iid_ancestors(iid), 1, None),
            )
        )
