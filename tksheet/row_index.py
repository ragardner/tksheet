from __future__ import annotations

import tkinter as tk
from collections import defaultdict
from collections.abc import Callable, Generator, Hashable, Iterator, Sequence
from functools import partial
from itertools import islice, repeat
from math import ceil
from re import findall
from typing import Any, Literal

from .colors import color_map
from .constants import (
    _test_str,
    text_editor_close_bindings,
    text_editor_newline_bindings,
    text_editor_to_unbind,
)
from .formatters import is_bool_like, try_to_bool
from .functions import (
    add_to_displayed,
    any_editor_or_dropdown_open,
    consecutive_ranges,
    event_dict,
    event_has_char_key,
    event_opens_dropdown_or_checkbox,
    get_bg_fg,
    get_menu_kwargs,
    get_n2a,
    get_new_indexes,
    int_x_tuple,
    is_contiguous,
    mod_event_val,
    new_tk_event,
    num2alpha,
    push_displayed,
    recursive_bind,
    remove_duplicates_outside_section,
    rounded_box_coords,
    safe_copy,
    stored_event_dict,
    try_b_index,
    try_binding,
    widget_descendants,
    wrap_text,
)
from .menus import build_empty_rc_menu, build_index_rc_menu
from .other_classes import DraggedRowColumn, DropdownStorage, EventDataDict, Node, TextEditorStorage
from .sorting import sort_columns_by_row, sort_row
from .text_editor import TextEditor
from .tooltip import Tooltip


class RowIndex(tk.Canvas):
    def __init__(self, parent: tk.Misc, **kwargs) -> None:
        super().__init__(
            parent,
            background=parent.ops.index_bg,
            highlightthickness=0,
        )
        self.PAR = parent
        self.ops = self.PAR.ops
        self.MT = None  # is set from within MainTable() __init__
        self.CH = None  # is set from within MainTable() __init__
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
        self.new_iid_ctr = -1
        self.current_width = None
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
        self.extra_rc_func = None
        self.selection_binding_func = None
        self.shift_selection_binding_func = None
        self.ctrl_selection_binding_func = None
        self.drag_selection_binding_func = None
        self.ri_extra_begin_drag_drop_func = None
        self.ri_extra_end_drag_drop_func = None
        self.ri_extra_begin_sort_cols_func = None
        self.ri_extra_end_sort_cols_func = None
        self.extra_double_b1_func = None
        self.row_height_resize_func = None
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
        self.row_width_resize_bbox = ()
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
        self.disp_corners = set()
        self.hidd_text = {}
        self.hidd_high = {}
        self.hidd_grid = {}
        self.hidd_fill_sels = {}
        self.hidd_bord_sels = {}
        self.hidd_resize_lines = {}
        self.hidd_dropdown = {}
        self.hidd_checkbox = {}
        self.hidd_tree_arrow = {}
        self.hidd_corners = set()

        self.align = kwargs["row_index_align"]

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
            for b in self.ops.rc_bindings:
                self.bind(b, self.rc)
        else:
            self.unbind("<Motion>")
            self.unbind("<ButtonPress-1>")
            self.unbind("<B1-Motion>")
            self.unbind("<ButtonRelease-1>")
            self.unbind("<Double-Button-1>")
            for b in self.ops.rc_bindings:
                self.unbind(b)

    def set_width(self, new_width: int, set_TL: bool = False, recreate_selection_boxes: bool = True) -> None:
        try:
            self.config(width=new_width)
        except Exception:
            return
        if set_TL:
            self.TL.set_dimensions(new_w=new_width)
        if isinstance(self.current_width, int) and new_width > self.current_width and recreate_selection_boxes:
            self.MT.recreate_all_selection_boxes()
        self.current_width = new_width

    def is_readonly(self, datarn: int) -> bool:
        return datarn in self.cell_options and "readonly" in self.cell_options[datarn]

    def rc(self, event: Any) -> None:
        self.mouseclick_outside_editor_or_dropdown_all_canvases(inside=True)
        self.focus_set()
        popup_menu = None
        if self.MT.identify_row(y=event.y, allow_end=False) is None:
            self.MT.deselect("all")
            r = len(self.MT.row_positions) - 1
            if self.MT.rc_popup_menus_enabled:
                popup_menu = self.MT.empty_rc_popup_menu
                build_empty_rc_menu(self.MT, popup_menu)
        elif self.row_selection_enabled and not self.currently_resizing_width and not self.currently_resizing_height:
            r = self.MT.identify_row(y=event.y)
            if r < len(self.MT.row_positions) - 1:
                if self.MT.row_selected(r):
                    if self.MT.rc_popup_menus_enabled:
                        popup_menu = self.ri_rc_popup_menu
                        build_index_rc_menu(self.MT, popup_menu, self.MT.selected)
                else:
                    if self.MT.single_selection_enabled and self.MT.rc_select_enabled:
                        self.select_row(r, redraw=True)
                    elif self.MT.toggle_selection_enabled and self.MT.rc_select_enabled:
                        self.toggle_select_row(r, redraw=True)
                    if self.MT.rc_popup_menus_enabled:
                        popup_menu = self.ri_rc_popup_menu
                        build_index_rc_menu(self.MT, popup_menu, self.MT.selected)
        try_binding(self.extra_rc_func, event)
        if popup_menu is not None:
            self.popup_menu_loc = r
            popup_menu.tk_popup(event.x_root, event.y_root)

    def ctrl_b1_press(self, event: Any) -> None:
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

    def ctrl_shift_b1_press(self, event: Any) -> None:
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

    def shift_b1_press(self, event: Any) -> None:
        self.mouseclick_outside_editor_or_dropdown_all_canvases(inside=True)
        r = self.MT.identify_row(y=event.y)
        if (
            (self.drag_and_drop_enabled or self.row_selection_enabled)
            and self.rsz_h is None
            and self.rsz_w is None
            and r < len(self.MT.row_positions) - 1
        ):
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
        if r >= min_r:
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

    def mouse_motion(self, event: Any) -> None:
        if not self.currently_resizing_height and not self.currently_resizing_width:
            x = self.canvasx(event.x)
            y = self.canvasy(event.y)
            mouse_over_resize = False
            if self.height_resizing_enabled:
                r = self.check_mouse_position_height_resizers(x, y)
                if r is not None:
                    self.rsz_h, mouse_over_resize = r, True
                    if self.current_cursor != "sb_v_double_arrow":
                        self.config(cursor="sb_v_double_arrow")
                        self.current_cursor = "sb_v_double_arrow"
                else:
                    self.rsz_h = None
            if (
                self.width_resizing_enabled
                and not mouse_over_resize
                and self.ops.auto_resize_row_index is not True
                and not (
                    self.ops.auto_resize_row_index == "empty"
                    and not isinstance(self.MT._row_index, int)
                    and not self.MT._row_index
                )
            ):
                try:
                    x1, y1, x2, y2 = (
                        self.row_width_resize_bbox[0],
                        self.row_width_resize_bbox[1],
                        self.row_width_resize_bbox[2],
                        self.row_width_resize_bbox[3],
                    )
                    if x >= x1 and y >= y1 and x <= x2 and y <= y2:
                        self.rsz_w, mouse_over_resize = True, True
                        if self.current_cursor != "sb_h_double_arrow":
                            self.config(cursor="sb_h_double_arrow")
                            self.current_cursor = "sb_h_double_arrow"
                    else:
                        self.rsz_w = None
                except Exception:
                    self.rsz_w = None
            if not mouse_over_resize:
                if self.MT.row_selected(self.MT.identify_row(event, allow_end=False)):
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
        elif (self.row_selection_enabled or self.ops.treeview) and self.rsz_h is None and self.rsz_w is None:
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

    def b1_press(self, event: Any) -> None:
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
                fill=self.ops.resizing_line_fg,
                tag=("rh", "rhl"),
            )
            self.MT.create_resize_line(x1, y, x2, y, width=1, fill=self.ops.resizing_line_fg, tag=("rh", "rhl"))
            self.create_resize_line(
                0,
                line2y,
                self.current_width,
                line2y,
                width=1,
                fill=self.ops.resizing_line_fg,
                tag=("rh", "rhl2"),
            )
            self.MT.create_resize_line(
                x1, line2y, x2, line2y, width=1, fill=self.ops.resizing_line_fg, tag=("rh", "rhl2")
            )
        elif self.width_resizing_enabled and self.rsz_h is None and self.rsz_w is True:
            self.currently_resizing_width = True
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

    def b1_motion(self, event: Any) -> None:
        x1, y1, x2, y2 = self.MT.get_canvas_visible_area()
        if self.height_resizing_enabled and self.rsz_h is not None and self.currently_resizing_height:
            y = self.canvasy(event.y)
            size = y - self.MT.row_positions[self.rsz_h - 1]
            if size >= self.MT.min_row_height and size < self.ops.max_row_height:
                self.hide_resize_and_ctrl_lines(ctrl_lines=False)
                line2y = self.MT.row_positions[self.rsz_h - 1]
                self.create_resize_line(
                    0,
                    y,
                    self.current_width,
                    y,
                    width=1,
                    fill=self.ops.resizing_line_fg,
                    tag=("rh", "rhl"),
                )
                self.MT.create_resize_line(x1, y, x2, y, width=1, fill=self.ops.resizing_line_fg, tag=("rh", "rhl"))
                self.create_resize_line(
                    0,
                    line2y,
                    self.current_width,
                    line2y,
                    width=1,
                    fill=self.ops.resizing_line_fg,
                    tag=("rh", "rhl2"),
                )
                self.MT.create_resize_line(
                    x1,
                    line2y,
                    x2,
                    line2y,
                    width=1,
                    fill=self.ops.resizing_line_fg,
                    tag=("rh", "rhl2"),
                )
                self.drag_height_resize()
        elif self.width_resizing_enabled and self.rsz_w is not None and self.currently_resizing_width:
            evx = event.x
            if evx > self.current_width:
                if evx > self.ops.max_index_width:
                    evx = int(self.ops.max_index_width)
                self.drag_width_resize(evx)
            else:
                if evx < self.ops.min_column_width:
                    evx = self.ops.min_column_width
                self.drag_width_resize(evx)
        elif (
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

    def ctrl_b1_motion(self, event: Any) -> None:
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

    def drag_and_drop_motion(self, event: Any) -> float:
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
            fill=self.ops.drag_and_drop_bg,
            tag="move_rows",
        )
        self.MT.create_resize_line(x1, ypos, x2, ypos, width=3, fill=self.ops.drag_and_drop_bg, tag="move_rows")
        for boxst, boxend in consecutive_ranges(rows):
            self.MT.show_ctrl_outline(
                start_cell=(0, boxst),
                end_cell=(len(self.MT.col_positions) - 1, boxend),
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

    def event_over_dropdown(self, r: int, datarn: int, event: Any, canvasy: float) -> bool:
        return (
            canvasy < self.MT.row_positions[r] + self.MT.index_txt_height
            and self.get_cell_kwargs(datarn, key="dropdown")
            and event.x > self.current_width - self.MT.index_txt_height - 4
        )

    def event_over_checkbox(self, r: int, datarn: int, event: Any, canvasy: float) -> bool:
        return (
            canvasy < self.MT.row_positions[r] + self.MT.index_txt_height
            and self.get_cell_kwargs(datarn, key="checkbox")
            and event.x < self.MT.index_txt_height + 4
        )

    def drag_width_resize(self, width: int) -> None:
        self.set_width(width, set_TL=True)
        self.MT.main_table_redraw_grid_and_text(redraw_header=False, redraw_row_index=True, redraw_table=False)

    def drag_height_resize(self) -> None:
        new_row_pos = int(self.coords("rhl")[1])
        old_height = self.MT.row_positions[self.rsz_h] - self.MT.row_positions[self.rsz_h - 1]
        size = new_row_pos - self.MT.row_positions[self.rsz_h - 1]
        if size < self.MT.min_row_height:
            new_row_pos = ceil(self.MT.row_positions[self.rsz_h - 1] + self.MT.min_row_height)
        elif size > self.ops.max_row_height:
            new_row_pos = int(self.MT.row_positions[self.rsz_h - 1] + self.ops.max_row_height)
        increment = new_row_pos - self.MT.row_positions[self.rsz_h]
        self.MT.row_positions[self.rsz_h + 1 :] = [
            e + increment for e in islice(self.MT.row_positions, self.rsz_h + 1, None)
        ]
        self.MT.row_positions[self.rsz_h] = new_row_pos
        new_height = self.MT.row_positions[self.rsz_h] - self.MT.row_positions[self.rsz_h - 1]
        self.MT.allow_auto_resize_rows = False
        self.MT.recreate_all_selection_boxes()
        self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True, set_scrollregion=False)
        if self.row_height_resize_func is not None and old_height != new_height:
            self.row_height_resize_func(
                event_dict(
                    name="resize",
                    sheet=self.PAR.name,
                    resized_rows={self.rsz_h - 1: {"old_size": old_height, "new_size": new_height}},
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
        if self.height_resizing_enabled and self.rsz_h is not None and self.currently_resizing_height:
            self.drag_height_resize()
            self.hide_resize_and_ctrl_lines(ctrl_lines=False)
            self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
        elif (
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
                event_data = self.MT.new_event_dict("move_rows", state=True)
                event_data["value"] = r
                if try_binding(self.ri_extra_begin_drag_drop_func, event_data, "begin_move_rows"):
                    data_new_idxs, disp_new_idxs, event_data = self.MT.move_rows_adjust_options_dict(
                        *self.MT.get_args_for_move_rows(
                            move_to=r,
                            to_move=self.dragged_row.to_move,
                        ),
                        move_data=self.ops.row_drag_and_drop_perform,
                        move_heights=self.ops.row_drag_and_drop_perform,
                        event_data=event_data,
                    )
                    if data_new_idxs:
                        event_data["moved"]["rows"] = {
                            "data": data_new_idxs,
                            "displayed": disp_new_idxs,
                        }
                        if self.MT.undo_enabled:
                            self.MT.undo_stack.append(stored_event_dict(event_data))
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
                        self.select_row(r, ext=True, redraw=False)
                    self.PAR.item(iid, open_=iid not in self.tree_open_ids, undo=False)
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
        if self.ops.treeview and (
            canvasy < self.MT.row_positions[r] + self.MT.index_txt_height + 5
            and isinstance(self.MT._row_index, list)
            and (datarn := self.MT.datarn(r)) < len(self.MT._row_index)
            and eventx
            < (indent := self.get_iid_indent((iid := self.MT._row_index[datarn].iid))) + self.MT.index_txt_height + 4
            and eventx >= indent + 1
        ):
            return iid
        return None

    def _sort_rows(
        self,
        event: tk.Event | None = None,
        rows: Iterator[int] | None = None,
        reverse: bool = False,
        validation: bool = True,
        key: Callable | None = None,
        undo: bool = True,
    ) -> EventDataDict:
        if rows is None:
            rows = self.MT.get_selected_rows()
        if not rows:
            rows = list(range(0, len(self.MT.row_positions) - 1))
        event_data = self.MT.new_event_dict("edit_table")
        try_binding(self.MT.extra_begin_sort_cells_func, event_data)
        if key is None:
            key = self.ops.sort_key
        for r in rows:
            datarn = self.MT.datarn(r)
            for c, val in enumerate(sort_row(self.MT.data[datarn], reverse=reverse, key=key)):
                if (
                    not self.MT.edit_validation_func
                    or not validation
                    or (
                        self.MT.edit_validation_func
                        and (val := self.MT.edit_validation_func(mod_event_val(event_data, val, (datarn, c))))
                        is not None
                    )
                ):
                    event_data = self.MT.event_data_set_cell(
                        datarn=datarn,
                        datacn=c,
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

    def _sort_columns_by_row(
        self,
        event: tk.Event | None = None,
        row: int | None = None,
        reverse: bool = False,
        key: Callable | None = None,
        undo: bool = True,
    ) -> EventDataDict:
        event_data = self.MT.new_event_dict("move_columns", state=True)
        if not self.MT.data:
            return event_data
        if row is None:
            if not self.MT.selected:
                return event_data
            row = self.MT.datarn(self.MT.selected.row)
        if try_binding(self.ri_extra_begin_sort_cols_func, event_data, "begin_move_columns"):
            if key is None:
                key = self.ops.sort_key
            sorted_indices, data_new_idxs = sort_columns_by_row(self.MT.data, row=row, reverse=reverse, key=key)
            disp_new_idxs = {}
            if self.MT.all_columns_displayed:
                disp_new_idxs = data_new_idxs
            else:
                col_ctr = 0
                # idx is the displayed index, can just do range
                for old_idx in sorted_indices:
                    if (idx := try_b_index(self.MT.displayed_columns, old_idx)) is not None:
                        disp_new_idxs[idx] = col_ctr
                        col_ctr += 1
            data_new_idxs, disp_new_idxs, _ = self.PAR.mapping_move_columns(
                data_new_idxs=data_new_idxs,
                disp_new_idxs=disp_new_idxs,
                move_data=True,
                create_selections=False,
                undo=False,
                emit_event=False,
                redraw=True,
            )
            event_data["moved"]["columns"] = {
                "data": data_new_idxs,
                "displayed": disp_new_idxs,
            }
            if undo and self.MT.undo_enabled:
                self.MT.undo_stack.append(stored_event_dict(event_data))
            try_binding(self.ri_extra_end_sort_cols_func, event_data, "end_move_columns")
            self.MT.sheet_modified(event_data)
            self.PAR.emit_event("<<SheetModified>>", event_data)
            self.MT.refresh()
        return event_data

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
        r: int | Iterator[int],
        redraw: bool = False,
        run_binding_func: bool = True,
        ext: bool = False,
    ) -> int:
        boxes_to_hide = tuple(self.MT.selection_boxes)
        fill_iids = [
            self.MT.create_selection_box(
                start,
                0,
                end,
                len(self.MT.col_positions) - 1,
                "rows",
                set_current=True,
                ext=ext,
            )
            for start, end in consecutive_ranges(int_x_tuple(r))
        ]
        for iid in boxes_to_hide:
            self.MT.hide_selection_box(iid)
        if redraw:
            self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
        if run_binding_func:
            self.MT.run_selection_binding("rows")
        return fill_iids[0] if len(fill_iids) == 1 else fill_iids

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

    def get_cell_dimensions(self, datarn: int) -> tuple[int, int]:
        txt = self.cell_str(datarn, fix=False)
        if txt:
            lines = findall(r"[^\n]+", txt)
            h = self.MT.index_txt_height * len(lines) + 5
            w = max(sum(self.char_width_fn(c) for c in line) for line in lines) + 8
        else:
            w = self.ops.default_row_index_width
            h = self.MT.min_row_height
        # self.get_cell_kwargs not used here to boost performance
        if (datarn in self.cell_options and "dropdown" in self.cell_options[datarn]) or (
            datarn in self.cell_options and "checkbox" in self.cell_options[datarn]
        ):
            w += self.MT.index_txt_height + 2
        if self.ops.treeview:
            if datarn in self.cell_options and "align" in self.cell_options[datarn]:
                align = self.cell_options[datarn]["align"]
            else:
                align = self.align
            if align[-1] == "w":
                w += self.MT.index_txt_height
            w += self.get_iid_indent(self.MT._row_index[datarn].iid) + 10
        return w, h

    def get_cell_width(self, datarn: int) -> int:
        txt = self.cell_str(datarn, fix=False)
        if txt:
            w = max(sum(self.char_width_fn(c) for c in line) for line in findall(r"[^\n]+", txt)) + 8
        else:
            w = self.ops.default_row_index_width
        # self.get_cell_kwargs not used here to boost performance
        if (datarn in self.cell_options and "dropdown" in self.cell_options[datarn]) or (
            datarn in self.cell_options and "checkbox" in self.cell_options[datarn]
        ):
            w += self.MT.index_txt_height + 2
        if self.ops.treeview:
            if datarn in self.cell_options and "align" in self.cell_options[datarn]:
                align = self.cell_options[datarn]["align"]
            else:
                align = self.align
            if align[-1] == "w":
                w += self.MT.index_txt_height
            w += self.get_iid_indent(self.MT._row_index[datarn].iid) + 10
        return w

    def get_wrapped_cell_height(self, datarn: int) -> int:
        n_lines = max(
            1,
            sum(
                1
                for _ in wrap_text(
                    text=self.cell_str(datarn, fix=False),
                    max_width=self.current_width,
                    max_lines=float("inf"),
                    char_width_fn=self.char_width_fn,
                    widths=self.MT.char_widths[self.index_font],
                    wrap=self.ops.index_wrap,
                )
            ),
        )
        return 3 + (n_lines * self.MT.index_txt_height)

    def get_row_text_height(
        self,
        row: int,
        visible_only: bool = False,
        only_if_too_small: bool = False,
    ) -> int:
        h = self.MT.min_row_height
        datarn = row if self.MT.all_rows_displayed else self.MT.displayed_rows[row]
        # index
        ih = self.get_wrapped_cell_height(datarn)
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
            cell_heights = (
                self.MT.get_wrapped_cell_height(
                    datarn,
                    datacn,
                )
                for datacn in iterable
            )
            h = max(h, max(cell_heights, default=h))
            self.MT.cells_cache = None
        h = max(h, ih)
        if only_if_too_small and h < self.MT.row_positions[row + 1] - self.MT.row_positions[row]:
            return self.MT.row_positions[row + 1] - self.MT.row_positions[row]
        else:
            return max(int(min(h, self.ops.max_row_height)), self.MT.min_row_height)

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
        elif height > self.ops.max_row_height:
            height = int(self.ops.max_row_height)
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

    def get_index_text_width(self, only_rows: Iterator[int] | None = None) -> int:
        self.fix_index()
        w = self.ops.default_row_index_width
        if (not self.MT._row_index and isinstance(self.MT._row_index, list)) or (
            isinstance(self.MT._row_index, int) and self.MT._row_index >= len(self.MT.data)
        ):
            return w
        if only_rows:
            iterable = only_rows
        elif self.MT.all_rows_displayed:
            if isinstance(self.MT._row_index, list):
                iterable = range(len(self.MT._row_index))
            else:
                iterable = range(len(self.MT.data))
        else:
            iterable = self.MT.displayed_rows
        if (new_w := max(map(self.get_cell_width, iterable), default=w)) > w:
            w = new_w
        if w > self.ops.max_index_width:
            w = int(self.ops.max_index_width)
        return w

    def set_width_of_index_to_text(
        self,
        text: None | str = None,
        only_rows: list[int] | None = None,
    ) -> int:
        self.fix_index()
        w = self.ops.default_row_index_width
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
            w = self.get_index_text_width(only_rows=[] if only_rows is None else only_rows)
        if w > self.ops.max_index_width:
            w = int(self.ops.max_index_width)
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
            if self.ops.default_row_index == "letters":
                new_w = sum(self.char_width_fn(c) for c in num2alpha(end_row)) + 20
            elif self.ops.default_row_index == "numbers":
                new_w = sum(self.char_width_fn(c) for c in str(end_row)) + 20
            elif self.ops.default_row_index == "both":
                new_w = sum(self.char_width_fn(c) for c in f"{end_row + 1} {num2alpha(end_row)}") + 20
            elif self.ops.default_row_index is None:
                new_w = 20
        elif self.ops.auto_resize_row_index is True:
            new_w = self.get_index_text_width(only_rows=only_rows)
        else:
            new_w = None
        if new_w is not None and (sheet_w_x := int(self.PAR.winfo_width() * 0.7)) < new_w:
            new_w = sheet_w_x
        if new_w and (self.current_width - new_w > 20 or new_w - self.current_width > 3):
            if self.MT.find_window.open:
                self.MT.itemconfig(self.MT.find_window.canvas_id, state="hidden")
            self.set_width(new_w, set_TL=True, recreate_selection_boxes=False)
            return True
        return False

    def redraw_highlight_get_text_fg(
        self,
        fr: float,
        sr: float,
        r: int,
        sel_cells_bg: str,
        sel_rows_bg: str,
        selections: dict,
        datarn: int,
        has_dd: bool,
        tags: str | tuple[str],
    ) -> tuple[str, str, bool]:
        redrawn = False
        kwargs = self.get_cell_kwargs(datarn, key="highlight")
        if kwargs:
            high_bg = kwargs[0]
            if high_bg and not high_bg.startswith("#"):
                high_bg = color_map[high_bg]
            if "rows" in selections and r in selections["rows"]:
                txtfg = (
                    self.ops.index_selected_rows_fg
                    if kwargs[1] is None or self.ops.display_selected_fg_over_highlights
                    else kwargs[1]
                )
                redrawn = self.redraw_highlight(
                    0,
                    fr + 1,
                    self.current_width - 1,
                    sr,
                    fill=self.ops.index_selected_rows_bg
                    if high_bg is None
                    else (
                        f"#{int((int(high_bg[1:3], 16) + int(sel_rows_bg[1:3], 16)) / 2):02X}"
                        + f"{int((int(high_bg[3:5], 16) + int(sel_rows_bg[3:5], 16)) / 2):02X}"
                        + f"{int((int(high_bg[5:], 16) + int(sel_rows_bg[5:], 16)) / 2):02X}"
                    ),
                    outline=self.ops.index_fg if has_dd and self.ops.show_dropdown_borders else "",
                    tags=tags,
                )
            elif "cells" in selections and r in selections["cells"]:
                txtfg = (
                    self.ops.index_selected_cells_fg
                    if kwargs[1] is None or self.ops.display_selected_fg_over_highlights
                    else kwargs[1]
                )
                redrawn = self.redraw_highlight(
                    0,
                    fr + 1,
                    self.current_width - 1,
                    sr,
                    fill=self.ops.index_selected_cells_bg
                    if high_bg is None
                    else (
                        f"#{int((int(high_bg[1:3], 16) + int(sel_cells_bg[1:3], 16)) / 2):02X}"
                        + f"{int((int(high_bg[3:5], 16) + int(sel_cells_bg[3:5], 16)) / 2):02X}"
                        + f"{int((int(high_bg[5:], 16) + int(sel_cells_bg[5:], 16)) / 2):02X}"
                    ),
                    outline=self.ops.index_fg if has_dd and self.ops.show_dropdown_borders else "",
                    tags=tags,
                )
            else:
                txtfg = self.ops.index_fg if kwargs[1] is None else kwargs[1]
                if high_bg:
                    redrawn = self.redraw_highlight(
                        0,
                        fr + 1,
                        self.current_width - 1,
                        sr,
                        fill=high_bg,
                        outline=self.ops.index_fg if has_dd and self.ops.show_dropdown_borders else "",
                        tags=tags,
                    )
            tree_arrow_fg = txtfg
        elif not kwargs:
            if "rows" in selections and r in selections["rows"]:
                txtfg = self.ops.index_selected_rows_fg
                tree_arrow_fg = self.ops.selected_rows_tree_arrow_fg
                redrawn = self.redraw_highlight(
                    0,
                    fr + 1,
                    self.current_width - 1,
                    sr,
                    fill=self.ops.index_selected_rows_bg,
                    outline=self.ops.index_fg if has_dd and self.ops.show_dropdown_borders else "",
                    tags=tags,
                )
            elif "cells" in selections and r in selections["cells"]:
                txtfg = self.ops.index_selected_cells_fg
                tree_arrow_fg = self.ops.selected_cells_tree_arrow_fg
                redrawn = self.redraw_highlight(
                    0,
                    fr + 1,
                    self.current_width - 1,
                    sr,
                    fill=self.ops.index_selected_cells_bg,
                    outline=self.ops.index_fg if has_dd and self.ops.show_dropdown_borders else "",
                    tags=tags,
                )
            else:
                txtfg = self.ops.index_fg
                tree_arrow_fg = self.ops.tree_arrow_fg
                redrawn = self.redraw_highlight(
                    0,
                    fr + 1,
                    self.current_width - 1,
                    sr,
                    fill="",
                    outline=self.ops.index_fg if has_dd and self.ops.show_dropdown_borders else "",
                    tags=tags,
                )
        return txtfg, tree_arrow_fg, redrawn

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
        if not self.ops.show_horizontal_grid:
            y2 += 1
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
        points: tuple[float],
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

    def redraw_tree_arrow(
        self,
        x1: float,
        y1: float,
        y2: float,
        fill: str,
        indent: float,
        has_children: bool = False,
        open_: bool = False,
        level: int = 1,
    ) -> None:
        mod = (self.MT.index_txt_height - 1) if self.MT.index_txt_height % 2 else self.MT.index_txt_height
        small_mod = int(mod / 5)
        mid_y = int(self.MT.min_row_height / 2)
        if has_children:
            # up arrow
            if open_:
                points = (
                    # the left hand downward point
                    x1 + 5 + indent,
                    y1 + mid_y + small_mod,
                    # the middle upward point
                    x1 + 5 + indent + small_mod + small_mod,
                    y1 + mid_y - small_mod,
                    # the right hand downward point
                    x1 + 5 + indent + small_mod + small_mod + small_mod + small_mod,
                    y1 + mid_y + small_mod,
                )
            # right pointing arrow
            else:
                points = (
                    # the upper point
                    x1 + 5 + indent + small_mod + small_mod,
                    y1 + mid_y - small_mod - small_mod,
                    # the middle point
                    x1 + 5 + indent + small_mod + small_mod + small_mod + small_mod,
                    y1 + mid_y,
                    # the bottom point
                    x1 + 5 + indent + small_mod + small_mod,
                    y1 + mid_y + small_mod + small_mod,
                )
        else:
            # POINTS FOR A LINE THAT STOPS BELOW FIRST LINE OF TEXT
            # points = (
            #     # the upper point
            #     x1 + 5 + indent + small_mod + small_mod,
            #     y1 + mid_y - small_mod - small_mod,
            #     # the bottom point
            #     x1 + 5 + indent + small_mod + small_mod,
            #     y1 + mid_y + small_mod + small_mod,
            # )

            # POINTS FOR A LINE THAT STOPS AT ROW LINE
            # points = (
            #     # the upper point
            #     x1 + 5 + indent + small_mod + small_mod,
            #     y1 + mid_y - small_mod - small_mod,
            #     # the bottom point
            #     x1 + 5 + indent + small_mod + small_mod,
            #     y2 - mid_y + small_mod + small_mod,
            # )

            # POINTS FOR A HORIZONTAL LINE
            points = (
                # the left point
                x1 + 5 + indent,
                y1 + mid_y,
                # the right point
                x1 + 5 + indent + small_mod + small_mod + small_mod + small_mod,
                y1 + mid_y,
            )

        if self.hidd_tree_arrow:
            t, sh = self.hidd_tree_arrow.popitem()
            self.coords(t, points)
            if sh:
                self.itemconfig(t, fill=fill if has_children else self.ops.index_grid_fg)
            else:
                self.itemconfig(t, fill=fill if has_children else self.ops.index_grid_fg, state="normal")
        else:
            t = self.create_line(
                points,
                fill=fill if has_children else self.ops.index_grid_fg,
                width=2,
                capstyle=tk.ROUND,
                joinstyle=tk.BEVEL,
                tag="lift",
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
        draw_outline: bool = True,
        draw_arrow: bool = True,
        open_: bool = False,
    ) -> None:
        # if draw_outline and self.ops.show_dropdown_borders:
        #     self.redraw_highlight(x1 + 1, y1 + 1, x2, y2, fill="", outline=self.ops.index_fg)
        if draw_arrow:
            mod = (self.MT.index_txt_height - 1) if self.MT.index_txt_height % 2 else self.MT.index_txt_height
            small_mod = int(mod / 5)
            mid_y = int(self.MT.min_row_height / 2)
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
                t = self.create_line(
                    points,
                    fill=fill,
                    width=2,
                    capstyle=tk.ROUND,
                    joinstyle=tk.BEVEL,
                    tag="lift",
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

    def configure_scrollregion(self, last_row_line_pos: float) -> bool:
        try:
            self.configure(
                scrollregion=(
                    0,
                    0,
                    self.current_width,
                    last_row_line_pos + self.ops.empty_vertical + 2,
                )
            )
            return True
        except Exception:
            return False

    def char_width_fn(self, c: str) -> int:
        if c in self.MT.char_widths[self.index_font]:
            return self.MT.char_widths[self.index_font][c]
        else:
            self.MT.txt_measure_canvas.itemconfig(
                self.MT.txt_measure_canvas_text,
                text=_test_str + c,
                font=self.index_font,
            )
            b = self.MT.txt_measure_canvas.bbox(self.MT.txt_measure_canvas_text)
            wd = b[2] - b[0] - self.index_test_str_w
            self.MT.char_widths[self.index_font][c] = wd
            return wd

    def redraw_corner(self, x: float, y: float, tags: str | tuple[str]) -> None:
        if self.hidd_corners:
            iid = self.hidd_corners.pop()
            self.coords(iid, x - 10, y, x, y, x, y + 10)
            self.itemconfig(iid, fill=self.ops.index_grid_fg, state="normal", tags=tags)
            self.disp_corners.add(iid)
        else:
            self.disp_corners.add(
                self.create_polygon(x - 10, y, x, y, x, y + 10, fill=self.ops.index_grid_fg, tags=tags)
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
        set_scrollregion: bool,
    ) -> bool:
        if set_scrollregion and not self.configure_scrollregion(last_row_line_pos=last_row_line_pos):
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
        self.hidd_tree_arrow.update(self.disp_tree_arrow)
        self.disp_tree_arrow = {}
        self.hidd_corners.update(self.disp_corners)
        self.disp_corners = set()
        self.visible_row_dividers = {}
        self.row_width_resize_bbox = (
            self.current_width - 2,
            scrollpos_top,
            self.current_width,
            scrollpos_bot,
        )
        if (self.ops.show_horizontal_grid or self.height_resizing_enabled) and row_pos_exists:
            xend = self.current_width - 6
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
            self.redraw_gridline(points=points, fill=self.ops.index_grid_fg, width=1)
        sel_cells_bg = (
            self.ops.index_selected_cells_bg
            if self.ops.index_selected_cells_bg.startswith("#")
            else color_map[self.ops.index_selected_cells_bg]
        )
        sel_rows_bg = (
            self.ops.index_selected_rows_bg
            if self.ops.index_selected_rows_bg.startswith("#")
            else color_map[self.ops.index_selected_rows_bg]
        )
        font = self.ops.index_font
        selections = self.get_redraw_selections(text_start_row, grid_end_row)
        dd_coords = self.dropdown.get_coords()
        treeview = self.ops.treeview
        wrap = self.ops.index_wrap
        note_corners = self.ops.note_corners
        for r in range(text_start_row, text_end_row):
            rtopgridln = self.MT.row_positions[r]
            rbotgridln = self.MT.row_positions[r + 1]
            if rbotgridln - rtopgridln < self.MT.index_txt_height:
                continue
            checkbox_kwargs = {}
            datarn = r if self.MT.all_rows_displayed else self.MT.displayed_rows[r]
            dropdown_kwargs = self.get_cell_kwargs(datarn, key="dropdown")
            tag = f"{r}"
            fill, tree_arrow_fg, dd_drawn = self.redraw_highlight_get_text_fg(
                fr=rtopgridln,
                sr=rbotgridln,
                r=r,
                sel_cells_bg=sel_cells_bg,
                sel_rows_bg=sel_rows_bg,
                selections=selections,
                datarn=datarn,
                has_dd=bool(dropdown_kwargs),
                tags=("h", "c", tag),
            )
            if datarn in self.cell_options and "align" in self.cell_options[datarn]:
                align = self.cell_options[datarn]["align"]
            else:
                align = self.align
            if dropdown_kwargs:
                max_width = self.current_width - self.MT.index_txt_height - 5
                if align[-1] == "w":
                    draw_x = 3
                elif align[-1] == "e":
                    draw_x = self.current_width - 5 - self.MT.index_txt_height
                elif align[-1] == "n":
                    draw_x = (self.current_width - self.MT.index_txt_height) / 2
                self.redraw_dropdown(
                    0,
                    rtopgridln,
                    self.current_width - 1,
                    rbotgridln - 1,
                    fill=fill if dropdown_kwargs["state"] != "disabled" else self.ops.index_grid_fg,
                    outline=fill,
                    draw_outline=not dd_drawn,
                    draw_arrow=True,
                    open_=dd_coords == r,
                )
            else:
                max_width = self.current_width - 2
                if align[-1] == "w":
                    draw_x = 3
                elif align[-1] == "e":
                    draw_x = self.current_width - 3
                elif align[-1] == "n":
                    draw_x = self.current_width / 2
                if (
                    (checkbox_kwargs := self.get_cell_kwargs(datarn, key="checkbox"))
                    and not dropdown_kwargs
                    and max_width > self.MT.index_txt_height + 1
                ):
                    box_w = self.MT.index_txt_height + 1
                    if align[-1] == "w":
                        draw_x += box_w + 3
                    elif align[-1] == "n":
                        draw_x += box_w / 2 + 1
                    max_width -= box_w + 4
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
                        fill=fill if checkbox_kwargs["state"] == "normal" else self.ops.index_grid_fg,
                        outline="",
                        draw_check=draw_check,
                    )
            if treeview and isinstance(self.MT._row_index, list) and len(self.MT._row_index) > datarn:
                iid = self.MT._row_index[datarn].iid
                max_width -= self.MT.index_txt_height
                if align[-1] == "w":
                    draw_x += self.MT.index_txt_height + 3
                level, indent = self.get_iid_level_indent(iid)
                draw_x += indent + 5
                self.redraw_tree_arrow(
                    2,
                    rtopgridln,
                    rbotgridln - 1,
                    fill=tree_arrow_fg,
                    indent=indent,
                    has_children=bool(self.MT._row_index[datarn].children),
                    open_=self.MT._row_index[datarn].iid in self.tree_open_ids,
                    level=level,
                )
            tags = ("lift", "c", tag)
            if note_corners and max_width > 5 and datarn in self.cell_options and "note" in self.cell_options[datarn]:
                self.redraw_corner(self.current_width, rtopgridln, tags)
            if max_width <= 1:
                continue
            text = self.cell_str(datarn, fix=False)
            if not text:
                continue
            start_line = max(0, int((scrollpos_top - rtopgridln) / self.MT.index_txt_height))
            draw_y = rtopgridln + 3 + (start_line * self.MT.index_txt_height)
            gen_lines = wrap_text(
                text=text,
                max_width=max_width,
                max_lines=int((rbotgridln - rtopgridln - 2) / self.MT.index_txt_height),
                char_width_fn=self.char_width_fn,
                widths=self.MT.char_widths[font],
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
        for iid in self.hidd_corners:
            self.itemconfig(iid, state="hidden")
        self.tag_raise("lift")
        if self.disp_resize_lines:
            self.tag_raise("rh")
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

    def hide_tooltip(self) -> None:
        self.tooltip.withdraw()
        self.tooltip_showing, self.tooltip_coords = False, None

    def show_tooltip(self) -> None:
        r = int(self.tooltip_coords)
        datarn = self.MT.datarn(r)
        kws = self.get_cell_kwargs(datarn, key="note")
        if not self.ops.tooltips and not kws and not self.ops.user_can_create_notes:
            return
        self.MT.hide_tooltip()
        self.CH.hide_tooltip()
        cell_readonly = self.get_cell_kwargs(datarn, "readonly") or not self.MT.index_edit_cell_enabled()
        if kws:
            note = kws["note"]
            note_readonly = kws["readonly"]
        elif self.ops.user_can_create_notes:
            note = ""
            note_readonly = False
        else:
            note = None
            note_readonly = True
        self.tooltip_cell_content = f"{self.get_cell_data(datarn, none_to_empty_str=True)}"
        self.tooltip.reset(
            **{
                "text": self.tooltip_cell_content,
                "cell_readonly": cell_readonly,
                "note": note,
                "note_readonly": note_readonly,
                "row": r,
                "col": 0,
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
            r, _, cell, note = self.tooltip.get()
            datarn = self.MT.datarn(r)
            if not self.tooltip.cell_readonly and cell != self.tooltip_cell_content:
                event_data = self.new_single_edit_event(r, datarn, "??", self.get_cell_data(datarn), cell)
                self.do_single_edit(r, datarn, event_data, cell)
            if not self.tooltip.note_readonly:
                span = self.PAR.span(datarn).options(table=False, index=True)
                self.PAR.note(span, note=note if note else None, readonly=False)
            self.MT.refresh()
        self.hide_tooltip()
        self.focus_set()

    def get_redraw_selections(self, startr: int, endr: int) -> dict[str, set[int]]:
        d = defaultdict(set)
        for _, box in self.MT.get_selection_items():
            r1, _, r2, _ = box.coords
            for r in range(startr, endr):
                if r1 <= r and r2 > r:
                    d[box.type_ if box.type_ != "columns" else "cells"].add(r)
        return d

    def open_cell(self, event: Any = None, ignore_existing_editor: bool = False) -> None:
        if not self.MT.anything_selected() or (not ignore_existing_editor and self.text_editor.open):
            return
        if not self.MT.selected:
            return
        r = self.MT.selected.row
        datarn = r if self.MT.all_rows_displayed else self.MT.displayed_rows[r]
        if self.get_cell_kwargs(datarn, key="readonly"):
            return
        elif self.get_cell_kwargs(datarn, key="dropdown") or self.get_cell_kwargs(datarn, key="checkbox"):
            if event_opens_dropdown_or_checkbox(event):
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
        event: Any = None,
        r: int = 0,
        text: None | str = None,
        state: str = "normal",
        dropdown: bool = False,
    ) -> bool:
        text = f"{self.get_cell_data(self.MT.datarn(r), none_to_empty_str=True, redirect_int=True)}"
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
        if self.ops.cell_auto_resize_enabled:
            self.set_row_height_run_binding(r)
        if self.text_editor.open and r == self.text_editor.row:
            self.text_editor.set_text(self.text_editor.get() + "" if not isinstance(text, str) else text)
            return False
        self.hide_text_editor()
        self.hide_tooltip()
        if not self.MT.see(r, 0, keep_yscroll=True):
            self.MT.main_table_redraw_grid_and_text(True, True)
        x = 0
        y = self.MT.row_positions[r]
        w = self.current_width + 1
        h = self.MT.row_positions[r + 1] - y + 1
        kwargs = {
            "menu_kwargs": get_menu_kwargs(self.ops),
            "sheet_ops": self.ops,
            "font": self.ops.index_font,
            "border_color": self.ops.index_selected_rows_bg,
            "text": text,
            "state": state,
            "width": w,
            "height": h,
            "show_border": True,
            "bg": self.ops.index_editor_bg,
            "fg": self.ops.index_editor_fg,
            "select_bg": self.ops.index_editor_select_bg,
            "select_fg": self.ops.index_editor_select_fg,
            "align": self.get_cell_align(r),
            "r": r,
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

    def text_editor_newline_binding(self, event: Any = None, check_lines: bool = True) -> None:
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
            new_height = curr_height + self.MT.index_txt_height
            space_bot = self.MT.get_space_bot(r)
            if new_height > space_bot:
                new_height = space_bot
            if new_height != curr_height:
                self.set_row_height(r, new_height)
                self.MT.main_table_redraw_grid_and_text(True, True)
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
            self.text_editor.tktext.config(font=self.ops.index_font)
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

    def hide_text_editor(self) -> None:
        if self.text_editor.open:
            for binding in text_editor_to_unbind:
                self.text_editor.tktext.unbind(binding)
            self.itemconfig(self.text_editor.canvas_id, state="hidden")
            self.text_editor.open = False

    def do_single_edit(self, r: int, datarn: int, event_data: EventDataDict, val):
        edited = False
        set_data = partial(self.set_cell_data_undo, r=r, datarn=datarn, check_input_valid=False)
        if self.MT.edit_validation_func:
            val = self.MT.edit_validation_func(event_data)
            if val is not None and self.input_valid_for_cell(datarn, val):
                edited = set_data(value=val)
        elif self.input_valid_for_cell(datarn, val):
            edited = set_data(value=val)
        if edited:
            try_binding(self.extra_end_edit_cell_func, event_data)

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
            self.focus_set()
            return
        text_editor_value = self.text_editor.get()
        r = self.text_editor.row
        datarn = r if self.MT.all_rows_displayed else self.MT.displayed_rows[r]
        event_data = self.new_single_edit_event(r, datarn, event.keysym, self.get_cell_data(datarn), text_editor_value)
        self.do_single_edit(r, datarn, event_data, text_editor_value)
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
                win_h += self.MT.index_txt_height + (v_numlines * self.MT.index_txt_height) + 5  # end of cell
            else:
                win_h += self.MT.min_row_height
            if i == 5:
                break
        win_h = min(win_h, 500)
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
        event: tk.Misc,
    ) -> None:
        r = self.dropdown.get_coords()
        event_data = event_dict(
            name="table_dropdown_modified",
            sheet=self.PAR.name,
            value=self.text_editor.get(),
            loc=r,
            row=r,
            column=0,
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

    def open_dropdown_window(self, r: int, event: Any = None) -> None:
        self.hide_text_editor()
        kwargs = self.get_cell_kwargs(self.MT.datarn(r), key="dropdown")
        if kwargs["state"] == "disabled":
            return
        if kwargs["state"] == "normal" and not self.open_text_editor(event=event, r=r, dropdown=True):
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
            "bg": self.ops.index_editor_bg,
            "fg": self.ops.index_editor_fg,
            "select_bg": self.ops.index_editor_select_bg,
            "select_fg": self.ops.index_editor_select_fg,
            "width": win_w,
            "height": win_h,
            "font": self.ops.index_font,
            "ops": self.ops,
            "outline_color": self.ops.index_selected_rows_bg,
            "align": self.get_cell_align(r),
            "values": kwargs["values"],
            "search_function": kwargs["search_function"],
            "modified_function": kwargs["modified_function"],
        }
        if self.dropdown.window:
            self.dropdown.window.reset(**reset_kwargs)
            self.itemconfig(self.dropdown.canvas_id, state="normal", anchor=anchor)
            self.coords(self.dropdown.canvas_id, 0, ypos)
            self.dropdown.window.tkraise()
        else:
            self.dropdown.window = self.PAR._dropdown_cls(
                self.winfo_toplevel(),
                **reset_kwargs,
                single_index="r",
                close_dropdown_window=self.close_dropdown_window,
                arrowkey_RIGHT=self.MT.arrowkey_RIGHT,
                arrowkey_LEFT=self.MT.arrowkey_LEFT,
            )
            self.dropdown.canvas_id = self.create_window(
                (0, ypos),
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
            self.dropdown.window.bind("<FocusOut>", lambda _x: self.close_dropdown_window(r))
            self.dropdown.window.bind("<Escape>", self.close_dropdown_window)
            self.dropdown.window.focus_set()
            redraw = True
        self.dropdown.open = True
        if redraw:
            self.MT.main_table_redraw_grid_and_text(redraw_header=False, redraw_row_index=True, redraw_table=False)

    # r is displayed row
    def close_dropdown_window(self, r: int | None = None, selection: Any = None, redraw: bool = True) -> None:
        if r is not None and selection is not None:
            datarn = r if self.MT.all_rows_displayed else self.MT.displayed_rows[r]
            kwargs = self.get_cell_kwargs(datarn, key="dropdown")
            event_data = self.new_single_edit_event(r, datarn, "??", self.get_cell_data(datarn), selection)
            try_binding(kwargs["select_function"], event_data)
            selection = selection if not self.MT.edit_validation_func else self.MT.edit_validation_func(event_data)
            if selection is not None:
                edited = self.set_cell_data_undo(r, datarn=datarn, value=selection, redraw=not redraw)
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
            dispcn = self.MT.try_dispcn(self.MT._row_index)
            edited = self.MT.set_cell_data_undo(
                r=r,
                c=dispcn if isinstance(dispcn, int) else 0,
                datarn=datarn,
                datacn=self.MT._row_index,
                value=value,
                undo=True,
                cell_resize=isinstance(dispcn, int),
            )
        else:
            self.fix_index(datarn)
            if not check_input_valid or self.input_valid_for_cell(datarn, value):
                if self.MT.undo_enabled and undo:
                    self.MT.undo_stack.append(stored_event_dict(event_data))
                self.set_cell_data(datarn=datarn, value=value)
                edited = True
        if edited and cell_resize and self.ops.cell_auto_resize_enabled:
            self.set_row_height_run_binding(r, only_if_too_small=False)
        if redraw:
            self.MT.refresh()
        if edited:
            self.MT.sheet_modified(event_data)
        return edited

    def set_cell_data(self, datarn: int | None = None, value: Any = "") -> None:
        if isinstance(self.MT._row_index, int):
            self.MT.set_cell_data(datarn=datarn, datacn=self.MT._row_index, value=value)
        else:
            self.fix_index(datarn)
            if self.get_cell_kwargs(datarn, key="checkbox"):
                self.MT._row_index[datarn] = try_to_bool(value)
            elif self.ops.treeview:
                self.MT._row_index[datarn].text = value
            else:
                self.MT._row_index[datarn] = value

    def input_valid_for_cell(self, datarn: int, value: Any, check_readonly: bool = True) -> bool:
        if check_readonly and self.get_cell_kwargs(datarn, key="readonly"):
            return False
        elif self.get_cell_kwargs(datarn, key="checkbox"):
            return is_bool_like(value)
        else:
            return not (
                self.cell_equal_to(datarn, value)
                or (
                    (kwargs := self.get_cell_kwargs(datarn, key="dropdown"))
                    and kwargs["validate_input"]
                    and value not in kwargs["values"]
                )
            )

    def cell_equal_to(self, datarn: int, value: Any) -> bool:
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
    ) -> Any:
        if get_displayed:
            return self.cell_str(datarn, fix=False)
        if redirect_int and isinstance(self.MT._row_index, int):  # internal use
            return self.MT.get_cell_data(datarn, self.MT._row_index, none_to_empty_str=True)
        if (
            isinstance(self.MT._row_index, int)
            or not self.MT._row_index
            or datarn >= len(self.MT._row_index)
            or (self.MT._row_index[datarn] is None and none_to_empty_str)
        ):
            return ""
        if self.ops.treeview:
            return self.MT._row_index[datarn].text
        return self.MT._row_index[datarn]

    def cell_str(self, datarn: int, fix: bool = True) -> str:
        kwargs = self.get_cell_kwargs(datarn, key="dropdown")
        if kwargs:
            if kwargs["text"] is not None:
                return f"{kwargs['text']}"
        else:
            kwargs = self.get_cell_kwargs(datarn, key="checkbox")
            if kwargs:
                return f"{kwargs['text']}"
        if isinstance(self.MT._row_index, int):
            return self.MT.cell_str(datarn, self.MT._row_index, get_displayed=True)
        if fix:
            self.fix_index(datarn)
        try:
            value = self.MT._row_index[datarn]
            if value is None:
                value = ""
            elif isinstance(value, Node):
                value = value.text
            elif not isinstance(value, str):
                value = f"{value}"
        except Exception:
            value = ""
        if not value and self.ops.show_default_index_for_empty:
            value = get_n2a(datarn, self.ops.default_row_index)
        return value

    def get_value_for_empty_cell(self, datarn: int, r_ops: bool = True) -> Any:
        if self.ops.treeview:
            iid = self.new_iid()
            return Node(text=iid, iid=iid, parent=self.get_row_parent(datarn))
        if self.get_cell_kwargs(datarn, key="checkbox", cell=r_ops):
            return False
        elif (kwargs := self.get_cell_kwargs(datarn, key="dropdown", cell=r_ops)) and kwargs["validate_input"]:
            if kwargs["default_value"] is None:
                if kwargs["values"]:
                    return safe_copy(kwargs["values"][0])
                else:
                    return ""
            else:
                return safe_copy(kwargs["default_value"])
        else:
            return ""

    def get_empty_index_seq(self, end: int, start: int = 0, r_ops: bool = True) -> list[Any]:
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
                    if isinstance(self.MT.data[datarn][self.MT._row_index], bool)
                    else False
                )
            else:
                value = False
            self.set_cell_data_undo(r, datarn=datarn, value=value, cell_resize=False)
            event_data = self.new_single_edit_event(r, datarn, "??", pre_edit_value, value)
            if kwargs["check_function"] is not None:
                kwargs["check_function"](event_data)
            try_binding(self.extra_end_edit_cell_func, event_data)
        if redraw:
            self.MT.refresh()

    def get_cell_kwargs(self, datarn: int, key: Hashable | None = "dropdown", cell: bool = True) -> dict:
        if cell and datarn in self.cell_options:
            return self.cell_options[datarn] if key is None else self.cell_options[datarn].get(key, {})
        else:
            return {}

    def new_single_edit_event(self, r: int, datarn: int, k: str, before_val: Any, after_val: Any) -> EventDataDict:
        return event_dict(
            name="end_edit_index",
            sheet=self.PAR.name,
            widget=self,
            cells_index={datarn: before_val},
            key=k,
            value=after_val,
            loc=r,
            row=r,
            boxes=self.MT.get_boxes(),
            selected=self.MT.selected,
            data={datarn: after_val},
        )

    # Treeview Mode

    def tree_reset(self) -> None:
        self.tree_open_ids = set()
        self.rns = {}
        if self.MT:
            self.MT.displayed_rows = []
            self.MT._row_index = []
            self.MT.data = []
            self.MT.row_positions = [0]
            self.MT.saved_row_heights = {}

    def new_iid(self) -> str:
        self.new_iid_ctr += 1
        while (iid := f"{num2alpha(self.new_iid_ctr)}") in self.rns:
            self.new_iid_ctr += 1
        return iid

    def get_row_parent(self, r: int) -> str:
        if r >= len(self.MT._row_index):
            return ""
        else:
            return self.MT._row_index[r].parent

    def tree_del_rows(self, event_data: EventDataDict) -> EventDataDict:
        event_data["treeview"]["nodes"] = {}
        for node in event_data["deleted"]["index"].values():
            iid = node.iid
            if parent_node := self.parent_node(iid):
                if parent_node.iid not in event_data["treeview"]["nodes"]:
                    event_data["treeview"]["nodes"][parent_node.iid] = Node(
                        text=parent_node.text,
                        iid=parent_node.iid,
                        parent=parent_node.parent,
                        children=parent_node.children.copy(),
                    )
                self.remove_iid_from_parents_children(iid)
        for node in event_data["deleted"]["index"].values():
            iid = node.iid
            for did in self.get_iid_descendants(iid):
                self.tree_open_ids.discard(did)
            self.tree_open_ids.discard(iid)

        for datarn in reversed(event_data["deleted"]["index"]):
            del self.MT._row_index[datarn]

        return event_data

    def tree_add_rows(self, event_data: EventDataDict) -> EventDataDict:
        for rn, node in event_data["added"]["rows"]["index"].items():
            self.rns[node.iid] = rn

        if event_data["treeview"]["nodes"]:
            self.restore_nodes(event_data=event_data)
        else:
            row, a_node = next(iter(event_data["added"]["rows"]["index"].items()))
            if parent := a_node.parent:
                if self.iid_children(parent):
                    index = next(
                        (i for i, cid in enumerate(self.iid_children(parent)) if self.rns[cid] >= row),
                        len(self.iid_children(parent)),
                    )
                    self.iid_children(parent)[index:index] = [
                        n.iid for n in event_data["added"]["rows"]["index"].values()
                    ]
                else:
                    self.iid_children(parent).extend(n.iid for n in event_data["added"]["rows"]["index"].values())

            # handle displaying the new rows
            event_data["added"]["rows"]["row_heights"] = {}
            # parent exists and it's displayed and it's open
            if parent and self.PAR.item_displayed(parent) and parent in self.tree_open_ids:
                self.MT.displayed_rows = add_to_displayed(self.MT.displayed_rows, event_data["added"]["rows"]["index"])
                disp_idx = self.MT.disprn(self.rns[a_node.iid])  # first node, they're contiguous because not undo
                h = self.MT.get_default_row_height()
                for i in range(len(event_data["added"]["rows"]["index"])):
                    event_data["added"]["rows"]["row_heights"][disp_idx + i] = h
            # parent exists and either it's not displayed or not open
            elif parent:
                self.MT.displayed_rows = push_displayed(self.MT.displayed_rows, event_data["added"]["rows"]["index"])
            # no parent, top level
            elif not parent:
                self.MT.displayed_rows = add_to_displayed(self.MT.displayed_rows, event_data["added"]["rows"]["index"])
                h = self.MT.get_default_row_height()
                for datarn in event_data["added"]["rows"]["index"]:
                    event_data["added"]["rows"]["row_heights"][self.MT.disprn(datarn)] = h

        return event_data

    def remove_descendants(self, iids: set[str]) -> tuple[list[int], set[str]]:
        return (
            sorted(
                self.rns[iid] for iid in iids if not any(ancestor in iids for ancestor in self.get_iid_ancestors(iid))
            ),
            iids,
        )

    def move_rows_mod_nodes(
        self,
        data_new_idxs: dict[int, int],
        data_old_idxs: dict[int, int],
        disp_new_idxs: dict[int, int],
        maxidx: int,
        event_data: EventDataDict,
        undo_modification: EventDataDict | None = None,
        node_change: dict | None = None,
    ) -> Generator[tuple[dict[int, int], dict[int, int], dict[str, int], EventDataDict]] | None:
        # data_new_idxs is {old: new, old: new}
        # data_old_idxs is {new: old, new: old}
        if not event_data["treeview"]["nodes"]:
            if undo_modification:
                """
                Used by undo/redo
                """
                event_data = self.copy_nodes(undo_modification["treeview"]["nodes"], event_data)
                self.restore_nodes(undo_modification)

            elif event_data["moved"]["rows"]:
                if node_change:
                    moved_rows, iids = self.remove_descendants(node_change["iids"])
                    new_parent = node_change["new_par"]
                    move_to_index = node_change["index"]
                    if new_parent:
                        new_par_cn = self.iid_children(new_parent)
                        len_new_par_cn = len(new_par_cn)
                        if isinstance(move_to_index, int):
                            move_to_index = min(move_to_index, len_new_par_cn)
                        else:
                            move_to_index = len_new_par_cn
                        if not new_par_cn:
                            insert_row = self.rns[new_parent] + 1
                        elif move_to_index >= len_new_par_cn:
                            insert_row = self.rns[new_par_cn[-1]] + self.num_descendants(new_par_cn[-1]) + 1
                        else:
                            # To get the ids to the proper index and insert row -
                            # For every iid that is being moved under the same parent but to an
                            # index further along we have to adjust the insert row and move to index
                            insert_row = self.rns[new_par_cn[move_to_index]]
                            for iid in iids:
                                if new_parent == self.iid_parent(iid) and move_to_index > self.PAR.index(iid):
                                    insert_row += 1 + self.num_descendants(new_par_cn[move_to_index])
                                    move_to_index += 1
                                    if move_to_index >= len_new_par_cn:
                                        break
                    else:
                        if isinstance(move_to_index, int):
                            move_to_index = min(move_to_index, sum(1 for _ in self.gen_top_nodes()))
                        else:
                            move_to_index = sum(1 for _ in self.gen_top_nodes())
                        insert_row = self.PAR._get_id_insert_row(move_to_index, new_parent)
                else:
                    moved_rows, iids = self.remove_descendants(
                        {self.MT._row_index[r].iid for r in event_data["moved"]["rows"]["data"]}
                    )
                    if isinstance(event_data.value, int):
                        if event_data.value >= len(self.MT.displayed_rows):
                            insert_row = len(self.MT._row_index)
                        else:
                            insert_row = self.MT.datarn(event_data.value)
                        move_to_iid = self.MT._row_index[min(insert_row, len(self.MT._row_index) - 1)].iid
                    else:
                        min_from = min(event_data["moved"]["rows"]["data"])
                        min_to = min(event_data["moved"]["rows"]["data"].values())
                        max_to = max(event_data["moved"]["rows"]["data"].values())
                        insert_row = max_to if min_from <= min_to else min_to
                        move_to_iid = self.MT._row_index[insert_row].iid
                    move_to_index = self.PAR.index(move_to_iid)
                    new_parent = self.iid_parent(move_to_iid)
                    if any(
                        self.move_pid_causes_recursive_loop(self.MT._row_index[r].iid, new_parent) for r in moved_rows
                    ):
                        event_data["moved"]["rows"], data_new_idxs, data_old_idxs = {}, {}, {}
                        moved_rows = []

                new_loc_is_displayed = not new_parent or (
                    new_parent and new_parent in self.tree_open_ids and self.PAR.item_displayed(new_parent)
                )
                event_data["moved"]["rows"]["displayed"], disp_new_idxs = {}, {}

                if moved_rows:
                    event_data["moved"]["rows"]["data"] = {moved_rows[0]: insert_row}
                    iids = []
                    for r in moved_rows:
                        iid = self.MT._row_index[r].iid
                        event_data = self.move_node(event_data=event_data, item=iid, parent=new_parent)
                        iids.append(iid)
                    if new_parent:
                        self.iid_node(new_parent).children[move_to_index:move_to_index] = iids
                        self.iid_node(new_parent).children = remove_duplicates_outside_section(
                            self.iid_children(new_parent), move_to_index, len(iids)
                        )
                    event_data["moved"]["rows"]["data"] = get_new_indexes(
                        insert_row,
                        event_data["moved"]["rows"]["data"],
                    )
                    data_new_idxs = event_data["moved"]["rows"]["data"]
                    data_old_idxs = {v: k for k, v in data_new_idxs.items()}

        if data_new_idxs:
            self.MT.move_rows_data(data_new_idxs, data_old_idxs, maxidx)

        yield data_new_idxs, data_old_idxs, disp_new_idxs, event_data

        if not undo_modification and data_new_idxs:
            if new_parent and (not self.PAR.item_displayed(new_parent) or new_parent not in self.tree_open_ids):
                self.PAR.hide_rows(set(data_new_idxs.values()), data_indexes=True)
            if new_loc_is_displayed:
                self.PAR.show_rows(
                    (r for r in data_new_idxs.values() if self.ancestors_all_open(self.MT._row_index[r].iid))
                )

        yield None

    def move_node(self, event_data: EventDataDict, item: str, parent: str | None = None) -> EventDataDict:
        # also backs up nodes
        if parent is None:
            parent = self.iid_parent(item)
        item_node = self.iid_node(item)
        if parent:
            parent_node = self.iid_node(parent)
            # its the same parent, we're just moving index
            if parent == item_node.parent:
                event_data = self.copy_nodes((item, parent), event_data)
            else:
                if item_node.parent:
                    event_data = self.copy_nodes((item, item_node.parent, parent), event_data)
                else:
                    event_data = self.copy_nodes((item, parent), event_data)
                self.remove_iid_from_parents_children(item)
                item_node.parent = parent_node.iid
        else:
            if item_node.parent:
                event_data = self.copy_nodes((item, item_node.parent), event_data)
            else:
                event_data = self.copy_nodes((item,), event_data)
            self.remove_iid_from_parents_children(item)
            self.MT._row_index[self.rns[item]].parent = ""

        # last row in mapping + 1 is where to start from
        mapping = event_data["moved"]["rows"]["data"]
        row_ctr = next(reversed(mapping.values())) + 1

        rn = self.rns[item]
        if rn not in mapping:
            mapping[rn] = row_ctr
            row_ctr += 1
        for did in self.get_iid_descendants(item):
            mapping[self.rns[did]] = row_ctr
            row_ctr += 1

        event_data["moved"]["rows"]["data"] = mapping
        return event_data

    def restore_nodes(self, event_data: EventDataDict) -> None:
        for iid, node in event_data["treeview"]["nodes"].items():
            self.MT._row_index[self.rns[iid]] = node

    def copy_nodes(self, items: Iterator[str], event_data: EventDataDict) -> EventDataDict:
        for iid in items:
            if iid not in event_data["treeview"]["nodes"]:
                n = self.iid_node(iid)
                event_data["treeview"]["nodes"][iid] = Node(
                    text=n.text,
                    iid=n.iid,
                    parent=n.parent,
                    children=n.children.copy(),
                )
        return event_data

    def get_node_level(self, node: Node) -> int:
        level = 0
        while node.parent:
            level += 1
            node = self.MT._row_index[self.rns[node.parent]]
        return level

    def ancestors_all_open(self, iid: str, stop_at: str = "") -> bool:
        if stop_at:
            for i in self.get_iid_ancestors(iid):
                if i == stop_at:
                    return True
                elif i not in self.tree_open_ids:
                    return False
            return True
        return all(map(self.tree_open_ids.__contains__, self.get_iid_ancestors(iid)))

    def get_iid_ancestors(self, iid: str) -> Generator[str]:
        current_iid = iid
        while self.MT._row_index[self.rns[current_iid]].parent:
            parent_iid = self.MT._row_index[self.rns[current_iid]].parent
            yield parent_iid
            current_iid = parent_iid

    def get_iid_descendants(self, iid: str, check_open: bool = False) -> Generator[str]:
        index = self.MT._row_index
        stack = [iter(index[self.rns[iid]].children)]
        while stack:
            top_iterator = stack[-1]
            try:
                ciid = next(top_iterator)
                yield ciid
                if index[self.rns[ciid]].children and (not check_open or ciid in self.tree_open_ids):
                    stack.append(iter(index[self.rns[ciid]].children))
            except StopIteration:
                stack.pop()

    def num_descendants(self, iid: str) -> int:
        index = self.MT._row_index
        stack = [iter(index[self.rns[iid]].children)]
        num = 0
        while stack:
            top_iterator = stack[-1]
            try:
                ciid = next(top_iterator)
                num += 1
                if index[self.rns[ciid]].children:
                    stack.append(iter(index[self.rns[ciid]].children))
            except StopIteration:
                stack.pop()
        return num

    def iid_node(self, iid: str) -> Node:
        return self.MT._row_index[self.rns[iid]]

    def iid_parent(self, iid: str) -> str:
        return self.MT._row_index[self.rns[iid]].parent

    def iid_children(self, iid: str) -> list[str]:
        return self.MT._row_index[self.rns[iid]].children

    def parent_node(self, iid: str) -> Node | Literal[""]:
        if self.MT._row_index[self.rns[iid]].parent:
            return self.MT._row_index[self.rns[self.MT._row_index[self.rns[iid]].parent]]
        else:
            return ""

    def rename_iid(self, old: str, new: str, event_data: EventDataDict) -> EventDataDict:
        event_data.treeview.renamed[old] = new
        for ciid in self.iid_children(old):
            self.MT._row_index[self.rns[ciid]].parent = new
        if self.iid_parent(old):
            parent_node = self.parent_node(old)
            item_index = parent_node.children.index(old)
            parent_node.children[item_index] = new
        self.MT._row_index[self.rns[old]].iid = new
        self.rns[new] = self.rns.pop(old)
        if old in self.tree_open_ids:
            self.tree_open_ids.discard(old)
            self.tree_open_ids.add(new)
        return event_data

    def gen_top_nodes(self) -> Generator[Node]:
        yield from (node for node in self.MT._row_index if node.parent == "")

    def get_iid_indent(self, iid: str) -> int:
        if isinstance(self.ops.treeview_indent, str):
            indent = self.MT.index_txt_width * int(self.ops.treeview_indent)
        else:
            indent = self.ops.treeview_indent
        return indent * self.get_node_level(self.iid_node(iid))

    def get_iid_level_indent(self, iid: str) -> tuple[int, int]:
        if isinstance(self.ops.treeview_indent, str):
            indent = self.MT.index_txt_width * int(self.ops.treeview_indent)
        else:
            indent = self.ops.treeview_indent
        level = self.get_node_level(self.iid_node(iid))
        return level, indent * level

    def remove_iid_from_parents_children(self, iid: str) -> None:
        if parent_node := self.parent_node(iid):
            parent_node.children.remove(iid)
            if not parent_node.children:
                self.tree_open_ids.discard(parent_node.iid)

    def move_pid_causes_recursive_loop(self, to_move_iid: str, move_to_parent: str) -> bool:
        # if the parent the item is being moved under is one of the item's descendants
        # then it is a recursive loop
        return any(move_to_parent == diid for diid in self.get_iid_descendants(to_move_iid))

    def tree_build(
        self,
        data: list[list[Any]],
        iid_column: int,
        parent_column: int,
        text_column: None | int | list[str] = None,
        push_ops: bool = False,
        row_heights: Sequence[int] | None | False = None,
        open_ids: Iterator[str] | None = None,
        ncols: int | None = None,
        lower: bool = False,
        include_iid_column: bool = True,
        include_parent_column: bool = True,
        include_text_column: bool = True,
    ) -> None:
        data_rns = {}
        tree = {}
        if text_column is None:
            text_column = iid_column
        if not isinstance(ncols, int):
            ncols = max(map(len, data), default=0)
        for rn, row in enumerate(data):
            if lower:
                iid = row[iid_column].lower()
                pid = row[parent_column].lower()
            else:
                iid = row[iid_column]
                pid = row[parent_column]
            if iid in tree:
                tree[iid].text = row[text_column] if isinstance(text_column, int) else text_column[rn]
            else:
                tree[iid] = Node(
                    text=row[text_column] if isinstance(text_column, int) else text_column[rn],
                    iid=iid,
                    parent="",
                )
            data_rns[iid] = rn
            if pid:
                if pid in tree:
                    tree[pid].children.append(iid)
                else:
                    tree[pid] = Node(
                        text=pid,
                        iid=pid,
                        children=[iid],
                    )
                tree[iid].parent = pid
            else:
                tree[iid].parent = ""
        exclude = set()
        if not include_iid_column:
            exclude.add(iid_column)
        if not include_parent_column:
            exclude.add(parent_column)
        if isinstance(text_column, int) and not include_text_column:
            exclude.add(text_column)
        rows = []
        ctr = 0
        if exclude:
            for iid, node in tree.items():
                if node.parent == "":
                    row = [tree[iid]]
                    row.extend(e for i, e in enumerate(data[data_rns[iid]]) if i not in exclude)
                    rows.append(row)
                    self.rns[iid] = ctr
                    ctr += 1
                    for diid in self._build_get_descendants(iid, tree):
                        row = [tree[diid]]
                        row.extend(e for i, e in enumerate(data[data_rns[diid]]) if i not in exclude)
                        rows.append(row)
                        self.rns[diid] = ctr
                        ctr += 1
        else:
            for iid, node in tree.items():
                if node.parent == "":
                    row = [tree[iid]]
                    row.extend(data[data_rns[iid]])
                    rows.append(row)
                    self.rns[iid] = ctr
                    ctr += 1
                    for diid in self._build_get_descendants(iid, tree):
                        row = [tree[diid]]
                        row.extend(data[data_rns[diid]])
                        rows.append(row)
                        self.rns[diid] = ctr
                        ctr += 1
        self.PAR.insert_rows(
            rows=rows,
            idx=0,
            heights=[] if row_heights is False else row_heights,
            row_index=True,
            create_selections=False,
            fill=False,
            undo=False,
            push_ops=push_ops,
            redraw=False,
            tree=False,
        )
        self.MT.all_rows_displayed = False
        self.MT.displayed_rows = list(range(len(self.MT._row_index)))
        if open_ids:
            self.PAR.tree_set_open(open_ids=open_ids)
        else:
            index = self.MT._row_index
            self.PAR.hide_rows(
                {self.rns[iid] for iid in self.PAR.get_children() if index[self.rns[iid]].parent},
                deselect_all=False,
                data_indexes=True,
                row_heights=row_heights is not False,
            )

    def safe_tree_build(
        self,
        data: list[list[Any]],
        iid_column: int,
        parent_column: int,
        text_column: None | int | list[str] = None,
        push_ops: bool = False,
        row_heights: Sequence[int] | None | False = None,
        open_ids: Iterator[str] | None = None,
        ncols: int | None = None,
        lower: bool = False,
        include_iid_column: bool = True,
        include_parent_column: bool = True,
        include_text_column: bool = True,
    ) -> None:
        data_rns = {}
        tree = {}
        iids_missing_rows = set()
        if text_column is None:
            text_column = iid_column
        tally_of_ids = defaultdict(lambda: -1)
        if not isinstance(ncols, int):
            ncols = max(map(len, data), default=0)
        for rn, row in enumerate(data):
            if ncols > (lnr := len(row)):
                row += self.MT.get_empty_row_seq(rn, end=ncols, start=lnr)
            if lower:
                iid = row[iid_column].lower()
                pid = row[parent_column].lower()
            else:
                iid = row[iid_column]
                pid = row[parent_column]
            if not iid:
                continue
            tally_of_ids[iid] += 1
            if tally_of_ids[iid]:
                x = 1
                while iid in tally_of_ids:
                    new = f"{row[iid_column]}_DUPLICATED_{x}"
                    iid = new.lower() if lower else new
                    x += 1
                tally_of_ids[iid] += 1
                row[iid_column] = new
            if iid in tree:
                tree[iid].text = row[text_column] if isinstance(text_column, int) else text_column[rn]
            else:
                tree[iid] = Node(
                    text=row[text_column] if isinstance(text_column, int) else text_column[rn],
                    iid=iid,
                    parent="",
                )
            if iid in iids_missing_rows:
                iids_missing_rows.discard(iid)
            data_rns[iid] = rn
            if iid == pid or self.build_pid_causes_recursive_loop(iid, pid, tree):
                row[parent_column] = ""
                pid = ""
            if pid:
                if pid in tree:
                    tree[pid].children.append(iid)
                else:
                    tree[pid] = Node(
                        text=pid,
                        iid=pid,
                        children=[iid],
                    )
                    iids_missing_rows.add(pid)
                tree[iid].parent = pid
            else:
                tree[iid].parent = ""
        empty_rows = {}
        for iid in iids_missing_rows:
            node = tree[iid]
            newrow = self.MT.get_empty_row_seq(len(data), ncols)
            newrow[iid_column] = node.iid
            empty_rows[node.iid] = newrow
        exclude = set()
        if not include_iid_column:
            exclude.add(iid_column)
        if not include_parent_column:
            exclude.add(parent_column)
        if isinstance(text_column, int) and not include_text_column:
            exclude.add(text_column)
        rows = []
        ctr = 0
        if exclude:
            for iid, node in tree.items():
                if node.parent == "":
                    row = [tree[iid]]
                    if iid in empty_rows:
                        row.extend(e for i, e in enumerate(empty_rows[iid]) if i not in exclude)
                    else:
                        row.extend(e for i, e in enumerate(data[data_rns[iid]]) if i not in exclude)
                    rows.append(row)
                    self.rns[iid] = ctr
                    ctr += 1
                    for diid in self._build_get_descendants(iid, tree):
                        row = [tree[diid]]
                        row.extend(e for i, e in enumerate(data[data_rns[diid]]) if i not in exclude)
                        rows.append(row)
                        self.rns[diid] = ctr
                        ctr += 1
        else:
            for iid, node in tree.items():
                if node.parent == "":
                    row = [tree[iid]]
                    if iid in empty_rows:
                        row.extend(empty_rows[iid])
                    else:
                        row.extend(data[data_rns[iid]])
                    rows.append(row)
                    self.rns[iid] = ctr
                    ctr += 1
                    for diid in self._build_get_descendants(iid, tree):
                        row = [tree[diid]]
                        row.extend(data[data_rns[diid]])
                        rows.append(row)
                        self.rns[diid] = ctr
                        ctr += 1
        self.PAR.insert_rows(
            rows=rows,
            idx=0,
            heights=[] if row_heights is False else row_heights,
            row_index=True,
            create_selections=False,
            fill=False,
            undo=False,
            push_ops=push_ops,
            redraw=False,
            tree=False,
        )
        self.MT.all_rows_displayed = False
        self.MT.displayed_rows = list(range(len(self.MT._row_index)))
        if open_ids:
            self.PAR.tree_set_open(open_ids=open_ids)
        else:
            index = self.MT._row_index
            self.PAR.hide_rows(
                {self.rns[iid] for iid in self.PAR.get_children() if index[self.rns[iid]].parent},
                deselect_all=False,
                data_indexes=True,
                row_heights=row_heights is not False,
            )

    def _build_get_descendants(self, iid: str, tree: dict[str, Node]) -> Generator[str]:
        stack = [iter(tree[iid].children)]
        while stack:
            top_iterator = stack[-1]
            try:
                ciid = next(top_iterator)
                yield ciid
                if tree[ciid].children:
                    stack.append(iter(tree[ciid].children))
            except StopIteration:
                stack.pop()

    def build_pid_causes_recursive_loop(self, iid: str, pid: str, tree: dict[str, Node]) -> bool:
        # check descendants
        for diid in self._build_get_descendants(iid, tree):
            if diid == pid:
                return True
        # check ancestors
        current_iid = iid
        while tree[current_iid].parent:
            parent_iid = tree[current_iid].parent
            if parent_iid == pid:
                return True
            current_iid = parent_iid
        return False
