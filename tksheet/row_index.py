from __future__ import annotations

import tkinter as tk
from collections import defaultdict
from itertools import (
    chain,
    cycle,
    islice,
)
from math import (
    ceil,
    floor,
)

from .formatters import (
    is_bool_like,
    try_to_bool,
)
from .functions import (
    consecutive_chunks,
    coords_tag_to_int_tuple,
    ev_stack_dict,
    event_dict,
    get_checkbox_points,
    get_n2a,
    is_contiguous,
    num2alpha,
    try_binding,
)
from .other_classes import (
    DraggedRowColumn,
    DrawnItem,
    TextCfg,
    TextEditor,
)
from .vars import (
    USER_OS,
    Color_Map,
    rc_binding,
    symbols_set,
)


class RowIndex(tk.Canvas):
    def __init__(self, *args, **kwargs):
        tk.Canvas.__init__(
            self,
            kwargs["parentframe"],
            background=kwargs["index_bg"],
            highlightthickness=0,
        )
        self.parentframe = kwargs["parentframe"]
        self.MT = None  # is set from within MainTable() __init__
        self.CH = None  # is set from within MainTable() __init__
        self.TL = None  # is set from within TopLeftRectangle() __init__
        self.popup_menu_loc = None
        self.extra_begin_edit_cell_func = None
        self.extra_end_edit_cell_func = None
        self.text_editor = None
        self.text_editor_id = None
        self.text_editor_loc = None
        self.b1_pressed_loc = None
        self.existing_dropdown_canvas_id = None
        self.existing_dropdown_window = None
        self.closed_dropdown = None
        self.centre_alignment_text_mod_indexes = (slice(1, None), slice(None, -1))
        self.c_align_cyc = cycle(self.centre_alignment_text_mod_indexes)
        self.grid_cyctup = ("st", "end")
        self.grid_cyc = cycle(self.grid_cyctup)
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
        self.options = {}
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

        self.disp_text = defaultdict(set)
        self.disp_high = defaultdict(set)
        self.disp_grid = {}
        self.disp_fill_sels = {}
        self.disp_bord_sels = {}
        self.disp_resize_lines = {}
        self.disp_dropdown = {}
        self.disp_checkbox = {}
        self.hidd_text = defaultdict(set)
        self.hidd_high = defaultdict(set)
        self.hidd_grid = {}
        self.hidd_fill_sels = {}
        self.hidd_bord_sels = {}
        self.hidd_resize_lines = {}
        self.hidd_dropdown = {}
        self.hidd_checkbox = {}

        self.row_drag_and_drop_perform = kwargs["row_drag_and_drop_perform"]
        self.index_fg = kwargs["index_fg"]
        self.index_grid_fg = kwargs["index_grid_fg"]
        self.index_border_fg = kwargs["index_border_fg"]
        self.index_selected_cells_bg = kwargs["index_selected_cells_bg"]
        self.index_selected_cells_fg = kwargs["index_selected_cells_fg"]
        self.index_selected_rows_bg = kwargs["index_selected_rows_bg"]
        self.index_selected_rows_fg = kwargs["index_selected_rows_fg"]
        self.index_hidden_rows_expander_bg = kwargs["index_hidden_rows_expander_bg"]
        self.index_bg = kwargs["index_bg"]
        self.drag_and_drop_bg = kwargs["drag_and_drop_bg"]
        self.resizing_line_fg = kwargs["resizing_line_fg"]
        self.align = kwargs["row_index_align"]
        self.show_default_index_for_empty = kwargs["show_default_index_for_empty"]
        self.auto_resize_width = kwargs["auto_resize_width"]
        self.default_index = kwargs["default_row_index"].lower()
        self.basic_bindings()

    def basic_bindings(self, enable=True):
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

    def set_width(self, new_width, set_TL=False):
        self.current_width = new_width
        try:
            self.config(width=new_width)
        except Exception:
            return
        if set_TL:
            self.TL.set_dimensions(new_w=new_width)

    def enable_bindings(self, binding):
        if binding == "row_width_resize":
            self.width_resizing_enabled = True
        elif binding == "row_height_resize":
            self.height_resizing_enabled = True
        elif binding == "double_click_row_resize":
            self.double_click_resizing_enabled = True
        elif binding == "row_select":
            self.row_selection_enabled = True
        elif binding == "drag_and_drop":
            self.drag_and_drop_enabled = True

    def disable_bindings(self, binding):
        if binding == "row_width_resize":
            self.width_resizing_enabled = False
        elif binding == "row_height_resize":
            self.height_resizing_enabled = False
        elif binding == "double_click_row_resize":
            self.double_click_resizing_enabled = False
        elif binding == "row_select":
            self.row_selection_enabled = False
        elif binding == "drag_and_drop":
            self.drag_and_drop_enabled = False

    def check_mouse_position_height_resizers(self, x, y):
        for r, (x1, y1, x2, y2) in self.visible_row_dividers.items():
            if x >= x1 and y >= y1 and x <= x2 and y <= y2:
                return r

    def rc(self, event):
        self.mouseclick_outside_editor_or_dropdown_all_canvases()
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
        if self.extra_rc_func is not None:
            self.extra_rc_func(event)
        if popup_menu is not None:
            self.popup_menu_loc = r
            popup_menu.tk_popup(event.x_root, event.y_root)

    def ctrl_b1_press(self, event=None):
        self.mouseclick_outside_editor_or_dropdown_all_canvases()
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
                    if self.ctrl_selection_binding_func is not None:
                        self.ctrl_selection_binding_func(
                            self.MT.get_select_event(being_drawn_item=self.being_drawn_item),
                        )
                elif r_selected:
                    self.MT.deselect(r=r)
        elif not self.MT.ctrl_select_enabled:
            self.b1_press(event)

    def ctrl_shift_b1_press(self, event):
        self.mouseclick_outside_editor_or_dropdown_all_canvases()
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
                    currently_selected = self.MT.currently_selected()
                    self.MT.delete_item(self.MT.currently_selected().tags[2])
                    if currently_selected and currently_selected.type_ == "row":
                        box = self.get_shift_select_box(r, currently_selected.row)
                        self.being_drawn_item = self.MT.create_selection_box(*box, set_current=currently_selected)
                    else:
                        self.being_drawn_item = self.add_selection(r, run_binding_func=False, set_as_current=True)
                        box = self.MT.get_box_from_item(self.being_drawn_item, get_dict=True)
                    self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
                    if self.ctrl_selection_binding_func is not None:
                        self.ctrl_selection_binding_func(
                            self.MT.get_select_event(being_drawn_item=self.being_drawn_item),
                        )
                elif r_selected:
                    self.dragged_row = DraggedRowColumn(
                        dragged=r,
                        to_move=sorted(self.MT.get_selected_rows()),
                    )
        elif not self.MT.ctrl_select_enabled:
            self.shift_b1_press(event)

    def shift_b1_press(self, event):
        self.mouseclick_outside_editor_or_dropdown_all_canvases()
        y = event.y
        r = self.MT.identify_row(y=y)
        if (self.drag_and_drop_enabled or self.row_selection_enabled) and self.rsz_h is None and self.rsz_w is None:
            if r < len(self.MT.row_positions) - 1:
                r_selected = self.MT.row_selected(r)
                if not r_selected and self.row_selection_enabled:
                    currently_selected = self.MT.currently_selected()
                    if currently_selected and currently_selected.type_ == "row":
                        self.MT.deselect("all", redraw=False)
                        box = self.get_shift_select_box(r, currently_selected.row)
                        self.being_drawn_item = self.MT.create_selection_box(*box, set_current=currently_selected)
                    else:
                        self.being_drawn_item = self.add_selection(r, run_binding_func=False, set_as_current=True)
                        box = self.MT.get_box_from_item(self.being_drawn_item, get_dict=True)
                    self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
                    if self.shift_selection_binding_func is not None:
                        self.shift_selection_binding_func(
                            self.MT.get_select_event(being_drawn_item=self.being_drawn_item),
                        )
                elif r_selected:
                    self.dragged_row = DraggedRowColumn(
                        dragged=r,
                        to_move=sorted(self.MT.get_selected_rows()),
                    )

    def get_shift_selection_box(self, r, min_r):
        if r > min_r:
            return (min_r, 0, r + 1, len(self.MT.col_positions) - 1, "rows")
        elif r < min_r:
            return (r, 0, min_r + 1, len(self.MT.col_positions) - 1, "rows")

    def create_resize_line(self, x1, y1, x2, y2, width, fill, tag):
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

    def delete_resize_lines(self):
        self.hidd_resize_lines.update(self.disp_resize_lines)
        self.disp_resize_lines = {}
        for t, sh in self.hidd_resize_lines.items():
            if sh:
                self.itemconfig(t, state="hidden")
                self.hidd_resize_lines[t] = False

    def mouse_motion(self, event):
        if not self.currently_resizing_height and not self.currently_resizing_width:
            x = self.canvasx(event.x)
            y = self.canvasy(event.y)
            mouse_over_resize = False
            mouse_over_selected = False
            if self.height_resizing_enabled and not mouse_over_resize:
                r = self.check_mouse_position_height_resizers(x, y)
                if r is not None:
                    self.config(cursor="sb_v_double_arrow")
                    self.rsz_h = r
                    mouse_over_resize = True
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
                        self.config(cursor="sb_h_double_arrow")
                        self.rsz_w = True
                        mouse_over_resize = True
                    else:
                        self.rsz_w = None
                except Exception:
                    self.rsz_w = None
            if not mouse_over_resize:
                if self.MT.row_selected(self.MT.identify_row(event, allow_end=False)):
                    self.config(cursor="hand2")
                    mouse_over_selected = True
            if not mouse_over_resize and not mouse_over_selected:
                self.MT.reset_mouse_motion_creations()
        if self.extra_motion_func is not None:
            self.extra_motion_func(event)

    def double_b1(self, event=None):
        self.mouseclick_outside_editor_or_dropdown_all_canvases()
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
                        sheet=self.parentframe.name,
                        resized_rows={row: {"old_size": old_height, "new_size": new_height}},
                    )
                )
        elif self.width_resizing_enabled and self.rsz_h is None and self.rsz_w is True:
            self.set_width_of_index_to_text()
        elif self.row_selection_enabled and self.rsz_h is None and self.rsz_w is None:
            r = self.MT.identify_row(y=event.y)
            if r < len(self.MT.row_positions) - 1:
                if self.MT.single_selection_enabled:
                    self.select_row(r, redraw=True)
                elif self.MT.toggle_selection_enabled:
                    self.toggle_select_row(r, redraw=True)
                datarn = r if self.MT.all_rows_displayed else self.MT.displayed_rows[r]
                if (
                    self.get_cell_kwargs(datarn, key="dropdown")
                    or self.get_cell_kwargs(datarn, key="checkbox")
                    or self.edit_cell_enabled
                ):
                    self.open_cell(event)
        self.rsz_h = None
        self.mouse_motion(event)
        if self.extra_double_b1_func is not None:
            self.extra_double_b1_func(event)

    def b1_press(self, event=None):
        self.MT.unbind("<MouseWheel>")
        self.focus_set()
        self.closed_dropdown = self.mouseclick_outside_editor_or_dropdown_all_canvases()
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
                fill=self.resizing_line_fg,
                tag="rhl",
            )
            self.MT.create_resize_line(x1, y, x2, y, width=1, fill=self.resizing_line_fg, tag="rhl")
            self.create_resize_line(
                0,
                line2y,
                self.current_width,
                line2y,
                width=1,
                fill=self.resizing_line_fg,
                tag="rhl2",
            )
            self.MT.create_resize_line(x1, line2y, x2, line2y, width=1, fill=self.resizing_line_fg, tag="rhl2")
        elif self.width_resizing_enabled and self.rsz_h is None and self.rsz_w is True:
            self.currently_resizing_width = True
            x1, y1, x2, y2 = self.MT.get_canvas_visible_area()
            x = int(event.x)
            if x < self.MT.min_column_width:
                x = int(self.MT.min_column_width)
            self.new_row_width = x
            self.create_resize_line(x, y1, x, y2, width=1, fill=self.resizing_line_fg, tag="rwl")
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
        if self.extra_b1_press_func is not None:
            self.extra_b1_press_func(event)

    def b1_motion(self, event):
        x1, y1, x2, y2 = self.MT.get_canvas_visible_area()
        if self.height_resizing_enabled and self.rsz_h is not None and self.currently_resizing_height:
            y = self.canvasy(event.y)
            size = y - self.MT.row_positions[self.rsz_h - 1]
            if size >= self.MT.min_row_height and size < self.MT.max_row_height:
                self.delete_all_resize_and_ctrl_lines(ctrl_lines=False)
                line2y = self.MT.row_positions[self.rsz_h - 1]
                self.create_resize_line(
                    0,
                    y,
                    self.current_width,
                    y,
                    width=1,
                    fill=self.resizing_line_fg,
                    tag="rhl",
                )
                self.MT.create_resize_line(x1, y, x2, y, width=1, fill=self.resizing_line_fg, tag="rhl")
                self.create_resize_line(
                    0,
                    line2y,
                    self.current_width,
                    line2y,
                    width=1,
                    fill=self.resizing_line_fg,
                    tag="rhl2",
                )
                self.MT.create_resize_line(
                    x1,
                    line2y,
                    x2,
                    line2y,
                    width=1,
                    fill=self.resizing_line_fg,
                    tag="rhl2",
                )
        elif self.width_resizing_enabled and self.rsz_w is not None and self.currently_resizing_width:
            evx = event.x
            self.delete_all_resize_and_ctrl_lines(ctrl_lines=False)
            if evx > self.current_width:
                x = self.MT.canvasx(evx - self.current_width)
                if evx > self.MT.max_index_width:
                    evx = int(self.MT.max_index_width)
                    x = self.MT.canvasx(evx - self.current_width)
                self.new_row_width = evx
                self.MT.create_resize_line(x, y1, x, y2, width=1, fill=self.resizing_line_fg, tag="rwl")
            else:
                x = evx
                if x < self.MT.min_column_width:
                    x = int(self.MT.min_column_width)
                self.new_row_width = x
                self.create_resize_line(x, y1, x, y2, width=1, fill=self.resizing_line_fg, tag="rwl")
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
            currently_selected = self.MT.currently_selected()
            if end_row < len(self.MT.row_positions) - 1 and currently_selected:
                if currently_selected.type_ == "row":
                    box = self.get_b1_motion_box(currently_selected.row, end_row)
                    if (
                        box is not None
                        and self.being_drawn_item is not None
                        and self.MT.get_box_from_item(self.being_drawn_item) != box
                    ):
                        self.MT.deselect("all", redraw=False)
                        if box[2] - box[0] != 1:
                            self.being_drawn_item = self.MT.create_selection_box(*box, set_current=currently_selected)
                        else:
                            self.being_drawn_item = self.select_row(currently_selected.row, run_binding_func=False)
                        need_redraw = True
                        if self.drag_selection_binding_func is not None:
                            self.drag_selection_binding_func(
                                self.MT.get_select_event(being_drawn_item=self.being_drawn_item),
                            )
                if self.scroll_if_event_offscreen(event):
                    need_redraw = True
            if need_redraw:
                self.MT.main_table_redraw_grid_and_text(redraw_header=False, redraw_row_index=True)
        if self.extra_b1_motion_func is not None:
            self.extra_b1_motion_func(event)

    def get_b1_motion_box(self, start_row, end_row):
        if end_row >= start_row:
            return (start_row, 0, end_row + 1, len(self.MT.col_positions) - 1, "rows")
        elif end_row < start_row:
            return (end_row, 0, start_row + 1, len(self.MT.col_positions) - 1, "rows")

    def ctrl_b1_motion(self, event):
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
            currently_selected = self.MT.currently_selected()
            if end_row < len(self.MT.row_positions) - 1 and currently_selected:
                if currently_selected.type_ == "row":
                    box = self.get_b1_motion_box(currently_selected.row, end_row)
                    if (
                        box is not None
                        and self.being_drawn_item is not None
                        and self.MT.get_box_from_item(self.being_drawn_item) != box
                    ):
                        self.MT.delete_item(self.being_drawn_item)
                        if box[2] - box[0] != 1:
                            self.being_drawn_item = self.MT.create_selection_box(*box, set_current=currently_selected)
                        else:
                            self.being_drawn_item = self.add_selection(currently_selected.row, run_binding_func=False)
                        need_redraw = True
                        if self.drag_selection_binding_func is not None:
                            self.drag_selection_binding_func(
                                self.MT.get_select_event(being_drawn_item=self.being_drawn_item),
                            )
                if self.scroll_if_event_offscreen(event):
                    need_redraw = True
            if need_redraw:
                self.MT.main_table_redraw_grid_and_text(redraw_header=False, redraw_row_index=True)
        elif not self.MT.ctrl_select_enabled:
            self.b1_motion(event)

    def drag_and_drop_motion(self, event):
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
            self.MT.main_table_redraw_grid_and_text(redraw_row_index=True)
        elif y <= 0 and len(ycheck) > 1 and ycheck[0] > 0:
            if y >= -15:
                self.MT.yview_scroll(-1, "units")
                self.yview_scroll(-1, "units")
            else:
                self.MT.yview_scroll(-2, "units")
                self.yview_scroll(-2, "units")
            self.fix_yview()
            self.MT.main_table_redraw_grid_and_text(redraw_row_index=True)
        return self.MT.row_positions[self.MT.identify_row(y=y)]

    def show_drag_and_drop_indicators(
        self,
        ypos,
        x1,
        x2,
        rows,
    ):
        self.delete_all_resize_and_ctrl_lines()
        self.create_resize_line(
            0,
            ypos,
            self.current_width,
            ypos,
            width=3,
            fill=self.drag_and_drop_bg,
            tag="dd",
        )
        self.MT.create_resize_line(x1, ypos, x2, ypos, width=3, fill=self.drag_and_drop_bg, tag="dd")
        for chunk in consecutive_chunks(rows):
            self.MT.show_ctrl_outline(
                start_cell=(0, chunk[0]),
                end_cell=(len(self.MT.col_positions) - 1, chunk[-1] + 1),
                dash=(),
                outline=self.drag_and_drop_bg,
                delete_on_timer=False,
            )

    def delete_all_resize_and_ctrl_lines(self, ctrl_lines=True):
        self.delete_resize_lines()
        self.MT.delete_resize_lines()
        if ctrl_lines:
            self.MT.delete_ctrl_outlines()

    def scroll_if_event_offscreen(self, event):
        ycheck = self.yview()
        need_redraw = False
        if event.y > self.winfo_height() and len(ycheck) > 1 and ycheck[1] < 1:
            try:
                self.MT.yview_scroll(1, "units")
                self.yview_scroll(1, "units")
            except Exception:
                pass
            self.fix_yview()
            need_redraw = True
        elif event.y < 0 and self.canvasy(self.winfo_height()) > 0 and ycheck and ycheck[0] > 0:
            try:
                self.yview_scroll(-1, "units")
                self.MT.yview_scroll(-1, "units")
            except Exception:
                pass
            self.fix_yview()
            need_redraw = True
        return need_redraw

    def fix_yview(self):
        ycheck = self.yview()
        if ycheck and ycheck[0] < 0:
            self.MT.set_yviews("moveto", 0)
        if len(ycheck) > 1 and ycheck[1] > 1:
            self.MT.set_yviews("moveto", 1)

    def event_over_dropdown(self, r, datarn, event, canvasy):
        if (
            canvasy < self.MT.row_positions[r] + self.MT.index_txt_height
            and self.get_cell_kwargs(datarn, key="dropdown")
            and event.x > self.current_width - self.MT.index_txt_height - 4
        ):
            return True
        return False

    def event_over_checkbox(self, r, datarn, event, canvasy):
        if (
            canvasy < self.MT.row_positions[r] + self.MT.index_txt_height
            and self.get_cell_kwargs(datarn, key="checkbox")
            and event.x < self.MT.index_txt_height + 4
        ):
            return True
        return False

    def b1_release(self, event=None):
        if self.being_drawn_item is not None:
            currently_selected = self.MT.currently_selected()
            to_sel = self.MT.get_box_from_item(self.being_drawn_item)
            self.MT.delete_item(self.being_drawn_item)
            self.being_drawn_item = None
            self.MT.create_selection_box(*to_sel, set_current=currently_selected)
            if self.drag_selection_binding_func is not None:
                self.drag_selection_binding_func(
                    self.MT.get_select_event(being_drawn_item=self.being_drawn_item),
                )
        self.MT.bind("<MouseWheel>", self.MT.mousewheel)
        if self.height_resizing_enabled and self.rsz_h is not None and self.currently_resizing_height:
            self.currently_resizing_height = False
            new_row_pos = int(self.coords("rhl")[1])
            self.delete_all_resize_and_ctrl_lines(ctrl_lines=False)
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
                        sheet=self.parentframe.name,
                        resized_rows={self.rsz_h - 1: {"old_size": old_height, "new_size": new_height}},
                    )
                )
        elif self.width_resizing_enabled and self.rsz_w is not None and self.currently_resizing_width:
            self.currently_resizing_width = False
            self.delete_all_resize_and_ctrl_lines(ctrl_lines=False)
            self.set_width(self.new_row_width, set_TL=True)
            self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
        if (
            self.drag_and_drop_enabled
            and self.MT.anything_selected(exclude_cells=True, exclude_columns=True)
            and self.row_selection_enabled
            and self.rsz_h is None
            and self.rsz_w is None
            and self.dragged_row is not None
        ):
            self.delete_all_resize_and_ctrl_lines()
            y = event.y
            r = self.MT.identify_row(y=y)
            totalrows = len(self.dragged_row.to_move)
            if (
                r is not None
                and totalrows != (len(self.MT.row_positions) - 1)
                and not (r == self.dragged_row.to_move[0] and is_contiguous(self.dragged_row.to_move))
            ):
                if r >= len(self.MT.row_positions) - 1:
                    r -= 1
                event_data = event_dict(
                    name="move_rows",
                    sheet=self.parentframe.name,
                    boxes=self.MT.get_boxes(),
                    selected=self.MT.currently_selected(),
                    value=r,
                )
                if try_binding(self.ri_extra_begin_drag_drop_func, event_data, "begin_move_rows"):
                    data_new_idxs, disp_new_idxs, event_data = self.MT.move_rows_adjust_options_dict(
                        *self.MT.get_args_for_move_rows(
                            move_to=r,
                            to_move=self.dragged_row.to_move,
                        ),
                        move_data=self.row_drag_and_drop_perform,
                        event_data=event_data,
                    )
                    event_data["moved"]["rows"] = {
                        "data": data_new_idxs,
                        "displayed": disp_new_idxs,
                    }
                    if self.MT.undo_enabled:
                        self.MT.undo_stack.append(ev_stack_dict(event_data))
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
            else:
                self.mouseclick_outside_editor_or_dropdown_all_canvases()
            self.b1_pressed_loc = None
            self.closed_dropdown = None
        self.dragged_row = None
        self.currently_resizing_width = False
        self.currently_resizing_height = False
        self.rsz_w = None
        self.rsz_h = None
        self.mouse_motion(event)
        if self.extra_b1_release_func is not None:
            self.extra_b1_release_func(event)

    def toggle_select_row(
        self,
        row,
        add_selection=True,
        redraw=True,
        run_binding_func=True,
        set_as_current=True,
    ):
        if add_selection:
            if self.MT.row_selected(row):
                fill_iid = self.MT.deselect(r=row, redraw=redraw)
            else:
                fill_iid = self.add_selection(
                    r=row,
                    redraw=redraw,
                    run_binding_func=run_binding_func,
                    set_as_current=set_as_current,
                )
        else:
            if self.MT.row_selected(row):
                fill_iid = self.MT.deselect(r=row, redraw=redraw)
            else:
                fill_iid = self.select_row(row, redraw=redraw)
        return fill_iid

    def select_row(self, r, redraw=False, run_binding_func=True):
        self.MT.deselect("all", redraw=False)
        box = (r, 0, r + 1, len(self.MT.col_positions) - 1, "rows")
        fill_iid = self.MT.create_selection_box(*box)
        if redraw:
            self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
        if run_binding_func:
            self.MT.run_selection_binding("rows")
        return fill_iid

    def add_selection(self, r, redraw=False, run_binding_func=True, set_as_current=True):
        box = (r, 0, r + 1, len(self.MT.col_positions) - 1, "rows")
        fill_iid = self.MT.create_selection_box(*box, set_current=set_as_current)
        if redraw:
            self.MT.main_table_redraw_grid_and_text(redraw_header=False, redraw_row_index=True)
        if run_binding_func:
            self.MT.run_selection_binding("rows")
        return fill_iid

    def get_cell_dimensions(self, datarn):
        txt = self.get_valid_cell_data_as_str(datarn, fix=False)
        if txt:
            self.MT.txt_measure_canvas.itemconfig(self.MT.txt_measure_canvas_text, text=txt, font=self.MT.index_font)
            b = self.MT.txt_measure_canvas.bbox(self.MT.txt_measure_canvas_text)
            w = b[2] - b[0] + 7
            h = b[3] - b[1] + 5
        else:
            w = self.MT.default_index_width
            h = self.MT.min_row_height
        if self.get_cell_kwargs(datarn, key="dropdown") or self.get_cell_kwargs(datarn, key="checkbox"):
            return w + self.MT.index_txt_height, h
        return w, h

    def set_row_height(
        self,
        row,
        height=None,
        only_set_if_too_small=False,
        recreate=True,
        return_new_height=False,
        displayed_only=False,
    ):
        r_norm = row + 1
        r_extra = row + 2
        min_rh = self.MT.min_row_height
        datarn = row if self.MT.all_rows_displayed else self.MT.displayed_rows[row]
        if height is None:
            if self.MT.all_columns_displayed:
                if displayed_only:
                    x1, y1, x2, y2 = self.MT.get_canvas_visible_area()
                    start_col, end_col = self.MT.get_visible_columns(x1, x2)
                else:
                    start_col, end_col = (
                        0,
                        len(self.MT.data[row]) if self.MT.data else 0,
                    )
                iterable = range(start_col, end_col)
            else:
                if displayed_only:
                    x1, y1, x2, y2 = self.MT.get_canvas_visible_area()
                    start_col, end_col = self.MT.get_visible_columns(x1, x2)
                else:
                    start_col, end_col = 0, len(self.MT.displayed_columns)
                iterable = self.MT.displayed_columns[start_col:end_col]
            new_height = int(min_rh)
            w_, h = self.get_cell_dimensions(datarn)
            if h < min_rh:
                h = int(min_rh)
            elif h > self.MT.max_row_height:
                h = int(self.MT.max_row_height)
            if h > new_height:
                new_height = h
            if self.MT.data:
                for datacn in iterable:
                    txt = self.MT.get_valid_cell_data_as_str(datarn, datacn, get_displayed=True)
                    if txt:
                        h = self.MT.get_txt_h(txt) + 5
                    else:
                        h = min_rh
                    if h < min_rh:
                        h = int(min_rh)
                    elif h > self.MT.max_row_height:
                        h = int(self.MT.max_row_height)
                    if h > new_height:
                        new_height = h
        else:
            new_height = int(height)
        if new_height < min_rh:
            new_height = int(min_rh)
        elif new_height > self.MT.max_row_height:
            new_height = int(self.MT.max_row_height)
        if only_set_if_too_small and new_height <= self.MT.row_positions[row + 1] - self.MT.row_positions[row]:
            return self.MT.row_positions[row + 1] - self.MT.row_positions[row]
        if not return_new_height:
            new_row_pos = self.MT.row_positions[row] + new_height
            increment = new_row_pos - self.MT.row_positions[r_norm]
            self.MT.row_positions[r_extra:] = [
                e + increment for e in islice(self.MT.row_positions, r_extra, len(self.MT.row_positions))
            ]
            self.MT.row_positions[r_norm] = new_row_pos
            if recreate:
                self.MT.recreate_all_selection_boxes()
        return new_height

    def set_width_of_index_to_text(self, text=None):
        if (
            text is None
            and not self.MT._row_index
            and isinstance(self.MT._row_index, list)
            or isinstance(self.MT._row_index, int)
            and self.MT._row_index >= len(self.MT.data)
        ):
            return
        qconf = self.MT.txt_measure_canvas.itemconfig
        qbbox = self.MT.txt_measure_canvas.bbox
        qtxtm = self.MT.txt_measure_canvas_text
        new_width = int(self.MT.min_column_width)
        self.fix_index()
        if text is not None:
            if text:
                qconf(qtxtm, text=text)
                b = qbbox(qtxtm)
                w = b[2] - b[0] + 10
                if w > new_width:
                    new_width = w
            else:
                w = self.MT.default_index_width
        else:
            if self.MT.all_rows_displayed:
                if isinstance(self.MT._row_index, list):
                    iterable = range(len(self.MT._row_index))
                else:
                    iterable = range(len(self.MT.data))
            else:
                iterable = self.MT.displayed_rows
            if isinstance(self.MT._row_index, list):
                for datarn in iterable:
                    w, h_ = self.get_cell_dimensions(datarn)
                    if w < self.MT.min_column_width:
                        w = int(self.MT.min_column_width)
                    elif w > self.MT.max_index_width:
                        w = int(self.MT.max_index_width)
                    if self.get_cell_kwargs(datarn, key="checkbox"):
                        w += self.MT.index_txt_height + 6
                    elif self.get_cell_kwargs(datarn, key="dropdown"):
                        w += self.MT.index_txt_height + 4
                    if w > new_width:
                        new_width = w
            elif isinstance(self.MT._row_index, int):
                datacn = self.MT._row_index
                for datarn in iterable:
                    txt = self.MT.get_valid_cell_data_as_str(datarn, datacn, get_displayed=True)
                    if txt:
                        qconf(qtxtm, text=txt)
                        b = qbbox(qtxtm)
                        w = b[2] - b[0] + 10
                    else:
                        w = self.MT.default_index_width
                    if w < self.MT.min_column_width:
                        w = int(self.MT.min_column_width)
                    elif w > self.MT.max_index_width:
                        w = int(self.MT.max_index_width)
                    if w > new_width:
                        new_width = w
        if new_width == self.MT.min_column_width:
            new_width = self.MT.min_column_width + 10
        self.set_width(new_width, set_TL=True)
        self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)

    def set_height_of_all_rows(self, height=None, only_set_if_too_small=False, recreate=True):
        if height is None:
            self.MT.set_row_positions(
                itr=(
                    self.set_row_height(
                        rn,
                        only_set_if_too_small=only_set_if_too_small,
                        recreate=False,
                        return_new_height=True,
                    )
                    for rn in range(len(self.MT.data))
                )
            )
        else:
            self.MT.set_row_positions(itr=(height for r in range(len(self.MT.data))))
        if recreate:
            self.MT.recreate_all_selection_boxes()

    def auto_set_index_width(self, end_row):
        if not self.MT._row_index and not isinstance(self.MT._row_index, int) and self.auto_resize_width:
            if self.default_index == "letters":
                new_w = self.MT.get_txt_w(f"{num2alpha(end_row)}") + 20
                if self.current_width - new_w > 15 or new_w - self.current_width > 5:
                    self.set_width(new_w, set_TL=True)
                    return True
            elif self.default_index == "numbers":
                new_w = self.MT.get_txt_w(f"{end_row}") + 20
                if self.current_width - new_w > 15 or new_w - self.current_width > 5:
                    self.set_width(new_w, set_TL=True)
                    return True
            elif self.default_index == "both":
                new_w = self.MT.get_txt_w(f"{end_row + 1} {num2alpha(end_row)}") + 20
                if self.current_width - new_w > 15 or new_w - self.current_width > 5:
                    self.set_width(new_w, set_TL=True)
                    return True
        return False

    def redraw_highlight_get_text_fg(self, fr, sr, r, c_2, c_3, selections, datarn):
        redrawn = False
        kwargs = self.get_cell_kwargs(datarn, key="highlight")
        if kwargs:
            if kwargs[0] is not None:
                c_1 = kwargs[0] if kwargs[0].startswith("#") else Color_Map[kwargs[0]]
            if "rows" in selections and r in selections["rows"]:
                tf = (
                    self.index_selected_rows_fg
                    if kwargs[1] is None or self.MT.display_selected_fg_over_highlights
                    else kwargs[1]
                )
                if kwargs[0] is not None:
                    fill = (
                        f"#{int((int(c_1[1:3], 16) + int(c_3[1:3], 16)) / 2):02X}"
                        + f"{int((int(c_1[3:5], 16) + int(c_3[3:5], 16)) / 2):02X}"
                        + f"{int((int(c_1[5:], 16) + int(c_3[5:], 16)) / 2):02X}"
                    )
            elif "cells" in selections and r in selections["cells"]:
                tf = (
                    self.index_selected_cells_fg
                    if kwargs[1] is None or self.MT.display_selected_fg_over_highlights
                    else kwargs[1]
                )
                if kwargs[0] is not None:
                    fill = (
                        f"#{int((int(c_1[1:3], 16) + int(c_2[1:3], 16)) / 2):02X}"
                        + f"{int((int(c_1[3:5], 16) + int(c_2[3:5], 16)) / 2):02X}"
                        + f"{int((int(c_1[5:], 16) + int(c_2[5:], 16)) / 2):02X}"
                    )
            else:
                tf = self.index_fg if kwargs[1] is None else kwargs[1]
                if kwargs[0] is not None:
                    fill = kwargs[0]
            if kwargs[0] is not None:
                redrawn = self.redraw_highlight(
                    0,
                    fr + 1,
                    self.current_width - 1,
                    sr,
                    fill=fill,
                    outline=self.index_fg
                    if self.get_cell_kwargs(datarn, key="dropdown") and self.MT.show_dropdown_borders
                    else "",
                    tag="s",
                )
        elif not kwargs:
            if "rows" in selections and r in selections["rows"]:
                tf = self.index_selected_rows_fg
            elif "cells" in selections and r in selections["cells"]:
                tf = self.index_selected_cells_fg
            else:
                tf = self.index_fg
        return tf, redrawn

    def redraw_highlight(self, x1, y1, x2, y2, fill, outline, tag):
        config = (fill, outline)
        coords = (x1, y1, x2, y2)
        k = None
        if config in self.hidd_high:
            k = config
            iid, showing = self.hidd_high[k].pop()
            if all(int(crd1) == int(crd2) for crd1, crd2 in zip(self.coords(iid), coords)):
                option = 0 if showing else 2
            else:
                option = 1 if showing else 3

        elif self.hidd_high:
            k = next(iter(self.hidd_high))
            iid, showing = self.hidd_high[k].pop()
            if all(int(crd1) == int(crd2) for crd1, crd2 in zip(self.coords(iid), coords)):
                option = 2 if showing else 3
            else:
                option = 3

        else:
            iid, showing, option = (
                self.create_rectangle(coords, fill=fill, outline=outline, tag=tag),
                1,
                4,
            )

        if option in (1, 3):
            self.coords(iid, coords)
        if option in (2, 3):
            if showing:
                self.itemconfig(iid, fill=fill, outline=outline)
            else:
                self.itemconfig(iid, fill=fill, outline=outline, tag=tag, state="normal")

        if k is not None and not self.hidd_high[k]:
            del self.hidd_high[k]

        self.disp_high[config].add(DrawnItem(iid=iid, showing=1))
        return True

    def redraw_gridline(self, points, fill, width, tag):
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
        x1,
        y1,
        x2,
        y2,
        fill,
        outline,
        tag,
        draw_outline=True,
        draw_arrow=True,
        dd_is_open=False,
    ):
        if draw_outline and self.MT.show_dropdown_borders:
            self.redraw_highlight(x1 + 1, y1 + 1, x2, y2, fill="", outline=self.index_fg, tag=tag)
        if draw_arrow:
            topysub = floor(self.MT.index_half_txt_height / 2)
            mid_y = y1 + floor(self.MT.min_row_height / 2)
            if mid_y + topysub + 1 >= y1 + self.MT.index_txt_height - 1:
                mid_y -= 1
            if mid_y - topysub + 2 <= y1 + 4 + topysub:
                mid_y -= 1
                ty1 = mid_y + topysub + 1 if dd_is_open else mid_y - topysub + 3
                ty2 = mid_y - topysub + 3 if dd_is_open else mid_y + topysub + 1
                ty3 = mid_y + topysub + 1 if dd_is_open else mid_y - topysub + 3
            else:
                ty1 = mid_y + topysub + 1 if dd_is_open else mid_y - topysub + 2
                ty2 = mid_y - topysub + 2 if dd_is_open else mid_y + topysub + 1
                ty3 = mid_y + topysub + 1 if dd_is_open else mid_y - topysub + 2
            tx1 = x2 - self.MT.index_txt_height + 1
            tx2 = x2 - self.MT.index_half_txt_height - 1
            tx3 = x2 - 3
            if tx2 - tx1 > tx3 - tx2:
                tx1 += (tx2 - tx1) - (tx3 - tx2)
            elif tx2 - tx1 < tx3 - tx2:
                tx1 -= (tx3 - tx2) - (tx2 - tx1)
            points = (tx1, ty1, tx2, ty2, tx3, ty3)
            if self.hidd_dropdown:
                t, sh = self.hidd_dropdown.popitem()
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
                    width=2,
                    capstyle=tk.ROUND,
                    joinstyle=tk.ROUND,
                    tag=tag,
                )
            self.disp_dropdown[t] = True

    def redraw_checkbox(self, x1, y1, x2, y2, fill, outline, tag, draw_check=False):
        points = get_checkbox_points(x1, y1, x2, y2)
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
            points = get_checkbox_points(x1, y1, x2, y2, radius=4)
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

    def redraw_grid_and_text(
        self,
        last_row_line_pos,
        scrollpos_top,
        y_stop,
        start_row,
        end_row,
        scrollpos_bot,
        row_pos_exists,
    ):
        try:
            self.configure(
                scrollregion=(
                    0,
                    0,
                    self.current_width,
                    last_row_line_pos + self.MT.empty_vertical + 2,
                )
            )
        except Exception:
            return
        for k, v in self.disp_text.items():
            if k in self.hidd_text:
                self.hidd_text[k] = self.hidd_text[k] | self.disp_text[k]
            else:
                self.hidd_text[k] = v
        self.disp_text = defaultdict(set)
        for k, v in self.disp_high.items():
            if k in self.hidd_high:
                self.hidd_high[k] = self.hidd_high[k] | self.disp_high[k]
            else:
                self.hidd_high[k] = v
        self.disp_high = defaultdict(set)
        self.hidd_grid.update(self.disp_grid)
        self.disp_grid = {}
        self.hidd_dropdown.update(self.disp_dropdown)
        self.disp_dropdown = {}
        self.hidd_checkbox.update(self.disp_checkbox)
        self.disp_checkbox = {}
        self.visible_row_dividers = {}
        draw_y = self.MT.row_positions[start_row]
        xend = self.current_width - 6
        self.row_width_resize_bbox = (
            self.current_width - 2,
            scrollpos_top,
            self.current_width,
            scrollpos_bot,
        )
        if (self.MT.show_horizontal_grid or self.height_resizing_enabled) and row_pos_exists:
            self.grid_cyc = cycle(self.grid_cyctup)
            points = [
                self.current_width - 1,
                y_stop - 1,
                self.current_width - 1,
                scrollpos_top - 1,
                -1,
                scrollpos_top - 1,
            ]
            for r in range(start_row + 1, end_row):
                draw_y = self.MT.row_positions[r]
                if self.height_resizing_enabled:
                    self.visible_row_dividers[r] = (1, draw_y - 2, xend, draw_y + 2)
                st_or_end = next(self.grid_cyc)
                if st_or_end == "st":
                    points.extend(
                        [
                            -1,
                            draw_y,
                            self.current_width,
                            draw_y,
                            self.current_width,
                            self.MT.row_positions[r + 1] if len(self.MT.row_positions) - 1 > r else draw_y,
                        ]
                    )
                elif st_or_end == "end":
                    points.extend(
                        [
                            self.current_width,
                            draw_y,
                            -1,
                            draw_y,
                            -1,
                            self.MT.row_positions[r + 1] if len(self.MT.row_positions) - 1 > r else draw_y,
                        ]
                    )
                if points:
                    self.redraw_gridline(points=points, fill=self.index_grid_fg, width=1, tag="h")
        c_2 = (
            self.index_selected_cells_bg
            if self.index_selected_cells_bg.startswith("#")
            else Color_Map[self.index_selected_cells_bg]
        )
        c_3 = (
            self.index_selected_rows_bg
            if self.index_selected_rows_bg.startswith("#")
            else Color_Map[self.index_selected_rows_bg]
        )
        font = self.MT.index_font
        selections = self.get_redraw_selections(start_row, end_row)
        for r in range(start_row, end_row - 1):
            rtopgridln = self.MT.row_positions[r]
            rbotgridln = self.MT.row_positions[r + 1]
            if rbotgridln - rtopgridln < self.MT.index_txt_height:
                continue
            datarn = r if self.MT.all_rows_displayed else self.MT.displayed_rows[r]
            fill, dd_drawn = self.redraw_highlight_get_text_fg(rtopgridln, rbotgridln, r, c_2, c_3, selections, datarn)

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
                        dd_is_open=dropdown_kwargs["window"] != "no dropdown open",
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
                        dd_is_open=dropdown_kwargs["window"] != "no dropdown open",
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
                        dd_is_open=dropdown_kwargs["window"] != "no dropdown open",
                    )
                else:
                    mw = self.current_width - 1
                    draw_x = floor(self.current_width / 2)
            checkbox_kwargs = self.get_cell_kwargs(datarn, key="checkbox")
            if checkbox_kwargs and mw > self.MT.index_txt_height + 1:
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
                    fill=fill if checkbox_kwargs["state"] == "normal" else self.index_grid_fg,
                    outline="",
                    tag="cb",
                    draw_check=draw_check,
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
                        config = TextCfg(txt, fill, font, align)
                        k = None
                        if config in self.hidd_text:
                            k = config
                            iid, showing = self.hidd_text[k].pop()
                            cc1, cc2 = self.coords(iid)
                            if int(cc1) == int(draw_x) and int(cc2) == int(draw_y):
                                option = 0 if showing else 2
                            else:
                                option = 1 if showing else 3
                            self.tag_raise(iid)
                        elif self.hidd_text:
                            k = next(iter(self.hidd_text))
                            iid, showing = self.hidd_text[k].pop()
                            cc1, cc2 = self.coords(iid)
                            if int(cc1) == int(draw_x) and int(cc2) == int(draw_y):
                                option = 2 if showing else 3
                            else:
                                option = 3
                            self.tag_raise(iid)
                        else:
                            iid, showing, option = (
                                self.create_text(
                                    draw_x,
                                    draw_y,
                                    text=txt,
                                    fill=fill,
                                    font=font,
                                    anchor=align,
                                    tag="t",
                                ),
                                1,
                                4,
                            )
                        if option in (1, 3):
                            self.coords(iid, draw_x, draw_y)
                        if option in (2, 3):
                            if showing:
                                self.itemconfig(iid, text=txt, fill=fill, font=font, anchor=align)
                            else:
                                self.itemconfig(
                                    iid,
                                    text=txt,
                                    fill=fill,
                                    font=font,
                                    anchor=align,
                                    state="normal",
                                )
                        if k is not None and not self.hidd_text[k]:
                            del self.hidd_text[k]
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
                            elif align == "e" and dropdown_kwargs:
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
                            self.disp_text[config._replace(txt=txt)].add(DrawnItem(iid=iid, showing=True))
                        else:
                            self.disp_text[config].add(DrawnItem(iid=iid, showing=True))
                        draw_y += self.MT.index_xtra_lines_increment
                        if draw_y + self.MT.index_half_txt_height - 1 > rbotgridln:
                            break
        for cfg, set_ in self.hidd_text.items():
            for namedtup in tuple(set_):
                if namedtup.showing:
                    self.itemconfig(namedtup.iid, state="hidden")
                    self.hidd_text[cfg].discard(namedtup)
                    self.hidd_text[cfg].add(namedtup._replace(showing=False))
        for cfg, set_ in self.hidd_high.items():
            for namedtup in tuple(set_):
                if namedtup.showing:
                    self.itemconfig(namedtup.iid, state="hidden")
                    self.hidd_high[cfg].discard(namedtup)
                    self.hidd_high[cfg].add(namedtup._replace(showing=False))
        for t, sh in self.hidd_grid.items():
            if sh:
                self.itemconfig(t, state="hidden")
                self.hidd_grid[t] = False
        for t, sh in self.hidd_dropdown.items():
            if sh:
                self.itemconfig(t, state="hidden")
                self.hidd_dropdown[t] = False
        for t, sh in self.hidd_checkbox.items():
            if sh:
                self.itemconfig(t, state="hidden")
                self.hidd_checkbox[t] = False
        return True

    def get_redraw_selections(self, startr, endr):
        d = defaultdict(list)
        for item in chain(self.find_withtag("cells"), self.find_withtag("rows")):
            tags = self.gettags(item)
            if tags:
                d[tags[0]].append(coords_tag_to_int_tuple(tags[1]))
        d2 = {}
        if "cells" in d:
            d2["cells"] = {r for r in range(startr, endr) for r1, c1, r2, c2 in d["cells"] if r1 <= r and r2 > r}
        if "rows" in d:
            d2["rows"] = {r for r in range(startr, endr) for r1, c1, r2, c2 in d["rows"] if r1 <= r and r2 > r}
        return d2

    def open_cell(self, event=None, ignore_existing_editor=False):
        if not self.MT.anything_selected() or (not ignore_existing_editor and self.text_editor_id is not None):
            return
        currently_selected = self.MT.currently_selected()
        if not currently_selected:
            return
        r = int(currently_selected[0])
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
    def get_cell_align(self, r):
        datarn = r if self.MT.all_rows_displayed else self.MT.displayed_rows[r]
        if datarn in self.cell_options and "align" in self.cell_options[datarn]:
            align = self.cell_options[datarn]["align"]
        else:
            align = self.align
        return align

    # r is displayed row
    def open_text_editor(
        self,
        event=None,
        r=0,
        text=None,
        state="normal",
        see=True,
        set_data_on_close=True,
        binding=None,
        dropdown=False,
    ):
        text = None
        extra_func_key = "??"
        if event is None or self.MT.event_opens_dropdown_or_checkbox(event):
            if event is not None:
                if hasattr(event, "keysym") and event.keysym == "Return":
                    extra_func_key = "Return"
                elif hasattr(event, "keysym") and event.keysym == "F2":
                    extra_func_key = "F2"
            datarn = r if self.MT.all_rows_displayed else self.MT.displayed_rows[r]
            text = self.get_cell_data(datarn, none_to_empty_str=True, redirect_int=True)
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
        self.text_editor_loc = r
        if self.extra_begin_edit_cell_func is not None:
            try:
                text = self.extra_begin_edit_cell_func(
                    event_dict(
                        name="begin_edit_index",
                        sheet=self.parentframe.name,
                        key=extra_func_key,
                        value=text,
                        location=r,
                        boxes=self.MT.get_boxes(),
                        selected=self.MT.currently_selected(),
                    )
                )
            except Exception:
                return False
            if text is None:
                return False
            else:
                text = text if isinstance(text, str) else f"{text}"
        text = "" if text is None else text
        if self.MT.cell_auto_resize_enabled:
            self.set_row_height_run_binding(r)
        if r == self.text_editor_loc and self.text_editor is not None:
            self.text_editor.set_text(self.text_editor.get() + "" if not isinstance(text, str) else text)
            return
        if self.text_editor is not None:
            self.destroy_text_editor()
        if see:
            has_redrawn = self.MT.see(r=r, c=0, keep_yscroll=True, check_cell_visibility=True)
            if not has_redrawn:
                self.MT.refresh()
        self.text_editor_loc = r
        x = 0
        y = self.MT.row_positions[r] + 1
        w = self.current_width + 1
        h = self.MT.row_positions[r + 1] - y
        datarn = r if self.MT.all_rows_displayed else self.MT.displayed_rows[r]
        if text is None:
            text = self.get_cell_data(datarn, none_to_empty_str=True, redirect_int=True)
        bg, fg = self.index_bg, self.index_fg
        self.text_editor = TextEditor(
            self,
            text=text,
            font=self.MT.index_font,
            state=state,
            width=w,
            height=h,
            border_color=self.MT.table_selected_cells_border_fg,
            show_border=False,
            bg=bg,
            fg=fg,
            popup_menu_font=self.MT.popup_menu_font,
            popup_menu_fg=self.MT.popup_menu_fg,
            popup_menu_bg=self.MT.popup_menu_bg,
            popup_menu_highlight_bg=self.MT.popup_menu_highlight_bg,
            popup_menu_highlight_fg=self.MT.popup_menu_highlight_fg,
            binding=binding,
            align=self.get_cell_align(r),
            r=r,
            newline_binding=self.text_editor_newline_binding,
        )
        self.text_editor.update_idletasks()
        self.text_editor_id = self.create_window((x, y), window=self.text_editor, anchor="nw")
        if not dropdown:
            self.text_editor.textedit.focus_set()
            self.text_editor.scroll_to_bottom()
        self.text_editor.textedit.bind("<Alt-Return>", lambda x: self.text_editor_newline_binding(r=r))
        if USER_OS == "darwin":
            self.text_editor.textedit.bind("<Option-Return>", lambda x: self.text_editor_newline_binding(r=r))
        for key, func in self.MT.text_editor_user_bound_keys.items():
            self.text_editor.textedit.bind(key, func)
        if binding is not None:
            self.text_editor.textedit.bind("<Tab>", lambda x: binding((r, "Tab")))
            self.text_editor.textedit.bind("<Return>", lambda x: binding((r, "Return")))
            self.text_editor.textedit.bind("<FocusOut>", lambda x: binding((r, "FocusOut")))
            self.text_editor.textedit.bind("<Escape>", lambda x: binding((r, "Escape")))
        elif binding is None and set_data_on_close:
            self.text_editor.textedit.bind("<Tab>", lambda x: self.close_text_editor((r, "Tab")))
            self.text_editor.textedit.bind("<Return>", lambda x: self.close_text_editor((r, "Return")))
            if not dropdown:
                self.text_editor.textedit.bind("<FocusOut>", lambda x: self.close_text_editor((r, "FocusOut")))
            self.text_editor.textedit.bind("<Escape>", lambda x: self.close_text_editor((r, "Escape")))
        else:
            self.text_editor.textedit.bind("<Escape>", lambda x: self.destroy_text_editor("Escape"))
        return True

    def text_editor_newline_binding(self, r=0, c=0, event=None, check_lines=True):
        if self.height_resizing_enabled:
            datarn = r if self.MT.all_rows_displayed else self.MT.displayed_rows[r]
            curr_height = self.text_editor.winfo_height()
            if (
                not check_lines
                or self.MT.get_lines_cell_height(self.text_editor.get_num_lines() + 1, font=self.MT.index_font)
                > curr_height
            ):
                new_height = curr_height + self.MT.index_xtra_lines_increment
                space_bot = self.MT.get_space_bot(r)
                if new_height > space_bot:
                    new_height = space_bot
                if new_height != curr_height:
                    self.set_row_height(datarn, new_height)
                    self.MT.refresh()
                    self.text_editor.config(height=new_height)
                    self.coords(self.text_editor_id, 0, self.MT.row_positions[r] + 1)
                    kwargs = self.get_cell_kwargs(datarn, key="dropdown")
                    if kwargs:
                        text_editor_h = self.text_editor.winfo_height()
                        win_h, anchor = self.get_dropdown_height_anchor(datarn, text_editor_h)
                        if anchor == "nw":
                            self.coords(
                                kwargs["canvas_id"],
                                self.MT.col_positions[c],
                                self.MT.row_positions[r] + text_editor_h - 1,
                            )
                            self.itemconfig(kwargs["canvas_id"], anchor=anchor, height=win_h)
                        elif anchor == "sw":
                            self.coords(
                                kwargs["canvas_id"],
                                self.MT.col_positions[c],
                                self.MT.row_positions[r],
                            )
                            self.itemconfig(kwargs["canvas_id"], anchor=anchor, height=win_h)

    def bind_cell_edit(self, enable=True):
        if enable:
            self.edit_cell_enabled = True
        else:
            self.edit_cell_enabled = False

    def bind_text_editor_destroy(self, binding, r):
        self.text_editor.textedit.bind("<Return>", lambda x: binding((r, "Return")))
        self.text_editor.textedit.bind("<FocusOut>", lambda x: binding((r, "FocusOut")))
        self.text_editor.textedit.bind("<Escape>", lambda x: binding((r, "Escape")))
        self.text_editor.textedit.focus_set()

    def destroy_text_editor(self, event=None):
        self.text_editor_loc = None
        try:
            self.delete(self.text_editor_id)
        except Exception:
            pass
        try:
            self.text_editor.destroy()
        except Exception:
            pass
        self.text_editor = None
        self.text_editor_id = None
        if event is not None and len(event) >= 3 and "Escape" in event:
            self.focus_set()

    # r is displayed row
    def close_text_editor(
        self,
        editor_info=None,
        r=None,
        set_data_on_close=True,
        event=None,
        destroy=True,
        move_down=True,
        redraw=True,
        recreate=True,
    ):
        if self.focus_get() is None and editor_info:
            return "break"
        if editor_info is not None and len(editor_info) >= 2 and editor_info[1] == "Escape":
            self.destroy_text_editor("Escape")
            self.close_dropdown_window(r)
            return "break"
        if self.text_editor is not None:
            self.text_editor_value = self.text_editor.get()
        if destroy:
            self.destroy_text_editor()
        if set_data_on_close:
            if r is None and editor_info is not None and len(editor_info) >= 2:
                r = editor_info[0]
            datarn = r if self.MT.all_rows_displayed else self.MT.displayed_rows[r]
            event_data = event_dict(
                name="end_edit_index",
                sheet=self.parentframe.name,
                cells_index={datarn: self.get_cell_data(datarn)},
                key=editor_info[1] if len(editor_info) >= 2 else "FocusOut",
                value=self.text_editor_value,
                location=r,
                boxes=self.MT.get_boxes(),
                selected=self.MT.currently_selected(),
            )
            if self.extra_end_edit_cell_func is None and self.input_valid_for_cell(datarn, self.text_editor_value):
                self.set_cell_data_undo(
                    r,
                    datarn=datarn,
                    value=self.text_editor_value,
                    check_input_valid=False,
                )
            elif (
                self.extra_end_edit_cell_func is not None
                and not self.MT.edit_cell_validation
                and self.input_valid_for_cell(datarn, self.text_editor_value)
            ):
                self.set_cell_data_undo(
                    r,
                    datarn=datarn,
                    value=self.text_editor_value,
                    check_input_valid=False,
                )
                self.extra_end_edit_cell_func(event_data)
            elif self.extra_end_edit_cell_func is not None and self.MT.edit_cell_validation:
                validation = self.extra_end_edit_cell_func(event_data)
                if validation is not None:
                    self.text_editor_value = validation
                    if self.input_valid_for_cell(datarn, self.text_editor_value):
                        self.set_cell_data_undo(
                            r,
                            datarn=datarn,
                            value=self.text_editor_value,
                            check_input_valid=False,
                        )
        if move_down:
            pass
        self.close_dropdown_window(r)
        if recreate:
            self.MT.recreate_all_selection_boxes()
        if redraw:
            self.MT.refresh()
        if editor_info is not None and len(editor_info) >= 2 and editor_info[1] != "FocusOut":
            self.focus_set()
        return "break"

    # internal event use
    def set_cell_data_undo(
        self,
        r=0,
        datarn=None,
        value="",
        cell_resize=True,
        undo=True,
        redraw=True,
        check_input_valid=True,
    ):
        if datarn is None:
            datarn = r if self.MT.all_rows_displayed else self.MT.displayed_rows[r]
        event_data = event_dict(
            name="edit_index",
            sheet=self.parentframe.name,
            cells_index={datarn: self.get_cell_data(datarn)},
            boxes=self.MT.get_boxes(),
            selected=self.MT.currently_selected(),
        )
        if isinstance(self.MT._row_index, int):
            self.MT.set_cell_data_undo(r=r, c=self.MT._row_index, datarn=datarn, value=value, undo=True)
        else:
            self.fix_index(datarn)
            if not check_input_valid or self.input_valid_for_cell(datarn, value):
                if self.MT.undo_enabled and undo:
                    self.MT.undo_stack.append(ev_stack_dict(event_data))
                self.set_cell_data(datarn=datarn, value=value)
        if cell_resize and self.MT.cell_auto_resize_enabled:
            self.set_row_height_run_binding(r, only_set_if_too_small=False)
        if redraw:
            self.MT.refresh()
        self.MT.sheet_modified(event_data)

    def set_cell_data(self, datarn=None, value=""):
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

    def cell_equal_to(self, datarn, value):
        self.fix_index(datarn)
        if isinstance(self.MT._row_index, list):
            return self.MT._row_index[datarn] == value
        elif isinstance(self.MT._row_index, int):
            return self.MT.cell_equal_to(datarn, self.MT._row_index, value)

    def get_cell_data(self, datarn, get_displayed=False, none_to_empty_str=False, redirect_int=False):
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

    def get_valid_cell_data_as_str(self, datarn, fix=True) -> str:
        kwargs = self.get_cell_kwargs(datarn, key="dropdown")
        if kwargs and kwargs["text"] is not None:
            return f"{kwargs['text']}"
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
        if not value and self.show_default_index_for_empty:
            value = get_n2a(datarn, self.default_index)
        return value

    def get_value_for_empty_cell(self, datarn, r_ops=True):
        if self.get_cell_kwargs(datarn, key="checkbox", cell=r_ops):
            return False
        kwargs = self.get_cell_kwargs(datarn, key="dropdown", cell=r_ops)
        if kwargs and kwargs["validate_input"] and kwargs["values"]:
            return kwargs["values"][0]
        return ""

    def get_empty_index_seq(self, end, start=0, r_ops=True):
        return [self.get_value_for_empty_cell(datarn, r_ops=r_ops) for datarn in range(start, end)]

    def fix_index(self, datarn=None, fix_values=tuple()):
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
        if fix_values:
            for rn, v in enumerate(islice(self.MT._row_index, fix_values[0], fix_values[1])):
                if not self.input_valid_for_cell(rn, v):
                    self.MT._row_index[rn] = self.get_value_for_empty_cell(rn)

    def set_row_height_run_binding(self, r, only_set_if_too_small=True):
        old_height = self.MT.row_positions[r + 1] - self.MT.row_positions[r]
        new_height = self.set_row_height(r, only_set_if_too_small=only_set_if_too_small)
        if self.row_height_resize_func is not None and old_height != new_height:
            self.row_height_resize_func(
                event_dict(
                    name="resize",
                    sheet=self.parentframe.name,
                    resized_rows={r: {"old_size": old_height, "new_size": new_height}},
                )
            )

    # internal event use
    def click_checkbox(self, r, datarn=None, undo=True, redraw=True):
        if datarn is None:
            datarn = r if self.MT.all_rows_displayed else self.MT.displayed_rows[r]
        kwargs = self.get_cell_kwargs(datarn, key="checkbox")
        if kwargs["state"] == "normal":
            pre_edit_value = self.get_cell_data(datarn)
            if isinstance(self.MT._row_index, list):
                value = not self.MT._row_index[datarn] if type(self.MT._row_index[datarn]) == bool else False
            elif isinstance(self.MT._row_index, int):
                value = (
                    not self.MT.data[datarn][self.MT._row_index]
                    if type(self.MT.data[datarn][self.MT._row_index]) == bool
                    else False
                )
            else:
                value = False
            self.set_cell_data_undo(r, datarn=datarn, value=value, cell_resize=False)
            event_data = event_dict(
                name="end_edit_index",
                sheet=self.parentframe.name,
                cells_index={datarn: pre_edit_value},
                key="??",
                value=value,
                location=r,
                boxes=self.MT.get_boxes(),
                selected=self.MT.currently_selected(),
            )
            if kwargs["check_function"] is not None:
                kwargs["check_function"](event_data)
            try_binding(self.extra_end_edit_cell_func, event_data)
        if redraw:
            self.MT.refresh()

    def get_dropdown_height_anchor(self, datarn, text_editor_h=None):
        win_h = 5
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
        space_bot = self.MT.get_space_bot(0, text_editor_h)
        win_h2 = int(win_h)
        if win_h > space_bot:
            win_h = space_bot - 1
        if win_h < self.MT.index_txt_height + 5:
            win_h = self.MT.index_txt_height + 5
        elif win_h > win_h2:
            win_h = win_h2
        return win_h, "nw"

    # r is displayed row
    def open_dropdown_window(self, r, datarn=None, event=None):
        self.destroy_text_editor("Escape")
        self.destroy_opened_dropdown_window()
        if datarn is None:
            datarn = r if self.MT.all_rows_displayed else self.MT.displayed_rows[r]
        kwargs = self.get_cell_kwargs(datarn, key="dropdown")
        if kwargs["state"] == "normal":
            if not self.open_text_editor(event=event, r=r, dropdown=True):
                return
        win_h, anchor = self.get_dropdown_height_anchor(datarn)
        window = self.parentframe.dropdown_class(
            self.MT.winfo_toplevel(),
            r,
            0,
            width=self.current_width,
            height=win_h,
            font=self.MT.index_font,
            colors={
                "bg": self.MT.popup_menu_bg,
                "fg": self.MT.popup_menu_fg,
                "highlight_bg": self.MT.popup_menu_highlight_bg,
                "highlight_fg": self.MT.popup_menu_highlight_fg,
            },
            outline_color=self.MT.popup_menu_fg,
            values=kwargs["values"],
            close_dropdown_window=self.close_dropdown_window,
            search_function=kwargs["search_function"],
            arrowkey_RIGHT=self.MT.arrowkey_RIGHT,
            arrowkey_LEFT=self.MT.arrowkey_LEFT,
            align="w",
            single_index="r",
        )
        ypos = self.MT.row_positions[r + 1]
        kwargs["canvas_id"] = self.create_window((0, ypos), window=window, anchor=anchor)
        if kwargs["state"] == "normal":
            self.text_editor.textedit.bind(
                "<<TextModified>>",
                lambda x: window.search_and_see(
                    event_dict(
                        name="index_dropdown_modified",
                        sheet=self.parentframe.name,
                        value=self.text_editor.get(),
                        location=r,
                        boxes=self.MT.get_boxes(),
                        selected=self.MT.currently_selected(),
                    )
                ),
            )
            if kwargs["modified_function"] is not None:
                window.modified_function = kwargs["modified_function"]
            self.update_idletasks()
            try:
                self.after(1, lambda: self.text_editor.textedit.focus())
                self.after(2, self.text_editor.scroll_to_bottom())
            except Exception:
                return
            redraw = False
        else:
            window.bind("<FocusOut>", lambda x: self.close_dropdown_window(r))
            self.update_idletasks()
            window.focus_set()
            redraw = True
        self.existing_dropdown_window = window
        kwargs["window"] = window
        self.existing_dropdown_canvas_id = kwargs["canvas_id"]
        if redraw:
            self.MT.main_table_redraw_grid_and_text(redraw_header=False, redraw_row_index=True, redraw_table=False)

    # r is displayed row
    def close_dropdown_window(self, r=None, selection=None, redraw=True):
        if r is not None and selection is not None:
            datarn = r if self.MT.all_rows_displayed else self.MT.displayed_rows[r]
            kwargs = self.get_cell_kwargs(datarn, key="dropdown")
            pre_edit_value = self.get_cell_data(datarn)
            event_data = event_dict(
                name="end_edit_index",
                sheet=self.parentframe.name,
                cells_header={datarn: pre_edit_value},
                key="??",
                value=selection,
                location=r,
                boxes=self.MT.get_boxes(),
                selected=self.MT.currently_selected(),
            )
            if kwargs["select_function"] is not None:  # user has specified a selection function
                kwargs["select_function"](event_data)
            if self.extra_end_edit_cell_func is None:
                self.set_cell_data_undo(r, datarn=datarn, value=selection, redraw=not redraw)
            elif self.extra_end_edit_cell_func is not None and self.MT.edit_cell_validation:
                validation = self.extra_end_edit_cell_func(event_data)
                if validation is not None:
                    selection = validation
                self.set_cell_data_undo(r, datarn=datarn, value=selection, redraw=not redraw)
            elif self.extra_end_edit_cell_func is not None and not self.MT.edit_cell_validation:
                self.set_cell_data_undo(r, datarn=datarn, value=selection, redraw=not redraw)
                self.extra_end_edit_cell_func(event_data)
            self.focus_set()
            self.MT.recreate_all_selection_boxes()
        self.destroy_text_editor("Escape")
        self.destroy_opened_dropdown_window(r)
        if redraw:
            self.MT.refresh()

    def get_existing_dropdown_coords(self):
        if self.existing_dropdown_window is not None:
            return int(self.existing_dropdown_window.r)
        return None

    def mouseclick_outside_editor_or_dropdown(self):
        closed_dd_coords = self.get_existing_dropdown_coords()
        if self.text_editor_loc is not None and self.text_editor is not None:
            self.close_text_editor(editor_info=(self.text_editor_loc, "ButtonPress-1"))
        else:
            self.destroy_text_editor("Escape")
        if closed_dd_coords is not None:
            self.destroy_opened_dropdown_window(
                closed_dd_coords
            )  # displayed coords not data, necessary for b1 function
        return closed_dd_coords

    def mouseclick_outside_editor_or_dropdown_all_canvases(self):
        self.CH.mouseclick_outside_editor_or_dropdown()
        self.MT.mouseclick_outside_editor_or_dropdown()
        return self.mouseclick_outside_editor_or_dropdown()

    # r is displayed row, function can have two None args
    def destroy_opened_dropdown_window(self, r=None, datarn=None):
        if r is None and datarn is None and self.existing_dropdown_window is not None:
            r = self.get_existing_dropdown_coords()
        if r is not None or datarn is not None:
            if datarn is None:
                datarn_ = r if self.MT.all_rows_displayed else self.MT.displayed_rows[r]
            else:
                datarn_ = r
        else:
            datarn_ = None
        try:
            self.delete(self.existing_dropdown_canvas_id)
        except Exception:
            pass
        self.existing_dropdown_canvas_id = None
        try:
            self.existing_dropdown_window.destroy()
        except Exception:
            pass
        kwargs = self.get_cell_kwargs(datarn_, key="dropdown")
        if kwargs:
            kwargs["canvas_id"] = "no dropdown open"
            kwargs["window"] = "no dropdown open"
            try:
                self.delete(kwargs["canvas_id"])
            except Exception:
                pass
        self.existing_dropdown_window = None

    def get_cell_kwargs(self, datarn, key="dropdown", cell=True, entire=True):
        if cell and datarn in self.cell_options and key in self.cell_options[datarn]:
            return self.cell_options[datarn][key]
        if entire and key in self.options:
            return self.options[key]
        return {}
