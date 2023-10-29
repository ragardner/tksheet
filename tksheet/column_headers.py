from __future__ import annotations

import tkinter as tk
from collections import defaultdict
from itertools import (
    chain,
    cycle,
    islice,
)
from math import ceil, floor

from .formatters import is_bool_like, try_to_bool
from .functions import (
    consecutive_chunks,
    coords_tag_to_int_tuple,
    ev_stack_dict,
    event_dict,
    get_checkbox_points,
    get_n2a,
    is_contiguous,
    try_binding,
)
from .other_classes import (
    DraggedRowColumn,
    TextEditor,
)
from .vars import (
    USER_OS,
    Color_Map,
    rc_binding,
    symbols_set,
)


class ColumnHeaders(tk.Canvas):
    def __init__(self, *args, **kwargs):
        tk.Canvas.__init__(
            self,
            kwargs["parentframe"],
            background=kwargs["header_bg"],
            highlightthickness=0,
        )
        self.parentframe = kwargs["parentframe"]
        self.current_height = None  # is set from within MainTable() __init__ or from Sheet parameters
        self.MT = None  # is set from within MainTable() __init__
        self.RI = None  # is set from within MainTable() __init__
        self.TL = None  # is set from within TopLeftRectangle() __init__
        self.popup_menu_loc = None
        self.extra_begin_edit_cell_func = None
        self.extra_end_edit_cell_func = None
        self.text_editor = None
        self.text_editor_id = None
        self.text_editor_loc = None
        self.centre_alignment_text_mod_indexes = (slice(1, None), slice(None, -1))
        self.c_align_cyc = cycle(self.centre_alignment_text_mod_indexes)
        self.b1_pressed_loc = None
        self.existing_dropdown_canvas_id = None
        self.existing_dropdown_window = None
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

        self.disp_text = {}
        self.disp_high = {}
        self.disp_grid = {}
        self.disp_fill_sels = {}
        self.disp_resize_lines = {}
        self.disp_dropdown = {}
        self.disp_checkbox = {}
        self.hidd_text = {}
        self.hidd_high = {}
        self.hidd_grid = {}
        self.hidd_fill_sels = {}
        self.hidd_resize_lines = {}
        self.hidd_dropdown = {}
        self.hidd_checkbox = {}

        self.column_drag_and_drop_perform = kwargs["column_drag_and_drop_perform"]
        self.default_header = kwargs["default_header"].lower()
        self.header_bg = kwargs["header_bg"]
        self.header_fg = kwargs["header_fg"]
        self.header_grid_fg = kwargs["header_grid_fg"]
        self.header_border_fg = kwargs["header_border_fg"]
        self.header_selected_cells_bg = kwargs["header_selected_cells_bg"]
        self.header_selected_cells_fg = kwargs["header_selected_cells_fg"]
        self.header_selected_columns_bg = kwargs["header_selected_columns_bg"]
        self.header_selected_columns_fg = kwargs["header_selected_columns_fg"]
        self.header_hidden_columns_expander_bg = kwargs["header_hidden_columns_expander_bg"]
        self.show_default_header_for_empty = kwargs["show_default_header_for_empty"]
        self.drag_and_drop_bg = kwargs["drag_and_drop_bg"]
        self.resizing_line_fg = kwargs["resizing_line_fg"]
        self.align = kwargs["header_align"]
        self.basic_bindings()

    def basic_bindings(self, enable=True):
        if enable:
            self.bind("<Motion>", self.mouse_motion)
            self.bind("<ButtonPress-1>", self.b1_press)
            self.bind("<B1-Motion>", self.b1_motion)
            self.bind("<ButtonRelease-1>", self.b1_release)
            self.bind("<Double-Button-1>", self.double_b1)
            self.bind(rc_binding, self.rc)
            self.bind("<MouseWheel>", self.mousewheel)
        else:
            self.unbind("<Motion>")
            self.unbind("<ButtonPress-1>")
            self.unbind("<B1-Motion>")
            self.unbind("<ButtonRelease-1>")
            self.unbind("<Double-Button-1>")
            self.unbind(rc_binding)
            self.unbind("<MouseWheel>")

    def mousewheel(self, event: object):
        maxlines = 0
        if isinstance(self.MT._headers, int):
            if len(self.MT.data) > self.MT._headers:
                maxlines = max(
                    len(
                        self.MT.get_valid_cell_data_as_str(self.MT._headers, datacn, get_displayed=True)
                        .rstrip()
                        .split("\n")
                    )
                    for datacn in range(len(self.MT.data[self.MT._headers]))
                )
        elif isinstance(self.MT._headers, (list, tuple)):
            maxlines = max(
                len(e.rstrip().split("\n")) if isinstance(e, str) else len(f"{e}".rstrip().split("\n"))
                for e in self.MT._headers
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

    def set_height(self, new_height, set_TL=False):
        self.current_height = new_height
        try:
            self.config(height=new_height)
        except Exception:
            return
        if set_TL and self.TL is not None:
            self.TL.set_dimensions(new_h=new_height)

    def enable_bindings(self, binding):
        if binding == "column_width_resize":
            self.width_resizing_enabled = True
        if binding == "column_height_resize":
            self.height_resizing_enabled = True
        if binding == "double_click_column_resize":
            self.double_click_resizing_enabled = True
        if binding == "column_select":
            self.col_selection_enabled = True
        if binding == "drag_and_drop":
            self.drag_and_drop_enabled = True
        if binding == "hide_columns":
            self.hide_columns_enabled = True

    def disable_bindings(self, binding):
        if binding == "column_width_resize":
            self.width_resizing_enabled = False
        if binding == "column_height_resize":
            self.height_resizing_enabled = False
        if binding == "double_click_column_resize":
            self.double_click_resizing_enabled = False
        if binding == "column_select":
            self.col_selection_enabled = False
        if binding == "drag_and_drop":
            self.drag_and_drop_enabled = False
        if binding == "hide_columns":
            self.hide_columns_enabled = False

    def check_mouse_position_width_resizers(self, x, y):
        for c, (x1, y1, x2, y2) in self.visible_col_dividers.items():
            if x >= x1 and y >= y1 and x <= x2 and y <= y2:
                return c

    def rc(self, event: object):
        self.mouseclick_outside_editor_or_dropdown_all_canvases()
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
        if self.extra_rc_func is not None:
            self.extra_rc_func(event)
        if popup_menu is not None:
            self.popup_menu_loc = c
            popup_menu.tk_popup(event.x_root, event.y_root)

    def ctrl_b1_press(self, event: object):
        self.mouseclick_outside_editor_or_dropdown_all_canvases()
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
                    if self.ctrl_selection_binding_func is not None:
                        self.ctrl_selection_binding_func(
                            self.MT.get_select_event(being_drawn_item=self.being_drawn_item),
                        )
                elif c_selected:
                    self.MT.deselect(c=c)
        elif not self.MT.ctrl_select_enabled:
            self.b1_press(event)

    def ctrl_shift_b1_press(self, event: object):
        self.mouseclick_outside_editor_or_dropdown_all_canvases()
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
                    currently_selected = self.MT.currently_selected()
                    self.MT.delete_item(self.MT.currently_selected().tags[2])
                    if currently_selected and currently_selected.type_ == "column":
                        box = self.get_shift_select_box(c, currently_selected.column)
                        self.being_drawn_item = self.MT.create_selection_box(*box, set_current=currently_selected)
                    else:
                        self.being_drawn_item = self.add_selection(c, run_binding_func=False, set_as_current=True)
                        box = self.MT.get_box_from_item(self.being_drawn_item)
                    self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
                    if self.ctrl_selection_binding_func is not None:
                        self.ctrl_selection_binding_func(
                            self.MT.get_select_event(being_drawn_item=self.being_drawn_item),
                        )
                elif c_selected:
                    self.dragged_col = DraggedRowColumn(
                        dragged=c,
                        to_move=sorted(self.MT.get_selected_cols()),
                    )
        elif not self.MT.ctrl_select_enabled:
            self.shift_b1_press(event)

    def shift_b1_press(self, event: object):
        self.mouseclick_outside_editor_or_dropdown_all_canvases()
        x = event.x
        c = self.MT.identify_col(x=x)
        if (self.drag_and_drop_enabled or self.col_selection_enabled) and self.rsz_h is None and self.rsz_w is None:
            if c < len(self.MT.col_positions) - 1:
                c_selected = self.MT.col_selected(c)
                if not c_selected and self.col_selection_enabled:
                    currently_selected = self.MT.currently_selected()
                    if currently_selected and currently_selected.type_ == "column":
                        self.MT.deselect("all", redraw=False)
                        box = self.get_shift_select_box(c, currently_selected.column)
                        self.being_drawn_item = self.MT.create_selection_box(*box, set_current=currently_selected)
                    else:
                        self.being_drawn_item = self.add_selection(c, run_binding_func=False, set_as_current=True)
                        box = self.MT.get_box_from_item(self.being_drawn_item)
                    self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
                    if self.shift_selection_binding_func is not None:
                        self.shift_selection_binding_func(
                            self.MT.get_select_event(being_drawn_item=self.being_drawn_item),
                        )
                elif c_selected:
                    self.dragged_col = DraggedRowColumn(
                        dragged=c,
                        to_move=sorted(self.MT.get_selected_cols()),
                    )

    def get_shift_select_box(self, c, min_c):
        if c > min_c:
            return (0, min_c, len(self.MT.row_positions) - 1, c + 1, "columns")
        elif c < min_c:
            return (0, c, len(self.MT.row_positions) - 1, min_c + 1, "columns")

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
                self.itemconfig(t, tags=("",), state="hidden")
                self.hidd_resize_lines[t] = False

    def mouse_motion(self, event: object):
        if not self.currently_resizing_height and not self.currently_resizing_width:
            x = self.canvasx(event.x)
            y = self.canvasy(event.y)
            mouse_over_resize = False
            mouse_over_selected = False
            if self.width_resizing_enabled:
                c = self.check_mouse_position_width_resizers(x, y)
                if c is not None:
                    self.rsz_w, mouse_over_resize = c, True
                    self.config(cursor="sb_h_double_arrow")
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
                        self.config(cursor="sb_v_double_arrow")
                        self.rsz_h = True
                        mouse_over_resize = True
                    else:
                        self.rsz_h = None
                except Exception:
                    self.rsz_h = None
            if not mouse_over_resize:
                if self.MT.col_selected(self.MT.identify_col(event, allow_end=False)):
                    self.config(cursor="hand2")
                    mouse_over_selected = True
            if not mouse_over_resize and not mouse_over_selected:
                self.MT.reset_mouse_motion_creations()
        if self.extra_motion_func is not None:
            self.extra_motion_func(event)

    def double_b1(self, event: object):
        self.mouseclick_outside_editor_or_dropdown_all_canvases()
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
                        sheet=self.parentframe.name,
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
        if self.extra_double_b1_func is not None:
            self.extra_double_b1_func(event)

    def b1_press(self, event: object):
        self.MT.unbind("<MouseWheel>")
        self.focus_set()
        self.closed_dropdown = self.mouseclick_outside_editor_or_dropdown_all_canvases()
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
                fill=self.resizing_line_fg,
                tag="rwl",
            )
            self.MT.create_resize_line(x, y1, x, y2, width=1, fill=self.resizing_line_fg, tag="rwl")
            self.create_resize_line(
                line2x,
                0,
                line2x,
                self.current_height,
                width=1,
                fill=self.resizing_line_fg,
                tag="rwl2",
            )
            self.MT.create_resize_line(line2x, y1, line2x, y2, width=1, fill=self.resizing_line_fg, tag="rwl2")
        elif self.height_resizing_enabled and self.rsz_w is None and self.rsz_h is not None:
            x1, y1, x2, y2 = self.MT.get_canvas_visible_area()
            self.currently_resizing_height = True
            y = event.y
            if y < self.MT.min_header_height:
                y = int(self.MT.min_header_height)
            self.new_col_height = y
            self.create_resize_line(x1, y, x2, y, width=1, fill=self.resizing_line_fg, tag="rhl")
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
        if self.extra_b1_press_func is not None:
            self.extra_b1_press_func(event)

    def b1_motion(self, event: object):
        x1, y1, x2, y2 = self.MT.get_canvas_visible_area()
        if self.width_resizing_enabled and self.rsz_w is not None and self.currently_resizing_width:
            x = self.canvasx(event.x)
            size = x - self.MT.col_positions[self.rsz_w - 1]
            if size >= self.MT.min_column_width and size < self.MT.max_column_width:
                self.delete_all_resize_and_ctrl_lines(ctrl_lines=False)
                line2x = self.MT.col_positions[self.rsz_w - 1]
                self.create_resize_line(
                    x,
                    0,
                    x,
                    self.current_height,
                    width=1,
                    fill=self.resizing_line_fg,
                    tag="rwl",
                )
                self.MT.create_resize_line(x, y1, x, y2, width=1, fill=self.resizing_line_fg, tag="rwl")
                self.create_resize_line(
                    line2x,
                    0,
                    line2x,
                    self.current_height,
                    width=1,
                    fill=self.resizing_line_fg,
                    tag="rwl2",
                )
                self.MT.create_resize_line(
                    line2x,
                    y1,
                    line2x,
                    y2,
                    width=1,
                    fill=self.resizing_line_fg,
                    tag="rwl2",
                )
        elif self.height_resizing_enabled and self.rsz_h is not None and self.currently_resizing_height:
            evy = event.y
            self.delete_all_resize_and_ctrl_lines(ctrl_lines=False)
            if evy > self.current_height:
                y = self.MT.canvasy(evy - self.current_height)
                if evy > self.MT.max_header_height:
                    evy = int(self.MT.max_header_height)
                    y = self.MT.canvasy(evy - self.current_height)
                self.new_col_height = evy
                self.MT.create_resize_line(x1, y, x2, y, width=1, fill=self.resizing_line_fg, tag="rhl")
            else:
                y = evy
                if y < self.MT.min_header_height:
                    y = int(self.MT.min_header_height)
                self.new_col_height = y
                self.create_resize_line(x1, y, x2, y, width=1, fill=self.resizing_line_fg, tag="rhl")
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
            currently_selected = self.MT.currently_selected()
            if end_col < len(self.MT.col_positions) - 1 and currently_selected:
                if currently_selected.type_ == "column":
                    box = self.get_b1_motion_box(currently_selected.column, end_col)
                    if (
                        box is not None
                        and self.being_drawn_item is not None
                        and self.MT.get_box_from_item(self.being_drawn_item) != box
                    ):
                        self.MT.deselect("all", redraw=False)
                        if box[3] - box[1] != 1:
                            self.being_drawn_item = self.MT.create_selection_box(*box, set_current=currently_selected)
                        else:
                            self.being_drawn_item = self.select_col(currently_selected.column, run_binding_func=False)
                        need_redraw = True
                        if self.drag_selection_binding_func is not None:
                            self.drag_selection_binding_func(
                                self.MT.get_select_event(being_drawn_item=self.being_drawn_item),
                            )
                if self.scroll_if_event_offscreen(event):
                    need_redraw = True
            if need_redraw:
                self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=False)
        if self.extra_b1_motion_func is not None:
            self.extra_b1_motion_func(event)

    def get_b1_motion_box(self, start_col, end_col):
        if end_col >= start_col:
            return (
                0,
                start_col,
                len(self.MT.row_positions) - 1,
                end_col + 1,
                "columns",
            )
        elif end_col < start_col:
            return (
                0,
                end_col,
                len(self.MT.row_positions) - 1,
                start_col + 1,
                "columns",
            )

    def ctrl_b1_motion(self, event: object):
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
            currently_selected = self.MT.currently_selected()
            if end_col < len(self.MT.col_positions) - 1 and currently_selected:
                if currently_selected.type_ == "column":
                    box = self.get_b1_motion_box(currently_selected.column, end_col)
                    if (
                        box is not None
                        and self.being_drawn_item is not None
                        and self.MT.get_box_from_item(self.being_drawn_item) != box
                    ):
                        self.MT.delete_item(self.being_drawn_item)
                        if box[3] - box[1] != 1:
                            self.being_drawn_item = self.MT.create_selection_box(*box, set_current=currently_selected)
                        else:
                            self.being_drawn_item = self.add_selection(
                                currently_selected.column, run_binding_func=False
                            )
                        need_redraw = True
                        if self.drag_selection_binding_func is not None:
                            self.drag_selection_binding_func(
                                self.MT.get_select_event(being_drawn_item=self.being_drawn_item),
                            )
                if self.scroll_if_event_offscreen(event):
                    need_redraw = True
            if need_redraw:
                self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=False)
        elif not self.MT.ctrl_select_enabled:
            self.b1_motion(event)

    def drag_and_drop_motion(self, event: object):
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
        xpos,
        y1,
        y2,
        cols,
    ):
        self.delete_all_resize_and_ctrl_lines()
        self.create_resize_line(
            xpos,
            0,
            xpos,
            self.current_height,
            width=3,
            fill=self.drag_and_drop_bg,
            tag="move_columns",
        )
        self.MT.create_resize_line(xpos, y1, xpos, y2, width=3, fill=self.drag_and_drop_bg, tag="move_columns")
        for chunk in consecutive_chunks(cols):
            self.MT.show_ctrl_outline(
                start_cell=(chunk[0], 0),
                end_cell=(chunk[-1] + 1, len(self.MT.row_positions) - 1),
                dash=(),
                outline=self.drag_and_drop_bg,
                delete_on_timer=False,
            )

    def delete_all_resize_and_ctrl_lines(self, ctrl_lines=True):
        self.delete_resize_lines()
        self.MT.delete_resize_lines()
        if ctrl_lines:
            self.MT.delete_ctrl_outlines()

    def scroll_if_event_offscreen(self, event: object):
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

    def fix_xview(self):
        xcheck = self.xview()
        if xcheck and xcheck[0] < 0:
            self.MT.set_xviews("moveto", 0)
        elif len(xcheck) > 1 and xcheck[1] > 1:
            self.MT.set_xviews("moveto", 1)

    def event_over_dropdown(self, c, datacn, event: object, canvasx):
        if (
            event.y < self.MT.header_txt_height + 5
            and self.get_cell_kwargs(datacn, key="dropdown")
            and canvasx < self.MT.col_positions[c + 1]
            and canvasx > self.MT.col_positions[c + 1] - self.MT.header_txt_height - 4
        ):
            return True
        return False

    def event_over_checkbox(self, c, datacn, event: object, canvasx):
        if (
            event.y < self.MT.header_txt_height + 5
            and self.get_cell_kwargs(datacn, key="checkbox")
            and canvasx < self.MT.col_positions[c] + self.MT.header_txt_height + 4
        ):
            return True
        return False

    def b1_release(self, event: object):
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
        if self.width_resizing_enabled and self.rsz_w is not None and self.currently_resizing_width:
            self.currently_resizing_width = False
            new_col_pos = int(self.coords("rwl")[0])
            self.delete_all_resize_and_ctrl_lines(ctrl_lines=False)
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
                        sheet=self.parentframe.name,
                        resized_columns={self.rsz_w - 1: {"old_size": old_width, "new_size": new_width}},
                    )
                )
        elif self.height_resizing_enabled and self.rsz_h is not None and self.currently_resizing_height:
            self.currently_resizing_height = False
            self.delete_all_resize_and_ctrl_lines(ctrl_lines=False)
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
            self.delete_all_resize_and_ctrl_lines()
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
                if c >= len(self.MT.col_positions) - 1:
                    c -= 1
                event_data = event_dict(
                    name="move_columns",
                    sheet=self.parentframe.name,
                    boxes=self.MT.get_boxes(),
                    selected=self.MT.currently_selected(),
                    value=c,
                )
                if try_binding(self.ch_extra_begin_drag_drop_func, event_data, "begin_move_columns"):
                    data_new_idxs, disp_new_idxs, event_data = self.MT.move_columns_adjust_options_dict(
                        *self.MT.get_args_for_move_columns(
                            move_to=c,
                            to_move=self.dragged_col.to_move,
                        ),
                        move_data=self.column_drag_and_drop_perform,
                        event_data=event_data,
                    )
                    event_data["moved"]["columns"] = {
                        "data": data_new_idxs,
                        "displayed": disp_new_idxs,
                    }
                    if self.MT.undo_enabled:
                        self.MT.undo_stack.append(ev_stack_dict(event_data))
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
                if self.event_over_dropdown(c, datacn, event, canvasx) or self.event_over_checkbox(
                    c, datacn, event, canvasx
                ):
                    self.open_cell(event)
            else:
                self.mouseclick_outside_editor_or_dropdown_all_canvases()
            self.b1_pressed_loc = None
            self.closed_dropdown = None
        self.dragged_col = None
        self.currently_resizing_width = False
        self.currently_resizing_height = False
        self.rsz_w = None
        self.rsz_h = None
        self.mouse_motion(event)
        if self.extra_b1_release_func is not None:
            self.extra_b1_release_func(event)

    def toggle_select_col(
        self,
        column,
        add_selection=True,
        redraw=True,
        run_binding_func=True,
        set_as_current=True,
    ):
        if add_selection:
            if self.MT.col_selected(column):
                fill_iid = self.MT.deselect(c=column, redraw=redraw)
            else:
                fill_iid = self.add_selection(
                    c=column,
                    redraw=redraw,
                    run_binding_func=run_binding_func,
                    set_as_current=set_as_current,
                )
        else:
            if self.MT.col_selected(column):
                fill_iid = self.MT.deselect(c=column, redraw=redraw)
            else:
                fill_iid = self.select_col(column, redraw=redraw)
        return fill_iid

    def select_col(self, c, redraw=False, run_binding_func=True):
        self.MT.deselect("all", redraw=False)
        box = (0, c, len(self.MT.row_positions) - 1, c + 1, "columns")
        fill_iid = self.MT.create_selection_box(*box)
        if redraw:
            self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
        if run_binding_func:
            self.MT.run_selection_binding("columns")
        return fill_iid

    def add_selection(self, c, redraw=False, run_binding_func=True, set_as_current=True):
        box = (0, c, len(self.MT.row_positions) - 1, c + 1, "columns")
        fill_iid = self.MT.create_selection_box(*box, set_current=set_as_current)
        if redraw:
            self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
        if run_binding_func:
            self.MT.run_selection_binding("columns")
        return fill_iid

    def get_cell_dimensions(self, datacn):
        txt = self.get_valid_cell_data_as_str(datacn, fix=False)
        if txt:
            self.MT.txt_measure_canvas.itemconfig(self.MT.txt_measure_canvas_text, text=txt, font=self.MT.header_font)
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

    def set_height_of_header_to_text(self, text=None, only_increase=False):
        if (
            text is None
            and not self.MT._headers
            and isinstance(self.MT._headers, list)
            or isinstance(self.MT._headers, int)
            and self.MT._headers >= len(self.MT.data)
        ):
            return
        qconf = self.MT.txt_measure_canvas.itemconfig
        qbbox = self.MT.txt_measure_canvas.bbox
        qtxtm = self.MT.txt_measure_canvas_text
        new_height = self.MT.min_header_height
        self.fix_header()
        if text is not None:
            if text:
                qconf(qtxtm, text=text)
                b = qbbox(qtxtm)
                h = b[3] - b[1] + 5
                if h > new_height:
                    new_height = h
        else:
            if self.MT.all_columns_displayed:
                if isinstance(self.MT._headers, list):
                    iterable = range(len(self.MT._headers))
                else:
                    iterable = range(len(self.MT.data[self.MT._headers]))
            else:
                iterable = self.MT.displayed_columns
            if isinstance(self.MT._headers, list):
                for datacn in iterable:
                    w_, h = self.get_cell_dimensions(datacn)
                    if h < self.MT.min_header_height:
                        h = int(self.MT.min_header_height)
                    elif h > self.MT.max_header_height:
                        h = int(self.MT.max_header_height)
                    if h > new_height:
                        new_height = h
            elif isinstance(self.MT._headers, int):
                datarn = self.MT._headers
                for datacn in iterable:
                    txt = self.MT.get_valid_cell_data_as_str(datarn, datacn, get_displayed=True)
                    if txt:
                        qconf(qtxtm, text=txt)
                        b = qbbox(qtxtm)
                        h = b[3] - b[1] + 5
                    else:
                        h = self.MT.default_header_height
                    if h < self.MT.min_header_height:
                        h = int(self.MT.min_header_height)
                    elif h > self.MT.max_header_height:
                        h = int(self.MT.max_header_height)
                    if h > new_height:
                        new_height = h
        space_bot = self.MT.get_space_bot(0)
        if new_height > space_bot:
            new_height = space_bot
        if not only_increase or (only_increase and new_height > self.current_height):
            self.set_height(new_height, set_TL=True)
            self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
        return new_height

    def set_col_width(
        self,
        col,
        width=None,
        only_set_if_too_small=False,
        displayed_only=False,
        recreate=True,
        return_new_width=False,
    ):
        if col < 0:
            return
        qconf = self.MT.txt_measure_canvas.itemconfig
        qbbox = self.MT.txt_measure_canvas.bbox
        qtxtm = self.MT.txt_measure_canvas_text
        qtxth = self.MT.table_txt_height
        qfont = self.MT.table_font
        self.fix_header()
        if width is None:
            w = self.MT.min_column_width
            hw = self.MT.min_column_width
            if self.MT.all_rows_displayed:
                if displayed_only:
                    x1, y1, x2, y2 = self.MT.get_canvas_visible_area()
                    start_row, end_row = self.MT.get_visible_rows(y1, y2)
                else:
                    start_row, end_row = 0, len(self.MT.data)
                iterable = range(start_row, end_row)
            else:
                if displayed_only:
                    x1, y1, x2, y2 = self.MT.get_canvas_visible_area()
                    start_row, end_row = self.MT.get_visible_rows(y1, y2)
                else:
                    start_row, end_row = 0, len(self.MT.displayed_rows)
                iterable = self.MT.displayed_rows[start_row:end_row]
            datacn = col if self.MT.all_columns_displayed else self.MT.displayed_columns[col]
            # header
            hw, hh_ = self.get_cell_dimensions(datacn)
            # table
            if self.MT.data:
                for datarn in iterable:
                    txt = self.MT.get_valid_cell_data_as_str(datarn, datacn, get_displayed=True)
                    if txt:
                        qconf(qtxtm, text=txt, font=qfont)
                        b = qbbox(qtxtm)
                        if self.MT.get_cell_kwargs(datarn, datacn, key="dropdown") or self.MT.get_cell_kwargs(
                            datarn, datacn, key="checkbox"
                        ):
                            tw = b[2] - b[0] + qtxth + 7
                        else:
                            tw = b[2] - b[0] + 7
                        if tw > w:
                            w = tw
            if w > hw:
                new_width = w
            else:
                new_width = hw
        else:
            new_width = int(width)
        if new_width <= self.MT.min_column_width:
            new_width = int(self.MT.min_column_width)
        elif new_width > self.MT.max_column_width:
            new_width = int(self.MT.max_column_width)
        if only_set_if_too_small:
            if new_width <= self.MT.col_positions[col + 1] - self.MT.col_positions[col]:
                return self.MT.col_positions[col + 1] - self.MT.col_positions[col]
        if not return_new_width:
            new_col_pos = self.MT.col_positions[col] + new_width
            increment = new_col_pos - self.MT.col_positions[col + 1]
            self.MT.col_positions[col + 2 :] = [
                e + increment for e in islice(self.MT.col_positions, col + 2, len(self.MT.col_positions))
            ]
            self.MT.col_positions[col + 1] = new_col_pos
            if recreate:
                self.MT.recreate_all_selection_boxes()
        return new_width

    def set_width_of_all_cols(self, width=None, only_set_if_too_small=False, recreate=True):
        if width is None:
            if self.MT.all_columns_displayed:
                iterable = range(self.MT.total_data_cols())
            else:
                iterable = range(len(self.MT.displayed_columns))
            self.MT.set_col_positions(
                itr=(
                    self.set_col_width(
                        cn,
                        only_set_if_too_small=only_set_if_too_small,
                        recreate=False,
                        return_new_width=True,
                    )
                    for cn in iterable
                )
            )
        elif width is not None:
            if self.MT.all_columns_displayed:
                self.MT.set_col_positions(itr=(width for cn in range(self.MT.total_data_cols())))
            else:
                self.MT.set_col_positions(itr=(width for cn in range(len(self.MT.displayed_columns))))
        if recreate:
            self.MT.recreate_all_selection_boxes()

    def redraw_highlight_get_text_fg(self, fc, sc, c, c_2, c_3, selections, datacn):
        redrawn = False
        kwargs = self.get_cell_kwargs(datacn, key="highlight")
        if kwargs:
            if kwargs[0] is not None:
                c_1 = kwargs[0] if kwargs[0].startswith("#") else Color_Map[kwargs[0]]
            if "columns" in selections and c in selections["columns"]:
                tf = (
                    self.header_selected_columns_fg
                    if kwargs[1] is None or self.MT.display_selected_fg_over_highlights
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
                    self.header_selected_cells_fg
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
                tf = self.header_fg if kwargs[1] is None else kwargs[1]
                if kwargs[0] is not None:
                    fill = kwargs[0]
            if kwargs[0] is not None:
                redrawn = self.redraw_highlight(
                    fc + 1,
                    0,
                    sc,
                    self.current_height - 1,
                    fill=fill,
                    outline=self.header_fg
                    if self.get_cell_kwargs(datacn, key="dropdown") and self.MT.show_dropdown_borders
                    else "",
                    tag="hi",
                )
        elif not kwargs:
            if "columns" in selections and c in selections["columns"]:
                tf = self.header_selected_columns_fg
            elif "cells" in selections and c in selections["cells"]:
                tf = self.header_selected_cells_fg
            else:
                tf = self.header_fg
        return tf, redrawn

    def redraw_highlight(self, x1, y1, x2, y2, fill, outline, tag):
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

    def redraw_gridline(self, points, fill, width, tag):
        if self.hidd_grid:
            t, sh = self.hidd_grid.popitem()
            self.coords(t, points)
            if sh:
                self.itemconfig(
                    t,
                    fill=fill,
                    width=width,
                    tag=tag,
                    capstyle=tk.BUTT,
                    joinstyle=tk.ROUND,
                )
            else:
                self.itemconfig(
                    t,
                    fill=fill,
                    width=width,
                    tag=tag,
                    capstyle=tk.BUTT,
                    joinstyle=tk.ROUND,
                    state="normal",
                )
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
            self.redraw_highlight(x1 + 1, y1 + 1, x2, y2, fill="", outline=self.header_fg, tag=tag)
        if draw_arrow:
            topysub = floor(self.MT.header_half_txt_height / 2)
            mid_y = y1 + floor(self.MT.min_header_height / 2)
            if mid_y + topysub + 1 >= y1 + self.MT.header_txt_height - 1:
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
            tx1 = x2 - self.MT.header_txt_height + 1
            tx2 = x2 - self.MT.header_half_txt_height - 1
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
        last_col_line_pos,
        scrollpos_left,
        x_stop,
        start_col,
        end_col,
        scrollpos_right,
        col_pos_exists,
    ):
        try:
            self.configure(
                scrollregion=(
                    0,
                    0,
                    last_col_line_pos + self.MT.empty_horizontal + 2,
                    self.current_height,
                )
            )
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
        self.visible_col_dividers = {}
        self.col_height_resize_bbox = (
            scrollpos_left,
            self.current_height - 2,
            x_stop,
            self.current_height,
        )
        draw_x = self.MT.col_positions[start_col]
        yend = self.current_height - 5
        if (self.MT.show_vertical_grid or self.width_resizing_enabled) and col_pos_exists:
            points = [
                x_stop - 1,
                self.current_height - 1,
                scrollpos_left - 1,
                self.current_height - 1,
                scrollpos_left - 1,
                -1,
            ]
            for c in range(start_col + 1, end_col):
                draw_x = self.MT.col_positions[c]
                if self.width_resizing_enabled:
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
            self.redraw_gridline(points=points, fill=self.header_grid_fg, width=1, tag="v")
        top = self.canvasy(0)
        c_2 = (
            self.header_selected_cells_bg
            if self.header_selected_cells_bg.startswith("#")
            else Color_Map[self.header_selected_cells_bg]
        )
        c_3 = (
            self.header_selected_columns_bg
            if self.header_selected_columns_bg.startswith("#")
            else Color_Map[self.header_selected_columns_bg]
        )
        font = self.MT.header_font
        selections = self.get_redraw_selections(start_col, end_col)
        for c in range(start_col, end_col - 1):
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
                        dd_is_open=kwargs["window"] != "no dropdown open",
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
                        dd_is_open=kwargs["window"] != "no dropdown open",
                    )
                else:
                    mw = crightgridln - cleftgridln - 1
                    draw_x = crightgridln - 3

            elif align == "center":
                # stop = cleftgridln + 5
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
                        dd_is_open=kwargs["window"] != "no dropdown open",
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
                        mw -= box_w + 3
                    elif align == "center":
                        draw_x += ceil(box_w / 2) + 1
                        mw -= box_w + 2
                    else:
                        mw -= box_w + 1
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
                        fill=fill if kwargs["state"] == "normal" else self.header_grid_fg,
                        outline="",
                        tag="cb",
                        draw_check=draw_check,
                    )
            lns = self.get_valid_cell_data_as_str(datacn, fix=False)
            if not lns:
                continue
            lns = lns.split("\n")
            if mw > self.MT.header_txt_width and not (
                (align == "w" and (draw_x > x_stop))
                or (align == "e" and (draw_x > x_stop))
                or (align == "center" and (cleftgridln + 5 > x_stop))
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

    def get_redraw_selections(self, startc, endc):
        d = defaultdict(list)
        for item in chain(self.find_withtag("cells"), self.find_withtag("columns")):
            tags = self.gettags(item)
            if tags:
                d[tags[0]].append(coords_tag_to_int_tuple(tags[1]))
        d2 = {}
        if "cells" in d:
            d2["cells"] = {c for c in range(startc, endc) for r1, c1, r2, c2 in d["cells"] if c1 <= c and c2 > c}
        if "columns" in d:
            d2["columns"] = {c for c in range(startc, endc) for r1, c1, r2, c2 in d["columns"] if c1 <= c and c2 > c}
        return d2

    def open_cell(self, event: object = None, ignore_existing_editor=False):
        if not self.MT.anything_selected() or (not ignore_existing_editor and self.text_editor_id is not None):
            return
        currently_selected = self.MT.currently_selected()
        if not currently_selected:
            return
        x1 = currently_selected[1]
        datacn = x1 if self.MT.all_columns_displayed else self.MT.displayed_columns[x1]
        if self.get_cell_kwargs(datacn, key="readonly"):
            return
        elif self.get_cell_kwargs(datacn, key="dropdown") or self.get_cell_kwargs(datacn, key="checkbox"):
            if self.MT.event_opens_dropdown_or_checkbox(event):
                if self.get_cell_kwargs(datacn, key="dropdown"):
                    self.open_dropdown_window(x1, event=event)
                elif self.get_cell_kwargs(datacn, key="checkbox"):
                    self.click_checkbox(x1, datacn)
        elif self.edit_cell_enabled:
            self.open_text_editor(event=event, c=x1, dropdown=False)

    # displayed indexes
    def get_cell_align(self, c):
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
        c=0,
        text=None,
        state="normal",
        see=True,
        set_data_on_close=True,
        binding=None,
        dropdown=False,
    ):
        text = None
        extra_func_key = "??"
        datacn = c if self.MT.all_columns_displayed else self.MT.displayed_columns[c]
        if event is None or self.MT.event_opens_dropdown_or_checkbox(event):
            if event is not None:
                if hasattr(event, "keysym") and event.keysym == "Return":
                    extra_func_key = "Return"
                elif hasattr(event, "keysym") and event.keysym == "F2":
                    extra_func_key = "F2"
            text = self.get_cell_data(datacn, none_to_empty_str=True, redirect_int=True)
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
        self.text_editor_loc = c
        if self.extra_begin_edit_cell_func is not None:
            try:
                text = self.extra_begin_edit_cell_func(
                    event_dict(
                        name="begin_edit_header",
                        sheet=self.parentframe.name,
                        key=extra_func_key,
                        value=text,
                        location=c,
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
            if self.height_resizing_enabled:
                self.set_height_of_header_to_text(text)
            self.set_col_width_run_binding(c)
        if c == self.text_editor_loc and self.text_editor is not None:
            self.text_editor.set_text(self.text_editor.get() + "" if not isinstance(text, str) else text)
            return
        if self.text_editor is not None:
            self.destroy_text_editor()
        if see:
            has_redrawn = self.MT.see(r=0, c=c, keep_yscroll=True, check_cell_visibility=True)
            if not has_redrawn:
                self.MT.refresh()
        self.text_editor_loc = c
        x = self.MT.col_positions[c] + 1
        y = 0
        w = self.MT.col_positions[c + 1] - x
        h = self.current_height + 1
        datacn = c if self.MT.all_columns_displayed else self.MT.displayed_columns[c]
        if text is None:
            text = self.get_cell_data(datacn, none_to_empty_str=True, redirect_int=True)
        bg, fg = self.header_bg, self.header_fg
        self.text_editor = TextEditor(
            self,
            text=text,
            font=self.MT.header_font,
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
            align=self.get_cell_align(c),
            c=c,
            newline_binding=self.text_editor_has_wrapped,
        )
        self.text_editor.update_idletasks()
        self.text_editor_id = self.create_window((x, y), window=self.text_editor, anchor="nw")
        if not dropdown:
            self.text_editor.textedit.focus_set()
            self.text_editor.scroll_to_bottom()
        self.text_editor.textedit.bind("<Alt-Return>", lambda x: self.text_editor_newline_binding(c=c))
        if USER_OS == "darwin":
            self.text_editor.textedit.bind("<Option-Return>", lambda x: self.text_editor_newline_binding(c=c))
        for key, func in self.MT.text_editor_user_bound_keys.items():
            self.text_editor.textedit.bind(key, func)
        if binding is not None:
            self.text_editor.textedit.bind("<Tab>", lambda x: binding((c, "Tab")))
            self.text_editor.textedit.bind("<Return>", lambda x: binding((c, "Return")))
            self.text_editor.textedit.bind("<FocusOut>", lambda x: binding((c, "FocusOut")))
            self.text_editor.textedit.bind("<Escape>", lambda x: binding((c, "Escape")))
        elif binding is None and set_data_on_close:
            self.text_editor.textedit.bind("<Tab>", lambda x: self.close_text_editor((c, "Tab")))
            self.text_editor.textedit.bind("<Return>", lambda x: self.close_text_editor((c, "Return")))
            if not dropdown:
                self.text_editor.textedit.bind("<FocusOut>", lambda x: self.close_text_editor((c, "FocusOut")))
            self.text_editor.textedit.bind("<Escape>", lambda x: self.close_text_editor((c, "Escape")))
        else:
            self.text_editor.textedit.bind("<Escape>", lambda x: self.destroy_text_editor("Escape"))
        return True

    # displayed indexes                         #just here to receive text editor arg
    def text_editor_has_wrapped(self, r=0, c=0, check_lines=None):
        if self.width_resizing_enabled:
            datacn = c if self.MT.all_columns_displayed else self.MT.displayed_columns[c]
            curr_width = self.text_editor.winfo_width()
            new_width = curr_width + (self.MT.header_txt_height * 2)
            if new_width != curr_width:
                self.text_editor.config(width=new_width)
                self.set_col_width_run_binding(c, width=new_width, only_set_if_too_small=False)
                kwargs = self.get_cell_kwargs(datacn, key="dropdown")
                if kwargs:
                    self.itemconfig(kwargs["canvas_id"], width=new_width)
                    kwargs["window"].update_idletasks()
                    kwargs["window"]._reselect()
                self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=False, redraw_table=True)
                self.coords(self.text_editor_id, self.MT.col_positions[c] + 1, 0)

    # displayed indexes
    def text_editor_newline_binding(self, r=0, c=0, event: object = None, check_lines=True):
        if self.height_resizing_enabled:
            datacn = c if self.MT.all_columns_displayed else self.MT.displayed_columns[c]
            curr_height = self.text_editor.winfo_height()
            if (
                not check_lines
                or self.MT.get_lines_cell_height(self.text_editor.get_num_lines() + 1, font=self.MT.header_font)
                > curr_height
            ):
                new_height = curr_height + self.MT.header_xtra_lines_increment
                space_bot = self.MT.get_space_bot(0)
                if new_height > space_bot:
                    new_height = space_bot
                if new_height != curr_height:
                    self.text_editor.config(height=new_height)
                    self.set_height(new_height, set_TL=True)
                    kwargs = self.get_cell_kwargs(datacn, key="dropdown")
                    if kwargs:
                        win_h, anchor = self.get_dropdown_height_anchor(datacn, new_height)
                        self.coords(
                            kwargs["canvas_id"],
                            self.MT.col_positions[c],
                            new_height - 1,
                        )
                        self.itemconfig(kwargs["canvas_id"], anchor=anchor, height=win_h)

    def refresh_open_window_positions(self):
        if self.text_editor is not None:
            c = self.text_editor_loc
            self.text_editor.config(height=self.MT.col_positions[c + 1] - self.MT.col_positions[c])
            self.coords(
                self.text_editor_id,
                0,
                self.MT.col_positions[c],
            )
        if self.existing_dropdown_window is not None:
            c = self.get_existing_dropdown_coords()
            datacn = c if self.MT.all_columns_displayed else self.MT.displayed_columns[c]
            if self.text_editor is None:
                text_editor_h = self.MT.col_positions[c + 1] - self.MT.col_positions[c]
                anchor = self.itemcget(self.existing_dropdown_canvas_id, "anchor")
                win_h = 0
            else:
                text_editor_h = self.text_editor.winfo_height()
                win_h, anchor = self.get_dropdown_height_anchor(datacn, text_editor_h)
            if anchor == "nw":
                self.coords(
                    self.existing_dropdown_canvas_id,
                    0,
                    self.MT.col_positions[c] + text_editor_h - 1,
                )
                # self.itemconfig(self.existing_dropdown_canvas_id, anchor=anchor, height=win_h)
            elif anchor == "sw":
                self.coords(
                    self.existing_dropdown_canvas_id,
                    0,
                    self.MT.col_positions[c],
                )
                # self.itemconfig(self.existing_dropdown_canvas_id, anchor=anchor, height=win_h)

    def bind_cell_edit(self, enable=True):
        if enable:
            self.edit_cell_enabled = True
        else:
            self.edit_cell_enabled = False

    def bind_text_editor_destroy(self, binding, c):
        self.text_editor.textedit.bind("<Return>", lambda x: binding((c, "Return")))
        self.text_editor.textedit.bind("<FocusOut>", lambda x: binding((c, "FocusOut")))
        self.text_editor.textedit.bind("<Escape>", lambda x: binding((c, "Escape")))
        self.text_editor.textedit.focus_set()

    def destroy_text_editor(self, event: object = None):
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

    # c is displayed col
    def close_text_editor(
        self,
        editor_info=None,
        c=None,
        set_data_on_close=True,
        event: object = None,
        destroy=True,
        move_down=True,
        redraw=True,
        recreate=True,
    ):
        if self.focus_get() is None and editor_info:
            return
        if editor_info is not None and len(editor_info) >= 2 and editor_info[1] == "Escape":
            self.destroy_text_editor("Escape")
            self.close_dropdown_window(c)
            return
        if self.text_editor is not None:
            self.text_editor_value = self.text_editor.get()
        if destroy:
            self.destroy_text_editor()
        if set_data_on_close:
            if c is None and editor_info is not None and len(editor_info) >= 2:
                c = editor_info[0]
            datacn = c if self.MT.all_columns_displayed else self.MT.displayed_columns[c]
            event_data = event_dict(
                name="end_edit_header",
                sheet=self.parentframe.name,
                cells_header={datacn: self.get_cell_data(datacn)},
                key=editor_info[1] if len(editor_info) >= 2 else "FocusOut",
                value=self.text_editor_value,
                location=c,
                boxes=self.MT.get_boxes(),
                selected=self.MT.currently_selected(),
            )
            if self.extra_end_edit_cell_func is None and self.input_valid_for_cell(datacn, self.text_editor_value):
                self.set_cell_data_undo(
                    c,
                    datacn=datacn,
                    value=self.text_editor_value,
                    check_input_valid=False,
                )
            elif (
                self.extra_end_edit_cell_func is not None
                and not self.MT.edit_cell_validation
                and self.input_valid_for_cell(datacn, self.text_editor_value)
            ):
                self.set_cell_data_undo(
                    c,
                    datacn=datacn,
                    value=self.text_editor_value,
                    check_input_valid=False,
                )
                self.extra_end_edit_cell_func(event_data)
            elif self.extra_end_edit_cell_func is not None and self.MT.edit_cell_validation:
                validation = self.extra_end_edit_cell_func(event_data)
                if validation is not None:
                    self.text_editor_value = validation
                    if self.input_valid_for_cell(datacn, self.text_editor_value):
                        self.set_cell_data_undo(
                            c,
                            datacn=datacn,
                            value=self.text_editor_value,
                            check_input_valid=False,
                        )
        if move_down:
            pass
        self.close_dropdown_window(c)
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
        c=0,
        datacn=None,
        value="",
        cell_resize=True,
        undo=True,
        redraw=True,
        check_input_valid=True,
    ):
        if datacn is None:
            datacn = c if self.MT.all_columns_displayed else self.MT.displayed_columns[c]
        event_data = event_dict(
            name="edit_header",
            sheet=self.parentframe.name,
            cells_header={datacn: self.get_cell_data(datacn)},
            boxes=self.MT.get_boxes(),
            selected=self.MT.currently_selected(),
        )
        if isinstance(self.MT._headers, int):
            self.MT.set_cell_data_undo(r=self.MT._headers, c=c, datacn=datacn, value=value, undo=True)
        else:
            self.fix_header(datacn)
            if not check_input_valid or self.input_valid_for_cell(datacn, value):
                if self.MT.undo_enabled and undo:
                    self.MT.undo_stack.append(ev_stack_dict(event_data))
                self.set_cell_data(datacn=datacn, value=value)
        if cell_resize and self.MT.cell_auto_resize_enabled:
            if self.height_resizing_enabled:
                self.set_height_of_header_to_text(self.get_valid_cell_data_as_str(datacn, fix=False))
            self.set_col_width_run_binding(c)
        if redraw:
            self.MT.refresh()
        self.MT.sheet_modified(event_data)

    def set_cell_data(self, datacn=None, value=""):
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

    def cell_equal_to(self, datacn, value):
        self.fix_header(datacn)
        if isinstance(self.MT._headers, list):
            return self.MT._headers[datacn] == value
        elif isinstance(self.MT._headers, int):
            return self.MT.cell_equal_to(self.MT._headers, datacn, value)

    def get_cell_data(
        self,
        datacn,
        get_displayed=False,
        none_to_empty_str=False,
        redirect_int=False,
    ):
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

    def get_valid_cell_data_as_str(self, datacn, fix=True) -> str:
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
        if not value and self.show_default_header_for_empty:
            value = get_n2a(datacn, self.default_header)
        return value

    def get_value_for_empty_cell(self, datacn, c_ops=True):
        if self.get_cell_kwargs(datacn, key="checkbox", cell=c_ops):
            return False
        kwargs = self.get_cell_kwargs(datacn, key="dropdown", cell=c_ops)
        if kwargs and kwargs["validate_input"] and kwargs["values"]:
            return kwargs["values"][0]
        return ""

    def get_empty_header_seq(self, end, start=0, c_ops=True):
        return [self.get_value_for_empty_cell(datacn, c_ops=c_ops) for datacn in range(start, end)]

    def fix_header(self, datacn=None, fix_values=tuple()):
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
        if fix_values:
            for cn, v in enumerate(islice(self.MT._headers, fix_values[0], fix_values[1])):
                if not self.input_valid_for_cell(cn, v):
                    self.MT._headers[cn] = self.get_value_for_empty_cell(cn)

    # displayed indexes
    def set_col_width_run_binding(self, c, width=None, only_set_if_too_small=True):
        old_width = self.MT.col_positions[c + 1] - self.MT.col_positions[c]
        new_width = self.set_col_width(c, width=width, only_set_if_too_small=only_set_if_too_small)
        if self.column_width_resize_func is not None and old_width != new_width:
            self.column_width_resize_func(
                event_dict(
                    name="resize",
                    sheet=self.parentframe.name,
                    resized_columns={c: {"old_size": old_width, "new_size": new_width}},
                )
            )

    # internal event use
    def click_checkbox(self, c, datacn=None, undo=True, redraw=True):
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
                sheet=self.parentframe.name,
                cells_header={datacn: pre_edit_value},
                key="??",
                value=value,
                location=c,
                boxes=self.MT.get_boxes(),
                selected=self.MT.currently_selected(),
            )
            if kwargs["check_function"] is not None:
                kwargs["check_function"](event_data)
            try_binding(self.extra_end_edit_cell_func, event_data)
        if redraw:
            self.MT.refresh()

    def get_dropdown_height_anchor(self, datacn, text_editor_h=None):
        win_h = 5
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

    def open_dropdown_window(self, c, datacn=None, event: object = None):
        self.destroy_text_editor("Escape")
        self.destroy_opened_dropdown_window()
        if datacn is None:
            datacn = c if self.MT.all_columns_displayed else self.MT.displayed_columns[c]
        kwargs = self.get_cell_kwargs(datacn, key="dropdown")
        if kwargs["state"] == "normal":
            if not self.open_text_editor(event=event, c=c, dropdown=True):
                return
        win_h, anchor = self.get_dropdown_height_anchor(datacn)
        window = self.parentframe.dropdown_class(
            self.MT.winfo_toplevel(),
            0,
            c,
            width=self.MT.col_positions[c + 1] - self.MT.col_positions[c] + 1,
            height=win_h,
            font=self.MT.header_font,
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
            single_index="c",
        )
        ypos = self.current_height - 1
        kwargs["canvas_id"] = self.create_window((self.MT.col_positions[c], ypos), window=window, anchor=anchor)
        if kwargs["state"] == "normal":
            self.text_editor.textedit.bind(
                "<<TextModified>>",
                lambda x: window.search_and_see(
                    event_dict(
                        name="header_dropdown_modified",
                        sheet=self.parentframe.name,
                        value=self.text_editor.get(),
                        location=c,
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
            window.bind("<FocusOut>", lambda x: self.close_dropdown_window(c))
            self.update_idletasks()
            window.focus_set()
            redraw = True
        self.existing_dropdown_window = window
        kwargs["window"] = window
        self.existing_dropdown_canvas_id = kwargs["canvas_id"]
        if redraw:
            self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=False, redraw_table=False)

    def close_dropdown_window(self, c=None, selection=None, redraw=True):
        if c is not None and selection is not None:
            datacn = c if self.MT.all_columns_displayed else self.MT.displayed_columns[c]
            kwargs = self.get_cell_kwargs(datacn, key="dropdown")
            pre_edit_value = self.get_cell_data(datacn)
            event_data = event_dict(
                name="end_edit_header",
                sheet=self.parentframe.name,
                cells_header={datacn: pre_edit_value},
                key="??",
                value=selection,
                location=c,
                boxes=self.MT.get_boxes(),
                selected=self.MT.currently_selected(),
            )
            if kwargs["select_function"] is not None:
                kwargs["select_function"](event_data)
            if self.extra_end_edit_cell_func is None:
                self.set_cell_data_undo(c, datacn=datacn, value=selection, redraw=not redraw)
            elif self.extra_end_edit_cell_func is not None and self.MT.edit_cell_validation:
                validation = self.extra_end_edit_cell_func(event_data)
                if validation is not None:
                    selection = validation
                self.set_cell_data_undo(c, datacn=datacn, value=selection, redraw=not redraw)
            elif self.extra_end_edit_cell_func is not None and not self.MT.edit_cell_validation:
                self.set_cell_data_undo(c, datacn=datacn, value=selection, redraw=not redraw)
                self.extra_end_edit_cell_func(event_data)
            self.focus_set()
            self.MT.recreate_all_selection_boxes()
        self.destroy_text_editor("Escape")
        self.destroy_opened_dropdown_window(c)
        if redraw:
            self.MT.refresh()

    def get_existing_dropdown_coords(self):
        if self.existing_dropdown_window is not None:
            return int(self.existing_dropdown_window.c)
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
        self.RI.mouseclick_outside_editor_or_dropdown()
        self.MT.mouseclick_outside_editor_or_dropdown()
        return self.mouseclick_outside_editor_or_dropdown()

    # function can receive two None args
    def destroy_opened_dropdown_window(self, c=None, datacn=None):
        if c is None and datacn is None and self.existing_dropdown_window is not None:
            c = self.get_existing_dropdown_coords()
        if c is not None or datacn is not None:
            if datacn is None:
                datacn_ = c if self.MT.all_columns_displayed else self.MT.displayed_columns[c]
            else:
                datacn_ = datacn
        else:
            datacn_ = None
        try:
            self.delete(self.existing_dropdown_canvas_id)
        except Exception:
            pass
        self.existing_dropdown_canvas_id = None
        try:
            self.existing_dropdown_window.destroy()
        except Exception:
            pass
        kwargs = self.get_cell_kwargs(datacn_, key="dropdown")
        if kwargs:
            kwargs["canvas_id"] = "no dropdown open"
            kwargs["window"] = "no dropdown open"
            try:
                self.delete(kwargs["canvas_id"])
            except Exception:
                pass
        self.existing_dropdown_window = None

    def get_cell_kwargs(self, datacn, key="dropdown", cell=True):
        if cell and datacn in self.cell_options and key in self.cell_options[datacn]:
            return self.cell_options[datacn][key]
        return {}
