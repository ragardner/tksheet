import pickle
import tkinter as tk
import zlib
from collections import defaultdict
from itertools import accumulate, chain, cycle, islice
from math import ceil, floor

from ._tksheet_formatters import (
    is_bool_like,
    try_to_bool,
)
from ._tksheet_other_classes import (
    BeginDragDropEvent,
    DraggedRowColumn,
    DrawnItem,
    DropDownModifiedEvent,
    EditHeaderEvent,
    EndDragDropEvent,
    ResizeEvent,
    SelectColumnEvent,
    SelectionBoxEvent,
    TextCfg,
    TextEditor,
    get_checkbox_dict,
    get_dropdown_dict,
    get_n2a,
    get_seq_without_gaps_at_index,
)
from ._tksheet_vars import (
    USER_OS,
    Color_Map_,
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
        self.grid_cyctup = ("st", "end")
        self.grid_cyc = cycle(self.grid_cyctup)
        self.b1_pressed_loc = None
        self.existing_dropdown_canvas_id = None
        self.existing_dropdown_window = None
        self.closed_dropdown = None
        self.being_drawn_rect = None
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
        self.options = {}
        self.rsz_w = None
        self.rsz_h = None
        self.new_col_height = 0
        self.lines_start_at = 0
        self.currently_resizing_width = False
        self.currently_resizing_height = False
        self.ch_rc_popup_menu = None

        self.disp_text = defaultdict(set)
        self.disp_high = defaultdict(set)
        self.disp_grid = {}
        self.disp_fill_sels = {}
        self.disp_resize_lines = {}
        self.disp_dropdown = {}
        self.disp_checkbox = {}
        self.hidd_text = defaultdict(set)
        self.hidd_high = defaultdict(set)
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

    def mousewheel(self, event=None):
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

    def rc(self, event):
        self.mouseclick_outside_editor_or_dropdown_all_canvases()
        self.focus_set()
        popup_menu = None
        if self.MT.identify_col(x=event.x, allow_end=False) is None:
            self.MT.deselect("all")
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

    def ctrl_b1_press(self, event=None):
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
                    self.add_selection(c, set_as_current=True)
                    self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
                    if self.ctrl_selection_binding_func is not None:
                        self.ctrl_selection_binding_func(SelectionBoxEvent("ctrl_select_columns", (c, c + 1)))
                elif c_selected:
                    self.dragged_col = DraggedRowColumn(
                        dragged=c,
                        to_move=get_seq_without_gaps_at_index(sorted(self.MT.get_selected_cols()), c),
                    )
        elif not self.MT.ctrl_select_enabled:
            self.b1_press(event)

    def ctrl_shift_b1_press(self, event):
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
                    if currently_selected and currently_selected.type_ == "column":
                        min_c = int(currently_selected[1])
                        if c > min_c:
                            self.MT.create_selected(
                                0,
                                min_c,
                                len(self.MT.row_positions) - 1,
                                c + 1,
                                "columns",
                            )
                            func_event = tuple(range(min_c, c + 1))
                        elif c < min_c:
                            self.MT.create_selected(
                                0,
                                c,
                                len(self.MT.row_positions) - 1,
                                min_c + 1,
                                "columns",
                            )
                            func_event = tuple(range(c, min_c + 1))
                    else:
                        self.add_selection(c, set_as_current=True)
                        func_event = (c,)
                    self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
                    if self.ctrl_selection_binding_func is not None:
                        self.ctrl_selection_binding_func(SelectionBoxEvent("ctrl_select_columns", func_event))
                elif c_selected:
                    self.dragged_col = DraggedRowColumn(
                        dragged=c,
                        to_move=get_seq_without_gaps_at_index(sorted(self.MT.get_selected_cols()), c),
                    )
        elif not self.MT.ctrl_select_enabled:
            self.shift_b1_press(event)

    def shift_b1_press(self, event):
        self.mouseclick_outside_editor_or_dropdown_all_canvases()
        x = event.x
        c = self.MT.identify_col(x=x)
        if (self.drag_and_drop_enabled or self.col_selection_enabled) and self.rsz_h is None and self.rsz_w is None:
            if c < len(self.MT.col_positions) - 1:
                c_selected = self.MT.col_selected(c)
                if not c_selected and self.col_selection_enabled:
                    currently_selected = self.MT.currently_selected()
                    if currently_selected and currently_selected.type_ == "column":
                        min_c = int(currently_selected[1])
                        self.MT.delete_selection_rects(delete_current=False)
                        if c > min_c:
                            self.MT.create_selected(
                                0,
                                min_c,
                                len(self.MT.row_positions) - 1,
                                c + 1,
                                "columns",
                            )
                            func_event = tuple(range(min_c, c + 1))
                        elif c < min_c:
                            self.MT.create_selected(
                                0,
                                c,
                                len(self.MT.row_positions) - 1,
                                min_c + 1,
                                "columns",
                            )
                            func_event = tuple(range(c, min_c + 1))
                    else:
                        self.select_col(c)
                        func_event = (c,)
                    self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
                    if self.shift_selection_binding_func is not None:
                        self.shift_selection_binding_func(SelectionBoxEvent("shift_select_columns", func_event))
                elif c_selected:
                    self.dragged_col = DraggedRowColumn(
                        dragged=c,
                        to_move=get_seq_without_gaps_at_index(sorted(self.MT.get_selected_cols()), c),
                    )

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

    def double_b1(self, event=None):
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
                self.column_width_resize_func(ResizeEvent("column_width_resize", col, old_width, new_width))
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

    def b1_press(self, event=None):
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
                        to_move=get_seq_without_gaps_at_index(sorted(self.MT.get_selected_cols()), c),
                    )
                else:
                    self.being_drawn_rect = (
                        0,
                        c,
                        len(self.MT.row_positions) - 1,
                        c + 1,
                        "columns",
                    )
                    if self.MT.single_selection_enabled:
                        self.select_col(c, redraw=True)
                    elif self.MT.toggle_selection_enabled:
                        self.toggle_select_col(c, redraw=True)
        if self.extra_b1_press_func is not None:
            self.extra_b1_press_func(event)

    def b1_motion(self, event):
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
            if x > 0 and x < self.MT.col_positions[-1]:
                self.show_drag_and_drop_indicators(
                    self.drag_and_drop_motion(event),
                    y1,
                    y2,
                    self.dragged_col.to_move[0],
                    self.dragged_col.to_move[-1],
                )
        elif (
            self.MT.drag_selection_enabled and self.col_selection_enabled and self.rsz_h is None and self.rsz_w is None
        ):
            need_redraw = False
            end_col = self.MT.identify_col(x=event.x)
            currently_selected = self.MT.currently_selected()
            if end_col < len(self.MT.col_positions) - 1 and currently_selected:
                if currently_selected.type_ == "column":
                    start_col = currently_selected[1]
                    if end_col >= start_col:
                        rect = (
                            0,
                            start_col,
                            len(self.MT.row_positions) - 1,
                            end_col + 1,
                            "columns",
                        )
                        func_event = tuple(range(start_col, end_col + 1))
                    elif end_col < start_col:
                        rect = (
                            0,
                            end_col,
                            len(self.MT.row_positions) - 1,
                            start_col + 1,
                            "columns",
                        )
                        func_event = tuple(range(end_col, start_col + 1))
                    if self.being_drawn_rect != rect:
                        need_redraw = True
                        self.MT.delete_selection_rects(delete_current=False)
                        self.MT.create_selected(*rect)
                        self.being_drawn_rect = rect
                        if self.drag_selection_binding_func is not None:
                            self.drag_selection_binding_func(SelectionBoxEvent("drag_select_columns", func_event))
                if self.scroll_if_event_offscreen(event):
                    need_redraw = True
            if need_redraw:
                self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=False)
        if self.extra_b1_motion_func is not None:
            self.extra_b1_motion_func(event)

    def ctrl_b1_motion(self, event):
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
            if x > 0 and x < self.MT.col_positions[-1]:
                self.show_drag_and_drop_indicators(
                    self.drag_and_drop_motion(event),
                    y1,
                    y2,
                    self.dragged_col.to_move[0],
                    self.dragged_col.to_move[-1],
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
                    start_col = currently_selected[1]
                    if end_col >= start_col:
                        rect = (
                            0,
                            start_col,
                            len(self.MT.row_positions) - 1,
                            end_col + 1,
                            "columns",
                        )
                        func_event = tuple(range(start_col, end_col + 1))
                    elif end_col < start_col:
                        rect = (
                            0,
                            end_col,
                            len(self.MT.row_positions) - 1,
                            start_col + 1,
                            "columns",
                        )
                        func_event = tuple(range(end_col, start_col + 1))
                    if self.being_drawn_rect != rect:
                        need_redraw = True
                        if self.being_drawn_rect is not None:
                            self.MT.delete_selected(*self.being_drawn_rect)
                        self.MT.create_selected(*rect)
                        self.being_drawn_rect = rect
                        if self.drag_selection_binding_func is not None:
                            self.drag_selection_binding_func(SelectionBoxEvent("drag_select_columns", func_event))
                if self.scroll_if_event_offscreen(event):
                    need_redraw = True
            if need_redraw:
                self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=False)
        elif not self.MT.ctrl_select_enabled:
            self.b1_motion(event)

    def drag_and_drop_motion(self, event):
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
            self.MT.main_table_redraw_grid_and_text(redraw_header=True)
        elif x <= 0 and len(xcheck) > 1 and xcheck[0] > 0:
            if x >= -15:
                self.MT.xview_scroll(-1, "units")
                self.xview_scroll(-1, "units")
            else:
                self.MT.xview_scroll(-2, "units")
                self.xview_scroll(-2, "units")
            self.fix_xview()
            self.MT.main_table_redraw_grid_and_text(redraw_header=True)
        col = self.MT.identify_col(x=event.x)
        if col >= self.dragged_col.to_move[0] and col <= self.dragged_col.to_move[-1]:
            xpos = self.MT.col_positions[self.dragged_col.to_move[0]]
        else:
            if col < self.dragged_col.to_move[0]:
                xpos = self.MT.col_positions[col]
            else:
                xpos = (
                    self.MT.col_positions[col + 1]
                    if len(self.MT.col_positions) - 1 > col
                    else self.MT.col_positions[col]
                )
        return xpos

    def show_drag_and_drop_indicators(self, xpos, y1, y2, start_col, end_col):
        self.delete_all_resize_and_ctrl_lines()
        self.create_resize_line(
            xpos,
            0,
            xpos,
            self.current_height,
            width=3,
            fill=self.drag_and_drop_bg,
            tag="dd",
        )
        self.MT.create_resize_line(xpos, y1, xpos, y2, width=3, fill=self.drag_and_drop_bg, tag="dd")
        self.MT.show_ctrl_outline(
            start_cell=(start_col, 0),
            end_cell=(end_col + 1, len(self.MT.row_positions) - 1),
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
        xcheck = self.xview()
        need_redraw = False
        if event.x > self.winfo_width() and len(xcheck) > 1 and xcheck[1] < 1:
            try:
                self.MT.xview_scroll(1, "units")
                self.xview_scroll(1, "units")
            except Exception:
                pass
            self.fix_xview()
            need_redraw = True
        elif event.x < 0 and self.canvasx(self.winfo_width()) > 0 and xcheck and xcheck[0] > 0:
            try:
                self.xview_scroll(-1, "units")
                self.MT.xview_scroll(-1, "units")
            except Exception:
                pass
            self.fix_xview()
            need_redraw = True
        return need_redraw

    def fix_xview(self):
        xcheck = self.xview()
        if xcheck and xcheck[0] < 0:
            self.MT.set_xviews("moveto", 0)
        elif len(xcheck) > 1 and xcheck[1] > 1:
            self.MT.set_xviews("moveto", 1)

    def event_over_dropdown(self, c, datacn, event, canvasx):
        if (
            event.y < self.MT.header_txt_h + 5
            and self.get_cell_kwargs(datacn, key="dropdown")
            and canvasx < self.MT.col_positions[c + 1]
            and canvasx > self.MT.col_positions[c + 1] - self.MT.header_txt_h - 4
        ):
            return True
        return False

    def event_over_checkbox(self, c, datacn, event, canvasx):
        if (
            event.y < self.MT.header_txt_h + 5
            and self.get_cell_kwargs(datacn, key="checkbox")
            and canvasx < self.MT.col_positions[c] + self.MT.header_txt_h + 4
        ):
            return True
        return False

    def b1_release(self, event=None):
        if self.being_drawn_rect is not None:
            self.MT.delete_selected(*self.being_drawn_rect)
            to_sel = tuple(self.being_drawn_rect)
            self.being_drawn_rect = None
            self.MT.create_selected(*to_sel)
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
                self.column_width_resize_func(ResizeEvent("column_width_resize", self.rsz_w - 1, old_width, new_width))
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
        ):
            self.delete_all_resize_and_ctrl_lines()
            x = event.x
            c = self.MT.identify_col(x=x)
            orig_selected = self.dragged_col.to_move
            if (
                c != self.dragged_col.dragged
                and c is not None
                and (c < self.dragged_col.to_move[0] or c > self.dragged_col.to_move[-1])
                and len(orig_selected) != len(self.MT.col_positions) - 1
            ):
                rm1start = orig_selected[0]
                totalcols = len(orig_selected)
                extra_func_success = True
                if c >= len(self.MT.col_positions) - 1:
                    c -= 1
                if self.ch_extra_begin_drag_drop_func is not None:
                    try:
                        self.ch_extra_begin_drag_drop_func(
                            BeginDragDropEvent(
                                "begin_column_header_drag_drop",
                                tuple(orig_selected),
                                int(c),
                            )
                        )
                    except Exception:
                        extra_func_success = False
                if extra_func_success:
                    new_selected, dispset = self.MT.move_columns_adjust_options_dict(
                        c,
                        rm1start,
                        totalcols,
                        move_data=self.column_drag_and_drop_perform,
                    )
                    if self.MT.undo_enabled:
                        self.MT.undo_storage.append(
                            zlib.compress(pickle.dumps(("move_cols", orig_selected, new_selected)))
                        )
                    self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
                    if self.ch_extra_end_drag_drop_func is not None:
                        self.ch_extra_end_drag_drop_func(
                            EndDragDropEvent(
                                "end_column_header_drag_drop",
                                orig_selected,
                                new_selected,
                                int(c),
                            )
                        )
                    self.parentframe.emit_event("<<SheetModified>>")
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

    def readonly_header(self, columns=[], readonly=True):
        if isinstance(columns, int):
            columns_ = [columns]
        else:
            columns_ = columns
        if not readonly:
            for c in columns_:
                if c in self.cell_options and "readonly" in self.cell_options[c]:
                    del self.cell_options[c]["readonly"]
        else:
            for c in columns_:
                if c not in self.cell_options:
                    self.cell_options[c] = {}
                self.cell_options[c]["readonly"] = True

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
                self.MT.deselect(c=column, redraw=redraw)
            else:
                self.add_selection(
                    c=column,
                    redraw=redraw,
                    run_binding_func=run_binding_func,
                    set_as_current=set_as_current,
                )
        else:
            if self.MT.col_selected(column):
                self.MT.deselect(c=column, redraw=redraw)
            else:
                self.select_col(column, redraw=redraw)

    def select_col(self, c, redraw=False):
        self.MT.delete_selection_rects()
        self.MT.create_selected(0, c, len(self.MT.row_positions) - 1, c + 1, "columns")
        self.MT.set_currently_selected(0, c, type_="column")
        if redraw:
            self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
        if self.selection_binding_func is not None:
            self.selection_binding_func(SelectColumnEvent("select_column", int(c)))

    def add_selection(self, c, redraw=False, run_binding_func=True, set_as_current=True):
        if set_as_current:
            self.MT.set_currently_selected(0, c, type_="column")
        self.MT.create_selected(0, c, len(self.MT.row_positions) - 1, c + 1, "columns")
        if redraw:
            self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
        if self.selection_binding_func is not None and run_binding_func:
            self.selection_binding_func(("select_column", c))

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
            return w + self.MT.header_txt_h, h
        return w, h

    def set_height_of_header_to_text(self, text=None):
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
        qtxth = self.MT.txt_h
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
            self.MT.col_positions = list(
                accumulate(
                    chain(
                        [0],
                        (
                            self.set_col_width(
                                cn,
                                only_set_if_too_small=only_set_if_too_small,
                                recreate=False,
                                return_new_width=True,
                            )
                            for cn in iterable
                        ),
                    )
                )
            )
        elif width is not None:
            if self.MT.all_columns_displayed:
                self.MT.col_positions = list(accumulate(chain([0], (width for cn in range(self.MT.total_data_cols())))))
            else:
                self.MT.col_positions = list(
                    accumulate(chain([0], (width for cn in range(len(self.MT.displayed_columns)))))
                )
        if recreate:
            self.MT.recreate_all_selection_boxes()

    def align_cells(self, columns=[], align="global"):
        if isinstance(columns, int):
            cols = [columns]
        else:
            cols = columns
        if align == "global":
            for c in cols:
                if c in self.cell_options and "align" in self.cell_options[c]:
                    del self.cell_options[c]["align"]
        else:
            for c in cols:
                if c not in self.cell_options:
                    self.cell_options[c] = {}
                self.cell_options[c]["align"] = align

    def redraw_highlight_get_text_fg(self, fc, sc, c, c_2, c_3, selections, datacn):
        redrawn = False
        kwargs = self.get_cell_kwargs(datacn, key="highlight")
        if kwargs:
            if kwargs[0] is not None:
                c_1 = kwargs[0] if kwargs[0].startswith("#") else Color_Map_[kwargs[0]]
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
        config = (fill, outline)
        coords = (x1 - 1 if outline else x1, y1 - 1 if outline else y1, x2, y2)
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
            topysub = floor(self.MT.header_half_txt_h / 2)
            mid_y = y1 + floor(self.MT.min_header_height / 2)
            if mid_y + topysub + 1 >= y1 + self.MT.header_txt_h - 1:
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
            tx1 = x2 - self.MT.header_txt_h + 1
            tx2 = x2 - self.MT.header_half_txt_h - 1
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
        points = self.MT.get_checkbox_points(x1, y1, x2, y2)
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
            points = self.MT.get_checkbox_points(x1, y1, x2, y2, radius=4)
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
                    last_col_line_pos + self.MT.empty_horizontal,
                    self.current_height,
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
            self.grid_cyc = cycle(self.grid_cyctup)
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
                st_or_end = next(self.grid_cyc)
                if st_or_end == "st":
                    points.extend(
                        [
                            draw_x,
                            -1,
                            draw_x,
                            self.current_height,
                            self.MT.col_positions[c + 1] if len(self.MT.col_positions) - 1 > c else draw_x,
                            self.current_height,
                        ]
                    )
                elif st_or_end == "end":
                    points.extend(
                        [
                            draw_x,
                            self.current_height,
                            draw_x,
                            -1,
                            self.MT.col_positions[c + 1] if len(self.MT.col_positions) - 1 > c else draw_x,
                            -1,
                        ]
                    )
                if points:
                    self.redraw_gridline(points=points, fill=self.header_grid_fg, width=1, tag="v")
        top = self.canvasy(0)
        c_2 = (
            self.header_selected_cells_bg
            if self.header_selected_cells_bg.startswith("#")
            else Color_Map_[self.header_selected_cells_bg]
        )
        c_3 = (
            self.header_selected_columns_bg
            if self.header_selected_columns_bg.startswith("#")
            else Color_Map_[self.header_selected_columns_bg]
        )
        font = self.MT.header_font
        selections = self.get_redraw_selections(start_col, end_col)
        for c in range(start_col, end_col - 1):
            draw_y = self.MT.header_fl_ins
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
                    mw = crightgridln - cleftgridln - self.MT.header_txt_h - 2
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
                    mw = crightgridln - cleftgridln - self.MT.header_txt_h - 2
                    draw_x = crightgridln - 5 - self.MT.header_txt_h
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
                    mw = crightgridln - cleftgridln - self.MT.header_txt_h - 2
                    draw_x = cleftgridln + ceil((crightgridln - cleftgridln - self.MT.header_txt_h) / 2)
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
            kwargs = self.get_cell_kwargs(datacn, key="checkbox")
            if kwargs:
                if mw > self.MT.header_txt_h + 2:
                    box_w = self.MT.header_txt_h + 1
                    mw -= box_w
                    if align == "w":
                        draw_x += box_w + 1
                    elif align == "center":
                        draw_x += ceil(box_w / 2) + 1
                        mw -= 1
                    else:
                        mw -= 3
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
                        cleftgridln + self.MT.header_txt_h + 3,
                        self.MT.header_txt_h + 3,
                        fill=fill if kwargs["state"] == "normal" else self.header_grid_fg,
                        outline="",
                        tag="cb",
                        draw_check=draw_check,
                    )
            lns = self.get_valid_cell_data_as_str(datacn, fix=False).split("\n")
            if lns == [""]:
                if self.show_default_header_for_empty:
                    lns = (get_n2a(datacn, self.default_header),)
                else:
                    continue
            if mw > self.MT.header_txt_w and not (
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
                            self.disp_text[config._replace(txt=txt)].add(DrawnItem(iid=iid, showing=True))
                        else:
                            self.disp_text[config].add(DrawnItem(iid=iid, showing=True))
                    draw_y += self.MT.header_xtra_lines_increment
                    if draw_y - 1 > self.current_height:
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

    def get_redraw_selections(self, startc, endc):
        d = defaultdict(list)
        for item in chain(self.find_withtag("cells"), self.find_withtag("columns")):
            tags = self.gettags(item)
            d[tags[0]].append(tuple(int(e) for e in tags[1].split("_") if e))
        d2 = {}
        if "cells" in d:
            d2["cells"] = {c for c in range(startc, endc) for r1, c1, r2, c2 in d["cells"] if c1 <= c and c2 > c}
        if "columns" in d:
            d2["columns"] = {c for c in range(startc, endc) for r1, c1, r2, c2 in d["columns"] if c1 <= c and c2 > c}
        return d2

    def open_cell(self, event=None, ignore_existing_editor=False):
        if not self.MT.anything_selected() or (not ignore_existing_editor and self.text_editor_id is not None):
            return
        currently_selected = self.MT.currently_selected()
        if not currently_selected:
            return
        x1 = int(currently_selected[1])
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
        event=None,
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
        if event is None or self.MT.event_opens_dropdown_or_checkbox(event):
            if event is not None:
                if hasattr(event, "keysym") and event.keysym == "Return":
                    extra_func_key = "Return"
                elif hasattr(event, "keysym") and event.keysym == "F2":
                    extra_func_key = "F2"
            datacn = c if self.MT.all_columns_displayed else self.MT.displayed_columns[c]
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
                text = self.extra_begin_edit_cell_func(EditHeaderEvent(c, extra_func_key, text, "begin_edit_header"))
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

    # displayed indexes                             #just here to receive text editor arg
    def text_editor_has_wrapped(self, r=0, c=0, check_lines=None):
        if self.width_resizing_enabled:
            datacn = c if self.MT.all_columns_displayed else self.MT.displayed_columns[c]
            curr_width = self.text_editor.winfo_width()
            new_width = curr_width + (self.MT.header_txt_h * 2)
            if new_width != curr_width:
                self.text_editor.config(width=new_width)
                self.set_col_width_run_binding(c, width=new_width, only_set_if_too_small=False)
                kwargs = self.get_cell_kwargs(datacn, key="dropdown")
                if kwargs:
                    self.itemconfig(kwargs["canvas_id"], width=new_width)
                    kwargs["window"].update_idletasks()
                    kwargs["window"]._reselect()
                self.MT.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=False, redraw_table=True)

    # displayed indexes
    def text_editor_newline_binding(self, r=0, c=0, event=None, check_lines=True):
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

    def destroy_text_editor(self, event=None):
        if event is not None and self.extra_end_edit_cell_func is not None and self.text_editor_loc is not None:
            self.extra_end_edit_cell_func(
                EditHeaderEvent(int(self.text_editor_loc), "Escape", None, "escape_edit_header")
            )
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
        event=None,
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
                self.extra_end_edit_cell_func(
                    EditHeaderEvent(
                        c,
                        editor_info[1] if len(editor_info) >= 2 else "FocusOut",
                        f"{self.text_editor_value}",
                        "end_edit_header",
                    )
                )
            elif self.extra_end_edit_cell_func is not None and self.MT.edit_cell_validation:
                validation = self.extra_end_edit_cell_func(
                    EditHeaderEvent(
                        c,
                        editor_info[1] if len(editor_info) >= 2 else "FocusOut",
                        f"{self.text_editor_value}",
                        "end_edit_header",
                    )
                )
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
        if isinstance(self.MT._headers, int):
            self.MT.set_cell_data_undo(r=self.MT._headers, c=c, datacn=datacn, value=value, undo=True)
        else:
            self.fix_header(datacn)
            if not check_input_valid or self.input_valid_for_cell(datacn, value):
                if self.MT.undo_enabled and undo:
                    self.MT.undo_storage.append(
                        zlib.compress(
                            pickle.dumps(
                                (
                                    "edit_header",
                                    {datacn: self.MT._headers[datacn]},
                                    self.MT.get_boxes(include_current=False),
                                    self.MT.currently_selected(),
                                )
                            )
                        )
                    )
                self.set_cell_data(datacn=datacn, value=value)
        if cell_resize and self.MT.cell_auto_resize_enabled:
            if self.height_resizing_enabled:
                self.set_height_of_header_to_text(self.get_valid_cell_data_as_str(datacn, fix=False))
            self.set_col_width_run_binding(c)
        if redraw:
            self.MT.refresh()
        self.parentframe.emit_event("<<SheetModified>>")

    def set_cell_data(self, datacn=None, value=""):
        if isinstance(self.MT._headers, int):
            self.MT.set_cell_data(datarn=self.MT._headers, datacn=datacn, value=value)
        else:
            self.fix_header(datacn)
            if self.get_cell_kwargs(datacn, key="checkbox"):
                self.MT._headers[datacn] = try_to_bool(value)
            else:
                self.MT._headers[datacn] = value

    def input_valid_for_cell(self, datacn, value):
        if self.get_cell_kwargs(datacn, key="readonly"):
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

    def get_cell_data(self, datacn, get_displayed=False, none_to_empty_str=False, redirect_int=False):
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
        if kwargs and kwargs["text"] is not None:
            return f"{kwargs['text']}"
        kwargs = self.get_cell_kwargs(datacn, key="checkbox")
        if kwargs:
            return f"{kwargs['text']}"
        if isinstance(self.MT._headers, int):
            return self.MT.get_valid_cell_data_as_str(self.MT._headers, datacn, get_displayed=True)
        if fix:
            self.fix_header(datacn)
        try:
            return "" if self.MT._headers[datacn] is None else f"{self.MT._headers[datacn]}"
        except Exception:
            return ""

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
            self.column_width_resize_func(ResizeEvent("column_width_resize", c, old_width, new_width))

    # internal event use
    def click_checkbox(self, c, datacn=None, undo=True, redraw=True):
        if datacn is None:
            datacn = c if self.MT.all_columns_displayed else self.MT.displayed_columns[c]
        kwargs = self.get_cell_kwargs(datacn, key="checkbox")
        if kwargs["state"] == "normal":
            if isinstance(self.MT._headers, list):
                value = not self.MT._headers[datacn] if type(self.MT._headers[datacn]) == bool else False
            elif isinstance(self.MT._headers, int):
                value = (
                    not self.MT.data[self.MT._headers][datacn]
                    if type(self.MT.data[self.MT._headers][datacn]) == bool
                    else False
                )
            else:
                value = False
            self.set_cell_data_undo(c, datacn=datacn, value=value, cell_resize=False)
            if kwargs["check_function"] is not None:
                kwargs["check_function"](
                    (
                        0,
                        c,
                        "HeaderCheckboxClicked",
                        self.MT._headers[datacn]
                        if isinstance(self.MT._headers, list)
                        else self.MT.get_cell_data(self.MT._headers, datacn),
                    )
                )
            if self.extra_end_edit_cell_func is not None:
                self.extra_end_edit_cell_func(
                    EditHeaderEvent(
                        c,
                        "Return",
                        self.MT._headers[datacn]
                        if isinstance(self.MT._headers, list)
                        else self.MT.get_cell_data(self.MT._headers, datacn),
                        "end_edit_header",
                    )
                )
        if redraw:
            self.MT.refresh()

    def checkbox_header(self, **kwargs):
        self.destroy_opened_dropdown_window()
        if "dropdown" in self.options or "checkbox" in self.options:
            self.delete_options_dropdown_and_checkbox()
        if "checkbox" not in self.options:
            self.options["checkbox"] = {}
        self.options["checkbox"] = get_checkbox_dict(**kwargs)
        total_cols = self.MT.total_data_cols()
        if isinstance(self.MT._headers, int):
            for datacn in range(total_cols):
                self.MT.set_cell_data(datarn=self.MT._headers, datacn=datacn, value=kwargs["checked"])
        else:
            for datacn in range(total_cols):
                self.set_cell_data(datacn=datacn, value=kwargs["checked"])

    def dropdown_header(self, **kwargs):
        self.destroy_opened_dropdown_window()
        if "dropdown" in self.options or "checkbox" in self.options:
            self.delete_options_dropdown_and_checkbox()
        if "dropdown" not in self.options:
            self.options["dropdown"] = {}
        self.options["dropdown"] = get_dropdown_dict(**kwargs)
        total_cols = self.MT.total_data_cols()
        value = (
            kwargs["set_value"] if kwargs["set_value"] is not None else kwargs["values"][0] if kwargs["values"] else ""
        )
        if isinstance(self.MT._headers, int):
            for datacn in range(total_cols):
                self.MT.set_cell_data(datarn=self.MT._headers, datacn=datacn, value=value)
        else:
            for datacn in range(total_cols):
                self.set_cell_data(datacn=datacn, value=value)

    def create_checkbox(self, datacn=0, **kwargs):
        if datacn in self.cell_options and (
            "dropdown" in self.cell_options[datacn] or "checkbox" in self.cell_options[datacn]
        ):
            self.delete_cell_options_dropdown_and_checkbox(datacn)
        if datacn not in self.cell_options:
            self.cell_options[datacn] = {}
        self.cell_options[datacn]["checkbox"] = get_checkbox_dict(**kwargs)
        self.set_cell_data(datacn=datacn, value=kwargs["checked"])

    def create_dropdown(self, datacn=0, **kwargs):
        if datacn in self.cell_options and (
            "dropdown" in self.cell_options[datacn] or "checkbox" in self.cell_options[datacn]
        ):
            self.delete_cell_options_dropdown_and_checkbox(datacn)
        if datacn not in self.cell_options:
            self.cell_options[datacn] = {}
        self.cell_options[datacn]["dropdown"] = get_dropdown_dict(**kwargs)
        self.set_cell_data(
            datacn=datacn,
            value=kwargs["set_value"]
            if kwargs["set_value"] is not None
            else kwargs["values"][0]
            if kwargs["values"]
            else "",
        )

    def get_dropdown_height_anchor(self, datacn, text_editor_h=None):
        win_h = 5
        for i, v in enumerate(self.get_cell_kwargs(datacn, key="dropdown")["values"]):
            v_numlines = len(v.split("\n") if isinstance(v, str) else f"{v}".split("\n"))
            if v_numlines > 1:
                win_h += self.MT.header_fl_ins + (v_numlines * self.MT.header_xtra_lines_increment) + 5  # end of cell
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
        if win_h < self.MT.header_txt_h + 5:
            win_h = self.MT.header_txt_h + 5
        elif win_h > win_h2:
            win_h = win_h2
        return win_h, "nw"

    def open_dropdown_window(self, c, datacn=None, event=None):
        self.destroy_text_editor("Escape")
        self.destroy_opened_dropdown_window()
        if datacn is None:
            datacn = c if self.MT.all_columns_displayed else self.MT.displayed_columns[c]
        kwargs = self.get_cell_kwargs(datacn, key="dropdown")
        if kwargs["state"] == "normal":
            if not self.open_text_editor(event=event, c=c, dropdown=True):
                return
        win_h, anchor = self.get_dropdown_height_anchor(datacn)
        window = self.MT.parentframe.dropdown_class(
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
                    DropDownModifiedEvent("HeaderComboboxModified", 0, c, self.text_editor.get())
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
            if kwargs["select_function"] is not None:  # user has specified a selection function
                kwargs["select_function"](
                    EditHeaderEvent(c, "HeaderComboboxSelected", f"{selection}", "end_edit_header")
                )
            if self.extra_end_edit_cell_func is None:
                self.set_cell_data_undo(c, datacn=datacn, value=selection, redraw=not redraw)
            elif self.extra_end_edit_cell_func is not None and self.MT.edit_cell_validation:
                validation = self.extra_end_edit_cell_func(
                    EditHeaderEvent(c, "HeaderComboboxSelected", f"{selection}", "end_edit_header")
                )
                if validation is not None:
                    selection = validation
                self.set_cell_data_undo(c, datacn=datacn, value=selection, redraw=not redraw)
            elif self.extra_end_edit_cell_func is not None and not self.MT.edit_cell_validation:
                self.set_cell_data_undo(c, datacn=datacn, value=selection, redraw=not redraw)
                self.extra_end_edit_cell_func(
                    EditHeaderEvent(c, "HeaderComboboxSelected", f"{selection}", "end_edit_header")
                )
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

    def get_cell_kwargs(self, datacn, key="dropdown", cell=True, entire=True):
        if cell and datacn in self.cell_options and key in self.cell_options[datacn]:
            return self.cell_options[datacn][key]
        if entire and key in self.options:
            return self.options[key]
        return {}

    def delete_options_dropdown(self):
        self.destroy_opened_dropdown_window()
        if "dropdown" in self.options:
            del self.options["dropdown"]

    def delete_options_checkbox(self):
        if "checkbox" in self.options:
            del self.options["checkbox"]

    def delete_options_dropdown_and_checkbox(self):
        self.delete_options_dropdown()
        self.delete_options_checkbox()

    def delete_cell_options_dropdown(self, datacn):
        self.destroy_opened_dropdown_window(datacn=datacn)
        if datacn in self.cell_options and "dropdown" in self.cell_options[datacn]:
            del self.cell_options[datacn]["dropdown"]

    def delete_cell_options_checkbox(self, datacn):
        if datacn in self.cell_options and "checkbox" in self.cell_options[datacn]:
            del self.cell_options[datacn]["checkbox"]

    def delete_cell_options_dropdown_and_checkbox(self, datacn):
        self.delete_cell_options_dropdown(datacn)
        self.delete_cell_options_checkbox(datacn)
