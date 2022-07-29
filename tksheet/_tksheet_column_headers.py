from ._tksheet_vars import *
from ._tksheet_other_classes import *

from itertools import islice, repeat, accumulate, chain, cycle
from math import floor, ceil
import bisect
import pickle
import tkinter as tk
import zlib


class ColumnHeaders(tk.Canvas):
    def __init__(self,
                 parentframe = None,
                 main_canvas = None,
                 row_index_canvas = None,
                 max_colwidth = None,
                 max_header_height = None,
                 default_header = None,
                 header_align = None,
                 header_bg = None,
                 header_border_fg = None,
                 header_grid_fg = None,
                 header_fg = None,
                 header_selected_cells_bg = None,
                 header_selected_cells_fg = None,
                 header_selected_columns_bg = "#5f6368",
                 header_selected_columns_fg = "white",
                 header_select_bold = True,
                 drag_and_drop_bg = None,
                 header_hidden_columns_expander_bg = None,
                 column_drag_and_drop_perform = True,
                 resizing_line_fg = None,
                 show_default_header_for_empty = True):
        tk.Canvas.__init__(self,parentframe,
                           background = header_bg,
                           highlightthickness = 0)

        self.extra_begin_edit_cell_func = None
        self.extra_end_edit_cell_func = None
        
        self.text_editor = None
        self.text_editor_id = None
        self.text_editor_loc = None
        
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
        
        self.b1_pressed_loc = None
        self.existing_dropdown_canvas_id = None
        self.existing_dropdown_window = None
        self.closed_dropdown = None

        self.centre_alignment_text_mod_indexes = (slice(1, None), slice(None, -1))
        self.c_align_cyc = cycle(self.centre_alignment_text_mod_indexes)
        self.parentframe = parentframe
        self.column_drag_and_drop_perform = column_drag_and_drop_perform
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
        self.drag_selection_binding_func = None
        self.column_width_resize_func = None
        self.default_hdr = default_header.lower()
        self.max_cw = float(max_colwidth)
        self.max_header_height = float(max_header_height)
        self.current_height = None    # is set from within MainTable() __init__ or from Sheet parameters
        self.MT = main_canvas         # is set from within MainTable() __init__
        self.RI = row_index_canvas    # is set from within MainTable() __init__
        self.TL = None                # is set from within TopLeftRectangle() __init__
        self.header_bg = header_bg
        self.header_fg = header_fg
        self.header_grid_fg = header_grid_fg
        self.header_border_fg = header_border_fg
        self.header_selected_cells_bg = header_selected_cells_bg
        self.header_selected_cells_fg = header_selected_cells_fg
        self.header_selected_columns_bg = header_selected_columns_bg
        self.header_selected_columns_fg = header_selected_columns_fg
        self.header_hidden_columns_expander_bg = header_hidden_columns_expander_bg
        self.select_bold = header_select_bold
        self.drag_and_drop_bg = drag_and_drop_bg
        self.resizing_line_fg = resizing_line_fg
        self.align = header_align
        self.width_resizing_enabled = False
        self.height_resizing_enabled = False
        self.double_click_resizing_enabled = False
        self.col_selection_enabled = False
        self.drag_and_drop_enabled = False
        self.rc_delete_col_enabled = False
        self.rc_insert_col_enabled = False
        self.hide_columns_enabled = False
        self.edit_cell_enabled = False
        self.show_default_header_for_empty = show_default_header_for_empty
        self.dragged_col = None
        self.visible_col_dividers = []
        self.col_height_resize_bbox = tuple()
        self.cell_options = {}
        self.rsz_w = None
        self.rsz_h = None
        self.new_col_height = 0
        self.currently_resizing_width = False
        self.currently_resizing_height = False
        self.ch_rc_popup_menu = None
        self.basic_bindings()
        
    def basic_bindings(self, enable = True):
        if enable:
            self.bind("<Motion>", self.mouse_motion)
            self.bind("<ButtonPress-1>", self.b1_press)
            self.bind("<B1-Motion>", self.b1_motion)
            self.bind("<ButtonRelease-1>", self.b1_release)
            self.bind("<Double-Button-1>", self.double_b1)
            self.bind(get_rc_binding(), self.rc)
        else:
            self.unbind("<Motion>")
            self.unbind("<ButtonPress-1>")
            self.unbind("<B1-Motion>")
            self.unbind("<ButtonRelease-1>")
            self.unbind("<Double-Button-1>")
            self.unbind(get_rc_binding())

    def set_height(self, new_height, set_TL = False):
        self.current_height = new_height
        try:
            self.config(height = new_height)
        except:
            return
        if set_TL:
            self.TL.set_dimensions(new_h = new_height)

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

    def check_mouse_position_width_resizers(self, event):
        x = self.canvasx(event.x)
        y = self.canvasy(event.y)
        ov = None
        for x1, y1, x2, y2 in self.visible_col_dividers:
            if x >= x1 and y >= y1 and x <= x2 and y <= y2:
                ov = self.find_overlapping(x1, y1, x2, y2)
                break
        return ov

    def rc(self, event):
        self.MT.hide_dropdown_window()
        self.hide_dropdown_window()
        self.focus_set()
        popup_menu = None
        if self.MT.identify_col(x = event.x, allow_end = False) is None:
            self.MT.deselect("all")
            popup_menu = self.ch_rc_popup_menu
        elif self.col_selection_enabled and not self.currently_resizing_width and not self.currently_resizing_height:
            c = self.MT.identify_col(x = event.x)
            if c < len(self.MT.col_positions) - 1:
                if self.MT.col_selected(c):
                    if self.MT.rc_popup_menus_enabled:
                        popup_menu = self.ch_rc_popup_menu
                else:
                    if self.MT.single_selection_enabled and self.MT.rc_select_enabled:
                        self.select_col(c, redraw = True)
                    elif self.MT.toggle_selection_enabled and self.MT.rc_select_enabled:
                        self.toggle_select_col(c, redraw = True)
                    if self.MT.rc_popup_menus_enabled:
                        popup_menu = self.ch_rc_popup_menu
        if self.extra_rc_func is not None:
            self.extra_rc_func(event)
        if popup_menu is not None:
            popup_menu.tk_popup(event.x_root, event.y_root)

    def shift_b1_press(self, event):
        self.MT.hide_dropdown_window()
        x = event.x
        c = self.MT.identify_col(x = x)
        if self.drag_and_drop_enabled or self.col_selection_enabled and self.rsz_h is None and self.rsz_w is None:
            if c < len(self.MT.col_positions) - 1:
                c_selected = self.MT.col_selected(c)
                if not c_selected and self.col_selection_enabled:
                    c = int(c)
                    currently_selected = self.MT.currently_selected()
                    if currently_selected and currently_selected[0] == "column":
                        min_c = int(currently_selected[1])
                        self.MT.delete_selection_rects(delete_current = False)
                        if c > min_c:
                            self.MT.create_selected(0, min_c, len(self.MT.row_positions) - 1, c + 1, "cols")
                            func_event = tuple(range(min_c, c + 1))
                        elif c < min_c:
                            self.MT.create_selected(0, c, len(self.MT.row_positions) - 1, min_c + 1, "cols")
                            func_event = tuple(range(c, min_c + 1))
                    else:
                        self.select_col(c)
                    self.MT.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
                    if self.shift_selection_binding_func is not None:
                        self.shift_selection_binding_func(SelectionBoxEvent("shift_select_columns", func_event))
                elif c_selected:
                    self.dragged_col = c

    def create_resize_line(self, x1, y1, x2, y2, width, fill, tag):
        if self.hidd_resize_lines:
            t, sh = self.hidd_resize_lines.popitem()
            self.coords(t, x1, y1, x2, y2)
            if sh:
                self.itemconfig(t, width = width, fill = fill, tag = tag)
            else:
                self.itemconfig(t, width = width, fill = fill, tag = tag, state = "normal")
            self.lift(t)
        else:
            t = self.create_line(x1, y1, x2, y2, width = width, fill = fill, tag = tag)
        self.disp_resize_lines[t] = True

    def delete_resize_lines(self):
        self.hidd_resize_lines.update(self.disp_resize_lines)
        self.disp_resize_lines = {}
        for t, sh in self.hidd_resize_lines.items():
            if sh:
                self.itemconfig(t, state = "hidden")
                self.hidd_resize_lines[t] = False

    def mouse_motion(self, event):
        if not self.currently_resizing_height and not self.currently_resizing_width:
            x = self.canvasx(event.x)
            y = self.canvasy(event.y)
            mouse_over_resize = False
            if self.width_resizing_enabled:
                ov = self.check_mouse_position_width_resizers(event)
                if ov is not None:
                    tgs = tuple()
                    for itm in ov:
                        tgs = self.gettags(itm)
                        if tgs and "v" == tgs[0]:
                            c = int(tgs[1])
                            self.rsz_w, mouse_over_resize = c, True
                            self.config(cursor = "sb_h_double_arrow")
                            break
                else:
                    self.rsz_w = None
            if self.height_resizing_enabled and not mouse_over_resize:
                try:
                    x1, y1, x2, y2 = self.col_height_resize_bbox[0], self.col_height_resize_bbox[1], self.col_height_resize_bbox[2], self.col_height_resize_bbox[3]
                    if x >= x1 and y >= y1 and x <= x2 and y <= y2:
                        self.config(cursor = "sb_v_double_arrow")
                        self.rsz_h = True
                        mouse_over_resize = True
                    else:
                        self.rsz_h = None
                except:
                    self.rsz_h = None
            if not mouse_over_resize:
                self.MT.reset_mouse_motion_creations()
        if self.extra_motion_func is not None:
            self.extra_motion_func(event)

    def double_b1(self, event = None):
        self.MT.hide_dropdown_window()
        self.focus_set()
        if self.double_click_resizing_enabled and self.width_resizing_enabled and self.rsz_w is not None and not self.currently_resizing_width:
            col = self.rsz_w - 1
            old_width = self.MT.col_positions[self.rsz_w] - self.MT.col_positions[self.rsz_w - 1]
            new_width = self.set_col_width(col)
            self.MT.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
            if self.column_width_resize_func is not None and old_width != new_width:
                self.column_width_resize_func(ResizeEvent("column_width_resize", col, old_width, new_width))
        elif self.col_selection_enabled and self.rsz_h is None and self.rsz_w is None:
            c = self.MT.identify_col(x = event.x)
            if c < len(self.MT.col_positions) - 1:
                if self.MT.single_selection_enabled:
                    self.select_col(c, redraw = True)
                elif self.MT.toggle_selection_enabled:
                    self.toggle_select_col(c, redraw = True)
            dcol = c if self.MT.all_columns_displayed else self.MT.displayed_columns[c]
            if ((dcol in self.cell_options and 'checkbox' in self.cell_options[dcol]) or
                (dcol in self.cell_options and 'dropdown' in self.cell_options[dcol]) or
                self.edit_cell_enabled):
                self.open_cell(event)
        self.rsz_w = None
        self.mouse_motion(event)
        if self.extra_double_b1_func is not None:
            self.extra_double_b1_func(event)
        
    def b1_press(self, event = None):
        self.MT.hide_dropdown_window(b1 = True)
        self.focus_set()
        self.MT.unbind("<MouseWheel>")
        self.closed_dropdown = self.hide_dropdown_window(b1 = True)
        x1, y1, x2, y2 = self.MT.get_canvas_visible_area()
        if self.check_mouse_position_width_resizers(event) is None:
            self.rsz_w = None
        c = self.MT.identify_col(x = event.x)
        self.b1_pressed_loc = c
        if self.width_resizing_enabled and self.rsz_w is not None:
            self.currently_resizing_width = True
            x = self.MT.col_positions[self.rsz_w]
            line2x = self.MT.col_positions[self.rsz_w - 1]
            self.create_resize_line(x, 0, x, self.current_height, width = 1, fill = self.resizing_line_fg, tag = "rwl")
            self.MT.create_resize_line(x, y1, x, y2, width = 1, fill = self.resizing_line_fg, tag = "rwl")
            self.create_resize_line(line2x, 0, line2x, self.current_height,width = 1, fill = self.resizing_line_fg, tag = "rwl2")
            self.MT.create_resize_line(line2x, y1, line2x, y2, width = 1, fill = self.resizing_line_fg, tag = "rwl2")
        elif self.height_resizing_enabled and self.rsz_w is None and self.rsz_h is not None:
            self.currently_resizing_height = True
            y = event.y
            if y < self.MT.hdr_min_rh:
                y = int(self.MT.hdr_min_rh)
            self.new_col_height = y
            self.create_resize_line(x1, y, x2, y, width = 1, fill = self.resizing_line_fg, tag = "rhl")
        elif self.MT.identify_col(x = event.x, allow_end = False) is None:
            self.MT.deselect("all")
        elif self.col_selection_enabled and self.rsz_w is None and self.rsz_h is None:
            if c < len(self.MT.col_positions) - 1:
                if self.MT.single_selection_enabled:
                    self.select_col(c, redraw = True)
                elif self.MT.toggle_selection_enabled:
                    self.toggle_select_col(c, redraw = True)
        if self.extra_b1_press_func is not None:
            self.extra_b1_press_func(event)
    
    def b1_motion(self, event):
        x1, y1, x2, y2 = self.MT.get_canvas_visible_area()
        if self.width_resizing_enabled and self.rsz_w is not None and self.currently_resizing_width:
            x = self.canvasx(event.x)
            size = x - self.MT.col_positions[self.rsz_w - 1]
            if not size <= self.MT.min_cw and size < self.max_cw:
                self.delete_resize_lines()
                self.MT.delete_resize_lines()
                line2x = self.MT.col_positions[self.rsz_w - 1]
                self.create_resize_line(x, 0, x, self.current_height, width = 1, fill = self.resizing_line_fg, tag = "rwl")
                self.MT.create_resize_line(x, y1, x, y2, width = 1, fill = self.resizing_line_fg, tag = "rwl")
                self.create_resize_line(line2x, 0, line2x, self.current_height,width = 1, fill = self.resizing_line_fg, tag = "rwl2")
                self.MT.create_resize_line(line2x, y1, line2x, y2, width = 1, fill = self.resizing_line_fg, tag = "rwl2")
        elif self.height_resizing_enabled and self.rsz_h is not None and self.currently_resizing_height:
            evy = event.y
            self.delete_resize_lines()
            self.MT.delete_resize_lines()
            if evy > self.current_height:
                y = self.MT.canvasy(evy - self.current_height)
                if evy > self.max_header_height:
                    evy = int(self.max_header_height)
                    y = self.MT.canvasy(evy - self.current_height)
                self.new_col_height = evy
                self.MT.create_resize_line(x1, y, x2, y, width = 1, fill = self.resizing_line_fg, tag = "rhl")
            else:
                y = evy
                if y < self.MT.hdr_min_rh:
                    y = int(self.MT.hdr_min_rh)
                self.new_col_height = y
                self.create_resize_line(x1, y, x2, y, width = 1, fill = self.resizing_line_fg, tag = "rhl")
        elif self.drag_and_drop_enabled and self.col_selection_enabled and self.MT.anything_selected(exclude_cells = True, exclude_rows = True) and self.rsz_h is None and self.rsz_w is None and self.dragged_col is not None:
            x = self.canvasx(event.x)
            if x > 0 and x < self.MT.col_positions[-1]:
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
                    self.check_xview()
                    self.MT.main_table_redraw_grid_and_text(redraw_header = True)
                elif x <= 0 and len(xcheck) > 1 and xcheck[0] > 0:
                    if x >= -15:
                        self.MT.xview_scroll(-1, "units")
                        self.xview_scroll(-1, "units")
                    else:
                        self.MT.xview_scroll(-2, "units")
                        self.xview_scroll(-2, "units")
                    self.check_xview()
                    self.MT.main_table_redraw_grid_and_text(redraw_header = True)
                col = self.MT.identify_col(x = event.x)
                sels = self.MT.get_selected_cols()
                selsmin = min(sels)
                if col in sels:
                    xpos = self.MT.col_positions[selsmin]
                else:
                    if col < selsmin:
                        xpos = self.MT.col_positions[col]
                    else:
                        xpos = self.MT.col_positions[col + 1]
                self.delete_resize_lines()
                self.MT.delete_resize_lines()
                self.create_resize_line(xpos, 0, xpos, self.current_height, width = 3, fill = self.drag_and_drop_bg, tag = "dd")
                self.MT.create_resize_line(xpos, y1, xpos, y2, width = 3, fill = self.drag_and_drop_bg, tag = "dd")
        elif self.MT.drag_selection_enabled and self.col_selection_enabled and self.rsz_h is None and self.rsz_w is None:
            end_col = self.MT.identify_col(x = event.x)
            currently_selected = self.MT.currently_selected()
            if end_col < len(self.MT.col_positions) - 1 and currently_selected:
                if currently_selected[0] == "column":
                    start_col = currently_selected[1]
                    if end_col >= start_col:
                        rect = (0, start_col, len(self.MT.row_positions) - 1, end_col + 1, "cols")
                        func_event = tuple(range(start_col, end_col + 1))
                    elif end_col < start_col:
                        rect = (0, end_col, len(self.MT.row_positions) - 1, start_col + 1, "cols")
                        func_event = tuple(range(end_col, start_col + 1))
                    if self.being_drawn_rect != rect:
                        self.MT.delete_selection_rects(delete_current = False)
                        self.MT.create_selected(*rect)
                        self.being_drawn_rect = rect
                        if self.drag_selection_binding_func is not None:
                            self.drag_selection_binding_func(SelectionBoxEvent("drag_select_columns", func_event))
                xcheck = self.xview()
                if event.x > self.winfo_width() and len(xcheck) > 1 and xcheck[1] < 1:
                    try:
                        self.MT.xview_scroll(1, "units")
                        self.xview_scroll(1, "units")
                    except:
                        pass
                    self.check_xview()
                elif event.x < 0 and self.canvasx(self.winfo_width()) > 0 and xcheck and xcheck[0] > 0:
                    try:
                        self.xview_scroll(-1, "units")
                        self.MT.xview_scroll(-1, "units")
                    except:
                        pass
                    self.check_xview()
            self.MT.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = False)
        if self.extra_b1_motion_func is not None:
            self.extra_b1_motion_func(event)

    def check_xview(self):
        xcheck = self.xview()
        if xcheck and xcheck[0] < 0:
            self.MT.set_xviews("moveto", 0)
        elif len(xcheck) > 1 and xcheck[1] > 1:
            self.MT.set_xviews("moveto", 1)
            
    def b1_release(self, event = None):
        self.MT.bind("<MouseWheel>", self.MT.mousewheel)
        if self.width_resizing_enabled and self.rsz_w is not None and self.currently_resizing_width:
            self.currently_resizing_width = False
            new_col_pos = self.coords("rwl")[0]
            self.delete_resize_lines()
            self.MT.delete_resize_lines()
            old_width = self.MT.col_positions[self.rsz_w] - self.MT.col_positions[self.rsz_w - 1]
            size = new_col_pos - self.MT.col_positions[self.rsz_w - 1]
            if size < self.MT.min_cw:
                new_col_pos = ceil(self.MT.col_positions[self.rsz_w - 1] + self.MT.min_cw)
            elif size > self.max_cw:
                new_col_pos = floor(self.MT.col_positions[self.rsz_w - 1] + self.max_cw)
            increment = new_col_pos - self.MT.col_positions[self.rsz_w]
            self.MT.col_positions[self.rsz_w + 1:] = [e + increment for e in islice(self.MT.col_positions, self.rsz_w + 1, len(self.MT.col_positions))]
            self.MT.col_positions[self.rsz_w] = new_col_pos
            new_width = self.MT.col_positions[self.rsz_w] - self.MT.col_positions[self.rsz_w - 1]
            self.MT.recreate_all_selection_boxes()
            self.MT.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
            if self.column_width_resize_func is not None and old_width != new_width:
                self.column_width_resize_func(ResizeEvent("column_width_resize", self.rsz_w - 1, old_width, new_width))
        elif self.height_resizing_enabled and self.rsz_h is not None and self.currently_resizing_height:
            self.currently_resizing_height = False
            self.delete_resize_lines()
            self.MT.delete_resize_lines()
            self.set_height(self.new_col_height,set_TL = True)
            self.MT.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
        elif self.drag_and_drop_enabled and self.col_selection_enabled and self.MT.anything_selected(exclude_cells = True, exclude_rows = True) and self.rsz_h is None and self.rsz_w is None and self.dragged_col is not None:
            self.delete_resize_lines()
            self.MT.delete_resize_lines()
            x = event.x
            c = self.MT.identify_col(x = x)
            orig_selected_cols = self.MT.get_selected_cols()
            if c != self.dragged_col and c is not None and c not in orig_selected_cols and len(orig_selected_cols) != len(self.MT.col_positions) - 1:
                orig_selected_cols = sorted(orig_selected_cols)
                if len(orig_selected_cols) > 1:
                    start_idx = bisect.bisect_left(orig_selected_cols, self.dragged_col)
                    forward_gap = get_index_of_gap_in_sorted_integer_seq_forward(orig_selected_cols, start_idx)
                    reverse_gap = get_index_of_gap_in_sorted_integer_seq_reverse(orig_selected_cols, start_idx)
                    if forward_gap is not None:
                        orig_selected_cols[:] = orig_selected_cols[:forward_gap]
                    if reverse_gap is not None:
                        orig_selected_cols[:] = orig_selected_cols[reverse_gap:]
                rm1start = orig_selected_cols[0]
                totalcols = len(orig_selected_cols)
                extra_func_success = True
                if c >= len(self.MT.col_positions) - 1:
                    c -= 1
                if self.ch_extra_begin_drag_drop_func is not None:
                    try:
                        self.ch_extra_begin_drag_drop_func(BeginDragDropEvent("begin_column_header_drag_drop", tuple(orig_selected_cols), int(c)))
                    except:
                        extra_func_success = False
                if extra_func_success:
                    new_selected, dispset = self.MT.move_columns_adjust_options_dict(c,
                                                                                     rm1start, 
                                                                                     totalcols,
                                                                                     move_data = self.column_drag_and_drop_perform)
                    if self.MT.undo_enabled:
                        self.MT.undo_storage.append(zlib.compress(pickle.dumps(("move_cols",                 #0
                                                                                int(orig_selected_cols[0]),  #1
                                                                                int(new_selected[0]),        #2
                                                                                int(new_selected[-1]),       #3
                                                                                sorted(orig_selected_cols),  #4
                                                                                dispset))))                  #5
                    self.MT.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
                    if self.ch_extra_end_drag_drop_func is not None:
                        self.ch_extra_end_drag_drop_func(EndDragDropEvent("end_column_header_drag_drop", tuple(orig_selected_cols), new_selected, int(c)))
        elif self.b1_pressed_loc is not None and self.rsz_w is None and self.rsz_h is None:
            c = self.MT.identify_col(x = event.x)
            if c is not None and c == self.b1_pressed_loc and self.b1_pressed_loc != self.closed_dropdown:
                dcol = c if self.MT.all_columns_displayed else self.MT.displayed_columns[c]
                if ((dcol in self.cell_options and 'dropdown' in self.cell_options[dcol] and event.x < self.MT.col_positions[c + 1] and event.x > self.MT.col_positions[c + 1] - self.MT.hdr_txt_h - 4) or
                    (dcol in self.cell_options and 'checkbox' in self.cell_options[dcol] and event.x < self.MT.col_positions[c] + self.MT.hdr_txt_h + 5)):
                    if event.y < self.MT.hdr_txt_h + 5:
                        self.open_cell(event)
            else:
                self.hide_dropdown_window()
            self.b1_pressed_loc = None
            self.closed_dropdown = None

        self.dragged_col = None
        self.currently_resizing_width = False
        self.currently_resizing_height = False
        self.rsz_w = None
        self.rsz_h = None
        self.being_drawn_rect = None
        self.mouse_motion(event)
        if self.extra_b1_release_func is not None:
            self.extra_b1_release_func(event)
            
    def readonly_header(self, columns = [], readonly = True):
        if isinstance(columns, int):
            columns_ = [columns]
        else:
            columns_ = columns
        if not readonly:
            for c in columns_:
                if c in self.cell_options and 'readonly' in self.cell_options[c]:
                    del self.cell_options[c]['readonly']
        else:
            for c in columns_:
                if c not in self.cell_options:
                    self.cell_options[c] = {}
                self.cell_options[c]['readonly'] = True

    def highlight_cells(self, c = 0, cells = tuple(), bg = None, fg = None, redraw = False, overwrite = True):
        if bg is None and fg is None:
            return
        if cells and not isinstance(cells, int):
            iterable = cells
        elif isinstance(cells, int):
            iterable = (cells, )
        else:
            iterable = (c, )
        for c_ in iterable:
            if c_ not in self.cell_options:
                self.cell_options[c_] = {}
            if 'highlight' in self.cell_options[c_] and not overwrite:
                self.cell_options[c_]['highlight'] = (self.cell_options[c_]['highlight'][0] if bg is None else bg,
                                                      self.cell_options[c_]['highlight'][1] if fg is None else fg)
            else:
                self.cell_options[c_]['highlight'] = (bg, fg)
        if redraw:
            self.MT.main_table_redraw_grid_and_text(True, False)

    def select_col(self, c, redraw = False, keep_other_selections = False):
        ignore_keep = False
        if keep_other_selections:
            if self.MT.col_selected(c):
                self.MT.create_current(0, c, type_ = "col", inside = True)
            else:
                ignore_keep = True
        if ignore_keep or not keep_other_selections:
            self.MT.delete_selection_rects()
            self.MT.create_current(0, c, type_ = "col", inside = True)
            self.MT.create_selected(0, c, len(self.MT.row_positions) - 1, c + 1, "cols")
        if redraw:
            self.MT.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
        if self.selection_binding_func is not None:
            self.selection_binding_func(SelectColumnEvent("select_column", int(c)))

    def toggle_select_col(self, column, add_selection = True, redraw = True, run_binding_func = True, set_as_current = True):
        if add_selection:
            if self.MT.col_selected(column):
                self.MT.deselect(c = column, redraw = redraw)
            else:
                self.add_selection(c = column, redraw = redraw, run_binding_func = run_binding_func, set_as_current = set_as_current)
        else:
            if self.MT.col_selected(column):
                self.MT.deselect(c = column, redraw = redraw)
            else:
                self.select_col(column, redraw = redraw)

    def add_selection(self, c, redraw = False, run_binding_func = True, set_as_current = True):
        c = int(c)
        if set_as_current:
            create_new_sel = False
            current = self.MT.get_tags_of_current()
            if current:
                if current[0] == "Current_Outside":
                    create_new_sel = True
            self.MT.create_current(0, c, type_ = "col", inside = True)
            if create_new_sel:
                r1, c1, r2, c2 = tuple(int(e) for e in current[1].split("_") if e)
                self.MT.create_selected(r1, c1, r2, c2, current[2] + "s")
        self.MT.create_selected(0, c, len(self.MT.row_positions) - 1, c + 1, "cols")
        if redraw:
            self.MT.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
        if self.selection_binding_func is not None and run_binding_func:
            self.selection_binding_func(("select_column", int(c)))

    def set_col_width(self, col, width = None, only_set_if_too_small = False, displayed_only = False, recreate = True, return_new_width = False):
        if col < 0:
            return
        qconf = self.MT.txt_measure_canvas.itemconfig
        qbbox = self.MT.txt_measure_canvas.bbox
        qtxtm = self.MT.txt_measure_canvas_text
        qtxth = self.MT.txt_h
        qclop = self.MT.cell_options
        qfont = self.MT.my_font
        if width is None:
            w = self.MT.min_cw
            hw = self.MT.min_cw
            if displayed_only:
                x1, y1, x2, y2 = self.MT.get_canvas_visible_area()
                start_row, end_row = self.MT.get_visible_rows(y1, y2)
            else:
                start_row, end_row = 0, None
            if self.MT.all_columns_displayed:
                data_col = col
            else:
                data_col = self.MT.displayed_columns[col]
            # header
            if data_col in self.cell_options and 'checkbox' in self.cell_options[data_col]:
                qconf(qtxtm, text = self.cell_options[data_col]['checkbox']['text'], font = self.MT.my_hdr_font)
                b = qbbox(qtxtm)
                hw = b[2] - b[0] + 7 + self.MT.hdr_txt_h
            else:
                txt = ""
                if isinstance(self.MT.my_hdrs, int) or len(self.MT.my_hdrs) - 1 >= data_col:
                    if isinstance(self.MT.my_hdrs, int):
                        txt = self.MT.data_ref[self.MT.my_hdrs][data_col]
                    else:
                        txt = self.MT.my_hdrs[data_col]
                    if txt or (data_col in self.cell_options and 'dropdown' in self.cell_options[data_col]):
                        qconf(qtxtm, text = txt, font = self.MT.my_hdr_font)
                        b = qbbox(qtxtm)
                        if data_col in self.cell_options and 'dropdown' in self.cell_options[data_col]:
                            hw = b[2] - b[0] + 7 + self.MT.hdr_txt_h
                        else:
                            hw = b[2] - b[0] + 7
                if not isinstance(self.MT.my_hdrs, int) and ((not txt and self.show_default_header_for_empty) or len(self.MT.my_hdrs) < data_col):
                    if self.default_hdr == "letters":
                        hw = self.MT.GetHdrTextWidth(num2alpha(data_col)) + 7
                    elif self.default_hdr == "numbers":
                        hw = self.MT.GetHdrTextWidth(f"{data_col + 1}") + 7
                    else:
                        hw = self.MT.GetHdrTextWidth(f"{data_col + 1} {num2alpha(data_col)}") + 7
            # table
            for rn, r in enumerate(islice(self.MT.data_ref, start_row, end_row), start_row):
                if (rn, data_col) in qclop and 'checkbox' in qclop[(rn, data_col)]:
                    qconf(qtxtm, text = qclop[(rn, data_col)]['checkbox']['text'], font = qfont)
                    b = qbbox(qtxtm)
                    tw = qtxth + 7 + b[2] - b[0]
                    if tw > w:
                        w = tw
                else:
                    try:
                        if isinstance(r[data_col], str):
                            txt = r[data_col]
                        else:
                            txt = f"{r[data_col]}"
                    except:
                        continue
                    if txt:
                        qconf(qtxtm, text = txt, font = qfont)
                        b = qbbox(qtxtm)
                        tw = b[2] - b[0] + qtxth + 7 if (rn, data_col) in qclop and 'dropdown' in qclop[(rn, data_col)] else b[2] - b[0] + 7
                        if tw > w:
                            w = tw
            if w > hw:
                new_width = w
            else:
                new_width = hw
        else:
            new_width = int(width)
        if new_width <= self.MT.min_cw:
            new_width = int(self.MT.min_cw)
        elif new_width > self.max_cw:
            new_width = int(self.max_cw)
        if only_set_if_too_small:
            if new_width <= self.MT.col_positions[col + 1] - self.MT.col_positions[col]:
                return self.MT.col_positions[col + 1] - self.MT.col_positions[col]
        if return_new_width:
            return new_width
        else:
            new_col_pos = self.MT.col_positions[col] + new_width
            increment = new_col_pos - self.MT.col_positions[col + 1]
            self.MT.col_positions[col + 2:] = [e + increment for e in islice(self.MT.col_positions, col + 2, len(self.MT.col_positions))]
            self.MT.col_positions[col + 1] = new_col_pos
            if recreate:
                self.MT.recreate_all_selection_boxes()
        return new_width

    def set_width_of_all_cols(self, width = None, only_set_if_too_small = False, recreate = True):
        if width is None:
            if self.MT.all_columns_displayed:
                iterable = range(self.MT.total_data_cols())
            else:
                iterable = range(len(self.MT.displayed_columns))
            self.MT.col_positions = list(accumulate(chain([0], (self.set_col_width(cn, only_set_if_too_small = only_set_if_too_small, recreate = False, return_new_width = True) for cn in iterable))))
        elif width is not None:
            if self.MT.all_columns_displayed:
                self.MT.col_positions = list(accumulate(chain([0], (width for cn in range(self.MT.total_data_cols())))))
            else:
                self.MT.col_positions = list(accumulate(chain([0], (width for cn in range(len(self.MT.displayed_columns))))))
        if recreate:
            self.MT.recreate_all_selection_boxes()

    def GetLargestWidth(self, cell):
        return max(cell.split("\n"), key = self.MT.GetTextWidth)

    def align_cells(self, columns = [], align = "global"):
        if isinstance(columns, int):
            cols = [columns]
        else:
            cols = columns
        if align == "global":
            for c in cols:
                if c in self.cell_options and 'align' in self.cell_options[c]:
                    del self.cell_options[c]['align']
        else:
            for c in cols:
                if c not in self.cell_options:
                    self.cell_options[c] = {}
                self.cell_options[c]['align'] = align

    def redraw_highlight_get_text_fg(self, fc, sc, c, c_2, c_3, selected_cols, selected_rows, actual_selected_cols, hlcol):
        redrawn = False
        if hlcol in self.cell_options and 'highlight' in self.cell_options[hlcol] and c in actual_selected_cols:
            if self.cell_options[hlcol]['highlight'][0] is not None:
                c_1 = self.cell_options[hlcol]['highlight'][0] if self.cell_options[hlcol]['highlight'][0].startswith("#") else Color_Map_[self.cell_options[hlcol]['highlight'][0]]
                redrawn = self.redraw_highlight(fc + 1,
                                                0,
                                                sc,
                                                self.current_height - 1,
                                                fill = (f"#{int((int(c_1[1:3], 16) + int(c_3[1:3], 16)) / 2):02X}" +
                                                        f"{int((int(c_1[3:5], 16) + int(c_3[3:5], 16)) / 2):02X}" +
                                                        f"{int((int(c_1[5:], 16) + int(c_3[5:], 16)) / 2):02X}"),
                                                outline = self.header_fg if hlcol in self.cell_options and 'dropdown' in self.cell_options[hlcol] else "", 
                                                tag = "hi")
            tf = self.header_selected_columns_fg if self.cell_options[hlcol]['highlight'][1] is None or self.MT.display_selected_fg_over_highlights else self.cell_options[hlcol]['highlight'][1]
        elif hlcol in self.cell_options and 'highlight' in self.cell_options[hlcol] and (c in selected_cols or selected_rows):
            if self.cell_options[hlcol]['highlight'][0] is not None:
                c_1 = self.cell_options[hlcol]['highlight'][0] if self.cell_options[hlcol]['highlight'][0].startswith("#") else Color_Map_[self.cell_options[hlcol]['highlight'][0]]
                redrawn = self.redraw_highlight(fc + 1,
                                                0,
                                                sc,
                                                self.current_height - 1,
                                                fill = (f"#{int((int(c_1[1:3], 16) + int(c_2[1:3], 16)) / 2):02X}" +
                                                        f"{int((int(c_1[3:5], 16) + int(c_2[3:5], 16)) / 2):02X}" +
                                                        f"{int((int(c_1[5:], 16) + int(c_2[5:], 16)) / 2):02X}"),
                                                outline = self.header_fg if hlcol in self.cell_options and 'dropdown' in self.cell_options[hlcol] else "", 
                                                tag = "hi")
            tf = self.header_selected_cells_fg if self.cell_options[hlcol]['highlight'][1] is None or self.MT.display_selected_fg_over_highlights else self.cell_options[hlcol]['highlight'][1]
        elif c in actual_selected_cols:
            tf = self.header_selected_columns_fg
        elif c in selected_cols or selected_rows:
            tf = self.header_selected_cells_fg
        elif hlcol in self.cell_options and 'highlight' in self.cell_options[hlcol]:
            if self.cell_options[hlcol]['highlight'][0] is not None:
                redrawn = self.redraw_highlight(fc + 1,
                                                0, 
                                                sc, 
                                                self.current_height - 1, 
                                                fill = self.cell_options[hlcol]['highlight'][0], 
                                                outline = self.header_fg if hlcol in self.cell_options and 'dropdown' in self.cell_options[hlcol] else "", 
                                                tag = "hi")
            tf = self.header_fg if self.cell_options[hlcol]['highlight'][1] is None else self.cell_options[hlcol]['highlight'][1]
        else:
            tf = self.header_fg
        return tf, redrawn
            
    def redraw_highlight(self, x1, y1, x2, y2, fill, outline, tag, can_width = None):
        if self.hidd_high:
            t, sh = self.hidd_high.popitem()
            self.coords(t, x1 - 1 if outline else x1, y1 - 1 if outline else y1, x2 if can_width is None else x2 + can_width, y2)
            if sh:
                self.itemconfig(t, fill = fill, outline = outline)
            else:
                self.itemconfig(t, fill = fill, outline = outline, tag = tag, state = "normal")
            self.lift(t)
        else:
            t = self.create_rectangle(x1 - 1 if outline else x1, y1 - 1 if outline else y1, x2 if can_width is None else x2 + can_width, y2, fill = fill, outline = outline, tag = tag)
        self.disp_high[t] = True
        return True

    def redraw_text(self, x, y, text, fill, font, anchor, tag):
        if self.hidd_text:
            t, sh = self.hidd_text.popitem()
            self.coords(t, x, y)
            if sh:
                self.itemconfig(t, text = text, fill = fill, font = font, anchor = anchor)
            else:
                self.itemconfig(t, text = text, fill = fill, font = font, anchor = anchor, state = "normal")
            self.lift(t)
        else:
            t = self.create_text(x, y, text = text, fill = fill, font = font, anchor = anchor, tag = tag)
        self.disp_text[t] = True
        return t

    def redraw_gridline(self, x1, y1, x2, y2, fill, width, tag):
        if self.hidd_grid:
            t, sh = self.hidd_grid.popitem()
            self.coords(t, x1, y1, x2, y2)
            if sh:
                self.itemconfig(t, fill = fill, width = width, tag = tag, capstyle = tk.BUTT, joinstyle = tk.ROUND)
            else:
                self.itemconfig(t, fill = fill, width = width, tag = tag, capstyle = tk.BUTT, joinstyle = tk.ROUND, state = "normal")
            self.disp_grid[t] = True
        else:
            self.disp_grid[self.create_line(x1, y1, x2, y2, fill = fill, width = width, tag = tag)] = True
            
    def redraw_dropdown(self, x1, y1, x2, y2, fill, outline, tag, draw_outline = True, draw_arrow = True):
        if draw_outline:
            self.redraw_highlight(x1 + 1, y1 + 1, x2, y2, fill = "", outline = self.header_fg, tag = tag)
        if draw_arrow:
            topysub = floor(self.MT.hdr_half_txt_h / 2)
            mid_y = y1 + floor(self.MT.hdr_half_txt_h / 2) + 5
            #top left points for triangle
            ty1 = mid_y - topysub + 2
            tx1 = x2 - self.MT.hdr_txt_h + 1
            #bottom points for triangle
            ty2 = mid_y + self.MT.hdr_half_txt_h - 4
            tx2 = x2 - self.MT.hdr_half_txt_h - 1
            #top right points for triangle
            ty3 = mid_y - topysub + 2
            tx3 = x2 - 3
            
            points = (tx1, ty1, tx2, ty2, tx3, ty3)
            if self.hidd_dropdown:
                t, sh = self.hidd_dropdown.popitem()
                self.coords(t, points)
                if sh:
                    self.itemconfig(t, fill = fill)
                else:
                    self.itemconfig(t, fill = fill, tag = tag, state = "normal")
                self.lift(t)
            else:
                t = self.create_line(points, fill = fill, width = 2, capstyle = tk.ROUND, joinstyle = tk.ROUND, tag = tag)
            self.disp_dropdown[t] = True
            
    def redraw_checkbox(self, dcol, x1, y1, x2, y2, fill, outline, tag, draw_check = False):
        points = self.MT.get_checkbox_points(x1, y1, x2, y2)
        if self.hidd_checkbox:
            t, sh = self.hidd_checkbox.popitem()
            self.coords(t, points)
            if sh:
                self.itemconfig(t, fill = outline, outline = fill)
            else:
                self.itemconfig(t, fill = outline, outline = fill, tag = tag, state = "normal")
            self.lift(t)
        else:
            t = self.create_polygon(points, fill = outline, outline = fill, tag = tag, smooth = True)
        self.disp_checkbox[t] = True
        if draw_check:
            # draw filled box
            x1 = x1 + 2
            y1 = y1 + 2
            x2 = x2 - 1
            y2 = y2 - 1
            points = self.MT.get_checkbox_points(x1, y1, x2, y2)
            if self.hidd_checkbox:
                t, sh = self.hidd_checkbox.popitem()
                self.coords(t, points)
                if sh:
                    self.itemconfig(t, fill = fill, outline = outline)
                else:
                    self.itemconfig(t, fill = fill, outline = outline, tag = tag, state = "normal")
                self.lift(t)
            else:
                t = self.create_polygon(points, fill = fill, outline = outline, tag = tag, smooth = True)
            self.disp_checkbox[t] = True

            # draw one line of X
            if self.hidd_grid:
                t, sh = self.hidd_grid.popitem()
                self.coords(t, x1 + 2, y1 + 2, x2 - 2, y2 - 2)
                if sh:
                    self.itemconfig(t, fill = self.get_widget_bg_fg(dcol)[0], capstyle = tk.ROUND, joinstyle = tk.ROUND, width = 2)
                else:
                    self.itemconfig(t, fill = self.get_widget_bg_fg(dcol)[0], capstyle = tk.ROUND, joinstyle = tk.ROUND, width = 2, tag = tag, state = "normal")
                self.lift(t)
            else:
                t = self.create_line(x1 + 2, y1 + 2, x2 - 2, y2 - 2, fill = self.get_widget_bg_fg(dcol)[0], capstyle = tk.ROUND, joinstyle = tk.ROUND, width = 2, tag = tag)
            self.disp_grid[t] = True

            # draw other line of X
            if self.hidd_grid:
                t, sh = self.hidd_grid.popitem()
                self.coords(t, x2 - 2, y1 + 2, x1 + 2, y2 - 2)
                if sh:
                    self.itemconfig(t, fill = self.get_widget_bg_fg(dcol)[0], capstyle = tk.ROUND, joinstyle = tk.ROUND, width = 2)
                else:
                    self.itemconfig(t, fill = self.get_widget_bg_fg(dcol)[0], capstyle = tk.ROUND, joinstyle = tk.ROUND, width = 2, tag = tag, state = "normal")
                self.lift(t)
            else:
                t = self.create_line(x2 - 2, y1 + 2, x1 + 2, y2 - 2, fill = self.get_widget_bg_fg(dcol)[0], capstyle = tk.ROUND, joinstyle = tk.ROUND, width = 2, tag = tag)
            self.disp_grid[t] = True

    def redraw_grid_and_text(self, last_col_line_pos, x1, x_stop, start_col, end_col, selected_cols, selected_rows, actual_selected_cols):
        self.configure(scrollregion = (0,
                                       0,
                                       last_col_line_pos + self.MT.empty_horizontal,
                                       self.current_height))
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
        self.visible_col_dividers = []
        x = self.MT.col_positions[start_col]
        self.redraw_gridline(x, 0, x, self.current_height, fill = self.header_grid_fg, width = 1, tag = "fv")
        self.col_height_resize_bbox = (x1, self.current_height - 2, x_stop, self.current_height)
        yend = self.current_height - 5
        for c in range(start_col + 1, end_col):
            x = self.MT.col_positions[c]
            if self.width_resizing_enabled:
                self.visible_col_dividers.append((x - 2, 1, x + 2, yend))
            self.redraw_gridline(x, 0, x, self.current_height, fill = self.header_grid_fg, width = 1, tag = ("v", f"{c}"))
        top = self.canvasy(0)
        if self.MT.hdr_fl_ins + self.MT.hdr_half_txt_h - 1 > top:
            incfl = True
        else:
            incfl = False
        c_2 = self.header_selected_cells_bg if self.header_selected_cells_bg.startswith("#") else Color_Map_[self.header_selected_cells_bg]
        c_3 = self.header_selected_columns_bg if self.header_selected_columns_bg.startswith("#") else Color_Map_[self.header_selected_columns_bg]
        font = self.MT.my_hdr_font
        for c in range(start_col, end_col - 1):
            y = self.MT.hdr_fl_ins
            fc = self.MT.col_positions[c]
            sc = self.MT.col_positions[c + 1]
            if self.MT.all_columns_displayed:
                dcol = c
            else:
                dcol = self.MT.displayed_columns[c]
            tf, dd_drawn = self.redraw_highlight_get_text_fg(fc, sc, c, c_2, c_3, selected_cols, selected_rows, actual_selected_cols, dcol)

            if dcol in self.cell_options and 'align' in self.cell_options[dcol]:
                cell_alignment = self.cell_options[dcol]['align']
            elif dcol in self.MT.col_options and 'align' in self.MT.col_options[dcol]:
                cell_alignment = self.MT.col_options[dcol]['align']
            else:
                cell_alignment = self.align
                
            if cell_alignment == "w":
                x = fc + 5
                if dcol in self.cell_options and 'dropdown' in self.cell_options[dcol]:
                    mw = sc - fc - self.MT.hdr_txt_h - 2
                    self.redraw_dropdown(fc, 0, sc, self.current_height - 1, fill = tf, outline = tf, tag = "dd", draw_outline = not dd_drawn, draw_arrow = mw >= 5)
                else:
                    mw = sc - fc - 5

            elif cell_alignment == "e":
                if dcol in self.cell_options and 'dropdown' in self.cell_options[dcol]:
                    mw = sc - fc - self.MT.hdr_txt_h - 2
                    x = sc - 5 - self.MT.hdr_txt_h
                    self.redraw_dropdown(fc, 0, sc, self.current_height - 1, fill = tf, outline = tf, tag = "dd", draw_outline = not dd_drawn, draw_arrow = mw >= 5)
                else:
                    mw = sc - fc - 5
                    x = sc - 5

            elif cell_alignment == "center":
                #stop = fc + 5
                if dcol in self.cell_options and 'dropdown' in self.cell_options[dcol]:
                    mw = sc - fc - self.MT.hdr_txt_h - 2
                    x = fc + ceil((sc - fc - self.MT.hdr_txt_h) / 2)
                    self.redraw_dropdown(fc, 0, sc, self.current_height - 1, fill = tf, outline = tf, tag = "dd", draw_outline = not dd_drawn, draw_arrow = mw >= 5)
                else:
                    mw = sc - fc - 1
                    x = fc + floor((sc - fc) / 2)

            if dcol in self.cell_options and 'checkbox' in self.cell_options[dcol]:
                if mw > self.MT.hdr_txt_h + 2:
                    box_w = fc + self.MT.hdr_txt_h + 2 - fc
                    if cell_alignment == "w":
                        x = x + box_w 
                    elif cell_alignment == "center":
                        x = x + ceil(box_w / 2)
                    mw = mw - box_w
                    self.redraw_checkbox(dcol,
                                         fc + 2,
                                         0,
                                         fc + 2 + self.MT.hdr_txt_h + 2,
                                         self.MT.hdr_txt_h + 2,
                                         fill = tf if self.cell_options[dcol]['checkbox']['state'] == "normal" else self.header_grid_fg,
                                         outline = "",
                                         tag = "cb", 
                                         draw_check = self.MT.my_hdrs[dcol] if isinstance(self.MT.my_hdrs, (list, tuple)) else self.MT.data_ref[self.MT.my_hdrs][dcol])

            try:
                if dcol in self.cell_options and 'checkbox' in self.cell_options[dcol]:
                    lns = self.cell_options[dcol]['checkbox']['text'].split("\n") if isinstance(self.cell_options[dcol]['checkbox']['text'], str) else f"{self.cell_options[dcol]['checkbox']['text']}".split("\n")
                elif isinstance(self.MT.my_hdrs, int):
                    lns = self.MT.data_ref[self.MT.my_hdrs][dcol].split("\n") if isinstance(self.MT.data_ref[self.MT.my_hdrs][dcol], str) else f"{self.MT.data_ref[self.MT.my_hdrs][dcol]}".split("\n")
                else:
                    lns = self.MT.my_hdrs[dcol].split("\n") if isinstance(self.MT.my_hdrs[dcol], str) else f"{self.MT.my_hdrs[dcol]}".split("\n")
                got_hdr = True
            except:
                got_hdr = False
            if not got_hdr or (lns == [""] and self.show_default_header_for_empty):
                if self.default_hdr == "letters":
                    lns = (num2alpha(c), )
                elif self.default_hdr == "numbers":
                    lns = (f"{c + 1}", )
                else:
                    lns = (f"{c + 1} {num2alpha(c)}", )
            
            if cell_alignment == "center":
                if fc + 5 > x_stop or mw <= 5:
                    continue
                if incfl:
                    txt = lns[0]
                    t = self.redraw_text(x, y, text = txt, fill = tf, font = font, anchor = "center", tag = "t")
                    wd = self.bbox(t)
                    wd = wd[2] - wd[0]
                    if wd > mw:
                        tl = len(txt)
                        tmod = ceil((tl - int(tl * (mw / wd))) / 2)
                        txt = txt[tmod - 1:-tmod]
                        self.itemconfig(t, text = txt)
                        wd = self.bbox(t)
                        self.c_align_cyc = cycle(self.centre_alignment_text_mod_indexes)
                        while wd[2] - wd[0] > mw:
                            txt = txt[next(self.c_align_cyc)]
                            self.itemconfig(t, text = txt)
                            wd = self.bbox(t)
                        self.coords(t, x, y)
                if len(lns) > 1:
                    stl = int((top - y) / self.MT.hdr_xtra_lines_increment) - 1
                    if stl < 1:
                        stl = 1
                    y += (stl * self.MT.hdr_xtra_lines_increment)
                    if y + self.MT.hdr_half_txt_h - 1 < self.current_height:
                        for i in range(stl, len(lns)):
                            txt = lns[i]
                            t = self.redraw_text(x, y, text = txt, fill = tf, font = font, anchor = cell_alignment, tag = "t")
                            wd = self.bbox(t)
                            wd = wd[2] - wd[0]
                            if wd > mw:
                                tl = len(txt)
                                tmod = ceil((tl - int(tl * (mw / wd))) / 2)
                                txt = txt[tmod - 1:-tmod]
                                self.itemconfig(t, text = txt)
                                wd = self.bbox(t)
                                self.c_align_cyc = cycle(self.centre_alignment_text_mod_indexes)
                                while wd[2] - wd[0] > mw:
                                    txt = txt[next(self.c_align_cyc)]
                                    self.itemconfig(t, text = txt)
                                    wd = self.bbox(t)
                                self.coords(t, x, y)
                            y += self.MT.hdr_xtra_lines_increment
                            if y + self.MT.hdr_half_txt_h - 1 > self.current_height:
                                break
                            
            elif cell_alignment == "e":
                if x > x_stop or mw <= 5:
                    continue
                if incfl:
                    txt = lns[0]
                    t = self.redraw_text(x, y, text = txt, fill = tf, font = font, anchor = cell_alignment, tag = "t")
                    wd = self.bbox(t)
                    wd = wd[2] - wd[0]
                    if wd > mw:
                        txt = txt[len(txt) - int(len(txt) * (mw / wd)):]
                        self.itemconfig(t, text = txt)
                        wd = self.bbox(t)
                        while wd[2] - wd[0] > mw:
                            txt = txt[1:]
                            self.itemconfig(t, text = txt)
                            wd = self.bbox(t)
                if len(lns) > 1:
                    stl = int((top - y) / self.MT.hdr_xtra_lines_increment) - 1
                    if stl < 1:
                        stl = 1
                    y += (stl * self.MT.hdr_xtra_lines_increment)
                    if y + self.MT.hdr_half_txt_h - 1 < self.current_height:
                        for i in range(stl, len(lns)):
                            txt = lns[i]
                            t = self.redraw_text(x, y, text = txt, fill = tf, font = font, anchor = cell_alignment, tag = "t")
                            wd = self.bbox(t)
                            wd = wd[2] - wd[0]
                            if wd > mw:
                                txt = txt[len(txt) - int(len(txt) * (mw / wd)):]
                                self.itemconfig(t, text = txt)
                                wd = self.bbox(t)
                                while wd[2] - wd[0] > mw:
                                    txt = txt[1:]
                                    self.itemconfig(t, text = txt)
                                    wd = self.bbox(t)
                            y += self.MT.hdr_xtra_lines_increment
                            if y + self.MT.hdr_half_txt_h - 1 > self.current_height:
                                break
                
            elif cell_alignment == "w":
                if x > x_stop or mw <= 5:
                    continue
                if incfl:
                    txt = lns[0]
                    t = self.redraw_text(x, y, text = txt, fill = tf, font = font, anchor = cell_alignment, tag = "t")
                    wd = self.bbox(t)
                    wd = wd[2] - wd[0]
                    if wd > mw:
                        nl = int(len(txt) * (mw / wd))
                        self.itemconfig(t, text = txt[:nl])
                        wd = self.bbox(t)
                        while wd[2] - wd[0] > mw:
                            nl -= 1
                            self.dchars(t, nl)
                            wd = self.bbox(t)
                if len(lns) > 1:
                    stl = int((top - y) / self.MT.hdr_xtra_lines_increment) - 1
                    if stl < 1:
                        stl = 1
                    y += (stl * self.MT.hdr_xtra_lines_increment)
                    if y + self.MT.hdr_half_txt_h - 1 < self.current_height:
                        for i in range(stl, len(lns)):
                            txt = lns[i]
                            t = self.redraw_text(x, y, text = txt, fill = tf, font = font, anchor = cell_alignment, tag = "t")
                            wd = self.bbox(t)
                            wd = wd[2] - wd[0]
                            if wd > mw:
                                nl = int(len(txt) * (mw / wd))
                                self.itemconfig(t, text = txt[:nl])
                                wd = self.bbox(t)
                                while wd[2] - wd[0] > mw:
                                    nl -= 1
                                    self.dchars(t, nl)
                                    wd = self.bbox(t)
                            y += self.MT.hdr_xtra_lines_increment
                            if y + self.MT.hdr_half_txt_h - 1 > self.current_height:
                                break
                            
        self.redraw_gridline(x1, self.current_height - 1, x_stop, self.current_height - 1, fill = self.header_border_fg, width = 1, tag = "h")
        for t, sh in self.hidd_text.items():
            if sh:
                self.itemconfig(t, state = "hidden")
                self.hidd_text[t] = False
        for t, sh in self.hidd_high.items():
            if sh:
                self.itemconfig(t, state = "hidden")
                self.hidd_high[t] = False
        for t, sh in self.hidd_grid.items():
            if sh:
                self.itemconfig(t, state = "hidden")
                self.hidd_grid[t] = False
        for t, sh in self.hidd_dropdown.items():
            if sh:
                self.itemconfig(t, state = "hidden")
                self.hidd_dropdown[t] = False
        for t, sh in self.hidd_checkbox.items():
            if sh:
                self.itemconfig(t, state = "hidden")
                self.hidd_checkbox[t] = False
                
    def open_cell(self, event = None):
        if not self.MT.anything_selected() or self.text_editor_id is not None:
            return
        currently_selected = self.MT.currently_selected(get_coords = True)
        if not currently_selected:
            return
        x1 = int(currently_selected[1])
        dcol = x1 if self.MT.all_columns_displayed else self.MT.displayed_columns[x1]
        if dcol in self.cell_options and 'readonly' in self.cell_options[dcol]:
            return
        elif dcol in self.cell_options and ('dropdown' in self.cell_options[dcol] or 'checkbox' in self.cell_options[dcol]):
            if self.MT.event_opens_dropdown_or_checkbox(event):
                if 'dropdown' in self.cell_options[dcol]:
                    self.display_dropdown_window(x1, event = event)
                elif 'checkbox' in self.cell_options[dcol]:
                    self._click_checkbox(x1, dcol)
        elif self.edit_cell_enabled:
            self.edit_cell_(event, c = x1, dropdown = False)

    # c is displayed col
    def edit_cell_(self, event = None, c = None, dropdown = False):
        text = None
        extra_func_key = "??"
        if event is not None and ((hasattr(event, 'keysym') and event.keysym == 'BackSpace') or
                                  event.keycode in (8, 855638143)
                                  ):
            extra_func_key = "BackSpace"
            text = ""
        elif event is None or self.MT.event_opens_dropdown_or_checkbox(event):
            if event is not None:
                if hasattr(event, 'keysym') and event.keysym == 'Return':
                    extra_func_key = "Return"
                elif hasattr(event, 'keysym') and event.keysym == 'F2':
                    extra_func_key = "F2"
            dcol = c if self.MT.all_columns_displayed else self.MT.displayed_columns[c]
            if isinstance(self.MT.my_hdrs, list):
                if len(self.MT.my_hdrs) <= dcol:
                    self.MT.my_hdrs.extend(list(repeat("", dcol - len(self.MT.my_hdrs) + 1)))
                text = f"{self.MT.my_hdrs[dcol]}"
            elif isinstance(self.MT.my_hdrs, int):
                try:
                    text = f"{self.MT.data_ref[self.MT.my_hdrs][dcol]}"
                except:
                    text = ""
            if self.MT.cell_auto_resize_enabled:
                self.set_col_width_run_binding(c)
        elif event is not None and ((hasattr(event, "char") and event.char.isalpha()) or
                                    (hasattr(event, "char") and event.char.isdigit()) or
                                    (hasattr(event, "char") and event.char in symbols_set)):
            extra_func_key = event.char
            text = event.char
        else:
            return False
        self.text_editor_loc = (c)
        if self.extra_begin_edit_cell_func is not None:
            try:
                text2 = self.extra_begin_edit_cell_func(EditHeaderEvent(c, extra_func_key, text, "begin_edit_header"))
            except:
                return False
            if text2 is not None:
                text = text2
        text = "" if text is None else text
        self.select_col(c = c, keep_other_selections = True)
        self.create_text_editor(c = c, text = text, set_data_ref_on_destroy = True, dropdown = dropdown)
        return True

    # c is displayed col
    def create_text_editor(self,
                           c = 0,
                           text = None,
                           state = "normal",
                           see = True,
                           set_data_ref_on_destroy = False,
                           binding = None,
                           dropdown = False):
        if c == self.text_editor_loc and self.text_editor is not None:
            self.text_editor.set_text(self.text_editor.get() + "" if not isinstance(text, str) else text)
            return
        if self.text_editor is not None:
            self.destroy_text_editor()
        if see:
            has_redrawn = self.MT.see(r = 0, c = c, keep_yscroll = True, check_cell_visibility = True)
            if not has_redrawn:
                self.MT.refresh()
        self.text_editor_loc = c
        x = self.MT.col_positions[c] + 1
        y = 0
        w = self.MT.col_positions[c + 1] - x
        h = self.current_height + 1
        dcol = c if self.MT.all_columns_displayed else self.MT.displayed_columns[c]
        if text is None:
            if isinstance(self.MT.my_hdrs, list):
                if len(self.MT.my_hdrs) <= dcol:
                    self.MT.my_hdrs.extend(list(repeat("", dcol - len(self.MT.my_hdrs) + 1)))
                text = f"{self.MT.my_hdrs[dcol]}"
            elif isinstance(self.MT.my_hdrs, int):
                try:
                    text = f"{self.MT.data_ref[self.MT.my_hdrs][dcol]}"
                except:
                    text = ""
        bg, fg = self.get_widget_bg_fg(dcol)
        self.text_editor = TextEditor(self,
                                      text = text,
                                      font = self.MT.my_hdr_font,
                                      state = state,
                                      width = w,
                                      height = h,
                                      border_color = self.MT.table_selected_cells_border_fg,
                                      show_border = False,
                                      bg = bg,
                                      fg = fg,
                                      popup_menu_font = self.MT.popup_menu_font,
                                      popup_menu_fg = self.MT.popup_menu_fg,
                                      popup_menu_bg = self.MT.popup_menu_bg,
                                      popup_menu_highlight_bg = self.MT.popup_menu_highlight_bg,
                                      popup_menu_highlight_fg = self.MT.popup_menu_highlight_fg)
        self.text_editor_id = self.create_window((x, y), window = self.text_editor, anchor = "nw")
        if not dropdown:
            self.text_editor.textedit.focus_set()
            self.text_editor.scroll_to_bottom()
        self.text_editor.textedit.bind("<Alt-Return>", lambda x: self.text_editor_newline_binding(c))
        if USER_OS == 'Darwin':
            self.text_editor.textedit.bind("<Option-Return>", lambda x: self.text_editor_newline_binding(c))
        for key, func in self.MT.text_editor_user_bound_keys.items():
            self.text_editor.textedit.bind(key, func)
        if binding is not None:
            self.text_editor.textedit.bind("<Tab>", lambda x: binding((c, "Tab")))
            self.text_editor.textedit.bind("<Return>", lambda x: binding((c, "Return")))
            self.text_editor.textedit.bind("<FocusOut>", lambda x: binding((c, "FocusOut")))
            self.text_editor.textedit.bind("<Escape>", lambda x: binding((c, "Escape")))
        elif binding is None and set_data_ref_on_destroy:
            self.text_editor.textedit.bind("<Tab>", lambda x: self.get_text_editor_value((c, "Tab")))
            self.text_editor.textedit.bind("<Return>", lambda x: self.get_text_editor_value((c, "Return")))
            if not dropdown:
                self.text_editor.textedit.bind("<FocusOut>", lambda x: self.get_text_editor_value((c, "FocusOut")))
            self.text_editor.textedit.bind("<Escape>", lambda x: self.get_text_editor_value((c, "Escape")))
        else:
            self.text_editor.textedit.bind("<Escape>", lambda x: self.destroy_text_editor("Escape"))
            
    def bind_cell_edit(self, enable = True):
        if enable:
            self.edit_cell_enabled = True
        else:
            self.edit_cell_enabled = False

    def bind_text_editor_destroy(self, binding, c):
        self.text_editor.textedit.bind("<Return>", lambda x: binding((c, "Return")))
        self.text_editor.textedit.bind("<FocusOut>", lambda x: binding((c, "FocusOut")))
        self.text_editor.textedit.bind("<Escape>", lambda x: binding((c, "Escape")))
        self.text_editor.textedit.focus_set()

    def destroy_text_editor(self, event = None):
        if event is not None and self.extra_end_edit_cell_func is not None and self.text_editor_loc is not None:
            self.extra_end_edit_cell_func(EditHeaderEvent(int(self.text_editor_loc), "Escape", None, "escape_edit_header"))
        self.text_editor_loc = None
        try:
            self.delete(self.text_editor_id)
        except:
            pass
        try:
            self.text_editor.destroy()
        except:
            pass
        try:
            self.text_editor = None
        except:
            pass
        try:
            self.text_editor_id = None
        except:
            pass
        if event is not None and len(event) >= 3 and "Escape" in event:
            self.focus_set()

    # c is displayed col
    def get_text_editor_value(self, destroy_tup = None, c = None, set_data_ref_on_destroy = True, event = None, destroy = True, move_down = True, redraw = True, recreate = True):
        if self.focus_get() is None and destroy_tup:
            return
        if destroy_tup is not None and len(destroy_tup) >= 2 and destroy_tup[1] == "Escape":
            self.destroy_text_editor("Escape")
            self.hide_dropdown_window(c)
            return
        if self.text_editor is not None:
            self.text_editor_value = self.text_editor.get()
        if destroy:
            self.destroy_text_editor()
        if set_data_ref_on_destroy:
            if c is None and destroy_tup is not None and len(destroy_tup) >= 2:
                c = destroy_tup[0]
            if self.extra_end_edit_cell_func is not None:
                validation = self.extra_end_edit_cell_func(EditHeaderEvent(c, destroy_tup[1] if len(destroy_tup) >= 2 else "FocusOut", f"{self.text_editor_value}", "end_edit_header"))
                if validation is not None:
                    self.text_editor_value = validation
            self._set_cell_data(c, dcol = c if self.MT.all_columns_displayed else self.MT.displayed_columns[c], value = self.text_editor_value)
        if move_down:
            if c is None and destroy_tup is not None and len(destroy_tup) >= 2:
                c = destroy_tup[0]
            currently_selected = self.MT.currently_selected()
            if c is not None:
                if (
                    currently_selected and
                    c == currently_selected[1] and
                    (self.MT.single_selection_enabled or self.MT.toggle_selection_enabled)
                    ):
                    if destroy_tup is not None and len(destroy_tup) >= 2 and destroy_tup[1] == "Return":
                        self.select_col(c)
                        self.MT.see(0, c, keep_yscroll = True, bottom_right_corner = True, check_cell_visibility = True)
                    elif destroy_tup is not None and len(destroy_tup) >= 2 and destroy_tup[1] == "Tab":
                        self.select_col(c + 1 if c < len(self.MT.col_positions) - 2 else c)
                        self.MT.see(0, c + 1 if c < len(self.MT.col_positions) - 2 else c, keep_yscroll = True, bottom_right_corner = True, check_cell_visibility = True)
        self.hide_dropdown_window(c)
        if recreate:
            self.MT.recreate_all_selection_boxes()
        if redraw:
            self.MT.refresh()
        if destroy_tup is not None and len(destroy_tup) >= 2 and destroy_tup[1] != "FocusOut":
            self.focus_set()
        return self.text_editor_value
    
    #internal event use
    def _set_cell_data(self, c = 0, dcol = None, value = "", cell_resize = True):
        if dcol is None:
            dcol = c if self.MT.all_columns_displayed else self.MT.displayed_columns[c]
        if isinstance(self.MT.my_hdrs, list):
            if len(self.MT.my_hdrs) <= dcol:
                self.MT.my_hdrs.extend(list(repeat("", dcol - len(self.MT.my_hdrs) + 1)))
            self.MT.my_hdrs[dcol] = value
        elif isinstance(self.MT.my_hdrs, int):
            self.MT._set_cell_data(r = self.MT.my_hdrs, c = c, dcol = dcol, value = value, undo = True)
        if cell_resize and self.MT.cell_auto_resize_enabled:
            self.set_col_width_run_binding(c)
            self.MT.refresh()
            
    def set_col_width_run_binding(self, c):
        old_width = self.MT.col_positions[c + 1] - self.MT.col_positions[c]
        new_width = self.set_col_width(c, only_set_if_too_small = True)
        if self.column_width_resize_func is not None and old_width != new_width:
            self.column_width_resize_func(ResizeEvent("column_width_resize", c, old_width, new_width))

    #internal event use
    def _click_checkbox(self, c, dcol = None, undo = True, redraw = True):
        if dcol is None:
            dcol = c if self.MT.all_columns_displayed else self.MT.displayed_columns[c]
        if self.cell_options[dcol]['checkbox']['state'] == "normal":
            if isinstance(self.MT.my_hdrs, list):
                self._set_cell_data(c, dcol, value = not self.MT.my_hdrs[dcol] if type(self.MT.my_hdrs[dcol]) == bool else False, cell_resize = False)
            elif isinstance(self.MT.my_hdrs, int):
                self._set_cell_data(c, dcol, value = not self.MT.data_ref[self.MT.my_hdrs][dcol] if type(self.MT.data_ref[self.MT.my_hdrs][dcol]) == bool else False, cell_resize = False)
            if self.cell_options[dcol]['checkbox']['check_function'] is not None:
                self.cell_options[dcol]['checkbox']['check_function']((0, c, "HeaderCheckboxClicked", f"{self.MT.my_hdrs[dcol] if isinstance(self.MT.my_hdrs, list) else self.MT.data_ref[self.MT.my_hdrs][dcol]}"))
            if self.extra_end_edit_cell_func is not None:
                self.extra_end_edit_cell_func(EditHeaderEvent(c, "Return", f"{self.MT.my_hdrs[dcol] if isinstance(self.MT.my_hdrs, list) else self.MT.data_ref[self.MT.my_hdrs][dcol]}", "end_edit_header"))
        if redraw:
            self.MT.refresh()

    def create_checkbox(self, c = 0, checked = False, state = "normal", redraw = False, check_function = None, text = ""):
        if c in self.cell_options and any(x in self.cell_options[c] for x in ('dropdown', 'checkbox')):
            self.destroy_dropdown_and_checkbox(c)
        self._set_cell_data(dcol = c, value = checked, cell_resize = False)
        if c not in self.cell_options:
            self.cell_options[c] = {}
        self.cell_options[c]['checkbox'] = {'check_function': check_function,
                                            'state': state,
                                            'text': text}
        if redraw:
            self.MT.refresh()

    def create_dropdown(self, c = 0, values = [], set_value = None, state = "readonly", redraw = True, selection_function = None, modified_function = None):
        if c in self.cell_options and any(x in self.cell_options[c] for x in ('dropdown', 'checkbox')):
            self.destroy_dropdown_and_checkbox(c)
        if values:
            self.MT.headers(newheaders = set_value if set_value is not None else values[0], index = c)
        elif not values and set_value is not None:
            self.MT.headers(newheaders = set_value, index = c)
        if c not in self.cell_options:
            self.cell_options[c] = {}
        self.cell_options[c]['dropdown'] = {'values': values,
                                            'align': "w",
                                            'window': "no dropdown open",
                                            'canvas_id': "no dropdown open",
                                            'select_function': selection_function,
                                            'modified_function': modified_function,
                                            'state': state}
        if redraw:
            self.MT.refresh()
            
    def get_widget_bg_fg(self, c):
        bg = self.header_bg
        fg = self.header_fg
        if c in self.cell_options and 'highlight' in self.cell_options[c]:
            if self.cell_options[c]['highlight'][0] is not None:
                bg = self.cell_options[c]['highlight'][0]
            if self.cell_options[c]['highlight'][1] is not None:
                fg = self.cell_options[c]['highlight'][1]
        return bg, fg

    def get_dropdown_height_anchor(self, dcol, text_editor_h = None):
        numvalues = len(self.cell_options[dcol]['dropdown']['values'])
        xscroll_h = self.MT.parentframe.xscroll.winfo_height()
        if numvalues > 5:
            linespace = 6 * 5 + 3
            win_h = int(self.MT.hdr_txt_h * 6 + linespace + xscroll_h)
        else:
            linespace = numvalues * 5 + 3
            win_h = int(self.MT.hdr_txt_h * numvalues + linespace + xscroll_h)
        if win_h > 300:
            win_h = 300
        space_bot = self.MT.get_space_bot(0, text_editor_h)
        win_h2 = int(win_h)
        if win_h > space_bot:
            win_h = space_bot - 1
        if win_h < self.MT.hdr_txt_h + 5:
            win_h = self.MT.hdr_txt_h + 5
        elif win_h > win_h2:
            win_h = win_h2
        return win_h, "nw"

    def text_editor_newline_binding(self, c = None, event = None):
        if self.GetLinesHeight(self.text_editor.get_num_lines() + 1) > self.text_editor.winfo_height():
            self.text_editor.config(height = self.text_editor.winfo_height() + self.MT.hdr_xtra_lines_increment)
            dcol = c if self.MT.all_columns_displayed else self.MT.displayed_columns[c]
            if c if self.MT.all_columns_displayed else self.MT.displayed_columns[c] in self.cell_options and 'dropdown' in self.cell_options[dcol]:
                text_editor_h = self.text_editor.winfo_height()
                win_h, anchor = self.get_dropdown_height_anchor(dcol, text_editor_h)
                self.coords(self.cell_options[dcol]['dropdown']['canvas_id'],
                            self.MT.col_positions[c], self.MT.canvasy(0))
                self.itemconfig(self.cell_options[dcol]['dropdown']['canvas_id'],
                                anchor = anchor, height = win_h)
            
    # c is displayed col
    def display_dropdown_window(self, c, dcol = None, event = None):
        self.destroy_text_editor("Escape")
        self.delete_opened_dropdown_window()
        if dcol is None:
            dcol = c if self.MT.all_columns_displayed else self.MT.displayed_columns[c]
        if self.cell_options[dcol]['dropdown']['state'] == "normal":
            if not self.edit_cell_(c = c, dropdown = True, event = event):
                return
        bg, fg = self.get_widget_bg_fg(dcol)
        win_h, anchor = self.get_dropdown_height_anchor(dcol)
        window = self.MT.parentframe.dropdown_class(self.MT.winfo_toplevel(),
                                                    0,
                                                    c,
                                                    width = self.MT.col_positions[c + 1] - self.MT.col_positions[c] + 1,
                                                    height = win_h,
                                                    font = self.MT.my_hdr_font,
                                                    bg = bg,
                                                    fg = fg,
                                                    outline_color = fg,
                                                    values = self.cell_options[dcol]['dropdown']['values'],
                                                    hide_dropdown_window = self.hide_dropdown_window,
                                                    arrowkey_RIGHT = self.MT.arrowkey_RIGHT,
                                                    arrowkey_LEFT = self.MT.arrowkey_LEFT,
                                                    align = self.cell_options[dcol]['dropdown']['align'],
                                                    single_index = "c")
        ypos = self.current_height - 1
        self.cell_options[dcol]['dropdown']['canvas_id'] = self.create_window((self.MT.col_positions[c], ypos),
                                                                               window = window,
                                                                               anchor = anchor)
        if self.cell_options[dcol]['dropdown']['state'] == "normal":
            if self.cell_options[dcol]['dropdown']['modified_function'] is not None:
                self.text_editor.textedit.bind("<<TextModified>>", self.cell_options[dcol]['dropdown']['modified_function'])
            self.update()
            try:
                self.text_editor.textedit.focus_set()
                self.text_editor.scroll_to_bottom()
            except:
                return
        else:
            window.bind("<FocusOut>", lambda x: self.hide_dropdown_window(c))
            self.update()
            try:
                window.focus_set()
            except:
                return
        self.existing_dropdown_window = window
        self.cell_options[dcol]['dropdown']['window'] = window
        self.existing_dropdown_canvas_id = self.cell_options[dcol]['dropdown']['canvas_id']
            
    # c is displayed col
    def hide_dropdown_window(self, c = None, selection = None, b1 = False, redraw = True):
        if c is not None and selection is not None:
            dcol = c if self.MT.all_columns_displayed else self.MT.displayed_columns[c]
            if self.cell_options[dcol]['dropdown']['select_function'] is not None: # user has specified a selection function
                self.cell_options[dcol]['dropdown']['select_function'](EditHeaderEvent(c, "HeaderComboboxSelected", f"{selection}", "end_edit_header"))
            if self.extra_end_edit_cell_func is not None:
                validation = self.extra_end_edit_cell_func(EditHeaderEvent(c, "HeaderComboboxSelected", f"{selection}", "end_edit_header"))
                if validation is not None:
                    selection = validation
            self._set_cell_data(c, dcol, selection, cell_resize = True)
            self.focus_set()
            self.MT.recreate_all_selection_boxes()
            if redraw:
                self.MT.refresh()
        if self.existing_dropdown_window is not None:
            closedc, return_closedc = int(self.existing_dropdown_window.c), True
        else:
            return_closedc = False
        if b1 and self.text_editor_loc is not None and self.text_editor is not None:
            self.get_text_editor_value(destroy_tup = (self.text_editor_loc, "Return"))
        else:
            self.destroy_text_editor("Escape")
        self.delete_opened_dropdown_window(c)
        if return_closedc:
            return closedc
            
    # c is displayed col
    def delete_opened_dropdown_window(self, c = None, dcol = None):
        if c is not None and dcol is None:
            dcol = c if self.MT.all_columns_displayed else self.MT.displayed_columns[c]
        try:
            self.delete(self.existing_dropdown_canvas_id)
        except:
            pass
        self.existing_dropdown_canvas_id = None
        try:
            self.existing_dropdown_window.destroy()
        except:
            pass
        self.existing_dropdown_window = None
        if c is not None and dcol in self.cell_options and 'dropdown' in self.cell_options[dcol]:
            self.cell_options[dcol]['dropdown']['canvas_id'] = "no dropdown open"
            self.cell_options[dcol]['dropdown']['window'] = "no dropdown open"
            try:
                self.delete(self.cell_options[dcol]['dropdown']['canvas_id'])
            except:
                pass
            
    # c is dcol
    def destroy_dropdown(self, c):
        self.delete_opened_dropdown_window(c)
        if c in self.cell_options and 'dropdown' in self.cell_options[c]:
            del self.cell_options[c]['dropdown']
            
    # c is dcol
    def destroy_checkbox(self, c):
        if c in self.cell_options and 'checkbox' in self.cell_options[c]:
            del self.cell_options[c]['checkbox']

    # c is dcol
    def destroy_dropdown_and_checkbox(self, c):
        self.destroy_dropdown(c)
        self.destroy_checkbox(c)
