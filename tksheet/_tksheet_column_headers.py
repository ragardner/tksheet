from ._tksheet_vars import *
from ._tksheet_other_classes import *

from collections import defaultdict, deque
from itertools import islice, repeat, accumulate, chain, product, cycle
from math import floor, ceil
from tkinter import ttk
import bisect
import csv as csv_module
import io
import pickle
import re
import tkinter as tk
import zlib
# for mac bindings
from platform import system as get_os


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
                 measure_subset_header = True,
                 resizing_line_fg = None):
        tk.Canvas.__init__(self,parentframe,
                           background = header_bg,
                           highlightthickness = 0)

        self.disp_text = {}
        self.disp_high = {}
        self.disp_grid = {}
        self.disp_fill_sels = {}
        self.disp_col_exps = {}
        self.disp_resize_lines = {}
        self.hidd_text = {}
        self.hidd_high = {}
        self.hidd_grid = {}
        self.hidd_fill_sels = {}
        self.hidd_col_exps = {}
        self.hidd_resize_lines = {}
        
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
        self.default_hdr = default_header.lower()
        self.max_cw = float(max_colwidth)
        self.max_header_height = float(max_header_height)
        self.current_height = None    # is set from within MainTable() __init__ or from Sheet parameters
        self.MT = main_canvas         # is set from within MainTable() __init__
        self.RI = row_index_canvas    # is set from within MainTable() __init__
        self.TL = None                # is set from within TopLeftRectangle() __init__
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
        self.measure_subset_hdr = measure_subset_header
        self.dragged_col = None
        self.visible_col_dividers = []
        self.col_height_resize_bbox = tuple()
        self.cell_options = {}
        self.rsz_w = None
        self.rsz_h = None
        self.new_col_height = 0
        self.currently_resizing_width = False
        self.currently_resizing_height = False
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
        self.config(height = new_height)
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
        self.focus_set()
        if self.MT.identify_col(x = event.x, allow_end = False) is None:
            self.MT.deselect("all")
            self.ch_rc_popup_menu.tk_popup(event.x_root, event.y_root)
        elif self.col_selection_enabled and not self.currently_resizing_width and not self.currently_resizing_height:
            c = self.MT.identify_col(x = event.x)
            if c < len(self.MT.col_positions) - 1:
                if self.MT.col_selected(c):
                    if self.MT.rc_popup_menus_enabled:
                        self.ch_rc_popup_menu.tk_popup(event.x_root, event.y_root)
                else:
                    if self.MT.single_selection_enabled and self.MT.rc_select_enabled:
                        self.select_col(c, redraw = True)
                    elif self.MT.toggle_selection_enabled and self.MT.rc_select_enabled:
                        self.toggle_select_col(c, redraw = True)
                    if self.MT.rc_popup_menus_enabled:
                        self.ch_rc_popup_menu.tk_popup(event.x_root, event.y_root)
        if self.extra_rc_func is not None:
            self.extra_rc_func(event)

    def shift_b1_press(self, event):
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
                        elif c < min_c:
                            self.MT.create_selected(0, c, len(self.MT.row_positions) - 1, min_c + 1, "cols")
                    else:
                        self.select_col(c)
                    self.MT.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
                    if self.shift_selection_binding_func is not None:
                        self.shift_selection_binding_func(("shift_select_columns", tuple(sorted(self.MT.get_selected_cols()))))
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
                    for itm in ov:
                        tgs = self.gettags(itm)
                        if "v" == tgs[0]:
                            break
                    c = int(tgs[1])
                    self.rsz_w = c
                    self.config(cursor = "sb_h_double_arrow")
                    mouse_over_resize = True
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
        
    def b1_press(self, event = None):
        self.focus_set()
        self.MT.unbind("<MouseWheel>")
        x1, y1, x2, y2 = self.MT.get_canvas_visible_area()
        if self.check_mouse_position_width_resizers(event) is None:
            self.rsz_w = None
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
            c = self.MT.identify_col(x = event.x)
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
                            self.drag_selection_binding_func(("drag_select_columns", func_event))
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
            size = new_col_pos - self.MT.col_positions[self.rsz_w - 1]
            if size < self.MT.min_cw:
                new_row_pos = ceil(self.MT.col_positions[self.rsz_w - 1] + self.MT.min_cw)
            elif size > self.max_cw:
                new_col_pos = floor(self.MT.col_positions[self.rsz_w - 1] + self.max_cw)
            increment = new_col_pos - self.MT.col_positions[self.rsz_w]
            self.MT.col_positions[self.rsz_w + 1:] = [e + increment for e in islice(self.MT.col_positions, self.rsz_w + 1, len(self.MT.col_positions))]
            self.MT.col_positions[self.rsz_w] = new_col_pos
            self.MT.recreate_all_selection_boxes()
            self.MT.refresh_dropdowns()
            self.MT.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
        elif self.height_resizing_enabled and self.rsz_h is not None and self.currently_resizing_height:
            self.currently_resizing_height = False
            self.delete_resize_lines()
            self.MT.delete_resize_lines()
            self.set_height(self.new_col_height,set_TL = True)
            self.MT.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
        if self.drag_and_drop_enabled and self.col_selection_enabled and self.MT.anything_selected(exclude_cells = True, exclude_rows = True) and self.rsz_h is None and self.rsz_w is None and self.dragged_col is not None:
            self.delete_resize_lines()
            self.MT.delete_resize_lines()
            x = event.x
            c = self.MT.identify_col(x = x)
            orig_selected_cols = self.MT.get_selected_cols()
            if self.MT.all_columns_displayed and c != self.dragged_col and c is not None and c not in orig_selected_cols and len(orig_selected_cols) != len(self.MT.col_positions) - 1:
                orig_selected_cols = sorted(orig_selected_cols)
                if len(orig_selected_cols) > 1:
                    orig_min = orig_selected_cols[0]
                    orig_max = orig_selected_cols[1]
                    start_idx = bisect.bisect_left(orig_selected_cols, self.dragged_col)
                    forward_gap = get_index_of_gap_in_sorted_integer_seq_forward(orig_selected_cols, start_idx)
                    reverse_gap = get_index_of_gap_in_sorted_integer_seq_reverse(orig_selected_cols, start_idx)
                    if forward_gap is not None:
                        orig_selected_cols[:] = orig_selected_cols[:forward_gap]
                    if reverse_gap is not None:
                        orig_selected_cols[:] = orig_selected_cols[reverse_gap:]
                colsiter = orig_selected_cols.copy()
                rm1start = colsiter[0]
                rm1end = colsiter[-1] + 1
                rm2start = rm1start + (rm1end - rm1start)
                rm2end = rm1end + (rm1end - rm1start)
                totalcols = len(colsiter)
                if c >= len(self.MT.col_positions) - 1:
                    c -= 1
                c_ = int(c)
                if self.ch_extra_begin_drag_drop_func is not None:
                    self.ch_extra_begin_drag_drop_func(("begin_column_header_drag_drop", tuple(orig_selected_cols), int(c)))
                if self.column_drag_and_drop_perform:
                    if self.MT.all_columns_displayed:
                        if rm1start > c:
                            for rn in range(len(self.MT.data_ref)):
                                try:
                                    self.MT.data_ref[rn] = (self.MT.data_ref[rn][:c] +
                                                            self.MT.data_ref[rn][rm1start:rm1start + totalcols] +
                                                            self.MT.data_ref[rn][c:rm1start] +
                                                            self.MT.data_ref[rn][rm1start + totalcols:])
                                except:
                                    continue
                            if not isinstance(self.MT.my_hdrs, int) and self.MT.my_hdrs:
                                try:
                                    self.MT.my_hdrs = (self.MT.my_hdrs[:c] +
                                                       self.MT.my_hdrs[rm1start:rm1start + totalcols] +
                                                       self.MT.my_hdrs[c:rm1start] +
                                                       self.MT.my_hdrs[rm1start + totalcols:])
                                except:
                                    pass
                        else:
                            for rn in range(len(self.MT.data_ref)):
                                try:
                                    self.MT.data_ref[rn] = (self.MT.data_ref[rn][:rm1start] +
                                                            self.MT.data_ref[rn][rm1start + totalcols:c + 1] +
                                                            self.MT.data_ref[rn][rm1start:rm1start + totalcols] +
                                                            self.MT.data_ref[rn][c + 1:])
                                except:
                                    continue
                            if not isinstance(self.MT.my_hdrs, int) and self.MT.my_hdrs:
                                try:
                                    self.MT.my_hdrs = (self.MT.my_hdrs[:rm1start] +
                                                       self.MT.my_hdrs[rm1start + totalcols:c + 1] +
                                                       self.MT.my_hdrs[rm1start:rm1start + totalcols] +
                                                       self.MT.my_hdrs[c + 1:])
                                except:
                                    pass
                    else:
                        pass
                        # drag and drop columns when some are hidden disabled until correct code written
                cws = [int(b - a) for a, b in zip(self.MT.col_positions, islice(self.MT.col_positions, 1, len(self.MT.col_positions)))]
                if rm1start > c:
                    cws = (cws[:c] +
                           cws[rm1start:rm1start + totalcols] +
                           cws[c:rm1start] +
                           cws[rm1start + totalcols:])
                else:
                    cws = (cws[:rm1start] +
                           cws[rm1start + totalcols:c + 1] +
                           cws[rm1start:rm1start + totalcols] +
                           cws[c + 1:])
                self.MT.col_positions = list(accumulate(chain([0], (width for width in cws))))
                self.MT.deselect("all")
                if (c_ - 1) + totalcols > len(self.MT.col_positions) - 1:
                    new_selected = tuple(range(len(self.MT.col_positions) - 1 - totalcols, len(self.MT.col_positions) - 1))
                    self.MT.create_selected(0, len(self.MT.col_positions) - 1 - totalcols, len(self.MT.row_positions) - 1, len(self.MT.col_positions) - 1, "cols")
                else:
                    if rm1start > c:
                        new_selected = tuple(range(c_, c_ + totalcols))
                        self.MT.create_selected(0, c_, len(self.MT.row_positions) - 1, c_ + totalcols, "cols")
                    else:
                        new_selected = tuple(range(c_ + 1 - totalcols, c_ + 1))
                        self.MT.create_selected(0, c_ + 1 - totalcols, len(self.MT.row_positions) - 1, c_ + 1, "cols")
                self.MT.create_current(0, int(new_selected[0]), type_ = "col", inside = True)
                if self.MT.undo_enabled:
                    self.MT.undo_storage.append(zlib.compress(pickle.dumps(("move_cols",
                                                                            int(orig_selected_cols[0]),
                                                                            int(new_selected[0]),
                                                                             int(new_selected[-1]),
                                                                             sorted(orig_selected_cols)))))
                colset = set(colsiter)
                popped_ch = {t1: t2 for t1, t2 in self.cell_options.items() if t1 in colset}
                popped_cell = {t1: t2 for t1, t2 in self.MT.cell_options.items() if t1[1] in colset}
                popped_col = {t1: t2 for t1, t2 in self.MT.col_options.items() if t1 in colset}
                
                popped_ch = {t1: self.cell_options.pop(t1) for t1 in popped_ch}
                popped_cell = {t1: self.MT.cell_options.pop(t1) for t1 in popped_cell}
                popped_col = {t1: self.MT.col_options.pop(t1) for t1 in popped_col}

                self.cell_options = {t1 if t1 < rm1start else t1 - totalcols: t2 for t1, t2 in self.cell_options.items()}
                self.cell_options = {t1 if t1 < c_ else t1 + totalcols: t2 for t1, t2 in self.cell_options.items()}

                self.MT.col_options = {t1 if t1 < rm1start else t1 - totalcols: t2 for t1, t2 in self.MT.col_options.items()}
                self.MT.col_options = {t1 if t1 < c_ else t1 + totalcols: t2 for t1, t2 in self.MT.col_options.items()}

                self.MT.cell_options = {(t10, t11 if t11 < rm1start else t11 - totalcols): t2 for (t10, t11), t2 in self.MT.cell_options.items()}
                self.MT.cell_options = {(t10, t11 if t11 < c_ else t11 + totalcols): t2 for (t10, t11), t2 in self.MT.cell_options.items()}

                newcolsdct = {t1: t2 for t1, t2 in zip(colsiter, new_selected)}
                for t1, t2 in popped_ch.items():
                    self.cell_options[newcolsdct[t1]] = t2

                for t1, t2 in popped_col.items():
                    self.MT.col_options[newcolsdct[t1]] = t2

                for (t10, t11), t2 in popped_cell.items():
                    self.MT.cell_options[(t10, newcolsdct[t11])] = t2

                self.MT.refresh_dropdowns()
                
                self.MT.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
                if self.ch_extra_end_drag_drop_func is not None:
                    self.ch_extra_end_drag_drop_func(("end_column_header_drag_drop", tuple(orig_selected_cols), new_selected, int(c)))
        self.dragged_col = None
        self.currently_resizing_width = False
        self.currently_resizing_height = False
        self.rsz_w = None
        self.rsz_h = None
        self.being_drawn_rect = None
        self.mouse_motion(event)
        if self.extra_b1_release_func is not None:
            self.extra_b1_release_func(event)

    def double_b1(self, event = None):
        self.focus_set()
        if self.double_click_resizing_enabled and self.width_resizing_enabled and self.rsz_w is not None and not self.currently_resizing_width:
            col = self.rsz_w - 1
            self.set_col_width(col)
            self.MT.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
        elif self.col_selection_enabled and self.rsz_h is None and self.rsz_w is None:
            c = self.MT.identify_col(x = event.x)
            if c < len(self.MT.col_positions) - 1:
                if self.MT.single_selection_enabled:
                    self.select_col(c, redraw = True)
                elif self.MT.toggle_selection_enabled:
                    self.toggle_select_col(c, redraw = True)
        self.mouse_motion(event)
        self.rsz_w = None
        if self.extra_double_b1_func is not None:
            self.extra_double_b1_func(event)

    def highlight_cells(self, c = 0, cells = tuple(), bg = None, fg = None, redraw = False):
        if bg is None and fg is None:
            return
        if cells:
            for c_ in cells:
                if c_ not in self.cell_options:
                    self.cell_options[c_] = {}
                self.cell_options[c_]['highlight'] = (bg, fg)
        else:
            if c not in self.cell_options:
                self.cell_options[c] = {}
            self.cell_options[c]['highlight'] = (bg, fg)
        if redraw:
            self.MT.main_table_redraw_grid_and_text(True, False)

    def select_col(self, c, redraw = False, keep_other_selections = False):
        c = int(c)
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
            self.selection_binding_func(("select_column", int(c)))

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
        if width is None:
            w = self.MT.min_cw
            if displayed_only:
                x1, y1, x2, y2 = self.MT.get_canvas_visible_area()
                start_row, end_row = self.MT.get_visible_rows(y1, y2)
            else:
                start_row, end_row = 0, None
            if self.MT.all_columns_displayed:
                data_col = col
            else:
                data_col = self.MT.displayed_columns[col]
            try:
                if isinstance(self.MT.my_hdrs, int):
                    txt = self.MT.data_ref[self.MT.my_hdrs][data_col]
                else:
                    txt = self.MT.my_hdrs[data_col if self.measure_subset_hdr else col]
                if txt:
                    qconf(qtxtm, text = txt, font = self.MT.my_hdr_font)
                    b = qbbox(qtxtm)
                    hw = b[2] - b[0] + 5
                else:
                    hw = self.MT.min_cw
            except:
                if self.default_hdr == "letters":
                    hw = self.MT.GetHdrTextWidth(num2alpha(data_col)) + 5
                elif self.default_hdr == "numbers":
                    hw = self.MT.GetHdrTextWidth(f"{data_col + 1}") + 5
                else:
                    hw = self.MT.GetHdrTextWidth(f"{data_col + 1} {num2alpha(data_col)}") + 5
            for rn, r in enumerate(islice(self.MT.data_ref, start_row, end_row), start_row):
                try:
                    if isinstance(r[data_col], str):
                        txt = r[data_col]
                    else:
                        txt = f"{r[data_col]}"
                except:
                    txt = ""
                if txt:
                    qconf(qtxtm, text = txt, font = self.MT.my_font)
                    b = qbbox(qtxtm)
                    tw = b[2] - b[0] + 25 if (rn, data_col) in self.MT.cell_options and 'dropdown' in self.MT.cell_options[(rn, data_col)] else b[2] - b[0] + 5
                    if tw > w:
                        w = tw
                elif (rn, data_col) in self.MT.cell_options and 'dropdown' in self.MT.cell_options[(rn, data_col)]:
                    tw = 20
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
                self.MT.refresh_dropdowns()

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
            self.MT.refresh_dropdowns()

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
        if hlcol in self.cell_options and 'highlight' in self.cell_options[hlcol] and c in actual_selected_cols:
            if self.cell_options[hlcol]['highlight'][0] is not None:
                c_1 = self.cell_options[hlcol]['highlight'][0] if self.cell_options[hlcol]['highlight'][0].startswith("#") else Color_Map_[self.cell_options[hlcol]['highlight'][0]]
                self.redraw_highlight(fc + 1,
                                      0,
                                      sc,
                                      self.current_height - 1,
                                      fill = (f"#{int((int(c_1[1:3], 16) + int(c_3[1:3], 16)) / 2):02X}" +
                                              f"{int((int(c_1[3:5], 16) + int(c_3[3:5], 16)) / 2):02X}" +
                                              f"{int((int(c_1[5:], 16) + int(c_3[5:], 16)) / 2):02X}"),
                                      outline = "",
                                      tag = "s")
            tf = self.header_selected_columns_fg if self.cell_options[hlcol]['highlight'][1] is None or self.MT.display_selected_fg_over_highlights else self.cell_options[hlcol]['highlight'][1]
        elif hlcol in self.cell_options and 'highlight' in self.cell_options[hlcol] and (c in selected_cols or selected_rows):
            if self.cell_options[hlcol]['highlight'][0] is not None:
                c_1 = self.cell_options[hlcol]['highlight'][0] if self.cell_options[hlcol]['highlight'][0].startswith("#") else Color_Map_[self.cell_options[hlcol]['highlight'][0]]
                self.redraw_highlight(fc + 1,
                                      0,
                                      sc,
                                      self.current_height - 1,
                                      fill = (f"#{int((int(c_1[1:3], 16) + int(c_2[1:3], 16)) / 2):02X}" +
                                              f"{int((int(c_1[3:5], 16) + int(c_2[3:5], 16)) / 2):02X}" +
                                              f"{int((int(c_1[5:], 16) + int(c_2[5:], 16)) / 2):02X}"),
                                      outline = "",
                                      tag = "s")
            tf = self.header_selected_cells_fg if self.cell_options[hlcol]['highlight'][1] is None or self.MT.display_selected_fg_over_highlights else self.cell_options[hlcol]['highlight'][1]
        elif c in actual_selected_cols:
            tf = self.header_selected_columns_fg
        elif c in selected_cols or selected_rows:
            tf = self.header_selected_cells_fg
        elif hlcol in self.cell_options and 'highlight' in self.cell_options[hlcol]:
            if self.cell_options[hlcol]['highlight'][0] is not None:
                self.redraw_highlight(fc + 1, 0, sc, self.current_height - 1, fill = self.cell_options[hlcol]['highlight'][0], outline = "", tag = "s")
            tf = self.header_fg if self.cell_options[hlcol]['highlight'][1] is None else self.cell_options[hlcol]['highlight'][1]
        else:
            tf = self.header_fg
        return tf, self.MT.my_hdr_font

    def redraw_highlight(self, x1, y1, x2, y2, fill, outline, tag):
        if self.hidd_high:
            t, sh = self.hidd_high.popitem()
            self.coords(t, x1, y1, x2, y2)
            if sh:
                self.itemconfig(t, fill = fill, outline = outline, tag = tag)
            else:
                self.itemconfig(t, fill = fill, outline = outline, tag = tag, state = "normal")
            self.lift(t)
            self.disp_high[t] = True
        else:
            self.disp_high[self.create_rectangle(x1, y1, x2, y2, fill = fill, outline = outline, tag = tag)] = True

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
                self.itemconfig(t, fill = fill, width = width, tag = tag)
            else:
                self.itemconfig(t, fill = fill, width = width, tag = tag, state = "normal")
            self.disp_grid[t] = True
        else:
            self.disp_grid[self.create_line(x1, y1, x2, y2, fill = fill, width = width, tag = tag)] = True

    def redraw_hidden_col_expander(self, x1, y1, x2, y2, fill, outline, tag):
        if self.hidd_col_exps:
            t, sh = self.hidd_col_exps.popitem()
            self.coords(t, x1, y1, x2, y2)
            if sh:
                self.itemconfig(t, fill = fill, outline = outline, tag = tag)
            else:
                self.itemconfig(t, fill = fill, outline = outline, tag = tag, state = "normal")
            self.lift(t)
            self.disp_col_exps[t] = True
        else:
            t = self.create_rectangle(x1, y1, x2, y2, fill = fill, outline = outline, tag = tag)
            self.disp_col_exps[t] = True
        self.tag_bind(t, "<Button-1>", self.click_expander)

    def click_expander(self, event = None):
        c = self.MT.identify_col(x = event.x, allow_end = False)
        if c is not None and self.rsz_w is None and self.rsz_h is None:
            disp = sorted(self.MT.displayed_columns)
            col = self.MT.displayed_columns[c]
            idx = disp.index(col)
            ins = idx + 1
            if idx == len(disp) - 1:
                total = self.MT.total_data_cols()
                newcols = list(range(col + 1, total))
                self.MT.displayed_columns[ins:ins] = newcols
            else:
                newcols = list(range(disp[idx] + 1, disp[idx + 1]))
                self.MT.displayed_columns[ins:ins] = newcols
            self.MT.insert_col_positions(idx, len(newcols))
            self.MT.hidd_col_expander_idxs.discard(col)

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
        self.hidd_col_exps.update(self.disp_col_exps)
        self.disp_col_exps = {}
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
            if self.hide_columns_enabled and len(self.MT.displayed_columns) > c and self.MT.displayed_columns[c] in self.MT.hidd_col_expander_idxs:
                self.redraw_hidden_col_expander(self.MT.col_positions[c + 1] - 2, 2, self.MT.col_positions[c + 1] - 6, self.current_height - 2, fill = self.header_hidden_columns_expander_bg, outline = "",
                                                tag = ("hidd", f"{c}"))
        top = self.canvasy(0)
        if self.MT.hdr_fl_ins + self.MT.hdr_half_txt_h - 1 > top:
            incfl = True
        else:
            incfl = False
        c_2 = self.header_selected_cells_bg if self.header_selected_cells_bg.startswith("#") else Color_Map_[self.header_selected_cells_bg]
        c_3 = self.header_selected_columns_bg if self.header_selected_columns_bg.startswith("#") else Color_Map_[self.header_selected_columns_bg]
        for c in range(start_col, end_col - 1):
            fc = self.MT.col_positions[c]
            sc = self.MT.col_positions[c + 1]
            if self.MT.all_columns_displayed:
                dcol = c
            else:
                dcol = self.MT.displayed_columns[c]
            tf, font = self.redraw_highlight_get_text_fg(fc, sc, c, c_2, c_3, selected_cols, selected_rows, actual_selected_cols, dcol)

            if dcol in self.cell_options and 'align' in self.cell_options[dcol]:
                cell_alignment = self.cell_options[dcol]['align']
            elif dcol in self.MT.col_options and 'align' in self.MT.col_options[dcol]:
                cell_alignment = self.MT.col_options[dcol]['align']
            else:
                cell_alignment = self.align

            try:
                if isinstance(self.MT.my_hdrs, int):
                    lns = self.MT.data_ref[self.MT.my_hdrs][dcol].split("\n") if isinstance(self.MT.data_ref[self.MT.my_hdrs][dcol], str) else f"{self.MT.data_ref[self.MT.my_hdrs][dcol]}".split("\n")
                else:
                    lns = self.MT.my_hdrs[dcol].split("\n") if isinstance(self.MT.my_hdrs[dcol], str) else f"{self.MT.my_hdrs[dcol]}".split("\n")
            except:
                if self.default_hdr == "letters":
                    lns = (num2alpha(c), )
                elif self.default_hdr == "numbers":
                    lns = (f"{c + 1}", )
                else:
                    lns = (f"{c + 1} {num2alpha(c)}", )
            
            if cell_alignment == "center":
                mw = sc - fc - 1
                if fc + 5 > x_stop or mw <= 5:
                    continue
                x = fc + floor((sc - fc) / 2)
                y = self.MT.hdr_fl_ins
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
                mw = sc - fc - 5
                x = sc - 5
                if x > x_stop or mw <= 5:
                    continue
                y = self.MT.hdr_fl_ins
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
                mw = sc - fc - 5
                x = fc + 5
                if x > x_stop or mw <= 5:
                    continue
                y = self.MT.hdr_fl_ins
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
        for t, sh in self.hidd_col_exps.items():
            if sh:
                self.itemconfig(t, state = "hidden")
                self.hidd_col_exps[t] = False
        
    def GetCellCoords(self, event = None, r = None, c = None):
        pass

    
