from ._tksheet_vars import *
from ._tksheet_other_classes import *
from ._tksheet_top_left_rectangle import *
from ._tksheet_column_headers import *
from ._tksheet_row_index import *
from ._tksheet_main_table import *

from collections import defaultdict, deque
from itertools import islice, repeat, accumulate, chain
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


class Sheet(tk.Frame):
    def __init__(self,
                 C,
                 show = True,
                 show_top_left = True,
                 show_row_index = True,
                 show_header = True,
                 width = None,
                 height = None,
                 headers = None,
                 default_header = "letters", #letters or numbers
                 header_background = "#f8f9fa",
                 header_border_color = "#ababab",
                 header_grid_color = "#ababab",
                 header_foreground = "black",
                 header_select_background = "#e8eaed",
                 header_select_foreground = "black",
                 data_reference = None,
                 total_columns = None,
                 total_rows = None,
                 column_width = 120,
                 header_height = "1",
                 max_colwidth = "inf",
                 max_rh = "inf",
                 max_header_height = "inf",
                 max_row_width = "inf",
                 row_index = None,
                 row_index_width = 100,
                 row_height = "1",
                 row_index_background = "#f8f9fa",
                 row_index_border_color = "#ababab",
                 row_index_grid_color = "#ababab",
                 row_index_foreground = "black",
                 row_index_select_background = "#e8eaed",
                 row_index_select_foreground = "black",
                 top_left_background = "white",
                 top_left_foreground = "gray85",
                 font = ("Arial", 10, "normal"),
                 header_font = ("Arial", 10, "normal"),
                 align = "w",
                 header_align = "center",
                 row_index_align = "center",
                 frame_background = "white",
                 table_background = "white",
                 grid_color = "#d4d4d4",
                 text_color = "black",
                 show_selected_cells_border = True,
                 selected_cells_border_color = "#1a73e8",
                 selected_cells_background = "#e7f0fd",
                 selected_cells_foreground = "black",
                 resizing_line_color = "black",
                 drag_and_drop_color = "turquoise1",
                 displayed_columns = [],
                 all_columns_displayed = True,
                 outline_thickness = 0,
                 outline_color = "gray2",
                 theme = "light"):
        tk.Frame.__init__(self,
                          C,
                          background = frame_background,
                          highlightthickness = outline_thickness,
                          highlightbackground = outline_color)
        self.C = C
        if width is not None or height is not None:
            self.grid_propagate(0)
        if width is not None:
            self.config(width = width)
        if height is not None:
            self.config(height = height)
        self.grid_columnconfigure(1, weight = 1)
        self.grid_rowconfigure(1, weight = 1)
        self.RI = RowIndex(self,
                           max_rh = max_rh,
                           max_row_width = max_row_width,
                           row_index_width = row_index_width,
                           row_index_align = row_index_align,
                           row_index_background = row_index_background,
                           row_index_border_color = row_index_border_color,
                           row_index_grid_color = row_index_grid_color,
                           row_index_foreground = row_index_foreground,
                           row_index_select_background = row_index_select_background,
                           row_index_select_foreground = row_index_select_foreground,
                           drag_and_drop_color = drag_and_drop_color,
                           resizing_line_color = resizing_line_color)
        self.CH = ColumnHeaders(self,
                                max_colwidth = max_colwidth,
                                max_header_height = max_header_height,
                                default_header = default_header,
                                header_align = header_align,
                                header_background = header_background,
                                header_border_color = header_border_color,
                                header_grid_color = header_grid_color,
                                header_foreground = header_foreground,
                                header_select_background = header_select_background,
                                header_select_foreground = header_select_foreground,
                                drag_and_drop_color = drag_and_drop_color,
                                resizing_line_color = resizing_line_color)
        self.MT = MainTable(self,
                            column_width = column_width,
                            row_height = row_height,
                            column_headers_canvas = self.CH,
                            row_index_canvas = self.RI,
                            headers = headers,
                            header_height = header_height,
                            data_reference = data_reference,
                            total_cols = total_columns,
                            total_rows = total_rows,
                            row_index = row_index,
                            font = font,
                            header_font = header_font,
                            align = align,
                            table_background = table_background,
                            grid_color = grid_color,
                            text_color= text_color,
                            show_selected_cells_border = show_selected_cells_border,
                            selected_cells_border_color = selected_cells_border_color,
                            selected_cells_background = selected_cells_background,
                            selected_cells_foreground = selected_cells_foreground,
                            displayed_columns = displayed_columns,
                            all_columns_displayed = all_columns_displayed)
        self.TL = TopLeftRectangle(parentframe = self,
                                   main_canvas = self.MT,
                                   row_index_canvas = self.RI,
                                   header_canvas = self.CH,
                                   background = top_left_background,
                                   foreground = top_left_foreground)
        if theme != "light":
            self.change_theme(theme)
        self.yscroll = ttk.Scrollbar(self, command = self.MT.set_yviews, orient = "vertical")
        self.xscroll = ttk.Scrollbar(self, command = self.MT.set_xviews, orient = "horizontal")
        self.MT["xscrollcommand"] = self.xscroll.set
        self.MT["yscrollcommand"] = self.yscroll.set
        self.CH["xscrollcommand"] = self.xscroll.set
        self.RI["yscrollcommand"] = self.yscroll.set
        if show_top_left:
            self.TL.grid(row = 0, column = 0)
        if show_row_index:
            self.RI.grid(row = 1, column = 0, sticky = "nswe")
        if show_header:
            self.CH.grid(row = 0, column = 1, sticky = "nswe")
        if show:
            self.MT.grid(row = 1, column = 1, sticky = "nswe")
            self.yscroll.grid(row = 1, column = 2, sticky = "nswe")
            self.xscroll.grid(row = 2, column = 1, columnspan = 2, sticky = "nswe")
        self.MT.update()

    def height_and_width(self, height = None, width = None):
        if width is not None or height is not None:
            self.grid_propagate(0)
        elif width is None and height is None:
            self.grid_propagate(1)
        if width is not None:
            self.config(width = width)
        if height is not None:
            self.config(height = height)

    def focus_set(self, canvas = "table"):
        if canvas == "table":
            self.MT.focus_set()
        elif canvas == "header":
            self.CH.focus_set()
        elif canvas == "index":
            self.RI.focus_set()
        elif canvas == "topleft":
            self.TL.focus_set()

    def extra_bindings(self, bindings):
        for binding, func in bindings:
            if binding == "ctrl_c":
                self.MT.extra_ctrl_c_func = func
            if binding == "ctrl_x":
                self.MT.extra_ctrl_x_func = func
            if binding == "ctrl_v":
                self.MT.extra_ctrl_v_func = func
            if binding == "ctrl_z":
                self.MT.extra_ctrl_z_func = func
            if binding == "delete_key":
                self.MT.extra_delete_key_func = func
            if binding == "edit_cell":
                self.MT.extra_edit_cell_func = func
            if binding == "row_index_drag_drop":
                self.RI.ri_extra_drag_drop_func = func
            if binding == "column_header_drag_drop":
                self.CH.ch_extra_drag_drop_func = func
            if binding == "cell_select":
                self.MT.selection_binding_func = func
            if binding == "ctrl_a":
                self.MT.select_all_binding_func = func
            if binding == "row_select":
                self.RI.selection_binding_func = func
            if binding == "column_select":
                self.CH.selection_binding_func = func
            if binding == "drag_select_cells":
                self.MT.drag_selection_binding_func = func
            if binding == "drag_select_rows":
                self.RI.drag_selection_binding_func = func
            if binding == "drag_select_columns":
                self.CH.drag_selection_binding_func = func
            if binding == "shift_cell_select":
                self.MT.shift_selection_binding_func = func
            if binding == "shift_row_select":
                self.RI.shift_selection_binding_func = func
            if binding == "shift_column_select":
                self.CH.shift_selection_binding_func = func
            if binding == "deselect":
                self.MT.deselection_binding_func = func

    def bind(self, binding, func):
        if binding == "<ButtonPress-1>":
            self.MT.extra_b1_press_func = func
            self.CH.extra_b1_press_func = func
            self.RI.extra_b1_press_func = func
            self.TL.extra_b1_press_func = func
        elif binding == "<ButtonMotion-1>":
            self.MT.extra_b1_motion_func = func
            self.CH.extra_b1_motion_func = func
            self.RI.extra_b1_motion_func = func
            self.TL.extra_b1_motion_func = func
        elif binding == "<ButtonRelease-1>":
            self.MT.extra_b1_release_func = func
            self.CH.extra_b1_release_func = func
            self.RI.extra_b1_release_func = func
            self.TL.extra_b1_release_func = func
        elif binding == "<Double-Button-1>":
            self.MT.extra_double_b1_func = func
            self.CH.extra_double_b1_func = func
            self.RI.extra_double_b1_func = func
            self.TL.extra_double_b1_func = func
        elif binding == "<Motion>":
            self.MT.extra_motion_func = func
            self.CH.extra_motion_func = func
            self.RI.extra_motion_func = func
            self.TL.extra_motion_func = func
        else:
            self.MT.bind(binding, func)
            self.CH.bind(binding, func)
            self.RI.bind(binding, func)
            self.TL.bind(binding, func)

    def unbind(self, binding):
        if binding == "<ButtonPress-1>":
            self.MT.extra_b1_press_func = None
            self.CH.extra_b1_press_func = None
            self.RI.extra_b1_press_func = None
            self.TL.extra_b1_press_func = None
        elif binding == "<ButtonMotion-1>":
            self.MT.extra_b1_motion_func = None
            self.CH.extra_b1_motion_func = None
            self.RI.extra_b1_motion_func = None
            self.TL.extra_b1_motion_func = None
        elif binding == "<ButtonRelease-1>":
            self.MT.extra_b1_release_func = None
            self.CH.extra_b1_release_func = None
            self.RI.extra_b1_release_func = None
            self.TL.extra_b1_release_func = None
        elif binding == "<Double-Button-1>":
            self.MT.extra_double_b1_func = None
            self.CH.extra_double_b1_func = None
            self.RI.extra_double_b1_func = None
            self.TL.extra_double_b1_func = None
        elif binding == "<Motion>":
            self.MT.extra_motion_func = None
            self.CH.extra_motion_func = None
            self.RI.extra_motion_func = None
            self.TL.extra_motion_func = None
        else:
            self.MT.unbind(binding)
            self.CH.unbind(binding)
            self.RI.unbind(binding)
            self.TL.unbind(binding)

    def enable_bindings(self, bindings):
        self.MT.enable_bindings(bindings)

    def disable_bindings(self, bindings):
        self.MT.disable_bindings(bindings)

    def basic_bindings(self, enable = False):
        if enable:
            for canvas in (self.MT, self.CH, self.RI, self.TL):
                canvas.basic_bindings(onoff = "enable")
        elif not enable:
            for canvas in (self.MT, self.CH, self.RI, self.TL):
                canvas.basic_bindings(onoff = "disable")

    def edit_bindings(self, enable = False):
        if enable:
            self.MT.edit_bindings(True)
        elif not enable:
            self.MT.edit_bindings(False)

    def cell_edit_binding(self, enable = False):
        self.MT.bind_cell_edit(enable)

    def identify_region(self, event):
        # UNFINISHED, ADD SEPARATOR?
        if event.widget == self.MT:
            return "table"
        elif event.widget == self.RI:
            return "index"
        elif event.widget == self.CH:
            return "header"
        elif event.widget == self.TL:
            return "top left"

    def identify_row(self, event, exclude_index = False, allow_end = True):
        ev_w = event.widget
        if ev_w == self.MT:
            return self.MT.identify_row(y = event.y, allow_end = allow_end)
        elif ev_w == self.RI:
            if exclude_index:
                return None
            else:
                return self.MT.identify_row(y = event.y, allow_end = allow_end)
        elif ev_w == self.CH or ev_w == self.TL:
            return None

    def identify_column(self, event, exclude_header = False, allow_end = True):
        ev_w = event.widget
        if ev_w == self.MT:
            return self.MT.identify_col(x = event.x, allow_end = allow_end)
        elif ev_w == self.RI or ev_w == self.TL:
            return None
        elif ev_w == self.CH:
            if exclude_header:
                return None
            else:
                return self.MT.identify_col(x = event.x, allow_end = allow_end)

    def get_example_canvas_column_widths(self, total_cols = None):
        colpos = int(self.MT.default_cw)
        if total_cols is not None:
            return [0] + list(accumulate(colpos for c in range(total_cols)))
        return [0] + list(accumulate(colpos for c in range(len(self.MT.col_positions) - 1)))

    def get_example_canvas_row_heights(self, total_rows = None):
        rowpos = self.MT.GetLinesHeight(self.MT.default_rh)
        if total_rows is not None:
            return [0] + list(accumulate(rowpos for c in range(total_rows)))
        return [0] + list(accumulate(rowpos for c in range(len(self.MT.row_positions) - 1)))

    def get_column_widths(self, canvas_positions = False):
        if canvas_positions:
            return [int(n) for n in self.MT.col_positions]
        return [int(b - a) for a, b in zip(self.MT.col_positions, islice(self.MT.col_positions, 1, len(self.MT.col_positions)))]
    
    def get_row_heights(self, canvas_positions = False):
        if canvas_positions:
            return [int(n) for n in self.MT.row_positions]
        return [int(b - a) for a, b in zip(self.MT.row_positions, islice(self.MT.row_positions, 1, len(self.MT.row_positions)))]

    def column_width(self, column = None, width = None, only_set_if_too_small = False, redraw = True):
        if column == "all":
            if width == "default":
                self.MT.reset_col_positions()
        elif column == "displayed":
            if width == "text":
                sc, ec = self.MT.get_visible_columns(self.MT.canvasx(0), self.MT.canvasx(self.winfo_width()))
                for c in range(sc, ec - 1):
                    self.CH.set_col_width(c)
        elif width is not None and column is not None:
            self.CH.set_col_width(col = column, width = width, only_set_if_too_small = only_set_if_too_small)
        elif column is not None:
            return int(self.MT.col_positions[column + 1] - self.MT.col_positions[column])
        if redraw:
            self.refresh()

    def set_column_widths(self, column_widths = None, canvas_positions = False, reset = False, verify = False):
        cwx = None
        if reset:
            self.MT.reset_col_positions()
            return
        if verify:
            cwx = self.verify_column_widths(column_widths, canvas_positions)
        if isinstance(column_widths, list):
            if canvas_positions:
                self.MT.col_positions = column_widths
            else:
                self.MT.col_positions = [0] + list(accumulate(width for width in column_widths))
        return cwx
            
    def row_height(self, row = None, height = None, only_set_if_too_small = False, redraw = True):
        if row == "all":
            if height == "default":
                self.MT.reset_row_positions()
        elif row == "displayed":
            if height == "text":
                sr, er = self.MT.get_visible_rows(self.MT.canvasy(0), self.MT.canvasy(self.winfo_width()))
                for r in range(sr, er - 1):
                    self.RI.set_row_height(r)
        elif height is not None and row is not None:
            self.RI.set_row_height(row = row, height = height, only_set_if_too_small = only_set_if_too_small)
        elif row is not None:
            return int(self.MT.row_positions[row + 1] - self.MT.row_positions[row])
        if redraw:
            self.refresh()

    def set_row_heights(self, row_heights = None, canvas_positions = False, reset = False, verify = False):
        rhx = None
        if reset:
            self.MT.reset_row_positions()
            return
        if verify:
            rhx = self.verify_row_heights(row_heights, canvas_positions)
        if isinstance(row_heights, list):
            if canvas_positions:
                self.MT.row_positions = row_heights
            else:
                self.MT.row_positions = [0] + list(accumulate(height for height in row_heights))
        return rhx

    def verify_row_heights(self, row_heights, canvas_positions = False):
        if row_heights[0] != 0 or isinstance(row_heights[0],bool):
            return False
        if not isinstance(row_heights,list):
            return False
        if canvas_positions:
            if any(x - z < self.MT.min_rh or not isinstance(x, int) or isinstance(x, bool) for z, x in zip(islice(row_heights, 0, None), islice(row_heights, 1, None))):
                return False
        elif not canvas_positions:
            if any(z < self.MT.min_rh or not isinstance(z, int) or isinstance(z, bool) for z in row_heights):
                return False
        return True

    def verify_column_widths(self, column_widths, canvas_positions = False):
        if column_widths[0] != 0 or isinstance(column_widths[0],bool):
            return False
        if not isinstance(column_widths,list):
            return False
        if canvas_positions:
            if any(x - z < self.MT.min_cw or not isinstance(x, int) or isinstance(x, bool) for z, x in zip(islice(column_widths, 0, None), islice(column_widths, 1, None))):
                return False
        elif not canvas_positions:
            if any(z < self.MT.min_cw or not isinstance(z, int) or isinstance(z, bool) for z in column_widths):
                return False
        return True
                
    def default_row_height(self, height = None):
        if height is not None:
            self.MT.default_rh = int(height)
        return self.MT.default_rh

    def default_header_height(self, height = None):
        if height is not None:
            self.MT.default_hh = int(height)
        return self.MT.default_hh

    def create_dropdown(self,
                        r = 0,
                        c = 0,
                        values = [],
                        set_value = None,
                        state = "readonly",
                        see = True,
                        destroy_on_select = True,
                        current = False):
        self.MT.create_dropdown(r = r, c = c, values = values, set_value = set_value, state = state, see = see,
                                destroy_on_select = destroy_on_select, current = current)

    def get_dropdown_value(self, current = False, destroy = True):
        return self.MT.get_dropdown_value(current = current, destroy = destroy)

    def cut(self, event = None):
        self.MT.ctrl_x()

    def copy(self, event = None):
        self.MT.ctrl_c()

    def paste(self, event = None):
        self.MT.ctrl_v()

    def delete(self, event = None):
        self.MT.delete_key()

    def undo(self, event = None):
        self.MT.ctrl_z()

    def delete_row_position(self, idx, deselect_all = False, preserve_other_selections = False):
        self.MT.del_row_position(idx = idx,
                                 deselect_all = deselect_all,
                                 preserve_other_selections = preserve_other_selections)

    def delete_row(self, idx = -1, deselect_all = False, preserve_other_selections = False):
        del self.MT.data_ref[idx]
        self.MT.del_row_position(idx = idx,
                                 deselect_all = deselect_all,
                                 preserve_other_selections = preserve_other_selections)
        self.MT.total_rows -= 1

    def insert_row_position(self, idx, height = None, deselect_all = False, preserve_other_selections = False):
        self.MT.insert_row_position(idx = idx,
                                    height = height,
                                    deselect_all = deselect_all,
                                    preserve_other_selections = preserve_other_selections)

    def insert_row(self, row = None, idx = "end", height = None, deselect_all = False, preserve_other_selections = False):
        self.MT.insert_row(row = row,
                           idx = idx,
                           height = height,
                           deselect_all = deselect_all,
                           preserve_other_selections = preserve_other_selections)

    def total_rows(self, number = None, mod_positions = True, mod_data = True):
        if number is None:
            return int(self.MT.total_rows)
        if not isinstance(number, int) or number < 0:
            raise ValueError("number argument must be integer and > 0")
        total_cols = int(self.MT.total_cols)
        if number > len(self.MT.data_ref):
            if mod_positions:
                if mod_data:
                    height = self.MT.GetLinesHeight(self.MT.default_rh)
                    for r in range(number - len(self.MT.data_ref)):
                        self.MT.data_ref.append(list(repeat("", total_cols)))
                        self.MT.insert_row_position("end", height)
                else:
                    height = self.MT.GetLinesHeight(self.MT.default_rh)
                    for r in range(number - len(self.MT.data_ref)):
                        self.MT.insert_row_position("end", height)
            else:
                self.MT.data_ref.extend([list(repeat("", total_cols)) for r in range(number - len(self.MT.data_ref))])
        elif number < len(self.MT.data_ref):
            #here work needed if display subset of rows is implemented
            if mod_data:
                self.MT.data_ref[number:] = []
            self.MT.row_positions[number + 1:] = []
        self.MT.total_rows = len(self.MT.data_ref)  

    def total_columns(self, number = None, mod_positions = True, mod_data = True):
        if number is None:
            return int(self.MT.total_cols)
        if not isinstance(number, int) or number < 0:
            raise ValueError("number argument must be integer and > 0")
        total_cols = int(self.MT.total_cols)
        if number > total_cols:
            if mod_positions:
                width = self.MT.default_cw
                for c in range(number - total_cols):
                    self.MT.insert_col_position("end", width)
                if mod_data:
                    for r in range(len(self.MT.data_ref)):
                        self.MT.data_ref[r].extend(list(repeat("", number - total_cols)))
            else:
                self.MT.data_ref.extend([list(repeat("", total_cols)) for r in range(number - len(self.MT.data_ref))])
        elif number < total_cols:
            if not self.MT.all_columns_displayed:
                self.MT.display_columns(enable = False, reset_col_positions = False, deselect_all = True)
            if mod_data:
                for r in range(len(self.MT.data_ref)):
                    try:
                        self.MT.data_ref[r][number:] = []
                    except:
                        continue
            self.MT.col_positions[number + 1:] = []
        self.MT.total_cols = int(number)

    def move_row_position(self, row, moveto):
        self.MT.move_row_position(row, moveto)

    def move_row(self, row, moveto):
        self.MT.move_row_position(row, moveto)
        self.MT.data_ref.insert(moveto, self.MT.data_ref.pop(row))

    def delete_column_position(self, idx, deselect_all = False, preserve_other_selections = False):
        self.MT.del_col_position(idx,
                                 deselect_all = deselect_all,
                                 preserve_other_selections = preserve_other_selections)

    def delete_column(self, idx = -1, deselect_all = False, preserve_other_selections = False):
        for rn in range(len(self.MT.data_ref)):
            del self.MT.data_ref[rn][idx] 
        self.MT.del_col_position(idx,
                                 deselect_all = deselect_all,
                                 preserve_other_selections = preserve_other_selections)
        self.MT.total_cols -= 1

    def insert_column_position(self, idx, width = None, deselect_all = False, preserve_other_selections = False):
        self.MT.insert_col_position(idx = idx,
                                    width = width,
                                    deselect_all = deselect_all,
                                    preserve_other_selections = preserve_other_selections)

    def insert_column(self, column = None, idx = "end", width = None, deselect_all = False, preserve_other_selections = False):
        self.MT.insert_col(column = column,
                           idx = idx,
                           width = width,
                           deselect_all = deselect_all,
                           preserve_other_selections = preserve_other_selections)

    def move_column_position(self, column, moveto):
        self.MT.move_col_position(column, moveto)

    def move_column(self, column, moveto):
        self.MT.move_col_position(column, moveto)
        for rn in range(len(self.MT.data_ref)):
            self.MT.data_ref[rn].insert(moveto, self.MT.data_ref[rn].pop(column))

    def create_text_editor(self, row = 0, column = 0, text = None, state = "normal", see = True, set_data_ref_on_destroy = False):
        self.MT.create_text_editor(r = row, c = column, text = text, state = state, see = see, set_data_ref_on_destroy = set_data_ref_on_destroy)

    def get_xview(self):
        return self.MT.xview()

    def get_yview(self):
        return self.MT.yview()

    def set_xview(self, position, option = "moveto"):
        self.MT.set_xviews(option, position)

    def set_yview(self,position,option = "moveto"):
        self.MT.set_yviews(option, position)

    def set_view(self, x_args, y_args):
        self.MT.set_view(x_args, y_args)

    def see(self, row = 0, column = 0, keep_yscroll = False, keep_xscroll = False, bottom_right_corner = False, check_cell_visibility = True):
        self.MT.see(row, column, keep_yscroll, keep_xscroll, bottom_right_corner, check_cell_visibility = check_cell_visibility)

    def select_row(self, row, redraw = True):
        self.RI.select_row(row, redraw = redraw)

    def select_column(self, column, redraw = True):
        self.CH.select_col(column, redraw = redraw)

    def select_cell(self, row, column, redraw = True):
        self.MT.select_cell(row, column, redraw = redraw)

    def add_cell_selection(self, row, column, redraw = True, run_binding_func = True, set_as_current = True):
        self.MT.add_selection(r = row, c = column, redraw = redraw, run_binding_func = run_binding_func, set_as_current = set_as_current)

    def add_row_selection(self, row, redraw = True, run_binding_func = True, set_as_current = True):
        self.RI.add_selection(r = row, redraw = redraw, run_binding_func = run_binding_func, set_as_current = set_as_current)

    def add_column_selection(self, column, redraw = True, run_binding_func = True, set_as_current = True):
        self.CH.add_selection(c = column, redraw = redraw, run_binding_func = run_binding_func, set_as_current = set_as_current)

    def toggle_select_cell(self, row, column, add_selection = True, redraw = True, run_binding_func = True, set_as_current = True):
        self.MT.toggle_select_cell(row = row, column = column, add_selection = add_selection, redraw = redraw, run_binding_func = run_binding_func, set_as_current = set_as_current)

    def toggle_select_row(self, row, add_selection = True, redraw = True, run_binding_func = True, set_as_current = True):
        self.RI.toggle_select_row(row = row, add_selection = add_selection, redraw = redraw, run_binding_func = run_binding_func, set_as_current = set_as_current)

    def toggle_select_column(self, column, add_selection = True, redraw = True, run_binding_func = True, set_as_current = True):
        self.CH.toggle_select_col(column = column, add_selection = add_selection, redraw = redraw, run_binding_func = run_binding_func, set_as_current = set_as_current)

    def deselect(self, row = None, column = None, cell = None, redraw = True):
        self.MT.deselect(r = row, c = column, cell = cell, redraw = redraw)

    def get_selected_rows(self, get_cells = False):
        return self.MT.get_selected_rows(get_cells = get_cells)

    def get_selected_columns(self, get_cells = False):
        return self.MT.get_selected_cols(get_cells = get_cells)

    def get_selected_cells(self, get_rows = False, get_cols = False):
        return self.MT.get_selected_cells(get_rows = get_rows, get_cols = get_cols)

    def get_all_selection_boxes(self):
        return self.MT.get_all_selection_boxes()

    def is_cell_selected(self, r, c):
        return self.MT.is_cell_selected(r, c)

    def is_row_selected(self, r):
        return self.MT.is_row_selected(r)

    def is_column_selected(self, c):
        return self.MT.is_col_selected(c)

    def anything_selected(self, exclude_columns = False, exclude_rows = False, exclude_cells = False):
        return self.MT.anything_selected(exclude_columns = exclude_columns, exclude_rows = exclude_rows, exclude_cells = exclude_cells)

    def highlight_cells(self, row = 0, column = 0, cells = [], canvas = "table", bg = None, fg = None, redraw = False):
        if canvas == "table":
            self.MT.highlight_cells(r = row,
                                    c = column,
                                    cells = cells,
                                    bg = bg,
                                    fg = fg,
                                    redraw = redraw)
        elif canvas == "row_index":
            self.RI.highlight_cells(r = row,
                                    cells = cells,
                                    bg = bg,
                                    fg = fg,
                                    redraw = redraw)
        elif canvas == "header":
            self.CH.highlight_cells(c = column,
                                    cells = cells,
                                    bg = bg,
                                    fg = fg,
                                    redraw = redraw)

    def dehighlight_cells(self, row = 0, column = 0, cells = [], canvas = "table", all_ = False, redraw = True):
        if canvas == "table":
            if cells and not all_:
                for t in cells:
                    try:
                        del self.MT.highlighted_cells[t]
                    except:
                        pass
            elif not all_:
                if (row, column) in self.MT.highlighted_cells:
                    del self.MT.highlighted_cells[(row, column)]
            elif all_:
                self.MT.highlighted_cells = {}
        elif canvas == "row_index":
            if cells and not all_:
                for r in cells:
                    try:
                        del self.RI.highlighted_cells[r]
                    except:
                        pass
            elif not all_:
                if row in self.RI.highlighted_cells:
                    del self.RI.highlighted_cells[row]
            elif all_:
                self.RI.highlighted_cells = {}
        elif canvas == "header":
            if cells and not all_:
                for c in cells:
                    try:
                        del self.CH.highlighted_cells[c]
                    except:
                        pass
            elif not all_:
                if column in self.CH.highlighted_cells:
                    del self.CH.highlighted_cells[column]
            elif all_:
                self.CH.highlighted_cells = {}
        if redraw:
            self.refresh(True, True)

    def get_highlighted_cells(self, canvas = "table"):
        if canvas == "table":
            return self.MT.highlighted_cells
        elif canvas == "row_index":
            return self.RI.highlighted_cells
        elif canvas == "header":
            return self.CH.highlighted_cells
        
    def get_frame_y(self, y):
        return y + self.CH.current_height

    def get_frame_x(self, x):
        return x + self.RI.current_width

    def align(self, align = None, redraw = True):
        if align is None:
            return self.MT.align
        elif align in ("w", "center"):
            self.MT.align = align
        else:
            raise ValueError("Align must be either 'w' or 'center'")
        if redraw:
            self.refresh()

    def header_align(self, align = None, redraw = True):
        if align is None:
            return self.CH.align
        elif align in ("w", "center"):
            self.CH.align = align
        else:
            raise ValueError("Align must be either 'w' or 'center'")
        if redraw:
            self.refresh()

    def row_index_align(self, align = None, redraw = True):
        if align is None:
            return self.RI.align
        elif align in ("w", "center"):
            self.RI.align = align
        else:
            raise ValueError("Align must be either 'w' or 'center'")
        if redraw:
            self.refresh()

    def font(self, newfont = None):
        self.MT.font(newfont)

    def header_font(self, newfont = None):
        self.MT.header_font(newfont)

    def set_options(self,
                    font = None,
                    default_header = "letters",
                    header_font = None,
                    show_selected_cells_border = None,
                    theme = None,
                    max_colwidth = None,
                    max_row_height = None,
                    max_header_height = None,
                    max_row_width = None,
                    header_background = None,
                    header_border_color = None,
                    header_grid_color = None,
                    header_foreground = None,
                    header_select_background = None,
                    header_select_foreground = None,
                    row_index_background = None,
                    row_index_border_color = None,
                    row_index_grid_color = None,
                    row_index_foreground = None,
                    row_index_select_background = None,
                    row_index_select_foreground = None,
                    top_left_background = None,
                    top_left_foreground = None,
                    frame_background = None,
                    table_background = None,
                    grid_color = None,
                    text_color = None,
                    selected_cells_border_color = None,
                    selected_cells_background = None,
                    selected_cells_foreground = None,
                    resizing_line_color = None,
                    drag_and_drop_color = None,
                    outline_thickness = None,
                    outline_color = None,
                    redraw = True):
        if default_header is not None:
            self.CH.default_hdr = 1 if default_header.lower() == "letters" else 0
        if max_colwidth is not None:
            self.CH.max_cw = float(max_colwidth)
        if max_row_height is not None:
            self.RI.max_rh = float(max_row_height)
        if max_header_height is not None:
            self.CH.max_header_height = float(max_header_height)
        if max_row_width is not None:
            self.RI.max_row_width = float(max_row_width)
        if font is not None:
            self.MT.font(newfont)
        if header_font is not None:
            self.MT.header_font(newfont)
        if theme is not None:
            self.change_theme(theme)
        if show_selected_cells_border is not None:
            self.MT.show_selected_cells_border = show_selected_cells_border
        if header_background is not None:
            self.CH.config(background = header_background)
        if header_border_color is not None:
            self.CH.header_border_color = header_border_color
        if header_grid_color is not None:
            self.CH.grid_color = header_grid_color
        if header_foreground is not None:
            self.CH.text_color = header_foreground
        if header_select_background is not None:
            self.CH.selected_cells_background = header_select_background
        if header_select_foreground is not None:
            self.CH.selected_cells_foreground = header_select_foreground
        if row_index_background is not None:
            self.RI.config(background=row_index_background)
        if row_index_border_color is not None:
            self.RI.row_index_border_color = row_index_border_color
        if row_index_grid_color is not None:
            self.RI.grid_color = row_index_grid_color
        if row_index_foreground is not None:
            self.RI.text_color = row_index_foreground
        if row_index_select_background is not None:
            self.RI.selected_cells_background = row_index_select_background
        if row_index_select_foreground is not None:
            self.RI.selected_cells_foreground = row_index_select_foreground
        if top_left_background is not None:
            self.TL.config(background=top_left_background)
        if top_left_foreground is not None:
            self.TL.rectangle_foreground = top_left_foreground
            self.TL.itemconfig("rw", fill = top_left_foreground)
            self.TL.itemconfig("rh", fill = top_left_foreground)
        if frame_background is not None:
            self.config(background = frame_background)
        if table_background is not None:
            self.MT.config(background = table_background)
        if grid_color is not None:
            self.MT.grid_color = grid_color
        if text_color is not None:
            self.MT.text_color = text_color
        if selected_cells_border_color is not None:
            self.MT.selected_cells_border_color = selected_cells_border_color
        if selected_cells_background is not None:
            self.MT.selected_cells_background = selected_cells_background
        if selected_cells_foreground is not None:
            self.MT.selected_cells_foreground = selected_cells_foreground
        if resizing_line_color is not None:
            self.CH.resizing_line_color = resizing_line_color
            self.RI.resizing_line_color = resizing_line_color
        if drag_and_drop_color is not None:
            self.CH.drag_and_drop_color = drag_and_drop_color
            self.RI.drag_and_drop_color = drag_and_drop_color
        if outline_thickness is not None:
            self.config(highlightthickness = outline_thickness)
        if outline_color is not None:
            self.config(highlightbackground = outline_color)
        if redraw:
            self.refresh()

    def change_color(self,
                     header_background = None,
                     header_border_color = None,
                     header_grid_color = None,
                     header_foreground = None,
                     header_select_background = None,
                     header_select_foreground = None,
                     row_index_background = None,
                     row_index_border_color = None,
                     row_index_grid_color = None,
                     row_index_foreground = None,
                     row_index_select_background = None,
                     row_index_select_foreground = None,
                     top_left_background = None,
                     top_left_foreground = None,
                     frame_background = None,
                     table_background = None,
                     grid_color = None,
                     text_color = None,
                     selected_cells_border_color = None,
                     selected_cells_background = None,
                     selected_cells_foreground = None,
                     resizing_line_color = None,
                     drag_and_drop_color = None,
                     outline_thickness = None,
                     outline_color = None,
                     redraw = False):
        if header_background is not None:
            self.CH.config(background = header_background)
        if header_border_color is not None:
            self.CH.header_border_color = header_border_color
        if header_grid_color is not None:
            self.CH.grid_color = header_grid_color
        if header_foreground is not None:
            self.CH.text_color = header_foreground
        if header_select_background is not None:
            self.CH.selected_cells_background = header_select_background
        if header_select_foreground is not None:
            self.CH.selected_cells_foreground = header_select_foreground
        if row_index_background is not None:
            self.RI.config(background=row_index_background)
        if row_index_border_color is not None:
            self.RI.row_index_border_color = row_index_border_color
        if row_index_grid_color is not None:
            self.RI.grid_color = row_index_grid_color
        if row_index_foreground is not None:
            self.RI.text_color = row_index_foreground
        if row_index_select_background is not None:
            self.RI.selected_cells_background = row_index_select_background
        if row_index_select_foreground is not None:
            self.RI.selected_cells_foreground = row_index_select_foreground
        if top_left_background is not None:
            self.TL.config(background=top_left_background)
        if top_left_foreground is not None:
            self.TL.rectangle_foreground = top_left_foreground
            self.TL.itemconfig("rw", fill = top_left_foreground)
            self.TL.itemconfig("rh", fill = top_left_foreground)
        if frame_background is not None:
            self.config(background = frame_background)
        if table_background is not None:
            self.MT.config(background = table_background)
        if grid_color is not None:
            self.MT.grid_color = grid_color
        if text_color is not None:
            self.MT.text_color = text_color
        if selected_cells_border_color is not None:
            self.MT.selected_cells_border_color = selected_cells_border_color
        if selected_cells_background is not None:
            self.MT.selected_cells_background = selected_cells_background
        if selected_cells_foreground is not None:
            self.MT.selected_cells_foreground = selected_cells_foreground
        if resizing_line_color is not None:
            self.CH.resizing_line_color = resizing_line_color
            self.RI.resizing_line_color = resizing_line_color
        if drag_and_drop_color is not None:
            self.CH.drag_and_drop_color = drag_and_drop_color
            self.RI.drag_and_drop_color = drag_and_drop_color
        if outline_thickness is not None:
            self.config(highlightthickness = outline_thickness)
        if outline_color is not None:
            self.config(highlightbackground = outline_color)
        if redraw:
            self.refresh()

    def change_theme(self, theme = "light"):
        if theme == "light":
            self.change_color(header_background = "#f8f9fa",
                                header_border_color = "#ababab",
                                header_grid_color = "#ababab",
                                header_foreground = "black",
                                header_select_background = "#e8eaed",
                                header_select_foreground = "black",
                                row_index_background = "#f8f9fa",
                                row_index_border_color = "#ababab",
                                row_index_grid_color = "#ababab",
                                row_index_foreground = "black",
                                row_index_select_background = "#e8eaed",
                                row_index_select_foreground = "black",
                                top_left_background = "white",
                                top_left_foreground = "gray85",
                                table_background = "white",
                                grid_color = "#d4d4d4",
                                text_color = "black",
                              selected_cells_border_color = "#1a73e8",
                                selected_cells_background = "#e7f0fd",
                                selected_cells_foreground = "black",
                                resizing_line_color = "black",
                                drag_and_drop_color = "turquoise1",
                                outline_color = "gray2",
                                redraw = True)
            self.config(bg = "white")
        elif theme == "dark":
            self.change_color(header_background = "black",
                                header_border_color = "#353a41",
                                header_grid_color = "#353a41",
                                header_foreground = "gray95",
                                header_select_background = "gray95",
                                header_select_foreground = "black",
                                row_index_background = "black",
                                row_index_border_color = "#353a41",
                                row_index_grid_color = "#353a41",
                                row_index_foreground = "gray95",
                                row_index_select_background = "gray95",
                                row_index_select_foreground = "black",
                                top_left_background = "#353a41",
                                top_left_foreground = "#222222",
                                table_background = "#222222",
                                grid_color = "#353a41",
                                text_color = "gray95",
                              selected_cells_border_color = "#1a73e8",
                                selected_cells_background = "gray95",
                                selected_cells_foreground = "black",
                                resizing_line_color = "red",
                                drag_and_drop_color = "#9acd32",
                                outline_color = "gray2",
                                redraw = True)
            self.config(bg = "#222222")
            
    def data_reference(self,
                       newdataref = None,
                       total_cols = None,
                       total_rows = None,
                       reset_col_positions = True,
                       reset_row_positions = True,
                       redraw = False):
        return self.MT.data_reference(newdataref,
                                      total_cols,
                                      total_rows,
                                      reset_col_positions,
                                      reset_row_positions,
                                      redraw)

    def get_cell_data(self, r, c, return_copy = True):
        if return_copy:
            try:
                return f"{self.MT.data_ref[r][c]}"
            except:
                return None
        else:
            try:
                return self.MT.data_ref[r][c]
            except:
                return None

    def get_row_data(self, r, return_copy = True):
        if return_copy:
            try:
                return tuple(f"{e}" for e in self.MT.data_ref[r])
            except:
                return None
        else:
            try:
                self.MT.data_ref[r]
            except:
                return None

    def get_column_data(self, c, return_copy = True):
        res = []
        if return_copy:
            for r in self.MT.data_ref:
                try:
                    res.append(f"{r[c]}")
                except:
                    continue
            return tuple(res)
        else:
            for r in self.MT.data_ref:
                try:
                    res.append(r[c])
                except:
                    continue
            return res

    def set_cell_data(self, r, c, value = ""):
        self.MT.data_ref[r][c] = f"{value}"

    def set_row_data(self, r, values = tuple(), extend_row = True):
        if len(self.MT.data_ref) - 1 < r:
            raise Exception("Row number is out of range")
        if extend_row:
            maxidx = len(self.MT.data_ref[r]) - 1
            for c, v in enumerate(values):
                if c > maxidx:
                    self.MT.data_ref[r].append(v)
                else:
                    self.MT.data_ref[r][c] = v
            if c + 1 > self.MT.total_cols:
                new_total = c + 1
                self.total_columns(new_total)
        else:
            for c, v in enumerate(values):
                if c > maxidx:
                    self.MT.data_ref[r].append(v)
                else:
                    self.MT.data_ref[r][c] = v

    def set_column_data(self, c, values = tuple(), extend_column = True):
        if c > self.MT.total_cols - 1:
            raise Exception("Column number is out of range")
        if extend_column:
            maxidx = int(self.MT.total_rows) - 1
            for rn, v in enumerate(values):
                if rn > maxidx:
                    self.insert_row()
                    maxidx = int(self.MT.total_rows) - 1
                self.MT.data_ref[rn][c] = v
        else:
            for rn, v in enumerate(values):
                self.MT.data_ref[rn][c] = v
                    
    def display_subset_of_columns(self,
                                  indexes = None,
                                  enable = None,
                                  reset_col_positions = True,
                                  refresh = False,
                                  deselect_all = True):
        res = self.MT.display_columns(indexes = indexes,
                                      enable = enable,
                                      reset_col_positions = reset_col_positions,
                                      deselect_all = deselect_all)
        if refresh:
            self.refresh()
        return res

    def show_ctrl_outline(self, canvas = "table", start_cell = (0, 0), end_cell = (1, 1)):
        self.MT.show_ctrl_outline(canvas = canvas, start_cell = start_cell, end_cell = end_cell)

    def get_selected_min_max(self): # returns (min_y, min_x, max_y, max_x) of any selections including rows/columns
        return self.MT.get_selected_min_max()
        
    def headers(self, newheaders = None, index = None):
        return self.MT.headers(newheaders, index)
        
    def row_index(self, newindex = None, index = None):
        return self.MT.row_index(newindex,index)

    def reset_undos(self):
        self.MT.undo_storage = deque(maxlen = 20)

    def show(self, canvas = "all"):
        if canvas == "all":
            self.TL.grid(row = 0, column = 0)
            self.RI.grid(row = 1, column = 0, sticky = "nswe")
            self.CH.grid(row = 0, column = 1, sticky = "nswe")
            self.MT.grid(row = 1, column = 1, sticky = "nswe")
            self.yscroll.grid(row = 0, column = 2, rowspan = 3, sticky = "nswe")
            self.xscroll.grid(row = 2, column = 1, sticky = "nswe")
        elif canvas == "row_index":
            self.RI.grid(row = 1, column = 0, sticky = "nswe")
        elif canvas == "header":
            self.CH.grid(row = 0, column = 1, sticky = "nswe")
        elif canvas == "top_left":
            self.TL.grid(row = 0, column = 0)
        self.MT.update()

    def hide(self, canvas = "all"):
        if canvas.lower() == "all":
            self.TL.grid_forget()
            self.RI.grid_forget()
            self.CH.grid_forget()
            self.MT.grid_forget()
            self.yscroll.grid_forget()
            self.xscroll.grid_forget()
        elif canvas == "row_index":
            self.RI.grid_forget()
        elif canvas == "header":
            self.CH.grid_forget()
        elif canvas == "top_left":
            self.TL.grid_forget()

    def refresh(self, redraw_header = True, redraw_row_index = True):
        self.MT.main_table_redraw_grid_and_text(redraw_header = redraw_header, redraw_row_index = redraw_row_index)






