from ._tksheet_vars import *
from ._tksheet_other_classes import *
from ._tksheet_top_left_rectangle import *
from ._tksheet_column_headers import *
from ._tksheet_row_index import *
from ._tksheet_main_table import *

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


class Sheet(tk.Frame):
    def __init__(self,
                 parent,
                 show_table = True,
                 show_top_left = True,
                 show_row_index = True,
                 show_header = True,
                 show_x_scrollbar = True,
                 show_y_scrollbar = True,
                 width = None,
                 height = None,
                 headers = None,
                 measure_subset_header = True,
                 default_header = "letters", #letters, numbers or both
                 default_row_index = "numbers", #letters, numbers or both
                 page_up_down_select_row = True,
                 arrow_key_down_right_scroll_page = False,
                 enable_edit_cell_auto_resize = True,
                 data_reference = None,
                 data = None,
                 startup_select = None,
                 startup_focus = True,
                 total_columns = None,
                 total_rows = None,
                 column_width = 120,
                 header_height = "1",
                 max_colwidth = "inf",
                 max_rh = "inf",
                 max_header_height = "inf",
                 max_row_width = "inf",
                 row_index = None,
                 measure_subset_index = True,
                 after_redraw_time_ms = 100,
                 row_index_width = 100,
                 auto_resize_default_row_index = True,
                 set_all_heights_and_widths = False,
                 row_height = "1",
                 font = get_font(),
                 header_font = get_heading_font(),
                 popup_menu_font = get_font(),
                 align = "w",
                 header_align = "center",
                 row_index_align = "center",
                 displayed_columns = [],
                 all_columns_displayed = True,
                 max_undos = 20,
                 outline_thickness = 0,
                 outline_color = theme_light_blue['outline_color'],
                 column_drag_and_drop_perform = True,
                 row_drag_and_drop_perform = True,
                 empty_horizontal = 150,
                 empty_vertical = 100,
                 show_vertical_grid = True,
                 show_horizontal_grid = True,
                 display_selected_fg_over_highlights = False,
                 show_selected_cells_border = True,
                 theme                              = "light blue",
                 popup_menu_fg                      = "gray2",
                 popup_menu_bg                      = "#f2f2f2",
                 popup_menu_highlight_bg            = "#91c9f7",
                 popup_menu_highlight_fg            = "black",
                 frame_bg                           = theme_light_blue['table_bg'],
                 table_grid_fg                      = theme_light_blue['table_grid_fg'],
                 table_bg                           = theme_light_blue['table_bg'],
                 table_fg                           = theme_light_blue['table_fg'], 
                 table_selected_cells_border_fg     = theme_light_blue['table_selected_cells_border_fg'],
                 table_selected_cells_bg            = theme_light_blue['table_selected_cells_bg'],
                 table_selected_cells_fg            = theme_light_blue['table_selected_cells_fg'],
                 table_selected_rows_border_fg      = theme_light_blue['table_selected_rows_border_fg'],
                 table_selected_rows_bg             = theme_light_blue['table_selected_rows_bg'],
                 table_selected_rows_fg             = theme_light_blue['table_selected_rows_fg'],
                 table_selected_columns_border_fg   = theme_light_blue['table_selected_columns_border_fg'],
                 table_selected_columns_bg          = theme_light_blue['table_selected_columns_bg'],
                 table_selected_columns_fg          = theme_light_blue['table_selected_columns_fg'],
                 resizing_line_fg                   = theme_light_blue['resizing_line_fg'],
                 drag_and_drop_bg                   = theme_light_blue['drag_and_drop_bg'],
                 index_bg                           = theme_light_blue['index_bg'],
                 index_border_fg                    = theme_light_blue['index_border_fg'],
                 index_grid_fg                      = theme_light_blue['index_grid_fg'],
                 index_fg                           = theme_light_blue['index_fg'],
                 index_selected_cells_bg            = theme_light_blue['index_selected_cells_bg'],
                 index_selected_cells_fg            = theme_light_blue['index_selected_cells_fg'],
                 index_selected_rows_bg             = theme_light_blue['index_selected_rows_bg'],
                 index_selected_rows_fg             = theme_light_blue['index_selected_rows_fg'],
                 index_hidden_rows_expander_bg      = theme_light_blue['index_hidden_rows_expander_bg'],
                 header_bg                          = theme_light_blue['header_bg'],
                 header_border_fg                   = theme_light_blue['header_border_fg'],
                 header_grid_fg                     = theme_light_blue['header_grid_fg'],
                 header_fg                          = theme_light_blue['header_fg'],
                 header_selected_cells_bg           = theme_light_blue['header_selected_cells_bg'],
                 header_selected_cells_fg           = theme_light_blue['header_selected_cells_fg'],
                 header_selected_columns_bg         = theme_light_blue['header_selected_columns_bg'],
                 header_selected_columns_fg         = theme_light_blue['header_selected_columns_fg'],
                 header_hidden_columns_expander_bg  = theme_light_blue['header_hidden_columns_expander_bg'],
                 top_left_bg                        = theme_light_blue['top_left_bg'],
                 top_left_fg                        = theme_light_blue['top_left_fg'],
                 top_left_fg_highlight              = theme_light_blue['top_left_fg_highlight']):
        tk.Frame.__init__(self,
                          parent,
                          background = frame_bg,
                          highlightthickness = outline_thickness,
                          highlightbackground = outline_color)
        self.C = parent
        self.after_redraw_id = None
        self.after_redraw_time_ms = after_redraw_time_ms
        if width is not None or height is not None:
            self.grid_propagate(0)
        if width is not None:
            self.config(width = width)
        if height is not None:
            self.config(height = height)
        if width is not None and height is None:
            self.config(height = 300)
        if height is not None and width is None:
            self.config(width = 350)
        self.grid_columnconfigure(1, weight = 1)
        self.grid_rowconfigure(1, weight = 1)
        self.RI = RowIndex(self,
                           max_rh = max_rh,
                           max_row_width = max_row_width,
                           row_index_width = row_index_width,
                           row_index_align = row_index_align,
                           index_bg = index_bg,
                           index_border_fg = index_border_fg,
                           index_grid_fg = index_grid_fg,
                           index_fg = index_fg,
                           index_selected_cells_bg = index_selected_cells_bg,
                           index_selected_cells_fg = index_selected_cells_fg,
                           index_selected_rows_bg = index_selected_rows_bg,
                           index_selected_rows_fg = index_selected_rows_fg,
                           index_hidden_rows_expander_bg = index_hidden_rows_expander_bg,
                           drag_and_drop_bg = drag_and_drop_bg,
                           resizing_line_fg = resizing_line_fg,
                           row_drag_and_drop_perform = row_drag_and_drop_perform,
                           default_row_index = default_row_index,
                           measure_subset_index = measure_subset_index,
                           auto_resize_width = auto_resize_default_row_index)
        self.CH = ColumnHeaders(self,
                                max_colwidth = max_colwidth,
                                max_header_height = max_header_height,
                                default_header = default_header,
                                header_align = header_align,
                                header_bg = header_bg,
                                header_border_fg = header_border_fg,
                                header_grid_fg = header_grid_fg,
                                header_fg = header_fg,
                                header_selected_cells_bg = header_selected_cells_bg,
                                header_selected_cells_fg = header_selected_cells_fg,
                                header_selected_columns_bg = header_selected_columns_bg,
                                header_selected_columns_fg = header_selected_columns_fg,
                                header_hidden_columns_expander_bg = header_hidden_columns_expander_bg,
                                drag_and_drop_bg = drag_and_drop_bg,
                                column_drag_and_drop_perform = column_drag_and_drop_perform,
                                measure_subset_header = measure_subset_header,
                                resizing_line_fg = resizing_line_fg)
        self.MT = MainTable(self,
                            enable_edit_cell_auto_resize = enable_edit_cell_auto_resize,
                            page_up_down_select_row = page_up_down_select_row,
                            arrow_key_down_right_scroll_page = arrow_key_down_right_scroll_page,
                            display_selected_fg_over_highlights = display_selected_fg_over_highlights,
                            show_vertical_grid = show_vertical_grid,
                            show_horizontal_grid = show_horizontal_grid,
                            column_width = column_width,
                            row_height = row_height,
                            column_headers_canvas = self.CH,
                            row_index_canvas = self.RI,
                            headers = headers,
                            header_height = header_height,
                            data_reference = data if data_reference is None else data_reference,
                            total_cols = total_columns,
                            total_rows = total_rows,
                            row_index = row_index,
                            font = font,
                            header_font = header_font,
                            popup_menu_font = popup_menu_font,
                            popup_menu_fg = popup_menu_fg,
                            popup_menu_bg = popup_menu_bg,
                            popup_menu_highlight_bg = popup_menu_highlight_bg,
                            popup_menu_highlight_fg = popup_menu_highlight_fg,
                            align = align,
                            table_bg = table_bg,
                            table_grid_fg = table_grid_fg,
                            table_fg = table_fg,
                            show_selected_cells_border = show_selected_cells_border,
                            table_selected_cells_border_fg = table_selected_cells_border_fg,
                            table_selected_cells_bg = table_selected_cells_bg,
                            table_selected_cells_fg = table_selected_cells_fg,
                            table_selected_rows_border_fg = table_selected_rows_border_fg,
                            table_selected_rows_bg = table_selected_rows_bg,
                            table_selected_rows_fg = table_selected_rows_fg,
                            table_selected_columns_border_fg = table_selected_columns_border_fg,
                            table_selected_columns_bg = table_selected_columns_bg,
                            table_selected_columns_fg = table_selected_columns_fg,
                            displayed_columns = displayed_columns,
                            all_columns_displayed = all_columns_displayed,
                            empty_horizontal = empty_horizontal,
                            empty_vertical = empty_vertical,
                            max_undos = max_undos)
        self.TL = TopLeftRectangle(parentframe = self,
                                   main_canvas = self.MT,
                                   row_index_canvas = self.RI,
                                   header_canvas = self.CH,
                                   top_left_bg = top_left_bg,
                                   top_left_fg = top_left_fg,
                                   top_left_fg_highlight = top_left_fg_highlight)
        self.yscroll = ttk.Scrollbar(self, command = self.MT.set_yviews, orient = "vertical")
        self.xscroll = ttk.Scrollbar(self, command = self.MT.set_xviews, orient = "horizontal")
        if show_table:
            self.MT.grid(row = 1, column = 1, sticky = "nswe")
            self.MT["xscrollcommand"] = self.xscroll.set
            self.MT["yscrollcommand"] = self.yscroll.set
        if show_top_left:
            self.TL.grid(row = 0, column = 0)
        if show_row_index:
            self.RI.grid(row = 1, column = 0, sticky = "nswe")
            self.RI["yscrollcommand"] = self.yscroll.set
        if show_header:
            self.CH.grid(row = 0, column = 1, sticky = "nswe")
            self.CH["xscrollcommand"] = self.xscroll.set
        if show_x_scrollbar:
            self.xscroll.grid(row = 2, column = 1, columnspan = 2, sticky = "nswe")
            self.xscroll_showing = True
            self.xscroll_disabled = False
        else:
            self.xscroll_showing = False
            self.xscroll_disabled = True
        if show_y_scrollbar:
            self.yscroll.grid(row = 1, column = 2, sticky = "nswe")
            self.yscroll_showing = True
            self.yscroll_disabled = False
        else:
            self.yscroll_showing = False
            self.yscroll_disabled = True
        if theme != "light blue":
            self.MT.display_selected_fg_over_highlights = True
            self.change_theme(theme)
            for k, v in locals().items():
                if k in theme_light_blue and v != theme_light_blue[k]:
                    self.set_options(**{k: v})
        if set_all_heights_and_widths:
            self.set_all_cell_sizes_to_text()
        self.MT.update()
        self.update()
        self.refresh()
        if startup_select is not None:
            try:
                if startup_select[-1] == "cells":
                    self.create_selection_box(*startup_select)
                    self.set_currently_selected(startup_select[0], startup_select[1], selection_binding = False)
                    self.see(startup_select[0], startup_select[1])
                elif startup_select[-1] == "rows":
                    self.create_selection_box(startup_select[0], 0, startup_select[1], len(self.MT.col_positions) - 1, "rows")
                    self.set_currently_selected("row", startup_select[0], selection_binding = False)
                    self.see(startup_select[0], 0)
                elif startup_select[-1] in ("cols", "columns"):
                    self.create_selection_box(0, startup_select[0], len(self.MT.row_positions) - 1, startup_select[1], "cols")
                    self.set_currently_selected("col", startup_select[0], selection_binding = False)
                    self.see(0, startup_select[0])
            except:
                pass
        if startup_focus:
            self.MT.focus_set()

    def show(self, canvas = "all"):
        if canvas == "all":
            self.hide()
            self.TL.grid(row = 0, column = 0)
            self.RI.grid(row = 1, column = 0, sticky = "nswe")
            self.CH.grid(row = 0, column = 1, sticky = "nswe")
            self.MT.grid(row = 1, column = 1, sticky = "nswe")
            self.yscroll.grid(row = 0, column = 2, rowspan = 3, sticky = "nswe")
            self.xscroll.grid(row = 2, column = 1, sticky = "nswe")
            self.MT["xscrollcommand"] = self.xscroll.set
            self.CH["xscrollcommand"] = self.xscroll.set
            self.MT["yscrollcommand"] = self.yscroll.set
            self.RI["yscrollcommand"] = self.yscroll.set
            self.xscroll_showing = True
            self.yscroll_showing = True
            self.xscroll_disabled = False
            self.yscroll_disabled = False
        elif canvas == "row_index":
            self.RI.grid(row = 1, column = 0, sticky = "nswe")
            self.MT["yscrollcommand"] = self.yscroll.set
            self.RI["yscrollcommand"] = self.yscroll.set
            self.MT.show_index = True
        elif canvas == "header":
            self.CH.grid(row = 0, column = 1, sticky = "nswe")
            self.MT["xscrollcommand"] = self.xscroll.set
            self.CH["xscrollcommand"] = self.xscroll.set
            self.MT.show_header = True
        elif canvas == "top_left":
            self.TL.grid(row = 0, column = 0)
        elif canvas == "x_scrollbar":
            self.xscroll.grid(row = 2, column = 1, columnspan = 2, sticky = "nswe")
            self.xscroll_showing = True
            self.xscroll_disabled = False
        elif canvas == "y_scrollbar":
            self.yscroll.grid(row = 1, column = 2, sticky = "nswe")
            self.yscroll_showing = True
            self.yscroll_disabled = False
        self.MT.update()

    def hide(self, canvas = "all"):
        if canvas.lower() == "all":
            self.TL.grid_forget()
            self.RI.grid_forget()
            self.CH.grid_forget()
            self.MT.grid_forget()
            self.yscroll.grid_forget()
            self.xscroll.grid_forget()
            self.xscroll_showing = False
            self.yscroll_showing = False
            self.xscroll_disabled = True
            self.yscroll_disabled = True
        elif canvas == "row_index":
            self.RI.grid_forget()
            self.RI["yscrollcommand"] = 0
            self.MT.show_index = False
        elif canvas == "header":
            self.CH.grid_forget()
            self.CH["xscrollcommand"] = 0
            self.MT.show_header = False
        elif canvas == "top_left":
            self.TL.grid_forget()
        elif canvas == "x_scrollbar":
            self.xscroll.grid_forget()
            self.xscroll_showing = False
            self.xscroll_disabled = True
        elif canvas == "y_scrollbar":
            self.yscroll.grid_forget()
            self.yscroll_showing = False
            self.yscroll_disabled = True

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

    def popup_menu_add_command(self, label, func, table_menu = True, index_menu = True, header_menu = True):
        if label not in self.MT.extra_table_rc_menu_funcs and table_menu:
            self.MT.extra_table_rc_menu_funcs[label] = func
        if label not in self.MT.extra_index_rc_menu_funcs and index_menu:
            self.MT.extra_index_rc_menu_funcs[label] = func
        if label not in self.MT.extra_header_rc_menu_funcs and header_menu:
            self.MT.extra_header_rc_menu_funcs[label] = func
        self.MT.create_rc_menus()

    def popup_menu_del_command(self, label = None):
        if label is None:
            self.MT.extra_table_rc_menu_funcs = {}
            self.MT.extra_index_rc_menu_funcs = {}
            self.MT.extra_header_rc_menu_funcs = {}
        else:
            if label in self.MT.extra_table_rc_menu_funcs:
                del self.MT.extra_table_rc_menu_funcs[label]
            if label in self.MT.extra_index_rc_menu_funcs:
                del self.MT.extra_index_rc_menu_funcs[label]
            if label in self.MT.extra_header_rc_menu_funcs:
                del self.MT.extra_header_rc_menu_funcs[label]
        self.MT.create_rc_menus()

    def extra_bindings(self, bindings, func = "None"):
        if isinstance(bindings, str) and bindings.lower() in ("bind_all", "unbind_all"):
            self.MT.extra_begin_ctrl_c_func = None if func == "None" else func
            self.MT.extra_begin_ctrl_x_func = None if func == "None" else func
            self.MT.extra_begin_ctrl_v_func = None if func == "None" else func
            self.MT.extra_begin_ctrl_z_func = None if func == "None" else func
            self.MT.extra_begin_delete_key_func = None if func == "None" else func
            self.MT.extra_begin_edit_cell_func = None if func == "None" else func
            self.MT.extra_begin_edit_cell_func = None if func == "None" else func
            self.RI.ri_extra_begin_drag_drop_func = None if func == "None" else func
            self.CH.ch_extra_begin_drag_drop_func = None if func == "None" else func
            self.MT.extra_begin_del_rows_rc_func = None if func == "None" else func
            self.MT.extra_begin_del_cols_rc_func = None if func == "None" else func
            self.MT.extra_begin_insert_cols_rc_func = None if func == "None" else func
            self.MT.extra_begin_insert_rows_rc_func = None if func == "None" else func
            
            self.MT.extra_end_ctrl_c_func = None if func == "None" else func
            self.MT.extra_end_ctrl_x_func = None if func == "None" else func
            self.MT.extra_end_ctrl_v_func = None if func == "None" else func
            self.MT.extra_end_ctrl_z_func = None if func == "None" else func
            self.MT.extra_end_delete_key_func = None if func == "None" else func
            self.MT.extra_begin_edit_cell_func = None if func == "None" else func
            self.MT.extra_end_edit_cell_func = None if func == "None" else func
            self.RI.ri_extra_end_drag_drop_func = None if func == "None" else func
            self.CH.ch_extra_end_drag_drop_func = None if func == "None" else func
            self.MT.extra_end_del_rows_rc_func = None if func == "None" else func
            self.MT.extra_end_del_cols_rc_func = None if func == "None" else func
            self.MT.extra_end_insert_cols_rc_func = None if func == "None" else func
            self.MT.extra_end_insert_rows_rc_func = None if func == "None" else func
            
            self.MT.selection_binding_func = None if func == "None" else func
            self.MT.select_all_binding_func = None if func == "None" else func
            self.RI.selection_binding_func = None if func == "None" else func
            self.CH.selection_binding_func = None if func == "None" else func
            self.MT.drag_selection_binding_func = None if func == "None" else func
            self.RI.drag_selection_binding_func = None if func == "None" else func
            self.CH.drag_selection_binding_func = None if func == "None" else func
            self.MT.shift_selection_binding_func = None if func == "None" else func
            self.RI.shift_selection_binding_func = None if func == "None" else func
            self.CH.shift_selection_binding_func = None if func == "None" else func
            self.MT.deselection_binding_func = None
        else:
            if isinstance(bindings[0], str) and func == "None":
                iterable = [bindings]
            elif isinstance(bindings, str) and func != "None":
                iterable = [(bindings, func)]
            else:
                iterable = bindings
            for binding, func in iterable:
                binding == binding.lower()
                if binding in ("begin_copy", "begin_ctrl_c"):
                    self.MT.extra_begin_ctrl_c_func = func
                if binding in ("ctrl_c", "end_copy", "end_ctrl_c"):
                    self.MT.extra_end_ctrl_c_func = func

                if binding in ("begin_cut", "begin_ctrl_x"):
                    self.MT.extra_begin_ctrl_x_func = func
                if binding in ("ctrl_x", "end_cut", "end_ctrl_x"):
                    self.MT.extra_end_ctrl_x_func = func

                if binding in ("begin_paste", "begin_ctrl_v"):
                    self.MT.extra_begin_ctrl_v_func = func
                if binding in ("ctrl_v", "end_paste", "end_ctrl_v"):
                    self.MT.extra_end_ctrl_v_func = func

                if binding in ("begin_undo", "begin_ctrl_z"):
                    self.MT.extra_begin_ctrl_z_func = func
                if binding in ("ctrl_z", "end_undo", "end_ctrl_z"):
                    self.MT.extra_end_ctrl_z_func = func

                if binding in ("begin_delete_key", "begin_delete"):
                    self.MT.extra_begin_delete_key_func = func
                if binding in ("delete_key", "end_delete", "end_delete_key"):
                    self.MT.extra_end_delete_key_func = func
                    
                if binding == "begin_edit_cell":
                    self.MT.extra_begin_edit_cell_func = func
                if binding == "end_edit_cell" or binding == "edit_cell":
                    self.MT.extra_end_edit_cell_func = func

                if binding == "begin_row_index_drag_drop":
                    self.RI.ri_extra_begin_drag_drop_func = func
                if binding in ("row_index_drag_drop", "end_row_index_drag_drop"):
                    self.RI.ri_extra_end_drag_drop_func = func

                if binding == "begin_column_header_drag_drop":
                    self.CH.ch_extra_begin_drag_drop_func = func
                if binding in ("column_header_drag_drop", "end_column_header_drag_drop"):
                    self.CH.ch_extra_end_drag_drop_func = func

                if binding in ("begin_rc_delete_row", "begin_delete_rows"):
                    self.MT.extra_begin_del_rows_rc_func = func
                if binding in ("rc_delete_row", "end_rc_delete_row", "end_delete_rows"):
                    self.MT.extra_end_del_rows_rc_func = func

                if binding in ("begin_rc_delete_column", "begin_delete_columns"):
                    self.MT.extra_begin_del_cols_rc_func = func
                if binding in ("rc_delete_column", "end_rc_delete_column", "end_delete_columns"):
                    self.MT.extra_end_del_cols_rc_func = func
                    
                if binding in ("begin_rc_insert_column", "begin_insert_column", "begin_insert_columns"):
                    self.MT.extra_begin_insert_cols_rc_func = func
                if binding in ("rc_insert_column", "end_rc_insert_column", "end_insert_column", "end_insert_columns"):
                    self.MT.extra_end_insert_cols_rc_func = func

                if binding in ("begin_rc_insert_row", "begin_insert_row", "begin_insert_rows"):
                    self.MT.extra_begin_insert_rows_rc_func = func
                if binding in ("rc_insert_row", "end_rc_insert_row", "end_insert_row", "end_insert_rows"):
                    self.MT.extra_end_insert_rows_rc_func = func
                    
                if binding == "cell_select":
                    self.MT.selection_binding_func = func
                if binding in ("select_all", "ctrl_a"):
                    self.MT.select_all_binding_func = func
                if binding == "row_select":
                    self.RI.selection_binding_func = func
                if binding in ("col_select", "column_select"):
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
                    
                if binding == "all_select_events":
                    self.MT.selection_binding_func = func
                    self.MT.select_all_binding_func = func
                    self.RI.selection_binding_func = func
                    self.CH.selection_binding_func = func
                    self.MT.drag_selection_binding_func = func
                    self.RI.drag_selection_binding_func = func
                    self.CH.drag_selection_binding_func = func
                    self.MT.shift_selection_binding_func = func
                    self.RI.shift_selection_binding_func = func
                    self.CH.shift_selection_binding_func = func
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
        elif binding == get_rc_binding():
            self.MT.extra_rc_func = func
            self.CH.extra_rc_func = func
            self.RI.extra_rc_func = func
            self.TL.extra_rc_func = func
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
        elif binding == get_rc_binding():
            self.MT.extra_rc_func = None
            self.CH.extra_rc_func = None
            self.RI.extra_rc_func = None
            self.TL.extra_rc_func = None
        else:
            self.MT.unbind(binding)
            self.CH.unbind(binding)
            self.RI.unbind(binding)
            self.TL.unbind(binding)

    def enable_bindings(self, bindings = "all"):
        self.MT.enable_bindings(bindings)

    def disable_bindings(self, bindings = "all"):
        self.MT.disable_bindings(bindings)

    def basic_bindings(self, enable = False):
        for canvas in (self.MT, self.CH, self.RI, self.TL):
            canvas.basic_bindings(enable)

    def edit_bindings(self, enable = False):
        if enable:
            self.MT.edit_bindings(True)
        elif not enable:
            self.MT.edit_bindings(False)

    def cell_edit_binding(self, enable = False):
        self.MT.bind_cell_edit(enable)

    def identify_region(self, event):
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
            return list(accumulate(chain([0], (colpos for c in range(total_cols)))))
        return list(accumulate(chain([0], (colpos for c in range(len(self.MT.col_positions) - 1)))))

    def get_example_canvas_row_heights(self, total_rows = None):
        rowpos = self.MT.default_rh[1]
        if total_rows is not None:
            return list(accumulate(chain([0], (rowpos for c in range(total_rows)))))
        return list(accumulate(chain([0], (rowpos for c in range(len(self.MT.row_positions) - 1)))))

    def get_column_widths(self, canvas_positions = False):
        if canvas_positions:
            return [int(n) for n in self.MT.col_positions]
        return [int(b - a) for a, b in zip(self.MT.col_positions, islice(self.MT.col_positions, 1, len(self.MT.col_positions)))]
    
    def get_row_heights(self, canvas_positions = False):
        if canvas_positions:
            return [int(n) for n in self.MT.row_positions]
        return [int(b - a) for a, b in zip(self.MT.row_positions, islice(self.MT.row_positions, 1, len(self.MT.row_positions)))]

    def set_all_cell_sizes_to_text(self, redraw = True):
        self.MT.set_all_cell_sizes_to_text()
        if redraw:
            self.refresh()

    def set_all_column_widths(self, width = None, only_set_if_too_small = False, redraw = True, recreate_selection_boxes = True):
        self.CH.set_width_of_all_cols(width = width, only_set_if_too_small = only_set_if_too_small, recreate = recreate_selection_boxes)
        if redraw:
            self.refresh()

    def column_width(self, column = None, width = None, only_set_if_too_small = False, redraw = True):
        if column == "all":
            if width == "default":
                self.MT.reset_col_positions()
        elif column == "displayed":
            if width == "text":
                sc, ec = self.MT.get_visible_columns(self.MT.canvasx(0), self.MT.canvasx(self.winfo_width()))
                for c in range(sc, ec - 1):
                    self.CH.set_col_width(c)
        elif width == "text" and column is not None:
            self.CH.set_col_width(col = column, width = None, only_set_if_too_small = only_set_if_too_small)
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
                self.MT.col_positions = list(accumulate(chain([0], (width for width in column_widths))))
        return cwx

    def set_all_row_heights(self, height = None, only_set_if_too_small = False, redraw = True, recreate_selection_boxes = True):
        self.RI.set_height_of_all_rows(height = height, only_set_if_too_small = only_set_if_too_small, recreate = recreate_selection_boxes)
        if redraw:
            self.refresh()

    def set_width_of_index_to_text(self, recreate = True):
        self.RI.set_width_of_index_to_text(recreate = recreate)

    def row_height(self, row = None, height = None, only_set_if_too_small = False, redraw = True):
        if row == "all":
            if height == "default":
                self.MT.reset_row_positions()
        elif row == "displayed":
            if height == "text":
                sr, er = self.MT.get_visible_rows(self.MT.canvasy(0), self.MT.canvasy(self.winfo_width()))
                for r in range(sr, er - 1):
                    self.RI.set_row_height(r)
        elif height == "text" and row is not None:
            self.RI.set_row_height(row = row, height = None, only_set_if_too_small = only_set_if_too_small)
        elif height is not None and row is not None:
            self.RI.set_row_height(row = row, height = height, only_set_if_too_small = only_set_if_too_small)
        elif row is not None:
            return int(self.MT.row_positions[row + 1] - self.MT.row_positions[row])
        if redraw:
            self.refresh()

    def set_row_heights(self, row_heights = None, canvas_positions = False, reset = False, verify = False):
        if reset:
            self.MT.reset_row_positions()
            return
        if isinstance(row_heights, list):
            qmin = self.MT.min_rh
            if canvas_positions:
                if verify:
                    self.MT.row_positions = list(accumulate(chain([0], (height if qmin < height else qmin
                                                                        for height in [x - z for z, x in zip(islice(row_heights, 0, None),
                                                                                                             islice(row_heights, 1, None))]))))
                else:
                    self.MT.row_positions = row_heights
            else:
                if verify:
                    self.MT.row_positions = [qmin if z < qmin or not isinstance(z, int) or isinstance(z, bool) else z for z in row_heights]
                else:
                    self.MT.row_positions = list(accumulate(chain([0], (height for height in row_heights))))

    def verify_row_heights(self, row_heights, canvas_positions = False):
        if row_heights[0] != 0 or isinstance(row_heights[0], bool):
            return False
        if not isinstance(row_heights, list):
            return False
        if canvas_positions:
            if any(x - z < self.MT.min_rh or not isinstance(x, int) or isinstance(x, bool) for z, x in zip(islice(row_heights, 0, None), islice(row_heights, 1, None))):
                return False
        elif not canvas_positions:
            if any(z < self.MT.min_rh or not isinstance(z, int) or isinstance(z, bool) for z in row_heights):
                return False
        return True

    def verify_column_widths(self, column_widths, canvas_positions = False):
        if column_widths[0] != 0 or isinstance(column_widths[0], bool):
            return False
        if not isinstance(column_widths, list):
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
            self.MT.default_rh = (height if isinstance(height, str) else "pixels", height if isinstance(height, int) else self.MT.GetLinesHeight(int(height)))
        return self.MT.default_rh[1]

    def default_header_height(self, height = None):
        if height is not None:
            self.MT.default_hh = (height if isinstance(height, str) else "pixels", height if isinstance(height, int) else self.MT.GetHdrLinesHeight(int(height)))
        return self.MT.default_hh[1]

    def default_column_width(self, width = None):
        if width is not None:
            if width < self.MT.min_cw:
                self.MT.default_cw = self.MT.min_cw + 20
            else:
                self.MT.default_cw = int(width)
        return self.MT.default_cw

    def create_dropdown(self,
                        r = 0,
                        c = 0,
                        values = [],
                        set_value = None,
                        state = "readonly",
                        see = False,
                        destroy_on_leave = False,
                        destroy_on_select = True,
                        current = False,
                        set_cell_on_select = True,
                        redraw = True,
                        recreate_selection_boxes = True):
        self.MT.create_dropdown(r = r,
                                c = c,
                                values = values,
                                set_value = set_value,
                                state = state,
                                see = see,
                                destroy_on_leave = destroy_on_leave,
                                destroy_on_select = destroy_on_select,
                                current = current,
                                set_cell_on_select = set_cell_on_select,
                                redraw = redraw,
                                recreate = recreate_selection_boxes)

    def get_dropdown_value(self, r, c, current = False, destroy = True, set_cell_on_select = True, redraw = True, recreate = True):
        return self.MT.get_dropdown_value(r, c, current = current, destroy = destroy,
                                          set_cell_on_select = set_cell_on_select, redraw = redraw, recreate = recreate)

    def get_dropdown_values(self, r = 0, c = 0):
        return self.MT.cell_options[(r, c)]['dropdown'][0].dropdown['values']

    def set_dropdown_values(self, r = 0, c = 0, values = [], displayed = None):
        self.MT.cell_options[(r, c)]['dropdown'][0].dropdown['values'] = values
        if displayed is not None:
            self.MT.cell_options[(r, c)]['dropdown'][0].dropdown.set(displayed)
            self.MT.set_cell_data(r, c, displayed, undo = False, cell_resize = False)

    def delete_dropdown(self, r = 0, c = 0):
        if r == "all":
            for r, c in self.MT.cell_options:
                if 'dropdown' in self.MT.cell_options[(r, c)]:
                    self.MT.destroy_dropdown(r, c)
        else:
            self.MT.destroy_dropdown(r, c)

    def get_dropdowns(self):
        return {k: v['dropdown'] for k, v in self.MT.cell_options if 'dropdown' in v}

    def refresh_dropdowns(self, dropdowns = []):
        self.MT.refresh_dropdowns(dropdowns = dropdowns)

    def set_all_dropdown_values_to_sheet(self):
        for r, c in self.MT.cell_options:
            if 'dropdown' in self.MT.cell_options[(r, c)]:
                self.MT.cell_options[(r, c)]['dropdown'][0].set_displayed(self.MT.data_ref[r][c if self.MT.all_columns_displayed else self.MT.displayed_columns[c]])

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

    def delete_row_position(self, idx, deselect_all = False):
        self.MT.del_row_position(idx = idx,
                                 deselect_all = deselect_all)

    def delete_row(self, idx = 0, deselect_all = False):
        del self.MT.data_ref[idx]
        self.MT.del_row_position(idx = idx,
                                 deselect_all = deselect_all)
        self.MT.cell_options = {t1: t2 for t1, t2 in self.MT.cell_options.items() if t1[0] != idx}
        if idx in self.MT.row_options:
            del self.MT.row_options[idx]
        if idx in self.RI.cell_options:
            del self.RI.cell_options[idx]
        self.MT.cell_options = {(rn if rn < idx else rn - 1, cn): t2 for (rn, cn), t2 in self.MT.cell_options.items()}
        self.MT.row_options = {rn if rn < idx else rn - 1: t for rn, t in self.MT.row_options.items()}
        self.RI.cell_options = {rn if rn < idx else rn - 1: t for rn, t in self.RI.cell_options.items()}

    def insert_row_position(self, idx = "end", height = None, deselect_all = False, redraw = False):
        self.MT.insert_row_position(idx = idx,
                                    height = height,
                                    deselect_all = deselect_all)
        if redraw:
            self.redraw()

    def insert_row_positions(self, idx = "end", heights = None, deselect_all = False, redraw = False):
        self.MT.insert_row_positions(idx = idx,
                                     heights = heights,
                                     deselect_all = deselect_all)
        if redraw:
            self.redraw()

    def total_rows(self, number = None, mod_positions = True, mod_data = True):
        if number is None:
            return int(self.MT.total_data_rows())
        if not isinstance(number, int) or number < 0:
            raise ValueError("number argument must be integer and > 0")
        if number > len(self.MT.data_ref):
            if mod_positions:
                height = self.MT.GetLinesHeight(self.MT.default_rh)
                for r in range(number - len(self.MT.data_ref)):
                    self.MT.insert_row_position("end", height)
        elif number < len(self.MT.data_ref):
            self.MT.row_positions[number + 1:] = []
        if mod_data:
            self.MT.data_dimensions(total_rows = number)

    def total_columns(self, number = None, mod_positions = True, mod_data = True):
        total_cols = self.MT.total_data_cols()
        if number is None:
            return int(total_cols)
        if not isinstance(number, int) or number < 0:
            raise ValueError("number argument must be integer and > 0")
        if number > total_cols:
            if mod_positions:
                width = self.MT.default_cw
                for c in range(number - total_cols):
                    self.MT.insert_col_position("end", width)
        elif number < total_cols:
            if not self.MT.all_columns_displayed:
                self.MT.display_columns(enable = False, reset_col_positions = False, deselect_all = True)
            self.MT.col_positions[number + 1:] = []
        if mod_data:
            self.MT.data_dimensions(total_columns = number)

    def sheet_display_dimensions(self, total_rows = None, total_columns = None):
        if total_rows is None and total_columns is None:
            return len(self.MT.row_positions) - 1, len(self.MT.col_positions) - 1
        if total_rows is not None:
            height = self.MT.GetLinesHeight(self.MT.default_rh)
            self.MT.row_positions = list(accumulate(chain([0], (height for row in range(total_rows)))))
        if total_columns is not None:
            width = self.MT.default_cw
            self.MT.col_positions = list(accumulate(chain([0], (width for column in range(total_columns)))))

    def set_sheet_data_and_display_dimensions(self, total_rows = None, total_columns = None):
        self.sheet_display_dimensions(total_rows = total_rows, total_columns = total_columns)
        self.MT.data_dimensions(total_rows = total_rows, total_columns = total_columns)

    def move_row_position(self, row, moveto):
        self.MT.move_row_position(row, moveto)

    def move_row(self, row, moveto):
        self.MT.move_row_position(row, moveto)
        self.MT.data_ref.insert(moveto, self.MT.data_ref.pop(row))
        popped_ri = {t1: t2 for t1, t2 in self.RI.cell_options.items() if t1 == row}
        popped_cell = {t1: t2 for t1, t2 in self.MT.cell_options.items() if t1[0] == row}
        popped_row = {t1: t2 for t1, t2 in self.MT.row_options.items() if t1 == row}
        
        popped_ri = {t1: self.RI.cell_options.pop(t1) for t1 in popped_ri}
        popped_cell = {t1: self.MT.cell_options.pop(t1) for t1 in popped_cell}
        popped_row = {t1: self.MT.row_options.pop(t1) for t1 in popped_row}

        self.RI.cell_options = {t1 if t1 < row else t1 - 1: t2 for t1, t2 in self.RI.cell_options.items()}
        self.RI.cell_options = {t1 if t1 < moveto else t1 + 1: t2 for t1, t2 in self.RI.cell_options.items()}

        self.MT.row_options = {t1 if t1 < row else t1 - 1: t2 for t1, t2 in self.MT.row_options.items()}
        self.MT.row_options = {t1 if t1 < moveto else t1 + 1: t2 for t1, t2 in self.MT.row_options.items()}

        self.MT.cell_options = {(t10 if t10 < row else t10 - 1, t11): t2 for (t10, t11), t2 in self.MT.cell_options.items()}
        self.MT.cell_options = {(t10 if t10 < moveto else t10 + 1, t11): t2 for (t10, t11), t2 in self.MT.cell_options.items()}

        if popped_ri:
            self.RI.cell_options[moveto] = popped_ri[row]

        if popped_row:
            self.MT.row_options[moveto] = popped_row[row]

        if popped_cell:
            newrowsdct = {row: moveto}
            for (t10, t11), t2 in popped_cell.items():
                self.MT.cell_options[(newrowsdct[t10], t11)] = t2

    def delete_column_position(self, idx, deselect_all = False):
        self.MT.del_col_position(idx,
                                 deselect_all = deselect_all)

    def delete_column(self, idx = 0, deselect_all = False):
        for rn in range(len(self.MT.data_ref)):
            del self.MT.data_ref[rn][idx] 
        self.MT.del_col_position(idx,
                                 deselect_all = deselect_all)
        self.MT.cell_options = {t1: t2 for t1, t2 in self.MT.cell_options.items() if t1[1] != idx}
        if idx in self.MT.col_options:
            del self.MT.col_options[idx]
        if idx in self.CH.cell_options:
            del self.CH.cell_options[idx]
        self.MT.cell_options = {(rn, cn if cn < idx else cn - 1): t2 for (rn, cn), t2 in self.MT.cell_options.items()}
        self.MT.col_options = {cn if cn < idx else cn - 1: t for cn, t in self.MT.col_options.items()}
        self.CH.cell_options = {cn if cn < idx else cn - 1: t for cn, t in self.CH.cell_options.items()}

    def insert_column_position(self, idx = "end", width = None, deselect_all = False, redraw = False):
        self.MT.insert_col_position(idx = idx,
                                    width = width,
                                    deselect_all = deselect_all)
        if redraw:
            self.redraw()

    def insert_column_positions(self, idx = "end", widths = None, deselect_all = False, redraw = False):
        self.MT.insert_col_positions(idx = idx,
                                     widths = widths,
                                     deselect_all = deselect_all)
        if redraw:
            self.redraw()

    def move_column_position(self, column, moveto):
        self.MT.move_col_position(column, moveto)

    def move_column(self, column, moveto):
        self.MT.move_col_position(column, moveto)
        for rn in range(len(self.MT.data_ref)):
            self.MT.data_ref[rn].insert(moveto, self.MT.data_ref[rn].pop(column))
        popped_ch = {t1: t2 for t1, t2 in self.CH.cell_options.items() if t1 == column}
        popped_cell = {t1: t2 for t1, t2 in self.MT.cell_options.items() if t1[1] == column}
        popped_col = {t1: t2 for t1, t2 in self.MT.col_options.items() if t1 == column}
        
        popped_ch = {t1: self.CH.cell_options.pop(t1) for t1 in popped_ch}
        popped_cell = {t1: self.MT.cell_options.pop(t1) for t1 in popped_cell}
        popped_col = {t1: self.MT.col_options.pop(t1) for t1 in popped_col}

        self.CH.cell_options = {t1 if t1 < column else t1 - 1: t2 for t1, t2 in self.CH.cell_options.items()}
        self.CH.cell_options = {t1 if t1 < moveto else t1 + 1: t2 for t1, t2 in self.CH.cell_options.items()}

        self.MT.col_options = {t1 if t1 < column else t1 - 1: t2 for t1, t2 in self.MT.col_options.items()}
        self.MT.col_options = {t1 if t1 < moveto else t1 + 1: t2 for t1, t2 in self.MT.col_options.items()}

        self.MT.cell_options = {(t10, t11 if t11 < column else t11 - 1): t2 for (t10, t11), t2 in self.MT.cell_options.items()}
        self.MT.cell_options = {(t10, t11 if t11 < moveto else t11 + 1): t2 for (t10, t11), t2 in self.MT.cell_options.items()}

        if popped_ch:
            self.CH.cell_options[moveto] = popped_ch[column]

        if popped_col:
            self.MT.col_options[moveto] = popped_col[column]

        if popped_cell:
            newcolsdct = {column: moveto}
            for (t10, t11), t2 in popped_cell.items():
                self.MT.cell_options[(t10, newcolsdct[t11])] = t2

    def create_text_editor(self, row = 0, column = 0, text = None, state = "normal", see = True, set_data_ref_on_destroy = False,
                           binding = None):
        self.MT.create_text_editor(r = row, c = column, text = text, state = state, see = see, set_data_ref_on_destroy = set_data_ref_on_destroy,
                                   binding = binding)

    def set_text_editor_value(self, text = "", r = None, c = None):
        if self.MT.text_editor is not None and r is None and c is None:
            self.MT.text_editor.set_text(text)
        elif self.MT.text_editor is not None and self.MT.text_editor_loc == (r, c):
            self.MT.text_editor.set_text(text)

    def bind_text_editor_set(self, func, row, column):
        self.MT.bind_text_editor_destroy(func, row, column)

    def get_text_editor_value(self, destroy_tup = None, r = None, c = None, set_data_ref_on_destroy = True, event = None, destroy = True, move_down = True, redraw = True, recreate = True):
        return self.MT.get_text_editor_value(destroy_tup = destroy_tup,
                                             r = r,
                                             c = c,
                                             set_data_ref_on_destroy = set_data_ref_on_destroy,
                                             event = event,
                                             destroy = destroy,
                                             move_down = move_down,
                                             redraw = redraw,
                                             recreate = recreate)

    def destroy_text_editor(self, event = None):
        self.MT.destroy_text_editor(event = event)

    def get_text_editor_widget(self, event = None):
        try:
            return self.MT.text_editor.textedit
        except:
            return None

    def bind_key_text_editor(self, key, function):
        self.MT.text_editor_user_bound_keys[key] = function

    def unbind_key_text_editor(self, key):
        if key == "all":
            for key in self.MT.text_editor_user_bound_keys:
                try:
                    self.MT.text_editor.textedit.unbind(key)
                except:
                    pass
            self.MT.text_editor_user_bound_keys = {}
        else:
            if key in self.MT.text_editor_user_bound_keys:
                del self.MT.text_editor_user_bound_keys[key]
            try:
                self.MT.text_editor.textedit.unbind(key)
            except:
                pass

    def get_xview(self):
        return self.MT.xview()

    def get_yview(self):
        return self.MT.yview()

    def set_xview(self, position, option = "moveto"):
        self.MT.set_xviews(option, position)

    def set_yview(self, position, option = "moveto"):
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

    def move_down(self):
        self.MT.move_down()

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

    def get_currently_selected(self, get_coords = False, return_nones_if_not = False):
        curr = self.MT.currently_selected()
        if get_coords:
            if curr:
                if curr[0] == "row":
                    return curr[1], 0
                elif curr[0] == "column":
                    return 0, curr[1]
                else:
                    return curr
            elif not curr and return_nones_if_not:
                return (None, None)
            else:
                return curr
        else:
            if curr:
                return curr
            elif not curr and return_nones_if_not:
                return (None, None)
            else:
                return curr

    def set_currently_selected(self, current_tuple_0 = 0, current_tuple_1 = 0, selection_binding = True):
        if isinstance(current_tuple_0, int) and isinstance(current_tuple_1, int):
            self.MT.create_current(r = current_tuple_0,
                                   c = current_tuple_1,
                                   type_ = "cell",
                                   inside = True if self.MT.cell_selected(current_tuple_0, current_tuple_1) else False)
        elif current_tuple_0 == "row" and isinstance(current_tuple_1, int):
            self.MT.create_current(r = current_tuple_1,
                                   c = 0,
                                   type_ = "row",
                                   inside = True if self.MT.cell_selected(current_tuple_1, 0) else False)
        elif current_tuple_0 in ("col", "column") and isinstance(current_tuple_1, int):
            self.MT.create_current(r = 0,
                                   c = current_tuple_1,
                                   type_ = "col",
                                   inside = True if self.MT.cell_selected(0, current_tuple_1) else False)
        if selection_binding and self.MT.selection_binding_func is not None:
            self.MT.selection_binding_func(("select_cell", ) + tuple((current_tuple_0, current_tuple_1)))

    def get_selected_rows(self, get_cells = False, get_cells_as_rows = False, return_tuple = False):
        if return_tuple:
            return tuple(self.MT.get_selected_rows(get_cells = get_cells, get_cells_as_rows = get_cells_as_rows))
        else:
            return self.MT.get_selected_rows(get_cells = get_cells, get_cells_as_rows = get_cells_as_rows)

    def get_selected_columns(self, get_cells = False, get_cells_as_columns = False, return_tuple = False):
        if return_tuple:
            return tuple(self.MT.get_selected_cols(get_cells = get_cells, get_cells_as_cols = get_cells_as_columns))
        else:
            return self.MT.get_selected_cols(get_cells = get_cells, get_cells_as_cols = get_cells_as_columns)

    def get_selected_cells(self, get_rows = False, get_columns = False, sort_by_row = False, sort_by_column = False):
        if sort_by_row and sort_by_column:
            sels = sorted(self.MT.get_selected_cells(get_rows = get_rows, get_cols = get_columns),
                          key = lambda t: t[1])
            return sorted(sels,
                          key = lambda t: t[0])
        elif sort_by_row:
            return sorted(self.MT.get_selected_cells(get_rows = get_rows, get_cols = get_columns),
                          key = lambda t: t[0])
        elif sort_by_column:
            return sorted(self.MT.get_selected_cells(get_rows = get_rows, get_cols = get_columns),
                          key = lambda t: t[1])
        else:
            return self.MT.get_selected_cells(get_rows = get_rows, get_cols = get_columns)

    def get_all_selection_boxes(self):
        return self.MT.get_all_selection_boxes()

    def get_all_selection_boxes_with_types(self):
        return self.MT.get_all_selection_boxes_with_types()

    def create_selection_box(self, r1, c1, r2, c2, type_ = "cells"):
        return self.MT.create_selected(r1 = r1, c1 = c1, r2 = r2, c2 = c2, type_ = type_ if type_ != "columns" else "cols")

    def recreate_all_selection_boxes(self):
        self.MT.recreate_all_selection_boxes()

    def cell_selected(self, r, c):
        return self.MT.cell_selected(r, c)

    def row_selected(self, r):
        return self.MT.row_selected(r)

    def column_selected(self, c):
        return self.MT.col_selected(c)
    
    def anything_selected(self, exclude_columns = False, exclude_rows = False, exclude_cells = False):
        return self.MT.anything_selected(exclude_columns = exclude_columns, exclude_rows = exclude_rows, exclude_cells = exclude_cells)

    def all_selected(self):
        return self.MT.all_selected()

    def align_rows(self, rows = [], align = "global", align_index = False, redraw = True): #"center", "w", "e" or "global"
        self.MT.align_rows(rows = rows,
                           align = align,
                           align_index = align_index)
        if redraw:
            self.redraw()

    def align_columns(self, columns = [], align = "global", align_header = False, redraw = True): #"center", "w", "e" or "global"
        self.MT.align_columns(columns = columns,
                              align = align,
                              align_header = align_header)
        if redraw:
            self.redraw()

    def align_cells(self, row = 0, column = 0, cells = [], align = "global", redraw = True): #"center", "w", "e" or "global"
        self.MT.align_cells(row = row,
                            column = column,
                            cells = cells,
                            align = align)
        if redraw:
            self.redraw()

    def align_header(self, columns = [], align = "global", redraw = True):
        self.CH.align_cells(columns = columns,
                            align = align)
        if redraw:
            self.redraw()

    def align_index(self, rows = [], align = "global", redraw = True):
        self.RI.align_cells(rows = rows,
                            align = align)
        if redraw:
            self.redraw()

    def readonly_rows(self, rows = [], readonly = True, redraw = True):
        self.MT.readonly_rows(rows = rows,
                              readonly = readonly)
        if redraw:
            self.redraw()

    def readonly_columns(self, columns = [], readonly = True, redraw = True):
        self.MT.readonly_columns(columns = columns,
                                 readonly = readonly)
        if redraw:
            self.redraw()

    def readonly_cells(self, row = 0, column = 0, cells = [], readonly = True, redraw = True):
        self.MT.readonly_cells(row = row,
                               column = column,
                               cells = cells,
                               readonly = readonly)
        if redraw:
            self.redraw()

    def highlight_rows(self, rows = [], bg = None, fg = None, highlight_index = True, redraw = False):
        self.MT.highlight_rows(rows = rows,
                               bg = bg,
                               fg = fg,
                               highlight_index = highlight_index,
                               redraw = redraw)

    def highlight_columns(self, columns = [], bg = None, fg = None, highlight_header = True, redraw = False):
        self.MT.highlight_cols(cols = columns,
                                  bg = bg,
                                  fg = fg,
                                  highlight_header = highlight_header,
                                  redraw = redraw)

    def dehighlight_all(self):
        for k in self.MT.cell_options:
            if 'highlight' in self.MT.cell_options[k]:
                del self.MT.cell_options[k]['highlight']
                
        for k in self.MT.row_options:
            if 'highlight' in self.MT.row_options[k]:
                del self.MT.row_options[k]['highlight']
                
        for k in self.MT.col_options:
            if 'highlight' in self.MT.col_options[k]:
                del self.MT.col_options[k]['highlight']
                
        for k in self.RI.cell_options:
            if 'highlight' in self.RI.cell_options[k]:
                del self.RI.cell_options[k]['highlight']
                
        for k in self.CH.cell_options:
            if 'highlight' in self.CH.cell_options[k]:
                del self.CH.cell_options[k]['highlight']

    def dehighlight_rows(self, rows = [], redraw = False):
        if isinstance(rows, int):
            rows_ = [rows]
        else:
            rows_ = rows
        if not rows_ or rows_ == "all":
            for r in self.MT.row_options:
                if 'highlight' in self.MT.row_options[r]:
                    del self.MT.row_options[r]['highlight']

            for r in self.RI.cell_options:
                if 'highlight' in self.RI.cell_options[r]:
                    del self.RI.cell_options[r]['highlight']
        else:
            for r in rows_:
                try:
                    del self.MT.row_options[r]['highlight']
                except:
                    pass
                try:
                    del self.RI.cell_options[r]['highlight']
                except:
                    pass
        if redraw:
            self.refresh(True, True)

    def dehighlight_columns(self, columns = [], redraw = False):
        if isinstance(columns, int):
            columns_ = [columns]
        else:
            columns_ = columns
        if not columns_ or columns_ == "all":
            for c in self.MT.col_options:
                if 'highlight' in self.MT.col_options[c]:
                    del self.MT.col_options[c]['highlight']

            for c in self.CH.cell_options:
                if 'highlight' in self.CH.cell_options[c]:
                    del self.CH.cell_options[c]['highlight']
        else:
            for c in columns_:
                try:
                    del self.MT.col_options[c]['highlight']
                except:
                    pass
                try:
                    del self.CH.cell_options[c]['highlight']
                except:
                    pass
        if redraw:
            self.refresh(True, True)

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
        if row == "all" and canvas == "table":
            self.MT.cell_options = {}
        elif row == "all" and canvas == "row_index":
            self.RI.cell_options = {}
        elif row == "all" and canvas == "header":
            self.CH.cell_options = {}
        if canvas == "table":
            if cells and not all_:
                for t in cells:
                    try:
                        del self.MT.cell_options[t]['highlight']
                    except:
                        pass
            elif not all_:
                if (row, column) in self.MT.cell_options and 'highlight' in self.MT.cell_options[(row, column)]:
                    del self.MT.cell_options[(row, column)]['highlight']
            elif all_:
                for k in self.MT.cell_options:
                    if 'highlight' in self.MT.cell_options[k]:
                        del self.MT.cell_options[k]['highlight']
        elif canvas == "row_index":
            if cells and not all_:
                for r in cells:
                    try:
                        del self.RI.cell_options[r]['highlight']
                    except:
                        pass
            elif not all_:
                if row in self.RI.cell_options and 'highlight' in self.RI.cell_options[row]:
                    del self.RI.cell_options[row]['highlight']
            elif all_:
                for r in self.RI.cell_options:
                    if 'highlight' in self.RI.cell_options[r]:
                        del self.RI.cell_options[r]['highlight']
        elif canvas == "header":
            if cells and not all_:
                for c in cells:
                    try:
                        del self.CH.cell_options[c]['highlight']
                    except:
                        pass
            elif not all_:
                if column in self.CH.cell_options and 'highlight' in self.CH.cell_options[column]:
                    del self.CH.cell_options[column]['highlight']
            elif all_:
                for c in self.CH.cell_options:
                    if 'highlight' in self.CH.cell_options[c]:
                        del self.CH.cell_options[c]['highlight']
        if redraw:
            self.refresh(True, True)

    def get_cell_options(self, canvas = "table"):
        if canvas == "table":
            return self.MT.cell_options
        elif canvas == "row_index":
            return self.RI.cell_options
        elif canvas == "header":
            return self.CH.cell_options
        
    def get_frame_y(self, y):
        return y + self.CH.current_height

    def get_frame_x(self, x):
        return x + self.RI.current_width

    def align(self, align = None, redraw = True):
        if align is None:
            return self.MT.align
        elif align in ("w", "center", "e"):
            self.MT.align = align
        else:
            raise ValueError("Align must be either 'w' or 'center' or 'e'")
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

    def font(self, newfont = None, reset_row_positions = True):
        self.MT.font(newfont, reset_row_positions = reset_row_positions)

    def header_font(self, newfont = None):
        self.MT.header_font(newfont)

    def set_options(self,
                    enable_edit_cell_auto_resize = None,
                    page_up_down_select_row = None,
                    arrow_key_down_right_scroll_page = None,
                    display_selected_fg_over_highlights = None,
                    empty_horizontal = None,
                    empty_vertical = None,
                    show_horizontal_grid = None,
                    show_vertical_grid = None,
                    top_left_fg_highlight = None,
                    auto_resize_default_row_index = None,
                    font = None,
                    default_header = None,
                    default_row_index = None,
                    header_font = None,
                    show_selected_cells_border = None,
                    theme = None,
                    max_colwidth = None,
                    max_row_height = None,
                    max_header_height = None,
                    max_row_width = None,
                    header_height = None,
                    row_height = None,
                    column_width = None,
                    header_bg = None,
                    header_border_fg = None,
                    header_grid_fg = None,
                    header_fg = None,
                    header_selected_cells_bg = None,
                    header_selected_cells_fg = None,
                    header_hidden_columns_expander_bg = None,
                    index_bg = None,
                    index_border_fg = None,
                    index_grid_fg = None,
                    index_fg = None,
                    index_selected_cells_bg = None,
                    index_selected_cells_fg = None,
                    index_hidden_rows_expander_bg = None,
                    top_left_bg = None,
                    top_left_fg = None,
                    frame_bg = None,
                    table_bg = None,
                    table_grid_fg = None,
                    table_fg = None,
                    table_selected_cells_border_fg = None,
                    table_selected_cells_bg = None,
                    table_selected_cells_fg = None,
                    resizing_line_fg = None,
                    drag_and_drop_bg = None,
                    outline_thickness = None,
                    outline_color = None,
                    header_selected_columns_bg = None,
                    header_selected_columns_fg = None,
                    index_selected_rows_bg = None,
                    index_selected_rows_fg = None,
                    table_selected_rows_border_fg = None,
                    table_selected_rows_bg = None,
                    table_selected_rows_fg = None,
                    table_selected_columns_border_fg = None,
                    table_selected_columns_bg = None,
                    table_selected_columns_fg = None,
                    popup_menu_font = None,
                    popup_menu_fg = None,
                    popup_menu_bg = None,
                    popup_menu_highlight_bg = None,
                    popup_menu_highlight_fg = None,
                    row_drag_and_drop_perform = None,
                    column_drag_and_drop_perform = None,
                    measure_subset_index = None,
                    measure_subset_header = None,
                    redraw = True):
        if arrow_key_down_right_scroll_page is not None:
            self.MT.arrow_key_down_right_scroll_page = arrow_key_down_right_scroll_page
        if enable_edit_cell_auto_resize is not None:
            self.MT.enable_edit_cell_auto_resize = enable_edit_cell_auto_resize
        if header_hidden_columns_expander_bg is not None:
            self.CH.header_hidden_columns_expander_bg = header_hidden_columns_expander_bg
        if index_hidden_rows_expander_bg is not None:
            self.RI.index_hidden_rows_expander_bg = index_hidden_rows_expander_bg
        if page_up_down_select_row is not None:
            self.MT.page_up_down_select_row = page_up_down_select_row
        if display_selected_fg_over_highlights is not None:
            self.MT.display_selected_fg_over_highlights = display_selected_fg_over_highlights
        if show_horizontal_grid is not None:
            self.MT.show_horizontal_grid = show_horizontal_grid
        if show_vertical_grid is not None:
            self.MT.show_vertical_grid = show_vertical_grid
        if empty_horizontal is not None:
            self.MT.empty_horizontal = empty_horizontal
        if empty_vertical is not None:
            self.MT.empty_vertical = empty_vertical
        if row_height is not None:
            self.MT.default_rh = (row_height if isinstance(row_height, str) else "pixels", row_height if isinstance(row_height, int) else self.MT.GetLinesHeight(int(row_height)))
        if column_width is not None:
            self.MT.default_cw = self.MT.min_cw + 20 if column_width < self.MT.min_cw else int(column_width)
        if header_height is not None:
            self.MT.default_hh = (header_height if isinstance(header_height, str) else "pixels", header_height if isinstance(header_height, int) else self.MT.GetHdrLinesHeight(int(header_height)))
        if measure_subset_index is not None:
            self.RI.measure_subset_index = measure_subset_index
        if measure_subset_header is not None:
            self.CH.measure_subset_hdr = measure_subset_header
        if row_drag_and_drop_perform is not None:
            self.RI.row_drag_and_drop_perform = row_drag_and_drop_perform
        if column_drag_and_drop_perform is not None:
            self.CH.column_drag_and_drop_perform = column_drag_and_drop_perform
        if popup_menu_font is not None:
            self.MT.popup_menu_font = popup_menu_font
        if popup_menu_fg is not None:
            self.MT.popup_menu_fg = popup_menu_fg
        if popup_menu_bg is not None:
            self.MT.popup_menu_bg = popup_menu_bg
        if popup_menu_highlight_bg is not None:
            self.MT.popup_menu_highlight_bg = popup_menu_highlight_bg
        if popup_menu_highlight_fg is not None:
            self.MT.popup_menu_highlight_fg = popup_menu_highlight_fg
        if top_left_fg_highlight is not None:
            self.TL.top_left_fg_highlight = top_left_fg_highlight
        if auto_resize_default_row_index is not None:
            self.RI.auto_resize_width = auto_resize_default_row_index
        if header_selected_columns_bg is not None:
            self.CH.header_selected_columns_bg = header_selected_columns_bg
        if header_selected_columns_fg is not None:
            self.CH.header_selected_columns_fg = header_selected_columns_fg
        if index_selected_rows_bg is not None:
            self.RI.index_selected_rows_bg = index_selected_rows_bg
        if index_selected_rows_fg is not None:
            self.RI.index_selected_rows_fg = index_selected_rows_fg
        if table_selected_rows_border_fg is not None:
            self.MT.table_selected_rows_border_fg = table_selected_rows_border_fg
        if table_selected_rows_bg is not None:
            self.MT.table_selected_rows_bg = table_selected_rows_bg
        if table_selected_rows_fg is not None:
            self.MT.table_selected_rows_fg = table_selected_rows_fg
        if table_selected_columns_border_fg is not None:
            self.MT.table_selected_columns_border_fg = table_selected_columns_border_fg
        if table_selected_columns_bg is not None:
            self.MT.table_selected_columns_bg = table_selected_columns_bg
        if table_selected_columns_fg is not None:
            self.MT.table_selected_columns_fg = table_selected_columns_fg 
        if default_header is not None:
            self.CH.default_hdr = default_header.lower()
        if default_row_index is not None:
            self.RI.default_index = default_row_index.lower()
        if max_colwidth is not None:
            self.CH.max_cw = float(max_colwidth)
        if max_row_height is not None:
            self.RI.max_rh = float(max_row_height)
        if max_header_height is not None:
            self.CH.max_header_height = float(max_header_height)
        if max_row_width is not None:
            self.RI.max_row_width = float(max_row_width)
        if font is not None:
            self.MT.font(font)
        if header_font is not None:
            self.MT.header_font(header_font)
        if theme is not None:
            self.change_theme(theme)
        if show_selected_cells_border is not None:
            self.MT.show_selected_cells_border = show_selected_cells_border
        if header_bg is not None:
            self.CH.config(background = header_bg)
        if header_border_fg is not None:
            self.CH.header_border_fg = header_border_fg
        if header_grid_fg is not None:
            self.CH.header_grid_fg = header_grid_fg
        if header_fg is not None:
            self.CH.header_fg = header_fg
        if header_selected_cells_bg is not None:
            self.CH.header_selected_cells_bg = header_selected_cells_bg
        if header_selected_cells_fg is not None:
            self.CH.header_selected_cells_fg = header_selected_cells_fg
        if index_bg is not None:
            self.RI.config(background = index_bg)
        if index_border_fg is not None:
            self.RI.index_border_fg = index_border_fg
        if index_grid_fg is not None:
            self.RI.index_grid_fg = index_grid_fg
        if index_fg is not None:
            self.RI.index_fg = index_fg
        if index_selected_cells_bg is not None:
            self.RI.index_selected_cells_bg = index_selected_cells_bg
        if index_selected_cells_fg is not None:
            self.RI.index_selected_cells_fg = index_selected_cells_fg
        if top_left_bg is not None:
            self.TL.config(background = top_left_bg)
        if top_left_fg is not None:
            self.TL.top_left_fg = top_left_fg
            self.TL.itemconfig("rw", fill = top_left_fg)
            self.TL.itemconfig("rh", fill = top_left_fg)
        
        if frame_bg is not None:
            self.config(background = frame_bg)
        if table_bg is not None:
            self.MT.config(background = table_bg)
            self.MT.table_bg = table_bg
        if table_grid_fg is not None:
            self.MT.table_grid_fg = table_grid_fg
        if table_fg is not None:
            self.MT.table_fg = table_fg
        if table_selected_cells_border_fg is not None:
            self.MT.table_selected_cells_border_fg = table_selected_cells_border_fg
        if table_selected_cells_bg is not None:
            self.MT.table_selected_cells_bg = table_selected_cells_bg
        if table_selected_cells_fg is not None:
            self.MT.table_selected_cells_fg = table_selected_cells_fg
        if resizing_line_fg is not None:
            self.CH.resizing_line_fg = resizing_line_fg
            self.RI.resizing_line_fg = resizing_line_fg
        if drag_and_drop_bg is not None:
            self.CH.drag_and_drop_bg = drag_and_drop_bg
            self.RI.drag_and_drop_bg = drag_and_drop_bg
        if outline_thickness is not None:
            self.config(highlightthickness = outline_thickness)
        if outline_color is not None:
            self.config(highlightbackground = outline_color)
        self.MT.create_rc_menus()
        if redraw:
            self.refresh()

    def change_theme(self, theme = "light blue"):
        if theme == "light blue":
            self.MT.display_selected_fg_over_highlights = False
            self.set_options(**theme_light_blue,
                              redraw = True)
            self.config(bg = theme_light_blue['table_bg'])
        elif theme == "light green":
            self.MT.display_selected_fg_over_highlights = True
            self.set_options(**theme_light_green,
                              redraw = True)
            self.config(bg = theme_light_green['table_bg'])
        elif theme == "dark blue":
            self.MT.display_selected_fg_over_highlights = True
            self.set_options(**theme_dark_blue,
                              redraw = True)
            self.config(bg = theme_dark_blue['table_bg'])
        elif theme == "dark green":
            self.MT.display_selected_fg_over_highlights = True
            self.set_options(**theme_dark_green,
                              redraw = True)
            self.config(bg = theme_dark_green['table_bg'])
        self.MT.recreate_all_selection_boxes()
            
    def data_reference(self,
                       newdataref = None,
                       reset_col_positions = True,
                       reset_row_positions = True,
                       redraw = False):
        return self.MT.data_reference(newdataref,
                                      reset_col_positions,
                                      reset_row_positions,
                                      redraw)

    def set_sheet_data(self,
                       data = [[]],
                       reset_col_positions = True,
                       reset_row_positions = True,
                       redraw = True,
                       verify = False,
                       reset_highlights = False):
        if verify:
            if not isinstance(data, list) or not all(isinstance(row, list) for row in data):
                raise ValueError("data argument must be a list of lists, sublists being rows")
        if reset_highlights:
            self.dehighlight_all()
        return self.MT.data_reference(data,
                                      reset_col_positions,
                                      reset_row_positions,
                                      redraw,
                                      return_id = False)

    def get_sheet_data(self, return_copy = False, get_header = False, get_index = False):
        if return_copy:
            if get_header and get_index:
                index_limit = len(self.MT.my_row_index)
                return self.MT.my_hdrs.copy() + [[f"{self.MT.my_row_index[rn]}"] + r.copy() if rn < index_limit else [""] + r.copy() for rn, r in enumerate(self.MT.data_ref)]
            elif get_header and not get_index:
                return self.MT.my_hdrs.copy() + [r.copy() for r in self.MT.data_ref]
            elif get_index and not get_header:
                index_limit = len(self.MT.my_row_index)
                return [[f"{self.MT.my_row_index[rn]}"] + r.copy() if rn < index_limit else [""] + r.copy() for rn, r in enumerate(self.MT.data_ref)]
            elif not get_index and not get_header:
                return [r.copy() for r in self.MT.data_ref]
        else:
            if get_header and get_index:
                index_limit = len(self.MT.my_row_index)
                return self.MT.my_hdrs + [[self.MT.my_row_index[rn]] + r if rn < index_limit else [""] + r for rn, r in enumerate(self.MT.data_ref)]
            elif get_header and not get_index:
                return self.MT.my_hdrs + self.MT.data_ref
            elif get_index and not get_header:
                index_limit = len(self.MT.my_row_index)
                return [[self.MT.my_row_index[rn]] + r if rn < index_limit else [""] + r for rn, r in enumerate(self.MT.data_ref)]
            elif not get_index and not get_header:
                return self.MT.data_ref

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

    def set_cell_data(self, r, c, value = "", set_copy = True, redraw = False):
        self.MT.data_ref[r][c] = f"{value}" if set_copy else value
        if redraw and self.after_redraw_id is None:
            self.after_redraw_id = self.after(self.after_redraw_time_ms, self.after_redraw)

    def set_column_data(self, c, values = tuple(), add_rows = True, redraw = False):
        if add_rows:
            maxidx = len(self.MT.data_ref) - 1
            total_cols = None
            height = self.MT.default_rh[1]
            for rn, v in enumerate(values):
                if rn > maxidx:
                    if total_cols is None:
                        total_cols = self.MT.total_data_cols()
                    self.MT.data_ref.append(list(repeat("", total_cols)))
                    self.MT.insert_row_position("end", height = height)
                    maxidx += 1
                if c > len(self.MT.data_ref[rn]) - 1:
                    self.MT.data_ref[rn].extend(list(repeat("", c - len(self.MT.data_ref[rn]))))
                self.MT.data_ref[rn][c] = v
        else:
            for rn, v in enumerate(values):
                if c > len(self.MT.data_ref[rn]) - 1:
                    self.MT.data_ref[rn].extend(list(repeat("", c - len(self.MT.data_ref[rn]))))
                self.MT.data_ref[rn][c] = v
        if redraw and self.after_redraw_id is None:
            self.after_redraw_id = self.after(self.after_redraw_time_ms, self.after_redraw)

    def insert_column(self, values = None, idx = "end", width = None, deselect_all = False, add_rows = True, equalize_data_row_lengths = True,
                      mod_column_positions = True,
                      redraw = False):
        if mod_column_positions:
            self.MT.insert_col_position(idx = idx,
                                        width = width,
                                        deselect_all = deselect_all)
            if not self.MT.all_columns_displayed:
                try:
                    disp_next = max(self.MT.displayed_columns) + 1
                except:
                    disp_next = 0
                self.MT.displayed_columns.extend(list(range(disp_next, disp_next + 1)))
        if equalize_data_row_lengths:
            old_total = self.MT.equalize_data_row_lengths()
        else:
            old_total = self.MT.total_data_cols()
        if values is None:
            data = list(repeat("", self.MT.total_data_rows()))
        else:
            data = values
        maxidx = len(self.MT.data_ref) - 1
        if add_rows:
            height = self.MT.default_rh[1]
            if idx == "end":
                for rn, v in enumerate(data):
                    if rn > maxidx:
                        self.MT.data_ref.append(list(repeat("", old_total)))
                        self.MT.insert_row_position("end", height = height)
                        maxidx += 1
                    self.MT.data_ref[rn].append(v)
            else:
                for rn, v in enumerate(data):
                    if rn > maxidx:
                        self.MT.data_ref.append(list(repeat("", old_total)))
                        self.MT.insert_row_position("end", height = height)
                        maxidx += 1
                    self.MT.data_ref[rn].insert(idx, v)
        else:
            if idx == "end":
                for rn, v in enumerate(data):
                    if rn > maxidx:
                        break
                    self.MT.data_ref[rn].append(v)
            else:
                for rn, v in enumerate(data):
                    if rn > maxidx:
                        break
                    self.MT.data_ref[rn].insert(idx, v)
        if isinstance(idx, int):
            self.MT.cell_options = {(rn, cn if cn < idx else cn + 1): t2 for (rn, cn), t2 in self.MT.cell_options.items()}
            self.MT.col_options = {cn if cn < idx else cn + 1: t for cn, t in self.MT.col_options.items()}
            self.CH.cell_options = {cn if cn < idx else cn + 1: t for cn, t in self.CH.cell_options.items()}
        if redraw and self.after_redraw_id is None:
            self.after_redraw_id = self.after(self.after_redraw_time_ms, self.after_redraw)

    def insert_columns(self, columns = 1, idx = "end", widths = None, deselect_all = False, add_rows = True, equalize_data_row_lengths = True,
                       mod_column_positions = True,
                       redraw = False):
        if mod_column_positions:
            self.MT.insert_col_positions(idx = idx,
                                         widths = columns if isinstance(columns, int) and widths is None else widths,
                                         deselect_all = deselect_all)
        if equalize_data_row_lengths:
            old_total = self.MT.equalize_data_row_lengths()
        else:
            old_total = self.MT.total_data_cols()
        if isinstance(columns, int):
            total_rows = self.MT.total_data_rows()
            data = [list(repeat("", total_rows)) for i in range(columns)]
            numcols = columns
        else:
            data = columns
            numcols = len(columns)
        if mod_column_positions:
            if not self.MT.all_columns_displayed:
                try:
                    disp_next = max(self.MT.displayed_columns) + 1
                except:
                    disp_next = 0
                self.MT.displayed_columns.extend(list(range(disp_next, disp_next + numcols)))
        maxidx = len(self.MT.data_ref) - 1
        if add_rows:
            height = self.MT.default_rh[1]
            if idx == "end":
                for values in reversed(data):
                    for rn, v in enumerate(values):
                        if rn > maxidx:
                            self.MT.data_ref.append(list(repeat("", old_total)))
                            self.MT.insert_row_position("end", height = height)
                            maxidx += 1
                        self.MT.data_ref[rn].append(v)
            else:
                for values in reversed(data):
                    for rn, v in enumerate(values):
                        if rn > maxidx:
                            self.MT.data_ref.append(list(repeat("", old_total)))
                            self.MT.insert_row_position("end", height = height)
                            maxidx += 1
                        self.MT.data_ref[rn].insert(idx, v)
        else:
            if idx == "end":
                for values in reversed(data):
                    for rn, v in enumerate(values):
                        if rn > maxidx:
                            break
                        self.MT.data_ref[rn].append(v)
            else:
                for values in reversed(data):
                    for rn, v in enumerate(values):
                        if rn > maxidx:
                            break
                        self.MT.data_ref[rn].insert(idx, v)
        if isinstance(idx, int):
            num_add = len(data)
            self.MT.cell_options = {(rn, cn if cn < idx else cn + num_add): t2 for (rn, cn), t2 in self.MT.cell_options.items()}
            self.MT.col_options = {cn if cn < idx else cn + num_add: t for cn, t in self.MT.col_options.items()}
            self.CH.cell_options = {cn if cn < idx else cn + num_add: t for cn, t in self.CH.cell_options.items()}
        if redraw and self.after_redraw_id is None:
            self.after_redraw_id = self.after(self.after_redraw_time_ms, self.after_redraw)

    def set_row_data(self, r, values = tuple(), add_columns = True, redraw = False):
        if len(self.MT.data_ref) - 1 < r:
            raise Exception("Row number is out of range")
        maxidx = len(self.MT.data_ref[r]) - 1
        if add_columns:
            for c, v in enumerate(values):
                if c > maxidx:
                    self.MT.data_ref[r].append(v)
                    if self.MT.all_columns_displayed and c >= len(self.MT.col_positions) - 1:
                        self.MT.insert_col_position("end")
                else:
                    self.MT.data_ref[r][c] = v
        else:
            for c, v in enumerate(values):
                if c > maxidx:
                    self.MT.data_ref[r].append(v)
                else:
                    self.MT.data_ref[r][c] = v
        if redraw and self.after_redraw_id is None:
            self.after_redraw_id = self.after(self.after_redraw_time_ms, self.after_redraw)

    def insert_row(self, values = None, idx = "end", height = None, deselect_all = False, add_columns = True,
                   redraw = False):
        self.MT.insert_row_position(idx = idx,
                                    height = height,
                                    deselect_all = deselect_all)
        total_cols = None
        if values is None:
            total_cols = self.MT.total_data_cols()
            data = list(repeat("", total_cols))
        elif isinstance(values, list):
            data = values
        else:
            data = list(values)
        if add_columns:
            if total_cols is None:
                total_cols = self.MT.total_data_cols()
            if len(data) > total_cols:
                self.MT.data_ref[:] = [r + list(repeat("", len(data) - total_cols)) for r in self.MT.data_ref]
            elif len(data) < total_cols:
                data += list(repeat("", total_cols - len(data)))
            if self.MT.all_columns_displayed:
                if not self.MT.col_positions:
                    self.MT.col_positions = [0]
                if len(data) > len(self.MT.col_positions) - 1:
                    self.insert_column_positions("end", len(data) - (len(self.MT.col_positions) - 1))
        if isinstance(idx, str) and idx.lower() == "end":
            self.MT.data_ref.append(data)
        else:
            self.MT.data_ref.insert(idx, data)
            self.MT.cell_options = {(rn if rn < idx else rn + 1, cn): t2 for (rn, cn), t2 in self.MT.cell_options.items()}
            self.MT.row_options = {rn if rn < idx else rn + 1: t for rn, t in self.MT.row_options.items()}
            self.RI.cell_options = {rn if rn < idx else rn + 1: t for rn, t in self.RI.cell_options.items()}
        if redraw and self.after_redraw_id is None:
            self.after_redraw_id = self.after(self.after_redraw_time_ms, self.after_redraw)

    def insert_rows(self, rows = 1, idx = "end", heights = None, deselect_all = False, add_columns = True,
                    redraw = False):
        total_cols = None
        if isinstance(rows, int):
            total_cols = self.MT.total_data_cols()
            data = [list(repeat("", total_cols)) for i in range(rows)]
        elif isinstance(rows, list):
            data = rows
        else:
            data = list(rows)
        if heights is None:
            heights_ = len(data)
        else:
            heights_ = heights
        self.MT.insert_row_positions(idx = idx,
                                     heights = heights_,
                                     deselect_all = deselect_all)
        if add_columns:
            if total_cols is None:
                total_cols = self.MT.total_data_cols()
            data_max_cols = len(max(data, key = len))
            if data_max_cols > total_cols:
                self.MT.data_ref[:] = [r + list(repeat("", data_max_cols - total_cols)) for r in self.MT.data_ref]
            else:
                data[:] = [r + list(repeat("", total_cols - data_max_cols)) for r in data]
            if self.MT.all_columns_displayed:
                if not self.MT.col_positions:
                    self.MT.col_positions = [0]
                if data_max_cols > len(self.MT.col_positions) - 1:
                    self.insert_column_positions("end", data_max_cols - (len(self.MT.col_positions) - 1))
        if isinstance(idx, str) and idx.lower() == "end":
            self.MT.data_ref.extend(data)
        else:
            self.MT.data_ref[idx:idx] = data
            num_add = len(data)
            self.MT.cell_options = {(rn if rn < idx else rn + num_add, cn): t2 for (rn, cn), t2 in self.MT.cell_options.items()}
            self.MT.row_options = {rn if rn < idx else rn + num_add: t for rn, t in self.MT.row_options.items()}
            self.RI.cell_options = {rn if rn < idx else rn + num_add: t for rn, t in self.RI.cell_options.items()}
        if redraw and self.after_redraw_id is None:
            self.after_redraw_id = self.after(self.after_redraw_time_ms, self.after_redraw)

    def sheet_data_dimensions(self, total_rows = None, total_columns = None):
        self.MT.data_dimensions(total_rows, total_columns)

    def get_total_rows(self):
        return len(self.MT.data_ref)

    def equalize_data_row_lengths(self):
        return self.MT.equalize_data_row_lengths()
                    
    def displayed_columns(self,
                          indexes = None,
                          enable = None,
                          reset_col_positions = True,
                          set_col_positions = True,
                          refresh = False,
                          redraw = False,
                          deselect_all = True):
        res = self.MT.display_columns(indexes = indexes,
                                      enable = enable,
                                      reset_col_positions = reset_col_positions,
                                      set_col_positions = set_col_positions,
                                      deselect_all = deselect_all)
        if refresh or redraw:
            self.refresh()
        return res

    def display_subset_of_columns(self, indexes = None, enable = None, reset_col_positions = True, set_col_positions = True, refresh = False, redraw = False, deselect_all = True):
        return self.displayed_columns(indexes = indexes, enable = enable, reset_col_positions = reset_col_positions, set_col_positions = set_col_positions, refresh = refresh, redraw = redraw, deselect_all = deselect_all)

    def display_columns(self, indexes = None, enable = None, reset_col_positions = True, set_col_positions = True, refresh = False, redraw = False, deselect_all = True):
        return self.displayed_columns(indexes = indexes, enable = enable, reset_col_positions = reset_col_positions, set_col_positions = set_col_positions, refresh = refresh, redraw = redraw, deselect_all = deselect_all)

    def show_ctrl_outline(self, canvas = "table", start_cell = (0, 0), end_cell = (1, 1)):
        self.MT.show_ctrl_outline(canvas = canvas, start_cell = start_cell, end_cell = end_cell)

    def get_ctrl_x_c_boxes(self):
        return self.MT.get_ctrl_x_c_boxes()

    def get_selected_min_max(self): # returns (min_y, min_x, max_y, max_x) of any selections including rows/columns
        return self.MT.get_selected_min_max()
        
    def headers(self, newheaders = None, index = None, reset_col_positions = False, show_headers_if_not_sheet = True):
        return self.MT.headers(newheaders, index, reset_col_positions = reset_col_positions, show_headers_if_not_sheet = show_headers_if_not_sheet)
        
    def row_index(self, newindex = None, index = None, reset_row_positions = False, show_index_if_not_sheet = True):
        return self.MT.row_index(newindex, index, reset_row_positions = reset_row_positions, show_index_if_not_sheet = show_index_if_not_sheet)

    def reset_undos(self):
        self.MT.undo_storage = deque(maxlen = self.MT.max_undos)

    def after_redraw(self, redraw_header = True, redraw_row_index = True):
        self.MT.main_table_redraw_grid_and_text(redraw_header = redraw_header, redraw_row_index = redraw_row_index)
        self.after_redraw_id = None

    def redraw(self, redraw_header = True, redraw_row_index = True):
        self.MT.main_table_redraw_grid_and_text(redraw_header = redraw_header, redraw_row_index = redraw_row_index)

    def refresh(self, redraw_header = True, redraw_row_index = True):
        self.MT.main_table_redraw_grid_and_text(redraw_header = redraw_header, redraw_row_index = redraw_row_index)






