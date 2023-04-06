from ._tksheet_vars import *
from ._tksheet_other_classes import *
from ._tksheet_top_left_rectangle import *
from ._tksheet_column_headers import *
from ._tksheet_row_index import *
from ._tksheet_main_table import *

from collections import deque
from itertools import islice, repeat, accumulate, chain
from tkinter import ttk
import tkinter as tk


class Sheet(tk.Frame):
    def __init__(self,
                 parent,
                 show_table: bool = True,
                 show_top_left: bool = True,
                 show_row_index: bool = True,
                 show_header: bool = True,
                 show_x_scrollbar: bool = True,
                 show_y_scrollbar: bool = True,
                 width: int = None,
                 height: int = None,
                 headers: list = None,
                 header: list = None,
                 default_header: str = "letters", #letters, numbers or both
                 default_row_index: str = "numbers", #letters, numbers or both
                 show_default_header_for_empty: bool = True,
                 show_default_index_for_empty: bool = True,
                 page_up_down_select_row: bool = True,
                 expand_sheet_if_paste_too_big: bool = False,
                 paste_insert_column_limit: int = None,
                 paste_insert_row_limit: int = None,
                 show_dropdown_borders: bool = False,
                 ctrl_keys_over_dropdowns_enabled: bool = False,
                 arrow_key_down_right_scroll_page: bool = False,
                 enable_edit_cell_auto_resize: bool = True,
                 edit_cell_validation: bool = True,
                 data_reference: list = None,
                 data: list = None,
                 startup_select: tuple = None, #either (start row, end row, "rows"), (start column, end column, "rows") or (cells start row, cells start column, cells end row, cells end column, "cells")
                 startup_focus: bool = True,
                 total_columns: int = None,
                 total_rows: int = None,
                 column_width: int = 120,
                 header_height: str = "1", #str or int
                 max_colwidth: str = "inf", #str or int
                 max_rh: str = "inf", #str or int
                 max_header_height: str = "inf", #str or int
                 max_row_width: str = "inf", #str or int
                 row_index: list = None,
                 index: list = None,
                 after_redraw_time_ms: int = 100,
                 row_index_width: int = 100,
                 auto_resize_default_row_index: bool = True,
                 set_all_heights_and_widths: bool = False,
                 row_height: str = "1", #str or int
                 font: tuple = get_font(),
                 header_font: tuple = get_heading_font(),
                 index_font: tuple = get_index_font(), #currently has no effect
                 popup_menu_font: tuple = get_font(),
                 align: str = "w",
                 header_align: str = "center",
                 row_index_align: str = "center",
                 displayed_columns: list = [],
                 all_columns_displayed: bool = True,
                 max_undos: int = 30,
                 outline_thickness: int = 0,
                 outline_color: str = theme_light_blue['outline_color'],
                 column_drag_and_drop_perform: bool = True,
                 row_drag_and_drop_perform: bool = True,
                 empty_horizontal: int = 150,
                 empty_vertical: int = 100,
                 selected_rows_to_end_of_window: bool = False,
                 horizontal_grid_to_end_of_window: bool = False,
                 vertical_grid_to_end_of_window: bool = False,
                 show_vertical_grid: bool = True,
                 show_horizontal_grid: bool = True,
                 display_selected_fg_over_highlights: bool = False,
                 show_selected_cells_border: bool = True,
                 theme                              = "light blue",
                 popup_menu_fg                      = theme_light_blue['popup_menu_fg'],
                 popup_menu_bg                      = theme_light_blue['popup_menu_bg'],
                 popup_menu_highlight_bg            = theme_light_blue['popup_menu_highlight_bg'],
                 popup_menu_highlight_fg            = theme_light_blue['popup_menu_highlight_fg'],
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
                          highlightbackground = outline_color,
                          highlightcolor = outline_color)
        self.C = parent
        self.dropdown_class = Sheet_Dropdown
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
        self.RI = RowIndex(parentframe = self,
                           max_rh = max_rh,
                           max_row_width = max_row_width,
                           row_index_width = row_index_width,
                           row_index_align = self.convert_align(row_index_align),
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
                           auto_resize_width = auto_resize_default_row_index,
                           show_default_index_for_empty = show_default_index_for_empty)
        self.CH = ColumnHeaders(parentframe = self,
                                max_colwidth = max_colwidth,
                                max_header_height = max_header_height,
                                default_header = default_header,
                                header_align = self.convert_align(header_align),
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
                                resizing_line_fg = resizing_line_fg,
                                show_default_header_for_empty = show_default_header_for_empty)
        self.MT = MainTable(parentframe = self,
                            show_index = show_row_index,
                            show_header = show_header,
                            enable_edit_cell_auto_resize = enable_edit_cell_auto_resize,
                            edit_cell_validation = edit_cell_validation,
                            page_up_down_select_row = page_up_down_select_row,
                            expand_sheet_if_paste_too_big = expand_sheet_if_paste_too_big,
                            paste_insert_column_limit = paste_insert_column_limit,
                            paste_insert_row_limit = paste_insert_row_limit,
                            ctrl_keys_over_dropdowns_enabled = ctrl_keys_over_dropdowns_enabled,
                            show_dropdown_borders = show_dropdown_borders,
                            arrow_key_down_right_scroll_page = arrow_key_down_right_scroll_page,
                            display_selected_fg_over_highlights = display_selected_fg_over_highlights,
                            show_vertical_grid = show_vertical_grid,
                            show_horizontal_grid = show_horizontal_grid,
                            column_width = column_width,
                            row_height = row_height,
                            column_headers_canvas = self.CH,
                            row_index_canvas = self.RI,
                            headers = headers,
                            header = header,
                            header_height = header_height,
                            data_reference = data if data_reference is None else data_reference,
                            total_cols = total_columns,
                            total_rows = total_rows,
                            row_index = row_index,
                            index = index,
                            font = font,
                            header_font = header_font,
                            index_font = index_font,
                            popup_menu_font = popup_menu_font,
                            popup_menu_fg = popup_menu_fg,
                            popup_menu_bg = popup_menu_bg,
                            popup_menu_highlight_bg = popup_menu_highlight_bg,
                            popup_menu_highlight_fg = popup_menu_highlight_fg,
                            align = self.convert_align(align),
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
                            selected_rows_to_end_of_window = selected_rows_to_end_of_window,
                            horizontal_grid_to_end_of_window = horizontal_grid_to_end_of_window,
                            vertical_grid_to_end_of_window = vertical_grid_to_end_of_window,
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
        if show_top_left:
            self.TL.grid(row = 0, column = 0)
        if show_table:
            self.MT.grid(row = 1, column = 1, sticky = "nswe")
            self.MT["xscrollcommand"] = self.xscroll.set
            self.MT["yscrollcommand"] = self.yscroll.set
        if show_row_index:
            self.RI.grid(row = 1, column = 0, sticky = "nswe")
            self.RI["yscrollcommand"] = self.yscroll.set
        if show_header:
            self.CH.grid(row = 0, column = 1, sticky = "nswe")
            self.CH["xscrollcommand"] = self.xscroll.set
        if show_x_scrollbar:
            self.xscroll.grid(row = 2, column = 0, columnspan = 2, sticky = "nswe")
            self.xscroll_showing = True
            self.xscroll_disabled = False
        else:
            self.xscroll_showing = False
            self.xscroll_disabled = True
        if show_y_scrollbar:
            self.yscroll.grid(row = 0, column = 2, rowspan = 3, sticky = "nswe")
            self.yscroll_showing = True
            self.yscroll_disabled = False
        else:
            self.yscroll_showing = False
            self.yscroll_disabled = True
        self.update_idletasks()
        self.MT.update_idletasks()
        self.RI.update_idletasks()
        self.CH.update_idletasks()
        if theme != "light blue":
            self.change_theme(theme)
            for k, v in locals().items():
                if k in theme_light_blue and v != theme_light_blue[k]:
                    self.set_options(**{k: v})
        if set_all_heights_and_widths:
            self.set_all_cell_sizes_to_text()
        if startup_select is not None:
            try:
                if startup_select[-1] == "cells":
                    self.MT.create_selected(*startup_select)
                    self.MT.create_current(startup_select[0], startup_select[1], type_ = "cell", inside = True)
                    self.see(startup_select[0], startup_select[1])
                elif startup_select[-1] == "rows":
                    self.MT.create_selected(startup_select[0], 0, startup_select[1], len(self.MT.col_positions) - 1, "rows")
                    self.MT.create_current(startup_select[0], 0, type_ = "row", inside = True)
                    self.see(startup_select[0], 0)
                elif startup_select[-1] in ("cols", "columns"):
                    self.MT.create_selected(0, startup_select[0], len(self.MT.row_positions) - 1, startup_select[1], "columns")
                    self.MT.create_current(0, startup_select[0], type_ = "column", inside = True)
                    self.see(0, startup_select[0])
            except:
                pass
        self.refresh()
        if startup_focus:
            self.MT.focus_set()
            
    def set_refresh_timer(self, redraw = True):
        if redraw and self.after_redraw_id is None:
            self.after_redraw_id = self.after(self.after_redraw_time_ms, self.after_redraw)
            
    def after_redraw(self, redraw_header = True, redraw_row_index = True):
        self.MT.main_table_redraw_grid_and_text(redraw_header = redraw_header, redraw_row_index = redraw_row_index)
        self.after_redraw_id = None

    def show(self, canvas = "all"):
        if canvas == "all":
            self.hide()
            self.TL.grid(row = 0, column = 0)
            self.RI.grid(row = 1, column = 0, sticky = "nswe")
            self.CH.grid(row = 0, column = 1, sticky = "nswe")
            self.MT.grid(row = 1, column = 1, sticky = "nswe")
            self.yscroll.grid(row = 0, column = 2, rowspan = 3, sticky = "nswe")
            self.xscroll.grid(row = 2, column = 0, columnspan = 2, sticky = "nswe")
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
            self.xscroll.grid(row = 2, column = 0, columnspan = 2, sticky = "nswe")
            self.xscroll_showing = True
            self.xscroll_disabled = False
        elif canvas == "y_scrollbar":
            self.yscroll.grid(row = 0, column = 2, rowspan = 3, sticky = "nswe")
            self.yscroll_showing = True
            self.yscroll_disabled = False
        self.MT.update_idletasks()

    def hide(self, canvas = "all"):
        if canvas.lower() == "all":
            self.TL.grid_forget()
            self.RI.grid_forget()
            self.RI["yscrollcommand"] = 0
            self.MT.show_index = False
            self.CH.grid_forget()
            self.CH["xscrollcommand"] = 0
            self.MT.show_header = False
            self.MT.grid_forget()
            self.yscroll.grid_forget()
            self.xscroll.grid_forget()
            self.xscroll_showing = False
            self.yscroll_showing = False
            self.xscroll_disabled = True
            self.yscroll_disabled = True
        elif canvas.lower() == "row_index":
            self.RI.grid_forget()
            self.RI["yscrollcommand"] = 0
            self.MT.show_index = False
        elif canvas.lower() == "header":
            self.CH.grid_forget()
            self.CH["xscrollcommand"] = 0
            self.MT.show_header = False
        elif canvas.lower() == "top_left":
            self.TL.grid_forget()
        elif canvas.lower() == "x_scrollbar":
            self.xscroll.grid_forget()
            self.xscroll_showing = False
            self.xscroll_disabled = True
        elif canvas.lower() == "y_scrollbar":
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

    def popup_menu_add_command(self, label, func, table_menu = True, index_menu = True, header_menu = True, empty_space_menu = True):
        if label not in self.MT.extra_table_rc_menu_funcs and table_menu:
            self.MT.extra_table_rc_menu_funcs[label] = func
        if label not in self.MT.extra_index_rc_menu_funcs and index_menu:
            self.MT.extra_index_rc_menu_funcs[label] = func
        if label not in self.MT.extra_header_rc_menu_funcs and header_menu:
            self.MT.extra_header_rc_menu_funcs[label] = func
        if label not in self.MT.extra_empty_space_rc_menu_funcs and empty_space_menu:
            self.MT.extra_empty_space_rc_menu_funcs[label] = func
        self.MT.create_rc_menus()

    def popup_menu_del_command(self, label = None):
        if label is None:
            self.MT.extra_table_rc_menu_funcs = {}
            self.MT.extra_index_rc_menu_funcs = {}
            self.MT.extra_header_rc_menu_funcs = {}
            self.MT.extra_empty_space_rc_menu_funcs = {}
        else:
            if label in self.MT.extra_table_rc_menu_funcs:
                del self.MT.extra_table_rc_menu_funcs[label]
            if label in self.MT.extra_index_rc_menu_funcs:
                del self.MT.extra_index_rc_menu_funcs[label]
            if label in self.MT.extra_header_rc_menu_funcs:
                del self.MT.extra_header_rc_menu_funcs[label]
            if label in self.MT.extra_empty_space_rc_menu_funcs:
                del self.MT.extra_empty_space_rc_menu_funcs[label]
        self.MT.create_rc_menus()

    def extra_bindings(self, bindings, func = None):
        if isinstance(bindings, str) and bindings.lower() in ("all", "bind_all", "unbind_all"):
            self.MT.extra_begin_ctrl_c_func = func
            self.MT.extra_begin_ctrl_x_func = func
            self.MT.extra_begin_ctrl_v_func = func
            self.MT.extra_begin_ctrl_z_func = func
            self.MT.extra_begin_delete_key_func = func
            self.RI.ri_extra_begin_drag_drop_func = func
            self.CH.ch_extra_begin_drag_drop_func = func
            self.MT.extra_begin_del_rows_rc_func = func
            self.MT.extra_begin_del_cols_rc_func = func
            self.MT.extra_begin_insert_cols_rc_func = func
            self.MT.extra_begin_insert_rows_rc_func = func
            
            self.MT.extra_end_ctrl_c_func = func
            self.MT.extra_end_ctrl_x_func = func
            self.MT.extra_end_ctrl_v_func = func
            self.MT.extra_end_ctrl_z_func = func
            self.MT.extra_end_delete_key_func = func
            
            self.MT.extra_begin_edit_cell_func = func
            self.MT.extra_end_edit_cell_func = func
            self.CH.extra_begin_edit_cell_func = func
            self.CH.extra_end_edit_cell_func = func
            self.RI.extra_begin_edit_cell_func = func
            self.RI.extra_end_edit_cell_func = func
            
            self.RI.ri_extra_end_drag_drop_func = func
            self.CH.ch_extra_end_drag_drop_func = func
            self.MT.extra_end_del_rows_rc_func = func
            self.MT.extra_end_del_cols_rc_func = func
            self.MT.extra_end_insert_cols_rc_func = func
            self.MT.extra_end_insert_rows_rc_func = func
            
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

            self.CH.column_width_resize_func = func
            self.RI.row_height_resize_func = func
        else:
            if isinstance(bindings[0], str) and not isinstance(bindings, str):
                iterable = [bindings]
            elif isinstance(bindings, str):
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
                    
                if binding == "begin_edit_header":
                    self.CH.extra_begin_edit_cell_func = func
                if binding == "end_edit_header" or binding == "edit_header":
                    self.CH.extra_end_edit_cell_func = func
                    
                if binding == "begin_edit_index":
                    self.RI.extra_begin_edit_cell_func = func
                if binding == "end_edit_index" or binding == "edit_index":
                    self.RI.extra_end_edit_cell_func = func

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

                if binding == "column_width_resize":
                    self.CH.column_width_resize_func = func
                if binding == "row_height_resize":
                    self.RI.row_height_resize_func = func
                    
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

    def bind(self, binding, func, add = None):
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
            self.MT.bind(binding, func, add =  add)
            self.CH.bind(binding, func, add =  add)
            self.RI.bind(binding, func, add =  add)
            self.TL.bind(binding, func, add =  add)

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

    def enable_bindings(self, *bindings):
        self.MT.enable_bindings(bindings)

    def disable_bindings(self, *bindings):
        self.MT.disable_bindings(bindings)

    def basic_bindings(self, enable = False):
        for canvas in (self.MT, self.CH, self.RI, self.TL):
            canvas.basic_bindings(enable)

    def edit_bindings(self, enable = False):
        if enable:
            self.MT.edit_bindings(True)
        elif not enable:
            self.MT.edit_bindings(False)

    def cell_edit_binding(self, enable = False, keys = []):
        self.MT.bind_cell_edit(enable, keys = [])

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
        return self.MT.row_positions, self.MT.col_positions

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
        if is_iterable(column_widths):
            if canvas_positions and isinstance(column_widths, list):
                self.MT.col_positions = column_widths
            else:
                self.MT.col_positions = list(accumulate(chain([0], (width for width in column_widths))))
        return cwx

    def set_all_row_heights(self, height = None, only_set_if_too_small = False, redraw = True, recreate_selection_boxes = True):
        self.RI.set_height_of_all_rows(height = height, only_set_if_too_small = only_set_if_too_small, recreate = recreate_selection_boxes)
        if redraw:
            self.refresh()

    def set_cell_size_to_text(self, row, column, only_set_if_too_small = False, redraw = True):
        self.MT.set_cell_size_to_text(r = row, c = column, only_set_if_too_small = only_set_if_too_small)
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
        if is_iterable(row_heights):
            qmin = self.MT.min_rh
            if canvas_positions and isinstance(row_heights, list):
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

    def verify_row_heights(self, row_heights: list, canvas_positions = False):
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

    def verify_column_widths(self, column_widths: list, canvas_positions = False):
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

    def edit_cell(self, event = None, dropdown = False):
        self.MT.edit_cell_(event = event, dropdown = dropdown)

    def delete_row_position(self, idx: int, deselect_all = False):
        self.MT.del_row_position(idx = idx,
                                 deselect_all = deselect_all)

    def delete_row(self, idx = 0, deselect_all = False, redraw = True):
        self.delete_rows(rows = {idx}, deselect_all = deselect_all, redraw = redraw)

    def delete_rows(self, rows: set = set(), deselect_all = False, redraw = True):
        if deselect_all:
            self.deselect("all", redraw = False)
        if isinstance(rows, set):
            _rows = rows
        else:
            _rows = set(rows)
        self.MT.data[:] = [row for r, row in enumerate(self.MT.data) if r not in _rows]
        rhs = tuple(int(b - a) for a, b in zip(self.MT.row_positions, islice(self.MT.row_positions, 1, len(self.MT.row_positions))))
        if self.MT.all_rows_displayed:
            self.set_row_heights(row_heights = tuple(h for r, h in enumerate(rhs) if r not in _rows))
        else:
            dispset = set(self.MT.displayed_rows)
            heights_to_del = {i for i, r in enumerate(to_bis) if r in dispset}
            if heights_to_del:
                self.set_row_heights(row_heights = tuple(h for r, h in enumerate(rhs) if r not in heights_to_del))
            self.MT.displayed_rows = [r for r in self.MT.displayed_rows if r not in _rows]
        to_bis = sorted(_rows)
        self.MT.cell_options = {(r if not bisect.bisect_left(to_bis, r) else r - bisect.bisect_left(to_bis, r), c): v for (r, c), v in self.MT.cell_options.items() if r not in _rows}
        self.MT.row_options = {r if not bisect.bisect_left(to_bis, r) else r - bisect.bisect_left(to_bis, r): v for r, v in self.MT.row_options.items() if r not in _rows}
        self.RI.cell_options = {r if not bisect.bisect_left(to_bis, r) else r - bisect.bisect_left(to_bis, r): v for r, v in self.RI.cell_options.items() if r not in _rows}
        self.set_refresh_timer(redraw)

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
        if number > len(self.MT.data):
            if mod_positions:
                height = self.MT.GetLinesHeight(int(self.MT.default_rh[0]))
                for r in range(number - len(self.MT.data)):
                    self.MT.insert_row_position("end", height)
        elif number < len(self.MT.data):
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

    def move_row_position(self, row: int, moveto: int):
        self.MT.move_row_position(row, moveto)

    def move_row(self, row: int, moveto: int):
        self.move_rows(moveto, row, 1)

    def move_rows(self, moveto: int, to_move_min: int, number_of_rows: int, move_data: bool = True, index_type: str = "displayed", create_selections: bool = True, redraw = False):
        if index_type.lower() == "displayed" or self.MT.all_rows_displayed:
            new_selected, dispset = self.MT.move_rows_adjust_options_dict(moveto, to_move_min, number_of_rows, move_data, create_selections)
        else:
            new_selected, dispset = self.MT.move_rows_adjust_options_dict(moveto, to_move_min, number_of_rows, move_data, create_selections)
        self.set_refresh_timer(redraw)
        return new_selected, dispset

    def delete_column_position(self, idx: int, deselect_all = False):
        self.MT.del_col_position(idx,
                                 deselect_all = deselect_all)

    def delete_column(self, idx = 0, deselect_all = False, redraw = True):
        self.delete_columns(columns = {idx}, deselect_all = deselect_all, redraw = redraw)
        
    def delete_columns(self, columns: set = set(), deselect_all = False, redraw = True):
        if deselect_all:
            self.deselect("all", redraw = False)
        if isinstance(columns, set):
            to_del = columns
        else:
            to_del = set(columns)
        self.MT.data[:] = [[e for c, e in enumerate(r) if c not in to_del] for r in self.MT.data]
        to_bis = sorted(to_del)
        cws = tuple(int(b - a) for a, b in zip(self.MT.col_positions, islice(self.MT.col_positions, 1, len(self.MT.col_positions))))
        if self.MT.all_columns_displayed:
            self.set_column_widths(column_widths = tuple(w for c, w in enumerate(cws) if c not in to_del))
        else:
            dispset = set(self.MT.displayed_columns)
            widths_to_del = {i for i, c in enumerate(to_bis) if c in dispset}
            if widths_to_del:
                self.set_column_widths(column_widths = tuple(w for c, w in enumerate(cws) if c not in widths_to_del))
            self.MT.displayed_columns = [c if not bisect.bisect_left(to_bis, c) else c - bisect.bisect_left(to_bis, c) for c in self.MT.displayed_columns if c not in to_del]
        self.MT.cell_options = {(r, c if not bisect.bisect_left(to_bis, c) else c - bisect.bisect_left(to_bis, c)): v for (r, c), v in self.MT.cell_options.items() if c not in to_del}
        self.MT.col_options = {c if not bisect.bisect_left(to_bis, c) else c - bisect.bisect_left(to_bis, c): v for c, v in self.MT.col_options.items() if c not in to_del}
        self.CH.cell_options = {c if not bisect.bisect_left(to_bis, c) else c - bisect.bisect_left(to_bis, c): v for c, v in self.CH.cell_options.items() if c not in to_del}
        self.set_refresh_timer(redraw)

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

    def move_column_position(self, column: int, moveto: int):
        self.MT.move_col_position(column, moveto)

    def move_column(self, column: int, moveto: int):
        self.move_columns(moveto, column, 1)

    def move_columns(self, moveto: int, to_move_min: int, number_of_columns: int, move_data: bool = True, index_type: str = "displayed", create_selections: bool = True, redraw = False):
        if index_type.lower() == "displayed" or self.MT.all_columns_displayed:
            new_selected, dispset = self.MT.move_columns_adjust_options_dict(moveto, to_move_min, number_of_columns, move_data, create_selections)
            self.set_refresh_timer(redraw)
            return new_selected, dispset
        else:
            c = int(moveto)
            to_move_max = to_move_min + number_of_columns
            totalcols = int(number_of_columns)
            to_del = to_move_max + number_of_columns
            orig_selected = list(range(to_move_min, to_move_min + totalcols))
            num_data_cols = self.MT.total_data_cols()
            if c + totalcols > num_data_cols:
                new_selected = tuple(range(num_data_cols - totalcols, num_data_cols))
            else:
                if to_move_min > c:
                    new_selected = tuple(range(c, c + totalcols))
                else:
                    new_selected = tuple(range(c + 1 - totalcols, c + 1))
            newcolsdct = {t1: t2 for t1, t2 in zip(orig_selected, new_selected)}
            if to_move_min > c:
                for rn in range(len(self.MT.data)):
                    if len(self.MT.data[rn]) < to_move_max:
                        self.MT.data[rn].extend(list(repeat("", to_move_max - len(self.MT.data[rn]) + 1)))
                    self.MT.data[rn][c:c] = self.MT.data[rn][to_move_min:to_move_max]
                    self.MT.data[rn][to_move_max:to_del] = []
                if isinstance(self.MT._headers, list) and self.MT._headers:
                    if len(self.MT._headers) < to_move_max:
                        self.MT._headers.extend(list(repeat("", to_move_max - len(self.MT._headers) + 1)))
                    self.MT._headers[c:c] = self.MT._headers[to_move_min:to_move_max]
                    self.MT._headers[to_move_max:to_del] = []
                self.CH.cell_options = {
                    int(newcolsdct[k]) if k in newcolsdct else
                    k + totalcols if k < to_move_min and k >= c else
                    int(k): v for k, v in self.CH.cell_options.items()
                }
                self.MT.cell_options = {
                    (k[0], int(newcolsdct[k[1]])) if k[1] in newcolsdct else
                    (k[0], k[1] + totalcols) if k[1] < to_move_min and k[1] >= c else
                    k: v for k, v in self.MT.cell_options.items()
                }
                self.MT.col_options = {
                    int(newcolsdct[k]) if k in newcolsdct else
                    k + totalcols if k < to_move_min and k >= c else
                    int(k): v for k, v in self.MT.col_options.items()
                }
                self.MT.displayed_columns = sorted(int(newcolsdct[k]) if k in newcolsdct else k + totalcols if k < to_move_min and k >= c else int(k) for k in self.MT.displayed_columns)
            else:
                c += 1
                if move_data:
                    for rn in range(len(self.MT.data)):
                        if len(self.MT.data[rn]) < c - 1:
                            self.MT.data[rn].extend(list(repeat("", c - len(self.MT.data[rn]))))
                        self.MT.data[rn][c:c] = self.MT.data[rn][to_move_min:to_move_max]
                        self.MT.data[rn][to_move_min:to_move_max] = []
                    if isinstance(self.MT._headers, list) and self.MT._headers:
                        if len(self.MT._headers) < c:
                            self.MT._headers.extend(list(repeat("", c - len(self.MT._headers))))
                        self.MT._headers[c:c] = self.MT._headers[to_move_min:to_move_max]
                        self.MT._headers[to_move_min:to_move_max] = []
                self.CH.cell_options = {
                    int(newcolsdct[k]) if k in newcolsdct else
                    k - totalcols if k < c and k > to_move_min else
                    int(k): v for k, v in self.CH.cell_options.items()
                }
                self.MT.cell_options = {
                    (k[0], int(newcolsdct[k[1]])) if k[1] in newcolsdct else
                    (k[0], k[1] - totalcols) if k[1] < c and k[1] > to_move_min else
                    k: v for k, v in self.MT.cell_options.items()
                }
                self.MT.col_options = {
                    int(newcolsdct[k]) if k in newcolsdct else
                    k - totalcols if k < c and k > to_move_min else
                    int(k): v for k, v in self.MT.col_options.items()
                }
                self.MT.displayed_columns = sorted(int(newcolsdct[k]) if k in newcolsdct else k - totalcols if k < c and k > to_move_min else int(k) for k in self.MT.displayed_columns)
            self.set_refresh_timer(redraw)
            return new_selected, {}

    # works on currently selected box
    def open_cell(self, ignore_existing_editor = True):
        self.MT.open_cell(event = GeneratedMouseEvent(), ignore_existing_editor = ignore_existing_editor)
        
    def open_header_cell(self, ignore_existing_editor = True):
        self.CH.open_cell(event = GeneratedMouseEvent(), ignore_existing_editor = ignore_existing_editor)
        
    def open_index_cell(self, ignore_existing_editor = True):
        self.RI.open_cell(event = GeneratedMouseEvent(), ignore_existing_editor = ignore_existing_editor)

    def set_text_editor_value(self, text = "", r = None, c = None):
        if self.MT.text_editor is not None and r is None and c is None:
            self.MT.text_editor.set_text(text)
        elif self.MT.text_editor is not None and self.MT.text_editor_loc == (r, c):
            self.MT.text_editor.set_text(text)

    def bind_text_editor_set(self, func, row, column):
        self.MT.bind_text_editor_destroy(func, row, column)

    def destroy_text_editor(self, event = None):
        self.MT.destroy_text_editor(event = event)

    def get_text_editor_widget(self, event = None):
        try:
            return self.MT.text_editor.textedit
        except:
            return None

    def bind_key_text_editor(self, key: str, function):
        self.MT.text_editor_user_bound_keys[key] = function

    def unbind_key_text_editor(self, key: str):
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

    def see(self, row = 0, column = 0, keep_yscroll = False, keep_xscroll = False, bottom_right_corner = False, check_cell_visibility = True, redraw = True):
        self.MT.see(row, column, keep_yscroll, keep_xscroll, bottom_right_corner, check_cell_visibility = check_cell_visibility, redraw = redraw)

    def select_row(self, row, redraw = True):
        self.RI.select_row(int(row) if not isinstance(row, int) else row, redraw = redraw)

    def select_column(self, column, redraw = True):
        self.CH.select_col(int(column) if not isinstance(column, int) else column, redraw = redraw)

    def select_cell(self, row, column, redraw = True):
        self.MT.select_cell(int(row) if not isinstance(row, int) else row,
                            int(column) if not isinstance(column, int) else column,
                            redraw = redraw)

    def select_all(self, redraw = True, run_binding_func = True):
        self.MT.select_all(redraw = redraw, run_binding_func = run_binding_func)

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

    # (row, column, type_) e.g. (0, 0, "column") as a named tuple
    def get_currently_selected(self):
        return self.MT.currently_selected()

    def set_currently_selected(self, row, column, type_ = "cell", selection_binding = True):
        self.MT.create_current(r = row,
                               c = column,
                               type_ = type_,
                               inside = True if self.MT.cell_selected(row, column) else False)
        if selection_binding and self.MT.selection_binding_func is not None:
            self.MT.selection_binding_func(SelectCellEvent("select_cell", row, column))

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
        return self.MT.create_selected(r1 = r1, c1 = c1, r2 = r2, c2 = c2, type_ = "columns" if type_ == "cols" else type_)

    def recreate_all_selection_boxes(self):
        self.MT.recreate_all_selection_boxes()
        
    def cell_visible(self, r, c):
        return self.MT.cell_visible(r, c)
        
    def cell_completely_visible(self, r, c, seperate_axes = False):
        return self.MT.cell_completely_visible(r, c, seperate_axes)

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

    def readonly_rows(self, rows = [], readonly = True, redraw = False):
        self.MT.readonly_rows(rows = rows,
                              readonly = readonly)
        if redraw:
            self.redraw()

    def readonly_columns(self, columns = [], readonly = True, redraw = False):
        self.MT.readonly_columns(columns = columns,
                                 readonly = readonly)
        if redraw:
            self.redraw()

    def readonly_cells(self, row = 0, column = 0, cells = [], readonly = True, redraw = False):
        self.MT.readonly_cells(row = row,
                               column = column,
                               cells = cells,
                               readonly = readonly)
        if redraw:
            self.redraw()
            
    def readonly_header(self, columns = [], readonly = True, redraw = False):
        self.CH.readonly_header(columns = columns,
                                readonly = readonly)
        if redraw:
            self.redraw()
            
    def readonly_index(self, rows = [], readonly = True, redraw = False):
        self.RI.readonly_index(rows = rows,
                               readonly = readonly)
        if redraw:
            self.redraw()

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
            
    def highlight_rows(self, rows = [], bg = None, fg = None, highlight_index = True, redraw = False, end_of_screen = False, overwrite = True):
        self.MT.highlight_rows(rows = rows,
                               bg = bg,
                               fg = fg,
                               highlight_index = highlight_index,
                               redraw = redraw,
                               end_of_screen = end_of_screen,
                               overwrite = overwrite)

    def highlight_columns(self, columns = [], bg = None, fg = None, highlight_header = True, redraw = False, overwrite = True):
        self.MT.highlight_cols(cols = columns,
                               bg = bg,
                               fg = fg,
                               highlight_header = highlight_header,
                               redraw = redraw,
                               overwrite = overwrite)

    def highlight_cells(self, row = 0, column = 0, cells = [], canvas = "table", bg = None, fg = None, redraw = False, overwrite = True):
        if canvas == "table":
            self.MT.highlight_cells(r = row,
                                    c = column,
                                    cells = cells,
                                    bg = bg,
                                    fg = fg,
                                    redraw = redraw,
                                    overwrite = overwrite)
        elif canvas in ("row_index", "index"):
            self.RI.highlight_cells(r = row,
                                    cells = cells,
                                    bg = bg,
                                    fg = fg,
                                    redraw = redraw,
                                    overwrite = overwrite)
        elif canvas == "header":
            self.CH.highlight_cells(c = column,
                                    cells = cells,
                                    bg = bg,
                                    fg = fg,
                                    redraw = redraw,
                                    overwrite = overwrite)

    def dehighlight_cells(self, row = 0, column = 0, cells = [], canvas = "table", all_ = False, redraw = True):
        if row == "all" and canvas == "table":
            for k, v in self.MT.cell_options.items():
                if 'highlight' in v:
                    del self.MT.cell_options[k]['highlight']
        elif row == "all" and canvas == "row_index":
            for k, v in self.RI.cell_options.items():
                if 'highlight' in v:
                    del self.RI.cell_options[k]['highlight']
        elif row == "all" and canvas == "header":
            for k, v in self.CH.cell_options.items():
                if 'highlight' in v:
                    del self.CH.cell_options[k]['highlight']
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
            
    def delete_out_of_bounds_options(self):
        maxc = self.total_columns()
        maxr = self.total_rows()
        self.MT.cell_options = {k: v for k, v in self.MT.cell_options.items() if k[0] < maxr and k[1] < maxc}
        self.RI.cell_options = {k: v for k, v in self.RI.cell_options.items() if k < maxr}
        self.CH.cell_options = {k: v for k, v in self.CH.cell_options.items() if k < maxc}
        self.MT.col_options = {k: v for k, v in self.MT.col_options.items() if k < maxc}
        self.MT.row_options = {k: v for k, v in self.MT.row_options.items() if k < maxr}
        
    def reset_all_options(self):
        self.MT.cell_options = {}
        self.RI.cell_options = {}
        self.CH.cell_options = {}
        self.MT.col_options = {}
        self.MT.row_options = {}

    def get_cell_options(self, canvas = "table"):
        if canvas == "table":
            return self.MT.cell_options
        elif canvas == "row_index":
            return self.RI.cell_options
        elif canvas == "header":
            return self.CH.cell_options

    def get_highlighted_cells(self, canvas = "table"):
        if canvas == "table":
            return {k: v['highlight'] for k, v in self.MT.cell_options.items() if 'highlight' in v}
        elif canvas == "row_index":
            return {k: v['highlight'] for k, v in self.RI.cell_options.items() if 'highlight' in v}
        elif canvas == "header":
            return {k: v['highlight'] for k, v in self.CH.cell_options.items() if 'highlight' in v}
        
    def get_frame_y(self, y: int):
        return y + self.CH.current_height

    def get_frame_x(self, x: int):
        return x + self.RI.current_width
    
    def convert_align(self, align: str):
        a = align.lower()
        if a in ("c", "center", "centre"):
            return "center"
        elif a in ("w", "west", "left"):
            return "w"
        elif a in ("e", "east", "right"):
            return "e"
        else:
            raise ValueError("Align must be one of the following values: c, center, w, west, left, e, east, right")
        
    def get_cell_alignments(self):
        return {(r, c): v['align'] for (r, c), v in self.MT.cell_options.items() if 'align' in v}
    
    def get_column_alignments(self):
        return {c: v['align'] for c, v in self.MT.col_options.items() if 'align' in v}
    
    def get_row_alignments(self):
        return {r: v['align'] for r, v in self.MT.row_options.items() if 'align' in v}
        
    def align_rows(self, rows = [], align = "global", align_index = False, redraw = True): #"center", "w", "e" or "global"
        if align == "global" or self.convert_align(align):
            if isinstance(rows, dict):
                for k, v in rows.items():
                    self.MT.align_rows(rows = k,
                                       align = v,
                                       align_index = align_index)
            else:
                self.MT.align_rows(rows = rows,
                                   align = align if align == "global" else self.convert_align(align),
                                   align_index = align_index)
        if redraw:
            self.redraw()
        
    def align_columns(self, columns = [], align = "global", align_header = False, redraw = True): #"center", "w", "e" or "global"
        if align == "global" or self.convert_align(align):
            if isinstance(columns, dict):
                for k, v in columns.items():
                    self.MT.align_columns(columns = k,
                                          align = v,
                                          align_header = align_header)
            else:
                self.MT.align_columns(columns = columns,
                                      align = align if align == "global" else self.convert_align(align),
                                      align_header = align_header)
        if redraw:
            self.redraw()

    def align_cells(self, row = 0, column = 0, cells = [], align = "global", redraw = True): #"center", "w", "e" or "global"
        if align == "global" or self.convert_align(align):
            if isinstance(cells, dict):
                for (r, c), v in cells.items():
                    self.MT.align_cells(row = r,
                                        column = c,
                                        cells = [],
                                        align = v)
            else:
                self.MT.align_cells(row = row,
                                    column = column,
                                    cells = cells,
                                    align = align if align == "global" else self.convert_align(align))
        if redraw:
            self.redraw()

    def align_header(self, columns = [], align = "global", redraw = True):
        if align == "global" or self.convert_align(align):
            if isinstance(columns, dict):
                for k, v in columns.items():
                    self.CH.align_cells(columns = k,
                                        align = v)
            else:
                self.CH.align_cells(columns = columns,
                                    align = align if align == "global" else self.convert_align(align))
        if redraw:
            self.redraw()

    def align_index(self, rows = [], align = "global", redraw = True):
        if align == "global" or self.convert_align(align):
            if isinstance(rows, dict):
                for k, v in rows.items():
                    self.RI.align_cells(rows = rows,
                                        align = v)
            else:
                self.RI.align_cells(rows = rows,
                                    align = align if align == "global" else self.convert_align(align))
        if redraw:
            self.redraw()

    def align(self, align: str = None, redraw = True):
        if align is None:
            return self.MT.align
        elif self.convert_align(align):
            self.MT.align = self.convert_align(align)
        else:
            raise ValueError("Align must be one of the following values: c, center, w, west, e, east")
        if redraw:
            self.refresh()

    def header_align(self, align: str = None, redraw = True):
        if align is None:
            return self.CH.align
        elif self.convert_align(align):
            self.CH.align = self.convert_align(align)
        else:
            raise ValueError("Align must be one of the following values: c, center, w, west, e, east")
        if redraw:
            self.refresh()

    def row_index_align(self, align: str = None, redraw = True):
        if align is None:
            return self.RI.align
        elif self.convert_align(align):
            self.RI.align = self.convert_align(align)
        else:
            raise ValueError("Align must be one of the following values: c, center, w, west, e, east")
        if redraw:
            self.refresh()

    def font(self, newfont = None, reset_row_positions = True):
        self.MT.font(newfont, reset_row_positions = reset_row_positions)

    def header_font(self, newfont = None):
        self.MT.header_font(newfont)

    def set_options(self,
                    redraw = True,
                    **kwargs):
        if 'show_dropdown_borders' in kwargs:
            self.MT.show_dropdown_borders = kwargs['show_dropdown_borders']
        if 'edit_cell_validation' in kwargs:
            self.MT.edit_cell_validation = kwargs['edit_cell_validation']
        if 'ctrl_keys_over_dropdowns_enabled' in kwargs:
            self.MT.ctrl_keys_over_dropdowns_enabled = kwargs['ctrl_keys_over_dropdowns_enabled']
        if 'show_default_header_for_empty' in kwargs:
            self.CH.show_default_header_for_empty = kwargs['show_default_header_for_empty']
        if 'show_default_index_for_empty' in kwargs:
            self.RI.show_default_index_for_empty = kwargs['show_default_index_for_empty']
        if 'selected_rows_to_end_of_window' in kwargs:
            self.MT.selected_rows_to_end_of_window = kwargs['selected_rows_to_end_of_window']
        if 'horizontal_grid_to_end_of_window' in kwargs:
            self.MT.horizontal_grid_to_end_of_window = kwargs['horizontal_grid_to_end_of_window']
        if 'vertical_grid_to_end_of_window' in kwargs:
            self.MT.vertical_grid_to_end_of_window = kwargs['vertical_grid_to_end_of_window']
        if 'paste_insert_column_limit' in kwargs:
            self.MT.paste_insert_column_limit = kwargs['paste_insert_column_limit']
        if 'paste_insert_row_limit' in kwargs:
            self.MT.paste_insert_row_limit = kwargs['paste_insert_row_limit']
        if 'expand_sheet_if_paste_too_big' in kwargs:
            self.MT.expand_sheet_if_paste_too_big = kwargs['expand_sheet_if_paste_too_big']
        if 'arrow_key_down_right_scroll_page' in kwargs:
            self.MT.arrow_key_down_right_scroll_page = kwargs['arrow_key_down_right_scroll_page']
        if 'enable_edit_cell_auto_resize' in kwargs:
            self.MT.cell_auto_resize_enabled = kwargs['enable_edit_cell_auto_resize']
        if 'header_hidden_columns_expander_bg' in kwargs:
            self.CH.header_hidden_columns_expander_bg = kwargs['header_hidden_columns_expander_bg']
        if 'index_hidden_rows_expander_bg' in kwargs:
            self.RI.index_hidden_rows_expander_bg = kwargs['index_hidden_rows_expander_bg']
        if 'page_up_down_select_row' in kwargs:
            self.MT.page_up_down_select_row = kwargs['page_up_down_select_row']
        if 'display_selected_fg_over_highlights' in kwargs:
            self.MT.display_selected_fg_over_highlights = kwargs['display_selected_fg_over_highlights']
        if 'show_horizontal_grid' in kwargs:
            self.MT.show_horizontal_grid = kwargs['show_horizontal_grid']
        if 'show_vertical_grid' in kwargs:
            self.MT.show_vertical_grid = kwargs['show_vertical_grid']
        if 'empty_horizontal' in kwargs:
            self.MT.empty_horizontal = kwargs['empty_horizontal']
        if 'empty_vertical' in kwargs:
            self.MT.empty_vertical = kwargs['empty_vertical']
        if 'row_height' in kwargs:
            self.MT.default_rh = (kwargs['row_height'] if isinstance(kwargs['row_height'], str) else "pixels", kwargs['row_height'] if isinstance(kwargs['row_height'], int) else self.MT.GetLinesHeight(int(kwargs['row_height'])))
        if 'column_width' in kwargs:
            self.MT.default_cw = self.MT.min_cw + 20 if kwargs['column_width'] < self.MT.min_cw else int(kwargs['column_width'])
        if 'header_height' in kwargs:
            self.MT.default_hh = (kwargs['header_height'] if isinstance(kwargs['header_height'], str) else "pixels", kwargs['header_height'] if isinstance(kwargs['header_height'], int) else self.MT.GetHdrLinesHeight(int(kwargs['header_height'])))
        if 'row_drag_and_drop_perform' in kwargs:
            self.RI.row_drag_and_drop_perform = kwargs['row_drag_and_drop_perform']
        if 'column_drag_and_drop_perform' in kwargs:
            self.CH.column_drag_and_drop_perform = kwargs['column_drag_and_drop_perform']
        if 'popup_menu_font' in kwargs:
            self.MT.popup_menu_font = kwargs['popup_menu_font']
        if 'popup_menu_fg' in kwargs:
            self.MT.popup_menu_fg = kwargs['popup_menu_fg']
        if 'popup_menu_bg' in kwargs:
            self.MT.popup_menu_bg = kwargs['popup_menu_bg']
        if 'popup_menu_highlight_bg' in kwargs:
            self.MT.popup_menu_highlight_bg = kwargs['popup_menu_highlight_bg']
        if 'popup_menu_highlight_fg' in kwargs:
            self.MT.popup_menu_highlight_fg = kwargs['popup_menu_highlight_fg']
        if 'top_left_fg_highlight' in kwargs:
            self.TL.top_left_fg_highlight = kwargs['top_left_fg_highlight']
        if 'auto_resize_default_row_index' in kwargs:
            self.RI.auto_resize_width = kwargs['auto_resize_default_row_index']
        if 'header_selected_columns_bg' in kwargs:
            self.CH.header_selected_columns_bg = kwargs['header_selected_columns_bg']
        if 'header_selected_columns_fg' in kwargs:
            self.CH.header_selected_columns_fg = kwargs['header_selected_columns_fg']
        if 'index_selected_rows_bg' in kwargs:
            self.RI.index_selected_rows_bg = kwargs['index_selected_rows_bg']
        if 'index_selected_rows_fg' in kwargs:
            self.RI.index_selected_rows_fg = kwargs['index_selected_rows_fg']
        if 'table_selected_rows_border_fg' in kwargs:
            self.MT.table_selected_rows_border_fg = kwargs['table_selected_rows_border_fg']
        if 'table_selected_rows_bg' in kwargs:
            self.MT.table_selected_rows_bg = kwargs['table_selected_rows_bg']
        if 'table_selected_rows_fg' in kwargs:
            self.MT.table_selected_rows_fg = kwargs['table_selected_rows_fg']
        if 'table_selected_columns_border_fg' in kwargs:
            self.MT.table_selected_columns_border_fg = kwargs['table_selected_columns_border_fg']
        if 'table_selected_columns_bg' in kwargs:
            self.MT.table_selected_columns_bg = kwargs['table_selected_columns_bg']
        if 'table_selected_columns_fg' in kwargs:
            self.MT.table_selected_columns_fg = kwargs['table_selected_columns_fg']
        if 'default_header' in kwargs:
            self.CH.default_hdr = kwargs['default_header'].lower()
        if 'default_row_index' in kwargs:
            self.RI.default_index = kwargs['default_row_index'].lower()
        if 'max_colwidth' in kwargs:
            self.CH.max_cw = float(kwargs['max_colwidth'])
        if 'max_row_height' in kwargs:
            self.RI.max_rh = float(kwargs['max_row_height'])
        if 'max_header_height' in kwargs:
            self.CH.max_header_height = float(kwargs['max_header_height'])
        if 'max_row_width' in kwargs:
            self.RI.max_row_width = float(kwargs['max_row_width'])
        if 'font' in kwargs:
            self.MT.font(kwargs['font'])
        if 'header_font' in kwargs:
            self.MT.header_font(kwargs['header_font'])
        if 'theme' in kwargs:
            self.change_theme(kwargs['theme'])
        if 'show_selected_cells_border' in kwargs:
            self.MT.show_selected_cells_border = kwargs['show_selected_cells_border']
        if 'header_bg' in kwargs:
            self.CH.config(background = kwargs['header_bg'])
            self.CH.header_bg = kwargs['header_bg']
        if 'header_border_fg' in kwargs:
            self.CH.header_border_fg = kwargs['header_border_fg']
        if 'header_grid_fg' in kwargs:
            self.CH.header_grid_fg = kwargs['header_grid_fg']
        if 'header_fg' in kwargs:
            self.CH.header_fg = kwargs['header_fg']
        if 'header_selected_cells_bg' in kwargs:
            self.CH.header_selected_cells_bg = kwargs['header_selected_cells_bg']
        if 'header_selected_cells_fg' in kwargs:
            self.CH.header_selected_cells_fg = kwargs['header_selected_cells_fg']
        if 'index_bg' in kwargs:
            self.RI.config(background = kwargs['index_bg'])
            self.RI.index_bg = kwargs['index_bg']
        if 'index_border_fg' in kwargs:
            self.RI.index_border_fg = kwargs['index_border_fg']
        if 'index_grid_fg' in kwargs:
            self.RI.index_grid_fg = kwargs['index_grid_fg']
        if 'index_fg' in kwargs:
            self.RI.index_fg = kwargs['index_fg']
        if 'index_selected_cells_bg' in kwargs:
            self.RI.index_selected_cells_bg = kwargs['index_selected_cells_bg']
        if 'index_selected_cells_fg' in kwargs:
            self.RI.index_selected_cells_fg = kwargs['index_selected_cells_fg']
        if 'top_left_bg' in kwargs:
            self.TL.config(background = kwargs['top_left_bg'])
        if 'top_left_fg' in kwargs:
            self.TL.top_left_fg = kwargs['top_left_fg']
            self.TL.itemconfig("rw", fill = kwargs['top_left_fg'])
            self.TL.itemconfig("rh", fill = kwargs['top_left_fg'])
        if 'frame_bg' in kwargs:
            self.config(background = kwargs['frame_bg'])
        if 'table_bg' in kwargs:
            self.MT.config(background = kwargs['table_bg'])
            self.MT.table_bg = kwargs['table_bg']
        if 'table_grid_fg' in kwargs:
            self.MT.table_grid_fg = kwargs['table_grid_fg']
        if 'table_fg' in kwargs:
            self.MT.table_fg = kwargs['table_fg']
        if 'table_selected_cells_border_fg' in kwargs:
            self.MT.table_selected_cells_border_fg = kwargs['table_selected_cells_border_fg']
        if 'table_selected_cells_bg' in kwargs:
            self.MT.table_selected_cells_bg = kwargs['table_selected_cells_bg']
        if 'table_selected_cells_fg' in kwargs:
            self.MT.table_selected_cells_fg = kwargs['table_selected_cells_fg']
        if 'resizing_line_fg' in kwargs:
            self.CH.resizing_line_fg = kwargs['resizing_line_fg']
            self.RI.resizing_line_fg = kwargs['resizing_line_fg']
        if 'drag_and_drop_bg' in kwargs:
            self.CH.drag_and_drop_bg = kwargs['drag_and_drop_bg']
            self.RI.drag_and_drop_bg = kwargs['drag_and_drop_bg']
        if 'outline_thickness' in kwargs:
            self.config(highlightthickness = kwargs['outline_thickness'])
        if 'outline_color' in kwargs:
            self.config(highlightbackground = kwargs['outline_color'], highlightcolor = kwargs['outline_color'])
        self.MT.create_rc_menus()
        if redraw:
            self.refresh()

    def change_theme(self, theme = "light blue", redraw = True):
        if theme.lower() in ("light blue", "light_blue"):
            self.set_options(**theme_light_blue,
                             redraw = redraw)
            self.config(bg = theme_light_blue['table_bg'])
        elif theme.lower() == "dark":
            self.set_options(**theme_dark,
                             redraw = redraw)
            self.config(bg = theme_dark['table_bg'])
        elif theme.lower() in ("light green", "light_green"):
            self.set_options(**theme_light_green,
                             redraw = redraw)
            self.config(bg = theme_light_green['table_bg'])
        elif theme.lower() in ("dark blue", "dark_blue"):
            self.set_options(**theme_dark_blue,
                             redraw = redraw)
            self.config(bg = theme_dark_blue['table_bg'])
        elif theme.lower() in ("dark green", "dark_green"):
            self.set_options(**theme_dark_green,
                             redraw = redraw)
            self.config(bg = theme_dark_green['table_bg'])
        elif theme.lower() == "black":
            self.set_options(**theme_black,
                             redraw = redraw)
            self.config(bg = theme_black['table_bg'])
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
                raise ValueError("Data argument must be a list of lists, sublists being rows")
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
                index_limit = len(self.MT._row_index)
                return [[""] + self.MT._headers.copy()] + [[f"{self.MT._row_index[rn]}"] + r.copy() if rn < index_limit else [""] + r.copy() for rn, r in enumerate(self.MT.data)]
            elif get_header and not get_index:
                return [self.MT._headers.copy()] + [r.copy() for r in self.MT.data]
            elif get_index and not get_header:
                index_limit = len(self.MT._row_index)
                return [[f"{self.MT._row_index[rn]}"] + r.copy() if rn < index_limit else [""] + r.copy() for rn, r in enumerate(self.MT.data)]
            elif not get_index and not get_header:
                return [r.copy() for r in self.MT.data]
        else:
            if get_header and get_index:
                index_limit = len(self.MT._row_index)
                return [[""] + self.MT._headers] + [[self.MT._row_index[rn]] + r if rn < index_limit else [""] + r for rn, r in enumerate(self.MT.data)]
            elif get_header and not get_index:
                return [self.MT._headers] + self.MT.data
            elif get_index and not get_header:
                index_limit = len(self.MT._row_index)
                return [[self.MT._row_index[rn]] + r if rn < index_limit else [""] + r for rn, r in enumerate(self.MT.data)]
            elif not get_index and not get_header:
                return self.MT.data

    @property
    def data(self):
        return self.MT.data
            
    def yield_sheet_rows(self, get_header = False, get_index = False):
        if get_header:
            if get_index:
                yield [""] + [self.get_n2a(c, self.CH.default_hdr) for c in range(len(max(self.MT.data, key = len)))] if isinstance(self.MT._headers, int) or not self.MT._headers else self.MT._headers
            else:
                yield [self.get_n2a(c, self.CH.default_hdr) for c in range(len(max(self.MT.data, key = len)))] if isinstance(self.MT._headers, int) or not self.MT._headers else self.MT._headers
        if get_index:
            if isinstance(self.MT._row_index, int) or not self.MT._row_index:
                for rn, r in enumerate(self.MT.data):
                    yield [self.get_n2a(rn, self.RI.default_index)] + r
            else:
                index_limit = len(self.MT._row_index)
                for rn, r in enumerate(self.MT.data):
                    yield [self.MT._row_index[rn]] + r if rn < index_limit else [""] + r
        else:
            yield from self.MT.data
                
    def get_n2a(self, n = 0, _type = "numbers"):
        if _type == "letters":
            return num2alpha(n)
        elif _type == "numbers":
            return f"{n + 1}"
        else:
            return f"{num2alpha(n)} {n + 1}"

    def get_cell_data(self, r, c, return_copy = True):
        if return_copy:
            try:
                return f"{self.MT.data[r][c]}"
            except:
                return None
        else:
            try:
                return self.MT.data[r][c]
            except:
                return None

    def get_row_data(self, r, return_copy = True):
        if return_copy:
            try:
                return tuple(f"{e}" for e in self.MT.data[r])
            except:
                return None
        else:
            try:
                return self.MT.data[r]
            except:
                return None

    def get_column_data(self, c, return_copy = True):
        res = []
        if return_copy:
            for r in self.MT.data:
                try:
                    res.append(f"{r[c]}")
                except:
                    continue
            return tuple(res)
        else:
            for r in self.MT.data:
                try:
                    res.append(r[c])
                except:
                    continue
            return res

    def set_cell_data(self, r, c, value = "", set_copy = True, redraw = False):
        self.MT.data[r][c] = f"{value}" if set_copy else value
        self.set_refresh_timer(redraw)

    def set_column_data(self, c, values = tuple(), add_rows = True, redraw = False):
        if add_rows:
            maxidx = len(self.MT.data) - 1
            total_cols = None
            height = self.MT.default_rh[1]
            for rn, v in enumerate(values):
                if rn > maxidx:
                    if total_cols is None:
                        total_cols = self.MT.total_data_cols()
                    self.MT.data.append(list(repeat("", total_cols)))
                    self.MT.insert_row_position("end", height = height)
                    maxidx += 1
                if c > len(self.MT.data[rn]) - 1:
                    self.MT.data[rn].extend(list(repeat("", c - len(self.MT.data[rn]))))
                self.MT.data[rn][c] = v
        else:
            for rn, v in enumerate(values):
                if c > len(self.MT.data[rn]) - 1:
                    self.MT.data[rn].extend(list(repeat("", c - len(self.MT.data[rn]))))
                self.MT.data[rn][c] = v
        self.set_refresh_timer(redraw)

    def insert_column(self,
                      values = None, 
                      idx = "end", 
                      width = None, 
                      deselect_all = False, 
                      add_rows = True, 
                      equalize_data_row_lengths = True,
                      mod_column_positions = True,
                      redraw = False):
        if isinstance(values, (list, tuple)):
            _values = (values, )
        elif values is None:
            _values = 1
        else:
            _values = values
        if isinstance(width, int):
            _width = (width, )
        elif width is None:
            _width = width
        self.insert_columns(_values,
                            idx,
                            _width,
                            deselect_all, 
                            add_rows, 
                            equalize_data_row_lengths,
                            mod_column_positions,
                            redraw)

    def insert_columns(self, 
                       columns = 1, 
                       idx = "end", 
                       widths = None, 
                       deselect_all = False, 
                       add_rows = True, 
                       equalize_data_row_lengths = True,
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
        if not self.MT.all_columns_displayed:
            if idx != "end":
                self.MT.displayed_columns = [c if c < idx else c + numcols for c in self.MT.displayed_columns]
            if mod_column_positions:
                try:
                    disp_next = max(self.MT.displayed_columns) + 1
                except:
                    disp_next = 0
                self.MT.displayed_columns.extend(list(range(disp_next, disp_next + numcols)))
        maxidx = len(self.MT.data) - 1
        if add_rows:
            height = self.MT.default_rh[1]
            if idx == "end":
                for values in reversed(data):
                    for rn, v in enumerate(values):
                        if rn > maxidx:
                            self.MT.data.append(list(repeat("", old_total)))
                            self.MT.insert_row_position("end", height = height)
                            maxidx += 1
                        self.MT.data[rn].append(v)
            else:
                for values in reversed(data):
                    for rn, v in enumerate(values):
                        if rn > maxidx:
                            self.MT.data.append(list(repeat("", old_total)))
                            self.MT.insert_row_position("end", height = height)
                            maxidx += 1
                        self.MT.data[rn].insert(idx, v)
        else:
            if idx == "end":
                for values in reversed(data):
                    for rn, v in enumerate(values):
                        if rn > maxidx:
                            break
                        self.MT.data[rn].append(v)
            else:
                for values in reversed(data):
                    for rn, v in enumerate(values):
                        if rn > maxidx:
                            break
                        self.MT.data[rn].insert(idx, v)
        if isinstance(idx, int):
            num_add = len(data)
            self.MT.cell_options = {(rn, cn if cn < idx else cn + num_add): t2 for (rn, cn), t2 in self.MT.cell_options.items()}
            self.MT.col_options = {cn if cn < idx else cn + num_add: t for cn, t in self.MT.col_options.items()}
            self.CH.cell_options = {cn if cn < idx else cn + num_add: t for cn, t in self.CH.cell_options.items()}
        self.set_refresh_timer(redraw)

    def set_row_data(self, r, values = tuple(), add_columns = True, redraw = False):
        if len(self.MT.data) - 1 < r:
            raise Exception("Row number is out of range")
        maxidx = len(self.MT.data[r]) - 1
        if not values:
            self.MT.data[r][:] = list(repeat("", len(self.MT.data[r])))
        if add_columns:
            for c, v in enumerate(values):
                if c > maxidx:
                    self.MT.data[r].append(v)
                    if self.MT.all_columns_displayed and c >= len(self.MT.col_positions) - 1:
                        self.MT.insert_col_position("end")
                else:
                    self.MT.data[r][c] = v
        else:
            for c, v in enumerate(values):
                if c > maxidx:
                    self.MT.data[r].append(v)
                else:
                    self.MT.data[r][c] = v
        self.set_refresh_timer(redraw)

    def insert_row(self, values = None, idx = "end", height = None, deselect_all = False, add_columns = False,
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
                self.MT.data[:] = [r + list(repeat("", len(data) - total_cols)) for r in self.MT.data]
            elif len(data) < total_cols:
                data += list(repeat("", total_cols - len(data)))
            if self.MT.all_columns_displayed:
                if not self.MT.col_positions:
                    self.MT.col_positions = [0]
                if len(data) > len(self.MT.col_positions) - 1:
                    self.insert_column_positions("end", len(data) - (len(self.MT.col_positions) - 1))
        if isinstance(idx, str) and idx.lower() == "end":
            self.MT.data.append(data)
        else:
            self.MT.data.insert(idx, data)
            self.MT.cell_options = {(rn if rn < idx else rn + 1, cn): t2 for (rn, cn), t2 in self.MT.cell_options.items()}
            self.MT.row_options = {rn if rn < idx else rn + 1: t for rn, t in self.MT.row_options.items()}
            self.RI.cell_options = {rn if rn < idx else rn + 1: t for rn, t in self.RI.cell_options.items()}
        self.set_refresh_timer(redraw)

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
                self.MT.data[:] = [r + list(repeat("", data_max_cols - total_cols)) for r in self.MT.data]
            else:
                data[:] = [r + list(repeat("", total_cols - data_max_cols)) for r in data]
            if self.MT.all_columns_displayed:
                if not self.MT.col_positions:
                    self.MT.col_positions = [0]
                if data_max_cols > len(self.MT.col_positions) - 1:
                    self.insert_column_positions("end", data_max_cols - (len(self.MT.col_positions) - 1))
        if isinstance(idx, str) and idx.lower() == "end":
            self.MT.data.extend(data)
        else:
            self.MT.data[idx:idx] = data
            num_add = len(data)
            self.MT.cell_options = {(rn if rn < idx else rn + num_add, cn): t2 for (rn, cn), t2 in self.MT.cell_options.items()}
            self.MT.row_options = {rn if rn < idx else rn + num_add: t for rn, t in self.MT.row_options.items()}
            self.RI.cell_options = {rn if rn < idx else rn + num_add: t for rn, t in self.RI.cell_options.items()}
        self.set_refresh_timer(redraw)

    def sheet_data_dimensions(self, total_rows = None, total_columns = None):
        self.MT.data_dimensions(total_rows, total_columns)

    def get_total_rows(self):
        return len(self.MT.data)

    def equalize_data_row_lengths(self):
        return self.MT.equalize_data_row_lengths()
                    
    def display_columns(self,
                        columns = None,
                        all_columns_displayed = None,
                        reset_col_positions = True,
                        refresh = False,
                        redraw = False,
                        deselect_all = True):
        res = self.MT.display_columns(columns = None if isinstance(columns, str) and columns.lower() == "all" else columns,
                                      all_columns_displayed = True if isinstance(columns, str) and columns.lower() == "all" else all_columns_displayed,
                                      reset_col_positions = reset_col_positions,
                                      deselect_all = deselect_all)
        if refresh or redraw:
            self.refresh()
        return res
    
    def hide_columns(self, columns = set(), refresh = True, deselect_all = True):
        if isinstance(columns, int):
            _columns = {columns}
        elif isinstance(columns, set):
            _columns = columns
        else:
            _columns = set(columns)
        self.display_columns(columns = [c for c in range(self.MT.total_data_cols()) if c not in _columns] if self.MT.all_columns_displayed else [c for c in self.MT.displayed_columns if c not in _columns],
                             all_columns_displayed = False,
                             refresh = refresh, 
                             deselect_all = deselect_all)

    def show_ctrl_outline(self, canvas = "table", start_cell = (0, 0), end_cell = (1, 1)):
        self.MT.show_ctrl_outline(canvas = canvas, start_cell = start_cell, end_cell = end_cell)

    def get_ctrl_x_c_boxes(self):
        return self.MT.get_ctrl_x_c_boxes()

    def get_selected_min_max(self): # returns (min_y, min_x, max_y, max_x) of any selections including rows/columns
        return self.MT.get_selected_min_max()
        
    def headers(self, newheaders = None, index = None, reset_col_positions = False, show_headers_if_not_sheet = True, redraw = False):
        return self.MT.headers(newheaders, index, reset_col_positions = reset_col_positions, show_headers_if_not_sheet = show_headers_if_not_sheet, redraw = redraw)
        
    def row_index(self, newindex = None, index = None, reset_row_positions = False, show_index_if_not_sheet = True, redraw = False):
        return self.MT.row_index(newindex, index, reset_row_positions = reset_row_positions, show_index_if_not_sheet = show_index_if_not_sheet, redraw = redraw)

    def reset_undos(self):
        self.MT.undo_storage = deque(maxlen = self.MT.max_undos)

    def redraw(self, redraw_header = True, redraw_row_index = True):
        self.MT.main_table_redraw_grid_and_text(redraw_header = redraw_header, redraw_row_index = redraw_row_index)

    def refresh(self, redraw_header = True, redraw_row_index = True):
        self.MT.main_table_redraw_grid_and_text(redraw_header = redraw_header, redraw_row_index = redraw_row_index)
        
    def create_header_checkbox(self,
                               c,
                               checked = False,
                               state = "normal",
                               redraw = False,
                               check_function = None,
                               text = ""):
        if isinstance(c, str) and c.lower() == "all":
            for c_ in range(self.MT.total_data_cols()):
                self.CH.create_checkbox(c = c_,
                                        checked = checked,
                                        state = state,
                                        redraw = redraw,
                                        check_function = check_function,
                                        text = text)
        else:
            self.CH.create_checkbox(c = c,
                                    checked = checked,
                                    state = state,
                                    redraw = redraw,
                                    check_function = check_function,
                                    text = text)

    def click_header_checkbox(self,
                              c,
                              checked = None):
        if c in self.CH.cell_options and 'checkbox' in self.CH.cell_options[c]:
            if not type(self.MT._headers[c]) == bool:
                if checked is None:
                    self.MT._headers[c] = False
                else:
                    self.MT._headers[c] = bool(checked)
            else:
                self.MT._headers[c] = not self.MT._headers[c]

    def get_header_checkboxes(self):
        return {k: v['checkbox'] for k, v in self.CH.cell_options.items() if 'checkbox' in v}

    def delete_header_checkbox(self, c = 0):
        if c == "all":
            for c_ in self.CH.cell_options:
                if 'checkbox' in self.CH.cell_options[c_]:
                    self.CH.delete_checkbox(c_)
        else:
            self.CH.delete_checkbox(c)

    def header_checkbox(self,
                        c,
                        checked = None,
                        state = None,
                        check_function = "",
                        text = None):
        if type(checked) == bool:
            self.headers(newheaders = checked, index = c)
        if check_function != "":
            self.CH.cell_options[c]['checkbox']['check_function'] = check_function
        if state and state.lower() in ("normal", "disabled"):
            self.CH.cell_options[c]['checkbox']['state'] = state
        if text is not None:
            self.CH.cell_options[c]['checkbox']['text'] = text
        return {**self.CH.cell_options[c]['checkbox'], 'checked': self.MT._headers[c]}
    
    def create_index_checkbox(self,
                              r,
                              checked = False,
                              state = "normal",
                              redraw = False,
                              check_function = None,
                              text = ""):
        if isinstance(r, str) and r.lower() == "all":
            for r_ in range(self.MT.total_data_rows()):
                self.RI.create_checkbox(r = r_,
                                        checked = checked,
                                        state = state,
                                        redraw = redraw,
                                        check_function = check_function,
                                        text = text)
        else:
            self.RI.create_checkbox(r = r,
                                    checked = checked,
                                    state = state,
                                    redraw = redraw,
                                    check_function = check_function,
                                    text = text)

    def click_index_checkbox(self,
                             r,
                             checked = None):
        if r in self.RI.cell_options and 'checkbox' in self.RI.cell_options[r]:
            if not type(self.MT._row_index[r]) == bool:
                if checked is None:
                    self.MT._row_index[r] = False
                else:
                    self.MT._row_index[r] = bool(checked)
            else:
                self.MT._row_index[r] = not self.MT._row_index[r]

    def get_index_checkboxes(self):
        return {k: v['checkbox'] for k, v in self.RI.cell_options.items() if 'checkbox' in v}

    def delete_index_checkbox(self, r = 0):
        if r == "all":
            for r_ in self.RI.cell_options:
                if 'checkbox' in self.RI.cell_options[r_]:
                    self.RI.delete_checkbox(r_)
        else:
            self.RI.delete_checkbox(r)

    def index_checkbox(self,
                       r,
                       checked = None,
                       state = None,
                       check_function = "",
                       text = None):
        if type(checked) == bool:
            self.row_index(newindex = checked, index = r)
        if check_function != "":
            self.RI.cell_options[r]['checkbox']['check_function'] = check_function
        if state and state.lower() in ("normal", "disabled"):
            self.RI.cell_options[r]['checkbox']['state'] = state
        if text is not None:
            self.RI.cell_options[r]['checkbox']['text'] = text
        return {**self.RI.cell_options[r]['checkbox'], 'checked': self.MT._row_index[r]}

    def create_checkbox(self,
                        r,
                        c,
                        checked = False,
                        state = "normal",
                        redraw = False,
                        check_function = None,
                        text = ""):
        if isinstance(r, str) and r.lower() == "all" and isinstance(c, int):
            for r_ in range(self.MT.total_data_rows()):
                self.MT.create_checkbox(r = r_,
                                        c = c,
                                        checked = checked,
                                        state = state,
                                        redraw = redraw,
                                        check_function = check_function,
                                        text = text)
        elif isinstance(c, str) and c.lower() == "all" and isinstance(r, int):
            for c_ in range(self.MT.total_data_cols()):
                self.MT.create_checkbox(r = r,
                                        c = c_,
                                        checked = checked,
                                        state = state,
                                        redraw = redraw,
                                        check_function = check_function,
                                        text = text)
        elif isinstance(r, str) and r.lower() == "all" and isinstance(c, str) and c.lower() == "all":
            totalcols = self.MT.total_data_cols()
            for r_ in range(self.MT.total_data_rows()):
                for c_ in range(totalcols):
                    self.MT.create_checkbox(r = r_,
                                            c = c_,
                                            checked = checked,
                                            state = state,
                                            redraw = redraw,
                                            check_function = check_function,
                                            text = text)
        else:
            self.MT.create_checkbox(r = r,
                                    c = c,
                                    checked = checked,
                                    state = state,
                                    redraw = redraw,
                                    check_function = check_function,
                                    text = text)

    def click_checkbox(self,
                       r,
                       c,
                       checked = None):
        if (r, c) in self.MT.cell_options and 'checkbox' in self.MT.cell_options[(r, c)]:
            if not type(self.MT.data[r][c]) == bool:
                if checked is None:
                    self.MT.data[r][c] = False
                else:
                    self.MT.data[r][c] = bool(checked)
            else:
                self.MT.data[r][c] = not self.MT.data[r][c]

    def get_checkboxes(self):
        return {k: v['checkbox'] for k, v in self.MT.cell_options.items() if 'checkbox' in v}

    def delete_checkbox(self, r = 0, c = 0):
        if isinstance(r, str) and r.lower() == "all" and isinstance(c, int):
            for r_, c_ in self.MT.cell_options:
                if 'checkbox' in self.MT.cell_options[(r_, c)]:
                    self.MT.delete_checkbox(r_, c)
        elif isinstance(c, str) and c.lower() == "all" and isinstance(r, int):
            for r_, c_ in self.MT.cell_options:
                if 'checkbox' in self.MT.cell_options[(r, c_)]:
                    self.MT.delete_checkbox(r, c_)
        elif isinstance(r, str) and r.lower() == "all" and isinstance(c, str) and c.lower() == "all":
            for r_, c_ in self.MT.cell_options:
                if 'checkbox' in self.MT.cell_options[(r_, c_)]:
                    self.MT.delete_checkbox(r_, c_)
        else:
            self.MT.delete_checkbox(r, c)

    def checkbox(self,
                 r,
                 c,
                 checked = None,
                 state = None,
                 check_function = "",
                 text = None):
        if type(checked) == bool:
            self.set_cell_data(r, c, checked)
        if check_function != "":
            self.MT.cell_options[(r, c)]['checkbox']['check_function'] = check_function
        if state and state.lower() in ("normal", "disabled"):
            self.MT.cell_options[(r, c)]['checkbox']['state'] = state
        if text is not None:
            self.MT.cell_options[(r, c)]['checkbox']['text'] = text
        return {**self.MT.cell_options[(r, c)]['checkbox'], 'checked': self.MT.data[r][c]}
    
    def create_header_dropdown(self,
                               c = 0,
                               values = [],
                               set_value = None,
                               state = "readonly",
                               redraw = False,
                               selection_function = None,
                               modified_function = None):
        if isinstance(c, str) and c.lower() == "all":
            for c_ in range(self.MT.total_data_cols()):
                self.CH.create_dropdown(c = c_,
                                        values = values,
                                        set_value = set_value,
                                        state = state,
                                        redraw = redraw,
                                        selection_function = selection_function,
                                        modified_function = modified_function)
        else:
            self.CH.create_dropdown(c = c,
                                    values = values,
                                    set_value = set_value,
                                    state = state,
                                    redraw = redraw,
                                    selection_function = selection_function,
                                    modified_function = modified_function)

    def get_header_dropdown_value(self, c):
        if 'dropdown' in self.CH.cell_options[c]:
            return self.MT._headers[c]

    def get_header_dropdown_values(self, c = 0):
        return self.CH.cell_options[c]['dropdown']['values']

    def header_dropdown_functions(self, c, selection_function = "", modified_function = ""):
        if selection_function != "":
            self.CH.cell_options[c]['dropdown']['select_function'] = selection_function
        if modified_function != "":
            self.CH.cell_options[c]['dropdown']['modified_function'] = modified_function
        return self.CH.cell_options[c]['dropdown']['select_function'], self.CH.cell_options[c]['dropdown']['modified_function']

    def set_header_dropdown_values(self, c = 0, set_existing_dropdown = False, values = [], displayed = None):
        if set_existing_dropdown:
            if self.CH.existing_dropdown_window is not None:
                c_ = self.CH.existing_dropdown_window.c
            else:
                raise Exception("No dropdown box is currently open")
        else:
            c_ = c
        self.CH.cell_options[c_]['dropdown']['values'] = values
        if self.CH.cell_options[c_]['dropdown']['window'] != "no dropdown open":
            self.CH.cell_options[c_]['dropdown']['window'].values(values)
        if displayed is not None:
            self.MT.headers(newheaders = displayed, index = c_)

    def delete_header_dropdown(self, c = 0):
        if c == "all":
            for c_ in self.CH.cell_options:
                if 'dropdown' in self.CH.cell_options[c_]:
                    self.CH.delete_dropdown(c_)
        else:
            self.CH.delete_dropdown(c)

    def get_header_dropdowns(self):
        return {k: v['dropdown'] for k, v in self.CH.cell_options.items() if 'dropdown' in v}
    
    def open_header_dropdown(self, c):
        self.CH.display_dropdown_window(c)

    def close_header_dropdown(self, c):
        self.CH.hide_dropdown_window(c)
        
    def create_index_dropdown(self,
                              r = 0,
                              values = [],
                              set_value = None,
                              state = "readonly",
                              redraw = False,
                              selection_function = None,
                              modified_function = None):
        if isinstance(r, str) and r.lower() == "all":
            for r_ in range(self.MT.total_data_rows()):
                self.RI.create_dropdown(r = r_,
                                        values = values,
                                        set_value = set_value,
                                        state = state,
                                        redraw = redraw,
                                        selection_function = selection_function,
                                        modified_function = modified_function)
        else:
            self.RI.create_dropdown(r = r,
                                    values = values,
                                    set_value = set_value,
                                    state = state,
                                    redraw = redraw,
                                    selection_function = selection_function,
                                    modified_function = modified_function)

    def get_index_dropdown_value(self, r):
        if 'dropdown' in self.RI.cell_options[r]:
            return self.MT._row_index[r]

    def get_index_dropdown_values(self, r = 0):
        return self.RI.cell_options[r]['dropdown']['values']

    def index_dropdown_functions(self, r, selection_function = "", modified_function = ""):
        if selection_function != "":
            self.RI.cell_options[r]['dropdown']['select_function'] = selection_function
        if modified_function != "":
            self.RI.cell_options[r]['dropdown']['modified_function'] = modified_function
        return self.RI.cell_options[r]['dropdown']['select_function'], self.RI.cell_options[r]['dropdown']['modified_function']

    def set_index_dropdown_values(self, r = 0, set_existing_dropdown = False, values = [], displayed = None):
        if set_existing_dropdown:
            if self.RI.existing_dropdown_window is not None:
                r_ = self.RI.existing_dropdown_window.r
            else:
                raise Exception("No dropdown box is currently open")
        else:
            r_ = r
        self.RI.cell_options[r_]['dropdown']['values'] = values
        if self.RI.cell_options[r_]['dropdown']['window'] != "no dropdown open":
            self.RI.cell_options[r_]['dropdown']['window'].values(values)
        if displayed is not None:
            self.MT.row_index(newindex = displayed, index = r_)

    def delete_index_dropdown(self, r = 0):
        if r == "all":
            for r_ in self.RI.cell_options:
                if 'dropdown' in self.RI.cell_options[r_]:
                    self.RI.delete_dropdown(r_)
        else:
            self.RI.delete_dropdown(r)

    def get_index_dropdowns(self):
        return {k: v['dropdown'] for k, v in self.RI.cell_options.items() if 'dropdown' in v}
    
    def open_index_dropdown(self, r):
        self.RI.display_dropdown_window(r)

    def close_index_dropdown(self, r):
        self.RI.hide_dropdown_window(r)

    def create_dropdown(self,
                        r = 0,
                        c = 0,
                        values = [],
                        set_value = None,
                        state = "readonly",
                        redraw = False,
                        selection_function = None,
                        modified_function = None):
        if isinstance(r, str) and r.lower() == "all" and isinstance(c, int):
            for r_ in range(self.MT.total_data_rows()):
                self.MT.create_dropdown(r = r_,
                                        c = c,
                                        values = values,
                                        set_value = set_value,
                                        state = state,
                                        redraw = redraw,
                                        selection_function = selection_function,
                                        modified_function = modified_function)
        elif isinstance(c, str) and c.lower() == "all" and isinstance(r, int):
            for c_ in range(self.MT.total_data_cols()):
                self.MT.create_dropdown(r = r,
                                        c = c_,
                                        values = values,
                                        set_value = set_value,
                                        state = state,
                                        redraw = redraw,
                                        selection_function = selection_function,
                                        modified_function = modified_function)
        elif isinstance(r, str) and r.lower() == "all" and isinstance(c, str) and c.lower() == "all":
            totalcols = self.MT.total_data_cols()
            for r_ in range(self.MT.total_data_rows()):
                for c_ in range(totalcols):
                    self.MT.create_dropdown(r = r_,
                                            c = c_,
                                            values = values,
                                            set_value = set_value,
                                            state = state,
                                            redraw = redraw,
                                            selection_function = selection_function,
                                            modified_function = modified_function)
        
        else:
            self.MT.create_dropdown(r = r,
                                    c = c,
                                    values = values,
                                    set_value = set_value,
                                    state = state,
                                    redraw = redraw,
                                    selection_function = selection_function,
                                    modified_function = modified_function)

    def get_dropdown_value(self, r, c):
        return self.get_cell_data(r, c)

    def get_dropdown_values(self, r = 0, c = 0):
        return self.MT.cell_options[(r, c)]['dropdown']['values']

    def dropdown_functions(self, r, c, selection_function = "", modified_function = ""):
        if selection_function != "":
            self.MT.cell_options[(r, c)]['dropdown']['select_function'] = selection_function
        if modified_function != "":
            self.MT.cell_options[(r, c)]['dropdown']['modified_function'] = modified_function
        return self.MT.cell_options[(r, c)]['dropdown']['select_function'], self.MT.cell_options[(r, c)]['dropdown']['modified_function']

    def set_dropdown_values(self, r = 0, c = 0, set_existing_dropdown = False, values = [], displayed = None):
        if set_existing_dropdown:
            if self.MT.existing_dropdown_window is not None:
                r_ = self.MT.existing_dropdown_window.r
                c_ = self.MT.existing_dropdown_window.c
            else:
                raise Exception("No dropdown box is currently open")
        else:
            r_ = r
            c_ = c
        self.MT.cell_options[(r_, c_)]['dropdown']['values'] = values
        if self.MT.cell_options[(r_, c_)]['dropdown']['window'] != "no dropdown open":
            self.MT.cell_options[(r_, c_)]['dropdown']['window'].values(values)
        if displayed is not None:
            self.set_cell_data(r_, c_, displayed)
            if self.MT.cell_options[(r_, c_)]['dropdown']['window'] != "no dropdown open" and self.MT.text_editor_loc is not None and self.MT.text_editor is not None:
                self.MT.text_editor.set_text(displayed)

    def delete_dropdown(self, r = 0, c = 0):
        if isinstance(r, str) and r.lower() == "all" and isinstance(c, int):
            for r_, c_ in self.MT.cell_options:
                if 'dropdown' in self.MT.cell_options[(r_, c)]:
                    self.MT.delete_dropdown(r_, c)
        elif isinstance(c, str) and c.lower() == "all" and isinstance(r, int):
            for r_, c_ in self.MT.cell_options:
                if 'dropdown' in self.MT.cell_options[(r, c_)]:
                    self.MT.delete_dropdown(r, c_)
        elif isinstance(r, str) and r.lower() == "all" and isinstance(c, str) and c.lower() == "all":
            for r_, c_ in self.MT.cell_options:
                if 'dropdown' in self.MT.cell_options[(r_, c_)]:
                    self.MT.delete_dropdown(r_, c_)
        else:
            self.MT.delete_dropdown(r, c)

    def get_dropdowns(self):
        return {k: v['dropdown'] for k, v in self.MT.cell_options.items() if 'dropdown' in v}

    def open_dropdown(self, r, c):
        self.MT.display_dropdown_window(r, c)

    def close_dropdown(self, r, c):
        self.MT.hide_dropdown_window(r, c)


class Sheet_Dropdown(Sheet):
    def __init__(self,
                 parent,
                 r,
                 c,
                 width = None,
                 height = None,
                 font = None,
                 colors = {'bg': theme_light_blue['popup_menu_bg'],
                           'fg': theme_light_blue['popup_menu_fg'],
                           'highlight_bg': theme_light_blue['popup_menu_highlight_bg'],
                           'highlight_fg': theme_light_blue['popup_menu_highlight_fg']},
                 outline_color = theme_light_blue['table_fg'],
                 outline_thickness = 2,
                 values = [],
                 hide_dropdown_window = None,
                 arrowkey_RIGHT = None,
                 arrowkey_LEFT = None,
                 align = "w",
                 # False for using r, c "r" for r "c" for c
                 single_index = False):
        Sheet.__init__(self,
                       parent = parent,
                       outline_thickness = outline_thickness,
                       outline_color = outline_color,
                       table_grid_fg = colors['fg'],
                       show_horizontal_grid = True,
                       show_vertical_grid = False,
                       show_header = False,
                       show_row_index = False,
                       show_top_left = False,
                       align = "w", #alignments other than w for dropdown boxes are broken at the moment
                       empty_horizontal = 0,
                       empty_vertical = 0,
                       selected_rows_to_end_of_window = True,
                       horizontal_grid_to_end_of_window = True,
                       show_selected_cells_border = False,
                       table_selected_cells_border_fg = colors['fg'],
                       table_selected_cells_bg = colors['highlight_bg'],
                       table_selected_rows_border_fg = colors['fg'],
                       table_selected_rows_bg = colors['highlight_bg'],
                       table_selected_rows_fg = colors['highlight_fg'],
                       width = width,
                       height = height,
                       font = font if font else get_font(),
                       table_fg = colors['fg'],
                       table_bg = colors['bg'])
        self.parent = parent
        self.hide_dropdown_window = hide_dropdown_window
        self.arrowkey_RIGHT = arrowkey_RIGHT
        self.arrowkey_LEFT = arrowkey_LEFT
        self.h_ = height
        self.w_ = width
        self.r = r
        self.c = c
        self.row = -1
        self.single_index = single_index
        self.bind("<Motion>", self.mouse_motion)
        self.bind("<ButtonPress-1>", self.b1)
        self.bind("<Up>", self.arrowkey_UP)
        self.bind("<Tab>", self.arrowkey_RIGHT)
        self.bind("<Right>", self.arrowkey_RIGHT)
        self.bind("<Down>", self.arrowkey_DOWN)
        self.bind("<Left>", self.arrowkey_LEFT)
        self.bind("<Prior>", self.arrowkey_UP)
        self.bind("<Next>", self.arrowkey_DOWN)
        self.bind("<Return>", self.b1)
        if values:
            self.values(values, redraw = False)

    def arrowkey_UP(self, event = None):
        self.deselect("all")
        if self.row > 0:
            self.row -= 1
        else:
            self.row = 0
        self.see(self.row, 0, redraw = False)
        self.select_row(self.row)

    def arrowkey_DOWN(self, event = None):
        self.deselect("all")
        if len(self.MT.data) - 1 > self.row:
            self.row += 1
        self.see(self.row, 0, redraw = False)
        self.select_row(self.row)

    def mouse_motion(self, event = None):
        self.row = self.identify_row(event, exclude_index = True, allow_end = False)
        self.deselect("all")
        if self.row is not None:
            self.select_row(self.row)
            
    def _reselect(self):
        rows = self.get_selected_rows()
        if rows:
            self.select_row(next(iter(rows)))

    def b1(self, event = None):
        if event is None:
            row = None
        elif event.keycode == 13:
            row = self.get_selected_rows()
            if not row:
                row = None
            else:
                row = next(iter(row))
        else:
            row = self.identify_row(event, exclude_index = True, allow_end = False)
        if self.single_index:
            if row is None:
                self.hide_dropdown_window(self.r if self.single_index == "r" else self.c)
            else:
                self.hide_dropdown_window(self.r if self.single_index == "r" else self.c, self.get_cell_data(row, 0))
        else:
            if row is None:
                self.hide_dropdown_window(self.r, self.c)
            else:
                self.hide_dropdown_window(self.r, self.c, self.get_cell_data(row, 0))

    def values(self, values = [], redraw = True):
        self.set_sheet_data([[v] for v in values],
                            reset_col_positions = False,
                            reset_row_positions = False,
                            redraw = False,
                            verify = False)
        self.set_all_cell_sizes_to_text(redraw = True)



