from ._tksheet_vars import *
from ._tksheet_other_classes import *

from collections import defaultdict, deque
from itertools import islice, repeat, accumulate, chain, product, cycle
from math import floor, ceil
from tkinter import TclError
import bisect
import csv as csv_module
import io
import pickle
import tkinter as tk
import zlib


class MainTable(tk.Canvas):
    def __init__(self,
                 *args,
                 **kwargs):
        tk.Canvas.__init__(self,
                           kwargs['parentframe'],
                           background = kwargs['table_bg'],
                           highlightthickness = 0)
        self.parentframe = kwargs['parentframe']
        self.b1_pressed_loc = None
        self.existing_dropdown_canvas_id = None
        self.existing_dropdown_window = None
        self.closed_dropdown = None
        self.last_selected = tuple()
        self.centre_alignment_text_mod_indexes = (slice(1, None), slice(None, -1))
        self.c_align_cyc = cycle(self.centre_alignment_text_mod_indexes)
        self.grid_cyctup = ("st", "end")
        self.grid_cyc = cycle(self.grid_cyctup)

        self.disp_ctrl_outline = {}
        self.disp_text = defaultdict(set)
        self.disp_high = defaultdict(set)
        self.disp_grid = {}
        self.disp_fill_sels = {}
        self.disp_bord_sels = {}
        self.disp_resize_lines = {}
        self.disp_dropdown = {}
        self.disp_checkbox = {}
        self.hidd_ctrl_outline = {}
        self.hidd_text = defaultdict(set)
        self.hidd_high = defaultdict(set)
        self.hidd_grid = {}
        self.hidd_fill_sels = {}
        self.hidd_bord_sels = {}
        self.hidd_resize_lines = {}
        self.hidd_dropdown = {}
        self.hidd_checkbox = {}

        self.cell_options = {}
        self.col_options = {}
        self.row_options = {}

        """
        cell options dict looks like:
        {(row int, column int): {'dropdown': {'values': values,
                                              'window': "no dropdown open",
                                              'select_function': selection_function,
                                              'keypress_function': keypress_function,
                                              'state': state},
                                 'highlight: (bg, fg),
                                 'align': "e",
                                 'readonly': True}
        """

        self.extra_table_rc_menu_funcs = {}
        self.extra_index_rc_menu_funcs = {}
        self.extra_header_rc_menu_funcs = {}
        self.extra_empty_space_rc_menu_funcs = {}

        self.max_undos = kwargs['max_undos']
        self.undo_storage = deque(maxlen = kwargs['max_undos'])

        self.page_up_down_select_row = kwargs['page_up_down_select_row']
        self.expand_sheet_if_paste_too_big = kwargs['expand_sheet_if_paste_too_big']
        self.paste_insert_column_limit = kwargs['paste_insert_column_limit']
        self.paste_insert_row_limit = kwargs['paste_insert_row_limit']
        self.arrow_key_down_right_scroll_page = kwargs['arrow_key_down_right_scroll_page']
        self.cell_auto_resize_enabled = kwargs['enable_edit_cell_auto_resize']
        self.edit_cell_validation = kwargs['edit_cell_validation']
        self.display_selected_fg_over_highlights = kwargs['display_selected_fg_over_highlights']
        self.show_index = kwargs['show_index']
        self.show_header = kwargs['show_header']
        self.selected_rows_to_end_of_window = kwargs['selected_rows_to_end_of_window']
        self.horizontal_grid_to_end_of_window = kwargs['horizontal_grid_to_end_of_window']
        self.vertical_grid_to_end_of_window = kwargs['vertical_grid_to_end_of_window']
        self.empty_horizontal = kwargs['empty_horizontal']
        self.empty_vertical = kwargs['empty_vertical']
        self.show_vertical_grid = kwargs['show_vertical_grid']
        self.show_horizontal_grid = kwargs['show_horizontal_grid']
        self.min_rh = 0
        self.hdr_min_rh = 0
        self.being_drawn_rect = None
        self.extra_motion_func = None
        self.extra_b1_press_func = None
        self.extra_b1_motion_func = None
        self.extra_b1_release_func = None
        self.extra_double_b1_func = None
        self.extra_rc_func = None
        
        self.extra_begin_ctrl_c_func = None
        self.extra_end_ctrl_c_func = None
        
        self.extra_begin_ctrl_x_func = None
        self.extra_end_ctrl_x_func = None

        self.extra_begin_ctrl_v_func = None
        self.extra_end_ctrl_v_func = None

        self.extra_begin_ctrl_z_func = None
        self.extra_end_ctrl_z_func = None

        self.extra_begin_delete_key_func = None
        self.extra_end_delete_key_func = None
        
        self.extra_begin_edit_cell_func = None
        self.extra_end_edit_cell_func = None

        self.extra_begin_del_rows_rc_func = None
        self.extra_end_del_rows_rc_func = None

        self.extra_begin_del_cols_rc_func = None
        self.extra_end_del_cols_rc_func = None

        self.extra_begin_insert_cols_rc_func = None
        self.extra_end_insert_cols_rc_func = None

        self.extra_begin_insert_rows_rc_func = None
        self.extra_end_insert_rows_rc_func = None

        self.text_editor_user_bound_keys = {}

        self.selection_binding_func = None
        self.deselection_binding_func = None
        self.drag_selection_binding_func = None
        self.shift_selection_binding_func = None
        self.select_all_binding_func = None

        self.single_selection_enabled = False
        self.toggle_selection_enabled = False # with this mode every left click adds the cell to selected cells
        self.ctrl_keys_over_dropdowns_enabled = kwargs['ctrl_keys_over_dropdowns_enabled']
        self.show_dropdown_borders = kwargs['show_dropdown_borders']
        self.drag_selection_enabled = False
        self.select_all_enabled = False
        self.arrowkeys_enabled = False
        self.undo_enabled = False
        self.cut_enabled = False
        self.copy_enabled = False
        self.paste_enabled = False
        self.delete_key_enabled = False
        self.rc_select_enabled = False
        self.rc_delete_column_enabled = False
        self.rc_insert_column_enabled = False
        self.rc_delete_row_enabled = False
        self.rc_insert_row_enabled = False
        self.rc_popup_menus_enabled = False
        self.edit_cell_enabled = False
        self.text_editor_loc = None
        self.show_selected_cells_border = kwargs['show_selected_cells_border']
        self.new_row_width = 0
        self.new_header_height = 0
        self.row_width_resize_bbox = tuple()
        self.header_height_resize_bbox = tuple()
        self.CH = kwargs['column_headers_canvas']
        self.CH.MT = self
        self.CH.RI = kwargs['row_index_canvas']
        self.RI = kwargs['row_index_canvas']
        self.RI.MT = self
        self.RI.CH = kwargs['column_headers_canvas']
        self.TL = None                # is set from within TopLeftRectangle() __init__
        self.all_columns_displayed = True
        self.all_rows_displayed = True
        self.align = kwargs['align']
        self._font = kwargs['font']
        self.fnt_fam = kwargs['font'][0]
        self.fnt_sze = kwargs['font'][1]
        self.fnt_wgt = kwargs['font'][2]
        self._index_font = kwargs['index_font']
        self.index_fnt_fam = kwargs['index_font'][0]
        self.index_fnt_sze = kwargs['index_font'][1]
        self.index_fnt_wgt = kwargs['index_font'][2]
        self._hdr_font = kwargs['header_font']
        self.hdr_fnt_fam = kwargs['header_font'][0]
        self.hdr_fnt_sze = kwargs['header_font'][1]
        self.hdr_fnt_wgt = kwargs['header_font'][2]
        self.txt_measure_canvas = tk.Canvas(self)
        self.txt_measure_canvas_text = self.txt_measure_canvas.create_text(0, 0, text = "", font = self._font)
        self.text_editor = None
        self.text_editor_id = None
        self.default_cw = kwargs['column_width']
        self.default_rh = (kwargs['row_height'] if isinstance(kwargs['row_height'], str) else "pixels",
                           kwargs['row_height'] if isinstance(kwargs['row_height'], int) else self.GetLinesHeight(int(kwargs['row_height'])))
        self.default_hh = (kwargs['header_height'] if isinstance(kwargs['header_height'], str) else "pixels",
                           kwargs['header_height'] if isinstance(kwargs['header_height'], int) else self.GetHdrLinesHeight(int(kwargs['header_height'])))
        self.set_fnt_help()
        self.set_hdr_fnt_help()
        self.set_index_fnt_help()
        self.data = kwargs['data_reference']
        if isinstance(self.data, (list, tuple)):
            self.data = kwargs['data_reference']
        else:
            self.data = []
        if not self.data:
            if isinstance(kwargs['total_rows'], int) and isinstance(kwargs['total_cols'], int) and kwargs['total_rows'] > 0 and kwargs['total_cols'] > 0:
                self.data = [list(repeat("", kwargs['total_cols'])) for i in range(kwargs['total_rows'])]
        _header = kwargs['header'] if kwargs['header'] is not None else kwargs['headers']
        if isinstance(_header, int):
            self._headers = _header
        else:
            if _header:
                self._headers = _header
            else:
                self._headers = []
        _row_index = kwargs['index'] if kwargs['index'] is not None else kwargs['row_index']
        if isinstance(_row_index, int):
            self._row_index = _row_index
        else:
            if _row_index:
                self._row_index = _row_index
            else:
                self._row_index = []
        self.displayed_columns = []
        self.displayed_rows = []
        self.col_positions = [0]
        self.row_positions = [0]
        self.reset_row_positions()
        self.display_columns(columns = kwargs['displayed_columns'],
                             all_columns_displayed = kwargs['all_columns_displayed'],
                             reset_col_positions = False,
                             deselect_all = False)
        self.reset_col_positions()
        self.table_grid_fg = kwargs['table_grid_fg']
        self.table_fg = kwargs['table_fg']
        self.table_selected_cells_border_fg = kwargs['table_selected_cells_border_fg']
        self.table_selected_cells_bg = kwargs['table_selected_cells_bg']
        self.table_selected_cells_fg = kwargs['table_selected_cells_fg']
        self.table_selected_rows_border_fg = kwargs['table_selected_rows_border_fg']
        self.table_selected_rows_bg = kwargs['table_selected_rows_bg']
        self.table_selected_rows_fg = kwargs['table_selected_rows_fg']
        self.table_selected_columns_border_fg = kwargs['table_selected_columns_border_fg']
        self.table_selected_columns_bg = kwargs['table_selected_columns_bg']
        self.table_selected_columns_fg = kwargs['table_selected_columns_fg']
        self.table_bg = kwargs['table_bg']
        self.popup_menu_font = kwargs['popup_menu_font']
        self.popup_menu_fg = kwargs['popup_menu_fg']
        self.popup_menu_bg = kwargs['popup_menu_bg']
        self.popup_menu_highlight_bg = kwargs['popup_menu_highlight_bg']
        self.popup_menu_highlight_fg = kwargs['popup_menu_highlight_fg']
        self.rc_popup_menu = None
        self.empty_rc_popup_menu = None
        self.basic_bindings()
        self.create_rc_menus()
        
    def refresh(self, event = None):
        self.main_table_redraw_grid_and_text(True, True)

    def basic_bindings(self, enable = True):
        if enable:
            self.bind("<Configure>", self.refresh)
            self.bind("<Motion>", self.mouse_motion)
            self.bind("<ButtonPress-1>", self.b1_press)
            self.bind("<B1-Motion>", self.b1_motion)
            self.bind("<ButtonRelease-1>", self.b1_release)
            self.bind("<Double-Button-1>", self.double_b1)
            self.bind("<MouseWheel>", self.mousewheel)
            if USER_OS == "linux":
                for canvas in (self, self.RI):
                    canvas.bind("<Button-4>", self.mousewheel)
                    canvas.bind("<Button-5>", self.mousewheel)
                for canvas in (self, self.CH):
                    canvas.bind("<Shift-Button-4>", self.shift_mousewheel)
                    canvas.bind("<Shift-Button-5>", self.shift_mousewheel)
            self.bind("<Shift-MouseWheel>", self.shift_mousewheel)
            self.bind("<Shift-ButtonPress-1>", self.shift_b1_press)
            self.CH.bind("<Shift-ButtonPress-1>", self.CH.shift_b1_press)
            self.RI.bind("<Shift-ButtonPress-1>", self.RI.shift_b1_press)
            self.CH.bind("<Shift-MouseWheel>", self.shift_mousewheel)
            self.RI.bind("<MouseWheel>", self.mousewheel)
            self.bind(get_rc_binding(), self.rc)
        else:
            self.unbind("<Configure>")
            self.unbind("<Motion>")
            self.unbind("<ButtonPress-1>")
            self.unbind("<B1-Motion>")
            self.unbind("<ButtonRelease-1>")
            self.unbind("<Double-Button-1>")
            self.unbind("<MouseWheel>")
            if USER_OS == "linux":
                for canvas in (self, self.RI):
                    canvas.unbind("<Button-4>")
                    canvas.unbind("<Button-5>")
                for canvas in (self, self.CH):
                    canvas.unbind("<Shift-Button-4>")
                    canvas.unbind("<Shift-Button-5>")
            self.unbind("<Shift-ButtonPress-1>")
            self.CH.unbind("<Shift-ButtonPress-1>")
            self.RI.unbind("<Shift-ButtonPress-1>")
            self.unbind("<Shift-MouseWheel>")
            self.CH.unbind("<Shift-MouseWheel>")
            self.RI.unbind("<MouseWheel>")
            self.unbind(get_rc_binding())

    def show_ctrl_outline(self, canvas = "table", start_cell = (0, 0), end_cell = (0, 0)):
        self.create_ctrl_outline(self.col_positions[start_cell[0]] + 2,
                                 self.row_positions[start_cell[1]] + 2,
                                 self.col_positions[end_cell[0]] - 2,
                                 self.row_positions[end_cell[1]] - 2,
                                 fill = "",
                                 dash = (10, 15),
                                 width = 3 if end_cell[0] - start_cell[0] == 1 and end_cell[1] - start_cell[1] == 1 else 2,
                                 outline = self.table_selected_cells_border_fg,
                                 tag = "ctrl")
        self.after(1500, self.delete_ctrl_outlines)

    def create_ctrl_outline(self, x1, y1, x2, y2, fill, dash, width, outline, tag):
        if self.hidd_ctrl_outline:
            t, sh = self.hidd_ctrl_outline.popitem()
            self.coords(t, x1, y1, x2, y2)
            if sh:
                self.itemconfig(t, fill = fill, dash = dash, width = width, outline = outline, tag = tag)
            else:
                self.itemconfig(t, fill = fill, dash = dash, width = width, outline = outline, tag = tag, state = "normal")
            self.lift(t)
        else:
            t = self.create_rectangle(x1, y1, x2, y2, fill = fill, dash = dash, width = width, outline = outline, tag = tag)
        self.disp_ctrl_outline[t] = True

    def delete_ctrl_outlines(self):
        self.hidd_ctrl_outline.update(self.disp_ctrl_outline)
        self.disp_ctrl_outline = {}
        for t, sh in self.hidd_ctrl_outline.items():
            if sh:
                self.itemconfig(t, state = "hidden")
                self.hidd_ctrl_outline[t] = False

    def get_ctrl_x_c_boxes(self):
        currently_selected = self.currently_selected()
        boxes = {}
        if currently_selected.type_ in ("cell", "column"):
            for item in chain(self.find_withtag("CellSelectFill"), self.find_withtag("Current_Outside"), self.find_withtag("ColSelectFill")):
                alltags = self.gettags(item)
                if alltags[0] == "CellSelectFill" or alltags[0] == "Current_Outside":
                    boxes[tuple(int(e) for e in alltags[1].split("_") if e)] = "cells"
                elif alltags[0] == "ColSelectFill":
                    boxes[tuple(int(e) for e in alltags[1].split("_") if e)] = "columns"
            maxrows = 0
            for r1, c1, r2, c2 in boxes:
                if r2 - r1 > maxrows:
                    maxrows = r2 - r1
            for r1, c1, r2, c2 in tuple(boxes):
                if r2 - r1 < maxrows:
                    del boxes[(r1, c1, r2, c2)]
            return boxes, maxrows
        else:
            for item in self.find_withtag("RowSelectFill"):
                boxes[tuple(int(e) for e in self.gettags(item)[1].split("_") if e)] = "rows"
            return boxes

    def ctrl_c(self, event = None):
        currently_selected = self.currently_selected()
        if currently_selected:
            s = io.StringIO()
            writer = csv_module.writer(s, dialect = csv_module.excel_tab, lineterminator = "\n")
            rows = []
            if currently_selected.type_ in ("cell", "column"):
                boxes, maxrows = self.get_ctrl_x_c_boxes()
                if self.extra_begin_ctrl_c_func is not None:
                    try:
                        self.extra_begin_ctrl_c_func(CtrlKeyEvent("begin_ctrl_c", boxes, currently_selected, tuple()))
                    except:
                        return
                for rn in range(maxrows):
                    row = []
                    for r1, c1, r2, c2 in boxes:
                        if r2 - r1 < maxrows:
                            continue
                        data_ref_rn = r1 + rn
                        for c in range(c1, c2):
                            dcol = c if self.all_columns_displayed else self.displayed_columns[c]
                            try:
                                row.append(self.data[data_ref_rn][dcol])
                            except:
                                row.append("")
                    writer.writerow(row)
                    rows.append(row)
            else:
                boxes = self.get_ctrl_x_c_boxes()
                if self.extra_begin_ctrl_c_func is not None:
                    try:
                        self.extra_begin_ctrl_c_func(CtrlKeyEvent("begin_ctrl_c", boxes, currently_selected, tuple()))
                    except:
                        return
                for r1, c1, r2, c2 in boxes:
                    for rn in range(r2 - r1):
                        row = []
                        data_ref_rn = r1 + rn
                        for c in range(c1, c2):
                            dcol = c if self.all_columns_displayed else self.displayed_columns[c]
                            try:
                                row.append(self.data[data_ref_rn][dcol])
                            except:
                                row.append("")
                        writer.writerow(row)
                        rows.append(row)
            for r1, c1, r2, c2 in boxes:
                self.show_ctrl_outline(canvas = "table", start_cell = (c1, r1), end_cell = (c2, r2))
            self.clipboard_clear()
            self.clipboard_append(s.getvalue())
            self.update_idletasks()
            if self.extra_end_ctrl_c_func is not None:
                self.extra_end_ctrl_c_func(CtrlKeyEvent("end_ctrl_c", boxes, currently_selected, rows))
            
    def ctrl_x(self, event = None):
        if self.anything_selected():
            undo_storage = {}
            s = io.StringIO()
            writer = csv_module.writer(s, dialect = csv_module.excel_tab, lineterminator = "\n")
            currently_selected = self.currently_selected()
            rows = []
            if currently_selected.type_ in ("cell", "column"):
                boxes, maxrows = self.get_ctrl_x_c_boxes()
                if self.extra_begin_ctrl_x_func is not None:
                    try:
                        self.extra_begin_ctrl_x_func(CtrlKeyEvent("begin_ctrl_x", boxes, currently_selected, tuple()))
                    except:
                        return
                for rn in range(maxrows):
                    row = []
                    for r1, c1, r2, c2 in boxes:
                        if r2 - r1 < maxrows:
                            continue
                        data_ref_rn = r1 + rn
                        for c in range(c1, c2):
                            dcol = c if self.all_columns_displayed else self.displayed_columns[c]
                            try:
                                sx = f"{self.data[data_ref_rn][dcol]}"
                                row.append(sx)
                                if self.undo_enabled:
                                    undo_storage[(data_ref_rn, dcol)] = sx
                            except:
                                row.append("")
                    writer.writerow(row)
                    rows.append(row)
                for rn in range(maxrows):
                    for r1, c1, r2, c2 in boxes:
                        if r2 - r1 < maxrows:
                            continue
                        data_ref_rn = r1 + rn
                        if data_ref_rn in self.row_options and 'readonly' in self.row_options[data_ref_rn]:
                            continue
                        for c in range(c1, c2):
                            dcol = c if self.all_columns_displayed else self.displayed_columns[c]
                            if (
                                ((data_ref_rn, dcol) in self.cell_options and ('readonly' in self.cell_options[(data_ref_rn, dcol)] or 'checkbox' in self.cell_options[(data_ref_rn, dcol)])) or
                                (dcol in self.col_options and 'readonly' in self.col_options[dcol]) or
                                (not self.ctrl_keys_over_dropdowns_enabled and 
                                (data_ref_rn, dcol) in self.cell_options and 
                                'dropdown' in self.cell_options[(data_ref_rn, dcol)] and
                                "" not in self.cell_options[(data_ref_rn, dcol)]['dropdown']['values'])
                                ):
                                continue
                            try:
                                self.data[data_ref_rn][dcol] = ""
                            except:
                                continue
            else:
                boxes = self.get_ctrl_x_c_boxes()
                if self.extra_begin_ctrl_x_func is not None:
                    try:
                        self.extra_begin_ctrl_x_func(CtrlKeyEvent("begin_ctrl_x", boxes, currently_selected, tuple()))
                    except:
                        return
                for r1, c1, r2, c2 in boxes:
                    for rn in range(r2 - r1):
                        row = []
                        data_ref_rn = r1 + rn
                        for c in range(c1, c2):
                            dcol = c if self.all_columns_displayed else self.displayed_columns[c]
                            try:
                                sx = f"{self.data[data_ref_rn][dcol]}"
                                row.append(sx)
                                if self.undo_enabled:
                                    undo_storage[(data_ref_rn, dcol)] = sx
                            except:
                                row.append("")
                        writer.writerow(row)
                        rows.append(row)
                for r1, c1, r2, c2 in boxes:
                    for rn in range(r2 - r1):
                        data_ref_rn = r1 + rn
                        if data_ref_rn in self.row_options and 'readonly' in self.row_options[data_ref_rn]:
                            continue
                        for c in range(c1, c2):
                            dcol = c if self.all_columns_displayed else self.displayed_columns[c]
                            if (
                                ((data_ref_rn, dcol) in self.cell_options and ('readonly' in self.cell_options[(data_ref_rn, dcol)] or 'checkbox' in self.cell_options[(data_ref_rn, dcol)])) or
                                (dcol in self.col_options and 'readonly' in self.col_options[dcol]) or
                                (not self.ctrl_keys_over_dropdowns_enabled and 
                                (data_ref_rn, dcol) in self.cell_options and 
                                'dropdown' in self.cell_options[(data_ref_rn, dcol)] and
                                "" not in self.cell_options[(data_ref_rn, dcol)]['dropdown']['values'])
                                ):
                                continue
                            try:
                                self.data[data_ref_rn][dcol] = ""
                            except:
                                continue
            if self.undo_enabled:
                self.undo_storage.append(zlib.compress(pickle.dumps(("edit_cells", undo_storage, tuple(boxes.items()), currently_selected))))
            self.clipboard_clear()
            self.clipboard_append(s.getvalue())
            self.update_idletasks()
            self.refresh()
            for r1, c1, r2, c2 in boxes:
                self.show_ctrl_outline(canvas = "table", start_cell = (c1, r1), end_cell = (c2, r2))
            if self.extra_end_ctrl_x_func is not None:
                self.extra_end_ctrl_x_func(CtrlKeyEvent("end_ctrl_x", boxes, currently_selected, rows))

    def find_last_selected_box_with_current(self, currently_selected):
        if currently_selected.type_ in ("cell", "column"):
            boxes, maxrows = self.get_ctrl_x_c_boxes()
        else:
            boxes = self.get_ctrl_x_c_boxes()
        for (r1, c1, r2, c2), type_ in boxes.items():
            if (type_[:2] == currently_selected.type_[:2] and
                currently_selected.row >= r1 and
                currently_selected.row <= r2 and
                currently_selected.column >= c1 and
                currently_selected.column <= c2):
                if (self.last_selected and self.last_selected == (r1, c1, r2, c2, type_)) or not self.last_selected:
                    return (r1, c1, r2, c2)
        return (currently_selected.row, currently_selected.column, currently_selected.row + 1, currently_selected.column + 1)

    def ctrl_v(self, event = None):
        if not self.expand_sheet_if_paste_too_big and (len(self.col_positions) == 1 or len(self.row_positions) == 1):
            return
        currently_selected = self.currently_selected()
        if currently_selected:
            y1 = currently_selected[0]
            x1 = currently_selected[1]
        elif not currently_selected and not self.expand_sheet_if_paste_too_big:
            return
        else:
            if not self.data:
                x1, y1 = 0, 0
            else:
                if len(self.col_positions) == 1 and len(self.row_positions) > 1:
                    x1, y1 = 0, len(self.row_positions) - 1
                elif len(self.row_positions) == 1 and len(self.col_positions) > 1:
                    x1, y1 = len(self.col_positions) - 1, 0
                elif len(self.row_positions) > 1 and len(self.col_positions) > 1:
                    x1, y1 = 0, len(self.row_positions) - 1
        try:
            data = self.clipboard_get()
        except:
            return
        data = list(csv_module.reader(io.StringIO(data), delimiter = "\t", quotechar = '"', skipinitialspace = True))
        if not data:
            return
        numcols = len(max(data, key = len))
        numrows = len(data)
        for rn, r in enumerate(data):
            if len(r) < numcols:
                data[rn].extend(list(repeat("", numcols - len(r))))
        lastbox_r1, lastbox_c1, lastbox_r2, lastbox_c2 = self.find_last_selected_box_with_current(currently_selected)
        lastbox_numrows = lastbox_r2 - lastbox_r1
        lastbox_numcols = lastbox_c2 - lastbox_c1
        
        if lastbox_numrows > numrows and lastbox_numrows % numrows == 0:
            nd = []
            for times in range(int(lastbox_numrows / numrows)):
                nd.extend([r.copy() for r in data])
            data.extend(nd)
            numrows *= int(lastbox_numrows / numrows)
        if lastbox_numcols > numcols and lastbox_numcols % numcols == 0:
            for rn, r in enumerate(data):
                for times in range(int(lastbox_numcols / numcols)):
                    data[rn].extend(r.copy())
            numcols *= int(lastbox_numcols / numcols)
        undo_storage = {}
        if self.expand_sheet_if_paste_too_big:
            added_rows = 0
            added_cols = 0
            if x1 + numcols > len(self.col_positions) - 1:
                added_cols = x1 + numcols - len(self.col_positions) + 1
                if isinstance(self.paste_insert_column_limit, int) and self.paste_insert_column_limit < len(self.col_positions) - 1 + added_cols:
                    added_cols = self.paste_insert_column_limit - len(self.col_positions) - 1
                if added_cols > 0:
                    self.insert_col_positions(widths = int(added_cols))
                if not self.all_columns_displayed:
                    total_data_cols = self.total_data_cols()
                    self.displayed_columns.extend(list(range(total_data_cols, total_data_cols + added_cols)))
            if y1 + numrows > len(self.row_positions) - 1:
                added_rows = y1 + numrows - len(self.row_positions) + 1
                if isinstance(self.paste_insert_row_limit, int) and self.paste_insert_row_limit < len(self.row_positions) - 1 + added_rows:
                    added_rows = self.paste_insert_row_limit - len(self.row_positions) - 1
                if added_rows > 0:
                    self.insert_row_positions(heights = int(added_rows))
            added_rows_cols = (added_rows, added_cols)
        else:
            added_rows_cols = (0, 0)
        if x1 + numcols > len(self.col_positions) - 1:
            numcols = len(self.col_positions) - 1 - x1
        if y1 + numrows > len(self.row_positions) - 1:
            numrows = len(self.row_positions) - 1 - y1
        if self.extra_begin_ctrl_v_func is not None or self.extra_end_ctrl_v_func is not None:
            rows = [[data[ndr][ndc] for ndc, c in enumerate(range(x1, x1 + numcols))] for ndr, r in enumerate(range(y1, y1 + numrows))]
        if self.extra_begin_ctrl_v_func is not None:
            try:
                self.extra_begin_ctrl_v_func(PasteEvent("begin_ctrl_v", currently_selected, rows))
            except:
                return
        for ndr, r in enumerate(range(y1, y1 + numrows)):
            for ndc, c in enumerate(range(x1, x1 + numcols)):
                dcol = c if self.all_columns_displayed else self.displayed_columns[c]
                if r > len(self.data) - 1:
                    self.data.extend([list(repeat("", c + 1)) for r in range((r + 1) - len(self.data))])
                elif c > len(self.data[r]) - 1:
                    self.data[r].extend(list(repeat("", (c + 1) - len(self.data[r]))))
                if (
                    ((r, dcol) in self.cell_options and 'readonly' in self.cell_options[(r, dcol)]) or
                    ((r, dcol) in self.cell_options and 'checkbox' in self.cell_options[(r, dcol)]) or
                    (dcol in self.col_options and 'readonly' in self.col_options[dcol]) or
                    (r in self.row_options and 'readonly' in self.row_options[r]) or
                    # if pasting not allowed in dropdowns and paste value isn't in dropdown values
                    (not self.ctrl_keys_over_dropdowns_enabled and 
                     (r, dcol) in self.cell_options and 
                     'dropdown' in self.cell_options[(r, dcol)] and
                     data[ndr][ndc] not in self.cell_options[(r, dcol)]['dropdown']['values'])
                    ):
                    continue
                if self.undo_enabled:
                    undo_storage[(r, dcol)] = f"{self.data[r][dcol]}"
                self.data[r][dcol] = data[ndr][ndc]
        if self.expand_sheet_if_paste_too_big and self.undo_enabled:
            self.equalize_data_row_lengths()
        self.deselect("all")
        if self.undo_enabled:
            self.undo_storage.append(zlib.compress(pickle.dumps(("edit_cells_paste",
                                                                 undo_storage,
                                                                 (((y1, x1, y1 + numrows, x1 + numcols), "cells"), ), # boxes
                                                                 currently_selected,
                                                                 added_rows_cols))))
        self.create_selected(y1, x1, y1 + numrows, x1 + numcols, "cells")
        self.create_current(y1, x1, type_ = "cell", inside = True if numrows > 1 or numcols > 1 else False)
        self.see(r = y1, c = x1, keep_yscroll = False, keep_xscroll = False, bottom_right_corner = False, check_cell_visibility = True, redraw = False)
        self.refresh()
        if self.extra_end_ctrl_v_func is not None:
            self.extra_end_ctrl_v_func(PasteEvent("end_ctrl_v", currently_selected, rows))

    def delete_key(self, event = None):
        if self.anything_selected():
            currently_selected = self.currently_selected()
            undo_storage = {}
            boxes = []
            for item in chain(self.find_withtag("CellSelectFill"), self.find_withtag("RowSelectFill"), self.find_withtag("ColSelectFill"), self.find_withtag("Current_Outside")):
                alltags = self.gettags(item)
                box = tuple(int(e) for e in alltags[1].split("_") if e)
                if alltags[0] in ("CellSelectFill", "Current_Outside"):
                    boxes.append((box, "cells"))
                elif alltags[0] == "ColSelectFill":
                    boxes.append((box, "columns"))
                elif alltags[0] == "RowSelectFill":
                    boxes.append((box, "rows"))
            if self.extra_begin_delete_key_func is not None:
                try:
                    self.extra_begin_delete_key_func(CtrlKeyEvent("begin_delete_key", boxes, currently_selected, tuple()))
                except:
                    return
            for (r1, c1, r2, c2), _ in boxes:
                for r in range(r1, r2):
                    for c in range(c1, c2):
                        dcol = c if self.all_columns_displayed else self.displayed_columns[c]
                        if (
                            ((r, dcol) in self.cell_options and ('readonly' in self.cell_options[(r, dcol)] or 'checkbox' in self.cell_options[(r, dcol)])) or
                            # if del key not allowed in dropdowns and empty string isn't in dropdown values
                            (not self.ctrl_keys_over_dropdowns_enabled and 
                            (r, dcol) in self.cell_options and 
                            'dropdown' in self.cell_options[(r, dcol)] and
                            "" not in self.cell_options[(r, dcol)]['dropdown']['values']) or
                            (dcol in self.col_options and 'readonly' in self.col_options[dcol]) or
                            (r in self.row_options and 'readonly' in self.row_options[r])
                            ):
                            continue
                        try:
                            if self.undo_enabled:
                                undo_storage[(r, dcol)] = f"{self.data[r][dcol]}"
                            self.data[r][dcol] = ""
                        except:
                            continue
            if self.extra_end_delete_key_func is not None:
                self.extra_end_delete_key_func(CtrlKeyEvent("end_delete_key", boxes, currently_selected, undo_storage))
            if self.undo_enabled:
                self.undo_storage.append(zlib.compress(pickle.dumps(("edit_cells", undo_storage, boxes, currently_selected))))
            self.refresh()
            
    def move_columns_adjust_options_dict(self, col, to_move_min, num_cols, move_data = True, create_selections = True):
        c = int(col)
        to_move_max = to_move_min + num_cols
        to_del = to_move_max + num_cols
        orig_selected = list(range(to_move_min, to_move_min + num_cols))
        self.deselect("all", redraw = False)
        cws = [int(b - a) for a, b in zip(self.col_positions, islice(self.col_positions, 1, len(self.col_positions)))]
        if to_move_min > c:
            cws[c:c] = cws[to_move_min:to_move_max]
            cws[to_move_max:to_del] = []
        else:
            cws[c + 1:c + 1] = cws[to_move_min:to_move_max]
            cws[to_move_min:to_move_max] = []
        self.col_positions = list(accumulate(chain([0], (width for width in cws))))
        if c + num_cols > len(self.col_positions):
            new_selected = tuple(range(len(self.col_positions) - 1 - num_cols, len(self.col_positions) - 1))
            if create_selections:
                self.create_selected(0, len(self.col_positions) - 1 - num_cols, len(self.row_positions) - 1, len(self.col_positions) - 1, "columns")
        else:
            if to_move_min > c:
                new_selected = tuple(range(c, c + num_cols))
                if create_selections:
                    self.create_selected(0, c, len(self.row_positions) - 1, c + num_cols, "columns")
            else:
                new_selected = tuple(range(c + 1 - num_cols, c + 1))
                if create_selections:
                    self.create_selected(0, c + 1 - num_cols, len(self.row_positions) - 1, c + 1, "columns")
        if create_selections:
            self.create_current(0, int(new_selected[0]), type_ = "column", inside = True)
        newcolsdct = {t1: t2 for t1, t2 in zip(orig_selected, new_selected)}
        if self.all_columns_displayed:
            dispset = {}
            if to_move_min > c:
                if move_data:
                    for rn in range(len(self.data)):
                        if len(self.data[rn]) < to_move_max:
                            self.data[rn].extend(list(repeat("", to_move_max - len(self.data[rn]) + 1)))
                        self.data[rn][c:c] = self.data[rn][to_move_min:to_move_max]
                        self.data[rn][to_move_max:to_del] = []
                    if isinstance(self._headers, list) and self._headers:
                        if len(self._headers) < to_move_max:
                            self._headers.extend(list(repeat("", to_move_max - len(self._headers) + 1)))
                        self._headers[c:c] = self._headers[to_move_min:to_move_max]
                        self._headers[to_move_max:to_del] = []
                self.CH.cell_options = {
                    newcolsdct[k] if k in newcolsdct else
                    k + num_cols if k < to_move_min and k >= c else
                    k: v for k, v in self.CH.cell_options.items()
                }
                self.cell_options = {
                    (k[0], newcolsdct[k[1]]) if k[1] in newcolsdct else
                    (k[0], k[1] + num_cols) if k[1] < to_move_min and k[1] >= c else
                    k: v for k, v in self.cell_options.items()
                }
                self.col_options = {
                    newcolsdct[k] if k in newcolsdct else
                    k + num_cols if k < to_move_min and k >= c else
                    k: v for k, v in self.col_options.items()
                }
            else:
                c += 1
                if move_data:
                    for rn in range(len(self.data)):
                        if len(self.data[rn]) < c - 1:
                            self.data[rn].extend(list(repeat("", c - len(self.data[rn]))))
                        self.data[rn][c:c] = self.data[rn][to_move_min:to_move_max]
                        self.data[rn][to_move_min:to_move_max] = []
                    if isinstance(self._headers, list) and self._headers:
                        if len(self._headers) < c:
                            self._headers.extend(list(repeat("", c - len(self._headers))))
                        self._headers[c:c] = self._headers[to_move_min:to_move_max]
                        self._headers[to_move_min:to_move_max] = []
                self.CH.cell_options = {
                    newcolsdct[k] if k in newcolsdct else
                    k - num_cols if k < c and k > to_move_min else
                    k: v for k, v in self.CH.cell_options.items()
                }
                self.cell_options = {
                    (k[0], newcolsdct[k[1]]) if k[1] in newcolsdct else 
                    (k[0], k[1] - num_cols) if k[1] < c and k[1] > to_move_min else 
                    k: v for k, v in self.cell_options.items()
                }
                self.col_options = {
                    newcolsdct[k] if k in newcolsdct else
                    k - num_cols if k < c and k > to_move_min else
                    k: v for k, v in self.col_options.items()
                }
        else:
            # moves data around, not displayed columns indexes
            # which remain sorted and the same after drop and drop
            if to_move_min > c:
                dispset = {a: b for a, b in zip(self.displayed_columns, (self.displayed_columns[:c] +
                                                                         self.displayed_columns[to_move_min:to_move_min + num_cols] +
                                                                         self.displayed_columns[c:to_move_min] +
                                                                         self.displayed_columns[to_move_min + num_cols:]))}
            else:
                dispset = {a: b for a, b in zip(self.displayed_columns, (self.displayed_columns[:to_move_min] +
                                                                         self.displayed_columns[to_move_min + num_cols:c + 1] +
                                                                         self.displayed_columns[to_move_min:to_move_min + num_cols] +
                                                                         self.displayed_columns[c + 1:]))}
            # has to pick up elements from all over the place in the original row
            # building an entirely new row is best due to permutations of hidden columns
            if move_data:
                max_idx = max(chain(dispset, dispset.values())) + 1
                for rn in range(len(self.data)):
                    if len(self.data[rn]) < max_idx:
                        self.data[rn][:] = self.data[rn] + list(repeat("", max_idx - len(self.data[rn])))
                    new = []
                    idx = 0
                    done = set()
                    while len(new) < len(self.data[rn]):
                        if idx in dispset and idx not in done:
                            new.append(self.data[rn][dispset[idx]])
                            done.add(idx)
                        elif idx not in done:
                            new.append(self.data[rn][idx])
                            idx += 1
                        else:
                            idx += 1
                    self.data[rn] = new
                if isinstance(self._headers, list) and self._headers:
                    if len(self._headers) < max_idx:
                        self._headers[:] = self._headers + list(repeat("", max_idx - len(self._headers)))
                    new = []
                    idx = 0
                    done = set()
                    while len(new) < len(self._headers):
                        if idx in dispset and idx not in done:
                            new.append(self._headers[dispset[idx]])
                            done.add(idx)
                        elif idx not in done:
                            new.append(self._headers[idx])
                            idx += 1
                        else:
                            idx += 1
                    self._headers = new
            dispset = {b: a for a, b in dispset.items()}
            self.CH.cell_options = {dispset[k] if k in dispset else k: v for k, v in self.CH.cell_options.items()}
            self.cell_options = {(k[0], dispset[k[1]]) if k[1] in dispset else k: v for k, v in self.cell_options.items()}
            self.col_options = {dispset[k] if k in dispset else k: v for k, v in self.col_options.items()}
        return new_selected, dispset
    
    def move_rows_adjust_options_dict(self, row, to_move_min, num_rows, move_data = True, create_selections = True):
        r = int(row)
        to_move_max = to_move_min + num_rows
        to_del = to_move_max + num_rows
        orig_selected = list(range(to_move_min, to_move_min + num_rows))
        self.deselect("all", redraw = False)
        rhs = [int(b - a) for a, b in zip(self.row_positions, islice(self.row_positions, 1, len(self.row_positions)))]
        if to_move_min > r:
            rhs[r:r] = rhs[to_move_min:to_move_max]
            rhs[to_move_max:to_del] = []
        else:
            rhs[r + 1:r + 1] = rhs[to_move_min:to_move_max]
            rhs[to_move_min:to_move_max] = []
        self.row_positions = list(accumulate(chain([0], (height for height in rhs))))       
        if r + num_rows > len(self.row_positions):
            new_selected = tuple(range(len(self.row_positions) - 1 - num_rows, len(self.row_positions) - 1))
            if create_selections:
                self.create_selected(len(self.row_positions) - 1 - num_rows, 0, len(self.row_positions) - 1, len(self.col_positions) - 1, "rows")
        else:
            if to_move_min > r:
                new_selected = tuple(range(r, r + num_rows))
                if create_selections:
                    self.create_selected(r, 0, r + num_rows, len(self.col_positions) - 1, "rows")
            else:
                new_selected = tuple(range(r + 1 - num_rows, r + 1))
                if create_selections:
                    self.create_selected(r + 1 - num_rows, 0, r + 1, len(self.col_positions) - 1, "rows")
        if create_selections:
            self.create_current(int(new_selected[0]), 0, type_ = "row", inside = True)
        newrowsdct = {t1: t2 for t1, t2 in zip(orig_selected, new_selected)}
        if to_move_min > r:
            if move_data:
                self.data[r:r] = self.data[to_move_min:to_move_max]
                self.data[to_move_max:to_del] = []
                if isinstance(self._row_index, list) and self._row_index:
                    if len(self._row_index) < to_move_max:
                        self._row_index.extend(list(repeat("", to_move_max - len(self._row_index) + 1)))
                    self._row_index[r:r] = self._row_index[to_move_min:to_move_max]
                    self._row_index[to_move_max:to_del] = []
            self.RI.cell_options = {
                newrowsdct[k] if k in newrowsdct else
                k + num_rows if k < to_move_min and k >= r else
                k: v for k, v in self.RI.cell_options.items()
            }
            self.cell_options = {
                (newrowsdct[k[0]], k[1]) if k[0] in newrowsdct else
                (k[0] + num_rows, k[1]) if k[0] < to_move_min and k[0] >= r else
                k: v for k, v in self.cell_options.items()
            }
            self.row_options = {
                newrowsdct[k] if k in newrowsdct else
                k + num_rows if k < to_move_min and k >= r else
                k: v for k, v in self.row_options.items()
            }
        else:
            r += 1
            if move_data:
                self.data[r:r] = self.data[to_move_min:to_move_max]
                self.data[to_move_min:to_move_max] = []
                if isinstance(self._row_index, list) and self._row_index:
                    if len(self._row_index) < r:
                        self._row_index.extend(list(repeat("", r - len(self._row_index))))
                    self._row_index[r:r] = self._row_index[to_move_min:to_move_max]
                    self._row_index[to_move_min:to_move_max] = []
            self.RI.cell_options = {
                newrowsdct[k] if k in newrowsdct else
                k - num_rows if k < r and k > to_move_min else
                k: v for k, v in self.RI.cell_options.items()
            }
            self.cell_options = {
                (newrowsdct[k[0]], k[1]) if k[0] in newrowsdct else 
                (k[0] - num_rows, k[1]) if k[0] < r and k[0] > to_move_min else 
                k: v for k, v in self.cell_options.items()
            }
            self.row_options = {
                newrowsdct[k] if k in newrowsdct else
                k - num_rows if k < r and k > to_move_min else
                k: v for k, v in self.row_options.items()
            }
        return new_selected, {}

    def ctrl_z(self, event = None):
        if self.undo_storage:
            if not isinstance(self.undo_storage[-1], (tuple, dict)):
                undo_storage = pickle.loads(zlib.decompress(self.undo_storage[-1]))
            else:
                undo_storage = self.undo_storage[-1]
            self.deselect("all")
            if self.extra_begin_ctrl_z_func is not None:
                try:
                    self.extra_begin_ctrl_z_func(UndoEvent("begin_ctrl_z", undo_storage[0], undo_storage))
                except:
                    return
            self.undo_storage.pop()
            if undo_storage[0] in ("edit_header", ):
                for c, v in undo_storage[1].items():
                    self._headers[c] = v
                for box in undo_storage[2]:
                    r1, c1, r2, c2 = box[0]
                    if not self.expand_sheet_if_paste_too_big:
                        self.create_selected(r1, c1, r2, c2, box[1])
                self.create_current(0, undo_storage[3][1], type_ = "column", inside = True)
                
            if undo_storage[0] in ("edit_index", ):
                for r, v in undo_storage[1].items():
                    self._row_index[r] = v
                for box in undo_storage[2]:
                    r1, c1, r2, c2 = box[0]
                    if not self.expand_sheet_if_paste_too_big:
                        self.create_selected(r1, c1, r2, c2, box[1])
                self.create_current(0, undo_storage[3][1], type_ = "row", inside = True)
                            
            if undo_storage[0] in ("edit_cells", "edit_cells_paste"):
                for (r, c), v in undo_storage[1].items():
                    self.data[r][c] = v
                start_row = float("inf")
                start_col = float("inf")
                for box in undo_storage[2]:
                    r1, c1, r2, c2 = box[0]
                    if not self.expand_sheet_if_paste_too_big:
                        self.create_selected(r1, c1, r2, c2, box[1])
                    if r1 < start_row:
                        start_row = r1
                    if c1 < start_col:
                        start_col = c1
                if undo_storage[0] == "edit_cells_paste" and self.expand_sheet_if_paste_too_big:
                    if undo_storage[4][0] > 0:
                        self.del_row_positions(len(self.row_positions) - 1 - undo_storage[4][0], undo_storage[4][0])
                        self.data[:] = self.data[:-undo_storage[4][0]]
                    if undo_storage[4][1] > 0:
                        quick_added_cols = undo_storage[4][1]
                        self.del_col_positions(len(self.col_positions) - 1 - quick_added_cols, quick_added_cols)
                        for rn in range(len(self.data)):
                            self.data[rn][:] = self.data[rn][:-quick_added_cols]
                        if not self.all_columns_displayed:
                            self.displayed_columns[:] = self.displayed_columns[:-quick_added_cols]
                if undo_storage[3]:
                    if isinstance(undo_storage[3][0], int):
                        self.create_current(undo_storage[3][0], undo_storage[3][1], type_ = "cell", inside = True if self.cell_selected(undo_storage[3][0], undo_storage[3][1]) else False)
                    elif undo_storage[3][0] == "column":
                        self.create_current(0, undo_storage[3][1], type_ = "column", inside = True)
                    elif undo_storage[3][0] == "row":
                        self.create_current(undo_storage[3][1], 0, type_ = "row", inside = True)
                elif start_row < len(self.row_positions) - 1 and start_col < len(self.col_positions) - 1:
                    self.create_current(start_row, start_col, type_ = "cell", inside = True if self.cell_selected(start_row, start_col) else False)
                if start_row < len(self.row_positions) - 1 and start_col < len(self.col_positions) - 1:
                    self.see(r = start_row, c = start_col, keep_yscroll = False, keep_xscroll = False, bottom_right_corner = False, check_cell_visibility = True, redraw = False)
            
            elif undo_storage[0] == "move_cols":
                c = undo_storage[1]
                to_move_min = undo_storage[2]
                totalcols = len(undo_storage[4])
                if to_move_min < c:
                    c += totalcols - 1
                self.move_columns_adjust_options_dict(c, to_move_min, totalcols)
            
            elif undo_storage[0] == "move_rows":
                rhs = [int(b - a) for a, b in zip(self.row_positions, islice(self.row_positions, 1, len(self.row_positions)))]
                r = undo_storage[1]
                to_move_min = undo_storage[2]
                totalrows = len(undo_storage[4])
                if to_move_min < r:
                    r += totalrows - 1
                self.move_rows_adjust_options_dict(r, to_move_min, totalrows)
                        
            elif undo_storage[0] == "insert_row":
                self.data[undo_storage[1]['data_row_num']:undo_storage[1]['data_row_num'] + undo_storage[1]['numrows']] = []
                try:
                    self._row_index[undo_storage[1]['data_row_num']:undo_storage[1]['data_row_num'] + undo_storage[1]['numrows']] = []
                except:
                    pass
                self.del_row_positions(undo_storage[1]['sheet_row_num'],
                                       undo_storage[1]['numrows'],
                                       deselect_all = False)
                for r in range(undo_storage[1]['sheet_row_num'],
                               undo_storage[1]['sheet_row_num'] + undo_storage[1]['numrows']):
                    if r in self.row_options:
                        del self.row_options[r]
                    if r in self.RI.cell_options:
                        del self.RI.cell_options[r]
                numrows = undo_storage[1]['numrows']
                idx = undo_storage[1]['sheet_row_num'] + undo_storage[1]['numrows']
                self.cell_options = {(rn if rn < idx else rn - numrows, cn): t2 for (rn, cn), t2 in self.cell_options.items()}
                self.row_options = {rn if rn < idx else rn - numrows: t for rn, t in self.row_options.items()}
                self.RI.cell_options = {rn if rn < idx else rn - numrows: t for rn, t in self.RI.cell_options.items()}
                if len(self.row_positions) > 1:
                    start_row = undo_storage[1]['sheet_row_num'] if undo_storage[1]['sheet_row_num'] < len(self.row_positions) - 1 else undo_storage[1]['sheet_row_num'] - 1
                    self.RI.select_row(start_row)
                    self.see(r = start_row, c = 0, keep_yscroll = False, keep_xscroll = False, bottom_right_corner = False, check_cell_visibility = True, redraw = False)
                    
            elif undo_storage[0] == "insert_col":
                self.displayed_columns = undo_storage[1]['displayed_columns']
                qx = undo_storage[1]['data_col_num']
                qnum = undo_storage[1]['numcols']
                for rn in range(len(self.data)):
                    self.data[rn][qx:qx + qnum] = []
                try:
                    self._headers[qx:qx + qnum] = []
                except:
                    pass
                self.del_col_positions(undo_storage[1]['sheet_col_num'],
                                       undo_storage[1]['numcols'],
                                       deselect_all = False)
                for c in range(undo_storage[1]['sheet_col_num'],
                               undo_storage[1]['sheet_col_num'] + undo_storage[1]['numcols']):
                    if c in self.col_options:
                        del self.col_options[c]
                    if c in self.CH.cell_options:
                        del self.CH.cell_options[c]
                numcols = undo_storage[1]['numcols']
                idx = undo_storage[1]['sheet_col_num'] + undo_storage[1]['numcols']
                self.cell_options = {(rn, cn if cn < idx else cn - numcols): t2 for (rn, cn), t2 in self.cell_options.items()}
                self.col_options = {cn if cn < idx else cn - numcols: t for cn, t in self.col_options.items()}
                self.CH.cell_options = {cn if cn < idx else cn - numcols: t for cn, t in self.CH.cell_options.items()}
                if len(self.col_positions) > 1:
                    start_col = undo_storage[1]['sheet_col_num'] if undo_storage[1]['sheet_col_num'] < len(self.col_positions) - 1 else undo_storage[1]['sheet_col_num'] - 1
                    self.CH.select_col(start_col)
                    self.see(r = 0, c = start_col, keep_yscroll = False, keep_xscroll = False, bottom_right_corner = False, check_cell_visibility = True, redraw = False)
                    
            elif undo_storage[0] == "delete_rows":
                for rn, r, h in reversed(undo_storage[1]['deleted_rows']):
                    self.data.insert(rn, r)
                    self.insert_row_position(idx = rn, height = h)
                self.cell_options = undo_storage[1]['cell_options']
                self.row_options = undo_storage[1]['row_options']
                self.RI.cell_options = undo_storage[1]['RI_cell_options']
                for rn, r in reversed(undo_storage[1]['deleted_index_values']):
                    try:
                        self._row_index.insert(rn, r)
                    except:
                        continue
                self.reselect_from_get_boxes(undo_storage[1]['selection_boxes'])
                
            elif undo_storage[0] == "delete_cols":
                self.displayed_columns = undo_storage[1]['displayed_columns']
                self.cell_options = undo_storage[1]['cell_options']
                self.col_options = undo_storage[1]['col_options']
                self.CH.cell_options = undo_storage[1]['CH_cell_options']
                for cn, w in reversed(tuple(undo_storage[1]['colwidths'].items())):
                    self.insert_col_position(idx = cn, width = w)
                for cn, rowdict in reversed(tuple(undo_storage[1]['deleted_cols'].items())):
                    for rn, v in rowdict.items():
                        try:
                            self.data[rn].insert(cn, v)
                        except:
                            continue
                for cn, v in reversed(tuple(undo_storage[1]['deleted_hdr_values'].items())):
                    try:
                        self._headers.insert(cn, v)
                    except:
                        continue
                self.reselect_from_get_boxes(undo_storage[1]['selection_boxes'])
            self.refresh()
            if self.extra_end_ctrl_z_func is not None:
                self.extra_end_ctrl_z_func(UndoEvent("end_ctrl_z", undo_storage[0], undo_storage))
            
    def bind_arrowkeys(self, event = None):
        self.arrowkeys_enabled = True
        for canvas in (self, self.parentframe, self.CH, self.RI, self.TL):
            canvas.bind("<Up>", self.arrowkey_UP)
            canvas.bind("<Tab>", self.tab_key)
            canvas.bind("<Right>", self.arrowkey_RIGHT)
            canvas.bind("<Down>", self.arrowkey_DOWN)
            canvas.bind("<Left>", self.arrowkey_LEFT)
            canvas.bind("<Prior>", self.page_UP)
            canvas.bind("<Next>", self.page_DOWN)

    def unbind_arrowkeys(self, event = None):
        self.arrowkeys_enabled = False
        for canvas in (self, self.parentframe, self.CH, self.RI, self.TL):
            canvas.unbind("<Up>")
            canvas.unbind("<Right>")
            canvas.unbind("<Tab>")
            canvas.unbind("<Down>")
            canvas.unbind("<Left>")
            canvas.unbind("<Prior>")
            canvas.unbind("<Next>")

    def see(self,
            r = None,
            c = None,
            keep_yscroll = False,
            keep_xscroll = False,
            bottom_right_corner = False,
            check_cell_visibility = True,
            redraw = True):
        need_redraw = False
        if check_cell_visibility:
            yvis, xvis = self.cell_completely_visible(r = r, c = c, separate_axes = True)
        else:
            yvis, xvis = False, False
        if not yvis:
            if bottom_right_corner:
                if r is not None and not keep_yscroll:
                    winfo_height = self.winfo_height()
                    if self.row_positions[r + 1] - self.row_positions[r] > winfo_height:
                        y = self.row_positions[r]
                    else:
                        y = self.row_positions[r + 1] + 1 - winfo_height
                    args = ("moveto", y / (self.row_positions[-1] + self.empty_vertical))
                    if args[1] > 1:
                        args[1] = args[1] - 1
                    self.yview(*args)
                    self.RI.yview(*args)
                    need_redraw = True
            else:
                if r is not None and not keep_yscroll:
                    args = ("moveto", self.row_positions[r] / (self.row_positions[-1] + self.empty_vertical))
                    if args[1] > 1:
                        args[1] = args[1] - 1
                    self.yview(*args)
                    self.RI.yview(*args)
                    need_redraw = True
        if not xvis:
            if bottom_right_corner:
                if c is not None and not keep_xscroll:
                    winfo_width = self.winfo_width()
                    if self.col_positions[c + 1] - self.col_positions[c] > winfo_width:
                        x = self.col_positions[c]
                    else:
                        x = self.col_positions[c + 1] + 1 - winfo_width
                    args = ("moveto", x / (self.col_positions[-1] + self.empty_horizontal))
                    self.xview(*args)
                    self.CH.xview(*args)
                    need_redraw = True
            else:
                if c is not None and not keep_xscroll:
                    args = ("moveto", self.col_positions[c] / (self.col_positions[-1] + self.empty_horizontal))
                    self.xview(*args)
                    self.CH.xview(*args)
                    need_redraw = True
        if redraw and need_redraw:
            self.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
            return True
        else:
            return False
        
    def get_cell_coords(self, r = None, c = None):
        return (0 if not c else self.col_positions[c] + 1,
                0 if not r else self.row_positions[r] + 1, 
                0 if not c else self.col_positions[c + 1],
                0 if not r else self.row_positions[r + 1])

    def cell_completely_visible(self, r = 0, c = 0, separate_axes = False):
        cx1, cy1, cx2, cy2 = self.get_canvas_visible_area()
        x1, y1, x2, y2 = self.get_cell_coords(r, c)
        x_vis = True
        y_vis = True
        if cx1 > x1 or cx2 < x2:
            x_vis = False
        if cy1 > y1  or cy2 < y2:
            y_vis = False
        if separate_axes:
            return y_vis, x_vis
        else:
            if not y_vis or not x_vis:
                return False
            else:
                return True

    def cell_visible(self, r = 0, c = 0):
        cx1, cy1, cx2, cy2 = self.get_canvas_visible_area()
        x1, y1, x2, y2 = self.get_cell_coords(r, c)
        if x1 <= cx2 or y1 <= cy2 or x2 >= cx1 or y2 >= cy1:
            return True
        return False

    def select_all(self, redraw = True, run_binding_func = True):
        self.deselect("all")
        if len(self.row_positions) > 1 and len(self.col_positions) > 1:
            self.create_current(0, 0, type_ = "cell", inside = True)
            self.create_selected(0, 0, len(self.row_positions) - 1, len(self.col_positions) - 1)
            if redraw:
                self.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
            if self.select_all_binding_func is not None and run_binding_func:
                self.select_all_binding_func(SelectionBoxEvent("select_all_cells", (0, 0, len(self.row_positions) - 1, len(self.col_positions) - 1)))

    def select_cell(self, r, c, redraw = False, keep_other_selections = False):
        r = int(r)
        c = int(c)
        ignore_keep = False
        if keep_other_selections:
            if self.cell_selected(r, c):
                self.create_current(r, c, type_ = "cell", inside = True)
            else:
                ignore_keep = True
        if ignore_keep or not keep_other_selections:
            self.delete_selection_rects()
            self.create_current(r, c, type_ = "cell", inside = False)
        if redraw:
            self.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
        if self.selection_binding_func is not None:
            self.selection_binding_func(SelectCellEvent("select_cell", r, c))

    def move_down(self):
        currently_selected = self.currently_selected()
        if currently_selected:
            r, c = currently_selected.row, currently_selected.column
            if (
                r < len(self.row_positions) - 2 and
                (self.single_selection_enabled or self.toggle_selection_enabled)
                ):
                self.select_cell(r + 1, c)
                self.see(r + 1, c, keep_xscroll = True, bottom_right_corner = True, check_cell_visibility = True)

    def add_selection(self, r, c, redraw = False, run_binding_func = True, set_as_current = False):
        r = int(r)
        c = int(c)
        if set_as_current:
            items = self.find_withtag("Current_Outside")
            if items:
                alltags = self.gettags(items[0])
                if alltags[2] == "cell":
                    r1, c1, r2, c2 = tuple(int(e) for e in alltags[1].split("_") if e)
                    add_sel = (r1, c1)
                else:
                    add_sel = tuple()
            else:
                add_sel = tuple()
            self.create_current(r, c, type_ = "cell", inside = True if self.cell_selected(r, c) else False)
            if add_sel:
                self.add_selection(add_sel[0], add_sel[1], redraw = False, run_binding_func = False, set_as_current = False)
        else:
            self.create_selected(r, c, r + 1, c + 1)
        if redraw:
            self.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
        if self.selection_binding_func is not None and run_binding_func:
            self.selection_binding_func(SelectCellEvent("select_cell", r, c))

    def toggle_select_cell(self, row, column, add_selection = True, redraw = True, run_binding_func = True, set_as_current = True):
        if add_selection:
            if self.cell_selected(row, column, inc_rows = True, inc_cols = True):
                self.deselect(r = row, c = column, redraw = redraw)
            else:
                self.add_selection(r = row, c = column, redraw = redraw, run_binding_func = run_binding_func, set_as_current = set_as_current)
        else:
            if self.cell_selected(row, column, inc_rows = True, inc_cols = True):
                self.deselect(r = row, c = column, redraw = redraw)
            else:
                self.select_cell(row, column, redraw = redraw)

    def align_rows(self, rows = [], align = "global", align_index = False): #"center", "w", "e" or "global"
        if isinstance(rows, str) and rows.lower() == "all" and align == "global":
            for r in self.row_options:
                if 'align' in self.row_options[r]:
                    del self.row_options[r]['align']
            if align_index:
                for r in self.RI.cell_options:
                    if r in self.RI.cell_options and 'align' in self.RI.cell_options[r]:
                        del self.RI.cell_options[r]['align']
            return
        if isinstance(rows, int):
            rows_ = [rows]
        elif isinstance(rows, str) and rows.lower() == "all":
            rows_ = (r for r in range(self.total_data_rows()))
        else:
            rows_ = rows
        if align == "global":
            for r in rows_:
                if r in self.row_options and 'align' in self.row_options[r]:
                    del self.row_options[r]['align']
                if align_index and r in self.RI.cell_options and 'align' in self.RI.cell_options[r]:
                    del self.RI.cell_options[r]['align']
        else:
            for r in rows_:
                if r not in self.row_options:
                    self.row_options[r] = {}
                self.row_options[r]['align'] = align
                if align_index:
                    if r not in self.RI.cell_options:
                        self.RI.cell_options[r] = {}
                    self.RI.cell_options[r]['align'] = align

    def align_columns(self, columns = [], align = "global", align_header = False): #"center", "w", "e" or "global"
        if isinstance(columns, str) and columns.lower() == "all" and align == "global":
            for c in self.col_options:
                if 'align' in self.col_options[c]:
                    del self.col_options[c]['align']
            if align_header:
                for c in self.CH.cell_options:
                    if c in self.CH.cell_options and 'align' in self.CH.cell_options[c]:
                        del self.CH.cell_options[c]['align']
            return
        if isinstance(columns, int):
            cols_ = [columns]
        elif isinstance(columns, str) and columns.lower() == "all":
            cols_ = (c for c in range(self.total_data_cols()))
        else:
            cols_ = columns
        if align == "global":
            for c in cols_:
                if c in self.col_options and 'align' in self.col_options[c]:
                    del self.col_options[c]['align']
                if align_header and c in self.CH.cell_options and 'align' in self.CH.cell_options[c]:
                    del self.CH.cell_options[c]['align']
        else:
            for c in cols_:
                if c not in self.col_options:
                    self.col_options[c] = {}
                self.col_options[c]['align'] = align
                if align_header:
                    if c not in self.CH.cell_options:
                        self.CH.cell_options[c] = {}
                    self.CH.cell_options[c]['align'] = align

    def align_cells(self, row = 0, column = 0, cells = [], align = "global"): #"center", "w", "e" or "global"
        if isinstance(row, str) and row.lower() == "all" and align == "global":
            for (r, c) in self.cell_options:
                if 'align' in self.cell_options[(r, c)]:
                    del self.cell_options[(r, c)]['align']
            return
        if align == "global":
            if cells:
                for r, c in cells:
                    if (r, c) in self.cell_options and 'align' in self.cell_options[(r, c)]:
                        del self.cell_options[(r, c)]['align']
            else:
                if (row, column) in self.cell_options and 'align' in self.cell_options[(row, column)]:
                    del self.cell_options[(row, column)]['align']
        else:
            if cells:
                for r, c in cells:
                    if (r, c) not in self.cell_options:
                        self.cell_options[(r, c)] = {}
                    self.cell_options[(r, c)]['align'] = align
            else:
                if (row, column) not in self.cell_options:
                    self.cell_options[(row, column)] = {}
                self.cell_options[(row, column)]['align'] = align

    def readonly_rows(self, rows = [], readonly = True):
        if isinstance(rows, int):
            rows_ = [rows]
        else:
            rows_ = rows
        if not readonly:
            for r in rows_:
                if r in self.row_options and 'readonly' in self.row_options[r]:
                    del self.row_options[r]['readonly']
        else:
            for r in rows_:
                if r not in self.row_options:
                    self.row_options[r] = {}
                self.row_options[r]['readonly'] = True

    def readonly_columns(self, columns = [], readonly = True):
        if isinstance(columns, int):
            cols_ = [columns]
        else:
            cols_ = columns
        if not readonly:
            for c in cols_:
                if c in self.col_options and 'readonly' in self.col_options[c]:
                    del self.col_options[c]['readonly']
        else:
            for c in cols_:
                if c not in self.col_options:
                    self.col_options[c] = {}
                self.col_options[c]['readonly'] = True

    def readonly_cells(self, row = 0, column = 0, cells = [], readonly = True):
        if not readonly:
            if cells:
                for r, c in cells:
                    if (r, c) in self.cell_options and 'readonly' in self.cell_options[(r, c)]:
                        del self.cell_options[(r, c)]['readonly']
            else:
                if (row, column) in self.cell_options and 'readonly' in self.cell_options[(row, column)]:
                    del self.cell_options[(row, column)]['readonly']
        else:
            if cells:
                for (r, c) in cells:
                    if (r, c) not in self.cell_options:
                        self.cell_options[(r, c)] = {}
                    self.cell_options[(r, c)]['readonly'] = True
            else:
                if (row, column) not in self.cell_options:
                    self.cell_options[(row, column)] = {}
                self.cell_options[(row, column)]['readonly'] = True

    def highlight_cells(self, r = 0, c = 0, cells = tuple(), bg = None, fg = None, redraw = False, overwrite = True):
        if bg is None and fg is None:
            return
        if cells:
            for r_, c_ in cells:
                if (r_, c_) not in self.cell_options:
                    self.cell_options[(r_, c_)] = {}
                if 'highlight' in self.cell_options[(r_, c_)] and not overwrite:
                    self.cell_options[(r_, c_)]['highlight'] = (self.cell_options[(r_, c_)]['highlight'][0] if bg is None else bg,
                                                                self.cell_options[(r_, c_)]['highlight'][1] if fg is None else fg)
                else:
                    self.cell_options[(r_, c_)]['highlight'] = (bg, fg)
        else:
            if isinstance(r, str) and r.lower() == "all" and isinstance(c, int):
                riter = range(self.total_data_rows())
                citer = (c, )
            elif isinstance(c, str) and c.lower() == "all" and isinstance(r, int):
                riter = (r, )
                citer = range(self.total_data_cols())
            elif isinstance(r, int) and isinstance(c, int):
                riter = (r, )
                citer = (c, )
            for r_ in riter:
                for c_ in citer:
                    if (r_, c_) not in self.cell_options:
                        self.cell_options[(r_, c_)] = {}
                    if 'highlight' in self.cell_options[(r_, c_)] and not overwrite:
                        self.cell_options[(r_, c_)]['highlight'] = (self.cell_options[(r_, c_)]['highlight'][0] if bg is None else bg,
                                                                    self.cell_options[(r_, c_)]['highlight'][1] if fg is None else fg)
                    else:
                        self.cell_options[(r_, c_)]['highlight'] = (bg, fg)
        if redraw:
            self.main_table_redraw_grid_and_text()

    def highlight_cols(self, cols = [], bg = None, fg = None, highlight_header = False, redraw = False, overwrite = True):
        if bg is None and fg is None:
            return
        for c in (cols, ) if isinstance(cols, int) else cols:
            if c not in self.col_options:
                self.col_options[c] = {}
            if 'highlight' in self.col_options[c] and not overwrite:
                self.col_options[c]['highlight'] = (self.col_options[c]['highlight'][0] if bg is None else bg,
                                                    self.col_options[c]['highlight'][1] if fg is None else fg)
            else:
                self.col_options[c]['highlight'] = (bg, fg)
        if highlight_header:
            self.CH.highlight_cells(cells = cols, bg = bg, fg = fg)
        if redraw:
            self.main_table_redraw_grid_and_text(redraw_header = highlight_header)

    def highlight_rows(self, rows = [], bg = None, fg = None, highlight_index = False, redraw = False, end_of_screen = False, overwrite = True):
        if bg is None and fg is None:
            return
        for r in (rows, ) if isinstance(rows, int) else rows:
            if r not in self.row_options:
                self.row_options[r] = {}
            if 'highlight' in self.row_options[r] and not overwrite:
                self.row_options[r]['highlight'] = (self.row_options[r]['highlight'][0] if bg is None else bg,
                                                    self.row_options[r]['highlight'][1] if fg is None else fg,
                                                    self.row_options[r]['highlight'][2] if self.row_options[r]['highlight'][2] != end_of_screen else end_of_screen)
            else:
                self.row_options[r]['highlight'] = (bg, fg, end_of_screen)
        if highlight_index:
            self.RI.highlight_cells(cells = rows, bg = bg, fg = fg)
        if redraw:
            self.main_table_redraw_grid_and_text(redraw_row_index = highlight_index)

    def deselect(self, r = None, c = None, cell = None, redraw = True):
        deselected = tuple()
        deleted_boxes = {}
        if r == "all":
            deselected = ("deselect_all", self.delete_selection_rects())
        elif r == "allrows":
            for item in self.find_withtag("RowSelectFill"):
                alltags = self.gettags(item)
                if alltags:
                    r1, c1, r2, c2 = tuple(int(e) for e in alltags[1].split("_") if e)
                    deleted_boxes[r1, c1, r2, c2] = "rows"
                    self.delete(alltags[1])
                    self.RI.delete(alltags[1])
                    self.CH.delete(alltags[1])
            current = self.currently_selected()
            if current and current.type_ == "row":
                deleted_boxes[tuple(int(e) for e in self.get_tags_of_current()[1].split("_") if e)] = "cell"
                self.delete_current()
            deselected = ("deselect_all_rows", deleted_boxes)
        elif r == "allcols":
            for item in self.find_withtag("ColSelectFill"):
                alltags = self.gettags(item)
                if alltags:
                    r1, c1, r2, c2 = tuple(int(e) for e in alltags[1].split("_") if e)
                    deleted_boxes[r1, c1, r2, c2] = "columns"
                    self.delete(alltags[1])
                    self.RI.delete(alltags[1])
                    self.CH.delete(alltags[1])
            current = self.currently_selected()
            if current and current.type_ == "column":
                deleted_boxes[tuple(int(e) for e in self.get_tags_of_current()[1].split("_") if e)] = "cell"
                self.delete_current()
            deselected = ("deselect_all_cols", deleted_boxes)
        elif r is not None and c is None and cell is None:
            current = self.find_withtag("Current_Inside") + self.find_withtag("Current_Outside")
            current_tags = self.gettags(current[0]) if current else tuple()
            if current:
                curr_r1, curr_c1, curr_r2, curr_c2 = tuple(int(e) for e in current_tags[1].split("_") if e)
            reset_current = False
            for item in self.find_withtag("RowSelectFill"):
                alltags = self.gettags(item)
                if alltags:
                    r1, c1, r2, c2 = tuple(int(e) for e in alltags[1].split("_") if e)
                    if r >= r1 and r < r2:
                        self.delete(f"{r1}_{c1}_{r2}_{c2}")
                        self.RI.delete(f"{r1}_{c1}_{r2}_{c2}")
                        self.CH.delete(f"{r1}_{c1}_{r2}_{c2}")
                    if not reset_current and current and curr_r1 >= r1 and curr_r1 < r2:
                        reset_current = True
                        deleted_boxes[curr_r1, curr_c1, curr_r2, curr_c2] = "cell"
                    deleted_boxes[r1, c1, r2, c2] = "rows"
            if reset_current:
                self.delete_current()
                self.set_current_to_last()
            deselected = ("deselect_row", deleted_boxes)
        elif c is not None and r is None and cell is None:
            current = self.find_withtag("Current_Inside") + self.find_withtag("Current_Outside")
            current_tags = self.gettags(current[0]) if current else tuple()
            if current:
                curr_r1, curr_c1, curr_r2, curr_c2 = tuple(int(e) for e in current_tags[1].split("_") if e)
            reset_current = False
            for item in self.find_withtag("ColSelectFill"):
                alltags = self.gettags(item)
                if alltags:
                    r1, c1, r2, c2 = tuple(int(e) for e in alltags[1].split("_") if e)
                    if c >= c1 and c < c2:
                        self.delete(f"{r1}_{c1}_{r2}_{c2}")
                        self.RI.delete(f"{r1}_{c1}_{r2}_{c2}")
                        self.CH.delete(f"{r1}_{c1}_{r2}_{c2}")
                    if not reset_current and current and curr_c1 >= c1 and curr_c1 < c2:
                        reset_current = True
                        deleted_boxes[curr_r1, curr_c1, curr_r2, curr_c2] = "cell"
                    deleted_boxes[r1, c1, r2, c2] = "columns"
            if reset_current:
                self.delete_current()
                self.set_current_to_last()
            deselected = ("deselect_column", deleted_boxes)
        elif (r is not None and c is not None and cell is None) or cell is not None:
            set_curr = False
            if cell is not None:
                r, c = cell[0], cell[1]
            for item in chain(self.find_withtag("CellSelectFill"),
                              self.find_withtag("RowSelectFill"),
                              self.find_withtag("ColSelectFill"),
                              self.find_withtag("Current_Outside"),
                              self.find_withtag("Current_Inside")):
                alltags = self.gettags(item)
                if alltags:
                    r1, c1, r2, c2 = tuple(int(e) for e in alltags[1].split("_") if e)
                    if (r >= r1 and
                        c >= c1 and
                        r < r2 and
                        c < c2):
                        current = self.currently_selected()
                        if (not set_curr and
                            current and
                            r2 - r1 == 1 and
                            c2 - c1 == 1 and
                            r == current[0] and
                            c == current[1]):
                            set_curr = True
                        if current and not set_curr:
                            if (current[0] >= r1 and
                                current[0] < r2 and
                                current[1] >= c1 and
                                current[1] < c2):
                                set_curr = True
                        self.delete(f"{r1}_{c1}_{r2}_{c2}")
                        self.RI.delete(f"{r1}_{c1}_{r2}_{c2}")
                        self.CH.delete(f"{r1}_{c1}_{r2}_{c2}")
                        deleted_boxes[(r1, c1, r2, c2)] = "cells"
            if set_curr:
                try:
                    deleted_boxes[tuple(int(e) for e in self.get_tags_of_current()[1].split("_") if e)] = "cells"
                except:
                    pass
                self.delete_current()
                self.set_current_to_last()
            deselected = ("deselect_cell", deleted_boxes)
        if redraw:
            self.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
        if self.deselection_binding_func is not None:
            self.deselection_binding_func(DeselectionEvent(*deselected))

    def page_UP(self, event = None):
        if not self.arrowkeys_enabled:
            return
        height = self.winfo_height()
        top = self.canvasy(0)
        scrollto = top - height
        if scrollto < 0:
            scrollto = 0
        if self.page_up_down_select_row:
            r = bisect.bisect_left(self.row_positions, scrollto)
            current = self.currently_selected()
            if current and current[0] == r:
                r -= 1
            if r < 0:
                r = 0
            if self.RI.row_selection_enabled and (self.anything_selected(exclude_columns = True, exclude_cells = True) or not self.anything_selected()):
                self.RI.select_row(r)
                self.see(r, 0, keep_xscroll = True, check_cell_visibility = False)
            elif (self.single_selection_enabled or self.toggle_selection_enabled) and self.anything_selected(exclude_columns = True, exclude_rows = True):
                box = self.get_all_selection_boxes_with_types()[0][0]
                self.see(r, box[1], keep_xscroll = True, check_cell_visibility = False)
                self.select_cell(r, box[1])
        else:
            args = ("moveto", scrollto / (self.row_positions[-1] + 100))
            self.yview(*args)
            self.RI.yview(*args)
            self.main_table_redraw_grid_and_text(redraw_row_index = True)

    def page_DOWN(self, event = None):
        if not self.arrowkeys_enabled:
            return
        height = self.winfo_height()
        top = self.canvasy(0)
        scrollto = top + height
        if self.page_up_down_select_row and self.RI.row_selection_enabled:
            r = bisect.bisect_left(self.row_positions, scrollto) - 1
            current = self.currently_selected()
            if current and current[0] == r:
                r += 1
            if r > len(self.row_positions) - 2:
                r = len(self.row_positions) - 2
            if self.RI.row_selection_enabled and (self.anything_selected(exclude_columns = True, exclude_cells = True) or not self.anything_selected()):
                self.RI.select_row(r)
                self.see(r, 0, keep_xscroll = True, check_cell_visibility = False)
            elif (self.single_selection_enabled or self.toggle_selection_enabled) and self.anything_selected(exclude_columns = True, exclude_rows = True):
                box = self.get_all_selection_boxes_with_types()[0][0]
                self.see(r, box[1], keep_xscroll = True, check_cell_visibility = False)
                self.select_cell(r, box[1])
        else:
            end = self.row_positions[-1]
            if scrollto > end  + 100:
                scrollto = end
            args = ("moveto", scrollto / (end + 100))
            self.yview(*args)
            self.RI.yview(*args)
            self.main_table_redraw_grid_and_text(redraw_row_index = True)
        
    def arrowkey_UP(self, event = None):
        currently_selected = self.currently_selected()
        if not currently_selected or not self.arrowkeys_enabled:
            return
        if currently_selected.type_ == "row":
            r = currently_selected.row
            if r != 0 and self.RI.row_selection_enabled:
                if self.cell_completely_visible(r = r - 1, c = 0):
                    self.RI.select_row(r - 1, redraw = True)
                else:
                    self.RI.select_row(r - 1)
                    self.see(r - 1, 0, keep_xscroll = True, check_cell_visibility = False)
        elif currently_selected.type_ in ("cell", "column"):
            r = currently_selected[0]
            c = currently_selected[1]
            if r == 0 and self.CH.col_selection_enabled:
                if not self.cell_completely_visible(r = r, c = 0):
                    self.see(r, c, keep_xscroll = True, check_cell_visibility = False)
            elif r != 0 and (self.single_selection_enabled or self.toggle_selection_enabled):
                if self.cell_completely_visible(r = r - 1, c = c):
                    self.select_cell(r - 1, c, redraw = True)
                else:
                    self.select_cell(r - 1, c)
                    self.see(r - 1, c, keep_xscroll = True, check_cell_visibility = False)
                
    def arrowkey_RIGHT(self, event = None):
        currently_selected = self.currently_selected()
        if not currently_selected or not self.arrowkeys_enabled:
            return
        if currently_selected.type_ == "row":
            r = currently_selected.row
            if self.single_selection_enabled or self.toggle_selection_enabled:
                if self.cell_completely_visible(r = r, c = 0):
                    self.select_cell(r, 0, redraw = True)
                else:
                    self.select_cell(r, 0)
                    self.see(r, 0, keep_yscroll = True, bottom_right_corner = True, check_cell_visibility = False)
        elif currently_selected.type_ == "column":
            c = currently_selected.column
            if c < len(self.col_positions) - 2 and self.CH.col_selection_enabled:
                if self.cell_completely_visible(r = 0, c = c + 1):
                    self.CH.select_col(c + 1, redraw = True)
                else:
                    self.CH.select_col(c + 1)
                    self.see(0, c + 1, keep_yscroll = True, bottom_right_corner = False if self.arrow_key_down_right_scroll_page else True, check_cell_visibility = False)
        else:
            r = currently_selected[0]
            c = currently_selected[1]
            if c < len(self.col_positions) - 2 and (self.single_selection_enabled or self.toggle_selection_enabled):
                if self.cell_completely_visible(r = r, c = c + 1):
                    self.select_cell(r, c + 1, redraw =True)
                else:
                    self.select_cell(r, c + 1)
                    self.see(r, c + 1, keep_yscroll = True, bottom_right_corner = False if self.arrow_key_down_right_scroll_page else True, check_cell_visibility = False)

    def arrowkey_DOWN(self, event = None):
        currently_selected = self.currently_selected()
        if not currently_selected or not self.arrowkeys_enabled:
            return
        if currently_selected.type_ == "row":
            r = currently_selected.row
            if r < len(self.row_positions) - 2 and self.RI.row_selection_enabled:
                if self.cell_completely_visible(r = min(r + 2, len(self.row_positions) - 2), c = 0):
                    self.RI.select_row(r + 1, redraw = True)
                else:
                    self.RI.select_row(r + 1)
                    if r + 2 < len(self.row_positions) - 2 and (self.row_positions[r + 3] - self.row_positions[r + 2]) + (self.row_positions[r + 2] - self.row_positions[r + 1]) + 5 < self.winfo_height():
                        self.see(r + 2, 0, keep_xscroll = True, bottom_right_corner = True, check_cell_visibility = False)
                    elif not self.cell_completely_visible(r = r + 1, c = 0):
                        self.see(r + 1, 0, keep_xscroll = True, bottom_right_corner = False if self.arrow_key_down_right_scroll_page else True, check_cell_visibility = False)
        elif currently_selected.type_ == "column":
            c = currently_selected.column
            if self.single_selection_enabled or self.toggle_selection_enabled:
                if self.cell_completely_visible(r = 0, c = c):
                    self.select_cell(0, c, redraw = True)
                else:
                    self.select_cell(0, c)
                    self.see(0, c, keep_xscroll = True, bottom_right_corner = True, check_cell_visibility = False)
        else:
            r = currently_selected[0]
            c = currently_selected[1]
            if r < len(self.row_positions) - 2 and (self.single_selection_enabled or self.toggle_selection_enabled):
                if self.cell_completely_visible(r = min(r + 2, len(self.row_positions) - 2), c = c):
                    self.select_cell(r + 1, c, redraw = True)
                else:
                    self.select_cell(r + 1, c)
                    if r + 2 < len(self.row_positions) - 2 and (self.row_positions[r + 3] - self.row_positions[r + 2]) + (self.row_positions[r + 2] - self.row_positions[r + 1]) + 5 < self.winfo_height():
                        self.see(r + 2, c, keep_xscroll = True, bottom_right_corner = True, check_cell_visibility = False)
                    elif not self.cell_completely_visible(r = r + 1, c = c):
                        self.see(r + 1, c, keep_xscroll = True, bottom_right_corner = False if self.arrow_key_down_right_scroll_page else True, check_cell_visibility = False)
                    
    def arrowkey_LEFT(self, event = None):
        currently_selected = self.currently_selected()
        if not currently_selected or not self.arrowkeys_enabled:
            return
        if currently_selected.type_ == "column":
            c = currently_selected.column
            if c != 0 and self.CH.col_selection_enabled:
                if self.cell_completely_visible(r = 0, c = c - 1):
                    self.CH.select_col(c - 1, redraw = True)
                else:
                    self.CH.select_col(c - 1)
                    self.see(0, c - 1, keep_yscroll = True, bottom_right_corner = True, check_cell_visibility = False)
        elif currently_selected.type_ == "cell":
            r = currently_selected.row
            c = currently_selected.column
            if c == 0 and self.RI.row_selection_enabled:
                if not self.cell_completely_visible(r = r, c = 0):
                    self.see(r, c, keep_yscroll = True, check_cell_visibility = False)
            elif c != 0 and (self.single_selection_enabled or self.toggle_selection_enabled):
                if self.cell_completely_visible(r = r, c = c - 1):
                    self.select_cell(r, c - 1, redraw = True)
                else:
                    self.select_cell(r, c - 1)
                    self.see(r, c - 1, keep_yscroll = True, check_cell_visibility = False)

    def edit_bindings(self, enable = True, key = None):
        if key is None or key == "copy":
            if enable:
                for s2 in ("c", "C"):
                    for widget in (self, self.RI, self.CH, self.TL):
                        widget.bind(f"<{ctrl_key}-{s2}>", self.ctrl_c)
                self.copy_enabled = True
            else:
                for s1 in ("Control", "Command"):
                    for s2 in ("c", "C"):
                        for widget in (self, self.RI, self.CH, self.TL):
                            widget.unbind(f"<{s1}-{s2}>")
                self.copy_enabled = False
        if key is None or key == "cut":
            if enable:
                for s2 in ("x", "X"):
                    for widget in (self, self.RI, self.CH, self.TL):
                        widget.bind(f"<{ctrl_key}-{s2}>", self.ctrl_x)
                self.cut_enabled = True
            else:
                for s1 in ("Control", "Command"):
                    for s2 in ("x", "X"):
                        for widget in (self, self.RI, self.CH, self.TL):
                            widget.unbind(f"<{s1}-{s2}>")
                self.cut_enabled = False
        if key is None or key == "paste":
            if enable:
                for s2 in ("v", "V"):
                    for widget in (self, self.RI, self.CH, self.TL):
                        widget.bind(f"<{ctrl_key}-{s2}>", self.ctrl_v)
                self.paste_enabled = True
            else:
                for s1 in ("Control", "Command"):
                    for s2 in ("v", "V"):
                        for widget in (self, self.RI, self.CH, self.TL):
                            widget.unbind(f"<{s1}-{s2}>")
                self.paste_enabled = False
        if key is None or key == "undo":
            if enable:
                for s2 in ("z", "Z"):
                    for widget in (self, self.RI, self.CH, self.TL):
                        widget.bind(f"<{ctrl_key}-{s2}>", self.ctrl_z)
                self.undo_enabled = True
            else:
                for s1 in ("Control", "Command"):
                    for s2 in ("z", "Z"):
                        for widget in (self, self.RI, self.CH, self.TL):
                            widget.unbind(f"<{s1}-{s2}>")
                self.undo_enabled = False
        if key is None or key == "delete":
            if enable:
                for widget in (self, self.RI, self.CH, self.TL):
                    widget.bind("<Delete>", self.delete_key)
                self.delete_key_enabled = True
            else:
                for widget in (self, self.RI, self.CH, self.TL):
                    widget.unbind("<Delete>")
                self.delete_key_enabled = False
        if key is None or key == "edit_cell":
            if enable:
                self.bind_cell_edit(True)
            else:
                self.bind_cell_edit(False)
        # edit header with text editor (dropdowns and checkboxes not included)
        # this will not by enabled by using enable_bindings() to enable all bindings
        # must be enabled directly using enable_bindings("edit_header")
        # the same goes for "edit_index"
        if key == "edit_header":
            if enable:
                self.CH.bind_cell_edit(True)
            else:
                self.CH.bind_cell_edit(False)
        if key == "edit_index":
            if enable:
                self.RI.bind_cell_edit(True)
            else:
                self.RI.bind_cell_edit(False)

    def menu_add_command(self, menu: tk.Menu, **kwargs):
        if 'label' not in kwargs:
            return
        try:
            index = menu.index(kwargs['label'])
            menu.delete(index)
        except TclError:
            pass
        menu.add_command(**kwargs)

    def create_rc_menus(self):
        if not self.rc_popup_menu:
            self.rc_popup_menu = tk.Menu(self, tearoff = 0, background = self.popup_menu_bg)
        if not self.CH.ch_rc_popup_menu:
            self.CH.ch_rc_popup_menu = tk.Menu(self.CH, tearoff = 0, background = self.popup_menu_bg)
        if not self.RI.ri_rc_popup_menu:
            self.RI.ri_rc_popup_menu = tk.Menu(self.RI, tearoff = 0, background = self.popup_menu_bg)
        if not self.empty_rc_popup_menu:
            self.empty_rc_popup_menu = tk.Menu(self, tearoff = 0, background = self.popup_menu_bg)
        for menu in (self.rc_popup_menu,
                     self.CH.ch_rc_popup_menu,
                     self.RI.ri_rc_popup_menu,
                     self.empty_rc_popup_menu):
            menu.delete(0, 'end')
        if self.rc_popup_menus_enabled and self.CH.edit_cell_enabled:
            self.menu_add_command(self.CH.ch_rc_popup_menu,
                                  label = "Edit header",
                                  font = self.popup_menu_font,
                                  foreground = self.popup_menu_fg,
                                  background = self.popup_menu_bg,
                                  activebackground = self.popup_menu_highlight_bg,
                                  activeforeground = self.popup_menu_highlight_fg,
                                  command = lambda: self.CH.open_cell(event = "rc"))
        if self.rc_popup_menus_enabled and self.RI.edit_cell_enabled:
            self.menu_add_command(self.RI.ri_rc_popup_menu,
                                  label = "Edit index",
                                  font = self.popup_menu_font,
                                  foreground = self.popup_menu_fg,
                                  background = self.popup_menu_bg,
                                  activebackground = self.popup_menu_highlight_bg,
                                  activeforeground = self.popup_menu_highlight_fg,
                                  command = lambda: self.RI.open_cell(event = "rc"))
        if self.rc_popup_menus_enabled and self.edit_cell_enabled:
            self.menu_add_command(self.rc_popup_menu,
                                  label = "Edit cell",
                                  font = self.popup_menu_font,
                                  foreground = self.popup_menu_fg,
                                  background = self.popup_menu_bg,
                                  activebackground = self.popup_menu_highlight_bg,
                                  activeforeground = self.popup_menu_highlight_fg,
                                  command = lambda: self.open_cell(event = "rc"))
        if self.cut_enabled:
            self.menu_add_command(self.rc_popup_menu,
                                  label = "Cut",
                                  accelerator = "Ctrl+X",
                                  font = self.popup_menu_font,
                                  foreground = self.popup_menu_fg,
                                  background = self.popup_menu_bg,
                                  activebackground = self.popup_menu_highlight_bg,
                                  activeforeground = self.popup_menu_highlight_fg,
                                  command = self.ctrl_x)
            self.menu_add_command(self.CH.ch_rc_popup_menu,
                                  label = "Cut contents",
                                  accelerator = "Ctrl+X",
                                  font = self.popup_menu_font,
                                  foreground = self.popup_menu_fg,
                                  background = self.popup_menu_bg,
                                  activebackground = self.popup_menu_highlight_bg,
                                  activeforeground = self.popup_menu_highlight_fg,
                                  command = self.ctrl_x)
            self.menu_add_command(self.RI.ri_rc_popup_menu,
                                  label = "Cut contents",
                                  accelerator = "Ctrl+X",
                                  font = self.popup_menu_font,
                                  foreground = self.popup_menu_fg,
                                  background = self.popup_menu_bg,
                                  activebackground = self.popup_menu_highlight_bg,
                                  activeforeground = self.popup_menu_highlight_fg,
                                  command = self.ctrl_x)
        if self.copy_enabled:
            self.menu_add_command(self.rc_popup_menu, 
                                  label = "Copy",
                                  accelerator = "Ctrl+C",
                                  font = self.popup_menu_font,
                                  foreground = self.popup_menu_fg,
                                  background = self.popup_menu_bg,
                                  activebackground = self.popup_menu_highlight_bg,
                                  activeforeground = self.popup_menu_highlight_fg,
                                  command = self.ctrl_c)
            self.menu_add_command(self.CH.ch_rc_popup_menu, 
                                  label = "Copy contents",
                                  accelerator = "Ctrl+C",
                                  font = self.popup_menu_font,
                                  foreground = self.popup_menu_fg,
                                  background = self.popup_menu_bg,
                                  activebackground = self.popup_menu_highlight_bg,
                                  activeforeground = self.popup_menu_highlight_fg,
                                  command = self.ctrl_c)
            self.menu_add_command(self.RI.ri_rc_popup_menu, 
                                  label = "Copy contents",
                                  accelerator = "Ctrl+C",
                                  font = self.popup_menu_font,
                                  foreground = self.popup_menu_fg,
                                  background = self.popup_menu_bg,
                                  activebackground = self.popup_menu_highlight_bg,
                                  activeforeground = self.popup_menu_highlight_fg,
                                  command = self.ctrl_c)
        if self.paste_enabled:
            self.menu_add_command(self.rc_popup_menu, 
                                  label = "Paste",
                                  accelerator = "Ctrl+V",
                                  font = self.popup_menu_font,
                                  foreground = self.popup_menu_fg,
                                  background = self.popup_menu_bg,
                                  activebackground = self.popup_menu_highlight_bg,
                                  activeforeground = self.popup_menu_highlight_fg,
                                  command = self.ctrl_v)
            self.menu_add_command(self.CH.ch_rc_popup_menu, 
                                  label = "Paste",
                                  accelerator = "Ctrl+V",
                                  font = self.popup_menu_font,
                                  foreground = self.popup_menu_fg,
                                  background = self.popup_menu_bg,
                                  activebackground = self.popup_menu_highlight_bg,
                                  activeforeground = self.popup_menu_highlight_fg,
                                  command = self.ctrl_v)
            self.menu_add_command(self.RI.ri_rc_popup_menu, 
                                  label = "Paste",
                                  accelerator = "Ctrl+V",
                                  font = self.popup_menu_font,
                                  foreground = self.popup_menu_fg,
                                  background = self.popup_menu_bg,
                                  activebackground = self.popup_menu_highlight_bg,
                                  activeforeground = self.popup_menu_highlight_fg,
                                  command = self.ctrl_v)
            if self.expand_sheet_if_paste_too_big:
                self.menu_add_command(self.empty_rc_popup_menu, 
                                      label = "Paste",
                                      accelerator = "Ctrl+V",
                                      font = self.popup_menu_font,
                                      foreground = self.popup_menu_fg,
                                      background = self.popup_menu_bg,
                                      activebackground = self.popup_menu_highlight_bg,
                                      activeforeground = self.popup_menu_highlight_fg,
                                      command = self.ctrl_v)
        if self.delete_key_enabled:
            self.menu_add_command(self.rc_popup_menu, 
                                  label = "Delete",
                                  accelerator = "Del",
                                  font = self.popup_menu_font,
                                  foreground = self.popup_menu_fg,
                                  background = self.popup_menu_bg,
                                  activebackground = self.popup_menu_highlight_bg,
                                  activeforeground = self.popup_menu_highlight_fg,
                                  command = self.delete_key)
            self.menu_add_command(self.CH.ch_rc_popup_menu,
                                  label = "Clear contents",
                                  accelerator = "Del",
                                  font = self.popup_menu_font,
                                  foreground = self.popup_menu_fg,
                                  background = self.popup_menu_bg,
                                  activebackground = self.popup_menu_highlight_bg,
                                  activeforeground = self.popup_menu_highlight_fg,
                                  command = self.delete_key)
            self.menu_add_command(self.RI.ri_rc_popup_menu, 
                                  label = "Clear contents",
                                  accelerator = "Del",
                                  font = self.popup_menu_font,
                                  foreground = self.popup_menu_fg,
                                  background = self.popup_menu_bg,
                                  activebackground = self.popup_menu_highlight_bg,
                                  activeforeground = self.popup_menu_highlight_fg,
                                  command = self.delete_key)
        if self.rc_delete_column_enabled:
            self.menu_add_command(self.CH.ch_rc_popup_menu,
                                  label = "Delete columns",
                                  font = self.popup_menu_font,
                                  foreground = self.popup_menu_fg,
                                  background = self.popup_menu_bg,
                                  activebackground = self.popup_menu_highlight_bg,
                                  activeforeground = self.popup_menu_highlight_fg,
                                  command = self.del_cols_rc)
        if self.rc_insert_column_enabled:
            self.menu_add_command(self.CH.ch_rc_popup_menu,
                                  label = "Insert columns left",
                                  font = self.popup_menu_font,
                                  foreground = self.popup_menu_fg,
                                  background = self.popup_menu_bg,
                                  activebackground = self.popup_menu_highlight_bg,
                                  activeforeground = self.popup_menu_highlight_fg,
                                  command = lambda: self.insert_col_rc("left"))
            self.menu_add_command(self.empty_rc_popup_menu, 
                                  label = "Insert column",
                                  font = self.popup_menu_font,
                                  foreground = self.popup_menu_fg,
                                  background = self.popup_menu_bg,
                                  activebackground = self.popup_menu_highlight_bg,
                                  activeforeground = self.popup_menu_highlight_fg,
                                  command = lambda: self.insert_col_rc("left"))
            self.menu_add_command(self.CH.ch_rc_popup_menu, 
                                  label = "Insert columns right",
                                  font = self.popup_menu_font,
                                  foreground = self.popup_menu_fg,
                                  background = self.popup_menu_bg,
                                  activebackground = self.popup_menu_highlight_bg,
                                  activeforeground = self.popup_menu_highlight_fg,
                                  command = lambda: self.insert_col_rc("right"))
        if self.rc_delete_row_enabled:
            self.menu_add_command(self.RI.ri_rc_popup_menu, 
                                  label = "Delete rows",
                                  font = self.popup_menu_font,
                                  foreground = self.popup_menu_fg,
                                  background = self.popup_menu_bg,
                                  activebackground = self.popup_menu_highlight_bg,
                                  activeforeground = self.popup_menu_highlight_fg,
                                  command = self.del_rows_rc)
        if self.rc_insert_row_enabled:
            self.menu_add_command(self.RI.ri_rc_popup_menu,
                                  label = "Insert rows above",
                                  font = self.popup_menu_font,
                                  foreground = self.popup_menu_fg,
                                  background = self.popup_menu_bg,
                                  activebackground = self.popup_menu_highlight_bg,
                                  activeforeground = self.popup_menu_highlight_fg,
                                  command = lambda: self.insert_row_rc("above"))
            self.menu_add_command(self.RI.ri_rc_popup_menu, 
                                  label = "Insert rows below",
                                  font = self.popup_menu_font,
                                  foreground = self.popup_menu_fg,
                                  background = self.popup_menu_bg,
                                  activebackground = self.popup_menu_highlight_bg,
                                  activeforeground = self.popup_menu_highlight_fg,
                                  command = lambda: self.insert_row_rc("below"))
            self.menu_add_command(self.empty_rc_popup_menu, 
                                  label = "Insert row",
                                  font = self.popup_menu_font,
                                  foreground = self.popup_menu_fg,
                                  background = self.popup_menu_bg,
                                  activebackground = self.popup_menu_highlight_bg,
                                  activeforeground = self.popup_menu_highlight_fg,
                                  command = lambda: self.insert_row_rc("below"))
        for label, func in self.extra_table_rc_menu_funcs.items():
            self.menu_add_command(self.rc_popup_menu, 
                                  label = label,
                                  font = self.popup_menu_font,
                                  foreground = self.popup_menu_fg,
                                  background = self.popup_menu_bg,
                                  activebackground = self.popup_menu_highlight_bg,
                                  activeforeground = self.popup_menu_highlight_fg,
                                  command = func)
        for label, func in self.extra_index_rc_menu_funcs.items():
            self.menu_add_command(self.RI.ri_rc_popup_menu, 
                                  label = label,
                                  font = self.popup_menu_font,
                                  foreground = self.popup_menu_fg,
                                  background = self.popup_menu_bg,
                                  activebackground = self.popup_menu_highlight_bg,
                                  activeforeground = self.popup_menu_highlight_fg,
                                  command = func)
        for label, func in self.extra_header_rc_menu_funcs.items():
            self.menu_add_command(self.CH.ch_rc_popup_menu, 
                                  label = label,
                                  font = self.popup_menu_font,
                                  foreground = self.popup_menu_fg,
                                  background = self.popup_menu_bg,
                                  activebackground = self.popup_menu_highlight_bg,
                                  activeforeground = self.popup_menu_highlight_fg,
                                  command = func)
        for label, func in self.extra_empty_space_rc_menu_funcs.items():
            self.menu_add_command(self.empty_rc_popup_menu, 
                                  label = label,
                                  font = self.popup_menu_font,
                                  foreground = self.popup_menu_fg,
                                  background = self.popup_menu_bg,
                                  activebackground = self.popup_menu_highlight_bg,
                                  activeforeground = self.popup_menu_highlight_fg,
                                  command = func)

    def bind_cell_edit(self, enable = True, keys = []):
        if enable:
            self.edit_cell_enabled = True
            for w in (self, self.RI, self.CH):
                w.bind("<Key>", self.open_cell)
        else:
            self.edit_cell_enabled = False
            for w in (self, self.RI, self.CH):
                w.unbind("<Key>")

    def enable_bindings(self, bindings):
        if not bindings:
            self.enable_bindings_internal("all")
        elif isinstance(bindings, (list, tuple)):
            for binding in bindings:
                if isinstance(binding, (list, tuple)):
                    for bind in binding:
                        self.enable_bindings_internal(bind.lower())
                elif isinstance(binding, str):
                    self.enable_bindings_internal(binding.lower())
        elif isinstance(bindings, str):
            self.enable_bindings_internal(bindings.lower())
            
    def disable_bindings(self, bindings):
        if not bindings:
            self.disable_bindings_internal("all")
        elif isinstance(bindings, (list, tuple)):
            for binding in bindings:
                if isinstance(binding, (list, tuple)):
                    for bind in binding:
                        self.disable_bindings_internal(bind.lower())
                elif isinstance(binding, str):
                    self.disable_bindings_internal(binding.lower())
        elif isinstance(bindings, str):
            self.disable_bindings_internal(bindings)

    def enable_disable_select_all(self, enable = True):
        self.select_all_enabled = bool(enable)
        for s in ("A", "a"):
            binding = f"<{ctrl_key}-{s}>"
            for widget in (self, self.RI, self.CH, self.TL):
                if enable:
                    widget.bind(binding, self.select_all)
                else:
                    widget.unbind(binding)

    def enable_bindings_internal(self, binding):
        if binding in ("enable_all", "all"):
            self.single_selection_enabled = True
            self.toggle_selection_enabled = False
            self.drag_selection_enabled = True
            self.enable_disable_select_all(True)
            self.CH.enable_bindings("column_width_resize")
            self.CH.enable_bindings("column_select")
            self.CH.enable_bindings("column_height_resize")
            self.CH.enable_bindings("drag_and_drop")
            self.CH.enable_bindings("double_click_column_resize")
            self.RI.enable_bindings("row_height_resize")
            self.RI.enable_bindings("double_click_row_resize")
            self.RI.enable_bindings("row_width_resize")
            self.RI.enable_bindings("row_select")
            self.RI.enable_bindings("drag_and_drop")
            self.bind_arrowkeys()
            self.edit_bindings(True)
            self.rc_delete_column_enabled = True
            self.rc_delete_row_enabled = True
            self.rc_insert_column_enabled = True
            self.rc_insert_row_enabled = True
            self.rc_popup_menus_enabled = True
            self.rc_select_enabled = True
            self.TL.rh_state()
            self.TL.rw_state()
        elif binding in ("single", "single_selection_mode", "single_select"):
            self.single_selection_enabled = True
            self.toggle_selection_enabled = False
        elif binding in ("toggle", "toggle_selection_mode", "toggle_select"):
            self.toggle_selection_enabled = True
            self.single_selection_enabled = False
        elif binding == "drag_select":
            self.drag_selection_enabled = True
        elif binding == "select_all":
            self.enable_disable_select_all(True)
        elif binding == "column_width_resize":
            self.CH.enable_bindings("column_width_resize")
        elif binding == "column_select":
            self.CH.enable_bindings("column_select")
        elif binding == "column_height_resize":
            self.CH.enable_bindings("column_height_resize")
            self.TL.rh_state()
        elif binding == "column_drag_and_drop":
            self.CH.enable_bindings("drag_and_drop")
        elif binding == "double_click_column_resize":
            self.CH.enable_bindings("double_click_column_resize")
        elif binding == "row_height_resize":
            self.RI.enable_bindings("row_height_resize")
        elif binding == "double_click_row_resize":
            self.RI.enable_bindings("double_click_row_resize")
        elif binding == "row_width_resize":
            self.RI.enable_bindings("row_width_resize")
            self.TL.rw_state()
        elif binding == "row_select":
            self.RI.enable_bindings("row_select")
        elif binding == "row_drag_and_drop":
            self.RI.enable_bindings("drag_and_drop")
        elif binding == "arrowkeys":
            self.bind_arrowkeys()
        elif binding == "edit_bindings":
            self.edit_bindings(True)
        elif binding == "rc_delete_column":
            self.rc_delete_column_enabled = True
            self.rc_popup_menus_enabled = True
            self.rc_select_enabled = True
        elif binding == "rc_delete_row":
            self.rc_delete_row_enabled = True
            self.rc_popup_menus_enabled = True
            self.rc_select_enabled = True
        elif binding == "rc_insert_column":
            self.rc_insert_column_enabled = True
            self.rc_popup_menus_enabled = True
            self.rc_select_enabled = True
        elif binding == "rc_insert_row":
            self.rc_insert_row_enabled = True
            self.rc_popup_menus_enabled = True
            self.rc_select_enabled = True
        elif binding == "copy":
            self.edit_bindings(True, "copy")
        elif binding == "cut":
            self.edit_bindings(True, "cut")
        elif binding == "paste":
            self.edit_bindings(True, "paste")
        elif binding == "delete":
            self.edit_bindings(True, "delete")
        elif binding in ("right_click_popup_menu", "rc_popup_menu"):
            self.rc_popup_menus_enabled = True
            self.rc_select_enabled = True
        elif binding in ("right_click_select", "rc_select"):
            self.rc_select_enabled = True
        elif binding == "undo":
            self.edit_bindings(True, "undo")
        elif binding == "edit_cell":
            self.edit_bindings(True, "edit_cell")
        elif binding == "edit_header":
            self.edit_bindings(True, "edit_header")
        elif binding == "edit_index":
            self.edit_bindings(True, "edit_index")
        self.create_rc_menus()

    def disable_bindings_internal(self, binding):
        if binding in ("all", "disable_all"):
            self.single_selection_enabled = False
            self.toggle_selection_enabled = False
            self.drag_selection_enabled = False
            self.enable_disable_select_all(False)
            self.CH.disable_bindings("column_width_resize")
            self.CH.disable_bindings("column_select")
            self.CH.disable_bindings("column_height_resize")
            self.CH.disable_bindings("drag_and_drop")
            self.CH.disable_bindings("double_click_column_resize")
            self.RI.disable_bindings("row_height_resize")
            self.RI.disable_bindings("double_click_row_resize")
            self.RI.disable_bindings("row_width_resize")
            self.RI.disable_bindings("row_select")
            self.RI.disable_bindings("drag_and_drop")
            self.unbind_arrowkeys()
            self.edit_bindings(False)
            self.rc_delete_column_enabled = False
            self.rc_delete_row_enabled = False
            self.rc_insert_column_enabled = False
            self.rc_insert_row_enabled = False
            self.rc_popup_menus_enabled = False
            self.rc_select_enabled = False
            self.TL.rh_state("hidden")
            self.TL.rw_state("hidden")
        elif binding in ("single", "single_selection_mode", "single_select"):
            self.single_selection_enabled = False
        elif binding in ("toggle", "toggle_selection_mode", "toggle_select"):
            self.toggle_selection_enabled = False
        elif binding == "drag_select":
            self.drag_selection_enabled = False
        elif binding == "select_all":
            self.enable_disable_select_all(False)
        elif binding == "column_width_resize":
            self.CH.disable_bindings("column_width_resize")
        elif binding == "column_select":
            self.CH.disable_bindings("column_select")
        elif binding == "column_height_resize":
            self.CH.disable_bindings("column_height_resize")
            self.TL.rh_state("hidden")
        elif binding == "column_drag_and_drop":
            self.CH.disable_bindings("drag_and_drop")
        elif binding == "double_click_column_resize":
            self.CH.disable_bindings("double_click_column_resize")
        elif binding == "row_height_resize":
            self.RI.disable_bindings("row_height_resize")
        elif binding == "double_click_row_resize":
            self.RI.disable_bindings("double_click_row_resize")
        elif binding == "row_width_resize":
            self.RI.disable_bindings("row_width_resize")
            self.TL.rw_state("hidden")
        elif binding == "row_select":
            self.RI.disable_bindings("row_select")
        elif binding == "row_drag_and_drop":
            self.RI.disable_bindings("drag_and_drop")
        elif binding == "arrowkeys":
            self.unbind_arrowkeys()
        elif binding == "rc_delete_column":
            self.rc_delete_column_enabled = False
        elif binding == "rc_delete_row":
            self.rc_delete_row_enabled = False
        elif binding == "rc_insert_column":
            self.rc_insert_column_enabled = False
        elif binding == "rc_insert_row":
            self.rc_insert_row_enabled = False
        elif binding == "edit_bindings":
            self.edit_bindings(False)
        elif binding == "copy":
            self.edit_bindings(False, "copy")
        elif binding == "cut":
            self.edit_bindings(False, "cut")
        elif binding == "paste":
            self.edit_bindings(False, "paste")
        elif binding == "delete":
            self.edit_bindings(False, "delete")
        elif binding in ("right_click_popup_menu", "rc_popup_menu"):
            self.rc_popup_menus_enabled = False
        elif binding in ("right_click_select", "rc_select"):
            self.rc_select_enabled = False
        elif binding == "undo":
            self.edit_bindings(False, "undo")
        elif binding == "edit_cell":
            self.edit_bindings(False, "edit_cell")
        elif binding == "edit_header":
            self.edit_bindings(False, "edit_header")
        elif binding == "edit_index":
            self.edit_bindings(False, "edit_index")
        self.create_rc_menus()

    def reset_mouse_motion_creations(self, event = None):
        self.config(cursor = "")
        self.RI.config(cursor = "")
        self.CH.config(cursor = "")
        self.RI.rsz_w = None
        self.RI.rsz_h = None
        self.CH.rsz_w = None
        self.CH.rsz_h = None
    
    def mouse_motion(self, event):
        if (
            not self.RI.currently_resizing_height and
            not self.RI.currently_resizing_width and
            not self.CH.currently_resizing_height and
            not self.CH.currently_resizing_width
            ):
            mouse_over_resize = False
            x = self.canvasx(event.x)
            y = self.canvasy(event.y)
            if self.RI.width_resizing_enabled and self.show_index and not mouse_over_resize:
                try:
                    x1, y1, x2, y2 = self.row_width_resize_bbox[0], self.row_width_resize_bbox[1], self.row_width_resize_bbox[2], self.row_width_resize_bbox[3]
                    if x >= x1 and y >= y1 and x <= x2 and y <= y2:
                        self.config(cursor = "sb_h_double_arrow")
                        self.RI.config(cursor = "sb_h_double_arrow")
                        self.RI.rsz_w = True
                        mouse_over_resize = True
                except:
                    pass
            if self.CH.height_resizing_enabled and self.show_header and not mouse_over_resize:
                try:
                    x1, y1, x2, y2 = self.header_height_resize_bbox[0], self.header_height_resize_bbox[1], self.header_height_resize_bbox[2], self.header_height_resize_bbox[3]
                    if x >= x1 and y >= y1 and x <= x2 and y <= y2:
                        self.config(cursor = "sb_v_double_arrow")
                        self.CH.config(cursor = "sb_v_double_arrow")
                        self.CH.rsz_h = True
                        mouse_over_resize = True
                except:
                    pass
            if not mouse_over_resize:
                self.reset_mouse_motion_creations()
        if self.extra_motion_func is not None:
            self.extra_motion_func(event)

    def rc(self, event = None):
        self.mouseclick_outside_editor_or_dropdown()
        self.focus_set()
        popup_menu = None
        if self.single_selection_enabled and all(v is None for v in (self.RI.rsz_h, self.RI.rsz_w, self.CH.rsz_h, self.CH.rsz_w)):
            r = self.identify_row(y = event.y)
            c = self.identify_col(x = event.x)
            if r < len(self.row_positions) - 1 and c < len(self.col_positions) - 1:
                if self.col_selected(c):
                    if self.rc_popup_menus_enabled:
                        popup_menu = self.CH.ch_rc_popup_menu
                elif self.row_selected(r):
                    if self.rc_popup_menus_enabled:
                        popup_menu = self.RI.ri_rc_popup_menu
                elif self.cell_selected(r, c):
                    if self.rc_popup_menus_enabled:
                        popup_menu = self.rc_popup_menu
                else:
                    if self.rc_select_enabled:
                        self.select_cell(r, c, redraw = True)
                    if self.rc_popup_menus_enabled:
                        popup_menu = self.rc_popup_menu
            else:
                popup_menu = self.empty_rc_popup_menu
        elif self.toggle_selection_enabled and all(v is None for v in (self.RI.rsz_h, self.RI.rsz_w, self.CH.rsz_h, self.CH.rsz_w)):
            r = self.identify_row(y = event.y)
            c = self.identify_col(x = event.x)
            if r < len(self.row_positions) - 1 and c < len(self.col_positions) - 1:
                if self.col_selected(c):
                    if self.rc_popup_menus_enabled:
                        popup_menu = self.CH.ch_rc_popup_menu
                elif self.row_selected(r):
                    if self.rc_popup_menus_enabled:
                        popup_menu = self.RI.ri_rc_popup_menu
                elif self.cell_selected(r, c):
                    if self.rc_popup_menus_enabled:
                        popup_menu = self.rc_popup_menu
                else:
                    if self.rc_select_enabled:
                        self.toggle_select_cell(r, c, redraw = True)
                    if self.rc_popup_menus_enabled:
                        popup_menu = self.rc_popup_menu
            else:
                popup_menu = self.empty_rc_popup_menu
        if self.extra_rc_func is not None:
            self.extra_rc_func(event)
        if popup_menu is not None:
            popup_menu.tk_popup(event.x_root, event.y_root)

    def b1_press(self, event = None):
        self.closed_dropdown = self.mouseclick_outside_editor_or_dropdown()
        self.focus_set()
        x1, y1, x2, y2 = self.get_canvas_visible_area()
        if self.identify_col(x = event.x, allow_end = False) is None or self.identify_row(y = event.y, allow_end = False) is None:
            self.deselect("all")
        r = self.identify_row(y = event.y)
        c = self.identify_col(x = event.x)
        if self.single_selection_enabled and all(v is None for v in (self.RI.rsz_h, self.RI.rsz_w, self.CH.rsz_h, self.CH.rsz_w)):
            if r < len(self.row_positions) - 1 and c < len(self.col_positions) - 1:
                self.select_cell(r, c, redraw = True)
        elif self.toggle_selection_enabled and all(v is None for v in (self.RI.rsz_h, self.RI.rsz_w, self.CH.rsz_h, self.CH.rsz_w)):
            r = self.identify_row(y = event.y)
            c = self.identify_col(x = event.x)
            if r < len(self.row_positions) - 1 and c < len(self.col_positions) - 1:
                self.toggle_select_cell(r, c, redraw = True)
        elif self.RI.width_resizing_enabled and self.show_index and self.RI.rsz_h is None and self.RI.rsz_w == True:
            self.RI.currently_resizing_width = True
            self.new_row_width = self.RI.current_width + event.x
            x = self.canvasx(event.x)
            self.create_resize_line(x, y1, x, y2, width = 1, fill = self.RI.resizing_line_fg, tag = "rwl")
        elif self.CH.height_resizing_enabled and self.show_header and self.CH.rsz_w is None and self.CH.rsz_h == True:
            self.CH.currently_resizing_height = True
            self.new_header_height = self.CH.current_height + event.y
            y = self.canvasy(event.y)
            self.create_resize_line(x1, y, x2, y, width = 1, fill = self.RI.resizing_line_fg, tag = "rhl")
        self.b1_pressed_loc = (r, c)
        if self.extra_b1_press_func is not None:
            self.extra_b1_press_func(event)

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

    def shift_b1_press(self, event = None):
        self.mouseclick_outside_editor_or_dropdown()
        self.focus_set()
        if self.drag_selection_enabled and all(v is None for v in (self.RI.rsz_h, self.RI.rsz_w, self.CH.rsz_h, self.CH.rsz_w)):
            self.b1_pressed_loc = None
            rowsel = int(self.identify_row(y = event.y))
            colsel = int(self.identify_col(x = event.x))
            if rowsel < len(self.row_positions) - 1 and colsel < len(self.col_positions) - 1:
                currently_selected = self.currently_selected()
                if currently_selected and currently_selected.type_ == "cell":
                    min_r = currently_selected[0]
                    min_c = currently_selected[1]
                    self.delete_selection_rects(delete_current = False)
                    if rowsel >= min_r and colsel >= min_c:
                        self.create_selected(min_r, min_c, rowsel + 1, colsel + 1)
                    elif rowsel >= min_r and min_c >= colsel:
                        self.create_selected(min_r, colsel, rowsel + 1, min_c + 1)
                    elif min_r >= rowsel and colsel >= min_c:
                        self.create_selected(rowsel, min_c, min_r + 1, colsel + 1)
                    elif min_r >= rowsel and min_c >= colsel:
                        self.create_selected(rowsel, colsel, min_r + 1, min_c + 1)
                    last_selected = tuple(int(e) for e in self.gettags(self.find_withtag("CellSelectFill"))[1].split("_") if e)
                else:
                    self.select_cell(rowsel, colsel, redraw = False)
                    last_selected = tuple(int(e) for e in self.gettags(self.find_withtag("Current_Outside"))[1].split("_") if e)
                self.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True, redraw_table = True)
                if self.shift_selection_binding_func is not None:
                    self.shift_selection_binding_func(SelectionBoxEvent("shift_select_cells", last_selected))
        
    def b1_motion(self, event):
        x1, y1, x2, y2 = self.get_canvas_visible_area()
        if self.drag_selection_enabled and all(v is None for v in (self.RI.rsz_h, self.RI.rsz_w, self.CH.rsz_h, self.CH.rsz_w)):
            need_redraw = False
            end_row = self.identify_row(y = event.y)
            end_col = self.identify_col(x = event.x)
            currently_selected = self.currently_selected()
            if end_row < len(self.row_positions) - 1 and end_col < len(self.col_positions) - 1 and currently_selected and currently_selected.type_ == "cell":
                start_row = currently_selected[0]
                start_col = currently_selected[1]
                rect = None
                if (end_row >= start_row and 
                    end_col >= start_col and
                    (end_row - start_row or end_col - start_col)):
                    rect = (start_row, start_col, end_row + 1, end_col + 1)
                elif (end_row >= start_row and 
                      end_col < start_col and
                      (end_row - start_row or start_col - end_col)):
                    rect = (start_row, end_col, end_row + 1, start_col + 1)
                elif (end_row < start_row and
                      end_col >= start_col and
                      (start_row - end_row or end_col - start_col)):
                    rect = (end_row, start_col, start_row + 1, end_col + 1)
                elif (end_row < start_row and 
                      end_col < start_col and
                      (start_row - end_row or start_col - end_col)):
                    rect = (end_row, end_col, start_row + 1, start_col + 1)
                else:
                    rect = (start_row, start_col, start_row + 1, start_col + 1)
                if rect is not None and self.being_drawn_rect != rect:
                    self.delete_selection_rects(delete_current = False)
                    if rect[2] - rect[0] == 1 and rect[3] - rect[1] == 1:
                        self.being_drawn_rect = rect
                    else:
                        self.being_drawn_rect = rect
                        self.create_selected(*rect)
                    if self.drag_selection_binding_func is not None:
                        self.drag_selection_binding_func(SelectionBoxEvent("drag_select_cells", rect))
                    need_redraw = True
            if self.data:
                xcheck = self.xview()
                ycheck = self.yview()
                if len(xcheck) > 1 and xcheck[0] > 0 and event.x < 0:
                    try:
                        self.xview_scroll(-1, "units")
                        self.CH.xview_scroll(-1, "units")
                    except:
                        pass
                    need_redraw = True
                if len(ycheck) > 1 and ycheck[0] > 0 and event.y < 0:
                    try:
                        self.yview_scroll(-1, "units")
                        self.RI.yview_scroll(-1, "units")
                    except:
                        pass
                    need_redraw = True
                if len(xcheck) > 1 and xcheck[1] < 1 and event.x > self.winfo_width():
                    try:
                        self.xview_scroll(1, "units")
                        self.CH.xview_scroll(1, "units")
                    except:
                        pass
                    need_redraw = True
                if len(ycheck) > 1 and ycheck[1] < 1 and event.y > self.winfo_height():
                    try:
                        self.yview_scroll(1, "units")
                        self.RI.yview_scroll(1, "units")
                    except:
                        pass
                    need_redraw = True
            self.check_views()
            if need_redraw:
                self.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True, redraw_table = True)
        elif self.RI.width_resizing_enabled and self.RI.rsz_w is not None and self.RI.currently_resizing_width:
            self.RI.delete_resize_lines()
            self.delete_resize_lines()
            if event.x >= 0:
                x = self.canvasx(event.x)
                self.new_row_width = self.RI.current_width + event.x
                self.create_resize_line(x, y1, x, y2, width = 1, fill = self.RI.resizing_line_fg, tag = "rwl")
            else:
                x = self.RI.current_width + event.x
                if x < self.min_cw:
                    x = int(self.min_cw)
                self.new_row_width = x
                self.RI.create_resize_line(x, y1, x, y2, width = 1, fill = self.RI.resizing_line_fg, tag = "rwl")
        elif self.CH.height_resizing_enabled and self.CH.rsz_h is not None and self.CH.currently_resizing_height:
            self.CH.delete_resize_lines()
            self.delete_resize_lines()
            if event.y >= 0:
                y = self.canvasy(event.y)
                self.new_header_height = self.CH.current_height + event.y
                self.create_resize_line(x1, y, x2, y, width = 1, fill = self.RI.resizing_line_fg, tag = "rhl")
            else:
                y = self.CH.current_height + event.y
                if y < self.hdr_min_rh:
                    y = int(self.hdr_min_rh)
                self.new_header_height = y
                self.CH.create_resize_line(x1, y, x2, y, width = 1, fill = self.RI.resizing_line_fg, tag = "rhl")
        if self.extra_b1_motion_func is not None:
            self.extra_b1_motion_func(event)

    def b1_release(self, event = None):
        if self.being_drawn_rect is not None and (self.being_drawn_rect[2] - self.being_drawn_rect[0] > 1 or self.being_drawn_rect[3] - self.being_drawn_rect[1] > 1):
            to_sel = tuple(self.being_drawn_rect)
            self.being_drawn_rect = None
            self.create_selected(*to_sel)
        else:
            self.being_drawn_rect = None
        if self.RI.width_resizing_enabled and self.RI.rsz_w is not None and self.RI.currently_resizing_width:
            self.delete_resize_lines()
            self.RI.delete_resize_lines()
            self.RI.currently_resizing_width = False
            self.RI.set_width(self.new_row_width, set_TL = True)
            self.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
            self.b1_pressed_loc = None
        elif self.CH.height_resizing_enabled and self.CH.rsz_h is not None and self.CH.currently_resizing_height:
            self.delete_resize_lines()
            self.CH.delete_resize_lines()
            self.CH.currently_resizing_height = False
            self.CH.set_height(self.new_header_height, set_TL = True)
            self.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
            self.b1_pressed_loc = None
        self.RI.rsz_w = None
        self.CH.rsz_h = None
        if self.b1_pressed_loc is not None:
            r = self.identify_row(y = event.y, allow_end = False)
            c = self.identify_col(x = event.x, allow_end = False)
            if r is not None and c is not None and (r, c) == self.b1_pressed_loc:
                dcol = c if self.all_columns_displayed else self.displayed_columns[c]
                if (r, dcol) in self.cell_options and ('dropdown' in self.cell_options[(r, dcol)] or 'checkbox' in self.cell_options[(r, dcol)]):
                    canvasx = self.canvasx(event.x)
                    if (self.closed_dropdown != self.b1_pressed_loc and
                        'dropdown' in self.cell_options[(r, dcol)] and
                        canvasx > self.col_positions[c + 1] - self.txt_h - 5 and
                        canvasx < self.col_positions[c + 1] - 1):
                        self.open_cell(event)
                    elif 'checkbox' in self.cell_options[(r, dcol)] and event.x < self.col_positions[c] + self.txt_h + 5 and event.y < self.row_positions[r] + self.txt_h + 5:
                        self.open_cell(event)
                        self.mouseclick_outside_editor_or_dropdown()
                    else:
                        self.mouseclick_outside_editor_or_dropdown()
            else:
                self.mouseclick_outside_editor_or_dropdown()
        self.b1_pressed_loc = None
        self.closed_dropdown = None
        if self.extra_b1_release_func is not None:
            self.extra_b1_release_func(event)

    def double_b1(self, event = None):
        self.mouseclick_outside_editor_or_dropdown()
        self.focus_set()
        x1, y1, x2, y2 = self.get_canvas_visible_area()
        if self.identify_col(x = event.x, allow_end = False) is None or self.identify_row(y = event.y, allow_end = False) is None:
            self.deselect("all")
        elif self.single_selection_enabled and all(v is None for v in (self.RI.rsz_h, self.RI.rsz_w, self.CH.rsz_h, self.CH.rsz_w)):
            r = self.identify_row(y = event.y)
            c = self.identify_col(x = event.x)
            if r < len(self.row_positions) - 1 and c < len(self.col_positions) - 1:
                self.select_cell(r, c, redraw = True)
                if self.edit_cell_enabled:
                    self.open_cell(event)
        elif self.toggle_selection_enabled and all(v is None for v in (self.RI.rsz_h, self.RI.rsz_w, self.CH.rsz_h, self.CH.rsz_w)):
            r = self.identify_row(y = event.y)
            c = self.identify_col(x = event.x)
            if r < len(self.row_positions) - 1 and c < len(self.col_positions) - 1:
                self.toggle_select_cell(r, c, redraw = True)
                if self.edit_cell_enabled:
                    self.open_cell(event)
        if self.extra_double_b1_func is not None:
            self.extra_double_b1_func(event)

    def identify_row(self, event = None, y = None, allow_end = True):
        if event is None:
            y2 = self.canvasy(y)
        elif y is None:
            y2 = self.canvasy(event.y)
        r = bisect.bisect_left(self.row_positions, y2)
        if r != 0:
            r -= 1
        if not allow_end and r >= len(self.row_positions) - 1:
            return None
        return r

    def identify_col(self, event = None, x = None, allow_end = True):
        if event is None:
            x2 = self.canvasx(x)
        elif x is None:
            x2 = self.canvasx(event.x)
        c = bisect.bisect_left(self.col_positions, x2)
        if c != 0:
            c -= 1
        if not allow_end and c >= len(self.col_positions) - 1:
            return None
        return c

    def check_views(self):
        xcheck = self.xview()
        ycheck = self.yview()
        if xcheck and xcheck[0] <= 0:
            self.xview(*("moveto", 0))
            if self.show_header:
                self.CH.xview(*("moveto", 0))

        elif len(xcheck) > 1 and xcheck[1] >= 1:
            self.xview(*("moveto", 1))
            if self.show_header:
                self.CH.xview(*("moveto", 1))
                
        if ycheck and ycheck[0] <= 0:
            self.yview(*("moveto", 0))
            if self.show_index:
                self.RI.yview(*("moveto", 0))
        elif len(ycheck) > 1 and ycheck[1] >= 1:
            self.yview(*("moveto", 1))
            if self.show_index:
                self.RI.yview(*("moveto", 1))

    def set_xviews(self, *args):
        self.xview(*args)
        if self.show_header:
            self.CH.xview(*args)
        self.check_views()
        self.main_table_redraw_grid_and_text(redraw_header = True if self.show_header else False)

    def set_yviews(self, *args):
        self.yview(*args)
        if self.show_index:
            self.RI.yview(*args)
        self.check_views()
        self.main_table_redraw_grid_and_text(redraw_row_index = True if self.show_index else False)

    def set_view(self, x_args, y_args):
        self.xview(*x_args)
        if self.show_header:
            self.CH.xview(*x_args)
        self.yview(*y_args)
        if self.show_index:
            self.RI.yview(*y_args)
        self.check_views()
        self.main_table_redraw_grid_and_text(redraw_row_index = True if self.show_index else False,
                                             redraw_header = True if self.show_header else False)

    def mousewheel(self, event = None):
        if event.delta < 0 or event.num == 5:
            self.yview_scroll(1, "units")
            self.RI.yview_scroll(1, "units")
        elif event.delta >= 0 or event.num == 4:
            if self.canvasy(0) <= 0:
                return
            self.yview_scroll(-1, "units")
            self.RI.yview_scroll(-1, "units")
        self.main_table_redraw_grid_and_text(redraw_row_index = True)

    def shift_mousewheel(self, event = None):
        if event.delta < 0 or event.num == 5:
            self.xview_scroll(1, "units")
            self.CH.xview_scroll(1, "units")
        elif event.delta >= 0 or event.num == 4:
            if self.canvasx(0) <= 0:
                return
            self.xview_scroll(-1, "units")
            self.CH.xview_scroll(-1, "units")
        self.main_table_redraw_grid_and_text(redraw_header = True)

    def GetWidthChars(self, width):
        char_w = self.GetTextWidth("_")
        return int(width / char_w)

    def GetTextWidth(self, txt):
        self.txt_measure_canvas.itemconfig(self.txt_measure_canvas_text, text = txt, font = self._font)
        b = self.txt_measure_canvas.bbox(self.txt_measure_canvas_text)
        return b[2] - b[0]

    def GetTextHeight(self, txt):
        self.txt_measure_canvas.itemconfig(self.txt_measure_canvas_text, text = txt, font = self._font)
        b = self.txt_measure_canvas.bbox(self.txt_measure_canvas_text)
        return b[3] - b[1]

    def GetHdrTextWidth(self, txt):
        self.txt_measure_canvas.itemconfig(self.txt_measure_canvas_text, text = txt, font = self._hdr_font)
        b = self.txt_measure_canvas.bbox(self.txt_measure_canvas_text)
        return b[2] - b[0]

    def GetHdrTextHeight(self, txt):
        self.txt_measure_canvas.itemconfig(self.txt_measure_canvas_text, text = txt, font = self._hdr_font)
        b = self.txt_measure_canvas.bbox(self.txt_measure_canvas_text)
        return b[3] - b[1]
    
    def GetLinesHeight(self, n, old_method = False):
        if old_method:
            if n == 1:
                return int(self.min_rh)
            else:
                return int(self.fl_ins) + (self.xtra_lines_increment * n) - 2
        else:
            x = self.txt_measure_canvas.create_text(0, 0,
                                                    text = "\n".join(["j^|" for lines in range(n)]) if n > 1 else "j^|",
                                                    font = self._font)
            b = self.txt_measure_canvas.bbox(x)
            h = b[3] - b[1] + 5
            self.txt_measure_canvas.delete(x)
            return h

    def GetHdrLinesHeight(self, n, old_method = False):
        if old_method:
            if n == 1:
                return int(self.hdr_min_rh)
            else:
                return int(self.hdr_fl_ins) + (self.hdr_xtra_lines_increment * n) - 2
        else:
            x = self.txt_measure_canvas.create_text(0, 0,
                                                    text = "\n".join(["j^|" for lines in range(n)]) if n > 1 else "j^|",
                                                    font = self._hdr_font)
            b = self.txt_measure_canvas.bbox(x)
            h = b[3] - b[1] + 5
            self.txt_measure_canvas.delete(x)
            return h

    def set_min_cw(self):
        #w1 = self.GetHdrTextWidth("X") + 5
        #w2 = self.GetTextWidth("X") + 5
        #if w1 >= w2:
        #    self.min_cw = w1
        #else:
        #    self.min_cw = w2
        self.min_cw = 5
        if self.min_cw > self.CH.max_cw:
            self.CH.max_cw = self.min_cw + 20
        if self.min_cw > self.default_cw:
            self.default_cw = self.min_cw + 20

    def font(self, newfont = None, reset_row_positions = False):
        if newfont:
            if not isinstance(newfont, tuple):
                raise ValueError("Argument must be tuple e.g. ('Carlito',12,'normal')")
            if len(newfont) != 3:
                raise ValueError("Argument must be three-tuple")
            if (
                not isinstance(newfont[0], str) or
                not isinstance(newfont[1], int) or
                not isinstance(newfont[2], str)
                ):
                raise ValueError("Argument must be font, size and 'normal', 'bold' or 'italic' e.g. ('Carlito',12,'normal')")
            else:
                self._font = newfont
            self.fnt_fam = newfont[0]
            self.fnt_sze = newfont[1]
            self.fnt_wgt = newfont[2]
            self.set_fnt_help()
            if reset_row_positions:
                self.reset_row_positions()
        else:
            return self._font

    def set_fnt_help(self):
        self.txt_h = self.GetTextHeight("|ZXjy*'^")
        self.txt_w = self.GetTextWidth("I")
        self.half_txt_h = ceil(self.txt_h / 2)
        if self.half_txt_h % 2 == 0:
            self.fl_ins = self.half_txt_h + 2
        else:
            self.fl_ins = self.half_txt_h + 3
        self.xtra_lines_increment = int(self.txt_h)
        self.min_rh = self.txt_h + 5
        if self.min_rh < 12:
            self.min_rh = 12
        #self.min_rh = 5
        if self.default_rh[0] != "pixels":
            self.default_rh = (self.default_rh[0] if self.default_rh[0] != "pixels" else "pixels",
                               self.GetLinesHeight(int(self.default_rh[0])) if self.default_rh[0] != "pixels" else self.default_rh[1])
        self.set_min_cw()
        
    def header_font(self, newfont = None):
        if newfont:
            if not isinstance(newfont, tuple):
                raise ValueError("Argument must be tuple e.g. ('Carlito', 12, 'normal')")
            if len(newfont) != 3:
                raise ValueError("Argument must be three-tuple")
            if (
                not isinstance(newfont[0], str) or
                not isinstance(newfont[1], int) or
                not isinstance(newfont[2], str)
                ):
                raise ValueError("Argument must be font, size and 'normal', 'bold' or 'italic' e.g. ('Carlito', 12, 'normal')")
            else:
                self._hdr_font = newfont
            self.hdr_fnt_fam = newfont[0]
            self.hdr_fnt_sze = newfont[1]
            self.hdr_fnt_wgt = newfont[2]
            self.set_hdr_fnt_help()
        else:
            return self._hdr_font

    def set_hdr_fnt_help(self):
        self.hdr_txt_h = self.GetHdrTextHeight("|ZXj*'^")
        self.hdr_txt_w = self.GetHdrTextWidth("I")
        self.hdr_half_txt_h = ceil(self.hdr_txt_h / 2)
        if self.hdr_half_txt_h % 2 == 0:
            self.hdr_fl_ins = self.hdr_half_txt_h + 2
        else:
            self.hdr_fl_ins = self.hdr_half_txt_h + 3
        self.hdr_xtra_lines_increment = self.hdr_txt_h
        self.hdr_min_rh = self.hdr_txt_h + 5
        if self.default_hh[0] != "pixels":
            self.default_hh = (self.default_hh[0] if self.default_hh[0] != "pixels" else "pixels",
                               self.GetHdrLinesHeight(int(self.default_hh[0])) if self.default_hh[0] != "pixels" else self.default_hh[1])
        self.set_min_cw()
        self.CH.set_height(self.default_hh[1])
        
    def set_index_fnt_help(self):
        pass

    def data_reference(self, newdataref = None, reset_col_positions = True, reset_row_positions = True, redraw = False, return_id = True):
        if isinstance(newdataref, (list, tuple)):
            self.data = newdataref
            self.undo_storage = deque(maxlen = self.max_undos)
            if reset_col_positions:
                self.reset_col_positions()
            if reset_row_positions:
                self.reset_row_positions()
            if redraw:
                self.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
        if return_id:
            return id(self.data)
        else:
            return self.data

    def set_cell_size_to_text(self, r, c, only_set_if_too_small = False, redraw = True, run_binding = False):
        min_cw = self.min_cw
        min_rh = self.min_rh
        h = int(min_rh)
        w = int(min_cw)
        if self.all_columns_displayed:
            cn = int(c)
        else:
            cn = self.displayed_columns[c]
        rn = int(r)
        if (rn, cn) in self.cell_options and 'checkbox' in self.cell_options[(rn, cn)]:
            self.txt_measure_canvas.itemconfig(self.txt_measure_canvas_text, 
                                               text = self.cell_options[(rn, cn)]['checkbox']['text'],
                                               font = self._hdr_font)
            b = self.txt_measure_canvas.bbox(self.txt_measure_canvas_text)
            tw = b[2] - b[0] + 7 + self.txt_h
            if b[3] - b[1] + 5 > h:
                h = b[3] - b[1] + 5
        else:
            try:
                if isinstance(self.data[r][cn], str):
                    txt = self.data[r][cn]
                else:
                    txt = f"{self.data[r][cn]}"
            except:
                txt = ""
            if txt:
                self.txt_measure_canvas.itemconfig(self.txt_measure_canvas_text, text = txt, font = self._font)
                b = self.txt_measure_canvas.bbox(self.txt_measure_canvas_text)
                tw = b[2] - b[0] + self.txt_h + 7 if (rn, cn) in self.cell_options and 'dropdown' in self.cell_options[(rn, cn)] else b[2] - b[0] + 7
                if b[3] - b[1] + 5 > h:
                    h = b[3] - b[1] + 5
            else:
                if (rn, cn) in self.cell_options and 'dropdown' in self.cell_options[(rn, cn)]:
                    tw = self.txt_h + 7
                else:
                    tw = min_cw
        if tw > w:
            w = tw
        if h < min_rh:
            h = int(min_rh)
        elif h > self.RI.max_rh:
            h = int(self.RI.max_rh)
        if w < min_cw:
            w = int(min_cw)
        elif w > self.CH.max_cw:
            w = int(self.CH.max_cw)
        cell_needs_resize_w = False
        cell_needs_resize_h = False
        if only_set_if_too_small:
            if w > self.col_positions[c + 1] - self.col_positions[c]:
                cell_needs_resize_w = True
            if h > self.row_positions[r + 1] - self.row_positions[r]:
                cell_needs_resize_h = True
        else:
            if w != self.col_positions[c + 1] - self.col_positions[c]:
                cell_needs_resize_w = True
            if h != self.row_positions[r + 1] - self.row_positions[r]:
                cell_needs_resize_h = True
        if cell_needs_resize_w:
            old_width = self.col_positions[c + 1] - self.col_positions[c]
            new_col_pos = self.col_positions[c] + w
            increment = new_col_pos - self.col_positions[c + 1]
            self.col_positions[c + 2:] = [e + increment for e in islice(self.col_positions, c + 2, len(self.col_positions))]
            self.col_positions[c + 1] = new_col_pos
            new_width = self.col_positions[c + 1] - self.col_positions[c]
            if run_binding and self.CH.column_width_resize_func is not None and old_width != new_width:
                self.CH.column_width_resize_func(ResizeEvent("column_width_resize", c, old_width, new_width))
        if cell_needs_resize_h:
            old_height = self.row_positions[r + 1] - self.row_positions[r]
            new_row_pos = self.row_positions[r] + h
            increment = new_row_pos - self.row_positions[r + 1]
            self.row_positions[r + 2:] = [e + increment for e in islice(self.row_positions, r + 2, len(self.row_positions))]
            self.row_positions[r + 1] = new_row_pos
            new_height = self.row_positions[r + 1] - self.row_positions[r]
            if run_binding and self.RI.row_height_resize_func is not None and old_height != new_height:
                self.RI.row_height_resize_func(ResizeEvent("row_height_resize", r, old_height, new_height))
        if cell_needs_resize_w or cell_needs_resize_h:
            self.recreate_all_selection_boxes()
            if redraw:
                self.refresh()
                return True
            else:
                return False

    def set_all_cell_sizes_to_text(self, include_index = False):
        min_cw = self.min_cw
        min_rh = self.min_rh
        rhs = defaultdict(lambda: int(min_rh))
        cws = []
        x = self.txt_measure_canvas.create_text(0, 0, text = "", font = self._font)
        x2 = self.txt_measure_canvas.create_text(0, 0, text = "", font = self._hdr_font)
        itmcon = self.txt_measure_canvas.itemconfig
        itmbbx = self.txt_measure_canvas.bbox
        if self.all_columns_displayed:
            iterable = range(self.total_data_cols())
        else:
            iterable = self.displayed_columns
        if isinstance(self._row_index, list):
            for rn in range(self.total_data_rows()):
                if rn in self.RI.cell_options and 'checkbox' in self.RI.cell_options[rn]:
                    txt = self.RI.cell_options[rn]['checkbox']['text']
                else:
                    try:
                        if isinstance(self._row_index[rn], str):
                            txt = self._row_index[rn]
                        else:
                            txt = f"{self._row_index[rn]}"
                    except:
                        txt = ""
                if txt:
                    itmcon(x, text = txt)
                    b = itmbbx(x)
                    h = b[3] - b[1] + 7
                else:
                    h = min_rh
                if h < min_rh:
                    h = int(min_rh)
                elif h > self.RI.max_rh:
                    h = int(self.RI.max_rh)
                if h > rhs[rn]:
                    rhs[rn] = h
        for cn in iterable:
            if cn in self.CH.cell_options and 'checkbox' in self.CH.cell_options[cn]:
                txt = self.CH.cell_options[cn]['checkbox']['text']
                if txt:
                    itmcon(x2, text = txt)
                    b = itmbbx(x2)
                    w = b[2] - b[0] + 7 + self.hdr_txt_h
                else:
                    w = self.min_cw
            else:
                try:
                    if isinstance(self._headers, int):
                        txt = self.data[self._headers][cn]
                    else:
                        txt = self._headers[cn]
                    if txt:
                        itmcon(x2, text = txt)
                        b = itmbbx(x2)
                        w = b[2] - b[0] + self.hdr_txt_h + 7 if cn in self.CH.cell_options and 'dropdown' in self.CH.cell_options[cn] else b[2] - b[0] + 7
                    else:
                        w = self.min_cw + self.hdr_txt_h + 7 if cn in self.CH.cell_options and 'dropdown' in self.CH.cell_options[cn] else self.min_cw
                except:
                    if self.CH.default_hdr == "letters":
                        itmcon(x2, text = f"{num2alpha(cn)}")
                    elif self.CH.default_hdr == "numbers":
                        itmcon(x2, text = f"{cn + 1}")
                    else:
                        itmcon(x2, text = f"{cn + 1} {num2alpha(cn)}")
                    b = itmbbx(x2)
                    w = b[2] - b[0] + 7
            for rn, r in enumerate(self.data):
                if (rn, cn) in self.cell_options and 'checkbox' in self.cell_options[(rn, cn)]:
                    txt = self.cell_options[(rn, cn)]['checkbox']['text']
                    if txt:
                        itmcon(x, text = txt)
                        b = itmbbx(x)
                        tw = b[2] - b[0] + 7
                        h = b[3] - b[1] + 5
                    else:
                        tw = min_cw
                        h = min_rh
                else:
                    try:
                        if isinstance(r[cn], str):
                            txt = r[cn]
                        else:
                            txt = f"{r[cn]}"
                    except:
                        txt = ""
                    if txt:
                        itmcon(x, text = txt)
                        b = itmbbx(x)
                        tw = b[2] - b[0] + self.txt_h + 7 if (rn, cn) in self.cell_options and 'dropdown' in self.cell_options[(rn, cn)] else b[2] - b[0] + 7
                        h = b[3] - b[1] + 5
                    else:
                        tw = self.txt_h + 7 if (rn, cn) in self.cell_options and 'dropdown' in self.cell_options[(rn, cn)] else min_cw
                        h = min_rh
                if tw > w:
                    w = tw
                if h < min_rh:
                    h = int(min_rh)
                elif h > self.RI.max_rh:
                    h = int(self.RI.max_rh)
                if h > rhs[rn]:
                    rhs[rn] = h
            if w < min_cw:
                w = int(min_cw)
            elif w > self.CH.max_cw:
                w = int(self.CH.max_cw)
            cws.append(w)
        self.txt_measure_canvas.delete(x)
        self.txt_measure_canvas.delete(x2)
        self.row_positions = list(accumulate(chain([0], (height for height in rhs.values()))))
        self.col_positions = list(accumulate(chain([0], (width for width in cws))))
        self.recreate_all_selection_boxes()
        return self.row_positions, self.col_positions

    def reset_col_positions(self, ncols = None):
        colpos = int(self.default_cw)
        if self.all_columns_displayed:
            self.col_positions = list(accumulate(chain([0], (colpos for c in range(ncols if ncols is not None else self.total_data_cols())))))
        else:
            self.col_positions = list(accumulate(chain([0], (colpos for c in range(ncols if ncols is not None else len(self.displayed_columns))))))

    def del_col_position(self, idx, deselect_all = False):
        if deselect_all:
            self.deselect("all", redraw = False)
        if idx == "end" or len(self.col_positions) <= idx + 1:
            del self.col_positions[-1]
        else:
            w = self.col_positions[idx + 1] - self.col_positions[idx]
            idx += 1
            del self.col_positions[idx]
            self.col_positions[idx:] = [e - w for e in islice(self.col_positions, idx, len(self.col_positions))]

    def del_col_positions(self, idx, num = 1, deselect_all = False):
        if deselect_all:
            self.deselect("all", redraw = False)
        if idx == "end" or len(self.col_positions) <= idx + 1:
            del self.col_positions[-1]
        else:
            cws = [int(b - a) for a, b in zip(self.col_positions, islice(self.col_positions, 1, len(self.col_positions)))]
            cws[idx:idx + num] = []
            self.col_positions = list(accumulate(chain([0], (width for width in cws))))

    def insert_col_position(self, idx = "end", width = None, deselect_all = False):
        if deselect_all:
            self.deselect("all", redraw = False)
        if width is None:
            w = self.default_cw
        else:
            w = width
        if idx == "end" or len(self.col_positions) == idx + 1:
            self.col_positions.append(self.col_positions[-1] + w)
        else:
            idx += 1
            self.col_positions.insert(idx, self.col_positions[idx - 1] + w)
            idx += 1
            self.col_positions[idx:] = [e + w for e in islice(self.col_positions, idx, len(self.col_positions))]

    def insert_col_positions(self, idx = "end", widths = None, deselect_all = False):
        if deselect_all:
            self.deselect("all", redraw = False)
        if widths is None:
            w = [self.default_cw]
        elif isinstance(widths, int):
            w = list(repeat(self.default_cw, widths))
        else:
            w = widths
        if idx == "end" or len(self.col_positions) == idx + 1:
            if len(w) > 1:
                self.col_positions += list(accumulate(chain([self.col_positions[-1] + w[0]], islice(w, 1, None))))
            else:
                self.col_positions.append(self.col_positions[-1] + w[0])
        else:
            if len(w) > 1:
                idx += 1
                self.col_positions[idx:idx] = list(accumulate(chain([self.col_positions[idx - 1] + w[0]], islice(w, 1, None))))
                idx += len(w)
                sumw = sum(w)
                self.col_positions[idx:] = [e + sumw for e in islice(self.col_positions, idx, len(self.col_positions))]
            else:
                w = w[0]
                idx += 1
                self.col_positions.insert(idx, self.col_positions[idx - 1] + w)
                idx += 1
                self.col_positions[idx:] = [e + w for e in islice(self.col_positions, idx, len(self.col_positions))]

    def insert_col_rc(self, event = None):
        if self.anything_selected(exclude_rows = True, exclude_cells = True):
            selcols = self.get_selected_cols()
            numcols = len(selcols)
            displayed_ins_col = min(selcols) if event == "left" else max(selcols) + 1
            if self.all_columns_displayed:
                data_ins_col = int(displayed_ins_col)
            else:
                if displayed_ins_col == len(self.col_positions) - 1:
                    rowlen = len(max(self.data, key = len)) if self.data else 0
                    data_ins_col = rowlen
                else:
                    try:
                        data_ins_col = int(self.displayed_columns[displayed_ins_col])
                    except:
                        data_ins_col = int(self.displayed_columns[displayed_ins_col - 1])
        else:
            numcols = 1
            displayed_ins_col = len(self.col_positions) - 1
            data_ins_col = int(displayed_ins_col)
        if isinstance(self.paste_insert_column_limit, int) and self.paste_insert_column_limit < displayed_ins_col + numcols:
            numcols = self.paste_insert_column_limit - len(self.col_positions) - 1
            if numcols < 1:
                return
        if self.extra_begin_insert_cols_rc_func is not None:
            try:
                self.extra_begin_insert_cols_rc_func(InsertEvent("begin_insert_columns", data_ins_col, displayed_ins_col, numcols))
            except:
                return
        saved_displayed_columns = list(self.displayed_columns)
        if not self.all_columns_displayed:
            if displayed_ins_col == len(self.col_positions) - 1:
                self.displayed_columns += list(range(rowlen, rowlen + numcols))
            else:
                if displayed_ins_col > len(self.displayed_columns) - 1:
                    adj_ins = displayed_ins_col - 1
                else:
                    adj_ins = displayed_ins_col
                part1 = self.displayed_columns[:adj_ins]
                part2 = list(range(self.displayed_columns[adj_ins], self.displayed_columns[adj_ins] + numcols + 1))
                part3 = [] if displayed_ins_col > len(self.displayed_columns) - 1 else [cn + numcols for cn in islice(self.displayed_columns, adj_ins + 1, None)]
                self.displayed_columns = (part1 +
                                          part2 +
                                          part3)
        self.insert_col_positions(idx = displayed_ins_col,
                                  widths = numcols,
                                  deselect_all = True)
        self.cell_options = {(rn, cn if cn < data_ins_col else cn + numcols): t2 for (rn, cn), t2 in self.cell_options.items()}
        self.col_options = {cn if cn < data_ins_col else cn + numcols: t for cn, t in self.col_options.items()}
        self.CH.cell_options = {cn if cn < data_ins_col else cn + numcols: t for cn, t in self.CH.cell_options.items()}
        if self._headers and isinstance(self._headers, list):
            try:
                self._headers[data_ins_col:data_ins_col] = list(repeat("", numcols))
            except:
                pass
        if self.row_positions == [0] and not self.data:
            self.insert_row_position(idx = "end",
                                     height = int(self.min_rh),
                                     deselect_all = False)
            self.data.append(list(repeat("", numcols)))
        else:
            for rn in range(len(self.data)):
                self.data[rn][data_ins_col:data_ins_col] = list(repeat("", numcols))
        self.create_selected(0, displayed_ins_col, len(self.row_positions) - 1, displayed_ins_col + numcols, "columns")
        self.create_current(0, displayed_ins_col, "column", inside = True)
        if self.undo_enabled:
            self.undo_storage.append(zlib.compress(pickle.dumps(("insert_col", {"data_col_num": data_ins_col,
                                                                                "displayed_columns": saved_displayed_columns,
                                                                                "sheet_col_num": displayed_ins_col,
                                                                                "numcols": numcols}))))
        self.refresh()
        if self.extra_end_insert_cols_rc_func is not None:
            self.extra_end_insert_cols_rc_func(InsertEvent("end_insert_columns", data_ins_col, displayed_ins_col, numcols))

    def insert_row_rc(self, event = None):
        if self.anything_selected(exclude_columns = True, exclude_cells = True):
            selrows = self.get_selected_rows()
            numrows = len(selrows)
            stidx = min(selrows) if event == "above" else max(selrows) + 1
            posidx = int(stidx)
        else:
            selrows = [0]
            numrows = 1
            stidx = self.total_data_rows()
            posidx = len(self.row_positions) - 1
        if isinstance(self.paste_insert_row_limit, int) and self.paste_insert_row_limit < posidx + numrows:
            numrows = self.paste_insert_row_limit - len(self.row_positions) - 1
            if numrows < 1:
                return
        if self.extra_begin_insert_rows_rc_func is not None:
            try:
                self.extra_begin_insert_rows_rc_func(InsertEvent("begin_insert_rows", stidx, posidx, numrows))
            except:
                return
        self.insert_row_positions(idx = posidx,
                                  heights = numrows,
                                  deselect_all = True)
        self.cell_options = {(rn if rn < posidx else rn + numrows, cn): t2 for (rn, cn), t2 in self.cell_options.items()}
        self.row_options = {rn if rn < posidx else rn + numrows: t for rn, t in self.row_options.items()}
        self.RI.cell_options = {rn if rn < posidx else rn + numrows: t for rn, t in self.RI.cell_options.items()}
        if self._row_index and isinstance(self._row_index, list):
            try:
                self._row_index[stidx:stidx] = list(repeat("", numrows))
            except:
                pass
        if self.col_positions == [0] and not self.data:
            self.insert_col_position(idx = "end",
                                     width = None,
                                     deselect_all = False)
            self.data.append([""])
        else:
            total_data_cols = self.total_data_cols()
            self.data[stidx:stidx] = [list(repeat("", total_data_cols)) for rn in range(numrows)]
        self.create_selected(posidx, 0, posidx + numrows, len(self.col_positions) - 1, "rows")
        self.create_current(posidx, 0, "row", inside = True)
        if self.undo_enabled:
            self.undo_storage.append(zlib.compress(pickle.dumps(("insert_row", {"data_row_num": stidx,
                                                                                "sheet_row_num": posidx,
                                                                                "numrows": numrows}))))
        self.refresh()
        if self.extra_end_insert_rows_rc_func is not None:
            self.extra_end_insert_rows_rc_func(InsertEvent("end_insert_rows", stidx, posidx, numrows))
            
    def del_cols_rc(self, event = None):
        seld_cols = sorted(self.get_selected_cols())
        if seld_cols:
            if self.extra_begin_del_cols_rc_func is not None:
                try:
                    self.extra_begin_del_cols_rc_func(DeleteRowColumnEvent("begin_delete_columns", seld_cols))
                except:
                    return
            seldset = set(seld_cols) if self.all_columns_displayed else set(self.displayed_columns[c] for c in seld_cols)
            list_of_coords = tuple((r, c) for (r, c) in self.cell_options if c in seldset)
            if self.undo_enabled:
                undo_storage = {'deleted_cols': {},
                                'colwidths': {},
                                'deleted_hdr_values': {},
                                'selection_boxes': self.get_boxes(),
                                'displayed_columns': list(self.displayed_columns),
                                'cell_options': {k: v.copy() for k, v in self.cell_options.items()},
                                'col_options': {k: v.copy() for k, v in self.col_options.items()},
                                'CH_cell_options': {k: v.copy() for k, v in self.CH.cell_options.items()}}
            if self.all_columns_displayed:
                if self.undo_enabled:
                    for c in reversed(seld_cols):
                        undo_storage['colwidths'][c] = self.col_positions[c + 1] - self.col_positions[c]
                        for rn in range(len(self.data)):
                            if c not in undo_storage['deleted_cols']:
                                undo_storage['deleted_cols'][c] = {}
                            try:
                                undo_storage['deleted_cols'][c][rn] = self.data[rn].pop(c)
                            except:
                                continue
                    if self._headers and isinstance(self._headers, list):
                        for c in reversed(seld_cols):
                            try:
                                undo_storage['deleted_hdr_values'][c] = self._headers.pop(c)
                            except:
                                continue
                else:
                    for rn in range(len(self.data)):
                        for c in reversed(seld_cols):
                            del self.data[rn][c]
                    if self._headers and isinstance(self._headers, list):
                        for c in reversed(seld_cols):
                            try:
                                del self._headers[c]
                            except:
                                continue
            else:
                if self.undo_enabled:
                    for c in reversed(seld_cols):
                        undo_storage['colwidths'][c] = self.col_positions[c + 1] - self.col_positions[c]
                        for rn in range(len(self.data)):
                            if self.displayed_columns[c] not in undo_storage['deleted_cols']:
                                undo_storage['deleted_cols'][self.displayed_columns[c]] = {}
                            try:
                                undo_storage['deleted_cols'][self.displayed_columns[c]][rn] = self.data[rn].pop(self.displayed_columns[c])
                            except:
                                continue
                    if self._headers and isinstance(self._headers, list):
                        for c in reversed(seld_cols):
                            try:
                                undo_storage['deleted_hdr_values'][self.displayed_columns[c]] = self._headers.pop(self.displayed_columns[c])
                            except:
                                continue
                else:
                    for rn in range(len(self.data)):
                        for c in reversed(seld_cols):
                            del self.data[rn][self.displayed_columns[c]]
                    if self._headers and isinstance(self._headers, list):
                        for c in reversed(seld_cols):
                            try:
                                del self._headers[self.displayed_columns[c]]
                            except:
                                continue
            if self.undo_enabled:
                self.undo_storage.append(("delete_cols", undo_storage))
            self.del_cell_options(list_of_coords)
            for c in reversed(seld_cols):
                dcol = c if self.all_columns_displayed else self.displayed_columns[c]
                self.del_col_position(c,
                                      deselect_all = False)
                if dcol in self.col_options:
                    del self.col_options[dcol]
                if dcol in self.CH.cell_options:
                    del self.CH.cell_options[dcol]
            numcols = len(seld_cols)
            idx = seld_cols[-1]
            self.cell_options = {(rn, cn if cn < idx else cn - numcols): t2 for (rn, cn), t2 in self.cell_options.items()}
            self.col_options = {cn if cn < idx else cn - numcols: t for cn, t in self.col_options.items()}
            self.CH.cell_options = {cn if cn < idx else cn - numcols: t for cn, t in self.CH.cell_options.items()}
            self.deselect("allcols", redraw = False)
            self.set_current_to_last()
            if not self.all_columns_displayed:
                self.displayed_columns = [c for c in self.displayed_columns if c not in seldset]
                for c in sorted(seldset):
                    self.displayed_columns = [dc if c > dc else dc - 1 for dc in self.displayed_columns]
            self.refresh()
            if self.extra_end_del_cols_rc_func is not None:
                self.extra_end_del_cols_rc_func(DeleteRowColumnEvent("end_delete_columns", seld_cols))

    def del_cell_options(self, list_of_coords):
        for r, dcol in list_of_coords:
            if (r, dcol) in self.cell_options and 'dropdown' in self.cell_options[(r, dcol)]:
                self.delete_dropdown(r, dcol)
            del self.cell_options[(r, dcol)]

    def del_rows_rc(self, event = None):
        seld_rows = sorted(self.get_selected_rows())
        if seld_rows:
            if self.extra_begin_del_rows_rc_func is not None:
                try:
                    self.extra_begin_del_rows_rc_func(DeleteRowColumnEvent("begin_delete_rows", seld_rows))
                except:
                    return
            seldset = set(seld_rows)
            list_of_coords = tuple((r, c) for (r, c) in self.cell_options if r in seldset)
            if self.undo_enabled:
                undo_storage = {'deleted_rows': [],
                                'deleted_index_values': [],
                                'selection_boxes': self.get_boxes(),
                                'cell_options': {k: v.copy() for k, v in self.cell_options.items()},
                                'row_options': {k: v.copy() for k, v in self.row_options.items()},
                                'RI_cell_options': {k: v.copy() for k, v in self.RI.cell_options.items()}}
                for r in reversed(seld_rows):
                    undo_storage['deleted_rows'].append((r, self.data.pop(r), self.row_positions[r + 1] - self.row_positions[r]))
            else:
                for r in reversed(seld_rows):
                    del self.data[r]
            if self._row_index and isinstance(self._row_index, list):
                if self.undo_enabled:
                    for r in reversed(seld_rows):
                        try:
                            undo_storage['deleted_index_values'].append((r, self._row_index.pop(r)))
                        except:
                            continue
                else:
                    for r in reversed(seld_rows):
                        try:
                            del self._row_index[r]
                        except:
                            continue
            if self.undo_enabled:
                self.undo_storage.append(("delete_rows", undo_storage))
            self.del_cell_options(list_of_coords)
            for r in reversed(seld_rows):
                self.del_row_position(r,
                                      deselect_all = False)
                if r in self.row_options:
                    del self.row_options[r]
                if r in self.RI.cell_options:
                    del self.RI.cell_options[r]
            numrows = len(seld_rows)
            idx = seld_rows[-1]
            self.cell_options = {(rn if rn < idx else rn - numrows, cn): t2 for (rn, cn), t2 in self.cell_options.items()}
            self.row_options = {rn if rn < idx else rn - numrows: t for rn, t in self.row_options.items()}
            self.RI.cell_options = {rn if rn < idx else rn - numrows: t for rn, t in self.RI.cell_options.items()}
            self.deselect("allrows", redraw = False)
            self.set_current_to_last()
            self.refresh()
            if self.extra_end_del_rows_rc_func is not None:
                self.extra_end_del_rows_rc_func(DeleteRowColumnEvent("end_delete_rows", seld_rows))

    def reset_row_positions(self):
        rowpos = self.default_rh[1]
        self.row_positions = list(accumulate(chain([0], (rowpos for r in range(self.total_data_rows())))))

    def del_row_position(self, idx, deselect_all = False):
        if deselect_all:
            self.deselect("all", redraw = False)
        if idx == "end" or len(self.row_positions) <= idx + 1:
            del self.row_positions[-1]
        else:
            w = self.row_positions[idx + 1] - self.row_positions[idx]
            idx += 1
            del self.row_positions[idx]
            self.row_positions[idx:] = [e - w for e in islice(self.row_positions, idx, len(self.row_positions))]

    def del_row_positions(self, idx, numrows = 1, deselect_all = False):
        if deselect_all:
            self.deselect("all", redraw = False)
        if idx == "end" or len(self.row_positions) <= idx + 1:
            del self.row_positions[-1]
        else:
            rhs = [int(b - a) for a, b in zip(self.row_positions, islice(self.row_positions, 1, len(self.row_positions)))]
            rhs[idx:idx + numrows] = []
            self.row_positions = list(accumulate(chain([0], (height for height in rhs))))

    def insert_row_position(self, idx, height = None, deselect_all = False):
        if deselect_all:
            self.deselect("all", redraw = False)
        if height is None:
            h = self.default_rh[1]
        else:
            h = height
        if idx == "end" or len(self.row_positions) == idx + 1:
            self.row_positions.append(self.row_positions[-1] + h)
        else:
            idx += 1
            self.row_positions.insert(idx, self.row_positions[idx - 1] + h)
            idx += 1
            self.row_positions[idx:] = [e + h for e in islice(self.row_positions, idx, len(self.row_positions))]

    def insert_row_positions(self, idx = "end", heights = None, deselect_all = False):
        if deselect_all:
            self.deselect("all", redraw = False)
        if heights is None:
            h = [self.default_rh[1]]
        elif isinstance(heights, int):
            h = list(repeat(self.default_rh[1], heights))
        else:
            h = heights
        if idx == "end" or len(self.row_positions) == idx + 1:
            if len(h) > 1:
                self.row_positions += list(accumulate(chain([self.row_positions[-1] + h[0]], islice(h, 1, None))))
            else:
                self.row_positions.append(self.row_positions[-1] + h[0])
        else:
            if len(h) > 1:
                idx += 1
                self.row_positions[idx:idx] = list(accumulate(chain([self.row_positions[idx - 1] + h[0]], islice(h, 1, None))))
                idx += len(h)
                sumh = sum(h)
                self.row_positions[idx:] = [e + sumh for e in islice(self.row_positions, idx, len(self.row_positions))]
            else:
                h = h[0]
                idx += 1
                self.row_positions.insert(idx, self.row_positions[idx - 1] + h)
                idx += 1
                self.row_positions[idx:] = [e + h for e in islice(self.row_positions, idx, len(self.row_positions))]

    def move_row_position(self, idx1, idx2):
        if not len(self.row_positions) <= 2:
            if idx1 < idx2:
                height = self.row_positions[idx1 + 1] - self.row_positions[idx1]
                self.row_positions.insert(idx2 + 1, self.row_positions.pop(idx1 + 1))
                for i in range(idx1 + 1, idx2 + 1):
                    self.row_positions[i] -= height
                self.row_positions[idx2 + 1] = self.row_positions[idx2] + height
            else:
                height = self.row_positions[idx1 + 1] - self.row_positions[idx1]
                self.row_positions.insert(idx2 + 1, self.row_positions.pop(idx1 + 1))
                for i in range(idx2 + 2, idx1 + 2):
                    self.row_positions[i] += height
                self.row_positions[idx2 + 1] = self.row_positions[idx2] + height

    def move_col_position(self, idx1, idx2):
        if not len(self.col_positions) <= 2:
            if idx1 < idx2:
                width = self.col_positions[idx1 + 1] - self.col_positions[idx1]
                self.col_positions.insert(idx2 + 1, self.col_positions.pop(idx1 + 1))
                for i in range(idx1 + 1, idx2 + 1):
                    self.col_positions[i] -= width
                self.col_positions[idx2 + 1] = self.col_positions[idx2] + width
            else:
                width = self.col_positions[idx1 + 1] - self.col_positions[idx1]
                self.col_positions.insert(idx2 + 1, self.col_positions.pop(idx1 + 1))
                for i in range(idx2 + 2, idx1 + 2):
                    self.col_positions[i] += width
                self.col_positions[idx2 + 1] = self.col_positions[idx2] + width

    def display_columns(self, columns = None, all_columns_displayed = None, reset_col_positions = True, deselect_all = True):
        if columns is None and all_columns_displayed is None:
            return list(range(self.total_data_cols())) if self.all_columns_displayed else self.displayed_columns
        total_data_cols = None
        if (
            (columns is not None and columns != self.displayed_columns) or 
            (all_columns_displayed and not self.all_columns_displayed)
            ):
            self.undo_storage = deque(maxlen = self.max_undos)
        if columns is not None and columns != self.displayed_columns:
            self.displayed_columns = sorted(columns)
        if all_columns_displayed:
            if not self.all_columns_displayed:
                total_data_cols = self.total_data_cols()
                self.displayed_columns = list(range(total_data_cols))
            self.all_columns_displayed = True
        elif all_columns_displayed is not None and not all_columns_displayed:
            self.all_columns_displayed = False
        if reset_col_positions:
            self.reset_col_positions(ncols = total_data_cols)
        if deselect_all:
            self.deselect("all", redraw = False)
                
    def headers(self, newheaders = None, index = None, reset_col_positions = False, show_headers_if_not_sheet = True, redraw = False):
        if newheaders is not None:
            if isinstance(newheaders, (list, tuple)):
                self._headers = list(newheaders) if isinstance(newheaders, tuple) else newheaders
            elif isinstance(newheaders, int):
                self._headers = int(newheaders)
            elif isinstance(self._headers, list) and isinstance(index, int):
                if len(self._headers) <= index:
                    self._headers.extend(list(repeat("", index - len(self._headers) + 1)))
                self._headers[index] = f"{newheaders}"
            elif not isinstance(newheaders, (list, tuple, int)) and index is None:
                try:
                    self._headers = list(newheaders)
                except:
                    raise ValueError("New header must be iterable or int (use int to use a row as the header")
            if reset_col_positions:
                self.reset_col_positions()
            elif show_headers_if_not_sheet and isinstance(self._headers, list) and (self.col_positions == [0] or not self.col_positions):
                colpos = int(self.default_cw)
                if self.all_columns_displayed:
                    self.col_positions = list(accumulate(chain([0], (colpos for c in range(len(self._headers))))))
                else:
                    self.col_positions = list(accumulate(chain([0], (colpos for c in range(len(self.displayed_columns))))))
            if redraw:
                self.refresh()
        else:
            if index is not None:
                if isinstance(index, int):
                    return self._headers[index]
            else:
                return self._headers

    def row_index(self, newindex = None, index = None, reset_row_positions = False, show_index_if_not_sheet = True, redraw = False):
        if newindex is not None:
            if not self._row_index and not isinstance(self._row_index, int):
                self.RI.set_width(self.RI.default_width, set_TL = True)
            if isinstance(newindex, (list, tuple)):
                self._row_index = list(newindex) if isinstance(newindex, tuple) else newindex
            elif isinstance(newindex, int):
                self._row_index = int(newindex)
            elif isinstance(index, int):
                if len(self._row_index) <= index:
                    self._row_index.extend(list(repeat("", index - len(self._row_index) + 1)))
                self._row_index[index] = f"{newindex}"
            elif not isinstance(newindex, (list, tuple, int)) and index is None:
                try:
                    self._row_index = list(newindex)
                except:
                    raise ValueError("New index must be iterable or int (use int to use a column as the index")
            if reset_row_positions:
                self.reset_row_positions()
            elif show_index_if_not_sheet and isinstance(self._row_index, list) and (self.row_positions == [0] or not self.row_positions):
                rowpos = self.default_rh[1]
                self.row_positions = list(accumulate(chain([0], (rowpos for c in range(len(self._row_index))))))
            if redraw:
                self.refresh()
        else:
            if index is not None:
                if isinstance(index, int):
                    return self._row_index[index]
            else:
                return self._row_index

    def total_data_cols(self, include_headers = True):
        h_total = 0
        d_total = 0
        if include_headers:
            if isinstance(self._headers, list):
                h_total = len(self._headers)
        try:
            d_total = len(max(self.data, key = len))
        except:
            pass
        return h_total if h_total > d_total else d_total

    def total_data_rows(self):
        i_total = 0
        d_total = 0
        if isinstance(self._row_index, list):
            i_total = len(self._row_index)
        d_total = len(self.data)
        return i_total if i_total > d_total else d_total

    def data_dimensions(self, total_rows = None, total_columns = None):
        if total_rows is None and total_columns is None:
            return self.total_data_rows(), self.total_data_cols()
        if total_rows is not None:
            if len(self.data) < total_rows:
                if total_columns is None:
                    total_data_cols = self.total_data_cols()
                    self.data.extend([list(repeat("", total_data_cols)) for r in range(total_rows - len(self.data))])
                else:
                    self.data.extend([list(repeat("", total_columns)) for r in range(total_rows - len(self.data))])
            else:
                self.data[total_rows:] = []
        if total_columns is not None:
            self.data[:] = [r[:total_columns] if len(r) > total_columns else r + list(repeat("", total_columns - len(r))) for r in self.data]

    def equalize_data_row_lengths(self, include_header = False):
        total_columns = self.total_data_cols()
        if include_header and total_columns > len(self._headers):
            self._headers[:] = self._headers + list(repeat("", total_columns - len(self._headers)))
        self.data[:] = [r + list(repeat("", total_columns - len(r))) if total_columns > len(r) else r for r in self.data]
        return total_columns

    def get_canvas_visible_area(self):
        return self.canvasx(0), self.canvasy(0), self.canvasx(self.winfo_width()), self.canvasy(self.winfo_height())

    def get_visible_rows(self, y1, y2):
        start_row = bisect.bisect_left(self.row_positions, y1)
        end_row = bisect.bisect_right(self.row_positions, y2)
        if not y2 >= self.row_positions[-1]:
            end_row += 1
        return start_row, end_row

    def get_visible_columns(self, x1, x2):
        start_col = bisect.bisect_left(self.col_positions, x1)
        end_col = bisect.bisect_right(self.col_positions, x2)
        if not x2 >= self.col_positions[-1]:
            end_col += 1
        return start_col, end_col

    def redraw_highlight_get_text_fg(self, r, c, fc, fr, sc, sr, c_2_, c_3_, c_4_, selected_cells, actual_selected_rows, actual_selected_cols, dcol, can_width):
        redrawn = False
        # ________________________ CELL IS HIGHLIGHTED AND IN SELECTED CELLS ________________________
        if (r, dcol) in self.cell_options and 'highlight' in self.cell_options[(r, dcol)] and (r, c) in selected_cells:
            tf = self.table_selected_cells_fg if self.cell_options[(r, dcol)]['highlight'][1] is None or self.display_selected_fg_over_highlights else self.cell_options[(r, dcol)]['highlight'][1]
            if self.cell_options[(r, dcol)]['highlight'][0] is not None:
                c_1 = self.cell_options[(r, dcol)]['highlight'][0] if self.cell_options[(r, dcol)]['highlight'][0].startswith("#") else Color_Map_[self.cell_options[(r, dcol)]['highlight'][0]]
                redrawn = self.redraw_highlight(fc + 1, fr + 1, sc, sr, fill = (f"#{int((int(c_1[1:3], 16) + c_2_[0]) / 2):02X}" +
                                                                      f"{int((int(c_1[3:5], 16) + c_2_[1]) / 2):02X}" +
                                                                      f"{int((int(c_1[5:], 16) + c_2_[2]) / 2):02X}"),
                                                outline = self.table_fg if (r, dcol) in self.cell_options and 'dropdown' in self.cell_options[(r, dcol)] and self.show_dropdown_borders else "", 
                                                tag = "hi")
            
        elif r in self.row_options and 'highlight' in self.row_options[r] and (r, c) in selected_cells:
            tf = self.table_selected_cells_fg if self.row_options[r]['highlight'][1] is None or self.display_selected_fg_over_highlights else self.row_options[r]['highlight'][1]
            if self.row_options[r]['highlight'][0] is not None:
                c_1 = self.row_options[r]['highlight'][0] if self.row_options[r]['highlight'][0].startswith("#") else Color_Map_[self.row_options[r]['highlight'][0]]
                redrawn = self.redraw_highlight(fc + 1, fr + 1, sc, sr, fill = (f"#{int((int(c_1[1:3], 16) + c_2_[0]) / 2):02X}" +
                                                                      f"{int((int(c_1[3:5], 16) + c_2_[1]) / 2):02X}" +
                                                                      f"{int((int(c_1[5:], 16) + c_2_[2]) / 2):02X}"),
                                                outline = self.table_fg if (r, dcol) in self.cell_options and 'dropdown' in self.cell_options[(r, dcol)] and self.show_dropdown_borders else "", 
                                                tag = "hi",
                                                can_width = can_width if self.row_options[r]['highlight'][2] else None)
            
        elif dcol in self.col_options and 'highlight' in self.col_options[dcol] and (r, c) in selected_cells:
            tf = self.table_selected_cells_fg if self.col_options[dcol]['highlight'][1] is None or self.display_selected_fg_over_highlights else self.col_options[dcol]['highlight'][1]
            if self.col_options[dcol]['highlight'][0] is not None:
                c_1 = self.col_options[dcol]['highlight'][0] if self.col_options[dcol]['highlight'][0].startswith("#") else Color_Map_[self.col_options[dcol]['highlight'][0]]
                redrawn = self.redraw_highlight(fc + 1, fr + 1, sc, sr, fill = (f"#{int((int(c_1[1:3], 16) + c_2_[0]) / 2):02X}" +
                                                                      f"{int((int(c_1[3:5], 16) + c_2_[1]) / 2):02X}" +
                                                                      f"{int((int(c_1[5:], 16) + c_2_[2]) / 2):02X}"),
                                                outline = self.table_fg if (r, dcol) in self.cell_options and 'dropdown' in self.cell_options[(r, dcol)] and self.show_dropdown_borders else "", 
                                                tag = "hi")
            
        # ________________________ CELL IS HIGHLIGHTED AND IN SELECTED ROWS ________________________
        elif (r, dcol) in self.cell_options and 'highlight' in self.cell_options[(r, dcol)] and r in actual_selected_rows:
            tf = self.table_selected_rows_fg if self.cell_options[(r, dcol)]['highlight'][1] is None or self.display_selected_fg_over_highlights else self.cell_options[(r, dcol)]['highlight'][1]
            if self.cell_options[(r, dcol)]['highlight'][0] is not None:
                c_1 = self.cell_options[(r, dcol)]['highlight'][0] if self.cell_options[(r, dcol)]['highlight'][0].startswith("#") else Color_Map_[self.cell_options[(r, dcol)]['highlight'][0]]
                redrawn = self.redraw_highlight(fc + 1, fr + 1, sc, sr, fill = (f"#{int((int(c_1[1:3], 16) + c_4_[0]) / 2):02X}" +
                                                                      f"{int((int(c_1[3:5], 16) + c_4_[1]) / 2):02X}" +
                                                                      f"{int((int(c_1[5:], 16) + c_4_[2]) / 2):02X}"),
                                                outline = self.table_fg if (r, dcol) in self.cell_options and 'dropdown' in self.cell_options[(r, dcol)] and self.show_dropdown_borders else "", 
                                                tag = "hi")
            
        elif r in self.row_options and 'highlight' in self.row_options[r] and r in actual_selected_rows:
            tf = self.table_selected_rows_fg if self.row_options[r]['highlight'][1] is None or self.display_selected_fg_over_highlights else self.row_options[r]['highlight'][1]
            if self.row_options[r]['highlight'][0] is not None:
                c_1 = self.row_options[r]['highlight'][0] if self.row_options[r]['highlight'][0].startswith("#") else Color_Map_[self.row_options[r]['highlight'][0]]
                redrawn = self.redraw_highlight(fc + 1, fr + 1, sc, sr, fill = (f"#{int((int(c_1[1:3], 16) + c_4_[0]) / 2):02X}" +
                                                                      f"{int((int(c_1[3:5], 16) + c_4_[1]) / 2):02X}" +
                                                                      f"{int((int(c_1[5:], 16) + c_4_[2]) / 2):02X}"),
                                                outline = self.table_fg if (r, dcol) in self.cell_options and 'dropdown' in self.cell_options[(r, dcol)] and self.show_dropdown_borders else "", 
                                                tag = "hi",
                                                can_width = can_width if self.row_options[r]['highlight'][2] else None)
            
        elif dcol in self.col_options and 'highlight' in self.col_options[dcol] and r in actual_selected_rows:
            tf = self.table_selected_rows_fg if self.col_options[dcol]['highlight'][1] is None or self.display_selected_fg_over_highlights else self.col_options[dcol]['highlight'][1]
            if self.col_options[dcol]['highlight'][0] is not None:
                c_1 = self.col_options[dcol]['highlight'][0] if self.col_options[dcol]['highlight'][0].startswith("#") else Color_Map_[self.col_options[dcol]['highlight'][0]]
                redrawn = self.redraw_highlight(fc + 1, fr + 1, sc, sr, fill = (f"#{int((int(c_1[1:3], 16) + c_4_[0]) / 2):02X}" +
                                                                      f"{int((int(c_1[3:5], 16) + c_4_[1]) / 2):02X}" +
                                                                      f"{int((int(c_1[5:], 16) + c_4_[2]) / 2):02X}"),
                                                outline = self.table_fg if (r, dcol) in self.cell_options and 'dropdown' in self.cell_options[(r, dcol)] and self.show_dropdown_borders else "",
                                                tag = "hi")
            
        # ________________________ CELL IS HIGHLIGHTED AND IN SELECTED COLUMNS ________________________
        elif (r, dcol) in self.cell_options and 'highlight' in self.cell_options[(r, dcol)] and c in actual_selected_cols:
            tf = self.table_selected_columns_fg if self.cell_options[(r, dcol)]['highlight'][1] is None or self.display_selected_fg_over_highlights else self.cell_options[(r, dcol)]['highlight'][1]
            if self.cell_options[(r, dcol)]['highlight'][0] is not None:
                c_1 = self.cell_options[(r, dcol)]['highlight'][0] if self.cell_options[(r, dcol)]['highlight'][0].startswith("#") else Color_Map_[self.cell_options[(r, dcol)]['highlight'][0]]
                redrawn = self.redraw_highlight(fc + 1, fr + 1, sc, sr, fill = (f"#{int((int(c_1[1:3], 16) + c_3_[0]) / 2):02X}" +
                                                                      f"{int((int(c_1[3:5], 16) + c_3_[1]) / 2):02X}" +
                                                                      f"{int((int(c_1[5:], 16) + c_3_[2]) / 2):02X}"),
                                                outline = self.table_fg if (r, dcol) in self.cell_options and 'dropdown' in self.cell_options[(r, dcol)] and self.show_dropdown_borders else "",
                                                tag = "hi")
            
        elif r in self.row_options and 'highlight' in self.row_options[r] and c in actual_selected_cols:
            tf = self.table_selected_columns_fg if self.row_options[r]['highlight'][1] is None or self.display_selected_fg_over_highlights else self.row_options[r]['highlight'][1]
            if self.row_options[r]['highlight'][0] is not None:
                c_1 = self.row_options[r]['highlight'][0] if self.row_options[r]['highlight'][0].startswith("#") else Color_Map_[self.row_options[r]['highlight'][0]]
                redrawn = self.redraw_highlight(fc + 1, fr + 1, sc, sr, fill = (f"#{int((int(c_1[1:3], 16) + c_3_[0]) / 2):02X}" +
                                                                      f"{int((int(c_1[3:5], 16) + c_3_[1]) / 2):02X}" +
                                                                      f"{int((int(c_1[5:], 16) + c_3_[2]) / 2):02X}"),
                                                outline = self.table_fg if (r, dcol) in self.cell_options and 'dropdown' in self.cell_options[(r, dcol)] and self.show_dropdown_borders else "", 
                                                tag = "hi",
                                                can_width = can_width if self.row_options[r]['highlight'][2] else None)
            
        elif dcol in self.col_options and 'highlight' in self.col_options[dcol] and c in actual_selected_cols:
            tf = self.table_selected_columns_fg if self.col_options[dcol]['highlight'][1] is None or self.display_selected_fg_over_highlights else self.col_options[dcol]['highlight'][1]
            if self.col_options[dcol]['highlight'][0] is not None:
                c_1 = self.col_options[dcol]['highlight'][0] if self.col_options[dcol]['highlight'][0].startswith("#") else Color_Map_[self.col_options[dcol]['highlight'][0]]
                redrawn = self.redraw_highlight(fc + 1, fr + 1, sc, sr, fill = (f"#{int((int(c_1[1:3], 16) + c_3_[0]) / 2):02X}" +
                                                                      f"{int((int(c_1[3:5], 16) + c_3_[1]) / 2):02X}" +
                                                                      f"{int((int(c_1[5:], 16) + c_3_[2]) / 2):02X}"),
                                                outline = self.table_fg if (r, dcol) in self.cell_options and 'dropdown' in self.cell_options[(r, dcol)] and self.show_dropdown_borders else "",
                                                tag = "hi")

        # ________________________ CELL IS HIGHLIGHTED AND NOT SELECTED ________________________
        elif (r, dcol) in self.cell_options and 'highlight' in self.cell_options[(r, dcol)] and (r, c) not in selected_cells and r not in actual_selected_rows and c not in actual_selected_cols:
            tf = self.table_fg if self.cell_options[(r, dcol)]['highlight'][1] is None else self.cell_options[(r, dcol)]['highlight'][1]
            if self.cell_options[(r, dcol)]['highlight'][0] is not None:
                redrawn = self.redraw_highlight(fc + 1, fr + 1, sc, sr, fill = self.cell_options[(r, dcol)]['highlight'][0],
                                                outline = self.table_fg if (r, dcol) in self.cell_options and 'dropdown' in self.cell_options[(r, dcol)] and self.show_dropdown_borders else "", 
                                                tag = "hi")
            
        elif r in self.row_options and 'highlight' in self.row_options[r] and (r, c) not in selected_cells and r not in actual_selected_rows and c not in actual_selected_cols:
            tf = self.table_fg if self.row_options[r]['highlight'][1] is None else self.row_options[r]['highlight'][1]
            if self.row_options[r]['highlight'][0] is not None:
                redrawn = self.redraw_highlight(fc + 1, fr + 1, sc, sr, fill = self.row_options[r]['highlight'][0],
                                                outline = self.table_fg if (r, dcol) in self.cell_options and 'dropdown' in self.cell_options[(r, dcol)] and self.show_dropdown_borders else "", 
                                                tag = "hi",
                                                can_width = can_width if self.row_options[r]['highlight'][2] else None)
            
        elif dcol in self.col_options and 'highlight' in self.col_options[dcol] and (r, c) not in selected_cells and r not in actual_selected_rows and c not in actual_selected_cols:
            tf = self.table_fg if self.col_options[dcol]['highlight'][1] is None else self.col_options[dcol]['highlight'][1]
            if self.col_options[dcol]['highlight'][0] is not None:
                redrawn = self.redraw_highlight(fc + 1, fr + 1, sc, sr, fill = self.col_options[dcol]['highlight'][0],
                                                outline = self.table_fg if (r, dcol) in self.cell_options and 'dropdown' in self.cell_options[(r, dcol)] and self.show_dropdown_borders else "",
                                                tag = "hi")
        
        # ________________________ CELL IS JUST SELECTED ________________________
        elif (r, c) in selected_cells:
            tf = self.table_selected_cells_fg
        elif r in actual_selected_rows:
            tf = self.table_selected_rows_fg
        elif c in actual_selected_cols:
            tf = self.table_selected_columns_fg

        # ________________________ CELL IS NOT SELECTED ________________________
        else:
            tf = self.table_fg
        return tf, redrawn

    def redraw_highlight(self, x1, y1, x2, y2, fill, outline, tag, can_width = None):
        config = (fill, outline)
        coords = (x1 - 1 if outline else x1,
                  y1 - 1 if outline else y1,
                  x2 if can_width is None else x2 + can_width, 
                  y2)

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
            iid, showing, option = self.create_rectangle(coords, fill = fill, outline = outline, tag = tag), 1, 4

        if option in (1, 3):
            self.coords(iid, coords)
        if option in (2, 3):
            if showing:
                self.itemconfig(iid, fill = fill, outline = outline)
            else:
                self.itemconfig(iid, fill = fill, outline = outline, tag = tag, state = "normal")
        
        if k is not None and not self.hidd_high[k]:
            del self.hidd_high[k]

        self.disp_high[config].add(DrawnItem(iid = iid, showing = 1))
        return True

    def redraw_dropdown(self, x1, y1, x2, y2, fill, outline, tag, draw_outline = True, draw_arrow = True, dd_is_open = False):
        if draw_outline and self.show_dropdown_borders:
            self.redraw_highlight(x1 + 1, y1 + 1, x2, y2, fill = "", outline = self.table_fg, tag = tag)
        if draw_arrow:
            topysub = floor(self.half_txt_h / 2)
            mid_y = y1 + floor(self.min_rh / 2)
            if mid_y + topysub + 1 >= y1 + self.txt_h - 1:
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
            tx1 = x2 - self.txt_h + 1
            tx2 = x2 - self.half_txt_h - 1
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
                    self.itemconfig(t, fill = fill)
                else:
                    self.itemconfig(t, fill = fill, tag = tag, state = "normal")
                self.lift(t)
            else:
                t = self.create_line(points, fill = fill, width = 2, capstyle = tk.ROUND, joinstyle = tk.ROUND, tag = tag)
            self.disp_dropdown[t] = True

    def get_checkbox_points(self, x1, y1, x2, y2, radius = 8):
        return [x1+radius, y1,
                x1+radius, y1,
                x2-radius, y1,
                x2-radius, y1,
                x2, y1,
                x2, y1+radius,
                x2, y1+radius,
                x2, y2-radius,
                x2, y2-radius,
                x2, y2,
                x2-radius, y2,
                x2-radius, y2,
                x1+radius, y2,
                x1+radius, y2,
                x1, y2,
                x1, y2-radius,
                x1, y2-radius,
                x1, y1+radius,
                x1, y1+radius,
                x1, y1]

    def redraw_checkbox(self, r, dcol, x1, y1, x2, y2, fill, outline, tag, draw_check = False):
        points = self.get_checkbox_points(x1, y1, x2, y2)
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
            x1 = x1 + 4
            y1 = y1 + 4
            x2 = x2 - 3
            y2 = y2 - 3
            points = self.get_checkbox_points(x1, y1, x2, y2, radius = 4)
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

    def main_table_redraw_grid_and_text(self, redraw_header = False, redraw_row_index = False, redraw_table = True):
        try:
            last_col_line_pos = self.col_positions[-1] + 1
            last_row_line_pos = self.row_positions[-1] + 1
            can_width = self.winfo_width()
            can_height = self.winfo_height()
            self.configure(scrollregion = (0,
                                           0,
                                           last_col_line_pos + self.empty_horizontal,
                                           last_row_line_pos + self.empty_vertical))
            if can_width >= last_col_line_pos + self.empty_horizontal and self.parentframe.xscroll_showing:
                self.parentframe.xscroll.grid_forget()
                self.parentframe.xscroll_showing = False
            elif can_width < last_col_line_pos + self.empty_horizontal and not self.parentframe.xscroll_showing and not self.parentframe.xscroll_disabled and can_height > 45:
                self.parentframe.xscroll.grid(row = 2, column = 0, columnspan = 2, sticky = "nswe")
                self.parentframe.xscroll_showing = True
            if can_height >= last_row_line_pos + self.empty_vertical and self.parentframe.yscroll_showing:
                self.parentframe.yscroll.grid_forget()
                self.parentframe.yscroll_showing = False
            elif can_height < last_row_line_pos + self.empty_vertical and not self.parentframe.yscroll_showing and not self.parentframe.yscroll_disabled and can_width > 45:
                self.parentframe.yscroll.grid(row = 0, column = 2, rowspan = 3, sticky = "nswe")
                self.parentframe.yscroll_showing = True
            scrollpos_bot = self.canvasy(can_height)
            end_row = bisect.bisect_right(self.row_positions, scrollpos_bot)
            if not scrollpos_bot >= self.row_positions[-1]:
                end_row += 1
            if redraw_row_index and self.show_index:
                self.RI.auto_set_index_width(end_row - 1)
            scrollpos_left = self.canvasx(0)
            scrollpos_top = self.canvasy(0)
            scrollpos_right = self.canvasx(can_width)
            start_row = bisect.bisect_left(self.row_positions, scrollpos_top)
            self.row_width_resize_bbox = (scrollpos_left, scrollpos_top, scrollpos_left + 2, scrollpos_bot)
            self.header_height_resize_bbox = (scrollpos_left + 6, scrollpos_top, scrollpos_right, scrollpos_top + 2)
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
            start_col = bisect.bisect_left(self.col_positions, scrollpos_left)
            end_col = bisect.bisect_right(self.col_positions, scrollpos_right)
            if not scrollpos_right >= self.col_positions[-1]:
                end_col += 1
            if last_col_line_pos > scrollpos_right:
                x_stop = scrollpos_right
            else:
                x_stop = last_col_line_pos
            if last_row_line_pos > scrollpos_bot:
                y_stop = scrollpos_bot
            else:
                y_stop = last_row_line_pos
            scrollpos_bot_add2 = scrollpos_bot + 2
            if self.show_horizontal_grid:
                self.grid_cyc = cycle(self.grid_cyctup)
                points = []
                if self.horizontal_grid_to_end_of_window:
                    x_grid_stop = scrollpos_right + can_width
                else:
                    if last_col_line_pos > scrollpos_right:
                        x_grid_stop = x_stop + 1
                    else:
                        x_grid_stop = x_stop - 1
                for r in range(start_row - 1, end_row):
                    draw_y = self.row_positions[r]
                    st_or_end = next(self.grid_cyc)
                    if st_or_end == "st":
                        points.extend([self.canvasx(0) - 1, draw_y, 
                                       x_grid_stop, draw_y,
                                       x_grid_stop, self.row_positions[r + 1] if len(self.row_positions) - 1 > r else draw_y])
                    elif st_or_end == "end":
                        points.extend([x_grid_stop, draw_y,
                                       self.canvasx(0) - 1, draw_y,
                                       self.canvasx(0) - 1, self.row_positions[r + 1] if len(self.row_positions) - 1 > r else draw_y])
                if points:
                    if self.hidd_grid:
                        t, sh = self.hidd_grid.popitem()
                        self.coords(t, points)
                        if sh:
                            self.itemconfig(t, fill = self.table_grid_fg, capstyle = tk.BUTT, joinstyle = tk.ROUND, width = 1)
                        else:
                            self.itemconfig(t, fill = self.table_grid_fg, capstyle = tk.BUTT, joinstyle = tk.ROUND, width = 1, state = "normal")
                        self.disp_grid[t] = True
                    else:
                        self.disp_grid[self.create_line(points, fill = self.table_grid_fg, capstyle = tk.BUTT, joinstyle = tk.ROUND, width = 1, tag = "g")] = True
            if self.show_vertical_grid:
                self.grid_cyc = cycle(self.grid_cyctup)
                points = []
                if self.vertical_grid_to_end_of_window:
                    y_grid_stop = scrollpos_bot + can_height
                else:
                    if last_row_line_pos > scrollpos_bot:
                        y_grid_stop = y_stop + 1
                    else:
                        y_grid_stop = y_stop - 1
                for c in range(start_col - 1, end_col):
                    draw_x = self.col_positions[c]
                    st_or_end = next(self.grid_cyc)
                    if st_or_end == "st":
                        points.extend([draw_x, scrollpos_top - 1,
                                       draw_x, y_grid_stop,
                                       self.col_positions[c + 1] if len(self.col_positions) - 1 > c else draw_x, y_grid_stop])
                    elif st_or_end == "end":
                        points.extend([draw_x, y_grid_stop,
                                       draw_x, scrollpos_top - 1,
                                       self.col_positions[c + 1] if len(self.col_positions) - 1 > c else draw_x, scrollpos_top - 1])
                if points:
                    if self.hidd_grid:
                        t, sh = self.hidd_grid.popitem()
                        self.coords(t, points)
                        if sh:
                            self.itemconfig(t, fill = self.table_grid_fg, capstyle = tk.BUTT, joinstyle = tk.ROUND, width = 1)
                        else:
                            self.itemconfig(t, fill = self.table_grid_fg, capstyle = tk.BUTT, joinstyle = tk.ROUND, width = 1, state = "normal")
                        self.disp_grid[t] = True
                    else:
                        self.disp_grid[self.create_line(points, fill = self.table_grid_fg, capstyle = tk.BUTT, joinstyle = tk.ROUND, width = 1, tag = "g")] = True
            if start_row > 0:
                start_row -= 1
            if start_col > 0:
                start_col -= 1
            end_row -= 1
            c_2 = self.table_selected_cells_bg if self.table_selected_cells_bg.startswith("#") else Color_Map_[self.table_selected_cells_bg]
            c_2_ = (int(c_2[1:3], 16), int(c_2[3:5], 16), int(c_2[5:], 16))
            c_3 = self.table_selected_columns_bg if self.table_selected_columns_bg.startswith("#") else Color_Map_[self.table_selected_columns_bg]
            c_3_ = (int(c_3[1:3], 16), int(c_3[3:5], 16), int(c_3[5:], 16))
            c_4 = self.table_selected_rows_bg if self.table_selected_rows_bg.startswith("#") else Color_Map_[self.table_selected_rows_bg]
            c_4_ = (int(c_4[1:3], 16), int(c_4[3:5], 16), int(c_4[5:], 16))
            rows_ = tuple(range(start_row, end_row))
            selected_cells, selected_rows, selected_cols, actual_selected_rows, actual_selected_cols = self.get_redraw_selections((start_row, start_col, end_row, end_col - 1))
            font = self._font
            if redraw_table:
                for c in range(start_col, end_col - 1):
                    for r in rows_:
                        rtopgridln = self.row_positions[r]
                        rbotgridln = self.row_positions[r + 1]
                        if rbotgridln - rtopgridln < self.txt_h:
                            continue
                        if rbotgridln > scrollpos_bot_add2:
                            rbotgridln = scrollpos_bot_add2
                        cleftgridln = self.col_positions[c]
                        crightgridln = self.col_positions[c + 1]
                        
                        if self.all_columns_displayed:
                            dcol = c
                        else:
                            dcol = self.displayed_columns[c]
                        
                        fill, dd_drawn = self.redraw_highlight_get_text_fg(r, c, cleftgridln, rtopgridln, crightgridln, rbotgridln, c_2_, c_3_, c_4_, selected_cells, actual_selected_rows, actual_selected_cols, dcol, can_width)
                            
                        if (r, dcol) in self.cell_options and 'align' in self.cell_options[(r, dcol)]:
                            align = self.cell_options[(r, dcol)]['align']
                        elif r in self.row_options and 'align' in self.row_options[r]:
                            align = self.row_options[r]['align']
                        elif dcol in self.col_options and 'align' in self.col_options[dcol]:
                            align = self.col_options[dcol]['align']
                        else:
                            align = self.align
                        
                        if align == "w":
                            draw_x = cleftgridln + 3
                            if (r, dcol) in self.cell_options and 'dropdown' in self.cell_options[(r, dcol)]:
                                mw = crightgridln - cleftgridln - self.txt_h - 2
                                self.redraw_dropdown(cleftgridln, rtopgridln, crightgridln, self.row_positions[r + 1], 
                                                    fill = fill, outline = fill, tag = f"dd_{r}_{c}", draw_outline = not dd_drawn, draw_arrow = mw >= 5,
                                                    dd_is_open = self.cell_options[(r, dcol)]['dropdown']['window'] != "no dropdown open")
                            else:
                                mw = crightgridln - cleftgridln - 1
                            

                        elif align == "e":
                            if (r, dcol) in self.cell_options and 'dropdown' in self.cell_options[(r, dcol)]:
                                mw = crightgridln - cleftgridln - self.txt_h - 2
                                draw_x = crightgridln - 5 - self.txt_h
                                self.redraw_dropdown(cleftgridln, rtopgridln, crightgridln, self.row_positions[r + 1], 
                                                    fill = fill, outline = fill, tag = f"dd_{r}_{c}", draw_outline = not dd_drawn, draw_arrow = mw >= 5,
                                                    dd_is_open = self.cell_options[(r, dcol)]['dropdown']['window'] != "no dropdown open")
                            else:
                                mw = crightgridln - cleftgridln - 1
                                draw_x = crightgridln - 3

                        elif align == "center":
                            stop = cleftgridln + 5
                            if (r, dcol) in self.cell_options and 'dropdown' in self.cell_options[(r, dcol)]:
                                mw = crightgridln - cleftgridln - self.txt_h - 2
                                draw_x = cleftgridln + ceil((crightgridln - cleftgridln - self.txt_h) / 2)
                                self.redraw_dropdown(cleftgridln, rtopgridln, crightgridln, self.row_positions[r + 1], 
                                                    fill = fill, outline = fill, tag = f"dd_{r}_{c}", draw_outline = not dd_drawn, draw_arrow = mw >= 5,
                                                    dd_is_open = self.cell_options[(r, dcol)]['dropdown']['window'] != "no dropdown open")
                            else:
                                mw = crightgridln - cleftgridln - 1
                                draw_x = cleftgridln + floor((crightgridln - cleftgridln) / 2)

                        if (r, dcol) in self.cell_options and 'checkbox' in self.cell_options[(r, dcol)]:
                            if mw > self.txt_h + 2:
                                box_w = self.txt_h + 1
                                mw -= box_w
                                if align == "w":
                                    draw_x += box_w + 1
                                elif align == "center":
                                    draw_x += ceil(box_w / 2) + 1
                                    mw -= 1
                                else:
                                    mw -= 3
                                try:
                                    draw_check = self.data[r][dcol]
                                except:
                                    draw_check = False
                                self.redraw_checkbox(r,
                                                    dcol,
                                                    cleftgridln + 2,
                                                    rtopgridln + 2,
                                                    cleftgridln + self.txt_h + 3,
                                                    rtopgridln + self.txt_h + 3,
                                                    fill = fill if self.cell_options[(r, dcol)]['checkbox']['state'] == "normal" else self.table_grid_fg,
                                                    outline = "", tag = "cb", draw_check = draw_check)

                        if (r, dcol) in self.cell_options and 'checkbox' in self.cell_options[(r, dcol)]:
                            lns = self.cell_options[(r, dcol)]['checkbox']['text'].split("\n") if isinstance(self.cell_options[(r, dcol)]['checkbox']['text'], str) else f"{self.cell_options[(r, dcol)]['checkbox']['text']}".split("\n")
                        elif len(self.data) > r and len(self.data[r]) > dcol:
                            lns = self.data[r][dcol].split("\n") if isinstance(self.data[r][dcol], str) else f"{self.data[r][dcol]}".split("\n")
                        else:
                            continue
                        if lns != [''] and mw > self.txt_w and not ((align == "w" and draw_x > scrollpos_right) or
                                                                    (align == "e" and cleftgridln + 5 > scrollpos_right) or
                                                                    (align == "center" and stop > scrollpos_right)):
                            draw_y = rtopgridln + self.fl_ins
                            start_ln = int((scrollpos_top - rtopgridln) / self.xtra_lines_increment)
                            if start_ln < 0:
                                start_ln = 0
                            draw_y += start_ln * self.xtra_lines_increment
                            if draw_y + self.half_txt_h - 1 <= rbotgridln and len(lns) > start_ln:
                                for txt in islice(lns, start_ln, None):
                                    # for performance improvements in redrawing especially when just selecting cells
                                    # option 0: text doesn't need moving or config
                                    # option 1: text needs new x, y but has same config
                                    # option 2: text needs new config but has same x, y
                                    # option 3: text needs new x, y and new config
                                    # option 4: text needs to be created
                                    config = TextCfg(txt, fill, font, align)
                                    k = None
                                    if config in self.hidd_text:
                                        k = config
                                        iid, showing = self.hidd_text[k].pop()
                                        cc1, cc2 = self.coords(iid)
                                        if (int(cc1) == int(draw_x) and
                                            int(cc2) == int(draw_y)):
                                            option = 0 if showing else 2
                                        else:
                                            option = 1 if showing else 3
                                    elif self.hidd_text:
                                        k = next(iter(self.hidd_text))
                                        iid, showing = self.hidd_text[k].pop()
                                        cc1, cc2 = self.coords(iid)
                                        if (int(cc1) == int(draw_x) and
                                            int(cc2) == int(draw_y)):
                                            option = 2 if showing else 3
                                        else:
                                            option = 3
                                    else:
                                        iid, showing, option = self.create_text(draw_x, draw_y, text = txt, fill = fill, font = font, anchor = align, tag = "t"), 1, 4
                                    if option in (1, 3):
                                        self.coords(iid, draw_x, draw_y)
                                    if option in (2, 3):
                                        if showing:
                                            self.itemconfig(iid, text = txt, fill = fill, font = font, anchor = align)
                                        else:
                                            self.itemconfig(iid, text = txt, fill = fill, font = font, anchor = align, state = "normal")
                                    if k is not None and not self.hidd_text[k]:
                                        del self.hidd_text[k]
                                    wd = self.bbox(iid)
                                    wd = wd[2] - wd[0]
                                    if wd > mw:
                                        if align == "w":
                                            txt = txt[:int(len(txt) * (mw / wd))]
                                            self.itemconfig(iid, text = txt)
                                            wd = self.bbox(iid)
                                            while wd[2] - wd[0] > mw:
                                                txt = txt[:-1]
                                                self.itemconfig(iid, text = txt)
                                                wd = self.bbox(iid)
                                        elif align == "e":
                                            txt = txt[len(txt) - int(len(txt) * (mw / wd)):]
                                            self.itemconfig(iid, text = txt)
                                            wd = self.bbox(iid)
                                            while wd[2] - wd[0] > mw:
                                                txt = txt[1:]
                                                self.itemconfig(iid, text = txt)
                                                wd = self.bbox(iid)
                                        elif align == "center":
                                            self.c_align_cyc = cycle(self.centre_alignment_text_mod_indexes)
                                            tmod = ceil((len(txt) - int(len(txt) * (mw / wd))) / 2)
                                            txt = txt[tmod - 1:-tmod]
                                            self.itemconfig(iid, text = txt)
                                            wd = self.bbox(iid)
                                            while wd[2] - wd[0] > mw:
                                                txt = txt[next(self.c_align_cyc)]
                                                self.itemconfig(iid, text = txt)
                                                wd = self.bbox(iid)
                                            self.coords(iid, draw_x, draw_y)
                                        self.disp_text[config._replace(txt = txt)].add(DrawnItem(iid = iid, showing = 1))
                                    else:
                                        self.disp_text[config].add(DrawnItem(iid = iid, showing = 1))
                                    draw_y += self.xtra_lines_increment
                                    if draw_y + self.half_txt_h - 1 > rbotgridln:
                                        break
            if redraw_table:
                self.tag_raise("t")
                for cfg, set_ in self.hidd_text.items():
                    for namedtup in tuple(set_):
                        if namedtup.showing:
                            self.itemconfig(namedtup.iid, state = "hidden")
                            self.hidd_text[cfg].discard(namedtup)
                            self.hidd_text[cfg].add(namedtup._replace(showing = 0))
                for cfg, set_ in self.hidd_high.items():
                    for namedtup in tuple(set_):
                        if namedtup.showing:
                            self.itemconfig(namedtup.iid, state = "hidden")
                            self.hidd_high[cfg].discard(namedtup)
                            self.hidd_high[cfg].add(namedtup._replace(showing = 0))
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
                if self.show_selected_cells_border:
                    self.tag_raise("CellSelectBorder")
                    self.tag_raise("Current_Inside")
                    self.tag_raise("Current_Outside")
                    self.tag_raise("RowSelectBorder")
                    self.tag_raise("ColSelectBorder")
            if redraw_header and self.show_header:
                self.CH.redraw_grid_and_text(last_col_line_pos, scrollpos_left, x_stop, start_col, end_col, selected_cols, actual_selected_rows, actual_selected_cols)
            if redraw_row_index and self.show_index:
                self.RI.redraw_grid_and_text(last_row_line_pos, scrollpos_top, y_stop, start_row, end_row + 1, scrollpos_bot, selected_rows, actual_selected_cols, actual_selected_rows)
        except:
            return False
        return True

    def get_all_selection_items(self):
        return sorted(self.find_withtag("CellSelectFill") + self.find_withtag("RowSelectFill") + self.find_withtag("ColSelectFill") + self.find_withtag("Current_Inside") + self.find_withtag("Current_Outside"))

    def get_boxes(self):
        boxes = {}
        for item in self.get_all_selection_items():
            alltags = self.gettags(item)
            if alltags[0] == "CellSelectFill":
                boxes[tuple(int(e) for e in alltags[1].split("_") if e)] = "cells"
            elif alltags[0] == "RowSelectFill":
                boxes[tuple(int(e) for e in alltags[1].split("_") if e)] = "rows"
            elif alltags[0] == "ColSelectFill":
                boxes[tuple(int(e) for e in alltags[1].split("_") if e)] = "columns"
            elif alltags[0] == "Current_Inside":
                boxes[tuple(int(e) for e in alltags[1].split("_") if e)] = f"{alltags[2]}_inside"
            elif alltags[0] == "Current_Outside":
                boxes[tuple(int(e) for e in alltags[1].split("_") if e)] = f"{alltags[2]}_outside"
        return boxes

    def reselect_from_get_boxes(self, boxes):
        for k, v in boxes.items():
            if v == "cells":
                self.create_selected(k[0], k[1], k[2], k[3], "cells")
            elif v == "rows":
                self.create_selected(k[0], k[1], k[2], k[3], "rows")
            elif v == "columns":
                self.create_selected(k[0], k[1], k[2], k[3], "columns")
            elif v in ("cell_inside", "cell_outside", "row_inside", "row_outside", "col_outside", "col_inside"): #currently selected
                x = v.split("_")
                self.create_current(k[0], k[1], type_ = x[0], inside = True if x[1] == "inside" else False)

    def delete_selection_rects(self, cells = True, rows = True, cols = True, delete_current = True):
        deleted_boxes = {}
        if cells:
            for item in self.find_withtag("CellSelectFill"):
                alltags = self.gettags(item)
                if alltags:
                    deleted_boxes[tuple(int(e) for e in alltags[1].split("_") if e)] = "cells"
            self.delete("CellSelectFill", "CellSelectBorder")
            self.RI.delete("CellSelectFill", "CellSelectBorder")
            self.CH.delete("CellSelectFill", "CellSelectBorder")
        if rows:
            for item in self.find_withtag("RowSelectFill"):
                alltags = self.gettags(item)
                if alltags:
                    deleted_boxes[tuple(int(e) for e in alltags[1].split("_") if e)] = "rows"
            self.delete("RowSelectFill", "RowSelectBorder")
            self.RI.delete("RowSelectFill", "RowSelectBorder")
            self.CH.delete("RowSelectFill", "RowSelectBorder")
        if cols:
            for item in self.find_withtag("ColSelectFill"):
                alltags = self.gettags(item)
                if alltags:
                    deleted_boxes[tuple(int(e) for e in alltags[1].split("_") if e)] = "columns"
            self.delete("ColSelectFill", "ColSelectBorder")
            self.RI.delete("ColSelectFill", "ColSelectBorder")
            self.CH.delete("ColSelectFill", "ColSelectBorder")
        if delete_current:
            for item in chain(self.find_withtag("Current_Inside"), self.find_withtag("Current_Outside")):
                alltags = self.gettags(item)
                if alltags:
                    deleted_boxes[tuple(int(e) for e in alltags[1].split("_") if e)] = "cells"
            self.delete("Current_Inside", "Current_Outside")
            self.RI.delete("Current_Inside", "Current_Outside")
            self.CH.delete("Current_Inside", "Current_Outside")
        return deleted_boxes

    def currently_selected(self):
        items = self.find_withtag("Current_Inside") + self.find_withtag("Current_Outside")
        if not items:
            return tuple()
        alltags = self.gettags(items[0])
        box = tuple(int(e) for e in alltags[1].split("_") if e)
        return CurrentlySelectedClass(box[0], box[1], alltags[2])

    def get_tags_of_current(self):
        items = self.find_withtag("Current_Inside") + self.find_withtag("Current_Outside")
        if items:
            return self.gettags(items[0])
        else:
            return tuple()

    def create_current(self, r, c, type_ = "cell", inside = False): # cell, column or row
        r1, c1, r2, c2 = r, c, r + 1, c + 1
        self.delete("Current_Inside", "Current_Outside")
        self.RI.delete("Current_Inside", "Current_Outside")
        self.CH.delete("Current_Inside", "Current_Outside")
        if self.col_positions == [0]:
            c1 = 0
            c2 = 0
        if self.row_positions == [0]:
            r1 = 0
            r2 = 0
        if inside:
            tagr = ("Current_Inside", f"{r1}_{c1}_{r2}_{c2}", type_)
        else:
            tagr = ("Current_Outside", f"{r1}_{c1}_{r2}_{c2}", type_)
        if type_ == "cell":
            outline = self.table_selected_cells_border_fg
        elif type_ == "row":
            outline = self.table_selected_rows_border_fg
        elif type_ == "column":
            outline = self.table_selected_columns_border_fg
        if self.show_selected_cells_border:
            b = self.create_rectangle(self.col_positions[c1] + 1, self.row_positions[r1] + 1, self.col_positions[c2], self.row_positions[r2],
                                      fill = "",
                                      outline = outline,
                                      width = 2,
                                      tags = tagr)
            self.tag_raise(b)
        else:
            b = self.create_rectangle(self.col_positions[c1], self.row_positions[r1], self.col_positions[c2], self.row_positions[r2],
                                      fill = self.table_selected_cells_bg,
                                      outline = "",
                                      tags = tagr)
            self.tag_lower(b)
        ri = self.RI.create_rectangle(0, self.row_positions[r1], self.RI.current_width - 1, self.row_positions[r2],
                                      fill = self.RI.index_selected_cells_bg,
                                      outline = "",
                                      tags = tagr)
        ch = self.CH.create_rectangle(self.col_positions[c1], 0, self.col_positions[c2], self.CH.current_height - 1,
                                      fill = self.CH.header_selected_cells_bg,
                                      outline = "",
                                      tags = tagr)
        self.RI.tag_lower(ri)
        self.CH.tag_lower(ch)
        return b

    def set_current_to_last(self):
        if not self.currently_selected():
            items = sorted(self.find_withtag("CellSelectFill") + self.find_withtag("RowSelectFill") + self.find_withtag("ColSelectFill"))
            if items:
                last = self.gettags(items[-1])
                r1, c1, r2, c2 = tuple(int(e) for e in last[1].split("_") if e)
                if last[0] == "CellSelectFill":
                    return self.gettags(self.create_current(r1, c1, "cell", inside = True))
                elif last[0] == "RowSelectFill":
                    return self.gettags(self.create_current(r1, c1, "row", inside = True))
                elif last[0] == "ColSelectFill":
                    return self.gettags(self.create_current(r1, c1, "column", inside = True))
        return tuple()
   
    def delete_current(self):
        self.delete("Current_Inside", "Current_Outside")
        self.RI.delete("Current_Inside", "Current_Outside")
        self.CH.delete("Current_Inside", "Current_Outside")
            
    def create_selected(self, r1 = None, c1 = None, r2 = None, c2 = None, type_ = "cells", taglower = True):
        currently_selected = self.currently_selected()
        if currently_selected and currently_selected.type_ == "cell":
            if (currently_selected.row >= r1 and
                currently_selected.column >= c1 and
                currently_selected.row < r2 and
                currently_selected.column < c2):
                self.create_current(currently_selected.row, currently_selected.column, type_ = "cell", inside = True)
        if type_ == "cells":
            tagr = ("CellSelectFill", f"{r1}_{c1}_{r2}_{c2}")
            tagb = ("CellSelectBorder", f"{r1}_{c1}_{r2}_{c2}")
            taglower = "CellSelectFill"
            mt_bg = self.table_selected_cells_bg
            mt_border_col = self.table_selected_cells_border_fg
        elif type_ == "rows":
            tagr = ("RowSelectFill", f"{r1}_{c1}_{r2}_{c2}")
            tagb = ("RowSelectBorder", f"{r1}_{c1}_{r2}_{c2}")
            taglower = "RowSelectFill"
            mt_bg = self.table_selected_rows_bg
            mt_border_col = self.table_selected_rows_border_fg
        elif type_ == "columns":
            tagr = ("ColSelectFill", f"{r1}_{c1}_{r2}_{c2}")
            tagb = ("ColSelectBorder", f"{r1}_{c1}_{r2}_{c2}")
            taglower = "ColSelectFill"
            mt_bg = self.table_selected_columns_bg
            mt_border_col = self.table_selected_columns_border_fg
        self.last_selected = (r1, c1, r2, c2, type_)
        r = self.create_rectangle(self.col_positions[c1],
                                  self.row_positions[r1],
                                  self.canvasx(self.winfo_width()) if self.selected_rows_to_end_of_window else self.col_positions[c2],
                                  self.row_positions[r2],
                                  fill = mt_bg,
                                  outline = "",
                                  tags = tagr)
        
        self.RI.create_rectangle(0,
                                 self.row_positions[r1],
                                 self.RI.current_width - 1,
                                 self.row_positions[r2],
                                 fill = self.RI.index_selected_rows_bg if type_ == "rows" else self.RI.index_selected_cells_bg,
                                 outline = "",
                                 tags = tagr)
        self.CH.create_rectangle(self.col_positions[c1],
                                 0,
                                 self.col_positions[c2],
                                 self.CH.current_height - 1,
                                 fill = self.CH.header_selected_columns_bg if type_ == "columns" else self.CH.header_selected_cells_bg,
                                 outline = "",
                                 tags = tagr)
        if self.show_selected_cells_border and self.being_drawn_rect is None and self.RI.being_drawn_rect is None and self.CH.being_drawn_rect is None:
            b = self.create_rectangle(self.col_positions[c1], self.row_positions[r1], self.col_positions[c2], self.row_positions[r2],
                                      fill = "",
                                      outline = mt_border_col,
                                      tags = tagb)
        else:
            b = None
        if taglower:
            self.tag_lower(taglower)
            self.RI.tag_lower(taglower)
            self.CH.tag_lower(taglower)
            self.RI.tag_lower("Current_Inside")
            self.RI.tag_lower("Current_Outside")
            self.RI.tag_lower("CellSelectFill")
            self.CH.tag_lower("Current_Inside")
            self.CH.tag_lower("Current_Outside")
            self.CH.tag_lower("CellSelectFill")
        return r, b

    def recreate_all_selection_boxes(self):
        for item in chain(self.find_withtag("CellSelectFill"),
                          self.find_withtag("RowSelectFill"),
                          self.find_withtag("ColSelectFill"),
                          self.find_withtag("Current_Inside"),
                          self.find_withtag("Current_Outside")):
            full_tags = self.gettags(item)
            if full_tags:
                type_ = full_tags[0]
                r1, c1, r2, c2 = tuple(int(e) for e in full_tags[1].split("_") if e)
                self.delete(f"{r1}_{c1}_{r2}_{c2}")
                self.RI.delete(f"{r1}_{c1}_{r2}_{c2}")
                self.CH.delete(f"{r1}_{c1}_{r2}_{c2}")
                if r1 >= len(self.row_positions) - 1 or c1 >= len(self.col_positions) - 1:
                    continue
                if r2 > len(self.row_positions) - 1:
                    r2 = len(self.row_positions) - 1
                if c2 > len(self.col_positions) - 1:
                    c2 = len(self.col_positions) - 1
                if type_.startswith("CellSelect"):
                    self.create_selected(r1, c1, r2, c2, "cells")
                elif type_.startswith("RowSelect"):
                    self.create_selected(r1, c1, r2, c2, "rows")
                elif type_.startswith("ColSelect"):
                    self.create_selected(r1, c1, r2, c2, "columns")
                elif type_.startswith("Current"):
                    if type_ == "Current_Inside":
                        self.create_current(r1, c1, full_tags[2], inside = True)
                    elif type_ == "Current_Outside":
                        self.create_current(r1, c1, full_tags[2], inside = False)
        self.tag_lower("RowSelectFill")
        self.RI.tag_lower("RowSelectFill")
        self.CH.tag_lower("RowSelectFill")
        self.tag_lower("ColSelectFill")
        self.RI.tag_lower("ColSelectFill")
        self.CH.tag_lower("ColSelectFill")
        self.tag_lower("CellSelectFill")
        self.RI.tag_lower("CellSelectFill")
        self.CH.tag_lower("CellSelectFill")
        self.RI.tag_lower("Current_Inside")
        self.RI.tag_lower("Current_Outside")
        self.CH.tag_lower("Current_Inside")
        self.CH.tag_lower("Current_Outside")
        if not self.show_selected_cells_border:
            self.tag_lower("Current_Outside")

    def get_redraw_selections(self, within_range):
        scells = set()
        srows = set()
        scols = set()
        ac_srows = set()
        ac_scols = set()
        within_r1 = within_range[0]
        within_c1 = within_range[1]
        within_r2 = within_range[2]
        within_c2 = within_range[3]
        for item in self.find_withtag("RowSelectFill"):
            r1, c1, r2, c2 = tuple(int(e) for e in self.gettags(item)[1].split("_") if e)
            if (r1 >= within_r1 or
                r2 <= within_r2) or (within_r1 >= r1 and within_r2 <= r2):
                if r1 > within_r1:
                    start_row = r1
                else:
                    start_row = within_r1
                if r2 < within_r2:
                    end_row = r2
                else:
                    end_row = within_r2
                srows.update(set(range(start_row, end_row)))
                ac_srows.update(set(range(start_row, end_row)))
        for item in chain(self.find_withtag("Current_Outside"), self.find_withtag("Current_Inside")):
            r1, c1, r2, c2 = tuple(int(e) for e in self.gettags(item)[1].split("_") if e)
            if (r1 >= within_r1 or
                r2 <= within_r2):
                if r1 > within_r1:
                    start_row = r1
                else:
                    start_row = within_r1
                if r2 < within_r2:
                    end_row = r2
                else:
                    end_row = within_r2
                srows.update(set(range(start_row, end_row)))
        for item in self.find_withtag("ColSelectFill"): 
            r1, c1, r2, c2 = tuple(int(e) for e in self.gettags(item)[1].split("_") if e)
            if (c1 >= within_c1 or
                c2 <= within_c2) or (within_c1 >= c1 and within_c2 <= c2):
                if c1 > within_c1:
                    start_col = c1
                else:
                    start_col = within_c1
                if c2 < within_c2:
                    end_col = c2
                else:
                    end_col = within_c2
                scols.update(set(range(start_col, end_col)))
                ac_scols.update(set(range(start_col, end_col)))
        for item in self.find_withtag("Current_Outside"):
            r1, c1, r2, c2 = tuple(int(e) for e in self.gettags(item)[1].split("_") if e)
            if (c1 >= within_c1 or
                c2 <= within_c2):
                if c1 > within_c1:
                    start_col = c1
                else:
                    start_col = within_c1
                if c2 < within_c2:
                    end_col = c2
                else:
                    end_col = within_c2
                scols.update(set(range(start_col, end_col)))
        if not self.show_selected_cells_border:
            iterable = chain(self.find_withtag("CellSelectFill"), self.find_withtag("Current_Outside"))
        else:
            iterable = self.find_withtag("CellSelectFill")
        for item in iterable:
            tags = self.gettags(item)
            r1, c1, r2, c2 = tuple(int(e) for e in tags[1].split("_") if e)
            if (r1 >= within_r1 or
                c1 >= within_c1 or
                r2 <= within_r2 or
                c2 <= within_c2) or (within_c1 >= c1 and within_c2 <= c2) or (within_r1 >= r1 and within_r2 <= r2):
                if r1 > within_r1:
                    start_row = r1
                else:
                    start_row = within_r1
                if c1 > within_c1:
                    start_col = c1
                else:
                    start_col = within_c1
                if r2 < within_r2:
                    end_row = r2
                else:
                    end_row = within_r2
                if c2 < within_c2:
                    end_col = c2
                else:
                    end_col = within_c2
                colsr = tuple(range(start_col, end_col))
                rowsr = tuple(range(start_row, end_row))
                scells.update(set(product(rowsr, colsr)))
                srows.update(set(range(start_row, end_row)))
                scols.update(set(range(start_col, end_col)))
        return scells, srows, scols, ac_srows, ac_scols

    def get_selected_min_max(self):
        min_x = float("inf")
        min_y = float("inf")
        max_x = 0
        max_y = 0
        for item in chain(self.find_withtag("CellSelectFill"),
                          self.find_withtag("RowSelectFill"),
                          self.find_withtag("ColSelectFill"),
                          self.find_withtag("Current_Inside"),
                          self.find_withtag("Current_Outside")):
            r1, c1, r2, c2 = tuple(int(e) for e in self.gettags(item)[1].split("_") if e)
            if r1 < min_y:
                min_y = r1
            if c1 < min_x:
                min_x = c1
            if r2 > max_y:
                max_y = r2
            if c2 > max_x:
                max_x = c2
        if min_x != float("inf") and min_y != float("inf") and max_x > 0 and max_y > 0:
            return min_y, min_x, max_y, max_x
        else:
            return None, None, None, None

    def get_selected_rows(self, get_cells = False, within_range = None, get_cells_as_rows = False):
        s = set()
        if within_range is not None:
            within_r1 = within_range[0]
            within_r2 = within_range[1]
        if get_cells:
            if within_range is None:
                for item in self.find_withtag("RowSelectFill"):
                    r1, c1, r2, c2 = tuple(int(e) for e in self.gettags(item)[1].split("_") if e)
                    s.update(set(product(range(r1, r2), range(0, len(self.col_positions) - 1))))
                if get_cells_as_rows:
                    s.update(self.get_selected_cells())
            else:
                for item in self.find_withtag("RowSelectFill"):
                    r1, c1, r2, c2 = tuple(int(e) for e in self.gettags(item)[1].split("_") if e)
                    if (r1 >= within_r1 or
                        r2 <= within_r2):
                        if r1 > within_r1:
                            start_row = r1
                        else:
                            start_row = within_r1
                        if r2 < within_r2:
                            end_row = r2
                        else:
                            end_row = within_r2
                        s.update(set(product(range(start_row, end_row), range(0, len(self.col_positions) - 1))))
                if get_cells_as_rows:
                    s.update(self.get_selected_cells(within_range = (within_r1, 0, within_r2, len(self.col_positions) - 1)))
        else:
            if within_range is None:
                for item in self.find_withtag("RowSelectFill"):
                    r1, c1, r2, c2 = tuple(int(e) for e in self.gettags(item)[1].split("_") if e)
                    s.update(set(range(r1, r2)))
                if get_cells_as_rows:
                    s.update(set(tup[0] for tup in self.get_selected_cells()))
            else:
                for item in self.find_withtag("RowSelectFill"):
                    r1, c1, r2, c2 = tuple(int(e) for e in self.gettags(item)[1].split("_") if e)
                    if (r1 >= within_r1 or
                        r2 <= within_r2):
                        if r1 > within_r1:
                            start_row = r1
                        else:
                            start_row = within_r1
                        if r2 < within_r2:
                            end_row = r2
                        else:
                            end_row = within_r2
                        s.update(set(range(start_row, end_row)))
                if get_cells_as_rows:
                    s.update(set(tup[0] for tup in self.get_selected_cells(within_range = (within_r1, 0, within_r2, len(self.col_positions) - 1))))
        return s

    def get_selected_cols(self, get_cells = False, within_range = None, get_cells_as_cols = False):
        s = set()
        if within_range is not None:
            within_c1 = within_range[0]
            within_c2 = within_range[1]
        if get_cells:
            if within_range is None:
                for item in self.find_withtag("ColSelectFill"):
                    r1, c1, r2, c2 = tuple(int(e) for e in self.gettags(item)[1].split("_") if e)
                    s.update(set(product(range(c1, c2), range(0, len(self.row_positions) - 1))))
                if get_cells_as_cols:
                    s.update(self.get_selected_cells())
            else:
                for item in self.find_withtag("ColSelectFill"):
                    r1, c1, r2, c2 = tuple(int(e) for e in self.gettags(item)[1].split("_") if e)
                    if (c1 >= within_c1 or
                        c2 <= within_c2):
                        if c1 > within_c1:
                            start_col = c1
                        else:
                            start_col = within_c1
                        if c2 < within_c2:
                            end_col = c2
                        else:
                            end_col = within_c2
                        s.update(set(product(range(start_col, end_col), range(0, len(self.row_positions) - 1))))
                if get_cells_as_cols:
                    s.update(self.get_selected_cells(within_range = (0, within_c1, len(self.row_positions) - 1, within_c2)))
        else:
            if within_range is None:
                for item in self.find_withtag("ColSelectFill"):
                    r1, c1, r2, c2 = tuple(int(e) for e in self.gettags(item)[1].split("_") if e)
                    s.update(set(range(c1, c2)))
                if get_cells_as_cols:
                    s.update(set(tup[1] for tup in self.get_selected_cells()))
            else:
                for item in self.find_withtag("ColSelectFill"):
                    r1, c1, r2, c2 = tuple(int(e) for e in self.gettags(item)[1].split("_") if e)
                    if (c1 >= within_c1 or
                        c2 <= within_c2):
                        if c1 > within_c1:
                            start_col = c1
                        else:
                            start_col = within_c1
                        if c2 < within_c2:
                            end_col = c2
                        else:
                            end_col = within_c2
                        s.update(set(range(start_col, end_col)))
                if get_cells_as_cols:
                    s.update(set(tup[0] for tup in self.get_selected_cells(within_range = (0, within_c1, len(self.row_positions) - 1, within_c2))))
        return s

    def get_selected_cells(self, get_rows = False, get_cols = False, within_range = None):
        s = set()
        if within_range is not None:
            within_r1 = within_range[0]
            within_c1 = within_range[1]
            within_r2 = within_range[2]
            within_c2 = within_range[3]
        if get_cols and get_rows:
            iterable = chain(self.find_withtag("CellSelectFill"), self.find_withtag("RowSelectFill"), self.find_withtag("ColSelectFill"), self.find_withtag("Current_Outside"))
        elif get_rows and not get_cols:
            iterable = chain(self.find_withtag("CellSelectFill"), self.find_withtag("RowSelectFill"), self.find_withtag("Current_Outside"))
        elif get_cols and not get_rows:
            iterable = chain(self.find_withtag("CellSelectFill"), self.find_withtag("ColSelectFill"), self.find_withtag("Current_Outside"))
        else:
            iterable = chain(self.find_withtag("CellSelectFill"), self.find_withtag("Current_Outside"))
        if within_range is None:
            for item in iterable:
                r1, c1, r2, c2 = tuple(int(e) for e in self.gettags(item)[1].split("_") if e)
                s.update(set(product(range(r1, r2), range(c1, c2))))
        else:
            for item in iterable:
                r1, c1, r2, c2 = tuple(int(e) for e in self.gettags(item)[1].split("_") if e)
                if (r1 >= within_r1 or
                    c1 >= within_c1 or
                    r2 <= within_r2 or
                    c2 <= within_c2):
                    if r1 > within_r1:
                        start_row = r1
                    else:
                        start_row = within_r1
                    if c1 > within_c1:
                        start_col = c1
                    else:
                        start_col = within_c1
                    if r2 < within_r2:
                        end_row = r2
                    else:
                        end_row = within_r2
                    if c2 < within_c2:
                        end_col = c2
                    else:
                        end_col = within_c2
                    s.update(set(product(range(start_row, end_row), range(start_col, end_col))))
        return s

    def get_all_selection_boxes(self):
        return tuple(tuple(int(e) for e in self.gettags(item)[1].split("_") if e) for item in chain(self.find_withtag("CellSelectFill"),
                                                                                                    self.find_withtag("RowSelectFill"),
                                                                                                    self.find_withtag("ColSelectFill"),
                                                                                                    self.find_withtag("Current_Outside")))

    def get_all_selection_boxes_with_types(self):
        boxes = []
        for item in sorted(self.find_withtag("CellSelectFill") + self.find_withtag("RowSelectFill") + self.find_withtag("ColSelectFill") + self.find_withtag("Current_Outside")):
            tags = self.gettags(item)
            if tags:
                if tags[0].startswith(("Cell", "Current")):
                    boxes.append((tuple(int(e) for e in tags[1].split("_") if e), "cells"))
                elif tags[0].startswith("Row"):
                    boxes.append((tuple(int(e) for e in tags[1].split("_") if e), "rows"))
                elif tags[0].startswith("Col"):
                    boxes.append((tuple(int(e) for e in tags[1].split("_") if e), "columns"))
        return boxes

    def all_selected(self):
        for r1, c1, r2, c2 in self.get_all_selection_boxes():
            if not r1 and not c1 and r2 == len(self.row_positions) - 1 and c2 == len(self.col_positions) - 1:
                return True
        return False
    
    def cell_selected(self, r, c, inc_cols = False, inc_rows = False):
        if not isinstance(r, int) or not isinstance(c, int):
            return False
        if not inc_cols and not inc_rows:
            iterable = chain(self.find_withtag("CellSelectFill"), self.find_withtag("Current_Inside"), self.find_withtag("Current_Outside"))
        elif inc_cols and not inc_rows:
            iterable = chain(self.find_withtag("ColSelectFill"), self.find_withtag("CellSelectFill"), self.find_withtag("Current_Inside"), self.find_withtag("Current_Outside"))
        elif not inc_cols and inc_rows:
            iterable = chain(self.find_withtag("RowSelectFill"), self.find_withtag("CellSelectFill"), self.find_withtag("Current_Inside"), self.find_withtag("Current_Outside"))
        elif inc_cols and inc_rows:
            iterable = chain(self.find_withtag("RowSelectFill"), self.find_withtag("ColSelectFill"), self.find_withtag("CellSelectFill"), self.find_withtag("Current_Inside"), self.find_withtag("Current_Outside"))
        for item in iterable:
            r1, c1, r2, c2 = tuple(int(e) for e in self.gettags(item)[1].split("_") if e)
            if r1 <= r and c1 <= c and r2 > r and c2 > c:
                return True
        return False

    def col_selected(self, c):
        if not isinstance(c, int):
            return False
        for item in self.find_withtag("ColSelectFill"):
            r1, c1, r2, c2 = tuple(int(e) for e in self.gettags(item)[1].split("_") if e)
            if c1 <= c and c2 > c:
                return True
        return False

    def row_selected(self, r):
        if not isinstance(r, int):
            return False
        for item in self.find_withtag("RowSelectFill"):
            r1, c1, r2, c2 = tuple(int(e) for e in self.gettags(item)[1].split("_") if e)
            if r1 <= r and r2 > r:
                return True
        return False

    def anything_selected(self, exclude_columns = False, exclude_rows = False, exclude_cells = False):
        if exclude_columns and exclude_rows and not exclude_cells:
            if self.find_withtag("CellSelectFill") or self.find_withtag("Current_Outside"):
                return True
        elif exclude_columns and exclude_cells and not exclude_rows:
            if self.find_withtag("RowSelectFill"):
                return True
        elif exclude_rows and exclude_cells and not exclude_columns:
            if self.find_withtag("ColSelectFill"):
                return True
            
        elif exclude_columns and not exclude_rows and not exclude_cells:
            if self.find_withtag("CellSelectFill") or self.find_withtag("RowSelectFill") or self.find_withtag("Current_Outside"):
                return True
        elif exclude_rows and not exclude_columns and not exclude_cells:
            if self.find_withtag("CellSelectFill") or self.find_withtag("ColSelectFill") or self.find_withtag("Current_Outside"):
                return True
        elif exclude_cells and not exclude_columns and not exclude_rows:
            if self.find_withtag("RowSelectFill") or self.find_withtag("ColSelectFill"):
                return True
            
        elif not exclude_columns and not exclude_rows and not exclude_cells:
            if self.find_withtag("CellSelectFill") or self.find_withtag("RowSelectFill") or self.find_withtag("ColSelectFill") or self.find_withtag("Current_Outside"):
                return True
        return False

    def hide_current(self):
        for item in chain(self.find_withtag("Current_Inside"), self.find_withtag("Current_Outside")):
            self.itemconfig(item, state = "hidden")

    def show_current(self):
        for item in chain(self.find_withtag("Current_Inside"), self.find_withtag("Current_Outside")):
            self.itemconfig(item, state = "normal")

    def open_cell(self, event = None, ignore_existing_editor = False):
        if not self.anything_selected() or (not ignore_existing_editor and self.text_editor_id is not None):
            return
        currently_selected = self.currently_selected()
        if not currently_selected:
            return
        y1 = int(currently_selected[0])
        x1 = int(currently_selected[1])
        dcol = x1 if self.all_columns_displayed else self.displayed_columns[x1]
        if (
            ((y1, dcol) in self.cell_options and 'readonly' in self.cell_options[(y1, dcol)]) or
            (dcol in self.col_options and 'readonly' in self.col_options[dcol]) or
            (y1 in self.row_options and 'readonly' in self.row_options[y1])
            ):
            return
        elif (y1, dcol) in self.cell_options and ('dropdown' in self.cell_options[(y1, dcol)] or 'checkbox' in self.cell_options[(y1, dcol)]):
            if self.event_opens_dropdown_or_checkbox(event):
                if 'dropdown' in self.cell_options[(y1, dcol)]:
                    self.display_dropdown_window(y1, x1, event = event)
                elif 'checkbox' in self.cell_options[(y1, dcol)]:
                    self._click_checkbox(y1, x1, dcol)
        else:
            self.edit_cell_(event, r = y1, c = x1, dropdown = False)
            
    def event_opens_dropdown_or_checkbox(self, event = None):
        if event is None:
            return False
        elif event == "rc":
            return True
        elif ((hasattr(event, 'keysym') and event.keysym == 'Return') or  # enter or f2
              (hasattr(event, 'keysym') and event.keysym == 'F2') or
              (event is not None and hasattr(event, 'keycode') and event.keycode == "??" and hasattr(event, 'num') and event.num == 1) or
              (hasattr(event, 'keysym') and event.keysym == 'BackSpace')):
            return True
        else:
            return False

    # c is displayed col
    def edit_cell_(self, event = None, r = None, c = None, dropdown = False):
        text = None
        extra_func_key = "??"
        if event is None or self.event_opens_dropdown_or_checkbox(event):
            if event is not None:
                if hasattr(event, 'keysym') and event.keysym == 'Return':
                    extra_func_key = "Return"
                elif hasattr(event, 'keysym') and event.keysym == 'F2':
                    extra_func_key = "F2"
            text = f"{self.data[r][c]}" if self.all_columns_displayed else f"{self.data[r][self.displayed_columns[c]]}"
        elif event is not None and (hasattr(event, 'keysym') and event.keysym == 'BackSpace'):
            extra_func_key = "BackSpace"
            text = ""
        elif event is not None and ((hasattr(event, "char") and event.char.isalpha()) or
                                    (hasattr(event, "char") and event.char.isdigit()) or
                                    (hasattr(event, "char") and event.char in symbols_set)):
            extra_func_key = event.char
            text = event.char
        else:
            return False
        self.text_editor_loc = (r, c)
        if self.extra_begin_edit_cell_func is not None:
            try:
                text = self.extra_begin_edit_cell_func(EditCellEvent(r, c, extra_func_key, text, "begin_edit_cell"))
            except:
                return False
            if text is None:
                return False
            else:
                text = text if isinstance(text, str) else f"{text}"
        text = "" if text is None else text
        if self.cell_auto_resize_enabled:
            self.set_cell_size_to_text(r, c, only_set_if_too_small = True, redraw = True, run_binding = True)
        if not self.currently_selected():
            self.select_cell(r = r, c = c, keep_other_selections = True)
        self.create_text_editor(r = r, c = c, text = text, set_data_ref_on_destroy = True, dropdown = dropdown)
        return True
    
    # displayed indexes
    def get_cell_align(self, r, c):
        drow = r if self.all_rows_displayed else self.displayed_rows[r]
        dcol = c if self.all_columns_displayed else self.displayed_columns[c]
        if (drow, dcol) in self.cell_options and 'align' in self.cell_options[(drow, dcol)]:
            cell_alignment = self.cell_options[(drow, dcol)]['align']
        elif drow in self.row_options and 'align' in self.row_options[drow]:
            cell_alignment = self.row_options[drow]['align']
        elif dcol in self.col_options and 'align' in self.col_options[dcol]:
            cell_alignment = self.col_options[dcol]['align']
        else:
            cell_alignment = self.align
        return cell_alignment

    # displayed indexes
    def create_text_editor(self,
                           r = 0,
                           c = 0,
                           text = None,
                           state = "normal",
                           see = True,
                           set_data_ref_on_destroy = False,
                           binding = None,
                           dropdown = False):
        if (r, c) == self.text_editor_loc and self.text_editor is not None:
            self.text_editor.set_text(self.text_editor.get() + "" if not isinstance(text, str) else text)
            return
        if self.text_editor is not None:
            self.destroy_text_editor()
        if see:
            has_redrawn = self.see(r = r, c = c, check_cell_visibility = True)
            if not has_redrawn:
                self.refresh()
        self.text_editor_loc = (r, c)
        x = self.col_positions[c]
        y = self.row_positions[r]
        w = self.col_positions[c + 1] - x + 1
        h = self.row_positions[r + 1] - y + 1
        dcol = c if self.all_columns_displayed else self.displayed_columns[c]
        if text is None:
            text = self.data[r][dcol]
        self.hide_current()
        #bg, fg = self.get_widget_bg_fg(r, dcol)
        bg, fg = self.table_bg, self.table_fg
        self.text_editor = TextEditor(self, 
                                      text = text, 
                                      font = self._font, 
                                      state = state, 
                                      width = w, 
                                      height = h, 
                                      border_color = self.table_selected_cells_border_fg, 
                                      show_border = self.show_selected_cells_border,
                                      bg = bg, 
                                      fg = fg,
                                      popup_menu_font = self.popup_menu_font,
                                      popup_menu_fg = self.popup_menu_fg,
                                      popup_menu_bg = self.popup_menu_bg,
                                      popup_menu_highlight_bg = self.popup_menu_highlight_bg,
                                      popup_menu_highlight_fg = self.popup_menu_highlight_fg,
                                      binding = binding,
                                      align = self.get_cell_align(r, c),
                                      r = r,
                                      c = c,
                                      newline_binding = self.text_editor_newline_binding)
        self.text_editor.update_idletasks()
        self.text_editor_id = self.create_window((x, y), window = self.text_editor, anchor = "nw")
        if not dropdown:
            self.text_editor.textedit.focus_set()
            self.text_editor.scroll_to_bottom()
        self.text_editor.textedit.bind("<Alt-Return>", lambda x: self.text_editor_newline_binding(r, c))
        if USER_OS == 'darwin':
            self.text_editor.textedit.bind("<Option-Return>", lambda x: self.text_editor_newline_binding(r, c))
        for key, func in self.text_editor_user_bound_keys.items():
            self.text_editor.textedit.bind(key, func)
        if binding is not None:
            self.text_editor.textedit.bind("<Tab>", lambda x: binding((r, c, "Tab")))
            self.text_editor.textedit.bind("<Return>", lambda x: binding((r, c, "Return")))
            self.text_editor.textedit.bind("<FocusOut>", lambda x: binding((r, c, "FocusOut")))
            self.text_editor.textedit.bind("<Escape>", lambda x: binding((r, c, "Escape")))
        elif binding is None and set_data_ref_on_destroy:
            self.text_editor.textedit.bind("<Tab>", lambda x: self.close_text_editor((r, c, "Tab")))
            self.text_editor.textedit.bind("<Return>", lambda x: self.close_text_editor((r, c, "Return")))
            if not dropdown:
                self.text_editor.textedit.bind("<FocusOut>", lambda x: self.close_text_editor((r, c, "FocusOut")))
            self.text_editor.textedit.bind("<Escape>", lambda x: self.close_text_editor((r, c, "Escape")))
        else:
            self.text_editor.textedit.bind("<Escape>", lambda x: self.destroy_text_editor("Escape"))
    
    # displayed indexes
    def text_editor_newline_binding(self, r = 0, c = 0, event = None, check_lines = True):
        drow = r if self.all_rows_displayed else self.displayed_rows[r]
        dcol = c if self.all_columns_displayed else self.displayed_columns[c]
        curr_height = self.text_editor.winfo_height()
        if not check_lines or self.GetLinesHeight(self.text_editor.get_num_lines() + 1) > curr_height:
            new_height = curr_height + self.xtra_lines_increment
            space_bot = self.get_space_bot(r)
            if new_height > space_bot:
                new_height = space_bot
            if new_height != curr_height:
                self.text_editor.config(height = new_height)
                if ((r, dcol) in self.cell_options and
                    'dropdown' in self.cell_options[(r, dcol)]):
                    text_editor_h = self.text_editor.winfo_height()
                    win_h, anchor = self.get_dropdown_height_anchor(drow, dcol, text_editor_h)
                    if anchor == "nw":
                        self.coords(self.cell_options[(r, dcol)]['dropdown']['canvas_id'],
                                    self.col_positions[c], self.row_positions[r] + text_editor_h - 1)
                        self.itemconfig(self.cell_options[(r, dcol)]['dropdown']['canvas_id'],
                                        anchor = anchor, height = win_h)
                    elif anchor == "sw":
                        self.coords(self.cell_options[(r, dcol)]['dropdown']['canvas_id'],
                                    self.col_positions[c], self.row_positions[r])
                        self.itemconfig(self.cell_options[(r, dcol)]['dropdown']['canvas_id'],
                                        anchor = anchor, height = win_h)

    def destroy_text_editor(self, event = None):
        if event is not None and self.extra_end_edit_cell_func is not None and self.text_editor_loc is not None:
            self.extra_end_edit_cell_func(EditCellEvent(int(self.text_editor_loc[0]), int(self.text_editor_loc[1]), "Escape", None, "escape_edit_cell"))
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
        self.show_current()
        if event is not None and len(event) >= 3 and "Escape" in event:
            self.focus_set()

    # c is displayed col
    def close_text_editor(self, editor_info = None, r = None, c = None, set_data_ref_on_destroy = True, event = None, destroy = True, move_down = True, redraw = True, recreate = True):
        if self.focus_get() is None and editor_info:
            return "break"
        if editor_info is not None and len(editor_info) >= 3 and editor_info[2] == "Escape":
            self.destroy_text_editor("Escape")
            self.hide_dropdown_window(r, c)
            return "break"
        if self.text_editor is not None:
            self.text_editor_value = self.text_editor.get()
        if destroy:
            self.destroy_text_editor()
        if set_data_ref_on_destroy:
            if r is None and c is None and editor_info:
                r, c = editor_info[0], editor_info[1]
            if self.extra_end_edit_cell_func is None:
                self._set_cell_data(r, c, value = self.text_editor_value, redraw = False)
            elif self.extra_end_edit_cell_func is not None and not self.edit_cell_validation:
                self._set_cell_data(r, c, value = self.text_editor_value, redraw = False)
                self.extra_end_edit_cell_func(EditCellEvent(r, c, editor_info[2] if len(editor_info) >= 3 else "FocusOut", f"{self.text_editor_value}", "end_edit_cell"))
            elif self.extra_end_edit_cell_func is not None and self.edit_cell_validation:
                validation = self.extra_end_edit_cell_func(EditCellEvent(r, c, editor_info[2] if len(editor_info) >= 3 else "FocusOut", f"{self.text_editor_value}", "end_edit_cell"))
                if validation is not None:
                    self.text_editor_value = validation
                    self._set_cell_data(r, c, value = self.text_editor_value, redraw = False)
        if move_down:
            if r is None and c is None and editor_info:
                r, c = editor_info[0], editor_info[1]
            currently_selected = self.currently_selected()
            if (r is not None and
                c is not None and
                currently_selected and
                r == currently_selected[0] and
                c == currently_selected[1] and
                (self.single_selection_enabled or self.toggle_selection_enabled)
                ):
                r1, c1, r2, c2 = self.find_last_selected_box_with_current(currently_selected)
                numcols = c2 - c1
                numrows = r2 - r1
                if numcols == 1 and numrows == 1:
                    if editor_info is not None and len(editor_info) >= 3 and editor_info[2] == "Return":
                        self.select_cell(r + 1 if r < len(self.row_positions) - 2 else r, c)
                        self.see(r + 1 if r < len(self.row_positions) - 2 else r, c, keep_xscroll = True, bottom_right_corner = True, check_cell_visibility = True)
                    elif editor_info is not None and len(editor_info) >= 3 and editor_info[2] == "Tab":
                        self.select_cell(r, c + 1 if c < len(self.col_positions) - 2 else c)
                        self.see(r, c + 1 if c < len(self.col_positions) - 2 else c, keep_xscroll = True, bottom_right_corner = True, check_cell_visibility = True)
                else:
                    moved = False
                    new_r = r
                    new_c = c
                    if editor_info is not None and len(editor_info) >= 3 and editor_info[2] == "Return":
                        if r + 1 == r2:
                            new_r = r1
                        elif numrows > 1:
                            new_r = r + 1
                            moved = True
                        if not moved:
                            if c + 1 == c2:
                                new_c = c1
                            elif numcols > 1:
                                new_c = c + 1
                    elif editor_info is not None and len(editor_info) >= 3 and editor_info[2] == "Tab":
                        if c + 1 == c2:
                            new_c = c1
                        elif numcols > 1:
                            new_c = c + 1
                            moved = True
                        if not moved:
                            if r + 1 == r2:
                                new_r = r1
                            elif numrows > 1:
                                new_r = r + 1
                    self.create_current(new_r, new_c, type_ = currently_selected.type_, inside = True)
                    self.see(new_r, new_c, keep_xscroll = False, bottom_right_corner = True, check_cell_visibility = True)
        self.hide_dropdown_window(r, c)
        if recreate:
            self.recreate_all_selection_boxes()
        if redraw:
            self.refresh()
        if editor_info is not None and len(editor_info) >= 3 and editor_info[2] != "FocusOut":
            self.focus_set()
        return "break"
    
    def tab_key(self, event = None):
        currently_selected = self.currently_selected()
        if not currently_selected:
            return
        r = currently_selected.row
        c = currently_selected.column
        r1, c1, r2, c2 = self.find_last_selected_box_with_current(currently_selected)
        numcols = c2 - c1
        numrows = r2 - r1
        if numcols == 1 and numrows == 1:
            new_r = r
            new_c = c + 1 if c < len(self.col_positions) - 2 else c
        else:
            moved = False
            new_r = r
            new_c = c
            if c + 1 == c2:
                new_c = c1
            elif numcols > 1:
                new_c = c + 1
                moved = True
            if not moved:
                if r + 1 == r2:
                    new_r = r1
                elif numrows > 1:
                    new_r = r + 1
        self.create_current(new_r, new_c, type_ = currently_selected.type_, inside = True)
        self.see(new_r, new_c, keep_xscroll = False, bottom_right_corner = True, check_cell_visibility = True)
        return "break"

    #internal event use
    def _set_cell_data(self, r = 0, c = 0, drow = None, dcol = None, value = "", undo = True, cell_resize = True, redraw = True):
        if dcol is None:
            dcol = c if self.all_columns_displayed else self.displayed_columns[c]
        if drow is None:
            drow = r if self.all_rows_displayed else self.displayed_rows[r]
        if r > len(self.data) - 1:
            self.data.extend([list(repeat("", dcol + 1)) for i in range((r + 1) - len(self.data))])
        elif dcol > len(self.data[r]) - 1:
            self.data[r].extend(list(repeat("", (dcol + 1) - len(self.data[r]))))
        if self.undo_enabled and undo:
            if self.data[r][dcol] != value:
                self.undo_storage.append(zlib.compress(pickle.dumps(("edit_cells",
                                                                     {(r, dcol): self.data[r][dcol]},
                                                                     (((r, c, r + 1, c + 1), "cells"), ),
                                                                     self.currently_selected()))))
        self.data[r][dcol] = value
        if cell_resize and self.cell_auto_resize_enabled:
            self.set_cell_size_to_text(r, c, only_set_if_too_small = True, redraw = redraw, run_binding = True)
        return True

    #internal event use
    def _click_checkbox(self, r, c, dcol = None, undo = True, redraw = True):
        if dcol is None:
            dcol = c if self.all_columns_displayed else self.displayed_columns[c]
        if self.cell_options[(r, dcol)]['checkbox']['state'] == "normal":
            self._set_cell_data(r, c, dcol = dcol, value = not self.data[r][dcol] if type(self.data[r][dcol]) == bool else False, undo = undo, cell_resize = False)
            if self.cell_options[(r, dcol)]['checkbox']['check_function'] is not None:
                self.cell_options[(r, dcol)]['checkbox']['check_function']((r, c, "CheckboxClicked", f"{self.data[r][dcol]}"))
            if self.extra_end_edit_cell_func is not None:
                self.extra_end_edit_cell_func(EditCellEvent(r, c, "Return", f"{self.data[r][dcol]}", "end_edit_cell"))
        if redraw:
            self.refresh()

    def create_checkbox(self, r = 0, c = 0, checked = False, state = "normal", redraw = False, check_function = None, text = ""):
        if (r, c) in self.cell_options and any(x in self.cell_options[(r, c)] for x in ('dropdown', 'checkbox')):
            self.delete_dropdown_and_checkbox(r, c)
        self._set_cell_data(r, dcol = c, value = checked, cell_resize = False, undo = False) #only works because cell_resize is false and undo is false, otherwise needs c arg
        if (r, c) not in self.cell_options:
            self.cell_options[(r, c)] = {}
        self.cell_options[(r, c)]['checkbox'] = {'check_function': check_function,
                                                 'state': state,
                                                 'text': text}
        if redraw:
            self.refresh()

    def create_dropdown(self, r = 0, c = 0, values = [], set_value = None, state = "readonly", redraw = True, selection_function = None, modified_function = None):
        if (r, c) in self.cell_options and any(x in self.cell_options[(r, c)] for x in ('dropdown', 'checkbox')):
            self.delete_dropdown_and_checkbox(r, c)
        self._set_cell_data(r,
                            dcol = c,
                            value = set_value if set_value is not None else values[0] if values else "",
                            cell_resize = False,
                            undo = False) #only works because cell_resize is false and undo is false, otherwise needs c arg
        if (r, c) not in self.cell_options:
            self.cell_options[(r, c)] = {}
        self.cell_options[(r, c)]['dropdown'] = {'values': values,
                                                 'align': "w",
                                                 'window': "no dropdown open",
                                                 'canvas_id': "no dropdown open",
                                                 'select_function': selection_function,
                                                 'modified_function': modified_function,
                                                 'state': state}
        if redraw:
            self.refresh()

    def get_widget_bg_fg(self, r, c):
        bg = self.table_bg
        fg = self.table_fg
        if (r, c) in self.cell_options and 'highlight' in self.cell_options[(r, c)]:
            if self.cell_options[(r, c)]['highlight'][0] is not None:
                bg = self.cell_options[(r, c)]['highlight'][0]
            if self.cell_options[(r, c)]['highlight'][1] is not None:
                fg = self.cell_options[(r, c)]['highlight'][1]
        elif r in self.row_options and 'highlight' in self.row_options[r]:
            if self.row_options[r]['highlight'][0] is not None:
                bg = self.row_options[r]['highlight'][0]
            if self.row_options[r]['highlight'][1] is not None:
                fg = self.row_options[r]['highlight'][1]
        elif c in self.col_options and 'highlight' in self.col_options[c]:
            if self.col_options[c]['highlight'][0] is not None:
                bg = self.col_options[c]['highlight'][0]
            if self.col_options[c]['highlight'][1] is not None:
                fg = self.col_options[c]['highlight'][1]
        return bg, fg

    def get_space_bot(self, r, text_editor_h = None):
        if len(self.row_positions) <= 1:
            if text_editor_h is None:
                win_h = int(self.winfo_height())
                sheet_h = int(1 + self.empty_vertical)
            else:
                win_h = int(self.winfo_height() - text_editor_h)
                sheet_h = int(1 + self.empty_vertical - text_editor_h)
        else:
            if text_editor_h is None:
                win_h = int(self.canvasy(0) + self.winfo_height() - self.row_positions[r + 1])
                sheet_h = int(self.row_positions[-1] + 1 + self.empty_vertical - self.row_positions[r + 1])
            else:
                win_h = int(self.canvasy(0) + self.winfo_height() - (self.row_positions[r] + text_editor_h))
                sheet_h = int(self.row_positions[-1] + 1 + self.empty_vertical - (self.row_positions[r] + text_editor_h))
        if win_h > 0:
            win_h -= 1
        if sheet_h > 0:
            sheet_h -= 1
        return win_h if win_h >= sheet_h else sheet_h

    def get_dropdown_height_anchor(self, drow, dcol, text_editor_h = None):
        win_h = 5
        for i, v in enumerate(self.cell_options[(drow, dcol)]['dropdown']['values']):
            v_numlines = len(v.split("\n") if isinstance(v, str) else f"{v}".split("\n"))
            if v_numlines > 1:
                win_h += self.fl_ins + (v_numlines * self.xtra_lines_increment) + 5 # end of cell
            else:
                win_h += self.min_rh
            if i == 5:
                break
        if win_h > 500:
            win_h = 500
        space_bot = self.get_space_bot(drow, text_editor_h)
        space_top = int(self.row_positions[drow])
        anchor = "nw"
        win_h2 = int(win_h)
        if win_h > space_bot:
            if space_bot >= space_top:
                anchor = "nw"
                win_h = space_bot - 1
            elif space_top > space_bot:
                anchor = "sw"
                win_h = space_top - 1
        if win_h < self.txt_h + 5:
            win_h = self.txt_h + 5
        elif win_h > win_h2:
            win_h = win_h2
        return win_h, anchor

    # c is displayed col
    def display_dropdown_window(self, r, c, event = None):
        self.destroy_text_editor("Escape")
        self.destroy_opened_dropdown_window()
        drow = r if self.all_rows_displayed else self.displayed_rows[r]
        dcol = c if self.all_columns_displayed else self.displayed_columns[c]
        if self.cell_options[(r, dcol)]['dropdown']['state'] == "normal":
            if not self.edit_cell_(r = r, c = c, dropdown = True, event = event):
                return
        win_h, anchor = self.get_dropdown_height_anchor(drow, dcol)
        window = self.parentframe.dropdown_class(self.winfo_toplevel(),
                                                 r,
                                                 c,
                                                 width = self.col_positions[c + 1] - self.col_positions[c] + 1,
                                                 height = win_h,
                                                 font = self._font,
                                                 colors = {'bg': self.popup_menu_bg, 
                                                           'fg': self.popup_menu_fg, 
                                                           'highlight_bg': self.popup_menu_highlight_bg,
                                                           'highlight_fg': self.popup_menu_highlight_fg},
                                                 outline_color = self.table_selected_cells_border_fg,
                                                 outline_thickness = 2,
                                                 values = self.cell_options[(r, dcol)]['dropdown']['values'],
                                                 hide_dropdown_window = self.hide_dropdown_window,
                                                 arrowkey_RIGHT = self.arrowkey_RIGHT,
                                                 arrowkey_LEFT = self.arrowkey_LEFT,
                                                 align = self.cell_options[(r, dcol)]['dropdown']['align'])
        if self.cell_options[(r, dcol)]['dropdown']['state'] == "normal":
            if anchor == "nw":
                ypos = self.row_positions[r] + self.text_editor.h_ - 1
            else:
                ypos = self.row_positions[r]
            self.cell_options[(r, dcol)]['dropdown']['canvas_id'] = self.create_window((self.col_positions[c], ypos),
                                                                                        window = window,
                                                                                        anchor = anchor)
            if self.cell_options[(r, dcol)]['dropdown']['modified_function'] is not None:
                self.text_editor.textedit.bind("<<TextModified>>", lambda x: self.cell_options[(r, dcol)]['dropdown']['modified_function'](DropDownModifiedEvent("ComboboxModified", r, dcol, self.text_editor.get())))
            self.update_idletasks()
            try:
                self.after(1, lambda: self.text_editor.textedit.focus())
                self.after(2, self.text_editor.scroll_to_bottom())
            except:
                return
            redraw = False
        else:
            
            if anchor == "nw":
                ypos = self.row_positions[r + 1]
            else:
                ypos = self.row_positions[r]
            self.cell_options[(r, dcol)]['dropdown']['canvas_id'] = self.create_window((self.col_positions[c], ypos),
                                                                                        window = window,
                                                                                        anchor = anchor)
            self.update_idletasks()
            window.bind("<FocusOut>", lambda x: self.hide_dropdown_window(r, c))
            window.focus()
            redraw = True
        self.existing_dropdown_window = window
        self.cell_options[(r, dcol)]['dropdown']['window'] = window
        self.existing_dropdown_canvas_id = self.cell_options[(r, dcol)]['dropdown']['canvas_id']
        if redraw:
            self.main_table_redraw_grid_and_text(redraw_header = False, redraw_row_index = False)

    # displayed indexes, not data
    def hide_dropdown_window(self, r = None, c = None, selection = None, redraw = True):
        if r is not None and c is not None and selection is not None:
            dcol = c if self.all_columns_displayed else self.displayed_columns[c]
            drow = r if self.all_rows_displayed else self.displayed_rows[r]
            if self.cell_options[(r, dcol)]['dropdown']['select_function'] is not None: # user has specified a selection function
                self.cell_options[(r, dcol)]['dropdown']['select_function'](EditCellEvent(r, c, "ComboboxSelected", f"{selection}", "end_edit_cell"))
            if self.extra_end_edit_cell_func is None:
                self._set_cell_data(r, c, dcol = dcol, value = selection, redraw = not redraw)
            elif self.extra_end_edit_cell_func is not None and self.edit_cell_validation:
                validation = self.extra_end_edit_cell_func(EditCellEvent(r, c, "ComboboxSelected", f"{selection}", "end_edit_cell"))
                if validation is not None:
                    selection = validation
                self._set_cell_data(r, c, dcol = dcol, value = selection, redraw = not redraw)
            elif self.extra_end_edit_cell_func is not None and not self.edit_cell_validation:
                self._set_cell_data(r, c, dcol = dcol, value = selection, redraw = not redraw)
                self.extra_end_edit_cell_func(EditCellEvent(r, c, "ComboboxSelected", f"{selection}", "end_edit_cell"))
            self.focus_set()
            self.recreate_all_selection_boxes()
        self.destroy_text_editor("Escape")
        self.destroy_opened_dropdown_window(r, c)
        if redraw:
            self.refresh()
        
    def mouseclick_outside_editor_or_dropdown(self):
        if self.existing_dropdown_window is not None:
            closed_dd_coords = (int(self.existing_dropdown_window.r),
                                int(self.existing_dropdown_window.c))
        else:
            closed_dd_coords = None
        if self.text_editor_loc is not None and self.text_editor is not None:
            self.close_text_editor(editor_info = self.text_editor_loc + ("ButtonPress-1", ))
        else:
            self.destroy_text_editor("Escape")
        if closed_dd_coords:
            self.destroy_opened_dropdown_window(closed_dd_coords[0], closed_dd_coords[1]) #displayed coords not data, necessary for b1 function
        return closed_dd_coords

    # function can receive 4 None args
    def destroy_opened_dropdown_window(self, r = None, c = None, drow = None, dcol = None):
        if c is not None or dcol is not None:
            if dcol is None:
                dcol_ = c if self.all_columns_displayed else self.displayed_columns[c]
            else:
                dcol_ = dcol
        else:
            dcol_ = None
        if r is not None or drow is not None:
            if drow is None:
                drow_ = r if self.all_rows_displayed else self.displayed_rows[r]
            else:
                drow_ = drow
        else:
            drow_ = None
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
        if (drow_, dcol_) in self.cell_options and 'dropdown' in self.cell_options[(drow_, dcol_)]:
            self.cell_options[(drow_, dcol_)]['dropdown']['canvas_id'] = "no dropdown open"
            self.cell_options[(drow_, dcol_)]['dropdown']['window'] = "no dropdown open"
            try:
                self.delete(self.cell_options[(drow_, dcol_)]['dropdown']['canvas_id'])
            except:
                pass

    def get_displayed_col_from_dcol(self, dcol):
        try:
            return self.displayed_columns.index(dcol)
        except:
            return None
    
    # c is dcol
    def delete_dropdown(self, r, c):
        self.destroy_opened_dropdown_window(r, dcol = c)
        if (r, c) in self.cell_options and 'dropdown' in self.cell_options[(r, c)]:
            del self.cell_options[(r, c)]['dropdown']

    # c is dcol
    def delete_checkbox(self, r, c):
        if (r, c) in self.cell_options and 'checkbox' in self.cell_options[(r, c)]:
            del self.cell_options[(r, c)]['checkbox']

    # c is dcol
    def delete_dropdown_and_checkbox(self, r, c):
        self.delete_dropdown(r, c)
        self.delete_checkbox(r, c)

    # deprecated
    def refresh_dropdowns(self, dropdowns = []):
        pass

