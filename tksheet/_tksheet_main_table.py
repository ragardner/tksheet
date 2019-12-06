from ._tksheet_vars import *
from ._tksheet_other_classes import *

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


class MainTable(tk.Canvas):
    def __init__(self,
                 parentframe = None,
                 column_width = None,
                 column_headers_canvas = None,
                 row_index_canvas = None,
                 headers = None,
                 header_height = None,
                 row_height = None,
                 data_reference = None,
                 total_cols = None,
                 total_rows = None,
                 row_index = None,
                 font = None,
                 header_font = None,
                 align = None,
                 width = None,
                 height = None,
                 table_background = "white",
                 grid_color = "gray15",
                 text_color = "black",
                 selected_cells_background = "gray85",
                 selected_cells_foreground = "white",
                 displayed_columns = [],
                 all_columns_displayed = True):
        tk.Canvas.__init__(self,
                           parentframe,
                           width = width,
                           height = height,
                           background = table_background,
                           highlightthickness = 0)
        self.min_rh = 0
        self.hdr_min_rh = 0
        self.extra_motion_func = None
        self.extra_b1_press_func = None
        self.extra_b1_motion_func = None
        self.extra_b1_release_func = None
        self.extra_double_b1_func = None
        self.extra_ctrl_c_func = None
        self.extra_ctrl_x_func = None
        self.extra_ctrl_v_func = None
        self.extra_ctrl_z_func = None
        self.extra_delete_key_func = None
        self.extra_edit_cell_func = None
        self.selection_binding_func = None # function to run when a spreadsheet selection event occurs
        self.drag_selection_binding_func = None
        self.single_selection_enabled = False
        self.drag_selection_enabled = False
        self.multiple_selection_enabled = False
        self.arrowkeys_enabled = False
        self.undo_enabled = False
        self.new_row_width = 0
        self.new_header_height = 0
        self.parentframe = parentframe
        self.row_width_resize_bbox = tuple()
        self.header_height_resize_bbox = tuple()
        self.CH = column_headers_canvas
        self.CH.MT = self
        self.CH.RI = row_index_canvas
        self.RI = row_index_canvas
        self.RI.MT = self
        self.RI.CH = column_headers_canvas
        self.TL = None                # is set from within TopLeftRectangle() __init__
        self.total_rows = 0
        self.total_cols = 0
        self.data_ref = data_reference
        if isinstance(self.data_ref, list):
            self.data_ref = data_reference
        else:
            self.data_ref = []
        if self.data_ref and total_rows is None:
            self.total_rows = len(self.data_ref)
        if self.data_ref and total_cols is None:
            self.total_cols = len(max(self.data_ref, key = len))
        if not self.data_ref:
            if isinstance(total_rows, int) and isinstance(total_cols, int) and total_rows > 0 and total_cols > 0:
                self.data_ref = [["" for c in range(total_cols)] for r in range(total_rows)]
                self.total_rows = total_rows
                self.total_cols = total_cols
        self.displayed_columns = displayed_columns
        self.all_columns_displayed = all_columns_displayed

        self.grid_color = grid_color
        self.text_color = text_color
        self.selected_cells_background = selected_cells_background
        self.selected_cells_foreground = selected_cells_foreground
        self.table_background = table_background
        
        self.align = align
        self.my_font = font
        self.fnt_fam = font[0]
        self.fnt_sze = font[1]
        self.fnt_wgt = font[2]
        self.my_hdr_font = header_font
        self.hdr_fnt_fam = header_font[0]
        self.hdr_fnt_sze = header_font[1]
        self.hdr_fnt_wgt = header_font[2]

        self.txt_measure_canvas = tk.Canvas(self)
        self.table_dropdown = None
        self.table_dropdown_id = None
        self.table_dropdown_value = None
        self.default_cw = column_width
        self.default_rh = int(row_height)
        self.default_hh = int(header_height)
        self.set_fnt_help()
        self.set_hdr_fnt_help()
        
        if headers:
            self.my_hdrs = headers
        else:
            self.my_hdrs = []
        
        if isinstance(row_index, int):
            self.my_row_index = row_index
        else:
            if row_index:
                self.my_row_index = row_index
            else:
                self.my_row_index = []

        self.col_positions = [0]
        self.row_positions = [0]
        self.reset_col_positions()
        self.reset_row_positions()

        self.currently_selected = tuple() # can be a row ("row",row number) or column ("column",column number) or cell (row number,column number)
        self.selected_cells = set()
        self.selected_cols = set()
        self.selected_rows = set()
        self.highlighted_cells = {}

        self.undo_storage = deque(maxlen = 20)

        self.bind("<Motion>", self.mouse_motion)
        self.bind("<Shift-ButtonPress-1>",self.shift_b1_press)
        self.bind("<ButtonPress-1>", self.b1_press)
        self.bind("<B1-Motion>", self.b1_motion)
        self.bind("<ButtonRelease-1>", self.b1_release)
        self.bind("<Double-Button-1>", self.double_b1)
        self.bind("<Configure>", self.refresh)
        self.bind("<MouseWheel>", self.mousewheel)

        self.rc_popup_menu = tk.Menu(self, tearoff = 0)
        self.rc_popup_menu.add_command(label = "Select all (Ctrl-a)",
                                       command = self.select_all)
        self.rc_popup_menu.add_separator()
        self.rc_popup_menu.add_command(label = "Cut (Ctrl-x)",
                                       command = self.ctrl_x,
                                       state = "disabled")
        self.rc_popup_menu.add_separator()
        self.rc_popup_menu.add_command(label = "Copy (Ctrl-c)",
                                       command = self.ctrl_c,
                                       state = "disabled")
        self.rc_popup_menu.add_separator()
        self.rc_popup_menu.add_command(label = "Paste (Ctrl-v)",
                                       command = self.ctrl_v,
                                       state = "disabled")
        self.rc_popup_menu.add_separator()
        self.rc_popup_menu.add_command(label = "Delete (Del)",
                                       command = self.delete_key,
                                       state = "disabled")

        self.CH.ch_rc_popup_menu = tk.Menu(self.CH, tearoff = 0)
        self.CH.ch_rc_popup_menu.add_command(label = "Cut Contents (Ctrl-x)",
                                               command = self.ctrl_x,
                                       state = "disabled")
        self.CH.ch_rc_popup_menu.add_separator()
        self.CH.ch_rc_popup_menu.add_command(label = "Copy Contents (Ctrl-c)",
                                               command = self.ctrl_c,
                                       state = "disabled")
        self.CH.ch_rc_popup_menu.add_separator()
        self.CH.ch_rc_popup_menu.add_command(label = "Paste (Ctrl-v)",
                                               command = self.ctrl_v,
                                       state = "disabled")
        self.CH.ch_rc_popup_menu.add_separator()
        self.CH.ch_rc_popup_menu.add_command(label = "Clear Contents (Del)",
                                               command = self.delete_key,
                                       state = "disabled")
        self.CH.ch_rc_popup_menu.add_separator()
        self.CH.ch_rc_popup_menu.add_command(label = "Delete Columns",
                                           command = self.del_cols_rc,
                                           state = "disabled")
        self.CH.ch_rc_popup_menu.add_separator()
        self.CH.ch_rc_popup_menu.add_command(label = "Insert Column",
                                               command = self.insert_col_rc,
                                               state = "disabled")
        
        self.RI.ri_rc_popup_menu = tk.Menu(self.RI, tearoff = 0)
        self.RI.ri_rc_popup_menu.add_command(label = "Cut Contents (Ctrl-x)",
                                               command = self.ctrl_x,
                                       state = "disabled")
        self.RI.ri_rc_popup_menu.add_separator()
        self.RI.ri_rc_popup_menu.add_command(label = "Copy Contents (Ctrl-c)",
                                               command = self.ctrl_c,
                                       state = "disabled")
        self.RI.ri_rc_popup_menu.add_separator()
        self.RI.ri_rc_popup_menu.add_command(label = "Paste (Ctrl-v)",
                                               command = self.ctrl_v,
                                       state = "disabled")
        self.RI.ri_rc_popup_menu.add_separator()
        self.RI.ri_rc_popup_menu.add_command(label = "Clear Contents (Del)",
                                               command = self.delete_key,
                                       state = "disabled")
        self.RI.ri_rc_popup_menu.add_separator()
        self.RI.ri_rc_popup_menu.add_command(label = "Delete Rows",
                                           command = self.del_rows_rc,
                                           state = "disabled")
        self.RI.ri_rc_popup_menu.add_separator()
        self.RI.ri_rc_popup_menu.add_command(label = "Insert Row",
                                               command = self.insert_row_rc,
                                               state = "disabled")
        
    def refresh(self, event = None):
        self.main_table_redraw_grid_and_text(True, True)

    def shift_b1_press(self, event = None):
        if self.drag_selection_enabled and all(v is None for v in (self.RI.rsz_h, self.RI.rsz_w, self.CH.rsz_h, self.CH.rsz_w)):
            rowsel = int(self.identify_row(y = event.y))
            colsel = int(self.identify_col(x = event.x))
            if rowsel < len(self.row_positions) - 1 and colsel < len(self.col_positions) - 1 and (rowsel, colsel) not in self.selected_cells:
                if self.currently_selected and isinstance(self.currently_selected[0], int):
                    min_r = self.currently_selected[0]
                    min_c = self.currently_selected[1]
                    self.selected_cells = set()
                    self.CH.selected_cells = defaultdict(int)
                    self.RI.selected_cells = defaultdict(int)
                    self.selected_cols = set()
                    self.selected_rows = set()
                    if rowsel >= min_r and colsel >= min_c:
                        for r in range(min_r, rowsel + 1):
                            for c in range(min_c, colsel + 1):
                                self.selected_cells.add((r, c))
                                self.RI.selected_cells[r] += 1
                                self.CH.selected_cells[c] += 1
                    elif rowsel >= min_r and min_c >= colsel:
                        for r in range(min_r, rowsel + 1):
                            for c in range(colsel, min_c + 1):
                                self.selected_cells.add((r, c))
                                self.RI.selected_cells[r] += 1
                                self.CH.selected_cells[c] += 1
                    elif min_r >= rowsel and colsel >= min_c:
                        for r in range(rowsel, min_r + 1):
                            for c in range(min_c, colsel + 1):
                                self.selected_cells.add((r, c))
                                self.RI.selected_cells[r] += 1
                                self.CH.selected_cells[c] += 1
                    elif min_r >= rowsel and min_c >= colsel:
                        for r in range(rowsel, min_r + 1):
                            for c in range(colsel, min_c + 1):
                                self.selected_cells.add((r, c))
                                self.RI.selected_cells[r] += 1
                                self.CH.selected_cells[c] += 1
                else:
                    self.select_cell(rowsel, colsel, redraw = False)
                self.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
                if self.selection_binding_func is not None:
                    self.selection_binding_func(("cell", ) + tuple(self.currently_selected))

    def basic_bindings(self, onoff = "enable"):
        if onoff == "enable":
            self.bind("<Motion>", self.mouse_motion)
            self.bind("<ButtonPress-1>", self.b1_press)
            self.bind("<B1-Motion>", self.b1_motion)
            self.bind("<ButtonRelease-1>", self.b1_release)
            self.bind("<Double-Button-1>", self.double_b1)
            self.bind("<MouseWheel>", self.mousewheel)
        elif onoff == "disable":
            self.unbind("<Motion>")
            self.unbind("<ButtonPress-1>")
            self.unbind("<B1-Motion>")
            self.unbind("<ButtonRelease-1>")
            self.unbind("<Double-Button-1>")
            self.unbind("<MouseWheel>")

    def show_ctrl_outline(self, t = "cut", canvas = "table", start_cell = (0, 0), end_cell = (0, 0)):
        if t != "cut":
            self.create_rectangle(self.col_positions[start_cell[0]] + 1,
                                  self.row_positions[start_cell[1]] + 1,
                                  self.col_positions[end_cell[0]],
                                  self.row_positions[end_cell[1]],
                                  fill = "",
                                  width = 4,
                                  outline = self.grid_color,
                                  tag = "ctrl")
        else:
            self.create_rectangle(self.col_positions[start_cell[0]] + 1,
                                  self.row_positions[start_cell[1]] + 1,
                                  self.col_positions[end_cell[0]],
                                  self.row_positions[end_cell[1]],
                                  fill = "",
                                  dash = (10, 5),
                                  width = 4,
                                  outline = self.table_background,
                                  tag = "ctrl")
        self.tag_raise("ctrl")
        self.after(1000, self.del_ctrl_outline)

    def del_ctrl_outline(self, event = None):
        self.delete("ctrl")

    def ctrl_c(self, event = None):
        if self.anything_selected():
            s = io.StringIO()
            writer = csv_module.writer(s, dialect = csv_module.excel_tab, lineterminator = "\n")
            if self.selected_cols:
                x1 = self.get_min_selected_cell_x()
                x2 = self.get_max_selected_cell_x() + 1
                y1 = 0
                y2 = len(self.row_positions) - 1
            elif self.selected_rows:
                x1 = 0
                x2 = len(self.col_positions) - 1
                y1 = self.get_min_selected_cell_y()
                y2 = self.get_max_selected_cell_y() + 1
            else:
                x1 = self.get_min_selected_cell_x()
                x2 = self.get_max_selected_cell_x() + 1
                y1 = self.get_min_selected_cell_y()
                y2 = self.get_max_selected_cell_y() + 1
            if self.all_columns_displayed:
                for r in range(y1, y2):
                    l_ = [self.data_ref[r][c] for c in range(x1, x2)]
                    writer.writerow(l_)
            else:
                for r in range(y1, y2):
                    l_ = [self.data_ref[r][self.displayed_columns[c]] for c in range(x1, x2)]
                    writer.writerow(l_)
            self.clipboard_clear()
            s = s.getvalue().rstrip()
            self.clipboard_append(s)
            self.update()
            if self.extra_ctrl_c_func is not None:
                self.extra_ctrl_c_func()
            
    def ctrl_x(self, event = None):
        if self.anything_selected():
            if self.undo_enabled:
                undo_storage = {}
            s = io.StringIO()
            writer = csv_module.writer(s, dialect = csv_module.excel_tab, lineterminator = "\n")
            if self.selected_cols:
                x1 = self.get_min_selected_cell_x()
                x2 = self.get_max_selected_cell_x() + 1
                y1 = 0
                y2 = len(self.row_positions) - 1
            elif self.selected_rows:
                x1 = 0
                x2 = len(self.col_positions) - 1
                y1 = self.get_min_selected_cell_y()
                y2 = self.get_max_selected_cell_y() + 1
            else:
                x1 = self.get_min_selected_cell_x()
                x2 = self.get_max_selected_cell_x() + 1
                y1 = self.get_min_selected_cell_y()
                y2 = self.get_max_selected_cell_y() + 1
            if self.undo_enabled:
                undo_storage = {}
            if self.all_columns_displayed:
                for r in range(y1, y2):
                    l_ = []
                    for c in range(x1, x2):
                        sl = f"{self.data_ref[r][c]}"
                        l_.append(sl)
                        if self.undo_enabled:
                            undo_storage[(r, c)] = sl
                        self.data_ref[r][c] = ""
                    writer.writerow(l_)
            else:
                for r in range(y1, y2):
                    l_ = []
                    for c in range(x1, x2):
                        sl = f"{self.data_ref[r][self.displayed_columns[c]]}"
                        l_.append(sl)
                        if self.undo_enabled:
                            undo_storage[(r, self.displayed_columns[c])] = sl
                        self.data_ref[r][self.displayed_columns[c]] = ""
                    writer.writerow(l_)
            if self.undo_enabled:
                self.undo_storage.append(zlib.compress(pickle.dumps(("edit_cells", undo_storage))))
            self.clipboard_clear()
            s = s.getvalue().rstrip()
            self.clipboard_append(s)
            self.update()
            self.refresh()
            self.show_ctrl_outline(t = "cut", canvas = "table", start_cell = (x1, y1), end_cell = (x2, y2))
            if self.extra_ctrl_x_func is not None:
                self.extra_ctrl_x_func()

    def ctrl_v(self, event = None):
        if self.anything_selected():
            try:
                data = self.clipboard_get()
            except:
                return
            nd = []
            for r in csv_module.reader(io.StringIO(data), delimiter = "\t", quotechar = '"', skipinitialspace = True):
                try:
                    nd.append(r[:len(r) - next(i for i, c in enumerate(reversed(r)) if c)])
                except:
                    continue
            if not nd:
                return
            data = nd
            numcols = len(max(data, key = len))
            numrows = len(data)
            for rn, r in enumerate(data):
                if len(r) < numcols:
                    data[rn].extend(list(repeat("", numcols - len(r))))
            if self.undo_enabled:
                undo_storage = {}
            if self.selected_cols:
                x1 = self.get_min_selected_cell_x()
                y1 = 0
            elif self.selected_rows:
                x1 = 0
                y1 = self.get_min_selected_cell_y()
            else:
                x1 = self.get_min_selected_cell_x()
                y1 = self.get_min_selected_cell_y()
            if x1 + numcols > len(self.col_positions) - 1:
                numcols = len(self.col_positions) - 1 - x1
            if y1 + numrows > len(self.row_positions) - 1:
                numrows = len(self.row_positions) - 1 - y1 
            if self.all_columns_displayed:
                for ndr, r in enumerate(range(y1, y1 + numrows)):
                    l_ = []
                    for ndc, c in enumerate(range(x1, x1 + numcols)):
                        s = f"{self.data_ref[r][c]}"
                        l_.append(s)
                        if self.undo_enabled:
                            undo_storage[(r, c)] = s
                        self.data_ref[r][c] = data[ndr][ndc]
            else:
                for ndr, r in enumerate(range(y1, y1 + numrows)):
                    l_ = []
                    for ndc, c in enumerate(range(x1, x1 + numcols)):
                        s = f"{self.data_ref[r][self.displayed_columns[c]]}"
                        l_.append(s)
                        if self.undo_enabled:
                            undo_storage[(r, self.displayed_columns[c])] = s
                        self.data_ref[r][self.displayed_columns[c]] = data[ndr][ndc]
            if self.undo_enabled:
                self.undo_storage.append(zlib.compress(pickle.dumps(("edit_cells", undo_storage))))
            self.refresh()
            self.show_ctrl_outline(t = "paste", canvas = "table", start_cell = (x1, y1), end_cell = (x1 + numcols, y1 + numrows))
            if self.extra_ctrl_v_func is not None:
                self.extra_ctrl_v_func()

    def ctrl_z(self, event = None):
        if self.undo_storage:
            start_row = float("inf")
            start_col = float("inf")
            undo_storage = pickle.loads(zlib.decompress(self.undo_storage.pop()))
            if undo_storage[0] == "edit_cells":
                for (r, c), v in undo_storage[1].items():
                    if r < start_row:
                        start_row = r
                    if c < start_col:
                        start_col = c
                    self.data_ref[r][c] = v
            elif undo_storage[0] == "move_rows":
                pass

            elif undo_storage[0] == "move_cols":
                pass

            elif undo_storage[0] == "insert_row":
                pass

            elif undo_storage[0] == "insert_col":
                pass

            elif undo_storage[0] == "delete_rows":
                pass

            elif undo_storage[0] == "delete_cols":
                pass

            
            self.select_cell(start_row, start_col)
            self.see(r = start_row, c = start_col, keep_yscroll = False, keep_xscroll = False, bottom_right_corner = False, check_cell_visibility = True, redraw = False)
            self.refresh()
            if self.extra_ctrl_z_func is not None:
                self.extra_ctrl_z_func()

    def delete_key(self, event = None):
        if self.anything_selected():
            if self.undo_enabled:
                undo_storage = {}
            if self.selected_cols:
                x1 = self.get_min_selected_cell_x()
                x2 = self.get_max_selected_cell_x() + 1
                y1 = 0
                y2 = len(self.row_positions) - 1
            elif self.selected_rows:
                x1 = 0
                x2 = len(self.col_positions) - 1
                y1 = self.get_min_selected_cell_y()
                y2 = self.get_max_selected_cell_y() + 1
            else:
                x1 = self.get_min_selected_cell_x()
                x2 = self.get_max_selected_cell_x() + 1
                y1 = self.get_min_selected_cell_y()
                y2 = self.get_max_selected_cell_y() + 1
            if self.all_columns_displayed:
                for r in range(y1, y2):
                    for c in range(x1, x2):
                        if self.undo_enabled:
                            undo_storage[(r, c)] = f"{self.data_ref[r][c]}"
                        self.data_ref[r][c] = ""
            else:
                for r in range(y1, y2):
                    for c in range(x1, x2):
                        if self.undo_enabled:
                            undo_storage[(r, self.displayed_columns[c])] = f"{self.data_ref[r][self.displayed_columns[c]]}"
                        self.data_ref[r][self.displayed_columns[c]] = ""
            if self.undo_enabled:
                self.undo_storage.append(zlib.compress(pickle.dumps(("edit_cells", undo_storage))))
            self.refresh()
            if self.extra_delete_key_func is not None:
                self.extra_delete_key_func()
            
    def bind_arrowkeys(self, event = None):
        self.arrowkeys_enabled = True
        for canvas in (self, self.CH, self.RI, self.TL):
            canvas.bind("<Up>", self.arrowkey_UP)
            canvas.bind("<Right>", self.arrowkey_RIGHT)
            canvas.bind("<Down>", self.arrowkey_DOWN)
            canvas.bind("<Left>", self.arrowkey_LEFT)
            canvas.bind("<Prior>", self.page_UP)
            canvas.bind("<Next>", self.page_DOWN)

    def unbind_arrowkeys(self, event = None):
        self.arrowkeys_enabled = False
        for canvas in (self, self.CH, self.RI, self.TL):
            canvas.unbind("<Up>")
            canvas.unbind("<Right>")
            canvas.unbind("<Down>")
            canvas.unbind("<Left>")
            canvas.unbind("<Prior>")
            canvas.unbind("<Next>")

    def see(self, r = None, c = None, keep_yscroll = False, keep_xscroll = False, bottom_right_corner = False, check_cell_visibility = True,
            redraw = True):
        if check_cell_visibility:
            visible = self.cell_is_completely_visible(r = r, c = c)
        else:
            visible = False
        if not visible:
            if bottom_right_corner:
                if r is not None and not keep_yscroll:
                    y = self.row_positions[r + 1] + 1 - self.winfo_height()
                    args = ("moveto", y / (self.row_positions[-1] + 100))
                    self.yview(*args)
                    self.RI.yview(*args)
                    if redraw:
                        self.main_table_redraw_grid_and_text(redraw_row_index = True)
                if c is not None and not keep_xscroll:
                    x = self.col_positions[c + 1] + 1 - self.winfo_width()
                    args = ("moveto",x / (self.col_positions[-1] + 150))
                    self.xview(*args)
                    self.CH.xview(*args)
                    if redraw:
                        self.main_table_redraw_grid_and_text(redraw_header = True)
            else:
                if r is not None and not keep_yscroll:
                    args = ("moveto", self.row_positions[r] / (self.row_positions[-1] + 100))
                    self.yview(*args)
                    self.RI.yview(*args)
                    self.main_table_redraw_grid_and_text(redraw_row_index = True)
                if c is not None and not keep_xscroll:
                    args = ("moveto", self.col_positions[c] / (self.col_positions[-1] + 150))
                    self.xview(*args)
                    self.CH.xview(*args)
                    if redraw:
                        self.main_table_redraw_grid_and_text(redraw_header = True)

    def cell_is_completely_visible(self, r = 0, c = 0, cell_coords = None):
        cx1, cy1, cx2, cy2 = self.get_canvas_visible_area()
        if cell_coords is None:
            x1, y1, x2, y2 = self.GetCellCoords(r = r, c = c, sel = True)
        else:
            x1, y1, x2, y2 = cell_coords
        if cx1 > x1 or cy1 > y1 or cx2 < x2 or cy2 < y2:
            return False
        return True

    def cell_is_visible(self,r = 0, c = 0, cell_coords = None):
        cx1, cy1, cx2, cy2 = self.get_canvas_visible_area()
        if cell_coords is None:
            x1, y1, x2, y2 = self.GetCellCoords(r = r, c = c, sel = True)
        else:
            x1, y1, x2, y2 = cell_coords
        if x1 <= cx2 or y1 <= cy2 or x2 >= cx1 or y2 >= cy1:
            return True
        return False

    def select(self, r, c, cell,redraw = True):
        if cell is not None:
            r, c = cell[0], cell[1]
            self.RI.selected_cells[r] += 1
            self.CH.selected_cells[c] += 1
            self.selected_cells.add(cell)
            if cell == self.currently_selected:
                self.currently_selected = cell
        elif r is not None:
            self.RI.selected_cells[r] += 1
            self.selected_rows.add(r)
        elif c is not None:
            self.CH.selected_cells[c] += 1
            self.selected_cols.add(c)
        if redraw:
            self.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)

    def select_cell(self, r, c, redraw = False):
        r = int(r)
        c = int(c)
        self.currently_selected = (r, c)
        self.selected_cells = {(r, c)}
        self.CH.selected_cells = defaultdict(int)
        self.CH.selected_cells[c] += 1
        self.RI.selected_cells = defaultdict(int)
        self.RI.selected_cells[r] += 1
        self.selected_cols = set()
        self.selected_rows = set()
        if redraw:
            self.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
        if self.selection_binding_func is not None:
            self.selection_binding_func(("cell", ) + tuple(self.currently_selected))

    def highlight_cells(self, r = 0, c = 0, cells = tuple(), bg = None, fg = None, redraw = False):
        if bg is None and fg is None:
            return
        if cells:
            self.highlighted_cells = {t: (bg, fg) for t in cells}
        else:
            self.highlighted_cells[(r, c)] = (bg, fg)
        if redraw:
            self.main_table_redraw_grid_and_text()

    def add_selection(self, r, c, redraw = False, run_binding_func = True):
        r = int(r)
        c = int(c)
        self.currently_selected = (r, c)
        if (r, c) not in self.selected_cells:
            self.selected_cells.add((r, c))
            self.CH.selected_cells[c] += 1
            self.RI.selected_cells[r] += 1
        self.selected_cols = set()
        self.selected_rows = set()
        if redraw:
            self.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
        if self.selection_binding_func is not None and run_binding_func:
            self.selection_binding_func(("cell", ) + tuple(self.currently_selected))

    def select_all(self, redraw = True, run_binding_func = True):
        self.deselect("all")
        if len(self.row_positions) > 1 and len(self.col_positions) > 1:
            self.currently_selected = (0, 0)
            cols = tuple(range(len(self.col_positions) - 1))
            for r in range(len(self.row_positions) - 1):
                for c in cols:
                    self.selected_cells.add((r, c))
                    self.CH.selected_cells[c] += 1
                    self.RI.selected_cells[r] += 1
        if redraw:
            self.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
        if self.selection_binding_func is not None and run_binding_func:
            self.selection_binding_func("all")

    def page_UP(self, event = None):
        if not self.arrowkeys_enabled:
            return
        height = self.winfo_height()
        top = self.canvasy(0)
        scrollto = top - height
        if scrollto < 0:
            scrollto = 0
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
        end = self.row_positions[-1]
        if scrollto > end  + 100:
            scrollto = end
        args = ("moveto", scrollto / (end + 100))
        self.yview(*args)
        self.RI.yview(*args)
        self.main_table_redraw_grid_and_text(redraw_row_index = True)
        
    def arrowkey_UP(self, event = None):
        if not self.currently_selected or not self.arrowkeys_enabled:
            return
        if self.currently_selected[0] == "row":
            r = self.currently_selected[1]
            if r != 0 and self.RI.row_selection_enabled:
                if self.cell_is_completely_visible(r = r - 1, c = 0):
                    self.RI.select_row(r - 1, redraw = True)
                else:
                    self.RI.select_row(r - 1)
                    self.see(r - 1, 0, keep_xscroll = True, check_cell_visibility = False)
        elif isinstance(self.currently_selected[0],int):
            r = self.currently_selected[0]
            c = self.currently_selected[1]
            if r == 0 and self.CH.col_selection_enabled:
                if self.cell_is_completely_visible(r = r, c = 0):
                    self.CH.select_col(c, redraw = True)
                else:
                    self.CH.select_col(c)
                    self.see(r, c, keep_xscroll = True, check_cell_visibility = False)
            elif r != 0 and (self.single_selection_enabled or self.multiple_selection_enabled):
                if self.cell_is_completely_visible(r = r - 1, c = c):
                    self.select_cell(r - 1, c, redraw = True)
                else:
                    self.select_cell(r - 1, c)
                    self.see(r - 1, c, keep_xscroll = True, check_cell_visibility = False)
                
    def arrowkey_RIGHT(self, event = None):
        if not self.currently_selected or not self.arrowkeys_enabled:
            return
        if self.currently_selected[0] == "row":
            r = self.currently_selected[1]
            if self.single_selection_enabled or self.multiple_selection_enabled:
                if self.cell_is_completely_visible(r = r, c = 0):
                    self.select_cell(r, 0, redraw = True)
                else:
                    self.select_cell(r, 0)
                    self.see(r, 0, keep_yscroll = True, bottom_right_corner = True, check_cell_visibility = False)
        elif self.currently_selected[0] == "column":
            c = self.currently_selected[1]
            if c < len(self.col_positions) - 2 and self.CH.col_selection_enabled:
                if self.cell_is_completely_visible(r = 0, c = c + 1):
                    self.CH.select_col(c + 1, redraw = True)
                else:
                    self.CH.select_col(c + 1)
                    self.see(0, c + 1, keep_yscroll = True, bottom_right_corner = True, check_cell_visibility = False)
        elif isinstance(self.currently_selected[0], int):
            r = self.currently_selected[0]
            c = self.currently_selected[1]
            if c < len(self.col_positions) - 2 and (self.single_selection_enabled or self.multiple_selection_enabled):
                if self.cell_is_completely_visible(r = r, c = c + 1):
                    self.select_cell(r, c + 1, redraw =True)
                else:
                    self.select_cell(r, c + 1)
                    self.see(r, c + 1, keep_yscroll = True, bottom_right_corner = True, check_cell_visibility = False)

    def arrowkey_DOWN(self, event = None):
        if not self.currently_selected or not self.arrowkeys_enabled:
            return
        if self.currently_selected[0] == "row":
            r = self.currently_selected[1]
            if r < len(self.row_positions) - 2 and self.RI.row_selection_enabled:
                if self.cell_is_completely_visible(r = r + 1, c = 0):
                    self.RI.select_row(r + 1, redraw = True)
                else:
                    self.RI.select_row(r + 1)
                    self.see(r + 1, 0, keep_xscroll = True, bottom_right_corner = True, check_cell_visibility = False)
        elif self.currently_selected[0] == "column":
            c = self.currently_selected[1]
            if self.single_selection_enabled or self.multiple_selection_enabled:
                if self.cell_is_completely_visible(r = 0, c = c):
                    self.select_cell(0, c, redraw = True)
                else:
                    self.select_cell(0, c)
                    self.see(0, c, keep_xscroll = True, bottom_right_corner = True, check_cell_visibility = False)
        elif isinstance(self.currently_selected[0],int):
            r = self.currently_selected[0]
            c = self.currently_selected[1]
            if r < len(self.row_positions) - 2 and (self.single_selection_enabled or self.multiple_selection_enabled):
                if self.cell_is_completely_visible(r = r + 1, c = c):
                    self.select_cell(r + 1, c, redraw = True)
                else:
                    self.select_cell(r + 1, c)
                    self.see(r + 1, c, keep_xscroll = True, bottom_right_corner = True, check_cell_visibility = False)
                    
    def arrowkey_LEFT(self, event = None):
        if not self.currently_selected or not self.arrowkeys_enabled:
            return
        if self.currently_selected[0] == "column":
            c = self.currently_selected[1]
            if c != 0 and self.CH.col_selection_enabled:
                if self.cell_is_completely_visible(r = 0, c = c - 1):
                    self.CH.select_col(c - 1, redraw = True)
                else:
                    self.CH.select_col(c - 1)
                    self.see(0, c - 1, keep_yscroll = True, bottom_right_corner = True, check_cell_visibility = False)
        elif isinstance(self.currently_selected[0], int):
            r = self.currently_selected[0]
            c = self.currently_selected[1]
            if c == 0 and self.RI.row_selection_enabled:
                if self.cell_is_completely_visible(r = r, c = 0):
                    self.RI.select_row(r, redraw = True)
                else:
                    self.RI.select_row(r)
                    self.see(r, c, keep_yscroll = True, check_cell_visibility = False)
            elif c != 0 and (self.single_selection_enabled or self.multiple_selection_enabled):
                if self.cell_is_completely_visible(r = r, c = c - 1):
                    self.select_cell(r, c - 1, redraw = True)
                else:
                    self.select_cell(r, c - 1)
                    self.see(r, c - 1, keep_yscroll = True, check_cell_visibility = False)

    def edit_bindings(self, enable = True, key = None):
        if key is None or key == "copy":
            if enable:
                self.bind("<Control-c>", self.ctrl_c)
                self.bind("<Control-C>", self.ctrl_c)
                self.RI.bind("<Control-c>", self.ctrl_c)
                self.RI.bind("<Control-C>", self.ctrl_c)
                self.CH.bind("<Control-c>", self.ctrl_c)
                self.CH.bind("<Control-C>", self.ctrl_c)
                self.rc_popup_menu.entryconfig("Copy (Ctrl-c)", state = "normal")
                self.CH.ch_rc_popup_menu.entryconfig("Copy Contents (Ctrl-c)", state = "normal")
                self.RI.ri_rc_popup_menu.entryconfig("Copy Contents (Ctrl-c)", state = "normal")
            else:
                self.unbind("<Control-c>")
                self.unbind("<Control-C>")
                self.RI.unbind("<Control-c>")
                self.RI.unbind("<Control-C>")
                self.CH.unbind("<Control-c>")
                self.CH.unbind("<Control-C>")
                self.rc_popup_menu.entryconfig("Copy (Ctrl-c)", state = "disabled")
                self.CH.ch_rc_popup_menu.entryconfig("Copy Contents (Ctrl-c)", state = "disabled")
                self.RI.ri_rc_popup_menu.entryconfig("Copy Contents (Ctrl-c)", state = "disabled")
        if key is None or key == "cut":
            if enable:
                self.bind("<Control-x>", self.ctrl_x)
                self.bind("<Control-X>", self.ctrl_x)
                self.RI.bind("<Control-x>", self.ctrl_x)
                self.RI.bind("<Control-X>", self.ctrl_x)
                self.CH.bind("<Control-x>", self.ctrl_x)
                self.CH.bind("<Control-X>", self.ctrl_x)
                self.rc_popup_menu.entryconfig("Cut (Ctrl-x)", state = "normal")
                self.CH.ch_rc_popup_menu.entryconfig("Cut Contents (Ctrl-x)", state = "normal")
                self.RI.ri_rc_popup_menu.entryconfig("Cut Contents (Ctrl-x)", state = "normal")
            else:
                self.unbind("<Control-x>")
                self.unbind("<Control-X>")
                self.RI.unbind("<Control-x>")
                self.RI.unbind("<Control-X>")
                self.CH.unbind("<Control-x>")
                self.CH.unbind("<Control-X>")
                self.rc_popup_menu.entryconfig("Cut (Ctrl-x)", state = "disabled")
                self.CH.ch_rc_popup_menu.entryconfig("Cut Contents (Ctrl-x)", state = "disabled")
                self.RI.ri_rc_popup_menu.entryconfig("Cut Contents (Ctrl-x)", state = "disabled")
        if key is None or key == "paste":
            if enable:
                self.bind("<Control-v>", self.ctrl_v)
                self.bind("<Control-V>", self.ctrl_v)
                self.RI.bind("<Control-v>", self.ctrl_v)
                self.RI.bind("<Control-V>", self.ctrl_v)
                self.CH.bind("<Control-v>", self.ctrl_v)
                self.CH.bind("<Control-V>", self.ctrl_v)
                self.rc_popup_menu.entryconfig("Paste (Ctrl-v)", state = "normal")
                self.CH.ch_rc_popup_menu.entryconfig("Paste (Ctrl-v)", state = "normal")
                self.RI.ri_rc_popup_menu.entryconfig("Paste (Ctrl-v)", state = "normal")
            else:
                self.unbind("<Control-v>")
                self.unbind("<Control-V>")
                self.RI.unbind("<Control-v>")
                self.RI.unbind("<Control-V>")
                self.CH.unbind("<Control-v>")
                self.CH.unbind("<Control-V>")
                self.rc_popup_menu.entryconfig("Paste (Ctrl-v)", state = "disabled")
                self.CH.ch_rc_popup_menu.entryconfig("Paste (Ctrl-v)", state = "disabled")
                self.RI.ri_rc_popup_menu.entryconfig("Paste (Ctrl-v)", state = "disabled")
        if key is None or key == "undo":
            if enable:
                self.undo_enabled = True
                self.bind("<Control-z>", self.ctrl_z)
                self.bind("<Control-Z>", self.ctrl_z)
                self.RI.bind("<Control-z>", self.ctrl_z)
                self.RI.bind("<Control-Z>", self.ctrl_z)
                self.CH.bind("<Control-z>", self.ctrl_z)
                self.CH.bind("<Control-Z>", self.ctrl_z)
            else:
                self.undo_enabled = False
                self.unbind("<Control-z>")
                self.unbind("<Control-Z>")
                self.RI.unbind("<Control-z>")
                self.RI.unbind("<Control-Z>")
                self.CH.unbind("<Control-z>")
                self.CH.unbind("<Control-Z>")
        if key is None or key == "delete":
            if enable:
                self.bind("<Delete>", self.delete_key)
                self.RI.bind("<Delete>", self.delete_key)
                self.CH.bind("<Delete>", self.delete_key)
                self.rc_popup_menu.entryconfig("Delete (Del)", state = "normal")
                self.RI.ri_rc_popup_menu.entryconfig("Clear Contents (Del)", state = "normal")
                self.CH.ch_rc_popup_menu.entryconfig("Clear Contents (Del)", state = "normal")
            else:
                self.unbind("<Delete>")
                self.RI.unbind("<Delete>")
                self.CH.unbind("<Delete>")
                self.rc_popup_menu.entryconfig("Delete (Del)", state = "disabled")
                self.RI.ri_rc_popup_menu.entryconfig("Clear Contents (Del)", state = "disabled")
                self.CH.ch_rc_popup_menu.entryconfig("Clear Contents (Del)", state = "disabled")
        if key is None or key == "rc":
            self.bind_rc(enable)
        if key is None or key == "edit_cell":
            if enable:
                self.bind_cell_edit(True)
            else:
                self.bind_cell_edit(False)

    def bind_rc(self, enable = True):
        if enable:
            if str(get_os()) == "Darwin":
                self.bind("<2>", self.rc)
                self.RI.bind("<2>", self.RI.rc)
                self.CH.bind("<2>", self.CH.rc)
            else:
                self.bind("<3>", self.rc)
                self.RI.bind("<3>", self.RI.rc)
                self.CH.bind("<3>", self.CH.rc)
        else:
            if str(get_os()) == "Darwin":
                self.unbind("<2>")
                self.RI.unbind("<2>")
                self.CH.unbind("<2>")
            else:
                self.unbind("<3>")
                self.RI.unbind("<3>")
                self.CH.unbind("<3>")

    def bind_cell_edit(self, enable = True):
        if enable:
            for c in chain(lowercase_letters, uppercase_letters):
                self.bind(f"<{c}>", self.edit_cell_)
            for c in chain(numbers, symbols, other_symbols):
                self.bind(c, self.edit_cell_)
            self.bind("<F2>", self.edit_cell_)
            self.bind("<Double-Button-1>", self.edit_cell_)
        else:
            for c in chain(lowercase_letters, uppercase_letters):
                self.unbind(f"<{c}>")
            for c in chain(numbers, symbols, other_symbols):
                self.unbind(c)
            self.unbind("<F2>")
            self.unbind("<Double-Button-1>")

    def enable_bindings(self, bindings):
        if isinstance(bindings,(list, tuple)):
            for binding in bindings:
                self.enable_bindings_internal(binding)
        elif isinstance(bindings, str):
            self.enable_bindings_internal(bindings)

    def enable_bindings_internal(self, binding):
        if binding == "single":
            self.single_selection_enabled = True
            self.multiple_selection_enabled = False
        elif binding == "multiple":
            self.multiple_selection_enabled = True
            self.single_selection_enabled = False
        elif binding == "drag_select":
            self.drag_selection_enabled = True
            self.bind("<Control-a>", self.select_all)
            self.bind("<Control-A>", self.select_all)
            self.RI.bind("<Control-a>", self.select_all)
            self.RI.bind("<Control-A>", self.select_all)
            self.CH.bind("<Control-a>", self.select_all)
            self.CH.bind("<Control-A>", self.select_all)
        elif binding == "column_width_resize":
            self.CH.enable_bindings("column_width_resize")
        elif binding == "column_select":
            self.CH.enable_bindings("column_select")
        elif binding == "column_height_resize":
            self.CH.enable_bindings("column_height_resize")
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
        elif binding == "row_select":
            self.RI.enable_bindings("row_select")
        elif binding == "row_drag_and_drop":
            self.RI.enable_bindings("drag_and_drop")
        elif binding == "arrowkeys":
            self.bind_arrowkeys()
        elif binding == "edit_bindings":
            self.edit_bindings(True)
        elif binding == "rc_delete_column":
            self.CH.enable_bindings("rc_delete_column")
            self.bind_rc(True)
        elif binding == "rc_delete_row":
            self.RI.enable_bindings("rc_delete_row")
            self.bind_rc(True)
        elif binding == "rc_insert_column":
            self.CH.enable_bindings("rc_insert_column")
            self.bind_rc(True)
        elif binding == "rc_insert_row":
            self.RI.enable_bindings("rc_insert_row")
            self.bind_rc(True)
        elif binding == "copy":
            self.edit_bindings(True, "copy")
        elif binding == "cut":
            self.edit_bindings(True, "cut")
        elif binding == "paste":
            self.edit_bindings(True, "paste")
        elif binding == "delete":
            self.edit_bindings(True, "delete")
        elif binding == "rc":
            self.edit_bindings(True, "rc")
        elif binding == "undo":
            self.edit_bindings(True, "undo")
        elif binding == "edit_cell":
            self.edit_bindings(True, "edit_cell")
        
    def disable_bindings(self, bindings):
        if isinstance(bindings,(list, tuple)):
            for binding in bindings:
                self.disable_bindings_internal(binding)
        elif isinstance(bindings, str):
            self.disable_bindings_internal(bindings)

    def disable_bindings_internal(self, binding):
        if binding == "single":
            self.single_selection_enabled = False
        elif binding == "multiple":
            self.multiple_selection_enabled = False
        elif binding == "drag_select":
            self.drag_selection_enabled = False
            self.unbind("<Control-a>")
            self.unbind("<Control-A>")
            self.RI.unbind("<Control-a>")
            self.RI.unbind("<Control-A>")
            self.CH.unbind("<Control-a>")
            self.CH.unbind("<Control-A>")
        elif binding == "column_width_resize":
            self.CH.disable_bindings("column_width_resize")
        elif binding == "column_select":
            self.CH.disable_bindings("column_select")
        elif binding == "column_height_resize":
            self.CH.disable_bindings("column_height_resize")
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
        elif binding == "row_select":
            self.RI.disable_bindings("row_select")
        elif binding == "row_drag_and_drop":
            self.RI.disable_bindings("drag_and_drop")
        elif binding == "arrowkeys":
            self.unbind_arrowkeys()
        elif binding == "rc_delete_column":
            self.CH.disable_bindings("rc_delete_column")
            self.bind_rc(False)
        elif binding == "rc_delete_row":
            self.RI.disable_bindings("rc_delete_row")
            self.bind_rc(False)
        elif binding == "rc_insert_column":
            self.CH.disable_bindings("rc_insert_column")
            self.bind_rc(False)
        elif binding == "rc_insert_row":
            self.RI.disable_bindings("rc_insert_row")
            self.bind_rc(False)
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
        elif binding == "rc":
            self.edit_bindings(False, "rc")
        elif binding == "undo":
            self.edit_bindings(False, "undo")
        elif binding == "edit_cell":
            self.edit_bindings(False, "edit_cell")

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
            if self.RI.width_resizing_enabled and not mouse_over_resize:
                try:
                    x1, y1, x2, y2 = self.row_width_resize_bbox[0], self.row_width_resize_bbox[1], self.row_width_resize_bbox[2], self.row_width_resize_bbox[3]
                    if x >= x1 and y >= y1 and x <= x2 and y <= y2:
                        self.config(cursor = "sb_h_double_arrow")
                        self.RI.config(cursor = "sb_h_double_arrow")
                        self.RI.rsz_w = True
                        mouse_over_resize = True
                except:
                    pass
            if self.CH.height_resizing_enabled and not mouse_over_resize:
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
        self.focus_set()
        if self.identify_col(x = event.x, allow_end = False) is None or self.identify_row(y = event.y, allow_end = False) is None:
            self.deselect("all")
        elif self.single_selection_enabled and all(v is None for v in (self.RI.rsz_h, self.RI.rsz_w, self.CH.rsz_h, self.CH.rsz_w)):
            r = self.identify_row(y = event.y)
            c = self.identify_col(x = event.x)
            if r < len(self.row_positions) - 1 and c < len(self.col_positions) - 1:
                cols_selected = self.anything_selected(exclude_rows = True, exclude_cells = True)
                rows_selected = self.anything_selected(exclude_columns = True, exclude_cells = True)
                if rows_selected and not cols_selected:
                    x1 = 0
                    x2 = len(self.col_positions) - 1
                    y1 = self.get_min_selected_cell_y()
                    y2 = self.get_max_selected_cell_y()
                elif cols_selected and not rows_selected:
                    x1 = self.get_min_selected_cell_x()
                    x2 = self.get_max_selected_cell_x()
                    y1 = 0
                    y2 = len(self.row_positions) - 1
                else:
                    x1 = self.get_min_selected_cell_x()
                    x2 = self.get_max_selected_cell_x()
                    y1 = self.get_min_selected_cell_y()
                    y2 = self.get_max_selected_cell_y()
                if all(e is not None for e in (x1, x2, y1, y2)) and r >= y1 and c >= x1 and r <= y2 and c <= x2:
                    self.rc_popup_menu.tk_popup(event.x_root, event.y_root)
                else:
                    self.select_cell(r, c, redraw = True)
                    self.rc_popup_menu.tk_popup(event.x_root, event.y_root)
        elif self.multiple_selection_enabled and all(v is None for v in (self.RI.rsz_h, self.RI.rsz_w, self.CH.rsz_h, self.CH.rsz_w)):
            r = self.identify_row(y = event.y)
            c = self.identify_col(x = event.x)
            if r < len(self.row_positions) - 1 and c < len(self.col_positions) - 1:
                cols_selected = self.anything_selected(exclude_rows = True, exclude_cells = True)
                rows_selected = self.anything_selected(exclude_columns = True, exclude_cells = True)
                if rows_selected and not cols_selected:
                    x1 = 0
                    x2 = len(self.col_positions) - 1
                    y1 = self.get_min_selected_cell_y()
                    y2 = self.get_max_selected_cell_y()
                elif cols_selected and not rows_selected:
                    x1 = self.get_min_selected_cell_x()
                    x2 = self.get_max_selected_cell_x()
                    y1 = 0
                    y2 = len(self.row_positions) - 1
                else:
                    x1 = self.get_min_selected_cell_x()
                    x2 = self.get_max_selected_cell_x()
                    y1 = self.get_min_selected_cell_y()
                    y2 = self.get_max_selected_cell_y()
                if all(e is not None for e in (x1, x2, y1, y2)) and r >= y1 and c >= x1 and r <= y2 and c <= x2:
                    self.rc_popup_menu.tk_popup(event.x_root, event.y_root)
                else:
                    self.add_selection(r, c, redraw = True)
                    self.rc_popup_menu.tk_popup(event.x_root, event.y_root)

    def b1_press(self, event = None):
        self.focus_set()
        x1, y1, x2, y2 = self.get_canvas_visible_area()
        if self.identify_col(x = event.x, allow_end = False) is None or self.identify_row(y = event.y, allow_end = False) is None:
            self.deselect("all")
        elif self.single_selection_enabled and all(v is None for v in (self.RI.rsz_h, self.RI.rsz_w, self.CH.rsz_h, self.CH.rsz_w)):
            r = self.identify_row(y = event.y)
            c = self.identify_col(x = event.x)
            if r < len(self.row_positions) - 1 and c < len(self.col_positions) - 1:
                self.select_cell(r, c, redraw = True)
        elif self.multiple_selection_enabled and all(v is None for v in (self.RI.rsz_h, self.RI.rsz_w, self.CH.rsz_h, self.CH.rsz_w)):
            r = self.identify_row(y = event.y)
            c = self.identify_col(x = event.x)
            if r < len(self.row_positions) - 1 and c < len(self.col_positions) - 1:
                self.add_selection(r, c, redraw = True)
        elif self.RI.width_resizing_enabled and self.RI.rsz_h is None and self.RI.rsz_w == True:
            self.RI.currently_resizing_width = True
            self.new_row_width = self.RI.current_width + event.x
            x = self.canvasx(event.x)
            self.create_line(x, y1, x, y2, width = 1, fill = self.RI.resizing_line_color, tag = "rwl")
        elif self.CH.height_resizing_enabled and self.CH.rsz_w is None and self.CH.rsz_h == True:
            self.CH.currently_resizing_height = True
            self.new_header_height = self.CH.current_height + event.y
            y = self.canvasy(event.y)
            self.create_line(x1, y, x2, y, width = 1, fill = self.RI.resizing_line_color, tag = "rhl")
        if self.extra_b1_press_func is not None:
            self.extra_b1_press_func(event)

    def get_max_selected_cell_x(self):
        try:
            return max(self.CH.selected_cells)
        except:
            return None

    def get_max_selected_cell_y(self):
        try:
            return max(self.RI.selected_cells)
        except:
            return None

    def get_min_selected_cell_y(self):
        try:
            return min(self.RI.selected_cells)
        except:
            return None

    def get_min_selected_cell_x(self):
        try:
            return min(self.CH.selected_cells)
        except:
            return None
        
    def b1_motion(self, event):
        x1, y1, x2, y2 = self.get_canvas_visible_area()
        if self.drag_selection_enabled and all(v is None for v in (self.RI.rsz_h, self.RI.rsz_w, self.CH.rsz_h, self.CH.rsz_w)): 
            end_row = self.identify_row(y = event.y)
            end_col = self.identify_col(x = event.x)
            if end_row < len(self.row_positions) - 1 and end_col < len(self.col_positions) - 1 and len(self.currently_selected) == 2:
                if isinstance(self.currently_selected[0], int):
                    start_row = self.currently_selected[0]
                    start_col = self.currently_selected[1]
                    self.selected_cols = set()
                    self.selected_rows = set()
                    self.selected_cells = set()
                    self.CH.selected_cells = defaultdict(int)
                    self.RI.selected_cells = defaultdict(int)
                    if end_row >= start_row and end_col >= start_col:
                        for c in range(start_col, end_col + 1):
                            self.CH.selected_cells[c] += 1
                            for r in range(start_row, end_row + 1):
                                self.RI.selected_cells[r] += 1
                                self.selected_cells.add((r, c))
                    elif end_row >= start_row and end_col < start_col:
                        for c in range(end_col,start_col + 1):
                            self.CH.selected_cells[c] += 1
                            for r in range(start_row,end_row + 1):
                                self.RI.selected_cells[r] += 1
                                self.selected_cells.add((r,c))
                    elif end_row < start_row and end_col >= start_col:
                        for c in range(start_col, end_col + 1):
                            self.CH.selected_cells[c] += 1
                            for r in range(end_row, start_row + 1):
                                self.RI.selected_cells[r] += 1
                                self.selected_cells.add((r, c))
                    elif end_row < start_row and end_col < start_col:
                        for c in range(end_col,start_col + 1):
                            self.CH.selected_cells[c] += 1
                            for r in range(end_row,start_row + 1):
                                self.RI.selected_cells[r] += 1
                                self.selected_cells.add((r, c))
                    if self.drag_selection_binding_func is not None:
                        self.drag_selection_binding_func(sorted([start_row, end_row]) + sorted([start_col, end_col]))
            if event.x > self.winfo_width():
                try:
                    self.xview_scroll(1, "units")
                    self.CH.xview_scroll(1, "units")
                except:
                    pass
            elif event.x < 0:
                try:
                    self.xview_scroll(-1, "units")
                    self.CH.xview_scroll(-1, "units")
                except:
                    pass
            if event.y > self.winfo_height():
                try:
                    self.yview_scroll(1, "units")
                    self.RI.yview_scroll(1, "units")
                except:
                    pass
            elif event.y < 0:
                try:
                    self.yview_scroll(-1, "units")
                    self.RI.yview_scroll(-1, "units")
                except:
                    pass
            self.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
        elif self.RI.width_resizing_enabled and self.RI.rsz_w is not None and self.RI.currently_resizing_width:
            self.RI.delete("rwl")
            self.delete("rwl")
            if event.x >= 0:
                x = self.canvasx(event.x)
                self.new_row_width = self.RI.current_width + event.x
                self.create_line(x, y1, x, y2, width = 1, fill = self.RI.resizing_line_color, tag = "rwl")
            else:
                x = self.RI.current_width + event.x
                if x < self.min_cw:
                    x = int(self.min_cw)
                self.new_row_width = x
                self.RI.create_line(x, y1, x, y2, width = 1, fill = self.RI.resizing_line_color, tag = "rwl")
        elif self.CH.height_resizing_enabled and self.CH.rsz_h is not None and self.CH.currently_resizing_height:
            self.CH.delete("rhl")
            self.delete("rhl")
            if event.y >= 0:
                y = self.canvasy(event.y)
                self.new_header_height = self.CH.current_height + event.y
                self.create_line(x1, y, x2, y, width = 1, fill = self.RI.resizing_line_color, tag = "rhl")
            else:
                y = self.CH.current_height + event.y
                if y < self.hdr_min_rh:
                    y = int(self.hdr_min_rh)
                self.new_header_height = y
                self.CH.create_line(x1, y, x2, y, width = 1, fill = self.RI.resizing_line_color, tag = "rhl")
        
        if self.extra_b1_motion_func is not None:
            self.extra_b1_motion_func(event)
        
    def b1_release(self, event = None):
        if self.RI.width_resizing_enabled and self.RI.rsz_w is not None and self.RI.currently_resizing_width:
            self.delete("rwl")
            self.RI.delete("rwl")
            self.RI.currently_resizing_width = False
            self.RI.set_width(self.new_row_width, set_TL = True)
            self.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
        elif self.CH.height_resizing_enabled and self.CH.rsz_h is not None and self.CH.currently_resizing_height:
            self.delete("rhl")
            self.CH.delete("rhl")
            self.CH.currently_resizing_height = False
            self.CH.set_height(self.new_header_height, set_TL = True)
            self.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
        self.RI.rsz_w = None
        self.CH.rsz_h = None
        self.mouse_motion(event)
        if self.extra_b1_release_func is not None:
            self.extra_b1_release_func(event)

    def double_b1(self, event = None):
        self.focus_set()
        x1, y1, x2, y2 = self.get_canvas_visible_area()
        if self.identify_col(x = event.x, allow_end = False) is None or self.identify_row(y = event.y, allow_end = False) is None:
            self.deselect("all")
        elif self.single_selection_enabled and all(v is None for v in (self.RI.rsz_h, self.RI.rsz_w, self.CH.rsz_h, self.CH.rsz_w)):
            r = self.identify_row(y = event.y)
            c = self.identify_col(x = event.x)
            if r < len(self.row_positions) - 1 and c < len(self.col_positions) - 1:
                self.select_cell(r, c, redraw = True)
        elif self.multiple_selection_enabled and all(v is None for v in (self.RI.rsz_h, self.RI.rsz_w, self.CH.rsz_h, self.CH.rsz_w)):
            r = self.identify_row(y = event.y)
            c = self.identify_col(x = event.x)
            if r < len(self.row_positions) - 1 and c < len(self.col_positions) - 1:
                self.add_selection(r, c, redraw = True)
        if self.extra_double_b1_func is not None:
            self.extra_double_b1_func(event)

    def deselect(self, r = None, c = None, cell = None, redraw = True):
        if r == "all":
            self.selected_cells = set()
            self.currently_selected = tuple()
            self.CH.selected_cells = defaultdict(int)
            self.RI.selected_cells = defaultdict(int)
            self.selected_cols = set()
            self.selected_rows = set()
        if r != "all" and r is not None:
            try: del self.RI.selected_cells[r]
            except: pass
            self.selected_rows.discard(r)
        if c is not None:
            try: del self.CH.selected_cells[c]
            except: pass
            self.selected_cols.discard(c)
        if cell is not None:
            r, c = cell[0], cell[1]
            self.selected_cells.discard(cell)
            self.RI.selected_cells[r] -= 1
            if self.RI.selected_cells[r] < 1:
                del self.RI.selected_cells[r]
            self.CH.selected_cells[c] -= 1
            if self.CH.selected_cells[c] < 1:
                del self.CH.selected_cells[c]
            if cell == self.currently_selected:
                self.currently_selected = tuple()
        if redraw:
            self.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)

    def identify_row(self, event = None, y = None, allow_end = True):
        if event is None:
            y2 = self.canvasy(y)
        elif y is None:
            y2 = self.canvasy(event.y)
        r = bisect.bisect_left(self.row_positions, y2)
        if r != 0:
            r -= 1
        if not allow_end:
            if r >= len(self.row_positions) - 1:
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
        if not allow_end:
            if c >= len(self.col_positions) - 1:
                return None
        return c

    def GetCellCoords(self, event = None, r = None, c = None, sel = False):
        # event takes priority as parameter
        if event is not None:
            r = self.identify_row(event)
            c = self.identify_col(event)
        elif r is not None and c is not None:
            if sel:
                return self.col_positions[c] + 1,self.row_positions[r] + 1, self.col_positions[c + 1], self.row_positions[r + 1]
            else:
                return self.col_positions[c], self.row_positions[r], self.col_positions[c + 1], self.row_positions[r + 1]

    def set_xviews(self, *args):
        self.xview(*args)
        self.CH.xview(*args)
        self.main_table_redraw_grid_and_text(redraw_header = True)

    def set_yviews(self, *args):
        self.yview(*args)
        self.RI.yview(*args)
        self.main_table_redraw_grid_and_text(redraw_row_index = True)

    def set_view(self, x_args, y_args):
        self.xview(*x_args)
        self.CH.xview(*x_args)
        self.yview(*y_args)
        self.RI.yview(*y_args)
        self.main_table_redraw_grid_and_text(redraw_row_index = True, redraw_header = True)

    def mousewheel(self, event = None):
        if event.num == 5 or event.delta == -120:
            self.yview_scroll(1, "units")
            self.RI.yview_scroll(1, "units")
        if event.num == 4 or event.delta == 120:
            if self.canvasy(0) <= 0:
                return
            self.yview_scroll(-1, "units")
            self.RI.yview_scroll(-1, "units")
        self.main_table_redraw_grid_and_text(redraw_row_index = True)

    def GetWidthChars(self, width):
        char_w = self.GetTextWidth("_")
        return int(width / char_w)

    def GetTextWidth(self, txt):
        x = self.txt_measure_canvas.create_text(0, 0, text = txt, font = self.my_font)
        b = self.txt_measure_canvas.bbox(x)
        self.txt_measure_canvas.delete(x)
        return b[2] - b[0]

    def GetTextHeight(self, txt):
        x = self.txt_measure_canvas.create_text(0, 0, text = txt, font = self.my_font)
        b = self.txt_measure_canvas.bbox(x)
        self.txt_measure_canvas.delete(x)
        return b[3] - b[1]

    def GetHdrTextWidth(self, txt):
        x = self.txt_measure_canvas.create_text(0, 0, text = txt, font = self.my_hdr_font)
        b = self.txt_measure_canvas.bbox(x)
        self.txt_measure_canvas.delete(x)
        return b[2] - b[0]

    def GetHdrTextHeight(self, txt):
        x = self.txt_measure_canvas.create_text(0, 0, text = txt, font = self.my_hdr_font)
        b = self.txt_measure_canvas.bbox(x)
        self.txt_measure_canvas.delete(x)
        return b[3] - b[1]

    def set_min_cw(self):
        w1 = self.GetHdrTextWidth("XXXX")
        w2 = self.GetTextWidth("XXXX")
        if w1 >= w2:
            self.min_cw = w1
        else:
            self.min_cw = w2
        if self.min_cw > self.CH.max_cw:
            self.CH.max_cw = self.min_cw * 2
        if self.min_cw > self.default_cw:
            self.default_cw = self.min_cw * 2

    def font(self, newfont = None):
        if newfont:
            if (
                not isinstance(newfont, tuple) or
                not isinstance(newfont[0], str) or
                not isinstance(newfont[1], int)
                ):
                raise ValueError("Parameter must be tuple e.g. ('Arial',12,'normal')")
            if len(newfont) > 2:
                raise ValueError("Parameter must be three-tuple")
            else:
                self.my_font = newfont
            self.fnt_fam = newfont[0]
            self.fnt_sze = newfont[1]
            self.fnt_wgt = newfont[2]
            self.set_fnt_help()
        else:
            return self.my_font

    def set_fnt_help(self):
        self.txt_h = self.GetTextHeight("|ZX*'^")
        self.half_txt_h = ceil(self.txt_h / 2)
        self.fl_ins = self.half_txt_h + 3
        self.xtra_lines_increment = int(self.txt_h)
        self.min_rh = self.txt_h + 6
        if self.min_rh < 12:
            self.min_rh = 12
        self.set_min_cw()
        
    def header_font(self, newfont = None):
        if newfont:
            if (
                not isinstance(newfont, tuple) or
                not isinstance(newfont[0], str) or
                not isinstance(newfont[1], int)
                ):
                raise ValueError("Parameter must be tuple e.g. ('Arial',12,'bold')")
            if len(newfont) == 3:
                if not isinstance(newfont[2], str):
                    raise ValueError("Parameter must be tuple e.g. ('Arial',12,'bold')")
            if len(newfont) > 3:
                raise ValueError("Parameter must be three tuple")
            else:
                self.my_hdr_font = newfont
            self.hdr_fnt_fam = newfont[0]
            self.hdr_fnt_sze = newfont[1]
            self.hdr_fnt_wgt = newfont[2]
            self.set_hdr_fnt_help()
        else:
            return self.my_hdr_font

    def set_hdr_fnt_help(self):
        self.hdr_txt_h = self.GetHdrTextHeight("|ZX*'^")
        self.hdr_half_txt_h = ceil(self.hdr_txt_h / 2)
        self.hdr_fl_ins = self.hdr_half_txt_h + 5
        self.hdr_xtra_lines_increment = self.hdr_txt_h
        self.hdr_min_rh = self.hdr_txt_h + 10
        self.set_min_cw()
        self.CH.set_height(self.GetHdrLinesHeight(self.default_hh))

    def data_reference(self, newdataref = None, total_cols = None, total_rows = None, reset_col_positions = True, reset_row_positions = True, redraw = False):
        if isinstance(newdataref, (list, tuple)):
            self.data_ref = newdataref
            self.undo_storage = deque(maxlen = 20)
            if total_cols is None:
                try:
                    self.total_cols = len(max(newdataref, key = len))
                except:
                    self.total_cols = 0
            elif isinstance(total_cols, int):
                self.total_cols = int(total_cols)
            if isinstance(total_rows, int):
                self.total_rows = int(total_rows)
            else:
                self.total_rows = len(newdataref)
            if reset_col_positions:
                self.reset_col_positions()
            if reset_row_positions:
                self.reset_row_positions()
            if redraw:
                self.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
        else:
            return id(self.data_ref)

    def reset_col_positions(self):
        colpos = int(self.default_cw)
        if self.all_columns_displayed:
            self.col_positions = [0] + list(accumulate(colpos for c in range(self.total_cols)))
        else:
            self.col_positions = [0] + list(accumulate(colpos for c in range(len(self.displayed_columns))))

    def del_col_position(self, idx, deselect_all = False, preserve_other_selections = False):
        # WORK NEEDED FOR PRESERVE SELECTIONS ?
        if deselect_all:
            self.deselect("all", redraw = False)
        if idx == "end" or len(self.col_positions) <= idx + 1:
            del self.col_positions[-1]
        else:
            w = self.col_positions[idx + 1] - self.col_positions[idx]
            idx += 1
            del self.col_positions[idx]
            self.col_positions[idx:] = [e - w for e in islice(self.col_positions, idx, len(self.col_positions))]

    def insert_col_position(self, idx, width = None, deselect_all = False, preserve_other_selections = False):
        # WORK NEEDED FOR PRESERVE SELECTIONS ?
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

    def insert_col_rc(self, event = None):
        if not self.anything_selected():
            stidx = int(self.total_cols)
            posidx = len(self.col_positions) - 1
            if not self.all_columns_displayed:
                stidx = self.displayed_columns[-1] + 1
                self.displayed_columns = [e + 1 if i >= self.displayed_columns[-1] else e for i, e in enumerate(self.displayed_columns)]
                self.displayed_columns.append(int(self.displayed_columns[-1]) + 1)
            self.insert_col_position(idx = posidx,
                                        width = int(self.default_cw),
                                        deselect_all = True,
                                        preserve_other_selections = False)
            if self.my_hdrs and isinstance(self.my_hdrs, list):
                try:
                    self.my_hdrs.insert(stidx, "")
                except:
                    pass
            for rn in range(len(self.data_ref)):
                self.data_ref[rn].insert(stidx, "")
            self.CH.select_col(c = posidx)
            self.total_cols += 1
        else:
            stidx = self.get_min_selected_cell_x()
            if stidx is not None:
                posidx = int(stidx)
                if not self.all_columns_displayed:
                    stidx = int(self.displayed_columns[posidx])
                    self.displayed_columns = [e + 1 if i >= posidx else e for i, e in enumerate(self.displayed_columns)]
                    self.displayed_columns.insert(posidx, int(stidx))
                self.insert_col_position(idx = posidx,
                                            width = int(self.default_cw),
                                            deselect_all = True,
                                            preserve_other_selections = False)
                if self.my_hdrs and isinstance(self.my_hdrs, list):
                    try:
                        self.my_hdrs.insert(stidx, "")
                    except:
                        pass
                for rn in range(len(self.data_ref)):
                    self.data_ref[rn].insert(stidx, "")
                self.CH.select_col(c = posidx)
                self.total_cols += 1
        self.refresh()

    def insert_row_rc(self, event = None):
        if not self.anything_selected():
            stidx = int(self.total_rows)
            posidx = len(self.row_positions) - 1
            #if not self.all_columns_displayed: #work here for subset of rows
            #    stidx = self.displayed_columns[-1] + 1
            #    self.displayed_columns = [e + 1 if i >= self.displayed_columns[-1] else e for i, e in enumerate(self.displayed_columns)]
            #    self.displayed_columns.append(int(self.displayed_columns[-1]) + 1)
            self.insert_row_position(idx = posidx,
                                        height = self.GetLinesHeight(self.default_rh),
                                        deselect_all = True,
                                        preserve_other_selections = False)
            if self.my_row_index and isinstance(self.my_row_index, list):
                try:
                    self.my_row_index.insert(stidx, "")
                except:
                    pass
            self.data_ref.insert(stidx, list(repeat("", self.total_cols)))
            self.RI.select_row(r = posidx)
            self.total_rows += 1
        else:
            stidx = self.get_min_selected_cell_y()
            if stidx is not None:
                posidx = int(stidx)
                #if not self.all_columns_displayed:
                #    stidx = int(self.displayed_columns[posidx])
                #    self.displayed_columns = [e + 1 if i >= posidx else e for i, e in enumerate(self.displayed_columns)]
                #    self.displayed_columns.insert(posidx, int(stidx))
                self.insert_row_position(idx = posidx,
                                        height = self.GetLinesHeight(self.default_rh),
                                        deselect_all = True,
                                        preserve_other_selections = False)
                if self.my_row_index and isinstance(self.my_row_index, list):
                    try:
                        self.my_row_index.insert(stidx, "")
                    except:
                        pass
                self.data_ref.insert(stidx, list(repeat("", self.total_cols)))
                self.RI.select_row(r = posidx)
                self.total_rows += 1
        self.refresh()

    def del_cols_rc(self, event = None):
        seld_cols = self.get_selected_cols()
        if seld_cols:
            if self.all_columns_displayed:
                for rn in range(len(self.data_ref)):
                    for c in reversed(seld_cols):
                        del self.data_ref[rn][c]
                if self.my_hdrs and isinstance(self.my_hdrs, list):
                    for c in reversed(seld_cols):
                        try:
                            del self.my_hdrs[c]
                        except:
                            pass
                for c in reversed(seld_cols):
                    self.del_col_position(c,
                                          deselect_all = False,
                                          preserve_other_selections = False)
                    self.total_cols -= 1
            else:
                for rn in range(len(self.data_ref)):
                    for c in reversed(seld_cols):
                        del self.data_ref[rn][self.displayed_columns[c]]
                if self.my_hdrs and isinstance(self.my_hdrs, list):
                    for c in reversed(seld_cols):
                        try:
                            del self.my_hdrs[self.displayed_columns[c]]
                        except:
                            pass
                for c in reversed(seld_cols):
                    self.del_col_position(c,
                                          deselect_all = False,
                                          preserve_other_selections = False)
                    self.total_cols -= 1
            self.deselect("all", redraw = True)

    def del_rows_rc(self, event = None):
        seld_rows = self.get_selected_rows()
        if seld_rows:
            #if self.all_columns_displayed: work here if subset rows done
            for r in reversed(seld_rows):
                del self.data_ref[r]
            if self.my_row_index and isinstance(self.my_row_index, list):
                for r in reversed(seld_rows):
                    try:
                        del self.my_row_index[r]
                    except:
                        pass
            for r in reversed(seld_rows):
                self.del_row_position(r,
                                      deselect_all = False,
                                      preserve_other_selections = False)
                self.total_rows -= 1
            #else:
                #for rn in range(len(self.data_ref)):
                #    for c in reversed(seld_cols):
                #        del self.data_ref[rn][self.displayed_columns[c]]
                #if self.my_hdrs and isinstance(self.my_hdrs, list):
                #    for c in reversed(seld_cols):
                #        try:
                #            del self.my_hdrs[self.displayed_columns[c]]
                #        except:
                #            pass
                #for c in reversed(seld_cols):
                #    self.del_col_position(c,
                #                          deselect_all = False,
                #                          preserve_other_selections = False)
                #    self.total_cols -= 1
            self.deselect("all", redraw = True)

    def insert_col(self, column = None, idx = "end", width = None, deselect_all = False, preserve_other_selections = False):
        self.insert_col_position(idx = idx,
                                    width = width,
                                    deselect_all = deselect_all,
                                    preserve_other_selections = preserve_other_selections)
        if isinstance(idx, str) and idx.lower() == "end":
            if column is None:
                for rn in range(len(self.data_ref)):
                    self.data_ref[rn].append("")
            else:
                for rn, col_value in zip(range(len(self.data_ref)), column):
                    self.data_ref[rn].append(col_value)
        else:
            if column is None:
                for rn in range(len(self.data_ref)):
                    self.data_ref[rn].insert(idx, "")
            else:
                for rn, col_value in zip(range(len(self.data_ref)), column):
                    self.data_ref[rn].insert(idx, col_value)
        self.total_cols += 1

    def reset_row_positions(self):
        rowpos = self.GetLinesHeight(self.default_rh)
        self.row_positions = [0] + list(accumulate(rowpos for r in range(self.total_rows)))

    def del_row_position(self, idx, deselect_all = False, preserve_other_selections = False):
        # WORK NEEDED FOR PRESERVE SELECTIONS ?
        if deselect_all:
            self.deselect("all", redraw = False)
        if idx == "end" or len(self.row_positions) <= idx + 1:
            del self.row_positions[-1]
        else:
            w = self.row_positions[idx + 1] - self.row_positions[idx]
            idx += 1
            del self.row_positions[idx]
            self.row_positions[idx:] = [e - w for e in islice(self.row_positions, idx, len(self.row_positions))]

    def insert_row_position(self, idx, height = None, deselect_all = False, preserve_other_selections = False):
        # WORK NEEDED FOR PRESERVE SELECTIONS ?
        if deselect_all:
            self.deselect("all", redraw = False)
        if height is None:
            h = self.GetLinesHeight(self.default_rh)
        else:
            h = height
        if idx == "end" or len(self.row_positions) == idx + 1:
            self.row_positions.append(self.row_positions[-1] + h)
        else:
            idx += 1
            self.row_positions.insert(idx, self.row_positions[idx - 1] + h)
            idx += 1
            self.row_positions[idx:] = [e + h for e in islice(self.row_positions, idx, len(self.row_positions))]

    def insert_row(self, row = None, idx = "end", height = None, increase_numcols = True, deselect_all = False, preserve_other_selections = False):
        self.insert_row_position(idx = idx, height = height,
                                    deselect_all = deselect_all,
                                    preserve_other_selections = preserve_other_selections)
        if isinstance(idx, str) and idx.lower() == "end":
            if isinstance(row, list):
                self.data_ref.append(row if row else list(repeat("", len(self.col_positions) - 1)))
            elif row is not None:
                self.data_ref.append([e for e in row])
            elif row is None:
                self.data_ref.append(list(repeat("", len(self.col_positions) - 1)))
        else:
            if isinstance(row, list):
                self.data_ref.insert(idx, row if row else list(repeat("", len(self.col_positions) - 1)))
            elif row is not None:
                self.data_ref.insert(idx, [e for e in row])
            elif row is None:
                self.data_ref.append(list(repeat("", len(self.col_positions) - 1)))
        self.total_rows += 1

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

    def GetLinesHeight(self, n):
        y = int(self.fl_ins)
        if n == 1:
            y += 6
        else:
            for i in range(n):
                y += self.xtra_lines_increment
        if y < self.min_rh:
            y = int(self.min_rh)
        return y

    def GetHdrLinesHeight(self, n):
        y = int(self.hdr_fl_ins)
        if n == 1:
            y + 10
        else:
            for i in range(n):
                y += self.hdr_xtra_lines_increment
        if y < self.hdr_min_rh:
            y = int(self.hdr_min_rh)
        return y

    def display_columns(self, indexes = None, enable = None, reset_col_positions = True, deselect_all = True):
        if deselect_all:
            self.deselect("all")
        if indexes is None and enable is None:
            return tuple(self.displayed_columns)
        if indexes is not None:
            self.displayed_columns = indexes
        if enable == True:
            self.all_columns_displayed = False
        if enable == False:
            self.all_columns_displayed = True
        if reset_col_positions:
            self.reset_col_positions()
                
    def headers(self, newheaders = None, index = None):
        if newheaders is not None:
            if isinstance(newheaders, (list, tuple)):
                if (
                    isinstance(newheaders,(str, int, float)) and
                    isinstance(index, int)
                    ):
                    self.my_hdrs[index] = f"{newheaders}"
                else:
                    self.my_hdrs = newheaders
            elif isinstance(newheaders, int):
                self.my_hdrs = newheaders
        else:
            if index is not None:
                if isinstance(index, int):
                    return self.my_hdrs[index]
            else:
                return self.my_hdrs

    def row_index(self, newindex = None, index = None):
        if isinstance(newindex, int):
            self.my_row_index = newindex
        else:
            if isinstance(newindex, (list, tuple)):
                self.my_row_index = newindex
            else:
                if index is not None:
                    if isinstance(index, int):
                        return self.my_row_index[index]
                else:
                    return self.my_row_index

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

    def main_table_redraw_grid_and_text(self, redraw_header = False, redraw_row_index = False):
        try:
            last_col_line_pos = self.col_positions[-1] + 1
            last_row_line_pos = self.row_positions[-1] + 1
            self.configure(scrollregion=(0, 0, last_col_line_pos + 150, last_row_line_pos + 100))
            self.delete("all")
            x1 = self.canvasx(0)
            y1 = self.canvasy(0)
            x2 = self.canvasx(self.winfo_width())
            y2 = self.canvasy(self.winfo_height())
            self.row_width_resize_bbox = (x1, y1, x1 + 5, y2)
            self.header_height_resize_bbox = (x1 + 6, y1, x2, y1 + 3)
            start_row = bisect.bisect_left(self.row_positions, y1)
            end_row = bisect.bisect_right(self.row_positions, y2)
            if not y2 >= self.row_positions[-1]:
                end_row += 1
            start_col = bisect.bisect_left(self.col_positions, x1)
            end_col = bisect.bisect_right(self.col_positions, x2)
            if not x2 >= self.col_positions[-1]:
                end_col += 1
            if last_col_line_pos > x2:
                x_stop = x2
            else:
                x_stop = last_col_line_pos
            if last_row_line_pos > y2:
                y_stop = y2
            else:
                y_stop = last_row_line_pos
            cr_ = self.create_rectangle
            ct_ = self.create_text
            sb = y2 + 2
            if start_row > 0:
                selsr = start_row - 1
            else:
                selsr = start_row
            if start_col > 0:
                selsc = start_col - 1
            else:
                selsc = start_col
            if not self.selected_cells:
                if self.selected_rows:
                    for r in range(selsr, end_row - 1):
                        fr = self.row_positions[r]
                        sr = self.row_positions[r + 1]
                        if sr > sb:
                            sr = sb
                        if r in self.selected_rows:
                            cr_(x1, fr + 1, x_stop, sr, fill = self.selected_cells_background, outline = "")
                elif self.selected_cols:
                    for c in range(selsc, end_col - 1):
                        fc = self.col_positions[c]
                        sc = self.col_positions[c + 1]
                        if sc > x2 + 2:
                            sc = x2 + 2
                        if c in self.selected_cols:
                            cr_(fc + 1, y1, sc, y_stop, fill = self.selected_cells_background, outline = "")
            for r in range(start_row - 1, end_row):
                y = self.row_positions[r]
                self.create_line(x1, y, x_stop, y, fill= self.grid_color, width = 1)
            for c in range(start_col - 1, end_col):
                x = self.col_positions[c]
                self.create_line(x, y1, x, y_stop, fill = self.grid_color, width = 1)
            if start_row > 0:
                start_row -= 1
            if start_col > 0:
                start_col -= 1   
            end_row -= 1
            c_2 = self.selected_cells_background if self.selected_cells_background.startswith("#") else Color_Map_[self.selected_cells_background]
            c_2_ = (int(c_2[1:3], 16), int(c_2[3:5], 16), int(c_2[5:], 16))
            rows_ = tuple(range(start_row, end_row))
            if self.all_columns_displayed:
                if self.align == "w":
                    for c in range(start_col, end_col - 1):
                        fc = self.col_positions[c]
                        sc = self.col_positions[c + 1]
                        x = fc + 5
                        mw = sc - fc - 5
                        for r in rows_:
                            fr = self.row_positions[r]
                            sr = self.row_positions[r + 1]
                            if sr > sb:
                                sr = sb
                            if (r, c) in self.highlighted_cells and ((r, c) in self.selected_cells or r in self.selected_rows or c in self.selected_cols):
                                c_1 = self.highlighted_cells[(r, c)][0] if self.highlighted_cells[(r, c)][0].startswith("#") else Color_Map_[self.highlighted_cells[(r, c)][0]]
                                cr_(fc + 1,
                                    fr + 1,
                                    sc,
                                    sr,
                                    fill = (f"#{int((int(c_1[1:3], 16) + c_2_[0]) / 2):02X}" +
                                            f"{int((int(c_1[3:5], 16) + c_2_[1]) / 2):02X}" +
                                            f"{int((int(c_1[5:], 16) + c_2_[2]) / 2):02X}"),
                                    outline = "")
                                tf = self.selected_cells_foreground if self.highlighted_cells[(r, c)][1] is None else self.highlighted_cells[(r, c)][1]
                            elif (r, c) in self.selected_cells:
                                cr_(fc + 1, fr + 1, sc, sr, fill = self.selected_cells_background, outline = "")
                                tf = self.selected_cells_foreground
                            elif (r, c) in self.highlighted_cells and r not in self.selected_rows and c not in self.selected_cols:
                                cr_(fc + 1, fr + 1, sc, sr, fill = self.highlighted_cells[(r, c)][0], outline = "")
                                tf = self.text_color if self.highlighted_cells[(r, c)][1] is None else self.highlighted_cells[(r, c)][1]
                            elif r in self.selected_rows or c in self.selected_cols:
                                tf = self.selected_cells_foreground
                            else:
                                tf = self.text_color
                            if x > x2:
                                continue
                            try:
                                lns = self.data_ref[r][c]
                                if isinstance(lns, str):
                                    lns = lns.split("\n")
                                else:
                                    lns = (f"{lns}",)
                                y = fr + self.fl_ins
                                if y + self.half_txt_h > y1:
                                    fl = lns[0]
                                    t = ct_(x, y, text = fl, fill = tf, font = self.my_font, anchor = "w")
                                    wd = self.bbox(t)
                                    wd = wd[2] - wd[0]
                                    if wd > mw:
                                        #nl = int(mw * (len(fl) / wd)) - 1
                                        nl = int(len(fl) * (mw / wd)) - 1
                                        self.itemconfig(t,text=fl[:nl])
                                        wd = self.bbox(t)
                                        while wd[2] - wd[0] > mw:
                                            nl -= 1
                                            self.dchars(t,nl)
                                            wd = self.bbox(t)
                                if len(lns) > 1:
                                    stl = int((y1 - y) / self.xtra_lines_increment) - 1
                                    if stl < 1:
                                        stl = 1
                                    y += (stl * self.xtra_lines_increment)
                                    if y + self.half_txt_h < sr:
                                        for i in range(stl, len(lns)):
                                            txt = lns[i]
                                            t = ct_(x, y, text = txt, fill = tf, font = self.my_font, anchor = "w")
                                            wd = self.bbox(t)
                                            wd = wd[2] - wd[0]
                                            if wd > mw:
                                                #nl = int(mw * (len(txt) / wd)) - 1
                                                nl = int(len(txt) * (mw / wd)) - 1
                                                self.itemconfig(t,text=txt[:nl])
                                                wd = self.bbox(t)
                                                while wd[2] - wd[0] > mw:
                                                    nl -= 1
                                                    self.dchars(t,nl)
                                                    wd = self.bbox(t)
                                            y += self.xtra_lines_increment
                                            if y + self.half_txt_h > sr:
                                                break
                            except:
                                continue
                elif self.align == "center":
                    for c in range(start_col, end_col - 1):
                        fc = self.col_positions[c]
                        stop = fc + 5
                        sc = self.col_positions[c + 1]
                        mw = sc - fc - 5
                        x = fc + floor(mw / 2)
                        for r in rows_:
                            fr = self.row_positions[r]
                            sr = self.row_positions[r + 1]
                            if sr > sb:
                                sr = sb
                            if (r, c) in self.highlighted_cells and ((r, c) in self.selected_cells or r in self.selected_rows or c in self.selected_cols):
                                c_1 = self.highlighted_cells[(r, c)][0] if self.highlighted_cells[(r, c)][0].startswith("#") else Color_Map_[self.highlighted_cells[(r, c)][0]]
                                cr_(fc + 1,
                                    fr + 1,
                                    sc,
                                    sr,
                                    fill = (f"#{int((int(c_1[1:3], 16) + c_2_[0]) / 2):02X}" +
                                            f"{int((int(c_1[3:5], 16) + c_2_[1]) / 2):02X}" +
                                            f"{int((int(c_1[5:], 16) + c_2_[2]) / 2):02X}"),
                                    outline = "")
                                tf = self.selected_cells_foreground if self.highlighted_cells[(r, c)][1] is None else self.highlighted_cells[(r, c)][1]
                            elif (r, c) in self.selected_cells:
                                cr_(fc + 1, fr + 1, sc, sr, fill = self.selected_cells_background, outline = "")
                                tf = self.selected_cells_foreground
                            elif (r, c) in self.highlighted_cells and r not in self.selected_rows and c not in self.selected_cols:
                                cr_(fc + 1, fr + 1, sc, sr, fill = self.highlighted_cells[(r, c)][0], outline = "")
                                tf = self.text_color if self.highlighted_cells[(r, c)][1] is None else self.highlighted_cells[(r, c)][1]
                            elif r in self.selected_rows or c in self.selected_cols:
                                tf = self.selected_cells_foreground
                            else:
                                tf = self.text_color
                            if stop > x2:
                                continue
                            try:
                                lns = self.data_ref[r][c]
                                if isinstance(lns, str):
                                    lns = lns.split("\n")
                                else:
                                    lns = (f"{lns}", )
                                fl = lns[0]
                                y = fr + self.fl_ins
                                if y + self.half_txt_h > y1:
                                    t = ct_(x, y, text = fl, fill = tf, font = self.my_font, anchor = "center")
                                    wd = self.bbox(t)
                                    wd = wd[2] - wd[0]
                                    if wd > mw:
                                        tl = len(fl)
                                        slce = tl - floor(tl * (mw / wd))
                                        if slce % 2:
                                            slce += 1
                                        else:
                                            slce += 2
                                        slce = int(slce / 2)
                                        fl = fl[slce:tl - slce]
                                        self.itemconfig(t, text = fl)
                                        wd = self.bbox(t)
                                        while wd[2] - wd[0] > mw:
                                            fl = fl[1: - 1]
                                            self.itemconfig(t, text = fl)
                                            wd = self.bbox(t)
                                if len(lns) > 1:
                                    stl = int((y1 - y) / self.xtra_lines_increment) - 1
                                    if stl < 1:
                                        stl = 1
                                    y += (stl * self.xtra_lines_increment)
                                    if y + self.half_txt_h < sr:
                                        for i in range(stl,len(lns)):
                                            txt = lns[i]
                                            t = ct_(x, y, text = txt, fill = tf, font = self.my_font, anchor = "center")
                                            wd = self.bbox(t)
                                            wd = wd[2] - wd[0]
                                            if wd > mw:
                                                tl = len(txt)
                                                slce = tl - floor(tl * (mw / wd))
                                                if slce % 2:
                                                    slce += 1
                                                else:
                                                    slce += 2
                                                slce = int(slce / 2)
                                                txt = txt[slce:tl - slce]
                                                self.itemconfig(t, text = txt)
                                                wd = self.bbox(t)
                                                while wd[2] - wd[0] > mw:
                                                    txt = txt[1: - 1]
                                                    self.itemconfig(t, text = txt)
                                                    wd = self.bbox(t)
                                            y += self.xtra_lines_increment
                                            if y + self.half_txt_h > sr:
                                                break
                            except:
                                continue
            else:
                if self.align == "w":
                    for c in range(start_col, end_col - 1):
                        fc = self.col_positions[c]
                        sc = self.col_positions[c + 1]
                        x = fc + 5
                        mw = sc - fc - 5
                        for r in rows_:
                            fr = self.row_positions[r]
                            sr = self.row_positions[r + 1]
                            if sr > sb:
                                sr = sb
                            if (r, self.displayed_columns[c]) in self.highlighted_cells and ((r, c) in self.selected_cells or r in self.selected_rows or c in self.selected_cols):
                                c_1 = self.highlighted_cells[(r, self.displayed_columns[c])][0] if self.highlighted_cells[(r, self.displayed_columns[c])][0].startswith("#") else Color_Map_[self.highlighted_cells[(r, self.displayed_columns[c])][0]]
                                cr_(fc + 1,
                                    fr + 1,
                                    sc,
                                    sr,
                                    fill = (f"#{int((int(c_1[1:3], 16) + c_2_[0]) / 2):02X}" +
                                            f"{int((int(c_1[3:5], 16) + c_2_[1]) / 2):02X}" +
                                            f"{int((int(c_1[5:], 16) + c_2_[2]) / 2):02X}"),
                                    outline = "")
                                tf = self.selected_cells_foreground if self.highlighted_cells[(r, self.displayed_columns[c])][1] is None else self.highlighted_cells[(r, self.displayed_columns[c])][1]
                            elif (r, c) in self.selected_cells:
                                cr_(fc + 1, fr + 1, sc, sr, fill = self.selected_cells_background, outline = "")
                                tf = self.selected_cells_foreground
                            elif (r, self.displayed_columns[c]) in self.highlighted_cells and r not in self.selected_rows and c not in self.selected_cols:
                                cr_(fc + 1, fr + 1, sc, sr, fill = self.highlighted_cells[(r, self.displayed_columns[c])][0], outline = "")
                                tf = self.text_color if self.highlighted_cells[(r, self.displayed_columns[c])][1] is None else self.highlighted_cells[(r, self.displayed_columns[c])][1]
                            elif r in self.selected_rows or c in self.selected_cols:
                                tf = self.selected_cells_foreground
                            else:
                                tf = self.text_color
                            if x > x2:
                                continue
                            try:
                                lns = self.data_ref[r][self.displayed_columns[c]]
                                if isinstance(lns, str):
                                    lns = lns.split("\n")
                                else:
                                    lns = (f"{lns}", )
                                y = fr + self.fl_ins
                                if y + self.half_txt_h > y1:
                                    fl = lns[0]
                                    t = ct_(x, y, text = fl, fill = tf, font = self.my_font, anchor = "w")
                                    wd = self.bbox(t)
                                    wd = wd[2] - wd[0]
                                    if wd > mw:
                                        nl = int(len(fl) * (mw / wd)) - 1
                                        self.itemconfig(t, text = fl[:nl])
                                        wd = self.bbox(t)
                                        while wd[2] - wd[0] > mw:
                                            nl -= 1
                                            self.dchars(t, nl)
                                            wd = self.bbox(t)
                                if len(lns) > 1:
                                    stl = int((y1 - y) / self.xtra_lines_increment) - 1
                                    if stl < 1:
                                        stl = 1
                                    y += (stl * self.xtra_lines_increment)
                                    if y + self.half_txt_h < sr:
                                        for i in range(stl, len(lns)):
                                            txt = lns[i]
                                            t = ct_(x, y, text = txt, fill = tf, font = self.my_font, anchor = "w")
                                            wd = self.bbox(t)
                                            wd = wd[2] - wd[0]
                                            if wd > mw:
                                                nl = int(len(txt) * (mw / wd)) - 1
                                                self.itemconfig(t,text=txt[:nl])
                                                wd = self.bbox(t)
                                                while wd[2] - wd[0] > mw:
                                                    nl -= 1
                                                    self.dchars(t,nl)
                                                    wd = self.bbox(t)
                                            y += self.xtra_lines_increment
                                            if y + self.half_txt_h > sr:
                                                break
                            except:
                                continue
                elif self.align == "center":
                    for c in range(start_col, end_col - 1):
                        fc = self.col_positions[c]
                        stop = fc + 5
                        sc = self.col_positions[c + 1]
                        mw = sc - fc - 5
                        x = fc + floor(mw / 2)
                        for r in rows_:
                            fr = self.row_positions[r]
                            sr = self.row_positions[r + 1]
                            if sr > sb:
                                sr = sb
                            if (r, self.displayed_columns[c]) in self.highlighted_cells and ((r, c) in self.selected_cells or r in self.selected_rows or c in self.selected_cols):
                                c_1 = self.highlighted_cells[(r, self.displayed_columns[c])][0] if self.highlighted_cells[(r, self.displayed_columns[c])][0].startswith("#") else Color_Map_[self.highlighted_cells[(r, self.displayed_columns[c])][0]]
                                cr_(fc + 1,
                                    fr + 1,
                                    sc,
                                    sr,
                                    fill = (f"#{int((int(c_1[1:3], 16) + c_2_[0]) / 2):02X}" +
                                            f"{int((int(c_1[3:5], 16) + c_2_[1]) / 2):02X}" +
                                            f"{int((int(c_1[5:], 16) + c_2_[2]) / 2):02X}"),
                                    outline = "")
                                tf = self.selected_cells_foreground if self.highlighted_cells[(r, self.displayed_columns[c])][1] is None else self.highlighted_cells[(r, self.displayed_columns[c])][1]
                            elif (r, c) in self.selected_cells:
                                cr_(fc + 1, fr + 1, sc, sr, fill = self.selected_cells_background, outline = "")
                                tf = self.selected_cells_foreground
                            elif (r, self.displayed_columns[c]) in self.highlighted_cells and r not in self.selected_rows and c not in self.selected_cols:
                                cr_(fc + 1, fr + 1, sc, sr, fill = self.highlighted_cells[(r, self.displayed_columns[c])][0], outline = "")
                                tf = self.text_color if self.highlighted_cells[(r, self.displayed_columns[c])][1] is None else self.highlighted_cells[(r, self.displayed_columns[c])][1]
                            elif r in self.selected_rows or c in self.selected_cols:
                                tf = self.selected_cells_foreground
                            else:
                                tf = self.text_color
                            if stop > x2:
                                continue
                            try:
                                lns = self.data_ref[r][self.displayed_columns[c]]
                                if isinstance(lns, str):
                                    lns = lns.split("\n")
                                else:
                                    lns = (f"{lns}", )
                                fl = lns[0]
                                y = fr + self.fl_ins
                                if y + self.half_txt_h > y1:
                                    t = ct_(x, y, text = fl, fill = tf, font = self.my_font, anchor = "center")
                                    wd = self.bbox(t)
                                    wd = wd[2] - wd[0]
                                    if wd > mw:
                                        tl = len(fl)
                                        slce = tl - floor(tl * (mw / wd))
                                        if slce % 2:
                                            slce += 1
                                        else:
                                            slce += 2
                                        slce = int(slce / 2)
                                        fl = fl[slce:tl - slce]
                                        self.itemconfig(t, text = fl)
                                        wd = self.bbox(t)
                                        while wd[2] - wd[0] > mw:
                                            fl = fl[1:-1]
                                            self.itemconfig(t, text = fl)
                                            wd = self.bbox(t)
                                if len(lns) > 1:
                                    stl = int((y1 - y) / self.xtra_lines_increment) - 1
                                    if stl < 1:
                                        stl = 1
                                    y += (stl * self.xtra_lines_increment)
                                    if y + self.half_txt_h < sr:
                                        for i in range(stl, len(lns)):
                                            txt = lns[i]
                                            t = ct_(x, y, text = txt, fill = tf, font = self.my_font, anchor = "center")
                                            wd = self.bbox(t)
                                            wd = wd[2] - wd[0]
                                            if wd > mw:
                                                tl = len(txt)
                                                slce = tl - floor(tl * (mw / wd))
                                                if slce % 2:
                                                    slce += 1
                                                else:
                                                    slce += 2
                                                slce = int(slce / 2)
                                                txt = txt[slce:tl - slce]
                                                self.itemconfig(t, text = txt)
                                                wd = self.bbox(t)
                                                while wd[2] - wd[0] > mw:
                                                    txt = txt[1:-1]
                                                    self.itemconfig(t, text = txt)
                                                    wd = self.bbox(t)
                                            y += self.xtra_lines_increment
                                            if y + self.half_txt_h > sr:
                                                break
                            except:
                                continue
        except:
            return
        if redraw_header:
            self.CH.redraw_grid_and_text(last_col_line_pos, x1, x_stop, start_col, end_col)
        if redraw_row_index:
            self.RI.redraw_grid_and_text(last_row_line_pos, y1, y_stop, start_row, end_row + 1, y2, x1, x_stop)

    def GetColCoords(self, c, sel = False):
        last_col_line_pos = self.col_positions[-1] + 1
        last_row_line_pos = self.row_positions[-1] + 1
        x1 = self.col_positions[c]
        x2 = self.col_positions[c + 1]
        y1 = self.canvasy(0)
        y2 = self.canvasy(self.winfo_height())
        if last_row_line_pos < y2:
            y2 = last_col_line_pos
        if sel:
            return x1, y1 + 1, x2, y2
        else:
            return x1, y1, x2, y2

    def GetRowCoords(self, r, sel = False):
        last_col_line_pos = self.col_positions[-1] + 1
        x1 = self.canvasx(0)
        x2 = self.canvasx(self.winfo_width())
        if last_col_line_pos < x2:
            x2 = last_col_line_pos
        y1 = self.row_positions[r]
        y2 = self.row_positions[r + 1]
        if sel:
            return x1, y1 + 1, x2, y2
        else:
            return x1, y1, x2, y2

    def get_selected_rows(self, get_cells = False):
        if get_cells:
            return sorted(set([cll[0] for cll in self.selected_cells] + list(self.selected_rows)))
        return sorted(self.selected_rows)

    def get_selected_cols(self, get_cells = False):
        if get_cells:
            return sorted(set([cll[0] for cll in self.selected_cells] + list(self.selected_cols)))
        return sorted(self.selected_cols)

    def anything_selected(self, exclude_columns = False, exclude_rows = False, exclude_cells = False):
        if exclude_columns and exclude_rows and not exclude_cells:
            if self.selected_cells:
                return True
        elif exclude_columns and exclude_cells and not exclude_rows:
            if self.selected_rows:
                return True
        elif exclude_rows and exclude_cells and not exclude_columns:
            if self.selected_cols:
                return True
            
        elif exclude_columns and not exclude_rows and not exclude_cells:
            if self.selected_rows or self.selected_cells:
                return True
        elif exclude_rows and not exclude_columns and not exclude_cells:
            if self.selected_cols or self.selected_cells:
                return True
        elif exclude_cells and not exclude_columns and not exclude_rows:
            if self.selected_cols or self.selected_rows:
                return True
            
        elif not exclude_columns and not exclude_rows and not exclude_cells:
            if self.selected_cols or self.selected_rows or self.selected_cells:
                return True
        return False

    def get_selected_cells(self, get_rows = False, get_cols = False):
        res = []
        # IMPROVE SPEED
        if get_cols:
            if self.selected_cols:
                for c in self.selected_cols:
                    for r in range(len(self.row_positions) - 1):
                        res.append((r, c))
        if get_rows:
            if self.selected_rows:
                for r in self.selected_rows:
                    for c in range(len(self.col_positions) - 1):
                        res.append((r, c))
        return res + list(self.selected_cells)

    def edit_cell_(self, event = None):
        if not self.anything_selected():
            return
        if not self.selected_cells:
            return
        x1 = self.get_min_selected_cell_x()
        y1 = self.get_min_selected_cell_y()
        if event.char in all_chars:
            text = event.char
        else:
            text = self.data_ref[y1][x1]
        self.select_cell(r = y1, c = x1)
        self.see(r = y1, c = x1, keep_yscroll = False, keep_xscroll = False, bottom_right_corner = False, check_cell_visibility = True)
        self.RI.set_row_height(y1, only_set_if_too_small = True)
        self.CH.set_col_width(x1, only_set_if_too_small = True)
        self.refresh()
        self.create_text_editor(r = y1, c = x1, text = text, set_data_ref_on_destroy = True)   
        
    def create_text_editor(self, r = 0, c = 0, text = None, state = "normal", see = True, set_data_ref_on_destroy = False):
        if see:
            self.see(r = r, c = c, check_cell_visibility = True)
        x = self.col_positions[c]
        y = self.row_positions[r]
        w = self.col_positions[c + 1] - x
        h = self.row_positions[r + 1] - y
        if text is None:
            text = ""
        self.text_editor = TextEditor(self, text = text, font = self.my_font, state = state, width = w, height = h)
        self.text_editor_id = self.create_window((x, y), window = self.text_editor, anchor = "nw")
        self.text_editor.textedit.bind("<Alt-Return>", self.text_editor_newline_binding)
        if set_data_ref_on_destroy:
            self.text_editor.textedit.bind("<Return>", lambda x: self.get_text_editor_value((r, c)))
            self.text_editor.textedit.bind("<FocusOut>", lambda x: self.get_text_editor_value((r, c)))
            self.text_editor.textedit.focus_set()

    def bind_text_editor_destroy(self, binding, r, c):
        self.text_editor.textedit.bind("<Return>", lambda x: binding((r, c)))
        self.text_editor.textedit.bind("<FocusOut>", lambda x: binding((r, c)))
        self.text_editor.textedit.focus_set()

    def destroy_text_editor(self):
        try:
            self.delete(self.text_editor_id)
            self.text_editor.destroy()
            self.text_editor_id = None
        except:
            pass
        self.text_editor = None

    def text_editor_newline_binding(self, event = None):
        self.text_editor.config(height = self.text_editor.winfo_height() + self.xtra_lines_increment)

    def get_text_editor_value(self, destroy_tup = None, r = None, c = None, set_data_ref_on_destroy = True, event = None, destroy = True, move_down = True):
        if self.text_editor is not None:
            self.text_editor_value = self.text_editor.get()
        if destroy:
            try:
                self.delete(self.text_editor_id)
                self.text_editor.destroy()
                self.text_editor_id = None
            except:
                pass
            self.text_editor = None
        if set_data_ref_on_destroy:
            if r is None and c is None and destroy_tup:
                r, c = destroy_tup[0], destroy_tup[1]
            if self.undo_enabled:
                self.undo_storage.append(zlib.compress(pickle.dumps(("edit_cells", {(r, c): f"{self.data_ref[r][c]}"}))))
            self.data_ref[r][c] = self.text_editor_value
            self.RI.set_row_height(r)
            self.CH.set_col_width(c, only_set_if_too_small = True)
            if self.extra_edit_cell_func is not None:
                self.extra_edit_cell_func((r, c))
        if move_down:
            self.arrowkey_DOWN()
        self.refresh()
        self.focus_set()
        return self.text_editor_value

    def create_dropdown(self, r = 0, c = 0, values = [], set_value = None, state = "readonly", see = True, destroy_on_select = True, current = False):
        if see:
            if not self.cell_is_completely_visible(r = r, c = c):
                self.see(r = r, c = c, check_cell_visibility = False)
        x = self.col_positions[c]
        y = self.row_positions[r]
        w = self.GetWidthChars(self.col_positions[c + 1] - x)
        self.table_dropdown = TableDropdown(self, font = self.my_font, state = state, values = values, set_value = set_value, width = w)
        self.table_dropdown_id = self.create_window((x, y), window = self.table_dropdown, anchor = "nw")
        if destroy_on_select:
            self.table_dropdown.bind("<<ComboboxSelected>>", lambda event: self.get_dropdown_value(current = current))

    def get_dropdown_value(self, event = None, current = False, destroy = True):
        if self.table_dropdown is not None:
            if current:
                self.table_dropdown_value = self.table_dropdown.current()
            else:
                self.table_dropdown_value = self.table_dropdown.get_my_value()
        if destroy:
            try:
                self.delete(self.table_dropdown_id)
                self.table_dropdown.destroy()
            except:
                pass
            self.table_dropdown = None
        return self.table_dropdown_value
    

