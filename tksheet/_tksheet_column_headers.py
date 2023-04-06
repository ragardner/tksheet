from ._tksheet_vars import *
from ._tksheet_other_classes import *

from itertools import islice, accumulate, chain, cycle, repeat
from collections import defaultdict
from math import floor, ceil
import bisect
import pickle
import tkinter as tk
import zlib


class ColumnHeaders(tk.Canvas):
    def __init__(self,
                 *args,
                 **kwargs):
        tk.Canvas.__init__(self,
                           kwargs['parentframe'],
                           background = kwargs['header_bg'],
                           highlightthickness = 0)
        self.parentframe = kwargs['parentframe']
        self.current_height = None    # is set from within MainTable() __init__ or from Sheet parameters
        self.MT = None         # is set from within MainTable() __init__
        self.RI = None    # is set from within MainTable() __init__
        self.TL = None                # is set from within TopLeftRectangle() __init__
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
        
        self.column_drag_and_drop_perform = kwargs['column_drag_and_drop_perform']
        self.default_hdr = kwargs['default_header'].lower()
        self.max_cw = float(kwargs['max_colwidth'])
        self.max_header_height = float(kwargs['max_header_height'])
        self.header_bg = kwargs['header_bg']
        self.header_fg = kwargs['header_fg']
        self.header_grid_fg = kwargs['header_grid_fg']
        self.header_border_fg = kwargs['header_border_fg']
        self.header_selected_cells_bg = kwargs['header_selected_cells_bg']
        self.header_selected_cells_fg = kwargs['header_selected_cells_fg']
        self.header_selected_columns_bg = kwargs['header_selected_columns_bg']
        self.header_selected_columns_fg = kwargs['header_selected_columns_fg']
        self.header_hidden_columns_expander_bg = kwargs['header_hidden_columns_expander_bg']
        self.show_default_header_for_empty = kwargs['show_default_header_for_empty']
        self.drag_and_drop_bg = kwargs['drag_and_drop_bg']
        self.resizing_line_fg = kwargs['resizing_line_fg']
        self.align = kwargs['header_align']
        self.basic_bindings()
        
    def basic_bindings(self, enable = True):
        if enable:
            self.bind("<Motion>", self.mouse_motion)
            self.bind("<ButtonPress-1>", self.b1_press)
            self.bind("<B1-Motion>", self.b1_motion)
            self.bind("<ButtonRelease-1>", self.b1_release)
            self.bind("<Double-Button-1>", self.double_b1)
            self.bind(get_rc_binding(), self.rc)
            self.bind("<MouseWheel>", self.mousewheel)
        else:
            self.unbind("<Motion>")
            self.unbind("<ButtonPress-1>")
            self.unbind("<B1-Motion>")
            self.unbind("<ButtonRelease-1>")
            self.unbind("<Double-Button-1>")
            self.unbind(get_rc_binding())
            self.unbind("<MouseWheel>")
            
    def mousewheel(self, event = None):
        maxlines = 0
        if isinstance(self.MT._headers, int):
            if len(self.MT.data) > self.MT._headers:
                maxlines = max(len(e.rstrip().split("\n")) if isinstance(e, str) else len(f"{e}".rstrip().split("\n")) for e in self.MT.data[self.MT._headers])
        elif self.MT._headers:
            maxlines = max(len(e.rstrip().split("\n")) if isinstance(e, str) else len(f"{e}".rstrip().split("\n")) for e in self.MT._headers)
        if maxlines == 1:
            maxlines = 0
        if self.lines_start_at > maxlines:
            self.lines_start_at = maxlines
        if (event.delta < 0 or event.num == 5) and self.lines_start_at < maxlines:
            self.lines_start_at += 1
        elif (event.delta >= 0 or event.num == 4) and self.lines_start_at > 0:
            self.lines_start_at -= 1
        self.MT.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = False, redraw_table = False)

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

    def check_mouse_position_width_resizers(self, x, y):
        for c, (x1, y1, x2, y2) in self.visible_col_dividers.items():
            if (x >= x1 and
                y >= y1 and
                x <= x2 and
                y <= y2):
                return c

    def rc(self, event):
        self.MT.mouseclick_outside_editor_or_dropdown()
        self.mouseclick_outside_editor_or_dropdown()
        self.focus_set()
        popup_menu = None
        if self.MT.identify_col(x = event.x, allow_end = False) is None:
            self.MT.deselect("all")
            if self.MT.rc_popup_menus_enabled:
                popup_menu = self.MT.empty_rc_popup_menu
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
        self.MT.mouseclick_outside_editor_or_dropdown()
        self.mouseclick_outside_editor_or_dropdown()
        x = event.x
        c = self.MT.identify_col(x = x)
        if self.drag_and_drop_enabled or self.col_selection_enabled and self.rsz_h is None and self.rsz_w is None:
            if c < len(self.MT.col_positions) - 1:
                c_selected = self.MT.col_selected(c)
                if not c_selected and self.col_selection_enabled:
                    c = int(c)
                    currently_selected = self.MT.currently_selected()
                    if currently_selected and currently_selected.type_ == "column":
                        min_c = int(currently_selected[1])
                        self.MT.delete_selection_rects(delete_current = False)
                        if c > min_c:
                            self.MT.create_selected(0, min_c, len(self.MT.row_positions) - 1, c + 1, "columns")
                            func_event = tuple(range(min_c, c + 1))
                        elif c < min_c:
                            self.MT.create_selected(0, c, len(self.MT.row_positions) - 1, min_c + 1, "columns")
                            func_event = tuple(range(c, min_c + 1))
                    else:
                        self.select_col(c)
                        func_event = (c, )
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
            mouse_over_selected = False
            if self.width_resizing_enabled:
                c = self.check_mouse_position_width_resizers(x, y)
                if c is not None:
                    self.rsz_w, mouse_over_resize = c, True
                    self.config(cursor = "sb_h_double_arrow")
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
                if self.MT.col_selected(self.MT.identify_col(event, allow_end = False)):
                    self.config(cursor = "hand2")
                    mouse_over_selected = True
            if not mouse_over_resize and not mouse_over_selected:
                self.MT.reset_mouse_motion_creations()
        if self.extra_motion_func is not None:
            self.extra_motion_func(event)

    def double_b1(self, event = None):
        self.MT.mouseclick_outside_editor_or_dropdown()
        self.mouseclick_outside_editor_or_dropdown()
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
        self.MT.unbind("<MouseWheel>")
        self.focus_set()
        self.MT.mouseclick_outside_editor_or_dropdown()
        self.closed_dropdown = self.mouseclick_outside_editor_or_dropdown()
        x = self.canvasx(event.x)
        y = self.canvasy(event.y)
        c = self.MT.identify_col(x = event.x)
        self.b1_pressed_loc = c
        if self.check_mouse_position_width_resizers(x, y) is None:
            self.rsz_w = None
        if self.width_resizing_enabled and self.rsz_w is not None:
            x1, y1, x2, y2 = self.MT.get_canvas_visible_area()
            self.currently_resizing_width = True
            x = self.MT.col_positions[self.rsz_w]
            line2x = self.MT.col_positions[self.rsz_w - 1]
            self.create_resize_line(x, 0, x, self.current_height, width = 1, fill = self.resizing_line_fg, tag = "rwl")
            self.MT.create_resize_line(x, y1, x, y2, width = 1, fill = self.resizing_line_fg, tag = "rwl")
            self.create_resize_line(line2x, 0, line2x, self.current_height,width = 1, fill = self.resizing_line_fg, tag = "rwl2")
            self.MT.create_resize_line(line2x, y1, line2x, y2, width = 1, fill = self.resizing_line_fg, tag = "rwl2")
        elif self.height_resizing_enabled and self.rsz_w is None and self.rsz_h is not None:
            x1, y1, x2, y2 = self.MT.get_canvas_visible_area()
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
                if self.MT.col_selected(c):
                    dcol = c if self.MT.all_columns_displayed else self.MT.displayed_columns[c]
                    if ((dcol in self.cell_options and 'dropdown' in self.cell_options[dcol] and x < self.MT.col_positions[c + 1] and x > self.MT.col_positions[c + 1] - self.MT.hdr_txt_h - 4) or
                        (dcol in self.cell_options and 'checkbox' in self.cell_options[dcol] and x < self.MT.col_positions[c] + self.MT.hdr_txt_h + 5)) and event.y < self.MT.hdr_txt_h + 5:
                        pass
                    else:
                        self.dragged_col = c
                else:
                    self.being_drawn_rect = (0, c, len(self.MT.row_positions) - 1, c + 1, "columns")
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
            if size >= self.MT.min_cw and size < self.max_cw:
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
            need_redraw = False
            end_col = self.MT.identify_col(x = event.x)
            currently_selected = self.MT.currently_selected()
            if end_col < len(self.MT.col_positions) - 1 and currently_selected:
                if currently_selected.type_ == "column":
                    start_col = currently_selected[1]
                    if end_col >= start_col:
                        rect = (0, start_col, len(self.MT.row_positions) - 1, end_col + 1, "columns")
                        func_event = tuple(range(start_col, end_col + 1))
                    elif end_col < start_col:
                        rect = (0, end_col, len(self.MT.row_positions) - 1, start_col + 1, "columns")
                        func_event = tuple(range(end_col, start_col + 1))
                    if self.being_drawn_rect != rect:
                        need_redraw = True
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
                    need_redraw = True
                elif event.x < 0 and self.canvasx(self.winfo_width()) > 0 and xcheck and xcheck[0] > 0:
                    try:
                        self.xview_scroll(-1, "units")
                        self.MT.xview_scroll(-1, "units")
                    except:
                        pass
                    self.check_xview()
                    need_redraw = True
            if need_redraw:
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
        if self.being_drawn_rect is not None:
            to_sel = tuple(self.being_drawn_rect)
            self.being_drawn_rect = None
            self.MT.create_selected(*to_sel)
        self.MT.bind("<MouseWheel>", self.MT.mousewheel)
        if self.width_resizing_enabled and self.rsz_w is not None and self.currently_resizing_width:
            self.currently_resizing_width = False
            new_col_pos = int(self.coords("rwl")[0])
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
            orig_selected = self.MT.get_selected_cols()
            if c != self.dragged_col and c is not None and c not in orig_selected and len(orig_selected) != len(self.MT.col_positions) - 1:
                orig_selected = sorted(orig_selected)
                if len(orig_selected) > 1:
                    start_idx = bisect.bisect_left(orig_selected, self.dragged_col)
                    forward_gap = get_index_of_gap_in_sorted_integer_seq_forward(orig_selected, start_idx)
                    reverse_gap = get_index_of_gap_in_sorted_integer_seq_reverse(orig_selected, start_idx)
                    if forward_gap is not None:
                        orig_selected[:] = orig_selected[:forward_gap]
                    if reverse_gap is not None:
                        orig_selected[:] = orig_selected[reverse_gap:]
                rm1start = orig_selected[0]
                totalcols = len(orig_selected)
                extra_func_success = True
                if c >= len(self.MT.col_positions) - 1:
                    c -= 1
                if self.ch_extra_begin_drag_drop_func is not None:
                    try:
                        self.ch_extra_begin_drag_drop_func(BeginDragDropEvent("begin_column_header_drag_drop", tuple(orig_selected), int(c)))
                    except:
                        extra_func_success = False
                if extra_func_success:
                    new_selected, dispset = self.MT.move_columns_adjust_options_dict(c,
                                                                                     rm1start, 
                                                                                     totalcols,
                                                                                     move_data = self.column_drag_and_drop_perform)
                    if self.MT.undo_enabled:
                        self.MT.undo_storage.append(zlib.compress(pickle.dumps(("move_cols",                 #0
                                                                                int(orig_selected[0]),  #1
                                                                                int(new_selected[0]),        #2
                                                                                int(new_selected[-1]),       #3
                                                                                sorted(orig_selected),  #4
                                                                                dispset))))                  #5
                    self.MT.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
                    if self.ch_extra_end_drag_drop_func is not None:
                        self.ch_extra_end_drag_drop_func(EndDragDropEvent("end_column_header_drag_drop", tuple(orig_selected), new_selected, int(c)))
        elif self.b1_pressed_loc is not None and self.rsz_w is None and self.rsz_h is None:
            c = self.MT.identify_col(x = event.x)
            if c is not None and c == self.b1_pressed_loc and self.b1_pressed_loc != self.closed_dropdown:
                dcol = c if self.MT.all_columns_displayed else self.MT.displayed_columns[c]
                canvasx = self.canvasx(event.x)
                if ((dcol in self.cell_options and 'dropdown' in self.cell_options[dcol] and canvasx < self.MT.col_positions[c + 1] and canvasx > self.MT.col_positions[c + 1] - self.MT.hdr_txt_h - 4) or
                    (dcol in self.cell_options and 'checkbox' in self.cell_options[dcol] and canvasx < self.MT.col_positions[c] + self.MT.hdr_txt_h + 5)):
                    if event.y < self.MT.hdr_txt_h + 5:
                        self.open_cell(event)
            else:
                self.mouseclick_outside_editor_or_dropdown()
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
                self.MT.create_current(0, c, type_ = "column", inside = True)
            else:
                ignore_keep = True
        if ignore_keep or not keep_other_selections:
            self.MT.delete_selection_rects()
            self.MT.create_current(0, c, type_ = "column", inside = True)
            self.MT.create_selected(0, c, len(self.MT.row_positions) - 1, c + 1, "columns")
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
            self.MT.create_current(0, c, type_ = "column", inside = True)
            if create_new_sel:
                r1, c1, r2, c2 = tuple(int(e) for e in current[1].split("_") if e)
                self.MT.create_selected(r1, c1, r2, c2, current[2] + "s")
        self.MT.create_selected(0, c, len(self.MT.row_positions) - 1, c + 1, "columns")
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
        qfont = self.MT._font
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
                qconf(qtxtm, text = self.cell_options[data_col]['checkbox']['text'], font = self.MT._hdr_font)
                b = qbbox(qtxtm)
                hw = b[2] - b[0] + 7 + self.MT.hdr_txt_h
            else:
                txt = ""
                if isinstance(self.MT._headers, int) or len(self.MT._headers) - 1 >= data_col:
                    if isinstance(self.MT._headers, int):
                        txt = self.MT.data[self.MT._headers][data_col]
                    else:
                        txt = self.MT._headers[data_col]
                    if txt or (data_col in self.cell_options and 'dropdown' in self.cell_options[data_col]):
                        qconf(qtxtm, text = txt, font = self.MT._hdr_font)
                        b = qbbox(qtxtm)
                        if data_col in self.cell_options and 'dropdown' in self.cell_options[data_col]:
                            hw = b[2] - b[0] + 7 + self.MT.hdr_txt_h
                        else:
                            hw = b[2] - b[0] + 7
                if not isinstance(self.MT._headers, int) and ((not txt and self.show_default_header_for_empty) or len(self.MT._headers) < data_col):
                    if self.default_hdr == "letters":
                        hw = self.MT.GetHdrTextWidth(num2alpha(data_col)) + 7
                    elif self.default_hdr == "numbers":
                        hw = self.MT.GetHdrTextWidth(f"{data_col + 1}") + 7
                    else:
                        hw = self.MT.GetHdrTextWidth(f"{data_col + 1} {num2alpha(data_col)}") + 7
            # table
            for rn, r in enumerate(islice(self.MT.data, start_row, end_row), start_row):
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
        if not return_new_width:
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
                                                outline = self.header_fg if hlcol in self.cell_options and 'dropdown' in self.cell_options[hlcol] and self.MT.show_dropdown_borders else "",
                                                tag = "hi")
            fill = self.header_selected_columns_fg if self.cell_options[hlcol]['highlight'][1] is None or self.MT.display_selected_fg_over_highlights else self.cell_options[hlcol]['highlight'][1]
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
                                                outline = self.header_fg if hlcol in self.cell_options and 'dropdown' in self.cell_options[hlcol] and self.MT.show_dropdown_borders else "", 
                                                tag = "hi")
            fill = self.header_selected_cells_fg if self.cell_options[hlcol]['highlight'][1] is None or self.MT.display_selected_fg_over_highlights else self.cell_options[hlcol]['highlight'][1]
        elif c in actual_selected_cols:
            fill = self.header_selected_columns_fg
        elif c in selected_cols or selected_rows:
            fill = self.header_selected_cells_fg
        elif hlcol in self.cell_options and 'highlight' in self.cell_options[hlcol]:
            if self.cell_options[hlcol]['highlight'][0] is not None:
                redrawn = self.redraw_highlight(fc + 1,
                                                0, 
                                                sc, 
                                                self.current_height - 1, 
                                                fill = self.cell_options[hlcol]['highlight'][0], 
                                                outline = self.header_fg if hlcol in self.cell_options and 'dropdown' in self.cell_options[hlcol] and self.MT.show_dropdown_borders else "", 
                                                tag = "hi")
            fill = self.header_fg if self.cell_options[hlcol]['highlight'][1] is None else self.cell_options[hlcol]['highlight'][1]
        else:
            fill = self.header_fg
        return fill, redrawn
            
    def redraw_highlight(self, x1, y1, x2, y2, fill, outline, tag):
        config = (fill, outline)
        coords = (x1 - 1 if outline else x1,
                  y1 - 1 if outline else y1,
                  x2, 
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

    def redraw_gridline(self,points, fill, width, tag):
        if self.hidd_grid:
            t, sh = self.hidd_grid.popitem()
            self.coords(t, points)
            if sh:
                self.itemconfig(t, fill = fill, width = width, tag = tag, capstyle = tk.BUTT, joinstyle = tk.ROUND)
            else:
                self.itemconfig(t, fill = fill, width = width, tag = tag, capstyle = tk.BUTT, joinstyle = tk.ROUND, state = "normal")
            self.disp_grid[t] = True
        else:
            self.disp_grid[self.create_line(points, fill = fill, width = width, tag = tag)] = True
            
    def redraw_dropdown(self, x1, y1, x2, y2, fill, outline, tag, draw_outline = True, draw_arrow = True, dd_is_open = False):
        if draw_outline and self.MT.show_dropdown_borders:
            self.redraw_highlight(x1 + 1, y1 + 1, x2, y2, fill = "", outline = self.header_fg, tag = tag)
        if draw_arrow:
            topysub = floor(self.MT.hdr_half_txt_h / 2)
            mid_y = y1 + floor(self.MT.hdr_min_rh / 2)
            if mid_y + topysub + 1 >= y1 + self.MT.hdr_txt_h - 1:
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
            tx1 = x2 - self.MT.hdr_txt_h + 1
            tx2 = x2 - self.MT.hdr_half_txt_h - 1
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
            x1 = x1 + 4
            y1 = y1 + 4
            x2 = x2 - 3
            y2 = y2 - 3
            points = self.MT.get_checkbox_points(x1, y1, x2, y2, radius = 4)
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

    def redraw_grid_and_text(self, last_col_line_pos, scrollpos_left, x_stop, start_col, end_col, selected_cols, selected_rows, actual_selected_cols):
        self.configure(scrollregion = (0,
                                       0,
                                       last_col_line_pos + self.MT.empty_horizontal,
                                       self.current_height))
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
        self.col_height_resize_bbox = (scrollpos_left, self.current_height - 2, x_stop, self.current_height)
        draw_x = self.MT.col_positions[start_col]
        yend = self.current_height - 5
        if self.MT.show_vertical_grid or self.width_resizing_enabled:
            self.grid_cyc = cycle(self.grid_cyctup)
            points = []
            for c in range(start_col + 1, end_col):
                draw_x = self.MT.col_positions[c]
                if self.width_resizing_enabled:
                    self.visible_col_dividers[c] = (draw_x - 2, 1, draw_x + 2, yend)
                st_or_end = next(self.grid_cyc)
                if st_or_end == "st":
                    points.extend([draw_x, -1,
                                   draw_x, self.current_height,
                                   self.MT.col_positions[c + 1] if len(self.MT.col_positions) - 1 > c else draw_x, self.current_height])
                elif st_or_end == "end":
                    points.extend([draw_x, self.current_height,
                                   draw_x, -1,
                                   self.MT.col_positions[c + 1] if len(self.MT.col_positions) - 1 > c else draw_x, -1])
                if points:
                    self.redraw_gridline(points = points, fill = self.header_grid_fg, width = 1, tag = "v")
        self.redraw_gridline(points = (draw_x, 0, draw_x, self.current_height), fill = self.header_grid_fg, width = 1, tag = "fv")
        self.redraw_gridline(points = (scrollpos_left, self.current_height - 1, x_stop, self.current_height - 1), fill = self.header_border_fg, width = 1, tag = "h")
        top = self.canvasy(0)
        c_2 = self.header_selected_cells_bg if self.header_selected_cells_bg.startswith("#") else Color_Map_[self.header_selected_cells_bg]
        c_3 = self.header_selected_columns_bg if self.header_selected_columns_bg.startswith("#") else Color_Map_[self.header_selected_columns_bg]
        font = self.MT._hdr_font
        for c in range(start_col, end_col - 1):
            draw_y = self.MT.hdr_fl_ins
            cleftgridln = self.MT.col_positions[c]
            crightgridln = self.MT.col_positions[c + 1]
            if self.MT.all_columns_displayed:
                dcol = c
            else:
                dcol = self.MT.displayed_columns[c]
            fill, dd_drawn = self.redraw_highlight_get_text_fg(cleftgridln, crightgridln, c, c_2, c_3, selected_cols, selected_rows, actual_selected_cols, dcol)

            if dcol in self.cell_options and 'align' in self.cell_options[dcol]:
                align = self.cell_options[dcol]['align']
            else:
                align = self.align
                
            if align == "w":
                draw_x = cleftgridln + 3
                if dcol in self.cell_options and 'dropdown' in self.cell_options[dcol]:
                    mw = crightgridln - cleftgridln - self.MT.hdr_txt_h - 2
                    self.redraw_dropdown(cleftgridln, 0, crightgridln, self.current_height - 1, 
                                         fill = fill, outline = fill, tag = "dd", draw_outline = not dd_drawn, draw_arrow = mw >= 5,
                                         dd_is_open = self.cell_options[dcol]['dropdown']['window'] != "no dropdown open")
                else:
                    mw = crightgridln - cleftgridln - 1

            elif align == "e":
                if dcol in self.cell_options and 'dropdown' in self.cell_options[dcol]:
                    mw = crightgridln - cleftgridln - self.MT.hdr_txt_h - 2
                    draw_x = crightgridln - 5 - self.MT.hdr_txt_h
                    self.redraw_dropdown(cleftgridln, 0, crightgridln, self.current_height - 1,
                                         fill = fill, outline = fill, tag = "dd", draw_outline = not dd_drawn, draw_arrow = mw >= 5,
                                         dd_is_open = self.cell_options[dcol]['dropdown']['window'] != "no dropdown open")
                else:
                    mw = crightgridln - cleftgridln - 1
                    draw_x = crightgridln - 3

            elif align == "center":
                #stop = cleftgridln + 5
                if dcol in self.cell_options and 'dropdown' in self.cell_options[dcol]:
                    mw = crightgridln - cleftgridln - self.MT.hdr_txt_h - 2
                    draw_x = cleftgridln + ceil((crightgridln - cleftgridln - self.MT.hdr_txt_h) / 2)
                    self.redraw_dropdown(cleftgridln, 0, crightgridln, self.current_height - 1, 
                                         fill = fill, outline = fill, tag = "dd", draw_outline = not dd_drawn, draw_arrow = mw >= 5,
                                         dd_is_open = self.cell_options[dcol]['dropdown']['window'] != "no dropdown open")
                else:
                    mw = crightgridln - cleftgridln - 1
                    draw_x = cleftgridln + floor((crightgridln - cleftgridln) / 2)

            if dcol in self.cell_options and 'checkbox' in self.cell_options[dcol]:
                if mw > self.MT.hdr_txt_h + 2:
                    box_w = self.MT.hdr_txt_h + 1
                    mw -= box_w
                    if align == "w":
                        draw_x += box_w + 1
                    elif align == "center":
                        draw_x += ceil(box_w / 2) + 1
                        mw -= 1
                    else:
                        mw -= 3
                    try:
                        draw_check = self.MT._headers[dcol] if isinstance(self.MT._headers, (list, tuple)) else self.MT.data[self.MT._headers][dcol]
                    except:
                        draw_check = False
                    self.redraw_checkbox(dcol,
                                         cleftgridln + 2,
                                         2,
                                         cleftgridln + self.MT.hdr_txt_h + 3,
                                         self.MT.hdr_txt_h + 3,
                                         fill = fill if self.cell_options[dcol]['checkbox']['state'] == "normal" else self.header_grid_fg,
                                         outline = "",
                                         tag = "cb", 
                                         draw_check = draw_check)

            try:
                if dcol in self.cell_options and 'checkbox' in self.cell_options[dcol]:
                    lns = self.cell_options[dcol]['checkbox']['text'].split("\n") if isinstance(self.cell_options[dcol]['checkbox']['text'], str) else f"{self.cell_options[dcol]['checkbox']['text']}".split("\n")
                elif isinstance(self.MT._headers, int):
                    lns = self.MT.data[self.MT._headers][dcol].split("\n") if isinstance(self.MT.data[self.MT._headers][dcol], str) else f"{self.MT.data[self.MT._headers][dcol]}".split("\n")
                else:
                    lns = self.MT._headers[dcol].split("\n") if isinstance(self.MT._headers[dcol], str) else f"{self.MT._headers[dcol]}".split("\n")
                got_hdr = True
            except:
                got_hdr = False
            if not got_hdr or (lns == [""] and self.show_default_header_for_empty):
                if self.default_hdr == "letters":
                    lns = (num2alpha(dcol), )
                elif self.default_hdr == "numbers":
                    lns = (f"{dcol + 1}", )
                else:
                    lns = (f"{dcol + 1} {num2alpha(dcol)}", )
            if mw > self.MT.hdr_txt_w and not ((align == "w" and (draw_x > x_stop)) or
                                               (align == "e" and (draw_x > x_stop)) or
                                               (align == "center" and (cleftgridln + 5 > x_stop))):
                for txt in islice(lns, self.lines_start_at if self.lines_start_at < len(lns) else len(lns) - 1, None):
                    if draw_y > top:
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
                    draw_y += self.MT.hdr_xtra_lines_increment
                    if draw_y - 1 > self.current_height:
                        break
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
                
    def open_cell(self, event = None, ignore_existing_editor = False):
        if not self.MT.anything_selected()  or (not ignore_existing_editor and self.text_editor_id is not None):
            return
        currently_selected = self.MT.currently_selected()
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
        if event is None or self.MT.event_opens_dropdown_or_checkbox(event):
            if event is not None:
                if hasattr(event, 'keysym') and event.keysym == 'Return':
                    extra_func_key = "Return"
                elif hasattr(event, 'keysym') and event.keysym == 'F2':
                    extra_func_key = "F2"
            dcol = c if self.MT.all_columns_displayed else self.MT.displayed_columns[c]
            if isinstance(self.MT._headers, list):
                if len(self.MT._headers) <= dcol:
                    self.MT._headers.extend(list(repeat("", dcol - len(self.MT._headers) + 1)))
                text = f"{self.MT._headers[dcol]}"
            elif isinstance(self.MT._headers, int):
                try:
                    text = f"{self.MT.data[self.MT._headers][dcol]}"
                except:
                    text = ""
            
        elif event is not None and ((hasattr(event, 'keysym') and event.keysym == 'BackSpace') or
                                  event.keycode in (8, 855638143)
                                  ):
            extra_func_key = "BackSpace"
            text = ""
        elif event is not None and ((hasattr(event, "char") and event.char.isalpha()) or
                                    (hasattr(event, "char") and event.char.isdigit()) or
                                    (hasattr(event, "char") and event.char in symbols_set)):
            extra_func_key = event.char
            text = event.char
        else:
            return False
        self.text_editor_loc = c
        if self.extra_begin_edit_cell_func is not None:
            try:
                text = self.extra_begin_edit_cell_func(EditHeaderEvent(c, extra_func_key, text, "begin_edit_header"))
            except:
                return False
            if text is None:
                return False
            else:
                text = text if isinstance(text, str) else f"{text}"
        text = "" if text is None else text
        if self.MT.cell_auto_resize_enabled:
            if self.height_resizing_enabled:
                self.set_current_height_to_cell(dcol)
            self.set_col_width_run_binding(c)
        self.select_col(c = c, keep_other_selections = True)
        self.create_text_editor(c = c, text = text, set_data_ref_on_destroy = True, dropdown = dropdown)
        return True
    
    # displayed indexes
    def get_cell_align(self, c):
        dcol = c if self.MT.all_columns_displayed else self.MT.displayed_columns[c]
        if dcol in self.cell_options and 'align' in self.cell_options[dcol]:
            align = self.cell_options[dcol]['align']
        else:
            align = self.align
        return align

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
            if isinstance(self.MT._headers, list):
                if len(self.MT._headers) <= dcol:
                    self.MT._headers.extend(list(repeat("", dcol - len(self.MT._headers) + 1)))
                text = f"{self.MT._headers[dcol]}"
            elif isinstance(self.MT._headers, int):
                try:
                    text = f"{self.MT.data[self.MT._headers][dcol]}"
                except:
                    text = ""
        #bg, fg = self.get_widget_bg_fg(dcol)
        bg, fg = self.header_bg, self.header_fg
        self.text_editor = TextEditor(self,
                                      text = text,
                                      font = self.MT._hdr_font,
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
                                      popup_menu_highlight_fg = self.MT.popup_menu_highlight_fg,
                                      binding = binding,
                                      align = self.get_cell_align(c),
                                      c = c,
                                      newline_binding = self.text_editor_has_wrapped)
        self.text_editor.update_idletasks()
        self.text_editor_id = self.create_window((x, y), window = self.text_editor, anchor = "nw")
        if not dropdown:
            self.text_editor.textedit.focus_set()
            self.text_editor.scroll_to_bottom()
        self.text_editor.textedit.bind("<Alt-Return>", lambda x: self.text_editor_newline_binding(c = c))
        if USER_OS == 'darwin':
            self.text_editor.textedit.bind("<Option-Return>", lambda x: self.text_editor_newline_binding(c = c))
        for key, func in self.MT.text_editor_user_bound_keys.items():
            self.text_editor.textedit.bind(key, func)
        if binding is not None:
            self.text_editor.textedit.bind("<Tab>", lambda x: binding((c, "Tab")))
            self.text_editor.textedit.bind("<Return>", lambda x: binding((c, "Return")))
            self.text_editor.textedit.bind("<FocusOut>", lambda x: binding((c, "FocusOut")))
            self.text_editor.textedit.bind("<Escape>", lambda x: binding((c, "Escape")))
        elif binding is None and set_data_ref_on_destroy:
            self.text_editor.textedit.bind("<Tab>", lambda x: self.close_text_editor((c, "Tab")))
            self.text_editor.textedit.bind("<Return>", lambda x: self.close_text_editor((c, "Return")))
            if not dropdown:
                self.text_editor.textedit.bind("<FocusOut>", lambda x: self.close_text_editor((c, "FocusOut")))
            self.text_editor.textedit.bind("<Escape>", lambda x: self.close_text_editor((c, "Escape")))
        else:
            self.text_editor.textedit.bind("<Escape>", lambda x: self.destroy_text_editor("Escape"))
    
    # displayed indexes                            #just here to receive text editor arg
    def text_editor_has_wrapped(self, r = 0, c = 0, check_lines = None):
        if self.width_resizing_enabled:
            dcol = c if self.MT.all_columns_displayed else self.MT.displayed_columns[c]
            curr_width = self.text_editor.winfo_width()
            new_width = curr_width + (self.MT.hdr_txt_h * 2)
            if new_width != curr_width:
                self.text_editor.config(width = new_width)
                self.set_col_width_run_binding(c, width = new_width, only_set_if_too_small = False)
                if dcol in self.cell_options and 'dropdown' in self.cell_options[dcol]:
                    self.itemconfig(self.cell_options[dcol]['dropdown']['canvas_id'],
                                    width = new_width)
                    self.cell_options[dcol]['dropdown']['window'].update_idletasks()
                    self.cell_options[dcol]['dropdown']['window']._reselect()
                self.MT.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = False, redraw_table = True)
    
    # displayed indexes
    def text_editor_newline_binding(self, r = 0, c = 0, event = None, check_lines = True):
        if self.height_resizing_enabled:
            dcol = c if self.MT.all_columns_displayed else self.MT.displayed_columns[c]
            curr_height = self.text_editor.winfo_height()
            if not check_lines or self.MT.GetHdrLinesHeight(self.text_editor.get_num_lines() + 1) > curr_height:
                new_height = curr_height + self.MT.hdr_xtra_lines_increment
                space_bot = self.MT.get_space_bot(0)
                if new_height > space_bot:
                    new_height = space_bot
                if new_height != curr_height:
                    self.text_editor.config(height = new_height)
                    self.set_height(new_height, set_TL = True)
                    if dcol in self.cell_options and 'dropdown' in self.cell_options[dcol]:
                        win_h, anchor = self.get_dropdown_height_anchor(dcol, new_height)
                        self.coords(self.cell_options[dcol]['dropdown']['canvas_id'],
                                    self.MT.col_positions[c], new_height - 1)
                        self.itemconfig(self.cell_options[dcol]['dropdown']['canvas_id'],
                                        anchor = anchor, height = win_h)
            
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
    def close_text_editor(self, editor_info = None, c = None, set_data_ref_on_destroy = True, event = None, destroy = True, move_down = True, redraw = True, recreate = True):
        if self.focus_get() is None and editor_info:
            return
        if editor_info is not None and len(editor_info) >= 2 and editor_info[1] == "Escape":
            self.destroy_text_editor("Escape")
            self.hide_dropdown_window(c)
            return
        if self.text_editor is not None:
            self.text_editor_value = self.text_editor.get()
        if destroy:
            self.destroy_text_editor()
        if set_data_ref_on_destroy:
            if c is None and editor_info is not None and len(editor_info) >= 2:
                c = editor_info[0]
            if self.extra_end_edit_cell_func is None:
                self._set_cell_data(c, dcol = c if self.MT.all_columns_displayed else self.MT.displayed_columns[c], value = self.text_editor_value)
            elif self.extra_end_edit_cell_func is not None and not self.MT.edit_cell_validation:
                self._set_cell_data(c, dcol = c if self.MT.all_columns_displayed else self.MT.displayed_columns[c], value = self.text_editor_value)
                self.extra_end_edit_cell_func(EditHeaderEvent(c, editor_info[1] if len(editor_info) >= 2 else "FocusOut", f"{self.text_editor_value}", "end_edit_header"))
            elif self.extra_end_edit_cell_func is not None and self.MT.edit_cell_validation:
                validation = self.extra_end_edit_cell_func(EditHeaderEvent(c, editor_info[1] if len(editor_info) >= 2 else "FocusOut", f"{self.text_editor_value}", "end_edit_header"))
                if validation is not None:
                    self.text_editor_value = validation
                    self._set_cell_data(c, dcol = c if self.MT.all_columns_displayed else self.MT.displayed_columns[c], value = self.text_editor_value)
        if move_down:
            pass
        self.hide_dropdown_window(c)
        if recreate:
            self.MT.recreate_all_selection_boxes()
        if redraw:
            self.MT.refresh()
        if editor_info is not None and len(editor_info) >= 2 and editor_info[1] != "FocusOut":
            self.focus_set()
        return "break"
    
    #internal event use
    def _set_cell_data(self, c = 0, dcol = None, value = "", cell_resize = True, undo = True, redraw = True):
        if dcol is None:
            dcol = c if self.MT.all_columns_displayed else self.MT.displayed_columns[c]
        if isinstance(self.MT._headers, list):
            if len(self.MT._headers) <= dcol:
                self.MT._headers.extend(list(repeat("", dcol - len(self.MT._headers) + 1)))
            if self.MT.undo_enabled and undo:
                if self.MT._headers[dcol] != value:
                    self.MT.undo_storage.append(zlib.compress(pickle.dumps(("edit_header",
                                                                           {dcol: self.MT._headers[dcol]},
                                                                           (((0, c, len(self.MT.row_positions) - 1, c + 1), "columns"), ),
                                                                           self.MT.currently_selected()))))
            self.MT._headers[dcol] = value
        elif isinstance(self.MT._headers, int):
            self.MT._set_cell_data(r = self.MT._headers, c = c, dcol = dcol, value = value, undo = True)
        if cell_resize and self.MT.cell_auto_resize_enabled:
            if self.height_resizing_enabled:
                self.set_current_height_to_cell(dcol)
            self.set_col_width_run_binding(c)
        if redraw:
            self.MT.refresh()
        return True
    
    # displayed indexes
    def set_col_width_run_binding(self, c, width = None, only_set_if_too_small = True):
        old_width = self.MT.col_positions[c + 1] - self.MT.col_positions[c]
        new_width = self.set_col_width(c, width = width, only_set_if_too_small = only_set_if_too_small)
        if self.column_width_resize_func is not None and old_width != new_width:
            self.column_width_resize_func(ResizeEvent("column_width_resize", c, old_width, new_width))

    #internal event use
    def _click_checkbox(self, c, dcol = None, undo = True, redraw = True):
        if dcol is None:
            dcol = c if self.MT.all_columns_displayed else self.MT.displayed_columns[c]
        if self.cell_options[dcol]['checkbox']['state'] == "normal":
            if isinstance(self.MT._headers, list):
                self._set_cell_data(c, dcol = dcol, value = not self.MT._headers[dcol] if type(self.MT._headers[dcol]) == bool else False, cell_resize = False)
            elif isinstance(self.MT._headers, int):
                self._set_cell_data(c, dcol = dcol, value = not self.MT.data[self.MT._headers][dcol] if type(self.MT.data[self.MT._headers][dcol]) == bool else False, cell_resize = False)
            if self.cell_options[dcol]['checkbox']['check_function'] is not None:
                self.cell_options[dcol]['checkbox']['check_function']((0, c, "HeaderCheckboxClicked", f"{self.MT._headers[dcol] if isinstance(self.MT._headers, list) else self.MT.data[self.MT._headers][dcol]}"))
            if self.extra_end_edit_cell_func is not None:
                self.extra_end_edit_cell_func(EditHeaderEvent(c, "Return", f"{self.MT._headers[dcol] if isinstance(self.MT._headers, list) else self.MT.data[self.MT._headers][dcol]}", "end_edit_header"))
        if redraw:
            self.MT.refresh()

    def create_checkbox(self, c = 0, checked = False, state = "normal", redraw = False, check_function = None, text = ""):
        if c in self.cell_options and any(x in self.cell_options[c] for x in ('dropdown', 'checkbox')):
            self.delete_dropdown_and_checkbox(c)
        self._set_cell_data(dcol = c, value = checked, cell_resize = False, undo = False) # only works because cell_resize and undo are false else needs c arg
        if c not in self.cell_options:
            self.cell_options[c] = {}
        self.cell_options[c]['checkbox'] = {'check_function': check_function,
                                            'state': state,
                                            'text': text}
        if redraw:
            self.MT.refresh()

    def create_dropdown(self, c = 0, values = [], set_value = None, state = "readonly", redraw = True, selection_function = None, modified_function = None):
        if c in self.cell_options and any(x in self.cell_options[c] for x in ('dropdown', 'checkbox')):
            self.delete_dropdown_and_checkbox(c)
        self._set_cell_data(dcol = c, value = set_value if set_value is not None else values[0] if values else "", cell_resize = False, undo = False) # only works because cell_resize and undo are false else needs c arg
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
        win_h = 5
        for i, v in enumerate(self.cell_options[dcol]['dropdown']['values']):
            v_numlines = len(v.split("\n") if isinstance(v, str) else f"{v}".split("\n"))
            if v_numlines > 1:
                win_h += self.MT.hdr_fl_ins + (v_numlines * self.MT.hdr_xtra_lines_increment) + 5 # end of cell
            else:
                win_h += self.MT.hdr_min_rh
            if i == 5:
                break
        if win_h > 500:
            win_h = 500
        space_bot = self.MT.get_space_bot(0, text_editor_h)
        win_h2 = int(win_h)
        if win_h > space_bot:
            win_h = space_bot - 1
        if win_h < self.MT.hdr_txt_h + 5:
            win_h = self.MT.hdr_txt_h + 5
        elif win_h > win_h2:
            win_h = win_h2
        return win_h, "nw"
    
    # data indexes
    def set_current_height_to_cell(self, dcol):
        x = self.MT.txt_measure_canvas.create_text(0,
                                                   0,
                                                   text = self.MT.data[self.MT._headers][dcol] if isinstance(self.MT._headers, int) else self.MT._headers[dcol],
                                                   font = self.MT._hdr_font)
        b = self.MT.txt_measure_canvas.bbox(x)
        self.MT.txt_measure_canvas.delete(x)
        new_height = b[3] - b[1] + 5
        space_bot = self.MT.get_space_bot(0)
        if new_height > space_bot:
            new_height = space_bot
        self.set_height(new_height, set_TL = True)
            
    # c is displayed col
    def display_dropdown_window(self, c, dcol = None, event = None):
        self.destroy_text_editor("Escape")
        self.destroy_opened_dropdown_window()
        if dcol is None:
            dcol = c if self.MT.all_columns_displayed else self.MT.displayed_columns[c]
        if self.cell_options[dcol]['dropdown']['state'] == "normal":
            if not self.edit_cell_(c = c, dropdown = True, event = event):
                return
        win_h, anchor = self.get_dropdown_height_anchor(dcol)
        window = self.MT.parentframe.dropdown_class(self.MT.winfo_toplevel(),
                                                    0,
                                                    c,
                                                    width = self.MT.col_positions[c + 1] - self.MT.col_positions[c] + 1,
                                                    height = win_h,
                                                    font = self.MT._hdr_font,
                                                    colors = {'bg': self.MT.popup_menu_bg, 
                                                              'fg': self.MT.popup_menu_fg, 
                                                              'highlight_bg': self.MT.popup_menu_highlight_bg,
                                                              'highlight_fg': self.MT.popup_menu_highlight_fg},
                                                    outline_color = self.MT.popup_menu_fg,
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
            self.update_idletasks()
            try:
                self.after(1, lambda: self.text_editor.textedit.focus())
                self.after(2, self.text_editor.scroll_to_bottom())
            except:
                return
            redraw = False
        else:
            window.bind("<FocusOut>", lambda x: self.hide_dropdown_window(c))
            self.update_idletasks()
            window.focus_set()
            redraw = True
        self.existing_dropdown_window = window
        self.cell_options[dcol]['dropdown']['window'] = window
        self.existing_dropdown_canvas_id = self.cell_options[dcol]['dropdown']['canvas_id']
        if redraw:
            self.MT.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = False, redraw_table = False)
            
    # c is displayed col
    def hide_dropdown_window(self, c = None, selection = None, redraw = True):
        if c is not None and selection is not None:
            dcol = c if self.MT.all_columns_displayed else self.MT.displayed_columns[c]
            if self.cell_options[dcol]['dropdown']['select_function'] is not None: # user has specified a selection function
                self.cell_options[dcol]['dropdown']['select_function'](EditHeaderEvent(c, "HeaderComboboxSelected", f"{selection}", "end_edit_header"))
            if self.extra_end_edit_cell_func is None:
                self._set_cell_data(c, dcol = dcol, value = selection, redraw = not redraw)
            elif self.extra_end_edit_cell_func is not None and self.MT.edit_cell_validation:
                validation = self.extra_end_edit_cell_func(EditHeaderEvent(c, "HeaderComboboxSelected", f"{selection}", "end_edit_header"))
                if validation is not None:
                    selection = validation
                self._set_cell_data(c, dcol = dcol, value = selection, redraw = not redraw)
            elif self.extra_end_edit_cell_func is not None and not self.MT.edit_cell_validation:
                self._set_cell_data(c, dcol = dcol, value = selection, redraw = not redraw)
                self.extra_end_edit_cell_func(EditHeaderEvent(c, "HeaderComboboxSelected", f"{selection}", "end_edit_header"))
            self.focus_set()
            self.MT.recreate_all_selection_boxes()
        self.destroy_text_editor("Escape")
        self.destroy_opened_dropdown_window(c)
        if redraw:
            self.MT.refresh()
        
    def mouseclick_outside_editor_or_dropdown(self):
        if self.existing_dropdown_window is not None:
            closed_dd_coords = int(self.existing_dropdown_window.c)
        else:
            closed_dd_coords = None
        if self.text_editor_loc is not None and self.text_editor is not None:
            self.close_text_editor(editor_info = (self.text_editor_loc, "ButtonPress-1"))
        else:
            self.destroy_text_editor("Escape")
        if closed_dd_coords:
            self.destroy_opened_dropdown_window(closed_dd_coords) #displayed coords not data, necessary for b1 function
        return closed_dd_coords
            
    # function can receive two None args
    def destroy_opened_dropdown_window(self, c = None, dcol = None):
        if c is not None or dcol is not None:
            if dcol is None:
                dcol_ = c if self.MT.all_columns_displayed else self.MT.displayed_columns[c]
            else:
                dcol_ = dcol
        else:
            dcol_ = None
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
        if dcol_ in self.cell_options and 'dropdown' in self.cell_options[dcol_]:
            self.cell_options[dcol_]['dropdown']['canvas_id'] = "no dropdown open"
            self.cell_options[dcol_]['dropdown']['window'] = "no dropdown open"
            try:
                self.delete(self.cell_options[dcol_]['dropdown']['canvas_id'])
            except:
                pass
            
    # c is dcol
    def delete_dropdown(self, c):
        self.destroy_opened_dropdown_window(dcol = c)
        if c in self.cell_options and 'dropdown' in self.cell_options[c]:
            del self.cell_options[c]['dropdown']
            
    # c is dcol
    def delete_checkbox(self, c):
        if c in self.cell_options and 'checkbox' in self.cell_options[c]:
            del self.cell_options[c]['checkbox']

    # c is dcol
    def delete_dropdown_and_checkbox(self, c):
        self.delete_dropdown(dcol = c)
        self.delete_checkbox(c)
