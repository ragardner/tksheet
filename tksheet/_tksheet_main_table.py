from __future__ import annotations

import bisect
import csv as csv
import io
import pickle
import tkinter as tk
import zlib
from collections import defaultdict, deque
from itertools import accumulate, chain, cycle, islice, product, repeat
from math import ceil, floor
from tkinter import TclError
from typing import Any, Union
from ._tksheet_formatters import (
    data_to_str,
    format_data,
    get_clipboard_data,
    get_data_with_valid_check,
    is_bool_like,
    try_to_bool,
)
from ._tksheet_other_classes import (
    CtrlKeyEvent,
    CurrentlySelectedClass,
    DeleteRowColumnEvent,
    DeselectionEvent,
    DrawnItem,
    DropDownModifiedEvent,
    EditCellEvent,
    InsertEvent,
    PasteEvent,
    ResizeEvent,
    SelectCellEvent,
    SelectionBoxEvent,
    TextCfg,
    TextEditor,
    UndoEvent,
    get_checkbox_dict,
    get_dropdown_dict,
    get_seq_without_gaps_at_index,
    is_iterable,
)
from ._tksheet_vars import (
    USER_OS,
    Color_Map_,
    arrowkey_bindings_helper,
    ctrl_key,
    rc_binding,
    symbols_set,
)


class MainTable(tk.Canvas):
    def __init__(self, *args, **kwargs):
        tk.Canvas.__init__(
            self,
            kwargs["parentframe"],
            background=kwargs["table_bg"],
            highlightthickness=0,
        )
        self.parentframe = kwargs["parentframe"]
        self.parentframe_width = 0
        self.parentframe_height = 0
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
        self.options = {}

        self.arrowkey_binding_functions = {
            "tab": self.tab_key,
            "up": self.arrowkey_UP,
            "right": self.arrowkey_RIGHT,
            "down": self.arrowkey_DOWN,
            "left": self.arrowkey_LEFT,
            "prior": self.page_UP,
            "next": self.page_DOWN,
        }
        self.extra_table_rc_menu_funcs = {}
        self.extra_index_rc_menu_funcs = {}
        self.extra_header_rc_menu_funcs = {}
        self.extra_empty_space_rc_menu_funcs = {}

        self.max_undos = kwargs["max_undos"]
        self.undo_storage = deque(maxlen=kwargs["max_undos"])

        self.to_clipboard_delimiter = kwargs["to_clipboard_delimiter"]
        self.to_clipboard_quotechar = kwargs["to_clipboard_quotechar"]
        self.to_clipboard_lineterminator = kwargs["to_clipboard_lineterminator"]
        self.from_clipboard_delimiters = (
            kwargs["from_clipboard_delimiters"]
            if isinstance(kwargs["from_clipboard_delimiters"], str)
            else "".join(kwargs["from_clipboard_delimiters"])
        )
        self.page_up_down_select_row = kwargs["page_up_down_select_row"]
        self.expand_sheet_if_paste_too_big = kwargs["expand_sheet_if_paste_too_big"]
        self.paste_insert_column_limit = kwargs["paste_insert_column_limit"]
        self.paste_insert_row_limit = kwargs["paste_insert_row_limit"]
        self.arrow_key_down_right_scroll_page = kwargs["arrow_key_down_right_scroll_page"]
        self.cell_auto_resize_enabled = kwargs["enable_edit_cell_auto_resize"]
        self.auto_resize_columns = kwargs["auto_resize_columns"]
        self.auto_resize_rows = kwargs["auto_resize_rows"]
        self.allow_auto_resize_columns = True
        self.allow_auto_resize_rows = True
        self.edit_cell_validation = kwargs["edit_cell_validation"]
        self.display_selected_fg_over_highlights = kwargs["display_selected_fg_over_highlights"]
        self.show_index = kwargs["show_index"]
        self.show_header = kwargs["show_header"]
        self.selected_rows_to_end_of_window = kwargs["selected_rows_to_end_of_window"]
        self.horizontal_grid_to_end_of_window = kwargs["horizontal_grid_to_end_of_window"]
        self.vertical_grid_to_end_of_window = kwargs["vertical_grid_to_end_of_window"]
        self.empty_horizontal = kwargs["empty_horizontal"]
        self.empty_vertical = kwargs["empty_vertical"]
        self.show_vertical_grid = kwargs["show_vertical_grid"]
        self.show_horizontal_grid = kwargs["show_horizontal_grid"]
        self.min_row_height = 0
        self.min_header_height = 0
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
        self.ctrl_selection_binding_func = None
        self.select_all_binding_func = None

        self.single_selection_enabled = False
        # with this mode every left click adds the cell to selected cells
        self.toggle_selection_enabled = False
        self.show_dropdown_borders = kwargs["show_dropdown_borders"]
        self.drag_selection_enabled = False
        self.select_all_enabled = False
        self.undo_enabled = False
        self.cut_enabled = False
        self.copy_enabled = False
        self.paste_enabled = False
        self.delete_key_enabled = False
        self.rc_select_enabled = False
        self.ctrl_select_enabled = False
        self.rc_delete_column_enabled = False
        self.rc_insert_column_enabled = False
        self.rc_delete_row_enabled = False
        self.rc_insert_row_enabled = False
        self.rc_popup_menus_enabled = False
        self.edit_cell_enabled = False
        self.text_editor_loc = None
        self.show_selected_cells_border = kwargs["show_selected_cells_border"]
        self.new_row_width = 0
        self.new_header_height = 0
        self.row_width_resize_bbox = tuple()
        self.header_height_resize_bbox = tuple()
        self.CH = kwargs["column_headers_canvas"]
        self.CH.MT = self
        self.CH.RI = kwargs["row_index_canvas"]
        self.RI = kwargs["row_index_canvas"]
        self.RI.MT = self
        self.RI.CH = kwargs["column_headers_canvas"]
        self.TL = None  # is set from within TopLeftRectangle() __init__
        self.all_columns_displayed = True
        self.all_rows_displayed = True
        self.align = kwargs["align"]
        self.table_font = kwargs["font"]
        self.font_fam = kwargs["font"][0]
        self.font_sze = kwargs["font"][1]
        self.font_wgt = kwargs["font"][2]
        self.index_font = kwargs["index_font"]
        self.index_font_fam = kwargs["index_font"][0]
        self.index_font_sze = kwargs["index_font"][1]
        self.index_font_wgt = kwargs["index_font"][2]
        self.header_font = kwargs["header_font"]
        self.header_font_fam = kwargs["header_font"][0]
        self.header_font_sze = kwargs["header_font"][1]
        self.header_font_wgt = kwargs["header_font"][2]
        self.txt_measure_canvas = tk.Canvas(self)
        self.txt_measure_canvas_text = self.txt_measure_canvas.create_text(0, 0, text="", font=self.table_font)
        self.text_editor = None
        self.text_editor_id = None

        self.max_row_height = float(kwargs["max_row_height"])
        self.max_index_width = float(kwargs["max_index_width"])
        self.max_column_width = float(kwargs["max_column_width"])
        self.max_header_height = float(kwargs["max_header_height"])
        if kwargs["row_index_width"] is None:
            self.RI.set_width(70)
            self.default_index_width = 70
        else:
            self.RI.set_width(kwargs["row_index_width"])
            self.default_index_width = kwargs["row_index_width"]
        self.default_header_height = (
            kwargs["header_height"] if isinstance(kwargs["header_height"], str) else "pixels",
            kwargs["header_height"]
            if isinstance(kwargs["header_height"], int)
            else self.get_lines_cell_height(int(kwargs["header_height"]), font=self.header_font),
        )
        self.default_column_width = kwargs["column_width"]
        self.default_row_height = (
            kwargs["row_height"] if isinstance(kwargs["row_height"], str) else "pixels",
            kwargs["row_height"]
            if isinstance(kwargs["row_height"], int)
            else self.get_lines_cell_height(int(kwargs["row_height"])),
        )
        self.set_table_font_help()
        self.set_header_font_help()
        self.set_index_font_help()
        self.data = kwargs["data_reference"]
        if isinstance(self.data, (list, tuple)):
            self.data = kwargs["data_reference"]
        else:
            self.data = []
        if not self.data:
            if (
                isinstance(kwargs["total_rows"], int)
                and isinstance(kwargs["total_cols"], int)
                and kwargs["total_rows"] > 0
                and kwargs["total_cols"] > 0
            ):
                self.data = [list(repeat("", kwargs["total_cols"])) for i in range(kwargs["total_rows"])]
        _header = kwargs["header"] if kwargs["header"] is not None else kwargs["headers"]
        if isinstance(_header, int):
            self._headers = _header
        else:
            if _header:
                self._headers = _header
            else:
                self._headers = []
        _row_index = kwargs["index"] if kwargs["index"] is not None else kwargs["row_index"]
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
        self.display_rows(
            rows=kwargs["displayed_rows"],
            all_rows_displayed=kwargs["all_rows_displayed"],
            reset_row_positions=False,
            deselect_all=False,
        )
        self.reset_row_positions()
        self.display_columns(
            columns=kwargs["displayed_columns"],
            all_columns_displayed=kwargs["all_columns_displayed"],
            reset_col_positions=False,
            deselect_all=False,
        )
        self.reset_col_positions()
        self.table_grid_fg = kwargs["table_grid_fg"]
        self.table_fg = kwargs["table_fg"]
        self.table_selected_cells_border_fg = kwargs["table_selected_cells_border_fg"]
        self.table_selected_cells_bg = kwargs["table_selected_cells_bg"]
        self.table_selected_cells_fg = kwargs["table_selected_cells_fg"]
        self.table_selected_rows_border_fg = kwargs["table_selected_rows_border_fg"]
        self.table_selected_rows_bg = kwargs["table_selected_rows_bg"]
        self.table_selected_rows_fg = kwargs["table_selected_rows_fg"]
        self.table_selected_columns_border_fg = kwargs["table_selected_columns_border_fg"]
        self.table_selected_columns_bg = kwargs["table_selected_columns_bg"]
        self.table_selected_columns_fg = kwargs["table_selected_columns_fg"]
        self.table_bg = kwargs["table_bg"]
        self.popup_menu_font = kwargs["popup_menu_font"]
        self.popup_menu_fg = kwargs["popup_menu_fg"]
        self.popup_menu_bg = kwargs["popup_menu_bg"]
        self.popup_menu_highlight_bg = kwargs["popup_menu_highlight_bg"]
        self.popup_menu_highlight_fg = kwargs["popup_menu_highlight_fg"]
        self.rc_popup_menu = None
        self.empty_rc_popup_menu = None
        self.basic_bindings()
        self.create_rc_menus()

    def refresh(self, event=None):
        self.main_table_redraw_grid_and_text(True, True)

    def window_configured(self, event):
        w = self.parentframe.winfo_width()
        if w != self.parentframe_width:
            self.parentframe_width = w
            self.allow_auto_resize_columns = True
        h = self.parentframe.winfo_height()
        if h != self.parentframe_height:
            self.parentframe_height = h
            self.allow_auto_resize_rows = True
        self.main_table_redraw_grid_and_text(True, True)

    def basic_bindings(self, enable=True):
        if enable:
            self.bind("<Configure>", self.window_configured)
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
            self.bind(rc_binding, self.rc)
            self.bind(f"<{ctrl_key}-ButtonPress-1>", self.ctrl_b1_press)
            self.CH.bind(f"<{ctrl_key}-ButtonPress-1>", self.CH.ctrl_b1_press)
            self.RI.bind(f"<{ctrl_key}-ButtonPress-1>", self.RI.ctrl_b1_press)
            self.bind(f"<{ctrl_key}-Shift-ButtonPress-1>", self.ctrl_shift_b1_press)
            self.CH.bind(f"<{ctrl_key}-Shift-ButtonPress-1>", self.CH.ctrl_shift_b1_press)
            self.RI.bind(f"<{ctrl_key}-Shift-ButtonPress-1>", self.RI.ctrl_shift_b1_press)
            self.bind(f"<{ctrl_key}-B1-Motion>", self.ctrl_b1_motion)
            self.CH.bind(f"<{ctrl_key}-B1-Motion>", self.CH.ctrl_b1_motion)
            self.RI.bind(f"<{ctrl_key}-B1-Motion>", self.RI.ctrl_b1_motion)
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
            self.unbind(rc_binding)
            self.unbind(f"<{ctrl_key}-ButtonPress-1>")
            self.CH.unbind(f"<{ctrl_key}-ButtonPress-1>")
            self.RI.unbind(f"<{ctrl_key}-ButtonPress-1>")
            self.unbind(f"<{ctrl_key}-Shift-ButtonPress-1>")
            self.CH.unbind(f"<{ctrl_key}-Shift-ButtonPress-1>")
            self.RI.unbind(f"<{ctrl_key}-Shift-ButtonPress-1>")
            self.unbind(f"<{ctrl_key}-B1-Motion>")
            self.CH.unbind(f"<{ctrl_key}-B1-Motion>")
            self.RI.unbind(f"<{ctrl_key}-B1-Motion>")

    def show_ctrl_outline(
        self,
        canvas="table",
        start_cell=(0, 0),
        end_cell=(0, 0),
        dash=(20, 20),
        outline=None,
        delete_on_timer=True,
    ):
        self.create_ctrl_outline(
            self.col_positions[start_cell[0]] + 1,
            self.row_positions[start_cell[1]] + 1,
            self.col_positions[end_cell[0]] - 1,
            self.row_positions[end_cell[1]] - 1,
            fill="",
            dash=dash,
            width=3,
            outline=self.RI.resizing_line_fg if outline is None else outline,
            tag="ctrl",
        )
        if delete_on_timer:
            self.after(1500, self.delete_ctrl_outlines)

    def create_ctrl_outline(self, x1, y1, x2, y2, fill, dash, width, outline, tag):
        if self.hidd_ctrl_outline:
            t, sh = self.hidd_ctrl_outline.popitem()
            self.coords(t, x1, y1, x2, y2)
            if sh:
                self.itemconfig(t, fill=fill, dash=dash, width=width, outline=outline, tag=tag)
            else:
                self.itemconfig(
                    t,
                    fill=fill,
                    dash=dash,
                    width=width,
                    outline=outline,
                    tag=tag,
                    state="normal",
                )
            self.lift(t)
        else:
            t = self.create_rectangle(
                x1,
                y1,
                x2,
                y2,
                fill=fill,
                dash=dash,
                width=width,
                outline=outline,
                tag=tag,
            )
        self.disp_ctrl_outline[t] = True

    def delete_ctrl_outlines(self):
        self.hidd_ctrl_outline.update(self.disp_ctrl_outline)
        self.disp_ctrl_outline = {}
        for t, sh in self.hidd_ctrl_outline.items():
            if sh:
                self.itemconfig(t, state="hidden")
                self.hidd_ctrl_outline[t] = False

    def get_ctrl_x_c_boxes(self):
        currently_selected = self.currently_selected()
        boxes = {}
        if currently_selected.type_ in ("cell", "column"):
            for item in chain(self.find_withtag("cells"), self.find_withtag("columns")):
                alltags = self.gettags(item)
                boxes[tuple(int(e) for e in alltags[1].split("_") if e)] = alltags[0]
            curr_box = self.find_last_selected_box_with_current_from_boxes(currently_selected, boxes)
            maxrows = curr_box[2] - curr_box[0]
            for box in tuple(boxes):
                if box[2] - box[0] != maxrows:
                    del boxes[box]
            return boxes, maxrows
        else:
            for item in self.find_withtag("rows"):
                boxes[tuple(int(e) for e in self.gettags(item)[1].split("_") if e)] = "rows"
            return boxes

    def ctrl_c(self, event=None):
        currently_selected = self.currently_selected()
        if currently_selected:
            s = io.StringIO()
            writer = csv.writer(
                s,
                dialect=csv.excel_tab,
                delimiter=self.to_clipboard_delimiter,
                quotechar=self.to_clipboard_quotechar,
                lineterminator=self.to_clipboard_lineterminator,
            )
            rows = []
            if currently_selected.type_ in ("cell", "column"):
                boxes, maxrows = self.get_ctrl_x_c_boxes()
                if self.extra_begin_ctrl_c_func is not None:
                    try:
                        self.extra_begin_ctrl_c_func(CtrlKeyEvent("begin_ctrl_c", boxes, currently_selected, tuple()))
                    except Exception:
                        return
                for rn in range(maxrows):
                    row = []
                    for r1, c1, r2, c2 in boxes:
                        if r2 - r1 < maxrows:
                            continue
                        datarn = (r1 + rn) if self.all_rows_displayed else self.displayed_rows[r1 + rn]
                        for c in range(c1, c2):
                            datacn = c if self.all_columns_displayed else self.displayed_columns[c]
                            row.append(self.get_cell_clipboard(datarn, datacn))
                    writer.writerow(row)
                    rows.append(row)
            else:
                boxes = self.get_ctrl_x_c_boxes()
                if self.extra_begin_ctrl_c_func is not None:
                    try:
                        self.extra_begin_ctrl_c_func(CtrlKeyEvent("begin_ctrl_c", boxes, currently_selected, tuple()))
                    except Exception:
                        return
                for r1, c1, r2, c2 in boxes:
                    for rn in range(r2 - r1):
                        row = []
                        datarn = (r1 + rn) if self.all_rows_displayed else self.displayed_rows[r1 + rn]
                        for c in range(c1, c2):
                            datacn = c if self.all_columns_displayed else self.displayed_columns[c]
                            row.append(self.get_cell_clipboard(datarn, datacn))
                        writer.writerow(row)
                        rows.append(row)
            for r1, c1, r2, c2 in boxes:
                self.show_ctrl_outline(canvas="table", start_cell=(c1, r1), end_cell=(c2, r2))
            self.clipboard_clear()
            self.clipboard_append(s.getvalue())
            self.update_idletasks()
            if self.extra_end_ctrl_c_func is not None:
                self.extra_end_ctrl_c_func(CtrlKeyEvent("end_ctrl_c", boxes, currently_selected, rows))

    def ctrl_x(self, event=None):
        if not self.anything_selected():
            return
        undo_storage = {}
        s = io.StringIO()
        writer = csv.writer(
            s,
            dialect=csv.excel_tab,
            delimiter=self.to_clipboard_delimiter,
            quotechar=self.to_clipboard_quotechar,
            lineterminator=self.to_clipboard_lineterminator,
        )
        currently_selected = self.currently_selected()
        rows = []
        changes = 0
        if currently_selected.type_ in ("cell", "column"):
            boxes, maxrows = self.get_ctrl_x_c_boxes()
            if self.extra_begin_ctrl_x_func is not None:
                try:
                    self.extra_begin_ctrl_x_func(CtrlKeyEvent("begin_ctrl_x", boxes, currently_selected, tuple()))
                except Exception:
                    return
            for rn in range(maxrows):
                row = []
                for r1, c1, r2, c2 in boxes:
                    if r2 - r1 < maxrows:
                        continue
                    datarn = (r1 + rn) if self.all_rows_displayed else self.displayed_rows[r1 + rn]
                    for c in range(c1, c2):
                        datacn = c if self.all_columns_displayed else self.displayed_columns[c]
                        row.append(self.get_cell_clipboard(datarn, datacn))
                writer.writerow(row)
                rows.append(row)
            for rn in range(maxrows):
                for r1, c1, r2, c2 in boxes:
                    if r2 - r1 < maxrows:
                        continue
                    datarn = (r1 + rn) if self.all_rows_displayed else self.displayed_rows[r1 + rn]
                    for c in range(c1, c2):
                        datacn = c if self.all_columns_displayed else self.displayed_columns[c]
                        if self.input_valid_for_cell(datarn, datacn, ""):
                            if self.undo_enabled:
                                undo_storage[(datarn, datacn)] = self.get_cell_data(datarn, datacn)
                            self.set_cell_data(datarn, datacn, "")
                            changes += 1
        else:
            boxes = self.get_ctrl_x_c_boxes()
            if self.extra_begin_ctrl_x_func is not None:
                try:
                    self.extra_begin_ctrl_x_func(CtrlKeyEvent("begin_ctrl_x", boxes, currently_selected, tuple()))
                except Exception:
                    return
            for r1, c1, r2, c2 in boxes:
                for rn in range(r2 - r1):
                    row = []
                    datarn = (r1 + rn) if self.all_rows_displayed else self.displayed_rows[r1 + rn]
                    for c in range(c1, c2):
                        datacn = c if self.all_columns_displayed else self.displayed_columns[c]
                        row.append(self.get_cell_data(datarn, datacn))
                    writer.writerow(row)
                    rows.append(row)
            for r1, c1, r2, c2 in boxes:
                for rn in range(r2 - r1):
                    datarn = (r1 + rn) if self.all_rows_displayed else self.displayed_rows[r1 + rn]
                    for c in range(c1, c2):
                        datacn = c if self.all_columns_displayed else self.displayed_columns[c]
                        if self.input_valid_for_cell(datarn, datacn, ""):
                            if self.undo_enabled:
                                undo_storage[(datarn, datacn)] = self.get_cell_data(datarn, datacn)
                            self.set_cell_data(datarn, datacn, "")
                            changes += 1
        if changes and self.undo_enabled:
            self.undo_storage.append(
                zlib.compress(pickle.dumps(("edit_cells", undo_storage, boxes, currently_selected)))
            )
        self.clipboard_clear()
        self.clipboard_append(s.getvalue())
        self.update_idletasks()
        self.refresh()
        for r1, c1, r2, c2 in boxes:
            self.show_ctrl_outline(canvas="table", start_cell=(c1, r1), end_cell=(c2, r2))
        if self.extra_end_ctrl_x_func is not None:
            self.extra_end_ctrl_x_func(CtrlKeyEvent("end_ctrl_x", boxes, currently_selected, rows))
        self.parentframe.emit_event("<<SheetModified>>")

    def find_last_selected_box_with_current(self, currently_selected):
        if currently_selected.type_ in ("cell", "column"):
            boxes, maxrows = self.get_ctrl_x_c_boxes()
        else:
            boxes = self.get_ctrl_x_c_boxes()
        return self.find_last_selected_box_with_current_from_boxes(currently_selected, boxes)

    def find_last_selected_box_with_current_from_boxes(self, currently_selected, boxes):
        for (r1, c1, r2, c2), type_ in boxes.items():
            if (
                type_[:2] == currently_selected.type_[:2]
                and currently_selected.row >= r1
                and currently_selected.row <= r2
                and currently_selected.column >= c1
                and currently_selected.column <= c2
            ):
                if (self.last_selected and self.last_selected == (r1, c1, r2, c2, type_)) or not self.last_selected:
                    return (r1, c1, r2, c2)
        return (
            currently_selected.row,
            currently_selected.column,
            currently_selected.row + 1,
            currently_selected.column + 1,
        )

    def ctrl_v(self, event=None):
        if not self.expand_sheet_if_paste_too_big and (len(self.col_positions) == 1 or len(self.row_positions) == 1):
            return
        currently_selected = self.currently_selected()
        if currently_selected:
            selected_r = currently_selected[0]
            selected_c = currently_selected[1]
        elif not currently_selected and not self.expand_sheet_if_paste_too_big:
            return
        else:
            if not self.data:
                selected_c, selected_r = 0, 0
            else:
                if len(self.col_positions) == 1 and len(self.row_positions) > 1:
                    selected_c, selected_r = 0, len(self.row_positions) - 1
                elif len(self.row_positions) == 1 and len(self.col_positions) > 1:
                    selected_c, selected_r = len(self.col_positions) - 1, 0
                elif len(self.row_positions) > 1 and len(self.col_positions) > 1:
                    selected_c, selected_r = 0, len(self.row_positions) - 1
        try:
            data = self.clipboard_get()
        except Exception:
            return
        try:
            dialect = csv.Sniffer().sniff(data, delimiters=self.from_clipboard_delimiters)
        except Exception:
            dialect = csv.excel_tab
        data = list(csv.reader(io.StringIO(data), dialect=dialect, skipinitialspace=True))
        if not data:
            return
        numcols = len(max(data, key=len))
        numrows = len(data)
        for rn, r in enumerate(data):
            if len(r) < numcols:
                data[rn].extend(list(repeat("", numcols - len(r))))
        (
            lastbox_r1,
            lastbox_c1,
            lastbox_r2,
            lastbox_c2,
        ) = self.find_last_selected_box_with_current(currently_selected)
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
            if selected_c + numcols > len(self.col_positions) - 1:
                added_cols = selected_c + numcols - len(self.col_positions) + 1
                if (
                    isinstance(self.paste_insert_column_limit, int)
                    and self.paste_insert_column_limit < len(self.col_positions) - 1 + added_cols
                ):
                    added_cols = self.paste_insert_column_limit - len(self.col_positions) - 1
                if added_cols > 0:
                    self.insert_col_positions(widths=int(added_cols))
                if not self.all_columns_displayed:
                    total_data_cols = self.total_data_cols()
                    self.displayed_columns.extend(list(range(total_data_cols, total_data_cols + added_cols)))
            if selected_r + numrows > len(self.row_positions) - 1:
                added_rows = selected_r + numrows - len(self.row_positions) + 1
                if (
                    isinstance(self.paste_insert_row_limit, int)
                    and self.paste_insert_row_limit < len(self.row_positions) - 1 + added_rows
                ):
                    added_rows = self.paste_insert_row_limit - len(self.row_positions) - 1
                if added_rows > 0:
                    self.insert_row_positions(heights=int(added_rows))
                if not self.all_rows_displayed:
                    total_data_rows = self.total_data_rows()
                    self.displayed_rows.extend(list(range(total_data_rows, total_data_rows + added_rows)))
            added_rows_cols = (added_rows, added_cols)
        else:
            added_rows_cols = (0, 0)
        if selected_c + numcols > len(self.col_positions) - 1:
            numcols = len(self.col_positions) - 1 - selected_c
        if selected_r + numrows > len(self.row_positions) - 1:
            numrows = len(self.row_positions) - 1 - selected_r
        if self.extra_begin_ctrl_v_func is not None or self.extra_end_ctrl_v_func is not None:
            rows = [
                [data[ndr][ndc] for ndc, c in enumerate(range(selected_c, selected_c + numcols))]
                for ndr, r in enumerate(range(selected_r, selected_r + numrows))
            ]
        if self.extra_begin_ctrl_v_func is not None:
            try:
                self.extra_begin_ctrl_v_func(PasteEvent("begin_ctrl_v", currently_selected, rows))
            except Exception:
                return
        changes = 0
        for ndr, r in enumerate(range(selected_r, selected_r + numrows)):
            datarn = r if self.all_rows_displayed else self.displayed_rows[r]
            for ndc, c in enumerate(range(selected_c, selected_c + numcols)):
                datacn = c if self.all_columns_displayed else self.displayed_columns[c]
                if self.input_valid_for_cell(datarn, datacn, data[ndr][ndc]):
                    if self.undo_enabled:
                        undo_storage[(datarn, datacn)] = self.get_cell_data(datarn, datacn)
                    self.set_cell_data(datarn, datacn, data[ndr][ndc])
                    changes += 1
        if self.expand_sheet_if_paste_too_big and self.undo_enabled:
            self.equalize_data_row_lengths()
        self.deselect("all")
        if changes and self.undo_enabled:
            self.undo_storage.append(
                zlib.compress(
                    pickle.dumps(
                        (
                            "edit_cells_paste",
                            undo_storage,
                            {
                                (
                                    selected_r,
                                    selected_c,
                                    selected_r + numrows,
                                    selected_c + numcols,
                                ): "cells"
                            },  # boxes
                            currently_selected,
                            added_rows_cols,
                        )
                    )
                )
            )
        self.create_selected(selected_r, selected_c, selected_r + numrows, selected_c + numcols, "cells")
        self.set_currently_selected(selected_r, selected_c, type_="cell")
        self.see(
            r=selected_r,
            c=selected_c,
            keep_yscroll=False,
            keep_xscroll=False,
            bottom_right_corner=False,
            check_cell_visibility=True,
            redraw=False,
        )
        self.refresh()
        if self.extra_end_ctrl_v_func is not None:
            self.extra_end_ctrl_v_func(PasteEvent("end_ctrl_v", currently_selected, rows))
        self.parentframe.emit_event("<<SheetModified>>")

    def delete_key(self, event=None):
        if not self.anything_selected():
            return
        currently_selected = self.currently_selected()
        undo_storage = {}
        boxes = {}
        for item in chain(
            self.find_withtag("cells"),
            self.find_withtag("rows"),
            self.find_withtag("columns"),
        ):
            alltags = self.gettags(item)
            box = tuple(int(e) for e in alltags[1].split("_") if e)
            boxes[box] = alltags[0]
        if self.extra_begin_delete_key_func is not None:
            try:
                self.extra_begin_delete_key_func(CtrlKeyEvent("begin_delete_key", boxes, currently_selected, tuple()))
            except Exception:
                return
        changes = 0
        for r1, c1, r2, c2 in boxes:
            for r in range(r1, r2):
                datarn = r if self.all_rows_displayed else self.displayed_rows[r]
                for c in range(c1, c2):
                    datacn = c if self.all_columns_displayed else self.displayed_columns[c]
                    if self.input_valid_for_cell(datarn, datacn, ""):
                        if self.undo_enabled:
                            undo_storage[(datarn, datacn)] = self.get_cell_data(datarn, datacn)
                        self.set_cell_data(datarn, datacn, "")
                        changes += 1
        if self.extra_end_delete_key_func is not None:
            self.extra_end_delete_key_func(CtrlKeyEvent("end_delete_key", boxes, currently_selected, undo_storage))
        if changes and self.undo_enabled:
            self.undo_storage.append(
                zlib.compress(pickle.dumps(("edit_cells", undo_storage, boxes, currently_selected)))
            )
        self.refresh()
        self.parentframe.emit_event("<<SheetModified>>")

    def move_columns_adjust_options_dict(
        self,
        col,
        to_move_min,
        num_cols,
        move_data=True,
        create_selections=True,
        index_type="displayed",
    ):
        c = int(col)
        to_move_max = to_move_min + num_cols
        to_del = to_move_max + num_cols
        orig_selected = list(range(to_move_min, to_move_min + num_cols))
        if index_type == "displayed":
            self.deselect("all", redraw=False)
            cws = [
                int(b - a)
                for a, b in zip(
                    self.col_positions,
                    islice(self.col_positions, 1, len(self.col_positions)),
                )
            ]
            if to_move_min > c:
                cws[c:c] = cws[to_move_min:to_move_max]
                cws[to_move_max:to_del] = []
            else:
                cws[c + 1 : c + 1] = cws[to_move_min:to_move_max]
                cws[to_move_min:to_move_max] = []
            self.col_positions = list(accumulate(chain([0], (width for width in cws))))
            if c + num_cols > len(self.col_positions):
                new_selected = tuple(
                    range(
                        len(self.col_positions) - 1 - num_cols,
                        len(self.col_positions) - 1,
                    )
                )
                if create_selections:
                    self.create_selected(
                        0,
                        len(self.col_positions) - 1 - num_cols,
                        len(self.row_positions) - 1,
                        len(self.col_positions) - 1,
                        "columns",
                    )
            else:
                if to_move_min > c:
                    new_selected = tuple(range(c, c + num_cols))
                    if create_selections:
                        self.create_selected(0, c, len(self.row_positions) - 1, c + num_cols, "columns")
                else:
                    new_selected = tuple(range(c + 1 - num_cols, c + 1))
                    if create_selections:
                        self.create_selected(
                            0,
                            c + 1 - num_cols,
                            len(self.row_positions) - 1,
                            c + 1,
                            "columns",
                        )
        elif index_type == "data":
            if to_move_min > c:
                new_selected = tuple(range(c, c + num_cols))
            else:
                new_selected = tuple(range(c + 1 - num_cols, c + 1))
        newcolsdct = {t1: t2 for t1, t2 in zip(orig_selected, new_selected)}
        if self.all_columns_displayed or index_type != "displayed":
            dispset = {}
            if to_move_min > c:
                if move_data:
                    extend_idx = to_move_max - 1
                    for rn in range(len(self.data)):
                        if to_move_max > len(self.data[rn]):
                            self.fix_row_len(rn, extend_idx)
                        self.data[rn][c:c] = self.data[rn][to_move_min:to_move_max]
                        self.data[rn][to_move_max:to_del] = []
                    self.CH.fix_header(extend_idx)
                    if isinstance(self._headers, list) and self._headers:
                        self._headers[c:c] = self._headers[to_move_min:to_move_max]
                        self._headers[to_move_max:to_del] = []
                self.CH.cell_options = {
                    newcolsdct[k] if k in newcolsdct else k + num_cols if k < to_move_min and k >= c else k: v
                    for k, v in self.CH.cell_options.items()
                }
                self.cell_options = {
                    (k[0], newcolsdct[k[1]])
                    if k[1] in newcolsdct
                    else (k[0], k[1] + num_cols)
                    if k[1] < to_move_min and k[1] >= c
                    else k: v
                    for k, v in self.cell_options.items()
                }
                self.col_options = {
                    newcolsdct[k] if k in newcolsdct else k + num_cols if k < to_move_min and k >= c else k: v
                    for k, v in self.col_options.items()
                }
                if index_type != "displayed":
                    self.displayed_columns = sorted(
                        int(newcolsdct[k])
                        if k in newcolsdct
                        else k + num_cols
                        if k < to_move_min and k >= c
                        else int(k)
                        for k in self.displayed_columns
                    )
            else:
                c += 1
                if move_data:
                    extend_idx = c - 1
                    for rn in range(len(self.data)):
                        if c > len(self.data[rn]):
                            self.fix_row_len(rn, extend_idx)
                        self.data[rn][c:c] = self.data[rn][to_move_min:to_move_max]
                        self.data[rn][to_move_min:to_move_max] = []
                    self.CH.fix_header(extend_idx)
                    if isinstance(self._headers, list) and self._headers:
                        self._headers[c:c] = self._headers[to_move_min:to_move_max]
                        self._headers[to_move_min:to_move_max] = []
                self.CH.cell_options = {
                    newcolsdct[k] if k in newcolsdct else k - num_cols if k < c and k > to_move_min else k: v
                    for k, v in self.CH.cell_options.items()
                }
                self.cell_options = {
                    (k[0], newcolsdct[k[1]])
                    if k[1] in newcolsdct
                    else (k[0], k[1] - num_cols)
                    if k[1] < c and k[1] > to_move_min
                    else k: v
                    for k, v in self.cell_options.items()
                }
                self.col_options = {
                    newcolsdct[k] if k in newcolsdct else k - num_cols if k < c and k > to_move_min else k: v
                    for k, v in self.col_options.items()
                }
                if index_type != "displayed":
                    self.displayed_columns = sorted(
                        int(newcolsdct[k]) if k in newcolsdct else k - num_cols if k < c and k > to_move_min else int(k)
                        for k in self.displayed_columns
                    )
        else:
            # moves data around, not displayed columns indexes
            # which remain sorted and the same after drop and drop
            if to_move_min > c:
                dispset = {
                    a: b
                    for a, b in zip(
                        self.displayed_columns,
                        (
                            self.displayed_columns[:c]
                            + self.displayed_columns[to_move_min : to_move_min + num_cols]
                            + self.displayed_columns[c:to_move_min]
                            + self.displayed_columns[to_move_min + num_cols :]
                        ),
                    )
                }
            else:
                dispset = {
                    a: b
                    for a, b in zip(
                        self.displayed_columns,
                        (
                            self.displayed_columns[:to_move_min]
                            + self.displayed_columns[to_move_min + num_cols : c + 1]
                            + self.displayed_columns[to_move_min : to_move_min + num_cols]
                            + self.displayed_columns[c + 1 :]
                        ),
                    )
                }
            # has to pick up elements from all over the place in the original row
            # building an entirely new row is best due to permutations of hidden columns
            if move_data:
                max_len = max(chain(dispset, dispset.values())) + 1
                max_idx = max_len - 1
                for rn in range(len(self.data)):
                    if max_len > len(self.data[rn]):
                        self.fix_row_len(rn, max_idx)
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
                self.CH.fix_header(max_idx)
                if isinstance(self._headers, list) and self._headers:
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
                self.cell_options = {
                    (k[0], dispset[k[1]]) if k[1] in dispset else k: v for k, v in self.cell_options.items()
                }
                self.col_options = {dispset[k] if k in dispset else k: v for k, v in self.col_options.items()}
        return new_selected, {b: a for a, b in dispset.items()}

    def move_rows_adjust_options_dict(
        self,
        row,
        to_move_min,
        num_rows,
        move_data=True,
        create_selections=True,
        index_type="displayed",
    ):
        r = int(row)
        to_move_max = to_move_min + num_rows
        to_del = to_move_max + num_rows
        orig_selected = list(range(to_move_min, to_move_min + num_rows))
        if index_type == "displayed":
            self.deselect("all", redraw=False)
            rhs = [
                int(b - a)
                for a, b in zip(
                    self.row_positions,
                    islice(self.row_positions, 1, len(self.row_positions)),
                )
            ]
            if to_move_min > r:
                rhs[r:r] = rhs[to_move_min:to_move_max]
                rhs[to_move_max:to_del] = []
            else:
                rhs[r + 1 : r + 1] = rhs[to_move_min:to_move_max]
                rhs[to_move_min:to_move_max] = []
            self.row_positions = list(accumulate(chain([0], (height for height in rhs))))
            if r + num_rows > len(self.row_positions):
                new_selected = tuple(
                    range(
                        len(self.row_positions) - 1 - num_rows,
                        len(self.row_positions) - 1,
                    )
                )
                if create_selections:
                    self.create_selected(
                        len(self.row_positions) - 1 - num_rows,
                        0,
                        len(self.row_positions) - 1,
                        len(self.col_positions) - 1,
                        "rows",
                    )
            else:
                if to_move_min > r:
                    new_selected = tuple(range(r, r + num_rows))
                    if create_selections:
                        self.create_selected(r, 0, r + num_rows, len(self.col_positions) - 1, "rows")
                else:
                    new_selected = tuple(range(r + 1 - num_rows, r + 1))
                    if create_selections:
                        self.create_selected(
                            r + 1 - num_rows,
                            0,
                            r + 1,
                            len(self.col_positions) - 1,
                            "rows",
                        )
        elif index_type == "data":
            if to_move_min > r:
                new_selected = tuple(range(r, r + num_rows))
            else:
                new_selected = tuple(range(r + 1 - num_rows, r + 1))
        newrowsdct = {t1: t2 for t1, t2 in zip(orig_selected, new_selected)}
        if self.all_rows_displayed or index_type != "displayed":
            dispset = {}
            if to_move_min > r:
                if move_data:
                    extend_idx = to_move_max - 1
                    if to_move_max > len(self.data):
                        self.fix_data_len(extend_idx)
                    self.data[r:r] = self.data[to_move_min:to_move_max]
                    self.data[to_move_max:to_del] = []
                    self.RI.fix_index(extend_idx)
                    if isinstance(self._row_index, list) and self._row_index:
                        self._row_index[r:r] = self._row_index[to_move_min:to_move_max]
                        self._row_index[to_move_max:to_del] = []
                self.RI.cell_options = {
                    newrowsdct[k] if k in newrowsdct else k + num_rows if k < to_move_min and k >= r else k: v
                    for k, v in self.RI.cell_options.items()
                }
                self.cell_options = {
                    (newrowsdct[k[0]], k[1])
                    if k[0] in newrowsdct
                    else (k[0] + num_rows, k[1])
                    if k[0] < to_move_min and k[0] >= r
                    else k: v
                    for k, v in self.cell_options.items()
                }
                self.row_options = {
                    newrowsdct[k] if k in newrowsdct else k + num_rows if k < to_move_min and k >= r else k: v
                    for k, v in self.row_options.items()
                }
                if index_type != "displayed":
                    self.displayed_rows = sorted(
                        int(newrowsdct[k])
                        if k in newrowsdct
                        else k + num_rows
                        if k < to_move_min and k >= r
                        else int(k)
                        for k in self.displayed_rows
                    )
            else:
                r += 1
                if move_data:
                    extend_idx = r - 1
                    if r > len(self.data):
                        self.fix_data_len(extend_idx)
                    self.data[r:r] = self.data[to_move_min:to_move_max]
                    self.data[to_move_min:to_move_max] = []
                    self.RI.fix_index(extend_idx)
                    if isinstance(self._row_index, list) and self._row_index:
                        self._row_index[r:r] = self._row_index[to_move_min:to_move_max]
                        self._row_index[to_move_min:to_move_max] = []
                self.RI.cell_options = {
                    newrowsdct[k] if k in newrowsdct else k - num_rows if k < r and k > to_move_min else k: v
                    for k, v in self.RI.cell_options.items()
                }
                self.cell_options = {
                    (newrowsdct[k[0]], k[1])
                    if k[0] in newrowsdct
                    else (k[0] - num_rows, k[1])
                    if k[0] < r and k[0] > to_move_min
                    else k: v
                    for k, v in self.cell_options.items()
                }
                self.row_options = {
                    newrowsdct[k] if k in newrowsdct else k - num_rows if k < r and k > to_move_min else k: v
                    for k, v in self.row_options.items()
                }
                if index_type != "displayed":
                    self.displayed_rows = sorted(
                        int(newrowsdct[k]) if k in newrowsdct else k - num_rows if k < r and k > to_move_min else int(k)
                        for k in self.displayed_rows
                    )
        else:
            # moves data around, not displayed rows indexes
            # which remain sorted and the same after drop and drop
            if to_move_min > r:
                dispset = {
                    a: b
                    for a, b in zip(
                        self.displayed_rows,
                        (
                            self.displayed_rows[:r]
                            + self.displayed_rows[to_move_min : to_move_min + num_rows]
                            + self.displayed_rows[r:to_move_min]
                            + self.displayed_rows[to_move_min + num_rows :]
                        ),
                    )
                }
            else:
                dispset = {
                    a: b
                    for a, b in zip(
                        self.displayed_rows,
                        (
                            self.displayed_rows[:to_move_min]
                            + self.displayed_rows[to_move_min + num_rows : r + 1]
                            + self.displayed_rows[to_move_min : to_move_min + num_rows]
                            + self.displayed_rows[r + 1 :]
                        ),
                    )
                }
            # has to pick up rows from all over the place in the original sheet
            # building an entirely new sheet is best due to permutations of hidden rows
            if move_data:
                max_len = max(chain(dispset, dispset.values())) + 1
                if len(self.data) < max_len:
                    self.fix_data_len(max_len - 1)
                new = []
                idx = 0
                done = set()
                while len(new) < len(self.data):
                    if idx in dispset and idx not in done:
                        new.append(self.data[dispset[idx]])
                        done.add(idx)
                    elif idx not in done:
                        new.append(self.data[idx])
                        idx += 1
                    else:
                        idx += 1
                self.data = new
                self.RI.fix_index(max_len - 1)
                if isinstance(self._row_index, list) and self._row_index:
                    new = []
                    idx = 0
                    done = set()
                    while len(new) < len(self._row_index):
                        if idx in dispset and idx not in done:
                            new.append(self._row_index[dispset[idx]])
                            done.add(idx)
                        elif idx not in done:
                            new.append(self._row_index[idx])
                            idx += 1
                        else:
                            idx += 1
                    self._row_index = new
                dispset = {b: a for a, b in dispset.items()}
                self.RI.cell_options = {dispset[k] if k in dispset else k: v for k, v in self.RI.cell_options.items()}
                self.cell_options = {
                    (dispset[k[0]], k[1]) if k[0] in dispset else k: v for k, v in self.cell_options.items()
                }
                self.row_options = {dispset[k] if k in dispset else k: v for k, v in self.row_options.items()}
        return new_selected, {b: a for a, b in dispset.items()}

    def ctrl_z(self, event=None):
        if not self.undo_storage:
            return
        if not isinstance(self.undo_storage[-1], (tuple, dict)):
            undo_storage = pickle.loads(zlib.decompress(self.undo_storage[-1]))
        else:
            undo_storage = self.undo_storage[-1]
        self.deselect("all")
        if self.extra_begin_ctrl_z_func is not None:
            try:
                self.extra_begin_ctrl_z_func(UndoEvent("begin_ctrl_z", undo_storage[0], undo_storage))
            except Exception:
                return
        self.undo_storage.pop()
        if undo_storage[0] in ("edit_header",):
            for c, v in undo_storage[1].items():
                self._headers[c] = v
            self.reselect_from_get_boxes(undo_storage[2])
            self.set_currently_selected(0, undo_storage[3][1], type_="column")

        if undo_storage[0] in ("edit_index",):
            for r, v in undo_storage[1].items():
                self._row_index[r] = v
            self.reselect_from_get_boxes(undo_storage[2])
            self.set_currently_selected(0, undo_storage[3][1], type_="row")

        if undo_storage[0] in ("edit_cells", "edit_cells_paste"):
            for (datarn, datacn), v in undo_storage[1].items():
                self.set_cell_data(datarn, datacn, v)
            start_row = float("inf")
            start_col = float("inf")
            if undo_storage[0] == "edit_cells_paste" and self.expand_sheet_if_paste_too_big:
                if undo_storage[4][0] > 0:
                    self.del_row_positions(
                        len(self.row_positions) - 1 - undo_storage[4][0],
                        undo_storage[4][0],
                    )
                    self.data[:] = self.data[: -undo_storage[4][0]]
                    if not self.all_rows_displayed:
                        self.displayed_rows[:] = self.displayed_rows[: -undo_storage[4][0]]
                if undo_storage[4][1] > 0:
                    quick_added_cols = undo_storage[4][1]
                    self.del_col_positions(len(self.col_positions) - 1 - quick_added_cols, quick_added_cols)
                    for rn in range(len(self.data)):
                        self.data[rn][:] = self.data[rn][:-quick_added_cols]
                    if not self.all_columns_displayed:
                        self.displayed_columns[:] = self.displayed_columns[:-quick_added_cols]
            self.reselect_from_get_boxes(undo_storage[2])
            if undo_storage[3]:
                self.set_currently_selected(
                    undo_storage[3].row,
                    undo_storage[3].column,
                    type_=undo_storage[3].type_,
                )
            elif start_row < len(self.row_positions) - 1 and start_col < len(self.col_positions) - 1:
                self.set_currently_selected(start_row, start_col, type_="cell")
            if start_row < len(self.row_positions) - 1 and start_col < len(self.col_positions) - 1:
                self.see(
                    r=start_row,
                    c=start_col,
                    keep_yscroll=False,
                    keep_xscroll=False,
                    bottom_right_corner=False,
                    check_cell_visibility=True,
                    redraw=False,
                )

        elif undo_storage[0] == "move_cols":
            c = undo_storage[1][0]
            to_move_min = undo_storage[2][0]
            totalcols = len(undo_storage[2])
            if to_move_min < c:
                c += totalcols - 1
            self.move_columns_adjust_options_dict(c, to_move_min, totalcols)
            self.see(
                r=0,
                c=c,
                keep_yscroll=False,
                keep_xscroll=False,
                bottom_right_corner=False,
                check_cell_visibility=True,
                redraw=False,
            )

        elif undo_storage[0] == "move_rows":
            r = undo_storage[1][0]
            to_move_min = undo_storage[2][0]
            totalrows = len(undo_storage[2])
            if to_move_min < r:
                r += totalrows - 1
            self.move_rows_adjust_options_dict(r, to_move_min, totalrows)
            self.see(
                r=r,
                c=0,
                keep_yscroll=False,
                keep_xscroll=False,
                bottom_right_corner=False,
                check_cell_visibility=True,
                redraw=False,
            )

        elif undo_storage[0] == "insert_rows":
            self.displayed_rows = undo_storage[1]["displayed_rows"]
            self.data[
                undo_storage[1]["data_row_num"] : undo_storage[1]["data_row_num"] + undo_storage[1]["numrows"]
            ] = []
            try:
                self._row_index[
                    undo_storage[1]["data_row_num"] : undo_storage[1]["data_row_num"] + undo_storage[1]["numrows"]
                ] = []
            except Exception:
                pass
            self.del_row_positions(
                undo_storage[1]["sheet_row_num"],
                undo_storage[1]["numrows"],
                deselect_all=False,
            )
            to_del = set(
                range(
                    undo_storage[1]["sheet_row_num"],
                    undo_storage[1]["sheet_row_num"] + undo_storage[1]["numrows"],
                )
            )
            numrows = undo_storage[1]["numrows"]
            idx = undo_storage[1]["sheet_row_num"] + undo_storage[1]["numrows"]
            self.cell_options = {
                (rn if rn < idx else rn - numrows, cn): t2
                for (rn, cn), t2 in self.cell_options.items()
                if rn not in to_del
            }
            self.row_options = {
                rn if rn < idx else rn - numrows: t for rn, t in self.row_options.items() if rn not in to_del
            }
            self.RI.cell_options = {
                rn if rn < idx else rn - numrows: t for rn, t in self.RI.cell_options.items() if rn not in to_del
            }
            if len(self.row_positions) > 1:
                start_row = (
                    undo_storage[1]["sheet_row_num"]
                    if undo_storage[1]["sheet_row_num"] < len(self.row_positions) - 1
                    else undo_storage[1]["sheet_row_num"] - 1
                )
                self.RI.select_row(start_row)
                self.see(
                    r=start_row,
                    c=0,
                    keep_yscroll=False,
                    keep_xscroll=False,
                    bottom_right_corner=False,
                    check_cell_visibility=True,
                    redraw=False,
                )

        elif undo_storage[0] == "insert_cols":
            self.displayed_columns = undo_storage[1]["displayed_columns"]
            qx = undo_storage[1]["data_col_num"]
            qnum = undo_storage[1]["numcols"]
            for rn in range(len(self.data)):
                self.data[rn][qx : qx + qnum] = []
            try:
                self._headers[qx : qx + qnum] = []
            except Exception:
                pass
            self.del_col_positions(
                undo_storage[1]["sheet_col_num"],
                undo_storage[1]["numcols"],
                deselect_all=False,
            )
            to_del = set(
                range(
                    undo_storage[1]["sheet_col_num"],
                    undo_storage[1]["sheet_col_num"] + undo_storage[1]["numcols"],
                )
            )
            numcols = undo_storage[1]["numcols"]
            idx = undo_storage[1]["sheet_col_num"] + undo_storage[1]["numcols"]
            self.cell_options = {
                (rn, cn if cn < idx else cn - numcols): t2
                for (rn, cn), t2 in self.cell_options.items()
                if cn not in to_del
            }
            self.col_options = {
                cn if cn < idx else cn - numcols: t for cn, t in self.col_options.items() if cn not in to_del
            }
            self.CH.cell_options = {
                cn if cn < idx else cn - numcols: t for cn, t in self.CH.cell_options.items() if cn not in to_del
            }
            if len(self.col_positions) > 1:
                start_col = (
                    undo_storage[1]["sheet_col_num"]
                    if undo_storage[1]["sheet_col_num"] < len(self.col_positions) - 1
                    else undo_storage[1]["sheet_col_num"] - 1
                )
                self.CH.select_col(start_col)
                self.see(
                    r=0,
                    c=start_col,
                    keep_yscroll=False,
                    keep_xscroll=False,
                    bottom_right_corner=False,
                    check_cell_visibility=True,
                    redraw=False,
                )

        elif undo_storage[0] == "delete_rows":
            self.displayed_rows = undo_storage[1]["displayed_rows"]
            for rn, r in reversed(undo_storage[1]["deleted_rows"]):
                self.data.insert(rn, r)
            for rn, h in reversed(tuple(undo_storage[1]["rowheights"].items())):
                self.insert_row_position(idx=rn, height=h)
            self.cell_options = undo_storage[1]["cell_options"]
            self.row_options = undo_storage[1]["row_options"]
            self.RI.cell_options = undo_storage[1]["RI_cell_options"]
            for rn, r in reversed(undo_storage[1]["deleted_index_values"]):
                try:
                    self._row_index.insert(rn, r)
                except Exception:
                    continue
            self.reselect_from_get_boxes(undo_storage[1]["selection_boxes"])

        elif undo_storage[0] == "delete_cols":
            self.displayed_columns = undo_storage[1]["displayed_columns"]
            self.cell_options = undo_storage[1]["cell_options"]
            self.col_options = undo_storage[1]["col_options"]
            self.CH.cell_options = undo_storage[1]["CH_cell_options"]
            for cn, w in reversed(tuple(undo_storage[1]["colwidths"].items())):
                self.insert_col_position(idx=cn, width=w)
            for cn, rowdict in reversed(tuple(undo_storage[1]["deleted_cols"].items())):
                for rn, v in rowdict.items():
                    try:
                        self.data[rn].insert(cn, v)
                    except Exception:
                        continue
            for cn, v in reversed(tuple(undo_storage[1]["deleted_header_values"].items())):
                try:
                    self._headers.insert(cn, v)
                except Exception:
                    continue
            self.reselect_from_get_boxes(undo_storage[1]["selection_boxes"])
        self.refresh()
        if self.extra_end_ctrl_z_func is not None:
            self.extra_end_ctrl_z_func(UndoEvent("end_ctrl_z", undo_storage[0], undo_storage))
        self.parentframe.emit_event("<<SheetModified>>")

    def bind_arrowkeys(self, keys: dict = {}):
        for canvas in (self, self.parentframe, self.CH, self.RI, self.TL):
            for k, func in keys.items():
                canvas.bind(f"<{arrowkey_bindings_helper[k.lower()]}>", func)

    def unbind_arrowkeys(self, keys: dict = {}):
        for canvas in (self, self.parentframe, self.CH, self.RI, self.TL):
            for k, func in keys.items():
                canvas.unbind(f"<{arrowkey_bindings_helper[k.lower()]}>")

    def see(
        self,
        r=None,
        c=None,
        keep_yscroll=False,
        keep_xscroll=False,
        bottom_right_corner=False,
        check_cell_visibility=True,
        redraw=True,
    ):
        need_redraw = False
        if check_cell_visibility:
            yvis, xvis = self.cell_completely_visible(r=r, c=c, separate_axes=True)
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
                    args = (
                        "moveto",
                        y / (self.row_positions[-1] + self.empty_vertical),
                    )
                    if args[1] > 1:
                        args[1] = args[1] - 1
                    self.yview(*args)
                    self.RI.yview(*args)
                    need_redraw = True
            else:
                if r is not None and not keep_yscroll:
                    args = (
                        "moveto",
                        self.row_positions[r] / (self.row_positions[-1] + self.empty_vertical),
                    )
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
                    args = (
                        "moveto",
                        x / (self.col_positions[-1] + self.empty_horizontal),
                    )
                    self.xview(*args)
                    self.CH.xview(*args)
                    need_redraw = True
            else:
                if c is not None and not keep_xscroll:
                    args = (
                        "moveto",
                        self.col_positions[c] / (self.col_positions[-1] + self.empty_horizontal),
                    )
                    self.xview(*args)
                    self.CH.xview(*args)
                    need_redraw = True
        if redraw and need_redraw:
            self.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
            return True
        else:
            return False

    def get_cell_coords(self, r=None, c=None):
        return (
            0 if not c else self.col_positions[c] + 1,
            0 if not r else self.row_positions[r] + 1,
            0 if not c else self.col_positions[c + 1],
            0 if not r else self.row_positions[r + 1],
        )

    def cell_completely_visible(self, r=0, c=0, separate_axes=False):
        cx1, cy1, cx2, cy2 = self.get_canvas_visible_area()
        x1, y1, x2, y2 = self.get_cell_coords(r, c)
        x_vis = True
        y_vis = True
        if cx1 > x1 or cx2 < x2:
            x_vis = False
        if cy1 > y1 or cy2 < y2:
            y_vis = False
        if separate_axes:
            return y_vis, x_vis
        else:
            if not y_vis or not x_vis:
                return False
            else:
                return True

    def cell_visible(self, r=0, c=0):
        cx1, cy1, cx2, cy2 = self.get_canvas_visible_area()
        x1, y1, x2, y2 = self.get_cell_coords(r, c)
        if x1 <= cx2 or y1 <= cy2 or x2 >= cx1 or y2 >= cy1:
            return True
        return False

    def select_all(self, redraw=True, run_binding_func=True):
        currently_selected = self.currently_selected()
        self.deselect("all")
        if len(self.row_positions) > 1 and len(self.col_positions) > 1:
            if currently_selected:
                self.set_currently_selected(currently_selected.row, currently_selected.column, type_="cell")
            else:
                self.set_currently_selected(0, 0, type_="cell")
            self.create_selected(0, 0, len(self.row_positions) - 1, len(self.col_positions) - 1)
            if redraw:
                self.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
            if self.select_all_binding_func is not None and run_binding_func:
                self.select_all_binding_func(
                    SelectionBoxEvent(
                        "select_all_cells",
                        (
                            0,
                            0,
                            len(self.row_positions) - 1,
                            len(self.col_positions) - 1,
                        ),
                    )
                )

    def select_cell(self, r, c, redraw=False):
        self.delete_selection_rects()
        self.create_selected(r, c, r + 1, c + 1, state="hidden")
        self.set_currently_selected(r, c, type_="cell")
        if redraw:
            self.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
        if self.selection_binding_func is not None:
            self.selection_binding_func(SelectCellEvent("select_cell", r, c))

    def add_selection(self, r, c, redraw=False, run_binding_func=True, set_as_current=False):
        self.create_selected(r, c, r + 1, c + 1, state="hidden")
        if set_as_current:
            self.set_currently_selected(r, c, type_="cell")
        if redraw:
            self.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
        if self.selection_binding_func is not None and run_binding_func:
            self.selection_binding_func(SelectCellEvent("select_cell", r, c))

    def toggle_select_cell(
        self,
        row,
        column,
        add_selection=True,
        redraw=True,
        run_binding_func=True,
        set_as_current=True,
    ):
        if add_selection:
            if self.cell_selected(row, column, inc_rows=True, inc_cols=True):
                self.deselect(r=row, c=column, redraw=redraw)
            else:
                self.add_selection(
                    r=row,
                    c=column,
                    redraw=redraw,
                    run_binding_func=run_binding_func,
                    set_as_current=set_as_current,
                )
        else:
            if self.cell_selected(row, column, inc_rows=True, inc_cols=True):
                self.deselect(r=row, c=column, redraw=redraw)
            else:
                self.select_cell(row, column, redraw=redraw)

    def align_rows(self, rows=[], align="global", align_index=False):  # "center", "w", "e" or "global"
        if isinstance(rows, str) and rows.lower() == "all" and align == "global":
            for r in self.row_options:
                if "align" in self.row_options[r]:
                    del self.row_options[r]["align"]
            if align_index:
                for r in self.RI.cell_options:
                    if r in self.RI.cell_options and "align" in self.RI.cell_options[r]:
                        del self.RI.cell_options[r]["align"]
            return
        if isinstance(rows, int):
            rows_ = [rows]
        elif isinstance(rows, str) and rows.lower() == "all":
            rows_ = (r for r in range(self.total_data_rows()))
        else:
            rows_ = rows
        if align == "global":
            for r in rows_:
                if r in self.row_options and "align" in self.row_options[r]:
                    del self.row_options[r]["align"]
                if align_index and r in self.RI.cell_options and "align" in self.RI.cell_options[r]:
                    del self.RI.cell_options[r]["align"]
        else:
            for r in rows_:
                if r not in self.row_options:
                    self.row_options[r] = {}
                self.row_options[r]["align"] = align
                if align_index:
                    if r not in self.RI.cell_options:
                        self.RI.cell_options[r] = {}
                    self.RI.cell_options[r]["align"] = align

    def align_columns(self, columns=[], align="global", align_header=False):  # "center", "w", "e" or "global"
        if isinstance(columns, str) and columns.lower() == "all" and align == "global":
            for c in self.col_options:
                if "align" in self.col_options[c]:
                    del self.col_options[c]["align"]
            if align_header:
                for c in self.CH.cell_options:
                    if c in self.CH.cell_options and "align" in self.CH.cell_options[c]:
                        del self.CH.cell_options[c]["align"]
            return
        if isinstance(columns, int):
            cols_ = [columns]
        elif isinstance(columns, str) and columns.lower() == "all":
            cols_ = (c for c in range(self.total_data_cols()))
        else:
            cols_ = columns
        if align == "global":
            for c in cols_:
                if c in self.col_options and "align" in self.col_options[c]:
                    del self.col_options[c]["align"]
                if align_header and c in self.CH.cell_options and "align" in self.CH.cell_options[c]:
                    del self.CH.cell_options[c]["align"]
        else:
            for c in cols_:
                if c not in self.col_options:
                    self.col_options[c] = {}
                self.col_options[c]["align"] = align
                if align_header:
                    if c not in self.CH.cell_options:
                        self.CH.cell_options[c] = {}
                    self.CH.cell_options[c]["align"] = align

    def align_cells(self, row=0, column=0, cells=[], align="global"):  # "center", "w", "e" or "global"
        if isinstance(row, str) and row.lower() == "all" and align == "global":
            for r, c in self.cell_options:
                if "align" in self.cell_options[(r, c)]:
                    del self.cell_options[(r, c)]["align"]
            return
        if align == "global":
            if cells:
                for r, c in cells:
                    if (r, c) in self.cell_options and "align" in self.cell_options[(r, c)]:
                        del self.cell_options[(r, c)]["align"]
            else:
                if (row, column) in self.cell_options and "align" in self.cell_options[(row, column)]:
                    del self.cell_options[(row, column)]["align"]
        else:
            if cells:
                for r, c in cells:
                    if (r, c) not in self.cell_options:
                        self.cell_options[(r, c)] = {}
                    self.cell_options[(r, c)]["align"] = align
            else:
                if (row, column) not in self.cell_options:
                    self.cell_options[(row, column)] = {}
                self.cell_options[(row, column)]["align"] = align

    def deselect(self, r=None, c=None, cell=None, redraw=True):
        deselected = tuple()
        deleted_boxes = {}
        if r == "all":
            deselected = ("deselect_all", self.delete_selection_rects())
        elif r == "allrows":
            for item in self.find_withtag("rows"):
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
            for item in self.find_withtag("columns"):
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
            current = self.find_withtag("selected")
            current_tags = self.gettags(current[0]) if current else tuple()
            if current:
                curr_r1, curr_c1, curr_r2, curr_c2 = tuple(int(e) for e in current_tags[1].split("_") if e)
            reset_current = False
            for item in self.find_withtag("rows"):
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
            current = self.find_withtag("selected")
            current_tags = self.gettags(current[0]) if current else tuple()
            if current:
                curr_r1, curr_c1, curr_r2, curr_c2 = tuple(int(e) for e in current_tags[1].split("_") if e)
            reset_current = False
            for item in self.find_withtag("columns"):
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
            for item in chain(
                self.find_withtag("cells"),
                self.find_withtag("rows"),
                self.find_withtag("columns"),
                self.find_withtag("selected"),
            ):
                alltags = self.gettags(item)
                if alltags:
                    r1, c1, r2, c2 = tuple(int(e) for e in alltags[1].split("_") if e)
                    if r >= r1 and c >= c1 and r < r2 and c < c2:
                        current = self.currently_selected()
                        if (
                            not set_curr
                            and current
                            and r2 - r1 == 1
                            and c2 - c1 == 1
                            and r == current[0]
                            and c == current[1]
                        ):
                            set_curr = True
                        if current and not set_curr:
                            if current[0] >= r1 and current[0] < r2 and current[1] >= c1 and current[1] < c2:
                                set_curr = True
                        self.delete(f"{r1}_{c1}_{r2}_{c2}")
                        self.RI.delete(f"{r1}_{c1}_{r2}_{c2}")
                        self.CH.delete(f"{r1}_{c1}_{r2}_{c2}")
                        deleted_boxes[(r1, c1, r2, c2)] = "cells"
            if set_curr:
                try:
                    deleted_boxes[tuple(int(e) for e in self.get_tags_of_current()[1].split("_") if e)] = "cells"
                except Exception:
                    pass
                self.delete_current()
                self.set_current_to_last()
            deselected = ("deselect_cell", deleted_boxes)
        if redraw:
            self.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
        if self.deselection_binding_func is not None:
            self.deselection_binding_func(DeselectionEvent(*deselected))

    def page_UP(self, event=None):
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
            if self.RI.row_selection_enabled and (
                self.anything_selected(exclude_columns=True, exclude_cells=True) or not self.anything_selected()
            ):
                self.RI.select_row(r)
                self.see(r, 0, keep_xscroll=True, check_cell_visibility=False)
            elif (self.single_selection_enabled or self.toggle_selection_enabled) and self.anything_selected(
                exclude_columns=True, exclude_rows=True
            ):
                box = self.get_all_selection_boxes_with_types()[0][0]
                self.see(r, box[1], keep_xscroll=True, check_cell_visibility=False)
                self.select_cell(r, box[1])
        else:
            args = ("moveto", scrollto / (self.row_positions[-1] + 100))
            self.yview(*args)
            self.RI.yview(*args)
            self.main_table_redraw_grid_and_text(redraw_row_index=True)

    def page_DOWN(self, event=None):
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
            if self.RI.row_selection_enabled and (
                self.anything_selected(exclude_columns=True, exclude_cells=True) or not self.anything_selected()
            ):
                self.RI.select_row(r)
                self.see(r, 0, keep_xscroll=True, check_cell_visibility=False)
            elif (self.single_selection_enabled or self.toggle_selection_enabled) and self.anything_selected(
                exclude_columns=True, exclude_rows=True
            ):
                box = self.get_all_selection_boxes_with_types()[0][0]
                self.see(r, box[1], keep_xscroll=True, check_cell_visibility=False)
                self.select_cell(r, box[1])
        else:
            end = self.row_positions[-1]
            if scrollto > end + 100:
                scrollto = end
            args = ("moveto", scrollto / (end + 100))
            self.yview(*args)
            self.RI.yview(*args)
            self.main_table_redraw_grid_and_text(redraw_row_index=True)

    def arrowkey_UP(self, event=None):
        currently_selected = self.currently_selected()
        if not currently_selected:
            return
        if currently_selected.type_ == "row":
            r = currently_selected.row
            if r != 0 and self.RI.row_selection_enabled:
                if self.cell_completely_visible(r=r - 1, c=0):
                    self.RI.select_row(r - 1, redraw=True)
                else:
                    self.RI.select_row(r - 1)
                    self.see(r - 1, 0, keep_xscroll=True, check_cell_visibility=False)
        elif currently_selected.type_ in ("cell", "column"):
            r = currently_selected[0]
            c = currently_selected[1]
            if r == 0 and self.CH.col_selection_enabled:
                if not self.cell_completely_visible(r=r, c=0):
                    self.see(r, c, keep_xscroll=True, check_cell_visibility=False)
            elif r != 0 and (self.single_selection_enabled or self.toggle_selection_enabled):
                if self.cell_completely_visible(r=r - 1, c=c):
                    self.select_cell(r - 1, c, redraw=True)
                else:
                    self.select_cell(r - 1, c)
                    self.see(r - 1, c, keep_xscroll=True, check_cell_visibility=False)

    def arrowkey_RIGHT(self, event=None):
        currently_selected = self.currently_selected()
        if not currently_selected:
            return
        if currently_selected.type_ == "row":
            r = currently_selected.row
            if self.single_selection_enabled or self.toggle_selection_enabled:
                if self.cell_completely_visible(r=r, c=0):
                    self.select_cell(r, 0, redraw=True)
                else:
                    self.select_cell(r, 0)
                    self.see(
                        r,
                        0,
                        keep_yscroll=True,
                        bottom_right_corner=True,
                        check_cell_visibility=False,
                    )
        elif currently_selected.type_ == "column":
            c = currently_selected.column
            if c < len(self.col_positions) - 2 and self.CH.col_selection_enabled:
                if self.cell_completely_visible(r=0, c=c + 1):
                    self.CH.select_col(c + 1, redraw=True)
                else:
                    self.CH.select_col(c + 1)
                    self.see(
                        0,
                        c + 1,
                        keep_yscroll=True,
                        bottom_right_corner=False if self.arrow_key_down_right_scroll_page else True,
                        check_cell_visibility=False,
                    )
        else:
            r = currently_selected[0]
            c = currently_selected[1]
            if c < len(self.col_positions) - 2 and (self.single_selection_enabled or self.toggle_selection_enabled):
                if self.cell_completely_visible(r=r, c=c + 1):
                    self.select_cell(r, c + 1, redraw=True)
                else:
                    self.select_cell(r, c + 1)
                    self.see(
                        r,
                        c + 1,
                        keep_yscroll=True,
                        bottom_right_corner=False if self.arrow_key_down_right_scroll_page else True,
                        check_cell_visibility=False,
                    )

    def arrowkey_DOWN(self, event=None):
        currently_selected = self.currently_selected()
        if not currently_selected:
            return
        if currently_selected.type_ == "row":
            r = currently_selected.row
            if r < len(self.row_positions) - 2 and self.RI.row_selection_enabled:
                if self.cell_completely_visible(r=min(r + 2, len(self.row_positions) - 2), c=0):
                    self.RI.select_row(r + 1, redraw=True)
                else:
                    self.RI.select_row(r + 1)
                    if (
                        r + 2 < len(self.row_positions) - 2
                        and (self.row_positions[r + 3] - self.row_positions[r + 2])
                        + (self.row_positions[r + 2] - self.row_positions[r + 1])
                        + 5
                        < self.winfo_height()
                    ):
                        self.see(
                            r + 2,
                            0,
                            keep_xscroll=True,
                            bottom_right_corner=True,
                            check_cell_visibility=False,
                        )
                    elif not self.cell_completely_visible(r=r + 1, c=0):
                        self.see(
                            r + 1,
                            0,
                            keep_xscroll=True,
                            bottom_right_corner=False if self.arrow_key_down_right_scroll_page else True,
                            check_cell_visibility=False,
                        )
        elif currently_selected.type_ == "column":
            c = currently_selected.column
            if self.single_selection_enabled or self.toggle_selection_enabled:
                if self.cell_completely_visible(r=0, c=c):
                    self.select_cell(0, c, redraw=True)
                else:
                    self.select_cell(0, c)
                    self.see(
                        0,
                        c,
                        keep_xscroll=True,
                        bottom_right_corner=True,
                        check_cell_visibility=False,
                    )
        else:
            r = currently_selected[0]
            c = currently_selected[1]
            if r < len(self.row_positions) - 2 and (self.single_selection_enabled or self.toggle_selection_enabled):
                if self.cell_completely_visible(r=min(r + 2, len(self.row_positions) - 2), c=c):
                    self.select_cell(r + 1, c, redraw=True)
                else:
                    self.select_cell(r + 1, c)
                    if (
                        r + 2 < len(self.row_positions) - 2
                        and (self.row_positions[r + 3] - self.row_positions[r + 2])
                        + (self.row_positions[r + 2] - self.row_positions[r + 1])
                        + 5
                        < self.winfo_height()
                    ):
                        self.see(
                            r + 2,
                            c,
                            keep_xscroll=True,
                            bottom_right_corner=True,
                            check_cell_visibility=False,
                        )
                    elif not self.cell_completely_visible(r=r + 1, c=c):
                        self.see(
                            r + 1,
                            c,
                            keep_xscroll=True,
                            bottom_right_corner=False if self.arrow_key_down_right_scroll_page else True,
                            check_cell_visibility=False,
                        )

    def arrowkey_LEFT(self, event=None):
        currently_selected = self.currently_selected()
        if not currently_selected:
            return
        if currently_selected.type_ == "column":
            c = currently_selected.column
            if c != 0 and self.CH.col_selection_enabled:
                if self.cell_completely_visible(r=0, c=c - 1):
                    self.CH.select_col(c - 1, redraw=True)
                else:
                    self.CH.select_col(c - 1)
                    self.see(
                        0,
                        c - 1,
                        keep_yscroll=True,
                        bottom_right_corner=True,
                        check_cell_visibility=False,
                    )
        elif currently_selected.type_ == "cell":
            r = currently_selected.row
            c = currently_selected.column
            if c == 0 and self.RI.row_selection_enabled:
                if not self.cell_completely_visible(r=r, c=0):
                    self.see(r, c, keep_yscroll=True, check_cell_visibility=False)
            elif c != 0 and (self.single_selection_enabled or self.toggle_selection_enabled):
                if self.cell_completely_visible(r=r, c=c - 1):
                    self.select_cell(r, c - 1, redraw=True)
                else:
                    self.select_cell(r, c - 1)
                    self.see(r, c - 1, keep_yscroll=True, check_cell_visibility=False)

    def edit_bindings(self, enable=True, key=None):
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
        if key is None or key == "edit_header":
            if enable:
                self.CH.bind_cell_edit(True)
            else:
                self.CH.bind_cell_edit(False)
        if key is None or key == "edit_index":
            if enable:
                self.RI.bind_cell_edit(True)
            else:
                self.RI.bind_cell_edit(False)

    def menu_add_command(self, menu: tk.Menu, **kwargs):
        if "label" not in kwargs:
            return
        try:
            index = menu.index(kwargs["label"])
            menu.delete(index)
        except TclError:
            pass
        menu.add_command(**kwargs)

    def create_rc_menus(self):
        if not self.rc_popup_menu:
            self.rc_popup_menu = tk.Menu(self, tearoff=0, background=self.popup_menu_bg)
        if not self.CH.ch_rc_popup_menu:
            self.CH.ch_rc_popup_menu = tk.Menu(self.CH, tearoff=0, background=self.popup_menu_bg)
        if not self.RI.ri_rc_popup_menu:
            self.RI.ri_rc_popup_menu = tk.Menu(self.RI, tearoff=0, background=self.popup_menu_bg)
        if not self.empty_rc_popup_menu:
            self.empty_rc_popup_menu = tk.Menu(self, tearoff=0, background=self.popup_menu_bg)
        for menu in (
            self.rc_popup_menu,
            self.CH.ch_rc_popup_menu,
            self.RI.ri_rc_popup_menu,
            self.empty_rc_popup_menu,
        ):
            menu.delete(0, "end")
        if self.rc_popup_menus_enabled and self.CH.edit_cell_enabled:
            self.menu_add_command(
                self.CH.ch_rc_popup_menu,
                label="Edit header",
                font=self.popup_menu_font,
                foreground=self.popup_menu_fg,
                background=self.popup_menu_bg,
                activebackground=self.popup_menu_highlight_bg,
                activeforeground=self.popup_menu_highlight_fg,
                command=lambda: self.CH.open_cell(event="rc"),
            )
        if self.rc_popup_menus_enabled and self.RI.edit_cell_enabled:
            self.menu_add_command(
                self.RI.ri_rc_popup_menu,
                label="Edit index",
                font=self.popup_menu_font,
                foreground=self.popup_menu_fg,
                background=self.popup_menu_bg,
                activebackground=self.popup_menu_highlight_bg,
                activeforeground=self.popup_menu_highlight_fg,
                command=lambda: self.RI.open_cell(event="rc"),
            )
        if self.rc_popup_menus_enabled and self.edit_cell_enabled:
            self.menu_add_command(
                self.rc_popup_menu,
                label="Edit cell",
                font=self.popup_menu_font,
                foreground=self.popup_menu_fg,
                background=self.popup_menu_bg,
                activebackground=self.popup_menu_highlight_bg,
                activeforeground=self.popup_menu_highlight_fg,
                command=lambda: self.open_cell(event="rc"),
            )
        if self.cut_enabled:
            self.menu_add_command(
                self.rc_popup_menu,
                label="Cut",
                accelerator="Ctrl+X",
                font=self.popup_menu_font,
                foreground=self.popup_menu_fg,
                background=self.popup_menu_bg,
                activebackground=self.popup_menu_highlight_bg,
                activeforeground=self.popup_menu_highlight_fg,
                command=self.ctrl_x,
            )
            self.menu_add_command(
                self.CH.ch_rc_popup_menu,
                label="Cut contents",
                accelerator="Ctrl+X",
                font=self.popup_menu_font,
                foreground=self.popup_menu_fg,
                background=self.popup_menu_bg,
                activebackground=self.popup_menu_highlight_bg,
                activeforeground=self.popup_menu_highlight_fg,
                command=self.ctrl_x,
            )
            self.menu_add_command(
                self.RI.ri_rc_popup_menu,
                label="Cut contents",
                accelerator="Ctrl+X",
                font=self.popup_menu_font,
                foreground=self.popup_menu_fg,
                background=self.popup_menu_bg,
                activebackground=self.popup_menu_highlight_bg,
                activeforeground=self.popup_menu_highlight_fg,
                command=self.ctrl_x,
            )
        if self.copy_enabled:
            self.menu_add_command(
                self.rc_popup_menu,
                label="Copy",
                accelerator="Ctrl+C",
                font=self.popup_menu_font,
                foreground=self.popup_menu_fg,
                background=self.popup_menu_bg,
                activebackground=self.popup_menu_highlight_bg,
                activeforeground=self.popup_menu_highlight_fg,
                command=self.ctrl_c,
            )
            self.menu_add_command(
                self.CH.ch_rc_popup_menu,
                label="Copy contents",
                accelerator="Ctrl+C",
                font=self.popup_menu_font,
                foreground=self.popup_menu_fg,
                background=self.popup_menu_bg,
                activebackground=self.popup_menu_highlight_bg,
                activeforeground=self.popup_menu_highlight_fg,
                command=self.ctrl_c,
            )
            self.menu_add_command(
                self.RI.ri_rc_popup_menu,
                label="Copy contents",
                accelerator="Ctrl+C",
                font=self.popup_menu_font,
                foreground=self.popup_menu_fg,
                background=self.popup_menu_bg,
                activebackground=self.popup_menu_highlight_bg,
                activeforeground=self.popup_menu_highlight_fg,
                command=self.ctrl_c,
            )
        if self.paste_enabled:
            self.menu_add_command(
                self.rc_popup_menu,
                label="Paste",
                accelerator="Ctrl+V",
                font=self.popup_menu_font,
                foreground=self.popup_menu_fg,
                background=self.popup_menu_bg,
                activebackground=self.popup_menu_highlight_bg,
                activeforeground=self.popup_menu_highlight_fg,
                command=self.ctrl_v,
            )
            self.menu_add_command(
                self.CH.ch_rc_popup_menu,
                label="Paste",
                accelerator="Ctrl+V",
                font=self.popup_menu_font,
                foreground=self.popup_menu_fg,
                background=self.popup_menu_bg,
                activebackground=self.popup_menu_highlight_bg,
                activeforeground=self.popup_menu_highlight_fg,
                command=self.ctrl_v,
            )
            self.menu_add_command(
                self.RI.ri_rc_popup_menu,
                label="Paste",
                accelerator="Ctrl+V",
                font=self.popup_menu_font,
                foreground=self.popup_menu_fg,
                background=self.popup_menu_bg,
                activebackground=self.popup_menu_highlight_bg,
                activeforeground=self.popup_menu_highlight_fg,
                command=self.ctrl_v,
            )
            if self.expand_sheet_if_paste_too_big:
                self.menu_add_command(
                    self.empty_rc_popup_menu,
                    label="Paste",
                    accelerator="Ctrl+V",
                    font=self.popup_menu_font,
                    foreground=self.popup_menu_fg,
                    background=self.popup_menu_bg,
                    activebackground=self.popup_menu_highlight_bg,
                    activeforeground=self.popup_menu_highlight_fg,
                    command=self.ctrl_v,
                )
        if self.delete_key_enabled:
            self.menu_add_command(
                self.rc_popup_menu,
                label="Delete",
                accelerator="Del",
                font=self.popup_menu_font,
                foreground=self.popup_menu_fg,
                background=self.popup_menu_bg,
                activebackground=self.popup_menu_highlight_bg,
                activeforeground=self.popup_menu_highlight_fg,
                command=self.delete_key,
            )
            self.menu_add_command(
                self.CH.ch_rc_popup_menu,
                label="Clear contents",
                accelerator="Del",
                font=self.popup_menu_font,
                foreground=self.popup_menu_fg,
                background=self.popup_menu_bg,
                activebackground=self.popup_menu_highlight_bg,
                activeforeground=self.popup_menu_highlight_fg,
                command=self.delete_key,
            )
            self.menu_add_command(
                self.RI.ri_rc_popup_menu,
                label="Clear contents",
                accelerator="Del",
                font=self.popup_menu_font,
                foreground=self.popup_menu_fg,
                background=self.popup_menu_bg,
                activebackground=self.popup_menu_highlight_bg,
                activeforeground=self.popup_menu_highlight_fg,
                command=self.delete_key,
            )
        if self.rc_delete_column_enabled:
            self.menu_add_command(
                self.CH.ch_rc_popup_menu,
                label="Delete columns",
                font=self.popup_menu_font,
                foreground=self.popup_menu_fg,
                background=self.popup_menu_bg,
                activebackground=self.popup_menu_highlight_bg,
                activeforeground=self.popup_menu_highlight_fg,
                command=self.del_cols_rc,
            )
        if self.rc_insert_column_enabled:
            self.menu_add_command(
                self.CH.ch_rc_popup_menu,
                label="Insert columns left",
                font=self.popup_menu_font,
                foreground=self.popup_menu_fg,
                background=self.popup_menu_bg,
                activebackground=self.popup_menu_highlight_bg,
                activeforeground=self.popup_menu_highlight_fg,
                command=lambda: self.insert_cols_rc("left"),
            )
            self.menu_add_command(
                self.empty_rc_popup_menu,
                label="Insert column",
                font=self.popup_menu_font,
                foreground=self.popup_menu_fg,
                background=self.popup_menu_bg,
                activebackground=self.popup_menu_highlight_bg,
                activeforeground=self.popup_menu_highlight_fg,
                command=lambda: self.insert_cols_rc("left"),
            )
            self.menu_add_command(
                self.CH.ch_rc_popup_menu,
                label="Insert columns right",
                font=self.popup_menu_font,
                foreground=self.popup_menu_fg,
                background=self.popup_menu_bg,
                activebackground=self.popup_menu_highlight_bg,
                activeforeground=self.popup_menu_highlight_fg,
                command=lambda: self.insert_cols_rc("right"),
            )
        if self.rc_delete_row_enabled:
            self.menu_add_command(
                self.RI.ri_rc_popup_menu,
                label="Delete rows",
                font=self.popup_menu_font,
                foreground=self.popup_menu_fg,
                background=self.popup_menu_bg,
                activebackground=self.popup_menu_highlight_bg,
                activeforeground=self.popup_menu_highlight_fg,
                command=self.del_rows_rc,
            )
        if self.rc_insert_row_enabled:
            self.menu_add_command(
                self.RI.ri_rc_popup_menu,
                label="Insert rows above",
                font=self.popup_menu_font,
                foreground=self.popup_menu_fg,
                background=self.popup_menu_bg,
                activebackground=self.popup_menu_highlight_bg,
                activeforeground=self.popup_menu_highlight_fg,
                command=lambda: self.insert_rows_rc("above"),
            )
            self.menu_add_command(
                self.RI.ri_rc_popup_menu,
                label="Insert rows below",
                font=self.popup_menu_font,
                foreground=self.popup_menu_fg,
                background=self.popup_menu_bg,
                activebackground=self.popup_menu_highlight_bg,
                activeforeground=self.popup_menu_highlight_fg,
                command=lambda: self.insert_rows_rc("below"),
            )
            self.menu_add_command(
                self.empty_rc_popup_menu,
                label="Insert row",
                font=self.popup_menu_font,
                foreground=self.popup_menu_fg,
                background=self.popup_menu_bg,
                activebackground=self.popup_menu_highlight_bg,
                activeforeground=self.popup_menu_highlight_fg,
                command=lambda: self.insert_rows_rc("below"),
            )
        for label, func in self.extra_table_rc_menu_funcs.items():
            self.menu_add_command(
                self.rc_popup_menu,
                label=label,
                font=self.popup_menu_font,
                foreground=self.popup_menu_fg,
                background=self.popup_menu_bg,
                activebackground=self.popup_menu_highlight_bg,
                activeforeground=self.popup_menu_highlight_fg,
                command=func,
            )
        for label, func in self.extra_index_rc_menu_funcs.items():
            self.menu_add_command(
                self.RI.ri_rc_popup_menu,
                label=label,
                font=self.popup_menu_font,
                foreground=self.popup_menu_fg,
                background=self.popup_menu_bg,
                activebackground=self.popup_menu_highlight_bg,
                activeforeground=self.popup_menu_highlight_fg,
                command=func,
            )
        for label, func in self.extra_header_rc_menu_funcs.items():
            self.menu_add_command(
                self.CH.ch_rc_popup_menu,
                label=label,
                font=self.popup_menu_font,
                foreground=self.popup_menu_fg,
                background=self.popup_menu_bg,
                activebackground=self.popup_menu_highlight_bg,
                activeforeground=self.popup_menu_highlight_fg,
                command=func,
            )
        for label, func in self.extra_empty_space_rc_menu_funcs.items():
            self.menu_add_command(
                self.empty_rc_popup_menu,
                label=label,
                font=self.popup_menu_font,
                foreground=self.popup_menu_fg,
                background=self.popup_menu_bg,
                activebackground=self.popup_menu_highlight_bg,
                activeforeground=self.popup_menu_highlight_fg,
                command=func,
            )

    def bind_cell_edit(self, enable=True, keys=[]):
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

    def enable_disable_select_all(self, enable=True):
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
            self.bind_arrowkeys(self.arrowkey_binding_functions)
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
            self.bind_arrowkeys(self.arrowkey_binding_functions)
        elif binding in ("tab", "up", "right", "down", "left", "prior", "next"):
            self.bind_arrowkeys(keys={binding: self.arrowkey_binding_functions[binding]})
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
        elif binding in ("ctrl_click_select", "ctrl_select"):
            self.ctrl_select_enabled = True
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
            self.unbind_arrowkeys(self.arrowkey_binding_functions)
            self.edit_bindings(False)
            self.rc_delete_column_enabled = False
            self.rc_delete_row_enabled = False
            self.rc_insert_column_enabled = False
            self.rc_insert_row_enabled = False
            self.rc_popup_menus_enabled = False
            self.rc_select_enabled = False
            self.ctrl_select_enabled = False
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
            self.unbind_arrowkeys(self.arrowkey_binding_functions)
        elif binding in ("tab", "up", "right", "down", "left", "prior", "next"):
            self.unbind_arrowkeys(keys={binding: self.arrowkey_binding_functions[binding]})
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
        elif binding in ("ctrl_click_select", "ctrl_select"):
            self.ctrl_select_enabled = False
        elif binding == "undo":
            self.edit_bindings(False, "undo")
        elif binding == "edit_cell":
            self.edit_bindings(False, "edit_cell")
        elif binding == "edit_header":
            self.edit_bindings(False, "edit_header")
        elif binding == "edit_index":
            self.edit_bindings(False, "edit_index")
        self.create_rc_menus()

    def reset_mouse_motion_creations(self, event=None):
        self.config(cursor="")
        self.RI.config(cursor="")
        self.CH.config(cursor="")
        self.RI.rsz_w = None
        self.RI.rsz_h = None
        self.CH.rsz_w = None
        self.CH.rsz_h = None

    def mouse_motion(self, event):
        if (
            not self.RI.currently_resizing_height
            and not self.RI.currently_resizing_width
            and not self.CH.currently_resizing_height
            and not self.CH.currently_resizing_width
        ):
            mouse_over_resize = False
            x = self.canvasx(event.x)
            y = self.canvasy(event.y)
            if self.RI.width_resizing_enabled and self.show_index and not mouse_over_resize:
                try:
                    x1, y1, x2, y2 = (
                        self.row_width_resize_bbox[0],
                        self.row_width_resize_bbox[1],
                        self.row_width_resize_bbox[2],
                        self.row_width_resize_bbox[3],
                    )
                    if x >= x1 and y >= y1 and x <= x2 and y <= y2:
                        self.config(cursor="sb_h_double_arrow")
                        self.RI.config(cursor="sb_h_double_arrow")
                        self.RI.rsz_w = True
                        mouse_over_resize = True
                except Exception:
                    pass
            if self.CH.height_resizing_enabled and self.show_header and not mouse_over_resize:
                try:
                    x1, y1, x2, y2 = (
                        self.header_height_resize_bbox[0],
                        self.header_height_resize_bbox[1],
                        self.header_height_resize_bbox[2],
                        self.header_height_resize_bbox[3],
                    )
                    if x >= x1 and y >= y1 and x <= x2 and y <= y2:
                        self.config(cursor="sb_v_double_arrow")
                        self.CH.config(cursor="sb_v_double_arrow")
                        self.CH.rsz_h = True
                        mouse_over_resize = True
                except Exception:
                    pass
            if not mouse_over_resize:
                self.reset_mouse_motion_creations()
        if self.extra_motion_func is not None:
            self.extra_motion_func(event)

    def rc(self, event=None):
        self.mouseclick_outside_editor_or_dropdown_all_canvases()
        self.focus_set()
        popup_menu = None
        if self.single_selection_enabled and all(
            v is None for v in (self.RI.rsz_h, self.RI.rsz_w, self.CH.rsz_h, self.CH.rsz_w)
        ):
            r = self.identify_row(y=event.y)
            c = self.identify_col(x=event.x)
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
                        self.select_cell(r, c, redraw=True)
                    if self.rc_popup_menus_enabled:
                        popup_menu = self.rc_popup_menu
            else:
                popup_menu = self.empty_rc_popup_menu
        elif self.toggle_selection_enabled and all(
            v is None for v in (self.RI.rsz_h, self.RI.rsz_w, self.CH.rsz_h, self.CH.rsz_w)
        ):
            r = self.identify_row(y=event.y)
            c = self.identify_col(x=event.x)
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
                        self.toggle_select_cell(r, c, redraw=True)
                    if self.rc_popup_menus_enabled:
                        popup_menu = self.rc_popup_menu
            else:
                popup_menu = self.empty_rc_popup_menu
        if self.extra_rc_func is not None:
            self.extra_rc_func(event)
        if popup_menu is not None:
            popup_menu.tk_popup(event.x_root, event.y_root)

    def b1_press(self, event=None):
        self.closed_dropdown = self.mouseclick_outside_editor_or_dropdown_all_canvases()
        self.focus_set()
        x1, y1, x2, y2 = self.get_canvas_visible_area()
        if (
            self.identify_col(x=event.x, allow_end=False) is None
            or self.identify_row(y=event.y, allow_end=False) is None
        ):
            self.deselect("all")
        r = self.identify_row(y=event.y)
        c = self.identify_col(x=event.x)
        if self.single_selection_enabled and all(
            v is None for v in (self.RI.rsz_h, self.RI.rsz_w, self.CH.rsz_h, self.CH.rsz_w)
        ):
            if r < len(self.row_positions) - 1 and c < len(self.col_positions) - 1:
                self.select_cell(r, c, redraw=True)
        elif self.toggle_selection_enabled and all(
            v is None for v in (self.RI.rsz_h, self.RI.rsz_w, self.CH.rsz_h, self.CH.rsz_w)
        ):
            r = self.identify_row(y=event.y)
            c = self.identify_col(x=event.x)
            if r < len(self.row_positions) - 1 and c < len(self.col_positions) - 1:
                self.toggle_select_cell(r, c, redraw=True)
        elif self.RI.width_resizing_enabled and self.show_index and self.RI.rsz_h is None and self.RI.rsz_w:
            self.RI.currently_resizing_width = True
            self.new_row_width = self.RI.current_width + event.x
            x = self.canvasx(event.x)
            self.create_resize_line(x, y1, x, y2, width=1, fill=self.RI.resizing_line_fg, tag="rwl")
        elif self.CH.height_resizing_enabled and self.show_header and self.CH.rsz_w is None and self.CH.rsz_h:
            self.CH.currently_resizing_height = True
            self.new_header_height = self.CH.current_height + event.y
            y = self.canvasy(event.y)
            self.create_resize_line(x1, y, x2, y, width=1, fill=self.RI.resizing_line_fg, tag="rhl")
        self.b1_pressed_loc = (r, c)
        if self.extra_b1_press_func is not None:
            self.extra_b1_press_func(event)

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

    def ctrl_b1_press(self, event=None):
        self.mouseclick_outside_editor_or_dropdown_all_canvases()
        self.focus_set()
        if self.ctrl_select_enabled and all(
            v is None for v in (self.RI.rsz_h, self.RI.rsz_w, self.CH.rsz_h, self.CH.rsz_w)
        ):
            self.b1_pressed_loc = None
            rowsel = int(self.identify_row(y=event.y))
            colsel = int(self.identify_col(x=event.x))
            if rowsel < len(self.row_positions) - 1 and colsel < len(self.col_positions) - 1:
                currently_selected = self.currently_selected()
                if not currently_selected or currently_selected.row != rowsel or currently_selected.column != colsel:
                    self.add_selection(rowsel, colsel, set_as_current=True)
                self.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True, redraw_table=True)
                if self.ctrl_selection_binding_func is not None:
                    self.ctrl_selection_binding_func(
                        SelectionBoxEvent(
                            "ctrl_select_cells",
                            (rowsel, colsel, rowsel + 1, colsel + 1),
                        )
                    )
        elif not self.ctrl_select_enabled:
            self.b1_press(event)

    def ctrl_shift_b1_press(self, event=None):
        self.mouseclick_outside_editor_or_dropdown_all_canvases()
        self.focus_set()
        if (
            self.ctrl_select_enabled
            and self.drag_selection_enabled
            and all(v is None for v in (self.RI.rsz_h, self.RI.rsz_w, self.CH.rsz_h, self.CH.rsz_w))
        ):
            self.b1_pressed_loc = None
            rowsel = int(self.identify_row(y=event.y))
            colsel = int(self.identify_col(x=event.x))
            if rowsel < len(self.row_positions) - 1 and colsel < len(self.col_positions) - 1:
                currently_selected = self.currently_selected()
                if currently_selected and currently_selected.type_ == "cell":
                    min_r = currently_selected[0]
                    min_c = currently_selected[1]
                    if rowsel >= min_r and colsel >= min_c:
                        last_selected = (min_r, min_c, rowsel + 1, colsel + 1)
                    elif rowsel >= min_r and min_c >= colsel:
                        last_selected = (min_r, colsel, rowsel + 1, min_c + 1)
                    elif min_r >= rowsel and colsel >= min_c:
                        last_selected = (rowsel, min_c, min_r + 1, colsel + 1)
                    elif min_r >= rowsel and min_c >= colsel:
                        last_selected = (rowsel, colsel, min_r + 1, min_c + 1)
                    self.create_selected(*last_selected)
                else:
                    self.add_selection(rowsel, colsel, set_as_current=True)
                    last_selected = (rowsel, colsel, rowsel + 1, colsel + 1)
                self.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True, redraw_table=True)
                if self.shift_selection_binding_func is not None:
                    self.shift_selection_binding_func(SelectionBoxEvent("shift_select_cells", last_selected))
        elif not self.ctrl_select_enabled:
            self.shift_b1_press(event)

    def shift_b1_press(self, event=None):
        self.mouseclick_outside_editor_or_dropdown_all_canvases()
        self.focus_set()
        if self.drag_selection_enabled and all(
            v is None for v in (self.RI.rsz_h, self.RI.rsz_w, self.CH.rsz_h, self.CH.rsz_w)
        ):
            self.b1_pressed_loc = None
            rowsel = int(self.identify_row(y=event.y))
            colsel = int(self.identify_col(x=event.x))
            if rowsel < len(self.row_positions) - 1 and colsel < len(self.col_positions) - 1:
                currently_selected = self.currently_selected()
                if currently_selected and currently_selected.type_ == "cell":
                    min_r = currently_selected[0]
                    min_c = currently_selected[1]
                    self.delete_selection_rects(delete_current=False)
                    if rowsel >= min_r and colsel >= min_c:
                        self.create_selected(min_r, min_c, rowsel + 1, colsel + 1)
                    elif rowsel >= min_r and min_c >= colsel:
                        self.create_selected(min_r, colsel, rowsel + 1, min_c + 1)
                    elif min_r >= rowsel and colsel >= min_c:
                        self.create_selected(rowsel, min_c, min_r + 1, colsel + 1)
                    elif min_r >= rowsel and min_c >= colsel:
                        self.create_selected(rowsel, colsel, min_r + 1, min_c + 1)
                    last_selected = tuple(int(e) for e in self.gettags(self.find_withtag("cells"))[1].split("_") if e)
                else:
                    self.select_cell(rowsel, colsel, redraw=False)
                    last_selected = tuple(
                        int(e) for e in self.gettags(self.find_withtag("selected"))[1].split("_") if e
                    )
                self.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True, redraw_table=True)
                if self.shift_selection_binding_func is not None:
                    self.shift_selection_binding_func(SelectionBoxEvent("shift_select_cells", last_selected))

    def get_b1_motion_rect(self, start_row, start_col, end_row, end_col):
        if end_row >= start_row and end_col >= start_col and (end_row - start_row or end_col - start_col):
            return (start_row, start_col, end_row + 1, end_col + 1, "cells")
        elif end_row >= start_row and end_col < start_col and (end_row - start_row or start_col - end_col):
            return (start_row, end_col, end_row + 1, start_col + 1, "cells")
        elif end_row < start_row and end_col >= start_col and (start_row - end_row or end_col - start_col):
            return (end_row, start_col, start_row + 1, end_col + 1, "cells")
        elif end_row < start_row and end_col < start_col and (start_row - end_row or start_col - end_col):
            return (end_row, end_col, start_row + 1, start_col + 1, "cells")
        else:
            return (start_row, start_col, start_row + 1, start_col + 1, "cells")
        return None

    def b1_motion(self, event):
        x1, y1, x2, y2 = self.get_canvas_visible_area()
        if self.drag_selection_enabled and all(
            v is None for v in (self.RI.rsz_h, self.RI.rsz_w, self.CH.rsz_h, self.CH.rsz_w)
        ):
            need_redraw = False
            end_row = self.identify_row(y=event.y)
            end_col = self.identify_col(x=event.x)
            currently_selected = self.currently_selected()
            if (
                end_row < len(self.row_positions) - 1
                and end_col < len(self.col_positions) - 1
                and currently_selected
                and currently_selected.type_ == "cell"
            ):
                rect = self.get_b1_motion_rect(*(currently_selected[0], currently_selected[1], end_row, end_col))
                if rect is not None and self.being_drawn_rect != rect:
                    self.delete_selection_rects(delete_current=False)
                    if rect[2] - rect[0] == 1 and rect[3] - rect[1] == 1:
                        self.being_drawn_rect = rect
                    else:
                        self.being_drawn_rect = rect
                        self.create_selected(*rect)
                    if self.drag_selection_binding_func is not None:
                        self.drag_selection_binding_func(SelectionBoxEvent("drag_select_cells", rect[:-1]))
                    need_redraw = True
            if self.scroll_if_event_offscreen(event):
                need_redraw = True
            if need_redraw:
                self.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True, redraw_table=True)
        elif self.RI.width_resizing_enabled and self.RI.rsz_w is not None and self.RI.currently_resizing_width:
            self.RI.delete_resize_lines()
            self.delete_resize_lines()
            if event.x >= 0:
                x = self.canvasx(event.x)
                self.new_row_width = self.RI.current_width + event.x
                self.create_resize_line(x, y1, x, y2, width=1, fill=self.RI.resizing_line_fg, tag="rwl")
            else:
                x = self.RI.current_width + event.x
                if x < self.min_column_width:
                    x = int(self.min_column_width)
                self.new_row_width = x
                self.RI.create_resize_line(x, y1, x, y2, width=1, fill=self.RI.resizing_line_fg, tag="rwl")
        elif self.CH.height_resizing_enabled and self.CH.rsz_h is not None and self.CH.currently_resizing_height:
            self.CH.delete_resize_lines()
            self.delete_resize_lines()
            if event.y >= 0:
                y = self.canvasy(event.y)
                self.new_header_height = self.CH.current_height + event.y
                self.create_resize_line(x1, y, x2, y, width=1, fill=self.RI.resizing_line_fg, tag="rhl")
            else:
                y = self.CH.current_height + event.y
                if y < self.min_header_height:
                    y = int(self.min_header_height)
                self.new_header_height = y
                self.CH.create_resize_line(x1, y, x2, y, width=1, fill=self.RI.resizing_line_fg, tag="rhl")
        if self.extra_b1_motion_func is not None:
            self.extra_b1_motion_func(event)

    def ctrl_b1_motion(self, event):
        x1, y1, x2, y2 = self.get_canvas_visible_area()
        if (
            self.ctrl_select_enabled
            and self.drag_selection_enabled
            and all(v is None for v in (self.RI.rsz_h, self.RI.rsz_w, self.CH.rsz_h, self.CH.rsz_w))
        ):
            need_redraw = False
            end_row = self.identify_row(y=event.y)
            end_col = self.identify_col(x=event.x)
            currently_selected = self.currently_selected()
            if (
                end_row < len(self.row_positions) - 1
                and end_col < len(self.col_positions) - 1
                and currently_selected
                and currently_selected.type_ == "cell"
            ):
                rect = self.get_b1_motion_rect(*(currently_selected[0], currently_selected[1], end_row, end_col))
                if rect is not None and self.being_drawn_rect != rect:
                    if rect[2] - rect[0] == 1 and rect[3] - rect[1] == 1:
                        self.being_drawn_rect = rect
                    else:
                        if self.being_drawn_rect is not None:
                            self.delete_selected(*self.being_drawn_rect)
                        self.being_drawn_rect = rect
                        self.create_selected(*rect)
                    if self.drag_selection_binding_func is not None:
                        self.drag_selection_binding_func(SelectionBoxEvent("drag_select_cells", rect[:-1]))
                    need_redraw = True
            if self.scroll_if_event_offscreen(event):
                need_redraw = True
            if need_redraw:
                self.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True, redraw_table=True)
        elif not self.ctrl_select_enabled:
            self.b1_motion(event)

    def b1_release(self, event=None):
        if self.being_drawn_rect is not None and (
            self.being_drawn_rect[2] - self.being_drawn_rect[0] > 1
            or self.being_drawn_rect[3] - self.being_drawn_rect[1] > 1
        ):
            self.delete_selected(*self.being_drawn_rect)
            to_sel = tuple(self.being_drawn_rect)
            self.being_drawn_rect = None
            self.create_selected(*to_sel)
        else:
            self.being_drawn_rect = None
        if self.RI.width_resizing_enabled and self.RI.rsz_w is not None and self.RI.currently_resizing_width:
            self.delete_resize_lines()
            self.RI.delete_resize_lines()
            self.RI.currently_resizing_width = False
            self.RI.set_width(self.new_row_width, set_TL=True)
            self.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
            self.b1_pressed_loc = None
        elif self.CH.height_resizing_enabled and self.CH.rsz_h is not None and self.CH.currently_resizing_height:
            self.delete_resize_lines()
            self.CH.delete_resize_lines()
            self.CH.currently_resizing_height = False
            self.CH.set_height(self.new_header_height, set_TL=True)
            self.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
            self.b1_pressed_loc = None
        self.RI.rsz_w = None
        self.CH.rsz_h = None
        if self.b1_pressed_loc is not None:
            r = self.identify_row(y=event.y, allow_end=False)
            c = self.identify_col(x=event.x, allow_end=False)
            if r is not None and c is not None and (r, c) == self.b1_pressed_loc:
                datarn = r if self.all_rows_displayed else self.displayed_rows[r]
                datacn = c if self.all_columns_displayed else self.displayed_columns[c]
                if self.get_cell_kwargs(datarn, datacn, key="dropdown") or self.get_cell_kwargs(
                    datarn, datacn, key="checkbox"
                ):
                    canvasx = self.canvasx(event.x)
                    if (
                        self.closed_dropdown != self.b1_pressed_loc
                        and self.get_cell_kwargs(datarn, datacn, key="dropdown")
                        and canvasx > self.col_positions[c + 1] - self.txt_h - 4
                        and canvasx < self.col_positions[c + 1] - 1
                    ) or (
                        self.get_cell_kwargs(datarn, datacn, key="checkbox")
                        and canvasx < self.col_positions[c] + self.txt_h + 4
                        and self.canvasy(event.y) < self.row_positions[r] + self.txt_h + 4
                    ):
                        self.open_cell(event)
            else:
                self.mouseclick_outside_editor_or_dropdown_all_canvases()
        self.b1_pressed_loc = None
        self.closed_dropdown = None
        if self.extra_b1_release_func is not None:
            self.extra_b1_release_func(event)

    def double_b1(self, event=None):
        self.mouseclick_outside_editor_or_dropdown_all_canvases()
        self.focus_set()
        x1, y1, x2, y2 = self.get_canvas_visible_area()
        if (
            self.identify_col(x=event.x, allow_end=False) is None
            or self.identify_row(y=event.y, allow_end=False) is None
        ):
            self.deselect("all")
        elif self.single_selection_enabled and all(
            v is None for v in (self.RI.rsz_h, self.RI.rsz_w, self.CH.rsz_h, self.CH.rsz_w)
        ):
            r = self.identify_row(y=event.y)
            c = self.identify_col(x=event.x)
            if r < len(self.row_positions) - 1 and c < len(self.col_positions) - 1:
                self.select_cell(r, c, redraw=True)
                if self.edit_cell_enabled:
                    self.open_cell(event)
        elif self.toggle_selection_enabled and all(
            v is None for v in (self.RI.rsz_h, self.RI.rsz_w, self.CH.rsz_h, self.CH.rsz_w)
        ):
            r = self.identify_row(y=event.y)
            c = self.identify_col(x=event.x)
            if r < len(self.row_positions) - 1 and c < len(self.col_positions) - 1:
                self.toggle_select_cell(r, c, redraw=True)
                if self.edit_cell_enabled:
                    self.open_cell(event)
        if self.extra_double_b1_func is not None:
            self.extra_double_b1_func(event)

    def identify_row(self, event=None, y=None, allow_end=True):
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

    def identify_col(self, event=None, x=None, allow_end=True):
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

    def fix_views(self):
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

    def scroll_if_event_offscreen(self, event):
        need_redraw = False
        if self.data:
            xcheck = self.xview()
            ycheck = self.yview()
            if len(xcheck) > 1 and xcheck[0] > 0 and event.x < 0:
                try:
                    self.xview_scroll(-1, "units")
                    self.CH.xview_scroll(-1, "units")
                except Exception:
                    pass
                need_redraw = True
            if len(ycheck) > 1 and ycheck[0] > 0 and event.y < 0:
                try:
                    self.yview_scroll(-1, "units")
                    self.RI.yview_scroll(-1, "units")
                except Exception:
                    pass
                need_redraw = True
            if len(xcheck) > 1 and xcheck[1] < 1 and event.x > self.winfo_width():
                try:
                    self.xview_scroll(1, "units")
                    self.CH.xview_scroll(1, "units")
                except Exception:
                    pass
                need_redraw = True
            if len(ycheck) > 1 and ycheck[1] < 1 and event.y > self.winfo_height():
                try:
                    self.yview_scroll(1, "units")
                    self.RI.yview_scroll(1, "units")
                except Exception:
                    pass
                need_redraw = True
        if need_redraw:
            self.fix_views()
        return need_redraw

    def set_xviews(self, *args):
        self.xview(*args)
        if self.show_header:
            self.CH.xview(*args)
        self.fix_views()
        self.main_table_redraw_grid_and_text(redraw_header=True if self.show_header else False)

    def set_yviews(self, *args):
        self.yview(*args)
        if self.show_index:
            self.RI.yview(*args)
        self.fix_views()
        self.main_table_redraw_grid_and_text(redraw_row_index=True if self.show_index else False)

    def set_view(self, x_args, y_args):
        self.xview(*x_args)
        if self.show_header:
            self.CH.xview(*x_args)
        self.yview(*y_args)
        if self.show_index:
            self.RI.yview(*y_args)
        self.fix_views()
        self.main_table_redraw_grid_and_text(
            redraw_row_index=True if self.show_index else False,
            redraw_header=True if self.show_header else False,
        )

    def mousewheel(self, event=None):
        if event.delta < 0 or event.num == 5:
            self.yview_scroll(1, "units")
            self.RI.yview_scroll(1, "units")
        elif event.delta >= 0 or event.num == 4:
            if self.canvasy(0) <= 0:
                return
            self.yview_scroll(-1, "units")
            self.RI.yview_scroll(-1, "units")
        self.main_table_redraw_grid_and_text(redraw_row_index=True)

    def shift_mousewheel(self, event=None):
        if event.delta < 0 or event.num == 5:
            self.xview_scroll(1, "units")
            self.CH.xview_scroll(1, "units")
        elif event.delta >= 0 or event.num == 4:
            if self.canvasx(0) <= 0:
                return
            self.xview_scroll(-1, "units")
            self.CH.xview_scroll(-1, "units")
        self.main_table_redraw_grid_and_text(redraw_header=True)

    def get_txt_w(self, txt, font=None):
        self.txt_measure_canvas.itemconfig(
            self.txt_measure_canvas_text,
            text=txt,
            font=self.table_font if font is None else font,
        )
        b = self.txt_measure_canvas.bbox(self.txt_measure_canvas_text)
        return b[2] - b[0]

    def get_txt_h(self, txt, font=None):
        self.txt_measure_canvas.itemconfig(
            self.txt_measure_canvas_text,
            text=txt,
            font=self.table_font if font is None else font,
        )
        b = self.txt_measure_canvas.bbox(self.txt_measure_canvas_text)
        return b[3] - b[1]

    def get_txt_dimensions(self, txt, font=None):
        self.txt_measure_canvas.itemconfig(
            self.txt_measure_canvas_text,
            text=txt,
            font=self.table_font if font is None else font,
        )
        b = self.txt_measure_canvas.bbox(self.txt_measure_canvas_text)
        return b[2] - b[0], b[3] - b[1]

    def get_lines_cell_height(self, n, font=None):
        return (
            self.get_txt_h(
                txt="\n".join(["j^|" for lines in range(n)]) if n > 1 else "j^|",
                font=self.table_font if font is None else font,
            )
            + 5
        )

    def set_min_column_width(self):
        self.min_column_width = 5
        if self.min_column_width > self.max_column_width:
            self.max_column_width = self.min_column_width + 20
        if self.min_column_width > self.default_column_width:
            self.default_column_width = self.min_column_width + 20

    def set_table_font(self, newfont=None, reset_row_positions=False):
        if newfont:
            if not isinstance(newfont, tuple):
                raise ValueError("Argument must be tuple e.g. ('Carlito',12,'normal')")
            if len(newfont) != 3:
                raise ValueError("Argument must be three-tuple")
            if not isinstance(newfont[0], str) or not isinstance(newfont[1], int) or not isinstance(newfont[2], str):
                raise ValueError(
                    "Argument must be font, size and 'normal', 'bold' or 'italic' e.g. ('Carlito',12,'normal')"
                )
            self.table_font = newfont
            self.font_fam = newfont[0]
            self.font_sze = newfont[1]
            self.font_wgt = newfont[2]
            self.set_table_font_help()
            if reset_row_positions:
                self.reset_row_positions()
            self.recreate_all_selection_boxes()
        else:
            return self.table_font

    def set_table_font_help(self):
        self.txt_h = self.get_txt_h("|ZXjy*'^")
        self.txt_w = self.get_txt_w("|")
        self.half_txt_h = ceil(self.txt_h / 2)
        if self.half_txt_h % 2 == 0:
            self.fl_ins = self.half_txt_h + 2
        else:
            self.fl_ins = self.half_txt_h + 3
        self.xtra_lines_increment = int(self.txt_h)
        self.min_row_height = self.txt_h + 5
        if self.min_row_height < 12:
            self.min_row_height = 12
        if self.default_row_height[0] != "pixels":
            self.default_row_height = (
                self.default_row_height[0] if self.default_row_height[0] != "pixels" else "pixels",
                self.get_lines_cell_height(int(self.default_row_height[0]))
                if self.default_row_height[0] != "pixels"
                else self.default_row_height[1],
            )
        self.set_min_column_width()

    def set_header_font(self, newfont=None):
        if newfont:
            if not isinstance(newfont, tuple):
                raise ValueError("Argument must be tuple e.g. ('Carlito', 12, 'normal')")
            if len(newfont) != 3:
                raise ValueError("Argument must be three-tuple")
            if not isinstance(newfont[0], str) or not isinstance(newfont[1], int) or not isinstance(newfont[2], str):
                raise ValueError(
                    "Argument must be font, size and 'normal', 'bold' or 'italic' e.g. ('Carlito', 12, 'normal')"
                )
            self.header_font = newfont
            self.header_font_fam = newfont[0]
            self.header_font_sze = newfont[1]
            self.header_font_wgt = newfont[2]
            self.set_header_font_help()
            self.recreate_all_selection_boxes()
        else:
            return self.header_font

    def set_header_font_help(self):
        self.header_txt_w, self.header_txt_h = self.get_txt_dimensions("|", self.header_font)
        self.header_half_txt_h = ceil(self.header_txt_h / 2)
        if self.header_half_txt_h % 2 == 0:
            self.header_fl_ins = self.header_half_txt_h + 2
        else:
            self.header_fl_ins = self.header_half_txt_h + 3
        self.header_xtra_lines_increment = self.header_txt_h
        self.min_header_height = self.header_txt_h + 5
        if self.default_header_height[0] != "pixels":
            self.default_header_height = (
                self.default_header_height[0] if self.default_header_height[0] != "pixels" else "pixels",
                self.get_lines_cell_height(int(self.default_header_height[0]), font=self.header_font)
                if self.default_header_height[0] != "pixels"
                else self.default_header_height[1],
            )
        self.set_min_column_width()
        self.CH.set_height(self.default_header_height[1], set_TL=True)

    def set_index_font(self, newfont=None):
        if newfont:
            if not isinstance(newfont, tuple):
                raise ValueError("Argument must be tuple e.g. ('Carlito', 12, 'normal')")
            if len(newfont) != 3:
                raise ValueError("Argument must be three-tuple")
            if not isinstance(newfont[0], str) or not isinstance(newfont[1], int) or not isinstance(newfont[2], str):
                raise ValueError(
                    "Argument must be font, size and 'normal', 'bold' or" "'italic' e.g. ('Carlito',12,'normal')"
                )
            self.index_font = newfont
            self.index_font_fam = newfont[0]
            self.index_font_sze = newfont[1]
            self.index_font_wgt = newfont[2]
            self.set_index_font_help()
        return self.index_font

    def set_index_font_help(self):
        self.index_txt_width, self.index_txt_height = self.get_txt_dimensions("|", self.index_font)
        self.index_half_txt_height = ceil(self.index_txt_height / 2)
        if self.index_half_txt_height % 2 == 0:
            self.index_first_ln_ins = self.index_half_txt_height + 2
        else:
            self.index_first_ln_ins = self.index_half_txt_height + 3
        self.index_xtra_lines_increment = self.index_txt_height
        self.min_index_width = 5

    def data_reference(
        self,
        newdataref=None,
        reset_col_positions=True,
        reset_row_positions=True,
        redraw=False,
        return_id=True,
        keep_formatting=True,
    ):
        if isinstance(newdataref, (list, tuple)):
            self.data = newdataref
            if keep_formatting:
                self.reapply_formatting()
            else:
                self.delete_all_formatting(clear_values=False)
            self.undo_storage = deque(maxlen=self.max_undos)
            if reset_col_positions:
                self.reset_col_positions()
            if reset_row_positions:
                self.reset_row_positions()
            if redraw:
                self.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
        if return_id:
            return id(self.data)
        else:
            return self.data

    def get_cell_dimensions(self, datarn, datacn):
        txt = self.get_valid_cell_data_as_str(datarn, datacn, get_displayed=True)
        if txt:
            self.txt_measure_canvas.itemconfig(self.txt_measure_canvas_text, text=txt, font=self.table_font)
            b = self.txt_measure_canvas.bbox(self.txt_measure_canvas_text)
            w = b[2] - b[0] + 7
            h = b[3] - b[1] + 5
        else:
            w = self.min_column_width
            h = self.min_row_height
        if self.get_cell_kwargs(datarn, datacn, key="dropdown") or self.get_cell_kwargs(datarn, datacn, key="checkbox"):
            return w + self.txt_h, h
        return w, h

    def set_cell_size_to_text(self, r, c, only_set_if_too_small=False, redraw=True, run_binding=False):
        min_column_width = int(self.min_column_width)
        min_rh = int(self.min_row_height)
        w = min_column_width
        h = min_rh
        datacn = c if self.all_columns_displayed else self.displayed_columns[c]
        datarn = r if self.all_rows_displayed else self.displayed_rows[r]
        tw, h = self.get_cell_dimensions(datarn, datacn)
        if tw > w:
            w = tw
        if h < min_rh:
            h = int(min_rh)
        elif h > self.max_row_height:
            h = int(self.max_row_height)
        if w < min_column_width:
            w = int(min_column_width)
        elif w > self.max_column_width:
            w = int(self.max_column_width)
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
            self.col_positions[c + 2 :] = [
                e + increment for e in islice(self.col_positions, c + 2, len(self.col_positions))
            ]
            self.col_positions[c + 1] = new_col_pos
            new_width = self.col_positions[c + 1] - self.col_positions[c]
            if run_binding and self.CH.column_width_resize_func is not None and old_width != new_width:
                self.CH.column_width_resize_func(ResizeEvent("column_width_resize", c, old_width, new_width))
        if cell_needs_resize_h:
            old_height = self.row_positions[r + 1] - self.row_positions[r]
            new_row_pos = self.row_positions[r] + h
            increment = new_row_pos - self.row_positions[r + 1]
            self.row_positions[r + 2 :] = [
                e + increment for e in islice(self.row_positions, r + 2, len(self.row_positions))
            ]
            self.row_positions[r + 1] = new_row_pos
            new_height = self.row_positions[r + 1] - self.row_positions[r]
            if run_binding and self.RI.row_height_resize_func is not None and old_height != new_height:
                self.RI.row_height_resize_func(ResizeEvent("row_height_resize", r, old_height, new_height))
        if cell_needs_resize_w or cell_needs_resize_h:
            self.recreate_all_selection_boxes()
            self.allow_auto_resize_columns = not cell_needs_resize_w
            self.allow_auto_resize_rows = not cell_needs_resize_h
            if redraw:
                self.refresh()
                return True
            else:
                return False

    def set_all_cell_sizes_to_text(self, include_index=False):
        min_column_width = int(self.min_column_width)
        min_rh = int(self.min_row_height)
        w = min_column_width
        h = min_rh
        rhs = defaultdict(lambda: int(min_rh))
        cws = []
        x = self.txt_measure_canvas.create_text(0, 0, text="", font=self.table_font)
        x2 = self.txt_measure_canvas.create_text(0, 0, text="", font=self.header_font)
        itmcon = self.txt_measure_canvas.itemconfig
        itmbbx = self.txt_measure_canvas.bbox
        if self.all_columns_displayed:
            itercols = range(self.total_data_cols())
        else:
            itercols = self.displayed_columns
        if self.all_rows_displayed:
            iterrows = range(self.total_data_rows())
        else:
            iterrows = self.displayed_rows
        if is_iterable(self._row_index):
            for datarn in iterrows:
                w_, h = self.RI.get_cell_dimensions(datarn)
                if h < min_rh:
                    h = int(min_rh)
                elif h > self.max_row_height:
                    h = int(self.max_row_height)
                if h > rhs[datarn]:
                    rhs[datarn] = h
        for datacn in itercols:
            w, h_ = self.CH.get_cell_dimensions(datacn)
            if self.all_rows_displayed:
                # refresh range generator if needed
                iterrows = range(self.total_data_rows())
            for datarn in iterrows:
                txt = self.get_valid_cell_data_as_str(datarn, datacn, get_displayed=True)
                if txt:
                    itmcon(x, text=txt)
                    b = itmbbx(x)
                    tw = b[2] - b[0] + 7
                    h = b[3] - b[1] + 5
                else:
                    tw = min_column_width
                    h = min_rh
                if self.get_cell_kwargs(datarn, datacn, key="dropdown") or self.get_cell_kwargs(
                    datarn, datacn, key="checkbox"
                ):
                    tw += self.txt_h
                if tw > w:
                    w = tw
                if h < min_rh:
                    h = int(min_rh)
                elif h > self.max_row_height:
                    h = int(self.max_row_height)
                if h > rhs[datarn]:
                    rhs[datarn] = h
            if w < min_column_width:
                w = int(min_column_width)
            elif w > self.max_column_width:
                w = int(self.max_column_width)
            cws.append(w)
        self.txt_measure_canvas.delete(x)
        self.txt_measure_canvas.delete(x2)
        self.row_positions = list(accumulate(chain([0], (height for height in rhs.values()))))
        self.col_positions = list(accumulate(chain([0], (width for width in cws))))
        self.recreate_all_selection_boxes()
        return self.row_positions, self.col_positions

    def reset_col_positions(self, ncols=None):
        colpos = int(self.default_column_width)
        if self.all_columns_displayed:
            self.col_positions = list(
                accumulate(
                    chain(
                        [0],
                        (colpos for c in range(ncols if ncols is not None else self.total_data_cols())),
                    )
                )
            )
        else:
            self.col_positions = list(
                accumulate(
                    chain(
                        [0],
                        (colpos for c in range(ncols if ncols is not None else len(self.displayed_columns))),
                    )
                )
            )

    def reset_row_positions(self, nrows=None):
        rowpos = self.default_row_height[1]
        if self.all_rows_displayed:
            self.row_positions = list(
                accumulate(
                    chain(
                        [0],
                        (rowpos for r in range(nrows if nrows is not None else self.total_data_rows())),
                    )
                )
            )
        else:
            self.row_positions = list(
                accumulate(
                    chain(
                        [0],
                        (rowpos for r in range(nrows if nrows is not None else len(self.displayed_rows))),
                    )
                )
            )

    def del_col_position(self, idx, deselect_all=False):
        if deselect_all:
            self.deselect("all", redraw=False)
        if idx == "end" or len(self.col_positions) <= idx + 1:
            del self.col_positions[-1]
        else:
            w = self.col_positions[idx + 1] - self.col_positions[idx]
            idx += 1
            del self.col_positions[idx]
            self.col_positions[idx:] = [e - w for e in islice(self.col_positions, idx, len(self.col_positions))]

    def del_row_position(self, idx, deselect_all=False):
        if deselect_all:
            self.deselect("all", redraw=False)
        if idx == "end" or len(self.row_positions) <= idx + 1:
            del self.row_positions[-1]
        else:
            w = self.row_positions[idx + 1] - self.row_positions[idx]
            idx += 1
            del self.row_positions[idx]
            self.row_positions[idx:] = [e - w for e in islice(self.row_positions, idx, len(self.row_positions))]

    def del_col_positions(self, idx, num=1, deselect_all=False):
        if deselect_all:
            self.deselect("all", redraw=False)
        if idx == "end" or len(self.col_positions) <= idx + 1:
            del self.col_positions[-1]
        else:
            cws = [
                int(b - a)
                for a, b in zip(
                    self.col_positions,
                    islice(self.col_positions, 1, len(self.col_positions)),
                )
            ]
            cws[idx : idx + num] = []
            self.col_positions = list(accumulate(chain([0], (width for width in cws))))

    def del_row_positions(self, idx, numrows=1, deselect_all=False):
        if deselect_all:
            self.deselect("all", redraw=False)
        if idx == "end" or len(self.row_positions) <= idx + 1:
            del self.row_positions[-1]
        else:
            rhs = [
                int(b - a)
                for a, b in zip(
                    self.row_positions,
                    islice(self.row_positions, 1, len(self.row_positions)),
                )
            ]
            rhs[idx : idx + numrows] = []
            self.row_positions = list(accumulate(chain([0], (height for height in rhs))))

    def insert_col_position(self, idx="end", width=None, deselect_all=False):
        if deselect_all:
            self.deselect("all", redraw=False)
        if width is None:
            w = self.default_column_width
        else:
            w = width
        if idx == "end" or len(self.col_positions) == idx + 1:
            self.col_positions.append(self.col_positions[-1] + w)
        else:
            idx += 1
            self.col_positions.insert(idx, self.col_positions[idx - 1] + w)
            idx += 1
            self.col_positions[idx:] = [e + w for e in islice(self.col_positions, idx, len(self.col_positions))]

    def insert_row_position(self, idx, height=None, deselect_all=False):
        if deselect_all:
            self.deselect("all", redraw=False)
        if height is None:
            h = self.default_row_height[1]
        else:
            h = height
        if idx == "end" or len(self.row_positions) == idx + 1:
            self.row_positions.append(self.row_positions[-1] + h)
        else:
            idx += 1
            self.row_positions.insert(idx, self.row_positions[idx - 1] + h)
            idx += 1
            self.row_positions[idx:] = [e + h for e in islice(self.row_positions, idx, len(self.row_positions))]

    def insert_col_positions(self, idx="end", widths=None, deselect_all=False):
        if deselect_all:
            self.deselect("all", redraw=False)
        if widths is None:
            w = [self.default_column_width]
        elif isinstance(widths, int):
            w = list(repeat(self.default_column_width, widths))
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
                self.col_positions[idx:idx] = list(
                    accumulate(chain([self.col_positions[idx - 1] + w[0]], islice(w, 1, None)))
                )
                idx += len(w)
                sumw = sum(w)
                self.col_positions[idx:] = [e + sumw for e in islice(self.col_positions, idx, len(self.col_positions))]
            else:
                w = w[0]
                idx += 1
                self.col_positions.insert(idx, self.col_positions[idx - 1] + w)
                idx += 1
                self.col_positions[idx:] = [e + w for e in islice(self.col_positions, idx, len(self.col_positions))]

    def insert_row_positions(self, idx="end", heights=None, deselect_all=False):
        if deselect_all:
            self.deselect("all", redraw=False)
        if heights is None:
            h = [self.default_row_height[1]]
        elif isinstance(heights, int):
            h = list(repeat(self.default_row_height[1], heights))
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
                self.row_positions[idx:idx] = list(
                    accumulate(chain([self.row_positions[idx - 1] + h[0]], islice(h, 1, None)))
                )
                idx += len(h)
                sumh = sum(h)
                self.row_positions[idx:] = [e + sumh for e in islice(self.row_positions, idx, len(self.row_positions))]
            else:
                h = h[0]
                idx += 1
                self.row_positions.insert(idx, self.row_positions[idx - 1] + h)
                idx += 1
                self.row_positions[idx:] = [e + h for e in islice(self.row_positions, idx, len(self.row_positions))]

    def insert_cols_rc(self, event=None):
        if self.anything_selected(exclude_rows=True, exclude_cells=True):
            selcols = self.get_selected_cols()
            numcols = len(selcols)
            displayed_ins_col = min(selcols) if event == "left" else max(selcols) + 1
            if self.all_columns_displayed:
                data_ins_col = int(displayed_ins_col)
            else:
                if displayed_ins_col == len(self.col_positions) - 1:
                    rowlen = len(max(self.data, key=len)) if self.data else 0
                    data_ins_col = rowlen
                else:
                    try:
                        data_ins_col = int(self.displayed_columns[displayed_ins_col])
                    except Exception:
                        data_ins_col = int(self.displayed_columns[displayed_ins_col - 1])
        else:
            numcols = 1
            displayed_ins_col = len(self.col_positions) - 1
            data_ins_col = int(displayed_ins_col)
        if (
            isinstance(self.paste_insert_column_limit, int)
            and self.paste_insert_column_limit < displayed_ins_col + numcols
        ):
            numcols = self.paste_insert_column_limit - len(self.col_positions) - 1
            if numcols < 1:
                return
        if self.extra_begin_insert_cols_rc_func is not None:
            try:
                self.extra_begin_insert_cols_rc_func(
                    InsertEvent("begin_insert_columns", data_ins_col, displayed_ins_col, numcols)
                )
            except Exception:
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
                part2 = list(
                    range(
                        self.displayed_columns[adj_ins],
                        self.displayed_columns[adj_ins] + numcols + 1,
                    )
                )
                part3 = (
                    []
                    if displayed_ins_col > len(self.displayed_columns) - 1
                    else [cn + numcols for cn in islice(self.displayed_columns, adj_ins + 1, None)]
                )
                self.displayed_columns = part1 + part2 + part3
        self.insert_col_positions(idx=displayed_ins_col, widths=numcols, deselect_all=True)
        self.cell_options = {
            (rn, cn if cn < data_ins_col else cn + numcols): t2 for (rn, cn), t2 in self.cell_options.items()
        }
        self.col_options = {cn if cn < data_ins_col else cn + numcols: t for cn, t in self.col_options.items()}
        self.CH.cell_options = {cn if cn < data_ins_col else cn + numcols: t for cn, t in self.CH.cell_options.items()}
        self.CH.fix_header()
        if self._headers and isinstance(self._headers, list):
            if data_ins_col >= len(self._headers):
                self.CH.fix_header(
                    datacn=data_ins_col,
                    fix_values=(data_ins_col, data_ins_col + numcols),
                )
            else:
                self._headers[data_ins_col:data_ins_col] = self.CH.get_empty_header_seq(
                    end=data_ins_col + numcols, start=data_ins_col, c_ops=False
                )
        if self.row_positions == [0] and not self.data:
            self.insert_row_position(idx="end", height=int(self.min_row_height), deselect_all=False)
            self.data.append(self.get_empty_row_seq(0, end=data_ins_col + numcols, start=data_ins_col, c_ops=False))
        else:
            end = data_ins_col + numcols
            for rn in range(len(self.data)):
                self.data[rn][data_ins_col:data_ins_col] = self.get_empty_row_seq(rn, end, data_ins_col, c_ops=False)
        self.create_selected(
            0,
            displayed_ins_col,
            len(self.row_positions) - 1,
            displayed_ins_col + numcols,
            "columns",
        )
        self.set_currently_selected(0, displayed_ins_col, "column")
        if self.undo_enabled:
            self.undo_storage.append(
                zlib.compress(
                    pickle.dumps(
                        (
                            "insert_cols",
                            {
                                "data_col_num": data_ins_col,
                                "displayed_columns": saved_displayed_columns,
                                "sheet_col_num": displayed_ins_col,
                                "numcols": numcols,
                            },
                        )
                    )
                )
            )
        self.refresh()
        if self.extra_end_insert_cols_rc_func is not None:
            self.extra_end_insert_cols_rc_func(
                InsertEvent("end_insert_columns", data_ins_col, displayed_ins_col, numcols)
            )
        self.parentframe.emit_event("<<SheetModified>>")

    def insert_rows_rc(self, event=None):
        if self.anything_selected(exclude_columns=True, exclude_cells=True):
            selrows = self.get_selected_rows()
            numrows = len(selrows)
            displayed_ins_row = min(selrows) if event == "above" else max(selrows) + 1
            if self.all_rows_displayed:
                data_ins_row = int(displayed_ins_row)
            else:
                if displayed_ins_row == len(self.row_positions) - 1:
                    datalen = len(self.data)
                    data_ins_row = datalen
                else:
                    try:
                        data_ins_row = int(self.displayed_rows[displayed_ins_row])
                    except Exception:
                        data_ins_row = int(self.displayed_rows[displayed_ins_row - 1])
        else:
            numrows = 1
            displayed_ins_row = len(self.row_positions) - 1
            data_ins_row = int(displayed_ins_row)
        if isinstance(self.paste_insert_row_limit, int) and self.paste_insert_row_limit < displayed_ins_row + numrows:
            numrows = self.paste_insert_row_limit - len(self.row_positions) - 1
            if numrows < 1:
                return
        if self.extra_begin_insert_rows_rc_func is not None:
            try:
                self.extra_begin_insert_rows_rc_func(
                    InsertEvent("begin_insert_rows", data_ins_row, displayed_ins_row, numrows)
                )
            except Exception:
                return
        saved_displayed_rows = list(self.displayed_rows)
        if not self.all_rows_displayed:
            if displayed_ins_row == len(self.row_positions) - 1:
                self.displayed_rows += list(range(datalen, datalen + numrows))
            else:
                if displayed_ins_row > len(self.displayed_rows) - 1:
                    adj_ins = displayed_ins_row - 1
                else:
                    adj_ins = displayed_ins_row
                part1 = self.displayed_rows[:adj_ins]
                part2 = list(
                    range(
                        self.displayed_rows[adj_ins],
                        self.displayed_rows[adj_ins] + numrows + 1,
                    )
                )
                part3 = (
                    []
                    if displayed_ins_row > len(self.displayed_rows) - 1
                    else [rn + numrows for rn in islice(self.displayed_rows, adj_ins + 1, None)]
                )
                self.displayed_rows = part1 + part2 + part3
        self.insert_row_positions(idx=displayed_ins_row, heights=numrows, deselect_all=True)
        self.cell_options = {
            (rn if rn < data_ins_row else rn + numrows, cn): t2 for (rn, cn), t2 in self.cell_options.items()
        }
        self.row_options = {rn if rn < data_ins_row else rn + numrows: t for rn, t in self.row_options.items()}
        self.RI.cell_options = {rn if rn < data_ins_row else rn + numrows: t for rn, t in self.RI.cell_options.items()}
        self.RI.fix_index()
        if self._row_index and isinstance(self._row_index, list):
            if data_ins_row >= len(self._row_index):
                self.RI.fix_index(
                    datacn=data_ins_row,
                    fix_values=(data_ins_row, data_ins_row + numrows),
                )
            else:
                self._row_index[data_ins_row:data_ins_row] = self.RI.get_empty_index_seq(
                    data_ins_row + numrows, data_ins_row, r_ops=False
                )
        if self.col_positions == [0] and not self.data:
            self.insert_col_position(idx="end", width=None, deselect_all=False)
            self.data.append(self.get_empty_row_seq(0, end=data_ins_row + numrows, start=data_ins_row, r_ops=False))
        else:
            total_data_cols = self.total_data_cols()
            self.data[data_ins_row:data_ins_row] = [
                self.get_empty_row_seq(rn, total_data_cols, r_ops=False)
                for rn in range(data_ins_row, data_ins_row + numrows)
            ]
        self.create_selected(
            displayed_ins_row,
            0,
            displayed_ins_row + numrows,
            len(self.col_positions) - 1,
            "rows",
        )
        self.set_currently_selected(displayed_ins_row, 0, "row")
        if self.undo_enabled:
            self.undo_storage.append(
                zlib.compress(
                    pickle.dumps(
                        (
                            "insert_rows",
                            {
                                "data_row_num": data_ins_row,
                                "displayed_rows": saved_displayed_rows,
                                "sheet_row_num": displayed_ins_row,
                                "numrows": numrows,
                            },
                        )
                    )
                )
            )
        self.refresh()
        if self.extra_end_insert_rows_rc_func is not None:
            self.extra_end_insert_rows_rc_func(InsertEvent("end_insert_rows", data_ins_row, displayed_ins_row, numrows))
        self.parentframe.emit_event("<<SheetModified>>")

    def del_cols_rc(self, event=None, c=None):
        seld_cols = sorted(self.get_selected_cols())
        curr = self.currently_selected()
        if not seld_cols or not curr:
            return
        if (
            self.CH.popup_menu_loc is None
            or self.CH.popup_menu_loc < seld_cols[0]
            or self.CH.popup_menu_loc > seld_cols[-1]
        ):
            c = seld_cols[0]
        else:
            c = self.CH.popup_menu_loc
        seld_cols = get_seq_without_gaps_at_index(seld_cols, c)
        self.deselect("all", redraw=False)
        self.create_selected(
            0,
            seld_cols[0],
            len(self.row_positions) - 1,
            seld_cols[-1] + 1,
            "columns",
        )
        self.set_currently_selected(0, seld_cols[0], type_="column")
        seldmax = seld_cols[-1] if self.all_columns_displayed else self.displayed_columns[seld_cols[-1]]
        if self.extra_begin_del_cols_rc_func is not None:
            try:
                self.extra_begin_del_cols_rc_func(DeleteRowColumnEvent("begin_delete_columns", seld_cols))
            except Exception:
                return
        seldset = set(seld_cols) if self.all_columns_displayed else set(self.displayed_columns[c] for c in seld_cols)
        if self.undo_enabled:
            undo_storage = {
                "deleted_cols": {},
                "colwidths": {},
                "deleted_header_values": {},
                "selection_boxes": self.get_boxes(),
                "displayed_columns": list(self.displayed_columns)
                if not isinstance(self.displayed_columns, int)
                else int(self.displayed_columns),
                "cell_options": {k: v.copy() for k, v in self.cell_options.items()},
                "col_options": {k: v.copy() for k, v in self.col_options.items()},
                "CH_cell_options": {k: v.copy() for k, v in self.CH.cell_options.items()},
            }
            for c in reversed(seld_cols):
                undo_storage["colwidths"][c] = self.col_positions[c + 1] - self.col_positions[c]
                datacn = c if self.all_columns_displayed else self.displayed_columns[c]
                for rn in range(len(self.data)):
                    if datacn not in undo_storage["deleted_cols"]:
                        undo_storage["deleted_cols"][datacn] = {}
                    try:
                        undo_storage["deleted_cols"][datacn][rn] = self.data[rn].pop(datacn)
                    except Exception:
                        continue
                try:
                    undo_storage["deleted_header_values"][datacn] = self._headers.pop(datacn)
                except Exception:
                    continue
        else:
            for c in reversed(seld_cols):
                datacn = c if self.all_columns_displayed else self.displayed_columns[c]
                for rn in range(len(self.data)):
                    del self.data[rn][datacn]
                try:
                    del self._headers[datacn]
                except Exception:
                    continue
        if self.undo_enabled:
            self.undo_storage.append(("delete_cols", undo_storage))
        for c in reversed(seld_cols):
            self.del_col_position(c, deselect_all=False)
        numcols = len(seld_cols)
        self.cell_options = {
            (rn, cn if cn < seldmax else cn - numcols): t2
            for (rn, cn), t2 in self.cell_options.items()
            if cn not in seldset
        }
        self.col_options = {
            cn if cn < seldmax else cn - numcols: t for cn, t in self.col_options.items() if cn not in seldset
        }
        self.CH.cell_options = {
            cn if cn < seldmax else cn - numcols: t for cn, t in self.CH.cell_options.items() if cn not in seldset
        }
        self.deselect("all", redraw=False)
        if not self.all_columns_displayed:
            self.displayed_columns = [c for c in self.displayed_columns if c not in seldset]
            for c in sorted(seldset):
                self.displayed_columns = [dc if c > dc else dc - 1 for dc in self.displayed_columns]
        self.refresh()
        if self.extra_end_del_cols_rc_func is not None:
            self.extra_end_del_cols_rc_func(DeleteRowColumnEvent("end_delete_columns", seld_cols))
        self.parentframe.emit_event("<<SheetModified>>")

    def del_rows_rc(self, event=None, r=None):
        seld_rows = sorted(self.get_selected_rows())
        curr = self.currently_selected()
        if not seld_rows or not curr:
            return
        if (
            self.RI.popup_menu_loc is None
            or self.RI.popup_menu_loc < seld_rows[0]
            or self.RI.popup_menu_loc > seld_rows[-1]
        ):
            r = seld_rows[0]
        else:
            r = self.RI.popup_menu_loc
        seld_rows = get_seq_without_gaps_at_index(seld_rows, r)
        self.deselect("all", redraw=False)
        self.create_selected(
            seld_rows[0],
            0,
            seld_rows[-1] + 1,
            len(self.col_positions) - 1,
            "rows",
        )
        self.set_currently_selected(seld_rows[0], 0, type_="row")
        seldmax = seld_rows[-1] if self.all_rows_displayed else self.displayed_rows[seld_rows[-1]]
        if self.extra_begin_del_rows_rc_func is not None:
            try:
                self.extra_begin_del_rows_rc_func(DeleteRowColumnEvent("begin_delete_rows", seld_rows))
            except Exception:
                return
        seldset = set(seld_rows) if self.all_rows_displayed else set(self.displayed_rows[r] for r in seld_rows)
        if self.undo_enabled:
            undo_storage = {
                "deleted_rows": [],
                "rowheights": {},
                "deleted_index_values": [],
                "selection_boxes": self.get_boxes(),
                "displayed_rows": list(self.displayed_rows)
                if not isinstance(self.displayed_rows, int)
                else int(self.displayed_rows),
                "cell_options": {k: v.copy() for k, v in self.cell_options.items()},
                "row_options": {k: v.copy() for k, v in self.row_options.items()},
                "RI_cell_options": {k: v.copy() for k, v in self.RI.cell_options.items()},
            }
            for r in reversed(seld_rows):
                undo_storage["rowheights"][r] = self.row_positions[r + 1] - self.row_positions[r]
                datarn = r if self.all_rows_displayed else self.displayed_rows[r]
                undo_storage["deleted_rows"].append((datarn, self.data.pop(datarn)))
                try:
                    undo_storage["deleted_index_values"].append((datarn, self._row_index.pop(datarn)))
                except Exception:
                    continue
        else:
            for r in reversed(seld_rows):
                datarn = r if self.all_rows_displayed else self.displayed_rows[r]
                del self.data[datarn]
                try:
                    del self._row_index[datarn]
                except Exception:
                    continue
        if self.undo_enabled:
            self.undo_storage.append(("delete_rows", undo_storage))
        for r in reversed(seld_rows):
            self.del_row_position(r, deselect_all=False)
        numrows = len(seld_rows)
        self.cell_options = {
            (rn if rn < seldmax else rn - numrows, cn): t2
            for (rn, cn), t2 in self.cell_options.items()
            if rn not in seldset
        }
        self.row_options = {
            rn if rn < seldmax else rn - numrows: t for rn, t in self.row_options.items() if rn not in seldset
        }
        self.RI.cell_options = {
            rn if rn < seldmax else rn - numrows: t for rn, t in self.RI.cell_options.items() if rn not in seldset
        }
        self.deselect("all", redraw=False)
        if not self.all_rows_displayed:
            self.displayed_rows = [r for r in self.displayed_rows if r not in seldset]
            for r in sorted(seldset):
                self.displayed_rows = [dr if r > dr else dr - 1 for dr in self.displayed_rows]
        self.refresh()
        if self.extra_end_del_rows_rc_func is not None:
            self.extra_end_del_rows_rc_func(DeleteRowColumnEvent("end_delete_rows", seld_rows))
        self.parentframe.emit_event("<<SheetModified>>")

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

    def display_rows(
        self,
        rows=None,
        all_rows_displayed=None,
        reset_row_positions=True,
        deselect_all=True,
    ):
        if rows is None and all_rows_displayed is None:
            return list(range(self.total_data_rows())) if self.all_rows_displayed else self.displayed_rows
        total_data_rows = None
        if (rows is not None and rows != self.displayed_rows) or (all_rows_displayed and not self.all_rows_displayed):
            self.undo_storage = deque(maxlen=self.max_undos)
        if rows is not None and rows != self.displayed_rows:
            self.displayed_rows = sorted(rows)
        if all_rows_displayed:
            if not self.all_rows_displayed:
                total_data_rows = self.total_data_rows()
                self.displayed_rows = list(range(total_data_rows))
            self.all_rows_displayed = True
        elif all_rows_displayed is not None and not all_rows_displayed:
            self.all_rows_displayed = False
        if reset_row_positions:
            self.reset_row_positions(nrows=total_data_rows)
        if deselect_all:
            self.deselect("all", redraw=False)

    def display_columns(
        self,
        columns=None,
        all_columns_displayed=None,
        reset_col_positions=True,
        deselect_all=True,
    ):
        if columns is None and all_columns_displayed is None:
            return list(range(self.total_data_cols())) if self.all_columns_displayed else self.displayed_columns
        total_data_cols = None
        if (columns is not None and columns != self.displayed_columns) or (
            all_columns_displayed and not self.all_columns_displayed
        ):
            self.undo_storage = deque(maxlen=self.max_undos)
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
            self.reset_col_positions(ncols=total_data_cols)
        if deselect_all:
            self.deselect("all", redraw=False)

    def headers(
        self,
        newheaders=None,
        index=None,
        reset_col_positions=False,
        show_headers_if_not_sheet=True,
        redraw=False,
    ):
        if newheaders is not None:
            if isinstance(newheaders, (list, tuple)):
                self._headers = list(newheaders) if isinstance(newheaders, tuple) else newheaders
            elif isinstance(newheaders, int):
                self._headers = int(newheaders)
            elif isinstance(self._headers, list) and isinstance(index, int):
                if index >= len(self._headers):
                    self.CH.fix_header(index)
                self._headers[index] = f"{newheaders}"
            elif not isinstance(newheaders, (list, tuple, int)) and index is None:
                try:
                    self._headers = list(newheaders)
                except Exception:
                    raise ValueError("New header must be iterable or int (use int to use a row as the header")
            if reset_col_positions:
                self.reset_col_positions()
            elif (
                show_headers_if_not_sheet
                and isinstance(self._headers, list)
                and (self.col_positions == [0] or not self.col_positions)
            ):
                colpos = int(self.default_column_width)
                if self.all_columns_displayed:
                    self.col_positions = list(accumulate(chain([0], (colpos for c in range(len(self._headers))))))
                else:
                    self.col_positions = list(
                        accumulate(
                            chain(
                                [0],
                                (colpos for c in range(len(self.displayed_columns))),
                            )
                        )
                    )
            if redraw:
                self.refresh()
        else:
            if not isinstance(self._headers, int) and index is not None and isinstance(index, int):
                return self._headers[index]
            else:
                return self._headers

    def row_index(
        self,
        newindex=None,
        index=None,
        reset_row_positions=False,
        show_index_if_not_sheet=True,
        redraw=False,
    ):
        if newindex is not None:
            if not self._row_index and not isinstance(self._row_index, int):
                self.RI.set_width(self.default_index_width, set_TL=True)
            if isinstance(newindex, (list, tuple)):
                self._row_index = list(newindex) if isinstance(newindex, tuple) else newindex
            elif isinstance(newindex, int):
                self._row_index = int(newindex)
            elif isinstance(index, int):
                if index >= len(self._row_index):
                    self.RI.fix_index(index)
                self._row_index[index] = f"{newindex}"
            elif not isinstance(newindex, (list, tuple, int)) and index is None:
                try:
                    self._row_index = list(newindex)
                except Exception:
                    raise ValueError("New index must be iterable or int (use int to use a column as the index")
            if reset_row_positions:
                self.reset_row_positions()
            elif (
                show_index_if_not_sheet
                and isinstance(self._row_index, list)
                and (self.row_positions == [0] or not self.row_positions)
            ):
                rowpos = self.default_row_height[1]
                if self.all_rows_displayed:
                    self.row_positions = list(accumulate(chain([0], (rowpos for r in range(len(self._row_index))))))
                else:
                    self.row_positions = list(accumulate(chain([0], (rowpos for r in range(len(self.displayed_rows))))))

            if redraw:
                self.refresh()
        else:
            if not isinstance(self._row_index, int) and index is not None and isinstance(index, int):
                return self._row_index[index]
            else:
                return self._row_index

    def total_data_cols(self, include_header=True):
        h_total = 0
        d_total = 0
        if include_header:
            if isinstance(self._headers, (list, tuple)):
                h_total = len(self._headers)
        try:
            d_total = len(max(self.data, key=len))
        except Exception:
            pass
        return h_total if h_total > d_total else d_total

    def total_data_rows(self, include_index=True):
        i_total = 0
        d_total = 0
        if include_index:
            if isinstance(self._row_index, (list, tuple)):
                i_total = len(self._row_index)
        d_total = len(self.data)
        return i_total if i_total > d_total else d_total

    def data_dimensions(self, total_rows=None, total_columns=None):
        if total_rows is None and total_columns is None:
            return self.total_data_rows(), self.total_data_cols()
        if total_rows is not None:
            if len(self.data) < total_rows:
                ncols = self.total_data_cols() if total_columns is None else total_columns
                self.data.extend([self.get_empty_row_seq(r, ncols) for r in range(total_rows - len(self.data))])
            else:
                self.data[total_rows:] = []
        if total_columns is not None:
            self.data[:] = [
                r[:total_columns]
                if len(r) > total_columns
                else r + self.get_empty_row_seq(rn, end=len(r) + total_columns - len(r), start=len(r))
                for rn, r in enumerate(self.data)
            ]

    def equalize_data_row_lengths(self, include_header=False, total_columns=None):
        total_columns = self.total_data_cols() if total_columns is None else total_columns
        if include_header and total_columns > len(self._headers):
            self.CH.fix_header(total_columns)
        self.data[:] = [
            (r + self.get_empty_row_seq(rn, end=len(r) + total_columns - len(r), start=len(r)))
            if total_columns > len(r)
            else r
            for rn, r in enumerate(self.data)
        ]
        return total_columns

    def get_canvas_visible_area(self):
        return (
            self.canvasx(0),
            self.canvasy(0),
            self.canvasx(self.winfo_width()),
            self.canvasy(self.winfo_height()),
        )

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

    def redraw_highlight_get_text_fg(
        self,
        r,
        c,
        fc,
        fr,
        sc,
        sr,
        c_2_,
        c_3_,
        c_4_,
        selections,
        datarn,
        datacn,
        can_width,
    ):
        redrawn = False
        kwargs = self.get_cell_kwargs(datarn, datacn, key="highlight")
        if kwargs:
            if kwargs[0] is not None:
                c_1 = kwargs[0] if kwargs[0].startswith("#") else Color_Map_[kwargs[0]]
            if "cells" in selections and (r, c) in selections["cells"]:
                tf = (
                    self.table_selected_cells_fg
                    if kwargs[1] is None or self.display_selected_fg_over_highlights
                    else kwargs[1]
                )
                if kwargs[0] is not None:
                    fill = (
                        f"#{int((int(c_1[1:3], 16) + c_2_[0]) / 2):02X}"
                        + f"{int((int(c_1[3:5], 16) + c_2_[1]) / 2):02X}"
                        + f"{int((int(c_1[5:], 16) + c_2_[2]) / 2):02X}"
                    )
            elif "rows" in selections and r in selections["rows"]:
                tf = (
                    self.table_selected_rows_fg
                    if kwargs[1] is None or self.display_selected_fg_over_highlights
                    else kwargs[1]
                )
                if kwargs[0] is not None:
                    fill = (
                        f"#{int((int(c_1[1:3], 16) + c_4_[0]) / 2):02X}"
                        + f"{int((int(c_1[3:5], 16) + c_4_[1]) / 2):02X}"
                        + f"{int((int(c_1[5:], 16) + c_4_[2]) / 2):02X}"
                    )
            elif "columns" in selections and c in selections["columns"]:
                tf = (
                    self.table_selected_columns_fg
                    if kwargs[1] is None or self.display_selected_fg_over_highlights
                    else kwargs[1]
                )
                if kwargs[0] is not None:
                    fill = (
                        f"#{int((int(c_1[1:3], 16) + c_3_[0]) / 2):02X}"
                        + f"{int((int(c_1[3:5], 16) + c_3_[1]) / 2):02X}"
                        + f"{int((int(c_1[5:], 16) + c_3_[2]) / 2):02X}"
                    )
            else:
                tf = self.table_fg if kwargs[1] is None else kwargs[1]
                if kwargs[0] is not None:
                    fill = kwargs[0]
            if kwargs[0] is not None:
                redrawn = self.redraw_highlight(
                    fc + 1,
                    fr + 1,
                    sc,
                    sr,
                    fill=fill,
                    outline=self.table_fg
                    if self.get_cell_kwargs(datarn, datacn, key="dropdown") and self.show_dropdown_borders
                    else "",
                    tag="hi",
                    can_width=can_width if (len(kwargs) > 2 and kwargs[2]) else None,
                )
        elif not kwargs:
            if "cells" in selections and (r, c) in selections["cells"]:
                tf = self.table_selected_cells_fg
            elif "rows" in selections and r in selections["rows"]:
                tf = self.table_selected_rows_fg
            elif "columns" in selections and c in selections["columns"]:
                tf = self.table_selected_columns_fg
            else:
                tf = self.table_fg
        return tf, redrawn

    def redraw_highlight(self, x1, y1, x2, y2, fill, outline, tag, can_width=None, pc=None):
        config = (fill, outline)
        if type(pc) != int or pc >= 100 or pc <= 0:
            coords = (
                x1 - 1 if outline else x1,
                y1 - 1 if outline else y1,
                x2 if can_width is None else x2 + can_width,
                y2,
            )
        else:
            coords = (x1, y1, (x2 - x1) * (pc / 100), y2)
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
        if draw_outline and self.show_dropdown_borders:
            self.redraw_highlight(x1 + 1, y1 + 1, x2, y2, fill="", outline=self.table_fg, tag=tag)
        if draw_arrow:
            topysub = floor(self.half_txt_h / 2)
            mid_y = y1 + floor(self.min_row_height / 2)
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

    def get_checkbox_points(self, x1, y1, x2, y2, radius=8):
        return [
            x1 + radius,
            y1,
            x1 + radius,
            y1,
            x2 - radius,
            y1,
            x2 - radius,
            y1,
            x2,
            y1,
            x2,
            y1 + radius,
            x2,
            y1 + radius,
            x2,
            y2 - radius,
            x2,
            y2 - radius,
            x2,
            y2,
            x2 - radius,
            y2,
            x2 - radius,
            y2,
            x1 + radius,
            y2,
            x1 + radius,
            y2,
            x1,
            y2,
            x1,
            y2 - radius,
            x1,
            y2 - radius,
            x1,
            y1 + radius,
            x1,
            y1 + radius,
            x1,
            y1,
        ]

    def redraw_checkbox(self, x1, y1, x2, y2, fill, outline, tag, draw_check=False):
        points = self.get_checkbox_points(x1, y1, x2, y2)
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
            x1 = x1 + 4
            y1 = y1 + 4
            x2 = x2 - 3
            y2 = y2 - 3
            points = self.get_checkbox_points(x1, y1, x2, y2, radius=4)
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

    def main_table_redraw_grid_and_text(
        self,
        redraw_header=False,
        redraw_row_index=False,
        redraw_table=True,
    ):
        try:
            can_width = self.winfo_width()
            can_height = self.winfo_height()
        except Exception:
            return False
        row_pos_exists = self.row_positions != [0] and self.row_positions
        col_pos_exists = self.col_positions != [0] and self.col_positions
        resized_cols = False
        resized_rows = False
        if self.auto_resize_columns and self.allow_auto_resize_columns and col_pos_exists:
            max_w = int(can_width)
            max_w -= self.empty_horizontal
            if (len(self.col_positions) - 1) * self.auto_resize_columns < max_w:
                resized_cols = True
                change = int((max_w - self.col_positions[-1]) / (len(self.col_positions) - 1))
                widths = [
                    int(b - a) + change - 1
                    for a, b in zip(self.col_positions, islice(self.col_positions, 1, len(self.col_positions)))
                ]
                diffs = {}
                for i, w in enumerate(widths):
                    if w < self.auto_resize_columns:
                        diffs[i] = self.auto_resize_columns - w
                        widths[i] = self.auto_resize_columns
                if diffs and len(diffs) < len(widths):
                    change = sum(diffs.values()) / (len(widths) - len(diffs))
                    for i, w in enumerate(widths):
                        if i not in diffs:
                            widths[i] -= change
                self.col_positions = list(accumulate(chain([0], widths)))
        if self.auto_resize_rows and self.allow_auto_resize_rows and row_pos_exists:
            max_h = int(can_height)
            max_h -= self.empty_vertical
            if (len(self.row_positions) - 1) * self.auto_resize_rows < max_h:
                resized_rows = True
                change = int((max_h - self.row_positions[-1]) / (len(self.row_positions) - 1))
                heights = [
                    int(b - a) + change - 1
                    for a, b in zip(self.row_positions, islice(self.row_positions, 1, len(self.row_positions)))
                ]
                diffs = {}
                for i, h in enumerate(heights):
                    if h < self.auto_resize_rows:
                        diffs[i] = self.auto_resize_rows - h
                        heights[i] = self.auto_resize_rows
                if diffs and len(diffs) < len(heights):
                    change = sum(diffs.values()) / (len(heights) - len(diffs))
                    for i, h in enumerate(heights):
                        if i not in diffs:
                            heights[i] -= change
                self.row_positions = list(accumulate(chain([0], heights)))
        if resized_cols or resized_rows:
            self.recreate_all_selection_boxes()
        last_col_line_pos = self.col_positions[-1] + 1
        last_row_line_pos = self.row_positions[-1] + 1
        if can_width >= last_col_line_pos + self.empty_horizontal and self.parentframe.xscroll_showing:
            self.parentframe.xscroll.grid_forget()
            self.parentframe.xscroll_showing = False
        elif (
            can_width < last_col_line_pos + self.empty_horizontal
            and not self.parentframe.xscroll_showing
            and not self.parentframe.xscroll_disabled
            and can_height > 40
        ):
            self.parentframe.xscroll.grid(row=2, column=0, columnspan=2, sticky="nswe")
            self.parentframe.xscroll_showing = True
        if can_height >= last_row_line_pos + self.empty_vertical and self.parentframe.yscroll_showing:
            self.parentframe.yscroll.grid_forget()
            self.parentframe.yscroll_showing = False
        elif (
            can_height < last_row_line_pos + self.empty_vertical
            and not self.parentframe.yscroll_showing
            and not self.parentframe.yscroll_disabled
            and can_width > 40
        ):
            self.parentframe.yscroll.grid(row=0, column=2, rowspan=3, sticky="nswe")
            self.parentframe.yscroll_showing = True
        self.configure(
            scrollregion=(
                0,
                0,
                last_col_line_pos + self.empty_horizontal + 2,
                last_row_line_pos + self.empty_vertical + 2,
            )
        )
        scrollpos_bot = self.canvasy(can_height)
        end_row = bisect.bisect_right(self.row_positions, scrollpos_bot)
        if not scrollpos_bot >= self.row_positions[-1]:
            end_row += 1
        if redraw_row_index and self.show_index:
            self.RI.auto_set_index_width(end_row - 1)
            # return
        scrollpos_left = self.canvasx(0)
        scrollpos_top = self.canvasy(0)
        scrollpos_right = self.canvasx(can_width)
        start_row = bisect.bisect_left(self.row_positions, scrollpos_top)
        start_col = bisect.bisect_left(self.col_positions, scrollpos_left)
        end_col = bisect.bisect_right(self.col_positions, scrollpos_right)
        self.row_width_resize_bbox = (
            scrollpos_left,
            scrollpos_top,
            scrollpos_left + 2,
            scrollpos_bot,
        )
        self.header_height_resize_bbox = (
            scrollpos_left + 6,
            scrollpos_top,
            scrollpos_right,
            scrollpos_top + 2,
        )
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
        if self.show_horizontal_grid and row_pos_exists:
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
                    points.extend(
                        [
                            self.canvasx(0) - 1,
                            draw_y,
                            x_grid_stop,
                            draw_y,
                            x_grid_stop,
                            self.row_positions[r + 1] if len(self.row_positions) - 1 > r else draw_y,
                        ]
                    )
                elif st_or_end == "end":
                    points.extend(
                        [
                            x_grid_stop,
                            draw_y,
                            self.canvasx(0) - 1,
                            draw_y,
                            self.canvasx(0) - 1,
                            self.row_positions[r + 1] if len(self.row_positions) - 1 > r else draw_y,
                        ]
                    )
            if points:
                if self.hidd_grid:
                    t, sh = self.hidd_grid.popitem()
                    self.coords(t, points)
                    if sh:
                        self.itemconfig(
                            t,
                            fill=self.table_grid_fg,
                            capstyle=tk.BUTT,
                            joinstyle=tk.ROUND,
                            width=1,
                        )
                    else:
                        self.itemconfig(
                            t,
                            fill=self.table_grid_fg,
                            capstyle=tk.BUTT,
                            joinstyle=tk.ROUND,
                            width=1,
                            state="normal",
                        )
                    self.disp_grid[t] = True
                else:
                    self.disp_grid[
                        self.create_line(
                            points,
                            fill=self.table_grid_fg,
                            capstyle=tk.BUTT,
                            joinstyle=tk.ROUND,
                            width=1,
                            tag="g",
                        )
                    ] = True
        if self.show_vertical_grid and col_pos_exists:
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
                    points.extend(
                        [
                            draw_x,
                            scrollpos_top - 1,
                            draw_x,
                            y_grid_stop,
                            self.col_positions[c + 1] if len(self.col_positions) - 1 > c else draw_x,
                            y_grid_stop,
                        ]
                    )
                elif st_or_end == "end":
                    points.extend(
                        [
                            draw_x,
                            y_grid_stop,
                            draw_x,
                            scrollpos_top - 1,
                            self.col_positions[c + 1] if len(self.col_positions) - 1 > c else draw_x,
                            scrollpos_top - 1,
                        ]
                    )
            if points:
                if self.hidd_grid:
                    t, sh = self.hidd_grid.popitem()
                    self.coords(t, points)
                    if sh:
                        self.itemconfig(
                            t,
                            fill=self.table_grid_fg,
                            capstyle=tk.BUTT,
                            joinstyle=tk.ROUND,
                            width=1,
                        )
                    else:
                        self.itemconfig(
                            t,
                            fill=self.table_grid_fg,
                            capstyle=tk.BUTT,
                            joinstyle=tk.ROUND,
                            width=1,
                            state="normal",
                        )
                    self.disp_grid[t] = True
                else:
                    self.disp_grid[
                        self.create_line(
                            points,
                            fill=self.table_grid_fg,
                            capstyle=tk.BUTT,
                            joinstyle=tk.ROUND,
                            width=1,
                            tag="g",
                        )
                    ] = True
        if start_row > 0:
            start_row -= 1
        if start_col > 0:
            start_col -= 1
        end_row -= 1
        selections = self.get_redraw_selections(start_row, end_row, start_col, end_col)
        c_2 = (
            self.table_selected_cells_bg
            if self.table_selected_cells_bg.startswith("#")
            else Color_Map_[self.table_selected_cells_bg]
        )
        c_2_ = (int(c_2[1:3], 16), int(c_2[3:5], 16), int(c_2[5:], 16))
        c_3 = (
            self.table_selected_columns_bg
            if self.table_selected_columns_bg.startswith("#")
            else Color_Map_[self.table_selected_columns_bg]
        )
        c_3_ = (int(c_3[1:3], 16), int(c_3[3:5], 16), int(c_3[5:], 16))
        c_4 = (
            self.table_selected_rows_bg
            if self.table_selected_rows_bg.startswith("#")
            else Color_Map_[self.table_selected_rows_bg]
        )
        c_4_ = (int(c_4[1:3], 16), int(c_4[3:5], 16), int(c_4[5:], 16))
        rows_ = tuple(range(start_row, end_row))
        font = self.table_font
        if redraw_table:
            for c in range(start_col, end_col - 1):
                for r in rows_:
                    rtopgridln = self.row_positions[r]
                    rbotgridln = self.row_positions[r + 1]
                    if rbotgridln - rtopgridln < self.txt_h:
                        continue
                    cleftgridln = self.col_positions[c]
                    crightgridln = self.col_positions[c + 1]

                    datarn = r if self.all_rows_displayed else self.displayed_rows[r]
                    datacn = c if self.all_columns_displayed else self.displayed_columns[c]

                    fill, dd_drawn = self.redraw_highlight_get_text_fg(
                        r,
                        c,
                        cleftgridln,
                        rtopgridln,
                        crightgridln,
                        rbotgridln,
                        c_2_,
                        c_3_,
                        c_4_,
                        selections,
                        datarn,
                        datacn,
                        can_width,
                    )
                    align = self.get_cell_kwargs(datarn, datacn, key="align")
                    if align:
                        align = align
                    else:
                        align = self.align
                    kwargs = self.get_cell_kwargs(datarn, datacn, key="dropdown")
                    if align == "w":
                        draw_x = cleftgridln + 3
                        if kwargs:
                            mw = crightgridln - cleftgridln - self.txt_h - 2
                            self.redraw_dropdown(
                                cleftgridln,
                                rtopgridln,
                                crightgridln,
                                self.row_positions[r + 1],
                                fill=fill,
                                outline=fill,
                                tag=f"dd_{r}_{c}",
                                draw_outline=not dd_drawn,
                                draw_arrow=mw >= 5,
                                dd_is_open=kwargs["window"] != "no dropdown open",
                            )
                        else:
                            mw = crightgridln - cleftgridln - 1
                    elif align == "e":
                        if kwargs:
                            mw = crightgridln - cleftgridln - self.txt_h - 2
                            draw_x = crightgridln - 5 - self.txt_h
                            self.redraw_dropdown(
                                cleftgridln,
                                rtopgridln,
                                crightgridln,
                                self.row_positions[r + 1],
                                fill=fill,
                                outline=fill,
                                tag=f"dd_{r}_{c}",
                                draw_outline=not dd_drawn,
                                draw_arrow=mw >= 5,
                                dd_is_open=kwargs["window"] != "no dropdown open",
                            )
                        else:
                            mw = crightgridln - cleftgridln - 1
                            draw_x = crightgridln - 3
                    elif align == "center":
                        stop = cleftgridln + 5
                        if kwargs:
                            mw = crightgridln - cleftgridln - self.txt_h - 2
                            draw_x = cleftgridln + ceil((crightgridln - cleftgridln - self.txt_h) / 2)
                            self.redraw_dropdown(
                                cleftgridln,
                                rtopgridln,
                                crightgridln,
                                self.row_positions[r + 1],
                                fill=fill,
                                outline=fill,
                                tag=f"dd_{r}_{c}",
                                draw_outline=not dd_drawn,
                                draw_arrow=mw >= 5,
                                dd_is_open=kwargs["window"] != "no dropdown open",
                            )
                        else:
                            mw = crightgridln - cleftgridln - 1
                            draw_x = cleftgridln + floor((crightgridln - cleftgridln) / 2)
                    kwargs = self.get_cell_kwargs(datarn, datacn, key="checkbox")
                    if kwargs:
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
                                draw_check = self.data[datarn][datacn]
                            except Exception:
                                draw_check = False
                            self.redraw_checkbox(
                                cleftgridln + 2,
                                rtopgridln + 2,
                                cleftgridln + self.txt_h + 3,
                                rtopgridln + self.txt_h + 3,
                                fill=fill if kwargs["state"] == "normal" else self.table_grid_fg,
                                outline="",
                                tag="cb",
                                draw_check=draw_check,
                            )
                    lns = self.get_valid_cell_data_as_str(datarn, datacn, get_displayed=True).split("\n")
                    if (
                        lns != [""]
                        and mw > self.txt_w
                        and not (
                            (align == "w" and draw_x > scrollpos_right)
                            or (align == "e" and cleftgridln + 5 > scrollpos_right)
                            or (align == "center" and stop > scrollpos_right)
                        )
                    ):
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
                                draw_y += self.xtra_lines_increment
                                if draw_y + self.half_txt_h - 1 > rbotgridln:
                                    break
        if redraw_table:
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
            if self.show_selected_cells_border:
                self.tag_raise("cellsbd")
                self.tag_raise("selected")
                self.tag_raise("rowsbd")
                self.tag_raise("columnsbd")
        if redraw_header and self.show_header:
            self.CH.redraw_grid_and_text(
                last_col_line_pos,
                scrollpos_left,
                x_stop,
                start_col,
                end_col,
                scrollpos_right,
                col_pos_exists,
            )
        if redraw_row_index and self.show_index:
            self.RI.redraw_grid_and_text(
                last_row_line_pos,
                scrollpos_top,
                y_stop,
                start_row,
                end_row + 1,
                scrollpos_bot,
                row_pos_exists,
            )
        self.parentframe.emit_event("<<SheetRedrawn>>")
        return True

    def get_all_selection_items(self):
        return sorted(
            self.find_withtag("cells")
            + self.find_withtag("rows")
            + self.find_withtag("columns")
            + self.find_withtag("selected")
        )

    def get_boxes(self, include_current=True):
        boxes = {}
        for item in self.get_all_selection_items():
            alltags = self.gettags(item)
            if alltags[0] == "cells":
                boxes[tuple(int(e) for e in alltags[1].split("_") if e)] = "cells"
            elif alltags[0] == "rows":
                boxes[tuple(int(e) for e in alltags[1].split("_") if e)] = "rows"
            elif alltags[0] == "columns":
                boxes[tuple(int(e) for e in alltags[1].split("_") if e)] = "columns"
            elif include_current and alltags[0] == "selected":
                boxes[tuple(int(e) for e in alltags[1].split("_") if e)] = f"{alltags[2]}"
        return boxes

    def reselect_from_get_boxes(self, boxes):
        for (r1, c1, r2, c2), v in boxes.items():
            if r2 < len(self.row_positions) and c2 < len(self.col_positions):
                if v == "cells":
                    self.create_selected(r1, c1, r2, c2, "cells")
                elif v == "rows":
                    self.create_selected(r1, c1, r2, c2, "rows")
                elif v == "columns":
                    self.create_selected(r1, c1, r2, c2, "columns")
                elif v in ("cell", "row", "column"):  # currently selected
                    self.set_currently_selected(r1, c1, type_=v)

    def delete_selected(self, r1=None, c1=None, r2=None, c2=None, type_=None):
        deleted_boxes = {}
        tags_to_del = set()
        box1 = (r1, c1, r2, c2)
        for s in self.get_selection_tags_from_type(type_):
            for item in self.find_withtag(s):
                alltags = self.gettags(item)
                if alltags:
                    box2 = tuple(int(e) for e in alltags[1].split("_") if e)
                    if box1 == box2:
                        tags_to_del.add(alltags)
                        deleted_boxes[box2] = (
                            "cells"
                            if alltags[0].startswith("cell")
                            else "rows"
                            if alltags[0].startswith("row")
                            else "columns"
                        )
                        self.delete(item)
        for canvas in (self.RI, self.CH):
            for item in canvas.find_withtag(s):
                if canvas.gettags(item) in tags_to_del:
                    canvas.delete(item)

    def get_selection_tags_from_type(self, type_):
        if type_ == "cells":
            return {"cells", "cellsbd"}
        if type_ == "rows":
            return {"rows", "rowsbd"}
        elif type_ == "columns":
            return {"columns", "columnsbd"}
        else:
            return {"cells", "cellsbd", "rows", "rowsbd", "columns", "columnsbd"}

    def delete_selection_rects(self, cells=True, rows=True, cols=True, delete_current=True):
        deleted_boxes = {}
        if cells:
            for item in self.find_withtag("cells"):
                alltags = self.gettags(item)
                if alltags:
                    deleted_boxes[tuple(int(e) for e in alltags[1].split("_") if e)] = "cells"
            self.delete("cells", "cellsbd")
            self.RI.delete("cells", "cellsbd")
            self.CH.delete("cells", "cellsbd")
        if rows:
            for item in self.find_withtag("rows"):
                alltags = self.gettags(item)
                if alltags:
                    deleted_boxes[tuple(int(e) for e in alltags[1].split("_") if e)] = "rows"
            self.delete("rows", "rowsbd")
            self.RI.delete("rows", "rowsbd")
            self.CH.delete("rows", "rowsbd")
        if cols:
            for item in self.find_withtag("columns"):
                alltags = self.gettags(item)
                if alltags:
                    deleted_boxes[tuple(int(e) for e in alltags[1].split("_") if e)] = "columns"
            self.delete("columns", "columnsbd")
            self.RI.delete("columns", "columnsbd")
            self.CH.delete("columns", "columnsbd")
        if delete_current:
            self.delete("selected")
            self.RI.delete("selected")
            self.CH.delete("selected")
        return deleted_boxes

    def currently_selected(self):
        items = self.find_withtag("selected")
        if not items:
            return tuple()
        alltags = self.gettags(items[0])
        box = tuple(int(e) for e in alltags[1].split("_") if e)
        return CurrentlySelectedClass(box[0], box[1], alltags[2])

    def get_tags_of_current(self):
        items = self.find_withtag("selected")
        if items:
            return self.gettags(items[0])
        else:
            return tuple()

    def set_currently_selected(self, r, c, type_="cell"):  # cell, column or row
        r1, c1, r2, c2 = r, c, r + 1, c + 1
        self.delete("selected")
        self.RI.delete("selected")
        self.CH.delete("selected")
        if self.col_positions == [0]:
            c1 = 0
            c2 = 0
        if self.row_positions == [0]:
            r1 = 0
            r2 = 0
        tagr = ("selected", f"{r1}_{c1}_{r2}_{c2}", type_)
        tag_index_header = ("cells", f"{r1}_{c1}_{r2}_{c2}", "selected")
        if type_ == "cell":
            outline = self.table_selected_cells_border_fg
        elif type_ == "row":
            outline = self.table_selected_rows_border_fg
        elif type_ == "column":
            outline = self.table_selected_columns_border_fg
        if self.show_selected_cells_border:
            b = self.create_rectangle(
                self.col_positions[c1] + 1,
                self.row_positions[r1] + 1,
                self.col_positions[c2],
                self.row_positions[r2],
                fill="",
                outline=outline,
                width=2,
                tags=tagr,
            )
            self.tag_raise(b)
        else:
            b = self.create_rectangle(
                self.col_positions[c1],
                self.row_positions[r1],
                self.col_positions[c2],
                self.row_positions[r2],
                fill=self.table_selected_cells_bg,
                outline="",
                tags=tagr,
            )
            self.tag_lower(b)
        ri = self.RI.create_rectangle(
            0,
            self.row_positions[r1],
            self.RI.current_width - 1,
            self.row_positions[r2],
            fill=self.RI.index_selected_cells_bg,
            outline="",
            tags=tag_index_header,
        )
        ch = self.CH.create_rectangle(
            self.col_positions[c1],
            0,
            self.col_positions[c2],
            self.CH.current_height - 1,
            fill=self.CH.header_selected_cells_bg,
            outline="",
            tags=tag_index_header,
        )
        self.RI.tag_lower(ri)
        self.CH.tag_lower(ch)
        return b

    def set_current_to_last(self):
        if not self.currently_selected():
            items = sorted(self.find_withtag("cells") + self.find_withtag("rows") + self.find_withtag("columns"))
            if items:
                last = self.gettags(items[-1])
                r1, c1, r2, c2 = tuple(int(e) for e in last[1].split("_") if e)
                if last[0] == "cells":
                    return self.gettags(self.set_currently_selected(r1, c1, "cell"))
                elif last[0] == "rows":
                    return self.gettags(self.set_currently_selected(r1, c1, "row"))
                elif last[0] == "columns":
                    return self.gettags(self.set_currently_selected(r1, c1, "column"))
        return tuple()

    def delete_current(self):
        self.delete("selected")
        self.RI.delete("selected")
        self.CH.delete("selected")

    def create_selected(
        self,
        r1=None,
        c1=None,
        r2=None,
        c2=None,
        type_="cells",
        taglower=True,
        state="normal",
    ):
        self.itemconfig("cells", state="normal")
        coords = f"{r1}_{c1}_{r2}_{c2}"
        if type_ == "cells":
            tagr = ("cells", coords)
            tagb = ("cellsbd", coords)
            mt_bg = self.table_selected_cells_bg
            mt_border_col = self.table_selected_cells_border_fg
        elif type_ == "rows":
            tagr = ("rows", coords)
            tagb = ("rowsbd", coords)
            tag_index_header = ("cells", coords)
            mt_bg = self.table_selected_rows_bg
            mt_border_col = self.table_selected_rows_border_fg
        elif type_ == "columns":
            tagr = ("columns", coords)
            tagb = ("columnsbd", coords)
            tag_index_header = ("cells", coords)
            mt_bg = self.table_selected_columns_bg
            mt_border_col = self.table_selected_columns_border_fg
        self.last_selected = (r1, c1, r2, c2, type_)
        ch_tags = tag_index_header if type_ == "rows" else tagr
        ri_tags = tag_index_header if type_ == "columns" else tagr
        r = self.create_rectangle(
            self.col_positions[c1],
            self.row_positions[r1],
            self.canvasx(self.winfo_width()) if self.selected_rows_to_end_of_window else self.col_positions[c2],
            self.row_positions[r2],
            fill=mt_bg,
            outline="",
            state=state,
            tags=tagr,
        )
        self.RI.create_rectangle(
            0,
            self.row_positions[r1],
            self.RI.current_width - 1,
            self.row_positions[r2],
            fill=self.RI.index_selected_rows_bg if type_ == "rows" else self.RI.index_selected_cells_bg,
            outline="",
            tags=ri_tags,
        )
        self.CH.create_rectangle(
            self.col_positions[c1],
            0,
            self.col_positions[c2],
            self.CH.current_height - 1,
            fill=self.CH.header_selected_columns_bg if type_ == "columns" else self.CH.header_selected_cells_bg,
            outline="",
            tags=ch_tags,
        )
        if self.show_selected_cells_border and (
            (self.being_drawn_rect is None and self.RI.being_drawn_rect is None and self.CH.being_drawn_rect is None)
            or len(self.anything_selected()) > 1
        ):
            b = self.create_rectangle(
                self.col_positions[c1],
                self.row_positions[r1],
                self.col_positions[c2],
                self.row_positions[r2],
                fill="",
                outline=mt_border_col,
                tags=tagb,
            )
        else:
            b = None
        if taglower:
            self.tag_lower("rows")
            self.RI.tag_lower("rows")
            self.tag_lower("columns")
            self.CH.tag_lower("columns")
            self.tag_lower("cells")
            self.RI.tag_lower("cells")
            self.CH.tag_lower("cells")
        return r, b

    def recreate_all_selection_boxes(self):
        curr = self.currently_selected()
        for item in chain(
            self.find_withtag("cells"),
            self.find_withtag("rows"),
            self.find_withtag("columns"),
        ):
            tags = self.gettags(item)
            if tags:
                r1, c1, r2, c2 = tuple(int(e) for e in tags[1].split("_") if e)
                state = self.itemcget(item, "state")
                self.delete(f"{r1}_{c1}_{r2}_{c2}")
                self.RI.delete(f"{r1}_{c1}_{r2}_{c2}")
                self.CH.delete(f"{r1}_{c1}_{r2}_{c2}")
                if r1 >= len(self.row_positions) - 1 or c1 >= len(self.col_positions) - 1:
                    continue
                if r2 > len(self.row_positions) - 1:
                    r2 = len(self.row_positions) - 1
                if c2 > len(self.col_positions) - 1:
                    c2 = len(self.col_positions) - 1
                self.create_selected(r1, c1, r2, c2, tags[0], state=state)
        if curr:
            self.set_currently_selected(curr.row, curr.column, curr.type_)
        self.tag_lower("rows")
        self.RI.tag_lower("rows")
        self.tag_lower("columns")
        self.CH.tag_lower("columns")
        self.tag_lower("cells")
        self.RI.tag_lower("cells")
        self.CH.tag_lower("cells")
        if not self.show_selected_cells_border:
            self.tag_lower("selected")

    def get_redraw_selections(self, startr, endr, startc, endc):
        d = defaultdict(list)
        for item in chain(
            self.find_withtag("cells"),
            self.find_withtag("rows"),
            self.find_withtag("columns"),
        ):
            tags = self.gettags(item)
            d[tags[0]].append(tuple(int(e) for e in tags[1].split("_") if e))
        d2 = {}
        if "cells" in d:
            d2["cells"] = {
                (r, c)
                for r in range(startr, endr)
                for c in range(startc, endc)
                for r1, c1, r2, c2 in d["cells"]
                if r1 <= r and c1 <= c and r2 > r and c2 > c
            }
        if "rows" in d:
            d2["rows"] = {r for r in range(startr, endr) for r1, c1, r2, c2 in d["rows"] if r1 <= r and r2 > r}
        if "columns" in d:
            d2["columns"] = {c for c in range(startc, endc) for r1, c1, r2, c2 in d["columns"] if c1 <= c and c2 > c}
        return d2

    def get_selected_min_max(self):
        min_x = float("inf")
        min_y = float("inf")
        max_x = 0
        max_y = 0
        for item in chain(
            self.find_withtag("cells"),
            self.find_withtag("rows"),
            self.find_withtag("columns"),
            self.find_withtag("selected"),
        ):
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

    def get_selected_rows(self, get_cells=False, within_range=None, get_cells_as_rows=False):
        s = set()
        if within_range is not None:
            within_r1 = within_range[0]
            within_r2 = within_range[1]
        if get_cells:
            if within_range is None:
                for item in self.find_withtag("rows"):
                    r1, c1, r2, c2 = tuple(int(e) for e in self.gettags(item)[1].split("_") if e)
                    s.update(set(product(range(r1, r2), range(0, len(self.col_positions) - 1))))
                if get_cells_as_rows:
                    s.update(self.get_selected_cells())
            else:
                for item in self.find_withtag("rows"):
                    r1, c1, r2, c2 = tuple(int(e) for e in self.gettags(item)[1].split("_") if e)
                    if r1 >= within_r1 or r2 <= within_r2:
                        if r1 > within_r1:
                            start_row = r1
                        else:
                            start_row = within_r1
                        if r2 < within_r2:
                            end_row = r2
                        else:
                            end_row = within_r2
                        s.update(
                            set(
                                product(
                                    range(start_row, end_row),
                                    range(0, len(self.col_positions) - 1),
                                )
                            )
                        )
                if get_cells_as_rows:
                    s.update(
                        self.get_selected_cells(
                            within_range=(
                                within_r1,
                                0,
                                within_r2,
                                len(self.col_positions) - 1,
                            )
                        )
                    )
        else:
            if within_range is None:
                for item in self.find_withtag("rows"):
                    r1, c1, r2, c2 = tuple(int(e) for e in self.gettags(item)[1].split("_") if e)
                    s.update(set(range(r1, r2)))
                if get_cells_as_rows:
                    s.update(set(tup[0] for tup in self.get_selected_cells()))
            else:
                for item in self.find_withtag("rows"):
                    r1, c1, r2, c2 = tuple(int(e) for e in self.gettags(item)[1].split("_") if e)
                    if r1 >= within_r1 or r2 <= within_r2:
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
                    s.update(
                        set(
                            tup[0]
                            for tup in self.get_selected_cells(
                                within_range=(
                                    within_r1,
                                    0,
                                    within_r2,
                                    len(self.col_positions) - 1,
                                )
                            )
                        )
                    )
        return s

    def get_selected_cols(self, get_cells=False, within_range=None, get_cells_as_cols=False):
        s = set()
        if within_range is not None:
            within_c1 = within_range[0]
            within_c2 = within_range[1]
        if get_cells:
            if within_range is None:
                for item in self.find_withtag("columns"):
                    r1, c1, r2, c2 = tuple(int(e) for e in self.gettags(item)[1].split("_") if e)
                    s.update(set(product(range(c1, c2), range(0, len(self.row_positions) - 1))))
                if get_cells_as_cols:
                    s.update(self.get_selected_cells())
            else:
                for item in self.find_withtag("columns"):
                    r1, c1, r2, c2 = tuple(int(e) for e in self.gettags(item)[1].split("_") if e)
                    if c1 >= within_c1 or c2 <= within_c2:
                        if c1 > within_c1:
                            start_col = c1
                        else:
                            start_col = within_c1
                        if c2 < within_c2:
                            end_col = c2
                        else:
                            end_col = within_c2
                        s.update(
                            set(
                                product(
                                    range(start_col, end_col),
                                    range(0, len(self.row_positions) - 1),
                                )
                            )
                        )
                if get_cells_as_cols:
                    s.update(
                        self.get_selected_cells(
                            within_range=(
                                0,
                                within_c1,
                                len(self.row_positions) - 1,
                                within_c2,
                            )
                        )
                    )
        else:
            if within_range is None:
                for item in self.find_withtag("columns"):
                    r1, c1, r2, c2 = tuple(int(e) for e in self.gettags(item)[1].split("_") if e)
                    s.update(set(range(c1, c2)))
                if get_cells_as_cols:
                    s.update(set(tup[1] for tup in self.get_selected_cells()))
            else:
                for item in self.find_withtag("columns"):
                    r1, c1, r2, c2 = tuple(int(e) for e in self.gettags(item)[1].split("_") if e)
                    if c1 >= within_c1 or c2 <= within_c2:
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
                    s.update(
                        set(
                            tup[0]
                            for tup in self.get_selected_cells(
                                within_range=(
                                    0,
                                    within_c1,
                                    len(self.row_positions) - 1,
                                    within_c2,
                                )
                            )
                        )
                    )
        return s

    def get_selected_cells(self, get_rows=False, get_cols=False, within_range=None):
        s = set()
        if within_range is not None:
            within_r1 = within_range[0]
            within_c1 = within_range[1]
            within_r2 = within_range[2]
            within_c2 = within_range[3]
        if get_cols and get_rows:
            iterable = chain(
                self.find_withtag("cells"),
                self.find_withtag("rows"),
                self.find_withtag("columns"),
            )
        elif get_rows and not get_cols:
            iterable = chain(self.find_withtag("cells"), self.find_withtag("rows"))
        elif get_cols and not get_rows:
            iterable = chain(self.find_withtag("cells"), self.find_withtag("columns"))
        else:
            iterable = chain(self.find_withtag("cells"))
        if within_range is None:
            for item in iterable:
                r1, c1, r2, c2 = tuple(int(e) for e in self.gettags(item)[1].split("_") if e)
                s.update(set(product(range(r1, r2), range(c1, c2))))
        else:
            for item in iterable:
                r1, c1, r2, c2 = tuple(int(e) for e in self.gettags(item)[1].split("_") if e)
                if r1 >= within_r1 or c1 >= within_c1 or r2 <= within_r2 or c2 <= within_c2:
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
        return tuple(
            tuple(int(e) for e in self.gettags(item)[1].split("_") if e)
            for item in chain(
                self.find_withtag("cells"),
                self.find_withtag("rows"),
                self.find_withtag("columns"),
            )
        )

    def get_all_selection_boxes_with_types(self):
        boxes = []
        for item in sorted(self.find_withtag("cells") + self.find_withtag("rows") + self.find_withtag("columns")):
            tags = self.gettags(item)
            boxes.append((tuple(int(e) for e in tags[1].split("_") if e), tags[0]))
        return boxes

    def all_selected(self):
        for r1, c1, r2, c2 in self.get_all_selection_boxes():
            if not r1 and not c1 and r2 == len(self.row_positions) - 1 and c2 == len(self.col_positions) - 1:
                return True
        return False

    # don't have to use "selected" because you can't have a current without a selection box
    def cell_selected(self, r, c, inc_cols=False, inc_rows=False):
        if not isinstance(r, int) or not isinstance(c, int):
            return False
        if not inc_cols and not inc_rows:
            iterable = self.find_withtag("cells")
        elif inc_cols and not inc_rows:
            iterable = chain(self.find_withtag("columns"), self.find_withtag("cells"))
        elif not inc_cols and inc_rows:
            iterable = chain(self.find_withtag("rows"), self.find_withtag("cells"))
        elif inc_cols and inc_rows:
            iterable = chain(
                self.find_withtag("rows"),
                self.find_withtag("columns"),
                self.find_withtag("cells"),
            )
        for item in iterable:
            r1, c1, r2, c2 = tuple(int(e) for e in self.gettags(item)[1].split("_") if e)
            if r1 <= r and c1 <= c and r2 > r and c2 > c:
                return True
        return False

    def col_selected(self, c):
        if not isinstance(c, int):
            return False
        for item in self.find_withtag("columns"):
            r1, c1, r2, c2 = tuple(int(e) for e in self.gettags(item)[1].split("_") if e)
            if c1 <= c and c2 > c:
                return True
        return False

    def row_selected(self, r):
        if not isinstance(r, int):
            return False
        for item in self.find_withtag("rows"):
            r1, c1, r2, c2 = tuple(int(e) for e in self.gettags(item)[1].split("_") if e)
            if r1 <= r and r2 > r:
                return True
        return False

    def anything_selected(self, exclude_columns=False, exclude_rows=False, exclude_cells=False):
        if exclude_columns and exclude_rows and not exclude_cells:
            return self.find_withtag("cells")
        elif exclude_columns and exclude_cells and not exclude_rows:
            return self.find_withtag("rows")
        elif exclude_rows and exclude_cells and not exclude_columns:
            return self.find_withtag("columns")

        elif exclude_columns and not exclude_rows and not exclude_cells:
            return self.find_withtag("cells") + self.find_withtag("rows")
        elif exclude_rows and not exclude_columns and not exclude_cells:
            return self.find_withtag("cells") + self.find_withtag("columns")

        elif exclude_cells and not exclude_columns and not exclude_rows:
            return self.find_withtag("rows") + self.find_withtag("columns")

        elif not exclude_columns and not exclude_rows and not exclude_cells:
            return self.find_withtag("cells") + self.find_withtag("rows") + self.find_withtag("columns")
        return tuple()

    def hide_current(self):
        for item in self.find_withtag("selected"):
            self.itemconfig(item, state="hidden")

    def show_current(self):
        for item in self.find_withtag("selected"):
            self.itemconfig(item, state="normal")

    def open_cell(self, event=None, ignore_existing_editor=False):
        if not self.anything_selected() or (not ignore_existing_editor and self.text_editor_id is not None):
            return
        currently_selected = self.currently_selected()
        if not currently_selected:
            return
        r, c = int(currently_selected[0]), int(currently_selected[1])
        datacn = c if self.all_columns_displayed else self.displayed_columns[c]
        datarn = r if self.all_rows_displayed else self.displayed_rows[r]
        if self.get_cell_kwargs(datarn, datacn, key="readonly"):
            return
        elif self.get_cell_kwargs(datarn, datacn, key="dropdown") or self.get_cell_kwargs(
            datarn, datacn, key="checkbox"
        ):
            if self.event_opens_dropdown_or_checkbox(event):
                if self.get_cell_kwargs(datarn, datacn, key="dropdown"):
                    self.open_dropdown_window(r, c, event=event)
                elif self.get_cell_kwargs(datarn, datacn, key="checkbox"):
                    self.click_checkbox(r=r, c=c, datarn=datarn, datacn=datacn)
        else:
            self.open_text_editor(event=event, r=r, c=c, dropdown=False)

    def event_opens_dropdown_or_checkbox(self, event=None):
        if event is None:
            return False
        elif event == "rc":
            return True
        elif (
            (hasattr(event, "keysym") and event.keysym == "Return")
            or (hasattr(event, "keysym") and event.keysym == "F2")  # enter or f2
            or (
                event is not None
                and hasattr(event, "keycode")
                and event.keycode == "??"
                and hasattr(event, "num")
                and event.num == 1
            )
            or (hasattr(event, "keysym") and event.keysym == "BackSpace")
        ):
            return True
        else:
            return False

    # displayed indexes
    def get_cell_align(self, r, c):
        datarn = r if self.all_rows_displayed else self.displayed_rows[r]
        datacn = c if self.all_columns_displayed else self.displayed_columns[c]
        cell_alignment = self.get_cell_kwargs(datarn, datacn, key="align")
        if cell_alignment:
            return cell_alignment
        return self.align

    # displayed indexes
    def open_text_editor(
        self,
        event=None,
        r=0,
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
        if event is None or self.event_opens_dropdown_or_checkbox(event):
            if event is not None:
                if hasattr(event, "keysym") and event.keysym == "Return":
                    extra_func_key = "Return"
                elif hasattr(event, "keysym") and event.keysym == "F2":
                    extra_func_key = "F2"
            datarn = r if self.all_rows_displayed else self.displayed_rows[r]
            datacn = c if self.all_columns_displayed else self.displayed_columns[c]
            if event is not None and (hasattr(event, "keysym") and event.keysym == "BackSpace"):
                extra_func_key = "BackSpace"
                text = ""
            else:
                text = f"{self.get_cell_data(datarn, datacn, none_to_empty_str = True)}"
        elif event is not None and (
            (hasattr(event, "char") and event.char.isalpha())
            or (hasattr(event, "char") and event.char.isdigit())
            or (hasattr(event, "char") and event.char in symbols_set)
        ):
            extra_func_key = event.char
            text = event.char
        else:
            return False
        self.text_editor_loc = (r, c)
        if self.extra_begin_edit_cell_func is not None:
            try:
                text = self.extra_begin_edit_cell_func(EditCellEvent(r, c, extra_func_key, text, "begin_edit_cell"))
            except Exception:
                return False
            if text is None:
                return False
            else:
                text = text if isinstance(text, str) else f"{text}"
        text = "" if text is None else text
        if self.cell_auto_resize_enabled:
            self.set_cell_size_to_text(r, c, only_set_if_too_small=True, redraw=True, run_binding=True)

        if (r, c) == self.text_editor_loc and self.text_editor is not None:
            self.text_editor.set_text(self.text_editor.get() + "" if not isinstance(text, str) else text)
            return
        if self.text_editor is not None:
            self.destroy_text_editor()
        if see:
            has_redrawn = self.see(r=r, c=c, check_cell_visibility=True)
            if not has_redrawn:
                self.refresh()
        self.text_editor_loc = (r, c)
        x = self.col_positions[c]
        y = self.row_positions[r]
        w = self.col_positions[c + 1] - x + 1
        h = self.row_positions[r + 1] - y + 1
        if text is None:
            text = f"""{self.get_cell_data(r if self.all_rows_displayed else self.displayed_rows[r],
                                           c if self.all_columns_displayed else self.displayed_columns[c],
                                           none_to_empty_str = True)}"""
        self.hide_current()
        bg, fg = self.table_bg, self.table_fg
        self.text_editor = TextEditor(
            self,
            text=text,
            font=self.table_font,
            state=state,
            width=w,
            height=h,
            border_color=self.table_selected_cells_border_fg,
            show_border=self.show_selected_cells_border,
            bg=bg,
            fg=fg,
            popup_menu_font=self.popup_menu_font,
            popup_menu_fg=self.popup_menu_fg,
            popup_menu_bg=self.popup_menu_bg,
            popup_menu_highlight_bg=self.popup_menu_highlight_bg,
            popup_menu_highlight_fg=self.popup_menu_highlight_fg,
            binding=binding,
            align=self.get_cell_align(r, c),
            r=r,
            c=c,
            newline_binding=self.text_editor_newline_binding,
        )
        self.text_editor.update_idletasks()
        self.text_editor_id = self.create_window((x, y), window=self.text_editor, anchor="nw")
        if not dropdown:
            self.text_editor.textedit.focus_set()
            self.text_editor.scroll_to_bottom()
        self.text_editor.textedit.bind("<Alt-Return>", lambda x: self.text_editor_newline_binding(r, c))
        if USER_OS == "darwin":
            self.text_editor.textedit.bind("<Option-Return>", lambda x: self.text_editor_newline_binding(r, c))
        for key, func in self.text_editor_user_bound_keys.items():
            self.text_editor.textedit.bind(key, func)
        if binding is not None:
            self.text_editor.textedit.bind("<Tab>", lambda x: binding((r, c, "Tab")))
            self.text_editor.textedit.bind("<Return>", lambda x: binding((r, c, "Return")))
            self.text_editor.textedit.bind("<FocusOut>", lambda x: binding((r, c, "FocusOut")))
            self.text_editor.textedit.bind("<Escape>", lambda x: binding((r, c, "Escape")))
        elif binding is None and set_data_on_close:
            self.text_editor.textedit.bind("<Tab>", lambda x: self.close_text_editor((r, c, "Tab")))
            self.text_editor.textedit.bind("<Return>", lambda x: self.close_text_editor((r, c, "Return")))
            if not dropdown:
                self.text_editor.textedit.bind("<FocusOut>", lambda x: self.close_text_editor((r, c, "FocusOut")))
            self.text_editor.textedit.bind("<Escape>", lambda x: self.close_text_editor((r, c, "Escape")))
        else:
            self.text_editor.textedit.bind("<Escape>", lambda x: self.destroy_text_editor("Escape"))
        return True

    # displayed indexes
    def text_editor_newline_binding(self, r=0, c=0, event=None, check_lines=True):
        datarn = r if self.all_rows_displayed else self.displayed_rows[r]
        datacn = c if self.all_columns_displayed else self.displayed_columns[c]
        curr_height = self.text_editor.winfo_height()
        if not check_lines or self.get_lines_cell_height(self.text_editor.get_num_lines() + 1) > curr_height:
            new_height = curr_height + self.xtra_lines_increment
            space_bot = self.get_space_bot(r)
            if new_height > space_bot:
                new_height = space_bot
            if new_height != curr_height:
                self.text_editor.config(height=new_height)
                kwargs = self.get_cell_kwargs(datarn, datacn, key="dropdown")
                if kwargs:
                    text_editor_h = self.text_editor.winfo_height()
                    win_h, anchor = self.get_dropdown_height_anchor(datarn, datacn, text_editor_h)
                    if anchor == "nw":
                        self.coords(
                            kwargs["canvas_id"],
                            self.col_positions[c],
                            self.row_positions[r] + text_editor_h - 1,
                        )
                        self.itemconfig(kwargs["canvas_id"], anchor=anchor, height=win_h)
                    elif anchor == "sw":
                        self.coords(
                            kwargs["canvas_id"],
                            self.col_positions[c],
                            self.row_positions[r],
                        )
                        self.itemconfig(kwargs["canvas_id"], anchor=anchor, height=win_h)

    def destroy_text_editor(self, event=None):
        if event is not None and self.extra_end_edit_cell_func is not None and self.text_editor_loc is not None:
            self.extra_end_edit_cell_func(
                EditCellEvent(
                    int(self.text_editor_loc[0]),
                    int(self.text_editor_loc[1]),
                    "Escape",
                    None,
                    "escape_edit_cell",
                )
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
        self.show_current()
        if event is not None and len(event) >= 3 and "Escape" in event:
            self.focus_set()

    # c is displayed col
    def close_text_editor(
        self,
        editor_info=None,
        r=None,
        c=None,
        set_data_on_close=True,
        event=None,
        destroy=True,
        move_down=True,
        redraw=True,
        recreate=True,
    ):
        if self.focus_get() is None and editor_info:
            return "break"
        if editor_info is not None and len(editor_info) >= 3 and editor_info[2] == "Escape":
            self.destroy_text_editor("Escape")
            self.close_dropdown_window(r, c)
            return "break"
        if self.text_editor is not None:
            self.text_editor_value = self.text_editor.get()
        if destroy:
            self.destroy_text_editor()
        if set_data_on_close:
            if r is None and c is None and editor_info:
                r, c = editor_info[0], editor_info[1]
            datarn = r if self.all_rows_displayed else self.displayed_rows[r]
            datacn = c if self.all_columns_displayed else self.displayed_columns[c]
            if self.extra_end_edit_cell_func is None and self.input_valid_for_cell(
                datarn, datacn, self.text_editor_value
            ):
                self.set_cell_data_undo(
                    r,
                    c,
                    datarn=datarn,
                    datacn=datacn,
                    value=self.text_editor_value,
                    redraw=False,
                    check_input_valid=False,
                )
            elif (
                self.extra_end_edit_cell_func is not None
                and not self.edit_cell_validation
                and self.input_valid_for_cell(datarn, datacn, self.text_editor_value)
            ):
                self.set_cell_data_undo(
                    r,
                    c,
                    datarn=datarn,
                    datacn=datacn,
                    value=self.text_editor_value,
                    redraw=False,
                    check_input_valid=False,
                )
                self.extra_end_edit_cell_func(
                    EditCellEvent(
                        r,
                        c,
                        editor_info[2] if len(editor_info) >= 3 else "FocusOut",
                        f"{self.text_editor_value}",
                        "end_edit_cell",
                    )
                )
            elif self.extra_end_edit_cell_func is not None and self.edit_cell_validation:
                validation = self.extra_end_edit_cell_func(
                    EditCellEvent(
                        r,
                        c,
                        editor_info[2] if len(editor_info) >= 3 else "FocusOut",
                        f"{self.text_editor_value}",
                        "end_edit_cell",
                    )
                )
                self.text_editor_value = validation
                if validation is not None and self.input_valid_for_cell(datarn, datacn, self.text_editor_value):
                    self.set_cell_data_undo(
                        r,
                        c,
                        datarn=datarn,
                        datacn=datacn,
                        value=self.text_editor_value,
                        redraw=False,
                        check_input_valid=False,
                    )
        if move_down:
            if r is None and c is None and editor_info:
                r, c = editor_info[0], editor_info[1]
            currently_selected = self.currently_selected()
            if (
                r is not None
                and c is not None
                and currently_selected
                and r == currently_selected[0]
                and c == currently_selected[1]
                and (self.single_selection_enabled or self.toggle_selection_enabled)
            ):
                r1, c1, r2, c2 = self.find_last_selected_box_with_current(currently_selected)
                numcols = c2 - c1
                numrows = r2 - r1
                if numcols == 1 and numrows == 1:
                    if editor_info is not None and len(editor_info) >= 3 and editor_info[2] == "Return":
                        self.select_cell(r + 1 if r < len(self.row_positions) - 2 else r, c)
                        self.see(
                            r + 1 if r < len(self.row_positions) - 2 else r,
                            c,
                            keep_xscroll=True,
                            bottom_right_corner=True,
                            check_cell_visibility=True,
                        )
                    elif editor_info is not None and len(editor_info) >= 3 and editor_info[2] == "Tab":
                        self.select_cell(r, c + 1 if c < len(self.col_positions) - 2 else c)
                        self.see(
                            r,
                            c + 1 if c < len(self.col_positions) - 2 else c,
                            keep_xscroll=True,
                            bottom_right_corner=True,
                            check_cell_visibility=True,
                        )
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
                    self.set_currently_selected(new_r, new_c, type_=currently_selected.type_)
                    self.see(
                        new_r,
                        new_c,
                        keep_xscroll=False,
                        bottom_right_corner=True,
                        check_cell_visibility=True,
                    )
        self.close_dropdown_window(r, c)
        if recreate:
            self.recreate_all_selection_boxes()
        if redraw:
            self.refresh()
        if editor_info is not None and len(editor_info) >= 3 and editor_info[2] != "FocusOut":
            self.focus_set()
        return "break"

    def tab_key(self, event=None):
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
        self.set_currently_selected(new_r, new_c, type_=currently_selected.type_)
        self.see(
            new_r,
            new_c,
            keep_xscroll=False,
            bottom_right_corner=True,
            check_cell_visibility=True,
        )
        return "break"

    # internal event use
    def set_cell_data_undo(
        self,
        r=0,
        c=0,
        datarn=None,
        datacn=None,
        value="",
        undo=True,
        cell_resize=True,
        redraw=True,
        check_input_valid=True,
    ):
        if datacn is None:
            datacn = c if self.all_columns_displayed else self.displayed_columns[c]
        if datarn is None:
            datarn = r if self.all_rows_displayed else self.displayed_rows[r]
        if not check_input_valid or self.input_valid_for_cell(datarn, datacn, value):
            if self.undo_enabled and undo:
                self.undo_storage.append(
                    zlib.compress(
                        pickle.dumps(
                            (
                                "edit_cells",
                                {(datarn, datacn): self.get_cell_data(datarn, datacn)},
                                self.get_boxes(include_current=False),
                                self.currently_selected(),
                            )
                        )
                    )
                )
            self.set_cell_data(datarn, datacn, value)
        if cell_resize and self.cell_auto_resize_enabled:
            self.set_cell_size_to_text(r, c, only_set_if_too_small=True, redraw=redraw, run_binding=True)
        self.parentframe.emit_event("<<SheetModified>>")
        return True

    def set_cell_data(self, datarn, datacn, value, kwargs={}, expand_sheet=True):
        if expand_sheet:
            if datarn >= len(self.data):
                self.fix_data_len(datarn, datacn)
            elif datacn >= len(self.data[datarn]):
                self.fix_row_len(datarn, datacn)
        if expand_sheet or (len(self.data) > datarn and len(self.data[datarn]) > datacn):
            if (
                datarn,
                datacn,
            ) in self.cell_options and "checkbox" in self.cell_options[(datarn, datacn)]:
                self.data[datarn][datacn] = try_to_bool(value)
            else:
                if not kwargs:
                    kwargs = self.get_cell_kwargs(datarn, datacn, key="format")
                if kwargs:
                    if kwargs["formatter"] is None:
                        self.data[datarn][datacn] = format_data(value=value, **kwargs)
                    else:
                        self.data[datarn][datacn] = kwargs["formatter"](value, **kwargs)
                else:
                    self.data[datarn][datacn] = value

    def get_value_for_empty_cell(self, datarn, datacn, r_ops=True, c_ops=True):
        if self.get_cell_kwargs(
            datarn,
            datacn,
            key="checkbox",
            cell=r_ops and c_ops,
            row=r_ops,
            column=c_ops,
        ):
            return False
        kwargs = self.get_cell_kwargs(
            datarn,
            datacn,
            key="dropdown",
            cell=r_ops and c_ops,
            row=r_ops,
            column=c_ops,
        )
        if kwargs and kwargs["validate_input"] and kwargs["values"]:
            return kwargs["values"][0]
        return ""

    def get_empty_row_seq(self, datarn, end, start=0, r_ops=True, c_ops=True):
        return [self.get_value_for_empty_cell(datarn, datacn, r_ops=r_ops, c_ops=c_ops) for datacn in range(start, end)]

    def fix_row_len(self, datarn, datacn):
        self.data[datarn].extend(self.get_empty_row_seq(datarn, end=datacn + 1, start=len(self.data[datarn])))

    def fix_row_values(self, datarn, start=None, end=None):
        if datarn < len(self.data):
            for datacn, v in enumerate(islice(self.data[datarn], start, end)):
                if not self.input_valid_for_cell(datarn, datacn, v):
                    self.data[datarn][datacn] = self.get_value_for_empty_cell(datarn, datacn)

    def fix_data_len(self, datarn, datacn):
        ncols = self.total_data_cols() if datacn is None else datacn + 1
        self.data.extend([self.get_empty_row_seq(rn, end=ncols, start=0) for rn in range(len(self.data), datarn + 1)])

    # internal event use
    def click_checkbox(self, r, c, datarn=None, datacn=None, undo=True, redraw=True):
        if datarn is None:
            datarn = r if self.all_rows_displayed else self.displayed_rows[r]
        if datacn is None:
            datacn = c if self.all_columns_displayed else self.displayed_columns[c]
        kwargs = self.get_cell_kwargs(datarn, datacn, key="checkbox")
        if kwargs["state"] == "normal":
            self.set_cell_data_undo(
                r,
                c,
                value=not self.data[datarn][datacn] if type(self.data[datarn][datacn]) == bool else False,
                undo=undo,
                cell_resize=False,
                check_input_valid=False,
            )
            if kwargs["check_function"] is not None:
                kwargs["check_function"]((r, c, "CheckboxClicked", self.data[datarn][datacn]))
            if self.extra_end_edit_cell_func is not None:
                self.extra_end_edit_cell_func(EditCellEvent(r, c, "Return", self.data[datarn][datacn], "end_edit_cell"))
        if redraw:
            self.refresh()

    def create_checkbox(self, datarn=0, datacn=0, **kwargs):
        self.delete_cell_format(datarn, datacn, clear_values=False)
        if (datarn, datacn) in self.cell_options and (
            "dropdown" in self.cell_options[(datarn, datacn)] or "checkbox" in self.cell_options[(datarn, datacn)]
        ):
            self.delete_cell_options_dropdown_and_checkbox(datarn, datacn)
        if (datarn, datacn) not in self.cell_options:
            self.cell_options[(datarn, datacn)] = {}
        self.cell_options[(datarn, datacn)]["checkbox"] = get_checkbox_dict(**kwargs)
        self.set_cell_data(datarn, datacn, kwargs["checked"])

    def checkbox_row(self, datarn=0, **kwargs):
        self.delete_row_format(datarn, clear_values=False)
        if datarn in self.row_options and (
            "dropdown" in self.row_options[datarn] or "checkbox" in self.row_options[datarn]
        ):
            self.delete_row_options_dropdown_and_checkbox(datarn)
        if datarn not in self.row_options:
            self.row_options[datarn] = {}
        self.row_options[datarn]["checkbox"] = get_checkbox_dict(**kwargs)
        for datacn in range(self.total_data_cols()):
            self.set_cell_data(datarn, datacn, kwargs["checked"])

    def checkbox_column(self, datacn=0, **kwargs):
        self.delete_column_format(datacn, clear_values=False)
        if datacn in self.col_options and (
            "dropdown" in self.col_options[datacn] or "checkbox" in self.col_options[datacn]
        ):
            self.delete_column_options_dropdown_and_checkbox(datacn)
        if datacn not in self.col_options:
            self.col_options[datacn] = {}
        self.col_options[datacn]["checkbox"] = get_checkbox_dict(**kwargs)
        for datarn in range(self.total_data_rows()):
            self.set_cell_data(datarn, datacn, kwargs["checked"])

    def checkbox_sheet(self, **kwargs):
        self.delete_sheet_format(clear_values=False)
        if "dropdown" in self.options or "checkbox" in self.options:
            self.delete_options_dropdown_and_checkbox()
        self.options["checkbox"] = get_checkbox_dict(**kwargs)
        total_cols = self.total_data_cols()
        for datarn in range(self.total_data_rows()):
            for datacn in range(total_cols):
                self.set_cell_data(datarn, datacn, kwargs["checked"])

    def create_dropdown(self, datarn=0, datacn=0, **kwargs):
        if (datarn, datacn) in self.cell_options and (
            "dropdown" in self.cell_options[(datarn, datacn)] or "checkbox" in self.cell_options[(datarn, datacn)]
        ):
            self.delete_cell_options_dropdown_and_checkbox(datarn, datacn)
        if (datarn, datacn) not in self.cell_options:
            self.cell_options[(datarn, datacn)] = {}
        self.cell_options[(datarn, datacn)]["dropdown"] = get_dropdown_dict(**kwargs)
        self.set_cell_data(
            datarn,
            datacn,
            kwargs["set_value"] if kwargs["set_value"] is not None else kwargs["values"][0] if kwargs["values"] else "",
        )

    def dropdown_row(self, datarn=0, **kwargs):
        if datarn in self.row_options and (
            "dropdown" in self.row_options[datarn] or "checkbox" in self.row_options[datarn]
        ):
            self.delete_row_options_dropdown_and_checkbox(datarn)
        if datarn not in self.row_options:
            self.row_options[datarn] = {}
        self.row_options[datarn]["dropdown"] = get_dropdown_dict(**kwargs)
        value = (
            kwargs["set_value"] if kwargs["set_value"] is not None else kwargs["values"][0] if kwargs["values"] else ""
        )
        for datacn in range(self.total_data_cols()):
            self.set_cell_data(datarn, datacn, value)

    def dropdown_column(self, datacn=0, **kwargs):
        if datacn in self.col_options and (
            "dropdown" in self.col_options[datacn] or "checkbox" in self.col_options[datacn]
        ):
            self.delete_column_options_dropdown_and_checkbox(datacn)
        if datacn not in self.col_options:
            self.col_options[datacn] = {}
        self.col_options[datacn]["dropdown"] = get_dropdown_dict(**kwargs)
        value = (
            kwargs["set_value"] if kwargs["set_value"] is not None else kwargs["values"][0] if kwargs["values"] else ""
        )
        for datarn in range(self.total_data_rows()):
            self.set_cell_data(datarn, datacn, value)

    def dropdown_sheet(self, **kwargs):
        if "dropdown" in self.options or "checkbox" in self.options:
            self.delete_options_dropdown_and_checkbox()
        self.options["dropdown"] = get_dropdown_dict(**kwargs)
        value = (
            kwargs["set_value"] if kwargs["set_value"] is not None else kwargs["values"][0] if kwargs["values"] else ""
        )
        total_cols = self.total_data_cols()
        for datarn in range(self.total_data_rows()):
            for datacn in range(total_cols):
                self.set_cell_data(datarn, datacn, value)

    def format_cell(self, datarn, datacn, **kwargs):
        if (datarn, datacn) in self.cell_options and "checkbox" in self.cell_options[(datarn, datacn)]:
            return
        kwargs = self.format_fix_kwargs(kwargs)
        if (datarn, datacn) not in self.cell_options:
            self.cell_options[(datarn, datacn)] = {}
        self.cell_options[(datarn, datacn)]["format"] = kwargs
        self.set_cell_data(
            datarn,
            datacn,
            value=kwargs["value"] if "value" in kwargs else self.get_cell_data(datarn, datacn),
            kwargs=kwargs,
        )

    def format_row(self, datarn, **kwargs):
        if datarn in self.row_options and "checkbox" in self.row_options[datarn]:
            return
        kwargs = self.format_fix_kwargs(kwargs)
        if datarn not in self.row_options:
            self.row_options[datarn] = {}
        self.row_options[datarn]["format"] = kwargs
        for datacn in range(self.total_data_cols()):
            self.set_cell_data(
                datarn,
                datacn,
                value=kwargs["value"] if "value" in kwargs else self.get_cell_data(datarn, datacn),
                kwargs=kwargs,
            )

    def format_column(self, datacn, **kwargs):
        if datacn in self.col_options and "checkbox" in self.col_options[datacn]:
            return
        kwargs = self.format_fix_kwargs(kwargs)
        if datacn not in self.col_options:
            self.col_options[datacn] = {}
        self.col_options[datacn]["format"] = kwargs
        for datarn in range(self.total_data_rows()):
            self.set_cell_data(
                datarn,
                datacn,
                value=kwargs["value"] if "value" in kwargs else self.get_cell_data(datarn, datacn),
                kwargs=kwargs,
            )

    def format_sheet(self, **kwargs):
        kwargs = self.format_fix_kwargs(kwargs)
        self.options["format"] = kwargs
        for datarn in range(self.total_data_rows()):
            for datacn in range(self.total_data_cols()):
                self.set_cell_data(
                    datarn,
                    datacn,
                    value=kwargs["value"] if "value" in kwargs else self.get_cell_data(datarn, datacn),
                    kwargs=kwargs,
                )

    def format_fix_kwargs(self, kwargs):
        if kwargs["formatter"] is None:
            if kwargs["nullable"]:
                if isinstance(kwargs["datatypes"], (list, tuple)):
                    kwargs["datatypes"] = tuple(kwargs["datatypes"]) + (type(None),)
                else:
                    kwargs["datatypes"] = (kwargs["datatypes"], type(None))
            elif (isinstance(kwargs["datatypes"], (list, tuple)) and type(None) in kwargs["datatypes"]) or kwargs[
                "datatypes"
            ] is type(None):
                raise TypeError("Non-nullable cells cannot have NoneType as a datatype.")
        if not isinstance(kwargs["invalid_value"], str):
            kwargs["invalid_value"] = f"{kwargs['invalid_value']}"
        return kwargs

    def reapply_formatting(self):
        if "format" in self.options:
            for r in range(len(self.data)):
                if r not in self.row_options:
                    for c in range(len(self.data[r])):
                        if not (
                            (r, c) in self.cell_options
                            and "format" in self.cell_options[(r, c)]
                            or c in self.col_options
                            and "format" in self.col_options[c]
                        ):
                            self.set_cell_data(r, c, value=self.data[r][c])
        for c in self.yield_formatted_columns():
            for r in range(len(self.data)):
                if not (
                    (r, c) in self.cell_options
                    and "format" in self.cell_options[(r, c)]
                    or r in self.row_options
                    and "format" in self.row_options[r]
                ):
                    self.set_cell_data(r, c, value=self.data[r][c])
        for r in self.yield_formatted_rows():
            for c in range(len(self.data[r])):
                if not ((r, c) in self.cell_options and "format" in self.cell_options[(r, c)]):
                    self.set_cell_data(r, c, value=self.data[r][c])
        for r, c in self.yield_formatted_cells():
            if len(self.data) > r and len(self.data[r]) > c:
                self.set_cell_data(r, c, value=self.data[r][c])

    def delete_all_formatting(self, clear_values=False):
        self.delete_cell_format("all", clear_values=clear_values)
        self.delete_row_format("all", clear_values=clear_values)
        self.delete_column_format("all", clear_values=clear_values)
        self.delete_sheet_format(clear_values=clear_values)

    def delete_cell_format(self, datarn="all", datacn=0, clear_values=False):
        if isinstance(datarn, str) and datarn.lower() == "all":
            for datarn, datacn in self.yield_formatted_cells():
                del self.cell_options[(datarn, datacn)]["format"]
                if clear_values:
                    self.set_cell_data(datarn, datacn, "", expand_sheet=False)
        else:
            if (datarn, datacn) in self.cell_options and "format" in self.cell_options[(datarn, datacn)]:
                del self.cell_options[(datarn, datacn)]["format"]
                if clear_values:
                    self.set_cell_data(datarn, datacn, "", expand_sheet=False)

    def delete_row_format(self, datarn="all", clear_values=False):
        if isinstance(datarn, str) and datarn.lower() == "all":
            for datarn in self.yield_formatted_rows():
                del self.row_options[datarn]["format"]
                if clear_values:
                    for datacn in range(len(self.data[datarn])):
                        self.set_cell_data(datarn, datacn, "", expand_sheet=False)
        else:
            if datarn in self.row_options and "format" in self.row_options[datarn]:
                del self.row_options[datarn]["format"]
                if clear_values:
                    for datacn in range(len(self.data[datarn])):
                        self.set_cell_data(datarn, datacn, "", expand_sheet=False)

    def delete_column_format(self, datacn="all", clear_values=False):
        if isinstance(datacn, str) and datacn.lower() == "all":
            for datacn in self.yield_formatted_columns():
                del self.col_options[datacn]["format"]
                if clear_values:
                    for datarn in range(len(self.data)):
                        self.set_cell_data(datarn, datacn, "", expand_sheet=False)
        else:
            if datacn in self.col_options and "format" in self.col_options[datacn]:
                del self.col_options[datacn]["format"]
                if clear_values:
                    for datarn in range(len(self.data)):
                        self.set_cell_data(datarn, datacn, "", expand_sheet=False)

    def delete_sheet_format(self, clear_values=False):
        if "format" in self.options:
            del self.options["format"]
            if clear_values:
                total_cols = self.total_data_cols()
                self.data = [
                    [self.get_value_for_empty_cell(r, c) for c in range(total_cols)]
                    for r in range(self.total_data_rows())
                ]

    # deals with possibility of formatter class being in self.data cell
    # if cell is formatted - possibly returns invalid_value kwarg if cell value is not in datatypes kwarg
    # if get displayed is true then Nones are replaced by ""
    def get_valid_cell_data_as_str(self, datarn, datacn, get_displayed=False, **kwargs) -> str:
        if get_displayed:
            kwargs = self.get_cell_kwargs(datarn, datacn, key="dropdown")
            if kwargs and kwargs["text"] is not None:
                return f"{kwargs['text']}"
            kwargs = self.get_cell_kwargs(datarn, datacn, key="checkbox")
            if kwargs:
                return f"{kwargs['text']}"
        value = self.data[datarn][datacn] if len(self.data) > datarn and len(self.data[datarn]) > datacn else ""
        kwargs = self.get_cell_kwargs(datarn, datacn, key="format")
        if kwargs:
            if kwargs["formatter"] is None:
                if get_displayed:
                    return data_to_str(value, **kwargs)
                else:
                    return f"{get_data_with_valid_check(value, **kwargs)}"
            else:
                if get_displayed:
                    # assumed given formatter class has __str__() function
                    return f"{value}"
                else:
                    # assumed given formatter class has get_data_with_valid_check() function
                    return f"{value.get_data_with_valid_check()}"
        return "" if value is None else f"{value}"

    def get_cell_data(self, datarn, datacn, get_displayed=False, none_to_empty_str=False, **kwargs) -> Any:
        if get_displayed:
            return self.get_valid_cell_data_as_str(datarn, datacn, get_displayed=True)
        value = self.data[datarn][datacn] if len(self.data) > datarn and len(self.data[datarn]) > datacn else ""
        kwargs = self.get_cell_kwargs(datarn, datacn, key="format")
        if kwargs and kwargs["formatter"] is not None:
            value = value.value  # assumed given formatter class has value attribute
        return "" if (value is None and none_to_empty_str) else value

    def input_valid_for_cell(self, datarn, datacn, value):
        if self.get_cell_kwargs(datarn, datacn, key="readonly"):
            return False
        if self.cell_equal_to(datarn, datacn, value):
            return False
        if self.get_cell_kwargs(datarn, datacn, key="format"):
            return True
        if self.get_cell_kwargs(datarn, datacn, key="checkbox"):
            return is_bool_like(value)
        kwargs = self.get_cell_kwargs(datarn, datacn, key="dropdown")
        if kwargs and kwargs["validate_input"] and value not in kwargs["values"]:
            return False
        return True

    def cell_equal_to(self, datarn, datacn, value, **kwargs):
        v = self.get_cell_data(datarn, datacn)
        kwargs = self.get_cell_kwargs(datarn, datacn, key="format")
        if kwargs and kwargs["formatter"] is None:
            return v == format_data(value=value, **kwargs)
        # assumed if there is a formatter class in cell then it has a __eq__() function anyway
        # else if there is not a formatter class in cell and cell is not formatted
        # then compare value as is
        return v == value

    def get_cell_clipboard(self, datarn, datacn) -> Union[str, int, float, bool]:
        value = self.data[datarn][datacn] if len(self.data) > datarn and len(self.data[datarn]) > datacn else ""
        kwargs = self.get_cell_kwargs(datarn, datacn, key="format")
        if kwargs:
            if kwargs["formatter"] is None:
                return get_clipboard_data(value, **kwargs)
            else:
                # assumed given formatter class has get_clipboard_data()
                # function and it returns one of above type hints
                return value.get_clipboard_data()
        return f"{value}"

    def yield_formatted_cells(self, formatter=None):
        if formatter is None:
            yield from (
                cell
                for cell, options in self.cell_options.items()
                if "format" in options and options["format"]["formatter"] == formatter
            )
        else:
            yield from (cell for cell, options in self.cell_options.items() if "format" in options)

    def yield_formatted_rows(self, formatter=None):
        if formatter is None:
            yield from (r for r, options in self.row_options.items() if "format" in options)
        else:
            yield from (
                r
                for r, options in self.row_options.items()
                if "format" in options and options["format"]["formatter"] == formatter
            )

    def yield_formatted_columns(self, formatter=None):
        if formatter is None:
            yield from (c for c, options in self.col_options.items() if "format" in options)
        else:
            yield from (
                c
                for c, options in self.col_options.items()
                if "format" in options and options["format"]["formatter"] == formatter
            )

    def get_cell_kwargs(
        self,
        datarn,
        datacn,
        key="format",
        cell=True,
        row=True,
        column=True,
        entire=True,
    ):
        if cell and (datarn, datacn) in self.cell_options and key in self.cell_options[(datarn, datacn)]:
            return self.cell_options[(datarn, datacn)][key]
        if row and datarn in self.row_options and key in self.row_options[datarn]:
            return self.row_options[datarn][key]
        if column and datacn in self.col_options and key in self.col_options[datacn]:
            return self.col_options[datacn][key]
        if entire and key in self.options:
            return self.options[key]
        return {}

    def get_space_bot(self, r, text_editor_h=None):
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
                sheet_h = int(
                    self.row_positions[-1] + 1 + self.empty_vertical - (self.row_positions[r] + text_editor_h)
                )
        if win_h > 0:
            win_h -= 1
        if sheet_h > 0:
            sheet_h -= 1
        return win_h if win_h >= sheet_h else sheet_h

    def get_dropdown_height_anchor(self, datarn, datacn, text_editor_h=None):
        win_h = 5
        for i, v in enumerate(self.get_cell_kwargs(datarn, datacn, key="dropdown")["values"]):
            v_numlines = len(v.split("\n") if isinstance(v, str) else f"{v}".split("\n"))
            if v_numlines > 1:
                win_h += self.fl_ins + (v_numlines * self.xtra_lines_increment) + 5  # end of cell
            else:
                win_h += self.min_row_height
            if i == 5:
                break
        if win_h > 500:
            win_h = 500
        space_bot = self.get_space_bot(datarn, text_editor_h)
        space_top = int(self.row_positions[datarn])
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
    def open_dropdown_window(self, r, c, event=None):
        self.destroy_text_editor("Escape")
        self.destroy_opened_dropdown_window()
        datarn = r if self.all_rows_displayed else self.displayed_rows[r]
        datacn = c if self.all_columns_displayed else self.displayed_columns[c]
        kwargs = self.get_cell_kwargs(datarn, datacn, key="dropdown")
        if kwargs["state"] == "normal":
            if not self.open_text_editor(event=event, r=r, c=c, dropdown=True):
                return
        win_h, anchor = self.get_dropdown_height_anchor(datarn, datacn)
        window = self.parentframe.dropdown_class(
            self.winfo_toplevel(),
            r,
            c,
            width=self.col_positions[c + 1] - self.col_positions[c] + 1,
            height=win_h,
            font=self.table_font,
            colors={
                "bg": self.popup_menu_bg,
                "fg": self.popup_menu_fg,
                "highlight_bg": self.popup_menu_highlight_bg,
                "highlight_fg": self.popup_menu_highlight_fg,
            },
            outline_color=self.table_selected_cells_border_fg,
            outline_thickness=2,
            values=kwargs["values"],
            close_dropdown_window=self.close_dropdown_window,
            search_function=kwargs["search_function"],
            arrowkey_RIGHT=self.arrowkey_RIGHT,
            arrowkey_LEFT=self.arrowkey_LEFT,
            align="w",
        )  # self.get_cell_align(r, c)
        if kwargs["state"] == "normal":
            if anchor == "nw":
                ypos = self.row_positions[r] + self.text_editor.h_ - 1
            else:
                ypos = self.row_positions[r]
            kwargs["canvas_id"] = self.create_window((self.col_positions[c], ypos), window=window, anchor=anchor)
            self.text_editor.textedit.bind(
                "<<TextModified>>",
                lambda x: window.search_and_see(
                    DropDownModifiedEvent("ComboboxModified", r, c, self.text_editor.get())
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
            if anchor == "nw":
                ypos = self.row_positions[r + 1]
            else:
                ypos = self.row_positions[r]
            kwargs["canvas_id"] = self.create_window((self.col_positions[c], ypos), window=window, anchor=anchor)
            self.update_idletasks()
            window.bind("<FocusOut>", lambda x: self.close_dropdown_window(r, c))
            window.focus()
            redraw = True
        self.existing_dropdown_window = window
        kwargs["window"] = window
        self.existing_dropdown_canvas_id = kwargs["canvas_id"]
        if redraw:
            self.main_table_redraw_grid_and_text(redraw_header=False, redraw_row_index=False)

    # displayed indexes, not data
    def close_dropdown_window(self, r=None, c=None, selection=None, redraw=True):
        if r is not None and c is not None and selection is not None:
            datacn = c if self.all_columns_displayed else self.displayed_columns[c]
            datarn = r if self.all_rows_displayed else self.displayed_rows[r]
            kwargs = self.get_cell_kwargs(datarn, datacn, key="dropdown")
            if kwargs["select_function"] is not None:  # user has specified a selection function
                kwargs["select_function"](EditCellEvent(r, c, "ComboboxSelected", f"{selection}", "end_edit_cell"))
            if self.extra_end_edit_cell_func is None:
                self.set_cell_data_undo(r, c, value=selection, redraw=not redraw)
            elif self.extra_end_edit_cell_func is not None and self.edit_cell_validation:
                validation = self.extra_end_edit_cell_func(
                    EditCellEvent(r, c, "ComboboxSelected", f"{selection}", "end_edit_cell")
                )
                if validation is not None:
                    selection = validation
                self.set_cell_data_undo(r, c, value=selection, redraw=not redraw)
            elif self.extra_end_edit_cell_func is not None and not self.edit_cell_validation:
                self.set_cell_data_undo(r, c, value=selection, redraw=not redraw)
                self.extra_end_edit_cell_func(EditCellEvent(r, c, "ComboboxSelected", f"{selection}", "end_edit_cell"))
            self.focus_set()
            self.recreate_all_selection_boxes()
        self.destroy_text_editor("Escape")
        self.destroy_opened_dropdown_window(r, c)
        if redraw:
            self.refresh()

    def get_existing_dropdown_coords(self):
        if self.existing_dropdown_window is not None:
            return int(self.existing_dropdown_window.r), int(self.existing_dropdown_window.c)
        return None

    def mouseclick_outside_editor_or_dropdown(self):
        closed_dd_coords = self.get_existing_dropdown_coords()
        if self.text_editor_loc is not None and self.text_editor is not None:
            self.close_text_editor(editor_info=self.text_editor_loc + ("ButtonPress-1",))
        else:
            self.destroy_text_editor("Escape")
        if closed_dd_coords is not None:
            self.destroy_opened_dropdown_window(
                closed_dd_coords[0], closed_dd_coords[1]
            )  # displayed coords not data, necessary for b1 function
        return closed_dd_coords

    def mouseclick_outside_editor_or_dropdown_all_canvases(self):
        self.CH.mouseclick_outside_editor_or_dropdown()
        self.RI.mouseclick_outside_editor_or_dropdown()
        return self.mouseclick_outside_editor_or_dropdown()

    # function can receive 4 None args
    def destroy_opened_dropdown_window(self, r=None, c=None, datarn=None, datacn=None):
        if r is None and datarn is None and c is None and datacn is None and self.existing_dropdown_window is not None:
            r, c = self.get_existing_dropdown_coords()
        if c is not None or datacn is not None:
            if datacn is None:
                datacn_ = c if self.all_columns_displayed else self.displayed_columns[c]
            else:
                datacn_ = datacn
        else:
            datacn_ = None
        if r is not None or datarn is not None:
            if datarn is None:
                datarn_ = r if self.all_rows_displayed else self.displayed_rows[r]
            else:
                datarn_ = datarn
        else:
            datarn_ = None
        try:
            self.delete(self.existing_dropdown_canvas_id)
        except Exception:
            pass
        self.existing_dropdown_canvas_id = None
        try:
            self.existing_dropdown_window.destroy()
        except Exception:
            pass
        kwargs = self.get_cell_kwargs(datarn_, datacn_, key="dropdown")
        if kwargs:
            kwargs["canvas_id"] = "no dropdown open"
            kwargs["window"] = "no dropdown open"
            try:
                self.delete(kwargs["canvas_id"])
            except Exception:
                pass
        self.existing_dropdown_window = None

    def get_displayed_col_from_datacn(self, datacn):
        try:
            return self.displayed_columns.index(datacn)
        except Exception:
            return None

    def delete_cell_options_dropdown(self, datarn, datacn):
        self.destroy_opened_dropdown_window()
        if (datarn, datacn) in self.cell_options and "dropdown" in self.cell_options[(datarn, datacn)]:
            del self.cell_options[(datarn, datacn)]["dropdown"]

    def delete_cell_options_checkbox(self, datarn, datacn):
        if (datarn, datacn) in self.cell_options and "checkbox" in self.cell_options[(datarn, datacn)]:
            del self.cell_options[(datarn, datacn)]["checkbox"]

    def delete_cell_options_dropdown_and_checkbox(self, datarn, datacn):
        self.delete_cell_options_dropdown(datarn, datacn)
        self.delete_cell_options_checkbox(datarn, datacn)

    def delete_row_options_dropdown(self, datarn):
        self.destroy_opened_dropdown_window()
        if datarn in self.row_options and "dropdown" in self.row_options[datarn]:
            del self.row_options[datarn]["dropdown"]

    def delete_row_options_checkbox(self, datarn):
        if datarn in self.row_options and "checkbox" in self.row_options[datarn]:
            del self.row_options[datarn]["checkbox"]

    def delete_row_options_dropdown_and_checkbox(self, datarn):
        self.delete_row_options_dropdown(datarn)
        self.delete_row_options_checkbox(datarn)

    def delete_column_options_dropdown(self, datacn):
        self.destroy_opened_dropdown_window()
        if datacn in self.col_options and "dropdown" in self.col_options[datacn]:
            del self.col_options[datacn]["dropdown"]

    def delete_column_options_checkbox(self, datacn):
        if datacn in self.col_options and "checkbox" in self.col_options[datacn]:
            del self.col_options[datacn]["checkbox"]

    def delete_column_options_dropdown_and_checkbox(self, datacn):
        self.delete_column_options_dropdown(datacn)
        self.delete_column_options_checkbox(datacn)

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
