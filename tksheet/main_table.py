from __future__ import annotations

import csv as csv
import io
import tkinter as tk
from bisect import (
    bisect_left,
    bisect_right,
)
from collections import (
    defaultdict,
    deque,
)
from collections.abc import (
    Callable,
    Generator,
    Hashable,
    Iterator,
    Sequence,
)
from functools import (
    partial,
)
from itertools import (
    accumulate,
    chain,
    cycle,
    islice,
    product,
    repeat,
)
from math import (
    ceil,
    floor,
)
from operator import itemgetter
from tkinter import TclError
from typing import Literal

from .colors import (
    color_map,
)
from .formatters import (
    data_to_str,
    format_data,
    get_clipboard_data,
    get_data_with_valid_check,
    is_bool_like,
    try_to_bool,
)
from .functions import (
    consecutive_chunks,
    coords_tag_to_box_nt,
    coords_tag_to_int_tuple,
    decompress_load,
    diff_gen,
    diff_list,
    ev_stack_dict,
    event_dict,
    gen_formatted,
    get_checkbox_points,
    get_new_indexes,
    get_seq_without_gaps_at_index,
    index_exists,
    insert_items,
    is_iterable,
    is_type_int,
    len_to_idx,
    mod_event_val,
    mod_span,
    move_elements_by_mapping,
    pickle_obj,
    span_idxs_post_move,
    try_binding,
    unpickle_obj,
)
from .other_classes import (
    Box_t,
    CurrentlySelectedClass,
    EventDataDict,
    TextEditor,
)
from .vars import (
    USER_OS,
    arrowkey_bindings_helper,
    ctrl_key,
    rc_binding,
    symbols_set,
    val_modifying_options,
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
        self.current_cursor = ""
        self.b1_pressed_loc = None
        self.existing_dropdown_canvas_id = None
        self.existing_dropdown_window = None
        self.closed_dropdown = None
        self.last_selected = None
        self.centre_alignment_text_mod_indexes = (slice(1, None), slice(None, -1))
        self.c_align_cyc = cycle(self.centre_alignment_text_mod_indexes)
        self.span = self.parentframe.span
        self.synced_scrolls = set()
        self.set_cell_sizes_on_zoom = kwargs["set_cell_sizes_on_zoom"]

        self.disp_ctrl_outline = {}
        self.disp_text = {}
        self.disp_high = {}
        self.disp_grid = {}
        self.disp_fill_sels = {}
        self.disp_bord_sels = {}
        self.disp_resize_lines = {}
        self.disp_dropdown = {}
        self.disp_checkbox = {}
        self.hidd_ctrl_outline = {}
        self.hidd_text = {}
        self.hidd_high = {}
        self.hidd_grid = {}
        self.hidd_fill_sels = {}
        self.hidd_bord_sels = {}
        self.hidd_resize_lines = {}
        self.hidd_dropdown = {}
        self.hidd_checkbox = {}

        self.named_spans = {}
        self.cell_options = {}
        self.col_options = {}
        self.row_options = {}

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
        self.purge_undo_and_redo_stack()

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
        self.being_drawn_item = None
        self.extra_motion_func = None
        self.extra_b1_press_func = None
        self.extra_b1_motion_func = None
        self.extra_b1_release_func = None
        self.extra_double_b1_func = None
        self.extra_rc_func = None

        self.edit_validation_func = None

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

        self.table_font = [
            kwargs["font"][0],
            int(kwargs["font"][1] * kwargs["zoom"] / 100),
            kwargs["font"][2],
        ]

        self.index_font = [
            kwargs["index_font"][0],
            int(kwargs["index_font"][1] * kwargs["zoom"] / 100),
            kwargs["index_font"][2],
        ]

        self.header_font = [
            kwargs["header_font"][0],
            int(kwargs["header_font"][1] * kwargs["zoom"] / 100),
            kwargs["header_font"][2],
        ]
        for fnt in (self.table_font, self.index_font, self.header_font):
            if fnt[1] < 1:
                fnt[1] = 1
        self.table_font = tuple(self.table_font)
        self.index_font = tuple(self.index_font)
        self.header_font = tuple(self.header_font)

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
        self.set_col_positions(itr=[])
        self.set_row_positions(itr=[])
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
        self.table_selected_box_cells_fg = kwargs["table_selected_box_cells_fg"]
        self.table_selected_box_rows_fg = kwargs["table_selected_box_rows_fg"]
        self.table_selected_box_columns_fg = kwargs["table_selected_box_columns_fg"]
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

    def refresh(self, event: object = None) -> None:
        self.main_table_redraw_grid_and_text(True, True)

    def window_configured(self, event: object) -> None:
        w = self.parentframe.winfo_width()
        if w != self.parentframe_width:
            self.parentframe_width = w
            self.allow_auto_resize_columns = True
        h = self.parentframe.winfo_height()
        if h != self.parentframe_height:
            self.parentframe_height = h
            self.allow_auto_resize_rows = True
        self.main_table_redraw_grid_and_text(True, True)

    def basic_bindings(self, enable: bool = True) -> None:
        bindings = (
            ("<Configure>", self, self.window_configured),
            ("<Motion>", self, self.mouse_motion),
            ("<ButtonPress-1>", self, self.b1_press),
            ("<B1-Motion>", self, self.b1_motion),
            ("<ButtonRelease-1>", self, self.b1_release),
            ("<Double-Button-1>", self, self.double_b1),
            ("<MouseWheel>", self, self.mousewheel),
            ("<Shift-ButtonPress-1>", self, self.shift_b1_press),
            ("<Shift-ButtonPress-1>", self.CH, self.CH.shift_b1_press),
            ("<Shift-ButtonPress-1>", self.RI, self.RI.shift_b1_press),
            ("<MouseWheel>", self.RI, self.mousewheel),
            (rc_binding, self, self.rc),
            (f"<{ctrl_key}-ButtonPress-1>", self, self.ctrl_b1_press),
            (f"<{ctrl_key}-ButtonPress-1>", self.CH, self.CH.ctrl_b1_press),
            (f"<{ctrl_key}-ButtonPress-1>", self.RI, self.RI.ctrl_b1_press),
            (f"<{ctrl_key}-Shift-ButtonPress-1>", self, self.ctrl_shift_b1_press),
            (f"<{ctrl_key}-Shift-ButtonPress-1>", self.CH, self.CH.ctrl_shift_b1_press),
            (f"<{ctrl_key}-Shift-ButtonPress-1>", self.RI, self.RI.ctrl_shift_b1_press),
            (f"<{ctrl_key}-B1-Motion>", self, self.ctrl_b1_motion),
            (f"<{ctrl_key}-B1-Motion>", self.CH, self.CH.ctrl_b1_motion),
            (f"<{ctrl_key}-B1-Motion>", self.RI, self.RI.ctrl_b1_motion),
        )
        all_canvas_bindings = (
            ("<Shift-MouseWheel>", self.shift_mousewheel),
            ("<Control-MouseWheel>", self.ctrl_mousewheel),
            ("<Control-plus>", self.zoom_in),
            ("<Control-equal>", self.zoom_in),
            ("<Control-minus>", self.zoom_out),
        )
        all_canvas_linux_bindings = {
            ("<Button-4>", self.mousewheel),
            ("<Button-5>", self.mousewheel),
            ("<Shift-Button-4>", self.shift_mousewheel),
            ("<Shift-Button-5>", self.shift_mousewheel),
            ("<Control-Button-4>", self.ctrl_mousewheel),
            ("<Control-Button-5>", self.ctrl_mousewheel),
        }
        if enable:
            for b in bindings:
                b[1].bind(b[0], b[2])
            for b in all_canvas_bindings:
                for canvas in (self, self.RI, self.CH):
                    canvas.bind(b[0], b[1])
            if USER_OS == "linux":
                for b in all_canvas_linux_bindings:
                    for canvas in (self, self.RI, self.CH):
                        canvas.bind(b[0], b[1])
        else:
            for b in bindings:
                b[1].unbind(b[0])
            for b in all_canvas_bindings:
                for canvas in (self, self.RI, self.CH):
                    canvas.unbind(b[0])
            if USER_OS == "linux":
                for b in all_canvas_linux_bindings:
                    for canvas in (self, self.RI, self.CH):
                        canvas.unbind(b[0])

    def show_ctrl_outline(
        self,
        canvas: Literal["table"] = "table",
        start_cell: tuple[int, int] = (0, 0),
        end_cell: tuple[int, int] = (0, 0),
        dash: tuple[int, int] = (20, 20),
        outline: str | None = None,
        delete_on_timer: bool = True,
    ) -> None:
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

    def create_ctrl_outline(
        self,
        x1: int,
        y1: int,
        x2: int,
        y2: int,
        fill: str,
        dash: tuple[int, int],
        width: int,
        outline: str,
        tag: str | tuple[str, ...],
    ) -> None:
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

    def delete_ctrl_outlines(self) -> None:
        self.hidd_ctrl_outline.update(self.disp_ctrl_outline)
        self.disp_ctrl_outline = {}
        for t, sh in self.hidd_ctrl_outline.items():
            if sh:
                self.itemconfig(t, state="hidden")
                self.hidd_ctrl_outline[t] = False

    def get_ctrl_x_c_boxes(self) -> tuple[dict[tuple[int, int, int, int], str], int]:
        currently_selected = self.currently_selected()
        boxes = {}
        maxrows = 0
        if currently_selected.type_ in ("cell", "column"):
            curr_box = self.get_box_containing_current()
            maxrows = curr_box[2] - curr_box[0]
            for item in self.get_selection_items(rows=False, current=False):
                tags = self.gettags(item)
                box = coords_tag_to_box_nt(tags[1])
                if maxrows >= box[2] - box[0]:
                    boxes[box] = tags[0]
        else:
            for item in self.get_selection_items(columns=False, cells=False, current=False):
                boxes[coords_tag_to_box_nt(self.gettags(item)[1])] = "rows"
        return boxes, maxrows

    def io_csv_writer(self) -> tuple[io.StringIO, csv.writer]:
        s = io.StringIO()
        writer = csv.writer(
            s,
            dialect=csv.excel_tab,
            delimiter=self.to_clipboard_delimiter,
            quotechar=self.to_clipboard_quotechar,
            lineterminator=self.to_clipboard_lineterminator,
        )
        return s, writer

    def ctrl_c(self, event=None) -> None:
        if not self.anything_selected():
            return
        currently_selected = self.currently_selected()
        event_data = event_dict(
            sheet=self.parentframe.name,
            selected=currently_selected,
        )
        event_data["eventname"] = "begin_ctrl_c"
        boxes, maxrows = self.get_ctrl_x_c_boxes()
        event_data["selection_boxes"] = boxes
        s, writer = self.io_csv_writer()
        if not try_binding(self.extra_begin_ctrl_c_func, event_data):
            return
        if currently_selected.type_ in ("cell", "column"):
            for rn in range(maxrows):
                row = []
                for r1, c1, r2, c2 in boxes:
                    datarn = (r1 + rn) if self.all_rows_displayed else self.displayed_rows[r1 + rn]
                    for c in range(c1, c2):
                        datacn = self.datacn(c)
                        v = self.get_cell_clipboard(datarn, datacn)
                        event_data["cells"]["table"][(datarn, datacn)] = v
                        row.append(v)
                writer.writerow(row)
        else:
            for r1, c1, r2, c2 in boxes:
                for rn in range(r2 - r1):
                    row = []
                    datarn = (r1 + rn) if self.all_rows_displayed else self.displayed_rows[r1 + rn]
                    for c in range(c1, c2):
                        datacn = self.datacn(c)
                        v = self.get_cell_clipboard(datarn, datacn)
                        event_data["cells"]["table"][(datarn, datacn)] = v
                        row.append(v)
                    writer.writerow(row)
        for r1, c1, r2, c2 in boxes:
            self.show_ctrl_outline(canvas="table", start_cell=(c1, r1), end_cell=(c2, r2))
        self.clipboard_clear()
        self.clipboard_append(s.getvalue())
        self.update_idletasks()
        try_binding(self.extra_end_ctrl_c_func, event_data, "end_ctrl_c")

    def ctrl_x(self, event=None) -> None:
        if not self.anything_selected():
            return
        currently_selected = self.currently_selected()
        event_data = event_dict(
            name="edit_table",
            sheet=self.parentframe.name,
            selected=currently_selected,
        )
        boxes, maxrows = self.get_ctrl_x_c_boxes()
        event_data["selection_boxes"] = boxes
        s, writer = self.io_csv_writer()
        if not try_binding(self.extra_begin_ctrl_x_func, event_data, "begin_ctrl_x"):
            return
        if currently_selected.type_ in ("cell", "column"):
            for rn in range(maxrows):
                row = []
                for r1, c1, r2, c2 in boxes:
                    if r2 - r1 < maxrows:
                        continue
                    datarn = (r1 + rn) if self.all_rows_displayed else self.displayed_rows[r1 + rn]
                    for c in range(c1, c2):
                        datacn = self.datacn(c)
                        row.append(self.get_cell_clipboard(datarn, datacn))
                        val = self.get_value_for_empty_cell(datarn, datacn)
                        if not self.edit_validation_func or (
                            self.edit_validation_func
                            and (val := self.edit_validation_func(mod_event_val(event_data, val, (r1 + rn, c))))
                            is not None
                        ):
                            event_data = self.event_data_set_cell(
                                datarn,
                                datacn,
                                val,
                                event_data,
                            )
                writer.writerow(row)
        else:
            for r1, c1, r2, c2 in boxes:
                for rn in range(r2 - r1):
                    row = []
                    datarn = (r1 + rn) if self.all_rows_displayed else self.displayed_rows[r1 + rn]
                    for c in range(c1, c2):
                        datacn = self.datacn(c)
                        row.append(self.get_cell_clipboard(datarn, datacn))
                        val = self.get_value_for_empty_cell(datarn, datacn)
                        if not self.edit_validation_func or (
                            self.edit_validation_func
                            and (val := self.edit_validation_func(mod_event_val(event_data, val, (r1 + rn, c))))
                            is not None
                        ):
                            event_data = self.event_data_set_cell(
                                datarn,
                                datacn,
                                val,
                                event_data,
                            )
                    writer.writerow(row)
        if event_data["cells"]["table"]:
            self.undo_stack.append(ev_stack_dict(event_data))
        self.clipboard_clear()
        self.clipboard_append(s.getvalue())
        self.update_idletasks()
        self.refresh()
        for r1, c1, r2, c2 in boxes:
            self.show_ctrl_outline(canvas="table", start_cell=(c1, r1), end_cell=(c2, r2))
        if event_data["cells"]["table"]:
            try_binding(self.extra_end_ctrl_x_func, event_data, "end_ctrl_x")
        self.sheet_modified(event_data)

    def get_box_containing_current(self) -> tuple[int, int, int, int]:
        item = self.get_selection_items(cells=False, rows=False, columns=False)[-1]
        return coords_tag_to_box_nt(self.gettags(item)[1])

    def ctrl_v(self, event: object = None) -> None:
        if not self.expand_sheet_if_paste_too_big and (len(self.col_positions) == 1 or len(self.row_positions) == 1):
            return
        currently_selected = self.currently_selected()
        event_data = event_dict(
            name="edit_table",
            sheet=self.parentframe.name,
            selected=currently_selected,
        )
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
        new_data_numcols = max(map(len, data))
        new_data_numrows = len(data)
        for rn, r in enumerate(data):
            if len(r) < new_data_numcols:
                data[rn] += list(repeat("", new_data_numcols - len(r)))
        (
            lastbox_r1,
            lastbox_c1,
            lastbox_r2,
            lastbox_c2,
        ) = self.get_box_containing_current()
        lastbox_numrows = lastbox_r2 - lastbox_r1
        lastbox_numcols = lastbox_c2 - lastbox_c1
        if lastbox_numrows > new_data_numrows and not lastbox_numrows % new_data_numrows:
            nd = []
            for _ in range(int(lastbox_numrows / new_data_numrows)):
                nd.extend([r.copy() for r in data])
            data.extend(nd)
            new_data_numrows *= int(lastbox_numrows / new_data_numrows)

        if lastbox_numcols > new_data_numcols and not lastbox_numcols % new_data_numcols:
            for rn, r in enumerate(data):
                for _ in range(int(lastbox_numcols / new_data_numcols)):
                    data[rn].extend(r.copy())
            new_data_numcols *= int(lastbox_numcols / new_data_numcols)
        event_data["data"] = data
        added_rows = 0
        added_cols = 0
        total_data_cols = None
        if self.expand_sheet_if_paste_too_big:
            # determine number of columns and/or rows to add to sheet
            if selected_c + new_data_numcols > len(self.col_positions) - 1:
                total_data_cols = self.equalize_data_row_lengths()
                added_cols = selected_c + new_data_numcols - len(self.col_positions) + 1
                if (
                    isinstance(self.paste_insert_column_limit, int)
                    and self.paste_insert_column_limit < len(self.col_positions) - 1 + added_cols
                ):
                    added_cols = self.paste_insert_column_limit - len(self.col_positions) - 1
            if selected_r + new_data_numrows > len(self.row_positions) - 1:
                added_rows = selected_r + new_data_numrows - len(self.row_positions) + 1
                if (
                    isinstance(self.paste_insert_row_limit, int)
                    and self.paste_insert_row_limit < len(self.row_positions) - 1 + added_rows
                ):
                    added_rows = self.paste_insert_row_limit - len(self.row_positions) - 1
        if selected_c + new_data_numcols > len(self.col_positions) - 1:
            adjusted_new_data_numcols = len(self.col_positions) - 1 - selected_c
        else:
            adjusted_new_data_numcols = new_data_numcols
        if selected_r + new_data_numrows > len(self.row_positions) - 1:
            adjusted_new_data_numrows = len(self.row_positions) - 1 - selected_r
        else:
            adjusted_new_data_numrows = new_data_numrows
        selected_r_adjusted_new_data_numrows = selected_r + adjusted_new_data_numrows
        selected_c_adjusted_new_data_numcols = selected_c + adjusted_new_data_numcols
        boxes = {
            (
                selected_r,
                selected_c,
                selected_r_adjusted_new_data_numrows,
                selected_c_adjusted_new_data_numcols,
            ): "cells"
        }
        event_data["selection_boxes"] = boxes
        if not try_binding(self.extra_begin_ctrl_v_func, event_data, "begin_ctrl_v"):
            return
        # the order of actions here is important:
        # edit existing sheet (not including any added rows/columns)

        # then if there are any added rows/columns:
        # create empty rows/columns dicts for any added rows/columns
        # edit those dicts with so far unused cells of data from clipboard
        # instead of editing table using set cell data, add any new rows then columns with pasted data
        for ndr, r in enumerate(range(selected_r, selected_r_adjusted_new_data_numrows)):
            for ndc, c in enumerate(range(selected_c, selected_c_adjusted_new_data_numcols)):
                val = data[ndr][ndc]
                if not self.edit_validation_func or (
                    self.edit_validation_func
                    and (val := self.edit_validation_func(mod_event_val(event_data, val, (r, c)))) is not None
                ):
                    event_data = self.event_data_set_cell(
                        datarn=self.datarn(r),
                        datacn=self.datacn(c),
                        value=val,
                        event_data=event_data,
                    )
        if added_rows > 0:
            ctr = 0
            data_ins_row = len(self.data)
            displayed_ins_row = len(self.row_positions) - 1
            if total_data_cols is None:
                total_data_cols = self.total_data_cols()
            rows, index, row_heights = self.get_args_for_add_rows(
                data_ins_row=data_ins_row,
                displayed_ins_row=displayed_ins_row,
                numrows=added_rows,
            )
            for ndr, r in zip(
                range(
                    adjusted_new_data_numrows,
                    new_data_numrows,
                ),
                reversed(rows),
            ):
                for ndc, c in enumerate(
                    range(
                        selected_c,
                        selected_c_adjusted_new_data_numcols,
                    )
                ):
                    val = data[ndr][ndc]
                    datacn = self.datacn(c)
                    if not self.edit_validation_func or (
                        self.edit_validation_func
                        and (val := self.edit_validation_func(mod_event_val(event_data, val, (r, c)))) is not None
                        and self.input_valid_for_cell(r, datacn, val, ignore_empty=True)
                    ):
                        rows[r][datacn] = val
                        ctr += 1
            if ctr:
                event_data = self.add_rows(
                    rows=rows,
                    index=index,
                    row_heights=row_heights,
                    event_data=event_data,
                )
        if added_cols > 0:
            ctr = 0
            data_ins_col = total_data_cols
            displayed_ins_col = len(self.col_positions) - 1
            columns, headers, column_widths = self.get_args_for_add_columns(
                data_ins_col=data_ins_col,
                displayed_ins_col=displayed_ins_col,
                numcols=added_cols,
            )
            for ndr, r in enumerate(
                range(
                    selected_r,
                    selected_r + new_data_numrows,
                )
            ):
                for ndc, c in zip(
                    range(
                        adjusted_new_data_numcols,
                        new_data_numcols,
                    ),
                    reversed(columns),
                ):
                    val = data[ndr][ndc]
                    datarn = self.datarn(r)
                    if not self.edit_validation_func or (
                        self.edit_validation_func
                        and (val := self.edit_validation_func(mod_event_val(event_data, val, (r, c)))) is not None
                        and self.input_valid_for_cell(datarn, c, val, ignore_empty=True)
                    ):
                        columns[c][datarn] = val
                        ctr += 1
            if ctr:
                event_data = self.add_columns(
                    columns=columns,
                    header=headers,
                    column_widths=column_widths,
                    event_data=event_data,
                )
        self.deselect("all", redraw=False)
        if event_data["cells"]["table"] or event_data["added"]["rows"] or event_data["added"]["columns"]:
            self.undo_stack.append(ev_stack_dict(event_data))
        self.create_selection_box(
            selected_r,
            selected_c,
            selected_r_adjusted_new_data_numrows,
            selected_c_adjusted_new_data_numcols,
            "cells",
            run_binding=True,
        )
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
        if event_data["cells"]["table"] or event_data["added"]["rows"] or event_data["added"]["columns"]:
            try_binding(self.extra_end_ctrl_v_func, event_data, "end_ctrl_v")
        self.sheet_modified(event_data)

    def delete_key(self, event: object = None) -> None:
        if not self.anything_selected():
            return
        currently_selected = self.currently_selected()
        event_data = event_dict(
            name="edit_table",
            sheet=self.parentframe.name,
            selected=currently_selected,
        )
        boxes = self.get_boxes()
        event_data["selection_boxes"] = boxes
        if not try_binding(self.extra_begin_delete_key_func, event_data, "begin_delete"):
            return
        for r1, c1, r2, c2 in boxes:
            for r in range(r1, r2):
                for c in range(c1, c2):
                    datarn, datacn = self.datarn(r), self.datacn(c)
                    val = self.get_value_for_empty_cell(datarn, datacn)
                    if not self.edit_validation_func or (
                        self.edit_validation_func
                        and (val := self.edit_validation_func(mod_event_val(event_data, val, (r, c)))) is not None
                    ):
                        event_data = self.event_data_set_cell(
                            datarn,
                            datacn,
                            val,
                            event_data,
                        )
        if event_data["cells"]["table"]:
            self.undo_stack.append(ev_stack_dict(event_data))
            try_binding(self.extra_end_delete_key_func, event_data, "end_delete")
        self.refresh()
        self.sheet_modified(event_data)

    def event_data_set_cell(self, datarn: int, datacn: int, value: object, event_data: dict) -> EventDataDict:
        if self.input_valid_for_cell(datarn, datacn, value):
            event_data["cells"]["table"][(datarn, datacn)] = self.get_cell_data(datarn, datacn)
            self.set_cell_data(datarn, datacn, value)
        return event_data

    def get_args_for_move_columns(
        self,
        move_to: int,
        to_move: list[int],
        data_indexes: bool = False,
    ) -> tuple[dict[int, int], dict[int, int], int, dict[int, int]]:
        if not data_indexes or self.all_columns_displayed:
            disp_new_idxs = get_new_indexes(
                seqlen=len(self.col_positions) - 1,
                move_to=move_to,
                to_move=to_move,
            )
        else:
            disp_new_idxs = {}
        totalcols = self.equalize_data_row_lengths(at_least_cols=move_to + 1)
        if self.all_columns_displayed or data_indexes:
            data_new_idxs = get_new_indexes(seqlen=totalcols, move_to=move_to, to_move=to_move)
        elif not self.all_columns_displayed and not data_indexes:
            data_new_idxs = get_new_indexes(seqlen=len(self.displayed_columns), move_to=move_to, to_move=to_move)
            data_old_idxs = dict(zip(data_new_idxs.values(), data_new_idxs))
            data_new_idxs = dict(
                zip(
                    move_elements_by_mapping(self.displayed_columns, data_new_idxs, data_old_idxs),
                    self.displayed_columns,
                )
            )
        return data_new_idxs, dict(zip(data_new_idxs.values(), data_new_idxs)), totalcols, disp_new_idxs

    def move_columns_adjust_options_dict(
        self,
        data_new_idxs: dict[int, int],
        data_old_idxs: dict[int, int],
        totalcols: int | None,
        disp_new_idxs: None | dict[int, int] = None,
        move_data: bool = True,
        create_selections: bool = True,
        data_indexes: bool = False,
        event_data: EventDataDict | None = None,
    ) -> tuple[dict[int, int], dict[int, int], EventDataDict]:
        if not isinstance(totalcols, int):
            totalcols = max(data_new_idxs.values(), default=0)
            if totalcols:
                totalcols += 1
            totalcols = self.equalize_data_row_lengths(at_least_cols=totalcols)
        if event_data is None:
            event_data = event_dict(
                name="move_columns",
                sheet=self.parentframe.name,
                boxes=self.get_boxes(),
                selected=self.currently_selected(),
            )
            event_data["moved"]["columns"] = {
                "data": data_new_idxs,
                "displayed": {} if disp_new_idxs is None else disp_new_idxs,
            }
        event_data["options"] = self.pickle_options()
        event_data["named_spans"] = {k: span.pickle_self() for k, span in self.named_spans.items()}
        if disp_new_idxs and (not data_indexes or self.all_columns_displayed):
            self.deselect("all", run_binding=False, redraw=False)
            self.set_col_positions(
                itr=move_elements_by_mapping(
                    self.get_column_widths(),
                    disp_new_idxs,
                    dict(
                        zip(
                            disp_new_idxs.values(),
                            disp_new_idxs,
                        )
                    ),
                )
            )
            if create_selections:
                for chunk in consecutive_chunks(sorted(disp_new_idxs.values())):
                    self.create_selection_box(
                        0,
                        chunk[0],
                        len(self.row_positions) - 1,
                        chunk[-1] + 1,
                        "columns",
                        run_binding=True,
                    )
        if move_data:
            self.data[:] = [
                move_elements_by_mapping(
                    r,
                    data_new_idxs,
                    data_old_idxs,
                )
                for r in self.data
            ]
            maxidx = len_to_idx(totalcols)
            self.CH.fix_header(maxidx)
            if isinstance(self._headers, list) and self._headers:
                self._headers = move_elements_by_mapping(self._headers, data_new_idxs, data_old_idxs)
            maxidx = self.get_max_column_idx(maxidx)
            full_new_idxs = self.get_full_new_idxs(
                max_idx=maxidx,
                new_idxs=data_new_idxs,
                old_idxs=data_old_idxs,
            )
            full_old_idxs = dict(zip(full_new_idxs.values(), full_new_idxs))
            self.cell_options = {(k[0], full_new_idxs[k[1]]): v for k, v in self.cell_options.items()}
            self.col_options = {full_new_idxs[k]: v for k, v in self.col_options.items()}
            self.CH.cell_options = {full_new_idxs[k]: v for k, v in self.CH.cell_options.items()}
            totalrows = self.total_data_rows()
            new_ops = self.parentframe.create_options_from_span
            qkspan = self.span()
            for name, span in self.named_spans.items():
                # span is neither a cell options nor col options span, continue
                if not isinstance(span["from_c"], int):
                    continue
                oldupto_colrange, newupto_colrange, newfrom, newupto = span_idxs_post_move(
                    data_new_idxs,
                    full_new_idxs,
                    totalcols,
                    span,
                    "c",
                )
                # add cell/col kwargs for columns that are new to the span
                old_span_idxs = set(full_new_idxs[k] for k in range(span["from_c"], oldupto_colrange))
                for k in range(newfrom, newupto_colrange):
                    if k not in old_span_idxs:
                        oldidx = full_old_idxs[k]
                        # event_data is used to preserve old cell value
                        # in case cells are modified by
                        # formatting, checkboxes, dropdown boxes
                        if (
                            span["type_"] in val_modifying_options
                            and span["header"]
                            and oldidx not in event_data["cells"]["header"]
                        ):
                            event_data["cells"]["header"][oldidx] = self.CH.get_cell_data(k)
                        # the span targets columns
                        if span["from_r"] is None:
                            if span["type_"] in val_modifying_options:
                                for datarn in range(len(self.data)):
                                    if (datarn, oldidx) not in event_data["cells"]["table"]:
                                        event_data["cells"]["table"][(datarn, oldidx)] = self.get_cell_data(datarn, k)
                            # create new col options
                            new_ops(
                                mod_span(
                                    qkspan,
                                    span,
                                    from_c=k,
                                    upto_c=k + 1,
                                )
                            )
                        # the span targets cells
                        else:
                            rng_upto_r = totalrows if span["upto_r"] is None else span["upto_r"]
                            for datarn in range(span["from_r"], rng_upto_r):
                                if (
                                    span["type_"] in val_modifying_options
                                    and (datarn, oldidx) not in event_data["cells"]["table"]
                                ):
                                    event_data["cells"]["table"][(datarn, oldidx)] = self.get_cell_data(datarn, k)
                                # create new cell options
                                new_ops(
                                    mod_span(
                                        qkspan,
                                        span,
                                        from_r=datarn,
                                        upto_r=datarn + 1,
                                        from_c=k,
                                        upto_c=k + 1,
                                    )
                                )
                # remove span specific kwargs from cells/columns
                # that are no longer in the span,
                # cell options/col options keys are new idxs
                for k in range(span["from_c"], oldupto_colrange):
                    # has it moved outside of new span coords
                    if (isinstance(newupto, int) and (full_new_idxs[k] < newfrom or full_new_idxs[k] >= newupto)) or (
                        newupto is None and full_new_idxs[k] < newfrom
                    ):
                        # span includes header
                        if (
                            span["header"]
                            and full_new_idxs[k] in self.CH.cell_options
                            and span["type_"] in self.CH.cell_options[full_new_idxs[k]]
                        ):
                            del self.CH.cell_options[full_new_idxs[k]][span["type_"]]
                        # span is for col options
                        if span["from_r"] is None:
                            if (
                                full_new_idxs[k] in self.col_options
                                and span["type_"] in self.col_options[full_new_idxs[k]]
                            ):
                                del self.col_options[full_new_idxs[k]][span["type_"]]
                        # span is for cell options
                        else:
                            rng_upto_r = totalrows if span["upto_r"] is None else span["upto_r"]
                            for r in range(span["from_r"], rng_upto_r):
                                if (r, full_new_idxs[k]) in self.cell_options and span["type_"] in self.cell_options[
                                    (r, full_new_idxs[k])
                                ]:
                                    del self.cell_options[(r, full_new_idxs[k])][span["type_"]]
                # finally, change the span coords
                span["from_c"], span["upto_c"] = newfrom, newupto
            if data_indexes:
                self.displayed_columns = sorted(full_new_idxs[k] for k in self.displayed_columns)
        return data_new_idxs, disp_new_idxs, event_data

    def get_max_column_idx(self, maxidx: int | None = None) -> int:
        if maxidx is None:
            maxidx = len_to_idx(self.total_data_cols())
        # max column number in cell_options
        if maxidx < (maxk := max(self.cell_options, key=itemgetter(1), default=(0, 0))[1]):
            maxidx = maxk
        # max column number in column_options, index cell options
        for d in (self.col_options, self.CH.cell_options):
            if maxidx < (maxk := max(d, default=0)):
                maxidx = maxk
        # max column number in named spans
        if maxidx < (
            maxk := max(
                (d["from_c"] for d in self.named_spans.values() if isinstance(d["from_c"], int)),
                default=0,
            )
        ):
            maxidx = maxk
        if maxidx < (
            maxk := max(
                (d["upto_c"] for d in self.named_spans.values() if isinstance(d["upto_c"], int)),
                default=0,
            )
        ):
            maxidx = maxk
        return maxidx

    def get_args_for_move_rows(
        self,
        move_to: int,
        to_move: list[int],
        data_indexes: bool = False,
    ) -> tuple[dict[int, int], dict[int, int], int, dict[int, int]]:
        if not data_indexes or self.all_rows_displayed:
            disp_new_idxs = get_new_indexes(
                seqlen=len(self.row_positions) - 1,
                move_to=move_to,
                to_move=to_move,
            )
        else:
            disp_new_idxs = {}
        totalrows = self.fix_data_len(move_to)
        if self.all_rows_displayed or data_indexes:
            data_new_idxs = get_new_indexes(seqlen=totalrows, move_to=move_to, to_move=to_move)
        elif not self.all_rows_displayed and not data_indexes:
            data_new_idxs = get_new_indexes(seqlen=len(self.displayed_rows), move_to=move_to, to_move=to_move)
            data_old_idxs = dict(zip(data_new_idxs.values(), data_new_idxs))
            data_new_idxs = dict(
                zip(
                    move_elements_by_mapping(self.displayed_rows, data_new_idxs, data_old_idxs),
                    self.displayed_rows,
                )
            )
        return data_new_idxs, dict(zip(data_new_idxs.values(), data_new_idxs)), totalrows, disp_new_idxs

    def move_rows_adjust_options_dict(
        self,
        data_new_idxs: dict[int, int],
        data_old_idxs: dict[int, int],
        totalrows: int | None,
        disp_new_idxs: None | dict[int, int] = None,
        move_data: bool = True,
        create_selections: bool = True,
        data_indexes: bool = False,
        event_data: EventDataDict | None = None,
    ) -> tuple[dict[int, int], dict[int, int], EventDataDict]:
        if not isinstance(totalrows, int):
            totalrows = self.fix_data_len(max(data_new_idxs.values(), default=0))
        if event_data is None:
            event_data = event_dict(
                name="move_rows",
                sheet=self.parentframe.name,
                boxes=self.get_boxes(),
                selected=self.currently_selected(),
            )
            event_data["moved"]["rows"] = {
                "data": data_new_idxs,
                "displayed": {} if disp_new_idxs is None else disp_new_idxs,
            }
        event_data["options"] = self.pickle_options()
        event_data["named_spans"] = {k: span.pickle_self() for k, span in self.named_spans.items()}
        if disp_new_idxs and (not data_indexes or self.all_rows_displayed):
            self.deselect("all", run_binding=False, redraw=False)
            self.set_row_positions(
                itr=move_elements_by_mapping(
                    self.get_row_heights(),
                    disp_new_idxs,
                    dict(
                        zip(
                            disp_new_idxs.values(),
                            disp_new_idxs,
                        )
                    ),
                )
            )
            if create_selections:
                for chunk in consecutive_chunks(sorted(disp_new_idxs.values())):
                    self.create_selection_box(
                        chunk[0],
                        0,
                        chunk[-1] + 1,
                        len(self.col_positions) - 1,
                        "rows",
                        run_binding=True,
                    )
        if move_data:
            self.data[:] = move_elements_by_mapping(
                self.data,
                data_new_idxs,
                data_old_idxs,
            )
            maxidx = len_to_idx(totalrows)
            self.RI.fix_index(maxidx)
            if isinstance(self._row_index, list) and self._row_index:
                self._row_index = move_elements_by_mapping(self._row_index, data_new_idxs, data_old_idxs)
            maxidx = self.get_max_row_idx(maxidx)
            full_new_idxs = self.get_full_new_idxs(
                max_idx=maxidx,
                new_idxs=data_new_idxs,
                old_idxs=data_old_idxs,
            )
            full_old_idxs = dict(zip(full_new_idxs.values(), full_new_idxs))
            self.cell_options = {(full_new_idxs[k[0]], k[1]): v for k, v in self.cell_options.items()}
            self.row_options = {full_new_idxs[k]: v for k, v in self.row_options.items()}
            self.RI.cell_options = {full_new_idxs[k]: v for k, v in self.RI.cell_options.items()}
            totalcols = self.total_data_cols()
            new_ops = self.parentframe.create_options_from_span
            qkspan = self.span()
            for name, span in self.named_spans.items():
                # span is neither a cell options nor row options span, continue
                if not isinstance(span["from_r"], int):
                    continue
                oldupto_rowrange, newupto_rowrange, newfrom, newupto = span_idxs_post_move(
                    data_new_idxs,
                    full_new_idxs,
                    totalrows,
                    span,
                    "r",
                )
                # add cell/row kwargs for rows that are new to the span
                old_span_idxs = set(full_new_idxs[k] for k in range(span["from_r"], oldupto_rowrange))
                for k in range(newfrom, newupto_rowrange):
                    if k not in old_span_idxs:
                        oldidx = full_old_idxs[k]
                        # event_data is used to preserve old cell value
                        # in case cells are modified by
                        # formatting, checkboxes, dropdown boxes
                        if (
                            span["type_"] in val_modifying_options
                            and span["index"]
                            and oldidx not in event_data["cells"]["index"]
                        ):
                            event_data["cells"]["index"][oldidx] = self.RI.get_cell_data(k)
                        # the span targets rows
                        if span["from_c"] is None:
                            if span["type_"] in val_modifying_options:
                                for datacn in range(len(self.data[k])):
                                    if (oldidx, datacn) not in event_data["cells"]["table"]:
                                        event_data["cells"]["table"][(oldidx, datacn)] = self.get_cell_data(k, datacn)
                            # create new row options
                            new_ops(
                                mod_span(
                                    qkspan,
                                    span,
                                    from_r=k,
                                    upto_r=k + 1,
                                )
                            )
                        # the span targets cells
                        else:
                            rng_upto_c = totalcols if span["upto_c"] is None else span["upto_c"]
                            for datacn in range(span["from_c"], rng_upto_c):
                                if (
                                    span["type_"] in val_modifying_options
                                    and (oldidx, datacn) not in event_data["cells"]["table"]
                                ):
                                    event_data["cells"]["table"][(oldidx, datacn)] = self.get_cell_data(k, datacn)
                                # create new cell options
                                new_ops(
                                    mod_span(
                                        qkspan,
                                        span,
                                        from_r=k,
                                        upto_r=k + 1,
                                        from_c=datacn,
                                        upto_c=datacn + 1,
                                    )
                                )
                # remove span specific kwargs from cells/rows
                # that are no longer in the span,
                # cell options/row options keys are new idxs
                for k in range(span["from_r"], oldupto_rowrange):
                    # has it moved outside of new span coords
                    if (isinstance(newupto, int) and (full_new_idxs[k] < newfrom or full_new_idxs[k] >= newupto)) or (
                        newupto is None and full_new_idxs[k] < newfrom
                    ):
                        # span includes index
                        if (
                            span["index"]
                            and full_new_idxs[k] in self.RI.cell_options
                            and span["type_"] in self.RI.cell_options[full_new_idxs[k]]
                        ):
                            del self.RI.cell_options[full_new_idxs[k]][span["type_"]]
                        # span is for row options
                        if span["from_c"] is None:
                            if (
                                full_new_idxs[k] in self.row_options
                                and span["type_"] in self.row_options[full_new_idxs[k]]
                            ):
                                del self.row_options[full_new_idxs[k]][span["type_"]]
                        # span is for cell options
                        else:
                            rng_upto_c = totalcols if span["upto_c"] is None else span["upto_c"]
                            for c in range(span["from_c"], rng_upto_c):
                                if (full_new_idxs[k], c) in self.cell_options and span["type_"] in self.cell_options[
                                    (full_new_idxs[k], c)
                                ]:
                                    del self.cell_options[(full_new_idxs[k], c)][span["type_"]]
                # finally, change the span coords
                span["from_r"], span["upto_r"] = newfrom, newupto
            if data_indexes:
                self.displayed_rows = sorted(full_new_idxs[k] for k in self.displayed_rows)
        return data_new_idxs, disp_new_idxs, event_data

    def get_max_row_idx(self, maxidx: int | None = None) -> int:
        if maxidx is None:
            maxidx = len_to_idx(self.total_data_rows())
        if maxidx < (maxk := max(self.cell_options, key=itemgetter(0), default=(0, 0))[0]):
            maxidx = maxk
        # max row number in row_options, index cell options
        for d in (self.row_options, self.RI.cell_options):
            if maxidx < (maxk := max(d, default=0)):
                maxidx = maxk
        # max row number in named spans
        if maxidx < (
            maxk := max(
                (d["from_r"] for d in self.named_spans.values() if isinstance(d["from_r"], int)),
                default=0,
            )
        ):
            maxidx = maxk
        if maxidx < (
            maxk := max(
                (d["upto_r"] for d in self.named_spans.values() if isinstance(d["upto_r"], int)),
                default=0,
            )
        ):
            maxidx = maxk
        return maxidx

    def get_full_new_idxs(
        self,
        max_idx: int,
        new_idxs: dict[int, int],
        old_idxs: None | dict[int, int] = None,
    ) -> dict[int, int]:
        # return a dict of all row or column indexes
        # old indexes and new indexes, not just the
        # ones that were moved e.g.
        # {old index: new index, ...}
        # all the way from 0 to max_idx
        if old_idxs is None:
            old_idxs = dict(zip(new_idxs.values(), new_idxs))
        seq = tuple(range(max_idx + 1))
        return dict(
            zip(
                move_elements_by_mapping(seq, new_idxs, old_idxs),
                seq,
            )
        )

    def undo(self, event: object = None) -> None:
        if not self.undo_stack:
            return
        if isinstance(self.undo_stack[-1]["data"], dict):
            modification = self.undo_stack[-1]["data"]
        else:
            modification = decompress_load(self.undo_stack[-1]["data"])
        if not try_binding(self.extra_begin_ctrl_z_func, modification, "begin_undo"):
            return
        self.redo_stack.append(self.undo_modification_invert_event(modification))
        self.undo_stack.pop()
        try_binding(self.extra_end_ctrl_z_func, modification, "end_undo")

    def redo(self, event: object = None) -> None:
        if not self.redo_stack:
            return
        if isinstance(self.redo_stack[-1]["data"], dict):
            modification = self.redo_stack[-1]["data"]
        else:
            modification = decompress_load(self.redo_stack[-1]["data"])
        if not try_binding(self.extra_begin_ctrl_z_func, modification, "begin_redo"):
            return
        self.undo_stack.append(self.undo_modification_invert_event(modification, name="redo"))
        self.redo_stack.pop()
        try_binding(self.extra_end_ctrl_z_func, modification, "end_redo")

    def sheet_modified(self, event_data: EventDataDict, purge_redo: bool = True) -> None:
        self.parentframe.emit_event("<<SheetModified>>", event_data)
        if purge_redo:
            self.purge_redo_stack()

    def edit_cells_using_modification(self, modification: dict, event_data: dict) -> EventDataDict:
        for datarn, v in modification["cells"]["index"].items():
            self._row_index[datarn] = v
        for datacn, v in modification["cells"]["header"].items():
            self._headers[datacn] = v
        for (datarn, datacn), v in modification["cells"]["table"].items():
            self.set_cell_data(datarn, datacn, v)
        return event_data

    def save_cells_using_modification(self, modification: EventDataDict, event_data: EventDataDict) -> EventDataDict:
        for datarn, v in modification["cells"]["index"].items():
            event_data["cells"]["index"][datarn] = self.RI.get_cell_data(datarn)
        for datacn, v in modification["cells"]["header"].items():
            event_data["cells"]["header"][datacn] = self.CH.get_cell_data(datacn)
        for (datarn, datacn), v in modification["cells"]["table"].items():
            event_data["cells"]["table"][(datarn, datacn)] = self.get_cell_data(datarn, datacn)
        return event_data

    def restore_options_named_spans(self, modification: EventDataDict) -> None:
        if not isinstance(modification["options"], dict):
            modification["options"] = unpickle_obj(modification["options"])
        if "cell_options" in modification["options"]:
            self.cell_options = modification["options"]["cell_options"]
        if "column_options" in modification["options"]:
            self.col_options = modification["options"]["column_options"]
        if "row_options" in modification["options"]:
            self.row_options = modification["options"]["row_options"]
        if "CH_cell_options" in modification["options"]:
            self.CH.cell_options = modification["options"]["CH_cell_options"]
        if "RI_cell_options" in modification["options"]:
            self.RI.cell_options = modification["options"]["RI_cell_options"]
        self.named_spans = {k: unpickle_obj(v) for k, v in modification["named_spans"].items()}

    def undo_modification_invert_event(self, modification: EventDataDict, name: str = "undo") -> bytes | EventDataDict:
        self.deselect("all", redraw=False)
        event_data = event_dict(
            name=modification["eventname"],
            sheet=self.parentframe.name,
        )
        event_data["selection_boxes"] = modification["selection_boxes"]
        event_data["selected"] = modification["selected"]
        saved_cells = False
        curr = tuple()

        if modification["added"]["rows"] or modification["added"]["columns"]:
            event_data = self.save_cells_using_modification(modification, event_data)
            saved_cells = True

        if modification["moved"]["columns"]:
            totalcols = self.equalize_data_row_lengths()
            if totalcols < (mx_k := max(modification["moved"]["columns"]["data"].values())):
                totalcols = mx_k
            data_new_idxs, disp_new_idxs, event_data = self.move_columns_adjust_options_dict(
                data_new_idxs=dict(
                    zip(
                        modification["moved"]["columns"]["data"].values(),
                        modification["moved"]["columns"]["data"],
                    )
                ),
                data_old_idxs=modification["moved"]["columns"]["data"],
                totalcols=totalcols,
                disp_new_idxs=dict(
                    zip(
                        modification["moved"]["columns"]["displayed"].values(),
                        modification["moved"]["columns"]["displayed"],
                    )
                ),
                event_data=event_data,
            )
            event_data["moved"]["columns"] = {
                "data": data_new_idxs,
                "displayed": disp_new_idxs,
            }
            self.restore_options_named_spans(modification)
            curr = self.currently_selected()

        if modification["moved"]["rows"]:
            totalrows = self.total_data_rows()
            if totalrows < (mx_k := max(modification["moved"]["rows"]["data"].values())):
                totalrows = mx_k
            data_new_idxs, disp_new_idxs, event_data = self.move_rows_adjust_options_dict(
                data_new_idxs=dict(
                    zip(
                        modification["moved"]["rows"]["data"].values(),
                        modification["moved"]["rows"]["data"],
                    )
                ),
                data_old_idxs=modification["moved"]["rows"]["data"],
                totalrows=totalrows,
                disp_new_idxs=dict(
                    zip(
                        modification["moved"]["rows"]["displayed"].values(),
                        modification["moved"]["rows"]["displayed"],
                    )
                ),
                event_data=event_data,
            )
            event_data["moved"]["rows"] = {
                "data": data_new_idxs,
                "displayed": disp_new_idxs,
            }
            self.restore_options_named_spans(modification)
            curr = self.currently_selected()

        if modification["added"]["rows"]:
            self.deselect("all", run_binding=False, redraw=False)
            event_data = self.delete_rows_data(
                rows=tuple(reversed(modification["added"]["rows"]["table"])),
                event_data=event_data,
            )
            event_data = self.delete_rows_displayed(
                rows=tuple(reversed(modification["added"]["rows"]["row_heights"])),
                event_data=event_data,
            )
            self.displayed_rows = modification["added"]["rows"]["displayed_rows"]
            if len(self.row_positions) > 1:
                self.reselect_from_get_boxes(
                    modification["selection_boxes"],
                    modification["selected"],
                )
                curr = self.currently_selected()

        if modification["added"]["columns"]:
            self.deselect("all", run_binding=False, redraw=False)
            event_data = self.delete_columns_data(
                cols=tuple(reversed(modification["added"]["columns"]["table"])),
                event_data=event_data,
            )
            event_data = self.delete_columns_displayed(
                cols=tuple(reversed(modification["added"]["columns"]["column_widths"])),
                event_data=event_data,
            )
            self.displayed_columns = modification["added"]["columns"]["displayed_columns"]
            if len(self.col_positions) > 1:
                self.reselect_from_get_boxes(
                    modification["selection_boxes"],
                    modification["selected"],
                )
                curr = self.currently_selected()

        if modification["deleted"]["rows"]:
            self.add_rows(
                rows=modification["deleted"]["rows"],
                index=modification["deleted"]["index"],
                row_heights=modification["deleted"]["row_heights"],
                event_data=event_data,
                displayed_rows=modification["deleted"]["displayed_rows"],
                create_ops=False,
                create_selections=not modification["eventname"].startswith("edit"),
                add_col_positions=False,
            )
            self.restore_options_named_spans(modification)

        if modification["deleted"]["columns"]:
            self.add_columns(
                columns=modification["deleted"]["columns"],
                header=modification["deleted"]["header"],
                column_widths=modification["deleted"]["column_widths"],
                event_data=event_data,
                displayed_columns=modification["deleted"]["displayed_columns"],
                create_ops=False,
                create_selections=not modification["eventname"].startswith("edit"),
                add_row_positions=False,
            )
            self.restore_options_named_spans(modification)

        if modification["eventname"].startswith("edit") and (
            modification["deleted"]["columns"] or modification["deleted"]["rows"]
        ):
            self.reselect_from_get_boxes(
                modification["selection_boxes"],
                modification["selected"],
            )
            curr = self.currently_selected()

        if modification["eventname"].startswith(("edit", "move")):
            if not saved_cells:
                event_data = self.save_cells_using_modification(modification, event_data)
            event_data = self.edit_cells_using_modification(modification, event_data)
            if (
                not modification["deleted"]["columns"]
                and not modification["deleted"]["rows"]
                and not modification["eventname"].startswith("move")
            ):
                self.reselect_from_get_boxes(
                    modification["selection_boxes"],
                    modification["selected"],
                )
            curr = self.currently_selected()

        elif modification["eventname"].startswith("add"):
            event_data["eventname"] = modification["eventname"].replace("add", "delete")

        elif modification["eventname"].startswith("delete"):
            event_data["eventname"] = modification["eventname"].replace("delete", "add")

        if curr:
            self.see(
                r=curr.row,
                c=curr.column,
                keep_yscroll=False,
                keep_xscroll=False,
                bottom_right_corner=False,
                check_cell_visibility=True,
                redraw=False,
            )

        self.sheet_modified(event_data, purge_redo=False)
        self.refresh()
        return ev_stack_dict(event_data)

    def bind_arrowkeys(self, keys: dict = {}) -> None:
        for canvas in (self, self.parentframe, self.CH, self.RI, self.TL):
            for k, func in keys.items():
                canvas.bind(f"<{arrowkey_bindings_helper[k.lower()]}>", func)

    def unbind_arrowkeys(self, keys: dict = {}) -> None:
        for canvas in (self, self.parentframe, self.CH, self.RI, self.TL):
            for k, func in keys.items():
                canvas.unbind(f"<{arrowkey_bindings_helper[k.lower()]}>")

    def see(
        self,
        r: int | None = None,
        c: int | None = None,
        keep_yscroll: bool = False,
        keep_xscroll: bool = False,
        bottom_right_corner: bool = False,
        check_cell_visibility: bool = True,
        redraw: bool = True,
        r_pc: float = 0.0,
        c_pc: float = 0.0,
    ) -> bool:
        need_redraw = False
        yvis, xvis = False, False
        if check_cell_visibility:
            yvis, xvis = self.cell_completely_visible(r=r, c=c, separate_axes=True)
        if not yvis and len(self.row_positions) > 1:
            if bottom_right_corner:
                if r is not None and not keep_yscroll:
                    winfo_height = self.winfo_height()
                    if self.row_positions[r + 1] - self.row_positions[r] > winfo_height:
                        y = self.row_positions[r]
                    else:
                        y = self.row_positions[r + 1] + 1 - winfo_height
                    args = [
                        "moveto",
                        y / (self.row_positions[-1] + self.empty_vertical),
                    ]
                    if args[1] > 1:
                        args[1] = args[1] - 1
                    self.set_yviews(*args, redraw=False)
                    need_redraw = True
            else:
                if r is not None and not keep_yscroll:
                    y = self.row_positions[r] + ((self.row_positions[r + 1] - self.row_positions[r]) * r_pc)
                    args = [
                        "moveto",
                        y / (self.row_positions[-1] + self.empty_vertical),
                    ]
                    if args[1] > 1:
                        args[1] = args[1] - 1
                    self.set_yviews(*args, redraw=False)
                    need_redraw = True
        if not xvis and len(self.col_positions) > 1:
            if bottom_right_corner:
                if c is not None and not keep_xscroll:
                    winfo_width = self.winfo_width()
                    if self.col_positions[c + 1] - self.col_positions[c] > winfo_width:
                        x = self.col_positions[c]
                    else:
                        x = self.col_positions[c + 1] + 1 - winfo_width
                    args = [
                        "moveto",
                        x / (self.col_positions[-1] + self.empty_horizontal),
                    ]
                    self.set_xviews(*args, redraw=False)
                    need_redraw = True
            else:
                if c is not None and not keep_xscroll:
                    x = self.col_positions[c] + ((self.col_positions[c + 1] - self.col_positions[c]) * c_pc)
                    args = [
                        "moveto",
                        x / (self.col_positions[-1] + self.empty_horizontal),
                    ]
                    self.set_xviews(*args, redraw=False)
                    need_redraw = True
        if redraw and need_redraw:
            self.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
            return True
        return False

    def get_cell_coords(self, r: int | None = None, c: int | None = None) -> tuple[float, float, float, float]:
        return (
            0 if not c else self.col_positions[c] + 1,
            0 if not r else self.row_positions[r] + 1,
            0 if not c else self.col_positions[c + 1],
            0 if not r else self.row_positions[r + 1],
        )

    def cell_completely_visible(self, r: int = 0, c: int = 0, separate_axes: bool = False) -> bool:
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
        if not y_vis or not x_vis:
            return False
        return True

    def cell_visible(self, r: int = 0, c: int = 0) -> bool:
        cx1, cy1, cx2, cy2 = self.get_canvas_visible_area()
        x1, y1, x2, y2 = self.get_cell_coords(r, c)
        if x1 <= cx2 or y1 <= cy2 or x2 >= cx1 or y2 >= cy1:
            return True
        return False

    def select_all(self, redraw: bool = True, run_binding_func: bool = True) -> None:
        currently_selected = self.currently_selected()
        self.deselect("all", redraw=False)
        if len(self.row_positions) > 1 and len(self.col_positions) > 1:
            item = self.create_selection_box(
                0,
                0,
                len(self.row_positions) - 1,
                len(self.col_positions) - 1,
                set_current=False,
            )
            if currently_selected:
                self.set_currently_selected(currently_selected.row, currently_selected.column, item=item)
            else:
                self.set_currently_selected(0, 0, item=item)
            if redraw:
                self.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
            if self.select_all_binding_func and run_binding_func:
                self.select_all_binding_func(
                    self.get_select_event(being_drawn_item=self.being_drawn_item),
                )

    def select_cell(
        self,
        r: int,
        c: int,
        redraw: bool = False,
        run_binding_func: bool = True,
    ) -> int:
        self.deselect("all", redraw=False)
        fill_iid = self.create_selection_box(r, c, r + 1, c + 1, state="hidden")
        if redraw:
            self.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
        if run_binding_func:
            self.run_selection_binding("cells")
        return fill_iid

    def add_selection(
        self,
        r: int,
        c: int,
        redraw: bool = False,
        run_binding_func: bool = True,
        set_as_current: bool = False,
    ) -> int:
        fill_iid = self.create_selection_box(r, c, r + 1, c + 1, state="hidden", set_current=set_as_current)
        if redraw:
            self.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
        if run_binding_func:
            self.run_selection_binding("cells")
        return fill_iid

    def toggle_select_cell(
        self,
        row: int,
        column: int,
        add_selection: bool = True,
        redraw: bool = True,
        run_binding_func: bool = True,
        set_as_current: bool = True,
    ) -> int:
        if add_selection:
            if self.cell_selected(row, column, inc_rows=True, inc_cols=True):
                fill_iid = self.deselect(r=row, c=column, redraw=redraw)
            else:
                fill_iid = self.add_selection(
                    r=row,
                    c=column,
                    redraw=redraw,
                    run_binding_func=run_binding_func,
                    set_as_current=set_as_current,
                )
        else:
            if self.cell_selected(row, column, inc_rows=True, inc_cols=True):
                fill_iid = self.deselect(r=row, c=column, redraw=redraw)
            else:
                fill_iid = self.select_cell(row, column, redraw=redraw)
        return fill_iid

    def get_select_event(self, being_drawn_item: None | int) -> EventDataDict:
        return event_dict(
            name="select",
            sheet=self.parentframe.name,
            selected=self.currently_selected(),
            being_selected=self.get_box_from_item(being_drawn_item),
            boxes=self.get_boxes(),
        )

    def deselect(
        self,
        r: int | None | str = None,
        c: int | None = None,
        cell: tuple[int, int] | None = None,
        redraw: bool = True,
        run_binding: bool = True,
    ) -> None:
        if not self.anything_selected():
            return
        # saved_current = self.currently_selected()
        set_curr = False
        current = self.currently_selected().tags
        if r == "all" or (r is None and c is None and cell is None):
            for item in self.get_selection_items(current=False):
                tags = self.gettags(item)
                self.delete_item(item)
        elif r in ("allrows", "allcols"):
            for item in self.get_selection_items(
                columns=r == "allcols", rows=r == "allrows", cells=False, current=False
            ):
                tags = self.gettags(item)
                r1, c1, r2, c2 = coords_tag_to_int_tuple(tags[1])
                self.delete_item(item)
                if current[2] == tags[2]:
                    set_curr = True
        elif r is not None and c is None and cell is None:
            for item in self.get_selection_items(columns=False, cells=False, current=False):
                tags = self.gettags(item)
                r1, c1, r2, c2 = coords_tag_to_int_tuple(tags[1])
                if r >= r1 and r < r2:
                    self.delete_item(item)
                    if current[2] == tags[2]:
                        set_curr = True
                    if r2 - r1 != 1:
                        if r == r1:
                            self.create_selection_box(
                                r1 + 1,
                                0,
                                r2,
                                len(self.col_positions) - 1,
                                "rows",
                                set_current=False,
                            )
                        elif r == r2 - 1:
                            self.create_selection_box(
                                r1,
                                0,
                                r2 - 1,
                                len(self.col_positions) - 1,
                                "rows",
                                set_current=False,
                            )
                        else:
                            self.create_selection_box(
                                r1,
                                0,
                                r,
                                len(self.col_positions) - 1,
                                "rows",
                                set_current=False,
                            )
                            self.create_selection_box(
                                r + 1,
                                0,
                                r2,
                                len(self.col_positions) - 1,
                                "rows",
                                set_current=False,
                            )
                    break
        elif c is not None and r is None and cell is None:
            for item in self.get_selection_items(rows=False, cells=False, current=False):
                tags = self.gettags(item)
                r1, c1, r2, c2 = coords_tag_to_int_tuple(tags[1])
                if c >= c1 and c < c2:
                    self.delete_item(item)
                    if current[2] == tags[2]:
                        set_curr = True
                    if c2 - c1 != 1:
                        if c == c1:
                            self.create_selection_box(
                                0,
                                c1 + 1,
                                len(self.row_positions) - 1,
                                c2,
                                "columns",
                                set_current=False,
                            )
                        elif c == c2 - 1:
                            self.create_selection_box(
                                0,
                                c1,
                                len(self.row_positions) - 1,
                                c2 - 1,
                                "columns",
                                set_current=False,
                            )
                        else:
                            self.create_selection_box(
                                0,
                                c1,
                                len(self.row_positions) - 1,
                                c,
                                "columns",
                                set_current=False,
                            )
                            self.create_selection_box(
                                0,
                                c + 1,
                                len(self.row_positions) - 1,
                                c2,
                                "columns",
                                set_current=False,
                            )
                    break
        elif (r is not None and c is not None and cell is None) or cell is not None:
            if cell is not None:
                r, c = cell[0], cell[1]
            for item in self.get_selection_items(current=False, reverse=True):
                tags = self.gettags(item)
                r1, c1, r2, c2 = coords_tag_to_int_tuple(tags[1])
                if r >= r1 and c >= c1 and r < r2 and c < c2:
                    self.delete_item(item)
                    if current[2] == tags[2]:
                        set_curr = True
                    break
        if set_curr:
            self.set_current_to_last()
        if redraw:
            self.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
        if run_binding:
            try_binding(
                self.deselection_binding_func,
                self.get_select_event(being_drawn_item=self.being_drawn_item),
            )

    def page_UP(self, event=None):
        height = self.winfo_height()
        top = self.canvasy(0)
        scrollto = top - height
        if scrollto < 0:
            scrollto = 0
        if self.page_up_down_select_row:
            r = bisect_left(self.row_positions, scrollto)
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
            r = bisect_left(self.row_positions, scrollto) - 1
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
                        widget.bind(f"<{ctrl_key}-{s2}>", self.undo)
                        widget.bind(f"<{ctrl_key}-Shift-{s2}>", self.redo)
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
                command=self.rc_delete_columns,
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
                command=lambda: self.rc_add_columns("left"),
            )
            self.menu_add_command(
                self.empty_rc_popup_menu,
                label="Insert column",
                font=self.popup_menu_font,
                foreground=self.popup_menu_fg,
                background=self.popup_menu_bg,
                activebackground=self.popup_menu_highlight_bg,
                activeforeground=self.popup_menu_highlight_fg,
                command=lambda: self.rc_add_columns("left"),
            )
            self.menu_add_command(
                self.CH.ch_rc_popup_menu,
                label="Insert columns right",
                font=self.popup_menu_font,
                foreground=self.popup_menu_fg,
                background=self.popup_menu_bg,
                activebackground=self.popup_menu_highlight_bg,
                activeforeground=self.popup_menu_highlight_fg,
                command=lambda: self.rc_add_columns("right"),
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
                command=self.rc_delete_rows,
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
                command=lambda: self.rc_add_rows("above"),
            )
            self.menu_add_command(
                self.RI.ri_rc_popup_menu,
                label="Insert rows below",
                font=self.popup_menu_font,
                foreground=self.popup_menu_fg,
                background=self.popup_menu_bg,
                activebackground=self.popup_menu_highlight_bg,
                activeforeground=self.popup_menu_highlight_fg,
                command=lambda: self.rc_add_rows("below"),
            )
            self.menu_add_command(
                self.empty_rc_popup_menu,
                label="Insert row",
                font=self.popup_menu_font,
                foreground=self.popup_menu_fg,
                background=self.popup_menu_bg,
                activebackground=self.popup_menu_highlight_bg,
                activeforeground=self.popup_menu_highlight_fg,
                command=lambda: self.rc_add_rows("below"),
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
            self.TL.sa_state()
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
            self.TL.sa_state()
        elif binding == "column_width_resize":
            self.CH.enable_bindings("column_width_resize")
        elif binding == "column_select":
            self.CH.enable_bindings("column_select")
        elif binding == "column_height_resize":
            self.CH.enable_bindings("column_height_resize")
            self.TL.rh_state()
        elif binding in ("column_drag_and_drop", "move_columns"):
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
        elif binding in ("row_drag_and_drop", "move_rows"):
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
            self.TL.sa_state("hidden")
        elif binding in ("single", "single_selection_mode", "single_select"):
            self.single_selection_enabled = False
        elif binding in ("toggle", "toggle_selection_mode", "toggle_select"):
            self.toggle_selection_enabled = False
        elif binding == "drag_select":
            self.drag_selection_enabled = False
        elif binding == "select_all":
            self.enable_disable_select_all(False)
            self.TL.sa_state("hidden")
        elif binding == "column_width_resize":
            self.CH.disable_bindings("column_width_resize")
        elif binding == "column_select":
            self.CH.disable_bindings("column_select")
        elif binding == "column_height_resize":
            self.CH.disable_bindings("column_height_resize")
            self.TL.rh_state("hidden")
        elif binding in ("column_drag_and_drop", "move_columns"):
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
        elif binding in ("row_drag_and_drop", "move_rows"):
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

    def reset_mouse_motion_creations(self) -> None:
        if self.current_cursor != "":
            self.config(cursor="")
            self.RI.config(cursor="")
            self.CH.config(cursor="")
            self.current_cursor = ""
        self.RI.rsz_w = None
        self.RI.rsz_h = None
        self.CH.rsz_w = None
        self.CH.rsz_h = None

    def mouse_motion(self, event: object):
        self.reset_mouse_motion_creations()
        if self.extra_motion_func:
            self.extra_motion_func(event)

    def not_currently_resizing(self) -> bool:
        return all(v is None for v in (self.RI.rsz_h, self.RI.rsz_w, self.CH.rsz_h, self.CH.rsz_w))

    def rc(self, event=None):
        self.mouseclick_outside_editor_or_dropdown_all_canvases()
        self.focus_set()
        popup_menu = None
        if self.single_selection_enabled and self.not_currently_resizing():
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
        elif self.toggle_selection_enabled and self.not_currently_resizing():
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
        if self.extra_rc_func:
            self.extra_rc_func(event)
        if popup_menu is not None:
            popup_menu.tk_popup(event.x_root, event.y_root)

    def b1_press(self, event=None):
        self.closed_dropdown = self.mouseclick_outside_editor_or_dropdown_all_canvases()
        self.focus_set()
        if (
            self.identify_col(x=event.x, allow_end=False) is None
            or self.identify_row(y=event.y, allow_end=False) is None
        ):
            self.deselect("all")
        r = self.identify_row(y=event.y)
        c = self.identify_col(x=event.x)
        if self.single_selection_enabled and self.not_currently_resizing():
            if r < len(self.row_positions) - 1 and c < len(self.col_positions) - 1:
                self.being_drawn_item = self.select_cell(r, c, redraw=True)
        elif self.toggle_selection_enabled and self.not_currently_resizing():
            r = self.identify_row(y=event.y)
            c = self.identify_col(x=event.x)
            if r < len(self.row_positions) - 1 and c < len(self.col_positions) - 1:
                self.toggle_select_cell(r, c, redraw=True)
        self.b1_pressed_loc = (r, c)
        if self.extra_b1_press_func:
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
        if self.ctrl_select_enabled and self.not_currently_resizing():
            self.b1_pressed_loc = None
            rowsel = int(self.identify_row(y=event.y))
            colsel = int(self.identify_col(x=event.x))
            if rowsel < len(self.row_positions) - 1 and colsel < len(self.col_positions) - 1:
                if self.cell_selected(rowsel, colsel):
                    self.deselect(rowsel, colsel)
                else:
                    self.being_drawn_item = True
                    self.being_drawn_item = self.add_selection(
                        rowsel, colsel, set_as_current=True, run_binding_func=False
                    )
                    if self.ctrl_selection_binding_func:
                        self.ctrl_selection_binding_func(
                            self.get_select_event(being_drawn_item=self.being_drawn_item),
                        )
                self.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True, redraw_table=True)

        elif not self.ctrl_select_enabled:
            self.b1_press(event)

    def ctrl_shift_b1_press(self, event=None):
        self.mouseclick_outside_editor_or_dropdown_all_canvases()
        self.focus_set()
        if self.ctrl_select_enabled and self.drag_selection_enabled and self.not_currently_resizing():
            self.b1_pressed_loc = None
            rowsel = int(self.identify_row(y=event.y))
            colsel = int(self.identify_col(x=event.x))
            if rowsel < len(self.row_positions) - 1 and colsel < len(self.col_positions) - 1:
                currently_selected = self.currently_selected()
                if currently_selected:
                    self.delete_item(currently_selected.tags[2])
                box = self.get_shift_select_box(currently_selected.row, rowsel, currently_selected.column, colsel)
                if currently_selected and currently_selected.type_ == "cell":
                    self.being_drawn_item = self.create_selection_box(*box, set_current=currently_selected)
                else:
                    self.being_drawn_item = self.add_selection(
                        rowsel, colsel, set_as_current=True, run_binding_func=False
                    )
                self.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True, redraw_table=True)
                if self.shift_selection_binding_func:
                    self.shift_selection_binding_func(
                        self.get_select_event(being_drawn_item=self.being_drawn_item),
                    )
        elif not self.ctrl_select_enabled:
            self.shift_b1_press(event)

    def shift_b1_press(self, event=None):
        self.mouseclick_outside_editor_or_dropdown_all_canvases()
        self.focus_set()
        if self.drag_selection_enabled and self.not_currently_resizing():
            self.b1_pressed_loc = None
            rowsel = int(self.identify_row(y=event.y))
            colsel = int(self.identify_col(x=event.x))
            if rowsel < len(self.row_positions) - 1 and colsel < len(self.col_positions) - 1:
                currently_selected = self.currently_selected()
                box = self.get_shift_select_box(currently_selected.row, rowsel, currently_selected.column, colsel)
                if currently_selected and currently_selected.type_ == "cell":
                    self.deselect("all", redraw=False)
                    self.being_drawn_item = self.create_selection_box(*box, set_current=currently_selected)
                else:
                    self.being_drawn_item = self.select_cell(rowsel, colsel, redraw=False, run_binding_func=False)
                self.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True, redraw_table=True)
                if self.shift_selection_binding_func:
                    self.shift_selection_binding_func(
                        self.get_select_event(being_drawn_item=self.being_drawn_item),
                    )

    def get_shift_select_box(self, min_r: int, rowsel: int, min_c: int, colsel: int):
        if rowsel >= min_r and colsel >= min_c:
            return min_r, min_c, rowsel + 1, colsel + 1, "cells"
        elif rowsel >= min_r and min_c >= colsel:
            return min_r, colsel, rowsel + 1, min_c + 1, "cells"
        elif min_r >= rowsel and colsel >= min_c:
            return rowsel, min_c, min_r + 1, colsel + 1, "cells"
        elif min_r >= rowsel and min_c >= colsel:
            return rowsel, colsel, min_r + 1, min_c + 1, "cells"

    def get_b1_motion_box(self, start_row: int, start_col: int, end_row: int, end_col: int):
        if end_row >= start_row and end_col >= start_col and (end_row - start_row or end_col - start_col):
            return start_row, start_col, end_row + 1, end_col + 1, "cells"
        elif end_row >= start_row and end_col < start_col and (end_row - start_row or start_col - end_col):
            return start_row, end_col, end_row + 1, start_col + 1, "cells"
        elif end_row < start_row and end_col >= start_col and (start_row - end_row or end_col - start_col):
            return end_row, start_col, start_row + 1, end_col + 1, "cells"
        elif end_row < start_row and end_col < start_col and (start_row - end_row or start_col - end_col):
            return end_row, end_col, start_row + 1, start_col + 1, "cells"
        else:
            return start_row, start_col, start_row + 1, start_col + 1, "cells"

    def b1_motion(self, event: object):
        if self.drag_selection_enabled and all(
            v is None
            for v in (
                self.RI.rsz_h,
                self.RI.rsz_w,
                self.CH.rsz_h,
                self.CH.rsz_w,
            )
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
                box = self.get_b1_motion_box(
                    *(
                        currently_selected.row,
                        currently_selected.column,
                        end_row,
                        end_col,
                    )
                )
                if (
                    box is not None
                    and self.being_drawn_item is not None
                    and self.get_box_from_item(self.being_drawn_item) != box
                ):
                    self.deselect("all", redraw=False)
                    if box[2] - box[0] != 1 or box[3] - box[1] != 1:
                        self.being_drawn_item = self.create_selection_box(*box, set_current=currently_selected)
                    else:
                        self.being_drawn_item = self.select_cell(
                            currently_selected.row, currently_selected.column, run_binding_func=False
                        )
                    need_redraw = True
                    if self.drag_selection_binding_func:
                        self.drag_selection_binding_func(
                            self.get_select_event(being_drawn_item=self.being_drawn_item),
                        )
            if self.scroll_if_event_offscreen(event):
                need_redraw = True
            if need_redraw:
                self.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True, redraw_table=True)
        if self.extra_b1_motion_func:
            self.extra_b1_motion_func(event)

    def ctrl_b1_motion(self, event: object):
        if self.ctrl_select_enabled and self.drag_selection_enabled and self.not_currently_resizing():
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
                box = self.get_b1_motion_box(
                    *(
                        currently_selected.row,
                        currently_selected.column,
                        end_row,
                        end_col,
                    )
                )
                if (
                    box is not None
                    and self.being_drawn_item is not None
                    and self.get_box_from_item(self.being_drawn_item) != box
                ):
                    self.delete_item(self.being_drawn_item)
                    if box[2] - box[0] != 1 or box[3] - box[1] != 1:
                        self.being_drawn_item = self.create_selection_box(*box, set_current=currently_selected)
                    else:
                        self.being_drawn_item = self.add_selection(
                            currently_selected.row,
                            currently_selected.column,
                            set_as_current=True,
                        )
                    need_redraw = True
                    if self.drag_selection_binding_func:
                        self.drag_selection_binding_func(
                            self.get_select_event(being_drawn_item=self.being_drawn_item),
                        )
            if self.scroll_if_event_offscreen(event):
                need_redraw = True
            if need_redraw:
                self.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True, redraw_table=True)
        elif not self.ctrl_select_enabled:
            self.b1_motion(event)

    def b1_release(self, event=None):
        if self.being_drawn_item is not None:
            currently_selected = self.currently_selected()
            to_sel = self.get_box_from_item(self.being_drawn_item)
            self.delete_item(self.being_drawn_item)
            self.being_drawn_item = None
            self.create_selection_box(
                *to_sel,
                state="hidden" if to_sel[2] - to_sel[0] == 1 and to_sel[3] - to_sel[1] == 1 else "normal",
                set_current=currently_selected,
            )
            if self.drag_selection_binding_func:
                self.drag_selection_binding_func(
                    self.get_select_event(being_drawn_item=self.being_drawn_item),
                )
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
                datarn = self.datarn(r)
                datacn = self.datacn(c)
                if self.get_cell_kwargs(datarn, datacn, key="dropdown") or self.get_cell_kwargs(
                    datarn, datacn, key="checkbox"
                ):
                    canvasx = self.canvasx(event.x)
                    if (
                        self.closed_dropdown != self.b1_pressed_loc
                        and self.get_cell_kwargs(datarn, datacn, key="dropdown")
                        and canvasx > self.col_positions[c + 1] - self.table_txt_height - 4
                        and canvasx < self.col_positions[c + 1] - 1
                    ) or (
                        self.get_cell_kwargs(datarn, datacn, key="checkbox")
                        and canvasx < self.col_positions[c] + self.table_txt_height + 4
                        and self.canvasy(event.y) < self.row_positions[r] + self.table_txt_height + 4
                    ):
                        self.open_cell(event)
            else:
                self.mouseclick_outside_editor_or_dropdown_all_canvases()
        self.b1_pressed_loc = None
        self.closed_dropdown = None
        if self.extra_b1_release_func:
            self.extra_b1_release_func(event)

    def double_b1(self, event=None):
        self.mouseclick_outside_editor_or_dropdown_all_canvases()
        self.focus_set()
        if (
            self.identify_col(x=event.x, allow_end=False) is None
            or self.identify_row(y=event.y, allow_end=False) is None
        ):
            self.deselect("all")
        elif self.single_selection_enabled and self.not_currently_resizing():
            r = self.identify_row(y=event.y)
            c = self.identify_col(x=event.x)
            if r < len(self.row_positions) - 1 and c < len(self.col_positions) - 1:
                self.select_cell(r, c, redraw=True)
                if self.edit_cell_enabled:
                    self.open_cell(event)
        elif self.toggle_selection_enabled and self.not_currently_resizing():
            r = self.identify_row(y=event.y)
            c = self.identify_col(x=event.x)
            if r < len(self.row_positions) - 1 and c < len(self.col_positions) - 1:
                self.toggle_select_cell(r, c, redraw=True)
                if self.edit_cell_enabled:
                    self.open_cell(event)
        if self.extra_double_b1_func:
            self.extra_double_b1_func(event)

    def identify_row(self, event=None, y=None, allow_end=True):
        if event is None:
            y2 = self.canvasy(y)
        elif y is None:
            y2 = self.canvasy(event.y)
        r = bisect_left(self.row_positions, y2)
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
        c = bisect_left(self.col_positions, x2)
        if c != 0:
            c -= 1
        if not allow_end and c >= len(self.col_positions) - 1:
            return None
        return c

    def fix_views(self) -> None:
        xcheck = self.xview()
        ycheck = self.yview()
        if xcheck and xcheck[0] <= 0:
            self.xview("moveto", 0)
            if self.show_header:
                self.CH.xview("moveto", 0)
        elif len(xcheck) > 1 and xcheck[1] >= 1:
            self.xview("moveto", 1)
            if self.show_header:
                self.CH.xview("moveto", 1)
        if ycheck and ycheck[0] <= 0:
            self.yview("moveto", 0)
            if self.show_index:
                self.RI.yview("moveto", 0)
        elif len(ycheck) > 1 and ycheck[1] >= 1:
            self.yview("moveto", 1)
            if self.show_index:
                self.RI.yview("moveto", 1)

    def scroll_if_event_offscreen(self, event: object) -> bool:
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
            self.x_move_synced_scrolls("moveto", self.xview()[0])
            self.y_move_synced_scrolls("moveto", self.yview()[0])
        return need_redraw

    def x_move_synced_scrolls(self, *args, redraw: bool = True):
        for widget in self.synced_scrolls:
            # try:
            if hasattr(widget, "MT"):
                widget.MT.set_xviews(*args, move_synced=False, redraw=redraw)
            else:
                widget.xview(*args)
            # except Exception:
            #     continue

    def y_move_synced_scrolls(self, *args, redraw: bool = True):
        for widget in self.synced_scrolls:
            # try:
            if hasattr(widget, "MT"):
                widget.MT.set_yviews(*args, move_synced=False, redraw=redraw)
            else:
                widget.yview(*args)
            # except Exception:
            #     continue

    def set_xviews(self, *args, move_synced: bool = True, redraw: bool = True) -> None:
        self.xview(*args)
        if self.show_header:
            self.CH.xview(*args)
        if move_synced:
            self.x_move_synced_scrolls(*args)
        self.fix_views()
        if redraw:
            self.main_table_redraw_grid_and_text(redraw_header=True if self.show_header else False)

    def set_yviews(self, *args, move_synced: bool = True, redraw: bool = True) -> None:
        self.yview(*args)
        if self.show_index:
            self.RI.yview(*args)
        if move_synced:
            self.y_move_synced_scrolls(*args)
        self.fix_views()
        if redraw:
            self.main_table_redraw_grid_and_text(redraw_row_index=True if self.show_index else False)

    def set_view(self, x_args: list[str, float], y_args: list[str, float]) -> None:
        self.xview(*x_args)
        if self.show_header:
            self.CH.xview(*x_args)
        self.yview(*y_args)
        if self.show_index:
            self.RI.yview(*y_args)
        self.x_move_synced_scrolls(*x_args)
        self.y_move_synced_scrolls(*y_args)
        self.fix_views()
        self.main_table_redraw_grid_and_text(
            redraw_row_index=True if self.show_index else False,
            redraw_header=True if self.show_header else False,
        )

    def mousewheel(self, event: object) -> None:
        if event.delta < 0 or event.num == 5:
            self.yview_scroll(1, "units")
            self.RI.yview_scroll(1, "units")
            self.y_move_synced_scrolls("moveto", self.yview()[0])
        elif event.delta >= 0 or event.num == 4:
            if self.canvasy(0) <= 0:
                return
            self.yview_scroll(-1, "units")
            self.RI.yview_scroll(-1, "units")
            self.y_move_synced_scrolls("moveto", self.yview()[0])
        self.main_table_redraw_grid_and_text(redraw_row_index=True)

    def shift_mousewheel(self, event: object) -> None:
        if event.delta < 0 or event.num == 5:
            self.xview_scroll(1, "units")
            self.CH.xview_scroll(1, "units")
            self.x_move_synced_scrolls("moveto", self.xview()[0])
        elif event.delta >= 0 or event.num == 4:
            if self.canvasx(0) <= 0:
                return
            self.xview_scroll(-1, "units")
            self.CH.xview_scroll(-1, "units")
            self.x_move_synced_scrolls("moveto", self.xview()[0])
        self.main_table_redraw_grid_and_text(redraw_header=True)

    def ctrl_mousewheel(self, event):
        if event.delta < 0 or event.num == 5:
            if self.table_font[1] < 2 or self.index_font[1] < 2 or self.header_font[1] < 2:
                return
            self.zoom_out()
        elif event.delta >= 0 or event.num == 4:
            self.zoom_in()

    def zoom_in(self, event=None):
        self.zoom_font(
            (self.table_font[0], self.table_font[1] + 1, self.table_font[2]),
            (self.header_font[0], self.header_font[1] + 1, self.header_font[2]),
        )

    def zoom_out(self, event=None):
        self.zoom_font(
            (self.table_font[0], self.table_font[1] - 1, self.table_font[2]),
            (self.header_font[0], self.header_font[1] - 1, self.header_font[2]),
        )

    def zoom_font(self, table_font: tuple, header_font: tuple):
        # should record position prior to change and then see after change
        y = self.canvasy(0)
        x = self.canvasx(0)
        r = self.identify_row(y=0)
        c = self.identify_col(x=0)
        try:
            r_pc = (y - self.row_positions[r]) / (self.row_positions[r + 1] - self.row_positions[r])
        except Exception:
            r_pc = 0.0
        try:
            c_pc = (x - self.col_positions[c]) / (self.col_positions[c + 1] - self.col_positions[c])
        except Exception:
            c_pc = 0.0
        old_min_row_height = int(self.min_row_height)
        old_default_row_height = int(self.default_row_height[1])
        self.set_table_font(
            table_font,
            reset_row_positions=False,
        )
        self.set_index_font(table_font)
        self.set_header_font(header_font)
        if self.set_cell_sizes_on_zoom:
            self.set_all_cell_sizes_to_text()
            self.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
        elif not self.set_cell_sizes_on_zoom:
            self.row_positions = list(
                accumulate(
                    chain(
                        [0],
                        (
                            self.min_row_height
                            if h == old_min_row_height
                            else self.default_row_height[1]
                            if h == old_default_row_height
                            else self.min_row_height
                            if h < self.min_row_height
                            else h
                            for h in self.gen_row_heights()
                        ),
                    )
                )
            )
            self.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)
            self.recreate_all_selection_boxes()
        self.refresh_open_window_positions()
        self.RI.refresh_open_window_positions()
        self.CH.refresh_open_window_positions()
        self.see(
            r,
            c,
            check_cell_visibility=False,
            redraw=False,
            r_pc=r_pc,
            c_pc=c_pc,
        )
        self.main_table_redraw_grid_and_text(redraw_header=True, redraw_row_index=True)

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
                txt="\n".join("|" for lines in range(n)) if n > 1 else "|",
                font=self.table_font if font is None else font,
            )
            + 5
        )

    def set_min_column_width(self):
        self.min_column_width = 1
        if self.min_column_width > self.max_column_width:
            self.max_column_width = self.min_column_width + 20
        if self.min_column_width > self.default_column_width:
            self.default_column_width = self.min_column_width + 20
        if isinstance(self.auto_resize_columns, (int, float)) and self.auto_resize_columns < self.min_column_width:
            self.auto_resize_columns = self.min_column_width

    def set_table_font(self, newfont: tuple | None = None, reset_row_positions: bool = False) -> tuple[str, int, str]:
        if newfont:
            if not isinstance(newfont, tuple):
                raise ValueError("Argument must be tuple e.g. " "('Carlito', 12, 'normal')")
            if len(newfont) != 3:
                raise ValueError("Argument must be three-tuple")
            if not isinstance(newfont[0], str) or not isinstance(newfont[1], int) or not isinstance(newfont[2], str):
                raise ValueError(
                    "Argument must be font, size and 'normal', 'bold' or" "'italic' e.g. ('Carlito',12,'normal')"
                )
            self.table_font = newfont
            self.set_table_font_help()
            if reset_row_positions:
                if isinstance(reset_row_positions, bool):
                    self.reset_row_positions()
                else:
                    self.set_row_positions(itr=reset_row_positions)
                self.recreate_all_selection_boxes()
        return self.table_font

    def set_table_font_help(self):
        self.table_txt_width, self.table_txt_height = self.get_txt_dimensions("|", self.table_font)
        self.table_half_txt_height = ceil(self.table_txt_height / 2)
        if self.table_half_txt_height % 2 == 0:
            self.table_first_ln_ins = self.table_half_txt_height + 2
        else:
            self.table_first_ln_ins = self.table_half_txt_height + 3
        self.table_xtra_lines_increment = int(self.table_txt_height)
        self.min_row_height = self.table_txt_height + 5
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

    def set_header_font(self, newfont: tuple | None = None) -> tuple[str, int, str]:
        if newfont:
            if not isinstance(newfont, tuple):
                raise ValueError("Argument must be tuple e.g. ('Carlito', 12, 'normal')")
            if len(newfont) != 3:
                raise ValueError("Argument must be three-tuple")
            if not isinstance(newfont[0], str) or not isinstance(newfont[1], int) or not isinstance(newfont[2], str):
                raise ValueError(
                    "Argument must be font, size and 'normal', 'bold' or" "'italic' e.g. ('Carlito',12,'normal')"
                )
            self.header_font = newfont
            self.set_header_font_help()
            self.recreate_all_selection_boxes()
        return self.header_font

    def set_header_font_help(self):
        self.header_txt_width, self.header_txt_height = self.get_txt_dimensions("|", self.header_font)
        self.header_half_txt_height = ceil(self.header_txt_height / 2)
        if self.header_half_txt_height % 2 == 0:
            self.header_first_ln_ins = self.header_half_txt_height + 2
        else:
            self.header_first_ln_ins = self.header_half_txt_height + 3
        self.header_xtra_lines_increment = self.header_txt_height
        self.min_header_height = self.header_txt_height + 5
        if self.default_header_height[0] != "pixels":
            self.default_header_height = (
                self.default_header_height[0] if self.default_header_height[0] != "pixels" else "pixels",
                self.get_lines_cell_height(int(self.default_header_height[0]), font=self.header_font)
                if self.default_header_height[0] != "pixels"
                else self.default_header_height[1],
            )
        self.set_min_column_width()
        self.CH.set_height(self.default_header_height[1], set_TL=True)

    def set_index_font(self, newfont: tuple | None = None) -> tuple[str, int, str]:
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

    def purge_undo_and_redo_stack(self):
        self.undo_stack = deque(maxlen=self.max_undos)
        self.redo_stack = deque(maxlen=self.max_undos)

    def purge_redo_stack(self):
        self.redo_stack = deque(maxlen=self.max_undos)

    def data_reference(
        self,
        newdataref: list | tuple | None = None,
        reset_col_positions: bool = True,
        reset_row_positions: bool = True,
        redraw: bool = False,
        return_id: bool = True,
        keep_formatting: bool = True,
    ) -> object:
        if isinstance(newdataref, (list, tuple)):
            self.data = newdataref
            if keep_formatting:
                self.reapply_formatting()
            else:
                self.delete_all_formatting(clear_values=False)
            self.purge_undo_and_redo_stack()
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
            return w + self.table_txt_height, h
        return w, h

    def set_cell_size_to_text(self, r, c, only_set_if_too_small=False, redraw: bool = True, run_binding=False):
        min_column_width = int(self.min_column_width)
        min_rh = int(self.min_row_height)
        w = min_column_width
        h = min_rh
        datacn = self.datacn(c)
        datarn = self.datarn(r)
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
            if run_binding and self.CH.column_width_resize_func and old_width != new_width:
                self.CH.column_width_resize_func(
                    event_dict(
                        name="resize",
                        sheet=self.parentframe.name,
                        resized_columns={c: {"old_size": old_width, "new_size": new_width}},
                    )
                )
        if cell_needs_resize_h:
            old_height = self.row_positions[r + 1] - self.row_positions[r]
            new_row_pos = self.row_positions[r] + h
            increment = new_row_pos - self.row_positions[r + 1]
            self.row_positions[r + 2 :] = [
                e + increment for e in islice(self.row_positions, r + 2, len(self.row_positions))
            ]
            self.row_positions[r + 1] = new_row_pos
            new_height = self.row_positions[r + 1] - self.row_positions[r]
            if run_binding and self.RI.row_height_resize_func and old_height != new_height:
                self.RI.row_height_resize_func(
                    event_dict(
                        name="resize",
                        sheet=self.parentframe.name,
                        resized_rows={r: {"old_size": old_height, "new_size": new_height}},
                    )
                )
        if cell_needs_resize_w or cell_needs_resize_h:
            self.recreate_all_selection_boxes()
            if redraw:
                self.refresh()
                return True
            else:
                return False

    def set_all_cell_sizes_to_text(self, include_index: bool = False) -> tuple[list[float], list[float]]:
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
        numrows = self.total_data_rows()
        if self.all_columns_displayed:
            itercols = range(self.total_data_cols())
        else:
            itercols = self.displayed_columns
        if self.all_rows_displayed:
            iterrows = range(numrows)
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
                iterrows = range(numrows)
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
                    tw += self.table_txt_height
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
        self.set_row_positions(itr=(height for height in rhs.values()))
        self.set_col_positions(itr=(width for width in cws))
        self.recreate_all_selection_boxes()
        return self.row_positions, self.col_positions

    def set_col_positions(self, itr: Iterator[float]) -> None:
        self.col_positions = list(accumulate(chain([0], itr)))

    def reset_col_positions(self, ncols: int | None = None):
        colpos = int(self.default_column_width)
        if self.all_columns_displayed:
            self.set_col_positions(itr=(colpos for c in range(ncols if ncols is not None else self.total_data_cols())))
        else:
            self.set_col_positions(
                itr=(colpos for c in range(ncols if ncols is not None else len(self.displayed_columns)))
            )

    def set_row_positions(self, itr: Iterator[float]) -> None:
        self.row_positions = list(accumulate(chain([0], itr)))

    def reset_row_positions(self, nrows: int | None = None):
        rowpos = self.default_row_height[1]
        if self.all_rows_displayed:
            self.set_row_positions(itr=(rowpos for r in range(nrows if nrows is not None else self.total_data_rows())))
        else:
            self.set_row_positions(
                itr=(rowpos for r in range(nrows if nrows is not None else len(self.displayed_rows)))
            )

    def del_col_position(self, idx: int, deselect_all: bool = False):
        if deselect_all:
            self.deselect("all", redraw=False)
        if idx == "end" or len(self.col_positions) <= idx + 1:
            del self.col_positions[-1]
        else:
            w = self.col_positions[idx + 1] - self.col_positions[idx]
            idx += 1
            del self.col_positions[idx]
            self.col_positions[idx:] = [e - w for e in islice(self.col_positions, idx, len(self.col_positions))]

    def del_row_position(self, idx: int, deselect_all: bool = False):
        if deselect_all:
            self.deselect("all", redraw=False)
        if idx == "end" or len(self.row_positions) <= idx + 1:
            del self.row_positions[-1]
        else:
            w = self.row_positions[idx + 1] - self.row_positions[idx]
            idx += 1
            del self.row_positions[idx]
            self.row_positions[idx:] = [e - w for e in islice(self.row_positions, idx, len(self.row_positions))]

    def get_column_widths(self) -> list[int]:
        return diff_list(self.col_positions)

    def get_row_heights(self) -> list[int]:
        return diff_list(self.row_positions)

    def gen_column_widths(self) -> Generator[int]:
        return diff_gen(self.col_positions)

    def gen_row_heights(self) -> Generator[int]:
        return diff_gen(self.row_positions)

    def del_col_positions(self, idx: int, num: int = 1, deselect_all: bool = False):
        if deselect_all:
            self.deselect("all", redraw=False)
        if idx == "end" or len(self.col_positions) <= idx + 1:
            del self.col_positions[-1]
        else:
            cws = self.get_column_widths()
            cws[idx : idx + num] = []
            self.set_col_positions(itr=cws)

    def del_row_positions(self, idx: int, numrows: int = 1, deselect_all: bool = False):
        if deselect_all:
            self.deselect("all", redraw=False)
        if idx == "end" or len(self.row_positions) <= idx + 1:
            del self.row_positions[-1]
        else:
            rhs = self.get_row_heights()
            rhs[idx : idx + numrows] = []
            self.set_row_positions(itr=rhs)

    def insert_col_position(
        self,
        idx: Literal["end"] | int = "end",
        width: int | None = None,
        deselect_all: bool = False,
    ) -> None:
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

    def insert_row_position(
        self,
        idx: Literal["end"] | int = "end",
        height: int | None = None,
        deselect_all: bool = False,
    ) -> None:
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

    def insert_col_positions(
        self,
        idx: Literal["end"] | int = "end",
        widths: Sequence[float] | int | None = None,
        deselect_all: bool = False,
    ) -> None:
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

    def insert_row_positions(
        self,
        idx: Literal["end"] | int = "end",
        heights: Sequence[float] | int | None = None,
        deselect_all: bool = False,
    ) -> None:
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

    def named_span_coords(self, name: str | dict) -> tuple[int, int, int | None, int | None]:
        dct = self.named_spans[name] if isinstance(name, str) else name
        return (
            dct["from_r"],
            dct["from_c"],
            dct["upto_r"],
            dct["upto_c"],
        )

    def adjust_options_post_add_columns(
        self,
        cols: list | tuple,
        create_ops: bool = True,
    ) -> None:
        # self.cell_options = dict(to_add["cell_options"]) here
        self.cell_options = {
            (r, c if not (num := bisect_right(cols, c)) else c + num): v for (r, c), v in self.cell_options.items()
        }
        self.col_options = {
            c if not (num := bisect_right(cols, c)) else c + num: v for c, v in self.col_options.items()
        }
        self.CH.cell_options = {
            c if not (num := bisect_right(cols, c)) else c + num: v for c, v in self.CH.cell_options.items()
        }
        # if there are named spans where columns were added
        # add options to gap which was created by adding columns
        totalrows = None
        new_ops = self.parentframe.create_options_from_span
        qkspan = self.span()
        for name, span in self.named_spans.items():
            if isinstance(span["from_c"], int):
                for datacn in cols:
                    if span["from_c"] > datacn:
                        span["from_c"] += 1
                        if isinstance(span["upto_c"], int):
                            span["upto_c"] += 1
                    elif span["from_c"] <= datacn and (
                        (isinstance(span["upto_c"], int) and span["upto_c"] > datacn) or span["upto_c"] is None
                    ):
                        if isinstance(span["upto_c"], int):
                            span["upto_c"] += 1
                        # if to_add then it's an undo/redo and don't
                        # need to create fresh options
                        if create_ops:
                            # if rows are none it's a column options span
                            if span["from_r"] is None:
                                new_ops(
                                    mod_span(
                                        qkspan,
                                        span,
                                        from_c=datacn,
                                        upto_c=datacn + 1,
                                    )
                                )
                            # cells
                            else:
                                if totalrows is None:
                                    totalrows = self.total_data_rows()
                                rng_upto_r = totalrows if span["upto_r"] is None else span["upto_r"]
                                for rn in range(span["from_r"], rng_upto_r):
                                    new_ops(
                                        mod_span(
                                            qkspan,
                                            span,
                                            from_r=rn,
                                            from_c=datacn,
                                            upto_r=rn + 1,
                                            upto_c=datacn + 1,
                                        )
                                    )

    def adjust_options_post_add_rows(
        self,
        rows: list | tuple,
        create_ops: bool = True,
    ) -> None:
        self.cell_options = {
            (r if not (num := bisect_right(rows, r)) else r + num, c): v for (r, c), v in self.cell_options.items()
        }
        self.row_options = {
            r if not (num := bisect_right(rows, r)) else r + num: v for r, v in self.row_options.items()
        }
        self.RI.cell_options = {
            r if not (num := bisect_right(rows, r)) else r + num: v for r, v in self.RI.cell_options.items()
        }
        # if there are named spans where rows were added
        # add options to gap which was created by adding rows
        totalcols = None
        new_ops = self.parentframe.create_options_from_span
        qkspan = self.span()
        for name, span in self.named_spans.items():
            if isinstance(span["from_r"], int):
                for datarn in rows:
                    if span["from_r"] > datarn:
                        span["from_r"] += 1
                        if isinstance(span["upto_r"], int):
                            span["upto_r"] += 1
                    elif span["from_r"] <= datarn and (
                        (isinstance(span["upto_r"], int) and span["upto_r"] > datarn) or span["upto_r"] is None
                    ):
                        if isinstance(span["upto_r"], int):
                            span["upto_r"] += 1
                        # if to_add then it's an undo/redo and don't
                        # need to create fresh options
                        if create_ops:
                            # if rows are none it's a row options span
                            if span["from_c"] is None:
                                new_ops(
                                    mod_span(
                                        qkspan,
                                        span,
                                        from_r=datarn,
                                        upto_r=datarn + 1,
                                    )
                                )
                            # cells
                            else:
                                if totalcols is None:
                                    totalcols = self.total_data_cols()
                                rng_upto_c = totalcols if span["upto_c"] is None else span["upto_c"]
                                for cn in range(span["from_c"], rng_upto_c):
                                    new_ops(
                                        mod_span(
                                            qkspan,
                                            span,
                                            from_r=datarn,
                                            from_c=cn,
                                            upto_r=datarn + 1,
                                            upto_c=cn + 1,
                                        )
                                    )

    def adjust_options_post_delete_columns(
        self,
        to_del: None | set = None,
        to_bis: None | list = None,
        named_spans: None | set = None,
    ) -> list[int]:
        if to_del is None:
            to_del = set()
        if not to_bis:
            to_bis = sorted(to_del)
        self.cell_options = {
            (
                r,
                c if not (num := bisect_left(to_bis, c)) else c - num,
            ): v
            for (r, c), v in self.cell_options.items()
            if c not in to_del
        }
        self.col_options = {
            c if not (num := bisect_left(to_bis, c)) else c - num: v
            for c, v in self.col_options.items()
            if c not in to_del
        }
        self.CH.cell_options = {
            c if not (num := bisect_left(to_bis, c)) else c - num: v
            for c, v in self.CH.cell_options.items()
            if c not in to_del
        }
        self.del_columns_from_named_spans(
            to_del=to_del,
            to_bis=to_bis,
            named_spans=named_spans,
        )
        return to_bis

    def del_columns_from_named_spans(
        self,
        to_del: set,
        to_bis: list,
        named_spans: None | set = None,
    ) -> None:
        if named_spans is None:
            named_spans = self.get_spans_to_del_from_cols(cols=to_del)
        for name in named_spans:
            del self.named_spans[name]
        for name, span in self.named_spans.items():
            if isinstance(span["from_c"], int):
                for c in to_bis:
                    if span["from_c"] > c:
                        span["from_c"] -= 1
                    if isinstance(span["upto_c"], int) and span["upto_c"] > c:
                        span["upto_c"] -= 1

    def get_spans_to_del_from_cols(
        self,
        cols: set,
    ) -> set:
        total = self.total_data_cols()
        return {
            nm
            for nm, sp in self.named_spans.items()
            if isinstance(sp["from_c"], int)
            and all(c in cols for c in range(sp["from_c"], total if sp["upto_c"] is None else sp["upto_c"]))
        }

    def adjust_options_post_delete_rows(
        self,
        to_del: None | set = None,
        to_bis: None | list = None,
        named_spans: None | set = None,
    ) -> list[int]:
        if to_del is None:
            to_del = set()
        if not to_bis:
            to_bis = sorted(to_del)
        self.cell_options = {
            (
                r if not (num := bisect_left(to_bis, r)) else r - num,
                c,
            ): v
            for (r, c), v in self.cell_options.items()
            if r not in to_del
        }
        self.row_options = {
            r if not (num := bisect_left(to_bis, r)) else r - num: v
            for r, v in self.row_options.items()
            if r not in to_del
        }
        self.RI.cell_options = {
            r if not (num := bisect_left(to_bis, r)) else r - num: v
            for r, v in self.RI.cell_options.items()
            if r not in to_del
        }
        self.del_rows_from_named_spans(
            to_del=to_del,
            to_bis=to_bis,
            named_spans=named_spans,
        )
        return to_bis

    def del_rows_from_named_spans(
        self,
        to_del: set,
        to_bis: list,
        named_spans: None | set = None,
    ) -> None:
        if named_spans is None:
            named_spans = self.get_spans_to_del_from_rows(rows=to_del)
        for name in named_spans:
            del self.named_spans[name]
        for name, span in self.named_spans.items():
            if isinstance(span["from_r"], int):
                for r in to_bis:
                    if span["from_r"] > r:
                        span["from_r"] -= 1
                    if isinstance(span["upto_r"], int) and span["upto_r"] > r:
                        span["upto_r"] -= 1

    def get_spans_to_del_from_rows(
        self,
        rows: set,
    ) -> set:
        total = self.total_data_rows()
        return {
            nm
            for nm, sp in self.named_spans.items()
            if isinstance(sp["from_r"], int)
            and all(r in rows for r in range(sp["from_r"], total if sp["upto_r"] is None else sp["upto_r"]))
        }

    def add_columns(
        self,
        columns: dict,
        header: dict,
        column_widths: dict,
        event_data: dict,
        displayed_columns: None | list[int] = None,
        create_ops: bool = True,
        create_selections: bool = True,
        add_row_positions: bool = True,
    ) -> EventDataDict:
        saved_displayed_columns = list(self.displayed_columns)
        if isinstance(displayed_columns, list):
            self.displayed_columns = displayed_columns
        elif not self.all_columns_displayed:
            # push displayed indexes by one for every inserted column
            self.displayed_columns.sort()
            # highest index is first in columns
            up_to = len(self.displayed_columns)
            for cn in columns:
                self.displayed_columns.insert((last_ins := bisect_left(self.displayed_columns, cn)), cn)
                self.displayed_columns[last_ins + 1 : up_to] = [
                    i + 1 for i in islice(self.displayed_columns, last_ins + 1, up_to)
                ]
                up_to = last_ins
        cws = self.get_column_widths()
        if column_widths and next(reversed(column_widths)) > len(cws):
            for i in reversed(range(len(cws), len(cws) + next(reversed(column_widths)) - len(cws))):
                column_widths[i] = self.default_column_width
        self.set_col_positions(
            itr=insert_items(
                cws,
                column_widths,
            )
        )
        # we're inserting so we can use indexes == len for
        # fix functions, the values will go on the end
        maxrn = 0
        for cn, rowdict in reversed(columns.items()):
            for rn, v in rowdict.items():
                if rn > len(self.data):
                    self.fix_data_len(rn - 1, cn - 1)
                if rn > maxrn:
                    maxrn = rn
                self.data[rn].insert(cn, v)
        # if not hiding rows then we can extend row positions if necessary
        if add_row_positions and self.all_rows_displayed and maxrn + 1 > len(self.row_positions) - 1:
            self.set_row_positions(
                itr=chain(
                    self.gen_row_heights(),
                    (self.default_row_height[1] for i in range(len(self.row_positions) - 1, maxrn + 1)),
                )
            )
        if isinstance(self._headers, list):
            self._headers = insert_items(self._headers, header, self.CH.fix_header)
        self.adjust_options_post_add_columns(
            cols=tuple(reversed(columns)),
            create_ops=create_ops,
        )
        if create_selections:
            self.deselect("all")
            for chunk in consecutive_chunks(tuple(reversed(column_widths))):
                self.create_selection_box(
                    0,
                    chunk[0],
                    len(self.row_positions) - 1,
                    chunk[-1] + 1,
                    "columns",
                    run_binding=True,
                )
        event_data["added"]["columns"] = {
            "table": columns,
            "header": header,
            "column_widths": column_widths,
            "displayed_columns": saved_displayed_columns,
        }
        return event_data

    def rc_add_columns(self, event: object = None):
        rowlen = self.equalize_data_row_lengths(include_header=False)
        selcols = sorted(self.get_selected_cols())
        if (
            selcols
            and isinstance(self.CH.popup_menu_loc, int)
            and selcols[0] <= self.CH.popup_menu_loc
            and selcols[-1] >= self.CH.popup_menu_loc
        ):
            selcols = get_seq_without_gaps_at_index(selcols, self.CH.popup_menu_loc)
            numcols = len(selcols)
            displayed_ins_col = selcols[0] if event == "left" else selcols[-1] + 1
            if self.all_columns_displayed:
                data_ins_col = int(displayed_ins_col)
            else:
                if displayed_ins_col == len(self.col_positions) - 1:
                    data_ins_col = rowlen
                else:
                    if len(self.displayed_columns) > displayed_ins_col:
                        data_ins_col = int(self.displayed_columns[displayed_ins_col])
                    else:
                        data_ins_col = int(self.displayed_columns[-1])
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
        event_data = event_dict(
            name="add_columns",
            sheet=self.parentframe.name,
            boxes=self.get_boxes(),
            selected=self.currently_selected(),
        )
        if not try_binding(self.extra_begin_insert_cols_rc_func, event_data, "begin_add_columns"):
            return
        event_data = self.add_columns(
            *self.get_args_for_add_columns(data_ins_col, displayed_ins_col, numcols),
            event_data=event_data,
        )
        if self.undo_enabled:
            self.undo_stack.append(ev_stack_dict(event_data))
        self.refresh()
        try_binding(self.extra_end_insert_cols_rc_func, event_data, "end_add_columns")
        self.sheet_modified(event_data)

    def add_rows(
        self,
        rows: dict,
        index: dict,
        row_heights: dict,
        event_data: dict,
        displayed_rows: None | list[int] = None,
        create_ops: bool = True,
        create_selections: bool = True,
        add_col_positions: bool = True,
    ) -> EventDataDict:
        saved_displayed_rows = list(self.displayed_rows)
        if isinstance(displayed_rows, list):
            self.displayed_rows = displayed_rows
        elif not self.all_rows_displayed:
            # push displayed indexes by one for every inserted column
            self.displayed_rows.sort()
            # highest index is first in rows
            up_to = len(self.displayed_rows)
            for rn in rows:
                self.displayed_rows.insert((last_ins := bisect_left(self.displayed_rows, rn)), rn)
                self.displayed_rows[last_ins + 1 : up_to] = [
                    i + 1 for i in islice(self.displayed_rows, last_ins + 1, up_to)
                ]
                up_to = last_ins
        rhs = self.get_row_heights()
        if row_heights and next(reversed(row_heights)) > len(rhs):
            for i in reversed(range(len(rhs), len(rhs) + next(reversed(row_heights)) - len(rhs))):
                row_heights[i] = self.default_row_height[1]
        self.set_row_positions(
            itr=insert_items(
                rhs,
                row_heights,
            )
        )
        maxcn = 0
        for rn, row in reversed(rows.items()):
            cn = len(row) - 1
            if rn > len(self.data):
                self.fix_data_len(rn - 1, cn)
            self.data.insert(rn, row)
            if cn > maxcn:
                maxcn = cn
        if isinstance(self.row_index, list):
            self._row_index = insert_items(self._row_index, index, self.RI.fix_index)
        # if not hiding columns then we can extend col positions if necessary
        if add_col_positions and self.all_columns_displayed and maxcn + 1 > len(self.col_positions) - 1:
            self.set_col_positions(
                itr=chain(
                    self.gen_column_widths(),
                    (self.default_column_width for i in range(len(self.col_positions) - 1, maxcn + 1)),
                )
            )
        self.adjust_options_post_add_rows(
            rows=tuple(reversed(rows)),
            create_ops=create_ops,
        )
        if create_selections:
            self.deselect("all")
            for chunk in consecutive_chunks(tuple(reversed(row_heights))):
                self.create_selection_box(
                    chunk[0],
                    0,
                    chunk[-1] + 1,
                    len(self.col_positions) - 1,
                    "rows",
                    run_binding=True,
                )
        event_data["added"]["rows"] = {
            "table": rows,
            "index": index,
            "row_heights": row_heights,
            "displayed_rows": saved_displayed_rows,
        }
        return event_data

    def rc_add_rows(self, event: object = None):
        total_data_rows = self.total_data_rows()
        selrows = sorted(self.get_selected_rows())
        if (
            selrows
            and isinstance(self.RI.popup_menu_loc, int)
            and selrows[0] <= self.RI.popup_menu_loc
            and selrows[-1] >= self.RI.popup_menu_loc
        ):
            selrows = get_seq_without_gaps_at_index(selrows, self.RI.popup_menu_loc)
            numrows = len(selrows)
            displayed_ins_row = selrows[0] if event == "above" else selrows[-1] + 1
            if self.all_rows_displayed:
                data_ins_row = int(displayed_ins_row)
            else:
                if displayed_ins_row == len(self.row_positions) - 1:
                    data_ins_row = total_data_rows
                else:
                    if len(self.displayed_rows) > displayed_ins_row:
                        data_ins_row = int(self.displayed_rows[displayed_ins_row])
                    else:
                        data_ins_row = int(self.displayed_rows[-1])
        else:
            numrows = 1
            displayed_ins_row = len(self.row_positions) - 1
            data_ins_row = int(displayed_ins_row)
        if isinstance(self.paste_insert_row_limit, int) and self.paste_insert_row_limit < displayed_ins_row + numrows:
            numrows = self.paste_insert_row_limit - len(self.row_positions) - 1
            if numrows < 1:
                return
        event_data = event_dict(
            name="add_rows",
            sheet=self.parentframe.name,
            boxes=self.get_boxes(),
            selected=self.currently_selected(),
        )
        if not try_binding(self.extra_begin_insert_rows_rc_func, event_data, "begin_add_rows"):
            return
        event_data = self.add_rows(
            *self.get_args_for_add_rows(data_ins_row, displayed_ins_row, numrows),
            event_data=event_data,
        )
        if self.undo_enabled:
            self.undo_stack.append(ev_stack_dict(event_data))
        self.refresh()
        try_binding(self.extra_end_insert_rows_rc_func, event_data, "end_add_rows")
        self.sheet_modified(event_data)

    def get_args_for_add_columns(
        self,
        data_ins_col: int,
        displayed_ins_col: int,
        numcols: int,
        columns: list | None = None,
        widths: list | None = None,
        headers: bool = False,
    ) -> tuple[dict, dict, dict]:
        if isinstance(self._headers, list):
            if headers and columns:
                header_data = {
                    datacn: column[0]
                    for datacn, column in zip(reversed(range(data_ins_col, data_ins_col + numcols)), reversed(columns))
                }
            else:
                header_data = {
                    datacn: self.CH.get_value_for_empty_cell(datacn, c_ops=False)
                    for datacn in reversed(range(data_ins_col, data_ins_col + numcols))
                }
        else:
            header_data = {}
        if columns is None:
            columns = {
                datacn: {
                    datarn: self.get_value_for_empty_cell(datarn, datacn, c_ops=False)
                    for datarn in range(len(self.data))
                }
                for datacn in reversed(range(data_ins_col, data_ins_col + numcols))
            }
        else:
            if headers:
                start = 1
            else:
                start = 0
            columns = {
                datacn: {datarn: v for datarn, v in enumerate(islice(column, start, None))}
                for datacn, column in zip(reversed(range(data_ins_col, data_ins_col + numcols)), reversed(columns))
            }
        if widths is None:
            widths = {
                c: self.default_column_width for c in reversed(range(displayed_ins_col, displayed_ins_col + numcols))
            }
        else:
            widths = {
                c: width
                for c, width in zip(reversed(range(displayed_ins_col, displayed_ins_col + numcols)), reversed(widths))
            }
        return columns, header_data, widths

    def get_args_for_add_rows(
        self,
        data_ins_row: int,
        displayed_ins_row: int,
        numrows: int,
        rows: list | None = None,
        heights: list | None = None,
        row_index: bool = False,
        total_data_cols=None,
    ) -> tuple:
        if isinstance(self._row_index, list):
            if row_index and rows:
                index_data = {
                    datarn: v[0]
                    for datarn, v in zip(reversed(range(data_ins_row, data_ins_row + numrows)), reversed(rows))
                }
            else:
                index_data = {
                    datarn: self.RI.get_value_for_empty_cell(datarn, r_ops=False)
                    for datarn in reversed(range(data_ins_row, data_ins_row + numrows))
                }
        else:
            index_data = {}
        if rows is None:
            if total_data_cols is None:
                total_data_cols = self.total_data_cols()
            rows = {
                datarn: [self.get_value_for_empty_cell(datarn, c, c_ops=False) for c in range(total_data_cols)]
                for datarn in reversed(range(data_ins_row, data_ins_row + numrows))
            }
        else:
            if row_index:
                start = 1
            else:
                start = 0
            rows = {
                datarn: v[start:] if start and v else v
                for datarn, v in zip(reversed(range(data_ins_row, data_ins_row + numrows)), reversed(rows))
            }
        if heights is None:
            heights = {
                r: self.default_row_height[1] for r in reversed(range(displayed_ins_row, displayed_ins_row + numrows))
            }
        else:
            heights = {
                r: height
                for r, height in zip(reversed(range(displayed_ins_row, displayed_ins_row + numrows)), reversed(heights))
            }
        return rows, index_data, heights

    def pickle_options(self) -> bytes:
        return pickle_obj(
            {
                "cell_options": self.cell_options,
                "column_options": self.col_options,
                "row_options": self.row_options,
                "CH_cell_options": self.CH.cell_options,
                "RI_cell_options": self.RI.cell_options,
            }
        )

    def delete_columns_data(self, cols: list, event_data: dict) -> EventDataDict:
        self.mouseclick_outside_editor_or_dropdown_all_canvases()
        event_data["deleted"]["displayed_columns"] = (
            list(self.displayed_columns) if not isinstance(self.displayed_columns, int) else int(self.displayed_columns)
        )
        event_data["options"] = self.pickle_options()
        event_data["named_spans"] = {k: span.pickle_self() for k, span in self.named_spans.items()}
        for datacn in reversed(cols):
            for rn in range(len(self.data)):
                if datacn not in event_data["deleted"]["columns"]:
                    event_data["deleted"]["columns"][datacn] = {}
                try:
                    event_data["deleted"]["columns"][datacn][rn] = self.data[rn].pop(datacn)
                except Exception:
                    continue
            try:
                event_data["deleted"]["header"][datacn] = self._headers.pop(datacn)
            except Exception:
                continue
        cols_set = set(cols)
        self.adjust_options_post_delete_columns(
            to_del=cols_set,
            to_bis=cols,
            named_spans=self.get_spans_to_del_from_cols(cols=cols_set),
        )
        if not self.all_columns_displayed:
            self.displayed_columns = [c for c in self.displayed_columns if c not in cols_set]
            for c in cols:
                self.displayed_columns = [dc if c > dc else dc - 1 for dc in self.displayed_columns]
        return event_data

    def delete_columns_displayed(self, cols: list, event_data: dict) -> EventDataDict:
        cols_set = set(cols)
        for c in reversed(cols):
            event_data["deleted"]["column_widths"][c] = self.col_positions[c + 1] - self.col_positions[c]
        self.set_col_positions(itr=(width for c, width in enumerate(self.gen_column_widths()) if c not in cols_set))
        return event_data

    def rc_delete_columns(self, event: object = None):
        selected = sorted(self.get_selected_cols())
        curr = self.currently_selected()
        if not selected or not curr:
            return
        event_data = event_dict(
            name="delete_columns",
            sheet=self.parentframe.name,
            boxes=self.get_boxes(),
            selected=self.currently_selected(),
        )
        if not try_binding(self.extra_begin_del_cols_rc_func, event_data, "begin_delete_columns"):
            return
        event_data = self.delete_columns_displayed(selected, event_data)
        data_cols = selected if self.all_columns_displayed else [self.displayed_columns[c] for c in selected]
        event_data = self.delete_columns_data(data_cols, event_data)
        if self.undo_enabled:
            self.undo_stack.append(ev_stack_dict(event_data))
        self.deselect("all")
        try_binding(self.extra_end_del_cols_rc_func, event_data, "end_delete_columns")
        self.sheet_modified(event_data)

    def delete_rows_data(self, rows: list, event_data: dict) -> EventDataDict:
        self.mouseclick_outside_editor_or_dropdown_all_canvases()
        event_data["deleted"]["displayed_rows"] = (
            list(self.displayed_rows) if not isinstance(self.displayed_rows, int) else int(self.displayed_rows)
        )
        event_data["options"] = self.pickle_options()
        event_data["named_spans"] = {k: span.pickle_self() for k, span in self.named_spans.items()}
        for datarn in reversed(rows):
            event_data["deleted"]["rows"][datarn] = self.data.pop(datarn)
            try:
                event_data["deleted"]["index"][datarn] = self._row_index.pop(datarn)
            except Exception:
                continue
        rows_set = set(rows)
        self.adjust_options_post_delete_rows(
            to_del=rows_set,
            to_bis=rows,
            named_spans=self.get_spans_to_del_from_rows(rows=rows_set),
        )
        if not self.all_rows_displayed:
            self.displayed_rows = [r for r in self.displayed_rows if r not in rows_set]
            for r in rows:
                self.displayed_rows = [dr if r > dr else dr - 1 for dr in self.displayed_rows]
        return event_data

    def delete_rows_displayed(self, rows: list, event_data: dict) -> EventDataDict:
        rows_set = set(rows)
        for r in reversed(rows):
            event_data["deleted"]["row_heights"][r] = self.row_positions[r + 1] - self.row_positions[r]
        self.set_row_positions(itr=(height for r, height in enumerate(self.gen_row_heights()) if r not in rows_set))
        return event_data

    def rc_delete_rows(self, event: object = None):
        selected = sorted(self.get_selected_rows())
        curr = self.currently_selected()
        if not selected or not curr:
            return
        event_data = event_dict(
            name="delete_rows",
            sheet=self.parentframe.name,
            boxes=self.get_boxes(),
            selected=self.currently_selected(),
        )
        if not try_binding(self.extra_begin_del_rows_rc_func, event_data, "begin_delete_rows"):
            return
        event_data = self.delete_rows_displayed(selected, event_data)
        data_rows = selected if self.all_rows_displayed else [self.displayed_rows[r] for r in selected]
        event_data = self.delete_rows_data(data_rows, event_data)
        if self.undo_enabled:
            self.undo_stack.append(ev_stack_dict(event_data))
        self.deselect("all")
        try_binding(self.extra_end_del_rows_rc_func, event_data, "end_delete_rows")
        self.sheet_modified(event_data)

    def move_row_position(self, idx1: int, idx2: int):
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

    def move_col_position(self, idx1: int, idx2: int):
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
        rows: int | Iterator | None = None,
        all_rows_displayed: bool | None = None,
        reset_row_positions: bool = True,
        deselect_all: bool = True,
    ) -> list[int] | None:
        if rows is None and all_rows_displayed is None:
            return list(range(self.total_data_rows())) if self.all_rows_displayed else self.displayed_rows
        total_data_rows = None
        if (rows is not None and rows != self.displayed_rows) or (all_rows_displayed and not self.all_rows_displayed):
            self.purge_undo_and_redo_stack()
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
        columns: int | Iterator | None = None,
        all_columns_displayed: bool | None = None,
        reset_col_positions: bool = True,
        deselect_all: bool = True,
    ) -> list[int] | None:
        if columns is None and all_columns_displayed is None:
            return list(range(self.total_data_cols())) if self.all_columns_displayed else self.displayed_columns
        total_data_cols = None
        if (columns is not None and columns != self.displayed_columns) or (
            all_columns_displayed and not self.all_columns_displayed
        ):
            self.purge_undo_and_redo_stack()
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
        newheaders: object = None,
        index: int | None = None,
        reset_col_positions: bool = False,
        show_headers_if_not_sheet: bool = True,
        redraw: bool = False,
    ) -> object:
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
                    raise ValueError(
                        """
                        New header must be iterable or int \
                        (use int to use a row as the header
                        """
                    )
            if reset_col_positions:
                self.reset_col_positions()
            elif (
                show_headers_if_not_sheet
                and isinstance(self._headers, list)
                and (self.col_positions == [0] or not self.col_positions)
            ):
                colpos = int(self.default_column_width)
                if self.all_columns_displayed:
                    self.set_col_positions(itr=repeat(colpos, len(self._headers)))
                else:
                    self.set_col_positions(itr=repeat(colpos, len(self.displayed_columns)))
            if redraw:
                self.refresh()
        else:
            if not isinstance(self._headers, int) and index is not None and isinstance(index, int):
                return self._headers[index]
            else:
                return self._headers

    def row_index(
        self,
        newindex: object = None,
        index: int | None = None,
        reset_row_positions: bool = False,
        show_index_if_not_sheet: bool = True,
        redraw: bool = False,
    ) -> object:
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
                    raise ValueError(
                        """
                        New index must be iterable or int \
                        (use int to use a column as the index
                        """
                    )
            if reset_row_positions:
                self.reset_row_positions()
            elif (
                show_index_if_not_sheet
                and isinstance(self._row_index, list)
                and (self.row_positions == [0] or not self.row_positions)
            ):
                rowpos = self.default_row_height[1]
                if self.all_rows_displayed:
                    self.set_row_positions(itr=repeat(rowpos, len(self._row_index)))
                else:
                    self.set_row_positions(itr=repeat(rowpos, len(self.displayed_rows)))
            if redraw:
                self.refresh()
        else:
            if not isinstance(self._row_index, int) and index is not None and isinstance(index, int):
                return self._row_index[index]
            else:
                return self._row_index

    def total_data_cols(self, include_header: bool = True) -> int:
        h_total = 0
        d_total = 0
        if include_header:
            if isinstance(self._headers, (list, tuple)):
                h_total = len(self._headers)
        # map() for some reason is 15% faster than max(key=len)
        # using python 3.11 windows 11
        d_total = max(map(len, self.data), default=0)
        return h_total if h_total > d_total else d_total

    def total_data_rows(self, include_index: bool = True) -> int:
        i_total = 0
        d_total = 0
        if include_index:
            if isinstance(self._row_index, (list, tuple)):
                i_total = len(self._row_index)
        d_total = len(self.data)
        return i_total if i_total > d_total else d_total

    def data_dimensions(self, total_rows: int | None = None, total_columns: int | None = None):
        if total_rows is None and total_columns is None:
            return self.total_data_rows(), self.total_data_cols()
        if total_rows is not None:
            if len(self.data) < total_rows:
                ncols = self.total_data_cols() if total_columns is None else total_columns
                self.data.extend([self.get_empty_row_seq(r, ncols) for r in range(total_rows - len(self.data))])
            else:
                self.data[total_rows:] = []
        if total_columns is not None:
            for rn, r in enumerate(self.data):
                if (lnr := len(r)) > total_columns:
                    r = r[:total_columns]
                elif lnr < total_columns:
                    r += self.get_empty_row_seq(rn, end=total_columns, start=lnr)

    def equalize_data_row_lengths(
        self,
        include_header: bool = False,
        total_data_cols: int | None = None,
        at_least_cols: int | None = None,
    ) -> int:
        if not isinstance(total_data_cols, int):
            total_data_cols = self.total_data_cols(include_header=include_header)
        if isinstance(at_least_cols, int) and at_least_cols > total_data_cols:
            total_data_cols = at_least_cols
        if include_header and total_data_cols > len(self._headers):
            self.CH.fix_header(total_data_cols)
        for rn, r in enumerate(self.data):
            if total_data_cols > (lnr := len(r)):
                r += self.get_empty_row_seq(rn, end=total_data_cols, start=lnr)
        return total_data_cols

    def get_canvas_visible_area(self) -> tuple[float, float, float, float]:
        return (
            self.canvasx(0),
            self.canvasy(0),
            self.canvasx(self.winfo_width()),
            self.canvasy(self.winfo_height()),
        )

    def get_visible_rows(self, y1: int | float, y2: int | float) -> tuple[int, int]:
        start_row = bisect_left(self.row_positions, y1)
        end_row = bisect_right(self.row_positions, y2)
        if not y2 >= self.row_positions[-1]:
            end_row += 1
        return start_row, end_row

    def get_visible_columns(self, x1: int | float, x2: int | float) -> tuple[int, int]:
        start_col = bisect_left(self.col_positions, x1)
        end_col = bisect_right(self.col_positions, x2)
        if not x2 >= self.col_positions[-1]:
            end_col += 1
        return start_col, end_col

    def redraw_highlight_get_text_fg(
        self,
        r: int,
        c: int,
        fc: int | float,
        fr: int | float,
        sc: int | float,
        sr: int | float,
        c_2_: tuple[int, int, int],
        c_3_: tuple[int, int, int],
        c_4_: tuple[int, int, int],
        selections: dict,
        datarn: int,
        datacn: int,
        can_width: int | None,
    ) -> str:
        redrawn = False
        kwargs = self.get_cell_kwargs(datarn, datacn, key="highlight")
        if kwargs:
            if kwargs[0] is not None:
                c_1 = kwargs[0] if kwargs[0].startswith("#") else color_map[kwargs[0]]
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
        if not is_type_int(pc) or pc >= 100 or pc <= 0:
            coords = (
                x1 - 1 if outline else x1,
                y1 - 1 if outline else y1,
                x2 if can_width is None else x2 + can_width,
                y2,
            )
        else:
            coords = (x1, y1, (x2 - x1) * (pc / 100), y2)
        if self.hidd_high:
            iid, showing = self.hidd_high.popitem()
            self.coords(iid, coords)
            if showing:
                self.itemconfig(iid, fill=fill, outline=outline)
            else:
                self.itemconfig(iid, fill=fill, outline=outline, tag=tag, state="normal")
        else:
            iid = self.create_rectangle(coords, fill=fill, outline=outline, tag=tag)
        self.disp_high[iid] = True
        return True

    def redraw_gridline(
        self,
        points,
        fill,
        width,
        tag,
    ):
        if self.hidd_grid:
            iid, sh = self.hidd_grid.popitem()
            self.coords(iid, points)
            if sh:
                self.itemconfig(
                    iid,
                    fill=fill,
                    width=width,
                    capstyle=tk.BUTT,
                    joinstyle=tk.ROUND,
                )
            else:
                self.itemconfig(
                    iid,
                    fill=fill,
                    width=width,
                    capstyle=tk.BUTT,
                    joinstyle=tk.ROUND,
                    state="normal",
                )
        else:
            iid = self.create_line(
                points,
                fill=fill,
                width=width,
                capstyle=tk.BUTT,
                joinstyle=tk.ROUND,
                tag=tag,
            )
        self.disp_grid[iid] = True
        return iid

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
            topysub = floor(self.table_half_txt_height / 2)
            mid_y = y1 + floor(self.min_row_height / 2)
            if mid_y + topysub + 1 >= y1 + self.table_txt_height - 1:
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
            tx1 = x2 - self.table_txt_height + 1
            tx2 = x2 - self.table_half_txt_height - 1
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

    def redraw_checkbox(
        self,
        x1: int | float,
        y1: int | float,
        x2: int | float,
        y2: int | float,
        fill: str,
        outline: str,
        tag: str | tuple,
        draw_check: bool = False,
    ) -> None:
        points = get_checkbox_points(x1, y1, x2, y2)
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
            points = get_checkbox_points(x1, y1, x2, y2, radius=4)
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
        redraw_header: bool = False,
        redraw_row_index: bool = False,
        redraw_table: bool = True,
    ) -> bool:
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
            if self.auto_resize_columns < self.min_column_width:
                min_column_width = self.column_width
            else:
                min_column_width = self.auto_resize_columns
            if (len(self.col_positions) - 1) * min_column_width < max_w:
                resized_cols = True
                change = int((max_w - self.col_positions[-1]) / (len(self.col_positions) - 1))
                widths = [
                    int(b - a) + change - 1
                    for a, b in zip(self.col_positions, islice(self.col_positions, 1, len(self.col_positions)))
                ]
                diffs = {}
                for i, w in enumerate(widths):
                    if w < min_column_width:
                        diffs[i] = min_column_width - w
                        widths[i] = min_column_width
                if diffs and len(diffs) < len(widths):
                    change = sum(diffs.values()) / (len(widths) - len(diffs))
                    for i, w in enumerate(widths):
                        if i not in diffs:
                            widths[i] -= change
                self.set_col_positions(itr=widths)
        if self.auto_resize_rows and self.allow_auto_resize_rows and row_pos_exists:
            max_h = int(can_height)
            max_h -= self.empty_vertical
            if self.auto_resize_rows < self.min_row_height:
                min_row_height = self.min_row_height
            else:
                min_row_height = self.auto_resize_rows
            if (len(self.row_positions) - 1) * min_row_height < max_h:
                resized_rows = True
                change = int((max_h - self.row_positions[-1]) / (len(self.row_positions) - 1))
                heights = [
                    int(b - a) + change - 1
                    for a, b in zip(self.row_positions, islice(self.row_positions, 1, len(self.row_positions)))
                ]
                diffs = {}
                for i, h in enumerate(heights):
                    if h < min_row_height:
                        diffs[i] = min_row_height - h
                        heights[i] = min_row_height
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
        end_row = bisect_right(self.row_positions, scrollpos_bot)
        if not scrollpos_bot >= self.row_positions[-1]:
            end_row += 1
        if redraw_row_index and self.show_index:
            self.RI.auto_set_index_width(end_row - 1)
            # return
        scrollpos_left = self.canvasx(0)
        scrollpos_top = self.canvasy(0)
        scrollpos_right = self.canvasx(can_width)
        start_row = bisect_left(self.row_positions, scrollpos_top)
        start_col = bisect_left(self.col_positions, scrollpos_left)
        end_col = bisect_right(self.col_positions, scrollpos_right)
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
        if redraw_table and self.show_horizontal_grid and row_pos_exists:
            if self.horizontal_grid_to_end_of_window:
                x_grid_stop = scrollpos_right + can_width
            else:
                if last_col_line_pos > scrollpos_right:
                    x_grid_stop = x_stop + 1
                else:
                    x_grid_stop = x_stop - 1
            points = list(
                chain.from_iterable(
                    [
                        (
                            self.canvasx(0) - 1,
                            self.row_positions[r],
                            x_grid_stop,
                            self.row_positions[r],
                            self.canvasx(0) - 1,
                            self.row_positions[r],
                            self.canvasx(0) - 1,
                            self.row_positions[r + 1] if len(self.row_positions) - 1 > r else self.row_positions[r],
                        )
                        for r in range(start_row - 1, end_row)
                    ]
                )
            )
            if points:
                self.redraw_gridline(
                    points=points,
                    fill=self.table_grid_fg,
                    width=1,
                    tag="g",
                )
        if redraw_table and self.show_vertical_grid and col_pos_exists:
            if self.vertical_grid_to_end_of_window:
                y_grid_stop = scrollpos_bot + can_height
            else:
                if last_row_line_pos > scrollpos_bot:
                    y_grid_stop = y_stop + 1
                else:
                    y_grid_stop = y_stop - 1
            points = list(
                chain.from_iterable(
                    [
                        (
                            self.col_positions[c],
                            scrollpos_top - 1,
                            self.col_positions[c],
                            y_grid_stop,
                            self.col_positions[c],
                            scrollpos_top - 1,
                            self.col_positions[c + 1] if len(self.col_positions) - 1 > c else self.col_positions[c],
                            scrollpos_top - 1,
                        )
                        for c in range(start_col - 1, end_col)
                    ]
                )
            )
            if points:
                self.redraw_gridline(
                    points=points,
                    fill=self.table_grid_fg,
                    width=1,
                    tag="g",
                )
        if start_row > 0:
            start_row -= 1
        if start_col > 0:
            start_col -= 1
        end_row -= 1
        if redraw_table:
            selections = self.get_redraw_selections(start_row, end_row, start_col, end_col)
            c_2 = (
                self.table_selected_cells_bg
                if self.table_selected_cells_bg.startswith("#")
                else color_map[self.table_selected_cells_bg]
            )
            c_2_ = (int(c_2[1:3], 16), int(c_2[3:5], 16), int(c_2[5:], 16))
            c_3 = (
                self.table_selected_columns_bg
                if self.table_selected_columns_bg.startswith("#")
                else color_map[self.table_selected_columns_bg]
            )
            c_3_ = (int(c_3[1:3], 16), int(c_3[3:5], 16), int(c_3[5:], 16))
            c_4 = (
                self.table_selected_rows_bg
                if self.table_selected_rows_bg.startswith("#")
                else color_map[self.table_selected_rows_bg]
            )
            c_4_ = (int(c_4[1:3], 16), int(c_4[3:5], 16), int(c_4[5:], 16))
            rows_ = tuple(range(start_row, end_row))
            font = self.table_font
            dd_coords = self.get_existing_dropdown_coords()
            for c in range(start_col, end_col - 1):
                for r in rows_:
                    rtopgridln = self.row_positions[r]
                    rbotgridln = self.row_positions[r + 1]
                    if rbotgridln - rtopgridln < self.table_txt_height:
                        continue
                    cleftgridln = self.col_positions[c]
                    crightgridln = self.col_positions[c + 1]

                    datarn = self.datarn(r)
                    datacn = self.datacn(c)

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
                            mw = crightgridln - cleftgridln - self.table_txt_height - 2
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
                                dd_is_open=dd_coords == (r, c),
                            )
                        else:
                            mw = crightgridln - cleftgridln - 1
                    elif align == "e":
                        if kwargs:
                            mw = crightgridln - cleftgridln - self.table_txt_height - 2
                            draw_x = crightgridln - 5 - self.table_txt_height
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
                                dd_is_open=dd_coords == (r, c),
                            )
                        else:
                            mw = crightgridln - cleftgridln - 1
                            draw_x = crightgridln - 3
                    elif align == "center":
                        stop = cleftgridln + 5
                        if kwargs:
                            mw = crightgridln - cleftgridln - self.table_txt_height - 2
                            draw_x = cleftgridln + ceil((crightgridln - cleftgridln - self.table_txt_height) / 2)
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
                                dd_is_open=dd_coords == (r, c),
                            )
                        else:
                            mw = crightgridln - cleftgridln - 1
                            draw_x = cleftgridln + floor((crightgridln - cleftgridln) / 2)
                    if not kwargs:
                        kwargs = self.get_cell_kwargs(datarn, datacn, key="checkbox")
                        if kwargs and mw > self.table_txt_height + 1:
                            box_w = self.table_txt_height + 1
                            if align == "w":
                                draw_x += box_w + 3
                                mw -= box_w + 3
                            elif align == "center":
                                draw_x += ceil(box_w / 2) + 1
                                mw -= box_w + 2
                            else:
                                mw -= box_w + 1
                            try:
                                draw_check = self.data[datarn][datacn]
                            except Exception:
                                draw_check = False
                            self.redraw_checkbox(
                                cleftgridln + 2,
                                rtopgridln + 2,
                                cleftgridln + self.table_txt_height + 3,
                                rtopgridln + self.table_txt_height + 3,
                                fill=fill if kwargs["state"] == "normal" else self.table_grid_fg,
                                outline="",
                                tag="cb",
                                draw_check=draw_check,
                            )
                    lns = self.get_valid_cell_data_as_str(datarn, datacn, get_displayed=True).split("\n")
                    if (
                        lns != [""]
                        and mw > self.table_txt_width
                        and not (
                            (align == "w" and draw_x > scrollpos_right)
                            or (align == "e" and cleftgridln + 5 > scrollpos_right)
                            or (align == "center" and stop > scrollpos_right)
                        )
                    ):
                        draw_y = rtopgridln + self.table_first_ln_ins
                        start_ln = int((scrollpos_top - rtopgridln) / self.table_xtra_lines_increment)
                        if start_ln < 0:
                            start_ln = 0
                        draw_y += start_ln * self.table_xtra_lines_increment
                        if draw_y + self.table_half_txt_height - 1 <= rbotgridln and len(lns) > start_ln:
                            for txt in islice(lns, start_ln, None):
                                if self.hidd_text:
                                    iid, showing = self.hidd_text.popitem()
                                    self.coords(iid, draw_x, draw_y)
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
                                    self.tag_raise(iid)
                                else:
                                    iid = self.create_text(
                                        draw_x,
                                        draw_y,
                                        text=txt,
                                        fill=fill,
                                        font=font,
                                        anchor=align,
                                        tag="t",
                                    )
                                self.disp_text[iid] = True
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
                                draw_y += self.table_xtra_lines_increment
                                if draw_y + self.table_half_txt_height - 1 > rbotgridln:
                                    break
        if redraw_table:
            for dct in (self.hidd_text, self.hidd_high, self.hidd_grid, self.hidd_dropdown, self.hidd_checkbox):
                for iid, showing in dct.items():
                    if showing:
                        self.itemconfig(iid, state="hidden")
                        dct[iid] = False
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
        event_data = {"sheetname": "", "header": redraw_header, "row_index": redraw_row_index, "table": redraw_table}
        self.parentframe.emit_event("<<SheetRedrawn>>", data=event_data)
        return True

    def get_selection_items(
        self,
        cells: bool = True,
        rows: bool = True,
        columns: bool = True,
        current: bool = True,
        reverse: bool = False,
    ) -> list:
        return sorted(
            (self.find_withtag("cells") if cells else tuple())
            + (self.find_withtag("rows") if rows else tuple())
            + (self.find_withtag("columns") if columns else tuple())
            + (self.find_withtag("selected") if current else tuple()),
            reverse=reverse,
        )

    def get_boxes(self) -> dict:
        boxes = {}
        for item in self.get_selection_items(current=False):
            tags = self.gettags(item)
            boxes[coords_tag_to_box_nt(tags[1])] = tags[0]
        return boxes

    def reselect_from_get_boxes(
        self,
        boxes: dict,
        curr: tuple = tuple(),
    ) -> None:
        for (r1, c1, r2, c2), v in boxes.items():
            if r2 < len(self.row_positions) and c2 < len(self.col_positions):
                self.create_selection_box(r1, c1, r2, c2, v, run_binding=True)
        if curr:
            self.set_currently_selected(tags=curr.tags)

    def currently_selected(self, get_item: bool = False) -> int | tuple | CurrentlySelectedClass:
        items = self.get_selection_items(cells=False, rows=False, columns=False)
        if not items:
            return tuple()
        # more than one currently selected box shouldn't be allowed
        # but just to make sure we get the most recent one anyway
        if get_item:
            return items[-1]
        tags = self.gettags(items[-1])
        r, c = coords_tag_to_int_tuple(tags[3])
        # remove "s" from end
        type_ = tags[4].split("_")[1][:-1]
        return CurrentlySelectedClass(r, c, type_, tags)

    def move_currently_selected_within_box(self, r: int, c: int) -> None:
        curr = self.currently_selected()
        if curr:
            self.set_currently_selected(r=r, c=c, item=self.to_item_int(curr.tags), tags=curr.tags)

    def set_currently_selected(self, r: int | None = None, c: int | None = None, **kwargs) -> None:
        # if r and c are provided it will attempt to
        # put currently box at that coordinate
        # _________
        # "selected" tags have the most information about the box
        if "tags" in kwargs and kwargs["tags"]:
            r_, c_ = coords_tag_to_int_tuple(kwargs["tags"][3])
            if r is None:
                r = r_
            if c is None:
                c = c_
            if "item" in kwargs:
                self.set_currently_selected(
                    r=r, c=c, item=kwargs["item"], box=coords_tag_to_int_tuple(kwargs["tags"][1])
                )
            else:
                self.set_currently_selected(r=r, c=c, box=coords_tag_to_int_tuple(kwargs["tags"][1]))
            return
        # place at item if r and c are in bounds
        if "item" in kwargs:
            tags = self.gettags(self.to_item_int(kwargs["item"]))
            if tags:
                r1, c1, r2, c2 = coords_tag_to_int_tuple(tags[1])
                if r is None:
                    r = r1
                if c is None:
                    c = c1
                if r1 <= r and c1 <= c and r2 > r and c2 > c:
                    self.create_currently_selected_box(
                        r,
                        c,
                        ("selected", tags[1], tags[2], f"{r}_{c}", f"type_{tags[0]}"),
                    )
                    return
        # currently selected is pointed at any selection box with "box" coordinates
        if "box" in kwargs:
            if r is None:
                r = kwargs["box"][0]
            if c is None:
                c = kwargs["box"][1]
            for item in self.get_selection_items(current=False):
                tags = self.gettags(item)
                r1, c1, r2, c2 = coords_tag_to_int_tuple(tags[1])
                if kwargs["box"] == (r1, c1, r2, c2) and r1 <= r and c1 <= c and r2 > r and c2 > c:
                    self.create_currently_selected_box(
                        r,
                        c,
                        ("selected", tags[1], tags[2], f"{r}_{c}", f"type_{tags[0]}"),
                    )
                    return
        # currently selected is just pointed at a coordinate
        # find the top most box there, requires r and c
        if r is not None and c is not None:
            for item in self.get_selection_items(current=False):
                tags = self.gettags(item)
                r1, c1, r2, c2 = coords_tag_to_int_tuple(tags[1])
                if r1 <= r and c1 <= c and r2 > r and c2 > c:
                    self.create_currently_selected_box(
                        r,
                        c,
                        ("selected", tags[1], tags[2], f"{r}_{c}", f"type_{tags[0]}"),
                    )
                    return
            # wasn't provided an item and couldn't find a box at coords so select cell
            self.select_cell(r, c, redraw=True)

    def set_current_to_last(self) -> None:
        items = self.get_selection_items(current=False)
        if items:
            item = items[-1]
            r1, c1, r2, c2, type_ = self.get_box_from_item(item)
            if r2 - r1 == 1 and c2 - c1 == 1:
                self.itemconfig(item, state="hidden")
            self.set_currently_selected(item=item)

    def delete_item(self, item: int, set_current: bool = True) -> None:
        if item is None:
            return
        item = self.to_item_int(item)
        self.delete(f"iid_{item}")
        self.RI.delete(f"iid_{item}")
        self.CH.delete(f"iid_{item}")

    def get_box_from_item(self, item: int, get_dict: bool = False) -> dict | tuple:
        if item is None or item is True:
            return tuple()
        # it's in the form of an item tag f"iid_{item}"
        item = self.to_item_int(item)
        tags = self.gettags(item)
        if tags:
            if get_dict:
                if tags[0] == "selected":
                    return {coords_tag_to_box_nt(tags[1]): tags[4].split("_")[1]}
                else:
                    return {coords_tag_to_box_nt(tags[1]): tags[0]}
            else:
                if tags[0] == "selected":
                    return Box_t(*(coords_tag_to_int_tuple(tags[1]) + (tags[4].split("_")[1],)))
                else:
                    return Box_t(*(coords_tag_to_int_tuple(tags[1]) + (tags[0],)))
        return tuple()

    def get_selected_box_bg_fg(self, type_: str) -> tuple:
        if type_ == "cells":
            return self.table_selected_cells_bg, self.table_selected_box_cells_fg
        elif type_ == "rows":
            return self.table_selected_rows_bg, self.table_selected_box_rows_fg
        elif type_ == "columns":
            return self.table_selected_columns_bg, self.table_selected_box_columns_fg

    def create_currently_selected_box(self, r: int, c: int, tags: tuple) -> int:
        type_ = tags[4].split("_")[1]
        fill, outline = self.get_selected_box_bg_fg(type_=type_)
        iid = self.currently_selected(get_item=True)
        x1 = self.col_positions[c] + 1
        y1 = self.row_positions[r] + 1
        x2 = self.col_positions[c + 1] if index_exists(self.col_positions, c + 1) else self.col_positions[c]
        y2 = self.row_positions[r + 1] if index_exists(self.row_positions, r + 1) else self.row_positions[r]
        if isinstance(iid, int):
            self.coords(
                iid,
                x1,
                y1,
                x2,
                y2,
            )
            if self.show_selected_cells_border:
                self.itemconfig(
                    iid,
                    fill="",
                    outline=outline,
                    width=2,
                    tags=tags,
                )
                self.tag_raise(iid)
            else:
                self.itemconfig(
                    iid,
                    fill=fill,
                    outline="",
                    tags=tags,
                )
        else:
            if self.show_selected_cells_border:
                iid = self.create_rectangle(
                    x1,
                    y1,
                    x2,
                    y2,
                    fill="",
                    outline=outline,
                    width=2,
                    tags=tags,
                )
            else:
                iid = self.create_rectangle(
                    x1,
                    y1,
                    x2,
                    y2,
                    fill=fill,
                    outline="",
                    tags=tags,
                )
        if not self.show_selected_cells_border:
            self.tag_lower(iid)
            self.lower_selection_boxes()
        return iid

    def create_selection_box(
        self,
        r1: int,
        c1: int,
        r2: int,
        c2: int,
        type_: str = "cells",
        state: str = "normal",
        set_current: bool | tuple[int, int] = True,
        run_binding: bool = False,
    ) -> int:
        if self.col_positions == [0]:
            c1 = 0
            c2 = 0
        if self.row_positions == [0]:
            r1 = 0
            r2 = 0
        self.itemconfig("cells", state="normal")
        if type_ == "cells":
            fill_tags = ("cells", f"{r1}_{c1}_{r2}_{c2}")
            border_tags = ("cellsbd", f"{r1}_{c1}_{r2}_{c2}")
            mt_bg = self.table_selected_cells_bg
            mt_border_col = self.table_selected_cells_border_fg
        elif type_ == "rows":
            fill_tags = ("rows", f"{r1}_{c1}_{r2}_{c2}")
            border_tags = ("rowsbd", f"{r1}_{c1}_{r2}_{c2}")
            mt_bg = self.table_selected_rows_bg
            mt_border_col = self.table_selected_rows_border_fg
        elif type_ == "columns":
            fill_tags = ("columns", f"{r1}_{c1}_{r2}_{c2}")
            border_tags = ("columnsbd", f"{r1}_{c1}_{r2}_{c2}")
            mt_bg = self.table_selected_columns_bg
            mt_border_col = self.table_selected_columns_border_fg
        fill_iid = self.create_rectangle(
            self.col_positions[c1],
            self.row_positions[r1],
            self.canvasx(self.winfo_width()) if self.selected_rows_to_end_of_window else self.col_positions[c2],
            self.row_positions[r2],
            fill=mt_bg,
            outline="",
            state=state,
            tags=fill_tags,
        )
        tag_addon = f"iid_{fill_iid}"
        self.addtag_withtag(tag_addon, fill_iid)
        self.last_selected = fill_iid
        tag_index_header_type = fill_tags + (tag_addon,)
        tag_index_header = ("cells", f"{r1}_{c1}_{r2}_{c2}", tag_addon)
        ch_tags = tag_index_header if type_ == "rows" else tag_index_header_type
        ri_tags = tag_index_header if type_ == "columns" else tag_index_header_type
        border_tags = border_tags + (tag_addon,)
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
        if set_current:
            if set_current is True:
                curr_r = r1
                curr_c = c1
            elif isinstance(set_current, tuple):
                curr_r = set_current[0]
                curr_c = set_current[1]
            currently_selected_tags = (
                "selected",  # tags[0] name
                f"{r1}_{c1}_{r2}_{c2}",  # tags[1] dimensions of box it's attached to
                tag_addon,  # tags[2] iid of box it's attached to
                f"{curr_r}_{curr_c}",  # tags[3] position of currently selected box
                f"type_{type_}",  # tags[4] type of box it's attached to
            )
            self.create_currently_selected_box(curr_r, curr_c, tags=currently_selected_tags)
        if self.show_selected_cells_border and (
            (self.being_drawn_item is None and self.RI.being_drawn_item is None and self.CH.being_drawn_item is None)
            or len(self.anything_selected()) > 1
        ):
            self.create_rectangle(
                self.col_positions[c1],
                self.row_positions[r1],
                self.col_positions[c2],
                self.row_positions[r2],
                fill="",
                outline=mt_border_col,
                tags=border_tags,
            )
        self.lower_selection_boxes()
        if run_binding:
            self.run_selection_binding(type_)
        return fill_iid

    def lower_selection_boxes(self) -> None:
        self.tag_lower("rows")
        self.RI.tag_lower("rows")
        self.tag_lower("columns")
        self.CH.tag_lower("columns")
        self.tag_lower("cells")
        self.RI.tag_lower("cells")
        self.CH.tag_lower("cells")
        if self.show_selected_cells_border:
            self.tag_raise("selected")

    def recreate_selection_box(
        self,
        r1: int,
        c1: int,
        r2: int,
        c2: int,
        fill_iid: int,
        run_binding: bool = False,
    ) -> None:
        alltags = self.gettags(fill_iid)
        type_ = alltags[0]
        tag_addon = f"iid_{fill_iid}"
        if type_ == "cells":
            fill_tags = ("cells", f"{r1}_{c1}_{r2}_{c2}", tag_addon)
            border_tags = ("cellsbd", f"{r1}_{c1}_{r2}_{c2}", tag_addon)
            mt_bg = self.table_selected_cells_bg
            mt_border_col = self.table_selected_cells_border_fg
        elif type_ == "rows":
            fill_tags = ("rows", f"{r1}_{c1}_{r2}_{c2}", tag_addon)
            border_tags = ("rowsbd", f"{r1}_{c1}_{r2}_{c2}", tag_addon)
            mt_bg = self.table_selected_rows_bg
            mt_border_col = self.table_selected_rows_border_fg
        elif type_ == "columns":
            fill_tags = ("columns", f"{r1}_{c1}_{r2}_{c2}", tag_addon)
            border_tags = ("columnsbd", f"{r1}_{c1}_{r2}_{c2}", tag_addon)
            mt_bg = self.table_selected_columns_bg
            mt_border_col = self.table_selected_columns_border_fg
        self.coords(
            fill_iid,
            self.col_positions[c1],
            self.row_positions[r1],
            self.canvasx(self.winfo_width()) if self.selected_rows_to_end_of_window else self.col_positions[c2],
            self.row_positions[r2],
        )
        self.itemconfig(
            fill_iid,
            fill=mt_bg,
            outline="",
            tags=fill_tags,
        )
        tag_index_header_type = fill_tags
        tag_index_header = ("cells", f"{r1}_{c1}_{r2}_{c2}", tag_addon)
        ch_tags = tag_index_header if type_ == "rows" else tag_index_header_type
        ri_tags = tag_index_header if type_ == "columns" else tag_index_header_type
        self.RI.coords(
            tag_addon,
            0,
            self.row_positions[r1],
            self.RI.current_width - 1,
            self.row_positions[r2],
        )
        self.RI.itemconfig(
            tag_addon,
            fill=self.RI.index_selected_rows_bg if type_ == "rows" else self.RI.index_selected_cells_bg,
            outline="",
            tags=ri_tags,
        )
        self.CH.coords(
            tag_addon,
            self.col_positions[c1],
            0,
            self.col_positions[c2],
            self.CH.current_height - 1,
        )
        self.CH.itemconfig(
            tag_addon,
            fill=self.CH.header_selected_columns_bg if type_ == "columns" else self.CH.header_selected_cells_bg,
            outline="",
            tags=ch_tags,
        )
        # check for border of selection box which is a separate item
        border_item = [item for item in self.find_withtag(tag_addon) if self.gettags(item)[0].endswith("bd")]
        if border_item:
            border_item = border_item[0]
            if self.show_selected_cells_border:
                self.coords(
                    border_item,
                    self.col_positions[c1],
                    self.row_positions[r1],
                    self.col_positions[c2],
                    self.row_positions[r2],
                )
                self.itemconfig(
                    border_item,
                    fill="",
                    outline=mt_border_col,
                    tags=border_tags,
                )
            else:
                self.delete(border_item)
        if run_binding:
            self.run_selection_binding(type_)

    def run_selection_binding(self, type_: str) -> None:
        if type_ == "cells":
            if self.selection_binding_func:
                self.selection_binding_func(
                    self.get_select_event(being_drawn_item=self.being_drawn_item),
                )
        elif type_ == "rows":
            if self.RI.selection_binding_func:
                self.RI.selection_binding_func(
                    self.get_select_event(being_drawn_item=self.RI.being_drawn_item),
                )
        elif type_ == "columns":
            if self.CH.selection_binding_func:
                self.CH.selection_binding_func(
                    self.get_select_event(being_drawn_item=self.CH.being_drawn_item),
                )

    def recreate_all_selection_boxes(self) -> None:
        curr = self.currently_selected()
        if not curr:
            return
        for item in self.get_selection_items(current=False):
            tags = self.gettags(item)
            r1, c1, r2, c2 = coords_tag_to_int_tuple(tags[1])
            if r1 >= len(self.row_positions) - 1 or c1 >= len(self.col_positions) - 1:
                self.delete_item(item)
                continue
            if r2 > len(self.row_positions) - 1:
                r2 = len(self.row_positions) - 1
            if c2 > len(self.col_positions) - 1:
                c2 = len(self.col_positions) - 1
            self.recreate_selection_box(r1, c1, r2, c2, item)
        self.set_currently_selected(tags=curr.tags)

    def to_item_int(self, item: int | str | tuple) -> int:
        # item arg is a tuple of tags
        if isinstance(item, tuple):
            if isinstance(item, CurrentlySelectedClass):
                return int(item[3][2].split("_")[1])
            return int(item[2].split("_")[1])
        # item arg is one of those tags
        if isinstance(item, str):
            return int(item.split("_")[1])
        # item is probably an int
        return item

    def get_redraw_selections(self, startr: int, endr: int, startc: int, endc: int) -> dict:
        d = defaultdict(list)
        for item in self.get_selection_items(current=False):
            tags = self.gettags(item)
            d[tags[0]].append(coords_tag_to_int_tuple(tags[1]))
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

    def get_selected_min_max(self) -> tuple[int, int, int, int] | tuple[None, None, None, None]:
        min_x = float("inf")
        min_y = float("inf")
        max_x = 0
        max_y = 0
        for item in self.get_selection_items():
            r1, c1, r2, c2 = coords_tag_to_int_tuple(self.gettags(item)[1])
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
        return None, None, None, None

    def get_selected_rows(
        self,
        get_cells: bool = False,
        within_range: tuple | None = None,
        get_cells_as_rows: bool = False,
    ) -> set[int] | set[tuple[int, int]]:
        s = set()
        if within_range is not None:
            within_r1 = within_range[0]
            within_r2 = within_range[1]
        if get_cells:
            if within_range is None:
                for item in self.get_selection_items(cells=False, columns=False, current=False):
                    r1, c1, r2, c2 = coords_tag_to_int_tuple(self.gettags(item)[1])
                    s.update(set(product(range(r1, r2), range(0, len(self.col_positions) - 1))))
                if get_cells_as_rows:
                    s.update(self.get_selected_cells())
            else:
                for item in self.get_selection_items(cells=False, columns=False, current=False):
                    r1, c1, r2, c2 = coords_tag_to_int_tuple(self.gettags(item)[1])
                    if r1 >= within_r1 or r2 <= within_r2:
                        s.update(
                            set(
                                product(
                                    range(r1 if r1 > within_r1 else within_r1, r2 if r2 < within_r2 else within_r2),
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
                for item in self.get_selection_items(cells=False, columns=False, current=False):
                    r1, c1, r2, c2 = coords_tag_to_int_tuple(self.gettags(item)[1])
                    s.update(set(range(r1, r2)))
                if get_cells_as_rows:
                    s.update(set(tup[0] for tup in self.get_selected_cells()))
            else:
                for item in self.get_selection_items(cells=False, columns=False, current=False):
                    r1, c1, r2, c2 = coords_tag_to_int_tuple(self.gettags(item)[1])
                    if r1 >= within_r1 or r2 <= within_r2:
                        s.update(set(range(r1 if r1 > within_r1 else within_r1, r2 if r2 < within_r2 else within_r2)))
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

    def get_selected_cols(
        self,
        get_cells: bool = False,
        within_range: tuple | None = None,
        get_cells_as_cols: bool = False,
    ) -> set[int] | set[tuple[int, int]]:
        s = set()
        if within_range is not None:
            within_c1 = within_range[0]
            within_c2 = within_range[1]
        if get_cells:
            if within_range is None:
                for item in self.get_selection_items(cells=False, rows=False, current=False):
                    r1, c1, r2, c2 = coords_tag_to_int_tuple(self.gettags(item)[1])
                    s.update(set(product(range(c1, c2), range(0, len(self.row_positions) - 1))))
                if get_cells_as_cols:
                    s.update(self.get_selected_cells())
            else:
                for item in self.get_selection_items(cells=False, rows=False, current=False):
                    r1, c1, r2, c2 = coords_tag_to_int_tuple(self.gettags(item)[1])
                    if c1 >= within_c1 or c2 <= within_c2:
                        s.update(
                            set(
                                product(
                                    range(c1 if c1 > within_c1 else within_c1, c2 if c2 < within_c2 else within_c2),
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
                for item in self.get_selection_items(cells=False, rows=False, current=False):
                    r1, c1, r2, c2 = coords_tag_to_int_tuple(self.gettags(item)[1])
                    s.update(set(range(c1, c2)))
                if get_cells_as_cols:
                    s.update(set(tup[1] for tup in self.get_selected_cells()))
            else:
                for item in self.get_selection_items(cells=False, rows=False, current=False):
                    r1, c1, r2, c2 = coords_tag_to_int_tuple(self.gettags(item)[1])
                    if c1 >= within_c1 or c2 <= within_c2:
                        s.update(set(range(c1 if c1 > within_c1 else within_c1, c2 if c2 < within_c2 else within_c2)))
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

    def get_selected_cells(
        self,
        get_rows: bool = False,
        get_cols: bool = False,
        within_range: bool = None,
    ) -> set[tuple[int, int]]:
        s = set()
        if within_range is not None:
            within_r1 = within_range[0]
            within_c1 = within_range[1]
            within_r2 = within_range[2]
            within_c2 = within_range[3]
        if within_range is None:
            for item in self.get_selection_items(rows=get_rows, columns=get_cols, current=False):
                r1, c1, r2, c2 = coords_tag_to_int_tuple(self.gettags(item)[1])
                s.update(set(product(range(r1, r2), range(c1, c2))))
        else:
            for item in self.get_selection_items(rows=get_rows, columns=get_cols, current=False):
                r1, c1, r2, c2 = coords_tag_to_int_tuple(self.gettags(item)[1])
                if r1 >= within_r1 or c1 >= within_c1 or r2 <= within_r2 or c2 <= within_c2:
                    s.update(
                        set(
                            product(
                                range(r1 if r1 > within_r1 else within_r1, r2 if r2 < within_r2 else within_r2),
                                range(c1 if c1 > within_c1 else within_c1, c2 if c2 < within_c2 else within_c2),
                            )
                        )
                    )
        return s

    def get_all_selection_boxes(self) -> tuple[tuple[int, int, int, int]]:
        return tuple(coords_tag_to_int_tuple(self.gettags(item)[1]) for item in self.get_selection_items(current=False))

    def get_all_selection_boxes_with_types(self) -> list[tuple[tuple[int, int, int, int], str]]:
        boxes = []
        for item in self.get_selection_items(current=False):
            tags = self.gettags(item)
            boxes.append((coords_tag_to_int_tuple(tags[1]), tags[0]))
        return boxes

    def all_selected(self) -> bool:
        for r1, c1, r2, c2 in self.get_all_selection_boxes():
            if not r1 and not c1 and r2 == len(self.row_positions) - 1 and c2 == len(self.col_positions) - 1:
                return True
        return False

    # don't have to use "selected" because you can't
    # have a current without a selection box
    def cell_selected(
        self,
        r: int,
        c: int,
        inc_cols: bool = False,
        inc_rows: bool = False,
    ) -> bool:
        if not isinstance(r, int) or not isinstance(c, int):
            return False
        for item in self.get_selection_items(rows=inc_rows, columns=inc_cols, current=False):
            r1, c1, r2, c2 = coords_tag_to_int_tuple(self.gettags(item)[1])
            if r1 <= r and c1 <= c and r2 > r and c2 > c:
                return True
        return False

    def col_selected(self, c: int) -> bool:
        if not isinstance(c, int):
            return False
        for item in self.get_selection_items(cells=False, rows=False, current=False):
            r1, c1, r2, c2 = coords_tag_to_int_tuple(self.gettags(item)[1])
            if c1 <= c and c2 > c:
                return True
        return False

    def row_selected(self, r: int) -> bool:
        if not isinstance(r, int):
            return False
        for item in self.get_selection_items(cells=False, columns=False, current=False):
            r1, c1, r2, c2 = coords_tag_to_int_tuple(self.gettags(item)[1])
            if r1 <= r and r2 > r:
                return True
        return False

    def anything_selected(
        self,
        exclude_columns: bool = False,
        exclude_rows: bool = False,
        exclude_cells: bool = False,
    ) -> list[int]:
        return self.get_selection_items(
            columns=not exclude_columns,
            rows=not exclude_rows,
            cells=not exclude_cells,
            current=False,
        )

    def open_cell(
        self,
        event: object = None,
        ignore_existing_editor: bool = False,
    ) -> None:
        if not self.anything_selected() or (not ignore_existing_editor and self.text_editor_id is not None):
            return
        currently_selected = self.currently_selected()
        if not currently_selected:
            return
        r, c = int(currently_selected[0]), int(currently_selected[1])
        datacn = self.datacn(c)
        datarn = self.datarn(r)
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

    def event_opens_dropdown_or_checkbox(self, event=None) -> bool:
        if event is None:
            return False
        elif event == "rc":
            return True
        elif (
            (hasattr(event, "keysym") and event.keysym == "Return")
            or (hasattr(event, "keysym") and event.keysym == "F2")
            or (
                event is not None
                and hasattr(event, "keycode")
                and event.keycode == "??"
                and hasattr(event, "num")
                and event.num == 1
            )  # mouseclick
            or (hasattr(event, "keysym") and event.keysym == "BackSpace")
        ):
            return True
        else:
            return False

    # displayed indexes
    def get_cell_align(self, r: int, c: int) -> str:
        datarn = self.datarn(r)
        datacn = self.datacn(c)
        cell_alignment = self.get_cell_kwargs(datarn, datacn, key="align")
        if cell_alignment:
            return cell_alignment
        return self.align

    # displayed indexes
    def open_text_editor(
        self,
        event: object = None,
        r: int = 0,
        c: int = 0,
        text: str | None = None,
        state: str = "normal",
        dropdown: bool = False,
    ) -> bool:
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
        if self.extra_begin_edit_cell_func:
            try:
                text = self.extra_begin_edit_cell_func(
                    event_dict(
                        name="begin_edit_table",
                        sheet=self.parentframe.name,
                        key=extra_func_key,
                        value=text,
                        loc=tuple(self.text_editor_loc),
                        boxes=self.get_boxes(),
                        selected=self.currently_selected(),
                    )
                )
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
        if not self.see(r=r, c=c, check_cell_visibility=True):
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
        self.text_editor.textedit.bind("<Alt-Return>", lambda _x: self.text_editor_newline_binding(r, c))
        if USER_OS == "darwin":
            self.text_editor.textedit.bind("<Option-Return>", lambda _x: self.text_editor_newline_binding(r, c))
        for key, func in self.text_editor_user_bound_keys.items():
            self.text_editor.textedit.bind(key, func)
        self.text_editor.textedit.bind("<Tab>", lambda _x: self.close_text_editor((r, c, "Tab")))
        self.text_editor.textedit.bind("<Return>", lambda _x: self.close_text_editor((r, c, "Return")))
        if not dropdown:
            self.text_editor.textedit.bind("<FocusOut>", lambda _x: self.close_text_editor((r, c, "FocusOut")))
        self.text_editor.textedit.bind("<Escape>", lambda _x: self.close_text_editor((r, c, "Escape")))
        return True

    # displayed indexes
    def text_editor_newline_binding(
        self,
        r: int = 0,
        c: int = 0,
        event: object = None,
        check_lines: bool = True,
    ) -> None:
        datarn = self.datarn(r)
        datacn = self.datacn(c)
        curr_height = self.text_editor.winfo_height()
        if not check_lines or self.get_lines_cell_height(self.text_editor.get_num_lines() + 1) > curr_height:
            new_height = curr_height + self.table_xtra_lines_increment
            space_bot = self.get_space_bot(r)
            if new_height > space_bot:
                new_height = space_bot
            if new_height != curr_height:
                self.text_editor.config(height=new_height)
                kwargs = self.get_cell_kwargs(datarn, datacn, key="dropdown")
                if kwargs:
                    text_editor_h = self.text_editor.winfo_height()
                    win_h, anchor = self.get_dropdown_height_anchor(r, c, text_editor_h)
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

    def refresh_open_window_positions(self):
        if self.text_editor is not None:
            r, c = self.text_editor_loc
            self.text_editor.config(height=self.row_positions[r + 1] - self.row_positions[r])
            self.coords(
                self.text_editor_id,
                self.col_positions[c],
                self.row_positions[r],
            )
        if self.existing_dropdown_window is not None:
            r, c = self.get_existing_dropdown_coords()
            if self.text_editor is None:
                text_editor_h = self.row_positions[r + 1] - self.row_positions[r]
                anchor = self.itemcget(self.existing_dropdown_canvas_id, "anchor")
                win_h = 0
            else:
                text_editor_h = self.text_editor.winfo_height()
                win_h, anchor = self.get_dropdown_height_anchor(r, c, text_editor_h)
            if anchor == "nw":
                self.coords(
                    self.existing_dropdown_canvas_id,
                    self.col_positions[c],
                    self.row_positions[r] + text_editor_h - 1,
                )
                # self.itemconfig(self.existing_dropdown_canvas_id, anchor=anchor, height=win_h)
            elif anchor == "sw":
                self.coords(
                    self.existing_dropdown_canvas_id,
                    self.col_positions[c],
                    self.row_positions[r],
                )
                # self.itemconfig(self.existing_dropdown_canvas_id, anchor=anchor, height=win_h)

    def destroy_text_editor(self, reason: None | str = None) -> None:
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
        if reason == "Escape":
            self.focus_set()

    # c is displayed col
    def close_text_editor(
        self,
        editor_info: tuple | None = None,
    ) -> str | None:
        # checking if text editor should be closed or not
        focused = self.focus_get()
        try:
            if focused == self.text_editor.textedit.rc_popup_menu:
                return "break"
        except Exception:
            pass
        if focused is None and editor_info:
            return "break"
        if editor_info[2] == "Escape":
            self.destroy_text_editor("Escape")
            self.close_dropdown_window()
            return
        # setting cell data with text editor value
        self.text_editor_value = self.text_editor.get()
        self.destroy_text_editor()
        currently_selected = self.currently_selected()
        r, c = editor_info[0], editor_info[1]
        datarn, datacn = self.datarn(r), self.datacn(c)
        event_data = event_dict(
            name="end_edit_table",
            sheet=self.parentframe.name,
            cells_table={(datarn, datacn): self.text_editor_value},
            key=editor_info[2],
            value=self.text_editor_value,
            loc=(r, c),
            boxes=self.get_boxes(),
            selected=currently_selected,
        )
        move_down = False
        set_data = partial(
            self.set_cell_data_undo,
            r=r,
            c=c,
            datarn=datarn,
            datacn=datacn,
            redraw=False,
            check_input_valid=False,
        )
        if self.edit_validation_func:
            self.text_editor_value = self.edit_validation_func(event_data)
            if self.text_editor_value is not None and self.input_valid_for_cell(datarn, datacn, self.text_editor_value):
                move_down = set_data(value=self.text_editor_value)
        elif self.input_valid_for_cell(datarn, datacn, self.text_editor_value):
            move_down = set_data(value=self.text_editor_value)
        if move_down:
            try_binding(self.extra_end_edit_cell_func, event_data)
        if (
            move_down
            and r is not None
            and c is not None
            and currently_selected
            and r == currently_selected[0]
            and c == currently_selected[1]
            and (self.single_selection_enabled or self.toggle_selection_enabled)
        ):
            r1, c1, r2, c2 = self.get_box_containing_current()
            numcols = c2 - c1
            numrows = r2 - r1
            if numcols == 1 and numrows == 1:
                if editor_info[2] == "Return":
                    self.select_cell(r + 1 if r < len(self.row_positions) - 2 else r, c)
                    self.see(
                        r + 1 if r < len(self.row_positions) - 2 else r,
                        c,
                        keep_xscroll=True,
                        bottom_right_corner=True,
                        check_cell_visibility=True,
                    )
                elif editor_info[2] == "Tab":
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
                if editor_info[2] == "Return":
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
                elif editor_info[2] == "Tab":
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
                self.move_currently_selected_within_box(new_r, new_c)
                self.see(
                    new_r,
                    new_c,
                    keep_xscroll=False,
                    bottom_right_corner=True,
                    check_cell_visibility=True,
                )
        self.close_dropdown_window(r, c)
        self.recreate_all_selection_boxes()
        self.refresh()
        if editor_info[2] != "FocusOut":
            self.focus_set()
        return "break"

    def tab_key(self, event: object = None) -> str:
        currently_selected = self.currently_selected()
        if not currently_selected:
            return
        r = currently_selected.row
        c = currently_selected.column
        r1, c1, r2, c2 = self.get_box_containing_current()
        numcols = c2 - c1
        numrows = r2 - r1
        if numcols == 1 and numrows == 1:
            new_r = r
            new_c = c + 1 if c < len(self.col_positions) - 2 else c
            self.select_cell(new_r, new_c)
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
            self.move_currently_selected_within_box(new_r, new_c)
        self.see(
            new_r,
            new_c,
            keep_xscroll=False,
            bottom_right_corner=True,
            check_cell_visibility=True,
        )
        return "break"

    def get_space_bot(self, r: int, text_editor_h: int | None = None) -> int:
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

    def get_dropdown_height_anchor(self, r: int, c: int, text_editor_h: int | None = None) -> tuple:
        win_h = 5
        datarn, datacn = self.datarn(r), self.datacn(c)
        for i, v in enumerate(self.get_cell_kwargs(datarn, datacn, key="dropdown")["values"]):
            v_numlines = len(v.split("\n") if isinstance(v, str) else f"{v}".split("\n"))
            if v_numlines > 1:
                win_h += self.table_first_ln_ins + (v_numlines * self.table_xtra_lines_increment) + 5  # end of cell
            else:
                win_h += self.min_row_height
            if i == 5:
                break
        if win_h > 500:
            win_h = 500
        space_bot = self.get_space_bot(r, text_editor_h)
        space_top = int(self.row_positions[r])
        anchor = "nw"
        win_h2 = int(win_h)
        if win_h > space_bot:
            if space_bot >= space_top:
                anchor = "nw"
                win_h = space_bot - 1
            elif space_top > space_bot:
                anchor = "sw"
                win_h = space_top - 1
        if win_h < self.table_txt_height + 5:
            win_h = self.table_txt_height + 5
        elif win_h > win_h2:
            win_h = win_h2
        return win_h, anchor

    def dropdown_text_editor_modified(
        self,
        dd_window: object,
        event: dict,
        modified_func: Callable | None,
    ) -> None:
        if modified_func:
            modified_func(event)
        dd_window.search_and_see(event)

    # c is displayed col
    def open_dropdown_window(
        self,
        r: int,
        c: int,
        event: object = None,
    ) -> None:
        self.destroy_text_editor("Escape")
        self.destroy_opened_dropdown_window()
        datarn = self.datarn(r)
        datacn = self.datacn(c)
        kwargs = self.get_cell_kwargs(datarn, datacn, key="dropdown")
        if kwargs["state"] == "normal":
            if not self.open_text_editor(event=event, r=r, c=c, dropdown=True):
                return
        win_h, anchor = self.get_dropdown_height_anchor(r, c)
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
            outline_color=self.get_selected_box_bg_fg(type_="cells")[1],
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
                    event_dict(
                        name="table_dropdown_modified",
                        sheet=self.parentframe.name,
                        value=self.text_editor.get(),
                        loc=(r, c),
                        boxes=self.get_boxes(),
                        selected=self.currently_selected(),
                    )
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
    def close_dropdown_window(
        self,
        r: int | None = None,
        c: int | None = None,
        selection: object = None,
        redraw: bool = True,
    ) -> None:
        if r is not None and c is not None and selection is not None:
            datacn = self.datacn(c)
            datarn = self.datarn(r)
            kwargs = self.get_cell_kwargs(datarn, datacn, key="dropdown")
            pre_edit_value = self.get_cell_data(datarn, datacn)
            event_data = event_dict(
                name="end_edit_table",
                sheet=self.parentframe.name,
                cells_table={(datarn, datacn): pre_edit_value},
                key="??",
                value=selection,
                loc=(r, c),
                boxes=self.get_boxes(),
                selected=self.currently_selected(),
            )
            if kwargs["select_function"] is not None:
                kwargs["select_function"](event_data)
            if self.edit_validation_func:
                selection, edited = self.edit_validation_func(event_data), False
                if selection is not None:
                    edited = self.set_cell_data_undo(r, c, value=selection, redraw=not redraw)
            else:
                edited = self.set_cell_data_undo(r, c, value=selection, redraw=not redraw)
            if edited:
                try_binding(self.extra_end_edit_cell_func, event_data)
            self.focus_set()
            self.recreate_all_selection_boxes()
        self.destroy_text_editor("Escape")
        self.destroy_opened_dropdown_window(r, c)
        if redraw:
            self.refresh()

    def get_existing_dropdown_coords(self) -> tuple[int, int] | None:
        if self.existing_dropdown_window is not None:
            return int(self.existing_dropdown_window.r), int(self.existing_dropdown_window.c)
        return None

    def mouseclick_outside_editor_or_dropdown(self) -> tuple:
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
    def destroy_opened_dropdown_window(
        self,
        r: int | None = None,
        c: int | None = None,
        datarn: int | None = None,
        datacn: int | None = None,
    ):
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

    def click_checkbox(
        self,
        r: int,
        c: int,
        datarn: int | None = None,
        datacn: int | None = None,
        undo: bool = True,
        redraw: bool = True,
    ) -> None:
        if datarn is None:
            datarn = self.datarn(r)
        if datacn is None:
            datacn = self.datacn(c)
        kwargs = self.get_cell_kwargs(datarn, datacn, key="checkbox")
        if kwargs["state"] == "normal":
            pre_edit_value = self.get_cell_data(datarn, datacn)
            value = not self.data[datarn][datacn] if isinstance(self.data[datarn][datacn], bool) else False
            self.set_cell_data_undo(
                r,
                c,
                value=value,
                undo=undo,
                cell_resize=False,
                check_input_valid=False,
            )
            event_data = event_dict(
                name="end_edit_table",
                sheet=self.parentframe.name,
                cells_table={(datarn, datacn): pre_edit_value},
                key="??",
                value=value,
                loc=(r, c),
                boxes=self.get_boxes(),
                selected=self.currently_selected(),
            )
            if kwargs["check_function"] is not None:
                kwargs["check_function"](event_data)
            try_binding(self.extra_end_edit_cell_func, event_data)
        if redraw:
            self.refresh()

    # internal event use
    def set_cell_data_undo(
        self,
        r: int = 0,
        c: int = 0,
        datarn: int | None = None,
        datacn: int | None = None,
        value: str = "",
        undo: bool = True,
        cell_resize: bool = True,
        redraw: bool = True,
        check_input_valid: bool = True,
    ) -> bool:
        if datacn is None:
            datacn = self.datacn(c)
        if datarn is None:
            datarn = self.datarn(r)
        event_data = event_dict(
            name="edit_table",
            sheet=self.parentframe.name,
            cells_table={(datarn, datacn): self.get_cell_data(datarn, datacn)},
            boxes=self.get_boxes(),
            selected=self.currently_selected(),
        )
        if not check_input_valid or self.input_valid_for_cell(datarn, datacn, value):
            if self.undo_enabled and undo:
                self.undo_stack.append(ev_stack_dict(event_data))
            self.set_cell_data(datarn, datacn, value)
            if cell_resize and self.cell_auto_resize_enabled:
                self.set_cell_size_to_text(r, c, only_set_if_too_small=True, redraw=redraw, run_binding=True)
            self.sheet_modified(event_data)
            return True
        return False

    def set_cell_data(
        self,
        datarn: int,
        datacn: int,
        value: object,
        kwargs: dict = {},
        expand_sheet: bool = True,
    ) -> None:
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

    def get_value_for_empty_cell(self, datarn: int, datacn: int, r_ops: bool = True, c_ops: bool = True) -> object:
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

    def get_empty_row_seq(
        self,
        datarn: int,
        end: int,
        start: int = 0,
        r_ops: bool = True,
        c_ops: bool = True,
    ) -> list[object]:
        return [self.get_value_for_empty_cell(datarn, datacn, r_ops=r_ops, c_ops=c_ops) for datacn in range(start, end)]

    def fix_row_len(self, datarn: int, datacn: int) -> None:
        self.data[datarn].extend(self.get_empty_row_seq(datarn, end=datacn + 1, start=len(self.data[datarn])))

    def fix_row_values(self, datarn: int, start: int | None = None, end: int | None = None):
        if datarn < len(self.data):
            for datacn, v in enumerate(islice(self.data[datarn], start, end)):
                if not self.input_valid_for_cell(datarn, datacn, v):
                    self.data[datarn][datacn] = self.get_value_for_empty_cell(datarn, datacn)

    def fix_data_len(self, datarn: int, datacn: int | None = None) -> int:
        ncols = self.total_data_cols() if datacn is None else datacn + 1
        self.data.extend([self.get_empty_row_seq(rn, end=ncols, start=0) for rn in range(len(self.data), datarn + 1)])
        return len(self.data)

    def reapply_formatting(self) -> None:
        for c in gen_formatted(self.col_options):
            for r in range(len(self.data)):
                if not (
                    (r, c) in self.cell_options
                    and "format" in self.cell_options[(r, c)]
                    or r in self.row_options
                    and "format" in self.row_options[r]
                ):
                    self.set_cell_data(r, c, value=self.data[r][c])
        for r in gen_formatted(self.row_options):
            for c in range(len(self.data[r])):
                if not ((r, c) in self.cell_options and "format" in self.cell_options[(r, c)]):
                    self.set_cell_data(r, c, value=self.data[r][c])
        for r, c in gen_formatted(self.cell_options):
            if len(self.data) > r and len(self.data[r]) > c:
                self.set_cell_data(r, c, value=self.data[r][c])

    def delete_all_formatting(self, clear_values: bool = False) -> None:
        self.delete_cell_format("all", clear_values=clear_values)
        self.delete_row_format("all", clear_values=clear_values)
        self.delete_column_format("all", clear_values=clear_values)

    def delete_cell_format(self, datarn: str | int = "all", datacn: int = 0, clear_values: bool = False) -> None:
        if isinstance(datarn, str) and datarn.lower() == "all":
            itr = gen_formatted(self.cell_options)
        else:
            itr = ((datarn, datacn),)
        get_val = self.get_value_for_empty_cell
        for key in itr:
            try:
                del self.cell_options[key]["format"]
            except Exception:
                continue
            if clear_values:
                self.set_cell_data(*key, get_val(*key), expand_sheet=False)

    def delete_row_format(self, datarn: str | int = "all", clear_values: bool = False) -> None:
        if isinstance(datarn, str) and datarn.lower() == "all":
            itr = gen_formatted(self.row_options)
        else:
            itr = (datarn,)
        get_val = self.get_value_for_empty_cell
        for datarn in itr:
            try:
                del self.row_options[datarn]["format"]
            except Exception:
                continue
            if clear_values:
                for datacn in range(len(self.data[datarn])):
                    self.set_cell_data(datarn, datacn, get_val(datarn, datacn), expand_sheet=False)

    def delete_column_format(self, datacn: str | int = "all", clear_values: bool = False) -> None:
        if isinstance(datacn, str) and datacn.lower() == "all":
            itr = gen_formatted(self.col_options)
        else:
            itr = (datacn,)
        get_val = self.get_value_for_empty_cell
        for datacn in itr:
            try:
                del self.col_options[datacn]["format"]
            except Exception:
                continue
            if clear_values:
                for datarn in range(len(self.data)):
                    self.set_cell_data(datarn, datacn, get_val(datarn, datacn), expand_sheet=False)

    # deals with possibility of formatter class being in self.data cell
    # if cell is formatted - possibly returns invalid_value kwarg if
    # cell value is not in datatypes kwarg
    # if get displayed is true then Nones are replaced by ""
    def get_valid_cell_data_as_str(self, datarn: int, datacn: int, get_displayed: bool = False, **kwargs) -> str:
        if get_displayed:
            kwargs = self.get_cell_kwargs(datarn, datacn, key="dropdown")
            if kwargs:
                if kwargs["text"] is not None:
                    return f"{kwargs['text']}"
            else:
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
                    # assumed given formatter class has __str__()
                    return f"{value}"
                else:
                    # assumed given formatter class has get_data_with_valid_check()
                    return f"{value.get_data_with_valid_check()}"
        return "" if value is None else f"{value}"

    def get_cell_data(
        self,
        datarn: int,
        datacn: int,
        get_displayed: bool = False,
        none_to_empty_str: bool = False,
        fmt_kw: dict | None = None,
        **kwargs,
    ) -> object:
        if get_displayed:
            return self.get_valid_cell_data_as_str(datarn, datacn, get_displayed=True)
        value = self.data[datarn][datacn] if len(self.data) > datarn and len(self.data[datarn]) > datacn else ""
        kwargs = self.get_cell_kwargs(datarn, datacn, key="format")
        if kwargs and kwargs["formatter"] is not None:
            value = value.value  # assumed given formatter class has value attribute
        if isinstance(fmt_kw, dict):
            value = format_data(value=value, **fmt_kw)
        return "" if (value is None and none_to_empty_str) else value

    def input_valid_for_cell(
        self,
        datarn: int,
        datacn: int,
        value: object,
        check_readonly: bool = True,
        ignore_empty: bool = False,
    ) -> bool:
        if check_readonly and self.get_cell_kwargs(datarn, datacn, key="readonly"):
            return False
        if self.get_cell_kwargs(datarn, datacn, key="format"):
            return True
        if self.cell_equal_to(datarn, datacn, value, ignore_empty=ignore_empty):
            return False
        if self.get_cell_kwargs(datarn, datacn, key="checkbox"):
            return is_bool_like(value)
        kwargs = self.get_cell_kwargs(datarn, datacn, key="dropdown")
        if kwargs and kwargs["validate_input"] and value not in kwargs["values"]:
            return False
        return True

    def cell_equal_to(self, datarn: int, datacn: int, value: object, ignore_empty: bool = False, **kwargs) -> bool:
        v = self.get_cell_data(datarn, datacn)
        kwargs = self.get_cell_kwargs(datarn, datacn, key="format")
        if kwargs and kwargs["formatter"] is None:
            if ignore_empty:
                if not (x := format_data(value=value, **kwargs)) and not v:
                    return False
                return v == x
            return v == format_data(value=value, **kwargs)
        # assumed if there is a formatter class in cell then it has a
        # __eq__() function anyway
        # else if there is not a formatter class in cell and cell is not formatted
        # then compare value as is
        if ignore_empty and not v and not value:
            return False
        return v == value

    def get_cell_clipboard(self, datarn: int, datacn: int) -> str | int | float | bool:
        value = self.data[datarn][datacn] if len(self.data) > datarn and len(self.data[datarn]) > datacn else ""
        kwargs = self.get_cell_kwargs(datarn, datacn, key="format")
        if kwargs:
            if kwargs["formatter"] is None:
                return get_clipboard_data(value, **kwargs)
            else:
                # assumed given formatter class has get_clipboard_data() function
                # and it returns one of above type hints
                return value.get_clipboard_data()
        return f"{value}"

    def get_cell_kwargs(
        self,
        datarn: int,
        datacn: int,
        key: Hashable = "format",
        cell: bool = True,
        row: bool = True,
        column: bool = True,
    ) -> dict:
        if cell and (datarn, datacn) in self.cell_options and key in self.cell_options[(datarn, datacn)]:
            return self.cell_options[(datarn, datacn)][key]
        if row and datarn in self.row_options and key in self.row_options[datarn]:
            return self.row_options[datarn][key]
        if column and datacn in self.col_options and key in self.col_options[datacn]:
            return self.col_options[datacn][key]
        return {}

    def datacn(self, c: int) -> int:
        return c if self.all_columns_displayed else self.displayed_columns[c]

    def datarn(self, r: int) -> int:
        return r if self.all_rows_displayed else self.displayed_rows[r]
