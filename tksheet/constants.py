from __future__ import annotations

from platform import system as get_os
from typing import Hashable

USER_OS: str = f"{get_os()}".lower()
ctrl_key: str = "Command" if USER_OS == "darwin" else "Control"
rc_binding: str = "<2>" if USER_OS == "darwin" else "<3>"
symbols_set: set[str] = set("""!#$%&'()*+,-./:;"@[]^_`{|}~>?= \\""")
nonelike: set[Hashable] = {None, "none", ""}
truthy: set[Hashable] = {True, "true", "t", "yes", "y", "on", "1"}
falsy: set[Hashable] = {False, "false", "f", "no", "n", "off", "0"}
_test_str: str = "aiW_-|"

val_modifying_options: set[str] = {"checkbox", "format", "dropdown"}

named_span_types: set[str] = {
    "format",
    "highlight",
    "dropdown",
    "checkbox",
    "readonly",
    "align",
}

emitted_events: set[str] = {
    "<<SheetModified>>",
    "<<SheetRedrawn>>",
    "<<SheetSelect>>",
    "<<Copy>>",
    "<<Cut>>",
    "<<Paste>>",
    "<<Delete>>",
    "<<Undo>>",
    "<<Redo>>",
    "<<SelectAll>>",
}

align_helper: dict[str, str] = {
    "w": "nw",
    "e": "ne",
    "center": "n",
    "n": "center",
    "nw": "left",
    "ne": "right",
}

align_value_error: str = "Align must be one of the following values: c, center, w, nw, west, left, e, ne, east, right"
font_value_error: str = "Argument must be font, size and 'normal', 'bold' or'italic' e.g. ('Carlito',12,'normal')"

backwards_compatibility_keys: dict[str, str] = {
    "font": "table_font",
}

text_editor_newline_bindings: set[str] = {
    "<Alt-Return>",
    "<Alt-KP_Enter>",
    "<Shift-Return>",
    f"<{ctrl_key}-Return>",
}
if USER_OS == "darwin":
    text_editor_newline_bindings.add("<Option-Return>")

text_editor_close_bindings: dict[str, str] = {
    "<Tab>": "Tab",
    "<Return>": "Return",
    "<KP_Enter>": "Return",
    "<Escape>": "Escape",
}

text_editor_to_unbind: set[str] = (
    text_editor_newline_bindings | set(text_editor_close_bindings) | {"<FocusOut>", "<KeyRelease>"}
)

scrollbar_options_keys: set[str] = {
    "vertical_scroll_background",
    "horizontal_scroll_background",
    "vertical_scroll_troughcolor",
    "horizontal_scroll_troughcolor",
    "vertical_scroll_lightcolor",
    "horizontal_scroll_lightcolor",
    "vertical_scroll_darkcolor",
    "horizontal_scroll_darkcolor",
    "vertical_scroll_relief",
    "horizontal_scroll_relief",
    "vertical_scroll_troughrelief",
    "horizontal_scroll_troughrelief",
    "vertical_scroll_bordercolor",
    "horizontal_scroll_bordercolor",
    "vertical_scroll_borderwidth",
    "horizontal_scroll_borderwidth",
    "vertical_scroll_gripcount",
    "horizontal_scroll_gripcount",
    "vertical_scroll_arrowsize",
    "horizontal_scroll_arrowsize",
    "vertical_scroll_active_bg",
    "horizontal_scroll_active_bg",
    "vertical_scroll_not_active_bg",
    "horizontal_scroll_not_active_bg",
    "vertical_scroll_pressed_bg",
    "horizontal_scroll_pressed_bg",
    "vertical_scroll_active_fg",
    "horizontal_scroll_active_fg",
    "vertical_scroll_not_active_fg",
    "horizontal_scroll_not_active_fg",
    "vertical_scroll_pressed_fg",
    "horizontal_scroll_pressed_fg",
}

bind_add_columns: set[str] = {
    "all",
    "rc_insert_column",
    "insert_column",
    "add_column",
    "insert_columns",
    "add_columns",
    "insert column",
    "add column",
}

bind_del_columns: set[str] = {
    "all",
    "rc_delete_column",
    "delete_column",
    "del_column",
    "delete_columns",
    "del_columns",
    "delete columns",
    "del columns",
}

bind_add_rows: set[str] = {
    "all",
    "rc_insert_row",
    "insert_row",
    "add_row",
    "insert_rows",
    "add_rows",
    "insert row",
    "add row",
}

bind_del_rows: set[str] = {
    "all",
    "rc_delete_row",
    "delete_row",
    "del_row",
    "delete_rows",
    "del_rows",
    "delete rows",
    "del rows",
}

BINDING_TO_ATTR = {
    "replace_all": ("MT", "extra_end_replace_all_func"),
    # begin sort cells
    "begin_sort_cells": ("MT", "extra_begin_sort_cells_func"),
    # end sort cells
    "sort_cells": ("MT", "extra_end_sort_cells_func"),
    "end_sort_cells": ("MT", "extra_end_sort_cells_func"),
    # begin sort rows
    "begin_sort_rows": ("CH", "ch_extra_begin_sort_rows_func"),
    # end sort rows
    "sort_rows": ("CH", "ch_extra_end_sort_rows_func"),
    "end_sort_rows": ("CH", "ch_extra_end_sort_rows_func"),
    # begin sort cols
    "begin_sort_columns": ("RI", "ri_extra_begin_sort_cols_func"),
    # end sort cols
    "sort_columns": ("RI", "ri_extra_end_sort_cols_func"),
    "end_sort_columns": ("RI", "ri_extra_end_sort_cols_func"),
    # begin copy
    "begin_copy": ("MT", "extra_begin_ctrl_c_func"),
    "begin_ctrl_c": ("MT", "extra_begin_ctrl_c_func"),
    # end copy
    "ctrl_c": ("MT", "extra_end_ctrl_c_func"),
    "end_copy": ("MT", "extra_end_ctrl_c_func"),
    "end_ctrl_c": ("MT", "extra_end_ctrl_c_func"),
    "copy": ("MT", "extra_end_ctrl_c_func"),
    # begin cut
    "begin_cut": ("MT", "extra_begin_ctrl_x_func"),
    "begin_ctrl_x": ("MT", "extra_begin_ctrl_x_func"),
    # end cut
    "ctrl_x": ("MT", "extra_end_ctrl_x_func"),
    "end_cut": ("MT", "extra_end_ctrl_x_func"),
    "end_ctrl_x": ("MT", "extra_end_ctrl_x_func"),
    "cut": ("MT", "extra_end_ctrl_x_func"),
    # begin paste
    "begin_paste": ("MT", "extra_begin_ctrl_v_func"),
    "begin_ctrl_v": ("MT", "extra_begin_ctrl_v_func"),
    # end paste
    "ctrl_v": ("MT", "extra_end_ctrl_v_func"),
    "end_paste": ("MT", "extra_end_ctrl_v_func"),
    "end_ctrl_v": ("MT", "extra_end_ctrl_v_func"),
    "paste": ("MT", "extra_end_ctrl_v_func"),
    # begin undo/redo
    "begin_undo": ("MT", "extra_begin_ctrl_z_func"),
    "begin_ctrl_z": ("MT", "extra_begin_ctrl_z_func"),
    # end undo/redo
    "ctrl_z": ("MT", "extra_end_ctrl_z_func"),
    "end_undo": ("MT", "extra_end_ctrl_z_func"),
    "end_ctrl_z": ("MT", "extra_end_ctrl_z_func"),
    "undo": ("MT", "extra_end_ctrl_z_func"),
    # begin del key
    "begin_delete_key": ("MT", "extra_begin_delete_key_func"),
    "begin_delete": ("MT", "extra_begin_delete_key_func"),
    # end del key
    "delete_key": ("MT", "extra_end_delete_key_func"),
    "end_delete": ("MT", "extra_end_delete_key_func"),
    "end_delete_key": ("MT", "extra_end_delete_key_func"),
    "delete": ("MT", "extra_end_delete_key_func"),
    # begin edit table
    "begin_edit_cell": ("MT", "extra_begin_edit_cell_func"),
    "begin_edit_table": ("MT", "extra_begin_edit_cell_func"),
    # end edit table
    "end_edit_cell": ("MT", "extra_end_edit_cell_func"),
    "edit_cell": ("MT", "extra_end_edit_cell_func"),
    "edit_table": ("MT", "extra_end_edit_cell_func"),
    # begin edit header
    "begin_edit_header": ("CH", "extra_begin_edit_cell_func"),
    # end edit header
    "end_edit_header": ("CH", "extra_end_edit_cell_func"),
    "edit_header": ("CH", "extra_end_edit_cell_func"),
    # begin edit index
    "begin_edit_index": ("RI", "extra_begin_edit_cell_func"),
    # end edit index
    "end_edit_index": ("RI", "extra_end_edit_cell_func"),
    "edit_index": ("RI", "extra_end_edit_cell_func"),
    # begin move rows
    "begin_row_index_drag_drop": ("RI", "ri_extra_begin_drag_drop_func"),
    "begin_move_rows": ("RI", "ri_extra_begin_drag_drop_func"),
    # end move rows
    "row_index_drag_drop": ("RI", "ri_extra_end_drag_drop_func"),
    "move_rows": ("RI", "ri_extra_end_drag_drop_func"),
    "end_move_rows": ("RI", "ri_extra_end_drag_drop_func"),
    "end_row_index_drag_drop": ("RI", "ri_extra_end_drag_drop_func"),
    # begin move cols
    "begin_column_header_drag_drop": ("CH", "ch_extra_begin_drag_drop_func"),
    "begin_move_columns": ("CH", "ch_extra_begin_drag_drop_func"),
    # end move cols
    "column_header_drag_drop": ("CH", "ch_extra_end_drag_drop_func"),
    "move_columns": ("CH", "ch_extra_end_drag_drop_func"),
    "end_move_columns": ("CH", "ch_extra_end_drag_drop_func"),
    "end_column_header_drag_drop": ("CH", "ch_extra_end_drag_drop_func"),
    # begin del rows
    "begin_rc_delete_row": ("MT", "extra_begin_del_rows_rc_func"),
    "begin_delete_rows": ("MT", "extra_begin_del_rows_rc_func"),
    # end del rows
    "rc_delete_row": ("MT", "extra_end_del_rows_rc_func"),
    "end_rc_delete_row": ("MT", "extra_end_del_rows_rc_func"),
    "end_delete_rows": ("MT", "extra_end_del_rows_rc_func"),
    "delete_rows": ("MT", "extra_end_del_rows_rc_func"),
    # begin del cols
    "begin_rc_delete_column": ("MT", "extra_begin_del_cols_rc_func"),
    "begin_delete_columns": ("MT", "extra_begin_del_cols_rc_func"),
    # end del cols
    "rc_delete_column": ("MT", "extra_end_del_cols_rc_func"),
    "end_rc_delete_column": ("MT", "extra_end_del_cols_rc_func"),
    "end_delete_columns": ("MT", "extra_end_del_cols_rc_func"),
    "delete_columns": ("MT", "extra_end_del_cols_rc_func"),
    # begin add cols
    "begin_rc_insert_column": ("MT", "extra_begin_insert_cols_rc_func"),
    "begin_insert_column": ("MT", "extra_begin_insert_cols_rc_func"),
    "begin_insert_columns": ("MT", "extra_begin_insert_cols_rc_func"),
    "begin_add_column": ("MT", "extra_begin_insert_cols_rc_func"),
    "begin_rc_add_column": ("MT", "extra_begin_insert_cols_rc_func"),
    "begin_add_columns": ("MT", "extra_begin_insert_cols_rc_func"),
    # end add cols
    "rc_insert_column": ("MT", "extra_end_insert_cols_rc_func"),
    "end_rc_insert_column": ("MT", "extra_end_insert_cols_rc_func"),
    "end_insert_column": ("MT", "extra_end_insert_cols_rc_func"),
    "end_insert_columns": ("MT", "extra_end_insert_cols_rc_func"),
    "rc_add_column": ("MT", "extra_end_insert_cols_rc_func"),
    "end_rc_add_column": ("MT", "extra_end_insert_cols_rc_func"),
    "end_add_column": ("MT", "extra_end_insert_cols_rc_func"),
    "end_add_columns": ("MT", "extra_end_insert_cols_rc_func"),
    "add_columns": ("MT", "extra_end_insert_cols_rc_func"),
    # begin add rows
    "begin_rc_insert_row": ("MT", "extra_begin_insert_rows_rc_func"),
    "begin_insert_row": ("MT", "extra_begin_insert_rows_rc_func"),
    "begin_insert_rows": ("MT", "extra_begin_insert_rows_rc_func"),
    "begin_rc_add_row": ("MT", "extra_begin_insert_rows_rc_func"),
    "begin_add_row": ("MT", "extra_begin_insert_rows_rc_func"),
    "begin_add_rows": ("MT", "extra_begin_insert_rows_rc_func"),
    # end add rows
    "rc_insert_row": ("MT", "extra_end_insert_rows_rc_func"),
    "end_rc_insert_row": ("MT", "extra_end_insert_rows_rc_func"),
    "end_insert_row": ("MT", "extra_end_insert_rows_rc_func"),
    "end_insert_rows": ("MT", "extra_end_insert_rows_rc_func"),
    "rc_add_row": ("MT", "extra_end_insert_rows_rc_func"),
    "end_rc_add_row": ("MT", "extra_end_insert_rows_rc_func"),
    "end_add_row": ("MT", "extra_end_insert_rows_rc_func"),
    "end_add_rows": ("MT", "extra_end_insert_rows_rc_func"),
    "add_rows": ("MT", "extra_end_insert_rows_rc_func"),
    # resize
    "row_height_resize": ("RI", "row_height_resize_func"),
    "column_width_resize": ("CH", "column_width_resize_func"),
    # selection
    "cell_select": ("MT", "selection_binding_func"),
    "select_all": ("MT", "select_all_binding_func"),
    "all_select": ("MT", "select_all_binding_func"),
    "row_select": ("RI", "selection_binding_func"),
    "column_select": ("CH", "selection_binding_func"),
    "col_select": ("CH", "selection_binding_func"),
    "drag_select_cells": ("MT", "drag_selection_binding_func"),
    "drag_select_rows": ("RI", "drag_selection_binding_func"),
    "drag_select_columns": ("CH", "drag_selection_binding_func"),
    "shift_cell_select": ("MT", "shift_selection_binding_func"),
    "shift_row_select": ("RI", "shift_selection_binding_func"),
    "shift_column_select": ("CH", "shift_selection_binding_func"),
    "ctrl_cell_select": ("MT", "ctrl_selection_binding_func"),
    "ctrl_row_select": ("RI", "ctrl_selection_binding_func"),
    "ctrl_column_select": ("CH", "ctrl_selection_binding_func"),
    "deselect": ("MT", "deselection_binding_func"),
}

MODIFIED_BINDINGS = [
    ("MT", "extra_end_replace_all_func"),
    ("MT", "extra_end_sort_cells_func"),
    ("CH", "ch_extra_end_sort_rows_func"),
    ("RI", "ri_extra_end_sort_cols_func"),
    ("MT", "extra_end_ctrl_c_func"),
    ("MT", "extra_end_ctrl_x_func"),
    ("MT", "extra_end_ctrl_v_func"),
    ("MT", "extra_end_ctrl_z_func"),
    ("MT", "extra_end_delete_key_func"),
    ("RI", "ri_extra_end_drag_drop_func"),
    ("CH", "ch_extra_end_drag_drop_func"),
    ("MT", "extra_end_del_rows_rc_func"),
    ("MT", "extra_end_del_cols_rc_func"),
    ("MT", "extra_end_insert_cols_rc_func"),
    ("MT", "extra_end_insert_rows_rc_func"),
    ("MT", "extra_end_edit_cell_func"),
    ("CH", "extra_end_edit_cell_func"),
    ("RI", "extra_end_edit_cell_func"),
]

visited = set()
ALL_BINDINGS = []
SELECT_BINDINGS = []
for o, b in BINDING_TO_ATTR.values():
    if b not in visited:
        ALL_BINDINGS.append((o, b))
        visited.add(b)
    if "select" in b and (o, b) not in visited:
        SELECT_BINDINGS.append((o, b))
        visited.add((o, b))
visited = set()
