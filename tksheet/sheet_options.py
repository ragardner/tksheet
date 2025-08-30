from __future__ import annotations

import csv
import tkinter as tk

from .constants import (  # noqa: F401
    ICON_ADD,
    ICON_CLEAR,
    ICON_COPY,
    ICON_COPY_PLAIN,
    ICON_CUT,
    ICON_DEL,
    ICON_EDIT,
    ICON_PASTE,
    ICON_REDO,
    ICON_SELECT_ALL,
    ICON_SORT_ASC,
    ICON_SORT_DESC,
    ICON_UNDO,
    USER_OS,
    alt_key,
    ctrl_key,
)
from .other_classes import DotDict, FontTuple
from .sorting import fast_sort_key, natural_sort_key, version_sort_key  # noqa: F401
from .themes import theme_light_blue


def new_sheet_options() -> DotDict:
    return DotDict(
        {
            **theme_light_blue,
            "popup_menu_font": FontTuple(
                "Calibri",
                13 if USER_OS == "darwin" else 11,
                "normal",
            ),
            "table_font": FontTuple(
                "Calibri",
                13 if USER_OS == "darwin" else 11,
                "normal",
            ),
            "header_font": FontTuple(
                "Calibri",
                13 if USER_OS == "darwin" else 11,
                "normal",
            ),
            "index_font": FontTuple(
                "Calibri",
                13 if USER_OS == "darwin" else 11,
                "normal",
            ),
            # edit header
            "edit_header_label": "Edit header",
            "edit_header_accelerator": "",
            "edit_header_image": tk.PhotoImage(data=ICON_EDIT),
            "edit_header_compound": "left",
            # edit index
            "edit_index_label": "Edit index",
            "edit_index_accelerator": "",
            "edit_index_image": tk.PhotoImage(data=ICON_EDIT),
            "edit_index_compound": "left",
            # edit cell
            "edit_cell_label": "Edit cell",
            "edit_cell_accelerator": "",
            "edit_cell_image": tk.PhotoImage(data=ICON_EDIT),
            "edit_cell_compound": "left",
            # cut
            "cut_label": "Cut",
            "cut_accelerator": "Ctrl+X",
            "cut_image": tk.PhotoImage(data=ICON_CUT),
            "cut_compound": "left",
            # copy
            "copy_label": "Copy",
            "copy_accelerator": "Ctrl+C",
            "copy_image": tk.PhotoImage(data=ICON_COPY),
            "copy_compound": "left",
            # copy plain
            "copy_plain_label": "Copy text",
            "copy_plain_accelerator": "Ctrl+Ins",
            "copy_plain_image": tk.PhotoImage(data=ICON_COPY_PLAIN),
            "copy_plain_compound": "left",
            # paste
            "paste_label": "Paste",
            "paste_accelerator": "Ctrl+V",
            "paste_image": tk.PhotoImage(data=ICON_PASTE),
            "paste_compound": "left",
            # delete
            "delete_label": "Delete",
            "delete_accelerator": "Del",
            "delete_image": tk.PhotoImage(data=ICON_CLEAR),
            "delete_compound": "left",
            # clear contents
            "clear_contents_label": "Clear contents",
            "clear_contents_accelerator": "Del",
            "clear_contents_image": tk.PhotoImage(data=ICON_CLEAR),
            "clear_contents_compound": "left",
            # del columns
            "delete_columns_label": "Delete columns",
            "delete_columns_accelerator": "",
            "delete_columns_image": tk.PhotoImage(data=ICON_DEL),
            "delete_columns_compound": "left",
            # insert columns left
            "insert_columns_left_label": "Insert columns left",
            "insert_columns_left_accelerator": "",
            "insert_columns_left_image": tk.PhotoImage(data=ICON_ADD),
            "insert_columns_left_compound": "left",
            # insert columns right
            "insert_columns_right_label": "Insert columns right",
            "insert_columns_right_accelerator": "",
            "insert_columns_right_image": tk.PhotoImage(data=ICON_ADD),
            "insert_columns_right_compound": "left",
            # insert single column
            "insert_column_label": "Insert column",
            "insert_column_accelerator": "",
            "insert_column_image": tk.PhotoImage(data=ICON_ADD),
            "insert_column_compound": "left",
            # del rows
            "delete_rows_label": "Delete rows",
            "delete_rows_accelerator": "",
            "delete_rows_image": tk.PhotoImage(data=ICON_DEL),
            "delete_rows_compound": "left",
            # insert rows above
            "insert_rows_above_label": "Insert rows above",
            "insert_rows_above_accelerator": "",
            "insert_rows_above_image": tk.PhotoImage(data=ICON_ADD),
            "insert_rows_above_compound": "left",
            # insert rows below
            "insert_rows_below_label": "Insert rows below",
            "insert_rows_below_accelerator": "",
            "insert_rows_below_image": tk.PhotoImage(data=ICON_ADD),
            "insert_rows_below_compound": "left",
            # insert single row
            "insert_row_label": "Insert row",
            "insert_row_accelerator": "",
            "insert_row_image": tk.PhotoImage(data=ICON_ADD),
            "insert_row_compound": "left",
            # sorting
            # labels
            "sort_cells_label": "Sort Asc.",
            "sort_cells_x_label": "Sort row-wise Asc.",
            "sort_row_label": "Sort values Asc.",
            "sort_column_label": "Sort values Asc.",
            "sort_rows_label": "Sort rows Asc.",
            "sort_columns_label": "Sort columns Asc.",
            # reverse labels
            "sort_cells_reverse_label": "Sort Desc.",
            "sort_cells_x_reverse_label": "Sort row-wise Desc.",
            "sort_row_reverse_label": "Sort values Desc.",
            "sort_column_reverse_label": "Sort values Desc.",
            "sort_rows_reverse_label": "Sort rows Desc.",
            "sort_columns_reverse_label": "Sort columns Desc.",
            # accelerators
            "sort_cells_accelerator": "",
            "sort_cells_x_accelerator": "",
            "sort_row_accelerator": "",
            "sort_column_accelerator": "",
            "sort_rows_accelerator": "",
            "sort_columns_accelerator": "",
            # reverse accelerators
            "sort_cells_reverse_accelerator": "",
            "sort_cells_x_reverse_accelerator": "",
            "sort_row_reverse_accelerator": "",
            "sort_column_reverse_accelerator": "",
            "sort_rows_reverse_accelerator": "",
            "sort_columns_reverse_accelerator": "",
            # images
            "sort_cells_image": tk.PhotoImage(data=ICON_SORT_ASC),
            "sort_cells_x_image": tk.PhotoImage(data=ICON_SORT_ASC),
            "sort_row_image": tk.PhotoImage(data=ICON_SORT_ASC),
            "sort_column_image": tk.PhotoImage(data=ICON_SORT_ASC),
            "sort_rows_image": tk.PhotoImage(data=ICON_SORT_ASC),
            "sort_columns_image": tk.PhotoImage(data=ICON_SORT_ASC),
            # compounds
            "sort_cells_compound": "left",
            "sort_cells_x_compound": "left",
            "sort_row_compound": "left",
            "sort_column_compound": "left",
            "sort_rows_compound": "left",
            "sort_columns_compound": "left",
            # reverse images
            "sort_cells_reverse_image": tk.PhotoImage(data=ICON_SORT_DESC),
            "sort_cells_x_reverse_image": tk.PhotoImage(data=ICON_SORT_DESC),
            "sort_row_reverse_image": tk.PhotoImage(data=ICON_SORT_DESC),
            "sort_column_reverse_image": tk.PhotoImage(data=ICON_SORT_DESC),
            "sort_rows_reverse_image": tk.PhotoImage(data=ICON_SORT_DESC),
            "sort_columns_reverse_image": tk.PhotoImage(data=ICON_SORT_DESC),
            # reverse compounds
            "sort_cells_reverse_compound": "left",
            "sort_cells_x_reverse_compound": "left",
            "sort_row_reverse_compound": "left",
            "sort_column_reverse_compound": "left",
            "sort_rows_reverse_compound": "left",
            "sort_columns_reverse_compound": "left",
            # select all
            "select_all_label": "Select all",
            "select_all_accelerator": "Ctrl+A",
            "select_all_image": tk.PhotoImage(data=ICON_SELECT_ALL),
            "select_all_compound": "left",
            # undo
            "undo_label": "Undo",
            "undo_accelerator": "Ctrl+Z",
            "undo_image": tk.PhotoImage(data=ICON_UNDO),
            "undo_compound": "left",
            # redo
            "redo_label": "Redo",
            "redo_accelerator": "Ctrl+Shift+Z",
            "redo_image": tk.PhotoImage(data=ICON_REDO),
            "redo_compound": "left",
            # bindings
            "rc_bindings": ["<2>", "<3>"] if USER_OS == "darwin" else ["<3>"],
            "copy_bindings": [
                f"<{ctrl_key}-c>",
                f"<{ctrl_key}-C>",
            ],
            "copy_plain_bindings": [
                f"<{ctrl_key}-Insert>",
            ],
            "cut_bindings": [
                f"<{ctrl_key}-x>",
                f"<{ctrl_key}-X>",
            ],
            "paste_bindings": [
                f"<{ctrl_key}-v>",
                f"<{ctrl_key}-V>",
            ],
            "undo_bindings": [
                f"<{ctrl_key}-z>",
                f"<{ctrl_key}-Z>",
            ],
            "redo_bindings": [
                f"<{ctrl_key}-Shift-z>",
                f"<{ctrl_key}-Shift-Z>",
            ],
            "delete_bindings": [
                "<Delete>",
            ],
            "select_all_bindings": [
                f"<{ctrl_key}-a>",
                f"<{ctrl_key}-A>",
                f"<{ctrl_key}-Shift-space>",
            ],
            "select_columns_bindings": [
                "<Control-space>",
            ],
            "select_rows_bindings": [
                "<Shift-space>",
            ],
            "row_start_bindings": [
                "<Command-Left>",
                "<Home>",
            ]
            if USER_OS == "darwin"
            else ["<Home>"],
            "table_start_bindings": [
                f"<{ctrl_key}-Home>",
            ],
            "tab_bindings": [
                "<Tab>",
            ],
            "up_bindings": [
                "<Up>",
            ],
            "right_bindings": [
                "<Right>",
            ],
            "down_bindings": [
                "<Down>",
            ],
            "left_bindings": [
                "<Left>",
            ],
            "shift_up_bindings": [
                "<Shift-Up>",
            ],
            "shift_right_bindings": [
                "<Shift-Right>",
            ],
            "shift_down_bindings": [
                "<Shift-Down>",
            ],
            "shift_left_bindings": [
                "<Shift-Left>",
            ],
            "prior_bindings": [
                "<Prior>",
            ],
            "next_bindings": [
                "<Next>",
            ],
            "find_bindings": [
                f"<{ctrl_key}-f>",
                f"<{ctrl_key}-F>",
            ],
            "find_next_bindings": [
                f"<{ctrl_key}-g>",
                f"<{ctrl_key}-G>",
            ],
            "find_previous_bindings": [
                f"<{ctrl_key}-Shift-g>",
                f"<{ctrl_key}-Shift-G>",
            ],
            "toggle_replace_bindings": [
                f"<{ctrl_key}-h>",
                f"<{ctrl_key}-H>",
            ],
            "escape_bindings": [
                "<Escape>",
            ],
            # other
            "vertical_scroll_borderwidth": 1,
            "horizontal_scroll_borderwidth": 1,
            "vertical_scroll_gripcount": 0,
            "horizontal_scroll_gripcount": 0,
            "vertical_scroll_arrowsize": "",
            "horizontal_scroll_arrowsize": "",
            "set_cell_sizes_on_zoom": False,
            "auto_resize_columns": None,
            "auto_resize_rows": None,
            "to_clipboard_dialect": csv.excel_tab,
            "to_clipboard_delimiter": "\t",
            "to_clipboard_quotechar": '"',
            "to_clipboard_lineterminator": "\n",
            "from_clipboard_delimiters": ["\t"],
            "show_dropdown_borders": False,
            "show_default_header_for_empty": True,
            "show_default_index_for_empty": True,
            "default_header_height": "1",
            "default_row_height": "1",
            "default_column_width": 120,
            "default_row_index_width": 70,
            "default_row_index": "numbers",
            "default_header": "letters",
            "page_up_down_select_row": True,
            "paste_can_expand_x": False,
            "paste_can_expand_y": False,
            "paste_insert_column_limit": None,
            "paste_insert_row_limit": None,
            "arrow_key_down_right_scroll_page": False,
            "cell_auto_resize_enabled": True,
            "auto_resize_row_index": True,
            "max_undos": 30,
            "column_drag_and_drop_perform": True,
            "row_drag_and_drop_perform": True,
            "empty_horizontal": 50,
            "empty_vertical": 50,
            "horizontal_grid_to_end_of_window": False,
            "vertical_grid_to_end_of_window": False,
            "show_vertical_grid": True,
            "show_horizontal_grid": True,
            "display_selected_fg_over_highlights": False,
            "show_selected_cells_border": True,
            "edit_cell_tab": "right",
            "edit_cell_return": "down",
            "editor_del_key": "forward",
            "treeview": False,
            "treeview_indent": "2",
            "rounded_boxes": True,
            "alternate_color": "",
            "allow_cell_overflow": False,
            "table_wrap": "c",
            "header_wrap": "c",
            "index_wrap": "c",
            "min_column_width": 1,
            "max_column_width": float("inf"),
            "max_header_height": float("inf"),
            "max_row_height": float("inf"),
            "max_index_width": float("inf"),
            "show_top_left": None,
            "sort_key": natural_sort_key,
            "tooltips": False,
            "user_can_create_notes": False,
            "note_corners": False,
            "tooltip_width": 210,
            "tooltip_height": 210,
            "tooltip_hover_delay": 1200,
        }
    )
