from __future__ import annotations

from .other_classes import (
    DotDict,
    FontTuple,
)
from .themes import (
    theme_light_blue,
)
from .vars import (
    USER_OS,
    ctrl_key,
)


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
            "edit_header_label": "Edit header",
            "edit_header_accelerator": "",
            "edit_index_label": "Edit index",
            "edit_index_accelerator": "",
            "edit_cell_label": "Edit cell",
            "edit_cell_accelerator": "",
            "cut_label": "Cut",
            "cut_accelerator": "Ctrl+X",
            "cut_contents_label": "Cut contents",
            "cut_contents_accelerator": "Ctrl+X",
            "copy_label": "Copy",
            "copy_accelerator": "Ctrl+C",
            "copy_contents_label": "Copy contents",
            "copy_contents_accelerator": "Ctrl+C",
            "paste_label": "Paste",
            "paste_accelerator": "Ctrl+V",
            "delete_label": "Delete",
            "delete_accelerator": "Del",
            "clear_contents_label": "Clear contents",
            "clear_contents_accelerator": "Del",
            "delete_columns_label": "Delete columns",
            "delete_columns_accelerator": "",
            "insert_columns_left_label": "Insert columns left",
            "insert_columns_left_accelerator": "",
            "insert_column_label": "Insert column",
            "insert_column_accelerator": "",
            "insert_columns_right_label": "Insert columns right",
            "insert_columns_right_accelerator": "",
            "delete_rows_label": "Delete rows",
            "delete_rows_accelerator": "",
            "insert_rows_above_label": "Insert rows above",
            "insert_rows_above_accelerator": "",
            "insert_rows_below_label": "Insert rows below",
            "insert_rows_below_accelerator": "",
            "insert_row_label": "Insert row",
            "insert_row_accelerator": "",
            "select_all_label": "Select all",
            "select_all_accelerator": "Ctrl+A",
            "undo_label": "Undo",
            "undo_accelerator": "Ctrl+Z",
            "copy_bindings": [
                f"<{ctrl_key}-c>",
                f"<{ctrl_key}-C>",
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
            ],
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
            "vertical_scroll_borderwidth": 1,
            "horizontal_scroll_borderwidth": 1,
            "vertical_scroll_gripcount": 0,
            "horizontal_scroll_gripcount": 0,
            "vertical_scroll_arrowsize": "",
            "horizontal_scroll_arrowsize": "",
            "set_cell_sizes_on_zoom": False,
            "auto_resize_columns": None,
            "auto_resize_rows": None,
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
            "selected_rows_to_end_of_window": False,
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
            "treeview_indent": "5",
            "rounded_boxes": True,
            "alternate_color": "",
            "min_column_width": 1,
            "max_column_width": float("inf"),
            "max_header_height": float("inf"),
            "max_row_height": float("inf"),
            "max_index_width": float("inf"),
        }
    )
