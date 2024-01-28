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
)


def new_sheet_options() -> DotDict:
    return DotDict(
        {
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
            "edit_header": DotDict(
                {
                    "label": "Edit header",
                    "accelerator": "",
                }
            ),
            "edit_index": DotDict(
                {
                    "label": "Edit index",
                    "accelerator": "",
                }
            ),
            "edit_cell": DotDict(
                {
                    "label": "Edit cell",
                    "accelerator": "",
                }
            ),
            "cut": DotDict(
                {
                    "label": "Cut",
                    "accelerator": "Ctrl+X",
                }
            ),
            "cut_contents": DotDict(
                {
                    "label": "Cut contents",
                    "accelerator": "Ctrl+X",
                }
            ),
            "copy_": DotDict(
                {
                    "label": "Copy",
                    "accelerator": "Ctrl+C",
                }
            ),
            "copy_contents": DotDict(
                {
                    "label": "Copy contents",
                    "accelerator": "Ctrl+C",
                }
            ),
            "paste": DotDict(
                {
                    "label": "Paste",
                    "accelerator": "Ctrl+V",
                }
            ),
            "delete": DotDict(
                {
                    "label": "Delete",
                    "accelerator": "Del",
                }
            ),
            "clear_contents": DotDict(
                {
                    "label": "Clear contents",
                    "accelerator": "Del",
                }
            ),
            "delete_columns": DotDict(
                {
                    "label": "Delete columns",
                    "accelerator": "",
                }
            ),
            "insert_columns_left": DotDict(
                {
                    "label": "Insert columns left",
                    "accelerator": "",
                }
            ),
            "insert_column": DotDict(
                {
                    "label": "Insert column",
                    "accelerator": "",
                }
            ),
            "insert_columns_right": DotDict(
                {
                    "label": "Insert columns right",
                    "accelerator": "",
                }
            ),
            "delete_rows": DotDict(
                {
                    "label": "Delete rows",
                    "accelerator": "",
                }
            ),
            "insert_rows_above": DotDict(
                {
                    "label": "Insert rows above",
                    "accelerator": "",
                }
            ),
            "insert_rows_below": DotDict(
                {
                    "label": "Insert rows below",
                    "accelerator": "",
                }
            ),
            "insert_row": DotDict(
                {
                    "label": "Insert row",
                    "accelerator": "",
                }
            ),
            "popup_menu_fg": theme_light_blue["popup_menu_fg"],
            "popup_menu_bg": theme_light_blue["popup_menu_bg"],
            "popup_menu_highlight_bg": theme_light_blue["popup_menu_highlight_bg"],
            "popup_menu_highlight_fg": theme_light_blue["popup_menu_highlight_fg"],
            "index_hidden_rows_expander_bg": theme_light_blue["index_hidden_rows_expander_bg"],
            "header_hidden_columns_expander_bg": theme_light_blue["header_hidden_columns_expander_bg"],
            "header_bg": theme_light_blue["header_bg"],
            "header_border_fg": theme_light_blue["header_border_fg"],
            "header_grid_fg": theme_light_blue["header_grid_fg"],
            "header_fg": theme_light_blue["header_fg"],
            "header_selected_cells_bg": theme_light_blue["header_selected_cells_bg"],
            "header_selected_cells_fg": theme_light_blue["header_selected_cells_fg"],
            "index_bg": theme_light_blue["index_bg"],
            "index_border_fg": theme_light_blue["index_border_fg"],
            "index_grid_fg": theme_light_blue["index_grid_fg"],
            "index_fg": theme_light_blue["index_fg"],
            "index_selected_cells_bg": theme_light_blue["index_selected_cells_bg"],
            "index_selected_cells_fg": theme_light_blue["index_selected_cells_fg"],
            "top_left_bg": theme_light_blue["top_left_bg"],
            "top_left_fg": theme_light_blue["top_left_fg"],
            "top_left_fg_highlight": theme_light_blue["top_left_fg_highlight"],
            "table_bg": theme_light_blue["table_bg"],
            "table_grid_fg": theme_light_blue["table_grid_fg"],
            "table_fg": theme_light_blue["table_fg"],
            "table_selected_box_cells_fg": theme_light_blue["table_selected_box_cells_fg"],
            "table_selected_box_rows_fg": theme_light_blue["table_selected_box_rows_fg"],
            "table_selected_box_columns_fg": theme_light_blue["table_selected_box_columns_fg"],
            "table_selected_cells_border_fg": theme_light_blue["table_selected_cells_border_fg"],
            "table_selected_cells_bg": theme_light_blue["table_selected_cells_bg"],
            "table_selected_cells_fg": theme_light_blue["table_selected_cells_fg"],
            "resizing_line_fg": theme_light_blue["resizing_line_fg"],
            "drag_and_drop_bg": theme_light_blue["drag_and_drop_bg"],
            "outline_color": theme_light_blue["outline_color"],
            "header_selected_columns_bg": theme_light_blue["header_selected_columns_bg"],
            "header_selected_columns_fg": theme_light_blue["header_selected_columns_fg"],
            "index_selected_rows_bg": theme_light_blue["index_selected_rows_bg"],
            "index_selected_rows_fg": theme_light_blue["index_selected_rows_fg"],
            "table_selected_rows_border_fg": theme_light_blue["table_selected_rows_border_fg"],
            "table_selected_rows_bg": theme_light_blue["table_selected_rows_bg"],
            "table_selected_rows_fg": theme_light_blue["table_selected_rows_fg"],
            "table_selected_columns_border_fg": theme_light_blue["table_selected_columns_border_fg"],
            "table_selected_columns_bg": theme_light_blue["table_selected_columns_bg"],
            "table_selected_columns_fg": theme_light_blue["table_selected_columns_fg"],
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
            "page_up_down_select_row": True,
            "expand_sheet_if_paste_too_big": False,
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
        }
    )
