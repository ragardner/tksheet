"""Type definitions used by and useable for annotating code using TkSheet."""

from __future__ import annotations

from typing import NamedTuple, Union

from typing_extensions import Literal, TypeAlias, TypedDict


class Font(NamedTuple):
    family: str
    size: int
    style: str


Binding: TypeAlias = Literal[
    "all",
    "single_select",
    "toggle_select",
    "drag_select",
    "select_all",
    "column_drag_and_drop",
    "row_drag_and_drop",
    "column_select",
    "row_select",
    "column_width_resize",
    "double_click_column_resize",
    "row_width_resize",
    "column_height_resize",
    "arrowkeys",
    "up",
    "down",
    "left",
    "right",
    "prior",
    "next",
    "row_height_resize",
    "double_click_row_resize",
    "right_click_popup_menu",
    "rc_select",
    "rc_insert_column",
    "rc_delete_column",
    "rc_insert_row",
    "rc_delete_row",
    "ctrl_click_select",
    "ctrl_select",
    "copy",
    "cut",
    "paste",
    "delete",
    "undo",
    "edit_cell",
    "edit_header",
    "edit_index",
    "begin_copy",
    "end_copy",
    "begin_cut",
    "end_cut",
    "begin_paste",
    "end_paste",
    "begin_undo",
    "end_undo",
    "begin_delete",
    "end_delete",
    "begin_edit_cell",
    "end_edit_cell",
    "begin_row_index_drag_drop",
    "end_row_index_drag_drop",
    "begin_column_header_drag_drop",
    "end_column_header_drag_drop",
    "begin_delete_rows",
    "end_delete_rows",
    "begin_delete_columns",
    "end_delete_columns",
    "begin_insert_columns",
    "end_insert_columns",
    "begin_insert_rows",
    "end_insert_rows",
    "row_height_resize",
    "column_width_resize",
    "select_all",
    "row_select",
    "column_select",
    "cell_select",
    "drag_select_cells",
    "drag_select_rows",
    "drag_select_columns",
    "shift_cell_select",
    "shift_row_select",
    "shift_column_select",
    "deselect",
    "all_select_events",
    "select",
    "selectevents",
    "select_events",
    "all_modified_events",
    "sheetmodified",
    "sheet_modified",
    "modified_events",
    "modified",
    "bind_all",
    "unbind_all",
]

Canvas: TypeAlias = Literal["all", "row_index", "header", "top_left", "x_scrollbar", "y_scrollbar"]
Align: TypeAlias = Literal["nw", "n", "ne", "w", "center", "e", "sw", "s", "se"]
ScreenUnits: TypeAlias = Union[str, float]
State: TypeAlias = Literal["normal", "disabled"]


class NamedSpan(TypedDict):
    from_r: int | None
    from_c: int | None
    upto_r: int | None
    upto_c: int | None
    type_: str | None
    name: str | None
    table: bool | None
    header: bool | None
    index: bool | None
    kwargs: dict | None


class Theme(TypedDict):
    popup_menu_fg: str
    popup_menu_bg: str
    popup_menu_highlight_bg: str
    popup_menu_highlight_fg: str
    index_hidden_rows_expander_bg: str
    header_hidden_columns_expander_bg: str
    header_bg: str
    header_border_fg: str
    header_grid_fg: str
    header_fg: str
    header_selected_cells_bg: str
    header_selected_cells_fg: str
    index_bg: str
    index_border_fg: str
    index_grid_fg: str
    index_fg: str
    index_selected_cells_bg: str
    index_selected_cells_fg: str
    top_left_bg: str
    top_left_fg: str
    top_left_fg_highlight: str
    table_bg: str
    table_grid_fg: str
    table_fg: str
    table_selected_box_cells_fg: str
    table_selected_box_rows_fg: str
    table_selected_box_columns_fg: str
    table_selected_cells_border_fg: str
    table_selected_cells_bg: str
    table_selected_cells_fg: str
    resizing_line_fg: str
    drag_and_drop_bg: str
    outline_color: str
    header_selected_columns_bg: str
    header_selected_columns_fg: str
    index_selected_rows_bg: str
    index_selected_rows_fg: str
    table_selected_rows_border_fg: str
    table_selected_rows_bg: str
    table_selected_rows_fg: str
    table_selected_columns_border_fg: str
    table_selected_columns_bg: str
    table_selected_columns_fg: str


