# ruff: noqa: F401

"""
tksheet - A Python tkinter table widget
"""

__version__ = "7.2.7"

from .colors import (
    color_map,
)
from .column_headers import ColumnHeaders
from .formatters import (
    Formatter,
    bool_formatter,
    data_to_str,
    float_formatter,
    float_to_str,
    format_data,
    formatter,
    get_clipboard_data,
    get_data_with_valid_check,
    int_formatter,
    is_bool_like,
    is_none_like,
    percentage_formatter,
    percentage_to_str,
    to_bool,
    to_float,
    to_int,
    to_str,
    try_to_bool,
)
from .functions import (
    add_highlight,
    alpha2idx,
    alpha2num,
    consecutive_chunks,
    consecutive_ranges,
    data_to_displayed_idxs,
    displayed_to_data_idxs,
    dropdown_search_function,
    event_dict,
    get_checkbox_dict,
    get_checkbox_kwargs,
    get_csv_str_dialect,
    get_data_from_clipboard,
    get_dropdown_dict,
    get_dropdown_kwargs,
    get_index_of_gap_in_sorted_integer_seq_forward,
    get_index_of_gap_in_sorted_integer_seq_reverse,
    get_n2a,
    get_new_indexes,
    get_seq_without_gaps_at_index,
    insert_items,
    is_contiguous,
    is_iterable,
    move_elements_by_mapping,
    move_elements_to,
    new_tk_event,
    num2alpha,
    rounded_box_coords,
    span_dict,
    tksheet_type_error,
)
from .main_table import MainTable
from .other_classes import (
    DotDict,
    DraggedRowColumn,
    DrawnItem,
    EventDataDict,
    GeneratedMouseEvent,
    Highlight,
    Selected,
    Span,
    SpanRange,
    TextCfg,
)
from .row_index import RowIndex
from .sheet import Dropdown, Sheet
from .sheet_options import new_sheet_options
from .text_editor import (
    TextEditor,
    TextEditorTkText,
)
from .themes import (
    theme_black,
    theme_dark,
    theme_dark_blue,
    theme_dark_green,
    theme_light_blue,
    theme_light_green,
)
from .top_left_rectangle import TopLeftRectangle
from .vars import (
    USER_OS,
    ctrl_key,
    emitted_events,
    falsy,
    nonelike,
    rc_binding,
    symbols_set,
    truthy,
)
