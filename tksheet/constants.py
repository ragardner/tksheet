from __future__ import annotations

from platform import system as get_os
from typing import Hashable

USER_OS: str = f"{get_os()}".lower()
ctrl_key: str = "Command" if USER_OS == "darwin" else "Control"
alt_key: str = "Option" if USER_OS == "darwin" else "Alt"
symbols_set: set[str] = set("""!#$%&'()*+,-./:;"@[]^_`{|}~>?= \\""")
nonelike: set[Hashable] = {None, "none", ""}
truthy: set[Hashable] = {True, "true", "t", "yes", "y", "on", "1"}
falsy: set[Hashable] = {False, "false", "f", "no", "n", "off", "0"}
_test_str: str = "0"

val_modifying_options: set[str] = {"checkbox", "format", "dropdown"}

named_span_types: set[str] = {
    "format",
    "highlight",
    "dropdown",
    "checkbox",
    "readonly",
    "align",
    "note",
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
    f"<{alt_key}-Return>",
    f"<{alt_key}-KP_Enter>",
    "<Shift-Return>",
    f"<{ctrl_key}-Return>",
}

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

ICON_CUT = (
    """iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAAAXNSR0IArs4c6QAAAxxJREFUSEuFVstq"""
    """20AUPVcqBLLqoiZxVvK2u1JKoFCIocs0/5BlIP6AkWyJOI30AQ5k6X5Dkl1psiwhP6FV3JKuA1lYt7nz"""
    """sEaW3Qps7NHoPs4594yIADDal163XwRe+Vs/aQOsjkOgdTfqtOtKaBbV2rUoTvbZu+Oi+IAKIwD79vFr"""
    """BHyaqvj+35WuT7boQIJThTvTNFcEAQaBRirA7ihW9x4idURbXFYUUQhMZU8Wq76BXaLYDsZfi0sAX5h5"""
    """+uf37EiW32x3Lwh0COAqTdTBIqqPx0s5WV5ErxhTrrAHwrcsUYeOVz9BJdU+/nrYnEzOn4WDwfFgo9Pd"""
    """eSKARkMVCtfLpevKmafEtMeMMh2q3lJ+88j4LK8A4sfZw+bk3CYYDDY62ztP0muaqLApOEJW5FHIGpY9"""
    """AsrRUPWs3OpmXbbxmQ/Rw5Hg19nufgfw6eXzI03izwZT04YJTjo4gDJLVM/Iuk246cAo6M5oiqUbWRaS"""
    """5ZrNA3w8UaqUP4J5CJ6CSQdPk7jnR27MUJ2UMC7yJZnyDUBvCegKvvMQfUngYDHBVW+h8zUDu+q+I5+z"""
    """PI/Cim7BiEBcgqiEVguXWRL3ajk2FNCY8P+MpOBdRGGFW4AjzYJWS9yrLaRZY0NFNV5+nppMjbmWoqcW"""
    """Dcu6q2lOBopmbOHLrbME94bIqMUmHStVtg1xqRu/g7YX8Q2IXoPxHnaI6snPxVXKeYD+iYrLlfqUHS7B"""
    """OPe9CJVeZwQgqZdmaaJ2XKM+ZKIkk0SVzhBdEc673CRfAtTwos7Wzk8Q3rW8yPEixGt1eUlWIG4TFMaL"""
    """tFVMnmXRedFqq2iqS8NF6J/EcemmXWLUdu28SJvd5FkgOrZepM1Oe5Hg1fQCI2G+BShycGnyLU4LFTW9"""
    """aHYkGzpb3QsiOiTgajRUB2JkjerssaqtQ88JIoCtdej6fZLFJsh6kSXZHjgIsJsqdW8ecWe1L0frrBUy"""
    """2ZMmSs4Q00NLpnOMQLxvT7RrBDh1wRf8rTumW1PrTrQ1JtVAu/X24C+0brr6m28VvrO2sF6QVr/C+K8t"""
    """vvbr34S/2B2/MddUJi4AAAAASUVORK5CYII="""
)


ICON_EDIT = (
    """iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAAAXNSR0IArs4c6QAAAn5JREFUSEuFVu1h"""
    """qzAMPJEMQjYhk7x2DEIDoR9hDPomSTd5HiRYr5INCHBa/yK29XGnkxyCLALA+jWtzVbcSFxNO4ie5L75"""
    """XEWxJ2nPeqN9e79k2a5hhgPha8jQtmXpJHNSO2v8Y6YWLKHprvmO0YO5CCeyJElyw4Dnti6/DIJo/Iib"""
    """yCWB1UXTdfme0TOjAODuGY5imt39E2VZI0jqU6l7CZIe8yEnddfFzIPzuioPoxMJvPP4N+6HABtaEpzF"""
    """i0oL0MOrc1luyPjYlicnP5q3a7HL6CYw65eSQg02UQhKxapAmp1yHjKPAXKA3T2jIzGgwRkFD75t6upi"""
    """AliytsSNzkmMGa5+KQ9C1d7jBnAuhQXBiXP6RnWOtG0DxKwtgpkWVYs6F9haj6uiugHIo4pcXZ1iTYSJ"""
    """5JqJm2gRzokXxkLumrbgfO6nHxutkezAPZiiWk6HUaaS/6oPJmTrkiYxKL9zE01SHLEF5xTURFaqUSBT"""
    """3yTcr2E3VXmwSgvnEVks+Jz1soc2jRaaiHtiKjg2kTVR57EPVC2x4JYWG2zRaHPmXND3PAlSmyH/hizF"""
    """9WLYvV67W+CUXVOdIi3BbOHc0mJn1GZyIk5THbkfFx1Sq9mSVMsKmUW5RjHV4PWjkwGVDx7H9hzG7Dg1"""
    """7Xiwgy2R8ObxsQFUKHVVTnuhD+bZI84fjUE7k22vqLPX9+4JxD1An3VVPmsg6YN5asY+2EpwMShtPcZv"""
    """MWmluIzCe9/K71iLMSmnffD4iVBabNZK7vgEaIBO1KPStG+/Y+//NufqspHfhqflxoIiDSAUZfwHTLk4"""
    """9Xv6HB+QxGs0xUvHsWjGlzoth1//zjwys538H8mvnjDSt4tXAAAAAElFTkSuQmCC"""
)


ICON_SORT_DESC = (
    """iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAAAXNSR0IArs4c6QAAAf5JREFUSEuNVW12"""
    """gjAQnMWDVG+CJ+nzGJEWEIsco97EHoWDKGmzhLD5svL8ISG7szszmxCSD4GgodMfE6sEiN3yjdYXm5QX"""
    """fAC5RyZakXwAWYEHwMF2bzqpV9vzLl2hBi6ouhmG7UbjGxrbR4F9p9S40JWuVXZv6xeFuqrNpyU5aZQ2"""
    """6Xi3IK71iMYnFEm+ROUlrxsE4t94J+xPRzXyekKnSHmvA1E5JpQgmERbG7T8Hw1dJ2VB4NNiMaWMrhau"""
    """qOuHGzRM5WNdqV3XD8zQo8BuM+E2A+qxqY67mbq84Embns59S0XxbpKb4K6/cJ66UsTUTWjs+yF2mk+Q"""
    """b33xTSIvHRiAV/wum/J7exUgIvm1aReD5rd57getNVB/zB10F6ORLue0wmLsNGu3JQXhp67Ufj4RhO3k"""
    """MRBSZDViHbxntbI7jQi4flbqwKll5zKQAWwHuemNPJkYhmwHIUVebKjieoRFfogBbHBOg9n/QoOId1nK"""
    """okE0OEAIEGsQCAtc64V3cTtkXOQPWkStXUjr5ysmhshPI0XOAaQHMLis/LHWxrp8o537ixkDPip4Dobh"""
    """himYAyLrd9mSPKVI2lScr38o7fmrLah4qz+OBzNabde3tCka3rXS73hPONS5OHFN5ipKkJW/a5x1XFRu"""
    """7z85AlSpwUJRwqZPBUwMWno/8Au5DCkvTFG1swAAAABJRU5ErkJggg=="""
)


ICON_SORT_ASC = (
    """iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAAAXNSR0IArs4c6QAAAhZJREFUSEuVVu11"""
    """wjAMPAGDhE3KKIyRBBIaXuIxYBPYpCwCbmU7sfyV98qfFlv2Sbo7GcLyIQA6/GaWwnUT4Jb8DoGg3Wm5"""
    """yesxQPk+f3cmRgLLO1MAv7tkmj8sA8P/ZeECwAW53WGaHrzSNfWB/w5KPfDBl0yeCM9zWx+Wzrqzpl2i"""
    """Uy5T2UfgOirNjHRtTRx7uY4X2lBvg+2HgPv5VB8FddmOZ2olXMdpAUi6FpQhhJBwE5Ac4vgKGjLqKhGb"""
    """TW/W4z8Ahkk9oCUHGqap3EfiBOg58xXKtyCGYVTGFGscREfvneEj7F9BRYAB0EB3ci0qaTVjxpiibBeH"""
    """cQoq8B7PObnkiZkDqV+XqVeRrWDmwGybvkveHR9ev46PFZIlB3zuexgvtN30Zf8GFrif2/rovOJh5cDK"""
    """+6CsVakcKWvj0rm/kq/YyQEHCeGxcubJyup1AHb2EN7A8btpXpLkXqlqq3HjVLqmORTHdYZLK1MysjRG"""
    """AvB6b/Rh96Efzvq9wd5c/sEXSN/7tjmm1STDJHgzzJd+UtX2AwapQHgBqCy4fkKTAe7aep9MMz/4Yo9Z"""
    """KiQ5vZosCHRlZ6vVJIFe51O9j6dmxmOJyAIAvrLjSjSDuAqCzOVzuGYu/3ymo+IvZyZ1p3HTGhW3JX6h"""
    """bGUhWCLTuXXJ4UySPqa0WZ7Zi4pCM2RUUe6I/DESRa2MipSt9Zcrsuty/BfWPCAzd/TdsgAAAABJRU5E"""
    """rkJggg=="""
)


ICON_PASTE = (
    """iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAAAXNSR0IArs4c6QAAAlFJREFUSEt9Ve2h"""
    """oyAQnI2NmE5MJe9SBppINKeWkXeVJJ08CjnlDgRcRPWXIDjsfCyE8BAAvQzX83ufkx18IYGWoXtj3+Vz"""
    """KLITXgTkEbQZkAaB1N8J1+YuPhYn/rc9r5lyDwcgEDSafnhD6yJaltao6kqcN0oPmAHZQDye3YNOpy8A"""
    """ud8kK0Fb5LXdwKfV/xqauiq/OdVRBc+u/6U1XqEuR0VdlaYkQMdCPLtBb1B3nUFmzmKAvn9rjWKccDG8"""
    """tpYiFHvl2/kTPrUQFzkMeTbhB4CSlTh74Aig7Xo7X1fCil8bkTPIAxA1WpHLj6Gl/T0Y5e3+zQo4wOGp"""
    """dz623WAqyBeAlT28aL4CTzmnf8OJAc7Q1AhhxI6MExYkALGBV0afP1qdoDESXZtSKGMEY/GgAT/RAlBS"""
    """SAlXywVnnpp3MiOokXAxIFGVfOBtFyg6rGDeKfshzzTeLjdqPOHSCqH8IZiLAFuBBuqbqyARM86BH8mh"""
    """z7OJjKVzECyI12IGcCv3Kghc+0zYAB76TNU3cV71ImABKMkKxQ7czC1EHv7WN0Giz70Sl1nw8BDWQeNi"""
    """+uAsfWZpjrLv80zTywVyI8luLbcpP+mSg7il27SbFqH1izQVGkg6K7sP9imKpY3jMNvUtvSI9w0XxRRx"""
    """6pJMMPv69p7cCazY8C8ucsL1NmcBajHFKsncpiyV17oS396+CUVJAAF3vb71NDXyXj02u2ly4dhVzvR7"""
    """3l/NJwD8dO7SKED4Wu6ADYAEzE4oPU5/ZH17cHqP87juu6vGlxgh0Qn4B4YcZzGhG9dFAAAAAElFTkSu"""
    """QmCC"""
)


ICON_CLEAR = (
    """iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAAAXNSR0IArs4c6QAAAjRJREFUSEuNVotx"""
    """wjAMleC6R9gkTEIzBiF2yAcYo3QS2KSeozRxT/4kcuIAuYMzlqPPe3oyCOxBANDm97iiNYJ2+6GFvzt7"""
    """zRnJE1ta98+eMPQkIW70Xvle1ZyOiLgDgORZRnrmaKhSQd9XUhZXnrZZ1/XpE1b4FWQfSTdeQQg09n0m"""
    """5MEEQf9C3Z5vAJB2uttWQtyXoOM8jdzY3bK8JOsP+EHQShT5xrNpmKubsyFAFjlOqYxAO6FpPFG3F03B"""
    """rB/ryULU8gBLNL8ONQQQOYJGDlEYOd6CYcvGGmFWgckpApHHl7ihM38PzKpqr2gtDdbaNIQs8u0A6cxP"""
    """SDLjYMy/bi83AJ0CgOoeuCWLdY60d5Vinw0qdFBTEsJzwLooGsB2hqYOI20o80FIQYOS1CkTWkIOYiQb"""
    """csK+LsszD0JG63z6cIg8yZ6ounEkO4MfAmR3VRDmBAulpbpf3BInXLGUF1VAGYogQKRNfeWhc1QucAKA"""
    """qntAEGSupyckj9yMJBMsIyeYAGglRb55k+S4DmgArhB3gjlyQUrCShb7jCv/TSW7/L1IKEW35ndDXGjh"""
    """RLCjIiKQsI+sGN2XM0VUvcDlMIu4QOLT6MVNZxLlXTQZ1wiQ6l5n9sLwYliaP/PBVzZNusY1ifJuRwgX"""
    """WuzCmeG0NGUn+5Rkebh66ozVdMa6T2FFVybNGU/sADnX3ugxRE3pXn+X4nDkAhwO89sq7m2pJAfjrBEY"""
    """RGGRr29ej+/i/xAX7B9y8mgsRjNIxAAAAABJRU5ErkJggg=="""
)

ICON_DEL = (
    """iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAAAXNSR0IArs4c6QAAAiJJREFUSEuNVdFh"""
    """ozAMlWCQdhQySZsxDIkJpuAxLpvQUTpIG9/ZFljGojl/JIAlPz1JT0bYFgKCA4cA9JO2xKdgCEB/R8Z+"""
    """+1cjM9n3fyf1APCaH+K+EHC4durOYgyYfKEcQGQzfNh3QPcnvsXF7f3z9wNOw1V9ysS95+bFKROAtQs8"""
    """oPmhQzhAP9qmrmABhE/dqtMeeE11BBBiCwwmGwLXnSI7qhPZm2mm/RZTbnKojYGZ7AIOmgjm/cgHhRLR"""
    """9mEXEKuVcThhMNOtqqp+V6PijCdNs9rf9UWdQ5MlB8p7kZasJ2IrC0k13s8B6Aulk5U0i9J8WAfoQHc+"""
    """r2wVVcxT9x8A8YRguBZWOPQo778DMDpS55h59u0IWqmTBzCzXUIQbXz3a5ys85FF5qmSmdA8TgRw0Hct"""
    """rmbcObVvnsZko3Z15erx0Qkp4t9SEKQPknDSBAfIuihSzQBIUOM0E/3oPIRG4ALM/fLOZAzk6Lhzi7LC"""
    """c5s0ZbNZdMxgPxJKlrzIlKJMB2wUSjoIBWQiijYAfae2RpBAQ3L20o+GeYcM43TDqnrRnTp7+xt7j7wR"""
    """eJGTVtYUMfWskaTo8purmEW8iwLLfLIW94FEX1ZvLnO5TXdFDl3kVepcA4Bnza9DYcCtEP04N3WFdPm0"""
    """pzQQBQBtbFPXbpHvA34RiM9FUMWo0Na+1g9oEOHNbRfQ0Yjbvn958vqi7vtZXt7JT896XhFu8RczqWgy"""
    """PKwKugAAAABJRU5ErkJggg=="""
)


ICON_ADD = (
    """iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAAAXNSR0IArs4c6QAAAYZJREFUSEuVVdt1"""
    """wjAMldNBWKUdhTEC1IVAyBpsQjdLjx+JJesqTvPB4ThGV/ch4RwRzfGDdjz5Ir/PC4A64Qg+qka8JSuo"""
    """5jh+fqkB/kcpN6cqr027Uk/27GiWqgFgiwEnmwC4SBYDRzSM0ztc9af+SwtWKchIrdJisPw6ADym2Is/"""
    """96Z3RZvUefsiM3YYn/En/tIHBfUD0rUjReVXw7iHgTRce2A1FiXKDLJE2K46LMYUxBRVpgQG4eg7AoCZ"""
    """UGFZPNhIDtdaAshqPNa1FYLD8JzeNM+fjrCP8XIwWLiXDxz9hghvAtzu49V13Q9gW45iPY6Sv8/08pfT"""
    """kU8W68PWdHljS2RtDEfGJNcWpgLrHPBBU9MpU5MY7DZZxlTsHLZyhAe4gWTxElMuUVkV+Y5Up7Sfm2YM"""
    """gCxgDvAuwg2FiiJsrclsmmxsATjL6PB2f1xd93Hw5/4I1zU3hUlXb9hNwNZKr7esngO1ci28Rj4RAzy9"""
    """Mi0Yrt6g5e+WDRroSLHhYPsY/AFsVMwga5Ld0wAAAABJRU5ErkJggg=="""
)


ICON_COPY = (
    """iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAAAXNSR0IArs4c6QAAAjdJREFUSEuFVuuZ"""
    """qjAQPYOFXLcTrGTXMiIIImLK0FuJdrI2smRvJgEmJLn6w4+Qx8x5zATC/CMQDAwB/m+ZEk9ymp8Tq+R7"""
    """mgZdP5yoKD4BbOc9drddsA7gxq9/s11TqXsyE35J4ADdRX+BcEtnJKMYu2Wd9T4ZxB/G+Z2v+gGD8mfE"""
    """rmvU838UTdm2+rrdjPQNg1dTq480CpuODTBoprKp1ExIwOPMdYiRaaXiT1OrvdRSKpMNkOGdM3onbLDX"""
    """a2DJDRB4jVaneae9cU7oTEvRRRuQQVMdVp4Buv5yomLzCYOtXcO2SrhLxAzc5SjiAB6B4ODM7jK3hHOW"""
    """80KTwSbB6w32Ta3uIkCMwLnLlD8j7bqjeq6ZSdVkq/V2M+Lb1klTqY+pDiKKWJucu96U8+LKA7kAiYMW"""
    """+3pkkX3SRrb7+kEby5zVNNZA8BBok+8H0YysK4EgoYFEFhEeQlpGhPNwnQvXIdD6gRElAftjre5Tq3CZ"""
    """JOwr+szUfbkTeywpDb4A3HxxPZuD2qW1yfEeouEABmhq5TRw1jIliGxBWXtxb5kKsK0O5DLNNZBQhkjk"""
    """qcGsz0hRlPfOFITQD1fvIo9gmQqTjLus7EX5UJGLHPK4T04IIoq8o6SwkiSmiDXwdRAwKBzS+YsIcH3F"""
    """uit96JJc2+tyU+ABgjNLrr97kfkqTdeYvD7nBsdE+DP5KnUBMnQG7hpRzh8AuXbt3r/MOP5tj9VpvvQn"""
    """98lqZLmjwPLTRk7Lygvx/gLRN1wxA1X+ogAAAABJRU5ErkJggg=="""
)


ICON_COPY_PLAIN = (
    """iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAC2ElEQVR4AZxViXHbQAxcyI3YndiVxC5D"""
    """1kOReliGk0roVGL1kZEuiwUonuR8EwwPwAE4AMTymeF/yD4fmkyTBhiqAsYt/o1KhFkIckOaqJfII2fJ"""
    """ArlptvtNt+s/uMp21xfKARGOK1J8WDym2+09tnT7/qPrDs/yFPEswE3X9c8zmzU033ONHR0xam5ksSq3"""
    """PKWcv8vlrOCeGd88F0TGLZJm+OLaqeBpvZyb1mr+4rYpaQF7cdNlNavFZr1cKP70Aw9yzNBIYhyRduXR"""
    """Rbuav7tU0uIaGBbymisCkIjAtp3zjoWHpgBSgKwgi1ijlVccocJLpoldV6wDGXt9WY6IQZxlS/EVSQFe"""
    """XyhL64BvXT8E+PvQ6RuiqzxE4XkoMk81Ip9ls+TMWYUBODt46tp305JbDJ77eDs/z0P8hB1IVyPSuUy6"""
    """cfAWCfYI+vJVYDKB5MobYpI/XVFAmcWu55vFUtzkSauEGPyWkmGkmVxiaZJu2O76gTMWBpNMDByT3SF8"""
    """wmY/6LR6FEMUA2baihlE0oGTY5C67GKf3wOZYXw8Q3NuzvJsjMgNV6ugdQw0+1fjrDXz9YjB6hYbPhyX"""
    """80p/2f2mwMVPJVuhhvG+axOcMmmK2h0YeMz4vGUQ5z7w48U5H+LZz7kTG9r6QUcqlsdkqfXAgGYlJLBU"""
    """ddUY1B1JN2jmUyJZ9QR2u8PAF3MYb7Ye0SMPPCJJ74Fmvcj5V3NfxMwzbZ4YRfEcvmSoCuQTUmQPJr2M"""
    """zYSt4mwod5OGm+goIL/pKxrfckNNpd5U+mQPrdn2Y+fvI6RRwP3n8k1n7/DGOY7ADrJdmN30B3TErcuX"""
    """7s6geDsjcsHyawq+WCd7p+MFBboTr0k/wTSIJHKMMoBu4nqu/mjAsZRTu17Pp6+pzgFo+bNY0cEP2RNX"""
    """AJtg0g0WlpiYydQ0/ke7fAQfmtVyk80pNEYk1dSR1L8x84AS8cX1X9SX2fATAAD//4U3kXMAAAAGSURB"""
    """VAMAz6lfSAvBe+gAAAAASUVORK5CYII="""
)


ICON_SELECT_ALL = (
    """iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAAAXNSR0IArs4c6QAAAiZJREFUSEuNVe11"""
    """gzAMPJlBmm4Ck6QZwyF8tsAYTSdJNmkG6cONhQk2tkn4kUeELEun04lgPQRA8Y8xGgNBGZP56Dna/nZE"""
    """QLsGnlgg7/bH2eULYUlodUH92dVCiD0It+ooM11I2w8XKKQgXMujzHSghm0qBdG1zGWmy2O/ETsQmjKX"""
    """Z05dWRe0Xf8B0PeUkjqX+fGg35rPriYhKgDnMpe+7SQPOhD7kagMJge+5H7+AZHOgBTSvxFZU8hrGLoY"""
    """mlO61TDskhG/AG7lSb47FZhM3+YsYxeEu7DgruMIId4Krkzb5zr4z8yWV/pu+9ukmGCZiEggO6NNRgWY"""
    """6J0NBHN6oGPMrFgq83ntzMrq4rbvLzpzzTi3yV3PIBW5JC7tKVwBHwLar4Hvn+LAQHQP2bRMR7fJmxMb"""
    """xmMdx++BHdR7txsYgC7CjbhYxJVhdSbGonnQNHad3RyDb4Am3BtPBJfKdJNNL1lCFhbNTT5JWkvq84Yv"""
    """BbXdwCNQOk1mzfmqSSSmyc+nI9aNtSI4cr0J+StTZQ/PDKPHxBg1vcOB0bZNnlxrPdfY8QTGIdoEj8li"""
    """x1mzSFEKgLU8hogtZMtunQKV7ZAmibrwIvKlYvgAMC0cs730q7vRjplmVDPw9kqJcC1MoIff/TuBDsVJ"""
    """nhe5vstrNfS7ZCS9GvdQuPFeYAlhdr220YTYq3H8qYq8nokblWtH063hWsMyD51td/iwSZoYUV6WEOAf"""
    """jjY7LKuzEksAAAAASUVORK5CYII="""
)


ICON_UNDO = (
    """iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAAAXNSR0IArs4c6QAAAjpJREFUSEuNVv19"""
    """wiAQPdBB4iZmk3YMYg0h/jRjtJ1EN5E92kB7BMIRiJp/QHLc53svMsCHAYB1u8KevvQ2ryz+Gi7zU3a1"""
    """Fj1NKFqlXpIAz6vxBZLsrNszYGDTJjyuYLrQDkO1sfCJgdtG1MvOvNK8vEX+FjrfAnyCgT0AfLUH8V4c"""
    """1pMozFfohhxsZcgcnTPQ7UHsXplryYZUMLmfnVuXedn5q73BpGkF1DkD0EefObWhQHiOuhDA13a6DFfr"""
    """MkdSML/4Pa1/OtL+6DZyUL0QOlBpmcR8NQZIvU2YwqcQbDLVxpjv7njo0jngTcJe1yID13/UVNiiXw61"""
    """Eo2OCIhNQVu8yn/NG+Nc4htjjJKLIBnRaBDMbORQu/IfDFYOl2pj2B0s6NHad3VsbgGSSQVhmPLiK7FQ"""
    """Ya8xiGqEplBeQrI7nTvOuAQON9mIemppaNH8wyLrXUdiJbYCYLr9ELtZCzLAM3BVjHDHhORB7Kh2pqpF"""
    """Lvsg0kmFYzLlctozb3tPubOoIGr2EgtOzxK2zxah4tOw3zC4zi0K9msyW1bWXIpDYHU6d4imqFt+Bhk4"""
    """6AGV4uw8Srfsh/2WA5IURgt1fxS3dAYFR2uBacU4WP5j3tjGZQ52NEq2H138NiQzSD8mcQqTy/4yXMHa"""
    """vWO1I3TKajuTLG/jDI1sHsRW9ecOM3VH0be2xnybLf9SQgRtmjgWupIp5Rpo15hcmBn9BzHJdYE4DyGb"""
    """QSzlh5tBDtO1GdDoBZgutXmR7R8TrUAvyr2/1AAAAABJRU5ErkJggg=="""
)


ICON_REDO = (
    """iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAAAXNSR0IArs4c6QAAAkJJREFUSEuVVe2B"""
    """oyAQncErRCu5tZJNykAjijlDGclVEjuJhZywgqggmN3zT3QY5uPNmxeE/YMAoKzRvpuft3YEBOX7OCHW"""
    """1yOnrYajjG5BbiUA+iv6RAqG2X1pL3ItcjwnCKIdeDqx25t46qsjwplTOgRdxiAKwD8aEAK03e0JCj8A"""
    """YBgJ5FuSH0J0BN1ir4VIEwm6i3SfZA/If0Br2WKzBEkQcl4UgyGLB9Hb2blBw9HXN5EmKt7JMt4oGro6"""
    """8k+eMCG/QYHG2jJoIZ6ucXs3ZRiTGlhZZDNxEDBYosmHX7sGCam9zDae8Z8DOQkcT4SeFTRfLM4ezFA0"""
    """JjjW+ktKySWQvq1o7y23g5SZg1J3wygFA7sUmWGjs5NreiZE+kvCSzuMEnPOaG/aXGXAH9YcHO4aQpzo"""
    """WpU0iynPart24q4UnJSSvK7KJlzA7Xotbmmi0AQHUENdFpmnVzEWtZ14aV6PBLKW0uFQFMyiiacd/sBK"""
    """msVmGbCo7YQhMLtQT6NiusSvfxokyacOvqpNTHG3zAjtTa8/fIxS5W1V9EvLbnXxSl2V9afgVdp24j6t"""
    """/WmC6cFKel64HChoEO+7BPbcsGKEl16WaUV4VdKH334oZLHBbgqtGbh7eNs1mJglG9Qo/0okPa9o7/L6"""
    """MKjDHF+LbErNd70LWiIIIfUhiwj0jNI8LmG+IAYdLA3ZpasBITV09KXnwS707EOxKehbiL77H9jOY+Sd"""
    """T/esW+8EFIwYXK0Pki0K6khLCFEAbBzpn3b6BQruMyxajTd+AAAAAElFTkSuQmCC"""
)
