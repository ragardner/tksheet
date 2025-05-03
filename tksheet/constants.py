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

ICON_CUT = (
    """iVBORw0KGgoAAAANSUhEUgAAABQAAAAUCAYAAACNiR0NAAAAAXNSR0IArs4c6QAAArtJREFUOE91VTFu"""
    """GzEQnOEpQADDroJULvSTSA+Qq3Ru5NL+w3FFXh0glV1GjR8Q+wGSX5IUfoEMAwGiu42W5ElnC2EhSHvH"""
    """2Z3dmRXRHwIEoPtA+dHH3zwoL9mFcovleQr1R0TOnHNf7HfbdU8xhE1+VlL1bx+BH0rJ+Xfwi8Vi5uju"""
    """AZwWkBfV7lJEHt+ADhKQQO1lHEP4ndKSCQ1e/Nmoqp4JvIKIamFVD+Kk3bbnIYbNoReHaprYrBQ6adt2"""
    """GmJY91ywCGHmHB8I3vi6vjNGTdNcq+qtdnohQR6HzTXIGONKgQmAtfd+aoVZPHUlhDAj+UDypvb1ncWa"""
    """EK8VuFXViwPtTL5p4koVEwLrWmSaRqk9ZQLeS6KMRJkxkVJ8I9Rt2+3nEJpNJqqpMoClsno6VAk0V5lO"""
    """qfKewKmm2/wD6kcolt77KyMTm7iCGk1de18qG6gkU86XU34Rf0bnvqSg6i8lvzpSQC6hOt5RnCix9rX1"""
    """LI01URkKOFd4pOicVkTGjpzTOSlU1nUawLGQ3kv16JU+RxpA37N60LP/1HHklEwZ6Nr2qRqNfpKY7Pq2"""
    """ryw2cS61X+7NUhDMeta5wVAWMzp3D+WpOQeKvwA/pAGIn5osYoxzgD8AXYr3VwY67FYain2YbKpR9Qzi"""
    """lWA0/qr4bjnbdvsphLixqyL12LlqThsU+ukX0DLc5D4J5uPslCxsmniTsLu2u1gszM+9EmRMurlzBVT8"""
    """1XsXJQ0aIIpTjHEsgMl6tiAG3LyXcVW5UimX4utE3/imNpr+RtXoWaGvJGNyEtVDcbJt2/O8xjJijyu7"""
    """LeOqUqnyqvb1cr8c7IstiKqsL7MYwZdO20uRxWOv/X6J9rCm010vxyHkTdP7ZG8eW7B0TE6xBduEsNHe"""
    """DW+UOti0eZ2mk6c+WMbFw+//CA5uPXJUQRts8H9EAVc2ZV3PAAAAAABJRU5ErkJggg=="""
)


ICON_EDIT = (
    """iVBORw0KGgoAAAANSUhEUgAAABQAAAAUCAYAAACNiR0NAAAAAXNSR0IArs4c6QAAAkVJREFUOE9tVe2h"""
    """4jAMs0mY4zpKu0kZhORiBik3SdmE7vEa/LCdtCkcf/qRRrEkWyDoD+0CXK4IgLw/6npdq9+3zztKRSqg"""
    """3x/t5+F2Royhc85NANCXCu5rzolSWgwQSwVstyzXWpPe1IoRFMy7pxTM/EqIpz8IMDLAPYRwQdkp7ASJ"""
    """9SsBK/cVrMhhYF7Bcl4HSvSQrTe6TQw8xhC0tP33IVW7GEK0ykzCJf/kIREtiAwp0QSAYwjXBvBT90oa"""
    """AUII3dn5pyqMcAeGEQCWvObh/Sx6zoD4iNcwVPW+tStaRKnM+RmAu5zzkFJ6ENEICFLVAswdAi6rSEBq"""
    """ymeL7ArEKGDmZgWrqymlv6cTxne1y5rzRQ4yAg1ey1rAvHMTA/RrzsOtGCB2hRh77/zMQjuvCmZ7Fe7b"""
    """lSBg3k3I2AsVO922xBj7IsGSa2VNJUZZsW0WWpqqSyKjouYImJvVkAPNw4ztjS2t4Z2fEFhpamVFEgU7"""
    """+xmYVTMqmm38SkUbZaF59m5i/jQAIYZr77yfi5uNAcfKdD7qqCeiGdWAfQJkTTTz3s0srfGzXojMza3P"""
    """tjix6VINY7xW18o8bnp+aEaPYwpVYMHnkgEos0gjM0xtr5mbxYA1X4jowTXSaoJ8xZp1NRClJwJ2Moui"""
    """5QlPIzprWnHTQqBEUHWhCrZdt25RQHm9AEC35Z9NwCAZ17xrgkTyUY8xK94hIQTUFCJLCwBemF//Xpnv"""
    """iQRo79iN5X/Cqc3OfVIOb0v877l6+DsoaWnHfaTUL8BEfy3rFX1+AAAAAElFTkSuQmCC"""
)


ICON_DOWN = (
    """iVBORw0KGgoAAAANSUhEUgAAABQAAAAUCAYAAACNiR0NAAAAAXNSR0IArs4c6QAAAh5JREFUOE9lVdt1"""
    """KzEIhCiF2J3YncSFGIRcSJxK4k6uC9k93AWhx27yE1tCMAwzGBEAAAFU7cP05xcAoPahXmLEISAoKqD2"""
    """GzsBBbtBzweqGA8isx/605YOLEON3RceiT3XhCCwEPOliLzanYjkluID8X0nevozQ91Rxvfo2OFaA9aW"""
    """SDEMTyK+lSJfAPA9s7Gu61lE3v1sYGp1RrjdlUdRVXgS0836ZeZTi5Aib+cgKBkcBKCKu/FeOZNSbBTP"""
    """O9HNron45AOJUCcRETrKvwjjJKqWItEy3XKWnBLycRCWXVUlM+fDjKIZG08QbC2jVoTWbkrpoqq97Son"""
    """Bc42LEPblTWmjIhW0WNLGUPpsjm01VBNx05I1fWUzA5Hy3xjpktK6Xv/cCa+3hDz2VA3owzCtwLyeCio"""
    """umyI6fSZ0m+ld2+pVgQBTJvXWuZvaW/Zpkwbh0dHOknO9fxwOCpadt6+7nd6OoePYn7zhDYUS1pmIe9M"""
    """P2kyyrhwP1P6p4hvvtNZQjabI35SSr8A8KJoabhjtL/HGsre7JoRPxhAzVInAPT/NuVlXa4i5bX302HO"""
    """k5R7XEs6GQeWZbkWKa+hhGExT9kHVXm0r8Oa24wk54wfyMb5uqxX8a0zySS2TNXzWG5N3bv11R4S0wVU"""
    """39ug6kaZIXdn2P6cdmWYIhK6l4YHD6KYVrVLpkfPK2By8NT0pAfXm86/ANFArJ0eulfqfyvsKyd4CFrk"""
    """AAAAAElFTkSuQmCC"""
)


ICON_UP = (
    """iVBORw0KGgoAAAANSUhEUgAAABQAAAAUCAYAAACNiR0NAAAAAXNSR0IArs4c6QAAAg5JREFUOE99VO1h"""
    """qzAMlOp2j3STZJMwCFIlBok3KZvUezyIX+UvDOE9/wkE6XQ+nYTQDgJABLAfOxHSM0aEmF7yKVH7t5aK"""
    """gDW/YmwFziBydMrHmIt2x/7v8dqnSfW+rOssIqHyMp4WbmyNRcJqgBtvTB/L9SxGRR8AeAeMYV3WQVTm"""
    """lthdrcrQQaUKjaE9iBaw8iEihHVZBhGd99qdCZMjTHZTvjCDOyDMGOEaAWYAuCJgWNY/CfQVphck3z8B"""
    """MtPVOff9CzAT001Frd1+XRdxzv0AQGCizyxbY5Kyk1ydB1KX7T9VvTOxN21UDRA8Ew/EdLH4rUH/4pk7"""
    """/NplTI3JgExDZfB/DTfz7ppS2VeGRDQw89W590e+abbD3uqZsUkSk4AnpQ0QAfzINBDx5cO57+qsVLTL"""
    """KfCBiG/W3BNjI6hKBiQe+nHouVXxDz3JDHdThACTaIyAnmgcSkO+tlZsU/B8Pv2xWXmaLb2bJJ2mCDF6"""
    """JhpEpztAfBznrcR707m4r+G8+CBpiODH0YLBfHqpwlnpqpOohiOZnYZ5ixTbIHgqgMR82UAkZIn6PVcm"""
    """Js9dPdulJ9UYC6CIfCEit7mKEJZ1vYloyGqVttTFUVnV3WqV7YpZbHseL859XGN8pol5w7cwMvnkyaPl"""
    """GsOK2sz7yjrZr+7BsgvT9RATy4p9Onr9Jm778qV1J/R+E/8CGGoXMBOU6gcAAAAASUVORK5CYII="""
)


ICON_PASTE = (
    """iVBORw0KGgoAAAANSUhEUgAAABQAAAAUCAYAAACNiR0NAAAAAXNSR0IArs4c6QAAAmFJREFUOE9tVYt1"""
    """2zAMPFDqHJUnqT1J3UFCEOwgcSepMkm0R0QjAj8S7cR+zxJFCLwD7mAiAGo/AEjL1dZ275kn1GfIgRYE"""
    """iMTFFu3RcZO369vtxZITMcZXIrqqdhv1YChm7/2lhu4Xgn2fPsw8OeeujhwrdNb7/a1ElVAi+gXgDKIb"""
    """VN88+1vey+cW1AUxEf6K/FfQOfNVgt7vwoFDf6aIBCLi7tniPZ8I2iqyU7wCeAVwSynJMA7vUCx7nY4M"""
    """kyFY13QqTIjXtF5ijLNl3CkHCWEgx2tKJxFZYox2wG8QcmOeGiBG1XuefgzD+12NSQhGjFqgBAnkiFNN"""
    """WArWSaC1u4PM7KdxHHlD+y+KzFUEJaLVpiX8hurDGez5LCLzfmZuCnUIa7HTmk4SZfkqiR0yQo01qsIS"""
    """SjuqlltxnhF2EVUSRx23+k3jMFoDz6oqmxCCBRXKldu3CYlAqtUxteK1UApM4zAwQGe1pjCHapNCRSRk"""
    """fT00ZdspBzneYe7C6GkCH5uMYpSldNn8aV3OCddT8WrpMns/DcNo1KYMriI0fs4RmzcIOr+wv1iiKhuC"""
    """mI6qbKLIkgdG+WnVzuvmCPM6QFds9mTvL/0M2Smbf9f0URA+j6E2gjZkRfRkTZm9f3kYEp2wDeGQKTeE"""
    """1dJH+WodrMuDc9fAHLohVne/dHm9iMT5QTZ5kPRSbyOPwOyzwA8d1nlovhzH8R2qs6q+maq+TDfVPJWa"""
    """xB3RTwWua0oXs94x5PKcA4I1JkvEuqXQhurZh61QxW437/2fcso+D49pZvXJGJo8nj3YrdtfwWFK4BNM"""
    """FnEr6HTQYwAAAABJRU5ErkJggg=="""
)


ICON_CLEAR = (
    """iVBORw0KGgoAAAANSUhEUgAAABQAAAAUCAYAAACNiR0NAAAAAXNSR0IArs4c6QAAAktJREFUOE9tVOuB"""
    """o0AI/tAU4lZyppLVQoTMWEhMJclVkhTiLifDjFEv/vGBMN8DIOwuAqAANvf8SKRQC9lVwuWVsMYstF4i"""
    """0tSnuk0Zln1I3L6XEDNP2zOoBGKMHYDrLpjwHqs6QlLnAuD1q3q7iFwskBESYgyW+yDFzf4q7JzigVNW"""
    """JR/1B4qWmb/s3c5CQkd6haIvFD5IlX5PRRI8LzeOsVNNzFJuQjhmukroeTBNLFFTTnpS3YFcnSFFDKtU"""
    """PUsuGOPYEfSqQC/Mk9MlCA/Ngv61pW/GhRBexe1dwYJwNYTIKS+MRLg51fVTQRPz0Ls04Q6gmef5bAcZ"""
    """vfA20xGa3mELO5/CIk1VVR1RJYCaNo2CWkAfwnI2/Qx5HGOHnYZkGo6dQl1Y0zC7akXrquoqIsm0HyJy"""
    """Nk0/U5Yp2WY6EOGqapSHae0SjxnNNnf/JMy996aJoAj/IcwuK3AlQj8kl/2KY7xbjxHw+FX9S0SyaDgx"""
    """c9LUOyR0SlhbLiEcQ/SPoJ6HYbJvMqQxfFqxQfgsAzdVXZum3z8/83nR/eXsQpowAvqBjXIR1sZO0S/u"""
    """Tpqbdj90x8WxsuhIsbbcG6HPcZ6UbXLWP30yZz32YQf0XBCKSHuq6zs2Gu7Q5WVwnHH7x3rY9If18OCG"""
    """IrVHXT8XNx8E3MqkrCsi2V6s2syNn/oNovZnnr9CDNbsPreXEC4VVeLtul0CuVD6syxSz1kRrzsgb5uy"""
    """pX3cTq0XfRfyzLIE8+LIYRu3feydenjau7ovU4xJO6m0ZDr0H9VKXi8tUg0XAAAAAElFTkSuQmCC"""
)

ICON_DEL = (
    """iVBORw0KGgoAAAANSUhEUgAAABQAAAAUCAYAAACNiR0NAAAAAXNSR0IArs4c6QAAAmJJREFUOE9tVe2h"""
    """mzAMlDCdo2SS5m1CBkGynUGgk7xuErpHjN6T/BGT1j8wjp3zSXcSiIAAICAAkF/zOj8BmHnKb2jPGMKu"""
    """R99HPa/zC0cXBZCZr865FQAmPYDSrtlTSrcQwp//A/fQjSZAjPETQK4iEoyhIKDDnwB4BREgossJsFCz"""
    """yViWcDV4fQ0xioJ5772I3WTDh+AHHDil50XDf+0UnDrFGFZAuJaQp/L7nvN3GmUP9lcqwBgrEVRqTDQ5"""
    """5xgB5yLPXjNR2WeS7/nBCUF2AdyZ6EMT3ULWw8oSAeaFKP/eUNUL5aqObIxRETZe6FbvU/2acWKMqupM"""
    """TKiqZqv8E3KDjPGuam1EdGta9NlpgJWh+i7GmYg2TY0qxMy+so8hM6SFbzXXJeQcmwEizLTkkEPwHk3R"""
    """dFGeP5x7HCKhglrIAGeGqH6RHNg9xlV6QB/8MCA/UzLPOTc+RI7gmb1VTQfY5OryDjGUHJaQQwgeETk9"""
    """00ULaBzHhxyHMaxebQxbrNnRVlqhilIA1cQOe4YuA2oq8vnvJ25Ei4liPtSQayVkUXBmWkxkH7xVxfOZ"""
    """Lnqvc+6h1WM5/F7fVRSAbWG+ac7UE5VfFuUeVxSciRdUI2unGUe3Lgt96Ltz45rSM+TGYL7tRHl1qeZg"""
    """ZXgydu+pU+nkDbONqsx069te+5uprMYuOWztUU+8+dtEuUe1x8ZUKyUXSRulB37qrSLyV+1Uk3LqJSIw"""
    """DMMvAbimlD763ljrroFaexoGzj2wK+dTbdtiP+T4rZ6sW2Xuk5Mrm5gms7p1/hpG/gzUz0Fon4Jz4/gC"""
    """TY9qMPGIU0oAAAAASUVORK5CYII="""
)


ICON_ADD = (
    """iVBORw0KGgoAAAANSUhEUgAAABQAAAAUCAYAAACNiR0NAAAAAXNSR0IArs4c6QAAAZdJREFUOE99VNtx"""
    """AyEMXI1dyF0nbiWFGOZIISkl6cQuJB4yeoGAi/11BqFdaVcihB8BqPZfvgkgO/Dzt/EACERAtfCYka8a"""
    """QJVvDq8G2e9GGhoXWcmzznS5nO4jH34oCWPWXprVW4Gc843jjqP8KJRDrtCxrsCsP+Aij1IenDCntA89"""
    """NjbcZ+43Y3VyC5hfVXyW8uAH6Z52r2gRaUk4dc57wyUc5RCGKaddVA86Rg2c8JkEGmeopZQHA9yN4SzU"""
    """ZLde2ojmopCWjIqUuOQTD4QjBm5mzjndLpfr12AjDdgAPMUWLEBAfr1+P1R9hRp8mHPertfLtxi4qpnl"""
    """MWEzBZ+jGISW0KQW2yzjFRTjBNxDEUVKnp2vwc5sYDjHutKSMNjmbHp6yb4FDMUN2rpBFeUwhjnvbe6t"""
    """IbNENn2hM0Sg6itA0/qk9JI5hgsNfQ774HTk9FCBlh6eLJTujGaF/5cfqw9g68shrhQfYB2fRZSu4rRR"""
    """gofOtmLwbnDrNEODgw2Jk7VFazYYn4XJa3TfDL/IFMYlEJfjPxtQyRxJrJ10AAAAAElFTkSuQmCC"""
)


ICON_COPY = (
    """iVBORw0KGgoAAAANSUhEUgAAABQAAAAUCAYAAACNiR0NAAAAAXNSR0IArs4c6QAAAlNJREFUOE9lVeF5"""
    """qzAMlDDdAzZJJ+l7g0Sy3UHaTtJsQt4cDVaRZIPJ84/AZ/BxujspCIgAAoAgQMwT9EtAH+hjvdhKKd39"""
    """Dm1PUPyFuuw9ZrqEMH4AwBmwofSIAHcQSET8WZlAj2lH8ntedLMU+XIqUlkpb2fXrgPim94T8dyoKVc/"""
    """USvJOYuIpMgxip1uNTiUq4KAIpBy/qNFEdFsD2wdZSMzTyGERQGZOR5fQ0CUHbuxPPh2YruYznDTYhpf"""
    """wiLFAXfelaR+cD/aE+/w3KhaDTNNYXxZpJS04cUmNBNfwjh+gEgFrBx7QXd38b5RTET0iQYYxkWkHBqq"""
    """USkvpo7IVyNjxtQf07RaMeDwpgTpSjNq9sZdw8pQAdWoIilGrka19B2xa9LFmOKAyOu6zmaKAhbVMLmG"""
    """TDyF8azryZQWvLqZUoo4DLz+PBRQNRzNlMjOxgF1z3W10g7EU1do6SnFiGeG41JKSTFx1AiqDC9jZW3O"""
    """9y3TJb06awwboHVKfv8GkKlsBmi4mRrrsn2iRqmLyfOtl4y8PtYZ1Sviq/Yy64tE/KoymK5So9RF5WkW"""
    """gHaLxgNx4HV9zLWWs0DePR4l07DrmL1fW2gQIJnLA/8Y4H9iIzBfHVCd70puL3fTyo7H5LF5aGx2vK4R"""
    """1BSLzeo5rO1/Hjud7SnGiGFwDd1AHSdt/Hj+c87fPh/xJrL+s9zY1PEe6ZeWKyg3utJrg9uHbmsoZr6E"""
    """EHgDvJxGcp0exwS0jfu6Pv6mlG915vise9ZGiVyJpzpzG8kTO5s09W9EH/wCNJx/qLBSvIoAAAAASUVO"""
    """RK5CYII="""
)


ICON_SELECT_ALL = (
    """iVBORw0KGgoAAAANSUhEUgAAABQAAAAUCAYAAACNiR0NAAAAAXNSR0IArs4c6QAAAjNJREFUOE91VYt1"""
    """4zAMoyp3D3cTe5PrICYtepD0Jqk3ifZodOqDKNlyrvXLi6MPKRCAGEfkyFGm7MieXN90rBwTzjnK2TY4"""
    """R5S7IIRjaJ+aRYRHjEMIEbMsMiJYQ4jYxizj4P3EzB9IKyKjBo0Z8TVNw0Wq2x8iuhHlnZlnbPaDv1Om"""
    """mNJj1rDFoJqJckwpzThUN8U65X/577rKikO6hArE+yI8t9JZZHJE8UDMMpLLFIJGhLLw+Oq9ELlp4eWt"""
    """UIEv4YLmBmSFJ8qIo2xEnaRefhkW0DR4f1vYgFw4fN5faMmOnMtVgD65JbS1hi0XscoD+CFs8SJzXWsK"""
    """mhVMAKvAhqhQ1cQp+cAVYKeU3oOGvYh+xBoF7bHkdRHJauwDsSHsPYf3qt5+CT5zd/5r57nCIdyQHmkG"""
    """mFpxZ+4K5nTn4eKOYqv5LLtK0Ur+RcYnbP0uE6p48NTjqIAW4Wnwww08aNC9CXOgbCIctj05BYe+8X9w"""
    """aJOfJ4c/XOz+bnXuqQkttnDYnPiM/WgOEKDX+WeCWjXH1bsydI7grHNTl/iimu3HVHWUI1W9p/Q14562"""
    """G2cxT5Gny8sKOpIf/Gf6KiXHcrwwj96/3tFpirnRrmp7wrsf+9q+2joEIaIpPdKbJaxPWNfVvbyII7cv"""
    """vFj78gMOiY+UZvRERftyFIuJQ4ib6p3IjUT5fWH5ANcXDuF6a03GiAhPVNtXuaIsIwKApCG0/ZUYCHzl"""
    """61ma/zvxb82s/ifQN7JWXStNNSIQAAAAAElFTkSuQmCC"""
)


ICON_UNDO = (
    """iVBORw0KGgoAAAANSUhEUgAAABQAAAAUCAYAAACNiR0NAAAAAXNSR0IArs4c6QAAAhFJREFUOE9tVNth"""
    """qzAMlQJ3D5ikdJNkEKRIDNJskozCHsXxjfwA28U/wcEWOg8dhLQQAHx4tidbcVev41Q4iQDexxv5NO67"""
    """8K9tfXxZnAqfCJcREH0osq90LreRf/8UsDuqy5VoflQtnDSNvuywaTii9qCiTwCYnNtGkWXdCcmdY8QS"""
    """iyVkGAAXnNnuA0tUQjEAfBDPt0xQzWDNXalFxYfK8gTwdbG20olcpwVVI0wANIjSaBOwbM6tIvJqRckG"""
    """MeUDImaeur57Gie7aWo2SjOtnw+/iOh2ZqxABhMPXdf9AMIEHuyChBfhC8lwx+0vBLh6wJWIxtKze4fW"""
    """DHEqCjC932+5f9YfwyWLssj9gshGD9E87uKXZCbosVOAyfu3sBU9MbLdU1U7d3WbG0VltaLHpGRWP+5i"""
    """pqHr/v0gwLSZD1VWK7qLnR6YaOr6/umc+zahrNgxKQ27zDyAh0HUFI3HgpULC6nqFQGsy9tM9IhnmhmO"""
    """WwRMA3sMfaI+6WTTsagEyL+bG3U5UDSjGluo0ic0mAc2jhkxT33fPYPS8zzWxi7mMwrgj2hCACIaMmwP"""
    """MHSXy4RBYVg3524q+or9Jw5LydsUFLnfL3hho6HKLYTVbe4mqVhjmybUUiAmaGajyXs/hHT5AGbme6J+"""
    """T5u9w/ziNFGSB8q8C7DakC1SOsXXEf5Rjn2aY9JVFbNoKdvr4/AfC84KLpEV0gkAAAAASUVORK5CYII="""
)


ICON_REDO = (
    """iVBORw0KGgoAAAANSUhEUgAAABQAAAAUCAYAAACNiR0NAAAAAXNSR0IArs4c6QAAAflJREFUOE+FVO2V"""
    """4zAIhHMKyXYSd2IXYohIIfZVEl8ll0Ls5SKBPpxN3umP9CQYBhiEiJCWqu1xxSuND+3l4TEbVZ/6nLwd"""
    """pMK9s0x2/1vO78Us3xYAjwoAIreBaVo+ZfACiICgTqQ9W0BmPndd9xcAHkT0hYigXpYcMpUwE2nPVsif"""
    """pQghXBGREeAxMX2Zc4FLp2a1D5lp6VKxC+GaQDNTA0ytbQERAJ2rb8x0PnWnS+2F2xhzVoBzST8pI917"""
    """rYotgkiYn4ZD5VuTiOCluk5s37Y+BFlz8oekbyJ3RbiAwqqqf/AXPiJxK5Vp0wPFlCPDlZj6RM9NSlNu"""
    """crsDaExxJKKlLXcbVURmAByetisR9a18D00RkUhkIaKxLXTbawODITHLYE1k5wkgQQYAnBF0nIiXqsfK"""
    """jZkvp667K+LK09SX5FudZIoiMijAvO17LyGkAr8TVJwUommxflr/U01dCo6HECUSp0D1OzBfr2b4yvMg"""
    """fY9nd5lY9LGJQIQQwh0Qzvu2j5HlYQhqk9seHIYihj80xWb1FDsdI6z7tv2Ows1eIYRyNk8EVDXFlJRf"""
    """snDQOconl8nEnHxGYlqqcztwuZ7N94CKoK5iYrax0+8o3rT4WdwP+ZaIDveh2KipkYdZbum1v7x/8Mff"""
    """JuG++dB+fGW1F+ZSR/IfEnP6G73UrosAAAAASUVORK5CYII="""
)
