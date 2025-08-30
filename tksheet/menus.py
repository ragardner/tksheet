import tkinter as tk
from tkinter import TclError

from .functions import get_menu_kwargs
from .other_classes import Selected


def menu_add_command(menu: tk.Menu, **kwargs) -> None:
    if "label" not in kwargs:
        return
    try:
        index = menu.index(kwargs["label"])
        menu.delete(index)
    except TclError:
        pass
    menu.add_command(**kwargs)


def build_table_rc_menu(MT, popup_menu: tk.Menu, selected: Selected) -> None:
    popup_menu.delete(0, "end")
    mnkwgs = get_menu_kwargs(MT.PAR.ops)
    if selected and (MT.cut_enabled or MT.paste_enabled or MT.delete_key_enabled or MT.rc_sort_cells_enabled):
        selections_readonly = MT.estimate_selections_readonly()
    else:
        selections_readonly = False
    if selected:
        datarn, datacn = MT.datarn(selected.row), MT.datacn(selected.column)
        if MT.table_edit_cell_enabled() and not MT.is_readonly(datarn, datacn):
            menu_add_command(
                popup_menu,
                label=MT.PAR.ops.edit_cell_label,
                command=lambda: MT.open_cell(event="rc"),
                image=MT.PAR.ops.edit_cell_image,
                compound=MT.PAR.ops.edit_cell_compound,
                **mnkwgs,
            )
    if (
        MT.cut_enabled
        and selected
        and not selections_readonly
        and any(x in MT.enabled_bindings_menu_entries for x in ("all", "cut", "edit_bindings", "edit"))
    ):
        menu_add_command(
            popup_menu,
            label=MT.PAR.ops.cut_label,
            accelerator=MT.PAR.ops.cut_accelerator,
            command=MT.ctrl_x,
            image=MT.PAR.ops.cut_image,
            compound=MT.PAR.ops.cut_compound,
            **mnkwgs,
        )
    if (
        MT.copy_enabled
        and selected
        and any(x in MT.enabled_bindings_menu_entries for x in ("all", "copy", "edit_bindings", "edit"))
    ):
        menu_add_command(
            popup_menu,
            label=MT.PAR.ops.copy_label,
            accelerator=MT.PAR.ops.copy_accelerator,
            command=MT.ctrl_c,
            image=MT.PAR.ops.copy_image,
            compound=MT.PAR.ops.copy_compound,
            **mnkwgs,
        )
        if not MT.single_cell_selected():
            menu_add_command(
                popup_menu,
                label=MT.PAR.ops.copy_plain_label,
                accelerator=MT.PAR.ops.copy_plain_accelerator,
                command=MT.ctrl_c_plain,
                image=MT.PAR.ops.copy_plain_image,
                compound=MT.PAR.ops.copy_plain_compound,
                **mnkwgs,
            )
    if (
        MT.paste_enabled
        and (not selections_readonly or MT.PAR.ops.paste_can_expand_x or MT.PAR.ops.paste_can_expand_y)
        and any(x in MT.enabled_bindings_menu_entries for x in ("all", "paste", "edit_bindings", "edit"))
    ):
        menu_add_command(
            popup_menu,
            label=MT.PAR.ops.paste_label,
            accelerator=MT.PAR.ops.paste_accelerator,
            command=MT.ctrl_v,
            image=MT.PAR.ops.paste_image,
            compound=MT.PAR.ops.paste_compound,
            **mnkwgs,
        )
    if (
        MT.delete_key_enabled
        and selected
        and not selections_readonly
        and any(x in MT.enabled_bindings_menu_entries for x in ("all", "delete", "edit_bindings", "edit"))
    ):
        menu_add_command(
            popup_menu,
            label=MT.PAR.ops.delete_label,
            accelerator=MT.PAR.ops.delete_accelerator,
            command=MT.delete_key,
            image=MT.PAR.ops.delete_image,
            compound=MT.PAR.ops.delete_compound,
            **mnkwgs,
        )
    if MT.rc_sort_cells_enabled and selected and not selections_readonly and not MT.single_cell_selected():
        menu_add_command(
            popup_menu,
            label=MT.PAR.ops.sort_cells_label,
            accelerator=MT.PAR.ops.sort_cells_accelerator,
            command=MT.sort_boxes,
            image=MT.PAR.ops.sort_cells_image,
            compound=MT.PAR.ops.sort_cells_compound,
            **mnkwgs,
        )
        menu_add_command(
            popup_menu,
            label=MT.PAR.ops.sort_cells_reverse_label,
            accelerator=MT.PAR.ops.sort_cells_reverse_accelerator,
            command=lambda: MT.sort_boxes(reverse=True),
            image=MT.PAR.ops.sort_cells_reverse_image,
            compound=MT.PAR.ops.sort_cells_reverse_compound,
            **mnkwgs,
        )
        menu_add_command(
            popup_menu,
            label=MT.PAR.ops.sort_cells_x_label,
            accelerator=MT.PAR.ops.sort_cells_x_accelerator,
            command=lambda: MT.sort_boxes(row_wise=True),
            image=MT.PAR.ops.sort_cells_x_image,
            compound=MT.PAR.ops.sort_cells_x_compound,
            **mnkwgs,
        )
        menu_add_command(
            popup_menu,
            label=MT.PAR.ops.sort_cells_x_reverse_label,
            accelerator=MT.PAR.ops.sort_cells_x_reverse_accelerator,
            command=lambda: MT.sort_boxes(reverse=True, row_wise=True),
            image=MT.PAR.ops.sort_cells_x_reverse_image,
            compound=MT.PAR.ops.sort_cells_x_reverse_compound,
            **mnkwgs,
        )
    if MT.undo_enabled and any(
        x in MT.enabled_bindings_menu_entries for x in ("all", "undo", "redo", "edit_bindings", "edit")
    ):
        if MT.undo_stack:
            menu_add_command(
                popup_menu,
                label=MT.PAR.ops.undo_label,
                accelerator=MT.PAR.ops.undo_accelerator,
                command=MT.undo,
                image=MT.PAR.ops.undo_image,
                compound=MT.PAR.ops.undo_compound,
                **mnkwgs,
            )
        if MT.redo_stack:
            menu_add_command(
                popup_menu,
                label=MT.PAR.ops.redo_label,
                accelerator=MT.PAR.ops.redo_accelerator,
                command=MT.redo,
                image=MT.PAR.ops.redo_image,
                compound=MT.PAR.ops.redo_compound,
                **mnkwgs,
            )
    for label, kws in MT.extra_table_rc_menu_funcs.items():
        menu_add_command(popup_menu, label=label, **{**mnkwgs, **kws})


def build_index_rc_menu(MT, popup_menu: tk.Menu, selected: Selected) -> None:
    popup_menu.delete(0, "end")
    mnkwgs = get_menu_kwargs(MT.PAR.ops)
    if selected and (MT.cut_enabled or MT.paste_enabled or MT.delete_key_enabled or MT.rc_sort_row_enabled):
        selections_readonly = MT.estimate_selections_readonly()
    else:
        selections_readonly = False
    if selected:
        datarn = MT.datarn(selected.row)
        if MT.index_edit_cell_enabled() and not MT.RI.is_readonly(datarn):
            menu_add_command(
                popup_menu,
                label=MT.PAR.ops.edit_index_label,
                command=lambda: MT.RI.open_cell(event="rc"),
                image=MT.PAR.ops.edit_index_image,
                compound=MT.PAR.ops.edit_index_compound,
                **mnkwgs,
            )
    if (
        MT.cut_enabled
        and selected
        and not selections_readonly
        and any(x in MT.enabled_bindings_menu_entries for x in ("all", "cut", "edit_bindings", "edit"))
    ):
        menu_add_command(
            popup_menu,
            label=MT.PAR.ops.cut_label,
            accelerator=MT.PAR.ops.cut_accelerator,
            command=MT.ctrl_x,
            image=MT.PAR.ops.cut_image,
            compound=MT.PAR.ops.cut_compound,
            **mnkwgs,
        )
    if (
        MT.copy_enabled
        and selected
        and any(x in MT.enabled_bindings_menu_entries for x in ("all", "copy", "edit_bindings", "edit"))
    ):
        menu_add_command(
            popup_menu,
            label=MT.PAR.ops.copy_label,
            accelerator=MT.PAR.ops.copy_accelerator,
            command=MT.ctrl_c,
            image=MT.PAR.ops.copy_image,
            compound=MT.PAR.ops.copy_compound,
            **mnkwgs,
        )
        if not MT.single_cell_selected():
            menu_add_command(
                popup_menu,
                label=MT.PAR.ops.copy_plain_label,
                accelerator=MT.PAR.ops.copy_plain_accelerator,
                command=MT.ctrl_c_plain,
                image=MT.PAR.ops.copy_plain_image,
                compound=MT.PAR.ops.copy_plain_compound,
                **mnkwgs,
            )
    if (
        MT.paste_enabled
        and (not selections_readonly or MT.PAR.ops.paste_can_expand_x or MT.PAR.ops.paste_can_expand_y)
        and any(x in MT.enabled_bindings_menu_entries for x in ("all", "paste", "edit_bindings", "edit"))
    ):
        menu_add_command(
            popup_menu,
            label=MT.PAR.ops.paste_label,
            accelerator=MT.PAR.ops.paste_accelerator,
            command=MT.ctrl_v,
            image=MT.PAR.ops.paste_image,
            compound=MT.PAR.ops.paste_compound,
            **mnkwgs,
        )
    if (
        MT.delete_key_enabled
        and selected
        and not selections_readonly
        and any(x in MT.enabled_bindings_menu_entries for x in ("all", "delete", "edit_bindings", "edit"))
    ):
        menu_add_command(
            popup_menu,
            label=MT.PAR.ops.clear_contents_label,
            accelerator=MT.PAR.ops.clear_contents_accelerator,
            command=MT.delete_key,
            image=MT.PAR.ops.clear_contents_image,
            compound=MT.PAR.ops.clear_contents_compound,
            **mnkwgs,
        )
    if MT.rc_delete_row_enabled and selected:
        menu_add_command(
            popup_menu,
            label=MT.PAR.ops.delete_rows_label,
            command=MT.delete_rows,
            image=MT.PAR.ops.delete_rows_image,
            compound=MT.PAR.ops.delete_rows_compound,
            **mnkwgs,
        )
    if MT.rc_insert_row_enabled:
        menu_add_command(
            popup_menu,
            label=MT.PAR.ops.insert_rows_above_label,
            command=lambda: MT.rc_add_rows("above"),
            image=MT.PAR.ops.insert_rows_above_image,
            compound=MT.PAR.ops.insert_rows_above_compound,
            **mnkwgs,
        )
        menu_add_command(
            popup_menu,
            label=MT.PAR.ops.insert_rows_below_label,
            command=lambda: MT.rc_add_rows("below"),
            image=MT.PAR.ops.insert_rows_below_image,
            compound=MT.PAR.ops.insert_rows_below_compound,
            **mnkwgs,
        )
    if MT.rc_sort_row_enabled and selected and not selections_readonly and not MT.single_cell_selected():
        menu_add_command(
            popup_menu,
            label=MT.PAR.ops.sort_row_label,
            accelerator=MT.PAR.ops.sort_row_accelerator,
            command=MT.RI._sort_rows,
            image=MT.PAR.ops.sort_row_image,
            compound=MT.PAR.ops.sort_row_compound,
            **mnkwgs,
        )
        menu_add_command(
            popup_menu,
            label=MT.PAR.ops.sort_row_reverse_label,
            accelerator=MT.PAR.ops.sort_row_reverse_accelerator,
            command=lambda: MT.RI._sort_rows(reverse=True),
            image=MT.PAR.ops.sort_row_reverse_image,
            compound=MT.PAR.ops.sort_row_reverse_compound,
            **mnkwgs,
        )
    if MT.rc_sort_columns_enabled and selected and not MT.single_cell_selected():
        menu_add_command(
            popup_menu,
            label=MT.PAR.ops.sort_columns_label,
            accelerator=MT.PAR.ops.sort_columns_accelerator,
            command=MT.RI._sort_columns_by_row,
            image=MT.PAR.ops.sort_columns_image,
            compound=MT.PAR.ops.sort_columns_compound,
            **mnkwgs,
        )
        menu_add_command(
            popup_menu,
            label=MT.PAR.ops.sort_columns_reverse_label,
            accelerator=MT.PAR.ops.sort_columns_reverse_accelerator,
            command=lambda: MT.RI._sort_columns_by_row(reverse=True),
            image=MT.PAR.ops.sort_columns_reverse_image,
            compound=MT.PAR.ops.sort_columns_reverse_compound,
            **mnkwgs,
        )
    if MT.undo_enabled and any(
        x in MT.enabled_bindings_menu_entries for x in ("all", "undo", "redo", "edit_bindings", "edit")
    ):
        if MT.undo_stack:
            menu_add_command(
                popup_menu,
                label=MT.PAR.ops.undo_label,
                accelerator=MT.PAR.ops.undo_accelerator,
                command=MT.undo,
                image=MT.PAR.ops.undo_image,
                compound=MT.PAR.ops.undo_compound,
                **mnkwgs,
            )
        if MT.redo_stack:
            menu_add_command(
                popup_menu,
                label=MT.PAR.ops.redo_label,
                accelerator=MT.PAR.ops.redo_accelerator,
                command=MT.redo,
                image=MT.PAR.ops.redo_image,
                compound=MT.PAR.ops.redo_compound,
                **mnkwgs,
            )
    for label, kws in MT.extra_index_rc_menu_funcs.items():
        menu_add_command(popup_menu, label=label, **{**mnkwgs, **kws})


def build_header_rc_menu(MT, popup_menu: tk.Menu, selected: Selected) -> None:
    popup_menu.delete(0, "end")
    mnkwgs = get_menu_kwargs(MT.PAR.ops)
    if selected and (MT.cut_enabled or MT.paste_enabled or MT.delete_key_enabled or MT.rc_sort_column_enabled):
        selections_readonly = MT.estimate_selections_readonly()
    else:
        selections_readonly = False
    if selected:
        datacn = MT.datacn(selected.column)
        if MT.header_edit_cell_enabled() and not MT.CH.is_readonly(datacn):
            menu_add_command(
                popup_menu,
                label=MT.PAR.ops.edit_header_label,
                command=lambda: MT.CH.open_cell(event="rc"),
                image=MT.PAR.ops.edit_header_image,
                compound=MT.PAR.ops.edit_header_compound,
                **mnkwgs,
            )
    if (
        MT.cut_enabled
        and selected
        and not selections_readonly
        and any(x in MT.enabled_bindings_menu_entries for x in ("all", "cut", "edit_bindings", "edit"))
    ):
        menu_add_command(
            popup_menu,
            label=MT.PAR.ops.cut_label,
            accelerator=MT.PAR.ops.cut_accelerator,
            command=MT.ctrl_x,
            image=MT.PAR.ops.cut_image,
            compound=MT.PAR.ops.cut_compound,
            **mnkwgs,
        )
    if (
        MT.copy_enabled
        and selected
        and any(x in MT.enabled_bindings_menu_entries for x in ("all", "copy", "edit_bindings", "edit"))
    ):
        menu_add_command(
            popup_menu,
            label=MT.PAR.ops.copy_label,
            accelerator=MT.PAR.ops.copy_accelerator,
            command=MT.ctrl_c,
            image=MT.PAR.ops.copy_image,
            compound=MT.PAR.ops.copy_compound,
            **mnkwgs,
        )
        if not MT.single_cell_selected():
            menu_add_command(
                popup_menu,
                label=MT.PAR.ops.copy_plain_label,
                accelerator=MT.PAR.ops.copy_plain_accelerator,
                command=MT.ctrl_c_plain,
                image=MT.PAR.ops.copy_plain_image,
                compound=MT.PAR.ops.copy_plain_compound,
                **mnkwgs,
            )
    if (
        MT.paste_enabled
        and (not selections_readonly or MT.PAR.ops.paste_can_expand_x or MT.PAR.ops.paste_can_expand_y)
        and any(x in MT.enabled_bindings_menu_entries for x in ("all", "paste", "edit_bindings", "edit"))
    ):
        menu_add_command(
            popup_menu,
            label=MT.PAR.ops.paste_label,
            accelerator=MT.PAR.ops.paste_accelerator,
            command=MT.ctrl_v,
            image=MT.PAR.ops.paste_image,
            compound=MT.PAR.ops.paste_compound,
            **mnkwgs,
        )
    if (
        MT.delete_key_enabled
        and selected
        and not selections_readonly
        and any(x in MT.enabled_bindings_menu_entries for x in ("all", "delete", "edit_bindings", "edit"))
    ):
        menu_add_command(
            popup_menu,
            label=MT.PAR.ops.clear_contents_label,
            accelerator=MT.PAR.ops.clear_contents_accelerator,
            command=MT.delete_key,
            image=MT.PAR.ops.clear_contents_image,
            compound=MT.PAR.ops.clear_contents_compound,
            **mnkwgs,
        )
    if MT.rc_delete_column_enabled and selected:
        menu_add_command(
            popup_menu,
            label=MT.PAR.ops.delete_columns_label,
            command=MT.delete_columns,
            image=MT.PAR.ops.delete_columns_image,
            compound=MT.PAR.ops.delete_columns_compound,
            **mnkwgs,
        )
    if MT.rc_insert_column_enabled:
        menu_add_command(
            popup_menu,
            label=MT.PAR.ops.insert_columns_left_label,
            command=lambda: MT.rc_add_columns("left"),
            image=MT.PAR.ops.insert_columns_left_image,
            compound=MT.PAR.ops.insert_columns_left_compound,
            **mnkwgs,
        )
        menu_add_command(
            popup_menu,
            label=MT.PAR.ops.insert_columns_right_label,
            command=lambda: MT.rc_add_columns("right"),
            image=MT.PAR.ops.insert_columns_right_image,
            compound=MT.PAR.ops.insert_columns_right_compound,
            **mnkwgs,
        )
    if MT.rc_sort_column_enabled and selected and not selections_readonly and not MT.single_cell_selected():
        menu_add_command(
            popup_menu,
            label=MT.PAR.ops.sort_column_label,
            accelerator=MT.PAR.ops.sort_column_accelerator,
            command=MT.CH._sort_columns,
            image=MT.PAR.ops.sort_column_image,
            compound=MT.PAR.ops.sort_column_compound,
            **mnkwgs,
        )
        menu_add_command(
            popup_menu,
            label=MT.PAR.ops.sort_column_reverse_label,
            accelerator=MT.PAR.ops.sort_column_reverse_accelerator,
            command=lambda: MT.CH._sort_columns(reverse=True),
            image=MT.PAR.ops.sort_column_reverse_image,
            compound=MT.PAR.ops.sort_column_reverse_compound,
            **mnkwgs,
        )
    if MT.rc_sort_rows_enabled and selected and not MT.single_cell_selected():
        menu_add_command(
            popup_menu,
            label=MT.PAR.ops.sort_rows_label,
            accelerator=MT.PAR.ops.sort_rows_accelerator,
            command=MT.CH._sort_rows_by_column,
            image=MT.PAR.ops.sort_rows_image,
            compound=MT.PAR.ops.sort_rows_compound,
            **mnkwgs,
        )
        menu_add_command(
            popup_menu,
            label=MT.PAR.ops.sort_rows_reverse_label,
            accelerator=MT.PAR.ops.sort_rows_reverse_accelerator,
            command=lambda: MT.CH._sort_rows_by_column(reverse=True),
            image=MT.PAR.ops.sort_rows_reverse_image,
            compound=MT.PAR.ops.sort_rows_reverse_compound,
            **mnkwgs,
        )
    if MT.undo_enabled and any(
        x in MT.enabled_bindings_menu_entries for x in ("all", "undo", "redo", "edit_bindings", "edit")
    ):
        if MT.undo_stack:
            menu_add_command(
                popup_menu,
                label=MT.PAR.ops.undo_label,
                accelerator=MT.PAR.ops.undo_accelerator,
                command=MT.undo,
                image=MT.PAR.ops.undo_image,
                compound=MT.PAR.ops.undo_compound,
                **mnkwgs,
            )
        if MT.redo_stack:
            menu_add_command(
                popup_menu,
                label=MT.PAR.ops.redo_label,
                accelerator=MT.PAR.ops.redo_accelerator,
                command=MT.redo,
                image=MT.PAR.ops.redo_image,
                compound=MT.PAR.ops.redo_compound,
                **mnkwgs,
            )
    for label, kws in MT.extra_header_rc_menu_funcs.items():
        menu_add_command(popup_menu, label=label, **{**mnkwgs, **kws})


def build_empty_rc_menu(MT, popup_menu: tk.Menu) -> None:
    popup_menu.delete(0, "end")
    mnkwgs = get_menu_kwargs(MT.PAR.ops)
    if (
        MT.paste_enabled
        and any(x in MT.enabled_bindings_menu_entries for x in ("all", "paste", "edit_bindings", "edit"))
        and (MT.PAR.ops.paste_can_expand_x or MT.PAR.ops.paste_can_expand_y)
    ):
        menu_add_command(
            popup_menu,
            label=MT.PAR.ops.paste_label,
            accelerator=MT.PAR.ops.paste_accelerator,
            command=MT.ctrl_v,
            image=MT.PAR.ops.paste_image,
            compound=MT.PAR.ops.paste_compound,
            **mnkwgs,
        )
    if MT.rc_insert_column_enabled:
        menu_add_command(
            popup_menu,
            label=MT.PAR.ops.insert_column_label,
            command=lambda: MT.rc_add_columns("left"),
            image=MT.PAR.ops.insert_column_image,
            compound=MT.PAR.ops.insert_column_compound,
            **mnkwgs,
        )
    if MT.rc_insert_row_enabled:
        menu_add_command(
            popup_menu,
            label=MT.PAR.ops.insert_row_label,
            command=lambda: MT.rc_add_rows("below"),
            image=MT.PAR.ops.insert_row_image,
            compound=MT.PAR.ops.insert_row_compound,
            **mnkwgs,
        )
    if MT.undo_enabled and any(
        x in MT.enabled_bindings_menu_entries for x in ("all", "undo", "redo", "edit_bindings", "edit")
    ):
        if MT.undo_stack:
            menu_add_command(
                popup_menu,
                label=MT.PAR.ops.undo_label,
                accelerator=MT.PAR.ops.undo_accelerator,
                command=MT.undo,
                image=MT.PAR.ops.undo_image,
                compound=MT.PAR.ops.undo_compound,
                **mnkwgs,
            )
        if MT.redo_stack:
            menu_add_command(
                popup_menu,
                label=MT.PAR.ops.redo_label,
                accelerator=MT.PAR.ops.redo_accelerator,
                command=MT.redo,
                image=MT.PAR.ops.redo_image,
                compound=MT.PAR.ops.redo_compound,
                **mnkwgs,
            )
    for label, kws in MT.extra_empty_space_rc_menu_funcs.items():
        menu_add_command(popup_menu, label=label, **{**mnkwgs, **kws})
