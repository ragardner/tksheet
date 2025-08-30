from __future__ import annotations

import copy
import csv
import io
import re
import tkinter as tk
from bisect import bisect_left
from collections import deque
from collections.abc import Callable, Generator, Hashable, Iterable, Iterator, Sequence
from difflib import SequenceMatcher
from itertools import chain, islice, repeat
from types import ModuleType
from typing import Any, Literal

from .colors import color_map
from .constants import align_value_error, symbols_set
from .formatters import to_bool
from .other_classes import DotDict, EventDataDict, Highlight, Loc, Span

ORD_A = ord("A")


def wrap_text(
    text: str,
    max_width: int,
    max_lines: int,
    char_width_fn: Callable,
    widths: dict[str, int],
    wrap: Literal["", "c", "w"] = "",
    start_line: int = 0,
) -> Generator[str]:
    total_lines = 0
    line_width = 0
    if wrap == "c":
        current_line = []
        for line in text.split("\n"):
            for c in line:
                try:
                    char_width = widths[c]
                except KeyError:
                    char_width = char_width_fn(c)
                # most of the time the char will not be a tab
                if c != "\t":
                    # adding char to line would result in wrap
                    if line_width + char_width >= max_width:
                        if total_lines >= start_line:
                            yield "".join(current_line)

                        total_lines += 1
                        if total_lines >= max_lines:
                            return
                        current_line = []
                        line_width = 0

                        if char_width <= max_width:
                            current_line.append(c)
                            line_width = char_width
                    # adding char to line is okay
                    else:
                        current_line.append(c)
                        line_width += char_width
                # if char happens to be a tab
                else:
                    if (remainder := line_width % char_width) > 0:
                        char_width = char_width - remainder

                    # adding tab to line would result in wrap
                    if line_width + char_width >= max_width:
                        if total_lines >= start_line:
                            yield "".join(current_line)

                        total_lines += 1
                        if total_lines >= max_lines:
                            return
                        current_line = []
                        line_width = 0

                        if widths["\t"] <= max_width:
                            current_line.append(c)
                            line_width = widths["\t"]
                        else:
                            try:
                                space_width = widths[" "]
                            except KeyError:
                                space_width = char_width_fn(" ")
                            if space_width <= max_width:
                                current_line.append(" ")
                                line_width = space_width
                    # adding tab to line is okay
                    else:
                        current_line.append(c)
                        line_width += char_width

            if total_lines >= start_line:
                yield "".join(current_line)

            total_lines += 1
            if total_lines >= max_lines:
                return
            current_line = []  # Reset for next line
            line_width = 0

    elif wrap == "w":
        try:
            space_width = widths[" "]
        except KeyError:
            space_width = char_width_fn(" ")
        current_line = []

        for line in text.split("\n"):
            for i, word in enumerate(line.split()):
                # if we're going to next word and
                # if a space fits on the end of the current line we add one
                if i and line_width + space_width < max_width:
                    current_line.append(" ")
                    line_width += space_width

                # check if word will fit
                word_width = 0
                word_char_widths = []
                for c in word:
                    try:
                        char_width = widths[c]
                    except KeyError:
                        char_width = char_width_fn(c)
                    word_char_widths.append(char_width)
                    word_width += char_width

                # we only wrap by character if the whole word alone wont fit max width
                # word won't fit at all we resort to char wrapping it
                if word_width >= max_width:
                    # yield current line before char wrapping word
                    if current_line:
                        if total_lines >= start_line:
                            yield "".join(current_line)

                        total_lines += 1
                        if total_lines >= max_lines:
                            return
                        current_line = []
                        line_width = 0

                    for c, w in zip(word, word_char_widths):
                        # adding char to line would result in wrap
                        if line_width + w >= max_width:
                            if total_lines >= start_line:
                                yield "".join(current_line)

                            total_lines += 1
                            if total_lines >= max_lines:
                                return
                            current_line = []
                            line_width = 0

                            if w <= max_width:
                                current_line.append(c)
                                line_width = w
                        # adding char to line is okay
                        else:
                            current_line.append(c)
                            line_width += w

                # word won't fit on current line but will fit on a newline
                elif line_width + word_width >= max_width:
                    if total_lines >= start_line:
                        yield "".join(current_line)

                    total_lines += 1
                    if total_lines >= max_lines:
                        return
                    current_line = [word]
                    line_width = word_width

                # word will fit we put it on the current line
                else:
                    current_line.append(word)
                    line_width += word_width

            if total_lines >= start_line:
                yield "".join(current_line)

            total_lines += 1
            if total_lines >= max_lines:
                return

            current_line = []  # Reset for next line
            line_width = 0

    else:
        for line in text.split("\n"):
            line_width = 0
            current_line = []
            for c in line:
                try:
                    char_width = widths[c]
                except KeyError:
                    char_width = char_width_fn(c)
                if c == "\t" and (remainder := line_width % char_width) > 0:
                    char_width = char_width - remainder

                line_width += char_width
                if line_width >= max_width:
                    break
                current_line.append(c)

            if total_lines >= start_line:
                yield "".join(current_line)

            # Count the line whether it's empty or not
            total_lines += 1
            if total_lines >= max_lines:
                return


def estimate_max_visible_cells(table: tk.Canvas) -> int:
    """
    This is used limit for readonly checks to avoid perf issues on large selections.
    Calculate the rough maximum number of cells that could fit in the visible portion of the table.
    This uses a sort of minimum cell size in pixels and ignores headers, scrollbars, borders, etc.
    """
    # Get current table dimensions in pixels
    widget_width = table.winfo_width()
    widget_height = table.winfo_height()

    # If widget isn't realized yet (e.g., before pack/grid), winfo returns 1; handle that
    if widget_width <= 1 or widget_height <= 1:
        return 10000  # Fallback to a reasonable default

    # Calculate max columns and rows that fit
    max_cols = widget_width // 30
    max_rows = widget_height // 26

    # Rough max cells
    return max_rows * max_cols


def get_csv_str_dialect(s: str, delimiters: str) -> csv.Dialect:
    if len(s) > 6000:
        try:
            _upto = next(
                match.start() + 1 for i, match in enumerate(re.finditer("\n", s), 1) if i == 300 or match.start() > 6000
            )
        except Exception:
            _upto = len(s)
    else:
        _upto = len(s)
    try:
        return csv.Sniffer().sniff(s[:_upto] if len(s) > 6000 else s, delimiters=delimiters)
    except Exception:
        return csv.excel_tab


def get_data_from_clipboard(
    widget: tk.Misc,
    delimiters: str,
    lineterminator: str = "\n",
) -> list[list[str]]:
    data = widget.clipboard_get()
    dialect = get_csv_str_dialect(data, delimiters=delimiters)
    if dialect.delimiter in data or lineterminator in data:
        return list(csv.reader(io.StringIO(data), dialect=dialect, skipinitialspace=True))
    return [[data]]


def widget_descendants(widget: tk.Misc) -> list[tk.Misc]:
    result = []
    queue = deque([widget])
    while queue:
        current_widget = queue.popleft()
        result.append(current_widget)
        queue.extend(current_widget.winfo_children())
    return result


def recursive_bind(widget: tk.Misc, event: str, callback: Callable) -> None:
    widget.bind(event, callback)
    for child in widget.winfo_children():
        recursive_bind(child, event, callback)


def recursive_unbind(widget: tk.Misc, event: str) -> None:
    widget.unbind(event)
    for child in widget.winfo_children():
        recursive_unbind(child, event)


def tksheet_type_error(kwarg: str, valid_types: list[str], not_type: Any) -> str:
    valid_types = ", ".join(f"{type_}" for type_ in valid_types)
    return f"Argument '{kwarg}' must be one of the following types: {valid_types}, not {type(not_type)}."


def new_tk_event(keysym: str) -> tk.Event:
    event = tk.Event()
    event.keysym = keysym
    return event


def event_has_char_key(event: Any) -> bool:
    return (
        event and hasattr(event, "char") and (event.char.isalpha() or event.char.isdigit() or event.char in symbols_set)
    )


def event_opens_dropdown_or_checkbox(event=None) -> bool:
    if event is None:
        return False
    elif event == "rc":
        return True
    return (
        (hasattr(event, "keysym") and event.keysym in {"Return", "F2", "BackSpace"})
        or (
            hasattr(event, "keycode") and event.keycode == "??" and hasattr(event, "num") and event.num == 1
        )  # mouseclick
    )


def dropdown_search_function(search_for: str, data: Iterable[Any]) -> None | int:
    search_for = search_for.lower()
    search_len = len(search_for)
    if not search_len:
        return next((i for i, v in enumerate(data) if not str(v)), None)

    matcher = SequenceMatcher(None, search_for, "", autojunk=False)

    match_rn = None
    match_st = float("inf")
    match_len_diff = float("inf")

    fallback_rn = None
    fallback_match_length = 0
    fallback_st = float("inf")

    for rn, value in enumerate(data):
        value = str(value).lower()
        if not value:
            continue
        st = value.find(search_for)
        if st != -1:
            len_diff = len(value) - search_len
            if st < match_st or (st == match_st and len_diff < match_len_diff):
                match_rn = rn
                match_st = st
                match_len_diff = len_diff

        elif match_rn is None:
            matcher.set_seq2(value)
            match = matcher.find_longest_match(0, search_len, 0, len(value))
            match_length = match.size
            start = match.b if match_length > 0 else -1
            if match_length > fallback_match_length or (match_length == fallback_match_length and start < fallback_st):
                fallback_rn = rn
                fallback_match_length = match_length
                fallback_st = start

    return match_rn if match_rn is not None else fallback_rn


def float_to_int(f: int | float) -> int | float:
    if f == float("inf"):
        return f
    return int(f)


def event_dict(
    name: str = None,
    sheet: Any = None,
    widget: tk.Canvas | None = None,
    boxes: None | dict | tuple = None,
    cells_table: None | dict = None,
    cells_header: None | dict = None,
    cells_index: None | dict = None,
    selected: None | tuple = None,
    data: Any = None,
    key: None | str = None,
    value: Any = None,
    loc: None | int | tuple[int] = None,
    row: None | int = None,
    column: None | int = None,
    resized_rows: None | dict = None,
    resized_columns: None | dict = None,
    # resized_index: None, dict] = None,
    # resized_header: None, dict] = None,
    being_selected: None | tuple = None,
    named_spans: None | dict = None,
    sheet_state: None | dict = None,
    treeview: None | dict = None,
    **kwargs,
) -> EventDataDict:
    return EventDataDict(
        eventname="" if name is None else name,
        sheetname="!sheet" if sheet is None else sheet,
        cells=DotDict(
            table=DotDict() if cells_table is None else cells_table,
            header=DotDict() if cells_header is None else cells_header,
            index=DotDict() if cells_index is None else cells_index,
        ),
        moved=DotDict(
            rows=DotDict(),
            columns=DotDict(),
        ),
        added=DotDict(
            rows=DotDict(),
            columns=DotDict(),
        ),
        deleted=DotDict(
            rows=DotDict(),
            columns=DotDict(),
            header=DotDict(),
            index=DotDict(),
            column_widths=DotDict(),
            row_heights=DotDict(),
        ),
        named_spans=DotDict() if named_spans is None else named_spans,
        options=DotDict(),
        selection_boxes={} if boxes is None else boxes,
        selected=() if selected is None else selected,
        being_selected=() if being_selected is None else being_selected,
        data={} if data is None else data,
        key="" if key is None else key,
        value=None if value is None else value,
        loc=() if loc is None else loc,
        row=row,
        column=column,
        resized=DotDict(
            rows=DotDict() if resized_rows is None else resized_rows,
            columns=DotDict() if resized_columns is None else resized_columns,
            # "header": DotDict() if resized_header is None else resized_header,
            # "index": DotDict() if resized_index is None else resized_index,
        ),
        widget=widget,
        sheet_state=DotDict() if sheet_state is None else sheet_state,
        treeview=DotDict(
            nodes={},
            renamed={},
            text={},
        )
        if treeview is None
        else treeview,
    )


def change_eventname(event_dict: EventDataDict, newname: str) -> EventDataDict:
    return EventDataDict({**event_dict, **{"eventname": newname}})


def stored_event_dict(d: DotDict) -> DotDict:
    return DotDict(name=d["eventname"], data=DotDict(kv for kv in d.items() if kv[0] != "widget"))


def len_to_idx(n: int) -> int:
    if n < 1:
        return 0
    else:
        return n - 1


def b_index(sorted_seq: Sequence[int], num_to_index: int) -> int:
    """
    Designed to be a faster way of finding the index of an int
    in a sorted list of ints than list.index()
    """
    if (idx := bisect_left(sorted_seq, num_to_index)) == len(sorted_seq) or sorted_seq[idx] != num_to_index:
        raise ValueError(f"{num_to_index} is not in Sequence")
    else:
        return idx


def try_b_index(sorted_seq: Sequence[int], num_to_index: int) -> int | None:
    if (idx := bisect_left(sorted_seq, num_to_index)) == len(sorted_seq) or sorted_seq[idx] != num_to_index:
        return None
    else:
        return idx


def bisect_in(sorted_seq: Sequence[int], num: int) -> bool:
    """
    Faster than 'num in sorted_seq'
    """
    try:
        return sorted_seq[bisect_left(sorted_seq, num)] == num
    except Exception:
        return False


def push_n(num: int, sorted_seq: Sequence[int]) -> int:
    if num < sorted_seq[0]:
        return num
    else:
        hi = len(sorted_seq)
        lo = 0
        while lo < hi:
            mid = (lo + hi) // 2
            if sorted_seq[mid] < num + mid + 1:
                lo = mid + 1
            else:
                hi = mid
        return num + lo


def get_menu_kwargs(ops: DotDict[str, Any]) -> DotDict[str, Any]:
    return DotDict(
        {
            "font": ops.popup_menu_font,
            "foreground": ops.popup_menu_fg,
            "background": ops.popup_menu_bg,
            "activebackground": ops.popup_menu_highlight_bg,
            "activeforeground": ops.popup_menu_highlight_fg,
        }
    )


def get_bg_fg(ops: DotDict[str, Any]) -> dict[str, str]:
    return {
        "bg": ops.table_editor_bg,
        "fg": ops.table_editor_fg,
        "select_bg": ops.table_editor_select_bg,
        "select_fg": ops.table_editor_select_fg,
    }


def get_dropdown_kwargs(
    values: list[Any] | None = None,
    set_value: Any = None,
    state: str = "normal",
    redraw: bool = True,
    selection_function: Callable | None = None,
    modified_function: Callable | None = None,
    search_function: Callable = dropdown_search_function,
    validate_input: bool = True,
    text: None | str = None,
    edit_data: bool = True,
    default_value: Any = None,
) -> dict:
    return {
        "values": [] if values is None else values,
        "set_value": set_value,
        "state": state,
        "redraw": redraw,
        "selection_function": selection_function,
        "modified_function": modified_function,
        "search_function": search_function,
        "validate_input": validate_input,
        "text": text,
        "edit_data": edit_data,
        "default_value": default_value,
    }


def get_dropdown_dict(**kwargs) -> dict:
    return {
        "values": kwargs["values"],
        "select_function": kwargs["selection_function"],
        "modified_function": kwargs["modified_function"],
        "search_function": kwargs["search_function"],
        "validate_input": kwargs["validate_input"],
        "text": kwargs["text"],
        "state": kwargs["state"],
        "default_value": kwargs["default_value"],
    }


def get_checkbox_kwargs(
    checked: bool = False,
    state: str = "normal",
    redraw: bool = True,
    check_function: Callable | None = None,
    text: str = "",
) -> dict:
    return {
        "checked": checked,
        "state": state,
        "redraw": redraw,
        "check_function": check_function,
        "text": text,
    }


def get_checkbox_dict(**kwargs) -> dict:
    return {
        "check_function": kwargs["check_function"],
        "state": kwargs["state"],
        "text": kwargs["text"],
    }


def is_iterable(o: Any) -> bool:
    if isinstance(o, str):
        return False
    try:
        iter(o)
        return True
    except Exception:
        return False


def int_x_iter(i: Iterator[int] | int) -> Iterator[int]:
    if isinstance(i, int):
        return (i,)
    return i


def int_x_tuple(i: Iterator[int] | int) -> tuple[int]:
    if isinstance(i, int):
        return (i,)
    return tuple(i)


def unpack(t: tuple[Any] | tuple[Iterator[Any]]) -> tuple[Any]:
    if not len(t):
        return t
    if is_iterable(t[0]) and len(t) == 1:
        return t[0]
    return t


def is_type_int(o: Any) -> bool:
    return isinstance(o, int) and not isinstance(o, bool)


def force_bool(o: Any) -> bool:
    try:
        return to_bool(o)
    except Exception:
        return False


def alpha2idx(a: str) -> int | None:
    try:
        n = 0
        for c in a.upper():
            n = n * 26 + ord(c) - ORD_A + 1
        return n - 1
    except Exception:
        return None


def alpha2num(a: str) -> int | None:
    n = alpha2idx(a)
    return n if n is None else n + 1


def num2alpha(n: int) -> str | None:
    try:
        s = ""
        n += 1
        while n > 0:
            n, r = divmod(n - 1, 26)
            s = chr(65 + r) + s
        return s
    except Exception:
        return None


def idx_param_to_int(idx: str | int | None) -> int | None:
    if idx is None or isinstance(idx, int):
        return idx
    return alpha2idx(idx)


def get_n2a(n: int = 0, _type: Literal["letters", "numbers", "both"] | None = "numbers") -> str:
    if _type == "letters":
        return num2alpha(n)
    elif _type == "numbers":
        return f"{n + 1}"
    elif _type == "both":
        return f"{num2alpha(n)} {n + 1}"
    elif _type is None:
        return ""


def get_index_of_gap_in_sorted_integer_seq_forward(
    seq: list[int],
    start: int = 0,
) -> int | None:
    prevn = seq[start]
    for idx, n in enumerate(islice(seq, start + 1, None), start + 1):
        if n != prevn + 1:
            return idx
        prevn = n
    return None


def get_index_of_gap_in_sorted_integer_seq_reverse(
    seq: list[int],
    start: int = 0,
) -> int | None:
    prevn = seq[start]
    for idx, n in zip(range(start, -1, -1), reversed(seq[:start])):
        if n != prevn - 1:
            return idx
        prevn = n
    return None


def get_seq_without_gaps_at_index(
    seq: list[int],
    position: int,
    get_st_end: bool = False,
) -> tuple[int, int] | list[int]:
    start_idx = bisect_left(seq, position)
    forward_gap = get_index_of_gap_in_sorted_integer_seq_forward(seq, start_idx)
    reverse_gap = get_index_of_gap_in_sorted_integer_seq_reverse(seq, start_idx)
    if forward_gap is not None:
        seq = seq[:forward_gap]
    if reverse_gap is not None:
        seq = seq[reverse_gap:]
    if get_st_end:
        return seq[0], seq[-1]
    return seq


def consecutive_chunks(seq: list[int]) -> Generator[list[int]]:
    start = 0
    for index, value in enumerate(seq, 1):
        try:
            if seq[index] > value + 1:
                yield seq[start : (start := index)]
        except Exception:
            yield seq[start : len(seq)]


def consecutive_ranges(seq: Sequence[int]) -> Generator[tuple[int, int]]:
    seq_iter = iter(seq)
    try:
        start = next(seq_iter)
    except StopIteration:
        return
    prev = start
    for curr in seq_iter:
        if curr > prev + 1:
            yield start, prev + 1
            start = curr
        prev = curr
    yield start, prev + 1


def is_contiguous(iterable: Iterable[int]) -> bool:
    itr = iter(iterable)
    prev = next(itr)
    return all(i == (prev := prev + 1) for i in itr)


def box_is_single_cell(
    r1: int,
    c1: int,
    r2: int,
    c2: int,
) -> bool:
    return r2 - r1 == 1 and c2 - c1 == 1


def color_tup(color: str) -> tuple[int, int, int]:
    res = color if color.startswith("#") else color_map[color]
    return int(res[1:3], 16), int(res[3:5], 16), int(res[5:], 16)


def cell_down_within_box(
    r: int,
    c: int,
    r1: int,
    c1: int,
    r2: int,
    c2: int,
    numrows: int,
    numcols: int,
) -> tuple[int, int]:
    moved = False
    new_r = r
    new_c = c
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
    return new_r, new_c


def cell_right_within_box(
    r: int,
    c: int,
    r1: int,
    c1: int,
    r2: int,
    c2: int,
    numrows: int,
    numcols: int,
) -> tuple[int, int]:
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
    return new_r, new_c


def get_last(
    it: Iterator[Any],
) -> Any:
    if hasattr(it, "__reversed__"):
        try:
            return next(reversed(it))
        except Exception:
            return None
    else:
        try:
            return deque(it, maxlen=1)[0]
        except Exception:
            return None


def index_exists(seq: Sequence[Any], index: int) -> bool:
    try:
        seq[index]
        return True
    except Exception:
        return False


def add_to_displayed(displayed: list[int], to_add: Iterable[int]) -> list[int]:
    # assumes to_add is sorted
    for i in to_add:
        ins = bisect_left(displayed, i)
        displayed[ins:] = [i] + [e + 1 for e in islice(displayed, ins, None)]
    return displayed


def push_displayed(displayed: list[int], to_add: Iterable[int]) -> list[int]:
    # assumes to_add is sorted
    for i in to_add:
        ins = bisect_left(displayed, i)
        displayed[ins:] = [e + 1 for e in islice(displayed, ins, None)]
    return displayed


def move_elements_by_mapping(
    seq: list[Any],
    new_idxs: dict[int, int],
    old_idxs: dict[int, int] | None = None,
) -> list[Any]:
    # move elements of a list around
    # displacing other elements based on mapping
    # new_idxs = {old index: new index, ...}
    # old_idxs = {new index: old index, ...}
    if old_idxs is None:
        old_idxs = dict(zip(new_idxs.values(), new_idxs))
    remaining_values = (e for i, e in enumerate(seq) if i not in new_idxs)
    return [seq[old_idxs[i]] if i in old_idxs else next(remaining_values) for i in range(len(seq))]


def move_elements_by_mapping_gen(
    seq: list[Any],
    new_idxs: dict[int, int],
    old_idxs: dict[int, int] | None = None,
) -> Generator[Any]:
    if old_idxs is None:
        old_idxs = dict(zip(new_idxs.values(), new_idxs))
    remaining_values = (e for i, e in enumerate(seq) if i not in new_idxs)
    return (seq[old_idxs[i]] if i in old_idxs else next(remaining_values) for i in range(len(seq)))


def move_fast(seq: list[Any], new_idxs: dict[int, int], old_idxs: dict[int, int]) -> list[Any]:
    remaining_values = (e for i, e in enumerate(seq) if i not in new_idxs)
    return [seq[old_idxs[i]] if i in old_idxs else next(remaining_values) for i in range(len(seq))]


def move_elements_to(
    seq: list[Any],
    move_to: int,
    to_move: list[int],
) -> list[Any]:
    return move_elements_by_mapping(
        seq,
        *get_new_indexes(
            move_to,
            to_move,
            get_inverse=True,
        ),
    )


def remove_duplicates_outside_section(strings: list[str], section_start: int, section_size: int) -> list[str]:
    if section_start == 0 and section_size >= len(strings):
        return strings

    section_end = section_start + section_size
    section_set = set(strings[section_start:section_end])

    return [s for i, s in enumerate(strings) if (section_start <= i < section_end) or (s not in section_set)]


def any_editor_or_dropdown_open(MT) -> bool:
    return (
        MT.dropdown.open
        or MT.RI.dropdown.open
        or MT.CH.dropdown.open
        or MT.text_editor.open
        or MT.RI.text_editor.open
        or MT.CH.text_editor.open
    )


def get_new_indexes(
    move_to: int,
    to_move: Iterable[int],
    get_inverse: bool = False,
) -> tuple[dict[int, int]] | dict[int, int]:
    """
    move_to: A positive int, could possibly be the same as an element of to_move
    to_move: An iterable of ints, could be a dict, could be in any order
    returns {old idx: new idx, ...}
    """
    offset = sum(1 for i in to_move if i < move_to)
    correct_move_to = move_to - offset
    if not get_inverse:
        return {elem: correct_move_to + i for i, elem in enumerate(to_move)}
    else:
        new_idxs = {}
        old_idxs = {}
        for i, elem in enumerate(to_move):
            value = correct_move_to + i
            new_idxs[elem] = value
            old_idxs[value] = elem
        return new_idxs, old_idxs


def insert_items(
    seq: list[Any],
    to_insert: dict[int, Any],
    seq_len_func: Callable | None = None,
) -> list[Any]:
    """
    seq: list[Any]
    to_insert: keys are ints sorted, representing list indexes to insert items.
               Values are any, e.g. {0: 200, 1: 200}
    """
    if to_insert:
        if seq_len_func and next(reversed(to_insert)) >= len(seq) + len(to_insert):
            seq_len_func(next(reversed(to_insert)) - len(to_insert))
        for idx, v in to_insert.items():
            seq[idx:idx] = [v]
    return seq


def data_to_displayed_idxs(
    to_convert: list[int],
    displayed: list[int],
) -> list[int]:
    return [i for i, e in enumerate(displayed) if bisect_in(to_convert, e)]


def displayed_to_data_idxs(
    to_convert: list[int],
    displayed: list[int],
) -> list[int]:
    return [displayed[e] for e in to_convert]


def rounded_box_coords(
    x1: float,
    y1: float,
    x2: float,
    y2: float,
    radius: int = 5,
) -> tuple[float, ...]:
    # Handle case where rectangle is too small for rounding
    if y2 - y1 < 2 or x2 - x1 < 2:
        return (x1, y1, x2, y1, x2, y2, x1, y2, x1, y1)
    # Coordinates for a closed rectangle with rounded corners
    return (
        x1 + radius,
        y1,  # Top side start
        x2 - radius,
        y1,  # Top side end
        x2,
        y1,  # Top-right corner
        x2,
        y1 + radius,
        x2,
        y2 - radius,  # Right side
        x2,
        y2,  # Bottom-right corner
        x2 - radius,
        y2,
        x1 + radius,
        y2,  # Bottom side
        x1,
        y2,  # Bottom-left corner
        x1,
        y2 - radius,
        x1,
        y1 + radius,  # Left side
        x1,
        y1,  # Top-left corner
        x1 + radius,
        y1,  # Close the shape
    )


def diff_gen(seq: list[float]) -> Generator[int]:
    it = iter(seq)
    a = next(it)
    for b in it:
        yield int(b - a)
        a = b


def gen_coords(
    start_row: int,
    start_col: int,
    end_row: int,
    end_col: int,
    reverse: bool = False,
) -> Generator[tuple[int, int]]:
    if reverse:
        for r in reversed(range(start_row, end_row)):
            for c in reversed(range(start_col, end_col)):
                yield (r, c)
    else:
        for r in range(start_row, end_row):
            for c in range(start_col, end_col):
                yield (r, c)


def box_gen_coords(
    from_r: int,
    from_c: int,
    upto_r: int,
    upto_c: int,
    start_r: int,
    start_c: int,
    reverse: bool,
    all_rows_displayed: bool = True,
    all_cols_displayed: bool = True,
    displayed_cols: list[int] | None = None,
    displayed_rows: list[int] | None = None,
    no_wrap: bool = False,
) -> Generator[tuple[int, int]]:
    # Initialize empty lists if None
    if displayed_rows is None:
        displayed_rows = []
    if displayed_cols is None:
        displayed_cols = []

    # Adjust row indices based on displayed_rows
    if not all_rows_displayed:
        from_r = displayed_rows[from_r]
        upto_r = displayed_rows[upto_r - 1] + 1
        start_r = displayed_rows[start_r]
    # Adjust column indices based on displayed_cols (fixing original bug)
    if not all_cols_displayed:
        from_c = displayed_cols[from_c]
        upto_c = displayed_cols[upto_c - 1] + 1
        start_c = displayed_cols[start_c]

    if not reverse:
        # Forward direction
        # Part 1: From (start_r, start_c) to the end of the box
        for c in range(start_c, upto_c):
            yield (start_r, c)
        for r in range(start_r + 1, upto_r):
            for c in range(from_c, upto_c):
                yield (r, c)
        if not no_wrap:
            # Part 2: Wrap around from beginning to just before (start_r, start_c)
            for r in range(from_r, start_r):
                for c in range(from_c, upto_c):
                    yield (r, c)
            if start_c > from_c:  # Only if there are columns before start_c
                for c in range(from_c, start_c):
                    yield (start_r, c)
    else:
        # Reverse direction
        # Part 1: From (start_r, start_c) backwards to the start of the box
        for c in range(start_c, from_c - 1, -1):
            yield (start_r, c)
        for r in range(start_r - 1, from_r - 1, -1):
            for c in range(upto_c - 1, from_c - 1, -1):
                yield (r, c)
        if not no_wrap:
            # Part 2: Wrap around from end to just after (start_r, start_c)
            for r in range(upto_r - 1, start_r, -1):
                for c in range(upto_c - 1, from_c - 1, -1):
                    yield (r, c)
            if start_c < upto_c - 1:  # Only if there are columns after start_c
                for c in range(upto_c - 1, start_c, -1):
                    yield (start_r, c)


def next_cell(
    start_row: int,
    start_col: int,
    end_row: int,
    end_col: int,
    row: int,
    col: int,
    reverse: bool = False,
) -> tuple[int, int]:
    if reverse:
        col -= 1
        if col < start_col:
            col = end_col - 1
            row -= 1
            if row < start_row:
                row = end_row - 1
    else:
        col += 1
        if col == end_col:
            col = start_col
            row += 1
            if row == end_row:
                row = start_row
    return row, col


def is_last_cell(
    start_row: int,
    start_col: int,
    end_row: int,
    end_col: int,
    row: int,
    col: int,
    reverse: bool = False,
) -> bool:
    if reverse:
        return row == start_row and col == start_col
    return row == end_row - 1 and col == end_col - 1


def zip_fill_2nd_value(x: Iterator[Any], o: Any) -> Generator[Any, Any]:
    return zip(x, repeat(o))


def str_to_int(s: str) -> int | None:
    if s.startswith(("-", "+")):
        if s[1:].isdigit():
            return int(s)
        return None
    elif not s.startswith(("-", "+")):
        if s.isdigit():
            return int(s)
        return None


def gen_formatted(
    options: dict,
    formatter: Any = None,
) -> Generator[tuple[int, int]] | Generator[int]:
    if formatter is None:
        return (k for k, dct in options.items() if "format" in dct)
    return (k for k, dct in options.items() if "format" in dct and dct["format"]["formatter"] == formatter)


def options_with_key(
    options: dict,
    key: str,
) -> Generator[tuple[int, int]] | Generator[int]:
    return (k for k, dct in options.items() if key in dct)


def try_binding(
    binding: None | Callable,
    event: dict,
    new_name: None | str = None,
) -> bool:
    if binding:
        try:
            if new_name is None:
                binding(event)
            else:
                binding(change_eventname(event, new_name))
        except Exception:
            return False
    return True


def menu_item_exists(menu: tk.Menu, label: str) -> bool:
    try:
        menu.index(label)
        return True
    except Exception:
        return False


def span_dict(
    from_r: int | None = None,
    from_c: int | None = None,
    upto_r: int | None = None,
    upto_c: int | None = None,
    type_: str | None = None,
    name: str | None = None,
    kwargs: dict | None = None,
    table: bool = True,
    header: bool = False,
    index: bool = False,
    tdisp: bool = False,
    idisp: bool = True,
    hdisp: bool = True,
    transposed: bool = False,
    ndim: int | None = None,
    convert: Callable | None = None,
    undo: bool = False,
    emit_event: bool = False,
    widget: Any = None,
) -> Span:
    d: Span = Span(
        from_r=from_r,
        from_c=from_c,
        upto_r=upto_r,
        upto_c=upto_c,
        type_="" if type_ is None else type_,
        name="" if name is None else name,
        kwargs={} if kwargs is None else kwargs,
        table=table,
        index=index,
        header=header,
        tdisp=tdisp,
        idisp=idisp,
        hdisp=hdisp,
        transposed=transposed,
        ndim=ndim,
        convert=convert,
        undo=undo,
        emit_event=emit_event,
        widget=widget,
    )
    return d


def coords_to_span(
    widget: Any,
    from_r: int | None = None,
    from_c: int | None = None,
    upto_r: int | None = None,
    upto_c: int | None = None,
) -> Span:
    if not isinstance(from_r, int) and from_r is not None:
        from_r = None
    if not isinstance(from_c, int) and from_c is not None:
        from_c = None
    if not isinstance(upto_r, int) and upto_r is not None:
        upto_r = None
    if not isinstance(upto_c, int) and upto_c is not None:
        upto_c = None
    if from_r is None and from_c is None:
        from_r = 0
    return span_dict(
        from_r=from_r,
        from_c=from_c,
        upto_r=upto_r,
        upto_c=upto_c,
        widget=widget,
    )


PATTERN_ROW = re.compile(r"^(\d+)$")  # "1"
PATTERN_COL = re.compile(r"^([A-Z]+)$")  # "A"
PATTERN_CELL = re.compile(r"^([A-Z]+)(\d+)$")  # "A1"
PATTERN_RANGE = re.compile(r"^([A-Z]+)(\d+):([A-Z]+)(\d+)$")  # "A1:B2"
PATTERN_ROW_RANGE = re.compile(r"^(\d+):(\d+)$")  # "1:2"
PATTERN_ROW_START = re.compile(r"^(\d+):$")  # "2:"
PATTERN_ROW_END = re.compile(r"^:(\d+)$")  # ":2"
PATTERN_COL_RANGE = re.compile(r"^([A-Z]+):([A-Z]+)$")  # "A:B"
PATTERN_COL_START = re.compile(r"^([A-Z]+):$")  # "A:"
PATTERN_COL_END = re.compile(r"^:([A-Z]+)$")  # ":B"
PATTERN_CELL_START = re.compile(r"^([A-Z]+)(\d+):$")  # "A1:"
PATTERN_CELL_END = re.compile(r"^:([A-Z]+)(\d+)$")  # ":B1"
PATTERN_CELL_COL = re.compile(r"^([A-Z]+)(\d+):([A-Z]+)$")  # "A1:B"
PATTERN_CELL_ROW = re.compile(r"^([A-Z]+)(\d+):(\d+)$")  # "A1:2"
PATTERN_ALL = re.compile(r"^:$")  # ":"


def span_a2i(a: str) -> int | None:
    n = 0
    for c in a.upper():
        n = n * 26 + ord(c) - ORD_A + 1
    return n - 1


def span_a2n(a: str) -> int | None:
    n = 0
    for c in a.upper():
        n = n * 26 + ord(c) - ORD_A + 1
    return n


def key_to_span(
    key: (
        str
        | int
        | slice
        | Sequence[int | None, int | None]
        | Sequence[int | None, int | None, int | None, int | None]
        | Sequence[Sequence[int | None, int | None], Sequence[int | None, int | None]]
    ),
    spans: dict[str, Span],
    widget: Any = None,
) -> Span:
    """
    Convert various input types to a Span object representing a 2D range.

    Args:
        key: The input to convert (str, int, slice, sequence, or None).
        spans: A dictionary of named spans (e.g., {"<name>": Span(...)}).
        widget: Optional widget context for span creation.

    Returns:
        A Span object or an error message string if the key is invalid.
    """
    # Handle Span object directly
    if isinstance(key, Span):
        return key

    # Handle None as full span
    elif key is None:
        return coords_to_span(widget=widget, from_r=None, from_c=None, upto_r=None, upto_c=None)

    # Validate input type
    elif not isinstance(key, (str, int, slice, tuple, list)):
        return f"Key type must be either str, int, list, tuple or slice, not '{type(key).__name__}'."

    try:
        # Integer key: whole row
        if isinstance(key, int):
            return span_dict(
                from_r=key,
                from_c=None,
                upto_r=key + 1,
                upto_c=None,
                widget=widget,
            )

        # Slice key: row range
        elif isinstance(key, slice):
            start = 0 if key.start is None else key.start
            return span_dict(
                from_r=start,
                from_c=None,
                upto_r=key.stop,
                upto_c=None,
                widget=widget,
            )

        # Sequence key: various span formats
        elif isinstance(key, (tuple, list)):
            if len(key) == 2:
                if (isinstance(key[0], int) or key[0] is None) and (isinstance(key[1], int) or key[1] is None):
                    # Single cell or partial span: (row, col)
                    r_int = isinstance(key[0], int)
                    c_int = isinstance(key[1], int)
                    return span_dict(
                        from_r=key[0] if r_int else 0,
                        from_c=key[1] if c_int else 0,
                        upto_r=key[0] + 1 if r_int else None,
                        upto_c=key[1] + 1 if c_int else None,
                        widget=widget,
                    )

                elif isinstance(key[0], int) and isinstance(key[1], str):
                    # Single cell with column letter: (row 0, col A)
                    c_int = span_a2i(key[1])
                    return span_dict(
                        from_r=key[0],
                        from_c=c_int,
                        upto_r=key[0] + 1,
                        upto_c=c_int + 1,
                        widget=widget,
                    )

                else:
                    return f"'{key}' could not be converted to span."

            elif len(key) == 4:
                # Full span coordinates: (from_r, from_c, upto_r, upto_c)
                return coords_to_span(
                    widget=widget,
                    from_r=key[0],
                    from_c=key[1],
                    upto_r=key[2],
                    upto_c=key[3],
                )
            elif len(key) == 2 and all(isinstance(k, (tuple, list)) for k in key):
                # Start and end points: ((from_r, from_c), (upto_r, upto_c))
                return coords_to_span(
                    widget=widget,
                    from_r=key[0][0],
                    from_c=key[0][1],
                    upto_r=key[1][0],
                    upto_c=key[1][1],
                )

        # String key: parse various span formats
        elif isinstance(key, str):
            if not key:
                # Empty string treated as full span
                return span_dict(
                    from_r=0,
                    from_c=None,
                    upto_r=None,
                    upto_c=None,
                    widget=widget,
                )
            elif key.startswith("<") and key.endswith(">"):
                name = key[1:-1]
                return spans.get(name, f"'{name}' not in named spans.")

            key = key.upper()  # Case-insensitive parsing

            # Match string against precompiled patterns
            if m := PATTERN_ROW.match(key):
                return span_dict(
                    from_r=int(m[1]) - 1,
                    from_c=None,
                    upto_r=int(m[1]),
                    upto_c=None,
                    widget=widget,
                )
            elif m := PATTERN_COL.match(key):
                c_int = span_a2i(m[1])
                return span_dict(
                    from_r=None,
                    from_c=c_int,
                    upto_r=None,
                    upto_c=c_int + 1,
                    widget=widget,
                )
            elif m := PATTERN_CELL.match(key):
                c = span_a2i(m[1])
                r = int(m[2]) - 1
                return span_dict(
                    from_r=r,
                    from_c=c,
                    upto_r=r + 1,
                    upto_c=c + 1,
                    widget=widget,
                )
            elif m := PATTERN_RANGE.match(key):
                return span_dict(
                    from_r=int(m[2]) - 1,
                    from_c=span_a2i(m[1]),
                    upto_r=int(m[4]),
                    upto_c=span_a2n(m[3]),
                    widget=widget,
                )
            elif m := PATTERN_ROW_RANGE.match(key):
                return span_dict(
                    from_r=int(m[1]) - 1,
                    from_c=None,
                    upto_r=int(m[2]),
                    upto_c=None,
                    widget=widget,
                )
            elif m := PATTERN_ROW_START.match(key):
                return span_dict(
                    from_r=int(m[1]) - 1,
                    from_c=None,
                    upto_r=None,
                    upto_c=None,
                    widget=widget,
                )
            elif m := PATTERN_ROW_END.match(key):
                return span_dict(
                    from_r=0,
                    from_c=None,
                    upto_r=int(m[1]),
                    upto_c=None,
                    widget=widget,
                )
            elif m := PATTERN_COL_RANGE.match(key):
                return span_dict(
                    from_r=None,
                    from_c=span_a2i(m[1]),
                    upto_r=None,
                    upto_c=span_a2n(m[2]),
                    widget=widget,
                )
            elif m := PATTERN_COL_START.match(key):
                return span_dict(
                    from_r=None,
                    from_c=span_a2i(m[1]),
                    upto_r=None,
                    upto_c=None,
                    widget=widget,
                )
            elif m := PATTERN_COL_END.match(key):
                return span_dict(
                    from_r=None,
                    from_c=0,
                    upto_r=None,
                    upto_c=span_a2n(m[1]),
                    widget=widget,
                )
            elif m := PATTERN_CELL_START.match(key):
                return span_dict(
                    from_r=int(m[2]) - 1,
                    from_c=span_a2i(m[1]),
                    upto_r=None,
                    upto_c=None,
                    widget=widget,
                )
            elif m := PATTERN_CELL_END.match(key):
                return span_dict(
                    from_r=0,
                    from_c=0,
                    upto_r=int(m[2]),
                    upto_c=span_a2n(m[1]),
                    widget=widget,
                )
            elif m := PATTERN_CELL_COL.match(key):
                return span_dict(
                    from_r=int(m[2]) - 1,
                    from_c=span_a2i(m[1]),
                    upto_r=None,
                    upto_c=span_a2n(m[3]),
                    widget=widget,
                )
            elif m := PATTERN_CELL_ROW.match(key):
                return span_dict(
                    from_r=int(m[2]) - 1,
                    from_c=span_a2i(m[1]),
                    upto_r=int(m[3]),
                    upto_c=None,
                    widget=widget,
                )
            elif PATTERN_ALL.match(key):
                return span_dict(
                    from_r=0,
                    from_c=None,
                    upto_r=None,
                    upto_c=None,
                    widget=widget,
                )
            else:
                return f"'{key}' could not be converted to span."

    except ValueError as error:
        return f"Error, '{key}' could not be converted to span: {error}"


def span_is_cell(span: Span) -> bool:
    return (
        isinstance(span["from_r"], int)
        and isinstance(span["from_c"], int)
        and isinstance(span["upto_r"], int)
        and isinstance(span["upto_c"], int)
        and span["upto_r"] - span["from_r"] == 1
        and span["upto_c"] - span["from_c"] == 1
    )


def span_to_cell(span: Span) -> tuple[int, int]:
    # assumed that span arg has been tested by 'span_is_cell()'
    return (span["from_r"], span["from_c"])


def span_ranges(
    span: Span,
    totalrows: int | Callable,
    totalcols: int | Callable,
) -> tuple[Generator[int], Generator[int]]:
    rng_from_r = 0 if span.from_r is None else span.from_r
    rng_from_c = 0 if span.from_c is None else span.from_c
    rng_upto_r = (totalrows() if isinstance(totalrows, Callable) else totalrows) if span.upto_r is None else span.upto_r
    rng_upto_c = (totalcols() if isinstance(totalcols, Callable) else totalcols) if span.upto_c is None else span.upto_c
    return range(rng_from_r, rng_upto_r), range(rng_from_c, rng_upto_c)


def span_froms(
    span: Span,
) -> tuple[int, int]:
    from_r = 0 if span.from_r is None else span.from_r
    from_c = 0 if span.from_c is None else span.from_c
    return from_r, from_c


def del_named_span_options(options: dict, itr: Iterator[Hashable], type_: str) -> None:
    for k in itr:
        if k in options and type_ in options[k]:
            del options[k][type_]


def del_named_span_options_nested(
    options: dict, itr1: Iterator[Hashable], itr2: Iterator[Hashable], type_: str
) -> None:
    for k1 in itr1:
        for k2 in itr2:
            k = (k1, k2)
            if k in options and type_ in options[k]:
                del options[k][type_]


def mod_note(options: dict, key: int | tuple[int, int], note: str | None, readonly: bool = True) -> dict:
    if note is not None:
        if key not in options:
            options[key] = {}
        options[key]["note"] = {"note": note, "readonly": readonly}
    else:
        if key in options and "note" in options[key]:
            del options[key]["note"]
    return options


def add_highlight(
    options: dict,
    key: int | tuple[int, int],
    bg: bool | None | str = False,
    fg: bool | None | str = False,
    end: bool | None = None,
    overwrite: bool = True,
) -> dict:
    if key not in options:
        options[key] = {}
    if overwrite or "highlight" not in options[key]:
        options[key]["highlight"] = Highlight(
            bg=None if bg is False else bg,
            fg=None if fg is False else fg,
            end=False if end is None else end,
        )
    else:
        options[key]["highlight"] = Highlight(
            bg=options[key]["highlight"].bg if bg is False else bg,
            fg=options[key]["highlight"].fg if fg is False else fg,
            end=options[key]["highlight"].end if end is None else end,
        )
    return options


def set_readonly(
    options: dict,
    key: int | tuple[int, int],
    readonly: bool = True,
) -> dict:
    if readonly:
        if key not in options:
            options[key] = {}
        options[key]["readonly"] = True
    else:
        if key in options and "readonly" in options[key]:
            del options[key]["readonly"]
    return options


def convert_align(align: str | None) -> str | None:
    if isinstance(align, str):
        a = align.lower()
        if a == "global":
            return None
        elif a in ("c", "center", "centre", "n"):
            return "n"
        elif a in ("w", "west", "left", "nw"):
            return "nw"
        elif a in ("e", "east", "right", "ne"):
            return "ne"
    elif align is None:
        return None
    raise ValueError(align_value_error)


def set_align(
    options: dict,
    key: int | tuple[int, int],
    align: str | None = None,
) -> dict:
    if align:
        if key not in options:
            options[key] = {}
        options[key]["align"] = align
    else:
        if key in options and "align" in options[key]:
            del options[key]["align"]


def del_from_options(
    options: dict,
    key: str,
    coords: int | Iterator[int | tuple[int, int]] | None = None,
) -> dict:
    if isinstance(coords, int):
        if coords in options and key in options[coords]:
            del options[coords]
    elif is_iterable(coords):
        for coord in coords:
            if coord in options and key in options[coord]:
                del options[coord]
    else:
        for d in options.values():
            if key in d:
                del d[key]


def add_to_options(
    options: dict,
    coords: int | tuple[int, int],
    key: str,
    value: Any,
) -> dict:
    if coords not in options:
        options[coords] = {}
    options[coords][key] = value


def fix_format_kwargs(kwargs: dict) -> dict:
    if kwargs["formatter"] is None:
        if kwargs["nullable"]:
            if isinstance(kwargs["datatypes"], (list, tuple)):
                kwargs["datatypes"] = tuple(kwargs["datatypes"]) + (type(None),)
            else:
                kwargs["datatypes"] = (kwargs["datatypes"], type(None))
        elif (isinstance(kwargs["datatypes"], (list, tuple)) and type(None) in kwargs["datatypes"]) or kwargs[
            "datatypes"
        ] is type(None):
            raise TypeError("Non-nullable cells cannot have NoneType as a datatype.")
    if not isinstance(kwargs["invalid_value"], str):
        kwargs["invalid_value"] = f"{kwargs['invalid_value']}"
    return kwargs


def span_idxs_post_move(
    new_idxs: dict[int, int],
    full_new_idxs: dict[int, int],
    total: int,
    span: Span,
    axis: str,  # 'r' or 'c'
) -> tuple[int | None]:
    """
    Calculates the position of a span after moving rows/columns
    """
    if isinstance(span[f"upto_{axis}"], int):
        oldfrom, oldupto = int(span[f"from_{axis}"]), int(span[f"upto_{axis}"]) - 1
        newfrom = full_new_idxs[oldfrom]
        newupto = full_new_idxs[oldupto]
        if newfrom > newupto:
            newfrom, newupto = newupto, newfrom
        newupto += 1
        oldupto_colrange = int(span[f"upto_{axis}"])
        newupto_colrange = newupto
    else:
        oldfrom = int(span[f"from_{axis}"])
        newfrom = 0 if not oldfrom else full_new_idxs[oldfrom]
        newupto = None
        oldupto_colrange = total
        newupto_colrange = oldupto_colrange
    return oldupto_colrange, newupto_colrange, newfrom, newupto


def mod_span(
    to_set_to: Span,
    to_set_from: Span,
    from_r: int | None = None,
    from_c: int | None = None,
    upto_r: int | None = None,
    upto_c: int | None = None,
) -> Span:
    to_set_to.kwargs = to_set_from.kwargs
    to_set_to.type_ = to_set_from.type_
    to_set_to.table = to_set_from.table
    to_set_to.index = to_set_from.index
    to_set_to.header = to_set_from.header
    to_set_to.from_r = from_r
    to_set_to.from_c = from_c
    to_set_to.upto_r = upto_r
    to_set_to.upto_c = upto_c
    return to_set_to


def mod_span_widget(span: Span, widget: Any) -> Span:
    span.widget = widget
    return span


def mod_event_val(
    event_data: EventDataDict,
    val: Any,
    loc: Loc | None = None,
    row: int | None = None,
    column: int | None = None,
) -> EventDataDict:
    event_data.value = val
    if isinstance(loc, tuple):
        event_data.loc = Loc(*loc)
        event_data.row = loc[0]
        event_data.column = loc[1]

    elif isinstance(row, int):
        event_data.loc = Loc(row=row)
        event_data.row = row

    elif isinstance(column, int):
        event_data.loc = Loc(column=column)
        event_data.column = column

    return event_data


def pop_positions(
    itr: Callable,
    to_pop: dict[int, int],  # displayed index: data index
    save_to: dict[int, int],
) -> Iterator[int]:
    for i, pos in enumerate(itr()):
        if i in to_pop:
            save_to[to_pop[i]] = pos
        else:
            yield pos


def get_horizontal_gridline_points(
    left: float,
    stop: float,
    positions: list[float],
    start: int,
    end: int,
) -> list[int | float]:
    return list(
        chain.from_iterable(
            (
                left - 1,
                positions[r],
                stop,
                positions[r],
                left - 1,
                positions[r],
                left - 1,
                positions[r + 1] if len(positions) - 1 > r else positions[r],
            )
            for r in range(start, end)
        )
    )


def get_vertical_gridline_points(
    top: float,
    stop: float,
    positions: list[float],
    start: int,
    end: int,
) -> list[int | float]:
    return list(
        chain.from_iterable(
            (
                positions[c],
                top - 1,
                positions[c],
                stop,
                positions[c],
                top - 1,
                positions[c + 1] if len(positions) - 1 > c else positions[c],
                top - 1,
            )
            for c in range(start, end)
        )
    )


def safe_copy(value: Any) -> Any:
    """
    Attempts to create a deep copy of the input value. If copying fails,
    returns the original value.

    Args:
        value: Any Python object to be copied

    Returns:
        A deep copy of the value if possible, otherwise the original value
    """
    try:
        # Try deep copy first for most objects
        return copy.deepcopy(value)
    except Exception:
        try:
            # For types that deepcopy might fail on, try shallow copy
            return copy.copy(value)
        except Exception:
            try:
                # For built-in immutable types, return as-is
                if isinstance(value, (int, float, str, bool, bytes, tuple, frozenset)):
                    return value
                # For None
                if value is None:
                    return None
                # For functions, return same function
                if isinstance(value, Callable):
                    return value
                # For modules
                if isinstance(value, ModuleType):
                    return value
                # For classes
                if isinstance(value, type):
                    return value
            except Exception:
                pass
            # If all copy attempts fail, return original value
            return value
