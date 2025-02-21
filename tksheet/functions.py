from __future__ import annotations

import csv
import io
import pickle
import re
import tkinter as tk
from bisect import bisect_left, bisect_right
from collections import deque
from collections.abc import Callable, Generator, Hashable, Iterable, Iterator, Sequence
from difflib import SequenceMatcher
from itertools import islice, repeat
from typing import Literal

from .colors import color_map
from .constants import align_value_error, symbols_set
from .formatters import to_bool
from .other_classes import Box_nt, DotDict, EventDataDict, Highlight, Loc, Span
from .tksheet_types import AnyIter

unpickle_obj = pickle.loads
lines_re = re.compile(r"[^\n]+")


def wrap_text(
    text: str,
    max_width: int,
    max_lines: int,
    char_width_fn: Callable,
    widths: dict[str, int],
    wrap: Literal["", "c", "w"] = "",
    start_line: int = 0,
) -> Generator[str]:
    lines = (match.group() for match in lines_re.finditer(text))
    current_line = []
    total_lines = 0
    line_width = 0

    if not wrap:
        for line in lines:
            line_width = 0
            current_line = []
            for char in line:
                char_width = widths.get(char, char_width_fn(char))
                line_width += char_width
                if line_width >= max_width:
                    break
                current_line.append(char)

            if total_lines >= start_line:
                yield "".join(current_line)

            # Count the line whether it's empty or not
            total_lines += 1
            if total_lines >= max_lines:
                return

    elif wrap == "c":
        for line in lines:
            for char in line:
                char_width = widths.get(char, char_width_fn(char))

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
                        current_line.append(char)
                        line_width = char_width
                # adding char to line is okay
                else:
                    current_line.append(char)
                    line_width += char_width

            if total_lines >= start_line:
                yield "".join(current_line)

            total_lines += 1
            if total_lines >= max_lines:
                return
            current_line = []  # Reset for next line
            line_width = 0

    elif wrap == "w":
        space_width = widths.get(" ", char_width_fn(" "))

        for line in lines:
            words = line.split()
            for i, word in enumerate(words):
                # if we're going to next word and
                # if a space fits on the end of the current line we add one
                if i and line_width + space_width < max_width:
                    current_line.append(" ")
                    line_width += space_width

                # check if word will fit
                word_width = 0
                word_char_widths = []
                for char in word:
                    word_char_widths.append((w := widths.get(char, char_width_fn(char))))
                    word_width += w

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

                    for char, w in zip(word, word_char_widths):
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
                                current_line.append(char)
                                line_width = w
                        # adding char to line is okay
                        else:
                            current_line.append(char)
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


def recursive_bind(widget: tk.Misc, event: str, callback: Callable) -> None:
    widget.bind(event, callback)
    for child in widget.winfo_children():
        recursive_bind(child, event, callback)


def tksheet_type_error(kwarg: str, valid_types: list[str], not_type: object) -> str:
    valid_types = ", ".join(f"{type_}" for type_ in valid_types)
    return f"Argument '{kwarg}' must be one of the following types: {valid_types}, not {type(not_type)}."


def new_tk_event(keysym: str) -> tk.Event:
    event = tk.Event()
    event.keysym = keysym
    return event


def event_has_char_key(event: object) -> bool:
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


def dropdown_search_function(search_for: str, data: Iterable[object]) -> None | int:
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


def selection_box_tup_to_dict(box: tuple) -> dict:
    return {Box_nt(*box[:-1]): box[-1]}


def event_dict(
    name: str = None,
    sheet: object = None,
    widget: tk.Canvas | None = None,
    boxes: None | dict | tuple = None,
    cells_table: None | dict = None,
    cells_header: None | dict = None,
    cells_index: None | dict = None,
    selected: None | tuple = None,
    data: object = None,
    key: None | str = None,
    value: object = None,
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
        selection_boxes=(
            {} if boxes is None else selection_box_tup_to_dict(boxes) if isinstance(boxes, tuple) else boxes
        ),
        selected=tuple() if selected is None else selected,
        being_selected=tuple() if being_selected is None else being_selected,
        data=[] if data is None else data,
        key="" if key is None else key,
        value=None if value is None else value,
        loc=tuple() if loc is None else loc,
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
    return n - 1


def b_index(sorted_seq: Sequence[int], num_to_index: int) -> int:
    """
    Designed to be a faster way of finding the index of an int
    in a sorted list of ints than list.index()
    """
    if (idx := bisect_left(sorted_seq, num_to_index)) == len(sorted_seq) or sorted_seq[idx] != num_to_index:
        raise ValueError(f"{num_to_index} is not in Sequence")
    return idx


def try_b_index(sorted_seq: Sequence[int], num_to_index: int) -> int | None:
    if (idx := bisect_left(sorted_seq, num_to_index)) == len(sorted_seq) or sorted_seq[idx] != num_to_index:
        return None
    return idx


def bisect_in(sorted_seq: Sequence[int], num: int) -> bool:
    """
    Faster than 'num in sorted_seq'
    """
    try:
        return sorted_seq[bisect_left(sorted_seq, num)] == num
    except Exception:
        return False


# def push_n(num: int, sorted_seq: Sequence[int]) -> int:
#     if num < sorted_seq[0]:
#         return num
#     else:
#         for e in sorted_seq:
#             if num >= e:
#                 num += 1
#             else:
#                 return num
#         return num


def push_n(num: int, sorted_seq: Sequence[int]) -> int:
    if num < sorted_seq[0]:
        return num
    low = num
    high = num + len(sorted_seq)
    while low < high:
        mid = (low + high) // 2
        if num + bisect_right(sorted_seq, mid) <= mid:
            high = mid
        else:
            low = mid + 1
    return low


def get_dropdown_kwargs(
    values: list = [],
    set_value: object = None,
    state: str = "normal",
    redraw: bool = True,
    selection_function: Callable | None = None,
    modified_function: Callable | None = None,
    search_function: Callable = dropdown_search_function,
    validate_input: bool = True,
    text: None | str = None,
) -> dict:
    return {
        "values": values,
        "set_value": set_value,
        "state": state,
        "redraw": redraw,
        "selection_function": selection_function,
        "modified_function": modified_function,
        "search_function": search_function,
        "validate_input": validate_input,
        "text": text,
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


def is_iterable(o: object) -> bool:
    if isinstance(o, str):
        return False
    try:
        iter(o)
        return True
    except Exception:
        return False


def int_x_iter(i: AnyIter[int] | int) -> AnyIter[int]:
    if isinstance(i, int):
        return (i,)
    return i


def int_x_tuple(i: AnyIter[int] | int) -> tuple[int]:
    if isinstance(i, int):
        return (i,)
    return tuple(i)


def unpack(t: tuple[object] | tuple[AnyIter[object]]) -> tuple[object]:
    if not len(t):
        return t
    if is_iterable(t[0]) and len(t) == 1:
        return t[0]
    return t


def is_type_int(o: object) -> bool:
    return isinstance(o, int) and not isinstance(o, bool)


def force_bool(o: object) -> bool:
    try:
        return to_bool(o)
    except Exception:
        return False


def alpha2idx(a: str) -> int | None:
    try:
        n = 0
        orda = ord("A")
        for c in a.upper():
            n = n * 26 + ord(c) - orda + 1
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


def down_cell_within_box(
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
    it: AnyIter[object],
) -> object:
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


def index_exists(seq: Sequence[object], index: int) -> bool:
    try:
        seq[index]
        return True
    except Exception:
        return False


def add_to_displayed(displayed: list[int], to_add: Iterable[int]) -> list[int]:
    # assumes to_add is sorted in reverse
    for i in reversed(to_add):
        ins = bisect_left(displayed, i)
        displayed[ins:] = [i] + [e + 1 for e in islice(displayed, ins, None)]
    return displayed


def move_elements_by_mapping(
    seq: list[object],
    new_idxs: dict[int, int],
    old_idxs: dict[int, int] | None = None,
) -> list[object]:
    # move elements of a list around
    # displacing other elements based on mapping
    # new_idxs = {old index: new index, ...}
    # old_idxs = {new index: old index, ...}
    if old_idxs is None:
        old_idxs = dict(zip(new_idxs.values(), new_idxs))
    remaining_values = (e for i, e in enumerate(seq) if i not in new_idxs)
    return [seq[old_idxs[i]] if i in old_idxs else next(remaining_values) for i in range(len(seq))]


def move_elements_to(
    seq: list[object],
    move_to: int,
    to_move: list[int],
) -> list[object]:
    return move_elements_by_mapping(
        seq,
        *get_new_indexes(
            move_to,
            to_move,
            get_inverse=True,
        ),
    )


def get_new_indexes(
    move_to: int,
    to_move: list[int],
    get_inverse: bool = False,
) -> tuple[dict]:
    """
    returns {old idx: new idx, ...}
    """
    offset = sum(1 for i in to_move if i < move_to)
    new_idxs = range(move_to - offset, move_to - offset + len(to_move))
    new_idxs = {old: new for old, new in zip(to_move, new_idxs)}
    if get_inverse:
        return new_idxs, dict(zip(new_idxs.values(), new_idxs))
    return new_idxs


def insert_items(
    seq: list[object],
    to_insert: dict[int, object],
    seq_len_func: Callable | None = None,
) -> list[object]:
    if to_insert:
        if seq_len_func and next(iter(to_insert)) >= len(seq) + len(to_insert):
            seq_len_func(next(iter(to_insert)) - len(to_insert))
        for idx, v in reversed(to_insert.items()):
            seq[idx:idx] = [v]
    return seq


def del_placeholder_dict_key(
    d: dict[Hashable, object],
    k: Hashable,
    v: object,
    p: tuple = tuple(),
) -> dict[Hashable, object]:
    if p in d:
        del d[p]
    d[k] = v
    return d


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
) -> tuple[float]:
    if y2 - y1 < 2 or x2 - x1 < 2:
        return x1, y1, x2, y1, x2, y2, x1, y2
    return (
        x1 + radius,
        y1,
        x1 + radius,
        y1,
        x2 - radius,
        y1,
        x2 - radius,
        y1,
        x2,
        y1,
        x2,
        y1 + radius,
        x2,
        y1 + radius,
        x2,
        y2 - radius,
        x2,
        y2 - radius,
        x2,
        y2,
        x2 - radius,
        y2,
        x2 - radius,
        y2,
        x1 + radius,
        y2,
        x1 + radius,
        y2,
        x1,
        y2,
        x1,
        y2 - radius,
        x1,
        y2 - radius,
        x1,
        y1 + radius,
        x1,
        y1 + radius,
        x1,
        y1,
    )


def diff_list(seq: list[float]) -> list[int]:
    return [
        int(b - a)
        for a, b in zip(
            seq,
            islice(seq, 1, None),
        )
    ]


def diff_gen(seq: list[float]) -> Generator[int]:
    return (
        int(b - a)
        for a, b in zip(
            seq,
            islice(seq, 1, None),
        )
    )


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
    start_row: int,
    start_col: int,
    total_cols: int,
    total_rows: int,
    reverse: bool = False,
) -> Generator[tuple[int, int]]:
    if reverse:
        # yield start cell
        yield (start_row, start_col)
        # yield any remaining cells in the starting row before the start column
        if start_col:
            for col in reversed(range(start_col)):
                yield (start_row, col)
        # yield any cells above start row
        for row in reversed(range(start_row)):
            for col in reversed(range(total_cols)):
                yield (row, col)
        # yield cells from bottom of table upward
        for row in range(total_rows - 1, start_row, -1):
            for col in reversed(range(total_cols)):
                yield (row, col)
        # yield any remaining cells in start row
        for col in range(total_cols - 1, start_col, -1):
            yield (start_row, col)
    else:
        # Yield cells from the start position to the end of the current row
        for col in range(start_col, total_cols):
            yield (start_row, col)
        # yield from the next row to the last row
        for row in range(start_row + 1, total_rows):
            for col in range(total_cols):
                yield (row, col)
        # yield from the beginning up to the start
        for row in range(start_row):
            for col in range(total_cols):
                yield (row, col)
        # yield any remaining cells in the starting row before the start column
        for col in range(start_col):
            yield (start_row, col)


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


def zip_fill_2nd_value(x: AnyIter[object], o: object) -> Generator[object, object]:
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
    formatter: object = None,
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
    widget: object = None,
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
    widget: object,
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
    widget: object = None,
) -> Span:
    if isinstance(key, Span):
        return key
    elif key is None:
        key = (None, None, None, None)
    elif not isinstance(key, (str, int, slice, list, tuple)):
        return f"Key type must be either str, int, list, tuple or slice, not '{type(key)}'."
    try:
        if isinstance(key, (list, tuple)):
            if isinstance(key[0], int) or key[0] is None:
                if len(key) == 2:
                    """
                    (int | None, int | None) -
                    (0, 0) - row 0, column 0 - the first cell
                    (0, None) - row 0, all columns
                    (None, 0) - column 0, all rows
                    """
                    return span_dict(
                        from_r=key[0] if isinstance(key[0], int) else 0,
                        from_c=key[1] if isinstance(key[1], int) else 0,
                        upto_r=(key[0] + 1) if isinstance(key[0], int) else None,
                        upto_c=(key[1] + 1) if isinstance(key[1], int) else None,
                        widget=widget,
                    )

                elif len(key) == 4:
                    """
                    (int | None, int | None, int | None, int | None) -
                    (from row,  from column, up to row, up to column)
                    """
                    return coords_to_span(
                        widget=widget,
                        from_r=key[0],
                        from_c=key[1],
                        upto_r=key[2],
                        upto_c=key[3],
                    )
                    # return span_dict(
                    #     from_r=key[0] if isinstance(key[0], int) else 0,
                    #     from_c=key[1] if isinstance(key[1], int) else 0,
                    #     upto_r=key[2] if isinstance(key[2], int) else None,
                    #     upto_c=key[3] if isinstance(key[3], int) else None,
                    #     widget=widget,
                    # )

            elif isinstance(key[0], (list, tuple)):
                """
                ((int | None, int | None), (int | None, int | None))

                First Sequence is start row and column
                Second Sequence is up to but not including row and column
                """
                return coords_to_span(
                    widget=widget,
                    from_r=key[0][0],
                    from_c=key[0][1],
                    upto_r=key[1][0],
                    upto_c=key[1][1],
                )
                # return span_dict(
                #     from_r=key[0][0] if isinstance(key[0][0], int) else 0,
                #     from_c=key[0][1] if isinstance(key[0][1], int) else 0,
                #     upto_r=key[1][0],
                #     upto_c=key[1][1],
                #     widget=widget,
                # )

        elif isinstance(key, int):
            """
            [int] - Whole row at that index
            """
            return span_dict(
                from_r=key,
                from_c=None,
                upto_r=key + 1,
                upto_c=None,
                widget=widget,
            )

        elif isinstance(key, slice):
            """
            [slice]
            """
            """
            [:] - All rows
            """
            if key.start is None and key.stop is None:
                """
                [:]
                """
                return span_dict(
                    from_r=0,
                    from_c=None,
                    upto_r=None,
                    upto_c=None,
                    widget=widget,
                )
            """
            [1:3] - Rows 1, 2
            [:2] - Rows up to but not including 2
            [2:] - Rows starting from and including 2
            """
            start = 0 if key.start is None else key.start
            return span_dict(
                from_r=start,
                from_c=None,
                upto_r=key.stop,
                upto_c=None,
                widget=widget,
            )

        elif isinstance(key, str):
            if not key:
                key = ":"

            if key.startswith("<") and key.endswith(">"):
                if (key := key[1:-1]) in spans:
                    """
                    ["<name>"] - Surrounded by "<" ">" cells from a named range
                    """
                    return spans[key]
                return f"'{key}' not in named spans."

            key = key.upper()

            if key.isdigit():
                """
                ["1"] - Row 0
                """
                return span_dict(
                    from_r=int(key) - 1,
                    from_c=None,
                    upto_r=int(key),
                    upto_c=None,
                    widget=widget,
                )

            if key.isalpha():
                """
                ["A"] - Column 0
                """
                return span_dict(
                    from_r=None,
                    from_c=alpha2idx(key),
                    upto_r=None,
                    upto_c=alpha2idx(key) + 1,
                    widget=widget,
                )

            splitk = key.split(":")
            if len(splitk) > 2:
                return f"'{key}' could not be converted to span."

            if len(splitk) == 1 and not splitk[0].isdigit() and not splitk[0].isalpha() and not splitk[0][0].isdigit():
                """
                ["A1"] - Cell (0, 0)
                """
                keys_digits = re.search(r"\d", splitk[0])
                if keys_digits:
                    digits_start = keys_digits.start()
                    if not digits_start:
                        return f"'{key}' could not be converted to span."
                    if digits_start:
                        key_row = splitk[0][digits_start:]
                        key_column = splitk[0][:digits_start]
                        return span_dict(
                            from_r=int(key_row) - 1,
                            from_c=alpha2idx(key_column),
                            upto_r=int(key_row),
                            upto_c=alpha2idx(key_column) + 1,
                            widget=widget,
                        )

            if not splitk[0] and not splitk[1]:
                """
                [":"] - All rows
                """
                return span_dict(
                    from_r=0,
                    from_c=None,
                    upto_r=None,
                    upto_c=None,
                    widget=widget,
                )

            if splitk[0].isdigit() and not splitk[1]:
                """
                ["2:"] - Rows starting from and including 1
                """
                return span_dict(
                    from_r=int(splitk[0]) - 1,
                    from_c=None,
                    upto_r=None,
                    upto_c=None,
                    widget=widget,
                )

            if splitk[1].isdigit() and not splitk[0]:
                """
                [":2"] - Rows up to and including 1
                """
                return span_dict(
                    from_r=0,
                    from_c=None,
                    upto_r=int(splitk[1]),
                    upto_c=None,
                    widget=widget,
                )

            if splitk[0].isdigit() and splitk[1].isdigit():
                """
                ["1:2"] - Rows 0, 1
                """
                return span_dict(
                    from_r=int(splitk[0]) - 1,
                    from_c=None,
                    upto_r=int(splitk[1]),
                    upto_c=None,
                    widget=widget,
                )

            if splitk[0].isalpha() and not splitk[1]:
                """
                ["B:"] - Columns starting from and including 2
                """
                return span_dict(
                    from_r=None,
                    from_c=alpha2idx(splitk[0]),
                    upto_r=None,
                    upto_c=None,
                    widget=widget,
                )

            if splitk[1].isalpha() and not splitk[0]:
                """
                [":B"] - Columns up to and including 2
                """
                return span_dict(
                    from_r=None,
                    from_c=0,
                    upto_r=None,
                    upto_c=alpha2idx(splitk[1]) + 1,
                    widget=widget,
                )

            if splitk[0].isalpha() and splitk[1].isalpha():
                """
                ["A:B"] - Columns 0, 1
                """
                return span_dict(
                    from_r=None,
                    from_c=alpha2idx(splitk[0]),
                    upto_r=None,
                    upto_c=alpha2idx(splitk[1]) + 1,
                    widget=widget,
                )

            m1 = re.search(r"\d", splitk[0])
            m2 = re.search(r"\d", splitk[1])
            m1start = m1.start() if m1 else None
            m2start = m2.start() if m2 else None
            if m1start and m2start:
                """
                ["A1:B1"] - Cells (0, 0), (0, 1)
                """
                c1 = splitk[0][:m1start]
                r1 = splitk[0][m1start:]
                c2 = splitk[1][:m2start]
                r2 = splitk[1][m2start:]
                return span_dict(
                    from_r=int(r1) - 1,
                    from_c=alpha2idx(c1),
                    upto_r=int(r2),
                    upto_c=alpha2idx(c2) + 1,
                    widget=widget,
                )

            if not splitk[0] and m2start:
                """
                [":B1"] - Cells (0, 0), (0, 1)
                """
                c2 = splitk[1][:m2start]
                r2 = splitk[1][m2start:]
                return span_dict(
                    from_r=0,
                    from_c=0,
                    upto_r=int(r2),
                    upto_c=alpha2idx(c2) + 1,
                    widget=widget,
                )

            if not splitk[1] and m1start:
                """
                ["A1:"] - Cells starting from and including (0, 0)
                """
                c1 = splitk[0][:m1start]
                r1 = splitk[0][m1start:]
                return span_dict(
                    from_r=int(r1) - 1,
                    from_c=alpha2idx(c1),
                    upto_r=None,
                    upto_c=None,
                    widget=widget,
                )

            if m1start and splitk[1].isalpha():
                """
                ["A1:B"] - All the cells starting from (0, 0)
                        expanding out to include column 1
                        but not including cells beyond column
                        1 and expanding down to include all rows
                    A   B   C   D
                1   x   x
                2   x   x
                3   x   x
                4   x   x
                ...
                """
                c1 = splitk[0][:m1start]
                r1 = splitk[0][m1start:]
                return span_dict(
                    from_r=int(r1) - 1,
                    from_c=alpha2idx(c1),
                    upto_r=None,
                    upto_c=alpha2idx(splitk[1]) + 1,
                    widget=widget,
                )

            if m1start and splitk[1].isdigit():
                """
                ["A1:2"] - All the cells starting from (0, 0)
                        expanding down to include row 1
                        but not including cells beyond row
                        1 and expanding out to include all
                        columns
                    A   B   C   D
                1   x   x   x   x
                2   x   x   x   x
                3
                4
                ...
                """
                c1 = splitk[0][:m1start]
                r1 = splitk[0][m1start:]
                return span_dict(
                    from_r=int(r1) - 1,
                    from_c=alpha2idx(c1),
                    upto_r=int(splitk[1]),
                    upto_c=None,
                    widget=widget,
                )

    except ValueError as error:
        return f"Error, '{key}' could not be converted to span: {error}"
    else:
        return f"'{key}' could not be converted to span."


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
    if span.upto_r is None:
        rng_upto_r = totalrows() if isinstance(totalrows, Callable) else totalrows
    else:
        rng_upto_r = span.upto_r
    if span.upto_c is None:
        rng_upto_c = totalcols() if isinstance(totalcols, Callable) else totalcols
    else:
        rng_upto_c = span.upto_c
    return range(rng_from_r, rng_upto_r), range(rng_from_c, rng_upto_c)


def span_froms(
    span: Span,
) -> tuple[int, int]:
    from_r = 0 if span.from_r is None else span.from_r
    from_c = 0 if span.from_c is None else span.from_c
    return from_r, from_c


def del_named_span_options(options: dict, itr: AnyIter[Hashable], type_: str) -> None:
    for k in itr:
        if k in options and type_ in options[k]:
            del options[k][type_]


def del_named_span_options_nested(options: dict, itr1: AnyIter[Hashable], itr2: AnyIter[Hashable], type_: str) -> None:
    for k1 in itr1:
        for k2 in itr2:
            k = (k1, k2)
            if k in options and type_ in options[k]:
                del options[k][type_]


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
    coords: int | AnyIter[int | tuple[int, int]] | None = None,
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
    value: object,
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
        if not oldfrom:
            newfrom = 0
        else:
            newfrom = full_new_idxs[oldfrom]
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


def mod_span_widget(span: Span, widget: object) -> Span:
    span.widget = widget
    return span


def mod_event_val(
    event_data: EventDataDict,
    val: object,
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
