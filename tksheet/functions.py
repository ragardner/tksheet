from __future__ import annotations

import bisect
import csv
import io
import pickle
import re
import tkinter as tk
import zlib
from collections import deque
from collections.abc import (
    Callable,
    Generator,
    Iterator,
    Sequence,
)
from functools import partial
from itertools import islice, repeat

from .formatters import (
    to_bool,
)
from .other_classes import (
    Box_nt,
    DotDict,
    EventDataDict,
    Highlight,
    Loc,
    Span,
)

compress = partial(zlib.compress, level=1)
pickle_obj = partial(pickle.dumps, protocol=pickle.HIGHEST_PROTOCOL)
unpickle_obj = pickle.loads


def get_csv_str_dialect(s: str, delimiters: str) -> csv.Dialect:
    try:
        return csv.Sniffer().sniff(s[:5000] if len(s) > 5000 else s, delimiters=delimiters)
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


def pickle_compress(obj: object) -> bytes:
    return compress(pickle_obj(obj))


def decompress_load(b: bytes) -> object:
    return pickle.loads(zlib.decompress(b))


def tksheet_type_error(kwarg: str, valid_types: list[str], not_type: object) -> str:
    valid_types = ", ".join(f"{type_}" for type_ in valid_types)
    return f"Argument '{kwarg}' must be one of the following types: {valid_types}, " f"not {type(not_type)}."


def new_tk_event(keysym: str) -> tk.Event:
    event = tk.Event()
    event.keysym = keysym
    return event


def dropdown_search_function(
    search_for: object,
    data: Sequence[object],
) -> None | int:
    search_len = len(search_for)
    # search_for in data
    match_rn = float("inf")
    match_st = float("inf")
    match_len_diff = float("inf")
    # data in search_for in case no match
    match_data_rn = float("inf")
    match_data_st = float("inf")
    match_data_numchars = 0
    for rn, row in enumerate(data):
        dd_val = rf"{row[0]}".lower()
        # checking if search text is in dropdown row
        st = dd_val.find(search_for)
        if st > -1:
            # priority is start index
            # if there's already a matching start
            # then compare the len difference
            len_diff = len(dd_val) - search_len
            if st < match_st or (st == match_st and len_diff < match_len_diff):
                match_rn = rn
                match_st = st
                match_len_diff = len_diff
        # fall back in case of no existing match
        elif match_rn == float("inf"):
            for numchars in range(2, search_len - 1):
                for from_idx in range(search_len - 1):
                    if from_idx + numchars > search_len:
                        break
                    st = dd_val.find(search_for[from_idx : from_idx + numchars])
                    if st > -1 and (
                        numchars > match_data_numchars or (numchars == match_data_numchars and st < match_data_st)
                    ):
                        match_data_rn = rn
                        match_data_st = st
                        match_data_numchars = numchars
    if match_rn != float("inf"):
        return match_rn
    elif match_data_rn != float("inf"):
        return match_data_rn
    return None


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
            displayed_rows=None,
            displayed_columns=None,
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
    )


def change_eventname(event_dict: EventDataDict, newname: str) -> EventDataDict:
    return EventDataDict({**event_dict, **{"eventname": newname}})


def pickled_event_dict(d: DotDict) -> DotDict:
    return DotDict(name=d["eventname"], data=pickle_compress(DotDict({k: v for k, v in d.items() if k != "widget"})))


def len_to_idx(n: int) -> int:
    if n < 1:
        return 0
    return n - 1


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


def int_x_iter(i: Iterator[int] | int) -> Iterator[int]:
    if isinstance(i, int):
        return (i,)
    return i


def unpack(t: tuple[object] | tuple[Iterator[object]]) -> tuple[object]:
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


def str_to_coords(s: str) -> None | tuple[int]:
    s = s.split(":")


def alpha2idx(a: str) -> int | None:
    try:
        a = a.upper()
        n = 0
        orda = ord("A")
        for c in a:
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


def get_n2a(n: int = 0, _type: str = "numbers") -> str:
    if _type == "letters":
        return num2alpha(n)
    elif _type == "numbers":
        return f"{n + 1}"
    return f"{num2alpha(n)} {n + 1}"


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
    start_idx = bisect.bisect_left(seq, position)
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
    start = 0
    for index, value in enumerate(seq, 1):
        try:
            if seq[index] > value + 1:
                yield seq[start], seq[index - 1] + 1
                start = index
        except Exception:
            yield seq[start], seq[-1] + 1


def is_contiguous(iterable: Iterator[int]) -> bool:
    itr = iter(iterable)
    prev = next(itr)
    return all(i == (prev := prev + 1) for i in itr)


def get_last(
    it: Iterator,
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


def move_elements_by_mapping(
    seq: list[object],
    new_idxs: dict[int, int],
    old_idxs: dict[int, int] | None = None,
) -> list[object]:
    # move elements of a list around, displacing
    # other elements based on mapping
    # of {old index: new index, ...}
    if old_idxs is None:
        old_idxs = dict(zip(new_idxs.values(), new_idxs))

    # create dummy list
    res = [0] * len(seq)

    # create generator of values yet to be put into res
    remaining = (e for i, e in enumerate(seq) if i not in new_idxs)

    # goes over res twice:
    # once to put elements being moved in new spots
    # then to fill remaining spots with remaining elements

    # fill new indexes in res
    if len(new_idxs) > int(len(seq) / 2) - 1:
        # if moving a lot of items better to do comprehension
        return [
            next(remaining) if i_ not in old_idxs else e_
            for i_, e_ in enumerate(seq[old_idxs[i]] if i in old_idxs else e for i, e in enumerate(res))
        ]
    else:
        # if just moving a few items assignments are fine
        for old, new in new_idxs.items():
            res[new] = seq[old]
        # fill remaining indexes
        return [next(remaining) if i not in old_idxs else e for i, e in enumerate(res)]


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


def insert_items(seq: list | tuple, to_insert: dict, seq_len_func: Callable | None = None) -> list:
    # inserts many items into a list using a dict of reverse sorted order of
    # {index: value, index: value, ...}
    res = []
    extended = 0
    for i, (idx, v) in enumerate(reversed(to_insert.items())):
        # need to extend seq if it's not long enough
        if seq_len_func and idx - i > len(seq):
            seq_len_func(idx - i - 1)
        res.extend(seq[extended : idx - i])
        extended += idx - i - extended
        res.append(v)
    res.extend(seq[extended:])
    seq = res
    return seq


def data_to_displayed_idxs(
    to_convert: list[int],
    displayed: list[int],
) -> list[int]:
    data_indexes = set(to_convert)
    return [i for i, e in enumerate(displayed) if e in data_indexes]


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
    radius: int = 8,
) -> tuple[float]:
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


def zip_fill_2nd_value(x: Iterator, o: object) -> Generator[object, object]:
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


def del_named_span_options(options: dict, itr: Iterator, type_: str) -> None:
    for k in itr:
        if k in options and type_ in options[k]:
            del options[k][type_]


def del_named_span_options_nested(options: dict, itr1: Iterator, itr2: Iterator, type_: str) -> None:
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
        elif a in ("c", "center", "centre"):
            return "center"
        elif a in ("w", "west", "left"):
            return "w"
        elif a in ("e", "east", "right"):
            return "e"
    elif align is None:
        return None
    raise ValueError("Align must be one of the following values: c, center, w, west, left, e, east, right")


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
    coords: int | Iterator | None = None,
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
    loc: Loc | int,
) -> EventDataDict:
    event_data.value = val
    event_data.loc = Loc(*loc)
    event_data.row = loc[0]
    event_data.column = loc[1]
    return event_data


def pop_positions(
    itr: Iterator[int],
    to_pop: dict[int, int],  # displayed index: data index
    save_to: dict[int, int],
) -> Iterator[int]:
    for i, pos in enumerate(itr()):
        if i in to_pop:
            save_to[to_pop[i]] = pos
        else:
            yield pos
