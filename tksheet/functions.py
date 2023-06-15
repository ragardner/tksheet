from __future__ import annotations

import bisect
import pickle
import zlib
from collections.abc import (
    Callable,
    Iterator,
    Generator,
)
from .types import (
    Span,
)
from .other_classes import (
    DotDict,
    SpanDict,
)
from functools import partial
from itertools import islice
import re

compress = partial(zlib.compress, level=1)
pickle_obj = partial(pickle.dumps, protocol=pickle.HIGHEST_PROTOCOL)
unpickle_obj = pickle.loads


def pickle_compress(obj: object) -> bytes:
    return compress(pickle_obj(obj))


def decompress_load(b: bytes) -> object:
    return pickle.loads(zlib.decompress(b))


def tksheet_type_error(kwarg, valid_types, not_type):
    valid_types = ", ".join(f"{type_}" for type_ in valid_types)
    return f"Argument '{kwarg}' must be one of the following types: {valid_types}, " f"not {type(not_type)}."


def dropdown_search_function(search_for, data):
    search_len = len(search_for)
    match_rn = float("inf")
    match_st = float("inf")
    match_len_diff = float("inf")
    for rn, row in enumerate(data):
        dd_val = rf"{row[0]}".lower()
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
    if match_rn != float("inf"):
        return match_rn
    return None


def event_dict(
    name: str = None,
    sheet: object = None,
    boxes: None | dict | tuple = None,
    cells_table: None | dict = None,
    cells_header: None | dict = None,
    cells_index: None | dict = None,
    selected: None | tuple = None,
    data: object = None,
    key: None | str = None,
    value: object = None,
    location: None | int | tuple[int] = None,
    resized_rows: None | dict = None,
    resized_columns: None | dict = None,
    # resized_index: None, dict] = None,
    # resized_header: None, dict] = None,
    being_selected: None | tuple = None,
    named_spans: None | dict = None,
    **kwargs,
) -> dict:
    return DotDict(
        {
            "eventname": "" if name is None else name,
            "sheetname": "!sheet" if sheet is None else sheet,
            "cells": {
                "table": {} if cells_table is None else cells_table,
                "header": {} if cells_header is None else cells_header,
                "index": {} if cells_index is None else cells_index,
            },
            "moved": {
                "rows": {},
                "columns": {},
            },
            "added": {
                "rows": {},
                "columns": {},
            },
            "deleted": {
                "rows": {},
                "columns": {},
                "header": {},
                "index": {},
                "column_widths": {},
                "row_heights": {},
                "options": {},
                "displayed_rows": None,
                "displayed_columns": None,
            },
            "named_spans": {} if named_spans is None else named_spans,
            "selection_boxes": {} if boxes is None else {boxes[:-1]: boxes[-1]} if isinstance(boxes, tuple) else boxes,
            "selected": tuple() if selected is None else selected,
            "being_selected": tuple() if being_selected is None else being_selected,
            "data": [] if data is None else data,
            "key": "" if key is None else key,
            "value": None if value is None else value,
            "location": tuple() if location is None else location,
            "resized": {
                "rows": {} if resized_rows is None else resized_rows,
                "columns": {} if resized_columns is None else resized_columns,
                # "header": {} if resized_header is None else resized_header,
                # "index": {} if resized_index is None else resized_index,
            },
        }
    )


def change_eventname(event_dict: dict, newname: str) -> dict:
    return {**event_dict, **{"eventname": newname}}


def ev_stack_dict(d):
    return DotDict(
        {
            "name": d["eventname"],
            "data": pickle_compress(d),
        }
    )


def len_to_idx(n: int) -> int:
    if n < 1:
        return 0
    return n - 1


def get_dropdown_kwargs(
    values=[],
    set_value=None,
    state="normal",
    redraw=True,
    selection_function=None,
    modified_function=None,
    search_function=dropdown_search_function,
    validate_input=True,
    text=None,
    **kwargs,
):
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


def get_dropdown_dict(**kwargs):
    return {
        "values": kwargs["values"],
        "window": "no dropdown open",
        "canvas_id": "no dropdown open",
        "select_function": kwargs["selection_function"],
        "modified_function": kwargs["modified_function"],
        "search_function": kwargs["search_function"],
        "validate_input": kwargs["validate_input"],
        "text": kwargs["text"],
        "state": kwargs["state"],
    }


def get_checkbox_kwargs(
    checked=False,
    state="normal",
    redraw=True,
    check_function=None,
    text="",
    **kwargs,
):
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


def str_to_coords(s: str) -> None | tuple[int, ...]:
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


def get_n2a(n: int = 0, _type: str = "numbers") -> str:
    if _type == "letters":
        return num2alpha(n)
    elif _type == "numbers":
        return f"{n + 1}"
    else:
        return f"{num2alpha(n)} {n + 1}"


def get_index_of_gap_in_sorted_integer_seq_forward(seq, start=0):
    prevn = seq[start]
    for idx, n in enumerate(islice(seq, start + 1, None), start + 1):
        if n != prevn + 1:
            return idx
        prevn = n
    return None


def get_index_of_gap_in_sorted_integer_seq_reverse(seq, start=0):
    prevn = seq[start]
    for idx, n in zip(range(start, -1, -1), reversed(seq[:start])):
        if n != prevn - 1:
            return idx
        prevn = n
    return None


def get_seq_without_gaps_at_index(seq, position, get_st_end=False):
    start_idx = bisect.bisect_left(seq, position)
    forward_gap = get_index_of_gap_in_sorted_integer_seq_forward(seq, start_idx)
    reverse_gap = get_index_of_gap_in_sorted_integer_seq_reverse(seq, start_idx)
    if forward_gap is not None:
        seq[:] = seq[:forward_gap]
    if reverse_gap is not None:
        seq[:] = seq[reverse_gap:]
    if get_st_end:
        return seq[0], seq[-1]
    return seq


def consecutive_chunks(seq: list) -> list:
    if not seq:
        yield seq
    start = 0
    end = 0
    for index, value in enumerate(seq):
        if index < len(seq) - 1:
            if seq[index + 1] > value + 1:
                end = index + 1
                yield seq[start:end]
                start = end
        else:
            yield seq[start : len(seq)]


def is_contiguous(seq: list[int]) -> bool:
    itr = iter(seq)
    prev = next(itr)
    if not all(i == (prev := prev + 1) for i in itr):
        return False
    return True


def move_elements_by_mapping(
    seq: list[...],
    new_idxs: dict,
    old_idxs: dict,
) -> list[...]:
    # move elements of a list around, displacing
    # other elements based on mapping
    # of {old index: new index, ...}

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
    seq: list[...],
    move_to: int,
    to_move: list[int],
) -> list[...]:
    return move_elements_by_mapping(
        seq,
        *get_new_indexes(
            len(seq),
            move_to,
            to_move,
            get_inverse=True,
        ),
    )


def get_new_indexes(
    seqlen: int,
    move_to: int,
    to_move: list[int],
    keep_len: bool = True,
    get_inverse: bool = False,
) -> tuple[dict]:
    # returns
    # {old idx: new idx, ...}

    # if new indexes have to be pushed back by
    # moving elements towards the end of the list
    if keep_len and len(to_move) > seqlen - move_to:
        offset = len(to_move) - 1
        new_idxs = {to_move[i]: move_to - offset + i for i in range(len(to_move))}
    else:
        new_idxs = {to_move[i]: move_to + i for i in range(len(to_move))}
    if get_inverse:
        return new_idxs, dict(zip(new_idxs.values(), new_idxs))
    return new_idxs


def insert_items(seq: list | tuple, to_insert: dict, seq_len_func: Callable | None = None) -> list:
    # faster method of inserting many items into a list
    # using a dict of reverse sorted order of
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


def get_checkbox_points(x1, y1, x2, y2, radius=8):
    return [
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
    ]


def diff_list(seq: list[float, ...]) -> list[int, ...]:
    return [
        int(b - a)
        for a, b in zip(
            seq,
            islice(seq, 1, None),
        )
    ]


def diff_gen(seq: list[float, ...]) -> Generator[int, ...]:
    return (
        int(b - a)
        for a, b in zip(
            seq,
            islice(seq, 1, None),
        )
    )


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
    self,
    options: dict,
    formatter: object = None,
) -> Generator[tuple[int, int], ...] | Generator[int, ...]:
    if formatter is None:
        return (k for k, dct in options.items() if "format" in dct)
    else:
        return (k for k, dct in options.items() if "format" in dct and dct["format"]["formatter"] == formatter)


def try_binding(
    binding: None | Callable,
    event: dict,
    new_name: None | str = None,
) -> bool:
    if binding is not None:
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
    transpose: bool = False,
    ndim: int | None = None,
    convert: object = None,
    undo: bool = False,
    widget: object = None,
) -> Span:
    d: Span = SpanDict(
        {
            "from_r": from_r,
            "from_c": from_c,
            "upto_r": upto_r,
            "upto_c": upto_c,
            "type_": "" if type_ is None else type_,
            "name": "" if name is None else name,
            "kwargs": {} if kwargs is None else kwargs,
            "table": table,
            "header": header,
            "index": index,
            "tdisp": tdisp,
            "idisp": idisp,
            "hdisp": hdisp,
            "transpose": transpose,
            "ndim": ndim,
            "convert": convert,
            "undo": undo,
            "widget": widget,
        }
    )
    return d


def coords_to_span(
    from_r: int | None = None,
    from_c: int | None = None,
    upto_r: int | None = None,
    upto_c: int | None = None,
    name: str | None = None,
    spans: dict[str, Span] | None = None,
) -> Span:
    if isinstance(name, str) and isinstance(spans, dict) and name in spans:
        return spans[name]
    return span_dict(
        from_r=from_r,
        from_c=from_c,
        upto_r=upto_r,
        upto_c=upto_c,
    )


def key_to_span(
    key: str | int | slice,
    spans: dict[str, Span],
    widget: object = None,
) -> Span:
    if not isinstance(key, (str, int, slice)):
        return f"Key type must be either str, int or slice, not '{type(key)}'."

    if isinstance(key, int):
        """
        [int] - Whole row at that index
        """
        return span_dict(
            from_r=key - 1,
            from_c=None,
            upto_r=key,
            upto_c=None,
            widget=widget,
        )

    elif isinstance(key, slice):
        """
        [slice]
        """
        """
        [:] - All cells
        """
        if key.start is None and key.stop is None:
            """
            [:]
            """
            return span_dict(
                from_r=0,
                from_c=0,
                upto_r=None,
                upto_c=None,
                widget=widget,
            )
        """
        [1:3] - Rows 1, 2
        [:2] - Rows up to but not including 2
        [2:] - Rows starting from and including 2
        """
        start = 0 if key.start is None else key.start - 1
        return span_dict(
            from_r=start,
            from_c=None,
            upto_r=key.stop,
            upto_c=None,
            widget=widget,
        )

    elif isinstance(key, str):
        if not key:
            return f"Key cannot be an empty string."  # noqa: F541

        if key.startswith("<") and key.endswith(">"):
            if (key := key[1:-1]) in spans:
                """
                ["<name>"] - Surrounded by "<" ">" cells from a named range
                """
                return spans[key]
            return f"'{key}' not in named spans."

        key = key.upper()

        try:
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
                [":"] - All cells
                """
                return span_dict(
                    from_r=0,
                    from_c=0,
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
    if (
        isinstance(span["from_r"], int)
        and isinstance(span["from_c"], int)
        and isinstance(span["upto_r"], int)
        and isinstance(span["upto_c"], int)
        and span["upto_r"] - span["from_r"] == 1
        and span["upto_c"] - span["from_c"] == 1
    ):
        return True
    return False


def span_to_cell(span: Span) -> tuple[int, int]:
    # assumed that span arg has been tested by 'span_is_cell()'
    return (span["from_r"], span["from_c"])


def span_ranges(
    span: Span,
    totalrows: int | Callable,
    totalcols: int | Callable,
) -> tuple[Generator, Generator]:
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


def coords_tag_to_int_tuple(s: str) -> tuple[int, int, int, int] | tuple[int, int]:
    return tuple(map(int, filter(None, s.split("_"))))
