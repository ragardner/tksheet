from __future__ import annotations

import bisect
import pickle
import zlib
from collections.abc import (
    Generator,
)
from functools import partial
from itertools import islice

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
    return {
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


def ev_stack_dict(d):
    return {
        "name": d["eventname"],
        "data": pickle_compress(d),
    }


def len_to_idx(n: int) -> int:
    if not n:
        return 0
    return n - 1


def named_span_dict(
    from_r: int | None = None,
    from_c: int | None = None,
    upto_r: int | None = None,
    upto_c: int | None = None,
    type_: str | None = None,
    name: str | None = None,
    table: bool | None = None,
    header: bool | None = None,
    index: bool | None = None,
    kwargs: dict | None = None,
) -> dict:
    return {
        "from_r": None if from_r is None else from_r,
        "from_c": None if from_c is None else from_c,
        "upto_r": None if upto_r is None else upto_r,
        "upto_c": None if upto_c is None else upto_c,
        "type_": "" if type_ is None else type_,
        "name": "" if name is None else name,
        "table": True if table is None else table,
        "header": False if header is None else header,
        "index": False if index is None else index,
        "kwargs": {} if kwargs is None else kwargs,
    }


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


def get_checkbox_kwargs(checked=False, state="normal", redraw=True, check_function=None, text="", **kwargs):
    return {
        "checked": checked,
        "state": state,
        "redraw": redraw,
        "check_function": check_function,
        "text": text,
    }


def get_checkbox_dict(**kwargs):
    return {
        "check_function": kwargs["check_function"],
        "state": kwargs["state"],
        "text": kwargs["text"],
    }


def is_iterable(o):
    if isinstance(o, str):
        return False
    try:
        iter(o)
        return True
    except Exception:
        return False


def str_to_coords(s: str) -> None | tuple[int, ...]:
    s = s.split(":")


def alpha2num(a: str) -> int:
    a = a.upper()
    n = 0
    orda = ord("A")
    for c in a:
        n = n * 26 + ord(c) - orda + 1
    return n - 1


def num2alpha(n: int) -> str:
    s = ""
    n += 1
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def get_n2a(n=0, _type="numbers"):
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


def convert_data_to_displayed_indexes(
    to_convert: list[int],
    displayed: list[int],
) -> list[int]:
    data_indexes = set(to_convert)
    return [i for i, e in enumerate(displayed) if e in data_indexes]


def convert_displayed_to_data_indexes(
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
