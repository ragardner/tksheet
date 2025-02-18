from __future__ import annotations

from collections.abc import Iterable, Iterator
from typing import Literal, Tuple, Union

from .other_classes import Span

CreateSpanTypes = Union[
    Tuple[()],
    None,
    str,
    int,
    slice,
    Tuple[int, None, int, None],
    Tuple[Tuple[int, None, int, None], Tuple[int, None, int, None]],
    Tuple[int, None, int, None, int, None, int, None],
    Span,
]

CellPropertyKey = Literal[
    "format",
    "highlight",
    "dropdown",
    "checkbox",
    "readonly",
    "align",
]

AnyIter = Iterable | Iterator
