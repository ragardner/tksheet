from typing import Union, Tuple

from .other_classes import (
    Span,
)

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
