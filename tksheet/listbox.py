from __future__ import annotations
import tkinter as tk
# from collections.abc import Callable, Generator, Iterator, Sequence

from .sheet import (
    Sheet,
)


class ListBox(Sheet):
    def __init__(
        self,
        parent: tk.Misc,
    ) -> None:
        Sheet.__init__(
            self,
            parent=parent,
        )
        self.parent = parent
