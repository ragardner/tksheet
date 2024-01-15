from __future__ import annotations

from platform import system as get_os
from collections import namedtuple


USER_OS: str = f"{get_os()}".lower()
ctrl_key: str = "Command" if USER_OS == "darwin" else "Control"
rc_binding: str = "<2>" if USER_OS == "darwin" else "<3>"
symbols_set: set[str] = set("""!#$%&'()*+,-./:;"@[]^_`{|}~>?= \\""")
nonelike: set[object] = {None, "none", ""}
truthy: set[object] = {True, "true", "t", "yes", "y", "on", "1", 1, 1.0}
falsy: set[object] = {False, "false", "f", "no", "n", "off", "0", 0, 0.0}
val_modifying_options: set[str, str, str] = {"checkbox", "format", "dropdown"}
named_span_types = (
    "format",
    "highlight",
    "dropdown",
    "checkbox",
    "readonly",
    "align",
)
arrowkey_bindings_helper: dict[str, str] = {
    "tab": "Tab",
    "up": "Up",
    "right": "Right",
    "left": "Left",
    "down": "Down",
    "prior": "Prior",
    "next": "Next",
}
emitted_events: set[str] = {
    "<<SheetModified>>",
    "<<SheetRedrawn>>",
}
FontTuple = namedtuple("FontTuple", "family size style")


def get_font() -> FontTuple:
    return FontTuple("Calibri", 13 if USER_OS == "darwin" else 11, "normal")


def get_index_font() -> FontTuple:
    return FontTuple("Calibri", 13 if USER_OS == "darwin" else 11, "normal")


def get_header_font() -> FontTuple:
    return FontTuple("Calibri", 13 if USER_OS == "darwin" else 11, "normal")

