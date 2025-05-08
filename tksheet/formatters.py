from __future__ import annotations

from collections.abc import Callable
from contextlib import suppress
from typing import Any

from .constants import falsy, nonelike, truthy


def is_none_like(o: Any) -> bool:
    return (isinstance(o, str) and o.lower().replace(" ", "") in nonelike) or o in nonelike


def to_int(o: Any, **kwargs) -> int:
    if isinstance(o, int):
        return o
    return int(float(o))


def to_float(o: Any, **kwargs) -> float:
    if isinstance(o, float):
        return o
    return float(o)


def to_percentage(o: Any, **kwargs) -> float:
    if isinstance(o, float):
        return o
    if isinstance(o, str) and o.endswith("%"):
        return float(o.replace("%", "")) / 100
    return float(o)


def alt_to_percentage(o: Any, **kwargs) -> float:
    if isinstance(o, float):
        return o
    if isinstance(o, str) and o.endswith("%"):
        return float(o.replace("%", ""))
    return float(o)


def to_bool(val: Any, **kwargs) -> bool:
    if isinstance(val, bool):
        return val
    v = val.lower() if isinstance(val, str) else val
    _truthy = kwargs.get("truthy", truthy)
    _falsy = kwargs.get("falsy", falsy)
    if v in _truthy:
        return True
    elif v in _falsy:
        return False
    raise ValueError(f'Cannot map "{val}" to bool.')


def try_to_bool(o: Any, **kwargs) -> Any:
    try:
        return to_bool(o)
    except Exception:
        return o


def is_bool_like(o: Any, **kwargs) -> bool:
    try:
        to_bool(o)
        return True
    except Exception:
        return False


def to_str(o: Any, **kwargs: dict) -> str:
    return f"{o}"


def float_to_str(v: Any, **kwargs: dict) -> str:
    if isinstance(v, float):
        if v.is_integer():
            return f"{int(v)}"
        if "decimals" in kwargs and isinstance(kwargs["decimals"], int):
            if kwargs["decimals"]:
                return f"{round(v, kwargs['decimals'])}"
            return f"{int(round(v, kwargs['decimals']))}"
    return f"{v}"


def percentage_to_str(v: Any, **kwargs: dict) -> str:
    if isinstance(v, (int, float)):
        x = v * 100
        if isinstance(x, float):
            if x.is_integer():
                return f"{int(x)}%"
            if "decimals" in kwargs and isinstance(kwargs["decimals"], int):
                if kwargs["decimals"]:
                    return f"{round(x, kwargs['decimals'])}%"
                return f"{int(round(x, kwargs['decimals']))}%"
        return f"{x}%"
    return f"{v}%"


def alt_percentage_to_str(v: Any, **kwargs: dict) -> str:
    return f"{float_to_str(v)}%"


def bool_to_str(v: Any, **kwargs: dict) -> str:
    return f"{v}"


def int_formatter(
    datatypes: tuple[Any] | Any = int,
    format_function: Callable = to_int,
    to_str_function: Callable = to_str,
    **kwargs,
) -> dict:
    return formatter(
        datatypes=datatypes,
        format_function=format_function,
        to_str_function=to_str_function,
        **kwargs,
    )


def float_formatter(
    datatypes: tuple[Any] | Any = float,
    format_function: Callable = to_float,
    to_str_function: Callable = float_to_str,
    decimals: int = 2,
    **kwargs,
) -> dict:
    return formatter(
        datatypes=datatypes,
        format_function=format_function,
        to_str_function=to_str_function,
        decimals=decimals,
        **kwargs,
    )


def percentage_formatter(
    datatypes: tuple[Any] | Any = float,
    format_function: Callable = to_percentage,
    to_str_function: Callable = percentage_to_str,
    decimals: int = 2,
    **kwargs,
) -> dict:
    return formatter(
        datatypes=datatypes,
        format_function=format_function,
        to_str_function=to_str_function,
        decimals=decimals,
        **kwargs,
    )


def bool_formatter(
    datatypes: tuple[Any] | Any = bool,
    format_function: Callable = to_bool,
    to_str_function: Callable = bool_to_str,
    invalid_value: Any = "NA",
    truthy_values: set[Any] = truthy,
    falsy_values: set[Any] = falsy,
    **kwargs,
) -> dict:
    return formatter(
        datatypes=datatypes,
        format_function=format_function,
        to_str_function=to_str_function,
        invalid_value=invalid_value,
        truthy_values=truthy_values,
        falsy_values=falsy_values,
        **kwargs,
    )


def formatter(
    datatypes: tuple[Any] | Any,
    format_function: Callable,
    to_str_function: Callable = to_str,
    invalid_value: Any = "NaN",
    nullable: bool = True,
    pre_format_function: Callable | None = None,
    post_format_function: Callable | None = None,
    clipboard_function: Callable | None = None,
    **kwargs,
) -> dict:
    return {
        **{
            "datatypes": datatypes,
            "format_function": format_function,
            "to_str_function": to_str_function,
            "invalid_value": invalid_value,
            "nullable": nullable,
            "pre_format_function": pre_format_function,
            "post_format_function": post_format_function,
            "clipboard_function": clipboard_function,
        },
        **kwargs,
    }


def format_data(
    value: Any = "",
    datatypes: tuple[Any] | Any = int,
    nullable: bool = True,
    pre_format_function: Callable | None = None,
    format_function: Callable | None = to_int,
    post_format_function: Callable | None = None,
    **kwargs,
) -> Any:
    if pre_format_function:
        value = pre_format_function(value)
    if nullable and is_none_like(value):
        value = None
    else:
        with suppress(Exception):
            value = format_function(value, **kwargs)
    if post_format_function and isinstance(value, datatypes):
        value = post_format_function(value)
    return value


def data_to_str(
    value: Any = "",
    datatypes: tuple[Any] | Any = int,
    nullable: bool = True,
    invalid_value: Any = "NaN",
    to_str_function: Callable | None = None,
    **kwargs,
) -> str:
    if not isinstance(value, datatypes):
        return invalid_value
    elif value is None and nullable:
        return ""
    else:
        return to_str_function(value, **kwargs)


# Only used if MainTable.cell_str is called with get_displayed=False
def get_data_with_valid_check(value="", datatypes: tuple[()] | tuple[Any] | Any = (), invalid_value="NA"):
    if isinstance(value, datatypes):
        return value
    return invalid_value


def get_clipboard_data(value: Any = "", clipboard_function: Callable | None = None, **kwargs: dict) -> Any:
    if clipboard_function is not None:
        return clipboard_function(value, **kwargs)
    if isinstance(value, (str, int, float, bool)):
        return value
    return data_to_str(value, **kwargs)


class Formatter:
    def __init__(
        self,
        value: Any,
        datatypes: tuple[Any],
        format_function: Callable = to_int,
        to_str_function: Callable = to_str,
        nullable: bool = True,
        invalid_value: str = "NaN",
        pre_format_function: Callable | None = None,
        post_format_function: Callable | None = None,
        clipboard_function: Callable | None = None,
        **kwargs,
    ) -> None:
        if nullable:
            if isinstance(datatypes, (list, tuple)):
                datatypes = tuple(set(datatypes) | {type(None)})
            else:
                datatypes = (datatypes, type(None))
        elif isinstance(datatypes, (list, tuple)) and type(None) in datatypes or datatypes is type(None):
            raise TypeError("Non-nullable cells cannot have NoneType as a datatype.")
        self.kwargs = kwargs
        self.valid_datatypes = datatypes
        self.format_function = format_function
        self.to_str_function = to_str_function
        self.nullable = nullable
        self.invalid_value = invalid_value
        self.pre_format_function = pre_format_function
        self.post_format_function = post_format_function
        self.clipboard_function = clipboard_function
        try:
            self.value = self.format_data(value)
        except Exception:
            self.value = f"{value}"

    def __str__(self) -> Any:
        if not self.valid():
            return self.invalid_value
        if self.value is None and self.nullable:
            return ""
        return self.to_str_function(self.value, **self.kwargs)

    def valid(self, value: Any = None) -> bool:
        if value is None:
            value = self.value
        return isinstance(value, self.valid_datatypes)

    def format_data(self, value: Any) -> Any:
        if self.pre_format_function:
            value = self.pre_format_function(value)
        value = None if (self.nullable and is_none_like(value)) else self.format_function(value, **self.kwargs)
        if self.post_format_function and self.valid(value):
            value = self.post_format_function(value)
        return value

    def get_data_with_valid_check(self) -> Any:
        if self.valid():
            return self.value
        return self.invalid_value

    def get_clipboard_data(self) -> Any:
        if self.clipboard_function is not None:
            return self.clipboard_function(self.value, **self.kwargs)
        if isinstance(self.value, (int, float, bool)):
            return self.value
        return self.__str__()

    def __eq__(self, __value: Any) -> bool:
        # in case of custom formatter class
        # compare the values
        try:
            if hasattr(__value, "value"):
                return self.value == __value.value
        except Exception:
            pass
        # if comparing to a string, format the string and compare
        if isinstance(__value, str):
            try:
                return self.value == self.format_data(__value)
            except Exception:
                pass
        # if comparing to anything else, compare the values
        return self.value == __value
