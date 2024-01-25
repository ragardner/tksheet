from __future__ import annotations

from collections.abc import Callable

from .vars import falsy, nonelike, truthy


def is_none_like(o: object):
    if (isinstance(o, str) and o.lower().replace(" ", "") in nonelike) or o in nonelike:
        return True
    return False


def to_int(o: object, **kwargs):
    if isinstance(o, int):
        return o
    return int(float(o))


def to_float(o: object, **kwargs):
    if isinstance(o, float):
        return o
    if isinstance(o, str) and o.endswith("%"):
        return float(o.replace("%", "")) / 100
    return float(o)


def to_bool(val: object, **kwargs):
    if isinstance(val, bool):
        return val
    if isinstance(val, str):
        v = val.lower()
    else:
        v = val
    if "truthy" in kwargs:
        _truthy = kwargs["truthy"]
    else:
        _truthy = truthy
    if "falsy" in kwargs:
        _falsy = kwargs["falsy"]
    else:
        _falsy = falsy
    if v in _truthy:
        return True
    elif v in _falsy:
        return False
    raise ValueError(f'Cannot map "{val}" to bool.')


def try_to_bool(o: object, **kwargs):
    try:
        return to_bool(o)
    except Exception:
        return o


def is_bool_like(o: object, **kwargs):
    try:
        to_bool(o)
        return True
    except Exception:
        return False


def to_str(o: object, **kwargs: dict) -> str:
    return f"{o}"


def float_to_str(v: int | float, **kwargs: dict) -> str:
    if isinstance(v, float):
        if v.is_integer():
            return f"{int(v)}"
        if "decimals" in kwargs and isinstance(kwargs["decimals"], int):
            if kwargs["decimals"]:
                return f"{round(v, kwargs['decimals'])}"
            return f"{int(round(v, kwargs['decimals']))}"
    return f"{v}"


def percentage_to_str(v: int | float, **kwargs: dict) -> str:
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


def bool_to_str(v: object, **kwargs: dict) -> str:
    return f"{v}"


def int_formatter(
    datatypes: tuple[object] | object = int,
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
    datatypes: tuple[object] | object = float,
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
    datatypes: tuple[object] | object = float,
    format_function: Callable = to_float,
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
    datatypes: tuple[object] | object = bool,
    format_function: Callable = to_bool,
    to_str_function: Callable = bool_to_str,
    invalid_value: object = "NA",
    truthy_values: set[object] = truthy,
    falsy_values: set[object] = falsy,
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
    datatypes: tuple[object] | object,
    format_function: Callable,
    to_str_function: Callable = to_str,
    invalid_value: object = "NaN",
    nullable: bool = True,
    pre_format_function: Callable | None = None,
    post_format_function: Callable | None = None,
    clipboard_function: Callable | None = None,
    **kwargs,
) -> dict:
    return {
        **dict(
            datatypes=datatypes,
            format_function=format_function,
            to_str_function=to_str_function,
            invalid_value=invalid_value,
            nullable=nullable,
            pre_format_function=pre_format_function,
            post_format_function=post_format_function,
            clipboard_function=clipboard_function,
        ),
        **kwargs,
    }


def format_data(
    value: object = "",
    datatypes: tuple[object] | object = int,
    nullable: bool = True,
    pre_format_function: Callable | None = None,
    format_function: Callable | None = to_int,
    post_format_function: Callable | None = None,
    **kwargs,
) -> object:
    if pre_format_function:
        value = pre_format_function(value)
    if nullable and is_none_like(value):
        value = None
    else:
        try:
            value = format_function(value, **kwargs)
        except Exception:
            pass
    if post_format_function and isinstance(value, datatypes):
        value = post_format_function(value)
    return value


def data_to_str(
    value: object = "",
    datatypes: tuple[object] | object = int,
    nullable: bool = True,
    invalid_value: object = "NaN",
    to_str_function: Callable | None = None,
    **kwargs,
) -> str:
    if not isinstance(value, datatypes):
        return invalid_value
    if value is None and nullable:
        return ""
    return to_str_function(value, **kwargs)


def get_data_with_valid_check(value="", datatypes: tuple[()] | tuple[object] | object = tuple(), invalid_value="NA"):
    if isinstance(value, datatypes):
        return value
    return invalid_value


def get_clipboard_data(value: object = "", clipboard_function: Callable | None = None, **kwargs: dict) -> object:
    if clipboard_function is not None:
        return clipboard_function(value, **kwargs)
    if isinstance(value, (str, int, float, bool)):
        return value
    return data_to_str(value, **kwargs)


class Formatter:
    def __init__(
        self,
        value: object,
        datatypes: tuple[object],
        object=int,
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
                datatypes = tuple({type_ for type_ in datatypes} | {type(None)})
            else:
                datatypes = (datatypes, type(None))
        elif isinstance(datatypes, (list, tuple)) and type(None) in datatypes:
            raise TypeError("Non-nullable cells cannot have NoneType as a datatype.")
        elif datatypes is type(None):
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

    def __str__(self) -> object:
        if not self.valid():
            return self.invalid_value
        if self.value is None and self.nullable:
            return ""
        return self.to_str_function(self.value, **self.kwargs)

    def valid(self, value: object = None) -> bool:
        if value is None:
            value = self.value
        if isinstance(value, self.valid_datatypes):
            return True
        return False

    def format_data(self, value: object) -> object:
        if self.pre_format_function:
            value = self.pre_format_function(value)
        value = None if (self.nullable and is_none_like(value)) else self.format_function(value, **self.kwargs)
        if self.post_format_function and self.valid(value):
            value = self.post_format_function(value)
        return value

    def get_data_with_valid_check(self) -> object:
        if self.valid():
            return self.value
        return self.invalid_value

    def get_clipboard_data(self) -> object:
        if self.clipboard_function is not None:
            return self.clipboard_function(self.value, **self.kwargs)
        if isinstance(self.value, (int, float, bool)):
            return self.value
        return self.__str__()

    def __eq__(self, __value: object) -> bool:
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
