from abc import ABC, abstractmethod
from typing import Union, Any, Type, Callable
from ._tksheet_vars import *

def is_nonelike(n: Any):
    if n is None:
        return True
    if isinstance(n, str):
        return n.lower().replace(" ", "") in nonelike
    else:
        return False

def to_int(x: Any, **kwargs):
    if isinstance(x, int):
        return x
    return int(float(x))

def to_float(x: Any, **kwargs):
    if isinstance(x, float):
        return x
    if isinstance(x, str) and x.endswith('%'):
        try:
            return float(x.replace('%', "")) / 100
        except:
            raise ValueError(f'Cannot map {x} to float')
    return float(x)

def to_bool(val: Any, **kwargs):
    if isinstance(val, bool):
        return val
    if isinstance(val, str):
        v = val.lower()
    else:
        v = val
    if v in truthy:
        return True
    elif v in falsy:
        return False
    raise ValueError(f'Cannot map "{val}" to bool.')

def to_str(v: Any, 
                         **kwargs: dict) -> str:
    return f"{v}"

def float_to_str(v: Any,
                           **kwargs: dict) -> str:
    if 'decimals' in kwargs:
        if kwargs['decimals']:
            return f"{round(v, kwargs['decimals'])}"
        return f"{int(round(v, kwargs['decimals']))}"
    if v.isinteger():
        return f"{int(v)}"
    return f"{v}"

def percentage_to_str(v: Any,
                                **kwargs: dict) -> str:
    if 'decimals' in kwargs:
        if kwargs['decimals']:
            return f"{round(v * 100, kwargs['decimals'])}%"
        return f"{int(round(v * 100, kwargs['decimals']))}%"
    if v.isinteger():
        return f"{int(v) * 100}%"
    return f"{v * 100}%"

def bool_to_str(v: Any,
                          **kwargs: dict) -> str:
    return f"{v}"

def int_formatter(datatypes = int,
                         format_func = to_int,
                         to_str_func = to_str,
                         invalid_value = "NaN",
                         **kwargs,
                         ):
    return {**dict(datatypes = datatypes,
                   format_func = format_func,
                   to_str_func = to_str_func,
                   invalid_value = invalid_value),
            **kwargs}

def float_formatter(datatypes = float,
                           format_func = to_float,
                           to_str_func = float_to_str,
                           invalid_value = "NaN",
                           decimals = 1,
                           **kwargs,
                           ):
    return {**dict(datatypes = datatypes,
                   format_func = format_func,
                   to_str_func = to_str_func,
                   invalid_value = invalid_value,
                   decimals = decimals),
            **kwargs}

def percentage_formatter(datatypes = float,
                                format_func = to_float,
                                to_str_func = percentage_to_str,
                                invalid_value = "NaN",
                                decimals = 0,
                                **kwargs,
                                ):
    return {**dict(datatypes = datatypes,
                   format_func = format_func,
                   to_str_func = to_str_func,
                   invalid_value = invalid_value,
                   decimals = decimals),
            **kwargs}

def bool_formatter(datatypes = bool,
                          format_func = to_bool,
                          to_str_func = bool_to_str,
                          invalid_value = "NA",
                          **kwargs,
                          ):
    return {**dict(datatypes = datatypes,
                   format_func = format_func,
                   to_str_func = to_str_func,
                   invalid_value = invalid_value),
            **kwargs}

def formatter(datatypes,
              format_func,
              to_str_func = to_str,
              invalid_value = "NaN",
              nullable = True,
              pre_format_func = None,
              post_format_func = None,
              clipboard_func = None,
              **kwargs):
    return {**dict(datatypes = datatypes,
                   format_func = format_func,
                   to_str_func = to_str_func,
                   invalid_value = invalid_value,
                   nullable = nullable,
                   pre_format_func = pre_format_func,
                   post_format_func = post_format_func,
                   clipboard_func = clipboard_func),
            **kwargs}


class Formatter:
    def __init__(self,
                 value,
                 datatypes = int,
                 format_func = to_int,
                 to_str_func = to_str,
                 nullable = True,
                 invalid_value = "NaN",
                 pre_format_func = None,
                 post_format_func = None,
                 clipboard_func = None,
                 **kwargs,
                 ):
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
        self.format_func = format_func
        self.to_str_func = to_str_func
        self.nullable = nullable
        self.invalid_value = invalid_value
        self.pre_format_func = pre_format_func
        self.post_format_func = post_format_func
        self.clipboard_func = clipboard_func
        try:
            self.value = self._format(value)
        except Exception as e:
            self.value = f"{value}"

    def __str__(self):
        if not self.valid():
            return self.invalid_value
        if self.value is None and self.nullable:
            return ""
        return self.to_str_func(self.value, **self.kwargs)

    def valid(self, value = None) -> bool:
        if value is None:
            value = self.value
        if isinstance(value, self.valid_datatypes):
            return True
        else:
            return False
    
    def _format(self, value):
        if self.pre_format_func:
            value = self.pre_format_func(value)
        value = None if (self.nullable and is_nonelike(value)) else self.format_func(value, **self.kwargs)
        if self.post_format_func and self.valid(value):
            value = self.post_format_func(value)
        return value

    def data(self):
        if self.valid():
            return self.value
        return self.invalid_value
    
    def clipboard(self):
        if self.clipboard_func is not None:
            return self.clipboard_func(self.value, **self.kwargs)
        if isinstance(self.value, (int, float, bool)):
            return self.value
        return self.__str__()
    
    def __eq__(self, __value: object) -> bool:
        if isinstance(__value, Formatter): # if comparing to another formatter, compare the values
            return self.value == __value.value
        if isinstance(__value, str): # if comparing to a string, format the string and compare
            try:
                value = self._format(__value)
                return self.value == value
            except Exception as e:
                pass
        return f"{self.value}" == __value # if comparing to anything else, compare the values
