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
    
def to_nullable_int(x: Any):
    return None if is_nonelike(x) else to_int(x)

def to_nullable_float(x: Any):
    return None if is_nonelike(x) else to_float(x)

def to_nullable_bool(b: Any):
    return None if is_nonelike(b) else to_bool(b)

def int_formatter_to_str(v: Any, 
                         **kwargs: dict) -> str:
    return f"{v}"

def float_formatter_to_str(v: Any,
                           **kwargs: dict) -> str:
    if 'decimals' in kwargs:
        if kwargs['decimals']:
            return f"{round(v, kwargs['decimals'])}"
        return f"{int(round(v, kwargs['decimals']))}"
    if v.isinteger():
        return f"{int(v)}"
    return f"{v}"

def percentage_formatter_to_str(v: Any,
                                **kwargs: dict) -> str:
    if 'decimals' in kwargs:
        if kwargs['decimals']:
            return f"{round(v * 100, kwargs['decimals'])}%"
        return f"{int(round(v * 100, kwargs['decimals']))}%"
    if v.isinteger():
        return f"{int(v) * 100}%"
    return f"{v * 100}%"

def bool_formatter_to_str(v: Any,
                          **kwargs: dict) -> str:
    return f"{v}"

def int_formatter_kwargs(datatypes = int,
                         format_func = to_int,
                         to_str_func = int_formatter_to_str,
                         invalid_value = "NaN",
                         **kwargs,
                         ):
    return {**dict(datatypes = datatypes,
                   format_func = format_func,
                   to_str_func = to_str_func,
                   invalid_value = invalid_value),
            **kwargs}

def float_formatter_kwargs(datatypes = float,
                           format_func = to_float,
                           to_str_func = float_formatter_to_str,
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

def percentage_formatter_kwargs(datatypes = float,
                                format_func = to_float,
                                to_str_func = percentage_formatter_to_str,
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

def bool_formatter_kwargs(datatypes = bool,
                          format_func = to_bool,
                          to_str_func = bool_formatter_to_str,
                          invalid_value = "NA",
                          **kwargs,
                          ):
    return {**dict(datatypes = datatypes,
                   format_func = format_func,
                   to_str_func = to_str_func,
                   invalid_value = invalid_value),
            **kwargs}


class Formatter:
    def __init__(self,
                 value,
                 datatypes = int,
                 format_func = to_int,
                 to_str_func = int_formatter_to_str,
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


class AbstractFormatterClass(ABC):
    '''
    The base class for all cell formatters. Subclasses must define a __str__ method which determines how the data will be displayed on the table.
    Parameters:
        value: The value, usaually a string which is passed to the class on initialisaton. 
        datatypes: The target datatype(s) of the class. Used to validate itself by the data() method.
        format_func: A function that allows formats between the input string value and the target datatype.
        invalid_value: The object used in place of data that cannot be format_funced to the target datatype(s)
        pre_format_func: A function run on the initial value BEFORE it reaches the format_func
        post_format_func: A function run on the value AFTER it reaches the format_func IF the value has been correctly _formated.

    Methods:
        valid: returns Bool. used to validate whether self.value has been correctly _formated
        data: Returns either the _formated value or invalid_value depending on whether self.value has been correctly _formated
        _format(value): Tried to _format value using self.format_func.
    '''
    def __init__(self, 
                 value, 
                 datatypes, 
                 format_func,
                 nullable = True,
                 invalid_value = "NA",
                 pre_format_func = None,
                 post_format_func = None,
                 ):
        if nullable:
            if isinstance(datatypes, (list, tuple)):
                datatypes = [t for t in datatypes]
                datatypes.append(type(None))
                datatypes = tuple(datatypes)
            else:
                datatypes = (datatypes, type(None))
        elif isinstance(datatypes, (list, tuple)) and type(None) in datatypes:
            raise TypeError("Non-nullable cells cannot have NoneType as a datatype!")
        elif datatypes is type(None):
            raise TypeError("Non-nullable cells cannot have NoneType as a datatype!")
        self.nullable = nullable
        self.format_func = format_func
        self.pre_format_func = pre_format_func
        self.post_format_func = post_format_func
        self.valid_datatypes = datatypes
        self.invalid_value = invalid_value
        try:
            self.value = self._format(value)
        except (ValueError, TypeError):
            self.value = f"{value}"

    @abstractmethod
    def __str__(self):
        if not self.valid():
            return self.invalid_value
        if self.value is None and self.nullable:
            return ""

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
        if self.nullable:
            value = None if is_nonelike(value) else self.format_func(value)
        else:
            value = self.format_func(value)
        if self.post_format_func and self.valid(value):
            value = self.post_format_func(value)
        return value
    
    def data(self):
        if self.valid():
            return self.value
        return self.invalid_value
    
    def clipboard(self):
        if isinstance(self.value, (int, float, bool)):
            return self.value
        return self.__str__()
    

class CellFormatter(AbstractFormatterClass):
    '''
    Cell formatter class:
    Parameters:
        value: The value, usaually a string which is passed to the class on initialisaton. 
        datatypes: The target datatype(s) of the class. Used to validate itself by the data() method.
        format_func: A function that allows formats between the input string value and the target datatype.
        to_str: Function that _formats between the datatype and a representative string
        invalid_value: The object used in place of data that cannot be format_funced to the target datatype(s). Can be any type with a __str__ method.
        is_nullable: whether the None values can be stored in addtion to the other datatypes.
        pre_format_func: A function run on the initial value BEFORE it reaches the format_func
        post_format_func: A function run on the value AFTER it reaches the format_func IF the value has been correctly _formated.
    
    '''

    def __init__(self, 
                 value,
                 datatypes: Union[Type, tuple[Type]],
                 format_func: Callable[[Any], Any], 
                 to_str: Union[Callable[[Any], str], None] = None,
                 nullable = True,
                 invalid_value: Any = "NA",
                 pre_format_func: Union[Callable, None] = None, 
                 post_format_func: Union[Callable, None] = None
                 ):
        super().__init__(value, 
                         datatypes, 
                         format_func, 
                         nullable, 
                         invalid_value, 
                         pre_format_func, 
                         post_format_func)
        self.to_str = to_str

    def __str__(self):
        s = super().__str__()
        if s is not None:
            return s
        if self.to_str is not None:
            return self.to_str(self.value)
        return f"{self.value}"


class FloatFormatter(AbstractFormatterClass):
    def __init__(self, 
                 value, 
                 decimals = None,
                 nullable = True,
                 invalid_value = "NaN",
                 format_func = to_float,
                 pre_format_func = None,
                 post_format_func = None,
                 ):
        self.decimals = decimals
        super().__init__(value, 
                         float, 
                         format_func,
                         nullable,
                         invalid_value,
                         pre_format_func,
                         post_format_func,
                         )

    def __str__(self):
        s = super().__str__()
        if s is not None:
            return s
        if self.decimals is not None:
            return f"{round(self.value, self.decimals)}"
        return f"{self.value}"
    

class PercentageFormatter(AbstractFormatterClass):
    def __init__(self, 
                 value, 
                 decimals = None,
                 nullable = True,
                 invalid_value = "NaN",
                 format_func = to_float,
                 pre_format_func = None,
                 post_format_func = None,
                 ):
        self.decimals = decimals
        super().__init__(value, 
                         float, 
                         format_func,
                         nullable,
                         invalid_value,
                         pre_format_func,
                         post_format_func,
                         )

    def __str__(self):
        s = super().__str__()
        if s is not None:
            return s
        if self.decimals is not None:
            return f"{(round(self.value, self.decimals)) * 100}%"
        return f"{self.value * 100}%"


class IntFormatter(AbstractFormatterClass):
    def __init__(self, 
                 value,
                 nullable = True,
                 invalid_value = "NaN",
                 format_func = to_int,
                 pre_format_func = None,
                 post_format_func = None,
                 ):
        super().__init__(value, 
                         int, 
                         format_func,
                         nullable,
                         invalid_value,
                         pre_format_func,
                         post_format_func,
                         )

    def __str__(self):
        s = super().__str__()
        if s is not None:
            return s
        return f"{self.value}"


class BoolFormatter(AbstractFormatterClass):
    def __init__(self, 
                 value,
                 invalid_value = "NA",
                 format_func = to_bool,
                 nullable = True,
                 pre_format_func = None,
                 post_format_func = None,                 
                 ):
        super().__init__(value, 
                         bool,
                         format_func,
                         nullable,
                         invalid_value,
                         pre_format_func,
                         post_format_func,
                         )
    
    def __str__(self):
        s = super().__str__()
        if s is not None:
            return s
        return f"{self.value}"
