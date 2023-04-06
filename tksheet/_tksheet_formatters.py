from abc import ABC, abstractmethod
from typing import Union, Any, Type, Callable

def is_nonelike(n: Any):
    nonelike = {'none', ''}
    if n is None:
        return True
    if isinstance(n, str):
        return n.lower().replace(' ', '') in nonelike
    else:
        return False

def to_int(x: Any):
    if isinstance(x, int):
        return x
    return int(float(x))

def to_float(x: Any):
    if isinstance(x, float):
        return x
    if isinstance(x, str) and x.endswith('%'):
        try:
            return float(x.replace('%', ''))/100
        except:
            raise ValueError(f'Cannot map {x} to float')
    return float(x)

def to_bool(val: Any):
    if isinstance(val, bool):
        return val
    v = val.lower()
    truthy = {True, "true", "t", "yes", "y", "on", "1", 1}
    falsy = {False, "false", "f", "no", "n", "off", "0", 0}
    if v in truthy:
        return True
    if v in falsy:
        return False
    else:
        raise ValueError(f'Cannot map "{val}" to bool.')
    
def to_nullable_int(x: Any):
    return None if is_nonelike(x) else to_int(x)

def to_nullable_float(x: Any):
    return None if is_nonelike(x) else to_float(x)


def to_nullable_bool(b: Any):
    return None if is_nonelike(b) else to_bool(b)


class AbstractFormatterClass(ABC):
    '''
    The base class for all cell formatters. Subclasses must define a __str__ method which determines how the data will be displayed on the table.
    Parameters:
        value: The value, usaually a string which is passed to the class on initialisaton. 
        datatypes: The target datatype(s) of the class. Used to validate itself by the data() method.
        converter: A function that allows conversions between the input string value and the target datatype.
        missing_values: The object used in place of data that cannot be convertered to the target datatype(s)
        pre_conversion_func: A function run on the initial value BEFORE it reaches the converter
        post_conversion_func: A function run on the value AFTER it reaches the converter IF the value has been correctly converted.

    Methods:
        validator: returns Bool. used to validate whether self.value has been correctly converted
        data: Returns either the converted value or missing_values depending on whether self.value has been correctly converted
        convert(value): Tried to convert value using self.converter.
    '''
    def __init__(self, 
                 value, 
                 datatypes, 
                 converter,
                 nullable = True,
                 missing_values = "NA",
                 pre_conversion_func = None,
                 post_conversion_func = None,
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
        self.converter = converter
        self.pre_conversion_func = pre_conversion_func
        self.post_conversion_func = post_conversion_func
        self.valid_datatypes = datatypes
        self.missing_values = missing_values
        try:
            self.value = self.convert(value)
        except (ValueError, TypeError):
            self.value = str(value)

    @abstractmethod
    def __str__(self):
        if not self.validator():
            return self.missing_values
        if self.value == None and self.nullable:
            return ''

    def validator(self, value = None) -> bool:
        if value is None:
            value = self.value
        if isinstance(value, self.valid_datatypes):
            return True
        else:
            return False
    
    def convert(self, value):
        if self.pre_conversion_func:
            value = self.pre_conversion_func(value)
        if self.nullable:
            value = None if is_nonelike(value) else self.converter(value)
        else:
            value = self.converter(value)
        if self.post_conversion_func and self.validator(value):
             value = self.post_conversion_func(value)
        return value
    
    def data(self):
        if self.validator():
            return self.value
        return self.missing_values
    

class CellFormatter(AbstractFormatterClass):
    '''
    Cell formatter class:
    Parameters:
        value: The value, usaually a string which is passed to the class on initialisaton. 
        datatypes: The target datatype(s) of the class. Used to validate itself by the data() method.
        converter: A function that allows conversions between the input string value and the target datatype.
        to_str: Function that converts between the datatype and a representative string
        missing_values: The object used in place of data that cannot be convertered to the target datatype(s). Can be any type with a __str__ method.
        is_nullable: whether the None values can be stored in addtion to the other datatypes.
        pre_conversion_func: A function run on the initial value BEFORE it reaches the converter
        post_conversion_func: A function run on the value AFTER it reaches the converter IF the value has been correctly converted.
    
    '''

    def __init__(self, 
                 value,
                 datatypes: Union[Type, tuple[Type]],
                 converter: Callable[[Any], Any], 
                 to_str: Union[Callable[[Any], str], None] = None,
                 nullable = True,
                 missing_values: Any = "NA",
                 pre_conversion_func: Union[Callable, None] = None, 
                 post_conversion_func: Union[Callable, None] = None
                 ):
        super().__init__(value, 
                         datatypes, 
                         converter, 
                         nullable, 
                         missing_values, 
                         pre_conversion_func, 
                         post_conversion_func)
        self.to_str = to_str

    def __str__(self):
        s = super().__str__()
        if s != None:
            return s
        if self.to_str != None:
            return self.to_str(self.value)
        return f'{self.value}'
    

class FloatFormatter(AbstractFormatterClass):
    def __init__(self, 
                 value, 
                 decimals = None,
                 nullable = True,
                 missing_values = "NaN",
                 converter = to_float,
                 pre_conversion_func = None,
                 post_conversion_func = None,
                 ):
        self.decimals = decimals
        super().__init__(value, 
                         float, 
                         converter,
                         nullable,
                         missing_values,
                         pre_conversion_func,
                         post_conversion_func,
                         )

    def __str__(self):
        s = super().__str__()
        if s != None:
            return s
        if self.decimals != None:
            return (f"%.{self.decimals}f" % round(self.value, self.decimals))
        return str(self.value)
    

class PercentageFormatter(AbstractFormatterClass):
    def __init__(self, 
                 value, 
                 decimals = None,
                 nullable = True,
                 missing_values = "NaN",
                 converter = to_float,
                 pre_conversion_func = None,
                 post_conversion_func = None,
                 ):
        self.decimals = decimals
        super().__init__(value, 
                         float, 
                         converter,
                         nullable,
                         missing_values,
                         pre_conversion_func,
                         post_conversion_func,
                         )

    def __str__(self):
        s = super().__str__()
        if s != None:
            return s
        if self.decimals != None:
            return (f"%.{self.decimals}f" % round(self.value*100, self.decimals))+'%'
        return str(self.value*100)+'%'


class IntFormatter(AbstractFormatterClass):
    def __init__(self, 
                 value,
                 nullable = True,
                 missing_values = "NaN",
                 converter = to_int,
                 pre_conversion_func = None,
                 post_conversion_func = None,
                 ):
        super().__init__(value, 
                         int, 
                         converter,
                         nullable,
                         missing_values,
                         pre_conversion_func,
                         post_conversion_func,
                         )

    def __str__(self):
        s = super().__str__()
        if s != None:
            return s
        return str(self.value)


class BoolFormatter(AbstractFormatterClass):
    def __init__(self, 
                 value,
                 missing_values = "NA",
                 converter = to_bool,
                 nullable = True,
                 pre_conversion_func = None,
                 post_conversion_func = None,                 
                 ):
        super().__init__(value, 
                         bool,
                         converter,
                         nullable,
                         missing_values,
                         pre_conversion_func,
                         post_conversion_func,
                         )
    
    def __str__(self):
        s = super().__str__()
        if s != None:
            return s
        return str(self.value)
    
