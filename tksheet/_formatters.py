from datetime import datetime, date, timedelta, time
from dateutil import parser, tz
from abc import ABC, abstractmethod

def is_nonelike(n: str):
    nonelike = {'none', '-', ''}
    if n is None:
        return True
    if isinstance(n, str):
        return n.lower().replace(' ', '') in nonelike
    else:
        return False

def to_int(x: str):
    if isinstance(x, int):
        return x
    return int(float(x))

def to_float(x: str):
    if isinstance(x, float):
        return x
    if isinstance(x, str) and x.endswith('%'):
        try:
            return float(x.replace('%', ''))/100
        except:
            raise ValueError(f'Cannot map {x} to float')
    return float(x)

def to_date(d: str):
    if isinstance(d, date):
        return d
    elif isinstance(d, datetime):
        return d.date()
    if isinstance(d, str):
        try:
            return parser.parse(d).replace(tzinfo=None).date()
        except Exception as ex:
            raise ValueError(f'Cannot map "{d}" to date.')
    else:
        raise ValueError(f'Cannot map "{d}" to date.')

def to_datetime(dt: str, time_zone=None):
    if isinstance(dt, datetime):
        pass
    elif isinstance(dt, date):
        dt = datetime(dt.year, dt.month, dt.day)
    elif isinstance(dt, str):
        try:
            dt = parser.parse(dt)
        except Exception as ex:
            raise ValueError(f'Cannot map "{dt}" to datetime.')
    else:
        raise ValueError(f'Cannot map "{dt}" to datetime.')
    if dt.tzinfo is None:
        dt.replace(tzinfo=tz.tzlocal())
    dt = dt.astimezone(tz.tzlocal() if time_zone is None else time_zone)
    return dt.replace(tzinfo=None)

def to_utc_datetime(dt: str):
    return to_datetime(dt, tz.UTC)

def to_time(t: str):
    if isinstance(t, timedelta):
        return t
    elif isinstance(t, datetime):
        return timedelta(
            hours=t.hour, 
            minutes=t.minute, 
            seconds=t.second,
            microseconds=t.microsecond,
        )
    elif isinstance(t, str):
        try:
            return parser.parse(t).time()
        except Exception as ex:
            raise ValueError(f'Cannot map "{t}" to timedelta.')
    else:
        raise ValueError(f'Cannot map "{t}" to timedelta.')

def to_bool(val: str):
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
    
def to_nullable_int(x):
    return None if is_nonelike(x) else to_int(x)

def to_nullable_float(x):
    return None if is_nonelike(x) else to_float(x)

def to_nullable_date(t):
    return None if is_nonelike(t) else to_date(t)

def to_nullable_datetime(t, time_zone=None):
    return None if is_nonelike(t) else to_datetime(t, time_zone=time_zone)
    
def to_nullable_utc_datetime(t):
    return None if is_nonelike(t) else to_utc_datetime(t)

def to_nullable_time(t):
    return None if is_nonelike(t) else to_time(t)
    
def to_nullable_bool(b):
    return None if is_nonelike(b) else to_bool(b)

class AbstractCellClass(ABC):
    '''
    The base class for all cell formatters. Subclasses must define a __str__ method which determines how the data will be displayed on the table.
    Parameters:
        value: The value, usaually a string which is pass to the class on initialisaton. 
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
                 missing_values = "NA",
                 pre_conversion_func = None,
                 post_conversion_func = None,
                 ):
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
        pass

    def validator(self) -> bool:
        if isinstance(self.value, self.valid_datatypes):
            self.is_valid = True
        else:
            self.is_valid  = False
        return self.is_valid
    
    def convert(self, value):
        if self.pre_conversion_func:
            value = self.pre_conversion_func(value)
        value = self.converter(value)
        if self.post_conversion_func and self.validator():
             value = self.post_conversion_func(value)
        return value
    
    def data(self):
        if self.validator():
            return self.value
        return self.missing_values

class DateCell(AbstractCellClass):
    def __init__(self, 
                 value, 
                 format = "%m/%d/%Y",
                 missing_values = "NaT",
                 converter = to_date,
                 pre_conversion_func = None,
                 post_conversion_func = None,
                 ):
        self.format = format
        super().__init__(value, 
                         date, 
                         converter,
                         missing_values,
                         pre_conversion_func,
                         post_conversion_func,
                         )

    def __str__(self):
        if not self.validator():
            return self.missing_values
        return self.value.strftime(self.format)
    
class NullableDateCell(AbstractCellClass):
    def __init__(self, 
                 value, 
                 format = "%m/%d/%Y",
                 missing_values = "NaT",
                 converter = to_nullable_date,
                 pre_conversion_func = None,
                 post_conversion_func = None,
                 ):
        self.format = format
        super().__init__(value, 
                         (date, type(None)), 
                         converter,
                         missing_values,
                         pre_conversion_func,
                         post_conversion_func,
                         )

    def __str__(self):
        if not self.validator():
            return self.missing_values   
        elif isinstance(self.value, date):
            return self.value.strftime(self.format)
        return ""
        
class NullableDatetimeCell(AbstractCellClass):
    def __init__(self, 
                 value, 
                 format = "%d/%m/%Y %H:%M:%S",
                 missing_values = "NaT",
                 converter = to_nullable_datetime,
                 pre_conversion_func = None,
                 post_conversion_func = None,
                 ):
        self.format = format
        super().__init__(value, 
                         (datetime, type(None)), 
                         converter,
                         missing_values,
                         pre_conversion_func,
                         post_conversion_func,
                         )

    def __str__(self):
        if not self.validator():
            return self.missing_values  
        elif isinstance(self.value, datetime):
            return self.value.strftime(self.format)
        return ""

class DatetimeCell(AbstractCellClass):
    def __init__(self, 
                 value, 
                 format = "%d/%m/%Y %H:%M:%S",
                 missing_values = "NaT",
                 converter = to_datetime,
                 pre_conversion_func = None,
                 post_conversion_func = None,
                 ):
        self.format = format
        super().__init__(value, 
                         datetime, 
                         converter,
                         missing_values,
                         pre_conversion_func,
                         post_conversion_func,
                         )

    def __str__(self):
        if not self.validator():
            return self.missing_values  
        return self.value.strftime(self.format)

class NullableTimeCell(AbstractCellClass):
    def __init__(self, 
                 value, 
                 format = "%H:%M:%S",
                 missing_values = "NaT",
                 converter = to_nullable_time,
                 pre_conversion_func = None,
                 post_conversion_func = None,
                 ):
        self.format = format
        super().__init__(value, 
                         (time, type(None)), 
                         converter,
                         missing_values,
                         pre_conversion_func,
                         post_conversion_func,
                         )

    def __str__(self):
        if not self.validator():
            return self.missing_values  
        elif isinstance(self.value, datetime):
            return self.value.strftime(self.format)
        return ""

class TimeCell(AbstractCellClass):
    def __init__(self, 
                 value, 
                 format = "%H:%M:%S",
                 missing_values = "NaT",
                 converter = to_time,
                 pre_conversion_func = None,
                 post_conversion_func = None,
                 ):
        self.format = format
        super().__init__(value, 
                         time, 
                         converter,
                         missing_values,
                         pre_conversion_func,
                         post_conversion_func,
                         )

    def __str__(self):
        if not self.validator():
            return self.missing_values  
        return self.value.strftime(self.format)

class NullableFloatCell(AbstractCellClass):
    def __init__(self, 
                 value, 
                 decimals = None,
                 missing_values = "NaN",
                 converter = to_nullable_float,
                 pre_conversion_func = None,
                 post_conversion_func = None,
                 ):
        self.decimals = decimals
        super().__init__(value, 
                         (float, type(None)), 
                         converter,
                         missing_values,
                         pre_conversion_func,
                         post_conversion_func,
                         )

    def __str__(self):
        if not self.validator():
            return self.missing_values
        if isinstance(self.value, float):
            if self.decimals != None:
                return (f"%.{self.decimals}f" % round(self.value, self.decimals))
            return str(self.value)
        return ""

class FloatCell(AbstractCellClass):
    def __init__(self, 
                 value, 
                 decimals = None,
                 missing_values = "NaN",
                 converter = to_float,
                 pre_conversion_func = None,
                 post_conversion_func = None,
                 ):
        self.decimals = decimals
        super().__init__(value, 
                         float, 
                         converter,
                         missing_values,
                         pre_conversion_func,
                         post_conversion_func,
                         )

    def __str__(self):
        if not self.validator():
            return self.missing_values
        if self.decimals != None:
            return (f"%.{self.decimals}f" % round(self.value, self.decimals))
        return str(self.value)
    

class PercentageCell(AbstractCellClass):
    def __init__(self, 
                 value, 
                 decimals = None,
                 missing_values = "NaN",
                 converter = to_float,
                 pre_conversion_func = None,
                 post_conversion_func = None,
                 ):
        self.decimals = decimals
        super().__init__(value, 
                         float, 
                         converter,
                         missing_values,
                         pre_conversion_func,
                         post_conversion_func,
                         )

    def __str__(self):
        if not self.validator():
            return self.missing_values
        if self.decimals != None:
            return (f"%.{self.decimals}f" % round(self.value*100, self.decimals))+'%'
        return str(self.value*100)+'%'

class NullablePercentageCell(AbstractCellClass):
    def __init__(self, 
                 value, 
                 decimals = None,
                 missing_values = "NaN",
                 converter = to_nullable_float,
                 pre_conversion_func = None,
                 post_conversion_func = None,
                 ):
        self.decimals = decimals
        super().__init__(value, 
                         (float, type(None)), 
                         converter,
                         missing_values,
                         pre_conversion_func,
                         post_conversion_func,
                         )

    def __str__(self):
        if not self.validator():
            return self.missing_values
        if isinstance(self.value, float):
            if self.decimals != None:
                return (f"%.{self.decimals}f" % round(self.value*100, self.decimals))+'%'
            return str(self.value*100)+'%'
        return ""

class NullableIntCell(AbstractCellClass):
    def __init__(self, 
                 value,
                 missing_values = "NaN",
                 converter = to_nullable_int,
                 pre_conversion_func = None,
                 post_conversion_func = None,
                 ):
        super().__init__(value, 
                         (int, type(None)), 
                         converter,
                         missing_values,
                         pre_conversion_func,
                         post_conversion_func,
                         )

    def __str__(self):
        if not self.validator():
            return self.missing_values
        if isinstance(self.value, int):
            return str(self.value)
        return ""
    
class IntCell(AbstractCellClass):
    def __init__(self, 
                 value,
                 missing_values = "NaN",
                 converter = to_int,
                 pre_conversion_func = None,
                 post_conversion_func = None,
                 ):
        super().__init__(value, 
                         int, 
                         converter,
                         missing_values,
                         pre_conversion_func,
                         post_conversion_func,
                         )

    def __str__(self):
        if not self.validator():
            return self.missing_values
        return str(self.value)
    
class NullableBoolCell(AbstractCellClass):
    def __init__(self, 
                 value,
                 missing_values = "NA",
                 converter = to_nullable_bool,
                 pre_conversion_func = None,
                 post_conversion_func = None,                 
                 ):
        super().__init__(value, 
                         (bool, type(None)),
                         converter,
                         missing_values,
                         pre_conversion_func,
                         post_conversion_func,
                         )
    
    def __str__(self):
        if not self.validator():
            return self.missing_values
        if isinstance(self.value, bool):
            return str(self.value)
        return ""

class BoolCell(AbstractCellClass):
    def __init__(self, 
                 value,
                 missing_values = "NA",
                 converter = to_bool,
                 pre_conversion_func = None,
                 post_conversion_func = None,                 
                 ):
        super().__init__(value, 
                         bool,
                         converter,
                         missing_values,
                         pre_conversion_func,
                         post_conversion_func,
                         )
    
    def __str__(self):
        if not self.validator():
            return self.missing_values
        return str(self.value)
    
