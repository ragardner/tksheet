from datetime import datetime, date, timedelta
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

def to_timedelta(t: str):
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
    return None if is_nonelike(x) else float(x)

def to_nullable_date(t):
    return None if is_nonelike(t) else to_date(t)

def to_nullable_datetime(t, time_zone=None):
    return None if is_nonelike(t) else to_datetime(t, time_zone=time_zone)
    
def to_nullable_utc_datetime(t):
    return None if is_nonelike(t) else to_utc_datetime(t)

def to_nullable_timedelta(t):
    return None if is_nonelike(t) else to_timedelta(t)
    
def to_nullable_bool(b):
    return None if is_nonelike(b) else to_bool(b)

class AbstractCellClass(ABC):
    def __init__(self, value, datatypes, missing_value = "NA"):
        try:
            self.value = self.converter(value)
        except ValueError:
            self.value = str(value)
        self.valid_datatypes = datatypes
        self.validator()
        self.missing_value = missing_value

    def validator(self) -> bool:
        if isinstance(self.value, self.valid_datatypes):
            self.is_valid = True
        else:
            self.is_valid  = False
        return self.is_valid
    
    @abstractmethod
    def converter(self, value):
        pass

    def data(self):
        if self.validator:
            return self.value
        return self.missing_value

class DateCell(AbstractCellClass):
    def __init__(self, value, format = "%m/%d/%Y"):
        self.format = format
        super().__init__(value, date, 'NaT')

    def __str__(self):
        if not self.validator():
            return self.missing_value
        return self.value.strftime(self.format)
        
    def converter(self, value):
        return to_date(value)
    
class NullableDateCell(AbstractCellClass):
    def __init__(self, value, format = "%m/%d/%Y"):
        self.format = format
        super().__init__(value, (date, type(None)), 'NaT')

    def __str__(self):
        if not self.validator():
            return self.missing_value   
        elif isinstance(self.value, date):
            return self.value.strftime(self.format)
        return ""
        
    def converter(self, value):
        return to_nullable_date(value)

class NullableDatetimeCell(AbstractCellClass):
    def __init__(self, value, format = "%m/%d/%Y %H:%M:%S"):
        self.format = format
        super().__init__(value, (datetime, type(None)), 'Nat')

    def __str__(self):
        if not self.validator():
            return self.missing_value  
        elif isinstance(self.value, datetime):
            return self.value.strftime(self.format)
        return ""
            
    def converter(self, value):
        return to_nullable_datetime(value)

class DatetimeCell(AbstractCellClass):
    def __init__(self, value, format = "%m/%d/%Y %H:%M:%S"):
        self.format = format
        super().__init__(value, datetime, 'NaT')

    def __str__(self):
        if not self.validator():
            return self.missing_value  
        return self.value.strftime(self.format)
            
    def converter(self, value):
        return to_datetime(value)
    
class NullableFloatCell(AbstractCellClass):
    def __init__(self, value, decimals = None):
        self.decimals = decimals
        super().__init__(value, (float, type(None)), 'NaN')

    def __str__(self):
        if not self.validator():
            return self.missing_value
        if isinstance(self.value, float):
            return (f"%.{self.decimals}f" % round(self.value, self.decimals))
        return ""
            
    def converter(self, value):
        return to_nullable_float(value)

class FloatCell(AbstractCellClass):
    def __init__(self, value, decimals = None):
        self.decimals = decimals
        super().__init__(value, float, 'NaN')

    def __str__(self):
        if not self.validator():
            return self.missing_value
        return (f"%.{self.decimals}f" % round(self.value, self.decimals))
            
    def converter(self, value):
        return to_float(value)
    
class NullableIntCell(AbstractCellClass):
    def __init__(self, value):
        super().__init__(value, (int, type(None)), 'NaN')

    def __str__(self):
        if not self.validator():
            return self.missing_value
        if isinstance(self.value, int):
            return str(self.value)
        return ""
    
    def converter(self, value):
        return to_nullable_int(value)
    
class IntCell(AbstractCellClass):
    def __init__(self, value):
        super().__init__(value, int, 'NaN')

    def __str__(self):
        if not self.validator():
            return self.missing_value
        return str(self.value)
    
    def converter(self, value):
        return to_int(value)
    
class NullableBoolCell(AbstractCellClass):
    def __init__(self, value):
        super().__init__(value, (bool, type(None)))
    
    def __str__(self):
        if not self.validator():
            return self.missing_value
        if isinstance(self.value, bool):
            return str(self.value)
        return ""
    
    def converter(self, value):
        return to_nullable_bool(value)

class BoolCell(AbstractCellClass):
    def __init__(self, value):
        super().__init__(value, bool)
    
    def __str__(self):
        if not self.validator():
            return self.missing_value
        return str(self.value)
    
    def converter(self, value):
        return to_bool(value)