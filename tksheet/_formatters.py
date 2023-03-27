from datetime import datetime, date, timedelta
from dateutil import parser, tz
from abc import ABC, abstractmethod

def is_nonelike(n: str):
    nonelike = {'none', 'na', 'n/a', 'nan', 'null', 'nada', 'nat', 'nil', '-', '', 'nad'}
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
    def __init__(self, value, datatypes):
        try:
            self.value = self.converter(value)
        except ValueError:
            self.value = str(value)
        self._datatypes = datatypes
        self.validator()

    def validator(self) -> bool:
        if isinstance(self.value, self._datatypes):
            self._is_valid = True
        else:
            self._is_valid  = False
        return self._is_valid
    
    @abstractmethod
    def converter(self, value):
        pass

class DateCell(AbstractCellClass):
    def __init__(self, value, format = "%m/%d/%Y"):
        self.format = format
        super().__init__(value, date)

    def __str__(self):
        if not self.validator():
            return "NAT"
        return self.value.strftime(self.format)
        
    def converter(self, value):
        return to_date(value)
    
class NullableDateCell(AbstractCellClass):
    def __init__(self, value, format = "%m/%d/%Y"):
        self.format = format
        super().__init__(value, (date, type(None)))

    def __str__(self):
        if not self.validator():
            return "NAT"    
        elif isinstance(self.value, date):
            return self.value.strftime(self.format)
        return ""
        
    def converter(self, value):
        return to_nullable_date(value)

class NullableDatetimeCell(AbstractCellClass):
    def __init__(self, value, format = "%m/%d/%Y %H:%M:%S"):
        self.format = format
        super().__init__(value, datatypes=(datetime, type(None)))

    def __str__(self):
        if not self.validator():
            return "NAT"    
        elif isinstance(self.value, datetime):
            return self.value.strftime(self.format)
        return ""
            
    def converter(self, value):
        return to_nullable_datetime(value)

class DatetimeCell(AbstractCellClass):
    def __init__(self, value, format = "%m/%d/%Y %H:%M:%S"):
        self.format = format
        super().__init__(value, datatypes=datetime)

    def __str__(self):
        if not self.validator():
            return "NAT"   
        return self.value.strftime(self.format)
            
    def converter(self, value):
        return to_datetime(value)
    
class NullableFloatCell(AbstractCellClass):
    def __init__(self, value, decimal_places = None):
        self._decimal_places = decimal_places
        super().__init__(value, datatypes=(float, type(None)))

    def __str__(self):
        if not self.validator():
            return "NA"
        if isinstance(self.value, float):
            return (f"%.{self._decimal_places}f" % self.value)
        return ""
            
    def converter(self, value):
        value = to_nullable_float(value)
        if self._decimal_places and isinstance(value, float):
            return round(value, self._decimal_places)
        return value

class FloatCell(AbstractCellClass):
    def __init__(self, value, decimal_places = None):
        self._decimal_places = decimal_places
        super().__init__(value, datatypes=float)

    def __str__(self):
        if not self.validator():
            return "NA"
        return (f"%.{self._decimal_places}f" % self.value)
            
    def converter(self, value):
        value = to_float(value)
        if self._decimal_places and isinstance(value, float):
            value = round(value, self._decimal_places)
        return value
    
class NullableIntCell(AbstractCellClass):
    def __init__(self, value):
        super().__init__(value, datatypes=(int, type(None)))

    def __str__(self):
        if not self.validator():
            return "NA"
        if isinstance(self.value, int):
            return str(self.value)
        return ""
    
    def converter(self, value):
        return to_nullable_int(value)
    
class IntCell(AbstractCellClass):
    def __init__(self, value):
        super().__init__(value, datatypes=int)

    def __str__(self):
        if not self.validator():
            return "NA"
        return str(self.value)
    
    def converter(self, value):
        return to_int(value)
    
class NullableBoolCell(AbstractCellClass):
    def __init__(self, value):
        super().__init__(value, datatypes=(bool, type(None)))
    
    def __str__(self):
        if not self.validator():
            return "NA"
        if isinstance(self.value, bool):
            return str(self.value)
        return ""
    
    def converter(self, value):
        return to_nullable_bool(value)

class BoolCell(AbstractCellClass):
    def __init__(self, value):
        super().__init__(value, datatypes=bool)
    
    def __str__(self):
        if not self.validator():
            return "NA"
        return str(self.value)
    
    def converter(self, value):
        return to_bool(value)