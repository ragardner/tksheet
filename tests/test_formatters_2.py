from tksheet import *
from tksheet._tksheet_formatters import *
import tkinter as tk
from datetime import datetime, date, timedelta, time
from dateutil import parser, tz
from math import ceil


# --------------------- Special formatter arguements ---------------------
def round_up(x):
    return float(ceil(x))

def only_numeric(s):
    return ''.join([n for n in s if n.isnumeric() or n == '.'])

def convert_to_local_datetime(dt: str):
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
    dt = dt.astimezone(tz.tzlocal())
    return dt.replace(tzinfo=None)

def datetime_to_string(dt: datetime):
    return dt.strftime('%d %b, %Y, %H:%M:%S')

# --------------------- Custom Formatter Class ---------------------
class CustomDatetimeFormatter(AbstractFormatterClass):
    def __init__(self, 
                 value, 
                 format = "%d/%m/%Y %H:%M:%S", #Additional arguement that can be passed through the formatter kwargs
                 nullable = True,
                 invalid_value = 'NaT',
                 format_func = convert_to_local_datetime,
                 pre_format_func = None,
                 post_format_func = None,
                 ):
        self.format = format
        super().__init__(value, 
                         datetime, 
                         format_func,
                         nullable,
                         invalid_value,
                         pre_format_func,
                         post_format_func,
                         )

    def __str__(self):
        if (s:=super().__str__()) != None:
            return s
        return self.value.strftime(self.format)
    

class demo(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)
        self.grid_columnconfigure(0, weight = 1)
        self.grid_rowconfigure(0, weight = 1)
        self.frame = tk.Frame(self)
        self.frame.grid_columnconfigure(0, weight = 1)
        self.frame.grid_rowconfigure(0, weight = 1)
        self.sheet = Sheet(self.frame,
                           empty_horizontal=0,
                           empty_vertical=0,
                           data = [[f"{r}"]*11 for r in range(20)]
                           )
        self.sheet.enable_bindings()
        self.frame.grid(row = 0, column = 0, sticky = "nswe")
        self.sheet.grid(row = 0, column = 0, sticky = "nswe")
        self.sheet.headers(['Non-Nullable Float Cell\n1 decimals places', 
                            'Float Cell', 
                            'Int Cell', 
                            'Bool Cell', 
                            'Percentage Cell\n0 decimal places', 
                            'Custom Datetime Cell',
                            'Custom Datetime Cell Class',
                            'Float Cell that\nrounds up', 
                            'Float cell that\n strips non-numeric', 
                            'Dropdown Over Nullable\nPercentage Cell', 
                            'Checkboxes always\noverride formatting'])
        
        # ---------- Some examples of cell formatting --------
        self.sheet.format_cell('all', 0, FloatFormatter, dict(decimals=1, nullable=False))  
        self.sheet.format_cell('all', 1, FloatFormatter, dict(decimals=5))
        self.sheet.format_cell('all', 2, IntFormatter)
        self.sheet.format_cell('all', 3, BoolFormatter)
        self.sheet.format_cell('all', 4, PercentageFormatter, dict(decimals=0))


        # ---------------- Custom Formatters -----------------
        # Totally custom cell format using the inbuilt cell formatter
        self.sheet.format_cell('all', 5, CellFormatter, datatypes = datetime, 
                                                              format_func = convert_to_local_datetime, 
                                                              to_str = datetime_to_string, 
                                                              nullable = False,
                                                              invalid_value = 'NaT')
        # Custom format generated from the base class
        self.sheet.format_cell('all', 6, CustomDatetimeFormatter, format = '%Y-%m-%d %H:%M %p', nullable = True)
        
        # Unique cell behaviour using the post_format_function
        self.sheet.format_cell('all', 7, FloatFormatter, dict(post_format_func = round_up)) # Custom display format
        self.sheet.format_cell('all', 8, FloatFormatter, dict(pre_format_func = only_numeric))

        self.sheet.create_dropdown('all', 9, values=['', '104%',.24,"300%",'not a number'], set_value=1)
        self.sheet.format_cell('all', 9, PercentageFormatter)
        self.sheet.format_cell('all', 10, PercentageFormatter)
        self.sheet.create_checkbox('all', 10, text='Checkbox') # Check boxes always override formatters

app = demo()
app.mainloop()