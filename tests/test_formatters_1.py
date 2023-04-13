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
    return ''.join(n for n in s if n.isnumeric() or n == '.')

def convert_to_local_datetime(dt: str, **kwargs):
    if isinstance(dt, datetime):
        pass
    elif isinstance(dt, date):
        dt = datetime(dt.year, dt.month, dt.day)
    try:
        dt = parser.parse(dt)
    except:
        return dt
    if dt.tzinfo is None:
        dt.replace(tzinfo = tz.tzlocal())
    dt = dt.astimezone(tz.tzlocal())
    return dt.replace(tzinfo = None)

def datetime_to_string(dt: datetime, **kwargs):
    return dt.strftime('%d %b, %Y, %H:%M:%S')

# --------------------- Custom Formatter Class ---------------------

def custom_datetime_to_str(dt: datetime, **kwargs):
    return dt.strftime(kwargs['format'])


class demo(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)
        self.grid_columnconfigure(0, weight = 1)
        self.grid_rowconfigure(0, weight = 1)
        self.frame = tk.Frame(self)
        self.frame.grid_columnconfigure(0, weight = 1)
        self.frame.grid_rowconfigure(0, weight = 1)
        self.sheet = Sheet(self.frame,
                           empty_vertical = 0,
                           empty_horizontal = 0,
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
                            'Percentage Cell\n2 decimal places'])
        
        # ---------- Some examples of cell formatting --------
        self.sheet.format_cell('all', 0, formatter_kwargs = float_formatter_kwargs(nullable = False))
        self.sheet.format_cell('all', 1, formatter_kwargs = float_formatter_kwargs())
        self.sheet.format_cell('all', 2, formatter_kwargs = int_formatter_kwargs())
        self.sheet.format_cell('all', 3, formatter_kwargs = bool_formatter_kwargs())
        self.sheet.format_cell('all', 4, formatter_kwargs = percentage_formatter_kwargs())


        # ---------------- Custom Formatters -----------------
        # Customformat
        self.sheet.format_cell('all', 5, datatypes = datetime, 
                                         format_func = convert_to_local_datetime, 
                                         to_str_func = datetime_to_string, 
                                         nullable = False,
                                         invalid_value = 'NaT',
                                         )
        # Custom format
        self.sheet.format_cell('all', 6, datatypes = datetime, 
                                         format_func = convert_to_local_datetime, 
                                         to_str_func = custom_datetime_to_str, 
                                         nullable = True,
                                         invalid_value = 'NaT',
                                         format = '%Y-%m-%d %H:%M %p'
                                         )
        
        # Unique cell behaviour using the post_conversion_function
        self.sheet.format_cell('all', 7, formatter_kwargs = float_formatter_kwargs(post_format_func = round_up))
        self.sheet.format_cell('all', 8, formatter_kwargs = float_formatter_kwargs(), pre_format_func = only_numeric)

        self.sheet.create_dropdown('all', 9, values = ['', '104%', .24, "300%", 'not a number'], set_value = 1)
        self.sheet.format_cell('all', 9, formatter_kwargs = percentage_formatter_kwargs(), decimals = 0)
        self.sheet.format_cell('all', 10, formatter_kwargs = percentage_formatter_kwargs(decimals = 5))


app = demo()
app.mainloop()