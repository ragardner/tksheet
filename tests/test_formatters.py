from tksheet import *
from tksheet._tksheet_formatters import *
import tkinter as tk
from datetime import datetime, date, timedelta, time
from dateutil import parser, tz
from math import ceil
import re

def round_up(x):
    return float(ceil(x))

def only_numeric(s):
    return ''.join(n for n in f"{s}" if n.isnumeric() or n == '.')

date_replace = re.compile('|'.join(['\(','\)','\[','\]','\<','\>',]))

def convert_to_local_datetime(dt: str, **kwargs):
    if isinstance(dt, datetime):
        pass
    elif isinstance(dt, date):
        dt = datetime(dt.year, dt.month, dt.day)
    else:
        if isinstance(dt, str):
            dt = date_replace.sub("", dt)
        dt = parser.parse(dt)
    if dt.tzinfo is None:
        dt.replace(tzinfo = tz.tzlocal())
    dt = dt.astimezone(tz.tzlocal())
    return dt.replace(tzinfo = None)

def datetime_to_string(dt: datetime, **kwargs):
    return dt.strftime('%d %b, %Y, %H:%M:%S')

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
                           data = [[f"{r}"]*11 for r in range(10)],
                           #header = 0,
                           theme = "dark",
                           )
        self.sheet.enable_bindings("all", "edit_header", "edit_index")
        self.frame.grid(row = 0, column = 0, sticky = "nswe")
        self.sheet.grid(row = 0, column = 0, sticky = "nswe")

        #self.sheet.format_cell('all', 0, formatter_options = float_formatter(nullable = False))
        #self.sheet.format_cell('all', 1, formatter_options = float_formatter())
        #self.sheet.format_cell('all', 2, formatter_options = int_formatter())
        #self.sheet.format_cell('all', 3, formatter_options = bool_formatter())
        #self.sheet.format_cell('all', 4, formatter_options = percentage_formatter())
        #self.sheet.format_cell('all', 5, formatter_options = formatter(datatypes = datetime, 
        #                                                               format_function = convert_to_local_datetime, 
        #                                                               to_str_function = datetime_to_string, 
        #                                                               nullable = False,
        #                                                               invalid_value = 'NaT',
        #                                                               ))
        #self.sheet.format_cell('all', 6, datatypes = datetime, 
        #                                 format_function = convert_to_local_datetime, 
        #                                 to_str_function = custom_datetime_to_str, 
        #                                 nullable = True,
        #                                 invalid_value = 'NaT',
        #                                 format = '%Y/%m/%d %H:%M %p'
        #                                 )

        #self.sheet.format_cell('all', 7, formatter_options = float_formatter(post_format_function = round_up))
        #self.sheet.format_cell('all', 8, formatter_options = float_formatter(), pre_format_function = only_numeric)

        #self.sheet.create_dropdown('all', 0, values = ['', '104%', .24, , "300%", 'not a number'], set_value = 1)
        #self.sheet.format_column(9, formatter_options = percentage_formatter(), decimals = 0)
        #self.sheet.format_cell('all', 10, formatter_options = percentage_formatter(decimals = 5), formatter_class = Formatter)
        
        self.sheet.format_sheet(formatter(datatypes = datetime, 
                                          format_function = convert_to_local_datetime, 
                                          to_str_function = datetime_to_string, 
                                          nullable = False,
                                          invalid_value = 'NaT',
                                          ))
        self.sheet.format_row(0, int_formatter())
        self.sheet.format_column(0, bool_formatter(nullable = False))
        self.sheet.format_cell(0, 0, percentage_formatter(nullable = False))
        self.sheet.highlight_rows(0, bg = "purple", fg = "white")
        self.sheet.set_cell_data(0, 0, "10.546445%")
        

    def del_rows(self, event = None):
        self.sheet.delete_rows(self.sheet.get_selected_rows())
        
    def del_cols(self, event = None):
        self.sheet.delete_columns(self.sheet.get_selected_columns())


app = demo()
app.mainloop()