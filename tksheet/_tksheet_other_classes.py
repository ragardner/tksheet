from ._tksheet_vars import *
from collections import defaultdict, deque
from itertools import islice, repeat, accumulate, chain, product, cycle
from math import floor, ceil
from tkinter import ttk
import bisect
import csv as csv_module
import io
import pickle
import re
import tkinter as tk
import zlib
# for mac bindings
from platform import system as get_os


class TextEditor_(tk.Text):
    def __init__(self,
                 parent,
                 font = get_font(),
                 text = None,
                 state = "normal"):
        tk.Text.__init__(self,
                         parent,
                         font = font,
                         state = state,
                         spacing1 = 2,
                         spacing2 = 2,
                         bd = 0,
                         highlightthickness = 0,
                         undo = True,
                         maxundo = 20,
                         background = parent.parent.table_bg,
                         foreground = parent.parent.table_fg,
                         insertbackground = parent.parent.table_fg)
        self.parent = parent
        if text is not None:
            self.insert(1.0, text)
        self.yview_moveto(1)
        self.rc_popup_menu = tk.Menu(self, tearoff = 0)
        self.rc_popup_menu.add_command(label = "Select all",
                                       accelerator = "Ctrl+A",
                                       font = self.parent.parent.popup_menu_font,
                                       foreground = self.parent.parent.popup_menu_fg,
                                       background = self.parent.parent.popup_menu_bg,
                                       activebackground = self.parent.parent.popup_menu_highlight_bg,
                                       activeforeground = self.parent.parent.popup_menu_highlight_fg,
                                       command = self.select_all)
        self.rc_popup_menu.add_command(label = "Cut",
                                       accelerator = "Ctrl+X",
                                       font = self.parent.parent.popup_menu_font,
                                       foreground = self.parent.parent.popup_menu_fg,
                                       background = self.parent.parent.popup_menu_bg,
                                       activebackground = self.parent.parent.popup_menu_highlight_bg,
                                       activeforeground = self.parent.parent.popup_menu_highlight_fg,
                                       command = self.cut)
        self.rc_popup_menu.add_command(label = "Copy",
                                       accelerator = "Ctrl+C",
                                       font = self.parent.parent.popup_menu_font,
                                       foreground = self.parent.parent.popup_menu_fg,
                                       background = self.parent.parent.popup_menu_bg,
                                       activebackground = self.parent.parent.popup_menu_highlight_bg,
                                       activeforeground = self.parent.parent.popup_menu_highlight_fg,
                                       command = self.copy)
        self.rc_popup_menu.add_command(label = "Paste",
                                       accelerator = "Ctrl+V",
                                       font = self.parent.parent.popup_menu_font,
                                       foreground = self.parent.parent.popup_menu_fg,
                                       background = self.parent.parent.popup_menu_bg,
                                       activebackground = self.parent.parent.popup_menu_highlight_bg,
                                       activeforeground = self.parent.parent.popup_menu_highlight_fg,
                                       command = self.paste)
        self.rc_popup_menu.add_command(label = "Undo",
                                       accelerator = "Ctrl+Z",
                                       font = self.parent.parent.popup_menu_font,
                                       foreground = self.parent.parent.popup_menu_fg,
                                       background = self.parent.parent.popup_menu_bg,
                                       activebackground = self.parent.parent.popup_menu_highlight_bg,
                                       activeforeground = self.parent.parent.popup_menu_highlight_fg,
                                       command = self.undo)
        self.bind("<1>", lambda event: self.focus_set())
        if str(get_os()) == "Darwin":
            self.bind("<2>", self.rc)
        else:
            self.bind("<3>", self.rc)
    
    def rc(self,event):
        self.focus_set()
        self.rc_popup_menu.tk_popup(event.x_root, event.y_root)
        
    def select_all(self, event = None):
        self.event_generate("<Control-a>")
        return "break"
    
    def cut(self, event = None):
        self.event_generate("<Control-x>")
        return "break"
    
    def copy(self, event = None):
        self.event_generate("<Control-c>")
        return "break"
    
    def paste(self, event = None):
        self.event_generate("<Control-v>")
        return "break"

    def undo(self, event = None):
        self.event_generate("<Control-z>")
        return "break"


class TextEditor(tk.Frame):
    def __init__(self,
                 parent,
                 font = get_font(),
                 text = None,
                 state = "normal",
                 width = None,
                 height = None,
                 border_color = "black",
                 show_border = True):
        tk.Frame.__init__(self,
                          parent,
                          height = height,
                          width = width,
                          highlightbackground = border_color,
                          highlightcolor = border_color,
                          highlightthickness = 2 if show_border else 0,
                          bd = 0)
        self.parent = parent
        self.textedit = TextEditor_(self,
                                    font = font,
                                    text = text,
                                    state = state)
        self.textedit.grid(row = 0,
                           column = 0,
                           sticky = "nswe")
        self.grid_columnconfigure(0, weight = 1)
        self.grid_rowconfigure(0, weight = 1)
        self.grid_propagate(False)
        self.textedit.focus_set()
        
    def get(self):
        return self.textedit.get("1.0", "end").rstrip()

    def get_num_lines(self):
        return int(self.textedit.index('end-1c').split('.')[0])

    def set_text(self, text):
        self.textedit.delete(1.0, "end")
        self.textedit.insert(1.0, text)

    def scroll_to_bottom(self):
        self.textedit.yview_moveto(1)


class TableDropdown(tk.Frame):
    def __init__(self,
                 parent,
                 font,
                 state,
                 values = [],
                 set_value = None,
                 width = None,
                 height = None):
        tk.Frame.__init__(self,
                          parent)
        if width:
            self.config(width = width)
        if height:
            self.config(height = height)
        self.parent = parent
        self.grid_columnconfigure(0, weight = 1)
        self.grid_rowconfigure(0, weight = 1)
        self.dropdown = TableDropdown_(self,
                                       font,
                                       state,
                                       values = values,
                                       set_value = set_value)
        self.dropdown.grid(row = 0,
                           column = 0,
                           sticky = "nswe")
        self.grid_propagate(False)
        self.dropdown.focus_set()
        
    def get_my_value(self, event = None):
        return self.dropdown.displayed.get()
    
    def set_displayed(self, value, event = None):
        self.dropdown.displayed.set(value)

    def set_my_values(self, values = []):
        self.dropdown['values'] = values


class TableDropdown_(ttk.Combobox):
    def __init__(self,
                 parent,
                 font,
                 state,
                 values = [],
                 set_value = None):
        self.displayed = tk.StringVar()
        ttk.Combobox.__init__(self,
                              parent,
                              font = font,
                              state = state,
                              values = values,
                              textvariable = self.displayed)
        if set_value is not None:
            self.displayed.set(set_value)
        elif values:
            self.displayed.set(values[0])
            
    def get_my_value(self, event = None):
        return self.displayed.get()
    
    def set_my_value(self, value, event = None):
        self.displayed.set(value)

    def set_my_values(self, values = []):
        self['values'] = values


def num2alpha(n):
    s = ""
    n += 1
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

def get_index_of_gap_in_sorted_integer_seq_forward(seq, start = 0):
    prevn = seq[start]
    for idx, n in enumerate(islice(seq, start + 1, None), start + 1):
        if n != prevn + 1:
            return idx
        prevn = n
    return None

def get_index_of_gap_in_sorted_integer_seq_reverse(seq, start = 0):
    prevn = seq[start]
    for idx, n in zip(range(start, -1, -1), reversed(seq[:start])):
        if n != prevn - 1:
            return idx
        prevn = n
    return None

def get_rc_binding():
        if f"{get_os()}" == "Darwin":
            return "<2>"
        else:
            return "<3>"
        
