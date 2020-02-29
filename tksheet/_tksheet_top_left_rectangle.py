from ._tksheet_vars import *
from ._tksheet_other_classes import *

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


class TopLeftRectangle(tk.Canvas):
    def __init__(self,
                 parentframe = None,
                 main_canvas = None,
                 row_index_canvas = None,
                 header_canvas = None,
                 background = None,
                 foreground = None,
                 foreground_highlight = None):
        tk.Canvas.__init__(self,
                           parentframe,
                           background = background,
                           highlightthickness = 0)
        self.parentframe = parentframe
        self.rectangle_foreground = foreground
        self.resizers_highlight = foreground_highlight
        self.MT = main_canvas
        self.RI = row_index_canvas
        self.CH = header_canvas
        self.config(width = self.RI.current_width, height = self.CH.current_height)
        self.extra_motion_func = None
        self.extra_b1_press_func = None
        self.extra_b1_motion_func = None
        self.extra_b1_release_func = None
        self.extra_double_b1_func = None
        self.extra_rc_func = None
        self.MT.TL = self
        self.RI.TL = self
        self.CH.TL = self
        w = self.RI.current_width - 1
        h = self.CH.current_height - 1
        self.create_rectangle(0, h - 5, w - 5, h, fill = self.rectangle_foreground, outline = "", tag = "rw")
        self.create_rectangle(w - 5, 0, w, h, fill = self.rectangle_foreground, outline = "", tag = "rh")
        self.tag_bind("rw", "<Enter>", self.rw_enter)
        self.tag_bind("rh", "<Enter>", self.rh_enter)
        self.tag_bind("rw", "<Leave>", self.rw_leave)
        self.tag_bind("rh", "<Leave>", self.rh_leave)
        self.bind("<Motion>", self.mouse_motion)
        self.bind("<ButtonPress-1>", self.b1_press)
        self.bind("<B1-Motion>", self.b1_motion)
        self.bind("<ButtonRelease-1>", self.b1_release)
        self.bind("<Double-Button-1>", self.double_b1)
        self.bind(get_rc_binding(), self.rc)

    def rw_enter(self, event = None):
        self.itemconfig("rw", fill = self.resizers_highlight)

    def rh_enter(self, event = None):
        self.itemconfig("rh", fill = self.resizers_highlight)

    def rw_leave(self, event = None):
        self.itemconfig("rw", fill = self.rectangle_foreground)

    def rh_leave(self, event = None):
        self.itemconfig("rh", fill = self.rectangle_foreground)

    def basic_bindings(self, enable = True):
        if enable:
            self.bind("<Motion>", self.mouse_motion)
            self.bind("<ButtonPress-1>", self.b1_press)
            self.bind("<B1-Motion>", self.b1_motion)
            self.bind("<ButtonRelease-1>", self.b1_release)
            self.bind("<Double-Button-1>", self.double_b1)
            self.bind(get_rc_binding(), self.rc)
        else:
            self.unbind("<Motion>")
            self.unbind("<ButtonPress-1>")
            self.unbind("<B1-Motion>")
            self.unbind("<ButtonRelease-1>")
            self.unbind("<Double-Button-1>")
            self.unbind(get_rc_binding())

    def set_dimensions(self, new_w = None, new_h = None):
        if new_w:
            self.config(width = new_w)
            w = new_w - 1
            h = self.winfo_height() - 1
        if new_h:
            self.config(height = new_h)
            w = self.winfo_width() - 1
            h = new_h - 1
        self.coords("rw", 0, h - 5, w - 5, h)
        self.coords("rh", w - 5, 0, w, h)

    def mouse_motion(self, event = None):
        self.MT.reset_mouse_motion_creations()
        if self.extra_motion_func is not None:
            self.extra_motion_func(event)

    def b1_press(self, event = None):
        self.focus_set()
        rect = self.find_overlapping(event.x, event.y, event.x, event.y)
        if not rect:
            if self.MT.drag_selection_enabled:
                self.MT.select_all()
        elif rect[0] == 1:
            if self.RI.width_resizing_enabled:
                self.RI.set_width(self.RI.default_width, set_TL = True)
        elif rect[0] == 2:
            if self.CH.height_resizing_enabled:
                self.CH.set_height(self.MT.default_hh[1], set_TL = True)
        self.MT.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
        if self.extra_b1_press_func is not None:
            self.extra_b1_press_func(event)

    def b1_motion(self, event = None):
        self.focus_set()
        if self.extra_b1_motion_func is not None:
            self.extra_b1_motion_func(event)

    def b1_release(self, event = None):
        self.focus_set()
        if self.extra_b1_release_func is not None:
            self.extra_b1_release_func(event)

    def double_b1(self, event = None):
        self.focus_set()
        if self.extra_double_b1_func is not None:
            self.extra_double_b1_func(event)

    def rc(self, event = None):
        self.focus_set()
        if self.extra_rc_func is not None:
            self.extra_rc_func(event)


