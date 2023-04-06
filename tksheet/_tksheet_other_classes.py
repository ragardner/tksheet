from ._tksheet_vars import *
from collections import namedtuple
from itertools import islice
import tkinter as tk


CurrentlySelectedClass = namedtuple("CurrentlySelectedClass", "row column type_")
CtrlKeyEvent = namedtuple("CtrlKeyEvent", "eventname selectionboxes currentlyselected rows")
PasteEvent = namedtuple("PasteEvent", "eventname currentlyselected rows")
UndoEvent = namedtuple("UndoEvent", "eventname type storeddata")
SelectCellEvent = namedtuple("SelectCellEvent", "eventname row column")
SelectColumnEvent = namedtuple("SelectColumnEvent", "eventname column")
SelectRowEvent = namedtuple("SelectRowEvent", "eventname row")
DeselectionEvent = namedtuple("DeselectionEvent", "eventname selectionboxes")
SelectionBoxEvent = namedtuple("SelectionBoxEvent", "eventname selectionboxes")
InsertEvent = namedtuple("InsertEvent", "eventname dataindex displayindex quantity")
DeleteRowColumnEvent = namedtuple("DeleteRowColumnEvent", "eventname deleteindexes")
EditCellEvent = namedtuple("EditCellEvent", "row column key text eventname")
EditHeaderEvent = namedtuple("EditHeaderEvent", "column key text eventname")
EditIndexEvent = namedtuple("EditIndexEvent", "row key text eventname")
BeginDragDropEvent = namedtuple("BeginDragDropEvent", "eventname columnstomove movedto")
EndDragDropEvent = namedtuple("EndDragDropEvent", "eventname oldindexes newindexes movedto")
ResizeEvent = namedtuple("ResizeEvent", "eventname index oldsize newsize")
DropDownModifiedEvent = namedtuple("DropDownModifiedEvent", "eventname row column value")
DrawnItem = namedtuple("DrawnItem", "iid showing")
TextCfg = namedtuple("TextCfg", "txt tf font align")


class TextEditor_(tk.Text):
    def __init__(self,
                 parent,
                 font = get_font(),
                 text = None,
                 state = "normal",
                 bg = "white",
                 fg = "black",
                 popup_menu_font = ("Arial", 11, "normal"),
                 popup_menu_bg = "white",
                 popup_menu_fg = "black",
                 popup_menu_highlight_bg = "blue",
                 popup_menu_highlight_fg = "white",
                 align = "w",
                 newline_binding = None):
        tk.Text.__init__(self,
                         parent,
                         font = font,
                         state = state,
                         spacing1 = 0,
                         spacing2 = 0,
                         spacing3 = 0,
                         bd = 0,
                         highlightthickness = 0,
                         undo = True,
                         maxundo = 30,
                         background = bg,
                         foreground = fg,
                         insertbackground = fg)
        self.parent = parent
        self.newline_bindng = newline_binding
        if align == "w":
            self.align = "left"
        elif align == "center":
            self.align = "center"
        elif align == "e":
            self.align = "right"
        self.tag_configure("align", justify = self.align)
        if text:
            self.insert(1.0, text)
            self.yview_moveto(1)
        self.tag_add("align", 1.0, "end")
        self.rc_popup_menu = tk.Menu(self, tearoff = 0)
        self.rc_popup_menu.add_command(label = "Select all",
                                       accelerator = "Ctrl+A",
                                       font = popup_menu_font,
                                       foreground = popup_menu_fg,
                                       background = popup_menu_bg,
                                       activebackground = popup_menu_highlight_bg,
                                       activeforeground = popup_menu_highlight_fg,
                                       command = self.select_all)
        self.rc_popup_menu.add_command(label = "Cut",
                                       accelerator = "Ctrl+X",
                                       font = popup_menu_font,
                                       foreground = popup_menu_fg,
                                       background = popup_menu_bg,
                                       activebackground = popup_menu_highlight_bg,
                                       activeforeground = popup_menu_highlight_fg,
                                       command = self.cut)
        self.rc_popup_menu.add_command(label = "Copy",
                                       accelerator = "Ctrl+C",
                                       font = popup_menu_font,
                                       foreground = popup_menu_fg,
                                       background = popup_menu_bg,
                                       activebackground = popup_menu_highlight_bg,
                                       activeforeground = popup_menu_highlight_fg,
                                       command = self.copy)
        self.rc_popup_menu.add_command(label = "Paste",
                                       accelerator = "Ctrl+V",
                                       font = popup_menu_font,
                                       foreground = popup_menu_fg,
                                       background = popup_menu_bg,
                                       activebackground = popup_menu_highlight_bg,
                                       activeforeground = popup_menu_highlight_fg,
                                       command = self.paste)
        self.rc_popup_menu.add_command(label = "Undo",
                                       accelerator = "Ctrl+Z",
                                       font = popup_menu_font,
                                       foreground = popup_menu_fg,
                                       background = popup_menu_bg,
                                       activebackground = popup_menu_highlight_bg,
                                       activeforeground = popup_menu_highlight_fg,
                                       command = self.undo)
        self.bind("<1>", lambda event: self.focus_set())
        if USER_OS == "darwin":
            self.bind("<2>", self.rc)
        else:
            self.bind("<3>", self.rc)
        self._orig = self._w + "_orig"
        self.tk.call("rename", self._w, self._orig)
        self.tk.createcommand(self._w, self._proxy)

    def _proxy(self, command, *args):
        cmd = (self._orig, command) + args
        try:
            result = self.tk.call(cmd)
        except:
            return
        if command in ("insert", "delete", "replace"):
            self.tag_add("align", 1.0, "end")
            self.event_generate("<<TextModified>>")
            if args and len(args) > 1 and args[1] != '\n':
                out_of_bounds = self.yview()
                if out_of_bounds != (0.0, 1.0) and self.newline_bindng is not None:
                    self.newline_bindng(r = self.parent.r, c = self.parent.c, check_lines = False)
        return result
    
    def rc(self,event):
        self.focus_set()
        self.rc_popup_menu.tk_popup(event.x_root, event.y_root)
        
    def select_all(self, event = None):
        self.event_generate("<Command-a>" if is_mac() else "<Control-a>")
        return "break"
    
    def cut(self, event = None):
        self.event_generate("<Command-x>" if is_mac() else "<Control-x>")
        return "break"
    
    def copy(self, event = None):
        self.event_generate("<Command-c>" if is_mac() else "<Control-c>")
        return "break"
    
    def paste(self, event = None):
        self.event_generate("<Command-v>" if is_mac() else "<Control-v>")
        return "break"

    def undo(self, event = None):
        self.event_generate("<Command-z>" if is_mac() else "<Control-z>")
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
                 show_border = True,
                 bg = "white",
                 fg = "black",
                 popup_menu_font = ("Arial", 11, "normal"),
                 popup_menu_bg = "white",
                 popup_menu_fg = "black",
                 popup_menu_highlight_bg = "blue",
                 popup_menu_highlight_fg = "white",
                 binding = None,
                 align = "w",
                 r = 0,
                 c = 0,
                 newline_binding = None):
        tk.Frame.__init__(self,
                          parent,
                          height = height,
                          width = width,
                          highlightbackground = border_color,
                          highlightcolor = border_color,
                          highlightthickness = 2 if show_border else 0,
                          bd = 0)
        self.parent = parent
        self.r = r
        self.c = c
        self.textedit = TextEditor_(self,
                                    font = font,
                                    text = text,
                                    state = state,
                                    bg = bg,
                                    fg = fg,
                                    popup_menu_font = popup_menu_font,
                                    popup_menu_bg = popup_menu_bg,
                                    popup_menu_fg = popup_menu_fg,
                                    popup_menu_highlight_bg = popup_menu_highlight_bg,
                                    popup_menu_highlight_fg = popup_menu_highlight_fg,
                                    align = align,
                                    newline_binding = newline_binding)
        self.textedit.grid(row = 0,
                           column = 0,
                           sticky = "nswe")
        self.grid_columnconfigure(0, weight = 1)
        self.grid_rowconfigure(0, weight = 1)
        self.grid_propagate(False)
        self.w_ = width
        self.h_ = height
        self.binding = binding
        self.textedit.focus_set()
        
    def get(self):
        return self.textedit.get("1.0", "end-1c")

    def get_num_lines(self):
        return int(self.textedit.index('end-1c').split('.')[0])

    def set_text(self, text):
        self.textedit.delete(1.0, "end")
        self.textedit.insert(1.0, text)

    def scroll_to_bottom(self):
        self.textedit.yview_moveto(1)
        
        
class GeneratedMouseEvent:
    def __init__(self):
        self.keycode = "??"
        self.num = 1


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

def is_mac():
    if USER_OS == "darwin":
        return True
    else:
        return False

def get_rc_binding():
    if USER_OS == "darwin":
        return "<2>"
    else:
        return "<3>"
        
