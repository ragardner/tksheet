# tksheet [![PyPI version shields.io](https://img.shields.io/pypi/v/tksheet.svg)](https://pypi.python.org/pypi/tksheet/) ![python](https://img.shields.io/badge/python-3.6+-blue) [![License: MIT](https://img.shields.io/badge/License-MIT%20-blue.svg)](https://github.com/ragardner/tksheet/blob/master/LICENSE.txt) [![GitHub Release Date](https://img.shields.io/github/release-date-pre/ragardner/tksheet.svg)](https://github.com/ragardner/tksheet/releases)


Python tkinter table widget

```
pip install tksheet
```

### [Version release notes](https://github.com/ragardner/tksheet/blob/master/RELEASE_NOTES.md)
### [Documentation](https://github.com/ragardner/tksheet/blob/master/DOCUMENTATION.md) (unfinished)

### Features
 - Display and modify tabular data
 - Stores its display data as a Python list of lists, sublists being rows
 - Runs smoothly even with millions of rows/columns
 - Edit cells directly
 - Drag and drop columns and rows
 - Multiple line headers and rows
 - Expand row heights and column widths
 - Change fonts and font size
 - Change any colors in the sheet
 - Left or Centre text alignment
 - Cell elements can be any class with a `str` method

Work on this repository is ongoing, improvements in usability, functionality and finally documentation

Planned future changes:
 - Ctrl + click selection
 - Editing row index/header
 - Display subset of rows
 - More builtin functionality on right click
 - Better functions to access and manipulate table data
 - Better behavior from some functions e.g. maintaining correct selections when inserting columns
 - If it's popular enough or someone is willing to pay for my time I will integrate pandas dataframe
   to save memory when using large dataframes

### Basic Demo:

```python
from tksheet import Sheet
import tkinter as tk


class demo(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)
        self.grid_columnconfigure(0, weight = 1)
        self.grid_rowconfigure(0, weight = 1)
        self.frame = tk.Frame(self)
        self.frame.grid_columnconfigure(0, weight = 1)
        self.frame.grid_rowconfigure(0, weight = 1)
        self.sheet = Sheet(self.frame,
                           #empty_vertical = 0,
                           #empty_horizontal = 0,
                           #show_vertical_grid = False,
                           #show_horizontal_grid = False,
                           #auto_resize_numerical_row_index = False,
                           #header_height = "3",
                           #row_index_width = 100,
                           #align = "center",
                           #header_align = "w",
                            #row_index_align = "w",
                            #theme = "dark",
                            data = [[f"Row {r}, Column {c}\nnewline1\nnewline2" for c in range(30)] for r in range(2000)], #to set sheet data at startup
                            #headers = [f"Column {c}\nnewline1\nnewline2" for c in range(30)],
                            #row_index = [f"Row {r}\nnewline1\nnewline2" for r in range(2000)],
                            #set_all_heights_and_widths = True, #to fit all cell sizes to text at start up
                            #headers = 0, #to set headers as first row at startup
                            headers = [f"Column {c}\nnewline1\nnewline2" for c in range(30)],
                            #row_index = 0, #to set row_index as first column at startup
                            #total_rows = 2000, #if you want to set empty sheet dimensions at startup
                            #total_columns = 30, #if you want to set empty sheet dimensions at startup
                            height = 500, #height and width arguments are optional
                            width = 1200 #For full startup arguments see DOCUMENTATION.md
                            )
        #self.sheet.hide("row_index")
        #self.sheet.hide("header")
        #self.sheet.hide("top_left")
        self.sheet.enable_bindings(("single_select", #"single_select" or "toggle_select"
                                         "drag_select",   #enables shift click selection as well
                                         "column_drag_and_drop",
                                         "row_drag_and_drop",
                                         "column_select",
                                         "row_select",
                                         "column_width_resize",
                                         "double_click_column_resize",
                                         #"row_width_resize",
                                         #"column_height_resize",
                                         "arrowkeys",
                                         "row_height_resize",
                                         "double_click_row_resize",
                                         "right_click_popup_menu",
                                         "rc_select",
                                         "rc_insert_column",
                                         "rc_delete_column",
                                         "rc_insert_row",
                                         "rc_delete_row",
                                         "copy",
                                         "cut",
                                         "paste",
                                         "delete",
                                         "undo",
                                         "edit_cell"))
        #self.sheet.enable_bindings("enable_all")
        #self.sheet.disable_bindings() #uses the same strings
        #self.bind("<Configure>", self.window_resized)
        self.frame.grid(row = 0, column = 0, sticky = "nswe")
        self.sheet.grid(row = 0, column = 0, sticky = "nswe")
        
        """_________________________ EXAMPLES _________________________ """
        """_____________________________________________________________"""

        # __________ CHANGING THEME __________

        #self.sheet.change_theme("dark")

        # __________ HIGHLIGHT / DEHIGHLIGHT CELLS __________
        
        self.sheet.highlight_cells(row = 5, column = 5, bg = "#ed4337", fg = "white")
        self.sheet.highlight_cells(row = 5, column = 1, bg = "#ed4337", fg = "white")
        self.sheet.highlight_cells(row = 5, bg = "#ed4337", fg = "white", canvas = "row_index")
        self.sheet.highlight_cells(column = 0, bg = "#ed4337", fg = "white", canvas = "header")

        # __________ DISPLAY SUBSET OF COLUMNS __________

        #self.sheet.display_subset_of_columns(indexes = [3, 1, 2], enable = True) #any order

        # __________ DATA AND DISPLAY DIMENSIONS __________

        #self.sheet.total_rows(4) #will delete rows if set to less than current data rows
        #self.sheet.total_columns(2) #will delete columns if set to less than current data columns
        #self.sheet.sheet_data_dimensions(total_rows = 4, total_columns = 2)
        #self.sheet.sheet_display_dimensions(total_rows = 4, total_columns = 6) #currently resets widths and heights
        #self.sheet.set_sheet_data_and_display_dimensions(total_rows = 4, total_columns = 2) #currently resets widths and heights

        # __________ SETTING OR RESETTING TABLE DATA __________

        #.set_sheet_data() function returns the object you use as argument
        #verify checks if your data is a list of lists, raises error if not
        #self.data = self.sheet.set_sheet_data([[f"Row {r} Column {c}" for c in range(30)] for r in range(2000)], verify = False)

        # __________ SETTING ROW HEIGHTS AND COLUMN WIDTHS __________

        #self.sheet.set_cell_data(0, 0, "\n".join([f"Line {x}" for x in range(500)]))
        #self.sheet.set_column_data(1, ("" for i in range(2000)))
        #self.sheet.row_index((f"Row {r}" for r in range(2000))) #any iterable works
        #self.sheet.row_index("\n".join([f"Line {x}" for x in range(500)]), 2)
        #self.sheet.column_width(column = 0, width = 300)
        #self.sheet.row_height(row = 0, height = 60)
        #self.sheet.set_column_widths([120 for c in range(30)])
        #self.sheet.set_row_heights([30 for r in range(2000)])
        #self.sheet.set_all_column_widths()
        #self.sheet.set_all_row_heights()
        #self.sheet.set_all_cell_sizes_to_text()
        
        # __________ BINDING A FUNCTIONS TO USER ACTIONS __________

        self.sheet.extra_bindings([("cell_select", self.cell_select),
                                    #("begin_edit_cell", self.begin_edit_cell),
                                    ("shift_cell_select", self.shift_select_cells),
                                    ("drag_select_cells", self.drag_select_cells),
                                    ("ctrl_a", self.ctrl_a),
                                    ("row_select", self.row_select),
                                    ("shift_row_select", self.shift_select_rows),
                                    ("drag_select_rows", self.drag_select_rows),
                                    ("column_select", self.column_select),
                                    ("shift_column_select", self.shift_select_columns),
                                    ("drag_select_columns", self.drag_select_columns),
                                    ("deselect", self.deselect)
                                    ])
        
        #self.sheet.extra_bindings([("cell_select", None)]) #unbind cell select
        #self.sheet.extra_bindings("unbind_all") #remove all functions set by extra_bindings()

        # __________ BINDING NEW RIGHT CLICK FUNCTION __________
    
        self.sheet.bind("<3>", self.rc)

        # __________ SETTING HEADERS __________

        #self.sheet.headers((f"Header {c}" for c in range(30))) #any iterable works
        #self.sheet.headers("Change header example", 2)
        #print (self.sheet.headers())
        #print (self.sheet.headers(index = 2))

        # __________ SETTING ROW INDEX __________

        #self.sheet.row_index((f"Row {r}" for r in range(2000))) #any iterable works
        #self.sheet.row_index("Change index example", 2)
        #print (self.sheet.row_index())
        #print (self.sheet.row_index(index = 2))

        # __________ INSERTING A ROW __________

        #self.sheet.insert_row(values = (f"my new row here {c}" for c in range(30)), idx = 0) # a filled row at the start
        #self.sheet.insert_row() # an empty row at the end

        # __________ INSERTING A COLUMN __________

        #self.sheet.insert_column(values = (f"my new col here {r}" for r in range(2050)), idx = 0) # a filled column at the start
        #self.sheet.insert_column() # an empty column at the end

        # __________ SETTING A COLUMNS DATA __________

        # any iterable works
        #self.sheet.set_column_data(0, values = (0 for i in range(2050)))

        # __________ SETTING A ROWS DATA __________

        # any iterable works
        #self.sheet.set_row_data(0, values = (0 for i in range(35)))

        # __________ SETTING A CELLS DATA __________

        #self.sheet.set_cell_data(1, 2, "NEW VALUE")

        # __________ GETTING FULL SHEET DATA __________

        #self.all_data = self.sheet.get_sheet_data()

        # __________ GETTING CELL DATA __________

        #print (self.sheet.get_cell_data(0, 0))

        # __________ GETTING ROW DATA __________

        #print (self.sheet.get_row_data(0)) # only accessible by index

        # __________ GETTING COLUMN DATA __________

        #print (self.sheet.get_column_data(0)) # only accessible by index

        # __________ GETTING SELECTED __________

        #print (self.sheet.get_currently_selected())
        #print (self.sheet.get_selected_cells())
        #print (self.sheet.get_selected_rows())
        #print (self.sheet.get_selected_columns())
        #print (self.sheet.get_selection_boxes())
        #print (self.sheet.get_selection_boxes_with_types())

        # __________ SETTING SELECTED __________

        #self.sheet.deselect("all")
        #self.sheet.create_selection_box(0, 0, 2, 2, type_ = "cells") #type here is "cells", "cols" or "rows"
        #self.sheet.set_currently_selected(0, 0)
        #self.sheet.set_currently_selected("row", 0)
        #self.sheet.set_currently_selected("column", 0)

        # __________ CHECKING SELECTED __________

        #print (self.sheet.is_cell_selected(0, 0))
        #print (self.sheet.is_row_selected(0))
        #print (self.sheet.is_column_selected(0))
        #print (self.sheet.anything_selected())

        # __________ HIDING THE ROW INDEX AND HEADERS __________

        #self.sheet.hide("row_index")
        #self.sheet.hide("top_left")
        #self.sheet.hide("header")

        # __________ ADDITIONAL BINDINGS __________

        #self.sheet.bind("<Motion>", self.mouse_motion)

    """

    UNTIL DOCUMENTATION IS COMPLETE, PLEASE BROWSE THE FILE
    _tksheet.py FOR A FULL LIST OF FUNCTIONS AND THEIR PARAMETERS

    """
    
    def begin_edit_cell(self, event):
        pass

    def window_resized(self, event):
        pass
        #print (event)

    def mouse_motion(self, event):
        region = self.sheet.identify_region(event)
        row = self.sheet.identify_row(event, allow_end = False)
        column = self.sheet.identify_column(event, allow_end = False)
        print (region, row, column)

    def deselect(self, event):
        print (event, self.sheet.get_selected_cells())

    def rc(self, event):
        print (event)
        
    def cell_select(self, response):
        pass
        #print (response)

    def shift_select_cells(self, response):
        print (response)

    def drag_select_cells(self, response):
        pass
        #print (response)

    def ctrl_a(self, response):
        print (response)

    def row_select(self, response):
        print (response)

    def shift_select_rows(self, response):
        print (response)

    def drag_select_rows(self, response):
        pass
        #print (response)
        
    def column_select(self, response):
        print (response)

    def shift_select_columns(self, response):
        print (response)

    def drag_select_columns(self, response):
        pass
        #print (response)
        
app = demo()
app.mainloop()

```

----

### Light Theme

![alt text](https://i.imgur.com/PWJSPxf.jpg)


### Dark Theme

![alt text](https://i.imgur.com/OOqAARe.jpg)

----