# tksheet 
[![PyPI version shields.io](https://img.shields.io/pypi/v/tksheet.svg)](https://pypi.python.org/pypi/tksheet/)
[![License: MIT](https://img.shields.io/badge/License-MIT%20-blue.svg)](https://github.com/ragardner/tksheet/blob/master/LICENSE.txt)
[![GitHub Release Date](https://img.shields.io/github/release-date-pre/ragardner/tksheet.svg)](https://github.com/ragardner/tksheet/releases)

Python 3.6+ tkinter table widget for displaying tabular data

```
pip install tksheet
```

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
 - Copes with cell elements not being strings
 
Work on this repository is ongoing, improvements in usability, functionality and finally documentation
 
[Version release notes](https://github.com/ragardner/tksheet/blob/master/RELEASE_NOTES.md)

# Basic Demo:

```python
import tkinter as tk
from tksheet import Sheet

class demo(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)
        self.state("zoomed")
        self.grid_columnconfigure(0, weight = 1)
        self.grid_rowconfigure(0, weight = 1)
        
        self.sheet_demo = Sheet(self,
                                width = 1000,
                                height = 700,
                                align = "w",
                                header_align = "center",
                                row_index_align = "center",
                                column_width = 180,
                                row_index_width = 50,
                                total_rows = 5000,
                                total_columns = 100,
                                headers = [f"Header {c}" for c in range(100)])
        self.sheet_demo.enable_bindings(("single",
                                         "drag_select",
                                         "column_drag_and_drop",
                                         "row_drag_and_drop",
                                         "column_select",
                                         "row_select",
                                         "column_width_resize",
                                         "double_click_column_resize",
                                         "row_width_resize",
                                         "column_height_resize",
                                         "arrowkeys",
                                         "row_height_resize",
                                         "double_click_row_resize",
                                         "edit_bindings",
                                         "rc_insert_column",
                                         "rc_delete_column",
                                         "rc_insert_row",
                                         "rc_delete_row"))
        self.sheet_demo.grid(row = 0, column = 0, sticky = "nswe")
        self.sheet_demo.highlight_cells(row = 0, column = 0, bg = "orange", fg = "blue")
        self.sheet_demo.highlight_cells(row = 0, bg = "orange", fg = "blue", canvas = "row_index")
        self.sheet_demo.highlight_cells(column = 0, bg = "orange", fg = "blue", canvas = "header")

        """_________________________ EXAMPLES _________________________ """
        """_____________________________________________________________"""

        # __________ CHANGING THEME __________

        self.sheet_demo.change_theme("dark")

        # __________ SETTING OR RESETTING TABLE DATA __________
        
        self.data = [[f"Row {r} Column {c}" for c in range(100)] for r in range(5000)]
        self.sheet_demo.data_reference(self.data)

        # __________ DISPLAY SUBSET OF COLUMNS __________

        #self.sheet_demo.display_subset_of_columns(indexes = [5, 7, 9, 11], enable = True)

        # __________ SETTING HEADERS __________

        #self.headers = [f"Header {c}" for c in range(100)]
        #self.sheet_demo.headers(self.headers)

        # __________ INSERTING A ROW __________

        #self.sheet_demo.insert_row(row = (f"my new row here {c}" for c in range(100)), idx = 0) # a filled row at the start
        #self.sheet_demo.insert_row() # an empty row at the end

        # __________ INSERTING A COLUMN __________

        #self.sheet_demo.insert_column(column = (f"my new col here {r}" for r in range(5000)), idx = 0) # a filled column at the start
        #self.sheet_demo.insert_column() # an empty column at the end

        # __________ HIDING THE ROW INDEX AND HEADERS __________

        #self.sheet_demo.hide("row_index")
        #self.sheet_demo.hide("top_left")
        #self.sheet_demo.hide("header")
        

        
app = demo()
app.mainloop()

```

----

### Light Theme

![alt text](https://i.imgur.com/PyukzmG.jpg)


### Dark Theme

![alt text](https://i.imgur.com/gU3rJgw.jpg)

----

