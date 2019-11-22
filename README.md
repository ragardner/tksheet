# tksheet 2.6

Python 3.6+ tkinter table widget for displaying data

```
pip install tksheet
```

Display and modify tabular data using Python 3.6+ and tkinter. Stores its display data as a Python list of lists, sublists being rows.

Documentation coming soon (hopefully before 2020)

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
        self.data = [[f"Row {r} Column {c}" for c in range(100)] for r in range(5000)]
        self.sdem = Sheet(self,
                          width = 1000,
                          height = 700,
                          align = "w",
                          header_align = "center",
                          row_index_align = "center",
                          show = True,
                          column_width = 180,
                          row_index_width = 50,
                          data_reference = self.data,
                          headers=[f"Header {c}" for c in range(100)])
        self.sdem.enable_bindings(("single",
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
                                   "double_click_row_resize"))
        self.sdem.edit_bindings(True)
        self.sdem.grid(row = 0, column = 0, sticky = "nswe")
        self.sdem.highlight_cells(row = 0, column = 0, bg = "orange", fg = "blue")
        self.sdem.highlight_cells(row = 0, bg = "orange", fg = "blue", canvas = "row_index")
        self.sdem.highlight_cells(column = 0, bg = "orange", fg = "blue", canvas = "header")


app = demo()
app.mainloop()
```

![alt text](https://i.imgur.com/PyukzmG.jpg)

