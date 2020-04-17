from tksheet import Sheet
import tkinter as tk

class Sheet_Listbox(Sheet):
    def __init__(self,
                 parent,
                 values = []):
        Sheet.__init__(self,
                       parent = parent,
                       show_horizontal_grid = False,
                       show_vertical_grid = False,
                       show_header = False,
                       show_row_index = False,
                       show_top_left = False,
                       empty_horizontal = 0,
                       empty_vertical = 0)
        if values:
            self.values(values)

    def values(self, values = []):
        self.set_sheet_data([[v] for v in values],
                            reset_col_positions = False,
                            reset_row_positions = False,
                            redraw = False,
                            verify = False)
        self.set_all_cell_sizes_to_text()
        

class demo(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)
        self.grid_columnconfigure(0,
                                  weight = 1)
        self.grid_rowconfigure(0,
                               weight = 1)
        self.listbox = Sheet_Listbox(self,
                                     values = [f"_________  Item {i}  _________" for i in range(2000)])
        self.listbox.grid(row = 0,
                          column = 0,
                          sticky = "nswe")
        #self.listbox.values([f"new values {i}" for i in range(50)]) set values

        
app = demo()
app.mainloop()
