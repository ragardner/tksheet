from tksheet import Sheet
import tkinter as tk


class demo(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)
        self.grid_columnconfigure(0, weight = 1)
        self.grid_rowconfigure(0, weight = 1)
        self.sheet_demo = Sheet(self,
                                height = 500,
                                width = 700) #for full startup options see DOCUMENTATION.md
        self.sheet_demo.enable_bindings(("single_select", #"single_select" or "toggle_select"
                                         "drag_select",   #enables shift click selection as well
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
                                         "right_click_popup_menu",
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
        #self.sheet_demo.disable_bindings() #uses the same strings
        self.sheet_demo.grid(row = 0, column = 0, sticky = "nswe")
        

        """_________________________ EXAMPLES _________________________ """
        """_____________________________________________________________"""

        # __________ CHANGING THEME __________

        #self.sheet_demo.change_theme("dark")

        # __________ HIGHLIGHT / DEHIGHLIGHT CELLS __________

        self.sheet_demo.highlight_cells(row = 5, column = 0, bg = "#ed4337", fg = "white")
        self.sheet_demo.highlight_cells(row = 5, column = 5, bg = "#ed4337", fg = "white")
        self.sheet_demo.highlight_cells(row = 5, column = 1, bg = "#ed4337", fg = "white")
        self.sheet_demo.highlight_cells(row = 5, bg = "#ed4337", fg = "white", canvas = "row_index")
        self.sheet_demo.highlight_cells(column = 0, bg = "#ed4337", fg = "white", canvas = "header")

        # __________ SETTING OR RESETTING TABLE DATA __________
        
        self.data = [[f"Row {r} Column {c}" for c in range(30)] for r in range(200)]
        self.sheet_demo.data_reference(self.data)

        # __________ BINDING A FUNCTION TO USER SELECTS CELL __________

        self.sheet_demo.extra_bindings([
                                        ("cell_select", self.cell_select),
                                        ("shift_cell_select", self.shift_select_cells),
                                        ("drag_select_cells", self.drag_select_cells),
                                        ("ctrl_a", self.ctrl_a),
                                        ("row_select", self.row_select),
                                        ("shift_row_select", self.shift_select_rows),
                                        ("drag_select_rows", self.drag_select_rows),
                                        ("column_select", self.column_select),
                                        ("shift_column_select", self.shift_select_columns),
                                        ("drag_select_columns", self.drag_select_columns),
                                        ]
                                       )
        #self.sheet_demo.extra_bindings([("cell_select", None)]) #unbind cell select

        # __________ BINDING NEW RIGHT CLICK FUNCTION __________
    
        self.sheet_demo.bind("<3>", self.rc)

        # __________ DISPLAY SUBSET OF COLUMNS __________

        #self.sheet_demo.display_subset_of_columns(indexes = [3, 7, 9, 0], enable = True)

        # __________ SETTING HEADERS __________

        #self.sheet_demo.headers((f"Header {c}" for c in range(30))) #any iterable works
        #self.sheet_demo.headers("Change header example", 2)
        #print (self.sheet_demo.headers())
        #print (self.sheet_demo.headers(index = 2))

        # __________ SETTING ROW INDEX __________

        #self.sheet_demo.row_index((f"Row {r}" for r in range(200))) #any iterable works
        #self.sheet_demo.row_index("Change index example", 2)
        #print (self.sheet_demo.row_index())
        #print (self.sheet_demo.row_index(index = 2))

        # __________ INSERTING A ROW __________

        #self.sheet_demo.insert_row(row = (f"my new row here {c}" for c in range(100)), idx = 0) # a filled row at the start
        #self.sheet_demo.insert_row() # an empty row at the end

        # __________ INSERTING A COLUMN __________

        #self.sheet_demo.insert_column(column = (f"my new col here {r}" for r in range(5000)), idx = 0) # a filled column at the start
        #self.sheet_demo.insert_column() # an empty column at the end

        # __________ SETTING A COLUMNS DATA __________

        # any iterable works
        #self.sheet_demo.set_column_data(0, values = (0 for i in range(220)))

        # __________ SETTING A ROWS DATA __________

        # any iterable works
        #self.sheet_demo.set_row_data(0, values = (0 for i in range(35)))

        # __________ SETTING A CELLS DATA __________

        #self.sheet_demo.set_cell_data(1, 2, "NEW VALUE")

        # __________ GETTING FULL SHEET DATA __________

        #self.all_data = self.sheet_demo.get_sheet_data()

        # __________ GETTING CELL DATA __________

        #print (self.sheet_demo.get_cell_data(0, 0))

        # __________ GETTING ROW DATA __________

        #print (self.sheet_demo.get_row_data(0)) # only accessible by index

        # __________ GETTING COLUMN DATA __________

        #print (self.sheet_demo.get_column_data(0)) # only accessible by index

        # __________ HIDING THE ROW INDEX AND HEADERS __________

        #self.sheet_demo.hide("row_index")
        #self.sheet_demo.hide("top_left")
        #self.sheet_demo.hide("header")

    def rc(self, event):
        print (event)
        
    def cell_select(self, response):
        print (response)

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
