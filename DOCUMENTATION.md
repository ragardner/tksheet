# Table of Contents
1. [About tksheet](https://github.com/ragardner/tksheet/blob/master/DOCUMENTATION.md#1-About-tksheet)
2. [Installation and Requirements](https://github.com/ragardner/tksheet/blob/master/DOCUMENTATION.md#2-Installation-and-Requirements)
3. [Basic Initialization](https://github.com/ragardner/tksheet/blob/master/DOCUMENTATION.md#3-Basic-Initialization)
4. [Initialization Options](https://github.com/ragardner/tksheet/blob/master/DOCUMENTATION.md#4-Initialization-Options)
5. [Modifying Table Data and Dimensions](https://github.com/ragardner/tksheet/blob/master/DOCUMENTATION.md#5-Modifying-Table-Data-and-Dimensions)
6. [Getting Table Data](https://github.com/ragardner/tksheet/blob/master/DOCUMENTATION.md#6-Getting-Table-Data)
7. [Bindings and Functionality](https://github.com/ragardner/tksheet/blob/master/DOCUMENTATION.md#7-Bindings-and-Functionality)
8. [Identifying Bound Event Mouse Position](https://github.com/ragardner/tksheet/blob/master/DOCUMENTATION.md#8-Identifying-Bound-Event-Mouse-Position)
9. [Table Colors](https://github.com/ragardner/tksheet/blob/master/DOCUMENTATION.md#9-Table-Colors)
10. [Highlighting Cells](https://github.com/ragardner/tksheet/blob/master/DOCUMENTATION.md#10-Highlighting-Cells)
11. [Text Font and Alignment](https://github.com/ragardner/tksheet/blob/master/DOCUMENTATION.md#11-Text-Font-and-Alignment)
12. [Row Heights and Column Widths](https://github.com/ragardner/tksheet/blob/master/DOCUMENTATION.md#12-Row-Heights-and-Column-Widths)
13. [Getting Selected Cells](https://github.com/ragardner/tksheet/blob/master/DOCUMENTATION.md#13-Getting-Selected-Cells)
14. [Modifying Selected Cells](https://github.com/ragardner/tksheet/blob/master/DOCUMENTATION.md#14-Modifying-Selected-Cells)
15. [Modifying and Getting Scroll Positions](https://github.com/ragardner/tksheet/blob/master/DOCUMENTATION.md#15-Modifying-and-Getting-Scroll-Positions)
16. [Readonly Cells](https://github.com/ragardner/tksheet/blob/master/DOCUMENTATION.md#16-Readonly-Cells)
17. [Hiding Columns](https://github.com/ragardner/tksheet/blob/master/DOCUMENTATION.md#17-Hiding-Columns)
18. [Hiding Table Elements](https://github.com/ragardner/tksheet/blob/master/DOCUMENTATION.md#18-Hiding-Table-Elements)
19. [Cell Text Editor](https://github.com/ragardner/tksheet/blob/master/DOCUMENTATION.md#19-Cell-Text-Editor)
20. [Dropdown Boxes](https://github.com/ragardner/tksheet/blob/master/DOCUMENTATION.md#20-Dropdown-Boxes)
21. [Check Boxes](https://github.com/ragardner/tksheet/blob/master/DOCUMENTATION.md#21-Check-Boxes)
22. [Table Options and Other Functions](https://github.com/ragardner/tksheet/blob/master/DOCUMENTATION.md#22-Table-Options-and-Other-Functions)
23. [Example Loading Data from Excel](https://github.com/ragardner/tksheet/blob/master/DOCUMENTATION.md#23-Example-Loading-Data-from-Excel)
24. [Example Custom Right Click and Text Editor Functionality](https://github.com/ragardner/tksheet/blob/master/DOCUMENTATION.md#24-Example-Custom-Right-Click-and-Text-Editor-Functionality)
25. [Example Displaying Selections](https://github.com/ragardner/tksheet/blob/master/DOCUMENTATION.md#25-Example-Displaying-Selections)
26. [Example List Box](https://github.com/ragardner/tksheet/blob/master/DOCUMENTATION.md#26-Example-List-Box)


## 1 About tksheet
`tksheet` is a Python tkinter table widget written in pure python. It is licensed under the [MIT license](https://github.com/ragardner/tksheet/blob/master/LICENSE.txt).

It works using tkinter canvases to draw and moves lines and text for only the visible portion of the table.

Cell values can be any class with a `str` method.

See the [tests](https://github.com/ragardner/tksheet/tree/master/tests) folder for more examples.

Examples of things that are not possible with tksheet:
 - Cell merging
 - Cell text wrap
 - Changing font for individual cells
 - Different fonts for index and table
 - Mouse drag copy cells

## 2 Installation and Requirements
`tksheet` is available through PyPi (Python package index) and can be installed by using Pip through the command line `pip install tksheet`

Alternatively you can download the source code and (inside the tksheet directory) use the command line `python setup.py develop`

`tksheet` requires a Python version of `3.6` or higher.

## 3 Basic Initialization
Like other tkinter widgets you need only the `Sheet()`s parent as an argument to initialize a `Sheet()` e.g.
```python
sheet = Sheet(my_frame_widget)
```
 - `my_frame_widget` would be replaced by whatever widget is your `Sheet()`s parent.

___

As an example, this is a tkinter program involving a `Sheet()` widget
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
                           data = [[f"Row {r}, Column {c}\nnewline1\nnewline2" for c in range(50)] for r in range(500)])
        self.sheet.enable_bindings()
        self.frame.grid(row = 0, column = 0, sticky = "nswe")
        self.sheet.grid(row = 0, column = 0, sticky = "nswe")


app = demo()
app.mainloop()
```

## 4 Initialization Options
This is a full list of all the start up arguments, the only required argument is the sheets parent, everything else has default arguments.

```python
Sheet(
parent,
show_table = True,
show_top_left = True,
show_row_index = True,
show_header = True,
show_x_scrollbar = True,
show_y_scrollbar = True,
width = None,
height = None,
headers = None,
measure_subset_header = True,
default_header = "letters", #letters, numbers or both
default_row_index = "numbers", #letters, numbers or both
page_up_down_select_row = True,
expand_sheet_if_paste_too_big = False,
paste_insert_column_limit = None,
paste_insert_row_limit = None,
arrow_key_down_right_scroll_page = False,
enable_edit_cell_auto_resize = True,
data_reference = None,
data = None,
startup_select = None,
startup_focus = True,
total_columns = None,
total_rows = None,
column_width = 120,
header_height = "1",
max_colwidth = "inf",
max_rh = "inf",
max_header_height = "inf",
max_row_width = "inf",
row_index = None,
measure_subset_index = True,
after_redraw_time_ms = 100,
row_index_width = 100,
auto_resize_default_row_index = True,
set_all_heights_and_widths = False,
row_height = "1",
font = get_font(),
header_font = get_heading_font(),
popup_menu_font = get_font(),
align = "w",
header_align = "center",
row_index_align = "center",
displayed_columns = [],
all_columns_displayed = True,
max_undos = 20,
outline_thickness = 0,
outline_color = theme_light_blue['outline_color'],
column_drag_and_drop_perform = True,
row_drag_and_drop_perform = True,
empty_horizontal = 150,
empty_vertical = 100,
selected_rows_to_end_of_window = False,
horizontal_grid_to_end_of_window = False,
vertical_grid_to_end_of_window = False,
show_vertical_grid = True,
show_horizontal_grid = True,
display_selected_fg_over_highlights = False,
show_selected_cells_border = True,
theme                              = "light blue",
popup_menu_fg                      = "gray2",
popup_menu_bg                      = "#f2f2f2",
popup_menu_highlight_bg            = "#91c9f7",
popup_menu_highlight_fg            = "black",
frame_bg                           = theme_light_blue['table_bg'],
table_grid_fg                      = theme_light_blue['table_grid_fg'],
table_bg                           = theme_light_blue['table_bg'],
table_fg                           = theme_light_blue['table_fg'], 
table_selected_cells_border_fg     = theme_light_blue['table_selected_cells_border_fg'],
table_selected_cells_bg            = theme_light_blue['table_selected_cells_bg'],
table_selected_cells_fg            = theme_light_blue['table_selected_cells_fg'],
table_selected_rows_border_fg      = theme_light_blue['table_selected_rows_border_fg'],
table_selected_rows_bg             = theme_light_blue['table_selected_rows_bg'],
table_selected_rows_fg             = theme_light_blue['table_selected_rows_fg'],
table_selected_columns_border_fg   = theme_light_blue['table_selected_columns_border_fg'],
table_selected_columns_bg          = theme_light_blue['table_selected_columns_bg'],
table_selected_columns_fg          = theme_light_blue['table_selected_columns_fg'],
resizing_line_fg                   = theme_light_blue['resizing_line_fg'],
drag_and_drop_bg                   = theme_light_blue['drag_and_drop_bg'],
index_bg                           = theme_light_blue['index_bg'],
index_border_fg                    = theme_light_blue['index_border_fg'],
index_grid_fg                      = theme_light_blue['index_grid_fg'],
index_fg                           = theme_light_blue['index_fg'],
index_selected_cells_bg            = theme_light_blue['index_selected_cells_bg'],
index_selected_cells_fg            = theme_light_blue['index_selected_cells_fg'],
index_selected_rows_bg             = theme_light_blue['index_selected_rows_bg'],
index_selected_rows_fg             = theme_light_blue['index_selected_rows_fg'],
index_hidden_rows_expander_bg      = theme_light_blue['index_hidden_rows_expander_bg'],
header_bg                          = theme_light_blue['header_bg'],
header_border_fg                   = theme_light_blue['header_border_fg'],
header_grid_fg                     = theme_light_blue['header_grid_fg'],
header_fg                          = theme_light_blue['header_fg'],
header_selected_cells_bg           = theme_light_blue['header_selected_cells_bg'],
header_selected_cells_fg           = theme_light_blue['header_selected_cells_fg'],
header_selected_columns_bg         = theme_light_blue['header_selected_columns_bg'],
header_selected_columns_fg         = theme_light_blue['header_selected_columns_fg'],
header_hidden_columns_expander_bg  = theme_light_blue['header_hidden_columns_expander_bg'],
top_left_bg                        = theme_light_blue['top_left_bg'],
top_left_fg                        = theme_light_blue['top_left_fg'],
top_left_fg_highlight              = theme_light_blue['top_left_fg_highlight'])
```
 - `startup_select` selects cells, rows or columns at initialization by using a `tuple` e.g. `(0, 0, "cells")` for cell A0 or `(0, 5, "rows")` for rows 0 to 5.
 - `data_reference` and `data` are essentially the same

You can change these settings after initialization using the `set_options()` function.

## 5 Modifying Table Data and Dimensions

Set sheet data, overwrites any existing data.
```python
set_sheet_data(data = [[]],
               reset_col_positions = True,
               reset_row_positions = True,
               redraw = True,
               verify = False,
               reset_highlights = False)
```
 - `data` (`list`) has to be a list of lists for full functionality, for display only a list of tuples or a tuple of tuples will work.
 - `reset_col_positions` and `reset_row_positions` (`bool`) when `True` will reset column widths and row heights.
 - `redraw` (`bool`) refreses the table after setting new data.
 - `verify` (`bool`) goes through `data` and checks if it is a list of lists, will raise error if not, disabled by default.
 - `reset_highlights` (`bool`) resets all table cell highlights.

___

Set cell data, overwrites any existing data.
```python
set_cell_data(r, c, value = "", set_copy = True, redraw = False)
```
 - `set_copy` means `str()` will be used on the value before setting.

___

Insert a row into the sheet.
```python
insert_row(values = None, idx = "end", height = None, deselect_all = False, add_columns = False,
           redraw = False)
```
 - Leaving `values` as `None` inserts an empty row, e.g. `insert_row()` will append an empty row to the sheet.
 - `height` is the new rows displayed height in pixels, leave as `None` for default.
 - `add_columns` checks the rest of the sheets rows are at least the length as the new row, leave as `False` for better performance.

___

Set column data, overwrites any existing data.
```python
set_column_data(c, values = tuple(), add_rows = True, redraw = False)
```
 - `add_rows` adds extra rows to the sheet if the column data doesn't fit within current sheet dimensions.

___

Set the header to something non-default (if new header is shorter than total columns then default headers e.g. letters will be used on the end.
```python
headers(newheaders = None, index = None, reset_col_positions = False, show_headers_if_not_sheet = True)
```
 - Using an integer `int` for argument `newheaders` makes the sheet use that row as a header e.g. `headers(0)` means the first row will be used as a header (the first row will not be hidden in the sheet though).
 - Leaving `newheaders` as `None` and using the `index` argument returns the existing header value in that index.
 - Leaving all arguments as default e.g. `headers()` returns existing headers.

___

Set the index to something non-default (if new index is shorter than total rows then default index e.g. numbers will be used on the end.
```python
row_index(newindex = None, index = None, reset_row_positions = False, show_index_if_not_sheet = True)
```
 - Leaving `newindex` as `None` and using the `index` argument returns the existing row index value in that index.
 - Leaving all arguments as default e.g. `row_index()` returns the existing row index.

___

Insert a column into the sheet.
```python
insert_column(values = None, idx = "end", width = None, deselect_all = False, add_rows = True, equalize_data_row_lengths = True,
              mod_column_positions = True,
              redraw = False)
```

___

Insert multiple columns into the sheet.
```python
insert_columns(columns = 1, idx = "end", widths = None, deselect_all = False, add_rows = True, equalize_data_row_lengths = True,
               mod_column_positions = True,
               redraw = False)
```
 - `columns` can be either `int` or iterable of iterables.

___

Set row data, overwrites any existing data.
```python
set_row_data(r, values = tuple(), add_columns = True, redraw = False)
```

___

Insert multiple rows into the sheet.
```python
insert_rows(rows = 1, idx = "end", heights = None, deselect_all = False, add_columns = True,
            redraw = False)
```
 - `rows` can be either `int` or iterable of iterables.

___

```python
sheet_data_dimensions(total_rows = None, total_columns = None)
```

___

```python
delete_row(idx = 0, deselect_all = False, redraw = True)
```

___

```python
total_rows(number = None, mod_positions = True, mod_data = True)
```

___

```python
total_columns(number = None, mod_positions = True, mod_data = True)
```

___

```python
set_sheet_data_and_display_dimensions(total_rows = None, total_columns = None)
```

___

```python
move_row(row, moveto)
```

___

```python
delete_column(idx = 0, deselect_all = False, redraw = True)
```

___

```python
move_column(column, moveto)
```

___

Make all data rows the same length (same number of columns), goes by longest row. This will only affect the data variable, not visible columns.
```python
equalize_data_row_lengths()
```

___

Modify widget height and width in pixels
```python
height_and_width(height = None, width = None)
```
 - `height` (`int`) set a height in pixels
 - `width` (`int`) set a width in pixels
If both arguments are `None` then table will reset to default tkinter canvas dimensions.

## 6 Getting Table Data

Get sheet data and, if required, header and index data.
```python
get_sheet_data(return_copy = False, get_header = False, get_index = False)
```
 - `return_copy` (`bool`) will copy all cells if `True`, also copies header and index if they are `True`.
 - `get_header` (`bool`) will put the header as the first row if `True`.
 - `get_index` (`bool`) will put index items as the first item in every row.

___

```python
get_cell_data(r, c, return_copy = True)
```

___

```python
get_row_data(r, return_copy = True)
```

___

```python
get_column_data(c, return_copy = True)
```

___

Get number of rows in table data.
```python
get_total_rows()
```

## 7 Bindings and Functionality

Enable table functionality and bindings.
```python
enable_bindings(bindings = "all")
```
 - `bindings` (`str`) options are (rc stands for right click):
	- "all"
	- "single_select"
	- "toggle_select"
	- "drag_select"
        - "select_all"
	- "column_drag_and_drop"
	- "row_drag_and_drop"
	- "column_select"
	- "row_select"
	- "column_width_resize"
	- "double_click_column_resize"
	- "row_width_resize"
	- "column_height_resize"
	- "arrowkeys"
	- "row_height_resize"
	- "double_click_row_resize"
	- "right_click_popup_menu"
	- "rc_select"
	- "rc_insert_column"
	- "rc_delete_column"
	- "rc_insert_row"
	- "rc_delete_row"
	- "hide_columns"
	- "copy"
	- "cut"
	- "paste"
	- "delete"
	- "undo"
	- "edit_cell"

To allow table expansion when pasting data which doesn't fit in the table use either:
 - `expand_sheet_if_paste_too_big = True` in sheet initialization arguments or
 - `sheet.set_options(expand_sheet_if_paste_too_big = True)`

___

Disable table functionality and bindings, uses the same arguments as `enable_bindings()`
```python
disable_bindings(bindings = "all")
```

___

Bind various table functionality to your own functions. To unbind a function either set `func` argument to `None` or leave it as default e.g. `extra_bindings("begin_copy")` to unbind `"begin_copy"`.
```python
extra_bindings(bindings, func = "None")
```
Notes:
 - Upon an event being triggered the bound function will be sent a [namedtuple](https://docs.python.org/3/library/collections.html#collections.namedtuple) containing variables relevant to that event, use `print()` or similar to see all the variable names in the event. Each event contains different variable names with the exception of `eventname` e.g. `event.eventname`

Arguments:
 - `bindings` (`str`) options are:
	- "begin_copy"
	- "end_copy"
	- "begin_cut"
	- "end_cut"
	- "begin_paste"
	- "end_paste"
	- "begin_undo"
	- "end_undo"
	- "begin_delete"
	- "end_delete"
	- "begin_edit_cell"
	- "end_edit_cell"
	- "begin_row_index_drag_drop"
	- "end_row_index_drag_drop"
	- "begin_column_header_drag_drop"
	- "end_column_header_drag_drop"
	- "begin_delete_rows"
	- "end_delete_rows"
	- "begin_delete_columns"
	- "end_delete_columns"
	- "begin_insert_columns"
	- "end_insert_columns"
	- "begin_insert_rows"
	- "end_insert_rows"
    - "row_height_resize"
    - "column_width_resize"
	- "cell_select"
	- "select_all"
	- "row_select"
	- "column_select"
	- "drag_select_cells"
	- "drag_select_rows"
	- "drag_select_columns"
	- "shift_cell_select"
	- "shift_row_select"
	- "shift_column_select"
	- "deselect"
	- "all_select_events"
	- "bind_all"
	- "unbind_all"
 - `func` argument is the function you want to send the binding event to.

___

Add commands to the in-built right click popup menu.
```python
popup_menu_add_command(label, func, table_menu = True, index_menu = True, header_menu = True)
```

___

Remove the custom commands added using the above function from the in-built right click popup menu, if `label` is `None` then it removes all.
```python
popup_menu_del_command(label = None)
```

___

Disable table functionality and bindings (uses the same options as `enable_bindings()`.
```python
disable_bindings(bindings = "all")
```

___

Enable or disable mousewheel, left click etc.
```python
basic_bindings(enable = False)
```

___

Enable or disable cell edit functionality, including Undo.
```python
edit_bindings(enable = False)
```

___

Enable or disable the ability to edit a specific cell.
```python
cell_edit_binding(enable = False, keys = [])
```
 - `keys` can be used to bind more keys to open a cell edit window

___

```python
bind(binding, func, add = None)
```
 - `add` will only work for bindings which are not the following: `"<ButtonPress-1>"`, `"<ButtonMotion-1>"`, `"<ButtonRelease-1>"`, `"<Double-Button-1>"`, `"<Motion>"` and lastly whichever is your operating systems right mouse click button
___

```python
unbind(binding)
```

___

```python
cut(event = None)
copy(event = None)
paste(event = None)
delete(event = None)
undo(event = None)
```

## 8 Identifying Bound Event Mouse Position

The below functions require a mouse click event, for example you could bind right click, example [here](https://github.com/ragardner/tksheet/wiki#24-example-custom-right-click-and-text-editor-functionality), and then identify where the user has clicked.

___

```python
identify_region(event)
```

___

```python
identify_row(event, exclude_index = False, allow_end = True)
```

___

```python
identify_column(event, exclude_header = False, allow_end = True)
```

___

Sheet control actions for binding your own keys to e.g. `sheet.bind("<Control-B>", sheet.paste)`
```python
cut(self, event = None)
copy(self, event = None)
paste(self, event = None)
delete(self, event = None)
undo(self, event = None)
edit_cell(self, event = None, dropdown = False)
```

## 9 Table Colors

For specific colors use the function `set_options()`, arguments can be found [here](https://github.com/ragardner/tksheet/wiki#22-table-options-and-other-functions).

Otherwise you can change the theme using the below function.
```python
change_theme(theme = "light blue")
```
 - `theme` (`str`) options (themes) are `light blue`, `light green`, `dark blue` and `dark green`.

## 10 Highlighting Cells

 - `bg` and `fg` arguments use either a tkinter color or a hex `str` color.
 - Highlighting cells, rows or columns will also change the colors of dropdown boxes and check boxes.

___

```python
highlight_cells(row = 0, column = 0, cells = [], canvas = "table", bg = None, fg = None, redraw = False)
```

___

```python
dehighlight_cells(row = 0, column = 0, cells = [], canvas = "table", all_ = False, redraw = True)
```

___

```python
get_highlighted_cells(canvas = "table")
```

___

```python
highlight_rows(rows = [], bg = None, fg = None, highlight_index = True, redraw = False, end_of_screen = False)
```
 - `end_of_screen` when `True` makes the row highlight go past the last column line if there is any room there.

___

```python
highlight_columns(columns = [], bg = None, fg = None, highlight_header = True, redraw = False)
```

___

```python
dehighlight_all()
```

___

```python
dehighlight_rows(rows = [], redraw = False)
```

___

```python
dehighlight_columns(columns = [], redraw = False)
```

## 11 Text Font and Alignment

 - `newfont` arguments require a three tuple e.g. `("Arial", 12, "normal")`
 - `align` arguments (`str`) options are `w`, `e` or `center`.

___

```python
font(newfont = None, reset_row_positions = True)
```

___

```python
header_font(newfont = None)
```

___

```python
align(align = None, redraw = True)
```

___

```python
header_align(align = None, redraw = True)
```

___

```python
row_index_align(align = None, redraw = True)
```

___

Change the text alignment for **specific** rows, `"global"` resets to table setting.
```python
align_rows(rows = [], align = "global", align_index = False, redraw = True)
```

___

Change the text alignment for **specific** columns, `"global"` resets to table setting.
```python
align_columns(columns = [], align = "global", align_header = False, redraw = True)
```

___

Change the text alignment for **specific** cells inside the table, `"global"` resets to table setting.
```python
align_cells(row = 0, column = 0, cells = [], align = "global", redraw = True)
```

___

Change the text alignment for **specific** cells inside the header, `"global"` resets to header setting.
```python
align_header(columns = [], align = "global", redraw = True)
```

___

Change the text alignment for **specific** cells inside the index, `"global"` resets to index setting.
```python
align_index(rows = [], align = "global", redraw = True)
```

## 12 Row Heights and Column Widths

Set default column width in pixels.
```python
default_column_width(width = None)
```
 - `width` (`int`).

___

Set default row height in pixels or lines.
```python
default_row_height(height = None)
```
 - `height` (`int`, `str`) use a numerical `str` for number of lines e.g. `"3"` for a height that fits 3 lines or `int` for pixels.

___

Set default header bar height in pixels or lines.
```python
default_header_height(height = None)
```
 - `height` (`int`, `str`) use a numerical `str` for number of lines e.g. `"3"` for a height that fits 3 lines or `int` for pixels.

___
Set a specific cell size to its text
```python
set_cell_size_to_text(row, column, only_set_if_too_small = False, redraw = True)
```

___

Set all row heights and column widths to cell text sizes.
```python
set_all_cell_sizes_to_text(redraw = True)
```

___

Get the sheets column widths.
```python
get_column_widths(canvas_positions = False)
```
 - `canvas_positions` (`bool`) gets the actual canvas x coordinates of column lines.

___

Get the sheets row heights.
```python
get_row_heights(canvas_positions = False)
```
 - `canvas_positions` (`bool`) gets the actual canvas y coordinates of row lines.

___

Set all column widths to specific `width` in pixels (`int`) or leave `None` to set to cell text sizes for each column.
```python
set_all_column_widths(width = None, only_set_if_too_small = False, redraw = True, recreate_selection_boxes = True)
```

___

Set all row heights to specific `height` in pixels (`int`) or leave `None` to set to cell text sizes for each row.
```python
set_all_row_heights(height = None, only_set_if_too_small = False, redraw = True, recreate_selection_boxes = True)
```

___

Set a specific column width.
```python
column_width(column = None, width = None, only_set_if_too_small = False, redraw = True)
```

___

```python
set_column_widths(column_widths = None, canvas_positions = False, reset = False, verify = False)
```

___

```python
set_width_of_index_to_text(recreate = True)
```

___

```python
row_height(row = None, height = None, only_set_if_too_small = False, redraw = True)
```

___

```python
set_row_heights(row_heights = None, canvas_positions = False, reset = False, verify = False)
```

___

```python
delete_row_position(idx, deselect_all = False)
```

___

```python
insert_row_position(idx = "end", height = None, deselect_all = False, redraw = False)
```

___

```python
insert_row_positions(idx = "end", heights = None, deselect_all = False, redraw = False)
```

___

```python
sheet_display_dimensions(total_rows = None, total_columns = None)
```

___

```python
move_row_position(row, moveto)
```

___

```python
delete_column_position(idx, deselect_all = False)
```

___

```python
insert_column_position(idx = "end", width = None, deselect_all = False, redraw = False)
```

___

```python
insert_column_positions(idx = "end", widths = None, deselect_all = False, redraw = False)
```

___

```python
move_column_position(column, moveto)
```

___

```python
get_example_canvas_column_widths(total_cols = None)
```

___

```python
get_example_canvas_row_heights(total_rows = None)
```

___

```python
verify_row_heights(row_heights, canvas_positions = False)
```

___

```python
verify_column_widths(column_widths, canvas_positions = False)
```

## 13 Getting Selected Cells

```python
get_currently_selected(get_coords = False, return_nones_if_not = False)
```

___

```python
get_selected_rows(get_cells = False, get_cells_as_rows = False, return_tuple = False)
```

___

```python
get_selected_columns(get_cells = False, get_cells_as_columns = False, return_tuple = False)
```

___

```python
get_selected_cells(get_rows = False, get_columns = False, sort_by_row = False, sort_by_column = False)
```

___

```python
get_all_selection_boxes()
```

___

```python
get_all_selection_boxes_with_types()
```

___

Check if cell is selected, returns `bool`.
```python
cell_selected(r, c)
```

___

Check if row is selected, returns `bool`.
```python
row_selected(r)
```

___

Check if column is selected, returns `bool`.
```python
column_selected(c)
```

___

Check if any cells, rows or columns are selected, there are options for exclusions, returns `bool`.
```python
anything_selected(exclude_columns = False, exclude_rows = False, exclude_cells = False)
```

___

Check if user has entire table selected, returns `bool`.
```python
all_selected()
```

___

```python
get_ctrl_x_c_boxes()
```

___

```python
get_selected_min_max()
```
 - returns `(min_y, min_x, max_y, max_x)` of any selections including rows/columns.

## 14 Modifying Selected Cells

```python
set_currently_selected(current_tuple_0 = 0, current_tuple_1 = 0, selection_binding = True)
```

___

```python
select_row(row, redraw = True)
```

___

```python
select_column(column, redraw = True)
```

___

```python
select_cell(row, column, redraw = True)
```

___

```python
select_all(redraw = True, run_binding_func = True)
```

___

```python
move_down()
```

___

```python
add_cell_selection(row, column, redraw = True, run_binding_func = True, set_as_current = True)
```

___

```python
add_row_selection(row, redraw = True, run_binding_func = True, set_as_current = True)
```

___

```python
add_column_selection(column, redraw = True, run_binding_func = True, set_as_current = True)
```

___

```python
toggle_select_cell(row, column, add_selection = True, redraw = True, run_binding_func = True, set_as_current = True)
```

___

```python
toggle_select_row(row, add_selection = True, redraw = True, run_binding_func = True, set_as_current = True)
```

___

```python
toggle_select_column(column, add_selection = True, redraw = True, run_binding_func = True, set_as_current = True)
```

___

```python
create_selection_box(r1, c1, r2, c2, type_ = "cells")
```

___

```python
recreate_all_selection_boxes()
```

___

```python
deselect(row = None, column = None, cell = None, redraw = True)
```

## 15 Modifying and Getting Scroll Positions

```python
see(row = 0, column = 0, keep_yscroll = False, keep_xscroll = False, bottom_right_corner = False, check_cell_visibility = True)
```

___

```python
set_xview(position, option = "moveto")
```

___

```python
set_yview(position, option = "moveto")
```

___

```python
get_xview()
```

___

```python
get_yview()
```

___

```python
set_view(x_args, y_args)
```

___

```python
move_down()
```

## 16 Readonly Cells

```python
readonly_rows(rows = [], readonly = True, redraw = True)
```

___

```python
readonly_columns(columns = [], readonly = True, redraw = True)
```

___

```python
readonly_cells(row = 0, column = 0, cells = [], readonly = True, redraw = True)
```

## 17 Hiding Columns

Display only certain columns.
```python
display_columns(indexes = None,
                enable = None,
                reset_col_positions = True,
                set_col_positions = True,
                refresh = False,
                redraw = False,
                deselect_all = True)
```
 - If the chosen indexes are equal to a list of columns as long as the longest row in the sheets data then enable will be set to `False`.

## 18 Hiding Table Elements

Hide parts of the table or all of it
```python
hide(canvas = "all")
```
 - `canvas` (`str`) options are `all`, `row_index`, `header`, `top_left`, `x_scrollbar`, `y_scrollbar`
	- `all` hides the entire table and is the default.

___

Show parts of the table or all of it
```python
show(canvas = "all")
```
 - `canvas` (`str`) options are `all`, `row_index`, `header`, `top_left`, `x_scrollbar`, `y_scrollbar`
	- `all` shows the entire table and is the default.

## 19 Cell Text Editor

```python
create_text_editor(row = 0, column = 0, text = None, state = "normal", see = True, set_data_ref_on_destroy = False,
                           binding = None)
```

___

```python
set_text_editor_value(text = "", r = None, c = None)
```

___

```python
bind_text_editor_set(func, row, column)
```

___

```python
get_text_editor_value(destroy_tup = None, r = None, c = None, set_data_ref_on_destroy = True, event = None, destroy = True, move_down = True, redraw = True, recreate = True)
```

___

```python
destroy_text_editor(event = None)
```

___

```python
get_text_editor_widget(event = None)
```

___

```python
bind_key_text_editor(key, function)
```

___

```python
unbind_key_text_editor(key)
```

## 20 Dropdown Boxes

Create a dropdown box (only creates the arrow and border and sets it up for usage, does not pop open the box).
```python
def create_dropdown(r = 0,
                    c = 0,
                    values = [],
                    set_value = None,
                    state = "readonly",
                    redraw = False,
                    selection_function = None,
                    modified_function = None)
```
Notes:
 - When a user selects an item from the dropdown box the sheet will set the underlying cells data to the selected item, to bind this event use either the `selection_function` argument or see the function `extra_bindings()` with binding `"end_edit_cell"` [here](https://github.com/ragardner/tksheet/wiki#7-bindings-and-functionality).

 Arguments:
 - `values` are the values to appear when the dropdown box is popped open.
 - `state` determines whether or not there is also an editable text window at the top of the dropdown box when it is open.
 - `redraw` refreshes the sheet so the newly created box is visible.
 - `selection_function` can be used to trigger a specific function when an item from the dropdown box is selected, if you are using the above `extra_bindings()` as well it will also be triggered but after this function. e.g. `selection_function = my_function_name`
 - `modified_function` can be used to trigger a specific function when the `state` of the box is set to `"normal"` and there is an editable text window and a change of the text in that window has occurred.

___

Get chosen dropdown boxes values.
```python
get_dropdown_values(r = 0, c = 0)
```

___

Set the values and displayed value of a chosen dropdown box.
```python
set_dropdown_values(r = 0, c = 0, set_existing_dropdown = False, values = [], displayed = None)
```
 - `set_existing_dropdown` if `True` takes priority over `r` and `c` and sets the values of the last popped open dropdown box (if one one is popped open, if not then an `Exception` is raised).
 - `values` (`list`, `tuple`)
 - `displayed` (`str`, `None`) if not `None` will try to set the displayed value of the chosen dropdown box to given argument.

___

Set and get bound dropdown functions.
```python
dropdown_functions(r, c, selection_function = "", modified_function = "")
```

___

Delete dropdown boxes.
```python
delete_dropdown(r = 0, c = 0)
```
 - Set `r` to `"all"` to delete all dropdown boxes on the sheet.

___

Get a dictionary of all dropdown boxes; keys: `(row int, column int)` and values: `(ttk combobox widget, tk canvas window object)`
```python
get_dropdowns()
```

___

Pop open a dropdown box.
```python
open_dropdown(r, c)
```

___

Close an already open dropdown box.
```python
close_dropdown(r, c)
```
 - Also destroys any opened text editor windows.

## 21 Check Boxes

Create a check box.
```python
create_checkbox(r,
                c,
                checked = False,
                state = "normal",
                redraw = False,
                check_function = None)
```
Notes:
 - Use `highlight_cells()` or rows or columns to change the color of the checkbox.
 - Check boxes are always left aligned despite any align settings.

 Arguments:
 - `check_function` can be used to trigger a function when the user clicks a checkbox.
 - `state` can be `"normal"` or `"disabled"`. If `"disabled"` then color will be same as table grid lines, else it will be the cells text color.

___

Set or toggle a checkbox.
```python
click_checkbox(r, c, checked = None)
```

___

Get a dictionary of all check box dictionaries.
```python
get_checkboxes()
```

___

Delete a checkbox.
```python
delete_checkbox(r = 0, c = 0)
```
 - Set `r` to `"all"` to delete all check boxes.

___

Set or get information about a particular checkbox.
```python
checkbox(r,
         c,
         checked = None,
         state = None,
         check_function = "")
```

## 22 Table Options and Other Functions

```python
def set_options(
enable_edit_cell_auto_resize = None,
selected_rows_to_end_of_window = None,
horizontal_grid_to_end_of_window = None,
vertical_grid_to_end_of_window = None,
page_up_down_select_row = None,
expand_sheet_if_paste_too_big = None,
paste_insert_column_limit = None,
paste_insert_row_limit = None,
arrow_key_down_right_scroll_page = None,
display_selected_fg_over_highlights = None,
empty_horizontal = None,
empty_vertical = None,
show_horizontal_grid = None,
show_vertical_grid = None,
top_left_fg_highlight = None,
auto_resize_default_row_index = None,
font = None,
default_header = None,
default_row_index = None,
header_font = None,
show_selected_cells_border = None,
theme = None,
max_colwidth = None,
max_row_height = None,
max_header_height = None,
max_row_width = None,
header_height = None,
row_height = None,
column_width = None,
header_bg = None,
header_border_fg = None,
header_grid_fg = None,
header_fg = None,
header_selected_cells_bg = None,
header_selected_cells_fg = None,
header_hidden_columns_expander_bg = None,
index_bg = None,
index_border_fg = None,
index_grid_fg = None,
index_fg = None,
index_selected_cells_bg = None,
index_selected_cells_fg = None,
index_hidden_rows_expander_bg = None,
top_left_bg = None,
top_left_fg = None,
frame_bg = None,
table_bg = None,
table_grid_fg = None,
table_fg = None,
table_selected_cells_border_fg = None,
table_selected_cells_bg = None,
table_selected_cells_fg = None,
resizing_line_fg = None,
drag_and_drop_bg = None,
outline_thickness = None,
outline_color = None,
header_selected_columns_bg = None,
header_selected_columns_fg = None,
index_selected_rows_bg = None,
index_selected_rows_fg = None,
table_selected_rows_border_fg = None,
table_selected_rows_bg = None,
table_selected_rows_fg = None,
table_selected_columns_border_fg = None,
table_selected_columns_bg = None,
table_selected_columns_fg = None,
popup_menu_font = None,
popup_menu_fg = None,
popup_menu_bg = None,
popup_menu_highlight_bg = None,
popup_menu_highlight_fg = None,
row_drag_and_drop_perform = None,
column_drag_and_drop_perform = None,
measure_subset_index = None,
measure_subset_header = None,
redraw = True)
```

___

Get internal storage dictionary of highlights, readonly cells, dropdowns etc.
```python
get_cell_options(canvas = "table")
```

___

```python
get_frame_y(y)
```

___

```python
get_frame_x(x)
```

___

Flash a dashed box of chosen dimensions.
```python
show_ctrl_outline(canvas = "table", start_cell = (0, 0), end_cell = (1, 1))
```

___

Reset table undo storage.
```python
reset_undos()
```

___

Refresh the table.
```python
redraw(redraw_header = True, redraw_row_index = True)
```

___

Refresh the table.
```python
refresh(redraw_header = True, redraw_row_index = True)
```

## 23 Example Loading Data from Excel

Using `pandas` library, requires additional libraries:
 - `pandas`
 - `openpyxl`
```python
from tksheet import Sheet
import tkinter as tk
import pandas as pd


class demo(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)
        self.grid_columnconfigure(0, weight = 1)
        self.grid_rowconfigure(0, weight = 1)
        self.frame = tk.Frame(self)
        self.frame.grid_columnconfigure(0, weight = 1)
        self.frame.grid_rowconfigure(0, weight = 1)
        self.sheet = Sheet(self.frame,
                           data = pd.read_excel("excel_file.xlsx",      # filepath here
                                                #sheet_name = "sheet1", # optional sheet name here
                                                engine = "openpyxl",
                                                header = None).values.tolist())
        self.sheet.enable_bindings()
        self.frame.grid(row = 0, column = 0, sticky = "nswe")
        self.sheet.grid(row = 0, column = 0, sticky = "nswe")


app = demo()
app.mainloop()
```

## 24 Example Custom Right Click and Text Editor Functionality

This is to demonstrate adding your own commands to the in-built right click popup menu (or how you might start making your own right click menu functionality) and also creating a cell text editor the manual way.
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
                           data = [[f"Row {r}, Column {c}\nnewline1\nnewline2" for c in range(50)] for r in range(500)])
        self.sheet.enable_bindings(("single_select",
                                    "drag_select",
                                    "select_all",
                                    "column_select",
                                    "row_select",
                                    "column_width_resize",
                                    "double_click_column_resize",
                                    "arrowkeys",
                                    "row_height_resize",
                                    "double_click_row_resize",
                                    "right_click_popup_menu",
                                    "rc_select"
                                    ))
        self.sheet.popup_menu_add_command("Say Hello", self.new_right_click_button)
        self.sheet.popup_menu_add_command("Edit Cell", self.edit_cell, index_menu = False, header_menu = False)
        self.frame.grid(row = 0, column = 0, sticky = "nswe")
        self.sheet.grid(row = 0, column = 0, sticky = "nswe")

    def new_right_click_button(self, event = None):
        print ("Hello World!")

    def edit_cell(self, event = None):
        r, c = self.sheet.get_currently_selected()
        self.sheet.row_height(row = r, height = "text", only_set_if_too_small = True, redraw = False)
        self.sheet.column_width(column = c, width = "text", only_set_if_too_small = True, redraw = True)
        self.sheet.create_text_editor(row = r,
                                      column = c,
                                      text = self.sheet.get_cell_data(r, c),
                                      set_data_ref_on_destroy = False,
                                      binding = self.end_edit_cell)

    def end_edit_cell(self, event = None):
        newtext = self.sheet.get_text_editor_value(event,
                                                   r = event[0],
                                                   c = event[1],
                                                   set_data_ref_on_destroy = True,
                                                   move_down = True,
                                                   redraw = True,
                                                   recreate = True)
        print (newtext)


app = demo()
app.mainloop()
```
 - If you want to evaluate the value from the text editor you can set `set_data_ref_on_destroy` to `False` and do the evaluation to decide whether or not to use `set_cell_data()`.
 - If you want a totally new right click menu you can use `self.sheet.bind("<3>", <function>)` with a `tk.Menu` of your own design (right click is `<2>` on MacOS) and don't use `"right_click_popup_menu"` with `enable_bindings()`.

## 25 Example Displaying Selections

This is to demonstrate displaying what the user has selected in the sheet.
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
                           data = [[f"Row {r}, Column {c}\nnewline1\nnewline2" for c in range(50)] for r in range(500)])
        self.sheet.enable_bindings()
        self.sheet.extra_bindings([("all_select_events", self.sheet_select_event)])
        self.show_selections = tk.Label(self)
        self.frame.grid(row = 0, column = 0, sticky = "nswe")
        self.sheet.grid(row = 0, column = 0, sticky = "nswe")
        self.show_selections.grid(row = 1, column = 0, sticky = "nswe")

    def sheet_select_event(self, event = None):
        try:
            len(event)
        except:
            return
        try:
            if event[0] == "select_cell":
                self.show_selections.config(text = f"Cells: ({event[1] + 1},{event[2] + 1}) : ({event[1] + 1},{event[2] + 1})")
            elif "cells" in event[0]:
                self.show_selections.config(text = f"Cells: ({event[1] + 1},{event[2] + 1}) : ({event[3] + 1},{event[4]})")
            elif event[0] == "select_column":
                self.show_selections.config(text = f"Columns: {event[1] + 1} : {event[1] + 1}")
            elif "columns" in event[0]:
                self.show_selections.config(text = f"Columns: {event[1][0] + 1} : {event[1][-1] + 1}")
            elif event[0] == "select_row":
                self.show_selections.config(text = f"Rows: {event[1] + 1} : {event[1] + 1}")
            elif "rows" in event[0]:
                self.show_selections.config(text = f"Rows: {event[1][0] + 1} : {event[1][-1] + 1}")
            else:
                self.show_selections.config(text = "")
        except:
            self.show_selections.config(text = "")


app = demo()
app.mainloop()
```

## 26 Example List Box

This is to demonstrate some simple customization to make a different sort of widget (a list box).

```python
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
```
