# Table of Contents

1. [About tksheet](https://github.com/ragardner/tksheet/wiki#1-about-tksheet)
2. [Installation and Requirements](https://github.com/ragardner/tksheet/wiki#2-installation-and-requirements)
3. [Basic Initialization](https://github.com/ragardner/tksheet/wiki#3-basic-initialization)
4. [Initialization Options](https://github.com/ragardner/tksheet/wiki#4-initialization-options)
5. [Header and Index](https://github.com/ragardner/tksheet/wiki#5-header-and-index)
6. [Setting Table Data](https://github.com/ragardner/tksheet/wiki#6-setting-table-data)
7. [Getting Table Data](https://github.com/ragardner/tksheet/wiki#7-getting-table-data)
8. [Bindings and Functionality](https://github.com/ragardner/tksheet/wiki#8-bindings-and-functionality)
9. [Identifying Bound Event Mouse Position](https://github.com/ragardner/tksheet/wiki#9-identifying-bound-event-mouse-position)
10. [Table Colors](https://github.com/ragardner/tksheet/wiki#10-table-colors)
11. [Highlighting Cells](https://github.com/ragardner/tksheet/wiki#11-highlighting-cells)
12. [Text Font and Alignment](https://github.com/ragardner/tksheet/wiki#12-text-font-and-alignment)
13. [Row Heights and Column Widths](https://github.com/ragardner/tksheet/wiki#13-row-heights-and-column-widths)
14. [Getting Selected Cells](https://github.com/ragardner/tksheet/wiki#14-getting-selected-cells)
15. [Modifying Selected Cells](https://github.com/ragardner/tksheet/wiki#15-modifying-selected-cells)
16. [Modifying and Getting Scroll Positions](https://github.com/ragardner/tksheet/wiki#16-modifying-and-getting-scroll-positions)
17. [Readonly Cells](https://github.com/ragardner/tksheet/wiki#17-readonly-cells)
18. [Hiding Columns](https://github.com/ragardner/tksheet/wiki#18-hiding-columns)
19. [Hiding Table Elements](https://github.com/ragardner/tksheet/wiki#19-hiding-table-elements)
20. [Cell Text Editor](https://github.com/ragardner/tksheet/wiki#20-cell-text-editor)
21. [Dropdown Boxes](https://github.com/ragardner/tksheet/wiki#21-dropdown-boxes)
22. [Check Boxes](https://github.com/ragardner/tksheet/wiki#22-check-boxes)
23. [Table Options and Other Functions](https://github.com/ragardner/tksheet/wiki#23-table-options-and-other-functions)
24. [Example Loading Data from Excel](https://github.com/ragardner/tksheet/wiki#24-example-loading-data-from-excel)
25. [Example Custom Right Click and Text Editor Validation](https://github.com/ragardner/tksheet/wiki#25-example-custom-right-click-and-text-editor-validation)
26. [Example Displaying Selections](https://github.com/ragardner/tksheet/wiki#26-example-displaying-selections)
27. [Example List Box](https://github.com/ragardner/tksheet/wiki#27-example-list-box)
28. [Example Header Dropdown Boxes and Filtering](https://github.com/ragardner/tksheet/wiki#28-example-header-dropdown-boxes-and-filtering)
29. [Example ReadMe Screenshot Code](https://github.com/ragardner/tksheet/wiki#29-example-readme-screenshot-code)
30. [Example Saving tksheet as a csv File](https://github.com/ragardner/tksheet/wiki#30-example-saving-tksheet-as-a-csv-file)

## 1 About tksheet

`tksheet` is a Python tkinter table widget written in pure python. It is licensed under the [MIT license](https://github.com/ragardner/tksheet/blob/master/LICENSE.txt).

It works using tkinter canvases and moves lines, text and highlight rectangles around for only the visible portion of the table.

Cell values can be any class with a `str` method.

Some examples of things that are not possible with tksheet:
 - Cell merging
 - Cell text wrap
 - Changing font for individual cells
 - Different fonts for index and table
 - Mouse drag copy cells
 - Hide rows
 - Cell highlight borders
 - Highlighting continuous multiple cells with a single border

If you'd like to buy me a coffee for creating and supporting this library you can do so here: https://www.buymeacoffee.com/ragardner

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
show_table: bool = True,
show_top_left: bool = True,
show_row_index: bool = True,
show_header: bool = True,
show_x_scrollbar: bool = True,
show_y_scrollbar: bool = True,
width: int = None,
height: int = None,
headers: list = None,
header: list = None,
default_header: str = "letters", #letters, numbers or both
default_row_index: str = "numbers", #letters, numbers or both
show_default_header_for_empty: bool = True,
show_default_index_for_empty: bool = True,
page_up_down_select_row: bool = True,
expand_sheet_if_paste_too_big: bool = False,
paste_insert_column_limit: int = None,
paste_insert_row_limit: int = None,
show_dropdown_borders: bool = False,
ctrl_keys_over_dropdowns_enabled: bool = False,
arrow_key_down_right_scroll_page: bool = False,
enable_edit_cell_auto_resize: bool = True,
edit_cell_validation: bool = True,
data_reference: list = None,
data: list = None,
startup_select: tuple = None, #either (start row, end row, "rows"), (start column, end column, "rows") or (cells start row, cells start column, cells end row, cells end column, "cells")
startup_focus: bool = True,
total_columns: int = None,
total_rows: int = None,
column_width: int = 120,
header_height: str = "1", #str or int
max_colwidth: str = "inf", #str or int
max_rh: str = "inf", #str or int
max_header_height: str = "inf", #str or int
max_row_width: str = "inf", #str or int
row_index: list = None,
index: list = None,
after_redraw_time_ms: int = 100,
row_index_width: int = 100,
auto_resize_default_row_index: bool = True,
set_all_heights_and_widths: bool = False,
row_height: str = "1", #str or int
font: tuple = get_font(),
header_font: tuple = get_heading_font(),
index_font: tuple = get_index_font(), #currently has no effect
popup_menu_font: tuple = get_font(),
align: str = "w",
header_align: str = "center",
row_index_align: str = "center",
displayed_columns: list = [],
all_columns_displayed: bool = True,
max_undos: int = 20,
outline_thickness: int = 0,
outline_color: str = theme_light_blue['outline_color'],
column_drag_and_drop_perform: bool = True,
row_drag_and_drop_perform: bool = True,
empty_horizontal: int = 150,
empty_vertical: int = 100,
selected_rows_to_end_of_window: bool = False,
horizontal_grid_to_end_of_window: bool = False,
vertical_grid_to_end_of_window: bool = False,
show_vertical_grid: bool = True,
show_horizontal_grid: bool = True,
display_selected_fg_over_highlights: bool = False,
show_selected_cells_border: bool = True,
theme                              = "light blue",
popup_menu_fg                      = theme_light_blue['popup_menu_fg'],
popup_menu_bg                      = theme_light_blue['popup_menu_bg'],
popup_menu_highlight_bg            = theme_light_blue['popup_menu_highlight_bg'],
popup_menu_highlight_fg            = theme_light_blue['popup_menu_highlight_fg'],
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
 - `data_reference` and `data` are essentially the same.
 - `row_index` and `index` are the same, `index` takes priority, same as with `headers` and `header`.
 - `edit_cell_validation` (`bool`) is used when `extra_bindings()` have been set for cell edits. If a bound function returns something other than `None` it will be used as the cell value instead of the user input.
    - `True` makes data edits take place after the binding function is run.
    - `False` makes data edits take place before the binding function is run.

You can change these settings after initialization using the `set_options()` function.

## 5 Header and Index

Set the header to something non-default (if new header is shorter than total columns then default headers e.g. letters will be used on the end.
```python
headers(newheaders = None, index = None, reset_col_positions = False, show_headers_if_not_sheet = True, redraw = False)
```
 - Using an integer `int` for argument `newheaders` makes the sheet use that row as a header e.g. `headers(0)` means the first row will be used as a header (the first row will not be hidden in the sheet though), this is sort of equivalent to freezing the row.
 - Leaving `newheaders` as `None` and using the `index` argument returns the existing header value in that index.
 - Leaving all arguments as default e.g. `headers()` returns existing headers.

___

Set the index to something non-default (if new index is shorter than total rows then default index e.g. numbers will be used on the end.
```python
row_index(newindex = None, index = None, reset_row_positions = False, show_index_if_not_sheet = True, redraw = False)
```
 - Using an integer `int` for argument `newindex` makes the sheet use that column as an index e.g. `row_index(0)` means the first column will be used as an index (the first column will not be hidden in the sheet though), this is sort of equivalent to freezing the column.
 - Leaving `newindex` as `None` and using the `index` argument returns the existing row index value in that index.
 - Leaving all arguments as default e.g. `row_index()` returns the existing row index.

## 6 Setting Table Data

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
 - `redraw` (`bool`) refreshes the table after setting new data.
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
delete_rows(rows: set = set(), deselect_all = False, redraw = True)
```
 - Does not maintain selections.

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
move_rows(moveto: int, to_move_min: int, number_of_rows: int, move_data: bool = True, index_type = "displayed", create_selections: bool = True, redraw = False)
```
 - `to_move_min` is the first row in the series of rows.
 - `index_type` (`str`) either `"displayed"` or `"data"`

___

```python
delete_column(idx = 0, deselect_all = False, redraw = True)
```

___

```python
delete_columns(columns: set = set(), deselect_all = False, redraw = True)
```
 - Does not maintain selections.

___

```python
move_column(column, moveto)
```

___

```python
move_columns(moveto: int, to_move_min: int, number_of_columns: int, move_data: bool = True, index_type = "displayed", create_selections: bool = True, redraw = False)
```
 - `to_move_min` is the first column in the series of columns.
 - `index_type` (`str`) either `"displayed"` or `"data"`, e.g. if columns are hidden and you want to supply the function with data indexes not sheet displayed indexes.

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

## 7 Getting Table Data

Yield sheet rows one by one, includes default header and index if being used e.g. A, B, C, D, whereas `get_sheet_data()` does not.
```python
yield_sheet_rows(get_header = False, get_index = False)
```
 - `get_header` (`bool`) will put the header as the first row if `True`.
 - `get_index` (`bool`) will put index items as the first item in every row.

___

Get sheet data and, if required, header and index data.
```python
get_sheet_data(return_copy = False, get_header = False, get_index = False)
```
 - `return_copy` (`bool`) will copy all cells if `True`, also copies header and index if they are `True`.

___

Returns the main table data, readonly.
```python
@property
data()
```
 - e.g. `self.sheet.data`

___

The name of the actual internal sheet data list.
```python
.MT.data
```
 - You can use this to directly modify or retrieve the main table's data e.g. `cell_0_0 = my_sheet_name_here.MT.data[0][0]`

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

## 8 Bindings and Functionality

Enable table functionality and bindings.
```python
enable_bindings(*bindings)
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
    - "edit_header"
    - "edit_index"

Notes:
 - Dragging and dropping rows / columns is bound to shift - mouse left click and hold and drag.
 - `"edit_header"` and `"edit_index"` are not enabled by `bindings = "all"` and has to be enabled individually, double click or right click (if enabled) on header/index cells to edit.
 - To allow table expansion when pasting data which doesn't fit in the table use either:
    - `expand_sheet_if_paste_too_big = True` in sheet initialization arguments or
    - `sheet.set_options(expand_sheet_if_paste_too_big = True)`

___

Disable table functionality and bindings, uses the same arguments as `enable_bindings()`
```python
disable_bindings(*bindings)
```

___

Bind various table functionality to your own functions. To unbind a function either set `func` argument to `None` or leave it as default e.g. `extra_bindings("begin_copy")` to unbind `"begin_copy"`.
```python
extra_bindings(bindings, func = "None")
```

Notes:
 - Upon an event being triggered the bound function will be sent a [namedtuple](https://docs.python.org/3/library/collections.html#collections.namedtuple) containing variables relevant to that event, use `print()` or similar to see all the variable names in the event. Each event contains different variable names with the exception of `eventname` e.g. `event.eventname`
 - For most of the `"end_..."` events the bound function is run before the value is set.
 - The bound function for `"end_edit_cell"` is run before the cell data is set in order that a return value can set the cell instead of the user input. Using the event you can assess the user input and if needed override it with a return value which is not `None`. If `None` is the return value then the user input will NOT be overridden. The setting `edit_cell_validation` (see initialization or the function `set_options()`) can be used to turn off this return value checking. The `edit_cell` bindings also run if header/index editing is turned on.

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

## 9 Identifying Bound Event Mouse Position

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

## 10 Table Colors

To change the colors of individual cells, rows or columns use the functions listed under [highlighting cells](https://github.com/ragardner/tksheet/wiki#11-highlighting-cells).

For the colors of specific parts of the table such as gridlines and backgrounds use the function `set_options()`, arguments can be found [here](https://github.com/ragardner/tksheet/wiki#22-table-options-and-other-functions). Most of the `set_options()` arguments are the same as the sheet initialization arguments.

Otherwise you can change the theme using the below function.
```python
change_theme(theme = "light blue", redraw = True)
```
 - `theme` (`str`) options (themes) are `light blue`, `light green`, `dark`, `dark blue` and `dark green`.

## 11 Highlighting Cells

 - `bg` and `fg` arguments use either a tkinter color or a hex `str` color.
 - Highlighting cells, rows or columns will also change the colors of dropdown boxes and check boxes.

___

```python
highlight_cells(row = 0, column = 0, cells = [], canvas = "table", bg = None, fg = None, redraw = False, overwrite = True)
```
 - Setting `overwrite` to `False` allows a previously set `fg` to be kept while setting a new `bg` or vice versa.

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
highlight_rows(rows = [], bg = None, fg = None, highlight_index = True, redraw = False, end_of_screen = False, overwrite = True)
```
 - `end_of_screen` when `True` makes the row highlight go past the last column line if there is any room there.
 - Setting `overwrite` to `False` allows a previously set `fg` to be kept while setting a new `bg` or vice versa.

___

```python
highlight_columns(columns = [], bg = None, fg = None, highlight_header = True, redraw = False, overwrite = True)
```
 - Setting `overwrite` to `False` allows a previously set `fg` to be kept while setting a new `bg` or vice versa.

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

## 12 Text Font and Alignment

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
 - Use argument `"all"` for `rows` e.g. `align_rows("all")` to clear all specific row alignments.

___

Change the text alignment for **specific** columns, `"global"` resets to table setting.
```python
align_columns(columns = [], align = "global", align_header = False, redraw = True)
```
 - Use argument `"all"` for `columns` e.g. `align_columns("all")` to clear all specific column alignments.

___

Change the text alignment for **specific** cells inside the table, `"global"` resets to table setting.
```python
align_cells(row = 0, column = 0, cells = [], align = "global", redraw = True)
```
 - Use argument `"all"` for `row` e.g. `align_cells("all")` to clear all specific cell alignments.

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

___

```python
get_cell_alignments()
```

___

```python
get_row_alignments()
```

___

```python
get_column_alignments()
```

## 13 Row Heights and Column Widths

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

## 14 Getting Selected Cells

```python
get_currently_selected()
```
 - Returns `namedtuple` of `(row, column, type_)` e.g. `(0, 0, "column")`
    - `type_` can be `"row"`, `"column"` or `"cell"`

Usage example below:
```python
currently_selected = self.sheet.get_currently_selected()
if currently_selected:
    row = currently_selected.row
    column = currently_selected.column
    type_ = currently_selected.type_
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

## 15 Modifying Selected Cells

```python
set_currently_selected(row, column, type_ = "cell", selection_binding = True)
```
 - `type_` (`str`) either `"cell"`, `"row"` or `"column"`.
 - `selection_binding` if `True` runs extra bindings selection function if one has been specified using `extra_bindings()`.

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

## 16 Modifying and Getting Scroll Positions

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

## 17 Readonly Cells

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

___

```python
readonly_header(columns = [], readonly = True, redraw = True)
```

___

```python
readonly_index(rows = [], readonly = True, redraw = True)
```

## 18 Hiding Columns

Display only certain columns.
```python
display_columns(columns = None,
                all_columns_displayed = None,
                reset_col_positions = True,
                refresh = False,
                redraw = False,
                deselect_all = True)
```
 - `columns` (`int`, any iterable, `"all"`) are the columns to be displayed, omit the columns to be hidden.
 - Use argument `True` with `all_columns_displayed` to display all columns, however, there's no need to use `False` when `columns` is not `None`.

___

Hide specific columns.
```python
hide_columns(columns = set(),
             refresh = True, 
             deselect_all = True)
```
 - `columns` (`int`) uses data indexes not displayed, e.g. if you already have column 0 hidden and you want to hide the first column shown in the sheet you would use argument `columns = 1`.

## 19 Hiding Table Elements

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

## 20 Cell Text Editor

Open the currently selected cell in the main table.
```python
open_cell(ignore_existing_editor = True)
```
 - Function utilises the currently selected cell in the main table, even if a column/row is selected, to open a non selected cell first use `set_currently_selected()` to set the cell to open.

___

Open the currently selected cell but in the header.
```python
open_header_cell(ignore_existing_editor = True)
```
 - Also uses currently selected cell, which you can set with `set_currently_selected()`.

___

Open the currently selected cell but in the index.
```python
open_index_cell(ignore_existing_editor = True)
```
 - Also uses currently selected cell, which you can set with `set_currently_selected()`.

___

```python
set_text_editor_value(text = "", r = None, c = None)
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

## 21 Dropdown Boxes

Create a dropdown box (only creates the arrow and border and sets it up for usage, does not pop open the box).
```python
create_dropdown(r = 0,
                c = 0,
                values = [],
                set_value = None,
                state = "readonly",
                redraw = False,
                selection_function = None,
                modified_function = None)
```

```python
create_header_dropdown(c = 0,
                       values = [],
                       set_value = None,
                       state = "readonly",
                       redraw = False,
                       selection_function = None,
                       modified_function = None)
```

```python
create_index_dropdown(r = 0,
                      values = [],
                      set_value = None,
                      state = "readonly",
                      redraw = False,
                      selection_function = None,
                      modified_function = None)
```

Notes:
 - Use `selection_function`/`modified_function` like so `selection_function = my_function_name`. The function you use needs at least one argument because tksheet will send information to your function about the triggered dropdown.
 - When a user selects an item from the dropdown box the sheet will set the underlying cells data to the selected item, to bind this event use either the `selection_function` argument or see the function `extra_bindings()` with binding `"end_edit_cell"` [here](https://github.com/ragardner/tksheet/wiki#7-bindings-and-functionality).

 Arguments:
 - `r` and `c` (`int`, `str`) can be set to `"all"`
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

```python
get_header_dropdown_values(c = 0)
```

```python
get_index_dropdown_values(r = 0)
```

___

Set the values and displayed value of a chosen dropdown box.
```python
set_dropdown_values(r = 0, c = 0, set_existing_dropdown = False, values = [], displayed = None)
```

```python
set_header_dropdown_values(c = 0, set_existing_dropdown = False, values = [], displayed = None)
```

```python
set_index_dropdown_values(r = 0, set_existing_dropdown = False, values = [], displayed = None)
```

 - `set_existing_dropdown` if `True` takes priority over `r` and `c` and sets the values of the last popped open dropdown box (if one one is popped open, if not then an `Exception` is raised).
 - `values` (`list`, `tuple`)
 - `displayed` (`str`, `None`) if not `None` will try to set the displayed value of the chosen dropdown box to given argument.

___

Set and get bound dropdown functions.
```python
dropdown_functions(r, c, selection_function = "", modified_function = "")
```

```python
header_dropdown_functions(c, selection_function = "", modified_function = "")
```

```python
index_dropdown_functions(r, selection_function = "", modified_function = "")
```

___

Delete dropdown boxes.
```python
delete_dropdown(r = 0, c = 0)
```

```python
delete_header_dropdown(c = 0)
```

```python
delete_index_dropdown(r = 0)
```

 - Set first argument to `"all"` to delete all dropdown boxes on the sheet.

___

Get a dictionary of all dropdown boxes; keys: `(row int, column int)` and values: `(ttk combobox widget, tk canvas window object)`
```python
get_dropdowns()
```

```python
get_header_dropdowns()
```

```python
get_index_dropdowns()
```

___

Pop open a dropdown box.
```python
open_dropdown(r, c)
```

```python
open_header_dropdown(c)
```

```python
open_index_dropdown(r)
```

___

Close an already open dropdown box.
```python
close_dropdown(r, c)
```

```python
close_header_dropdown(c)
```

```python
close_index_dropdown(r)
```

 - Also destroys any opened text editor windows.

## 22 Check Boxes

Create a check box.
```python
create_checkbox(r,
                c,
                checked = False,
                state = "normal",
                redraw = False,
                check_function = None,
                text = "")
```

```python
create_header_checkbox(c,
                       checked = False,
                       state = "normal",
                       redraw = False,
                       check_function = None,
                       text = "")
```

```python
create_index_checkbox(r,
                      checked = False,
                      state = "normal",
                      redraw = False,
                      check_function = None,
                      text = "")
```

Notes:
 - Use `check_function` like so `check_function = my_function_name`. The function you use needs at least one argument because when the checkbox is clicked it will send information to your function about the clicked checkbox.
 - Use `highlight_cells()` or rows or columns to change the color of the checkbox.
 - Check boxes are always left aligned despite any align settings.

 Arguments:
 - `r` and `c` (`int`, `str`) can be set to `"all"`
 - `text` displays text next to the checkbox in the cell, but will not be used as data, data will either be `True` or `False`
 - `check_function` can be used to trigger a function when the user clicks a checkbox.
 - `state` can be `"normal"` or `"disabled"`. If `"disabled"` then color will be same as table grid lines, else it will be the cells text color.

___

Set or toggle a checkbox.
```python
click_checkbox(r, c, checked = None)
```

```python
click_header_checkbox(c, checked = None)
```

```python
click_index_checkbox(r, checked = None)
```

___

Get a dictionary of all check box dictionaries.
```python
get_checkboxes()
```

```python
get_header_checkboxes()
```

```python
get_index_checkboxes()
```

___

Delete a checkbox.
```python
delete_checkbox(r = 0, c = 0)
```

```python
delete_header_checkbox(c = 0)
```

```python
delete_index_checkbox(r = 0)
```

 - Set first argument to `"all"` to delete all check boxes.

___

Set or get information about a particular checkbox.
```python
checkbox(r,
         c,
         checked = None,
         state = None,
         check_function = "",
         text = None)
```

```python
header_checkbox(c,
                checked = None,
                state = None,
                check_function = "",
                text = None)
```

```python
index_checkbox(r,
               checked = None,
               state = None,
               check_function = "",
               text = None)
```

 - If any arguments are not default they will be set for the chosen checkbox.
 - If all arguments are default a dictionary of all the checkboxes information will be returned.

## 23 Table Options and Other Functions

The list of key word arguments available for `set_options()` are as follows, [see here](https://github.com/ragardner/tksheet/wiki#4-initialization-options) as a guide for what arguments to use.
```python
show_dropdown_borders
edit_cell_validation
show_default_header_for_empty
show_default_index_for_empty
enable_edit_cell_auto_resize
selected_rows_to_end_of_window
horizontal_grid_to_end_of_window
vertical_grid_to_end_of_window
page_up_down_select_row
expand_sheet_if_paste_too_big
paste_insert_column_limit
paste_insert_row_limit
ctrl_keys_over_dropdowns_enabled
arrow_key_down_right_scroll_page
display_selected_fg_over_highlights
empty_horizontal
empty_vertical
show_horizontal_grid
show_vertical_grid
top_left_fg_highlight
auto_resize_default_row_index
font
default_header
default_row_index
header_font
show_selected_cells_border
theme
max_colwidth
max_row_height
max_header_height
max_row_width
header_height
row_height
column_width
header_bg
header_border_fg
header_grid_fg
header_fg
header_selected_cells_bg
header_selected_cells_fg
header_hidden_columns_expander_bg
index_bg
index_border_fg
index_grid_fg
index_fg
index_selected_cells_bg
index_selected_cells_fg
index_hidden_rows_expander_bg
top_left_bg
top_left_fg
frame_bg
table_bg
table_grid_fg
table_fg
table_selected_cells_border_fg
table_selected_cells_bg
table_selected_cells_fg
resizing_line_fg
drag_and_drop_bg
outline_thickness
outline_color
header_selected_columns_bg
header_selected_columns_fg
index_selected_rows_bg
index_selected_rows_fg
table_selected_rows_border_fg
table_selected_rows_bg
table_selected_rows_fg
table_selected_columns_border_fg
table_selected_columns_bg
table_selected_columns_fg
popup_menu_font
popup_menu_fg
popup_menu_bg
popup_menu_highlight_bg
popup_menu_highlight_fg
row_drag_and_drop_perform
column_drag_and_drop_perform
redraw
```

___

Get internal storage dictionary of highlights, readonly cells, dropdowns etc.
```python
get_cell_options(canvas = "table")
```

___

Delete any alignments, dropdown boxes, checkboxes, highlights etc. that are larger than the sheets currently held data, includes row index and header in measurement of dimensions.
```python
delete_out_of_bounds_options()
```

___

Delete all alignments, dropdown boxes, checkboxes, highlights etc.
```python
reset_all_options()
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

## 24 Example Loading Data from Excel

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

## 25 Example Custom Right Click and Text Editor Validation

This is to demonstrate adding your own commands to the in-built right click popup menu (or how you might start making your own right click menu functionality) and also validating text editor input. In this demonstration the validation removes spaces from user input.
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
        self.sheet.enable_bindings("single_select",
                                   "drag_select",
                                   "edit_cell",
                                   "select_all",
                                   "column_select",
                                   "row_select",
                                   "column_width_resize",
                                   "double_click_column_resize",
                                   "arrowkeys",
                                   "row_height_resize",
                                   "double_click_row_resize",
                                   "right_click_popup_menu",
                                   "rc_select")
        self.sheet.extra_bindings([("begin_edit_cell", self.begin_edit_cell),
                                   ("end_edit_cell", self.end_edit_cell)])
        self.sheet.popup_menu_add_command("Say Hello", self.new_right_click_button)
        self.frame.grid(row = 0, column = 0, sticky = "nswe")
        self.sheet.grid(row = 0, column = 0, sticky = "nswe")

    def new_right_click_button(self, event = None):
        print ("Hello World!")

    def begin_edit_cell(self, event = None):
        return event.text

    def end_edit_cell(self, event = None):
        # remove spaces from user input
        if event.text is not None:
            return event.text.replace(" ", "")


app = demo()
app.mainloop()
```
 - If you want a totally new right click menu you can use `self.sheet.bind("<3>", <function>)` with a `tk.Menu` of your own design (right click is `<2>` on MacOS) and don't use `"right_click_popup_menu"` with `enable_bindings()`.

## 26 Example Displaying Selections

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
                self.show_selections.config(text = f"Cells: ({event.selectionboxes[0] + 1},{event.selectionboxes[1] + 1}) : ({event.selectionboxes[2]},{event.selectionboxes[3]})")
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

## 27 Example List Box

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

## 28 Example Header Dropdown Boxes and Filtering

A very simple demonstration of row filtering using header dropdown boxes.

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
        self.data = ([["3", "c", "z"],
                      ["1", "a", "x"],
                      ["1", "b", "y"],
                      ["2", "b", "y"],
                      ["2", "c", "z"]])
        self.sheet = Sheet(self.frame,
                           data = self.data,
                           theme = "dark",
                           height = 700,
                           width = 1100)
        self.sheet.enable_bindings("copy",
                                   "rc_select",
                                   "arrowkeys",
                                   "double_click_column_resize",
                                   "column_width_resize",
                                   "column_select",
                                   "row_select",
                                   "drag_select",
                                   "single_select",
                                   "select_all")
        self.frame.grid(row = 0, column = 0, sticky = "nswe")
        self.sheet.grid(row = 0, column = 0, sticky = "nswe")
        
        self.sheet.create_header_dropdown(c = 0,
                                            values = ["all", "1", "2", "3"],
                                            set_value = "all",
                                            selection_function = self.header_dropdown_selected)
        self.sheet.create_header_dropdown(c = 1,
                                            values = ["all", "a", "b", "c"],
                                            set_value = "all",
                                            selection_function = self.header_dropdown_selected)
        self.sheet.create_header_dropdown(c = 2,
                                            values = ["all", "x", "y", "z"],
                                            set_value = "all",
                                            selection_function = self.header_dropdown_selected)

    def header_dropdown_selected(self, event = None):
        hdrs = self.sheet.headers()
        # this function is run before header cell data is set by dropdown selection
        # so we have to get the new value from the event
        hdrs[event.column] = event.text
        if all(dd == "all" for dd in hdrs):
            self.sheet.set_sheet_data(self.data,
                                      reset_col_positions = False,
                                      reset_row_positions = False)
        else:
            self.sheet.set_sheet_data([row for row in self.data if all(row[c] == e or e == "all" for c, e in enumerate(hdrs))],
                                      reset_col_positions = False,
                                      reset_row_positions = False)
    
app = demo()
app.mainloop()
```

## 29 Example Readme Screenshot Code

The code used to make a screenshot for the readme file.

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
                           expand_sheet_if_paste_too_big = True,
                           empty_horizontal = 0,
                           empty_vertical = 0,
                           align = "w",
                           header_align = "c",
                           data = [[f"Row {r}, Column {c}\nnewline 1\nnewline 2" for c in range(6)] for r in range(21)],
                           headers = ["Dropdown Column", "Checkbox Column", "Center Aligned Column", "East Aligned Column", "", ""],
                           theme = "black",
                           height = 520,
                           width = 930)
        self.sheet.enable_bindings()
        self.sheet.enable_bindings("edit_header")
        self.frame.grid(row = 0, column = 0, sticky = "nswe")
        self.sheet.grid(row = 0, column = 0, sticky = "nswe")
        colors = ("#509f56",
                  "#64a85b",
                  "#78b160",
                  "#8cba66",
                  "#a0c36c",
                  "#b4cc71",
                  "#c8d576",
                  "#dcde7c",
                  "#f0e782",
                  "#ffec87",
                  "#ffe182",
                  "#ffdc7d",
                  "#ffd77b",
                  "#ffc873",
                  "#ffb469",
                  "#fea05f",
                  "#fc8c55",
                  "#fb784b",
                  "#fa6441",
                  "#f85037")
        self.sheet.align_columns(columns = 2, align = "c")
        self.sheet.align_columns(columns = 3, align = "e")
        self.sheet.create_dropdown(r = "all",
                                   c = 0,
                                   values = ["Dropdown"] + [f"{i}" for i in range(15)])
        self.sheet.create_checkbox(r = "all", c = 1, checked = True, text = "Checkbox")
        self.sheet.create_header_dropdown(c = 0, values = ["Header Dropdown"] + [f"{i}" for i in range(15)])
        self.sheet.create_header_checkbox(c = 1, checked = True, text = "Header Checkbox")
        self.sheet.align_cells(5, 0, align = "c")
        self.sheet.highlight_cells(5, 0, bg = "gray50", fg = "blue")
        self.sheet.highlight_cells(17, canvas = "index", bg = "yellow", fg = "black")
        self.sheet.highlight_cells(12, 1, bg = "gray90", fg = "purple")
        for r in range(len(colors)):
            self.sheet.highlight_cells(row = r,
                                       column = 3,
                                       fg = colors[r])
            self.sheet.highlight_cells(row = r,
                                       column = 4,
                                       bg = colors[r],
                                       fg = "black")
            self.sheet.highlight_cells(row = r,
                                       column = 5,
                                       bg = colors[r],
                                       fg = "purple")
        self.sheet.highlight_cells(column = 5,
                                   canvas = "header",
                                   bg = "white",
                                   fg = "purple")
        self.sheet.set_all_column_widths()
        self.sheet.extra_bindings("all", self.all_extra_bindings)
        
    def all_extra_bindings(self, event = None):
        #print (event)
        pass

        
app = demo()
app.mainloop()
```

## 30 Example Saving tksheet as a csv File

To both load a csv file and save tksheet data as a csv file not including headers and index.

```python
from tksheet import Sheet
import tkinter as tk
from tkinter import filedialog
import csv
from os.path import normpath
import io


class demo(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)
        self.withdraw()
        self.title("tksheet")
        self.grid_columnconfigure(0, weight = 1)
        self.grid_rowconfigure(0, weight = 1)
        self.frame = tk.Frame(self)
        self.frame.grid_columnconfigure(0, weight = 1)
        self.frame.grid_rowconfigure(0, weight = 1)
        self.sheet = Sheet(self.frame,
                           data = [[f"Row {r}, Column {c}" for c in range(6)] for r in range(21)])
        self.sheet.enable_bindings("all", "edit_header", "edit_index")
        self.frame.grid(row = 0, column = 0, sticky = "nswe")
        self.sheet.grid(row = 0, column = 0, sticky = "nswe")
        self.sheet.popup_menu_add_command("Open csv", self.open_csv)
        self.sheet.popup_menu_add_command("Save sheet", self.save_sheet)
        self.sheet.set_all_cell_sizes_to_text()
        self.sheet.change_theme("light green")
        
        # center the window and unhide
        self.update_idletasks()
        w = self.winfo_screenwidth() - 20
        h = self.winfo_screenheight() - 70
        size = (900, 500)
        x = (w/2 - size[0]/2)
        y = h/2 - size[1]/2
        self.geometry("%dx%d+%d+%d" % (size + ((w/2 - size[0]/2), h/2 - size[1]/2)))
        self.deiconify()

    def save_sheet(self):
        filepath = filedialog.asksaveasfilename(parent = self,
                                                title = "Save sheet as",
                                                filetypes = [('CSV File','.csv'),
                                                             ('TSV File','.tsv')],
                                                defaultextension = ".csv",
                                                confirmoverwrite = True)
        if not filepath or not filepath.lower().endswith((".csv", ".tsv")):
            return
        try:
            with open(normpath(filepath), "w", newline = "", encoding = "utf-8") as fh:
                writer = csv.writer(fh,
                                    dialect = csv.excel if filepath.lower().endswith(".csv") else csv.excel_tab,
                                    lineterminator = "\n")
                writer.writerows(self.sheet.yield_sheet_rows(get_header = False, get_index = False))
        except:
            return
                
    def open_csv(self):
        filepath = filedialog.askopenfilename(parent = self, title = "Select a csv file")
        if not filepath or not filepath.lower().endswith((".csv", ".tsv")):
            return
        try:
            with open(normpath(filepath), "r") as filehandle:
                filedata = filehandle.read()
            self.sheet.set_sheet_data([r for r in csv.reader(io.StringIO(filedata),
                                                             dialect = csv.Sniffer().sniff(filedata),
                                                             skipinitialspace = False)])
        except:
            return


app = demo()
app.mainloop()
```