# Table of Contents

- [About tksheet](https://github.com/ragardner/tksheet/wiki#about-tksheet)
- [Installation and Requirements](https://github.com/ragardner/tksheet/wiki#installation-and-requirements)
- [Basic Initialization](https://github.com/ragardner/tksheet/wiki#basic-initialization)
- [Initialization Options](https://github.com/ragardner/tksheet/wiki#initialization-options)
- [Header and Index](https://github.com/ragardner/tksheet/wiki#header-and-index)
- [Setting Table Data](https://github.com/ragardner/tksheet/wiki#setting-table-data)
- [Getting Table Data](https://github.com/ragardner/tksheet/wiki#getting-table-data)
- [Bindings and Functionality](https://github.com/ragardner/tksheet/wiki#bindings-and-functionality)
- [Identifying Bound Event Mouse Position](https://github.com/ragardner/tksheet/wiki#identifying-bound-event-mouse-position)
- [Table Colors](https://github.com/ragardner/tksheet/wiki#table-colors)
- [Highlighting Cells](https://github.com/ragardner/tksheet/wiki#highlighting-cells)
- [Text Font and Alignment](https://github.com/ragardner/tksheet/wiki#text-font-and-alignment)
- [Row Heights and Column Widths](https://github.com/ragardner/tksheet/wiki#row-heights-and-column-widths)
- [Getting Selected Cells](https://github.com/ragardner/tksheet/wiki#getting-selected-cells)
- [Modifying Selected Cells](https://github.com/ragardner/tksheet/wiki#modifying-selected-cells)
- [Modifying and Getting Scroll Positions](https://github.com/ragardner/tksheet/wiki#modifying-and-getting-scroll-positions)
- [Readonly Cells](https://github.com/ragardner/tksheet/wiki#readonly-cells)
- [Hiding Columns](https://github.com/ragardner/tksheet/wiki#hiding-columns)
- [Hiding Rows](https://github.com/ragardner/tksheet/wiki#hiding-rows)
- [Hiding Table Elements](https://github.com/ragardner/tksheet/wiki#hiding-table-elements)
- [Cell Text Editor](https://github.com/ragardner/tksheet/wiki#cell-text-editor)
- [Dropdown Boxes](https://github.com/ragardner/tksheet/wiki#dropdown-boxes)
- [Check Boxes](https://github.com/ragardner/tksheet/wiki#check-boxes)
- [Cell Formatting](https://github.com/ragardner/tksheet/wiki#cell-formatting)
- [Table Options and Other Functions](https://github.com/ragardner/tksheet/wiki#table-options-and-other-functions)
- [Example Loading Data from Excel](https://github.com/ragardner/tksheet/wiki#example-loading-data-from-excel)
- [Example Custom Right Click and Text Editor Validation](https://github.com/ragardner/tksheet/wiki#example-custom-right-click-and-text-editor-validation)
- [Example Displaying Selections](https://github.com/ragardner/tksheet/wiki#example-displaying-selections)
- [Example List Box](https://github.com/ragardner/tksheet/wiki#example-list-box)
- [Example Header Dropdown Boxes and Row Filtering](https://github.com/ragardner/tksheet/wiki#example-header-dropdown-boxes-and-row-filtering)
- [Example ReadMe Screenshot Code](https://github.com/ragardner/tksheet/wiki#example-readme-screenshot-code)
- [Example Saving tksheet as a csv File](https://github.com/ragardner/tksheet/wiki#example-saving-tksheet-as-a-csv-file)
- [Example Using and Creating Formatters](https://github.com/ragardner/tksheet/wiki#example-using-and-creating-formatters)
- [Contributing](https://github.com/ragardner/tksheet/wiki#contributing)

## **About tksheet**
----

`tksheet` is a Python tkinter table widget written in pure python. It is licensed under the [MIT license](https://github.com/ragardner/tksheet/blob/master/LICENSE.txt).

It works using tkinter canvases and moves lines, text and rectangles around for only the visible portion of the table.

Cell values can be any class with a `str` method.

Some examples of things that are not possible with tksheet:
- Cell merging
- Cell text wrap
- Changing font for individual cells
- Different fonts for index and table
- Mouse drag copy cells
- Cell highlight borders
- Highlighting continuous multiple cells with a single border

## **Installation and Requirements**
----

`tksheet` is available through PyPi (Python package index) and can be installed by using Pip through the command line `pip install tksheet`

Alternatively you can download the source code and (inside the tksheet directory) use the command line `python setup.py develop`

`tksheet` requires a Python version of `3.6` or higher.

## **Basic Initialization**
----

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

## **Initialization Options**
----

**This is a full list of all the start up arguments, the only required argument is the sheets parent, everything else has default arguments.**

```python
(
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
to_clipboard_delimiter = "\t",
to_clipboard_quotechar = '"',
to_clipboard_lineterminator = "\n",
from_clipboard_delimiters = ["\t"],
show_default_header_for_empty: bool = True,
show_default_index_for_empty: bool = True,
page_up_down_select_row: bool = True,
expand_sheet_if_paste_too_big: bool = False,
paste_insert_column_limit: int = None,
paste_insert_row_limit: int = None,
show_dropdown_borders: bool = False,
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
max_undos: int = 30,
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

## **Header and Index**
----

#### **Set the header.**
```python
headers(newheaders = None, index = None, reset_col_positions = False, show_headers_if_not_sheet = True, redraw = False)
```
- Using an integer `int` for argument `newheaders` makes the sheet use that row as a header e.g. `headers(0)` means the first row will be used as a header (the first row will not be hidden in the sheet though), this is sort of equivalent to freezing the row.
- Leaving `newheaders` as `None` and using the `index` argument returns the existing header value in that index.
- Leaving all arguments as default e.g. `headers()` returns existing headers.

___

#### **Set the index.**
```python
row_index(newindex = None, index = None, reset_row_positions = False, show_index_if_not_sheet = True, redraw = False)
```
- Using an integer `int` for argument `newindex` makes the sheet use that column as an index e.g. `row_index(0)` means the first column will be used as an index (the first column will not be hidden in the sheet though), this is sort of equivalent to freezing the column.
- Leaving `newindex` as `None` and using the `index` argument returns the existing row index value in that index.
- Leaving all arguments as default e.g. `row_index()` returns the existing row index.

## **Setting Table Data**
----

#### **Set sheet data, overwrites any existing data.**
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

#### **Set cell data, overwrites any existing data.**
```python
set_cell_data(r, c, value = "", redraw = False)
```

___

#### **Insert a row into the sheet.**
```python
insert_row(values: Union[list, int, None] = None, 
           idx: Union[str, int] = "end", 
           height = None, 
           deselect_all = False, 
           add_columns = False,
           mod_row_positions = True,
           redraw = True)
```
- Leaving `values` as `None` inserts an empty row, e.g. `insert_row()` will append an empty row to the sheet.
- `height` is the new rows displayed height in pixels, leave as `None` for default.
- `add_columns` checks the rest of the sheets rows are at least the length as the new row, leave as `False` for better performance.

___

#### **Set column data, overwrites any existing data.**
```python
set_column_data(c, values = tuple(), add_rows = True, redraw = False)
```
- `add_rows` adds extra rows to the sheet if the column data doesn't fit within current sheet dimensions.

___

#### **Insert a column into the sheet.**
```python
insert_column(values = None, idx = "end", width = None, deselect_all = False, add_rows = True, equalize_data_row_lengths = True,
              mod_column_positions = True,
              redraw = False)
```

___

#### **Insert multiple columns into the sheet.**
```python
insert_columns(columns = 1, idx = "end", widths = None, deselect_all = False, add_rows = True, equalize_data_row_lengths = True,
               mod_column_positions = True,
               redraw = False)
```
- `columns` can be either `int` or iterable of iterables.

___

#### **Set row data, overwrites any existing data.**
```python
set_row_data(r, values = tuple(), add_columns = True, redraw = False)
```

___

#### **Insert multiple rows into the sheet.**
```python
insert_rows(rows: Union[list, int] = 1, 
            idx: Union[str, int] = "end", 
            heights = None, 
            deselect_all = False, 
            add_columns = True,
            mod_row_positions = True, 
            redraw = True)
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

#### **Make all data rows the same length.**
```python
equalize_data_row_lengths()
```
- Makes every list in the table have the same number of elements, goes by longest list. This will only affect the data variable, not visible columns.

___

Modify widget height and width in pixels
```python
height_and_width(height = None, width = None)
```
- `height` (`int`) set a height in pixels
- `width` (`int`) set a width in pixels
If both arguments are `None` then table will reset to default tkinter canvas dimensions.

## **Getting Table Data**
----

#### **Yield / generate sheet rows one at a time.**

This is useful if your sheet is very large and you don't want to create an extra list in memory (it does actually generate one row at a time and not from a pre-built list).
```python
yield_sheet_rows(get_displayed = False, 
                 get_header = False,
                 get_index = False,
                 get_index_displayed = True,
                 get_header_displayed = True,
                 only_rows = None,
                 only_columns = None)
```

#### **Get sheet data as list of lists.**
```python
get_sheet_data(get_displayed = False, 
               get_header = False,
               get_index = False,
               get_index_displayed = True,
               get_header_displayed = True,
               only_rows = None,
               only_columns = None)
```

Note:
- The following keyword arguments both behave the same way for `yield_sheet_rows()` and `get_sheet_data()`.

Arguments:
- `get_displayed` (`bool`) if `True` it will return cell values as they are displayed on the screen. If `False` it will return any underlying data, for example if the cell is formatted.
- `get_header` (`bool`) if `True` it will return the header of the sheet even if there is not one.
- `get_index` (`bool`) if `True` it will return the index of the sheet even if there is not one.
- `get_index_displayed` (`bool`) if `True` it will return whatever index values are displayed on the screen, for example if there is a dropdown box with `text` set.
- `get_header_displayed` (`bool`) if `True` it will return whatever header values are displayed on the screen, for example if there is a dropdown box with `text` set.
- `only_rows` (`None`, `iterable`) with this argument you can supply an iterable of row indexes in any order to be the only rows that are returned.
- `only_columns` (`None`, `iterable`) with this argument you can supply an iterable of column indexes in any order to be the only columns that are returned.

___

#### **Get the main table data, readonly.**
```python
@property
data()
```
- e.g. `self.sheet.data`

___

#### **The name of the actual internal sheet data variable.**
```python
.MT.data
```
- You can use this to directly modify or retrieve the main table's data e.g. `cell_0_0 = my_sheet_name_here.MT.data[0][0]`. Note that this is the raw data and if there are cell formatters it will include the cell's formatter class, as such it is not recommended to use this to retrieve or modify data unless you know what you are doing.

___

```python
get_cell_data(r, c, get_displayed = False)
```

___

```python
get_row_data(r, 
             get_displayed = False, 
             get_index = False, 
             get_index_displayed = True, 
             only_columns = None)
```
- The above arguments behave the same way as for `get_sheet_data()`.

___

```python
get_column_data(c,
                get_displayed = False, 
                get_header = False,
                get_header_displayed = True, 
                only_rows = None)
```
- The above arguments behave the same way as for `get_sheet_data()`.

___

#### **Get number of rows in table data.**
```python
get_total_rows(include_index = False)
```

___

#### **Get number of columns in table data.**
```python
get_total_columns(include_header = False)
```

## **Bindings and Functionality**
----

#### **Enable table functionality and bindings.**
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
	- "arrowkeys" # all arrowkeys including page up and down
    - "up"
    - "down"
    - "left"
    - "right"
    - "prior" # page up
    - "next" # page down
	- "row_height_resize"
	- "double_click_row_resize"
	- "right_click_popup_menu"
	- "rc_select"
	- "rc_insert_column"
	- "rc_delete_column"
	- "rc_insert_row"
	- "rc_delete_row"
    - "ctrl_click_select"
    - "ctrl_select"
	- "copy"
	- "cut"
	- "paste"
	- "delete"
	- "undo"
	- "edit_cell"
    - "edit_header"
    - "edit_index"

Notes:
- `"edit_header"`, `"edit_index"`, `"ctrl_select"` and `"ctrl_click_select"` are not enabled by `bindings = "all"` and have to be enabled individually, double click or right click (if enabled) on header/index cells to edit.
- `"ctrl_select"` and `"ctrl_click_select"` are the same and you can use either one.
- To allow table expansion when pasting data which doesn't fit in the table use either:
   - `expand_sheet_if_paste_too_big = True` in sheet initialization arguments or
   - `sheet.set_options(expand_sheet_if_paste_too_big = True)`

Example:
- `sheet.enable_bindings("all", "edit_header", "edit_index", "ctrl_select")` to enable absolutely everything.

___

#### **Disable table functionality and bindings, uses the same arguments as `enable_bindings()`.**
```python
disable_bindings(*bindings)
```

___

#### **Bind various table functionality to your own functions. To unbind a function either set `func` argument to `None` or leave it as default e.g. `extra_bindings("begin_copy")` to unbind** `"begin_copy"`.
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
	- "all_select_events", "select", "selectevents", "select_events"
    - "all_modified_events", "sheetmodified", "sheet_modified" "modified_events", "modified"
	- "bind_all"
	- "unbind_all"
- `func` argument is the function you want to send the binding event to.
- Using one of the following `"all_modified_events", "sheetmodified", "sheet_modified" "modified_events", "modified"` will make any insert, delete or cell edit including pastes and undos send an event to your function. Please **note** that this will mean your function will have to return a value to use for cell edits unless the setting `edit_cell_validation` is `False`.

___

#### **Add commands to the in-built right click popup menu.**
```python
popup_menu_add_command(label, func, table_menu = True, index_menu = True, header_menu = True)
```

___

#### **Remove commands added using `popup_menu_add_command()` from the in-built right click popup menu.**
```python
popup_menu_del_command(label = None)
```
- If `label` is `None` then it removes all.

___

#### **Disable table functionality and bindings (uses the same options as `enable_bindings()`).**
```python
disable_bindings(bindings = "all")
```

___

#### **Enable or disable mousewheel, left click etc.**
```python
basic_bindings(enable = False)
```

___

#### **Enable or disable cell edit functionality, including Undo.**
```python
edit_bindings(enable = False)
```

___

#### **Enable or disable the ability to edit a specific cell using the inbuilt text editor.**
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

## **Identifying Bound Event Mouse Position**
----

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

#### **Sheet control actions for binding your own keys to e.g. `sheet.bind("<Control-B>", sheet.paste)`**
```python
cut(self, event = None)
copy(self, event = None)
paste(self, event = None)
delete(self, event = None)
undo(self, event = None)
edit_cell(self, event = None, dropdown = False)
```

## **Table Colors**
----

To change the colors of individual cells, rows or columns use the functions listed under [highlighting cells](https://github.com/ragardner/tksheet/wiki#11-highlighting-cells).

For the colors of specific parts of the table such as gridlines and backgrounds use the function `set_options()`, arguments can be found [here](https://github.com/ragardner/tksheet/wiki#22-table-options-and-other-functions). Most of the `set_options()` arguments are the same as the sheet initialization arguments.

Otherwise you can change the theme using the below function.
```python
change_theme(theme = "light blue", redraw = True)
```
- `theme` (`str`) options (themes) are `light blue`, `light green`, `dark`, `dark blue` and `dark green`.

## **Highlighting Cells**
----

- `bg` and `fg` arguments use either a tkinter color or a hex `str` color.
- Highlighting cells, rows or columns will also change the colors of dropdown boxes and check boxes.

___

```python
highlight_cells(row = 0, column = 0, cells = [], canvas = "table", bg = None, fg = None, redraw = True, overwrite = True)
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
highlight_rows(rows = [], bg = None, fg = None, highlight_index = True, redraw = True, end_of_screen = False, overwrite = True)
```
- `end_of_screen` when `True` makes the row highlight go past the last column line if there is any room there.
- Setting `overwrite` to `False` allows a previously set `fg` to be kept while setting a new `bg` or vice versa.

___

```python
highlight_columns(columns = [], bg = None, fg = None, highlight_header = True, redraw = True, overwrite = True)
```
- Setting `overwrite` to `False` allows a previously set `fg` to be kept while setting a new `bg` or vice versa.

___

```python
dehighlight_all()
```

___

```python
dehighlight_rows(rows = [], redraw = True)
```

___

```python
dehighlight_columns(columns = [], redraw = True)
```

## **Text Font and Alignment**
----

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

#### **Change the text alignment for specific rows, `"global"` resets to table setting.**
```python
align_rows(rows = [], align = "global", align_index = False, redraw = True)
```
- Use argument `"all"` for `rows` e.g. `align_rows("all")` to clear all specific row alignments.

___

#### **Change the text alignment for specific columns, `"global"` resets to table setting.**
```python
align_columns(columns = [], align = "global", align_header = False, redraw = True)
```
- Use argument `"all"` for `columns` e.g. `align_columns("all")` to clear all specific column alignments.

___

#### **Change the text alignment for specific cells inside the table, `"global"` resets to table setting.**
```python
align_cells(row = 0, column = 0, cells = [], align = "global", redraw = True)
```
- Use argument `"all"` for `row` e.g. `align_cells("all")` to clear all specific cell alignments.

___

#### **Change the text alignment for specific cells inside the header, `"global"` resets to header setting.**
```python
align_header(columns = [], align = "global", redraw = True)
```

___

#### **Change the text alignment for specific cells inside the index, `"global"` resets to index setting.**
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

## **Row Heights and Column Widths**
----

#### **Set default column width in pixels.**
```python
default_column_width(width = None)
```
- `width` (`int`).

___

#### **Set default row height in pixels or lines.**
```python
default_row_height(height = None)
```
- `height` (`int`, `str`) use a numerical `str` for number of lines e.g. `"3"` for a height that fits 3 lines or `int` for pixels.

___

#### **Set default header bar height in pixels or lines.**
```python
default_header_height(height = None)
```
- `height` (`int`, `str`) use a numerical `str` for number of lines e.g. `"3"` for a height that fits 3 lines or `int` for pixels.

___

#### **Set a specific cell size to its text.**
```python
set_cell_size_to_text(row, column, only_set_if_too_small = False, redraw = True)
```

___

#### **Set all row heights and column widths to cell text sizes.**
```python
set_all_cell_sizes_to_text(redraw = True)
```

___

#### **Get the sheets column widths.**
```python
get_column_widths(canvas_positions = False)
```
- `canvas_positions` (`bool`) gets the actual canvas x coordinates of column lines.

___

#### **Get the sheets row heights.**
```python
get_row_heights(canvas_positions = False)
```
- `canvas_positions` (`bool`) gets the actual canvas y coordinates of row lines.

___

#### **Set all column widths to specific `width` in pixels (`int`) or leave `None` to set to cell text sizes for each column.**
```python
set_all_column_widths(width = None, only_set_if_too_small = False, redraw = True, recreate_selection_boxes = True)
```

___

#### **Set all row heights to specific `height` in pixels (`int`) or leave `None` to set to cell text sizes for each row.**
```python
set_all_row_heights(height = None, only_set_if_too_small = False, redraw = True, recreate_selection_boxes = True)
```

___

#### **Set/get a specific column width.**
```python
column_width(column = None, width = None, only_set_if_too_small = False, redraw = True)
```

___

```python
set_column_widths(column_widths = None, canvas_positions = False, reset = False, verify = False)
```

___

```python
set_width_of_index_to_text(text = None)
```
- `text` (`str`, `None`) provide a `str` to set the width to or use `None` to set it to existing values in the index.

___

```python
set_height_of_header_to_text(text = None)
```
- `text` (`str`, `None`) provide a `str` to set the height to or use `None` to set it to existing values in the header.

___

#### **Set/get a specific row height.**
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

## **Getting Selected Cells**
----

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

#### **Check if cell is selected, returns `bool`.**
```python
cell_selected(r, c)
```

___

#### **Check if row is selected, returns `bool`.**
```python
row_selected(r)
```

___

#### **Check if column is selected, returns `bool`.**
```python
column_selected(c)
```

___

#### **Check if any cells, rows or columns are selected, there are options for exclusions, returns `bool`.**
```python
anything_selected(exclude_columns = False, exclude_rows = False, exclude_cells = False)
```

___

#### **Check if user has the entire table selected, returns `bool`.**
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

## **Modifying Selected Cells**
----

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

## **Modifying and Getting Scroll Positions**
----

#### **See / scroll to a specific cell on the sheet.**
```python
see(row = 0, column = 0, keep_yscroll = False, keep_xscroll = False, bottom_right_corner = False, check_cell_visibility = True)
```

___

#### **Check if a cell has any part of it visible, returns `bool`.**
```python
cell_visible(r, c)
```

___

#### **Check if a cell is totally visible, returns `bool`.**
```python
cell_completely_visible(r, c, seperate_axes = False)
```
- `separate_axes` returns tuple of bools e.g. `(cell y axis is visible, cell x axis is visible)`

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

## **Readonly Cells**
----

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

## **Hiding Columns**
----

#### **Display only certain columns.**
```python
display_columns(columns = None,
                all_columns_displayed = None,
                reset_col_positions = True,
                refresh = False,
                redraw = False,
                deselect_all = True,
                **kwargs)
```
- `columns` (`int`, `iterable`, `"all"`) are the columns to be displayed, omit the columns to be hidden.
- Use argument `True` with `all_columns_displayed` to display all columns, however, there's no need to use `False` when `columns` is not `None`.
- You can also use the keyword argument `all_displayed` instead of `all_columns_displayed`.
- Example usage to display all columns: `sheet.display_columns("all")`.
- Example usage to display specific columns only: `sheet.display_columns([2, 4, 7], all_displayed = False)`.

___

#### **Get all columns displayed boolean.**
```python
all_columns_displayed(a = None)
```
- `a` (`bool`, `None`) Either set by using `bool` or get by leaving `None` e.g. `all_columns_displayed()`.

___

#### **Hide specific columns.**
```python
hide_columns(columns = set(),
             redraw = True, 
             deselect_all = True)
```
- **NOTE**: `columns` (`int`) uses displayed column indexes, not data indexes. In other words the indexes of the columns displayed on the screen are the ones that are hidden, this is useful when uses in conjunction with `get_selected_columns()`.

## **Hiding Rows**
----

#### **Display only certain rows.**
```python
display_rows(rows = None,
             all_rows_displayed = None,
             reset_col_positions = True,
             refresh = False,
             redraw = False,
             deselect_all = True,
             **kwargs)
```
- `rows` (`int`, `iterable`, `"all"`) are the rows to be displayed, omit the rows to be hidden.
- Use argument `True` with `all_rows_displayed` to display all rows, however, there's no need to use `False` when `rows` is not `None`.
- You can also use the keyword argument `all_displayed` instead of `all_rows_displayed`.
- Example usage to display all rows: `sheet.display_rows("all")`.
- Example usage to display specific rows only: `sheet.display_rows([2, 4, 7], all_displayed = False)`.
- [This is a very simple example of row filtering](https://github.com/ragardner/tksheet/wiki#example-header-dropdown-boxes-and-row-filtering) using this function.

___

#### **Get all rows displayed boolean.**
```python
all_rows_displayed(a = None)
```
- `a` (`bool`, `None`) Either set by using `bool` or get by leaving `None` e.g. `all_rows_displayed()`.

___

#### **Hide specific rows.**
```python
hide_rows(rows = set(),
          redraw = True, 
          deselect_all = True)
```
- **NOTE**: `rows` (`int`) uses displayed row indexes, not data indexes. In other words the indexes of the rows displayed on the screen are the ones that are hidden, this is useful when uses in conjunction with `get_selected_rows()`.

## **Hiding Table Elements**
----

#### **Hide parts of the table or all of it.**
```python
hide(canvas = "all")
```
- `canvas` (`str`) options are `all`, `row_index`, `header`, `top_left`, `x_scrollbar`, `y_scrollbar`
	- `all` hides the entire table and is the default.

___

#### **Show parts of the table or all of it.**
```python
show(canvas = "all")
```
- `canvas` (`str`) options are `all`, `row_index`, `header`, `top_left`, `x_scrollbar`, `y_scrollbar`
	- `all` shows the entire table and is the default.

## **Cell Text Editor**
----

#### **Open the currently selected cell in the main table.**
```python
open_cell(ignore_existing_editor = True)
```
- Function utilises the currently selected cell in the main table, even if a column/row is selected, to open a non selected cell first use `set_currently_selected()` to set the cell to open.

___

#### **Open the currently selected cell but in the header.**
```python
open_header_cell(ignore_existing_editor = True)
```
- Also uses currently selected cell, which you can set with `set_currently_selected()`.

___

#### **Open the currently selected cell but in the index.**
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

## **Dropdown Boxes**
----

#### **Dropdown box creation**

When using the functions to create dropdown boxes these are the default arguments:
```python
values = [],
set_value = None,
state = "normal",
redraw = True,
selection_function = None,
modified_function = None,
search_function = dropdown_search_function,
validate_input = True,
text = None
```

Notes:
- Use `selection_function`/`modified_function` like so `selection_function = my_function_name`. The function you use needs at least one argument because tksheet will send information to your function about the triggered dropdown.
- When a user selects an item from the dropdown box the sheet will set the underlying cells data to the selected item, to bind this event use either the `selection_function` argument or see the function `extra_bindings()` with binding `"end_edit_cell"` [here](https://github.com/ragardner/tksheet/wiki#7-bindings-and-functionality).

 Arguments:
- `values` are the values to appear when the dropdown box is popped open.
- `state` determines whether or not there is also an editable text window at the top of the dropdown box when it is open.
- `redraw` refreshes the sheet so the newly created box is visible.
- `selection_function` can be used to trigger a specific function when an item from the dropdown box is selected, if you are using the above `extra_bindings()` as well it will also be triggered but after this function. e.g. `selection_function = my_function_name`
- `modified_function` can be used to trigger a specific function when the `state` of the box is set to `"normal"` and there is an editable text window and a change of the text in that window has occurred. Note that this function occurs before the dropdown boxes search feature.
- `search_function` (`None`, `callable`) sets the function that will be used to search the dropdown boxes values upon a dropdown text editor modified event when the dropdowns state is `normal`. Set to `None` to disable the search feature or use your own function with the following keyword arguments: `(search_for, data):` and make it return an row number (e.g. select and see the first value would be `0`) if positive and `None` if negative.
- `validate_input` (`bool`) when `True` will not allow cut, paste, delete or cell editor to input values to cell which are not in the dropdown boxes values.
- `text` (`None`, `str`) can be set to something other than `None` to always display over whatever value is in the cell, this is useful when you want to display a Header name over a dropdown box selection.

Box creation functions:
```python
create_dropdown(r = 0, c = 0, *args, **kwargs)
dropdown_cell(r = 0, c = 0, *args, **kwargs)
```
- `r` and `c` (`int`, `"all"`) can be set to `"all"` or `int`.

```python
dropdown_row(r = 0, *args, **kwargs)
```
- `r` (`int`, `"all"`, `iterable`) can be set to `"all"`, `int` or any `iterable` of `int`s.

```python
dropdown_column(c = 0, *args, **kwargs)
```
- `c` (`int`, `"all"`, `iterable`) can be set to `"all"`, `int` or any `iterable` of `int`s.

```python
dropdown_sheet(*args, **kwargs)
```
- Sets the entire sheet to always have a dropdown box in every cell with the specified arguments.

```python
create_index_dropdown(r = 0, *args, **kwargs)
```
- `r` (`int`, `None`, `iterable`) use `None` to set the entire index to always having dropdown boxes with the chosen args/kwargs.

```python
create_header_dropdown(c = 0, *args, **kwargs)
```
- `c` (`int`, `None`, `iterable`) use `None` to set the entire header to always having dropdown boxes with the chosen args/kwargs.

Examples:
```python
# dropdown boxes all the way down column index 5
sheet.dropdown_column(5, values = ["ON", "OFF"])

# dropdown boxes all the way across rows 0 and 1
sheet.dropdown_row(range(2), values = ["1", "2", "3"])

# headers always have dropdown boxes with the same values
sheet.create_header_dropdown(c = None, values = ["val 1", "val 2"])
```

___

#### **Get chosen dropdown boxes values.**
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

#### **Set the values and cell value of a chosen dropdown box.**
```python
set_dropdown_values(r = 0, c = 0, set_existing_dropdown = False, values = [], set_value = None)
```

```python
set_header_dropdown_values(c = 0, set_existing_dropdown = False, values = [], set_value = None)
```

```python
set_index_dropdown_values(r = 0, set_existing_dropdown = False, values = [], set_value = None)
```

- `set_existing_dropdown` if `True` takes priority over `r` and `c` and sets the values of the last popped open dropdown box (if one one is popped open, if not then an `Exception` is raised).
- `values` (`list`, `tuple`)
- `set_value` (`str`, `None`) if not `None` will try to set the value of the chosen cell to given argument.

___

#### **Set and get bound dropdown functions.**
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

#### **Delete dropdown boxes.**

```python
delete_dropdown(r = 0, c = 0)
delete_cell_dropdown(r = 0, c = 0)
```
- Deletes dropdown boxes created by `create_dropdown()`/`dropdown_cell()`.
- `r` and `c` (`int`, `"all"`).

```python
delete_row_dropdown(r = 0)
```
- Deletes dropdown boxes created by `dropdown_row()`.
- `r` (`int`, `"all"`, `iterable`).

```python
delete_column_dropdown(c = 0)
```
- Deletes dropdown boxes created by `dropdown_column()`.
- `c` (`int`, `"all"`, `iterable`).

```python
delete_sheet_dropdown()
```
- Deletes dropdown boxes created by `dropdown_sheet()`.

```python
delete_index_dropdown(r = 0)
```
- Deletes dropdown boxes created by `create_header_dropdown()`.
- `r` (`int`, `"all"`, `None`). Use `None` to delete dropdowns created using `None`.

```python
delete_header_dropdown(c = 0)
```
- Deletes dropdown boxes created by `create_header_dropdown()`.
- `c` (`int`, `"all"`, `None`). Use `None` to delete dropdowns created using `None`.

___

#### **Get a dictionary of all dropdown boxes**
```python
get_dropdowns()
```

```python
get_header_dropdowns()
```

```python
get_index_dropdowns()
```
Returns:
```python
{(row int, column int): {'values': values,
                         'window': "no dropdown open", #the actual frame object if one exists
                         'canvas_id': "no dropdown open", #the canvas id of the frame object if one exists
                         'select_function': selection_function,
                         'modified_function': modified_function,
                         'state': state,
                         'text': text}}
```
- **Note:** This dictionary will also contain a key named `"dropdown"` if you have used `dropdown_sheet()`/`create_index_dropdown(None)`/`create_header_dropdown(None)`.

___

#### **Pop open a dropdown box.**
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

#### **Close an open dropdown box.**
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

## **Check Boxes**
----

#### **Checkbox creation.**

When using the functions to create checkboxes these are the default arguments:
```python
checked = False,
state = "normal",
redraw = False,
check_function = None,
text = ""
```

Notes:
- Use `check_function` like so `check_function = my_function_name`. The function you use needs at least one argument because when the checkbox is clicked it will send information to your function about the clicked checkbox.
- Use `highlight_cells()` or rows or columns to change the color of the checkbox.
- Check boxes are always left aligned despite any align settings.

 Arguments:
- `r` and `c` (`int`, `str`) can be set to `"all"`.
- `checked` is the initial creation value to set the box to.
- `text` displays text next to the checkbox in the cell, but will not be used as data, data will either be `True` or `False`.
- `check_function` can be used to trigger a function when the user clicks a checkbox.
- `state` can be `"normal"` or `"disabled"`. If `"disabled"` then color will be same as table grid lines, else it will be the cells text color.

Box creation functions:
```python
create_checkbox(r = 0, c = 0, *args, **kwargs)
checkbox_cell(r = 0, c = 0, *args, **kwargs)
```
- `r` and `c` (`int`, `"all"`) can be set to `"all"` or `int`.

```python
checkbox_row(r = 0, *args, **kwargs)
```
- `r` (`int`, `"all"`, `iterable`) can be set to `"all"`, `int` or any `iterable` of `int`s.

```python
checkbox_column(c = 0, *args, **kwargs)
```
- `c` (`int`, `"all"`, `iterable`) can be set to `"all"`, `int` or any `iterable` of `int`s.

```python
checkbox_sheet(*args, **kwargs)
```
- Sets the entire sheet to always have a checkboxbox in every cell with the specified arguments.

```python
create_index_checkbox(r = 0, *args, **kwargs)
```
- `r` (`int`, `None`, `iterable`) use `None` to set the entire index to always having checkboxes with the chosen args/kwargs.

```python
create_header_checkbox(c = 0, *args, **kwargs)
```
- `c` (`int`, `None`, `iterable`) use `None` to set the entire header to always having checkboxes boxes with the chosen args/kwargs.

Examples:
```python
# boxes all the way down column index 5
sheet.checkbox_column(5, text = "Enabled", check_function = my_function)

# boxes all the way across rows 0 and 1
sheet.checkbox_row(range(2), text = "Row Checkbox" checked = True)

# headers always have boxes with the same values
sheet.create_header_checkbox(c = None, text = "Header Checkbox")
```

___

#### **Set or toggle a checkbox.**
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

#### **Get a dictionary of all check box dictionaries.**
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

#### **Delete a checkbox.**
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

#### **Set or get information about a particular checkbox.**
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

## **Cell Formatting**
----

By default tksheet stores all user inputted data as strings and while tksheet can store and display any datatype with a `__str__()` method this has some obvious limitations. 

Cell formatting aims to provide greater functionality when working with different datatypes and provide strict typing for the sheet. With formatting you can convert sheet data and user input to a specific datatype. 

Additionally, formatting also provides a function for displaying data on the table GUI (as a rounded float for example) and logic for handling invalid and missing data. 

tksheet has several basic built-in formatters and provides functionality for creating your own custom formats as well.

A demonstration of all the built-in and custom formatters can be found [here](https://github.com/ragardner/tksheet/wiki#example-using-and-creating-formatters).

### **To Note:**
----

1. When applying multiple overlapping formats with e.g. a formatted cell which overlaps a formatted row, the priority is as follows:
- Cell formats first.
- Row formats second.
- Column formats third.
- Sheet format (using `format_sheet()`) last.

2. Cell formatting will effectively override `validate_input = True` on cells with dropdown boxes.

3. When getting data take careful note of the `get_displayed` options, as these are the difference between getting the actual formatted data and what is simply displayed on the table GUI.

### **Basic Intialisation**
----

#### **Applying a format to cell:**

```python
format_cell(r, c, formatter_options = {}, formatter_class = None, **kwargs)
```
Notes:
- `format_cell()` using `all` will not make added cells (such as with a sheet expanding paste or right click insert rows) formatted. For that you will have to use `format_row()`/`format_column()`/`format_sheet()`.

Arguments:
- `r` (`int` or `"all"`) the row index to apply the formatter to.
- `c` (`int` or `"all"`) the column index to apply the formatter to.
- `formatter_options` (`dict`) a dictionary of keyword options/arguements to pass to the formatter.
- `formatter_class` (`class`) in case you want to use a custom class to store functions and information as opposed to using the built-in methods.
- `**kwargs` any additional keyword options/arguements to pass to the formatter.

#### **Clearing a format that was set by `format_cell`:**

```python
delete_cell_format(r = "all", c = "all", clear_values = False)
```
- Using `"all"` will not clear formats set by `format_row()`/`format_column()`.
- `r` (`int` or `"all"`) the row index to remove the cell formats from.
- `c` (`int` or `"all"`) the column index to remove the cell formats from.
- `clear_values` (`bool`) if true, the cell values will be set to `""`.

#### **Applying a format to a row:**

```python
format_row(r, formatter_options = {}, formatter_class = None, **kwargs)
```
- `r` (`int`, `Iterable[int]` or `"all"`) the row index to apply the formatter to.
- `formatter_options` (`dict`) a dictionary of keyword options/arguements to pass to the formatter.
- `formatter_class` (`class`) in case you want to use a custom class to store functions and information as opposed to using the built-in methods.
- `**kwargs` any additional keyword options/arguements to pass to the formatter.

#### **Clearing a format that was set by `format_row`:**

```python
delete_row_format(r = "all", clear_values = False)
```
- Using `"all"` will not clear formats set by `format_sheet()`.
- `r` (`int`, `Iterable[int]` or `"all"`) the row index to remove the cell formats from.
- `clear_values` (`bool`) if true, the cell values will be set to `""`.

#### **Applying a format to a column:**

```python
format_column(c, formatter_options = {}, formatter_class = None, **kwargs)
```
- `c` (`int`, `Iterable[int]` or `"all"`) the column index to apply the formatter to.
- `formatter_options` (`dict`) a dictionary of keyword options/arguements to pass to the formatter.
- `formatter_class` (`class`) in case you want to use a custom class to store functions and information as opposed to using the built-in methods.
- `**kwargs` any additional keyword options/arguements to pass to the formatter.

#### **Clearing a format that was set by `format_column`:**

```python
delete_column_format(c = "all", clear_values = False)
```
- Using `"all"` will not clear formats set by `format_sheet()`.
- `c` (`int`, `Iterable[int]` or `"all"`) the column index to remove the cell formats from.
- `clear_values` (`bool`) if true, the cell values will be set to `""`.

#### **Applying a format to the whole sheet:**

```python
format_sheet(formatter_options = {}, formatter_class = None, **kwargs)
```
- `formatter_options` (`dict`) a dictionary of keyword options/arguements to pass to the formatter.
- `formatter_class` (`class`) in case you want to use a custom class to store functions and information as opposed to using the built-in methods.
- `**kwargs` any additional keyword options/arguements to pass to the formatter.

#### **Clearing a format that was set by `format_sheet`:**

```python
delete_sheet_format(clear_values = False)
```
- `clear_values` (`bool`) if true, all the sheets cell values will be set to `""`.

#### **Delete all formatting, including cell, row, column and sheet formats:**

```python
delete_all_formatting(clear_values = False)
```
- `clear_values` (`bool`) if true, all the sheets cell values will be set to `""`.

#### **Reapply formatting to entire sheet:**

```python
reapply_formatting()
```
- Useful if you have manually changed the entire sheets data using `sheet.MT.data = ` and want to reformat the sheet using any existing formatting you have set.

### **Formatter Options and In-Built Formatters**
----

tksheet provides a number of in-built formatters, in addition to the base `formatter` function. These formatters are designed to provide a range of functionality for different datatypes. The following table lists the available formatters and their options.

```python
formatter(datatypes,
          format_function,
          to_str_function = to_str,
          invalid_value = "NA",
          nullable = True,
          pre_format_function = None,
          post_format_function = None,
          clipboard_function = None,
          **kwargs)
```

This is the generic formatter options interface. You can use this to create your own custom formatters. The following options are available. Note that all these options can also be passed to the `format_cell()` function as keyword arguments and are available as attributes for all formatters. You can provide functions of your own creation for all the below arguments which take functions if you require.

- `datatypes` (`list`) a list of datatypes that the formatter will accept. For example, `datatypes = [int, float]` will accept integers and floats.
- `format_function` (`function`) a function that takes a string and returns a value of the desired datatype. For example, `format_function = int` will convert a string to an integer.
- `to_str_function` (`function`) a function that takes a value of the desired datatype and returns a string. This determines how the formatter displays its data on the table. For example, `to_str_function = str` will convert an integer to a string. Defaults to `tksheet.to_str`.
- `invalid_value` (`any`) the value to return if the input string is invalid. For example, `invalid_value = "NA"` will return "NA" if the input string is invalid.
- `nullable` (`bool`) if true, the formatter will accept `None` as a valid input.
- `pre_format_function` (`function`) a function that takes a input string and returns a string. This function is called before the `format_function` and can be used to modify the input string before it is converted to the desired datatype. This can be useful if you want to strip out unwanted characters or convert a string to a different format before converting it to the desired datatype.
- `post_format_function` (`function`) a function that takes a value of the desired datatype and returns a value of the desired datatype. This function is called after the `format_function` and can be used to modify the output value after it is converted to the desired datatype. This can be useful if you want to round a float for example.
- `clipboard_function` (`function`) a function that takes a value of the desired datatype and returns a string. This function is called when the cell value is copied to the clipboard. This can be useful if you want to convert a value to a different format before it is copied to the clipboard.
- `**kwargs` any additional keyword options/arguements to pass to the formatter. These keyword arguments will be passed to the `format_function`, `to_str_function`, and the `clipboard_function`. These can be useful if you want to specifiy any additional formatting options, such as the number of decimal places to round to.

#### **Int Formatter**

```python
int_formatter(datatypes = int,
              format_function = to_int,
              to_str_function = to_str,
              invalid_value = "NaN",
              **kwargs,
              ):
```

The `int_formatter` is the basic configuration for a simple interger formatter.

 - `format_function` (`function`) a function that takes a string and returns an `int`. By default, this is set to the in-built `tksheet.to_int`. This function will always convert float-likes to its floor, for example `"5.9"` will be converted to `5`.
 - `to_str_function` (`function`) By default, this is set to the in-built `tksheet.to_str`, which is a very basic function that will displace the default string representation of the value.

Usage:

```python
sheet.format_cell(0, 0, formatter_options = tksheet.int_formatter())
```

#### **Float Formatter**

```python
float_formatter(datatypes = float,
                format_function = to_float,
                to_str_function = float_to_str,
                invalid_value = "NaN",
                decimals = 2,
                **kwargs
                )
```

The `float_formatter` is the basic configuration for a simple float formatter. It will always round float-likes to the specified number of decimal places, for example `"5.999"` will be converted to `"6.0"` if `decimals = 1`.

 - `format_function` (`function`) a function that takes a string and returns a `float`. By default, this is set to the in-built `tksheet.to_float`. This function will always convert percentages to their decimal equivalent, for example `"5%"` will be converted to `0.05`.
 - `to_str_function` (`function`) By default, this is set to the in-built `tksheet.float_to_str`, which will display the float to the specified number of decimal places.
 - `decimals` (`int`, `None`) the number of decimal places to round to. Defaults to `2`.

Usage:

```python
sheet.format_cell(0, 0, formatter_options = tksheet.float_formatter(decimals = None)) # A float formatter with maximum float() decimal places
```

#### **Percentage Formatter**

```python
percentage_formatter(datatypes = float,
                     format_function = to_float,
                     to_str_function = percentage_to_str,
                     invalid_value = "NaN",
                     decimals = 0,
                     **kwargs,
                     )
```

The `percentage_formatter` is the basic configuration for a simple percentage formatter. It will always round float-likes as a percentage to the specified number of decimal places, for example `"5.999%"` will be converted to `"6.0%"` if `decimals = 1`.

 - `format_function` (`function`) a function that takes a string and returns a `float`. By default, this is set to the in-built `tksheet.to_float`. This function will always convert percentages to their decimal equivalent, for example `"5%"` will be converted to `0.05`.
 - `to_str_function` (`function`) By default, this is set to the in-built `tksheet.percentage_to_str`, which will display the float as a percentage to the specified number of decimal places. For example, `0.05` will be displayed as `"5.0%"`.
 - `decimals` (`int`) the number of decimal places to round to. Defaults to `0`. 

Usage:

```python
sheet.format_cell(0, 0, formatter_options = tksheet.percentage_formatter(decimals = 1)) # A percentage formatter with 1 decimal place
```

#### **Bool Formatter**

```python
bool_formatter(datatypes = bool,
               format_function = to_bool,
               to_str_function = bool_to_str,
               invalid_value = "NA",
               truthy = truthy,
               falsy = falsy,
               **kwargs,
               )
```

 - `format_function` (`function`) a function that takes a string and returns a `bool`. By default, this is set to the in-built `tksheet.to_bool`.
 - `to_str_function` (`function`) By default, this is set to the in-built `tksheet.bool_to_str`, which will display the boolean as `"True"` or `"False"`.
 - `truthy` (`set`) a set of values that will be converted to `True`. Defaults to the in-built `tksheet.truthy`.
 - `falsy` (`set`) a set of values that will be converted to `False`. Defaults to the in-built `tksheet.falsy`.

Usage:

```python
# A bool formatter with custom truthy and falsy values to account for aussie and kiwi slang
sheet.format_cell(0, 0, formatter_options = tksheet.bool_formatter(truthy = tksheet.truthy | {"nah yeah"}, falsy = tksheet.falsy | {"yeah nah"}))
```

### **Datetime Formatters and Designing Your Own Custom Formatters**
----

tksheet is at the moment a dependency free library and so doesn't include a datetime parser as is.

You can however very easily make a datetime parser if you are willing to install a third-party package. Recommended are:

- [dateparser](https://dateparser.readthedocs.io/en/latest/)
- [dateutil](https://dateutil.readthedocs.io/en/stable/)

Both of these packages have a very comprehensive datetime parser which can be used to create a custom datetime formatter for tksheet.

Below is a simple example of how you might create a custom datetime formatter using the `dateutil` package.

```python
from tksheet import *
from datetime import datetime, date
from dateutil.parser import parse

def to_local_datetime(dt, **kwargs):
    '''
    Our custom format_function, converts a string or a date to a datetime object in the local timezone.
    '''
    if isinstance(dt, datetime):
        pass # Do nothing
    elif isinstance(dt, date):
        dt = datetime(dt.year, dt.month, dt.day) # Always good to account for unexpected inputs 
    else:
        try:
            dt = parser.parse(dt)
        except:
            raise ValueError(f"Could not parse {dt} as a datetime")
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo = tzlocal()) # If no timezone is specified, assume local timezone
    dt = dt.astimezone(tzlocal()) # Convert to local timezone
    return dt

def datetime_to_str(dt, **kwargs):
    '''
    Our custom to_str_function, converts a datetime object to a string with a format that can be specfied in kwargs.
    '''
    return dt.strftime(kwargs['format'])
# Now we can create our custom formatter dictionary from the generic formatter interface in tksheet
datetime_formatter = formatter(datatypes = datetime,
                               format_function = to_local_datetime,
                               to_str_function = datetime_to_str,
                               invalid_value = "NaT",
                               format = "%d/%m/%Y %H:%M:%S",
                               )
# From here we can pass our datetime_formatter into format_cells() just like any other formatter
```

For those wanting even more customisation of their formatters you also have the option of creating a custom formatter class. This is a more advanced topic and is not covered here, but it's recommended to create a new class which is a subclass of the `tksheet.Formatter` class and overriding the methods you would like to customise. This custom class can then be passed into the `format_cells()` `formatter_class` argument.

## **Table Options and Other Functions**
----

The list of key word arguments available for `set_options()` are as follows, [see here](https://github.com/ragardner/tksheet/wiki#4-initialization-options) as a guide for what arguments to use.
```python
to_clipboard_delimiter
to_clipboard_quotechar
to_clipboard_lineterminator
from_clipboard_delimiters
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

Delete any formats, alignments, dropdown boxes, checkboxes, highlights etc. that are larger than the sheets currently held data, includes row index and header in measurement of dimensions.
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

## **Example Loading Data from Excel**
----

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

## **Example Custom Right Click and Text Editor Validation**
----

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

## **Example Displaying Selections**
----

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

## **Example List Box**
----

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

## **Example Header Dropdown Boxes and Row Filtering**
----

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
                                          selection_function = self.header_dropdown_selected,
                                          text = "Header A Name")
        self.sheet.create_header_dropdown(c = 1,
                                          values = ["all", "a", "b", "c"],
                                          set_value = "all",
                                          selection_function = self.header_dropdown_selected,
                                          text = "Header B Name")
        self.sheet.create_header_dropdown(c = 2,
                                          values = ["all", "x", "y", "z"],
                                          set_value = "all",
                                          selection_function = self.header_dropdown_selected,
                                          text = "Header C Name")

    def header_dropdown_selected(self, event = None):
        hdrs = self.sheet.headers()
        # this function is run before header cell data is set by dropdown selection
        # so we have to get the new value from the event
        hdrs[event.column] = event.text
        if all(dd == "all" for dd in hdrs):
            self.sheet.display_rows("all")
        else:
            rows = [rn for rn, row in enumerate(self.data) if all(row[c] == e or e == "all" for c, e in enumerate(hdrs))]
            self.sheet.display_rows(rows = rows,
                                    all_displayed = False)
        self.sheet.redraw()
    
app = demo()
app.mainloop()
```

## **Example Readme Screenshot Code**
----

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
        self.sheet.enable_bindings("all", "edit_index", "edit header")
        self.sheet.popup_menu_add_command("Hide Rows", self.hide_rows, table_menu = False, header_menu = False, empty_space_menu = False)
        self.sheet.popup_menu_add_command("Show All Rows", self.show_rows, table_menu = False, header_menu = False, empty_space_menu = False)
        self.sheet.popup_menu_add_command("Hide Columns", self.hide_columns, table_menu = False, index_menu = False, empty_space_menu = False)
        self.sheet.popup_menu_add_command("Show All Columns", self.show_columns, table_menu = False, index_menu = False, empty_space_menu = False)
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
        
    def hide_rows(self, event = None):
        rows = self.sheet.get_selected_rows()
        if rows:
            self.sheet.hide_rows(rows)
        
    def show_rows(self, event = None):
        self.sheet.display_rows("all", redraw = True)
        
    def hide_columns(self, event = None):
        columns = self.sheet.get_selected_columns()
        if columns:
            self.sheet.hide_columns(columns)
        
    def show_columns(self, event = None):
        self.sheet.display_columns("all", redraw = True)
        
    def all_extra_bindings(self, event = None):
        #print (event)
        try:
            if hasattr(event, 'text'):
                return event.text
        except:
            pass

        
app = demo()
app.mainloop()
```

## **Example Saving tksheet as a csv File**
----

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
                writer.writerows(self.sheet.get_sheet_data(get_header = False, get_index = False))
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

## **Example Using and Creating Formatters**
----

```python
from tksheet import *
import tkinter as tk
from datetime import datetime, date, timedelta, time
from dateutil import parser, tz
from math import ceil
import re

date_replace = re.compile('|'.join(['\(', '\)', '\[', '\]', '\<', '\>']))

# --------------------- Custom formatter methods ---------------------
def round_up(x):
    return float(ceil(x))

def only_numeric(s):
    return ''.join(n for n in f"{s}" if n.isnumeric() or n == '.')

def convert_to_local_datetime(dt: str, **kwargs):
    if isinstance(dt, datetime):
        pass
    elif isinstance(dt, date):
        dt = datetime(dt.year, dt.month, dt.day)
    else:
        if isinstance(dt, str):
            dt = date_replace.sub("", dt)
        try:
            dt = parser.parse(dt)
        except:
            raise ValueError(f"Could not parse {dt} as a datetime")
    if dt.tzinfo is None:
        dt.replace(tzinfo = tz.tzlocal())
    dt = dt.astimezone(tz.tzlocal())
    return dt.replace(tzinfo = None)

def datetime_to_string(dt: datetime, **kwargs):
    return dt.strftime('%d %b, %Y, %H:%M:%S')

# --------------------- Custom Formatter with additional kwargs ---------------------

def custom_datetime_to_str(dt: datetime, **kwargs):
    return dt.strftime(kwargs['format'])


class demo(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)
        self.grid_columnconfigure(0, weight = 1)
        self.grid_rowconfigure(0, weight = 1)
        self.frame = tk.Frame(self)
        self.frame.grid_columnconfigure(0, weight = 1)
        self.frame.grid_rowconfigure(0, weight = 1)
        self.sheet = Sheet(self.frame,
                           empty_vertical = 0,
                           empty_horizontal = 0,
                           data = [[f"{r}"]*11 for r in range(20)]
                           )
        self.sheet.enable_bindings()
        self.frame.grid(row = 0, column = 0, sticky = "nswe")
        self.sheet.grid(row = 0, column = 0, sticky = "nswe")
        self.sheet.headers(['Non-Nullable Float Cell\n1 decimals places', 
                            'Float Cell', 
                            'Int Cell', 
                            'Bool Cell', 
                            'Percentage Cell\n0 decimal places', 
                            'Custom Datetime Cell',
                            'Custom Datetime Cell\nCustom Format String',
                            'Float Cell that\nrounds up', 
                            'Float cell that\n strips non-numeric', 
                            'Dropdown Over Nullable\nPercentage Cell', 
                            'Percentage Cell\n2 decimal places'])
        
        # ---------- Some examples of cell formatting --------
        self.sheet.format_cell('all', 0, formatter_options = float_formatter(nullable = False))
        self.sheet.format_cell('all', 1, formatter_options = float_formatter())
        self.sheet.format_cell('all', 2, formatter_options = int_formatter())
        self.sheet.format_cell('all', 3, formatter_options = bool_formatter(truthy = truthy | {"nah yeah"}, falsy = falsy | {"yeah nah"}))
        self.sheet.format_cell('all', 4, formatter_options = percentage_formatter())


        # ---------------- Custom Formatters -----------------
        # Custom using generic formatter interface
        self.sheet.format_cell('all', 5, formatter_options = formatter(datatypes = datetime, 
                                                                       format_function = convert_to_local_datetime, 
                                                                       to_str_function = datetime_to_string, 
                                                                       nullable = False,
                                                                       invalid_value = 'NaT',
                                                                       ))
        # Custom format
        self.sheet.format_cell('all', 6, datatypes = datetime, 
                                         format_function = convert_to_local_datetime, 
                                         to_str_function = custom_datetime_to_str, 
                                         nullable = True,
                                         invalid_value = 'NaT',
                                         format = '(%Y-%m-%d) %H:%M %p'
                                         )
        
        # Unique cell behaviour using the post_conversion_function
        self.sheet.format_cell('all', 7, formatter_options = float_formatter(post_format_function = round_up))
        self.sheet.format_cell('all', 8, formatter_options = float_formatter(), pre_format_function = only_numeric)

        self.sheet.create_dropdown('all', 9, values = ['', '104%', .24, "300%", 'not a number'], set_value = 1)
        self.sheet.format_cell('all', 9, formatter_options = percentage_formatter(), decimals = 0)
        self.sheet.format_cell('all', 10, formatter_options = percentage_formatter(decimals = 5))


app = demo()
app.mainloop()
```

## **Contributing**
----

Welcome and thank you for your interest in `tksheet`!

### **tksheet Goals**

- **Adaptable rather than comprehensive**: Prioritizes adaptability over comprehensiveness, providing essential features that can be easily extended or customized based on specific needs. This approach allows for flexibility in integrating tksheet into different projects and workflows.

- **Lightweight and performant**: Aims to provide a lightweight solution for creating spreadsheet-like functionality in tkinter applications, without additional dependencies and with a focus on efficiency and performance.

### **Dependencies**

tksheet is designed to only use built-in Python libraries (without third-party dependencies). Please ensure that your contributions do not introduce any new dependencies outside of Python's built-in libraries.

### **License**

tksheet is released under the MIT License. You can find the full text of the license [here](https://github.com/ragardner/tksheet/blob/master/LICENSE.txt).

By contributing to the tksheet project, you agree to license your contributions under the same MIT License. Please make sure to read and understand the terms and conditions of the license before contributing.

### **Contributing Code**

To contribute, please follow these steps:

1. Fork the tksheet repository.
2. If you are working on a new feature, create a new branch for your contribution. Use a descriptive name for the branch that reflects the feature you're working on.
3. Make your changes in your local branch, following the code style and conventions established in the project.
4. Test your changes thoroughly to ensure they do not introduce any new bugs or issues.
5. Submit a pull request to the `main` branch of the tksheet repository, including a clear title and detailed description of your changes. Pull requests ideally should include a small but comprehensive demonstration of the feature you are adding.
6. Don't forget to update the documentation!

***Note:*** If you're submitting a bugfix, it's generally preferred to submit it directly to the relevant branch, rather than creating a separate branch.

### **Asking Questions**

Got a question that hasn't been answered in the closed issues or is missing from the documentation? please follow these guidelines:

- Submit your question as an issue in the [Issues tab](https://github.com/ragardner/tksheet/issues).
- Provide a clear and concise description of your question, including any relevant details or examples that can help us understand your query better.

### **Issues**

Please use the [Issues tab](https://github.com/ragardner/tksheet/issues) to report any issues or ask for assistance.

When submitting an issue, please follow these guidelines:

- Check the existing issues to see if a similar bug or question has already been reported or discussed.
- If reporting a bug, provide a minimal example that can reproduce the issue, including any relevant code, error messages, and steps to reproduce.
- If asking a question or seeking help, provide a clear and concise description of your question or issue, including any relevant details or examples that can help people understand your query better.
- Include any relevant screenshots or gifs that can visually illustrate the issue or your question.

### **Enhancements or Suggestions**

If you have an idea for a new feature, improvement or change, please follow these guidelines:

- Submit your suggestion as an issue in the [Issues tab](https://github.com/ragardner/tksheet/issues).
- Include a clear and concise description of your idea, including any relevant details, screenshots, or mock-ups that can help contributors understand your suggestion better.
- You're also welcomed to become a contributor yourself and help implement your idea!