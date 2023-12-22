# Table of Contents

- [About tksheet](https://github.com/ragardner/tksheet/wiki/Version-7#about-tksheet)
- [Installation and Requirements](https://github.com/ragardner/tksheet/wiki/Version-7#installation-and-requirements)
- [Basic Initialization](https://github.com/ragardner/tksheet/wiki/Version-7#basic-initialization)
- [Basic Use](https://github.com/ragardner/tksheet/wiki/Version-7#basic-use)
- [Initialization Options](https://github.com/ragardner/tksheet/wiki/Version-7#initialization-options)
---
- [Table Colors](https://github.com/ragardner/tksheet/wiki/Version-7#table-colors)
- [Header and Index](https://github.com/ragardner/tksheet/wiki/Version-7#header-and-index)
- [Bindings and Functionality](https://github.com/ragardner/tksheet/wiki/Version-7#bindings-and-functionality)
---
- [Span Objects](https://github.com/ragardner/tksheet/wiki/Version-7#span-objects)
- [Named Spans](https://github.com/ragardner/tksheet/wiki/Version-7#named-spans)
- [Getting Sheet Data](https://github.com/ragardner/tksheet/wiki/Version-7#getting-sheet-data)
- [Setting Sheet Data](https://github.com/ragardner/tksheet/wiki/Version-7#setting-sheet-data)
---
- [Highlighting Cells](https://github.com/ragardner/tksheet/wiki/Version-7#highlighting-cells)
- [Dropdown Boxes](https://github.com/ragardner/tksheet/wiki/Version-7#dropdown-boxes)
- [Check Boxes](https://github.com/ragardner/tksheet/wiki/Version-7#check-boxes)
- [Data Formatting](https://github.com/ragardner/tksheet/wiki/Version-7#data-formatting)
- [Readonly Cells](https://github.com/ragardner/tksheet/wiki/Version-7#readonly-cells)
- [Text Font and Alignment](https://github.com/ragardner/tksheet/wiki/Version-7#text-font-and-alignment)
---
- [Getting Selected Cells](https://github.com/ragardner/tksheet/wiki/Version-7#getting-selected-cells)
- [Modifying Selected Cells](https://github.com/ragardner/tksheet/wiki/Version-7#modifying-selected-cells)
- [Row Heights and Column Widths](https://github.com/ragardner/tksheet/wiki/Version-7#row-heights-and-column-widths)
- [Identifying Bound Event Mouse Position](https://github.com/ragardner/tksheet/wiki/Version-7#identifying-bound-event-mouse-position)
- [Modifying and Getting Scroll Positions](https://github.com/ragardner/tksheet/wiki/Version-7#modifying-and-getting-scroll-positions)
---
- [Hiding Columns](https://github.com/ragardner/tksheet/wiki/Version-7#hiding-columns)
- [Hiding Rows](https://github.com/ragardner/tksheet/wiki/Version-7#hiding-rows)
- [Hiding Table Elements](https://github.com/ragardner/tksheet/wiki/Version-7#hiding-table-elements)
---
- [Cell Text Editor](https://github.com/ragardner/tksheet/wiki/Version-7#cell-text-editor)
- [Table Options and Other Functions](https://github.com/ragardner/tksheet/wiki/Version-7#table-options-and-other-functions)
---
- [Example Loading Data from Excel](https://github.com/ragardner/tksheet/wiki/Version-7#example-loading-data-from-excel)
- [Example Custom Right Click and Text Editor Validation](https://github.com/ragardner/tksheet/wiki/Version-7#example-custom-right-click-and-text-editor-validation)
- [Example Displaying Selections](https://github.com/ragardner/tksheet/wiki/Version-7#example-displaying-selections)
- [Example List Box](https://github.com/ragardner/tksheet/wiki/Version-7#example-list-box)
- [Example Header Dropdown Boxes and Row Filtering](https://github.com/ragardner/tksheet/wiki/Version-7#example-header-dropdown-boxes-and-row-filtering)
- [Example ReadMe Screenshot Code](https://github.com/ragardner/tksheet/wiki/Version-7#example-readme-screenshot-code)
- [Example Saving tksheet as a csv File](https://github.com/ragardner/tksheet/wiki/Version-7#example-saving-tksheet-as-a-csv-file)
- [Example Using and Creating Formatters](https://github.com/ragardner/tksheet/wiki/Version-7#example-using-and-creating-formatters)
- [Contributing](https://github.com/ragardner/tksheet/wiki/Version-7#contributing)

---
# **About tksheet**

`tksheet` is a Python tkinter table widget written in pure python. It is licensed under the [MIT license](https://github.com/ragardner/tksheet/blob/master/LICENSE.txt).

It works by using tkinter canvases and moving lines, text and rectangles around for only the visible portion of the table.

### **Limitations**

Some examples of things that are not possible with tksheet:
- Cell merging
- Cell text wrap
- Changing font for individual cells
- Different fonts for index and table
- Mouse drag copy cells
- Cell highlight borders

---
# **Installation and Requirements**

`tksheet` is available through PyPi (Python package index) and can be installed by using Pip through the command line `pip install tksheet`

```python
#To install using pip
pip install tksheet

#To update using pip
pip install tksheet --upgrade
```

Alternatively you can download the source code and (inside the tksheet directory) use the command line `python setup.py develop`

`tksheet` versions < `7.0.0` require a Python version of `3.6` or higher.

`tksheet` versions >= `7.0.0` require a Python version of `3.8` or higher.

---
# **Basic Initialization**

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

---
# **Basic Use**

```python
# here
```

---
# **Initialization Options**

This is a full list of all the start up arguments, the only required argument is the sheets parent, everything else has default arguments.

```python
(
parent,
name: str = "!sheet",
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
default_header: str = "letters",  # letters, numbers or both
default_row_index: str = "numbers",  # letters, numbers or both
to_clipboard_delimiter="\t",
to_clipboard_quotechar='"',
to_clipboard_lineterminator="\n",
from_clipboard_delimiters=["\t"],
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
# either (start row, end row, "rows"), (start column, end column, "rows") or
# (cells start row, cells start column, cells end row, cells end column, "cells")  # noqa: E501
startup_select: tuple = None,
startup_focus: bool = True,
total_columns: int = None,
total_rows: int = None,
column_width: int = 120,
header_height: str = "1",  # str or int
max_column_width: str = "inf",  # str or int
max_row_height: str = "inf",  # str or int
max_header_height: str = "inf",  # str or int
max_index_width: str = "inf",  # str or int
row_index: list = None,
index: list = None,
after_redraw_time_ms: int = 20,
row_index_width: int = None,
auto_resize_default_row_index: bool = True,
auto_resize_columns = None,
auto_resize_rows = None,
set_all_heights_and_widths: bool = False,
row_height: str = "1",  # str or int
font: tuple = get_font(),
header_font: tuple = get_header_font(),
index_font: tuple = get_index_font(),  # currently has no effect
popup_menu_font: tuple = get_font(),
align: str = "w",
header_align: str = "center",
row_index_align = None,
index_align: str = "center",
displayed_columns: list = [],
all_columns_displayed: bool = True,
displayed_rows: list = [],
all_rows_displayed: bool = True,
max_undos: int = 30,
outline_thickness: int = 0,
outline_color: str = theme_light_blue["outline_color"],
column_drag_and_drop_perform: bool = True,
row_drag_and_drop_perform: bool = True,
empty_horizontal: int = 50,
empty_vertical: int = 50,
selected_rows_to_end_of_window: bool = False,
horizontal_grid_to_end_of_window: bool = False,
vertical_grid_to_end_of_window: bool = False,
show_vertical_grid: bool = True,
show_horizontal_grid: bool = True,
display_selected_fg_over_highlights: bool = False,
show_selected_cells_border: bool = True,
theme: str = "light blue",
popup_menu_fg: str = theme_light_blue["popup_menu_fg"],
popup_menu_bg: str = theme_light_blue["popup_menu_bg"],
popup_menu_highlight_bg: str = theme_light_blue["popup_menu_highlight_bg"],
popup_menu_highlight_fg: str = theme_light_blue["popup_menu_highlight_fg"],
frame_bg: str = theme_light_blue["table_bg"],
table_grid_fg: str = theme_light_blue["table_grid_fg"],
table_bg: str = theme_light_blue["table_bg"],
table_fg: str = theme_light_blue["table_fg"],
table_selected_box_cells_fg: str = theme_light_blue["table_selected_box_cells_fg"],
table_selected_box_rows_fg: str = theme_light_blue["table_selected_box_rows_fg"],
table_selected_box_columns_fg: str = theme_light_blue["table_selected_box_columns_fg"],
table_selected_cells_border_fg: str = theme_light_blue["table_selected_cells_border_fg"],
table_selected_cells_bg: str = theme_light_blue["table_selected_cells_bg"],
table_selected_cells_fg: str = theme_light_blue["table_selected_cells_fg"],
table_selected_rows_border_fg: str = theme_light_blue["table_selected_rows_border_fg"],
table_selected_rows_bg: str = theme_light_blue["table_selected_rows_bg"],
table_selected_rows_fg: str = theme_light_blue["table_selected_rows_fg"],
table_selected_columns_border_fg: str = theme_light_blue["table_selected_columns_border_fg"],
table_selected_columns_bg: str = theme_light_blue["table_selected_columns_bg"],
table_selected_columns_fg: str = theme_light_blue["table_selected_columns_fg"],
resizing_line_fg: str = theme_light_blue["resizing_line_fg"],
drag_and_drop_bg: str = theme_light_blue["drag_and_drop_bg"],
index_bg: str = theme_light_blue["index_bg"],
index_border_fg: str = theme_light_blue["index_border_fg"],
index_grid_fg: str = theme_light_blue["index_grid_fg"],
index_fg: str = theme_light_blue["index_fg"],
index_selected_cells_bg: str = theme_light_blue["index_selected_cells_bg"],
index_selected_cells_fg: str = theme_light_blue["index_selected_cells_fg"],
index_selected_rows_bg: str = theme_light_blue["index_selected_rows_bg"],
index_selected_rows_fg: str = theme_light_blue["index_selected_rows_fg"],
header_bg: str = theme_light_blue["header_bg"],
header_border_fg: str = theme_light_blue["header_border_fg"],
header_grid_fg: str = theme_light_blue["header_grid_fg"],
header_fg: str = theme_light_blue["header_fg"],
header_selected_cells_bg: str = theme_light_blue["header_selected_cells_bg"],
header_selected_cells_fg: str = theme_light_blue["header_selected_cells_fg"],
header_selected_columns_bg: str = theme_light_blue["header_selected_columns_bg"],
header_selected_columns_fg: str = theme_light_blue["header_selected_columns_fg"],
top_left_bg: str = theme_light_blue["top_left_bg"],
top_left_fg: str = theme_light_blue["top_left_fg"],
top_left_fg_highlight: str = theme_light_blue["top_left_fg_highlight"],
)
```
- `name` setting a name for the sheet is useful when you have multiple sheets and you need to determine which one an event came from.
- `auto_resize_columns` (`int`, `None`) if set as an `int` the columns will automatically resize to fit the width of the window, the `int` value being the minimum of each column in pixels.
- `auto_resize_rows` (`int`, `None`) if set as an `int` the rows will automatically resize to fit the height of the window, the `int` value being the minimum height of each row in pixels.
- `startup_select` selects cells, rows or columns at initialization by using a `tuple` e.g. `(0, 0, "cells")` for cell A0 or `(0, 5, "rows")` for rows 0 to 5.
- `data_reference` and `data` are essentially the same.
- `row_index` and `index` are the same, `index` takes priority, same as with `headers` and `header`.
- `edit_cell_validation` (`bool`) is used when `extra_bindings()` have been set for cell edits. If a bound function returns something other than `None` it will be used as the cell value instead of the user input.
   - `True` makes data edits take place after the binding function is run.
   - `False` makes data edits take place before the binding function is run.

You can change these settings after initialization using the `set_options()` function.

---
# **Table Colors**

To change the colors of individual cells, rows or columns use the functions listed under [highlighting cells](https://github.com/ragardner/tksheet/wiki/Version-7#highlighting-cells).

For the colors of specific parts of the table such as gridlines and backgrounds use the function `set_options()`, keyword arguments specific to table colors are listed below. All the other `set_options()` arguments can be found [here](https://github.com/ragardner/tksheet/wiki/Version-7#table-options-and-other-functions).

Use a tkinter color or a hex string e.g.

```python
my_sheet_widget.set_options(table_bg="black")
my_sheet_widget.set_options(table_bg="#000000")
```

```python
set_options(
top_left_bg
top_left_fg
top_left_fg_highlight

table_bg
table_grid_fg
table_fg
table_selected_box_cells_fg
table_selected_box_rows_fg
table_selected_box_columns_fg
table_selected_cells_border_fg
table_selected_cells_bg
table_selected_cells_fg
table_selected_rows_border_fg
table_selected_rows_bg
table_selected_rows_fg
table_selected_columns_border_fg
table_selected_columns_bg
table_selected_columns_fg

header_bg
header_border_fg
header_grid_fg
header_fg
header_selected_cells_bg
header_selected_cells_fg
header_selected_columns_bg
header_selected_columns_fg

index_bg
index_border_fg
index_grid_fg
index_fg
index_selected_cells_bg
index_selected_cells_fg
index_selected_rows_bg
index_selected_rows_fg

resizing_line_fg
drag_and_drop_bg
outline_thickness
outline_color
frame_bg
popup_menu_font
popup_menu_fg
popup_menu_bg
popup_menu_highlight_bg
popup_menu_highlight_fg
)
```

Otherwise you can change the theme using the below function.
```python
change_theme(theme = "light blue", redraw = True)
```
- `theme` (`str`) options (themes) are `light blue`, `light green`, `dark`, `dark blue` and `dark green`.

---
# **Header and Index**

#### **Set the header**
```python
set_header_data(value, c = None, redraw = True)
```
- `value` (`iterable`, `int`, `Any`) if `c` is left as `None` then it attempts to set the whole header as the `value` (converting a generator to a list). If `value` is `int` it sets the header to display the row with that position.
- `c` (`iterable`, `int`, `None`) if both `value` and `c` are iterables it assumes `c` is an iterable of positions and `value` is an iterable of values and attempts to set each value to each position. If `c` is `int` it attempts to set the value at that position.

```python
headers(newheaders = None, index = None, reset_col_positions = False, show_headers_if_not_sheet = True, redraw = False)
```
- Using an integer `int` for argument `newheaders` makes the sheet use that row as a header e.g. `headers(0)` means the first row will be used as a header (the first row will not be hidden in the sheet though), this is sort of equivalent to freezing the row.
- Leaving `newheaders` as `None` and using the `index` argument returns the existing header value in that index.
- Leaving all arguments as default e.g. `headers()` returns existing headers.

___

#### **Set the index**
```python
set_index_data(value, r = None, redraw = True)
```
- `value` (`iterable`, `int`, `Any`) if `r` is left as `None` then it attempts to set the whole index as the `value` (converting a generator to a list). If `value` is `int` it sets the index to display the row with that position.
- `r` (`iterable`, `int`, `None`) if both `value` and `r` are iterables it assumes `r` is an iterable of positions and `value` is an iterable of values and attempts to set each value to each position. If `r` is `int` it attempts to set the value at that position.

```python
row_index(newindex = None, index = None, reset_row_positions = False, show_index_if_not_sheet = True, redraw = False)
```
- Using an integer `int` for argument `newindex` makes the sheet use that column as an index e.g. `row_index(0)` means the first column will be used as an index (the first column will not be hidden in the sheet though), this is sort of equivalent to freezing the column.
- Leaving `newindex` as `None` and using the `index` argument returns the existing row index value in that index.
- Leaving all arguments as default e.g. `row_index()` returns the existing row index.

---
# **Bindings and Functionality**

#### **Enable table functionality and bindings**
```python
enable_bindings(*bindings)
```
- `bindings` (`str`) options are (rc stands for right click):
	- "all"
	- "single_select"
	- "toggle_select"
	- "drag_select"
       - "select_all"
	- "column_drag_and_drop" / "move_columns"
	- "row_drag_and_drop" / "move_rows"
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

#### **Disable table functionality and bindings, uses the same arguments as `enable_bindings()`**
```python
disable_bindings(*bindings)
```

___

#### **Bind specific table functionality**

This function allows you to bind very specific table functionality to your own functions. If you want less specificity in event names you can also bind all sheet modifying events to a single function, [see here](https://github.com/ragardner/tksheet/wiki/Version-7#bind-tkinter-events).

```python
extra_bindings(bindings, func=None)
```

Notes:
- There are several ways to use this function:
    - `bindings` as a `str` and `func` as either `None` or a function. Using `None` as an argument for `func` will effectively unbind the function.
        - `extra_bindings("edit_cell", func=my_function)`
    - `bindings` as an `iterable` of `str`s and `func` as either `None` or a function. Using `None` as an argument for `func` will effectively unbind the function.
        - `extra_bindings(["all_select_events", "copy", "cut"], func=my_function)`
    - `bindings` as an `iterable` of `list`s or `tuple`s with length of two, e.g.
        - `extra_bindings([(binding, function), (binding, function), ...])` In this example you could also use `None` in the place of `function` to unbind the binding.
        - In this case the arg `func` is totally ignored.
- For most of the `"end_..."` events the bound function is run before the value is set.
- **To unbind** a function either set `func` argument to `None` or leave it as default e.g. `extra_bindings("begin_copy")` to unbind `"begin_copy"`.
- The bound function for `"end_edit_cell"` is run before the cell data is set in order that a return value can set the cell instead of the user input. Using the event you can assess the user input and if needed override it with a return value which is not `None`. If `None` is the return value then the user input will NOT be overridden. The setting `edit_cell_validation` (see initialization or the function `set_options()`) can be used to turn off this return value checking. The `edit_cell` bindings also run if header/index editing is turned on.

Parameters:
- `bindings` (`str`) options are:
	- `"begin_copy", "begin_ctrl_c"`
	- `"ctrl_c", "end_copy", "end_ctrl_c", "copy"`
	- `"begin_cut", "begin_ctrl_x"`
	- `"ctrl_x", "end_cut", "end_ctrl_x", "cut"`
	- `"begin_paste", "begin_ctrl_v"`
	- `"ctrl_v", "end_paste", "end_ctrl_v", "paste"`
	- `"begin_undo", "begin_ctrl_z"`
	- `"ctrl_z", "end_undo", "end_ctrl_z", "undo"`
	- `"begin_delete_key", "begin_delete"`
	- `"delete_key", "end_delete", "end_delete_key", "delete"`
	- `"begin_edit_cell", "begin_edit_table"`
	- `"end_edit_cell", "edit_cell", "edit_table"`
	- `"begin_edit_header"`
    - `"end_edit_header", "edit_header"`
    - `"begin_edit_index"`
	- `"end_edit_index", "edit_index"`
    - `"begin_row_index_drag_drop", "begin_move_rows"`
	- `"row_index_drag_drop", "move_rows", "end_move_rows", "end_row_index_drag_drop"`
	- `"begin_column_header_drag_drop", "begin_move_columns"`
	- `"column_header_drag_drop", "move_columns", "end_move_columns", "end_column_header_drag_drop"`
	- `"begin_rc_delete_row", "begin_delete_rows"`
	- `"rc_delete_row", "end_rc_delete_row", "end_delete_rows", "delete_rows"`
	- `"begin_rc_delete_column", "begin_delete_columns"`
	- `"rc_delete_column", "end_rc_delete_column","end_delete_columns", "delete_columns"`
	- `"begin_rc_insert_column", "begin_insert_column", "begin_insert_columns", "begin_add_column","begin_rc_add_column", "begin_add_columns"`
	- `"rc_insert_column", "end_rc_insert_column", "end_insert_column", "end_insert_columns", "rc_add_column", "end_rc_add_column", "end_add_column", "end_add_columns"`
	- `"begin_rc_insert_row", "begin_insert_row", "begin_insert_rows", "begin_rc_add_row", "begin_add_row", "begin_add_rows"`
    - `"rc_insert_row", "end_rc_insert_row", "end_insert_row", "end_insert_rows", "rc_add_row", "end_rc_add_row", "end_add_row", "end_add_rows"`
    - `"row_height_resize"`
    - `"column_width_resize"`
	- `"cell_select"`
	- `"select_all"`
	- `"row_select"`
	- `"column_select"`
	- `"drag_select_cells"`
	- `"drag_select_rows"`
	- `"drag_select_columns"`
	- `"shift_cell_select"`
	- `"shift_row_select"`
	- `"shift_column_select"`
	- `"deselect"`
	- `"all_select_events", "select", "selectevents", "select_events"`
    - `"all_modified_events", "sheetmodified", "sheet_modified" "modified_events", "modified"`
	- `"bind_all"`
	- `"unbind_all"`
- `func` argument is the function you want to send the binding event to.
- Using one of the following `"all_modified_events"`, `"sheetmodified"`, `"sheet_modified"`, `"modified_events"`, `"modified"` will make any insert, delete or cell edit including pastes and undos send an event to your function. Please **note** that this will mean your function will have to return a value to use for cell edits unless the setting `edit_cell_validation` is `False`.
- For events `"begin_move_columns"`/`"begin_move_rows"` the point where columns/rows will be moved to will be under the `event_data` key `"value"`.

**For tksheet versions earlier than `7.0.0`:**
- Upon an event being triggered the bound function will be sent a [namedtuple](https://docs.python.org/3/library/collections.html#collections.namedtuple) containing variables relevant to that event, use `print()` or similar to see all the variable names in the event. Each event contains different variable names with the exception of `eventname` e.g. `event.eventname`.

**For tksheet versions >= `7.0.0`:**

#### **Event Data**
- Using `extra_bindings()` the function you bind needs to have at least one argument which will receive a `dict` which has the following layout:

```python
{
"eventname": "",
"sheetname": "",
"cells": {
    "table": {},
    "header": {},
    "index": {},
},
"moved": {
    "rows": {},
    "columns": {},
},
"added": {
    "rows": {},
    "columns": {},
},
"deleted": {
    "rows": {},
    "columns": {},
    "header": {},
    "index": {},
    "column_widths": {},
    "row_heights": {},
    "options": {},
    "displayed_columns": None,
    "displayed_rows": None,
},
"named_spans": {},
"selection_boxes": {},
"selected": tuple(),
"being_selected": tuple(),
"data": [],
"key": "",
"value": None,
"loc": tuple(),
"resized": {
    "rows": {},
    "columns": {},
}
}
```

Keys:
- Key **`["eventname"]`** will be one of the following:
    - `"begin_ctrl_c"`
    - `"end_ctrl_c"`
    - `"begin_ctrl_x"`
    - `"end_ctrl_x"`
    - `"begin_ctrl_v"`
    - `"end_ctrl_v"`
    - `"begin_delete"`
    - `"end_delete"`
    - `"begin_undo"`
    - `"end_undo"`
    - `"begin_add_columns"`
    - `"end_add_columns"`
    - `"begin_add_rows"`
    - `"end_add_rows"`
    - `"begin_delete_columns"`
    - `"end_delete_columns"`
    - `"begin_delete_rows"`
    - `"end_delete_rows"`
    - `"begin_edit_table"`
    - `"end_edit_table"`
    - `"begin_edit_index"`
    - `"end_edit_index"`
    - `"begin_edit_header"`
    - `"end_edit_header"`
    - `"begin_move_rows"`
    - `"end_move_rows"`
    - `"begin_move_columns"`
    - `"end_move_columns"`
    - `"select"`
    - `"resize"`
- For events `"begin_move_columns"`/`"begin_move_rows"` the point where columns/rows will be moved to will be under the `event_data` key `"value"`.
- Key **`["sheetname"]`** is the [name given to the sheet widget on initialization](https://github.com/ragardner/tksheet/wiki/Version-7#initialization-options), useful if you have multiple sheets to determine which one emitted the event.
- Key **`["cells"]["table"]`** if any table cells have been modified by cut, paste, delete, cell editors, dropdown boxes, check boxes, undo or redo this will be a `dict` with `tuple` keys of `(data row index: int, data column index: int)` and the values will be the cell values at that location **prior** to the change. The `dict` will be empty if no such changes have taken place.
- Key **`["cells"]["header"]`** if any header cells have been modified by cell editors, dropdown boxes, check boxes, undo or redo this will be a `dict` with keys of `int: data column index` and the values will be the cell values at that location **prior** to the change. The `dict` will be empty if no such changes have taken place.
- Key **`["cells"]["index"]`** if any index cells have been modified by cell editors, dropdown boxes, check boxes, undo or redo this will be a `dict` with keys of `int: data column index` and the values will be the cell values at that location **prior** to the change. The `dict` will be empty if no such changes have taken place.
- Key **`["moved"]["rows"]`** if any rows have been moved by dragging and dropping or undoing/redoing of dragging and dropping rows this will be a `dict` with the following keys:
    - `{"data": {old data index: new data index, ...}, "displayed": {old displayed index: new displayed index, ...}}`
        - `"data"` will be a `dict` where the keys are the old data indexes of the rows and the values are the data indexes they have moved to.
        - `"displayed"` will be a `dict` where the keys are the old displayed indexes of the rows and the values are the displayed indexes they have moved to.
        - If no rows have been moved the `dict` under `["moved"]["rows"]` will be empty.
    - For events `"begin_move_rows"` the point where rows will be moved to will be under the `event_data` key `"value"`.
- Key **`["moved"]["columns"]`** if any columns have been moved by dragging and dropping or undoing/redoing of dragging and dropping columns this will be a `dict` with the following keys:
    - `{"data": {old data index: new data index, ...}, "displayed": {old displayed index: new displayed index, ...}}`
        - `"data"` will be a `dict` where the keys are the old data indexes of the columns and the values are the data indexes they have moved to.
        - `"displayed"` will be a `dict` where the keys are the old displayed indexes of the columns and the values are the displayed indexes they have moved to.
        - If no columns have been moved the `dict` under `["moved"]["columns"]` will be empty.
    - For events `"begin_move_columns"` the point where columns will be moved to will be under the `event_data` key `"value"`.
- Key **`["added"]["rows"]`** if any rows have been added by the inbuilt popup menu insert rows or by a paste which expands the sheet then this will be a `dict` with the following keys:
    - `{"data_index": int, "displayed_index": int, "num": int, "displayed": []}`
        - `"data_index"` is an `int` representing the row where the rows were added in the data.
        - `"displayed_index"` is an `int` representing the displayed table index where the rows were added (which will be different from the data index if there are hidden rows).
        - `"displayed"` is simply a copied list of the `Sheet()`s displayed rows immediately prior to the change.
        - If no rows have been added the `dict` will be empty.
- Key **`["added"]["columns"]`** if any columns have been added by the inbuilt popup menu insert columns or by a paste which expands the sheet then this will be a `dict` with the following keys:
    - `{"data_index": int, "displayed_index": int, "num": int, "displayed": []}`
        - `"data_index"` is an `int` representing the column where the columns were added in the data.
        - `"displayed_index"` is an `int` representing the displayed table index where the columns were added (which will be different from the data index if there are hidden columns).
        - `"displayed"` is simply a copied list of the `Sheet()`s displayed columns immediately prior to the change.
        - If no columns have been added the `dict` will be empty.
- Key **`["deleted"]["columns"]`** if any columns have been deleted by the inbuilt popup menu delete columns or by undoing a paste which added columns then this will be a `dict`. This `dict` will look like the following:
    - `{[column data index]: {[row data index]: cell value, [row data index]: cell value}, [column data index]: {...} ...}`
    - If no columns have been deleted then the `dict` value for `["deleted"]["columns"]` will be empty.
- Key **`["deleted"]["rows"]`** if any rows have been deleted by the inbuilt popup menu delete rows or by undoing a paste which added rows then this will be a `dict`. This `dict` will look like the following:
    - `{[row data index]: {[column data index]: cell value, [column data index]: cell value}, [row data index]: {...} ...}`
    - If no rows have been deleted then the `dict` value for `["deleted"]["rows"]` will be empty.
- Key **`["deleted"]["header"]`** if any header values have been deleted by the inbuilt popup menu delete columns or by undoing a paste which added columns and header values then this will be a `dict`. This `dict` will look like the following:
    - `{[column data index]: header cell value, [column data index]: header cell value, ...}`
    - If no columns have been deleted by the mentioned methods then the `dict` value for `["deleted"]["header"]` will be empty.
- Key **`["deleted"]["index"]`** if any index values have been deleted by the inbuilt popup menu delete rows or by undoing a paste which added rows and index values then this will be a `dict`. This `dict` will look like the following:
    - `{[row data index]: index cell value, [row data index]: index cell value, ...}`
    - If no index values have been deleted by the mentioned methods then the `dict` value for `["deleted"]["index"]` will be empty.
- Key **`["deleted"]["column_widths"]`** if any columns have been deleted by the inbuilt popup menu delete columns or by undoing a paste which added columns then this will be a `dict`. This `dict` will look like the following:
    - `{[column data index]: column width, [column data index]: column width, ...}`
    - If no columns have been deleted then the `dict` value for `["deleted"]["column_widths"]` will be empty.
- Key **`["deleted"]["row_heights"]`** if any rows have been deleted by the inbuilt popup menu delete rows or by undoing a paste which added rows then this will be a `dict`. This `dict` will look like the following:
    - `{[row data index]: row height, [row data index]: row height, ...}`
    - If no rows have been deleted then the `dict` value for `["deleted"]["row_heights"]` will be empty.
- Key **`["deleted"]["displayed_columns"]`**  if any columns have been deleted by the inbuilt popup menu delete columns or by undoing a paste which added columns then this will be a `list`. This `list` stores the displayed columns (the columns that are showing when others are hidden) immediately prior to the change.
- Key **`["deleted"]["displayed_rows"]`**  if any rows have been deleted by the inbuilt popup menu delete rows or by undoing a paste which added rows then this will be a `list`. This `list` stores the displayed rows (the rows that are showing when others are hidden) immediately prior to the change.
- Key **`["named_spans"]`** This `dict` serves as storage for the `Sheet()`s named spans. Each value in the `dict` is a pickled `span` object.
- Key **`["options"]`** This serves as storage for the `Sheet()`s options such as highlights, formatting, alignments, dropdown boxes, check boxes etc. It is a Python pickled `dict` where the values are the sheets internal cell/row/column options `dicts`.
- Key **`["selection_boxes"]`** the value of this is all selection boxes on the sheet in the form of a `dict` as shown below:
    - For every event except `"select"` events the selection boxes are those immediately prior to the modification, for `"select"` events they are the current selection boxes.
    - The layout is always: `"selection_boxes": {(start row, start column, up to but not including row, up to but not including column): selection box type}`.
        - The row/column indexes are `int`s and the selection box type is a `str` either `"cells"`, `"rows"` or `"columns"`.
    - The `dict` will be empty if there is nothing selected.
- Key **`["selected"]`** the value of this when there is something selected on the sheet is a `namedtuple` which contains the values: `(row: int, column: int, type_: str (either "cell", "row" or "column"), tags: tuple)`
    - The `tags` in this `namedtuple` are the tags of the rectangle on the canvas. They are the following:
        - Index `[0]` - `"selected"`.
        - Index `[1]` - `f"{start row}_{start column}_{up to row}_{up to column}"` - the dimensions of the box it's attached to.
        - Index `[2]` - `int` - the canvas id of the box it's attached to.
        - Index `[3]` - `f"{current box row}_{current box column}"` - the displayed position of currently selected box.
        - Index `[4]` - `f"type_{type_}"` - the type of the box it's attached to (either "cells", "rows" or "columns").
    - When nothing is selected or the event is not relevant to the currently selected box, such as a resize event it will be an empty `tuple`.
- Key **`["being_selected"]`** if any selection box is in the process of being drawn by holding down mouse button 1 and dragging then this will be a tuple with the following layout:
    - `(start row, start column, up to but not including row, up to but not including column, selection box type)`.
        - The selection box type is a `str` either `"cells"`, `"rows"` or `"columns"`.
    - If no box is in the process of being created then this will be a an empty `tuple`.
    - [See here](https://github.com/ragardner/tksheet/wiki/Version-7#example-displaying-selections) for an example.
- Key **`["data"]`** is primarily used for `paste` and it will contain the pasted data if any.
- Key **`["key"]`** - `str` - is primarily used for cell edit events where a key press has occurred. For `"begin_edit..."` events the value is the actual key which was pressed (or `"??"` for using the mouse to open a cell). It also might be one of the following for end edit events:
    - `"Return"` - enter key.
    - `"FocusOut"` - the editor or box lost focus, perhaps by mouse clicking elsewhere.
    - `"Tab"` - tab key.
- Key **`["value"]`** is used primarily by cell editing events. For `"begin_edit..."` events it's the value displayed in he text editor when it opens. For `"end_edit..."` events it's the value in the text editor when it was closed, for example by hitting `Return`. It also used by `"begin_move_columns"`/`"begin_move_rows"` - the point where columns/rows will be moved to will be under the `event_data` key `"value"`.
- Key **`["loc"]`** is for cell editing events to show the displayed (not data) coordinates of the event. It will be **either:**
    - A tuple of `(int displayed row index, int displayed column index)` in the case of editing table cells.
    - A single `int` in the case of editing index/header cells.
- Key **`["resized"]["rows"]`** is for row height resizing events, it will be a `dict` with the following layout:
    - `{int displayed row index: {"old_size": old_height, "new_size": new_height}}`.
    - If no rows have been resized then the value for `["resized"]["rows"]` will be an empty `dict`.
- Key **`["resized"]["columns"]`** is for column width resizing events, it will be a `dict` with the following layout:
    - `{int displayed column index: {"old_size": old_width, "new_size": new_width}}`.
    - If no columns have been resized then the value for `["resized"]["columns"]` will be an empty `dict`.

___

#### **Bind tkinter events**

With this function you can bind things in the usual way you would in tkinter and they will bind to all the `tksheet` canvases. There are also two special `tksheet` events you can bind, `"<<SheetModified>>"` and `"<<SheetRedrawn>>"`.
```python
bind(
    event: str,
    func: Callable,
    add: str | None = None,
)
```
- `add` may or may not work for various bindings depending on whether they are already in use by `tksheet`.
- **Note** that while a bound event after a paste/undo/redo might have the event name `"edit_table"` it also might have added/deleted rows/columns, refer to the docs on the event data `dict` for more info.
- `event` the two emitted events are:
    - `"<<SheetModified>>"` emitted whenever the sheet was modified by the end user by editing cells or adding or deleting rows/columns. The function you bind to this event must be able to receive a `dict` argument which will be the same as [the event data dict](https://github.com/ragardner/tksheet/wiki/Version-7#event-data) but with less specific event names. The possible event names are listed below:
        - `"edit_table"` when a user has cut, paste, delete or any cell edits including using dropdown boxes etc. in the table.
        - `"edit_index"` when a user has edited a index cell.
        - `"edit_header"` when a user has edited a header cell.
        - `"add_columns"` when a user has inserted columns.
        - `"add_rows"` when a user has inserted rows.
        - `"delete_columns"` when a user has deleted columns.
        - `"delete_rows"` when a user has deleted rows.
        - `"move_columns"` when a user has dragged and dropped columns.
        - `"move_rows"` when a user has dragged and dropped rows.
    - `"<<SheetRedrawn>>"` emitted whenever the sheet GUI was refreshed (redrawn). The data for this event will be different than the usual event data, it is simply:
        - `{"sheetname": name of your sheet, "header": bool True if the header was redrawn, "row_index": bool True if the index was redrawn, "table": bool True if the the table was redrawn}`

Example:
```python
# self.sheet_was_modified is your function
self.sheet.bind("<<SheetModified>>", self.sheet_was_modified)
```

___

With this function you can unbind things you have bound using the `bind()` function.
```python
unbind(binding)
```

___

#### **Add commands to the in-built right click popup menu**
```python
popup_menu_add_command(label, func, table_menu = True, index_menu = True, header_menu = True)
```

___

#### **Remove commands added using `popup_menu_add_command()` from the in-built right click popup menu**
```python
popup_menu_del_command(label = None)
```
- If `label` is `None` then it removes all.

___

#### **Enable or disable mousewheel, left click etc**
```python
basic_bindings(enable = False)
```

___

#### **Enable or disable cell edit functionality, including Undo**
```python
edit_bindings(enable = False)
```

___

#### **Enable or disable the ability to edit a specific cell using the inbuilt text editor**
```python
cell_edit_binding(enable = False, keys = [])
```
- `keys` can be used to bind more keys to open a cell edit window.

___

```python
cut(event = None)
copy(event = None)
paste(event = None)
delete(event = None)
undo(event = None)
```

---
# **Span Objects**

In `tksheet` versions > `7` there are functions which utilise an object named `Span`. These objects are a subclass of `dict` but with various additions and dot notation attribute access.

Spans basically represent an **contiguous** area of the sheet. They can be **one** of three **kinds**:
- `"cell"`
- `"row"`
- `"column"`

They can be used with some of the sheets functions such as data getting/setting and creation of things on the sheet such as dropdown boxes.

Spans store:
- A reference to the `Sheet()` they were created with.
- Variables which represent a particular range of cells and properties for accessing these ranges.
- Variables which represent options for those cells.
- Methods which can modify the above variables.
- Methods which can act upon the table using the above variables such as `highlight`, `format`, etc.

Whether cells, rows or columns are affected will depend on the spans [`kind`](https://github.com/ragardner/tksheet/wiki/Version-7#get-a-spans-kind).

### **Creating a span**

You can create a span by:

- Using the `span()` function e.g. `sheet.span("A1")` represents the cell `A1`

**or**

- Using square brackets on a Sheet object e.g. `sheet["A1"]` represents the cell `A1`

Both methods return the created span object.

```python
span(
    *key: CreateSpanTypes,
    type_: str = "",
    name: str = "",
    table: bool = True,
    index: bool = False,
    header: bool = False,
    tdisp: bool = False,
    idisp: bool = True,
    hdisp: bool = True,
    transposed: bool = False,
    ndim: int = 0,
    convert: object = None,
    undo: bool = False,
    widget: object = None,
    expand: None | str = None,
    formatter_options: dict | None = None,
    **kwargs,
)
"""
Create a span / get an existing span by name
Returns the created span
"""
```
- `key` you do not have to provide an argument for `key`, if no argument is provided then the span will be a full sheet span. Otherwise `key` can be the following types which are type hinted as `CreateSpanTypes`:
    - `None`
    - `str` e.g. `sheet.span("A1:F1")`
    - `int` e.g. `sheet.span(0)`
    - `slice` e.g. `sheet.span(slice(0, 4))`
    - `Sequence[int | None, int | None]` representing a cell of `row, column` e.g. `sheet.span(0, 0)`
    - `Sequence[Sequence[int | None, int | None], Sequence[int | None, int | None]]` representing `sheet.span(start row, start column, up to but not including row, up to but not including column)` e.g. `sheet.span(0, 0, 2, 2)`
    - `Span` e.g `sheet.span(another_span)`
- `type_` (`str`) must be either an empty string `""` or one of the following: `"format"`, `"highlight"`, `"dropdown"`, `"checkbox"`, `"readonly"`, `"align"`.
- `name` (`str`) used for named spans or for identification. If no name is provided then a name is generated for the span which is based on an internal integer ticker and then converted to a string in the same way column names are, e.g. `0` is `"A"`.
- `table` (`bool`) when `True` will make all functions used with the span target the main table as well as the header/index if those are `True`.
- `index` (`bool`) when `True` will make all functions used with the span target the index as well as the table/header if those are `True`.
- `header` (`bool`) when `True` will make all functions used with the span target the header as well as the table/index if those are `True`.
- `tdisp` (`bool`) is used by data getting functions that utilize spans and when `True` the function retrieves screen displayed data for the table, not underlying cell data.
- `idisp` (`bool`) is used by data getting functions that utilize spans and when `True` the function retrieves screen displayed data for the index, not underlying cell data.
- `hdisp` (`bool`) is used by data getting functions that utilize spans and when `True` the function retrieves screen displayed data for the header, not underlying cell data.
- `transposed` (`bool`) is used by data getting and setting functions that utilize spans. When `True`:
    - Returned sublists from data getting functions will represent columns rather than rows.
    - Data setting functions will assume that a single sequence is a column rather than row and that a list of lists is a list of columns rather than a list of rows.
- `ndim` (`int`) is used by data getting functions that utilize spans, it must be either `0` or `1` or `2`.
    - `0` is the default setting which will make the return value vary based on what it is. For example if the gathered data is only a single cell it will return a value instead of a list of lists with a single list containing a single value. A single row will be a single list.
    - `1` will force the return of a single list as opposed to a list of lists.
    - `2` will force the return of a list of lists.
- `convert` (`None`, `Callable`) can be used to modify the data using a function before returning it. The data sent to the `convert` function will be as it was before normally returning (after `ndim` has potentially modified it).
- `undo` (`bool`) is used by data modifying functions that utilize spans. When `True` and if undo is enabled for the sheet then the end user will be able to undo/redo the modification.
- `widget` (`object`) is the reference to the original sheet which created the span. This can be changed to a different sheet if required e.g. `my_span.widget = new_sheet`.
- `expand` (`None`, `str`) must be either `None` or:
    - `"table"`/`"both"` expand the span both down and right from the span start to the ends of the table.
    - `"right"` expand the span right to the end of the table `x` axis.
    - `"down"` expand the span downwards to the bottom of the table `y` axis.
- `formatter_options` (`dict`, `None`) must be either `None` or `dict`. If providing a `dict` it must be the same structure as used in format functions, see [here](https://github.com/ragardner/tksheet/wiki/Version-7#data-formatting) for more information. Used to turn the span into a format type span which:
    - When using `get_data()` will format the returned data.
    - When using `set_data()` will format the data being set but **NOT** create a new formatting rule on the sheet.
- `**kwargs` you can provide additional keyword arguments to the function for example those used in `span.highlight()` or `span.dropdown()` which are used when applying a named span to a table.

To create a named span see [here](https://github.com/ragardner/tksheet/wiki/Version-7#named-spans).

#### **Span creation syntax**

**When creating a span using the below methods:**
- `str`s use excel syntax and the indexing rule of up to **AND** including.
- `int`s use python syntax and the indexing rule of up to but **NOT** including.

For example python index `0` as in `[0]` is the first whereas excel index `1` as in `"A1"` is the first.

If you need to convert python indexes into column letters you can use the function `num2alpha` importable from `tksheet`:

```python
from tksheet import (
    Sheet,
    num2alpha as n2a,
)

# column index five as a letter
n2a(5)
```

#### **Span creation examples using square brackets**

```python
"""
EXAMPLES USING SQUARE BRACKETS
"""

span = sheet[0] # first row
span = sheet["1"] # first row

span = sheet[0:2] # first two rows
span = sheet["1:2"] # first two rows

span = sheet[:] # entire sheet
span = sheet[":"] # entire sheet

span = sheet[:2] # first two rows
span = sheet[":2"] # first two rows

""" THESE TWO HAVE DIFFERENT OUTCOMES """
span = sheet[2:] # all rows after and not inlcuding python index 1
span = sheet["2:"] # all rows after and not including python index 0

span = sheet["A"] # first column
span = sheet["A:C"] # first three columns

""" SOME CELL AREA EXAMPLES """
span = sheet["A1:C1"] # cells A1, B1, C1
span = sheet[0, 0, 1, 3] # cells A1, B1, C1
span = sheet[(0, 0, 1, 3)] # cells A1, B1, C1
span = sheet[(0, 0), (1, 3)] # cells A1, B1, C1
span = sheet[((0, 0), (1, 3))] # cells A1, B1, C1

span = sheet["A1:2"]
span = sheet[0, 0, 2, None]
"""
["A1:2"]
All the cells starting from (0, 0)
expanding down to include row 1
but not including cells beyond row
1 and expanding out to include all
columns

    A   B   C   D
1   x   x   x   x
2   x   x   x   x
3
4
...
"""

span = sheet["A1:B"]
span = sheet[0, 0, None, 2]
"""
["A1:B"]
All the cells starting from (0, 0)
expanding out to include column 1
but not including cells beyond column
1 and expanding down to include all
rows

    A   B   C   D
1   x   x
2   x   x
3   x   x
4   x   x
...
"""

""" GETTING AN EXISTING NAMED SPAN """
# you can retrieve an existing named span quickly by surrounding its name in <> e.g.
named_span_retrieval = sheet["<the name of the span goes here>"]
```

#### **Span creation examples using sheet.span()**

```python
"""
EXAMPLES USING span()
"""

"""
USING NO ARGUMENTS
"""
sheet.span() # entire sheet, in this case not including header or index

"""
USING ONE ARGUMENT

str or int or slice()
"""

# with one argument you can use the same string syntax used for square bracket span creation
sheet.span("A1")
sheet.span(0) # row at python index 0, all columns
sheet.span(slice(0, 2)) # rows at python indexes 0 and 1, all columns
sheet.span(":") # entire sheet

"""
USING TWO ARGUMENTS
int | None, int | None

or

(int | None, int | None), (int | None, int | None)
"""
sheet.span(0, 0) # row 0, column 0 - the first cell
sheet.span(0, None) # row 0, all columns
sheet.span(None, 0) # column 0, all rows

sheet.span((0, 0), (1, 1)) # row 0, column 0 - the first cell
sheet.span((0, 0), (None, 2)) # rows 0 - end, columns 0 and 1

"""
USING FOUR ARGUMENTS
int | None, int | None, int | None, int | None
"""

sheet.span(0, 0, 1, 1) # row 0, column 0 - the first cell
sheet.span(0, 0, None, 2) # rows 0 - end, columns 0 and 1
```

### **Span properties**

Spans have a few `@property` functions:
- `span.kind`
- `span.rows`
- `span.columns`

#### **Get a spans kind**

```python
span.kind
```
- Returns either `"cell"`, `"row"` or `"column"`.

```python
span = sheet.span("A1:C4")
print (span.kind)
# prints "cell"

span = sheet.span(":")
print (span.kind)
# prints "cell"

span = sheet.span("1:3")
print (span.kind)
# prints "row"

span = sheet.span("A:C")
print (span.kind)
# prints "column"
```

#### **Get span rows and columns**

```python
span.rows
span.columns
```
Returns a `SpanRange` object. The below examples are for `span.rows` but you can use `span.columns` for the spans columns exactly the same way.

```python
# use as an iterator
span = sheet.span("A1:C4")
for row in span.rows:
    pass
# use as a reversed iterator
for row in reversed(span.rows):
    pass

# check row membership
span = sheet.span("A1:C4")
print (2 in span.rows)
# prints True

# check span.rows equality, also can do not equal
span = self.sheet["A1:C4"]
span2 = self.sheet["1:4"]
print (span.rows == span2.rows)
# prints True

# check len
span = self.sheet["A1:C4"]
print (len(span.rows))
# prints 4
```

### **Span methods**

Spans have the following methods:

#### **Modify a spans attributes**

```python
span.options(
    type_: str | None = None,
    name: str | None = None,
    table: bool | None = None,
    index: bool | None = None,
    header: bool | None = None,
    tdisp: bool | None = None,
    idisp: bool | None  = None,
    hdisp: bool | None  = None,
    transposed: bool | None = None,
    ndim: int | None = None,
    convert: Callable | None = None,
    undo: bool | None = None,
    widget: object = None,
    expand: str | None = None,
    formatter_options: dict | None = None,
    **kwargs,
)
```
**Note:** that if `None` is used for any of the following parameters then that `Span`s attribute will be unchanged.
- `type_` (`str`, `None`) if not `None` then must be either an empty string `""` or one of the following: `"format"`, `"highlight"`, `"dropdown"`, `"checkbox"`, `"readonly"`, `"align"`.
- `name` (`str`, `None`) is used for named spans or for identification.
- `table` (`bool`, `None`) when `True` will make all functions used with the span target the main table as well as the header/index if those are `True`.
- `index` (`bool`, `None`) when `True` will make all functions used with the span target the index as well as the table/header if those are `True`.
- `header` (`bool`, `None`) when `True` will make all functions used with the span target the header as well as the table/index if those are `True`.
- `tdisp` (`bool`, `None`) is used by data getting functions that utilize spans and when `True` the function retrieves screen displayed data for the table, not underlying cell data.
- `idisp` (`bool`, `None`) is used by data getting functions that utilize spans and when `True` the function retrieves screen displayed data for the index, not underlying cell data.
- `hdisp` (`bool`, `None`) is used by data getting functions that utilize spans and when `True` the function retrieves screen displayed data for the header, not underlying cell data.
- `transposed` (`bool`, `None`) is used by data getting and setting functions that utilize spans. When `True`:
    - Returned sublists from data getting functions will represent columns rather than rows.
    - Data setting functions will assume that a single sequence is a column rather than row and that a list of lists is a list of columns rather than a list of rows.
- `ndim` (`int`, `None`) is used by data getting functions that utilize spans, it must be either `0` or `1` or `2`.
    - `0` is the default setting which will make the return value vary based on what it is. For example if the gathered data is only a single cell it will return a value instead of a list of lists with a single list containing a single value. A single row will be a single list.
    - `1` will force the return of a single list as opposed to a list of lists.
    - `2` will force the return of a list of lists.
- `convert` (`Callable`, `None`) can be used to modify the data using a function before returning it. The data sent to the `convert` function will be as it was before normally returning (after `ndim` has potentially modified it).
- `undo` (`bool`, `None`) is used by data modifying functions that utilize spans. When `True` and if undo is enabled for the sheet then the end user will be able to undo/redo the modification.
- `widget` (`object`) is the reference to the original sheet which created the span. This can be changed to a different sheet if required e.g. `my_span.widget = new_sheet`.
- `expand` (`str`, `None`) must be either `None` or:
    - `"table"`/`"both"` expand the span both down and right from the span start to the ends of the table.
    - `"right"` expand the span right to the end of the table `x` axis.
    - `"down"` expand the span downwards to the bottom of the table `y` axis.
- `formatter_options` (`dict`, `None`) must be either `None` or `dict`. If providing a `dict` it must be the same structure as used in format functions, see [here](https://github.com/ragardner/tksheet/wiki/Version-7#data-formatting) for more information. Used to turn the span into a format type span which:
    - When using `get_data()` will format the returned data.
    - When using `set_data()` will format the data being set but **NOT** create a new formatting rule on the sheet.
- `**kwargs` you can provide additional keyword arguments to the function for example those used in `span.highlight()` or `span.dropdown()` which are used when applying a named span to a table.

```python
# entire sheet
span = sheet["A1"].options(expand="both")

# column A
span = sheet["A1"].options(expand="down")

# row 0
span = sheet["A1"].options(
    expand="right",
    ndim=1, # to return a single list when getting data
)
```

All of a spans modifiable attributes are listed here:
- `from_r` (`int`) represents which row the span starts at, must be a positive `int`.
- `from_c` (`int`) represents which column the span starts at, must be a positive `int`.
- `upto_r` (`int`, `None`) represents which row the span ends at, must be a positive `int` or `None`. `None` means always up to and including the last row.
- `upto_c` (`int`, `None`) represents which column the span ends at, must be a positive `int` or `None`. `None` means always up to and including the last column.
- `type_` (`str`) must be either an empty string `""` or one of the following: `"format"`, `"highlight"`, `"dropdown"`, `"checkbox"`, `"readonly"`, `"align"`.
- `name` (`str`) used for named spans or for identification. If no name is provided then a name is generated for the span which is based on an internal integer ticker and then converted to a string in the same way column names are, e.g. `0` is `"A"`.
- `table` (`bool`) when `True` will make all functions used with the span target the main table as well as the header/index if those are `True`.
- `index` (`bool`) when `True` will make all functions used with the span target the index as well as the table/header if those are `True`.
- `header` (`bool`) when `True` will make all functions used with the span target the header as well as the table/index if those are `True`.
- `tdisp` (`bool`) is used by data getting functions that utilize spans and when `True` the function retrieves screen displayed data for the table, not underlying cell data.
- `idisp` (`bool`) is used by data getting functions that utilize spans and when `True` the function retrieves screen displayed data for the index, not underlying cell data.
- `hdisp` (`bool`) is used by data getting functions that utilize spans and when `True` the function retrieves screen displayed data for the header, not underlying cell data.
- `transposed` (`bool`) is used by data getting and setting functions that utilize spans. When `True`:
    - Returned sublists from data getting functions will represent columns rather than rows.
    - Data setting functions will assume that a single sequence is a column rather than row and that a list of lists is a list of columns rather than a list of rows.
- `ndim` (`int`) is used by data getting functions that utilize spans, it must be either `0` or `1` or `2`.
    - `0` is the default setting which will make the return value vary based on what it is. For example if the gathered data is only a single cell it will return a value instead of a list of lists with a single list containing a single value. A single row will be a single list.
    - `1` will force the return of a single list as opposed to a list of lists.
    - `2` will force the return of a list of lists.
- `convert` (`None`, `Callable`) can be used to modify the data using a function before returning it. The data sent to the `convert` function will be as it was before normally returning (after `ndim` has potentially modified it).
- `undo` (`bool`) is used by data modifying functions that utilize spans. When `True` and if undo is enabled for the sheet then the end user will be able to undo/redo the modification.
- `widget` (`object`) is the reference to the original sheet which created the span. This can be changed to a different sheet if required e.g. `my_span.widget = new_sheet`.
- `kwargs` a `dict` containing keyword arguments relevant for functions such as `span.highlight()` or `span.dropdown()` which are used when applying a named span to a table.

If necessary you can also modify these attributes the same way you would an objects. e.g.

```python
span = self.sheet("A")
span.upto_c = None
```

#### **Using a span to format data**

Formats table data, see the help on [formatting](https://github.com/ragardner/tksheet/wiki/Version-7#data-formatting) for more information. Note that using this function also creates a format rule for the affected table cells.

```python
span.format(
    formatter_options={},
    formatter_class=None,
    redraw: bool = True,
    **kwargs,
)
```

Example:
```python
# using square brackets
sheet[:].format(int_formatter())

# or instead using sheet.span()
sheet.span(":").format(int_formatter())
```

These examples show the formatting of the entire sheet (not including header and index) as `int` and creates a format rule for all currently existing cells. [Named spans](https://github.com/ragardner/tksheet/wiki/Version-7#named-spans) are required to create a rule for all future existing cells as well, for example those created by the end user inserting rows or columns.

#### **Using a span to delete data format rules**

Delete any currently existing format rules for parts of the table that are covered by the span. Should not be used where there are data formatting rules created by named spans, see [Named spans](https://github.com/ragardner/tksheet/wiki/Version-7#named-spans) for more information.

```python
span.del_format()
```

Example:
```python
span1 = sheet[2:4]
span1.format(float_formatter())
span1.del_format()
```

#### **Using a span to create highlights**

```python
span.highlight(
    bg: bool | None | str = False,
    fg: bool | None | str = False,
    end: bool | None = None,
    overwrite: bool = False,
    redraw: bool = True,
)
```

Example:
```python
# highlights column A background red, text color black
sheet["A"].highlight(bg="red", fg="black")

# or

# highlights column A background red, text color black
sheet["A"].bg = "red"
sheet["A"].fg = "black"
```

#### **Using a span to delete highlights**

Delete any currently existing highlights for parts of the sheet that are covered by the span. Should not be used where there are highlights created by named spans, see [Named spans](https://github.com/ragardner/tksheet/wiki/Version-7#named-spans) for more information.

```python
span.dehighlight()
```

Example:
```python
span1 = sheet[2:4].highlight(bg="red", fg="black")
span1.dehighlight()
```

#### **Using a span to create dropdown boxes**

Creates dropdown boxes for parts of the sheet that are covered by the span. For more information see [here](https://github.com/ragardner/tksheet/wiki/Version-7#dropdown-boxes).

```python
span.dropdown(
    values: list = [],
    set_value: object = None,
    state: str = "normal",
    redraw: bool = True,
    selection_function: Callable | None = None,
    modified_function: Callable | None = None,
    search_function: Callable = dropdown_search_function,
    validate_input: bool = True,
    text: None | str = None,
)
```

Example:
```python
sheet["D"].dropdown(
    values=["on", "off"],
    set_value="off",
)
```

#### **Using a span to delete dropdown boxes**

Delete dropdown boxes for parts of the sheet that are covered by the span. Should not be used where there are dropdown box rules created by named spans, see [Named spans](https://github.com/ragardner/tksheet/wiki/Version-7#named-spans) for more information.

```python
span.del_dropdown()
```

Example:
```python
dropdown_span = sheet["D"].dropdown(values=["on", "off"],
                                    set_value="off")
dropdown_span.del_dropdown()
```

#### **Using a span to create check boxes**

Create check boxes for parts of the sheet that are covered by the span.

```python
span.checkbox(
    checked: bool = False,
    state: str = "normal",
    redraw: bool = True,
    check_function: Callable | None = None,
    text: str = "",
)
```

Example:
```python
sheet["D"].checkbox(
    checked=True,
    text="Switch",
)
```

#### **Using a span to delete check boxes**

Delete check boxes for parts of the sheet that are covered by the span. Should not be used where there are check box rules created by named spans, see [Named spans](https://github.com/ragardner/tksheet/wiki/Version-7#named-spans) for more information.

```python
span.del_checkbox()
```

Example:
```python
checkbox_span = sheet["D"].checkbox(checked=True,
                                    text="Switch")
checkbox_span.del_checkbox()
```

#### **Using a span to set cells to read only**

Create a readonly rule for parts of the table that are covered by the span.

```python
span.readonly(readonly: bool = True)
```
- Using `span.readonly(False)` deletes any existing readonly rules for the span. Should not be used where there are readonly rules created by named spans, see [Named spans](https://github.com/ragardner/tksheet/wiki/Version-7#named-spans) for more information.

#### **Using a span to create text alignment rules**

Create a text alignment rule for parts of the sheet that are covered by the span.

```python
span.align(
    align: str | None,
    redraw: bool = True,
)
```
- `align` (`str`, `None`) must be either:
    - `None` - clears the alignment rule
    - `"c"`, `"center"`, `"centre"`
    - `"w"`, `"west"`, `"left"`
    - `"e"`, `"east"`, `"right"`

Example:
```python
sheet["D"].align("right")
```

#### **Using a span to delete text alignment rules**

Delete text alignment rules for parts of the sheet that are covered by the span. Should not be used where there are alignment rules created by named spans, see [Named spans](https://github.com/ragardner/tksheet/wiki/Version-7#named-spans) for more information.

```python
span.del_align()
```

Example:
```python
align_span = sheet["D"].align("right")
align_span.del_align()
```

#### **Using a span to clear cells**

Clear cell data from all cells that are covered by the span.

```python
span.clear(
    undo: bool | None = None,
    redraw: bool = True,
)
```
- `undo` (`bool`, `None`) When `True` if undo is enabled for the end user they will be able to undo the clear change.

Example:
```python
# clears column D
sheet["D"].clear()
```

#### **Set the spans orientation**

The attribute `span.transposed` (`bool`) is used by data getting and setting functions that utilize spans. When `True`:
    - Returned sublists from data getting functions will represent columns rather than rows.
    - Data setting functions will assume that a single sequence is a column rather than row and that a list of lists is a list of columns rather than a list of rows.

You can toggle the transpotition of the span by using:

```python
span.transpose()
```

If the attribute is already `True` this makes it `False` and vice versa.

```python
span = sheet["A:D"].transpose()
# this span is now transposed
print (span.transposed)
# prints True

span.transpose()
# this span is no longer transposed
print (span.transposed)
# prints False
```

#### **Expand the spans area**

Expand the spans area either all the way to the right (x axis) or all the way down (y axis) or both.

```python
span.expand(direction: str = "both")
```
- `direction` (`None`, `str`) must be either `None` or:
    - `"table"`/`"both"` expand the span both down and right from the span start to the ends of the table.
    - `"right"` expand the span right to the end of the table x axis.
    - `"down"` expand the span downwards to the bottom of the table y axis.

---
# **Named Spans**

Named spans are like spans but with a type, some keyword arguments saved in `span.kwargs` and then created by using a `Sheet()` function. Like spans, named spans are also **contiguous** areas of the sheet.

Named spans can be used to:
- Create options (rules) for the sheet which will expand/contract when new cells are added/removed. For example if a user were to insert rows in the middle of some already highlighted rows:
    - With ordinary row highlights the newly inserted rows would **NOT** be highlighted.
    - With named span row highlights the newly inserted rows would also be highlighted.
- Quickly delete an existing option from the table whereas an ordinary span would not keep track of where the options have been moved.

**Note** that generally when a user moves rows/columns around the dimensions of the named span essentially move with either end of the span:
- The new start of the span will be wherever the start row/column moves.
- The new end of the span will be wherever the end row/column moves.
The exceptions to this rule are when a span is expanded or has been created with `None`s or the start of `0` and no end or end of `None`.

For the end user, when a span is just a single row/column (and is not expanded/unlimited) it cannot be expanded but it can be deleted if the row/column is deleted.

#### **Creating a named span**

For a span to become a named span it needs:
- One of the following `type_`s: `"format"`, `"highlight"`, `"dropdown"`, `"checkbox"`, `"readonly"`, `"align"`.
- Relevant keyword arguments e.g. if the `type_` is `"highlight"` then arguments for `sheet.highlight()` found [here](https://github.com/ragardner/tksheet/wiki/Version-7#highlighting-cells).

After a span has the above items the following function has to be used to make it a named span and create the options on the sheet:

```python
named_span(span: Span)
"""
Adds a named span to the sheet
Returns the span
"""
```
- `span` must be an existing span with:
    - a `name` (a `name` is automatically generated upon span creation if one is not provided).
    - a `type_` as described above.
    - keyword arguments as described above.

Example of creating a named span which will always keep the entire sheet formatted as `int` no matter how many rows/columns are inserted:
```python
span = self.sheet.span(
    ":",
    # you don't have to provide a `type_` when using the `formatter_kwargs` argument
    formatter_options=int_formatter(),
)
self.sheet.named_span(span)
```

#### **Deleting a named span**

To delete a named span you simply have to provide the name.

```python
del_named_span(name: str)
```

Example, creating and deleting a span:
```python
# span covers the entire sheet
self.sheet.named_span(
    self.sheet.span(
        name="my highlight span",
        type_="highlight",
        bg="dark green",
        fg="#FFFFFF",
    )
)
self.sheet.del_named_span("my highlight span")

# ValueError is raised if name does not exist
self.sheet.del_named_span("this name doesnt exist")
# ValueError: Span 'B' does not exist.
```

#### **Other named span functions**

Sets the `Sheet`s internal dict of named spans:
```python
set_named_spans(named_spans: None | dict = None) -> Sheet
```
- Using `None` deletes all existing named spans

Get an existing named span:
```python
get_named_span(name: str) -> dict
```

Get all existing named spans:
```python
get_named_spans() -> dict
```

---
# **Getting Sheet Data**

#### **Using a span to get sheet data**

A `Span` object (more information [here](https://github.com/ragardner/tksheet/wiki/Version-7#span-objects)) is returned when using square brackets on a `Sheet` like so:

```python
span = self.sheet["A1"]
```

You can also use `sheet.span()`:

```python
span = self.sheet.span("A1")
```

The above span represents the cell `A1` - row 0, column 0. A reserved span attribute named `data` can then be used to retrieve the data for cell `A1`, example below:

```python
span = self.sheet["A1"]
cell_a1_data = span.data
```

The data that is retrieved entirely depends on the area the span represents. You can also use `span.value` to the same effect.

There are certain other span attributes which have an impact on the data returned, explained below:
- `table` (`bool`) when `True` will make all functions used with the span target the main table as well as the header/index if those are `True`.
- `index` (`bool`) when `True` will make all functions used with the span target the index as well as the table/header if those are `True`.
- `header` (`bool`) when `True` will make all functions used with the span target the header as well as the table/index if those are `True`.
- `tdisp` (`bool`) when `True` the function retrieves screen displayed data for the table, not underlying cell data.
- `idisp` (`bool`) when `True` the function retrieves screen displayed data for the index, not underlying cell data.
- `hdisp` (`bool`) when `True` the function retrieves screen displayed data for the header, not underlying cell data.
- `transposed` (`bool`) is used by data getting and setting functions that utilize spans. When `True`:
    - Returned sublists from **data getting** functions will represent columns rather than rows.
    - Data setting functions will assume that a single sequence is a column rather than row and that a list of lists is a list of columns rather than a list of rows.
- `ndim` (`int`) is used by data getting functions that utilize spans, it must be either `0` or `1` or `2`.
    - `0` is the default setting which will make the return value vary based on what it is. For example if the gathered data is only a single cell it will return a value instead of a list of lists with a single list containing a single value. A single row will be a single list.
    - `1` will force the return of a single list as opposed to a list of lists.
    - `2` will force the return of a list of lists.
- `convert` (`None`, `Callable`) can be used to modify the data using a function before returning it. The data sent to the `convert` function will be as it was before normally returning (after `ndim` has potentially modified it).
- `widget` (`object`) is the reference to the original sheet which created the span (this is the widget that data is retrieved from). This can be changed to a different sheet if required e.g. `my_span.widget = new_sheet`.

Some more complex examples of data retrieval:

```python
"single cell"
cell_a1_data = self.sheet["A1"].data

"entire sheet including headers and index"
entire_sheet_data = self.sheet["A1"].expand().options(header=True, index=True).data

"header data, no table or index data"
# a list of displayed header cells
header_data = self.sheet["A:C"].options(table=False, header=True).data

# a header value
header_data = self.sheet["A"].options(table=False, hdisp=False, header=True).data

"index data, no table or header data"
# a list of displayed index cells
index_data = self.sheet[:3].options(table=False, index=True).data

# or using sheet.span() a list of displayed index cells
index_data = self.sheet.span(slice(None, 3), table=False, index=True).data

# a row index value
index_data = self.sheet[3].options(table=False, idisp=False, index=True).data

"sheet data as columns instead of rows, with actual header data"
sheet_data = self.sheet[:].transpose().options(hdisp=False, header=True).data

# or instead using sheet.span() with only kwargs
sheet_data = self.sheet.span(transposed=True, hdisp=False, header=True).data
```

There is also a `Sheet()` function for data retrieval (it is used internally by the above data getting methods):

```python
sheet.get_data(
    *key: CreateSpanTypes,
) -> object
```

Examples:
```python
data = self.sheet.get_data("A1")
data = self.sheet.get_data(0, 0, 3, 3)
data = self.sheet.get_data(
    self.sheet.span(":D", transposed=True)
)
```

___

#### **Generate sheet rows one at a time**

This function is useful if you need a lot of sheet data, and can use one row at a time (may save memory use in certain scenarios). It does not use spans.

```python
yield_sheet_rows(get_displayed = False,
                 get_header = False,
                 get_index = False,
                 get_index_displayed = True,
                 get_header_displayed = True,
                 only_rows = None,
                 only_columns = None)
```
Note:
- The following keyword arguments both behave the same way for `yield_sheet_rows()` and `get_sheet_data()`.

Parameters:
- `get_displayed` (`bool`) if `True` it will return cell values as they are displayed on the screen. If `False` it will return any underlying data, for example if the cell is formatted.
- `get_header` (`bool`) if `True` it will return the header of the sheet even if there is not one.
- `get_index` (`bool`) if `True` it will return the index of the sheet even if there is not one.
- `get_index_displayed` (`bool`) if `True` it will return whatever index values are displayed on the screen, for example if there is a dropdown box with `text` set.
- `get_header_displayed` (`bool`) if `True` it will return whatever header values are displayed on the screen, for example if there is a dropdown box with `text` set.
- `only_rows` (`None`, `iterable`) with this argument you can supply an iterable of row indexes in any order to be the only rows that are returned.
- `only_columns` (`None`, `iterable`) with this argument you can supply an iterable of column indexes in any order to be the only columns that are returned.

___

#### **Get table data, readonly**

```python
@property
data()
```
- e.g. `self.sheet.data`
- Doesn't include header or index data.

___

#### **The name of the actual internal sheet data variable**

```python
.MT.data
```
- You can use this to directly modify or retrieve the main table's data e.g. `cell_0_0 = my_sheet_name_here.MT.data[0][0]` but only do so if you know what you're doing.

___

#### **Sheet methods**

`Sheet` objects also have some functions similar to lists. **Note** that these functions do **not** include the header or index.

Iterate over table rows:
```python
for row in self.sheet:
    print (row)

# and in reverse
for row in reversed(self.sheet):
    print (row)
```

Check if the table has a particular value (membership):
```python
# returns True or False
search_value = "the cell value I'm looking for"
print (search_value in self.sheet)
```
- Can also check if a row is in the sheet if a `list` is used.

___

#### **Other data getting functions**

Get a single cell:
```python
get_cell_data(r, c, get_displayed = False)
```

Get a row:
```python
get_row_data(r,
             get_displayed = False,
             get_index = False,
             get_index_displayed = True,
             only_columns = None)
```

Get a column:
```python
get_column_data(c,
                get_displayed = False,
                get_header = False,
                get_header_displayed = True,
                only_rows = None)
```
- The above arguments behave the same way as for `yield_sheet_rows()`.

___

#### **Get the number of rows in the sheet**
```python
get_total_rows(include_index = False)
```

___

#### **Get the number of columns in the sheet**
```python
get_total_columns(include_header = False)
```

---
# **Setting Sheet Data**

Fundamentally, there are two ways to set table data:
- Overwriting the entire table and setting the table data to a new object.
- Modifying the existing data.

#### **Overwriting the table**

```python
set_sheet_data(data: list | tuple | None = None,
               reset_col_positions: bool = True,
               reset_row_positions: bool = True,
               redraw: bool = True,
               verify: bool = False,
               reset_highlights: bool = False,
               keep_formatting: bool = True,
               delete_options: bool = False,
)
```
- `data` (`list`) has to be a list of lists for full functionality, for display only a list of tuples or a tuple of tuples will work.
- `reset_col_positions` and `reset_row_positions` (`bool`) when `True` will reset column widths and row heights.
- `redraw` (`bool`) refreshes the table after setting new data.
- `verify` (`bool`) goes through `data` and checks if it is a list of lists, will raise error if not, disabled by default.
- `reset_highlights` (`bool`) resets all table cell highlights.
- `keep_formatting` (`bool`) when `True` re-applies any prior formatting rules to the new data, if `False` all prior formatting rules are deleted.
- `delete_options` (`bool`) when `True` all table options such as dropdowns, check boxes, formatting, highlighting etc. are deleted.

**Note:** this function does not impact the sheet header or index.

___

```python
@data.setter
data(value: object)
```
- Acts like setting an attribute e.g. `sheet.data = [[1, 2, 3], [4, 5, 6]]`
- Uses the `set_sheet_data()` function and its default arguments.

___

#### **Modifying sheet data**

A `Span` object (more information [here](https://github.com/ragardner/tksheet/wiki/Version-7#span-objects)) is returned when using square brackets on a `Sheet` like so:

```python
span = self.sheet["A1"]
```

You can also use `sheet.span()`:

```python
span = self.sheet.span("A1")
```

The above span represents the cell `A1` - row 0, column 0. A reserved span attribute named `data` (you can also use `.value`) can then be used to modify sheet data **starting** from cell `A1`, example below:

```python
span = self.sheet["A1"]
span.data = "new value for cell A1"

# or even shorter:
self.sheet["A1"].data = "new value for cell A1"

# or with sheet.span()
self.sheet.span("A1").data = "new value for cell A1"
```

If you provide a list or tuple it will set more than one cell, starting from the spans start cell. In the example below three cells are set in the first row:

```python
self.sheet["B1"].data = ["row 0, column 1 new value (B1)",
                         "row 0, column 2 new value (C1)",
                         "row 0, column 3 new value (D1)"]
```

You can set data in column orientation with a transposed span:

```python
self.sheet["B1"].transpose().data = ["row 0, column 1 new value (B1)",
                                     "row 1, column 1 new value (B2)",
                                     "row 2, column 1 new value (B3)"]
```

When setting data only a spans start cell is taken into account, the end cell is ignored. The example below demonstrates this, the spans end - `"B1"` is ignored and 4 cells get new values:

```python
self.sheet["A1:B1"].data = ["A1 new val", "B1 new val", "C1 new val", "D1 new val"]
```

These are the span attributes which have an impact on the data set:
- `table` (`bool`) when `True` will make all functions used with the span target the main table as well as the header/index if those are `True`.
- `index` (`bool`) when `True` will make all functions used with the span target the index as well as the table/header if those are `True`.
- `header` (`bool`) when `True` will make all functions used with the span target the header as well as the table/index if those are `True`.
- `transposed` (`bool`) is used by data getting and setting functions that utilize spans. When `True`:
    - Returned sublists from data getting functions will represent columns rather than rows.
    - **Data setting** functions will assume that a single sequence is a column rather than row and that a list of lists is a list of columns rather than a list of rows.
- `widget` (`object`) is the reference to the original sheet which created the span (this is the widget that data is set to). This can be changed to a different sheet if required e.g. `my_span.widget = new_sheet`.

Some more complex examples of setting data:

```python
"""
SETTING ROW DATA
"""
# first row gets some new values and the index gets a new value also
self.sheet[0].options(index=True).data = ["index val", "row 0 col 0", "row 0 col 1", "row 0 col 2"]

# or instead using sheet.span() first row gets some new values and the index gets a new value also
self.sheet.span(0, index=True).data = ["index val", "row 0 col 0", "row 0 col 1", "row 0 col 2"]

# first two rows get some new values, index included
self.sheet[0].options(index=True).data = [["index 0", "row 0 col 0", "row 0 col 1", "row 0 col 2"],
                                          ["index 1", "row 1 col 0", "row 1 col 1", "row 1 col 2"]]

"""
SETTING COLUMN DATA
"""
# first column gets some new values and the header gets a new value also
self.sheet["A"].options(transposed=True, header=True).data = ["header val", "row 0 col 0", "row 1 col 0", "row 2 col 0"]

# or instead using sheet.span() first column gets some new values and the header gets a new value also
self.sheet.span("A", transposed=True, header=True).data = ["header val", "row 0 col 0", "row 1 col 0", "row 2 col 0"]

# first two columns get some new values, header included
self.sheet["A"].options(transposed=True, header=True).data = [["header 0", "row 0 col 0", "row 1 col 0", "row 2 col 0"],
                                                              ["header 1", "row 0 col 1", "row 1 col 1", "row 2 col 1"]]

"""
SETTING CELL AREA DATA
"""
# cells B2, C2, B3, C3 get new values
self.sheet["B2"].data = [["B2 new val", "C2 new val"],
                         ["B3 new val", "C3 new val"]]

# or instead using sheet.span() cells B2, C2, B3, C3 get new values
self.sheet.span("B2").data = [["B2 new val", "C2 new val"],
                              ["B3 new val", "C3 new val"]]
```

You can also use the `Sheet` function `set_data()` using the same methodology as above.

```python
set_data(
    *key: CreateSpanTypes,
    data: object = None,
    undo: bool | None = None,
    redraw: bool = True,
) -> None
```

Example:
```python
# create a span which encompasses the entire table, header and index
# all data values, no displayed values
self.sheet_span = self.sheet.span(
    header=True,
    index=True,
    hdisp=False,
    idisp=False,
)

# set data for the span which was created above
self.sheet_span.data = [["",  "A",  "B",  "C"]
                        ["1", "A1", "B1", "C1"],
                        ["2", "A2", "B2", "C2"]]
```

#### **Insert a row into the sheet**
```python
insert_row(values = None,
           idx = "end",
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

#### **Insert a column into the sheet**
```python
insert_column(values = None,
              idx = "end",
              width = None,
              deselect_all = False,
              add_rows = True,
              equalize_data_row_lengths = True,
              mod_column_positions = True,
              redraw = False)
```

___

#### **Insert multiple columns into the sheet**
```python
insert_columns(columns = 1, idx = "end", widths = None, deselect_all = False, add_rows = True, equalize_data_row_lengths = True,
               mod_column_positions = True,
               redraw = False)
```
- `columns` can be either `int` or iterable of iterables.

___

#### **Insert multiple rows into the sheet**
```python
insert_rows(rows = 1,
            idx = "end",
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

#### **Make all data rows the same length**
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

---
# **Highlighting Cells**

#### **Creating highlights**

`Span` objects (more information [here](https://github.com/ragardner/tksheet/wiki/Version-7#span-objects)) can be used to highlight cells, rows, columns, the entire sheet, headers and the index.

You can use either of the following methods:
- Using a span method e.g. `span.highlight()` more information [here](https://github.com/ragardner/tksheet/wiki/Version-7#using-a-span-to-create-highlights).
- Using a sheet method e.g. `sheet.highlight(Span)`

Or if you need user inserted row/columns in the middle of highlight areas to also be highlighted you can use named spans, more information [here](https://github.com/ragardner/tksheet/wiki/Version-7#named-spans).

Whether cells, rows or columns are highlighted depends on the [`kind`](https://github.com/ragardner/tksheet/wiki/Version-7#get-a-spans-kind) of span.

```python
highlight(
    *key: CreateSpanTypes,
    bg: bool | None | str = False,
    fg: bool | None | str = False,
    end: bool | None = None,
    overwrite: bool = False,
    redraw: bool = True,
) -> Span
```

Parameters:
- `key` (`CreateSpanTypes`) either a span or a type which can create a span. See [here](https://github.com/ragardner/tksheet/wiki/Version-7#creating-a-span) for more information on the types that can create a span.
- `bg` and `fg` arguments use either a tkinter color or a hex `str` color.
- `end` (`bool`) is used for row highlighting where `True` makes the highlight go to the end of the Sheet window on the x axis.
- `overwrite` (`bool`) when `True` overwrites the any previous highlight for that cell/row/column, whereas `False` will only impact the keyword arguments used.
- Highlighting cells, rows or columns will also change the colors of dropdown boxes and check boxes.

Example:
```python
# highlight cell - row 3, column 5
self.sheet.highlight(
    (3, 5),
    bg="dark green",
    fg="white",
)

# or

# same cells, background red, text color black
sheet[3, 5].bg = "red"
sheet[3, 5].fg = "black"
```

___

#### **Deleting highlights**

If the highlights were created by a named span then the named span must be deleted, more information [here](https://github.com/ragardner/tksheet/wiki/Version-7#deleting-a-named-span).

Otherwise you can use either of the following methods to delete/remove highlights:
- Using a span method e.g. `span.dehighlight()` more information [here](https://github.com/ragardner/tksheet/wiki/Version-7#using-a-span-to-delete-highlights).
- Using a sheet method e.g. `sheet.dehighlight(Span)` details below:

```python
dehighlight(
    *key: CreateSpanTypes,
    redraw: bool = True,
) -> Span
```

Parameters:
- `key` (`CreateSpanTypes`) either a span or a type which can create a span. See [here](https://github.com/ragardner/tksheet/wiki/Version-7#creating-a-span) for more information on the types that can create a span.

Example:
```python
# highlight column B
self.sheet.highlight(
    "B",
    bg="dark green",
    fg="white",
)

# dehighlight column B
self.sheet.dehighlight("B")
```

---
# **Dropdown Boxes**

#### **Creating dropdown boxes**

`Span` objects (more information [here](https://github.com/ragardner/tksheet/wiki/Version-7#span-objects)) can be used to create dropdown boxes for cells, rows, columns, the entire sheet, headers and the index.

You can use either of the following methods:
- Using a span method e.g. `span.dropdown()` more information [here](https://github.com/ragardner/tksheet/wiki/Version-7#using-a-span-to-create-dropdown-boxes).
- Using a sheet method e.g. `sheet.dropdown(Span)`

Or if you need user inserted row/columns in the middle of areas with dropdown boxes to also have dropdown boxes you can use named spans, more information [here](https://github.com/ragardner/tksheet/wiki/Version-7#named-spans).

Whether dropdown boxes are created for cells, rows or columns depends on the [`kind`](https://github.com/ragardner/tksheet/wiki/Version-7#get-a-spans-kind) of span.

```python
dropdown(
    *key: CreateSpanTypes,
    values: list = [],
    set_value: object = None,
    state: str = "normal",
    redraw: bool = True,
    selection_function: Callable | None = None,
    modified_function: Callable | None = None,
    search_function: Callable = dropdown_search_function,
    validate_input: bool = True,
    text: None | str = None,
) -> Span
```

Notes:
- `selection_function`/`modified_function` (`Callable`, `None`) parameters require either `None` or a function. The function you use needs at least one argument because tksheet will send information to your function about the triggered dropdown.
- When a user selects an item from the dropdown box the sheet will set the underlying cells data to the selected item, to bind this event use either the `selection_function` argument or see the function `extra_bindings()` with binding `"end_edit_cell"` [here](https://github.com/ragardner/tksheet/wiki/Version-7#bindings-and-functionality).

Parameters:
- `key` (`CreateSpanTypes`) either a span or a type which can create a span. See [here](https://github.com/ragardner/tksheet/wiki/Version-7#creating-a-span) for more information on the types that can create a span.
- `values` are the values to appear in a list view type interface when the dropdown box is open.
- `state` determines whether or not there is also an editable text window at the top of the dropdown box when it is open.
- `redraw` refreshes the sheet so the newly created box is visible.
- `selection_function` can be used to trigger a specific function when an item from the dropdown box is selected, if you are using the above `extra_bindings()` as well it will also be triggered but after this function. e.g. `selection_function = my_function_name`
- `modified_function` can be used to trigger a specific function when the `state` of the box is set to `"normal"` and there is an editable text window and a change of the text in that window has occurred. Note that this function occurs before the dropdown boxes search feature.
- `search_function` (`None`, `callable`) sets the function that will be used to search the dropdown boxes values upon a dropdown text editor modified event when the dropdowns state is `normal`. Set to `None` to disable the search feature or use your own function with the following keyword arguments: `(search_for, data):` and make it return an row number (e.g. select and see the first value would be `0`) if positive and `None` if negative.
- `validate_input` (`bool`) when `True` will not allow cut, paste, delete or cell editor to input values to cell which are not in the dropdown boxes values.
- `text` (`None`, `str`) can be set to something other than `None` to always display over whatever value is in the cell, this is useful when you want to display a Header name over a dropdown box selection.

Example:
```python
# create dropdown boxes in column "D"
self.sheet.dropdown(
    "D",
    values=[0, 1, 2, 3, 4],
)
```

___

#### **Deleting dropdown boxes**

If the dropdown boxes were created by a named span then the named span must be deleted, more information [here](https://github.com/ragardner/tksheet/wiki/Version-7#deleting-a-named-span).

Otherwise you can use either of the following methods to delete/remove dropdown boxes.
- Using a span method e.g. `span.del_dropdown()` more information [here](https://github.com/ragardner/tksheet/wiki/Version-7#using-a-span-to-delete-dropdown-boxes).
- Using a sheet method e.g. `sheet.del_dropdown(Span)` details below:

```python
del_dropdown(
    *key: CreateSpanTypes,
    redraw: bool = True,
) -> Span
```

Parameters:
- `key` (`CreateSpanTypes`) either a span or a type which can create a span. See [here](https://github.com/ragardner/tksheet/wiki/Version-7#creating-a-span) for more information on the types that can create a span.

Example:
```python
# create dropdown boxes in column "D"
self.sheet.dropdown(
    "D",
    values=[0, 1, 2, 3, 4],
)

# delete dropdown boxes in column "D"
self.sheet.del_dropdown("D")
```

___

#### **Get chosen dropdown boxes values**

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

#### **Set the values and cell value of a chosen dropdown box**

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

#### **Set and get bound dropdown functions**

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

#### **Get a dictionary of all dropdown boxes**

```python
get_dropdowns()
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

```python
get_header_dropdowns()
```

```python
get_index_dropdowns()
```

___

#### **Open a dropdown box**

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

#### **Close a dropdown box**

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

---
# **Check Boxes**

#### **Creating check boxes**

`Span` objects (more information [here](https://github.com/ragardner/tksheet/wiki/Version-7#span-objects)) can be used to create check boxes for cells, rows, columns, the entire sheet, headers and the index.

You can use either of the following methods:
- Using a span method e.g. `span.checkbox()` more information [here](https://github.com/ragardner/tksheet/wiki/Version-7#using-a-span-to-create-check-boxes).
- Using a sheet method e.g. `sheet.checkbox(Span)`

Or if you need user inserted row/columns in the middle of areas with check boxes to also have check boxes you can use named spans, more information [here](https://github.com/ragardner/tksheet/wiki/Version-7#named-spans).

Whether check boxes are created for cells, rows or columns depends on the [`kind`](https://github.com/ragardner/tksheet/wiki/Version-7#get-a-spans-kind) of span.

```python
checkbox(
    *key: CreateSpanTypes,
    checked: bool = False,
    state: str = "normal",
    redraw: bool = True,
    check_function: Callable | None = None,
    text: str = "",
) -> Span
```

Notes:
- `check_function` (`Callable`, `None`) requires either `None` or a function. The function you use needs at least one argument because when the checkbox is clicked it will send information to your function about the clicked checkbox.
- Use `highlight_cells()` or rows or columns to change the color of the checkbox.
- Check boxes are always left aligned despite any align settings.

 Parameters:
- `key` (`CreateSpanTypes`) either a span or a type which can create a span. See [here](https://github.com/ragardner/tksheet/wiki/Version-7#creating-a-span) for more information on the types that can create a span.
- `checked` is the initial creation value to set the box to.
- `text` displays text next to the checkbox in the cell, but will not be used as data, data will either be `True` or `False`.
- `check_function` can be used to trigger a function when the user clicks a checkbox.
- `state` can be `"normal"` or `"disabled"`. If `"disabled"` then color will be same as table grid lines, else it will be the cells text color.

Example:
```python
self.sheet.checkbox(
    "D",
    checked=True,
)
```

___

#### **Deleting check boxes**

If the check boxes were created by a named span then the named span must be deleted, more information [here](https://github.com/ragardner/tksheet/wiki/Version-7#deleting-a-named-span).

Otherwise you can use either of the following methods to delete/remove check boxes:
- Using a span method e.g. `span.del_checkbox()` more information [here](https://github.com/ragardner/tksheet/wiki/Version-7#using-a-span-to-delete-check-boxes).
- Using a sheet method e.g. `sheet.del_checkbox(Span)` details below:

```python
del_checkbox(
    *key: CreateSpanTypes,
    redraw: bool = True,
) -> Span
```

Parameters:
- `key` (`CreateSpanTypes`) either a span or a type which can create a span. See [here](https://github.com/ragardner/tksheet/wiki/Version-7#creating-a-span) for more information on the types that can create a span.

Example:
```python
# creating checkboxes in column D
self.sheet.checkbox(
    "D",
    checked=True,
)

# deleting checkboxes in column D
self.sheet.del_checkbox("D")
```

#### **Set or toggle a check box**
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

#### **Get a dictionary of all check box dictionaries**
```python
get_checkboxes()
```

```python
get_header_checkboxes()
```

```python
get_index_checkboxes()
```

---
# **Data Formatting**

By default tksheet stores all user inputted data as strings and while tksheet can store and display any datatype with a `__str__()` method this has some obvious limitations.

Data formatting aims to provide greater functionality when working with different datatypes and provide strict typing for the sheet. With formatting you can convert sheet data and user input to a specific datatype.

Additionally, formatting also provides a function for displaying data on the table GUI (as a rounded float for example) and logic for handling invalid and missing data.

tksheet has several basic built-in formatters and provides functionality for creating your own custom formats as well.

A demonstration of all the built-in and custom formatters can be found [here](https://github.com/ragardner/tksheet/wiki/Version-7#example-using-and-creating-formatters).

### **Creation and deletion of data formatting rules**

#### **Creating a data format rule**

`Span` objects (more information [here](https://github.com/ragardner/tksheet/wiki/Version-7#span-objects)) can be used to format data for cells, rows, columns, the entire sheet, headers and the index.

You can use either of the following methods:
- Using a span method e.g. `span.format()` more information [here](https://github.com/ragardner/tksheet/wiki/Version-7#using-a-span-to-format-data).
- Using a sheet method e.g. `sheet.format(Span)`

Or if you need user inserted row/columns in the middle of areas with data formatting to also be formatted you can use named spans, more information [here](https://github.com/ragardner/tksheet/wiki/Version-7#named-spans).

Whether data is formatted for cells, rows or columns depends on the [`kind`](https://github.com/ragardner/tksheet/wiki/Version-7#get-a-spans-kind) of span.

```python
format(
    *key: CreateSpanTypes,
    formatter_options: dict = {},
    formatter_class: object = None,
    redraw: bool = True,
    **kwargs,
) -> Span
```

Notes:
1. When applying multiple overlapping formats with e.g. a formatted cell which overlaps a formatted row, the priority is as follows:
    - Cell formats first.
    - Row formats second.
    - Column formats third.
2. Data formatting will effectively override `validate_input = True` on cells with dropdown boxes.
3. When getting data take careful note of the `get_displayed` options, as these are the difference between getting the actual formatted data and what is simply displayed on the table GUI.

Parameters:
- `key` (`CreateSpanTypes`) either a span or a type which can create a span. See [here](https://github.com/ragardner/tksheet/wiki/Version-7#creating-a-span) for more information on the types that can create a span.
- `formatter_options` (`dict`) a dictionary of keyword options/arguements to pass to the formatter, see [here](https://github.com/ragardner/tksheet/wiki/Version-7#formatters) for information on what argument to use.
- `formatter_class` (`class`) in case you want to use a custom class to store functions and information as opposed to using the built-in methods.
- `**kwargs` any additional keyword options/arguements to pass to the formatter.

#### **Deleting a data format rule**

If the data format rule was created by a named span then the named span must be deleted, more information [here](https://github.com/ragardner/tksheet/wiki/Version-7#deleting-a-named-span).

Otherwise you can use either of the following methods to delete/remove data formatting rules:
- Using a span method e.g. `span.del_format()` more information [here](https://github.com/ragardner/tksheet/wiki/Version-7#using-a-span-to-delete-check-boxes).
- Using a sheet method e.g. `sheet.del_format(Span)` details below:

```python
del_format(
    *key: CreateSpanTypes,
    clear_values: bool = False,
    redraw: bool = True,
) -> Span
```
- `key` (`CreateSpanTypes`) either a span or a type which can create a span. See [here](https://github.com/ragardner/tksheet/wiki/Version-7#creating-a-span) for more information on the types that can create a span.
- `clear_values` (`bool`) if true, all the cells covered by the span will have their values cleared.

#### **Delete all formatting**

```python
del_all_formatting(clear_values: bool = False)
```
- `clear_values` (`bool`) if true, all the sheets cell values will be cleared.

#### **Reapply formatting to entire sheet**

```python
reapply_formatting()
```
- Useful if you have manually changed the entire sheets data using `sheet.MT.data = ` and want to reformat the sheet using any existing formatting you have set.

### **Formatters**

`tksheet` provides a number of in-built formatters, in addition to the base `formatter` function. These formatters are designed to provide a range of functionality for different datatypes. The following table lists the available formatters and their options.

**You can use any of the below formatters as an argument for the parameter `formatter_options`.

```python
formatter(
    datatypes: tuple[object] | object,
    format_function: Callable,
    to_str_function: Callable = to_str,
    invalid_value: object = "NaN",
    nullable: bool = True,
    pre_format_function: Callable | None = None,
    post_format_function: Callable | None = None,
    clipboard_function: Callable | None = None,
    **kwargs,
)
```

This is the generic formatter options interface. You can use this to create your own custom formatters. The following options are available. Note that all these options can also be passed to the `format_cell()` function as keyword arguments and are available as attributes for all formatters. You can provide functions of your own creation for all the below arguments which take functions if you require.

- `datatypes` (`list`) a list of datatypes that the formatter will accept. For example, `datatypes = [int, float]` will accept integers and floats.
- `format_function` (`function`) a function that takes a string and returns a value of the desired datatype. For example, `format_function = int` will convert a string to an integer.
- `to_str_function` (`function`) a function that takes a value of the desired datatype and returns a string. This determines how the formatter displays its data on the table. For example, `to_str_function = str` will convert an integer to a string. Defaults to `tksheet.to_str`.
- `invalid_value` (`any`) the value to return if the input string is invalid. For example, `invalid_value = "NA"` will return "NA" if the input string is invalid.
- `nullable` (`bool`) if true, the formatter will accept `None` as a valid input.
- `pre_format_function` (`function`) a function that takes a input string and returns a string. This function is called before the `format_function` and can be used to modify the input string before it is converted to the desired datatype. This can be useful if you want to strip out unwanted characters or convert a string to a different format before converting it to the desired datatype.
- `post_format_function` (`function`) a function that takes a value **which might not be of the desired datatype, e.g. `None` if the cell is nullable and empty** and if successful returns a value of the desired datatype or if not successful returns the input value. This function is called after the `format_function` and can be used to modify the output value after it is converted to the desired datatype. This can be useful if you want to round a float for example.
- `clipboard_function` (`function`) a function that takes a value of the desired datatype and returns a string. This function is called when the cell value is copied to the clipboard. This can be useful if you want to convert a value to a different format before it is copied to the clipboard.
- `**kwargs` any additional keyword options/arguements to pass to the formatter. These keyword arguments will be passed to the `format_function`, `to_str_function`, and the `clipboard_function`. These can be useful if you want to specifiy any additional formatting options, such as the number of decimal places to round to.

#### **Int Formatter**

The `int_formatter` is the basic configuration for a simple interger formatter.

```python
int_formatter(
    datatypes: tuple[object] | object = int,
    format_function: Callable = to_int,
    to_str_function: Callable = to_str,
    invalid_value: object = "NaN",
    **kwargs,
)
```

Parameters:
 - `format_function` (`function`) a function that takes a string and returns an `int`. By default, this is set to the in-built `tksheet.to_int`. This function will always convert float-likes to its floor, for example `"5.9"` will be converted to `5`.
 - `to_str_function` (`function`) By default, this is set to the in-built `tksheet.to_str`, which is a very basic function that will displace the default string representation of the value.

Example:
```python
sheet.format_cell(0, 0, formatter_options = tksheet.int_formatter())
```

#### **Float Formatter**

The `float_formatter` is the basic configuration for a simple float formatter. It will always round float-likes to the specified number of decimal places, for example `"5.999"` will be converted to `"6.0"` if `decimals = 1`.

```python
float_formatter(
    datatypes: tuple[object] | object = float,
    format_function: Callable = to_float,
    to_str_function: Callable = float_to_str,
    invalid_value: object = "NaN",
    decimals: int = 2,
    **kwargs
)
```

Parameters:
 - `format_function` (`function`) a function that takes a string and returns a `float`. By default, this is set to the in-built `tksheet.to_float`. This function will always convert percentages to their decimal equivalent, for example `"5%"` will be converted to `0.05`.
 - `to_str_function` (`function`) By default, this is set to the in-built `tksheet.float_to_str`, which will display the float to the specified number of decimal places.
 - `decimals` (`int`, `None`) the number of decimal places to round to. Defaults to `2`.

Example:
```python
sheet.format_cell(0, 0, formatter_options = tksheet.float_formatter(decimals = None)) # A float formatter with maximum float() decimal places
```

#### **Percentage Formatter**

The `percentage_formatter` is the basic configuration for a simple percentage formatter. It will always round float-likes as a percentage to the specified number of decimal places, for example `"5.999%"` will be converted to `"6.0%"` if `decimals = 1`.

```python
percentage_formatter(
    datatypes: tuple[object] | object = float,
    format_function: Callable = to_float,
    to_str_function: Callable = percentage_to_str,
    invalid_value: object = "NaN",
    decimals: int = 0,
    **kwargs,
)
```

Parameters:
 - `format_function` (`function`) a function that takes a string and returns a `float`. By default, this is set to the in-built `tksheet.to_float`. This function will always convert percentages to their decimal equivalent, for example `"5%"` will be converted to `0.05`.
 - `to_str_function` (`function`) By default, this is set to the in-built `tksheet.percentage_to_str`, which will display the float as a percentage to the specified number of decimal places. For example, `0.05` will be displayed as `"5.0%"`.
 - `decimals` (`int`) the number of decimal places to round to. Defaults to `0`.

Example:
```python
sheet.format_cell(0, 0, formatter_options = tksheet.percentage_formatter(decimals = 1)) # A percentage formatter with 1 decimal place
```

#### **Bool Formatter**

```python
bool_formatter(
    datatypes: tuple[object] | object = bool,
    format_function: Callable = to_bool,
    to_str_function: Callable = bool_to_str,
    invalid_value: object = "NA",
    truthy: set = truthy,
    falsy: set = falsy,
    **kwargs,
)
```

Parameters:
 - `format_function` (`function`) a function that takes a string and returns a `bool`. By default, this is set to the in-built `tksheet.to_bool`.
 - `to_str_function` (`function`) By default, this is set to the in-built `tksheet.bool_to_str`, which will display the boolean as `"True"` or `"False"`.
 - `truthy` (`set`) a set of values that will be converted to `True`. Defaults to the in-built `tksheet.truthy`.
 - `falsy` (`set`) a set of values that will be converted to `False`. Defaults to the in-built `tksheet.falsy`.

Example:
```python
# A bool formatter with custom truthy and falsy values to account for aussie and kiwi slang
sheet.format_cell(0, 0, formatter_options = tksheet.bool_formatter(truthy = tksheet.truthy | {"nah yeah"}, falsy = tksheet.falsy | {"yeah nah"}))
```

### **Datetime Formatters and Designing Your Own Custom Formatters**

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

# From here we can pass our datetime_formatter into sheet.format() or span.format() just like any other formatter
```

For those wanting even more customisation of their formatters you also have the option of creating a custom formatter class. This is a more advanced topic and is not covered here, but it's recommended to create a new class which is a subclass of the `tksheet.Formatter` class and overriding the methods you would like to customise. This custom class can then be passed into the `format_cells()` `formatter_class` argument.

---
# **Readonly Cells**

#### **Creating a readonly rule**

`Span` objects (more information [here](https://github.com/ragardner/tksheet/wiki/Version-7#span-objects)) can be used to create readonly rules for cells, rows, columns, the entire sheet, headers and the index.

You can use either of the following methods:
- Using a span method e.g. `span.readonly()` more information [here](https://github.com/ragardner/tksheet/wiki/Version-7#using-a-span-to-set-cells-to-read-only).
- Using a sheet method e.g. `sheet.readonly(Span)`

Or if you need user inserted row/columns in the middle of areas with a readonly rule to also have a readonly rule you can use named spans, more information [here](https://github.com/ragardner/tksheet/wiki/Version-7#named-spans).

Whether cells, rows or columns are readonly depends on the [`kind`](https://github.com/ragardner/tksheet/wiki/Version-7#get-a-spans-kind) of span.

```python
readonly(
    *key: CreateSpanTypes,
    readonly: bool = True,
)
```

 Parameters:
- `key` (`CreateSpanTypes`) either a span or a type which can create a span. See [here](https://github.com/ragardner/tksheet/wiki/Version-7#creating-a-span) for more information on the types that can create a span.
- `readonly` (`bool`) `True` to create a rule and `False` to delete one created without the use of named spans.

___

#### **Deleting a readonly rule**

If the readonly rule was created by a named span then the named span must be deleted, more information [here](https://github.com/ragardner/tksheet/wiki/Version-7#deleting-a-named-span).

Otherwise you can use either of the following methods to delete/remove readonly rules:
- Using a span method e.g. `span.readonly()` with the keyword argument `readonly=False` more information [here](https://github.com/ragardner/tksheet/wiki/Version-7#using-a-span-to-set-cells-to-read-only).
- Using a sheet method e.g. `sheet.readonly(Span)` with the keyword argument `readonly=False` example below:

```python
# creating a readonly rule
self.sheet.readonly(
    self.sheet.span("A", header=True),
    readonly=True,
)

# deleting the readonly rule
self.sheet.readonly(
    self.sheet.span("A", header=True),
    readonly=False,
)
```

 Parameters:
- `key` (`CreateSpanTypes`) either a span or a type which can create a span. See [here](https://github.com/ragardner/tksheet/wiki/Version-7#creating-a-span) for more information on the types that can create a span.
- `readonly` (`bool`) `True` to create a rule and `False` to delete one created without the use of named spans.

---
# **Text Font and Alignment**

### **Font**

- Font arguments require a three tuple e.g. `("Arial", 12, "normal")` or `("Arial", 12, "bold")` or `("Arial", 12, "italic")`
- The table and index currently share a font, it's not possible to change the index font separate from the table font.

**Set the table and index font:**
```python
font(newfont = None, reset_row_positions = True)
```

**Set the header font:**
```python
header_font(newfont = None)
```

### **Text Alignment**

There are functions to set the text alignment for specific cells/rows/columns and also functions to set the text alignment for a whole part of the sheet (table/index/header).

- Alignment argument (`str`) options are:
    - `"w"`, `"west"`, `"left"`
    - `"e"`, `"east"`, `"right"`
    - `"c"`, `"center"`, `"centre"`

Unfortunately vertical alignment is not available.

#### **Whole widget text alignment**

Set the text alignment for the whole of the table (doesn't include index/header).
```python
table_align(align = None, redraw = True)
```

Set the text alignment for the whole of the header.
```python
header_align(align = None, redraw = True)
```

Set the text alignment for the whole of the index.
```python
row_index_align(align = None, redraw = True)

#same as row_index_align()
index_align(align = None, redraw = True)
```

#### **Creating a specific cell row or column text alignment rule**

The following function is for setting text alignment for specific cells, rows or columns in the table, header and index.

`Span` objects (more information [here](https://github.com/ragardner/tksheet/wiki/Version-7#span-objects)) can be used to create text alignment rules for cells, rows, columns, the entire sheet, headers and the index.

You can use either of the following methods:
- Using a span method e.g. `span.align()` more information [here](https://github.com/ragardner/tksheet/wiki/Version-7#using-a-span-to-create-text-alignment-rules).
- Using a sheet method e.g. `sheet.align(Span)`

Or if you need user inserted row/columns in the middle of areas with an alignment rule to also have an alignment rule you can use named spans, more information [here](https://github.com/ragardner/tksheet/wiki/Version-7#named-spans).

Whether cells, rows or columns are affected depends on the [`kind`](https://github.com/ragardner/tksheet/wiki/Version-7#get-a-spans-kind) of span.

```python
align(
    *key: CreateSpanTypes,
    align: str | None = None,
    redraw: bool = True,
)
```

 Parameters:
- `key` (`CreateSpanTypes`) either a span or a type which can create a span. See [here](https://github.com/ragardner/tksheet/wiki/Version-7#creating-a-span) for more information on the types that can create a span.
- `align` (`str`, `None`) must be one of the following:
    - `"w"`, `"west"`, `"left"`
    - `"e"`, `"east"`, `"right"`
    - `"c"`, `"center"`, `"centre"`

#### **Deleting a specific text alignment rule**

If the text alignment rule was created by a named span then the named span must be deleted, more information [here](https://github.com/ragardner/tksheet/wiki/Version-7#deleting-a-named-span).

Otherwise you can use either of the following methods to delete/remove specific text alignment rules:
- Using a span method e.g. `span.del_align()` more information [here](https://github.com/ragardner/tksheet/wiki/Version-7#using-a-span-to-delete-text-alignment-rules).
- Using a sheet method e.g. `sheet.del_align(Span)` details below:

```python
del_align(
    *key: CreateSpanTypes,
    redraw: bool = True,
)
```

 Parameters:
- `key` (`CreateSpanTypes`) either a span or a type which can create a span. See [here](https://github.com/ragardner/tksheet/wiki/Version-7#creating-a-span) for more information on the types that can create a span.

#### **Get existing specific text alignments**

Cell text alignments:
```python
get_cell_alignments()
```

Row text alignments:
```python
get_row_alignments()
```

Column text alignments:
```python
get_column_alignments()
```

---
# **Getting Selected Cells**

```python
get_currently_selected()
```
- Returns `namedtuple` of `(row, column, type_, tags)` e.g. `(0, 0, "column", (tags))`
   - `type_` can be `"row"`, `"column"` or `"cell"`
   - `tags` resembles the following:
```python
"""
As an example of currently selected tags
"""
(
    "selected",  # name
    "0_0_1_1",  # dimensions of box it's attached to
    250,  # canvas item id of currently selected rectangle
    "0_0",  # coordinates "row_column" of currently selected box
    "type_cells",  # type of box it's attached to, "type_cells", "type_rows" or "type_columns"
)
```

Example:
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

#### **Check if cell is selected**
```python
cell_selected(r, c)
```
- Returns `bool`.

___

#### **Check if row is selected**
```python
row_selected(r)
```
- Returns `bool`.

___

#### **Check if column is selected**
```python
column_selected(c)
```
- Returns `bool`.

___

#### **Check if any cells, rows or columns are selected, there are options for exclusions**
```python
anything_selected(exclude_columns = False, exclude_rows = False, exclude_cells = False)
```
- Returns `bool`.

___

#### **Check if user has the entire table selected**
```python
all_selected()
```
- Returns `bool`.

___

```python
get_ctrl_x_c_boxes()
```

___

```python
get_selected_min_max()
```
- returns `(min_y, min_x, max_y, max_x)` of any selections including rows/columns.

---
# **Modifying Selected Cells**

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

---
# **Row Heights and Column Widths**

#### **Auto resize column widths to fit the window**

```python
set_options(auto_resize_columns)
```
- `auto_resize_columns` (`int`, `None`) if set as an `int` the columns will automatically resize to fit the width of the window, the `int` value being the minimum of each column in pixels.

___

#### **Auto resize row heights to fit the window**
```python
set_options(auto_resize_rows)
```
- `auto_resize_rows` (`int`, `None`) if set as an `int` the rows will automatically resize to fit the height of the window, the `int` value being the minimum height of each row in pixels.

___

#### **Set default column width in pixels**
```python
default_column_width(width = None)
```
- `width` (`int`).

___

#### **Set default row height in pixels or lines**
```python
default_row_height(height = None)
```
- `height` (`int`, `str`) use a numerical `str` for number of lines e.g. `"3"` for a height that fits 3 lines or `int` for pixels.

___

#### **Set default header bar height in pixels or lines**
```python
default_header_height(height = None)
```
- `height` (`int`, `str`) use a numerical `str` for number of lines e.g. `"3"` for a height that fits 3 lines or `int` for pixels.

___

#### **Set a specific cell size to its text**
```python
set_cell_size_to_text(row, column, only_set_if_too_small = False, redraw = True)
```

___

#### **Set all row heights and column widths to cell text sizes**
```python
set_all_cell_sizes_to_text(redraw = True)
```

___

#### **Get the sheets column widths**
```python
get_column_widths(canvas_positions = False)
```
- `canvas_positions` (`bool`) gets the actual canvas x coordinates of column lines.

___

#### **Get the sheets row heights**
```python
get_row_heights(canvas_positions = False)
```
- `canvas_positions` (`bool`) gets the actual canvas y coordinates of row lines.

___

#### **Set all column widths to specific `width` in pixels (`int`) or leave `None` to set to cell text sizes for each column**
```python
set_all_column_widths(width = None, only_set_if_too_small = False, redraw = True, recreate_selection_boxes = True)
```

___

#### **Set all row heights to specific `height` in pixels (`int`) or leave `None` to set to cell text sizes for each row**
```python
set_all_row_heights(height = None, only_set_if_too_small = False, redraw = True, recreate_selection_boxes = True)
```

___

#### **Set/get a specific column width**
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

#### **Set/get a specific row height**
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

---
# **Identifying Bound Event Mouse Position**

The below functions require a mouse click event, for example you could bind right click, example [here](https://github.com/ragardner/tksheet/wiki/Version-7#example-custom-right-click-and-text-editor-functionality), and then identify where the user has clicked.

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

---
# **Modifying and Getting Scroll Positions**

#### **Sync scroll positions between widgets**

```python
sync_scroll(widget: object)
```
- Sync scroll positions between `Sheet`s, may or may not work with other widgets. Uses scrollbar positions.

Syncing two sheets:
```python
self.sheet1.sync_scroll(self.sheet2)
```

Syncing three sheets:
```python
# syncs sheet 1 and 2 between each other
self.sheet1.sync_scroll(self.sheet2)

# syncs sheet 1 and 3 between each other
self.sheet1.sync_scroll(self.sheet3)

# syncs sheet 2 and 3 between each other
self.sheet2.sync_scroll(self.sheet3)
```

#### **Unsync scroll positions between widgets**

```python
unsync_scroll(widget: None | Sheet = None)
```
- Leaving `widget` as `None` unsyncs all previously synced widgets.

#### **See / scroll to a specific cell on the sheet**
```python
see(
    row: int = 0,
    column: int = 0,
    keep_yscroll: bool = False,
    keep_xscroll: bool = False,
    bottom_right_corner: bool = False,
    check_cell_visibility: bool = True,
    redraw: bool = True,
) -> Sheet
```

___

#### **Check if a cell has any part of it visible**
```python
cell_visible(r, c)
```
- Returns `bool`.

___

#### **Check if a cell is totally visible**
```python
cell_completely_visible(r, c, seperate_axes = False)
```
- Returns `bool`.
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

---
# **Hiding Columns**

#### **Display only certain columns**
```python
display_columns(columns = None,
                all_columns_displayed = None,
                reset_col_positions = True,
                refresh = False,
                redraw = False,
                deselect_all = True,
                **kwargs)
```
Parameters:
- `columns` (`int`, `iterable`, `"all"`) are the columns to be displayed, omit the columns to be hidden.
- Use argument `True` with `all_columns_displayed` to display all columns, use `False` to display only the columns you've set using the `columns` arg.
- You can also use the keyword argument `all_displayed` instead of `all_columns_displayed`.

Examples:
```python
# display all columns
self.sheet.display_columns("all")

# displaying specific columns only
self.sheet.display_columns([2, 4, 7], all_displayed = False)
```

___

#### **Get all columns displayed boolean**
```python
all_columns_displayed(a = None)
```
- `a` (`bool`, `None`) Either set by using `bool` or get by leaving `None` e.g. `all_columns_displayed()`.

___

#### **Hide specific columns**
```python
hide_columns(columns = set(),
             redraw = True,
             deselect_all = True)
```
- **NOTE**: `columns` (`int`) uses displayed column indexes, not data indexes. In other words the indexes of the columns displayed on the screen are the ones that are hidden, this is useful when used in conjunction with `get_selected_columns()`.

---
# **Hiding Rows**

#### **Display only certain rows**
```python
display_rows(rows = None,
             all_rows_displayed = None,
             reset_col_positions = True,
             refresh = False,
             redraw = False,
             deselect_all = True,
             **kwargs)
```
Parameters:
- `rows` (`int`, `iterable`, `"all"`) are the rows to be displayed, omit the rows to be hidden.
- Use argument `True` with `all_rows_displayed` to display all rows, use `False` to display only the rows you've set using the `rows` arg.
- You can also use the keyword argument `all_displayed` instead of `all_rows_displayed`.

Examples:
- An example of row filtering using this function can be found [here](https://github.com/ragardner/tksheet/wiki/Version-7#example-header-dropdown-boxes-and-row-filtering).
- More examples below:
```python
# display all rows
self.sheet.display_rows("all")

# display specific rows only
self.sheet.display_rows([2, 4, 7], all_displayed = False)
```

___

#### **Get all rows displayed boolean**
```python
all_rows_displayed(a = None)
```
- `a` (`bool`, `None`) Either set by using `bool` or get by leaving `None` e.g. `all_rows_displayed()`.

___

#### **Hide specific rows**
```python
hide_rows(rows = set(),
          redraw = True,
          deselect_all = True)
```
- **NOTE**: `rows` (`int`) uses displayed row indexes, not data indexes. In other words the indexes of the rows displayed on the screen are the ones that are hidden, this is useful when used in conjunction with `get_selected_rows()`.

---
# **Hiding Table Elements**

#### **Hide parts of the table or all of it**
```python
hide(canvas = "all")
```
- `canvas` (`str`) options are `all`, `row_index`, `header`, `top_left`, `x_scrollbar`, `y_scrollbar`
	- `all` hides the entire table and is the default.

___

#### **Show parts of the table or all of it**
```python
show(canvas = "all")
```
- `canvas` (`str`) options are `all`, `row_index`, `header`, `top_left`, `x_scrollbar`, `y_scrollbar`
	- `all` shows the entire table and is the default.

---
# **Cell Text Editor**

#### **Open the currently selected cell in the main table**
```python
open_cell(ignore_existing_editor = True)
```
- Function utilises the currently selected cell in the main table, even if a column/row is selected, to open a non selected cell first use `set_currently_selected()` to set the cell to open.

___

#### **Open the currently selected cell but in the header**
```python
open_header_cell(ignore_existing_editor = True)
```
- Also uses currently selected cell, which you can set with `set_currently_selected()`.

___

#### **Open the currently selected cell but in the index**
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

---
# **Table Options and Other Functions**

```python
set_options(redraw: bool = True, **kwargs)
```

The list of key word arguments available for `set_options()` are as follows, [see here](https://github.com/ragardner/tksheet/wiki/Version-7#initialization-options) as a guide for what arguments to use.
```python
auto_resize_columns
auto_resize_rows
to_clipboard_delimiter
to_clipboard_quotechar
to_clipboard_lineterminator
from_clipboard_delimiters
show_dropdown_borders
edit_cell_validation
show_default_header_for_empty
show_default_index_for_empty
selected_rows_to_end_of_window
horizontal_grid_to_end_of_window
vertical_grid_to_end_of_window
paste_insert_column_limit
paste_insert_row_limit
expand_sheet_if_paste_too_big
arrow_key_down_right_scroll_page
enable_edit_cell_auto_resize
page_up_down_select_row
display_selected_fg_over_highlights
show_horizontal_grid
show_vertical_grid
empty_horizontal
empty_vertical
row_height
column_width
header_height
row_drag_and_drop_perform
column_drag_and_drop_perform
auto_resize_default_row_index
default_header
default_row_index
max_column_width
max_row_height
max_header_height
max_index_width
font
header_font
index_font

show_selected_cells_border
theme
top_left_bg
top_left_fg
top_left_fg_highlight

table_bg
table_grid_fg
table_fg
table_selected_box_cells_fg
table_selected_box_rows_fg
table_selected_box_columns_fg
table_selected_cells_border_fg
table_selected_cells_bg
table_selected_cells_fg
table_selected_rows_border_fg
table_selected_rows_bg
table_selected_rows_fg
table_selected_columns_border_fg
table_selected_columns_bg
table_selected_columns_fg

header_bg
header_border_fg
header_grid_fg
header_fg
header_selected_cells_bg
header_selected_cells_fg
header_selected_columns_bg
header_selected_columns_fg

index_bg
index_border_fg
index_grid_fg
index_fg
index_selected_cells_bg
index_selected_cells_fg
index_selected_rows_bg
index_selected_rows_fg

resizing_line_fg
drag_and_drop_bg
outline_thickness
outline_color
frame_bg
popup_menu_font
popup_menu_fg
popup_menu_bg
popup_menu_highlight_bg
popup_menu_highlight_fg
```

___

Get internal storage dictionary of highlights, readonly cells, dropdowns etc.
```python
get_cell_options(canvas = "table")
```

___

Delete any formats, alignments, dropdown boxes, check boxes, highlights etc. that are larger than the sheets currently held data, includes row index and header in measurement of dimensions.
```python
delete_out_of_bounds_options()
```

___

Delete all alignments, dropdown boxes, check boxes, highlights etc.
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

Various functions related to the Sheets internal undo and redo stacks.
```python
# clears both undos and redos
reset_undos()

# get the Sheets modifiable deque variables which store changes for undo and redo
get_undo_stack()
get_redo_stack()

# set the Sheets undo and redo stacks, returns Sheet widget
set_undo_stack(stack: deque)
set_redo_stack(stack: deque)
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

---
# **Example Loading Data from Excel**

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

---
# **Example Custom Right Click and Text Editor Validation**

This is to demonstrate:
- Adding your own commands to the in-built right click popup menu (or how you might start making your own right click menu functionality)
- Validating text editor input; in this demonstration the validation removes spaces from user input.

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
        return event.value

    def end_edit_cell(self, event = None):
        # remove spaces from user input
        if event.value is not None:
            return event.value.replace(" ", "")


app = demo()
app.mainloop()
```
- If you want a totally new right click menu you can use `self.sheet.bind("<3>", <function>)` with a `tk.Menu` of your own design (right click is `<2>` on MacOS) and don't use `"right_click_popup_menu"` with `enable_bindings()`.

---
# **Example Displaying Selections**

```python
from tksheet import (
    Sheet,
    num2alpha,
)
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
        self.sheet.enable_bindings("all", "ctrl_select")
        self.sheet.extra_bindings([("all_select_events", self.sheet_select_event)])
        self.show_selections = tk.Label(self)
        self.frame.grid(row = 0, column = 0, sticky = "nswe")
        self.sheet.grid(row = 0, column = 0, sticky = "nswe")
        self.show_selections.grid(row = 1, column = 0, sticky = "nsw")

    def sheet_select_event(self, event = None):
        if event.eventname == "select" and event.selection_boxes and event.selected:
            # get the most recently selected box in case there are multiple
            box = next(reversed(event.selection_boxes))
            type_ = event.selection_boxes[box]
            if type_ == "cells":
                self.show_selections.config(text=f"{type_.capitalize()}: {box.from_r + 1},{box.from_c + 1} : {box.upto_r},{box.upto_c}")
            elif type_ == "rows":
                self.show_selections.config(text=f"{type_.capitalize()}: {box.from_r + 1} : {box.upto_r}")
            elif type_ == "columns":
                self.show_selections.config(text=f"{type_.capitalize()}: {num2alpha(box.from_c)} : {num2alpha(box.upto_c - 1)}")
        else:
            self.show_selections.config(text="")


app = demo()
app.mainloop()
```

---
# **Example List Box**

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

---
# **Example Header Dropdown Boxes and Row Filtering**

A very simple demonstration of row filtering using header dropdown boxes.

```python
from tksheet import (
    Sheet,
    num2alpha as n2a,
)
import tkinter as tk


class demo(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.frame = tk.Frame(self)
        self.frame.grid_columnconfigure(0, weight=1)
        self.frame.grid_rowconfigure(0, weight=1)
        self.data = [
            ["3", "c", "z"],
            ["1", "a", "x"],
            ["1", "b", "y"],
            ["2", "b", "y"],
            ["2", "c", "z"],
        ]
        self.sheet = Sheet(
            self.frame,
            data=self.data,
            column_width=180,
            theme="dark",
            height=700,
            width=1100,
        )
        self.sheet.enable_bindings(
            "copy",
            "rc_select",
            "arrowkeys",
            "double_click_column_resize",
            "column_width_resize",
            "column_select",
            "row_select",
            "drag_select",
            "single_select",
            "select_all",
        )
        self.frame.grid(row=0, column=0, sticky="nswe")
        self.sheet.grid(row=0, column=0, sticky="nswe")

        self.sheet.dropdown(
            self.sheet.span(n2a(0), header=True, table=False),
            values=["all", "1", "2", "3"],
            set_value="all",
            selection_function=self.header_dropdown_selected,
            text="Header A Name",
        )
        self.sheet.dropdown(
            self.sheet.span(n2a(1), header=True, table=False),
            values=["all", "a", "b", "c"],
            set_value="all",
            selection_function=self.header_dropdown_selected,
            text="Header B Name",
        )
        self.sheet.dropdown(
            self.sheet.span(n2a(2), header=True, table=False),
            values=["all", "x", "y", "z"],
            set_value="all",
            selection_function=self.header_dropdown_selected,
            text="Header C Name",
        )

    def header_dropdown_selected(self, event=None):
        hdrs = self.sheet.headers()
        # this function is run before header cell data is set by dropdown selection
        # so we have to get the new value from the event
        hdrs[event.loc] = event.value
        if all(dd == "all" for dd in hdrs):
            self.sheet.display_rows("all")
        else:
            rows = [
                rn for rn, row in enumerate(self.data) if all(row[c] == e or e == "all" for c, e in enumerate(hdrs))
            ]
            self.sheet.display_rows(rows=rows, all_displayed=False)
        self.sheet.redraw()


app = demo()
app.mainloop()
```

---
# **Example Readme Screenshot Code**

The code used to make a screenshot for the readme file.

```python
from tksheet import (
    Sheet,
    num2alpha as n2a,
)
import tkinter as tk


class demo(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.frame = tk.Frame(self)
        self.frame.grid_columnconfigure(0, weight=1)
        self.frame.grid_rowconfigure(0, weight=1)
        self.sheet = Sheet(
            self.frame,
            expand_sheet_if_paste_too_big=True,
            empty_horizontal=0,
            empty_vertical=0,
            align="w",
            header_align="c",
            data=[[f"Row {r}, Column {c}\nnewline 1\nnewline 2" for c in range(6)] for r in range(21)],
            headers=[
                "Dropdown Column",
                "Checkbox Column",
                "Center Aligned Column",
                "East Aligned Column",
                "",
                "",
            ],
            theme="black",
            height=520,
            width=930,
        )
        self.sheet.enable_bindings("all", "edit_index", "edit header")
        self.sheet.popup_menu_add_command(
            "Hide Rows",
            self.hide_rows,
            table_menu=False,
            header_menu=False,
            empty_space_menu=False,
        )
        self.sheet.popup_menu_add_command(
            "Show All Rows",
            self.show_rows,
            table_menu=False,
            header_menu=False,
            empty_space_menu=False,
        )
        self.sheet.popup_menu_add_command(
            "Hide Columns",
            self.hide_columns,
            table_menu=False,
            index_menu=False,
            empty_space_menu=False,
        )
        self.sheet.popup_menu_add_command(
            "Show All Columns",
            self.show_columns,
            table_menu=False,
            index_menu=False,
            empty_space_menu=False,
        )
        self.frame.grid(row=0, column=0, sticky="nswe")
        self.sheet.grid(row=0, column=0, sticky="nswe")
        colors = (
            "#509f56",
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
            "#f85037",
        )
        # self.sheet.align_columns(columns=2, align="c")
        # self.sheet.align_columns(columns=3, align="e")
        # self.sheet.create_dropdown(r="all", c=0, values=["Dropdown"] + [f"{i}" for i in range(15)])
        # self.sheet.create_checkbox(r="all", c=1, checked=True, text="Checkbox")
        # self.sheet.create_header_dropdown(c=0, values=["Header Dropdown"] + [f"{i}" for i in range(15)])
        # self.sheet.create_header_checkbox(c=1, checked=True, text="Header Checkbox")
        # self.sheet.align_cells(5, 0, align="c")
        # self.sheet.highlight_cells(5, 0, bg="gray50", fg="blue")
        # self.sheet.highlight_cells(17, canvas="index", bg="yellow", fg="black")
        # self.sheet.highlight_cells(12, 1, bg="gray90", fg="purple")
        # for r in range(len(colors)):
        #     self.sheet.highlight_cells(row=r, column=3, fg=colors[r])
        #     self.sheet.highlight_cells(row=r, column=4, bg=colors[r], fg="black")
        #     self.sheet.highlight_cells(row=r, column=5, bg=colors[r], fg="purple")
        # self.sheet.highlight_cells(column=5, canvas="header", bg="white", fg="purple")
        self.sheet.align(n2a(2), align="c")
        self.sheet.align(n2a(3), align="e")
        self.sheet.dropdown(
            self.sheet.span("A", header=True),
            values=["Dropdown"] + [f"{i}" for i in range(15)],
        )
        self.sheet.checkbox(
            self.sheet.span("B", header=True),
            checked=True,
            text="Checkbox",
        )
        self.sheet.align(5, 0, align="c")
        self.sheet.highlight(5, 0, bg="gray50", fg="blue")
        self.sheet.highlight(
            self.sheet.span(17, index=True, table=False),
            bg="yellow",
            fg="black",
        )
        self.sheet.highlight(12, 1, bg="gray90", fg="purple")
        for r in range(len(colors)):
            self.sheet.highlight(r, 3, fg=colors[r])
            self.sheet.highlight(r, 4, bg=colors[r], fg="black")
            self.sheet.highlight(r, 5, bg=colors[r], fg="purple")
        self.sheet.highlight(
            self.sheet.span(n2a(5), header=True, table=False),
            bg="white",
            fg="purple",
        )
        self.sheet.set_all_column_widths()
        self.sheet.extra_bindings("all", self.all_extra_bindings)

    def hide_rows(self, event=None):
        rows = self.sheet.get_selected_rows()
        if rows:
            self.sheet.hide_rows(rows)

    def show_rows(self, event=None):
        self.sheet.display_rows("all", redraw=True)

    def hide_columns(self, event=None):
        columns = self.sheet.get_selected_columns()
        if columns:
            self.sheet.hide_columns(columns)

    def show_columns(self, event=None):
        self.sheet.display_columns("all", redraw=True)

    def all_extra_bindings(self, event=None):
        return event.value


app = demo()
app.mainloop()
```

---
# **Example Saving tksheet as a csv File**

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
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.frame = tk.Frame(self)
        self.frame.grid_columnconfigure(0, weight=1)
        self.frame.grid_rowconfigure(0, weight=1)
        self.sheet = Sheet(self.frame, data=[[f"Row {r}, Column {c}" for c in range(6)] for r in range(21)])
        self.sheet.enable_bindings("all", "edit_header", "edit_index")
        self.frame.grid(row=0, column=0, sticky="nswe")
        self.sheet.grid(row=0, column=0, sticky="nswe")
        self.sheet.popup_menu_add_command("Open csv", self.open_csv)
        self.sheet.popup_menu_add_command("Save sheet", self.save_sheet)
        self.sheet.set_all_cell_sizes_to_text()
        self.sheet.change_theme("light green")

        # create a span which encompasses the table, header and index
        # all data values, no displayed values
        self.sheet_span = self.sheet.span(
            header=True,
            index=True,
            hdisp=False,
            idisp=False,
        )

        # center the window and unhide
        self.update_idletasks()
        w = self.winfo_screenwidth() - 20
        h = self.winfo_screenheight() - 70
        size = (900, 500)
        self.geometry("%dx%d+%d+%d" % (size + ((w / 2 - size[0] / 2), h / 2 - size[1] / 2)))
        self.deiconify()

    def save_sheet(self):
        filepath = filedialog.asksaveasfilename(
            parent=self,
            title="Save sheet as",
            filetypes=[("CSV File", ".csv"), ("TSV File", ".tsv")],
            defaultextension=".csv",
            confirmoverwrite=True,
        )
        if not filepath or not filepath.lower().endswith((".csv", ".tsv")):
            return
        try:
            with open(normpath(filepath), "w", newline="", encoding="utf-8") as fh:
                writer = csv.writer(
                    fh,
                    dialect=csv.excel if filepath.lower().endswith(".csv") else csv.excel_tab,
                    lineterminator="\n",
                )
                writer.writerows(self.sheet_span.data)
        except Exception as error:
            print(error)
            return

    def open_csv(self):
        filepath = filedialog.askopenfilename(parent=self, title="Select a csv file")
        if not filepath or not filepath.lower().endswith((".csv", ".tsv")):
            return
        try:
            with open(normpath(filepath), "r") as filehandle:
                filedata = filehandle.read()
            self.sheet_span.data = [
                r
                for r in csv.reader(
                    io.StringIO(filedata),
                    dialect=csv.Sniffer().sniff(filedata),
                    skipinitialspace=False,
                )
            ]
        except Exception as error:
            print(error)
            return


app = demo()
app.mainloop()

```

---
# **Example Using and Creating Formatters**

```python
from tksheet import *
import tkinter as tk
from datetime import datetime, date, timedelta, time
from dateutil import parser, tz
from math import ceil
import re

date_replace = re.compile('|'.join(['\(', '\)', '\[', '\]', '\<', '\>']))

# Custom formatter methods
def round_up(x):
    try: # might not be a number if empty
        return float(ceil(x))
    except:
        return x

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

# Custom Formatter with additional kwargs

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

        # Some examples of data formatting
        self.sheet.format_cell('all', 0, formatter_options = float_formatter(nullable = False))
        self.sheet.format_cell('all', 1, formatter_options = float_formatter())
        self.sheet.format_cell('all', 2, formatter_options = int_formatter())
        self.sheet.format_cell('all', 3, formatter_options = bool_formatter(truthy = truthy | {"nah yeah"}, falsy = falsy | {"yeah nah"}))
        self.sheet.format_cell('all', 4, formatter_options = percentage_formatter())


        # Custom Formatters
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

---
# **Contributing**

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
- You're also welcome to become a contributor yourself and help implement your idea!

