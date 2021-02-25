# Table of Contents
1. [About tksheet](https://github.com/ragardner/tksheet/blob/master/DOCUMENTATION.md#About-tksheet)
2. [Installation and Requirements](https://github.com/ragardner/tksheet/blob/master/DOCUMENTATION.md#Installation-and-Requirements)
3. [Basic Initialization](https://github.com/ragardner/tksheet/blob/master/DOCUMENTATION.md#Basic-Initialization)
4. [Initialization Options](https://github.com/ragardner/tksheet/blob/master/DOCUMENTATION.md#Initialization-Options)
5. [Modifying Table Data](https://github.com/ragardner/tksheet/blob/master/DOCUMENTATION.md#Modifying-Table-Data)
6. [Retrieving Table Data](https://github.com/ragardner/tksheet/blob/master/DOCUMENTATION.md#Retrieving-Table-Data)
7. [Bindings and Functionality](https://github.com/ragardner/tksheet/blob/master/DOCUMENTATION.md#Bindings-and-Functionality)
8. [Identifying Bound Event Mouse Position](https://github.com/ragardner/tksheet/blob/master/DOCUMENTATION.md#Identifying-Bound-Event-Mouse-Position)
9. [Table Colors](https://github.com/ragardner/tksheet/blob/master/DOCUMENTATION.md#Table-Colors)
10. [Highlighting Cells](https://github.com/ragardner/tksheet/blob/master/DOCUMENTATION.md#Highlighting-Cells)
11. [Text Font and Alignment](https://github.com/ragardner/tksheet/blob/master/DOCUMENTATION.md#Text-Font-and-Alignment)
12. [Row Heights and Column Widths](https://github.com/ragardner/tksheet/blob/master/DOCUMENTATION.md#Row-Heights-and-Column-Widths)
13. [Retrieving Selected Cells](https://github.com/ragardner/tksheet/blob/master/DOCUMENTATION.md#Retrieving-Selected-Cells)
14. [Modifying Selected Cells](https://github.com/ragardner/tksheet/blob/master/DOCUMENTATION.md#Modifying-Selected-Cells)
15. [Modifying and Retrieving Scroll Positions](https://github.com/ragardner/tksheet/blob/master/DOCUMENTATION.md#Modifying-and-Retrieving-Scroll-Positions)
16. [Setting Readonly Cells](https://github.com/ragardner/tksheet/blob/master/DOCUMENTATION.md#Setting-Readonly-Cells)
17. [Hiding Columns](https://github.com/ragardner/tksheet/blob/master/DOCUMENTATION.md#Hiding-Columns)
18. [Hiding the Index and Header](https://github.com/ragardner/tksheet/blob/master/DOCUMENTATION.md#Hiding-the-Index-and-Header)
19. [Cell Text Editor](https://github.com/ragardner/tksheet/blob/master/DOCUMENTATION.md#Cell-Text-Editor)
20. [Dropdown Boxes](https://github.com/ragardner/tksheet/blob/master/DOCUMENTATION.md#Dropdown-Boxes)
21. [Example: Loading Data from Excel](https://github.com/ragardner/tksheet/blob/master/DOCUMENTATION.md#Example:-Loading-Data-from-Excel)


## About tksheet
`tksheet` is a Python tkinter table widget written in pure python. It is licensed under the [MIT license](https://github.com/ragardner/tksheet/blob/master/LICENSE.txt).

## Installation and Requirements
`tksheet` is available through PyPi (Python package index) and can be installed by using Pip through the command line `pip install tksheet`

Alternatively you can download the source code and (inside the tksheet directory) use the command line `python setup.py develop`

`tksheet` requires a Python version of `3.6` or higher.

## Basic Initialization
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
                           data = [[f"Row {r}, Column {c}\nnewline1\nnewline2" for c in range(5)] for r in range(5)])
        self.sheet.enable_bindings()
        self.frame.grid(row = 0, column = 0, sticky = "nswe")
        self.sheet.grid(row = 0, column = 0, sticky = "nswe")


app = demo()
app.mainloop()
```

## Initialization Options
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
default_header = "letters",     #letters, numbers or both
default_row_index = "numbers",  #letters, numbers or both
page_up_down_select_row = True,
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
show_vertical_grid = True,
show_horizontal_grid = True,
display_selected_fg_over_highlights = False,
show_selected_cells_border = True,
theme = "light blue",
popup_menu_fg                           = "gray2",
popup_menu_bg                           = "#f2f2f2",
popup_menu_highlight_bg                 = "#91c9f7",
popup_menu_highlight_fg                 = "black",
frame_bg                                = theme_light_blue['table_bg'],
table_grid_fg                           = theme_light_blue['table_grid_fg'],
table_bg                                = theme_light_blue['table_bg'],
table_fg                                = theme_light_blue['table_fg'], 
table_selected_cells_border_fg          = theme_light_blue['table_selected_cells_border_fg'],
table_selected_cells_bg                 = theme_light_blue['table_selected_cells_bg'],
table_selected_cells_fg                 = theme_light_blue['table_selected_cells_fg'],
table_selected_rows_border_fg           = theme_light_blue['table_selected_rows_border_fg'],
table_selected_rows_bg                  = theme_light_blue['table_selected_rows_bg'],
table_selected_rows_fg                  = theme_light_blue['table_selected_rows_fg'],
table_selected_columns_border_fg        = theme_light_blue['table_selected_columns_border_fg'],
table_selected_columns_bg               = theme_light_blue['table_selected_columns_bg'],
table_selected_columns_fg               = theme_light_blue['table_selected_columns_fg'],
resizing_line_fg                        = theme_light_blue['resizing_line_fg'],
drag_and_drop_bg                        = theme_light_blue['drag_and_drop_bg'],
index_bg                                = theme_light_blue['index_bg'],
index_border_fg                         = theme_light_blue['index_border_fg'],
index_grid_fg                           = theme_light_blue['index_grid_fg'],
index_fg                                = theme_light_blue['index_fg'],
index_selected_cells_bg                 = theme_light_blue['index_selected_cells_bg'],
index_selected_cells_fg                 = theme_light_blue['index_selected_cells_fg'],
index_selected_rows_bg                  = theme_light_blue['index_selected_rows_bg'],
index_selected_rows_fg                  = theme_light_blue['index_selected_rows_fg'],
header_bg                               = theme_light_blue['header_bg'],
header_border_fg                        = theme_light_blue['header_border_fg'],
header_grid_fg                          = theme_light_blue['header_grid_fg'],
header_fg                               = theme_light_blue['header_fg'],
header_selected_cells_bg                = theme_light_blue['header_selected_cells_bg'],
header_selected_cells_fg                = theme_light_blue['header_selected_cells_fg'],
header_selected_columns_bg              = theme_light_blue['header_selected_columns_bg'],
header_selected_columns_fg              = theme_light_blue['header_selected_columns_fg'],
top_left_bg                             = theme_light_blue['top_left_bg'],
top_left_fg                             = theme_light_blue['top_left_fg'],
top_left_fg_highlight                   = theme_light_blue['top_left_fg_highlight']
)
```

 - `startup_select` selects cells, rows or columns at initialization by using a `tuple` e.g. `(0, 0, "cells")` for cell A0 or `(0, 5, "rows")` for rows 0 to 5.
 - `data_reference` and `data` are essentially the same

You can change these settings after initialization using the `set_options()` function.

## Modifying Table Data

```python
delete_row(idx = 0, deselect_all = False)
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
delete_column(idx = 0, deselect_all = False)
```

___

```python
move_column(column, moveto)
```

## Retrieving Table Data


## Bindings and Functionality

```python
enable_bindings(bindings = "all")
```

___

```python
disable_bindings(bindings = "all")
```

___

```python
basic_bindings(enable = False)
```

___

```python
edit_bindings(enable = False)
```

___

```python
cell_edit_binding(enable = False)
```

___

```python
extra_bindings(bindings, func = "None")
```

___

```python
bind(binding, func)
```

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


## Identifying Bound Event Mouse Position

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



## Table Colors


## Highlighting Cells


## Text Font and Alignment

```python
font(self, newfont = None, reset_row_positions = True)
```

___

```python
header_font(self, newfont = None)
```

___

```python
align(self, align = None, redraw = True)
```

___

```python
header_align(self, align = None, redraw = True)
```

___

```python
row_index_align(self, align = None, redraw = True)
```

___

Change the text alignment for **specific** rows.
```python
align_rows(self, rows = [], align = "global", align_index = False, redraw = True)
```

___

Change the text alignment for **specific** columns.
```python
align_columns(self, columns = [], align = "global", align_header = False, redraw = True)
```

___

Change the text alignment for **specific** cells inside the table.
```python
align_cells(self, row = 0, column = 0, cells = [], align = "global", redraw = True)
```

___

Change the text alignment for **specific** cells inside the header.
```python
align_header(self, columns = [], align = "global", redraw = True)
```

___

Change the text alignment for **specific** cells inside the index.
```python
align_index(self, rows = [], align = "global", redraw = True)
```

## Row Heights and Column Widths

```python
get_example_canvas_column_widths(total_cols = None)
```

___

```python
get_example_canvas_row_heights(total_rows = None)
```

___

```python
get_column_widths(canvas_positions = False)
```

___

```python
get_row_heights(canvas_positions = False)
```

___

```python
set_all_cell_sizes_to_text(redraw = True)
```

___

```python
set_all_column_widths(width = None, only_set_if_too_small = False, redraw = True, recreate_selection_boxes = True)
```

___

```python
set_all_row_heights(height = None, only_set_if_too_small = False, redraw = True, recreate_selection_boxes = True)
```

___

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
verify_row_heights(row_heights, canvas_positions = False)
```

___

```python
verify_column_widths(column_widths, canvas_positions = False)
```

___

```python
default_row_height(height = None)
```

___

```python
default_header_height(height = None)
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

## Retrieving Selected Cells

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

```python
cell_selected(r, c)
```

___

```python
row_selected(r)
```

___

```python
column_selected(c)
```

___

```python
anything_selected(exclude_columns = False, exclude_rows = False, exclude_cells = False)
```

___

```python
all_selected()
```


## Modifying Selected Cells

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

## Modifying and Retrieving Scroll Positions

```python
get_xview()
```

___

```python
get_yview()
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
set_view(x_args, y_args)
```

___

```python
see(row = 0, column = 0, keep_yscroll = False, keep_xscroll = False, bottom_right_corner = False, check_cell_visibility = True)
```

___

```python
move_down()
```

## Setting Readonly Cells


## Hiding Columns


## Table Elements, Height and Width

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

___

Modify widget height and width in pixels
```python
height_and_width(height = None, width = None)
```
 - `height` (`int`) set a height in pixels
 - `width` (`int`) set a width in pixels
If both arguments are `None` then table will reset to default tkinter canvas dimensions.


## Cell Text Editor

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

## Dropdown Boxes

```python
create_dropdown(r = 0,
                c = 0,
                values = [],
                set_value = None,
                state = "readonly",
                see = True,
                destroy_on_leave = False,
                destroy_on_select = True,
                current = False,
                set_cell_on_select = True,
                redraw = True,
                recreate_selection_boxes = True)
```

___

```python
get_dropdown_value(current = False, destroy = True, set_cell_on_select = True, redraw = True, recreate = True)
```

___

```python
delete_dropdown(r = 0, c = 0)
```

___

```python
get_dropdowns()
```

___

```python
refresh_dropdowns(dropdowns = [])
```

___

```python
set_all_dropdown_values_to_sheet()
```


## Example: Loading Data from Excel






















