### Version 4.8.8
 - Fix bug with `highlight` functions where `fg` is set but not `bg`

### Version 4.8.7
 - Fix delete key bug

### Version 4.8.6
 - Change many values sent to functions set by `extra_bindings()` - begin is before action is taken, end is after

| Binding                        | Response                                                                |
|--------------------------------|------------------------------------------------------------------------:|
|"begin_ctrl_c"                  |("begin_ctrl_c", boxes, currently_selected)                              |
|"end_ctrl_c"                    |("end_ctrl_c", boxes, currently_selected, rows)                          |
|"begin_ctrl_x"                  |("begin_ctrl_x", boxes, currently_selected)                              |
|"end_ctrl_x"                    |("end_ctrl_x", boxes, currently_selected, rows)                          |
|"begin_ctrl_v"                  |("begin_ctrl_v", currently_selected, rows)                               |
|"end_ctrl_v"                    |("end_ctrl_v", currently_selected, rows)                                 |
|"begin_delete_key"              |("begin_delete_key", boxes, currently_selected)                          |
|"end_delete_key"                |("end_delete_key", boxes, currently_selected)                            |
|"begin_ctrl_z"                  |("begin_ctrl_z", change type)                                            |
|"end_ctrl_z"                    |("end_ctrl_z", change type)                                              |
|"begin_insert_columns"          |("begin_insert_columns", stidx, posidx, numcols)                         |
|"end_insert_columns"            |("end_insert_columns", stidx, posidx, numcols)                           |
|"begin_insert_rows"             |("begin_insert_rows", stidx, posidx, numrows)                            |
|"end_insert_rows"               |("end_insert_rows", stidx, posidx, numrows)                              |
|"begin_row_index_drag_drop"     |("begin_row_index_drag_drop", orig_selected_rows, r)                     |
|"end_row_index_drag_drop"       |("end_row_index_drag_drop", orig_selected_rows, new_selected, r)         |
|"begin_column_header_drag_drop" |("begin_column_header_drag_drop", orig_selected_cols, c)                 |
|"end_column_header_drag_drop"   |("end_column_header_drag_drop", orig_selected_cols, new_selected, c)   |

`boxes` are all selection box coordinates in the table, currently selected is cell coordinates

`rows` (in ctrl_c, ctrl_x and ctrl_v) is a list of lists which represents the data which was worked on

 - Right click insert columns/rows now inserts as many as selected or one if none selected
 - Some minor improvements

### Version 4.8.5
 - Fix line creation when click drag resizing index
 - Improve looks of top left rectangle resizers when one is disabled

### Version 4.8.4
 - Fix bug in `insert_row_positions()`

### Version 4.8.3
 - Make `enable_bindings()` and `disable_bindings()` without arguments do all bindings
 - Fix top left resizers not showing up if enabling/disabling all bindings
 - Improve look and feel of drag and drop
 - Fix bug in certain circumstances with `insert_row_positions()`
 - Fix `highlight_rows()` bug

### Version 4.8.2
 - Fix old color name, credit `ministiy`

### Version 4.8.1
 - Make drag selection only redraw table when rectangle dimensions change instead of mouse moves

### Version 4.8.0
 - Adjust light green theme colors slightly
 - Changed all color option argument names for better clarity

| Old Name                     | New Name                          |
|------------------------------|----------------------------------:|
|frame_background              |frame_bg                           |
|grid_color                    |table_grid_fg                      |
|table_background              |table_bg                           |
|text_color                    |table_fg                           |
|selected_cells_border_color   |table_selected_cells_border_fg     |
|selected_cells_background     |table_selected_cells_bg            |
|selected_cells_foreground     |table_selected_cells_fg            |
|selected_rows_border_color    |table_selected_rows_border_fg      |
|selected_rows_background      |table_selected_rows_bg             |
|selected_rows_foreground      |table_selected_rows_fg             |
|selected_columns_border_color |table_selected_columns_border_fg   |
|selected_columns_background   |table_selected_columns_bg          |
|selected_columns_foreground   |table_selected_columns_fg          |
|resizing_line_color           |resizing_line_fg                   |
|drag_and_drop_color           |drag_and_drop_bg                   |
|row_index_background          |index_bg                           |
|row_index_foreground          |index_fg                           |
|row_index_border_color        |index_border_fg                    |
|row_index_grid_color          |index_grid_fg                      |
|row_index_select_background   |index_selected_cells_bg            |
|row_index_select_foreground   |index_selected_cells_fg            |
|row_index_select_row_bg       |index_selected_rows_bg             |
|row_index_select_row_fg       |index_selected_rows_fg             |
|header_background             |header_bg                          |
|header_foreground             |header_fg                          |
|header_border_color           |header_border_fg                   |
|header_grid_color             |header_grid_fg                     |
|header_select_background      |header_selected_cells_bg           |
|header_select_foreground      |header_selected_cells_fg           |
|header_select_column_bg       |header_selected_columns_bg         |
|header_select_column_fg       |header_selected_columns_fg         |
|top_left_background           |top_left_bg                        |
|top_left_foreground           |top_left_fg                        |
|top_left_foreground_highlight |top_left_fg_highlight              |

### Version 4.7.9
 - Add option to use integer in `insert_row_positions()` and `insert_column_positions()` for how many columns to add
 - Fix error occurring in above functions when using python 3.6-3.7 due to itertools argument added in python 3.8

### Version 4.7.8
 - Fix show_horizontal and show_vertical grid

### Version 4.7.7
 - Reduce column minimum size even further
 - Fix bug with highlights introduced in 4.7.6

### Version 4.7.6
 - Make `startup_select` also see the chosen cell
 - Deprecated functions `is_cell_selected()`, `is_row_selected()`, `is_column_selected()` use the same functions but without the `is_` in the name
 - Add functions `highlight_columns()`, `highlight_rows()`, make `highlight_cells()` argument `cells` update dictionary rather than overwrite
 - Remove double click resizing of index
 - Reduce minimum column width to 1 character

### Version 4.7.5
 - Make arrowkey up and left not select column/row when at end, will add more shortcuts in a later update
 - Adjust width of auto-resized index
 - Hide header height/index width reset bars in top left if relevant options are disabled

### Version 4.7.4
 - Add option `page_up_down_select_row` will select row when using page up/down, default is `True`
 - Overhaul text, grid and highlight canvas item management, no longer deletes and redraws, keeps items using `"hidden"` and reuses them to prevent canvas item number getting high quickly
 - Edit cell resizing now only uses displayed rows/columns to fit to cell text
 - Fix `see()` not being correct when using non-default `empty_vertical` and `empty_horizontal`
 - Fix flicker when row index auto resizes
 - Rename `auto_resize_numerical_row_index` to `auto_resize_default_row_index`
 - Add option `startup_focus` to give `Sheet()` main table focus on initialization
 - Add startup argument `startup_select` use like so:

To create a cell selection box where cells 0,0 up to but not including 3,3 are selected:
```python
startup_select = (0, 0, 3, 3, "cells"),
```
To create a row selection box where rows 0 up to but not including 3 are selected:
```python
startup_select = (0, 3, "rows"),
```
To create a column selection box where columns 0 up to but not including 3 are selected:
```python
startup_select = (0, 3, "columns"),
```

### Version 4.7.3
 - Make `"begin_edit_cell"` and `"end_edit_cell"` the only two edit cell bindings for use with `extra_bindings()` function (although `"edit_cell"` still works and is the equivalent of `"end_edit_cell"`)
 - Make it so if `"begin_edit_cell"` binding returns anything other than `None` the text in the cell edit window will be that return value otherwise it will be the relevant keypress
 - Prevent auto resize from firing two redraw commands in one go
 - Fixed issue where clicking outside of cell edit window and Sheet would return focus to the Sheet widget anyway
 - Fixed issues where row/column select highlight would be lower in canvas than cell selections
 - Hopefully simplified cell edit code a bit
 - Fix undo delete columns error for python versions < 3.8

### Version 4.7.2
 - Adjustment to dark theme colors
 - Clean up repetitions of code which had virtually no performance gain anyway

### Version 4.7.1
 - Rename themes to "light blue", "light green", "dark blue", "dark green"
 - Clean up theme and color settings code

### Version 4.7.0
 - Add green theme
 - Change font to calibri

### Version 4.6.9
 - Minor fixes when using in-built ctrl functions

### Version 4.6.8
 - Fix button-1-motion scrolling issues
 - Many minor improvements

### Version 4.6.7
 - Some minor improvements
 - Fix minor display issue where sheet is empty and user clicks and drags upwards

### Version 4.6.6
 - Added multiple new functions
 - Many minor improvements
 - Decreased size of area that resize cursors will show
 - Much improved dark theme
 - Resizing cells will now resize dropdown boxes, still have to add support for moving rows/columns

### Version 4.6.5
 - Begin process of adding code to maintain multiple persistent dropdown boxes
 - Add function `delete_dropdown()` (use argument `"all"` for all dropdown boxes)
 - Add function `get_dropdowns()` to return locations and info for all boxes

### Version 4.6.4
 - Fix `create_dropdown()` issues

### Version 4.6.3
 - Fix github release

### Version 4.6.2
 - Add startup and `set_options()` arguments:
```
empty_horizontal #empty space on the x axis in pixels
empty_vertical #empty space on the y axis in pixels
show_horizontal_grid #grid lines for main table
show_vertical_grid #grid lines for main table
```
 - Fix scrollbar issues if hiding index or header
 - Change `C` in startup arguments to `parent` for clarity
 - Add Listbox using `Sheet()` recipe to tests

### Version 4.6.1
 - Clean up menus
 - Improve Dark theme

### Version 4.6.0
 - Fix bug with column drag and drop and row index drag and drop

### Version 4.5.9
 - Attempt made to remove cell editor window internal border, showing up on some platforms

### Version 4.5.8
 - Add function `set_text_editor_value()`
 - Moved internal `begin_cell_edit` code slightly
 - Make Alt-Return on text editor only increase text window height if too small

### Version 4.5.7
 - Fix download url in `setup.py`

### Version 4.5.6
 - Add `"begin_edit_cell"` to function `extra_bindings()`

### Version 4.5.5
 - Fix mouse motion binding in top left code not returning event object
 - Add some extra demonstration code to `test_tksheet.py`

### Version 4.5.4
 - Fix typo in row index code leading to text not slicing properly with center alignment

### Version 4.5.3
 - Improve text redraw slicing and positioning when cell size too small, especially for center alignments, very slight performance cost
 - If a number of lines has been set by using a string e.g. "1" for default row or header heights then the heights will be updated when changing fonts

### Version 4.5.2
 - Fix resize not taking into account header if header is set to a row in the sheet
 - Fix bugs with `display_subset_of_columns()`, argument `set_col_positions` does nothing until reworked or removed

### Version 4.5.1
 - Improve center alignment and auto resize widths issues

### Version 4.5.0
 - Make `grid_propagate()` only occur internally if both `height` and `width` arguments are used, instead of just one
 - Fix bug in `get_sheet_data()`
 - Fix mismatch with scrollbars in `show` scrollbar options

### Version 4.4.9
 - Add `header_height` argument to `set_options()`
 - Make `header_height` arguments accept integer representing pixels as well as string representing number of lines
 - Make top left header height resize bar use `header_height` and not minimum height

### Version 4.4.8
 - Fix readme python versions

### Version 4.4.7
 - Fix readme python versions

### Version 4.4.6
 - Fix Python version required

### Version 4.4.5
 - Fix typo in header code

### Version 4.4.4
 - Fix some issues with header and index not showing when values aren't string
 - Fix row index width when starting with default index and then switching later
 - Fix very minor issue with header text limit

### Version 4.4.3
 - Fix some issues with right click insert column/row
 - Fix error that occurs if row is too short on edit cell or not enough rows
 - Add function `recreate_all_selection_boxes()`
 - Add function `bind_text_editor_set()`
 - Add function `get_text_editor_value()`

### Version 4.4.2
 - Removed internal `total_rows` and `total_cols` variables, replaced with functions for better maintainability but at the cost of performance in some cases
 - Added `fix_data` argument to function `total_columns()` to even up all row lengths in sheet data
 - Removed `total_rows` and `total_cols` arguments from functions `data_reference()` and `set_sheet_data()`
 - Fix issues with function `total_columns()`
 - Redo functions `insert_column()` and `insert_row()`, hopefully they work better
 - Add argument `equalize_data_row_lengths` to function `insert_column()` default is `True`
 - Add arguments `get_index` and `get_header` to function `get_sheet_data()` defaults are `False`
 - Change some argument names for functions involving inserting/setting rows and columns
 - `display_subset_of_columns()` now tries to maintain some existing column widths
 - Add functions `sheet_data_dimensions()`, `sheet_display_dimensions()`, `set_sheet_data_and_display_dimensions()`
 - Fix bug with undo edit cell if displaying subset of columns
 - Fix max undos issue
 - `display_subset_of_columns()` now resets undo storage

### Version 4.4.1
 - Potential fix for functions `headers()` and `row_index()` when `show...if_not_sheet` arguments are utilized

### Version 4.4.0
 - Add argument `reset_col_positions` to function `headers()` default is `False`
 - Add argument `reset_row_positions` to function `row_index()` default is `False`
 - Add argument `show_headers_if_not_sheet` to function `headers()` default is `True`
 - Add argument `show_index_if_not_sheet` to function `row_index()` default is `True`
 - Adjust most `center` text draw positions slightly, `center` alignment should now actually be center
 - Improve header resizing accuracy for set to text size

### Version 4.3.9
 - Add `rc_insert_column` and `rc_insert_row` strings to extra bindings
 - Fix errors with insert column/row that occur if data is empty list
 - Fix resizing row index width/header height sometimes not working
 - Fix select all creating event when sheet is empty
 - Fix `create_current()` internal error if there are columns but no rows or rows but no columns
 - Add argument `get_cells_as_rows` to function `get_selected_rows()` to return selected cells as if they were selected rows
 - Add argument `get_cells_as_columns` to function `get_selected_columns()` to return selected cells as if they were selected columns
 - Add argument `return_tuple` to functions `get_selected_rows()`, `get_selected_columns()`, `get_selected_cells()`

### Version 4.3.8
 - Add function `set_currently_selected()`
 - Make `get_selected_min_max()` return `(None, None, None, None)` if nothing selected to prevent error with tuple unpacking e.g `r1, c1, r2, c2 = get_selected...`
 - Fix potential issue with internal function `set_col_width()` if `only_set...` argument is `True`
 - Fix typo in function `create_selection_box()`
 - Add `right_click_select` and `rc_select` to function `enable_bindings()`, is enabled as well if other right click functionality is enabled
 - Add `"all_select_events"` as option in function `extra_bindings()` to quickly bind all selection events in the table, including deselection to a single function

### Version 4.3.7
 - Make row height resize to text include row index values
 - Fix bug with function `.set_all_cell_sizes_to_text()` where empty cells would result in minimum sizes
 - Fix typo error in function `.set_all_cell_sizes_to_text()`

### Version 4.3.6
 - Add function `get_currently_selected()`
 - Add functions `self.sheet_demo.set_all_column_widths()`, `self.sheet_demo.set_all_row_heights()`, `self.sheet_demo.set_all_cell_sizes_to_text()`
 - Add argument `set_all_heights_and_widths` to startup, use `True` or `False`
 - Improve performance of all resizing functions
 - Make resizing detect text dimensions for full row/column, not just what is displayed

### Version 4.3.5
 - Fix bug with `row_index()` function
 - Fix `headers` set to `int` at start up not working if `headers` is `0`

### Version 4.3.4
 - Fix bug if `row_drag_and_drop_perform` or `column_drag_and_drop_perform` is False
 - Fix display issue with undo drag and drop rows/columns
 - Add `"unbind_all"` as possible argument for function `extra_bindings()`
 - Add `"enable_all"` as possible argument for function `enable_bindings()`
 - Add `"disable_all"` as possible argument for function `disable_bindings()`

### Version 4.3.3
 - Add function `set_sheet_data()` use argument `verify = False` to prevent verification of types
 - Add startup arguments `row_drag_and_drop_perform` and `column_drag_and_drop_perform` and add to `set_options()`
 - Fix double selection with drag and drop rows
 - Add `data` argument to startup arguments

### Version 4.3.2
 - Fix some highlighted cells not being reset when displaying subset of columns

### Version 4.3.1
 - Fix highlighted cells not blending with drag selection

### Version 4.3.0
 - Fix basic bindings bug

### Version 4.2.9
 - Fix typo in popup menu

### Version 4.2.8
 - Improve performance of drag selection
 - Improve `row_index()` and `headers()` functions, add examples
 - Add `"single_select"` and `"toggle_select"` to `enable_bindings()`

### Version 4.2.7
 - Fix bug with undo deletion

### Version 4.2.6
 - Fix bug with insert row
 - Deprecate `change_color()` use `set_options()` instead
 - Fix display bug where resizing row index width or header height would result in small selection boxes
 - Rework popup menus, use `enable_bindings()` with `"right_click_popup_menu"` to make them work again
 - Allow additional binding of right click to a function alongside `Sheet()` popup menus
 - Add function `get_sheet_data()`
 - Add `x_scrollbar` and `y_scrollbar` to `show()`, `hide()` and `Sheet()` startup
 - Removed `show` argument from startup, added `show_table`, default is `True`
 - Add option to automatically resize row index width when user has not set new indexes, default is `True`
 - Rework the top left rectangle and add select all on left click if drag selection is enabled
 - Change appearance of popup menus
 - Change popup menu color for dark theme
 - Fix wrong popup menu on right click with selected rows/columns

### Version 4.2.5
 - Fix bug where highlighted background or foreground might not be in the correct column when displaying a subset of columns
 - Change colors of selected rows and columns
 - Add color options: 
```python
header_select_column_bg
header_select_column_fg
row_index_select_row_bg
row_index_select_row_fg
selected_rows_border_color
selected_rows_background
selected_rows_foreground
selected_columns_border_color
selected_columns_background
selected_columns_foreground
```
 - Fix various issues with displaying correct colors in certain circumstances
 - Change dark theme colors slightly
                              
### Version 4.2.4
 - Fix PyPi release version

### Version 4.2.3
 - Fix bug with right click delete columns and undo
 - Right click in main table when over selected columns/rows now brings up column/row menu
 - Add internal cut, copy, paste, delete and undo to usable `Sheet()` functions

### Version 4.2.2
 - Fix bug with center alignment and display subset of columns
 - Fix bug with highlighted cells showing as selected when they're not

### Version 4.2.1
 - Fix resize row height bug
 - Fix deselect bug
 - Improve and fix many bugs with toggle select mode

### Version 4.2.0
 - Fix extra selection box drawing in certain circumstances
 - Fix bug in Paste
 - Remove some unnessecary code

### Version 4.1.9
 - Fix bugs introduced with 4.1.8

### Version 4.1.8
 - Overhaul internal selections variables and workings
 - Replace functions `get_min`/`get_max` `selected` `x` and `y` with `get_selected_min_max()`
 - Fix bugs with functions `total_columns()` and `total_rows()` and set default argument `mod_data` to `True`
 - Change drag and drop so that it modifies data with or without an extra binding set
 - Add functions `set_cell_data()`, `set_row_data()`, `set_column_data()`
 - Remove bloat in `get_highlighted_cells()`
 - Prepare many functions for implementation of control + click

### Version 4.1.7
 - Fix typo in `get_cell_data()`
 - Add function `height_and_width()`
 - Fix height and width in widget startup

### Version 4.1.6
 - Fix cell selection after editing cell ending with mouse click

### Version 4.1.5
 - Separate drag selection `extra_bindings()` into cells, rows and columns
 - Add shift select and separate it from left click and right click in `extra_bindings()`
 - Changed and hopefully made better all responses to `extra_bindings()` functions
 - Added selection box to selected rows/columns

### Version 4.1.4
 - Fix bugs with drag and drop rows/columns
 - Clean up drag and drop code

### Version 4.1.3
 - Improve speed of `get_selected_cells()`
 - Add function `get_selection_boxes()`
 - Remove more unnessecary loops
 - Added argument `show_selected_cells_border` to start up
 - Added function `set_options()`
 - Added undo to drag and drop rows/columns
 - Fixed bug with drag and drop columns when displaying a subset of columns

### Version 4.1.2
 - Move text up slightly inside cells
 - Fix selected cells border not showing sometimes
 - Replace some unnessecary loops

### Version 4.1.1
 - Improve select all speed by 20x
 - Shrink cell size slightly
 - Added options to hide row index and header at start up
 - Change looks of ctrl x border
 - Fix minor issue where two borders would draw in the same place

### Version 4.1.0
 - Added undo to insert row/column and delete rows/columns
 - Changed the default looks
 - Deprecated function `select()` and added `toggle_select_cell()`, `toggle_select_row()` and `toggle_select_column()`
 - Added functions `is_cell_selected()`, `is_row_selected()`, `is_column_selected()`,
   `get_cell_data()`, `get_row_data()` and `get_column_data()`
 - Changed behavior of `deselect()` function slightly to now accept a cell from arguments `row = ` and `column = ` as well as `cell = `
 - Added cell selection borders and start up option `selected_cells_border_color` as well as to the function `change_color()`
 - Removed the main table variable `selected_cells` and reduced memory use
 - Fixed a bug with editing a cell while displaying a subset of columns
 - Renamed `RowIndexes` class to `RowIndex`
 - Added separate test file

### Version 3.1 - 4.0.1
 - Fixed import errors

### Version 3.1:
 - Added `frame_background` parameter to `Sheet()` startup and `change_color()` function
 - Added more keys to edit cell binding
 - Separated `edit_bindings` for the function `enable_bindings()` into the following:
	- `"cut"`
	- `"copy"`
	- `"paste"`
	- `"delete"`
	- `"edit_cell"`
	- `"undo"`
	Note that the argument `"edit_bindings"` still works
 - Separated the tksheet classes and variables into different files

### Version 3.0:
 - Fixed bug with function `display_subset_of_columns()`
 - Fixed issues with row heights and column widths adjusting too small
 - Improved dark theme
 - Added top left rectangle to hide
 - Added insert and delete row and column to right click menu
 - Some general code improvements

### Version 2.9:
 - The function `display_columns()` has been changed to `display_subset_of_columns()` and some paramaters have been changed
 - The variables `total_rows` and `total_cols` have been reworked
 - Sheets can now be created with the parameters `total_rows` and `total_columns` to create blank sheets of certain dimensions
 - The functions `total_rows()` and `total_columns()` have been added to set or get the dimensions of the sheet
 - `"edit_bindings"` has been added as an input for the function `enable_bindings()`
 - `show()` and `hide()` have been reworked to allow hiding of the header and/or row index

### Version 2.6 - 2.9:
 - delete_row() is now delete_row_position()
 - insert_row() is now insert_row_position()
 - move_row() is now move_row_position()
 - delete_column() is now delete_column_position()
 - insert_column() is now insert_column_position()
 - move_column() is now move_column_position()
 - The original functions have been changed to behave as they read and actually modify the table data
