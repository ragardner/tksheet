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
