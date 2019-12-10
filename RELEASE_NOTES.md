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
