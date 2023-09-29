### Version 6.2.5
#### Fixed:
- Error when zooming or using `see()` with empty table

### Version 6.2.4
#### Added:
- Zoom in/out bindings control + mousewheel
- Zoom in bindings control + equals, control + plus
- Zoom out binding control + minus

### Version 6.2.3
#### Fixed:
- Bug with `format_row` using "all"

### Version 6.2.2
#### Fixed:
- 2 pixel misalignment of index/header and table when scrolling
- Undo cell edit not scrolling window

#### Added:
- Functionality to scroll multiple Sheets when scrolling one particular Sheet

### Version 6.2.1
#### Fixed:
- Editing header and overflowing text editor so text wraps while using `auto_resize_columns` causes text editor to be out of position
- `insert_columns()` when using an `int` for number of blank columns creates incorrect list layout

### Version 6.2.0
#### Fixed:
- Removed some type hinting that was only available to python versions >= 3.9

### Version 6.1.9
#### Fixed:
- Issues that follow selection boxes being recreated such as resizing columns/rows

### Version 6.1.8
#### Added:
- [#180](https://github.com/ragardner/tksheet/issues/180)

### Version 6.1.7
#### Fixed:
- [#183](https://github.com/ragardner/tksheet/issues/183)
- [#184](https://github.com/ragardner/tksheet/issues/184)

### Version 6.1.5
#### Fixed:
- `extra_bindings()` not binding functions
- [#181](https://github.com/ragardner/tksheet/issues/181)

### Version 6.1.4
#### Fixed:
- Error with setting/getting header font
- Setting fonts with `set_options()` not working
- Setting fonts after table initialization didn't refresh selection boxes or top left rectangle dimensions

#### Changed:
- Replaced wildcard imports
- Format code with 120 line length

### Version 6.1.3
#### Fixed:
- Missing imports
- Bug with `delete_rows()`
- Bug with hidden columns, cell options and deleting columns with the inbuilt right click menu
- Bug with a paste that expands the sheet where row lengths are uneven
- Poor box selection for cut and copy with multiple selected boxes
- Bug with `insert_columns()` with uneven row lengths
- Bug with `insert_rows()` if new data contains row lengths that are longer than sheet data
- Errors that occur when dragging and dropping rows/columns beyond the window

### Version 6.1.2
#### Fixed:
- Further potential issues with moving columns where row lengths are uneven
- Potential issues with using `move_rows()` where the provided index is larger than the number of rows
- `dropdown_sheet()` causing error

### Version 6.1.1
#### Fixed:
- [#177](https://github.com/ragardner/tksheet/issues/177)

### Version 6.1.0
#### Fixed:
- Error with using `None` to create a dropdown for the entire index when index is not `int`

### Version 6.0.5
#### Fixed:
- Bugs with using `None` to create dropdowns/checkboxes for entire header/index
- Bug with creating dropdowns in row index
- Bug with opening dropdowns in row index
- Bug with `insert_rows()` if `rows` had more columns than existing sheet
- Bug with `delete_column_dropdown()`
- Bug with using external move rows/columns functions where `index_type` wasn't `"displayed"`
- Bug with hidden rows and cutting/copying

#### Changed:
- Formatted code using `Black`

### Version 6.0.4
#### Fixed:
- Bug with setting header height when header is set to `int`
- Bug with clicking top left rectangle bars to reset header height / index width
- Wrong colors showing when having rows and columns selected at the same time
- Wrong colors showing for particular selection scenarios
- Control select enabled when using `enable_bindings("all")` when it should only be enabled if using `enable_bindings("ctrl_select")`
- When control select was disabled holding control/command would disable basic sheet bindings

#### Added:
- Add additional binding for `sheet.bind_event()`. `"<<SheetModified>>"` and `"<<SheetRedrawn>>"` are the two existing event options but currently only an empty dict is the associated data

### Version 6.0.3
#### Fixed:
- `disable_bindings()` disable all bindings not disabling edit header/index capability
- Undo not recreating all former selection boxes with cell edits
- Undoing paste where sheet was expanded and rows were hidden didn't remove added rows
- Bug with hidden rows and index highlights
- Bug with hidden rows and text editor newline binding in main table
- Bug with hidden rows and drag and drop rows
- Issue where columns/rows that where selected already would not open dropdown boxes or toggle checkboxes

#### Removed:
- External function `edit_cell()`, use `open_cell()` and `set_currently_selected()` as alternative

#### Added:
- Control / Command selecting multiple non-consecutive boxes
- Functions `checkbox_cell`, `checkbox_row`, `checkbox_column`, `checkbox_sheet`, `dropdown_cell`, `dropdown_row`, `dropdown_column`, `dropdown_sheet` and their related deletion functions

#### Changed:
- `extra_bindings()` undo event type `"insert_row"` renamed `"insert_rows"`, `"insert_col"` renamed `"insert_cols"`
- `extra_bindings()` undo event `.storeddata` has changed for row/column drag drop undo events to `("move_cols" or "move_rows", original_indexes, new_indexes)`
- Initialization argument `max_rh` renamed `max_row_height`
- Initialization argument `max_colwidth` renamed `max_column_width`
- Initialization argument `max_row_width` renamed `max_index_width`
- `header_dropdown_functions`/`index_dropdown_functions`/`dropdown_functions` now return dict
- Rename internal function `create_text_editor()`
- Select all now maintains existing currently selected cell when using Control/Command - a
- Renamed internal functions `check_views()` `check_xview()` `check_yview()`
- Default argument `redraw` changed to `True` for dropdown/checkbox creation functions
- Internal function `open_text_editor()` keyword argument `set_data_on_close` defaulted to `True`
- Merge internal functions `edit_cell_()` and `open_text_editor()`, remove function `edit_cell_()`
- Better graphical indicators for which columns/rows are being moved when dragging and dropping
- Better retainment of selection boxes after undo cell edits

### Version 6.0.2
#### Fixed:
- Right click delete rows bug [PR#171](https://github.com/ragardner/tksheet/pull/171)

### Version 6.0.1
#### Fixed:
- Data getting with `get_displayed = True` not returning checkbox text and also in the main table dropdown box text
- None being displayed on table instead of empty string
- `insert_columns()`/`insert_rows()` with hidden columns/rows not working quite right

#### Added:
- Hidden rows capability use `display_rows()`/`hide_rows()` and read the documentation for more info

#### Changed:
- Some functions have had their `redraw` keyword argument defaulted to `True` instead of `False` such as `highlight_cells()`
- `hide_rows()`/`hide_columns()` `refresh` arg changed to `redraw`
- `hide_rows()`/`hide_columns()` now uses displayed indexes not data

### Version 6.0.0
#### Fixed:
- Undo added to stack when no changes made with cut, paste, delete
- Using generator with `set_column_widths()`/`set_row_heights()` would result in lost first width/height
- Header/Index dropdown `modified_function` not sending modified event
- Escape out of dropdown box doesn't reset arrow orientation

#### Added:
- Cell formatters, thanks to [PR#158](https://github.com/ragardner/tksheet/pull/158)
- `format_cell()`, `format_row()`, `format_column()`, `format_sheet()`
- Options for changing output to clipboard delimiter, quotechar, lineterminator and option for setting paste delimiter detection
- Bindings `"up" "down" "left" "right" "prior" "next"` to enable arrowkey bindings individually use with `enable_bindings()`/`disable_bindings()`
= Dropdowns now have a search feature which searches their values after the entry box is modified (if dropdown is state `normal`)
- Dropdown kwargs `search_function`, `validate_input` and `text`

#### Removed:
- Startup arg `ctrl_keys_over_dropdowns_enabled`, also removed in `set_options()`
- `set_copy` arg in `set_cell_data()`
- `return_copy` arg in data getting functions but will not generate error if used as keyword arg

#### Changed:
- `index_border_fg` and `header_border_fg` no longer work, they now use the relevant grid foreground options
- `"dark"`/`"black"` themes
- Dropdowns now default to state `"normal"` and validate input by default
- `set_dropdown_values()`/`set_index_dropdown_values()`/`set_header_dropdown_values()` keyword argument `displayed` changed to `set_value` for clarity
- Checkbox click extra binding and edit cell extra binding (when associated with a checkbox click) return `bool` now, not `str`
- `get_cell_data()`/`get_row_data()`/`get_column_data()` have had an overhaul and have different keyword arguments, see documentation for more information. They also now return empty string/s if index is out of bounds (instead of `None`s)
- Rename some internal functions for consistency
- Extra bindings delete key now returns dict instead of list for boxes

### Version 5.6.8
#### Fixed:
 - [#159](https://github.com/ragardner/tksheet/issues/159)

#### Changed:
 - `"dark"` theme now looks more appropriate for MacOS dark theme

### Version 5.6.7
Fixed:
 - [#157](https://github.com/ragardner/tksheet/issues/157)
 - Double click row height resize not taking into account index checkbox text

Changed:
 - Some internal variables local to redrawing functions for clarity

### Version 5.6.6
Fixed:
 - `delete_rows()`/`delete_columns()` incorrect row heights/column widths after use
 - Tab key not seeing cell if out of sight
 - Row height / column width resizing with mouse incorrect by 1 pixel
 - Row index not extending if too short when changing a specific index
 - Selected rows/columns border fg not displaying in cell border
 - Various minor text placements
 - Edit index/header prematurely resizing height/width

Improved:
 - Significant performance improvements in redrawing table, especially when simply selecting cells
 - All themes
 - Can now drag and drop columns and rows with or without shift being held down, mouse cursor changes to hand when over selected

Changed:
 - Dropdown box colors now use popup menu colors
 - Checkboxes no longer have X inside, instead simply a smaller more distinct rectangle to improve redrawing performance

### Version 5.6.5
Fixed:
 - [#152](https://github.com/ragardner/tksheet/issues/152)

### Version 5.6.4
Fixed:
 - `set_row_heights()`/`set_column_widths()` failing to set if iterable was empty
 - `delete_rows()`/`delete_columns()` failing to delete row heights, column widths if arg is empty
 - Edges of grid lines appearing when not meant to

Changed:
 - Add `redraw` option for `change_theme()`

### Version 5.6.3
Fixed:
 - Dropdown boxes in main table not opening in certain circumstances
 - Scrollbars not having correct column/rowspan
 - Error when moving header height
 - Resizing arrows still showing up when hiding header/index if height/width resizing enabled

Added:
 - `header` startup argument, same as `headers`

Changed:
 - Increase default maximum undos from 20 to 30
 - Removed dropdown box border, to get them back use `show_dropdown_borders = True` on startup or with `set_options()`
 - Removed unnecessary variables and tidied `__init__` code

### Version 5.6.2
Fixed:
 - [#154](https://github.com/ragardner/tksheet/issues/154)
 - `delete_column()` not working with hidden columns
 - `insert_column()`/`insert_columns()` not working correctly with hidden columns

Added:
 - `delete_columns()`, `delete_rows()`

Changed:
 - `delete_column()`/`delete_row()` now use `delete_columns()`/`delete_rows()` internally
 - `insert_column()` now uses `insert_columns()` internally

### Version 5.6.1
Fixed:
 - [#153](https://github.com/ragardner/tksheet/issues/153)

Changed:
 - Adjust dropdown arrow sizes for more consistency for varying fonts

### Version 5.6.0
Fixed:
 - Issues and possible errors with dropdowns/checkboxes/cell edits
 - `delete_dropdown()`/`delete_checkbox()` issues

Changed:
 - Deprecated external functions `create_text_editor()`/`get_text_editor_value()`/`bind_text_editor_set()` as they no longer worked both externally and internally, use `open_cell()` instead
 - Renamed internal function `get_text_editor_value()` to `close_text_editor()`

Improved:
 - Slightly boost performance if there are many cells onscreen and gridlines are showing
 - You can now use the scroll wheel in the header to vertically scroll if there are multiple lines in the column headers
 - Improvements to text editor
 - Text now draws slightly closer to cell edges in certain scenarios
 - Improved visibility of dropdown box against sheet background
 - Improved dropdown window height
 - black theme selected cells border color
 - light green theme selected cells background color

### Version 5.5.3
Fixed:
 - Error on start

### Version 5.5.2
Fixed:
 - Row index missing itertools repeat
 - Header checkbox, index checkbox bugs
 - Row index editing errors
 - `move_columns()`/`move_rows()` redrawing when not supposed to
 - `move_row()`/`move_column()` error
 - `delete_out_of_bounds_options()` error
 - `dehighlight_cells()` incorrectly wiping all cell options when `"all"` is used

Changed:
 - `set_column_widths()`/`set_row_heights()` can now receive any iterable if `canvas_positions` is `False`
 - Tidied internal row/column moving code

### Version 5.5.1
Fixed:
 - `display_columns()` no longer redraws if `deselect_all` is `True` even when `redraw` is `False`
 - `extra_bindings()` cell editors carrying out cell edits even if validation function returns `None`
 - Index and header alignments wrongly associated with column and row alignments if `align_header`/`align_index` were `False`
 - Undo drag and drop wrong position if columns move back to higher index
 - Error when shift b1 press in headers/index while using extra bindings

Changed:
 - Internal variable name `data_ref` to `data`
 - `get_currently_selected()` and `currently_selected()` see documentation for more information
 - The way `extra_bindings()` + `begin_edit_cell` works, now if `None` is returned then the cell editor will not be opened
 - Paste repeats when selection box is larger than pasted items and is a multiple of pasted box
 - `move_row()`/`move_column()` now internally use `move_rows()`/`move_columns()`
 - `black` theme text to be a lot brighter

Added:
 - Spacebar to edit cell keys
 - Function `open_cell()` which uses currently selected box and mouse event
 - Function `data` see the documentation for more info
 - Functions `get_cell_alignments()`, `get_row_alignments()`, `get_column_alignments()`, `reset_all_options()`, `delete_out_of_bounds_options()`
 - Functions `move_columns()`, `move_rows()`

### Version 5.5.0
 - Deprecate functions `display_subset_of_columns()`, `displayed_columns()`
 - Change `get_text_editor_value()` function argument `destroy_tup` to `editor_info`
 - Change `display_columns()` function argument `indexes` to `columns`
 - Change `display_columns()` function argument `enable` to `all_columns_displayed`
 - Change some internal variable names
 - Remove all calls to `update()` from `tksheet`
 - Fix row index `"e"` text alignment bug
 - Fix checkbox text not showing in main table when not using west alignment
 - Fix `total_rows()` bug
 - Fix dropdown arrow being asymmetrical with different font sizes
 - Fix incorrect default headers while using hidden columns
 - Fix issue with `get_sheet_data()` if including index and header not putting a corner cell in the top left
 - Modify redrawing code, slightly improve redrawing efficiency in some scenarios
 - Add dropdown boxes, check boxes and cell editing to index
 - Add options `edit_cell_validation` which is an option for `extra_bindings`, see the documentation for more info
 - Add function `yield_sheet_rows()` which includes default index and header values if using them
 - Add function `hide_columns()` which allows input of columns to hide, instead of columns to display
 - Add functions related to index editing, checkboxes and dropdown boxes
 - Add default argument `show_default_index_for_empty`
 - Add `Edit header` option to in-built right click menu if both are enabled
 - Add `Edit index` option to in-built right click menu if both are enabled
 - Add `Edit cell` option to in-built right click menu if both are enabled
 - Documentation updates