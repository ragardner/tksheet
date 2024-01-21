### Version 7.0.0
#### Removed:
- `edit_cell_validation` from `Sheet()` initialization, `set_options()` and everywhere else due to confusion. Use the new function `edit_validation()` instead.
- Functions (use spans instead for the same purpose):
    - `header_checkbox`
    - `index_checkbox`
    - `format_sheet`
    - `delete_sheet_format`
    - `dropdown_sheet`
    - `delete_sheet_dropdown`
    - `checkbox_sheet`
    - `delete_sheet_checkbox`

- Parameters:
    - `verify` from functions `set_column_widths` and `set_row_heights`

- Old unused and deprecated parameters for:
    - `get_cell_data()`, `get_sheet_data()`, `get_row_data()`, `get_column_data()`, `yield_sheet_rows()`
    - All checkbox and dropdown creation functions

#### Fixed:
- Grid lines now properly raised above highlighted rows/columns
- Deselect events firing when unnecessary

#### Changed:
- Using `extra_bindings()` with `"end_edit_cell"`/`"edit_cell"` no longer requires a return value in your bound function to set the cell value to. For end user cell edit validation use the new function `edit_validation()` instead
- Changed functions:
    - Parameters and behavior:
        - `checkbox` now is used to create checkboxes and utilises spans
        - `delete_header_dropdown` default argument changed to required argument and can no longer delete a header wide dropdown
        - `delete_index_dropdown` default argument changed to required argument and can no longer delete an index wide dropdown

    - Parameters only:
        - `move_columns` most parameters changed
        - `move_rows` most parameters changed
        - `click_checkbox` most parameters changed
        - `insert_row` parameter `idx` default argument changed to `None`
        - `insert_column` parameter `idx` default argument changed to `None`
        - `insert_rows` parameter `idx` default argument changed to `None`
        - `insert_columns` parameter `idx` default argument changed to `None`

    - Renamed:
        - `align` -> `table_align`
- Reorganised order of functions in `sheet.py` to match documentation

#### Added:
- Method `edit_validation(func: Callable | None = None) -> None` to replace `edit_cell_validation`
- New methods for getting and setting data
- `bind` now also accepts `"<<SheetModified>>"` and `"<<SheetRedrawn>>"` arguments
- Redo, which is enabled when undo is enabled, use by pressing ctrl/cmd + shift + z
- Named spans for sheet options such as highlight, format
- Ctrl/cmd click deselect
- Ability to make currently selected box border different color to selection box border

#### Other Changes:
- Overhaul and totally change event data sent to functions bound by `extra_bindings()`
- Deselect events are now labelled as "select" events, see the docs on `"selection_boxes"` for more information
- Overhaul internal selection box workings
- Rename tksheet files
- Pressing escape on text editors no longer generates an edit cell/header/index event

### Version 6.3.5
#### Fixed:
- Error with function `set_currently_selected()`

#### Addressed:
- [207](https://github.com/ragardner/tksheet/issues/207)

### Version 6.3.4
#### Fixed:
- [#205](https://github.com/ragardner/tksheet/issues/205)

### Version 6.3.3
#### Improved:
- Term searching improved when typing in a normal state dropdown box

#### Removed:
- `bind_event()` function, `bind()` to be used instead

### Version 6.3.2
#### Changed:
- tksheet no longer supports Python 3.6, only versions 3.7+

#### Removed:
Due to errors caused when using Python versions < 3.8 the following functions have been removed:
- __bool__
- __len__

#### Fixed:
- Sheet init error since version `6.3.0` when running Python 3.7

### Version 6.3.1
#### Fixed:
- Two `EditCellEvent`s being emitted, removed the one with `None` as `text` attribute
- Selection box and currently selected box in different places when tab key pressed with single cell selected
- Return key not working in dropdown box when mouse pointer is outside of dropdown
- Visual issue: Dropdown arrow staying up when clicking on same cell with readonly dropdown state
- Visual issue: Dropdown row/column all arrows in up position when one box is open

#### Removed:
- Many parameters from internal functions dealing with text editor to simplify code

#### Changed:
- Cell selection doesn't move on Return key when cell edit using text editor was invalid
- Events for `extra_bindings()` end cut, delete and paste are no longer emitted if no changes were made

#### Improved:
- Term searching improved when typing in a normal state dropdown box

### Version 6.3.0
#### Fixed:
- Cell editor right click not working
- Cell editor select all not working

#### Added:
- Some methods to `Sheet` objects, see [documentation](https://github.com/ragardner/tksheet/wiki/Version-6#sheet-methods) for more information.
    - __bool__
    - __len__
    - __iter__
    - __reversed__
    - __contains__

### Version 6.2.9
#### Fixed:
- Incorrect row being targeted with hidden rows and text editor newline binding (potential error)

### Version 6.2.8
#### Fixed:
- [#203](https://github.com/ragardner/tksheet/issues/203)

### Version 6.2.7
#### Fixed:
- [#202](https://github.com/ragardner/tksheet/issues/202)

### Version 6.2.6
#### Fixed:
- [#201](https://github.com/ragardner/tksheet/issues/201)
- The ends of grid lines were incorrectly displaying connections with one another when only showing horizontal or vertical grid
- When a cell dropdown and a row checkbox were in the same cell both would be drawn but only one would function, this has been changed to give dropdown boxes priority
- Index text overlapping checkbox when alignment is `"right"`/`"e"`/`"east"` and index is not wide enough

#### Added:
- [#200](https://github.com/ragardner/tksheet/issues/200)

### Version 6.2.5
#### Fixed:
- Error when zooming or using `see()` with empty table
- Add initialization option `zoom` which is an `int` representing a percentage of the font size, `100` being no change

#### Removed:
- Some unnecessary internal variables

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

### Version 5.4.1
 - Fix bugs with functions `readonly_header()`, `checkbox()` and `header_checkbox()`
 - Clarify table colors documentation

### Version 5.4.0
 - Improve documentation a bit

### Version 5.3.9
 - Fix bug with paste without anything selected
 - Minor fix to highlight functions which could cause an error using certain args
 - Fix `ctrl_keys_over_dropdowns_enabled` initialization arg not setting

### Version 5.3.8
 - Fix focus out of table cell editor by clicking on header not setting cell

### Version 5.3.7
 - Fix issue with `enable_bindings()`/`disable_bindings()` no longer accept a tuple
 - Fix drag and drop bugs introduced in 5.3.6
 - Fix bugs with header set to integer introduced in 5.3.6

### Version 5.3.6
 - Editable and readonly dropdown boxes and checkboxes in the header
 - Add readonly functions to header
 - Editable header using `enable_bindings("edit_header")
 - Display text next to checkboxes in cells using `text` argument (is not considered data just for display, the data is either `True` or `False` in a checkbox)
 - Move position of checkboxes and dropdown boxes to top left of cells
 - Enable dragging and dropping columns when there are hidden columns, including undo
 - Checkboxes no longer editable using `Delete`, `Paste` or `Cut`
 - Dropdown boxes no longer editable using `Delete`, `Paste` or `Cut` unless dropdown box has a value which matches, override this behaviour with new option `ctrl_keys_over_dropdowns_enabled`
 - More logical behaviour around checkboxes and dropdown boxes and bindings
 - Many improvements and bug fixes

### Version 5.3.5
 - Fix control commands not working when top left has focus
 - Fix bug when undoing paste where extra rows/columns were added during the paste (disabled by default)

### Version 5.3.4
 - Unnecessary folder deletion
 - Add new theme `dark`

### Version 5.3.3
 - Add `add` parameter to `.bind()` function (only works for bindings which are not the following: `"<ButtonPress-1>"`, `"<ButtonMotion-1>"`, `"<ButtonRelease-1>"`, `"<Double-Button-1>"`, `"<Motion>"` and lastly whichever is your operating systems right mouse click button)

### Version 5.3.2
 - Fix Backspace binding in main table

### Version 5.3.1
 - Fix minor issues with `startup_select` argument
 - Fix extra menus being created in memory

### Version 5.3.0
 - Fix shift select error in main table when nothing else is selected

### Version 5.2.9
 - Fix text editor not opening in highlighted cells
 - Linux return key in cell editor fix
 - Fix potential error with check box / dropdown

### Version 5.2.8
 - Make dropdown box border consistent color
 - Fix error with modifier key press on editable dropdown boxes

### Version 5.2.7
 - Improve looks of dropdown arrows and fix slight overlapping with text
 - Check boxes no longer show `Tr` and `Fa` when cell is too small for check box for large enough for text

### Version 5.2.6
 - Fix issues with editable (`"normal"`) dropdown boxes and focus out bindings
 - Editable dropdown boxes now respond normally to character, backspace etc. key presses

### Version 5.2.5
 - Fix bug with highlighted cells being readonly

### Version 5.2.4
 - Fix user bound right click event not firing due to right click context menu
 - Fix issues with hidden columns and right click insert columns

### Version 5.2.3
 - Fix error on deleting row with dropdown box
 - Fix issues with deleting columns with check boxes, dropdown boxes

### Version 5.2.2
 - Improve looks for dropdown arrows
 - Fix grid lines in certain drop down boxes
 - Fix dropdown scroll bar showing up when not needed
 - After resizing rows/columns if mouse is in same position cursor reacts accordingly again

### Version 5.2.1
 - Remove `see` argument from `create_checkbox()` and `create_dropdown()` because they rely on data indexes whereas `see()` relies on displayed indexes
 - Fix undo not working with check box toggle
 - Fix error on clicking sheet empty space
 - Improve check box looks
 - Fix bugs related to hidden columns and dropdowns/checkboxes
 - Disable align from working with `create_dropdown()` because it was broken with anything except left alignment
 - Code cleanup

### Version 5.2.0
 - Adjust dropdown box heights slightly
 - Fix extra bindings begin edit cell making cell edit delete contents
 - Make many events namedtuples for better clarity
 - Add checkbox functionality

### Version 5.1.7
 - Fix double button-1 not editing cell
 - Fix error with edit cell extra bindings

### Version 5.1.6
 - Fix dropdown box and checkbox triggering off of non-Return keypresses
 - Fix issues with readonly cells and rows and control actions such as cut, delete, paste
 - Begin to add checkbox code
 - Some code cleanup

### Version 5.1.5
 - Fix center, e text alignment with dropdowns
 - Fix potential error when closing root Tk() window

### Version 5.1.4
 - Fix error when hiding columns and creating dropdown boxes
 - Fix dropdown box sizes

### Version 5.1.3
 - Major fixes for dropdown boxes and possibly cell editor
 - Fix mouse click outside cell edit window not setting cell value
 - International characters should work to edit a cell by default
 - Minor improvements

### Version 5.1.2
 - Bugfixes and improvements related to `5.1.0`

### Version 5.1.1
 - Bugfixes and improvements related to `5.1.0`

### Version 5.1.0
 - Overhaul dropdowns
 - Replace ttk dropdown widget

### Version 5.0.34
 - Fix `None` returned with `get_row_data()` when `get_copy` is `False`
 - Add bindings for column width resize and row height resize
 - Fix error with function `get_dropdowns()`

### Version 5.0.33
 - Add default argument `selection_function` to `create_dropdown()`

### Version 5.0.32
 - Add sheet refreshing to delete_row(), delete_column()

### Version 5.0.31
 - Fixed insert rows/columns below/left working incorrectly

### Version 5.0.30
 - Fixed issue with cell editor stripping whitespace from the end of value

### Version 5.0.29
 - Update documentation
 - `insert_row` argument `add_columns` now `False` by default for better performance

### Version 5.0.28
 - Various minor improvements

### Version 5.0.27
 - Internally set `all_columns_displayed` to `True` when user chooses full list of columns as argument for `display_columns()`/`display_subset_of_columns()`

### Version 5.0.26
 - Add option `expand_sheet_if_paste_too_big` in initialization and `set_options()` which adds rows/columns if paste is too large for sheet, disabled by default

### Version 5.0.25
 - Select all now needs to be enabled separately from `"drag_select"` using `"select_all"`
 - Separate Control / Command bindings based on OS

### Version 5.0.24
 - Fix documentation link

### Version 5.0.23
 - Fix page up/down cell select event not being created if user has cell selected

### Version 5.0.22
 - Put all begin extra functions inside try/except, returns on exception

### Version 5.0.21
 - Attempt to fix PyPi version issue

### Version 5.0.2
 - `set_sheet_data()` no longer verifies by default that data is list of lists (inbuilt functionality such as cell editing still requires list of lists though)
 - Initialization argument `data` allows tuple or list

### Version 5.0.19
 - Scrolling with arrowkeys improvements

### Version 5.0.18
 - Fix bug with hiding scrollbars not working

### Version 5.0.17
 - Cell highlight tkinter colors no longer case sensitive
 - Add additional items to `end_edit_cell` event responses

### Version 5.0.16
 - Add function `default_column_width()` and add `column_width` to function `set_options()`
 - Add functions `get_dropdown_values()` and `set_dropdown_values()`
 - Add row and column arguments to `get_dropdown_value()` to get a specific dropdown boxes value
 - `create_dropdown()` argument `see` is now `False` by default

### Version 5.0.15
 - Add functions `popup_menu_add_command()` and `popup_menu_del_command()` for extra commands on in-built right click popup menu
 - Update documentation

### Version 5.0.14
 - Remove some unnecessary code (`preserve_other_selections` arguments in some functions)
 - Add documentation file, update readme file

### Version 5.0.13
 - Add extra variable to undo event
 - Edit cell no longer creates undo if cell is left unchanged

### Version 5.0.12
 - Remove/add scrollbars depending on if window can be scrolled
 - Fix bug with delete key

### Version 5.0.11
 - Add default height and width if a height is used without a width or vice versa

### Version 5.0.1
 - Add right (`e`) cell alignment for index and header

### Version 5.0.0
 - Fix errors with `insert_column()` and `insert_columns()`

### Version 4.9.99
 - Fixed bugs with row copying where `list(repeat(list(repeat(` was used in code to create empty list of lists
 - Made cell resize to text (width only) take dropdown boxes into account

### Version 4.9.98
 - Fix error with dropdown box close while showing all columns

### Version 4.9.97
 - Fix bug with `delete_row()` and `delete_column()` functions when used with default arguments

### Version 4.9.96
 - Fix bug with dragging scrollbar when columns are shorter than window
 - Add startup argument `after_redraw_time_ms` default is `100`

### Version 4.9.95
 - Fix bug with `insert_rows()`
 - Add `redraw` default argument to many functions, default is `False`

### Version 4.9.92 - 4.9.94
 - Hopefully fix Linux mousewheel

### Version 4.9.91
 - Fix auto resize issue

### Version 4.9.9
 - Attempt to fix Linux mousewheel scrolling
 - Add `enable_edit_cell_auto_resize` option to startup and `set_options` default is `True`

### Version 4.9.8
 - Fix potential issue with undo and dictionary copying
 - Fix potential errors when moving/inserting/deleting rows/columns
 - Add drop down position refresh to delete columns/rows on right click menu

### Version 4.9.7
 - Various bug fixes and improvements
 - Add readonly cells/columns/rows

### Version 4.9.6
 - Fix bugs with `font()` functions
 - Fix edit cell bug when hiding columns

### Version 4.9.5
 - Attempt to fix scrolling issues

### Version 4.9.4
 - Make `display_subset_of_columns()` and other names of the same function always sort the showing columns
 - Make right click insert columns left shift data columns along
 - Fix issues with `insert_column()`/`insert_columns()` when hiding columns
 - Add default arguments `mod_column_positions` to functions `insert_column()`/`insert_columns()` which when set to `False` only changes data, not number of showing columns
 - Add `"e"` aka right hand side text alignment for main table, have not added to header or index yet
 - Add functions `align_cells()`, `align_rows()`, `align_columns()`

### Version 4.9.3
 - Fix paste bug
 - Fix mac os vertical scroll code

### Version 4.9.2
 - Add mac OS command c, x, v, z bindings
 - Make shift - mousewheel horizontal scroll

### Version 4.9.1
 - Make alt tab windows maintain open cell edit box
 - Fix bug in `insert_columns()`

### Version 4.8.9 - 4.9.0
 - Make functions `insert_row()`, `insert_column()`, `delete_row()`, `delete_column()` adjust highlighted cells/rows/columns to maintain correct highlight indexes
 - Built in right click functions now also auto-update highlighted cells/rows/columns
 - Add argument `reset_highlights` to `set_sheet_data()` default is `False`
 - Add function `dehighlight_all()`
 - Various bug fixes
 - Right click insert column when hiding columns has been changed to always insert fresh new columns

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
