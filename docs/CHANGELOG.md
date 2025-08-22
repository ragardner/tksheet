### Version 7.5.13
#### Added:
- The standard copy operation uses a CSV writer, which adds quote characters and escape characters when appropriate. Added a new alternate copy option for plain text copying [#311](https://github.com/ragardner/tksheet/issues/311).
- If copy is enabled, the key binding `platform control key` + `platform alt key` + `C`/`c` is now bound to perform a plain text copy.

#### Changed:
- The cell text editor's new line bindings no longer bind both `Alt` and `Option` on macOS and now only bind the appropriate modifier. Other similarly affected bindings have been changed in the same way.

#### Documentation:
- Fixed readme downloads badge.

### Version 7.5.12
#### Fixed:
- `selection_add()` was incorrectly using a `list` and not a `set`.

#### Improved:
- Refactoring to improve performance of the following functions:
    - tree_set_open()
    - tree_open()
    - internal function _tree_open()
    - display_item()
    - scroll_to_item()

### Version 7.5.11
#### Changed:
- Option `user_can_create_notes` now overrides cells made `readonly` but not notes created by the `note()` function that are readonly.

#### Added:
- Option `tooltip_hover_delay` to control how long the mouse cursor must stay still before a tooltip appears.

#### Improved:
- If tooltips/notes are being used hovering over any area of a cell will now make them appear rather than just the cells text.
- A cell edit will now only occur when using tooltips if the tooltip editor's content changed.

#### Fixed:
- Tooltip wrongly opens while cell text editor is open.
- Removed leftover debugging `print()` call from treeview move function.

### Version 7.5.10
#### Fixed:
- Error with notes when using readonly cells and writable notes. [#301](https://github.com/ragardner/tksheet/issues/301).

### Version 7.5.9
#### Fixed:
- Treeview mode regression: when opening (showing) a row using a mouse click with undo enabled the undo stack was mistakenly being added to, this is no longer the case.

#### Improved:
- Treeview mode function `set_children()` now performs all moves in a single action rather than multiple moves.
- Slight performance improvement to treeview mode move functions including `set_children()`.
- Re-implement a basic row selection option for treeview row moving functions.

### Version 7.5.8
#### Fixed:
- Currently selected cell box not showing.
- Potential issues when using `del_rows()` in treeview mode.

#### Improved:
- Slight improvement to performance of moving rows while in treeview mode.

### Version 7.5.7
#### Added:
- Notes and tooltips. Unfortunately this means that `<Enter>` and `<Leave>` have been `tag_bind()` bound to canvas `text` items.

#### Fixed:
- [#304](https://github.com/ragardner/tksheet/issues/304).
- Cells not showing new lines in the middle of text.

#### Removed:
- `selected_rows_to_end_of_window` from Sheet init and options.

#### Changed:
- Some code has been refactored but there shouldn't be any change to behavior as a result.

### Version 7.5.6
#### Addressed:
- [#299](https://github.com/ragardner/tksheet/issues/299).

### Version 7.5.5
#### Fixed:
- Treeview mode `insert()` function and inserting into a hidden section of the tree resulted in `displayed_rows` not being updated and incorrect row visibility in the treeview.
- Span creation using lower case letters and two-tuple keys with a column letter.

#### Improved:
- Performance for treeview mode `insert()` function.

### Version 7.5.4
#### Addressed:
- [#297](https://github.com/ragardner/tksheet/issues/297). More work is likely required but this is an improvement.

#### Changed:
- While treeview mode is enabled and due to issues modifying selection boxes after row moves, selection boxes will no longer be modified after moving items/rows when using treeview mode.

#### Improved:
- Performance improvement for treeview mode function `.index()`.

### Version 7.5.3
#### Fixed:
- Prevent row over-expansion potentially resulting in massive memory use upon paste due to lazy chain evaluation. [#295](https://github.com/ragardner/tksheet/issues/295).
- Regression causing some readonly functions to not work. [#296](https://github.com/ragardner/tksheet/issues/296).
- Wrong font in cell text editor and find window.
- Wrong font in header and index text editors.

### Version 7.5.2
#### Fixed:
- [#294](https://github.com/ragardner/tksheet/issues/294).
- [#296](https://github.com/ragardner/tksheet/issues/296).

#### Removed:
- `MainTable.extra_motion_func` variable as it's no longer required.

### Version 7.5.1
#### Fixed:
- Delete icon using clear icon.
- Wrong size clear icon.
- Popup menu font not scaling with zoom.
- [#293](https://github.com/ragardner/tksheet/issues/293). While tkinter doesn't allow floats for font sizes we can instead round up or down.

#### Added:
- A function for changing popup menu font. It also can be changed using `set_options()`.

### Version 7.5.0
#### Fixed:
- Select all in text editor and find window adding a newline.
- Popup menu font wrongly using table font.

#### Added:
- Capability to add images to built in menu commands and also commands added using `popup_menu_add_command`.
- Image icons to built in menu commands.
- Undo & redo to built in menu.

#### Changed:
- `popup_menu_add_command` will now overwrite existing menu commands with the same label. Previously they were ignored.
- Edit cell menu commands are now disabled when currently selected cell is read only.
- LICENSE file to accomodate open source icons.

#### Improved:
- Dashed selection box outlines that appear on modifications such as Cut and Paste.

### Version 7.4.24
#### Fixed:
- Redo stack was not being purged upon making a sheet modification using many Sheet() methods. [#291](https://github.com/ragardner/tksheet/issues/291).

#### Changed:
- Tab and Return keys to change cell selection no longer stop at the end of the row/column. [#290](https://github.com/ragardner/tksheet/issues/290).
- If nothing is selected in the table and tab key is pressed, select the first cell.

### Version 7.4.23
#### Fixed:
- Paste with nothing selected regression. [#289](https://github.com/ragardner/tksheet/issues/289).

### Version 7.4.22
#### Addressed:
- [#289](https://github.com/ragardner/tksheet/issues/289).

#### Changed:
- When inserting rows and columns any data formatting rules are now pre-applied.

### Version 7.4.21
#### Addressed:
- [#289](https://github.com/ragardner/tksheet/issues/289).

#### Fixed:
- Mouse motion error in top left rectangle (regression).

### Version 7.4.20
#### Addressed:
- [#288](https://github.com/ragardner/tksheet/issues/288).

### Version 7.4.19
#### Addressed:
- [#288](https://github.com/ragardner/tksheet/issues/288).

### Version 7.4.18
#### Addressed:
- [#281](https://github.com/ragardner/tksheet/issues/281).

### Version 7.4.17
#### Fixed:
- [#283](https://github.com/ragardner/tksheet/issues/283).

### Version 7.4.16
#### Fixed:
- [#282](https://github.com/ragardner/tksheet/issues/282).

### Version 7.4.15
#### Fixed:
- [#280](https://github.com/ragardner/tksheet/pull/280).

### Version 7.4.14
#### Fixed:
- Find in selection would sometimes fail.

#### Improved:
- Redraw optimizations.

### Version 7.4.13
#### Fixed:
- 7.4.12 deselection display regression.
- Header and index dropdown boxes with center aligned text made text spill out of the cell a little.

#### Improved:
- Further minor redrawing performance improvements.
- Minor tree build performance improvement.
- Minor performance improvement for moving rows/columns.

#### Addressed:
- Edge case error with moving rows while using treeview mode.

### Version 7.4.12
#### Addressed:
- [#276](https://github.com/ragardner/tksheet/issues/276).
- [#277](https://github.com/ragardner/tksheet/issues/277).
- Page up / down issues.

#### Added:
- html documentation.

### Version 7.4.11
#### Fixed:
- [#275](https://github.com/ragardner/tksheet/issues/275).

### Version 7.4.10
#### Fixed:
- Fixes and improvements to treeview mode.

#### Improved:
- Refactored adding/deleting rows/columns code, behavior should be unchanged.

#### Changed:
- Removed redundant `self.tree` attribute from treeview mode.
- Added `event_data: EventDataDict | None = None` parameter to `set_data()` as the last parameter.
- Added `undo` and `emit_event` parameters to `item()`.

#### Added:
- Treeview function `tree_traverse()`.

### Version 7.4.9
#### Fixed:
- [#274](https://github.com/ragardner/tksheet/issues/274).
- Getting data using spans or `get_data()` with `tdisp=True` and a `"format"` type `Span` would not result in formatted data.
- Incorrect docstring in `get_data()`.

#### Improved:
- Span data getting performance.
- Old data getting functions performance.

### Version 7.4.8
#### Fixed:
- Numerous fixes and improvements for using the find window with hidden rows/columns.
- Error on trying to resize row height when the row is empty.
- Rare minor issues with setting bindings using `set_options()`.
- Some events were adding to the undo stack even if undo was disabled.
- Error on trying to sort rows by column when using the treeview.
- Error on using certain non-treeview row moving functionality while using the treeview.

#### Changed:
- `enable_bindings()` with `"all"` or no args will now also enable the find and replace window.
- If you use `enable_bindings()` with `"all"` or no args and your software has its own Ctrl/Cmd g/G/h/H binding you may need to make changes such as using `set_options()` to remove tksheets find window bindings.
- The find window will no longer match displayed text in the table from checkboxes and dropdown boxes, only underlying cell data values. This is because the text values are not typically editable from the user end.

#### Added:
- Replace and replace all functionality in the find window.
- Using the find window without find within selection in treeview mode will show hidden rows if there's a cell match.
- The find window is now draggable horizontally.
- Ctrl/Cmd h/H binding for toggle replace.
- `**kwargs` to `Sheet()` initialization so that other settings can be changed such as bindings.

#### Improved:
- Refactored `extra_bindings()` code, behavior should be unchanged.
- Find window icons have been improved.

### Version 7.4.7
#### Changed:
- Add `push_n` to tksheet namespace

#### Addressed:
- [#267](https://github.com/ragardner/tksheet/issues/267) as it was not fixed in version `7.4.5`

#### Fixed:
- Bugs with treeview mode `item()` `iid` parameter, used for changing item ids
- `see()` function scrolls to wrong position if the row index auto-resized after scrolling

#### Improved:
- In-built find window performance

### Version 7.4.6
#### Fixed:
- Issues since version `7.4.4` with `get_cell_kwargs()` leading to cell options such as dropdowns sometimes not displaying

#### Changed:
- Remove unused parameter `object` from `Formatter` class

### Version 7.4.5
#### Changed:
- Functions which had mutable default arguments have been changed to either have:
    - no default argument where appropriate OR
    - `None` as the default and create an empty mutable if `None`

#### Fixed:
- [#267](https://github.com/ragardner/tksheet/issues/267)
- Treeview mode: changed recursive functions to `while` functions to avoid recursion depth limit during tree traversal
- Additional treeview function `move()` fixes

#### Improved:
- Code for `see()` and arrow key functions

### Version 7.4.4
#### Changed:
- **!:** Parameter for function `sort()`:
    - `boxes` is now `*box: CreateSpanTypes`
    - [More information](https://github.com/ragardner/tksheet/wiki/Version-7#sorting-cells)

#### Fixed:
- Sorting algorithm wasn't sorting some items correctly
- Boxes in event data not being correct when `sort()` was used
- Addressed some issues with treeview mode `move()`

#### Added:
- New in-built sort keys and default `sort_key` initialization and `set_options()` parameter
    - `natural_sort_key`
    - `version_sort_key`
    - `fast_sort_key`

#### Improved:
- Various minor performance improvements

### Version 7.4.3
#### Fixed:
- Functions bound using `extra_bindings()` were only called when a user performed an action, they're now always called when:
    - adding rows/columns
    - deleting rows/columns

#### Improved:
- Dropdown search fallback matching, performance

#### Addressed
- [#269](https://github.com/ragardner/tksheet/issues/269)

#### Changed:
- `show_top_left` parameter now defaults to `None` to represent tksheet handling of top left visibility
- Provide `natural_sort_key` as importable from `tksheet.natural_sort_key`
- `dropdown_search_function` now uses an iterable of objects rather than an iterable of Sequences of objects
- Rename internal parameter `restored_state` -> `from_undo`

### Version 7.4.2
#### Fixed:
- Strings that started with numbers not being sorted in the correct order

#### Changed:
- Add file paths to sort order using pathlib
- For better backwards compatibility use the usual "move_rows"/"move_columns" event names for sorting rows/columns events

### Version 7.4.1
#### Fixed:
- Issues with `7.4.0` sorting
- [#270](https://github.com/ragardner/tksheet/issues/270)
- Only add to index header when user adds rows/columns if index/header is populated

### Version 7.4.0
#### Changed:
- Text now wraps by character by default, can also disable wrapping or wrap by word
- Significant changes to how text is rendered
- Removed mousewheel scrolling lines in header, replaced with vertical axis table scroll
- Resizing row height to text is now based on the existing column width for the cell/cells, includes double click resizing
- Treeview mode `Node` class now uses `str` for parent and `list[str]` for children attributes
- Function `get_nodes()` renamed -> `get_iids()`
- Removed `data_indexes` parameter from `mapping_move...` functions
- Reduce default treeview indent
- `move()` function now returns the same as other move rows functions
- All `Sheet()` functions with an `undo` parameter have been set to `True` by default
- Using `Sheet.set_data()` or `Span`s to set an individual cell's data as `None` will now do so instead of returning
- Setting `show_top_left` during Sheet() initialization will now make the top left rectangle always show

#### Added:
- Natural sorting functionality [#238](https://github.com/ragardner/tksheet/issues/238)
- Treeview mode now works with all normal tksheet functions, including dragging and dropping rows
- Cell text overflow `allow_cell_overflow: bool = False` to adjacent cells for left and right alignments, disabled by default
- Text wrap for table, header and index
```python
# "" no wrap, "w" word wrap, "c" char wrap
table_wrap: Literal["", "w", "c"] = "c",
index_wrap: Literal["", "w", "c"] = "c",
header_wrap: Literal["", "w", "c"] = "c",
```

- `tree` parameter to `insert_rows()` function, used internally

#### Fixed:
- Index fonts now work correctly
- Functions `column_width()` and `row_height()` work correctly for any parameters
- Down sizing rows/columns when scrolled to the end of the axis would result in a rapid movement of row height/column width
- Address [#269](https://github.com/ragardner/tksheet/issues/269)

#### Improved:
- Minor performance improvements for:
    - `item_displayed()`
    - `show_rows()` / `show_columns()`
    - `move()`
    - Row insertion

### Version 7.3.4
#### Fixed:
- Error on find within selection without a selection
- Some canvas z fighting issues

#### Improved:
- Refactored redrawing code

#### Changed:
- rename Sheet() attr: dropdown_class -> _dropdown_cls

### Version 7.3.3
#### Improved:
- Find within selection performance + memory
- Moving rows/columns performance

### Version 7.3.2
#### Added:
- Built-in find window, use `enable_bindings("find")` to enable
- Escape binding which deselects
- `reverse` parameter to `get_selected_cells()` function

#### Fixed:
- Not redrawing after Control-space, Shift-space bindings for selecting columns/rows
- Not redrawing after Home, Control/Command-Home bindings for selecting start of the row and start of the table
- Dropdown boxes not reseting y scroll
- Issues with arrow key down
- Arrow key up wrongly moving scroll at row 0

#### Changed:
- Renamed `vars.py` -> `constants.py`
- Add `_version.py` to `.gitignore`
- `enable_bindings()` with `"find"` now enables an in-built find window which uses the following bindings:
    - Control/Command-f/F
    - Control/Command-g/G
    - Control/Command-Shift-g/G
    - Escape
    - Return (when the find window has focus)
    - Alt/Option-L to find within selection
- Added Escape binding when cell selection is enabled. Pressing Escape will now close the find window if its open, if its not open it will close any open text editor/dropdown box and deselect all cells.

### Version 7.3.1
#### Changed:
- `deselect()` also closes text editor / dropdown

#### Added:
- Mac OS Zoom in/out bindings
- Control-space, Shift-space bindings for selecting columns/rows if enabled
- Home, Control/Command-Home bindings for selecting start of the row and start of the table

#### Documentation:
- `bulk_insert()`: wrong example
- Docstrings: return values for `insert()`, `bulk_insert()`

#### Fixed:
- Wrong type hinting for `Iterator`

### Version 7.3.0
#### Changed:
- `"end_edit_cell"` events for single cell edits e.g. for cell text editor, dropdown box and checkbox edits now have the value **prior** to the edit in the event under keys `cells.table`
- Variables default values in `Sheet()` init have changed from `"inf"` -> `float("inf")`:
    - `max_column_width`
    - `max_row_height`
    - `max_index_width`
    - `max_header_height`
- Moved variables from `Sheet.MT` -> `Sheet.ops`:
    - `max_column_width`
    - `max_row_height`
    - `max_index_width`
    - `max_header_height`

#### Fixed:
- Lag on resize index width or header height

#### Added:
- Ability to set minimum column width: `min_column_width` in `Sheet()` init and `set_options()`

### Version 7.2.23
#### Changed:
- Edit validation will now also be triggered for undo and redo [256](https://github.com/ragardner/tksheet/issues/256)
- Cell text editors and dropdowns have their own shared background and foreground colors [255](https://github.com/ragardner/tksheet/issues/255)

#### Fixed:
- `7.2.22` regression when using cell text editors and normal dropdown boxes
- `7.2.22` regression `"normal"` dropdown box with `modified_function` not being triggered upon non-key release events such as by right clicking in the text editor and pasting
- Select events will no longer be emitted upon b1 release

#### Added:
- Options for changing table, index and header cell editor bg, fg [255](https://github.com/ragardner/tksheet/issues/255)
    - Added `Sheet()` parameters:
        - `table_editor_bg`
        - `table_editor_fg`
        - `table_editor_select_bg`
        - `table_editor_select_fg`
        - `index_editor_bg`
        - `index_editor_fg`
        - `index_editor_select_bg`
        - `index_editor_select_fg`
        - `header_editor_bg`
        - `header_editor_fg`
        - `header_editor_select_bg`
        - `header_editor_select_fg`

### Version 7.2.22
#### Added:
- Autocomplete for `"normal"` state dropdown boxes
- Python `3.13` support badge
- `"disabled"` state for dropdown boxes
- docstrings for enable/disable bindings and extra bindings functions to help with arguments

#### Fixed:
- Alternate row colors no longer blend with clear single cell selection box
- Highlights and alternate row colors not blending with `show_selected_cells_border=False`
- Tab key with `show_selected_cells_border=False` not updating selected colors when using highlights
- Treeview `selection_set()`/`selection_add()` now only use iids that are in the tree, will not generate an error if an iid is missing
- Error on close text editor

#### Improved:
- Code: minor type hinting
- Code: variable name clarity

### Version 7.2.21
#### Fixed:
- Storing functions in event data using `pickle` causes error [253](https://github.com/ragardner/tksheet/issues/253)

#### Added:
- Basic alternate row colors option

### Version 7.2.20
#### Changed:
- `float_formatter()` by default will no longer accept inputs that end in `"%"`
- `percentage_formatter()` default `format_function` renamed: `to_float()` -> `to_percentage()`
- Events no longer emitted for cut, paste or delete if the sheet has not been edited, for example due to read only cells [249](https://github.com/ragardner/tksheet/issues/249)
- `__init__()` calls now use `super()`

#### Added:
- Options for percentage formatter:
    - `format_function` option `to_percentage()`
    - `format_function` option `alt_to_percentage()`
    - `to_str_function` option `alt_percentage_to_str()`

#### Improved:
- Documentation for data formatting

### Version 7.2.19
#### Fixed:
- Error when pasting into empty sheet
- Potential error if using percentage formatter with more than just `int`, `float` and wrong type hint [250](https://github.com/ragardner/tksheet/issues/250)
- Only show the selection box outline while the mouse is dragged if the control key is down
- `index` and `header` `Span` parameters were not working with function `readonly()`

#### Added:
- Shift + arrowkey bindings for expanding/contracting a selection box
- Functions for getting cell properties (options) [249](https://github.com/ragardner/tksheet/issues/249)
- Ability to edit index in treeview mode, if the binding is enabled

#### Improved:
- Very slight performance boost to treeview insert, inserting rows when rows are hidden

### Version 7.2.18
#### Fixed:
- Inserting rows/columns with hidden rows/columns sometimes resulted in incorrect rows/columns being displayed
- Treeview function `insert()` when using parameter `index` sometimes resulted in treeview items being displayed in the wrong locations

#### Changed:
- Using function `set_currently_selected()` or any function/setter which does the same internally will now trigger a select event like creating selection boxes does
- iids in Treeview mode are now case sensitive
- Treeview function `tree_build()` when parameter `safety` is `False` will no longer check for missing ids in the iid column
- Treeview function `get_children()` now gets item ids from the row index which provides them in the same order as in the treeview
- Add parameter `run_binding` to treeview functions `selection_add()`, `selection_set()`
- Slight color change for top left rectangle bars

#### Added:
- Initialization parameters `default_header` and `default_row_index` can now optionally be set to `None` to not display anything in empty header cells / row index cells
- Parameters `lower` and `include_text_column` to treeview function `tree_build()`
- Treeview function `bulk_insert()`
- Treeview function `get_nodes()` behaves exactly the same as `get_children()` except it retrieves item ids from the tree nodes `dict` not the row index.
- Treeview function `descendants()` which returns a generator
- Treeview property `tree_selected`

### Version 7.2.17
#### Changed:
- Treeview mode slight appearance change

### Version 7.2.16
#### Added:
- Treeview mode documentation

### Version 7.2.15
#### Added:
- New option to address [247](https://github.com/ragardner/tksheet/issues/247)

### Version 7.2.14
#### Added:
- New options to address [246](https://github.com/ragardner/tksheet/issues/246)

#### Fixed:
- Redundant code causing potential redraw error
- Wrong version number in `__init__.py`

### Version 7.2.13
#### Fixed:
- [245](https://github.com/ragardner/tksheet/issues/245)

#### Changed:
- Moved variable `default_index` to sheet options as `default_row_index`
- Moved variable `default_header` to sheet options

#### Improved:
- Documentation

### Version 7.2.12
#### Pull Requests:
- [237](https://github.com/ragardner/tksheet/pull/237)

### Version 7.2.11
#### Fixed:
- [235](https://github.com/ragardner/tksheet/issues/235)
- Error with `show_columns()`

#### Pull Requests:
- [236](https://github.com/ragardner/tksheet/pull/236)

#### Changed:
- Resizing rows/columns with mouse button motion now does so during the motion also

### Version 7.2.10
#### Fixed:
- Fix index width resizing regression
- Sheet no longer steals focus when using `set_sheet_data()` and similar functions, issue [233](https://github.com/ragardner/tksheet/issues/233)

#### Improved:
- Don't show row index resize cursor when unnecessary

### Version 7.2.9
#### Fixed:
- [232](https://github.com/ragardner/tksheet/issues/232)

#### Addressed:
- [231](https://github.com/ragardner/tksheet/issues/231)

#### Changed:
- Dropdown box arrows changed back to lines as polygon triangles had issues with outlines
- Possible slight changes to logic of `display_rows()`/`display_columns()`
- `displayed_rows.setter`/`displayed_columns.setter` now reset row/column positions

### Version 7.2.8
#### Fixed:
- Minor flicker when selecting a row or column
- Top left rectangle sometimes not updating properly on change of dimensions

#### Improved:
- Dropdown triangle visuals

#### Changed:
- Top left rectangle colors
- Dropdown box outline in index/header

### Version 7.2.7
#### Fixed:
- [230](https://github.com/ragardner/tksheet/issues/230)
- Incorrect rows/columns after move with hidden rows/columns

#### Changed:
- Event data for moving rows/columns with hidden rows/columns e.g. under `event.moved.rows.data` now returns not just the moved rows/columns but all new row/column indexes for the displayed rows/columns. e.g. Use `event.moved.rows.displayed` with `Sheet.data_r()`/`Sheet.data_c()` to find only the originally moved rows/columns
- Moving rows/columns with hidden rows/columns will only modify these indexes if `move_data` parameter is `True`, default is `True`

### Version 7.2.6
#### Fixed:
- Drag and drop incorrect drop index
- `set_all_cell_sizes_to_text()` incorrect widths

#### Changed:
- Functions that move rows/columns such as `move_rows()`/`move_columns()` have had their move to indexes changed slightly, you may need to check yours still work as intended if using these functions

### Version 7.2.5
#### Fixed:
- `StopIteration` on drag and drop

#### Added:
- `gen_selected_cells()` function
- `is_contiguous` to `tksheet` namespace

### Version 7.2.4
#### Added:
- Progress bars

#### Fixed:
- Fix resizing cursor appearing at start of header/index and causing issues if clicked

### Version 7.2.3
#### Fixed:
- Fix add columns/add rows not inserting at correct index when index is larger than row or data len
- Fix add columns/add rows not undoing/redoing added heights/widths respectively
- Fix add columns/add rows regression

### Version 7.2.2
#### Added:
- Add functions to address [227](https://github.com/ragardner/tksheet/issues/227):
    - `get_column_text_width`
    - `get_row_text_height`
    - `visible_columns` - `@property`
    - `visible_rows` - `@property`

#### Changed:
- Internal parameter names:
    - `only_set_if_too_small` -> ` only_if_too_small`
- Internal functions:
    - `get_visible_rows`, `get_visible_columns` -> `visible_text_rows`, `visible_text_columns`, also changed to properties `@property`
- Internal function parameters:
    - `set_col_width`
    - `set_row_height`

### Version 7.2.1
#### Fixed:
- Regression since `7.2.0`: selection box borders not showing for rows/columns

#### Changed:
- Slightly reduced minimum row height since `7.2.0` change

### Version 7.2.0
#### Fixed:
- Cells in index/header not having correct colors when columns/rows were selected
- [226](https://github.com/ragardner/tksheet/issues/226)

#### Changed:
- A somewhat major change which warranted a significant version bump - **the minimum row height has increased slightly** to better accodomate the pipe character `|`

### Version 7.1.24
#### Fixed:
- Error on paste into empty sheet with `expand_sheet_if_paste_too_big`
- Data shape not being correct upon inserting a single column or row into an empty sheet
- Potential for `insert_columns()` to cause issues if provided columns are longer than current number of sheet rows
- [225](https://github.com/ragardner/tksheet/issues/225) Insert columns and insert rows not inserting blank cells into header/row index, issue seen with either:
    - Right clicking and using the in-built insert functionality
    - Using the `Sheet()` functions with an `int` for first parameter
- Wrong sheet dimensions caused by a paste which expands the sheet
- Wrong new selection box dimensions after paste which expands the sheet
- `set_options()` with `table_font=` not working

#### Changed:
- `expand_sheet_if_paste_too_big` replaced by `paste_can_expand_x` and `paste_can_expand_y`. To avoid compatibility issues using `expand_sheet_if_paste_too_big` will set both new options

#### Improved:
- Additional protection against issue [224](https://github.com/ragardner/tksheet/issues/224) but with `auto_resize_row_index="empty"` which is the default setting

### Version 7.1.23
#### Fixed:
- Redrawing loop with `auto_resize_row_index=True` [224](https://github.com/ragardner/tksheet/issues/224)

#### Changed:
- Unfortunately in order to fix [224](https://github.com/ragardner/tksheet/issues/224) while using `auto_resize_row_index=True` the x scroll bar will no longer be hidden if it is not needed

### Version 7.1.22
#### Addressed:
- Issue [222](https://github.com/ragardner/tksheet/issues/222)
- Issue [223](https://github.com/ragardner/tksheet/issues/223)

#### Fixed:
- Massive lag when pasting large amounts of data, caused by using Python `csv.Sniffer().sniff()` without setting sample size, now only samples 5000 characters max

#### Changed:
- Added `widget` key to emitted event `dict`s which can be either the header, index or table canvas. If you're using `pickle` on tksheet event `dict`s then you may need to delete the `widget` key beforehand
- Added `Highlight`, `add_highlight`, `new_tk_event` and `get_csv_str_dialect` to tksheet namespace
- Some old functions which used to return `None` now return `self` (`Sheet`)
- Re-add old highlighting functions to docs
- Event data from moving rows/columns while rows/columns are hidden no longer returns all displayed row/column indexes in the `dict` keys `["moved"]["rows"]["data"]`/`["moved"]["columns"]["data"]` but instead returns only the rows/columns which were originally moved.
- Add `Shift-Return` to text editor newline bindings
- Internal `close_text_editor()` function parameters
- Function `close_text_editor()` now closes all text editors that are open

#### Added:
- New parameters to `dropdown()`:
    - `edit_data` to disable editing of data when creating the dropdowns
    - `set_values` so a `dict` of cell/column/row coordinates and values can be provided instead of using `set_value` for every cell in the span
- New emitted events - issue [223](https://github.com/ragardner/tksheet/issues/223):
    - `"<<Copy>>"`
    - `"<<Cut>>"`
    - `"<<Paste>>"`
    - `"<<Delete>>"`
    - `"<<Undo>>"`
    - `"<<Redo>>"`
    - `"<<SelectAll>>"`

#### Improved:
- Dropdown box scrollbars will only display when necessary
- Some themes colors
- Old function `highlight_cells()` slight performance boost when calling it thousands of times
- Added parameters back to the functions `dropdown()`/`checkbox()` for help when writing code

### Version 7.1.21
#### Fixed:
- Error with `equalize_data_row_lengths()` while displaying a row as the header

#### Changed:
- Dropdown and treeview arrow appearance to triangles

### Version 7.1.20
#### Fixed:
- External right click popup menu being overwritten by internal when right clicking in empty table space
- Copying and pasting values from a single column not working correctly
- Mousewheel in header not working on `Linux`
- `get_cell_options()`/`get_row_options()`/`get_column_options()`/`get_index_options()`/`get_header_options()` values not being requested option. These functions now return a `dict` where the values are the requested option, e.g. for `"highlight"` the values will be `namedtuple`s

#### Changed:
- Functions `cut()`/`copy()`/`paste()`/`delete()`/`undo()`/`redo()` now return `EventDataDict` and not `self`/`Sheet`
- `get_cell_options()`/`get_row_options()`/`get_column_options()`/`get_index_options()`/`get_header_options()`, see above fix
- Clipboard operations where there's a single cell which contains newlines will now be clipboarded surrounded by quotechars

#### Added:
- Functions `del_row_positions()`/`del_column_positions()`
- `validation` parameter to `cut()`/`paste()`/`delete()`, `False` disables any bound `edit_validation()` function from running
- `row` and `column` keys to `cut()`/`paste()`/`delete()` events bound with `edit_validation()`

### Version 7.1.12
#### Fixed:
- Incorrect dropdown box opening coordinates following the opening of dropdown boxes on different sized rows
- Calling `grid_forget()` on a parent widget after opening a dropdown box would cause opened dropdown boxes to be hidden
- `disable_bindings()` with right click menu options such as insert_columns would disable the whole right click menu
- Removed old deprecated dropdown box `dict` values

#### Changed:
- To disable the whole right click menu using `disable_bindings()` it must be called with the specific string `"right_click_popup_menu"`
- themes `dict`s now `DotDict`s

### Version 7.1.11
#### Fixed:
- Rare error when selecting header cell with an empty table and `auto_resize_columns` enabled

#### Addressed:
- Issue [221](https://github.com/ragardner/tksheet/issues/221)

### Version 7.1.10
#### Fixed:
- Flickering when scrolling by using mouse to drag scroll bars
- Custom bindings overwritten when using `enable_bindings()`
- Rare situation where header/index is temporarily out of sync with table after setting xview/yview

### Version 7.1.9
#### Fixed:
- Potential error caused by moving rows/columns where:
    - There exists a header but data is empty
    - Data and index/header are all empty but displayed row/column lines exist

#### Changed:
- Function `equalize_data_row_lengths` parameter `include_header` now `True` by default
- Right clicking with right click select enabled in empty space in the table will now deselect all
- Right clicking on a selected area in the main table will no longer select the cell under the pointer if popup menus are not enabled. Use mouse button 1 to select a cell within a selected area

#### Added:
- Functions:
    - `has_focus()`
    - `ctrl_boxes`
    - `canvas_boxes`
    - `full_move_columns_idxs()`
    - `full_move_rows_idxs()`

### Version 7.1.8
#### Fixed:
- Issue with setting sheet xview/yview immediately after initialization
- Visual selection box issue with boxes with 0 length or width

#### Addressed:
- Issue [220](https://github.com/ragardner/tksheet/issues/220)

#### Added:
- Functions:
    - `boxes` setter for use with `Sheet.get_all_selection_boxes_with_types()` or `Sheet.boxes` property
    - `selected` setter
    - `all_rows` property to get `all_rows_displayed` `bool`
    - `all_columns` property to get `all_columns_displayed` `bool`
    - `all_rows` setter to set `all_rows_displayed` `bool`
    - `all_columns` setter to set `all_columns_displayed` `bool`
    - `displayed_rows` setter uses function `Sheet.display_rows()`
    - `displayed_columns` setter uses function `Sheet.display_columns()`

#### Removed:
- All parameters from function `Sheet.after_redraw()`
- Warnings about version changes from `Sheet()` initialization
- `within_range` parameters from internal `get_selected_cells`/`get_selected_rows`/`get_selected_columns` functions

### Version 7.1.7
#### Fixed:
- Using a cell text editor followed by setting a dropdown box with a text editor would set the previously open cell to the dropdown value

#### Changed:
- Selection boxes now have rounded corners
- Function `get_checkbox_points` renamed `rounded_box_coords`

#### Added:
- `rounded_boxes` Sheet option
- WIP `ListBox` class

### Version 7.1.6
#### Fixed:
- Undo error

#### Added:
- Function `deselect_any()` for non specific selection box type deselections

### Version 7.1.5
#### Fixed:
- `set_all_cell_sizes_to_text()` not working correctly if table font is different to index font, resulting in dropdown box values not showing properly
- Dropdown box colors and options not being updated after sheet color change
- Text editor alignments not working

#### Improved:
- Dropdown box alignment now uses cell alignment
- Minor changes to arrow key cell selection

#### Changed:
- When using `show_selected_cells_border=False` the colors for `table_selected_box_cells_fg`/`table_selected_box_rows_fg`/`table_selected_box_columns_fg` will be used for the currently selected cells background

### Version 7.1.4
#### Fixed:
- Fixed shift mouse click select rows/columns selecting cells instead of rows/columns
- Fixed `"<<SheetSelect>>"` event not being emitted after row/column select events

#### Added:
- Add new parameters to `cell_selected()`, `row_selected()`, `column_selected()`, no default behaviour change
- Functions:
    - `event_widget_is_sheet()`
    - `@property` function `boxes`, the same as `get_all_selection_boxes()`
    - `drow()`, `dcol()` functions the same as `displayed_row_to_data()`/`displayed_column_to_data()`

### Version 7.1.3
#### Fixed:
- Error with `get_all_selection_boxes_with_types()`

### Version 7.1.2
#### Fixed:
- Column selected detection bug
- Tagged cells/rows/columns not taken into account in max index detection, relevant for moving columns/rows

### Version 7.1.1
#### Fixed:
- Select all error
- Span widget attribute lost on delete rows/columns and undo
- Tagged cells/rows/columns lost on delete rows/columns and undo

### Version 7.1.0
#### Changed:
- Event data key `"selected"` and function `get_currently_selected()` values have changed:
    - `type_` attribute has been changed from either `"cell"`/`"row"`/`"column"` to `"cells"`/`"rows"`/`"columns"`
    - The attributes in the latter indexes have also changed
    - See the documentation for `get_currently_selected` for more information
---
- Rename class `TextEditor_` to `TextEditorTkText`
- Rename `TextEditor` attribute `textedit` to `tktext`
- Rename `namedtuple` `CurrentlySelectedClass` to `Selected`
---
- Overhaul how selection boxes are handled internally. `Sheet` functions dealing with selection boxes should behave the same
- Changed order of `Sheet()` init parameters
---
- `auto_resize_row_index` now has a different default value for its old behaviour:
    - `auto_resize_row_index: bool | Literal["empty"] = "empty"`
        - With `"empty"` it will only automatically resize if the row index is empty
        - With `True` it will always automatically resize
        - `False` it will never automatically resize
---
- Scrollbar appearance
- `hide_rows()`/`hide_columns()` functions now endeavour to save the row heights/column widths so that they may be reinserted when using new functions `show_rows()`/`show_columns()`
- Internal Dropdown Box information `dict`s no longer have the keys `"window"` and `"canvas_id"`
---
Span objects now have an additional two functions which link to the `Sheet` functions of the same names:
- `span.tag()`
- `span.untag()`

#### Removed:
- **Parameters**:
    - `set_text_editor_value()` parameters `r` and `c`

#### Added:
- **Functions**:
    - `show_rows()`, `show_columns()` which are designed to work alongside their `hide_rows()`/`hide_columns()` counterparts
    - `set_index_text_editor_value()` and `set_header_text_editor_value()`
    - `xview()`, `yview()`, `xview_moveto()`, `yview_moveto()`
- **Parameters**:
    - `data_indexes` `bool` parameters to functions: `hide_rows`, `hide_columns`, default value is `False` meaning there is no behavior change
    - `create_selections` `bool` parameters to functions: `insert_rows`, `insert_columns` default value is `True` meaning there is no behavior change
- **New tksheet functionality**:
    - Treeview mode (still a work in progress - functions are inside `sheet.py` under # Treeview Mode)
    - Cell, row and column tagging functions, also added to `Span`s
    - Ability to change the appearance of both scroll bars
    - New binding `"<<SheetSelect>>"` which encompasses all select events

#### Fixed:
- `mapping_move_rows()` error
- Potential issue with using `insert_rows` while also using an `int` as the row index to display a specific column in the index
- Potential error if a selection box ends up outside of rows/columns
- Pull request [#214](https://github.com/ragardner/tksheet/pull/214)
- Issue [215](https://github.com/ragardner/tksheet/issues/215)

#### Improved:
- Ctrl select now allows overlapping boxes which begin from within another box
- Ctrl click deselection
- The currently selected cell will no longer change after edits to individual cells in the main table which are not valid with a different value

### Version 7.0.6
#### Changed:
- The following `MainTable` attributes are now simply `int`s or `str`s which represent either pixels or number of lines, instead of `tuple`s:
    - `default_header_height`
    - `default_row_height`
- `Sheet()` init keyword argument `default_row_index_width` now only accepts `int`s
- Simplify internal use of `default_header_height`, `default_row_height`, `default_column_width`, `default_row_index_width`
- Move the following attribute locations from `MainTable` to `Sheet.ops`:
    - `default_header_height`
    - `default_row_height`
    - `default_column_width`
    - `default_row_index_width`
- Removed some protections for setting default row heights, default column widths smaller than minimum heights/widths

#### Added:
- Functions to address issue [#212](https://github.com/ragardner/tksheet/issues/212):
    - `get_text_editor_value`
    - `close_text_editor`

### Version 7.0.5
#### Fixed:
- Issue #210

### Version 7.0.4
#### Fixed:
- Additional header cells being created when using `set_data()` or data setting using spans under certain circumstances
- Additional index cells being created when using `set_data()` or data setting using spans under certain circumstances

#### Added:
- `Sheet.reset()` function

### Version 7.0.3
#### Fixed:
- Some classifiers in `pyproject.toml`

#### Added:
- Parameter `emit_event` to the functions:
    - `span`
    - `set_data`
    - `clear`
    - `insert_row`
    - `insert_column`
    - `insert_rows`
    - `insert_columns`
    - `del_row`
    - `del_column`
    - `del_rows`
    - `del_columns`
    - `move_rows`
    - `move_columns`
    - `mapping_move_rows`
    - `mapping_move_columns`

- Attribute `emit_event` to Span objects, default value is `False`
- Functions `mapping_move_rows` and `mapping_move_columns` to docs

#### Changed:
- Functions renamed:
    - `move_rows_using_mapping` -> `mapping_move_rows`
    - `move_columns_using_mapping` -> `mapping_move_columns`
- Order of parameters for functions:
    - `mapping_move_rows`
    - `mapping_move_columns`

### Version 7.0.2
#### Changed:
- `Sheet()` initialization parameter `row_index_width` renamed `default_row_index_width`
- `set_options()` keyword argument `row_index_width` renamed `default_row_index_width`
- Move doc files to new docs folder
- Delete `version.py` file, move `__version__` variable to `__init__.py`
- Add backwards compatibility for `Sheet()` initialization parameters:
    - `column_width`
    - `header_height`
    - `row_height`
    - `row_index_width`

### Version 7.0.1
#### Removed:
- Function `cell_edit_binding()` use `enable_bindings()` instead
- Function `edit_bindings()` use `enable_bindings()` instead

#### Fixed:
- Error when closing a dropdown in the row index
- Menus modified for every binding for one call of enable/disable bindings with more than one binding
- Editing the index/header cells being allowed when using `enable_bindings("all")`/`enable_bindings()` when it's not supposed to

#### Changed:
- `Sheet()` initialization parameter `column_width` renamed `default_column_width`
- `set_options()` keyword argument `column_width` renamed `default_column_width`
- `Sheet()` initialization parameter `header_height` renamed `default_header_height`
- `set_options()` keyword argument `header_height` renamed `default_header_height`
- `Sheet()` initialization parameter `row_height` renamed `default_row_height`
- `set_options()` keyword argument `row_height` renamed `default_row_height`
- `Sheet()` initialization parameter `auto_resize_default_row_index` renamed `auto_resize_row_index`
- `set_options()` keyword argument `auto_resize_default_row_index` renamed `auto_resize_row_index`
- `Sheet()` initialization parameter `enable_edit_cell_auto_resize` renamed `cell_auto_resize_enabled`
- `set_options()` keyword argument `enable_edit_cell_auto_resize` renamed `cell_auto_resize_enabled`
- Moved most changable sheet attributes to a dict variable named `ops` (which has dot notation access) accessible from the `Sheet` object
- The order of parameters has been changed for `Sheet()` initialization
- `Sheet` attribute `self.C` renamed `self.PAR`
- Various widget attributes named `self.parentframe` have been renamed `self.PAR`
- Return key on cell editor will move to the next cell regardless of whether the cell was edited
- Parameters for classes `TextEditor`, `TextEditor_`, `Dropdown
- Bindings for cut, copy, paste, delete, undo, redo no longer use both `Command` and `Control` but either one depending on the users operating system
- Internal arrow key binding methodology, if you were previously using this to change the arrow key bindings see [here](https://github.com/ragardner/tksheet/wiki/Version-7#changing-key-bindings) for info on adding arrow key bindings

#### Added:
- A way to change the in-built popup menu labels, info [here](https://github.com/ragardner/tksheet/wiki/Version-7#sheet-languages-and-bindings)
- A way to change the in-built bindings for cut, copy, paste, delete, undo, redo, select all, all the arrowkey bindings and page up/down, info [here](https://github.com/ragardner/tksheet/wiki/Version-7#changing-key-bindings)

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
