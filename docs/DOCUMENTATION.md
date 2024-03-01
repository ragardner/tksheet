# Table of Contents

- [About tksheet](https://github.com/ragardner/tksheet/wiki/Version-7#about-tksheet)
- [Installation and Requirements](https://github.com/ragardner/tksheet/wiki/Version-7#installation-and-requirements)
- [Basic Initialization](https://github.com/ragardner/tksheet/wiki/Version-7#basic-initialization)
- [Usage Examples](https://github.com/ragardner/tksheet/wiki/Version-7#usage-examples)
- [Initialization Options](https://github.com/ragardner/tksheet/wiki/Version-7#initialization-options)
---
- [Sheet Colors](https://github.com/ragardner/tksheet/wiki/Version-7#sheet-colors)
- [Header and Index](https://github.com/ragardner/tksheet/wiki/Version-7#header-and-index)
---
- [Bindings and Functionality](https://github.com/ragardner/tksheet/wiki/Version-7#bindings-and-functionality)
- [tkinter and tksheet Events](https://github.com/ragardner/tksheet/wiki/Version-7#tkinter-and-tksheet-events)
- [Sheet Languages and Bindings](https://github.com/ragardner/tksheet/wiki/Version-7#sheet-languages-and-bindings)
---
- [Span Objects](https://github.com/ragardner/tksheet/wiki/Version-7#span-objects)
- [Named Spans](https://github.com/ragardner/tksheet/wiki/Version-7#named-spans)
---
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
- [Scroll Positions and Cell Visibility](https://github.com/ragardner/tksheet/wiki/Version-7#scroll-positions-and-cell-visibility)
---
- [Hiding Columns](https://github.com/ragardner/tksheet/wiki/Version-7#hiding-columns)
- [Hiding Rows](https://github.com/ragardner/tksheet/wiki/Version-7#hiding-rows)
- [Hiding Sheet Elements](https://github.com/ragardner/tksheet/wiki/Version-7#hiding-sheet-elements)
- [Sheet Height and Width](https://github.com/ragardner/tksheet/wiki/Version-7#sheet-height-and-width)
---
- [Cell Text Editor](https://github.com/ragardner/tksheet/wiki/Version-7#cell-text-editor)
- [Sheet Options and Other Functions](https://github.com/ragardner/tksheet/wiki/Version-7#sheet-options-and-other-functions)
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

- `tksheet` is a Python tkinter table widget written in pure python.
- It is licensed under the [MIT license](https://github.com/ragardner/tksheet/blob/master/LICENSE.txt).
- It works by using tkinter canvases and moving lines, text and rectangles around for only the visible portion of the table.
- If you are using a version of tksheet that is older than `7.0.0` then you will need the documentation [here](https://github.com/ragardner/tksheet/wiki/Version-6) instead.
    - In tksheet versions >= `7.0.2` the current version will be at the top of the file `__init__.py`.

### **Limitations**

Some examples of things that are not possible with tksheet:
- Cell merging
- Cell text wrap
- Changing font for individual cells
- Different fonts for index and table
- Mouse drag copy cells
- Cell highlight borders
- At the present time the type hinting in tksheet is only meant to serve as a guide and not to be used with type checkers.

---
# **Installation and Requirements**

`tksheet` is available through PyPi (Python package index) and can be installed by using Pip through the command line `pip install tksheet`

```python
#To install using pip
pip install tksheet

#To update using pip
pip install tksheet --upgrade
```

Alternatively you can download the source code and inside the tksheet directory where the `pyproject.toml` file is located use the command line `pip install -e .`

- Versions **<** `7.0.0` require a Python version of **`3.7`** or higher.
- Versions **>=** `7.0.0` require a Python version of **`3.8`** or higher.

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
# **Usage Examples**

This is to demonstrate some of tksheets functionality.
- The functions which return the Sheet itself (have `-> Sheet`) can be chained with other Sheet functions.
- The functions which return a Span (have `-> Span`) can be chained with other Span functions.

```python
from tksheet import Sheet, num2alpha
import tkinter as tk


class demo(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.frame = tk.Frame(self)
        self.frame.grid_columnconfigure(0, weight=1)
        self.frame.grid_rowconfigure(0, weight=1)

        # create an instance of Sheet()
        self.sheet = Sheet(
            # set the Sheets parent widget
            self.frame,
            # optional: set the Sheets data at initialization
            data=[[f"Row {r}, Column {c}\nnewline1\nnewline2" for c in range(20)] for r in range(100)],
            theme="light green",
            height=520,
            width=1000,
        )
        # enable various bindings
        self.sheet.enable_bindings("all", "edit_index", "edit_header")

        # set a user edit validation function
        # AND bind all sheet modification events to a function
        # chained as two functions
        # more information at:
        # https://github.com/ragardner/tksheet/wiki/Version-7#validate-user-cell-edits
        self.sheet.edit_validation(self.validate_edits).bind("<<SheetModified>>", self.sheet_modified)

        # add some new commands to the in-built right click menu
        # setting data
        self.sheet.popup_menu_add_command(
            "Say Hello",
            self.say_hello,
            index_menu=False,
            header_menu=False,
            empty_space_menu=False,
        )
        # getting data
        self.sheet.popup_menu_add_command(
            "Print some data",
            self.print_data,
            empty_space_menu=False,
        )
        # overwrite Sheet data
        self.sheet.popup_menu_add_command("Reset Sheet data", self.reset)
        # set the header
        self.sheet.popup_menu_add_command(
            "Set header data",
            self.set_header,
            table_menu=False,
            index_menu=False,
            empty_space_menu=False,
        )
        # set the index
        self.sheet.popup_menu_add_command(
            "Set index data",
            self.set_index,
            table_menu=False,
            header_menu=False,
            empty_space_menu=False,
        )

        self.frame.grid(row=0, column=0, sticky="nswe")
        self.sheet.grid(row=0, column=0, sticky="nswe")

    def validate_edits(self, event):
        # print (event)
        if event.eventname.endswith("header"):
            return event.value + " edited header"
        elif event.eventname.endswith("index"):
            return event.value + " edited index"
        else:
            if not event.value:
                return "EMPTY"
            return event.value[:3]

    def say_hello(self):
        current_selection = self.sheet.get_currently_selected()
        if current_selection:
            box = (current_selection.row, current_selection.column)
            # set cell data, end user Undo enabled
            # more information at:
            # https://github.com/ragardner/tksheet/wiki/Version-7#setting-sheet-data
            self.sheet[box].options(undo=True).data = "Hello World!"
            # highlight the cell for 2 seconds
            self.highlight_area(box)

    def print_data(self):
        for box in self.sheet.get_all_selection_boxes():
            # get user selected area sheet data
            # more information at:
            # https://github.com/ragardner/tksheet/wiki/Version-7#getting-sheet-data
            data = self.sheet[box].data
            for row in data:
                print(row)

    def reset(self):
        # overwrites sheet data, more information at:
        # https://github.com/ragardner/tksheet/wiki/Version-7#setting-sheet-data
        self.sheet.set_sheet_data([[f"Row {r}, Column {c}\nnewline1\nnewline2" for c in range(20)] for r in range(100)])
        # reset header and index
        self.sheet.headers([])
        self.sheet.index([])

    def set_header(self):
        self.sheet.headers(
            [f"Header {(letter := num2alpha(i))} - {i + 1}\nHeader {letter} 2nd line!" for i in range(20)]
        )

    def set_index(self):
        self.sheet.set_index_width()
        self.sheet.row_index(
            [f"Index {(letter := num2alpha(i))} - {i + 1}\nIndex {letter} 2nd line!" for i in range(100)]
        )

    def sheet_modified(self, event):
        # uncomment below if you want to take a look at the event object
        # print ("The sheet was modified! Event object:")
        # for k, v in event.items():
        #     print (k, ":", v)
        # print ("\n")

        # otherwise more information at:
        # https://github.com/ragardner/tksheet/wiki/Version-7#event-data

        # highlight the modified cells briefly
        if event.eventname.startswith("move"):
            for box in self.sheet.get_all_selection_boxes():
                self.highlight_area(box)
        else:
            for box in event.selection_boxes:
                self.highlight_area(box)

    def highlight_area(self, box, time=800):
        # highlighting an area of the sheet
        # more information at:
        # https://github.com/ragardner/tksheet/wiki/Version-7#highlighting-cells
        self.sheet[box].bg = "indianred1"
        self.after(time, lambda: self.clear_highlight(box))

    def clear_highlight(self, box):
        self.sheet[box].dehighlight()


app = demo()
app.mainloop()
```

---
# **Initialization Options**

These are all the initialization parameters, the only required argument is the sheets `parent`, every other parameter has default arguments.

```python
def __init__(
    parent: tk.Misc,
    name: str = "!sheet",
    show_table: bool = True,
    show_top_left: bool = True,
    show_row_index: bool = True,
    show_header: bool = True,
    show_x_scrollbar: bool = True,
    show_y_scrollbar: bool = True,
    width: int | None = None,
    height: int | None = None,
    headers: None | list[object] = None,
    header: None | list[object] = None,
    row_index: None | list[object] = None,
    index: None | list[object] = None,
    default_header: Literal["letters", "numbers", "both"] = "letters",
    default_row_index: Literal["letters", "numbers", "both"] = "numbers",
    data_reference: None | Sequence[Sequence[object]] = None,
    data: None | Sequence[Sequence[object]] = None,
    startup_select: tuple[int, int, str] | tuple[int, int, int, int, str] = None,
    startup_focus: bool = True,
    total_columns: int | None = None,
    total_rows: int | None = None,
    default_column_width: int = 120,
    default_header_height: str | int = "1",
    default_row_index_width: int = 70,
    default_row_height: str | int = "1",
    max_column_width: Literal["inf"] | float = "inf",
    max_row_height: Literal["inf"] | float = "inf",
    max_header_height: Literal["inf"] | float = "inf",
    max_index_width: Literal["inf"] | float = "inf",
    after_redraw_time_ms: int = 20,
    set_all_heights_and_widths: bool = False,
    zoom: int = 100,
    align: str = "w",
    header_align: str = "center",
    row_index_align: str | None = None,
    index_align: str = "center",
    displayed_columns: list[int] = [],
    all_columns_displayed: bool = True,
    displayed_rows: list[int] = [],
    all_rows_displayed: bool = True,
    outline_thickness: int = 0,
    outline_color: str = theme_light_blue["outline_color"],
    theme: str = "light blue",
    frame_bg: str = theme_light_blue["table_bg"],
    popup_menu_fg: str = theme_light_blue["popup_menu_fg"],
    popup_menu_bg: str = theme_light_blue["popup_menu_bg"],
    popup_menu_highlight_bg: str = theme_light_blue["popup_menu_highlight_bg"],
    popup_menu_highlight_fg: str = theme_light_blue["popup_menu_highlight_fg"],
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
    index_hidden_rows_expander_bg: str = theme_light_blue["index_hidden_rows_expander_bg"],
    header_bg: str = theme_light_blue["header_bg"],
    header_border_fg: str = theme_light_blue["header_border_fg"],
    header_grid_fg: str = theme_light_blue["header_grid_fg"],
    header_fg: str = theme_light_blue["header_fg"],
    header_selected_cells_bg: str = theme_light_blue["header_selected_cells_bg"],
    header_selected_cells_fg: str = theme_light_blue["header_selected_cells_fg"],
    header_selected_columns_bg: str = theme_light_blue["header_selected_columns_bg"],
    header_selected_columns_fg: str = theme_light_blue["header_selected_columns_fg"],
    header_hidden_columns_expander_bg: str = theme_light_blue["header_hidden_columns_expander_bg"],
    top_left_bg: str = theme_light_blue["top_left_bg"],
    top_left_fg: str = theme_light_blue["top_left_fg"],
    top_left_fg_highlight: str = theme_light_blue["top_left_fg_highlight"],
    to_clipboard_delimiter: str = "\t",
    to_clipboard_quotechar: str = '"',
    to_clipboard_lineterminator: str = "\n",
    from_clipboard_delimiters: list[str] | str = ["\t"],
    show_default_header_for_empty: bool = True,
    show_default_index_for_empty: bool = True,
    page_up_down_select_row: bool = True,
    expand_sheet_if_paste_too_big: bool = False,
    paste_insert_column_limit: int | None = None,
    paste_insert_row_limit: int | None = None,
    show_dropdown_borders: bool = False,
    arrow_key_down_right_scroll_page: bool = False,
    cell_auto_resize_enabled: bool = True,
    auto_resize_row_index: bool = True,
    auto_resize_columns: int | None = None,
    auto_resize_rows: int | None = None,
    set_cell_sizes_on_zoom: bool = False,
    font: tuple[str, int, str] = FontTuple(
        "Calibri",
        13 if USER_OS == "darwin" else 11,
        "normal",
    ),
    header_font: tuple[str, int, str] = FontTuple(
        "Calibri",
        13 if USER_OS == "darwin" else 11,
        "normal",
    ),
    index_font: tuple[str, int, str] = FontTuple(
        "Calibri",
        13 if USER_OS == "darwin" else 11,
        "normal",
    ),  # currently has no effect
    popup_menu_font: tuple[str, int, str] = FontTuple(
        "Calibri",
        13 if USER_OS == "darwin" else 11,
        "normal",
    ),
    max_undos: int = 30,
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
) -> None
```
- `name` setting a name for the sheet is useful when you have multiple sheets and you need to determine which one an event came from.
- `auto_resize_columns` (`int`, `None`) if set as an `int` the columns will automatically resize to fit the width of the window, the `int` value being the minimum of each column in pixels.
- `auto_resize_rows` (`int`, `None`) if set as an `int` the rows will automatically resize to fit the height of the window, the `int` value being the minimum height of each row in pixels.
- `startup_select` selects cells, rows or columns at initialization by using a `tuple` e.g. `(0, 0, "cells")` for cell A0 or `(0, 5, "rows")` for rows 0 to 5.
- `data_reference` and `data` are essentially the same.
- `row_index` and `index` are the same, `index` takes priority, same as with `headers` and `header`.
- `startup_select` either `(start row, end row, "rows")`, `(start column, end column, "rows")` or `(start row, start column, end row, end column, "cells")`. The start/end row/column variables need to be `int`s.

You can change these settings after initialization using the [`set_options()` function](https://github.com/ragardner/tksheet/wiki/Version-7#sheet-options-and-other-functions).

---
# **Sheet Colors**

To change the colors of individual cells, rows or columns use the functions listed under [highlighting cells](https://github.com/ragardner/tksheet/wiki/Version-7#highlighting-cells).

For the colors of specific parts of the table such as gridlines and backgrounds use the function `set_options()`, keyword arguments specific to sheet colors are listed below. All the other `set_options()` arguments can be found [here](https://github.com/ragardner/tksheet/wiki/Version-7#sheet-options-and-other-functions).

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
change_theme(theme: str = "light blue", redraw: bool = True) -> Sheet
```
- `theme` (`str`) options (themes) are currently `light blue`, `light green`, `dark`, `dark blue` and `dark green`.

---
# **Header and Index**

#### **Set the header**

```python
set_header_data(value: object, c: int | None | Iterator = None, redraw: bool = True) -> Sheet
```
- `value` (`iterable`, `int`, `Any`) if `c` is left as `None` then it attempts to set the whole header as the `value` (converting a generator to a list). If `value` is `int` it sets the header to display the row with that position.
- `c` (`iterable`, `int`, `None`) if both `value` and `c` are iterables it assumes `c` is an iterable of positions and `value` is an iterable of values and attempts to set each value to each position. If `c` is `int` it attempts to set the value at that position.

```python
headers(
    newheaders: object = None,
    index: None | int = None,
    reset_col_positions: bool = False,
    show_headers_if_not_sheet: bool = True,
    redraw: bool = True,
) -> object
```
- Using an integer `int` for argument `newheaders` makes the sheet use that row as a header e.g. `headers(0)` means the first row will be used as a header (the first row will not be hidden in the sheet though), this is sort of equivalent to freezing the row.
- Leaving `newheaders` as `None` and using the `index` argument returns the existing header value in that index.
- Leaving all arguments as default e.g. `headers()` returns existing headers.

___

#### **Set the index**

```python
set_index_data(value: object, r: int | None | Iterator = None, redraw: bool = True) -> Sheet
```
- `value` (`iterable`, `int`, `Any`) if `r` is left as `None` then it attempts to set the whole index as the `value` (converting a generator to a list). If `value` is `int` it sets the index to display the row with that position.
- `r` (`iterable`, `int`, `None`) if both `value` and `r` are iterables it assumes `r` is an iterable of positions and `value` is an iterable of values and attempts to set each value to each position. If `r` is `int` it attempts to set the value at that position.

```python
row_index(
    newindex: object = None,
    index: None | int = None,
    reset_row_positions: bool = False,
    show_index_if_not_sheet: bool = True,
    redraw: bool = True,
) -> object
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
	- `"all"`
	- `"single_select"`
	- `"toggle_select"`
	- `"drag_select"`
       - `"select_all"`
	- `"column_drag_and_drop"` / `"move_columns"`
	- `"row_drag_and_drop"` / `"move_rows"`
	- `"column_select"`
	- `"row_select"`
	- `"column_width_resize"`
	- `"double_click_column_resize"`
	- `"row_width_resize"`
	- `"column_height_resize"`
	- `"arrowkeys"` # all arrowkeys including page up and down
    - `"up"`
    - `"down"`
    - `"left"`
    - `"right"`
    - `"prior"` # page up
    - `"next"` # page down
	- `"row_height_resize"`
	- `"double_click_row_resize"`
	- `"right_click_popup_menu"`
	- `"rc_select"`
	- `"rc_insert_column"`
	- `"rc_delete_column"`
	- `"rc_insert_row"`
	- `"rc_delete_row"`
    - `"ctrl_click_select"` / `"ctrl_select"`
	- `"copy"`
	- `"cut"`
	- `"paste"`
	- `"delete"`
	- `"undo"`
	- `"edit_cell"`
    - `"edit_header"`
    - `"edit_index"`

Notes:
- You can change the Sheets key bindings for functionality such as copy, paste, up, down etc. Instructions can be found [here](https://github.com/ragardner/tksheet/wiki/Version-7#changing-key-bindings).
- Control selection is **NOT** enabled with `"all"` and has to be specifically enabled.
- Header cell editing is **NOT** enabled with `"all"` and has to be specifically enabled.
- Index cell editing is **NOT** enabled with `"all"` and has to be specifically enabled.
- To allow table expansion when pasting data which doesn't fit in the table use either:
   - `expand_sheet_if_paste_too_big = True` in sheet initialization arguments or
   - `sheet.set_options(expand_sheet_if_paste_too_big = True)`

Example:
- `sheet.enable_bindings()` to enable absolutely everything.

___

#### **Disable table functionality and bindings**

```python
disable_bindings(*bindings)
```
Notes:
- Uses the same arguments as `enable_bindings()`.

___

#### **Bind specific table functionality**

This function allows you to bind very specific table functionality to your own functions.
- If you want less specificity in event names you can also bind all sheet modifying events to a single function, [see here](https://github.com/ragardner/tksheet/wiki/Version-7#tkinter-and-tksheet-events).
- If you want to validate/modify user cell edits [see here](https://github.com/ragardner/tksheet/wiki/Version-7#validate-user-cell-edits).

```python
extra_bindings(
    bindings: str | list | tuple,
    func: Callable | None = None,
) -> Sheet
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
- For `"end_..."` events the bound function is run before the value is set.
- **To unbind** a function either set `func` argument to `None` or leave it as default e.g. `extra_bindings("begin_copy")` to unbind `"begin_copy"`.

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
- Using one of the following `"all_modified_events"`, `"sheetmodified"`, `"sheet_modified"`, `"modified_events"`, `"modified"` will make any insert, delete or cell edit including pastes and undos send an event to your function.
- For events `"begin_move_columns"`/`"begin_move_rows"` the point where columns/rows will be moved to will be accessible by the key named `"value"`.
- For `"begin_edit..."` events the bound function must return a value to open the cell editor with, example [here](https://github.com/ragardner/tksheet/wiki/Version-7#example-custom-right-click-and-text-editor-validation).

#### **Event Data**

Using `extra_bindings()` the function you bind needs to have at least one argument which will receive a `dict`. The values of which can be accessed by dot notation e.g. `event.eventname` or `event.cells.table`:

```python
for (row, column), old_value in event.cells.table.items():
    print (f"R{row}", f"C{column}", "Old Value:", old_value)
```

It has the following layout and keys:

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
    },
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
- Key **`["cells"]["index"]`** if any index cells have been modified by cell editors, dropdown boxes, check boxes, undo or redo this will be a `dict` with keys of `int: data row index` and the values will be the cell values at that location **prior** to the change. The `dict` will be empty if no such changes have taken place.
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
- Key **`["selected"]`** the value of this when there is something selected on the sheet is a `namedtuple` which contains the values: `(row: int, column: int, type_: str (either "cell", "row" or "column"), tags: tuple)`.
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

#### **Validate user cell edits**

With this function you can validate (modify) most user sheet edits, includes cut, paste, delete (including column/row clear), dropdown boxes and cell edits.
```python
edit_validation(func: Callable | None = None) -> Sheet
```
Parameters:
- `func` (`Callable`, `None`) must either be a function which will receive a tksheet event dict which looks like [this](https://github.com/ragardner/tksheet/wiki/Version-7#event-data) or `None` which unbinds the function.

Notes:
- For examples of this function see [here](https://github.com/ragardner/tksheet/wiki/Version-7#basic-use) and [here](https://github.com/ragardner/tksheet/wiki/Version-7#example-custom-right-click-and-text-editor-validation).

___

#### **Add commands to the in-built right click popup menu**

```python
popup_menu_add_command(
    label: str,
    func: Callable,
    table_menu: bool = True,
    index_menu: bool = True,
    header_menu: bool = True,
    empty_space_menu: bool = True,
) -> Sheet
```

___

#### **Remove commands added using popup_menu_add_command from the in-built right click popup menu**

```python
popup_menu_del_command(label: str | None = None) -> Sheet
```
- If `label` is `None` then it removes all.

___

#### **Enable or disable mousewheel, left click etc**

```python
basic_bindings(enable: bool = False) -> Sheet
```

___

These functions are links to the Sheets own functionality. Functions such as `cut()` rely on whatever is currently selected on the Sheet.
```python
cut(event: object = None) -> Sheet
copy(event: object = None) -> Sheet
paste(event: object = None) -> Sheet
delete(event: object = None) -> Sheet
undo(event: object = None) -> Sheet
redo(event: object = None) -> Sheet
```

___

#### **Get the last event data dict**

```python
@property
def event() -> EventDataDict
```
- e.g. `last_event_data = sheet.event`
- Will be empty `EventDataDict` if there is no last event.

___

#### **Set focus to the Sheet**

```python
focus_set(
    canvas: Literal[
        "table",
        "header",
        "row_index",
        "index",
        "topleft",
        "top_left",
    ] = "table",
) -> Sheet
```

---
# **tkinter and tksheet Events**

- With the `Sheet.bind()` function you can bind things in the usual way you would in tkinter and they will bind to all the `tksheet` canvases.
- There are also two special `tksheet` events you can bind, `"<<SheetModified>>"` and `"<<SheetRedrawn>>"`.

```python
bind(
    event: str,
    func: Callable,
    add: str | None = None,
)
```
Parameters:
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
unbind(binding: str) -> Sheet
```

---
# **Sheet Languages and Bindings**

Listed in this section are ways to change some of tksheets language:
- The in-built right click menu.
- The in-built functionality keybindings, such as copy, paste etc.

Unfortunately these are currently the only modifications to tksheets language that are possible.

#### **Changing right click menu labels**

You can change the labels for tksheets in-built right click popup menu by using the [`set_options()` function](https://github.com/ragardner/tksheet/wiki/Version-7#sheet-options-and-other-functions) with any of the following keyword arguments:

```python
edit_header_label
edit_header_accelerator
edit_index_label
edit_index_accelerator
edit_cell_label
edit_cell_accelerator
cut_label
cut_accelerator
cut_contents_label
cut_contents_accelerator
copy_label
copy_accelerator
copy_contents_label
copy_contents_accelerator
paste_label
paste_accelerator
delete_label
delete_accelerator
clear_contents_label
clear_contents_accelerator
delete_columns_label
delete_columns_accelerator
insert_columns_left_label
insert_columns_left_accelerator
insert_column_label
insert_column_accelerator
insert_columns_right_label
insert_columns_right_accelerator
delete_rows_label
delete_rows_accelerator
insert_rows_above_label
insert_rows_above_accelerator
insert_rows_below_label
insert_rows_below_accelerator
insert_row_label
insert_row_accelerator
select_all_label
select_all_accelerator
undo_label
undo_accelerator
```

Example:

```python
# changing the copy label to the spanish for Copy
sheet.set_options(copy_label="Copiar")
```

#### **Changing key bindings**

You can change the bindings for tksheets in-built functionality such as cut, copy, paste by using the [`set_options()` function](https://github.com/ragardner/tksheet/wiki/Version-7#sheet-options-and-other-functions) with any the following keyword arguments:

```python
copy_bindings
cut_bindings
paste_bindings
undo_bindings
redo_bindings
delete_bindings
select_all_bindings
tab_bindings
up_bindings
right_bindings
down_bindings
left_bindings
prior_bindings
next_bindings
```

The argument must be a `list` of **tkinter** binding `str`s. In the below example the binding for copy is changed to `"<Control-e>"` and `"<Control-E>"`.

```python
# changing the binding for copy
sheet.set_options(copy_bindings=["<Control-e>", "<Control-E>"])
```

The default values for these bindings can be found in the tksheet file `sheet_options.py`.

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
    emit_event: bool = False,
    widget: object = None,
    expand: None | str = None,
    formatter_options: dict | None = None,
    **kwargs,
) -> Span
"""
Create a span / get an existing span by name
Returns the created span
"""
```
Parameters:
- `key` you do not have to provide an argument for `key`, if no argument is provided then the span will be a full sheet span. Otherwise `key` can be the following types which are type hinted as `CreateSpanTypes`:
    - `None`
    - `str` e.g. `sheet.span("A1:F1")`
    - `int` e.g. `sheet.span(0)`
    - `slice` e.g. `sheet.span(slice(0, 4))`
    - `Sequence[int | None, int | None]` representing a cell of `row, column` e.g. `sheet.span(0, 0)`
    - `Sequence[Sequence[int | None, int | None], Sequence[int | None, int | None]]` representing `sheet.span(start row, start column, up to but not including row, up to but not including column)` e.g. `sheet.span(0, 0, 2, 2)`
    - `Span` e.g `sheet.span(another_span)`
- `type_` (`str`) must be either an empty string `""` or one of the following: `"format"`, `"highlight"`, `"dropdown"`, `"checkbox"`, `"readonly"`, `"align"`.
- `name` (`str`) used for named spans or for identification. If no name is provided then a name is generated for the span which is based on an internal integer ticker and then converted to a string in the same way column names are.
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
- `emit_event` when `True` and when using data setting functions that utilize spans causes a `"<<SheetModified>>` event to occur if it has been bound, see [here](https://github.com/ragardner/tksheet/wiki/Version-7#tkinter-and-tksheet-events) for more information on binding this event.
- `widget` (`object`) is the reference to the original sheet which created the span. This can be changed to a different sheet if required e.g. `my_span.widget = new_sheet`.
- `expand` (`None`, `str`) must be either `None` or:
    - `"table"`/`"both"` expand the span both down and right from the span start to the ends of the table.
    - `"right"` expand the span right to the end of the table `x` axis.
    - `"down"` expand the span downwards to the bottom of the table `y` axis.
- `formatter_options` (`dict`, `None`) must be either `None` or `dict`. If providing a `dict` it must be the same structure as used in format functions, see [here](https://github.com/ragardner/tksheet/wiki/Version-7#data-formatting) for more information. Used to turn the span into a format type span which:
    - When using `get_data()` will format the returned data.
    - When using `set_data()` will format the data being set but **NOT** create a new formatting rule on the sheet.
- `**kwargs` you can provide additional keyword arguments to the function for example those used in `span.highlight()` or `span.dropdown()` which are used when applying a named span to a table.

Notes:
- To create a named span see [here](https://github.com/ragardner/tksheet/wiki/Version-7#named-spans).

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
span = sheet[0, 0] # cell A1
span = sheet[(0, 0)] # cell A1
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

# after importing num2alpha from tksheet
print (sheet[num2alpha(0)].kind)
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

Spans have the following methods, all of which return the span object itself so you can chain the functions e.g. `span.options(undo=True).clear().bg = "indianred1"`

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
    emit_event: bool | None = None,
    widget: object = None,
    expand: str | None = None,
    formatter_options: dict | None = None,
    **kwargs,
) -> Span
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
- `emit_event` (`bool`, `None`) is used by data modifying functions that utilize spans. When `True` causes a `"<<SheetModified>>` event to occur if it has been bound, see [here](https://github.com/ragardner/tksheet/wiki/Version-7#tkinter-and-tksheet-events) for more information.
- `widget` (`object`) is the reference to the original sheet which created the span. This can be changed to a different sheet if required e.g. `my_span.widget = new_sheet`.
- `expand` (`str`, `None`) must be either `None` or:
    - `"table"`/`"both"` expand the span both down and right from the span start to the ends of the table.
    - `"right"` expand the span right to the end of the table `x` axis.
    - `"down"` expand the span downwards to the bottom of the table `y` axis.
- `formatter_options` (`dict`, `None`) must be either `None` or `dict`. If providing a `dict` it must be the same structure as used in format functions, see [here](https://github.com/ragardner/tksheet/wiki/Version-7#data-formatting) for more information. Used to turn the span into a format type span which:
    - When using `get_data()` will format the returned data.
    - When using `set_data()` will format the data being set but **NOT** create a new formatting rule on the sheet.
- `**kwargs` you can provide additional keyword arguments to the function for example those used in `span.highlight()` or `span.dropdown()` which are used when applying a named span to a table.
- This function returns the span instance itself (`self`).

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
- `name` (`str`) used for named spans or for identification. If no name is provided then a name is generated for the span which is based on an internal integer ticker and then converted to a string in the same way column names are.
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
- `emit_event` (`bool`) is used by data modifying functions that utilize spans. When `True` causes a `"<<SheetModified>>` event to occur if it has been bound, see [here](https://github.com/ragardner/tksheet/wiki/Version-7#tkinter-and-tksheet-events) for more information.
- `widget` (`object`) is the reference to the original sheet which created the span. This can be changed to a different sheet if required e.g. `my_span.widget = new_sheet`.
- `kwargs` a `dict` containing keyword arguments relevant for functions such as `span.highlight()` or `span.dropdown()` which are used when applying a named span to a table.

If necessary you can also modify these attributes the same way you would an objects. e.g.

```python
# span now takes in all columns, including A
span = self.sheet("A")
span.upto_c = None

# span now adds to sheets undo stack when using data modifying functions that use spans
span = self.sheet("A")
span.undo = True
```

#### **Using a span to format data**

Formats table data, see the help on [formatting](https://github.com/ragardner/tksheet/wiki/Version-7#data-formatting) for more information. Note that using this function also creates a format rule for the affected table cells.

```python
span.format(
    formatter_options: dict = {},
    formatter_class: object = None,
    redraw: bool = True,
    **kwargs,
) -> Span
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
span.del_format() -> Span
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
) -> Span
```

There are two ways to create highlights using a span:

Method 1 example using `.highlight()`:
```python
# highlights column A background red, text color black
sheet["A"].highlight(bg="red", fg="black")

# the same but after having saved a span
my_span = sheet["A"]
my_span.highlight(bg="red", fg="black")
```

Method 2 example using `.bg`/`.fg`:
```python
# highlights column A background red, text color black
sheet["A"].bg = "red"
sheet["A"].fg = "black"

# the same but after having saved a span
my_span = sheet["A"]
my_span.bg = "red"
my_span.fg = "black"
```

#### **Using a span to delete highlights**

Delete any currently existing highlights for parts of the sheet that are covered by the span. Should not be used where there are highlights created by named spans, see [Named spans](https://github.com/ragardner/tksheet/wiki/Version-7#named-spans) for more information.

```python
span.dehighlight() -> Span
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
) -> Span
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
span.del_dropdown() -> Span
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
) -> Span
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
span.del_checkbox() -> Span
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
span.readonly(readonly: bool = True) -> Span
```
- Using `span.readonly(False)` deletes any existing readonly rules for the span. Should not be used where there are readonly rules created by named spans, see [Named spans](https://github.com/ragardner/tksheet/wiki/Version-7#named-spans) for more information.

#### **Using a span to create text alignment rules**

Create a text alignment rule for parts of the sheet that are covered by the span.

```python
span.align(
    align: str | None,
    redraw: bool = True,
) -> Span
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

There are two ways to create alignment rules using a span:

Method 1 example using `.align()`:
```python
# column D right text alignment
sheet["D"].align("right")

# the same but after having saved a span
my_span = sheet["D"]
my_span.align("right")
```

Method 2 example using `.align = `:
```python
# column D right text alignment
sheet["D"].align = "right"

# the same but after having saved a span
my_span = sheet["D"]
my_span.align = "right"
```

#### **Using a span to delete text alignment rules**

Delete text alignment rules for parts of the sheet that are covered by the span. Should not be used where there are alignment rules created by named spans, see [Named spans](https://github.com/ragardner/tksheet/wiki/Version-7#named-spans) for more information.

```python
span.del_align() -> Span
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
    emit_event: bool | None = None,
    redraw: bool = True,
) -> Span
```
Parameters:
- `undo` (`bool`, `None`) When `True` if undo is enabled for the end user they will be able to undo the clear change.
- `emit_event` when `True` causes a `"<<SheetModified>>` event to occur if it has been bound, see [here](https://github.com/ragardner/tksheet/wiki/Version-7#tkinter-and-tksheet-events) for more information.

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
span.transpose() -> Span
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
span.expand(direction: str = "both") -> Span
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

Examples of creating named spans:
```python
# Will highlight rows 3 up to and including 5
span1 = self.sheet.span(
    "3:5",
    type_="highlight",
    bg="green",
    fg="black",
)
self.sheet.named_span(span1)

#  Will always keep the entire sheet formatted as `int` no matter how many rows/columns are inserted
span2 = self.sheet.span(
    ":",
    # you don't have to provide a `type_` when using the `formatter_kwargs` argument
    formatter_options=int_formatter(),
)
self.sheet.named_span(span2)
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
# ValueError: Span 'this name doesnt exist' does not exist.
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

The above spans represent the cell `A1` - row 0, column 0.

A reserved span attribute named `data` can then be used to retrieve the data for cell `A1`, example below:

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

This function is useful if you need a lot of sheet data, and produces one row at a time (may save memory use in certain scenarios). It does not use spans.

```python
yield_sheet_rows(
    get_displayed: bool = False,
    get_header: bool = False,
    get_index: bool = False,
    get_index_displayed: bool = True,
    get_header_displayed: bool = True,
    only_rows: int | Iterator[int] | None = None,
    only_columns: int | Iterator[int] | None = None,
) -> Iterator[list[object]]
```
Parameters:
- `get_displayed` (`bool`) if `True` it will return cell values as they are displayed on the screen. If `False` it will return any underlying data, for example if the cell is formatted.
- `get_header` (`bool`) if `True` it will return the header of the sheet even if there is not one.
- `get_index` (`bool`) if `True` it will return the index of the sheet even if there is not one.
- `get_index_displayed` (`bool`) if `True` it will return whatever index values are displayed on the screen, for example if there is a dropdown box with `text` set.
- `get_header_displayed` (`bool`) if `True` it will return whatever header values are displayed on the screen, for example if there is a dropdown box with `text` set.
- `only_rows` (`None`, `iterable`) with this argument you can supply an iterable of `int` row indexes in any order to be the only rows that are returned.
- `only_columns` (`None`, `iterable`) with this argument you can supply an iterable of `int` column indexes in any order to be the only columns that are returned.

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

#### **Get the number of rows in the sheet**

```python
get_total_rows(include_index: bool = False) -> int
```

___

#### **Get the number of columns in the sheet**

```python
get_total_columns(include_header: bool = False) -> int
```

___

#### **Get a value for a particular cell if that cell was empty**

```python
get_value_for_empty_cell(
    r: int,
    c: int,
    r_ops: bool = True,
    c_ops: bool = True,
) -> object
```
- `r_ops`/`c_ops` when both are `True` it will take into account whatever cell/row/column options exist. When just `r_ops` is `True` it will take into account row options only and when just `c_ops` is `True` it will take into account column options only.

---
# **Setting Sheet Data**

Fundamentally, there are two ways to set table data:
- Overwriting the entire table and setting the table data to a new object.
- Modifying the existing data.

#### **Overwriting the table**

```python
set_sheet_data(
    data: list | tuple | None = None,
    reset_col_positions: bool = True,
    reset_row_positions: bool = True,
    redraw: bool = True,
    verify: bool = False,
    reset_highlights: bool = False,
    keep_formatting: bool = True,
    delete_options: bool = False,
) -> object
```
Parameters:
- `data` (`list`) has to be a list of lists for full functionality, for display only a list of tuples or a tuple of tuples will work.
- `reset_col_positions` and `reset_row_positions` (`bool`) when `True` will reset column widths and row heights.
- `redraw` (`bool`) refreshes the table after setting new data.
- `verify` (`bool`) goes through `data` and checks if it is a list of lists, will raise error if not, disabled by default.
- `reset_highlights` (`bool`) resets all table cell highlights.
- `keep_formatting` (`bool`) when `True` re-applies any prior formatting rules to the new data, if `False` all prior formatting rules are deleted.
- `delete_options` (`bool`) when `True` all table options such as dropdowns, check boxes, formatting, highlighting etc. are deleted.

Notes:
- This function does not impact the sheet header or index.

___

```python
@data.setter
data(value: object)
```
Notes:
- Acts like setting an attribute e.g. `sheet.data = [[1, 2, 3], [4, 5, 6]]`
- Uses the `set_sheet_data()` function and its default arguments.

___

#### **Reset all or specific sheet elements and attributes**

```python
reset(
    table: bool = True,
    header: bool = True,
    index: bool = True,
    row_heights: bool = True,
    column_widths: bool = True,
    cell_options: bool = True,
    undo_stack: bool = True,
    selections: bool = True,
    sheet_options: bool = False,
    redraw: bool = True,
) -> Sheet
```
Parameters:
- `table` when `True` resets the table to an empty list.
- `header` when `True` resets the header to an empty list.
- `index` when `True` resets the row index to an empty list.
- `row_heights` when `True` deletes all displayed row lines.
- `column_widths` when `True` deletes all displayed column lines.
- `cell_options` when `True` deletes all dropdowns, checkboxes, highlights, data formatting, etc.
- `undo_stack` when `True` resets the sheets undo stack to empty.
- `selections` when `True` deletes all selection boxes.
- `sheet_options` when `True` resets all the sheets options such as colors, font, popup menu labels and many more to default, for a full list of what's reset see the file `sheet_options.py`.

Notes:
- This function could be useful when a whole new sheet needs to be loaded.

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

The above span example represents the cell `A1` - row 0, column 0. A reserved span attribute named `data` (you can also use `.value`) can then be used to modify sheet data **starting** from cell `A1`. example below:

```python
span = self.sheet["A1"]
span.data = "new value for cell A1"

# or even shorter:
self.sheet["A1"].data = "new value for cell A1"

# or with sheet.span()
self.sheet.span("A1").data = "new value for cell A1"
```

If you provide a list or tuple it will set more than one cell, starting from the spans start cell. In the example below three cells are set in the first row, **starting from cell B1**:

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

"""
SETTING CELL AREA DATA INCLUDING HEADER AND INDEX
"""
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

#### **Sheet set data function**

You can also use the `Sheet` function `set_data()`.

```python
set_data(
    *key: CreateSpanTypes,
    data: object = None,
    undo: bool | None = None,
    emit_event: bool | None = None,
    redraw: bool = True,
) -> EventDataDict
```
Parameters:
- `undo` when `True` adds the change to the Sheets undo stack.
- `emit_event` when `True` causes a `"<<SheetModified>>` event to occur if it has been bound, see [here](https://github.com/ragardner/tksheet/wiki/Version-7#tkinter-and-tksheet-events) for more information.

Example:
```python
self.sheet.set_data(
    "A1",
    [["",  "A",  "B",  "C"]
     ["1", "A1", "B1", "C1"],
     ["2", "A2", "B2", "C2"]],
)
```

You can clear cells/rows/columns using a [spans `clear()` function](https://github.com/ragardner/tksheet/wiki/Version-7#using-a-span-to-clear-cells) or the Sheets `clear()` function. Below is the Sheets clear function:

```python
clear(
    *key: CreateSpanTypes,
    undo: bool | None = None,
    emit_event: bool | None = None,
    redraw: bool = True,
) -> EventDataDict
```
- `undo` when `True` adds the change to the Sheets undo stack.
- `emit_event` when `True` causes a `"<<SheetModified>>` event to occur if it has been bound, see [here](https://github.com/ragardner/tksheet/wiki/Version-7#tkinter-and-tksheet-events) for more information.

#### **Insert a row into the sheet**

```python
insert_row(
    row: list[object] | tuple[object] | None = None,
    idx: str | int | None = None,
    height: int | None = None,
    row_index: bool = False,
    fill: bool = True,
    undo: bool = False,
    emit_event: bool = False,
    redraw: bool = True,
) -> EventDataDict
```
Parameters:
- Leaving `row` as `None` inserts an empty row, e.g. `insert_row()` will append an empty row to the sheet.
- `height` is the new rows displayed height in pixels, leave as `None` for default.
- `row_index` when `True` assumes there is a row index value at the start of the row.
- `fill` when `True` any provided rows that are shorter than the Sheets longest row will be filled with empty values up to the length of the longest row.
- `undo` when `True` adds the change to the Sheets undo stack.
- `emit_event` when `True` causes a `"<<SheetModified>>` event to occur if it has been bound, see [here](https://github.com/ragardner/tksheet/wiki/Version-7#tkinter-and-tksheet-events) for more information.

___

#### **Insert a column into the sheet**

```python
insert_column(
    column: list[object] | tuple[object] | None = None,
    idx: str | int | None = None,
    width: int | None = None,
    header: bool = False,
    fill: bool = True,
    undo: bool = False,
    emit_event: bool = False,
    redraw: bool = True,
) -> EventDataDict
```
Parameters:
- Leaving `column` as `None` inserts an empty column, e.g. `insert_column()` will append an empty column to the sheet.
- `width` is the new columns displayed width in pixels, leave as `None` for default.
- `header` when `True` assumes there is a header value at the start of the column.
- `fill` when `True` any provided columns that are shorter than the Sheets longest column will be filled with empty values up to the length of the longest column.
- `undo` when `True` adds the change to the Sheets undo stack.
- `emit_event` when `True` causes a `"<<SheetModified>>` event to occur if it has been bound, see [here](https://github.com/ragardner/tksheet/wiki/Version-7#tkinter-and-tksheet-events) for more information.

___

#### **Insert multiple columns into the sheet**

```python
insert_columns(
    columns: list[tuple[object] | list[object]] | tuple[tuple[object] | list[object]] | int = 1,
    idx: str | int | None = None,
    widths: list[int] | tuple[int] | None = None,
    headers: bool = False,
    create_selections: bool = True,
    undo: bool = False,
    emit_event: bool = False,
    redraw: bool = True,
) -> EventDataDict
```
Parameters:
- `columns` if `int` will insert that number of blank columns.
- `idx` (`str`, `int`, `None`) either `str` e.g. `"A"` for `0`, `int` or `None` for end.
- `widths` are the new columns displayed widths in pixels, leave as `None` for default.
- `headers` when `True` assumes there are headers values at the start of each column.
- `fill` when `True` any provided columns that are shorter than the Sheets longest column will be filled with empty values up to the length of the longest column.
- `undo` when `True` adds the change to the Sheets undo stack.
- `emit_event` when `True` causes a `"<<SheetModified>>` event to occur if it has been bound, see [here](https://github.com/ragardner/tksheet/wiki/Version-7#tkinter-and-tksheet-events) for more information.

___

#### **Insert multiple rows into the sheet**

```python
insert_rows(
    rows: list[tuple[object] | list[object]] | tuple[tuple[object] | list[object]] | int = 1,
    idx: str | int | None = None,
    heights: list[int] | tuple[int] | None = None,
    row_index: bool = False,
    fill: bool = True,
    undo: bool = False,
    emit_event: bool = False,
    redraw: bool = True,
) -> EventDataDict
```
Parameters:
- `rows` if `int` will insert that number of blank rows.
- `idx` (`str`, `int`, `None`) either `str` e.g. `"A"` for `0`, `int` or `None` for end.
- `heights` are the new rows displayed heights in pixels, leave as `None` for default.
- `row_index` when `True` assumes there are row index values at the start of each row.
- `fill` when `True` any provided rows that are shorter than the Sheets longest row will be filled with empty values up to the length of the longest row.
- `undo` when `True` adds the change to the Sheets undo stack.
- `emit_event` when `True` causes a `"<<SheetModified>>` event to occur if it has been bound, see [here](https://github.com/ragardner/tksheet/wiki/Version-7#tkinter-and-tksheet-events) for more information.

___

#### **Delete a row from the sheet**

```python
del_row(
    idx: int = 0,
    data_indexes: bool = True,
    undo: bool = False,
    emit_event: bool = False,
    redraw: bool = True,
) -> EventDataDict
```
Parameters:
- `idx` is the row to delete.
- `data_indexes` only applicable when there are hidden rows. When `False` it makes the `idx` represent a displayed row and not the underlying Sheet data row. When `True` the index represent a data index.
- `undo` when `True` adds the change to the Sheets undo stack.
- `emit_event` when `True` causes a `"<<SheetModified>>` event to occur if it has been bound, see [here](https://github.com/ragardner/tksheet/wiki/Version-7#tkinter-and-tksheet-events) for more information.

___

#### **Delete multiple rows from the sheet**

```python
del_rows(
    rows: int | Iterator[int],
    data_indexes: bool = True,
    undo: bool = False,
    emit_event: bool = False,
    redraw: bool = True,
) -> EventDataDict
```
Parameters:
- `rows` can be either `int` or an iterable of `int`s representing row indexes.
- `data_indexes` only applicable when there are hidden rows. When `False` it makes the `rows` indexes represent displayed rows and not the underlying Sheet data rows. When `True` the indexes represent data indexes.
- `undo` when `True` adds the change to the Sheets undo stack.
- `emit_event` when `True` causes a `"<<SheetModified>>` event to occur if it has been bound, see [here](https://github.com/ragardner/tksheet/wiki/Version-7#tkinter-and-tksheet-events) for more information.

___

#### **Delete a column from the sheet**

```python
del_column(
    idx: int = 0,
    data_indexes: bool = True,
    undo: bool = False,
    emit_event: bool = False,
    redraw: bool = True,
) -> EventDataDict
```
Parameters:
- `idx` is the column to delete.
- `data_indexes` only applicable when there are hidden columns. When `False` it makes the `idx` represent a displayed column and not the underlying Sheet data column. When `True` the index represent a data index.
- `undo` when `True` adds the change to the Sheets undo stack.
- `emit_event` when `True` causes a `"<<SheetModified>>` event to occur if it has been bound, see [here](https://github.com/ragardner/tksheet/wiki/Version-7#tkinter-and-tksheet-events) for more information.

___

#### **Delete multiple columns from the sheet**

```python
del_columns(
    columns: int | Iterator[int],
    data_indexes: bool = True,
    undo: bool = False,
    emit_event: bool = False,
    redraw: bool = True,
) -> EventDataDict
```
Parameters:
- `columns` can be either `int` or an iterable of `int`s representing column indexes.
- `data_indexes` only applicable when there are hidden columns. When `False` it makes the `columns` indexes represent displayed columns and not the underlying Sheet data columns. When `True` the indexes represent data indexes.
- `undo` when `True` adds the change to the Sheets undo stack.
- `emit_event` when `True` causes a `"<<SheetModified>>` event to occur if it has been bound, see [here](https://github.com/ragardner/tksheet/wiki/Version-7#tkinter-and-tksheet-events) for more information.

___

Expands or contracts the sheet **data** dimensions.
```python
sheet_data_dimensions(
    total_rows: int | None = None,
    total_columns: int | None = None,
) -> Sheet
```
Parameters:
- `total_rows` sets the Sheets number of data rows.
- `total_columns` sets the Sheets number of data columns.

___

```python
set_sheet_data_and_display_dimensions(
    total_rows: int | None = None,
    total_columns: int | None = None,
) -> Sheet
```
Parameters:
- `total_rows` when `int` will set the number of the Sheets data and display rows by deleting or adding rows.
- `total_columns` when `int` will set the number of the Sheets data and display columns by deleting or adding columns.

___

```python
total_rows(
    number: int | None = None,
    mod_positions: bool = True,
    mod_data: bool = True,
) -> int | Sheet
```
Parameters:
- `number` sets the Sheets number of data rows. When `None` function will return the Sheets number of data rows including the number of rows in the index.
- `mod_positions` when `True` also sets the number of displayed rows.
- `mod_data` when `True` also sets the number of data rows.

___

```python
total_columns(
    number: int | None = None,
    mod_positions: bool = True,
    mod_data: bool = True,
) -> int | Sheet
```
Parameters:
- `number` sets the Sheets number of data columns. When `None` function will return the Sheets number of data columns including the number of columns in the header.
- `mod_positions` when `True` also sets the number of displayed columns.
- `mod_data` when `True` also sets the number of data columns.

___

#### **Move a single row to a new location**

```python
move_row(
    row: int,
    moveto: int,
) -> tuple[dict, dict, dict]
```
- Note that `row` and `moveto` indexes represent displayed indexes and not data. When there are hidden rows this is an important distinction, otherwise it is not at all important. To specifically use data indexes use the function `move_rows()`.

___

#### **Move a single column to a new location**

```python
move_column(
    column: int,
    moveto: int,
) -> tuple[dict, dict, dict]
```
- Note that `column` and `moveto` indexes represent displayed indexes and not data. When there are hidden columns this is an important distinction, otherwise it is not at all important. To specifically use data indexes use the function `move_columns()`.

___

#### **Move any rows to a new location**

```python
move_rows(
    move_to: int | None = None,
    to_move: list[int] | None = None,
    move_data: bool = True,
    data_indexes: bool = False,
    create_selections: bool = True,
    undo: bool = False,
    emit_event: bool = False,
    redraw: bool = True,
) -> tuple[dict, dict, dict]
```
Parameters:
- `move_to` is the new start index for the rows to be moved to.
- `to_move` is a `list` of row indexes to move to that new position, they will appear in the same order provided.
- `move_data` when `True` moves not just the displayed row positions but the Sheet data as well.
- `data_indexes` is only applicable when there are hidden rows. When `False` it makes the `move_to` and `to_move` indexes represent displayed rows and not the underlying Sheet data rows. When `True` the indexes represent data indexes.
- `create_selections` creates new selection boxes based on where the rows have moved.
- `undo` when `True` adds the change to the Sheets undo stack.
- `emit_event` when `True` causes a `"<<SheetModified>>` event to occur if it has been bound, see [here](https://github.com/ragardner/tksheet/wiki/Version-7#tkinter-and-tksheet-events) for more information.

Notes:
- The rows in `to_move` do **not** have to be contiguous.

___

#### **Move any columns to a new location**

```python
move_columns(
    move_to: int | None = None,
    to_move: list[int] | None = None,
    move_data: bool = True,
    data_indexes: bool = False,
    create_selections: bool = True,
    undo: bool = False,
    emit_event: bool = False,
    redraw: bool = True,
) -> tuple[dict, dict, dict]
```
Parameters:
- `move_to` is the new start index for the columns to be moved to.
- `to_move` is a `list` of column indexes to move to that new position, they will appear in the same order provided.
- `move_data` when `True` moves not just the displayed column positions but the Sheet data as well.
- `data_indexes` is only applicable when there are hidden columns. When `False` it makes the `move_to` and `to_move` indexes represent displayed columns and not the underlying Sheet data columns. When `True` the indexes represent data indexes.
- `create_selections` creates new selection boxes based on where the columns have moved.
- `undo` when `True` adds the change to the Sheets undo stack.
- `emit_event` when `True` causes a `"<<SheetModified>>` event to occur if it has been bound, see [here](https://github.com/ragardner/tksheet/wiki/Version-7#tkinter-and-tksheet-events) for more information.

Notes:
- The columns in `to_move` do **not** have to be contiguous.

___

#### **Move any columns to new locations**

```python
mapping_move_columns(
    data_new_idxs: dict[int, int],
    disp_new_idxs: None | dict[int, int] = None,
    move_data: bool = True,
    create_selections: bool = True,
    data_indexes: bool = False,
    undo: bool = False,
    emit_event: bool = False,
    redraw: bool = True,
) -> tuple[dict[int, int], dict[int, int], EventDataDict]
```
Parameters:
- `data_new_idxs` (`dict[int, int]`) must be a `dict` where the keys are the data columns to move as `int`s and the values are their new locations as `int`s.
- `disp_new_idxs` (`None | dict[int, int]`) either `None` or a `dict` where the keys are the displayed columns (basically the column widths) to move as `int`s and the values are their new locations as `int`s. If `None` then no column widths will be moved.
- `move_data` when `True` moves not just the displayed column positions but the Sheet data as well.
- `data_indexes` is only applicable when there are hidden columns. When `False` it makes the `move_to` and `to_move` indexes represent displayed columns and not the underlying Sheet data columns. When `True` the indexes represent data indexes.
- `create_selections` creates new selection boxes based on where the columns have moved.
- `undo` when `True` adds the change to the Sheets undo stack.
- `emit_event` when `True` causes a `"<<SheetModified>>` event to occur if it has been bound, see [here](https://github.com/ragardner/tksheet/wiki/Version-7#tkinter-and-tksheet-events) for more information.

___

#### **Move any rows to new locations**

```python
mapping_move_rows(
    data_new_idxs: dict[int, int],
    disp_new_idxs: None | dict[int, int] = None,
    move_data: bool = True,
    data_indexes: bool = False,
    create_selections: bool = True,
    undo: bool = False,
    emit_event: bool = False,
    redraw: bool = True,
) -> tuple[dict[int, int], dict[int, int], EventDataDict]
```
Parameters:
- `data_new_idxs` (`dict[int, int]`) must be a `dict` where the keys are the data rows to move as `int`s and the values are their new locations as `int`s.
- `disp_new_idxs` (`None | dict[int, int]`) either `None` or a `dict` where the keys are the displayed rows (basically the row heights) to move as `int`s and the values are their new locations as `int`s. If `None` then no row heights will be moved.
- `move_data` when `True` moves not just the displayed row positions but the Sheet data as well.
- `data_indexes` is only applicable when there are hidden rows. When `False` it makes the `move_to` and `to_move` indexes represent displayed rows and not the underlying Sheet data rows. When `True` the indexes represent data indexes.
- `create_selections` creates new selection boxes based on where the rows have moved.
- `undo` when `True` adds the change to the Sheets undo stack.
- `emit_event` when `True` causes a `"<<SheetModified>>` event to occur if it has been bound, see [here](https://github.com/ragardner/tksheet/wiki/Version-7#tkinter-and-tksheet-events) for more information.

___

#### **Make all data rows the same length**

```python
equalize_data_row_lengths(include_header: bool = False) -> int
```
- Makes every list in the table have the same number of elements, goes by longest list. This will only affect the data variable, not visible columns.
- Returns the new row length for all rows in the Sheet.

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
get_dropdown_values(r: int = 0, c: int = 0) -> None | list
```

```python
get_header_dropdown_values(c: int = 0) -> None | list
```

```python
get_index_dropdown_values(r: int = 0) -> None | list
```

___

#### **Set the values and cell value of a chosen dropdown box**

```python
set_dropdown_values(
    r: int = 0,
    c: int = 0,
    set_existing_dropdown: bool = False,
    values: list = [],
    set_value: object = None,
) -> Sheet
```

```python
set_header_dropdown_values(
    c: int = 0,
    set_existing_dropdown: bool = False,
    values: list = [],
    set_value: object = None,
) -> Sheet
```

```python
set_index_dropdown_values(
    r: int = 0,
    set_existing_dropdown: bool = False,
    values: list = [],
    set_value: object = None,
) -> Sheet
```
Parameters:
- `set_existing_dropdown` if `True` takes priority over `r` and `c` and sets the values of the last popped open dropdown box (if one one is popped open, if not then an `Exception` is raised).
- `values` (`list`, `tuple`)
- `set_value` (`str`, `None`) if not `None` will try to set the value of the chosen cell to given argument.

___

#### **Set or get bound dropdown functions**

```python
dropdown_functions(
    r: int,
    c: int,
    selection_function: str | Callable = "",
    modified_function: str | Callable = "",
) -> None | dict
```

```python
header_dropdown_functions(
    c: int,
    selection_function: str | Callable = "",
    modified_function: str | Callable = "",
) -> None | dict
```

```python
index_dropdown_functions(
    r: int,
    selection_function: str | Callable = "",
    modified_function: str | Callable = "",
) -> None | dict
```

___

#### **Get a dictionary of all dropdown boxes**

```python
get_dropdowns() -> dict
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
get_header_dropdowns() -> dict
```

```python
get_index_dropdowns() -> dict
```

___

#### **Open a dropdown box**

```python
open_dropdown(r: int, c: int) -> Sheet
```

```python
open_header_dropdown(c: int) -> Sheet
```

```python
open_index_dropdown(r: int) -> Sheet
```

___

#### **Close a dropdown box**

```python
close_dropdown(r: int, c: int) -> Sheet
```

```python
close_header_dropdown(c: int) -> Sheet
```

```python
close_index_dropdown(r: int) -> Sheet
```
Notes:
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
click_checkbox(
    *key: CreateSpanTypes,
    checked: bool | None = None,
    redraw: bool = True,
) -> Span
```

```python
click_header_checkbox(c: int, checked: bool | None = None) -> Sheet
```

```python
click_index_checkbox(r: int, checked: bool | None = None) -> Sheet
```

___

#### **Get a dictionary of all check box dictionaries**

```python
get_checkboxes() -> dict
```

```python
get_header_checkboxes() -> dict
```

```python
get_index_checkboxes() -> dict
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

`Span` objects (more information [here](https://github.com/ragardner/tksheet/wiki/Version-7#span-objects)) can be used to format data for cells, rows, columns and the entire sheet.

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
del_all_formatting(clear_values: bool = False) -> Sheet
```
- `clear_values` (`bool`) if true, all the sheets cell values will be cleared.

#### **Reapply formatting to entire sheet**

```python
reapply_formatting() -> Sheet
```
- Useful if you have manually changed the entire sheets data using `sheet.MT.data = ` and want to reformat the sheet using any existing formatting you have set.

#### **Check if a cell is formatted**

```python
formatted(r: int, c: int) -> dict
```
- If the cell is formatted function returns a `dict` with all the format keyword arguments. The `dict` will be empty if the cell is not formatted.

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
) -> dict
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
) -> dict
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
) -> dict
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
) -> dict
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
) -> dict
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
) -> Span
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

**Set the table and index font**

```python
font(
    newfont: tuple[str, int, str] | None = None,
    reset_row_positions: bool = True,
) -> tuple[str, int, str]
```

**Set the header font**

```python
header_font(newfont: tuple[str, int, str] | None = None) -> tuple[str, int, str]
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
table_align(
    align: str = None,
    redraw: bool = True,
) -> str | Sheet
```

Set the text alignment for the whole of the header.
```python
header_align(
    align: str = None,
    redraw: bool = True,
) -> str | Sheet
```

Set the text alignment for the whole of the index.
```python
row_index_align(
    align: str = None,
    redraw: bool = True,
) -> str | Sheet

# can also use index_align() which behaves the same
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
) -> Span
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
) -> Span
```
Parameters:
- `key` (`CreateSpanTypes`) either a span or a type which can create a span. See [here](https://github.com/ragardner/tksheet/wiki/Version-7#creating-a-span) for more information on the types that can create a span.

#### **Get existing specific text alignments**

Cell text alignments:
```python
get_cell_alignments() -> dict
```

Row text alignments:
```python
get_row_alignments() -> dict
```

Column text alignments:
```python
get_column_alignments() -> dict
```

---
# **Getting Selected Cells**

```python
get_currently_selected() -> tuple | CurrentlySelectedClass
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
get_selected_rows(
    get_cells: bool = False,
    get_cells_as_rows: bool = False,
    return_tuple: bool = False,
) -> tuple | set
```

___

```python
get_selected_columns(
    get_cells: bool = False,
    get_cells_as_columns: bool = False,
    return_tuple: bool = False,
) -> tuple | set
```

___

```python
get_selected_cells(
    get_rows: bool = False,
    get_columns: bool = False,
    sort_by_row: bool = False,
    sort_by_column: bool = False,
) -> list | set
```

___

```python
get_all_selection_boxes() -> tuple[tuple[int, int, int, int]]
```

___

```python
get_all_selection_boxes_with_types() -> list[tuple[tuple[int, int, int, int], str]]
```

___

#### **Check if cell is selected**

```python
cell_selected(r: int, c: int) -> bool
```

___

#### **Check if row is selected**

```python
row_selected(r: int) -> bool
```

___

#### **Check if column is selected**

```python
column_selected(c: int) -> bool
```

___

#### **Check if any cells, rows or columns are selected, there are options for exclusions**

```python
anything_selected(
    exclude_columns: bool = False,
    exclude_rows: bool = False,
    exclude_cells: bool = False,
) -> bool
```

___

#### **Check if user has the entire table selected**

```python
all_selected() -> bool
```

___

```python
get_ctrl_x_c_boxes() -> tuple[dict[tuple[int, int, int, int], str], int]
```

___

```python
get_selected_min_max() -> tuple[int, int, int, int] | tuple[None, None, None, None]
```
- returns `(min_y, min_x, max_y, max_x)` of any selections including rows/columns.

---
# **Modifying Selected Cells**

```python
set_currently_selected(row: int | None = None, column: int | None = None) -> Sheet
```

___

```python
select_row(row: int, redraw: bool = True, run_binding_func: bool = True) -> Sheet
```
- `run_binding_func` is only relevant if you have `extra_bindings()` with `"row_select"` bound.

___

```python
select_column(column: int, redraw: bool = True, run_binding_func: bool = True) -> Sheet
```
- `run_binding_func` is only relevant if you have `extra_bindings()` with `"column_select"` bound.

___

```python
select_cell(row: int, column: int, redraw: bool = True, run_binding_func: bool = True) -> Sheet
```
- `run_binding_func` is only relevant if you have `extra_bindings()` with `"cell_select"` bound.

___

```python
select_all(redraw: bool = True, run_binding_func: bool = True) -> Sheet
```
- `run_binding_func` is only relevant if you have `extra_bindings()` with `"select_all"` bound.

___

```python
add_cell_selection(
    row: int,
    column: int,
    redraw: bool = True,
    run_binding_func: bool = True,
    set_as_current: bool = True,
) -> Sheet
```
- `run_binding_func` is only relevant if you have `extra_bindings()` with `"cell_select"` bound.

___

```python
add_row_selection(
    row: int,
    redraw: bool = True,
    run_binding_func: bool = True,
    set_as_current: bool = True,
) -> Sheet
```
- `run_binding_func` is only relevant if you have `extra_bindings()` with `"row_select"` bound.

___

```python
add_column_selection(
    column: int,
    redraw: bool = True,
    run_binding_func: bool = True,
    set_as_current: bool = True,
) -> Sheet
```
- `run_binding_func` is only relevant if you have `extra_bindings()` with `"column_select"` bound.

___

```python
toggle_select_cell(
    row: int,
    column: int,
    add_selection: bool = True,
    redraw: bool = True,
    run_binding_func: bool = True,
    set_as_current: bool = True,
) -> Sheet
```
- `run_binding_func` is only relevant if you have `extra_bindings()` with `"cell_select"` bound.

___

```python
toggle_select_row(
    row: int,
    add_selection: bool = True,
    redraw: bool = True,
    run_binding_func: bool = True,
    set_as_current: bool = True,
) -> Sheet
```
- `run_binding_func` is only relevant if you have `extra_bindings()` with `"row_select"` bound.

___

```python
toggle_select_column(
    column: int,
    add_selection: bool = True,
    redraw: bool = True,
    run_binding_func: bool = True,
    set_as_current: bool = True,
) -> Sheet
```
- `run_binding_func` is only relevant if you have `extra_bindings()` with `"column_select"` bound.

___

```python
create_selection_box(
    r1: int,
    c1: int,
    r2: int,
    c2: int,
    type_: str = "cells",
) -> int
```
- `type_` either `"cells"` or `"rows"` or `"columns"`.
- Returns the canvas item id for the box.

___

```python
recreate_all_selection_boxes() -> Sheet
```

___

```python
deselect(
    row: int | None | str = None,
    column: int | None = None,
    cell: tuple | None = None,
    redraw: bool = True,
) -> Sheet
```

---
# **Row Heights and Column Widths**

#### **Auto resize column widths to fit the window**

To enable auto resizing of columns to the Sheet window use `set_options()` with the keyword argument `auto_resize_columns`. This argument can either be an `int` or `None`. If set as an `int` the columns will automatically resize to fit the width of the window, the `int` value being the minimum of each column in pixels. If `None` it will disable the auto resizing. Example:

```python
# auto resize columns, column minimum width set to 150 pixels
set_options(auto_resize_columns=150)
```

___

#### **Auto resize row heights to fit the window**

To enable auto resizing of rows to the Sheet window use `set_options()` with the keyword argument `auto_resize_rows`. This argument can either be an `int` or `None`. If set as an `int` the rows will automatically resize to fit the width of the window, the `int` value being the minimum of each row in pixels. If `None` it will disable the auto resizing. Example:

```python
# auto resize rows, row minimum width set to 30 pixels
set_options(auto_resize_rows=30)
```

___

#### **Set default column width in pixels**

```python
default_column_width(width: int | None = None) -> int
```
- `width` (`int`, `None`) use an `int` to set the width in pixels, `None` does not set the width.

___

#### **Set default row height in pixels or lines**

```python
default_row_height(height: int | str | None = None) -> int
```
- `height` (`int`, `str`, `None`) use a numerical `str` for number of lines e.g. `"3"` for a height that fits 3 lines OR `int` for pixels.

___

#### **Set default header bar height in pixels or lines**

```python
default_header_height(height: int | str | None = None) -> int
```
- `height` (`int`, `str`, `None`) use a numerical `str` for number of lines e.g. `"3"` for a height that fits 3 lines or `int` for pixels.

___

#### **Set a specific cell size to its text**

```python
set_cell_size_to_text(
    row: int,
    column: int,
    only_set_if_too_small: bool = False,
    redraw: bool = True,
) -> Sheet
```

___

#### **Set all row heights and column widths to cell text sizes**

```python
set_all_cell_sizes_to_text(redraw: bool = True) -> tuple[list[float], list[float]]
```
- Returns the Sheets row positions and column positions in that order.

___

#### **Set all column widths to a specific width in pixels**

```python
set_all_column_widths(
    width: int | None = None,
    only_set_if_too_small: bool = False,
    redraw: bool = True,
    recreate_selection_boxes: bool = True,
) -> Sheet
```
- `width` (`int`, `None`) leave `None` to set to cell text sizes for each column.

___

#### **Set all row heights to a specific height in pixels**

```python
set_all_row_heights(
    height: int | None = None,
    only_set_if_too_small: bool = False,
    redraw: bool = True,
    recreate_selection_boxes: bool = True,
) -> Sheet
```
- `height` (`int`, `None`) leave `None` to set to cell text sizes for each row.

___

#### **Set or get a specific column width**

```python
column_width(
    column: int | Literal["all", "displayed"] | None = None,
    width: int | Literal["default", "text"] | None = None,
    only_set_if_too_small: bool = False,
    redraw: bool = True,
) -> Sheet | int
```

___

#### **Set or get a specific row height**

```python
row_height(
    row: int | Literal["all", "displayed"] | None = None,
    height: int | Literal["default", "text"] | None = None,
    only_set_if_too_small: bool = False,
    redraw: bool = True,
) -> Sheet | int
```

___

#### **Get the sheets column widths**

```python
get_column_widths(canvas_positions: bool = False) -> list[float]
```
- `canvas_positions` (`bool`) gets the actual canvas x coordinates of column lines.

___

#### **Get the sheets row heights**

```python
get_row_heights(canvas_positions: bool = False) -> list[float]
```
- `canvas_positions` (`bool`) gets the actual canvas y coordinates of row lines.

___

```python
set_column_widths(
    column_widths: Iterator[int, float] | None = None,
    canvas_positions: bool = False,
    reset: bool = False,
) -> Sheet
```

___

```python
set_row_heights(
    row_heights: Iterator[int, float] | None = None,
    canvas_positions: bool = False,
    reset: bool = False,
) -> Sheet
```

___

```python
set_width_of_index_to_text(text: None | str = None, *args, **kwargs) -> Sheet
```
- `text` (`str`, `None`) provide a `str` to set the width to or use `None` to set it to existing values in the index.

___

```python
set_index_width(pixels: int, redraw: bool = True) -> Sheet
```
- Note that it disables auto resizing of index. Use [`set_options()`](https://github.com/ragardner/tksheet/wiki/Version-7#sheet-options-and-other-functions) to restore auto resizing.

___

```python
set_height_of_header_to_text(text: None | str = None) -> Sheet
```
- `text` (`str`, `None`) provide a `str` to set the height to or use `None` to set it to existing values in the header.

___

```python
set_header_height_pixels(pixels: int, redraw: bool = True) -> Sheet
```

___

```python
set_header_height_lines(nlines: int, redraw: bool = True) -> Sheet
```

___

```python
del_column_position(idx: int, deselect_all: bool = False) -> Sheet
```

___

```python
del_row_position(idx: int, deselect_all: bool = False) -> Sheet
```

___

```python
insert_row_position(
    idx: Literal["end"] | int = "end",
    height: int | None = None,
    deselect_all: bool = False,
    redraw: bool = False,
) -> Sheet
```

___

```python
insert_column_position(
    idx: Literal["end"] | int = "end",
    width: int | None = None,
    deselect_all: bool = False,
    redraw: bool = False,
) -> Sheet
```

___

```python
insert_row_positions(
    idx: Literal["end"] | int = "end",
    heights: Sequence[float] | int | None = None,
    deselect_all: bool = False,
    redraw: bool = False,
) -> Sheet
```

___

```python
insert_column_positions(
    idx: Literal["end"] | int = "end",
    widths: Sequence[float] | int | None = None,
    deselect_all: bool = False,
    redraw: bool = False,
) -> Sheet
```

___

```python
sheet_display_dimensions(
    total_rows: int | None =None,
    total_columns: int | None =None,
) -> tuple[int, int] | Sheet
```

___

```python
move_row_position(row: int, moveto: int) -> Sheet
```

___

```python
move_column_position(column: int, moveto: int) -> Sheet
```

___

```python
get_example_canvas_column_widths(total_cols: int | None = None) -> list[float]
```

___

```python
get_example_canvas_row_heights(total_rows: int | None = None) -> list[float]
```

___

```python
verify_row_heights(row_heights: list[float], canvas_positions: bool = False) -> bool
```

___

```python
verify_column_widths(column_widths: list[float], canvas_positions: bool = False) -> bool
```

---
# **Identifying Bound Event Mouse Position**

The below functions require a mouse click event, for example you could bind right click, example [here](https://github.com/ragardner/tksheet/wiki/Version-7#example-custom-right-click-and-text-editor-functionality), and then identify where the user has clicked.

___

```python
identify_region(event: object) -> Literal["table", "index", "header", "top left"]
```

___

```python
identify_row(
    event: object,
    exclude_index: bool = False,
    allow_end: bool = True,
) -> int | None
```

___

```python
identify_column(
    event: object,
    exclude_header: bool = False,
    allow_end: bool = True,
) -> int | None
```

___

#### **Sheet control actions for binding your own keys to**

For example: `sheet.bind("<Control-B>", sheet.paste)`

```python
cut(event: object = None) -> Sheet
copy(event: object = None) -> Sheet
paste(event: object = None) -> Sheet
delete(event: object = None) -> Sheet
undo(event: object = None) -> Sheet
redo(event: object = None) -> Sheet
```

---
# **Scroll Positions and Cell Visibility**

#### **Sync scroll positions between widgets**

```python
sync_scroll(widget: object) -> Sheet
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
unsync_scroll(widget: object = None) -> Sheet
```
- Leaving `widget` as `None` unsyncs all previously synced widgets.

#### **See or scroll to a specific cell on the sheet**

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
cell_visible(r: int, c: int) -> bool
```

___

#### **Check if a cell is totally visible**

```python
cell_completely_visible(r: int, c: int, seperate_axes: bool = False) -> bool
```
- `separate_axes` returns tuple of bools e.g. `(cell y axis is visible, cell x axis is visible)`

___

```python
set_xview(position: float, option: str = "moveto") -> Sheet
```

___

```python
set_yview(position: float, option: str = "moveto") -> Sheet
```

___

```python
get_xview() -> tuple[float, float]
```

___

```python
get_yview() -> tuple[float, float]
```

___

```python
set_view(x_args: [str, float], y_args: [str, float]) -> Sheet
```

---
# **Hiding Columns**

Note that once you have hidden columns you can use the function `displayed_column_to_data(column)` to retrieve a column data index from a displayed index.

#### **Display only certain columns**

```python
display_columns(
    columns: None | Literal["all"] | Iterator[int] = None,
    all_columns_displayed: None | bool = None,
    reset_col_positions: bool = True,
    refresh: bool = False,
    redraw: bool = False,
    deselect_all: bool = True,
    **kwargs,
) -> list[int] | None
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
all_columns_displayed(a: bool | None = None) -> bool
```
- `a` (`bool`, `None`) Either set by using `bool` or get by leaving `None` e.g. `all_columns_displayed()`.

___

#### **Hide specific columns**

```python
hide_columns(
    columns: int | set | Iterator[int] = set(),
    redraw: bool = True,
    deselect_all: bool = True,
) -> Sheet
```
- **NOTE**: `columns` (`int`) uses displayed column indexes, not data indexes. In other words the indexes of the columns displayed on the screen are the ones that are hidden, this is useful when used in conjunction with `get_selected_columns()`.

___

#### **Displayed column index to data**

Convert a displayed column index to a data index. If the internal `all_columns_displayed` attribute is `True` then it will simply return the provided argument.
```python
displayed_column_to_data(c)
```

___

#### **Get currently displayed columns**

```python
@property
displayed_columns() -> list[int]
```

---
# **Hiding Rows**

Note that once you have hidden rows you can use the function `displayed_row_to_data(row)` to retrieve a row data index from a displayed index.

#### **Display only certain rows**

```python
display_rows(
    rows: None | Literal["all"] | Iterator[int] = None,
    all_rows_displayed: None | bool = None,
    reset_row_positions: bool = True,
    refresh: bool = False,
    redraw: bool = False,
    deselect_all: bool = True,
    **kwargs,
) -> list[int] | None
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

#### **Hide specific rows**

```python
hide_rows(
    rows: int | set | Iterator[int] = set(),
    redraw: bool = True,
    deselect_all: bool = True,
) -> Sheet
```
- **NOTE**: `rows` (`int`) uses displayed row indexes, not data indexes. In other words the indexes of the rows displayed on the screen are the ones that are hidden, this is useful when used in conjunction with `get_selected_rows()`.

___

#### **Get all rows displayed boolean**

```python
all_rows_displayed(a: bool | None = None) -> bool
```
- `a` (`bool`, `None`) Either set by using `bool` or get by leaving `None` e.g. `all_rows_displayed()`.

___

#### **Displayed row index to data**

Convert a displayed row index to a data index. If the internal `all_rows_displayed` attribute is `True` then it will simply return the provided argument.
```python
displayed_row_to_data(r)
```

___

#### **Get currently displayed rows**

```python
@property
displayed_rows() -> list[int]
```

---
# **Hiding Sheet Elements**

#### **Hide parts of the Sheet or all of it**

```python
hide(
    canvas: Literal[
        "all",
        "row_index",
        "header",
        "top_left",
        "x_scrollbar",
        "y_scrollbar",
    ] = "all",
) -> Sheet
```
- `canvas` (`str`) options are `all`, `row_index`, `header`, `top_left`, `x_scrollbar`, `y_scrollbar`
	- `all` hides the entire table and is the default.

___

#### **Show parts of the Sheet or all of it**

```python
show(
    canvas: Literal[
        "all",
        "row_index",
        "header",
        "top_left",
        "x_scrollbar",
        "y_scrollbar",
    ] = "all",
) -> Sheet
```
- `canvas` (`str`) options are `all`, `row_index`, `header`, `top_left`, `x_scrollbar`, `y_scrollbar`
	- `all` shows the entire table and is the default.

---
# **Sheet Height and Width**

#### **Modify widget height and width in pixels**

```python
height_and_width(
    height: int | None = None,
    width: int | None = None,
) -> Sheet
```
- `height` (`int`, `None`) set a height in pixels.
- `width` (`int`, `None`) set a width in pixels.
If both arguments are `None` then table will reset to default tkinter canvas dimensions.

___

```python
get_frame_y(y: int) -> int
```
- Adds the height of the Sheets header to a y position.

___

```python
get_frame_x(x: int) -> int
```
- Adds the width of the Sheets index to an x position.

---
# **Cell Text Editor**

#### **Open the currently selected cell in the main table**

```python
open_cell(ignore_existing_editor: bool = True) -> Sheet
```
- Function utilises the currently selected cell in the main table, even if a column/row is selected, to open a non selected cell first use `set_currently_selected()` to set the cell to open.

___

#### **Open the currently selected cell but in the header**

```python
open_header_cell(ignore_existing_editor: bool = True) -> Sheet
```
- Also uses currently selected cell, which you can set with `set_currently_selected()`.

___

#### **Open the currently selected cell but in the index**

```python
open_index_cell(ignore_existing_editor: bool = True) -> Sheet
```
- Also uses currently selected cell, which you can set with `set_currently_selected()`.

___

```python
set_text_editor_value(
    text: str = "",
    r: int | None = None,
    c: int | None = None,
) -> Sheet
```

___

#### **Close any existing text editor**

```python
close_text_editor(set_data: bool = True) -> Sheet
```
Notes:
- Also closes any existing `"normal"` state dropdown box.

Parameters:
- `set_data` (`bool`) when `True` sets the cell data to the text editor value (if it is valid). When `False` the text editor is simply destroyed.

___

#### **Get an existing text editors value**

```python
get_text_editor_value() -> str | None
```
Notes:
- `None` is returned if no text editor exists, a `str` of the text editors value will be returned if it does.

___

```python
destroy_text_editor(event: object = None) -> Sheet
```

___

```python
get_text_editor_widget(event: object = None) -> tk.Text | None
```

___

```python
bind_key_text_editor(key: str, function: Callable) -> Sheet
```

___

```python
unbind_key_text_editor(key: str) -> Sheet
```

---
# **Sheet Options and Other Functions**

```python
set_options(redraw: bool = True, **kwargs) -> Sheet
```

The list of key word arguments available for `set_options()` are as follows, [see here](https://github.com/ragardner/tksheet/wiki/Version-7#initialization-options) as a guide for argument types.
```python
auto_resize_columns
auto_resize_rows
to_clipboard_delimiter
to_clipboard_quotechar
to_clipboard_lineterminator
from_clipboard_delimiters
show_dropdown_borders
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
default_row_height
default_column_width
default_header_height
default_row_index_width
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

# for changing the in-built right click menus labels
# use a string as an argument
edit_header_label
edit_header_accelerator
edit_index_label
edit_index_accelerator
edit_cell_label
edit_cell_accelerator
cut_label
cut_accelerator
cut_contents_label
cut_contents_accelerator
copy_label
copy_accelerator
copy_contents_label
copy_contents_accelerator
paste_label
paste_accelerator
delete_label
delete_accelerator
clear_contents_label
clear_contents_accelerator
delete_columns_label
delete_columns_accelerator
insert_columns_left_label
insert_columns_left_accelerator
insert_column_label
insert_column_accelerator
insert_columns_right_label
insert_columns_right_accelerator
delete_rows_label
delete_rows_accelerator
insert_rows_above_label
insert_rows_above_accelerator
insert_rows_below_label
insert_rows_below_accelerator
insert_row_label
insert_row_accelerator
select_all_label
select_all_accelerator
undo_label
undo_accelerator

# for changing the keyboard bindings for copy, paste, etc.
# use a list of strings as an argument
copy_bindings
cut_bindings
paste_bindings
undo_bindings
redo_bindings
delete_bindings
select_all_bindings
tab_bindings
up_bindings
right_bindings
down_bindings
left_bindings
prior_bindings
next_bindings
```
Notes:
- A dictionary can be provided instead of using the keyword arguments:

```python
kwargs = {
    "copy_bindings": [
        "<Control-g>",
        "<Control-G>",
    ],
    "cut_bindings": [
        "<Control-c>",
        "<Control-C>",
    ],
}
sheet.set_options(**kwargs)
```

___

Get internal storage dictionary of highlights, readonly cells, dropdowns etc. Specifically for cell options.
```python
get_cell_options(key: None | str = None, canvas: Literal["table", "row_index", "header"] = "table") -> dict
```

___

Get internal storage dictionary of highlights, readonly rows, dropdowns etc. Specifically for row options.
```python
get_row_options(key: None | str = None) -> dict
```

___

Get internal storage dictionary of highlights, readonly columns, dropdowns etc. Specifically for column options.
```python
get_column_options(key: None | str = None) -> dict
```

___

Get internal storage dictionary of highlights, readonly header cells, dropdowns etc. Specifically for header options.
```python
get_header_options(key: None | str = None) -> dict
```

___

Get internal storage dictionary of highlights, readonly row index cells, dropdowns etc. Specifically for row index options.
```python
get_index_options(key: None | str = None) -> dict
```

___

Delete any formats, alignments, dropdown boxes, check boxes, highlights etc. that are larger than the sheets currently held data, includes row index and header in measurement of dimensions.
```python
del_out_of_bounds_options() -> Sheet
```

___

Delete all alignments, dropdown boxes, check boxes, highlights etc.
```python
reset_all_options() -> Sheet
```

___

Flash a dashed box of chosen dimensions.
```python
show_ctrl_outline(
    canvas: Literal["table"] = "table",
    start_cell: tuple[int, int] = (0, 0),
    end_cell: tuple[int, int] = (1, 1),
) -> Sheet
```

___

Various functions related to the Sheets internal undo and redo stacks.
```python
# clears both undos and redos
reset_undos() -> Sheet

# get the Sheets modifiable deque variables which store changes for undo and redo
get_undo_stack() -> deque
get_redo_stack() -> deque

# set the Sheets undo and redo stacks, returns Sheet widget
set_undo_stack(stack: deque) -> Sheet
set_redo_stack(stack: deque) -> Sheet
```

___

Refresh the table.
```python
refresh(redraw_header: bool = True, redraw_row_index: bool = True) -> Sheet
```

___

Refresh the table.
```python
redraw(redraw_header: bool = True, redraw_row_index: bool = True) -> Sheet
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
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.frame = tk.Frame(self)
        self.frame.grid_columnconfigure(0, weight=1)
        self.frame.grid_rowconfigure(0, weight=1)
        self.sheet = Sheet(self.frame,
                           data=[[f"Row {r}, Column {c}\nnewline1\nnewline2" for c in range(50)] for r in range(500)])
        self.sheet.enable_bindings(
            "single_select",
            "drag_select",
            "edit_cell",
            "paste",
            "cut",
            "copy",
            "delete",
            "select_all",
            "column_select",
            "row_select",
            "column_width_resize",
            "double_click_column_resize",
            "arrowkeys",
            "row_height_resize",
            "double_click_row_resize",
            "right_click_popup_menu",
            "rc_select",
        )
        self.sheet.extra_bindings("begin_edit_cell", self.begin_edit_cell)
        self.sheet.edit_validation(self.validate_edits)
        self.sheet.popup_menu_add_command("Say Hello", self.new_right_click_button)
        self.frame.grid(row=0, column=0, sticky="nswe")
        self.sheet.grid(row=0, column=0, sticky="nswe")

    def new_right_click_button(self, event=None):
        print ("Hello World!")

    def begin_edit_cell(self, event=None):
        return event.value

    def validate_edits(self, event):
        # remove spaces from any cell edits, including paste
        if isinstance(event.value, str) and event.value:
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
        self.sheet.enable_bindings("all", "edit_index", "edit_header")
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
            self.sheet.reset()
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
from tksheet import (
    Sheet,
    formatter,
    float_formatter,
    int_formatter,
    percentage_formatter,
    bool_formatter,
    truthy,
    falsy,
    num2alpha,
)
import tkinter as tk
from datetime import datetime, date
from dateutil import parser, tz
from math import ceil
import re

date_replace = re.compile("|".join(["\(", "\)", "\[", "\]", "\<", "\>"]))


# Custom formatter methods
def round_up(x):
    try:  # might not be a number if empty
        return float(ceil(x))
    except Exception:
        return x


def only_numeric(s):
    return "".join(n for n in f"{s}" if n.isnumeric() or n == ".")


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
        except Exception:
            raise ValueError(f"Could not parse {dt} as a datetime")
    if dt.tzinfo is None:
        dt.replace(tzinfo=tz.tzlocal())
    dt = dt.astimezone(tz.tzlocal())
    return dt.replace(tzinfo=None)


def datetime_to_string(dt: datetime, **kwargs):
    return dt.strftime("%d %b, %Y, %H:%M:%S")


# Custom Formatter with additional kwargs


def custom_datetime_to_str(dt: datetime, **kwargs):
    return dt.strftime(kwargs["format"])


class demo(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.frame = tk.Frame(self)
        self.frame.grid_columnconfigure(0, weight=1)
        self.frame.grid_rowconfigure(0, weight=1)
        self.sheet = Sheet(self.frame, empty_vertical=0, empty_horizontal=0, data=[[f"{r}"] * 11 for r in range(20)])
        self.sheet.enable_bindings()
        self.frame.grid(row=0, column=0, sticky="nswe")
        self.sheet.grid(row=0, column=0, sticky="nswe")
        self.sheet.headers(
            [
                "Non-Nullable Float Cell\n1 decimals places",
                "Float Cell",
                "Int Cell",
                "Bool Cell",
                "Percentage Cell\n0 decimal places",
                "Custom Datetime Cell",
                "Custom Datetime Cell\nCustom Format String",
                "Float Cell that\nrounds up",
                "Float cell that\n strips non-numeric",
                "Dropdown Over Nullable\nPercentage Cell",
                "Percentage Cell\n2 decimal places",
            ]
        )

        # num2alpha converts column integer to letter

        # Some examples of data formatting
        self.sheet[num2alpha(0)].format(float_formatter(nullable=False))
        self.sheet[num2alpha(1)].format(float_formatter())
        self.sheet[num2alpha(2)].format(int_formatter())
        self.sheet[num2alpha(3)].format(bool_formatter(truthy=truthy | {"nah yeah"}, falsy=falsy | {"yeah nah"}))
        self.sheet[num2alpha(4)].format(percentage_formatter())

        # Custom Formatters
        # Custom using generic formatter interface
        self.sheet[num2alpha(5)].format(
            formatter(
                datatypes=datetime,
                format_function=convert_to_local_datetime,
                to_str_function=datetime_to_string,
                nullable=False,
                invalid_value="NaT",
            )
        )

        # Custom format
        self.sheet[num2alpha(6)].format(
            datatypes=datetime,
            format_function=convert_to_local_datetime,
            to_str_function=custom_datetime_to_str,
            nullable=True,
            invalid_value="NaT",
            format="(%Y-%m-%d) %H:%M %p",
        )

        # Unique cell behaviour using the post_conversion_function
        self.sheet[num2alpha(7)].format(float_formatter(post_format_function=round_up))
        self.sheet[num2alpha(8)].format(float_formatter(), pre_format_function=only_numeric)
        self.sheet[num2alpha(9)].dropdown(values=["", "104%", 0.24, "300%", "not a number"], set_value=1,)
        self.sheet[num2alpha(9)].format(percentage_formatter(), decimals=0)
        self.sheet[num2alpha(10)].format(percentage_formatter(decimals=5))


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

