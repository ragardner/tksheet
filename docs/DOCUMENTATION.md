# **About tksheet**

- `tksheet` is a Python tkinter table and treeview widget written in pure python.
- It is licensed under the [MIT license](https://github.com/ragardner/tksheet/blob/master/LICENSE.txt).
- It works by using tkinter canvases and moving lines, text and rectangles around for only the visible portion of the table.
- If you are using a version of tksheet that is older than `7.0.0` then you will need the documentation [here](https://github.com/ragardner/tksheet/wiki/Version-6) instead.
    - In tksheet versions >= `7.0.2` the current version will be at the top of the file `__init__.py`.

### **Known Issues**

1. When using `edit_validation()` to validate cell edits and pasting into the sheet:
    - If the sheets rows are expanded then the row numbers under the key `row` and `loc` are data indexes whereas the column numbers are displayed indexes.
    - If the sheets columns are expanded then the column numbers under the key `column` and `loc` are data indexes whereas the row numbers are displayed indexes.
    - This is only relevant when there are hidden rows or columns and you're using `edit_validation()` and you're using the event data keys `row`, `column` or `loc` in your bound `edit_validation()` function and you're using paste and the sheet can be expanded by paste.
2. There may be some issues with toggle select mode and deselection.

### **Limitations**

Some examples of things that are not possible with tksheet:

- Due to the limitations of the Tkinter Canvas right-to-left (RTL) languages are not supported.
- Cell merging.
- Changing font for individual cells.
- Mouse drag copy cells.
- Cell highlight borders.
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

This is to demonstrate some of tksheets functionality:

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
        # #validate-user-cell-edits
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
            # #setting-sheet-data
            self.sheet[box].options(undo=True).data = "Hello World!"
            # highlight the cell for 2 seconds
            self.highlight_area(box)

    def print_data(self):
        for box in self.sheet.get_all_selection_boxes():
            # get user selected area sheet data
            # more information at:
            # #getting-sheet-data
            data = self.sheet[box].data
            for row in data:
                print(row)

    def reset(self):
        # overwrites sheet data, more information at:
        # #setting-sheet-data
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
        # #event-data

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
        # #highlighting-cells
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
    show_top_left: bool | None = None,
    show_row_index: bool = True,
    show_header: bool = True,
    show_x_scrollbar: bool = True,
    show_y_scrollbar: bool = True,
    width: int | None = None,
    height: int | None = None,
    headers: None | list[Any] = None,
    header: None | list[Any] = None,
    row_index: None | list[Any] = None,
    index: None | list[Any] = None,
    default_header: Literal["letters", "numbers", "both"] | None = "letters",
    default_row_index: Literal["letters", "numbers", "both"] | None = "numbers",
    data_reference: None | Sequence[Sequence[Any]] = None,
    data: None | Sequence[Sequence[Any]] = None,
    # either (start row, end row, "rows"), (start column, end column, "rows") or
    # (cells start row, cells start column, cells end row, cells end column, "cells")  # noqa: E501
    startup_select: tuple[int, int, str] | tuple[int, int, int, int, str] = None,
    startup_focus: bool = True,
    total_columns: int | None = None,
    total_rows: int | None = None,
    default_column_width: int = 120,
    default_header_height: str | int = "1",
    default_row_index_width: int = 70,
    default_row_height: str | int = "1",
    min_column_width: int = 1,
    max_column_width: float = float("inf"),
    max_row_height: float = float("inf"),
    max_header_height: float = float("inf"),
    max_index_width: float = float("inf"),
    after_redraw_time_ms: int = 16,
    set_all_heights_and_widths: bool = False,
    zoom: int = 100,
    align: str = "nw",
    header_align: str = "n",
    row_index_align: str | None = None,
    index_align: str = "n",
    displayed_columns: list[int] | None = None,
    all_columns_displayed: bool = True,
    displayed_rows: list[int] | None = None,
    all_rows_displayed: bool = True,
    to_clipboard_delimiter: str = "\t",
    to_clipboard_quotechar: str = '"',
    to_clipboard_lineterminator: str = "\n",
    from_clipboard_delimiters: list[str] | str = "\t",
    show_default_header_for_empty: bool = True,
    show_default_index_for_empty: bool = True,
    page_up_down_select_row: bool = True,
    paste_can_expand_x: bool = False,
    paste_can_expand_y: bool = False,
    paste_insert_column_limit: int | None = None,
    paste_insert_row_limit: int | None = None,
    show_dropdown_borders: bool = False,
    arrow_key_down_right_scroll_page: bool = False,
    cell_auto_resize_enabled: bool = True,
    auto_resize_row_index: bool | Literal["empty"] = "empty",
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
    edit_cell_tab: Literal["right", "down", ""] = "right",
    edit_cell_return: Literal["right", "down", ""] = "down",
    editor_del_key: Literal["forward", "backward", ""] = "forward",
    treeview: bool = False,
    treeview_indent: str | int = "2",
    rounded_boxes: bool = True,
    alternate_color: str = "",
    allow_cell_overflow: bool = False,
    # "" no wrap, "w" word wrap, "c" char wrap
    table_wrap: Literal["", "w", "c"] = "c",
    index_wrap: Literal["", "w", "c"] = "c",
    header_wrap: Literal["", "w", "c"] = "c",
    sort_key: Callable = natural_sort_key,
    # colors
    outline_thickness: int = 0,
    theme: str = "light blue",
    outline_color: str = theme_light_blue["outline_color"],
    frame_bg: str = theme_light_blue["table_bg"],
    popup_menu_fg: str = theme_light_blue["popup_menu_fg"],
    popup_menu_bg: str = theme_light_blue["popup_menu_bg"],
    popup_menu_highlight_bg: str = theme_light_blue["popup_menu_highlight_bg"],
    popup_menu_highlight_fg: str = theme_light_blue["popup_menu_highlight_fg"],
    table_grid_fg: str = theme_light_blue["table_grid_fg"],
    table_bg: str = theme_light_blue["table_bg"],
    table_fg: str = theme_light_blue["table_fg"],
    table_editor_bg: str = theme_light_blue["table_editor_bg"],
    table_editor_fg: str = theme_light_blue["table_editor_fg"],
    table_editor_select_bg: str = theme_light_blue["table_editor_select_bg"],
    table_editor_select_fg: str = theme_light_blue["table_editor_select_fg"],
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
    index_editor_bg: str = theme_light_blue["index_editor_bg"],
    index_editor_fg: str = theme_light_blue["index_editor_fg"],
    index_editor_select_bg: str = theme_light_blue["index_editor_select_bg"],
    index_editor_select_fg: str = theme_light_blue["index_editor_select_fg"],
    index_selected_cells_bg: str = theme_light_blue["index_selected_cells_bg"],
    index_selected_cells_fg: str = theme_light_blue["index_selected_cells_fg"],
    index_selected_rows_bg: str = theme_light_blue["index_selected_rows_bg"],
    index_selected_rows_fg: str = theme_light_blue["index_selected_rows_fg"],
    index_hidden_rows_expander_bg: str = theme_light_blue["index_hidden_rows_expander_bg"],
    header_bg: str = theme_light_blue["header_bg"],
    header_border_fg: str = theme_light_blue["header_border_fg"],
    header_grid_fg: str = theme_light_blue["header_grid_fg"],
    header_fg: str = theme_light_blue["header_fg"],
    header_editor_bg: str = theme_light_blue["header_editor_bg"],
    header_editor_fg: str = theme_light_blue["header_editor_fg"],
    header_editor_select_bg: str = theme_light_blue["header_editor_select_bg"],
    header_editor_select_fg: str = theme_light_blue["header_editor_select_fg"],
    header_selected_cells_bg: str = theme_light_blue["header_selected_cells_bg"],
    header_selected_cells_fg: str = theme_light_blue["header_selected_cells_fg"],
    header_selected_columns_bg: str = theme_light_blue["header_selected_columns_bg"],
    header_selected_columns_fg: str = theme_light_blue["header_selected_columns_fg"],
    header_hidden_columns_expander_bg: str = theme_light_blue["header_hidden_columns_expander_bg"],
    top_left_bg: str = theme_light_blue["top_left_bg"],
    top_left_fg: str = theme_light_blue["top_left_fg"],
    top_left_fg_highlight: str = theme_light_blue["top_left_fg_highlight"],
    vertical_scroll_background: str = theme_light_blue["vertical_scroll_background"],
    horizontal_scroll_background: str = theme_light_blue["horizontal_scroll_background"],
    vertical_scroll_troughcolor: str = theme_light_blue["vertical_scroll_troughcolor"],
    horizontal_scroll_troughcolor: str = theme_light_blue["horizontal_scroll_troughcolor"],
    vertical_scroll_lightcolor: str = theme_light_blue["vertical_scroll_lightcolor"],
    horizontal_scroll_lightcolor: str = theme_light_blue["horizontal_scroll_lightcolor"],
    vertical_scroll_darkcolor: str = theme_light_blue["vertical_scroll_darkcolor"],
    horizontal_scroll_darkcolor: str = theme_light_blue["horizontal_scroll_darkcolor"],
    vertical_scroll_relief: str = theme_light_blue["vertical_scroll_relief"],
    horizontal_scroll_relief: str = theme_light_blue["horizontal_scroll_relief"],
    vertical_scroll_troughrelief: str = theme_light_blue["vertical_scroll_troughrelief"],
    horizontal_scroll_troughrelief: str = theme_light_blue["horizontal_scroll_troughrelief"],
    vertical_scroll_bordercolor: str = theme_light_blue["vertical_scroll_bordercolor"],
    horizontal_scroll_bordercolor: str = theme_light_blue["horizontal_scroll_bordercolor"],
    vertical_scroll_active_bg: str = theme_light_blue["vertical_scroll_active_bg"],
    horizontal_scroll_active_bg: str = theme_light_blue["horizontal_scroll_active_bg"],
    vertical_scroll_not_active_bg: str = theme_light_blue["vertical_scroll_not_active_bg"],
    horizontal_scroll_not_active_bg: str = theme_light_blue["horizontal_scroll_not_active_bg"],
    vertical_scroll_pressed_bg: str = theme_light_blue["vertical_scroll_pressed_bg"],
    horizontal_scroll_pressed_bg: str = theme_light_blue["horizontal_scroll_pressed_bg"],
    vertical_scroll_active_fg: str = theme_light_blue["vertical_scroll_active_fg"],
    horizontal_scroll_active_fg: str = theme_light_blue["horizontal_scroll_active_fg"],
    vertical_scroll_not_active_fg: str = theme_light_blue["vertical_scroll_not_active_fg"],
    horizontal_scroll_not_active_fg: str = theme_light_blue["horizontal_scroll_not_active_fg"],
    vertical_scroll_pressed_fg: str = theme_light_blue["vertical_scroll_pressed_fg"],
    horizontal_scroll_pressed_fg: str = theme_light_blue["horizontal_scroll_pressed_fg"],
    vertical_scroll_borderwidth: int = 1,
    horizontal_scroll_borderwidth: int = 1,
    vertical_scroll_gripcount: int = 0,
    horizontal_scroll_gripcount: int = 0,
    scrollbar_theme_inheritance: str = "default",
    scrollbar_show_arrows: bool = True,
    # changing the arrowsize (width) of the scrollbars
    # is not working with 'default' theme
    # use 'clam' theme instead if you want to change the width
    vertical_scroll_arrowsize: str | int = "",
    horizontal_scroll_arrowsize: str | int = "",
    **kwargs,
) -> None
```
- `name` setting a name for the sheet is useful when you have multiple sheets and you need to determine which one an event came from.
- `auto_resize_columns` (`int`, `None`) if set as an `int` the columns will automatically resize to fit the width of the window, the `int` value being the minimum of each column in pixels.
- `auto_resize_rows` (`int`, `None`) if set as an `int` the rows will automatically resize to fit the height of the window, the `int` value being the minimum height of each row in pixels.
- `startup_select` selects cells, rows or columns at initialization by using a `tuple` e.g. `(0, 0, "cells")` for cell A0 or `(0, 5, "rows")` for rows 0 to 5.
- `data_reference` and `data` are essentially the same.
- `row_index` and `index` are the same, `index` takes priority, same as with `headers` and `header`.
- `startup_select` either `(start row, end row, "rows")`, `(start column, end column, "rows")` or `(start row, start column, end row, end column, "cells")`. The start/end row/column variables need to be `int`s.
- `auto_resize_row_index` either `True`, `False` or `"empty"`.
    - `"empty"` it will only automatically resize if the row index is empty.
    - `True` it will always automatically resize.
    - `False` it will never automatically resize.
- If `show_selected_cells_border` is `False` then the colors for `table_selected_box_cells_fg` / `table_selected_box_rows_fg` / `table_selected_box_columns_fg` will be used for the currently selected cells background.
- Only set `show_top_left` to `True` if you want to always show the top left rectangle of the sheet. Leave as `None` to only show it when both the index and header are showing.
- For help with `treeview` mode see [here](#treeview-mode).

You can change most of these settings after initialization using the [`set_options()` function](#sheet-options).
- `scrollbar_theme_inheritance` and `scrollbar_show_arrows` will only work on `Sheet()` initialization, not with `set_options()`

---
# **Sheet Appearance**

### **Sheet Colors**

To change the colors of individual cells, rows or columns use the functions listed under [highlighting cells](#highlighting-cells).

For the colors of specific parts of the table such as gridlines and backgrounds use the function `set_options()`, keyword arguments specific to sheet colors are listed below. All the other `set_options()` arguments can be found [here](#sheet-options).

Use a tkinter color or a hex string e.g.

```python
my_sheet_widget.set_options(table_bg="black")
my_sheet_widget.set_options(table_bg="#000000")
my_sheet_widget.set_options(horizontal_scroll_pressed_bg="red")
```

#### **Set options**

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

# scroll bars
vertical_scroll_background
horizontal_scroll_background
vertical_scroll_troughcolor
horizontal_scroll_troughcolor
vertical_scroll_lightcolor
horizontal_scroll_lightcolor
vertical_scroll_darkcolor
horizontal_scroll_darkcolor
vertical_scroll_bordercolor
horizontal_scroll_bordercolor
vertical_scroll_active_bg
horizontal_scroll_active_bg
vertical_scroll_not_active_bg
horizontal_scroll_not_active_bg
vertical_scroll_pressed_bg
horizontal_scroll_pressed_bg
vertical_scroll_active_fg
horizontal_scroll_active_fg
vertical_scroll_not_active_fg
horizontal_scroll_not_active_fg
vertical_scroll_pressed_fg
horizontal_scroll_pressed_fg
)
```

Otherwise you can change the theme using the below function.
```python
change_theme(theme: str = "light blue", redraw: bool = True) -> Sheet
```
- `theme` (`str`) options (themes) are currently `"light blue"`, `"light green"`, `"dark"`, `"black"`, `"dark blue"` and `"dark green"`.

### **Scrollbar Appearance**

**Scrollbar colors:**

The above [function and keyword arguments](#set-options) can be used to change the colors of the scroll bars.

**Scrollbar relief, size, arrows, etc.**

Some scroll bar style options can only be changed on `Sheet()` initialization, others can be changed whenever using `set_options()`:

- Options that can only be set in the `= Sheet(...)` initialization:
    - `scrollbar_theme_inheritance: str = "default"`
        - This is which tkinter theme to inherit the new style from, changing the width of the scroll bar might not work with the `"default"` theme. If this is the case try using `"clam"` instead.
    - `scrollbar_show_arrows: bool`
        - When `False` the scroll bars arrow buttons on either end will be hidden, this may effect the width of the scroll bar.
- Options that can be set using `set_options()` also:
    - `vertical_scroll_borderwidth: int`
    - `horizontal_scroll_borderwidth: int`
    - `vertical_scroll_gripcount: int`
    - `horizontal_scroll_gripcount: int`
    - `vertical_scroll_arrowsize: str | int`
    - `horizontal_scroll_arrowsize: str | int`

### **Alternate Row Colors**

For basic alternate row colors in the main table either:

- Use the `Sheet()` initialization keyword argument `alternate_color` (`str`) or
- Use the `set_options()` function with the keyword argument `alternate_color`

Examples:

```python
set_options(alternate_color="#E2EAF4")
```

```python
my_sheet = Sheet(parent, alternate_color="gray80")
```

Note that any cell, row or column highlights will display over alternate row colors.

### **Refreshing the sheet**

Refresh the table.
```python
refresh(redraw_header: bool = True, redraw_row_index: bool = True) -> Sheet
```

___

Refresh the table.
```python
redraw(redraw_header: bool = True, redraw_row_index: bool = True) -> Sheet
```

___

Refresh after idle (prevents multiple redraws).
```python
set_refresh_timer(
    redraw: bool = True,
    index: bool = True,
    header: bool = True,
) -> Sheet
```

---
# **Sheet Options**

```python
set_options(redraw: bool = True, **kwargs) -> Sheet
```

Key word arguments available for `set_options()` (values are defaults):

```python
"popup_menu_fg": "#000000",
"popup_menu_bg": "#FFFFFF",
"popup_menu_highlight_bg": "#DCDEE0",
"popup_menu_highlight_fg": "#000000",
"index_hidden_rows_expander_bg": "#747775",
"header_hidden_columns_expander_bg": "#747775",
"header_bg": "#FFFFFF",
"header_border_fg": "#C4C7C5",
"header_grid_fg": "#C4C7C5",
"header_fg": "#444746",
"header_editor_bg": "#FFFFFF",
"header_editor_fg": "#444746",
"header_editor_select_bg": "#cfd1d1",
"header_editor_select_fg": "#000000",
"header_selected_cells_bg": "#D3E3FD",
"header_selected_cells_fg": "black",
"index_bg": "#FFFFFF",
"index_border_fg": "#C4C7C5",
"index_grid_fg": "#C4C7C5",
"index_fg": "black",
"index_editor_bg": "#FFFFFF",
"index_editor_fg": "black",
"index_editor_select_bg": "#cfd1d1",
"index_editor_select_fg": "#000000",
"index_selected_cells_bg": "#D3E3FD",
"index_selected_cells_fg": "black",
"top_left_bg": "#F9FBFD",
"top_left_fg": "#d9d9d9",
"top_left_fg_highlight": "#747775",
"table_bg": "#FFFFFF",
"table_grid_fg": "#E1E1E1",
"table_fg": "black",
"table_editor_bg": "#FFFFFF",
"table_editor_fg": "black",
"table_editor_select_bg": "#cfd1d1",
"table_editor_select_fg": "#000000",
"table_selected_box_cells_fg": "#0B57D0",
"table_selected_box_rows_fg": "#0B57D0",
"table_selected_box_columns_fg": "#0B57D0",
"table_selected_cells_border_fg": "#0B57D0",
"table_selected_cells_bg": "#E6EFFD",
"table_selected_cells_fg": "black",
"resizing_line_fg": "black",
"drag_and_drop_bg": "#0B57D0",
"outline_color": "gray2",
"header_selected_columns_bg": "#0B57D0",
"header_selected_columns_fg": "#FFFFFF",
"index_selected_rows_bg": "#0B57D0",
"index_selected_rows_fg": "#FFFFFF",
"table_selected_rows_border_fg": "#0B57D0",
"table_selected_rows_bg": "#E6EFFD",
"table_selected_rows_fg": "black",
"table_selected_columns_border_fg": "#0B57D0",
"table_selected_columns_bg": "#E6EFFD",
"table_selected_columns_fg": "black",
"tree_arrow_fg": "black",
"selected_cells_tree_arrow_fg": "black",
"selected_rows_tree_arrow_fg": "#FFFFFF",
"vertical_scroll_background": "#FFFFFF",
"horizontal_scroll_background": "#FFFFFF",
"vertical_scroll_troughcolor": "#f9fbfd",
"horizontal_scroll_troughcolor": "#f9fbfd",
"vertical_scroll_lightcolor": "#FFFFFF",
"horizontal_scroll_lightcolor": "#FFFFFF",
"vertical_scroll_darkcolor": "gray50",
"horizontal_scroll_darkcolor": "gray50",
"vertical_scroll_relief": "flat",
"horizontal_scroll_relief": "flat",
"vertical_scroll_troughrelief": "flat",
"horizontal_scroll_troughrelief": "flat",
"vertical_scroll_bordercolor": "#f9fbfd",
"horizontal_scroll_bordercolor": "#f9fbfd",
"vertical_scroll_active_bg": "#bdc1c6",
"horizontal_scroll_active_bg": "#bdc1c6",
"vertical_scroll_not_active_bg": "#DADCE0",
"horizontal_scroll_not_active_bg": "#DADCE0",
"vertical_scroll_pressed_bg": "#bdc1c6",
"horizontal_scroll_pressed_bg": "#bdc1c6",
"vertical_scroll_active_fg": "#bdc1c6",
"horizontal_scroll_active_fg": "#bdc1c6",
"vertical_scroll_not_active_fg": "#DADCE0",
"horizontal_scroll_not_active_fg": "#DADCE0",
"vertical_scroll_pressed_fg": "#bdc1c6",
"horizontal_scroll_pressed_fg": "#bdc1c6",
"popup_menu_font": FontTuple(
    "Calibri",
    13 if USER_OS == "darwin" else 11,
    "normal",
),
"table_font": FontTuple(
    "Calibri",
    13 if USER_OS == "darwin" else 11,
    "normal",
),
"header_font": FontTuple(
    "Calibri",
    13 if USER_OS == "darwin" else 11,
    "normal",
),
"index_font": FontTuple(
    "Calibri",
    13 if USER_OS == "darwin" else 11,
    "normal",
),
# edit header
"edit_header_label": "Edit header",
"edit_header_accelerator": "",
"edit_header_image": tk.PhotoImage(data=ICON_EDIT),
"edit_header_compound": "left",
# edit index
"edit_index_label": "Edit index",
"edit_index_accelerator": "",
"edit_index_image": tk.PhotoImage(data=ICON_EDIT),
"edit_index_compound": "left",
# edit cell
"edit_cell_label": "Edit cell",
"edit_cell_accelerator": "",
"edit_cell_image": tk.PhotoImage(data=ICON_EDIT),
"edit_cell_compound": "left",
# cut
"cut_label": "Cut",
"cut_accelerator": "Ctrl+X",
"cut_image": tk.PhotoImage(data=ICON_CUT),
"cut_compound": "left",
# cut contents
"cut_contents_label": "Cut contents",
"cut_contents_accelerator": "Ctrl+X",
"cut_contents_image": tk.PhotoImage(data=ICON_CUT),
"cut_contents_compound": "left",
# copy
"copy_label": "Copy",
"copy_accelerator": "Ctrl+C",
"copy_image": tk.PhotoImage(data=ICON_COPY),
"copy_compound": "left",
# copy contents
"copy_contents_label": "Copy contents",
"copy_contents_accelerator": "Ctrl+C",
"copy_contents_image": tk.PhotoImage(data=ICON_COPY),
"copy_contents_compound": "left",
# paste
"paste_label": "Paste",
"paste_accelerator": "Ctrl+V",
"paste_image": tk.PhotoImage(data=ICON_PASTE),
"paste_compound": "left",
# delete
"delete_label": "Delete",
"delete_accelerator": "Del",
"delete_image": tk.PhotoImage(data=ICON_CLEAR),
"delete_compound": "left",
# clear contents
"clear_contents_label": "Clear contents",
"clear_contents_accelerator": "Del",
"clear_contents_image": tk.PhotoImage(data=ICON_CLEAR),
"clear_contents_compound": "left",
# del columns
"delete_columns_label": "Delete columns",
"delete_columns_accelerator": "",
"delete_columns_image": tk.PhotoImage(data=ICON_DEL),
"delete_columns_compound": "left",
# insert columns left
"insert_columns_left_label": "Insert columns left",
"insert_columns_left_accelerator": "",
"insert_columns_left_image": tk.PhotoImage(data=ICON_ADD),
"insert_columns_left_compound": "left",
# insert columns right
"insert_columns_right_label": "Insert columns right",
"insert_columns_right_accelerator": "",
"insert_columns_right_image": tk.PhotoImage(data=ICON_ADD),
"insert_columns_right_compound": "left",
# insert single column
"insert_column_label": "Insert column",
"insert_column_accelerator": "",
"insert_column_image": tk.PhotoImage(data=ICON_ADD),
"insert_column_compound": "left",
# del rows
"delete_rows_label": "Delete rows",
"delete_rows_accelerator": "",
"delete_rows_image": tk.PhotoImage(data=ICON_DEL),
"delete_rows_compound": "left",
# insert rows above
"insert_rows_above_label": "Insert rows above",
"insert_rows_above_accelerator": "",
"insert_rows_above_image": tk.PhotoImage(data=ICON_ADD),
"insert_rows_above_compound": "left",
# insert rows below
"insert_rows_below_label": "Insert rows below",
"insert_rows_below_accelerator": "",
"insert_rows_below_image": tk.PhotoImage(data=ICON_ADD),
"insert_rows_below_compound": "left",
# insert single row
"insert_row_label": "Insert row",
"insert_row_accelerator": "",
"insert_row_image": tk.PhotoImage(data=ICON_ADD),
"insert_row_compound": "left",
# sorting
# labels
"sort_cells_label": "Sort Asc.",
"sort_cells_x_label": "Sort row-wise Asc.",
"sort_row_label": "Sort values Asc.",
"sort_column_label": "Sort values Asc.",
"sort_rows_label": "Sort rows Asc.",
"sort_columns_label": "Sort columns Asc.",
# reverse labels
"sort_cells_reverse_label": "Sort Desc.",
"sort_cells_x_reverse_label": "Sort row-wise Desc.",
"sort_row_reverse_label": "Sort values Desc.",
"sort_column_reverse_label": "Sort values Desc.",
"sort_rows_reverse_label": "Sort rows Desc.",
"sort_columns_reverse_label": "Sort columns Desc.",
# accelerators
"sort_cells_accelerator": "",
"sort_cells_x_accelerator": "",
"sort_row_accelerator": "",
"sort_column_accelerator": "",
"sort_rows_accelerator": "",
"sort_columns_accelerator": "",
# reverse accelerators
"sort_cells_reverse_accelerator": "",
"sort_cells_x_reverse_accelerator": "",
"sort_row_reverse_accelerator": "",
"sort_column_reverse_accelerator": "",
"sort_rows_reverse_accelerator": "",
"sort_columns_reverse_accelerator": "",
# images
"sort_cells_image": tk.PhotoImage(data=ICON_SORT_ASC),
"sort_cells_x_image": tk.PhotoImage(data=ICON_SORT_ASC),
"sort_row_image": tk.PhotoImage(data=ICON_SORT_ASC),
"sort_column_image": tk.PhotoImage(data=ICON_SORT_ASC),
"sort_rows_image": tk.PhotoImage(data=ICON_SORT_ASC),
"sort_columns_image": tk.PhotoImage(data=ICON_SORT_ASC),
# compounds
"sort_cells_compound": "left",
"sort_cells_x_compound": "left",
"sort_row_compound": "left",
"sort_column_compound": "left",
"sort_rows_compound": "left",
"sort_columns_compound": "left",
# reverse images
"sort_cells_reverse_image": tk.PhotoImage(data=ICON_SORT_DESC),
"sort_cells_x_reverse_image": tk.PhotoImage(data=ICON_SORT_DESC),
"sort_row_reverse_image": tk.PhotoImage(data=ICON_SORT_DESC),
"sort_column_reverse_image": tk.PhotoImage(data=ICON_SORT_DESC),
"sort_rows_reverse_image": tk.PhotoImage(data=ICON_SORT_DESC),
"sort_columns_reverse_image": tk.PhotoImage(data=ICON_SORT_DESC),
# reverse compounds
"sort_cells_reverse_compound": "left",
"sort_cells_x_reverse_compound": "left",
"sort_row_reverse_compound": "left",
"sort_column_reverse_compound": "left",
"sort_rows_reverse_compound": "left",
"sort_columns_reverse_compound": "left",
# select all
"select_all_label": "Select all",
"select_all_accelerator": "Ctrl+A",
"select_all_image": tk.PhotoImage(data=ICON_SELECT_ALL),
"select_all_compound": "left",
# undo
"undo_label": "Undo",
"undo_accelerator": "Ctrl+Z",
"undo_image": tk.PhotoImage(data=ICON_UNDO),
"undo_compound": "left",
# redo
"redo_label": "Redo",
"redo_accelerator": "Ctrl+Shift+Z",
"redo_image": tk.PhotoImage(data=ICON_REDO),
"redo_compound": "left",
# bindings
"copy_bindings": [
    f"<{ctrl_key}-c>",
    f"<{ctrl_key}-C>",
],
"cut_bindings": [
    f"<{ctrl_key}-x>",
    f"<{ctrl_key}-X>",
],
"paste_bindings": [
    f"<{ctrl_key}-v>",
    f"<{ctrl_key}-V>",
],
"undo_bindings": [
    f"<{ctrl_key}-z>",
    f"<{ctrl_key}-Z>",
],
"redo_bindings": [
    f"<{ctrl_key}-Shift-z>",
    f"<{ctrl_key}-Shift-Z>",
],
"delete_bindings": [
    "<Delete>",
],
"select_all_bindings": [
    f"<{ctrl_key}-a>",
    f"<{ctrl_key}-A>",
    f"<{ctrl_key}-Shift-space>",
],
"select_columns_bindings": [
    "<Control-space>",
],
"select_rows_bindings": [
    "<Shift-space>",
],
"row_start_bindings": [
    "<Command-Left>",
    "<Home>",
]
if USER_OS == "darwin"
else ["<Home>"],
"table_start_bindings": [
    f"<{ctrl_key}-Home>",
],
"tab_bindings": [
    "<Tab>",
],
"up_bindings": [
    "<Up>",
],
"right_bindings": [
    "<Right>",
],
"down_bindings": [
    "<Down>",
],
"left_bindings": [
    "<Left>",
],
"shift_up_bindings": [
    "<Shift-Up>",
],
"shift_right_bindings": [
    "<Shift-Right>",
],
"shift_down_bindings": [
    "<Shift-Down>",
],
"shift_left_bindings": [
    "<Shift-Left>",
],
"prior_bindings": [
    "<Prior>",
],
"next_bindings": [
    "<Next>",
],
"find_bindings": [
    f"<{ctrl_key}-f>",
    f"<{ctrl_key}-F>",
],
"find_next_bindings": [
    f"<{ctrl_key}-g>",
    f"<{ctrl_key}-G>",
],
"find_previous_bindings": [
    f"<{ctrl_key}-Shift-g>",
    f"<{ctrl_key}-Shift-G>",
],
"toggle_replace_bindings": [
    f"<{ctrl_key}-h>",
    f"<{ctrl_key}-H>",
],
"escape_bindings": [
    "<Escape>",
],
# other
"vertical_scroll_borderwidth": 1,
"horizontal_scroll_borderwidth": 1,
"vertical_scroll_gripcount": 0,
"horizontal_scroll_gripcount": 0,
"vertical_scroll_arrowsize": "",
"horizontal_scroll_arrowsize": "",
"set_cell_sizes_on_zoom": False,
"auto_resize_columns": None,
"auto_resize_rows": None,
"to_clipboard_delimiter": "\t",
"to_clipboard_quotechar": '"',
"to_clipboard_lineterminator": "\n",
"from_clipboard_delimiters": ["\t"],
"show_dropdown_borders": False,
"show_default_header_for_empty": True,
"show_default_index_for_empty": True,
"default_header_height": "1",
"default_row_height": "1",
"default_column_width": 120,
"default_row_index_width": 70,
"default_row_index": "numbers",
"default_header": "letters",
"page_up_down_select_row": True,
"paste_can_expand_x": False,
"paste_can_expand_y": False,
"paste_insert_column_limit": None,
"paste_insert_row_limit": None,
"arrow_key_down_right_scroll_page": False,
"cell_auto_resize_enabled": True,
"auto_resize_row_index": True,
"max_undos": 30,
"column_drag_and_drop_perform": True,
"row_drag_and_drop_perform": True,
"empty_horizontal": 50,
"empty_vertical": 50,
"selected_rows_to_end_of_window": False,
"horizontal_grid_to_end_of_window": False,
"vertical_grid_to_end_of_window": False,
"show_vertical_grid": True,
"show_horizontal_grid": True,
"display_selected_fg_over_highlights": False,
"show_selected_cells_border": True,
"edit_cell_tab": "right",
"edit_cell_return": "down",
"editor_del_key": "forward",
"treeview": False,
"treeview_indent": "2",
"rounded_boxes": True,
"alternate_color": "",
"allow_cell_overflow": False,
"table_wrap": "c",
"header_wrap": "c",
"index_wrap": "c",
"min_column_width": 1,
"max_column_width": float("inf"),
"max_header_height": float("inf"),
"max_row_height": float("inf"),
"max_index_width": float("inf"),
"show_top_left": None,
"sort_key": natural_sort_key,
```

Notes:

- The parameters ending in `_image` take `tk.PhotoImage` or an empty string `""` as types.
- The parameters ending in `_compound` take one of the following: `"left", "right", "bottom", "top", "none", None`.
- `sort_key` is `Callable` - a function.
- A dictionary can be provided to `set_options()` instead of using the keyword arguments, e.g.:

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

---
# **Header and Index**

#### **Set the header**

```python
set_header_data(value: Any, c: int | None | Iterator = None, redraw: bool = True) -> Sheet
```
- `value` (`iterable`, `int`, `Any`) if `c` is left as `None` then it attempts to set the whole header as the `value` (converting a generator to a list). If `value` is `int` it sets the header to display the row with that position.
- `c` (`iterable`, `int`, `None`) if both `value` and `c` are iterables it assumes `c` is an iterable of positions and `value` is an iterable of values and attempts to set each value to each position. If `c` is `int` it attempts to set the value at that position.

```python
headers(
    newheaders: Any = None,
    index: None | int = None,
    reset_col_positions: bool = False,
    show_headers_if_not_sheet: bool = True,
    redraw: bool = True,
) -> Any
```
- Using an integer `int` for argument `newheaders` makes the sheet use that row as a header e.g. `headers(0)` means the first row will be used as a header (the first row will not be hidden in the sheet though), this is sort of equivalent to freezing the row.
- Leaving `newheaders` as `None` and using the `index` argument returns the existing header value in that index.
- Leaving all arguments as default e.g. `headers()` returns existing headers.

___

#### **Set the index**

```python
set_index_data(value: Any, r: int | None | Iterator = None, redraw: bool = True) -> Sheet
```
- `value` (`iterable`, `int`, `Any`) if `r` is left as `None` then it attempts to set the whole index as the `value` (converting a generator to a list). If `value` is `int` it sets the index to display the row with that position.
- `r` (`iterable`, `int`, `None`) if both `value` and `r` are iterables it assumes `r` is an iterable of positions and `value` is an iterable of values and attempts to set each value to each position. If `r` is `int` it attempts to set the value at that position.

```python
row_index(
    newindex: Any = None,
    index: None | int = None,
    reset_row_positions: bool = False,
    show_index_if_not_sheet: bool = True,
    redraw: bool = True,
) -> Any
```
- Using an integer `int` for argument `newindex` makes the sheet use that column as an index e.g. `row_index(0)` means the first column will be used as an index (the first column will not be hidden in the sheet though), this is sort of equivalent to freezing the column.
- Leaving `newindex` as `None` and using the `index` argument returns the existing row index value in that index.
- Leaving all arguments as default e.g. `row_index()` returns the existing row index.

---
# **Text Wrap and Overflow**

#### **Control text wrapping**

You can set table, header and index text wrapping either at `Sheet()` initialization or using `set_options()`.

Make use of the following parameters:

- `table_wrap`
- `index_wrap`
- `header_wrap`

With one of the following arguments:

- `""` - For no text wrapping.
- `"c"` - For character wrapping.
- `"w"` - For word wrapping.

**Examples:**
```python
# for word wrap at initialization
my_sheet = Sheet(parent, table_wrap="w")

# for character wrap using set_options()
my_sheet.set_options(table_wrap="c")
```

#### **Control table text overflow**

This setting only works for cells that are not center (north) aligned. Cell text can be set to overflow adjacent empty cells in the table like so:

**Examples:**
```python
# for word wrap at initialization
my_sheet = Sheet(parent, allow_cell_overflow=True)

# for character wrap using set_options()
my_sheet.set_options(allow_cell_overflow=True)
```
- Set it to `False` to disable it.
- It is only available as a global setting for the table, not on a cell by cell basis.

---
# **Table functionality and bindings**

#### **Enable table functionality and bindings**

```python
enable_bindings(*bindings: Binding, menu: bool = True)
```

Parameters:

- `bindings` (`str`) options are (rc stands for right click):
	- `"all"` # enables all bindings with `single_select` mode, except the bindings that have to be specifically enabled by name.
	- `"single_select"` # normal selection mode
	- `"toggle_select"` # has issues but to enable a selection mode where cell/row/column selection is toggled
	- `"drag_select"` # to allow mouse click and drag selection of cells/rows/columns
       - `"select_all"` # drag_select also enables select_all
	- `"column_drag_and_drop"` / `"move_columns"` # to allow drag and drop of columns
	- `"row_drag_and_drop"` / `"move_rows"` # to allow drag and drop of rows
	- `"column_select"` # to allow column selection
	- `"row_select"` # to allow row selection
	- `"column_width_resize"` # for resizing columns
	- `"double_click_column_resize"` # for resizing columns to row text width
	- `"row_width_resize"` # to resize the index width
	- `"column_height_resize"` # to resize the header height
	- `"arrowkeys"` # all arrowkeys including page up and down
    - `"up"` # individual arrow key
    - `"down"` # individual arrow key
    - `"left"` # individual arrow key
    - `"right"` # individual arrow key
    - `"prior"` # page up
    - `"next"` # page down
	- `"row_height_resize"` # to resize rows
	- `"double_click_row_resize"` # for resizing rows to row text height
	- `"right_click_popup_menu"` / `"rc_popup_menu"` / `"rc_menu"` # for the in-built table context menu
	- `"rc_select"` # for selecting cells using right click
	- `"rc_insert_column"` # for a menu option to add columns
	- `"rc_delete_column"` # for a menu option to delete columns
	- `"rc_insert_row"` # for a menu option to add rows
	- `"rc_delete_row"` # for a menu option to delete rows
    - `"sort_cells"`
    - `"sort_row"`
    - `"sort_column"` / `"sort_col"`
    - `"sort_rows"`
    - `"sort_columns"` / `"sort_cols"`
	- `"copy"` # for copying to clipboard
	- `"cut"` # for cutting to clipboard
	- `"paste"` # for pasting into the table
	- `"delete"` # for clearing cells with the delete key
	- `"undo"` # for undo and redo
    - `"edit_cell"` # allow table cell editing
    - `"find"` # for a pop-up find window (does not find in index or header)
    - `"replace"` # additional functionality for the find window, replace and replace all
    - *`"ctrl_click_select"` / `"ctrl_select"` # for selecting multiple non-adjacent cells/rows/columns
    - *`"edit_header"` # allow header cell editing
    - *`"edit_index"` # allow index cell editing
    - ***has to be specifically enabled - See Notes.**
- `menu`(`bool`) when `True` adds the related functionality to the in-built popup menu. Only applicable for edit bindings such as Cut, Copy, Paste and Delete which have both a keyboard binding and a menu entry.

Notes:

- You can change the Sheets key bindings for functionality such as copy, paste, up, down etc. Instructions can be found [here](#changing-key-bindings).
- **Note** that the following functionalities are not enabled using `"all"` and have to be specifically enabled:
    - `"ctrl_click_select"` / `"ctrl_select"`
    - `"edit_header"`
    - `"edit_index"`
- To allow table expansion when pasting data which doesn't fit in the table use either:
   - `paste_can_expand_x=True`, `paste_can_expand_y=True` in sheet initialization arguments or the same keyword arguments with the function `set_options()`.

Example:

- `sheet.enable_bindings()` to enable everything except `"ctrl_select"`, `"edit_index"`, `"edit_header"`.

___

#### **Disable table functionality and bindings**

```python
disable_bindings(*bindings: Binding)
```
Notes:

- Uses the same arguments as `enable_bindings()`.

___

#### **Bind specific table functionality**

This function allows you to bind **very** specific table functionality to your own functions:

- If you want less specificity in event names you can also bind all sheet modifying events to a single function, [see here](#tkinter-and-tksheet-events).
- If you want to validate/modify user cell edits [see here](#validate-user-cell-edits).

```python
extra_bindings(
    bindings: str | list | tuple,
    func: Callable | None = None,
) -> Sheet
```

**There are several ways to use this function:**

- `bindings` as a `str` and `func` as either `None` or a function. Using `None` as an argument for `func` will effectively unbind the function.
    - `extra_bindings("edit_cell", func=my_function)`
- `bindings` as an `iterable` of `str`s and `func` as either `None` or a function. Using `None` as an argument for `func` will effectively unbind the function.
    - `extra_bindings(["all_select_events", "copy", "cut"], func=my_function)`
- `bindings` as an `iterable` of `list`s or `tuple`s with length of two, e.g.
    - `extra_bindings([(binding, function), (binding, function), ...])` In this example you could also use `None` in the place of `function` to unbind the binding.
    - In this case the arg `func` is totally ignored.
- For `"end_..."` events the bound function is run before the value is set.
- **To unbind** a function either set `func` argument to `None` or leave it as default e.g. `extra_bindings("begin_copy")` to unbind `"begin_copy"`.
- Even though undo/redo edits or adds or deletes rows/columns the bound functions for those actions will not be called. Undo/redo must be specifically bound in order for a function to be called.

**`bindings` (`str`) options:**

Undo/Redo:

- `"begin_undo", "begin_ctrl_z"`
- `"ctrl_z", "end_undo", "end_ctrl_z", "undo"`

Editing:

- `"begin_copy", "begin_ctrl_c"`
- `"ctrl_c", "end_copy", "end_ctrl_c", "copy"`
- `"begin_cut", "begin_ctrl_x"`
- `"ctrl_x", "end_cut", "end_ctrl_x", "cut"`
- `"begin_paste", "begin_ctrl_v"`
- `"ctrl_v", "end_paste", "end_ctrl_v", "paste"`
- `"begin_delete_key", "begin_delete"`
- `"delete_key", "end_delete", "end_delete_key", "delete"`
- `"begin_edit_cell", "begin_edit_table"`
- `"end_edit_cell", "edit_cell", "edit_table"`
- `"begin_edit_header"`
- `"end_edit_header", "edit_header"`
- `"begin_edit_index"`
- `"end_edit_index", "edit_index"`
- `"replace_all"`

Moving:

- `"begin_row_index_drag_drop", "begin_move_rows"`
- `"row_index_drag_drop", "move_rows", "end_move_rows", "end_row_index_drag_drop"`
- `"begin_column_header_drag_drop", "begin_move_columns"`
- `"column_header_drag_drop", "move_columns", "end_move_columns", "end_column_header_drag_drop"`
- `"begin_sort_cells"`
- `"sort_cells", "end_sort_cells"`
- `"begin_sort_rows"`
- `"sort_rows", "end_sort_rows"`
- `"begin_sort_columns"`
- `"sort_columns", "end_sort_columns"`

Deleting:

- `"begin_rc_delete_row", "begin_delete_rows"`
- `"rc_delete_row", "end_rc_delete_row", "end_delete_rows", "delete_rows"`
- `"begin_rc_delete_column", "begin_delete_columns"`
- `"rc_delete_column", "end_rc_delete_column", "end_delete_columns", "delete_columns"`

Adding:

- `"begin_rc_insert_column", "begin_insert_column", "begin_insert_columns", "begin_add_column", "begin_rc_add_column", "begin_add_columns"`
- `"rc_insert_column", "end_rc_insert_column", "end_insert_column", "end_insert_columns", "rc_add_column", "end_rc_add_column", "end_add_column", "end_add_columns", "add_columns"`
- `"begin_rc_insert_row", "begin_insert_row", "begin_insert_rows", "begin_rc_add_row", "begin_add_row", "begin_add_rows"`
- `"rc_insert_row", "end_rc_insert_row", "end_insert_row", "end_insert_rows", "rc_add_row", "end_rc_add_row", "end_add_row", "end_add_rows", "add_rows"`

Resizing rows/columns:

- `"row_height_resize"`
- `"column_width_resize"`

Selection:

- `"cell_select"`
- `"all_select"`
- `"row_select"`
- `"column_select"`
- `"drag_select_cells"`
- `"drag_select_rows"`
- `"drag_select_columns"`
- `"shift_cell_select"`
- `"shift_row_select"`
- `"shift_column_select"`
- `"ctrl_cell_select"`
- `"ctrl_row_select"`
- `"ctrl_column_select"`
- `"deselect"`

Event collections:

- `"all_select_events", "select", "selectevents", "select_events"`
- `"all_modified_events", "sheetmodified", "sheet_modified" "modified_events", "modified"`
- `"bind_all"`
- `"unbind_all"`

**Further Notes:**

- `func` argument is the function you want to send the binding event to.
- Using one of the following `"all_modified_events"`, `"sheetmodified"`, `"sheet_modified"`, `"modified_events"`, `"modified"` will make any insert, delete or cell edit including pastes and undos send an event to your function.
- For events `"begin_move_columns"`/`"begin_move_rows"` the point where columns/rows will be moved to will be accessible by the key named `"value"`.
- For `"begin_edit..."` events the bound function must return a value to open the cell editor with, example [here](#example-custom-right-click-and-text-editor-validation).

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
    "row": None,
    "column": None,
    "resized": {
        "rows": {},
        "columns": {},
    },
    "widget": None,
}
```

Keys:

- A function bound using `extra_bindings()` will receive event data with one of the following **`["eventname"]`** keys:
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
    - `"select"`
    - `"resize"`
    - *`"begin_move_rows"`
    - *`"end_move_rows"`
    - *`"begin_move_columns"`
    - *`"end_move_columns"`

- *is also used as the event name for sorting rows/columns events.

- `EventDataDict`s will otherwise have one of the following event names:
    - `"edit_table"` when a user has cut, paste, delete or made any cell edits including using dropdown boxes etc. in the table.
    - `"edit_index"` when a user has edited a index cell.
    - `"edit_header"` when a user has edited a header cell.
    - `"add_columns"` when a user has inserted columns.
    - `"add_rows"` when a user has inserted rows.
    - `"delete_columns"` when a user has deleted columns.
    - `"delete_rows"` when a user has deleted rows.
    - `"move_columns"` when a user has dragged and dropped OR sorted columns.
    - `"move_rows"` when a user has dragged and dropped OR sorted rows.
    - `"select"`
    - `"resize"`
- These event names would be used for `"<<SheetModified>>"` bound events for example.

- For events `"begin_move_columns"`/`"begin_move_rows"` the point where columns/rows will be moved to will be under the `event_data` key `"value"`.
- Key **`["sheetname"]`** is the [name given to the sheet widget on initialization](#initialization-options), useful if you have multiple sheets to determine which one emitted the event.
- Key **`["cells"]["table"]`** if any table cells have been modified by cut, paste, delete, cell editors, dropdown boxes, check boxes, undo or redo this will be a `dict` with `tuple` keys of `(data row index: int, data column index: int)` and the values will be the cell values at that location **prior** to the change. The `dict` will be empty if no such changes have taken place.
- Key **`["cells"]["header"]`** if any header cells have been modified by cell editors, dropdown boxes, check boxes, undo or redo this will be a `dict` with keys of `int: data column index` and the values will be the cell values at that location **prior** to the change. The `dict` will be empty if no such changes have taken place.
- Key **`["cells"]["index"]`** if any index cells have been modified by cell editors, dropdown boxes, check boxes, undo or redo this will be a `dict` with keys of `int: data row index` and the values will be the cell values at that location **prior** to the change. The `dict` will be empty if no such changes have taken place.
- Key **`["moved"]["rows"]`** if any rows have been moved by dragging and dropping or undoing/redoing of dragging and dropping rows this will be a `dict` with the following keys:
    - `{"data": {old data index: new data index, ...}, "displayed": {old displayed index: new displayed index, ...}}`
        - `"data"` will be a `dict` where the keys are the old data indexes of the rows and the values are the data indexes they have moved to.
        - `"displayed"` will be a `dict` where the keys are the old displayed indexes of the rows and the values are the displayed indexes they have moved to.
        - If no rows have been moved the `dict` under `["moved"]["rows"]` will be empty.
        - Note that if there are hidden rows the values for `"data"` will include all currently displayed row indexes and their new locations. If required and available, the values under `"displayed"` include only the directly moved rows, convert to data indexes using `Sheet.data_r()`.
    - For events `"begin_move_rows"` the point where rows will be moved to will be under the `event_data` key `"value"`.
- Key **`["moved"]["columns"]`** if any columns have been moved by dragging and dropping or undoing/redoing of dragging and dropping columns this will be a `dict` with the following keys:
    - `{"data": {old data index: new data index, ...}, "displayed": {old displayed index: new displayed index, ...}}`
        - `"data"` will be a `dict` where the keys are the old data indexes of the columns and the values are the data indexes they have moved to.
        - `"displayed"` will be a `dict` where the keys are the old displayed indexes of the columns and the values are the displayed indexes they have moved to.
        - If no columns have been moved the `dict` under `["moved"]["columns"]` will be empty.
        - Note that if there are hidden columns the values for `"data"` will include all currently displayed column indexes and their new locations. If required and available, the values under `"displayed"` include only the directly moved columns, convert to data indexes using `Sheet.data_c()`.
    - For events `"begin_move_columns"` the point where columns will be moved to will be under the `event_data` key `"value"`.
- Key **`["added"]["rows"]`** if any rows have been added by the inbuilt popup menu insert rows or by a paste which expands the sheet then this will be a `dict` with the following keys:
    - `{"data_index": int, "displayed_index": int, "num": int, "displayed": []}`
        - `"data_index"` is an `int` representing the row where the rows were added in the data.
        - `"displayed_index"` is an `int` representing the displayed table index where the rows were added (which will be different from the data index if there are hidden rows).
        - `"displayed"` is a copied list of the `Sheet()`s displayed rows immediately prior to the change.
        - If no rows have been added the `dict` will be empty.
- Key **`["added"]["columns"]`** if any columns have been added by the inbuilt popup menu insert columns or by a paste which expands the sheet then this will be a `dict` with the following keys:
    - `{"data_index": int, "displayed_index": int, "num": int, "displayed": []}`
        - `"data_index"` is an `int` representing the column where the columns were added in the data.
        - `"displayed_index"` is an `int` representing the displayed table index where the columns were added (which will be different from the data index if there are hidden columns).
        - `"displayed"` is a copied list of the `Sheet()`s displayed columns immediately prior to the change.
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
- Key **`["options"]`** This serves as storage for the `Sheet()`s options such as highlights, formatting, alignments, dropdown boxes, check boxes etc. It is a `dict` where the values are the sheets internal cell/row/column options `dicts`.
- Key **`["selection_boxes"]`** the value of this is all selection boxes on the sheet in the form of a `dict` as shown below:
    - For every event except `"select"` events the selection boxes are those immediately prior to the modification, for `"select"` events they are the current selection boxes.
    - The layout is always: `"selection_boxes": {(start row, start column, up to but not including row, up to but not including column): selection box type}`.
        - The row/column indexes are `int`s and the selection box type is a `str` either `"cells"`, `"rows"` or `"columns"`.
    - The `dict` will be empty if there is nothing selected.
- Key **`["selected"]`** the value of this when there is something selected on the sheet is a `namedtuple`. The values of which can be found [here](#get-the-currently-selected-cell).
    - When nothing is selected or the event is not relevant to the currently selected box, such as a resize event it will be an empty `tuple`.
- Key **`["being_selected"]`** if any selection box is in the process of being drawn by holding down mouse button 1 and dragging then this will be a tuple with the following layout:
    - `(start row, start column, up to but not including row, up to but not including column, selection box type)`.
        - The selection box type is a `str` either `"cells"`, `"rows"` or `"columns"`.
    - If no box is in the process of being created then this will be a an empty `tuple`.
    - [See here](#example-displaying-selections) for an example.
- Key **`["data"]`** - `dict[tuple[int, int], Any]` - changed from only being used by paste, now stores a `dict` of cell coordinates and values that make up a table edit event of more than one cell.
- Key **`["key"]`** - `str` - is primarily used for cell edit events where a key press has occurred. For `"begin_edit..."` events the value is the actual key which was pressed (or `"??"` for using the mouse to open a cell). It also might be one of the following for end edit events:
    - `"Return"` - enter key.
    - `"FocusOut"` - the editor or box lost focus, perhaps by mouse clicking elsewhere.
    - `"Tab"` - tab key.
- Key **`["value"]`** is used primarily by cell editing events. For `"begin_edit..."` events it's the value displayed in he text editor when it opens. For `"end_edit..."` events it's the value in the text editor when it was closed, for example by hitting `Return`. It also used by `"begin_move_columns"`/`"begin_move_rows"` - the point where columns/rows will be moved to will be under the `event_data` key `"value"`.
- Key **`["loc"]`** is for cell editing events to show the displayed (not data) coordinates of the event. It will be **either:**
    - A tuple of `(int displayed row index, int displayed column index)` in the case of editing table cells.
    - A single `int` in the case of editing index/header cells.
- Key **`["row"]`** is for cell editing events to show the displayed (not data) row number `int` of the event. If the event was not a cell editing event or a header cell was edited the value will be `None`.
- Key **`["column"]`** is for cell editing events to show the displayed (not data) column number `int` of the event. If the event was not a cell editing event or an index cell was edited the value will be `None`.
- Key **`["resized"]["rows"]`** is for row height resizing events, it will be a `dict` with the following layout:
    - `{int displayed row index: {"old_size": old_height, "new_size": new_height}}`.
    - If no rows have been resized then the value for `["resized"]["rows"]` will be an empty `dict`.
- Key **`["resized"]["columns"]`** is for column width resizing events, it will be a `dict` with the following layout:
    - `{int displayed column index: {"old_size": old_width, "new_size": new_width}}`.
    - If no columns have been resized then the value for `["resized"]["columns"]` will be an empty `dict`.
- Key **`["widget"]`** will contain the widget which emitted the event, either the `MainTable()`, `ColumnHeaders()` or `RowIndex()` which are all `tk.Canvas` widgets.

___

#### **Validate user cell edits**

With these functions you can validate or modify most user sheet edits, includes cut, paste, delete (including column/row clear), dropdown boxes and cell edits.

**Edit validation**

This function will be called for every cell edit in an action.

```python
edit_validation(func: Callable | None = None) -> Sheet
```
Parameters:

- `func` (`Callable`, `None`) must either be a function which will receive a tksheet event dict which looks like [this](#event-data) or `None` which unbinds the function.

Notes:

- If your bound function returns `None` then that specific cell edit will not be performed.
- For examples of this function see [here](#usage-examples) and [here](#example-custom-right-click-and-text-editor-validation).

**Bulk edit validation**

This function will be called at the end of an action and delay any edits until after validation.

```python
bulk_table_edit_validation(func: Callable | None = None) -> Sheet
```
Parameters:

- `func` (`Callable`, `None`) must either be a function which will receive a tksheet event dict which looks like [this](#event-data) or `None` which unbinds the function.

Notes:

- See the below example for more information on usage.

Example:

```python
from tksheet import Sheet
import tkinter as tk

from typing import Any


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
            data=[[f"Row {r}, Column {c}" for c in range(3)] for r in range(3)],
        )
        self.sheet.enable_bindings()
        self.sheet.bulk_table_edit_validation(self.validate)
        self.frame.grid(row=0, column=0, sticky="nswe")
        self.sheet.grid(row=0, column=0, sticky="nswe")

    def validate(self, event: dict) -> Any:
        """
        Whatever keys and values are left in event["data"]
        when the function returns are the edits that will be made

        An example below shows preventing edits if the proposed edit
        contains a space

        But you can also modify the values, or add more key, value
        pairs to event["data"]
        """
        not_valid = set()
        for (r, c), value in event.data.items():
            if " " in value:
                not_valid.add((r, c))
        event.data = {k: v for k, v in event.data.items() if k not in not_valid}


app = demo()
app.mainloop()
```

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
    image: tk.PhotoImage | Literal[""] = "",
    compound: Literal["top", "bottom", "left", "right", "none"] | None = None,
    accelerator: str | None = None,
) -> Sheet
```

Notes:

- Either creates or overwrites an existing menu command.

Example:

```python
self.sheet.popup_menu_add_command(
    label="Test insert rows",
    func=self.my_fn,
    image=tk.PhotoImage(file="filepath_to_img.png"),
    compound="left",
)
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
cut(event: Any = None) -> Sheet
copy(event: Any = None) -> Sheet
paste(event: Any = None) -> Sheet
delete(event: Any = None) -> Sheet
undo(event: Any = None) -> Sheet
redo(event: Any = None) -> Sheet
zoom_in() -> Sheet
zoom_out() -> Sheet
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

### **Sheet bind**

- With the `Sheet.bind()` function you can bind things in the usual way you would in tkinter and they will bind to all the `tksheet` canvases.
- There are also the following special `tksheet` events you can bind:

| Binding                | Usable with `event_generate()` |
| --------               | -------                        |
| `"<<SheetModified>>"`  | -                              |
| `"<<SheetRedrawn>>"`   | -                              |
| `"<<SheetSelect>>"`    | -                              |
| `"<<Copy>>"`           | X                              |
| `"<<Cut>>"`            | X                              |
| `"<<Paste>>"`          | X                              |
| `"<<Delete>>"`         | X                              |
| `"<<Undo>>"`           | X                              |
| `"<<Redo>>"`           | X                              |
| `"<<SelectAll>>"`      | X                              |

```python
bind(
    event: str,
    func: Callable,
    add: str | None = None,
)
```
Parameters:

- `add` may or may not work for various bindings depending on whether they are already in use by `tksheet`.
- **Note** that while a bound event after a paste/undo/redo might have the event name `"edit_table"` it also might have added/deleted rows/columns, refer to the docs on the event data `dict` for more information.
- `event` the emitted events are:
    - `"<<SheetModified>>"` emitted whenever the sheet was modified by the end user by editing cells or adding or deleting rows/columns. The function you bind to this event must be able to receive a `dict` argument which will be the same as [the event data dict](#event-data) but with less specific event names. The possible event names are listed below:
        - `"edit_table"` when a user has cut, paste, delete or made any cell edits including using dropdown boxes etc. in the table.
        - `"edit_index"` when a user has edited a index cell.
        - `"edit_header"` when a user has edited a header cell.
        - `"add_columns"` when a user has inserted columns.
        - `"add_rows"` when a user has inserted rows.
        - `"delete_columns"` when a user has deleted columns.
        - `"delete_rows"` when a user has deleted rows.
        - `"move_columns"` when a user has dragged and dropped columns.
        - `"move_rows"` when a user has dragged and dropped rows.
        - `"sort_rows"` when rows have been re-ordered by sorting.
        - `"sort_columns"` when columns have been re-ordered by sorting.
    - `"<<SheetRedrawn>>"` emitted whenever the sheet GUI was refreshed (redrawn). The data for this event will be different than the usual event data, it is:
        - `{"sheetname": name of your sheet, "header": bool True if the header was redrawn, "row_index": bool True if the index was redrawn, "table": bool True if the the table was redrawn}`
    - `"<<SheetSelect>>"` encompasses all select events and emits the same event as `"<<SheetModified>>"` but with the event name: `"select"`.
    - `"<<Copy>>"` emitted when a Sheet copy e.g. `<Control-c>` was performed and will have the `eventname` `"copy"`.
    - `"<<Cut>>"`
    - `"<<Paste>>"`
    - `"<<Delete>>"` emitted when a Sheet delete key function was performed.
    - `"<<SelectAll>>"`
    - `"<<Undo>>"`
    - `"<<Redo>>"`

Example:

```python
# self.sheet_was_modified is your function
self.sheet.bind("<<SheetModified>>", self.sheet_was_modified)
```

Example for `event_generate()`:
```python
self.sheet.event_generate("<<Copy>>")
```
- Tells the sheet to run its copy function.

___

### **Sheet unbind**

With this function you can unbind things you have bound using the `bind()` function.
```python
unbind(binding: str) -> Sheet
```

---
# **Sheet Languages and Bindings**

In this section are instructions to change some of tksheets in-built language and bindings:

- The in-built right click menu.
- The in-built functionality keybindings, such as copy, paste etc.

Please note that due to the limitations of the Tkinter Canvas tksheet doesn’t support right-to-left (RTL) languages.

#### **Changing right click menu labels**

You can change the labels for tksheets in-built right click popup menu by using the [`set_options()` function](#sheet-options) with any of the following keyword arguments:

```python
# edit header
"edit_header_label": "Edit header",
"edit_header_accelerator": "",
"edit_header_image": tk.PhotoImage(data=ICON_EDIT),
"edit_header_compound": "left",
# edit index
"edit_index_label": "Edit index",
"edit_index_accelerator": "",
"edit_index_image": tk.PhotoImage(data=ICON_EDIT),
"edit_index_compound": "left",
# edit cell
"edit_cell_label": "Edit cell",
"edit_cell_accelerator": "",
"edit_cell_image": tk.PhotoImage(data=ICON_EDIT),
"edit_cell_compound": "left",
# cut
"cut_label": "Cut",
"cut_accelerator": "Ctrl+X",
"cut_image": tk.PhotoImage(data=ICON_CUT),
"cut_compound": "left",
# cut contents
"cut_contents_label": "Cut contents",
"cut_contents_accelerator": "Ctrl+X",
"cut_contents_image": tk.PhotoImage(data=ICON_CUT),
"cut_contents_compound": "left",
# copy
"copy_label": "Copy",
"copy_accelerator": "Ctrl+C",
"copy_image": tk.PhotoImage(data=ICON_COPY),
"copy_compound": "left",
# copy contents
"copy_contents_label": "Copy contents",
"copy_contents_accelerator": "Ctrl+C",
"copy_contents_image": tk.PhotoImage(data=ICON_COPY),
"copy_contents_compound": "left",
# paste
"paste_label": "Paste",
"paste_accelerator": "Ctrl+V",
"paste_image": tk.PhotoImage(data=ICON_PASTE),
"paste_compound": "left",
# delete
"delete_label": "Delete",
"delete_accelerator": "Del",
"delete_image": tk.PhotoImage(data=ICON_CLEAR),
"delete_compound": "left",
# clear contents
"clear_contents_label": "Clear contents",
"clear_contents_accelerator": "Del",
"clear_contents_image": tk.PhotoImage(data=ICON_CLEAR),
"clear_contents_compound": "left",
# del columns
"delete_columns_label": "Delete columns",
"delete_columns_accelerator": "",
"delete_columns_image": tk.PhotoImage(data=ICON_DEL),
"delete_columns_compound": "left",
# insert columns left
"insert_columns_left_label": "Insert columns left",
"insert_columns_left_accelerator": "",
"insert_columns_left_image": tk.PhotoImage(data=ICON_ADD),
"insert_columns_left_compound": "left",
# insert columns right
"insert_columns_right_label": "Insert columns right",
"insert_columns_right_accelerator": "",
"insert_columns_right_image": tk.PhotoImage(data=ICON_ADD),
"insert_columns_right_compound": "left",
# insert single column
"insert_column_label": "Insert column",
"insert_column_accelerator": "",
"insert_column_image": tk.PhotoImage(data=ICON_ADD),
"insert_column_compound": "left",
# del rows
"delete_rows_label": "Delete rows",
"delete_rows_accelerator": "",
"delete_rows_image": tk.PhotoImage(data=ICON_DEL),
"delete_rows_compound": "left",
# insert rows above
"insert_rows_above_label": "Insert rows above",
"insert_rows_above_accelerator": "",
"insert_rows_above_image": tk.PhotoImage(data=ICON_ADD),
"insert_rows_above_compound": "left",
# insert rows below
"insert_rows_below_label": "Insert rows below",
"insert_rows_below_accelerator": "",
"insert_rows_below_image": tk.PhotoImage(data=ICON_ADD),
"insert_rows_below_compound": "left",
# insert single row
"insert_row_label": "Insert row",
"insert_row_accelerator": "",
"insert_row_image": tk.PhotoImage(data=ICON_ADD),
"insert_row_compound": "left",
# sorting
# labels
"sort_cells_label": "Sort Asc.",
"sort_cells_x_label": "Sort row-wise Asc.",
"sort_row_label": "Sort values Asc.",
"sort_column_label": "Sort values Asc.",
"sort_rows_label": "Sort rows Asc.",
"sort_columns_label": "Sort columns Asc.",
# reverse labels
"sort_cells_reverse_label": "Sort Desc.",
"sort_cells_x_reverse_label": "Sort row-wise Desc.",
"sort_row_reverse_label": "Sort values Desc.",
"sort_column_reverse_label": "Sort values Desc.",
"sort_rows_reverse_label": "Sort rows Desc.",
"sort_columns_reverse_label": "Sort columns Desc.",
# accelerators
"sort_cells_accelerator": "",
"sort_cells_x_accelerator": "",
"sort_row_accelerator": "",
"sort_column_accelerator": "",
"sort_rows_accelerator": "",
"sort_columns_accelerator": "",
# reverse accelerators
"sort_cells_reverse_accelerator": "",
"sort_cells_x_reverse_accelerator": "",
"sort_row_reverse_accelerator": "",
"sort_column_reverse_accelerator": "",
"sort_rows_reverse_accelerator": "",
"sort_columns_reverse_accelerator": "",
# images
"sort_cells_image": tk.PhotoImage(data=ICON_SORT_ASC),
"sort_cells_x_image": tk.PhotoImage(data=ICON_SORT_ASC),
"sort_row_image": tk.PhotoImage(data=ICON_SORT_ASC),
"sort_column_image": tk.PhotoImage(data=ICON_SORT_ASC),
"sort_rows_image": tk.PhotoImage(data=ICON_SORT_ASC),
"sort_columns_image": tk.PhotoImage(data=ICON_SORT_ASC),
# compounds
"sort_cells_compound": "left",
"sort_cells_x_compound": "left",
"sort_row_compound": "left",
"sort_column_compound": "left",
"sort_rows_compound": "left",
"sort_columns_compound": "left",
# reverse images
"sort_cells_reverse_image": tk.PhotoImage(data=ICON_SORT_DESC),
"sort_cells_x_reverse_image": tk.PhotoImage(data=ICON_SORT_DESC),
"sort_row_reverse_image": tk.PhotoImage(data=ICON_SORT_DESC),
"sort_column_reverse_image": tk.PhotoImage(data=ICON_SORT_DESC),
"sort_rows_reverse_image": tk.PhotoImage(data=ICON_SORT_DESC),
"sort_columns_reverse_image": tk.PhotoImage(data=ICON_SORT_DESC),
# reverse compounds
"sort_cells_reverse_compound": "left",
"sort_cells_x_reverse_compound": "left",
"sort_row_reverse_compound": "left",
"sort_column_reverse_compound": "left",
"sort_rows_reverse_compound": "left",
"sort_columns_reverse_compound": "left",
# select all
"select_all_label": "Select all",
"select_all_accelerator": "Ctrl+A",
"select_all_image": tk.PhotoImage(data=ICON_SELECT_ALL),
"select_all_compound": "left",
# undo
"undo_label": "Undo",
"undo_accelerator": "Ctrl+Z",
"undo_image": tk.PhotoImage(data=ICON_UNDO),
"undo_compound": "left",
# redo
"redo_label": "Redo",
"redo_accelerator": "Ctrl+Shift+Z",
"redo_image": tk.PhotoImage(data=ICON_REDO),
"redo_compound": "left",
```

Example:

```python
# changing the copy label to the spanish for Copy
sheet.set_options(copy_label="Copiar", copy_image=tk.PhotoImage(file="filepath_to_img.png"), copy_compound="left")
```

#### **Changing key bindings**

You can change the bindings for tksheets in-built functionality such as cut, copy, paste by using the [`set_options()` function](#sheet-options) with any the following keyword arguments:

```python
copy_bindings
cut_bindings
paste_bindings
undo_bindings
redo_bindings
delete_bindings
select_all_bindings
select_columns_bindings
select_rows_bindings
row_start_bindings
table_start_bindings
tab_bindings
up_bindings
right_bindings
down_bindings
left_bindings
shift_up_bindings
shift_right_bindings
shift_down_bindings
shift_left_bindings
prior_bindings
next_bindings
find_bindings
find_next_bindings
find_previous_bindings
escape_bindings
toggle_replace_bindings
```

The argument must be a `list` of **tkinter** binding `str`s. In the below example the binding for copy is changed to `"<Control-e>"` and `"<Control-E>"`.

```python
# changing the binding for copy
sheet.set_options(copy_bindings=["<Control-e>", "<Control-E>"])
```

The default values for these bindings can be found in the tksheet file `sheet_options.py`.

#### **Key bindings for other languages**

There is limited support in tkinter for keybindings in languages other than english, for example tkinters `.bind()` function doesn't cooperate with cyrillic characters.

There are ways around this however, see below for a limited example of how this might be achieved:

```python
from __future__ import annotations

import tkinter as tk

from tksheet import Sheet


class demo(tk.Tk):
    def __init__(self) -> None:
        tk.Tk.__init__(self)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.sheet = Sheet(
            parent=self,
            data=[[f"{r} {c}" for c in range(5)] for r in range(5)],
        )
        self.sheet.enable_bindings()
        self.sheet.grid(row=0, column=0, sticky="nswe")
        self.bind_all("<Key>", self.any_key)

    def any_key(self, event: tk.Event) -> None:
        """
        Establish that the Control key is held down
        """
        ctrl = (event.state & 4 > 0)
        if not ctrl:
            return
        """
        From here you can use event.keycode and event.keysym to determine
        which key has been pressed along with Control
        """
        print(event.keycode)
        print(event.keysym)
        """
        If the keys are the ones you want to have bound to Sheet functionality
        You can then call the Sheets functionality using event_generate()
        For example:
        """
        # if the key is correct then:
        self.sheet.event_generate("<<Copy>>")


app = demo()
app.mainloop()
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

Whether cells, rows or columns are affected will depend on the spans [`kind`](#get-a-spans-kind).

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
    convert: Callable | None = None,
    undo: bool = True,
    emit_event: bool = False,
    widget: Any = None,
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
- `emit_event` when `True` and when using data setting functions that utilize spans causes a `"<<SheetModified>>` event to occur if it has been bound, see [here](#tkinter-and-tksheet-events) for more information on binding this event.
- `widget` (`Any`) is the reference to the original sheet which created the span. This can be changed to a different sheet if required e.g. `my_span.widget = new_sheet`.
- `expand` (`None`, `str`) must be either `None` or:
    - `"table"`/`"both"` expand the span both down and right from the span start to the ends of the table.
    - `"right"` expand the span right to the end of the table `x` axis.
    - `"down"` expand the span downwards to the bottom of the table `y` axis.
- `formatter_options` (`dict`, `None`) must be either `None` or `dict`. If providing a `dict` it must be the same structure as used in format functions, see [here](#data-formatting) for more information. Used to turn the span into a format type span which:
    - When using `get_data()` will format the returned data.
    - When using `set_data()` will format the data being set but **NOT** create a new formatting rule on the sheet.
- `**kwargs` you can provide additional keyword arguments to the function for example those used in `span.highlight()` or `span.dropdown()` which are used when applying a named span to a table.

Notes:

- To create a named span see [here](#named-spans).

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
- `span.coords`

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
    widget: Any = None,
    expand: str | None = None,
    formatter_options: dict | None = None,
    **kwargs,
) -> Span
```
Note that if `None` is used for any of the following parameters then that `Span`s attribute will be unchanged:

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
- `emit_event` (`bool`, `None`) is used by data modifying functions that utilize spans. When `True` causes a `"<<SheetModified>>` event to occur if it has been bound, see [here](#tkinter-and-tksheet-events) for more information.
- `widget` (`Any`) is the reference to the original sheet which created the span. This can be changed to a different sheet if required e.g. `my_span.widget = new_sheet`.
- `expand` (`str`, `None`) must be either `None` or:
    - `"table"`/`"both"` expand the span both down and right from the span start to the ends of the table.
    - `"right"` expand the span right to the end of the table `x` axis.
    - `"down"` expand the span downwards to the bottom of the table `y` axis.
- `formatter_options` (`dict`, `None`) must be either `None` or `dict`. If providing a `dict` it must be the same structure as used in format functions, see [here](#data-formatting) for more information. Used to turn the span into a format type span which:
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
- `emit_event` (`bool`) is used by data modifying functions that utilize spans. When `True` causes a `"<<SheetModified>>` event to occur if it has been bound, see [here](#tkinter-and-tksheet-events) for more information.
- `widget` (`Any`) is the reference to the original sheet which created the span. This can be changed to a different sheet if required e.g. `my_span.widget = new_sheet`.
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

Formats table data, see the help on [formatting](#data-formatting) for more information. Note that using this function also creates a format rule for the affected table cells.

```python
span.format(
    formatter_options: dict = {},
    formatter_class: Any = None,
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

These examples show the formatting of the entire sheet (not including header and index) as `int` and creates a format rule for all currently existing cells. [Named spans](#named-spans) are required to create a rule for all future existing cells as well, for example those created by the end user inserting rows or columns.

#### **Using a span to delete data format rules**

Delete any currently existing format rules for parts of the table that are covered by the span. Should not be used where there are data formatting rules created by named spans, see [Named spans](#named-spans) for more information.

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

Delete any currently existing highlights for parts of the sheet that are covered by the span. Should not be used where there are highlights created by named spans, see [Named spans](#named-spans) for more information.

```python
span.dehighlight() -> Span
```

Example:
```python
span1 = sheet[2:4].highlight(bg="red", fg="black")
span1.dehighlight()
```

#### **Using a span to create dropdown boxes**

Creates dropdown boxes for parts of the sheet that are covered by the span. For more information see [here](#dropdown-boxes).

```python
span.dropdown(
    values: list = [],
    set_value: Any = None,
    state: Literal["normal", "readonly", "disabled"] = "normal",
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

Delete dropdown boxes for parts of the sheet that are covered by the span. Should not be used where there are dropdown box rules created by named spans, see [Named spans](#named-spans) for more information.

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
    edit_data: bool = True,
    checked: bool | None = None,
    state: Literal["normal", "disabled"] = "normal",
    redraw: bool = True,
    check_function: Callable | None = None,
    text: str = "",
) -> Span
```
Parameters:

- `edit_data` when `True` edits the underlying cell data to either `checked` if `checked` is a `bool` or tries to convert the existing cell data to a `bool`.
- `checked` is the initial creation value to set the box to, if `None` then and `edit_data` is `True` then it will try to convert the underlying cell data to a `bool`.
- `state` can be `"normal"` or `"disabled"`. If `"disabled"` then color will be same as table grid lines, else it will be the cells text color.
- `check_function` can be used to trigger a function when the user clicks a checkbox.
- `text` displays text next to the checkbox in the cell, but will not be used as data, data will either be `True` or `False`.

Notes:

- To get the current checkbox value either:
    - Get the cell data, more information [here](#getting-sheet-data).
    - Use the parameter `check_function` with a function of your own creation to be called when the checkbox is set by the user.

Example:
```python
sheet["D"].checkbox(
    checked=True,
    text="Switch",
)
```

#### **Using a span to delete check boxes**

Delete check boxes for parts of the sheet that are covered by the span. Should not be used where there are check box rules created by named spans, see [Named spans](#named-spans) for more information.

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
- Using `span.readonly(False)` deletes any existing readonly rules for the span. Should not be used where there are readonly rules created by named spans, see [Named spans](#named-spans) for more information.

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

Delete text alignment rules for parts of the sheet that are covered by the span. Should not be used where there are alignment rules created by named spans, see [Named spans](#named-spans) for more information.

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
- `emit_event` when `True` causes a `"<<SheetModified>>` event to occur if it has been bound, see [here](#tkinter-and-tksheet-events) for more information.

Example:
```python
# clears column D
sheet["D"].clear()
```

#### **Using a span to tag cells**

Tag cells, rows or columns depending on the spans kind, more information on tags [here](#tags).

```python
tag(*tags) -> Span
```
Notes:

- If `span.kind` is `"cell"` then cells will be tagged, if it's a row span then rows will be and so for columns.

Example:
```python
# tags rows 2, 3, 4 with "hello world"
sheet[2:5].tag("hello world")
```

#### **Using a span to untag cells**

Remove **all** tags from cells, rows or columns depending on the spans kind, more information on tags [here](#tags).

```python
untag() -> Span
```
Notes:

- If `span.kind` is `"cell"` then cells will be untagged, if it's a row span then rows will be and so for columns.

Example:
```python
# tags rows 2, 3, 4 with "hello" and "bye"
sheet[2:5].tag("hello", "bye")

# removes both "hello" and "bye" tags from rows 2, 3, 4
sheet[2:5].untag()
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
- Relevant keyword arguments e.g. if the `type_` is `"highlight"` then arguments for `sheet.highlight()` found [here](#highlighting-cells).

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

A `Span` object (more information [here](#span-objects)) is returned when using square brackets on a `Sheet` like so:

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
- `widget` (`Any`) is the reference to the original sheet which created the span (this is the widget that data is retrieved from). This can be changed to a different sheet if required e.g. `my_span.widget = new_sheet`.

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
) -> Any
```

Examples:
```python
data = self.sheet.get_data("A1")
data = self.sheet.get_data(0, 0, 3, 3)
data = self.sheet.get_data(
    self.sheet.span(":D", transposed=True)
)
```

#### **Get a single cells data**

This is a higher performance method to get a single cells data which may be useful for example when performing a very large number of single cell data retrievals in a loop.

```python
get_cell_data(r: int, c: int, get_displayed: bool = False) -> Any
```
- `get_displayed` (`bool`) when `True` retrieves the value that is displayed to the user in the sheet, not the underlying data.

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
) -> Iterator[list[Any]]
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
) -> Any
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
) -> Any
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
data(value: Any)
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

A `Span` object (more information [here](#span-objects)) is returned when using square brackets on a `Sheet` like so:

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
- `widget` (`Any`) is the reference to the original sheet which created the span (this is the widget that data is set to). This can be changed to a different sheet if required e.g. `my_span.widget = new_sheet`.

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

___

#### **Managing the undo and redo stacks**

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

#### **Sheet set data function**

You can also use the `Sheet` function `set_data()`.

```python
set_data(
    *key: CreateSpanTypes,
    data: Any = None,
    undo: bool | None = None,
    emit_event: bool | None = None,
    redraw: bool = True,
    event_data: EventDataDict | None = None,
) -> EventDataDict
```
Parameters:

- `undo` when `True` adds the change to the Sheets undo stack.
- `emit_event` when `True` causes a `"<<SheetModified>>` event to occur if it has been bound, see [here](#tkinter-and-tksheet-events) for more information.

Example:
```python
self.sheet.set_data(
    "A1",
    [["",  "A",  "B",  "C"]
     ["1", "A1", "B1", "C1"],
     ["2", "A2", "B2", "C2"]],
)
```

You can clear cells/rows/columns using a [spans `clear()` function](#using-a-span-to-clear-cells) or the Sheets `clear()` function. Below is the Sheets clear function:

```python
clear(
    *key: CreateSpanTypes,
    undo: bool | None = None,
    emit_event: bool | None = None,
    redraw: bool = True,
) -> EventDataDict
```
- `undo` when `True` adds the change to the Sheets undo stack.
- `emit_event` when `True` causes a `"<<SheetModified>>` event to occur if it has been bound, see [here](#tkinter-and-tksheet-events) for more information.

#### **Insert a row into the sheet**

```python
insert_row(
    row: list[Any] | tuple[Any] | None = None,
    idx: str | int | None = None,
    height: int | None = None,
    row_index: bool = False,
    fill: bool = True,
    undo: bool = True,
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
- `emit_event` when `True` causes a `"<<SheetModified>>` event to occur if it has been bound, see [here](#tkinter-and-tksheet-events) for more information.

___

#### **Insert a column into the sheet**

```python
insert_column(
    column: list[Any] | tuple[Any] | None = None,
    idx: str | int | None = None,
    width: int | None = None,
    header: bool = False,
    fill: bool = True,
    undo: bool = True,
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
- `emit_event` when `True` causes a `"<<SheetModified>>` event to occur if it has been bound, see [here](#tkinter-and-tksheet-events) for more information.

___

#### **Insert multiple columns into the sheet**

```python
insert_columns(
    columns: list[tuple[Any] | list[Any]] | tuple[tuple[Any] | list[Any]] | int = 1,
    idx: str | int | None = None,
    widths: list[int] | tuple[int] | None = None,
    headers: bool = False,
    fill: bool = True,
    undo: bool = True,
    emit_event: bool = False,
    create_selections: bool = True,
    add_row_heights: bool = True,
    push_ops: bool = True,
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
- `emit_event` when `True` causes a `"<<SheetModified>>` event to occur if it has been bound, see [here](#tkinter-and-tksheet-events) for more information.
- `create_selections` when `True` creates a selection box for the newly inserted columns.
- `add_row_heights` when `True` creates rows if there are no pre-existing rows.
- `push_ops` when `True` increases the indexes of all cell/column options such as dropdown boxes, highlights and data formatting.

___

#### **Insert multiple rows into the sheet**

```python
insert_rows(
    rows: list[tuple[Any] | list[Any]] | tuple[tuple[Any] | list[Any]] | int = 1,
    idx: str | int | None = None,
    heights: list[int] | tuple[int] | None = None,
    row_index: bool = False,
    fill: bool = True,
    undo: bool = True,
    emit_event: bool = False,
    create_selections: bool = True,
    add_column_widths: bool = True,
    push_ops: bool = True,
    tree: bool = True,
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
- `emit_event` when `True` causes a `"<<SheetModified>>` event to occur if it has been bound, see [here](#tkinter-and-tksheet-events) for more information.
- `create_selections` when `True` creates a selection box for the newly inserted rows.
- `add_column_widths` when `True` creates columns if there are no pre-existing columns.
- `push_ops` when `True` increases the indexes of all cell/row options such as dropdown boxes, highlights and data formatting.
- `tree` is mainly used internally but when `True` and also when treeview mode is enabled it performs the necessary actions to create new ids and add them to the tree.

___

#### **Delete a row from the sheet**

```python
del_row(
    idx: int = 0,
    data_indexes: bool = True,
    undo: bool = True,
    emit_event: bool = False,
    redraw: bool = True,
) -> EventDataDict
```
Parameters:

- `idx` is the row to delete.
- `data_indexes` only applicable when there are hidden rows. When `False` it makes the `idx` represent a displayed row and not the underlying Sheet data row. When `True` the index represent a data index.
- `undo` when `True` adds the change to the Sheets undo stack.
- `emit_event` when `True` causes a `"<<SheetModified>>` event to occur if it has been bound, see [here](#tkinter-and-tksheet-events) for more information.

___

#### **Delete multiple rows from the sheet**

```python
del_rows(
    rows: int | Iterator[int],
    data_indexes: bool = True,
    undo: bool = True,
    emit_event: bool = False,
    redraw: bool = True,
) -> EventDataDict
```
Parameters:

- `rows` can be either `int` or an iterable of `int`s representing row indexes.
- `data_indexes` only applicable when there are hidden rows. When `False` it makes the `rows` indexes represent displayed rows and not the underlying Sheet data rows. When `True` the indexes represent data indexes.
- `undo` when `True` adds the change to the Sheets undo stack.
- `emit_event` when `True` causes a `"<<SheetModified>>` event to occur if it has been bound, see [here](#tkinter-and-tksheet-events) for more information.

___

#### **Delete a column from the sheet**

```python
del_column(
    idx: int = 0,
    data_indexes: bool = True,
    undo: bool = True,
    emit_event: bool = False,
    redraw: bool = True,
) -> EventDataDict
```
Parameters:

- `idx` is the column to delete.
- `data_indexes` only applicable when there are hidden columns. When `False` it makes the `idx` represent a displayed column and not the underlying Sheet data column. When `True` the index represent a data index.
- `undo` when `True` adds the change to the Sheets undo stack.
- `emit_event` when `True` causes a `"<<SheetModified>>` event to occur if it has been bound, see [here](#tkinter-and-tksheet-events) for more information.

___

#### **Delete multiple columns from the sheet**

```python
del_columns(
    columns: int | Iterator[int],
    data_indexes: bool = True,
    undo: bool = True,
    emit_event: bool = False,
    redraw: bool = True,
) -> EventDataDict
```
Parameters:

- `columns` can be either `int` or an iterable of `int`s representing column indexes.
- `data_indexes` only applicable when there are hidden columns. When `False` it makes the `columns` indexes represent displayed columns and not the underlying Sheet data columns. When `True` the indexes represent data indexes.
- `undo` when `True` adds the change to the Sheets undo stack.
- `emit_event` when `True` causes a `"<<SheetModified>>` event to occur if it has been bound, see [here](#tkinter-and-tksheet-events) for more information.

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
) -> tuple[dict[int, int], dict[int, int], EventDataDict]:
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
    undo: bool = True,
    emit_event: bool = False,
    move_heights: bool = True,
    event_data: EventDataDict | None = None,
    redraw: bool = True,
) -> tuple[dict[int, int], dict[int, int], EventDataDict]:
```
Parameters:

- `move_to` is the new start index for the rows to be moved to.
- `to_move` is a `list` of row indexes to move to that new position, they will appear in the same order provided.
- `move_data` when `True` moves not just the displayed row positions but the Sheet data as well.
- `data_indexes` is only applicable when there are hidden rows. When `False` it makes the `move_to` and `to_move` indexes represent displayed rows and not the underlying Sheet data rows. When `True` the indexes represent data indexes.
- `create_selections` creates new selection boxes based on where the rows have moved.
- `undo` when `True` adds the change to the Sheets undo stack.
- `emit_event` when `True` causes a `"<<SheetModified>>` event to occur if it has been bound, see [here](#tkinter-and-tksheet-events) for more information.
- `move_heights` when `True` also moves the displayed row lines.

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
    undo: bool = True,
    emit_event: bool = False,
    move_widths: bool = True,
    event_data: EventDataDict | None = None,
    redraw: bool = True,
) -> tuple[dict[int, int], dict[int, int], EventDataDict]:
```
Parameters:

- `move_to` is the new start index for the columns to be moved to.
- `to_move` is a `list` of column indexes to move to that new position, they will appear in the same order provided.
- `move_data` when `True` moves not just the displayed column positions but the Sheet data as well.
- `data_indexes` is only applicable when there are hidden columns. When `False` it makes the `move_to` and `to_move` indexes represent displayed columns and not the underlying Sheet data columns. When `True` the indexes represent data indexes.
- `create_selections` creates new selection boxes based on where the columns have moved.
- `undo` when `True` adds the change to the Sheets undo stack.
- `emit_event` when `True` causes a `"<<SheetModified>>` event to occur if it has been bound, see [here](#tkinter-and-tksheet-events) for more information.
- `move_widths` when `True` also moves the displayed column lines.

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
    undo: bool = True,
    emit_event: bool = False,
    redraw: bool = True,
) -> tuple[dict[int, int], dict[int, int], EventDataDict]
```
Parameters:

- `data_new_idxs` (`dict[int, int]`) must be a `dict` where the keys are the data columns to move as `int`s and the values are their new locations as `int`s.
- `disp_new_idxs` (`None | dict[int, int]`) either `None` or a `dict` where the keys are the displayed columns (basically the column widths) to move as `int`s and the values are their new locations as `int`s. If `None` then no column widths will be moved.
- `move_data` when `True` moves not just the displayed column positions but the Sheet data as well.
- `create_selections` creates new selection boxes based on where the columns have moved.
- `undo` when `True` adds the change to the Sheets undo stack.
- `emit_event` when `True` causes a `"<<SheetModified>>` event to occur if it has been bound, see [here](#tkinter-and-tksheet-events) for more information.

___

Get a mapping (`dict`) of all `old: new` column indexes.

```python
full_move_columns_idxs(data_idxs: dict[int, int]) -> dict[int, int]
```
- e.g. converts `{0: 1}` to `{0: 1, 1: 0}` if the maximum Sheet column number is `1`.

___

#### **Move any rows to new locations**

```python
mapping_move_rows(
    data_new_idxs: dict[int, int],
    disp_new_idxs: None | dict[int, int] = None,
    move_data: bool = True,
    create_selections: bool = True,
    undo: bool = True,
    emit_event: bool = False,
    redraw: bool = True,
) -> tuple[dict[int, int], dict[int, int], EventDataDict]
```
Parameters:

- `data_new_idxs` (`dict[int, int]`) must be a `dict` where the keys are the data rows to move as `int`s and the values are their new locations as `int`s.
- `disp_new_idxs` (`None | dict[int, int]`) either `None` or a `dict` where the keys are the displayed rows (basically the row heights) to move as `int`s and the values are their new locations as `int`s. If `None` then no row heights will be moved.
- `move_data` when `True` moves not just the displayed row positions but the Sheet data as well.
- `create_selections` creates new selection boxes based on where the rows have moved.
- `undo` when `True` adds the change to the Sheets undo stack.
- `emit_event` when `True` causes a `"<<SheetModified>>` event to occur if it has been bound, see [here](#tkinter-and-tksheet-events) for more information.

___

Get a mapping (`dict`) of all `old: new` row indexes.

```python
full_move_rows_idxs(data_idxs: dict[int, int]) -> dict[int, int]
```
- e.g. converts `{0: 1}` to `{0: 1, 1: 0}` if the maximum Sheet row number is `1`.

___

#### **Make all data rows the same length**

```python
equalize_data_row_lengths(include_header: bool = True) -> int
```
- Makes every list in the table have the same number of elements, goes by longest list. This will only affect the data variable, not visible columns.
- Returns the new row length for all rows in the Sheet.

---
# **Sorting the Table**

There are three built-in sorting keys to choose from but you can always create your own and use that instead. See [here](#setting-the-default-sorting-key) for more information on how to set the default sorting key.

#### **natural_sort_key**

This is the **default** sorting key for natural sorting of various Python types:

- Won't sort string version numbers.
- Will convert strings to floats.
- Will sort strings that are file paths.

Order:

0. None
1. Empty strings
2. bool
3. int, float (inc. strings that are numbers)
4. datetime (inc. strings that are dates)
5. strings (including string file paths and paths as POSIX strings) & unknown objects with __str__
6. unknown objects

#### **version_sort_key**

An alternative sorting key that respects and sorts most version numbers:

- Won't convert strings to floats.
- Will sort string version numbers.
- Will sort strings that are file paths.

0. None
1. Empty strings
2. bool
3. int, float
4. datetime (inc. strings that are dates)
5. strings (including string file paths and paths as POSIX strings) & unknown objects with __str__
6. unknown objects

#### **fast_sort_key**

A faster key for natural sorting of various Python types. This key should probably be used if you intend on sorting sheets with over a million cells:

- Won't sort strings that are dates very well.
- Won't convert strings to floats.
- Won't sort string file paths very well.
- Will do ok with string version numbers.

0. None
1. Empty strings
2. bool
3. int, float
4. datetime
5. strings (including paths as POSIX strings) & unknown objects with __str__
6. unknown objects

#### **Setting the default sorting key**

Setting the sorting key at initialization:
```python
from tksheet import Sheet, natural_sort_key

my_sheet = Sheet(parent=parent, sort_key=natural_sort_key)
```

Setting the sorting key after initialization:
```python
from tksheet import Sheet, natural_sort_key

my_sheet.set_options(sort_key=natural_sort_key)
```

Using a sorting key with a tksheet sort function call:
```python
from tksheet import Sheet, natural_sort_key

my_sheet.sort_columns(0, key=natural_sort_key)
```
- Setting the key like this will, for this call, override whatever key was set at initialization or using `set_options()`.

#### **Notes about sorting**

- The readonly functions can be used to disallow sorting of particular cells/rows/columns values.

#### **Sorting cells**

```python
sort(
    *box: CreateSpanTypes,
    reverse: bool = False,
    row_wise: bool = False,
    validation: bool = True,
    key: Callable | None = None,
    undo: bool = True,
) -> EventDataDict
```
Parameters:

- `boxes` (`CreateSpanTypes`).
- `reverse` (`bool`) if `True` sorts in reverse (descending) order.
- `row-wise` (`bool`) if `True` sorts values row-wise. Default is column-wise.
- `key` (`Callable`, `None`) if `None` then uses the default sorting key.
- `undo` (`bool`) if `True` then adds the change (if a change was made) to the undo stack.

Notes:

- Sort the values of the box columns, or the values of the box rows if `row_wise` is `True`.
- **Will not shift cell options (properties) around, only cell values.**
- The event name in `EventDataDict` for sorting table values is `"edit_table"`.
- The readonly functions can be used to disallow sorting of particular cells values.

Example:
```python
# a box of cells row-wise
# start at row 0 & col 0, up to but not including row 10 & col 10
my_sheet.sort(0, 0, 10, 10, row_wise=True)

# sort the whole table
my_sheet.sort(None)
```

#### **Sorting row values**

```python
def sort_rows(
    rows: AnyIter[int] | Span | int | None = None,
    reverse: bool = False,
    validation: bool = True,
    key: Callable | None = None,
    undo: bool = True,
) -> EventDataDict
```
Parameters:

- `rows` (`AnyIter[int]` , `Span`, `int`, `None`) the rows to sort.
- `reverse` (`bool`) if `True` then sorts in reverse (descending) order.
- `key` (`Callable`, `None`) if `None` then uses the default sorting key.
- `undo` (`bool`) if `True` then adds the change (if a change was made) to the undo stack.

Notes:

- Sorts the values of each row independently.
- **Will not shift cell options (properties) around, only cell values.**
- The event name in `EventDataDict` for sorting table values is `"edit_table"`.
- The readonly functions can be used to disallow sorting of particular rows values.

#### **Sorting column values**

```python
def sort_columns(
    columns: AnyIter[int] | Span | int | None = None,
    reverse: bool = False,
    validation: bool = True,
    key: Callable | None = None,
    undo: bool = True,
) -> EventDataDict
```
Parameters:

- `columns` (`AnyIter[int]` , `Span`, `int`, `None`) the columns to sort.
- `reverse` (`bool`) if `True` then sorts in reverse (descending) order.
- `key` (`Callable`, `None`) if `None` then uses the default sorting key.
- `undo` (`bool`) if `True` then adds the change (if a change was made) to the undo stack.

Notes:

- Sorts the values of each column independently.
- **Will not shift cell options (properties) around, only cell values.**
- The event name in `EventDataDict` for sorting table values is `"edit_table"`.
- The readonly functions can be used to disallow sorting of particular columns values.

#### **Sorting the order of all rows using a column**

```python
def sort_rows_by_column(
    column: int | None = None,
    reverse: bool = False,
    key: Callable | None = None,
    undo: bool = True,
) -> EventDataDict
```
Parameters:

- `column` (`int`, `None`) if `None` then it uses the currently selected column to sort.
- `reverse` (`bool`) if `True` then sorts in reverse (descending) order.
- `key` (`Callable`, `None`) if `None` then uses the default sorting key.
- `undo` (`bool`) if `True` then adds the change (if a change was made) to the undo stack.

Notes:

- Sorts the tree if treeview mode is active.

#### **Sorting the order of all columns using a row**

```python
def sort_columns_by_row(
    row: int | None = None,
    reverse: bool = False,
    key: Callable | None = None,
    undo: bool = True,
) -> EventDataDict
```
Parameters:

- `row` (`int`, `None`) if `None` then it uses the currently selected row to sort.
- `reverse` (`bool`) if `True` then sorts in reverse (descending) order.
- `key` (`Callable`, `None`) if `None` then uses the default sorting key.
- `undo` (`bool`) if `True` then adds the change (if a change was made) to the undo stack.

---
# **Find and Replace**

An in-built find and replace window can be enabled using `enable_bindings()`, e.g:

```python
my_sheet.enable_bindings("find", "replace")

# all bindings, including find and replace
my_sheet.enable_bindings()
```

See [enable_bindings](#enable-table-functionality-and-bindings) for more information.

There are also some `Sheet()` functions that can be utilized, shown below.

#### **Check if the find window is open**

```python
@property
find_open() -> bool
```
e.g. `find_is_open = sheet.find_open`

#### **Open or close the find window**

```python
open_find(focus: bool = False) -> Sheet
```

```python
close_find() -> Sheet
```

#### **Find and select**

```python
next_match(within: bool | None = None, find: str | None = None) -> Sheet
```

```python
prev_match(within: bool | None = None, find: str | None = None) -> Sheet
```
Parameters:

- `within` (`bool`, `None`) if `bool` then will override the find windows within selection setting. If `None` then it will use the find windows setting.
- `find` (`str`, `None`) if `str` then will override the find windows search value. If `None` then it will use the find windows search value.

Notes:

- If looking within selection then hidden rows and columns will be skipped.

#### **Replace all using mapping**

```python
replace_all(mapping: dict[str, str], within: bool = False) -> EventDataDict
```
Parameters:

- `mapping` (`dict[str, str]`) a `dict` of keys to search for and values to replace them with.
- `within` (`bool`) when `True` will only do replaces inside existing selection boxes.

Notes:

- Will do partial cell data replaces also.
- If looking within selection then hidden rows and columns will be skipped.

---
# **Highlighting Cells**

### **Creating highlights**

#### **Using spans to create highlights**

`Span` objects (more information [here](#span-objects)) can be used to highlight cells, rows, columns, the entire sheet, headers and the index.

You can use either of the following methods:
- Using a span method e.g. `span.highlight()` more information [here](#using-a-span-to-create-highlights).
- Using a sheet method e.g. `sheet.highlight(Span)`

Or if you need user inserted row/columns in the middle of highlight areas to also be highlighted you can use named spans, more information [here](#named-spans).

Whether cells, rows or columns are highlighted depends on the [`kind`](#get-a-spans-kind) of span.

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

- `key` (`CreateSpanTypes`) either a span or a type which can create a span. See [here](#creating-a-span) for more information on the types that can create a span.
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

#### **Other ways to create highlights**

**Cells**

```python
highlight_cells(
    row: int | Literal["all"] = 0,
    column: int | Literal["all"] = 0,
    cells: list[tuple[int, int]] = [],
    canvas: Literal["table", "index", "header"] = "table",
    bg: bool | None | str = False,
    fg: bool | None | str = False,
    redraw: bool = True,
    overwrite: bool = True,
) -> Sheet
```

**Rows**

```python
highlight_rows(
    rows: Iterator[int] | int,
    bg: None | str = None,
    fg: None | str = None,
    highlight_index: bool = True,
    redraw: bool = True,
    end_of_screen: bool = False,
    overwrite: bool = True,
) -> Sheet
```

**Columns**

```python
highlight_columns(
    columns: Iterator[int] | int,
    bg: bool | None | str = False,
    fg: bool | None | str = False,
    highlight_header: bool = True,
    redraw: bool = True,
    overwrite: bool = True,
) -> Sheet
```

___

### **Deleting highlights**

#### **Using spans to delete highlights**

If the highlights were created by a named span then the named span must be deleted, more information [here](#deleting-a-named-span).

Otherwise you can use either of the following methods to delete/remove highlights:

- Using a span method e.g. `span.dehighlight()` more information [here](#using-a-span-to-delete-highlights).
- Using a sheet method e.g. `sheet.dehighlight(Span)` details below:

```python
dehighlight(
    *key: CreateSpanTypes,
    redraw: bool = True,
) -> Span
```
Parameters:

- `key` (`CreateSpanTypes`) either a span or a type which can create a span. See [here](#creating-a-span) for more information on the types that can create a span.

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

#### **Other ways to delete highlights**

**Cells**

```python
dehighlight_cells(
    row: int | Literal["all"] = 0,
    column: int = 0,
    cells: list[tuple[int, int]] = [],
    canvas: Literal["table", "row_index", "header"] = "table",
    all_: bool = False,
    redraw: bool = True,
) -> Sheet
```

**Rows**

```python
dehighlight_rows(
    rows: list[int] | Literal["all"] = [],
    redraw: bool = True,
) -> Sheet
```

**Columns**

```python
dehighlight_columns(
    columns: list[int] | Literal["all"] = [],
    redraw: bool = True,
) -> Sheet
```

**All**

```python
dehighlight_all(
    cells: bool = True,
    rows: bool = True,
    columns: bool = True,
    header: bool = True,
    index: bool = True,
    redraw: bool = True,
) -> Sheet
```

---
# **Dropdown Boxes**

#### **Creating dropdown boxes**

`Span` objects (more information [here](#span-objects)) can be used to create dropdown boxes for cells, rows, columns, the entire sheet, headers and the index.

You can use either of the following methods:

- Using a span method e.g. `span.dropdown()` more information [here](#using-a-span-to-create-dropdown-boxes).
- Using a sheet method e.g. `sheet.dropdown(Span)`

Or if you need user inserted row/columns in the middle of areas with dropdown boxes to also have dropdown boxes you can use named spans, more information [here](#named-spans).

Whether dropdown boxes are created for cells, rows or columns depends on the [`kind`](#get-a-spans-kind) of span.

```python
dropdown(
    *key: CreateSpanTypes,
    values: list = [],
    edit_data: bool = True,
    set_values: dict[tuple[int, int] | int, Any] | None = None,
    set_value: Any = None,
    state: Literal["normal", "readonly", "disabled"] = "normal",
    redraw: bool = True,
    selection_function: Callable | None = None,
    modified_function: Callable | None = None,
    search_function: Callable = dropdown_search_function,
    validate_input: bool = True,
    text: None | str = None,
) -> Span
```
Notes:

- `selection_function` / `modified_function` (`Callable`, `None`) parameters require either `None` or a function. The function you use needs at least one argument because tksheet will send information to your function about the triggered dropdown.
- When a user selects an item from the dropdown box the sheet will set the underlying cells data to the selected item, to bind this event use either the `selection_function` argument or see the function `extra_bindings()` with binding `"end_edit_cell"` [here](#table-functionality-and-bindings).

Parameters:

- `key` (`CreateSpanTypes`) either a span or a type which can create a span. See [here](#creating-a-span) for more information on the types that can create a span.
- `values` are the values to appear in a list view type interface when the dropdown box is open.
- `edit_data` when `True` makes edits in the table, header or index (depending on the span) based on `set_values`/`set_value`.
- `set_values` when combined with `edit_data=True` allows a `dict` to be provided of data coordinates (`tuple[int, int]` for a cell span or `int` for a row/column span) as `key`s and values to set the cell at that coordinate to.
    - e.g. `set_values={(0, 0): "new value for A1"}`.
    - The idea behind this parameter is that an entire column or row can have individual cell values and is not set to `set_value` alone.
- `set_value` when combined with `edit_data=True` sets every cell in the span to the value provided. If left as `None` and if `set_values` is also `None` then the topmost value from `values` will be used or if not `values` then `""`.
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

If the dropdown boxes were created by a named span then the named span must be deleted, more information [here](#deleting-a-named-span).

Otherwise you can use either of the following methods to delete/remove dropdown boxes.
- Using a span method e.g. `span.del_dropdown()` more information [here](#using-a-span-to-delete-dropdown-boxes).
- Using a sheet method e.g. `sheet.del_dropdown(Span)` details below:

```python
del_dropdown(
    *key: CreateSpanTypes,
    redraw: bool = True,
) -> Span
```
Parameters:

- `key` (`CreateSpanTypes`) either a span or a type which can create a span. See [here](#creating-a-span) for more information on the types that can create a span.

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
    set_value: Any = None,
) -> Sheet
```

```python
set_header_dropdown_values(
    c: int = 0,
    set_existing_dropdown: bool = False,
    values: list = [],
    set_value: Any = None,
) -> Sheet
```

```python
set_index_dropdown_values(
    r: int = 0,
    set_existing_dropdown: bool = False,
    values: list = [],
    set_value: Any = None,
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
close_dropdown(r: int | None = None, c: int | None = None) -> Sheet
```

```python
close_header_dropdown(c: int | None = None) -> Sheet
```

```python
close_index_dropdown(r: int | None = None) -> Sheet
```
Notes:
- Also destroys any opened text editor windows.

---
# **Check Boxes**

#### **Creating check boxes**

`Span` objects (more information [here](#span-objects)) can be used to create check boxes for cells, rows, columns, the entire sheet, headers and the index.

You can use either of the following methods:

- Using a span method e.g. `span.checkbox()` more information [here](#using-a-span-to-create-check-boxes).
- Using a sheet method e.g. `sheet.checkbox(Span)`

Or if you need user inserted row/columns in the middle of areas with check boxes to also have check boxes you can use named spans, more information [here](#named-spans).

Whether check boxes are created for cells, rows or columns depends on the [`kind`](#get-a-spans-kind) of span.

```python
checkbox(
    *key: CreateSpanTypes,
    edit_data: bool = True,
    checked: bool | None = None,
    state: Literal["normal", "disabled"] = "normal",
    redraw: bool = True,
    check_function: Callable | None = None,
    text: str = "",
) -> Span
```
Parameters:

- `key` (`CreateSpanTypes`) either a span or a type which can create a span. See [here](#creating-a-span) for more information on the types that can create a span.
- `edit_data` when `True` edits the underlying cell data to either `checked` if `checked` is a `bool` or tries to convert the existing cell data to a `bool`.
- `checked` is the initial creation value to set the box to, if `None` then and `edit_data` is `True` then it will try to convert the underlying cell data to a `bool`.
- `state` can be `"normal"` or `"disabled"`. If `"disabled"` then color will be same as table grid lines, else it will be the cells text color.
- `check_function` can be used to trigger a function when the user clicks a checkbox.
- `text` displays text next to the checkbox in the cell, but will not be used as data, data will either be `True` or `False`.

Notes:

- `check_function` (`Callable`, `None`) requires either `None` or a function. The function you use needs at least one argument because when the checkbox is set it will send information to your function about the clicked checkbox.
- Use `highlight_cells()` or rows or columns to change the color of the checkbox.
- Check boxes are always left aligned despite any align settings.
- To get the current checkbox value either:
    - Get the cell data, more information [here](#getting-sheet-data).
    - Use the parameter `check_function` with a function of your own creation to be called when the checkbox is set by the user.

Example:
```python
self.sheet.checkbox(
    "D",
    checked=True,
)
```

___

#### **Deleting check boxes**

If the check boxes were created by a named span then the named span must be deleted, more information [here](#deleting-a-named-span).

Otherwise you can use either of the following methods to delete/remove check boxes:

- Using a span method e.g. `span.del_checkbox()` more information [here](#using-a-span-to-delete-check-boxes).
- Using a sheet method e.g. `sheet.del_checkbox(Span)` details below:

```python
del_checkbox(
    *key: CreateSpanTypes,
    redraw: bool = True,
) -> Span
```
Parameters:

- `key` (`CreateSpanTypes`) either a span or a type which can create a span. See [here](#creating-a-span) for more information on the types that can create a span.

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

- An full explanation of what arguments to use for the `formatter_options` parameter can be found [here](#generic-formatter).
- A demonstration of all the built-in and custom formatters can be found [here](#example-using-and-creating-formatters).

### **Creation and deletion of data formatting rules**

#### **Creating a data format rule**

`Span` objects (more information [here](#span-objects)) can be used to format data for cells, rows, columns and the entire sheet.

You can use either of the following methods:
- Using a span method e.g. `span.format()` more information [here](#using-a-span-to-format-data).
- Using a sheet method e.g. `sheet.format(Span)`

Or if you need user inserted row/columns in the middle of areas with data formatting to also be formatted you can use named spans, more information [here](#named-spans).

Whether data is formatted for cells, rows or columns depends on the [`kind`](#get-a-spans-kind) of span.

```python
format(
    *key: CreateSpanTypes,
    formatter_options: dict = {},
    formatter_class: Any = None,
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
3. When getting data take careful note of the `get_displayed` options, as these are the difference between getting the actual formatted data and what is displayed on the table GUI.

Parameters:

- `key` (`CreateSpanTypes`) either a span or a type which can create a span. See [here](#creating-a-span) for more information on the types that can create a span.
- `formatter_options` (`dict`) a dictionary of keyword options/arguements to pass to the formatter, see [here](#formatters) for information on what argument to use.
- `formatter_class` (`class`) in case you want to use a custom class to store functions and information as opposed to using the built-in methods.
- `**kwargs` any additional keyword options/arguements to pass to the formatter.

___

#### **Deleting a data format rule**

If the data format rule was created by a named span then the named span must be deleted, more information [here](#deleting-a-named-span).

Otherwise you can use either of the following methods to delete/remove data formatting rules:

- Using a span method e.g. `span.del_format()` more information [here](#using-a-span-to-delete-check-boxes).
- Using a sheet method e.g. `sheet.del_format(Span)` details below:

```python
del_format(
    *key: CreateSpanTypes,
    clear_values: bool = False,
    redraw: bool = True,
) -> Span
```
- `key` (`CreateSpanTypes`) either a span or a type which can create a span. See [here](#creating-a-span) for more information on the types that can create a span.
- `clear_values` (`bool`) if true, all the cells covered by the span will have their values cleared.

___

#### **Delete all formatting**

```python
del_all_formatting(clear_values: bool = False) -> Sheet
```
- `clear_values` (`bool`) if true, all the sheets cell values will be cleared.

___

#### **Reapply formatting to entire sheet**

```python
reapply_formatting() -> Sheet
```
- Useful if you have manually changed the entire sheets data using `sheet.MT.data = ` and want to reformat the sheet using any existing formatting you have set.

___

#### **Check if a cell is formatted**

```python
formatted(r: int, c: int) -> dict
```
- If the cell is formatted function returns a `dict` with all the format keyword arguments. The `dict` will be empty if the cell is not formatted.

___

### **Formatters**

In addition to the [generic](#generic-formatter) formatter, `tksheet` provides formatters for many different data types.

A basic example showing how formatting some columns as `float` might be done:

```python
import tkinter as tk

from tksheet import (
    Sheet,
    float_formatter,
)
from tksheet import (
    num2alpha,
)


class demo(tk.Tk):
    def __init__(self):
        super().__init__()
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.sheet = Sheet(
            self,
            data=[[f"{r}", f"{r}"] for r in range(5)],
            expand_sheet_if_paste_too_big=True,
            theme="dark blue",
        )
        """
        Format example
        """
        # some keyword arguments inside float_formatter()
        self.sheet.format(
            num2alpha(0), # column A
            formatter_options=float_formatter(
                decimals=1,
                nullable=True,
            ),
        )
        # some keyword arguments outside
        # of float_formatter() instead
        self.sheet.format(
            "B", # column B
            formatter_options=float_formatter(),
            decimals=3,
            nullable=False,
        )
        """
        Rest of code
        """
        self.sheet.grid(row=0, column=0, sticky="nswe")
        self.sheet.enable_bindings(
            "all",
            "ctrl_select",
            "edit_header",
            "edit_index",
        )


app = demo()
app.mainloop()
```

**You can use any of the following formatters as an argument for the parameter `formatter_options`.**

A full list of keyword arguments available to these formatters is described [here](#generic-formatter).

___

#### **Int Formatter**

The `int_formatter` is the basic configuration for a simple interger formatter.

```python
int_formatter(
    datatypes: tuple[Any] | Any = int,
    format_function: Callable = to_int,
    to_str_function: Callable = to_str,
    invalid_value: Any = "NaN",
    **kwargs,
) -> dict
```
Parameters:

- `format_function` (`function`) a function that takes a string and returns an `int`. By default, this is set to the in-built `tksheet.to_int`. This function will always convert float-likes to its floor, for example `"5.9"` will be converted to `5`.
- `to_str_function` (`function`) By default, this is set to the in-built `tksheet.to_str`, which is a very basic function that will displace the default string representation of the value.
- A full explanation of keyword arguments available is described [here](#generic-formatter).

Example:
```python
sheet.format_cell(0, 0, formatter_options = tksheet.int_formatter())
```

___

#### **Float Formatter**

The `float_formatter` is the basic configuration for a simple float formatter. It will always round float-likes to the specified number of decimal places, for example `"5.999"` will be converted to `"6.0"` if `decimals = 1`.

```python
float_formatter(
    datatypes: tuple[Any] | Any = float,
    format_function: Callable = to_float,
    to_str_function: Callable = float_to_str,
    invalid_value: Any = "NaN",
    decimals: int = 2,
    **kwargs
) -> dict
```
Parameters:

- `format_function` (`function`) a function that takes a string and returns a `float`. By default, this is set to the in-built `tksheet.to_float`.
- `to_str_function` (`function`) By default, this is set to the in-built `tksheet.float_to_str`, which will display the float to the specified number of decimal places.
- `decimals` (`int`, `None`) the number of decimal places to round to. Defaults to `2`.
- A full explanation of keyword arguments available is described [here](#generic-formatter).

Example:
```python
sheet.format_cell(0, 0, formatter_options = tksheet.float_formatter(decimals = None)) # A float formatter with maximum float() decimal places
```

___

#### **Percentage Formatter**

The `percentage_formatter` is the basic configuration for a simple percentage formatter. It will always round float-likes as a percentage to the specified number of decimal places, for example `"5.999%"` will be converted to `"6.0%"` if `decimals = 1`.

```python
percentage_formatter(
    datatypes: tuple[Any] | Any = float,
    format_function: Callable = to_percentage,
    to_str_function: Callable = percentage_to_str,
    invalid_value: Any = "NaN",
    decimals: int = 0,
    **kwargs,
) -> dict
```
Parameters:

- `format_function` (`function`) a function that takes a string and returns a `float`. By default, this is set to the in-built `tksheet.to_percentage`. This function will always convert percentages to their decimal equivalent, for example `"5%"` will be converted to `0.05`.
- `to_str_function` (`function`) By default, this is set to the in-built `tksheet.percentage_to_str`, which will display the float as a percentage to the specified number of decimal places. For example, `0.05` will be displayed as `"5.0%"`.
- `decimals` (`int`) the number of decimal places to round to. Defaults to `0`.
- A full explanation of keyword arguments available is described [here](#generic-formatter).

Example:
```python
sheet.format_cell(0, 0, formatter_options=tksheet.percentage_formatter(decimals=1)) # A percentage formatter with 1 decimal place
```

**Note:**

By default `percentage_formatter()` converts user entry `21` to `2100%` and `21%` to `21%`. An example where `21` is converted to `21%` instead is shown below:

```python
# formats column A as percentage

# uses:
# format_function=alt_to_percentage
# to_str_function=alt_percentage_to_str
sheet.format(
    "A",
    formatter_options=tksheet.percentage_formatter(
        format_function=alt_to_percentage,
        to_str_function=alt_percentage_to_str,
    )
)
```

___

#### **Bool Formatter**

```python
bool_formatter(
    datatypes: tuple[Any] | Any = bool,
    format_function: Callable = to_bool,
    to_str_function: Callable = bool_to_str,
    invalid_value: Any = "NA",
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
- A full explanation of keyword arguments available is described [here](#generic-formatter).

Example:
```python
# A bool formatter with custom truthy and falsy values to account for aussie and kiwi slang
sheet.format_cell(0, 0, formatter_options = tksheet.bool_formatter(truthy = tksheet.truthy | {"nah yeah"}, falsy = tksheet.falsy | {"yeah nah"}))
```

___

#### **Generic Formatter**

```python
formatter(
    datatypes: tuple[Any] | Any,
    format_function: Callable,
    to_str_function: Callable = to_str,
    invalid_value: Any = "NaN",
    nullable: bool = True,
    pre_format_function: Callable | None = None,
    post_format_function: Callable | None = None,
    clipboard_function: Callable | None = None,
    **kwargs,
) -> dict
```

This is the generic formatter options interface. You can use this to create your own custom formatters. The following options are available.

Note that all these options can also be passed to the `Sheet()` format functions as keyword arguments and are available as attributes for all formatters.

You can also provide functions of your own creation for all the below arguments which take functions if you require.

- `datatypes` (`tuple[Any], Any`) a list of datatypes that the formatter will accept. For example, `datatypes=(int, float)` will accept integers and floats.
- `format_function` (`function`) a function that takes a string and returns a value of the desired datatype. For example, `format_function = int` will convert a string to an integer.
    - Exceptions are suppressed (ignored) for this function. If an `Exception` (error) occurs then it merely continues on to the `post_format_function` (if it is `Callable`) and then finally returns the value.
- `invalid_value` (`Any`) is the value to display in the cell if the cell value's `type` is not in `datatypes`.
    - `invalid_value` must have a `__str__` method - most python types do.
    - Pythons `isinstance()` is used to determine the `type`.
    - `invalid_value` is also relevant if using `Sheet` data retrieval functions with `tdisp` or `get_displayed` parameters as `True` as these retrieve the displayed value.
    - The cell's underlying value will not be set to `invalid_value` if it's invalid - it's only for display or displayed data retrieval.
- `to_str_function` (`function`) a function that takes a value of the desired datatype and returns a string. This determines how the formatter displays its data on the table. For example, `to_str_function = str` will convert an integer to a string. Defaults to `tksheet.to_str`.
    - If the cell value's `type` is not in `datatypes` then this function will **not** be called, the `invalid_value` will be returned instead.
    - If the cell's value is `None` and the format has `nullable=True` then this function will **not** be called, an empty `str` will be returned instead.
- `nullable` (`bool`) if `True` then it guarantees `type(None)` will be in `datatypes`, effectively allowing the cell's value to be `None`.
    - When `True` just before the given `format_function` is run the cell's value is checked to see if it's in a `set` named `nonelike`. You can import and modify this `set` using `from tksheet import nonelike`. If the value is in `nonelike` then it is set to `None` before being sent to the `format_function`.
    - When `False` `type(None)` is not allowed to be in `datatypes`. If the cell for some reason ends up as `None` then it will display as the `invalid_value`.
    - When `False` the earlier described `nonelike` check does not occur.
- `pre_format_function` (`function`) This function is called before the `format_function` and can be used to modify the cells value before it is formatted using `format_function`. This can be useful, for example, if you want to strip out unwanted characters or convert a string to a different format before converting it to the desired datatype.
    - Exceptions are NOT suppressed for this function.
- `post_format_function` (`function`) a function that takes a value **which might not be of the desired datatype, e.g. `None` if the cell is nullable and empty** and if successful returns a value of the desired datatype or if not successful returns the input value. This function is called after the `format_function` and can be used to modify the output value after it is converted to the desired datatype. This can be useful if you want to round a float for example.
    - Exceptions are NOT suppressed for this function.
- `clipboard_function` (`function`) a function that takes a value of the desired datatype and returns a string. This function is called when the cell value is copied to the clipboard. This can be useful if you want to convert a value to a different format before it is copied to the clipboard.
    - Exceptions are NOT suppressed for this function.
- `**kwargs` any additional keyword options/arguements to pass to the formatter. These keyword arguments will be passed to the `format_function`, `to_str_function`, and the `clipboard_function`. These can be useful if you want to specifiy any additional formatting options, such as the number of decimal places to round to.

___

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

For those wanting even more customisation of their formatters there is also the option of creating a custom formatter class.

This is a more advanced topic and is not covered here, but it's recommended to create a new class which is a subclass of `tksheet.Formatter` and override the methods to customise. This custom class can then be passed to the `format_cells()` `formatter_class` parameter.

---
# **Readonly Cells**

#### **Creating a readonly rule**

`Span` objects (more information [here](#span-objects)) can be used to create readonly rules for cells, rows, columns, the entire sheet, headers and the index.

You can use either of the following methods:

- Using a span method e.g. `span.readonly()` more information [here](#using-a-span-to-set-cells-to-read-only).
- Using a sheet method e.g. `sheet.readonly(Span)`

Or if you need user inserted row/columns in the middle of areas with a readonly rule to also have a readonly rule you can use named spans, more information [here](#named-spans).

Whether cells, rows or columns are readonly depends on the [`kind`](#get-a-spans-kind) of span.

```python
readonly(
    *key: CreateSpanTypes,
    readonly: bool = True,
) -> Span
```
Parameters:

- `key` (`CreateSpanTypes`) either a span or a type which can create a span. See [here](#creating-a-span) for more information on the types that can create a span.
- `readonly` (`bool`) `True` to create a rule and `False` to delete one created without the use of named spans.

___

#### **Deleting a readonly rule**

If the readonly rule was created by a named span then the named span must be deleted, more information [here](#deleting-a-named-span).

Otherwise you can use either of the following methods to delete/remove readonly rules:

- Using a span method e.g. `span.readonly()` with the keyword argument `readonly=False` more information [here](#using-a-span-to-set-cells-to-read-only).
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

- `key` (`CreateSpanTypes`) either a span or a type which can create a span. See [here](#creating-a-span) for more information on the types that can create a span.
- `readonly` (`bool`) `True` to create a rule and `False` to delete one created without the use of named spans.

---
# **Text Font and Alignment**

### **Font**

- Font arguments require a three tuple e.g. `("Arial", 12, "normal")` or `("Arial", 12, "bold")` or `("Arial", 12, "italic")`.

**Set the table font**

```python
font(newfont: tuple[str, int, str] | None = None) -> tuple[str, int, str]
```

**Set the index font**

```python
index_font(newfont: tuple[str, int, str] | None = None) -> tuple[str, int, str]
```

**Set the header font**

```python
header_font(newfont: tuple[str, int, str] | None = None) -> tuple[str, int, str]
```

**Set the in-built popup menu font**

```python
popup_menu_font(newfont: tuple[str, int, str] | None = None) -> tuple[str, int, str]
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

`Span` objects (more information [here](#span-objects)) can be used to create text alignment rules for cells, rows, columns, the entire sheet, headers and the index.

You can use either of the following methods:

- Using a span method e.g. `span.align()` more information [here](#using-a-span-to-create-text-alignment-rules).
- Using a sheet method e.g. `sheet.align(Span)`

Or if you need user inserted row/columns in the middle of areas with an alignment rule to also have an alignment rule you can use named spans, more information [here](#named-spans).

Whether cells, rows or columns are affected depends on the [`kind`](#get-a-spans-kind) of span.

```python
align(
    *key: CreateSpanTypes,
    align: str | None = None,
    redraw: bool = True,
) -> Span
```
Parameters:

- `key` (`CreateSpanTypes`) either a span or a type which can create a span. See [here](#creating-a-span) for more information on the types that can create a span.
- `align` (`str`, `None`) must be one of the following:
    - `"w"`, `"west"`, `"left"`
    - `"e"`, `"east"`, `"right"`
    - `"c"`, `"center"`, `"centre"`

#### **Deleting a specific text alignment rule**

If the text alignment rule was created by a named span then the named span must be deleted, more information [here](#deleting-a-named-span).

Otherwise you can use either of the following methods to delete/remove specific text alignment rules:

- Using a span method e.g. `span.del_align()` more information [here](#using-a-span-to-delete-text-alignment-rules).
- Using a sheet method e.g. `sheet.del_align(Span)` details below:

```python
del_align(
    *key: CreateSpanTypes,
    redraw: bool = True,
) -> Span
```
Parameters:

- `key` (`CreateSpanTypes`) either a span or a type which can create a span. See [here](#creating-a-span) for more information on the types that can create a span.

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
# **Getting Cell Properties**

The below functions can be used to retrieve cell options/properties such as highlights, format, readonly etc.

#### **Get table cell properties**

Retrieve options for a single cell in the main table. Also retrieves any row/column options impacting that cell.

```Python
props(
    row: int,
    column: int | str,
    key: None
    | Literal[
        "format",
        "highlight",
        "dropdown",
        "checkbox",
        "readonly",
        "align",
    ] = None,
    cellops: bool = True,
    rowops: bool = True,
    columnops: bool = True,
) -> dict
```
Parameters:

- `row` only `int`.
- `column` `int` or `str` e.g. `"A"` is index `0`.
- `key`:
    - If left as `None` then all existing properties for that cell will be returned in a `dict`.
    - If using a `str` e.g. `"highlight"` it will only look for highlight properties for that cell.
- `cellops` when `True` will look for cell options for the cell.
- `rowops` when `True` will look for row options for the cell.
- `columnops` when `True` will look for column options for the cell.

Example:

```python
# making column B, including header read only
sheet.readonly(sheet["B"].options(header=True))

# checking if row 0, column 1 (B) is readonly:
cell_is_readonly = sheet.props(0, 1, "readonly")

# can also use a string for the column:
cell_is_readonly = sheet.props(0, "b", "readonly")
```

___

#### **Get index cell properties**

Retrieve options for a single cell in the index.

```python
index_props(
    row: int,
    key: None
    | Literal[
        "format",
        "highlight",
        "dropdown",
        "checkbox",
        "readonly",
        "align",
    ] = None,
) -> dict
```
Parameters:

- `row` only `int`.
- `key`:
    - If left as `None` then all existing properties for that cell will be returned in a `dict`.
    - If using a `str` e.g. `"highlight"` it will only look for highlight properties for that cell.

___

#### **Get header cell properties**

Retrieve options for a single cell in the header.

```python
header_props(
    column: int | str,
    key: None
    | Literal[
        "format",
        "highlight",
        "dropdown",
        "checkbox",
        "readonly",
        "align",
    ] = None,
) -> dict
```
Parameters:

- `column` only `int`.
- `key`:
    - If left as `None` then all existing properties for that cell will be returned in a `dict`.
    - If using a `str` e.g. `"highlight"` it will only look for highlight properties for that cell.

___

#### **Other properties functions**

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

---
# **Getting Selected Cells**

All selected cell/box getting functions return or generate **displayed** cell coordinates.

- Displayed cell coordinates ignore hidden rows/columns when indexing cells.
- Data cell coordinates include hidden rows/columns in indexing cells.

#### **Get the currently selected cell**

This is always a single cell of displayed indices. If you have hidden rows or columns you can change the integers to data indices using the following functions:

- [Change a row](#displayed-row-index-to-data)
- [Change a column](#displayed-column-index-to-data)

```python
get_currently_selected() -> tuple | Selected
```
Notes:

- Returns either:
    - `namedtuple` of `(row, column, type_, box, iid, fill_iid)`.
        - `type_` can be `"rows"`, `"columns"` or `"cells"`.
        - `box` `tuple[int, int, int, int]` are the coordinates of the box that the currently selected box is attached to.
            - `(from row, from column, up to but not including row, up to but not including column)`.
        - `iid` is the canvas item id of the currently selected box.
        - `fill_iid` is the canvas item id of the box that the currently selected box is attached to.
    - An empty `tuple` if nothing is selected.
- Can also use `sheet.selected` as shorter `@property` version of the function.

Example:
```python
currently_selected = self.sheet.get_currently_selected()
if currently_selected:
    row = currently_selected.row
    column = currently_selected.column
    type_ = currently_selected.type_

if self.sheet.selected:
    ...
```

___

#### **Get selected rows**

```python
get_selected_rows(
    get_cells: bool = False,
    get_cells_as_rows: bool = False,
    return_tuple: bool = False,
) -> tuple[int] | tuple[tuple[int, int]] | set[int] | set[tuple[int, int]]
```
- Returns displayed indexes.

___

#### **Get selected columns**

```python
get_selected_columns(
    get_cells: bool = False,
    get_cells_as_columns: bool = False,
    return_tuple: bool = False,
) -> tuple[int] | tuple[tuple[int, int]] | set[int] | set[tuple[int, int]]
```
- Returns displayed indexes.

___

#### **Get selected cells**

```python
get_selected_cells(
    get_rows: bool = False,
    get_columns: bool = False,
    sort_by_row: bool = False,
    sort_by_column: bool = False,
    reverse: bool = False,
) -> list[tuple[int, int]] | set[tuple[int, int]]
```
- Returns displayed coordinates.

___

#### **Generate selected cells**

```python
gen_selected_cells(
    get_rows: bool = False,
    get_columns: bool = False,
) -> Generator[tuple[int, int]]
```
- Generates displayed coordinates.

___

#### **Get all selection boxes**

```python
get_all_selection_boxes() -> tuple[tuple[int, int, int, int]]
```
- Returns displayed coordinates.

___

#### **Get all selection boxes and their types**

```python
get_all_selection_boxes_with_types() -> list[tuple[tuple[int, int, int, int], str]]
```

Equivalent to `get_all_selection_boxes_with_types()` but shortened as `@property`.

```python
@property
boxes() -> list[tuple[tuple[int, int, int, int], str]]
```

___

#### **Check if cell is selected**

```python
cell_selected(
    r: int,
    c: int,
    rows: bool = False,
    columns: bool = False,
) -> bool
```
- `rows` if `True` also checks if provided cell is part of a selected row.
- `columns` if `True` also checks if provided cell is part of a selected column.

___

#### **Check if row is selected**

```python
row_selected(r: int, cells: bool = False) -> bool
```
- `cells` if `True` also checks if provided row is selected as part of a cell selection box.

___

#### **Check if column is selected**

```python
column_selected(c: int, cells: bool = False) -> bool
```
- `cells` if `True` also checks if provided column is selected as part of a cell selection box.

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
get_ctrl_x_c_boxes(nrows: bool = True) -> tuple[dict[tuple[int, int, int, int], str], int]
```

___

```python
@property
ctrl_boxes() -> dict[tuple[int, int, int, int], str]
```

___

```python
get_selected_min_max() -> tuple[int, int, int, int] | tuple[None, None, None, None]
```
- returns `(min_y, min_x, max_y, max_x)` of any selections including rows/columns.

---
# **Modifying Selected Cells**

All selected cell/box setting functions use **displayed** cell coordinates.

- Displayed cell coordinates ignore hidden rows/columns when indexing cells.
- Data cell coordinates include hidden rows/columns in indexing cells.

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

Deselect a specific cell, row or column.

```python
deselect(
    row: int | None | Literal["all"] = None,
    column: int | None = None,
    cell: tuple | None = None,
    redraw: bool = True,
) -> Sheet
```
- Leave parameters as `None`, or `row` as `"all"` to deselect everything.
- Can also close text editors and open dropdowns.

___

Deselect any cell, row or column selection box conflicting with `rows` and/or `columns`.

```python
deselect_any(
    rows: Iterator[int] | int | None,
    columns: Iterator[int] | int | None,
    redraw: bool = True,
) -> Sheet
```
- Leave parameters as `None` to deselect everything.
- Can also close text editors and open dropdowns.

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
@boxes.setter
boxes(boxes: Sequence[tuple[tuple[int, int, int, int], str]])
```
- Can be used to set the Sheets selection boxes, deselects everything before setting.

Example:
```python
sheet.boxes = [
    ((0, 0, 3, 3), "cells"),
    ((4, 0, 5, 10), "rows"),
]
```
- The above would select a cells box from cell `A1` up to and including cell `C3` and row `4` (in python index, `5` as excel index) where the sheet has 10 columns.
- The `str` in the type hint should be either `"cells"` or `"rows"` or `"columns"`.

___

```python
recreate_all_selection_boxes() -> Sheet
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
set_all_cell_sizes_to_text(
    redraw: bool = True,
    width: int | None = None,
    slim: bool = False,
) -> tuple[list[float], list[float]]
```
- Returns the Sheets row positions and column positions in that order.
- `width` a minimum width for all column widths set using this function.
- `slim` column widths will be set precisely to text width and not add any extra space.

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

#### **Get a rows text height**

```python
get_row_text_height(
    row: int,
    visible_only: bool = False,
    only_if_too_small: bool = False,
) -> int
```
- Returns a height in pixels which will fit all text in the specified row.
- `visible_only` if `True` only measures rows visible on the Sheet.
- `only_if_too_small` if `True` will only return a new height if the current row height is too short to accomodate its text.

___

#### **Get a columns text width**

```python
get_column_text_width(
    column: int,
    visible_only: bool = False,
    only_if_too_small: bool = False,
) -> int
```
- Returns a width in pixels which will fit all text in the specified column.
- `visible_only` if `True` only measures columns visible on the Sheet.
- `only_if_too_small` if `True` will only return a new width if the current column width is too thin to accomodate its text.

___

#### **Set or reset displayed column widths/lines**

```python
set_column_widths(
    column_widths: Iterator[int, float] | None = None,
    canvas_positions: bool = False,
    reset: bool = False,
) -> Sheet
```

___

#### **Set or reset displayed row heights/lines**

```python
set_row_heights(
    row_heights: Iterator[int, float] | None = None,
    canvas_positions: bool = False,
    reset: bool = False,
) -> Sheet
```

___

#### **Set the width of the index to a str or existing values**

```python
set_width_of_index_to_text(text: None | str = None, *args, **kwargs) -> Sheet
```
- `text` (`str`, `None`) provide a `str` to set the width to or use `None` to set it to existing values in the index.

___

#### **Set the width of the index to a value in pixels**

```python
set_index_width(pixels: int, redraw: bool = True) -> Sheet
```
- Note that it disables auto resizing of index. Use [`set_options()`](#sheet-options) to restore auto resizing.

___

#### **Set the height of the header to a str or existing values**

```python
set_height_of_header_to_text(text: None | str = None) -> Sheet
```
- `text` (`str`, `None`) provide a `str` to set the height to or use `None` to set it to existing values in the header.

___

#### **Set the height of the header to a value in pixels**

```python
set_header_height_pixels(pixels: int, redraw: bool = True) -> Sheet
```

___

#### **Set the height of the header to accomodate a number of lines**

```python
set_header_height_lines(nlines: int, redraw: bool = True) -> Sheet
```

___

#### **Delete displayed column lines**

```python
del_column_position(idx: int, deselect_all: bool = False) -> Sheet
del_column_positions(idxs: Iterator[int] | None = None) -> Sheet
```

___

#### **Delete displayed row lines**

```python
del_row_position(idx: int, deselect_all: bool = False) -> Sheet
del_row_positions(idxs: Iterator[int] | None = None) -> Sheet
```

___

#### **Insert a displayed row line**

```python
insert_row_position(
    idx: Literal["end"] | int = "end",
    height: int | None = None,
    deselect_all: bool = False,
    redraw: bool = False,
) -> Sheet
```

___

#### **Insert a displayed column line**

```python
insert_column_position(
    idx: Literal["end"] | int = "end",
    width: int | None = None,
    deselect_all: bool = False,
    redraw: bool = False,
) -> Sheet
```

___

#### **Insert multiple displayed row lines**

```python
insert_row_positions(
    idx: Literal["end"] | int = "end",
    heights: Sequence[float] | int | None = None,
    deselect_all: bool = False,
    redraw: bool = False,
) -> Sheet
```

___

#### **Insert multiple displayed column lines**

```python
insert_column_positions(
    idx: Literal["end"] | int = "end",
    widths: Sequence[float] | int | None = None,
    deselect_all: bool = False,
    redraw: bool = False,
) -> Sheet
```

___

#### **Set the number of displayed row lines and column lines**

```python
sheet_display_dimensions(
    total_rows: int | None =None,
    total_columns: int | None =None,
) -> tuple[int, int] | Sheet
```

___

#### **Move a displayed row line**

```python
move_row_position(row: int, moveto: int) -> Sheet
```

___

#### **Move a displayed column line**

```python
move_column_position(column: int, moveto: int) -> Sheet
```

___

#### **Get a list of default column width values**

```python
get_example_canvas_column_widths(total_cols: int | None = None) -> list[float]
```

___

#### **Get a list of default row height values**

```python
get_example_canvas_row_heights(total_rows: int | None = None) -> list[float]
```

___

#### **Verify a list of row heights or canvas y coordinates**

```python
verify_row_heights(row_heights: list[float], canvas_positions: bool = False) -> bool
```

___

#### **Verify a list of column widths or canvas x coordinates**

```python
verify_column_widths(column_widths: list[float], canvas_positions: bool = False) -> bool
```

___

#### **Make a row height valid**

```python
valid_row_height(height: int) -> int
```

___

#### **Make a column width valid**

```python
valid_column_width(width: int) -> int
```

___

#### **Get visible rows**

```python
@property
visible_rows() -> tuple[int, int]
```
- Returns start row, end row
- e.g. `start_row, end_row = sheet.visible_rows`

___

#### **Get visible columns**

```python
@property
visible_columns() -> tuple[int, int]
```
- Returns start column, end column
- e.g. `start_column, end_column = sheet.visible_columns`

---
# **Identifying Bound Event Mouse Position**

The below functions require a mouse click event, for example you could bind right click, example [here](#example-custom-right-click-and-text-editor-functionality), and then identify where the user has clicked.

___

Determine if a tk `event.widget` is the `Sheet`.

```python
event_widget_is_sheet(
    event: Any,
    table: bool = True,
    index: bool = True,
    header: bool = True,
    top_left: bool = True,
) -> bool
```
Notes:

- Parameters set to `True` will include events that occurred within that widget.
    - e.g. If an event occurs in the top left corner of the sheet but the parameter `top_left` is `False` the function will return `False`.

___

Check if any Sheet widgets have focus.

```python
has_focus() -> bool:
```
- Includes child widgets such as scroll bars.

___

```python
identify_region(event: Any) -> Literal["table", "index", "header", "top left"]
```

___

```python
identify_row(
    event: Any,
    exclude_index: bool = False,
    allow_end: bool = True,
) -> int | None
```

___

```python
identify_column(
    event: Any,
    exclude_header: bool = False,
    allow_end: bool = True,
) -> int | None
```

___

#### **Sheet control actions**

For example: `sheet.bind("<Control-B>", sheet.paste)`

```python
cut(event: Any = None, validation: bool = True) -> None | EventDataDict
paste(event: Any = None, validation: bool = True) -> None | EventDataDict
delete(event: Any = None, validation: bool = True) -> None | EventDataDict
copy(event: Any = None) -> None | EventDataDict
undo(event: Any = None) -> None | EventDataDict
redo(event: Any = None) -> None | EventDataDict
```
- `validation` (`bool`) when `False` disables any bound `edit_validation()` function from running.

---
# **Scroll Positions and Cell Visibility**

#### **Sync scroll positions between widgets**

```python
sync_scroll(widget: Any) -> Sheet
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
unsync_scroll(widget: Any = None) -> Sheet
```
- Leaving `widget` as `None` unsyncs all previously synced widgets.

#### **See or scroll to a specific cell on the sheet**

```python
see(
    row: int = 0,
    column: int = 0,
    keep_yscroll: bool = False,
    keep_xscroll: bool = False,
    bottom_right_corner: bool | None = None,
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
set_xview(position: None | float = None, option: str = "moveto") -> Sheet | tuple[float, float]
```
Notes:
- If `position` is `None` then `tuple[float, float]` of main table `xview()` is returned.
- `xview` and `xview_moveto` have the same behaviour.

___

```python
set_yview(position: None | float = None, option: str = "moveto") -> Sheet | tuple[float, float]
```
- If `position` is `None` then `tuple[float, float]` of main table `yview()` is returned.
- `yview` and `yview_moveto` have the same behaviour.

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

**Get the bool**
```python
@property
all_columns()
```
- e.g. `get_all_columns_displayed = sheet.all_columns`.

**Set the bool**
```python
@all_columns.setter
all_columns(a: bool)
```
e.g. `sheet.all_columns = True`.

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
    data_indexes: bool = False,
) -> Sheet
```
Parameters:

- **NOTE**: `columns` (`int`) by default uses displayed column indexes, not data indexes. In other words the indexes of the columns displayed on the screen are the ones that are hidden, this is useful when used in conjunction with `get_selected_columns()`.
- `data_indexes` when `False` it makes the `columns` parameter indexes represent displayed columns and not the underlying Sheet data columns. When `True` the indexes represent data indexes.

Example:
```python
columns_to_hide = set(sheet.data_c(c) for c in sheet.get_selected_columns())
sheet.hide_columns(
    columns_to_hide,
    data_indexes=True,
)
```

___

#### **Unhide specific columns**

```python
show_columns(
    columns: int | Iterator[int],
    redraw: bool = True,
    deselect_all: bool = True,
) -> Sheet
```
Parameters:

- **NOTE**: `columns` (`int`) uses data column indexes, not displayed indexes. In other words the indexes of the columns which represent the underlying data are shown.

Notes:

- Will return if all columns are currently displayed (`Sheet.all_columns`).

Example:
```python
# converting displayed column indexes to data indexes using data_c(c)
columns = set(sheet.data_c(c) for c in sheet.get_selected_columns())

# hiding columns
sheet.hide_columns(
    columns,
    data_indexes=True,
)

# showing them again
sheet.show_columns(columns)
```

___

#### **Displayed column index to data**

Convert a displayed column index to a data index. If the internal `all_columns_displayed` attribute is `True` then it will return the provided argument.
```python
displayed_column_to_data(c)
data_c(c)
```

___

#### **Get currently displayed columns**

```python
@property
displayed_columns() -> list[int]
```
- e.g. `columns = sheet.displayed_columns`

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

- An example of row filtering using this function can be found [here](#example-header-dropdown-boxes-and-row-filtering).
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
    data_indexes: bool = False,
) -> Sheet
```
Parameters:

- **NOTE**: `rows` (`int`) by default uses displayed row indexes, not data indexes. In other words the indexes of the rows displayed on the screen are the ones that are hidden, this is useful when used in conjunction with `get_selected_rows()`.
- `data_indexes` when `False` it makes the `rows` parameter indexes represent displayed rows and not the underlying Sheet data rows. When `True` the indexes represent data indexes.

Example:
```python
rows_to_hide = set(sheet.data_r(r) for r in sheet.get_selected_rows())
sheet.hide_rows(
    rows_to_hide,
    data_indexes=True,
)
```

___

#### **Unhide specific rows**

```python
show_rows(
    rows: int | Iterator[int],
    redraw: bool = True,
    deselect_all: bool = True,
) -> Sheet
```
Parameters:

- **NOTE**: `rows` (`int`) uses data row indexes, not displayed indexes. In other words the indexes of the rows which represent the underlying data are shown.

Notes:

- Will return if all rows are currently displayed (`Sheet.all_rows`).

Example:
```python
# converting displayed row indexes to data indexes using data_r(r)
rows = set(sheet.data_r(r) for r in sheet.get_selected_rows())

# hiding rows
sheet.hide_rows(
    rows,
    data_indexes=True,
)

# showing them again
sheet.show_rows(rows)
```

___

#### **Get or set the all rows displayed boolean**

**Get the bool**
```python
@property
all_rows()
```
- e.g. `get_all_rows_displayed = sheet.all_rows`.

**Set the bool**
```python
@all_rows.setter
all_rows(a: bool)
```
e.g. `sheet.all_rows = True`.

```python
all_rows_displayed(a: bool | None = None) -> bool
```
- `a` (`bool`, `None`) Either set by using `bool` or get by leaving `None` e.g. `all_rows_displayed()`.

___

#### **Displayed row index to data**

Convert a displayed row index to a data index. If the internal `all_rows_displayed` attribute is `True` then it will return the provided argument.
```python
displayed_row_to_data(r)
data_r(r)
```

___

#### **Get currently displayed rows**

```python
@property
displayed_rows() -> list[int]
```
- e.g. `rows = sheet.displayed_rows`

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

#### **Set the cell text editor value if it is open**

**Table:**

```python
set_text_editor_value(
    text: str = "",
) -> Sheet
```

**Index:**

```python
set_index_text_editor_value(
    text: str = "",
) -> Sheet
```

**Header:**

```python
set_header_text_editor_value(
    text: str = "",
) -> Sheet
```

___

#### **Close any existing text editor**

```python
close_text_editor(set_data: bool = True) -> Sheet
```
Notes:

- Closes any open text editors, including header and index.
- Also closes any existing `"normal"` state dropdown box.

Parameters:

- `set_data` (`bool`) when `True` sets the cell data to the text editor value (if it is valid). When `False` the text editor is closed without setting data.

___

#### **Get an existing text editors value**

```python
get_text_editor_value() -> str | None
```
Notes:

- `None` is returned if no text editor exists, a `str` of the text editors value will be returned if it does.

___

#### **Hide all text editors**

```python
destroy_text_editor(event: Any = None) -> Sheet
```

___

#### **Get the table tk Text widget which acts as the text editor**

```python
get_text_editor_widget(event: Any = None) -> tk.Text | None
```

___

#### **Bind additional keys to the text editor window**

```python
bind_key_text_editor(key: str, function: Callable) -> Sheet
```

___

#### **Unbind additional text editors bindings set using the above function**

```python
unbind_key_text_editor(key: str) -> Sheet
```

---
# **Treeview Mode**

tksheet has a treeview mode which behaves similarly to the ttk treeview widget, it is not a drop in replacement for it though.

Always either use a fresh `Sheet()` instance or use [Sheet.reset()](#reset-all-or-specific-sheet-elements-and-attributes) before enabling treeview mode.

### **TO NOTE:**

- When treeview mode is enabled the row index is a `list` of `Node` objects. The row index should not be modified by the usual `row_index()` function.
- Most other tksheet functions should work as normal.
- The index text alignment must be `"w"` aka west or left.

## **Creating a treeview mode sheet**

You can make a treeview mode sheet by using the initialization parameter `treeview`:

```python
sheet = Sheet(parent, treeview=True)
```

Or by using [`Sheet.reset()`](#reset-all-or-specific-sheet-elements-and-attributes) and [`Sheet.set_options()`](#sheet-options).

```python
my_sheet.reset()
my_sheet.set_options(treeview=True)
```

See the other sections on sheet initialization and examples for the other usual `Sheet()` parameters.

## **Treeview mode functions**

Functions designed for use with treeview mode.

#### **Insert an item**

```python
insert(
    parent: str = "",
    index: None | int | Literal["end"] = None,
    iid: None | str = None,
    text: None | str = None,
    values: None | list[Any] = None,
    create_selections: bool = False,
    undo: bool = True,
) -> str
```
Parameters:

- `parent` is the `iid` of the parent item (if any). If left as `""` then the item will not have a parent.
- `index` is the row number for the item to be placed at, leave as `None` for the end.
- `iid` is a new and unique item id. It will be generated automatically if left as `None`.
- `text` is the displayed text in the row index for the item.
- `values` is a list of values which will become the items row in the sheet.
- `create_selections` when `True` selects the row that has just been created.
- `undo` when `True` adds the change to the undo stack.

Notes:

- Returns the `iid`.

Example:
```python
sheet.insert(
    iid="top level",
    text="Top level",
    values=["cell A1", "cell B1"],
)
sheet.insert(
    parent="top level",
    iid="mid level",
    text="Mid level",
    values=["cell A2", "cell B2"],
)
```

___

#### **Insert multiple items**

```python
bulk_insert(
    data: list[list[Any]],
    parent: str = "",
    index: None | int | Literal["end"] = None,
    iid_column: int | None = None,
    text_column: int | None | str = None,
    create_selections: bool = False,
    include_iid_column: bool = True,
    include_text_column: bool = True,
    undo: bool = True,
) -> dict[str, int]
```
Parameters:

- `parent` is the `iid` of the parent item (if any). If left as `""` then the items will not have a parent.
- `index` is the row number for the items to be placed at, leave as `None` for the end.
- `iid_column` if left as `None` iids will be automatically generated for the new items, else you can specify a column in the `data` which contains the iids.
- `text_column`:
    - If left as `None` there will be no displayed text next to the items.
    - A text column can be provided in the data and `text_column` set to an `int` representing its index to provide the displayed text for the items.
    - Or if a `str` is used all items will have that `str` as their displayed text.
- `create_selections` when `True` selects the row that has just been created.
- `include_iid_column` when `False` excludes the iid column from the inserted rows.
- `include_text_column` when the `text_column` is an `int` setting this to `False` excludes that column from the treeview.
- `undo` when `True` adds the change to the undo stack.

Notes:

- Returns a `dict[str, int]` of key: new iids, value: their data row number.

Example:
```python
sheet.insert(iid="ID-1", text="ID-1", values=["ID 1 Value"])
sheet.bulk_insert(
    data=[
        ["CID-1", "CID 1 Value"],
        ["CID-2", "CID 2 Value"],
    ],
    parent="ID-1",
    iid_column=0,
    text_column=0,
    include_iid_column=False,
    include_text_column=False,
)
```

___

#### **Build a tree from data**

This takes a list of lists where sublists are rows and a few arguments to bulk insert items into the treeview. **Note that**:

- **It resets the sheet** so cannot be used to bulk add to an already existing treeview.

```python
tree_build(
    data: list[list[Any]],
    iid_column: int,
    parent_column: int,
    text_column: None | int = None,
    push_ops: bool = False,
    row_heights: Sequence[int] | None | False = None,
    open_ids: Iterator[str] | None = None,
    safety: bool = True,
    ncols: int | None = None,
    lower: bool = False,
    include_iid_column: bool = True,
    include_parent_column: bool = True,
    include_text_column: bool = True,
) -> Sheet
```
Parameters:

- `data` a list of lists, one column must be an iid column, another must be a parent iid column.
- `text_column` if an `int` is used then the values in that column will populate the row index.
- `push_ops` when `True` the newly inserted rows will push all existing sheet options such as highlights downwards.
- `row_heights` a `list` of `int`s can be used to provide the displayed row heights in pixels (does not include hidden items). Only use if you know what you're doing here.
- `open_ids` a list of iids which will be opened.
- `safety` when `True` checks for infinite loops, empty iid cells and duplicate iids. No error or warning will be generated.
    - In the case of infinite loops the parent iid cell will be cleared.
    - In the case of empty iid cells the row will be ignored.
    - In the case of duplicate iids they will be renamed and `"DUPLICATED_<number>"` will be attached to the end.
- `ncols` is like maximum columns, an `int` which limits the number of columns that are included in the loaded data.
- `lower` makes all item ids - iids lower case.
- `include_iid_column` when `False` excludes the iid column from the inserted rows.
- `include_parent_column` when `False` excludes the parent column from the inserted rows.
- `include_text_column` when the `text_column` is an `int` setting this to `False` excludes that column from the treeview.

Notes:

- Returns the `Sheet` object.

Example:
```python
data = [
    ["id1", "", "id1 val"],
    ["id2", "id1", "id2 val"],
]
sheet.tree_build(
    data=data,
    iid_column=0,
    parent_column=1,
    include_iid_column=False,
    include_parent_column=False,
)
```

___

#### **Reset the treeview**

```python
tree_reset() -> Sheet
```

___

#### **Get treeview iids that are open**

```python
tree_get_open() -> set[str]
```

___

#### **Set the open treeview iids**

```python
tree_set_open(open_ids: Iterator[str]) -> Sheet
```
- Any other iids are closed as a result.

___

#### **Open treeview iids**

```python
tree_open(*items, redraw: bool = True) -> Sheet
```
- Opens all given iids.

___

#### **Close treeview iids**

```python
tree_close(*items, redraw: bool = True) -> Sheet
```
- Closes all given iids.

___

#### **Set or get an iids attributes**

```python
item(
    item: str,
    iid: str | None = None,
    text: str | None = None,
    values: list | None = None,
    open_: bool | None = None,
    undo: bool = True,
    emit_event: bool = True,
    redraw: bool = True,
) -> DotDict | Sheet
```
Parameters:

- `item` iid, required argument.
- `iid` use a `str` to rename the iid.
- `text` use a `str` to get the iid new display text in the row index.
- `values` use a `list` of values to give the item a new row of values (does not include row index).
- `open_` use a `bool` to set the item as open or closed. `False` is closed.
- `undo` adds any changes to the undo stack.
- `emit_event` emits a `<<SheetModified>>` event if changes were made.

Notes:
- If no arguments are given a `DotDict` is returned with the item attributes.

```python
{
    "text": ...,
    "values": ...,
    "open_": ...,
}
```

___

#### **Get an iids row number**

```python
itemrow(item: str) -> int
```
- Includes hidden rows in counting row numbers.

___

#### **Get a row numbers item**

```python
rowitem(row: int, data_index: bool = False) -> str | None
```
- Includes hidden rows in counting row numbers. See [here](#hiding-rows) for more information.

___

#### **Get treeview children**

```python
get_children(item: None | str = None) -> Generator[str]
```
- `item`:
    - When left as `None` will return all iids currently in treeview, including hidden rows.
    - Use an empty `str` (`""`) to get all top level iids in the treeview.
    - Use an iid to get the children for that particular iid. Does not include all descendants.

```python
tree_traverse(item: None | str = None) -> Generator[str]:
```
- Exactly the same as above but instead of retrieving iids from the index list it gets top iids and uses traversal to iterate through descendants.

___

#### **Get item descendants**

```python
descendants(item: str, check_open: bool = False) -> Generator[str]:
```
- Returns a generator which yields item ids in depth first order.

___

#### **Get the currently selected treeview item**

```python
@property
tree_selected() -> str | None:
```
- Returns the item id of the currently selected box row. If nothing is selected returns `None`.

___

#### **Delete treeview items**

```python
del_items(*items) -> Sheet
```
- `*items` the iids of items to delete.
- Also deletes all item descendants.

___

#### **Move items to a new parent**

```python
set_children(parent: str, *newchildren) -> Sheet
```
- `parent` the new parent for the items.
- `*newchildren` the items to move.

___

#### **Move an item to a new parent**

```python
move(
    item: str,
    parent: str,
    index: int | None = None,
    select: bool = True,
    undo: bool = True,
    emit_event: bool = False,
) -> tuple[dict[int, int], dict[int, int], EventDataDict]
```
Parameters:

- `item` is the iid to move.
- `parent` is the new parent for the item.
    - Use an empty `str` (`""`) to move the item to the top.
- `index`:
    - Leave as `None` to move to item to the end of the top/children.
    - Use an `int` to move the item to an index within its parents children (or within top level items if moving to the top).
- `select` when `True` selects the moved rows.
- `undo` when `True` adds the change to the undo stack.
- `emit_event` when `True` emits events related to sheet modification.

Notes:

- Also moves all of the items descendants.
- The `reattach()` function is exactly the same as `move()`.
- Returns `dict[int, int]` of `{old index: new index, ...}` for data and displayed rows separately and also an `EventDataDict`.

___

#### **Check an item exists**

```python
exists(item: str) -> bool
```
- `item` - a treeview iid.

___

#### **Get an items parent**

```python
parent(item: str) -> str
```
- `item` - a treeview iid.

___

#### **Get an items index**

```python
index(item: str) -> int
```
- `item` - a treeview iid.

___

#### **Check if an item is currently displayed**

```python
item_displayed(item: str) -> bool
```
- `item` - a treeview iid.

___

#### **Display an item**

Make sure an items parents are all open, does **not** scroll to the item.

```python
display_item(item: str, redraw: bool = False) -> Sheet
```
- `item` - a treeview iid.

___

#### **Scroll to an item**

- Make sure an items parents are all open and scrolls to the item.

```python
scroll_to_item(item: str, redraw: bool = False) -> Sheet
```
- `item` - a treeview iid.

___

#### **Get currently selected items**

```python
selection(cells: bool = False) -> list[str]
```
Notes:
- Returns a list of selected iids (selected rows but as iids).

Parameters:
- `cells` when `True` any selected cells will also qualify as selected items.

___

#### **Set selected items**

```python
selection_set(*items, run_binding: bool = True, redraw: bool = True) -> Sheet
```
- Sets selected rows (items).

___

#### **Add selected items**

```python
selection_add(*items, run_binding: bool = True, redraw: bool = True) -> Sheet
```

___

#### **Remove selections**

```python
selection_remove(*items, redraw: bool = True) -> Sheet
```

___

#### **Toggle selections**

```python
selection_toggle(*items, redraw: bool = True) -> Sheet
```

---
# **Progress Bars**

Progress bars can be created for individual cells. They will only update when tkinter updates.

#### **Create a progress bar**

```python
create_progress_bar(
    row: int,
    column: int,
    bg: str,
    fg: str,
    name: Hashable,
    percent: int = 0,
    del_when_done: bool = False,
) -> Sheet
```
- `row` the row coordinate to create the bar at.
- `column` the column coordinate to create the bar at.
- `bg` the background color for the bar.
- `fg` the text color for the bar.
- `name` a name is required for easy referral to the bar later on.
    - Names can be re-used for multiple bars.
- `percent` the starting progress of the bar as an `int` either `0`, `100` or a number in between.
- `del_when_done` if `True` the `Sheet` will automatically delete the progress bar once it is modified with a percent of `100` or more.

___

#### **Modify progress bars**

```python
progress_bar(
    name: Hashable | None = None,
    cell: tuple[int, int] | None = None,
    percent: int | None = None,
    bg: str | None = None,
    fg: str | None = None,
) -> Sheet
```
Either `name` or `cell` can be used to refer to existing progress bars:

- `name` the name given to a progress bar, or multiple progress bars.
    - If this parameter is used then `cell` will not be used.
    - Will modify all progress bars with the given name.
- `cell` (`tuple[int, int]`) a tuple of two `int`s representing the progress bars location, `(row, column)`.
    - Can only refer to one progress bar.

Values that can be modified:

- `bg` the background color for the bar, leave as `None` for no change.
- `fg` the text color for the bar, leave as `None` for no change.
- `percent` the progress of the bar as an `int` either `0`, `100` or a number in between, leave as `None` for no change.

___

#### **Delete progress bars**

Note that this will delete the progress bars data from the Sheet as well.

```python
del_progress_bar(
    name: Hashable | None = None,
    cell: tuple[int, int] | None = None,
) -> Sheet
```
Either `name` or `cell` can be used to refer to existing progress bars:

- `name` the name given to a progress bar, or multiple progress bars.
    - Will delete all progress bars with the given name.
    - If this parameter is used then `cell` will not be used.
- `cell` (`tuple[int, int]`) a tuple of two `int`s representing the progress bars location, `(row, column)`.
    - Can only refer to one progress bar.

---
# **Tags**

Tags can be used to keep track of specific cells, rows and columns wherever they move. Note that:

- If rows/columns are deleted the the associated tags will be also.
- There is no equivalent `tag_bind` functionality at this time.
- All tagging functions use data indexes (not displayed indexes) - this is only relevant when there are hidden rows/columns.

#### **Tag a specific cell**

```python
tag_cell(
    cell: tuple[int, int],
    *tags,
) -> Sheet
```
Example:
```python
sheet.tag_cell((0, 0), "tag a1", "tag a1 no.2")
```

___

#### **Tag specific rows**

```python
tag_rows(
    rows: int | Iterator[int],
    *tags,
) -> Sheet
```

___

#### **Tag specific columns**

```python
tag_columns(
    columns: int | Iterator[int],
    *tags,
) -> Sheet
```

___

#### **Tag using a span**

```python
tag(
    *key: CreateSpanTypes,
    tags: Iterator[str] | str = "",
) -> Sheet
```

___

#### **Untag**

```python
untag(
    cell: tuple[int, int] | None = None,
    rows: int | Iterator[int] | None = None,
    columns: int | Iterator[int] | None = None,
) -> Sheet
```
- This removes all tags from the cell, rows or columns provided.

___

#### **Delete tags**

```python
tag_del(
    *tags,
    cells: bool = True,
    rows: bool = True,
    columns: bool = True,
) -> Sheet
```
- This deletes the provided tags from all `cells` if `True`, `rows` if `True` and `columns` if `True`.

___

#### **Get cells, rows or columns associated with tags**

```python
tag_has(
    *tags,
) -> DotDict
```
Notes:

- Returns all cells, rows and columns associated with **any** of the provided tags in the form of a `dict` with dot notation accessbility which has the following keys:
    - `"cells"` - with a value of `set[tuple[int, int]]` where the `tuple`s are cell coordinates - `(row, column)`.
    - `"rows"` - with a value of `set[int]` where the `int`s are rows.
    - `"columns"` - with a value of `set[int]` where the `int`s are columns.
- Returns data indexes.
- This function **updates** the `set`s with any cells/rows/columns associated with each tag, it does **not** return cells/rows/columns that have all the provided tags.

Example:
```python
sheet.tag_rows((0, 1), "row tag a", "row tag b")
sheet.tag_rows(4, "row tag b")
sheet.tag_rows(5, "row tag c")
sheet.tag_rows(6, "row tag d")
with_tags = sheet.tag_has("row tag b", "row tag c")

print (with_tags.rows)
# prints {0, 1, 4, 5}
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
            empty_horizontal=0,
            empty_vertical=0,
            paste_can_expand_x=True,
            paste_can_expand_y=True,
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
            theme="dark",
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
        self.sheet.align_columns(columns=2, align="c")
        self.sheet.align_columns(columns=3, align="e")
        self.sheet.create_index_dropdown(r=0, values=["Dropdown"] + [f"{i}" for i in range(15)])
        self.sheet.create_index_checkbox(r=3, checked=True, text="Checkbox")
        self.sheet.create_dropdown(r="all", c=0, values=["Dropdown"] + [f"{i}" for i in range(15)])
        self.sheet.create_checkbox(r="all", c=1, checked=True, text="Checkbox")
        self.sheet.create_header_dropdown(c=0, values=["Header Dropdown"] + [f"{i}" for i in range(15)])
        self.sheet.create_header_checkbox(c=1, checked=True, text="Header Checkbox")
        self.sheet.align_cells(5, 0, align="c")
        self.sheet.highlight_cells(5, 0, bg="gray50", fg="blue")
        self.sheet.highlight_cells(17, canvas="index", bg="yellow", fg="black")
        self.sheet.highlight_cells(12, 1, bg="gray90", fg="purple")
        for r in range(len(colors)):
            self.sheet.highlight_cells(row=r, column=3, fg=colors[r])
            self.sheet.highlight_cells(row=r, column=4, bg=colors[r], fg="black")
            self.sheet.highlight_cells(row=r, column=5, bg=colors[r], fg="purple")
        self.sheet.highlight_cells(column=5, canvas="header", bg="white", fg="purple")
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

date_replace = re.compile("|".join(re.escape(char) for char in "()[]<>"))


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

### **Contributors and Special Thanks**

A special thank you to:

- @CalJaDav for the very helpful ideas/pull requests, guidance in implementing them and helping me become a better developer.
- @demberto for providing pull requests and guidance to modernize and improve the project.
- All [contributors](https://github.com/ragardner/tksheet/graphs/contributors).
- Everyone who has reported an issue and helped me fix it.
