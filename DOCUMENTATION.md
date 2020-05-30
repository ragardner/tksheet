#### Sheet startup arguments
This is a full list of all the start up arguments, the only required argument is the sheets parent, everything else has default arguments

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
default_header = "letters", #letters, numbers or both
default_row_index = "numbers", #letters, numbers or both
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

You can change these settings after initialization using the `set_options()` function or other functions contained within `_tksheet.py`

Take a look inside the `tests` folder for demonstrations

Documentation not finished...





