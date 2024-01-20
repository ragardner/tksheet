# **tksheet** [![PyPI version shields.io](https://img.shields.io/pypi/v/tksheet.svg)](https://pypi.python.org/pypi/tksheet/) ![python](https://img.shields.io/badge/python-3.8+-blue) [![License: MIT](https://img.shields.io/badge/License-MIT%20-blue.svg)](https://github.com/ragardner/tksheet/blob/master/LICENSE.txt) [![GitHub Release Date](https://img.shields.io/github/release-date-pre/ragardner/tksheet.svg)](https://github.com/ragardner/tksheet/releases) [![Downloads](https://img.shields.io/pypi/dm/tksheet.svg)](https://pypi.org/project/tksheet/)

**A python tkinter table widget**

## **Help**

- [Installation](https://github.com/ragardner/tksheet/wiki/Version-7#installation-and-requirements)
- [Version 7+ Documentation](https://github.com/ragardner/tksheet/wiki/Version-7)
- [Version 6 Documentation](https://github.com/ragardner/tksheet/wiki/Version-6)
- [Changelog](https://github.com/ragardner/tksheet/blob/master/CHANGELOG.md)
- [Questions](https://github.com/ragardner/tksheet/wiki/Version-7#asking-questions)
- [Issues](https://github.com/ragardner/tksheet/wiki/Version-7#issues)
- [Suggestions](https://github.com/ragardner/tksheet/wiki/Version-7#enhancements-or-suggestions)
- This library is maintained with the help of **[others](https://github.com/ragardner/tksheet/graphs/contributors)**. If you would like to contribute please read [this help section](https://github.com/ragardner/tksheet/wiki/Version-7#contributing).

## **Changes for versions >= `7.0.0`:**

- **ALL** `extra_bindings()` event objects have changed, information [here](https://github.com/ragardner/tksheet/wiki/Version-7#bind-specific-table-functionality).
- The bound function for `extra_bindings()` with `"edit_cell"`/`"end_edit_cell"` **no longer** requires a return value and no longer sets the cell to the return value. Use [this](https://github.com/ragardner/tksheet/wiki/Version-7#validate-user-cell-edits) instead.
- `edit_cell_validation` has been removed and replaced with the function `edit_validation()`, information [here](https://github.com/ragardner/tksheet/wiki/Version-7#validate-user-cell-edits).
- Only Python versions >= `3.8` are supported.
- `tksheet` file names have been changed.
- Many other smaller changes, see the [changelog](https://github.com/ragardner/tksheet/blob/master/CHANGELOG.md) for more information.

## **Features**

- Display and modify tabular data
- Stores its display data as a Python list of lists, sublists being rows
- Runs smoothly even with millions of rows/columns
- Edit cells directly
- Cell values can potentially be [any class](https://github.com/ragardner/tksheet/wiki/Version-7#data-formatting), the default is any class with a `__str__` method
- Drag and drop columns and rows
- Multiple line header and index cells
- Expand row heights and column widths
- Change fonts and font size (not for individual cells)
- Change any colors in the sheet
- Create an unlimited number of high performance dropdown and check boxes
- [Hide rows and/or columns](https://github.com/ragardner/tksheet/wiki/Version-7#example-header-dropdown-boxes-and-row-filtering)
- Left `"w"`, Center `"center"` or Right `"e"` text alignment for any cell/row/column
- Versions >= 7 have succinct and easy to read syntax:

```python
# set data
sheet["A1"].data = "edited cell A1"
# get data
column_b = sheet["B"].data
```

### **light blue theme**

![tksheet light blue theme](https://i.imgur.com/ojU3IQi.jpeg)

### **black theme**

![tksheet black theme](https://i.imgur.com/JeF9vJe.jpeg)
