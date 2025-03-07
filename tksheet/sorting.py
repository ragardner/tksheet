from __future__ import annotations

from collections.abc import Callable, Iterator
from datetime import datetime
from pathlib import Path
from re import split
from typing import Any

# Possible date formats to try for the entire string
date_formats = (
    # Common formats
    "%d/%m/%Y",  # Day/Month/Year
    "%m/%d/%Y",  # Month/Day/Year (US format)
    "%Y/%m/%d",  # Year/Month/Day
    "%d.%m.%Y",  # Day.Month.Year (European format)
    "%d-%m-%Y",  # Day-Month-Year
    "%m-%d-%Y",  # Month-Day-Year
    "%Y-%m-%d",  # Year-Month-Day (ISO format without time)
    "%d/%m/%y",  # Day/Month/2-digit year
    "%m/%d/%y",  # Month/Day/2-digit year
    "%y/%m/%d",  # 2-digit year/Month/Day
    "%d,%m,%Y",  # Day,Month,Year
    "%m,%d,%Y",  # Month,Day,Year
    "%Y,%m,%d",  # Year,Month,Day
    "%d %m %Y",  # Day Month Year (with space)
    "%m %d %Y",  # Month Day Year (with space)
    "%Y %d %m",  # Year Month Day (with space)
    # With month names
    "%d %b %Y",  # Day Abbreviated Month Year
    "%b %d, %Y",  # Abbreviated Month Day, Year
    "%d %B %Y",  # Day Full Month Name Year
    "%B %d, %Y",  # Full Month Name Day, Year
    # ISO 8601 with/without time
    "%Y-%m-%dT%H:%M:%S",  # With time
    "%Y-%m-%d",  # Without time
    # Regional or less common formats
    # "%Y年%m月%d日",  # Japanese-style date
    "%Y%m%d",  # YYYYMMDD format, often used in logs or filenames
    "%y%m%d",  # YYMMDD
    "%d%m%Y",  # DDMMYYYY, sometimes used in Europe
    # Additional formats
    "%d/%m/%y %H:%M",  # Day/Month/Year Hour:Minute
    "%m/%d/%y %H:%M",  # Month/Day/Year Hour:Minute
    "%Y-%m-%d %H:%M:%S",  # Year-Month-Day Hour:Minute:Second
)


def _string_fallback(item: str) -> tuple[int, ...]:
    """
    In order to have reasonable file path sorting:
    - Split by path separators
    - Determine depth, more separators more depth
    - Deal with every dir by splitting by digits
    - Deal with file name by splitting by digits
    """
    components = split(r"[/\\]", item)
    if components[-1]:
        return (
            5,
            len(components),
            tuple(int(e) if e.isdigit() else e.lower() for comp in components[:-1] for e in split(r"(\d+)", comp) if e),
            tuple(int(e) if e.isdigit() else e.lower() for e in split(r"(\d+)", components[-1])),
        )
    else:
        return (
            5,
            len(components),
            tuple(int(e) if e.isdigit() else e.lower() for comp in components[:-1] for e in split(r"(\d+)", comp) if e),
            (),
        )


def version_sort_key(item: Any) -> tuple[int, ...]:
    """
    A key for natural sorting of various Python types.

    - Won't convert strings to floats
    - Will sort string version numbers

    0. None
    1. Empty strings
    2. bool
    3. int, float
    4. datetime (inc. strings that are dates)
    5. strings (including string file paths and paths as POSIX strings) & unknown objects with __str__
    6. unknown objects
    """
    if isinstance(item, str):
        if item:
            for date_format in date_formats:
                try:
                    return (4, datetime.strptime(item, date_format).timestamp())
                except ValueError:
                    continue
            # the same as _string_fallback
            components = split(r"[/\\]", item)
            if components[-1]:
                return (
                    5,
                    len(components),
                    tuple(
                        int(e) if e.isdigit() else e.lower()
                        for comp in components[:-1]
                        for e in split(r"(\d+)", comp)
                        if e
                    ),
                    tuple(int(e) if e.isdigit() else e.lower() for e in split(r"(\d+)", components[-1])),
                )
            else:
                return (
                    5,
                    len(components),
                    tuple(
                        int(e) if e.isdigit() else e.lower()
                        for comp in components[:-1]
                        for e in split(r"(\d+)", comp)
                        if e
                    ),
                    (),
                )
        else:
            return (1, item)

    elif item is None:
        return (0,)

    elif isinstance(item, bool):
        return (2, item)

    elif isinstance(item, (int, float)):
        return (3, item)

    elif isinstance(item, datetime):
        return (4, item.timestamp())

    elif isinstance(item, Path):
        return _string_fallback(item.as_posix())

    else:
        try:
            return _string_fallback(f"{item}")
        except Exception:
            return (6, item)


def natural_sort_key(item: Any) -> tuple[int, ...]:
    """
    A key for natural sorting of various Python types.

    - Won't sort string version numbers
    - Will convert strings to floats
    - Will sort strings that are file paths

    0. None
    1. Empty strings
    2. bool
    3. int, float (inc. strings that are numbers)
    4. datetime (inc. strings that are dates)
    5. strings (including string file paths and paths as POSIX strings) & unknown objects with __str__
    6. unknown objects
    """
    if isinstance(item, str):
        if item:
            for date_format in date_formats:
                try:
                    return (4, datetime.strptime(item, date_format).timestamp())
                except ValueError:
                    continue

            try:
                return (3, float(item))
            except ValueError:
                # the same as _string_fallback
                components = split(r"[/\\]", item)
                if components[-1]:
                    return (
                        5,
                        len(components),
                        tuple(
                            int(e) if e.isdigit() else e.lower()
                            for comp in components[:-1]
                            for e in split(r"(\d+)", comp)
                            if e
                        ),
                        tuple(int(e) if e.isdigit() else e.lower() for e in split(r"(\d+)", components[-1])),
                    )
                else:
                    return (
                        5,
                        len(components),
                        tuple(
                            int(e) if e.isdigit() else e.lower()
                            for comp in components[:-1]
                            for e in split(r"(\d+)", comp)
                            if e
                        ),
                        (),
                    )
        else:
            return (1, item)

    elif item is None:
        return (0,)

    elif isinstance(item, bool):
        return (2, item)

    elif isinstance(item, (int, float)):
        return (3, item)

    elif isinstance(item, datetime):
        return (4, item.timestamp())

    elif isinstance(item, Path):
        return _string_fallback(item.as_posix())

    else:
        try:
            return _string_fallback(f"{item}")
        except Exception:
            return (6, item)


def fast_sort_key(item: Any) -> tuple[int, ...]:
    """
    A faster key for natural sorting of various Python types.

    - Won't sort strings that are dates very well
    - Won't convert strings to floats
    - Won't sort string file paths very well
    - Will do ok with string version numbers

    0. None
    1. Empty strings
    2. bool
    3. int, float
    4. datetime
    5. strings (including paths as POSIX strings) & unknown objects with __str__
    6. unknown objects
    """
    if isinstance(item, str):
        if item:
            return (5, tuple(int(e) if e.isdigit() else e.lower() for e in split(r"(\d+)", item)))
        else:
            return (1, item)

    elif item is None:
        return (0,)

    elif isinstance(item, bool):
        return (2, item)

    elif isinstance(item, (int, float)):
        return (3, item)

    elif isinstance(item, datetime):
        return (4, item.timestamp())

    elif isinstance(item, Path):
        return _string_fallback(item.as_posix())

    else:
        try:
            return _string_fallback(f"{item}")
        except Exception:
            return (6, item)


def sort_selection(
    data: list[list[Any]],
    reverse: bool = False,
    key: Callable | None = None,
    row_wise: bool = False,
) -> list[list[Any]]:
    if not data or not isinstance(data[0], list):
        raise ValueError("Data must be a list of lists.")

    if key is None:
        key = natural_sort_key

    if row_wise:
        return [sorted(row, key=key, reverse=reverse) for row in data]
    else:
        return list(
            zip(
                *(
                    sorted(
                        (row[col] for row in data),
                        key=key,
                        reverse=reverse,
                    )
                    for col in range(len(data[0]))
                )
            )
        )


def sort_column(
    data: list[list[Any]] | list[Any] | Iterator[Any],
    column: int = 0,
    reverse: bool = False,
    key: Callable | None = None,
) -> list[list[Any]] | list[Any]:
    if not data:
        return data

    if key is None:
        key = natural_sort_key

    if isinstance(data, list) and isinstance(data[0], list):
        return sorted(data, key=lambda row: key(row[column]) if len(row) > column else key(None), reverse=reverse)
    else:
        return sorted(data, reverse=reverse, key=key)


def sort_row(
    data: list[list[Any]] | list[Any] | Iterator[Any],
    row: int = 0,
    reverse: bool = False,
    key: Callable | None = None,
) -> list[list[Any]]:
    if not data:
        return data

    if key is None:
        key = natural_sort_key

    if isinstance(data, list) and isinstance(data[0], list):
        if 0 <= row < len(data):
            data[row] = sorted(data[row], key=key, reverse=reverse)
            return data
        else:
            raise IndexError(f"Row index {row} out of range for data with {len(data)} rows.")
    else:
        return sorted(data, reverse=reverse, key=key)


def sort_rows_by_column(
    data: list[list[Any]],
    column: int = 0,
    reverse: bool = False,
    key: Callable | None = None,
) -> tuple[list[tuple[int, list[Any]]], dict[int, int]]:
    if not data:
        return data, {}

    # Check if data is a list of lists
    if not isinstance(data[0], list):
        raise ValueError("Data must be a list of lists for row sorting.")

    if key is None:
        key = natural_sort_key

    # Use a generator expression for sorting to avoid creating an intermediate list
    sorted_indexed_data = sorted(
        ((i, row) for i, row in enumerate(data)),
        key=lambda item: key(item[1][column]) if len(item[1]) > column else key(None),
        reverse=reverse,
    )

    # Return sorted rows [(old index, row), ...] and create the mapping dictionary
    return sorted_indexed_data, {old: new for new, (old, _) in enumerate(sorted_indexed_data)}


def sort_columns_by_row(
    data: list[list[Any]],
    row: int = 0,
    reverse: bool = False,
    key: Callable | None = None,
) -> tuple[list[int], dict[int, int]]:
    if not data:
        return data, {}

    # Check if data is a list of lists
    if not isinstance(data[0], list):
        raise ValueError("Data must be a list of lists for column sorting.")

    if row >= len(data) or row < 0:
        raise IndexError(f"Row index {row} out of range for data with {len(data)} rows.")

    if key is None:
        key = natural_sort_key

    # Get sorting indices based on the elements of the specified row
    sort_indices = sorted(range(len(data[row])), key=lambda i: key(data[row][i]), reverse=reverse)
    sort_indices_set = set(sort_indices)

    new_data = []
    for row_data in data:
        # Sort the columns based on sort_indices, then append any additional elements from longer rows
        sorted_part = [row_data[i] for i in sort_indices if i < len(row_data)]
        unsorted_part = [elem for idx, elem in enumerate(row_data) if idx not in sort_indices_set]
        new_data.append(sorted_part + unsorted_part)

    return sort_indices, dict(zip(range(len(data[row])), sort_indices))


def sort_tree_rows_by_column(
    data: list[list[Any]],
    column: int,
    index: list[Any],
    rns: dict[str, int],
    reverse: bool = False,
    key: Callable | None = None,
) -> tuple[list[Any], dict[int, int]]:
    """
    Sorts tree rows by a specified column in depth-first order, returning sorted nodes and a row mapping.

    Args:
        data: List of rows, where each row is a list of column values.
        column: Index of the column to sort by.
        index: List of nodes, where each node has 'iid', 'parent', and 'children' attributes.
        rns: Dictionary mapping item IDs (iid) to original row numbers in data.
        reverse: If True, sort in descending order; otherwise, ascending.
        key: Optional function to compute sort keys; defaults to natural_sort_key if None.

    Returns:
        Tuple containing:
        - List of nodes in sorted order.
        - Dictionary mapping original row numbers to new row numbers.
    """
    if not index or not rns:
        return [], {}

    if key is None:
        key = natural_sort_key  # Assuming natural_sort_key is defined elsewhere

    # Define the sort_reverse parameter to avoid unnecessary reversals
    sort_reverse = not reverse

    # Helper function to compute the sorting key for a node
    def get_key(node):
        row = data[rns[node.iid]]
        return key(row[column] if len(row) > column else None)

    # Initialize stack with sorted top-level nodes for efficiency
    stack = sorted(
        (node for node in index if node.parent == ""),
        key=get_key,
        reverse=sort_reverse,
    )

    # Initialize output structures
    sorted_nodes = []
    mapping = {}
    new_rn = 0

    # Process nodes iteratively in depth-first order
    while stack:
        current = stack.pop()  # Pop from the right end
        sorted_nodes.append(current)
        mapping[rns[current.iid]] = new_rn
        new_rn += 1
        if current.children:
            # Sort children
            sorted_children = sorted(
                (index[rns[ciid]] for ciid in current.children if ciid in rns),
                key=get_key,
                reverse=sort_reverse,
            )
            # Extend stack with sorted children
            stack.extend(sorted_children)  # Adds to the right end

    return sorted_nodes, mapping


# if __name__ == "__main__":
#     test_cases = {
#         "Filenames": ["file10.txt", "file2.txt", "file1.txt"],
#         "Versions": ["1.10", "1.2", "1.9", "1.21"],
#         "Mixed": ["z1.doc", "z10.doc", "z2.doc", "z100.doc"],
#         "Paths": [
#             "/folder/file.txt",
#             "/folder/file (1).txt",
#             "/folder (1)/file.txt",
#             "/folder (10)/file.txt",
#             "/dir/subdir/file1.txt",
#             "/dir/subdir/file10.txt",
#             "/dir/subdir/file2.txt",
#             "/dir/file.txt",
#             "/dir/sub/file123.txt",
#             "/dir/sub/file12.txt",
#             # New challenging cases
#             "/a/file.txt",
#             "/a/file1.txt",
#             "/a/b/file.txt",
#             "/a/b/file1.txt",
#             "/x/file-v1.2.txt",
#             "/x/file-v1.10.txt",
#             # My own new challenging cases
#             "/a/zzzzz.txt",
#             "/a/b/a.txt",
#         ],
#         "Case": ["Apple", "banana", "Corn", "apple", "Banana", "corn"],
#         "Leading Zeros": ["001", "010", "009", "100"],
#         "Complex": ["x-1-y-10", "x-1-y-2", "x-2-y-1", "x-10-y-1"],
#         "Lengths": ["2 ft 9 in", "2 ft 10 in", "10 ft 1 in", "10 ft 2 in"],
#         "Floats": [
#             "1.5",
#             "1.25",
#             "1.025",
#             "10.5",
#             "-10.2",
#             "-2.5",
#             "5.7",
#             "-1.25",
#             "0.0",
#             "1.5e3",
#             "2.5e2",
#             "1.23e4",
#             "1e-2",
#             "file1.2.txt",
#             "file1.5.txt",
#             "file1.10.txt",
#         ],
#         "Strings": [
#             "123abc",
#             "abc123",
#             "123abc456",
#             "abc123def",
#         ],
#     }

#     print("FAST SORT KEY:")

#     for name, data in test_cases.items():
#         sorted1 = sorted(data, key=fast_sort_key)
#         print(f"\n{name}:")
#         print(sorted1)

#     print("\nNATURAL SORT KEY:")

#     for name, data in test_cases.items():
#         sorted1 = sorted(data, key=natural_sort_key)
#         print(f"\n{name}:")
#         print(sorted1)

#     print("\nVERSION SORT KEY:")

#     for name, data in test_cases.items():
#         sorted1 = sorted(data, key=version_sort_key)
#         print(f"\n{name}:")
#         print(sorted1)
