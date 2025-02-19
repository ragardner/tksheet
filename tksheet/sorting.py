from __future__ import annotations

from collections.abc import Callable, Generator
from datetime import datetime
from re import finditer

from .other_classes import Node
from .tksheet_types import AnyIter

# Possible date formats to try for the entire string
date_formats = [
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
    "%m %d %Y",  # Month Day Year
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
]


def natural_sort_key(item: object) -> tuple[int, object]:
    """
    A key function for natural sorting that handles various Python types, including
    date-like strings in multiple formats.

    This function aims to sort elements in a human-readable order:
    - None values first
    - Booleans (False before True)
    - Numbers (integers, floats combined)
    - Datetime objects
    - Strings with natural sorting for embedded numbers and dates
    - Unknown types treated as strings or left at the end

    Args:
        item: Any Python object to be sorted.

    Returns:
        A tuple or value that can be used for sorting.
    """
    if item is None:
        return (0,)

    elif isinstance(item, bool):
        return (1, item)

    elif isinstance(item, (int, float)):
        return (2, item)

    elif isinstance(item, datetime):
        return (3, item.timestamp())

    elif isinstance(item, str):
        for date_format in date_formats:
            try:
                return (3, datetime.strptime(item, date_format).timestamp())
            except ValueError:
                continue

        try:
            return (2, float(item))
        except Exception:
            n = []
            s = []
            for match in finditer(r"\d+|[^\d\s]+", item):
                if (m := match.group()).isdigit():
                    n.append(int(m))
                else:
                    s.append(m.lower())
            return (5, s, n)

    else:
        # For unknown types, attempt to convert to string, or place at end
        try:
            return (6, f"{item}".lower())
        except Exception:
            return (7, item)


def sort_selection(
    data: list[list[object]],
    reverse: bool = False,
    key: Callable | None = None,
    row_wise: bool = False,
) -> list[list[object]]:
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
    data: list[list[object]] | list[object] | AnyIter[object],
    column: int = 0,
    reverse: bool = False,
    key: Callable | None = None,
) -> list[list[object]] | list[object]:
    if not data:
        return data

    if key is None:
        key = natural_sort_key

    if isinstance(data, list) and isinstance(data[0], list):
        return sorted(data, key=lambda row: key(row[column]) if len(row) > column else key(None), reverse=reverse)
    else:
        return sorted(data, reverse=reverse, key=key)


def sort_row(
    data: list[list[object]] | list[object] | AnyIter[object],
    row: int = 0,
    reverse: bool = False,
    key: Callable | None = None,
) -> list[list[object]]:
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
    data: list[list[object]],
    column: int = 0,
    reverse: bool = False,
    key: Callable | None = None,
) -> tuple[list[tuple[int, list[object]]], dict[int, int]]:
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
    data: list[list[object]],
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

    return sort_indices, {old: new for old, new in zip(range(len(data[row])), sort_indices)}


def _sort_node_children(
    node: Node,
    tree: dict[str, Node],
    reverse: bool,
    key: Callable,
) -> Generator[Node, None, None]:
    sorted_children = sorted(
        (tree[child_iid] for child_iid in node.children if child_iid in tree),
        key=lambda child: key(child.text),
        reverse=reverse,
    )
    for child in sorted_children:
        yield child
        if child.children:  # If the child node has children
            yield from _sort_node_children(child, tree, reverse, key)


def sort_tree_view(
    _row_index: list[Node],
    tree_rns: dict[str, int],
    tree: dict[str, Node],
    key: Callable = natural_sort_key,
    reverse: bool = False,
) -> tuple[list[Node], dict[int, int]]:
    if not _row_index or not tree_rns or not tree:
        return [], {}

    if key is None:
        key = natural_sort_key

    # Create the index map and sorted nodes list
    mapping = {}
    sorted_nodes = []
    new_index = 0

    # Sort top-level nodes
    for node in sorted(
        (node for node in _row_index if node.parent == ""),
        key=lambda node: key(node.text),
        reverse=reverse,
    ):
        iid = node.iid
        mapping[tree_rns[iid]] = new_index
        sorted_nodes.append(node)
        new_index += 1

        # Sort children recursively
        for descendant_node in _sort_node_children(node, tree, reverse, key):
            mapping[tree_rns[descendant_node.iid]] = new_index
            sorted_nodes.append(descendant_node)
            new_index += 1

    return sorted_nodes, mapping
