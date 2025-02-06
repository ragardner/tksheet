from __future__ import annotations

import re
import unittest
from collections.abc import Callable
from datetime import datetime

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
    "%Y年%m月%d日",  # Japanese-style date
    "%Y%m%d",  # YYYYMMDD format, often used in logs or filenames
    "%y%m%d",  # YYMMDD
    "%d%m%Y",  # DDMMYYYY, sometimes used in Europe
    # Additional formats
    "%d/%m/%y %H:%M",  # Day/Month/Year Hour:Minute
    "%m/%d/%y %H:%M",  # Month/Day/Year Hour:Minute
    "%Y-%m-%d %H:%M:%S",  # Year-Month-Day Hour:Minute:Second
]


def natural_sort_key(item: object) -> tuple[int, str | float | tuple[int | str, ...] | object]:
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

    With love from Grok ❤️

    Args:
        item: Any Python object to be sorted.

    Returns:
        A tuple or value that can be used for sorting.
    """
    if item is None:
        return (0, "")

    elif isinstance(item, bool):
        return (1, item)

    elif isinstance(item, (int, float)):
        return (2, (item,))  # Tuple to ensure float and int are sorted together

    elif isinstance(item, datetime):
        return (3, item.timestamp())

    elif isinstance(item, str):
        # Check if the whole string is a date
        for date_format in date_formats:
            try:
                # Use the same sort order as for datetime objects
                return (3, datetime.strptime(item, date_format).timestamp())
            except ValueError:
                continue

        # Check if the whole string is a number
        try:
            return (4, int(item))
        except Exception:
            pass
        try:
            return (4, float(item))
        except Exception:
            pass

        # Proceed with natural sorting
        return (4, [int(text) if text.isdigit() else text.lower() for text in re.split(r"(\d+)", item)])

    else:
        # For unknown types, attempt to convert to string, or place at end
        try:
            return (5, f"{item}".lower())
        except Exception:
            return (6, item)  # If conversion fails, place at the very end


def sort_column(
    data: list[list[object]] | list[object],
    column: int = 0,
    reverse: bool = False,
    key: Callable = natural_sort_key,
) -> list[list[object]] | list[object]:
    if not data:
        return data

    if isinstance(data[0], list):
        return sorted(data, key=lambda row: key(row[column]) if len(row) > column else key(None), reverse=reverse)
    else:
        return sorted(data, reverse=reverse, key=key)


def sort_row(
    data: list[list[object]] | list[object],
    row: int = 0,
    reverse: bool = False,
    key: Callable = natural_sort_key,
) -> list[list[object]]:
    if not data:
        return data

    if isinstance(data[0], list):
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
    key: Callable = natural_sort_key,
) -> list[list[object]]:
    if not data:
        return data

    # Check if data is a list of lists
    if not isinstance(data[0], list):
        raise ValueError("Data must be a list of lists for row sorting.")

    # Sort the data using the column as the key for sorting
    return sorted(data, key=lambda row: key(row[column]) if len(row) > column else key(None), reverse=reverse)


def sort_columns_by_row(
    data: list[list[object]],
    row: int = 0,
    reverse: bool = False,
    key: Callable = natural_sort_key,
) -> list[list[object]]:
    if not data:
        return data

    # Check if data is a list of lists
    if not isinstance(data[0], list):
        raise ValueError("Data must be a list of lists for column sorting.")

    if row >= len(data) or row < 0:
        raise IndexError(f"Row index {row} out of range for data with {len(data)} rows.")

    # Get sorting indices based on the elements of the specified row
    sort_indices = sorted(range(len(data[row])), key=lambda i: key(data[row][i]), reverse=reverse)
    sort_indices_set = set(sort_indices)

    new_data = []
    for row_data in data:
        # Sort the columns based on sort_indices, then append any additional elements from longer rows
        sorted_part = [row_data[i] for i in sort_indices if i < len(row_data)]
        unsorted_part = [elem for idx, elem in enumerate(row_data) if idx not in sort_indices_set]
        new_data.append(sorted_part + unsorted_part)

    return new_data


class TestNaturalSort(unittest.TestCase):
    def test_none_first(self):
        self.assertEqual(natural_sort_key(None), (0, ""))

    def test_booleans_order(self):
        self.assertLess(natural_sort_key(False), natural_sort_key(True))

    def test_numbers_order(self):
        self.assertLess(natural_sort_key(5), natural_sort_key(10))
        self.assertLess(natural_sort_key(5.5), natural_sort_key(6))

    def test_datetime_order(self):
        dt1 = datetime(2023, 1, 1)
        dt2 = datetime(2023, 1, 2)
        self.assertLess(natural_sort_key(dt1), natural_sort_key(dt2))

    def test_string_natural_sort(self):
        items = ["item2", "item10", "item1"]
        sorted_items = sorted(items, key=natural_sort_key)
        self.assertEqual(sorted_items, ["item1", "item2", "item10"])

    def test_date_string_recognition(self):
        # Test various date formats
        date_str1 = "01/01/2023"
        date_str2 = "2023-01-01"
        date_str3 = "Jan 1, 2023"

        dt = datetime(2023, 1, 1)

        self.assertEqual(natural_sort_key(date_str1)[0], 3)
        self.assertEqual(natural_sort_key(date_str2)[0], 3)
        self.assertEqual(natural_sort_key(date_str3)[0], 3)
        self.assertEqual(natural_sort_key(date_str1)[1], natural_sort_key(dt)[1])  # Timestamps should match

    def test_unknown_types(self):
        # Here we use a custom class for testing unknown types
        class Unknown:
            pass

        unknown = Unknown()
        self.assertEqual(natural_sort_key(unknown)[0], 5)  # Success case, string conversion works

    def test_unknown_types_failure(self):
        # Create an object where string conversion fails
        class Unconvertible:
            def __str__(self):
                raise Exception("String conversion fails")

            def __repr__(self):
                raise Exception("String conversion fails")

        unconvertible = Unconvertible()
        self.assertEqual(natural_sort_key(unconvertible)[0], 6)  # Failure case, string conversion fails


if __name__ == "__main__":
    unittest.main()
