"""
Extends `openpyxl` classes for easier operation on Excel tables.
"""

# __all__ = ["TableExt", "TableColumnExt"]

import os
import random
import re
from decimal import Decimal
from glob import glob
from typing import Dict, List, Tuple

import openpyxl
from openpyxl.cell.cell import TYPE_NUMERIC, Cell
from openpyxl.styles import Font
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.table import Table, TableColumn
from openpyxl.worksheet.worksheet import Worksheet


# * ------------------------------------------------------------------------------ # *
# * PYOPENXL CLASS EXTENSIONS * #


class TableExt:
    """
    Extends the `openpyxl.worksheet.table.Table` class from `openpyxl`.
    """

    def __init__(self, table: Table):
        self.table = table
        self.header_row_count = table.headerRowCount
        self.name = table.name
        self.ref = table.ref
        self.parent = 0
        self.table_columns = 0

    def _get_columns(self):
        """
        Test.
        """

        pass

    def _what(self):
        """
        Test.
        """

        pass


class TableColumnExt:
    """
    Extends the `openpyxl.worksheet.table.TableColumn` class from `openpyxl`.
    """

    def __init__(self, table_column: TableColumn):
        self.table_column = table_column
        self.name = table_column.name
        self.parent = None  # From `self.name` and `get_structure()`
        self.ref = None  # From `self.name` and `self.parent`
        self.cells = None  # From `self.ref`


# * ------------------------------------------------------------------------------ # *
# * WORKBOOK STRUCTURE * #


def get_structure(workbook_filename: str) -> Dict:
    """
    Get the workbook structure as a nested dictionary containing the sheet names in the
    workbook, the table names in each sheet, and the column names in each table.
    """
    # TODO: Build out the structure.
    # sheet names
    #   table names
    #     column names
    workbook_structure: Dict = {}
    return workbook_structure


# * ------------------------------------------------------------------------------ # *
# * EXCEL CELL/COLUMN TRANSLATORS * #


def split_xlscell(xlscell: str) -> Tuple[str, int]:
    """
    Split an Excel cell designation (e.g. `"B2"`) into its respective Excel-designated
    column and row (e.g. `"B"` and `"2"`).
    """

    xlscell_pattern = r"(?P<xlscol>[A-Z]+)(?P<row>[0-9]+)"
    match = re.match(xlscell_pattern, xlscell)

    if not match:
        raise ValueError(f"Argument '{xlscell}' is not an Excel cell designation.")

    return match.group("xlscol"), int(match.group("row"))


def num_to_xlscol(num: int) -> str:
    """
    Get the string representation of a column as in Excel (e.g. `"B"` or `"AC"`), given
    the number of that column counted starting from `"A"` (e.g. `2` or `29`).
    """

    quotient, remainder = divmod(num - 1, 26)
    last_chr = chr(ord("A") + remainder)
    if quotient > 0:
        xlscol = num_to_xlscol(quotient) + last_chr
    else:
        xlscol = last_chr

    return xlscol


def xlscol_to_num(xlscol: str):
    """
    Get the number of a column counted starting from `"A"` (e.g. `2` or `29`), given the
    string representation of a column as in Excel (e.g. `"B"` or `"AC"`).
    """

    num_last_char = ord(xlscol[-1]) - (ord("A") - 1)
    if len(xlscol) > 1:
        num = num_last_char + 26 * (xlscol_to_num(xlscol[:-1]))
    else:
        num = num_last_char

    return num
