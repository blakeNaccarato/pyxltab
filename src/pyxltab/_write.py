"""
Write the updated cells to the file.
"""

# TODO: Migrate useful content into `pyxltab`, remove the rest.

__all__ = ["write"]

from decimal import Decimal
import os
import random
import re
from typing import Dict, List, Tuple

import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell.cell import Cell, TYPE_NUMERIC
from openpyxl.styles import Font

FACTOR_MEAN = 1
FACTOR_STD = 0.02


# * ------------------------------------------------------------------------------ # *
# * MAIN * #


def write(
    file_structure: Dict,
    factor_mean: float = FACTOR_MEAN,
    factor_std: float = FACTOR_STD,
):
    """
    Produce Excel workbooks with column values multiplied by a factor from a Gaussian
    distribution.
    """

    # Reload the input book for each desired output book
    for book_id in file_structure["books"]:
        for out_file_id in book_id["out"]:
            book = openpyxl.load_workbook(book_id["in"])

            # For each table in each sheet of the book...
            for sheet_id in book_id["sheets"]:
                sheet = book[sheet_id["name"]]
                for table_id in sheet_id["tables"]:
                    table = sheet.tables[table_id["name"]]

                    # Get the table's position
                    table_range = get_range(table)
                    table_column_names = [col.name for col in table.tableColumns]

                    # For each column to be modified...
                    for col_id in table_id["columns"]:
                        col_num = table_column_names.index(col_id)

                        # Get the cell values of that column and modify them.
                        target_cells = get_target_cells(sheet, table_range, col_num)
                        modify_cells(target_cells, factor_mean, factor_std)

            # Save the modified input book to the output filename
            out_dir = os.path.dirname(out_file_id)
            if not os.path.exists(out_dir):
                os.makedirs(out_dir)
            book.save(out_file_id)
            book.close()


# * ------------------------------------------------------------------------------ # *
# * FUNCTIONS * #


def get_range(table):
    """
    Get an Excel cell designation (e.g. `"B2:E5"`) and return the first column letter,
    first row, and last row.
    """
    table_range = table.ref.split(":")
    [top_left, bot_right] = table_range
    (table_start_xlscol, table_start_row) = split_xlscell(top_left)
    (_, table_end_row) = split_xlscell(bot_right)
    table_range = (table_start_xlscol, table_start_row, table_end_row)

    return table_range


def get_target_cells(
    sheet: Worksheet, table_range: Tuple[str, int, int], col_num: int
) -> List[Cell]:
    """
    Get the cells corresponding to a column in the known range of a table.
    """

    (table_start_xlscol, table_start_row, table_end_row) = table_range
    xlscol = num_to_xlscol(xlscol_to_num(table_start_xlscol) + col_num)
    header_height = 1
    col_start_row = str(table_start_row + header_height)
    col_end_row = str(table_end_row)
    col_range = f"{xlscol}{col_start_row}:{xlscol}{col_end_row}"
    cells = [cell for row in sheet[col_range] for cell in row]

    return cells


def modify_cells(cells: List[Cell], factor_mean: float, factor_std: float):
    """
    Multiply cells in a sheet by a factor from a Gaussian distribution.
    """

    cell_values = [cell.value for cell in cells if cell.data_type == TYPE_NUMERIC]
    values_after_decimal = [
        str(value).split(".")[-1] for value in cell_values if "." in str(value)
    ]
    digits_after_decimal = [len(value) for value in values_after_decimal]
    if digits_after_decimal:
        max_digits_after_decimal = max(digits_after_decimal)
    else:
        max_digits_after_decimal = 0
    num_places = Decimal(10) ** -max_digits_after_decimal

    for cell in cells:
        if cell.data_type == TYPE_NUMERIC and cell.font.bold:
            cell.font = Font(bold=False)
            value = Decimal(cell.value)
            factor = Decimal(random.gauss(factor_mean, factor_std))
            cell.value = (value * factor).quantize(num_places)
        if cell.font.italic:
            cell.font = Font(italic=False)
            cell.value = None


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
