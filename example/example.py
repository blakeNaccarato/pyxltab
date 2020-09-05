"""
Example demonstrating the usage of `pyxltab`.
"""

import pyxltab

WORKBOOK_FILENAME = r"example\\example.xlsx"


def main():
    """
    Example demonstrating the usage of `pyxltab`.
    """

    workbook_structure = pyxltab.get_structure(WORKBOOK_FILENAME)
    assert True


if __name__ == "__main__":
    main()
