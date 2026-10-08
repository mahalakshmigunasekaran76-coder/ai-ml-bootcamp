"""
merge_excel.py
--------------
Merge Excel files from a folder (and all folders nested inside it) into ONE
workbook. Every sheet of every file becomes a sheet in the merged file, and a
"Contents" sheet at the front lists where each one came from.

Example layout:
    data/
        north/sales.xlsx          -> sheet "sales"
        south/2026/sales.xlsx     -> sheet "sales (2)"
        finance/budget.xlsx       -> sheets "budget - Q1", "budget - Q2"
        hr/headcount.csv          -> sheet "headcount"

Usage:
    pip install openpyxl
    python merge_excel.py "C:\\path\\to\\data" -o combined.xlsx

Options:
    -o / --output       Output file name (default: merged_workbook.xlsx)
    --top-level-only    Only look one folder deep (like the PPT tool)
    --keep-formulas     Keep formulas instead of their saved results
    --no-contents       Leave out the Contents sheet

The code is split into three files, which must sit in the same folder:
    excel_finder.py        - finds the Excel files          (part 1)
    excel_merge_engine.py  - copies the sheets              (part 2)
    excel_cli.py           - reads the command-line options (part 3)
This file brings them together; merge_excel_gui.py uses it too.
"""

from excel_finder import SUPPORTED, find_excel_files
from excel_merge_engine import copy_sheet, make_sheet_name, merge_excel_files
from excel_cli import main

__all__ = [
    "SUPPORTED",
    "find_excel_files",
    "copy_sheet",
    "make_sheet_name",
    "merge_excel_files",
    "main",
]

if __name__ == "__main__":
    main()
