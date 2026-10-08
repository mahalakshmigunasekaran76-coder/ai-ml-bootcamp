"""
excel_cli.py
------------
Part 3 of the Excel Merge Tool: the command-line entry point.

Reads the folder and options typed after the command, then uses
excel_finder to find the files and excel_merge_engine to merge them.
"""

import argparse
import sys
from pathlib import Path

from excel_finder import find_excel_files
from excel_merge_engine import merge_excel_files


def main():
    parser = argparse.ArgumentParser(
        description="Merge the Excel files in a folder and its nested folders "
                    "into one workbook, one sheet per source sheet.")
    parser.add_argument("root", help="Folder that contains the Excel files")
    parser.add_argument("-o", "--output", default="merged_workbook.xlsx",
                        help="Output file (default: merged_workbook.xlsx)")
    parser.add_argument("--top-level-only", action="store_true",
                        help="Only look one folder deep, not in every nested folder")
    parser.add_argument("--keep-formulas", action="store_true",
                        help="Keep formulas instead of their saved results")
    parser.add_argument("--no-contents", action="store_true",
                        help="Leave out the Contents sheet")
    args = parser.parse_args()

    root = Path(args.root)
    if not root.is_dir():
        sys.exit(f"Folder not found: {root}")

    output = Path(args.output)
    if output.suffix.lower() != ".xlsx":
        output = output.with_suffix(".xlsx")

    print(f"Scanning {root.resolve()} ...")
    files = find_excel_files(root, recursive=not args.top_level_only, exclude=output)
    if not files:
        sys.exit("No Excel files found.")

    print()
    total = merge_excel_files(files, output, root=root,
                              keep_formulas=args.keep_formulas,
                              contents=not args.no_contents)
    print(f"\nDone! {total} sheets saved to {output.resolve()}")


if __name__ == "__main__":
    main()
