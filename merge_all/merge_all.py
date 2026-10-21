"""
merge_all.py
------------
Merge Tool: finds every PowerPoint and Excel file in a folder and all the
folders nested inside it, then creates up to two files:

    combinedppt_2026-10-20_21-09-30.pptx     all the slides, in path order
    combinedexcel_2026-10-20_21-09-30.xlsx   all the sheets, one per sheet

A file is only created when that type was found. Both names share the
date and time the merge started.

Usage:
    pip install python-pptx openpyxl
    python merge_all_gui.py                          (the window)
    python merge_all.py "C:\\path\\to\\folder" -o "C:\\path\\to\\save"

These files must sit in the same folder:
    merge_all.py           finds the files and runs both merges (this file)
    merge_all_gui.py       the window
    ppt_merge_engine.py    merges the slides  (same as the PPT Merge Tool)
    excel_merge_engine.py  merges the sheets  (same as the Excel Merge Tool)
"""

import argparse
import sys
from datetime import datetime
from pathlib import Path

PPT_TYPES = {".pptx"}
EXCEL_TYPES = {".xlsx", ".xlsm", ".csv"}
OLD_TYPES = {".ppt": "old .ppt format, save it as .pptx",
             ".xls": "old .xls format, save it as .xlsx",
             ".xlsb": "binary .xlsb format, save it as .xlsx"}
OUTPUT_PREFIXES = ("combinedppt_", "combinedexcel_")


def find_files(root: Path):
    """Return (ppt_files, excel_files) found anywhere under root,
    including root itself, sorted by their path inside root."""
    ppt_files, excel_files = [], []
    for f in sorted(root.rglob("*"), key=lambda p: str(p.relative_to(root)).lower()):
        if not f.is_file():
            continue
        name, ext = f.name.lower(), f.suffix.lower()
        rel = f.relative_to(root)
        if name.startswith("~$"):                    # Office lock file
            continue
        if name.startswith(OUTPUT_PREFIXES):         # a result of an earlier run
            print(f"  [skip]  {rel}: created by an earlier merge")
            continue
        if ext in PPT_TYPES:
            print(f"  [found] {rel}")
            ppt_files.append(f)
        elif ext in EXCEL_TYPES:
            print(f"  [found] {rel}")
            excel_files.append(f)
        elif ext in OLD_TYPES:
            print(f"  [skip]  {rel}: {OLD_TYPES[ext]}")
    return ppt_files, excel_files


def output_names(stamp: datetime = None):
    """File names for this run, e.g. combinedppt_2026-10-20_21-09-30.pptx.
    Colons are not allowed in Windows file names, so the time uses dashes."""
    text = (stamp or datetime.now()).strftime("%Y-%m-%d_%H-%M-%S")
    return f"combinedppt_{text}.pptx", f"combinedexcel_{text}.xlsx"


def merge_all(input_dir: Path, output_dir: Path) -> list:
    """Find the files, run each merge, and return the paths created.
    Raises ValueError with a readable message when nothing can be made."""
    input_dir, output_dir = Path(input_dir), Path(output_dir)
    if not input_dir.is_dir():
        raise ValueError(f"Folder not found:\n{input_dir}")
    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"Scanning {input_dir} ...")
    ppt_files, excel_files = find_files(input_dir)
    if not ppt_files and not excel_files:
        raise ValueError("No PowerPoint (.pptx) or Excel (.xlsx, .xlsm, .csv) files "
                         f"were found in\n{input_dir}\nor the folders inside it.")

    ppt_name, excel_name = output_names()
    created, problems = [], []

    if ppt_files:
        print(f"\nPowerPoint: {len(ppt_files)} file(s)")
        from ppt_merge_engine import merge_with_python_pptx
        target = output_dir / ppt_name
        try:
            slides = merge_with_python_pptx(ppt_files, target)
            print(f"  -> {ppt_name} ({slides} slides)")
            created.append(target)
        except SystemExit as e:              # the engine stops with a message
            problems.append(f"PowerPoint: {e}")
            print(f"  PowerPoint merge stopped: {e}")

    if excel_files:
        print(f"\nExcel: {len(excel_files)} file(s)")
        from excel_merge_engine import merge_excel_files
        target = output_dir / excel_name
        try:
            sheets = merge_excel_files(excel_files, target, root=input_dir)
            print(f"  -> {excel_name} ({sheets} sheets)")
            created.append(target)
        except SystemExit as e:
            problems.append(f"Excel: {e}")
            print(f"  Excel merge stopped: {e}")

    if not created:
        raise ValueError("No files could be created.\n\n" + "\n".join(problems))
    return created


def main():
    parser = argparse.ArgumentParser(
        description="Merge all PowerPoint files into one .pptx and all Excel files "
                    "into one .xlsx, searching every nested folder.")
    parser.add_argument("root", help="Folder that contains the files")
    parser.add_argument("-o", "--output-dir", default=None,
                        help="Folder to save the created files (default: the same folder)")
    args = parser.parse_args()
    try:
        created = merge_all(Path(args.root), Path(args.output_dir or args.root))
    except ValueError as e:
        sys.exit(str(e))
    print("\nFiles created:")
    for f in created:
        print(f"  {f}")


if __name__ == "__main__":
    main()
