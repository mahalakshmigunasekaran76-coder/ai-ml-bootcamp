"""
excel_finder.py
---------------
Part 1 of the Excel Merge Tool: finds the Excel files.

Looks through the main folder and all folders nested inside it, and returns
the spreadsheet files it finds, sorted by their path.
"""

from pathlib import Path

SUPPORTED = {".xlsx", ".xlsm", ".csv"}
UNSUPPORTED = {".xls": "old .xls format", ".xlsb": "binary .xlsb format"}


def find_excel_files(root: Path, recursive: bool = True, exclude=None) -> list:
    """Return the supported spreadsheet files under root, sorted by path.

    recursive=True searches every nested folder; False only looks one
    level down (the files inside each sub-folder), like the PPT tool.
    exclude is a file to leave out, such as the merged output itself.
    """
    exclude = Path(exclude).resolve() if exclude else None
    if recursive:
        candidates = root.rglob("*")
    else:
        candidates = (f for folder in root.iterdir() if folder.is_dir() for f in folder.iterdir())

    files = []
    for f in sorted(candidates, key=lambda p: str(p.relative_to(root)).lower()):
        if not f.is_file() or f.name.startswith("~$"):      # skip lock files
            continue
        if exclude and f.resolve() == exclude:
            continue
        ext = f.suffix.lower()
        if ext in SUPPORTED:
            print(f"  [found] {f.relative_to(root)}")
            files.append(f)
        elif ext in UNSUPPORTED:
            print(f"  [skip]  {f.relative_to(root)}: {UNSUPPORTED[ext]}, "
                  "open it in Excel and save it as .xlsx")
    return files
