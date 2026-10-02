"""
ppt_finder.py
-------------
Part 1 of the PPT Merge Tool: finds the PowerPoint files.

Looks inside each sub-folder of a main folder (ppt1, ppt2, ...) and returns
the PowerPoint files it finds, in alphabetical folder order.
"""

from pathlib import Path

PPT_EXTENSIONS = {".pptx", ".ppt"}


def find_ppt_files(root: Path, recursive: bool = False) -> list:
    """Return PowerPoint files from each sub-folder of root, in sorted order."""
    files = []
    subfolders = sorted(p for p in root.iterdir() if p.is_dir())
    if not subfolders:
        print(f"No sub-folders found in {root}")

    for folder in subfolders:
        pattern = "**/*" if recursive else "*"
        found = sorted(
            f for f in folder.glob(pattern)
            if f.is_file()
            and f.suffix.lower() in PPT_EXTENSIONS
            and not f.name.startswith("~$")      # skip PowerPoint lock files
        )
        if not found:
            print(f"  [skip] no PowerPoint file in {folder.name}")
        for f in found:
            print(f"  [found] {f.relative_to(root)}")
        files.extend(found)
    return files
