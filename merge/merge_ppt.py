"""
merge_ppt.py
-------------
Collect PowerPoint files from every sub-folder of a root directory and merge
all of their slides into ONE presentation.

Example layout:
    my_decks/
        folder1/intro.pptx
        folder2/sales.pptx
        folder3/marketing.pptx
        folder4/finance.pptx
        folder5/summary.pptx

Usage:
    pip install python-pptx
    python merge_ppt.py my_decks -o combined.pptx

Options:
    -o / --output     Output file name (default: merged_presentation.pptx)
    --recursive       Also look in nested sub-folders (folder1/sub/...)
    --method com      Windows only: use installed PowerPoint (needs pywin32).
                      Handles .ppt files and complex content most faithfully.

Two merge engines:
  * "pptx" (default, any OS) - pure Python using python-pptx. Copies text,
    shapes, tables, pictures, charts, hyperlinks, slide backgrounds and
    speaker notes. The first deck is used as the base, so its theme/slide size
    is kept; slides from other decks are placed on the base's blank layout.
  * "com" (Windows + Microsoft PowerPoint) - asks PowerPoint itself to insert
    the slides, which also supports old .ppt files.

The code is split into three files, which must sit in the same folder:
    ppt_finder.py        - finds the PowerPoint files      (part 1)
    ppt_merge_engine.py  - merges the slides               (part 2)
    ppt_cli.py           - reads the command-line options  (part 3)
This file brings them together, so the tool still starts with
"python merge_ppt.py" and merge_ppt_gui.py keeps working unchanged.
"""

from ppt_finder import PPT_EXTENSIONS, find_ppt_files
from ppt_merge_engine import (
    copy_slide,
    merge_with_powerpoint_com,
    merge_with_python_pptx,
)
from ppt_cli import main

__all__ = [
    "PPT_EXTENSIONS",
    "find_ppt_files",
    "copy_slide",
    "merge_with_python_pptx",
    "merge_with_powerpoint_com",
    "main",
]

if __name__ == "__main__":
    main()
