"""
ppt_cli.py
----------
Part 3 of the PPT Merge Tool: the command-line entry point.

Reads the folder and options typed after the command, then uses
ppt_finder to find the files and ppt_merge_engine to merge them.
"""

import argparse
import sys
from pathlib import Path

from ppt_finder import find_ppt_files
from ppt_merge_engine import merge_with_powerpoint_com, merge_with_python_pptx


def main():
    parser = argparse.ArgumentParser(description="Merge PPT files from sub-folders.")
    parser.add_argument("root", help="Folder that contains the sub-folders")
    parser.add_argument("-o", "--output", default="merged_presentation.pptx",
                        help="Output file (default: merged_presentation.pptx)")
    parser.add_argument("--recursive", action="store_true",
                        help="Search nested sub-folders too")
    parser.add_argument("--method", choices=["pptx", "com"], default="pptx",
                        help="Merge engine (default: pptx)")
    args = parser.parse_args()

    root = Path(args.root)
    if not root.is_dir():
        sys.exit(f"Folder not found: {root}")

    output = Path(args.output)
    if output.suffix.lower() != ".pptx":
        output = output.with_suffix(".pptx")

    print(f"Scanning {root.resolve()} ...")
    files = [f for f in find_ppt_files(root, args.recursive)
             if f.resolve() != output.resolve()]
    if not files:
        sys.exit("No PowerPoint files found.")

    if args.method == "com":
        total = merge_with_powerpoint_com(files, output)
    else:
        total = merge_with_python_pptx(files, output)

    print(f"\nDone! {total} slides saved to {output.resolve()}")
