"""
merge_excel_gui.py
------------------
A simple window for the Excel Merge Tool.

Put this file in the SAME folder as merge_excel.py, then run:
    python merge_excel_gui.py
"""

import contextlib
import os
import platform
import queue
import subprocess
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

import merge_excel  # the command-line tool in the same folder

IS_WINDOWS = platform.system() == "Windows"


# --------------------------------------------------------------------------- #
# Merge logic (no window code here, so it can be tested on its own)
# --------------------------------------------------------------------------- #
def run_merge(input_dir: Path, output_file: Path, nested: bool = True,
              keep_formulas: bool = False, contents: bool = True) -> int:
    """Find and merge the workbooks. Returns the sheet count.
    Raises ValueError with a readable message if something is wrong."""
    files = merge_excel.find_excel_files(input_dir, recursive=nested, exclude=output_file)
    if not files:
        raise ValueError("No Excel files (.xlsx, .xlsm, .csv) were found in\n"
                         f"{input_dir}")
    print()
    try:
        return merge_excel.merge_excel_files(files, output_file, root=input_dir,
                                             keep_formulas=keep_formulas,
                                             contents=contents)
    except SystemExit as e:  # the engine stops with SystemExit("message")
        raise ValueError(str(e)) from None


def open_path(path: Path):
    """Open a file or folder with the system's default program."""
    if IS_WINDOWS:
        os.startfile(str(path))
    elif platform.system() == "Darwin":
        subprocess.run(["open", str(path)])
    else:
        subprocess.run(["xdg-open", str(path)])


class QueueWriter:
    """Sends print() output from the merge into the window's log."""
    def __init__(self, q):
        self.q = q

    def write(self, text):
        if text:
            self.q.put(("log", text))

    def flush(self):
        pass


# --------------------------------------------------------------------------- #
# The window
# --------------------------------------------------------------------------- #
class MergeApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        root.title("Excel Merge Tool")
        root.minsize(640, 500)

        self.q = queue.Queue()
        self.result_file = None

        self.input_var = tk.StringVar()
        self.out_dir_var = tk.StringVar()
        self.out_name_var = tk.StringVar(value="combined.xlsx")
        self.nested_var = tk.BooleanVar(value=True)
        self.formulas_var = tk.BooleanVar(value=False)
        self.contents_var = tk.BooleanVar(value=True)

        self._build()
        root.after(100, self._poll_queue)

    def _build(self):
        frm = ttk.Frame(self.root, padding=14)
        frm.pack(fill="both", expand=True)
        frm.columnconfigure(1, weight=1)
        pad = {"padx": 6, "pady": 5}

        ttk.Label(frm, text="Excel Merge Tool", font=("Segoe UI", 14, "bold")).grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0, 2))
        ttk.Label(frm, text="Combines Excel files from a folder and its sub-folders into one "
                            "workbook, one sheet per source sheet.",
                  foreground="gray").grid(row=1, column=0, columnspan=3, sticky="w", pady=(0, 10))

        ttk.Label(frm, text="Folder with the files:").grid(row=2, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.input_var).grid(row=2, column=1, sticky="ew", **pad)
        ttk.Button(frm, text="Browse...", command=self._pick_input).grid(row=2, column=2, **pad)

        ttk.Label(frm, text="Save to folder:").grid(row=3, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.out_dir_var).grid(row=3, column=1, sticky="ew", **pad)
        ttk.Button(frm, text="Browse...", command=self._pick_output).grid(row=3, column=2, **pad)

        ttk.Label(frm, text="File name:").grid(row=4, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.out_name_var).grid(row=4, column=1, sticky="ew", **pad)

        opts = ttk.Frame(frm)
        opts.grid(row=5, column=1, columnspan=2, sticky="w", **pad)
        ttk.Checkbutton(opts, text="Search all nested folders",
                        variable=self.nested_var).grid(row=0, column=0, sticky="w", padx=(0, 16))
        ttk.Checkbutton(opts, text="Add a Contents sheet",
                        variable=self.contents_var).grid(row=0, column=1, sticky="w")
        ttk.Checkbutton(opts, text="Keep formulas (instead of their results)",
                        variable=self.formulas_var).grid(row=1, column=0, columnspan=2, sticky="w")

        btns = ttk.Frame(frm)
        btns.grid(row=6, column=0, columnspan=3, sticky="ew", pady=(10, 6))
        self.merge_btn = ttk.Button(btns, text="Merge", command=self._start_merge)
        self.merge_btn.pack(side="left")
        self.open_file_btn = ttk.Button(btns, text="Open merged file", state="disabled",
                                        command=lambda: self._safe_open(self.result_file))
        self.open_file_btn.pack(side="left", padx=6)
        self.open_dir_btn = ttk.Button(btns, text="Open folder", state="disabled",
                                       command=lambda: self._safe_open(self.result_file.parent))
        self.open_dir_btn.pack(side="left")

        self.progress = ttk.Progressbar(frm, mode="indeterminate")
        self.progress.grid(row=7, column=0, columnspan=3, sticky="ew", pady=(0, 6))

        ttk.Label(frm, text="Progress:").grid(row=8, column=0, sticky="w")
        self.log = scrolledtext.ScrolledText(frm, height=12, font=("Consolas", 9),
                                             state="disabled", wrap="word")
        self.log.grid(row=9, column=0, columnspan=3, sticky="nsew", pady=(2, 0))
        frm.rowconfigure(9, weight=1)

    def _pick_input(self):
        folder = filedialog.askdirectory(title="Choose the folder that contains the Excel files")
        if folder:
            self.input_var.set(str(Path(folder)))
            if not self.out_dir_var.get():
                self.out_dir_var.set(str(Path(folder)))

    def _pick_output(self):
        folder = filedialog.askdirectory(title="Choose where to save the merged file")
        if folder:
            self.out_dir_var.set(str(Path(folder)))

    def _start_merge(self):
        input_dir = Path(self.input_var.get().strip().strip('"'))
        out_dir_text = self.out_dir_var.get().strip().strip('"')
        name = self.out_name_var.get().strip()

        if not self.input_var.get().strip() or not input_dir.is_dir():
            messagebox.showerror("Folder not found", "Please choose the folder that contains the Excel files.")
            return
        if not out_dir_text:
            messagebox.showerror("No save folder", "Please choose where to save the merged file.")
            return
        if not name:
            messagebox.showerror("No file name", "Please enter a name for the merged file.")
            return
        if any(c in name for c in '\\/:*?"<>|'):
            messagebox.showerror("Invalid file name", 'A file name cannot contain \\ / : * ? " < > |')
            return
        if not name.lower().endswith(".xlsx"):
            name = Path(name).stem + ".xlsx" if name.lower().endswith((".xls", ".xlsm")) else name + ".xlsx"
            self.out_name_var.set(name)

        out_dir = Path(out_dir_text)
        try:
            out_dir.mkdir(parents=True, exist_ok=True)
        except OSError as e:
            messagebox.showerror("Cannot use save folder", str(e))
            return
        output_file = out_dir / name

        if output_file.exists() and not messagebox.askyesno(
                "Replace file?", f"{output_file.name} already exists.\n\nReplace it?"):
            return

        self.merge_btn.config(state="disabled")
        self.open_file_btn.config(state="disabled")
        self.open_dir_btn.config(state="disabled")
        self._clear_log()
        self.progress.start(12)

        threading.Thread(
            target=self._worker,
            args=(input_dir, output_file, self.nested_var.get(),
                  self.formulas_var.get(), self.contents_var.get()),
            daemon=True,
        ).start()

    def _worker(self, input_dir, output_file, nested, keep_formulas, contents):
        """Runs in the background so the window stays responsive."""
        try:
            with contextlib.redirect_stdout(QueueWriter(self.q)):
                total = run_merge(input_dir, output_file, nested, keep_formulas, contents)
            self.q.put(("done", (total, output_file)))
        except PermissionError:
            self.q.put(("error", f"Could not save {output_file.name}.\n\n"
                                 "If it is open in Excel, close it and try again."))
        except Exception as e:
            self.q.put(("error", str(e) or type(e).__name__))

    def _poll_queue(self):
        try:
            while True:
                kind, data = self.q.get_nowait()
                if kind == "log":
                    self._append_log(data)
                elif kind == "done":
                    self._finish(*data)
                elif kind == "error":
                    self._fail(data)
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)

    def _finish(self, total, output_file):
        self.progress.stop()
        self.merge_btn.config(state="normal")
        self.result_file = output_file
        self.open_file_btn.config(state="normal")
        self.open_dir_btn.config(state="normal")
        self._append_log(f"\nDone! {total} sheets saved to {output_file}\n")
        if messagebox.askyesno("Merge complete",
                               f"{total} sheets saved to:\n{output_file}\n\nOpen it now?"):
            self._safe_open(output_file)

    def _fail(self, message):
        self.progress.stop()
        self.merge_btn.config(state="normal")
        self._append_log(f"\nERROR: {message}\n")
        messagebox.showerror("Merge failed", message)

    def _safe_open(self, path):
        try:
            open_path(path)
        except Exception as e:
            messagebox.showerror("Cannot open", str(e))

    def _append_log(self, text):
        self.log.config(state="normal")
        self.log.insert("end", text)
        self.log.see("end")
        self.log.config(state="disabled")

    def _clear_log(self):
        self.log.config(state="normal")
        self.log.delete("1.0", "end")
        self.log.config(state="disabled")


def main():
    root = tk.Tk()
    try:
        ttk.Style().theme_use("vista" if IS_WINDOWS else "clam")
    except tk.TclError:
        pass
    MergeApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
