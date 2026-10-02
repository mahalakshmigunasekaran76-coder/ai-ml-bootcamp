"""
merge_ppt_gui.py
----------------
A simple window for the PPT Merge Tool.

Put this file in the SAME folder as merge_ppt.py, then run:
    python merge_ppt_gui.py

The window lets you:
  1. Choose the folder that contains the deck sub-folders (ppt1, ppt2, ...)
  2. Choose where to save the merged file and what to call it
  3. Click "Merge" and watch the progress
  4. Open the merged file or its folder when it's done
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

import merge_ppt  # the command-line tool in the same folder

IS_WINDOWS = platform.system() == "Windows"


# --------------------------------------------------------------------------- #
# Merge logic (no window code here, so it can be tested on its own)
# --------------------------------------------------------------------------- #
def run_merge(input_dir: Path, output_file: Path,
              recursive: bool = False, use_powerpoint: bool = False) -> int:
    """Find and merge the decks. Returns the slide count.
    Raises ValueError with a readable message if something is wrong."""
    files = [f for f in merge_ppt.find_ppt_files(input_dir, recursive)
             if f.resolve() != output_file.resolve()]
    if not files:
        raise ValueError("No PowerPoint files were found in the sub-folders of\n"
                         f"{input_dir}\n\nEach deck must be in its own sub-folder.")
    try:
        if use_powerpoint:
            return merge_ppt.merge_with_powerpoint_com(files, output_file)
        return merge_ppt.merge_with_python_pptx(files, output_file)
    except SystemExit as e:  # merge_ppt stops with sys.exit("message")
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
        root.title("PPT Merge Tool")
        root.minsize(640, 480)

        self.q = queue.Queue()
        self.result_file = None

        self.input_var = tk.StringVar()
        self.out_dir_var = tk.StringVar()
        self.out_name_var = tk.StringVar(value="combined.pptx")
        self.recursive_var = tk.BooleanVar(value=False)
        self.com_var = tk.BooleanVar(value=False)

        self._build()
        root.after(100, self._poll_queue)

    # ---- layout -------------------------------------------------------------
    def _build(self):
        frm = ttk.Frame(self.root, padding=14)
        frm.pack(fill="both", expand=True)
        frm.columnconfigure(1, weight=1)
        pad = {"padx": 6, "pady": 5}

        ttk.Label(frm, text="PPT Merge Tool", font=("Segoe UI", 14, "bold")).grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0, 2))
        ttk.Label(frm, text="Combines the PowerPoint file in each sub-folder into one presentation.",
                  foreground="gray").grid(row=1, column=0, columnspan=3, sticky="w", pady=(0, 10))

        # 1. Input folder
        ttk.Label(frm, text="Folder with the decks:").grid(row=2, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.input_var).grid(row=2, column=1, sticky="ew", **pad)
        ttk.Button(frm, text="Browse...", command=self._pick_input).grid(row=2, column=2, **pad)

        # 2. Output folder
        ttk.Label(frm, text="Save to folder:").grid(row=3, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.out_dir_var).grid(row=3, column=1, sticky="ew", **pad)
        ttk.Button(frm, text="Browse...", command=self._pick_output).grid(row=3, column=2, **pad)

        # 3. Output file name
        ttk.Label(frm, text="File name:").grid(row=4, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.out_name_var).grid(row=4, column=1, sticky="ew", **pad)

        # Options
        opts = ttk.Frame(frm)
        opts.grid(row=5, column=1, sticky="w", **pad)
        ttk.Checkbutton(opts, text="Also search nested folders",
                        variable=self.recursive_var).pack(side="left", padx=(0, 16))
        if IS_WINDOWS:
            ttk.Checkbutton(opts, text="Use installed PowerPoint (also merges .ppt)",
                            variable=self.com_var).pack(side="left")

        # Buttons
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

        # Log
        ttk.Label(frm, text="Progress:").grid(row=8, column=0, sticky="w")
        self.log = scrolledtext.ScrolledText(frm, height=12, font=("Consolas", 9),
                                             state="disabled", wrap="word")
        self.log.grid(row=9, column=0, columnspan=3, sticky="nsew", pady=(2, 0))
        frm.rowconfigure(9, weight=1)

    # ---- folder pickers -----------------------------------------------------
    def _pick_input(self):
        folder = filedialog.askdirectory(title="Choose the folder that contains ppt1, ppt2, ...")
        if folder:
            self.input_var.set(str(Path(folder)))
            if not self.out_dir_var.get():          # default: save next to the decks
                self.out_dir_var.set(str(Path(folder)))

    def _pick_output(self):
        folder = filedialog.askdirectory(title="Choose where to save the merged file")
        if folder:
            self.out_dir_var.set(str(Path(folder)))

    # ---- merging ------------------------------------------------------------
    def _start_merge(self):
        input_dir = Path(self.input_var.get().strip().strip('"'))
        out_dir_text = self.out_dir_var.get().strip().strip('"')
        name = self.out_name_var.get().strip()

        if not self.input_var.get().strip() or not input_dir.is_dir():
            messagebox.showerror("Folder not found", "Please choose the folder that contains the decks.")
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
        if not name.lower().endswith(".pptx"):
            name += ".pptx"
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

        # Lock the UI while working
        self.merge_btn.config(state="disabled")
        self.open_file_btn.config(state="disabled")
        self.open_dir_btn.config(state="disabled")
        self._clear_log()
        self.progress.start(12)

        threading.Thread(
            target=self._worker,
            args=(input_dir, output_file, self.recursive_var.get(), self.com_var.get()),
            daemon=True,
        ).start()

    def _worker(self, input_dir, output_file, recursive, use_powerpoint):
        """Runs in the background so the window stays responsive."""
        try:
            if use_powerpoint:
                import pythoncom            # PowerPoint needs this on a background thread
                pythoncom.CoInitialize()
            with contextlib.redirect_stdout(QueueWriter(self.q)):
                total = run_merge(input_dir, output_file, recursive, use_powerpoint)
            self.q.put(("done", (total, output_file)))
        except PermissionError:
            self.q.put(("error", f"Could not save {output_file.name}.\n\n"
                                 "If it is open in PowerPoint, close it and try again."))
        except ImportError:
            self.q.put(("error", "The PowerPoint option needs pywin32.\n\n"
                                 "Run:  python -m pip install pywin32"))
        except Exception as e:
            self.q.put(("error", str(e) or type(e).__name__))

    def _poll_queue(self):
        """Moves messages from the background thread into the window."""
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
        if messagebox.askyesno("Merge complete",
                               f"{total} slides saved to:\n{output_file}\n\nOpen it now?"):
            self._safe_open(output_file)

    def _fail(self, message):
        self.progress.stop()
        self.merge_btn.config(state="normal")
        self._append_log(f"\nERROR: {message}\n")
        messagebox.showerror("Merge failed", message)

    # ---- helpers ------------------------------------------------------------
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
