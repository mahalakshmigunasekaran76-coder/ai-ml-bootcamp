"""
merge_all_gui.py
----------------
The window for the Merge Tool. Put it in the same folder as merge_all.py,
ppt_merge_engine.py and excel_merge_engine.py, then run:

    python merge_all_gui.py

1. Choose the folder with the files (every nested folder is included).
2. Choose the folder to save the created files.
3. Click "Create files". A message lists the files that were created.
"""

import contextlib
import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

import merge_all


class QueueWriter:
    """Sends print() output from the merge into the window's log."""
    def __init__(self, q):
        self.q = q

    def write(self, text):
        if text:
            self.q.put(("log", text))

    def flush(self):
        pass


class MergeApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        root.title("Merge Tool")
        root.minsize(620, 440)
        self.q = queue.Queue()
        self.input_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self._build()
        root.after(100, self._poll_queue)

    # ---- layout -------------------------------------------------------------
    def _build(self):
        frm = ttk.Frame(self.root, padding=14)
        frm.pack(fill="both", expand=True)
        frm.columnconfigure(1, weight=1)
        pad = {"padx": 6, "pady": 5}

        ttk.Label(frm, text="Merge Tool", font=("Segoe UI", 14, "bold")).grid(
            row=0, column=0, columnspan=3, sticky="w")
        ttk.Label(frm, text="Combines all PowerPoint files into one presentation and all "
                            "Excel files into one workbook. Nested folders are included.",
                  foreground="gray", wraplength=580).grid(
            row=1, column=0, columnspan=3, sticky="w", pady=(2, 12))

        ttk.Label(frm, text="Folder with the files:").grid(row=2, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.input_var).grid(row=2, column=1, sticky="ew", **pad)
        ttk.Button(frm, text="Browse...", command=self._pick_input).grid(row=2, column=2, **pad)

        ttk.Label(frm, text="Save created files to:").grid(row=3, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.output_var).grid(row=3, column=1, sticky="ew", **pad)
        ttk.Button(frm, text="Browse...", command=self._pick_output).grid(row=3, column=2, **pad)

        self.merge_btn = ttk.Button(frm, text="Create files", command=self._start)
        self.merge_btn.grid(row=4, column=0, sticky="w", padx=6, pady=(12, 6))

        self.progress = ttk.Progressbar(frm, mode="indeterminate")
        self.progress.grid(row=5, column=0, columnspan=3, sticky="ew", padx=6, pady=(0, 6))

        ttk.Label(frm, text="Progress:").grid(row=6, column=0, sticky="w", padx=6)
        self.log = scrolledtext.ScrolledText(frm, height=12, font=("Consolas", 9),
                                             state="disabled", wrap="word")
        self.log.grid(row=7, column=0, columnspan=3, sticky="nsew", padx=6, pady=(2, 0))
        frm.rowconfigure(7, weight=1)

    # ---- folder pickers -----------------------------------------------------
    def _pick_input(self):
        folder = filedialog.askdirectory(title="Choose the folder with the PowerPoint and Excel files")
        if folder:
            self.input_var.set(str(Path(folder)))
            if not self.output_var.get():            # default: save in the same folder
                self.output_var.set(str(Path(folder)))

    def _pick_output(self):
        folder = filedialog.askdirectory(title="Choose where to save the created files")
        if folder:
            self.output_var.set(str(Path(folder)))

    # ---- merging ------------------------------------------------------------
    def _start(self):
        in_text = self.input_var.get().strip().strip('"')
        out_text = self.output_var.get().strip().strip('"')
        if not in_text or not Path(in_text).is_dir():
            messagebox.showerror("Folder not found", "Choose the folder that contains the files.")
            return
        if not out_text:
            messagebox.showerror("No save folder", "Choose the folder to save the created files.")
            return
        self.merge_btn.config(state="disabled")
        self._clear_log()
        self.progress.start(12)
        threading.Thread(target=self._worker, args=(Path(in_text), Path(out_text)),
                         daemon=True).start()

    def _worker(self, input_dir, output_dir):
        """Runs in the background so the window stays responsive."""
        try:
            with contextlib.redirect_stdout(QueueWriter(self.q)):
                created = merge_all.merge_all(input_dir, output_dir)
            self.q.put(("done", created))
        except PermissionError as e:
            self.q.put(("error", f"A file could not be saved or read:\n{e.filename or e}\n\n"
                                 "Close it in PowerPoint or Excel, and check you can save "
                                 "to the chosen folder."))
        except Exception as e:
            self.q.put(("error", str(e) or type(e).__name__))

    def _poll_queue(self):
        try:
            while True:
                kind, data = self.q.get_nowait()
                if kind == "log":
                    self._append_log(data)
                elif kind == "done":
                    self._finish(data)
                elif kind == "error":
                    self._fail(data)
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)

    def _finish(self, created):
        self.progress.stop()
        self.merge_btn.config(state="normal")
        names = "\n".join(f.name for f in created)
        self._append_log(f"\nFiles created in {created[0].parent}:\n{names}\n")
        # Information message: an OK button, plus the window's close button
        messagebox.showinfo("Files created", f"Files created:\n\n{names}")

    def _fail(self, message):
        self.progress.stop()
        self.merge_btn.config(state="normal")
        self._append_log(f"\nERROR: {message}\n")
        messagebox.showerror("Merge failed", message)

    # ---- log helpers --------------------------------------------------------
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
        ttk.Style().theme_use("vista")
    except tk.TclError:
        pass
    MergeApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
