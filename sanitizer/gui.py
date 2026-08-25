import os
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

from .cli import collect_inputs
from .convert import to_docx
from .sanitize import sanitize_document

SUPPORTED_LABEL = "PDF / DOC / DOCX"


class SanitizerGUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Document Sanitizer")
        self.root.geometry("760x560")
        self.root.minsize(640, 480)

        self.items = []
        self.output_dir = tk.StringVar(
            value=os.path.join(os.path.expanduser("~"), "Documents", "Sanitized")
        )
        self.worker = None
        self.log_queue = queue.Queue()

        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 10, "pady": 6}

        top = ttk.LabelFrame(self.root, text=f"Input files & folders ({SUPPORTED_LABEL})")
        top.pack(fill="both", expand=True, **pad)

        btns = ttk.Frame(top)
        btns.pack(fill="x", padx=6, pady=4)
        ttk.Button(btns, text="Add files…", command=self.add_files).pack(side="left", padx=2)
        ttk.Button(btns, text="Add folder…", command=self.add_folder).pack(side="left", padx=2)
        ttk.Button(btns, text="Remove selected", command=self.remove_selected).pack(side="left", padx=2)
        ttk.Button(btns, text="Clear all", command=self.clear_all).pack(side="left", padx=2)

        self.listbox = tk.Listbox(top, selectmode=tk.EXTENDED, activestyle="dotbox")
        self.listbox.pack(fill="both", expand=True, padx=6, pady=(0, 6))

        out = ttk.LabelFrame(self.root, text="Output folder")
        out.pack(fill="x", **pad)
        ttk.Entry(out, textvariable=self.output_dir).pack(
            side="left", fill="x", expand=True, padx=6, pady=6
        )
        ttk.Button(out, text="Browse…", command=self.pick_output).pack(side="left", padx=6, pady=6)

        actions = ttk.Frame(self.root)
        actions.pack(fill="x", **pad)
        self.start_btn = ttk.Button(
            actions, text="Sanitize", command=self.start, state="disabled"
        )
        self.start_btn.pack(side="right", padx=6)
        self.progress = ttk.Progressbar(actions, mode="indeterminate", length=180)
        self.progress.pack(side="right")

        log = ttk.LabelFrame(self.root, text="Log")
        log.pack(fill="both", expand=True, **pad)
        self.log_text = scrolledtext.ScrolledText(log, height=10, state="disabled")
        self.log_text.pack(fill="both", expand=True, padx=6, pady=6)

    def add_files(self):
        paths = filedialog.askopenfilenames(
            title="Select documents",
            filetypes=[
                ("Documents", "*.pdf *.doc *.docx"),
                ("PDF", "*.pdf"),
                ("Word", "*.doc *.docx"),
                ("All files", "*.*"),
            ],
        )
        self._add(paths)

    def add_folder(self):
        path = filedialog.askdirectory(title="Select a folder of documents")
        if path:
            self._add([path])

    def _add(self, paths):
        known = set(self.items)
        for p in collect_inputs(paths):
            if p not in known:
                self.items.append(p)
                known.add(p)
                self.listbox.insert(tk.END, str(p))
        self._refresh_start()

    def remove_selected(self):
        for idx in reversed(self.listbox.curselection()):
            del self.items[idx]
            self.listbox.delete(idx)
        self._refresh_start()

    def clear_all(self):
        self.items.clear()
        self.listbox.delete(0, tk.END)
        self._refresh_start()

    def pick_output(self):
        path = filedialog.askdirectory(title="Select output folder")
        if path:
            self.output_dir.set(path)

    def _refresh_start(self):
        self.start_btn.config(state="normal" if self.items else "disabled")

    def log(self, msg: str):
        self.log_queue.put(msg)

    def start(self):
        if self.worker is not None:
            return
        out_dir = self.output_dir.get().strip()
        if not out_dir:
            messagebox.showerror("Document Sanitizer", "Please choose an output folder.")
            return
        try:
            os.makedirs(out_dir, exist_ok=True)
        except OSError as e:
            messagebox.showerror("Document Sanitizer", f"Cannot create output folder:\n{e}")
            return

        items = list(self.items)
        self.start_btn.config(state="disabled")
        self.progress.start(12)
        self.log(f"Starting batch of {len(items)} item(s)…")
        self.worker = threading.Thread(target=self._run, args=(items, out_dir), daemon=True)
        self.worker.start()

    def _run(self, items, out_dir):
        from docx import Document

        failures = 0
        try:
            for i, path in enumerate(items, 1):
                self.log(f"[{i}/{len(items)}] {os.path.basename(str(path))}")
                try:
                    docx_path, is_temp = to_docx(path)
                    doc = Document(str(docx_path))
                    stats = sanitize_document(doc)
                    tag = "" if path.suffix.lower() == ".docx" else "_" + path.suffix.lower().lstrip(".")
                    out_path = os.path.join(out_dir, f"{path.stem}{tag}_sanitized.docx")
                    doc.save(out_path)
                    if is_temp:
                        import shutil

                        shutil.rmtree(docx_path.parent, ignore_errors=True)
                    self.log(
                        f"    done -> {os.path.basename(out_path)} "
                        f"(watermarks removed: {stats['watermarks_removed']}, "
                        f"words replaced: {stats['words_replaced']}, "
                        f"sections fixed: {stats['sections_fixed']}, "
                        f"shapes scaled: {stats['shapes_scaled']})"
                    )
                except Exception as e:
                    failures += 1
                    self.log(f"    FAILED: {e}")
        finally:
            self.log_queue.put(
                f"Finished with {failures} failure(s). Output in: {out_dir}"
            )
            self.log_queue.put("__DONE__")


def main():
    root = tk.Tk()
    try:
        ttk.Style().theme_use("clam")
    except tk.TclError:
        pass

    def pump():
        try:
            while True:
                msg = gui.log_queue.get_nowait()
                if msg == "__DONE__":
                    gui.progress.stop()
                    gui.start_btn.config(state="normal" if gui.items else "disabled")
                    gui.worker = None
                else:
                    gui.log_text.config(state="normal")
                    gui.log_text.insert(tk.END, msg + "\n")
                    gui.log_text.see(tk.END)
                    gui.log_text.config(state="disabled")
        except queue.Empty:
            pass
        root.after(100, pump)

    gui = SanitizerGUI(root)
    pump()
    root.mainloop()


if __name__ == "__main__":
    main()
