import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox

from converter import (
    IMAGE_EXTS,
    convert_image_to_word,
    convert_pdf_to_word,
    convert_word_to_word,
)

FILETYPES = [
    ("Supported formats", "*.pdf *.docx *.jpg *.jpeg *.png *.bmp *.tiff *.tif *.webp"),
    ("PDF files",         "*.pdf"),
    ("Word files",        "*.docx"),
    ("Image files",       "*.jpg *.jpeg *.png *.bmp *.tiff *.tif *.webp"),
]


def _label_for(path: Path) -> str:
    ext = path.suffix.lower()
    if ext == ".pdf":
        return "PDF → Word"
    if ext == ".docx":
        return "Word → Word (table layout)"
    if ext in IMAGE_EXTS:
        return "Image → Word (OCR)"
    return "Unknown format"


def _convert(path: str) -> bytes:
    ext = Path(path).suffix.lower()
    if ext == ".pdf":
        return convert_pdf_to_word(path)
    if ext == ".docx":
        return convert_word_to_word(path)
    if ext in IMAGE_EXTS:
        return convert_image_to_word(path)
    raise ValueError(f"Unsupported format: {ext}")


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Document to Word Converter")
        self.resizable(False, False)
        self._build()

    def _build(self):
        pad = {"padx": 16, "pady": 8}

        # ── file picker row ──────────────────────────────────
        row1 = tk.Frame(self)
        row1.pack(fill="x", **pad)

        self._path_var = tk.StringVar()
        tk.Entry(row1, textvariable=self._path_var, width=45,
                 state="readonly").pack(side="left", fill="x", expand=True)
        tk.Button(row1, text="Select File", command=self._pick).pack(side="left", padx=(8, 0))

        # ── mode label ───────────────────────────────────────
        self._mode_var = tk.StringVar(value="")
        tk.Label(self, textvariable=self._mode_var, fg="#0055cc",
                 anchor="w").pack(fill="x", padx=16)

        # ── convert button ───────────────────────────────────
        self._btn = tk.Button(self, text="Convert", width=20,
                              command=self._start, state="disabled")
        self._btn.pack(**pad)

        # ── status bar ───────────────────────────────────────
        self._status = tk.StringVar(
            value="Select a file to convert  —  Supports: PDF, Word (.docx), JPG, PNG"
        )
        tk.Label(self, textvariable=self._status, anchor="w",
                 fg="#555555").pack(fill="x", padx=16, pady=(0, 12))

    def _pick(self):
        path = filedialog.askopenfilename(
            title="Select a file to convert",
            filetypes=FILETYPES,
        )
        if path:
            p = Path(path)
            self._path_var.set(path)
            self._mode_var.set("Mode: " + _label_for(p))
            self._status.set("Selected: " + p.name)
            self._btn.config(state="normal")

    def _start(self):
        src = self._path_var.get()
        if not src:
            return
        self._btn.config(state="disabled")
        self._status.set("Converting, please wait…")
        self.update_idletasks()
        threading.Thread(target=self._run, args=(src,), daemon=True).start()

    def _run(self, src: str):
        try:
            docx_bytes = _convert(src)
            out_path = Path(src).with_suffix(".docx")
            if out_path == Path(src):
                out_path = Path(src).with_stem(Path(src).stem + "_table")
            out_path.write_bytes(docx_bytes)
            self.after(0, self._done, str(out_path))
        except Exception as exc:
            self.after(0, self._error, str(exc))

    def _done(self, out_path: str):
        self._status.set("Done! Saved to: " + out_path)
        self._btn.config(state="normal")
        messagebox.showinfo("Conversion complete", f"Saved to:\n{out_path}")

    def _error(self, msg: str):
        self._status.set("Error: " + msg)
        self._btn.config(state="normal")
        messagebox.showerror("Conversion failed", msg)


if __name__ == "__main__":
    App().mainloop()
