import json
import os
import shutil
import subprocess
import sys
import tempfile
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk


def _ensure_ai_vision_deps():
    """
    Ensure packages needed for AI vision (anthropic, boto3) are available.
    In a PyInstaller bundle they must be compiled in — pip install won't work.
    In a normal Python environment, install silently if missing.
    Returns (ok: bool, missing: list[str]).
    """
    needed = []
    _import_errors = {}
    try:
        import anthropic  # noqa
    except Exception as _e:  # catch all — frozen exe may raise OSError/AttributeError
        needed.append("anthropic>=1.0.0")
        _import_errors["anthropic"] = f"{type(_e).__name__}: {_e}"
    try:
        import boto3  # noqa
    except Exception as _e:
        needed.append("boto3")
        _import_errors["boto3"] = f"{type(_e).__name__}: {_e}"

    if not needed:
        return True, []

    # PyInstaller bundle: cannot pip install at runtime.
    # Write error to log file for diagnosis, then fall back silently.
    if getattr(sys, "frozen", False):
        try:
            _log_path = Path.home() / ".pdf2word" / "ai_vision_error.log"
            _log_path.parent.mkdir(parents=True, exist_ok=True)
            _log_path.write_text(
                "\n".join(f"{k}: {v}" for k, v in _import_errors.items()),
                encoding="utf-8",
            )
        except Exception:
            pass
        err_detail = "; ".join(f"{k}: {v}" for k, v in _import_errors.items())
        return False, [f"{pkg} ({err_detail})" if err_detail else pkg for pkg in needed]

    try:
        subprocess.run(
            [sys.executable, "-m", "pip", "install", "--quiet"] + needed,
            check=True, capture_output=True,
        )
        return True, needed
    except Exception:
        return False, needed

from converter import (
    IMAGE_EXTS,
    convert_image_to_word,
    convert_pdf_to_word,
    convert_word_to_word,
    is_scanned_pdf,
    make_searchable_pdf,
)

FILETYPES = [
    ("Supported formats", "*.pdf *.docx *.jpg *.jpeg *.png *.bmp *.tiff *.tif *.webp"),
    ("PDF files",         "*.pdf"),
    ("Word files",        "*.docx"),
    ("Image files",       "*.jpg *.jpeg *.png *.bmp *.tiff *.tif *.webp"),
]

# ── AI 設定常數 ────────────────────────────────────────────────────────────────
_CONF_DIR       = Path.home() / ".pdf2word"
_AI_CONFIG_PATH = _CONF_DIR / "ai_config.json"
_AI_PROMPT_PATH = _CONF_DIR / "ai_prompt_override.txt"

_AI_PROMPT_DEFAULT = """\
請分析以下 Word 文件的內容（由 PDF 轉換而來），用繁體中文回答以下兩個問題：

1. **文法與用詞檢查**
   找出文法錯誤、不自然的句子、標點問題或用詞不一致之處，並條列每個問題與建議修正。

2. **Bookmark 建議**
   找出適合新增 Word Bookmark 的位置，例如主要章節標題、常被引用的定義、重要條款等。
   請列出：
   - Bookmark 名稱（英文或拼音，不含空格）
   - 對應的原文內容（前 20 字）"""

_PRESET_MODELS = [
    "",
    "claude-sonnet-5",
    "claude-opus-5",
    "claude-haiku-4-5-20251001",
    "us.anthropic.claude-sonnet-4-5-20250929-v1:0",
    "us.anthropic.claude-haiku-4-5-20251001-v1:0",
]

_AI_FIX_PROMPT = """\
你是文件校對助手。以下是由 PDF 轉換而來的 Word 文件內容，每一行對應文件中的一個段落或欄位。請找出並修正：

1. **拼字錯誤**：明顯的錯字、多字、少字、OCR 辨識錯誤（如「0」vs「O」、「l」vs「1」、「rn」vs「m」）
2. **標點符號問題**：多餘或缺少的標點符號
3. **文字黏連**：換行錯誤導致的詞語黏在一起

請直接輸出 JSON，不要有任何其他說明文字，格式如下：
{
  "corrections": [
    {
      "original": "原始錯誤文字（必須是單一段落內的原文，不可跨行，保留原始空格）",
      "corrected": "修正後文字",
      "reason": "修正原因"
    }
  ],
  "summary": "修正項目總數與簡短說明"
}

注意：
- original 必須是文件中單一行（段落）內實際存在的原文，不可包含換行符
- 保留原文的空格格式，不要自行新增或刪除空格
- 只列出確定錯誤的項目，不確定的不要列
- 若無需修正，corrections 設為空陣列 []"""


def _extract_docx_text(docx_bytes: bytes) -> str:
    """
    從 docx bytes 提取純文字，供 AI 分析。
    每個段落獨立一行，去重合併儲存格，並正規化多餘空白。
    """
    import io as _io, re as _re
    from docx import Document as _Document
    doc = _Document(_io.BytesIO(docx_bytes))
    parts = []

    def _add(text: str):
        t = _re.sub(r'\s+', ' ', text).strip()
        if t:
            parts.append(t)

    for para in doc.paragraphs:
        _add(para.text)

    seen_cells: set = set()
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                cid = id(cell._tc)
                if cid in seen_cells:
                    continue
                seen_cells.add(cid)
                for para in cell.paragraphs:
                    _add(para.text)

    return "\n".join(parts)


def _apply_text_corrections(docx_bytes: bytes, correction_list: list) -> tuple:
    """
    套用 AI JSON 修正至 docx。
    回傳 (corrected_bytes, applied_count, not_found_originals)。
    以 regex 彈性匹配多餘空白（Word run 每個字後加空格造成）。
    """
    import io as _io, re as _re
    from docx import Document as _Document

    doc = _Document(_io.BytesIO(docx_bytes))
    applied = 0
    not_found: list = []

    def _norm(s: str) -> str:
        return _re.sub(r'\s+', ' ', s).strip()

    def _build_pattern(norm_orig: str) -> str:
        return r'\s+'.join(_re.escape(w) for w in norm_orig.split())

    def _fix_para(para, norm_orig: str, corr: str, pattern: str) -> bool:
        nonlocal applied
        if _norm(para.text) and norm_orig not in _norm(para.text):
            return False
        # 先嘗試單一 run 匹配
        for run in para.runs:
            if _re.search(pattern, run.text):
                run.text = _re.sub(pattern, corr, run.text, count=1)
                applied += 1
                return True
        # 跨 run 情況：折疊至第一個 run
        new_text = _re.sub(pattern, corr, para.text, count=1)
        if new_text != para.text and para.runs:
            para.runs[0].text = new_text
            for r in para.runs[1:]:
                r.text = ""
            applied += 1
            return True
        return False

    seen_cells: set = set()
    for entry in correction_list:
        orig = entry.get("original", "").strip()
        repl = entry.get("corrected", "").strip()
        if not orig or orig == repl:
            continue
        norm_orig = _norm(orig)
        if not norm_orig:
            continue
        pattern = _build_pattern(norm_orig)
        found = False

        for para in doc.paragraphs:
            if _fix_para(para, norm_orig, repl, pattern):
                found = True
                break

        if not found:
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        cid = id(cell._tc)
                        if cid in seen_cells:
                            continue
                        seen_cells.add(cid)
                        for para in cell.paragraphs:
                            if _fix_para(para, norm_orig, repl, pattern):
                                found = True
                                break
                        if found:
                            break
                    if found:
                        break
                if found:
                    break
            seen_cells.clear()

        if not found:
            not_found.append(orig)

    buf = _io.BytesIO()
    doc.save(buf)
    return buf.getvalue(), applied, not_found


def _label_for(path: Path) -> str:
    ext = path.suffix.lower()
    if ext == ".pdf":
        return "PDF → Word"
    if ext == ".docx":
        return "Word → Word (table layout)"
    if ext in IMAGE_EXTS:
        return "Image → Word (OCR)"
    return "Unknown format"


def _convert(path: str, ai_vision_model: str = "", aws_profile: str = "",
             progress_callback=None) -> bytes:
    ext = Path(path).suffix.lower()
    if ext == ".pdf":
        return convert_pdf_to_word(path,
                                   ai_vision_model=ai_vision_model or None,
                                   aws_profile=aws_profile or None,
                                   progress_callback=progress_callback)
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
        self._ai_model_var        = tk.StringVar(value="claude-haiku-4-5-20251001")
        self._ai_vision_model_var = tk.StringVar(value="claude-sonnet-5")
        self._aws_profile_var     = tk.StringVar(value="genai-dev")
        self._auto_fix_var        = tk.BooleanVar(value=True)
        self._cancel_flag         = threading.Event()
        self._last_docx_path = None
        self._load_ai_config()
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

        # ── convert + AI 設定 row ────────────────────────────
        row_btn = tk.Frame(self)
        row_btn.pack(**pad)

        self._btn = tk.Button(row_btn, text="Convert", width=20,
                              command=self._start, state="disabled")
        self._btn.pack(side="left")
        ttk.Checkbutton(row_btn, text="轉換後自動修正拼字",
                        variable=self._auto_fix_var,
                        command=self._on_auto_fix_toggle).pack(side="left", padx=(10, 0))
        # Fix model button
        self._fix_model_btn = tk.Button(
            row_btn, text=self._ai_model_var.get(),
            font=("TkDefaultFont", 8), fg="#0055cc", relief="flat",
            cursor="hand2", command=self._pick_fix_model,
        )
        self._fix_model_btn.pack(side="left", padx=(4, 0))

        # ── AI Vision row ─────────────────────────────────────
        row_vision = tk.Frame(self)
        row_vision.pack(padx=16, pady=(0, 4), anchor="w")
        tk.Label(row_vision, text="AI Vision：", fg="#555555").pack(side="left")
        self._vision_model_btn = tk.Button(
            row_vision, text=self._ai_vision_model_var.get() or "（停用）",
            font=("TkDefaultFont", 8), fg="#0055cc", relief="flat",
            cursor="hand2", command=self._pick_vision_model,
        )
        self._vision_model_btn.pack(side="left")

        # ── status bar ───────────────────────────────────────
        self._status = tk.StringVar(
            value="Select a file to convert  —  Supports: PDF, Word (.docx), JPG, PNG"
        )
        tk.Label(self, textvariable=self._status, anchor="w",
                 fg="#555555").pack(fill="x", padx=16, pady=(0, 4))

        # ── separator ────────────────────────────────────────
        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=16, pady=(0, 8))

        # ── AI 分析 row ──────────────────────────────────────
        row_ai = tk.Frame(self)
        row_ai.pack(padx=16, pady=(0, 12))

        self._ai_btn = tk.Button(row_ai, text="🤖 AI 分析文件", width=18,
                                 command=self._start_ai_analyze, state="disabled")
        self._ai_btn.pack(side="left")

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
        self._cancel_flag.clear()
        self._btn.config(text="取消", command=self._cancel, state="normal")
        self._ai_btn.config(state="disabled")
        self._status.set("Converting, please wait…")
        self.update_idletasks()
        threading.Thread(target=self._run, args=(src,), daemon=True).start()

    def _cancel(self):
        self._cancel_flag.set()
        self._status.set("取消中…")
        self._btn.config(state="disabled")

    def _run(self, src: str):
        try:
            src_path = Path(src)

            # ── 掃描版 PDF：顯示警告後自動嘗試建 OCR 文字層 ────────────────────
            if src_path.suffix.lower() == ".pdf" and is_scanned_pdf(src):
                choice_evt  = threading.Event()
                user_choice = ["cancel"]  # "proceed" | "cancel"

                def _show_warning():
                    self._show_scanned_warning(
                        src,
                        on_choice=lambda c: (
                            user_choice.__setitem__(0, c),
                            choice_evt.set(),
                        ),
                    )

                self.after(0, _show_warning)
                choice_evt.wait()

                if user_choice[0] == "cancel" or self._cancel_flag.is_set():
                    self.after(0, lambda: (self._restore_convert_btn(), self._status.set("已取消")))
                    return

                # 使用者確認後，嘗試建立 OCR 文字層（專為掃描 PDF 設計，應該成功）
                self.after(0, lambda: self._status.set("OCR 辨識中，建立可搜尋文字層…"))
                searchable_path = src_path.with_name(src_path.stem + "_searchable.pdf")
                try:
                    make_searchable_pdf(src, str(searchable_path))
                    actual_src = str(searchable_path)
                    self.after(0, lambda: self._status.set("OCR 完成，開始轉換…"))
                except Exception as ocr_exc:
                    # OCR 失敗（罕見）→ 靜默降級，直接用原始 PDF
                    actual_src = src
            else:
                actual_src = src

            if self._cancel_flag.is_set():
                self.after(0, lambda: (self._restore_convert_btn(), self._status.set("已取消")))
                return

            # 啟用 AI vision 時，確保相依套件已安裝
            _vision_model = self._ai_vision_model_var.get().strip()
            if _vision_model:
                self.after(0, lambda: self._status.set("檢查 AI 套件…"))
                _ok, _missing = _ensure_ai_vision_deps()
                if not _ok:
                    _msg = f"AI vision 所需套件安裝失敗：{', '.join(_missing)}\n將改用 OCR 模式轉換。"
                    self.after(0, lambda m=_msg: messagebox.showwarning("套件安裝失敗", m))
                    _vision_model = ""
                elif _missing:
                    self.after(0, lambda: self._status.set("已安裝缺少的套件，開始轉換…"))

            self.after(0, lambda: self._status.set("轉換中… (第 1 頁)"))

            def _on_page_progress(page_idx, num_pages, msg):
                label = f"轉換中… ({page_idx+1}/{num_pages} 頁) {msg.split(':',1)[-1].strip()[:40]}"
                self.after(0, lambda m=label: self._status.set(m))

            docx_bytes = _convert(actual_src,
                                   ai_vision_model=_vision_model,
                                   aws_profile=self._aws_profile_var.get().strip(),
                                   progress_callback=_on_page_progress)
            out_path = src_path.with_suffix(".docx")
            if out_path == src_path:
                out_path = src_path.with_stem(src_path.stem + "_table")
            out_path.write_bytes(docx_bytes)
            self._last_docx_path = str(out_path)
            self.after(0, self._done, str(out_path))
        except Exception as exc:
            self.after(0, self._error, str(exc))

    def _show_scanned_warning(self, pdf_path: str, on_choice):
        """
        偵測到掃描版 PDF 時告知使用者品質限制並確認是否繼續。
        確認後程式會自動嘗試建立 OCR 文字層（ocrmypdf）再轉換，
        OCR 失敗時自動降級直接轉換。
        on_choice("proceed") → 繼續
        on_choice("cancel")  → 取消
        """
        has_ai = bool(self._ai_vision_model_var.get().strip())

        if has_ai:
            quality_note = (
                "Claude Vision 將分析版面，\n"
                "可還原背景色、表格結構與粗體文字。\n"
                "Logo / 印章圖片仍可能與原始版面略有差異。"
            )
        else:
            quality_note = (
                "轉換結果可能與原始版面差異較大：\n"
                "  • 字型、粗體、斜體 — 無法自動識別\n"
                "  • 印章 / Logo 圖片 — 不會嵌入\n"
                "  • 表格欄寬、對齊   — 僅為估算\n\n"
                "建議在「⚙ AI 設定」選擇 claude-sonnet-5\n"
                "以獲得接近原始版面的輸出品質。"
            )

        msg = (
            "此 PDF 是圖片掃描格式（非文字版）。\n\n"
            + quality_note +
            "\n\n程式將先用 OCR 建立文字層再轉換。\n"
            "是否繼續？"
        )

        proceed = messagebox.askyesno(
            "偵測到掃描版 PDF", msg, icon="warning"
        )
        on_choice("proceed" if proceed else "cancel")

    def _show_searchable_verify(self, searchable_path: str, on_confirm):
        """顯示驗證對話框：讓使用者開啟 searchable PDF 確認 OCR 文字後決定是否繼續。"""
        dlg = tk.Toplevel(self)
        dlg.title("驗證 OCR 文字")
        dlg.resizable(True, False)
        dlg.minsize(400, 200)
        dlg.geometry("460x240")
        dlg.grab_set()
        dlg.columnconfigure(0, weight=1)

        tk.Label(
            dlg,
            text=("已建立含 OCR 文字層的 PDF：\n"
                  f"{Path(searchable_path).name}\n\n"
                  "請開啟確認文字辨識正確後，\n"
                  "再選擇是否繼續轉換為 Word。"),
            wraplength=380, justify="left", padx=16, pady=12,
        ).grid(row=0, column=0, columnspan=2, sticky="ew")

        def _open():
            try:
                os.startfile(searchable_path)
            except Exception:
                pass

        ttk.Button(dlg, text="開啟 PDF 驗證", command=_open, width=16).grid(
            row=1, column=0, columnspan=2, padx=16, pady=(0, 8))

        btn_f = tk.Frame(dlg)
        btn_f.grid(row=2, column=0, columnspan=2, padx=16, pady=(0, 12), sticky="e")

        def _confirm():
            dlg.destroy()
            on_confirm(True)

        def _cancel():
            dlg.destroy()
            on_confirm(False)

        ttk.Button(btn_f, text="繼續轉 Word", command=_confirm, width=12).pack(
            side="right", padx=(4, 0))
        ttk.Button(btn_f, text="取消", command=_cancel, width=8).pack(side="right")

        dlg.protocol("WM_DELETE_WINDOW", _cancel)

    def _show_save_result(self, title: str, message: str, out_path: str):
        """顯示儲存結果對話框，並自動開啟資料夾（Explorer 選取該檔案）。"""
        # 自動開啟 Explorer 並選取檔案
        try:
            import subprocess as _sp
            _sp.Popen(["explorer", "/select,", str(out_path).replace("/", "\\")])
        except Exception:
            pass

        dlg = tk.Toplevel(self)
        dlg.title(title)
        dlg.resizable(False, False)
        dlg.grab_set()
        dlg.columnconfigure(0, weight=1)

        tk.Label(dlg, text=message, wraplength=420, justify="left",
                 padx=16, pady=12).grid(row=0, column=0, sticky="ew")

        btn_f = tk.Frame(dlg)
        btn_f.grid(row=1, column=0, padx=16, pady=(0, 12), sticky="e")

        def _open_file():
            try:
                os.startfile(out_path)
            except Exception:
                pass
            dlg.destroy()

        tk.Button(btn_f, text="開啟檔案", width=12, command=_open_file).pack(
            side="left", padx=(0, 4))
        tk.Button(btn_f, text="關閉", width=8, command=dlg.destroy).pack(side="left")

        dlg.after(100, lambda: dlg.lift())

    def _pick_vision_model(self):
        """Inline model picker for AI vision conversion."""
        dlg = tk.Toplevel(self)
        dlg.title("AI Vision Model")
        dlg.resizable(False, False)
        dlg.transient(self)
        dlg.grab_set()

        tk.Label(dlg, text="選擇掃描 PDF 轉換使用的 AI Vision Model：",
                 anchor="w").pack(padx=12, pady=(10, 4), fill="x")
        tk.Label(dlg, text="（留空 = 停用 AI vision，使用 OCR）",
                 fg="#888888", font=("TkDefaultFont", 8), anchor="w"
                 ).pack(padx=12, pady=(0, 6), fill="x")

        var = tk.StringVar(value=self._ai_vision_model_var.get())
        ttk.Radiobutton(dlg, text="（停用）", variable=var, value="").pack(
            anchor="w", padx=20, pady=1)
        for m in _PRESET_MODELS[1:]:
            ttk.Radiobutton(dlg, text=m, variable=var, value=m).pack(
                anchor="w", padx=20, pady=1)

        custom_frame = tk.Frame(dlg)
        custom_frame.pack(fill="x", padx=12, pady=(4, 0))
        ttk.Radiobutton(custom_frame, text="自訂：", variable=var, value="__custom__"
                        ).pack(side="left")
        custom_var = tk.StringVar()
        tk.Entry(custom_frame, textvariable=custom_var, width=30).pack(side="left")

        # AWS Profile
        pf = tk.Frame(dlg)
        pf.pack(fill="x", padx=12, pady=(8, 0))
        tk.Label(pf, text="AWS Profile（留空用 default）：", anchor="w").pack(side="left")
        profile_var = tk.StringVar(value=self._aws_profile_var.get())
        tk.Entry(pf, textvariable=profile_var, width=20).pack(side="left", padx=(4, 0))

        def _confirm():
            val = var.get()
            if val == "__custom__":
                val = custom_var.get().strip()
            self._ai_vision_model_var.set(val)
            self._vision_model_btn.config(text=val or "（停用）")
            self._aws_profile_var.set(profile_var.get().strip())
            self._save_ai_config()
            dlg.destroy()

        tk.Frame(dlg).pack(pady=4)
        bf = tk.Frame(dlg)
        bf.pack(pady=(0, 10))
        tk.Button(bf, text="確定", width=8, command=_confirm).pack(side="left", padx=4)
        tk.Button(bf, text="取消", width=8, command=dlg.destroy).pack(side="left")
        dlg.after(50, lambda: dlg.lift())

    def _pick_fix_model(self):
        """Inline model picker for auto-fix — small popup next to the button."""
        dlg = tk.Toplevel(self)
        dlg.title("修正拼字 Model")
        dlg.resizable(False, False)
        dlg.transient(self)
        dlg.grab_set()

        tk.Label(dlg, text="選擇修正拼字使用的 Model：",
                 anchor="w").pack(padx=12, pady=(10, 4), fill="x")

        var = tk.StringVar(value=self._ai_model_var.get())
        for m in _PRESET_MODELS[1:]:  # skip the empty default
            ttk.Radiobutton(dlg, text=m, variable=var, value=m).pack(
                anchor="w", padx=20, pady=1)

        # custom entry
        custom_frame = tk.Frame(dlg)
        custom_frame.pack(fill="x", padx=12, pady=(4, 0))
        ttk.Radiobutton(custom_frame, text="自訂：", variable=var, value="__custom__"
                        ).pack(side="left")
        custom_var = tk.StringVar()
        custom_entry = tk.Entry(custom_frame, textvariable=custom_var, width=25)
        custom_entry.pack(side="left")

        def _confirm():
            val = var.get()
            if val == "__custom__":
                val = custom_var.get().strip()
            if val:
                self._ai_model_var.set(val)
                self._fix_model_btn.config(text=val)
                self._save_ai_config()
            dlg.destroy()

        tk.Frame(dlg).pack(pady=4)
        bf = tk.Frame(dlg)
        bf.pack(pady=(0, 10))
        tk.Button(bf, text="確定", width=8, command=_confirm).pack(side="left", padx=4)
        tk.Button(bf, text="取消", width=8, command=dlg.destroy).pack(side="left")
        dlg.after(50, lambda: dlg.lift())

    def _on_auto_fix_toggle(self):
        """勾選「自動修正拼字」時，立即開啟 AI 設定並驗證連線。"""
        if self._auto_fix_var.get():
            self._open_ai_settings()

    def _restore_convert_btn(self):
        """轉換結束（完成/錯誤/取消）後恢復 Convert 按鈕。"""
        self._cancel_flag.clear()
        self._btn.config(text="Convert", command=self._start, state="normal")
        self._ai_btn.config(state="normal")

    def _done(self, out_path: str):
        if self._auto_fix_var.get() and not self._cancel_flag.is_set():
            # Keep "取消" button active during auto-fix so user can cancel
            self._status.set("轉換完成，AI 修正拼字中…")
            self._run_auto_fix_silent(out_path)
            return
        self._restore_convert_btn()
        self._status.set("Done! Saved to: " + out_path)
        self._show_save_result(
            "Conversion complete",
            f"轉換完成，已儲存至：\n{out_path}",
            out_path,
        )

    def _run_auto_fix_silent(self, base_path: str):
        """轉換後自動執行 AI 拼字修正（靜默模式，不開 popup）。"""
        def _finish_autofix(status_msg: str = ""):
            self._restore_convert_btn()
            if status_msg:
                self._status.set(status_msg)

        try:
            doc_bytes = Path(base_path).read_bytes()
        except Exception as e:
            _finish_autofix(f"轉換完成，讀取失敗：{e}")
            return
        doc_text = _extract_docx_text(doc_bytes)
        if not doc_text.strip():
            _finish_autofix(f"轉換完成（無文字可修正）：{base_path}")
            return

        # Check cancel before starting async
        if self._cancel_flag.is_set():
            _finish_autofix("已取消")
            return

        full_prompt = f"{_AI_FIX_PROMPT}\n\n=== 文件內容 ===\n{doc_text[:8000]}"

        def on_fix_done(raw_text):
            if self._cancel_flag.is_set():
                _finish_autofix("已取消")
                return
            import json as _json, re as _re
            json_match = _re.search(r'\{[\s\S]*\}', raw_text)
            if not json_match:
                _finish_autofix(f"轉換完成（AI 修正回應無效）：{base_path}")
                return
            try:
                corrections = _json.loads(json_match.group()).get("corrections", [])
            except Exception:
                _finish_autofix(f"轉換完成（AI 修正 JSON 解析失敗）：{base_path}")
                return
            if not corrections:
                _finish_autofix(f"轉換完成（AI 未發現拼字錯誤）：{base_path}")
                return
            try:
                corrected_bytes, applied, _ = _apply_text_corrections(doc_bytes, corrections)
            except Exception as e:
                _finish_autofix(f"轉換完成（修正套用失敗：{e}）：{base_path}")
                return
            orig_path = Path(base_path)
            out_path = orig_path.with_stem(orig_path.stem + "_corrected")
            try:
                out_path.write_bytes(corrected_bytes)
                self._last_docx_path = str(out_path)
                _finish_autofix(f"完成！已修正 {applied}/{len(corrections)} 項 → {out_path.name}")
                self._show_save_result(
                    "轉換並修正完成",
                    f"原始檔：{orig_path.name}\n修正版：{out_path.name}\n\n套用 {applied}/{len(corrections)} 項拼字修正",
                    str(out_path),
                )
            except Exception as e:
                _finish_autofix(f"修正版儲存失敗：{e}")

        def on_fix_error(text):
            _finish_autofix(f"轉換完成（AI 修正失敗）：{base_path}")

        self._run_claude_async(full_prompt, on_fix_done, on_fix_error)

    def _error(self, msg: str):
        self._restore_convert_btn()
        self._status.set("Error: " + msg)
        messagebox.showerror("Conversion failed", msg)

    # ── AI 設定 ────────────────────────────────────────────────────────────────

    def _load_ai_config(self):
        if _AI_CONFIG_PATH.exists():
            try:
                cfg = json.loads(_AI_CONFIG_PATH.read_text(encoding="utf-8"))
                self._ai_model_var.set(cfg.get("claude_model", "claude-haiku-4-5-20251001"))
                self._ai_vision_model_var.set(cfg.get("claude_vision_model", "claude-sonnet-5"))
                self._aws_profile_var.set(cfg.get("aws_profile", "genai-dev"))
            except Exception:
                pass
        # Sync model button texts after loading config (buttons may not exist yet on first call)
        try:
            self._fix_model_btn.config(text=self._ai_model_var.get())
        except AttributeError:
            pass
        try:
            self._vision_model_btn.config(
                text=self._ai_vision_model_var.get() or "（停用）")
        except AttributeError:
            pass

    def _save_ai_config(self):
        _CONF_DIR.mkdir(parents=True, exist_ok=True)
        cfg = {
            "claude_model":        self._ai_model_var.get().strip(),
            "claude_vision_model": self._ai_vision_model_var.get().strip(),
            "aws_profile":         self._aws_profile_var.get().strip(),
        }
        _AI_CONFIG_PATH.write_text(
            json.dumps(cfg, indent=2, ensure_ascii=False), encoding="utf-8"
        )

    def _open_ai_settings(self):
        dlg = tk.Toplevel(self)
        dlg.title("AI 設定")
        dlg.resizable(False, False)
        dlg.grab_set()
        dlg.columnconfigure(0, weight=1)

        ttk.Label(dlg, text="Claude Model：").grid(
            row=0, column=0, sticky="w", padx=10, pady=(12, 2))

        model_f = ttk.Frame(dlg)
        model_f.grid(row=1, column=0, padx=10, pady=(0, 2), sticky="ew")
        model_f.columnconfigure(0, weight=1)

        _model_var = tk.StringVar(value=self._ai_model_var.get())
        model_cb = ttk.Combobox(model_f, textvariable=_model_var,
                                values=_PRESET_MODELS, width=48)
        model_cb.grid(row=0, column=0, sticky="ew")

        verify_lbl = ttk.Label(model_f, text="", foreground="gray", width=18)
        verify_lbl.grid(row=0, column=1, padx=(6, 0), sticky="w")

        def _verify_model():
            model = _model_var.get().strip()
            verify_lbl.config(text="驗證中…", foreground="gray")
            dlg.update_idletasks()

            def _worker():
                try:
                    exe = shutil.which("claude") or "claude"
                    args = [exe, "-p", "--dangerously-skip-permissions"]
                    if model:
                        args += ["--model", model]
                    args.append("hi")
                    r = subprocess.run(args, capture_output=True, text=True,
                                       encoding="utf-8", errors="replace", timeout=30)
                    ok = r.returncode == 0
                    msg, color = ("✔ 可用", "#007700") if ok else ("✘ 無效", "#cc0000")
                except Exception:
                    msg, color = "✘ 錯誤", "#cc0000"
                dlg.after(0, lambda m=msg, c=color: verify_lbl.config(text=m, foreground=c))

            threading.Thread(target=_worker, daemon=True).start()

        ttk.Button(model_f, text="驗證", command=_verify_model, width=6).grid(
            row=0, column=2, padx=(4, 0))

        ttk.Label(dlg, text="留空 = 使用 Claude CLI 預設 model",
                  foreground="gray").grid(row=2, column=0, sticky="w", padx=10, pady=(0, 10))

        ttk.Separator(dlg, orient="horizontal").grid(
            row=3, column=0, sticky="ew", padx=10, pady=(0, 8))

        ttk.Label(dlg, text="掃描版 PDF 轉換 Model（建議 claude-sonnet-5）：").grid(
            row=4, column=0, sticky="w", padx=10, pady=(0, 2))

        vision_f = ttk.Frame(dlg)
        vision_f.grid(row=5, column=0, padx=10, pady=(0, 2), sticky="ew")
        vision_f.columnconfigure(0, weight=1)

        _vision_var = tk.StringVar(value=self._ai_vision_model_var.get())
        vision_cb = ttk.Combobox(vision_f, textvariable=_vision_var,
                                 values=_PRESET_MODELS, width=48)
        vision_cb.grid(row=0, column=0, sticky="ew")

        ttk.Label(dlg,
                  text="用於掃描版 PDF 的視覺分析（還原粗體、對齊、背景色、表格結構）。\n"
                       "建議使用 Sonnet — 多模態能力較 Haiku 強，版面還原準確率更高。\n"
                       "留空則改用本機 OCR（Tesseract）。",
                  foreground="gray").grid(row=6, column=0, sticky="w", padx=10, pady=(0, 10))

        btn_f = ttk.Frame(dlg)
        btn_f.grid(row=7, column=0, padx=10, pady=(4, 12), sticky="e")

        def save_close():
            self._ai_model_var.set(_model_var.get().strip())
            self._ai_vision_model_var.set(_vision_var.get().strip())
            self._save_ai_config()
            dlg.destroy()

        ttk.Button(btn_f, text="儲存", command=save_close, width=8).pack(side="right", padx=(4, 0))
        ttk.Button(btn_f, text="取消", command=dlg.destroy, width=8).pack(side="right")

    def _open_prompt_editor(self):
        ed = tk.Toplevel(self)
        ed.title("編輯 AI 分析 Prompt")
        ed.geometry("660x420")
        ed.grab_set()
        ed.rowconfigure(1, weight=1)
        ed.columnconfigure(0, weight=1)

        is_ov = _AI_PROMPT_PATH.exists()
        lbl = ttk.Label(
            ed,
            text="（使用自訂版本）" if is_ov else "（使用預設版本）",
            foreground="#886600" if is_ov else "gray",
        )
        lbl.grid(row=0, column=0, sticky="w", padx=8, pady=(8, 2))

        frm = ttk.Frame(ed)
        frm.grid(row=1, column=0, sticky="nsew", padx=8, pady=(0, 4))
        frm.rowconfigure(0, weight=1)
        frm.columnconfigure(0, weight=1)
        txt = tk.Text(frm, wrap="word", font=("Consolas", 10), undo=True)
        sb  = ttk.Scrollbar(frm, orient="vertical", command=txt.yview)
        txt.configure(yscrollcommand=sb.set)
        txt.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        txt.insert("1.0", _AI_PROMPT_PATH.read_text(encoding="utf-8") if is_ov else _AI_PROMPT_DEFAULT)

        btn_row = ttk.Frame(ed)
        btn_row.grid(row=2, column=0, padx=8, pady=(0, 8), sticky="ew")

        def _save():
            new_text = txt.get("1.0", "end-1c").strip()
            if not new_text:
                messagebox.showwarning("內容為空", "Prompt 不可為空。", parent=ed)
                return
            _CONF_DIR.mkdir(parents=True, exist_ok=True)
            _AI_PROMPT_PATH.write_text(new_text, encoding="utf-8")
            lbl.config(text="（使用自訂版本）", foreground="#886600")
            messagebox.showinfo("已儲存", "Prompt 已儲存。下次分析時生效。", parent=ed)

        def _reset():
            if not messagebox.askyesno("重設為預設", "確定要放棄自訂內容，恢復預設 Prompt？", parent=ed):
                return
            if _AI_PROMPT_PATH.exists():
                _AI_PROMPT_PATH.unlink()
            txt.delete("1.0", "end")
            txt.insert("1.0", _AI_PROMPT_DEFAULT)
            lbl.config(text="（使用預設版本）", foreground="gray")

        ttk.Button(btn_row, text="儲存", command=_save, width=10).pack(side="left")
        ttk.Button(btn_row, text="重設為預設", command=_reset, width=12).pack(side="left", padx=(8, 0))
        ttk.Button(btn_row, text="關閉", command=ed.destroy, width=10).pack(side="right")

    # ── AI 分析 ────────────────────────────────────────────────────────────────

    def _start_ai_analyze(self):
        if not self._last_docx_path:
            return

        popup = tk.Toplevel(self)
        popup.title("AI 文件分析")
        popup.geometry("700x520")
        popup.minsize(500, 360)
        popup.columnconfigure(0, weight=1)
        popup.rowconfigure(2, weight=1)

        # ── button row ───────────────────────────────
        hdr = tk.Frame(popup)
        hdr.grid(row=0, column=0, sticky="ew", padx=8, pady=(8, 4))

        analyze_btn = tk.Button(hdr, text="🤖 開始分析", width=14)
        analyze_btn.pack(side="left")
        tk.Button(hdr, text="✏ Prompt", width=10,
                  command=self._open_prompt_editor).pack(side="left", padx=(6, 0))
        tk.Button(hdr, text="⚙ AI 設定", width=10,
                  command=self._open_ai_settings).pack(side="left", padx=(6, 0))

        ttk.Separator(popup, orient="horizontal").grid(row=1, column=0, sticky="ew")

        # ── result text area ─────────────────────────
        result_frame = tk.Frame(popup)
        result_frame.grid(row=2, column=0, sticky="nsew", padx=8, pady=(4, 8))
        result_frame.rowconfigure(0, weight=1)
        result_frame.columnconfigure(0, weight=1)

        result_txt = tk.Text(result_frame, wrap="word", font=("Segoe UI", 10),
                             bg="#0d1117", fg="#c9d1d9")
        sb = ttk.Scrollbar(result_frame, command=result_txt.yview)
        result_txt.configure(yscrollcommand=sb.set)
        result_txt.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")

        def _block_edit(event):
            if event.state & 0x4:
                return None
            if event.keysym in (
                "Up", "Down", "Left", "Right", "Home", "End", "Prior", "Next",
                "Shift_L", "Shift_R", "Control_L", "Control_R",
            ):
                return None
            return "break"
        result_txt.bind("<Key>", _block_edit)

        _SPIN    = ["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"]
        _spin_i   = [0]
        _spin_job = [None]

        def _spin_tick():
            _spin_i[0] = (_spin_i[0] + 1) % len(_SPIN)
            result_txt.delete("1.0", "end")
            result_txt.insert("end", f"分析中… {_SPIN[_spin_i[0]]}\n")
            _spin_job[0] = result_txt.after(120, _spin_tick)

        def _stop_spinner():
            if _spin_job[0]:
                result_txt.after_cancel(_spin_job[0])
                _spin_job[0] = None

        _streamed = [False]

        def on_chunk(text, is_first):
            _streamed[0] = True
            if is_first:
                _stop_spinner()
                result_txt.delete("1.0", "end")
            result_txt.insert("end", text)
            result_txt.see("end")

        def on_result(text):
            _stop_spinner()
            if not _streamed[0]:
                result_txt.delete("1.0", "end")
                result_txt.insert("end", text)
            result_txt.see("1.0")
            analyze_btn.config(state="normal", text="🤖 開始分析")

        def on_error(text):
            _stop_spinner()
            result_txt.delete("1.0", "end")
            result_txt.insert("end", text)
            analyze_btn.config(state="normal", text="🤖 開始分析")

        def do_analyze():
            if not self._last_docx_path:
                on_error("❌ 尚未轉換任何文件，請先完成轉換。")
                return
            try:
                doc_bytes = Path(self._last_docx_path).read_bytes()
            except Exception as e:
                on_error(f"❌ 無法讀取檔案：{e}")
                return

            doc_text = _extract_docx_text(doc_bytes)
            if not doc_text.strip():
                on_error("❌ 無法從文件提取文字，請確認 .docx 檔案存在且含有文字。")
                return

            prompt_text = (
                _AI_PROMPT_PATH.read_text(encoding="utf-8").strip()
                if _AI_PROMPT_PATH.exists()
                else _AI_PROMPT_DEFAULT
            )
            full_prompt = f"{prompt_text}\n\n=== 文件內容 ===\n{doc_text[:8000]}"

            analyze_btn.config(state="disabled", text="分析中…")
            _streamed[0] = False
            result_txt.delete("1.0", "end")
            result_txt.insert("end", f"分析中… {_SPIN[0]}\n")
            _spin_job[0] = result_txt.after(120, _spin_tick)

            self._run_claude_async(full_prompt, on_result, on_error, on_chunk=on_chunk)

        analyze_btn.config(command=do_analyze)

        # ── 修正拼字功能 ──────────────────────────────
        def do_fix():
            if not self._last_docx_path:
                on_error("❌ 尚未轉換任何文件，請先完成轉換。")
                return
            try:
                doc_bytes = Path(self._last_docx_path).read_bytes()
            except Exception as e:
                on_error(f"❌ 無法讀取檔案：{e}")
                return

            doc_text = _extract_docx_text(doc_bytes)
            if not doc_text.strip():
                on_error("❌ 無法從文件提取文字。")
                return

            full_prompt = f"{_AI_FIX_PROMPT}\n\n=== 文件內容 ===\n{doc_text[:8000]}"

            fix_btn.config(state="disabled", text="修正中…")
            analyze_btn.config(state="disabled")
            _streamed[0] = False
            result_txt.delete("1.0", "end")
            result_txt.insert("end", f"AI 分析拼字錯誤中… {_SPIN[0]}\n")
            _spin_job[0] = result_txt.after(120, _spin_tick)

            def on_fix_result(raw_text):
                import json as _json
                import re as _re
                _stop_spinner()
                result_txt.delete("1.0", "end")

                json_match = _re.search(r'\{[\s\S]*\}', raw_text)
                if not json_match:
                    result_txt.insert("end", f"❌ AI 回應無法解析為 JSON：\n\n{raw_text}")
                    fix_btn.config(state="normal", text="🔧 修正拼字")
                    analyze_btn.config(state="normal")
                    return

                try:
                    data = _json.loads(json_match.group())
                    corrections = data.get("corrections", [])
                    summary = data.get("summary", "")
                except _json.JSONDecodeError as e:
                    result_txt.insert("end", f"❌ JSON 解析失敗：{e}\n\n{raw_text}")
                    fix_btn.config(state="normal", text="🔧 修正拼字")
                    analyze_btn.config(state="normal")
                    return

                if not corrections:
                    result_txt.insert("end", f"✅ AI 未發現需要修正的拼字錯誤。\n\n{summary}")
                    fix_btn.config(state="normal", text="🔧 修正拼字")
                    analyze_btn.config(state="normal")
                    return

                lines = [f"發現 {len(corrections)} 項修正：\n"]
                for i, c in enumerate(corrections, 1):
                    lines.append(f"{i}. 「{c.get('original', '')}」→「{c.get('corrected', '')}」")
                    if c.get("reason"):
                        lines.append(f"   {c['reason']}")
                lines.append("")
                result_txt.insert("end", "\n".join(lines))
                result_txt.update_idletasks()

                try:
                    corrected_bytes, applied, not_found = _apply_text_corrections(doc_bytes, corrections)
                except Exception as e:
                    result_txt.insert("end", f"\n❌ 套用修正失敗：{e}")
                    fix_btn.config(state="normal", text="🔧 修正拼字")
                    analyze_btn.config(state="normal")
                    return

                orig_path = Path(self._last_docx_path)
                out_path = orig_path.with_stem(orig_path.stem + "_corrected")
                try:
                    out_path.write_bytes(corrected_bytes)
                    status_line = f"✅ 已套用 {applied}/{len(corrections)} 項修正\n儲存至：{out_path}"
                    if not_found:
                        status_line += f"\n\n⚠️ 以下 {len(not_found)} 項未在文件中找到（可能是 AI 幻覺或上下文誤判）：\n"
                        status_line += "\n".join(f"  · 「{x}」" for x in not_found)
                    result_txt.insert("end", status_line)
                    self._last_docx_path = str(out_path)
                    self._status.set(f"修正版儲存至：{out_path.name}")
                    self._show_save_result(
                        "拼字修正完成",
                        f"套用 {applied}/{len(corrections)} 項修正\n\n修正版已儲存至：\n{out_path}",
                        str(out_path),
                    )
                except Exception as e:
                    result_txt.insert("end", f"\n❌ 儲存失敗：{e}")

                fix_btn.config(state="normal", text="🔧 修正拼字")
                analyze_btn.config(state="normal")
                result_txt.see("end")

            def on_fix_error(text):
                _stop_spinner()
                result_txt.delete("1.0", "end")
                result_txt.insert("end", text)
                fix_btn.config(state="normal", text="🔧 修正拼字")
                analyze_btn.config(state="normal")

            self._run_claude_async(full_prompt, on_fix_result, on_fix_error)

        fix_btn = tk.Button(hdr, text="🔧 修正拼字", width=12, command=do_fix)
        fix_btn.pack(side="left", padx=(6, 0))

    def _run_claude_async(self, full_prompt: str, on_result, on_error, on_chunk=None):
        """將 full_prompt 透過 claude -p 分析；callback 於 main thread 被呼叫。"""
        def worker():
            env = os.environ.copy()
            if sys.platform == "win32":
                for npm_bin in [
                    str(Path.home() / "AppData" / "Roaming" / "npm"),
                    str(Path.home() / "AppData" / "Local" / "npm"),
                ]:
                    if npm_bin.lower() not in env.get("PATH", "").lower():
                        env["PATH"] = npm_bin + os.pathsep + env.get("PATH", "")

            claude_exe = shutil.which("claude", path=env.get("PATH", ""))
            if not claude_exe:
                err_text = (
                    "❌ 找不到 claude 指令。\n\n"
                    "請確認 Claude Code CLI 已安裝，且安裝目錄在 PATH 中。\n"
                    "安裝說明：https://claude.ai/code\n\n"
                    f"目前 PATH（前 500 字）：\n{env.get('PATH', '')[:500]}"
                )
                self.after(0, lambda t=err_text: on_error(t))
                return

            tf = tempfile.NamedTemporaryFile(
                mode="w", encoding="utf-8", suffix=".txt", delete=False)
            try:
                tf.write(full_prompt)
                tf.close()
                tmp_path = tf.name
                model_id = self._ai_model_var.get().strip()

                if sys.platform == "win32":
                    model_flag = f' --model "{model_id}"' if model_id else ""
                    cmd_str = f'chcp 65001 > nul && type "{tmp_path}" | claude -p{model_flag}'
                    proc = subprocess.Popen(
                        cmd_str, shell=True,
                        stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                        env=env,
                    )
                else:
                    f_in = open(tmp_path, encoding="utf-8")
                    model_args = ["--model", model_id] if model_id else []
                    proc = subprocess.Popen(
                        [claude_exe, "-p"] + model_args, stdin=f_in,
                        stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                        env=env,
                    )

                def _flush_buf(raw_buf):
                    if not raw_buf:
                        return "", b""
                    for trim in range(0, 4):
                        try:
                            return raw_buf[:len(raw_buf) - trim].decode("utf-8"), raw_buf[len(raw_buf) - trim:]
                        except UnicodeDecodeError:
                            continue
                    return raw_buf.decode("utf-8", errors="replace"), b""

                chunks  = []
                raw_buf = b""
                _first  = [True]
                while True:
                    raw = proc.stdout.read(512)
                    if not raw:
                        if raw_buf:
                            text = raw_buf.decode("utf-8", errors="replace")
                            chunks.append(text)
                            if on_chunk:
                                is_first = _first[0]
                                _first[0] = False
                                self.after(0, lambda t=text, f=is_first: on_chunk(t, f))
                        break
                    raw_buf += raw
                    text, raw_buf = _flush_buf(raw_buf)
                    if text:
                        chunks.append(text)
                        if on_chunk:
                            is_first = _first[0]
                            _first[0] = False
                            self.after(0, lambda t=text, f=is_first: on_chunk(t, f))

                proc.wait()
                if sys.platform != "win32":
                    try:
                        f_in.close()
                    except Exception:
                        pass
                stderr_bytes = proc.stderr.read()
                stderr_txt   = stderr_bytes.decode("utf-8", errors="replace").strip()
                stdout_txt   = "".join(chunks).strip()

            finally:
                try:
                    os.unlink(tf.name)
                except Exception:
                    pass

            if proc.returncode == 0:
                result = stdout_txt or "（無回應）"
                self.after(0, lambda r=result: on_result(r))
            else:
                combined = stderr_txt or stdout_txt or f"exit code {proc.returncode}"
                if any(k in combined for k in ("not recognized", "找不到", "無法找到", "is not")):
                    err_text = (
                        "❌ 找不到 claude 指令（cmd 層級）。\n\n"
                        f"claude 路徑：{claude_exe}\n"
                        f"錯誤：{combined}"
                    )
                elif any(k in combined for k in (
                        "auth", "login", "sign in", "not logged", "unauthorized", "401", "403")):
                    err_text = (
                        "❌ Claude CLI 認證失敗。\n\n"
                        "請在終端機執行 claude 完成登入，再重試。\n\n"
                        f"詳細訊息：{combined}"
                    )
                else:
                    detail = ""
                    if stderr_txt:
                        detail += f"\nstderr：{stderr_txt}"
                    if stdout_txt and stdout_txt != stderr_txt:
                        detail += f"\nstdout：{stdout_txt[:500]}"
                    err_text = f"❌ claude 執行失敗（exit code {proc.returncode}）{detail}"
                self.after(0, lambda t=err_text: on_error(t))

        def wrapper():
            try:
                worker()
            except Exception as e:
                self.after(0, lambda err=e: on_error(f"❌ 呼叫失敗：{err}"))

        threading.Thread(target=wrapper, daemon=True).start()


if __name__ == "__main__":
    App().mainloop()
