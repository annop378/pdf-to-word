"""
Position-aware PDF → Word converter.

Every page is rebuilt as a single borderless Word table:
  - rows    = clusters of elements sharing the same vertical band
  - cells   = segments within a row separated by significant horizontal gaps
  - blanks  = horizontal rules → underlined spaces
  - images  = embedded PDF images → inline picture in cell
  - text    = preserves bold / italic / font-size
"""

import io
import os
import sys
import tempfile
from dataclasses import dataclass, field
from pathlib import Path

import fitz
import pdfplumber
import pytesseract
from PIL import Image as PILImage
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# 若在 PyInstaller bundle 內，使用內嵌的 Tesseract；否則使用系統安裝版
_BASE = Path(getattr(sys, "_MEIPASS", ""))
_BUNDLED_TESS = _BASE / "tesseract" / "tesseract.exe"
_SYSTEM_TESS  = Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe")

if _BUNDLED_TESS.exists():
    pytesseract.pytesseract.tesseract_cmd = str(_BUNDLED_TESS)
elif _SYSTEM_TESS.exists():
    pytesseract.pytesseract.tesseract_cmd = str(_SYSTEM_TESS)

ROW_TOLERANCE    = 10   # 垂直中點容差（pt）
COL_GAP          = 40
MIN_BLANK_WIDTH  = 20
BLANK_CHAR_WIDTH = 5.5


# ── data classes ──────────────────────────────────────────────────────────────
@dataclass
class Word:
    text:      str
    x0:        float
    top:       float
    x1:        float
    bottom:    float
    bold:      bool  = False
    italic:    bool  = False
    size:      float = 11.0
    underline: bool  = False
    fontname:  str   = ""


@dataclass
class Blank:
    x0:     float
    top:    float
    x1:     float
    bottom: float


@dataclass
class ImageElem:
    x0:     float
    top:    float
    x1:     float
    bottom: float
    data:   bytes = field(repr=False)
    ext:    str   = "png"


@dataclass
class BoxElem:
    """Placeholder for an empty bordered rectangle (checkbox / empty form cell)."""
    x0:     float
    top:    float
    x1:     float
    bottom: float


@dataclass
class FormFieldElem:
    """An AcroForm widget field extracted from the PDF (text box, checkbox, dropdown…)."""
    x0:         float
    top:        float
    x1:         float
    bottom:     float
    field_type: str = "text"  # "text" | "checkbox" | "radio" | "combobox" | "listbox"


@dataclass
class SeparatorElem:
    """Full-width horizontal rule — acts as a merge barrier between rows."""
    top:    float
    bottom: float


# ── extraction ────────────────────────────────────────────────────────────────
def _page_to_pil(pdf_path: str, idx: int, dpi: int = 200) -> PILImage.Image:
    doc = fitz.open(pdf_path)
    pix = doc[idx].get_pixmap(matrix=fitz.Matrix(dpi / 72, dpi / 72),
                               colorspace=fitz.csRGB)
    img = PILImage.frombytes("RGB", [pix.width, pix.height], pix.samples)
    doc.close()
    return img


def _extract_images(pdf_path: str, page_idx: int) -> list:
    """Extract embedded images with their bounding boxes (top-origin pts)."""
    images = []
    doc = fitz.open(pdf_path)
    page = doc[page_idx]
    seen = set()
    for img_info in page.get_images(full=True):
        xref = img_info[0]
        if xref in seen:
            continue
        seen.add(xref)
        rects = page.get_image_rects(xref)
        if not rects:
            continue
        try:
            base = doc.extract_image(xref)
        except Exception:
            continue
        for rect in rects:
            # skip thin wide rects (horizontal rules rendered as images)
            h = rect.y1 - rect.y0
            w = rect.x1 - rect.x0
            if h < 12 and w > MIN_BLANK_WIDTH and w / max(h, 0.1) > 10:
                continue
            # skip small solid-colour squares — these are checkbox/radio renders
            if w <= 24 and h <= 24:
                try:
                    from PIL import Image as _PILImage
                    import io as _io
                    _img = _PILImage.open(_io.BytesIO(base["image"])).convert("RGB")
                    _pixels = list(_img.getdata())
                    _unique = set(_pixels)
                    # pure black or near-black solid fill → form field icon, skip
                    if len(_unique) <= 4:
                        _avg = sum(sum(p) for p in _pixels) / (len(_pixels) * 3)
                        if _avg < 30 or _avg > 240:
                            continue
                except Exception:
                    pass
            # PyMuPDF rect uses top-origin (y0=top, y1=bottom)
            images.append(ImageElem(
                x0=rect.x0, top=rect.y0,
                x1=rect.x1, bottom=rect.y1,
                data=base["image"],
                ext=base.get("ext", "png"),
            ))
    doc.close()
    return images


def _extract_form_fields(pdf_path: str, page_idx: int) -> list:
    """Extract AcroForm widget fields from a PDF page as FormFieldElem list (top-origin pts)."""
    fields = []
    doc = fitz.open(pdf_path)
    page = doc[page_idx]
    ph = page.rect.height

    _TYPE_MAP = {
        fitz.PDF_WIDGET_TYPE_TEXT:     "text",
        fitz.PDF_WIDGET_TYPE_CHECKBOX: "checkbox",
        fitz.PDF_WIDGET_TYPE_RADIOBUTTON: "radio",
        fitz.PDF_WIDGET_TYPE_COMBOBOX: "combobox",
        fitz.PDF_WIDGET_TYPE_LISTBOX:  "listbox",
    }

    for w in page.widgets() or []:
        r = w.rect
        ftype = _TYPE_MAP.get(w.field_type, "text")
        # PyMuPDF uses top-origin already (y0=top, y1=bottom)
        fields.append(FormFieldElem(
            x0=r.x0, top=r.y0,
            x1=r.x1, bottom=r.y1,
            field_type=ftype,
        ))
    doc.close()
    return fields


def _extract_text_and_blanks(page) -> tuple:
    """Extract Words and Blanks from a pdfplumber page."""
    ph = page.height

    # 建立 char-level underline 查詢用的字典 {(x0,top): underline}
    char_underline: dict = {}
    for ch in page.chars:
        if ch.get("upright", 1):
            char_underline[(round(ch["x0"], 1), round(ch["top"], 1))] = bool(ch.get("underline"))

    # 水平線 → Blank（stroked lines）
    blanks = []
    seen_blanks: set = set()

    words = []
    for w in page.extract_words(extra_attrs=["fontname", "size"],
                                 use_text_flow=False,
                                 keep_blank_chars=False):
        # underscore-only words are fill-in blanks from Word source — treat as Blank
        if w["text"].strip("_") == "" and len(w["text"]) >= 2:
            blanks.append(Blank(x0=w["x0"], top=w["top"],
                                x1=w["x1"], bottom=w["bottom"]))
            continue

        fn = w.get("fontname") or ""
        ul = bool(char_underline.get((round(w["x0"], 1), round(w["top"], 1))))
        words.append(Word(
            text      = w["text"],
            x0        = w["x0"],   top    = w["top"],
            x1        = w["x1"],   bottom = w["bottom"],
            bold      = "Bold"   in fn or "bold"   in fn.lower(),
            italic    = "Italic" in fn or "Oblique" in fn or "italic" in fn.lower(),
            size      = float(w.get("size") or 11),
            underline = ul,
            fontname  = fn,
        ))

    def _add_blank(bx0, bx1, by_top, by_bottom):
        key = (round(bx0), round(by_top))
        if key in seen_blanks:
            return
        seen_blanks.add(key)
        blanks.append(Blank(x0=bx0, top=by_top, x1=bx1, bottom=by_bottom))

    # Full-width threshold: lines wider than 75% of page are section dividers.
    # They are NOT converted to Blank; instead returned as SeparatorElem so
    # _merge_parallel_rows can use them as merge barriers.
    full_width_threshold = page.width * 0.75
    separators = []
    seen_seps: set = set()

    def _add_sep(sy_top, sy_bottom):
        key = round(sy_top)
        if key not in seen_seps:
            seen_seps.add(key)
            separators.append(SeparatorElem(top=sy_top, bottom=sy_bottom))

    for ln in page.lines:
        lx0, lx1 = min(ln["x0"], ln["x1"]), max(ln["x0"], ln["x1"])
        ly0, ly1 = min(ln["y0"], ln["y1"]), max(ln["y0"], ln["y1"])
        if (ly1 - ly0) < 3 and (lx1 - lx0) >= MIN_BLANK_WIDTH:
            if (lx1 - lx0) >= full_width_threshold:
                _add_sep(ph - ly1, ph - ly0)
            else:
                _add_blank(lx0, lx1, ph - ly1, ph - ly0)

    # 薄矩形 → Blank（filled rectangles used as horizontal rules）
    for rect in page.rects:
        rx0, rx1 = min(rect["x0"], rect["x1"]), max(rect["x0"], rect["x1"])
        ry0, ry1 = min(rect["y0"], rect["y1"]), max(rect["y0"], rect["y1"])
        if (ry1 - ry0) < 6 and (rx1 - rx0) >= MIN_BLANK_WIDTH:
            if (rx1 - rx0) >= full_width_threshold:
                _add_sep(ph - ry1, ph - ry0)
            else:
                _add_blank(rx0, rx1, ph - ry1, ph - ry0)

    return words, blanks, separators


# ── row grouping ──────────────────────────────────────────────────────────────
def _group_rows(words, blanks, images, page_width=0, form_fields=None) -> list:
    """
    Group text+blanks into rows by vertical proximity.
    Images and form_fields are NOT used to drive row boundaries.
    Instead each is injected into the row whose vertical midpoint it best overlaps.
    """
    # --- step 1: build rows from text + blanks only -------------------------
    # 右側欄位元素（x0 > 頁面 72%）允許與目前行有小間距仍合併，
    # 但不更新 row_bottom，避免把下一行也拉進來。
    right_threshold = page_width * 0.72 if page_width else float("inf")
    RIGHT_GAP = 8  # pt，右欄允許的最大間距

    text_elems = sorted(words + blanks, key=lambda e: e.top)
    rows = []
    if text_elems:
        current    = []
        row_top    = None
        row_bottom = None
        for elem in text_elems:
            if row_top is None:
                current.append(elem)
                row_top, row_bottom = elem.top, elem.bottom
            else:
                overlap     = min(elem.bottom, row_bottom) - max(elem.top, row_top)
                is_right    = elem.x0 > right_threshold
                gap_allowed = RIGHT_GAP if is_right else 0
                if overlap > -gap_allowed:
                    current.append(elem)
                    # 右欄元素不擴展 row_bottom，避免拉入下一行
                    if not is_right:
                        row_bottom = max(row_bottom, elem.bottom)
                else:
                    rows.append(current)
                    current    = [elem]
                    row_top    = elem.top
                    row_bottom = elem.bottom
        if current:
            rows.append(current)

    # --- step 2: inject each image into the best-matching row ----------------
    for img in images:
        img_mid = (img.top + img.bottom) / 2
        best_idx, best_overlap = None, -1
        for i, row in enumerate(rows):
            row_top    = min(e.top    for e in row)
            row_bottom = max(e.bottom for e in row)
            overlap = min(img.bottom, row_bottom) - max(img.top, row_top)
            if overlap > best_overlap:
                best_overlap = overlap
                best_idx = i
        if best_idx is not None and best_overlap > 0:
            rows[best_idx].append(img)
        else:
            # no overlapping row — insert as its own row at the right position
            inserted = False
            for i, row in enumerate(rows):
                row_top = min(e.top for e in row)
                if img_mid < row_top:
                    rows.insert(i, [img])
                    inserted = True
                    break
            if not inserted:
                rows.append([img])

    # --- step 3: inject each form field (same logic as images) ---------------
    for ff in (form_fields or []):
        ff_mid = (ff.top + ff.bottom) / 2
        best_idx, best_overlap = None, -1
        for i, row in enumerate(rows):
            row_top    = min(e.top    for e in row)
            row_bottom = max(e.bottom for e in row)
            overlap = min(ff.bottom, row_bottom) - max(ff.top, row_top)
            if overlap > best_overlap:
                best_overlap = overlap
                best_idx = i
        if best_idx is not None and best_overlap > 0:
            rows[best_idx].append(ff)
        else:
            inserted = False
            for i, row in enumerate(rows):
                row_top = min(e.top for e in row)
                if ff_mid < row_top:
                    rows.insert(i, [ff])
                    inserted = True
                    break
            if not inserted:
                rows.append([ff])

    # --- step 4: sort each row by x0 ----------------------------------------
    return [sorted(row, key=lambda e: e.x0) for row in rows]


# ── column splitting ──────────────────────────────────────────────────────────
def _split_cells(row_elems) -> list:
    """
    Split row elements into cells.
    Rules (applied in order):
    1. Each Blank is always its own cell (so its bottom border width matches the underline).
    2. Split after a Word ending with ':'.
    3. Split on horizontal gap > COL_GAP.
    BoxElem: deduplicated by identity first (same box injected into multiple rows),
    then each unique box is its own cell.
    """
    if not row_elems:
        return []

    # Deduplicate BoxElems that were injected into multiple rows — keep first by x0
    seen_boxes: set = set()
    deduped = []
    for e in row_elems:
        if isinstance(e, BoxElem):
            if id(e) not in seen_boxes:
                seen_boxes.add(id(e))
                deduped.append(e)
        else:
            deduped.append(e)
    row_elems = deduped

    cells: list = []
    current: list = []

    for elem in row_elems:
        if isinstance(elem, FormFieldElem):
            # AcroForm widget fields are always their own cell
            if current:
                cells.append(current)
                current = []
            cells.append([elem])

        elif isinstance(elem, BoxElem):
            # Empty bordered boxes are always their own cell
            if current:
                cells.append(current)
                current = []
            cells.append([elem])

        elif isinstance(elem, Blank):
            # If the blank overlaps with text already in the current cell it is an
            # inline underline (e.g. the underline drawn beneath "Property Owner
            # Information"). Keep it in the same cell so the whole phrase stays
            # together and the cell gets a bottom border.
            # Only give the blank its own cell when it has no x-overlap with any
            # Word in current (i.e. it is a standalone fill-in blank).
            overlaps_current = any(
                isinstance(w, Word) and w.x0 < elem.x1 and w.x1 > elem.x0
                for w in current
            )
            if overlaps_current:
                current.append(elem)
            else:
                if current:
                    cells.append(current)
                    current = []
                cells.append([elem])
        else:
            if current:
                prev = current[-1]
                gap_split   = elem.x0 - prev.x1 > COL_GAP
                colon_split = isinstance(prev, Word) and prev.text.rstrip().endswith(":")
                if gap_split or colon_split:
                    cells.append(current)
                    current = []
            current.append(elem)

    if current:
        cells.append(current)

    return cells


# ── docx helpers ──────────────────────────────────────────────────────────────
def _remove_all_borders(table):
    tbl   = table._tbl
    tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    for old in tblPr.findall(qn("w:tblBorders")):
        tblPr.remove(old)
    tblBorders = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b = OxmlElement(f"w:{side}")
        b.set(qn("w:val"),   "none")
        b.set(qn("w:sz"),    "0")
        b.set(qn("w:space"), "0")
        b.set(qn("w:color"), "auto")
        tblBorders.append(b)
    tblPr.append(tblBorders)


def _remove_cell_borders(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:tcBorders")):
        tcPr.remove(old)
    tcBorders = OxmlElement("w:tcBorders")
    for side in ("top", "left", "bottom", "right"):
        b = OxmlElement(f"w:{side}")
        b.set(qn("w:val"),   "none")
        b.set(qn("w:sz"),    "0")
        b.set(qn("w:space"), "0")
        b.set(qn("w:color"), "auto")
        tcBorders.append(b)
    tcPr.append(tcBorders)


def _set_cell_bottom_border(cell):
    """在 cell 底部加一條實線框線（用來取代 Blank underline）。"""
    tcPr = cell._tc.get_or_add_tcPr()
    borders = tcPr.find(qn("w:tcBorders"))
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tcPr.append(borders)
    for old in borders.findall(qn("w:bottom")):
        borders.remove(old)
    b = OxmlElement("w:bottom")
    b.set(qn("w:val"),   "single")
    b.set(qn("w:sz"),    "6")
    b.set(qn("w:space"), "0")
    b.set(qn("w:color"), "000000")
    borders.append(b)


def _zero_para_spacing(para):
    """移除段落前後間距。"""
    pPr = para._p.get_or_add_pPr()
    spacing = pPr.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        pPr.append(spacing)
    spacing.set(qn("w:before"), "0")
    spacing.set(qn("w:after"),  "0")
    spacing.set(qn("w:line"),   "240")
    spacing.set(qn("w:lineRule"), "auto")


def _is_rule_image(elem: ImageElem) -> bool:
    """True when the image is a thin horizontal rule (not a real picture)."""
    height = elem.bottom - elem.top
    width  = elem.x1 - elem.x0
    return height < 12 and width > MIN_BLANK_WIDTH and width / max(height, 0.1) > 10


def _infer_alignment(elems, page_width: float):
    """Infer horizontal alignment from element positions relative to page width."""
    text_elems = [e for e in elems if isinstance(e, Word)]
    if not text_elems or page_width <= 0:
        return WD_ALIGN_PARAGRAPH.LEFT
    x0  = min(e.x0 for e in text_elems)
    x1  = max(e.x1 for e in text_elems)
    mid = (x0 + x1) / 2
    # text starting near the left margin is always left-aligned
    if x0 < page_width * 0.25:
        return WD_ALIGN_PARAGRAPH.LEFT
    # right-aligned: x0 > 60% of page width
    if x0 > page_width * 0.60:
        return WD_ALIGN_PARAGRAPH.RIGHT
    # centred: mid within 15% of page centre and not starting near left
    if abs(mid - page_width / 2) < page_width * 0.15:
        return WD_ALIGN_PARAGRAPH.CENTER
    return WD_ALIGN_PARAGRAPH.LEFT


def _insert_formcheckbox_field(para):
    """Insert a Word Legacy Forms Checkbox Field."""
    p = para._p

    r1 = OxmlElement("w:r")
    fc1 = OxmlElement("w:fldChar")
    fc1.set(qn("w:fldCharType"), "begin")
    ffData = OxmlElement("w:ffData")
    nm = OxmlElement("w:name")
    nm.set(qn("w:val"), "")
    ffData.append(nm)
    ffData.append(OxmlElement("w:enabled"))
    coe = OxmlElement("w:calcOnExit")
    coe.set(qn("w:val"), "0")
    ffData.append(coe)
    cb = OxmlElement("w:checkBox")
    cbsz = OxmlElement("w:sizeAuto")
    cb.append(cbsz)
    chk = OxmlElement("w:default")
    chk.set(qn("w:val"), "0")
    cb.append(chk)
    ffData.append(cb)
    fc1.append(ffData)
    r1.append(fc1)
    p.append(r1)

    r2 = OxmlElement("w:r")
    instr = OxmlElement("w:instrText")
    instr.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    instr.text = " FORMCHECKBOX "
    r2.append(instr)
    p.append(r2)

    r3 = OxmlElement("w:r")
    fc3 = OxmlElement("w:fldChar")
    fc3.set(qn("w:fldCharType"), "separate")
    r3.append(fc3)
    p.append(r3)

    r4 = OxmlElement("w:r")
    t4 = OxmlElement("w:t")
    t4.text = ""
    r4.append(t4)
    p.append(r4)

    r5 = OxmlElement("w:r")
    fc5 = OxmlElement("w:fldChar")
    fc5.set(qn("w:fldCharType"), "end")
    r5.append(fc5)
    p.append(r5)


def _insert_formdropdown_field(para):
    """Insert a Word Legacy Forms Dropdown Field."""
    p = para._p

    r1 = OxmlElement("w:r")
    fc1 = OxmlElement("w:fldChar")
    fc1.set(qn("w:fldCharType"), "begin")
    ffData = OxmlElement("w:ffData")
    nm = OxmlElement("w:name")
    nm.set(qn("w:val"), "")
    ffData.append(nm)
    ffData.append(OxmlElement("w:enabled"))
    coe = OxmlElement("w:calcOnExit")
    coe.set(qn("w:val"), "0")
    ffData.append(coe)
    dd = OxmlElement("w:ddList")
    dddef = OxmlElement("w:default")
    dddef.set(qn("w:val"), "0")
    dd.append(dddef)
    li = OxmlElement("w:listEntry")
    li.set(qn("w:val"), "")
    dd.append(li)
    ffData.append(dd)
    fc1.append(ffData)
    r1.append(fc1)
    p.append(r1)

    r2 = OxmlElement("w:r")
    instr = OxmlElement("w:instrText")
    instr.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    instr.text = " FORMDROPDOWN "
    r2.append(instr)
    p.append(r2)

    r3 = OxmlElement("w:r")
    fc3 = OxmlElement("w:fldChar")
    fc3.set(qn("w:fldCharType"), "separate")
    r3.append(fc3)
    p.append(r3)

    r4 = OxmlElement("w:r")
    t4 = OxmlElement("w:t")
    t4.text = ""
    r4.append(t4)
    p.append(r4)

    r5 = OxmlElement("w:r")
    fc5 = OxmlElement("w:fldChar")
    fc5.set(qn("w:fldCharType"), "end")
    r5.append(fc5)
    p.append(r5)


_FONT_MAP = {
    # PDF embedded font name fragment → Word font name
    "arial":            "Arial",
    "helvetica":        "Arial",
    "timesnewroman":    "Times New Roman",
    "times":            "Times New Roman",
    "courier":          "Courier New",
    "calibri":          "Calibri",
    "cambria":          "Cambria",
    "verdana":          "Verdana",
    "tahoma":           "Tahoma",
    "trebuchet":        "Trebuchet MS",
    "garamond":         "Garamond",
    "georgia":          "Georgia",
    "palatino":         "Palatino Linotype",
    "bookman":          "Bookman Old Style",
    "comic":            "Comic Sans MS",
    "impact":           "Impact",
    "symbol":           "Symbol",
    "wingdings":        "Wingdings",
}


def _resolve_font(pdf_fontname: str) -> str:
    """Map a PDF embedded font name (e.g. 'LWBXFC+ArialMT') to a Word font name."""
    if not pdf_fontname:
        return ""
    # strip subset prefix like 'LWBXFC+'
    base = pdf_fontname.split("+")[-1]
    key = base.lower().replace("-", "").replace(" ", "").replace("mt", "").replace("ps", "")
    for fragment, word_font in _FONT_MAP.items():
        if fragment in key:
            return word_font
    return base  # fallback: use as-is


def _insert_formtext_field(para, width_pts: float = 0):
    """Insert a Word Legacy Forms Text Field (Developer → Legacy Forms → Text Field)."""
    p = para._p

    r1 = OxmlElement("w:r")
    fc1 = OxmlElement("w:fldChar")
    fc1.set(qn("w:fldCharType"), "begin")
    ffData = OxmlElement("w:ffData")
    nm = OxmlElement("w:name")
    nm.set(qn("w:val"), "")
    ffData.append(nm)
    ffData.append(OxmlElement("w:enabled"))
    coe = OxmlElement("w:calcOnExit")
    coe.set(qn("w:val"), "0")
    ffData.append(coe)
    ffData.append(OxmlElement("w:textInput"))
    fc1.append(ffData)
    r1.append(fc1)
    p.append(r1)

    r2 = OxmlElement("w:r")
    instr = OxmlElement("w:instrText")
    instr.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    instr.text = " FORMTEXT "
    r2.append(instr)
    p.append(r2)

    r3 = OxmlElement("w:r")
    fc3 = OxmlElement("w:fldChar")
    fc3.set(qn("w:fldCharType"), "separate")
    r3.append(fc3)
    p.append(r3)

    r4 = OxmlElement("w:r")
    t4 = OxmlElement("w:t")
    t4.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t4.text = " " * max(1, int(width_pts / BLANK_CHAR_WIDTH))
    r4.append(t4)
    p.append(r4)

    r5 = OxmlElement("w:r")
    fc5 = OxmlElement("w:fldChar")
    fc5.set(qn("w:fldCharType"), "end")
    r5.append(fc5)
    p.append(r5)


def _write_cell(cell, elems, page_width: float = 0):
    if not elems:
        return False

    has_blank = any(
        isinstance(e, Blank) or (isinstance(e, ImageElem) and _is_rule_image(e))
        for e in elems
    )
    has_word = any(isinstance(e, Word) for e in elems)
    # Standalone-blank cells (no words) → FORMTEXT field; no bottom border needed.
    # Mixed cells (text + inline blank) → keep bottom border for underline appearance.
    use_formtext = has_blank and not has_word

    # Group elements into visual lines by vertical position.
    # When a cell contains elements merged from multiple original rows,
    # each original row becomes its own paragraph inside the cell.
    lines: list = [[]]
    prev_bottom: float = None
    for elem in elems:
        if prev_bottom is not None and elem.top > prev_bottom + 2:
            lines.append([])
        lines[-1].append(elem)
        prev_bottom = max(prev_bottom, elem.bottom) if prev_bottom is not None else elem.bottom

    for line_idx, line_elems in enumerate(lines):
        if not line_elems:
            continue
        if line_idx == 0:
            para = cell.paragraphs[0]
        else:
            para = cell.add_paragraph()
        _zero_para_spacing(para)
        para.alignment = _infer_alignment(line_elems, page_width)

        for elem in line_elems:
            if isinstance(elem, FormFieldElem):
                _insert_formtext_field(para, elem.x1 - elem.x0)

            elif isinstance(elem, BoxElem):
                if use_formtext:
                    _insert_formtext_field(para, elem.x1 - elem.x0)

            elif isinstance(elem, Blank):
                if use_formtext:
                    _insert_formtext_field(para, elem.x1 - elem.x0)
                # else: bottom border handled by cell border (mixed text+blank)

            elif isinstance(elem, ImageElem):
                if _is_rule_image(elem):
                    pass  # treat as blank — bottom border only
                else:
                    width_pts = elem.x1 - elem.x0
                    width_in  = max(0.3, min(width_pts / 72, 3.0))
                    buf = io.BytesIO(elem.data)
                    try:
                        run = para.add_run()
                        run.add_picture(buf, width=Inches(width_in))
                    except Exception:
                        pass

            else:  # Word
                run = para.add_run(elem.text + " ")
                run.bold      = elem.bold
                run.italic    = elem.italic
                run.font.size = Pt(max(7, min(float(elem.size), 28)))
                run.underline = False
                font_name = _resolve_font(elem.fontname)
                if font_name:
                    run.font.name = font_name
                    from docx.oxml.ns import qn as _qn
                    rPr = run._r.get_or_add_rPr()
                    rFonts = rPr.find(_qn("w:rFonts"))
                    if rFonts is None:
                        from docx.oxml import OxmlElement as _OxmlElement
                        rFonts = _OxmlElement("w:rFonts")
                        rPr.insert(0, rFonts)
                    rFonts.set(_qn("w:ascii"),    font_name)
                    rFonts.set(_qn("w:hAnsi"),    font_name)
                    rFonts.set(_qn("w:eastAsia"), font_name)
                    rFonts.set(_qn("w:cs"),       font_name)

    # Pure-blank cells use FORMTEXT — no bottom border needed.
    # Mixed cells (text with inline blank/underline) still need the bottom border.
    return has_blank and not use_formtext


# ── underline annotation ──────────────────────────────────────────────────────
def _annotate_underlines(rows: list, splits: list = None) -> tuple:
    """
    Scan rows for Blank-only rows that are drawn just below a text row
    (drawn underlines in PDF are separate graphic objects, not font attributes).
    Merge those blanks into the text row above as column-aligned bottom-border markers,
    then remove the blank-only rows from the list.

    Returns:
        (clean_rows, clean_splits, bottom_border_set)
        clean_rows        – rows with blank-only rows removed
        clean_splits      – matching pre-computed splits (or None when splits is None)
        bottom_border_set – set of (row_idx, x0, x1) tuples for cells needing bottom border
    """
    UNDERLINE_GAP = 8  # max pt gap between text bottom and line top

    bottom_border_set: set = set()
    rows_to_remove: set = set()

    for i, row in enumerate(rows):
        if not row or not all(
            isinstance(e, Blank) or (isinstance(e, ImageElem) and _is_rule_image(e))
            for e in row
        ):
            continue
        prev_idx = i - 1
        while prev_idx >= 0 and all(isinstance(e, Blank) for e in rows[prev_idx]):
            prev_idx -= 1
        if prev_idx < 0:
            continue
        prev_bottom = max(e.bottom for e in rows[prev_idx])
        row_top = min(e.top for e in row)
        if row_top - prev_bottom > UNDERLINE_GAP:
            continue

        for blank in row:
            bottom_border_set.add((prev_idx, blank.x0, blank.x1))
        rows_to_remove.add(i)

    clean_rows   = [r for idx, r in enumerate(rows) if idx not in rows_to_remove]
    clean_splits = (
        [s for idx, s in enumerate(splits) if idx not in rows_to_remove]
        if splits is not None else None
    )

    removed_before = [0] * len(rows)
    count = 0
    for idx in range(len(rows)):
        if idx in rows_to_remove:
            count += 1
        removed_before[idx] = count

    reindexed: set = set()
    for (r_idx, x0, x1) in bottom_border_set:
        reindexed.add((r_idx - removed_before[r_idx], x0, x1))

    return clean_rows, clean_splits, reindexed


def _cell_needs_border(row_idx: int, cell_elems: list,
                        bottom_border_set: set) -> bool:
    """True if this cell contains a Blank, or overlaps an underline span."""
    if any(isinstance(e, Blank) for e in cell_elems):
        return True
    if not cell_elems:
        return False
    cell_x0 = min(e.x0 for e in cell_elems)
    cell_x1 = max(e.x1 for e in cell_elems)
    for (r_idx, bx0, bx1) in bottom_border_set:
        if r_idx != row_idx:
            continue
        overlap = min(cell_x1, bx1) - max(cell_x0, bx0)
        if overlap > 2:
            return True
    return False


# ── parallel-row merging ──────────────────────────────────────────────────────
def _cells_parallel(cells_a: list, cells_b: list) -> bool:
    """
    True when two rows have the same column structure (parallel).
    Uses x0 alignment of Word elements only (ignoring Blank/BoxElem anchors):
    corresponding columns must start at roughly the same horizontal position.
    Falls back to overall cell x0 when a cell has no Word elements.
    """
    if len(cells_a) != len(cells_b):
        return False
    for ca, cb in zip(cells_a, cells_b):
        words_a = [e for e in ca if isinstance(e, Word)]
        words_b = [e for e in cb if isinstance(e, Word)]
        a_x0 = min(e.x0 for e in words_a) if words_a else min(e.x0 for e in ca)
        b_x0 = min(e.x0 for e in words_b) if words_b else min(e.x0 for e in cb)
        if abs(a_x0 - b_x0) > 20:
            return False
    return True


def _merge_parallel_rows(rows: list, separators: list = None) -> tuple:
    """
    Merge consecutive rows that (a) share the same column structure and
    (b) have no blank-line gap between them, and (c) are not separated by
    a full-width divider line (SeparatorElem barrier).

    Returns:
        (merged_rows, merged_splits)
    """
    if not rows:
        return [], []

    # Build a set of separator y-positions (top) for fast lookup
    sep_tops = sorted(s.top for s in (separators or []))

    def _barrier_between(top_a, top_b):
        """True if any separator falls between row A's bottom and row B's top."""
        for st in sep_tops:
            if top_a < st < top_b:
                return True
        return False

    def _row_bottom(r): return max(e.bottom for e in r)
    def _row_top(r):    return min(e.top    for e in r)
    def _avg_line_height(r):
        heights = [e.bottom - e.top for e in r if isinstance(e, Word)]
        return sum(heights) / len(heights) if heights else 10

    split_rows = [_split_cells(r) for r in rows]

    merged_rows   = [list(rows[0])]
    merged_splits = [list(split_rows[0])]

    for i, row in enumerate(rows[1:], start=1):
        prev_row   = merged_rows[-1]
        prev_split = merged_splits[-1]
        curr_split = split_rows[i]

        prev_bottom = _row_bottom(prev_row)
        curr_top    = _row_top(row)
        gap         = curr_top - prev_bottom
        line_h      = _avg_line_height(prev_row)
        no_gap      = gap <= line_h * 1.2
        parallel    = _cells_parallel(prev_split, curr_split)
        blocked     = _barrier_between(prev_bottom, curr_top)

        has_ff = any(isinstance(e, FormFieldElem) for e in prev_row) or \
                 any(isinstance(e, FormFieldElem) for e in row)

        if no_gap and parallel and not blocked and not has_ff:
            new_split = [ca + cb for ca, cb in zip(prev_split, curr_split)]
            merged_rows[-1]   = [e for cell in new_split for e in cell]
            merged_splits[-1] = new_split
        else:
            merged_rows.append(list(row))
            merged_splits.append(list(curr_split))

    return merged_rows, merged_splits


def _insert_parallel_spacers(rows: list, splits: list) -> tuple:
    """
    Between consecutive parallel rows that have a gap (> 1.2× line height),
    insert an empty spacer row so the table preserves the vertical white space.
    """
    if len(rows) <= 1:
        return rows, splits

    def _row_bottom(r): return max(e.bottom for e in r) if r else 0
    def _row_top(r):    return min(e.top    for e in r) if r else 0
    def _avg_line_height(r):
        heights = [e.bottom - e.top for e in r if isinstance(e, Word)]
        return sum(heights) / len(heights) if heights else 10

    out_rows   = [rows[0]]
    out_splits = [splits[0]]

    for i in range(1, len(rows)):
        prev_row   = rows[i - 1]
        curr_row   = rows[i]
        prev_split = splits[i - 1]
        curr_split = splits[i]

        if prev_row and curr_row and prev_split and curr_split:
            gap      = _row_top(curr_row) - _row_bottom(prev_row)
            line_h   = _avg_line_height(prev_row)
            has_gap  = gap > line_h * 1.2
            parallel = _cells_parallel(prev_split, curr_split)

            if has_gap and parallel:
                out_rows.append([])
                out_splits.append([])

        out_rows.append(curr_row)
        out_splits.append(curr_split)

    return out_rows, out_splits


# ── page builder ─────────────────────────────────────────────────────────────
def _build_page_table(doc: Document, rows: list, page_width: float = 0,
                      separators: list = None):
    if not rows:
        return

    # First pass: remove blank-only underline rows BEFORE merging so they don't
    # interrupt consecutive parallel rows (e.g. "ALL EXIT LIGHTS" underline rows
    # sit between text lines and break _cells_parallel chain).
    rows, _, _ = _annotate_underlines(rows, splits=None)
    if not rows:
        return

    rows, splits = _merge_parallel_rows(rows, separators=separators)
    rows, splits = _insert_parallel_spacers(rows, splits)
    rows, splits, bottom_border_set = _annotate_underlines(rows, splits)
    if not rows:
        return

    # Use pre-computed splits from merge step; fall back to recomputing if missing
    row_cells_list = splits if splits is not None else [_split_cells(r) for r in rows]
    # For rows not merged, splits may still need recomputation (shouldn't happen)
    row_cells_list = [
        s if s is not None else _split_cells(rows[i])
        for i, s in enumerate(row_cells_list)
    ]
    max_cols = max((len(c) for c in row_cells_list), default=1)

    tbl = doc.add_table(rows=len(rows), cols=max_cols)
    _remove_all_borders(tbl)

    tblPr = tbl._tbl.tblPr
    tblW  = OxmlElement("w:tblW")
    tblW.set(qn("w:type"), "pct")
    tblW.set(qn("w:w"),    "5000")
    tblPr.append(tblW)

    tblLayout = OxmlElement("w:tblLayout")
    tblLayout.set(qn("w:type"), "autofit")
    tblPr.append(tblLayout)

    for r_idx, (row, cells) in enumerate(zip(rows, row_cells_list)):
        tbl_row = tbl.rows[r_idx]
        n = len(cells)
        if n == 0:
            continue

        if n == 1:
            merged = tbl_row.cells[0]
            for c in range(1, max_cols):
                merged = merged.merge(tbl_row.cells[c])
            has_blank = _write_cell(merged, cells[0], page_width)
            _remove_cell_borders(merged)
            if has_blank or _cell_needs_border(r_idx, cells[0], bottom_border_set):
                _set_cell_bottom_border(merged)

        elif n == max_cols:
            for c_idx, cell_elems in enumerate(cells):
                cell = tbl_row.cells[c_idx]
                has_blank = _write_cell(cell, cell_elems, page_width)
                _remove_cell_borders(cell)
                if has_blank or _cell_needs_border(r_idx, cell_elems, bottom_border_set):
                    _set_cell_bottom_border(cell)

        else:
            cols_per = max_cols // n
            rem      = max_cols % n
            c_start  = 0
            for i, cell_elems in enumerate(cells):
                span  = cols_per + (1 if i < rem else 0)
                c_end = c_start + span - 1
                merged = tbl_row.cells[c_start]
                for c in range(c_start + 1, c_end + 1):
                    merged = merged.merge(tbl_row.cells[c])
                has_blank = _write_cell(merged, cell_elems, page_width)
                _remove_cell_borders(merged)
                if has_blank or _cell_needs_border(r_idx, cell_elems, bottom_border_set):
                    _set_cell_bottom_border(merged)
                c_start = c_end + 1


# ── OCR fallback ──────────────────────────────────────────────────────────────
def _build_ocr_page(doc: Document, pdf_path: str, page_idx: int):
    img      = _page_to_pil(pdf_path, page_idx)
    ocr_text = pytesseract.image_to_string(img, lang="eng", config="--psm 6")
    for line in ocr_text.splitlines():
        doc.add_paragraph(line.strip())


# ── main entry ────────────────────────────────────────────────────────────────
def convert_pdf_to_word(pdf_path: str) -> bytes:
    doc = Document()

    for p in list(doc.paragraphs):
        p._element.getparent().remove(p._element)

    for section in doc.sections:
        section.left_margin   = Inches(0.75)
        section.right_margin  = Inches(0.75)
        section.top_margin    = Inches(0.75)
        section.bottom_margin = Inches(0.75)

    with pdfplumber.open(pdf_path) as pdf:
        num_pages = len(pdf.pages)
        for page_idx, page in enumerate(pdf.pages):
            words, blanks, separators = _extract_text_and_blanks(page)
            images                    = _extract_images(pdf_path, page_idx)
            form_fields               = _extract_form_fields(pdf_path, page_idx)

            # Filter out any text / blanks that fall entirely inside a form field
            # bounding box — the widget already captures that content.
            if form_fields:
                def _inside_any_ff(e):
                    for ff in form_fields:
                        # significant overlap (> 50% of element area) with any form field
                        ox = min(e.x1, ff.x1) - max(e.x0, ff.x0)
                        oy = min(e.bottom, ff.bottom) - max(e.top, ff.top)
                        if ox > 0 and oy > 0:
                            elem_area = max((e.x1 - e.x0) * (e.bottom - e.top), 1)
                            if ox * oy / elem_area > 0.5:
                                return True
                    return False
                words  = [w for w in words  if not _inside_any_ff(w)]
                blanks = [b for b in blanks if not _inside_any_ff(b)]
                images = [i for i in images if not _inside_any_ff(i)]

            if not words and not images and not form_fields:
                _build_ocr_page(doc, pdf_path, page_idx)
            else:
                rows = _group_rows(words, blanks, images,
                                   page_width=page.width,
                                   form_fields=form_fields)
                _build_page_table(doc, rows, page_width=page.width,
                                  separators=separators)

            if page_idx < num_pages - 1:
                doc.add_page_break()

    buf = io.BytesIO()
    doc.save(buf)
    return _inject_form_protection(buf.getvalue())


def _inject_form_protection(docx_bytes: bytes) -> bytes:
    """
    Inject <w:documentProtection w:edit="forms" w:enforcement="0"/> into
    settings.xml so that Legacy Form Fields render as fillable controls in Word.
    enforcement="0" means the protection is declared but NOT enforced (password-free,
    users can still edit layout). Set to "1" to lock the document to forms-only.
    """
    import zipfile as _zf
    import re as _re

    out = io.BytesIO()
    with _zf.ZipFile(io.BytesIO(docx_bytes), "r") as zin, \
         _zf.ZipFile(out, "w", compression=_zf.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/settings.xml":
                xml = data.decode("utf-8")
                if "documentProtection" not in xml:
                    prot = '<w:documentProtection w:edit="forms" w:enforcement="0"/>'
                    xml = _re.sub(r"(<w:settings\b[^>]*>)", r"\1" + prot, xml, count=1)
                data = xml.encode("utf-8")
            zout.writestr(item, data)
    return out.getvalue()


# ── image → word ──────────────────────────────────────────────────────────────
IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif", ".webp"}

def convert_image_to_word(image_path: str) -> bytes:
    """OCR 辨識圖片各字的位置，以表格排版重建 Word（與 PDF 流程一致）。"""
    img = PILImage.open(image_path)

    # 取得圖片 DPI，用於 pixel → pt 換算（預設 96 DPI）
    dpi_info = img.info.get("dpi", (96, 96))
    dpi_x = float(dpi_info[0]) if hasattr(dpi_info, "__getitem__") else float(dpi_info)
    if dpi_x <= 0:
        dpi_x = 96.0
    scale = 72.0 / dpi_x  # pixel → pt

    data = pytesseract.image_to_data(
        img, lang="eng", config="--psm 6",
        output_type=pytesseract.Output.DATAFRAME,
    )
    # 過濾低信心與空白字
    data = data[(data["conf"] > 0) & (data["text"].str.strip() != "")]

    words = []
    for _, row in data.iterrows():
        x0     = row["left"]  * scale
        top    = row["top"]   * scale
        x1     = (row["left"] + row["width"])  * scale
        bottom = (row["top"]  + row["height"]) * scale
        words.append(Word(text=str(row["text"]), x0=x0, top=top, x1=x1, bottom=bottom))

    page_width = img.width * scale

    doc = Document()
    for p in list(doc.paragraphs):
        p._element.getparent().remove(p._element)
    for section in doc.sections:
        section.left_margin   = Inches(0.75)
        section.right_margin  = Inches(0.75)
        section.top_margin    = Inches(0.75)
        section.bottom_margin = Inches(0.75)

    if words:
        rows = _group_rows(words, [], [], page_width=page_width)
        _build_page_table(doc, rows, page_width=page_width)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ── word → word (表格版) ───────────────────────────────────────────────────────
def convert_word_to_word(docx_path: str) -> bytes:
    """Word → PDF（透過 Microsoft Word COM）→ 表格版 Word。"""
    from docx2pdf import convert as _docx2pdf

    tmp_fd, tmp_path = tempfile.mkstemp(suffix=".pdf")
    os.close(tmp_fd)

    try:
        # PyInstaller console=False 時 sys.stdout/stderr 為 None，
        # docx2pdf 內部的 tqdm 會嘗試寫入 stdout 而爆 'NoneType has no attribute write'
        # 先用 devnull 替代，轉換完再還原
        import sys
        _stdout, _stderr = sys.stdout, sys.stderr
        devnull = open(os.devnull, "w")
        try:
            sys.stdout = devnull
            sys.stderr = devnull
            _docx2pdf(docx_path, tmp_path)
        finally:
            sys.stdout = _stdout
            sys.stderr = _stderr
            devnull.close()

        return convert_pdf_to_word(tmp_path)
    finally:
        Path(tmp_path).unlink(missing_ok=True)
