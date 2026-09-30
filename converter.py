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
COL_GAP          = 55
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


@dataclass
class SpacerElem:
    """Phantom element representing vertical white-space between parallel rows."""
    height_pt: float
    x0:        float = 0.0
    top:       float = 0.0
    x1:        float = 0.0
    bottom:    float = 0.0


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


def _has_table_structure(page) -> bool:
    """True if the page has significant rectangular cell borders (a bordered form)."""
    count = 0
    for rect in page.rects:
        w = abs(rect.get("x1", 0) - rect.get("x0", 0))
        h = abs(rect.get("y1", 0) - rect.get("y0", 0))
        if w > 40 and h > 8:
            count += 1
    return count >= 4


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
                # Only split after ":" when next element is a fill-in blank or form field,
                # not when followed by plain text (avoids fragmenting body text like "Note: ...").
                colon_split = (
                    isinstance(prev, Word) and
                    prev.text.rstrip().endswith(":") and
                    isinstance(elem, (Blank, FormFieldElem, BoxElem))
                )
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


def _set_all_cell_borders(cell, color: str = "808080"):
    """Apply thin single-line borders on all 4 sides (for bordered-form pages)."""
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:tcBorders")):
        tcPr.remove(old)
    tcBorders = OxmlElement("w:tcBorders")
    for side in ("top", "left", "bottom", "right"):
        b = OxmlElement(f"w:{side}")
        b.set(qn("w:val"),   "single")
        b.set(qn("w:sz"),    "4")
        b.set(qn("w:space"), "0")
        b.set(qn("w:color"), color)
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


def _reset_doc_spacing(doc: Document):
    """重設 Normal 樣式段落間距，消除 python-docx 預設範本的 space_after 與 1.08 行距。"""
    try:
        style_elem = doc.styles['Normal'].element
        pPr = style_elem.get_or_add_pPr()
        for old in pPr.findall(qn("w:spacing")):
            pPr.remove(old)
        spacing = OxmlElement("w:spacing")
        spacing.set(qn("w:before"),   "0")
        spacing.set(qn("w:after"),    "0")
        spacing.set(qn("w:line"),     "240")
        spacing.set(qn("w:lineRule"), "auto")
        pPr.append(spacing)
    except Exception:
        pass


def _is_rule_image(elem: ImageElem) -> bool:
    """True when the image is a thin horizontal rule (not a real picture)."""
    height = elem.bottom - elem.top
    width  = elem.x1 - elem.x0
    return height < 12 and width > MIN_BLANK_WIDTH and width / max(height, 0.1) > 10


def _compute_col_widths(row_cells_list: list, max_cols: int,
                        content_x_min: float, content_x_max: float,
                        usable_width_pt: float) -> tuple:
    """
    從各列元素的 x 座標計算欄寬。
    以 max_cols 的列為樣本，取中位數邊界，轉換為 twips。
    回傳 (widths_twips: list[int], boundaries_pdf: list[float])。
    """
    TWI = 20  # 1 pt = 20 twips
    content_span = max(content_x_max - content_x_min, 1.0)

    if max_cols <= 1:
        w = round(usable_width_pt * TWI)
        return ([max(w, 200)], [content_x_min, content_x_max])

    col_x0_samples = [[] for _ in range(max_cols)]
    last_x1_samples: list = []

    for cells in row_cells_list:
        if len(cells) != max_cols:
            continue
        for i, elems in enumerate(cells):
            if not elems:
                continue
            x0 = min(e.x0 for e in elems)
            x1 = max(e.x1 for e in elems)
            col_x0_samples[i].append(x0)
            if i == max_cols - 1:
                last_x1_samples.append(x1)

    def _med(lst):
        if not lst:
            return None
        s = sorted(lst)
        return s[len(s) // 2]

    boundaries = [_med(s) for s in col_x0_samples]
    boundaries.append(_med(last_x1_samples) or content_x_max)

    # 線性插補缺失邊界
    for i in range(len(boundaries)):
        if boundaries[i] is not None:
            continue
        lo_i = next((j for j in range(i - 1, -1, -1) if boundaries[j] is not None), None)
        hi_i = next((j for j in range(i + 1, len(boundaries)) if boundaries[j] is not None), None)
        if lo_i is None and hi_i is None:
            boundaries[i] = content_x_min + content_span * i / max_cols
        elif lo_i is None:
            boundaries[i] = boundaries[hi_i] - content_span * (hi_i - i) / max_cols
        elif hi_i is None:
            boundaries[i] = boundaries[lo_i] + content_span * (i - lo_i) / max_cols
        else:
            t = (i - lo_i) / (hi_i - lo_i)
            boundaries[i] = boundaries[lo_i] + t * (boundaries[hi_i] - boundaries[lo_i])

    raw = [max(boundaries[i + 1] - boundaries[i], 1.0) for i in range(max_cols)]
    total = sum(raw)
    widths_twips = [max(round(w / total * usable_width_pt * TWI), 200) for w in raw]
    return (widths_twips, boundaries)


def _set_cell_width(cell, width_twips: int):
    """設定儲存格的明確寬度（twips）。"""
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:tcW")):
        tcPr.remove(old)
    tcW = OxmlElement("w:tcW")
    tcW.set(qn("w:type"), "dxa")
    tcW.set(qn("w:w"), str(max(width_twips, 1)))
    tcPr.insert(0, tcW)


def _set_cell_margins(cell, top_dxa: int = 0, bottom_dxa: int = 0,
                      left_dxa: int = 28, right_dxa: int = 28):
    """設定儲存格內部邊距（twips）。預設上下為 0，左右為 28 twips（≈1.4pt），
    消除 Word 預設 5.4pt 邊距造成的文字換行與列高膨脹。"""
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:tcMar")):
        tcPr.remove(old)
    tcMar = OxmlElement("w:tcMar")
    for side, val in (("top", top_dxa), ("left", left_dxa),
                      ("bottom", bottom_dxa), ("right", right_dxa)):
        m = OxmlElement(f"w:{side}")
        m.set(qn("w:w"), str(val))
        m.set(qn("w:type"), "dxa")
        tcMar.append(m)
    tcPr.append(tcMar)


def _set_row_height(tbl_row, height_pt: float, exact: bool = False):
    """設定列高（twips）。exact=True 用於 SpacerElem 列（精確高度），
    否則用 atLeast（最小高度，可因內容增高）。"""
    trPr = tbl_row._tr.get_or_add_trPr()
    for old in trPr.findall(qn("w:trHeight")):
        trPr.remove(old)
    trH = OxmlElement("w:trHeight")
    trH.set(qn("w:val"), str(max(1, round(height_pt * 20))))
    trH.set(qn("w:hRule"), "exact" if exact else "atLeast")
    trPr.append(trH)


def _cell_col_range(cell_elems: list, col_boundaries: list, max_cols: int) -> tuple:
    """依元素 x 範圍對應 (c_start, c_end) 欄索引。"""
    if not cell_elems:
        return (0, 0)
    x0 = min(e.x0 for e in cell_elems)
    x1 = max(e.x1 for e in cell_elems)

    def _find_col(x):
        for i in range(max_cols):
            if x < col_boundaries[i + 1]:
                return i
        return max_cols - 1

    c_start = _find_col(x0)
    c_end   = _find_col(max(x0, x1 - 2))
    return (c_start, max(c_start, min(c_end, max_cols - 1)))


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
                    # Use exact PDF dimensions; cap only at page boundary, not an arbitrary 3".
                    width_in = max(0.1, width_pts / 72)
                    if page_width > 0:
                        max_width_in = max(0.1, (page_width - 108) / 72)
                        width_in = min(width_in, max_width_in)
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
        while prev_idx >= 0 and (
            all(isinstance(e, Blank) for e in rows[prev_idx]) or
            all(isinstance(e, SpacerElem) for e in rows[prev_idx])
        ):
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
                spacer = SpacerElem(
                    height_pt=gap,
                    top=_row_bottom(prev_row),
                    bottom=_row_top(curr_row),
                )
                out_rows.append([spacer])
                out_splits.append([])

        out_rows.append(curr_row)
        out_splits.append(curr_split)

    return out_rows, out_splits


# ── page builder ─────────────────────────────────────────────────────────────
def _build_page_table(doc: Document, rows: list, page_width: float = 0,
                      separators: list = None, max_cols_cap: int = 0,
                      show_borders: bool = False):
    """
    max_cols_cap: when > 0, caps the number of table columns to this value.
    Use for OCR-derived rows to prevent extreme column fragmentation on forms.
    """
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
    if max_cols_cap > 0:
        max_cols = min(max_cols, max_cols_cap)

    # ── 計算欄寬：從元素 x 座標取樣，轉換為絕對寬度 ────────────────────────
    MARGIN_PT = 0.75 * 72  # 與 convert_pdf_to_word 的 section margin 一致
    usable_width_pt = max((page_width or 595.0) - 2 * MARGIN_PT, 200.0)

    all_elems_flat = [e for row in rows for e in row]
    if all_elems_flat and page_width > 0:
        content_x_min = min(e.x0 for e in all_elems_flat)
        content_x_max = max(e.x1 for e in all_elems_flat)
    else:
        content_x_min, content_x_max = 0.0, page_width or 595.0

    col_widths_twips, col_boundaries_pdf = _compute_col_widths(
        row_cells_list, max_cols, content_x_min, content_x_max, usable_width_pt
    )

    tbl = doc.add_table(rows=len(rows), cols=max_cols)
    _remove_all_borders(tbl)

    tblPr = tbl._tbl.tblPr
    tblW  = OxmlElement("w:tblW")
    tblW.set(qn("w:type"), "dxa")
    tblW.set(qn("w:w"),    str(sum(col_widths_twips)))
    tblPr.append(tblW)

    tblLayout = OxmlElement("w:tblLayout")
    tblLayout.set(qn("w:type"), "fixed")
    tblPr.append(tblLayout)

    # 設定欄格線（tblGrid）：移除 add_table 產生的舊 tblGrid，再插入新的
    for _old_grid in tbl._tbl.findall(qn("w:tblGrid")):
        tbl._tbl.remove(_old_grid)
    tblGrid = OxmlElement("w:tblGrid")
    for w in col_widths_twips:
        gridCol = OxmlElement("w:gridCol")
        gridCol.set(qn("w:w"), str(w))
        tblGrid.append(gridCol)
    tbl._tbl.insert(tbl._tbl.index(tblPr) + 1, tblGrid)

    for r_idx, (row, cells) in enumerate(zip(rows, row_cells_list)):
        tbl_row = tbl.rows[r_idx]

        # ── SpacerElem 列：精確空白高度，不寫入任何內容 ──────────────────────
        if row and all(isinstance(e, SpacerElem) for e in row):
            _set_row_height(tbl_row, row[0].height_pt, exact=True)
            merged = tbl_row.cells[0]
            for c in range(1, max_cols):
                merged = merged.merge(tbl_row.cells[c])
            _remove_cell_borders(merged)
            _set_cell_margins(merged, top_dxa=0, bottom_dxa=0, left_dxa=0, right_dxa=0)
            continue

        n = len(cells)
        if n == 0:
            continue

        # ── 依 PDF 座標設定最小列高（atLeast，允許內容撐高）───────────────────
        content_elems = [e for e in row if not isinstance(e, SpacerElem)]
        if content_elems:
            row_h = (max(e.bottom for e in content_elems) -
                     min(e.top    for e in content_elems))
            if row_h > 2:
                _set_row_height(tbl_row, row_h, exact=False)

        def _apply_borders(c, has_b, elems):
            if show_borders:
                _set_all_cell_borders(c)
            else:
                _remove_cell_borders(c)
            if has_b or _cell_needs_border(r_idx, elems, bottom_border_set):
                _set_cell_bottom_border(c)

        if n == 1:
            merged = tbl_row.cells[0]
            for c in range(1, max_cols):
                merged = merged.merge(tbl_row.cells[c])
            _set_cell_width(merged, sum(col_widths_twips))
            has_blank = _write_cell(merged, cells[0], page_width)
            _apply_borders(merged, has_blank, cells[0])
            _set_cell_margins(merged)

        elif n == max_cols:
            for c_idx, cell_elems in enumerate(cells):
                cell = tbl_row.cells[c_idx]
                _set_cell_width(cell, col_widths_twips[c_idx])
                has_blank = _write_cell(cell, cell_elems, page_width)
                _apply_borders(cell, has_blank, cell_elems)
                _set_cell_margins(cell)

        else:
            # 以 x 座標對應欄位範圍，取代均分法
            c_used_up_to = 0
            for i, cell_elems in enumerate(cells):
                c_start, c_end = _cell_col_range(cell_elems, col_boundaries_pdf, max_cols)
                c_start = max(c_start, c_used_up_to)
                if c_start >= max_cols:
                    break  # 欄位已耗盡，跳過剩餘格
                if i == n - 1:
                    c_end = max_cols - 1
                c_end = max(c_end, c_start)
                c_end = min(c_end, max_cols - 1)  # 防止越界
                merged = tbl_row.cells[c_start]
                for c in range(c_start + 1, c_end + 1):
                    merged = merged.merge(tbl_row.cells[c])
                span_w = sum(col_widths_twips[c_start:c_end + 1])
                _set_cell_width(merged, span_w)
                has_blank = _write_cell(merged, cell_elems, page_width)
                _apply_borders(merged, has_blank, cell_elems)
                _set_cell_margins(merged)
                c_used_up_to = c_end + 1

    # 補救掃描：確保空白格、spacer row 與未寫入格的段落也套用零行距
    # （_write_cell 只處理有內容的格；Normal 樣式預設有 space_after=8pt）
    _seen_tc: set = set()
    for _tbl_row in tbl.rows:
        for _cell in _tbl_row.cells:
            _tc_id = id(_cell._tc)
            if _tc_id in _seen_tc:
                continue
            _seen_tc.add(_tc_id)
            for _para in _cell.paragraphs:
                _zero_para_spacing(_para)


# ── OCR fallback ──────────────────────────────────────────────────────────────
def _build_ocr_page(doc: Document, pdf_path: str, page_idx: int):
    img      = _page_to_pil(pdf_path, page_idx)
    ocr_text = pytesseract.image_to_string(img, lang="eng", config="--psm 6")
    for line in ocr_text.splitlines():
        doc.add_paragraph(line.strip())


def make_searchable_pdf(pdf_path: str, output_path: str = None) -> bytes:
    """
    Add an OCR text layer to a scanned PDF using ocrmypdf + Tesseract.
    The original image is preserved; text becomes selectable/searchable.
    Returns the searchable PDF as bytes.
    If output_path is given, also writes the file to disk.
    Raises RuntimeError if ocrmypdf is not available.

    Windows note: ocrmypdf uses multiprocessing which requires the
    'spawn' start method. We set it here to avoid fork-related crashes
    when called from a GUI thread.
    """
    try:
        import ocrmypdf as _ocrmypdf
    except ImportError:
        raise RuntimeError(
            "ocrmypdf is not installed. Run: pip install ocrmypdf"
        )

    import io as _io
    import multiprocessing as _mp

    # On Windows (and frozen executables), multiprocessing defaults to
    # 'spawn'. Force it explicitly to prevent 'fork' errors from GUI threads.
    try:
        _mp.set_start_method("spawn", force=False)
    except RuntimeError:
        pass  # already set

    # Write to a temp file first (more reliable than BytesIO on Windows)
    if output_path:
        _ocrmypdf.ocr(
            pdf_path,
            output_path,
            language="eng",
            skip_text=True,
            progress_bar=False,
        )
        return Path(output_path).read_bytes()
    else:
        import tempfile as _tempfile
        tmp_fd, tmp_path = _tempfile.mkstemp(suffix=".pdf")
        os.close(tmp_fd)
        try:
            _ocrmypdf.ocr(
                pdf_path,
                tmp_path,
                language="eng",
                skip_text=True,
                progress_bar=False,
            )
            data = Path(tmp_path).read_bytes()
        finally:
            Path(tmp_path).unlink(missing_ok=True)
        return data


def is_scanned_pdf(pdf_path: str) -> bool:
    """Return True if the first page has no selectable text (likely scanned)."""
    with pdfplumber.open(pdf_path) as pdf:
        if not pdf.pages:
            return False
        return len(pdf.pages[0].extract_words()) == 0


# ── img2table table-structure detection ──────────────────────────────────────
def _ensure_tesseract_path():
    """Add the system Tesseract directory to PATH if not already present."""
    tess_dir = r"C:\Program Files\Tesseract-OCR"
    if os.path.isdir(tess_dir) and tess_dir.lower() not in os.environ.get("PATH", "").lower():
        os.environ["PATH"] = tess_dir + os.pathsep + os.environ.get("PATH", "")


def _build_img2table_page(doc: Document, pdf_path: str, page_idx: int,
                           page_width: float, page_height: float) -> bool:
    """
    Detect table structure from a scanned PDF page using img2table (visual line
    detection), then build Word tables with proper column widths and colspan.

    Returns True when at least one table was detected and rendered.
    Returns False so the caller can fall back to line-level OCR.
    """
    _ensure_tesseract_path()

    try:
        from img2table.document import PDF as _I2TPDF
        from img2table.ocr import TesseractOCR as _TesseractOCR
    except ImportError:
        return False

    try:
        _ocr = _TesseractOCR(lang="eng")
        _pdf = _I2TPDF(src=pdf_path, pages=[page_idx])
        raw_tables = _pdf.extract_tables(
            ocr=_ocr,
            implicit_rows=True,
            borderless_tables=False,
        )
    except Exception:
        return False

    page_tables = raw_tables.get(page_idx, [])
    if not page_tables:
        return False

    TWI = 20
    MARGIN_PT = 0.75 * 72
    usable_pt = max(page_width - 2 * MARGIN_PT, 200.0)

    for tbl_data in page_tables:
        content = tbl_data.content  # OrderedDict: row_idx → [TableCell, ...]
        if not content:
            continue

        # ── 建立欄邊界（所有格的 x1 聯集，排序後為欄起始座標） ──────────────
        all_x1 = sorted({int(c.bbox.x1) for row in content.values() for c in row})
        max_x2  = max(int(c.bbox.x2) for row in content.values() for c in row)
        num_cols = len(all_x1)
        if num_cols == 0:
            continue

        # 欄寬（pixel）→ 比例 → twips
        col_ends_px = all_x1[1:] + [max_x2]
        col_widths_px = [col_ends_px[i] - all_x1[i] for i in range(num_cols)]
        total_px = sum(col_widths_px) or 1
        col_widths_twips = [
            max(round(w / total_px * usable_pt * TWI), 200)
            for w in col_widths_px
        ]

        # ── 去重每列中重複的合併格，計算 colspan ─────────────────────────────
        def _col_idx(x):
            """Return the column index whose left boundary equals x (tolerance 5px)."""
            for i, v in enumerate(all_x1):
                if abs(v - x) <= 5:
                    return i
            return 0

        def _colspan(x1, x2):
            """Number of logical columns spanned by a cell with given pixel x-range."""
            c_start = _col_idx(int(x1))
            # c_end = largest col index whose start is still inside this cell
            c_end = c_start
            for i in range(c_start + 1, num_cols):
                if all_x1[i] < int(x2) - 5:
                    c_end = i
                else:
                    break
            return max(1, c_end - c_start + 1)

        # img2table renders PDFs at ~200 DPI internally
        IMG2TABLE_DPI = 200
        PX_TO_PT = 72.0 / IMG2TABLE_DPI  # 0.36

        # ── 去重每列中重複的合併格（相同 x1 → 同一合併格），計算 colspan ──────
        # 同時收集每列的 row_height_px 供字級估算
        deduped_rows = []
        for row_cells in content.values():
            seen_x1: set = set()
            unique_cells = []
            y1s = [int(c.bbox.y1) for c in row_cells]
            y2s = [int(c.bbox.y2) for c in row_cells]
            row_h_px = max(y2s) - min(y1s) if y2s and y1s else 20
            for cell in row_cells:
                key = int(cell.bbox.x1)
                if key not in seen_x1:
                    seen_x1.add(key)
                    span = _colspan(cell.bbox.x1, cell.bbox.x2)
                    unique_cells.append((cell, span, row_h_px))
            deduped_rows.append(unique_cells)

        # ── 格層級去重：跨列合併產生的重複格，在後續列中以空格取代 ──────────
        # img2table 不支援 rowspan，相同格內容會重複出現在多列。
        # 策略：追蹤前一列「原始」簽名（包含已被壓制的），後續列相同配對 → 清空。
        # 使用原始 row_items 簽名（壓制前），確保跨 3+ 列的 rowspan 也能正確去重。
        prev_row_all_sigs: set = set()
        normalized_rows: list = []
        for row_items in deduped_rows:
            new_row = []
            # 先收集本列所有原始簽名（不論是否被壓制）
            curr_all_sigs = {
                ((cell.value or "").strip(), int(cell.bbox.x1))
                for cell, _s, _h in row_items
                if (cell.value or "").strip()
            }
            for cell, span, row_h_px in row_items:
                sig = ((cell.value or "").strip(), int(cell.bbox.x1))
                if sig[0] and sig in prev_row_all_sigs:
                    # 前一列已出現此格 → 清空（模擬 rowspan）
                    class _EmptyCell:
                        def __init__(self, orig):
                            self.bbox  = orig.bbox
                            self.value = ""
                    new_row.append((_EmptyCell(cell), span, row_h_px))
                else:
                    new_row.append((cell, span, row_h_px))
            normalized_rows.append(new_row)
            prev_row_all_sigs = curr_all_sigs
        deduped_rows = normalized_rows

        if not deduped_rows:
            continue

        # ── 建 Word 表格 ─────────────────────────────────────────────────────
        tbl = doc.add_table(rows=len(deduped_rows), cols=num_cols)
        _remove_all_borders(tbl)

        tblPr = tbl._tbl.tblPr
        tblW = OxmlElement("w:tblW")
        tblW.set(qn("w:type"), "dxa")
        tblW.set(qn("w:w"), str(sum(col_widths_twips)))
        tblPr.append(tblW)
        tblLayout = OxmlElement("w:tblLayout")
        tblLayout.set(qn("w:type"), "fixed")
        tblPr.append(tblLayout)

        for _old in tbl._tbl.findall(qn("w:tblGrid")):
            tbl._tbl.remove(_old)
        tblGrid = OxmlElement("w:tblGrid")
        for w in col_widths_twips:
            gc = OxmlElement("w:gridCol")
            gc.set(qn("w:w"), str(w))
            tblGrid.append(gc)
        tbl._tbl.insert(tbl._tbl.index(tblPr) + 1, tblGrid)

        table_span_twips = sum(col_widths_twips)
        all_written_tc: set = set()  # for zero-spacing sweep only

        for r_idx, row_items in enumerate(deduped_rows):
            tbl_row = tbl.rows[r_idx]
            col_cursor = 0
            # Per-row seen_tc: prevents double-writing within a row (due to
            # python-docx returning the same merged _tc for multiple col indices),
            # but does NOT persist across rows (cross-row _tc sharing after
            # horizontal merges would otherwise cause rows to be silently skipped).
            row_seen_tc: set = set()

            for cell, span, row_h_px in row_items:
                if col_cursor >= num_cols:
                    break
                col_end = min(col_cursor + span - 1, num_cols - 1)

                word_cell = tbl_row.cells[col_cursor]
                # Use the actual _tc element object (not id()) to avoid false
                # positives from Python memory address reuse after GC.
                if word_cell._tc in row_seen_tc:
                    col_cursor = col_end + 1
                    continue

                for c in range(col_cursor + 1, col_end + 1):
                    word_cell = word_cell.merge(tbl_row.cells[c])
                row_seen_tc.add(word_cell._tc)
                all_written_tc.add(id(word_cell._tc))

                span_w = sum(col_widths_twips[col_cursor:col_end + 1])
                _set_cell_width(word_cell, span_w)
                _set_cell_margins(word_cell)
                _set_all_cell_borders(word_cell)

                # ── 字級：依列高估算，clamp 8-14pt ──────────────────────────
                # 標準 PDF 表單列高約 14-20pt，對應字級約 8-10pt；
                # 標題列高約 30-50pt，對應字級約 11-13pt。
                row_h_pt = row_h_px * PX_TO_PT
                font_pt  = max(8.0, min(round(row_h_pt * 0.40), 14))

                # ── 對齊：全寬格置中，否則靠左 ─────────────────────────────
                is_full_width = (span_w >= table_span_twips * 0.85)
                alignment = WD_ALIGN_PARAGRAPH.CENTER if is_full_width else WD_ALIGN_PARAGRAPH.LEFT

                # ── 文字清洗：過濾 OCR 雜訊 ─────────────────────────────────
                # 1. 去除 "None"（img2table 無法辨識時填入的字串）
                # 2. 過濾獨立行中僅 1-2 個字元的雜訊（如 "G"、短亂碼），
                #    但保留已知有意義的縮寫（MA、PD、F、M 等）
                # 已知有意義的短字（不過濾）
                _KEEP_SHORT = {
                    "MA", "PD", "F", "M", "W", "B", "A", "X",
                    "Y", "N", "AT", "OR", "OF", "TO", "IN", "BY",
                    "NO", "1", "2", "3", "HAZ", "DOB", "DOU", "PO",
                    "PAGE", "DATE", "HAIR", "RACE", "EYES", "MALE",
                    "FLED", "FLED", "NONE", "TYPE", "ONLY",
                }

                def _clean_lines(raw: str) -> list:
                    import re as _re
                    result = []
                    for ln in (raw or "").split("\n"):
                        ln = ln.strip()
                        if not ln:
                            continue
                        # 去除行內 standalone "None"（img2table 無值時填入）
                        # 只移除被空格或行首/行尾包圍的 None，避免誤刪正常單字
                        ln = _re.sub(r'(?<!\w)None(?!\w)', '', ln).strip()
                        if not ln:
                            continue
                        upper = ln.upper()
                        # 全大寫字母且 ≤ 4 字元：若不在保留清單則視為 OCR 尾段殘字
                        if (len(ln) <= 4 and upper.replace(" ", "").isalpha() and
                                upper not in _KEEP_SHORT):
                            continue
                        # 移除行尾跟著空格+單一大寫字母的 OCR 雜訊（如 "Department G"）
                        ln = _re.sub(r'\s+[A-Z]$', '', ln).strip()
                        if not ln:
                            continue
                        result.append(ln)
                    return result or [""]

                lines = _clean_lines(cell.value or "")

                for li, line in enumerate(lines):
                    para = word_cell.paragraphs[0] if li == 0 else word_cell.add_paragraph()
                    _zero_para_spacing(para)
                    para.alignment = alignment
                    if line:
                        run = para.add_run(line)
                        run.font.size = Pt(font_pt)

                col_cursor = col_end + 1

        # zero spacing for untouched cells (best-effort)
        seen_for_spacing: set = set()
        for tbl_row in tbl.rows:
            for word_cell in tbl_row.cells:
                tc_id = id(word_cell._tc)
                if tc_id not in seen_for_spacing:
                    seen_for_spacing.add(tc_id)
                    if tc_id not in all_written_tc:
                        for para in word_cell.paragraphs:
                            _zero_para_spacing(para)

    return True


def _is_fullpage_image(images: list, page_width: float, page_height: float) -> bool:
    """True when the only image on the page covers (nearly) the entire page — i.e. a scanned PDF."""
    if len(images) != 1:
        return False
    img = images[0]
    return (img.x0 < 10 and img.top < 10 and
            img.x1 > page_width - 10 and img.bottom > page_height - 10)


def _build_ocr_page_table(doc: Document, pdf_path: str, page_idx: int,
                           page_width: float, page_height: float, dpi: int = 200):
    """
    Rebuild a scanned PDF page as a positioned Word table.

    Priority:
    1. img2table visual table detection (detects grid lines → proper colspan)
    2. Line-level Tesseract OCR fallback (block/par/line groupby, max 6 cols)
    3. Plain-text OCR paragraph fallback
    """
    # ── 優先用 img2table 偵測表格格線結構 ──────────────────────────────────
    if _build_img2table_page(doc, pdf_path, page_idx, page_width, page_height):
        return

    # ── 降級：行級 OCR 聚合 ──────────────────────────────────────────────
    _ensure_tesseract_path()
    pil_img = _page_to_pil(pdf_path, page_idx, dpi=dpi)
    scale = 72.0 / dpi  # pixel → pt

    # psm 11 = sparse text (best for forms with multiple independent regions)
    data = pytesseract.image_to_data(
        pil_img, lang="eng", config="--psm 11",
        output_type=pytesseract.Output.DATAFRAME,
    )
    data = data[(data["conf"] > 0) & (data["text"].str.strip() != "")]

    if data.empty:
        _build_ocr_page(doc, pdf_path, page_idx)
        return

    # Aggregate words → lines using Tesseract's own block/par/line hierarchy.
    line_words: list = []
    for _, grp in data.groupby(["block_num", "par_num", "line_num"], sort=True):
        x0     = grp["left"].min() * scale
        top    = grp["top"].min() * scale
        x1     = (grp["left"] + grp["width"]).max() * scale
        bottom = (grp["top"]  + grp["height"]).max() * scale
        text   = " ".join(grp["text"].astype(str).str.strip()).strip()
        if text:
            line_words.append(Word(text=text, x0=x0, top=top, x1=x1, bottom=bottom))

    if not line_words:
        _build_ocr_page(doc, pdf_path, page_idx)
        return

    rows = _group_rows(line_words, [], [], page_width=page_width)
    _build_page_table(doc, rows, page_width=page_width, max_cols_cap=6)


# ── AI vision fallback (via claude CLI) ──────────────────────────────────────
_AI_LAYOUT_PROMPT = """\
@{image_path}

You are reconstructing a scanned court/legal form as a Word document using tables.
Your output drives the Word table builder directly — layout quality depends on you.

Return ONLY valid JSON. No markdown fences, no explanation.

─── SCHEMA ───────────────────────────────────────────────────────────────────
{
  "sections": [
    {
      "num_cols": <integer 1-8: number of LOGICAL columns in this section>,
      "col_widths_pct": [<pct>, ...],  // optional: relative width of each column, must sum to 100
      "rows": [
        {
          "height_pt": <row height in points; full letter page = 792pt>,
          "cells": [
            {
              "text": "<cell text; use actual newline char for multi-line content>",
              "bold": <true|false>,
              "italic": <true|false>,
              "font_size": <pt, typically 8-12>,
              "colspan": <1 to num_cols>,
              "align": "<left|center|right>",
              "bg_color": "<#RRGGBB if non-white fill, else null>",
              "border_top":    <true|false>,
              "border_bottom": <true|false>,
              "border_left":   <true|false>,
              "border_right":  <true|false>,
              "is_blank": <true if this is an empty fill-in field>,
              "is_image": <true if this cell contains a logo/seal/stamp/graphic (not just text)>,
              "image_region": {
                "x_pct": <left edge as % of page width, 0-100>,
                "y_pct": <top edge as % of page height, 0-100>,
                "w_pct": <width as % of page width>,
                "h_pct": <height as % of page height>
              }
            }
          ]
        }
      ]
    }
  ]
}

─── SECTION RULES ────────────────────────────────────────────────────────────
Split the page into 3-10 LOGICAL SECTIONS (e.g. header, court info, party info,
charges, checkboxes, signature). Each section gets its OWN num_cols.
Different sections may have different num_cols.
Target 4-15 rows per section.

─── ROW RULES ────────────────────────────────────────────────────────────────
One ROW = one horizontal band of the form, NOT one line of OCR text.
Form fields on the same horizontal level → SAME row.
A row with "Name: ___ DOB: ___" is ONE row with 4 cells: [Name:][blank][DOB:][blank]

─── CELL RULES ───────────────────────────────────────────────────────────────
Merge all text that logically belongs together into ONE cell (use \n for line breaks).
A form label + its fill-in blank = 2 adjacent cells in the same row.
Empty fill-in fields: text="", is_blank=true, border_bottom=true.
Checkboxes: text="[ ] Option label" or "[x] Option label".
Do NOT create separate rows for each checkbox option — group them in one cell.
Max 8 columns in any section.

─── IMAGE / LOGO RULES ──────────────────────────────────────────────────────
If a cell contains a logo, seal, stamp, emblem, or graphic image (not just text):
  - Set is_image=true
  - Set image_region with the approximate bounding box as % of page dimensions
  - text can still contain any caption/label text visible near the image
  - The image_region is used to crop and embed the actual image from the scan.

─── WIDTH / HEIGHT RULES ─────────────────────────────────────────────────────
col_widths_pct: set when columns are clearly unequal (e.g. [20,80] for label+value).
height_pt: estimate from the row's visible proportion of the 792pt page.
Tall header rows: 30-50pt. Normal form rows: 15-25pt. Compact rows: 10-14pt.

─── BACKGROUND & BORDERS ─────────────────────────────────────────────────────
bg_color: ONLY set when there is clearly visible non-white shading.
Borders: set true only when a visible line exists on that side.
"""


def _call_claude_sdk_with_image(image_path: str, model: str = "") -> str:
    """
    Call Claude directly via the Anthropic Python SDK with a base64-encoded image.
    Supports both direct API key (ANTHROPIC_API_KEY) and AWS Bedrock SSO.
    Bedrock is used when ANTHROPIC_API_KEY is absent but AWS credentials are available.
    Returns response text, or raises RuntimeError on failure.
    """
    import base64 as _b64
    try:
        import anthropic as _anthropic
    except ImportError:
        raise RuntimeError("anthropic package not installed")

    _prompt_text = _AI_LAYOUT_PROMPT.replace("{image_path}", "").strip()
    _prompt_lines = _prompt_text.splitlines()
    _prompt_clean = "\n".join(l for l in _prompt_lines if not l.strip().startswith("@"))

    with open(image_path, "rb") as _f:
        _img_b64 = _b64.b64encode(_f.read()).decode("utf-8")

    _messages = [{
        "role": "user",
        "content": [
            {
                "type": "image",
                "source": {
                    "type": "base64",
                    "media_type": "image/png",
                    "data": _img_b64,
                },
            },
            {"type": "text", "text": _prompt_clean},
        ],
    }]

    # ── 優先使用 ANTHROPIC_API_KEY；否則嘗試 AWS Bedrock SSO ─────────────────
    if os.environ.get("ANTHROPIC_API_KEY"):
        _model = model or "claude-sonnet-5"
        _client = _anthropic.Anthropic()
        _resp = _client.messages.create(
            model=_model, max_tokens=8192, messages=_messages)
        return _resp.content[0].text

    # AWS Bedrock：model id 格式不同，需轉換
    _bedrock_model_map = {
        "claude-sonnet-5":           "us.anthropic.claude-sonnet-5",
        "claude-sonnet-4-6":         "us.anthropic.claude-sonnet-4-6",
        "claude-opus-5-5":           "us.anthropic.claude-opus-5-5",
        "claude-opus-5":             "us.anthropic.claude-opus-5",
        "claude-haiku-4-5-20251001": "us.anthropic.claude-haiku-4-5-20251001-v1:0",
        "claude-fable-5-1":          "us.anthropic.claude-fable-5-1",
        "claude-fable-5":            "us.anthropic.claude-fable-5",
    }
    _raw_model = model or "claude-sonnet-5"
    _bedrock_id = _bedrock_model_map.get(_raw_model, _raw_model)

    try:
        # Resolve region: env var > AWS config > default us-east-1
        import os as _os
        _region = (
            _os.environ.get("AWS_REGION")
            or _os.environ.get("AWS_DEFAULT_REGION")
            or "us-east-1"
        )
        _client = _anthropic.AnthropicBedrock(aws_region=_region)
        # Use streaming to support larger responses (Bedrock requires streaming
        # for max_tokens > ~21000). Collect the full text via stream context.
        _text_parts = []
        with _client.messages.stream(
            model=_bedrock_id,
            max_tokens=32000,
            messages=_messages,
        ) as _stream:
            for _event in _stream.text_stream:
                _text_parts.append(_event)
        _full_text = "".join(_text_parts)
        if not _full_text:
            raise RuntimeError("Empty streaming response from Bedrock")
        return _full_text
    except RuntimeError:
        raise
    except Exception as _e:
        raise RuntimeError(f"Bedrock call failed: {_e}") from _e


def _call_claude_cli_with_image(image_path: str, model: str = "") -> str:
    """
    Call claude -p CLI with an image attachment using @path syntax in the prompt.
    Same auth mechanism as the existing text-correction flow in gui.py.
    Returns response text, or raises RuntimeError on failure.
    """
    import subprocess
    import shutil
    import tempfile as _tempfile

    env = os.environ.copy()
    if sys.platform == "win32":
        for npm_bin in [
            str(Path(os.path.expanduser("~")) / "AppData" / "Roaming" / "npm"),
            str(Path(os.path.expanduser("~")) / "AppData" / "Local"   / "npm"),
        ]:
            if npm_bin.lower() not in env.get("PATH", "").lower():
                env["PATH"] = npm_bin + os.pathsep + env.get("PATH", "")

    claude_exe = shutil.which("claude", path=env.get("PATH", ""))
    if not claude_exe:
        raise RuntimeError("claude CLI not found")

    prompt = _AI_LAYOUT_PROMPT.replace("{image_path}", image_path.replace("\\", "/"))

    tf = _tempfile.NamedTemporaryFile(
        mode="w", encoding="utf-8", suffix=".txt", delete=False)
    try:
        tf.write(prompt)
        tf.close()
        tmp_path = tf.name

        if sys.platform == "win32":
            model_flag = f' --model "{model}"' if model else ""
            cmd = f'chcp 65001 > nul && type "{tmp_path}" | claude -p{model_flag}'
            result = subprocess.run(
                cmd, shell=True, capture_output=True, env=env, timeout=180)
        else:
            args = [claude_exe, "-p"]
            if model:
                args += ["--model", model]
            with open(tmp_path, encoding="utf-8") as f_in:
                result = subprocess.run(
                    args, stdin=f_in, capture_output=True, env=env, timeout=180)

        if result.returncode != 0:
            err = result.stderr.decode("utf-8", errors="replace")[:300]
            raise RuntimeError(f"claude CLI exited {result.returncode}: {err}")

        return result.stdout.decode("utf-8", errors="replace").strip()
    finally:
        Path(tmp_path).unlink(missing_ok=True)


def _set_cell_shading(cell, hex_color: str):
    """Set cell background fill color (w:shd). hex_color e.g. 'FFD700' (no #)."""
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:shd")):
        tcPr.remove(old)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  hex_color.lstrip("#").upper())
    tcPr.append(shd)


def _set_cell_borders_ai(cell, top: bool, bottom: bool, left: bool, right: bool):
    """Set individual cell borders based on AI-detected border flags."""
    tcPr = cell._tc.get_or_add_tcPr()
    borders = tcPr.find(qn("w:tcBorders"))
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tcPr.append(borders)
    for side, enabled in (("top", top), ("bottom", bottom),
                           ("left", left), ("right", right)):
        for old in borders.findall(qn(f"w:{side}")):
            borders.remove(old)
        b = OxmlElement(f"w:{side}")
        if enabled:
            b.set(qn("w:val"),   "single")
            b.set(qn("w:sz"),    "6")
            b.set(qn("w:space"), "0")
            b.set(qn("w:color"), "000000")
        else:
            b.set(qn("w:val"),   "none")
            b.set(qn("w:sz"),    "0")
            b.set(qn("w:space"), "0")
            b.set(qn("w:color"), "auto")
        borders.append(b)


def _build_ai_section_table(doc: Document, section: dict,
                              col_widths_twips: list, usable_pt: float,
                              pil_img=None, page_width_pt: float = 0,
                              page_height_pt: float = 0):
    """
    Build one Word table for a single AI-returned section dict.
    col_widths_twips: list of twips widths for each column in this section.
    pil_img: optional PIL image of the full page (for cropping logo/seal regions).
    """
    TWI = 20
    num_cols = len(col_widths_twips)
    rows = section.get("rows", [])
    if not rows:
        return

    tbl = doc.add_table(rows=len(rows), cols=num_cols)
    _remove_all_borders(tbl)

    tblPr = tbl._tbl.tblPr
    total_w = sum(col_widths_twips)
    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:type"), "dxa")
    tblW.set(qn("w:w"),    str(total_w))
    tblPr.append(tblW)
    tblLayout = OxmlElement("w:tblLayout")
    tblLayout.set(qn("w:type"), "fixed")
    tblPr.append(tblLayout)

    for _old in tbl._tbl.findall(qn("w:tblGrid")):
        tbl._tbl.remove(_old)
    tblGrid = OxmlElement("w:tblGrid")
    for w in col_widths_twips:
        gc = OxmlElement("w:gridCol")
        gc.set(qn("w:w"), str(w))
        tblGrid.append(gc)
    tbl._tbl.insert(tbl._tbl.index(tblPr) + 1, tblGrid)

    all_written_tc: set = set()  # for zero-spacing sweep only

    for r_idx, ai_row in enumerate(rows):
        tbl_row = tbl.rows[r_idx]
        h_pt = float(ai_row.get("height_pt", 15))
        if h_pt > 2:
            _set_row_height(tbl_row, h_pt, exact=False)

        cells_def = ai_row.get("cells", [])
        col_cursor = 0
        # Per-row seen_tc — same fix as _build_img2table_page:
        # cross-row _tc sharing after horizontal merges must not gate later rows.
        row_seen_tc: set = set()

        for cell_def in cells_def:
            if col_cursor >= num_cols:
                break
            colspan = max(1, int(cell_def.get("colspan", 1)))
            col_end  = min(col_cursor + colspan - 1, num_cols - 1)

            word_cell = tbl_row.cells[col_cursor]
            if word_cell._tc in row_seen_tc:
                col_cursor = col_end + 1
                continue

            for c in range(col_cursor + 1, col_end + 1):
                word_cell = word_cell.merge(tbl_row.cells[c])
            row_seen_tc.add(word_cell._tc)
            all_written_tc.add(id(word_cell._tc))

            span_w = sum(col_widths_twips[col_cursor:col_end + 1])
            _set_cell_width(word_cell, span_w)
            _set_cell_margins(word_cell)

            _set_cell_borders_ai(
                word_cell,
                top    = bool(cell_def.get("border_top",    False)),
                bottom = bool(cell_def.get("border_bottom", False)),
                left   = bool(cell_def.get("border_left",   False)),
                right  = bool(cell_def.get("border_right",  False)),
            )

            bg = cell_def.get("bg_color")
            if bg:
                _set_cell_shading(word_cell, bg)

            # ── Image / Logo region ──────────────────────────────────────
            is_image = bool(cell_def.get("is_image", False))
            img_region = cell_def.get("image_region")
            if is_image and pil_img is not None and img_region and page_width_pt > 0:
                try:
                    iw, ih = pil_img.size
                    x_pct = float(img_region.get("x_pct", 0)) / 100
                    y_pct = float(img_region.get("y_pct", 0)) / 100
                    w_pct = float(img_region.get("w_pct", 10)) / 100
                    h_pct = float(img_region.get("h_pct", 10)) / 100
                    # clamp to image bounds
                    cx0 = max(0, int(x_pct * iw))
                    cy0 = max(0, int(y_pct * ih))
                    cx1 = min(iw, int((x_pct + w_pct) * iw))
                    cy1 = min(ih, int((y_pct + h_pct) * ih))
                    if cx1 > cx0 and cy1 > cy0:
                        cropped = pil_img.crop((cx0, cy0, cx1, cy1))
                        buf = io.BytesIO()
                        cropped.save(buf, format="PNG")
                        buf.seek(0)
                        para = word_cell.paragraphs[0]
                        _zero_para_spacing(para)
                        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        # scale to fit cell width
                        img_w_pt = (cx1 - cx0) * (72.0 / 200)  # assuming 200 DPI render
                        cell_w_pt = span_w / TWI
                        final_w = min(img_w_pt, cell_w_pt * 0.9)
                        para.add_run().add_picture(buf, width=Pt(max(final_w, 36)))
                except Exception:
                    pass  # image crop failed, fall through to text

            # ── Text ─────────────────────────────────────────────────────
            text      = str(cell_def.get("text", ""))
            is_blank  = bool(cell_def.get("is_blank", False))
            bold      = bool(cell_def.get("bold",   False))
            italic    = bool(cell_def.get("italic", False))
            fs        = max(7, min(float(cell_def.get("font_size", 10)), 28))
            align_str = cell_def.get("align", "left").lower()
            alignment = {
                "center": WD_ALIGN_PARAGRAPH.CENTER,
                "right":  WD_ALIGN_PARAGRAPH.RIGHT,
            }.get(align_str, WD_ALIGN_PARAGRAPH.LEFT)

            lines = text.split("\n") if text.strip() else [""]
            for line_idx, line in enumerate(lines):
                # if we already added an image, add text as separate paragraph
                start_idx = 1 if (is_image and img_region) else 0
                if line_idx == 0 and not (is_image and img_region):
                    para = word_cell.paragraphs[0]
                else:
                    para = word_cell.add_paragraph()
                _zero_para_spacing(para)
                para.alignment = alignment
                line = line.strip()
                if is_blank and not line:
                    _insert_formtext_field(para, span_w / TWI)
                elif line:
                    run = para.add_run(line)
                    run.bold      = bold
                    run.italic    = italic
                    run.font.size = Pt(fs)

            col_cursor = col_end + 1

    # Zero spacing for untouched cells
    seen_for_spacing: set = set()
    for tbl_row in tbl.rows:
        for word_cell in tbl_row.cells:
            tc_id = id(word_cell._tc)
            if tc_id not in seen_for_spacing:
                seen_for_spacing.add(tc_id)
                if tc_id not in all_written_tc:
                    for para in word_cell.paragraphs:
                        _zero_para_spacing(para)


def _build_ai_page_table(doc: Document, image_path: str, page_width: float,
                          page_height: float, model: str, pil_img=None):
    """
    Send the page image to Claude via CLI, receive sections-based JSON,
    then build one Word table per section.
    Returns True on success, False on failure (caller falls back to OCR).
    """
    import json as _json

    # Try SDK first (avoids Claude Code permission interception), fall back to CLI.
    raw = None
    _call_errors = []
    for _caller in (_call_claude_sdk_with_image, _call_claude_cli_with_image):
        try:
            raw = _caller(image_path, model=model)
            if raw and len(raw) > 20:
                break
        except Exception as _exc:
            err_msg = f"{_caller.__name__}: {type(_exc).__name__}: {_exc}"
            _call_errors.append(err_msg)
            # Always log SDK errors immediately (even if CLI succeeds)
            if _caller is _call_claude_sdk_with_image:
                try:
                    from pathlib import Path as _Px
                    import datetime as _dtx
                    _lx = _Px.home() / ".pdf2word" / "ai_vision_error.log"
                    _lx.parent.mkdir(parents=True, exist_ok=True)
                    _lx.write_text(
                        f"[{_dtx.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] SDK error: {err_msg}\n",
                        encoding="utf-8",
                    )
                except Exception:
                    pass
            continue
    if not raw:
        # Log the actual failure reason
        try:
            from pathlib import Path as _P
            import datetime as _dt
            _log = _P.home() / ".pdf2word" / "ai_vision_error.log"
            _log.parent.mkdir(parents=True, exist_ok=True)
            _ts = _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            _log.write_text(
                f"[{_ts}] _build_ai_page_table failed (raw={repr(raw)[:80]}):\n"
                + "\n".join(_call_errors) + "\n",
                encoding="utf-8",
            )
        except Exception:
            pass
        return False

    def _ai_fail_log(reason: str):
        try:
            from pathlib import Path as _P2
            import datetime as _dt2
            _log2 = _P2.home() / ".pdf2word" / "ai_vision_error.log"
            _log2.parent.mkdir(parents=True, exist_ok=True)
            _ts2 = _dt2.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            _log2.write_text(
                f"[{_ts2}] {reason}\nraw_len={len(raw)}, raw_preview={repr(raw[:200])}\n",
                encoding="utf-8",
            )
        except Exception:
            pass

    # Extract JSON from response (handles markdown fences, leading text, thinking preambles)
    import re as _re
    stripped = raw.strip()
    if stripped.startswith("```"):
        stripped = stripped.split("```")[1]
        if stripped.startswith("json"):
            stripped = stripped[4:]
        stripped = stripped.strip()
    if not stripped.startswith(("{", "[")):
        _m = _re.search(r'[{\[]', stripped)
        if _m:
            stripped = stripped[_m.start():]
        else:
            _ai_fail_log(f"no JSON start found in raw response")
            return False
    try:
        layout = _json.loads(stripped)
    except Exception as _je:
        try:
            _end = max(stripped.rfind("}"), stripped.rfind("]"))
            if _end > 0:
                layout = _json.loads(stripped[:_end + 1])
            else:
                _ai_fail_log(f"JSON parse failed (no closing brace): {_je}")
                return False
        except Exception as _je2:
            _ai_fail_log(f"JSON parse failed: {_je} / retry: {_je2}")
            return False

    sections = layout.get("sections")
    if not sections:
        flat_rows = layout.get("rows", [])
        if not flat_rows:
            _ai_fail_log(f"JSON parsed OK but no sections/rows. keys={list(layout.keys())}")
            return False
        num_cols = max(
            (sum(c.get("colspan", 1) for c in r.get("cells", [])) for r in flat_rows),
            default=1,
        )
        sections = [{"num_cols": num_cols, "rows": flat_rows}]

    if not sections:
        _ai_fail_log(f"sections empty after processing")
        return False

    TWI = 20
    MARGIN_PT = 0.75 * 72
    usable_pt = max(page_width - 2 * MARGIN_PT, 200.0)

    for section in sections:
        num_cols = max(1, int(section.get("num_cols", 1)))

        # col_widths_pct: optional proportional column widths
        pct_raw = section.get("col_widths_pct") or []
        if len(pct_raw) == num_cols and sum(pct_raw) > 0:
            total_pct = sum(pct_raw)
            col_widths = [
                max(round(p / total_pct * usable_pt * TWI), 200)
                for p in pct_raw
            ]
        else:
            col_w = max(round(usable_pt / num_cols * TWI), 200)
            col_widths = [col_w] * num_cols

        _build_ai_section_table(doc, section, col_widths, usable_pt,
                                pil_img=pil_img,
                                page_width_pt=page_width,
                                page_height_pt=page_height)

    # Log successful AI vision result to conversion.log
    try:
        from pathlib import Path as _P
        import datetime as _dt
        _log = _P.home() / ".pdf2word" / "conversion.log"
        _log.parent.mkdir(parents=True, exist_ok=True)
        _ts = _dt.datetime.now().strftime("%H:%M:%S")
        _entry = f"  [{_ts}] AI vision OK: model={model}, sections={len(sections)}, raw_len={len(raw)}\n"
        _existing = _log.read_text(encoding="utf-8") if _log.exists() else ""
        _log.write_text(_entry + _existing, encoding="utf-8")
    except Exception:
        pass

    return True


# ── main entry ────────────────────────────────────────────────────────────────
def _write_conversion_log(pdf_path: str, entries: list):
    """Write per-page conversion path info to ~/.pdf2word/conversion.log."""
    try:
        from pathlib import Path as _P
        import datetime as _dt
        log_dir = _P.home() / ".pdf2word"
        log_dir.mkdir(parents=True, exist_ok=True)
        ts = _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        lines = [f"[{ts}] {_P(pdf_path).name}"]
        lines += [f"  {e}" for e in entries]
        log_path = log_dir / "conversion.log"
        existing = log_path.read_text(encoding="utf-8") if log_path.exists() else ""
        log_path.write_text("\n".join(lines) + "\n\n" + existing, encoding="utf-8")
    except Exception:
        pass


def convert_pdf_to_word(pdf_path: str, *,
                        ai_vision_model: str = None) -> bytes:
    doc = Document()
    _reset_doc_spacing(doc)

    for p in list(doc.paragraphs):
        p._element.getparent().remove(p._element)

    for section in doc.sections:
        section.left_margin   = Inches(0.75)
        section.right_margin  = Inches(0.75)
        section.top_margin    = Inches(0.75)
        section.bottom_margin = Inches(0.75)

    _log_entries = []
    _ai_vision_disabled = False  # set True after first AI vision failure to skip remaining pages

    with pdfplumber.open(pdf_path) as pdf:
        num_pages = len(pdf.pages)
        # 匹配 PDF 頁面尺寸（避免 Word 預設 Letter/A4 與原始 PDF 不符）
        if pdf.pages:
            fp = pdf.pages[0]
            for section in doc.sections:
                section.page_width  = Pt(fp.width)
                section.page_height = Pt(fp.height)
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

            page_has_borders = _has_table_structure(page)

            # When ai_vision_model is set and page has a full-page background image,
            # always use AI vision — even if words were extracted from an OCR text layer.
            # OCR text layers (added by ocrmypdf) should NOT suppress AI vision analysis.
            is_fullpage = _is_fullpage_image(images, page.width, page.height)
            force_ai_vision = ai_vision_model and is_fullpage and words

            if force_ai_vision or (not words and not form_fields and is_fullpage):
                # 掃描 PDF：整頁是一張圖
                used_ai = False
                if ai_vision_model and not _ai_vision_disabled:
                    pil_img = _page_to_pil(pdf_path, page_idx, dpi=200)
                    # Save temp image next to the exe (or PDF) to ensure read permissions.
                    # Claude Code's sandbox blocks access to system temp dirs.
                    _exe_dir = Path(getattr(sys, "executable", __file__)).parent
                    tmp_img_path = str(_exe_dir / f"_ai_page_{page_idx}.png")
                    try:
                        pil_img.save(tmp_img_path, format="PNG")
                        used_ai = _build_ai_page_table(
                            doc, tmp_img_path,
                            page_width=page.width, page_height=page.height,
                            model=ai_vision_model,
                            pil_img=pil_img,
                        )
                    finally:
                        Path(tmp_img_path).unlink(missing_ok=True)
                if used_ai:
                    src = "forced from OCR-text PDF" if force_ai_vision else "scanned PDF"
                    _log_entries.append(f"page {page_idx+1}: AI vision ({ai_vision_model}) [{src}]")
                else:
                    if ai_vision_model and not _ai_vision_disabled:
                        _ai_vision_disabled = True  # stop retrying on subsequent pages
                    _log_entries.append(f"page {page_idx+1}: OCR/img2table (AI vision {'failed - disabled for remaining pages' if _ai_vision_disabled else 'not configured'})")
                    _build_ocr_page_table(doc, pdf_path, page_idx,
                                          page_width=page.width,
                                          page_height=page.height)
            elif not words and not images and not form_fields:
                _log_entries.append(f"page {page_idx+1}: plain-OCR (no content detected)")
                _build_ocr_page(doc, pdf_path, page_idx)
            else:
                _log_entries.append(f"page {page_idx+1}: text-PDF (pdfplumber, words={len(words)})")
                rows = _group_rows(words, blanks, images,
                                   page_width=page.width,
                                   form_fields=form_fields)
                _build_page_table(doc, rows, page_width=page.width,
                                  separators=separators,
                                  show_borders=page_has_borders)

            if page_idx < num_pages - 1:
                doc.add_page_break()

    _write_conversion_log(pdf_path, _log_entries)

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
    _reset_doc_spacing(doc)
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
