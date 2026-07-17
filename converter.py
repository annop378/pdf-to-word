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
from dataclasses import dataclass, field

import fitz
import pdfplumber
import pytesseract
from PIL import Image as PILImage
from docx import Document
from docx.shared import Pt, Inches
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

ROW_TOLERANCE    = 10   # 垂直中點容差（pt）
COL_GAP          = 40
MIN_BLANK_WIDTH  = 20
BLANK_CHAR_WIDTH = 5.5


# ── data classes ──────────────────────────────────────────────────────────────
@dataclass
class Word:
    text:   str
    x0:     float
    top:    float
    x1:     float
    bottom: float
    bold:   bool  = False
    italic: bool  = False
    size:   float = 11.0


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
            # PyMuPDF rect uses top-origin (y0=top, y1=bottom)
            images.append(ImageElem(
                x0=rect.x0, top=rect.y0,
                x1=rect.x1, bottom=rect.y1,
                data=base["image"],
                ext=base.get("ext", "png"),
            ))
    doc.close()
    return images


def _extract_text_and_blanks(page) -> tuple:
    """Extract Words and Blanks from a pdfplumber page."""
    ph = page.height
    words = []
    for w in page.extract_words(extra_attrs=["fontname", "size"],
                                 use_text_flow=False,
                                 keep_blank_chars=False):
        fn = w.get("fontname") or ""
        words.append(Word(
            text   = w["text"],
            x0     = w["x0"],   top    = w["top"],
            x1     = w["x1"],   bottom = w["bottom"],
            bold   = "Bold"   in fn or "bold"   in fn.lower(),
            italic = "Italic" in fn or "Oblique" in fn or "italic" in fn.lower(),
            size   = float(w.get("size") or 11),
        ))

    blanks = []
    for ln in page.lines:
        lx0, lx1 = min(ln["x0"], ln["x1"]), max(ln["x0"], ln["x1"])
        ly0, ly1 = min(ln["y0"], ln["y1"]), max(ln["y0"], ln["y1"])
        if (ly1 - ly0) < 3 and (lx1 - lx0) >= MIN_BLANK_WIDTH:
            blanks.append(Blank(x0=lx0, top=ph - ly1,
                                x1=lx1, bottom=ph - ly0))
    return words, blanks


# ── row grouping ──────────────────────────────────────────────────────────────
def _group_rows(words, blanks, images, page_width=0) -> list:
    """
    Group text+blanks into rows by vertical proximity.
    Images are NOT used to drive row boundaries (a tall logo would absorb
    many text rows into one). Instead each image is injected into the row
    whose vertical midpoint it best overlaps.
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

    # --- step 3: sort each row by x0 ----------------------------------------
    return [sorted(row, key=lambda e: e.x0) for row in rows]


# ── column splitting ──────────────────────────────────────────────────────────
def _split_cells(row_elems) -> list:
    if not row_elems:
        return []
    cells, current = [], [row_elems[0]]
    for elem in row_elems[1:]:
        if elem.x0 - current[-1].x1 > COL_GAP:
            cells.append(current)
            current = [elem]
        else:
            current.append(elem)
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


def _write_cell(cell, elems):
    para = cell.paragraphs[0]
    for elem in elems:
        if isinstance(elem, Blank):
            chars = max(4, int((elem.x1 - elem.x0) / BLANK_CHAR_WIDTH))
            run = para.add_run(" " * chars)
            run.underline = True

        elif isinstance(elem, ImageElem):
            width_pts = elem.x1 - elem.x0
            width_in  = max(0.3, min(width_pts / 72, 3.0))
            buf = io.BytesIO(elem.data)
            try:
                run = para.add_run()
                run.add_picture(buf, width=Inches(width_in))
            except Exception:
                pass  # skip unreadable images

        else:  # Word
            run = para.add_run(elem.text + " ")
            run.bold      = elem.bold
            run.italic    = elem.italic
            run.font.size = Pt(max(7, min(float(elem.size), 28)))


# ── page builder ─────────────────────────────────────────────────────────────
def _build_page_table(doc: Document, rows: list):
    if not rows:
        return

    row_cells_list = [_split_cells(row) for row in rows]
    max_cols = max((len(c) for c in row_cells_list), default=1)

    tbl = doc.add_table(rows=len(rows), cols=max_cols)
    _remove_all_borders(tbl)

    tblPr = tbl._tbl.tblPr
    tblW  = OxmlElement("w:tblW")
    tblW.set(qn("w:type"), "pct")
    tblW.set(qn("w:w"),    "5000")
    tblPr.append(tblW)

    for r_idx, (row, cells) in enumerate(zip(rows, row_cells_list)):
        tbl_row = tbl.rows[r_idx]
        n = len(cells)
        if n == 0:
            continue

        if n == 1:
            merged = tbl_row.cells[0]
            for c in range(1, max_cols):
                merged = merged.merge(tbl_row.cells[c])
            _write_cell(merged, cells[0])
            _remove_cell_borders(merged)

        elif n == max_cols:
            for c_idx, cell_elems in enumerate(cells):
                cell = tbl_row.cells[c_idx]
                _write_cell(cell, cell_elems)
                _remove_cell_borders(cell)

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
                _write_cell(merged, cell_elems)
                _remove_cell_borders(merged)
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
            words, blanks = _extract_text_and_blanks(page)
            images        = _extract_images(pdf_path, page_idx)

            if not words and not images:
                _build_ocr_page(doc, pdf_path, page_idx)
            else:
                rows = _group_rows(words, blanks, images, page_width=page.width)
                _build_page_table(doc, rows)

            if page_idx < num_pages - 1:
                doc.add_page_break()

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()
