# PDF to Word Converter

Upload a PDF (text-based or scanned image), automatically detects layout and tables, and downloads as a `.docx` file.

## Requirements

- Python 3.10+
- [Tesseract OCR 5.x](https://github.com/UB-Mannheim/tesseract/wiki) installed at `C:\Program Files\Tesseract-OCR\tesseract.exe`

## Setup

```bash
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
```

## Run

```bash
venv\Scripts\uvicorn main:app --reload
```

Open browser at http://localhost:8000

## How it works

| Situation | Handling |
|-----------|----------|
| Text-based PDF | `pdfplumber` extracts text with position |
| Scanned / image PDF | `pytesseract` OCR fallback |
| Tables with borders | `img2table` visual table detection |
| Form-style layout | Heuristic row/column grouping by coordinates |
| Embedded images (logos) | Extracted and placed in correct cell |

## Project Structure

```
pdf-to-word/
├── main.py          # FastAPI routes
├── converter.py     # Core conversion logic
├── requirements.txt
└── static/
    └── templates/
        └── index.html
```
