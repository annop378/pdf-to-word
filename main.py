import os
import tempfile
import uuid
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles

from converter import convert_pdf_to_word

app = FastAPI(title="PDF to Word Converter")

STATIC_DIR = Path(__file__).parent / "static"
TEMPLATES_DIR = STATIC_DIR / "templates"
OUTPUT_DIR = Path(tempfile.gettempdir()) / "pdf2word_output"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")


@app.get("/", response_class=HTMLResponse)
async def index():
    html_file = TEMPLATES_DIR / "index.html"
    return HTMLResponse(content=html_file.read_text(encoding="utf-8"))


@app.post("/convert")
async def convert(file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Only PDF files are accepted.")

    # save upload to a temp file
    suffix = Path(file.filename).suffix
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(await file.read())
        tmp_path = tmp.name

    try:
        docx_bytes = convert_pdf_to_word(tmp_path)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Conversion failed: {exc}")
    finally:
        os.unlink(tmp_path)

    # write output docx
    out_name = f"{uuid.uuid4().hex}.docx"
    out_path = OUTPUT_DIR / out_name
    out_path.write_bytes(docx_bytes)

    stem = Path(file.filename).stem
    return {"download_url": f"/download/{out_name}", "filename": f"{stem}.docx"}


@app.get("/download/{filename}")
async def download(filename: str):
    # prevent path traversal
    if "/" in filename or "\\" in filename or ".." in filename:
        raise HTTPException(status_code=400, detail="Invalid filename.")
    path = OUTPUT_DIR / filename
    if not path.exists():
        raise HTTPException(status_code=404, detail="File not found.")
    return FileResponse(
        path=str(path),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=filename,
    )
