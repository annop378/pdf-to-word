# pdf-to-word — Code Map

## 專案根目錄結構

```
pdf-to-word/
├── converter.py          轉換引擎（核心邏輯）
├── gui.py                Tkinter 桌面 GUI（PyInstaller 進入點）
├── launcher.py           Web 模式進入點：啟動 uvicorn + 開啟瀏覽器
├── main.py               FastAPI Web 應用程式（路由）
├── requirements.txt      Python 套件依賴
├── pdf2word.spec         PyInstaller 建構設定
├── build_exe.bat         Windows 一鍵建構腳本
├── README.md             使用說明與架構文件
└── static/
    └── templates/
        └── index.html    Web UI（純 HTML/CSS/JS，無外部框架）
```

---

## converter.py — 轉換引擎

### 資料類別（PDF 元素表示）

| 類別 | 說明 |
|---|---|
| `Word` | 單一文字，含座標、粗體、斜體、字號、底線、字型名稱 |
| `Blank` | 水平空白底線（填寫欄位以圖形呈現） |
| `ImageElem` | 嵌入圖片，含邊界框與原始位元組 |
| `BoxElem` | 空邊框矩形（勾選框或空白表格格） |
| `FormFieldElem` | AcroForm Widget（文字框、勾選框、單選、下拉選單） |
| `SeparatorElem` | 全寬水平分隔線，作為合併列的屏障 |
| `SpacerElem` | 垂直空白佔位元素，記錄平行列之間的間距（height_pt），用於精確還原 Word 列高 |

### 提取函式（PDF → 元素清單）

| 函式 | 說明 |
|---|---|
| `_page_to_pil(pdf_path, idx, dpi)` | 用 PyMuPDF 將頁面渲染為 PIL 圖片（OCR 備援） |
| `_extract_images(pdf_path, page_idx)` | 透過 PyMuPDF 抓取嵌入圖片；過濾細線與勾選框圖示 |
| `_extract_form_fields(pdf_path, page_idx)` | 透過 PyMuPDF 提取 AcroForm Widget |
| `_extract_text_and_blanks(page)` | 從 pdfplumber 頁面提取 `Word`、`Blank`、`SeparatorElem`；字元層級底線偵測 |

### 版面分析函式

| 函式 | 說明 |
|---|---|
| `_group_rows(words, blanks, images, page_width, form_fields)` | 依垂直距離（10 pt 容差）將元素分組為水平列；注入圖片與表單欄位 |
| `_split_cells(row_elems)` | 將列元素分割為儲存格：每個 `Blank` 獨立成格（除非與文字重疊），`FormFieldElem`/`BoxElem` 各自成格，文字以 > 40 pt 間距或冒號後切分 |
| `_cells_parallel(cells_a, cells_b)` | 若兩列欄數相同且 x0 位置差距在 20 pt 內則回傳 True |
| `_merge_parallel_rows(rows, separators)` | 合併結構相同且無間距/屏障的連續列 |
| `_insert_parallel_spacers(rows, splits)` | 在間距超過 1.2× 行高的平行列之間插入空白列 |
| `_annotate_underlines(rows, splits)` | 識別並移除代表上一列圖形底線的純 Blank 列，記錄需加下框線的儲存格 |
| `_infer_alignment(elems, page_width)` | 依元素 x 座標相對頁寬推斷 LEFT/CENTER/RIGHT 對齊 |
| `_compute_col_widths(row_cells_list, max_cols, content_x_min, content_x_max, usable_width_pt)` | 從 max_cols 列取樣 x 座標，計算各欄寬度（twips）；回傳 `(widths_twips, boundaries_pdf)` |
| `_set_cell_width(cell, width_twips)` | 設定儲存格明確寬度（`w:tcW dxa`） |
| `_set_cell_margins(cell, top_dxa, bottom_dxa, left_dxa, right_dxa)` | 設定儲存格內部邊距（twips）；預設上下 0、左右 28 twips（≈1.4pt），消除 Word 預設 5.4pt 邊距的換行與膨脹效果 |
| `_set_row_height(tbl_row, height_pt, exact)` | 設定列高；`exact=True` 用於 SpacerElem（精確高度），否則用 atLeast（最小高度） |
| `_cell_col_range(cell_elems, col_boundaries, max_cols)` | 依元素 x 範圍對應 `(c_start, c_end)` 欄索引，用於部分欄列的 span 計算 |

### docx 寫入輔助函式

| 函式 | 說明 |
|---|---|
| `_remove_all_borders(table)` | 移除所有表格框線（OOXML） |
| `_remove_cell_borders(cell)` | 移除儲存格四邊框線 |
| `_set_cell_bottom_border(cell)` | 為儲存格加下框線（代表底線） |
| `_zero_para_spacing(para)` | 移除段落前後間距，設為單倍行距 |
| `_is_rule_image(elem)` | 若圖片為細橫線（高度 < 12 pt，寬高比 > 10）則回傳 True |
| `_insert_formcheckbox_field(para)` | 插入 Word Legacy Forms FORMCHECKBOX 欄位 |
| `_insert_formdropdown_field(para)` | 插入 Word Legacy Forms FORMDROPDOWN 欄位 |
| `_insert_formtext_field(para, width_pts)` | 插入 Word Legacy Forms FORMTEXT 欄位，依空白寬度縮放 |
| `_write_cell(cell, elems, page_width)` | 將元素清單寫入 Word 表格儲存格 |
| `_resolve_font(pdf_fontname)` | 透過 `_FONT_MAP` 將 PDF 字型名稱對應至 Word 字型名稱 |
| `_build_page_table(doc, rows, page_width, separators)` | 為單一頁面建立無框線 Word 表格，協調合併、間距、底線與儲存格寫入 |
| `_build_ocr_page(doc, pdf_path, page_idx)` | 備援：渲染頁面為圖片，執行 Tesseract OCR，加入純文字段落 |
| `_inject_form_protection(docx_bytes)` | 後處理 .docx ZIP，在 `word/settings.xml` 注入表單保護設定 |

### 公開 API

| 函式 | 說明 |
|---|---|
| `convert_pdf_to_word(pdf_path: str) -> bytes` | 主進入點：開啟 PDF，處理每頁，回傳 .docx bytes |
| `convert_image_to_word(image_path: str) -> bytes` | 開啟圖片，執行 Tesseract，建立 Word 表格 |
| `convert_word_to_word(docx_path: str) -> bytes` | docx→PDF（Word COM）→ `convert_pdf_to_word` |

### 重要常數

| 名稱 | 值 | 意義 |
|---|---|---|
| `ROW_TOLERANCE` | 10 pt | 元素歸為同一列的垂直距離容差 |
| `COL_GAP` | 55 pt | 分割列為不同儲存格的最小水平間距 |
| `MIN_BLANK_WIDTH` | 20 pt | 判定為 Blank 的最小寬度 |
| `BLANK_CHAR_WIDTH` | 5.5 pt | 用於計算 FORMTEXT 欄位大小的平均字元寬度 |

---

## main.py — FastAPI Web 應用程式

| 符號 | 說明 |
|---|---|
| `app` | FastAPI 實例，掛載 `/static` |
| `OUTPUT_DIR` | 暫存目錄 `%TEMP%/pdf2word_output` |
| `index()` — `GET /` | 回傳 `index.html` |
| `convert(file)` — `POST /convert` | 接收 PDF 上傳，呼叫 `convert_pdf_to_word`，回傳 `{download_url, filename}` |
| `download(filename)` — `GET /download/{filename}` | 路徑穿越防護，提供 .docx 下載 |

---

## launcher.py — Web 模式啟動器

| 符號 | 說明 |
|---|---|
| `_open_browser()` | 延遲 1.2 秒後開啟 `http://127.0.0.1:8765` |
| `__main__` | 啟動瀏覽器執行緒，執行 `uvicorn.run(app, host="127.0.0.1", port=8765)` |

---

## gui.py — Tkinter 桌面 GUI

### 模組層級常數與函式

| 符號 | 說明 |
|---|---|
| `_CONF_DIR` | `~/.pdf2word/` 設定目錄 |
| `_AI_CONFIG_PATH` | `~/.pdf2word/ai_config.json` |
| `_AI_PROMPT_PATH` | `~/.pdf2word/ai_prompt_override.txt` |
| `_AI_PROMPT_DEFAULT` | 預設 AI 分析 Prompt（文法+書籤建議） |
| `_AI_FIX_PROMPT` | AI 修正拼字 Prompt（JSON 輸出格式） |
| `_PRESET_MODELS` | 預設 Claude model ID 清單 |
| `_extract_docx_text(docx_bytes)` | 提取 .docx 純文字供 AI 分析 |
| `_apply_text_corrections(docx_bytes, correction_list)` | 將 AI JSON 修正套用至 docx，回傳 `(corrected_bytes, applied_count)` |
| `_label_for(path)` | 回傳檔案格式對應的模式標籤 |
| `_convert(path)` | 依副檔名分派至對應的轉換函式 |

### App 類別（`tk.Tk` 子類別）

| 方法 | 說明 |
|---|---|
| `__init__` | 設定視窗、載入 AI 設定、建立 UI |
| `_build` | 建立所有 Widget：檔案選擇列、模式標籤、轉換按鈕、狀態列、AI 分析按鈕 |
| `_pick` | 開啟檔案選擇對話框 |
| `_start` | 停用 UI，在 daemon thread 啟動 `_run` |
| `_run(src)` | 呼叫 `_convert(src)`，寫入 .docx，完成後呼叫 `_done` 或 `_error` |
| `_done(out_path)` | 重新啟用 UI，顯示成功訊息框，啟用 AI 分析按鈕 |
| `_error(msg)` | 重新啟用 UI，顯示錯誤訊息框 |
| `_load_ai_config` | 讀取 `~/.pdf2word/ai_config.json` |
| `_save_ai_config` | 寫入 `~/.pdf2word/ai_config.json` |
| `_open_ai_settings` | Modal 對話框：選取/驗證 Claude model |
| `_open_prompt_editor` | Modal 文字編輯器：讀/寫 `~/.pdf2word/ai_prompt_override.txt` |
| `_start_ai_analyze` | 開啟 AI 分析 popup（含「開始分析」與「修正拼字」按鈕） |
| `_run_claude_async` | 在 daemon thread 執行 `claude -p`，串流 stdout 至 callback |

---

## 資料流

### PDF 轉換（Web 模式）

```
瀏覽器上傳 .pdf
  → POST /convert
  → convert_pdf_to_word(tmp_path)
      ├─ pdfplumber: _extract_text_and_blanks()  → Word, Blank, SeparatorElem
      ├─ PyMuPDF:   _extract_images()            → ImageElem
      ├─ PyMuPDF:   _extract_form_fields()       → FormFieldElem
      ├─ 過濾表單欄位框內的元素
      ├─ [無文字] → _build_ocr_page() (Tesseract)
      └─ [正常路徑] → _group_rows()
                       → _build_page_table()
                           ├─ _annotate_underlines()
                           ├─ _merge_parallel_rows() → _split_cells()
                           ├─ _insert_parallel_spacers()
                           └─ 每列每格 → _write_cell()
      → doc.save() → docx_bytes
      → _inject_form_protection()
  → 回傳 {download_url, filename}
```

### AI 修正拼字流程（桌面 GUI）

```
使用者點「🔧 修正拼字」
  → do_fix()
      ├─ 讀取 _last_docx_path 的 bytes
      ├─ _extract_docx_text() 提取純文字
      ├─ 組合 _AI_FIX_PROMPT + 文件內容
      ├─ _run_claude_async() → claude -p (JSON 輸出)
      ├─ on_fix_result():
      │     ├─ 解析 JSON corrections
      │     ├─ _apply_text_corrections(docx_bytes, corrections)
      │     └─ 儲存 <原檔名>_corrected.docx
      └─ 顯示套用項目數與儲存路徑
```

---

## 設定檔

| 檔案 | 位置 | 用途 |
|---|---|---|
| `requirements.txt` | 專案根目錄 | Python 套件依賴 |
| `pdf2word.spec` | 專案根目錄 | PyInstaller 設定（進入點 `gui.py`，打包 Tesseract） |
| `build_exe.bat` | 專案根目錄 | 一鍵建構腳本 |
| `~/.pdf2word/ai_config.json` | 使用者目錄 | 執行期 AI 設定（Claude model ID） |
| `~/.pdf2word/ai_prompt_override.txt` | 使用者目錄 | 自訂 AI 分析 Prompt（覆蓋預設值） |

### 硬編碼路徑

- `C:\Program Files\Tesseract-OCR\tesseract.exe` — 系統 Tesseract 備援（`converter.py`）
- `C:\Program Files\Tesseract-OCR` — PyInstaller 打包來源（`pdf2word.spec`）

---

## 進入點

| 進入點 | 指令 | 模式 |
|---|---|---|
| `launcher.py` | `python launcher.py` | Web：uvicorn `127.0.0.1:8765` |
| `main.py` | `uvicorn main:app --reload` | Web：FastAPI 開發模式（port 8000） |
| `gui.py` | `python gui.py` | 桌面 GUI（Tkinter） |
| `dist/PDFtoWord/PDFtoWord.exe` | 直接執行 | 獨立桌面 GUI（無需 Python） |

---

## 依賴套件

| 套件 | 用途 |
|---|---|
| `pdfplumber` | 字元層級文字提取（含座標）、線段/矩形偵測 |
| `PyMuPDF` (`fitz`) | 圖片提取、表單 Widget 提取、頁面渲染（OCR 備援） |
| `pytesseract` | OCR 備援（掃描 PDF 與圖片輸入） |
| `Pillow` | OCR 與圖片元素處理 |
| `python-docx` | 建立輸出 .docx 檔案 |
| `fastapi` + `uvicorn` | Web 伺服器（僅 Web 模式，EXE 不包含） |
| `docx2pdf` | Word→PDF 轉換（透過 Microsoft Word COM） |
| `tkinter` | 桌面 GUI（標準庫） |
