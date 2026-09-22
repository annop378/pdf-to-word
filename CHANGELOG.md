# Changelog

所有版本的重要變更記錄於此。格式依循 [Keep a Changelog](https://keepachangelog.com/zh-TW/1.0.0/)。

---

## [Unreleased] — 2026-09-22

### Added
- **converter.py** `SpacerElem`：新增垂直空白佔位資料類別，記錄平行列之間的間距（`height_pt`），取代舊的 `[]` 空列；配合 `_set_row_height(exact=True)` 在 Word 精確還原該間距
- **converter.py** `_set_cell_margins()`：設定儲存格內部邊距（`w:tcMar`）；預設上下 0、左右 28 twips（≈1.4pt），消除 Word 預設 5.4pt 邊距造成的文字換行與列高膨脹
- **converter.py** `_set_row_height()`：設定列高（`w:trHeight`）；`exact=True` 用於 SpacerElem 精確間距，`atLeast` 用於一般內容列以忠實還原 PDF 行高

### Changed
- **converter.py** `COL_GAP`：從 40pt 提高至 55pt，減少非必要欄分割（尤其對密排版面）
- **converter.py** `_split_cells()`：冒號分割規則限縮為「僅當冒號後方接 Blank/FormFieldElem/BoxElem 時才分割」，避免正文（如「Note: ...」）被不當切割為多格
- **converter.py** `_annotate_underlines()`：向前搜尋上一行文字時，同時跳過 SpacerElem 列，防止 SpacerElem 誤擋底線偵測
- **converter.py** `_insert_parallel_spacers()`：改以 `[SpacerElem(height_pt=gap)]` 取代 `[]`，攜帶精確間距資訊
- **converter.py** `_build_page_table()`：
  - SpacerElem 列以 `exact` 列高渲染，僅合併儲存格不寫入內容
  - 一般內容列加入 `atLeast` 最小列高，對應 PDF 行高
  - 所有儲存格套用 `_set_cell_margins()`，大幅減少預設 Word 邊距造成的版面偏移
- **converter.py** `_write_cell()`：圖片寬度改用 PDF 原始尺寸（`elem.x1 - elem.x0`），移除舊有 3 英吋上限，大型圖片（Logo、橫幅）不再被截短

### Fixed
- **build_exe.bat** Explorer 開著 `dist\PDFtoWord` 視窗時，PyInstaller 自身的 `shutil.rmtree` 也因 WinError 32 失敗的問題：在清除前透過 PowerShell COM（`Shell.Application`）自動關閉 URL 含「PDFtoWord」的所有 Explorer 視窗，釋放目錄鎖定後再執行 `rd /s /q`
- **gui.py** 轉換完成後自動以 Windows Explorer 開啟資料夾並選取輸出檔案；結果對話框新增「開啟檔案」按鈕，取代原本純文字的 messagebox
- **gui.py** 勾選「轉換後自動修正拼字」時自動彈出 AI 設定對話框（模型選擇與連線驗證），確保 AI 已正確設定再進行轉換
- **gui.py** 「轉換後自動修正拼字」checkbox：勾選後，轉換完成自動呼叫 AI 修正，靜默套用，完成後彈窗指向修正版檔案
- **gui.py** `🔧 修正拼字` 按鈕（AI 分析 popup 內）：以 AI 分析整份文件拼字，輸出 JSON 修正清單，套用後儲存 `<原檔名>_corrected.docx`
- **converter.py** `_reset_doc_spacing()`：重設 `Normal` 樣式的段落間距（`space_before=0, space_after=0, line=240/auto`），消除 python-docx 預設範本的 8pt 段後間距與 1.08 行距
- **converter.py** `_compute_col_widths()`：從 PDF 元素 x 座標取樣，計算各欄寬度（中位數邊界，回傳 twips）
- **converter.py** `_set_cell_width()`：設定儲存格明確絕對寬度（`w:tcW dxa`）
- **converter.py** `_cell_col_range()`：依元素 x 範圍對應欄索引，用於部分欄列的 span 計算
- `CODE_MAP.md`：完整記錄所有檔案、類別、函式及資料流的 code map
- `CLAUDE.md`：專案 AI 協作規範；新增 Changelog 維護規範，要求每次修正後更新 `[Unreleased]` 區塊

### Changed
- **converter.py** `_build_page_table()`：
  - 主迴圈結束後新增補救掃描，確保空白格、spacer row 及未寫入格的段落一律套用零行距
  - 表格由 `autofit` 改為 `fixed` layout，搭配明確 `w:tblGrid` 欄格線
  - 各儲存格設定絕對寬度（`_set_cell_width`），欄寬與原始 PDF 座標對應
  - 部分欄列（`n < max_cols`）改以 x 座標對應欄位邊界，取代舊的均分算法
- **converter.py** `convert_pdf_to_word()`：自動讀取 PDF 第一頁的 width/height，套用至 Word section，確保頁面尺寸一致
- **gui.py** `_extract_docx_text()`：改為逐段落輸出；以 `id(cell._tc)` 去重合併儲存格；正規化多餘空白
- **gui.py** `_apply_text_corrections()`：搜尋改用 regex（`\s+` 彈性匹配）；新增 `not_found` 清單回傳；套用後以 messagebox 提示修正版路徑
- **gui.py** `_AI_FIX_PROMPT`：補充說明每行對應單一段落、不可跨行、保留空格格式
- **build_exe.bat**：PyInstaller 後加入 5 秒等待；zip 步驟改用 `try/catch + exit 1` 明確回報失敗

### Fixed
- **converter.py** `_build_page_table`：修正 `tuple index out of range` 錯誤
  - 移除 `doc.add_table()` 自動建立的重複 `w:tblGrid` 元素
  - 在 `else` 分支加入欄位越界保護：`c_start >= max_cols` 時提前中止，並以 `min(c_end, max_cols - 1)` 防止越界
- 拼字修正套用後開檔無變更的問題（根因：`_extract_docx_text` 的 `" | "` 分隔符混入 `original`，導致 `applied = 0`）
- build_exe.bat 的 `[BUILD] Done.` 訊息即使 zip 失敗也會顯示的問題

---

## [0.4.0] — 2026-09-16

### Added
- README 新增輸入格式建議，說明 PDF 轉換效果最佳的原因

### Changed
- 移除主畫面多餘按鈕，簡化 UI
- 預設 AI 模型改為 `claude-haiku-4-5-20251001`

### Fixed
- 統一 EXE 名稱為 `PDFtoWord`，改用 venv 內的 PyInstaller 打包，解決路徑問題

---

## [0.3.0] — 2026-08-13

### Added
- Tkinter 桌面 GUI（`gui.py`）
- AI 分析文件功能（呼叫 Claude CLI，串流回應）
- AI 設定對話框（model 選擇與驗證）
- 自訂 AI Prompt 編輯器（讀寫 `~/.pdf2word/ai_prompt_override.txt`）
- FastAPI Web 模式（`main.py` + `launcher.py`）
- Web UI（`static/templates/index.html`，純 HTML/CSS/JS）
- PyInstaller 建構設定（`pdf2word.spec`）、一鍵建構腳本（`build_exe.bat`）
- Word → Word 重新排版（透過 `docx2pdf` COM 轉換後再處理）
- Image → Word（Tesseract OCR）

---

## [0.1.0] — 2026-07-17

### Added
- 初始版本：PDF → Word 轉換引擎（`converter.py`）
- 以 pdfplumber 提取文字座標、PyMuPDF 提取圖片與表單欄位
- 版面分析：`_group_rows`、`_split_cells`、`_merge_parallel_rows`
- 無框線 Word 表格重建版面
- AcroForm 表單欄位對應 Word Legacy Forms
- Tesseract OCR 備援（掃描 PDF）
- `_inject_form_protection`：注入 Word 表單保護設定
