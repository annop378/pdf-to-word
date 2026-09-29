# Changelog

所有版本的重要變更記錄於此。格式依循 [Keep a Changelog](https://keepachangelog.com/zh-TW/1.0.0/)。

---

## [Unreleased] — 2026-09-29

### Added
- **converter.py** `_set_all_cell_borders(cell, color)`：對掃描 PDF 的 img2table 路徑加回灰色細線格線，還原原始表單視覺結構
- **converter.py** `_has_table_structure(page)`：偵測 pdfplumber 頁面是否有向量格線（用於文字型 PDF 的 `show_borders` 開關）
- **converter.py** `_call_claude_sdk_with_image(image_path, model)`：直接透過 Anthropic Python SDK（含 AWS Bedrock SSO 支援）呼叫 Claude 視覺分析，避免 Claude Code 沙箱攔截 subprocess 的檔案讀取權限問題。優先嘗試 `ANTHROPIC_API_KEY`，否則自動使用 `AnthropicBedrock` 走 AWS credential chain；大型請求（`max_tokens=32000`）改用 Bedrock streaming API（`messages.stream`）
- **converter.py** Bedrock inference profile ID 對應表（`us.anthropic.claude-sonnet-5` 等格式）
- **gui.py** `_ensure_ai_vision_deps()`：exe 執行時自動偵測並透過 pip 安裝 `anthropic`、`boto3`，無需手動預裝

### Changed
- **converter.py** `_build_img2table_page()`：`_remove_cell_borders` 改為 `_set_all_cell_borders`，掃描表單轉換結果現在顯示格線
- **converter.py** `_build_ai_page_table()`：
  - 改用 SDK 優先（`_call_claude_sdk_with_image`）+ CLI 備援的雙路徑邏輯
  - JSON 提取強化：支援含 thinking 前言、markdown fences、trailing text 的各種回應格式
  - 修正 `max_tokens=8192` 導致大型表單頁面 JSON 被截斷的問題（改為 `32000` + streaming）
- **converter.py** `_build_page_table()`：新增 `show_borders` 參數；文字型 PDF 偵測到格線時也會顯示邊框
- **converter.py** 跨列 rowspan 去重邏輯：改用原始列的全部簽名（含已壓制的格），正確處理 3 列以上的 rowspan
- **converter.py** `_clean_lines()`：新增 regex 移除行尾跟著空格+單一大寫字母的 OCR 雜訊（如 "Department G"）

### Fixed
- **converter.py** `_build_img2table_page()` / `_build_ai_section_table()`：將 `row_seen_tc` 從 `set of id()` 改為 `set of _tc 物件`，修正 Python GC 記憶體位址重用導致欄位（如 HEIGHT）被誤判為 in_seen 而消失的 bug

### Added
- **converter.py** `make_searchable_pdf()`：使用 `ocrmypdf` 為掃描版 PDF 加入 OCR 文字層，回傳可搜尋的 PDF bytes；`output_path` 可選擇同時寫入磁碟
- **converter.py** `is_scanned_pdf()`：偵測第一頁是否無可選取文字（掃描版判斷）
- **gui.py** `_show_scanned_warning()`：偵測到掃描版 PDF 時**立刻**彈出警告 dialog（在任何 OCR 作業前），提示轉換品質限制（字型/粗體/Logo 無法自動識別），並讓使用者選擇「建立可搜尋 PDF 再轉換（建議）」/「直接轉換」/「取消」；未設 Vision Model 時建議使用 `claude-sonnet-5`（多模態能力較強）
- **gui.py** `_show_searchable_verify()`：可搜尋 PDF 建立完成後的驗證 dialog；使用者可開啟確認 OCR 文字正確後繼續或取消
- **gui.py** `_run()`：重構掃描版流程：①立刻顯示 `_show_scanned_warning` → ②使用者選擇後才開始 OCR → ③OCR 完成顯示 `_show_searchable_verify` → ④確認後用 searchable PDF 轉 Word；OCR 失敗時顯示錯誤 popup 並降級原始 PDF；修正舊版「OCR 先跑、對話框不出現」的問題
- **requirements.txt**：新增 `ocrmypdf>=17.0.0`
- **converter.py** `_is_fullpage_image()`：偵測整頁只有一張覆蓋全版的圖片（掃描 PDF 特徵）
- **converter.py** `_build_ocr_page_table()`：掃描 PDF 專用 OCR 路徑，使用 `pytesseract.image_to_data` 取得每個字的像素座標，換算為 pt 後走與一般 PDF 相同的 `_group_rows` + `_build_page_table` 流程，輸出有欄位結構的表格而非純段落
- **converter.py** `_call_claude_cli_with_image()`：使用與文字修正相同的 `claude -p` CLI 機制，透過 `@{image_path}` 語法將暫存 PNG 附加至 prompt，呼叫 Claude 取回 JSON 版面描述
- **converter.py** `_build_ai_section_table()`：將單一 AI section 字典轉換為一張 Word 表格；支援 `col_widths_pct`（比例欄寬）、`is_blank`（FORMTEXT 欄位）、多行文字（`\n` 分段）
- **converter.py** `_build_ai_page_table()`：重構為 sections-based 架構；每個 section 輸出一張獨立 Word 表格，不同 section 可有不同欄數；向下相容舊版 flat rows JSON；CLI 失敗自動降級 OCR
- **converter.py** `_set_cell_shading()`：設定 Word 儲存格背景填色（`w:shd`）
- **converter.py** `_set_cell_borders_ai()`：依 AI 回傳邊框旗標，分別設定儲存格四側邊框
- **converter.py** `convert_pdf_to_word()`：以 `ai_vision_model` keyword 參數取代舊的 `ai_api_key`/`ai_model`；全頁圖片時儲存暫存 PNG 後交給 `_build_ai_page_table`，無 model 時降級 OCR
- **gui.py**：AI 設定對話框新增「掃描版 PDF 轉換 Model」下拉選單，與現有文字分析 Model 各自獨立；設定值儲存於 `~/.pdf2word/ai_config.json` 的 `claude_vision_model` 欄位
- **main.py**：Web API 從環境變數 `AI_VISION_MODEL` 讀取視覺分析 model 並傳入轉換函式
- **converter.py** `SpacerElem`：新增垂直空白佔位資料類別，記錄平行列之間的間距（`height_pt`），取代舊的 `[]` 空列；配合 `_set_row_height(exact=True)` 在 Word 精確還原該間距
- **converter.py** `_set_cell_margins()`：設定儲存格內部邊距（`w:tcMar`）；預設上下 0、左右 28 twips（≈1.4pt），消除 Word 預設 5.4pt 邊距造成的文字換行與列高膨脹
- **converter.py** `_set_row_height()`：設定列高（`w:trHeight`）；`exact=True` 用於 SpacerElem 精確間距，`atLeast` 用於一般內容列以忠實還原 PDF 行高

### Changed
- **converter.py** `_AI_LAYOUT_PROMPT`：重寫為 sections-based 提示詞；要求 Claude 先辨識邏輯區塊（標題、法院資訊、當事人資訊、指控、簽名等），每個區塊建立獨立欄數的表格，以解決逐行 OCR 切割造成的欄位碎片化、欄寬失真、文字換行等問題；新增 `col_widths_pct`（比例欄寬）與 `is_blank` 欄位
- **converter.py** `_ensure_tesseract_path()`：自動將 `C:\Program Files\Tesseract-OCR` 加入 PATH（避免 img2table/pytesseract 找不到 tesseract）
- **converter.py** `_build_img2table_page()`：使用 `img2table` 偵測掃描頁面的視覺格線，推算欄邊界（唯一 x1 集合）與各格 colspan（bbox 覆蓋的欄數），以比例欄寬建立 Word 表格；依列高估算字級（row_h_pt × 0.40，clamp 8-14pt）；全寬格自動置中；格層級去重以模擬 rowspan；每列獨立 `row_seen_tc`；OCR 文字清洗（過濾 `None` 空值、≤4 字元全大寫殘字如 "ESS"/"TION" 等，保留已知縮寫）；是掃描版轉換的優先路徑
- **converter.py** `_build_ocr_page_table()`：重構為兩階段：①優先呼叫 `_build_img2table_page`，②失敗時降級為行級 Tesseract OCR 聚合（`--psm 11`，block/par/line groupby，`max_cols_cap=6`）
- **converter.py** `_build_page_table()`：新增 `max_cols_cap` 參數；當 > 0 時將 `max_cols` 限制在此值以內（預設 0 = 不限，保持文字型 PDF 原有行為）
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
- **converter.py** `_build_img2table_page()`：修正跨列 `seen_tc` bug — python-docx 做 horizontal merge 後，下一列相同欄位的 `_tc` id 可能與已合併格相同，導致後續列整列內容被跳過（全空）；改為每列獨立 `row_seen_tc`
- **converter.py** `_build_ai_section_table()`：同上，補齊相同修正（同樣使用跨列累積 `seen_tc`，有相同潛在問題）
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
