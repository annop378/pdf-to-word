# CODEMAP — pdf-to-word

## 檔案與資料夾

```
pdf-to-word/
├── converter.py          # 核心轉換邏輯（PDF / Word / 圖片 → docx）
├── gui.py                # Tkinter 桌面 GUI + AI 分析 / 修正功能
├── main.py               # FastAPI Web 路由
├── launcher.py           # 啟動 uvicorn 並自動開啟瀏覽器
├── pdf2word.spec         # PyInstaller 打包設定
├── build_exe.bat         # 一鍵打包腳本
├── requirements.txt      # Python 依賴
├── CLAUDE.md             # Claude Code 專案指引
├── README.md             # 使用說明
├── CODEMAP.md            # 本檔案
├── CHANGELOG.md          # 版本變更記錄
└── static/
    └── templates/
        └── index.html    # Web 前端 UI
```

---

## converter.py

### 資料類別

| 類別 | 用途 |
|------|------|
| `Word` | 單一文字詞，含座標、字型、粗體、斜體、底線 |
| `Blank` | 水平線 / 底線，代表填寫空白欄 |
| `ImageElem` | 嵌入圖片，含座標與 bytes |
| `BoxElem` | 空白帶框矩形（勾選框 / 空表格格） |
| `FormFieldElem` | AcroForm widget（文字框、勾選框、下拉） |
| `SeparatorElem` | 全寬水平分隔線，作為列合併的阻斷邊界 |
| `SpacerElem` | 平行列之間的垂直空白佔位，帶精確高度 |

### 常數

| 常數 | 值 | 說明 |
|------|----|------|
| `ROW_TOLERANCE` | 10 pt | 同列垂直中點容差 |
| `COL_GAP` | 55 pt | 水平間距超過此值時切分欄位 |
| `MIN_BLANK_WIDTH` | 20 pt | 空白線最小寬度 |
| `BLANK_CHAR_WIDTH` | 5.5 pt | FORMTEXT 空白字元寬度換算 |

### 提取函式

| 函式 | 輸入 | 輸出 | 說明 |
|------|------|------|------|
| `_page_to_pil` | pdf_path, page_idx, dpi | PIL.Image | 將 PDF 頁面渲染為 PIL 圖片 |
| `_extract_images` | pdf_path, page_idx | `list[ImageElem]` | 從 PyMuPDF 提取嵌入圖片；略過薄水平線與小色塊 |
| `_extract_form_fields` | pdf_path, page_idx | `list[FormFieldElem]` | 從 PyMuPDF 提取 AcroForm widget |
| `_extract_text_and_blanks` | pdfplumber page | `(words, blanks, separators)` | 提取文字、空白線、分隔線 |

### 分組 / 分欄函式

| 函式 | 說明 |
|------|------|
| `_group_rows` | 依垂直重疊將文字、空白、圖片、表單欄位分組成列；右欄元素允許小間距不擴展列高 |
| `_split_cells` | 將列元素依規則切分為儲存格：Blank 自成一格、冒號後接填寫欄時切分、水平間距 > COL_GAP 時切分 |

### 版面建構函式

| 函式 | 說明 |
|------|------|
| `_merge_parallel_rows` | 合併欄結構相同且無間距的連續列；全寬分隔線阻斷合併 |
| `_insert_parallel_spacers` | 在平行列間距 > 1.2× 行高時插入 SpacerElem 保留垂直間距 |
| `_annotate_underlines` | 偵測緊接在文字列下方的 Blank-only 列，將其轉換為底線標記後移除 |
| `_compute_col_widths` | 從各列元素 x 座標取中位數，計算各欄的 twips 寬度 |
| `_cell_col_range` | 依元素 x 範圍對應欄索引 (c_start, c_end) |
| `_infer_alignment` | 依元素相對頁面位置推斷對齊方式（左 / 中 / 右） |
| `_build_page_table` | 主要頁面重建函式：接受分組後的列，輸出 Word 無邊框表格；`max_cols_cap>0` 時限制最大欄數（OCR 路徑使用）；結尾只用 `_seen_tc` 做零行距補救掃描，**不以 `seen_tc` 控制格寫入**，因此不受跨列 `_tc` 共用影響 |

### Word XML 輔助函式

| 函式 | 說明 |
|------|------|
| `_remove_all_borders` | 移除表格所有邊框 |
| `_remove_cell_borders` | 移除儲存格所有邊框 |
| `_set_all_cell_borders` | 對掃描表單加回灰色細線格線（`color` 預設 `"808080"`） |
| `_set_cell_bottom_border` | 設定儲存格底部實線（取代底線） |
| `_set_cell_borders_ai` | 依 AI 回傳旗標設定儲存格四側邊框 |
| `_set_cell_shading` | 設定儲存格背景填色（`w:shd`） |
| `_set_cell_width` | 設定儲存格明確寬度（twips） |
| `_set_cell_margins` | 設定儲存格內部邊距（預設上下 0、左右 28 twips） |
| `_set_row_height` | 設定列高；`exact=True` 用於 SpacerElem，否則 `atLeast` |
| `_zero_para_spacing` | 移除段落前後間距 |
| `_reset_doc_spacing` | 重設 Normal 樣式行距（消除 python-docx 預設 space_after） |
| `_resolve_font` | PDF 字型名稱 → Word 字型名稱對照 |
| `_is_rule_image` | 判斷 ImageElem 是否為薄水平線（非實體圖片） |

### Legacy Forms 插入函式

| 函式 | 說明 |
|------|------|
| `_insert_formtext_field` | 插入 Word Legacy Forms 文字欄位（FORMTEXT） |
| `_insert_formcheckbox_field` | 插入 FORMCHECKBOX |
| `_insert_formdropdown_field` | 插入 FORMDROPDOWN |
| `_write_cell` | 將儲存格元素列表寫入 Word cell，含圖片、文字、Legacy Forms |

### OCR / AI 頁面重建函式

| 函式 | 說明 |
|------|------|
| `_build_ocr_page` | Tesseract OCR 純文字備援：辨識後逐行輸出段落（無表格結構） |
| `_is_fullpage_image` | 判斷頁面是否只有一張覆蓋整頁的圖（掃描 PDF 特徵） |
| `_ensure_tesseract_path` | 自動將系統 Tesseract 路徑加入 PATH（img2table/pytesseract 共用） |
| `_build_img2table_page` | `img2table` 視覺格線偵測：推算欄邊界與 colspan，比例欄寬建 Word 表格；依列高估算字級；格層級去重模擬 rowspan；**每列獨立 `row_seen_tc`** 避免 python-docx horizontal merge 後 `_tc` id 跨列共用誤判 |
| `_build_ocr_page_table` | 兩階段：①`_build_img2table_page` → ②失敗時降級行級 Tesseract OCR（`--psm 11`，block/par/line groupby，`max_cols_cap=6`） |
| `_call_claude_sdk_with_image` | 直接透過 Anthropic Python SDK 呼叫 Claude 視覺分析。優先使用 `ANTHROPIC_API_KEY`；否則走 AWS Bedrock SSO（`AnthropicBedrock`）。大型請求自動切換 streaming（`max_tokens=32000`）避免 Bedrock 非 streaming 限制。Bedrock inference profile ID 對應表內建 |
| `_call_claude_cli_with_image` | 透過 `claude -p` CLI（同文字修正認證），以 `@{image_path}` 語法將 PNG 附加至 prompt，回傳 Claude 回應文字。作為 SDK 失敗時的備援 |
| `_build_ai_section_table` | 將單一 AI section 字典建成一張 Word 表格；支援 `col_widths_pct`、`is_blank`（FORMTEXT）、`\n` 多行文字；每列獨立 `row_seen_tc`（使用 `_tc` 物件而非 `id()`，避免 GC 位址重用誤判）；**`is_image=true`** 時從 `pil_img` 裁切對應百分比區域並嵌入（還原 Logo/印章） |
| `_build_ai_page_table` | 解析 Claude 回傳的 sections JSON（支援含 thinking 前言、markdown fence、trailing text 的各種格式）；SDK 優先＋CLI 備援；每個 section 呼叫 `_build_ai_section_table`；相容舊版 flat rows JSON；失敗回傳 `False` |

### `seen_tc` 使用規範（已知跨列共用問題）

python-docx 在同一列做 horizontal merge 後，下一列相同欄位的 `_tc` 可能與已合併格的 Python 物件 id 相同（GC 釋放舊 `_tc` 後記憶體位址被新 `_tc` 重用）。因此：

- **不可用 `id(_tc)` 跨列累積 `seen_tc`**（位址重用會導致後續列格被誤判為已寫入而跳過，如 HEIGHT 欄位消失）
- **正確做法**：每列用獨立的 `row_seen_tc = set()`，存入 `_tc 物件本身`（不是 `id()`）；跨列僅用 `all_written_tc`（存 `id()`）供結尾零行距補救掃描
- **受影響函式**：`_build_img2table_page`、`_build_ai_section_table`（均已修正）
- **不受影響**：`_build_page_table`（`seen_tc` 只用在結尾補救掃描，不控制格寫入）

### 進入點

| 函式 | 簽名 | 說明 |
|------|------|------|
| `convert_pdf_to_word` | `(pdf_path, *, ai_vision_model=None) -> bytes` | 主要轉換函式；全頁圖時若有 `ai_vision_model` 則走 AI，否則走 OCR |
| `convert_image_to_word` | `(image_path) -> bytes` | 圖片 OCR → 表格 Word |
| `convert_word_to_word` | `(docx_path) -> bytes` | docx → PDF（Word COM）→ 表格 Word |
| `_inject_form_protection` | `(docx_bytes) -> bytes` | 注入 `documentProtection` 讓 Legacy Forms 可填寫 |

---

## gui.py

### 模組層級常數 / 函式

| 名稱 | 說明 |
|------|------|
| `_CONF_DIR` | `~/.pdf2word/`，儲存 AI 設定 |
| `_AI_CONFIG_PATH` | `~/.pdf2word/ai_config.json`，儲存 `claude_model`、`claude_vision_model` |
| `_AI_PROMPT_PATH` | `~/.pdf2word/ai_prompt_override.txt`，自訂分析 Prompt |
| `_PRESET_MODELS` | Claude model 預設清單 |
| `_AI_PROMPT_DEFAULT` | 預設文件分析 Prompt |
| `_AI_FIX_PROMPT` | 拼字修正 Prompt（要求 JSON 輸出） |
| `_ensure_ai_vision_deps` | 偵測並自動安裝 `anthropic`、`boto3`（AI vision 所需）；exe 在新電腦初次執行時使用 |
| `_extract_docx_text` | 從 docx bytes 提取純文字供 AI 分析 |
| `_apply_text_corrections` | 套用 AI JSON 修正至 docx，回傳 `(bytes, applied_count)` |
| `_label_for` | 依副檔名回傳中文格式標籤 |
| `_convert` | `(path, ai_vision_model="") -> bytes`；依副檔名呼叫對應轉換函式 |

### App（主視窗）

| 方法 | 說明 |
|------|------|
| `__init__` | 初始化 `_ai_model_var`（文字分析）、`_ai_vision_model_var`（掃描 PDF）、`_auto_fix_var`；載入設定 |
| `_build` | 建立 UI 元件 |
| `_pick` | 開啟檔案選擇器 |
| `_start` | 觸發轉換（啟動 `_run` 執行緒） |
| `_run` | 在背景執行緒呼叫 `_convert`；完成後回呼 `_done` 或 `_error` |
| `_show_save_result` | 顯示轉換完成對話框，含「開啟檔案」按鈕 |
| `_on_auto_fix_toggle` | 勾選「自動修正拼字」時立即觸發 AI 設定驗證 |
| `_done` | 轉換完成後處理：若勾選自動修正則呼叫 `_run_auto_fix_silent` |
| `_run_auto_fix_silent` | 靜默執行 AI 拼字修正，完成後指向 `_corrected.docx` |
| `_error` | 顯示錯誤訊息 |
| `_load_ai_config` | 從 `ai_config.json` 讀取 `claude_model`、`claude_vision_model` |
| `_save_ai_config` | 將 `claude_model`、`claude_vision_model` 儲存至 `ai_config.json` |
| `_open_ai_settings` | 開啟 AI 設定對話框：文字分析 Model + 掃描版 PDF 轉換 Model 兩個下拉，各自獨立 |
| `_open_prompt_editor` | 開啟自訂 Prompt 編輯器 |
| `_start_ai_analyze` | 開啟 AI 分析 popup（含串流輸出、修正拼字按鈕） |
| `_run_claude_async` | 以背景執行緒呼叫 `claude -p` CLI；支援串流 chunk 回呼 |

---

## main.py（FastAPI Web API）

| 路由 | 方法 | 說明 |
|------|------|------|
| `/` | GET | 回傳 `index.html` |
| `/convert` | POST | 接收 PDF 上傳，呼叫 `convert_pdf_to_word`（從環境變數 `AI_VISION_MODEL` 讀取 vision model），回傳下載 URL |
| `/download/{filename}` | GET | 下載已轉換的 docx；防止路徑穿越 |

---

## launcher.py

| 函式 | 說明 |
|------|------|
| `_open_browser` | 延遲 1.2 秒後開啟瀏覽器至 `http://localhost:8765` |
| `main`（模組層級） | 啟動 uvicorn（port 8765）並在背景執行緒開啟瀏覽器 |

---

## PDF 轉換資料流（文字型 PDF）

```
convert_pdf_to_word(pdf_path)
    ├─ pdfplumber.open → _extract_text_and_blanks → [Word], [Blank], [SeparatorElem]
    ├─ fitz.open → _extract_images → [ImageElem]
    ├─ fitz.open → _extract_form_fields → [FormFieldElem]
    ├─ 過濾與表單欄位重疊的文字 / 空白 / 圖片
    ↓
    _group_rows → [[列元素], ...]
    ↓
    _build_page_table
        ├─ _annotate_underlines（移除底線列、標記需底線的儲存格）
        ├─ _merge_parallel_rows（合併相同欄結構的相鄰列）
        ├─ _insert_parallel_spacers（插入 SpacerElem 保留間距）
        ├─ _compute_col_widths（從 x 座標取中位數計算欄寬）
        └─ 逐列 → _split_cells → 逐格 → _write_cell → Word table
```

## PDF 轉換資料流（掃描版 PDF）

```
convert_pdf_to_word(pdf_path, ai_vision_model="claude-sonnet-5")
    ├─ pdfplumber → words=[], blanks=[], separators=[]
    ├─ fitz → images=[ImageElem(全頁)]
    ├─ _is_fullpage_image → True
    ↓
    有 ai_vision_model？
    ├─ 是：_page_to_pil → 存暫存 PNG
    │       _build_ai_page_table
    │           └─ _call_claude_cli_with_image（claude -p CLI + @image_path）
    │               └─ 解析 JSON {sections:[{num_cols, col_widths_pct, rows:[...]}]}
    │               └─ 每個 section → _build_ai_section_table
    │                   → 獨立 Word 表格（含 _set_cell_shading, _set_cell_borders_ai）
    │       失敗 → 降級 OCR ↓
    └─ 否：_build_ocr_page_table
                └─ pytesseract.image_to_data → bounding box → _group_rows → _build_page_table
```

## AI 功能資料流

### 文字分析 / 拼字修正（gui.py）

```
_start_ai_analyze
    └─ _run_claude_async(full_prompt)
           └─ subprocess: type prompt.txt | claude -p [--model model]
                  ├─ 串流 stdout → on_chunk（即時顯示）
                  └─ 完成 → on_result / on_error

_run_auto_fix_silent / do_fix
    └─ _run_claude_async(_AI_FIX_PROMPT + 文件內容)
           └─ on_fix_done: JSON.parse corrections[]
                  └─ _apply_text_corrections(docx_bytes, corrections)
                         └─ 儲存 <原檔名>_corrected.docx
```

### 掃描版 PDF 視覺分析（converter.py）

```
_call_claude_cli_with_image(image_path, model)
    └─ subprocess: type prompt_with_@image | claude -p [--model model]
           └─ stdout → raw JSON string
_build_ai_page_table
    └─ _call_claude_cli_with_image → parse JSON → Word table
```

---

## 設定檔位置

| 檔案 | 路徑 | 內容 |
|------|------|------|
| AI 設定 | `~/.pdf2word/ai_config.json` | `claude_model`（文字分析）、`claude_vision_model`（掃描 PDF 轉換） |
| 自訂 Prompt | `~/.pdf2word/ai_prompt_override.txt` | 覆蓋預設文件分析 Prompt |
| Tesseract | `C:\Program Files\Tesseract-OCR\tesseract.exe` | 系統安裝路徑（打包版內嵌） |
