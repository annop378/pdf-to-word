# CLAUDE.md — pdf-to-word

## 首要閱讀

**修改任何程式碼前，請先閱讀 [CODE_MAP.md](CODE_MAP.md)。**

CODE_MAP.md 包含：
- 完整檔案與資料夾清單及用途
- 所有類別、函式與其關係
- PDF 轉換資料流（全流程）
- AI 修正拼字流程
- 設定檔位置與硬編碼路徑

---

## 語言

所有回答一律使用繁體中文。專有名詞（如 API、function、variable 等）可保留英文。

## 核心原則

- 閱讀相關檔案後再編輯。
- 不確定時，先明確說明假設再進行。
- 保持修改最小化，不超出任務範圍。
- 不發明不存在的架構；只描述程式碼中實際存在的內容。
- 每一步都必須可解釋。
- 每次修改後驗證輸出。

## 關鍵檔案對應

| 要修改的功能 | 主要檔案 |
|---|---|
| PDF/圖片/Word 轉換邏輯 | `converter.py` |
| 桌面 GUI 與 AI 分析功能 | `gui.py` |
| Web API 路由 | `main.py` |
| Web 啟動（uvicorn + 瀏覽器） | `launcher.py` |
| Web 前端 UI | `static/templates/index.html` |
| PyInstaller 打包設定 | `pdf2word.spec` |

## AI 修正功能說明

`gui.py` 包含兩個 AI 功能：

1. **AI 分析文件**（`_start_ai_analyze` → `do_analyze`）
   - 使用 `_AI_PROMPT_DEFAULT` 或自訂 Prompt
   - 回傳文法/書籤分析文字

2. **修正拼字**（`_start_ai_analyze` → `do_fix`）
   - 使用 `_AI_FIX_PROMPT`（要求 JSON 輸出）
   - 解析 `corrections` JSON 陣列
   - 呼叫 `_apply_text_corrections()` 套用修正
   - 儲存為 `<原檔名>_corrected.docx`

## 日誌與可追蹤性

- 所有產出的 `.exe` 應包含適當的日誌。
- 日誌需包含足夠的上下文資訊以追蹤錯誤。

## Changelog 維護

**每次完成修正或新增功能後，必須更新 [CHANGELOG.md](CHANGELOG.md)。**

- 尚未發版的變更記錄於 `[Unreleased]` 區塊。
- 依類別分類：`Added`（新功能）、`Changed`（改動）、`Fixed`（錯誤修正）。
- 描述要具體：說明修正的問題或新增的行為，不要只寫「修正 bug」。
- 格式遵循 [Keep a Changelog](https://keepachangelog.com/zh-TW/1.0.0/)。

## README 與 requirements.txt 確認

**每次完成修正或新增功能後，主動確認是否需要同步更新：**

- **[README.md](README.md)**：使用說明、功能描述、架構圖、已知限制等是否因此次變更而需更新。
- **[requirements.txt](requirements.txt)**：若新增、移除或升級了任何 Python 套件，必須同步更新此檔案，確保 `pip install -r requirements.txt` 可重現環境。

## 變更影響確認

- 修改一處前，主動確認其他相關位置是否需要同步修改。
- 包含：重複邏輯、平行實作、共用介面、設定、錯誤處理、日誌。
