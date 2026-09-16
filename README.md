# PDF to Word Converter

將 PDF、Word（`.docx`）或圖片檔轉換為格式保留的 `.docx`，支援文字 PDF、掃描版 PDF（OCR）、表格、嵌入圖片與多欄版面。

提供兩種使用方式：
- **Web 介面**：本地 FastAPI 伺服器 + 瀏覽器操作
- **桌面 GUI**：Tkinter 視窗應用程式（可打包為獨立 `.exe`）

---

## 輸入格式建議

**PDF 的轉換效果最好**，建議優先使用 PDF 作為輸入來源。

| 格式 | 效果 | 原因 |
|------|------|------|
| **PDF（文字型）** | ⭐⭐⭐ 最佳 | PDF 內嵌精確的字元座標、字型、字級資訊，可直接提取版面結構；欄位對齊、縮排、表格偵測準確率最高 |
| PDF（掃描版） | ⭐⭐ 良好 | 經 OCR 辨識，準確度取決於掃描品質；清晰掃描件效果接近文字型 PDF |
| 圖片（JPG / PNG） | ⭐⭐ 良好 | 同樣走 OCR，但圖片解析度與對比度直接影響辨識率 |
| Word（.docx） | ⭐ 有限 | 僅重新整理表格版面，不做完整版面重建；複雜樣式可能流失 |

PDF 格式保留了原始排版的向量資訊（文字位置精確到 pt），程式據此還原欄列結構，因此比圖片 OCR 或 Word 重排更能忠實呈現原始版面。

---

## 功能一覽

| 輸入格式 | 處理方式 |
|----------|----------|
| 文字型 PDF | `pdfplumber` 提取文字與座標位置 |
| 掃描版 / 圖片型 PDF | `pytesseract` OCR 辨識文字 |
| 帶邊框的表格 | `img2table` 視覺化表格偵測 |
| 多欄 / 表單版面 | 依座標進行啟發式欄列分組 |
| 嵌入圖片（如 Logo） | 提取後放置於對應儲存格 |
| Word 檔（`.docx`） | 重新整理表格版面後輸出 |
| 圖片（JPG / PNG 等） | OCR 辨識後輸出為 Word |

---

## 系統需求

- Python 3.10+
- [Tesseract OCR 5.x](https://github.com/UB-Mannheim/tesseract/wiki)（安裝於 `C:\Program Files\Tesseract-OCR\tesseract.exe`）
  - 打包成 EXE 時，Tesseract 會一併內嵌，無需另外安裝

---

## 安裝

```bash
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
```

---

## 使用方式

### Web 介面

```bash
# 方式一：自動開啟瀏覽器（port 8765）
python launcher.py

# 方式二：手動啟動 FastAPI（port 8000）
venv\Scripts\uvicorn main:app --reload
```

開啟瀏覽器前往 http://localhost:8000（或 http://localhost:8765）上傳 PDF，下載轉換完成的 `.docx`。

### 桌面 GUI

```bash
python gui.py
```

點選「Select File」選擇檔案，按「Convert」即可在原始路徑旁產生同名的 `.docx`。

---

## 打包成 EXE

執行以下指令，自動完成打包與壓縮：

```bat
build_exe.bat
```

打包完成後：
- 獨立應用程式資料夾：`dist\PdfToWord\`
- 可發佈的 ZIP 套件：`dist\PdfToWord.zip`

> 打包使用 PyInstaller，設定檔為 `pdf2word.spec`。

---

## 專案結構

```
pdf-to-word/
├── main.py           # FastAPI 路由（Web 模式）
├── launcher.py       # 啟動伺服器並自動開啟瀏覽器
├── gui.py            # Tkinter 桌面 GUI
├── converter.py      # 核心轉換邏輯（PDF / Word / 圖片 → docx）
├── pdf2word.spec     # PyInstaller 打包設定
├── build_exe.bat     # 一鍵打包腳本
├── requirements.txt
└── static/
    └── templates/
        └── index.html
```

---

## 轉換邏輯說明

每一頁輸出為一個無邊框 Word 表格，結構如下：

- **列（rows）**：垂直座標相近的元素歸為同一列（容差 10 pt）
- **欄（cells）**：同列內水平間距超過 40 pt 時切分為不同儲存格
- **空白欄位**：水平線轉換為底線空白
- **圖片**：從 PDF 提取後嵌入對應儲存格
- **文字樣式**：保留粗體、斜體、字型大小
