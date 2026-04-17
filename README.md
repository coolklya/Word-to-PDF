# ⚡ PDF 工程整合系統 — Streamlit 版

## 功能說明

| 分頁 | 功能 |
|------|------|
| ⚡ 批次轉檔與合併 | 上傳多個 .doc / .docx，自動智慧排序，可手動調整順序後轉換為 PDF，支援合併輸出或打包為 ZIP |
| ✂ PDF 擷取 / 重組 | 上傳 PDF，輸入頁碼範圍（如 `1, 3-5, 8`）進行擷取與重組後下載 |

## 部署方式（Streamlit Cloud）

1. 將以下三個檔案推送至 GitHub Repository 根目錄：
   - `app.py`
   - `requirements.txt`
   - `packages.txt`

2. 至 [share.streamlit.io](https://share.streamlit.io) 建立新 App，指向你的 repo。

3. Streamlit Cloud 會自動讀取 `packages.txt` 安裝 LibreOffice，以及 `requirements.txt` 安裝 Python 套件。

> **注意**：LibreOffice 安裝約需 2-3 分鐘，首次 Deploy 請耐心等候。

## 本機測試

```bash
pip install streamlit pypdf
# macOS
brew install libreoffice
# Ubuntu / Debian
sudo apt-get install libreoffice

streamlit run app.py
```

## 架構差異（V03 桌面版 vs 本 Streamlit 版）

| 項目 | V03 桌面版 | Streamlit 版 |
|------|-----------|-------------|
| Word 轉 PDF | `win32com.client`（需 Windows + Word） | `LibreOffice --headless`（Linux 相容） |
| GUI 框架 | `tkinter` | `streamlit` |
| 排序方式 | Treeview 拖曳 | ▲▼ 按鈕（session_state） |
| 檔案輸出 | 寫入本機磁碟 | 瀏覽器下載按鈕 |
| 排序邏輯 | `initial_sort_key`（完整沿用） | `initial_sort_key`（完整沿用） |

## 已知限制

- LibreOffice 轉換後的 PDF 排版可能與 Microsoft Word 原生輸出略有差異（字型嵌入、複雜版型）。
- 轉換速度約每頁 1-3 秒，大型標案文件（50+ 頁）請耐心等候進度條。
- 上傳單一檔案最大受限於 Streamlit Cloud 預設（200 MB），可在 `.streamlit/config.toml` 調整：
  ```toml
  [server]
  maxUploadSize = 500
  ```
