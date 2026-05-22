# 電子謄本結構化（tw-transcript-structurer）

把全國地政電子謄本系統下載的 PDF 自動解析成結構化 JSON 資料 — 透過 OCR 識別欄位、整理土地/建物清單、產出可程式化處理的 JSON。

> **僅供台灣地區使用**。輸入來源為內政部「全國地政電子謄本系統」下載的第一/二類謄本 PDF。

## 主要功能

| 功能 | 說明 |
|---|---|
| 謄本 PDF OCR 解析 | 使用 RapidOCR 解析掃描型謄本 PDF |
| 結構化輸出 | 土地標示部、所有權部、他項權利部 → JSON |
| 多筆合併 | 同案件多份謄本自動合併成一份 data.json |
| GUI 介面 | tkinter 介面方便操作 |

## 環境需求

- **作業系統**：Windows 10 / 11
- **Python**：3.9+

## 從原始碼執行

```bash
git clone https://github.com/windskyshao/tw-transcript-structurer.git
cd tw-transcript-structurer
pip install rapidocr_onnxruntime onnxruntime pyclipper PyMuPDF Pillow
python "GUI_transcript_pdf 1141021-01.py"
```

## 打包成 .exe（PyInstaller）

spec 採動態抓 `rapidocr_onnxruntime` 套件路徑，**不寫死任何電腦的路徑**，所以在任何裝好 RapidOCR 的環境都能打包：

```bash
pyinstaller "GUI_transcript_pdf 1141021-01.spec" -y
```

完成後 `dist\電子謄本結構化v1.5b.exe` 即可使用（exe 名稱由 spec 內定義）。

## 與「地籍資料查詢系統」的關係

本工具是 [tw-land-tools](https://github.com/windskyshao/tw-land-tools)（地籍資料查詢系統主程式）的附屬工具。主程式有「電子謄本結構化」按鈕會自動偵測同目錄下最新版本的 `電子謄本結構化v*.exe` 並啟動。

也可單獨使用，不依賴主程式。

## 工作流程

```
電子謄本 PDF
  → 本工具 OCR 解析
  → 輸出結構化 JSON（data.json / *_filtered.json）
  → 主程式或其他工具讀取 JSON 進行後續處理
```

## 注意事項

- 本工具只解析使用者本機已存在的 PDF，不會連任何網站
- OCR 結果可能有誤識別，**法律依據以官方文件為準**

## 授權

尚未設定 License。

## 問題回報

請至 [Issues](https://github.com/windskyshao/tw-transcript-structurer/issues) 開 issue。
