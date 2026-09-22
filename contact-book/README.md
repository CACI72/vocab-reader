# contact-book — 家庭聯絡簿「教師叮嚀」欄產生器

固化自 2026-09-22 兩次聯絡簿任務的檢討。那兩次合計耗時 2 小時 42 分，
其中一次產出的 .docx 無法開啟。本目錄把當時**暫存於容器、隨容器回收而滅失**的
產生腳本與規範文件固定下來，使同一任務不必再重建。

## 內容

| 檔案 | 用途 |
|---|---|
| `TASK-CARD.md` | **單一入口**任務卡：語氣規範、範例、版面參數、輸出路徑、品質檢核 |
| `build_contact_book.py` | .docx 列印表格產生器（2 欄 × 4 列，每格 8.5cm × 5cm，標楷體 14pt） |
| `entries.example.json` | 輸入格式範例（S 代號，已去識別化） |
| `test_build.py` | 回歸測試，含「無法開啟」bug 的防護驗證 |

## 使用

```bash
pip3 install python-docx
python3 build_contact_book.py entries.json -o 家庭聯絡簿_教師叮嚀欄_20260922.docx
python3 test_build.py      # 回歸測試
```

## 為什麼需要這支腳本

先前兩種做法都曾產出**無法開啟**的檔案，且兩者的檔案大小與 ZIP 結構看起來都正常：

1. **由模型在對話上下文中逐字搬運 base64** — Read 讀取無換行長檔會靜默截斷；
   即使分段重組，單字元誤植即造成 ZIP CRC 損毀，而檔案大小仍吻合。
2. **手工拼接 OOXML，把元素 append 到 `w:tblPr` 尾端** — `w:tblPr` 是有序序列
   （ECMA-376 CT_TblPrBase）且多數子元素不可重複；違反時 Word 判定「內容無法讀取」。
   本專案初版即犯此錯（重複的 `w:tblLayout` ＋ `w:tblCellMar` 排在 `w:tblLook` 之後）。

因此腳本有兩條硬性設計：

- **檔案位元組永不經過模型上下文**：直接寫入磁碟，再以檔案傳遞工具交付。
- **產出後自我驗證**：ZIP 完整性 → `w:tblPr` schema 順序與重複 → 欄寬列高；
  任一項失敗即以非零狀態離開，不會把壞檔交出去。

`test_build.py::test_schema_checker_catches_appended_element` 會刻意重現
第 2 種 bug，確認檢查器抓得到——該案例的 ZIP 完整性測試仍然通過，
正說明只驗 ZIP 是不夠的。

## 已知限制

- 本容器的 LibreOffice 無法載入任何 .docx（連純 python-docx 基準檔亦然），
  屬沙箱環境問題，因此**無法在此進行真實 Word／LibreOffice 開檔測試**。
  把關依靠 ZIP 完整性、OOXML schema 檢查與 python-docx 往返解析。
  首次於實機列印前，建議仍以 Word 開啟確認一次。
- 列高設為 exact 5cm，超過 96 字的內容會被截斷；腳本會事先警告，但不會自動縮字。
- 逐格獨立性檢查採關鍵詞比對，能攔下「同樣」「其他同學」「比…更」等明確跨生指涉，
  但無法理解語意；語意層面的跨格依賴仍需教師過目。
