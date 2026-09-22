# contact-book

`daily-contact-book/` 是一個可直接安裝的 Claude skill：家庭聯絡簿「教師叮嚀」欄的
撰寫規範與列印表格產生器。

本目錄即該 skill 的**唯一權威來源**。Drive 上的 `PROJECTS/daily-contact-book/workflow.md`
已改為指標，不再保留內容副本——今天（2026-09-22）修掉的問題之一，
就是規範散落在多份文件而彼此分歧。

## 由來

固化自 2026-09-22 兩次聯絡簿任務的檢討。那兩次合計耗時 2 小時 42 分，
其中一次產出的 .docx 無法開啟，而產生腳本僅暫存於容器、隨容器回收滅失。

三項對應的設計：

1. **檔案位元組永不經過模型上下文** — 腳本直接寫入磁碟。歷史上「模型逐字搬運
   base64」造成 ZIP CRC 損毀，而檔案大小仍吻合，難以察覺。
2. **產出後自我驗證** — ZIP 完整性 → `w:tblPr` schema 順序與重複 → 欄寬列高，
   任一項失敗即非零離開。初版就犯過 schema 順序錯誤，而 ZIP 測試照樣通過。
3. **規範隨 skill 同步** — 不再依賴容器或單一雲端硬碟位置。

## 結構

```
daily-contact-book/
├── SKILL.md                          撰寫規範、輸入來源、執行步驟
├── scripts/build_contact_book.py     .docx 產生器（2 欄 × 4 列，每格 8.5cm × 5cm）
├── scripts/test_build.py             回歸測試 10 項
└── references/entries.example.json   輸入格式範例（S 代號）
```

## 開發

```bash
pip3 install python-docx
python3 daily-contact-book/scripts/test_build.py
```

## 打包安裝

```bash
python -m scripts.package_skill daily-contact-book   # 於 skill-creator 目錄執行
```

產出的 `.skill` 檔可在 Claude 介面直接安裝。

## 已知限制

- 逐格獨立性檢查採關鍵詞比對，攔得下「同樣」「其他同學」「比…更」等明確跨生指涉，
  但讀不懂語意；語意層面的跨格依賴仍需教師過目。
- 列高固定 5cm，超過 96 字會被截斷；腳本事先警告，但不自動縮字。
