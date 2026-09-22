#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_contact_book.py — 家庭聯絡簿「教師叮嚀」欄列印表格產生器

用途
----
把每日各生的「教師叮嚀」欄文字，排成可直接列印、裁切後黏貼於紙本
「家庭聯絡暨學習紀錄」的表格：2 欄 × 4 列（共 8 格），每格精確
9cm × 5cm，字體標楷體 14pt。

設計原則（回應 2026-09-22 兩次 session 的失效檢討）
------------------------------------------------
1. 檔案位元組**永遠不經過模型的上下文**。本腳本直接把 .docx 寫到磁碟，
   再由檔案傳遞工具交付。歷史上「模型逐字搬運 base64」曾造成 ZIP CRC
   損毀、檔案大小卻吻合的靜默失敗（Word 無法開啟）。
2. 版面參數集中在 LAYOUT，不再每次重新推導。
3. 產出後自我驗證（ZIP 結構 + OOXML schema 順序 + 欄寬列高），驗證失敗即非零離開。

用法
----
    python3 build_contact_book.py entries.json -o 家庭聯絡簿_教師叮嚀欄_20260922.docx

entries.json 格式（陣列，最多 8 筆；不足 8 筆時其餘格留白）：
    [
      {"label": "S01", "text": "今日參與語文分組課程。..."},
      ...
    ]

`label` 為顯示於該格開頭的稱謂。依 TASK-CARD.md 之語氣規範，正式產出
應為學生本名之「名」，輸出即成為「〇〇今日表現：…」。
範例檔一律使用 S 代號，以符合去識別化規範。
"""

import argparse
import json
import sys
import zipfile
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.table import WD_ROW_HEIGHT_RULE
from docx.oxml.ns import qn
from docx.shared import Cm, Pt


# ── 版面參數（唯一權威；要改版面只改這裡）────────────────────────
LAYOUT = {
    "cols": 2,
    "rows": 4,
    "cell_w_cm": 9.0,        # 對應紙本教師叮嚀欄寬
    "cell_h_cm": 5.0,        # 對應紙本教師叮嚀欄高
    "font": "標楷體",
    "font_size_pt": 14,
    "line_spacing_pt": 20,   # 14pt 字在 5cm 內可容 6 行
    "cell_margin_cm": 0.15,
    "page_margin_cm": 1.5,
    "suffix": "今日表現：",  # 接在 label 之後的固定開頭
}

CELLS = LAYOUT["cols"] * LAYOUT["rows"]

# 5cm 列高扣掉上下內距後的可用高度，除以行高 → 可容行數
_USABLE_H_PT = (LAYOUT["cell_h_cm"] - 2 * LAYOUT["cell_margin_cm"]) * 28.3465
MAX_LINES = int(_USABLE_H_PT // LAYOUT["line_spacing_pt"])
# 9cm 扣掉左右內距，除以全形字寬（= 字級）→ 每行可容全形字數
CHARS_PER_LINE = int(
    ((LAYOUT["cell_w_cm"] - 2 * LAYOUT["cell_margin_cm"]) * 28.3465)
    // LAYOUT["font_size_pt"]
)
MAX_CHARS = MAX_LINES * CHARS_PER_LINE


def _set_cjk_font(run):
    """python-docx 的 font.name 只設 ascii/hAnsi，中日韓字型要另外設 eastAsia。"""
    run.font.name = LAYOUT["font"]
    run.font.size = Pt(LAYOUT["font_size_pt"])
    run._element.rPr.rFonts.set(qn("w:eastAsia"), LAYOUT["font"])


def _set_cell_margins(table, cm):
    """設定表格內距。

    w:tblPr 是 **有序序列**（CT_TblPrBase），tblCellMar 必須排在 tblLook
    之前。直接 append 會讓 Word 判定內容無法讀取——這正是本專案先前
    產出無法開啟的同一類成因，故此處明確插在 tblLook 之前。
    """
    tbl_pr = table._tbl.tblPr
    mar = tbl_pr.makeelement(qn("w:tblCellMar"), {})
    twips = str(int(cm * 567))
    for side in ("top", "left", "bottom", "right"):
        node = mar.makeelement(qn(f"w:{side}"), {qn("w:w"): twips, qn("w:type"): "dxa"})
        mar.append(node)

    look = tbl_pr.find(qn("w:tblLook"))
    if look is not None:
        look.addprevious(mar)
    else:
        tbl_pr.append(mar)


def build(entries, out_path):
    if len(entries) > CELLS:
        raise ValueError(f"最多 {CELLS} 筆，收到 {len(entries)} 筆")

    doc = Document()

    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.PORTRAIT
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    m = Cm(LAYOUT["page_margin_cm"])
    sec.left_margin = sec.right_margin = sec.top_margin = sec.bottom_margin = m

    # Normal 樣式也套標楷體，避免空格繼承 Calibri
    normal = doc.styles["Normal"]
    normal.font.name = LAYOUT["font"]
    normal.font.size = Pt(LAYOUT["font_size_pt"])
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), LAYOUT["font"])

    table = doc.add_table(rows=LAYOUT["rows"], cols=LAYOUT["cols"])
    table.style = "Table Grid"          # 實線框 = 裁切導引線
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False       # 由 python-docx 依序插入 tblLayout=fixed
    _set_cell_margins(table, LAYOUT["cell_margin_cm"])

    for col in table.columns:
        col.width = Cm(LAYOUT["cell_w_cm"])
    for row in table.rows:
        row.height = Cm(LAYOUT["cell_h_cm"])
        row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY

    warnings = []
    for idx in range(CELLS):
        r, c = divmod(idx, LAYOUT["cols"])
        cell = table.cell(r, c)
        cell.width = Cm(LAYOUT["cell_w_cm"])
        cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP

        para = cell.paragraphs[0]
        pf = para.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing = Pt(LAYOUT["line_spacing_pt"])

        if idx >= len(entries):
            _set_cjk_font(para.add_run(""))
            continue

        entry = entries[idx]
        body = str(entry.get("text", "")).strip()
        label = str(entry.get("label", "")).strip()
        text = f"{label}{LAYOUT['suffix']}{body}" if label else body

        for i, line in enumerate(text.split("\n")):
            target = para if i == 0 else cell.add_paragraph()
            if target is not para:
                tpf = target.paragraph_format
                tpf.space_before = tpf.space_after = Pt(0)
                tpf.line_spacing = Pt(LAYOUT["line_spacing_pt"])
            _set_cjk_font(target.add_run(line))

        if len(text) > MAX_CHARS:
            warnings.append(
                f"  第 {idx + 1} 格（{label or '未命名'}）{len(text)} 字，"
                f"超出 {MAX_CHARS} 字上限，列高固定 5cm 會截斷"
            )

    doc.save(out_path)
    return warnings


# w:tblPr 的合法子元素順序（ECMA-376 CT_TblPrBase 序列）
TBLPR_ORDER = [
    "tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize",
    "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd",
    "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook",
    "tblCaption", "tblDescription",
]


def _check_tblpr_schema(path):
    """檢查 w:tblPr 子元素的順序與重複。

    w:tblPr 是有序序列且多數子元素只能出現一次。把元素 append 到尾端
    看似無害，Word 卻會直接判定「內容無法讀取」而拒絕開啟；LibreOffice
    與 Google Drive 的容忍度不同，因此不能只靠「某個閱讀器打得開」。
    """
    import xml.etree.ElementTree as ET

    ns = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    problems = []
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read("word/document.xml"))

    for tbl_idx, tbl_pr in enumerate(root.iter(f"{ns}tblPr"), start=1):
        names = [c.tag.replace(ns, "") for c in tbl_pr]
        known = [n for n in names if n in TBLPR_ORDER]

        dupes = {n for n in known if known.count(n) > 1}
        if dupes:
            problems.append(f"表格 {tbl_idx} 的 w:tblPr 有重複元素：{sorted(dupes)}")

        ranks = [TBLPR_ORDER.index(n) for n in known]
        if ranks != sorted(ranks):
            problems.append(
                f"表格 {tbl_idx} 的 w:tblPr 子元素順序違反 schema：{known}"
            )

        unknown = [n for n in names if n not in TBLPR_ORDER]
        if unknown:
            problems.append(f"表格 {tbl_idx} 的 w:tblPr 含未知元素：{unknown}")

    return problems


def verify(path):
    """產出後自我驗證：ZIP 結構 + OOXML schema 順序 + 版面參數。"""
    problems = []

    try:
        with zipfile.ZipFile(path) as z:
            bad = z.testzip()
            if bad:
                problems.append(f"ZIP CRC 損毀於 {bad}")
            if "word/document.xml" not in z.namelist():
                problems.append("缺少 word/document.xml")
    except zipfile.BadZipFile as exc:
        return [f"不是有效的 ZIP/OOXML 檔：{exc}"]

    problems.extend(_check_tblpr_schema(path))

    doc = Document(path)
    if not doc.tables:
        return problems + ["文件中找不到表格"]

    t = doc.tables[0]
    if len(t.rows) != LAYOUT["rows"] or len(t.columns) != LAYOUT["cols"]:
        problems.append(f"表格為 {len(t.rows)}×{len(t.columns)}，"
                        f"應為 {LAYOUT['rows']}×{LAYOUT['cols']}")
    # OOXML 以 twip 儲存尺寸，換算回 EMU 會有次微米級誤差，故以 0.01cm 為容差
    tol = Cm(0.01)
    for i, row in enumerate(t.rows):
        if row.height is None or abs(row.height - Cm(LAYOUT["cell_h_cm"])) > tol:
            problems.append(f"第 {i + 1} 列列高非 {LAYOUT['cell_h_cm']}cm（實為 {row.height}）")
    for j, col in enumerate(t.columns):
        if col.width is None or abs(col.width - Cm(LAYOUT["cell_w_cm"])) > tol:
            problems.append(f"第 {j + 1} 欄欄寬非 {LAYOUT['cell_w_cm']}cm（實為 {col.width}）")

    return problems


def main():
    ap = argparse.ArgumentParser(description="家庭聯絡簿教師叮嚀欄列印表格產生器")
    ap.add_argument("entries", help="JSON 檔路徑（陣列，最多 8 筆）")
    ap.add_argument("-o", "--out", required=True, help="輸出 .docx 路徑")
    args = ap.parse_args()

    data = json.loads(Path(args.entries).read_text(encoding="utf-8"))
    if not isinstance(data, list):
        sys.exit("entries JSON 最外層必須是陣列")

    out = Path(args.out)
    warnings = build(data, out)

    problems = verify(out)
    if problems:
        print("✗ 驗證未通過：", file=sys.stderr)
        for p in problems:
            print(f"  {p}", file=sys.stderr)
        sys.exit(1)

    print(f"✓ 已產出 {out}（{out.stat().st_size:,} bytes）")
    print(f"  版面 {LAYOUT['cols']}欄×{LAYOUT['rows']}列，"
          f"每格 {LAYOUT['cell_w_cm']}cm×{LAYOUT['cell_h_cm']}cm，"
          f"{LAYOUT['font']} {LAYOUT['font_size_pt']}pt")
    print(f"  單格容量上限約 {MAX_CHARS} 字（{MAX_LINES} 行 × {CHARS_PER_LINE} 字）")
    print(f"  已填 {len(data)} 格，ZIP 結構與欄寬列高驗證通過")
    if warnings:
        print("\n⚠ 篇幅提醒：")
        for w in warnings:
            print(w)


if __name__ == "__main__":
    main()
