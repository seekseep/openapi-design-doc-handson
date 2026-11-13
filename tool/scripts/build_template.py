"""Build the template workbook (詳細設計書テンプレート.xlsm).

    python3 tool/scripts/build_template.py

1. Make the sheets with openpyxl (tool/build/template.xlsx, no macros)
2. Put the VBA in tool/vba into it (inject_vba.py)
   -> tool/dist/詳細設計書テンプレート.xlsm
   (also copied to the lectures that hand it out)
3. A workbook for the parser tests -> tool/test/parser-test.xlsm

Excel is not needed. The layout of the design document is defined here:
column A of the template sheets holds the markers read by modTemplate.bas.
"""
from __future__ import annotations

import shutil
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.hyperlink import Hyperlink

sys.path.insert(0, str(Path(__file__).resolve().parent))
from inject_vba import inject  # noqa: E402

TOOL = Path(__file__).resolve().parents[1]
BUILD = TOOL / "build"
DIST = TOOL / "dist"
ROOT = TOOL.parent
# Lectures that hand out the template workbook.
LECTURE_COPIES = [ROOT / "sections" / "01-design-doc" / "01-import" / "example"]
OUTPUT_NAME = "詳細設計書テンプレート.xlsm"
TEMPLATE_VERSION = "1"

FONT = "Meiryo UI"
THIN = Side(style="thin", color="A6A6A6")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEADER_FILL = PatternFill("solid", fgColor="DDEBF7")
SECTION_FILL = PatternFill("solid", fgColor="1F4E78")
SUB_FILL = PatternFill("solid", fgColor="F2F2F2")
NOTE_COLOR = "7F7F7F"
WRAP_TOP = Alignment(wrap_text=True, vertical="top")

# Columns of the endpoint sheet (A is the hidden marker column).
ENDPOINT_WIDTHS = {"A": 14, "B": 5, "C": 26, "D": 22, "E": 6, "F": 26, "G": 48, "H": 22, "I": 9}
TABLE_COLS = "CDEFGHI"


def font(**kw) -> Font:
    return Font(name=FONT, size=kw.pop("size", 10), **kw)


class SheetWriter:
    """Append rows to a template sheet. Each row may have a marker in A."""

    def __init__(self, ws, last_col: str):
        self.ws = ws
        self.row = 0
        self.last_col = last_col

    def add(self, mark: str = "", cells: dict[str, str] | None = None, *, hidden: bool = False,
            style: str = "", merge: str = "") -> int:
        self.row += 1
        r = self.row
        ws = self.ws
        if mark:
            ws[f"A{r}"] = mark
            ws[f"A{r}"].font = font(color=NOTE_COLOR, size=8)
        for col, value in (cells or {}).items():
            ws[f"{col}{r}"] = value
        if merge:
            first, last = merge.split(":")
            ws.merge_cells(f"{first}{r}:{last}{r}")
        self._style(r, style, cells or {}, merge)
        if hidden:
            ws.row_dimensions[r].hidden = True
        return r

    def _style(self, r: int, style: str, cells: dict, merge: str) -> None:
        ws = self.ws
        cols = [c for c in "BCDEFGHI" if c <= self.last_col]
        for c in cols:
            ws[f"{c}{r}"].font = font()
        if style == "title":
            ws[f"B{r}"].font = font(size=14, bold=True)
            ws.row_dimensions[r].height = 24
        elif style == "section":
            for c in cols:
                ws[f"{c}{r}"].fill = SECTION_FILL
                ws[f"{c}{r}"].font = font(size=11, bold=True, color="FFFFFF")
            ws.row_dimensions[r].height = 20
        elif style == "sub":
            ws[f"C{r}"].font = font(bold=True)
        elif style == "note":
            for c in cols:
                ws[f"{c}{r}"].font = font(size=9, color=NOTE_COLOR)
        elif style in ("header", "cells"):
            used = self._table_cols(cells, merge)
            for c in used:
                cell = ws[f"{c}{r}"]
                cell.border = BORDER
                cell.alignment = WRAP_TOP
                if style == "header":
                    cell.fill = HEADER_FILL
                    cell.font = font(bold=True)
        elif style == "label":
            # label in C, value in D (merged to the right)
            ws[f"C{r}"].fill = SUB_FILL
            for c in self._table_cols(cells, merge):
                ws[f"{c}{r}"].border = BORDER
                ws[f"{c}{r}"].alignment = WRAP_TOP
        elif style == "text":
            for c in self._table_cols(cells, merge):
                ws[f"{c}{r}"].alignment = WRAP_TOP

    def _table_cols(self, cells: dict, merge: str) -> list[str]:
        cols = set(cells)
        if merge:
            first, last = merge.split(":")
            cols |= {c for c in "BCDEFGHI" if first <= c <= last}
        return sorted(cols)


def field_table(w: SheetWriter, list_name: str, empty_text: str, with_in: bool) -> None:
    head = {"C": "項目", "D": "型", "E": "必須", "F": "制約", "G": "説明", "H": "例"}
    row = {"C": "{{name}}", "D": "{{type}}", "E": "{{required}}", "F": "{{constraints}}",
           "G": "{{description}}", "H": "{{example}}"}
    if with_in:
        head = {**head, "C": "名前", "I": "位置"}
        row = {**row, "I": "{{in}}"}
    w.add(cells=head, style="header")
    w.add(f"@row:{list_name}", row, hidden=True, style="cells")
    w.add(f"@empty:{list_name}", {"C": empty_text}, hidden=True, style="note")


def build_endpoint(ws) -> None:
    ws.title = "_tpl_endpoint"
    for col, width in ENDPOINT_WIDTHS.items():
        ws.column_dimensions[col].width = width
    ws.column_dimensions["A"].hidden = True
    ws.sheet_view.showGridLines = False
    w = SheetWriter(ws, "I")

    w.add("opid:{{operationId}}", {"B": "{{operationId}}　{{summary}}"}, style="title")
    r = w.add(cells={"B": "← 目次へ戻る"}, style="note")
    ws[f"B{r}"].hyperlink = Hyperlink(ref=f"B{r}", location="'目次'!A1", display="← 目次へ戻る")
    w.add()

    w.add(cells={"B": "■ 基本情報"}, style="section")
    w.add(cells={"C": "operationId", "D": "{{operationId}}"}, style="label", merge="D:I")
    w.add(cells={"C": "メソッド", "D": "{{method}}"}, style="label", merge="D:I")
    w.add(cells={"C": "パス", "D": "{{path}}"}, style="label", merge="D:I")
    w.add(cells={"C": "タグ", "D": "{{tags}}"}, style="label", merge="D:I")
    w.add(cells={"C": "概要", "D": "{{summary}}"}, style="label", merge="D:I")
    w.add("@if:deprecated", {"C": "非推奨", "D": "{{deprecated}}（新しく使わないこと）"}, hidden=True,
          style="label", merge="D:I")
    w.add()

    w.add(cells={"B": "■ 目的"}, style="section")
    w.add(cells={"C": "{{purpose}}"}, style="text", merge="C:I")
    w.add()

    w.add(cells={"B": "■ 権限"}, style="section")
    w.add(cells={"C": "認証", "D": "{{auth}}"}, style="label", merge="D:I")
    w.add(cells={"C": "ロール", "D": "条件"}, style="header", merge="D:I")
    w.add("@row:permissions", {"C": "{{role}}", "D": "{{condition}}"}, hidden=True, style="cells", merge="D:I")
    w.add("@empty:permissions", {"C": "（x-permissions の定義なし）"}, hidden=True, style="note")
    w.add()

    w.add(cells={"B": "■ リクエスト"}, style="section")
    w.add(cells={"C": "▼ パラメータ"}, style="sub")
    field_table(w, "params", "なし", with_in=True)
    w.add()
    w.add("@block:bodies", hidden=True)
    w.add(cells={"C": "▼ ボディ（{{contentType}}）"}, style="sub")
    field_table(w, "fields", "（項目なし）", with_in=False)
    w.add()
    w.add("@end:bodies", hidden=True)
    w.add("@empty:bodies", {"C": "▼ ボディ：なし"}, hidden=True, style="sub")
    w.add()

    w.add(cells={"B": "■ レスポンス"}, style="section")
    w.add("@block:responses", hidden=True)
    w.add(cells={"C": "▼ {{status}}　{{statusDescription}}"}, style="sub")
    w.add("@section:headers", hidden=True)
    w.add(cells={"C": "ヘッダー"}, style="note")
    w.add(cells={"C": "名前", "D": "型", "E": "必須", "F": "制約", "G": "説明"}, style="header")
    w.add("@row:headers", {"C": "{{name}}", "D": "{{type}}", "E": "{{required}}", "F": "{{constraints}}",
                           "G": "{{description}}"}, hidden=True, style="cells")
    w.add("@end:headers", hidden=True)
    w.add("@block:contents", hidden=True)
    w.add(cells={"C": "ボディ（{{contentType}}）"}, style="note")
    field_table(w, "fields", "（項目なし）", with_in=False)
    w.add("@end:contents", hidden=True)
    w.add("@empty:contents", {"C": "（ボディなし）"}, hidden=True, style="note")
    w.add()
    w.add("@end:responses", hidden=True)

    w.add(cells={"B": "■ 処理の内容"}, style="section")
    w.add(cells={"C": "※ この表は手で書く欄です。OpenAPI を取り込み直しても消えません。行は自由に追加・削除できます。"},
          style="note")
    w.add(cells={"B": "No", "C": "処理", "D": "詳細", "H": "備考"}, style="header", merge="D:G")
    ws[f"B{w.row}"].border = BORDER
    ws[f"B{w.row}"].fill = HEADER_FILL
    w.add("@processing:begin", hidden=True)
    r = w.add("@row:processing", {"B": "{{no}}", "C": "{{step}}", "D": "{{detail}}", "H": "{{note}}"},
              hidden=True, style="cells", merge="D:G")
    ws[f"B{r}"].border = BORDER
    ws[f"B{r}"].alignment = WRAP_TOP
    for i in range(5):
        # Empty rows to write in. No values here: an empty area means
        # "nothing written yet" when the sheet is imported again.
        r = w.add("@empty:processing", {"C": "", "H": ""}, hidden=True, style="cells", merge="D:G")
        ws[f"B{r}"].border = BORDER
    w.add("@processing:end", hidden=True)


def build_index(ws) -> None:
    ws.title = "_tpl_index"
    widths = {"A": 10, "B": 5, "C": 26, "D": 10, "E": 40, "F": 36, "G": 10}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    ws.column_dimensions["A"].hidden = True
    ws.sheet_view.showGridLines = False
    w = SheetWriter(ws, "G")
    w.add("@index", {"B": "{{title}}　詳細設計書"}, style="title")
    w.add(cells={"C": "バージョン", "D": "{{version}}"})
    w.add(cells={"C": "取り込み元", "D": "{{sourcePath}}"})
    w.add(cells={"C": "取り込み日時", "D": "{{importedAt}}"})
    w.add()
    head = {"B": "No", "C": "operationId", "D": "メソッド", "E": "パス", "F": "概要", "G": "状態"}
    w.add(cells=head)
    for c in head:
        ws[f"{c}{w.row}"].fill = HEADER_FILL
        ws[f"{c}{w.row}"].border = BORDER
        ws[f"{c}{w.row}"].font = font(bold=True)
    w.add("@row:indexRows", {"B": "{{no}}", "C": "{{sheetLink}}", "D": "{{method}}", "E": "{{path}}",
                             "F": "{{summary}}", "G": "{{status}}"}, hidden=True)
    for c in head:
        ws[f"{c}{w.row}"].border = BORDER
        ws[f"{c}{w.row}"].alignment = WRAP_TOP
    ws[f"C{w.row}"].font = font(color="0563C1", underline="single")


def build_guide(ws) -> None:
    ws.title = "使い方"
    ws.sheet_properties.codeName = "shGuide"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 90
    ws["B2"] = "OpenAPI から詳細設計書を作る"
    ws["B2"].font = font(size=16, bold=True)
    ws["B4"] = "▶ OpenAPI（YAML）を取り込む"
    ws["B4"].font = font(size=12, bold=True, color="FFFFFF")
    ws["B4"].fill = PatternFill("solid", fgColor="2F75B5")
    ws["B4"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[4].height = 32
    ws["B4"].hyperlink = Hyperlink(ref="B4", location="'使い方'!B4", display="▶ OpenAPI（YAML）を取り込む")
    lines = [
        "",
        "使い方",
        "1. 上のボタンを押して openapi.yaml を選ぶと、API ごとのシートと「目次」ができます。",
        "2. 各シートの「■ 処理の内容」に、処理の流れを書き足します。",
        "3. YAML を直したら、もう一度ボタンを押します。シートは作り直されますが、処理の内容は残ります。",
        "",
        "注意",
        "・「処理の内容」以外の場所に書いたものは、取り込み直すと消えます。",
        "・YAML から消えた API のシートは削除されず、目次で「廃止」と表示されます。",
        "・取り込む前に、このブックのコピーを backup フォルダに保存します。",
        "・見た目は非表示のシート「_tpl_endpoint」で変えられます（シートの表示から再表示できます）。",
    ]
    for i, text in enumerate(lines, start=5):
        ws[f"B{i}"] = text
        ws[f"B{i}"].font = font(bold=text in ("使い方", "注意"))
    return None


def build_config(ws) -> None:
    ws.title = "_config"
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 60
    rows = [
        ("key", "value"),
        ("templateVersion", TEMPLATE_VERSION),
        ("backup", "on"),
        ("indexSheetName", "目次"),
        ("yes", "○"),
        ("root", "（ボディ全体）"),
        ("recursive", "（再帰：同じ型が入れ子になる）"),
        ("authNone", "認証なし"),
        ("deprecated", "非推奨"),
        ("removed", "廃止"),
        ("oneOf", "次のいずれか 1 つ"),
        ("anyOf", "次のうち 1 つ以上"),
        ("msgDone", "取り込みが完了しました。"),
        ("msgFailed", "取り込みに失敗しました。"),
        ("lastPath", ""),
        ("lastImport", ""),
    ]
    for r, (k, v) in enumerate(rows, start=1):
        ws.cell(r, 1, k).font = font(bold=(r == 1))
        ws.cell(r, 2, v).font = font(bold=(r == 1))


def main() -> int:
    BUILD.mkdir(exist_ok=True)
    DIST.mkdir(exist_ok=True)
    wb = Workbook()
    build_guide(wb.active)
    endpoint = wb.create_sheet()
    build_endpoint(endpoint)
    index = wb.create_sheet()
    build_index(index)
    config = wb.create_sheet()
    build_config(config)
    for ws in (endpoint, index, config):
        ws.sheet_state = "hidden"
    wb.defined_names["ImportButton"] = DefinedName("ImportButton", attr_text="'使い方'!$B$4")
    xlsx = BUILD / "template.xlsx"
    wb.save(xlsx)
    inject(xlsx, TOOL / "vba", DIST / OUTPUT_NAME, exclude={"modTest"})
    for dest in LECTURE_COPIES:
        shutil.copyfile(DIST / OUTPUT_NAME, dest / OUTPUT_NAME)
        print(f"copied to {dest.relative_to(ROOT)}")

    # A workbook for running the YAML parser tests by hand (macro RunAllTests).
    test_wb = Workbook()
    test_wb.active.title = "Test"
    test_wb.active["A1"] = "開発 → マクロ → RunAllTests を実行すると、tool/fixtures のテストを流します。"
    test_xlsx = BUILD / "parser-test.xlsx"
    test_wb.save(test_xlsx)
    inject(test_xlsx, TOOL / "vba", TOOL / "test" / "parser-test.xlsm")
    return 0


if __name__ == "__main__":
    sys.exit(main())
