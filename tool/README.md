# tool: 取り込みマクロとテンプレート

ハンズオンで配る `詳細設計書テンプレート.xlsm` の元です。受講者向けの説明は `sections/` にあり、ここは作る側の資料です。

## 構成

```
tool/
├── vba/                    # マクロのソース（正本）
│   ├── modMain.bas         # 入口（ImportOpenApi / ImportFromPath）、シートの作り直し、処理の内容の移植
│   ├── modOpenApi.bas      # OpenAPI → シートに書く値（$ref の解決、スキーマの平坦化）
│   ├── modTemplate.bas     # ひな形シートの目印（A 列）を読んで行を展開する
│   ├── modYaml.bas         # YAML パーサ
│   ├── modFile.bas         # ファイル選択、UTF-8 の読み込み
│   ├── modTest.bas         # パーサのテスト（配布物には入れない）
│   ├── clsNode.cls         # YAML のノード（マップ／配列／スカラー）
│   └── sheets/shGuide.cls  # 「使い方」シートのボタン（ハイパーリンク）のイベント
├── scripts/
│   ├── build_template.py   # テンプレートのブックを組み立てる（Excel 不要）
│   ├── inject_vba.py       # xlsx に VBA を入れて xlsm にする
│   ├── vbaproject.py       # vbaProject.bin を MS-OVBA / MS-CFB の仕様どおりに作る
│   ├── vba_lint.py         # 未宣言の変数などを簡易チェックする
│   ├── make_expected.py    # パーサのテストの期待値を PyYAML で作る
│   └── excel_dev.py        # （Mac 用）Excel を AppleScript で動かしてテストする
├── fixtures/               # パーサのテスト用 YAML と期待値
├── dist/                   # 配布するテンプレート（build_template.py の出力）
├── test/parser-test.xlsm   # パーサのテスト用ブック（build_template.py の出力）
├── DESIGN.md               # 最初の設計メモ
└── redocly.yaml            # サンプル YAML の lint 設定
```

## テンプレートを作り直す

Python 3 と openpyxl が要ります。Excel は要りません。

```sh
pip install openpyxl
python3 tool/scripts/vba_lint.py tool/vba/*.bas tool/vba/*.cls tool/vba/sheets/*.cls
python3 tool/scripts/build_template.py
```

`tool/dist/詳細設計書テンプレート.xlsm` ができ、テンプレートを配る節（`sections/01-design-doc/01-import/example/`）にもコピーされます。シートの見た目と目印は `build_template.py` で定義しています。

VBA は p-code（コンパイル済みのコード）を持たない形で入れています。Excel が開いたときにソースからコンパイルし直します。

## パーサのテスト

`tool/test/parser-test.xlsm` を開き、「開発」→「マクロ」→ `RunAllTests` を実行します。`tool/fixtures/` の YAML を読み、PyYAML で作った期待値（`*.expected`）と比べた結果を新しいシートに出します。

fixtures を足したら期待値を作り直します。

```sh
pip install pyyaml
python3 tool/scripts/make_expected.py
```

`fixtures/openapi.yaml` はハンズオンのサンプル（`sections/01-design-doc/01-import/example/openapi.yaml`）のコピーです。サンプルを直したらこちらも更新します。

## VBA を書くときの決まり

- **ASCII だけで書く**（コメントも英語）。VBE はソースを Shift_JIS などの ANSI コードページで読むため、UTF-8 の日本語は文字化けする。シートに出す日本語は、ひな形シートと `_config` シートに置く
- **バックスラッシュ（`\`）を書かない**。日本語の Mac Excel では 0x5C が円記号として読まれる。整数除算は `Int(a / b)`、文字としては `ChrW(92)` を使う
- モジュール変数の宣言は、最初のプロシージャより前に書く（`vba_lint.py` が検出する）
- `Scripting.Dictionary` / `FileSystemObject` / `ADODB.Stream` は使わない（Mac に無い。連想配列は `clsNode`）

`inject_vba.py` は ASCII 以外の文字とバックスラッシュを見つけると止まります。

## しくみ

- ひな形のシート `_tpl_endpoint` / `_tpl_index` の A 列（非表示）に、行の目印（`@row:` `@block:` など）を書いている。`modTemplate.bas` が目印を読み、YAML から作った値（`modOpenApi.bas`）で行を展開する。目印の一覧は `modTemplate.bas` の冒頭と、教材の「テンプレートの見た目を変える」にある
- シートと API の対応は、シートの A1 の `opid:<operationId>` で取る。シート名は見ない
- 取り込み直すときは、ひな形から新しいシートを作り、古いシートの `@processing:begin` と `@processing:end` の間の行を移してから、古いシートを消す
- 「使い方」シートのボタンは、自分自身へのハイパーリンクを持つセル。`Worksheet_FollowHyperlink` で取り込みを始める（フォームのボタンは openpyxl で作れないため）
