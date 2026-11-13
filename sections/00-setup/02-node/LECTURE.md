---
docs: true
title: Node.js の準備
---

# Node.js の準備

第 2 章では、OpenAPI から TypeScript の型を作ります。そのために Node.js（npm）とエディタを入れます。第 1 章（Excel）だけ進める場合は不要です。

## Node.js を入れる

[Node.js の公式サイト](https://nodejs.org/ja/download) から **LTS** 版のインストーラー（Windows Installer `.msi`）を落として実行します。途中の選択肢は既定のままで構いません。

`winget` が使える場合は、PowerShell で次のコマンドでも入ります。

```powershell
winget install OpenJS.NodeJS.LTS
```

入ったら、**新しく開いた** PowerShell で確かめます。バージョンが表示されれば完了です。

```powershell
node -v
npm -v
```

## エディタを入れる

[Visual Studio Code](https://code.visualstudio.com/) を入れておくと、YAML と TypeScript を色付きで読めます。型の間違いもその場で赤線で教えてくれます。

フォルダを開くときは「ファイル」→「フォルダーを開く」で、ZIP を展開したフォルダを選びます。「ターミナル」→「新しいターミナル」で、そのフォルダを開いた PowerShell が下に出ます。

:::faq

### `node` が見つからないと言われる

インストール前から開いていたターミナルは、新しい設定を読み込んでいません。ターミナル（VS Code ごと）を開き直してください。

### `npm install` で「スクリプトの実行が無効」と言われる

PowerShell の実行ポリシーで `npm.ps1` が止められています。次のコマンドで、自分のユーザーだけ署名付きスクリプトを許可します。

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

:::
