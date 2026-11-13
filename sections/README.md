---
docs: true
title: 進め方
sidebar:
  label: 進め方
  order: 0
---

# 進め方

## 必要なもの

| 章 | 必要なもの |
|---|---|
| 01. 詳細設計書を作る | Windows の Excel（[Excel の準備](./00-setup/01-excel/LECTURE.md)） |
| 02. TypeScript の型を作る | Node.js とエディタ（[Node.js の準備](./00-setup/02-node/LECTURE.md)） |

第 1 章だけ、第 2 章だけでも進められます。第 2 章は第 1 章の最後の YAML を使いますが、必要なファイルは節ごとに配っています。

## 節の進め方

各節の冒頭に、その節で使うファイルのダウンロードボタンがあります。ZIP を展開し、本文に沿って手を動かします。

- 第 1 章は、最初の節で作った `詳細設計書.xlsm` を最後まで育てていきます。途中の節の ZIP には、その節を終えた状態の `openapi.yaml` が入っています
- 第 2 章は、節ごとに独立したプロジェクトです

## 題材

JWT で認証する「タスク管理 API」です。ログイン、タスクの一覧・作成・更新・削除、添付ファイル、履歴、カテゴリなど 17 本の API があります。OpenAPI のいろいろな書き方（`$ref`、`allOf`、`oneOf` と `discriminator`、`multipart/form-data`、再帰する型、`additionalProperties` など）がひととおり入るように作ってあります。
