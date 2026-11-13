# API を変えて取り込み直す

本文は [LECTURE.md](./LECTURE.md)。

`example/openapi.yaml` は 01-import の YAML に次の変更を入れたものです。

- `listTasks` に `priority` パラメータを追加
- `addComment`（POST /tasks/{taskId}/comments）を追加
- `getHealth`（GET /health）を削除
