# 型を使って API を呼ぶ

本文は [LECTURE.md](./LECTURE.md)。

```powershell
cd example
npm install
npm run gen:ts      # openapi.yaml -> src/generated/api.d.ts
npm run typecheck   # src/client.ts が YAML どおりか確かめる
```
