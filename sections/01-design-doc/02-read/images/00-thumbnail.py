import sys, os
sys.path.insert(0, os.path.expanduser("~/.claude/skills/genfig"))
from genfig import Canvas


c = Canvas(760, 400)
c.text(170, 40, "openapi.yaml", scale="heading")
c.text(590, 40, "設計書のシート", scale="heading")
rows = [("x-purpose", "■ 目的"), ("security / x-permissions", "■ 権限"), ("parameters / requestBody", "■ リクエスト"), ("responses", "■ レスポンス")]
for i, (l, r) in enumerate(rows):
    y = 100 + i * 80
    a = c.node(170, y, l, shape="sticky", color="yellow", w=280, h=56, label_scale="label")
    b = c.node(590, y, r, shape="sticky", color="green", w=200, h=56, label_scale="label")
    c.link(a, b)
c.save("00-thumbnail.svg")
