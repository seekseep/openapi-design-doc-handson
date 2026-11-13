import sys, os
sys.path.insert(0, os.path.expanduser("~/.claude/skills/genfig"))
from genfig import Canvas


c = Canvas(760, 330)
yaml = c.node(130, 165, "openapi.yaml", emoji_cp="1f4c4", label_scale="label")
doc = c.node(610, 80, "詳細設計書", emoji_cp="1f4ca", label_scale="label", w=110, h=80)
ts = c.node(610, 250, "TypeScript の型", emoji_cp="1f537", label_scale="label", w=110, h=80)
c.link(yaml, doc, label="Excel マクロ")
c.link(yaml, ts, label="openapi-typescript")
c.save("00-thumbnail.svg")
