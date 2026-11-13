import sys, os
sys.path.insert(0, os.path.expanduser("~/.claude/skills/genfig"))
from genfig import Canvas


c = Canvas(760, 330)
yaml = c.node(140, 110, "openapi.yaml", emoji_cp="1f4c4", label_scale="label")
tpl = c.node(140, 250, "テンプレート", emoji_cp="1f5c2", label_scale="label")
macro = c.node(320, 180, "取り込む", shape="hexagon", color="blue", w=150, h=90, label_scale="label")
doc = c.node(650, 180, "詳細設計書", emoji_cp="1f4ca", label_scale="label")
c.link(yaml, macro)
c.link(tpl, macro)
c.link(macro, doc, label="API ごとのシート")
c.save("00-thumbnail.svg")
