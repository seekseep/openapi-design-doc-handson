import sys, os
sys.path.insert(0, os.path.expanduser("~/.claude/skills/genfig"))
from genfig import Canvas


c = Canvas(760, 380)
yaml = c.node(110, 130, "openapi.yaml", emoji_cp="1f4c4", label_scale="label")
types = c.node(380, 130, "api.d.ts", emoji_cp="1f537", label_scale="label")
code = c.node(650, 130, "client.ts", emoji_cp="1f4bb", label_scale="label")
c.link(yaml, types, label="gen:ts")
c.link(types, code, label="型を使う")
warn = c.node(600, 310, "食い違いは実行前にエラー", shape="sticky", color="red", w=270, h=50, label_scale="label")
c.link(code, warn, primary=False, dash="dashed", label="typecheck", label_pos="above", start="s", end="n")
c.save("00-thumbnail.svg")
