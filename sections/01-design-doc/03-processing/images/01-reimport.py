import sys, os
sys.path.insert(0, os.path.expanduser("~/.claude/skills/genfig"))
from genfig import Canvas


c = Canvas(760, 360)
tpl = c.node(130, 90, "ひな形", shape="sticky", color="purple", w=170, h=60, label_scale="label")
old = c.node(130, 270, "古いシート", shape="sticky", color="gray", w=170, h=60, label_scale="label")
new = c.node(560, 180, "新しいシート", shape="sticky", color="green", w=200, h=80, label_scale="label")
c.link(tpl, new, label="作り直す", label_pos="above")
c.link(old, new, label="処理の内容を移す", primary=False, label_pos="below")
c.text(130, 340, "移したあとで削除", scale="caption")
c.save("01-reimport.svg")
