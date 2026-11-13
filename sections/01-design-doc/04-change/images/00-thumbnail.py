import sys, os
sys.path.insert(0, os.path.expanduser("~/.claude/skills/genfig"))
from genfig import Canvas


c = Canvas(760, 380)
yaml = c.node(120, 190, "変更した YAML", emoji_cp="1f4c4", label_scale="label")
add = c.node(520, 80, "addComment：シートが増える", shape="sticky", color="green", w=330, h=60, label_scale="label")
chg = c.node(520, 190, "listTasks：表の行が増える", shape="sticky", color="blue", w=330, h=60, label_scale="label")
rm = c.node(520, 300, "getHealth：残して「廃止」", shape="sticky", color="gray", w=330, h=60, label_scale="label")
for n in (add, chg, rm):
    c.link(yaml, n, end="w")
c.save("00-thumbnail.svg")
