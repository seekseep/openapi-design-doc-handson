import sys, os
sys.path.insert(0, os.path.expanduser("~/.claude/skills/genfig"))
from genfig import Canvas


c = Canvas(760, 340)
tpl = c.node(150, 170, "ひな形を直す", emoji_cp="1f3a8", label_scale="label")
names = ["login", "listTasks", "createTask"]
for i, n in enumerate(names):
    s = c.node(570, 70 + i * 100, n, shape="sticky", color="green", w=200, h=60, label_scale="label")
    c.link(tpl, s, end="w", label="取り込み直す" if i == 1 else None)
c.save("00-thumbnail.svg")
