import sys, os
sys.path.insert(0, os.path.expanduser("~/.claude/skills/genfig"))
from genfig import Canvas


c = Canvas(760, 300)
c.sticky(30, 40, 700, 220, color="gray")
c.text(380, 80, "詳細設計書.xlsm", scale="heading")
names = [("使い方", "yellow"), ("目次", "blue"), ("login", "green"), ("listTasks", "green"), ("…17 枚", "green")]
for i, (n, col) in enumerate(names):
    c.node(110 + i * 135, 170, n, shape="sticky", color=col, w=118, h=70, label_scale="label")
c.save("01-sheets.svg")
