import sys, os
sys.path.insert(0, os.path.expanduser("~/.claude/skills/genfig"))
from genfig import Canvas


c = Canvas(760, 360)
c.text(380, 40, "目次", scale="heading")
rows = [("listTasks", "GET /tasks", "", "gray"), ("addComment", "POST /tasks/{taskId}/comments", "（増えた）", "green"),
        ("receiveWebhook", "POST /webhooks/task-events", "非推奨", "gray"), ("getHealth", "（YAML から消えた）", "廃止", "red")]
for i, (op, path, st, col) in enumerate(rows):
    y = 100 + i * 64
    c.node(150, y, op, shape="sticky", color=col if col != "gray" else "yellow", w=200, h=50, label_scale="label")
    c.text(270, y + 6, path, scale="body", align="left")
    if st:
        c.text(690, y + 6, st, scale="label", fill="#b91c1c" if col == "red" else None)
c.save("01-index.svg")
