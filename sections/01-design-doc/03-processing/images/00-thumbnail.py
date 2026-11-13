import sys, os
sys.path.insert(0, os.path.expanduser("~/.claude/skills/genfig"))
from genfig import Canvas


c = Canvas(760, 340)
yaml = c.node(110, 110, "openapi.yaml", emoji_cp="1f4c4", label_scale="label")
person = c.node(110, 250, "設計する人", emoji_cp="270d", label_scale="label")
c.sticky(330, 50, 390, 250, color="gray")
c.text(525, 86, "API のシート", scale="label")
auto = c.node(525, 140, "目的・権限・リクエスト・レスポンス", shape="sticky", color="blue", w=340, h=60, label_scale="label")
hand = c.node(525, 240, "処理の内容（取り込み直しても残る）", shape="sticky", color="yellow", w=340, h=60, label_scale="label")
c.link(yaml, auto, label="取り込む")
c.link(person, hand, label="書き足す")
c.save("00-thumbnail.svg")
