#!/usr/bin/env python3
"""amend.json から新旧対照表の原稿（content/01b_shinkyu.txt）を作る。"""
import json, sys
a = json.load(open(sys.argv[1], encoding="utf8"))
L = ["@header 学校法人明珠学園　就業規則 新旧対照表", "# 就業規則　新旧対照表",
     "@center （こども性暴力防止法の施行に伴う改正案）", "",
     "> 改正のある箇所のみを掲げています。「（新設）」は現行規則に対応する規定がないことを示します。施行日は理事会の決定に合わせて記入してください。", "",
     "@cols 12,27,41,20", "| 条 | 現行 | 改正案 | 改正理由 |", "|---|---|---|---|"]
for r in a:
    f = lambda xs: "<br>".join(x.replace("|", "｜") for x in xs)
    L.append(f"| {r['jo']} | {f(r['old'])} | {f(r['new'])} | {r['note']} |")
open(sys.argv[2], "w", encoding="utf8").write("\n".join(L) + "\n")
