# -*- coding: utf-8 -*-
import re
from pathlib import Path

root = Path(r"D:\zwr\idea\project\111\innerworld\精神治疗\抑郁症")
files = sorted(root.glob("模块*/*.md")) + sorted(root.glob("附录：术语与参考文献/*.md"))
files = [p for p in files if p.is_file()]

stats = {k: 0 for k in ["h1","h2","h3","h4","table","fence","quote","check","link","bold","code","br","escaped_pipe"]}
max_cols = 0
weird = []
for p in files:
    lines = p.read_text(encoding="utf-8").splitlines()
    in_fence = False
    for i, line in enumerate(lines, 1):
        if line.startswith("```"):
            stats["fence"] += 1
            in_fence = not in_fence
            continue
        if in_fence:
            continue
        if line.startswith("# "): stats["h1"] += 1
        elif line.startswith("## "): stats["h2"] += 1
        elif line.startswith("### "): stats["h3"] += 1
        elif line.startswith("#### "): stats["h4"] += 1
        if line.startswith("|") and line.count("|") >= 2:
            stats["table"] += 1
            max_cols = max(max_cols, line.count("|") - 1)
        if line.startswith(">"): stats["quote"] += 1
        if re.match(r"^\s*- \[[ xX]\]", line): stats["check"] += 1
        stats["link"] += len(re.findall(r"\[[^\]]+\]\([^)]+\)", line))
        stats["bold"] += len(re.findall(r"\*\*[^*]+\*\*", line))
        stats["code"] += len(re.findall(r"`[^`]+`", line))
        stats["br"] += line.count("<br>")
        stats["escaped_pipe"] += line.count(r"\|")
        if "|" in line and not line.startswith("|") and line.count("|") >= 2 and not line.startswith("```"):
            weird.append((str(p.relative_to(root)), i, line[:120]))
print("files", len(files))
print(stats)
print("max_cols", max_cols)
print("weird pipe lines", len(weird))
for w in weird[:20]: print(w)
print("total lines", sum(len(p.read_text(encoding='utf-8').splitlines()) for p in files))
