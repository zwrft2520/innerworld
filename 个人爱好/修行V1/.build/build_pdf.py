# -*- coding: utf-8 -*-
"""把《修行V1 中国传统文化与实修知识体系》28 篇 Markdown 合并为一个带中文排版的 HTML，
供 Chrome 无头模式打印为 PDF。模块顺序严格按《目录.md》中出现的顺序。"""
import glob
import os
import re
import markdown

BASE = r"D:\zwr\idea\project\111\innerworld\个人爱好\修行V1"
OUT_HTML = os.path.join(BASE, ".build", "修行V1实修知识体系.html")

CSS = """
@page { size: A4; margin: 18mm 16mm 16mm 16mm; }
* { box-sizing: border-box; }
html { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
body {
  font-family: "Microsoft YaHei", "PingFang SC", "Source Han Serif SC", "SimSun", sans-serif;
  font-size: 11pt; line-height: 1.75; color: #2d2a26;
  margin: 0;
  -webkit-print-color-adjust: exact; print-color-adjust: exact;
}
.cover { text-align: center; padding: 88mm 0 0 0; page-break-after: always; }
.cover h1 { font-size: 30pt; color: #3c4a3a; margin: 0 0 12px 0; letter-spacing: 4px; }
.cover .sub { font-size: 13pt; color: #5f7059; letter-spacing: 2px; }
.cover .rule { width: 120px; height: 3px; background: #8da284; margin: 28px auto; }
.cover .meta { font-size: 10.5pt; color: #7d8a78; margin-top: 18px; }
.module-divider { page-break-before: always; border-top: 2px solid #cdd8c8; margin-top: 2px; padding-top: 14px; }
.module-divider h2 { color: #3c4a3a; font-size: 18pt; margin: 0 0 6px 0; }
.module-divider .goal { color: #6b7a66; font-size: 10.5pt; font-style: italic; margin: 0 0 10px 0; }
section.theme { page-break-before: always; }
h1 { font-size: 20pt; color: #33402f; border-bottom: 1px solid #cdd8c8; padding-bottom: 6px; }
h2 { font-size: 15pt; color: #3c4a3a; }
h3 { font-size: 12.5pt; color: #4c5c48; }
blockquote {
  margin: 10px 0; padding: 10px 16px;
  background: #eef2eb; border-left: 4px solid #8da284; color: #52584e;
  border-radius: 0 6px 6px 0;
}
blockquote p { margin: 4px 0; }
strong { color: #33402f; }
table { border-collapse: collapse; width: 100%; margin: 12px 0; font-size: 10pt; }
th, td { border: 1px solid #cdd8c8; padding: 6px 9px; vertical-align: top; text-align: left; }
th { background: #dde6d6; color: #33402f; }
tr:nth-child(even) td { background: #f4f7f1; }
code {
  font-family: Consolas, "Courier New", monospace;
  background: #eef2eb; color: #3c4a3a; padding: 1px 5px; border-radius: 3px; font-size: 9.5pt;
}
pre { background: #f1f5ee; border: 1px solid #d6e0cf; padding: 10px 14px; border-radius: 6px; overflow: hidden; }
pre code { background: none; padding: 0; font-size: 9.5pt; white-space: pre-wrap; }
ul, ol { margin: 6px 0 6px 0; padding-left: 22px; }
li { margin: 3px 0; }
hr { border: none; border-top: 1px solid #cdd8c8; margin: 18px 0; }
a { color: #5f7059; text-decoration: none; }
.references { margin-top: 10px; font-size: 9.5pt; color: #5c6459; }
.references h3 { font-size: 11pt; }
"""


def md_to_html(text):
    return markdown.markdown(
        text,
        extensions=["tables", "fenced_code", "sane_lists"],
    )


def parse_modules(toc_text):
    """按《目录.md》出现顺序解析模块名与其简介段落。"""
    modules = []
    lines = toc_text.splitlines()
    i = 0
    while i < len(lines):
        m = re.match(r"^##\s+(模块.+?)\s*$", lines[i])
        if m:
            name = m.group(1).strip()
            i += 1
            intro = []
            while i < len(lines):
                line = lines[i]
                if re.match(r"^#{1,6}\s", line):
                    break
                s = line.strip()
                if s:
                    intro.append(s)
                i += 1
            modules.append({"name": name, "intro": " ".join(intro).strip()})
            continue
        i += 1
    return modules


def main():
    with open(os.path.join(BASE, "目录.md"), "r", encoding="utf-8") as f:
        toc_text = f.read()
    toc_html = md_to_html(toc_text)

    parts = []
    parts.append(
        '<div class="cover">'
        '<h1>中国传统文化与实修知识体系</h1>'
        '<div class="rule"></div>'
        '<div class="sub">儒·释·道 三家实修总览</div>'
        '<div class="meta">总论 · 道家 · 佛家 · 儒家 · 通用实操 · 纠偏 · 典籍</div>'
        "</div>"
    )
    parts.append('<section class="toc">' + toc_html + "</section>")

    # 严格按《目录.md》顺序处理模块
    for mod in parse_modules(toc_text):
        mod_dir = os.path.join(BASE, mod["name"])
        if not os.path.isdir(mod_dir):
            continue
        divider = (
            '<div class="module-divider">'
            f"<h2>{mod['name']}</h2>"
            + (f'<div class="goal">{mod["intro"]}</div>' if mod["intro"] else "")
            + "</div>"
        )
        parts.append(divider)

        files = sorted(
            glob.glob(os.path.join(mod_dir, "*.md")),
            key=lambda p: os.path.basename(p),
        )
        for fp in files:
            with open(fp, "r", encoding="utf-8") as f:
                body = md_to_html(f.read())
            parts.append(f'<section class="theme">{body}</section>')

    html = (
        "<!DOCTYPE html><html lang='zh-CN'><head><meta charset='utf-8'>"
        "<title>中国传统文化与实修知识体系</title><style>" + CSS + "</style></head><body>"
        + "".join(parts)
        + "</body></html>"
    )
    with open(OUT_HTML, "w", encoding="utf-8") as f:
        f.write(html)

    print("已生成:", OUT_HTML)
    print("解析到模块数:", len(parse_modules(toc_text)))


if __name__ == "__main__":
    main()
