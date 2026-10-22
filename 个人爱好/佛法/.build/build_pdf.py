# -*- coding: utf-8 -*-
"""把《佛法智慧资料库》36 篇 Markdown 合并为一个带中文排版的 HTML，
供 Chrome 无头模式打印为 PDF。"""
import glob
import os
import re
import markdown

BASE = r"D:\zwr\idea\project\111\innerworld\个人爱好\佛法"
OUT_HTML = os.path.join(BASE, ".build", "佛法智慧资料库.html")

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
.cover { text-align: center; padding: 90mm 0 0 0; page-break-after: always; }
.cover h1 { font-size: 34pt; color: #5b4636; margin: 0 0 12px 0; letter-spacing: 6px; }
.cover .sub { font-size: 14pt; color: #8a6d4b; letter-spacing: 2px; }
.cover .rule { width: 120px; height: 3px; background: #b49a7e; margin: 28px auto; }
.cover .meta { font-size: 10.5pt; color: #9a8b78; margin-top: 18px; }
.module-divider { page-break-before: always; border-top: 2px solid #d8c7b3; margin-top: 2px; padding-top: 14px; }
.module-divider h2 { color: #5b4636; font-size: 18pt; margin: 0 0 6px 0; }
.module-divider .goal { color: #7a6a55; font-size: 10.5pt; font-style: italic; margin: 0 0 10px 0; }
section.theme { page-break-before: always; }
h1 { font-size: 20pt; color: #4a3a2c; border-bottom: 1px solid #d8c7b3; padding-bottom: 6px; }
h2 { font-size: 15pt; color: #5b4636; }
h3 { font-size: 12.5pt; color: #6b5640; }
blockquote {
  margin: 10px 0; padding: 10px 16px;
  background: #f5efe6; border-left: 4px solid #b49a7e; color: #5a534a;
  border-radius: 0 6px 6px 0;
}
blockquote p { margin: 4px 0; }
strong { color: #4a3a2c; }
table { border-collapse: collapse; width: 100%; margin: 12px 0; font-size: 10pt; }
th, td { border: 1px solid #d8c7b3; padding: 6px 9px; vertical-align: top; text-align: left; }
th { background: #efe4d2; color: #4a3a2c; }
tr:nth-child(even) td { background: #faf6f0; }
code {
  font-family: Consolas, "Courier New", monospace;
  background: #f2ede4; color: #5b4636; padding: 1px 5px; border-radius: 3px; font-size: 9.5pt;
}
pre { background: #f6f1e8; border: 1px solid #e0d4c2; padding: 10px 14px; border-radius: 6px; overflow: hidden; }
pre code { background: none; padding: 0; font-size: 9.5pt; white-space: pre-wrap; }
ul, ol { margin: 6px 0 6px 0; padding-left: 22px; }
li { margin: 3px 0; }
hr { border: none; border-top: 1px solid #d8c7b3; margin: 18px 0; }
a { color: #8a6d4b; text-decoration: none; }
.references { margin-top: 10px; font-size: 9.5pt; color: #6b6157; }
.references h3 { font-size: 11pt; }
"""


def md_to_html(text):
    return markdown.markdown(
        text,
        extensions=["tables", "fenced_code", "sane_lists"],
    )


def module_goal(lines):
    """从模块目录.md 的 '> 目标：...' 块提取目标描述。"""
    goal = ""
    for line in lines:
        m = re.match(r">\s*目标[:：]\s*(.*)", line)
        if m:
            goal = m.group(1).strip()
    return goal


def main():
    # 读取顶层目录
    with open(os.path.join(BASE, "目录.md"), "r", encoding="utf-8") as f:
        toc_lines = f.read()
    toc_html = md_to_html(toc_lines)

    parts = []
    parts.append(
        '<div class="cover">'
        '<h1>佛法智慧资料库</h1>'
        '<div class="rule"></div>'
        '<div class="sub">把佛法智慧应用到生活与工作</div>'
        '<div class="meta">认识世界 · 心态情绪 · 工作职场 · 处世为人 · 日常生活 · 逆境苦难</div>'
        "</div>"
    )
    parts.append('<section class="toc">' + toc_html + "</section>")

    # 按模块顺序处理
    module_dirs = sorted(glob.glob(os.path.join(BASE, "模块*")))
    for mod_dir in module_dirs:
        mod_name = os.path.basename(mod_dir)
        # 模块目标：从目录.md 中该模块标题下的目标行读取
        goal = ""
        lines = toc_lines.splitlines()
        for i, line in enumerate(lines):
            if line.strip().startswith("##") and mod_name in line:
                # 向后找目标行
                for j in range(i + 1, min(i + 6, len(lines))):
                    m = re.match(r">\s*目标[:：]\s*(.*)", lines[j])
                    if m:
                        goal = m.group(1).strip()
                        break
                break

        divider = (
            '<div class="module-divider">'
            f"<h2>{mod_name}</h2>"
            + (f'<div class="goal">目标：{goal}</div>' if goal else "")
            + "</div>"
        )
        parts.append(divider)

        # 模块内主题文件（按编号排序）
        files = sorted(
            glob.glob(os.path.join(mod_dir, "*.md")),
            key=lambda p: os.path.basename(p),
        )
        for fp in files:
            with open(fp, "r", encoding="utf-8") as f:
                content = f.read()
            body = md_to_html(content)
            parts.append(f'<section class="theme">{body}</section>')

    html = (
        "<!DOCTYPE html><html lang='zh-CN'><head><meta charset='utf-8'>"
        "<title>佛法智慧资料库</title><style>" + CSS + "</style></head><body>"
        + "".join(parts)
        + "</body></html>"
    )

    with open(OUT_HTML, "w", encoding="utf-8") as f:
        f.write(html)

    print("已生成:", OUT_HTML)
    print("模块数:", len(module_dirs))


if __name__ == "__main__":
    main()
