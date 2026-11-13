"""docs/ の Markdown から Word 版の手順書を作る。
使い方（リポジトリの一番上で）：python3 tools/build/build_manual.py <フローチャート画像.png> <出力.docx>
必要なもの：pandoc、python-docx、tools/build/reference.docx（スタイル定義）
"""
import os, re, subprocess, sys, tempfile, shutil
import docx
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

img, out = sys.argv[1], sys.argv[2]
work = tempfile.mkdtemp()
shutil.copy(img, os.path.join(work, "flow.png"))

def clean(t):
    t = re.sub(r"^> IPL標準フロー.*\n", "", t, flags=re.M)
    t = re.sub(r"```mermaid.*?```", "@@FLOW@@", t, flags=re.S)
    t = re.sub(r"\[([^\]]+)\]\((?!http)[^)]+\)", r"\1", t)
    t = re.sub(r"(?:\.\./)*templates/(\d\d_[^\s）)、。]+?)\.(?:md|csv)", r"ワークシート集（Excel）の「\1」シート", t)
    t = re.sub(r"(?:\.\./)*templates/", "ワークシート集（Excel）", t)
    t = re.sub(r"(?:\.\./)*(?:docs/)?(?:steps/)?STEP(\d)_[^\s）)、。]+?\.md", r"STEP\1の章", t)
    t = re.sub(r"(?:\.\./)*(?:docs/)?(0[01]|9\d)_([^\s）)、。]+?)\.md", r"「\2」の章", t)
    t = t.replace("（steps/）", "（第1部の各STEPの章）")
    t = re.sub(r"(?:\.\./)*(copilot-studio|sharing)/([^\s）)、。]+?)\.(md|txt)", r"リポジトリの \1/\2.\3", t)
    t = t.replace("`python3 tools/check_flow.py`", "python3 tools/check_flow.py（リポジトリ内）")
    t = re.sub(r"^(#+) ", lambda m: "#" + m.group(1) + " ", t, flags=re.M)  # 1段下げる
    return t.replace("@@FLOW@@", "![IPL標準フロー フローチャート（実線＝進む、点線＝戻る）](flow.png){width=100%}")

part1 = ["docs/00_はじめに.md", "docs/01_全体像.md"] + ["docs/steps/" + f for f in sorted(os.listdir("docs/steps"))]
part2 = ["docs/90_発明塾手法との対応.md", "docs/91_呼び方の決め方.md", "docs/92_フロー検証の進め方.md", "docs/93_体制・品質・蓄積の論点.md", "docs/99_出典と改訂履歴.md"]
md = ['---', 'title: "IPL標準フロー 手順書"', 'subtitle: "事業・研究の一人ひとりが自らIPランドスケープを進めるために　v1.2（たたき台）"',
      'date: "2026年11月"', 'lang: ja-JP', 'toc-title: "目次"', '---', '', '# 第1部　利用者向け（事業・研究の担当者）', '']
md += [clean(open(p, encoding="utf-8").read()) + "\n" for p in part1]
md += ["# 第2部　標準化チーム・IPL担当向け（付録）", ""]
md += [clean(open(p, encoding="utf-8").read()) + "\n" for p in part2]
src = os.path.join(work, "manual.md")
open(src, "w", encoding="utf-8").write("\n".join(md))
ref = os.path.abspath("tools/build/reference.docx")
subprocess.run(["pandoc", "manual.md", "-o", "manual.docx", "--reference-doc", ref, "--toc", "--toc-depth=2",
                "-f", "markdown-implicit_figures"], cwd=work, check=True)

d = docx.Document(os.path.join(work, "manual.docx"))
for t in d.tables:
    for c in t.rows[0].cells:
        tcPr = c._tc.get_or_add_tcPr(); sh = OxmlElement("w:shd")
        sh.set(qn("w:val"), "clear"); sh.set(qn("w:fill"), "DCE9ED"); tcPr.append(sh)
        for p in c.paragraphs:
            for r in p.runs:
                r.bold = True
seen = 0
for p in d.paragraphs:
    n = p.style.name.lower()
    if n == "heading 1":
        p.paragraph_format.page_break_before = True; seen = 1
    elif n == "heading 2":
        if seen == 1: seen = 2
        else: p.paragraph_format.page_break_before = True
u = OxmlElement("w:updateFields"); u.set(qn("w:val"), "true"); d.settings.element.append(u)
d.core_properties.title = "IPL標準フロー 手順書 v1.2"
d.save(out)
print("saved", out)
