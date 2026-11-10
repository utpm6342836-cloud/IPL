"""IPL標準フローの手順書が、決めた構造ルールを満たしているかを確かめる。

使い方：リポジトリの一番上で  python3 tools/check_flow.py
（docs/92_フロー検証の進め方.md の「検証1：構造チェック」を自動で行う部分）
"""
import glob
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
REQUIRED = {
    "目的": "## このSTEPでやること",
    "始める条件": "## 始める前にそろえるもの",
    "できあがるもの": "## できあがるもの",
    "チェック": "## 次に進む前のチェック",
    "困ったら": "## 困ったら",
    "発明塾のコツ": "## 発明塾のコツ",
    "根拠": "## 根拠",
}

problems = 0
for path in sorted(glob.glob(os.path.join(ROOT, "docs", "steps", "STEP*.md"))):
    text = open(path, encoding="utf-8").read()
    name = os.path.basename(path)
    issues = [f"見出し「{k}」がない" for k, v in REQUIRED.items() if v not in text]
    steps = re.findall(r"^### (\d-\d+.*)$", text, flags=re.M)
    no_time = [s for s in steps if not re.search(r"（[^）]*(分|時間|会議)[^）]*）", s)]
    issues += [f"目安時間がない手順：{s}" for s in no_time]
    if "エージェントに頼めること" not in text:
        issues.append("エージェントに頼めること／自分でやることの分担がない")
    if "templates/" not in text:
        issues.append("ひな形へのリンクがない")
    for target in re.findall(r"\]\(([^)]+)\)", text):
        if target.startswith("http"):
            continue
        p = os.path.normpath(os.path.join(os.path.dirname(path), target.split("#")[0]))
        if not os.path.exists(p):
            issues.append(f"リンク切れ：{target}")
    problems += len(issues)
    print(f"{name}: " + ("OK" if not issues else "; ".join(issues)))

sys.exit(1 if problems else 0)
