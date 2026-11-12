"""IPL標準フロー 修正リスト（Excel）を作る。
使い方：python3 tools/build/revision_list.py <出力.xlsx>
"""
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter

FONT = "Meiryo"
DK, TEAL, HEAD, LT = "1F3A4D", "0F7C8C", "DCE9ED", "EEF3F5"
thin = Side(style="thin", color="A9C7CF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
NROWS = 500


def f(size=10, bold=False, color="1E2B33", italic=False, underline=None):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic, underline=underline)


wb = Workbook()

# ---------------- 使い方
ws0 = wb.active
ws0.title = "使い方"
ws0.sheet_view.showGridLines = False
ws0.column_dimensions["A"].width = 4
ws0.column_dimensions["B"].width = 26
ws0.column_dimensions["C"].width = 90
ws0["B1"] = "IPL標準フロー 修正リスト"
ws0["B1"].font = f(16, True, DK)
ws0["B2"] = "手順書・ワークシート・説明資料・エージェントへの修正案を、誰でも書き込める一覧です。判断の結果と理由を残し、同じ議論のくり返しを防ぎます。"
ws0["B2"].font = f(10, False, "5B7083", True)

rows = [
    ("書く人（だれでも）", None),
    ("1. まず検索する", "「修正リスト」シートで、対象STEPと手順番号のフィルター（見出しの▼）を使い、同じ箇所への過去の修正がないか確かめる。「同じ箇所の件数」が2以上なら過去の記録がある。"),
    ("2. 1行に1件書く", "灰色の見出しの列（受付日〜理由・背景）を書く。1つの意見に複数の修正が含まれる場合は行を分ける。"),
    ("3. 手順番号は「1-3」の形", "STEP1の手順1-3なら「1-3」。全体に関わる場合は空欄。手順番号は改訂で変わることがあるため、「手順の名前」の列も書く。"),
    ("4. 過去の記録があった場合", "「関連No」に過去のNoを書き、何が違うのか（状況が変わった、新しい理由がある）を理由・背景に書く。"),
    ("判断する人（標準化チーム）", None),
    ("5. 状態を更新する", "受付 → 検討中 → 採用／一部採用／見送り／保留 → 反映済み。見送り・保留の場合も、判断理由を必ず書く（同じ提案が再び出たときの答えになる）。"),
    ("6. 反映したら", "「反映した版」「反映日」を書き、状態を「反映済み」にする。手順書の改訂履歴にもNoを書く。"),
    ("7. 意見が食い違うとき", "片方を削除せず、両方の行を残し、判断理由にどちらをなぜ選んだかを書く。"),
    ("注意", None),
    ("個人名の扱い", "提案者の名前は社内での確認のために書く。このファイルを社外に出したり、AIエージェントのナレッジに入れたりしない。"),
    ("社外秘の情報", "具体的な案件名・顧客名・出願前の発明の内容は書かない。必要なら「案件A」のように書く。"),
    ("並べ替え・フィルター", "見出しの▼で並べ替え・絞り込みができる。共同編集中に並べ替えると他の人の表示も変わるため、絞り込み（フィルター）を使うか「シートビュー」を使う。"),
]
r = 4
for a, b in rows:
    if b is None:
        ws0.cell(row=r, column=2, value=a).font = f(12, True, TEAL)
    else:
        ws0.cell(row=r, column=2, value=a).font = f(10, True)
        c = ws0.cell(row=r, column=3, value=b)
        c.font = f(10)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws0.row_dimensions[r].height = 30
    r += 1

ws0.page_setup.fitToWidth = 1
ws0.page_setup.fitToHeight = 0
ws0.sheet_properties.pageSetUpPr.fitToPage = True

# ---------------- 修正リスト
ws = wb.create_sheet("修正リスト")
ws.sheet_view.showGridLines = False
cols = [
    # (見出し, 幅, 記入者)  記入者: P=提案者 J=判断者 A=自動
    ("No", 6, "A"), ("受付日", 11, "P"), ("提案者", 12, "P"), ("所属", 14, "P"),
    ("対象", 12, "P"), ("手順番号", 9, "P"), ("手順の名前", 22, "P"), ("修正の種類", 10, "P"),
    ("今の記載（要約）", 30, "P"), ("修正案", 40, "P"), ("理由・背景", 36, "P"), ("優先度", 8, "P"),
    ("関連No", 9, "P"), ("同じ箇所の件数", 9, "A"),
    ("状態", 10, "J"), ("判断者", 12, "J"), ("判断日", 11, "J"), ("判断理由", 36, "J"),
    ("反映した版", 10, "J"), ("反映日", 11, "J"), ("反映したファイル", 22, "J"), ("備考", 24, "J"),
]
ws["A1"] = "IPL標準フロー 修正リスト"
ws["A1"].font = f(14, True, DK)
ws["A2"] = "灰色の見出し＝提案する人が書く／青緑の見出し＝標準化チームが書く／No と「同じ箇所の件数」は自動"
ws["A2"].font = f(9, False, "5B7083", True)
HR = 4
for j, (h, w, who) in enumerate(cols, 1):
    c = ws.cell(row=HR, column=j, value=h)
    c.font = f(10, True, "FFFFFF" if who == "J" else "1E2B33")
    c.fill = PatternFill("solid", fgColor=TEAL if who == "J" else (HEAD if who == "P" else LT))
    c.alignment = Alignment(wrap_text=True, vertical="center")
    c.border = BORDER
    ws.column_dimensions[get_column_letter(j)].width = w
ws.row_dimensions[HR].height = 34

col = {h: get_column_letter(j) for j, (h, _, _) in enumerate(cols, 1)}
first, last = HR + 1, HR + NROWS
for i in range(NROWS):
    rr = first + i
    for j in range(1, len(cols) + 1):
        c = ws.cell(row=rr, column=j)
        c.font = f(10)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        c.border = BORDER
    ws[f"{col['No']}{rr}"] = f'=IF({col["修正案"]}{rr}="","",ROW()-{HR})'
    ws[f"{col['同じ箇所の件数']}{rr}"] = (
        f'=IF(OR({col["対象"]}{rr}="",{col["修正案"]}{rr}=""),"",'
        f'COUNTIFS(${col["対象"]}${first}:${col["対象"]}${last},{col["対象"]}{rr},'
        f'${col["手順番号"]}${first}:${col["手順番号"]}${last},{col["手順番号"]}{rr}&""))'
    )
    for h in ("受付日", "判断日", "反映日"):
        ws[f"{col[h]}{rr}"].number_format = "yyyy/mm/dd"
    ws[f"{col['手順番号']}{rr}"].number_format = "@"  # 「1-3」が日付に変わらないよう文字列にする
    ws[f"{col['反映した版']}{rr}"].number_format = "@"

lists = {
    "対象": ["全体", "STEP0", "STEP1", "STEP2", "STEP3", "STEP4", "STEP5", "STEP6", "STEP7", "STEP8", "ワークシート", "説明資料", "エージェント", "その他"],
    "修正の種類": ["追加", "変更", "削除", "誤り", "分かりにくい", "質問"],
    "優先度": ["高", "中", "低"],
    "状態": ["受付", "検討中", "採用", "一部採用", "見送り", "保留", "反映済み"],
}
for h, opts in lists.items():
    dv = DataValidation(type="list", formula1='"' + ",".join(opts) + '"', allow_blank=True)
    dv.error = "リストから選んでください"
    dv.showErrorMessage = True
    ws.add_data_validation(dv)
    dv.add(f"{col[h]}{first}:{col[h]}{last}")
dvd = DataValidation(type="date", operator="greaterThan", formula1="DATE(2026,1,1)", allow_blank=True)
dvd.error = "日付（例：2026/11/12）を入れてください"
dvd.showErrorMessage = True
ws.add_data_validation(dvd)
for h in ("受付日", "判断日", "反映日"):
    dvd.add(f"{col[h]}{first}:{col[h]}{last}")

# 状態で色分け
rng = f"A{first}:{get_column_letter(len(cols))}{last}"
S = col["状態"]
for val, color in (("見送り", "F2F2F2"), ("反映済み", "E8F3EA"), ("保留", "FFF4E0")):
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'${S}{first}="{val}"'], fill=PatternFill("solid", fgColor=color)))
# 同じ箇所の件数が2以上なら強調
K = col["同じ箇所の件数"]
ws.conditional_formatting.add(f"{K}{first}:{K}{last}", FormulaRule(formula=[f"AND(ISNUMBER({K}{first}),{K}{first}>=2)"], fill=PatternFill("solid", fgColor="FBE3D6"), font=Font(name=FONT, bold=True, color="B4472F")))

ws.auto_filter.ref = f"A{HR}:{get_column_letter(len(cols))}{last}"
ws.freeze_panes = ws[f"{col['対象']}{first}"]
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.sheet_properties.pageSetUpPr.fitToPage = True

# ---------------- 記入例
ex = wb.create_sheet("記入例")
ex.sheet_view.showGridLines = False
ex["A1"] = "記入例（架空）"
ex["A1"].font = f(14, True, DK)
ex["A2"] = "このシートは見本です。実際の記入は「修正リスト」シートに書いてください。"
ex["A2"].font = f(9, False, "5B7083", True)
heads = ["No", "受付日", "提案者", "所属", "対象", "手順番号", "手順の名前", "修正の種類", "今の記載（要約）", "修正案", "理由・背景", "優先度", "関連No", "状態", "判断者", "判断日", "判断理由", "反映した版"]
samples = [
    [1, "2026/11/20", "研究 Aさん", "研究開発部", "STEP1", "1-3", "ファーストユーザーを1つ決める", "追加", "5項目（困りごとの強さ等）で仮に評価", "評価項目に「自社技術の適合度」を足す", "技術的に合わない会社を候補に入れてしまった", "中", "", "採用", "標準化チーム", "2026/11/27", "他部署からも同様の声。6項目にする", "v1.2"],
    [2, "2026/11/21", "事業 Bさん", "事業企画部", "STEP6", "6-5", "自分で聞きに行く", "変更", "本人がヒアリングに行く", "営業担当が聞いてもよいことにする", "顧客との窓口が営業に限られている", "高", "", "一部採用", "標準化チーム", "2026/11/27", "本人の同席を条件に、営業が主に話すことは可。本人が仮説を説明する部分は残す", "v1.2"],
    [3, "2027/02/03", "研究 Cさん", "研究開発部", "STEP6", "6-5", "自分で聞きに行く", "変更", "本人がヒアリングに行く", "営業担当だけで聞いてよいことにする", "時間が取れない", "中", "2", "見送り", "標準化チーム", "2027/02/10", "No.2で検討済み。本人が仮説を説明しないと本音が聞けないため。同席の形で対応", ""],
]
for j, h in enumerate(heads, 1):
    c = ex.cell(row=4, column=j, value=h)
    c.font = f(10, True)
    c.fill = PatternFill("solid", fgColor=HEAD)
    c.border = BORDER
    c.alignment = Alignment(wrap_text=True, vertical="center")
    ex.column_dimensions[get_column_letter(j)].width = [6, 11, 12, 12, 8, 8, 20, 9, 24, 28, 28, 7, 7, 9, 12, 11, 34, 9][j - 1]
for i, row in enumerate(samples, 5):
    for j, v in enumerate(row, 1):
        c = ex.cell(row=i, column=j, value=v)
        c.font = f(10)
        c.border = BORDER
        c.alignment = Alignment(wrap_text=True, vertical="top")
    ex.row_dimensions[i].height = 48
ex.cell(row=9, column=1, value="No.3 のように、過去に検討済みの提案は「関連No」で過去の行を示し、判断理由に前回の結論を書くと、同じ議論をくり返さずに済みます。").font = f(10, False, TEAL)

wb.active = 1
wb.save(sys.argv[1])
print("saved", sys.argv[1])
