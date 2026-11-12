"""IPL担当 相談ログ（Excel）を作る。
使い方：python3 tools/build/consult_log.py <出力.xlsx>
"""
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter

FONT = "Meiryo"
DK, TEAL, HEAD = "1F3A4D", "0F7C8C", "DCE9ED"
thin = Side(style="thin", color="A9C7CF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
N = 500


def f(size=10, bold=False, color="1E2B33", italic=False):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic)


wb = Workbook()
ws0 = wb.active
ws0.title = "使い方"
ws0.sheet_view.showGridLines = False
ws0.column_dimensions["A"].width = 4
ws0.column_dimensions["B"].width = 24
ws0.column_dimensions["C"].width = 90
ws0["B1"] = "IPL担当 相談ログ"
ws0["B1"].font = f(16, True, DK)
ws0["B2"] = "IPL担当が受けた相談を1行ずつ記録し、月1回のIPL担当会で標準フローに戻す候補を選びます。"
ws0["B2"].font = f(10, False, "5B7083", True)
rows = [
    ("記録するとき（IPL担当）", None),
    ("いつ", "相談を受けた日のうちに。1件5分を目安に。"),
    ("いちばん大事な列", "「手順書に書いてあったか」。書いていなかった・分かりにくかった相談が、標準フローを直す手がかりになる。"),
    ("書かないこと", "相談者の個人名、案件名・顧客名、出願前の発明の内容。案件はコードで書く。"),
    ("月1回のIPL担当会", None),
    ("選ぶ", "「標準化の候補」が「はい」「会で相談」の行を見て、修正リストに登録するものを選ぶ。"),
    ("登録したら", "修正リストのNoを「修正リストNo」に書く。修正リスト側の「関連No」には相談ログのNoを書く。"),
    ("集計", "「集計」シートで、STEP別の件数と「書いていなかった」「分かりにくい」の件数を見る。"),
    ("保存場所", "IPL担当と標準化チームだけが見られる場所に置く。AIエージェントのナレッジには入れない。"),
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
    r += 1
ws0.page_setup.fitToWidth = 1
ws0.page_setup.fitToHeight = 0
ws0.sheet_properties.pageSetUpPr.fitToPage = True

ws = wb.create_sheet("相談ログ")
ws.sheet_view.showGridLines = False
ws["A1"] = "IPL担当 相談ログ"
ws["A1"].font = f(14, True, DK)
cols = [("No", 6), ("日付", 11), ("IPL担当", 12), ("相談者の部署", 14), ("案件コード", 10), ("STEP", 9),
        ("手順番号", 9), ("相談の種類", 16), ("相談の内容", 36), ("答えたこと", 36), ("手順書に書いてあったか", 16),
        ("標準化の候補", 10), ("修正リストNo", 10), ("タグ（業界・分野など）", 18), ("備考", 20)]
HR = 3
for j, (h, w) in enumerate(cols, 1):
    c = ws.cell(row=HR, column=j, value=h)
    c.font = f(10, True)
    c.fill = PatternFill("solid", fgColor=HEAD)
    c.alignment = Alignment(wrap_text=True, vertical="center")
    c.border = BORDER
    ws.column_dimensions[get_column_letter(j)].width = w
ws.row_dimensions[HR].height = 32
col = {h: get_column_letter(j) for j, (h, _) in enumerate(cols, 1)}
first, last = HR + 1, HR + N
for i in range(N):
    rr = first + i
    for j in range(1, len(cols) + 1):
        c = ws.cell(row=rr, column=j)
        c.font = f(10)
        c.border = BORDER
        c.alignment = Alignment(wrap_text=True, vertical="top")
    ws[f"{col['No']}{rr}"] = f'=IF({col["相談の内容"]}{rr}="","",ROW()-{HR})'
    ws[f"{col['日付']}{rr}"].number_format = "yyyy/mm/dd"
    ws[f"{col['手順番号']}{rr}"].number_format = "@"
lists = {
    "STEP": ["全体", "STEP0", "STEP1", "STEP2", "STEP3", "STEP4", "STEP5", "STEP6", "STEP7", "STEP8", "エージェント", "その他"],
    "相談の種類": ["手順が分からない", "ひな形の書き方", "調べ方・検索式", "判断の迷い", "エージェントの使い方・誤り", "その他"],
    "手順書に書いてあったか": ["書いてあった", "分かりにくい", "書いていなかった"],
    "標準化の候補": ["はい", "いいえ", "会で相談"],
}
for h, opts in lists.items():
    dv = DataValidation(type="list", formula1='"' + ",".join(opts) + '"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"{col[h]}{first}:{col[h]}{last}")
K = col["手順書に書いてあったか"]
ws.conditional_formatting.add(f"A{first}:{get_column_letter(len(cols))}{last}",
                              FormulaRule(formula=[f'${K}{first}="書いていなかった"'], fill=PatternFill("solid", fgColor="FBE3D6")))
ws.conditional_formatting.add(f"A{first}:{get_column_letter(len(cols))}{last}",
                              FormulaRule(formula=[f'${K}{first}="分かりにくい"'], fill=PatternFill("solid", fgColor="FFF4E0")))
ws.auto_filter.ref = f"A{HR}:{get_column_letter(len(cols))}{last}"
ws.freeze_panes = ws[f"{col['STEP']}{first}"]

# 集計
sm = wb.create_sheet("集計")
sm.sheet_view.showGridLines = False
sm["A1"] = "STEP別の集計（自動）"
sm["A1"].font = f(14, True, DK)
heads = ["STEP", "相談件数", "書いてあった", "分かりにくい", "書いていなかった", "標準化の候補（はい）", "足りなさの割合"]
for j, h in enumerate(heads, 1):
    c = sm.cell(row=3, column=j, value=h)
    c.font = f(10, True)
    c.fill = PatternFill("solid", fgColor=HEAD)
    c.border = BORDER
    c.alignment = Alignment(wrap_text=True, vertical="center")
    sm.column_dimensions[get_column_letter(j)].width = [14, 10, 12, 12, 14, 16, 14][j - 1]
S = f"相談ログ!${col['STEP']}${first}:${col['STEP']}${last}"
Kr = f"相談ログ!${K}${first}:${K}${last}"
Cr = f"相談ログ!${col['標準化の候補']}${first}:${col['標準化の候補']}${last}"
for i, s in enumerate(lists["STEP"] + ["合計"], 4):
    sm.cell(row=i, column=1, value=s)
    if s == "合計":
        for j, L in enumerate("BCDEF", 2):
            sm.cell(row=i, column=j, value=f"=SUM({L}4:{L}{i-1})")
    else:
        sm.cell(row=i, column=2, value=f'=COUNTIF({S},A{i})')
        sm.cell(row=i, column=3, value=f'=COUNTIFS({S},A{i},{Kr},"書いてあった")')
        sm.cell(row=i, column=4, value=f'=COUNTIFS({S},A{i},{Kr},"分かりにくい")')
        sm.cell(row=i, column=5, value=f'=COUNTIFS({S},A{i},{Kr},"書いていなかった")')
        sm.cell(row=i, column=6, value=f'=COUNTIFS({S},A{i},{Cr},"はい")')
    sm.cell(row=i, column=7, value=f'=IF(B{i}=0,"",(D{i}+E{i})/B{i})').number_format = "0%"
    for j in range(1, 8):
        sm.cell(row=i, column=j).border = BORDER
        sm.cell(row=i, column=j).font = f(10, s == "合計")
sm.cell(row=4 + len(lists["STEP"]) + 2, column=1, value="足りなさの割合＝（分かりにくい＋書いていなかった）÷相談件数。改訂を重ねると下がるのが望ましい。").font = f(9, False, "5B7083", True)

wb.active = 1
wb.save(sys.argv[1])
print("saved", sys.argv[1])
