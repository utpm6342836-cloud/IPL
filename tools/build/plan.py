"""IPL標準フロー 推進計画（Excel）を作る。
使い方：python3 tools/build/plan.py <出力.xlsx>
日程は 2026-11-12 時点の案。週の列は月曜始まり。
"""
import sys
import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

FONT = "Meiryo"
DK, TEAL, HEAD = "1F3A4D", "0F7C8C", "DCE9ED"
thin = Side(style="thin", color="A9C7CF")
hair = Side(style="hair", color="D5DEE2")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
D = dt.date


def f(size=10, bold=False, color="1E2B33", italic=False):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic)


CAT = {  # 区分: 色
    "A 準備・確認": "5B7083",
    "B 手順書の改訂": "0F7C8C",
    "C AIエージェント": "D9862B",
    "D パイロット": "7A9A3A",
    "E 配布": "B4472F",
    "F 2027年度": "A9C7CF",
}
# (ID, 区分, 作業, 開始, 終了, 担当, 前提, できあがるもの, マイルストーン?)
T = [
    ("A1", "A 準備・確認", "推進計画の確定・関係者の合意", D(2026, 11, 16), D(2026, 11, 27), "標準化チーム", "", "推進計画（本ファイル）", False),
    ("A2", "A 準備・確認", "ヒアリング対象者の選定・日程調整", D(2026, 11, 16), D(2026, 11, 27), "標準化チーム", "A1", "ヒアリング計画（日程入り）", False),
    ("A3", "A 準備・確認", "修正リスト・相談ログを SharePoint に公開", D(2026, 11, 16), D(2026, 11, 20), "標準化チーム", "", "修正リスト、相談ログ", False),
    ("A4", "A 準備・確認", "発明塾資料の利用条件の確認（U1）", D(2026, 11, 16), D(2026, 12, 11), "標準化チーム", "", "利用可否の回答", False),
    ("A5", "A 準備・確認", "Web検索で検索語が外部に送られることの社内確認（U12）", D(2026, 11, 16), D(2026, 12, 25), "情報システム・法務", "", "可否と条件", False),
    ("A6", "A 準備・確認", "特許庁ガイドブックとの照合（V3）", D(2026, 11, 30), D(2026, 12, 11), "標準化チーム", "", "照合表", False),
    ("B1", "B 手順書の改訂", "ヒアリング①：STEP0〜2（事業・研究の担当者、依頼者）", D(2026, 11, 30), D(2026, 12, 11), "標準化チーム", "A2", "修正リストへの登録", False),
    ("B2", "B 手順書の改訂", "ヒアリング②：STEP3〜5（IPL担当、調査担当）", D(2026, 12, 7), D(2026, 12, 18), "標準化チーム", "A2", "修正リストへの登録", False),
    ("B3", "B 手順書の改訂", "ヒアリング③：STEP6〜8（事業・研究の担当者、意思決定者）", D(2026, 12, 14), D(2026, 12, 25), "標準化チーム", "A2", "修正リストへの登録", False),
    ("B4", "B 手順書の改訂", "修正の判断・反映（手順書・ワークシート・説明資料）", D(2026, 12, 14), D(2026, 12, 25), "標準化チーム", "B1〜B3", "v1.2", False),
    ("M1", "B 手順書の改訂", "◆ 手順書 v1.2 確定（12月の目標）", D(2026, 12, 25), D(2026, 12, 25), "標準化チーム", "B4", "手順書 v1.2", True),
    ("C1", "C AIエージェント", "案内役のα版を作る（ナレッジ＝手順書 v1.2）", D(2026, 12, 14), D(2027, 1, 15), "エージェント作成担当", "B4（下書きで先行）", "案内役α版", False),
    ("C2", "C AIエージェント", "子エージェント「Web調査」の設定・テスト", D(2027, 1, 12), D(2027, 1, 29), "エージェント作成担当", "A5", "Web調査（一次情報サイト限定）", False),
    ("C3", "C AIエージェント", "検証エージェント（チェック役）の試行", D(2027, 1, 19), D(2027, 2, 5), "エージェント作成担当", "C1", "チェック役（試行）", False),
    ("C4", "C AIエージェント", "公開前テスト（テストセット28問、合格基準の確認）", D(2027, 1, 25), D(2027, 2, 12), "標準化チーム", "C1〜C3", "テスト結果", False),
    ("M2", "C AIエージェント", "◆ β版をパイロット参加者に公開", D(2027, 1, 29), D(2027, 1, 29), "エージェント作成担当", "C1・C2", "β版", True),
    ("D1", "D パイロット", "パイロット案件（2〜3件）と参加者の選定", D(2026, 12, 14), D(2027, 1, 8), "標準化チーム", "", "パイロット計画", False),
    ("D2", "D パイロット", "パイロット実施（IPL担当が伴走、ワークシート・β版を使う）", D(2027, 1, 12), D(2027, 3, 5), "参加者・IPL担当", "D1・M2", "判断メモ、振り返りアンケート", False),
    ("D3", "D パイロット", "月1回 IPL担当会（相談ログから候補を選ぶ）1/28・2/25・3/25", D(2027, 1, 25), D(2027, 3, 26), "IPL担当・標準化チーム", "A3", "修正リストへの登録", False),
    ("D4", "D パイロット", "パイロットの振り返りと v1.3 への反映", D(2027, 3, 1), D(2027, 3, 12), "標準化チーム", "D2", "手順書 v1.3、エージェント更新", False),
    ("E1", "E 配布", "利用ガイド・動画の作成（動画シリーズ構成案を使う）", D(2027, 2, 15), D(2027, 3, 12), "標準化チーム", "B4", "利用ガイド、動画", False),
    ("E2", "E 配布", "説明会（興味のある人向け）", D(2027, 3, 15), D(2027, 3, 19), "標準化チーム", "E1・D4", "説明会", False),
    ("M3", "E 配布", "◆ 興味のある人へ配布開始（限定公開）", D(2027, 3, 22), D(2027, 3, 22), "標準化チーム", "C4合格・A4・A5", "AIツール配布", True),
    ("E3", "E 配布", "年度の振り返りと2027年度計画", D(2027, 3, 22), D(2027, 3, 31), "標準化チーム", "", "2027年度計画", False),
    ("F1", "F 2027年度", "配布範囲の拡大、月次の品質点検（会話の抜き取り）", D(2027, 4, 1), D(2027, 6, 30), "標準化チーム", "M3", "", False),
    ("F2", "F 2027年度", "四半期改訂（v1.4）", D(2027, 6, 1), D(2027, 6, 25), "標準化チーム", "", "v1.4", False),
    ("F3", "F 2027年度", "特許調査エージェント（API）の検討", D(2027, 4, 1), D(2027, 6, 30), "標準化チーム", "U13", "検討結果", False),
]
HOLIDAY_WEEKS = {D(2026, 12, 28), D(2027, 1, 4)}  # 年末年始（目安）

wb = Workbook()
ws = wb.active
ws.title = "全体計画"
ws.sheet_view.showGridLines = False
ws["A1"] = "IPL標準フロー 推進計画（2026年度）"
ws["A1"].font = f(16, True, DK)
ws["A2"] = "2026-11-12 時点の案。◆はマイルストーン。灰色の列は年末年始（目安）。状態は▼から選ぶ。"
ws["A2"].font = f(9, False, "5B7083", True)
fixed = [("ID", 5), ("区分", 13), ("作業", 46), ("開始", 10), ("終了", 10), ("担当", 15), ("前提", 13), ("できあがるもの", 22), ("状態", 8)]
start = D(2026, 11, 16)
weeks = []
w = start
while w <= D(2027, 6, 28):
    weeks.append(w)
    w += dt.timedelta(days=7)
HR = 5
# month row
for j, (h, wd) in enumerate(fixed, 1):
    c = ws.cell(row=HR, column=j, value=h)
    c.font = f(10, True)
    c.fill = PatternFill("solid", fgColor=HEAD)
    c.border = BORDER
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.merge_cells(start_row=HR - 1, start_column=j, end_row=HR, end_column=j)
    ws.cell(row=HR - 1, column=j).fill = PatternFill("solid", fgColor=HEAD)
    ws.cell(row=HR - 1, column=j).value = h
    ws.cell(row=HR - 1, column=j).font = f(10, True)
    ws.cell(row=HR - 1, column=j).alignment = Alignment(horizontal="center", vertical="center")
    ws.column_dimensions[get_column_letter(j)].width = wd
W0 = len(fixed) + 1
prev_month = None
mstart = W0
for k, wk in enumerate(weeks):
    cidx = W0 + k
    ws.column_dimensions[get_column_letter(cidx)].width = 4.2
    c = ws.cell(row=HR, column=cidx, value=wk.day)
    c.font = f(8, False)
    c.alignment = Alignment(horizontal="center")
    c.fill = PatternFill("solid", fgColor="D9D9D9" if wk in HOLIDAY_WEEKS else HEAD)
    c.border = Border(bottom=thin)
    if (wk.year, wk.month) != prev_month:
        if prev_month is not None and cidx - 1 > mstart:
            ws.merge_cells(start_row=HR - 1, start_column=mstart, end_row=HR - 1, end_column=cidx - 1)
        mc = ws.cell(row=HR - 1, column=cidx, value=f"{wk.year}年{wk.month}月")
        mc.font = f(9, True, DK)
        mc.alignment = Alignment(horizontal="left")
        mstart = cidx
        prev_month = (wk.year, wk.month)
ws.merge_cells(start_row=HR - 1, start_column=mstart, end_row=HR - 1, end_column=W0 + len(weeks) - 1)

state_dv = DataValidation(type="list", formula1='"未着手,進行中,完了,遅れ,中止"', allow_blank=True)
ws.add_data_validation(state_dv)
r = HR + 1
for (tid, cat, name, s, e, owner, pre, out, ms) in T:
    vals = [tid, cat, name, s, e, owner, pre, out, "未着手"]
    for j, v in enumerate(vals, 1):
        c = ws.cell(row=r, column=j, value=v)
        c.font = f(10, ms or j == 3 and False, DK if ms else "1E2B33")
        if ms:
            c.font = f(10, True, DK)
        c.border = BORDER
        c.alignment = Alignment(vertical="center", wrap_text=(j in (3, 7, 8)))
        if j in (4, 5):
            c.number_format = "m/d"
    state_dv.add(f"I{r}")
    color = CAT[cat]
    for k, wk in enumerate(weeks):
        cidx = W0 + k
        cell = ws.cell(row=r, column=cidx)
        cell.border = Border(left=hair, right=hair, top=hair, bottom=hair)
        we = wk + dt.timedelta(days=6)
        if wk in HOLIDAY_WEEKS:
            cell.fill = PatternFill("solid", fgColor="EFEFEF")
        if s <= we and e >= wk:
            if ms:
                cell.value = "◆"
                cell.font = Font(name=FONT, size=11, bold=True, color="B4472F")
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.fill = PatternFill("solid", fgColor=color)
    ws.row_dimensions[r].height = 28
    r += 1
ws.freeze_panes = ws.cell(row=HR + 1, column=W0)
# legend
r += 1
ws.cell(row=r, column=3, value="区分の色").font = f(10, True)
for i, (k, v) in enumerate(CAT.items()):
    ws.cell(row=r + 1 + i, column=2).fill = PatternFill("solid", fgColor=v)
    ws.cell(row=r + 1 + i, column=3, value=k).font = f(10)
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.sheet_properties.pageSetUpPr.fitToPage = True

# ---------------- マイルストーンと判定基準
ms = wb.create_sheet("マイルストーン")
ms.sheet_view.showGridLines = False
ms["A1"] = "マイルストーンと進める条件"
ms["A1"].font = f(14, True, DK)
heads = ["◆", "日付", "内容", "進める条件（満たさなければ日程を見直す）", "満たせないときの代わりの案"]
rows = [
    ("M1", "2026/12/25", "手順書 v1.2 確定", "ヒアリング①〜③で各STEP最低2人から聞けた／修正リストの「高」の提案に判断が付いた", "聞けていないSTEPは v1.1 のまま確定し、パイロットで補う"),
    ("M2", "2027/01/29", "β版をパイロット参加者に公開", "案内役がテスト質問の「してはいけないこと」7問に全問合格", "Web調査を外し、案内役のみでβ版を出す"),
    ("M3", "2027/03/22", "興味のある人へ配布開始", "テストセット合格（禁止事項は全問、その他8割）／発明塾資料の利用条件の確認済み（U1）／外部送信の確認済み（U12。未了ならWeb調査は外す）／パイロット1件以上で判断メモまで到達", "配布対象をパイロット参加者の周辺に限る、または4月に延期"),
]
for j, h in enumerate(heads, 1):
    c = ms.cell(row=3, column=j, value=h)
    c.font = f(10, True)
    c.fill = PatternFill("solid", fgColor=HEAD)
    c.border = BORDER
    ms.column_dimensions[get_column_letter(j)].width = [6, 12, 26, 60, 44][j - 1]
for i, row in enumerate(rows, 4):
    for j, v in enumerate(row, 1):
        c = ms.cell(row=i, column=j, value=v)
        c.font = f(10)
        c.border = BORDER
        c.alignment = Alignment(wrap_text=True, vertical="top")
    ms.row_dimensions[i].height = 60

# ---------------- ヒアリング計画
hp = wb.create_sheet("ヒアリング計画")
hp.sheet_view.showGridLines = False
hp["A1"] = "ヒアリング計画（STEPごと）"
hp["A1"].font = f(14, True, DK)
hp["A2"] = "1回60分、1人につき2〜3STEPを目安。聞いた内容は修正リストに登録する（提案者＝ヒアリング相手、または聞き手）。"
hp["A2"].font = f(9, False, "5B7083", True)
heads = ["回", "対象STEP", "話を聞く相手（役割）", "人数の目安", "特に聞きたいこと", "見せるもの", "時期", "聞き手", "相手の名前", "日時", "状態", "修正リストNo"]
plan_rows = [
    ("①", "STEP0〜2", "事業部門の企画担当、研究開発のテーマリーダー、依頼する側の管理職", "3〜4人", "テーマの決め方の実際、ファーストユーザーを1社に絞れるか、事業仮説を誰と作っているか、手元の情報で何日かけられるか", "手順書 STEP0〜2、ワークシート01〜05、フローチャート", "11/30〜12/11", "", "", "", "未着手", ""),
    ("②", "STEP3〜5", "IPL担当、特許調査の担当者（ベテラン）", "2〜3人", "調査計画の書き方の実際、検索式の確かめ方、名寄せ・ファミリーの扱い、目安時間は妥当か、事業・研究の人に任せられる範囲", "手順書 STEP3〜5、ワークシート06・07", "12/7〜12/18", "", "", "", "未着手", ""),
    ("③", "STEP6〜8", "事業・研究の担当者（ヒアリング経験者）、意思決定者（部長など）", "3〜4人", "顧客ヒアリングを本人が行えるか、判断メモで決められるか（足りない項目）、Go/No Goの基準、振り返りを書く時間があるか", "手順書 STEP6〜8、ワークシート08〜13", "12/14〜12/25", "", "", "", "未着手", ""),
    ("追加", "エージェント", "パイロット参加者", "2〜3人", "エージェントに任せたいこと・任せたくないこと、回答で困った点", "β版", "2月", "", "", "", "未着手", ""),
]
for j, h in enumerate(heads, 1):
    c = hp.cell(row=4, column=j, value=h)
    c.font = f(10, True)
    c.fill = PatternFill("solid", fgColor=HEAD)
    c.border = BORDER
    c.alignment = Alignment(wrap_text=True, vertical="center")
    hp.column_dimensions[get_column_letter(j)].width = [5, 10, 30, 9, 50, 26, 12, 10, 12, 12, 8, 10][j - 1]
dv = DataValidation(type="list", formula1='"未着手,日程調整中,実施済み,記録済み"', allow_blank=True)
hp.add_data_validation(dv)
for i, row in enumerate(plan_rows, 5):
    for j, v in enumerate(row, 1):
        c = hp.cell(row=i, column=j, value=v)
        c.font = f(10)
        c.border = BORDER
        c.alignment = Alignment(wrap_text=True, vertical="top")
    hp.row_dimensions[i].height = 70
    dv.add(f"K{i}")
# empty rows for individual sessions
r0 = 5 + len(plan_rows) + 1
hp.cell(row=r0, column=1, value="1人ずつの予定（実施ごとに1行）").font = f(12, True, TEAL)
for j, h in enumerate(["回", "対象STEP", "相手の名前", "所属・役割", "日時", "場所・形式", "聞き手", "状態", "記録（修正リストNo）", "メモ"], 1):
    c = hp.cell(row=r0 + 1, column=j, value=h)
    c.font = f(10, True)
    c.fill = PatternFill("solid", fgColor=HEAD)
    c.border = BORDER
for i in range(20):
    for j in range(1, 11):
        hp.cell(row=r0 + 2 + i, column=j).border = BORDER
    dv.add(f"H{r0 + 2 + i}")

# ---------------- ヒアリングの質問
q = wb.create_sheet("ヒアリングの質問")
q.sheet_view.showGridLines = False
q["A1"] = "ヒアリングで聞くこと"
q["A1"].font = f(14, True, DK)
q["A2"] = "手順書の該当STEPを見せながら聞く。答えは「事実（今やっていること）」と「意見（こうしたい）」を分けて書く。"
q["A2"].font = f(9, False, "5B7083", True)
qs = [
    ("共通", "今の仕事で、このSTEPに当たることを実際にどうやっていますか（最近の1件で）"),
    ("共通", "手順書の手順で、抜けていること・実際と違うことはありますか"),
    ("共通", "目安時間は現実的ですか。どの手順に一番時間がかかりますか"),
    ("共通", "ワークシートで書きにくい欄、要らない欄はありますか"),
    ("共通", "AIに任せたいこと、任せたくないことはありますか"),
    ("共通", "このSTEPでよくある失敗や、後で困ったことはありますか"),
    ("STEP0", "依頼を受けたとき、何を決めるための調査かを誰とどう決めていますか。やめる条件を決めたことはありますか"),
    ("STEP1", "最初に買う一社を具体的に決められそうですか。決められない場合、何が足りませんか"),
    ("STEP2", "勝てる・儲かる・できそうのうち、いつも根拠が弱くなるのはどれですか。守れる（特許）は誰が考えていますか"),
    ("STEP3〜5", "事業・研究の人が自分で検索するとしたら、どこまでできて、どこからIPL担当が要りますか"),
    ("STEP6", "顧客に仮説を持ち込んで聞いたことはありますか。本人が行けない理由は何ですか"),
    ("STEP7", "判断メモの項目で、決めるのに足りないもの・要らないものは何ですか（意思決定者に）"),
    ("STEP8", "振り返りに使える時間はどのくらいですか。何が残っていれば次の人の役に立ちますか"),
]
for j, h in enumerate(["対象", "質問"], 1):
    c = q.cell(row=4, column=j, value=h)
    c.font = f(10, True)
    c.fill = PatternFill("solid", fgColor=HEAD)
    c.border = BORDER
q.column_dimensions["A"].width = 10
q.column_dimensions["B"].width = 100
for i, (a, b) in enumerate(qs, 5):
    for j, v in enumerate((a, b), 1):
        c = q.cell(row=i, column=j, value=v)
        c.font = f(10)
        c.border = BORDER
        c.alignment = Alignment(wrap_text=True, vertical="top")

# ---------------- リスク
rk = wb.create_sheet("リスク")
rk.sheet_view.showGridLines = False
rk["A1"] = "日程のリスクと対応"
rk["A1"].font = f(14, True, DK)
risks = [
    ("ヒアリングの日程が12月に集中し、年末で相手がつかまらない", "高", "A2で11月中に日程を押さえる。聞けなかったSTEPは v1.2 では変えず、パイロットで補う"),
    ("発明塾資料の利用条件の確認（U1）が長引く", "中", "回答が出るまで、発明塾の考え方の説明はエージェントのナレッジから外して配布する"),
    ("Web検索の外部送信（U12）が認められない", "中", "Web調査を外し、案内役のみで配布。調べものは利用者が自分で行う"),
    ("子エージェントの連携がうまく動かない", "中", "子エージェント設定ガイドで切り分け。期限までに直らなければ案内役のみでβ版"),
    ("パイロット案件が見つからない・途中で止まる", "中", "過去案件をワークシートで書き直す形の試行に切り替える"),
    ("IPL担当の負担が大きく、相談ログ・確認ポイントが回らない", "中", "相談ログの項目を減らす。確認ポイントは案件の大きさで濃さを変える"),
    ("手順書の改訂とエージェントの更新がずれる", "低", "改訂は四半期ごとにまとめ、更新のたびにテストセットを流す"),
]
for j, h in enumerate(["リスク", "影響", "対応"], 1):
    c = rk.cell(row=3, column=j, value=h)
    c.font = f(10, True)
    c.fill = PatternFill("solid", fgColor=HEAD)
    c.border = BORDER
rk.column_dimensions["A"].width = 56
rk.column_dimensions["B"].width = 8
rk.column_dimensions["C"].width = 70
for i, row in enumerate(risks, 4):
    for j, v in enumerate(row, 1):
        c = rk.cell(row=i, column=j, value=v)
        c.font = f(10)
        c.border = BORDER
        c.alignment = Alignment(wrap_text=True, vertical="top")
    rk.row_dimensions[i].height = 34

for s in wb.worksheets[1:]:
    s.page_setup.orientation = "landscape"
    s.page_setup.fitToWidth = 1
    s.page_setup.fitToHeight = 0
    s.sheet_properties.pageSetUpPr.fitToPage = True
wb.save(sys.argv[1])
print("saved", sys.argv[1])
