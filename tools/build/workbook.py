import sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

FONT = "Meiryo"
DK, TEAL, LT, HEAD = "1F3A4D", "0F7C8C", "EEF3F5", "DCE9ED"
thin = Side(style="thin", color="A9C7CF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
WRAP = Alignment(wrap_text=True, vertical="top")

wb = Workbook()
wb.remove(wb.active)
index = []


def f(size=10, bold=False, color="1E2B33", italic=False):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic)


def sheet(name, step, purpose, widths):
    ws = wb.create_sheet(name)
    ws.sheet_view.showGridLines = False
    ws["A1"] = name.split("_", 1)[1] if "_" in name else name
    ws["A1"].font = f(16, True, DK)
    ws["A2"] = f"{step}｜{purpose}"
    ws["A2"].font = f(10, False, "5B7083", True)
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    index.append((name, step, purpose))
    return ws


def section(ws, row, text):
    c = ws.cell(row=row, column=1, value=text)
    c.font = f(12, True, TEAL)
    return row + 1


def form(ws, row, items, value_cols=3, height=36):
    """2-column form: label in A, merged input B..(B+value_cols-1)."""
    for item in items:
        label, hint = (item if isinstance(item, tuple) else (item, ""))
        a = ws.cell(row=row, column=1, value=label)
        a.font = f(10, True)
        a.fill = PatternFill("solid", fgColor=LT)
        a.alignment = WRAP
        a.border = BORDER
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=1 + value_cols)
        b = ws.cell(row=row, column=2, value=hint or None)
        b.font = f(10, False, "8A99A6" if hint else "1E2B33", bool(hint))
        b.alignment = WRAP
        for col in range(2, 2 + value_cols):
            ws.cell(row=row, column=col).border = BORDER
        ws.row_dimensions[row].height = height
        row += 1
    return row + 1


def table(ws, row, headers, nrows=30, prefill=None, dv=None, height=30, filt=False):
    for j, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=j, value=h)
        c.font = f(10, True)
        c.fill = PatternFill("solid", fgColor=HEAD)
        c.alignment = Alignment(wrap_text=True, vertical="center")
        c.border = BORDER
    ws.row_dimensions[row].height = 32
    hdr = row
    for i in range(nrows):
        r = row + 1 + i
        for j in range(1, len(headers) + 1):
            c = ws.cell(row=r, column=j)
            c.border = BORDER
            c.alignment = WRAP
            c.font = f(10)
        if prefill and i < len(prefill):
            for j, v in enumerate(prefill[i], 1):
                if v is not None:
                    ws.cell(row=r, column=j, value=v)
        ws.row_dimensions[r].height = height
    last = row + nrows
    for col, options in (dv or {}).items():
        v = DataValidation(type="list", formula1='"' + ",".join(options) + '"', allow_blank=True)
        ws.add_data_validation(v)
        letter = get_column_letter(col)
        v.add(f"{letter}{hdr + 1}:{letter}{last}")
    if filt:
        ws.auto_filter.ref = f"A{hdr}:{get_column_letter(len(headers))}{last}"
    return last + 2, hdr


RANK = ["A", "B", "C", "D"]
STATE = ["事実", "予想", "未確認"]

# ---------------- 使い方
ws0 = wb.create_sheet("使い方", 0)
ws0.sheet_view.showGridLines = False
ws0["A1"] = "IPL標準フロー ワークシート集 v1.2（たたき台）"
ws0["A1"].font = f(16, True, DK)
notes = [
    "案件ごとにこのファイルをコピーして使います（ファイル名例：IPL_案件名_YYYYMMDD.xlsx）。",
    "手順は「IPL標準フロー 手順書（Word）」の各STEPの章を見てください。シート名の番号は手順書のひな形番号と同じです。",
    "灰色の斜体は記入例・ヒントです。上書きしてください。▼の付いた欄はリストから選びます。",
    "事実には出典（URL・特許番号など）を必ず書きます。出典を示せないことは「未確認」とします。",
    "出願前の発明・社外秘の情報を書く場合は、ファイルの保存場所とアクセス権に注意してください。",
]
for i, t in enumerate(notes, 3):
    ws0.cell(row=i, column=1, value="・" + t).font = f(10)
ws0.column_dimensions["A"].width = 30
ws0.column_dimensions["B"].width = 26
ws0.column_dimensions["C"].width = 70

# ---------------- 00 案件ファイル
ws = sheet("00_案件ファイル", "全STEP", "案件の進み具合を1か所で見る", [26, 34, 26, 30, 30])
r = section(ws, 4, "基本情報")
r = form(ws, r, ["案件名", "進め方（小・中・大）", "目的の種類と、軽くする・飛ばすSTEP", "今のSTEP", "意思決定者", "依頼者", "担当（あなた）／伴走するIPL担当／反対の立場で見る人", "期限（意思決定日）", "AIに入力してよい範囲（社内規程の確認日・確認者）"], value_cols=4, height=24)
r = section(ws, r, "できあがったもの")
stepsrow = [
    ("0 お題設定", "お題シート、エッジ情報カード"), ("1 顧客価値仮説", "困りごとシート、価値の仮説"),
    ("2 事業仮説", "企画骨子、事業仮説シート"), ("3 調査目的設定", "調査目的シート（確認①）"), ("4 調査設計", "調査計画"),
    ("5 情報探索", "情報リスト"), ("6 構造化・分析", "整理表・マップ"),
    ("7 仮説検証・反証", "仮説の記録、ヒアリング票、分かったこと／まだ分からないこと（確認②）"),
    ("8 意思決定", "判断メモ、意思決定の記録、振り返りアンケート、ノウハウ集"),
]
r, _ = table(ws, r, ["STEP", "もの", "状態▼", "保存場所", "メモ"], nrows=9, prefill=[(a, b, "未") for a, b in stepsrow], dv={3: ["未", "途中", "済", "飛ばした"]}, height=22)
r = section(ws, r, "まだ分からないこと（判断への影響が大きい順）")
r, _ = table(ws, r, ["No", "まだ分からないこと", "判断への影響（大/中/小）▼", "確かめ方", "期限"], nrows=8, prefill=[(i + 1,) for i in range(8)], dv={3: ["大", "中", "小"]}, height=22)
r = section(ws, r, "新たなお題（STEP8で出たもの。次の周回でSTEP0から回す）")
r, _ = table(ws, r, ["No", "新たなお題", "出たきっかけ", "次に回す時期"], nrows=5, prefill=[(i + 1,) for i in range(5)], height=22)
r = section(ws, r, "チェックポイントの記録")
table(ws, r, ["日付", "STEP", "確認した人", "結果▼", "戻った先と理由"], nrows=15, dv={4: ["通過", "戻る"]}, height=22)

# ---------------- 01 お題シート
ws = sheet("01_お題シート", "STEP0 お題設定", "気になっていることを1文のお題にする（何を決めるかはSTEP3で決める）", [30, 30, 30, 30])
r = section(ws, 4, "お題シート")
form(ws, r, [
    "気になっていること（言い換えずに）",
    ("お題（1文）", "例：［技術・能力］は、［どんな人・場面］で役に立つのではないか。／［変化］は当社にとって機会（または脅威）ではないか。"),
    "背景（なぜ今気になるのか）", ("きっかけになった情報（エッジ情報カードのID）", "E1, E2…"),
    "依頼者（いれば）／その先で判断する人（分かる範囲で）",
    "出願前の発明・社外秘の情報が含まれるか／AIに入力してよい範囲",
    ("進め方の大きさ（仮）", "小／中／大"), "今の時点の思いつき・仮説（粗くてよい）", "一緒に考えたい人・話を聞きたい人",
], value_cols=3, height=40)

# ---------------- 02 エッジ情報カード
ws = sheet("02_エッジ情報カード", "STEP0 お題設定", "きっかけの情報を「事実→解釈→将来の予想」で書く（1〜3件）", [8, 24, 30, 30, 16, 12, 12, 9, 30, 30, 26, 26, 26])
r = section(ws, 4, "エッジ情報カード（1行1件）")
r, _ = table(ws, r, ["ID", "読む前に確かめたいこと", "事実（何が起きたか）", "出典（URL・特許番号・DOI）", "発行主体", "公開日", "取得日", "ランク▼", "解釈［予想］なぜ今・どの制約が変わるか", "将来の予想：誰に新しい困りごと・機会が生じるか", "誰がどこまで取り組み、どこで止まっているか", "外れている可能性（反証）", "次に調べること"], nrows=8, prefill=[(f"E{i+1}",) for i in range(8)], dv={8: RANK}, height=60)
r = section(ws, r, "変化と困りごとの表")
r, _ = table(ws, r, ["No", "変化（予想が変わる項目だけ）", "出典", "困る人（直接／その先）", "必要になる機能", "状態▼"], nrows=8, prefill=[(i + 1,) for i in range(8)], dv={6: STATE}, height=36)
r = section(ws, r, "今日の結論（毎日1行）")
table(ws, r, ["日付", "今日の結論", "次に調べること"], nrows=15, height=22)

# ---------------- 03 困りごと
ws = sheet("03_困りごとシート_ロジックツリー", "STEP1 顧客価値仮説", "ファーストユーザーと困りごとを、トレードオフまで掘る", [8, 12, 26, 26, 10, 10, 10, 10, 10, 10, 34])
r = section(ws, 4, "顧客候補（点数は会いに行く順番を決めるためだけに使う）")
pre = [(f"C{i+1}", None, None, None, None, None, None, None, None, f"=IF(COUNT(E{0}:I{0})=0,\"\",SUM(E{0}:I{0}))") for i in range(6)]
r0 = r
r, hdr = table(ws, r, ["候補", "直接／その先▼", "会社名・部署（または企業コード）", "使う人／進める人／決裁者", "困りごとの強さ▼", "頻度▼", "今の手段への不満▼", "予算・決裁権▼", "会いに行ける可能性▼", "合計", "会いたい理由（1〜3行）"], nrows=6, dv={2: ["直接", "その先"], 5: list("12345"), 6: list("12345"), 7: list("12345"), 8: list("12345"), 9: list("12345")}, height=36)
for i in range(6):
    rr = hdr + 1 + i
    ws.cell(row=rr, column=1, value=f"C{i+1}")
    ws.cell(row=rr, column=10, value=f'=IF(COUNT(E{rr}:I{rr})=0,"",SUM(E{rr}:I{rr}))')
r = section(ws, r, "ファーストユーザー（1つ）")
r = form(ws, r, ["会社名・人", "この相手に決めた理由（今分かる範囲で）"], value_cols=4)
r = section(ws, r, "今の状態とありたい状態")
r = form(ws, r, ["今の状態（As-Is）", "妨げているもの", "ありたい状態（To-Be）", "判断に使う指標（最大3つ）", "今は何で代用しているか", "その手段の費用・損失・不満"], value_cols=4)
r = section(ws, r, "課題‐解決ロジックツリー（左から右へ）")
r, _ = table(ws, r, ["No", "顧客の成果", "妨げる現象", "原因", "解決の原理（案）"], nrows=8, prefill=[(i + 1,) for i in range(8)], height=30)
r = section(ws, r, "上位の言葉への言い換え（用途を広げる）")
r, _ = table(ws, r, ["No", "元の言葉", "上位の言葉", "そこから出る別の対象・機能"], nrows=5, prefill=[(i + 1,) for i in range(5)], height=24)
r = section(ws, r, "トレードオフ")
r, _ = table(ws, r, ["No", "Aを良くすると", "Bが悪くなる", "既存の製品・特許はどちらを優先しているか（出典）"], nrows=4, prefill=[(i + 1,) for i in range(4)], height=30)
r = section(ws, r, "2つの理由を分けて書く")
form(ws, r, ["技術的に難しい理由", "顧客にとって重要な理由"], value_cols=4)

# ---------------- 04 価値仮説
ws = sheet("04_価値仮説_企画骨子", "STEP1・STEP2", "価値の仮説を1文にし、企画骨子5項目をまず書く", [30, 40, 40, 16])
r = section(ws, 4, "価値の仮説（STEP1の最後に1文で）")
r = form(ws, r, [("価値の仮説", "［顧客］は［場面］で［今の手段］のために［具体的な損失］を受けている。［指標］を［目標値］まで改善できれば、［お金を払う人］が［対価］を払う可能性がある。"), ("目標値は", "顧客の実測／こちらの予想"), "判断に使う指標（3つまで）"], value_cols=3, height=48)
r = section(ws, r, "企画骨子（STEP2の最初に。空欄があってもまず書く）")
r, _ = table(ws, r, ["項目", "記入", "メモ", "状態▼"], nrows=5, prefill=[("① 顧客価値の仮説",), ("② 最小限の構成（MVP）",), ("③ ファーストユーザー（会社名・人）",), ("④ 話を聞くための準備物",), ("⑤ なぜ自分たちがやるのか（背景・経緯）",)], dv={4: STATE}, height=48)

# ---------------- 05 事業仮説
ws = sheet("05_事業仮説シート", "STEP2 事業仮説", "いきなり○△×を付けない。崩れる条件と、一番あやしい仮説の印", [6, 34, 40, 12, 28, 34, 10])
r = section(ws, 4, "事業仮説シート")
pre = [("H1", "売れる（STEP1から）"), ("H2", "勝てる：なぜ競合では解けないか"), ("H3", "勝てる：なぜ他の人はまだやっていないか、自社なら越えられる理由"), ("H4", "勝てる（守れる）：この会社から買わなければならない理由（特許・ノウハウ・認証など）"), ("H5", "儲かる：1社分の売上（使用量×採用率×単価）、広がり方"), ("H6", "できそう（図の名前にはないが必ず確かめる）：性能・量産・品質・規制・社内外の資源"), ("H7",), ("H8",)]
r, _ = table(ws, r, ["ID", "観点", "仮説", "状態▼", "根拠（Source ID・URL）", "崩れる条件", "あやしい印▼"], nrows=8, prefill=pre, dv={4: STATE, 7: ["★"]}, height=48)
ws.cell(row=r, column=1, value="※出願の余地とFTO（他社特許を侵害せずに実施できるか）は別の問題。FTOはこのフローでは判断しない。").font = f(9, False, "5B7083", True)
r = section(ws, r + 2, "IPL担当の確認【確認ポイント①】")
form(ws, r, ["確認日", "確認者", "指摘"], value_cols=5, height=24)

# ---------------- 06 調査目的シート
ws = sheet("06_調査目的シート", "STEP3 調査目的設定", "何を決めるために、どの仮説を確かめるか。Goとやめる条件", [30, 30, 30, 30])
r = section(ws, 4, "調査目的シート")
form(ws, r, [
    "お題（STEP0）", "価値の仮説（STEP1、1文）",
    ("今回決めたいこと（1文）", "当社は［技術・能力］を使って、［顧客］の［困りごと］に対し、［用途・事業の形］で取り組む価値があるかを、［期限］までに判断する。"),
    ("判断したあとにやること", "研究する／出願する／組む／買う／見送る／その他"), "会社の方針（中計・事業部方針）との整合",
    ("確かめる仮説と順番（事業仮説シートのID）", "1. H□（理由：　） 2. H□ 3. H□"), ("確かめない仮説と、その理由", "答えがどうであれ判断が変わらない、など"),
    "意思決定者", "意思決定日・途中の確認日", "対象の市場・地域・期間", "使えるデータ（社内で契約しているDB・社内データ）", "投資の上限・技術の制約",
    ("Goにする条件（1文）", "例：最初の顧客候補1社が有償評価に進む意思を示し、主要な他社特許の壁を避けられる見込みがある"),
    ("やめる条件（1文）", "例：3社に聞いて、今の手段で困っていないという答えが続く"),
    ("進め方の大きさ（見直し後）", "小／中／大"), "軽くする・飛ばすSTEP", "最後に渡すもの",
    "確認：意思決定者（氏名・日付）", "IPL担当の確認【確認ポイント①】（氏名・日付・指摘）",
], value_cols=3, height=40)

# ---------------- 06 調査計画
ws = sheet("07_調査計画", "STEP4 調査設計", "仮説 → 必要な情報 → 場所 → 検索式 の順で決める", [8, 30, 8, 30, 22, 14, 34, 30, 24, 10])
r = section(ws, 4, "調査計画")
r, _ = table(ws, r, ["仮説ID", "仮説", "優先度▼", "確かめるのに必要な情報", "調べる場所", "検索の層▼", "検索式ID・検索式（案）", "外れている証拠の探し方", "やめる条件", "時間の上限"], nrows=12, dv={3: ["1", "2", "3"], 6: ["A課題", "B技術", "C会社", "A×B"]}, height=48)
r = section(ws, r, "検索式の記録")
table(ws, r, ["検索式ID", "データベース", "検索式", "検索日", "件数（公報）", "件数（ファミリー）", "代表的な特許の拾い漏れ", "関係ないものの割合（n=20）", "確認したIPL担当"], nrows=12, prefill=[(f"Q{i+1}",) for i in range(12)], height=36)

# ---------------- 07 情報リスト
cols = ["Source ID", "種類▼", "文献番号・URL", "優先日", "出願日", "公開日（または発行日）", "取得日", "検索式ID", "会社名（原文）", "まとめた会社名", "ファミリー番号", "権利の状態▼", "重要さの目安", "発明者・著者", "共同出願先・共同研究先", "技術", "機能", "課題", "解決手段", "効果", "用途", "お客様（直接/その先）", "市場", "該当箇所（ページ・段落）", "事実か予想か▼", "確かさのランク▼", "書いた人▼", "関係する仮説ID（+支える/−外れる）", "メモ"]
ws = sheet("08_情報リスト", "STEP5 情報探索", "1行1件。出典のない行は作らない", [10, 10, 30, 11, 11, 13, 11, 9, 18, 18, 14, 11, 12, 18, 20, 18, 18, 24, 24, 20, 18, 12, 14, 14, 10, 10, 12, 18, 24])
r, hdr = table(ws, 4, cols, nrows=200, prefill=[(f"S{i+1:03d}",) for i in range(200)], dv={2: ["特許", "論文", "公的資料", "決算資料", "会社の発表", "ニュース", "社内情報", "ヒアリング"], 12: ["審査中", "登録", "消滅", "取下・放棄", "不明"], 25: STATE, 26: RANK, 27: ["AI", "人", "AI＋人確認"]}, height=20, filt=True)
ws.freeze_panes = "B5"

# ---------------- 08 仮説の記録
ws = sheet("09_仮説の記録", "STEP7 仮説検証・反証", "残す・直す・捨てる・保留と、その理由（いちばん大事な欄）", [12, 8, 34, 14, 14, 12, 10, 28, 40])
r = section(ws, 4, "仮説の記録")
r, _ = table(ws, r, ["日付", "仮説ID", "仮説（変更後）", "支える情報（ID）", "外れている証拠（ID）", "判断▼", "確かさ▼", "次に調べること", "なぜそうしたか・なぜ次にそれを調べるか"], nrows=25, dv={6: ["残す", "直す", "捨てる", "保留"], 7: ["高", "中", "低"]}, height=40)
r = section(ws, r, "分かったこと・まだ分からないこと・新しい疑問（7-8）")
r, _ = table(ws, r, ["仮説ID", "分かったこと（結論と根拠ID）", "まだ分からないこと（理由と確かめ方）", "新しく出てきた疑問（次のお題の候補）"], nrows=8, height=40)
r = section(ws, r, "反対の立場からのレビュー")
r, _ = table(ws, r, ["日付", "レビューした人（またはAI）", "指摘", "確かめた結果", "仮説への反映"], nrows=10, height=36)
r = section(ws, r, "よくある思い込みの確認")
table(ws, r, ["確認項目", "確認▼", "メモ"], nrows=6, prefill=[("特許が見つからない理由（秘密・未公開・検索漏れ）を確かめた",), ("出願件数だけで強さを判断していない",), ("直近1年半の出願減少を未公開の影響と区別した",), ("別の技術で同じ困りごとが解決されていないか確かめた",), ("直接のお客様とその先のお客様の両方を確かめた",), ("判断の条件を変えた場合、理由を書いた",)], dv={2: ["済", "未", "該当なし"]}, height=24)

# ---------------- 09 ヒアリング票
ws = sheet("10_ヒアリング票", "STEP7 仮説検証・反証", "聞いた事実と自分の解釈を分ける。1回につき1シートをコピーして使う", [34, 44, 44])
r = section(ws, 4, "基本情報")
r = form(ws, r, ["日付／形式／記録者", "相手（会社・部署・役割。実名は別管理なら企業コード）", ("秘密保持契約・記録の扱い", "あり／なし、社外秘／社内のみ"), "持ち込んだ提案書（ファイル名・版）", "確かめたい仮説ID"], value_cols=2, height=24)
r, _ = table(ws, r, ["質問", "相手から聞いた事実", "自分の解釈（別欄）"], nrows=10, prefill=[("① 最近いつ、どんな場面で困ったか",), ("② 今はどう対応しているか（代わりの手段）",), ("③ その手段の費用・損失",), ("④ 新しい手段を入れない理由・妨げ",), ("⑤ 理想の仕様・指標",), ("⑥ お金を払う人・予算・決裁の流れ",), ("⑦ 価格の感覚",), ("⑧ 有償評価など次の行動に進めるか",), ("「違う」と言われた点・理解されなかった点",), ("紹介された次の相手",)], height=48)
r = form(ws, r, [("証拠の強さ", "強（発注・有償評価・支払条件）／中（予算と導入手順を確認）／弱（好意的な回答のみ）"), "技術評価の担当と購買承認の担当は同じか"], value_cols=2, height=30)

# ---------------- 10 検証実験カード
ws = sheet("11_検証実験カード", "STEP7・STEP8", "一番大きな不確かさを消す、一番小さな実験", [34, 60])
r = section(ws, 4, "検証実験カード")
form(ws, r, ["仮説ID・仮説", "今ある証拠（ID・ランク）", "最も困る反証（これが出たら仮説が崩れる）", "最小の実験（ヒアリング・サンプル評価・試作・机上試算など）", "成功の基準", "予算", "担当", "期限", "結果", ("判断", "Go／No Go（やめる・方向を変えて新たなお題へ）／条件付きGo")], value_cols=1, height=36)

# ---------------- 11 判断メモ
ws = sheet("12_判断メモ", "STEP8 依頼者が意思決定", "1枚目に結論・次の行動・やめる条件・まだ分からないこと", [34, 30, 30, 30])
r = section(ws, 4, "1枚目（意思決定者向け）")
r = form(ws, r, ["つまり何か（結論）", ("提案", "Go／No Go（やめる・方向を変えて新たなお題へ）／条件付きGo（追加の検証）"), "次に何をするか（誰が・いつまでに・費用）", "やめる条件", "まだ分からないこと（ランクC・D）", "判断の根拠", "残る一番大きな不確かさ", "それを消す一番小さな実験", "次に判断する日", "確かめずに判断した部分（期限・予算で打ち切った部分）", "新たなお題（次にSTEP0から回すもの）"], value_cols=3, height=40)
r = section(ws, r, "4観点の評価（合計点でGoにしない）")
r, _ = table(ws, r, ["観点", "評価▼", "主な根拠ID（ランクA・B）", "致命的な問題の有無▼"], nrows=4, prefill=[("売れる",), ("勝てる（守れるを含む）",), ("儲かる",), ("できそう",)], dv={2: ["①確かめた", "②有力だが未確認", "③根拠不足", "④外れている証拠が優勢"], 4: ["なし", "あり"]}, height=30)
r = section(ws, r, "儲かるの試算（最初の顧客から積み上げる）")
r, hdr = table(ws, r, ["項目", "悲観", "標準", "楽観"], nrows=6, prefill=[("1社の年間使用量",), ("採用率",), ("自社の単価",), ("1社分の売上（自動計算）",), ("対象社数",), ("売上の見込み（自動計算）",)], height=22)
for col in "BCD":
    ws[f"{col}{hdr+4}"] = f'=IF(COUNT({col}{hdr+1}:{col}{hdr+3})<3,"",{col}{hdr+1}*{col}{hdr+2}*{col}{hdr+3})'
    ws[f"{col}{hdr+6}"] = f'=IF(OR({col}{hdr+4}="",{col}{hdr+5}=""),"",{col}{hdr+4}*{col}{hdr+5})'
r = section(ws, r, "意思決定の記録")
form(ws, r, ["日付／意思決定者", "判断", "出た質問・指摘"], value_cols=3, height=30)

# ---------------- 12 振り返り
ws = sheet("13_振り返りアンケート", "STEP8（最後）振り返り", "標準化チームが集計し、手順書とエージェントの改訂に使う（15分）", [20, 18, 14, 14, 14, 50])
r = section(ws, 4, "基本情報")
r = form(ws, r, ["案件名（または案件ID）", "あなたの部署", "進め方（小・中・大）／期間"], value_cols=5, height=22)
r = section(ws, r, "STEPごと")
names = ["0 お題設定", "1 顧客価値仮説", "2 事業仮説", "3 調査目的設定", "4 調査設計", "5 情報探索", "6 構造化・分析", "7 仮説検証・反証", "8 意思決定（振り返りを含む）"]
r, _ = table(ws, r, ["STEP", "実施▼", "かかった時間", "分かりやすさ▼", "役に立ったか▼", "手順と違うやり方をしたところ・理由"], nrows=9, prefill=[(n,) for n in names], dv={2: ["実施した", "軽くした", "飛ばした"], 4: list("12345"), 5: list("12345")}, height=26)
r = section(ws, r, "全体")
form(ws, r, ["戻った回数（STEP4へ／STEP3へ／STEP1へ／STEP2へ）ときっかけ。新たなお題として回したか", "手順書に書いてなくて困ったこと", "手順書のとおりにやって、かえって遠回りになったこと", "エージェントが役に立った場面／間違えた場面", "IPL担当に相談した内容（手順書に足すべきもの）", "判断の結果と、その判断に一番効いた情報"], value_cols=5, height=40)

# ---------------- 13 ノウハウ集
ws = sheet("14_ノウハウ集", "STEP8（最後）振り返り", "次の人が同じ手順をやり直せるように残す", [36, 70])
r = section(ws, 4, "ノウハウ集")
form(ws, r, ["案件名／期間／担当", "使った検索式と、その検索式にした理由", "分類の切り口と名寄せの対応表（保存場所）", "採用した仮説・捨てた仮説と、その理由", "「なぜ次にその調べものをしたか」の要点", "ヒアリングで分かった、公開情報からは読めなかったこと", "AIが間違えた点と、人が直した理由", "エージェントの指示文の改善案", "標準フローへの改善提案", "振り返りの日（判断の1〜2年後）", ("社外に出してよいか", "社外可／社内のみ／社外秘")], value_cols=1, height=40)

# ---------------- 14 フロー検証
ws = sheet("15_フロー検証チェックリスト", "標準化チーム", "フローの各STEPが正しく設計されているかを、改訂のたびに確かめる", [6, 50, 7, 7, 7, 7, 7, 7, 7, 7, 7])
r = section(ws, 4, "A. 構造のチェック（○／△／×）")
items = ["目的が1文で書かれている", "始める条件（前のSTEPから受け取るもの）が書かれている", "できあがるもの（ひな形）が1つ以上ある", "できあがるものが、次のSTEPの始める条件になっている", "チェックする人と、だめなときの戻り先が書かれている", "手順に目安時間がある", "エージェントに頼めること／自分でやることが分かれている", "事実・予想・未確認を分けるルールが効いている", "反証を探す手順がある（フロー全体で）"]
r, _ = table(ws, r, ["No", "確認項目"] + [f"STEP{i}" for i in range(9)], nrows=9, prefill=[(f"A{i+1}", t) for i, t in enumerate(items)], dv={c: ["○", "△", "×"] for c in range(3, 12)}, height=24)
r = section(ws, r, "B. 外部の基準との照合")
r, _ = table(ws, r, ["No", "基準", "照合した版・ページ", "対応するSTEP", "抜けている要素", "対応方針"], nrows=4, prefill=[(1, "特許庁「IPランドスケープ実践ガイドブック」"), (2, "特許庁「経営戦略に資する知財情報分析・活用に関する調査研究」"), (3, "発明塾の手法（手順書 付録）"), (4, "社内の既存の取り組み（R&Dプラットフォームなど）")], height=36)
r = section(ws, r, "C. 実際の案件での確認（過去案件の当てはめ・パイロット）")
r, _ = table(ws, r, ["No", "案件", "当てはめ方▼", "うまく当てはまったSTEP", "当てはまらなかったSTEP・理由", "手順書の修正案"], nrows=6, prefill=[(i + 1,) for i in range(6)], dv={3: ["過去案件", "パイロット"]}, height=36)
r = section(ws, r, "D. 判定")
form(ws, r, ["修正が必要なSTEP", "次の版で直すこと", "判定者・日付"], value_cols=5, height=30)

# index on 使い方
ws0.cell(row=9, column=1, value="シート").font = f(10, True)
ws0.cell(row=9, column=2, value="STEP").font = f(10, True)
ws0.cell(row=9, column=3, value="使いどころ").font = f(10, True)
for c in range(1, 4):
    ws0.cell(row=9, column=c).fill = PatternFill("solid", fgColor=HEAD)
    ws0.cell(row=9, column=c).border = BORDER
for i, (name, step, purpose) in enumerate(index, 10):
    c = ws0.cell(row=i, column=1, value=name)
    c.hyperlink = f"#'{name}'!A1"
    c.font = Font(name=FONT, size=10, color=TEAL, underline="single")
    ws0.cell(row=i, column=2, value=step).font = f(10)
    ws0.cell(row=i, column=3, value=purpose).font = f(10)
    for col in range(1, 4):
        ws0.cell(row=i, column=col).border = BORDER

for w in wb.worksheets:
    w.page_setup.orientation = "landscape"
    w.page_setup.fitToWidth = 1
    w.page_setup.fitToHeight = 0
    w.sheet_properties.pageSetUpPr.fitToPage = True

wb.save(sys.argv[1])
print("saved", sys.argv[1], len(wb.worksheets), "sheets")
