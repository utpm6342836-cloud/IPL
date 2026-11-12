"""Draw the IPL flowchart onto slide 4 of the deck with glued connectors."""
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.dml import MSO_LINE
from pptx.oxml.ns import qn
from lxml import etree

src, dst, slide_no = sys.argv[1], sys.argv[2], int(sys.argv[3])
prs = Presentation(src)
slide = prs.slides[slide_no - 1]
sh = slide.shapes

DK = RGBColor(0x1F, 0x3A, 0x4D)
TEAL = RGBColor(0x0F, 0x7C, 0x8C)
AMBER = RGBColor(0xD9, 0x86, 0x2B)
GRAY = RGBColor(0x8A, 0x99, 0xA6)
LT = RGBColor(0xEE, 0xF3, 0xF5)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
TXT = RGBColor(0x1E, 0x2B, 0x33)
FONT = "Meiryo"


def text(shape, lines, size=11, color=TXT, bold_first=True):
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.04)
    tf.margin_top = tf.margin_bottom = Inches(0.02)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = ln
        r.font.size = Pt(size if i else size + 0.5)
        r.font.bold = bold_first and i == 0
        r.font.color.rgb = color
        r.font.name = FONT
        rpr = r._r.get_or_add_rPr()
        ea = rpr.find(qn("a:ea"))
        if ea is None:
            ea = etree.SubElement(rpr, qn("a:ea"))
        ea.set("typeface", FONT)


def box(name, kind, x, y, w, h, lines, fill=LT, line=TEAL, fcolor=TXT, dash=False, size=11):
    s = sh.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    s.name = name
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    s.line.color.rgb = line
    s.line.width = Pt(1.5)
    if dash:
        s.line.dash_style = MSO_LINE.DASH
    s.shadow.inherit = False
    text(s, lines, size=size, color=fcolor)
    return s


def label(name, x, y, w, h, t, size=10, color=TXT, align=PP_ALIGN.LEFT, bold=False):
    tb = sh.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tb.name = name
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = t
    r.font.size = Pt(size)
    r.font.color.rgb = color
    r.font.bold = bold
    r.font.name = FONT
    rpr = r._r.get_or_add_rPr()
    etree.SubElement(rpr, qn("a:ea")).set("typeface", FONT)
    return tb


def arrowhead(shape):
    ln = shape.line._get_or_add_ln()
    for tag in ("a:headEnd", "a:tailEnd"):
        old = ln.find(qn(tag))
        if old is not None:
            ln.remove(old)
    etree.SubElement(ln, qn("a:tailEnd")).set("type", "triangle")


def connect(name, a, a_site, b, b_site, color=DK, dash=False, kind=MSO_CONNECTOR.STRAIGHT):
    c = sh.add_connector(kind, 0, 0, Inches(1), Inches(1))
    c.begin_connect(a, a_site)
    c.end_connect(b, b_site)
    c.name = name
    c.line.color.rgb = color
    c.line.width = Pt(1.75)
    if dash:
        c.line.dash_style = MSO_LINE.DASH
    arrowhead(c)
    return c


def poly(name, pts, color=AMBER, dash=True, arrow=True, width=1.5):
    fb = sh.build_freeform(Inches(pts[0][0]), Inches(pts[0][1]), scale=1.0)
    fb.add_line_segments([(Inches(x), Inches(y)) for x, y in pts[1:]], close=False)
    s = fb.convert_to_shape()
    s.name = name
    s.fill.background()
    s.line.color.rgb = color
    s.line.width = Pt(width)
    if dash:
        s.line.dash_style = MSO_LINE.DASH
    s.shadow.inherit = False
    if arrow:
        arrowhead(s)
    return s


# ---- geometry (inches)
ROW_Y, ROW_H = 3.35, 1.0
SW, GAP = 1.12, 0.27
xs = {}
x = 0.4
start = None

# loop band first (behind)
band = sh.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.62), Inches(2.95), Inches(7.43), Inches(1.75))
band.name = "検証ループ枠"
band.adjustments[0] = 0.08
band.fill.solid(); band.fill.fore_color.rgb = RGBColor(0xF6, 0xF9, 0xFA)
band.line.color.rgb = RGBColor(0xA9, 0xC7, 0xCF); band.line.width = Pt(1); band.line.dash_style = MSO_LINE.ROUND_DOT
band.shadow.inherit = False
label("検証ループ見出し", 5.75, 2.99, 3.5, 0.28, "検証ループ（STEP3〜STEP6）", size=10, color=TEAL, bold=True)

start = box("開始", MSO_SHAPE.ROUNDED_RECTANGLE, 0.4, ROW_Y + 0.15, 0.95, 0.7, ["個別案件", "の依頼"], fill=DK, line=DK, fcolor=WHITE, size=10)
start.adjustments[0] = 0.5
names = [
    ("STEP0", ["目的設定"]),
    ("STEP1", ["顧客価値仮説", "（売れる）"]),
    ("STEP2", ["事業仮説", "（勝てる・儲かる", "・できそう）"]),  # small font
    ("STEP3", ["調査設計"]),
    ("STEP4", ["情報探索"]),
    ("STEP5", ["構造化・分析"]),
    ("STEP6", ["仮説検証", "・反証"]),
]
boxes = []
x = 1.62
for nm, sub in names:
    b = box(nm, MSO_SHAPE.ROUNDED_RECTANGLE, x, ROW_Y, SW, ROW_H, [nm] + sub, size=(8.5 if nm == "STEP2" else 10))
    b.adjustments[0] = 0.12
    boxes.append(b)
    xs[nm] = x + SW / 2
    x += SW + GAP

D = box("判断", MSO_SHAPE.DIAMOND, 11.4, ROW_Y - 0.15, 1.5, 1.3, ["判断に足る", "Factは", "そろった？"], fill=WHITE, line=DK, size=9)
DX = 11.4 + 0.75

# forward connectors (glued). sites: 0 top, 1 left, 2 bottom, 3 right
connect("進む_開始-STEP0", start, 3, boxes[0], 1)
for i in range(len(boxes) - 1):
    connect(f"進む_{names[i][0]}-{names[i+1][0]}", boxes[i], 3, boxes[i + 1], 1)
connect("進む_STEP6-判断", boxes[-1], 3, D, 1)

# lower row
S7 = box("STEP7", MSO_SHAPE.ROUNDED_RECTANGLE, 11.4, 5.1, 1.5, 0.95, ["STEP7", "依頼者が意思決定", "（Go / No Go）"], fill=LT, size=10)
S7.adjustments[0] = 0.12
S8 = box("STEP8", MSO_SHAPE.ROUNDED_RECTANGLE, 8.8, 5.1, 2.0, 0.95, ["STEP8", "学習・標準化", "（振り返り・ノウハウ集）"], fill=LT, size=10)
S8.adjustments[0] = 0.12
END = box("横展開", MSO_SHAPE.ROUNDED_RECTANGLE, 6.45, 5.22, 1.75, 0.7, ["次の案件へ", "横展開"], fill=DK, line=DK, fcolor=WHITE, size=10)
END.adjustments[0] = 0.5
MGT = box("経営層へ提言", MSO_SHAPE.ROUNDED_RECTANGLE, 11.4, 6.3, 1.5, 0.6, ["経営層へ提言", "（必要に応じて）"], fill=WHITE, line=GRAY, fcolor=TXT, dash=True, size=9)

connect("進む_判断-STEP7", D, 2, S7, 0)
connect("進む_STEP7-STEP8", S7, 1, S8, 3)
connect("進む_STEP8-横展開", S8, 1, END, 3)
connect("任意_STEP7-経営層", S7, 2, MGT, 0, color=GRAY, dash=True)
label("Yesラベル", 12.22, 4.6, 0.6, 0.3, "Yes", size=11, color=DK, bold=True)

# feedback trunk + branches (free lines)
TOP_Y = ROW_Y - 0.15  # diamond top
poly("戻る_幹線", [(DX, TOP_Y), (DX, 1.55)], arrow=False)
label("Noラベル", DX + 0.08, 2.75, 0.5, 0.3, "No", size=11, color=AMBER, bold=True)
branches = [
    (2.75, "STEP3", "Factが不足 → STEP3で調査設計を見直す", AMBER, True),
    (2.35, "STEP2", "事業仮説が弱い → STEP2（提案：共有済みの図にはない）", GRAY, True),
    (1.95, "STEP1", "顧客価値が弱い → STEP1へ", AMBER, True),
    (1.55, "STEP0", "目的がずれた → STEP0へ（依頼者と相談し直す）", AMBER, True),
]
for y, tgt, t, col, dash in branches:
    poly(f"戻る_{tgt}", [(DX, y), (xs[tgt], y), (xs[tgt], ROW_Y)], color=col, dash=dash)
    label(f"戻るラベル_{tgt}", xs[tgt] + 0.12, y - 0.3, 6.5, 0.28, t, size=10, color=(GRAY if col == GRAY else AMBER), bold=(col != GRAY))

# back-to-start (lateral deployment)
poly("戻る_横展開-開始", [(6.45, 5.57), (0.875, 5.57), (0.875, ROW_Y + 0.85)], color=DK, dash=True)
label("横展開ラベル", 1.05, 5.63, 5.0, 0.28, "知見を次の案件・IPL活動に生かす", size=10, color=DK)

# notes and legend
label("注記1", 1.05, 4.5, 5.0, 0.3, "※目的に応じてSTEPを軽く・飛ばす（STEP0・7は除く）", size=10)
label("注記2", 1.05, 4.78, 5.0, 0.3, "※STEP1・2は既存情報（社内含む）で立てる", size=10)
label("注記3", 1.05, 5.06, 5.0, 0.3, "※STEP2後・STEP6後にIPL担当が確認（確認ポイント①②）", size=10)

ly = 6.55
poly("凡例_進む", [(0.4, ly + 0.12), (1.1, ly + 0.12)], color=DK, dash=False, width=1.75)
label("凡例_進む_t", 1.2, ly, 1.0, 0.28, "進む", size=10)
poly("凡例_戻る", [(2.1, ly + 0.12), (2.8, ly + 0.12)], color=AMBER, dash=True)
label("凡例_戻る_t", 2.9, ly, 2.2, 0.28, "戻る（フィードバック）", size=10)
poly("凡例_提案", [(5.1, ly + 0.12), (5.8, ly + 0.12)], color=GRAY, dash=True)
label("凡例_提案_t", 5.9, ly, 3.6, 0.28, "提案・任意（共有済みの図にはない線を含む）", size=10)

prs.save(dst)
print("saved", dst)
