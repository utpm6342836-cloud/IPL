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


# ---- geometry (inches): row1 STEP0->3 (left to right), row2 STEP4->7 (right to left), row3 STEP8
COLS = [0.85, 3.95, 7.05, 10.15]
BW, BH = 2.4, 0.9
R1, R2, R3 = 1.5, 3.6, 5.55


def cx(i):
    return COLS[i] + BW / 2


band = sh.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.7), Inches(3.38), Inches(12.05), Inches(1.82))
band.name = "検証ループ枠"
band.adjustments[0] = 0.08
band.fill.solid(); band.fill.fore_color.rgb = RGBColor(0xF6, 0xF9, 0xFA)
band.line.color.rgb = RGBColor(0xA9, 0xC7, 0xCF); band.line.width = Pt(1); band.line.dash_style = MSO_LINE.ROUND_DOT
band.shadow.inherit = False
label("検証ループ見出し", 9.4, 4.93, 3.2, 0.24, "検証ループ（STEP4〜STEP7）", size=10, color=TEAL, bold=True, align=PP_ALIGN.RIGHT)

def step(name, col, y, lines, size=11):
    b = box(name, MSO_SHAPE.ROUNDED_RECTANGLE, COLS[col], y, BW, BH, lines, size=size)
    b.adjustments[0] = 0.15
    return b

S0 = step("STEP0", 0, R1, ["STEP0", "お題設定"])
S1 = step("STEP1", 1, R1, ["STEP1", "顧客価値仮説", "（売れる）"])
S2 = step("STEP2", 2, R1, ["STEP2", "事業仮説", "（勝てる・儲かる）"])
S3 = step("STEP3", 3, R1, ["STEP3", "調査目的設定"])
S4 = step("STEP4", 3, R2, ["STEP4", "調査設計"])
S5 = step("STEP5", 2, R2, ["STEP5", "情報探索"])
S6 = step("STEP6", 1, R2, ["STEP6", "構造化・分析"])
S7 = step("STEP7", 0, R2, ["STEP7", "仮説検証・反証"])
S8 = step("STEP8", 0, R3, ["STEP8", "依頼者が意思決定", "（Go / No Go）"])
MGT = box("経営層へ提言", MSO_SHAPE.ROUNDED_RECTANGLE, COLS[1], R3 + 0.1, BW, 0.7, ["経営層へ提言", "（必要に応じて）"], fill=WHITE, line=GRAY, fcolor=TXT, dash=True, size=10)

# forward (glued). sites: 0 top, 1 left, 2 bottom, 3 right
connect("進む_STEP0-STEP1", S0, 3, S1, 1)
connect("進む_STEP1-STEP2", S1, 3, S2, 1)
connect("進む_STEP2-STEP3", S2, 3, S3, 1)
connect("進む_STEP3-STEP4", S3, 2, S4, 0)
connect("進む_STEP4-STEP5", S4, 1, S5, 3)
connect("進む_STEP5-STEP6", S5, 1, S6, 3)
connect("進む_STEP6-STEP7", S6, 1, S7, 3)
connect("進む_STEP7-STEP8", S7, 2, S8, 0)
connect("任意_STEP8-経営層", S8, 3, MGT, 1, color=GRAY, dash=True)
label("判断ラベル", cx(0) + 0.12, 5.24, 3.0, 0.26, "判断に足るFactがそろった", size=10, color=DK, bold=True)

# feedback from STEP7 top to STEP1 / STEP2 (proposal) / STEP3
top7 = R2
bot1 = R1 + BH
x7 = [COLS[0] + 0.35, COLS[0] + 1.0, COLS[0] + 1.65]
lanes = [
    (2.70, x7[0], COLS[1] + 0.5, AMBER, "顧客価値が弱い → STEP1へ", "STEP1"),
    (2.95, x7[1], COLS[2] + 0.5, GRAY, "事業仮説が弱い → STEP2へ（提案）", "STEP2"),
    (3.20, x7[2], COLS[3] + 0.5, AMBER, "調査目的とずれる → STEP3へ", "STEP3"),
]
for y, xs_, xt, col, t, tgt in lanes:
    poly(f"戻る_STEP7-{tgt}", [(xs_, top7), (xs_, y), (xt, y), (xt, bot1)], color=col)
    if tgt == "STEP3":
        label(f"戻るラベル_{tgt}", xt - 2.75, y - 0.27, 2.6, 0.25, t, size=10, color=AMBER, bold=True, align=PP_ALIGN.RIGHT)
    else:
        label(f"戻るラベル_{tgt}", xt + 0.12, y - 0.27, 3.4, 0.25, t, size=10, color=(GRAY if col == GRAY else AMBER), bold=(col != GRAY))

# Fact shortage: STEP7 bottom -> below row2 -> STEP4 bottom
fx = COLS[0] + 2.05
poly("戻る_STEP7-STEP4", [(fx, R2 + BH), (fx, 4.8), (cx(3), 4.8), (cx(3), R2 + BH)], color=AMBER)
label("戻るラベル_STEP4", 3.6, 4.85, 5.0, 0.26, "Factが不足 → STEP4で調査設計を見直す", size=10, color=AMBER, bold=True)

# new investigation: STEP8 left -> STEP0 left
poly("戻る_STEP8-STEP0", [(COLS[0], R3 + BH / 2), (0.5, R3 + BH / 2), (0.5, R1 + BH / 2), (COLS[0], R1 + BH / 2)], color=DK)
label("新たな調査ラベル", 0.14, 2.45, 0.3, 1.35, "新\nた\nな\n調\n査\nへ", size=10, color=DK, bold=True, align=PP_ALIGN.CENTER)

# notes and legend
label("注記1", 7.05, 5.55, 5.8, 0.26, "※STEP1・2は既存情報（社内含む）で立てる", size=10)
label("注記2", 7.05, 5.83, 5.8, 0.26, "※IPL担当の確認：①STEP3の後（調べ始める前）②STEP7の後（決める前）", size=10)
label("注記3", 7.05, 6.11, 5.8, 0.26, "※振り返り（学習・標準化）はSTEP8の最後に行う", size=10)
ly = 6.62
poly("凡例_進む", [(0.85, ly + 0.12), (1.55, ly + 0.12)], color=DK, dash=False, width=1.75)
label("凡例_進む_t", 1.65, ly, 1.0, 0.26, "プロセス", size=10)
poly("凡例_戻る", [(2.7, ly + 0.12), (3.4, ly + 0.12)], color=AMBER, dash=True)
label("凡例_戻る_t", 3.5, ly, 2.2, 0.26, "フィードバック", size=10)
poly("凡例_提案", [(5.0, ly + 0.12), (5.7, ly + 0.12)], color=GRAY, dash=True)
label("凡例_提案_t", 5.8, ly, 4.2, 0.26, "提案・任意（共有済みの図にはない線を含む）", size=10)

prs.save(dst)
print("saved", dst)
