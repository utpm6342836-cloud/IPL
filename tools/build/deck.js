const pptxgen = require("pptxgenjs");
const { applyTheme } = require("/root/.claude/skills/synced/e592ccf7-0a74-4018-9798-a3c1f17b50e9_d3c19d20-13f0-4b33-926a-3395a4da9b50/pptx/scripts/apply_theme.js");

const OUT = process.argv[2] || "deck_base.pptx";
const THEME = {
  name: "IPL Flow",
  headFontFace: "Meiryo",
  bodyFontFace: "Meiryo",
  colors: {
    dk1: "1E2B33", lt1: "FFFFFF", dk2: "1F3A4D", lt2: "EEF3F5",
    accent1: "0F7C8C", accent2: "D9862B", accent3: "5B7083", accent4: "A9C7CF",
    accent5: "7A9A3A", accent6: "B4472F", hlink: "0F7C8C", folHlink: "5B7083",
  },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.title = "IPL標準フロー v1.1 説明資料";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;

pres.defineSlideMaster({
  title: "TITLE_DARK",
  background: { color: THEME.colors.dk2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.2, w: 11.7, h: 1.6, fontSize: 40, bold: true, color: C.background1, valign: "bottom", align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.8, y: 4.0, w: 11.7, h: 1.4, fontSize: 20, color: THEME.colors.accent4 ? C.accent4 : C.background1, valign: "top", margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: THEME.colors.lt1 },
  margin: [0.5, 0.5, 0.6, 0.5],
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.35, w: 12.1, h: 0.8, fontSize: 28, bold: true, color: C.text2, valign: "middle", align: "left", margin: 0 }, text: "" } },
    { text: { text: "IPL標準フロー v1.1（たたき台）", options: { x: 0.6, y: 7.05, w: 6, h: 0.3, fontSize: 10, color: C.accent3, margin: 0 } } },
  ],
  slideNumber: { x: 12.2, y: 7.05, w: 0.5, h: 0.3, fontSize: 10, color: C.accent3 },
});

function title(slide, t) { slide.addText(t, { placeholder: "title" }); }
function card(slide, name, x, y, w, h, fill) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: fill || C.background2 }, line: { color: fill || C.background2 }, objectName: name });
}
function numCircle(slide, name, x, y, d, label, fill) {
  slide.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill || C.accent1 }, line: { color: fill || C.accent1 }, objectName: name });
  slide.addText(label, { x, y, w: d, h: d, align: "center", valign: "middle", fontSize: d > 0.9 ? 26 : 16, bold: true, color: C.background1, margin: 0, isTextBox: true, objectName: name + "_t" });
}
function bullets(items, size) {
  return items.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < items.length - 1, paraSpaceAfter: 6, fontSize: size || 14 } }));
}

// ---------- 1 表紙
pres.addSection({ title: "はじめに" });
let s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "はじめに" });
title(s, "IPL標準フロー");
s.addText("事業・研究の一人ひとりが、自らIPランドスケープを進めるために\nv1.1（たたき台）／2026年11月", { placeholder: "body" });
s.addNotes("本資料は、IPL標準フロー v1.1（たたき台）の説明資料です。手順書（Word）とワークシート集（Excel）とあわせて使います。");

// ---------- 2 目指すこと
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "はじめに" });
title(s, "目指すのは、各部門の人が自らIPLを進められる体制");
const why = [
  ["誰のためか", "事業部門（SBU）・研究開発部門で、新しい事業・製品・研究テーマを考えている一人ひとり"],
  ["なぜ自分でやるのか", "顧客と技術を一番知っているのは本人。自分で立てた仮説を自分で確かめるのが、早く確か"],
  ["一人ではない", "2026年度はIPL担当が伴走。2027年度からは手順書を読み込んだAIエージェントとIPL担当が支える"],
];
why.forEach((w, i) => {
  const x = 0.6 + i * 4.1;
  card(s, "why_card_" + i, x, 1.6, 3.8, 3.4);
  numCircle(s, "why_num_" + i, x + 0.3, 1.9, 0.7, String(i + 1));
  s.addText(w[0], { x: x + 0.3, y: 2.8, w: 3.2, h: 0.6, fontSize: 20, bold: true, color: C.text2, margin: 0, isTextBox: true, objectName: "why_head_" + i });
  s.addText(w[1], { x: x + 0.3, y: 3.45, w: 3.2, h: 1.9, fontSize: 15, color: C.text1, valign: "top", margin: 0, isTextBox: true, objectName: "why_body_" + i });
});
s.addText("IPL＝特許などの知財情報と、市場・顧客・競合などの事業情報を組み合わせて、事業や研究の判断に使うこと", { x: 0.6, y: 5.4, w: 12.1, h: 0.6, fontSize: 15, color: C.accent3, italic: true, margin: 0, isTextBox: true, objectName: "def" });
s.addNotes("このフローは、IPL担当に調査をお願いするためではなく、事業・研究の担当者が自分で仮説を立て、確かめ、決めるための手順書です。");

// ---------- 3 体制とスケジュール
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "はじめに" });
title(s, "2026年度は伴走、2027年度からはエージェントと進める");
const phases = [
  ["2026年度", "IPL担当と一緒に進める", ["IPL担当が手順書を使って伴走", "仮説の立て方・確かめ方を身につける", "〜12月：手順書を確定", "〜3月：AIエージェントを作成"]],
  ["2027年度〜", "エージェントとチャットしながら進める", ["手順書を読み込んだエージェントが案内", "分からない点はIPL担当に相談", "IPL担当は標準化チームに改善点を伝える", "振り返りで手順書とエージェントを改訂"]],
];
phases.forEach((p, i) => {
  const x = 0.6 + i * 6.25;
  card(s, "phase_card_" + i, x, 1.6, 5.85, 4.2, i === 0 ? C.background2 : C.background2);
  s.addText(p[0], { x: x + 0.35, y: 1.85, w: 5.2, h: 0.55, fontSize: 24, bold: true, color: i === 0 ? C.accent1 : C.accent2, margin: 0, isTextBox: true, objectName: "phase_year_" + i });
  s.addText(p[1], { x: x + 0.35, y: 2.45, w: 5.2, h: 0.5, fontSize: 17, bold: true, color: C.text2, margin: 0, isTextBox: true, objectName: "phase_head_" + i });
  s.addText(bullets(p[2], 15), { x: x + 0.35, y: 3.1, w: 5.2, h: 2.5, valign: "top", color: C.text1, margin: 0, isTextBox: true, objectName: "phase_list_" + i });
});
s.addShape(pres.shapes.CHEVRON, { x: 6.47, y: 3.45, w: 0.3, h: 0.5, fill: { color: C.accent3 }, line: { color: C.accent3 }, objectName: "phase_arrow" });
s.addText("どちらの時期も「調べ始める前（STEP2後）」と「決める前（STEP6後）」にIPL担当が確認する", { x: 0.6, y: 6.05, w: 12.1, h: 0.6, fontSize: 15, bold: true, color: C.text2, margin: 0, isTextBox: true, objectName: "phase_note" });
s.addNotes("ロードマップの図にもとづく体制です。確認ポイントは品質を保つための仕組みで、詳しくは後半のスライドで説明します。");

// ---------- 4 フローチャート（図形はpythonで追加）
pres.addSection({ title: "全体の流れ" });
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "全体の流れ" });
title(s, "全体の流れ：決める → 検証ループ → 意思決定 → 残す");
s.addNotes("実線は進む流れ、点線は戻る流れ（フィードバック）です。灰色の点線は、共有済みの図にはない戻り先の提案です。図形はPowerPointのコネクタで結んであるので、箱を動かすと線が追従します（戻りの線は自由線のため、箱を動かしたら手で直してください）。");

// ---------- 5 共通ルール
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "全体の流れ" });
title(s, "全STEPで守るのは3つだけ");
const rules = [
  ["事実と予想を分ける", "出典（URL・特許番号など）を示せないことは「未確認」と書く"],
  ["外れている証拠も探す", "都合のよい情報だけ集めない。見つからないことは正しい証拠にならない"],
  ["秘密情報は入れる前に確かめる", "出願前の発明・社外秘は、社内規程を確認してからAIに入力する"],
];
rules.forEach((r, i) => {
  const y = 1.55 + i * 1.2;
  numCircle(s, "rule_num_" + i, 0.6, y + 0.12, 0.75, String(i + 1));
  s.addText(r[0], { x: 1.6, y, w: 5.0, h: 0.5, fontSize: 19, bold: true, color: C.text2, margin: 0, isTextBox: true, objectName: "rule_head_" + i });
  s.addText(r[1], { x: 1.6, y: y + 0.5, w: 5.2, h: 0.55, fontSize: 14, color: C.text1, margin: 0, isTextBox: true, objectName: "rule_body_" + i });
});
s.addText("確かさのランク（判断の主な根拠はA・Bだけ）", { x: 7.3, y: 1.55, w: 5.4, h: 0.5, fontSize: 17, bold: true, color: C.text2, margin: 0, isTextBox: true, objectName: "rank_head" });
const rankRows = [
  [{ text: "ランク", options: { bold: true, color: C.background1, fill: { color: C.text2 } } }, { text: "条件", options: { bold: true, color: C.background1, fill: { color: C.text2 } } }],
  ["A", "元の資料を人が開いて確かめた"],
  ["B", "元の資料をAIが読んだだけ／独立した記事2つ以上"],
  ["C", "まとめ記事が1つだけ"],
  ["D", "出典を確かめられない（出典のないAIの回答など）"],
];
s.addTable(rankRows, { x: 7.3, y: 2.15, w: 5.4, colW: [0.9, 4.5], fontSize: 14, color: C.text1, border: { type: "solid", pt: 0.75, color: "C9D6DC" }, rowH: 0.55, valign: "middle", objectName: "rank_table" });
s.addText("AIが要約しただけの情報は最高でもB", { x: 7.3, y: 5.1, w: 5.4, h: 0.5, fontSize: 14, italic: true, color: C.accent3, margin: 0, isTextBox: true, objectName: "rank_note" });
s.addNotes("共通ルールは手順書の全体像6章にあります。確かさのランクはSTEP4で情報リストに付けます。");

// ---------- 6-14 STEPごと
pres.addSection({ title: "STEPごとの進め方" });
const steps = [
  ["0", "目的設定", "何を調べるかではなく、何を決めるために調べるかを1文にする", ["依頼を言い換えずに書き写す", "決めたいことを1文にする", "会社の方針と合っているか確かめる", "意思決定者・期限・範囲を書く", "Goとやめる条件を1文ずつ書く", "進め方の大きさと飛ばすSTEPを決める"], "依頼シート", "意思決定者が「何を決めるか」「Go／やめる条件」を1文で言える", "企画書はゴールではなくスタート。分からないことを見える化する"],
  ["1", "顧客価値仮説（売れる）", "最初に誰が買うか（一社・一人）と、なぜ買うかを手元の情報で書く", ["気になる情報1件から「だから何か」を書く", "顧客候補を3〜5つ挙げる（直接／その先）", "ファーストユーザーを1つ決める", "今の状態とありたい状態を対にする", "ロジックツリーとトレードオフを書く", "価値の仮説を1文にする"], "エッジ情報カード、困りごとシート、価値の仮説", "ファーストユーザーが会社名・部署まで書け、価値の仮説を顧客に聞いて賛否が分かる", "顧客はセグメントではなく、会いたいと思える一社・一人"],
  ["2", "事業仮説（勝てる・儲かる・できそう）", "勝てる（守れるを含む）・儲かる・できそうの仮説と、崩れる条件を書く", ["企画骨子5項目をまず書く", "勝てる：なぜ他はまだやっていないか、守れるか", "儲かる：1社分の売上から考える", "できそう：量産・規制・資源", "全仮説に崩れる条件、一番あやしい仮説に印", "IPL担当に見てもらう【確認①】"], "企画骨子、事業仮説シート", "すべての仮説に崩れる条件がある／IPL担当の確認①", "差別化ではなく「この会社から買う理由（独占）」を考える"],
  ["3", "調査設計", "仮説 → 必要な情報 → 場所 → 検索式 の順で調べ方を決める", ["仮説ごとに必要な情報を書く", "調べる場所を決める", "検索を課題・技術・会社の3層に分ける", "検索式の案を作る", "代表的な特許が拾えるか確かめる", "外れている証拠の探し方・やめる条件を書く"], "調査計画", "全仮説に外れている証拠の探し方とやめる条件がある", "課題を先読みして検索する（仮説検索）"],
  ["4", "情報探索", "情報を集め、1件ごとに出典・取得日・確かさのランクを付ける", ["印を付けた仮説から集める", "情報リストに1件ずつ書く（出典なしの行は作らない）", "大事な特許は人・技術・権利で読む", "確かさのランクを付ける", "1日の終わりに今日の結論を1行"], "情報リスト", "全行に出典・ランクがあり、外れている証拠の行もある", "特許から意図・キーマン・共同研究先を読む"],
  ["5", "構造化・分析", "まとめて分類し、表や図にして「だから何か」を書く", ["同じ発明をまとめる（ファミリー）", "会社名をまとめる（名寄せ）", "仮説に合わせて切り口を決める", "表や図にする（数え方を明記）", "件数・空白の読み方の注意を書く", "表ごとに「だから何か」を3行"], "整理表・マップ", "どの表も確かめる仮説IDがあり、他の人が作り直せる", "網羅性を求めすぎない。仮説を動かす表を1つ"],
  ["6", "仮説検証・反証", "反対の立場から見て、仮説を持ち込んで顧客に聞き、残す・直す・捨てるを決める", ["支える情報と外れている証拠を並べる", "反対の立場から見てもらう", "よくある思い込みを確かめる", "尖った仮説を提案書にする", "自分で聞きに行く（違うと言われたことも記録）", "IPL担当に見てもらう【確認②】"], "仮説の記録、ヒアリング票", "反証を探した記録があり、根拠がランクA・B／IPL担当の確認②", "自分の仮説は自分で確かめる。先行例との違いで育てる"],
  ["7", "依頼者が意思決定（Go / No Go）", "根拠と残る不確かさをそろえた判断メモで、依頼者が決める", ["4観点を4段階で評価する", "儲かるを最初の顧客から積み上げる", "選択肢を並べる（No Goはやめる／方向を変える）", "判断メモを書く", "依頼者が決めて記録する", "必要なら経営層へ提言"], "判断メモ", "結論・次の行動・やめる条件・まだ分からないことがそろっている", "行き詰まったら方向転換（Pivot）で育てる"],
  ["8", "学習・標準化", "振り返って、次の人が使えるように残す", ["振り返りアンケートに答える（15分）", "考えた経過と捨てた仮説を残す", "AIの間違いを残す", "社外に出してよいか決めて保存", "1〜2年後の振り返り日を決める"], "振り返りアンケート、ノウハウ集", "他の人が同じ手順をやり直せる", "分析途中まで残し、考えの道筋を追えるようにする"],
];
steps.forEach((st) => {
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "STEPごとの進め方" });
  title(s, "STEP" + st[0] + "　" + st[1]);
  numCircle(s, "step_num", 0.6, 1.45, 1.0, st[0]);
  s.addText(st[2], { x: 1.85, y: 1.45, w: 10.85, h: 1.0, fontSize: 18, bold: true, color: C.text2, valign: "middle", margin: 0, isTextBox: true, objectName: "step_goal" });
  s.addText("やり方", { x: 0.6, y: 2.75, w: 6.6, h: 0.45, fontSize: 16, bold: true, color: C.accent1, margin: 0, isTextBox: true, objectName: "step_how_h" });
  s.addText(st[3].map((t, i) => ({ text: t, options: { bullet: { type: "number" }, breakLine: i < st[3].length - 1, paraSpaceAfter: 6, fontSize: 15 } })), { x: 0.6, y: 3.25, w: 6.6, h: 3.5, valign: "top", color: C.text1, margin: 0, isTextBox: true, objectName: "step_how" });
  const boxes = [["できあがるもの", st[4]], ["次に進む前のチェック", st[5]], ["発明塾のコツ", st[6]]];
  boxes.forEach((b, i) => {
    const y = 2.75 + i * 1.33;
    card(s, "step_box_" + i, 7.6, y, 5.1, 1.18);
    s.addText(b[0], { x: 7.85, y: y + 0.1, w: 4.7, h: 0.35, fontSize: 13, bold: true, color: i === 2 ? C.accent2 : C.accent1, margin: 0, isTextBox: true, objectName: "step_box_h_" + i });
    s.addText(b[1], { x: 7.85, y: y + 0.45, w: 4.7, h: 0.68, fontSize: 14, color: C.text1, valign: "top", margin: 0, isTextBox: true, objectName: "step_box_b_" + i });
  });
  s.addNotes("詳しい手順と目安時間は、手順書（Word）のSTEP" + st[0] + "を参照してください。ひな形はワークシート集（Excel）にあります。");
});

// ---------- 15 品質
pres.addSection({ title: "仕組み" });
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "仕組み" });
title(s, "品質は2つの確認ポイントと、確かさのランクで保つ");
const q = [
  ["確認ポイント①", "調べ始める前（STEP2の後）", "目的・仮説・調べる順番がずれていないか"],
  ["確認ポイント②", "決める前（STEP6の後）", "根拠がランクA・Bか、反証を探したか"],
  ["エージェントのテスト", "改訂のたび", "テスト質問20問で、手順書どおりに案内するか"],
  ["事後の振り返り", "判断の1〜2年後", "判断は当たっていたか、何を見落としたか"],
];
q.forEach((r, i) => {
  const x = 0.6 + (i % 2) * 6.25, y = 1.55 + Math.floor(i / 2) * 2.45;
  card(s, "q_card_" + i, x, y, 5.85, 2.15);
  s.addText(r[0], { x: x + 0.35, y: y + 0.25, w: 5.2, h: 0.5, fontSize: 20, bold: true, color: C.accent1, margin: 0, isTextBox: true, objectName: "q_h_" + i });
  s.addText(r[1], { x: x + 0.35, y: y + 0.8, w: 5.2, h: 0.45, fontSize: 15, bold: true, color: C.text2, margin: 0, isTextBox: true, objectName: "q_when_" + i });
  s.addText(r[2], { x: x + 0.35, y: y + 1.3, w: 5.2, h: 0.65, fontSize: 15, color: C.text1, valign: "top", margin: 0, isTextBox: true, objectName: "q_b_" + i });
});
s.addText("未確認：IPL担当の人数で、確認ポイントの件数を回せるか", { x: 0.6, y: 6.5, w: 12.1, h: 0.4, fontSize: 13, italic: true, color: C.accent6, margin: 0, isTextBox: true, objectName: "q_note" });
s.addNotes("ロードマップ図の「品質はどう担保？」への案です。案件の大きさによって確認の濃さを変える案もあります（手順書 付録 体制・品質・蓄積の論点）。");

// ---------- 16 検証
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "仕組み" });
title(s, "「このSTEPで正しいか」は6つの検証で確かめる");
const vRows = [
  ["検証", "何が分かるか", "v1.1の状況"],
  ["V1 構造チェック", "目的・入力・出力・チェック・戻り先がそろっているか", "実施済み（5件修正後、全STEP合格）"],
  ["V2 既存資料との照合", "発明塾・過去の手順書の要素が抜けていないか", "実施済み（2資料は未照合）"],
  ["V3 公的な基準との照合", "特許庁ガイドブックと比べた抜け", "未実施（作成環境から接続できず）"],
  ["V4 過去案件への当てはめ", "実際の企画に合うか", "1件試行（STEP0・守れるが抜けやすい）"],
  ["V5 パイロット運用", "事業・研究の人が実際に使えるか", "2026年度に2〜3件"],
  ["V6 エージェントの検証", "手順書どおりに案内するか", "テスト質問20問を用意済み"],
].map((r, i) => r.map((c) => (i === 0 ? { text: c, options: { bold: true, color: C.background1, fill: { color: C.text2 } } } : c)));
s.addTable(vRows, { x: 0.6, y: 1.5, w: 12.1, colW: [3.2, 5.0, 3.9], fontSize: 14, color: C.text1, border: { type: "solid", pt: 0.75, color: "C9D6DC" }, rowH: 0.62, valign: "middle", objectName: "v_table" });
s.addText("最も重いのはV5。1件の当てはめや自動チェックだけで「正しい」とは言えない", { x: 0.6, y: 6.15, w: 12.1, h: 0.5, fontSize: 15, bold: true, color: C.text2, margin: 0, isTextBox: true, objectName: "v_note" });
s.addNotes("詳しくは手順書の付録「フロー検証の進め方」。構造チェックは tools/check_flow.py で自動化しています。");

// ---------- 17 発明塾との関係
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "仕組み" });
title(s, "STEP名は社内の言葉のまま、発明塾の手法は各STEPの中で使う");
const hj = [
  ["STEP1", "エッジ情報、ファーストユーザー、課題‐解決ロジックツリー、トレードオフ"],
  ["STEP2", "企画骨子5項目、独占、他の人がやっていない理由"],
  ["STEP3", "課題を先読みした検索（仮説検索）"],
  ["STEP4", "人・技術・権利の3つの見方"],
  ["STEP6", "仮説を持ち込むヒアリング、先行例との差分で育てる、良い仲間と良い議論"],
  ["STEP7・8", "Pivot、考えの道筋を残す"],
];
hj.forEach((r, i) => {
  const y = 1.5 + i * 0.78;
  card(s, "hj_tag_" + i, 0.6, y, 1.6, 0.6, C.accent1);
  s.addText(r[0], { x: 0.6, y, w: 1.6, h: 0.6, fontSize: 15, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "hj_tag_t_" + i });
  s.addText(r[1], { x: 2.45, y, w: 5.6, h: 0.6, fontSize: 15, color: C.text1, valign: "middle", margin: 0, isTextBox: true, objectName: "hj_b_" + i });
});
card(s, "hj_reason", 8.5, 1.5, 4.2, 4.5);
s.addText("この形にした理由", { x: 8.8, y: 1.75, w: 3.7, h: 0.5, fontSize: 17, bold: true, color: C.text2, margin: 0, isTextBox: true, objectName: "hj_reason_h" });
s.addText(bullets(["共有済みのフローと発明塾は、順番の考え方がほぼ同じ（顧客価値が先、仮説を立ててから調べる）", "受講していない人にも意味が伝わる", "社外講座の利用条件に左右されにくい"], 14), { x: 8.8, y: 2.35, w: 3.7, h: 3.5, valign: "top", color: C.text1, margin: 0, isTextBox: true, objectName: "hj_reason_b" });
s.addNotes("STEP名から何をするか分からないという声が3割以上あれば、STEP名に発明塾の言葉を副題として付ける案に切り替えます。");

// ---------- 18 次の進め方
pres.addSection({ title: "まとめ" });
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "まとめ" });
title(s, "次に決めること・確かめること");
const nx = [
  ["発明塾資料の利用条件", "社内共有・エージェントのナレッジに使ってよいか、契約窓口に確認"],
  ["共有済みの図の見直し", "守れる、ヒアリング、STEP2への戻り、方向転換を図に足すか"],
  ["パイロット案件の選定", "2〜3件で試し、目安時間とワークシートを直す"],
  ["社内の他の取り組みとの連携", "R&Dプラットフォームなどと、審査・置き場所が重ならないか"],
];
nx.forEach((r, i) => {
  const y = 1.5 + i * 1.25;
  numCircle(s, "nx_num_" + i, 0.6, y + 0.1, 0.8, String(i + 1), C.accent2);
  s.addText(r[0], { x: 1.7, y, w: 10.9, h: 0.5, fontSize: 19, bold: true, color: C.text2, margin: 0, isTextBox: true, objectName: "nx_h_" + i });
  s.addText(r[1], { x: 1.7, y: y + 0.52, w: 10.9, h: 0.5, fontSize: 15, color: C.text1, margin: 0, isTextBox: true, objectName: "nx_b_" + i });
});
s.addNotes("いずれも手順書の付録（出典と改訂履歴の未確認事項、呼び方の決め方、体制・品質・蓄積の論点）に詳しく書いています。");

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written", OUT);
})();
