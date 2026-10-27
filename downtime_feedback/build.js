const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const { applyTheme } = require("/root/.claude/skills/synced/6be4650e-3a94-4c90-b923-b63667a03645_8a9e4964-db73-460c-97dd-30feef43c407/pptx/scripts/apply_theme.js");

const OUT = process.argv[2] || "downtime_feedback.pptx";
const FONT = "BIZ UDPゴシック";

const HEX = {
  plum: "4A2245", rose: "C7487A", roseTint: "F7E8EF", teal: "17796F",
  tealTint: "E3F2EF", ink: "2B2230", gray: "6E6470", white: "FFFFFF",
};
const THEME = {
  name: "Downtime Feedback",
  headFontFace: FONT, bodyFontFace: FONT,
  colors: {
    dk1: HEX.ink, lt1: HEX.white, dk2: HEX.plum, lt2: HEX.roseTint,
    accent1: HEX.rose, accent2: HEX.teal, accent3: HEX.plum, accent4: "E9A8C1",
    accent5: HEX.tealTint, accent6: HEX.gray, hlink: HEX.teal, folHlink: HEX.plum,
  },
};

async function icon(Comp, color) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color: "#" + color, size: 256 }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
  pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
  pres.title = "ダウンタイム 振り返りフィードバック";
  const C = pres.SchemeColor;
  // scheme aliases
  const PLUM = C.text2, INK = C.text1, WHITE = C.background1, ROSE_T = C.background2;
  const ROSE = C.accent1, TEAL = C.accent2, TEAL_T = C.accent5, GRAY = C.accent6;

  const ICONS = {
    quote: await icon(fa.FaQuoteLeft, HEX.white),
    q: await icon(fa.FaQuestion, HEX.white),
    eye: await icon(fa.FaEye, HEX.white),
    sync: await icon(fa.FaExchangeAlt, HEX.white),
    heart: await icon(fa.FaHeartbeat, HEX.white),
    hands: await icon(fa.FaHandsHelping, HEX.white),
    users: await icon(fa.FaUsers, HEX.white),
    flag: await icon(fa.FaFlagCheckered, HEX.white),
    stetho: await icon(fa.FaStethoscope, HEX.white),
    light: await icon(fa.FaLightbulb, HEX.white),
  };

  // ---------- layouts ----------
  pres.defineSlideMaster({
    title: "TITLE_DARK",
    background: { color: PLUM },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.9, y: 1.5, w: 11.5, h: 1.9, fontSize: 48, bold: true, color: WHITE, valign: "bottom", align: "left" }, text: "" } },
      { placeholder: { options: { name: "body", type: "body", x: 0.9, y: 3.7, w: 11.5, h: 3.0, fontSize: 24, color: "E9A8C1", valign: "top", align: "left" }, text: "" } },
    ],
  });
  pres.defineSlideMaster({
    title: "SECTION",
    background: { color: PLUM },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 3.4, y: 2.6, w: 9.3, h: 1.3, fontSize: 44, bold: true, color: WHITE, valign: "middle", align: "left" }, text: "" } },
      { placeholder: { options: { name: "body", type: "body", x: 3.4, y: 3.9, w: 9.3, h: 1.0, fontSize: 24, color: "E9A8C1", valign: "top" }, text: "" } },
    ],
    slideNumber: { x: 12.3, y: 6.9, w: 0.7, h: 0.4, fontSize: 14, color: "E9A8C1", align: "right" },
  });
  pres.defineSlideMaster({
    title: "CONTENT",
    background: { color: WHITE },
    objects: [
      { placeholder: { options: { name: "tag", type: "body", x: 0.7, y: 0.3, w: 8, h: 0.4, fontSize: 16, color: ROSE, bold: true, margin: 0 }, text: "" } },
      { placeholder: { options: { name: "title", type: "title", x: 0.7, y: 0.7, w: 11.9, h: 1.0, fontSize: 36, bold: true, color: PLUM, valign: "middle", margin: 0, align: "left" }, text: "" } },
    ],
    slideNumber: { x: 12.3, y: 6.9, w: 0.7, h: 0.4, fontSize: 14, color: GRAY, align: "right" },
  });
  pres.defineSlideMaster({
    title: "QUESTION",
    background: { color: TEAL },
    objects: [
      { placeholder: { options: { name: "tag", type: "body", x: 0.9, y: 0.6, w: 10, h: 0.5, fontSize: 22, color: WHITE, bold: true, margin: 0 }, text: "" } },
      { placeholder: { options: { name: "title", type: "title", x: 0.9, y: 1.5, w: 11.5, h: 3.6, fontSize: 44, bold: true, color: WHITE, valign: "middle", margin: 0, align: "left" }, text: "" } },
      { placeholder: { options: { name: "body", type: "body", x: 0.9, y: 5.4, w: 11.5, h: 1.3, fontSize: 24, color: WHITE, valign: "top", margin: 0 }, text: "" } },
    ],
    slideNumber: { x: 12.3, y: 6.9, w: 0.7, h: 0.4, fontSize: 14, color: WHITE, align: "right" },
  });

  let section = "";
  const add = (master, sec) => {
    if (sec && sec !== section) { pres.addSection({ title: sec }); section = sec; }
    return pres.addSlide({ masterName: master, sectionTitle: section });
  };
  const content = (sec, tag, title) => {
    const s = add("CONTENT", sec);
    s.addText(tag, { placeholder: "tag" });
    s.addText(title, { placeholder: "title" });
    return s;
  };
  const circleIcon = (s, data, x, y, d, fill) => {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { type: "none" }, objectName: "icon-circle" });
    const p = d * 0.26;
    s.addImage({ data, x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p, altText: "", objectName: "icon" });
  };
  // quote card: rose tint, round, quote icon top-left
  const quote = (s, x, y, w, h, text, size = 24, fill = ROSE_T, iconFill = ROSE) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.15, fill: { color: fill }, line: { type: "none" }, objectName: "quote-card" });
    circleIcon(s, ICONS.quote, x + 0.3, y + 0.3, 0.6, iconFill);
    s.addText(text, { x: x + 1.1, y: y + 0.25, w: w - 1.4, h: h - 0.5, fontSize: size, color: INK, valign: "middle", isTextBox: true, margin: 0, lineSpacingMultiple: 1.15, objectName: "quote-text" });
  };
  const point = (s, text, y = 6.05) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.7, y, w: 11.9, h: 0.85, rectRadius: 0.12, fill: { color: PLUM }, line: { type: "none" }, objectName: "point-bar" });
    s.addText([{ text: "POINT  ", options: { bold: true, color: "E9A8C1" } }, { text, options: { color: WHITE, bold: true } }],
      { x: 1.0, y, w: 11.4, h: 0.85, fontSize: 24, valign: "middle", isTextBox: true, margin: 0, objectName: "point-text" });
  };
  const question = (sec, tag, title, body, notes) => {
    const s = add("QUESTION", sec);
    s.addText(tag, { placeholder: "tag" });
    s.addText(title, { placeholder: "title" });
    if (body) s.addText(body, { placeholder: "body" });
    s.addShape(pres.shapes.OVAL, { x: 11.2, y: 0.4, w: 1.4, h: 1.4, fill: { color: WHITE, transparency: 80 }, line: { type: "none" }, objectName: "q-circle" });
    s.addImage({ data: ICONS.q, x: 11.55, y: 0.75, w: 0.7, h: 0.7, altText: "", objectName: "q-icon" });
    if (notes) s.addNotes(notes);
    return s;
  };
  const divider = (sec, num, title, sub, ic) => {
    const s = add("SECTION", sec);
    circleIcon(s, ic, 0.9, 2.55, 2.0, ROSE);
    s.addText(num, { x: 3.4, y: 1.9, w: 6, h: 0.6, fontSize: 22, color: "E9A8C1", bold: true, isTextBox: true, margin: 0 });
    s.addText(title, { placeholder: "title" });
    s.addText(sub, { placeholder: "body" });
    return s;
  };

  // ================= 0. Opening =================
  let s = add("TITLE_DARK", "はじめに");
  s.addText("「ダウンタイム」を観て\nみんなが考えたこと", { placeholder: "title" });
  s.addText("振り返りフィードバック ― 看護科 2年", { placeholder: "body" });
  s.addNotes("前回のワークシートと振り返りから、クラスの考えを紹介します。今日は「正解」を出すのではなく、自分とは違う見方に出会うことが目的です。");

  s = content("はじめに", "今日のねらい", "友だちの言葉で、自分の考えを一歩深める");
  const goals = [
    [ICONS.eye, "自分が気づかなかった見方に出会う", "同じドラマでも、心に残った場面も理由も人それぞれ"],
    [ICONS.light, "「なぜそう思う？」をもう一段深く", "答えが同じでも、理由の深さはちがう"],
    [ICONS.stetho, "看護職をめざす立場で考える", "患者さんの「大切なもの」にどう向き合うか"],
  ];
  goals.forEach(([ic, h, b], i) => {
    const y = 2.1 + i * 1.3;
    circleIcon(s, ic, 0.9, y, 0.95, i === 2 ? TEAL : ROSE);
    s.addText(h, { x: 2.2, y: y - 0.05, w: 10, h: 0.6, fontSize: 28, bold: true, color: INK, isTextBox: true, margin: 0 });
    s.addText(b, { x: 2.2, y: y + 0.55, w: 10, h: 0.45, fontSize: 20, color: GRAY, isTextBox: true, margin: 0 });
  });
  s.addText("※ 紹介する意見はすべて匿名です。「誰が書いたか」は探さないでください。", { x: 0.7, y: 6.2, w: 11.9, h: 0.5, fontSize: 18, color: GRAY, isTextBox: true, margin: 0 });
  s.addNotes("引用は誤字のみ最小限に整えています。誰の意見かを詮索しないこと、どの意見も否定しないことを最初に確認します。");

  // ================= 1. 観る前 =================
  const S1 = "1 観る前の私たち";
  divider(S1, "PART 1", "観る前の私たち", "「美容医療」ってどんなイメージだった？", ICONS.eye);

  s = content(S1, "PART 1 観る前の私たち", "「美容医療」と聞いて思い浮かんだ言葉");
  const words = [
    ["整形", 54, 0.9, 2.0, ROSE], ["脱毛", 44, 3.3, 2.15, PLUM], ["脂肪吸引", 40, 5.3, 2.2, ROSE],
    ["二重", 34, 8.3, 2.25, PLUM], ["ボトックス", 30, 0.9, 3.25, PLUM], ["ヒアルロン酸", 28, 4.0, 3.3, ROSE],
    ["綺麗にする", 26, 7.6, 3.32, GRAY], ["豊胸", 26, 10.4, 3.32, PLUM], ["高い", 22, 0.9, 4.2, GRAY],
    ["お金持ち", 22, 2.6, 4.2, GRAY], ["コンプレックスを治す", 22, 4.9, 4.2, GRAY], ["顔変わる", 22, 9.0, 4.2, GRAY],
  ];
  words.forEach(([t, sz, x, y, c]) => s.addText(t, { x, y, w: 4.2, h: 0.9, fontSize: sz, bold: true, color: c, isTextBox: true, margin: 0, objectName: "word" }));
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.7, y: 5.25, w: 11.9, h: 1.55, rectRadius: 0.15, fill: { color: TEAL_T }, line: { type: "none" }, objectName: "callout" });
  s.addText([
    { text: "「リスク」「失敗」「医師」「病院」と書いた人は ", options: { color: INK } },
    { text: "0人", options: { color: TEAL, bold: true, fontSize: 40 } },
  ], { x: 1.0, y: 5.25, w: 11.4, h: 0.95, fontSize: 26, valign: "middle", isTextBox: true, margin: 0 });
  s.addText("観る前は「何ができるか（技術）」が中心。観た後、何が見えるようになった？", { x: 1.0, y: 6.1, w: 11.4, h: 0.6, fontSize: 20, color: GRAY, isTextBox: true, margin: 0 });
  s.addNotes("クラス全員が書いた言葉を大きさでざっくり表しています（大きいほど多い）。ほとんどが「施術名」で、リスクや医療者に関する言葉はゼロでした。これは悪いことではなく、それだけ身近になっている証拠。観た後に何が見えるようになったかを意識しながら次へ。");

  s = content(S1, "PART 1 観る前の私たち", "「医療」だと思うのはどれ？（26人が回答）");
  s.addChart(pres.charts.BAR, [{ name: "人数", labels: ["骨折の手術", "口唇裂の形成手術", "乳がん術後の乳房再建", "歯列の矯正", "二重まぶた手術", "脱毛"], values: [26, 21, 21, 18, 14, 8] }], {
    x: 0.7, y: 1.9, w: 7.6, h: 4.9, barDir: "bar", catAxisOrientation: "maxMin",
    chartColors: [HEX.rose], showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 20, dataLabelColor: HEX.ink, dataLabelFontBold: true,
    catAxisLabelFontSize: 20, catAxisLabelColor: HEX.ink, catAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt",
    valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" }, valAxisMaxVal: 30, barGapWidthPct: 45, showLegend: false,
  });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 8.7, y: 2.0, w: 3.9, h: 4.7, rectRadius: 0.15, fill: { color: ROSE_T }, line: { type: "none" }, objectName: "side-card" });
  s.addText([
    { text: "全員一致は\n「骨折」だけ", options: { fontSize: 30, bold: true, color: PLUM, breakLine: true } },
    { text: " ", options: { fontSize: 14, breakLine: true } },
    { text: "そのほかは「どこで線を引くか」が一人ひとり違った。\n\n医療とそうでないものの線は、何で決まる？", options: { fontSize: 22, color: INK } },
  ], { x: 9.0, y: 2.2, w: 3.4, h: 4.3, valign: "top", isTextBox: true, margin: 0 });
  s.addNotes("骨折は全員一致。脱毛は8人。グラデーションになっていることに注目させる。");

  s = content(S1, "PART 1 観る前の私たち", "みんなは何を基準に線を引いた？");
  quote(s, 0.7, 1.95, 11.9, 1.75, "怪我とか生まれつきの病気を「治したり戻したり」しているもの。選んでないのは、病気じゃなくて自分の意思や見た目の変換をしているから。", 24);
  quote(s, 0.7, 3.95, 11.9, 1.75, "綺麗になりたい人の手伝いをしているのが美容。でも、医療ではないから「美容医療」じゃない名前にした方が良い。", 24);
  point(s, "「治す・戻す」は医療、「変える」は医療じゃない？");
  s.addNotes("多くの人がこの基準で線を引いていました。とても筋の通った考え方です。では本当にこの基準ですっきり分けられるのか、次のスライドで揺さぶります。");

  question(S1, "考えてみよう ― となりの人と 2分", "口唇裂の形成手術も、乳房再建も\n「見た目」を整える手術。\n二重まぶた手術と、何がちがう？",
    "ヒント：ドラマの問い「この世界に病気じゃない人なんていない？」",
    "発問のねらい：「治す/変える」の線引きは、実は見た目・生きづらさ・本人の苦しみと深く関わっていることに気づかせる。口唇裂も乳房再建も、命には直接関わらないが、その人の生活の質や心にとって重要。では二重まぶたで深く悩む人の苦しみはどうか？『誰が』『何を基準に』病気と決めるのか、を考えさせる。眼瞼下垂（まぶたが下がる病気）は保険診療になる、という例も境界を考える材料になる。");

  s = content(S1, "PART 1 観る前の私たち", "知っておこう：美容医療も「医療」です");
  const facts = [
    [ICONS.stetho, "医師だけが行える「医療行為」", "メスを入れる・注射をする・レーザーを当てる ― 美容目的でも医療"],
    [ICONS.flag, "多くは「自由診療」（保険がきかない）", "料金はクリニックが決める。だから広告や価格の競争も起きる"],
    [ICONS.hands, "だからこそ「説明と同意」が欠かせない", "病気の治療ではない分、本人が納得して選ぶことがより大切に"],
  ];
  facts.forEach(([ic, h, b], i) => {
    const y = 1.95 + i * 1.3;
    circleIcon(s, ic, 0.9, y, 0.95, TEAL);
    s.addText(h, { x: 2.2, y: y - 0.05, w: 10.3, h: 0.6, fontSize: 28, bold: true, color: INK, isTextBox: true, margin: 0 });
    s.addText(b, { x: 2.2, y: y + 0.55, w: 10.3, h: 0.45, fontSize: 20, color: GRAY, isTextBox: true, margin: 0 });
  });
  s.addText("厚生労働省「美容医療を受ける前に」 https://aesthetic-medicine-caution.mhlw.go.jp/", { x: 0.7, y: 6.1, w: 11.9, h: 0.5, fontSize: 18, color: TEAL, isTextBox: true, margin: 0, hyperlink: { url: "https://aesthetic-medicine-caution.mhlw.go.jp/" } });
  s.addNotes("「美容医療じゃない名前にした方が良い」という意見への応答にもなる。名前を変えると医療としてのルール（医師が行う、説明と同意など）から外れてしまう、という視点も紹介できる。ワークシートに載せた厚労省サイトをここで改めて確認。");

  // ================= 2. 観る前→観た後 =================
  const S2 = "2 観る前→観た後";
  divider(S2, "PART 2", "観る前 → 観た後", "考えは変わった？ 変わらなかった？", ICONS.sync);

  s = content(S2, "PART 2 観る前 → 観た後", "選んだ答えは、あまり変わらなかった");
  s.addChart(pres.charts.BAR, [
    { name: "観る前（美容医療）", labels: ["自分も他の人も「あり」", "自分は「なし」・他の人は「あり」", "自分も他の人も「なし」", "自分は「あり」・他の人は「なし」"], values: [18, 7, 0, 0] },
    { name: "観た後（美容整形）", labels: ["自分も他の人も「あり」", "自分は「なし」・他の人は「あり」", "自分も他の人も「なし」", "自分は「あり」・他の人は「なし」"], values: [14, 8, 1, 0] },
  ], {
    x: 0.7, y: 1.85, w: 7.9, h: 4.95, barDir: "bar", barGrouping: "clustered", catAxisOrientation: "maxMin",
    chartColors: ["E9A8C1", HEX.plum], showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 18, dataLabelColor: HEX.ink, dataLabelFontBold: true,
    catAxisLabelFontSize: 16, catAxisLabelColor: HEX.ink, catAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt",
    valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" }, valAxisMaxVal: 21, barGapWidthPct: 40,
    showLegend: true, legendPos: "t", legendFontSize: 18, legendFontFace: "+mn-lt", legendColor: HEX.ink,
  });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 8.9, y: 1.95, w: 3.7, h: 4.85, rectRadius: 0.15, fill: { color: ROSE_T }, line: { type: "none" }, objectName: "side-card" });
  s.addText([
    { text: "「自分はあり・\n他の人はなし」", options: { fontSize: 24, bold: true, color: PLUM, breakLine: true } },
    { text: "前も後も0人", options: { fontSize: 24, bold: true, color: ROSE, breakLine: true } },
    { text: " ", options: { fontSize: 12, breakLine: true } },
    { text: "「他の人の選択は否定しない」がクラスの共通点。\n\nでも、変わったのは数字より「理由」だった。", options: { fontSize: 20, color: INK } },
  ], { x: 9.15, y: 2.15, w: 3.25, h: 4.5, valign: "top", isTextBox: true, margin: 0 });
  s.addNotes("観る前は「美容医療」、観た後は「美容整形」についての問いなので厳密には同じ質問ではない点に注意。観る前は無回答・複数選択が各1人、観た後は無回答4人。「あり」から「自分はなし」へ動いた人が2人、「なし」まで動いた人が1人。");

  s = content(S2, "PART 2 観る前 → 観た後", "答えは同じ「あり」でも、理由が変わった");
  const pairs = [
    ["好きな自分になれるならいい", "ちゃんと検査とかしてから手術した方がいい"],
    ["気になるなら整形したらいいじゃん", "命を守るためにシリコンを抜いて、そこまで追い込まれるならだめ"],
    ["自分の好きなようにしたらいいやん", "失敗した時に感情をコントロールできて、依存しないならやってもいい"],
  ];
  s.addText("観る前", { x: 0.7, y: 1.8, w: 4.3, h: 0.45, fontSize: 20, bold: true, color: GRAY, isTextBox: true, margin: 0 });
  s.addText("観た後", { x: 6.4, y: 1.8, w: 6.2, h: 0.45, fontSize: 20, bold: true, color: ROSE, isTextBox: true, margin: 0 });
  pairs.forEach(([a, b], i) => {
    const y = 2.3 + i * 1.22;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.7, y, w: 5.0, h: 1.05, rectRadius: 0.12, fill: { color: "F1EEF2" }, line: { type: "none" }, objectName: "before" });
    s.addText(a, { x: 0.95, y, w: 4.6, h: 1.05, fontSize: 20, color: GRAY, valign: "middle", isTextBox: true, margin: 0 });
    s.addShape(pres.shapes.RIGHT_ARROW, { x: 5.8, y: y + 0.32, w: 0.5, h: 0.42, fill: { color: ROSE }, line: { type: "none" }, objectName: "arrow" });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.4, y, w: 6.2, h: 1.05, rectRadius: 0.12, fill: { color: ROSE_T }, line: { type: "none" }, objectName: "after" });
    s.addText(b, { x: 6.65, y, w: 5.8, h: 1.05, fontSize: 20, bold: true, color: INK, valign: "middle", isTextBox: true, margin: 0 });
  });
  point(s, "「好きにすれば」から「条件つきのあり」へ");
  s.addNotes("答えの選択肢は同じでも、「どんな条件なら」を考えられるようになったことが大きな成長。何人かに「あなたの『条件』は？」と聞いてみる。");

  s = content(S2, "PART 2 観る前 → 観た後", "「自己責任」をもう一歩深く");
  quote(s, 0.7, 1.95, 11.9, 1.75, "整形は本人の自己責任だから好きにすれば良いと思ってたけど、患者が納得いかず文句を言えば、医者は理不尽に炎上してしまう。「したければすれば良い」も簡単には言えないと思った。", 22);
  quote(s, 0.7, 3.95, 11.9, 1.75, "本人が生きやすくなるなら良い。そのことによって起こるリスクや問題について、しっかりと理解しているなら。", 22, TEAL_T, TEAL);
  point(s, "「自己責任」は、十分に知って・理解して選んだときに成り立つ");
  s.addNotes("「自己責任」という言葉はクラスで何度も出てきた。でも、リスクを知らされていない人、理解できていない人、追い詰められて冷静に選べない人に「自己責任」は言えるだろうか？ → これが後半のインフォームド・コンセントにつながる。");

  question(S2, "考えてみよう ― となりの人と 2分", "「他の人がするのはあり」\nその「他の人」が\n家族や親友だったら？",
    "「仲の良い人や家族だったらやって欲しくない。でも私はその人の気持ちを大切にしたい」（クラスの声）",
    "多くの人が「他人の選択は自由」と答えたが、その他人を具体的な大切な人に置き換えると気持ちが揺れる。揺れること自体が大事。看護師にとって患者さんは『他人』だが、家族にとっては大切な人。その間で、看護師はどう関わる？");

  s = content(S2, "PART 2 観る前 → 観た後", "「変わらなかった」も立派な答え");
  quote(s, 0.7, 1.95, 11.9, 2.45, "観る前から他の人の整形は肯定していて、ドラマを見ても考えは変わらなかった。相手がどれだけ自分の顔や体型について悩み、苦労して、その辛さを改善するためにとった行動に対して、僕の考えを強要するような理不尽なことはしたくないので。", 24);
  s.addText("自分はしない。でも、相手の悩みの重さは自分にはわからない ―\nだから否定しない。理由がはっきり言葉になっています。", { x: 0.7, y: 4.65, w: 11.9, h: 1.1, fontSize: 22, color: GRAY, isTextBox: true, margin: 0 });
  point(s, "大事なのは「変わったか」ではなく「理由を言えるか」");
  s.addNotes("「特になし」「変わらなかった」と書いた人も、ここで自分の理由を考え直すきっかけに。");

  // ================= 3. 命と大切なもの =================
  const S3 = "3 命と大切なもの";
  divider(S3, "PART 3", "命と、大切なもの", "ぶつかったとき、どう判断する？", ICONS.heart);

  s = content(S3, "PART 3 命と、大切なもの", "22人の判断は、大きく3つに分かれた");
  const groups = [
    ["10人", "命を最優先", "「命がないと、大切にしていることさえ守れない」", ROSE, ROSE_T],
    ["8人", "命を守りつつ、本人の思いも", "「まず対話する」「治ってから安全に」", TEAL, TEAL_T],
    ["4人", "本人の価値観を尊重", "「その人が今後どうしたいかを聞き、合わせる」", PLUM, "EDE4EC"],
  ];
  groups.forEach(([n, h, b, c, t], i) => {
    const x = 0.7 + i * 4.05;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.95, w: 3.8, h: 4.85, rectRadius: 0.15, fill: { color: t }, line: { type: "none" }, objectName: "group-card" });
    s.addText(n, { x: x + 0.3, y: 2.15, w: 3.2, h: 1.3, fontSize: 66, bold: true, color: c, isTextBox: true, margin: 0 });
    s.addText(h, { x: x + 0.3, y: 3.5, w: 3.2, h: 1.1, fontSize: 26, bold: true, color: INK, valign: "top", isTextBox: true, margin: 0 });
    s.addText(b, { x: x + 0.3, y: 4.7, w: 3.2, h: 1.9, fontSize: 20, color: GRAY, valign: "top", isTextBox: true, margin: 0 });
  });
  s.addNotes("記述をもとに教員が3つに分類したもの（無回答5人）。どれも間違いではない。それぞれの理由に注目する。");

  s = content(S3, "PART 3 命と、大切なもの", "2つの声を並べてみると");
  quote(s, 0.7, 1.95, 5.8, 3.8, "命は1番大事。命を失うと、大事なものなど何にも変えられない。命がないと、その人が大切にしていることさえ守ることができないから。", 22);
  quote(s, 6.8, 1.95, 5.8, 3.8, "命や体を大切にしている人もいれば、命よりも大切なものがある人もいる。自分が自分らしく生きていくための選択は、患者によって違うから。", 22, TEAL_T, TEAL);
  point(s, "「命を守る」の“命”って、何のこと？");
  s.addNotes("どちらも「命を大切にしている」点では同じ。前者は生命（いのち）を、後者はその人の生き方・人生を見ている。英語ではどちらも life。看護ではQOL（生活の質）という言葉で語られることを紹介してもよい。");

  s = content(S3, "PART 3 命と、大切なもの", "「どちらか」を選ぶ前にできること");
  quote(s, 0.7, 1.95, 11.9, 2.0, "どちらか一方だけを切り捨てることはできない。まずは本人が「何にそれほどこだわっているのか」を深く対話し、命を守りながらも大切なものを残せる代替案を必死に探す。", 22);
  quote(s, 0.7, 4.15, 11.9, 1.6, "してもしなくても死にゆく未来なら、本人が悔いなく旅立てるように、やれることはさせてあげたい。", 22, TEAL_T, TEAL);
  point(s, "答えは「対話」と「その人の状況」によって変わる");
  s.addNotes("二者択一ではなく第三の道を探す姿勢。また、同じ選択でも状況（回復が見込めるか、終末期か）によって判断が変わる、という視点。");

  question(S3, "グループで考えよう ― 4人で 5分", "「シリコンだけは抜かないで」\n手術の前に、どんな対話が\nできていたらよかった？",
    "誰が・いつ・何を聞いておく？　あなたが看護師なら？",
    "※配慮：ドラマでは患者が自ら命を絶つ場面がある。体調や気持ちがしんどい生徒がいないか表情を見ながら進め、必要なら個別に声をかける。\n\n発問のねらい：危篤状態の緊急時には本人の同意をとる時間がない。だからこそ『事前に』本人の大切なものを知っておくことが大事。豊胸手術を受ける前の診察で、腫瘍の有無を確認していたら？ 本人が胸をどれほど大切にしているかを医療者が知っていたら？ → 次のインフォームド・コンセントへ。");

  // ================= 4. IC =================
  const S4 = "4 インフォームド・コンセント";
  divider(S4, "PART 4", "インフォームド・コンセント", "「同意書にサイン」のことではない", ICONS.hands);

  s = content(S4, "PART 4 インフォームド・コンセント", "インフォームド・コンセント ＝ 4つのステップ");
  const steps = [["1", "説明を受ける", "目的・リスク・\nほかの方法"], ["2", "理解する", "わからないことを\n質問できる"], ["3", "納得する", "自分の大切な\nものと照らす"], ["4", "自分で選ぶ", "「同意しない」も\n選べる"]];
  steps.forEach(([n, h, b], i) => {
    const x = 0.7 + i * 3.05;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 2.0, w: 2.8, h: 3.75, rectRadius: 0.15, fill: { color: i === 3 ? ROSE_T : TEAL_T }, line: { type: "none" }, objectName: "step-card" });
    s.addShape(pres.shapes.OVAL, { x: x + 0.95, y: 2.25, w: 0.9, h: 0.9, fill: { color: i === 3 ? ROSE : TEAL }, line: { type: "none" }, objectName: "step-num" });
    s.addText(n, { x: x + 0.95, y: 2.25, w: 0.9, h: 0.9, fontSize: 32, bold: true, color: WHITE, align: "center", valign: "middle", isTextBox: true, margin: 0 });
    s.addText(h, { x: x + 0.15, y: 3.35, w: 2.5, h: 0.7, fontSize: 28, bold: true, color: INK, align: "center", isTextBox: true, margin: 0 });
    s.addText(b, { x: x + 0.1, y: 4.15, w: 2.6, h: 1.3, fontSize: 20, color: GRAY, align: "center", valign: "top", isTextBox: true, margin: 0 });
  });
  point(s, "サインはゴールではなく、この4つがそろった「しるし」");
  s.addNotes("インフォームド・コンセント（説明と同意）。医療者が説明し、患者が理解・納得したうえで自分の意思で決めること。「同意しない（断る）」権利も含まれる。");

  s = content(S4, "PART 4 インフォームド・コンセント", "クラスの中に、すでに答えのヒントがあった");
  quote(s, 0.7, 1.95, 11.9, 2.0, "患者の命の優先順位や、何が嬉しいことなのか ― 価値観を知らないと、知らないうちに患者を追い詰めてしまう。コミュニケーションをとって患者について知っていく努力が重要。", 22);
  quote(s, 0.7, 4.15, 11.9, 1.6, "最後は患者の意志が大切。これまで大切にしてきたことや考えを、しっかり聞いて受け止めて尊重すること。", 22, TEAL_T, TEAL);
  point(s, "医療者は「情報」を、患者は「大切なもの」を ― 対話は双方向");
  s.addNotes("インフォームド・コンセントは、医療者が一方的に説明するだけではない。患者の価値観を聴くことが含まれる。生徒自身の言葉で、すでにその本質に迫っていることを伝えて称える。");

  s = content(S4, "PART 4 インフォームド・コンセント", "看護師だからできること");
  circleIcon(s, ICONS.stetho, 0.9, 2.1, 1.6, TEAL);
  s.addText([
    { text: "医師の説明のあと、患者さんのいちばん近くにいるのは看護師。", options: { bold: true, color: INK, breakLine: true } },
    { text: " ", options: { fontSize: 10, breakLine: true } },
    { text: "「本当はよくわかっていない」", options: { color: TEAL, bold: true, breakLine: true } },
    { text: "「本当はまだ迷っている」", options: { color: TEAL, bold: true, breakLine: true } },
    { text: "「誰にも言えない大切なものがある」", options: { color: TEAL, bold: true, breakLine: true } },
    { text: " ", options: { fontSize: 10, breakLine: true } },
    { text: "そんな声に気づき、医師や家族につなぐことができる。", options: { color: INK } },
  ], { x: 3.0, y: 1.95, w: 9.6, h: 3.9, fontSize: 24, lineSpacingMultiple: 1.2, valign: "top", isTextBox: true, margin: 0 });
  point(s, "看護師は、患者さんの「選ぶ」を支える人");
  s.addNotes("看護職には、患者の意思決定を支援し、患者の思いを代弁する（アドボカシー）役割がある。看護科の生徒として、ドラマの中に看護師がいたら何ができたかを想像させてもよい。");

  // ================= 5. 社会と私 =================
  const S5 = "5 社会と私";
  divider(S5, "PART 5", "社会と、私", "そのコンプレックスはどこから来た？", ICONS.users);

  s = content(S5, "PART 5 社会と、私", "「直せるなら直したい」と答えた人");
  s.addText([{ text: "19", options: { fontSize: 120, bold: true, color: ROSE } }, { text: " / 25人", options: { fontSize: 36, bold: true, color: GRAY } }],
    { x: 0.7, y: 1.9, w: 5.4, h: 2.4, valign: "middle", isTextBox: true, margin: 0 });
  s.addText("その理由として書かれた言葉", { x: 0.7, y: 4.4, w: 5.4, h: 0.5, fontSize: 20, color: GRAY, isTextBox: true, margin: 0 });
  s.addText("自信が持てる・自分を好きになれる・堂々と笑える・コンプレックスがなくなる", { x: 0.7, y: 4.9, w: 5.4, h: 1.4, fontSize: 22, bold: true, color: PLUM, valign: "top", isTextBox: true, margin: 0 });
  quote(s, 6.5, 1.95, 6.1, 2.15, "差別の目で見られないで、周りの人達と溶け込むことができるため。", 22);
  quote(s, 6.5, 4.3, 6.1, 2.15, "努力でなんとかなるものなら、一旦自分で頑張ってみることが大事だと思う。", 22, TEAL_T, TEAL);
  s.addNotes("具体的な部位には触れない（個人が特定されないよう配慮）。「自分のため」の理由が多い一方で、「周りの目」を理由にした声もある。");

  s = content(S5, "PART 5 社会と、私", "「理想の顔」の基準は、誰が決めた？");
  quote(s, 0.7, 1.95, 11.9, 1.25, "SNSで知らない人の顔を簡単に見られるから、勝手に比べてしまい、美容医療に手を出す人が増えている。", 21);
  quote(s, 0.7, 3.4, 11.9, 1.25, "ルッキズムが強くなって、整形が「努力」みたいになっている。でも、気にしているなら心の健康のためにするのも良いのかも。", 21);
  quote(s, 0.7, 4.85, 11.9, 1.0, "他人に愛されるより、自分が自分を愛せることを求める人が多いんだな。", 21, TEAL_T, TEAL);
  point(s, "「自分のため」の理想にも、社会の目が入りこんでいない？");
  s.addNotes("ルッキズム（外見至上主義）の視点。自分で選んだつもりの「理想」も、SNSや広告、周りの言葉の影響を受けている。それでも本人の苦しみは本物。両方を同時に見る力が大切。");

  question(S5, "考えてみよう ― ひとりで 1分 → となりと共有", "クラスの誰かがこう書いていました。\n「自分も周りも、人の体を大切にする\n考え方が根本にあったらいいな」",
    "そんな社会にするために、あなたにできることは？",
    "「周りが人の容姿を自分の価値観でとやかく言う権利はない」という言葉も同じ生徒から。小さな行動（見た目をいじらない、SNSでの言葉を選ぶ等）から考えさせる。");

  // ================= まとめ =================
  const S6 = "まとめ";
  s = content(S6, "まとめ", "今日、持ち帰ってほしい3つの問い");
  const takeaways = [
    ["1", "「治す」と「変える」の境目は、誰が、何を基準に決める？"],
    ["2", "「自己責任」が成り立つために、何が必要？"],
    ["3", "命と大切なものがぶつかる前に、どんな対話ができる？"],
  ];
  takeaways.forEach(([n, t], i) => {
    const y = 2.0 + i * 1.35;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.7, y, w: 11.9, h: 1.1, rectRadius: 0.15, fill: { color: i === 2 ? TEAL_T : ROSE_T }, line: { type: "none" }, objectName: "take-card" });
    s.addShape(pres.shapes.OVAL, { x: 0.95, y: y + 0.15, w: 0.8, h: 0.8, fill: { color: i === 2 ? TEAL : ROSE }, line: { type: "none" }, objectName: "take-num" });
    s.addText(n, { x: 0.95, y: y + 0.15, w: 0.8, h: 0.8, fontSize: 30, bold: true, color: WHITE, align: "center", valign: "middle", isTextBox: true, margin: 0 });
    s.addText(t, { x: 2.0, y, w: 10.4, h: 1.1, fontSize: 26, bold: true, color: INK, valign: "middle", isTextBox: true, margin: 0 });
  });
  s.addText("どれも「正解」はひとつではありません。看護の現場で、何度も出会う問いです。", { x: 0.7, y: 6.1, w: 11.9, h: 0.6, fontSize: 20, color: GRAY, isTextBox: true, margin: 0 });
  s.addNotes("3つの問いはPART1〜3に対応。最後に振り返りシートに「今日、友だちの意見で一番揺さぶられたもの」を書かせる。");

  s = add("TITLE_DARK", S6);
  s.addText("次回へ", { placeholder: "title" });
  s.addText("「続きが見たい」「2話を見たい」という声が多数。\n「実際の医療現場で、医師や看護師がどう向き合っているのか知りたい」という声も。\n\n振り返り：今日、友だちの意見で一番揺さぶられたのは？", { placeholder: "body" });
  s.addNotes("次回の予告。振り返りの問いを書かせて終了。");

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
