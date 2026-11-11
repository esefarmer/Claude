#!/usr/bin/env node
/*
 * 簡易マークアップ（content/*.txt）から Word 文書を生成する。
 *   node build.js <content.txt> <out.docx>
 *
 * 記法:
 *   @header 文字列        ヘッダー右肩の文書名
 *   # 見出し              文書タイトル（中央）
 *   ## 見出し             章・大見出し
 *   ### 見出し            小見出し
 *   （見出し）            条文見出し
 *   第１条　本文          条（ぶら下げ）  /  ２　本文  項  /  (1)　本文  号  /  イ　本文  細目
 *   - 本文 /   - 本文     箇条書き（2階層）
 *   > 本文                注記（網掛け）
 *   @right / @center 本文 右寄せ・中央寄せ
 *   @pb                   改ページ
 *   @sp                   空行
 *   @cols 20,80           次の表の列幅（%）
 *   | a | b |             表。2行目が |---| なら1行目は見出し行。セル内 <br> で改行。
 *                         セル先頭 ! で項目名セル（網掛け）。セル内容 ^ で左のセルと結合。
 *   **太字**              行内の太字
 */
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType,
  BorderStyle, ShadingType, Header, Footer, PageNumber, LevelFormat, HeadingLevel, VerticalAlign,
  PageBreak,
} = require("docx");

const [src, out] = process.argv.slice(2);
const lines = fs.readFileSync(src, "utf8").split(/\r?\n/);

const MIN = "ＭＳ 明朝", GOT = "ＭＳ ゴシック";
const font = (f) => ({ ascii: f, hAnsi: f, eastAsia: f, cs: f });
const SZ = 21;                // 10.5pt
const CH = 210;               // 1字分（twip）
const PAGE_W = 11906, MARGIN = 1247;
const CONTENT_W = PAGE_W - MARGIN * 2;

function runs(text, opt = {}) {
  const parts = text.split(/(\*\*[^*]+\*\*)/).filter((s) => s !== "");
  return parts.map((s) => {
    const bold = /^\*\*.+\*\*$/.test(s);
    return new TextRun({
      text: bold ? s.slice(2, -2) : s,
      bold: bold || opt.bold,
      font: font(bold || opt.gothic ? GOT : MIN),
      size: opt.size || SZ,
    });
  });
}

function para(text, o = {}) {
  return new Paragraph({
    children: runs(text, o),
    alignment: o.align,
    indent: o.indent,
    spacing: { before: o.before || 0, after: o.after || 0, line: o.line || 340 },
    heading: o.heading,
    keepNext: o.keepNext,
    shading: o.shade ? { type: ShadingType.CLEAR, fill: "F2F2F2", color: "auto" } : undefined,
    numbering: o.numbering,
  });
}

const border = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
const borders = { top: border, bottom: border, left: border, right: border };

function cellParas(text, o = {}) {
  return text.split(/<br>/).map((t) => new Paragraph({
    children: runs(t, { size: 19, bold: o.bold, gothic: o.gothic }),
    alignment: o.align,
    spacing: { before: 0, after: 0, line: 290 },
  }));
}

function buildTable(rows, cols) {
  const n = Math.max(...rows.map((r) => r.cells.length));
  let pct = cols && cols.length === n ? cols : Array(n).fill(100 / n);
  const tot = pct.reduce((a, b) => a + b, 0);
  const widths = pct.map((p) => Math.floor((CONTENT_W * p) / tot));
  widths[n - 1] += CONTENT_W - widths.reduce((a, b) => a + b, 0);
  const trs = rows.map((r) => {
    const cells = [];
    for (let i = 0; i < n; i++) {
      let t = (r.cells[i] ?? "").trim();
      if (t === "^") continue;
      let span = 1, w = widths[i];
      while (i + span < n && (r.cells[i + span] ?? "").trim() === "^") { w += widths[i + span]; span++; }
      const label = t.startsWith("!");
      if (label) t = t.slice(1).trim();
      const head = r.header;
      cells.push(new TableCell({
        children: cellParas(t, { bold: head, gothic: head || label, align: head ? AlignmentType.CENTER : undefined }),
        width: { size: w, type: WidthType.DXA },
        columnSpan: span > 1 ? span : undefined,
        borders,
        verticalAlign: head || label ? VerticalAlign.CENTER : VerticalAlign.TOP,
        shading: head || label ? { type: ShadingType.CLEAR, fill: "E7E6E6", color: "auto" } : undefined,
        margins: { top: 50, bottom: 50, left: 90, right: 90 },
      }));
    }
    const long = r.cells.join("").length > 260;   // 長い行（新旧対照表等）はページをまたいでよい
    return new TableRow({ children: cells, tableHeader: r.header, cantSplit: !long });
  });
  return new Table({ width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: widths, rows: trs });
}

const body = [];
let headerText = "";
let cols = null;
let i = 0;
while (i < lines.length) {
  let l = lines[i];
  if (/^\|/.test(l)) {
    const rows = [];
    while (i < lines.length && /^\|/.test(lines[i])) {
      const raw = lines[i].replace(/^\|/, "").replace(/\|\s*$/, "").split("|");
      if (raw.every((c) => /^\s*:?-+:?\s*$/.test(c))) {
        if (rows.length) rows[rows.length - 1].header = true;
      } else rows.push({ cells: raw, header: false });
      i++;
    }
    body.push(buildTable(rows, cols));
    body.push(para("", { line: 160 }));
    cols = null;
    continue;
  }
  i++;
  if (l.trim() === "") continue;
  let m;
  if ((m = l.match(/^@header (.*)/))) { headerText = m[1]; continue; }
  if ((m = l.match(/^@cols (.*)/))) { cols = m[1].split(",").map(Number); continue; }
  if (l.trim() === "@pb") { body.push(new Paragraph({ children: [new PageBreak()] })); continue; }
  if (l.trim() === "@sp") { body.push(para("")); continue; }
  if ((m = l.match(/^@right (.*)/))) {
    // 末尾が空白の行（署名欄など）は、記入欄が残るよう左インデントで右側に置く
    if (/　{3,}$/.test(m[1])) body.push(para(m[1].replace(/　+$/, ""), { indent: { left: Math.round(CONTENT_W * 0.5) }, before: 120 }));
    else body.push(para(m[1], { align: AlignmentType.RIGHT }));
    continue;
  }
  if ((m = l.match(/^@center (.*)/))) { body.push(para(m[1], { align: AlignmentType.CENTER })); continue; }
  if ((m = l.match(/^# (.*)/))) {
    body.push(para(m[1], { align: AlignmentType.CENTER, bold: true, gothic: true, size: 30, before: 120, after: 240, heading: HeadingLevel.TITLE, line: 440 }));
    continue;
  }
  if ((m = l.match(/^## (.*)/))) {
    body.push(para(m[1], { bold: true, gothic: true, size: 24, before: 300, after: 120, heading: HeadingLevel.HEADING_1, keepNext: true, line: 380 }));
    continue;
  }
  if ((m = l.match(/^### (.*)/))) {
    body.push(para(m[1], { bold: true, gothic: true, size: 22, before: 200, after: 60, heading: HeadingLevel.HEADING_2, keepNext: true }));
    continue;
  }
  if ((m = l.match(/^> (.*)/))) {
    // 注記：網掛けの1セル表（段落罫線は docx ライブラリの出力順がスキーマに合わないため使わない）
    const g = { style: BorderStyle.SINGLE, size: 4, color: "808080" };
    body.push(new Table({
      width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [CONTENT_W],
      rows: [new TableRow({ children: [new TableCell({
        children: [new Paragraph({ children: runs(m[1], { size: 19 }), spacing: { before: 0, after: 0, line: 300 } })],
        width: { size: CONTENT_W, type: WidthType.DXA },
        borders: { top: g, bottom: g, left: g, right: g },
        shading: { type: ShadingType.CLEAR, fill: "F2F2F2", color: "auto" },
        margins: { top: 70, bottom: 70, left: 120, right: 120 },
      })] })],
    }));
    body.push(para("", { line: 160 }));
    continue;
  }
  if ((m = l.match(/^( *)- (.*)/))) {
    const lv = m[1].length >= 2 ? 1 : 0;
    body.push(para(m[2], { numbering: { reference: "bul", level: lv } }));
    continue;
  }
  if (/^（[^）]*）$/.test(l.trim()) && l.trim().length <= 34) {
    body.push(para(l.trim(), { indent: { left: CH }, before: 140, keepNext: true }));
    continue;
  }
  if (/^第[０-９0-9]+条(の[０-９0-9]+)?　/.test(l) || /^[０-９]+　/.test(l)) {
    body.push(para(l, { indent: { left: CH, hanging: CH } }));
    continue;
  }
  if (/^[(（][0-9０-９]+[)）]/.test(l)) {
    body.push(para(l, { indent: { left: CH * 3, hanging: CH * 2 } }));
    continue;
  }
  if (/^[イロハニホヘトアウエオ①-⑳]　/.test(l)) {
    body.push(para(l, { indent: { left: CH * 4, hanging: CH } }));
    continue;
  }
  if (/^[□☐]/.test(l)) {
    body.push(para(l, { indent: { left: CH * 2, hanging: CH } }));
    continue;
  }
  body.push(para(l));
}

const doc = new Document({
  creator: "学校法人明珠学園",
  title: headerText,
  styles: {
    default: { document: { run: { font: font(MIN), size: SZ } } },
    paragraphStyles: [
      { id: "Title", name: "Title", basedOn: "Normal", run: { font: font(GOT), size: 30, bold: true } },
      { id: "Heading1", name: "heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: font(GOT), size: 24, bold: true }, paragraph: { outlineLevel: 0 } },
      { id: "Heading2", name: "heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: font(GOT), size: 22, bold: true }, paragraph: { outlineLevel: 1 } },
    ],
  },
  numbering: {
    config: [{
      reference: "bul",
      levels: [
        { level: 0, format: LevelFormat.BULLET, text: "・", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: CH * 2, hanging: CH } }, run: { font: font(MIN) } } },
        { level: 1, format: LevelFormat.BULLET, text: "－", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: CH * 4, hanging: CH } }, run: { font: font(MIN) } } },
      ],
    }],
  },
  sections: [{
    properties: { page: { size: { width: PAGE_W, height: 16838 }, margin: { top: 1247, bottom: 1134, left: MARGIN, right: MARGIN } } },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT,
      children: [new TextRun({ text: headerText, font: font(GOT), size: 16, color: "595959" })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
      children: [new TextRun({ children: [PageNumber.CURRENT], font: font(MIN), size: 18 })] })] }) },
    children: body,
  }],
});

Packer.toBuffer(doc).then((b) => { fs.writeFileSync(out, b); console.log("wrote", out); });
