// 書類生成の共通部品（A4縦・明朝体）
const {
  Document, Paragraph, TextRun, Table, TableRow, TableCell, WidthType,
  AlignmentType, BorderStyle, VerticalAlign, ShadingType, PageOrientation,
  Footer, PageNumber, TableLayoutType,
} = require('docx');

const FONT = { ascii: 'Century', hAnsi: 'Century', eastAsia: 'ＭＳ 明朝' };
const GOTHIC = { ascii: 'ＭＳ ゴシック', hAnsi: 'ＭＳ ゴシック', eastAsia: 'ＭＳ ゴシック' };

const A4 = { width: 11906, height: 16838 };
const MARGIN = 1134; // 20mm
const CONTENT_W = A4.width - MARGIN * 2; // 9638
const CONTENT_W_LAND = A4.height - MARGIN * 2; // 14570

// 文字列中の **太字** を TextRun に分解する
function runs(text, opt = {}) {
  const parts = String(text).split(/(\*\*[^*]+\*\*)/);
  return parts.filter((s) => s !== '').map((s) => {
    const bold = s.startsWith('**');
    return new TextRun({
      text: bold ? s.slice(2, -2) : s,
      bold: bold || opt.bold,
      size: opt.size,
      font: opt.gothic ? GOTHIC : undefined,
      underline: opt.underline ? {} : undefined,
    });
  });
}

function p(text, opt = {}) {
  return new Paragraph({
    children: runs(text, opt),
    alignment: opt.align,
    indent: opt.indent,
    spacing: { before: opt.before ?? 0, after: opt.after ?? 0, line: opt.line },
    keepNext: opt.keepNext,
    border: opt.border,
  });
}

const blank = (n = 1) => Array.from({ length: n }, () => p(''));
const right = (t, o = {}) => p(t, { ...o, align: AlignmentType.RIGHT });
const center = (t, o = {}) => p(t, { ...o, align: AlignmentType.CENTER });
const title = (t) => center(t, { size: 32, bold: true, gothic: true, before: 120, after: 240 });
const formNo = (n) => right(`（様式第${n}号）`);

// ぶら下げインデント付きの項目（「(1)」「１」など）
function item(label, text, depth = 0, opt = {}) {
  const base = 210 * (depth + 1);
  const hang = opt.hang ?? 420;
  return new Paragraph({
    children: [...runs(label, opt), ...runs(text, opt)],
    indent: { left: base + hang - 210, hanging: hang },
    spacing: { before: 0, after: opt.after ?? 0 },
  });
}

const B = { style: BorderStyle.SINGLE, size: 4, color: '000000' };
const B_THICK = { style: BorderStyle.SINGLE, size: 12, color: '000000' };
const NO_B = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };

// セル: c(内容, 幅) — 内容は文字列 / 文字列配列 / Paragraph配列
function c(content, w, opt = {}) {
  const lines = Array.isArray(content) ? content : [content];
  const children = lines.map((l) =>
    l instanceof Paragraph ? l : p(l, { align: opt.align, size: opt.size, bold: opt.bold }));
  return new TableCell({
    children,
    width: { size: w, type: WidthType.DXA },
    columnSpan: opt.span,
    rowSpan: opt.rowSpan,
    verticalAlign: opt.valign ?? VerticalAlign.CENTER,
    shading: opt.shade ? { type: ShadingType.CLEAR, color: 'auto', fill: opt.shade } : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    borders: opt.borders,
  });
}
// 見出しセル（網掛け・中央揃え）
const h = (t, w, opt = {}) => c(t, w, { shade: 'EDEDED', align: AlignmentType.CENTER, ...opt });

function table(widths, rows, opt = {}) {
  return new Table({
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    columnWidths: widths,
    layout: TableLayoutType.FIXED,
    alignment: opt.align ?? AlignmentType.CENTER,
    rows: rows.map((cells, i) => new TableRow({
      children: cells,
      height: opt.heights?.[i] ?? opt.height ? { value: opt.heights?.[i] ?? opt.height, rule: 'atLeast' } : undefined,
      cantSplit: true,
      tableHeader: opt.header && i === 0,
    })),
    borders: opt.borders ?? { top: B, bottom: B, left: B, right: B, insideHorizontal: B, insideVertical: B },
  });
}

// 囲み枠（1セルの表）
function box(children, opt = {}) {
  const bd = opt.thick ? B_THICK : B;
  return table([CONTENT_W], [[c(children, CONTENT_W, { valign: VerticalAlign.TOP, shade: opt.shade })]], {
    borders: { top: bd, bottom: bd, left: bd, right: bd, insideHorizontal: bd, insideVertical: bd },
  });
}

function doc(sections) {
  return new Document({
    styles: { default: { document: { run: { font: FONT, size: 21 } } } },
    sections: sections.map((s) => ({
      properties: {
        page: {
          size: s.landscape
            ? { width: A4.width, height: A4.height, orientation: PageOrientation.LANDSCAPE }
            : A4,
          margin: { top: s.marginTop ?? MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN },
        },
      },
      footers: s.pageNumber ? {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ children: ['－ ', PageNumber.CURRENT, ' －'], size: 18 })],
          })],
        }),
      } : undefined,
      children: s.children,
    })),
  });
}

module.exports = {
  p, blank, right, center, title, formNo, item, c, h, table, box, doc, runs,
  B, B_THICK, NO_B, CONTENT_W, CONTENT_W_LAND, AlignmentType, VerticalAlign, BorderStyle, Paragraph,
};
