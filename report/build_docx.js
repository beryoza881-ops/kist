const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, ShadingType, AlignmentType, BorderStyle, HeadingLevel,
  VerticalAlign, PageOrientation
} = require('docx');

const FONT = '맑은 고딕';
const SRC = '/home/user/kist/report/해외_자율실험실_현황.md';
const OUT = '/home/user/kist/report/해외_자율실험실_현황.docx';

const CONTENT_W = 9746;            // A4 11906 - 2*1080
const COLS = [1500, 1150, 4300, 2796];
if (COLS.reduce((a, b) => a + b, 0) !== CONTENT_W) throw new Error('column widths must sum to ' + CONTENT_W);

const lines = fs.readFileSync(SRC, 'utf8').split('\n');

// --- markdown inline: strip nothing fancy, source has no bold/italic inside cells
function runs(text, opts = {}) {
  return [new TextRun({ text, font: FONT, size: opts.size || 19, bold: !!opts.bold, color: opts.color })];
}

function cell(text, { header = false, width, align } = {}) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    verticalAlign: VerticalAlign.CENTER,
    shading: header ? { type: ShadingType.CLEAR, fill: 'D9E2F3', color: 'auto' } : undefined,
    margins: { top: 70, bottom: 70, left: 90, right: 90 },
    children: [new Paragraph({
      alignment: align || (header ? AlignmentType.CENTER : AlignmentType.LEFT),
      spacing: { line: 264 },
      children: runs(text, { bold: header, size: header ? 19 : 18 }),
    })],
  });
}

function splitRow(line) {
  return line.replace(/^\||\|$/g, '').split('|').map(s => s.trim());
}

const children = [];
let i = 0;

while (i < lines.length) {
  const line = lines[i];

  // table block
  if (line.trim().startsWith('|')) {
    const block = [];
    while (i < lines.length && lines[i].trim().startsWith('|')) { block.push(lines[i]); i++; }
    const head = splitRow(block[0]);
    const body = block.slice(2).map(splitRow);
    const rows = [
      new TableRow({
        tableHeader: true,
        children: head.map((t, c) => cell(t, { header: true, width: COLS[c] })),
      }),
      ...body.map(r => new TableRow({
        children: r.map((t, c) => cell(t, {
          width: COLS[c],
          align: c < 2 ? AlignmentType.CENTER : AlignmentType.LEFT,
        })),
      })),
    ];
    children.push(new Table({
      columnWidths: COLS,
      width: { size: CONTENT_W, type: WidthType.DXA },
      rows,
    }));
    children.push(new Paragraph({ spacing: { after: 260 }, children: [] }));
    continue;
  }

  const t = line.trim();

  if (t.startsWith('# ')) {
    children.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 400 },
      children: runs(t.slice(2), { bold: true, size: 32 }),
    }));
  } else if (t.startsWith('□')) {
    children.push(new Paragraph({
      spacing: { before: 320, after: 160 },
      children: runs(t, { bold: true, size: 24 }),
    }));
  } else if (t.startsWith('ㅇ')) {
    children.push(new Paragraph({
      indent: { left: 220 },
      spacing: { before: 140, after: 90 },
      children: runs(t, { bold: true, size: 21 }),
    }));
  } else if (t.startsWith('- ')) {
    children.push(new Paragraph({
      indent: { left: 560, hanging: 180 },
      spacing: { after: 110, line: 288 },
      children: runs('- ' + t.slice(2), { size: 20 }),
    }));
  } else if (t.length) {
    children.push(new Paragraph({ spacing: { after: 110 }, children: runs(t, { size: 20 }) }));
  }
  i++;
}

const doc = new Document({
  styles: { default: { document: { run: { font: FONT, size: 20 } } } },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838, orientation: PageOrientation.PORTRAIT },
        margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 },
      },
    },
    children,
  }],
});

Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log('wrote', OUT, b.length, 'bytes'); });
