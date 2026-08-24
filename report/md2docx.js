// 간이 마크다운 → docx 변환기 (제목/표/목록/굵게/링크/인용 지원)
const path = require('path');
const SCRATCH = '/tmp/claude-0/-home-user-kist/f7349fbc-9d61-5776-85de-581602fe2d75/scratchpad';
const docx = require(path.join(SCRATCH, 'node_modules', 'docx'));
const { Document, Packer, Paragraph, TextRun, ExternalHyperlink, Table, TableRow, TableCell,
  WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign, convertMillimetersToTwip } = docx;
const fs = require('fs');

const KO = { ascii: '맑은 고딕', eastAsia: '맑은 고딕', hAnsi: '맑은 고딕', cs: '맑은 고딕' };
const NAVY = '1D3567', LINE = 'B9C2D4', HEADFILL = '1D3567', ALTFILL = 'F7F9FC';
const PAGE_W = convertMillimetersToTwip(210), MARGIN = convertMillimetersToTwip(14);
const CONTENT_W = PAGE_W - 2 * MARGIN;

// 인라인: **굵게**, `코드`, [텍스트](url)
function inline(text, { size = 16, bold = false, color = '222222' } = {}) {
  const out = [];
  const re = /\[([^\]]+)\]\(([^)]+)\)|\*\*([^*]+)\*\*|`([^`]+)`/g;
  let last = 0, m;
  while ((m = re.exec(text)) !== null) {
    if (m.index > last) out.push(new TextRun({ text: text.slice(last, m.index), size, bold, color, font: KO }));
    if (m[1] !== undefined) {
      const label = m[1].replace(/\*\*/g, '');
      out.push(new ExternalHyperlink({ link: m[2], children: [new TextRun({ text: label, size, bold: bold || /\*\*/.test(m[1]), font: KO, color: '1B54B5', underline: {} })] }));
    } else if (m[3] !== undefined) {
      out.push(new TextRun({ text: m[3], size, bold: true, color, font: KO }));
    } else {
      out.push(new TextRun({ text: m[4], size: size - 1, color: '444444', font: { ascii: 'Consolas', eastAsia: '맑은 고딕', hAnsi: 'Consolas' } }));
    }
    last = re.lastIndex;
  }
  if (last < text.length) out.push(new TextRun({ text: text.slice(last), size, bold, color, font: KO }));
  return out;
}
const splitRow = (l) => l.replace(/^\||\|$/g, '').split('|').map((s) => s.trim());

function convert(mdPath, outPath, titleOverride) {
  const lines = fs.readFileSync(mdPath, 'utf8').split('\n');
  const c = [];
  let i = 0, titleDone = false;
  while (i < lines.length) {
    const l = lines[i];
    // 표
    if (/^\s*\|/.test(l) && i + 1 < lines.length && /^\s*\|[\s:|-]+\|\s*$/.test(lines[i + 1])) {
      const header = splitRow(l); const body = [];
      i += 2;
      while (i < lines.length && /^\s*\|/.test(lines[i])) { body.push(splitRow(lines[i])); i++; }
      const n = header.length;
      const w = Math.floor(CONTENT_W / n);
      const widths = Array(n).fill(w); widths[n - 1] = CONTENT_W - w * (n - 1);
      const mk = (t, idx, opt) => new TableCell({ width: { size: widths[idx], type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER,
        shading: opt.fill ? { type: ShadingType.CLEAR, fill: opt.fill, color: 'auto' } : undefined,
        margins: { top: 75, bottom: 75, left: 75, right: 75 },
        children: [new Paragraph({ alignment: opt.align, spacing: { before: 10, after: 10, line: 200, lineRule: 'auto' },
          children: inline(t, { size: 14, bold: opt.bold, color: opt.color || '222222' }) })] });
      c.push(new Table({ columnWidths: widths, width: { size: CONTENT_W, type: WidthType.DXA },
        borders: { top: { style: BorderStyle.SINGLE, size: 3, color: LINE }, bottom: { style: BorderStyle.SINGLE, size: 3, color: LINE },
          left: { style: BorderStyle.SINGLE, size: 3, color: LINE }, right: { style: BorderStyle.SINGLE, size: 3, color: LINE },
          insideHorizontal: { style: BorderStyle.SINGLE, size: 3, color: LINE }, insideVertical: { style: BorderStyle.SINGLE, size: 3, color: LINE } },
        rows: [new TableRow({ tableHeader: true, children: header.map((t, k) => mk(t, k, { fill: HEADFILL, bold: true, color: 'FFFFFF', align: AlignmentType.CENTER })) }),
          ...body.map((r, ri) => new TableRow({ cantSplit: true, children: Array.from({ length: n }, (_, k) =>
            mk(r[k] || '', k, { fill: ri % 2 === 1 ? ALTFILL : undefined, bold: k === 0 })) }))] }));
      c.push(new Paragraph({ spacing: { after: 60 }, children: [] }));
      continue;
    }
    const hm = l.match(/^(#{1,4})\s+(.*)$/);
    if (hm) {
      const lvl = hm[1].length; const txt = hm[2].trim();
      if (lvl === 1 && !titleDone) {
        titleDone = true;
        c.push(new Paragraph({ spacing: { after: 40, line: 240, lineRule: 'auto' },
          children: inline(txt, { size: 25, bold: true, color: NAVY }) }));
        c.push(new Paragraph({ spacing: { after: 130 }, border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: NAVY, space: 4 } }, children: [] }));
      } else {
        c.push(new Paragraph({ spacing: { before: lvl <= 2 ? 300 : 200, after: 70 },
          children: inline(txt, { size: lvl <= 2 ? 18 : 16, bold: true, color: NAVY }) }));
      }
      i++; continue;
    }
    if (/^\s*---\s*$/.test(l)) { c.push(new Paragraph({ spacing: { before: 100, after: 100 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: LINE, space: 2 } }, children: [] })); i++; continue; }
    if (/^\s*[-*]\s+/.test(l)) {
      c.push(new Paragraph({ spacing: { before: 26, after: 26, line: 212, lineRule: 'auto' }, indent: { left: 240, hanging: 170 },
        children: [new TextRun({ text: '▪  ', size: 15, color: NAVY, font: KO }), ...inline(l.replace(/^\s*[-*]\s+/, ''))] }));
      i++; continue;
    }
    if (/^\s*>\s?/.test(l)) {
      const buf = [];
      while (i < lines.length && /^\s*>\s?/.test(lines[i])) { buf.push(lines[i].replace(/^\s*>\s?/, '')); i++; }
      c.push(new Table({ columnWidths: [CONTENT_W], width: { size: CONTENT_W, type: WidthType.DXA },
        borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE },
          left: { style: BorderStyle.SINGLE, size: 14, color: NAVY }, insideHorizontal: { style: BorderStyle.NONE }, insideVertical: { style: BorderStyle.NONE } },
        rows: [new TableRow({ children: [new TableCell({ width: { size: CONTENT_W, type: WidthType.DXA },
          shading: { type: ShadingType.CLEAR, fill: 'F7F9FC', color: 'auto' }, margins: { top: 95, bottom: 95, left: 170, right: 140 },
          children: buf.filter((x) => x.trim()).map((x) => new Paragraph({ spacing: { before: 16, after: 16, line: 215, lineRule: 'auto' }, children: inline(x) })) })] })] }));
      c.push(new Paragraph({ spacing: { after: 60 }, children: [] }));
      continue;
    }
    if (/^\s*```/.test(l)) { i++; const buf = [];
      while (i < lines.length && !/^\s*```/.test(lines[i])) { buf.push(lines[i]); i++; } i++;
      buf.forEach((x) => c.push(new Paragraph({ spacing: { before: 0, after: 0, line: 200, lineRule: 'auto' },
        children: [new TextRun({ text: x, size: 14, color: '333333', font: { ascii: 'Consolas', eastAsia: '맑은 고딕', hAnsi: 'Consolas' } })] })));
      c.push(new Paragraph({ spacing: { after: 80 }, children: [] })); continue; }
    if (/^\s*\*(.+)\*\s*$/.test(l.trim()) && l.trim().startsWith('*') && !l.trim().startsWith('**')) {
      c.push(new Paragraph({ spacing: { before: 90, line: 205, lineRule: 'auto' },
        children: inline(l.trim().replace(/^\*|\*$/g, ''), { size: 13, color: '666666' }) })); i++; continue; }
    if (l.trim() === '') { i++; continue; }
    c.push(new Paragraph({ spacing: { before: 40, after: 40, line: 215, lineRule: 'auto' }, children: inline(l) }));
    i++;
  }
  if (!titleDone && titleOverride) c.unshift(new Paragraph({ spacing: { after: 130 }, children: inline(titleOverride, { size: 25, bold: true, color: NAVY }) }));
  const doc = new Document({
    styles: { default: { document: { run: { font: KO, size: 16, color: '222222' }, paragraph: { spacing: { line: 220, lineRule: 'auto' } } } },
      characterStyles: [{ id: 'Hyperlink', name: 'Hyperlink', basedOn: 'DefaultParagraphFont', run: { color: '1B54B5', underline: {} } }] },
    sections: [{ properties: { page: { size: { width: PAGE_W, height: convertMillimetersToTwip(297) },
      margin: { top: MARGIN, bottom: convertMillimetersToTwip(12), left: MARGIN, right: MARGIN } } }, children: c }],
  });
  return Packer.toBuffer(doc).then((b) => { fs.writeFileSync(outPath, b); console.log('written:', outPath, b.length); });
}
convert(process.argv[2], process.argv[3], process.argv[4]);
