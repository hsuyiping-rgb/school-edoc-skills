#!/usr/bin/env node
/**
 * 由 JSON 產生 Word 文件。兩種：
 *   kind: "review" 招標文件審查報告
 *   kind: "spec"   採購品項規格補正建議表
 *
 *   node build_report.js <資料.json> <輸出.docx>
 *
 * 格式與行內標記詳見 references/report-format.md
 */
const path = require('path');
const fs = require('fs');

// docx 可能裝在全域，找不到就往 npm root 找
let D;
try {
  D = require('docx');
} catch (e) {
  try {
    const root = require('child_process').execSync('npm root -g', { encoding: 'utf8' }).trim();
    D = require(path.join(root, 'docx'));
  } catch (e2) {
    console.error('找不到 docx 套件。請執行：npm install -g docx');
    process.exit(1);
  }
}
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, PageBreak,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle, LevelFormat,
} = D;

// ---------- 版面常數 ----------
const FONT = '微軟正黑體';
const SERIF = '標楷體';
const MONO = 'Consolas';
const INK = '1A2621', SOFT = '3C4A44', MUTED = '5E6B64', PRIMARY = '1F4D3D', RULE = 'D8DDD8';
const SEV = {
  high: { label: '重大', color: 'A32A21' },
  fix: { label: '應補正', color: '7E5400' },
  advice: { label: '建議', color: '28557E' },
  ok: { label: '已符合', color: '2E6B4F' },
};
const CONTENT_W = 9330;
const B = (c, sz) => ({ style: BorderStyle.SINGLE, size: sz || 4, color: c || RULE });
const NO_B = { style: BorderStyle.NONE, size: 0, color: 'auto' };

// ---------- 行內標記 ----------
// **粗體**  `法條`  其餘為內文
function inline(str, base = {}) {
  const size = base.size || 20;
  const runs = [];
  const re = /(\*\*[^*]+\*\*|`[^`]+`)/g;
  let last = 0, m;
  const plain = s => { if (s) runs.push(new TextRun({ text: s, font: FONT, size, color: base.color || SOFT })); };
  while ((m = re.exec(str)) !== null) {
    plain(str.slice(last, m.index));
    const tok = m[0];
    if (tok.startsWith('**')) {
      runs.push(new TextRun({ text: tok.slice(2, -2), font: FONT, size, bold: true, color: INK }));
    } else {
      runs.push(new TextRun({ text: tok.slice(1, -1), font: MONO, size: size - 2, bold: true, color: PRIMARY }));
    }
    last = m.index + tok.length;
  }
  plain(str.slice(last));
  return runs.length ? runs : [new TextRun({ text: '', font: FONT, size })];
}

const t = (text, o = {}) => new TextRun({
  text, font: o.font || FONT, size: o.size || 20, bold: o.bold, color: o.color || INK,
});
const p = (children, o = {}) => new Paragraph({
  children: Array.isArray(children) ? children : [children],
  spacing: { before: o.before ?? 0, after: o.after ?? 100, line: o.line ?? 290 },
  indent: o.indent, border: o.border, keepNext: o.keepNext, heading: o.heading,
});
const bullet = (str, o = {}) => new Paragraph({
  children: inline(str, o),
  numbering: { reference: 'dot', level: 0 },
  spacing: { after: 60, line: 285 },
  indent: { left: o.left ?? 300, hanging: 240 },
});
const cell = (children, o = {}) => new TableCell({
  children,
  width: { size: o.w, type: WidthType.DXA },
  shading: o.fill ? { type: ShadingType.CLEAR, color: 'auto', fill: o.fill } : undefined,
  margins: { top: o.pad ?? 100, bottom: o.pad ?? 100, left: o.padx ?? 150, right: o.padx ?? 150 },
  verticalAlign: o.mid ? 'center' : 'top',
  borders: o.plain
    ? { top: NO_B, bottom: NO_B, left: o.leftBorder ? B() : NO_B, right: NO_B }
    : { top: B(), bottom: B(), left: B(), right: B() },
});
const secHead = s => new Paragraph({
  children: [t(s, { font: SERIF, size: 28, bold: true, color: INK })],
  heading: HeadingLevel.HEADING_1,
  spacing: { before: 400, after: 60, line: 300 },
  border: { bottom: B(PRIMARY, 12) },
});
const note = s => p(t(s, { size: 18, color: MUTED }), { after: 180, line: 280 });
const spacer = (after = 140) => p(t(''), { after });

// ---------- 檔頭 ----------
function masthead(meta) {
  const titles = (meta.title || '').split('\n');
  const tags = [
    meta.fileNo && ['檔　　號', meta.fileNo],
    meta.caseNo && ['採購案號', meta.caseNo],
    meta.reviewDate && ['審查日期', meta.reviewDate],
  ].filter(Boolean);
  const left = [
    ...titles.map((ln, i) => p(t(ln, { font: SERIF, size: 32, bold: true }), { after: i === titles.length - 1 ? 90 : 40, line: 340 })),
    meta.subtitle ? p(t(meta.subtitle, { size: 18, color: MUTED }), { after: 0 }) : null,
  ].filter(Boolean);
  if (!tags.length) return [...left, p(t(''), { after: 140, border: { bottom: B(INK, 12) } })];
  return [
    new Table({
      columnWidths: [6730, 2600],
      width: { size: CONTENT_W, type: WidthType.DXA },
      rows: [new TableRow({
        children: [
          cell(left, { w: 6730, plain: true, pad: 0, padx: 0 }),
          cell(tags.map(([k, v], i) => p([
            t(k + '　', { size: 17, color: MUTED }), t(v, { font: MONO, size: 17, color: INK }),
          ], { after: i === tags.length - 1 ? 0 : 30, line: 260 })), { w: 2600, plain: true, leftBorder: true, pad: 0, padx: 0 }),
        ],
      })],
    }),
    p(t(''), { after: 60, border: { bottom: B(INK, 12) } }),
    spacer(120),
  ];
}

// ---------- 案件資料格 ----------
function factsGrid(facts) {
  if (!facts || !facts.length) return [];
  const cols = 3, w = Math.floor(CONTENT_W / cols), rows = [];
  for (let i = 0; i < facts.length; i += cols) {
    const chunk = facts.slice(i, i + cols);
    while (chunk.length < cols) chunk.push(['', '']);
    rows.push(new TableRow({
      children: chunk.map(([k, v]) => cell([
        p(t(k, { size: 16, color: MUTED }), { after: 30, line: 240 }),
        p(t(v, { size: 20, bold: true, color: INK }), { after: 0, line: 260 }),
      ], { w, fill: 'FFFFFF' })),
    }));
  }
  return [new Table({ columnWidths: Array(cols).fill(w), width: { size: CONTENT_W, type: WidthType.DXA }, rows }), spacer(200)];
}

// ---------- 結論方塊 ----------
function calloutBox(o) {
  const kids = [];
  if (o.heading) kids.push(p(t(o.heading, { font: SERIF, size: 26, bold: true, color: INK }), { after: 140, line: 320 }));
  (o.paragraphs || []).forEach((s, i, a) => kids.push(p(inline(s), { after: i === a.length - 1 && !o.tally ? 0 : 120, line: 320 })));
  (o.items || []).forEach((s, i) => kids.push(new Paragraph({
    children: [t(String(i + 1) + '. ', { font: MONO, size: 20, bold: true, color: PRIMARY }), ...inline(s)],
    spacing: { after: 90, line: 300 }, indent: { left: 420, hanging: 300 },
  })));
  if (o.tally) {
    const runs = [];
    o.tally.forEach(([k, n], i) => {
      if (i) runs.push(t('　│　', { size: 19, color: RULE }));
      const s = SEV[k] || {};
      runs.push(t(`${s.label || k} ${n}`, { size: 19, bold: true, color: s.color || PRIMARY }));
    });
    kids.push(p(runs, { after: 0 }));
  }
  return new Table({
    columnWidths: [CONTENT_W], width: { size: CONTENT_W, type: WidthType.DXA },
    rows: [new TableRow({
      cantSplit: !!o.keepTogether,
      children: [new TableCell({
        width: { size: CONTENT_W, type: WidthType.DXA },
        shading: { type: ShadingType.CLEAR, color: 'auto', fill: o.fill || 'F3F5F2' },
        margins: { top: 220, bottom: 220, left: 260, right: 260 },
        borders: { top: B(o.topColor || RULE, o.topColor ? 18 : 4), bottom: B(), left: B(), right: B() },
        children: kids,
      })],
    })],
  });
}

// ---------- 發現 ----------
function finding(f) {
  const sev = SEV[f.severity] || SEV.advice;
  const out = [new Paragraph({
    children: [
      t((f.id || '') + '  ', { font: MONO, size: 22, bold: true, color: PRIMARY }),
      t('［' + sev.label + '］', { size: 18, bold: true, color: sev.color }),
      t(' ' + f.title, { font: SERIF, size: 24, bold: true, color: INK }),
    ],
    spacing: { before: 300, after: 120, line: 300 },
    keepNext: true,
  })];
  (f.blocks || []).forEach(b => {
    const isFix = /建議|處理|作法/.test(b.label || '');
    if (b.text) {
      out.push(new Paragraph({
        children: [t('【' + b.label + '】', { size: 19, bold: true, color: isFix ? PRIMARY : MUTED }), ...inline(b.text)],
        spacing: { after: (b.bullets && b.bullets.length) ? 60 : 110, line: 300 },
        indent: { left: 360 },
      }));
    } else if (b.label) {
      out.push(new Paragraph({
        children: [t('【' + b.label + '】', { size: 19, bold: true, color: isFix ? PRIMARY : MUTED })],
        spacing: { after: 60, line: 300 }, indent: { left: 360 },
      }));
    }
    (b.bullets || []).forEach(s => out.push(bullet(s, { left: 900 })));
    if (b.bullets && b.bullets.length) out.push(p(t(''), { after: 40, line: 120 }));
  });
  return out;
}

// ---------- 通用表格 ----------
function gridTable(spec) {
  const widths = spec.widths;
  const rows = [];
  if (spec.header) {
    rows.push(new TableRow({
      tableHeader: true,
      children: spec.header.map((h, i) => cell([p(t(h, { size: 17, bold: true, color: MUTED }), { after: 0 })], { w: widths[i], fill: 'EDEFEC' })),
    }));
  }
  spec.rows.forEach(r => {
    rows.push(new TableRow({
      children: r.map((c, i) => {
        const isMono = spec.mono && spec.mono.includes(i);
        const isKey = spec.keyCol === i;
        const lines = String(c).split('\n');
        return cell(lines.map((ln, j) => p(
          isMono || isKey
            ? [t(ln, { font: isMono ? MONO : FONT, size: isKey && j === 0 ? 19 : 18, bold: j === 0, color: j === 0 ? (isMono ? PRIMARY : INK) : MUTED })]
            : inline(ln, { size: 19 }),
          { after: j === lines.length - 1 ? 0 : 30, line: 275 },
        )), { w: widths[i], fill: isKey ? 'F7F8F6' : 'FFFFFF', mid: !!spec.midCols && spec.midCols.includes(i) });
      }),
    }));
  });
  return new Table({ columnWidths: widths, width: { size: CONTENT_W, type: WidthType.DXA }, rows });
}

// 左窄右寬、右欄為條列的表（規格補正、共通條款都用這個）
function listTable(spec) {
  const widths = spec.widths || [1900, CONTENT_W - 1900];
  const rows = [];
  if (spec.header) {
    rows.push(new TableRow({
      tableHeader: true,
      children: spec.header.map((h, i) => cell([p(t(h, { size: 17, bold: true, color: MUTED }), { after: 0 })], { w: widths[i], fill: 'EDEFEC' })),
    }));
  }
  spec.rows.forEach(r => {
    const keyLines = String(r.key).split('\n');
    const right = Array.isArray(r.bullets) && r.bullets.length
      ? r.bullets.map(s => bullet(s))
      : [p(inline(r.text || '', { size: 19 }), { after: 0 })];
    rows.push(new TableRow({
      children: [
        cell(keyLines.map((ln, j) => p(
          t(ln, { font: j === 0 && spec.monoKey ? MONO : FONT, size: j === 0 ? 19 : 18, bold: j <= (spec.boldKeyLines ?? 0), color: j === 0 && spec.monoKey ? PRIMARY : (j === 0 ? INK : MUTED) }),
          { after: j === keyLines.length - 1 ? 0 : 30, line: 265 },
        )), { w: widths[0], fill: 'F7F8F6', mid: !!spec.midKey }),
        cell(right, { w: widths[1] }),
      ],
    }));
  });
  return new Table({ columnWidths: widths, width: { size: CONTENT_W, type: WidthType.DXA }, rows });
}

// ---------- 章節組裝 ----------
function renderSections(sections) {
  const out = [];
  (sections || []).forEach(s => {
    if (s.pageBreak) out.push(new Paragraph({ children: [new PageBreak()] }));
    if (s.heading) out.push(secHead(s.heading));
    if (s.note) out.push(note(s.note));
    (s.paragraphs || []).forEach(str => out.push(p(inline(str), { after: 110 })));
    (s.findings || []).forEach(f => out.push(...finding(f)));
    if (s.table) out.push(gridTable(s.table));
    if (s.listTable) out.push(listTable(s.listTable));
    (s.numbered || []).forEach((str, i) => out.push(new Paragraph({
      children: [t(String(i + 1) + '. ', { font: MONO, size: 20, bold: true, color: PRIMARY }), ...inline(str)],
      spacing: { after: 80, line: 290 }, indent: { left: 420, hanging: 300 },
    })));
    (s.bullets || []).forEach(str => out.push(bullet(str, { left: 620 })));
    // 表格之後的收尾段落（renderSections 的順序是固定的，這個欄位讓你把話說在表格後面）
    (s.afterParagraphs || []).forEach(str => out.push(p(inline(str), { before: 120, after: 110 })));
    if (s.callout) out.push(calloutBox(s.callout));
    if (s.spacerAfter !== false) out.push(spacer(s.spacerAfter || 60));
  });
  return out;
}

// ---------- 主流程 ----------
function build(data) {
  const children = [
    ...masthead(data.meta || {}),
    ...factsGrid(data.facts),
  ];
  if (data.verdict) {
    children.push(calloutBox({ ...data.verdict, topColor: data.verdict.topColor || SEV.high.color }), spacer(60));
  }
  children.push(...renderSections(data.sections));
  if (data.closing) children.push(calloutBox({ ...data.closing, keepTogether: true }));
  if (data.footer) {
    children.push(p(t(''), { before: 280, after: 60, border: { top: B(RULE, 4) } }));
    children.push(p(t(data.footer, { size: 17, color: MUTED }), { line: 280 }));
  }

  return new Document({
    creator: (data.meta && data.meta.creator) || '招標文件審查',
    title: (data.meta && (data.meta.title || '').replace(/\n/g, ' ')) || '審查報告',
    numbering: {
      config: [{
        reference: 'dot',
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: '▪', alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 300, hanging: 240 } }, run: { color: PRIMARY, font: FONT, size: 16 } },
        }],
      }],
    },
    styles: { default: { document: { run: { font: FONT, size: 20, color: INK }, paragraph: { spacing: { line: 290 } } } } },
    sections: [{
      properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1080, bottom: 1080, left: 1080, right: 1080 } } },
      children,
    }],
  });
}

const [, , jsonPath, outPath] = process.argv;
if (!jsonPath || !outPath) {
  console.error('用法：node build_report.js <資料.json> <輸出.docx>');
  process.exit(1);
}
const data = JSON.parse(fs.readFileSync(jsonPath, 'utf8'));
Packer.toBuffer(build(data)).then(buf => {
  fs.writeFileSync(outPath, buf);
  console.log(`已產生 ${outPath}（${buf.length} bytes）`);
  console.log('接著執行 render_check.py 檢查排版。');
}).catch(e => { console.error(e); process.exit(1); });
