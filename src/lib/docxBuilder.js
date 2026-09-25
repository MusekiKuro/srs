import {
  AlignmentType,
  Document,
  Footer,
  HeadingLevel,
  LineRuleType,
  Packer,
  PageNumber,
  Paragraph,
  TableOfContents,
  TextRun,
} from 'docx';
import { GOST } from './gost.js';

function run(text, opts = {}) {
  return new TextRun({
    text: String(text ?? ''),
    font: GOST.font,
    size: opts.size ?? GOST.bodySize,
    bold: opts.bold ?? false,
    color: GOST.black,
  });
}

function bodyPara(text) {
  return new Paragraph({
    children: [run(text)],
    alignment: AlignmentType.JUSTIFIED,
    indent: { firstLine: GOST.firstLineIndent },
    spacing: { before: 0, after: 0, line: GOST.lineOneHalf, lineRule: LineRuleType.AUTO },
  });
}

function headingPara(text, level, breakBefore) {
  return new Paragraph({
    children: [run(text, { bold: true })],
    heading: level === 2 ? HeadingLevel.HEADING_2 : HeadingLevel.HEADING_1,
    alignment: level === 2 ? AlignmentType.LEFT : AlignmentType.CENTER,
    pageBreakBefore: breakBefore,
    spacing: { before: 0, after: 120 },
  });
}

export function buildGostDocument(doc) {
  const footer = new Footer({
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ children: [PageNumber.CURRENT], font: GOST.font, size: GOST.footerSize })],
      }),
    ],
  });
  const children = [
    new Paragraph({
      children: [run('СОДЕРЖАНИЕ', { bold: true })],
      heading: HeadingLevel.HEADING_1,
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 120 },
    }),
    new TableOfContents('СОДЕРЖАНИЕ', { hyperlink: true, headingStyleRange: '1-2' }),
  ];
  for (const ch of doc.chapters ?? []) {
    children.push(headingPara(ch.heading, ch.level, true));
    for (const p of ch.paragraphs ?? []) {
      if (String(p ?? '').trim()) children.push(bodyPara(p));
    }
  }
  if ((doc.bibliography ?? []).length) {
    children.push(headingPara('СПИСОК ИСПОЛЬЗОВАННОЙ ЛИТЕРАТУРЫ', 1, true));
    for (const src of doc.bibliography) {
      children.push(new Paragraph({
        children: [run(src)],
        alignment: AlignmentType.JUSTIFIED,
        spacing: { before: 0, after: 0, line: GOST.lineOneHalf, lineRule: LineRuleType.AUTO },
      }));
    }
  }
  return new Document({
    sections: [{
      properties: { page: { size: { ...GOST.pageSize }, margin: { ...GOST.margins } } },
      footers: { default: footer },
      children,
    }],
  });
}

export function docxBlob(doc) {
  return Packer.toBlob(buildGostDocument(doc));
}

export function downloadDocx(blob, filename) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}
