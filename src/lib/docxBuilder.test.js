import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { Packer } from 'docx';
import { buildGostDocument } from './docxBuilder.js';

const doc = {
  kind: 'srs',
  title: 'Тест',
  chapters: [
    { heading: 'ВВЕДЕНИЕ', level: 1, paragraphs: ['Первый абзац.', 'Второй абзац.'] },
    { heading: '1. Теория', level: 1, paragraphs: ['Суть.'] },
  ],
  bibliography: ['Иванов, И.И. Основы. – М.: Юрайт, 2023. – 10 с.'],
};

async function documentXml() {
  const buf = await Packer.toBuffer(buildGostDocument(doc));
  const zip = await JSZip.loadAsync(buf);
  return zip.file('word/document.xml').async('text');
}

describe('docxBuilder', () => {
  it('embeds TOC field, headings and GOST values', async () => {
    const xml = await documentXml();
    expect(xml).toContain('TOC');
    expect(xml).toContain('Heading1');
    expect(xml).toContain('Times New Roman');
    expect(xml).toContain('w:left="1701"');
    expect(xml).toContain('Первый абзац.');
  });
  it('has no service hints and no title page', async () => {
    const xml = await documentXml();
    expect(xml).not.toContain('Обновить поле');
    expect(xml).not.toContain('МИНИСТЕРСТВО');
    expect(xml).toContain('СПИСОК ИСПОЛЬЗОВАННОЙ ЛИТЕРАТУРЫ');
  });
});
