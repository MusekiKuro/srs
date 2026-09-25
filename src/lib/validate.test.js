import { describe, expect, it } from 'vitest';
import { canBiblio, canDownload, canWriteText, makeFilename, settingsReady } from './validate.js';

const emptyDoc = { kind: 'srs', title: 'Тема', chapters: [], bibliography: [] };
const readyDoc = { kind: 'practice', title: 'Учет дебиторки: анализ!', chapters: [{ heading: 'ВВЕДЕНИЕ', paragraphs: ['текст'] }], bibliography: [] };

describe('validate', () => {
  it('checks settings', () => {
    expect(settingsReady({ baseUrl: 'https://x/v1', apiKey: 'k', model: 'm' })).toBe(true);
    expect(settingsReady({ baseUrl: 'https://x/v1', apiKey: '  ', model: 'm' })).toBe(false);
    expect(settingsReady(null)).toBe(false);
  });
  it('gates writing on plan presence', () => {
    expect(canWriteText(emptyDoc)).toBe(false);
    expect(canWriteText(readyDoc)).toBe(true);
  });
  it('gates download and biblio on text presence', () => {
    expect(canDownload(emptyDoc)).toBe(false);
    expect(canBiblio(emptyDoc)).toBe(false);
    expect(canDownload({ ...emptyDoc, chapters: [{ heading: 'Г', paragraphs: [''] }] })).toBe(false);
    expect(canDownload(readyDoc)).toBe(true);
    expect(canBiblio(readyDoc)).toBe(true);
  });
  it('makes safe filenames', () => {
    expect(makeFilename(readyDoc)).toBe('Отчет_практика_Учет-дебиторки-анализ.docx');
    expect(makeFilename({ kind: 'srs', title: 'ИИ в образовании' })).toBe('СРС_ИИ-в-образовании.docx');
  });
});
