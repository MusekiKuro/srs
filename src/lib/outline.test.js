import { describe, expect, it } from 'vitest';
import { buildChapterPrompt, buildPlanPrompt, countPages, countWords, extractPlanArray, sanitizeMarkdown } from './outline.js';

describe('outline', () => {
  it('extracts plan array from fenced response', () => {
    const raw = 'Вот план:\n```json\n[{"heading": "ВВЕДЕНИЕ", "level": 1}]\n```\nУдачи!';
    expect(extractPlanArray(raw)).toEqual([{ heading: 'ВВЕДЕНИЕ', level: 1 }]);
  });
  it('extracts plan array from bare response', () => {
    expect(extractPlanArray('[{"heading": "1. Теория", "level": 1}]')).toEqual([{ heading: '1. Теория', level: 1 }]);
  });
  it('throws clear error when no array', () => {
    expect(() => extractPlanArray('Просто текст без списка.')).toThrow('В ответе нет списка глав.');
  });
  it('throws on bad chapter shape', () => {
    expect(() => extractPlanArray('[{"foo": 1}]')).toThrow(/неверный формат/i);
  });
  it('sanitizes markdown to plain text', () => {
    expect(sanitizeMarkdown('# Заголовок\nТекст **жирный** и *курсив* и `код`.')).toBe('Заголовок\nТекст жирный и курсив и код.');
  });
  it('counts words and pages', () => {
    const t = Array(650).fill('слово').join(' ');
    expect(countWords(t)).toBe(650);
    expect(countPages(t)).toBeCloseTo(2, 1);
  });
  it('builds prompts with topic and volume', () => {
    const p = buildPlanPrompt({ title: 'ИИ в образовании', kind: 'srs', pages: 8 });
    expect(p).toContain('ИИ в образовании');
    expect(p).toContain('JSON');
    const c = buildChapterPrompt('Тема', { heading: 'ВВЕДЕНИЕ' }, 300);
    expect(c).toContain('ВВЕДЕНИЕ');
  });
});
