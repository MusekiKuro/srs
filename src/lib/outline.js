import { GOST } from './gost.js';

const KIND_LABEL = { srs: 'СРС (самостоятельная работа студента)', practice: 'отчет по производственной практике' };

export function buildPlanPrompt({ title, kind, pages }) {
  return [
    `Составь структуру работы «${title}» (${KIND_LABEL[kind] ?? kind}, объем ${pages} стр).`,
    'Ответь СТРОГО массивом JSON без пояснений, вида [{"heading": "...", "level": 1}].',
    'Правила: первый раздел — "ВВЕДЕНИЕ" (level 1), последний текстовый — "ЗАКЛЮЧЕНИЕ" (level 1),',
    'между ними 2–4 главы (level 1) с параграфами (level 2) где нужно. Без markdown, без комментариев.',
  ].join('\n');
}

export function buildChapterPrompt(title, chapter, targetWords) {
  return [
    `Пишем работу «${title}». Раздел: ${chapter.heading}.`,
    `Напиши около ${targetWords} слов обычным текстом, 3–6 абзацев, научным студенческим стилем.`,
    'Без заголовка в начале, без markdown-разметки, без списков литературы внутри.',
  ].join('\n');
}

export function buildBiblioPrompt(title, chapters) {
  const heads = chapters.map((c) => c.heading).join('; ');
  return [
    `Работа «${title}», разделы: ${heads}.`,
    'Составь список использованной литературы (7–12 источников) готовыми строками по ГОСТ Р 7.0.5-2008,',
    'каждый источник с новой строки, без нумерации: автор, название, город: издательство, год.',
  ].join('\n');
}

export function stripCodeFences(text) {
  return String(text ?? '').replace(/```(?:json)?\s*/gi, '').replace(/```/g, '');
}

export function extractPlanArray(text) {
  const clean = stripCodeFences(text);
  const start = clean.indexOf('[');
  const end = clean.lastIndexOf(']');
  if (start < 0 || end <= start) throw new Error('В ответе нет списка глав.');
  let arr;
  try {
    arr = JSON.parse(clean.slice(start, end + 1));
  } catch {
    throw new Error('Неверный формат плана, попробуйте ещё раз.');
  }
  if (!Array.isArray(arr) || !arr.length) throw new Error('Неверный формат плана, попробуйте ещё раз.');
  return arr.map((item) => {
    if (!item || typeof item.heading !== 'string' || ![1, 2].includes(item.level)) {
      throw new Error('Неверный формат плана, попробуйте ещё раз.');
    }
    return { heading: item.heading.trim(), level: item.level };
  });
}

export function sanitizeMarkdown(text) {
  return String(text ?? '')
    .replace(/^#{1,6}\s+/gm, '')
    .replace(/\*\*([^*]+)\*\*/g, '$1')
    .replace(/(^|\W)\*([^*\n]+)\*/g, '$1$2')
    .replace(/`([^`]+)`/g, '$1')
    .trim();
}

export function countWords(text) {
  const words = String(text ?? '').trim().split(/\s+/).filter(Boolean);
  return words.length === 1 && words[0] === '' ? 0 : words.length;
}

export function countPages(text) {
  return countWords(text) / GOST.wordsPerPage;
}
