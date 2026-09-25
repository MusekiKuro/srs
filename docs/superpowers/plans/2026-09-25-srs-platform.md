# SRS Platform Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a static React + Vite SPA where a student enters topic/type/pages, generates plan and chapter text via their own LLM key, edits in browser, and downloads a GOST `.docx`.

**Architecture:** No backend. Four pure lib modules (`storage`, `llm`, `outline`, `validate`, `docxBuilder`) with colocated vitest tests, a pure reducer + thin React context for state, five components for three screens. `docx@9.7.2` builds the file in-browser; `Packer.toBuffer` is used in Node tests.

**Tech Stack:** React 19.3.0, react-dom 19.3.0, Vite 8.3.1, @vitejs/plugin-react 6.1.1, docx 9.7.2, vitest 5.0.1, jszip ^3.10.2 (dev, tests only).

**Spec:** `docs/superpowers/specs/2026-09-25-srs-platform-design.md`

## Global Constraints

- Static SPA, no backend, no server functions, no secrets in code — one line.
- LLM key lives in `localStorage` only, sent only to the user-chosen provider — one line.
- GOST exact values: Times New Roman, body 14pt (size 28 half-points), footer 12pt (size 24), line 1.5 (spacing line 360, AUTO rule), first-line indent 1.25cm (709 twips), margins mm 30/10/20/20 → twips left 1701 / right 567 / top 1134 / bottom 1134, A4 11906x16838, headings via HeadingLevel styles, TOC field `headingStyleRange "1-2"`, page numbers bottom-center, first page is СОДЕРЖАНИЕ, no title page, no helper text inside the file — one line.
- Russian UI copy — one line.
- Tests colocated as `src/**/*.test.js`, run with `npm test` (`vitest run`), follow existing `kp-generator` patterns — one line.
- ~300–350 words per page (`wordsPerPage: 325`) — one line.

## Review Focus

- Model wraps plan JSON in ```json fences or prose → outline must extract the array or throw a clear error; test pins it in Task 4.
- Model returns markdown (`#`, `**`, backticks) in chapter text → sanitize to plain text before storing; test pins it in Task 4.
- `baseUrl` with trailing slash or duplicated path → normalize before fetch; test pins it in Task 3.
- Stream aborts mid-chapter → already-streamed tokens must stay, error surfaces for retry; test pins it in Task 3.
- Cyrillic/spaces in download filename → sanitize to safe name; test pins it in Task 7.

---

## File Structure

New files (repo root is the app; skill lives in `srs-practice-report/` untouched):

- `package.json` — deps + scripts (dev/build/test).
- `vite.config.js` — react plugin.
- `index.html` — root div + script to `/src/main.jsx`.
- `src/main.jsx` — React root, imports App + styles.
- `src/styles.css` — деловой стиль (тёмно-синий акцент, белый фон).
- `src/App.jsx` — screen router: list | new | editor; DocsProvider wrapper.
- `src/lib/gost.js` — GOST constants + `mmToTwips`.
- `src/lib/gost.test.js`
- `src/lib/storage.js` — localStorage CRUD for documents + settings.
- `src/lib/storage.test.js`
- `src/lib/llm.js` — presets, URL normalize, SSE streaming client.
- `src/lib/llm.test.js`
- `src/lib/outline.js` — prompts, plan-JSON extract, markdown sanitize, word/page count.
- `src/lib/outline.test.js`
- `src/lib/validate.js` — settingsReady, canWriteText, canDownload, canBiblio, makeFilename.
- `src/lib/validate.test.js`
- `src/lib/docxBuilder.js` — GOST Document assembly, blob, download.
- `src/lib/docxBuilder.test.js`
- `src/state/docReducer.js` — pure `docReducer` + `docsReducer`.
- `src/state/docReducer.test.js`
- `src/state/store.jsx` — thin DocsProvider + useDocs (no test; wiring only).
- `src/components/WorksList.jsx`, `NewWorkForm.jsx`, `SettingsModal.jsx`, `Editor.jsx`, `ChapterView.jsx` — UI (verified via build + manual checklist).

---

### Task 1: Scaffold Vite + React

**Files:**
- Create: `package.json`, `vite.config.js`, `index.html`, `src/main.jsx`, `src/App.jsx`, `src/styles.css`

**Interfaces:**
- Consumes: nothing.
- Produces: `npm run dev` serves App; `npm run build` outputs `dist/`; `npm test` runs (0 tests, exit 0).

- [ ] **Step 1: Write package.json**

```json
{
  "name": "srs-platform",
  "private": true,
  "version": "0.1.0",
  "type": "module",
  "scripts": {
    "dev": "vite",
    "build": "vite build",
    "test": "vitest run"
  },
  "dependencies": {
    "docx": "9.7.2",
    "react": "19.3.0",
    "react-dom": "19.3.0"
  },
  "devDependencies": {
    "@vitejs/plugin-react": "6.1.1",
    "jszip": "^3.10.2",
    "vite": "8.3.1",
    "vitest": "5.0.1"
  }
}
```

- [ ] **Step 2: Write vite.config.js**

```js
import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig({
  plugins: [react()],
});
```

- [ ] **Step 3: Write index.html**

```html
<!doctype html>
<html lang="ru">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>SRS Platform — СРС и отчеты по практике</title>
  </head>
  <body>
    <div id="root"></div>
    <script type="module" src="/src/main.jsx"></script>
  </body>
</html>
```

- [ ] **Step 4: Write src/main.jsx, src/App.jsx, src/styles.css**

```jsx
// src/main.jsx
import React from 'react';
import { createRoot } from 'react-dom/client';
import App from './App.jsx';
import './styles.css';

createRoot(document.getElementById('root')).render(<App />);
```

```jsx
// src/App.jsx
export default function App() {
  return (
    <main className="app">
      <h1>SRS Platform</h1>
      <p>СРС и отчеты по практике — от темы до docx по ГОСТу.</p>
    </main>
  );
}
```

```css
/* src/styles.css */
:root { --accent: #13243b; --bg: #ffffff; --border: #d9e0ea; }
body { margin: 0; font-family: 'Times New Roman', Georgia, serif; background: var(--bg); color: #111; }
.app { max-width: 960px; margin: 0 auto; padding: 24px; }
```

- [ ] **Step 5: Install and verify**

```bash
npm install
npm test
```

Expected: install succeeds; `vitest run` exits 0 with "No test files found".

- [ ] **Step 6: Verify build**

```bash
npm run build
```

Expected: `dist/index.html` exists, no errors.

- [ ] **Step 7: Commit**

```bash
git add package.json vite.config.js index.html src/main.jsx src/App.jsx src/styles.css
git commit -m "feat: scaffold React Vite app"
```

---

### Task 2: GOST constants + storage

**Files:**
- Create: `src/lib/gost.js`, `src/lib/gost.test.js`, `src/lib/storage.js`, `src/lib/storage.test.js`

**Interfaces:**
- Consumes: nothing.
- Produces: `GOST`, `mmToTwips(mm)`; `loadDocuments/saveDocuments/loadSettings/saveSettings/newId` — used by Tasks 6–9.

- [ ] **Step 1: Write failing test for gost**

```js
// src/lib/gost.test.js
import { describe, expect, it } from 'vitest';
import { GOST, mmToTwips } from './gost.js';

describe('gost', () => {
  it('converts mm to twips', () => {
    expect(mmToTwips(30)).toBe(1701);
    expect(mmToTwips(10)).toBe(567);
  });
  it('holds exact GOST values', () => {
    expect(GOST.font).toBe('Times New Roman');
    expect(GOST.bodySize).toBe(28);
    expect(GOST.margins).toEqual({ top: 1134, right: 567, bottom: 1134, left: 1701 });
    expect(GOST.firstLineIndent).toBe(709);
    expect(GOST.lineOneHalf).toBe(360);
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npm test -- src/lib/gost.test.js`
Expected: FAIL with "Cannot find module './gost.js'".

- [ ] **Step 3: Write minimal implementation**

```js
// src/lib/gost.js
export const MM_TO_TWIPS = 1440 / 25.4;

export function mmToTwips(mm) {
  return Math.round(mm * MM_TO_TWIPS);
}

export const GOST = {
  font: 'Times New Roman',
  black: '000000',
  bodySize: 28, // 14pt in half-points
  footerSize: 24, // 12pt in half-points
  lineSingle: 240,
  lineOneHalf: 360, // 1.5 spacing with AUTO rule
  firstLineIndent: 709, // 1.25cm in twips
  margins: { top: 1134, right: 567, bottom: 1134, left: 1701 }, // 20/10/20/30 mm
  pageSize: { width: 11906, height: 16838 }, // A4 in twips
  wordsPerPage: 325,
};
```

- [ ] **Step 4: Write failing test for storage**

```js
// src/lib/storage.test.js
import { beforeEach, describe, expect, it } from 'vitest';
import { loadDocuments, loadSettings, saveDocuments, saveSettings } from './storage.js';

function mockLocalStorage() {
  let store = {};
  globalThis.localStorage = {
    getItem: (k) => (k in store ? store[k] : null),
    setItem: (k, v) => { store[k] = String(v); },
    removeItem: (k) => { delete store[k]; },
    clear: () => { store = {}; },
  };
}

describe('storage', () => {
  beforeEach(mockLocalStorage);
  it('round-trips documents', () => {
    const docs = [{ id: 'a', title: 'Тема' }];
    expect(saveDocuments(docs)).toEqual({ ok: true });
    expect(loadDocuments()).toEqual(docs);
  });
  it('returns fallbacks on empty or broken data', () => {
    expect(loadDocuments()).toEqual([]);
    localStorage.setItem('srs-platform:documents', 'not-json{{{');
    expect(loadDocuments()).toEqual([]);
  });
  it('round-trips settings', () => {
    const s = { baseUrl: 'https://x/v1', apiKey: 'k', model: 'm' };
    expect(saveSettings(s)).toEqual({ ok: true });
    expect(loadSettings()).toEqual(s);
  });
  it('reports quota errors', () => {
    const err = new Error('quota'); err.name = 'QuotaExceededError';
    globalThis.localStorage = { getItem: () => null, setItem: () => { throw err; }, removeItem: () => {} };
    expect(saveDocuments([{ id: 'a' }])).toEqual({ ok: false, reason: 'quota' });
  });
});
```

- [ ] **Step 5: Run test to verify it fails**

Run: `npm test -- src/lib/storage.test.js`
Expected: FAIL with "Cannot find module './storage.js'".

- [ ] **Step 6: Write minimal implementation**

```js
// src/lib/storage.js
const DOCS_KEY = 'srs-platform:documents';
const SETTINGS_KEY = 'srs-platform:settings';

function read(key, fallback) {
  try {
    const raw = localStorage.getItem(key);
    return raw ? JSON.parse(raw) : fallback;
  } catch {
    return fallback;
  }
}

function write(key, value) {
  try {
    localStorage.setItem(key, JSON.stringify(value));
    return { ok: true };
  } catch (e) {
    if (e && (e.name === 'QuotaExceededError' || e.code === 22)) {
      return { ok: false, reason: 'quota' };
    }
    return { ok: false, reason: 'unknown' };
  }
}

export function loadDocuments() { return read(DOCS_KEY, []); }
export function saveDocuments(docs) { return write(DOCS_KEY, docs); }
export function loadSettings() { return read(SETTINGS_KEY, { baseUrl: '', apiKey: '', model: '' }); }
export function saveSettings(s) { return write(SETTINGS_KEY, s); }

export function newId() {
  if (globalThis.crypto?.randomUUID) return globalThis.crypto.randomUUID();
  return `id-${Date.now()}-${Math.floor(Math.random() * 1e9)}`;
}
```

- [ ] **Step 7: Run tests to verify they pass**

Run: `npm test -- src/lib/gost.test.js src/lib/storage.test.js`
Expected: PASS, 7 tests.

- [ ] **Step 8: Commit**

```bash
git add src/lib/gost.js src/lib/gost.test.js src/lib/storage.js src/lib/storage.test.js
git commit -m "feat: add GOST constants and localStorage layer"
```

---

### Task 3: LLM streaming client

**Files:**
- Create: `src/lib/llm.js`, `src/lib/llm.test.js`

**Interfaces:**
- Consumes: nothing.
- Produces: `PRESETS`, `normalizeBaseUrl(url)`, `chatUrl(baseUrl)`, `streamChat({baseUrl, apiKey, model, messages, onToken, signal})` — used by Task 9. Error objects carry `.status`.

- [ ] **Step 1: Write the failing test**

```js
// src/lib/llm.test.js
import { describe, expect, it, vi } from 'vitest';
import { chatUrl, normalizeBaseUrl, streamChat } from './llm.js';

function sseResponse(chunks) {
  const encoder = new TextEncoder();
  const payload = chunks.map((c) => `data: ${JSON.stringify(c)}\n\n`).join('') + 'data: [DONE]\n\n';
  return new Response(encoder.encode(payload));
}

describe('llm', () => {
  it('normalizes base url', () => {
    expect(normalizeBaseUrl('https://x/v1///')).toBe('https://x/v1');
    expect(chatUrl('https://x/v1/')).toBe('https://x/v1/chat/completions');
    expect(chatUrl('https://x/v1')).toBe('https://x/v1/chat/completions');
  });
  it('streams tokens and returns full text', async () => {
    globalThis.fetch = vi.fn().mockResolvedValue(sseResponse([
      { choices: [{ delta: { content: 'При' } }] },
      { choices: [{ delta: { content: 'вет' } }] },
    ]));
    const seen = [];
    const full = await streamChat({
      baseUrl: 'https://x/v1', apiKey: 'k', model: 'm',
      messages: [{ role: 'user', content: 'hi' }],
      onToken: (t) => seen.push(t),
    });
    expect(full).toBe('Привет');
    expect(seen).toEqual(['При', 'вет']);
  });
  it('throws status error on 401', async () => {
    globalThis.fetch = vi.fn().mockResolvedValue(new Response('no', { status: 401 }));
    await expect(streamChat({ baseUrl: 'https://x/v1', apiKey: 'bad', model: 'm', messages: [] }))
      .rejects.toMatchObject({ status: 401 });
  });
  it('keeps partial text when stream breaks', async () => {
    const encoder = new TextEncoder();
    const broken = encoder.encode('data: {"choices":[{"delta":{"content":"Часть"}}]}\n\ndata: BROKEN{{{');
    globalThis.fetch = vi.fn().mockResolvedValue(new Response(broken));
    const seen = [];
    await expect(streamChat({
      baseUrl: 'https://x/v1', apiKey: 'k', model: 'm', messages: [],
      onToken: (t) => seen.push(t),
    })).rejects.toThrow();
    expect(seen).toEqual(['Часть']);
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npm test -- src/lib/llm.test.js`
Expected: FAIL with "Cannot find module './llm.js'".

- [ ] **Step 3: Write minimal implementation**

```js
// src/lib/llm.js
export const PRESETS = [
  { id: 'openai', label: 'OpenAI', baseUrl: 'https://api.openai.com/v1', model: 'gpt-4o-mini' },
  { id: 'openrouter', label: 'OpenRouter', baseUrl: 'https://openrouter.ai/api/v1', model: 'openai/gpt-4o-mini' },
  { id: 'deepseek', label: 'DeepSeek', baseUrl: 'https://api.deepseek.com/v1', model: 'deepseek-chat' },
  { id: 'groq', label: 'Groq', baseUrl: 'https://api.groq.com/openai/v1', model: 'llama-3.1-8b-instant' },
  { id: 'custom', label: 'Свой URL', baseUrl: '', model: '' },
];

export function normalizeBaseUrl(url) {
  return String(url ?? '').trim().replace(/\/+$/, '');
}

export function chatUrl(baseUrl) {
  return `${normalizeBaseUrl(baseUrl)}/chat/completions`;
}

export async function streamChat({ baseUrl, apiKey, model, messages, onToken, signal }) {
  const res = await fetch(chatUrl(baseUrl), {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${apiKey}` },
    body: JSON.stringify({ model, messages, stream: true }),
    signal,
  });
  if (!res.ok) {
    const err = new Error(res.status === 401 ? 'Неверный ключ. Проверьте ключ в настройках.' : `Сервер вернул ${res.status}.`);
    err.status = res.status;
    throw err;
  }
  const reader = res.body.getReader();
  const decoder = new TextDecoder();
  let buf = '';
  let full = '';
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    buf += decoder.decode(value, { stream: true });
    const parts = buf.split('\n\n');
    buf = parts.pop();
    for (const part of parts) {
      const line = part.trim();
      if (!line.startsWith('data:')) continue;
      const payload = line.slice(5).trim();
      if (payload === '[DONE]') continue;
      const json = JSON.parse(payload);
      const piece = json?.choices?.[0]?.delta?.content ?? '';
      if (piece) { full += piece; onToken?.(piece); }
    }
  }
  return full;
}
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `npm test -- src/lib/llm.test.js`
Expected: PASS, 4 tests.

- [ ] **Step 5: Commit**

```bash
git add src/lib/llm.js src/lib/llm.test.js
git commit -m "feat: add LLM streaming client"
```

---

### Task 4: Outline prompts + parsing

**Files:**
- Create: `src/lib/outline.js`, `src/lib/outline.test.js`

**Interfaces:**
- Consumes: `GOST.wordsPerPage` from Task 2.
- Produces: `buildPlanPrompt(meta)`, `buildChapterPrompt(title, chapter, targetWords)`, `buildBiblioPrompt(title, chapters)`, `extractPlanArray(text)`, `sanitizeMarkdown(text)`, `countWords(text)`, `countPages(text)` — used by Task 9.

- [ ] **Step 1: Write the failing test**

```js
// src/lib/outline.test.js
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
    expect(() => extractPlanArray('[{"foo": 1}]')).toThrow('неверный формат');
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npm test -- src/lib/outline.test.js`
Expected: FAIL with "Cannot find module './outline.js'".

- [ ] **Step 3: Write minimal implementation**

```js
// src/lib/outline.js
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
    throw new Error('План в неверном формате, попробуйте ещё раз.');
  }
  if (!Array.isArray(arr) || !arr.length) throw new Error('План в неверном формате, попробуйте ещё раз.');
  return arr.map((item) => {
    if (!item || typeof item.heading !== 'string' || ![1, 2].includes(item.level)) {
      throw new Error('План в неверном формате, попробуйте ещё раз.');
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
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `npm test -- src/lib/outline.test.js`
Expected: PASS, 7 tests.

- [ ] **Step 5: Commit**

```bash
git add src/lib/outline.js src/lib/outline.test.js
git commit -m "feat: add outline prompts and parsing"
```

---

### Task 5: Validation helpers

**Files:**
- Create: `src/lib/validate.js`, `src/lib/validate.test.js`

**Interfaces:**
- Consumes: nothing (takes plain doc/settings objects shaped in Task 6).
- Produces: `settingsReady(s)`, `canWriteText(doc)`, `canDownload(doc)`, `canBiblio(doc)`, `makeFilename(doc)` — used by Task 9.

- [ ] **Step 1: Write the failing test**

```js
// src/lib/validate.test.js
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
    expect(makeFilename(readyDoc)).toBe('Отчет_практика-Учет-дебиторки-анализ.docx');
    expect(makeFilename({ kind: 'srs', title: 'ИИ в образовании' })).toBe('СРС-ИИ-в-образовании.docx');
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npm test -- src/lib/validate.test.js`
Expected: FAIL with "Cannot find module './validate.js'".

- [ ] **Step 3: Write minimal implementation**

```js
// src/lib/validate.js
export function settingsReady(s) {
  return Boolean(s && String(s.baseUrl ?? '').trim() && String(s.apiKey ?? '').trim() && String(s.model ?? '').trim());
}

function chapterText(ch) {
  return (ch?.paragraphs ?? []).join('\n').trim();
}

export function hasPlan(doc) {
  return (doc?.chapters ?? []).length > 0;
}

export function hasAnyText(doc) {
  return (doc?.chapters ?? []).some((ch) => chapterText(ch).length > 0);
}

export function canWriteText(doc) { return hasPlan(doc); }
export function canDownload(doc) { return hasAnyText(doc); }
export function canBiblio(doc) { return hasAnyText(doc); }

export function makeFilename(doc) {
  const prefix = doc?.kind === 'practice' ? 'Отчет_практика' : 'СРС';
  const base = `${prefix}_${doc?.title ?? 'работа'}`;
  const safe = base.replace(/[^a-zа-яё0-9_-]+/gi, '-').replace(/-+/g, '-').replace(/^-|-$/g, '').slice(0, 80);
  return `${safe || 'document'}.docx`;
}
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `npm test -- src/lib/validate.test.js`
Expected: PASS, 4 tests.

- [ ] **Step 5: Commit**

```bash
git add src/lib/validate.js src/lib/validate.test.js
git commit -m "feat: add validation helpers"
```

---

### Task 6: Document reducer + store

**Files:**
- Create: `src/state/docReducer.js`, `src/state/docReducer.test.js`, `src/state/store.jsx`

**Interfaces:**
- Consumes: `newId` from Task 2.
- Produces: `docReducer(doc, action)`, `docsReducer(state, action)`, `DocsProvider`, `useDocs()` — used by Tasks 8–9. `store.jsx` is thin wiring, no test.

- [ ] **Step 1: Write the failing test**

```js
// src/state/docReducer.test.js
import { describe, expect, it } from 'vitest';
import { docReducer, docsReducer, initialDoc } from './docReducer.js';

describe('docReducer', () => {
  it('creates chapters from plan', () => {
    const doc = initialDoc({ title: 'Т', kind: 'srs', pages: 8 });
    const next = docReducer(doc, { type: 'SET_PLAN', plan: [{ heading: 'ВВЕДЕНИЕ', level: 1 }] });
    expect(next.chapters).toHaveLength(1);
    expect(next.chapters[0].paragraphs).toEqual([]);
    expect(next.chapters[0].id).toBeTruthy();
  });
  it('appends streamed text to a chapter', () => {
    const doc = docReducer(initialDoc({ title: 'Т', kind: 'srs', pages: 8 }),
      { type: 'SET_PLAN', plan: [{ heading: 'Г', level: 1 }] });
    const id = doc.chapters[0].id;
    const a = docReducer(doc, { type: 'APPEND_CHAPTER_TEXT', chapterId: id, text: 'При' });
    const b = docReducer(a, { type: 'APPEND_CHAPTER_TEXT', chapterId: id, text: 'вет' });
    expect(b.chapters[0].paragraphs).toEqual(['Привет']);
  });
  it('moves chapters', () => {
    let doc = docReducer(initialDoc({ title: 'Т', kind: 'srs', pages: 8 }),
      { type: 'SET_PLAN', plan: [{ heading: 'A', level: 1 }, { heading: 'B', level: 1 }] });
    doc = docReducer(doc, { type: 'MOVE_CHAPTER', from: 0, to: 1 });
    expect(doc.chapters.map((c) => c.heading)).toEqual(['B', 'A']);
  });
});

describe('docsReducer', () => {
  it('creates, selects and deletes documents', () => {
    let s = docsReducer({ documents: [], currentId: null }, { type: 'CREATE', title: 'Т', kind: 'srs', pages: 5 });
    expect(s.documents).toHaveLength(1);
    const id = s.currentId;
    s = docsReducer(s, { type: 'DELETE', id });
    expect(s.documents).toHaveLength(0);
    expect(s.currentId).toBeNull();
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npm test -- src/state/docReducer.test.js`
Expected: FAIL with "Cannot find module './docReducer.js'".

- [ ] **Step 3: Write minimal implementation**

```js
// src/state/docReducer.js
import { newId } from '../lib/storage.js';

export function initialDoc({ title, kind, pages }) {
  const now = Date.now();
  return { id: newId(), title, kind, pages, createdAt: now, updatedAt: now, chapters: [], bibliography: [] };
}

function touch(doc) {
  return { ...doc, updatedAt: Date.now() };
}

export function docReducer(doc, action) {
  switch (action.type) {
    case 'SET_META':
      return touch({ ...doc, title: action.title ?? doc.title, kind: action.kind ?? doc.kind, pages: action.pages ?? doc.pages });
    case 'SET_PLAN':
      return touch({ ...doc, chapters: action.plan.map((p) => ({ id: newId(), heading: p.heading, level: p.level, paragraphs: [] })) });
    case 'SET_CHAPTER_TEXT':
      return touch({ ...doc, chapters: doc.chapters.map((c) => (c.id === action.chapterId ? { ...c, paragraphs: [action.text] } : c)) });
    case 'APPEND_CHAPTER_TEXT':
      return touch({
        ...doc,
        chapters: doc.chapters.map((c) => {
          if (c.id !== action.chapterId) return c;
          const parts = [...c.paragraphs];
          parts[parts.length - 1] = `${parts[parts.length - 1] ?? ''}${action.text}`;
          return { ...c, paragraphs: parts.length ? parts : [action.text] };
        }),
      });
    case 'UPDATE_CHAPTER':
      return touch({ ...doc, chapters: doc.chapters.map((c) => (c.id === action.chapterId ? { ...c, ...action.patch } : c)) });
    case 'ADD_CHAPTER':
      return touch({ ...doc, chapters: [...doc.chapters, { id: newId(), heading: action.heading || 'Новая глава', level: action.level || 1, paragraphs: [] }] });
    case 'REMOVE_CHAPTER':
      return touch({ ...doc, chapters: doc.chapters.filter((c) => c.id !== action.chapterId) });
    case 'MOVE_CHAPTER': {
      const list = [...doc.chapters];
      const [moved] = list.splice(action.from, 1);
      list.splice(action.to, 0, moved);
      return touch({ ...doc, chapters: list });
    }
    case 'SET_BIBLIO':
      return touch({ ...doc, bibliography: action.items });
    default:
      return doc;
  }
}

export function docsReducer(state, action) {
  switch (action.type) {
    case 'HYDRATE':
      return { documents: action.documents, currentId: action.currentId };
    case 'CREATE': {
      const doc = initialDoc(action);
      return { documents: [doc, ...state.documents], currentId: doc.id };
    }
    case 'SELECT':
      return { ...state, currentId: action.id };
    case 'DELETE':
      return { documents: state.documents.filter((d) => d.id !== action.id), currentId: state.currentId === action.id ? null : state.currentId };
    case 'UPDATE_DOC':
      return { ...state, documents: state.documents.map((d) => (d.id === action.doc.id ? action.doc : d)) };
    default:
      return state;
  }
}
```

- [ ] **Step 4: Write store.jsx (thin wiring, no test)**

```jsx
// src/state/store.jsx
import { createContext, useContext, useEffect, useReducer } from 'react';
import { loadDocuments, saveDocuments } from '../lib/storage.js';
import { docsReducer } from './docReducer.js';

const DocsContext = createContext(null);

export function DocsProvider({ children }) {
  const [state, dispatch] = useReducer(docsReducer, { documents: [], currentId: null });
  useEffect(() => {
    dispatch({ type: 'HYDRATE', documents: loadDocuments(), currentId: null });
  }, []);
  useEffect(() => {
    saveDocuments(state.documents);
  }, [state.documents]);
  return <DocsContext.Provider value={{ state, dispatch }}>{children}</DocsContext.Provider>;
}

export function useDocs() {
  const ctx = useContext(DocsContext);
  if (!ctx) throw new Error('useDocs вне DocsProvider');
  return ctx;
}
```

- [ ] **Step 5: Run tests to verify they pass**

Run: `npm test -- src/state/docReducer.test.js`
Expected: PASS, 4 tests.

- [ ] **Step 6: Commit**

```bash
git add src/state/docReducer.js src/state/docReducer.test.js src/state/store.jsx
git commit -m "feat: add document reducer and store"
```

---

### Task 7: GOST docx builder

**Files:**
- Create: `src/lib/docxBuilder.js`, `src/lib/docxBuilder.test.js`

**Interfaces:**
- Consumes: `GOST` from Task 2, doc shapes from Task 6.
- Produces: `buildGostDocument(doc)`, `docxBlob(doc)`, `downloadDocx(blob, filename)` — used by Task 9. Test uses `Packer.toBuffer` + `jszip` (devDep from Task 1).

- [ ] **Step 1: Write the failing test**

```js
// src/lib/docxBuilder.test.js
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npm test -- src/lib/docxBuilder.test.js`
Expected: FAIL with "Cannot find module './docxBuilder.js'".

- [ ] **Step 3: Write minimal implementation**

```js
// src/lib/docxBuilder.js
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
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `npm test -- src/lib/docxBuilder.test.js`
Expected: PASS, 2 tests.

- [ ] **Step 5: Commit**

```bash
git add src/lib/docxBuilder.js src/lib/docxBuilder.test.js
git commit -m "feat: add GOST docx builder"
```

---

### Task 8: List, form, settings UI

**Files:**
- Create: `src/components/WorksList.jsx`, `src/components/NewWorkForm.jsx`, `src/components/SettingsModal.jsx`

**Interfaces:**
- Consumes: `useDocs` from Task 6, `PRESETS` + `loadSettings/saveSettings` from Tasks 2–3.
- Produces: three presentational components with callback props — used by Task 9. Verified via build + manual checklist (no unit tests for JSX).

- [ ] **Step 1: Write WorksList.jsx**

```jsx
// src/components/WorksList.jsx
export default function WorksList({ documents, onOpen, onDelete, onNew }) {
  return (
    <section>
      <h2>Мои работы</h2>
      <button type="button" onClick={onNew}>Новая работа</button>
      {documents.length === 0 ? (
        <p>Пока пусто. Создайте первую работу — это займёт минуту.</p>
      ) : (
        <ul className="works">
          {documents.map((d) => (
            <li key={d.id} className="work-card">
              <button type="button" className="work-open" onClick={() => onOpen(d.id)}>
                {d.title} — {d.kind === 'practice' ? 'отчет' : 'СРС'}, {d.pages} стр., глав: {d.chapters.length}
              </button>
              <button type="button" onClick={() => onDelete(d.id)}>Удалить</button>
            </li>
          ))}
        </ul>
      )}
    </section>
  );
}
```

- [ ] **Step 2: Write NewWorkForm.jsx**

```jsx
// src/components/NewWorkForm.jsx
import { useState } from 'react';

export default function NewWorkForm({ onCreate, onCancel }) {
  const [title, setTitle] = useState('');
  const [kind, setKind] = useState('srs');
  const [pages, setPages] = useState(8);
  const valid = title.trim().length > 1 && Number(pages) >= 1 && Number(pages) <= 100;
  return (
    <form onSubmit={(e) => { e.preventDefault(); if (valid) onCreate({ title: title.trim(), kind, pages: Number(pages) }); }}>
      <h2>Новая работа</h2>
      <label>Тема<input value={title} onChange={(e) => setTitle(e.target.value)} placeholder="Например: Искусственный интеллект в образовании" /></label>
      <label>Тип
        <select value={kind} onChange={(e) => setKind(e.target.value)}>
          <option value="srs">СРС / реферат</option>
          <option value="practice">Отчет по практике</option>
        </select>
      </label>
      <label>Страниц<input type="number" min="1" max="100" value={pages} onChange={(e) => setPages(e.target.value)} /></label>
      <button type="submit" disabled={!valid}>Создать</button>
      <button type="button" onClick={onCancel}>Назад</button>
    </form>
  );
}
```

- [ ] **Step 3: Write SettingsModal.jsx**

```jsx
// src/components/SettingsModal.jsx
import { useState } from 'react';
import { PRESETS } from '../lib/llm.js';

export default function SettingsModal({ initial, onSave, onClose }) {
  const [baseUrl, setBaseUrl] = useState(initial.baseUrl || '');
  const [apiKey, setApiKey] = useState(initial.apiKey || '');
  const [model, setModel] = useState(initial.model || '');
  return (
    <div className="modal">
      <h2>Настройки ИИ</h2>
      <p className="warn">Ключ хранится только в этом браузере. Не вводите ключ на чужом компьютере.</p>
      <label>Провайдер
        <select value={PRESETS.find((p) => p.baseUrl === baseUrl)?.id || 'custom'} onChange={(e) => {
          const p = PRESETS.find((x) => x.id === e.target.value);
          setBaseUrl(p.baseUrl); if (p.model) setModel(p.model);
        }}>
          {PRESETS.map((p) => <option key={p.id} value={p.id}>{p.label}</option>)}
        </select>
      </label>
      <label>Base URL<input value={baseUrl} onChange={(e) => setBaseUrl(e.target.value)} placeholder="https://api.openai.com/v1" /></label>
      <label>Ключ<input type="password" value={apiKey} onChange={(e) => setApiKey(e.target.value)} /></label>
      <label>Модель<input value={model} onChange={(e) => setModel(e.target.value)} placeholder="gpt-4o-mini" /></label>
      <button type="button" onClick={() => { onSave({ baseUrl, apiKey, model }); onClose(); }}>Сохранить</button>
      <button type="button" onClick={onClose}>Закрыть</button>
    </div>
  );
}
```

- [ ] **Step 4: Verify build still passes**

Run: `npm run build`
Expected: success (components unused yet — no import errors).

- [ ] **Step 5: Commit**

```bash
git add src/components/WorksList.jsx src/components/NewWorkForm.jsx src/components/SettingsModal.jsx
git commit -m "feat: add list, form and settings UI"
```

---

### Task 9: Editor + App wiring

**Files:**
- Create: `src/components/ChapterView.jsx`, `src/components/Editor.jsx`
- Modify: `src/App.jsx` (replace Task 1 placeholder with router + DocsProvider)

**Interfaces:**
- Consumes: everything from Tasks 2–8.
- Produces: working app. Verified via build + manual checklist (no unit tests for JSX).

- [ ] **Step 1: Write ChapterView.jsx**

```jsx
// src/components/ChapterView.jsx
export default function ChapterView({ chapter, onHeading, onText, onWriteOne, writing, canWrite }) {
  return (
    <article className="chapter">
      <input value={chapter.heading} onChange={(e) => onHeading(chapter.id, e.target.value)} aria-label="Заголовок главы" />
      {(chapter.paragraphs.length ? chapter.paragraphs : ['']).map((p, i) => (
        <textarea key={i} value={p} rows={4}
          onChange={(e) => onText(chapter.id, i, e.target.value)} aria-label={`Абзац ${i + 1}`} />
      ))}
      <button type="button" disabled={!canWrite || writing} onClick={() => onWriteOne(chapter.id)}>
        {writing ? 'Пишу…' : 'Написать эту главу'}
      </button>
    </article>
  );
}
```

- [ ] **Step 2: Write Editor.jsx**

```jsx
// src/components/Editor.jsx
import { useState } from 'react';
import { streamChat } from '../lib/llm.js';
import { buildBiblioPrompt, buildChapterPrompt, buildPlanPrompt, countPages, extractPlanArray, sanitizeMarkdown } from '../lib/outline.js';
import { canBiblio, canDownload, canWriteText, makeFilename, settingsReady } from '../lib/validate.js';
import { downloadDocx, docxBlob } from '../lib/docxBuilder.js';
import { docReducer } from '../state/docReducer.js';
import ChapterView from './ChapterView.jsx';
import SettingsModal from './SettingsModal.jsx';

const WORDS_PER_CHAPTER = 350;

export default function Editor({ doc, settings, onDoc, onSettings, onBack }) {
  const [status, setStatus] = useState('');
  const [writing, setWriting] = useState(false);
  const [showSettings, setShowSettings] = useState(false);
  const ready = settingsReady(settings);

  async function runPlan() {
    setWriting(true); setStatus('Составляю план…');
    try {
      const raw = await streamChat({ ...settings, messages: [{ role: 'user', content: buildPlanPrompt(doc) }] });
      onDoc(docReducer(doc, { type: 'SET_PLAN', plan: extractPlanArray(raw) }));
      setStatus('План готов. Проверьте главы и нажмите «План принят», либо правьте.');
    } catch (e) { setStatus(e.status === 401 ? 'Неверный ключ. Проверьте настройки.' : `Не вышло: ${e.message}`); }
    finally { setWriting(false); }
  }

  async function writeChapter(id) {
    const ch = doc.chapters.find((c) => c.id === id);
    if (!ch) return;
    setWriting(true); setStatus(`Пишу: ${ch.heading}`);
    try {
      const prompt = buildChapterPrompt(doc.title, ch, WORDS_PER_CHAPTER);
      const full = await streamChat({ ...settings, messages: [{ role: 'user', content: prompt }],
        onToken: (t) => onDoc((d) => docReducer(d, { type: 'APPEND_CHAPTER_TEXT', chapterId: id, text: t })) });
      onDoc((d) => docReducer(d, { type: 'SET_CHAPTER_TEXT', chapterId: id, text: sanitizeMarkdown(full) }));
      setStatus('Глава готова.');
    } catch (e) { setStatus(e.status === 401 ? 'Неверный ключ. Проверьте настройки.' : `Остановилось: ${e.message}. Написанное сохранено, можно повторить.`); }
    finally { setWriting(false); }
  }

  async function writeAll() {
    for (const ch of doc.chapters) {
      const current = doc.chapters.find((c) => c.id === ch.id);
      if (current && current.paragraphs.join('').trim()) continue;
      await writeChapter(ch.id);
    }
  }

  async function runBiblio() {
    setWriting(true); setStatus('Собираю литературу…');
    try {
      const raw = await streamChat({ ...settings, messages: [{ role: 'user', content: buildBiblioPrompt(doc.title, doc.chapters) }] });
      const items = sanitizeMarkdown(raw).split('\n').map((s) => s.trim()).filter(Boolean);
      onDoc(docReducer(doc, { type: 'SET_BIBLIO', items }));
      setStatus('Литература добавлена, проверьте строки.');
    } catch (e) { setStatus(`Не вышло: ${e.message}`); }
    finally { setWriting(false); }
  }

  async function runDownload() {
    try {
      const blob = await docxBlob(doc);
      downloadDocx(blob, makeFilename(doc));
      setStatus('Файл скачан. В Word: правый клик по содержанию → Обновить поле.');
    } catch (e) { setStatus(`Не удалось собрать файл: ${e.message}`); }
  }

  const pages = countPages(doc.chapters.flatMap((c) => c.paragraphs).join('\n'));
  return (
    <section>
      <button type="button" onClick={onBack}>← Мои работы</button>
      <h2>{doc.title} ({doc.kind === 'practice' ? 'отчет' : 'СРС'}, нужно {doc.pages} стр., сейчас ~{pages.toFixed(1)})</h2>
      {!ready && <p className="warn">Укажите ИИ-ключ в настройках, иначе генерация недоступна. <button type="button" onClick={() => setShowSettings(true)}>Настройки</button></p>}
      <div className="toolbar">
        <button type="button" disabled={!ready || writing} onClick={runPlan}>Составить план</button>
        <button type="button" disabled={!ready || !canWriteText(doc) || writing} onClick={writeAll}>Написать текст</button>
        <button type="button" disabled={!ready || !canBiblio(doc) || writing} onClick={runBiblio}>Литература</button>
        <button type="button" disabled={!canDownload(doc)} onClick={runDownload}>Скачать .docx</button>
        <button type="button" onClick={() => setShowSettings(true)}>Настройки</button>
      </div>
      {status && <p role="status">{status}</p>}
      <div className="editor">
        <nav className="toc">{doc.chapters.map((c) => <div key={c.id}>{c.heading}</div>)}</nav>
        <div className="chapters">
          {doc.chapters.map((c) => (
            <ChapterView key={c.id} chapter={c} writing={writing} canWrite={ready}
              onHeading={(id, heading) => onDoc(docReducer(doc, { type: 'UPDATE_CHAPTER', chapterId: id, patch: { heading } }))}
              onText={(id, i, value) => onDoc(docReducer(doc, {
                type: 'UPDATE_CHAPTER', chapterId: id,
                patch: { paragraphs: doc.chapters.find((x) => x.id === id).paragraphs.map((p, j) => (j === i ? value : p)) },
              }))}
              onWriteOne={writeChapter} />
          ))}
        </div>
      </div>
      {showSettings && <SettingsModal initial={settings} onSave={onSettings} onClose={() => setShowSettings(false)} />}
    </section>
  );
}
```

- [ ] **Step 3: Rewrite App.jsx with router**

```jsx
// src/App.jsx
import { useState } from 'react';
import { DocsProvider, useDocs } from './state/store.jsx';
import { loadSettings, saveSettings } from './lib/storage.js';
import WorksList from './components/WorksList.jsx';
import NewWorkForm from './components/NewWorkForm.jsx';
import Editor from './components/Editor.jsx';

function Shell() {
  const { state, dispatch } = useDocs();
  const [screen, setScreen] = useState('list');
  const [settings, setSettings] = useState(loadSettings);
  const current = state.documents.find((d) => d.id === state.currentId);

  function save(s) { setSettings(s); saveSettings(s); }
  function updateDoc(next) {
    const doc = typeof next === 'function' ? next(state.documents.find((d) => d.id === state.currentId)) : next;
    dispatch({ type: 'UPDATE_DOC', doc });
  }

  if (screen === 'new') {
    return <NewWorkForm onCreate={(m) => { dispatch({ type: 'CREATE', ...m }); setScreen('editor'); }} onCancel={() => setScreen('list')} />;
  }
  if (screen === 'editor' && current) {
    return <Editor doc={current} settings={settings} onDoc={updateDoc} onSettings={save} onBack={() => setScreen('list')} />;
  }
  return (
    <main className="app">
      <h1>SRS Platform</h1>
      <WorksList documents={state.documents}
        onOpen={(id) => { dispatch({ type: 'SELECT', id }); setScreen('editor'); }}
        onDelete={(id) => dispatch({ type: 'DELETE', id })}
        onNew={() => setScreen('new')} />
    </main>
  );
}

export default function App() {
  return <DocsProvider><Shell /></DocsProvider>;
}
```

- [ ] **Step 4: Run full test suite and build**

```bash
npm test
npm run build
```

Expected: all tests PASS; build outputs `dist/index.html` with no errors.

- [ ] **Step 5: Commit**

```bash
git add src/components/ChapterView.jsx src/components/Editor.jsx src/App.jsx
git commit -m "feat: add editor and app wiring"
```

---

### Task 10: Manual checklist + deploy

**Files:** none (verification only).

- [ ] **Step 1: Run the checklist from the spec**
  - `npm test` — все тесты зеленые.
  - `npm run dev` — создать работу, сгенерировать план и 1 главу (нужен личный ключ), поправить текст, скачать docx.
  - Открыть docx в Word/LibreOffice: поля 30/10/20/20, Times 14, интервал 1.5, содержание обновляется, титульника нет, подсказок внутри нет.
  - Узкий экран (375px): колонки не ломаются.
- [ ] **Step 2: Commit any fixes** (если чек-лист нашел баги — чинить отдельными коммитами).
- [ ] **Step 3: Push**

```bash
git push origin main
```

- [ ] **Step 4: Deploy**

```bash
npx vercel --prod
```

Expected: production URL отвечает, приложение открывается. `VERCEL_TOKEN` не задан — вход интерактивный, держать терминал открытым.

- [ ] **Step 5: Verify production**

Открыть production URL, создать тестовую работу без ключа (кнопки ИИ неактивны с подсказкой), проверить скачивание невозможно без текста — по матрице валидации из спеки.
