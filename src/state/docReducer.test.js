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
