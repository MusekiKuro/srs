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
