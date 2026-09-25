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
