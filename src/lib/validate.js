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
