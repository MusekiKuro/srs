import { useState } from 'react';
import { streamChat } from '../lib/llm.js';
import { buildBiblioPrompt, buildChapterPrompt, buildPlanPrompt, countPages, extractPlanArray, sanitizeMarkdown } from '../lib/outline.js';
import { canBiblio, canDownload, canWriteText, makeFilename, settingsReady } from '../lib/validate.js';
import { downloadDocx, docxBlob } from '../lib/docxBuilder.js';
import { docReducer } from '../state/docReducer.js';
import ChapterView from './ChapterView.jsx';
import SettingsModal from './SettingsModal.jsx';

const WORDS_PER_CHAPTER = 350;

function setParagraph(paragraphs, i, value) {
  const next = [...paragraphs];
  while (next.length <= i) next.push('');
  next[i] = value;
  return next;
}

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
      setStatus('План готов. Проверьте главы и правьте при необходимости.');
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
      if ((ch.paragraphs ?? []).join('').trim()) continue;
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
                patch: { paragraphs: setParagraph(doc.chapters.find((x) => x.id === id).paragraphs, i, value) },
              }))}
              onWriteOne={writeChapter} />
          ))}
        </div>
      </div>
      {showSettings && <SettingsModal initial={settings} onSave={onSettings} onClose={() => setShowSettings(false)} />}
    </section>
  );
}
