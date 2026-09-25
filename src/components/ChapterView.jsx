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
