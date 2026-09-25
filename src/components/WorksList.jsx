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
