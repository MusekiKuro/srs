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
