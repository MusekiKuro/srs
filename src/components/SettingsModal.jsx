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
