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
