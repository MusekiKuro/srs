import { createContext, useContext, useEffect, useReducer } from 'react';
import { loadDocuments, saveDocuments } from '../lib/storage.js';
import { docsReducer } from './docReducer.js';

const DocsContext = createContext(null);

export function DocsProvider({ children }) {
  const [state, dispatch] = useReducer(docsReducer, { documents: [], currentId: null });
  useEffect(() => {
    dispatch({ type: 'HYDRATE', documents: loadDocuments(), currentId: null });
  }, []);
  useEffect(() => {
    saveDocuments(state.documents);
  }, [state.documents]);
  return <DocsContext.Provider value={{ state, dispatch }}>{children}</DocsContext.Provider>;
}

export function useDocs() {
  const ctx = useContext(DocsContext);
  if (!ctx) throw new Error('useDocs вне DocsProvider');
  return ctx;
}
