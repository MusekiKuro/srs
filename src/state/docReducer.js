import { newId } from '../lib/storage.js';

export function initialDoc({ title, kind, pages }) {
  const now = Date.now();
  return { id: newId(), title, kind, pages, createdAt: now, updatedAt: now, chapters: [], bibliography: [] };
}

function touch(doc) {
  return { ...doc, updatedAt: Date.now() };
}

export function docReducer(doc, action) {
  switch (action.type) {
    case 'SET_META':
      return touch({ ...doc, title: action.title ?? doc.title, kind: action.kind ?? doc.kind, pages: action.pages ?? doc.pages });
    case 'SET_PLAN':
      return touch({ ...doc, chapters: action.plan.map((p) => ({ id: newId(), heading: p.heading, level: p.level, paragraphs: [] })) });
    case 'SET_CHAPTER_TEXT':
      return touch({ ...doc, chapters: doc.chapters.map((c) => (c.id === action.chapterId ? { ...c, paragraphs: [action.text] } : c)) });
    case 'APPEND_CHAPTER_TEXT':
      return touch({
        ...doc,
        chapters: doc.chapters.map((c) => {
          if (c.id !== action.chapterId) return c;
          const parts = [...c.paragraphs];
          parts[parts.length - 1] = `${parts[parts.length - 1] ?? ''}${action.text}`;
          return { ...c, paragraphs: parts.length ? parts : [action.text] };
        }),
      });
    case 'UPDATE_CHAPTER':
      return touch({ ...doc, chapters: doc.chapters.map((c) => (c.id === action.chapterId ? { ...c, ...action.patch } : c)) });
    case 'ADD_CHAPTER':
      return touch({ ...doc, chapters: [...doc.chapters, { id: newId(), heading: action.heading || 'Новая глава', level: action.level || 1, paragraphs: [] }] });
    case 'REMOVE_CHAPTER':
      return touch({ ...doc, chapters: doc.chapters.filter((c) => c.id !== action.chapterId) });
    case 'MOVE_CHAPTER': {
      const list = [...doc.chapters];
      const [moved] = list.splice(action.from, 1);
      list.splice(action.to, 0, moved);
      return touch({ ...doc, chapters: list });
    }
    case 'SET_BIBLIO':
      return touch({ ...doc, bibliography: action.items });
    default:
      return doc;
  }
}

export function docsReducer(state, action) {
  switch (action.type) {
    case 'HYDRATE':
      return { documents: action.documents, currentId: action.currentId };
    case 'CREATE': {
      const doc = initialDoc(action);
      return { documents: [doc, ...state.documents], currentId: doc.id };
    }
    case 'SELECT':
      return { ...state, currentId: action.id };
    case 'DELETE':
      return { documents: state.documents.filter((d) => d.id !== action.id), currentId: state.currentId === action.id ? null : state.currentId };
    case 'UPDATE_DOC':
      return { ...state, documents: state.documents.map((d) => (d.id === action.doc.id ? action.doc : d)) };
    default:
      return state;
  }
}
