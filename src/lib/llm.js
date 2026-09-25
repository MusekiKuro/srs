export const PRESETS = [
  { id: 'openai', label: 'OpenAI', baseUrl: 'https://api.openai.com/v1', model: 'gpt-4o-mini' },
  { id: 'openrouter', label: 'OpenRouter', baseUrl: 'https://openrouter.ai/api/v1', model: 'openai/gpt-4o-mini' },
  { id: 'deepseek', label: 'DeepSeek', baseUrl: 'https://api.deepseek.com/v1', model: 'deepseek-chat' },
  { id: 'groq', label: 'Groq', baseUrl: 'https://api.groq.com/openai/v1', model: 'llama-3.1-8b-instant' },
  { id: 'custom', label: 'Свой URL', baseUrl: '', model: '' },
];

export function normalizeBaseUrl(url) {
  return String(url ?? '').trim().replace(/\/+$/, '');
}

export function chatUrl(baseUrl) {
  return `${normalizeBaseUrl(baseUrl)}/chat/completions`;
}

export async function streamChat({ baseUrl, apiKey, model, messages, onToken, signal }) {
  const res = await fetch(chatUrl(baseUrl), {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${apiKey}` },
    body: JSON.stringify({ model, messages, stream: true }),
    signal,
  });
  if (!res.ok) {
    const err = new Error(res.status === 401 ? 'Неверный ключ. Проверьте ключ в настройках.' : `Сервер вернул ${res.status}.`);
    err.status = res.status;
    throw err;
  }
  const reader = res.body.getReader();
  const decoder = new TextDecoder();
  let buf = '';
  let full = '';
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    buf += decoder.decode(value, { stream: true });    const parts = buf.split('\n\n');
    buf = parts.pop();
    for (const part of parts) {
      const line = part.trim();
      if (!line.startsWith('data:')) continue;
      const payload = line.slice(5).trim();
      if (payload === '[DONE]') continue;
      const json = JSON.parse(payload);
      const piece = json?.choices?.[0]?.delta?.content ?? '';
      if (piece) { full += piece; onToken?.(piece); }
    }
  }
  if (buf.trim() && buf.trim() !== 'data: [DONE]') {
    throw new Error('Поток оборвался, ответ неполный.');
  }
  return full;
}
