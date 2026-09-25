import { describe, expect, it, vi } from 'vitest';
import { chatUrl, normalizeBaseUrl, streamChat } from './llm.js';

function sseResponse(chunks) {
  const encoder = new TextEncoder();
  const payload = chunks.map((c) => `data: ${JSON.stringify(c)}\n\n`).join('') + 'data: [DONE]\n\n';
  return new Response(encoder.encode(payload));
}

describe('llm', () => {
  it('normalizes base url', () => {
    expect(normalizeBaseUrl('https://x/v1///')).toBe('https://x/v1');
    expect(chatUrl('https://x/v1/')).toBe('https://x/v1/chat/completions');
    expect(chatUrl('https://x/v1')).toBe('https://x/v1/chat/completions');
  });
  it('streams tokens and returns full text', async () => {
    globalThis.fetch = vi.fn().mockResolvedValue(sseResponse([
      { choices: [{ delta: { content: 'При' } }] },
      { choices: [{ delta: { content: 'вет' } }] },
    ]));
    const seen = [];
    const full = await streamChat({
      baseUrl: 'https://x/v1', apiKey: 'k', model: 'm',
      messages: [{ role: 'user', content: 'hi' }],
      onToken: (t) => seen.push(t),
    });
    expect(full).toBe('Привет');
    expect(seen).toEqual(['При', 'вет']);
  });
  it('throws status error on 401', async () => {
    globalThis.fetch = vi.fn().mockResolvedValue(new Response('no', { status: 401 }));
    await expect(streamChat({ baseUrl: 'https://x/v1', apiKey: 'bad', model: 'm', messages: [] }))
      .rejects.toMatchObject({ status: 401 });
  });
  it('keeps partial text when stream breaks', async () => {
    const encoder = new TextEncoder();
    const broken = encoder.encode('data: {"choices":[{"delta":{"content":"Часть"}}]}\n\ndata: BROKEN{{{');
    globalThis.fetch = vi.fn().mockResolvedValue(new Response(broken));
    const seen = [];
    await expect(streamChat({
      baseUrl: 'https://x/v1', apiKey: 'k', model: 'm', messages: [],
      onToken: (t) => seen.push(t),
    })).rejects.toThrow();
    expect(seen).toEqual(['Часть']);
  });
});
