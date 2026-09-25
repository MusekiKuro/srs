import { describe, expect, it } from 'vitest';
import { GOST, mmToTwips } from './gost.js';

describe('gost', () => {
  it('converts mm to twips', () => {
    expect(mmToTwips(30)).toBe(1701);
    expect(mmToTwips(10)).toBe(567);
  });
  it('holds exact GOST values', () => {
    expect(GOST.font).toBe('Times New Roman');
    expect(GOST.bodySize).toBe(28);
    expect(GOST.margins).toEqual({ top: 1134, right: 567, bottom: 1134, left: 1701 });
    expect(GOST.firstLineIndent).toBe(709);
    expect(GOST.lineOneHalf).toBe(360);
  });
});
