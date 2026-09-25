export const MM_TO_TWIPS = 1440 / 25.4;

export function mmToTwips(mm) {
  return Math.round(mm * MM_TO_TWIPS);
}

export const GOST = {
  font: 'Times New Roman',
  black: '000000',
  bodySize: 28, // 14pt in half-points
  footerSize: 24, // 12pt in half-points
  lineSingle: 240,
  lineOneHalf: 360, // 1.5 spacing with AUTO rule
  firstLineIndent: 709, // 1.25cm in twips
  margins: { top: 1134, right: 567, bottom: 1134, left: 1701 }, // 20/10/20/30 mm
  pageSize: { width: 11906, height: 16838 }, // A4 in twips
  wordsPerPage: 325,
};
