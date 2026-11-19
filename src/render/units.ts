/** 1 inch = 914400 EMU = 96 CSS px. */
export const EMU_PER_PX = 914400 / 96;

/** Multiply an EMU value by this to get CSS pixels. */
export const slideFactor = 96 / 914400;

/** Point to pixel factor used for font sizes (kept from v1 for identical output). */
export const fontSizeFactor = 4 / 3.2;

/** Languages rendered right-to-left. */
export const rtl_langs_array = ["he-IL", "ar-AE", "ar-SA", "dv-MV", "fa-IR", "ur-PK"];

/** Converts an OOXML angle (60000ths of a degree) to whole degrees. */
export function angleToDegrees(angle: unknown): number {
  if (angle === "" || angle === null || angle === undefined) return 0;
  return Math.round(Number(angle) / 60000);
}

export function emuToPx(emu: number | string): number {
  return Number(emu) * slideFactor;
}

export function pxToEmu(px: number): number {
  return Math.round(px * EMU_PER_PX);
}
