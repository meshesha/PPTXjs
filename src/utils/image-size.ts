/**
 * Reads the pixel size of PNG, JPEG, GIF, BMP and WebP images from their bytes.
 * Synchronous and DOM-free (replaces v1's busy-wait on `new Image()`, which
 * could freeze the page).
 */
export function imageSize(bytes: Uint8Array): [number, number] | null {
  const b = bytes;
  const u16be = (o: number) => (b[o] << 8) | b[o + 1];
  const u16le = (o: number) => b[o] | (b[o + 1] << 8);
  const u32be = (o: number) => ((b[o] << 24) | (b[o + 1] << 16) | (b[o + 2] << 8) | b[o + 3]) >>> 0;
  const u32le = (o: number) => (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0;
  if (b.length < 24) return null;
  // PNG
  if (b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) return [u32be(16), u32be(20)];
  // GIF
  if (b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46) return [u16le(6), u16le(8)];
  // BMP
  if (b[0] === 0x42 && b[1] === 0x4d) return [u32le(18), Math.abs(u32le(22) | 0)];
  // WebP
  if (b[0] === 0x52 && b[1] === 0x49 && b[8] === 0x57 && b[9] === 0x45) {
    const chunk = String.fromCharCode(b[12], b[13], b[14], b[15]);
    if (chunk === "VP8 ") return [u16le(26) & 0x3fff, u16le(28) & 0x3fff];
    if (chunk === "VP8L") {
      const bits = u32le(21);
      return [(bits & 0x3fff) + 1, ((bits >> 14) & 0x3fff) + 1];
    }
    if (chunk === "VP8X") return [1 + (b[24] | (b[25] << 8) | (b[26] << 16)), 1 + (b[27] | (b[28] << 8) | (b[29] << 16))];
  }
  // JPEG: walk the segments to the SOFn marker
  if (b[0] === 0xff && b[1] === 0xd8) {
    let o = 2;
    while (o + 9 < b.length) {
      if (b[o] !== 0xff) {
        o++;
        continue;
      }
      const marker = b[o + 1];
      if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
        return [u16be(o + 7), u16be(o + 5)];
      }
      o += 2 + u16be(o + 2);
    }
  }
  return null;
}

/** Size of an image given as a `data:` URL, or `[0, 0]` when unknown. */
export function imageSizeFromDataUrl(dataUrl: string): [number, number] {
  const comma = dataUrl.indexOf(",");
  if (!dataUrl.startsWith("data:") || comma === -1 || !/;base64$/.test(dataUrl.slice(0, comma))) return [0, 0];
  // only the first ~64KB are needed for the header
  const b64 = dataUrl.slice(comma + 1, comma + 1 + 87384);
  let bin: string;
  try {
    bin = atob(b64.slice(0, b64.length - (b64.length % 4)));
  } catch {
    return [0, 0];
  }
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return imageSize(bytes) ?? [0, 0];
}
