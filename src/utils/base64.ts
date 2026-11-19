const CHARS = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";

/** Base64-encodes binary data (works in browsers, workers and Node). */
export function base64ArrayBuffer(input: ArrayBuffer | Uint8Array | null | undefined): string {
  if (!input) return "";
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
  let out = "";
  const len = bytes.length;
  const rem = len % 3;
  const main = len - rem;
  for (let i = 0; i < main; i += 3) {
    const n = (bytes[i] << 16) | (bytes[i + 1] << 8) | bytes[i + 2];
    out += CHARS[(n >> 18) & 63] + CHARS[(n >> 12) & 63] + CHARS[(n >> 6) & 63] + CHARS[n & 63];
  }
  if (rem === 1) {
    const n = bytes[main];
    out += CHARS[n >> 2] + CHARS[(n & 3) << 4] + "==";
  } else if (rem === 2) {
    const n = (bytes[main] << 8) | bytes[main + 1];
    out += CHARS[n >> 10] + CHARS[(n >> 4) & 63] + CHARS[(n & 15) << 2] + "=";
  }
  return out;
}
