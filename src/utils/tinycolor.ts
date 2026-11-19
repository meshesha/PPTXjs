/**
 * The subset of TinyColor v1.4.2 used by the renderer, as a dependency-free
 * TypeScript module. The parsing and rounding rules follow TinyColor exactly so
 * colours match PPTXjs v1 output byte for byte.
 *
 * TinyColor: https://github.com/bgrins/TinyColor - Brian Grinstead, MIT License
 */
type Num = number | string;

export interface Rgba {
  r: number;
  g: number;
  b: number;
  a: number;
}
export interface Hsla {
  h: number;
  s: number;
  l: number;
  a: number;
}
type ColorInput = string | Partial<{ r: Num; g: Num; b: Num; h: Num; s: Num; l: Num; a: Num }> | TinyColor | undefined | null;

const round = Math.round;

function isOnePointZero(n: Num): boolean {
  return typeof n === "string" && n.indexOf(".") !== -1 && parseFloat(n) === 1;
}
function isPercentage(n: Num): boolean {
  return typeof n === "string" && n.indexOf("%") !== -1;
}
/** Takes input from [0, max] (or a percentage string) and returns it as [0, 1]. */
function bound01(n: Num, max: number): number {
  if (isOnePointZero(n)) n = "100%";
  const processPercent = isPercentage(n);
  let v = Math.min(max, Math.max(0, parseFloat(String(n))));
  if (processPercent) v = parseInt(String(v * max), 10) / 100;
  if (Math.abs(v - max) < 0.000001) return 1;
  return (v % max) / max;
}
function clamp01(v: number): number {
  return Math.min(1, Math.max(0, v));
}
function boundAlpha(a: Num | undefined): number {
  let v = parseFloat(String(a));
  if (isNaN(v) || v < 0 || v > 1) v = 1;
  return v;
}
function convertToPercentage(n: Num): Num {
  if (Number(n) <= 1) return Number(n) * 100 + "%";
  return n;
}
function pad2(c: string): string {
  return c.length === 1 ? "0" + c : c;
}

function rgbToHsl(r: number, g: number, b: number): { h: number; s: number; l: number } {
  r = bound01(r, 255);
  g = bound01(g, 255);
  b = bound01(b, 255);
  const max = Math.max(r, g, b),
    min = Math.min(r, g, b);
  let h = 0,
    s: number;
  const l = (max + min) / 2;
  if (max === min) {
    h = s = 0;
  } else {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    switch (max) {
      case r:
        h = (g - b) / d + (g < b ? 6 : 0);
        break;
      case g:
        h = (b - r) / d + 2;
        break;
      case b:
        h = (r - g) / d + 4;
        break;
    }
    h /= 6;
  }
  return { h, s, l };
}

function hue2rgb(p: number, q: number, t: number): number {
  if (t < 0) t += 1;
  if (t > 1) t -= 1;
  if (t < 1 / 6) return p + (q - p) * 6 * t;
  if (t < 1 / 2) return q;
  if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
  return p;
}

function hslToRgb(h: Num, s: Num, l: Num): { r: number; g: number; b: number } {
  const hh = bound01(h, 360),
    ss = bound01(s, 100),
    ll = bound01(l, 100);
  let r: number, g: number, b: number;
  if (ss === 0) {
    r = g = b = ll;
  } else {
    const q = ll < 0.5 ? ll * (1 + ss) : ll + ss - ll * ss;
    const p = 2 * ll - q;
    r = hue2rgb(p, q, hh + 1 / 3);
    g = hue2rgb(p, q, hh);
    b = hue2rgb(p, q, hh - 1 / 3);
  }
  return { r: r * 255, g: g * 255, b: b * 255 };
}

const CSS_UNIT = "(?:[-\\+]?\\d*\\.\\d+%?)|(?:[-\\+]?\\d+%?)";
const PERMISSIVE3 = `[\\s|\\(]+(${CSS_UNIT})[,|\\s]+(${CSS_UNIT})[,|\\s]+(${CSS_UNIT})\\s*\\)?`;
const PERMISSIVE4 = `[\\s|\\(]+(${CSS_UNIT})[,|\\s]+(${CSS_UNIT})[,|\\s]+(${CSS_UNIT})[,|\\s]+(${CSS_UNIT})\\s*\\)?`;
const MATCHERS = {
  CSS_UNIT: new RegExp(CSS_UNIT),
  rgb: new RegExp("rgb" + PERMISSIVE3),
  rgba: new RegExp("rgba" + PERMISSIVE4),
  hsl: new RegExp("hsl" + PERMISSIVE3),
  hsla: new RegExp("hsla" + PERMISSIVE4),
  hex3: /^#?([0-9a-fA-F]{1})([0-9a-fA-F]{1})([0-9a-fA-F]{1})$/,
  hex6: /^#?([0-9a-fA-F]{2})([0-9a-fA-F]{2})([0-9a-fA-F]{2})$/,
  hex4: /^#?([0-9a-fA-F]{1})([0-9a-fA-F]{1})([0-9a-fA-F]{1})([0-9a-fA-F]{1})$/,
  hex8: /^#?([0-9a-fA-F]{2})([0-9a-fA-F]{2})([0-9a-fA-F]{2})([0-9a-fA-F]{2})$/,
};

function isValidCSSUnit(c: unknown): boolean {
  return c !== undefined && c !== null && !!MATCHERS.CSS_UNIT.exec(String(c));
}

const hexInt = (v: string) => parseInt(v, 16);

function stringInputToObject(color: string): Record<string, Num> | false {
  color = color.trim().toLowerCase();
  if (color === "transparent") return { r: 0, g: 0, b: 0, a: 0 };
  const named = NAMES[color];
  if (named) color = named;
  let m: RegExpExecArray | null;
  if ((m = MATCHERS.rgb.exec(color))) return { r: m[1], g: m[2], b: m[3] };
  if ((m = MATCHERS.rgba.exec(color))) return { r: m[1], g: m[2], b: m[3], a: m[4] };
  if ((m = MATCHERS.hsl.exec(color))) return { h: m[1], s: m[2], l: m[3] };
  if ((m = MATCHERS.hsla.exec(color))) return { h: m[1], s: m[2], l: m[3], a: m[4] };
  if ((m = MATCHERS.hex8.exec(color))) return { r: hexInt(m[1]), g: hexInt(m[2]), b: hexInt(m[3]), a: hexInt(m[4]) / 255 };
  if ((m = MATCHERS.hex6.exec(color))) return { r: hexInt(m[1]), g: hexInt(m[2]), b: hexInt(m[3]) };
  if ((m = MATCHERS.hex4.exec(color)))
    return { r: hexInt(m[1] + m[1]), g: hexInt(m[2] + m[2]), b: hexInt(m[3] + m[3]), a: hexInt(m[4] + m[4]) / 255 };
  if ((m = MATCHERS.hex3.exec(color))) return { r: hexInt(m[1] + m[1]), g: hexInt(m[2] + m[2]), b: hexInt(m[3] + m[3]) };
  return false;
}

function inputToRGB(input: ColorInput): Rgba & { ok: boolean } {
  let rgb = { r: 0, g: 0, b: 0 };
  let a: Num | undefined = 1;
  let ok = false;
  let color: any = input;
  if (typeof color === "string") color = stringInputToObject(color);
  if (color && typeof color === "object") {
    if (isValidCSSUnit(color.r) && isValidCSSUnit(color.g) && isValidCSSUnit(color.b)) {
      rgb = { r: bound01(color.r, 255) * 255, g: bound01(color.g, 255) * 255, b: bound01(color.b, 255) * 255 };
      ok = true;
    } else if (isValidCSSUnit(color.h) && isValidCSSUnit(color.s) && isValidCSSUnit(color.l)) {
      rgb = hslToRgb(color.h, convertToPercentage(color.s), convertToPercentage(color.l));
      ok = true;
    }
    if (Object.prototype.hasOwnProperty.call(color, "a")) a = color.a;
  }
  return {
    ok,
    r: Math.min(255, Math.max(rgb.r, 0)),
    g: Math.min(255, Math.max(rgb.g, 0)),
    b: Math.min(255, Math.max(rgb.b, 0)),
    a: boundAlpha(a),
  };
}

export class TinyColor {
  private _r: number;
  private _g: number;
  private _b: number;
  private _a: number;
  private _roundA: number;
  private readonly _ok: boolean;

  constructor(input?: ColorInput) {
    if (input instanceof TinyColor) {
      this._r = input._r;
      this._g = input._g;
      this._b = input._b;
      this._a = input._a;
      this._roundA = input._roundA;
      this._ok = input._ok;
      return;
    }
    const rgb = inputToRGB(input ?? "");
    this._r = rgb.r < 1 ? round(rgb.r) : rgb.r;
    this._g = rgb.g < 1 ? round(rgb.g) : rgb.g;
    this._b = rgb.b < 1 ? round(rgb.b) : rgb.b;
    this._a = rgb.a;
    this._roundA = round(100 * this._a) / 100;
    this._ok = rgb.ok;
  }

  isValid(): boolean {
    return this._ok;
  }
  getAlpha(): number {
    return this._a;
  }
  setAlpha(value: Num): this {
    this._a = boundAlpha(value);
    this._roundA = round(100 * this._a) / 100;
    return this;
  }
  toHsl(): Hsla {
    const hsl = rgbToHsl(this._r, this._g, this._b);
    return { h: hsl.h * 360, s: hsl.s, l: hsl.l, a: this._a };
  }
  toHslString(): string {
    const hsl = rgbToHsl(this._r, this._g, this._b);
    const h = round(hsl.h * 360),
      s = round(hsl.s * 100),
      l = round(hsl.l * 100);
    return this._a === 1 ? `hsl(${h}, ${s}%, ${l}%)` : `hsla(${h}, ${s}%, ${l}%, ${this._roundA})`;
  }
  toHex(): string {
    return pad2(round(this._r).toString(16)) + pad2(round(this._g).toString(16)) + pad2(round(this._b).toString(16));
  }
  toHexString(): string {
    return "#" + this.toHex();
  }
  toHex8(): string {
    return this.toHex() + pad2(round(parseFloat(String(this._a)) * 255).toString(16));
  }
  toRgb(): Rgba {
    return { r: round(this._r), g: round(this._g), b: round(this._b), a: this._a };
  }
  toRgbString(): string {
    return this._a === 1
      ? `rgb(${round(this._r)}, ${round(this._g)}, ${round(this._b)})`
      : `rgba(${round(this._r)}, ${round(this._g)}, ${round(this._b)}, ${this._roundA})`;
  }
  saturate(amount = 10): TinyColor {
    const hsl = this.toHsl();
    hsl.s = clamp01(hsl.s + amount / 100);
    return new TinyColor(hsl);
  }
}

/** Factory with the same call style as the original `tinycolor(...)`. */
export function tinycolor(input?: ColorInput): TinyColor {
  return new TinyColor(input);
}

// CSS named colours (lower case) -> hex
const NAMES: Record<string, string> = Object.fromEntries(
  (
    "aliceblue:f0f8ff,antiquewhite:faebd7,aqua:0ff,aquamarine:7fffd4,azure:f0ffff,beige:f5f5dc,bisque:ffe4c4,black:000," +
    "blanchedalmond:ffebcd,blue:00f,blueviolet:8a2be2,brown:a52a2a,burlywood:deb887,burntsienna:ea7e5d,cadetblue:5f9ea0," +
    "chartreuse:7fff00,chocolate:d2691e,coral:ff7f50,cornflowerblue:6495ed,cornsilk:fff8dc,crimson:dc143c,cyan:0ff," +
    "darkblue:00008b,darkcyan:008b8b,darkgoldenrod:b8860b,darkgray:a9a9a9,darkgreen:006400,darkgrey:a9a9a9,darkkhaki:bdb76b," +
    "darkmagenta:8b008b,darkolivegreen:556b2f,darkorange:ff8c00,darkorchid:9932cc,darkred:8b0000,darksalmon:e9967a," +
    "darkseagreen:8fbc8f,darkslateblue:483d8b,darkslategray:2f4f4f,darkslategrey:2f4f4f,darkturquoise:00ced1," +
    "darkviolet:9400d3,deeppink:ff1493,deepskyblue:00bfff,dimgray:696969,dimgrey:696969,dodgerblue:1e90ff," +
    "firebrick:b22222,floralwhite:fffaf0,forestgreen:228b22,fuchsia:f0f,gainsboro:dcdcdc,ghostwhite:f8f8ff,gold:ffd700," +
    "goldenrod:daa520,gray:808080,green:008000,greenyellow:adff2f,grey:808080,honeydew:f0fff0,hotpink:ff69b4," +
    "indianred:cd5c5c,indigo:4b0082,ivory:fffff0,khaki:f0e68c,lavender:e6e6fa,lavenderblush:fff0f5,lawngreen:7cfc00," +
    "lemonchiffon:fffacd,lightblue:add8e6,lightcoral:f08080,lightcyan:e0ffff,lightgoldenrodyellow:fafad2,lightgray:d3d3d3," +
    "lightgreen:90ee90,lightgrey:d3d3d3,lightpink:ffb6c1,lightsalmon:ffa07a,lightseagreen:20b2aa,lightskyblue:87cefa," +
    "lightslategray:789,lightslategrey:789,lightsteelblue:b0c4de,lightyellow:ffffe0,lime:0f0,limegreen:32cd32,linen:faf0e6," +
    "magenta:f0f,maroon:800000,mediumaquamarine:66cdaa,mediumblue:0000cd,mediumorchid:ba55d3,mediumpurple:9370db," +
    "mediumseagreen:3cb371,mediumslateblue:7b68ee,mediumspringgreen:00fa9a,mediumturquoise:48d1cc,mediumvioletred:c71585," +
    "midnightblue:191970,mintcream:f5fffa,mistyrose:ffe4e1,moccasin:ffe4b5,navajowhite:ffdead,navy:000080,oldlace:fdf5e6," +
    "olive:808000,olivedrab:6b8e23,orange:ffa500,orangered:ff4500,orchid:da70d6,palegoldenrod:eee8aa,palegreen:98fb98," +
    "paleturquoise:afeeee,palevioletred:db7093,papayawhip:ffefd5,peachpuff:ffdab9,peru:cd853f,pink:ffc0cb,plum:dda0dd," +
    "powderblue:b0e0e6,purple:800080,rebeccapurple:663399,red:f00,rosybrown:bc8f8f,royalblue:4169e1,saddlebrown:8b4513," +
    "salmon:fa8072,sandybrown:f4a460,seagreen:2e8b57,seashell:fff5ee,sienna:a0522d,silver:c0c0c0,skyblue:87ceeb," +
    "slateblue:6a5acd,slategray:708090,slategrey:708090,snow:fffafa,springgreen:00ff7f,steelblue:4682b4,tan:d2b48c," +
    "teal:008080,thistle:d8bfd8,tomato:ff6347,turquoise:40e0d0,violet:ee82ee,wheat:f5deb3,white:fff,whitesmoke:f5f5f5," +
    "yellow:ff0,yellowgreen:9acd32"
  )
    .split(",")
    .map((p) => p.split(":") as [string, string]),
);
