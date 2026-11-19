/**
 * Auto-numbered bullets (`<a:buAutoNum>`). v1 filled the numbers in with a
 * jQuery pass over the DOM after rendering; this computes them on the HTML
 * string of each text body instead, so it also works without a DOM.
 */

export function romanize(num: number): string {
  if (!(num > 0)) return "";
  const digits = String(Math.floor(num)).split("");
  const key = [
    "", "C", "CC", "CCC", "CD", "D", "DC", "DCC", "DCCC", "CM",
    "", "X", "XX", "XXX", "XL", "L", "LX", "LXX", "LXXX", "XC",
    "", "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX",
  ];
  let roman = "";
  let i = 3;
  while (i--) roman = (key[+(digits.pop() ?? 0) + i * 10] || "") + roman;
  return "M".repeat(+digits.join("") || 0) + roman;
}

/** 1 -> a, 26 -> z, 27 -> aa, 28 -> bb ... (PowerPoint style) */
export function alphaNumeric(num: number, upper: boolean): string {
  if (!(num > 0)) return "";
  const letter = String.fromCharCode(65 + ((num - 1) % 26));
  const s = letter.repeat(Math.floor((num - 1) / 26) + 1);
  return upper ? s : s.toLowerCase();
}

const HEBREW: Array<[number, string]> = [
  [400, "ת"], [300, "ש"], [200, "ר"], [100, "ק"], [90, "צ"], [80, "פ"], [70, "ע"], [60, "ס"], [50, "נ"],
  [40, "מ"], [30, "ל"], [20, "כ"], [10, "י"], [9, "ט"], [8, "ח"], [7, "ז"], [6, "ו"], [5, "ה"], [4, "ד"],
  [3, "ג"], [2, "ב"], [1, "א"],
];

export function hebrewNumeral(n: number): string {
  let ret = "";
  n = n % 1000;
  for (const [v, ch] of HEBREW) for (; n >= v; n -= v) ret += ch;
  return ret.replace(/יה/, "טו").replace(/יו/, "טז");
}

/** Formats `num` according to an OOXML `ST_TextAutonumberScheme`. */
export function getNumTypeNum(scheme: string, num: number): string {
  const m = /^(arabic|alphaLc|alphaUc|romanLc|romanUc|hebrew2|circleNum\w*|ea1\w*|thai\w*|arabicDb\w*)(.*)$/.exec(scheme || "");
  let core: string;
  let style = "Period";
  if (m) {
    style = m[2] || "Plain";
    switch (m[1]) {
      case "alphaLc":
        core = alphaNumeric(num, false);
        break;
      case "alphaUc":
        core = alphaNumeric(num, true);
        break;
      case "romanLc":
        core = romanize(num).toLowerCase();
        break;
      case "romanUc":
        core = romanize(num);
        break;
      case "hebrew2":
        core = hebrewNumeral(num);
        break;
      default:
        core = String(num);
    }
  } else {
    core = String(num);
  }
  switch (style) {
    case "Period":
      return core + ". ";
    case "ParenR":
      return core + ") ";
    case "ParenBoth":
      return "(" + core + ") ";
    case "Minus":
      return core + "-";
    case "Plain":
      return core + " ";
    default:
      return core + ". ";
  }
}

const BULLET_RE = /data-bulltname = '([^']*)' data-bulltlvl = '([^']*)'(?: data-bulltstart = '([^']*)')? class='numeric-bullet-style'><\/div>/g;

/**
 * Fills the numbers into the empty numeric-bullet placeholders of one text
 * body, using v1's counting rules (a counter per level and scheme; it resets
 * when the scheme changes at the same level or a deeper level starts).
 */
export function applyNumericBullets(html: string): string {
  if (html.indexOf("numeric-bullet-style") === -1) return html;
  const counters: Array<{ lvl: number; type: string; n: number }> = [];
  return html.replace(BULLET_RE, (_m, type: string, lvlStr: string, startStr: string | undefined) => {
    const lvl = Number(lvlStr) || 0;
    const start = Number(startStr) || 1;
    // drop counters of deeper levels
    while (counters.length && counters[counters.length - 1].lvl > lvl) counters.pop();
    const top = counters[counters.length - 1];
    let n: number;
    if (top && top.lvl === lvl && top.type === type) {
      n = ++top.n;
    } else {
      if (top && top.lvl === lvl) counters.pop();
      counters.push({ lvl, type, n: start });
      n = start;
    }
    const label = getNumTypeNum(type, n);
    return `data-bulltname = '${type}' data-bulltlvl = '${lvlStr}' class='numeric-bullet-style'>${label}</div>`;
  });
}
