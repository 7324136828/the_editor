/**
 * LaTeX support: live HTML preview (math as MathML), LaTeX <-> word model
 * conversion, templates. Plain ES module, no DOM dependency (Node-testable).
 *
 * Required exports (used by viewer/viewer.js and viewer/latex.js):
 *   newLatexDocument(template?) -> string        template: "article" (default) | "report" |
 *                                                "letter" | "beamer" | "math" (homework) | "cv"
 *   TEMPLATES                   -> [{ id, label, description }]
 *   latexToHtml(source)         -> { html, outline:[{level, title, line, num}],
 *                                    errors:[{line, message, severity:"error"|"warning"}] }
 *                                  (body fragment for the live preview; block elements
 *                                  carry data-line="N" (1-based source line))
 *   latexToHtmlDocument(source, title) -> string (standalone HTML page; math as MathML)
 *   mathToMathML(tex, display:boolean) -> string ("<math>…</math>"; never throws)
 *   latexToWordModel(source)    -> word model { kind:"word", meta:{title?, author?, warnings?}, blocks }
 *                                  (see src/office-backend/src/model.rs WORD CONTRACT)
 *   wordModelToLatex(model)     -> string (standalone, compilable LaTeX document)
 * Extra exports: tokenize, mathToUnicode, PREVIEW_CSS, COMMAND_COMPLETIONS,
 *   ENVIRONMENT_COMPLETIONS, SYMBOL_GROUPS, userCommandNames.
 *
 * Design: tokenize() -> DocParser (a small TeX-like reader with an input stack
 * for macro expansion) builds a block/inline AST -> renderHtml() or
 * toWordModel(). Math is handed to MathParser as a string, which builds a math
 * AST serialised either to MathML Core or to a Unicode approximation.
 * Everything is defensive: caps on nesting, macro expansions and loop counts,
 * and no function here throws on malformed input.
 */

// ===========================================================================
// small utilities
// ===========================================================================

const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const MAX_DEPTH = 120;          // nesting of groups / environments / math
const MAX_EXPANSIONS = 200000;  // macro expansions per document
const MAX_MATH_EXPANSIONS = 5000;

function roman(n, upper) {
  if (!(n > 0) || n >= 4000) return String(n);
  const t = [[1000, "m"], [900, "cm"], [500, "d"], [400, "cd"], [100, "c"], [90, "xc"], [50, "l"], [40, "xl"], [10, "x"], [9, "ix"], [5, "v"], [4, "iv"], [1, "i"]];
  let s = "";
  for (const [v, r] of t) while (n >= v) { s += r; n -= v; }
  return upper ? s.toUpperCase() : s;
}
function alpha(n, upper) {
  if (!(n > 0)) return String(n);
  let s = "";
  while (n > 0) { n--; s = String.fromCharCode(97 + (n % 26)) + s; n = Math.floor(n / 26); }
  return upper ? s.toUpperCase() : s;
}
function fmtCounter(n, style) {
  switch (style) {
    case "alph": return alpha(n, false);
    case "Alph": return alpha(n, true);
    case "roman": return roman(n, false);
    case "Roman": return roman(n, true);
    case "fnsymbol": return ["*", "†", "‡", "§", "¶", "‖", "**", "††", "‡‡"][n - 1] || String(n);
    default: return String(n);
  }
}

function todayString() {
  const d = new Date();
  const m = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  return `${m[d.getMonth()]} ${d.getDate()}, ${d.getFullYear()}`;
}

// ===========================================================================
// colors
// ===========================================================================

const COLORS = {
  red: "#FF0000", green: "#00FF00", blue: "#0000FF", cyan: "#00FFFF", magenta: "#FF00FF", yellow: "#FFFF00",
  black: "#000000", white: "#FFFFFF", gray: "#808080", grey: "#808080", darkgray: "#404040", lightgray: "#BFBFBF",
  brown: "#BF8040", lime: "#BFFF00", olive: "#808000", orange: "#FF8000", pink: "#FFBFBF", purple: "#BF0040",
  teal: "#008080", violet: "#800080", navy: "#000080", maroon: "#800000", darkgreen: "#006400", gold: "#FFD700",
  // dvipsnames (approximate sRGB values used by xcolor)
  apricot: "#FBB982", aquamarine: "#00B5BE", bittersweet: "#C04F17", bluegreen: "#00B3B8", blueviolet: "#473992",
  brickred: "#B6321C", burntorange: "#F7921D", cadetblue: "#74729A", carnationpink: "#F282B4", cerulean: "#00A2E3",
  cornflowerblue: "#41B0E4", dandelion: "#FDBC42", darkorchid: "#A4538A", emerald: "#00A99D", forestgreen: "#009B55",
  fuchsia: "#8C368C", goldenrod: "#FFDF42", greenyellow: "#DFE674", junglegreen: "#00A99A", lavender: "#F49EC4",
  limegreen: "#8DC73E", mahogany: "#A9341F", melon: "#F89E7B", midnightblue: "#006795", mulberry: "#A93C93",
  navyblue: "#006EB8", olivegreen: "#3C8031", orangered: "#ED135A", orchid: "#AF72B0", peach: "#F7965A",
  periwinkle: "#7977B8", pinegreen: "#008B72", plum: "#92268F", processblue: "#00B0F0", rawsienna: "#974006",
  redorange: "#F26035", redviolet: "#A1246B", rhodamine: "#EF559F", royalblue: "#0071BC", royalpurple: "#613F99",
  rubinered: "#ED017D", salmon: "#F69289", seagreen: "#3FBC9D", sepia: "#671800", skyblue: "#46C5DD",
  springgreen: "#C6DC67", tan: "#DA9D76", tealblue: "#00AEB3", thistle: "#D883B7", turquoise: "#00B4CE",
  violetred: "#EF58A0", wildstrawberry: "#EE2967", yellowgreen: "#98CC70", yelloworange: "#FAA21A",
};

const hex2 = (n) => Math.max(0, Math.min(255, Math.round(n))).toString(16).padStart(2, "0").toUpperCase();
function rgbHex(r, g, b) { return `#${hex2(r)}${hex2(g)}${hex2(b)}`; }
function hexRgb(h) { return [parseInt(h.slice(1, 3), 16), parseInt(h.slice(3, 5), 16), parseInt(h.slice(5, 7), 16)]; }

/** Color from a model ("HTML", "rgb", "RGB", "gray", "cmyk") and spec. */
function colorFromModel(model, spec) {
  const m = String(model || "").trim();
  const s = String(spec || "").trim();
  try {
    if (/^html$/i.test(m)) return /^[0-9a-f]{6}$/i.test(s) ? `#${s.toUpperCase()}` : null;
    const nums = s.split(/\s*,\s*|\s+/).map(Number);
    if (m === "rgb" && nums.length === 3 && nums.every((x) => x >= 0 && x <= 1)) return rgbHex(nums[0] * 255, nums[1] * 255, nums[2] * 255);
    if (m === "RGB" && nums.length === 3) return rgbHex(nums[0], nums[1], nums[2]);
    if (/^gray$/i.test(m) && nums.length === 1) return rgbHex(nums[0] * 255, nums[0] * 255, nums[0] * 255);
    if (/^cmyk$/i.test(m) && nums.length === 4) {
      const [c, y0, k] = [nums[0], nums[2], nums[3]];
      const mm = nums[1];
      return rgbHex(255 * (1 - c) * (1 - k), 255 * (1 - mm) * (1 - k), 255 * (1 - y0) * (1 - k));
    }
  } catch { /* fall through */ }
  return null;
}

/** Resolve an xcolor expression ("red", "red!40", "blue!30!black", "#hex") to "#RRGGBB" or null. */
function resolveColor(expr, defined) {
  const s = String(expr || "").trim();
  if (!s) return null;
  if (/^#[0-9a-f]{6}$/i.test(s)) return s.toUpperCase();
  const parts = s.split("!").map((p) => p.trim());
  const base = (name) => {
    if (defined && defined.has(name)) return defined.get(name);
    return COLORS[name.toLowerCase()] || null;
  };
  let cur = base(parts[0]);
  if (!cur) return null;
  for (let k = 1; k < parts.length; k += 2) {
    const pct = Math.max(0, Math.min(100, Number(parts[k])));
    if (!Number.isFinite(pct)) return cur;
    const other = parts[k + 1] !== undefined ? base(parts[k + 1]) : "#FFFFFF";
    if (!other) return cur;
    const a = hexRgb(cur), b = hexRgb(other), f = pct / 100;
    cur = rgbHex(a[0] * f + b[0] * (1 - f), a[1] * f + b[1] * (1 - f), a[2] * f + b[2] * (1 - f));
  }
  return cur;
}

// ===========================================================================
// text-mode symbol tables
// ===========================================================================

const TEXT_SYMBOLS = {
  "%": "%", "&": "&", "#": "#", "_": "_", "$": "$", "{": "{", "}": "}", " ": " ", "@": "", "/": "", "-": "",
  ",": " ", ";": " ", ":": " ", "!": "", ">": " ", "|": "‖",
  textbackslash: "\\", textasciitilde: "~", textasciicircum: "^", textbar: "|", textless: "<", textgreater: ">",
  textbullet: "•", textperiodcentered: "·", textdegree: "°", degree: "°", textregistered: "®", texttrademark: "™",
  copyright: "©", textcopyright: "©", pounds: "£", textsterling: "£", euro: "€", texteuro: "€", EUR: "€",
  textyen: "¥", yen: "¥", textcent: "¢", S: "§", P: "¶", textsection: "§", textparagraph: "¶", dag: "†", ddag: "‡",
  textdagger: "†", textdaggerdbl: "‡", ldots: "…", dots: "…", textellipsis: "…", textendash: "–", textemdash: "—",
  textquoteleft: "‘", textquoteright: "’", textquotedblleft: "“", textquotedblright: "”", textquotesingle: "'",
  textquotedbl: '"', textasciigrave: "`", textasciiacute: "´", textasciidieresis: "¨", textasciimacron: "¯", guillemotleft: "«", guillemotright: "»", guillemetleft: "«", guillemetright: "»",
  guilsinglleft: "‹", guilsinglright: "›", quotedblbase: "„", quotesinglbase: "‚", ss: "ß", SS: "SS", ae: "æ",
  AE: "Æ", oe: "œ", OE: "Œ", o: "ø", O: "Ø", aa: "å", AA: "Å", l: "ł", L: "Ł", i: "ı", j: "ȷ", dh: "ð", DH: "Ð",
  th: "þ", TH: "Þ", ng: "ŋ", NG: "Ŋ", LaTeX: "LaTeX", TeX: "TeX", LaTeXe: "LaTeX2ε", XeLaTeX: "XeLaTeX",
  LuaLaTeX: "LuaLaTeX", BibTeX: "BibTeX", slash: "/", textunderscore: "_", checkmark: "✓", textmu: "µ",
  textonehalf: "½", textonequarter: "¼", textthreequarters: "¾", texttimes: "×", textdiv: "÷", textpm: "±",
  textminus: "−", textordfeminine: "ª", textordmasculine: "º", textexclamdown: "¡", textquestiondown: "¿",
  textbraceleft: "{", textbraceright: "}", textdollar: "$", textnumero: "№", textcelsius: "℃", textohm: "Ω",
  textlangle: "⟨", textrangle: "⟩", textleftarrow: "←", textrightarrow: "→", textuparrow: "↑", textdownarrow: "↓",
  textasteriskcentered: "∗", textvisiblespace: "␣", textcircled: "", quad: " ", qquad: "  ",
  enspace: " ", enskip: " ", thinspace: " ", negthinspace: "", nobreakspace: " ",
  space: " ", hfil: " ", hss: "", dotfill: " … ", textellipsis_: "…", ldotp: ".", cdotp: "·",
  lq: "‘", rq: "’", textsuperscript_: "", Cross: "✗", ding: "", textopenbullet: "◦", ensp: " ",
  bullet: "•",
};

/** Combining marks for text accents (\'e, \"{o}, \c{c} ...). */
const TEXT_ACCENTS = {
  "'": "́", "`": "̀", "^": "̂", '"': "̈", "~": "̃", "=": "̄", ".": "̇",
  u: "̆", v: "̌", H: "̋", c: "̧", d: "̣", b: "̱", r: "̊", k: "̨",
  t: "͡", G: "̏", h: "̉",
};

/** Font sizes in pt for a 10pt class (scaled for 11pt/12pt). */
const SIZES = {
  tiny: 5, scriptsize: 7, footnotesize: 8, small: 9, normalsize: 10, large: 12, Large: 14.4, LARGE: 17.28,
  huge: 20.74, Huge: 24.88,
};

const LOREM = [
  "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Ut purus elit, vestibulum ut, placerat ac, adipiscing vitae, felis. Curabitur dictum gravida mauris. Nam arcu libero, nonummy eget, consectetuer id, vulputate a, magna. Donec vehicula augue eu neque.",
  "Nam dui ligula, fringilla a, euismod sodales, sollicitudin vel, wisi. Morbi auctor lorem non justo. Nam lacus libero, pretium at, lobortis vitae, ultricies et, tellus. Donec aliquet, tortor sed accumsan bibendum, erat ligula aliquet magna, vitae ornare odio metus a mi.",
  "Nulla malesuada porttitor diam. Donec felis erat, congue non, volutpat at, tincidunt tristique, libero. Vivamus viverra fermentum felis. Donec nonummy pellentesque ante. Phasellus adipiscing semper elit. Proin fermentum massa ac quam.",
  "Quisque ullamcorper placerat ipsum. Cras nibh. Morbi vel justo vitae lacus tincidunt ultrices. Lorem ipsum dolor sit amet, consectetuer adipiscing elit. In hac habitasse platea dictumst. Integer tempus convallis augue.",
  "Etiam facilisis. Nunc elementum fermentum wisi. Aenean placerat. Ut imperdiet, enim sed gravida sollicitudin, felis odio placerat quam, ac pulvinar elit purus eget enim. Nunc vitae tortor. Proin tempus nibh sit amet nisl.",
];

// ===========================================================================
// tokenizer
// ===========================================================================

/*
 * Token kinds:
 *   {k:"cmd", v:name, l}      control word or control symbol (v = "\\" for \\)
 *   {k:"txt", v:text, l}      text without specials; whitespace collapsed to " "
 *   {k:"ch",  v:char, l}      one of { } $ $$ & ~ ^ _ # [ ] *
 *   {k:"par", l}              paragraph break (blank line)
 *   {k:"param", v:"1".."9", l}  macro parameter #n
 *   {k:"verb", env, v, opt, l, end}   verbatim-like environment body (raw)
 *   {k:"vinl", v, l}          \verb|...|
 *   {k:"skip", env, l, end}   \begin{comment}...\end{comment}
 */

const VERB_ENVS = new Set(["verbatim", "verbatim*", "Verbatim", "BVerbatim", "lstlisting", "minted", "comment", "filecontents", "filecontents*"]);
const SPECIAL = new Uint8Array(128);
for (const c of "\\{}$&~^_#%[]* \t\n\r") SPECIAL[c.charCodeAt(0)] = 1;

const isLetterCode = (c) => (c >= 65 && c <= 90) || (c >= 97 && c <= 122) || c === 64;

export function tokenize(src, line0 = 1) {
  const s = String(src == null ? "" : src);
  const n = s.length;
  const out = [];
  let i = 0, line = line0, lineStart = true;
  let buf = "", bufLine = line;
  const flush = () => { if (buf) { out.push({ k: "txt", v: buf, l: bufLine }); buf = ""; } };
  const addText = (t) => { if (!buf) bufLine = line; buf += t; };
  while (i < n) {
    const c = s.charCodeAt(i);
    if (c === 32 || c === 9 || c === 10 || c === 13) {
      let nl = 0, j = i;
      const startLine = line;
      while (j < n) {
        const d = s.charCodeAt(j);
        if (d === 10) { nl++; line++; } else if (d !== 32 && d !== 9 && d !== 13) break;
        j++;
      }
      i = j;
      if (nl >= 2 || (nl >= 1 && lineStart)) {
        flush();
        out.push({ k: "par", l: startLine });
        lineStart = true;
      } else if (!lineStart) {
        if (!buf.endsWith(" ")) addText(" ");
      }
      continue;
    }
    if (c === 37) { // % comment: drop to end of line, eat the newline and leading blanks
      const e = s.indexOf("\n", i);
      if (e < 0) { i = n; break; }
      i = e + 1;
      line++;
      while (i < n && (s.charCodeAt(i) === 32 || s.charCodeAt(i) === 9)) i++;
      lineStart = true;
      continue;
    }
    lineStart = false;
    if (c === 92) { // backslash
      flush();
      let j = i + 1;
      if (j >= n) { addText("\\"); i = j; continue; }
      const d = s.charCodeAt(j);
      if (isLetterCode(d)) {
        while (j < n && isLetterCode(s.charCodeAt(j))) j++;
        const name = s.slice(i + 1, j);
        const tl = line;
        if (name === "begin") {
          const m = /^[ \t]*\{([A-Za-z*]+)\}/.exec(s.slice(j, j + 40));
          if (m && VERB_ENVS.has(m[1])) {
            const env = m[1];
            let k = j + m[0].length;
            // header options on the same line: [opts] and/or {lang}
            const hm = /^[ \t]*(\[[^\]\n]*\])?[ \t]*(\{[^}\n]*\})?/.exec(s.slice(k, k + 400));
            let opt = "";
            if (hm && (env === "minted" || env === "lstlisting" || env === "Verbatim" || env === "BVerbatim" || env.startsWith("filecontents"))) {
              opt = ((hm[1] || "") + (hm[2] || "")).trim();
              k += hm[0].length;
            }
            const eol = s.indexOf("\n", k);
            if (eol >= 0 && /^[ \t\r]*$/.test(s.slice(k, eol))) k = eol + 1;
            const endMark = `\\end{${env}}`;
            let e = s.indexOf(endMark, k);
            const closed = e >= 0;
            if (!closed) e = n;
            let body = s.slice(k, e).replace(/\r/g, "");
            for (let p = i; p < e; p++) if (s.charCodeAt(p) === 10) line++;
            body = body.replace(/\n[ \t]*$/, "");
            out.push(env === "comment" ? { k: "skip", env, l: tl, end: line, closed } : { k: "verb", env, v: body, opt, l: tl, end: line, closed });
            i = closed ? e + endMark.length : n;
            continue;
          }
        } else if (name === "verb") {
          let k = j;
          if (s[k] === "*") k++;
          const delim = s[k];
          if (delim && !/[\sA-Za-z]/.test(delim)) {
            let e = s.indexOf(delim, k + 1);
            const nl = s.indexOf("\n", k + 1);
            if (e < 0 || (nl >= 0 && nl < e)) e = nl >= 0 ? nl : n;
            out.push({ k: "vinl", v: s.slice(k + 1, e).replace(/\r/g, ""), l: tl });
            i = Math.min(n, e + 1);
            continue;
          }
        }
        out.push({ k: "cmd", v: name, l: tl });
        i = j;
        // TeX skips blanks after a control word, including one end of line.
        while (i < n && (s.charCodeAt(i) === 32 || s.charCodeAt(i) === 9 || s.charCodeAt(i) === 13)) i++;
        if (i < n && s.charCodeAt(i) === 10) { i++; line++; lineStart = true; while (i < n && (s.charCodeAt(i) === 32 || s.charCodeAt(i) === 9)) i++; }
        continue;
      }
      // control symbol (handle surrogate pairs)
      const cp = s.codePointAt(j);
      const sym = String.fromCodePoint(cp);
      if (sym === "\n" || sym === "\r") { out.push({ k: "cmd", v: " ", l: line }); i = j; continue; }
      out.push({ k: "cmd", v: sym, l: line });
      i = j + sym.length;
      continue;
    }
    if (c < 128 && SPECIAL[c]) {
      flush();
      const ch = s[i];
      if (ch === "$" && s[i + 1] === "$") { out.push({ k: "ch", v: "$$", l: line }); i += 2; continue; }
      if (ch === "#") {
        const nx = s[i + 1];
        if (nx >= "1" && nx <= "9") { out.push({ k: "param", v: nx, l: line }); i += 2; continue; }
        if (nx === "#") { out.push({ k: "param", v: "#", l: line }); i += 2; continue; }
      }
      out.push({ k: "ch", v: ch, l: line });
      i++;
      continue;
    }
    // plain text run
    let j = i + 1;
    while (j < n) {
      const d = s.charCodeAt(j);
      if (d < 128 && SPECIAL[d]) break;
      j++;
    }
    addText(s.slice(i, j));
    i = j;
  }
  flush();
  return out;
}

/** Serialize tokens back to LaTeX source (used for math, URLs, macro bodies). */
function serialize(toks) {
  let s = "";
  for (let k = 0; k < toks.length; k++) {
    const t = toks[k];
    switch (t.k) {
      case "cmd":
        s += "\\" + t.v;
        if (/^[A-Za-z@]+$/.test(t.v)) {
          const nx = toks[k + 1];
          if (nx && ((nx.k === "txt" && /^[A-Za-z@]/.test(nx.v)) || nx.k === "param")) s += " ";
        }
        break;
      case "txt": s += t.v; break;
      case "ch": s += t.v; break;
      case "par": s += "\n\n"; break;
      case "param": s += "#" + t.v; break;
      case "verb": s += `\\begin{${t.env}}${t.opt || ""}\n${t.v}\n\\end{${t.env}}`; break;
      case "vinl": s += `\\verb|${t.v}|`; break;
      default: break;
    }
  }
  return s;
}

/** Plain text of tokens (for labels, keys, URLs, color names...). */
function plainText(toks) {
  let s = "";
  for (const t of toks) {
    if (t.k === "txt") s += t.v;
    else if (t.k === "ch") s += t.v === "~" ? "~" : t.v;
    else if (t.k === "cmd") s += t.v.length === 1 && !/[A-Za-z]/.test(t.v) ? (t.v === "\\" ? "" : t.v) : (TEXT_SYMBOLS[t.v] ?? "");
    else if (t.k === "param") s += "#" + t.v;
    else if (t.k === "vinl") s += t.v;
    else if (t.k === "par") s += " ";
  }
  return s;
}

// ===========================================================================
// math: symbol tables
// ===========================================================================

const GREEK_LOWER = {
  alpha: "α", beta: "β", gamma: "γ", delta: "δ", epsilon: "ϵ", varepsilon: "ε", zeta: "ζ", eta: "η", theta: "θ",
  vartheta: "ϑ", iota: "ι", kappa: "κ", varkappa: "ϰ", lambda: "λ", mu: "μ", nu: "ν", xi: "ξ", omicron: "ο", pi: "π",
  varpi: "ϖ", rho: "ρ", varrho: "ϱ", sigma: "σ", varsigma: "ς", tau: "τ", upsilon: "υ", phi: "ϕ", varphi: "φ",
  chi: "χ", psi: "ψ", omega: "ω", digamma: "ϝ",
};
const GREEK_UPPER = {
  Alpha: "Α", Beta: "Β", Gamma: "Γ", Delta: "Δ", Epsilon: "Ε", Zeta: "Ζ", Eta: "Η", Theta: "Θ", Iota: "Ι", Kappa: "Κ",
  Lambda: "Λ", Mu: "Μ", Nu: "Ν", Xi: "Ξ", Omicron: "Ο", Pi: "Π", Rho: "Ρ", Sigma: "Σ", Tau: "Τ", Upsilon: "Υ",
  Phi: "Φ", Chi: "Χ", Psi: "Ψ", Omega: "Ω", varGamma: "𝛤", varDelta: "𝛥", varTheta: "𝛩", varLambda: "𝛬",
  varXi: "𝛯", varPi: "𝛱", varSigma: "𝛴", varUpsilon: "𝛶", varPhi: "𝛷", varPsi: "𝛹", varOmega: "𝛺",
};
/** Ordinary symbols rendered as upright identifiers. */
const MATH_ORD = {
  infty: "∞", partial: "∂", nabla: "∇", emptyset: "∅", varnothing: "∅", aleph: "ℵ", beth: "ℶ", gimel: "ℷ",
  hbar: "ℏ", hslash: "ℏ", ell: "ℓ", wp: "℘", Re: "ℜ", Im: "ℑ", imath: "ı", jmath: "ȷ", angle: "∠",
  measuredangle: "∡", sphericalangle: "∢", triangle: "△", top: "⊤", bot: "⊥", clubsuit: "♣", diamondsuit: "♢",
  heartsuit: "♡", spadesuit: "♠", flat: "♭", natural: "♮", sharp: "♯", checkmark: "✓", square: "□", Box: "□",
  blacksquare: "■", lozenge: "◊", blacklozenge: "⧫", mho: "℧", eth: "ð", Finv: "Ⅎ", complement: "∁",
  backprime: "‵", degree: "°", S: "§", P: "¶", dag: "†", ddag: "‡", copyright: "©", pounds: "£", varkappa_: "ϰ",
  bigstar: "★", blacktriangle: "▴", blacktriangledown: "▾", vartriangle: "△", triangledown: "▿", diagup: "╱",
  diagdown: "╲", circledcirc: "⊚", maltese: "✠", yen: "¥", euro: "€", "_": "_", "#": "#", "%": "%", "&": "&",
  "$": "$", AA: "Å",
};
const MATH_OPS = {
  pm: "±", mp: "∓", times: "×", div: "÷", cdot: "⋅", ast: "∗", star: "⋆", circ: "∘", bullet: "∙", oplus: "⊕",
  ominus: "⊖", otimes: "⊗", oslash: "⊘", odot: "⊙", cap: "∩", cup: "∪", sqcap: "⊓", sqcup: "⊔", vee: "∨",
  lor: "∨", wedge: "∧", land: "∧", setminus: "∖", smallsetminus: "∖", wr: "≀", diamond: "⋄", bigtriangleup: "△",
  bigtriangledown: "▽", triangleleft: "◃", triangleright: "▹", uplus: "⊎", amalg: "⨿", dagger: "†",
  ddagger: "‡", ltimes: "⋉", rtimes: "⋊", bowtie: "⋈", boxplus: "⊞", boxminus: "⊟", boxtimes: "⊠", boxdot: "⊡",
  cdotp: "⋅", centerdot: "⋅", divideontimes: "⋇", dotplus: "∔", intercal: "⊺", barwedge: "⌅", veebar: "⊻",
  curlyvee: "⋎", curlywedge: "⋏", circledast: "⊛", circleddash: "⊝", Cap: "⋒", Cup: "⋓",
  leq: "≤", le: "≤", geq: "≥", ge: "≥", neq: "≠", ne: "≠", leqslant: "⩽", geqslant: "⩾", ll: "≪", gg: "≫",
  lll: "⋘", ggg: "⋙", prec: "≺", succ: "≻", preceq: "⪯", succeq: "⪰", sim: "∼", simeq: "≃", approx: "≈",
  cong: "≅", equiv: "≡", doteq: "≐", propto: "∝", asymp: "≍", subset: "⊂", supset: "⊃", subseteq: "⊆",
  supseteq: "⊇", subsetneq: "⊊", supsetneq: "⊋", sqsubset: "⊏", sqsupset: "⊐", sqsubseteq: "⊑", sqsupseteq: "⊒",
  in: "∈", ni: "∋", owns: "∋", notin: "∉", vdash: "⊢", dashv: "⊣", models: "⊨", vDash: "⊨", Vdash: "⊩",
  perp: "⊥", mid: "∣", parallel: "∥", nmid: "∤", nparallel: "∦", smile: "⌣", frown: "⌢", lesssim: "≲",
  gtrsim: "≳", approxeq: "≊", triangleq: "≜", coloneqq: "≔", coloneq: "≔", eqqcolon: "≕", nleq: "≰", ngeq: "≱",
  nless: "≮", ngtr: "≯", nsim: "≁", ncong: "≇", nsubseteq: "⊈", nsupseteq: "⊉", leqq: "≦", geqq: "≧", lneq: "⪇",
  gneq: "⪈", therefore: "∴", because: "∵", bumpeq: "≏", Bumpeq: "≎", circeq: "≗", eqcirc: "≖", lessgtr: "≶",
  gtrless: "≷", trianglelefteq: "⊴", trianglerighteq: "⊵", vartriangleleft: "⊲", vartriangleright: "⊳",
  lhd: "⊲", rhd: "⊳", unlhd: "⊴", unrhd: "⊵", Subset: "⋐", Supset: "⋑", varpropto: "∝", backsim: "∽",
  thicksim: "∼", thickapprox: "≈", pitchfork: "⋔", between: "≬", nvdash: "⊬", nvDash: "⊭", prec_: "≺",
  to: "→", rightarrow: "→", leftarrow: "←", gets: "←", leftrightarrow: "↔", Rightarrow: "⇒", Leftarrow: "⇐",
  Leftrightarrow: "⇔", implies: "⟹", impliedby: "⟸", iff: "⟺", mapsto: "↦", longmapsto: "⟼",
  longrightarrow: "⟶", longleftarrow: "⟵", longleftrightarrow: "⟷", Longrightarrow: "⟹", Longleftarrow: "⟸",
  Longleftrightarrow: "⟺", uparrow: "↑", downarrow: "↓", updownarrow: "↕", Uparrow: "⇑", Downarrow: "⇓",
  Updownarrow: "⇕", nearrow: "↗", searrow: "↘", swarrow: "↙", nwarrow: "↖", hookrightarrow: "↪",
  hookleftarrow: "↩", rightharpoonup: "⇀", rightharpoondown: "⇁", leftharpoonup: "↼", leftharpoondown: "↽",
  rightleftharpoons: "⇌", leftrightharpoons: "⇋", leadsto: "⇝", rightsquigarrow: "⇝", circlearrowleft: "↺",
  circlearrowright: "↻", curvearrowleft: "↶", curvearrowright: "↷", twoheadrightarrow: "↠",
  twoheadleftarrow: "↞", nrightarrow: "↛", nleftarrow: "↚", nRightarrow: "⇏", nLeftarrow: "⇍",
  nLeftrightarrow: "⇎", nleftrightarrow: "↮", rightleftarrows: "⇄", leftrightarrows: "⇆", upuparrows: "⇈",
  downdownarrows: "⇊", Lsh: "↰", Rsh: "↱", dashrightarrow: "⇢", dashleftarrow: "⇠", rightarrowtail: "↣",
  leftarrowtail: "↢", restriction: "↾", multimap: "⊸", longleftarrow_: "⟵",
  forall: "∀", exists: "∃", nexists: "∄", neg: "¬", lnot: "¬", colon: ":", ldots: "…", cdots: "⋯",
  vdots: "⋮", ddots: "⋱", dots: "…", dotsc: "…", dotsb: "⋯", dotsm: "⋯", dotsi: "⋯", dotso: "…", iddots: "⋰",
  adots: "⋰", hdots: "…", prime: "′", surd: "√", backslash: "\\", lbrack: "[", rbrack: "]", lparen: "(", rparen: ")",
};
/** Fence-like operators that must not stretch unless used with \left/\right. */
const FENCE_CHARS = new Set(["(", ")", "[", "]", "{", "}", "|", "‖", "⟨", "⟩", "⌊", "⌋", "⌈", "⌉", "⟮", "⟯", "/", "\\", "⌜", "⌝", "⌞", "⌟"]);
const DELIMS = {
  lbrace: "{", rbrace: "}", "{": "{", "}": "}", langle: "⟨", rangle: "⟩", lfloor: "⌊", rfloor: "⌋", lceil: "⌈",
  rceil: "⌉", vert: "|", Vert: "‖", lvert: "|", rvert: "|", lVert: "‖", rVert: "‖", "|": "‖", lgroup: "⟮",
  rgroup: "⟯", backslash: "\\", uparrow: "↑", downarrow: "↓", updownarrow: "↕", Uparrow: "⇑", Downarrow: "⇓",
  Updownarrow: "⇕", lbrack: "[", rbrack: "]", ulcorner: "⌜", urcorner: "⌝", llcorner: "⌞", lrcorner: "⌟",
  lmoustache: "⎰", rmoustache: "⎱", "<": "⟨", ">": "⟩",
};
const LARGE_OPS = {
  sum: ["∑", true], prod: ["∏", true], coprod: ["∐", true], bigcup: ["⋃", true], bigcap: ["⋂", true],
  bigoplus: ["⨁", true], bigotimes: ["⨂", true], bigodot: ["⨀", true], bigvee: ["⋁", true], bigwedge: ["⋀", true],
  bigsqcup: ["⨆", true], biguplus: ["⨄", true], int: ["∫", false], iint: ["∬", false], iiint: ["∭", false],
  iiiint: ["⨌", false], oint: ["∮", false], oiint: ["∯", false], intop: ["∫", true], smallint: ["∫", false],
  idotsint: ["∫⋯∫", false],
};
const FUNCS = {
  sin: 0, cos: 0, tan: 0, cot: 0, sec: 0, csc: 0, arcsin: 0, arccos: 0, arctan: 0, sinh: 0, cosh: 0, tanh: 0, coth: 0,
  log: 0, ln: 0, lg: 0, exp: 0, dim: 0, ker: 0, hom: 0, arg: 0, deg: 0, lcm: 0, sgn: 0, tr: 0, Tr: 0, rank: 0,
  det: 1, gcd: 1, Pr: 1, min: 1, max: 1, sup: 1, inf: 1, lim: 1, liminf: 1, limsup: 1, injlim: 1, projlim: 1, argmin: 1, argmax: 1,
};
const FUNC_TEXT = { liminf: "lim inf", limsup: "lim sup", injlim: "inj lim", projlim: "proj lim", argmin: "arg min", argmax: "arg max" };
const MATH_ACCENTS = {
  hat: ["^", false], widehat: ["^", true], check: ["ˇ", false], widecheck: ["ˇ", true], tilde: ["~", false],
  widetilde: ["~", true], acute: ["´", false], grave: ["`", false], dot: ["˙", false], ddot: ["¨", false],
  dddot: ["⃛", false], breve: ["˘", false], bar: ["¯", false], overline: ["‾", true], vec: ["→", false],
  overrightarrow: ["→", true], overleftarrow: ["←", true], overleftrightarrow: ["↔", true], mathring: ["˚", false],
  overbrace: ["⏞", true], Overrightarrow: ["⇒", true],
  underline: ["_", true, true], underbrace: ["⏟", true, true], underrightarrow: ["→", true, true],
  underleftarrow: ["←", true, true], underleftrightarrow: ["↔", true, true],
};
/** Unicode combining marks used by the plain-text approximation. */
const COMBINING = {
  hat: "̂", widehat: "̂", check: "̌", widecheck: "̌", tilde: "̃", widetilde: "̃",
  acute: "́", grave: "̀", dot: "̇", ddot: "̈", dddot: "⃛", breve: "̆", bar: "̄",
  overline: "̅", vec: "⃗", overrightarrow: "⃗", overleftarrow: "⃖", mathring: "̊",
  underline: "̲",
};
const NOT_MAP = {
  "=": "≠", "∈": "∉", "⊂": "⊄", "⊃": "⊅", "⊆": "⊈", "⊇": "⊉", "<": "≮", ">": "≯", "≤": "≰", "≥": "≱", "∼": "≁",
  "≅": "≇", "≡": "≢", "≈": "≉", "∣": "∤", "∥": "∦", "∃": "∄", "≃": "≄", "⊢": "⊬", "⊨": "⊭", "→": "↛", "←": "↚",
  "↔": "↮", "⇒": "⇏", "⇐": "⇍", "⇔": "⇎", "≺": "⊀", "≻": "⊁", "∋": "∌", "⪯": "⋠", "⪰": "⋡",
};
const SPACES = {
  ",": 0.1667, ":": 0.2222, ">": 0.2222, ";": 0.2778, "!": -0.1667, " ": 0.25, quad: 1, qquad: 2, enspace: 0.5,
  thinspace: 0.1667, medspace: 0.2222, thickspace: 0.2778, negthinspace: -0.1667, negmedspace: -0.2222,
  negthickspace: -0.2778, enskip: 0.5, space: 0.25, nobreakspace: 0.25,
};
const BIG_SIZES = { big: "1.2em", Big: "1.623em", bigg: "2.047em", Bigg: "2.470em" };
const MATH_FONTS = {
  mathbf: "bf", mathit: "it", mathrm: "rm", mathsf: "sf", mathtt: "tt", mathbb: "bb", mathcal: "cal", mathscr: "cal",
  mathfrak: "frak", boldsymbol: "bi", bm: "bi", pmb: "bi", mathbfit: "bi", Bbb: "bb", mathnormal: "", mathup: "rm",
  mathsfbf: "sfbf", mathbfsf: "sfbf", symbf: "bf", symit: "it", symbb: "bb", symcal: "cal", symrm: "rm",
};
const OLD_FONTS = { rm: "rm", bf: "bf", it: "it", sf: "sf", tt: "tt", cal: "cal", mit: "", frak: "frak", Bbb: "bb" };
const MATH_TEXT_CMDS = {
  text: {}, textrm: {}, textnormal: {}, mbox: {}, hbox: {}, textup: {}, textmd: {}, mathrm_: {},
  textbf: { b: 1 }, textit: { i: 1 }, textsl: { i: 1 }, emph: { i: 1 }, texttt: { tt: 1 }, textsf: { sf: 1 },
  fbox: { box: 1 }, framebox: { box: 1 },
};

// Mathematical alphanumeric symbols: [capital A, small a, digit 0]
const ALNUM = {
  bf: [0x1D400, 0x1D41A, 0x1D7CE], it: [0x1D434, 0x1D44E, 0], bi: [0x1D468, 0x1D482, 0x1D7CE],
  cal: [0x1D49C, 0x1D4B6, 0], frak: [0x1D504, 0x1D51E, 0], bb: [0x1D538, 0x1D552, 0x1D7D8],
  sf: [0x1D5A0, 0x1D5BA, 0x1D7E2], sfbf: [0x1D5D4, 0x1D5EE, 0x1D7EC], tt: [0x1D670, 0x1D68A, 0x1D7F6],
};
const ALNUM_HOLES = {
  it: { h: "ℎ" },
  cal: { B: "ℬ", E: "ℰ", F: "ℱ", H: "ℋ", I: "ℐ", L: "ℒ", M: "ℳ", R: "ℛ", e: "ℯ", g: "ℊ", o: "ℴ" },
  frak: { C: "ℭ", H: "ℌ", I: "ℑ", R: "ℜ", Z: "ℨ" },
  bb: { C: "ℂ", H: "ℍ", N: "ℕ", P: "ℙ", Q: "ℚ", R: "ℝ", Z: "ℤ" },
};

/** Map one character to the given math font (returns it unchanged when there is no variant). */
function mathFontChar(ch, font) {
  if (!font || font === "rm" || font === "it") return ch;
  const t = ALNUM[font];
  if (!t) return ch;
  const holes = ALNUM_HOLES[font];
  if (holes && holes[ch]) return holes[ch];
  const c = ch.charCodeAt(0);
  if (c >= 65 && c <= 90) return String.fromCodePoint(t[0] + c - 65);
  if (c >= 97 && c <= 122) return String.fromCodePoint(t[1] + c - 97);
  if (c >= 48 && c <= 57 && t[2]) return String.fromCodePoint(t[2] + c - 48);
  if (font === "bf" || font === "bi") {
    const cp = ch.codePointAt(0);
    if (cp >= 0x391 && cp <= 0x3A9) return String.fromCodePoint((font === "bf" ? 0x1D6A8 : 0x1D71C) + cp - 0x391);
    if (cp >= 0x3B1 && cp <= 0x3C9 && font === "bi") return String.fromCodePoint(0x1D736 + cp - 0x3B1);
    if (cp >= 0x3B1 && cp <= 0x3C9 && font === "bf") return String.fromCodePoint(0x1D6C2 + cp - 0x3B1);
  }
  return ch;
}

const SUP_MAP = {
  0: "⁰", 1: "¹", 2: "²", 3: "³", 4: "⁴", 5: "⁵", 6: "⁶", 7: "⁷", 8: "⁸", 9: "⁹", "+": "⁺", "−": "⁻", "-": "⁻", "=": "⁼",
  "(": "⁽", ")": "⁾", n: "ⁿ", i: "ⁱ", a: "ᵃ", b: "ᵇ", c: "ᶜ", d: "ᵈ", e: "ᵉ", f: "ᶠ", g: "ᵍ", h: "ʰ", j: "ʲ", k: "ᵏ",
  l: "ˡ", m: "ᵐ", o: "ᵒ", p: "ᵖ", r: "ʳ", s: "ˢ", t: "ᵗ", u: "ᵘ", v: "ᵛ", w: "ʷ", x: "ˣ", y: "ʸ", z: "ᶻ", T: "ᵀ",
  A: "ᴬ", B: "ᴮ", D: "ᴰ", E: "ᴱ", G: "ᴳ", H: "ᴴ", I: "ᴵ", J: "ᴶ", K: "ᴷ", L: "ᴸ", M: "ᴹ", N: "ᴺ", O: "ᴼ", P: "ᴾ",
  R: "ᴿ", U: "ᵁ", V: "ⱽ", W: "ᵂ", "′": "′", "″": "″", "‴": "‴", "∗": "*", "*": "*", " ": "",
};
const SUB_MAP = {
  0: "₀", 1: "₁", 2: "₂", 3: "₃", 4: "₄", 5: "₅", 6: "₆", 7: "₇", 8: "₈", 9: "₉", "+": "₊", "−": "₋", "-": "₋", "=": "₌",
  "(": "₍", ")": "₎", a: "ₐ", e: "ₑ", h: "ₕ", i: "ᵢ", j: "ⱼ", k: "ₖ", l: "ₗ", m: "ₘ", n: "ₙ", o: "ₒ", p: "ₚ",
  r: "ᵣ", s: "ₛ", t: "ₜ", u: "ᵤ", v: "ᵥ", x: "ₓ", "β": "ᵦ", "γ": "ᵧ", "ρ": "ᵨ", "φ": "ᵩ", "χ": "ᵪ", " ": "",
};

/** Text-mode processing of a raw string (for \text{} in math and quick plain conversions). */
function textModeString(raw) {
  let s = String(raw);
  s = s.replace(/\\([A-Za-z]+)\s*(\{\})?/g, (m, name) => (name in TEXT_SYMBOLS ? TEXT_SYMBOLS[name] : name === "today" ? todayString() : ""));
  s = s.replace(/\\([%&#_${}\s,;:!])/g, (m, c) => (c === "," ? " " : /\s/.test(c) ? " " : c === ";" || c === ":" || c === "!" ? " " : c));
  s = s.replace(/~/g, " ").replace(/---/g, "—").replace(/--/g, "–").replace(/``/g, "“").replace(/''/g, "”");
  s = s.replace(/[{}]/g, "");
  return s;
}

// ===========================================================================
// math: parser (string -> AST)
// ===========================================================================

const isAsciiLetter = (c) => /^[A-Za-z]$/.test(c);

class MathParser {
  constructor(src, opts = {}) {
    this.s = String(src);
    this.i = 0;
    this.n = this.s.length;
    this.errors = opts.errors || [];
    this.macros = opts.macros || null;
    this.display = !!opts.display;
    this.font = "";
    this.depth = opts.depth || 0;
    this.exp = 0;
    this.steps = 0;
  }

  err(msg) { if (this.errors.length < 40 && !this.errors.includes(msg)) this.errors.push(msg); }

  ws() {
    const s = this.s;
    while (this.i < this.n) {
      const c = s.charCodeAt(this.i);
      if (c === 32 || c === 9 || c === 10 || c === 13) this.i++;
      else if (c === 37) { const e = s.indexOf("\n", this.i); this.i = e < 0 ? this.n : e + 1; } else break;
    }
  }

  lex(adv) {
    this.ws();
    if (this.i >= this.n) return null;
    const s = this.s;
    let j = this.i;
    let tok;
    if (s[j] === "\\") {
      j++;
      if (j >= this.n) tok = { k: "ch", v: "\\" };
      else if (isAsciiLetter(s[j])) {
        let e = j;
        while (e < this.n && isAsciiLetter(s[e])) e++;
        tok = { k: "cmd", v: s.slice(j, e) };
        j = e;
      } else {
        const ch = String.fromCodePoint(s.codePointAt(j));
        tok = { k: "cmd", v: ch };
        j += ch.length;
      }
    } else {
      const ch = String.fromCodePoint(s.codePointAt(j));
      tok = { k: "ch", v: ch };
      j += ch.length;
    }
    if (adv) this.i = j;
    return tok;
  }
  peek() { return this.lex(false); }
  next() { return this.lex(true); }

  /** Raw balanced {...} content (or one token) as a string. */
  readRawArg() {
    this.ws();
    if (this.i >= this.n) return "";
    if (this.s[this.i] === "{") {
      let depth = 0;
      const start = this.i + 1;
      for (let j = this.i; j < this.n; j++) {
        const c = this.s[j];
        if (c === "\\") { j++; continue; }
        if (c === "{") depth++;
        else if (c === "}") { depth--; if (depth === 0) { this.i = j + 1; return this.s.slice(start, j); } }
      }
      this.err("Missing } in math");
      const r = this.s.slice(start);
      this.i = this.n;
      return r;
    }
    const t = this.next();
    if (!t) return "";
    return t.k === "cmd" ? "\\" + t.v : t.v;
  }

  readRawOpt() {
    this.ws();
    if (this.s[this.i] !== "[") return null;
    let depth = 0;
    for (let j = this.i + 1; j < this.n; j++) {
      const c = this.s[j];
      if (c === "\\") { j++; continue; }
      if (c === "{") depth++;
      else if (c === "}") depth--;
      else if (c === "]" && depth === 0) { const r = this.s.slice(this.i + 1, j); this.i = j + 1; return r; }
    }
    return null;
  }

  sub(str, display = this.display) {
    if (this.depth > MAX_DEPTH) { this.err("Formula nested too deeply"); return { t: "err", v: "…" }; }
    const p = new MathParser(str, { errors: this.errors, macros: this.macros, display, depth: this.depth + 1 });
    p.font = this.font;
    p.colors = this.colors;
    p.exp = this.exp;
    const nodes = p.parseList(new Set());
    this.exp = p.exp;
    return { t: "row", c: nodes };
  }

  parseTop() { return this.parseList(new Set()); }

  /** Parse atoms until one of `stops` ("}", "&", "\\\\", "end", "right", "]") is next. */
  parseList(stops) {
    if (++this.depth > MAX_DEPTH) { this.depth--; this.err("Formula nested too deeply"); this.i = this.n; return [{ t: "err", v: "…" }]; }
    const out = [];
    let infix = null;
    try {
      for (;;) {
        if (++this.steps > 400000) { this.i = this.n; break; }
        const t = this.peek();
        if (!t) break;
        if (t.k === "ch") {
          if (t.v === "}") { if (stops.has("}")) break; this.next(); this.err("Extra } in a formula"); out.push({ t: "err", v: "}" }); continue; }
          if (t.v === "&") { if (stops.has("&")) break; this.next(); this.err("Misplaced & (only allowed inside align or matrix environments)"); continue; }
          if (t.v === "]" && stops.has("]")) break;
        } else {
          const v = t.v;
          if (v === "\\" || v === "cr" || v === "newline") {
            if (stops.has("\\\\")) break;
            this.next(); this.skipRowOpts(); continue;
          }
          if (v === "end") {
            if (stops.has("end")) break;
            this.next(); const nm = this.readRawArg();
            this.err(`\\end{${nm}} without a matching \\begin{${nm}}`);
            continue;
          }
          if (v === "right") {
            if (stops.has("right")) break;
            this.next(); this.readDelim(); this.err("\\right without a matching \\left"); continue;
          }
          if (v === "over" || v === "choose" || v === "atop" || v === "above" || v === "brack" || v === "brace") {
            this.next();
            if (v === "above") this.readDimen();
            infix = { num: out.splice(0), kind: v };
            continue;
          }
          if (v === "displaystyle" || v === "textstyle" || v === "scriptstyle" || v === "scriptscriptstyle") {
            this.next();
            const saved = this.display;
            this.display = v === "displaystyle";
            const rest = this.parseList(stops);
            this.display = saved;
            out.push({ t: "style", disp: v, c: rest });
            break;
          }
          if (v === "color") {
            this.next();
            const model = this.readRawOpt();
            const spec = this.readRawArg();
            const rest = this.parseList(stops);
            out.push({ t: "style", color: model ? colorFromModel(model, spec) : resolveColor(spec, this.colors), c: rest });
            break;
          }
          if (Object.prototype.hasOwnProperty.call(OLD_FONTS, v) && !(this.macros && this.macros.has(v))) {
            this.next();
            const saved = this.font;
            this.font = OLD_FONTS[v];
            const rest = this.parseList(stops);
            this.font = saved;
            out.push({ t: "row", c: rest });
            break;
          }
          if (v === "hline" || v === "hdashline" || v === "midrule" || v === "toprule" || v === "bottomrule") { this.next(); out.push({ t: "hline" }); continue; }
        }
        const atom = this.parseScripted();
        if (atom) out.push(atom);
      }
    } finally {
      this.depth--;
    }
    if (infix) {
      const den = out.splice(0);
      const n = { t: "row", c: infix.num }, d = { t: "row", c: den };
      if (infix.kind === "over" || infix.kind === "above") out.push({ t: "frac", n, d });
      else if (infix.kind === "atop") out.push({ t: "frac", n, d, thick: "0" });
      else out.push({ t: "frac", n, d, thick: "0", fence: infix.kind === "choose" ? ["(", ")"] : infix.kind === "brack" ? ["[", "]"] : ["{", "}"] });
    }
    return out;
  }

  skipRowOpts() {
    this.ws();
    if (this.s[this.i] === "*") this.i++;
    this.readRawOpt();
  }

  readDimen() {
    this.ws();
    const m = /^-?\s*[\d.]+\s*(em|ex|pt|mm|cm|in|mu|bp|pc|sp)?/.exec(this.s.slice(this.i, this.i + 30));
    if (m) { this.i += m[0].length; return m[0]; }
    return "";
  }

  /** Atom plus its scripts / primes / \limits. */
  parseScripted() {
    let base = null;
    const t = this.peek();
    if (t && t.k === "ch" && (t.v === "^" || t.v === "_" || t.v === "'")) base = { t: "row", c: [] };
    else {
      base = this.parseAtom();
      if (!base) return null;
    }
    for (;;) {
      const p = this.peek();
      if (p && p.k === "cmd" && (p.v === "limits" || p.v === "nolimits" || p.v === "displaylimits")) {
        this.next();
        base.limits = p.v === "limits" ? true : p.v === "nolimits" ? false : undefined;
      } else break;
    }
    let sub = null, sup = null, primes = 0;
    for (let guard = 0; guard < 20; guard++) {
      const p = this.peek();
      if (!p || p.k !== "ch") break;
      if (p.v === "'") { this.next(); primes++; continue; }
      if (p.v === "^" || p.v === "_") {
        this.next();
        const savedD = this.display;
        this.display = false;
        const arg = this.parseArg();
        this.display = savedD;
        if (p.v === "^") { if (sup) this.err("Double superscript (use braces: x^{ab})"); sup = arg; } else { if (sub) this.err("Double subscript (use braces: x_{ab})"); sub = arg; }
        continue;
      }
      break;
    }
    if (primes) {
      const pr = { t: "o", v: primes === 1 ? "′" : primes === 2 ? "″" : primes === 3 ? "‴" : "′".repeat(primes), prime: true };
      sup = sup ? { t: "row", c: [pr, sup] } : pr;
    }
    if (!sub && !sup) return base;
    return { t: "scripts", base, sub, sup, disp: this.display };
  }

  /** One argument: {group} or a single token. */
  parseArg() {
    const t = this.peek();
    if (!t) { this.err("Missing argument in a formula"); return { t: "row", c: [] }; }
    if (t.k === "ch" && t.v === "{") return this.parseAtom();
    if (t.k === "ch") {
      this.next();
      if (/[0-9]/.test(t.v)) return { t: "n", v: mathFontChar(t.v, this.font) };
      return this.charAtom(t.v, true);
    }
    return this.parseAtom();
  }

  parseAtom() {
    const t = this.next();
    if (!t) return null;
    if (t.k === "ch") return this.charAtom(t.v, false);
    return this.command(t.v);
  }

  charAtom(c, single) {
    if (c === "{") {
      const nodes = this.parseList(new Set(["}"]));
      const close = this.next();
      if (!close || close.v !== "}") this.err("Missing } in a formula");
      return { t: "row", c: nodes, group: true };
    }
    if (isAsciiLetter(c)) {
      if (this.font === "rm" && !single) {
        let w = c;
        while (this.i < this.n && isAsciiLetter(this.s[this.i])) w += this.s[this.i++];
        return { t: "i", v: w, normal: true };
      }
      return { t: "i", v: mathFontChar(c, this.font), normal: this.font === "rm" };
    }
    if (/[0-9]/.test(c) || (c === "." && /[0-9]/.test(this.s[this.i] || ""))) {
      let num = c;
      if (!single) {
        while (this.i < this.n) {
          const d = this.s[this.i];
          if (/[0-9]/.test(d) || (d === "." && /[0-9]/.test(this.s[this.i + 1] || ""))) { num += d; this.i++; } else break;
        }
      }
      return { t: "n", v: [...num].map((ch) => mathFontChar(ch, this.font)).join("") };
    }
    switch (c) {
      case "-": return { t: "o", v: "−" };
      case "*": return { t: "o", v: "∗" };
      case "~": return { t: "space", w: 0.25 };
      case "$": return null;
      case "^": case "_": this.err(`Misplaced ${c} in a formula`); return null;
      case "#": this.err("Macro parameter character # in a formula"); return { t: "err", v: "#" };
      case "\\": return null;
      default: break;
    }
    if ("+=<>!?,;:/|()[].".includes(c)) {
      const o = { t: "o", v: c };
      if (FENCE_CHARS.has(c)) o.stretchy = false;
      return o;
    }
    if (/\p{L}/u.test(c)) return { t: "i", v: mathFontChar(c, this.font) };
    if (/\p{N}/u.test(c)) return { t: "n", v: c };
    return { t: "o", v: c };
  }

  readDelim() {
    const t = this.next();
    if (!t) { this.err("Missing delimiter after \\left or \\right"); return ""; }
    if (t.k === "ch") {
      if (t.v === ".") return "";
      if ("()[]|/".includes(t.v)) return t.v;
      if (t.v === "<") return "⟨";
      if (t.v === ">") return "⟩";
      this.err(`Bad delimiter "${t.v}" after \\left or \\right`);
      return t.v;
    }
    if (DELIMS[t.v] !== undefined) return DELIMS[t.v];
    this.err(`Bad delimiter \\${t.v} after \\left or \\right`);
    return "";
  }

  expandMacro(name) {
    const m = this.macros.get(name);
    if (++this.exp > MAX_MATH_EXPANSIONS || this.n > 400000) {
      this.err(`Macro \\${name} expands too often (recursive definition?)`);
      this.i = this.n;
      return false;
    }
    const args = [];
    let first = 0;
    if (m.def !== null && m.def !== undefined) {
      const o = this.readRawOpt();
      args.push(o === null ? m.def : o);
      first = 1;
    }
    for (let k = first; k < m.n; k++) args.push(this.readRawArg());
    const body = m.src.replace(/##|#([1-9])/g, (all, d) => (d ? (args[d - 1] ?? "") : "#"));
    this.s = body + " " + this.s.slice(this.i);
    this.i = 0;
    this.n = this.s.length;
    return true;
  }

  command(name) {
    if (this.macros && this.macros.has(name)) {
      if (!this.expandMacro(name)) return { t: "err", v: "\\" + name };
      return this.parseAtom();
    }
    if (GREEK_LOWER[name]) return { t: "i", v: mathFontChar(GREEK_LOWER[name], this.font) };
    if (GREEK_UPPER[name]) return { t: "i", v: mathFontChar(GREEK_UPPER[name], this.font), normal: !this.font || this.font === "rm" };
    if (MATH_ORD[name]) return { t: "i", v: MATH_ORD[name], normal: true };
    if (LARGE_OPS[name]) return { t: "o", v: LARGE_OPS[name][0], large: true, lim: LARGE_OPS[name][1] };
    if (Object.prototype.hasOwnProperty.call(FUNCS, name)) return { t: "fn", v: FUNC_TEXT[name] || name, lim: !!FUNCS[name] };
    if (MATH_OPS[name]) {
      const v = MATH_OPS[name];
      return FENCE_CHARS.has(v) ? { t: "o", v, stretchy: false } : { t: "o", v };
    }
    if (DELIMS[name] && name.length === 1) return { t: "o", v: DELIMS[name], stretchy: false };
    if (Object.prototype.hasOwnProperty.call(SPACES, name)) return { t: "space", w: SPACES[name] };
    if (MATH_FONTS[name] !== undefined) {
      const saved = this.font;
      this.font = MATH_FONTS[name];
      const arg = this.parseArg();
      this.font = saved;
      return arg && arg.t === "row" ? { ...arg, group: true } : arg;
    }
    if (MATH_TEXT_CMDS[name]) return this.textCommand(MATH_TEXT_CMDS[name]);
    if (MATH_ACCENTS[name]) {
      const [acc, stretchy, under] = MATH_ACCENTS[name];
      const base = this.parseArg();
      const node = { t: "accent", c: base, acc, stretchy, under: !!under, name };
      if (name === "overbrace" || name === "underbrace") node.lim = true, node.limits = true;
      return node;
    }
    switch (name) {
      case "frac": case "dfrac": case "tfrac": case "cfrac": {
        if (name === "cfrac") this.readRawOpt();
        const saved = this.display;
        this.display = false;
        const n = this.parseArg();
        const d = this.parseArg();
        this.display = saved;
        return { t: "frac", n, d, style: name === "dfrac" || name === "cfrac" ? "d" : name === "tfrac" ? "t" : null };
      }
      case "binom": case "dbinom": case "tbinom": {
        const saved = this.display;
        this.display = false;
        const n = this.parseArg();
        const d = this.parseArg();
        this.display = saved;
        return { t: "frac", n, d, thick: "0", fence: ["(", ")"], style: name === "dbinom" ? "d" : name === "tbinom" ? "t" : null };
      }
      case "genfrac": {
        const l = this.readRawArg().trim(), r = this.readRawArg().trim(), th = this.readRawArg().trim(), st = this.readRawArg().trim();
        const n = this.parseArg(), d = this.parseArg();
        const dl = (x) => (x === "." || !x ? "" : x.startsWith("\\") ? (DELIMS[x.slice(1)] ?? "") : x);
        return { t: "frac", n, d, thick: th === "0pt" || th === "0" ? "0" : null, fence: (l || r) ? [dl(l), dl(r)] : null, style: st === "0" ? "d" : st === "1" ? "t" : null };
      }
      case "sqrt": {
        const idx = this.readRawOpt();
        const body = this.parseArg();
        return { t: "sqrt", c: body, idx: idx !== null && idx.trim() ? this.sub(idx, false) : null };
      }
      case "left": {
        const open = this.readDelim();
        const c = this.parseList(new Set(["right"]));
        const t = this.peek();
        let close = "";
        if (t && t.k === "cmd" && t.v === "right") { this.next(); close = this.readDelim(); } else this.err("Missing \\right after \\left");
        return { t: "fenced", open, close, c };
      }
      case "middle": return { t: "o", v: this.readDelim(), stretchy: true, fence: true };
      case "big": case "Big": case "bigg": case "Bigg":
      case "bigl": case "Bigl": case "biggl": case "Biggl":
      case "bigr": case "Bigr": case "biggr": case "Biggr":
      case "bigm": case "Bigm": case "biggm": case "Biggm": {
        const size = BIG_SIZES[name.replace(/[lrm]$/, "")];
        const v = this.readDelim();
        return v ? { t: "o", v, size, stretchy: true } : { t: "row", c: [] };
      }
      case "not": {
        const nx = this.parseAtom();
        if (nx && (nx.t === "o" || nx.t === "i") && nx.v) {
          return { t: "o", v: NOT_MAP[nx.v] || nx.v + "̸" };
        }
        return { t: "row", c: [{ t: "o", v: "/" }, nx || { t: "row", c: [] }] };
      }
      case "neq": return { t: "o", v: "≠" };
      case "textcolor": {
        const model = this.readRawOpt();
        const spec = this.readRawArg();
        const body = this.parseArg();
        return { t: "style", color: model ? colorFromModel(model, spec) : resolveColor(spec, this.colors), c: [body] };
      }
      case "colorbox": {
        const model = this.readRawOpt();
        const spec = this.readRawArg();
        const body = this.parseArg();
        return { t: "style", bg: model ? colorFromModel(model, spec) : resolveColor(spec, this.colors), c: [body] };
      }
      case "boxed": return { t: "box", c: this.parseArg() };
      case "phantom": case "hphantom": return { t: "phantom", c: this.parseArg() };
      case "vphantom": this.parseArg(); return { t: "row", c: [] };
      case "smash": this.readRawOpt(); return this.parseArg();
      case "cancel": case "bcancel": case "xcancel": case "sout": return { t: "cancel", c: this.parseArg() };
      case "overset": case "stackrel": {
        const top = this.parseArg();
        const base = this.parseArg();
        return { t: "scripts", base: { ...wrapRow(base), limits: true }, sup: top, sub: null, disp: true };
      }
      case "underset": {
        const bot = this.parseArg();
        const base = this.parseArg();
        return { t: "scripts", base: { ...wrapRow(base), limits: true }, sub: bot, sup: null, disp: true };
      }
      case "xrightarrow": case "xleftarrow": case "xRightarrow": case "xLeftarrow": case "xleftrightarrow":
      case "xLeftrightarrow": case "xmapsto": case "xhookrightarrow": case "xhookleftarrow": case "xlongequal": {
        const below = this.readRawOpt();
        const above = this.parseArg();
        const arrow = { xrightarrow: "→", xleftarrow: "←", xRightarrow: "⇒", xLeftarrow: "⇐", xleftrightarrow: "↔", xLeftrightarrow: "⇔", xmapsto: "↦", xhookrightarrow: "↪", xhookleftarrow: "↩", xlongequal: "=" }[name];
        return { t: "scripts", base: { t: "o", v: arrow, stretchy: true, limits: true, xarrow: true }, sup: above, sub: below ? this.sub(below, false) : null, disp: true };
      }
      case "substack": {
        const raw = this.readRawArg();
        const p = new MathParser(raw, { errors: this.errors, macros: this.macros, depth: this.depth + 1 });
        return p.tableBody("substack", { align: ["c"], small: true });
      }
      case "pmod": {
        const arg = this.parseArg();
        return { t: "row", c: [{ t: "space", w: 1 }, { t: "o", v: "(", stretchy: false }, { t: "fn", v: "mod" }, { t: "space", w: 0.3333 }, arg, { t: "o", v: ")", stretchy: false }] };
      }
      case "pod": {
        const arg = this.parseArg();
        return { t: "row", c: [{ t: "space", w: 1 }, { t: "o", v: "(", stretchy: false }, arg, { t: "o", v: ")", stretchy: false }] };
      }
      case "bmod": return { t: "o", v: "mod" };
      case "mod": return { t: "row", c: [{ t: "space", w: 1 }, { t: "fn", v: "mod" }, { t: "space", w: 0.3333 }] };
      case "operatorname": case "operatornamewithlimits": {
        let star = false;
        if (this.s[this.i] === "*") { star = true; this.i++; }
        const raw = this.readRawArg();
        return { t: "fn", v: textModeString(raw.replace(/\\,/g, " ")).trim() || "?", lim: star || name === "operatornamewithlimits" };
      }
      case "mathop": {
        const a = this.parseArg();
        return a && a.t === "row" && a.c.length === 1 ? { ...a.c[0], lim: true } : { ...wrapRow(a), lim: true };
      }
      case "mathbin": case "mathrel": case "mathord": case "mathopen": case "mathclose": case "mathpunct": case "mathinner":
      case "ensuremath": case "displaylines": case "hbox_": return this.parseArg();
      case "begin": {
        const env = this.readRawArg().trim();
        return this.environment(env);
      }
      case "hspace": case "hskip": case "kern": case "mkern": case "mskip": case "mspace": {
        let raw;
        if (name === "hspace") { if (this.s[this.i] === "*") this.i++; raw = this.readRawArg(); } else if (name === "mspace") raw = this.readRawArg(); else raw = this.readDimen();
        return { t: "space", w: dimToEm(raw) };
      }
      case "label": case "tag": case "eqref_": {
        if (this.s[this.i] === "*") this.i++;
        this.readRawArg();
        return null;
      }
      case "nonumber": case "notag": case "displaybreak": case "allowdisplaybreaks": case "mathstrut": case "strut":
      case "nolimits": case "limits": case "relax": case "protect": case "nobreak": case "allowbreak": case "noindent":
      case "centering": case "hfill": case "hfil": case "vfill": case "par": case "quad_":
        return null;
      case "eqref": case "ref": return { t: "text", v: "(" + this.readRawArg() + ")" };
      case "lbrace": return { t: "o", v: "{", stretchy: false };
      case "matrix": case "pmatrix_": {
        const raw = this.readRawArg();
        const p = new MathParser(raw, { errors: this.errors, macros: this.macros, depth: this.depth + 1 });
        return p.tableBody("matrix", { align: ["c"] });
      }
      case "LaTeX": return { t: "text", v: "LaTeX" };
      case "TeX": return { t: "text", v: "TeX" };
      case "dots": case "ldots": return { t: "o", v: "…" };
      case "unicode": this.readRawArg(); return { t: "o", v: "?" };
      case "vcenter": case "raisebox": { if (name === "raisebox") this.readRawArg(); return this.parseArg(); }
      default: break;
    }
    this.err(`Unknown command \\${name} in a formula`);
    return { t: "err", v: "\\" + name };
  }

  textCommand(style) {
    const raw = this.readRawArg();
    const parts = [];
    // nested math inside \text{...}: split on unescaped $
    const re = /(?<!\\)\$([^$]*)\$/g;
    let last = 0, m;
    while ((m = re.exec(raw))) {
      if (m.index > last) parts.push({ t: "text", v: textModeString(raw.slice(last, m.index)), style });
      parts.push(this.sub(m[1], false));
      last = re.lastIndex;
    }
    if (last < raw.length) parts.push({ t: "text", v: textModeString(raw.slice(last)), style });
    if (!parts.length) return { t: "text", v: "", style };
    return parts.length === 1 ? parts[0] : { t: "row", c: parts };
  }

  environment(env) {
    const name = env.replace(/\*$/, "");
    const fences = {
      pmatrix: ["(", ")"], bmatrix: ["[", "]"], Bmatrix: ["{", "}"], vmatrix: ["|", "|"], Vmatrix: ["‖", "‖"],
      cases: ["{", ""], dcases: ["{", ""], rcases: ["", "}"], drcases: ["", "}"], psmallmatrix: ["(", ")"],
      bsmallmatrix: ["[", "]"],
    };
    let align = ["c"], aligned = false, small = false, disp = null;
    let colLines = null;
    if (name === "array" || name === "subarray" || name === "darray") {
      this.readRawOpt();
      const spec = this.readRawArg();
      const r = parseColSpec(spec);
      align = r.aligns.length ? r.aligns : ["c"];
      colLines = r.lines;
    } else if (name === "aligned" || name === "split" || name === "alignedat" || name === "align" || name === "alignat" || name === "flalign" || name === "eqnarray") {
      if (name === "alignedat" || name === "alignat") this.readRawArg();
      align = name === "eqnarray" ? ["r", "c", "l"] : ["r", "l"];
      aligned = name !== "eqnarray";
      disp = true;
    } else if (name === "gathered" || name === "gather" || name === "multline" || name === "equation") {
      align = ["c"]; disp = true;
    } else if (name === "cases" || name === "dcases" || name === "rcases" || name === "drcases") {
      align = ["l", "l"];
      if (name.startsWith("d")) disp = true;
    } else if (name === "smallmatrix" || name === "psmallmatrix" || name === "bsmallmatrix") {
      small = true;
    } else if (name.endsWith("matrix")) {
      const o = this.readRawOpt();
      if (o && /^[lcr]$/.test(o.trim())) align = [o.trim()];
    } else if (name !== "matrix") {
      this.err(`Unknown environment ${env} in a formula`);
    }
    const node = this.tableBody(env, { align, aligned, small, fence: fences[name] || null, disp, colLines });
    return node;
  }

  /** Rows/cells until \end{env} (or end of input for \substack/\matrix bodies). */
  tableBody(env, o) {
    const rows = [];
    let row = [];
    const stops = new Set(["&", "\\\\", "end"]);
    const saved = this.display;
    if (o.disp !== null && o.disp !== undefined) this.display = o.disp; else this.display = false;
    for (let guard = 0; guard < 100000; guard++) {
      const cell = this.parseList(stops);
      row.push(cell);
      const t = this.next();
      if (!t) {
        if (env !== "substack" && env !== "matrix") this.err(`Missing \\end{${env}} in a formula`);
        break;
      }
      if (t.k === "ch" && t.v === "&") continue;
      if (t.k === "cmd" && (t.v === "\\" || t.v === "cr" || t.v === "newline")) { this.skipRowOpts(); rows.push(row); row = []; continue; }
      if (t.k === "cmd" && t.v === "end") {
        const nm = this.readRawArg().trim();
        if (nm !== env) this.err(`\\begin{${env}} ended by \\end{${nm}}`);
        break;
      }
    }
    this.display = saved;
    if (row.length > 1 || (row.length === 1 && row[0].length)) rows.push(row);
    // \hline markers -> row lines
    const lines = [];
    const outRows = [];
    for (const r of rows) {
      let top = false;
      while (r[0] && r[0].length && r[0][0].t === "hline") { r[0].shift(); top = true; }
      const onlyLine = r.length === 1 && r[0].length === 0;
      if (onlyLine && top) { if (outRows.length) lines[outRows.length - 1] = (lines[outRows.length - 1] || 0) | 2; else lines[0] = 1; continue; }
      if (top) lines[outRows.length] = (lines[outRows.length] || 0) | 1;
      outRows.push(r.map((c) => c.filter((x) => x.t !== "hline")));
    }
    return { t: "table", rows: outRows, align: o.align, aligned: !!o.aligned, small: !!o.small, fence: o.fence, rowLines: lines, colLines: o.colLines || null, env };
  }
}

function wrapRow(n) { return n && n.t === "row" ? n : { t: "row", c: n ? [n] : [] }; }

function dimToEm(raw) {
  const m = /(-?\s*[\d.]+)\s*(em|ex|pt|mm|cm|in|mu|bp|pc)?/.exec(String(raw || ""));
  if (!m) return 0.25;
  const v = parseFloat(m[1].replace(/\s/g, ""));
  const f = { em: 1, ex: 0.43, pt: 0.1, bp: 0.1, mm: 0.284, cm: 2.84, in: 7.2, mu: 1 / 18, pc: 1.2 }[m[2] || "pt"] || 0.1;
  return Math.max(-3, Math.min(20, v * f));
}

/** Column spec of tabular/array: aligns per column and vertical line flags per boundary. */
function parseColSpec(spec) {
  const s = String(spec || "");
  const aligns = [], widths = [], lines = [0];
  let i = 0, guard = 0;
  let expanded = s;
  // expand *{n}{spec}
  for (let k = 0; k < 5 && /\*\s*\{/.test(expanded); k++) {
    expanded = expanded.replace(/\*\s*\{(\d+)\}\s*\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}/g, (m, n, sp) => sp.repeat(Math.min(50, Number(n) || 0)));
  }
  const S = expanded;
  i = 0;
  const read = () => {
    while (S[i] === " ") i++;
    if (S[i] !== "{") { const c = S[i] || ""; i++; return c; }
    let d = 0;
    const st = i;
    for (; i < S.length; i++) {
      if (S[i] === "{") d++;
      else if (S[i] === "}") { d--; if (d === 0) { i++; return S.slice(st + 1, i - 1); } }
    }
    return S.slice(st + 1);
  };
  while (i < S.length && guard++ < 500) {
    const c = S[i++];
    if (c === "l" || c === "c" || c === "r") { aligns.push(c); widths.push(null); lines[aligns.length] = 0; }
    else if (c === "p" || c === "m" || c === "b") { const w = read(); aligns.push("l"); widths.push(w); lines[aligns.length] = 0; }
    else if (c === "X") { aligns.push("l"); widths.push("X"); lines[aligns.length] = 0; }
    else if (c === "S") { aligns.push("c"); widths.push(null); lines[aligns.length] = 0; }
    else if (c === "L" || c === "C" || c === "R") { aligns.push(c.toLowerCase()); widths.push(null); lines[aligns.length] = 0; }
    else if (c === "|") lines[aligns.length] = (lines[aligns.length] || 0) + 1;
    else if (c === "@" || c === "!" || c === ">" || c === "<") read();
    else if (c === "\\") { while (i < S.length && /[A-Za-z]/.test(S[i])) i++; }
  }
  return { aligns, widths, lines };
}

// ===========================================================================
// math: AST -> MathML Core
// ===========================================================================

function isFenceOpen(n) { return n && n.t === "o" && /^[([{⟨⌊⌈|‖]$/.test(n.v) && n.stretchy === false; }
function baseOf(n) { let b = n; while (b && b.t === "scripts") b = b.base; return b; }

function mmlRowContent(nodes) {
  let out = "";
  for (let k = 0; k < nodes.length; k++) {
    const nd = nodes[k];
    if (!nd) continue;
    out += mml(nd);
    const b = baseOf(nd);
    if (b && b.t === "fn") {
      const nx = nodes[k + 1];
      const nb = baseOf(nx);
      if (nx && !(nb && nb.t === "o" && (isFenceOpen(nb) || nb.v === "(")) && !(nb && nb.t === "fenced") && !(nb && nb.t === "space")) out += '<mspace width="0.1667em"/>';
    }
  }
  return out;
}
function mmlWrap(n) {
  if (!n) return "<mrow></mrow>";
  if (n.t === "row") return n.c.length === 1 ? mml(n.c[0]) : `<mrow>${mmlRowContent(n.c)}</mrow>`;
  return mml(n);
}
const colorAttr = (n) => [n.color ? `color:${n.color}` : "", n.bg ? `background-color:${n.bg}` : ""].filter(Boolean).join(";");

function mml(n) {
  if (!n) return "";
  switch (n.t) {
    case "i": return `<mi${n.normal && [...n.v].length === 1 ? ' mathvariant="normal"' : ""}>${esc(n.v)}</mi>`;
    case "n": return `<mn>${esc(n.v)}</mn>`;
    case "o": {
      let a = "";
      if (n.stretchy === false) a += ' stretchy="false"';
      if (n.stretchy === true) a += ' stretchy="true"';
      if (n.size) a += ` minsize="${n.size}" maxsize="${n.size}"`;
      if (n.fence) a += ' fence="true"';
      if (n.form) a += ` form="${n.form}"`;
      if (n.movable === false) a += ' movablelimits="false"';
      if (n.prime) a += ' lspace="0" rspace="0"';
      return `<mo${a}>${esc(n.v)}</mo>`;
    }
    case "fn": return `<mi${[...n.v].length === 1 ? ' mathvariant="normal"' : ""}>${esc(n.v)}</mi>`;
    case "text": {
      const st = n.style || {};
      const css = [st.b ? "font-weight:bold" : "", st.i ? "font-style:italic" : "", st.tt ? "font-family:monospace" : "", st.sf ? "font-family:sans-serif" : "", st.box ? "border:1px solid currentColor;padding:1px 3px" : ""].filter(Boolean).join(";");
      const v = esc(n.v).replace(/^ +| +$/g, (m) => "&#160;".repeat(m.length));
      return `<mtext${css ? ` style="${css}"` : ""}>${v}</mtext>`;
    }
    case "row": return n.c.length === 1 && !n.c[0].limits ? mml(n.c[0]) : `<mrow>${mmlRowContent(n.c)}</mrow>`;
    case "frac": {
      let f = `<mfrac${n.thick === "0" ? ' linethickness="0"' : ""}>${mmlWrap(n.n)}${mmlWrap(n.d)}</mfrac>`;
      if (n.style === "d") f = `<mstyle displaystyle="true" scriptlevel="0">${f}</mstyle>`;
      else if (n.style === "t") f = `<mstyle displaystyle="false">${f}</mstyle>`;
      if (n.fence) f = `<mrow>${n.fence[0] ? `<mo stretchy="true" fence="true" form="prefix">${esc(n.fence[0])}</mo>` : ""}${f}${n.fence[1] ? `<mo stretchy="true" fence="true" form="postfix">${esc(n.fence[1])}</mo>` : ""}</mrow>`;
      return f;
    }
    case "sqrt": return n.idx ? `<mroot>${mmlWrap(n.c)}${mmlWrap(n.idx)}</mroot>` : `<msqrt>${n.c && n.c.t === "row" ? mmlRowContent(n.c.c) : mml(n.c)}</msqrt>`;
    case "scripts": {
      const b = baseOf(n.base) || n.base;
      const limits = n.base.limits === true || (n.base.limits !== false && (n.base.lim || b.lim) && n.disp);
      let baseXml;
      if (limits && n.base.t === "o") baseXml = mml({ ...n.base, movable: false });
      else baseXml = mmlWrap(n.base);
      if (limits) {
        if (n.sub && n.sup) return `<munderover>${baseXml}${mmlWrap(n.sub)}${mmlWrap(n.sup)}</munderover>`;
        if (n.sub) return `<munder>${baseXml}${mmlWrap(n.sub)}</munder>`;
        return `<mover>${baseXml}${mmlWrap(n.sup)}</mover>`;
      }
      if (n.sub && n.sup) return `<msubsup>${baseXml}${mmlWrap(n.sub)}${mmlWrap(n.sup)}</msubsup>`;
      if (n.sub) return `<msub>${baseXml}${mmlWrap(n.sub)}</msub>`;
      return `<msup>${baseXml}${mmlWrap(n.sup)}</msup>`;
    }
    case "fenced": {
      const o = n.open ? `<mo fence="true" stretchy="true" symmetric="true" form="prefix">${esc(n.open)}</mo>` : "";
      const c = n.close ? `<mo fence="true" stretchy="true" symmetric="true" form="postfix">${esc(n.close)}</mo>` : "";
      return `<mrow>${o}${mmlRowContent(n.c)}${c}</mrow>`;
    }
    case "accent": {
      const mo = `<mo stretchy="${n.stretchy ? "true" : "false"}">${esc(n.acc)}</mo>`;
      if (n.under) return `<munder accentunder="${n.name === "underbrace" ? "false" : "true"}">${mmlWrap(n.c)}${mo}</munder>`;
      return `<mover accent="${n.name === "overbrace" ? "false" : "true"}">${mmlWrap(n.c)}${mo}</mover>`;
    }
    case "table": return mmlTable(n);
    case "space": return `<mspace width="${Math.round(n.w * 10000) / 10000}em"/>`;
    case "style": {
      const inner = mmlRowContent(n.c);
      if (n.disp) {
        const map = { displaystyle: ' displaystyle="true" scriptlevel="0"', textstyle: ' displaystyle="false" scriptlevel="0"', scriptstyle: ' displaystyle="false" scriptlevel="1"', scriptscriptstyle: ' displaystyle="false" scriptlevel="2"' };
        return `<mstyle${map[n.disp]}>${inner}</mstyle>`;
      }
      const css = colorAttr(n);
      return `<mrow${css ? ` style="${css}"` : ""}>${inner}</mrow>`;
    }
    case "box": return `<mrow style="border:1px solid currentColor;padding:2px 4px">${mmlWrap(n.c)}</mrow>`;
    case "phantom": return `<mphantom>${mmlWrap(n.c)}</mphantom>`;
    case "cancel": return `<mrow style="background:linear-gradient(to top right,transparent calc(50% - 0.6px),currentColor,transparent calc(50% + 0.6px))">${mmlWrap(n.c)}</mrow>`;
    case "err": return `<merror><mtext>${esc(n.v)}</mtext></merror>`;
    case "hline": return "";
    default: return "";
  }
}

function mmlTable(n) {
  const ncols = Math.max(1, ...n.rows.map((r) => r.length));
  let rows = "";
  n.rows.forEach((r, ri) => {
    let cells = "";
    for (let ci = 0; ci < ncols; ci++) {
      const cell = r[ci] || [];
      let al;
      if (n.aligned) al = ci % 2 === 0 ? "right" : "left";
      else al = { l: "left", c: "center", r: "right" }[n.align[Math.min(ci, n.align.length - 1)]] || "center";
      const css = [`text-align:${al}`];
      if (n.aligned) {
        css.push(ci % 2 === 0 ? `padding-right:0;padding-left:${ci === 0 ? 0 : 1}em` : "padding-left:0;padding-right:0");
      } else if (n.env && /cases/.test(n.env)) {
        css.push(ci === 0 ? "padding-left:0.2em;padding-right:1em" : "padding-left:0");
      }
      if (n.colLines) {
        if (ci === 0 && n.colLines[0]) css.push("border-left:1px solid currentColor");
        if (n.colLines[ci + 1]) css.push("border-right:1px solid currentColor");
      }
      if (n.rowLines) {
        if (n.rowLines[ri] & 1) css.push("border-top:1px solid currentColor");
        if (n.rowLines[ri] & 2) css.push("border-bottom:1px solid currentColor");
      }
      // In aligned environments the right column starts with an operator: add an
      // empty mrow before it so relations get their infix spacing.
      const lead = n.aligned && ci % 2 === 1 ? "<mrow></mrow>" : "";
      cells += `<mtd columnalign="${al}" style="${css.join(";")}">${lead}${mmlRowContent(cell)}</mtd>`;
    }
    rows += `<mtr>${cells}</mtr>`;
  });
  const attrs = [];
  if (n.small) attrs.push('displaystyle="false" scriptlevel="1"');
  else if (n.aligned || n.env === "gathered" || n.env === "dcases" || n.env === "eqnarray") attrs.push('displaystyle="true"');
  else attrs.push('displaystyle="false"');
  const colalign = n.aligned ? Array.from({ length: ncols }, (_, k) => (k % 2 ? "left" : "right")).join(" ") : null;
  let t = `<mtable ${attrs.join(" ")}${colalign ? ` columnalign="${colalign}"` : ""}${n.aligned ? ' columnspacing="0em"' : ""}>${rows}</mtable>`;
  if (n.fence) {
    t = `<mrow>${n.fence[0] ? `<mo fence="true" stretchy="true" symmetric="true" form="prefix">${esc(n.fence[0])}</mo>` : ""}${t}${n.fence[1] ? `<mo fence="true" stretchy="true" symmetric="true" form="postfix">${esc(n.fence[1])}</mo>` : ""}</mrow>`;
  }
  return t;
}

// ===========================================================================
// math: AST -> Unicode approximation (word model, plain text)
// ===========================================================================

const isSimpleText = (s) => /^[\p{L}\p{N}.′″]+$/u.test(s) && [...s].length <= 6;
function scriptText(s, map, mark) {
  if (!s) return "";
  const chars = [...s];
  if (chars.every((c) => map[c] !== undefined)) return chars.map((c) => map[c]).join("");
  return mark + (chars.length === 1 ? s : `(${s})`);
}

const UNI_SPACED = new Set(["=", "<", ">", "≤", "≥", "≠", "≈", "≡", "∼", "≃", "≅", "∝", "→", "←", "↔", "⇒", "⇐", "⇔", "⟹", "⟸", "⟺", "↦", "∈", "∉", "⊂", "⊃", "⊆", "⊇", "+", "−", "±", "∓", "×", "⋅", "÷", "∪", "∩", "≪", "≫", "∣", "∥", "⊥", "≔"]);
function uniList(nodes) {
  let out = "";
  for (let k = 0; k < nodes.length; k++) {
    const nd = nodes[k];
    if (!nd) continue;
    const s = uni(nd);
    if (nd.t === "o" && UNI_SPACED.has(nd.v)) {
      if (out.trim()) { out = out.replace(/ +$/, "") + " " + s + " "; continue; }
      if (!"+−±∓".includes(nd.v)) { out += s + " "; continue; }
    }
    out += s;
    const b = baseOf(nd);
    if (b && b.t === "fn") {
      const nx = nodes[k + 1];
      const nb = baseOf(nx);
      if (nx && nb && (nb.t === "i" || nb.t === "n" || nb.t === "fn" || nb.t === "frac")) out += " ";
    }
  }
  return out;
}
function uni(n) {
  if (!n) return "";
  switch (n.t) {
    case "i": case "n": case "fn": return n.v;
    case "o": return n.v;
    case "text": return n.v;
    case "row": return uniList(n.c).replace(/  +/g, " ");
    case "frac": {
      const a = uni(n.n).trim(), b = uni(n.d).trim();
      if (n.thick === "0") return `${n.fence ? n.fence[0] : ""}${a} over ${b}${n.fence ? n.fence[1] : ""}`.replace(/ over /, n.fence && n.fence[0] === "(" ? " choose " : " over ");
      return `${isSimpleText(a) ? a : `(${a})`}/${isSimpleText(b) ? b : `(${b})`}`;
    }
    case "sqrt": {
      const a = uni(n.c).trim();
      const ix = n.idx ? uni(n.idx).trim() : "";
      const root = ix === "3" ? "∛" : ix === "4" ? "∜" : ix ? `${scriptText(ix, SUP_MAP, "")}√` : "√";
      return root + (isSimpleText(a) ? a : `(${a})`);
    }
    case "scripts": {
      let s = uni(n.base);
      if (n.base && n.base.t === "row" && n.base.c.length > 1 && !n.base.group) s = `(${s})`;
      if (n.sub) s += scriptText(uni(n.sub).trim(), SUB_MAP, "_");
      if (n.sup) s += scriptText(uni(n.sup).trim(), SUP_MAP, "^");
      return s;
    }
    case "fenced": return `${n.open}${uniList(n.c)}${n.close}`;
    case "accent": {
      const s = uni(n.c);
      const mark = COMBINING[n.name];
      if (n.name === "overbrace" || n.name === "underbrace") return s;
      if (mark && [...s].length === 1) return s + mark;
      if (n.name === "vec" || n.name === "overrightarrow") return s + "⃗";
      return s + (mark || "");
    }
    case "table": {
      const rows = n.rows.map((r) => r.map((c) => uniList(c).trim()).join(n.aligned ? "" : " ")).join("; ");
      return n.fence ? `${n.fence[0]}${rows}${n.fence[1]}` : rows;
    }
    case "space": return n.w >= 0.9 ? " " : n.w > 0.2 ? " " : n.w > 0 ? " " : "";
    case "style": case "cancel": return uniList(n.c);
    case "box": return uni(n.c);
    case "phantom": return "";
    case "err": return n.v;
    default: return "";
  }
}

function convertMath(tex, display, opts = {}) {
  const src = String(tex == null ? "" : tex);
  try {
    const p = new MathParser(src, { display, macros: opts.macros || null });
    p.colors = opts.colors || null;
    const nodes = p.parseTop();
    const body = mmlRowContent(nodes);
    const alt = opts.alt === false ? "" : ` alttext="${esc(src.replace(/\s+/g, " ").trim().slice(0, 300))}"`;
    const xml = `<math${display ? ' display="block"' : ""}${alt}>${nodes.length === 1 ? body : `<mrow>${body}</mrow>`}</math>`;
    return { xml, errors: p.errors, nodes };
  } catch (e) {
    return { xml: `<math${display ? ' display="block"' : ""}><merror><mtext>${esc(src)}</mtext></merror></math>`, errors: ["This formula could not be read"], nodes: [{ t: "text", v: src }] };
  }
}

/** TeX math -> MathML Core markup ("<math>…</math>"). Never throws; problems become <merror>. */
export function mathToMathML(tex, display = false, opts = {}) {
  return convertMath(tex, !!display, opts).xml;
}

/** TeX math -> readable Unicode approximation ("x² + √(a+b)"). Never throws. */
export function mathToUnicode(tex, opts = {}) {
  try {
    const r = convertMath(tex, false, { ...opts, alt: false });
    return uniList(r.nodes).replace(/\s+/g, " ").trim();
  } catch {
    return String(tex);
  }
}

// ===========================================================================
// document parser (tokens -> block/inline AST)
// ===========================================================================

/*
 * Blocks:
 *   {type:"par", inl, line, align?, noindent?}
 *   {type:"heading", cmd, level, num, title:inl, line, ids}
 *   {type:"title", line}                      (\maketitle / \titlepage)
 *   {type:"list", kind:"ul"|"ol"|"dl", items:[{label, mark, blocks, line}], line}
 *   {type:"quote"|"abstract"|"div"|"proof", blocks, align?, width?, line}
 *   {type:"code", text, lang, line}
 *   {type:"math", env, rows:[{cells:[{xml, nodes}], num, ids}], line}
 *   {type:"table", rows:[{cells:[{blocks, colspan, rowspan, align, lines}], top, bottom, line}], line}
 *   {type:"float", kind, blocks, caption, num, captionTop, ids, line}
 *   {type:"theorem", title, num, note, style, blocks, ids, line}
 *   {type:"frame", title, subtitle, blocks, line}
 *   {type:"pagebreak"|"toc"|"hr"|"vspace", line}
 *   {type:"bib", title, items:[{num, key, inl, line}], line}
 *   {type:"placeholder", text, line}
 *   {type:"letter", to, blocks, line}
 *   {type:"columns", blocks:[div...], line}
 * Inline nodes:
 *   {t:"text", v, st}  {t:"math", xml, nodes, tex, st}  {t:"br"}  {t:"fn", n, st}
 *   {t:"ref", key, kind, st}  {t:"cite", keys, note, st}  {t:"code", v, st}  {t:"img", file, width, st}
 */

const BASE_STYLE = Object.freeze({});
const styleKeys = new WeakMap();
function styleKey(st) {
  let k = styleKeys.get(st);
  if (k === undefined) { k = JSON.stringify(st); styleKeys.set(st, k); }
  return k;
}
function withStyle(st, patch) {
  const n = { ...st };
  for (const key of Object.keys(patch)) {
    const v = patch[key];
    if (v === undefined || v === false || v === null) delete n[key];
    else n[key] = v;
  }
  return Object.freeze(n);
}

const SECTION_LEVEL = { part: 0, chapter: 1, section: 2, subsection: 3, subsubsection: 4, paragraph: 5, subparagraph: 6 };
const THEOREM_FALLBACK = {
  theorem: "Theorem", lemma: "Lemma", proposition: "Proposition", corollary: "Corollary", definition: "Definition",
  example: "Example", remark: "Remark", note: "Note", exercise: "Exercise", problem: "Problem", solution: "Solution",
  claim: "Claim", conjecture: "Conjecture", fact: "Fact", observation: "Observation", question: "Question",
  answer: "Answer", assumption: "Assumption", hypothesis: "Hypothesis", notation: "Notation", axiom: "Axiom",
  thm: "Theorem", lem: "Lemma", prop: "Proposition", cor: "Corollary", defn: "Definition", rem: "Remark",
};
const DISPLAY_MATH_ENVS = new Set(["equation", "equation*", "displaymath", "align", "align*", "alignat", "alignat*", "flalign", "flalign*", "gather", "gather*", "multline", "multline*", "eqnarray", "eqnarray*", "math", "dmath", "dmath*"]);
const LIST_ENVS = { itemize: "ul", enumerate: "ol", description: "dl", compactitem: "ul", compactenum: "ol", compactdesc: "dl", inparaenum: "ol", asparaenum: "ol", asparaitem: "ul", inparaitem: "ul", list: "ul", checklist: "ul", tasks: "ol" };
const FLOAT_ENVS = new Set(["figure", "figure*", "table", "table*", "wrapfigure", "wraptable", "subfigure", "subtable", "sidewaysfigure", "sidewaystable", "SCfigure", "algorithm", "algorithm*", "listing", "marginfigure", "margintable"]);
const TABULAR_ENVS = new Set(["tabular", "tabular*", "tabularx", "tabulary", "longtable", "longtable*", "array", "tabu", "supertabular", "xtabular", "NiceTabular", "tblr", "longtblr"]);
const SKIP_ENVS = new Set(["tikzpicture", "pgfpicture", "picture", "circuitikz", "forest", "tikzcd", "axis", "pspicture", "asy", "dot2tex", "sagesilent", "pycode"]);
const PASS_ENVS = new Set(["titlepage", "small", "footnotesize", "scriptsize", "large", "Large", "normalsize", "singlespace", "doublespace", "onehalfspace", "spacing", "adjustbox", "landscape", "samepage", "fullwidth", "sloppypar", "frame_", "appendices", "subappendices", "otherlanguage", "otherlanguage*", "english", "refsection", "linenumbers", "mdframed", "tcolorbox", "framed", "shaded", "leftbar", "resizebox", "scope", "threeparttable", "tablenotes", "noindent", "addmargin", "onlyenv", "overlayarea", "overprint", "actionenv", "uncoverenv", "visibleenv", "alertenv", "structureenv", "itemize_"]);
const SIZE_CMDS = SIZES;

const IGNORE_ARGS = {
  // name: [optional args before, mandatory args]
  usepackage: [1, 1], RequirePackage: [1, 1], documentstyle: [1, 1], pagestyle: [0, 1], thispagestyle: [0, 1],
  pagenumbering: [0, 1], setlength: [0, 2], addtolength: [0, 2], settowidth: [0, 2], settoheight: [0, 2],
  linespread: [0, 1], geometry: [0, 1], hypersetup: [0, 1], lstset: [0, 1], graphicspath: [0, 1],
  bibliographystyle: [0, 1], captionsetup: [1, 1], newlength: [0, 1], usetheme: [1, 1], usecolortheme: [1, 1],
  usefonttheme: [1, 1], useinnertheme: [1, 1], useoutertheme: [1, 1], setbeamertemplate: [0, 1], setbeamercolor: [0, 2],
  setbeamerfont: [0, 2], setbeamersize: [0, 1], titlespacing: [0, 4], titleformat: [1, 5], newgeometry: [0, 1],
  addbibresource: [1, 1], newcounter: [1, 1], numberwithin: [1, 2], setcounter_: [0, 2], counterwithin: [0, 2],
  counterwithout: [0, 2], DeclareGraphicsExtensions: [0, 1], setmainfont: [1, 1], setsansfont: [1, 1],
  setmonofont: [1, 1], newfontfamily: [1, 2], urlstyle: [0, 1], definecolorset: [0, 4], selectlanguage: [0, 1],
  setlist: [1, 1], tcbuselibrary: [0, 1], usetikzlibrary: [0, 1], pgfplotsset: [0, 1], allowdisplaybreaks_: [1, 0],
  fancyhf: [0, 1], fancyhead: [1, 1], fancyfoot: [1, 1], lhead: [1, 1], chead: [1, 1], rhead: [1, 1], lfoot: [1, 1],
  cfoot: [1, 1], rfoot: [1, 1], setstretch: [0, 1], onehalfspacing_: [0, 0], hyphenation: [0, 1],
  vspace: [0, 1], addvspace: [0, 1], enlargethispage: [0, 1], nocite: [0, 1], index: [0, 1], glossary: [0, 1],
  addcontentsline: [0, 3], addtocontents: [0, 2], markboth: [0, 2], markright: [0, 1], setkomafont: [0, 2],
  theoremstyle_: [0, 1], newtheoremstyle: [0, 9], AtBeginSection: [1, 1], AtBeginSubsection: [1, 1],
  logo: [0, 1], titlegraphic: [0, 1], mode: [0, 0], resetcounteronoverlays: [0, 1], includeonly: [0, 1],
  DeclareUnicodeCharacter: [0, 2], newunicodechar: [0, 2], SetKwInOut: [0, 2], crefname: [0, 3], Crefname: [0, 3],
  bibpunct: [1, 6], setcitestyle: [0, 1], babelprovide: [1, 1], sisetup: [0, 1], date_: [0, 0], frenchspacing_: [0, 0],
  pdfbookmark: [1, 2], phantomsection_: [0, 0], ExecuteOptions: [0, 1], ProcessOptions_: [0, 0], PassOptionsToPackage: [0, 2],
  setbeameroption: [0, 1], institute_: [1, 1], hbox_: [0, 0], vskip_: [0, 0], arraystretch_: [0, 0],
};
const SILENT = new Set([
  "noindent", "indent", "centering_", "par_", "relax", "protect", "null", "leavevmode", "nobreak", "allowbreak", "sloppy",
  "fussy", "frenchspacing", "nonfrenchspacing", "selectfont", "makeatletter", "makeatother", "onehalfspacing",
  "doublespacing", "singlespacing", "raggedbottom", "flushbottom", "normalcolor", "unskip", "ignorespaces",
  "vfill", "vfil", "break", "filbreak", "goodbreak", "nopagebreak", "nolinebreak", "linebreak_", "strut", "mbox_",
  "phantomsection", "listoffigures", "listoftables", "printindex", "makeindex", "makeglossaries", "printglossaries",
  "maketitle_", "pause", "displaybreak", "allowdisplaybreaks", "arraybackslash", "tabularnewline_", "hline", "cline_",
  "toprule", "midrule", "bottomrule", "endhead", "endfirsthead", "endfoot", "endlastfoot", "addlinespace",
  "cmidrule_", "specialrule_", "hdashline", "FloatBarrier", "clearpage_", "justifying", "RaggedRight",
  "itshape_", "nointerlineskip", "offinterlineskip", "smallbreak", "medbreak", "bigbreak", "titlepage_",
  "tableofcontents_", "textwidth", "linewidth", "columnwidth", "textheight", "paperwidth", "hsize", "vsize",
  "baselineskip", "parindent", "parskip", "fill", "stretch_", "the", "today_", "cleardoublepage_", "numberwithin_",
  "headrule", "footrule", "rmfamily_", "boldmath", "unboldmath", "label_", "appendixpage", "addappheadtotoc",
  "begingroup", "endgroup", "bgroup", "egroup", "inputencoding", "frontmatter", "mainmatter", "backmatter",
  "maketitle__", "tightlist", "bfseries_", "AND", "And", "endinput", "printbibliography_", "newblock",
  "item_", "lineskip", "small_", "nocite_", "noexpand", "expandafter", "global", "long", "outer", "immediate",
]);

function braceMatch(toks) {
  const m = new Int32Array(toks.length).fill(-1);
  const st = [];
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    if (t.k !== "ch") continue;
    if (t.v === "{") st.push(i);
    else if (t.v === "}" && st.length) m[st.pop()] = i;
  }
  return m;
}

/** Split a math body into rows on top-level \\ (outside braces and nested environments). */
function splitTopLevel(src, sep) {
  const out = [];
  let depth = 0, envDepth = 0, start = 0;
  const s = src;
  for (let i = 0; i < s.length; i++) {
    const c = s[i];
    if (c === "\\") {
      const nx = s[i + 1];
      if (nx === "\\" && sep === "\\\\" && depth === 0 && envDepth === 0) {
        out.push(s.slice(start, i));
        i += 2;
        while (s[i] === " " || s[i] === "*") i++;
        if (s[i] === "[") { const e = s.indexOf("]", i); if (e > 0) i = e + 1; }
        start = i;
        i--;
        continue;
      }
      if (/[A-Za-z]/.test(nx || "")) {
        let j = i + 1;
        while (j < s.length && /[A-Za-z]/.test(s[j])) j++;
        const name = s.slice(i + 1, j);
        if (name === "begin") envDepth++;
        else if (name === "end") envDepth = Math.max(0, envDepth - 1);
        i = j - 1;
        continue;
      }
      i++;
      continue;
    }
    if (c === "%") { const e = s.indexOf("\n", i); i = e < 0 ? s.length : e; continue; }
    if (c === "{") depth++;
    else if (c === "}") depth = Math.max(0, depth - 1);
    else if (c === "&" && sep === "&" && depth === 0 && envDepth === 0) { out.push(s.slice(start, i)); start = i + 1; }
  }
  out.push(s.slice(start));
  return out;
}

class DocParser {
  constructor(src) {
    this.src = String(src == null ? "" : src);
    this.errors = [];
    this.toks = tokenize(this.src);
    this.match = braceMatch(this.toks);
    this.pos = 0;
    this.stack = [];
    this.line = 1;
    this.macros = new Map();
    this.envDefs = new Map();
    this.theorems = new Map();
    this.thmCounters = new Map();
    this.thmStyle = "plain";
    this.colors = new Map();
    this.labels = new Map();
    this.refNodes = [];
    this.citeNodes = [];
    this.footnotes = [];
    this.bibMap = new Map();
    this.unknownCmds = new Map();
    this.outline = [];
    this.docclass = "article";
    this.hasChapters = false;
    this.isBeamer = false;
    this.isLetter = false;
    this.title = null; this.author = null; this.date = null; this.subtitle = null; this.institute = null;
    this.address = null; this.signature = null;
    this.counters = { part: 0, chapter: 0, section: 0, subsection: 0, subsubsection: 0, paragraph: 0, subparagraph: 0, equation: 0, figure: 0, table: 0, footnote: 0, algorithm: 0, frame: 0 };
    this.appendix = false;
    this.root = [];
    this.cur = this.root;
    this.para = null;
    this.paraLine = 1;
    this.noindent = false;
    this.style = BASE_STYLE;
    this.align = null;
    this.frames = [];
    this.hasDocEnv = /\\begin\s*\{document\}/.test(this.src);
    this.inBody = !this.hasDocEnv;
    this.finished = false;
    this.expansions = 0;
    this.lastTarget = null;
    this.listDepth = 0;
    this.enumDepth = 0;
    this.curFloat = null;
    this.curFrame = null;
    this.tableCtx = null;
    this.baseSize = 10;
    this.idSeq = 0;
    this.preambleTextWarned = false;
    this.imageCount = 0;
    this.skipCount = 0;
    this.depth = 0;
  }

  // ---- diagnostics
  error(line, message) { this.diag(line, message, "error"); }
  warn(line, message) { this.diag(line, message, "warning"); }
  diag(line, message, severity) {
    const l = line || this.line;
    const key = l + "|" + message;
    if (!this.diagSeen) this.diagSeen = new Set();
    if (this.diagSeen.has(key) || this.errors.length >= 500) return;
    this.diagSeen.add(key);
    this.errors.push({ line: l, message, severity });
  }

  // ---- input
  next() {
    while (this.stack.length) {
      const f = this.stack[this.stack.length - 1];
      if (f.p < f.t.length) { const t = f.t[f.p++]; if (t.l) this.line = t.l; return t; }
      this.stack.pop();
    }
    if (this.pos < this.toks.length) { const t = this.toks[this.pos++]; this.line = t.l; return t; }
    return null;
  }
  peek() {
    for (let k = this.stack.length - 1; k >= 0; k--) {
      const f = this.stack[k];
      if (f.p < f.t.length) return f.t[f.p];
    }
    return this.pos < this.toks.length ? this.toks[this.pos] : null;
  }
  fromMain() {
    for (let k = this.stack.length - 1; k >= 0; k--) if (this.stack[k].p < this.stack[k].t.length) return false;
    return true;
  }
  unread(tok) { this.pushTokens([tok]); }
  pushTokens(arr) {
    while (this.stack.length && this.stack[this.stack.length - 1].p >= this.stack[this.stack.length - 1].t.length) this.stack.pop();
    if (!arr.length) return;
    if (this.stack.length > 2000) { this.error(this.line, "Macros nested too deeply (is a command defined in terms of itself?)"); return; }
    this.stack.push({ t: arr, p: 0 });
  }
  skipSpaces() {
    for (;;) {
      const t = this.peek();
      if (t && t.k === "txt" && /^\s/.test(t.v)) {
        this.next();
        const rest = t.v.replace(/^\s+/, "");
        if (rest) { this.unread({ k: "txt", v: rest, l: t.l }); return; }
        continue;
      }
      return;
    }
  }
  readStar() {
    const t = this.peek();
    if (t && t.k === "ch" && t.v === "*") { this.next(); return true; }
    return false;
  }

  /** Mandatory argument: token array (balanced group or one token). */
  readArg(cmd) {
    this.skipSpaces();
    let t = this.peek();
    if (!t) { this.error(this.line, `Missing argument for \\${cmd}`); return []; }
    if (t.k === "ch" && t.v === "{") {
      if (this.fromMain()) {
        const start = this.pos;
        const end = this.match[start];
        if (end >= 0) { this.pos = end + 1; this.line = this.toks[end].l; return this.toks.slice(start + 1, end); }
        let e = start + 1;
        while (e < this.toks.length && this.toks[e].k !== "par") e++;
        this.error(t.l, `Missing } for the argument of \\${cmd} that starts on line ${t.l}`);
        this.pos = e;
        return this.toks.slice(start + 1, e);
      }
      this.next();
      const out = [];
      let depth = 1;
      for (let guard = 0; guard < 2000000; guard++) {
        const x = this.next();
        if (!x) { this.error(t.l, `Missing } for the argument of \\${cmd} that starts on line ${t.l}`); break; }
        if (x.k === "sentinel") { this.unread(x); this.error(t.l, `Missing } for the argument of \\${cmd} that starts on line ${t.l}`); break; }
        if (x.k === "ch" && x.v === "{") depth++;
        else if (x.k === "ch" && x.v === "}") { if (--depth === 0) break; }
        out.push(x);
      }
      return out;
    }
    if (t.k === "par" || (t.k === "ch" && t.v === "}") || t.k === "sentinel") {
      this.error(t.l || this.line, `Missing argument for \\${cmd}`);
      return [];
    }
    this.next();
    if (t.k === "txt") {
      const ch = String.fromCodePoint(t.v.codePointAt(0));
      const rest = t.v.slice(ch.length);
      if (rest) this.unread({ k: "txt", v: rest, l: t.l });
      return [{ k: "txt", v: ch, l: t.l }];
    }
    return [t];
  }

  /** Optional [..] argument: token array or null. */
  readOpt() {
    let t = this.peek();
    let skipped = null;
    if (t && t.k === "txt" && /^\s+$/.test(t.v)) { skipped = this.next(); t = this.peek(); }
    if (!t || t.k !== "ch" || t.v !== "[") { if (skipped) this.unread(skipped); return null; }
    this.next();
    const out = [];
    let depth = 0, br = 0;
    for (let guard = 0; guard < 100000; guard++) {
      const x = this.next();
      if (!x) { this.error(t.l, "Missing ] for an optional argument"); break; }
      if (x.k === "par" || x.k === "sentinel") { this.unread(x); this.error(t.l, "Missing ] for an optional argument"); break; }
      if (x.k === "ch") {
        if (x.v === "{") depth++;
        else if (x.v === "}") depth--;
        else if (x.v === "[" && depth === 0) br++;
        else if (x.v === "]" && depth === 0) { if (br === 0) break; br--; }
      }
      out.push(x);
    }
    return out;
  }
  argText(cmd) { return plainText(this.readArg(cmd)).trim(); }
  /** Beamer overlay specification <...> right after a command. */
  skipOverlay() {
    const t = this.peek();
    if (t && t.k === "txt" && /^\s*<[^>]*>/.test(t.v)) {
      this.next();
      const rest = t.v.replace(/^\s*<[^>]*>/, "");
      if (rest) this.unread({ k: "txt", v: rest, l: t.l });
    }
  }

  // ---- output helpers
  addText(v, line) {
    if (!this.inBody) {
      if (!this.preambleTextWarned && /\S/.test(v)) {
        this.preambleTextWarned = true;
        this.error(line, "Text before \\begin{document} (move it into the document body)");
      }
      return;
    }
    if (!this.para) {
      if (!/\S/.test(v)) return;
      this.para = [];
      this.paraLine = line || this.line;
    }
    const last = this.para[this.para.length - 1];
    if (last && last.t === "text" && last.st === this.style) last.v += v;
    else this.para.push({ t: "text", v, st: this.style });
  }
  addInline(node) {
    if (!this.inBody) return;
    if (!this.para) { this.para = []; this.paraLine = node.l || this.line; }
    if (node.st === undefined) node.st = this.style;
    this.para.push(node);
  }
  flushPar() {
    const p = this.para;
    this.para = null;
    if (!p || !p.length) { this.noindent = false; return; }
    // trim
    while (p.length && p[0].t === "text" && !p[0].v.trim()) p.shift();
    if (p.length && p[0].t === "text") p[0] = { ...p[0], v: p[0].v.replace(/^\s+/, "") };
    while (p.length && p[p.length - 1].t === "text" && !p[p.length - 1].v.trim()) p.pop();
    while (p.length && p[p.length - 1].t === "br") p.pop();
    if (p.length && p[p.length - 1].t === "text") p[p.length - 1] = { ...p[p.length - 1], v: p[p.length - 1].v.replace(/\s+$/, "") };
    if (!p.length) { this.noindent = false; return; }
    const b = { type: "par", inl: p, line: this.paraLine };
    if (this.align) b.align = this.align;
    if (this.noindent) b.noindent = true;
    this.noindent = false;
    this.cur.push(b);
  }
  addBlock(b) {
    this.flushPar();
    if (!this.inBody) return b;
    if (!b.line) b.line = this.line;
    this.cur.push(b);
    return b;
  }
  saveCtx() { return { cur: this.cur, para: this.para, paraLine: this.paraLine, style: this.style, align: this.align, noindent: this.noindent }; }
  restoreCtx(c) { this.cur = c.cur; this.para = c.para; this.paraLine = c.paraLine; this.style = c.style; this.align = c.align; this.noindent = c.noindent; }
  /** Run `fn` collecting blocks into a fresh container; returns the blocks. */
  container(fn) {
    const saved = this.saveCtx();
    this.cur = [];
    this.para = null;
    let reason;
    try { reason = fn(); } finally { this.flushPar(); }
    const blocks = this.cur;
    const style = saved.style;
    this.restoreCtx(saved);
    this.style = style;
    return { blocks, reason };
  }
  /** Parse tokens (an argument) in the current context. */
  parseTokens(toks) {
    if (!toks || !toks.length) return "sentinel";
    if (this.depth > MAX_DEPTH) { this.error(this.line, "Document nested too deeply"); return "sentinel"; }
    const sentinel = { k: "sentinel" };
    this.pushTokens([...toks, sentinel]);
    this.frames.push({ type: "arg", line: this.line });
    const r = this.parseUntil({ sentinel });
    this.frames.pop();
    return r;
  }
  /** Inline content of tokens (titles, captions, footnotes); paragraphs are joined. */
  captureInline(toks, resetStyle = false) {
    const saved = this.saveCtx();
    const savedBody = this.inBody;
    this.inBody = true;
    this.cur = [];
    this.para = null;
    if (resetStyle) this.style = BASE_STYLE;
    this.parseTokens(toks);
    this.flushPar();
    const out = [];
    const collect = (blocks) => {
      for (const b of blocks) {
        if (b.type === "par") { if (out.length) out.push({ t: "text", v: " ", st: BASE_STYLE }); out.push(...b.inl); }
        else if (b.blocks) collect(b.blocks);
      }
    };
    collect(this.cur);
    this.inBody = savedBody;
    this.restoreCtx(saved);
    return out;
  }
  newId(prefix) { return `ltx-${prefix}-${++this.idSeq}`; }
  setTarget(num, kind, block, title) {
    this.lastTarget = { num: num == null ? "" : String(num), kind, block, title: title || "" };
  }

  // ---- main loop
  /**
   * mode: {group, env, item, cell, sentinel, top}
   * returns "}" | "end" | "item" | "&" | "\\\\" | "sentinel" | "eof" | "unwind" | "docend"
   */
  parseUntil(mode) {
    if (++this.depth > MAX_DEPTH) {
      this.depth--;
      this.error(this.line, "Document nested too deeply (too many groups or environments inside each other)");
      // skip to the end of input to stay safe
      this.stack = [];
      this.pos = this.toks.length;
      return "eof";
    }
    try {
      for (;;) {
        if (this.finished) return "docend";
        const t = this.next();
        if (!t) return "eof";
        switch (t.k) {
          case "txt": this.addText(textLigatures(t.v), t.l); break;
          case "par":
            if (mode.cell && this.tableCtx && !this.tableCtx.allowPar) { this.addText(" ", t.l); break; }
            this.flushPar();
            break;
          case "ch": {
            const r = this.special(t, mode);
            if (r) return r;
            break;
          }
          case "param": this.addText(t.v === "#" ? "#" : "#" + t.v, t.l); break;
          case "verb": this.verbBlock(t); break;
          case "skip": if (!t.closed) this.error(t.l, `\\begin{${t.env}} on line ${t.l} is never closed`); break;
          case "vinl": this.addInline({ t: "code", v: t.v, l: t.l }); break;
          case "sentinel": if (t === mode.sentinel) return "sentinel"; break;
          case "cmd": {
            const r = this.command(t, mode);
            if (r) return r;
            break;
          }
          default: break;
        }
      }
    } finally {
      this.depth--;
    }
  }

  special(t, mode) {
    switch (t.v) {
      case "{": this.group(t); return null;
      case "}": {
        if (mode.group) return "}";
        if (this.frames.some((f) => f.type === "group")) {
          const f = this.frames[this.frames.length - 1];
          if (f && f.type === "env") this.error(f.line, `\\begin{${f.name}} on line ${f.line} is not closed before the } on line ${t.l} (missing \\end{${f.name}})`);
          this.unread(t);
          return "unwind";
        }
        this.error(t.l, "Extra } (it has no matching {)");
        return null;
      }
      case "$": case "$$": this.inlineOrDisplayMath(t); return null;
      case "&":
        if (mode.cell) return "&";
        if (this.tableCtx && this.frames.some((f) => f.type === "env" && f.name === this.tableCtx.env)) {
          this.error(t.l, "& inside a { } group in a table cell (close the group before &)");
          return null;
        }
        this.error(t.l, "Misplaced & (only allowed in tables and align environments; write \\& for an ampersand)");
        return null;
      case "~": this.addText(" ", t.l); return null;
      case "^": case "_":
        this.error(t.l, `${t.v} is only allowed in math (write \\${t.v} or put it inside $...$)`);
        this.addText(t.v, t.l);
        return null;
      case "#": this.error(t.l, "# is only allowed in macro definitions (write \\#)"); this.addText("#", t.l); return null;
      default: this.addText(t.v, t.l); return null;
    }
  }

  group(t) {
    const savedStyle = this.style, savedAlign = this.align, savedTarget = this.lastTarget;
    this.frames.push({ type: "group", line: t.l });
    const r = this.parseUntil({ group: true });
    this.frames.pop();
    if (r === "eof") this.error(t.l, `Missing } for the { on line ${t.l}`);
    this.style = savedStyle;
    // \centering inside a group applies to paragraphs ended inside it
    this.align = savedAlign;
    this.lastTarget = savedTarget;
  }

  // ---- math
  collectMath(close, open, envName) {
    const out = [];
    let depth = 0, nest = 0;
    for (let guard = 0; guard < 5000000; guard++) {
      const t = this.next();
      if (!t) { this.error(open.l, `Math started on line ${open.l} is never closed (missing ${close})`); break; }
      if (t.k === "sentinel") { this.unread(t); this.error(open.l, `Math started on line ${open.l} is not closed (missing ${close})`); break; }
      if (t.k === "par") {
        if (envName && envName !== "math") { this.error(t.l, `Blank line inside \\begin{${envName}} (line ${open.l}); remove it or use % on that line`); continue; }
        this.unread(t);
        this.error(open.l, `Math started on line ${open.l} is not closed before the blank line (missing ${close})`);
        break;
      }
      if (t.k === "ch") {
        if (t.v === "{") depth++;
        else if (t.v === "}") depth--;
        if (close === "$" && depth <= 0 && t.v === "$") break;
        if (close === "$" && depth <= 0 && t.v === "$$") { this.unread({ k: "ch", v: "$", l: t.l }); break; }
        if (close === "$$" && t.v === "$$") break;
        if (close === "$$" && t.v === "$" && depth <= 0) { this.error(t.l, "Display math started with $$ must end with $$"); break; }
      } else if (t.k === "cmd") {
        if (close === "\\)" && t.v === ")") break;
        if (close === "\\]" && t.v === "]") break;
        if ((t.v === ")" || t.v === "]") && close !== "\\)" && close !== "\\]" && depth <= 0) {
          this.error(t.l, `\\${t.v} without a matching \\${t.v === ")" ? "(" : "["}`);
          continue;
        }
        if (t.v === "end" || t.v === "begin") {
          const arg = this.readArg(t.v);
          const name = plainText(arg).trim();
          if (envName && name === envName) {
            if (t.v === "begin") nest++;
            else if (nest === 0) break;
            else nest--;
          }
          if (t.v === "end" && name === "document") {
            this.error(open.l, `Math started on line ${open.l} is never closed (missing ${close})`);
            this.pushTokens([t, { k: "ch", v: "{", l: t.l }, ...arg, { k: "ch", v: "}", l: t.l }]);
            break;
          }
          out.push(t, { k: "ch", v: "{", l: t.l }, ...arg, { k: "ch", v: "}", l: t.l });
          continue;
        }
      }
      out.push(t);
    }
    return out;
  }

  mathNode(tex, display, line) {
    const r = convertMath(tex, display, { macros: this.macros, colors: this.colors });
    for (const e of r.errors) {
      if (/^Unknown command/.test(e)) this.warn(line, e);
      else this.error(line, e);
    }
    return { xml: r.xml, nodes: r.nodes, tex };
  }

  inlineOrDisplayMath(t) {
    if (t.v === "$$") {
      const body = this.collectMath("$$", t);
      this.displayMath("displaymath", serialize(body), t.l);
      return;
    }
    const body = this.collectMath("$", t);
    const tex = serialize(body);
    if (!this.inBody) return;
    const m = this.mathNode(tex, false, t.l);
    this.addInline({ t: "math", ...m, l: t.l });
  }

  displayMath(env, src, line) {
    if (!this.inBody) return;
    const name = env.replace(/\*$/, "");
    const starred = env.endsWith("*") || name === "displaymath" || name === "math";
    let body = src;
    if (name === "alignat") body = body.replace(/^\s*\{[^}]*\}/, "");
    const multi = ["align", "flalign", "alignat", "eqnarray", "gather", "multline"].includes(name);
    let rowsSrc = multi ? splitTopLevel(body, "\\\\") : [body];
    if (rowsSrc.length > 1 && !rowsSrc[rowsSrc.length - 1].replace(/\\(label|tag)\s*\{[^}]*\}|\\(nonumber|notag)|%[^\n]*/g, "").trim()) {
      const lastRow = rowsSrc.pop();
      if (/\\label/.test(lastRow)) rowsSrc[rowsSrc.length - 1] += lastRow;
    }
    const rows = rowsSrc.map((r) => {
      let tex = r, label = null, tag = null, nonum = false;
      tex = tex.replace(/\\label\s*\{([^}]*)\}/g, (m, k) => { label = label || k.trim(); return ""; });
      tex = tex.replace(/\\tag(\*?)\s*\{([^}]*)\}/g, (m, star, x) => { tag = star ? textModeString(x) : `(${textModeString(x)})`; return ""; });
      if (/\\(nonumber|notag)(?![A-Za-z])/.test(tex)) { nonum = true; tex = tex.replace(/\\(nonumber|notag)(?![A-Za-z])/g, ""); }
      return { tex, label, tag, nonum };
    });
    const eqNum = () => {
      this.counters.equation++;
      return this.hasChapters ? `${this.chapterNum()}.${this.counters.equation}` : String(this.counters.equation);
    };
    const block = { type: "math", env: name, rows: [], line };
    rows.forEach((r, k) => {
      let num = null;
      if (r.tag) num = r.tag;
      else if (!starred && !r.nonum) {
        if (name === "equation" || name === "dmath") num = k === 0 ? `(${eqNum()})` : null;
        else if (name === "multline") num = k === rows.length - 1 ? `(${eqNum()})` : null;
        else num = `(${eqNum()})`;
      }
      const cellsSrc = name === "align" || name === "flalign" || name === "alignat" || name === "eqnarray" ? splitTopLevel(r.tex, "&") : [r.tex];
      const cells = cellsSrc.map((c, k) => this.mathNode(k % 2 && name !== "eqnarray" ? "{}" + c : c, true, line));
      const row = { cells, num, ids: [] };
      if (r.label) {
        const id = this.newId("eq");
        row.ids.push(id);
        this.defineLabel(r.label, num ? num.replace(/^\((.*)\)$/, "$1") : "", "equation", id, line);
      }
      block.rows.push(row);
    });
    this.addBlock(block);
  }

  // ---- labels
  defineLabel(key, num, kind, id, line, title) {
    if (!key) return;
    if (this.labels.has(key)) this.warn(line, `Label "${key}" is defined more than once`);
    this.labels.set(key, { num: num == null ? "" : String(num), kind, id, title: title || "" });
  }
  chapterNum() { return this.appendix && this.appendixLevel === "chapter" ? alpha(this.counters.chapter, true) : String(this.counters.chapter); }

  // ---- verbatim
  verbBlock(t) {
    if (!t.closed) this.error(t.l, `\\begin{${t.env}} on line ${t.l} is never closed (missing \\end{${t.env}})`);
    if (t.env.startsWith("filecontents")) return;
    let lang = "";
    const m = /\{([^}]*)\}/.exec(t.opt || "");
    if (t.env === "minted" && m) lang = m[1];
    const lm = /language\s*=\s*([\w+#-]+)/i.exec(t.opt || "");
    if (lm) lang = lm[1];
    const cap = /caption\s*=\s*\{?([^,}\]]*)/i.exec(t.opt || "");
    const b = this.addBlock({ type: "code", text: t.v, lang, line: t.l });
    if (cap && cap[1].trim()) b.caption = cap[1].trim();
  }

  // ---- commands
  command(t, mode) {
    const name = t.v;
    if (this.macros.has(name)) { this.expand(name, t); return null; }
    // control symbols
    if (name.length === 1 && !/[A-Za-z]/.test(name)) return this.controlSymbol(t, mode);
    const h = CMD[name];
    if (h) return h.call(this, t, mode) || null;
    if (SIZE_CMDS[name] !== undefined) { this.style = withStyle(this.style, { size: name === "normalsize" ? undefined : SIZE_CMDS[name] * (this.baseSize / 10) }); return null; }
    if (Object.prototype.hasOwnProperty.call(TEXT_SYMBOLS, name)) { this.addText(TEXT_SYMBOLS[name], t.l); return null; }
    if (IGNORE_ARGS[name]) {
      const [o, m] = IGNORE_ARGS[name];
      this.readStar();
      for (let k = 0; k < o; k++) this.readOpt();
      for (let k = 0; k < m; k++) this.readArg(name);
      return null;
    }
    if (SILENT.has(name)) return null;
    if (TEXT_ACCENTS[name] && name.length === 1) return this.accent(t);
    // math-only commands with arguments used in text (\mathbb{R}, \frac{a}{b}): render as math, tell the user
    const nArgs = MATH_FONTS[name] !== undefined ? 1 : { frac: 2, dfrac: 2, tfrac: 2, binom: 2, sqrt: 1, hat: 1, vec: 1, bar: 1, tilde: 1, dot: 1, overrightarrow: 1, boldsymbol: 1, operatorname: 1 }[name];
    if (nArgs) {
      let tex = "\\" + name;
      if (name === "sqrt") { const o = this.readOpt(); if (o) tex += `[${serialize(o)}]`; }
      for (let k = 0; k < nArgs; k++) tex += `{${serialize(this.readArg(name))}}`;
      if (this.inBody) this.addInline({ t: "math", ...this.mathNode(tex, false, t.l), l: t.l });
      if (!this.mathInTextWarned) { this.mathInTextWarned = true; this.warn(t.l, `\\${name} is a math command: put it inside $...$`); }
      return null;
    }
    // math-only symbols used in text: render them, but tell the user
    const sym =GREEK_LOWER[name] || GREEK_UPPER[name] || MATH_ORD[name] || MATH_OPS[name];
    if (sym) {
      this.addText(sym, t.l);
      if (!this.mathInTextWarned) { this.mathInTextWarned = true; this.warn(t.l, `\\${name} is a math command: put it inside $...$`); }
      return null;
    }
    // unknown: drop the command, keep its arguments
    if (!this.unknownCmds.has(name)) this.unknownCmds.set(name, t.l);
    this.readStar();
    this.readOpt();
    return null;
  }

  controlSymbol(t, mode) {
    const v = t.v;
    if (v === "\\") {
      this.readStar();
      this.readOpt();
      if (mode.cell) return "\\\\";
      if (this.tableCtx && mode.group && this.frames.some((f) => f.type === "env" && f.name === this.tableCtx.env)) { this.addInline({ t: "br" }); return null; }
      if (this.para) this.addInline({ t: "br" });
      return null;
    }
    if (v === "(") {
      const body = this.collectMath("\\)", t);
      if (this.inBody) this.addInline({ t: "math", ...this.mathNode(serialize(body), false, t.l), l: t.l });
      return null;
    }
    if (v === "[") {
      const body = this.collectMath("\\]", t);
      this.displayMath("displaymath", serialize(body), t.l);
      return null;
    }
    if (v === ")" || v === "]") { this.error(t.l, `\\${v} without a matching \\${v === ")" ? "(" : "["}`); return null; }
    if (TEXT_ACCENTS[v]) return this.accent(t);
    if (v === " ") { this.addText(" ", t.l); return null; }
    if (Object.prototype.hasOwnProperty.call(TEXT_SYMBOLS, v)) { this.addText(TEXT_SYMBOLS[v], t.l); return null; }
    this.addText(v, t.l);
    return null;
  }

  accent(t) {
    const mark = TEXT_ACCENTS[t.v];
    const arg = this.readArg(t.v);
    let base = plainText(arg);
    if (!base && arg.length === 1 && arg[0].k === "cmd") base = TEXT_SYMBOLS[arg[0].v] || "";
    if (base === "ı") base = "i";
    if (base === "ȷ") base = "j";
    const first = base ? String.fromCodePoint(base.codePointAt(0)) : " ";
    this.addText((first + mark).normalize("NFC") + base.slice(first.length), t.l);
    return null;
  }

  expand(name, t) {
    if (++this.expansions > MAX_EXPANSIONS) {
      if (!this.expansionError) { this.expansionError = true; this.error(t.l, `Too many macro expansions (is \\${name} defined in terms of itself?)`); }
      return;
    }
    const m = this.macros.get(name);
    if (m.alias) { this.pushTokens([{ k: "cmd", v: m.alias, l: t.l }]); return; }
    const args = [];
    let first = 0;
    if (m.def !== null && m.def !== undefined) {
      const o = this.readOpt();
      args.push(o === null ? m.defToks : o);
      first = 1;
    }
    for (let k = first; k < m.n; k++) args.push(this.readArg(name));
    const out = [];
    for (const bt of m.body) {
      if (bt.k === "param") {
        if (bt.v === "#") out.push({ k: "ch", v: "#", l: t.l });
        else { const a = args[Number(bt.v) - 1]; if (a) for (const x of a) out.push(x); }
      } else out.push(bt.l === t.l ? bt : { ...bt, l: t.l });
      if (out.length > 200000) break;
    }
    this.pushTokens(out);
  }

  defineMacro(nameTokens, n, defToks, body, provide, line) {
    const nt = nameTokens.find((x) => x.k === "cmd");
    if (!nt) { this.error(line, "\\newcommand needs a command name like \\mycmd"); return; }
    if (provide && this.macros.has(nt.v)) return;
    const cnt = Math.max(0, Math.min(9, Number(n) || 0));
    this.macros.set(nt.v, { n: cnt, def: defToks ? serialize(defToks) : null, defToks, body, src: serialize(body) });
  }

  // ---- sectioning
  section(t) {
    const cmd = t.v;
    const star = this.readStar();
    this.readOpt();
    const arg = this.readArg(cmd);
    if (!this.inBody) return;
    const level = SECTION_LEVEL[cmd];
    const title = this.captureInline(arg);
    const maxNumbered = this.hasChapters ? 3 : 4;
    let num = null;
    if (!star) {
      const order = ["part", "chapter", "section", "subsection", "subsubsection", "paragraph", "subparagraph"];
      const idx = order.indexOf(cmd);
      this.counters[cmd]++;
      for (let k = idx + 1; k < order.length; k++) this.counters[order[k]] = 0;
      if (cmd === "chapter") {
        this.counters.equation = 0; this.counters.figure = 0; this.counters.table = 0; this.counters.footnote = 0;
      }
      if (cmd === "part") num = roman(this.counters.part, true);
      else if (level <= maxNumbered) {
        const parts = [];
        const top = this.hasChapters ? 1 : 2;
        for (let k = top; k <= level; k++) {
          const c = order[k];
          let v = this.counters[c];
          if (this.appendix && k === top) v = alpha(v, true);
          parts.push(String(v));
        }
        num = parts.join(".");
      }
    }
    const plain = inlinePlain(title).trim();
    const block = { type: "heading", cmd, level, num, title, line: t.l, ids: [this.newId("sec")] };
    if (this.isBeamer) block.beamer = true;
    this.addBlock(block);
    this.outline.push({ level, title: plain || "(untitled)", line: t.l, num: num || "", id: block.ids[0] });
    this.setTarget(num || "", cmd, block, plain);
    this.noindent = true;
  }

  caption(t) {
    this.readStar();
    this.readOpt();
    const arg = this.readArg("caption");
    if (!this.inBody) return;
    const f = this.curFloat;
    if (!f) {
      if (this.tableCtx && this.tableCtx.env.startsWith("longtable")) {
        this.counters.table++;
        const num = this.hasChapters ? `${this.chapterNum()}.${this.counters.table}` : String(this.counters.table);
        this.tableCtx.caption = { inl: this.captureInline(arg), num };
        this.setTarget(num, "table", null);
        return;
      }
      this.error(t.l, "\\caption outside a figure or table");
      this.addBlock({ type: "par", inl: this.captureInline(arg), line: t.l, align: "center" });
      return;
    }
    let num;
    if (f.kind === "subfigure" || f.kind === "subtable") {
      const parent = f.parent;
      if (parent) parent.subCount = (parent.subCount || 0) + 1;
      num = `(${alpha(parent ? parent.subCount : 1, false)})`;
      f.refNum = parent && parent.num ? `${parent.num}${alpha(parent.subCount, false)}` : alpha(parent ? parent.subCount : 1, false);
    } else {
      const key = f.kind === "table" ? "table" : f.kind === "algorithm" ? "algorithm" : "figure";
      this.counters[key]++;
      num = this.hasChapters && key !== "algorithm" ? `${this.chapterNum()}.${this.counters[key]}` : String(this.counters[key]);
    }
    f.num = num;
    f.caption = this.captureInline(arg);
    f.captionTop = this.cur.length === 0 && !this.para;
    this.setTarget(f.refNum || num, f.kind, f);
  }

  // ---- environments
  beginEnv(t, mode) {
    const name = this.argText("begin");
    if (!name) { this.error(t.l, "\\begin needs an environment name"); return null; }
    if (name === "document") {
      if (this.inBody && this.hasDocEnv) { this.error(t.l, "\\begin{document} appears twice"); return null; }
      this.inBody = true;
      this.frames.push({ type: "env", name, line: t.l });
      const r = this.parseUntil({ env: "document" });
      this.frames.pop();
      this.flushPar();
      if (r === "end") { this.sawEnd = true; this.finished = true; return "docend"; }
      if (r !== "docend") this.error(this.lastLine(), "Missing \\end{document} at the end of the file");
      this.finished = true;
      return "docend";
    }
    if (this.frames.length > MAX_DEPTH) { this.error(t.l, "Environments nested too deeply"); return null; }
    const savedTarget = this.lastTarget;
    const savedAlign = this.align;
    const savedStyle = this.style;
    this.frames.push({ type: "env", name, line: t.l });
    try {
      this.environment(name, t, mode);
    } finally {
      this.frames.pop();
      this.style = savedStyle;
      this.align = savedAlign;
      this.lastTarget = savedTarget; // \label after \end{...} refers to the enclosing target again
    }
    return null;
  }
  lastLine() { return this.toks.length ? this.toks[this.toks.length - 1].l : 1; }

  /** Parse an environment body until \end{name}; reports unclosed environments. */
  body(name, t, extra = {}) {
    const r = this.parseUntil({ env: name, ...extra });
    if (r === "eof" || r === "docend") this.error(t.l, `\\begin{${name}} on line ${t.l} is never closed (missing \\end{${name}})`);
    return r;
  }

  environment(name, t, mode) {
    if (DISPLAY_MATH_ENVS.has(name)) {
      const body = this.collectMath(`\\end{${name}}`, t, name);
      this.displayMathFromEnv(name, body, t);
      return;
    }
    if (LIST_ENVS[name]) return this.list(name, t);
    if (name === "thebibliography") return this.bibliography(t);
    if (TABULAR_ENVS.has(name)) return this.tabular(name, t);
    if (FLOAT_ENVS.has(name)) return this.float(name, t);
    if (this.theorems.has(name)) return this.theorem(name, this.theorems.get(name), t);
    if (this.envDefs.has(name)) return this.userEnv(name, t);
    if (SKIP_ENVS.has(name)) return this.skipEnv(name, t);
    switch (name) {
      case "quote": case "quotation": case "verse": case "displayquote": case "quoting": {
        const { blocks } = this.container(() => this.body(name, t));
        this.addBlock({ type: "quote", env: name, blocks, line: t.l });
        return;
      }
      case "abstract": {
        const { blocks } = this.container(() => this.body(name, t));
        this.addBlock({ type: "abstract", blocks, line: t.l });
        return;
      }
      case "center": case "flushleft": case "flushright": case "centering": case "raggedright": case "raggedleft": {
        const align = name === "center" || name === "centering" ? "center" : name === "flushright" || name === "raggedleft" ? "right" : "left";
        const { blocks } = this.container(() => { this.align = align; return this.body(name, t); });
        this.addBlock({ type: "div", align, blocks, line: t.l });
        return;
      }
      case "minipage": case "column": case "parbox": {
        this.readOpt(); this.readOpt(); this.readOpt();
        const width = widthPercent(plainText(this.readArg(name)));
        let align = null;
        const { blocks } = this.container(() => { this.align = null; const res = this.body(name, t); align = this.align; return res; });
        this.addBlock({ type: "div", mini: true, width, blocks, line: t.l, column: name === "column", align });
        return;
      }
      case "columns": {
        this.readOpt();
        const { blocks } = this.container(() => this.body(name, t));
        this.addBlock({ type: "columns", blocks, line: t.l });
        return;
      }
      case "multicols": case "multicols*": {
        const n = Number(this.argText(name)) || 2;
        this.readOpt();
        const { blocks } = this.container(() => this.body(name, t));
        this.addBlock({ type: "div", cols: Math.min(6, Math.max(1, n)), blocks, line: t.l });
        return;
      }
      case "proof": {
        const note = this.readOpt();
        const noteInl = note ? this.captureInline(note) : null;
        const { blocks } = this.container(() => { this.style = BASE_STYLE; return this.body(name, t); });
        this.addBlock({ type: "proof", note: noteInl, blocks, line: t.l });
        return;
      }
      case "frame": return this.frame(t);
      case "block": case "alertblock": case "exampleblock": {
        this.skipOverlay();
        const title = this.peekIsBrace() ? this.captureInline(this.readArg(name)) : [];
        const { blocks } = this.container(() => this.body(name, t));
        this.addBlock({ type: "div", box: name, boxTitle: title, blocks, line: t.l });
        return;
      }
      case "letter": {
        const to = this.captureInline(this.readArg("letter"));
        const { blocks } = this.container(() => this.body(name, t));
        this.addBlock({ type: "letter", to, blocks, line: t.l });
        return;
      }
      case "algorithmic": case "algorithmic*": case "algpseudocode": {
        this.readOpt();
        const { blocks } = this.container(() => { this.algo = (this.algo || 0) + 1; const r = this.body(name, t); this.algo--; return r; });
        this.addBlock({ type: "div", algo: true, blocks, line: t.l });
        return;
      }
      case "comment": return;
      default: break;
    }
    if (PASS_ENVS.has(name) || SIZE_CMDS[name] !== undefined) {
      if (SIZE_CMDS[name] !== undefined) this.style = withStyle(this.style, { size: SIZE_CMDS[name] * (this.baseSize / 10) });
      if (name === "tcolorbox" || name === "mdframed" || name === "framed") this.readOpt();
      this.body(name, t);
      return;
    }
    if (THEOREM_FALLBACK[name]) {
      this.warn(t.l, `Environment "${name}" is not defined: add \\newtheorem{${name}}{${THEOREM_FALLBACK[name]}} to the preamble`);
      const def = { title: THEOREM_FALLBACK[name], counter: name, style: /definition|example|exercise|problem/.test(name) ? "definition" : /remark|note/.test(name) ? "remark" : "plain" };
      this.theorems.set(name, def);
      return this.theorem(name, def, t);
    }
    this.warn(t.l, `Unknown environment "${name}": its content is shown without special formatting`);
    const { blocks } = this.container(() => this.body(name, t));
    this.addBlock({ type: "div", blocks, line: t.l });
  }

  peekIsBrace() {
    this.skipSpaces();
    const p = this.peek();
    return !!(p && p.k === "ch" && p.v === "{");
  }

  skipEnv(name, t) {
    let depth = 0;
    for (let guard = 0; guard < 5000000; guard++) {
      const x = this.next();
      if (!x) { this.error(t.l, `\\begin{${name}} on line ${t.l} is never closed (missing \\end{${name}})`); break; }
      if (x.k === "cmd" && (x.v === "begin" || x.v === "end")) {
        const nm = this.argText(x.v);
        if (nm === name) { if (x.v === "begin") depth++; else if (depth === 0) break; else depth--; }
        if (nm === "document" && x.v === "end") { this.pushTokens([x, { k: "ch", v: "{", l: x.l }, { k: "txt", v: "document", l: x.l }, { k: "ch", v: "}", l: x.l }]); this.error(t.l, `\\begin{${name}} on line ${t.l} is never closed`); break; }
      }
    }
    const label = /tikz|pgf|forest|axis|circuit/.test(name) ? "TikZ drawing" : name === "picture" ? "Picture" : `${name} content`;
    this.addBlock({ type: "placeholder", text: `${label} (shown in the compiled PDF only)`, line: t.l });
  }

  displayMathFromEnv(name, bodyToks, t) {
    if (name === "math") {
      if (this.inBody) this.addInline({ t: "math", ...this.mathNode(serialize(bodyToks), false, t.l), l: t.l });
      return;
    }
    this.displayMath(name, serialize(bodyToks), t.l);
  }

  list(name, t) {
    const kind = LIST_ENVS[name];
    let opts = this.readOpt();
    if (name === "list") { this.readArg(name); this.readArg(name); }
    if (name === "tasks") this.readOpt(); // tasks(n)
    const isEnum = kind === "ol";
    if (isEnum) this.enumDepth++;
    this.listDepth++;
    const block = { type: "list", kind, items: [], line: t.l, depth: this.listDepth };
    // enumerate label format
    let fmt = null;
    let counter = 0;
    if (isEnum) {
      const lvl = Math.min(4, this.enumDepth);
      fmt = [null, { pre: "", style: "arabic", post: "." }, { pre: "(", style: "alph", post: ")" }, { pre: "", style: "roman", post: "." }, { pre: "", style: "Alph", post: "." }][lvl];
      if (opts) {
        const raw = serialize(opts);
        const lm = /label\s*=\s*(\{[^}]*\}|[^,]*)/.exec(raw);
        const startM = /start\s*=\s*(\d+)/.exec(raw);
        if (startM) counter = Number(startM[1]) - 1;
        const tpl = lm ? lm[1].replace(/^\{|\}$/g, "") : (/label|start|leftmargin|itemsep|noitemsep|nosep|topsep|resume|ref|align|wide|font/.test(raw) ? null : raw);
        const f = parseEnumLabel(tpl);
        if (f) fmt = f;
      }
    }
    block.fmt = fmt;
    try {
      // anything before the first \item
      const pre = this.container(() => this.parseUntil({ env: name, item: true }));
      let reason = pre.reason;
      if (pre.blocks.some((b) => b.type === "par")) this.error(t.l, `Text before the first \\item in ${name} (line ${t.l})`);
      for (let guard = 0; guard < 1000000 && reason === "item"; guard++) {
        const itemLine = this.line;
        this.skipOverlay();
        const label = this.readOpt();
        const item = { label: null, mark: null, blocks: [], line: itemLine };
        if (label !== null) item.label = this.captureInline(label);
        else if (isEnum) {
          counter++;
          const v = fmtCounter(counter, fmt.style);
          item.mark = `${fmt.pre}${v}${fmt.post}`;
          this.setTarget(v, "item", null);
        }
        const r = this.container(() => this.parseUntil({ env: name, item: true }));
        item.blocks = r.blocks;
        reason = r.reason;
        block.items.push(item);
      }
      if (reason === "eof" || reason === "docend") this.error(t.l, `\\begin{${name}} on line ${t.l} is never closed (missing \\end{${name}})`);
    } finally {
      this.listDepth--;
      if (isEnum) this.enumDepth--;
    }
    this.addBlock(block);
  }

  bibliography(t) {
    this.readArg("thebibliography");
    const block = { type: "bib", title: this.hasChapters ? "Bibliography" : "References", items: [], line: t.l };
    const pre = this.container(() => this.parseUntil({ env: "thebibliography", item: true }));
    let reason = pre.reason;
    let n = 0;
    for (let guard = 0; guard < 1000000 && (reason === "bibitem" || reason === "item"); guard++) {
      const line = this.line;
      const label = this.readOpt();
      const key = reason === "bibitem" ? this.argText("bibitem") : "";
      n++;
      const num = label ? plainText(label).trim() : String(n);
      if (key) this.bibMap.set(key, num);
      const r = this.container(() => this.parseUntil({ env: "thebibliography", item: true }));
      const inl = [];
      for (const b of r.blocks) if (b.type === "par") { if (inl.length) inl.push({ t: "text", v: " ", st: BASE_STYLE }); inl.push(...b.inl); }
      block.items.push({ num, key, inl, line, id: key ? `ltx-bib-${key.replace(/[^\w-]/g, "_")}` : null });
      reason = r.reason;
    }
    if (reason === "eof" || reason === "docend") this.error(t.l, `\\begin{thebibliography} on line ${t.l} is never closed`);
    this.addBlock(block);
  }

  theorem(name, def, t) {
    const note = this.readOpt();
    let num = null;
    if (!def.star) {
      const key = def.counter || name;
      const within = def.within ? this.sectionNumber(def.within) : "";
      let c = this.thmCounters.get(key);
      if (!c || c.within !== within) c = { n: 0, within };
      c.n++;
      this.thmCounters.set(key, c);
      num = within ? `${within}.${c.n}` : String(c.n);
    }
    const block = { type: "theorem", env: name, title: def.title, num, note: note ? this.captureInline(note) : null, style: def.style || "plain", blocks: [], line: t.l, ids: [] };
    this.setTarget(num || "", "theorem", block, def.title);
    const r = this.container(() => { this.style = BASE_STYLE; return this.body(name, t); });
    block.blocks = r.blocks;
    this.addBlock(block);
  }

  sectionNumber(level) {
    const order = ["part", "chapter", "section", "subsection", "subsubsection"];
    const idx = order.indexOf(level);
    if (idx < 0) return "";
    const parts = [];
    for (let k = this.hasChapters ? 1 : 2; k <= idx; k++) parts.push(this.counters[order[k]]);
    return parts.join(".");
  }

  float(name, t) {
    const kind = name.replace(/\*$/, "").replace(/^(wrap|sideways|SC|margin)/, "").replace(/^figure$/, "figure").toLowerCase();
    let width = null;
    if (name === "wrapfigure" || name === "wraptable") { this.readOpt(); this.readArg(name); this.readOpt(); width = widthPercent(plainText(this.readArg(name))); }
    else if (name === "subfigure" || name === "subtable") { this.readOpt(); width = widthPercent(plainText(this.readArg(name))); }
    else this.readOpt();
    const block = { type: "float", kind: kind === "listing" ? "listing" : kind, blocks: [], caption: null, num: null, ids: [], line: t.l, width };
    block.parent = this.curFloat;
    const saved = this.curFloat;
    this.curFloat = block;
    try {
      const r = this.container(() => { const res = this.body(name, t); block.align = this.align; return res; });
      block.blocks = r.blocks;
    } finally {
      this.curFloat = saved;
    }
    delete block.parent;
    this.addBlock(block);
  }

  frame(t) {
    this.skipOverlay();
    this.readOpt();
    let title = null, subtitle = null;
    if (this.peekIsBrace()) {
      title = this.captureInline(this.readArg("frame"));
      if (this.peekIsBrace()) subtitle = this.captureInline(this.readArg("frame"));
    }
    this.counters.frame++;
    const block = { type: "frame", title, subtitle, blocks: [], line: t.l, num: this.counters.frame };
    const saved = this.curFrame;
    this.curFrame = block;
    try {
      const r = this.container(() => this.body("frame", t));
      block.blocks = r.blocks;
    } finally {
      this.curFrame = saved;
    }
    this.addBlock(block);
    const plain = block.title ? inlinePlain(block.title).trim() : "";
    const isTitle = !plain && block.blocks.some((x) => x.type === "title");
    this.outline.push({ level: 3, title: plain || (isTitle ? "Title page" : `Slide ${block.num}`), line: t.l, num: "", frame: true });
  }

  userEnv(name, t) {
    const d = this.envDefs.get(name);
    const args = [];
    let first = 0;
    if (d.def !== null && d.def !== undefined) { const o = this.readOpt(); args.push(o === null ? d.defToks : o); first = 1; }
    for (let k = first; k < d.n; k++) args.push(this.readArg(name));
    const subst = (toks) => {
      const out = [];
      for (const bt of toks) {
        if (bt.k === "param") { const a = args[Number(bt.v) - 1]; if (a) out.push(...a); } else out.push({ ...bt, l: t.l });
      }
      return out;
    };
    if (++this.expansions > MAX_EXPANSIONS) return;
    this.parseTokens(subst(d.begin));
    this.body(name, t);
    this.parseTokens(subst(d.end));
  }

  tabular(name, t) {
    if (name === "tabular*" || name === "tabularx" || name === "tabulary" || name === "xtabular") this.readArg(name);
    if (name === "tblr" || name === "longtblr") this.readOpt();
    this.readOpt();
    const specToks = this.readArg(name);
    const spec = parseColSpec(serialize(specToks));
    const ctx = { env: name, spec, rows: [], caption: null, allowPar: spec.widths.some((w) => w) };
    const savedCtx = this.tableCtx;
    this.tableCtx = ctx;
    const block = { type: "table", rows: [], spec, line: t.l, env: name };
    try {
      let row = { cells: [], top: 0, bottom: 0, clines: [], line: this.line };
      let col = 0;
      for (let guard = 0; guard < 2000000; guard++) {
        const cellLine = this.line;
        const cell = { blocks: [], colspan: 1, rowspan: 1, align: spec.aligns[col] || "l", line: cellLine };
        // leading rules and \multicolumn / \multirow
        this.skipSpaces();
        for (let g2 = 0; g2 < 50; g2++) {
          const p = this.peek();
          if (!p) break;
          if (p.k === "txt" && !p.v.trim()) { this.next(); continue; }
          if (p.k === "par") { this.next(); continue; }
          if (p.k === "cmd" && /^(hline|toprule|midrule|bottomrule|hdashline|Xhline|specialrule|cmidrule|cline|addlinespace|endhead|endfirsthead|endfoot|endlastfoot|rowcolor|noalign)$/.test(p.v)) {
            this.next();
            if (p.v === "cline" || p.v === "cmidrule") {
              this.readOpt();
              const rng = /(\d+)\s*-\s*(\d+)/.exec(plainText(this.readArg(p.v)));
              if (rng && !row.cells.length) row.clines.push([Number(rng[1]), Number(rng[2])]);
            } else if (p.v === "specialrule") { this.readArg(p.v); this.readArg(p.v); this.readArg(p.v); row.top++; }
            else if (p.v === "Xhline" || p.v === "noalign") { this.readArg(p.v); row.top++; }
            else if (p.v === "addlinespace") this.readOpt();
            else if (p.v === "rowcolor") { this.readOpt(); const c = resolveColor(plainText(this.readArg(p.v)), this.colors); if (c) row.bg = c; }
            else if (!/^end/.test(p.v)) { if (!row.cells.length) row.top++; }
            continue;
          }
          if (p.k === "cmd" && p.v === "caption" && name.startsWith("longtable")) { this.next(); this.caption(p); this.skipSpaces(); const q = this.peek(); if (q && q.k === "cmd" && q.v === "\\") { this.next(); this.readStar(); this.readOpt(); } continue; }
          if (p.k === "cmd" && p.v === "label" && name.startsWith("longtable")) { this.next(); const key = this.argText("label"); this.defineLabel(key, this.lastTarget ? this.lastTarget.num : "", "table", block.id || (block.id = this.newId("tab")), p.l); continue; }
          break;
        }
        const p = this.peek();
        if (!row.cells.length && p && p.l) row.line = p.l;
        let reason;
        if (p && p.k === "cmd" && p.v === "multicolumn") {
          this.next();
          const n = Number(plainText(this.readArg("multicolumn")).trim()) || 1;
          const mspec = parseColSpec(serialize(this.readArg("multicolumn")));
          const content = this.readArg("multicolumn");
          cell.colspan = Math.max(1, Math.min(100, n));
          cell.align = mspec.aligns[0] || "c";
          cell.lines = [mspec.lines[0] || 0, mspec.lines[1] || 0];
          const r = this.container(() => { this.parseTokens(content); return this.parseUntil({ cell: true, env: name }); });
          cell.blocks = r.blocks;
          reason = r.reason;
        } else {
          const r = this.container(() => this.parseUntil({ cell: true, env: name }));
          cell.blocks = r.blocks;
          reason = r.reason;
        }
        hoistMultirow(cell);
        row.cells.push(cell);
        col += cell.colspan;
        if (reason === "&") continue;
        if (reason === "\\\\") {
          ctx.rows.push(row);
          row = { cells: [], top: 0, bottom: 0, clines: [], line: this.line };
          col = 0;
          continue;
        }
        if (reason !== "end") this.error(t.l, `\\begin{${name}} on line ${t.l} is never closed (missing \\end{${name}})`);
        break;
      }
      // the last row: keep only if it has content; a trailing \hline becomes a bottom rule
      const hasContent = row.cells.length > 1 || (row.cells[0] && row.cells[0].blocks.length);
      if (hasContent) ctx.rows.push(row);
      else if (row.top && ctx.rows.length) ctx.rows[ctx.rows.length - 1].bottom += row.top;
    } finally {
      this.tableCtx = savedCtx;
    }
    // validate column counts
    const ncols = Math.max(spec.aligns.length, 1);
    for (const r of ctx.rows) {
      const used = r.cells.reduce((a, c) => a + c.colspan, 0);
      if (used > ncols && spec.aligns.length) { this.error(r.line, `Extra & in table row: ${used} cells but the column spec "${serialize(specToks)}" has ${ncols}`); break; }
    }
    block.rows = ctx.rows;
    block.ncols = Math.max(ncols, ...ctx.rows.map((r) => r.cells.reduce((a, c) => a + c.colspan, 0)));
    if (ctx.caption) block.caption = ctx.caption;
    applyRowspans(block);
    this.addBlock(block);
  }
}

/** \multirow{n}{width}{text} as the first thing in a cell -> rowspan */
function hoistMultirow(cell) {
  const b = cell.blocks[0];
  if (b && b.type === "par" && b.inl[0] && b.inl[0].t === "multirow") {
    cell.rowspan = b.inl[0].n;
    b.inl.shift();
    if (!b.inl.length) cell.blocks.shift();
  }
}
function applyRowspans(block) {
  block.rows.forEach((r, ri) => {
    let col = 0;
    for (const c of r.cells) {
      if (c.rowspan > 1) {
        for (let k = 1; k < c.rowspan && ri + k < block.rows.length; k++) {
          const other = block.rows[ri + k];
          let oc = 0;
          for (const x of other.cells) {
            if (oc === col && !x.blocks.length) { x.covered = true; break; }
            oc += x.colspan;
          }
        }
      }
      col += c.colspan;
    }
  });
}

function widthPercent(w) {
  const s = String(w || "").trim();
  let m = /^([\d.]*)\s*\\?(textwidth|linewidth|columnwidth|hsize|paperwidth)/.exec(s);
  if (m) return Math.max(5, Math.min(100, (m[1] === "" ? 1 : parseFloat(m[1])) * 100));
  m = /^([\d.]+)\s*(cm|mm|in|pt|em)/.exec(s);
  if (m) {
    const cm = parseFloat(m[1]) * ({ cm: 1, mm: 0.1, in: 2.54, pt: 0.0351, em: 0.35 }[m[2]]);
    return Math.max(5, Math.min(100, (cm / 15) * 100));
  }
  return null;
}

/** enumerate label templates: "(a)", "i.", "\alph*)", "\Roman*." ... */
function parseEnumLabel(tpl) {
  if (tpl == null) return null;
  const s = String(tpl).trim();
  if (!s) return null;
  let m = /^(.*?)\\(arabic|alph|Alph|roman|Roman)\*?(.*)$/.exec(s);
  if (m) return { pre: textModeString(m[1]), style: m[2], post: textModeString(m[3]) };
  m = /^([^aAiI1]*)([aAiI1])(.*)$/.exec(s);
  if (m) {
    const style = { a: "alph", A: "Alph", i: "roman", I: "Roman", 1: "arabic" }[m[2]];
    return { pre: textModeString(m[1]), style, post: textModeString(m[3]) };
  }
  return null;
}

function textLigatures(v) {
  if (!/[-`'!?]/.test(v)) return v;
  return v.replace(/---/g, "—").replace(/--/g, "–").replace(/``/g, "“").replace(/''/g, "”").replace(/!`/g, "¡").replace(/\?`/g, "¿").replace(/`/g, "‘");
}

/** Plain text of inline nodes (outline titles, alt text). */
function inlinePlain(inl, labels) {
  let s = "";
  for (const n of inl || []) {
    switch (n.t) {
      case "text": case "code": s += n.v; break;
      case "math": s += uniList(n.nodes || []); break;
      case "br": s += " "; break;
      case "ref": { const L = labels && labels.get(n.key); s += L ? L.num : "??"; break; }
      case "cite": s += "[" + n.keys.join(",") + "]"; break;
      case "img": s += `[Image: ${n.file}]`; break;
      default: break;
    }
  }
  return s;
}

// @@CMD@@
// ---------------------------------------------------------------------------
// command handlers (this = DocParser). Return a stop reason to end the
// current parseUntil(), or nothing.
// ---------------------------------------------------------------------------

function styled(patch) {
  return function (t) {
    const arg = this.readArg(t.v);
    const saved = this.style;
    this.style = withStyle(this.style, typeof patch === "function" ? patch(this.style) : patch);
    this.parseTokens(arg);
    this.style = saved;
  };
}
function declare(patch) {
  return function () { this.style = withStyle(this.style, typeof patch === "function" ? patch(this.style) : patch); };
}
const RESET_FONT = { b: undefined, i: undefined, sc: undefined, sf: undefined, tt: undefined };

function defCommand(t) {
  this.readStar();
  const nameToks = this.readArg(t.v);
  const n = this.readOpt();
  const def = this.readOpt();
  const body = this.readArg(t.v);
  this.defineMacro(nameToks, n ? plainText(n).trim() : 0, def, body, t.v === "providecommand", t.l);
}
function sectionCmd(t) { this.section(t); }
function pageBreak(t) { this.readOpt(); if (this.inBody) this.addBlock({ type: "pagebreak", line: t.l }); }
function refCmd(t) {
  this.readStar();
  const keys = this.argText(t.v).split(",").map((s) => s.trim()).filter(Boolean);
  if (!this.inBody) return;
  keys.forEach((key, k) => {
    if (k) this.addText(", ", t.l);
    const node = { t: "ref", key, kind: t.v, l: t.l };
    this.refNodes.push(node);
    this.addInline(node);
  });
}
function citeCmd(t) {
  this.readStar();
  const o1 = this.readOpt();
  const o2 = this.readOpt();
  const keys = this.argText(t.v).split(",").map((s) => s.trim()).filter(Boolean);
  if (!this.inBody) return;
  const noteToks = o2 || o1;
  const note = noteToks ? inlinePlain(this.captureInline(noteToks), this.labels).trim() : "";
  const node = { t: "cite", keys, note, kind: t.v, l: t.l };
  this.citeNodes.push(node);
  this.addInline(node);
}
function overlayArg(t) {
  this.skipOverlay();
  if (this.peekIsBrace()) {
    const a = this.readArg(t.v);
    if (t.v !== "invisible") this.parseTokens(a);
  }
}
function lastArgOf(nOpt, nMand) {
  return function (t) {
    for (let k = 0; k < nOpt; k++) this.readOpt();
    for (let k = 0; k < nMand - 1; k++) this.readArg(t.v);
    this.readOpt();
    this.parseTokens(this.readArg(t.v));
  };
}
function algoLine(prefix, indentDelta, args = 0, suffix = "") {
  return function (t) {
    this.flushPar();
    if (indentDelta < 0) this.algoIndent = Math.max(0, (this.algoIndent || 0) + indentDelta);
    const parts = [];
    for (let k = 0; k < args; k++) parts.push(this.readArg(t.v));
    this.paraIndent = this.algoIndent || 0;
    const saved = this.style;
    this.style = withStyle(this.style, { b: true });
    if (prefix) this.addText(prefix + (args ? " " : ""), t.l);
    this.style = saved;
    parts.forEach((a, k) => {
      if (k) this.addText(", ", t.l);
      this.addMathOrText(a, t.l);
    });
    if (suffix) { this.style = withStyle(this.style, { b: true }); this.addText(" " + suffix, t.l); this.style = saved; }
    if (indentDelta > 0) this.algoIndent = (this.algoIndent || 0) + indentDelta;
  };
}

DocParser.prototype.addMathOrText = function (toks, line) {
  // algorithmic conditions are usually math-heavy: render $...$ as usual, bare text as math when it has operators
  const s = serialize(toks);
  if (/\$|\\\(/.test(s) || !/[=<>\\^_]/.test(s)) this.parseTokens(toks);
  else this.addInline({ t: "math", ...this.mathNode(s, false, line), l: line });
};

const _flushPar = DocParser.prototype.flushPar;
DocParser.prototype.flushPar = function () {
  if (this.para && this.paraIndent) {
    const n = this.cur.length;
    _flushPar.call(this);
    if (this.cur.length > n) this.cur[this.cur.length - 1].indent = this.paraIndent;
    this.paraIndent = 0;
    return;
  }
  this.paraIndent = 0;
  _flushPar.call(this);
};

const CIRCLED = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳";

const CMD = {
  documentclass(t) {
    const o = this.readOpt();
    const cls = this.argText("documentclass");
    this.docclass = cls || "article";
    this.hasChapters = /^(report|book|memoir|scrbook|scrreprt|thesis|ociamthesis)$/.test(this.docclass);
    this.isBeamer = this.docclass === "beamer";
    this.isLetter = this.docclass === "letter" || this.docclass === "scrlttr2";
    const m = /(\d+)pt/.exec(o ? plainText(o) : "");
    if (m) this.baseSize = Number(m[1]);
    this.classLine = t.l;
  },
  begin(t, mode) { return this.beginEnv(t, mode); },
  end(t, mode) {
    const name = this.argText("end");
    if (mode.env === name) return "end";
    let open = -1;
    for (let k = this.frames.length - 1; k >= 0; k--) if (this.frames[k].type === "env" && this.frames[k].name === name) { open = k; break; }
    if (open >= 0) {
      const f = this.frames[this.frames.length - 1];
      if (f.type === "group") this.error(f.line, `Missing } for the { on line ${f.line} (found \\end{${name}} on line ${t.l} first)`);
      else if (f.type === "env") this.error(f.line, `\\begin{${f.name}} on line ${f.line} is ended by \\end{${name}} on line ${t.l}`);
      this.pushTokens([t, { k: "ch", v: "{", l: t.l }, { k: "txt", v: name, l: t.l }, { k: "ch", v: "}", l: t.l }]);
      return "unwind";
    }
    const top = this.frames[this.frames.length - 1];
    if (mode.env && top && top.type === "env" && top.name === mode.env && mode.env !== "document") {
      // probably a typo (\begin{itemize} ... \end{itemise}): close the open environment
      this.error(t.l, `\\end{${name}} on line ${t.l} does not match \\begin{${mode.env}} on line ${top.line}`);
      return "end";
    }
    this.error(t.l, `\\end{${name}} on line ${t.l} has no matching \\begin{${name}}`);
    return null;
  },
  item(t, mode) {
    if (mode.item) return "item";
    this.error(t.l, "\\item outside a list (use it inside itemize or enumerate)");
    this.readOpt();
    this.addText("• ", t.l);
    return null;
  },
  bibitem(t, mode) {
    if (mode.item) return "bibitem";
    this.error(t.l, "\\bibitem outside thebibliography");
    this.readOpt();
    this.readArg("bibitem");
    return null;
  },
  par() { this.flushPar(); },
  newline(t, mode) { void mode; if (this.para) this.addInline({ t: "br" }); },
  linebreak(t) { this.readOpt(); if (this.para) this.addInline({ t: "br" }); },
  tabularnewline(t, mode) { if (mode.cell) return "\\\\"; return null; },
  cr(t, mode) { if (mode.cell) return "\\\\"; if (this.para) this.addInline({ t: "br" }); return null; },

  // ---- front matter
  title(t) { this.readOpt(); this.title = this.captureInline(this.readArg("title"), true); },
  author(t) { this.readOpt(); this.author = this.captureInline(this.readArg("author"), true); },
  date(t) { this.readOpt(); this.date = this.captureInline(this.readArg("date"), true); },
  subtitle(t) { this.readOpt(); this.subtitle = this.captureInline(this.readArg("subtitle"), true); },
  institute(t) { this.readOpt(); this.institute = this.captureInline(this.readArg("institute"), true); },
  address(t) { this.address = this.captureInline(this.readArg("address"), true); },
  signature(t) { this.signature = this.captureInline(this.readArg("signature"), true); },
  maketitle(t) { if (this.inBody) this.addBlock({ type: "title", line: t.l }); },
  titlepage(t) { if (this.inBody) this.addBlock({ type: "title", line: t.l }); },
  today(t) { this.addText(todayString(), t.l); },
  and(t) { this.addInline({ t: "and", l: t.l }); },
  thanks(t) { CMD.footnote.call(this, t); },

  // ---- definitions
  newcommand: defCommand, renewcommand: defCommand, providecommand: defCommand, DeclareRobustCommand: defCommand,
  newcommandx: defCommand, renewcommandx: defCommand,
  def(t) {
    const nt = this.next();
    let n = 0;
    for (let guard = 0; guard < 50; guard++) {
      const p = this.peek();
      if (!p || (p.k === "ch" && p.v === "{")) break;
      this.next();
      if (p.k === "param") n = Math.max(n, Number(p.v) || 0);
    }
    const body = this.readArg("def");
    if (nt && nt.k === "cmd") this.macros.set(nt.v, { n, def: null, body, src: serialize(body) });
  },
  gdef(t) { CMD.def.call(this, t); }, edef(t) { CMD.def.call(this, t); }, xdef(t) { CMD.def.call(this, t); },
  let(t) {
    const a = this.next();
    let b = this.next();
    if (b && b.k === "txt" && b.v.trim() === "=") b = this.next();
    else if (b && b.k === "txt" && /^\s*=?\s*$/.test(b.v)) b = this.next();
    if (a && a.k === "cmd" && b && b.k === "cmd") {
      if (this.macros.has(b.v)) this.macros.set(a.v, this.macros.get(b.v));
      else this.macros.set(a.v, { alias: b.v, n: 0, body: [{ k: "cmd", v: b.v }], src: "\\" + b.v + " " });
    }
  },
  newenvironment(t) {
    this.readStar();
    const name = this.argText(t.v);
    const n = this.readOpt();
    const def = this.readOpt();
    const begin = this.readArg(t.v);
    const end = this.readArg(t.v);
    if (name) this.envDefs.set(name, { n: Math.min(9, Number(n ? plainText(n) : 0) || 0), def: def ? serialize(def) : null, defToks: def, begin, end });
  },
  renewenvironment(t) { CMD.newenvironment.call(this, t); },
  newtheorem(t) {
    const star = this.readStar();
    const name = this.argText("newtheorem");
    const shared = this.readOpt();
    const title = plainText(this.readArg("newtheorem")).trim();
    const within = this.readOpt();
    if (name) this.theorems.set(name, { title: textModeString(title), counter: shared ? plainText(shared).trim() : name, within: within ? plainText(within).trim() : null, star, style: this.thmStyle });
  },
  theoremstyle(t) { this.thmStyle = this.argText(t.v) || "plain"; },
  definecolor(t) {
    this.readOpt();
    const name = this.argText(t.v), model = this.argText(t.v), spec = this.argText(t.v);
    const c = colorFromModel(model, spec);
    if (name && c) this.colors.set(name, c);
    else if (name) this.warn(t.l, `Color "${name}" could not be understood (model ${model})`);
  },
  colorlet(t) {
    this.readOpt();
    const name = this.argText(t.v), expr = this.argText(t.v);
    const c = resolveColor(expr, this.colors);
    if (name && c) this.colors.set(name, c);
  },
  DeclareMathOperator(t) {
    const star = this.readStar();
    const nameToks = this.readArg(t.v);
    const text = this.readArg(t.v);
    const nt = nameToks.find((x) => x.k === "cmd");
    if (!nt) return;
    const src = `\\operatorname${star ? "*" : ""}{${serialize(text)}}`;
    this.macros.set(nt.v, { n: 0, def: null, body: tokenize(src, t.l), src });
  },
  setcounter(t) {
    const name = this.argText(t.v);
    const v = parseInt(this.argText(t.v), 10);
    if (Object.prototype.hasOwnProperty.call(this.counters, name) && Number.isFinite(v)) this.counters[name] = v;
  },
  addtocounter(t) {
    const name = this.argText(t.v);
    const v = parseInt(this.argText(t.v), 10);
    if (Object.prototype.hasOwnProperty.call(this.counters, name) && Number.isFinite(v)) this.counters[name] += v;
  },
  stepcounter(t) { const name = this.argText(t.v); if (Object.prototype.hasOwnProperty.call(this.counters, name)) this.counters[name]++; },
  refstepcounter(t) { CMD.stepcounter.call(this, t); },
  appendix() {
    this.appendix = true;
    if (this.hasChapters) { this.counters.chapter = 0; this.appendixLevel = "chapter"; } else { this.counters.section = 0; this.appendixLevel = "section"; }
  },

  // ---- sectioning
  part: sectionCmd, chapter: sectionCmd, section: sectionCmd, subsection: sectionCmd, subsubsection: sectionCmd,
  paragraph: sectionCmd, subparagraph: sectionCmd,
  addchap(t) { t = { ...t, v: "chapter" }; this.unread({ k: "ch", v: "*", l: t.l }); this.section(t); },
  addsec(t) { t = { ...t, v: "section" }; this.unread({ k: "ch", v: "*", l: t.l }); this.section(t); },
  frametitle(t) {
    this.skipOverlay();
    this.readOpt();
    const inl = this.captureInline(this.readArg(t.v));
    if (this.curFrame) this.curFrame.title = inl;
  },
  framesubtitle(t) {
    this.skipOverlay();
    const inl = this.captureInline(this.readArg(t.v));
    if (this.curFrame) this.curFrame.subtitle = inl;
  },

  // ---- text styles
  textbf: styled({ b: true }), textit: styled({ i: true }), textsl: styled({ i: true }), emph: styled((s) => ({ i: !s.i })),
  underline: styled({ u: true }), uline: styled({ u: true }), uuline: styled({ u: true }), uwave: styled({ u: true }),
  sout: styled({ s: true }), st: styled({ s: true }), xout: styled({ s: true }), strike: styled({ s: true }),
  texttt: styled({ tt: true, sf: undefined }), textsc: styled({ sc: true }), textsf: styled({ sf: true, tt: undefined }),
  textrm: styled({ sf: undefined, tt: undefined }), textup: styled({ i: undefined }), textmd: styled({ b: undefined }),
  textnormal: styled(RESET_FONT), textsuperscript: styled({ sup: true }), textsubscript: styled({ sub: true }),
  bfseries: declare({ b: true }), bf: declare({ b: true }), itshape: declare({ i: true }), it: declare({ i: true }),
  slshape: declare({ i: true }), sl: declare({ i: true }), em: declare((s) => ({ i: !s.i })), upshape: declare({ i: undefined }),
  mdseries: declare({ b: undefined }), ttfamily: declare({ tt: true, sf: undefined }), tt: declare({ tt: true, sf: undefined }),
  sffamily: declare({ sf: true, tt: undefined }), sf: declare({ sf: true, tt: undefined }), rmfamily: declare({ sf: undefined, tt: undefined }),
  rm: declare({ sf: undefined, tt: undefined }), scshape: declare({ sc: true }), sc: declare({ sc: true }), normalfont: declare(RESET_FONT),
  centering() { this.align = "center"; },
  raggedright() { this.align = "left"; },
  raggedleft() { this.align = "right"; },
  noindent() { this.noindent = true; },
  textcolor(t) {
    const model = this.readOpt();
    const spec = this.argText(t.v);
    const color = model ? colorFromModel(plainText(model).trim(), spec) : resolveColor(spec, this.colors);
    if (!color) this.warn(t.l, `Unknown color "${spec}"`);
    const arg = this.readArg(t.v);
    const saved = this.style;
    this.style = withStyle(this.style, { color: color || undefined });
    this.parseTokens(arg);
    this.style = saved;
  },
  color(t) {
    const model = this.readOpt();
    const spec = this.argText(t.v);
    const color = model ? colorFromModel(plainText(model).trim(), spec) : resolveColor(spec, this.colors);
    if (!color) this.warn(t.l, `Unknown color "${spec}"`);
    this.style = withStyle(this.style, { color: color || undefined });
    this.skipSpaces();
  },
  colorbox(t) {
    const model = this.readOpt();
    const spec = this.argText(t.v);
    const color = model ? colorFromModel(plainText(model).trim(), spec) : resolveColor(spec, this.colors);
    const arg = this.readArg(t.v);
    const saved = this.style;
    this.style = withStyle(this.style, { hl: color || undefined });
    this.parseTokens(arg);
    this.style = saved;
  },
  fcolorbox(t) {
    this.readOpt();
    this.argText(t.v);
    const model = this.readOpt();
    const spec = this.argText(t.v);
    const color = model ? colorFromModel(plainText(model).trim(), spec) : resolveColor(spec, this.colors);
    const arg = this.readArg(t.v);
    const saved = this.style;
    this.style = withStyle(this.style, { hl: color || undefined, box: true });
    this.parseTokens(arg);
    this.style = saved;
  },
  hl: styled({ hl: "#FFFF00" }), highlight: styled({ hl: "#FFFF00" }), hlc(t) { this.readOpt(); styled({ hl: "#FFFF00" }).call(this, t); },
  fbox: styled({ box: true }),
  framebox(t) { this.readOpt(); this.readOpt(); styled({ box: true }).call(this, t); },
  mbox(t) { this.parseTokens(this.readArg(t.v)); },
  hbox(t) { this.parseTokens(this.readArg(t.v)); },
  text(t) { this.parseTokens(this.readArg(t.v)); },
  makebox(t) { this.readOpt(); this.readOpt(); this.parseTokens(this.readArg(t.v)); },
  raisebox(t) { this.readArg(t.v); this.readOpt(); this.readOpt(); this.parseTokens(this.readArg(t.v)); },
  scalebox: lastArgOf(0, 2), resizebox: lastArgOf(0, 3), rotatebox: lastArgOf(1, 2), reflectbox: lastArgOf(0, 1),
  parbox(t) { this.readOpt(); this.readOpt(); this.readOpt(); this.readArg(t.v); this.parseTokens(this.readArg(t.v)); },
  hyperlink(t) { this.readArg(t.v); this.parseTokens(this.readArg(t.v)); },
  hypertarget(t) { this.readArg(t.v); this.parseTokens(this.readArg(t.v)); },
  hyperref(t) { this.readOpt(); this.parseTokens(this.readArg(t.v)); },
  centerline(t) {
    const arg = this.readArg(t.v);
    this.flushPar();
    const saved = this.align;
    this.align = "center";
    this.parseTokens(arg);
    this.flushPar();
    this.align = saved;
  },
  url(t) {
    const url = plainText(this.readArg(t.v)).trim();
    const saved = this.style;
    this.style = withStyle(this.style, { link: url });
    this.addText(url, t.l);
    this.style = saved;
  },
  nolinkurl(t) { this.addText(plainText(this.readArg(t.v)).trim(), t.l); },
  href(t) {
    const url = plainText(this.readArg(t.v)).trim();
    const arg = this.readArg(t.v);
    const saved = this.style;
    this.style = withStyle(this.style, { link: url });
    this.parseTokens(arg);
    this.style = saved;
  },
  footnote(t) {
    const o = this.readOpt();
    const arg = this.readArg(t.v);
    if (!this.inBody) return;
    const n = o ? (parseInt(plainText(o), 10) || ++this.counters.footnote) : ++this.counters.footnote;
    const inl = this.captureInline(arg, true);
    this.footnotes.push({ n, inl, line: t.l });
    this.addInline({ t: "fn", n, l: t.l });
  },
  footnotemark(t) {
    const o = this.readOpt();
    if (!this.inBody) return;
    const n = o ? (parseInt(plainText(o), 10) || ++this.counters.footnote) : ++this.counters.footnote;
    this.addInline({ t: "fn", n, l: t.l });
  },
  footnotetext(t) {
    const o = this.readOpt();
    const arg = this.readArg(t.v);
    if (!this.inBody) return;
    const n = o ? (parseInt(plainText(o), 10) || this.counters.footnote) : this.counters.footnote;
    this.footnotes.push({ n, inl: this.captureInline(arg, true), line: t.l });
  },
  label(t) {
    const key = this.argText(t.v);
    if (!this.inBody || !key) return;
    const tg = this.lastTarget;
    let id = null;
    if (tg && tg.block) {
      tg.block.ids = tg.block.ids || [];
      if (!tg.block.ids.length) tg.block.ids.push(this.newId("lbl"));
      id = tg.block.ids[0];
    }
    this.defineLabel(key, tg ? tg.num : "", tg ? tg.kind : "", id, t.l, tg ? tg.title : "");
  },
  ref: refCmd, eqref: refCmd, pageref: refCmd, autoref: refCmd, Autoref: refCmd, cref: refCmd, Cref: refCmd,
  nameref: refCmd, vref: refCmd, cpageref: refCmd, labelcref: refCmd, subref: refCmd, namecref: refCmd,
  cite: citeCmd, citep: citeCmd, citet: citeCmd, parencite: citeCmd, textcite: citeCmd, autocite: citeCmd,
  footcite: citeCmd, Cite: citeCmd, Citep: citeCmd, Citet: citeCmd, citealp: citeCmd, citealt: citeCmd,
  supercite: citeCmd, smartcite: citeCmd, citeauthor: citeCmd, citeyear: citeCmd,
  bibliography(t) {
    const file = this.argText(t.v);
    if (!this.inBody) return;
    this.usesBibFile = true;
    this.warn(t.l, `The bibliography file "${file}.bib" cannot be read in the preview: citations show as [?] (Compile PDF shows them)`);
    this.addBlock({ type: "placeholder", text: `Bibliography from ${file}.bib (shown in the compiled PDF only)`, line: t.l });
  },
  printbibliography(t) {
    this.readOpt();
    if (!this.inBody) return;
    this.usesBibFile = true;
    this.warn(t.l, "biblatex bibliographies cannot be built in the preview: citations show as [?] (Compile PDF shows them)");
    this.addBlock({ type: "placeholder", text: "Bibliography (shown in the compiled PDF only)", line: t.l });
  },
  includegraphics(t) {
    this.readStar();
    const o = this.readOpt();
    this.readOpt();
    const file = this.argText(t.v);
    if (!this.inBody) return;
    const opts = o ? plainText(o) : "";
    let width = null;
    const wm = /width\s*=\s*([^,]+)/.exec(opts);
    if (wm) width = widthPercent(wm[1].trim());
    const sm = /scale\s*=\s*([\d.]+)/.exec(opts);
    if (!width && sm) width = Math.min(100, parseFloat(sm[1]) * 60);
    this.imageCount++;
    this.addInline({ t: "img", file, width, l: t.l });
  },
  includepdf(t) { this.readOpt(); const f = this.argText(t.v); this.addBlock({ type: "placeholder", text: `PDF pages from ${f} (shown in the compiled PDF only)`, line: t.l }); },
  input(t) {
    const file = this.argText(t.v);
    if (!this.inBody && this.hasDocEnv) { this.warn(t.l, `"${file}" is loaded from another file, which the preview cannot read`); return; }
    this.warn(t.l, `Included file "${file}" cannot be shown in the preview (only this file is rendered)`);
    this.addBlock({ type: "placeholder", text: `Contents of ${file} (only this file is shown in the preview)`, line: t.l });
  },
  include(t) { CMD.input.call(this, t); }, subfile(t) { CMD.input.call(this, t); },
  import(t) { this.readArg(t.v); CMD.input.call(this, t); },
  tableofcontents(t) { if (this.inBody) this.addBlock({ type: "toc", line: t.l }); },
  newpage: pageBreak, clearpage: pageBreak, cleardoublepage: pageBreak, pagebreak: pageBreak,
  smallskip(t) { this.addBlock({ type: "vspace", size: 0.5, line: t.l }); },
  medskip(t) { this.addBlock({ type: "vspace", size: 1, line: t.l }); },
  bigskip(t) { this.addBlock({ type: "vspace", size: 1.6, line: t.l }); },
  hspace(t) {
    this.readStar();
    const em = dimToEm(plainText(this.readArg(t.v)));
    if (em > 0.05) this.addText(em >= 0.9 ? " ".repeat(Math.min(12, Math.round(em))) : " ", t.l);
  },
  vspace(t) { this.readStar(); const em = dimToEm(plainText(this.readArg(t.v))); if (em > 0.4 && this.inBody && !this.para) this.addBlock({ type: "vspace", size: Math.min(6, em), line: t.l }); },
  rule(t) {
    this.readOpt();
    const w = this.argText(t.v);
    this.argText(t.v);
    const p = widthPercent(w);
    if (p && p >= 50) this.addBlock({ type: "hr", line: t.l });
  },
  hrule(t) { this.addBlock({ type: "hr", line: t.l }); },
  hfill(t) { if (this.para) this.addInline({ t: "hfill", l: t.l }); else this.noindent = true; },
  hrulefill(t) { if (this.para) this.addInline({ t: "fill", l: t.l }); else this.addBlock({ type: "hr", line: t.l }); },
  lipsum(t) {
    const o = this.readOpt();
    let a = 1, b = 7;
    if (o) { const m = /(\d+)(?:\s*-\s*(\d+))?/.exec(plainText(o)); if (m) { a = Number(m[1]); b = m[2] ? Number(m[2]) : a; } }
    for (let k = a; k <= Math.min(b, a + 20); k++) { this.flushPar(); this.addText(LOREM[(k - 1) % LOREM.length], t.l); this.flushPar(); }
  },
  blindtext(t) { this.readOpt(); this.flushPar(); this.addText(LOREM[0] + " " + LOREM[1], t.l); this.flushPar(); },
  Blindtext(t) { this.readOpt(); for (let k = 0; k < 4; k++) { this.flushPar(); this.addText(LOREM[k], t.l); } this.flushPar(); },
  multicolumn(t) { this.readArg(t.v); this.readArg(t.v); this.parseTokens(this.readArg(t.v)); },
  multirow(t) {
    this.readOpt();
    const n = parseInt(this.argText(t.v), 10) || 1;
    this.readOpt();
    this.readArg(t.v);
    this.readOpt();
    const content = this.readArg(t.v);
    if (this.tableCtx && !this.para) this.addInline({ t: "multirow", n: Math.min(100, Math.max(1, n)), l: t.l });
    this.parseTokens(content);
  },
  ensuremath(t) {
    const tex = serialize(this.readArg(t.v));
    if (this.inBody) this.addInline({ t: "math", ...this.mathNode(tex, false, t.l), l: t.l });
  },
  marginpar(t) {
    this.readOpt();
    const arg = this.readArg(t.v);
    const saved = this.style;
    this.style = withStyle(this.style, { size: 8, color: "#666666" });
    this.addText(" [", t.l);
    this.parseTokens(arg);
    this.addText("] ", t.l);
    this.style = saved;
  },
  todo(t) { this.readOpt(); const arg = this.readArg(t.v); const saved = this.style; this.style = withStyle(this.style, { hl: "#FFF3B0" }); this.addText("TODO: ", t.l); this.parseTokens(arg); this.style = saved; },
  alert(t) { this.skipOverlay(); styled({ color: "#CC0000" }).call(this, t); },
  structure(t) { this.skipOverlay(); styled({ color: "#33339A" }).call(this, t); },
  only: overlayArg, uncover: overlayArg, visible: overlayArg, onslide: overlayArg, invisible: overlayArg,
  alt(t) { this.skipOverlay(); const a = this.readArg(t.v); this.readArg(t.v); this.parseTokens(a); },
  temporal(t) { this.skipOverlay(); this.readArg(t.v); const a = this.readArg(t.v); this.readArg(t.v); this.parseTokens(a); },
  opening(t) {
    const arg = this.readArg(t.v);
    if (!this.inBody) return;
    this.addBlock({ type: "letterhead", line: t.l });
    this.parseTokens(arg);
    this.flushPar();
  },
  closing(t) {
    const arg = this.readArg(t.v);
    if (!this.inBody) return;
    this.flushPar();
    this.parseTokens(arg);
    this.flushPar();
    this.addBlock({ type: "vspace", size: 2.5, line: t.l });
    if (this.signature) { this.para = [...this.signature]; this.paraLine = t.l; this.flushPar(); }
  },
  ps(t) { const a = this.readArg(t.v); this.flushPar(); this.parseTokens(a); this.flushPar(); },
  cc(t) { const a = this.readArg(t.v); this.flushPar(); this.addText("cc: ", t.l); this.parseTokens(a); this.flushPar(); },
  encl(t) { const a = this.readArg(t.v); this.flushPar(); this.addText("Enclosure: ", t.l); this.parseTokens(a); this.flushPar(); },
  textcircled(t) {
    const s = plainText(this.readArg(t.v)).trim();
    const n = parseInt(s, 10);
    this.addText(n >= 1 && n <= 20 ? CIRCLED[n - 1] : `(${s})`, t.l);
  },
  caption(t) { this.caption(t); },
  captionof(t) {
    this.readStar();
    const kind = this.argText(t.v);
    this.readOpt();
    const arg = this.readArg(t.v);
    if (!this.inBody) return;
    const key = kind === "table" ? "table" : "figure";
    this.counters[key]++;
    const num = this.hasChapters ? `${this.chapterNum()}.${this.counters[key]}` : String(this.counters[key]);
    const inl = this.captureInline(arg);
    this.flushPar();
    const b = { type: "par", inl: [{ t: "text", v: `${key === "table" ? "Table" : "Figure"} ${num}: `, st: withStyle(BASE_STYLE, { b: true }) }, ...inl], line: t.l, align: "center", ids: [] };
    this.cur.push(b);
    this.setTarget(num, key, b);
  },
  phantom(t) { this.readArg(t.v); }, hphantom(t) { this.readArg(t.v); }, vphantom(t) { this.readArg(t.v); },
  ding(t) { this.readArg(t.v); this.addText("✓", t.l); },
  // algorithmic (algpseudocode)
  State: algoLine("", 0), Statex: algoLine("", 0),
  If: algoLine("if", 1, 1, "then"), ElsIf: algoLine("else if", 0, 1, "then"), Else: algoLine("else", 0), EndIf: algoLine("end if", -1),
  For: algoLine("for", 1, 1, "do"), ForAll: algoLine("for all", 1, 1, "do"), EndFor: algoLine("end for", -1),
  While: algoLine("while", 1, 1, "do"), EndWhile: algoLine("end while", -1), Repeat: algoLine("repeat", 1), Until: algoLine("until", -1, 1),
  Loop: algoLine("loop", 1), EndLoop: algoLine("end loop", -1), Return: algoLine("return", 0), Require: algoLine("Require:", 0), Ensure: algoLine("Ensure:", 0),
  Procedure: algoLine("procedure", 1, 2), EndProcedure: algoLine("end procedure", -1), Function: algoLine("function", 1, 2), EndFunction: algoLine("end function", -1),
  Comment(t) { const a = this.readArg(t.v); this.addText("  ▷ ", t.l); this.parseTokens(a); },
  Call(t) { const n = plainText(this.readArg(t.v)); const a = this.readArg(t.v); const saved = this.style; this.style = withStyle(this.style, { sc: true }); this.addText(n, t.l); this.style = saved; this.addText("(", t.l); this.addMathOrText(a, t.l); this.addText(")", t.l); },
};
CMD.newline = function (t, mode) { void mode; if (this.para) this.addInline({ t: "br" }); };

DocParser.prototype.run = function () {
  try {
    this.parseUntil({ top: true });
  } catch (e) {
    this.error(this.line, `The preview stopped early: ${e && e.message ? e.message : e}`);
  }
  this.flushPar();
  this.finish();
  return this;
};

DocParser.prototype.finish = function () {
  if (this.classLine && !this.hasDocEnv) this.error(this.classLine, "Missing \\begin{document} (the document body must be inside \\begin{document} ... \\end{document})");
  if (this.unknownCmds.size) {
    const list = [...this.unknownCmds].slice(0, 25).map(([n, l]) => `\\${n} (line ${l})`).join(", ");
    const first = Math.min(...this.unknownCmds.values());
    this.warn(first, `Unknown command${this.unknownCmds.size === 1 ? "" : "s"} (text kept, formatting dropped): ${list}${this.unknownCmds.size > 25 ? ", …" : ""}`);
  }
  const seenRef = new Set();
  for (const r of this.refNodes) {
    if (!this.labels.has(r.key) && !seenRef.has(r.key)) { seenRef.add(r.key); this.warn(r.l, `Reference "${r.key}" is not defined (shown as ??); add \\label{${r.key}}`); }
  }
  if (!this.usesBibFile) {
    const seen = new Set();
    for (const c of this.citeNodes) for (const k of c.keys) {
      if (!this.bibMap.has(k) && !seen.has(k)) { seen.add(k); this.warn(c.l, `Citation "${k}" is not defined (shown as [?]); add \\bibitem{${k}} to thebibliography`); }
    }
  }
  this.errors.sort((a, b) => a.line - b.line || (a.severity === b.severity ? 0 : a.severity === "error" ? -1 : 1));
};

// @@RENDER@@
// ===========================================================================
// HTML renderer
// ===========================================================================

function safeUrl(u) {
  const s = String(u || "").trim();
  if (/^(javascript|data|vbscript|file|chrome|chrome-extension):/i.test(s)) return "#";
  if (/^www\./i.test(s)) return "https://" + s;
  if (/^[\w.+-]+@[\w-]+\.[\w.-]+$/.test(s)) return "mailto:" + s;
  return s;
}

function styleCss(st, base) {
  const css = [];
  if (st.b) css.push("font-weight:bold");
  if (st.i) css.push("font-style:italic");
  const deco = [st.u && "underline", st.s && "line-through"].filter(Boolean);
  if (deco.length) css.push(`text-decoration:${deco.join(" ")}`);
  if (st.sc) css.push("font-variant:small-caps");
  if (st.tt) css.push("font-family:var(--ltx-mono)");
  else if (st.sf) css.push("font-family:var(--ltx-sans)");
  if (st.size) css.push(`font-size:${+(st.size / base).toFixed(3)}em`);
  if (st.color) css.push(`color:${st.color}`);
  if (st.hl) css.push(`background-color:${st.hl}`);
  if (st.box) css.push("border:1px solid currentColor;padding:0 3px");
  return css.join(";");
}
function wrapStyle(html, st, base) {
  let out = html;
  const css = styleCss(st, base);
  if (css) out = `<span style="${css}">${out}</span>`;
  if (st.sup) out = `<sup>${out}</sup>`;
  else if (st.sub) out = `<sub>${out}</sub>`;
  if (st.link) out = `<a class="ltx-link" href="${esc(safeUrl(st.link))}" target="_blank" rel="noopener noreferrer">${out}</a>`;
  return out;
}

const IMG_ICON = '<svg class="ltx-img-icon" viewBox="0 0 24 24" width="28" height="28" aria-hidden="true"><rect x="3" y="4" width="18" height="16" rx="2" fill="none" stroke="currentColor" stroke-width="1.6"/><circle cx="9" cy="9.5" r="1.8" fill="currentColor"/><path d="M4.5 18l5-5.5 3.5 3.5 2.5-2.5 4 4.5" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linejoin="round"/></svg>';
const ITEM_MARKS = ["•", "–", "∗", "·"];
const REF_NAMES = { section: "Section", subsection: "Section", subsubsection: "Section", chapter: "Chapter", part: "Part", paragraph: "Paragraph", subparagraph: "Paragraph", figure: "Figure", subfigure: "Figure", table: "Table", subtable: "Table", equation: "Equation", item: "Item", algorithm: "Algorithm", listing: "Listing" };

class HtmlRenderer {
  constructor(P) {
    this.P = P;
    this.base = P.baseSize || 10;
  }

  refText(n) {
    const L = this.P.labels.get(n.key);
    if (!L) return { text: "??", id: null, missing: true };
    const num = L.num || "";
    switch (n.kind) {
      case "eqref": return { text: `(${num})`, id: L.id };
      case "pageref": case "cpageref": return { text: "?", id: L.id, title: "Page numbers are only known in the compiled PDF" };
      case "nameref": return { text: L.title || num, id: L.id };
      case "autoref": case "Autoref": case "cref": case "Cref": case "vref": case "namecref": {
        const kindName = L.kind === "theorem" ? (L.title || "Theorem") : REF_NAMES[L.kind] || "";
        if (n.kind === "namecref") return { text: kindName, id: L.id };
        const s = kindName ? `${kindName} ${L.kind === "equation" && n.kind !== "autoref" && n.kind !== "Autoref" ? `(${num})` : num}` : num;
        return { text: n.kind === "cref" || n.kind === "vref" ? s.charAt(0).toLowerCase() + s.slice(1) : s, id: L.id };
      }
      default: return { text: num, id: L.id };
    }
  }

  citeText(n) {
    const nums = n.keys.map((k) => this.P.bibMap.get(k) ?? "?");
    if (n.kind === "citeauthor") return n.keys.join(", ");
    return `[${nums.join(", ")}${n.note ? `, ${n.note}` : ""}]`;
  }

  node(n) {
    switch (n.t) {
      case "text": return esc(n.v);
      case "math": return n.xml;
      case "br": return "<br>";
      case "fn": return `<sup class="ltx-fnref"><a href="#ltx-fn-${n.n}" data-fn="${n.n}">${n.n}</a></sup>`;
      case "ref": {
        const r = this.refText(n);
        if (r.missing) return `<span class="ltx-ref ltx-missing" title="Undefined reference: ${esc(n.key)}">??</span>`;
        return r.id ? `<a class="ltx-ref" href="#${r.id}"${r.title ? ` title="${esc(r.title)}"` : ""}>${esc(r.text)}</a>` : `<span class="ltx-ref">${esc(r.text)}</span>`;
      }
      case "cite": {
        const t = this.citeText(n);
        const k = n.keys.find((x) => this.P.bibMap.has(x));
        return k ? `<a class="ltx-cite" href="#ltx-bib-${esc(k.replace(/[^\w-]/g, "_"))}">${esc(t)}</a>` : `<span class="ltx-cite${t.includes("?") ? " ltx-missing" : ""}">${esc(t)}</span>`;
      }
      case "code": return `<code class="ltx-verb">${esc(n.v)}</code>`;
      case "img": return `<span class="ltx-img" style="width:${n.width ? Math.round(n.width) : 50}%" title="${esc(n.file)}">${IMG_ICON}<span class="ltx-img-name">${esc(n.file || "image")}</span><span class="ltx-img-note">image (not available in the preview)</span></span>`;
      case "and": return '<span class="ltx-and"></span>';
      case "fill": return '<span class="ltx-fill"></span>';
      case "hfill": return '<span class="ltx-hfill"></span>';
      default: return "";
    }
  }

  inl(nodes) {
    let out = "";
    let k = 0;
    while (k < nodes.length) {
      const st = nodes[k].st || BASE_STYLE;
      const key = styleKey(st);
      let html = "";
      while (k < nodes.length && styleKey(nodes[k].st || BASE_STYLE) === key) html += this.node(nodes[k++]);
      out += key === "{}" ? html : wrapStyle(html, st, this.base);
    }
    return out;
  }

  ids(b) { return b.ids && b.ids.length ? ` id="${b.ids[0]}"` : ""; }
  anchors(b) { return b.ids && b.ids.length > 1 ? b.ids.slice(1).map((id) => `<a id="${id}"></a>`).join("") : ""; }

  blocks(list) {
    let out = "";
    for (const b of list) out += this.block(b);
    return out;
  }

  par(b, prefix = "", suffix = "") {
    const cls = b.noindent ? ' class="ltx-noindent"' : "";
    const css = [];
    if (b.align) css.push(`text-align:${b.align}`);
    if (b.indent) css.push(`margin-left:${b.indent * 1.5}em`);
    let body;
    if (b.inl.some((n) => n.t === "hfill")) {
      // \hfill: every line (split at \\) becomes a flex row where \hfill takes the free space
      const lines = [[]];
      for (const n of b.inl) { if (n.t === "br") lines.push([]); else lines[lines.length - 1].push(n); }
      body = lines.map((ln, k) => `<span class="ltx-fline">${k === 0 ? prefix : ""}${this.inl(ln)}</span>`).join("");
      return `<p data-line="${b.line}"${this.ids(b)} class="ltx-noindent ltx-flex"${css.length ? ` style="${css.join(";")}"` : ""}>${body}${suffix}</p>`;
    }
    return `<p data-line="${b.line}"${this.ids(b)}${cls}${css.length ? ` style="${css.join(";")}"` : ""}>${prefix}${this.inl(b.inl)}${suffix}</p>`;
  }

  block(b) {
    switch (b.type) {
      case "par": return this.par(b);
      case "heading": return this.heading(b);
      case "title": return this.titleBlock(b);
      case "list": return this.list(b);
      case "quote": return `<blockquote class="ltx-quote ltx-${b.env}" data-line="${b.line}">${this.blocks(b.blocks)}</blockquote>`;
      case "abstract": return `<div class="ltx-abstract" data-line="${b.line}"><div class="ltx-abstract-title">Abstract</div>${this.blocks(b.blocks)}</div>`;
      case "div": return this.div(b);
      case "columns": return `<div class="ltx-columns" data-line="${b.line}">${this.blocks(b.blocks)}</div>`;
      case "proof": return this.proof(b);
      case "theorem": return this.theorem(b);
      case "code": return `<figure class="ltx-codeblock" data-line="${b.line}">${b.caption ? `<figcaption>${esc(b.caption)}</figcaption>` : ""}<pre class="ltx-code"${b.lang ? ` data-lang="${esc(b.lang)}"` : ""}><code>${esc(b.text)}</code></pre></figure>`;
      case "math": return this.math(b);
      case "table": return this.table(b);
      case "float": return this.float(b);
      case "frame": return this.frame(b);
      case "bib": return this.bib(b);
      case "toc": return this.toc(b);
      case "placeholder": return `<div class="ltx-placeholder" data-line="${b.line}">${esc(b.text)}</div>`;
      case "pagebreak": return `<div class="ltx-pagebreak" data-line="${b.line}" title="Page break"></div>`;
      case "hr": return `<hr class="ltx-hr" data-line="${b.line}">`;
      case "vspace": return `<div class="ltx-vspace" data-line="${b.line}" style="height:${b.size}em"></div>`;
      case "letter": return this.letter(b);
      case "letterhead": return "";
      default: return "";
    }
  }

  heading(b) {
    const title = this.inl(b.title);
    const line = b.line;
    if (b.beamer) return `<div class="ltx-beamer-section" data-line="${line}"${this.ids(b)}>${esc(b.cmd === "subsection" ? "Subsection" : "Section")}: ${title}</div>`;
    if (b.cmd === "part") return `<h1 class="ltx-part" data-line="${line}"${this.ids(b)}>${b.num ? `<span class="ltx-partnum">Part ${esc(b.num)}</span>` : ""}${title}</h1>${this.anchors(b)}`;
    if (b.cmd === "chapter") {
      const label = this.P.appendix && /^[A-Z]+$/.test(b.num || "") ? "Appendix" : "Chapter";
      return `<h1 class="ltx-chapter" data-line="${line}"${this.ids(b)}>${b.num ? `<span class="ltx-chapnum">${label} ${esc(b.num)}</span>` : ""}${title}</h1>${this.anchors(b)}`;
    }
    const tag = { section: "h2", subsection: "h3", subsubsection: "h4", paragraph: "h5", subparagraph: "h6" }[b.cmd] || "h2";
    return `<${tag} class="ltx-${b.cmd}" data-line="${line}"${this.ids(b)}>${b.num ? `<span class="ltx-secnum">${esc(b.num)}</span>` : ""}${title}</${tag}>${this.anchors(b)}`;
  }

  titleBlock(b) {
    const P = this.P;
    const parts = [];
    if (P.title) parts.push(`<div class="ltx-title">${this.inl(P.title)}</div>`);
    else parts.push('<div class="ltx-title ltx-missing" title="Add \\title{...} to the preamble">(no title)</div>');
    if (P.subtitle) parts.push(`<div class="ltx-subtitle">${this.inl(P.subtitle)}</div>`);
    if (P.author && P.author.length) parts.push(`<div class="ltx-author">${this.inl(P.author)}</div>`);
    if (P.institute && P.institute.length) parts.push(`<div class="ltx-institute">${this.inl(P.institute)}</div>`);
    const date = P.date === null ? esc(todayString()) : this.inl(P.date);
    if (date) parts.push(`<div class="ltx-date">${date}</div>`);
    const inner = `<header class="ltx-titleblock">${parts.join("")}</header>`;
    if (P.isBeamer && !this.inFrame) return `<section class="ltx-frame ltx-titleframe" data-line="${b.line}"><div class="ltx-frame-body">${inner}</div></section>`;
    if (P.isBeamer) return inner.replace("<header ", `<header data-line="${b.line}" `).replace('class="ltx-titleblock"', 'class="ltx-titleblock ltx-intitleframe"');
    return inner.replace("<header ", `<header data-line="${b.line}" `);
  }

  list(b) {
    if (b.kind === "dl") {
      let out = `<div class="ltx-desc" data-line="${b.line}">`;
      for (const it of b.items) {
        const label = it.label ? `<span class="ltx-dt">${this.inl(it.label)}</span> ` : "";
        out += `<div class="ltx-desc-item" data-line="${it.line}">${this.itemBody(it, label)}</div>`;
      }
      return out + "</div>";
    }
    const tag = b.kind === "ol" ? "ol" : "ul";
    let out = `<${tag} class="ltx-list ltx-${tag}" data-line="${b.line}">`;
    for (const it of b.items) {
      let mark;
      if (it.label) mark = this.inl(it.label);
      else if (b.kind === "ol") mark = esc(it.mark || "");
      else mark = ITEM_MARKS[Math.min(3, (b.depth || 1) - 1)];
      out += `<li data-line="${it.line}"><span class="ltx-mark">${mark}</span><div class="ltx-item-body">${this.blocks(it.blocks)}</div></li>`;
    }
    return out + `</${tag}>`;
  }

  /** Block list whose first paragraph gets an inline prefix (run-in heads). */
  itemBody(it, prefix, suffix = "") {
    const blocks = it.blocks || [];
    if (blocks.length && blocks[0].type === "par") {
      const rest = blocks.slice(1);
      if (!rest.length) return this.par({ ...blocks[0], noindent: true }, prefix, suffix);
      const last = rest[rest.length - 1];
      if (suffix && last.type === "par") return this.par({ ...blocks[0], noindent: true }, prefix) + this.blocks(rest.slice(0, -1)) + this.par(last, "", suffix);
      return this.par({ ...blocks[0], noindent: true }, prefix) + this.blocks(rest) + (suffix ? `<p class="ltx-noindent ltx-qedline">${suffix}</p>` : "");
    }
    return (prefix ? `<p class="ltx-noindent">${prefix}</p>` : "") + this.blocks(blocks) + (suffix ? `<p class="ltx-noindent ltx-qedline">${suffix}</p>` : "");
  }

  theorem(b) {
    const head = `<span class="ltx-thm-head">${esc(b.title)}${b.num ? ` ${esc(b.num)}` : ""}</span>${b.note ? ` <span class="ltx-thm-note">(${this.inl(b.note)})</span>` : ""}<span class="ltx-thm-head">.</span> `;
    return `<div class="ltx-thm ltx-thm-${esc(b.style)}" data-line="${b.line}"${this.ids(b)}>${this.itemBody(b, head)}</div>`;
  }

  proof(b) {
    const head = `<span class="ltx-proof-head">${b.note ? this.inl(b.note) : "Proof"}.</span> `;
    return `<div class="ltx-proof" data-line="${b.line}">${this.itemBody(b, head, '<span class="ltx-qed">∎</span>')}</div>`;
  }

  div(b) {
    const inner = this.blocks(b.blocks);
    if (b.box) {
      return `<div class="ltx-block ltx-${b.box}" data-line="${b.line}">${b.boxTitle && b.boxTitle.length ? `<div class="ltx-block-title">${this.inl(b.boxTitle)}</div>` : ""}<div class="ltx-block-body">${inner}</div></div>`;
    }
    if (b.mini) return `<div class="ltx-minipage${b.column ? " ltx-column" : ""}" data-line="${b.line}" style="width:${b.width ? Math.round(b.width) : 48}%${b.align ? `;text-align:${b.align}` : ""}">${inner}</div>`;
    if (b.cols) return `<div class="ltx-multicols" data-line="${b.line}" style="column-count:${b.cols}">${inner}</div>`;
    if (b.algo) return `<div class="ltx-algo" data-line="${b.line}">${inner}</div>`;
    return `<div class="ltx-div${b.align ? ` ltx-${b.align}` : ""}" data-line="${b.line}"${b.align ? ` style="text-align:${b.align}"` : ""}>${inner}</div>`;
  }

  math(b) {
    const blockXml = (c) => c.xml;
    const inlineDisp = (c) => c.xml.replace('<math display="block"', '<math displaystyle="true"');
    const num = (r) => (r.num ? `<span class="ltx-eqno">${esc(r.num)}</span>` : "");
    const ids = (r) => (r.ids && r.ids.length ? ` id="${r.ids[0]}"` : "");
    if (["align", "flalign", "alignat", "eqnarray"].includes(b.env)) {
      const ncells = Math.max(1, ...b.rows.map((r) => r.cells.length));
      let out = `<div class="ltx-eqwrap" data-line="${b.line}"><table class="ltx-align"><tbody>`;
      for (const r of b.rows) {
        let cells = "";
        for (let k = 0; k < ncells; k++) {
          const c = r.cells[k];
          const cls = b.env === "eqnarray" ? ["r", "c", "l"][k % 3] : k % 2 === 0 ? "r" : "l";
          cells += `<td class="ltx-al-${cls}">${c ? inlineDisp(c) : ""}</td>`;
        }
        out += `<tr${ids(r)}><td class="ltx-eqpad"></td>${cells}<td class="ltx-eqpad ltx-eqnocell">${num(r)}</td></tr>`;
      }
      return out + "</tbody></table></div>";
    }
    if (b.rows.length > 1) {
      let out = `<div class="ltx-eqgroup" data-line="${b.line}">`;
      b.rows.forEach((r, k) => {
        const pos = b.env === "multline" ? (k === 0 ? " ltx-ml-first" : k === b.rows.length - 1 ? " ltx-ml-last" : "") : "";
        out += `<div class="ltx-eq${pos}"${ids(r)}><div class="ltx-eq-body">${blockXml(r.cells[0])}</div>${num(r)}</div>`;
      });
      return out + "</div>";
    }
    const r = b.rows[0] || { cells: [{ xml: "" }] };
    return `<div class="ltx-eq" data-line="${b.line}"${ids(r)}><div class="ltx-eq-body">${blockXml(r.cells[0])}</div>${num(r)}</div>`;
  }

  table(b) {
    const spec = b.spec || { aligns: [], widths: [], lines: [] };
    const border = (n) => (n >= 2 ? "3px double currentColor" : "1px solid currentColor");
    let rows = "";
    b.rows.forEach((r, ri) => {
      let col = 0;
      let cells = "";
      for (const c of r.cells) {
        const span = c.colspan || 1;
        if (c.covered) { col += span; continue; }
        const css = [`text-align:${{ l: "left", c: "center", r: "right" }[c.align] || "left"}`];
        const leftLine = c.lines ? c.lines[0] : col === 0 ? spec.lines[0] : 0;
        const rightLine = c.lines ? c.lines[1] : spec.lines[col + span];
        if (leftLine) css.push(`border-left:${border(leftLine)}`);
        if (rightLine) css.push(`border-right:${border(rightLine)}`);
        if (r.top) css.push(`border-top:${border(r.top)}`);
        else if (r.clines && r.clines.some(([a, z]) => col + 1 >= a && col + 1 <= z)) css.push("border-top:1px solid currentColor");
        if (r.bottom || (ri === b.rows.length - 1 && r.bottom)) css.push(`border-bottom:${border(r.bottom)}`);
        const w = spec.widths[col];
        if (w && w !== "X" && span === 1) { const pw = widthPercent(w); if (pw) css.push(`width:${Math.round(pw)}%`); }
        const content = c.blocks.length === 1 && c.blocks[0].type === "par" ? this.inl(c.blocks[0].inl) : this.blocks(c.blocks);
        cells += `<td${span > 1 ? ` colspan="${span}"` : ""}${c.rowspan > 1 ? ` rowspan="${c.rowspan}"` : ""} style="${css.join(";")}">${content}</td>`;
        col += span;
      }
      rows += `<tr data-line="${r.line}"${r.bg ? ` style="background:${r.bg}"` : ""}>${cells}</tr>`;
    });
    const cap = b.caption ? `<caption><span class="ltx-caplabel">Table ${esc(b.caption.num)}:</span> ${this.inl(b.caption.inl)}</caption>` : "";
    return `<table class="ltx-tabular" data-line="${b.line}"${b.id ? ` id="${b.id}"` : ""}>${cap}<tbody>${rows}</tbody></table>`;
  }

  float(b) {
    const name = { figure: "Figure", table: "Table", algorithm: "Algorithm", listing: "Listing", subfigure: "", subtable: "" }[b.kind] ?? "Figure";
    let cap = "";
    if (b.caption) {
      const label = name ? `${name} ${b.num}:` : `${b.num}`;
      cap = `<figcaption><span class="ltx-caplabel">${esc(label)}</span> ${this.inl(b.caption)}</figcaption>`;
    }
    const sub = b.kind === "subfigure" || b.kind === "subtable";
    const css = [];
    if (b.align) css.push(`text-align:${b.align}`);
    if (sub || b.width) css.push(`width:${Math.round(b.width || 45)}%`);
    return `<figure class="ltx-float ltx-${esc(b.kind)}${sub ? " ltx-sub" : ""}" data-line="${b.line}"${this.ids(b)}${css.length ? ` style="${css.join(";")}"` : ""}>${b.captionTop ? cap : ""}${this.blocks(b.blocks)}${b.captionTop ? "" : cap}</figure>`;
  }

  frame(b) {
    const title = b.title && b.title.length ? `<div class="ltx-frame-title">${this.inl(b.title)}${b.subtitle ? `<div class="ltx-frame-subtitle">${this.inl(b.subtitle)}</div>` : ""}</div>` : "";
    const saved = this.inFrame;
    this.inFrame = true;
    const body = this.blocks(b.blocks);
    this.inFrame = saved;
    const isTitle = !title && b.blocks.some((x) => x.type === "title");
    return `<section class="ltx-frame${isTitle ? " ltx-titleframe" : ""}" data-line="${b.line}">${title}<div class="ltx-frame-body">${body}</div><div class="ltx-frame-num">${b.num}</div></section>`;
  }

  bib(b) {
    let out = `<section class="ltx-bib" data-line="${b.line}"><h2 class="ltx-section">${esc(b.title)}</h2><div class="ltx-biblist">`;
    for (const it of b.items) out += `<p class="ltx-bibitem" data-line="${it.line}"${it.id ? ` id="${it.id}"` : ""}><span class="ltx-bibnum">[${esc(it.num)}]</span> ${this.inl(it.inl)}</p>`;
    return out + "</div></section>";
  }

  toc(b) {
    const items = this.P.outline.filter((o) => !o.frame && o.level <= (this.P.hasChapters ? 3 : 4));
    let out = `<nav class="ltx-toc" data-line="${b.line}"><h2 class="ltx-toc-title">Contents</h2>`;
    for (const o of items) out += `<div class="ltx-toc-item ltx-toc-l${o.level}"><a href="#${o.id || ""}" data-line="${o.line}">${o.num ? `<span class="ltx-toc-num">${esc(o.num)}</span>` : ""}${esc(o.title)}</a></div>`;
    return out + "</nav>";
  }

  letter(b) {
    const P = this.P;
    const from = P.address ? `<div class="ltx-letter-from">${this.inl(P.address)}</div>` : "";
    const date = `<div class="ltx-letter-date">${P.date === null ? esc(todayString()) : this.inl(P.date)}</div>`;
    return `<div class="ltx-letter" data-line="${b.line}">${from}${date}<div class="ltx-letter-to">${this.inl(b.to)}</div>${this.blocks(b.blocks)}</div>`;
  }

  footnotes() {
    const fns = this.P.footnotes;
    if (!fns.length) return "";
    let out = '<section class="ltx-footnotes"><div class="ltx-footnotes-rule"></div>';
    for (const f of fns) out += `<p class="ltx-footnote" id="ltx-fn-${f.n}" data-line="${f.line}"><sup>${f.n}</sup> ${this.inl(f.inl)}</p>`;
    return out + "</section>";
  }

  render() {
    const P = this.P;
    const cls = `ltx-doc ltx-class-${esc(String(P.docclass).replace(/[^\w-]/g, ""))}${P.isBeamer ? " ltx-beamer" : ""}${P.hasChapters ? " ltx-chapters" : ""}`;
    return `<article class="${cls}">${this.blocks(P.root)}${this.footnotes()}</article>`;
  }
}

/** Parse + render LaTeX for the live preview. Never throws. */
export function latexToHtml(source) {
  try {
    const P = new DocParser(source).run();
    for (const o of P.outline) if (!o.id) o.id = "";
    const html = new HtmlRenderer(P).render();
    return { html, outline: P.outline.map((o) => ({ level: o.level, title: o.title, line: o.line, num: o.num, ...(o.frame ? { frame: true } : {}) })), errors: P.errors };
  } catch (e) {
    return {
      html: `<article class="ltx-doc"><pre class="ltx-code">${esc(source)}</pre></article>`,
      outline: [],
      errors: [{ line: 1, message: `The preview could not be built (${e && e.message ? e.message : e})`, severity: "error" }],
    };
  }
}

export const PREVIEW_CSS = `
.ltx-doc{--ltx-serif:"Latin Modern Roman","CMU Serif","Cambria","Georgia","Times New Roman",serif;--ltx-sans:"Segoe UI","Helvetica Neue",Arial,sans-serif;--ltx-mono:"Cascadia Mono",Consolas,"Courier New",monospace;
 font-family:var(--ltx-serif);font-size:16px;line-height:1.5;color:#1b1b1b;background:#fff;max-width:760px;margin:0 auto;padding:56px 64px 72px;box-sizing:border-box;text-align:justify;hyphens:auto;overflow-wrap:break-word}
.ltx-doc *{box-sizing:border-box}
.ltx-doc p{margin:0 0 .25em;text-indent:1.5em}
.ltx-doc p.ltx-noindent,.ltx-doc h1+p,.ltx-doc h2+p,.ltx-doc h3+p,.ltx-doc h4+p,.ltx-doc header+p,.ltx-doc .ltx-item-body>p:first-child,.ltx-doc td p,.ltx-doc figure p,.ltx-doc .ltx-abstract p:first-of-type,.ltx-doc .ltx-frame-body p,.ltx-doc .ltx-block-body p,.ltx-doc .ltx-div>p:first-child,.ltx-doc blockquote>p:first-child,.ltx-doc .ltx-minipage p,.ltx-doc .ltx-letter p{text-indent:0}
.ltx-doc .ltx-letter p{margin-bottom:.8em}
.ltx-doc h1,.ltx-doc h2,.ltx-doc h3,.ltx-doc h4,.ltx-doc h5,.ltx-doc h6{font-family:var(--ltx-serif);text-align:left;line-height:1.25;color:#111}
.ltx-doc h1.ltx-part{font-size:2.1em;text-align:center;margin:1.6em 0 1em}
.ltx-doc .ltx-partnum{display:block;font-size:.75em;margin-bottom:.3em}
.ltx-doc h1.ltx-chapter{font-size:1.9em;margin:1.4em 0 .9em}
.ltx-doc .ltx-chapnum{display:block;font-size:.7em;margin-bottom:.45em}
.ltx-doc h2.ltx-section{font-size:1.42em;margin:1.4em 0 .55em}
.ltx-doc h3.ltx-subsection{font-size:1.18em;margin:1.2em 0 .45em}
.ltx-doc h4.ltx-subsubsection{font-size:1.02em;margin:1em 0 .4em}
.ltx-doc h5.ltx-paragraph,.ltx-doc h6.ltx-subparagraph{font-size:1em;margin:.9em 0 .25em}
.ltx-doc h6.ltx-subparagraph{margin-left:1.5em}
.ltx-doc .ltx-secnum{margin-right:.9em}
.ltx-doc .ltx-titleblock{text-align:center;margin:0 0 2.2em}
.ltx-doc .ltx-title{font-size:1.75em;line-height:1.25;margin-bottom:.7em}
.ltx-doc .ltx-subtitle{font-size:1.2em;margin:-.3em 0 .7em}
.ltx-doc .ltx-author,.ltx-doc .ltx-institute{font-size:1.1em;margin-bottom:.35em}
.ltx-doc .ltx-date{font-size:1.05em}
.ltx-doc .ltx-and{display:inline-block;width:2.5em}
.ltx-doc .ltx-missing{color:#b42318}
.ltx-doc .ltx-abstract{margin:1.2em 3em 1.6em;font-size:.92em}
.ltx-doc .ltx-abstract-title{font-weight:bold;text-align:center;margin-bottom:.4em}
.ltx-doc blockquote{margin:.8em 2.5em}
.ltx-doc .ltx-verse p{text-align:left}
.ltx-doc .ltx-list{list-style:none;margin:.5em 0;padding-left:2.2em}
.ltx-doc .ltx-list>li{position:relative;margin:.2em 0}
.ltx-doc .ltx-mark{position:absolute;right:100%;margin-right:.55em;white-space:nowrap;text-align:right}
.ltx-doc .ltx-desc{margin:.5em 0 .5em 1.5em}
.ltx-doc .ltx-desc-item>p:first-child{text-indent:-1.5em}
.ltx-doc .ltx-dt{font-weight:bold}
.ltx-doc .ltx-thm,.ltx-doc .ltx-proof{margin:.9em 0}
.ltx-doc .ltx-thm-plain .ltx-item-body,.ltx-doc .ltx-thm-plain p{font-style:italic}
.ltx-doc .ltx-thm-head{font-weight:bold;font-style:normal}
.ltx-doc .ltx-thm-remark .ltx-thm-head{font-weight:normal;font-style:italic}
.ltx-doc .ltx-thm-note{font-style:normal}
.ltx-doc .ltx-proof-head{font-style:italic}
.ltx-doc .ltx-qed{float:right;margin-left:1em}
.ltx-doc .ltx-qedline{text-align:right}.ltx-doc .ltx-qedline .ltx-qed{float:none}
.ltx-doc .ltx-eq{display:flex;align-items:center;margin:.7em 0;gap:1em}
.ltx-doc .ltx-eq-body{flex:1;min-width:0;overflow-x:auto;overflow-y:hidden;text-align:center}
.ltx-doc .ltx-eq-body math[display=block]{margin:0}
.ltx-doc .ltx-ml-first .ltx-eq-body{text-align:left}.ltx-doc .ltx-ml-first math{display:inline-math}
.ltx-doc .ltx-ml-last .ltx-eq-body{text-align:right}.ltx-doc .ltx-ml-last math{display:inline-math}
.ltx-doc .ltx-eqno{flex:none;white-space:nowrap}
.ltx-doc .ltx-eqwrap{overflow-x:auto;overflow-y:hidden;margin:.7em 0;padding:2px 0}
.ltx-doc table.ltx-align{width:100%;border-collapse:collapse}
.ltx-doc table.ltx-align td{padding:.25em 0;vertical-align:middle}
.ltx-doc .ltx-al-r{text-align:right}.ltx-doc .ltx-al-l{text-align:left;padding-right:1.5em!important}.ltx-doc .ltx-al-c{text-align:center;padding:0 .4em!important}
.ltx-doc .ltx-eqpad{width:50%}.ltx-doc .ltx-eqnocell{text-align:right;white-space:nowrap}
.ltx-doc math{font-size:1.06em}
.ltx-doc table.ltx-tabular{display:inline-table;border-collapse:collapse;margin:.6em 0;text-align:left;text-indent:0;font-size:.95em;vertical-align:top}
.ltx-doc table.ltx-tabular td{padding:.22em .6em;vertical-align:top}
.ltx-doc table.ltx-tabular caption{caption-side:top;padding-bottom:.4em;text-align:center}
.ltx-doc figure{margin:1.2em 0;text-indent:0}
.ltx-doc figure.ltx-sub{display:inline-block;vertical-align:top;margin:.4em 1%}
.ltx-doc figcaption{margin:.5em 0;font-size:.95em;text-align:center}
.ltx-doc .ltx-table figcaption{margin-top:0}
.ltx-doc .ltx-caplabel{font-weight:bold}
.ltx-doc .ltx-img{display:inline-flex;flex-direction:column;align-items:center;justify-content:center;gap:4px;min-height:120px;max-width:100%;vertical-align:middle;border:1.5px dashed #9aa4ae;border-radius:6px;background:#f4f6f8;color:#5b6570;font-family:var(--ltx-sans);font-size:13px;padding:10px;text-indent:0;text-align:center}
.ltx-doc .ltx-img-name{font-weight:600;word-break:break-all;color:#36404a}
.ltx-doc .ltx-img-note{font-size:11px}
.ltx-doc .ltx-placeholder{border:1.5px dashed #9aa4ae;border-radius:6px;background:#f4f6f8;color:#5b6570;font-family:var(--ltx-sans);font-size:13px;padding:18px;text-align:center;margin:.8em 0}
.ltx-doc pre.ltx-code{font-family:var(--ltx-mono);font-size:.82em;line-height:1.45;background:#f6f8fa;border:1px solid #e1e4e8;border-radius:6px;padding:.7em .9em;overflow-x:auto;text-align:left;white-space:pre;margin:0}
.ltx-doc .ltx-codeblock{margin:.8em 0}
.ltx-doc code.ltx-verb{font-family:var(--ltx-mono);font-size:.9em}
.ltx-doc a.ltx-link,.ltx-doc a.ltx-ref,.ltx-doc a.ltx-cite{color:#1a56db;text-decoration:none}
.ltx-doc a.ltx-link:hover,.ltx-doc a.ltx-ref:hover{text-decoration:underline}
.ltx-doc .ltx-fnref a{color:inherit;text-decoration:none}
.ltx-doc sup,.ltx-doc sub{line-height:0}
.ltx-doc .ltx-footnotes{margin-top:2em;font-size:.86em}
.ltx-doc .ltx-footnotes-rule{width:35%;border-top:1px solid #555;margin-bottom:.6em}
.ltx-doc .ltx-footnote{text-indent:0;margin:.15em 0}
.ltx-doc .ltx-pagebreak{border-top:1px dashed #b8c0c8;margin:1.6em -64px;position:relative}
.ltx-doc .ltx-pagebreak::after{content:"page break";position:absolute;left:50%;top:-.75em;transform:translateX(-50%);background:#fff;color:#8a949e;font:11px var(--ltx-sans);padding:0 8px}
.ltx-doc .ltx-hr{border:0;border-top:1px solid #444;margin:1em 0}
.ltx-doc .ltx-fill{display:inline-block;width:6em;border-bottom:1px solid currentColor}
.ltx-doc p.ltx-flex .ltx-fline{display:flex;flex-wrap:wrap;align-items:baseline;column-gap:.3em}.ltx-doc .ltx-hfill{flex:1 1 0;min-width:.5em}
.ltx-doc .ltx-bib .ltx-bibitem{padding-left:2.4em;text-indent:-2.4em;margin:.3em 0}
.ltx-doc .ltx-bibnum{display:inline-block;min-width:2.1em}
.ltx-doc .ltx-toc{margin:1em 0 2em}
.ltx-doc .ltx-toc-title{font-size:1.42em}
.ltx-doc .ltx-toc-item a{color:inherit;text-decoration:none;display:block;padding:1px 0}
.ltx-doc .ltx-toc-item a:hover{color:#1a56db}
.ltx-doc .ltx-toc-num{display:inline-block;min-width:2.6em}
.ltx-doc .ltx-toc-l1{font-weight:bold;margin-top:.5em}.ltx-doc .ltx-toc-l2{font-weight:bold;margin-top:.3em}
.ltx-doc.ltx-chapters .ltx-toc-l2{font-weight:normal;margin-left:1.5em;margin-top:0}
.ltx-doc .ltx-toc-l3{margin-left:1.5em}.ltx-doc.ltx-chapters .ltx-toc-l3{margin-left:3.6em}.ltx-doc .ltx-toc-l4{margin-left:3.6em}
.ltx-doc .ltx-minipage{display:inline-block;vertical-align:top;margin:.3em 1%;text-indent:0}
.ltx-doc .ltx-columns{display:flex;gap:4%;align-items:flex-start}
.ltx-doc .ltx-columns>.ltx-minipage{margin:0}
.ltx-doc .ltx-multicols{column-gap:2em}
.ltx-doc .ltx-algo{border-top:1.5px solid #333;border-bottom:1.5px solid #333;padding:.3em 0;text-align:left}
.ltx-doc .ltx-algo p{margin:.1em 0}
.ltx-doc .ltx-letter-from{text-align:right;margin-bottom:1em;white-space:normal}
.ltx-doc .ltx-letter-date{text-align:right;margin-bottom:1.5em}
.ltx-doc .ltx-letter-to{margin-bottom:1.5em}
.ltx-doc .ltx-vspace{display:block}
.ltx-doc.ltx-beamer{max-width:none;padding:24px;background:transparent;text-align:left}
.ltx-doc .ltx-frame{background:#fff;border:1px solid #cfd6dd;border-radius:6px;box-shadow:0 2px 8px rgba(0,0,0,.08);aspect-ratio:4/3;max-width:640px;margin:0 auto 24px;display:flex;flex-direction:column;overflow:hidden;position:relative;font-family:var(--ltx-sans);font-size:15px}
.ltx-doc .ltx-frame-title{background:#2b3a67;color:#fff;font-size:1.35em;padding:12px 22px}
.ltx-doc .ltx-frame-subtitle{font-size:.7em;opacity:.85;margin-top:2px}
.ltx-doc .ltx-frame-body{flex:1;padding:14px 26px;overflow:hidden;display:flex;flex-direction:column;justify-content:center}
.ltx-doc .ltx-frame-num{position:absolute;right:10px;bottom:6px;font-size:11px;color:#8a949e}
.ltx-doc .ltx-titleframe .ltx-titleblock{margin:0}
.ltx-doc .ltx-titleframe .ltx-title{color:#2b3a67}
.ltx-doc .ltx-beamer-section{max-width:640px;margin:8px auto;color:#6a737d;font:12px var(--ltx-sans);text-transform:uppercase;letter-spacing:.06em}
.ltx-doc .ltx-block{border-radius:5px;overflow:hidden;margin:.6em 0;background:#eef0f7}
.ltx-doc .ltx-block-title{background:#3e4f86;color:#fff;padding:4px 10px}
.ltx-doc .ltx-alertblock .ltx-block-title{background:#b42318}.ltx-doc .ltx-exampleblock .ltx-block-title{background:#18794e}
.ltx-doc .ltx-block-body{padding:6px 10px}
@media print{.ltx-doc{max-width:none;padding:0}.ltx-doc .ltx-pagebreak{break-after:page;border:0;margin:0}.ltx-doc .ltx-pagebreak::after{display:none}.ltx-doc .ltx-frame{break-inside:avoid;box-shadow:none}.ltx-doc .ltx-img,.ltx-doc .ltx-placeholder{break-inside:avoid}}
`;

/** Standalone HTML page (embedded CSS, MathML) for Save As .html. */
export function latexToHtmlDocument(source, title = "Document") {
  const r = latexToHtml(source);
  return `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${esc(title)}</title>
<style>
html{background:#eef0f3}body{margin:0;padding:24px 12px}
.ltx-doc{box-shadow:0 1px 4px rgba(0,0,0,.12)}
${PREVIEW_CSS}
</style>
</head>
<body>
${r.html}
</body>
</html>
`;
}

// ===========================================================================
// LaTeX -> word model
// ===========================================================================

class WordRenderer {
  constructor(P) {
    this.P = P;
    this.out = [];
    this.warnings = new Set();
    this.html = new HtmlRenderer(P);
  }

  runsOf(inl, extra = {}) {
    const runs = [];
    const push = (r) => {
      if (!r.t) return;
      const last = runs[runs.length - 1];
      if (last && sameRunProps(last, r)) last.t += r.t;
      else runs.push(r);
    };
    for (const n of inl || []) {
      const props = { ...runProps(n.st || BASE_STYLE, this.P.baseSize), ...extra };
      switch (n.t) {
        case "text": push({ t: n.v, ...props }); break;
        case "math": push({ t: uniList(n.nodes || []).replace(/\s+/g, " ").trim(), font: "Cambria Math", ...pick(props, ["color", "link", "b"]) }); break;
        case "br": push({ t: "\n", ...props }); break;
        case "fn": push({ t: `[${n.n}]`, sup: true }); break;
        case "ref": push({ t: this.html.refText(n).text, ...props }); break;
        case "cite": push({ t: this.html.citeText(n), ...props }); break;
        case "code": push({ t: n.v, ...props, font: "Courier New" }); break;
        case "img": this.imageWarning(); push({ t: `[Image: ${n.file}]`, ...props }); break;
        case "and": if (runs.length) runs[runs.length - 1].t = runs[runs.length - 1].t.replace(/\s+$/, ""); push({ t: ", ", ...props }); break;
        case "fill": case "hfill": push({ t: "\t", ...props }); break;
        default: break;
      }
    }
    // trim
    if (runs.length) {
      runs[0].t = runs[0].t.replace(/^[ \t]+/, "");
      runs[runs.length - 1].t = runs[runs.length - 1].t.replace(/\s+$/, "");
    }
    return runs.filter((r) => r.t);
  }

  imageWarning() { this.warnings.add("Images are not embedded: each \\includegraphics becomes the text [Image: file name]."); }

  para(type, runs, align) {
    if (!runs.length && type !== "p") return;
    const b = { type, runs };
    if (align && align !== "left" && type !== "code") b.align = align;
    this.out.push(b);
  }

  blocks(list, ctx) {
    for (const b of list) this.block(b, ctx);
  }

  block(b, ctx) {
    const P = this.P;
    const ptype = ctx.quote ? "quote" : "p";
    switch (b.type) {
      case "par": {
        const runs = this.runsOf(b.inl, ctx.italic ? { i: true } : {});
        if (ctx.prefix) { runs.unshift(...ctx.prefix); ctx.prefix = null; }
        if (runs.length) this.para(ptype, mergeRuns(runs), b.align || ctx.align);
        break;
      }
      case "heading": {
        let lvl;
        if (P.isBeamer) lvl = 2;
        else if (b.cmd === "part") lvl = 1;
        else if (P.hasChapters) lvl = Math.min(6, b.level);
        else lvl = Math.min(6, Math.max(1, b.level - 1));
        const runs = this.runsOf(b.title);
        if (runs.length) this.out.push({ type: `h${lvl}`, runs });
        break;
      }
      case "title": {
        if (P.title) { const r = this.runsOf(P.title); if (r.length) this.out.push({ type: "h1", align: "center", runs: r }); }
        if (P.subtitle) { const r = this.runsOf(P.subtitle); if (r.length) this.out.push({ type: "p", align: "center", runs: r }); }
        if (P.author) { const r = this.runsOf(P.author); if (r.length) this.out.push({ type: "p", align: "center", runs: r }); }
        if (P.institute) { const r = this.runsOf(P.institute); if (r.length) this.out.push({ type: "p", align: "center", runs: r }); }
        const d = P.date === null ? [{ t: todayString() }] : this.runsOf(P.date);
        if (d.length) this.out.push({ type: "p", align: "center", runs: d });
        break;
      }
      case "list": this.list(b, ctx); break;
      case "quote": this.blocks(b.blocks, { ...ctx, quote: true }); break;
      case "abstract":
        this.out.push({ type: "p", align: "center", runs: [{ t: "Abstract", b: true }] });
        this.blocks(b.blocks, ctx);
        break;
      case "div":
        if (b.box && b.boxTitle && b.boxTitle.length) this.para("p", this.runsOf(b.boxTitle, { b: true }));
        this.blocks(b.blocks, { ...ctx, align: b.align || ctx.align });
        break;
      case "columns": this.blocks(b.blocks, ctx); break;
      case "proof": {
        const prefix = [{ t: b.note ? `${inlinePlain(b.note, P.labels)}. ` : "Proof. ", i: true }];
        this.withPrefix(b.blocks, ctx, prefix, [{ t: " ∎" }]);
        break;
      }
      case "theorem": {
        const head = [{ t: `${b.title}${b.num ? ` ${b.num}` : ""}`, b: true }];
        if (b.note) head.push({ t: ` (${inlinePlain(b.note, P.labels)})` });
        head.push({ t: ". ", b: true });
        this.withPrefix(b.blocks, { ...ctx, italic: b.style === "plain" }, head, null);
        break;
      }
      case "code":
        for (const line of String(b.text).split("\n")) this.out.push({ type: "code", runs: line ? [{ t: line }] : [] });
        break;
      case "math":
        for (const r of b.rows) {
          const t = r.cells.map((c) => uniList(c.nodes || []).replace(/\s+/g, " ").trim()).join(" ").replace(/\s+/g, " ").trim();
          const runs = t ? [{ t, font: "Cambria Math" }] : [];
          if (r.num) runs.push({ t: `\t${r.num}` });
          if (runs.length) this.out.push({ type: "p", align: "center", runs });
        }
        break;
      case "table": this.table(b, ctx); break;
      case "float": {
        const cap = b.caption ? [{ t: `${{ table: "Table", algorithm: "Algorithm", listing: "Listing", subfigure: "", subtable: "" }[b.kind] ?? "Figure"} ${b.num}: `.trimStart(), b: true }, ...this.runsOf(b.caption)] : null;
        if (cap && b.captionTop) this.out.push({ type: "p", align: "center", runs: mergeRuns(cap) });
        this.blocks(b.blocks, { ...ctx, align: b.align || ctx.align });
        if (cap && !b.captionTop) this.out.push({ type: "p", align: "center", runs: mergeRuns(cap) });
        break;
      }
      case "frame": {
        if (this.out.length) this.out.push({ type: "pagebreak" });
        if (b.title && b.title.length) this.out.push({ type: "h2", runs: this.runsOf(b.title) });
        if (b.subtitle && b.subtitle.length) this.out.push({ type: "p", runs: this.runsOf(b.subtitle, { i: true }) });
        this.blocks(b.blocks, ctx);
        break;
      }
      case "bib":
        this.out.push({ type: P.hasChapters ? "h1" : "h1", runs: [{ t: b.title }] });
        for (const it of b.items) this.out.push({ type: "p", runs: mergeRuns([{ t: `[${it.num}] ` }, ...this.runsOf(it.inl)]) });
        break;
      case "placeholder": this.out.push({ type: "p", runs: [{ t: `[${b.text}]`, i: true }] }); break;
      case "pagebreak": this.out.push({ type: "pagebreak" }); break;
      case "letter": {
        if (P.address) this.out.push({ type: "p", align: "right", runs: this.runsOf(P.address) });
        this.out.push({ type: "p", align: "right", runs: P.date === null ? [{ t: todayString() }] : this.runsOf(P.date) });
        const to = this.runsOf(b.to);
        if (to.length) this.out.push({ type: "p", runs: to });
        this.blocks(b.blocks, ctx);
        break;
      }
      default: break;
    }
  }

  withPrefix(blocks, ctx, prefix, suffix) {
    const start = this.out.length;
    if (blocks.length && blocks[0].type === "par") {
      this.blocks(blocks, { ...ctx, prefix: prefix.slice() });
    } else {
      this.out.push({ type: ctx.quote ? "quote" : "p", runs: prefix.slice() });
      this.blocks(blocks, ctx);
    }
    if (suffix) {
      const last = this.out[this.out.length - 1];
      if (this.out.length > start && last && Array.isArray(last.runs) && last.type !== "code") last.runs.push(...suffix);
      else this.out.push({ type: "p", align: "right", runs: suffix });
    }
  }

  list(b, ctx) {
    const level = Math.min(8, ctx.level === undefined ? 0 : ctx.level);
    for (const it of b.items) {
      const blocks = it.blocks || [];
      let runs = [];
      let rest = blocks;
      if (blocks.length && blocks[0].type === "par") { runs = this.runsOf(blocks[0].inl); rest = blocks.slice(1); }
      if (it.label && it.label.length) {
        const lab = this.runsOf(it.label, b.kind === "dl" ? { b: true } : {});
        runs = mergeRuns([...lab, { t: " " }, ...runs]);
      }
      this.out.push({ type: "list", list: b.kind === "ol" ? "ol" : "ul", level, runs });
      for (const r of rest) {
        if (r.type === "list") this.list(r, { ...ctx, level: level + 1 });
        else this.block(r, ctx);
      }
    }
  }

  table(b, ctx) {
    const rows = [];
    for (const r of b.rows) {
      const cells = [];
      for (const c of r.cells) {
        const sub = new WordRenderer(this.P);
        sub.warnings = this.warnings;
        sub.blocks(c.blocks, { align: null });
        const cell = { blocks: sub.out.length ? sub.out : [{ type: "p", runs: [] }] };
        if (c.colspan > 1) cell.colspan = c.colspan;
        if (c.align === "c") cell.blocks.forEach((x) => { if (x.type === "p" && !x.align) x.align = "center"; });
        if (c.align === "r") cell.blocks.forEach((x) => { if (x.type === "p" && !x.align) x.align = "right"; });
        cells.push(cell);
      }
      if (cells.length) rows.push(cells);
    }
    if (b.caption) this.out.push({ type: "p", align: "center", runs: mergeRuns([{ t: `Table ${b.caption.num}: `, b: true }, ...this.runsOf(b.caption.inl)]) });
    if (rows.length) this.out.push({ type: "table", rows });
    void ctx;
  }

  render() {
    const P = this.P;
    this.blocks(P.root, { align: null });
    if (P.footnotes.length) {
      this.out.push({ type: "h1", runs: [{ t: "Notes" }] });
      for (const f of P.footnotes) this.out.push({ type: "p", runs: mergeRuns([{ t: `${f.n}. ` }, ...this.runsOf(f.inl)]) });
    }
    const meta = {};
    if (P.title) { const t = inlinePlain(P.title, P.labels).replace(/\s+/g, " ").trim(); if (t) meta.title = t; }
    if (P.author) { const a = this.runsOf(P.author).map((r) => r.t).join("").replace(/\s+/g, " ").trim(); if (a) meta.author = a; }
    for (const e of P.errors) if (e.severity === "warning" && /^Unknown command/.test(e.message)) this.warnings.add(e.message);
    if (this.warnings.size) meta.warnings = [...this.warnings];
    return { kind: "word", meta, blocks: this.out };
  }
}

function pick(o, keys) { const r = {}; for (const k of keys) if (o[k] !== undefined) r[k] = o[k]; return r; }
function runProps(st, base) {
  const r = {};
  if (st.b) r.b = true;
  if (st.i) r.i = true;
  if (st.u) r.u = true;
  if (st.s) r.s = true;
  if (st.sup) r.sup = true;
  if (st.sub) r.sub = true;
  if (st.color) r.color = st.color;
  if (st.hl) r.hl = st.hl;
  if (st.link) r.link = safeUrl(st.link);
  if (st.tt) r.font = "Courier New";
  if (st.size && Math.abs(st.size - (base || 10)) > 0.01) r.size = Math.round(st.size * 1.1 * 2) / 2;
  return r;
}
const RUN_KEYS = ["b", "i", "u", "s", "sup", "sub", "color", "hl", "link", "font", "size"];
function sameRunProps(a, b) { return RUN_KEYS.every((k) => a[k] === b[k]); }
function mergeRuns(runs) {
  const out = [];
  for (const r of runs) {
    if (!r || !r.t) continue;
    const last = out[out.length - 1];
    if (last && sameRunProps(last, r)) last.t += r.t;
    else out.push({ ...r });
  }
  return out;
}

/** LaTeX source -> word model (WORD CONTRACT). Never throws. */
export function latexToWordModel(source) {
  try {
    const P = new DocParser(source).run();
    return new WordRenderer(P).render();
  } catch (e) {
    return { kind: "word", meta: { warnings: [`The LaTeX could not be converted completely (${e && e.message ? e.message : e})`] }, blocks: String(source || "").split(/\n{2,}/).map((t) => ({ type: "p", runs: [{ t }] })) };
  }
}

// ===========================================================================
// word model -> LaTeX
// ===========================================================================

/** Unicode -> LaTeX math command (inverse of the symbol tables). */
const UNI_MATH = (() => {
  const m = new Map();
  const add = (tbl, filter) => { for (const [name, ch] of Object.entries(tbl)) { const c = Array.isArray(ch) ? ch[0] : ch; if (c && !m.has(c) && /^[A-Za-z]+$/.test(name) && (!filter || filter(name))) m.set(c, "\\" + name); } };
  add(GREEK_LOWER); add(GREEK_UPPER, (n) => !n.startsWith("var"));
  add(MATH_OPS, (n) => !["colon", "lbrack", "rbrack", "lparen", "rparen", "backslash", "dots", "hdots", "dotsc", "dotso", "prime"].includes(n));
  add(MATH_ORD, (n) => !["S", "P", "dag", "ddag", "copyright", "pounds", "yen", "euro", "degree", "AA", "imath", "jmath", "square", "lozenge"].includes(n));
  add(LARGE_OPS, (n) => n !== "intop" && n !== "smallint" && n !== "idotsint");
  for (const [k, v] of Object.entries({ "−": "-", "∙": "\\bullet", "′": "'", "″": "''", "‴": "'''", "√": "\\surd", "…": "\\ldots", "⋅": "\\cdot", "·": "\\cdot", "∗": "*", "⟨": "\\langle", "⟩": "\\rangle", "‖": "\\|", "∖": "\\setminus", "ℝ": "\\mathbb{R}", "ℕ": "\\mathbb{N}", "ℤ": "\\mathbb{Z}", "ℚ": "\\mathbb{Q}", "ℂ": "\\mathbb{C}", "ℍ": "\\mathbb{H}", "ℙ": "\\mathbb{P}", "°": "^{\\circ}", "÷": "\\div", "×": "\\times", "±": "\\pm", "¬": "\\neg", "∁": "\\complement", "ϵ": "\\epsilon" })) m.set(k, v);
  return m;
})();
const UNI_SUP = Object.fromEntries(Object.entries(SUP_MAP).filter(([k, v]) => v && v !== k && v.length === 1).map(([k, v]) => [v, k === "−" ? "-" : k]));
const UNI_SUB = Object.fromEntries(Object.entries(SUB_MAP).filter(([k, v]) => v && v !== k && v.length === 1).map(([k, v]) => [v, k === "−" ? "-" : k]));
const UNI_COMBINING = { "̂": "\\hat", "̄": "\\bar", "⃗": "\\vec", "̇": "\\dot", "̈": "\\ddot", "̃": "\\tilde", "̅": "\\overline", "̌": "\\check", "́": "\\acute", "̀": "\\grave", "̆": "\\breve", "̊": "\\mathring" };
const ALNUM_FONT_CMD = { bf: "\\mathbf", it: "\\mathit", bi: "\\boldsymbol", cal: "\\mathcal", frak: "\\mathfrak", bb: "\\mathbb", sf: "\\mathsf", sfbf: "\\mathsf", tt: "\\mathtt" };
/** Mathematical alphanumeric symbol -> ["\\mathbf", "x"] or null. */
function unAlnum(ch) {
  const cp = ch.codePointAt(0);
  if (cp < 0x1D400 || cp > 0x1D7FF) {
    for (const [font, holes] of Object.entries(ALNUM_HOLES)) for (const [l, c] of Object.entries(holes)) if (c === ch) return [ALNUM_FONT_CMD[font], l];
    return null;
  }
  for (const [font, [A, a, d]] of Object.entries(ALNUM)) {
    if (cp >= A && cp < A + 26) return [ALNUM_FONT_CMD[font], String.fromCharCode(65 + cp - A)];
    if (cp >= a && cp < a + 26) return [ALNUM_FONT_CMD[font], String.fromCharCode(97 + cp - a)];
    if (d && cp >= d && cp < d + 10) return [ALNUM_FONT_CMD[font], String.fromCharCode(48 + cp - d)];
  }
  if (cp >= 0x1D6A8 && cp <= 0x1D6C0) { const g = String.fromCodePoint(0x391 + cp - 0x1D6A8); return ["\\mathbf", UNI_MATH.get(g) || g]; }
  if (cp >= 0x1D6C2 && cp <= 0x1D6DA) { const g = String.fromCodePoint(0x3B1 + cp - 0x1D6C2); return ["\\boldsymbol", UNI_MATH.get(g) || g]; }
  return null;
}

/** Characters pdflatex can typeset with inputenc(utf8) + fontenc(T1) + textcomp. */
function textSafe(cp) {
  return (cp >= 0xA0 && cp <= 0x17F && cp !== 0x149) || [0x2013, 0x2014, 0x2018, 0x2019, 0x201A, 0x201C, 0x201D, 0x201E, 0x2026, 0x2022, 0x2020, 0x2021, 0x2030, 0x20AC, 0x2122, 0x2039, 0x203A, 0x2190, 0x2191, 0x2192, 0x2193, 0x2116, 0x2103, 0x02C6, 0x02DC, 0x0218, 0x0219, 0x021A, 0x021B].includes(cp);
}

class LatexWriter {
  constructor() { this.replaced = 0; }

  /** Escape plain text for LaTeX text mode. */
  text(s) {
    let out = "";
    const chars = [...String(s)];
    for (let k = 0; k < chars.length; k++) {
      const c = chars[k];
      const nx = chars[k + 1];
      switch (c) {
        case "\\": out += "\\textbackslash{}"; continue;
        case "{": out += "\\{"; continue;
        case "}": out += "\\}"; continue;
        case "$": out += "\\$"; continue;
        case "&": out += "\\&"; continue;
        case "#": out += "\\#"; continue;
        case "%": out += "\\%"; continue;
        case "_": out += "\\_"; continue;
        case "^": out += "\\textasciicircum{}"; continue;
        case "~": out += "\\textasciitilde{}"; continue;
        case "<": out += "\\textless{}"; continue;
        case ">": out += "\\textgreater{}"; continue;
        case "|": out += "\\textbar{}"; continue;
        case '"': out += "\\textquotedbl{}"; continue;
        case " ": out += "~"; continue;
        case "\t": out += "\\quad{}"; continue;
        case "\r": continue;
        case "\n": out += "\\newline{}"; continue;
        case "-": out += nx === "-" ? "-{}" : "-"; continue;
        case "`": out += "\\textasciigrave{}"; continue;
        case "'": out += nx === "'" || nx === "`" ? "'{}" : "'"; continue;
        case ",": out += nx === "," ? ",{}" : ","; continue;
        case "!": case "?": out += nx === "`" ? `${c}{}` : c; continue;
        case "‘": out += "`"; continue;
        case "’": out += "'"; continue;
        case "“": out += "``"; continue;
        case "”": out += "''"; continue;
        case "–": out += "--"; continue;
        case "—": out += "---"; continue;
        case "…": out += "\\ldots{}"; continue;
        case "€": out += "\\texteuro{}"; continue;
        case "​": case "﻿": continue;
        default: break;
      }
      const cp = c.codePointAt(0);
      if (cp < 0x80) { out += cp < 0x20 ? " " : c; continue; }
      if (textSafe(cp)) { out += c; continue; }
      if (UNI_SUP[c] !== undefined) { out += `\\textsuperscript{${this.text(UNI_SUP[c])}}`; continue; }
      if (UNI_SUB[c] !== undefined) { out += `\\textsubscript{${this.text(UNI_SUB[c])}}`; continue; }
      if (UNI_MATH.has(c)) { out += `$${UNI_MATH.get(c)}$`; continue; }
      const al = unAlnum(c);
      if (al) { out += `$${al[0]}{${al[1]}}$`; continue; }
      if (c === "✓" || c === "✔") { out += "$\\checkmark$"; continue; }
      const nfd = c.normalize("NFD");
      if (nfd !== c && textSafe(nfd.codePointAt(0)) || (nfd !== c && nfd.codePointAt(0) < 0x80)) {
        const base = String.fromCodePoint(nfd.codePointAt(0));
        const acc = { "́": "'", "̀": "`", "̂": "^", "̈": '"', "̃": "~", "̄": "=", "̇": ".", "̆": "u", "̌": "v", "̋": "H", "̧": "c", "̣": "d", "̊": "r", "̨": "k" }[nfd.slice(base.length, base.length + 1)];
        if (acc && /^[A-Za-z]$/.test(base)) { out += `\\${acc}{${base}}`; continue; }
      }
      if (/\p{M}/u.test(c)) continue;
      this.replaced++;
      out += "?";
    }
    return out;
  }

  /** Unicode approximation of math (Cambria Math run) -> LaTeX math body. */
  math(s) {
    const chars = [...String(s)];
    let out = "";
    const group = (k) => {
      // "(...)" after ^ or _ -> {...}; returns [content, nextIndex]
      if (chars[k] === "(") {
        let d = 0;
        for (let j = k; j < chars.length; j++) {
          if (chars[j] === "(") d++;
          else if (chars[j] === ")") { d--; if (d === 0) return [this.math(chars.slice(k + 1, j).join("")), j + 1]; }
        }
      }
      if (k < chars.length) return [this.math(chars[k]), k + 1];
      return ["{}", k];
    };
    for (let k = 0; k < chars.length; k++) {
      const c = chars[k];
      if (c === "^" || c === "_") { const [g, nk] = group(k + 1); out += `${c}{${g}}`; k = nk - 1; continue; }
      if (c === "√" || c === "∛" || c === "∜") {
        const [g, nk] = group(k + 1);
        out += c === "√" ? `\\sqrt{${g}}` : `\\sqrt[${c === "∛" ? 3 : 4}]{${g}}`;
        k = nk - 1;
        continue;
      }
      if (UNI_SUP[c] !== undefined) {
        let run = "";
        while (k < chars.length && UNI_SUP[chars[k]] !== undefined) run += UNI_SUP[chars[k++]];
        k--;
        out += `^{${this.math(run)}}`;
        continue;
      }
      if (UNI_SUB[c] !== undefined) {
        let run = "";
        while (k < chars.length && UNI_SUB[chars[k]] !== undefined) run += UNI_SUB[chars[k++]];
        k--;
        out += `_{${this.math(run)}}`;
        continue;
      }
      if (chars[k + 1] && UNI_COMBINING[chars[k + 1]]) {
        out += `${UNI_COMBINING[chars[k + 1]]}{${this.math(c)}}`;
        k++;
        continue;
      }
      switch (c) {
        case "{": out += "\\{"; continue;
        case "}": out += "\\}"; continue;
        case "#": out += "\\#"; continue;
        case "$": out += "\\$"; continue;
        case "%": out += "\\%"; continue;
        case "&": out += "\\&"; continue;
        case "~": out += "\\sim "; continue;
        case "\\": out += "\\backslash "; continue;
        case "\n": case "\t": case " ": case " ": case " ": case " ": out += " "; continue;
        default: break;
      }
      const cp = c.codePointAt(0);
      if (cp < 0x80) { out += cp < 0x20 ? " " : c; continue; }
      if (UNI_MATH.has(c)) { const cmd = UNI_MATH.get(c); out += /[A-Za-z]$/.test(cmd) ? cmd + " " : cmd; continue; }
      const al = unAlnum(c);
      if (al) { out += `${al[0]}{${al[1]}}`; continue; }
      if (UNI_COMBINING[c]) continue;
      if (textSafe(cp)) { out += `\\text{${c}}`; continue; }
      this.replaced++;
      out += "?";
    }
    return out.replace(/\s+$/, "");
  }

  url(u) {
    return String(u).replace(/[\\{}^\s]/g, (c) => "%" + c.charCodeAt(0).toString(16).toUpperCase().padStart(2, "0")).replace(/[#%&_~]/g, (c) => (c === "%" ? "\\%" : c === "#" ? "\\#" : c));
  }

  runs(runs, opts = {}) {
    let out = "";
    const list = (runs || []).filter((r) => r && typeof r.t === "string" && r.t.length);
    for (let k = 0; k < list.length; k++) {
      const r = list[k];
      if (r.font === "Cambria Math" || r.font === "Cambria  Math") {
        const m = this.math(r.t).trim();
        if (m) out += `$${m}$`;
        continue;
      }
      const pieces = String(r.t).split("\n");
      pieces.forEach((piece, pi) => {
        if (pi > 0) out += opts.noBreaks ? " " : "\\newline{}";
        if (!piece) return;
        let s = this.text(piece);
        if (r.font && /courier|consolas|mono/i.test(r.font)) s = `\\texttt{${s}}`;
        if (r.b) s = `\\textbf{${s}}`;
        if (r.i) s = `\\textit{${s}}`;
        if (r.u) s = `\\underline{${s}}`;
        if (r.s) s = `\\sout{${s}}`;
        if (r.sup) s = `\\textsuperscript{${s}}`;
        else if (r.sub) s = `\\textsubscript{${s}}`;
        if (r.color && /^#[0-9a-f]{6}$/i.test(r.color)) s = `\\textcolor[HTML]{${r.color.slice(1).toUpperCase()}}{${s}}`;
        if (r.hl && /^#[0-9a-f]{6}$/i.test(r.hl)) s = `\\colorbox[HTML]{${r.hl.slice(1).toUpperCase()}}{${s}}`;
        if (r.link && !String(r.link).startsWith("#") && !opts.noLinks) s = `\\href{${this.url(r.link)}}{${s}}`;
        out += s;
      });
    }
    // no line break at the very start or end of a paragraph
    out = out.replace(/^(\\newline\{\})+/, "").replace(/(\\newline\{\})+$/, "");
    return out;
  }
}

const plainRuns = (b) => ((b && b.runs) || []).map((r) => r.t || "").join("").replace(/\s+/g, " ").trim();

/** word model -> standalone, compilable LaTeX (article class). */
export function wordModelToLatex(model) {
  const W = new LatexWriter();
  const meta = (model && model.meta) || {};
  let blocks = Array.isArray(model && model.blocks) ? model.blocks.slice() : [];
  const title = typeof meta.title === "string" ? meta.title.trim() : "";
  const author = typeof meta.author === "string" ? meta.author.trim() : "";
  let date = null;
  // drop the title/author/date paragraphs that duplicate the metadata (round trip of \maketitle)
  if (title && blocks.length && /^(h1|p)$/.test(blocks[0].type) && plainRuns(blocks[0]) === title) {
    blocks.shift();
    if (author && blocks.length && blocks[0].type === "p" && plainRuns(blocks[0]) === author) blocks.shift();
    if (blocks.length && blocks[0].type === "p" && blocks[0].align === "center" && /^[A-Z][a-z]+ \d{1,2}, \d{4}$/.test(plainRuns(blocks[0]))) date = plainRuns(blocks.shift());
  }
  const body = [];
  let needLongtable = false, needArray = false;
  const listStack = [];
  let listChunk = -1;
  // all lines of one (nested) list go into a single body entry
  const listLine = (line) => {
    if (listChunk >= 0 && listChunk === body.length - 1) body[listChunk] += "\n" + line;
    else { body.push(line); listChunk = body.length - 1; }
  };
  const closeLists = (toLevel) => {
    while (listStack.length > toLevel) listLine(`${"  ".repeat(listStack.length - 1)}\\end{${listStack.pop()}}`);
  };
  for (let k = 0; k < blocks.length; k++) {
    const b = blocks[k];
    if (!b || typeof b !== "object") continue;
    if (b.type !== "list") closeLists(0);
    switch (b.type) {
      case "p": case "quote": {
        if (b.type === "quote") {
          const group = [];
          while (k < blocks.length && blocks[k].type === "quote") group.push(blocks[k++]);
          k--;
          const parts = group.map((q) => W.runs(q.runs)).filter((s) => s.trim());
          if (parts.length) body.push(`\\begin{quote}\n${parts.join("\n\n")}\n\\end{quote}`);
          break;
        }
        const txt = W.runs(b.runs);
        if (!txt.trim()) break;
        if (b.align === "center" || b.align === "right") {
          const env = b.align === "center" ? "center" : "flushright";
          const group = [txt];
          while (k + 1 < blocks.length && blocks[k + 1].type === "p" && blocks[k + 1].align === b.align) {
            const t2 = W.runs(blocks[++k].runs);
            if (t2.trim()) group.push(t2);
          }
          body.push(`\\begin{${env}}\n${group.join("\n\n")}\n\\end{${env}}`);
        } else body.push(txt);
        break;
      }
      case "h1": case "h2": case "h3": case "h4": case "h5": case "h6": {
        const txt = W.runs(b.runs, { noBreaks: true });
        if (!txt.trim()) break;
        const cmd = { h1: "section", h2: "subsection", h3: "subsubsection" }[b.type] || "paragraph";
        body.push(`\\${cmd}{${txt}}`);
        break;
      }
      case "list": {
        const level = Math.max(0, Math.min(3, Number(b.level) || 0));
        const env = b.list === "ol" ? "enumerate" : "itemize";
        // close deeper lists, or a list of the other type at this level
        closeLists(level + 1);
        if (listStack.length === level + 1 && listStack[level] !== env) closeLists(level);
        while (listStack.length < level + 1) {
          const lvl = listStack.length;
          const e = lvl === level ? env : "itemize";
          listLine(`${"  ".repeat(lvl)}\\begin{${e}}`);
          listStack.push(e);
          if (lvl < level) listLine(`${"  ".repeat(lvl + 1)}\\item[]`);
        }
        const txt = W.runs(b.runs);
        listLine(`${"  ".repeat(level + 1)}\\item ${txt.startsWith("[") ? "{}" : ""}${txt}`);
        // keep the list open only if the next block is also a list item
        break;
      }
      case "code": {
        const lines = [];
        while (k < blocks.length && blocks[k].type === "code") {
          const t = ((blocks[k].runs || []).map((r) => r.t || "").join(""));
          for (const ln of t.split("\n")) lines.push(ln);
          k++;
        }
        k--;
        const safe = lines.map((ln) => [...ln.replace(/\\end\{verbatim\}/g, "\\end {verbatim}")].map((c) => {
          const cp = c.codePointAt(0);
          if (cp < 0x80 || textSafe(cp)) return c === "\t" ? "    " : c;
          W.replaced++;
          return "?";
        }).join(""));
        body.push(`\\begin{verbatim}\n${safe.join("\n")}\n\\end{verbatim}`);
        break;
      }
      case "table": {
        const rows = (b.rows || []).filter((r) => Array.isArray(r));
        if (!rows.length) break;
        const ncols = Math.max(1, ...rows.map((r) => r.reduce((a, c) => a + Math.max(1, Number(c && c.colspan) || 1), 0)));
        const cellText = (c) => ((c && c.blocks) || []).map((x) => {
          if (x.type === "table") return "[table]";
          if (x.type === "img") return `[Image: ${W.text(x.alt || "image")}]`;
          if (!x.runs) return "";
          const t = W.runs(x.runs, { noBreaks: true });
          return x.type === "list" ? `\\textbullet{} ${t}` : t;
        }).filter((s) => s.trim());
        const long = rows.some((r) => r.some((c) => cellText(c).join(" ").replace(/\\[a-z]+\{?|[{}]/gi, "").length > 30));
        const width = (n) => `p{${((0.94 * n) / ncols).toFixed(3)}\\linewidth}`;
        const colSpec = long ? width(1) : "l";
        if (long) needArray = true;
        const spec = "|" + Array.from({ length: ncols }, () => `${colSpec}|`).join("");
        const longtab = rows.length > 40;
        if (longtab) needLongtable = true;
        const env = longtab ? "longtable" : "tabular";
        const lines = [`\\begin{${env}}{${spec}}`, "\\hline"];
        for (const r of rows) {
          const cells = [];
          let used = 0;
          r.forEach((c) => {
            const span = Math.max(1, Math.min(ncols - used, Number(c && c.colspan) || 1));
            if (span <= 0) return;
            const content = cellText(c).join(long ? " \\newline " : " ");
            if (span > 1) cells.push(`\\multicolumn{${span}}{${used === 0 ? "|" : ""}${long ? width(span) : "l"}|}{${content}}`);
            else cells.push(content);
            used += span;
          });
          while (used < ncols) { cells.push(""); used++; }
          lines.push(`${cells.join(" & ")} \\\\`, "\\hline");
        }
        lines.push(`\\end{${env}}`);
        body.push(longtab ? lines.join("\n") : `\\noindent\n${lines.join("\n")}`);
        break;
      }
      case "img": {
        const alt = String(b.alt || "image").replace(/\s+/g, " ").trim() || "image";
        body.push(`% image omitted: ${alt.replace(/[\r\n]/g, " ")}\n\\fbox{Image: ${W.text(alt)}}`);
        break;
      }
      case "pagebreak": body.push("\\newpage"); break;
      default: break;
    }
    if (b.type === "list" && !(blocks[k + 1] && blocks[k + 1].type === "list")) closeLists(0);
  }
  closeLists(0);

  const pre = [
    "\\documentclass[11pt]{article}",
    "\\usepackage[utf8]{inputenc}",
    "\\usepackage[T1]{fontenc}",
    "\\usepackage{textcomp}",
    "\\usepackage{amsmath}",
    "\\usepackage{amssymb}",
    "\\usepackage{graphicx}",
    "\\usepackage{xcolor}",
    "\\usepackage[normalem]{ulem}",
  ];
  if (needArray) pre.push("\\usepackage{array}");
  if (needLongtable) pre.push("\\usepackage{longtable}");
  pre.push("\\usepackage{hyperref}");
  if (W.replaced) pre.unshift(`% Note: ${W.replaced} character${W.replaced === 1 ? "" : "s"} that pdflatex cannot typeset were replaced by "?".`);
  const head = [];
  if (title || author) {
    head.push(`\\title{${W.text(title).replace(/\\newline\{\}/g, " ")}}`);
    head.push(`\\author{${W.text(author).replace(/\\newline\{\}/g, " ")}}`);
    head.push(`\\date{${date ? W.text(date) : ""}}`);
  }
  const doc = [...pre, "", ...head, ...(head.length ? [""] : []), "\\begin{document}", ...(head.length ? ["\\maketitle", ""] : [""])];
  return `${doc.join("\n")}${body.join("\n\n")}\n\n\\end{document}\n`;
}

// ===========================================================================
// templates
// ===========================================================================

export const TEMPLATES = [
  { id: "article", label: "Article", description: "Sections, equations, a table and a bibliography-ready layout" },
  { id: "report", label: "Report", description: "Chapters, table of contents and references" },
  { id: "letter", label: "Letter", description: "A formal letter with addresses and signature" },
  { id: "beamer", label: "Presentation (Beamer)", description: "Slides with a title page, lists, blocks and columns" },
  { id: "math", label: "Math homework", description: "Numbered problems with proofs and solutions (amsmath/amsthm)" },
  { id: "cv", label: "CV / Résumé", description: "A clean one-page curriculum vitae" },
];

const TPL = {
  article: String.raw`\documentclass[11pt]{article}
\usepackage[utf8]{inputenc}
\usepackage[T1]{fontenc}
\usepackage{amsmath, amssymb}
\usepackage{graphicx}
\usepackage{hyperref}

\title{An Example Article}
\author{Your Name}
\date{\today}

\begin{document}
\maketitle

\begin{abstract}
Summarize the article in two or three sentences: the question, the method and the main result.
\end{abstract}

\section{Introduction}
\label{sec:intro}
Start writing here. Text can be \textbf{bold}, \textit{italic} or \underline{underlined}, and
footnotes are numbered automatically.\footnote{Like this one.} Section~\ref{sec:results} shows
some mathematics.

\section{Results}
\label{sec:results}
Inline math such as $e^{i\pi} + 1 = 0$ sits in the text; numbered equations stand on their own:
\begin{equation}
  \int_0^\infty e^{-x^2}\,dx = \frac{\sqrt{\pi}}{2}.
  \label{eq:gauss}
\end{equation}
Equation~\eqref{eq:gauss} is the Gaussian integral. Lists are easy too:
\begin{itemize}
  \item a first point,
  \item a second point.
\end{itemize}

\begin{table}[h]
  \centering
  \begin{tabular}{|l|c|r|}
    \hline
    Item & Quantity & Price \\
    \hline
    Apples & 3 & 1.20 \\
    Pears  & 5 & 2.50 \\
    \hline
  \end{tabular}
  \caption{A small table.}
  \label{tab:prices}
\end{table}

\section{Conclusion}
Table~\ref{tab:prices} lists the prices. Summarize the main points here.

\end{document}
`,
  report: String.raw`\documentclass[11pt]{report}
\usepackage[utf8]{inputenc}
\usepackage[T1]{fontenc}
\usepackage{amsmath, amssymb}
\usepackage{graphicx}
\usepackage{hyperref}

\title{Project Report}
\author{Your Name}
\date{\today}

\begin{document}
\maketitle
\tableofcontents

\chapter{Introduction}
\label{ch:intro}
This report describes the project. Chapter~\ref{ch:method} explains the method and
Chapter~\ref{ch:results} presents the results~\cite{knuth}.

\section{Background}
Write the background here.

\section{Goals}
\begin{enumerate}
  \item Define the problem.
  \item Build a solution.
  \item Evaluate it.
\end{enumerate}

\chapter{Method}
\label{ch:method}
\section{Approach}
Describe the approach. A key relation is
\begin{equation}
  E = m c^2 .
  \label{eq:energy}
\end{equation}

\chapter{Results}
\label{ch:results}
\section{Findings}
Equation~\eqref{eq:energy} was confirmed.

\begin{thebibliography}{9}
\bibitem{knuth} Donald E. Knuth. \emph{The \TeX book}. Addison-Wesley, 1984.
\end{thebibliography}

\end{document}
`,
  letter: String.raw`\documentclass[11pt]{letter}
\usepackage[utf8]{inputenc}
\usepackage[T1]{fontenc}

\signature{Your Name}
\address{Your Street 1 \\ 12345 Your Town}

\begin{document}

\begin{letter}{Ms.~Jane Smith \\ Example Company \\ Main Street 2 \\ 54321 Other Town}

\opening{Dear Ms.~Smith,}

Thank you for your letter. Write the first paragraph here: say why you are writing.

Use the second paragraph for the details, and the last one for what should happen next.

\closing{Yours sincerely,}

\end{letter}

\end{document}
`,
  beamer: String.raw`\documentclass{beamer}
\usetheme{Madrid}
\usepackage[utf8]{inputenc}
\usepackage[T1]{fontenc}

\title{Presentation Title}
\subtitle{A short subtitle}
\author{Your Name}
\institute{Your Organization}
\date{\today}

\begin{document}

\begin{frame}
  \titlepage
\end{frame}

\section{Introduction}

\begin{frame}{Why this matters}
  \begin{itemize}
    \item The first reason
    \item The second reason
    \item The third reason
  \end{itemize}
\end{frame}

\begin{frame}{A formula}
  \begin{block}{Pythagoras}
    In a right triangle: $a^2 + b^2 = c^2$.
  \end{block}
  \begin{alertblock}{Remember}
    Units must match.
  \end{alertblock}
\end{frame}

\section{Details}

\begin{frame}{Two columns}
  \begin{columns}
    \begin{column}{0.5\textwidth}
      Text on the left side.
    \end{column}
    \begin{column}{0.5\textwidth}
      \begin{enumerate}
        \item One
        \item Two
      \end{enumerate}
    \end{column}
  \end{columns}
\end{frame}

\begin{frame}{Thank you}
  \centering
  Questions?
\end{frame}

\end{document}
`,
  math: String.raw`\documentclass[11pt]{article}
\usepackage[utf8]{inputenc}
\usepackage[T1]{fontenc}
\usepackage[margin=1in]{geometry}
\usepackage{amsmath, amssymb, amsthm}

\newtheorem{problem}{Problem}
\newcommand{\R}{\mathbb{R}}
\newcommand{\N}{\mathbb{N}}

\title{Homework 1}
\author{Your Name}
\date{\today}

\begin{document}
\maketitle

\begin{problem}
Show that $\sum_{k=1}^{n} k = \frac{n(n+1)}{2}$ for every $n \in \N$.
\end{problem}

\begin{proof}
By induction on $n$. For $n = 1$ both sides equal $1$. If the formula holds for $n$, then
\begin{align*}
  \sum_{k=1}^{n+1} k &= \frac{n(n+1)}{2} + (n+1) \\
                     &= \frac{(n+1)(n+2)}{2}.
\end{align*}
\end{proof}

\begin{problem}
Find all $x \in \R$ with $x^2 - 5x + 6 = 0$.
\end{problem}

\begin{proof}[Solution]
Factor: $x^2 - 5x + 6 = (x-2)(x-3)$, so $x \in \{2, 3\}$.
\end{proof}

\begin{problem}
Compute $\det A$ for
\[
  A = \begin{pmatrix} 2 & 1 \\ 1 & 3 \end{pmatrix}.
\]
\end{problem}

\begin{proof}[Solution]
$\det A = 2 \cdot 3 - 1 \cdot 1 = 5$.
\end{proof}

\end{document}
`,
  cv: String.raw`\documentclass[11pt]{article}
\usepackage[utf8]{inputenc}
\usepackage[T1]{fontenc}
\usepackage[margin=2cm]{geometry}
\usepackage{xcolor}
\usepackage{enumitem}
\usepackage{hyperref}

\pagestyle{empty}
\definecolor{accent}{HTML}{008080}
\newcommand{\cvsection}[1]{\section*{\textcolor{accent}{#1}}}
\newcommand{\entry}[4]{\noindent\textbf{#1} \hfill #2 \\ \textit{#3} \hfill \textit{#4} \par\smallskip}

\begin{document}

\begin{center}
  {\Huge\textbf{Jane Doe}} \\[6pt]
  Main Street 1, 12345 Town \quad\textbullet\quad +1 234 567 890 \quad\textbullet\quad
  \href{mailto:jane.doe@example.com}{jane.doe@example.com}
\end{center}

\cvsection{Profile}
Analyst with five years of experience in data analysis and reporting. Calm under pressure,
careful with details and happy to explain numbers to people who do not like numbers.

\cvsection{Experience}
\entry{Senior Analyst}{2021 -- today}{Example Corp, Town}{Full time}
\begin{itemize}[noitemsep]
  \item Built the monthly reporting used by 40 managers.
  \item Cut the time to close the books from 10 to 6 days.
\end{itemize}

\entry{Analyst}{2018 -- 2021}{Other Company, City}{Full time}
\begin{itemize}[noitemsep]
  \item Automated data checks with Python and SQL.
\end{itemize}

\cvsection{Education}
\entry{M.Sc. Economics}{2016 -- 2018}{University of Somewhere}{Grade: 1.3}

\cvsection{Skills}
\begin{tabular}{@{}ll}
  \textbf{Languages} & English (native), German (fluent) \\
  \textbf{Tools} & Python, SQL, Excel, \LaTeX \\
\end{tabular}

\end{document}
`,
};

/** A new LaTeX document from a template id (default: article). */
export function newLatexDocument(template) {
  return TPL[template] || TPL.article;
}

// ===========================================================================
// editor support data (completion, symbols)
// ===========================================================================

/** [command (without backslash), text inserted after the backslash, description] */
export const COMMAND_COMPLETIONS = [
  ["section", "section{}", "Section heading"], ["subsection", "subsection{}", "Subsection heading"],
  ["subsubsection", "subsubsection{}", "Sub-subsection heading"], ["chapter", "chapter{}", "Chapter (report/book)"],
  ["paragraph", "paragraph{}", "Run-in paragraph heading"], ["part", "part{}", "Part"],
  ["textbf", "textbf{}", "Bold text"], ["textit", "textit{}", "Italic text"], ["emph", "emph{}", "Emphasis"],
  ["underline", "underline{}", "Underlined text"], ["texttt", "texttt{}", "Typewriter text"],
  ["textsc", "textsc{}", "Small caps"], ["textcolor", "textcolor{red}{}", "Colored text"],
  ["textsuperscript", "textsuperscript{}", "Superscript text"], ["textsubscript", "textsubscript{}", "Subscript text"],
  ["footnote", "footnote{}", "Footnote"], ["label", "label{}", "Label for cross references"],
  ["ref", "ref{}", "Reference to a label"], ["eqref", "eqref{}", "Reference to an equation"],
  ["pageref", "pageref{}", "Page of a label"], ["cite", "cite{}", "Citation"],
  ["href", "href{https://}{}", "Hyperlink"], ["url", "url{}", "URL"],
  ["includegraphics", "includegraphics[width=0.8\\textwidth]{}", "Insert an image"],
  ["caption", "caption{}", "Caption of a figure or table"], ["centering", "centering", "Center the content"],
  ["item", "item ", "List item"], ["begin", "begin{}", "Begin an environment"], ["end", "end{}", "End an environment"],
  ["frac", "frac{}{}", "Fraction"], ["dfrac", "dfrac{}{}", "Display-style fraction"], ["sqrt", "sqrt{}", "Square root"],
  ["sum", "sum_{}^{}", "Sum"], ["prod", "prod_{}^{}", "Product"], ["int", "int_{}^{}", "Integral"],
  ["lim", "lim_{}", "Limit"], ["infty", "infty", "Infinity ∞"], ["partial", "partial", "Partial ∂"],
  ["left", "left( ", "Growing left delimiter"], ["right", "right)", "Growing right delimiter"],
  ["mathbb", "mathbb{}", "Blackboard bold (ℝ, ℕ)"], ["mathbf", "mathbf{}", "Bold math"], ["mathrm", "mathrm{}", "Upright math"],
  ["mathcal", "mathcal{}", "Calligraphic letters"], ["text", "text{}", "Text inside math"], ["operatorname", "operatorname{}", "Function name"],
  ["hat", "hat{}", "Hat accent"], ["bar", "bar{}", "Bar accent"], ["vec", "vec{}", "Vector arrow"], ["overline", "overline{}", "Overline"],
  ["cdot", "cdot", "Centered dot ⋅"], ["times", "times", "Times ×"], ["leq", "leq", "Less or equal ≤"], ["geq", "geq", "Greater or equal ≥"],
  ["neq", "neq", "Not equal ≠"], ["approx", "approx", "Approximately ≈"], ["in", "in", "Element of ∈"], ["to", "to", "Arrow →"],
  ["Rightarrow", "Rightarrow", "Implies ⇒"], ["ldots", "ldots", "Dots …"], ["cdots", "cdots", "Centered dots ⋯"],
  ["alpha", "alpha", "α"], ["beta", "beta", "β"], ["gamma", "gamma", "γ"], ["delta", "delta", "δ"], ["epsilon", "epsilon", "ϵ"],
  ["theta", "theta", "θ"], ["lambda", "lambda", "λ"], ["mu", "mu", "μ"], ["pi", "pi", "π"], ["sigma", "sigma", "σ"],
  ["phi", "phi", "ϕ"], ["omega", "omega", "ω"], ["Delta", "Delta", "Δ"], ["Sigma", "Sigma", "Σ"], ["Omega", "Omega", "Ω"],
  ["newcommand", "newcommand{\\}{}", "Define a command"], ["renewcommand", "renewcommand{\\}{}", "Redefine a command"],
  ["usepackage", "usepackage{}", "Load a package"], ["documentclass", "documentclass{article}", "Document class"],
  ["title", "title{}", "Document title"], ["author", "author{}", "Document author"], ["date", "date{\\today}", "Document date"],
  ["maketitle", "maketitle", "Print the title block"], ["tableofcontents", "tableofcontents", "Table of contents"],
  ["newpage", "newpage", "Start a new page"], ["clearpage", "clearpage", "New page (flush floats)"],
  ["noindent", "noindent", "No paragraph indent"], ["vspace", "vspace{1em}", "Vertical space"], ["hspace", "hspace{1em}", "Horizontal space"],
  ["quad", "quad", "Wide space"], ["qquad", "qquad", "Wider space"], ["today", "today", "Today's date"], ["LaTeX", "LaTeX", "LaTeX logo"],
  ["multicolumn", "multicolumn{2}{c}{}", "Cell spanning columns"], ["hline", "hline", "Horizontal table rule"],
  ["cline", "cline{1-2}", "Partial table rule"], ["newtheorem", "newtheorem{theorem}{Theorem}", "Define a theorem environment"],
  ["bibitem", "bibitem{}", "Bibliography entry"], ["frametitle", "frametitle{}", "Slide title (beamer)"],
  ["small", "small", "Smaller text"], ["large", "large", "Larger text"], ["Large", "Large", "Even larger text"], ["huge", "huge", "Huge text"],
  ["color", "color{red}", "Switch text color"], ["boxed", "boxed{}", "Boxed formula"], ["binom", "binom{}{}", "Binomial coefficient"],
  ["overbrace", "overbrace{}^{}", "Brace above"], ["underbrace", "underbrace{}_{}", "Brace below"],
];

export const ENVIRONMENT_COMPLETIONS = [
  "itemize", "enumerate", "description", "equation", "equation*", "align", "align*", "gather", "gather*", "multline",
  "figure", "table", "tabular", "center", "flushleft", "flushright", "quote", "quotation", "verse", "abstract",
  "verbatim", "lstlisting", "minipage", "theorem", "lemma", "proof", "definition", "example", "remark", "corollary",
  "proposition", "frame", "block", "alertblock", "columns", "column", "thebibliography", "matrix", "pmatrix", "bmatrix",
  "vmatrix", "cases", "array", "split", "aligned", "document", "letter", "multicols", "subfigure",
];

const sym = (names, tbl) => names.map((n) => [`\\${n}`, tbl[n]]).filter((x) => x[1]);
export const SYMBOL_GROUPS = [
  { label: "Greek", items: [...sym(Object.keys(GREEK_LOWER).filter((n) => n !== "omicron" && n !== "digamma"), GREEK_LOWER), ...sym(["Gamma", "Delta", "Theta", "Lambda", "Xi", "Pi", "Sigma", "Upsilon", "Phi", "Psi", "Omega"], GREEK_UPPER)] },
  { label: "Operators", items: [...sym(["pm", "mp", "times", "div", "cdot", "ast", "star", "circ", "bullet", "oplus", "otimes", "cap", "cup", "setminus", "wedge", "vee"], MATH_OPS), ["\\sum", "∑"], ["\\prod", "∏"], ["\\int", "∫"], ["\\oint", "∮"], ["\\iint", "∬"], ["\\bigcup", "⋃"], ["\\bigcap", "⋂"], ["\\sqrt{x}", "√"], ["\\frac{a}{b}", "a⁄b"], ["\\partial", "∂"], ["\\nabla", "∇"], ["\\infty", "∞"]] },
  { label: "Relations", items: sym(["leq", "geq", "neq", "approx", "equiv", "sim", "simeq", "cong", "propto", "ll", "gg", "subset", "supset", "subseteq", "supseteq", "in", "notin", "ni", "perp", "parallel", "mid", "models", "vdash", "prec", "succ"], MATH_OPS) },
  { label: "Arrows", items: sym(["to", "leftarrow", "leftrightarrow", "Rightarrow", "Leftarrow", "Leftrightarrow", "mapsto", "longrightarrow", "Longrightarrow", "implies", "iff", "uparrow", "downarrow", "nearrow", "searrow", "hookrightarrow", "rightleftharpoons"], MATH_OPS) },
  { label: "Logic & sets", items: [...sym(["forall", "exists", "nexists", "neg", "therefore", "because"], MATH_OPS), ["\\emptyset", "∅"], ["\\mathbb{R}", "ℝ"], ["\\mathbb{N}", "ℕ"], ["\\mathbb{Z}", "ℤ"], ["\\mathbb{Q}", "ℚ"], ["\\mathbb{C}", "ℂ"], ["\\aleph", "ℵ"], ["\\ell", "ℓ"], ["\\hbar", "ℏ"], ["\\ldots", "…"], ["\\cdots", "⋯"], ["\\vdots", "⋮"], ["\\ddots", "⋱"]] },
];

/** Names of commands defined in the source (\newcommand, \def, \DeclareMathOperator). */
export function userCommandNames(source) {
  const out = new Set();
  const re = /\\(?:(?:re)?newcommand|providecommand|DeclareRobustCommand|DeclareMathOperator)\*?\s*\{?\s*\\([A-Za-z@]+)|\\[gex]?def\s*\\([A-Za-z@]+)/g;
  let m;
  const s = String(source || "");
  let guard = 0;
  while ((m = re.exec(s)) && guard++ < 5000) out.add(m[1] || m[2]);
  return [...out];
}

/** Labels, bibliography keys and environment names defined in the source (for completion). */
export function latexKeys(source) {
  const s = String(source || "");
  const grab = (re) => { const o = new Set(); let m, g = 0; while ((m = re.exec(s)) && g++ < 20000) o.add(m[1].trim()); return [...o]; };
  return {
    labels: grab(/\\label\s*\{([^}]+)\}/g),
    cites: grab(/\\bibitem(?:\[[^\]]*\])?\s*\{([^}]+)\}/g),
    envs: grab(/\\(?:newtheorem|newenvironment|renewenvironment)\*?\s*\{([^}]+)\}/g),
  };
}
