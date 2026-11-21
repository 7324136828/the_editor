/**
 * LaTeX support (lib/latex.js): tokenizer, document parser, live-preview HTML,
 * TeX math -> MathML, LaTeX <-> word model, templates, robustness.
 * Run: node --test --test-isolation=none tests/latex.test.js
 */

import { test } from "node:test";
import assert from "node:assert/strict";

import {
  tokenize, latexToHtml, latexToHtmlDocument, mathToMathML, mathToUnicode,
  latexToWordModel, wordModelToLatex, newLatexDocument, TEMPLATES,
  COMMAND_COMPLETIONS, ENVIRONMENT_COMPLETIONS, SYMBOL_GROUPS, userCommandNames, latexKeys, PREVIEW_CSS,
} from "../lib/latex.js";

const doc = (body, pre = "") => `\\documentclass{article}\n${pre}\\begin{document}\n${body}\n\\end{document}\n`;
const html = (body, pre) => latexToHtml(doc(body, pre));
const errs = (r, sev) => r.errors.filter((e) => !sev || e.severity === sev);
const text = (h) => h.replace(/<math[\s\S]*?<\/math>/g, "[M]").replace(/<[^>]+>/g, "").replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"');
const runsText = (b) => (b.runs || []).map((r) => r.t).join("");
const BT = "`";

// ---------------------------------------------------------------- tokenizer

test("tokenizer: commands, groups, text, paragraphs and line numbers", () => {
  const t = tokenize("Hello \\textbf{world}\n\nNext");
  assert.deepEqual(t.map((x) => x.k), ["txt", "cmd", "ch", "txt", "ch", "par", "txt"]);
  assert.equal(t[1].v, "textbf");
  assert.equal(t[6].l, 3);
});

test("tokenizer: comments eat the newline and leading blanks; \\% is not a comment", () => {
  const t = tokenize("a% comment\n   b \\% c");
  assert.equal(t.filter((x) => x.k === "txt").map((x) => x.v).join(""), "ab  c");
  assert.ok(t.some((x) => x.k === "cmd" && x.v === "%"));
});

test("tokenizer: spaces after control words are skipped, control symbols keep them", () => {
  const t = tokenize("\\LaTeX  is \\& more");
  assert.equal(t[1].v, "is ");
  assert.equal(t[3].v, " more");
});

test("tokenizer: verbatim, \\verb, comment environment and $$", () => {
  const t = tokenize("\\verb|a{b}%c| x\n\\begin{verbatim}\n\\section{no}\n% not a comment\n\\end{verbatim}\n\\begin{comment}\nhidden\n\\end{comment}\n$$x$$");
  assert.equal(t[0].k, "vinl");
  assert.equal(t[0].v, "a{b}%c");
  const v = t.find((x) => x.k === "verb");
  assert.equal(v.v, "\\section{no}\n% not a comment");
  assert.equal(v.l, 2);
  assert.ok(t.some((x) => x.k === "skip"));
  assert.ok(t.some((x) => x.k === "ch" && x.v === "$$"));
});

test("tokenizer: macro parameters and blank line containing only spaces", () => {
  const t = tokenize("#1 ## x\n   \n y");
  assert.equal(t[0].k, "param");
  assert.equal(t[2].k, "param");
  assert.ok(t.some((x) => x.k === "par"));
});

// ---------------------------------------------------------------- structure

test("sections are numbered like LaTeX (article) and appear in the outline", () => {
  const r = html("\\section{Intro}\n\\subsection{Sub}\n\\subsubsection{Deep}\n\\paragraph{Para}\n\\section*{Star}\n\\section{Two}");
  assert.deepEqual(r.outline.map((o) => [o.level, o.num, o.title]), [[2, "1", "Intro"], [3, "1.1", "Sub"], [4, "1.1.1", "Deep"], [5, "", "Para"], [2, "", "Star"], [2, "2", "Two"]]);
  assert.match(r.html, /<h2 class="ltx-section" data-line="3"[^>]*><span class="ltx-secnum">1<\/span>Intro<\/h2>/);
  assert.equal(r.outline[0].line, 3);
});

test("report class: chapters number sections, subsubsections are not numbered", () => {
  const r = latexToHtml("\\documentclass{report}\n\\begin{document}\n\\chapter{One}\n\\section{A}\n\\subsection{B}\n\\subsubsection{C}\n\\chapter{Two}\n\\section{D}\n\\end{document}");
  assert.deepEqual(r.outline.map((o) => o.num), ["1", "1.1", "1.1.1", "", "2", "2.1"]);
  assert.match(r.html, /Chapter 1<\/span>One/);
});

test("appendix switches to letters", () => {
  const r = html("\\section{A}\\appendix\\section{Data}\\subsection{More}");
  assert.deepEqual(r.outline.map((o) => o.num), ["1", "A", "A.1"]);
});

test("title block from \\title/\\author/\\date and \\maketitle", () => {
  const r = html("\\maketitle", "\\title{My \\textbf{Title}}\\author{Ann \\and Bob}\\date{1 May 2024}\n");
  assert.match(r.html, /<div class="ltx-title">My <span style="font-weight:bold">Title<\/span><\/div>/);
  assert.match(r.html, /ltx-author">Ann <span class="ltx-and"><\/span>Bob/);
  assert.match(r.html, /ltx-date">1 May 2024/);
});

test("paragraphs split on blank lines and carry data-line", () => {
  const r = html("First line\ncontinues.\n\nSecond.");
  const ps = r.html.match(/<p data-line="(\d+)"/g);
  assert.equal(ps.length, 2);
  assert.match(r.html, /<p data-line="3"[^>]*>First line continues\.<\/p>/);
  assert.match(r.html, /<p data-line="6"/);
});

test("text formatting, sizes, colors and declarations scoped by groups", () => {
  const r = html("\\textbf{b} \\textit{i} \\emph{e \\emph{u}} \\underline{u} \\sout{s} \\texttt{t} \\textsc{sc} \\textsf{sf} {\\bfseries B} {\\large L} \\textcolor{red}{R} {\\color[HTML]{00AA00} G} \\textcolor[rgb]{0,0,1}{blue} \\colorbox{yellow}{hl} \\textsuperscript{2} \\textsubscript{x} after", "\\usepackage{xcolor}\n");
  const h = r.html;
  assert.match(h, /font-weight:bold">b</);
  assert.match(h, /font-style:italic">i</);
  assert.match(h, /font-style:italic">e <\/span>u/, "nested \\emph turns italics off");
  assert.match(h, /text-decoration:underline">u</);
  assert.match(h, /text-decoration:line-through">s</);
  assert.match(h, /font-family:var\(--ltx-mono\)">t</);
  assert.match(h, /font-variant:small-caps">sc</);
  assert.match(h, /font-size:1\.2em">L</);
  assert.match(h, /color:#FF0000">R</);
  assert.match(h, /color:#00AA00">G</);
  assert.match(h, /color:#0000FF">blue</);
  assert.match(h, /background-color:#FFFF00">hl</);
  assert.match(h, /<sup>2<\/sup>/);
  assert.match(h, /<sub>x<\/sub>/);
  assert.match(h, /\safter<\/p>/, "styles end with their group");
});

test("special characters, quotes, dashes and accents", () => {
  const r = html(`50\\% \\& \\$ \\# \\_ \\{ \\} \\textbackslash{} a~b -- --- ${BT}${BT}q'' \\ldots{} \\'e \\"{o} \\c{c} \\ss{} \\LaTeX\\ rocks`);
  const t = text(r.html);
  assert.ok(t.includes("50% & $ # _ { } \\ a\u00A0b – — “q” … é ö ç ß LaTeX rocks"), t);
});

test("lists: itemize, enumerate with nested numbering, description and \\item[label]", () => {
  const r = html("\\begin{enumerate}\\item one\\begin{enumerate}\\item sub\\item sub2\\end{enumerate}\\item two\\end{enumerate}\n\\begin{itemize}\\item a\\item[--] b\\end{itemize}\n\\begin{description}\\item[Key] value\\end{description}");
  const h = r.html;
  assert.deepEqual([...h.matchAll(/ltx-mark">([^<]*)</g)].map((m) => m[1]), ["1.", "(a)", "(b)", "2.", "•", "–"]);
  assert.match(h, /<span class="ltx-dt">Key<\/span> value/);
  assert.equal(errs(r, "error").length, 0);
});

test("enumerate label options (enumitem and enumerate package style)", () => {
  const r = html("\\begin{enumerate}[label=(\\roman*)]\\item a\\item b\\end{enumerate}\\begin{enumerate}[A.]\\item c\\end{enumerate}");
  assert.deepEqual([...r.html.matchAll(/ltx-mark">([^<]*)</g)].map((m) => m[1]), ["(i)", "(ii)", "A."]);
});

test("quote, center, flushright, verbatim and \\verb", () => {
  const r = html("\\begin{quote}Q\\end{quote}\\begin{center}C\\end{center}\\begin{flushright}R\\end{flushright}\n\\begin{verbatim}\n  x < y & z\n\\end{verbatim}\n\\verb|$x$|");
  assert.match(r.html, /<blockquote class="ltx-quote ltx-quote"/);
  assert.match(r.html, /text-align:center[^>]*>C<\/p>/);
  assert.match(r.html, /text-align:right[^>]*>R<\/p>/);
  assert.match(r.html, /<pre class="ltx-code"><code>  x &lt; y &amp; z<\/code><\/pre>/);
  assert.match(r.html, /<code class="ltx-verb">\$x\$<\/code>/);
});

test("tables: column spec, rules, \\multicolumn and \\cline", () => {
  const r = html("\\begin{tabular}{|l|c|r|}\\hline A & B & C \\\\ \\hline \\multicolumn{2}{|c|}{wide} & x \\\\ \\cline{1-2} d & e & f \\\\ \\hline\\end{tabular}");
  const h = r.html;
  assert.equal((h.match(/<tr /g) || []).length, 3);
  assert.match(h, /<td colspan="2" style="text-align:center;border-left:1px solid currentColor;border-right:1px solid currentColor;border-top/);
  assert.match(h, /text-align:right[^"]*">C</);
  assert.match(h, /border-bottom:1px solid currentColor[^"]*">f</, "trailing \\hline -> bottom rule");
  assert.equal(errs(r, "error").length, 0);
});

test("tables: too many cells is reported with the row's line", () => {
  const r = html("\\begin{tabular}{ll}\na & b & c \\\\\n\\end{tabular}");
  const e = errs(r, "error");
  assert.equal(e.length, 1);
  assert.equal(e[0].line, 4);
  assert.match(e[0].message, /Extra &/);
});

test("figures and tables: numbered captions, labels and \\ref", () => {
  const r = html("\\begin{figure}\\centering\\includegraphics[width=0.5\\textwidth]{img/cat.png}\\caption{A cat}\\label{fig:cat}\\end{figure}\n\\begin{table}\\caption{Data}\\label{tab:d}\\begin{tabular}{l}x\\end{tabular}\\end{table}\nSee Figure~\\ref{fig:cat} and Table~\\ref{tab:d}.");
  const h = r.html;
  assert.match(h, /<span class="ltx-caplabel">Figure 1:<\/span> A cat/);
  assert.match(h, /<span class="ltx-caplabel">Table 1:<\/span> Data/);
  assert.match(h, /class="ltx-img" style="width:50%"[^>]*>.*cat\.png/);
  assert.match(text(h), /See Figure\u00A01 and Table\u00A01\./);
  assert.ok(h.indexOf("Table 1:") < h.indexOf("<table"), "caption above the table when written first");
});

test("references: sections, equations, \\eqref, unknown -> ?? with a warning", () => {
  const r = html("\\section{A}\\label{s:a}\n\\begin{equation}x=1\\label{eq:x}\\end{equation}\nSee \\ref{s:a}, \\eqref{eq:x}, \\autoref{eq:x} and \\ref{missing}.");
  const t = text(r.html);
  assert.match(t, /See 1, \(1\), Equation 1 and \?\?\./);
  const w = errs(r, "warning");
  assert.equal(w.length, 1);
  assert.match(w[0].message, /"missing"/);
  assert.equal(w[0].line, 5);
});

test("citations numbered from thebibliography; unknown -> [?]", () => {
  const r = html("\\cite{b} \\cite[p.~5]{a} \\cite{a,b} \\cite{zz}\n\\begin{thebibliography}{9}\\bibitem{a} Alpha.\\bibitem{b} Beta.\\end{thebibliography}");
  const t = text(r.html);
  assert.match(t, /\[2\] \[1, p.\u00A05\] \[1, 2\] \[\?\]/);
  assert.match(r.html, /ltx-bibnum">\[1\]<\/span> Alpha\./);
  assert.ok(errs(r, "warning").some((e) => /"zz"/.test(e.message)));
});

test("footnotes are numbered and collected at the end", () => {
  const r = html("A\\footnote{First.} B\\footnote{Second \\textbf{bold}.}");
  assert.match(r.html, /<sup class="ltx-fnref"><a href="#ltx-fn-1" data-fn="1">1<\/a><\/sup>/);
  assert.match(r.html, /<section class="ltx-footnotes">.*id="ltx-fn-2"[^>]*><sup>2<\/sup> Second <span style="font-weight:bold">bold<\/span>\./);
});

test("macros: \\newcommand with arguments and default, \\def, \\renewcommand", () => {
  const r = html("\\hi{Ann} \\opt{} \\opt[X] \\twice{ab} $\\R$","\\newcommand{\\hi}[1]{Hello #1!}\n\\newcommand{\\opt}[1][D]{(#1)}\n\\def\\twice#1{#1#1}\n\\newcommand\\R{\\mathbb{R}}\n");
  assert.match(text(r.html), /Hello Ann! \(D\) \(X\) abab/);
  assert.equal(errs(r).length, 0, JSON.stringify(r.errors));
});

test("macros used inside math are expanded by the math converter", () => {
  const r = html("$x\\in\\R$, $\\norm{v}$", "\\newcommand{\\R}{\\mathbb{R}}\\newcommand{\\norm}[1]{\\left\\| #1 \\right\\|}\n");
  assert.match(r.html, /<mi>ℝ<\/mi>/);
  assert.match(r.html, /<mo fence="true" stretchy="true" symmetric="true" form="prefix">‖<\/mo><mi>v<\/mi>/);
});

test("recursive macros stop with one error instead of hanging", () => {
  const t0 = Date.now();
  const r = latexToHtml("\\def\\a{x\\a}\\a and \\newcommand{\\b}{\\b}\\b");
  assert.ok(Date.now() - t0 < 2000);
  assert.ok(errs(r, "error").some((e) => /Too many macro expansions/.test(e.message)));
});

test("theorems from \\newtheorem, shared counters, proof with QED", () => {
  const r = html("\\begin{theorem}[Main]A\\end{theorem}\\begin{lemma}B\\end{lemma}\\begin{theorem}C\\end{theorem}\\begin{definition}D\\end{definition}\\begin{proof}Done.\\end{proof}", "\\newtheorem{theorem}{Theorem}\\newtheorem{lemma}[theorem]{Lemma}\\theoremstyle{definition}\\newtheorem{definition}{Definition}\n");
  const t = text(r.html);
  assert.match(t, /Theorem 1 \(Main\)\. A/);
  assert.match(t, /Lemma 2\. B/);
  assert.match(t, /Theorem 3\. C/);
  assert.match(t, /Definition 1\. D/);
  assert.match(r.html, /ltx-thm-plain/);
  assert.match(r.html, /ltx-thm-definition/);
  assert.match(r.html, /<span class="ltx-proof-head">Proof\.<\/span> Done\.<span class="ltx-qed">∎<\/span>/);
});

test("display math: numbered equation, align rows, \\nonumber, \\tag, starred", () => {
  const r = html("\\begin{equation}a=b\\end{equation}\\begin{align}x&=1\\\\y&=2\\nonumber\\\\z&=3\\label{eq:z}\\end{align}\\begin{equation*}c\\end{equation*}\\begin{equation}d\\tag{$*$}\\end{equation}\\[e\\]$$f$$ \\eqref{eq:z}");
  const nums = [...r.html.matchAll(/ltx-eqno">([^<]*)</g)].map((m) => m[1]);
  assert.deepEqual(nums, ["(1)", "(2)", "(3)", "($*$)"]);
  assert.match(text(r.html), /\(3\)\s*$/);
  assert.match(r.html, /<table class="ltx-align">/);
  assert.equal((r.html.match(/<math display="block"/g) || []).length, 5);
});

test("report class numbers equations and figures by chapter", () => {
  const r = latexToHtml("\\documentclass{report}\\begin{document}\\chapter{A}\\chapter{B}\\begin{equation}x\\end{equation}\\begin{figure}\\caption{F}\\end{figure}\\end{document}");
  assert.match(r.html, /ltx-eqno">\(2\.1\)/);
  assert.match(r.html, /Figure 2\.1:/);
});

test("beamer frames render as slides with titles", () => {
  const r = latexToHtml(newLatexDocument("beamer"));
  assert.equal(errs(r, "error").length, 0);
  assert.ok((r.html.match(/class="ltx-frame"/g) || []).length >= 4);
  assert.match(r.html, /<div class="ltx-frame-title">Why this matters<\/div>/);
  assert.match(r.html, /ltx-titleframe/);
  assert.ok(r.outline.some((o) => o.frame && o.title === "A formula"));
});

test("unknown commands keep their text and produce ONE warning listing them", () => {
  const r = html("\\foo{alpha} \\baz[x]{beta} \\foo{gamma}");
  assert.match(text(r.html), /alpha beta gamma/);
  const w = errs(r, "warning");
  assert.equal(w.length, 1);
  assert.match(w[0].message, /\\foo \(line 3\), \\baz \(line 3\)/);
});

test("unknown environment: content shown, warning with line", () => {
  const r = html("x\n\\begin{weird}inside\\end{weird}");
  assert.match(text(r.html), /inside/);
  assert.ok(errs(r, "warning").some((e) => e.line === 4 && /weird/.test(e.message)));
});

test("errors: unbalanced braces with line numbers", () => {
  let r = html("ok\n\\textbf{never closed\n\nnext");
  assert.ok(errs(r, "error").some((e) => e.line === 4 && /Missing \}/.test(e.message)), JSON.stringify(r.errors));
  assert.match(text(r.html), /never closed/);
  r = html("a } b");
  assert.ok(errs(r, "error").some((e) => e.line === 3 && /Extra \}/.test(e.message)));
  r = html("{ open group");
  assert.ok(errs(r, "error").some((e) => e.line === 3 && /Missing \}/.test(e.message)));
});

test("errors: unclosed and mismatched environments", () => {
  let r = latexToHtml("\\documentclass{article}\n\\begin{document}\n\\begin{itemize}\n\\item a\n\\end{document}");
  assert.ok(errs(r, "error").some((e) => e.line === 3 && /itemize/.test(e.message)), JSON.stringify(r.errors));
  r = html("\\begin{itemize}\n\\item a\n\\end{enumerate}\nafter");
  const e = errs(r, "error");
  assert.ok(e.some((x) => x.line === 5 && /does not match \\begin\{itemize\}/.test(x.message)), JSON.stringify(e));
  assert.match(text(r.html), /after/);
});

test("errors: missing \\end{document} and missing \\begin{document}", () => {
  let r = latexToHtml("\\documentclass{article}\n\\begin{document}\nHello\n");
  assert.ok(errs(r, "error").some((e) => /Missing \\end\{document\}/.test(e.message)));
  assert.match(r.html, /Hello/);
  r = latexToHtml("\\documentclass{article}\n\\title{x}\n");
  assert.ok(errs(r, "error").some((e) => e.line === 1 && /Missing \\begin\{document\}/.test(e.message)));
});

test("errors: unclosed inline math stops at the blank line", () => {
  const r = html("$x^2\n\nnext paragraph");
  assert.ok(errs(r, "error").some((e) => e.line === 3 && /not closed/.test(e.message)));
  assert.match(text(r.html), /next paragraph/);
});

test("text after \\end{document} is ignored; preamble text is reported", () => {
  const r = latexToHtml("\\documentclass{article}\nstray\n\\begin{document}\nIn\n\\end{document}\nAfter");
  assert.doesNotMatch(r.html, /After|stray/);
  assert.ok(errs(r, "error").some((e) => e.line === 2));
});

test("fragments without \\documentclass render directly", () => {
  const r = latexToHtml("\\section{Hi} $a+b$");
  assert.match(r.html, /Hi<\/h2>/);
  assert.equal(r.errors.length, 0);
});

test("tikz pictures and \\input are placeholders, not errors", () => {
  const r = html("\\begin{tikzpicture}\\draw (0,0) -- (1,1);\\end{tikzpicture}\\input{chapter1}");
  assert.match(r.html, /TikZ drawing/);
  assert.equal(errs(r, "error").length, 0);
});

// ---------------------------------------------------------------- math

test("math: identifiers, numbers and operators", () => {
  const m = mathToMathML("x+12.5=y-z");
  assert.equal(m, '<math alttext="x+12.5=y-z"><mrow><mi>x</mi><mo>+</mo><mn>12.5</mn><mo>=</mo><mi>y</mi><mo>−</mo><mi>z</mi></mrow></math>');
});

test("math: fractions, binomials, roots", () => {
  assert.match(mathToMathML("\\frac{a+b}{c}"), /<mfrac><mrow><mi>a<\/mi><mo>\+<\/mo><mi>b<\/mi><\/mrow><mi>c<\/mi><\/mfrac>/);
  assert.match(mathToMathML("\\frac12"), /<mfrac><mn>1<\/mn><mn>2<\/mn><\/mfrac>/);
  assert.match(mathToMathML("\\dfrac{1}{x}"), /<mstyle displaystyle="true" scriptlevel="0"><mfrac>/);
  assert.match(mathToMathML("\\binom{n}{k}"), /<mfrac linethickness="0"><mi>n<\/mi><mi>k<\/mi><\/mfrac>/);
  assert.match(mathToMathML("\\sqrt{2}"), /<msqrt><mn>2<\/mn><\/msqrt>/);
  assert.match(mathToMathML("\\sqrt[3]{x}"), /<mroot><mi>x<\/mi><mn>3<\/mn><\/mroot>/);
});

test("math: scripts, primes, limits in display vs inline", () => {
  assert.match(mathToMathML("x_i^2"), /<msubsup><mi>x<\/mi><mi>i<\/mi><mn>2<\/mn><\/msubsup>/);
  assert.match(mathToMathML("f''(x)"), /<msup><mi>f<\/mi><mo[^>]*>″<\/mo><\/msup>/);
  assert.match(mathToMathML("\\sum_{i=1}^n i", true), /<munderover><mo movablelimits="false">∑<\/mo>/);
  assert.match(mathToMathML("\\sum_{i=1}^n i", false), /<msubsup><mo>∑<\/mo>/);
  assert.match(mathToMathML("\\int_0^1 f", true), /<msubsup><mo>∫<\/mo>/);
  assert.match(mathToMathML("\\int\\limits_0^1 f", true), /<munderover><mo movablelimits="false">∫<\/mo>/);
  assert.match(mathToMathML("\\lim_{x\\to0} f", true), /<munder><mi>lim<\/mi><mrow><mi>x<\/mi><mo>→<\/mo><mn>0<\/mn><\/mrow><\/munder>/);
  assert.match(mathToMathML("\\max_{k} a_k", false), /<msub><mi>max<\/mi><mi>k<\/mi><\/msub>/);
});

test("math: \\left \\right, \\big, \\middle", () => {
  const m = mathToMathML("\\left( \\frac{a}{b} \\middle| c \\right]");
  assert.match(m, /<mo fence="true" stretchy="true" symmetric="true" form="prefix">\(<\/mo>/);
  assert.match(m, /<mo stretchy="true" fence="true">\|<\/mo>/);
  assert.match(m, /form="postfix">\]<\/mo>/);
  assert.match(mathToMathML("\\left. x \\right\\}"), /<mrow><mi>x<\/mi><mo[^>]*form="postfix">\}<\/mo><\/mrow>/);
  assert.match(mathToMathML("\\big( x \\Big)"), /<mo stretchy="true" minsize="1.2em" maxsize="1.2em">\(<\/mo>/);
  assert.match(mathToMathML("(x)"), /<mo stretchy="false">\(<\/mo>/, "plain parentheses do not stretch");
});

test("math: matrices, cases, array with lines, aligned", () => {
  const p = mathToMathML("\\begin{pmatrix} a & b \\\\ c & d \\end{pmatrix}");
  assert.match(p, /form="prefix">\(<\/mo><mtable displaystyle="false">/);
  assert.equal((p.match(/<mtr>/g) || []).length, 2);
  assert.equal((p.match(/<mtd /g) || []).length, 4);
  assert.match(mathToMathML("\\begin{bmatrix}1\\end{bmatrix}"), /form="prefix">\[<\/mo>/);
  assert.match(mathToMathML("\\begin{vmatrix}1\\end{vmatrix}"), /form="prefix">\|<\/mo>/);
  const c = mathToMathML("f = \\begin{cases} 1 & x>0 \\\\ 0 & \\text{otherwise} \\end{cases}");
  assert.match(c, /form="prefix">\{<\/mo><mtable/);
  assert.match(c, /<mtext>otherwise<\/mtext>/);
  const a = mathToMathML("\\begin{array}{c|l} a & b \\\\ \\hline c & d \\end{array}");
  assert.match(a, /border-right:1px solid currentColor/);
  assert.match(a, /border-top:1px solid currentColor/);
  const al = mathToMathML("\\begin{aligned} x &= 1 \\\\ y &= 2 \\end{aligned}", true);
  assert.match(al, /columnalign="right left"/);
  assert.match(al, /<mtd columnalign="left"[^>]*><mrow><\/mrow><mo>=<\/mo>/);
});

test("math: Greek, fonts (Unicode math alphanumerics), accents, text, functions", () => {
  assert.match(mathToMathML("\\alpha\\Gamma\\varepsilon"), /<mi>α<\/mi><mi mathvariant="normal">Γ<\/mi><mi>ε<\/mi>/);
  assert.match(mathToMathML("\\mathbb{R}\\mathbf{x}\\mathcal{A}\\mathfrak{g}\\mathsf{s}\\mathtt{t}"), /<mi>ℝ<\/mi><mi>𝐱<\/mi><mi>𝒜<\/mi><mi>𝔤<\/mi><mi>𝗌<\/mi><mi>𝚝<\/mi>/);
  assert.match(mathToMathML("\\mathrm{d}x"), /<mi mathvariant="normal">d<\/mi><mi>x<\/mi>/);
  assert.match(mathToMathML("\\boldsymbol{\\alpha}"), /<mi>𝜶<\/mi>/);
  assert.match(mathToMathML("\\hat{x}\\vec{v}\\overline{AB}"), /<mover accent="true"><mi>x<\/mi><mo stretchy="false">\^<\/mo><\/mover>.*<mo stretchy="true">‾<\/mo>/);
  assert.match(mathToMathML("\\underbrace{a+b}_{n}", true), /<munder><munder accentunder="false">/);
  assert.match(mathToMathML("\\text{if } x"), /<mtext>if&#160;<\/mtext>/);
  assert.match(mathToMathML("\\sin x"), /<mi>sin<\/mi><mspace width="0.1667em"\/><mi>x<\/mi>/);
  assert.match(mathToMathML("\\operatorname{rank} A"), /<mi>rank<\/mi>/);
});

test("math: relations, arrows, dots, spacing, \\not, color, primes", () => {
  const m = mathToMathML("a\\leq b \\neq c \\to d \\Rightarrow e \\ldots \\cdots \\, \\quad \\not\\in \\not= \\color{red} z");
  for (const s of ["≤", "≠", "→", "⇒", "…", "⋯", "∉"]) assert.ok(m.includes(`<mo>${s}</mo>`), s);
  assert.match(m, /<mspace width="0.1667em"\/>/);
  assert.match(m, /<mspace width="1em"\/>/);
  assert.match(m, /<mrow style="color:#FF0000"><mi>z<\/mi><\/mrow>/);
});

test("math: malformed input never throws and marks the error", () => {
  for (const bad of ["\\frac{a", "x^", "}", "\\left(", "\\begin{pmatrix} a", "\\undefinedcmd", "{{{{", "a_b_c", "\\sqrt[", "\\end{x}", "", "&", "\\\\"]) {
    const m = mathToMathML(bad, true);
    assert.match(m, /^<math[^>]*>[\s\S]*<\/math>$/, bad);
  }
  assert.match(mathToMathML("\\undefinedcmd x"), /<merror><mtext>\\undefinedcmd<\/mtext><\/merror>/);
  assert.match(mathToMathML("a}"), /<merror>/);
  const deep = "{".repeat(5000) + "x" + "}".repeat(5000);
  assert.match(mathToMathML(deep), /<math/);
});

test("math errors are reported in the problems list with the line", () => {
  const r = html("ok\n$\\frac{a$ and $\\foo$");
  assert.ok(r.errors.some((e) => e.line === 4 && /Missing \}/.test(e.message)), JSON.stringify(r.errors));
  assert.ok(r.errors.some((e) => e.line === 4 && e.severity === "warning" && /\\foo/.test(e.message)));
});

test("mathToUnicode: readable approximations", () => {
  assert.equal(mathToUnicode("x^2 + y_i"), "x² + yᵢ");
  assert.equal(mathToUnicode("\\frac{a+b}{2}"), "(a + b)/2");
  assert.equal(mathToUnicode("\\frac{1}{2}"), "1/2");
  assert.equal(mathToUnicode("\\sqrt{x+1}"), "√(x + 1)");
  assert.equal(mathToUnicode("\\alpha \\leq \\beta"), "α ≤ β");
  assert.equal(mathToUnicode("\\sum_{i=1}^{n} i"), "∑ᵢ₌₁ⁿi");
  assert.equal(mathToUnicode("\\mathbb{R}^n"), "ℝⁿ");
  assert.equal(mathToUnicode("x^{a+b}"), "xᵃ⁺ᵇ");
  assert.equal(mathToUnicode("e^{i\\pi}"), "e^(iπ)");
});

// ---------------------------------------------------------------- word model

test("LaTeX -> word: title, headings, paragraphs and runs per the mapping", () => {
  const m = latexToWordModel(doc("\\maketitle\n\\section{One}\nA \\textbf{b} \\textit{i} \\underline{u} \\sout{s} \\texttt{t} \\textsuperscript{2}\\textsubscript{x} \\textcolor{red}{r} \\colorbox{yellow}{h} \\href{https://x.org}{link} \\url{https://y.org}\n\\subsection{Two}\\subsubsection{Three}\\paragraph{Four}x\\subparagraph{Five}y", "\\usepackage{xcolor}\\title{T}\\author{Au}\\date{}\n"));
  assert.equal(m.kind, "word");
  assert.equal(m.meta.title, "T");
  assert.equal(m.meta.author, "Au");
  assert.deepEqual(m.blocks[0], { type: "h1", align: "center", runs: [{ t: "T" }] });
  assert.deepEqual(m.blocks[1], { type: "p", align: "center", runs: [{ t: "Au" }] });
  assert.deepEqual(m.blocks.filter((b) => /^h/.test(b.type)).map((b) => b.type), ["h1", "h1", "h2", "h3", "h4", "h5"]);
  const p = m.blocks.find((b) => b.type === "p" && runsText(b).startsWith("A "));
  const find = (t) => p.runs.find((r) => r.t === t);
  assert.ok(find("b").b && find("i").i && find("u").u && find("s").s);
  assert.equal(find("t").font, "Courier New");
  assert.ok(find("2").sup && find("x").sub);
  assert.equal(find("r").color, "#FF0000");
  assert.equal(find("h").hl, "#FFFF00");
  assert.equal(find("link").link, "https://x.org");
  assert.equal(find("https://y.org").link, "https://y.org");
});

test("LaTeX -> word: report class shifts sections by one level", () => {
  const m = latexToWordModel("\\documentclass{report}\\begin{document}\\chapter{C}\\section{S}\\subsection{U}\\end{document}");
  assert.deepEqual(m.blocks.map((b) => b.type), ["h1", "h2", "h3"]);
});

test("LaTeX -> word: lists, description, quote, code, table, footnotes, page break, images", () => {
  const m = latexToWordModel(doc("\\begin{itemize}\\item a\\begin{enumerate}\\item b\\end{enumerate}\\end{itemize}\n\\begin{description}\\item[K] v\\end{description}\n\\begin{quote}q\\end{quote}\n\\begin{verbatim}\nl1\nl2\n\\end{verbatim}\n\\begin{tabular}{|l|l|}\\hline x & y\\\\\\multicolumn{2}{|c|}{z}\\\\\\hline\\end{tabular}\n\nText\\footnote{Note.}\\newpage\n\\begin{center}\\includegraphics{pic.png}\\end{center}"));
  const sig = m.blocks.map((b) => `${b.type}${b.list ? `:${b.list}${b.level}` : ""}${b.align ? `@${b.align}` : ""}:${b.type === "table" ? b.rows.length : runsText(b)}`);
  assert.deepEqual(sig, ["list:ul0:a", "list:ol1:b", "list:ul0:K v", "quote:q", "code:l1", "code:l2", "table:2", "p:Text[1]", "pagebreak:", "p@center:[Image: pic.png]", "h1:Notes", "p:1. Note."]);
  const desc = m.blocks[2];
  assert.ok(desc.runs[0].b && desc.runs[0].t === "K");
  const tbl = m.blocks[6];
  assert.equal(tbl.rows[1][0].colspan, 2);
  assert.equal(runsText(tbl.rows[0][1].blocks[0]), "y");
  assert.ok(m.blocks[7].runs.find((r) => r.t === "[1]").sup);
  assert.ok(m.meta.warnings.some((w) => /Images are not embedded/.test(w)));
});

test("LaTeX -> word: math becomes Cambria Math runs; display math centered, one paragraph per align row", () => {
  const m = latexToWordModel(doc("Let $\\alpha^2 \\leq \\frac{1}{2}$.\n\\begin{align*}a &= b\\\\c &\\to d\\end{align*}"));
  const p = m.blocks[0];
  const mr = p.runs.find((r) => r.font === "Cambria Math");
  assert.equal(mr.t, "α² ≤ 1/2");
  assert.deepEqual(m.blocks.slice(1).map((b) => [b.type, b.align, runsText(b), b.runs[0].font]), [["p", "center", "a = b", "Cambria Math"], ["p", "center", "c → d", "Cambria Math"]]);
});

test("LaTeX -> word: unknown commands are kept as text with one warning", () => {
  const m = latexToWordModel(doc("\\mystery{kept} text"));
  assert.equal(runsText(m.blocks[0]), "kept text");
  assert.equal(m.meta.warnings.length, 1);
  assert.match(m.meta.warnings[0], /\\mystery/);
});

test("word -> LaTeX: preamble, title, headings, formatting and escaping", () => {
  const tex = wordModelToLatex({
    kind: "word", meta: { title: "R&D #1", author: "Ann" },
    blocks: [
      { type: "h1", runs: [{ t: "R&D #1" }] },
      { type: "h1", runs: [{ t: "Intro" }] },
      { type: "h2", runs: [{ t: "Sub" }] }, { type: "h3", runs: [{ t: "Deep" }] }, { type: "h5", runs: [{ t: "Low" }] },
      { type: "p", runs: [{ t: "50% of $5 & #3 a_b {x} ~ ^ \\ " }, { t: "bold", b: true }, { t: "it", i: true }, { t: "u", u: true }, { t: "s", s: true }, { t: "2", sup: true }, { t: "i", sub: true }, { t: "red", color: "#ff0000" }, { t: "hl", hl: "#FFFF00" }, { t: "site", link: "https://e.com/a_b?x=1&y=%20#top" }, { t: "code", font: "Courier New" }] },
      { type: "p", align: "center", runs: [{ t: "centered" }] },
      { type: "p", align: "right", runs: [{ t: "right" }] },
    ],
  });
  assert.match(tex, /^\\documentclass\[11pt\]\{article\}/);
  for (const pkg of ["\\usepackage[utf8]{inputenc}", "\\usepackage[T1]{fontenc}", "\\usepackage{amsmath}", "\\usepackage{graphicx}", "\\usepackage{xcolor}", "\\usepackage{hyperref}", "\\usepackage[normalem]{ulem}"]) assert.ok(tex.includes(pkg), pkg);
  assert.match(tex, /\\title\{R\\&D \\#1\}\n\\author\{Ann\}/);
  assert.match(tex, /\\maketitle/);
  assert.equal((tex.match(/\\section\{/g) || []).length, 1, "title paragraph is not repeated");
  assert.match(tex, /\\section\{Intro\}\n\n\\subsection\{Sub\}\n\n\\subsubsection\{Deep\}\n\n\\paragraph\{Low\}/);
  assert.ok(tex.includes("50\\% of \\$5 \\& \\#3 a\\_b \\{x\\} \\textasciitilde{} \\textasciicircum{} \\textbackslash{} "));
  for (const s of ["\\textbf{bold}", "\\textit{it}", "\\underline{u}", "\\sout{s}", "\\textsuperscript{2}", "\\textsubscript{i}", "\\textcolor[HTML]{FF0000}{red}", "\\colorbox[HTML]{FFFF00}{hl}", "\\href{https://e.com/a_b?x=1&y=\\%20\\#top}{site}", "\\texttt{code}"]) assert.ok(tex.includes(s), s);
  assert.match(tex, /\\begin\{center\}\ncentered\n\\end\{center\}/);
  assert.match(tex, /\\begin\{flushright\}\nright\n\\end\{flushright\}/);
  assert.match(tex, /\\end\{document\}\n$/);
});

test("word -> LaTeX: nested lists, tables with colspan, code, quote, image, page break, math", () => {
  const tex = wordModelToLatex({
    kind: "word", meta: {},
    blocks: [
      { type: "list", list: "ul", level: 0, runs: [{ t: "a" }] },
      { type: "list", list: "ol", level: 1, runs: [{ t: "b" }] },
      { type: "list", list: "ol", level: 1, runs: [{ t: "[c]" }] },
      { type: "list", list: "ul", level: 0, runs: [{ t: "d" }] },
      { type: "table", rows: [[{ blocks: [{ type: "p", runs: [{ t: "A&B" }] }] }, { blocks: [{ type: "p", runs: [{ t: "C" }] }] }], [{ colspan: 2, blocks: [{ type: "p", runs: [{ t: "wide" }] }] }]] },
      { type: "code", runs: [{ t: "if (a) {" }] }, { type: "code", runs: [{ t: "  \\end{verbatim} 100%" }] },
      { type: "quote", runs: [{ t: "q1" }] }, { type: "quote", runs: [{ t: "q2" }] },
      { type: "img", src: "data:image/png;base64,AAAA", alt: "Logo" },
      { type: "pagebreak" },
      { type: "p", runs: [{ t: "x² + α ≤ √(a+b)", font: "Cambria Math" }] },
    ],
  });
  assert.match(tex, /\\begin\{itemize\}\n {2}\\item a\n {2}\\begin\{enumerate\}\n {4}\\item b\n {4}\\item \{\}\[c\]\n {2}\\end\{enumerate\}\n {2}\\item d\n\\end\{itemize\}/);
  assert.match(tex, /\\begin\{tabular\}\{\|l\|l\|\}\n\\hline\nA\\&B & C \\\\\n\\hline\n\\multicolumn\{2\}\{\|l\|\}\{wide\} \\\\\n\\hline\n\\end\{tabular\}/);
  assert.match(tex, /\\begin\{verbatim\}\nif \(a\) \{\n {2}\\end \{verbatim\} 100%\n\\end\{verbatim\}/);
  assert.match(tex, /\\begin\{quote\}\nq1\n\nq2\n\\end\{quote\}/);
  assert.match(tex, /% image omitted: Logo\n\\fbox\{Image: Logo\}/);
  assert.match(tex, /\\newpage/);
  assert.match(tex, /\$x\^\{2\} \+ \\alpha {1,2}\\leq {1,2}\\sqrt\{a\+b\}\$/);
});

test("word -> LaTeX: Unicode pdflatex cannot typeset is replaced, Greek goes to math", () => {
  const tex = wordModelToLatex({ kind: "word", meta: {}, blocks: [{ type: "p", runs: [{ t: "café – “ok” … € α 日本 😀" }] }] });
  assert.ok(tex.includes("café -- ``ok'' \\ldots{} \\texteuro{} $\\alpha$ ?? ?"), tex);
  assert.match(tex, /^% Note: 3 characters that pdflatex cannot typeset were replaced by "\?"\./);
});

test("round trip: word -> LaTeX -> word keeps text, escaping and formatting", () => {
  const original = {
    kind: "word", meta: { title: "Plan" },
    blocks: [
      { type: "h1", runs: [{ t: "Costs & {budget}" }] },
      { type: "p", runs: [{ t: "Save 50% on $5_000 items #1 ~ ^ \\ -- ``x'' " }, { t: "bold", b: true }, { t: " " }, { t: "link", link: "https://ex.com/?a=1&b=2" }] },
      { type: "list", list: "ol", level: 0, runs: [{ t: "first" }] },
      { type: "list", list: "ol", level: 1, runs: [{ t: "nested" }] },
      { type: "quote", runs: [{ t: "wise words" }] },
      { type: "table", rows: [[{ blocks: [{ type: "p", runs: [{ t: "a_b" }] }] }, { blocks: [{ type: "p", runs: [{ t: "c%d" }] }] }]] },
      { type: "code", runs: [{ t: "x = {1, 2}" }] },
      { type: "pagebreak" },
      { type: "p", align: "center", runs: [{ t: "end" }] },
    ],
  };
  const back = latexToWordModel(wordModelToLatex(original));
  assert.equal(back.meta.title, "Plan");
  const sig = back.blocks.map((b) => `${b.type}${b.list ? `:${b.list}${b.level}` : ""}${b.align ? `@${b.align}` : ""}:${b.type === "table" ? b.rows.map((r) => r.map((c) => runsText(c.blocks[0])).join("|")).join("/") : runsText(b)}`);
  assert.deepEqual(sig, [
    "h1@center:Plan",
    "h1:Costs & {budget}",
    "p:Save 50% on $5_000 items #1 ~ ^ \\ -- ``x'' bold link",
    "list:ol0:first", "list:ol1:nested", "quote:wise words", "table:a_b|c%d", "code:x = {1, 2}", "pagebreak:", "p@center:end",
  ]);
  const p = back.blocks[2];
  assert.ok(p.runs.find((r) => r.t === "bold").b);
  assert.equal(p.runs.find((r) => r.t === "link").link, "https://ex.com/?a=1&b=2");
});

test("round trip: LaTeX -> word -> LaTeX -> word is stable", () => {
  const src = newLatexDocument("article");
  const m1 = latexToWordModel(src);
  const m2 = latexToWordModel(wordModelToLatex(m1));
  const sig = (m) => m.blocks.map((b) => `${b.type}:${b.type === "table" ? b.rows.length : runsText(b).replace(/\s+/g, "")}`);
  assert.deepEqual(sig(m2), sig(m1));
  assert.equal(latexToHtml(wordModelToLatex(m1)).errors.filter((e) => e.severity === "error").length, 0);
});

// ---------------------------------------------------------------- templates & misc

test("every template parses without errors and has content", () => {
  assert.ok(TEMPLATES.length >= 6);
  for (const t of TEMPLATES) {
    const src = newLatexDocument(t.id);
    assert.match(src, /\\begin\{document\}[\s\S]*\\end\{document\}/, t.id);
    const r = latexToHtml(src);
    assert.deepEqual(r.errors.filter((e) => e.severity === "error"), [], `${t.id}: ${JSON.stringify(r.errors)}`);
    assert.deepEqual(r.errors.filter((e) => /Unknown/.test(e.message)), [], `${t.id}: ${JSON.stringify(r.errors)}`);
    assert.ok(r.html.length > 500, t.id);
    const m = latexToWordModel(src);
    assert.ok(m.blocks.length > 3, t.id);
  }
  assert.equal(newLatexDocument(), newLatexDocument("article"));
  assert.equal(newLatexDocument("nope"), newLatexDocument("article"));
});

test("standalone HTML document embeds CSS and MathML", () => {
  const h = latexToHtmlDocument(newLatexDocument("math"), "Home <work>");
  assert.match(h, /^<!doctype html>/);
  assert.match(h, /<title>Home &lt;work&gt;<\/title>/);
  assert.ok(h.includes(".ltx-doc{"));
  assert.match(h, /<math display="block"/);
  assert.doesNotMatch(h, /<script/i);
  assert.ok(PREVIEW_CSS.includes("@media print"));
});

test("links are sanitized", () => {
  const r = html("\\href{javascript:alert(1)}{x} \\url{www.example.com}");
  assert.doesNotMatch(r.html, /javascript:/);
  assert.match(r.html, /href="https:\/\/www\.example\.com"/);
});

test("HTML output escapes text and attributes", () => {
  const r = html("<script>alert(1)</script> \"q\" & \\label{a\"b}\\ref{a\"b}");
  assert.doesNotMatch(r.html, /<script>/);
  assert.match(r.html, /&lt;script&gt;/);
});

test("completion data and helpers", () => {
  assert.ok(COMMAND_COMPLETIONS.length > 80);
  assert.ok(COMMAND_COMPLETIONS.every((c) => c.length === 3 && typeof c[1] === "string"));
  assert.ok(ENVIRONMENT_COMPLETIONS.includes("itemize"));
  assert.ok(SYMBOL_GROUPS.every((g) => g.items.length > 5 && g.items.every((i) => i[0].startsWith("\\") && i[1])));
  assert.deepEqual(userCommandNames("\\newcommand{\\R}{x}\\renewcommand\\foo[1]{}\\def\\bar#1{}\\DeclareMathOperator{\\tr}{tr}").sort(), ["R", "bar", "foo", "tr"]);
  assert.deepEqual(latexKeys("\\label{a}\\label{b}\\bibitem[x]{k1}\\newtheorem{thm}{T}"), { labels: ["a", "b"], cites: ["k1"], envs: ["thm"] });
});

test("robustness: hostile input terminates quickly", () => {
  const inputs = [
    "{".repeat(20000), "}".repeat(20000), "\\begin{itemize}".repeat(3000), "$".repeat(10001), "\\item ".repeat(5000),
    "\\frac".repeat(5000), "\\\\".repeat(5000), "&".repeat(5000), "\\left(".repeat(3000), "%".repeat(1000) + "\n" + "a",
    "\\newcommand{\\x}{\\x\\x}\\x", "\\begin{tabular}{" + "l".repeat(5000) + "}" + "a&".repeat(5000) + "\\end{tabular}",
  ];
  for (const s of inputs) {
    const t0 = Date.now();
    const r = latexToHtml(s);
    assert.equal(typeof r.html, "string");
    assert.ok(Date.now() - t0 < 3000, `slow on ${s.slice(0, 30)}`);
    latexToWordModel(s);
  }
});

test("performance: a 1 MB document renders in about a second", () => {
  const base = newLatexDocument("article");
  const head = base.slice(0, base.indexOf("\\section{Introduction}"));
  const body = base.slice(base.indexOf("\\section{Introduction}"), base.indexOf("\\end{document}"));
  let big = head;
  let n = 0;
  while (big.length < 1_000_000) big += body.replace(/\\(label|ref|eqref)\{([^}]+)\}/g, (m, c, k) => `\\${c}{${k}-${n}}`), n++;
  big += "\\end{document}\n";
  const t0 = Date.now();
  const r = latexToHtml(big);
  const ms = Date.now() - t0;
  assert.ok(big.length >= 1_000_000);
  assert.ok(ms < 1500, `took ${ms} ms`);
  assert.equal(r.errors.filter((e) => e.severity === "error").length, 0);
  assert.equal(r.outline.length, n * 3);
});

test("viewer: TeX log excerpts become problems with line numbers", async () => {
  const { parseTexLog } = await import("../viewer/latex.js");
  const p = parseTexLog("This is pdfTeX\n! Undefined control sequence.\nl.12 Some \oops\n\n./main.tex:30: Missing $ inserted.");
  assert.deepEqual(p.map((e) => [e.line, e.severity, e.source]), [[12, "error", "PDF"], [30, "error", "PDF"]]);
  assert.match(p[0].message, /Undefined control sequence\. \(near “Some \oops”\)/);
  assert.equal(parseTexLog("PDF output needs a LaTeX engine")[0].line, 1);
});
