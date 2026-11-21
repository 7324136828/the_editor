/**
 * Spreadsheet editor view: virtualized grid, selection, in-cell editing,
 * clipboard, undo/redo, formatting toolbar, row/column operations, find,
 * sheet tabs and formula bar.
 *
 * Operates on the excel model returned by getModel() (see lib/model-excel.js).
 * Every listener is registered with an AbortSignal, so destroy() removes all
 * of them (including those on the shared #name-box/#formula-bar/#sheet-tabs/
 * #btn-add-sheet elements) and removes every element this view created.
 *
 * Focus model: a <textarea> (".xv-editor") always sits on the active cell.
 * When the grid has focus that textarea has focus; in navigation mode it is
 * invisible and empty, so typing (including IME composition) lands in it and
 * starts an edit that replaces the cell, exactly like a spreadsheet.
 */

import {
  addrToRC, rcToAddr, indexToCol, parseRange, MAX_ROWS, MAX_COLS,
  formatCell, cellRaw, cellFromInput, cellText, cellAlign, cellStyleCss,
  csvParse, csvSerialize, escapeHtmlText, charsToPx, pxToChars, renderSheetHtml,
  validateSheetName, sanitizeSheetName, isDateFormat, isDateOnlyOrDateTimeFormat,
  isoToSerial, serialToIso, parseCellInput, cellIsEmpty, ymdToSerial, formatGeneral,
  rangeToRef, currentRegion, detectHeaderRow, sortRange, mergesIn, filterText, filterValues,
  filterHiddenRows, FILTER_OPS, fillSeries, CHART_TYPES, chartSvg, sortKeyOf,
} from "../lib/model-excel.js";
import {
  recalcBook, invalidateSheet, translateFormula, normalizeFormula,
  formulaAutoFormat, shiftCells, renameSheet, deleteSheet, sheetBounds,
  FUNCTION_INFO, functionArgIndex,
} from "../lib/formula.js";

const ROW_H = 24;
const HEAD_H = 24;
const DEFAULT_COL_PX = 88;
const MAX_UNDO = 200;
const BIG_SNAPSHOT_CELLS = 300000;
const GRID_FONT_FAMILY = 'Calibri, Carlito, "Segoe UI", Arial, sans-serif';
const GRID_FONT_SIZE = 14;

const NUMBER_FORMATS = [
  ["General", ""],
  ["Number (1234.56)", "0.00"],
  ["Thousands (1,235)", "#,##0"],
  ["Thousands (1,234.56)", "#,##0.00"],
  ["Percent (12%)", "0%"],
  ["Percent (12.34%)", "0.00%"],
  ["Currency ($1,234.56)", '"$"#,##0.00'],
  ["Currency (€1,234.56)", '"€"#,##0.00'],
  ["Scientific (1.23E+03)", "0.00E+00"],
  ["Date (2024-01-31)", "yyyy-mm-dd"],
  ["Date (1/31/2024)", "m/d/yyyy"],
  ["Date (31-Jan-24)", "d-mmm-yy"],
  ["Time (13:30)", "h:mm"],
  ["Time (1:30 PM)", "h:mm AM/PM"],
  ["Date & time", "yyyy-mm-dd hh:mm"],
  ["Text (@)", "@"],
];

const clone = (v) => (v === undefined ? undefined : JSON.parse(JSON.stringify(v)));
const sameCell = (a, b) => JSON.stringify(a ?? null) === JSON.stringify(b ?? null);
const clamp = (v, lo, hi) => Math.max(lo, Math.min(hi, v));

function svgIcon(path, title) {
  return `<svg viewBox="0 0 20 20" aria-hidden="true" focusable="false"><path d="${path}" fill="currentColor"/></svg><span class="xv-sr">${title}</span>`;
}

const ICONS = {
  undo: "M7 4 2 8.5 7 13v-3h5a3.5 3.5 0 0 1 0 7H9v2h3a5.5 5.5 0 0 0 0-11H7V4z",
  redo: "M13 4l5 4.5-5 4.5v-3H8a3.5 3.5 0 0 0 0 7h3v2H8a5.5 5.5 0 0 1 0-11h5V4z",
  left: "M3 4h14v2H3zm0 4h9v2H3zm0 4h14v2H3zm0 4h9v2H3z",
  center: "M3 4h14v2H3zm3 4h8v2H6zm-3 4h14v2H3zm3 4h8v2H6z",
  right: "M3 4h14v2H3zm5 4h9v2H8zm-5 4h14v2H3zm5 4h9v2H8z",
  find: "M8.5 3a5.5 5.5 0 0 1 4.4 8.8l4 4-1.4 1.4-4-4A5.5 5.5 0 1 1 8.5 3zm0 2a3.5 3.5 0 1 0 0 7 3.5 3.5 0 0 0 0-7z",
  clear: "M4 15h7v2H4zM13.6 3.6l2.8 2.8-6 6H7.6V9.6l6-6zM15 12l1.4-1.4 1.6 1.6 1.6-1.6L21 12l-1.6 1.6L21 15.2l-1.4 1.4-1.6-1.6-1.6 1.6L15 15.2l1.6-1.6z",
  sortAsc: "M5 3h2v10h2.5L6 17.5 2.5 13H5zM11 4h4.5v1.6l-2.6 3.4h2.6V10.5H11V8.9l2.6-3.4H11zm.2 8.5h1.8l2.5 5.5h-1.6l-.45-1.1h-2.2l-.45 1.1h-1.6zm.9 1.9-.6 1.5h1.2z",
  sortDesc: "M5 3h2v10h2.5L6 17.5 2.5 13H5zM11.2 3h1.8l2.5 5.5h-1.6l-.45-1.1h-2.2l-.45 1.1h-1.6zm.9 1.9-.6 1.5h1.2zM11 11.5h4.5v1.6l-2.6 3.4h2.6V18H11v-1.6l2.6-3.4H11z",
  filter: "M2.5 4h15l-5.8 6.8V16l-3.4 1.6v-6.8z",
  chart: "M3 16h14v1.5H3zM4.5 9h2.5v6H4.5zm4.2-4h2.5v10H8.7zm4.3 3h2.5v7H13z",
  funnel: "M3 4h14l-5.5 6.5V16l-3 1.5v-7z",
  caret: "M5.5 8h9L10 13z",
};

export class ExcelView {
  /**
   * @param {object} opts
   * @param {HTMLElement} opts.root      container (.excel-view)
   * @param {() => object} opts.getModel returns current excel model
   * @param {(msg:string)=>void} opts.onStatus status bar message
   * @param {(dirty:boolean)=>void} opts.onDirty dirty flag changed
   * @param {() => void} opts.onStats    refresh status stats
   */
  /**
   * confirm (optional): async (message, {title, ok}) => boolean, used for
   * destructive questions so the host can show its own dialog; defaults to
   * window.confirm.
   */
  constructor({ root, getModel, onStatus, onDirty, onStats, confirm }) {
    this.root = root;
    this.confirm = confirm || ((msg) => Promise.resolve(this.win.confirm(msg)));
    this.getModel = getModel;
    this.onStatus = onStatus || (() => {});
    this.onDirty = onDirty || (() => {});
    this.onStats = onStats || (() => {});
    this.doc = root.ownerDocument;
    this.win = this.doc.defaultView || window;
    this.ac = new AbortController();

    this.cur = { r: 0, c: 0 };
    this.anchor = { r: 0, c: 0 };
    this.focusCell = { r: 0, c: 0 };
    this.selKind = "cells"; // cells | rows | cols | all
    this.editing = null;    // {r, c, mode:"enter"|"edit"}
    this.fbDirty = false;
    this.undoStack = [];
    this.redoStack = [];
    this.clip = null;       // internal clipboard
    this.nrows = 200;
    this.ncols = 30;
    this._cx = [0];
    this._win = null;
    this._modelRef = null;
    this._sheetRef = null;
    this._used = { r2: -1, c2: -1 };
    this._measureCache = new Map();
    this._statsMsg = false;
    this._destroyed = false;
    this._hidden = [];             // rows hidden by the AutoFilter (sorted)
    this._hiddenSet = new Set();
    this._filterState = new WeakMap(); // sheet -> { [col]: criteria } (view state, not saved)
    this._fz = { rows: 0, cols: 0 };
    this._selChart = -1;
    this._keepChartSel = null;

    this.elTabs = this.doc.getElementById("sheet-tabs");
    this.elNameBox = this.doc.getElementById("name-box");
    this.elFormula = this.doc.getElementById("formula-bar");
    this.elAdd = this.doc.getElementById("btn-add-sheet");
    this.elBar = this.doc.getElementById("excel-bar");

    this._build();
    this._bind();
  }

  // =====================================================================
  // DOM construction
  // =====================================================================

  _build() {
    const d = this.doc;
    this.root.textContent = "";
    this.root.classList.add("xv-root");
    const wrap = d.createElement("div");
    wrap.className = "xv-wrap";
    wrap.innerHTML = `
      <div class="xv-corner" title="Select all (Ctrl+A)"></div>
      <div class="xv-colhdr"><div class="xv-colhdr-inner"></div><div class="xv-colhdr-fz"></div></div>
      <div class="xv-rowhdr"><div class="xv-rowhdr-inner"></div><div class="xv-rowhdr-fz"></div></div>
      <div class="xv-scroll">
        <div class="xv-canvas">
          <div class="xv-fzl"><div class="xv-fz-in"></div></div>
          <div class="xv-fzt"><div class="xv-fz-in"></div><div class="xv-fzc"><div class="xv-fz-in"></div></div></div>
          <div class="xv-lines"></div>
          <div class="xv-cells"></div>
          <div class="xv-charts"></div>
          <div class="xv-selrange" hidden></div>
          <div class="xv-copyrange" hidden></div>
          <div class="xv-fillprev" hidden></div>
          <div class="xv-cursor"></div>
          <div class="xv-fillh" title="Drag to fill a series or copy (hold Ctrl to switch). Double-click to fill down." hidden></div>
          <textarea class="xv-editor" spellcheck="false" autocomplete="off" autocapitalize="off"
            aria-label="Cell editor" rows="1" wrap="off"></textarea>
        </div>
      </div>
      <div class="xv-find" role="dialog" aria-label="Find and replace" hidden>
        <div class="xv-find-row">
          <label class="xv-find-lbl" for="xv-find-input">Find</label>
          <input class="xv-find-input" id="xv-find-input" placeholder="Find what" spellcheck="false" aria-label="Find what">
          <span class="xv-find-count" aria-live="polite"></span>
          <button type="button" class="xv-find-prev" title="Find previous (Shift+Enter)" aria-label="Find previous">&#8593;</button>
          <button type="button" class="xv-find-next" title="Find next (Enter)" aria-label="Find next">&#8595;</button>
          <button type="button" class="xv-find-more" title="Show replace and options (Ctrl+H)" aria-label="Replace and options" aria-expanded="false">&#8942;</button>
          <button type="button" class="xv-find-close" title="Close (Esc)" aria-label="Close find">&#215;</button>
        </div>
        <div class="xv-find-row xv-find-rep">
          <label class="xv-find-lbl" for="xv-rep-input">Replace</label>
          <input class="xv-rep-input" id="xv-rep-input" placeholder="Replace with" spellcheck="false" aria-label="Replace with">
          <button type="button" class="xv-btn xv-rep-one" title="Replace the selected match and find the next one">Replace</button>
          <button type="button" class="xv-btn xv-rep-all" title="Replace every match">Replace all</button>
        </div>
        <div class="xv-find-row xv-find-opts">
          <label title="Search the active sheet or every sheet">Within
            <select class="xv-find-within" aria-label="Search within"><option value="sheet">Sheet</option><option value="book">Workbook</option></select></label>
          <label title="Search what cells show (values) or what they contain (formulas)">Look in
            <select class="xv-find-lookin" aria-label="Look in"><option value="values">Values</option><option value="formulas">Formulas</option></select></label>
          <label title="Upper and lower case must match"><input type="checkbox" class="xv-find-case"> Match case</label>
          <label title="The whole cell must equal the search text"><input type="checkbox" class="xv-find-whole"> Entire cell</label>
        </div>
      </div>
      <div class="xv-menu" role="menu" hidden></div>
      <div class="xv-notice" role="status" aria-live="polite" hidden></div>`;
    this.root.appendChild(wrap);
    const print = d.createElement("div");
    print.className = "xv-print";
    this.root.appendChild(print);

    const q = (s) => wrap.querySelector(s);
    this.elWrap = wrap;
    this.elPrint = print;
    this.elCorner = q(".xv-corner");
    this.elColHdr = q(".xv-colhdr");
    this.elColInner = q(".xv-colhdr-inner");
    this.elRowHdr = q(".xv-rowhdr");
    this.elRowInner = q(".xv-rowhdr-inner");
    this.elScroll = q(".xv-scroll");
    this.elGrid = this.elScroll; // backwards-compatible name
    this.elCanvas = q(".xv-canvas");
    this.elLines = q(".xv-lines");
    this.elCells = q(".xv-cells");
    this.elSel = q(".xv-selrange");
    this.elCopy = q(".xv-copyrange");
    this.elCursor = q(".xv-cursor");
    this.elEditor = q(".xv-editor");
    this.elFind = q(".xv-find");
    this.elFindInput = q(".xv-find-input");
    this.elFindCount = q(".xv-find-count");
    this.elRepInput = q(".xv-rep-input");
    this.elMenu = q(".xv-menu");
    this.elNotice = q(".xv-notice");
    this.elColFz = q(".xv-colhdr-fz");
    this.elRowFz = q(".xv-rowhdr-fz");
    this.elFzL = q(".xv-fzl");
    this.elFzT = q(".xv-fzt");
    this.elFzC = q(".xv-fzc");
    this.elFzLIn = this.elFzL.firstElementChild;
    this.elFzTIn = this.elFzT.firstElementChild;
    this.elFzCIn = this.elFzC.firstElementChild;
    this.elCharts = q(".xv-charts");
    this.elFillH = q(".xv-fillh");
    this.elFillPrev = q(".xv-fillprev");
    this._popups = []; // elements appended outside the root (removed in destroy)

    this._buildToolbar();
  }

  _buildToolbar() {
    if (!this.elBar) return;
    const d = this.doc;
    const tb = d.createElement("div");
    tb.className = "xv-toolbar";
    tb.setAttribute("role", "toolbar");
    tb.setAttribute("aria-label", "Spreadsheet formatting");
    const fmtOptions = NUMBER_FORMATS.map(([label, z]) => `<option value="${escapeHtmlText(z)}">${escapeHtmlText(label)}</option>`).join("");
    tb.innerHTML = `
      <button class="fb xv-tb" data-act="undo" title="Undo (Ctrl+Z)" aria-label="Undo">${svgIcon(ICONS.undo, "Undo")}</button>
      <button class="fb xv-tb" data-act="redo" title="Redo (Ctrl+Y)" aria-label="Redo">${svgIcon(ICONS.redo, "Redo")}</button>
      <span class="sep" aria-hidden="true"></span>
      <button class="fb xv-tb" data-act="b" title="Bold (Ctrl+B)" aria-label="Bold" aria-pressed="false"><b>B</b></button>
      <button class="fb xv-tb" data-act="i" title="Italic (Ctrl+I)" aria-label="Italic" aria-pressed="false"><i>I</i></button>
      <button class="fb xv-tb" data-act="u" title="Underline (Ctrl+U)" aria-label="Underline" aria-pressed="false"><u>U</u></button>
      <label class="fb xv-tb xv-color" title="Text color"><span class="xv-color-a">A</span><span class="xv-swatch xv-swatch-color"></span><input type="color" data-act="color" value="#c00000" aria-label="Text color"></label>
      <label class="fb xv-tb xv-color" title="Fill color"><span class="xv-color-fill">&#9632;</span><span class="xv-swatch xv-swatch-bg"></span><input type="color" data-act="bg" value="#fff2cc" aria-label="Fill color"></label>
      <span class="sep" aria-hidden="true"></span>
      <button class="fb xv-tb" data-act="align-left" title="Align left" aria-label="Align left" aria-pressed="false">${svgIcon(ICONS.left, "Align left")}</button>
      <button class="fb xv-tb" data-act="align-center" title="Center" aria-label="Center" aria-pressed="false">${svgIcon(ICONS.center, "Center")}</button>
      <button class="fb xv-tb" data-act="align-right" title="Align right" aria-label="Align right" aria-pressed="false">${svgIcon(ICONS.right, "Align right")}</button>
      <select class="fb-select xv-numfmt" data-act="numfmt" title="Number format" aria-label="Number format">${fmtOptions}</select>
      <button class="fb xv-tb" data-act="clearfmt" title="Clear formatting" aria-label="Clear formatting">${svgIcon(ICONS.clear, "Clear formatting")}</button>
      <span class="sep" aria-hidden="true"></span>
      <button class="fb xv-tb small" data-act="ins-row" title="Insert row(s) above">+Row</button>
      <button class="fb xv-tb small" data-act="del-row" title="Delete selected row(s)">&#8722;Row</button>
      <button class="fb xv-tb small" data-act="ins-col" title="Insert column(s) to the left">+Col</button>
      <button class="fb xv-tb small" data-act="del-col" title="Delete selected column(s)">&#8722;Col</button>
      <span class="sep" aria-hidden="true"></span>
      <button class="fb xv-tb" data-act="find" title="Find and replace (Ctrl+F / Ctrl+H)" aria-label="Find and replace">${svgIcon(ICONS.find, "Find and replace")}</button>
      <span class="sep" aria-hidden="true"></span>
      <button class="fb xv-tb small" data-act="sort-asc" title="Sort A to Z / smallest to largest by the current column" aria-label="Sort A to Z">${svgIcon(ICONS.sortAsc, "Sort A to Z")}</button>
      <button class="fb xv-tb small" data-act="sort-desc" title="Sort Z to A / largest to smallest by the current column" aria-label="Sort Z to A">${svgIcon(ICONS.sortDesc, "Sort Z to A")}</button>
      <button class="fb xv-tb small" data-act="sort-dialog" title="Custom sort: several columns, headers (Ctrl+Shift+S)">Sort&#8230;</button>
      <button class="fb xv-tb small" data-act="filter" title="Filter: add drop-down buttons to the headers of the current table (Ctrl+Shift+L)" aria-pressed="false">${svgIcon(ICONS.filter, "Filter")}<span class="xv-tblabel">Filter</span></button>
      <button class="fb xv-tb small" data-act="freeze" title="Freeze rows or columns so they stay visible while scrolling" aria-haspopup="menu">Freeze &#9662;</button>
      <button class="fb xv-tb small" data-act="chart" title="Insert a chart from the selected data (header row + first column as labels)">${svgIcon(ICONS.chart, "Insert chart")}<span class="xv-tblabel">Chart</span></button>`;
    this.elBar.insertBefore(tb, this.elBar.firstChild);
    this.elToolbar = tb;
    this.elNumFmt = tb.querySelector(".xv-numfmt");
  }

  // =====================================================================
  // Event binding (all with the abort signal)
  // =====================================================================

  _bind() {
    const sig = { signal: this.ac.signal };
    const on = (el, type, fn, opts = {}) => el && el.addEventListener(type, fn, { ...opts, ...sig });

    // editor / keyboard sink
    on(this.elEditor, "keydown", (e) => this._onKeyDown(e));
    on(this.elEditor, "input", () => this._onEditorInput());
    on(this.elEditor, "keyup", (e) => { if (this.editing && /^(Arrow(Left|Right)|Home|End)$/.test(e.key)) this._acUpdate(this.elEditor); });
    on(this.elEditor, "compositionstart", () => { if (!this.editing) this._beginEdit("enter", null); });
    on(this.elEditor, "copy", (e) => this._onCopy(e, false));
    on(this.elEditor, "cut", (e) => this._onCopy(e, true));
    on(this.elEditor, "paste", (e) => this._onPaste(e));
    on(this.elEditor, "blur", () => this._onEditorBlur());

    // grid mouse
    on(this.elScroll, "scroll", () => this._onScroll(), { passive: true });
    on(this.elScroll, "mousedown", (e) => this._onCellMouseDown(e));
    on(this.elScroll, "dblclick", (e) => this._onCellDblClick(e));
    on(this.elScroll, "contextmenu", (e) => this._onContextMenu(e, "cell"));
    on(this.elColHdr, "mousedown", (e) => this._onHeaderMouseDown(e, "col"));
    on(this.elRowHdr, "mousedown", (e) => this._onHeaderMouseDown(e, "row"));
    on(this.elColHdr, "dblclick", (e) => this._onHeaderDblClick(e));
    on(this.elColHdr, "contextmenu", (e) => this._onContextMenu(e, "col"));
    on(this.elRowHdr, "contextmenu", (e) => this._onContextMenu(e, "row"));
    on(this.elColHdr, "wheel", (e) => { this.elScroll.scrollLeft += e.deltaX || e.deltaY; e.preventDefault(); }, { passive: false });
    on(this.elRowHdr, "wheel", (e) => { this.elScroll.scrollTop += e.deltaY; e.preventDefault(); }, { passive: false });
    on(this.elCorner, "mousedown", (e) => { e.preventDefault(); this._commitAll(); this.selectAll(); this._focusGrid(); });
    on(this.win, "mousemove", (e) => this._onDragMove(e));
    on(this.win, "mouseup", (e) => this._endDrag(e));
    on(this.doc, "mousedown", (e) => {
      if (!this.elMenu.hidden && !this.elMenu.contains(e.target)) this._hideMenu();
      if (this.elFilterPanel && !this.elFilterPanel.contains(e.target) && !(e.target.closest && e.target.closest(".xv-fbtn"))) this._closeFilterPanel();
      if (this.elAc && !this.elAc.contains(e.target)) this._hideAutocomplete();
    }, { capture: true });
    on(this.elCharts, "keydown", (e) => this._onChartKey(e));
    on(this.elCharts, "dblclick", (e) => {
      const el = e.target.closest(".xv-chart");
      if (el) { e.stopPropagation(); this._openChartPanel(+el.dataset.i); }
    });
    on(this.win, "blur", () => this._hideMenu());

    // find & replace panel
    on(this.elFindInput, "input", () => this._runFind(true));
    on(this.elFind, "change", () => this._runFind(false));
    on(this.elFind, "keydown", (e) => {
      const mod = e.ctrlKey || e.metaKey;
      if (mod && !e.altKey && (e.key === "f" || e.key === "h" || e.key === "F" || e.key === "H")) {
        e.preventDefault();
        this.openFind(e.key.toLowerCase() === "h" || this.elFind.classList.contains("xv-find-replace"));
        return;
      }
      if (e.key === "Escape") { e.preventDefault(); this.closeFind(); return; }
      if (e.key !== "Enter" || e.target.tagName === "BUTTON" || e.target.tagName === "SELECT") return;
      e.preventDefault();
      if (e.target === this.elRepInput) { if (mod) this._replaceAll(); else this._replaceOne(); }
      else this._findStep(e.shiftKey ? -1 : 1);
    });
    on(this.elFind.querySelector(".xv-find-prev"), "click", () => this._findStep(-1));
    on(this.elFind.querySelector(".xv-find-next"), "click", () => this._findStep(1));
    on(this.elFind.querySelector(".xv-find-close"), "click", () => this.closeFind());
    on(this.elFind.querySelector(".xv-find-more"), "click", () => {
      this._setFindMode(!this.elFind.classList.contains("xv-find-replace"));
      (this.elFind.classList.contains("xv-find-replace") ? this.elRepInput : this.elFindInput).focus();
    });
    on(this.elFind.querySelector(".xv-rep-one"), "click", () => this._replaceOne());
    on(this.elFind.querySelector(".xv-rep-all"), "click", () => this._replaceAll());

    // menu
    on(this.elMenu, "click", (e) => this._onMenuClick(e));
    on(this.elMenu, "keydown", (e) => { if (e.key === "Escape") { this._hideMenu(); this._focusGrid(); } });

    // shared chrome: add sheet, name box, formula bar, tabs
    on(this.elAdd, "click", () => this.addSheet());
    on(this.elNameBox, "focus", () => this.elNameBox.select());
    on(this.elNameBox, "keydown", (e) => {
      if (e.key === "Enter") { e.preventDefault(); this._gotoRef(this.elNameBox.value); }
      else if (e.key === "Escape") { e.preventDefault(); this._syncBar(); this._focusGrid(); }
    });
    on(this.elFormula, "focus", () => {
      if (this.editing) {
        // Continue the in-progress edit in the formula bar.
        const text = this.elEditor.value;
        this._endEditUi();
        this.elFormula.value = text;
        this.fbDirty = text !== cellRaw(this._curCell());
      }
    });
    on(this.elFormula, "input", () => { this.fbDirty = true; this._acUpdate(this.elFormula); });
    on(this.elFormula, "keyup", (e) => { if (/^(Arrow(Left|Right)|Home|End)$/.test(e.key)) this._acUpdate(this.elFormula); });
    on(this.elFormula, "click", () => this._acUpdate(this.elFormula));
    on(this.elFormula, "keydown", (e) => this._onFormulaKey(e));
    on(this.elFormula, "blur", () => { this._hideAutocomplete(); if (this.fbDirty) this._commitFormulaBar(0, 0, false); });
    on(this.elTabs, "click", (e) => this._onTabClick(e));
    on(this.elTabs, "dblclick", (e) => this._onTabDblClick(e));
    on(this.elTabs, "contextmenu", (e) => this._onTabContextMenu(e));

    // toolbar
    if (this.elToolbar) {
      on(this.elToolbar, "mousedown", (e) => {
        // keep grid focus/selection when clicking buttons (not for inputs/selects)
        if (e.target.closest("button")) e.preventDefault();
      });
      on(this.elToolbar, "click", (e) => {
        const b = e.target.closest("button[data-act]");
        if (b) this._toolbarAction(b.dataset.act);
      });
      on(this.elToolbar, "change", (e) => {
        const t = e.target;
        if (t.dataset.act === "numfmt") { this._applyNumberFormat(t.value); this._focusGrid(); }
        else if (t.dataset.act === "color" || t.dataset.act === "bg") { this._applyStyle({ [t.dataset.act]: t.value.toUpperCase() }); this._focusGrid(); }
      });
    }

    // printing renders a static table of the active sheet
    on(this.win, "beforeprint", () => this._preparePrint());
    on(this.win, "afterprint", () => { this.elPrint.textContent = ""; });

    if (typeof ResizeObserver !== "undefined") {
      this._ro = new ResizeObserver(() => this._renderGrid());
      this._ro.observe(this.elScroll);
    }
  }

  destroy() {
    if (this._destroyed) return;
    this._destroyed = true;
    try { this._cancelEdit(); } catch { /* ignore */ }
    this.ac.abort();
    if (this._ro) this._ro.disconnect();
    if (this._dragTimer) clearInterval(this._dragTimer);
    if (this._raf) cancelAnimationFrame(this._raf);
    if (this._chartRaf) cancelAnimationFrame(this._chartRaf);
    clearTimeout(this._noticeTimer);
    for (const el of this._popups || []) el.remove();
    this._popups = [];
    if (this.elToolbar) this.elToolbar.remove();
    if (this.elTabs) this.elTabs.textContent = "";
    this.root.textContent = "";
    this.root.classList.remove("xv-root");
  }

  // =====================================================================
  // Model access
  // =====================================================================

  _model() { return this.getModel(); }
  _si() {
    const m = this._model();
    return clamp(m.active || 0, 0, Math.max(0, m.sheets.length - 1));
  }
  _sheet() {
    const m = this._model();
    return m && m.sheets ? m.sheets[this._si()] : null;
  }
  _cellAt(r, c) {
    const s = this._sheet();
    return s ? s.cells[rcToAddr(r, c)] : undefined;
  }
  _curCell() {
    const m = this._mergeAt(this.cur.r, this.cur.c);
    return m ? this._cellAt(m.r1, m.c1) : this._cellAt(this.cur.r, this.cur.c);
  }

  // =====================================================================
  // Rendering
  // =====================================================================

  render() {
    const m = this._model();
    if (!m) return;
    if (!m.sheets || !m.sheets.length) m.sheets = [{ name: "Sheet1", cells: {} }];
    let fresh = false;
    if (this._modelRef !== m) {
      this._modelRef = m;
      this.undoStack = [];
      this.redoStack = [];
      this.clip = null;
      this._recalc(null, true);
      this._resetSheetState();
      fresh = true;
    } else if (this._sheetRef !== this._sheet()) {
      this._resetSheetState();
    }
    this._renderTabs();
    this._renderGrid(true);
    this._syncBar();
    this._updateToolbar();
    this.onStats();
    // A freshly opened workbook gets keyboard focus so arrows/typing work at once.
    const ae = this.doc.activeElement;
    if (fresh && (!ae || ae === this.doc.body) && !this.root.hidden) this._focusGrid();
  }

  _resetSheetState() {
    const sheet = this._sheet();
    this._sheetRef = sheet;
    if (!sheet.cells) sheet.cells = {};
    this._used = sheetBounds(sheet);
    this._mergeList = (sheet.merges || []).map(parseRange).filter(Boolean);
    this.cur = { r: 0, c: 0 };
    this.anchor = { r: 0, c: 0 };
    this.focusCell = { r: 0, c: 0 };
    this.selKind = "cells";
    this.nrows = clamp(Math.max(200, this._used.r2 + 60), 1, MAX_ROWS);
    this.ncols = clamp(Math.max(30, this._used.c2 + 12), 1, MAX_COLS);
    this._selChart = this._keepChartSel != null && this._keepChartSel >= 0 ? this._keepChartSel : -1;
    this._keepChartSel = null;
    this._closeFilterPanel();
    this._computeHidden();
    this._readFreeze();
    this._layoutCols();
    this.elScroll.scrollTop = 0;
    this.elScroll.scrollLeft = 0;
    this._win = null;
    this._renderCharts();
    if (!this.elFind.hidden) this._runFind(false);
  }

  _layoutCols() {
    const sheet = this._sheet();
    const widths = (sheet && sheet.colWidths) || {};
    const cx = new Array(this.ncols + 1);
    cx[0] = 0;
    for (let c = 0; c < this.ncols; c++) {
      const w = widths[indexToCol(c)];
      cx[c + 1] = cx[c] + (w != null ? charsToPx(w) : DEFAULT_COL_PX);
    }
    this._cx = cx;
    this._colLetters = [];
    const totalH = this._ry(this.nrows);
    this.elCanvas.style.width = `${cx[this.ncols]}px`;
    this.elCanvas.style.height = `${totalH}px`;
    const digits = String(this.nrows).length;
    const rh = Math.max(40, digits * 8 + 16);
    this.elWrap.style.setProperty("--xv-rh", `${rh}px`);
    this.elWrap.style.setProperty("--xv-hh", `${HEAD_H}px`);
    this._layoutPanes();
  }

  // ---- rows: filtered (hidden) rows have zero height ----

  /** Recompute the rows hidden by the sheet's AutoFilter criteria. */
  _computeHidden() {
    const sheet = this._sheet();
    const st = sheet && this._filterState.get(sheet);
    const list = sheet && sheet.autoFilter && st && Object.keys(st).length ? filterHiddenRows(sheet, sheet.autoFilter, st) : [];
    this._hidden = list;
    this._hiddenSet = new Set(list);
  }

  _isHidden(r) { return this._hiddenSet.size > 0 && this._hiddenSet.has(r); }

  /** Number of hidden rows above row r. */
  _hiddenBefore(r) {
    const h = this._hidden;
    if (!h.length || r <= h[0]) return 0;
    let lo = 0, hi = h.length;
    while (lo < hi) { const mid = (lo + hi) >> 1; if (h[mid] < r) lo = mid + 1; else hi = mid; }
    return lo;
  }

  /** Top of row r in canvas px. */
  _ry(r) { return (r - this._hiddenBefore(r)) * ROW_H; }
  _rh(r) { return this._isHidden(r) ? 0 : ROW_H; }

  /** Visible row at canvas y. */
  _rowAtY(y) {
    const v = Math.max(0, Math.floor(y / ROW_H));
    if (!this._hidden.length) return clamp(v, 0, this.nrows - 1);
    let lo = 0, hi = this.nrows - 1;
    while (lo < hi) { // smallest r with (visible rows in 0..r) > v
      const mid = (lo + hi) >> 1;
      if (mid + 1 - this._hiddenBefore(mid + 1) > v) hi = mid; else lo = mid + 1;
    }
    return lo;
  }

  /** Move |n| visible rows from r (sign = direction); stays inside the sheet. */
  _stepRow(r, n) {
    const d = n < 0 ? -1 : 1;
    let left = Math.abs(n);
    let x = r;
    while (left > 0) {
      let y = x + d;
      while (y >= 0 && y < MAX_ROWS && this._isHidden(y)) y += d;
      if (y < 0 || y >= MAX_ROWS) break;
      x = y;
      left--;
    }
    return x;
  }

  _visibleRow(r) {
    if (!this._isHidden(r)) return r;
    const down = this._stepRow(r, 1);
    return this._isHidden(down) ? this._stepRow(r, -1) : down;
  }

  // ---- frozen panes ----

  _readFreeze() {
    const f = this._sheet() && this._sheet().freeze;
    this._fz = { rows: clamp(Math.floor((f && f.rows) || 0), 0, 10000), cols: clamp(Math.floor((f && f.cols) || 0), 0, 1000) };
  }

  /** Size of the frozen area in px: {fh, fw}. */
  _fzSize() {
    const fz = this._fz || { rows: 0, cols: 0 };
    return { fh: fz.rows ? this._ry(fz.rows) : 0, fw: fz.cols ? this._cx[Math.min(fz.cols, this.ncols)] : 0 };
  }

  _layoutPanes() {
    if (!this.elFzL) return;
    const { fh, fw } = this._fzSize();
    const totalH = this._ry(this.nrows), totalW = this._cx[this.ncols];
    const on = fh > 0 || fw > 0;
    this.elWrap.classList.toggle("xv-frozen", on);
    this.elFzL.style.width = `${fw}px`;
    this.elFzL.style.height = `${on ? totalH : 0}px`;
    this.elFzL.hidden = !fw;
    this.elFzT.style.width = `${totalW}px`;
    this.elFzT.style.height = `${fh}px`;
    this.elFzT.style.marginTop = fw ? `${-totalH}px` : "0px";
    this.elFzT.hidden = !fh;
    this.elFzC.style.width = `${fw}px`;
    this.elFzC.style.height = `${fh}px`;
    this.elFzC.hidden = !(fh && fw);
    this.elColFz.style.width = `${fw}px`;
    this.elColFz.hidden = !fw;
    this.elRowFz.style.height = `${fh}px`;
    this.elRowFz.hidden = !fh;
  }

  _col(c) { return this._colLetters[c] || (this._colLetters[c] = indexToCol(c)); }
  _colW(c) { return this._cx[c + 1] - this._cx[c]; }

  _colAtX(x) {
    const cx = this._cx;
    let lo = 0, hi = this.ncols - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (cx[mid] <= x) lo = mid; else hi = mid - 1;
    }
    return lo;
  }

  /** Grow the virtual extent so (r, c) and a margin beyond it exist. */
  _ensureExtent(r, c) {
    let changed = false;
    if (r >= this.nrows - 20 && this.nrows < MAX_ROWS) { this.nrows = clamp(r + 200, 1, MAX_ROWS); changed = true; }
    if (c >= this.ncols - 5 && this.ncols < MAX_COLS) { this.ncols = clamp(c + 20, 1, MAX_COLS); changed = true; }
    if (changed) this._layoutCols();
    return changed;
  }

  _onScroll() {
    const sl = this.elScroll.scrollLeft, st = this.elScroll.scrollTop;
    this.elColInner.style.transform = `translateX(${-sl}px)`;
    this.elRowInner.style.transform = `translateY(${-st}px)`;
    // Extend the sheet when scrolling near the end (spreadsheets are "infinite").
    const nearBottom = st + this.elScroll.clientHeight > this._ry(this.nrows) - 10 * ROW_H;
    const nearRight = sl + this.elScroll.clientWidth > this._cx[this.ncols] - 3 * DEFAULT_COL_PX;
    if ((nearBottom && this.nrows < MAX_ROWS) || (nearRight && this.ncols < MAX_COLS)) {
      if (nearBottom) this.nrows = clamp(this.nrows + 500, 1, MAX_ROWS);
      if (nearRight) this.ncols = clamp(this.ncols + 20, 1, MAX_COLS);
      this._layoutCols();
    }
    if (this._fz && (this._fz.rows || this._fz.cols)) this._renderOverlays(); // overlays on frozen cells follow at once
    if (!this.elFilterPanel || this.elFilterPanel.hidden) { /* nothing */ } else this._closeFilterPanel();
    this._hideAutocomplete();
    if (this._raf) return;
    this._raf = requestAnimationFrame(() => {
      this._raf = 0;
      this._renderGrid(false);
    });
  }

  _visibleWindow() {
    const st = this.elScroll.scrollTop, sl = this.elScroll.scrollLeft;
    const h = this.elScroll.clientHeight || 600, w = this.elScroll.clientWidth || 1000;
    const r0 = this._stepRow(this._rowAtY(st), -3);
    const r1 = clamp(this._stepRow(this._rowAtY(st + h), 3), 0, this.nrows - 1);
    const c0 = clamp(this._colAtX(sl) - 1, 0, this.ncols - 1);
    const c1 = clamp(this._colAtX(sl + w) + 1, 0, this.ncols - 1);
    return { r0, r1, c0, c1 };
  }

  _renderGrid(force = true) {
    const sheet = this._sheet();
    if (!sheet || this._destroyed) return;
    const win = this._visibleWindow();
    const same = this._win && this._win.r0 === win.r0 && this._win.r1 === win.r1 && this._win.c0 === win.c0 && this._win.c1 === win.c1;
    if (same && !force) { this._renderOverlays(); return; }
    this._win = win;
    this._renderCells(sheet, win);
    this._renderHeaders(win);
    this._renderOverlays();
  }

  _measure(text, font) {
    const key = font + "\u0000" + text;
    let w = this._measureCache.get(key);
    if (w != null) return w;
    if (!this._ctx) {
      const cv = this.doc.createElement("canvas");
      this._ctx = cv.getContext && cv.getContext("2d");
    }
    if (!this._ctx) return text.length * 7;
    this._ctx.font = font;
    w = this._ctx.measureText(text).width;
    if (this._measureCache.size > 20000) this._measureCache.clear();
    this._measureCache.set(key, w);
    return w;
  }

  _fontFor(cell) {
    const s = cell && cell.s;
    return `${s && s.i ? "italic " : ""}${s && s.b ? "700 " : ""}${GRID_FONT_SIZE}px ${GRID_FONT_FAMILY}`;
  }

  _mergeAt(r, c) {
    const list = this._mergeList;
    if (!list || !list.length) return null;
    for (const m of list) if (r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2) return m;
    return null;
  }

  _cellHtml(cell, left, top, width, height, opts = {}) {
    const { text, color } = formatCell(cell);
    const align = cellAlign(cell);
    let css = cellStyleCss(cell, null);
    if (color) css += `color:${color};`;
    let cls = "xv-c";
    if (align === "right") cls += " xv-r";
    else if (align === "center") cls += " xv-ctr";
    if (cell.t === "e") cls += " xv-err";
    if (cell.s && cell.s.wrap) cls += " xv-wrapc";
    let shown = text;
    let w = width;
    if (shown && !(cell.s && cell.s.wrap)) {
      const tw = this._measure(shown, this._fontFor(cell));
      if (tw > width - 8) {
        const general = cell.t === "n" && typeof cell.v === "number" && Number.isFinite(cell.v) && (!cell.z || /^general$/i.test(cell.z));
        const shorter = general ? this._fitGeneral(cell.v, width - 8, this._fontFor(cell)) : null;
        if (shorter != null) {
          // Like Excel, a General number shows fewer decimals in a narrow column.
          shown = shorter;
        } else if (cell.t === "n" || cell.t === "d") {
          // Excel shows #### when a number does not fit.
          const hw = this._measure("#", this._fontFor(cell));
          shown = "#".repeat(Math.max(1, Math.floor((width - 8) / hw)));
        } else if (align === "left" && opts.spill) {
          w = opts.spill(tw + 10);
          if (w > width) cls += " xv-spill";
        }
      }
    }
    return `<div class="${cls}" style="left:${left}px;top:${top}px;width:${w}px;height:${height}px;${css}">${escapeHtmlText(shown)}</div>`;
  }

  /** Fewer significant digits of a General non-integer that fit `room` px, or null. */
  _fitGeneral(v, room, font) {
    if (Number.isInteger(v)) return null;
    const a = Math.abs(v);
    const intDigits = a >= 1 ? Math.floor(Math.log10(a)) + 1 : 1;
    for (let d = 10; d >= intDigits; d--) {
      const t = formatGeneral(v, d);
      if (this._measure(t, font) <= room) return t;
    }
    return null;
  }

  /** Cells + vertical grid lines of rows r0..r1 x cols c0..c1 as HTML. */
  _cellsHtml(sheet, r0, r1, c0, c1) {
    const cells = sheet.cells;
    const cx = this._cx;
    let html = "";
    // merges intersecting the window
    const covered = new Set();
    for (const m of this._mergeList || []) {
      if (m.r2 < r0 || m.r1 > r1 || m.c2 < c0 || m.c1 > c1) continue;
      for (let r = Math.max(m.r1, r0); r <= Math.min(m.r2, r1); r++) {
        for (let c = Math.max(m.c1, c0); c <= Math.min(m.c2, c1); c++) covered.add(r * 16384 + c);
      }
      const cell = cells[rcToAddr(m.r1, m.c1)] || { t: "s", v: null };
      const c2 = Math.min(m.c2, this.ncols - 1);
      const top = this._ry(m.r1), height = this._ry(m.r2 + 1) - top;
      if (height <= 0) continue;
      html += this._cellHtml(cell, cx[m.c1], top, cx[c2 + 1] - cx[m.c1], height)
        .replace('class="xv-c', 'class="xv-c xv-merge');
    }
    const af = this._afRange();
    const crit = af ? this._filterState.get(sheet) || {} : null;
    for (let r = r0; r <= r1; r++) {
      if (this._isHidden(r)) continue;
      const rs = String(r + 1);
      const top = this._ry(r);
      for (let c = c0; c <= c1; c++) {
        const cell = cells[this._col(c) + rs];
        if (!cell || covered.has(r * 16384 + c)) continue;
        if (cell.v == null && !cell.f && !(cell.s && cell.s.bg)) continue;
        const spill = (need) => {
          let w = this._colW(c);
          for (let k = c + 1; k < Math.min(this.ncols, c + 40) && w < need; k++) {
            const nb = cells[this._col(k) + rs];
            if ((nb && (nb.v != null || nb.f)) || covered.has(r * 16384 + k) || this._mergeAt(r, k)) break;
            if (af && r === af.r1 && k >= af.c1 && k <= af.c2) break;
            w += this._colW(k);
          }
          return Math.min(w, Math.max(this._colW(c), need));
        };
        html += this._cellHtml(cell, cx[c], top, this._colW(c), ROW_H, { spill });
      }
      if (af && r === af.r1) {
        for (let c = Math.max(c0, af.c1); c <= Math.min(c1, af.c2); c++) {
          const active = !!crit[c];
          const name = filterText(cells[this._col(c) + rs]) || `Column ${this._col(c)}`;
          html += `<button type="button" class="xv-fbtn${active ? " xv-fbtn-on" : ""}" data-fc="${c}" tabindex="-1" style="left:${cx[c + 1] - 21}px;top:${top + 3}px"` +
            ` title="${escapeHtmlText(active ? `${name}: filtered. Click to change the filter` : `Filter or sort ${name}`)} (Alt+Down)" aria-label="Filter ${escapeHtmlText(name)}">` +
            `${active ? svgIcon(ICONS.funnel, "") : svgIcon(ICONS.caret, "")}</button>`;
        }
      }
    }
    let lines = "";
    const top = this._ry(r0), height = this._ry(r1 + 1) - top;
    for (let c = c0; c <= c1; c++) {
      lines += `<div class="xv-vl" style="left:${cx[c + 1] - 1}px;top:${top}px;height:${height}px"></div>`;
    }
    return { html, lines };
  }

  _renderCells(sheet, win) {
    const main = this._cellsHtml(sheet, win.r0, win.r1, win.c0, win.c1);
    this.elCells.innerHTML = main.html;
    this.elLines.innerHTML = main.lines;
    const fz = this._fz || { rows: 0, cols: 0 };
    if (!fz.rows && !fz.cols) {
      if (this._fzRendered) { this.elFzTIn.innerHTML = ""; this.elFzLIn.innerHTML = ""; this.elFzCIn.innerHTML = ""; this._fzRendered = false; }
      return;
    }
    this._fzRendered = true;
    const h = this.elScroll.clientHeight || 600, w = this.elScroll.clientWidth || 1000;
    const fr = Math.min(fz.rows, this._rowAtY(h) + 1), fc = Math.min(fz.cols, this._colAtX(w) + 1);
    const pane = (el, r0, r1, c0, c1) => {
      if (r1 < r0 || c1 < c0) { el.innerHTML = ""; return; }
      const p = this._cellsHtml(sheet, r0, r1, c0, c1);
      el.innerHTML = p.lines + p.html;
    };
    if (fz.rows) pane(this.elFzTIn, 0, fr - 1, Math.max(win.c0, fz.cols), win.c1);
    if (fz.cols) pane(this.elFzLIn, Math.max(win.r0, fz.rows), win.r1, 0, fc - 1);
    if (fz.rows && fz.cols) pane(this.elFzCIn, 0, fr - 1, 0, fc - 1);
  }

  _selRange() {
    const a = this.anchor, f = this.focusCell;
    let rg = { r1: Math.min(a.r, f.r), r2: Math.max(a.r, f.r), c1: Math.min(a.c, f.c), c2: Math.max(a.c, f.c) };
    // expand to cover merges that intersect the selection
    if (this._mergeList && this._mergeList.length) {
      let changed = true;
      while (changed) {
        changed = false;
        for (const m of this._mergeList) {
          if (m.r2 < rg.r1 || m.r1 > rg.r2 || m.c2 < rg.c1 || m.c1 > rg.c2) continue;
          if (m.r1 < rg.r1 || m.r2 > rg.r2 || m.c1 < rg.c1 || m.c2 > rg.c2) {
            rg = { r1: Math.min(rg.r1, m.r1), r2: Math.max(rg.r2, m.r2), c1: Math.min(rg.c1, m.c1), c2: Math.max(rg.c2, m.c2) };
            changed = true;
          }
        }
      }
    }
    return rg;
  }

  /** Selection clipped to the used area (for whole row/column selections). */
  _selClipped() {
    const rg = this._selRange();
    return {
      r1: rg.r1, c1: rg.c1,
      r2: Math.min(rg.r2, Math.max(rg.r1, this._used.r2)),
      c2: Math.min(rg.c2, Math.max(rg.c1, this._used.c2)),
    };
  }

  _rectOf(rg) {
    const cx = this._cx;
    const c2 = Math.min(rg.c2, this.ncols - 1), r2 = Math.min(rg.r2, this.nrows - 1);
    const top = this._ry(rg.r1);
    return { left: cx[rg.c1], top, width: cx[c2 + 1] - cx[rg.c1], height: this._ry(r2 + 1) - top };
  }

  /**
   * Where a range is drawn right now, in canvas px, taking frozen panes into
   * account (frozen rows/cols are pinned at the scroll position; unfrozen
   * parts hidden behind the panes are clipped). null when nothing shows.
   */
  _viewRect(rg) {
    const fz = this._fz || { rows: 0, cols: 0 };
    const r = this._rectOf(rg);
    if (!fz.rows && !fz.cols) return r;
    const { fh, fw } = this._fzSize();
    const st = this.elScroll.scrollTop, sl = this.elScroll.scrollLeft;
    let top = r.top, bottom = r.top + r.height, left = r.left, right = r.left + r.width;
    if (fz.rows) {
      top = rg.r1 < fz.rows ? st + top : Math.max(top, st + fh);
      bottom = rg.r2 < fz.rows ? st + bottom : Math.max(bottom, st + fh);
    }
    if (fz.cols) {
      left = rg.c1 < fz.cols ? sl + left : Math.max(left, sl + fw);
      right = rg.c2 < fz.cols ? sl + right : Math.max(right, sl + fw);
    }
    if (bottom - top <= 0 || right - left <= 0) return null;
    return { left, top, width: right - left, height: bottom - top };
  }

  _place(el, rect) {
    if (!rect) { el.style.visibility = "hidden"; return; }
    el.style.visibility = "";
    el.style.left = `${rect.left}px`;
    el.style.top = `${rect.top}px`;
    el.style.width = `${rect.width}px`;
    el.style.height = `${rect.height}px`;
  }

  _renderOverlays() {
    const rg = this._selRange();
    const multi = rg.r1 !== rg.r2 || rg.c1 !== rg.c2;
    const m = this._mergeAt(this.cur.r, this.cur.c);
    const curRg = m || { r1: this.cur.r, r2: this.cur.r, c1: this.cur.c, c2: this.cur.c };
    const curRect = this._viewRect(curRg);
    this._place(this.elCursor, curRect);
    // Overlays stay under floating charts, except on frozen cells (above the panes).
    const fz = this._fz || { rows: 0, cols: 0 };
    const frozenPart = (g) => !!g && ((fz.rows && g.r1 < fz.rows) || (fz.cols && g.c1 < fz.cols));
    this.elCursor.style.zIndex = frozenPart(curRg) ? "" : "3";
    this.elSel.style.zIndex = frozenPart(rg) ? "" : "2";
    this.elFillH.style.zIndex = frozenPart(rg) ? "" : "3";
    this.elCopy.style.zIndex = this.clip && frozenPart(this.clip.range) ? "" : "3";
    const isMergeOnly = m && rg.r1 === m.r1 && rg.r2 === m.r2 && rg.c1 === m.c1 && rg.c2 === m.c2;
    this.elSel.hidden = !multi || isMergeOnly;
    if (multi) this._place(this.elSel, this._viewRect(rg));
    if (this.clip && this.clip.si === this._si() && this.clip.sheetRef === this._sheet()) {
      this.elCopy.hidden = false;
      this._place(this.elCopy, this._viewRect(this.clip.range));
    } else if (!this._point) this.elCopy.hidden = true;
    if (!this.editing) {
      this._place(this.elEditor, curRect || this._rectOf(curRg));
    } else this._placeEditor();
    // fill handle at the bottom-right corner of the selection (cells only)
    const showFill = !this.editing && this.selKind === "cells" && !(this.drag && this.drag.kind === "fill");
    const selRect = multi && !isMergeOnly ? this._viewRect(rg) : curRect;
    this.elFillH.hidden = !showFill || !selRect;
    if (showFill && selRect) {
      this.elFillH.style.left = `${selRect.left + selRect.width - 4}px`;
      this.elFillH.style.top = `${selRect.top + selRect.height - 4}px`;
    }
    // header highlight
    const sel = rg;
    const hdrs = (a, b) => [...a.children, ...b.children];
    for (const el of hdrs(this.elColInner, this.elColFz)) {
      const c = +el.dataset.c;
      el.classList.toggle("xv-hsel", c >= sel.c1 && c <= sel.c2);
      el.classList.toggle("xv-hfull", this.selKind === "cols" || this.selKind === "all");
    }
    for (const el of hdrs(this.elRowInner, this.elRowFz)) {
      const r = +el.dataset.r;
      el.classList.toggle("xv-hsel", r >= sel.r1 && r <= sel.r2);
      el.classList.toggle("xv-hfull", this.selKind === "rows" || this.selKind === "all");
    }
  }

  _renderHeaders(win) {
    const cx = this._cx;
    const fz = this._fz || { rows: 0, cols: 0 };
    const colHtml = (c0, c1) => {
      let ch = "";
      for (let c = c0; c <= c1; c++) {
        ch += `<div class="xv-ch" data-c="${c}" style="left:${cx[c]}px;width:${cx[c + 1] - cx[c]}px">${this._col(c)}<span class="xv-rz" data-c="${c}" title="Drag to resize, double-click to fit"></span></div>`;
      }
      return ch;
    };
    const af = this._afRange();
    const filtered = af && this._hidden.length > 0;
    const rowHtml = (r0, r1) => {
      let rh = "";
      for (let r = r0; r <= r1; r++) {
        if (this._isHidden(r)) continue;
        const cls = filtered && r > af.r1 && r <= af.r2 ? " xv-rfilt" : "";
        const gap = this._isHidden(r + 1) ? " xv-rgap" : "";
        rh += `<div class="xv-rhd${cls}${gap}" data-r="${r}" style="top:${this._ry(r)}px">${r + 1}</div>`;
      }
      return rh;
    };
    this.elColInner.innerHTML = colHtml(win.c0, win.c1);
    this.elRowInner.innerHTML = rowHtml(win.r0, win.r1);
    this.elColFz.innerHTML = fz.cols ? colHtml(0, Math.min(fz.cols, this.ncols) - 1) : "";
    this.elRowFz.innerHTML = fz.rows ? rowHtml(0, Math.min(fz.rows, this.nrows) - 1) : "";
    this.elColInner.style.transform = `translateX(${-this.elScroll.scrollLeft}px)`;
    this.elRowInner.style.transform = `translateY(${-this.elScroll.scrollTop}px)`;
  }

  _renderTabs() {
    if (!this.elTabs) return;
    const m = this._model();
    const d = this.doc;
    this.elTabs.textContent = "";
    m.sheets.forEach((s, i) => {
      const tab = d.createElement("div");
      tab.className = "sheet-tab" + (i === this._si() ? " active" : "") + (s.hidden ? " hidden-tab" : "");
      tab.setAttribute("role", "tab");
      tab.setAttribute("aria-selected", i === this._si() ? "true" : "false");
      tab.tabIndex = i === this._si() ? 0 : -1;
      tab.dataset.index = String(i);
      tab.title = `${s.name}${s.hidden ? " (hidden)" : ""} — double-click to rename, right-click for more`;
      const name = d.createElement("span");
      name.className = "tab-name";
      name.textContent = s.name;
      tab.appendChild(name);
      if (m.sheets.length > 1) {
        const x = d.createElement("button");
        x.className = "tab-x";
        x.type = "button";
        x.textContent = "×";
        x.title = `Delete sheet "${s.name}"`;
        x.setAttribute("aria-label", `Delete sheet ${s.name}`);
        x.dataset.del = String(i);
        tab.appendChild(x);
      }
      this.elTabs.appendChild(tab);
    });
  }

  _syncBar() {
    const rg = this._selRange();
    const multi = rg.r1 !== rg.r2 || rg.c1 !== rg.c2;
    if (this.doc.activeElement !== this.elNameBox) {
      if (this.selKind === "all") this.elNameBox.value = "All";
      else if (this.selKind === "rows") this.elNameBox.value = `${rg.r1 + 1}:${rg.r2 + 1}`;
      else if (this.selKind === "cols") this.elNameBox.value = `${this._col(rg.c1)}:${this._col(rg.c2)}`;
      else {
        const m = this._mergeAt(this.cur.r, this.cur.c);
        this.elNameBox.value = multi && !(m && m.r1 === rg.r1 && m.c1 === rg.c1 && m.r2 === rg.r2 && m.c2 === rg.c2)
          ? `${rcToAddr(rg.r1, rg.c1)}:${rcToAddr(rg.r2, rg.c2)}`
          : rcToAddr(m ? m.r1 : this.cur.r, m ? m.c1 : this.cur.c);
      }
    }
    // Keep the formula bar in sync unless the user is typing in it right now.
    if (!this.editing && !this.fbDirty) this.elFormula.value = cellRaw(this._curCell());
    this._updateSelectionStats(rg, multi);
  }

  _updateSelectionStats(rg, multi) {
    if (!multi) {
      if (this._statsMsg) { this.onStatus(""); this._statsMsg = false; }
      return;
    }
    const cl = this._selClipped();
    const area = (cl.r2 - cl.r1 + 1) * (cl.c2 - cl.c1 + 1);
    if (area > 200000) return;
    const sheet = this._sheet();
    let sum = 0, n = 0, count = 0;
    for (let r = cl.r1; r <= cl.r2; r++) {
      for (let c = cl.c1; c <= cl.c2; c++) {
        const cell = sheet.cells[rcToAddr(r, c)];
        if (cellIsEmpty(cell)) continue;
        count++;
        if (cell.t === "n" && typeof cell.v === "number") { sum += cell.v; n++; }
      }
    }
    const fmt = (x) => String(parseFloat(x.toPrecision(12)));
    const msg = n ? `Sum: ${fmt(sum)}   Average: ${fmt(sum / n)}   Count: ${count}` : count ? `Count: ${count}` : "";
    this.onStatus(msg);
    this._statsMsg = true;
  }

  _updateToolbar() {
    if (!this.elToolbar) return;
    const cell = this._curCell();
    const s = (cell && cell.s) || {};
    const set = (act, on) => {
      const b = this.elToolbar.querySelector(`[data-act="${act}"]`);
      if (b) b.setAttribute("aria-pressed", on ? "true" : "false");
    };
    set("b", !!s.b); set("i", !!s.i); set("u", !!s.u);
    const align = s.align || "";
    set("align-left", align === "left"); set("align-center", align === "center"); set("align-right", align === "right");
    const z = (cell && cell.z) || (cell && cell.t === "d" ? (/T/.test(String(cell.v)) ? "yyyy-mm-dd hh:mm" : "yyyy-mm-dd") : "");
    const sel = this.elNumFmt;
    if (sel) {
      let opt = [...sel.options].find((o) => o.value === z);
      sel.querySelectorAll("option[data-custom]").forEach((o) => { if (o !== opt) o.remove(); });
      if (!opt) {
        opt = this.doc.createElement("option");
        opt.value = z;
        opt.dataset.custom = "1";
        opt.textContent = `Custom (${z})`;
        sel.appendChild(opt);
      }
      sel.value = z;
    }
    const sw = this.elToolbar.querySelector(".xv-swatch-color");
    if (sw) sw.style.background = s.color || "#c00000";
    const bw = this.elToolbar.querySelector(".xv-swatch-bg");
    if (bw) bw.style.background = s.bg || "#fff2cc";
    set("filter", !!(this._sheet() && this._sheet().autoFilter));
    const undo = this.elToolbar.querySelector('[data-act="undo"]');
    const redo = this.elToolbar.querySelector('[data-act="redo"]');
    if (undo) undo.disabled = !this.undoStack.length;
    if (redo) redo.disabled = !this.redoStack.length;
  }

  // =====================================================================
  // Selection & navigation
  // =====================================================================

  /** Select a single cell by address (public, kept for compatibility). */
  /** Select a cell ("B2") or a range ("B2:D5"); the cursor goes to the first cell. */
  select(addr) {
    const [a, b] = String(addr || "").split(":");
    const rc = addrToRC(a);
    if (!rc) return;
    const end = b ? addrToRC(b) : rc;
    if (!end) return;
    this._commitAll();
    this._setSelection(rc, end, "cells", rc);
  }

  selectAll() {
    this._setSelection({ r: 0, c: 0 }, { r: MAX_ROWS - 1, c: MAX_COLS - 1 }, "all", { r: 0, c: 0 });
  }

  /** Set selection from anchor to focus; `cur` defaults to anchor. */
  _setSelection(anchor, focus, kind = "cells", cur = null, scroll = true) {
    // Any selection change ends a Tab data-entry run (see _moveAfterEntry).
    this._tabStartCol = null;
    this.anchor = { r: clamp(anchor.r, 0, MAX_ROWS - 1), c: clamp(anchor.c, 0, MAX_COLS - 1) };
    this.focusCell = { r: clamp(focus.r, 0, MAX_ROWS - 1), c: clamp(focus.c, 0, MAX_COLS - 1) };
    const cc = cur || anchor;
    this.cur = { r: clamp(cc.r, 0, MAX_ROWS - 1), c: clamp(cc.c, 0, MAX_COLS - 1) };
    const m = this._mergeAt(this.cur.r, this.cur.c);
    if (m && kind === "cells" && anchor.r === focus.r && anchor.c === focus.c) {
      this.cur = { r: m.r1, c: m.c1 };
      this.anchor = { r: m.r1, c: m.c1 };
      this.focusCell = { r: m.r1, c: m.c1 };
    }
    this.selKind = kind;
    // Whole-row/column selections extend to the sheet edge; only grow the
    // virtual grid for the part the user actually navigated to.
    const er = kind === "cols" || kind === "all" ? this.cur.r : this.focusCell.r;
    const ec = kind === "rows" || kind === "all" ? this.cur.c : this.focusCell.c;
    const grew = this._ensureExtent(er, ec);
    if (scroll) this._scrollIntoView(kind === "cells" ? this.focusCell : this.cur);
    if (grew) this._renderGrid(true); else this._renderGrid(false);
    this._syncBar();
    this._updateToolbar();
  }

  _scrollIntoView(rc) {
    const sc = this.elScroll;
    const r = clamp(rc.r, 0, this.nrows - 1), c = clamp(rc.c, 0, this.ncols - 1);
    const fz = this._fz || { rows: 0, cols: 0 };
    const { fh, fw } = this._fzSize();
    const h = sc.clientHeight, w = sc.clientWidth;
    if (r >= fz.rows) {
      const top = this._ry(r), bottom = top + ROW_H;
      if (top < sc.scrollTop + fh) sc.scrollTop = top - fh;
      else if (bottom > sc.scrollTop + h) sc.scrollTop = bottom - h;
    }
    if (c >= fz.cols) {
      const left = this._cx[c], right = this._cx[c + 1];
      if (left < sc.scrollLeft + fw) sc.scrollLeft = left - fw;
      else if (right > sc.scrollLeft + w) sc.scrollLeft = Math.min(left - fw, right - w);
    }
  }

  _move(dr, dc, extend = false) {
    if (extend) {
      const f = this.focusCell;
      this._setSelection(this.anchor, { r: dr ? this._stepRow(f.r, dr) : f.r, c: f.c + dc }, this.selKind === "cells" ? "cells" : this.selKind, this.cur);
      return;
    }
    let { r, c } = this.cur;
    const m = this._mergeAt(r, c);
    if (m) {
      if (dr > 0) r = m.r2; else if (dr < 0) r = m.r1;
      if (dc > 0) c = m.c2; else if (dc < 0) c = m.c1;
    }
    const t = { r: dr ? this._stepRow(r, dr) : r, c: clamp(c + dc, 0, MAX_COLS - 1) };
    this._setSelection(t, t, "cells");
  }

  /**
   * Move after committing an entry, with Excel's data-entry behaviour: Tab
   * moves right and remembers the column the run started in; Enter then goes
   * to the next row *in that starting column* (A1 Tab B1 Tab C1 Enter -> A2).
   */
  _moveAfterEntry(dr, dc) {
    const start = this._tabStartCol;
    if (dc && !dr) {
      const keep = start == null ? this.cur.c : start;
      this._move(0, dc);
      this._tabStartCol = keep;
      return;
    }
    if (dr > 0 && !dc && start != null) {
      const m = this._mergeAt(this.cur.r, this.cur.c);
      const r = clamp((m ? m.r2 : this.cur.r) + dr, 0, MAX_ROWS - 1);
      this._setSelection({ r, c: start }, { r, c: start }, "cells");
      return;
    }
    this._move(dr, dc);
  }

  _hasValue(r, c) {
    const cell = this._cellAt(r, c);
    return !cellIsEmpty(cell);
  }

  /** Ctrl+Arrow: jump to the edge of the current data region. */
  _jump(dr, dc, extend) {
    const start = extend ? this.focusCell : this.cur;
    let { r, c } = start;
    const maxR = Math.max(this._used.r2, 0), maxC = Math.max(this._used.c2, 0);
    const limR = dr > 0 ? Math.max(maxR, r) : 0, limC = dc > 0 ? Math.max(maxC, c) : 0;
    const inBounds = (rr, cc) => (dr > 0 ? rr <= limR : dr < 0 ? rr >= 0 : true) && (dc > 0 ? cc <= limC : dc < 0 ? cc >= 0 : true);
    const nr = r + dr, nc = c + dc;
    if (!inBounds(nr, nc)) { /* stay */ }
    else if (this._hasValue(r, c) && this._hasValue(nr, nc)) {
      while (inBounds(r + dr, c + dc) && this._hasValue(r + dr, c + dc)) { r += dr; c += dc; }
    } else {
      r = nr; c = nc;
      while (!this._hasValue(r, c) && inBounds(r + dr, c + dc)) { r += dr; c += dc; }
    }
    const t = { r: this._visibleRow(clamp(r, 0, MAX_ROWS - 1)), c: clamp(c, 0, MAX_COLS - 1) };
    if (extend) this._setSelection(this.anchor, t, "cells", this.cur);
    else this._setSelection(t, t, "cells");
  }

  _gotoRef(text) {
    let ref = String(text || "").trim();
    if (!ref) { this._focusGrid(); return; }
    const m = this._model();
    let sheetIdx = this._si();
    const bang = ref.lastIndexOf("!");
    if (bang > 0) {
      let name = ref.slice(0, bang);
      if (/^'.*'$/.test(name)) name = name.slice(1, -1).replace(/''/g, "'");
      const i = m.sheets.findIndex((s) => s.name.toLowerCase() === name.toLowerCase());
      if (i < 0) { this.onStatus(`No sheet named "${name}"`); return; }
      sheetIdx = i;
      ref = ref.slice(bang + 1);
    }
    const rg = parseRange(ref.replace(/\$/g, "").toUpperCase());
    if (!rg) { this.onStatus(`Not a cell reference: ${text}`); this._syncBar(); return; }
    if (sheetIdx !== this._si()) this._switchSheet(sheetIdx);
    this._setSelection({ r: rg.r1, c: rg.c1 }, { r: rg.r2, c: rg.c2 }, "cells", { r: rg.r1, c: rg.c1 });
    this._scrollIntoView({ r: rg.r1, c: rg.c1 });
    this._renderGrid(false);
    this._focusGrid();
  }

  /** Put keyboard focus in the grid (public). */
  focus() {
    this._focusGrid();
  }

  /**
   * Cheap status information for the host page (no full scan of the cells):
   * {sheet, sheets, ref:"A1:D20"|"", rows, cols, selection:"B2:C5"}.
   */
  getStats() {
    const m = this._model();
    const sheet = this._sheet();
    if (!m || !sheet) return null;
    const used = this._used || { r2: -1, c2: -1 };
    const rg = this._selRange();
    return {
      sheet: sheet.name,
      sheets: m.sheets.length,
      rows: used.r2 + 1,
      cols: used.c2 + 1,
      ref: used.r2 >= 0 ? `A1:${rcToAddr(used.r2, used.c2)}` : "",
      selection: this.elNameBox ? this.elNameBox.value : rcToAddr(rg.r1, rg.c1),
    };
  }

  _focusGrid() {
    if (this._destroyed) return;
    if (this.doc.activeElement !== this.elEditor) this.elEditor.focus({ preventScroll: true });
  }

  // =====================================================================
  // Mouse
  // =====================================================================

  _cellFromEvent(e) {
    const rect = this.elScroll.getBoundingClientRect();
    const vx = e.clientX - rect.left, vy = e.clientY - rect.top;
    const { fh, fw } = this._fzSize();
    const x = fw && vx < fw ? vx : vx + this.elScroll.scrollLeft;
    const y = fh && vy < fh ? vy : vy + this.elScroll.scrollTop;
    return { r: clamp(this._rowAtY(Math.max(0, y)), 0, this.nrows - 1), c: this._colAtX(Math.max(0, x)) };
  }

  _onCellMouseDown(e) {
    if (e.target === this.elEditor && this.editing) return; // caret placement inside the editor
    const chartEl = e.target.closest && e.target.closest(".xv-chart");
    if (chartEl) { this._onChartMouseDown(e, chartEl); return; }
    if (this._selChart >= 0) this._selectChart(-1, false);
    this._closeChartPanel();
    const fb = e.target.closest && e.target.closest(".xv-fbtn");
    if (fb && e.button === 0) {
      e.preventDefault();
      e.stopPropagation();
      this._commitAll();
      this._openFilterPanel(+fb.dataset.fc);
      return;
    }
    if (e.target === this.elFillH && e.button === 0) {
      e.preventDefault();
      this._commitAll();
      this.drag = { kind: "fill", src: this._selClipped(), target: null, ctrl: e.ctrlKey };
      this._renderOverlays();
      return;
    }
    const sc = this.elScroll;
    // ignore clicks on the scrollbars
    if (e.target === sc && (e.offsetX >= sc.clientWidth || e.offsetY >= sc.clientHeight)) return;
    if (e.button !== 0 && e.button !== 2) return;
    const rc = this._cellFromEvent(e);
    if (e.button === 0 && this._canPointRef()) {
      // Formula "point mode": clicking a cell while typing a formula inserts its reference.
      e.preventDefault();
      this._pointRef(rc, e.shiftKey && this._point ? this._point.anchor : rc);
      this.drag = { kind: "point" };
      return;
    }
    if (e.button === 2) {
      const rg = this._selRange();
      if (rc.r < rg.r1 || rc.r > rg.r2 || rc.c < rg.c1 || rc.c > rg.c2) {
        this._commitAll();
        this._setSelection(rc, rc, "cells", null, false);
      }
      return;
    }
    e.preventDefault();
    this._commitAll();
    this._hideMenu();
    if (e.shiftKey) {
      this._setSelection(this.anchor, rc, "cells", this.cur, false);
    } else {
      this._setSelection(rc, rc, "cells", null, false);
    }
    this.drag = { kind: "cells" };
    this._focusGrid();
  }

  _onCellDblClick(e) {
    if (e.target === this.elEditor) return;
    if (e.target === this.elFillH) { this._fillDownAuto(); return; }
    if (e.target.closest && (e.target.closest(".xv-chart") || e.target.closest(".xv-fbtn"))) return;
    const rc = this._cellFromEvent(e);
    this._setSelection(rc, rc, "cells", null, false);
    this._beginEdit("edit", cellRaw(this._curCell()));
  }

  _onHeaderMouseDown(e, axis) {
    if (e.button !== 0 && e.button !== 2) return;
    const rz = e.target.closest(".xv-rz");
    if (rz && axis === "col" && e.button === 0) {
      e.preventDefault();
      const c = +rz.dataset.c;
      this._commitAll();
      this.drag = { kind: "resize", c, startX: e.clientX, startW: this._colW(c) };
      this.elWrap.classList.add("xv-resizing");
      this._focusGrid();
      return;
    }
    const hd = e.target.closest(axis === "col" ? ".xv-ch" : ".xv-rhd");
    if (!hd) return;
    e.preventDefault();
    const i = +(axis === "col" ? hd.dataset.c : hd.dataset.r);
    if (e.button === 2) {
      const rg = this._selRange();
      const inside = axis === "col" ? i >= rg.c1 && i <= rg.c2 && this.selKind !== "rows" : i >= rg.r1 && i <= rg.r2 && this.selKind !== "cols";
      if (inside && (this.selKind === (axis === "col" ? "cols" : "rows") || this.selKind === "all")) return;
    }
    this._commitAll();
    this._selectLine(axis, e.shiftKey ? null : i, i);
    if (e.button === 0) this.drag = { kind: axis === "col" ? "cols" : "rows" };
    this._focusGrid();
  }

  /** Select whole columns/rows from anchor index `a` (null = keep) to `b`. */
  _selectLine(axis, a, b) {
    if (axis === "col") {
      const ac = a == null ? this.anchor.c : a;
      const top = clamp(this._rowAtY(this.elScroll.scrollTop + this._fzSize().fh), 0, MAX_ROWS - 1);
      this._setSelection({ r: 0, c: ac }, { r: MAX_ROWS - 1, c: b }, "cols", { r: a == null ? this.cur.r : top, c: a == null ? this.cur.c : ac }, false);
    } else {
      const ar = a == null ? this.anchor.r : a;
      this._setSelection({ r: ar, c: 0 }, { r: b, c: MAX_COLS - 1 }, "rows", { r: a == null ? this.cur.r : ar, c: 0 }, false);
    }
  }

  _onHeaderDblClick(e) {
    const rz = e.target.closest(".xv-rz");
    if (!rz) return;
    this._autofit(+rz.dataset.c);
  }

  _onDragMove(e) {
    const d = this.drag;
    if (!d) return;
    if (d.kind === "resize") {
      const w = Math.max(8, d.startW + (e.clientX - d.startX));
      d.newW = w;
      this._setColPx(d.c, w, false);
      return;
    }
    if (d.kind === "chart") { this._chartDragMove(e); return; }
    const rect = this.elScroll.getBoundingClientRect();
    const out = e.clientY < rect.top || e.clientY > rect.bottom || e.clientX < rect.left || e.clientX > rect.right;
    this._lastDragEvent = { clientX: e.clientX, clientY: e.clientY, ctrlKey: e.ctrlKey || e.metaKey };
    if (out && !this._dragTimer) {
      this._dragTimer = setInterval(() => this._autoScrollDrag(), 50);
    } else if (!out && this._dragTimer) {
      clearInterval(this._dragTimer);
      this._dragTimer = null;
    }
    this._dragTo(e);
  }

  _dragTo(e) {
    const d = this.drag;
    const rect = this.elScroll.getBoundingClientRect();
    const cx = clamp(e.clientX, rect.left, rect.left + this.elScroll.clientWidth - 1);
    const cy = clamp(e.clientY, rect.top, rect.top + this.elScroll.clientHeight - 1);
    const rc = this._cellFromEvent({ clientX: cx, clientY: cy });
    if (d.kind === "fill") {
      this._fillDragTo(rc, !!e.ctrlKey);
    } else if (d.kind === "point") {
      if (this._point && (rc.r !== this._point.focus.r || rc.c !== this._point.focus.c)) this._pointRef(rc, this._point.anchor);
    } else if (d.kind === "cells") {
      if (rc.r !== this.focusCell.r || rc.c !== this.focusCell.c) this._setSelection(this.anchor, rc, "cells", this.cur, false);
    } else if (d.kind === "cols") {
      if (rc.c !== this.focusCell.c) this._selectLine("col", null, rc.c);
    } else if (d.kind === "rows") {
      if (rc.r !== this.focusCell.r) this._selectLine("row", null, rc.r);
    }
  }

  _autoScrollDrag() {
    const e = this._lastDragEvent;
    if (!this.drag || !e) return;
    const rect = this.elScroll.getBoundingClientRect();
    const sc = this.elScroll;
    if (e.clientY > rect.bottom) sc.scrollTop += ROW_H;
    else if (e.clientY < rect.top) sc.scrollTop -= ROW_H;
    if (e.clientX > rect.right) sc.scrollLeft += 40;
    else if (e.clientX < rect.left) sc.scrollLeft -= 40;
    this._dragTo(e);
  }

  _endDrag(e) {
    const d = this.drag;
    if (!d) return;
    this.drag = null;
    if (this._dragTimer) { clearInterval(this._dragTimer); this._dragTimer = null; }
    if (d.kind === "chart") { this._chartDragEnd(d); return; }
    if (d.kind === "fill") {
      this.elFillPrev.hidden = true;
      if (d.target) this._fillTo(d.src, d.target.dir, d.target.n, e ? !!(e.ctrlKey || e.metaKey) : !!d.ctrl);
      else this._renderOverlays();
      this._focusGrid();
      return;
    }
    if (d.kind === "resize") {
      this.elWrap.classList.remove("xv-resizing");
      if (d.newW != null && Math.abs(d.newW - d.startW) >= 1) this._setColPx(d.c, d.newW, true);
      else this._layoutCols();
    }
  }

  /** Set a column width in px. commit=true records undo and marks dirty. */
  _setColPx(c, px, commit) {
    const sheet = this._sheet();
    if (!commit) {
      // live preview only
      const delta = px - this._colW(c);
      for (let k = c + 1; k <= this.ncols; k++) this._cx[k] += delta;
      this.elCanvas.style.width = `${this._cx[this.ncols]}px`;
      this._renderGrid(true);
      return;
    }
    const rg = this._selRange();
    const cols = this.selKind === "cols" && c >= rg.c1 && c <= rg.c2
      ? Array.from({ length: Math.min(rg.c2 - rg.c1 + 1, 500) }, (_, i) => rg.c1 + i)
      : [c];
    const after = { ...(sheet.colWidths || {}) };
    for (const k of cols) after[indexToCol(k)] = pxToChars(px);
    this._setColWidths(after, "Resize column");
  }

  _setColWidths(after, label) {
    const sheet = this._sheet();
    this._metaOp({ kind: "meta", label, si: this._si(), key: "colWidths", before: clone(sheet.colWidths) || null, after });
  }

  _autofit(c) {
    const sheet = this._sheet();
    let max = 0;
    const col = indexToCol(c);
    const r2 = Math.min(this._used.r2, 20000);
    for (let r = 0; r <= r2; r++) {
      const cell = sheet.cells[col + (r + 1)];
      if (cellIsEmpty(cell)) continue;
      const { text } = formatCell(cell);
      max = Math.max(max, this._measure(text, this._fontFor(cell)));
    }
    const px = clamp(Math.ceil(max + 14), 24, 800);
    this._setColWidths({ ...(sheet.colWidths || {}), [col]: pxToChars(px) }, "Fit column width");
  }

  // =====================================================================
  // Keyboard
  // =====================================================================

  _onKeyDown(e) {
    if (e.isComposing || e.keyCode === 229) return;
    if (this.editing) { this._onEditKey(e); return; }
    const mod = e.ctrlKey || e.metaKey;
    const k = e.key;
    if (!this.elMenu.hidden && k === "Escape") { this._hideMenu(); return; }
    if (mod && !e.altKey) {
      switch (k.toLowerCase()) {
        case "z": e.preventDefault(); if (e.shiftKey) this.redo(); else this.undo(); return;
        case "y": e.preventDefault(); this.redo(); return;
        case "a": e.preventDefault(); this.selectAll(); return;
        case "b": e.preventDefault(); this._toggleStyle("b"); return;
        case "i": e.preventDefault(); this._toggleStyle("i"); return;
        case "u": e.preventDefault(); this._toggleStyle("u"); return;
        case "f": e.preventDefault(); this.openFind(false); return;
        case "h": e.preventDefault(); this.openFind(true); return;
        case "l": if (e.shiftKey) { e.preventDefault(); this._toggleFilter(); } return;
        case "d": e.preventDefault(); this._fill("down"); return;
        case "r": e.preventDefault(); this._fill("right"); return;
        case ";": e.preventDefault(); this._insertNow(false); return;
        case ":": e.preventDefault(); this._insertNow(true); return;
        case "home": e.preventDefault(); this._setSelection({ r: 0, c: 0 }, { r: 0, c: 0 }); return;
        case "end": {
          e.preventDefault();
          const t = { r: Math.max(0, this._used.r2), c: Math.max(0, this._used.c2) };
          this._setSelection(t, t);
          return;
        }
        case "arrowup": e.preventDefault(); this._jump(-1, 0, e.shiftKey); return;
        case "arrowdown": e.preventDefault(); this._jump(1, 0, e.shiftKey); return;
        case "arrowleft": e.preventDefault(); this._jump(0, -1, e.shiftKey); return;
        case "arrowright": e.preventDefault(); this._jump(0, 1, e.shiftKey); return;
        // Ctrl+Shift+= / Ctrl+- insert/delete whole selected rows/columns; otherwise
        // leave them to the browser (page zoom).
        case "+": case "=":
          if ((e.shiftKey || k === "+") && (this.selKind === "rows" || this.selKind === "cols")) { e.preventDefault(); this._insertForSelection(); }
          return;
        case "-":
          if (this.selKind === "rows" || this.selKind === "cols") { e.preventDefault(); this._deleteForSelection(); }
          return;
        case " ": e.preventDefault(); this._selectLine("col", this.cur.c, this.cur.c); return;
        default: return; // Ctrl+C/X/V -> clipboard events; Ctrl+S/O/P bubble to the window
      }
    }
    if (e.altKey && k === "ArrowDown") {
      const af = this._afRange();
      if (af && this.cur.r === af.r1 && this.cur.c >= af.c1 && this.cur.c <= af.c2) { e.preventDefault(); this._openFilterPanel(this.cur.c); return; }
    }
    switch (k) {
      case "ArrowUp": e.preventDefault(); this._move(-1, 0, e.shiftKey); return;
      case "ArrowDown": e.preventDefault(); this._move(1, 0, e.shiftKey); return;
      case "ArrowLeft": e.preventDefault(); this._move(0, -1, e.shiftKey); return;
      case "ArrowRight": e.preventDefault(); this._move(0, 1, e.shiftKey); return;
      case "Tab": e.preventDefault(); this._moveAfterEntry(0, e.shiftKey ? -1 : 1); return;
      case "Enter":
        e.preventDefault();
        if (e.shiftKey) this._move(-1, 0);
        else this._beginEdit("edit", cellRaw(this._curCell()));
        return;
      case "F2": e.preventDefault(); this._beginEdit("edit", cellRaw(this._curCell())); return;
      case "Delete": case "Backspace": e.preventDefault(); this.clearSelection(); return;
      case "Escape":
        if (this.clip) { this.clip = null; this._renderOverlays(); }
        if (this.elNotice && !this.elNotice.hidden) this.elNotice.hidden = true;
        return;
      case "PageDown": case "PageUp": {
        e.preventDefault();
        const n = Math.max(1, Math.floor(this.elScroll.clientHeight / ROW_H) - 1) * (k === "PageDown" ? 1 : -1);
        this.elScroll.scrollTop += n * ROW_H;
        this._move(n, 0, e.shiftKey);
        return;
      }
      case "Home": e.preventDefault(); this._setSelection({ r: this.cur.r, c: 0 }, { r: this.cur.r, c: 0 }); return;
      case " ":
        if (e.shiftKey) { e.preventDefault(); this._selectLine("row", this.cur.r, this.cur.r); }
        return;
      default:
        // Printable characters fall through to the textarea and start an edit via "input".
        return;
    }
  }

  _onEditKey(e) {
    if (this._acKey(e)) return;
    const mod = e.ctrlKey || e.metaKey;
    const k = e.key;
    if (mod && !e.altKey) {
      const lk = k.toLowerCase();
      if (lk === "s" || lk === "o" || lk === "p") {
        // Commit first, then let the shortcut reach the window handler (save/open/print).
        this._commitEdit(0, 0);
        return;
      }
      if (k === "Enter") { e.preventDefault(); this._commitEdit(0, 0, { fill: true }); return; }
      return; // native text editing shortcuts (undo, copy, paste, select all...)
    }
    switch (k) {
      case "Enter":
        e.preventDefault();
        if (e.altKey) { this._insertAtCaret("\n"); return; }
        this._commitEdit(e.shiftKey ? -1 : 1, 0);
        return;
      case "Tab":
        e.preventDefault();
        this._commitEdit(0, e.shiftKey ? -1 : 1);
        return;
      case "Escape":
        e.preventDefault();
        this._cancelEdit();
        return;
      case "F2":
        e.preventDefault();
        this.editing.mode = this.editing.mode === "enter" ? "edit" : "enter";
        this.onStatus(this.editing.mode === "edit" ? "Edit mode: arrow keys move the caret" : "Enter mode: arrow keys commit and move");
        return;
      case "ArrowUp": case "ArrowDown": case "ArrowLeft": case "ArrowRight":
        if (this.editing.mode === "enter" && !this.elEditor.value.startsWith("=")) {
          e.preventDefault();
          const d = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[k];
          this._commitEdit(d[0], d[1]);
        }
        return;
      default:
        this._point = null;
        return;
    }
  }

  /** True when the in-cell editor holds a formula and the caret is where a reference may go. */
  _canPointRef() {
    if (!this.editing) return false;
    const ta = this.elEditor;
    const v = ta.value;
    if (!v.startsWith("=")) return false;
    const pos = ta.selectionStart;
    const p = this._point;
    if (p && p.end === pos && v.slice(p.start, p.end) === p.text) return true;
    const before = v.slice(0, pos).replace(/\s+$/, "");
    return /[=(,+\-*/^&<>:;]$/.test(before);
  }

  /** Insert (or replace the last inserted) reference for the range anchor..rc at the caret. */
  _pointRef(rc, anchor) {
    const ta = this.elEditor;
    const r1 = Math.min(anchor.r, rc.r), r2 = Math.max(anchor.r, rc.r);
    const c1 = Math.min(anchor.c, rc.c), c2 = Math.max(anchor.c, rc.c);
    const text = r1 === r2 && c1 === c2 ? rcToAddr(r1, c1) : `${rcToAddr(r1, c1)}:${rcToAddr(r2, c2)}`;
    const p = this._point;
    const replacing = p && ta.value.slice(p.start, p.end) === p.text && ta.selectionStart === p.end;
    const start = replacing ? p.start : ta.selectionStart;
    const end = replacing ? p.end : ta.selectionEnd;
    ta.setRangeText(text, start, end, "end");
    this._point = { start, end: start + text.length, text, anchor, focus: rc };
    this.elFormula.value = ta.value;
    this._sizeEditor();
    this._place(this.elCopy, this._rectOf({ r1, r2, c1, c2 }));
    this.elCopy.hidden = false;
    this._focusGrid();
  }

  _insertAtCaret(text) {
    const ta = this.elEditor;
    const s = ta.selectionStart, en = ta.selectionEnd;
    ta.setRangeText(text, s, en, "end");
    this._onEditorInput();
  }

  _onEditorInput() {
    if (!this.editing) {
      if (this.elEditor.value === "") return;
      this._beginEdit("enter", null);
    }
    this.elFormula.value = this.elEditor.value;
    this._sizeEditor();
    this._acUpdate(this.elEditor);
  }

  _onEditorBlur() {
    // Commit when focus leaves the grid entirely (e.g. user clicks a toolbar control).
    setTimeout(() => {
      if (this._destroyed || !this.editing) return;
      const ae = this.doc.activeElement;
      if (ae === this.elEditor || ae === this.elFormula) return;
      this._commitEdit(0, 0, { keepFocus: false });
    }, 0);
  }

  _onFormulaKey(e) {
    if (this._acKey(e)) return;
    const mod = e.ctrlKey || e.metaKey;
    if (mod && !e.altKey && ["s", "o", "p"].includes(e.key.toLowerCase())) {
      if (this.fbDirty) this._commitFormulaBar(0, 0, false);
      return; // bubble to the window shortcut handler
    }
    if (e.key === "Enter") {
      e.preventDefault();
      this._commitFormulaBar(e.shiftKey ? -1 : 1, 0, true);
    } else if (e.key === "Tab") {
      e.preventDefault();
      this._commitFormulaBar(0, e.shiftKey ? -1 : 1, true);
    } else if (e.key === "Escape") {
      e.preventDefault();
      this.fbDirty = false;
      this.elFormula.value = cellRaw(this._curCell());
      this._focusGrid();
    }
  }

  // =====================================================================
  // Editing
  // =====================================================================

  /**
   * Start editing the active cell. mode "edit" (F2/Enter/double-click) keeps
   * arrows inside the text; mode "enter" (typing) commits on arrows.
   * text=null keeps what is already in the editor (typed char / IME composition).
   */
  _beginEdit(mode, text) {
    if (this.editing) return;
    this._hideMenu();
    this.elFillH.hidden = true;
    const m = this._mergeAt(this.cur.r, this.cur.c);
    const r = m ? m.r1 : this.cur.r, c = m ? m.c1 : this.cur.c;
    this._scrollIntoView({ r, c });
    this.editing = { r, c, mode };
    const ta = this.elEditor;
    if (text != null) ta.value = text;
    ta.classList.add("xv-editing");
    const cell = this._cellAt(r, c);
    ta.style.font = this._fontFor(cell);
    ta.style.textAlign = "left";
    this._placeEditor();
    this._focusGrid();
    if (text != null) ta.setSelectionRange(ta.value.length, ta.value.length);
    this.elFormula.value = ta.value;
  }

  _placeEditor() {
    if (!this.editing) return;
    const m = this._mergeAt(this.editing.r, this.editing.c);
    const rg = m || { r1: this.editing.r, r2: this.editing.r, c1: this.editing.c, c2: this.editing.c };
    const rect = this._viewRect(rg) || this._rectOf(rg);
    this._editorBase = rect;
    this._place(this.elEditor, rect);
    this._sizeEditor();
  }

  _sizeEditor() {
    if (!this.editing || !this._editorBase) return;
    const ta = this.elEditor;
    const base = this._editorBase;
    const lines = ta.value.split("\n");
    const font = ta.style.font || `${GRID_FONT_SIZE}px ${GRID_FONT_FAMILY}`;
    const widest = Math.max(...lines.map((l) => this._measure(l, font)));
    const maxW = Math.max(base.width, this.elScroll.clientWidth - (base.left - this.elScroll.scrollLeft) - 4);
    ta.style.width = `${clamp(widest + 20, base.width, Math.max(base.width, maxW))}px`;
    ta.style.height = `${Math.max(base.height, lines.length * 18 + 6)}px`;
  }

  _endEditUi() {
    this.editing = null;
    this._point = null;
    this._hideAutocomplete();
    const ta = this.elEditor;
    ta.value = "";
    ta.classList.remove("xv-editing");
    ta.style.width = "";
    ta.style.height = "";
    this._renderOverlays();
  }

  _cancelEdit() {
    if (!this.editing) return;
    this._endEditUi();
    this._syncBar();
    this._focusGrid();
  }

  /**
   * Commit the in-cell editor. dr/dc move the selection afterwards.
   * opts.fill: write the value into every cell of the selection (Ctrl+Enter).
   */
  _commitEdit(dr, dc, opts = {}) {
    if (!this.editing) return;
    const { r, c } = this.editing;
    const text = this.elEditor.value;
    this._endEditUi();
    if (opts.fill) {
      const rg = this._selClipped();
      const changes = [];
      for (let rr = rg.r1; rr <= rg.r2; rr++) {
        for (let cc = rg.c1; cc <= rg.c2; cc++) {
          const addr = rcToAddr(rr, cc);
          const base = this._prepareEntry(text, this._cellAt(rr, cc));
          let next = base;
          if (base && base.f && (rr !== r || cc !== c)) next = { ...base, f: translateFormula(base.f, rr - r, cc - c), v: null, t: "f" };
          changes.push({ addr, after: next });
        }
      }
      this._applyChanges(changes, "Fill");
    } else {
      const addr = rcToAddr(r, c);
      const prev = this._cellAt(r, c);
      const next = this._prepareEntry(text, prev);
      if (next !== prev) this._applyChanges([{ addr, after: next }], "Edit");
    }
    if (dr || dc) this._moveAfterEntry(dr, dc);
    else this._syncBar();
    if (opts.keepFocus !== false) this._focusGrid();
  }

  /** Text typed by the user -> new cell (or the same object when unchanged). */
  _prepareEntry(text, prev) {
    const next = cellFromInput(text, prev);
    if (next === prev || !next) return next;
    if (next.f) {
      next.f = normalizeFormula(next.f);
      if (!next.z) {
        const z = formulaAutoFormat(this._model(), this._si(), next.f);
        if (z) next.z = z;
      }
    }
    return next;
  }

  _commitFormulaBar(dr, dc, focusGrid) {
    const text = this.elFormula.value;
    this.fbDirty = false;
    const m = this._mergeAt(this.cur.r, this.cur.c);
    const r = m ? m.r1 : this.cur.r, c = m ? m.c1 : this.cur.c;
    const prev = this._cellAt(r, c);
    const next = this._prepareEntry(text, prev);
    if (next !== prev) this._applyChanges([{ addr: rcToAddr(r, c), after: next }], "Edit");
    if (dr || dc) this._moveAfterEntry(dr, dc);
    else this._syncBar();
    if (focusGrid) this._focusGrid();
  }

  /** Flush any in-progress edit (cell editor or formula bar) into the model. */
  commitEdit() {
    if (this.editing) this._commitEdit(0, 0, { keepFocus: false });
    if (this.fbDirty) this._commitFormulaBar(0, 0, false);
  }

  _commitAll() {
    if (this.editing) this._commitEdit(0, 0, { keepFocus: false });
    if (this.fbDirty) this._commitFormulaBar(0, 0, false);
  }

  // =====================================================================
  // Changes, undo/redo, recalculation
  // =====================================================================

  /**
   * Apply cell changes to the active sheet as one undoable step.
   * changes: [{addr, after}] (after null deletes the cell).
   */
  _applyChanges(changes, label = "Edit", opts = {}) {
    const si = opts.si ?? this._si();
    const sheet = this._model().sheets[si];
    const recs = [];
    for (const ch of changes) {
      const before = sheet.cells[ch.addr];
      if (sameCell(before, ch.after)) continue;
      recs.push({ addr: ch.addr, before: clone(before), after: clone(ch.after) });
      if (ch.after == null) delete sheet.cells[ch.addr];
      else sheet.cells[ch.addr] = ch.after;
    }
    if (!recs.length) return false;
    this._pushUndo({ kind: "cells", si, recs, label, sel: this._selState() });
    this._afterCellsChanged(si, recs.map((x) => x.addr));
    return true;
  }

  _selState() {
    return { cur: { ...this.cur }, anchor: { ...this.anchor }, focus: { ...this.focusCell }, kind: this.selKind };
  }

  _pushUndo(entry) {
    this.undoStack.push(entry);
    if (this.undoStack.length > MAX_UNDO) this.undoStack.shift();
    this.redoStack = [];
  }

  _afterCellsChanged(si, addrs) {
    const sheet = this._model().sheets[si];
    for (const a of addrs) {
      const rc = addrToRC(a);
      if (rc && sheet.cells[a]) {
        if (si === this._si()) {
          if (rc.r > this._used.r2) this._used.r2 = rc.r;
          if (rc.c > this._used.c2) this._used.c2 = rc.c;
        }
      }
    }
    this._recalc([{ sheet: si, addrs }]);
    this._scheduleCharts();
    this._renderGrid(true);
    this._syncBar();
    this._updateToolbar();
    this.onDirty(true);
    this.onStats();
  }

  _recalc(changed, initial = false) {
    try {
      const res = recalcBook(this._model(), changed ? { changed } : {});
      if (res.cycles.length) {
        this.onStatus(`Circular reference: ${res.cycles.slice(0, 3).join(", ")}${res.cycles.length > 3 ? "…" : ""} (shown as #CIRC!)`);
        this._statsMsg = false;
      }
      return res;
    } catch (e) {
      if (typeof console !== "undefined") console.warn("recalc failed", e);
      if (!initial) this.onStatus(`Recalculation failed: ${e.message}`);
      return null;
    }
  }

  _snapshot() {
    const m = this._model();
    return JSON.stringify({ active: m.active || 0, sheets: m.sheets });
  }

  _restore(snap) {
    const m = this._model();
    const data = JSON.parse(snap);
    m.sheets.length = 0;
    for (const sh of data.sheets) m.sheets.push(sh);
    m.active = data.active;
  }

  /**
   * Run a whole-workbook operation (insert/delete rows/columns, delete,
   * duplicate or rename sheets) as one undoable step using JSON snapshots.
   * Very large workbooks skip the snapshot (the step cannot be undone).
   */
  _bookOp(label, fn) {
    const m = this._model();
    let total = 0;
    for (const sh of m.sheets) total += Object.keys(sh.cells || {}).length;
    const canUndo = total <= BIG_SNAPSHOT_CELLS;
    const before = canUndo ? this._snapshot() : null;
    const sel = this._selState();
    fn();
    if (canUndo) {
      this._pushUndo({ kind: "book", label, before, after: this._snapshot(), sel });
      // keep memory bounded: at most 30 snapshot entries
      let books = 0;
      for (let k = this.undoStack.length - 1; k >= 0; k--) {
        if (this.undoStack[k].kind === "book" && ++books > 30) this.undoStack.splice(k, 1);
      }
    } else {
      this.onStatus(`${label} cannot be undone (workbook too large for undo).`);
      this._statsMsg = false;
      this.undoStack = [];
      this.redoStack = [];
    }
    this._afterBookChange(true);
  }

  /**
   * Cheap, index-based workbook changes that never touch cells:
   *  {kind:"meta", si, key, before, after}  (colWidths / hidden)
   *  {kind:"move", from, to}
   *  {kind:"addSheet", index, sheet}
   * plus the active sheet before/after.
   */
  _metaOp(entry) {
    const m = this._model();
    entry.activeBefore = m.active || 0;
    entry.sel = this._selState();
    this._applyMeta(entry, "after");
    entry.activeAfter = m.active || 0;
    this._pushUndo(entry);
    this._afterBookChange(false);
  }

  _applyMeta(e, which) {
    const m = this._model();
    const undo = which === "before";
    switch (e.kind) {
      case "meta": {
        const sh = m.sheets[e.si];
        if (!sh) break;
        const v = clone(e[which]);
        if (v == null || v === false) delete sh[e.key]; else sh[e.key] = v;
        break;
      }
      case "move": {
        const [from, to] = undo ? [e.to, e.from] : [e.from, e.to];
        const [sh] = m.sheets.splice(from, 1);
        m.sheets.splice(to, 0, sh);
        break;
      }
      case "addSheet":
        if (undo) m.sheets.splice(e.index, 1);
        else m.sheets.splice(e.index, 0, clone(e.sheet));
        break;
      default: break;
    }
    if (which === "after" && e.activeAfter != null) m.active = e.activeAfter;
    else if (which === "after" && e.activeTo != null) m.active = e.activeTo;
    else if (undo && e.activeBefore != null) m.active = e.activeBefore;
    m.active = clamp(m.active || 0, 0, m.sheets.length - 1);
  }

  _afterBookChange(recalc) {
    const keepSel = this._selState();
    for (const s of this._model().sheets) invalidateSheet(s);
    if (recalc) this._recalc(null);
    const sheet = this._sheet();
    const sameSheet = this._sheetRef === sheet;
    const scroll = { top: this.elScroll.scrollTop, left: this.elScroll.scrollLeft };
    this._resetSheetState();
    if (sameSheet) {
      this.elScroll.scrollTop = scroll.top;
      this.elScroll.scrollLeft = scroll.left;
      this._restoreSel(keepSel, false);
    }
    this._renderTabs();
    this._renderGrid(true);
    this._syncBar();
    this._updateToolbar();
    this.onDirty(true);
    this.onStats();
  }

  _restoreSel(st, scroll = true) {
    if (!st) return;
    this._setSelection(st.anchor, st.focus, st.kind, st.cur, scroll);
  }

  undo() {
    this._commitAll();
    const e = this.undoStack.pop();
    if (!e) { this.onStatus("Nothing to undo"); return; }
    this.redoStack.push(e);
    this._applyHistory(e, "before");
    this.onStatus(`Undo: ${e.label}`);
    this._statsMsg = false;
  }

  redo() {
    this._commitAll();
    const e = this.redoStack.pop();
    if (!e) { this.onStatus("Nothing to redo"); return; }
    this.undoStack.push(e);
    this._applyHistory(e, "after");
    this.onStatus(`Redo: ${e.label}`);
    this._statsMsg = false;
  }

  _applyHistory(e, which) {
    if (e.kind === "cells") {
      const m = this._model();
      const sheet = m.sheets[e.si];
      if (!sheet) return;
      for (const rec of e.recs) {
        const v = clone(rec[which]);
        if (v == null) delete sheet.cells[rec.addr]; else sheet.cells[rec.addr] = v;
      }
      if (this._si() !== e.si) {
        m.active = e.si;
        this._resetSheetState();
        this._renderTabs();
      }
      this._afterCellsChanged(e.si, e.recs.map((r) => r.addr));
      this._restoreSel(e.sel);
      return;
    }
    if (e.kind === "custom") {
      if (which === "before") e.undo(); else e.redo();
      this._afterCustom(e);
    } else if (e.kind === "book") {
      this._restore(e[which]);
      this._afterBookChange(true);
    } else {
      this._applyMeta(e, which);
      this._afterBookChange(false);
    }
    this._restoreSel(e.sel);
  }

  // =====================================================================
  // Clear, fill, formatting
  // =====================================================================

  /** Delete/Backspace: clear values in the selection (formats are kept). */
  clearSelection() {
    const rg = this._selClipped();
    const sheet = this._sheet();
    const changes = [];
    this._forEachExisting(sheet, rg, (addr, cell, r) => {
      if (this._isHidden(r)) return; // like Excel: filtered-out rows are left alone
      const next = cellFromInput("", cell);
      changes.push({ addr, after: next });
    });
    if (changes.length > 1e6) { this.onStatus("Selection too large to clear at once"); return; }
    this._applyChanges(changes, "Clear");
  }

  /** Visit existing cells inside a range (fast for sparse sheets). */
  _forEachExisting(sheet, rg, fn) {
    const area = (rg.r2 - rg.r1 + 1) * (rg.c2 - rg.c1 + 1);
    const keys = Object.keys(sheet.cells);
    if (area > keys.length * 2) {
      for (const addr of keys) {
        const rc = addrToRC(addr);
        if (rc && rc.r >= rg.r1 && rc.r <= rg.r2 && rc.c >= rg.c1 && rc.c <= rg.c2) fn(addr, sheet.cells[addr], rc.r, rc.c);
      }
    } else {
      for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) {
        const addr = rcToAddr(r, c);
        const cell = sheet.cells[addr];
        if (cell) fn(addr, cell, r, c);
      }
    }
  }

  /** Ctrl+D / Ctrl+R: copy the first row/column of the selection into the rest. */
  _fill(dir) {
    const rg = this._selClipped();
    const sheet = this._sheet();
    const changes = [];
    if (dir === "down") {
      const src = rg.r1 === rg.r2 ? rg.r1 - 1 : rg.r1;
      if (src < 0) return;
      for (let c = rg.c1; c <= rg.c2; c++) {
        const s = sheet.cells[rcToAddr(src, c)];
        for (let r = Math.max(src + 1, rg.r1); r <= rg.r2; r++) changes.push({ addr: rcToAddr(r, c), after: this._copyCell(s, r - src, 0) });
      }
    } else {
      const src = rg.c1 === rg.c2 ? rg.c1 - 1 : rg.c1;
      if (src < 0) return;
      for (let r = rg.r1; r <= rg.r2; r++) {
        const s = sheet.cells[rcToAddr(r, src)];
        for (let c = Math.max(src + 1, rg.c1); c <= rg.c2; c++) changes.push({ addr: rcToAddr(r, c), after: this._copyCell(s, 0, c - src) });
      }
    }
    this._applyChanges(changes, dir === "down" ? "Fill down" : "Fill right");
  }

  _copyCell(cell, dr, dc) {
    if (!cell) return null;
    const out = clone(cell);
    if (out.f) {
      out.f = translateFormula(out.f, dr, dc);
      out.t = "f";
      out.v = null;
    }
    return out;
  }

  _insertNow(time) {
    const d = new Date();
    const cell = time
      ? { t: "n", v: (d.getHours() * 3600 + d.getMinutes() * 60) / 86400, z: "h:mm" }
      : { t: "d", v: serialToIso(ymdToSerial(d.getFullYear(), d.getMonth() + 1, d.getDate())), z: "yyyy-mm-dd" };
    const prev = this._curCell();
    if (prev && prev.s) cell.s = { ...prev.s };
    const m = this._mergeAt(this.cur.r, this.cur.c);
    this._applyChanges([{ addr: rcToAddr(m ? m.r1 : this.cur.r, m ? m.c1 : this.cur.c), after: cell }], time ? "Insert time" : "Insert date");
  }

  /** Apply a function to every cell of the selection (creating format-only cells for blanks). */
  _mapSelection(label, fn) {
    const rg = this._selClipped();
    const area = (rg.r2 - rg.r1 + 1) * (rg.c2 - rg.c1 + 1);
    const sheet = this._sheet();
    const changes = [];
    if (area <= 50000) {
      for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) {
        const addr = rcToAddr(r, c);
        const cell = sheet.cells[addr];
        const next = fn(cell ? clone(cell) : { t: "s", v: null });
        changes.push({ addr, after: next && (next.v != null || next.f || next.s || next.z) ? next : null });
      }
    } else {
      this._forEachExisting(sheet, rg, (addr, cell) => changes.push({ addr, after: fn(clone(cell)) }));
    }
    this._applyChanges(changes, label);
  }

  _toggleStyle(key) {
    const cur = this._curCell();
    const on = !(cur && cur.s && cur.s[key]);
    this._applyStyle({ [key]: on });
  }

  /** Merge style properties into the selection; false/"" removes a property. */
  _applyStyle(patch) {
    this._mapSelection("Format", (cell) => {
      const s = { ...(cell.s || {}) };
      for (const [k, v] of Object.entries(patch)) {
        if (v === false || v === "" || v == null) delete s[k]; else s[k] = v;
      }
      if (Object.keys(s).length) cell.s = s; else delete cell.s;
      return cell;
    });
  }

  _applyNumberFormat(z) {
    this._mapSelection("Number format", (cell) => {
      if (!z) delete cell.z; else cell.z = z;
      if (cell.f) return cell; // formula results are re-typed by recalc
      if (cell.t === "n" && z && isDateOnlyOrDateTimeFormat(z) && typeof cell.v === "number" && cell.v >= 1) {
        cell.t = "d"; cell.v = serialToIso(cell.v);
      } else if (cell.t === "d" && !(z && isDateFormat(z))) {
        const serial = isoToSerial(cell.v);
        if (serial != null) { cell.t = "n"; cell.v = serial; }
      } else if (cell.t === "s" && z && z !== "@" && typeof cell.v === "string") {
        // "1234" stored as text becomes a number when given a number format
        const p = parseCellInput(cell.v);
        if (p && (p.t === "n" || p.t === "d")) { cell.t = p.t; cell.v = p.v; }
      }
      return cell;
    });
  }

  _clearFormatting() {
    this._mapSelection("Clear formatting", (cell) => {
      delete cell.s;
      delete cell.z;
      if (cell.t === "d") {
        const serial = isoToSerial(cell.v);
        if (serial != null) { cell.t = "n"; cell.v = serial; }
      }
      return cell;
    });
  }

  _toolbarAction(act) {
    if (act !== "undo" && act !== "redo") this._commitAll();
    switch (act) {
      case "undo": this.undo(); break;
      case "redo": this.redo(); break;
      case "b": case "i": case "u": this._toggleStyle(act); break;
      case "align-left": case "align-center": case "align-right": {
        const a = act.slice(6);
        const cur = this._curCell();
        this._applyStyle({ align: cur && cur.s && cur.s.align === a ? false : a });
        break;
      }
      case "clearfmt": this._clearFormatting(); break;
      case "ins-row": this._structural("row", "insert"); break;
      case "del-row": this._structural("row", "delete"); break;
      case "ins-col": this._structural("col", "insert"); break;
      case "del-col": this._structural("col", "delete"); break;
      case "find": this.openFind(); return;
      case "sort-asc": this._quickSort(false); break;
      case "sort-desc": this._quickSort(true); break;
      case "sort-dialog": this._openSortDialog(); return;
      case "filter": this._toggleFilter(); break;
      case "freeze": this._showFreezeMenu(this.elToolbar.querySelector('[data-act="freeze"]')); return;
      case "chart": this._insertChart(); if (this._selChart >= 0) { this._selectChart(this._selChart); return; } break;
      default: break;
    }
    this._focusGrid();
  }

  // =====================================================================
  // Rows / columns
  // =====================================================================

  /** Insert/delete rows or columns covering the current selection. where: "before"|"after". */
  _structural(axis, op, where = "before") {
    this._commitAll();
    const rg = this._selRange();
    const isRow = axis === "row";
    let at = isRow ? rg.r1 : rg.c1;
    let n = isRow ? rg.r2 - rg.r1 + 1 : rg.c2 - rg.c1 + 1;
    if ((isRow && this.selKind === "cols") || (!isRow && this.selKind === "rows") || this.selKind === "all") {
      if (op === "delete" && this.selKind === "all") { this.onStatus("Select rows or columns to delete"); return; }
      n = 1;
      at = isRow ? this.cur.r : this.cur.c;
    }
    if (op === "insert" && where === "after") at += n;
    n = Math.min(n, isRow ? 100000 : 1000);
    const si = this._si();
    const label = `${op === "insert" ? "Insert" : "Delete"} ${n} ${isRow ? "row" : "column"}${n > 1 ? "s" : ""}`;
    if (!isRow) this._filterState.delete(this._sheet()); // criteria are keyed by column
    this._bookOp(label, () => {
      shiftCells(this._model(), si, axis, at, op === "insert" ? n : -n);
    });
    // keep a sensible selection
    if (isRow) this._setSelection({ r: at, c: 0 }, { r: op === "insert" ? at + n - 1 : at, c: MAX_COLS - 1 }, "rows", { r: at, c: this.cur.c });
    else this._setSelection({ r: 0, c: at }, { r: MAX_ROWS - 1, c: op === "insert" ? at + n - 1 : at }, "cols", { r: this.cur.r, c: at });
    this.onStatus(label);
    this._statsMsg = false;
  }

  _insertForSelection() {
    if (this.selKind === "rows") this._structural("row", "insert");
    else if (this.selKind === "cols") this._structural("col", "insert");
    else this.onStatus("Select whole rows or columns (click a header) to insert, or use the right-click menu");
  }

  _deleteForSelection() {
    if (this.selKind === "rows") this._structural("row", "delete");
    else if (this.selKind === "cols") this._structural("col", "delete");
    else this.onStatus("Select whole rows or columns (click a header) to delete, or use the right-click menu");
  }

  // =====================================================================
  // Clipboard
  // =====================================================================

  _onCopy(e, cut) {
    if (this.editing) return; // native copy/cut of the editor text
    e.preventDefault();
    const rg = this._selClipped();
    const area = (rg.r2 - rg.r1 + 1) * (rg.c2 - rg.c1 + 1);
    if (area > 2000000) { this.onStatus("Selection too large to copy"); return; }
    const sheet = this._sheet();
    const rows = [], cells = [];
    let html = '<table style="border-collapse:collapse">';
    for (let r = rg.r1; r <= rg.r2; r++) {
      if (this._isHidden(r) && rg.r1 !== rg.r2) continue; // copy only the visible (filtered) rows
      const row = [], crow = [];
      html += "<tr>";
      for (let c = rg.c1; c <= rg.c2; c++) {
        const cell = sheet.cells[rcToAddr(r, c)];
        const text = cellText(cell);
        row.push(text);
        crow.push(clone(cell) || null);
        const css = cell ? cellStyleCss(cell, cellAlign(cell)) : "";
        html += `<td${css ? ` style="${css}"` : ""}>${escapeHtmlText(text)}</td>`;
      }
      html += "</tr>";
      rows.push(row);
      cells.push(crow);
    }
    html += "</table>";
    const text = csvSerialize(rows, "\t", { eol: "\r\n", trailingNewline: true });
    e.clipboardData.setData("text/plain", text);
    e.clipboardData.setData("text/html", html);
    this.clip = { text, cells, range: { ...rg }, si: this._si(), sheetRef: this._sheet(), cut };
    this._renderOverlays();
    this.onStatus(`${cut ? "Cut" : "Copied"} ${rows.length}×${rows[0].length} cells`);
    this._statsMsg = false;
  }

  _onPaste(e) {
    if (this.editing) return; // paste into the editor text
    const dt = e.clipboardData;
    if (!dt) return;
    if (dt.files && dt.files.length) return; // let the viewer open pasted files
    const text = dt.getData("text/plain");
    e.preventDefault();
    if (!text && !dt.getData("text/html")) return;
    this.pasteText(text || this._htmlToTsv(dt.getData("text/html")));
  }

  _htmlToTsv(html) {
    try {
      const doc = new DOMParser().parseFromString(html, "text/html");
      const rows = [...doc.querySelectorAll("tr")].map((tr) => [...tr.querySelectorAll("td,th")].map((td) => td.textContent));
      return rows.length ? csvSerialize(rows, "\t") : doc.body.textContent || "";
    } catch { return ""; }
  }

  /** Paste TSV/plain text (or our own internal clipboard) at the selection. */
  pasteText(text) {
    const rg = this._selRange();
    const sheet = this._sheet();
    const si = this._si();
    let block; // 2D array of cells (after objects)
    const internal = this.clip && this.clip.text === text;
    if (internal) {
      block = this.clip.cells;
    } else {
      let t = String(text).replace(/\r\n?$/, "");
      if (t.endsWith("\n")) t = t.slice(0, -1);
      const rows = /[\t\n\r"]/.test(t) ? csvParse(t, "\t") : [[t]];
      block = rows.map((row) => row.map((f) => f));
    }
    const h = block.length, w = Math.max(...block.map((r) => r.length));
    if (!h || !w) return;
    if (h * w > 2000000) { this.onStatus("Clipboard data too large to paste"); return; }
    // tile into the selection when it is an exact multiple of the block
    const selH = rg.r2 - rg.r1 + 1, selW = rg.c2 - rg.c1 + 1;
    const tile = this.selKind === "cells" && selH % h === 0 && selW % w === 0 && selH * selW <= 1e6;
    const outH = tile ? selH : h, outW = tile ? selW : w;
    const top = rg.r1, left = rg.c1;
    if (top + outH > MAX_ROWS || left + outW > MAX_COLS) { this.onStatus("Paste area extends beyond the sheet"); return; }
    const changes = [];
    for (let i = 0; i < outH; i++) {
      for (let j = 0; j < outW; j++) {
        const r = top + i, c = left + j;
        const addr = rcToAddr(r, c);
        const src = block[i % h][j % w];
        let next;
        if (internal) {
          const sr = this.clip.range.r1 + (i % h), sc = this.clip.range.c1 + (j % w);
          next = this._copyCell(src, r - sr, c - sc);
        } else {
          next = src == null || src === "" ? cellFromInput("", sheet.cells[addr]) : this._prepareEntry(String(src), sheet.cells[addr]);
          if (next === sheet.cells[addr]) continue;
        }
        changes.push({ addr, after: next });
      }
    }
    if (internal && this.clip.cut && this.clip.si === si && this.clip.sheetRef === sheet) {
      // move: clear the source cells that are not overwritten
      const src = this.clip.range;
      const targets = new Set(changes.map((c) => c.addr));
      for (let r = src.r1; r <= src.r2; r++) for (let c = src.c1; c <= src.c2; c++) {
        const addr = rcToAddr(r, c);
        if (!targets.has(addr) && sheet.cells[addr]) changes.push({ addr, after: null });
      }
      this.clip = null;
    }
    this._applyChanges(changes, "Paste");
    this._setSelection({ r: top, c: left }, { r: top + outH - 1, c: left + outW - 1 }, "cells", { r: top, c: left });
    this.onStatus(`Pasted ${outH}×${outW} cells`);
    this._statsMsg = false;
  }

  // =====================================================================
  // Context menus
  // =====================================================================

  _onContextMenu(e, where) {
    if (this.editing && e.target === this.elEditor) return; // native menu inside the editor
    e.preventDefault();
    const rg = this._selRange();
    const nr = rg.r2 - rg.r1 + 1, nc = rg.c2 - rg.c1 + 1;
    const rows = this.selKind === "cols" || this.selKind === "all" ? 1 : Math.min(nr, 100000);
    const cols = this.selKind === "rows" || this.selKind === "all" ? 1 : Math.min(nc, 1000);
    const items = [
      ["cut", "Cut", "Ctrl+X"], ["copy", "Copy", "Ctrl+C"], ["paste", "Paste", "Ctrl+V"], null,
    ];
    if (where !== "col") {
      items.push(["ins-row-above", `Insert ${rows > 1 ? rows + " rows" : "row"} above`]);
      items.push(["ins-row-below", `Insert ${rows > 1 ? rows + " rows" : "row"} below`]);
      items.push(["del-row", `Delete ${rows > 1 ? rows + " rows" : "row"}`]);
    }
    if (where !== "row") {
      if (where !== "col") items.push(null);
      items.push(["ins-col-left", `Insert ${cols > 1 ? cols + " columns" : "column"} left`]);
      items.push(["ins-col-right", `Insert ${cols > 1 ? cols + " columns" : "column"} right`]);
      items.push(["del-col", `Delete ${cols > 1 ? cols + " columns" : "column"}`]);
    }
    items.push(null, ["clear", "Clear contents", "Del"], ["clearfmt", "Clear formatting"]);
    if (where === "col") items.push(["autofit", "Fit column width"]);
    if (where === "cell") {
      items.push(null, ["sort-asc", "Sort A to Z"], ["sort-desc", "Sort Z to A"], ["sort-dialog", "Custom sort…"],
        ["filter", this._sheet().autoFilter ? "Remove filter" : "Filter", "Ctrl+Shift+L"], ["chart", "Insert chart"]);
    }
    this._showMenu(e.clientX, e.clientY, items, { where });
  }

  _onTabContextMenu(e) {
    const tab = e.target.closest(".sheet-tab");
    if (!tab) return;
    e.preventDefault();
    const i = +tab.dataset.index;
    const m = this._model();
    const s = m.sheets[i];
    const hidden = m.sheets.map((x, k) => [x, k]).filter(([x]) => x.hidden);
    const items = [
      ["tab-rename", "Rename…"], ["tab-dup", "Duplicate"],
      ["tab-left", "Move left", "", i === 0], ["tab-right", "Move right", "", i === m.sheets.length - 1],
      null,
      s.hidden ? ["tab-unhide", "Unhide"] : ["tab-hide", "Hide", "", m.sheets.filter((x) => !x.hidden).length <= 1],
      ...hidden.filter(([, k]) => k !== i).map(([x, k]) => [`tab-unhide-${k}`, `Unhide "${x.name}"`]),
      null,
      ["tab-delete", "Delete", "", m.sheets.length <= 1],
    ];
    this._showMenu(e.clientX, e.clientY, items, { tab: i });
  }

  _showMenu(x, y, items, ctx) {
    this._menuCtx = ctx;
    let html = "";
    for (const it of items) {
      if (!it) { html += '<div class="xv-menu-sep" role="separator"></div>'; continue; }
      const [act, label, hint, disabled] = it;
      html += `<button type="button" role="menuitem" data-act="${act}"${disabled ? " disabled" : ""}><span>${escapeHtmlText(label)}</span>${hint ? `<kbd>${escapeHtmlText(hint)}</kbd>` : ""}</button>`;
    }
    const menu = this.elMenu;
    menu.innerHTML = html;
    menu.hidden = false;
    const wr = this.elWrap.getBoundingClientRect();
    const mw = menu.offsetWidth || 220, mh = menu.offsetHeight || 300;
    const vw = this.win.innerWidth, vh = this.win.innerHeight;
    // The menu lives inside the (clipped) grid frame: keep it inside the frame.
    menu.style.left = `${clamp(x, Math.max(4, wr.left + 2), Math.min(vw, wr.right) - mw - 4) - wr.left}px`;
    menu.style.top = `${clamp(y, Math.max(4, wr.top + 2), Math.min(vh, wr.bottom) - mh - 4) - wr.top}px`;
    const first = menu.querySelector("button:not([disabled])");
    if (first) first.focus({ preventScroll: true });
  }

  _hideMenu() {
    if (this.elMenu && !this.elMenu.hidden) {
      this.elMenu.hidden = true;
      this.elMenu.textContent = "";
    }
  }

  _onMenuClick(e) {
    const b = e.target.closest("button[data-act]");
    if (!b || b.disabled) return;
    const act = b.dataset.act;
    const ctx = this._menuCtx || {};
    this._hideMenu();
    switch (act) {
      case "cut": case "copy":
        this._focusGrid();
        if (!this.doc.execCommand(act)) this.onStatus(`Use Ctrl+${act === "cut" ? "X" : "C"} to ${act}`);
        return;
      case "paste":
        this._focusGrid();
        if (navigator.clipboard && navigator.clipboard.readText) {
          navigator.clipboard.readText().then((t) => this.pasteText(t), () => this.onStatus("Use Ctrl+V to paste"));
        } else this.onStatus("Use Ctrl+V to paste");
        return;
      case "ins-row-above": this._structural("row", "insert", "before"); break;
      case "ins-row-below": this._structural("row", "insert", "after"); break;
      case "del-row": this._structural("row", "delete"); break;
      case "ins-col-left": this._structural("col", "insert", "before"); break;
      case "ins-col-right": this._structural("col", "insert", "after"); break;
      case "del-col": this._structural("col", "delete"); break;
      case "clear": this.clearSelection(); break;
      case "clearfmt": this._clearFormatting(); break;
      case "autofit": {
        const rg = this._selRange();
        for (let c = rg.c1; c <= Math.min(rg.c2, rg.c1 + 50); c++) this._autofit(c);
        break;
      }
      case "sort-asc": this._quickSort(false); break;
      case "sort-desc": this._quickSort(true); break;
      case "sort-dialog": this._openSortDialog(); return;
      case "filter": this._toggleFilter(); break;
      case "chart": this._insertChart(); if (this._selChart >= 0) { this._selectChart(this._selChart); return; } break;
      case "fz-row": this._setFreeze(1, 0); break;
      case "fz-col": this._setFreeze(0, 1); break;
      case "fz-sel": this._setFreeze(this.cur.r, this.cur.c); break;
      case "fz-off": this._setFreeze(0, 0); break;
      case "tab-rename": this._renameSheetUi(ctx.tab); return;
      case "tab-dup": this._duplicateSheet(ctx.tab); break;
      case "tab-left": this._moveSheet(ctx.tab, -1); break;
      case "tab-right": this._moveSheet(ctx.tab, 1); break;
      case "tab-hide": this._setHidden(ctx.tab, true); break;
      case "tab-unhide": this._setHidden(ctx.tab, false); break;
      case "tab-delete": this._deleteSheet(ctx.tab); break;
      default:
        if (act.startsWith("tab-unhide-")) this._setHidden(+act.slice(11), false);
        break;
    }
    this._focusGrid();
  }

  // =====================================================================
  // Sheets
  // =====================================================================

  _onTabClick(e) {
    const x = e.target.closest(".tab-x");
    if (x) {
      e.stopPropagation();
      this._deleteSheet(+x.dataset.del);
      return;
    }
    const tab = e.target.closest(".sheet-tab");
    if (!tab || tab.querySelector("input")) return;
    this._switchSheet(+tab.dataset.index);
    this._focusGrid();
  }

  _onTabDblClick(e) {
    const tab = e.target.closest(".sheet-tab");
    if (!tab || e.target.closest(".tab-x")) return;
    this._renameSheetUi(+tab.dataset.index);
  }

  _switchSheet(i) {
    const m = this._model();
    if (!m.sheets[i]) return;
    this._commitAll();
    if (i === this._si()) return;
    m.active = i;
    this._resetSheetState();
    this._renderTabs();
    this._renderGrid(true);
    this._syncBar();
    this._updateToolbar();
    this.onStats();
  }

  addSheet() {
    this._commitAll();
    const m = this._model();
    let n = m.sheets.length + 1;
    let name = `Sheet${n}`;
    while (m.sheets.some((s) => s.name.toLowerCase() === name.toLowerCase())) name = `Sheet${++n}`;
    this._metaOp({ kind: "addSheet", label: "Add sheet", index: m.sheets.length, sheet: { name, cells: {} }, activeTo: m.sheets.length });
    this._focusGrid();
  }

  async _deleteSheet(i) {
    this._commitAll();
    const m = this._model();
    if (m.sheets.length <= 1) { this.onStatus("A workbook must contain at least one sheet"); return; }
    const s = m.sheets[i];
    if (!s) return;
    const hasData = Object.keys(s.cells || {}).length > 0;
    if (hasData) {
      const yes = await this.confirm(`The sheet "${s.name}" contains data. You can bring it back with Undo (Ctrl+Z).`, { title: `Delete sheet "${s.name}"?`, ok: "Delete sheet" });
      // The view may have been destroyed or the sheet list changed meanwhile.
      if (!yes || this.ac.signal.aborted || m !== this._model() || m.sheets[i] !== s) return;
    }
    const wasActive = i === this._si();
    this._bookOp(`Delete sheet "${s.name}"`, () => {
      deleteSheet(m, i); // keeps the same sheet active when a sheet to its left is deleted
      if (wasActive && m.sheets[m.active || 0] && m.sheets[m.active || 0].hidden) {
        const v = m.sheets.findIndex((x) => !x.hidden);
        if (v >= 0) m.active = v;
      }
    });
  }

  _renameSheetUi(i) {
    if (!this.elTabs) return;
    const tab = this.elTabs.querySelector(`.sheet-tab[data-index="${i}"]`);
    const m = this._model();
    const s = m.sheets[i];
    if (!tab || !s) return;
    const name = tab.querySelector(".tab-name");
    const input = this.doc.createElement("input");
    input.className = "xv-tab-input";
    input.value = s.name;
    input.maxLength = 31;
    input.setAttribute("aria-label", "Sheet name");
    name.replaceWith(input);
    input.focus();
    input.select();
    let done = false;
    const finish = (commit) => {
      if (done) return;
      done = true;
      if (commit) {
        const nn = input.value.trim();
        if (nn !== s.name) {
          const err = validateSheetName(nn, m.sheets, i);
          if (err) {
            this.onStatus(err);
            this._statsMsg = false;
            done = false;
            input.setCustomValidity(err);
            input.reportValidity && input.reportValidity();
            input.focus();
            return;
          }
          this._bookOp("Rename sheet", () => renameSheet(m, i, nn));
          this._focusGrid();
          return;
        }
      }
      this._renderTabs();
      this._focusGrid();
    };
    input.addEventListener("keydown", (e) => {
      e.stopPropagation();
      if (e.key === "Enter") { e.preventDefault(); finish(true); }
      else if (e.key === "Escape") { e.preventDefault(); finish(false); }
    }, { signal: this.ac.signal });
    input.addEventListener("input", () => input.setCustomValidity(""), { signal: this.ac.signal });
    input.addEventListener("blur", () => finish(true), { signal: this.ac.signal });
  }

  _duplicateSheet(i) {
    const m = this._model();
    const s = m.sheets[i];
    const name = sanitizeSheetName(`${s.name} (2)`, m.sheets);
    this._bookOp("Duplicate sheet", () => {
      const copy = clone(s);
      copy.name = name;
      delete copy.hidden;
      m.sheets.splice(i + 1, 0, copy);
      m.active = i + 1;
    });
  }

  _moveSheet(i, d) {
    const m = this._model();
    const j = i + d;
    if (j < 0 || j >= m.sheets.length) return;
    const order = m.sheets.map((_, k) => k);
    const [moved] = order.splice(i, 1);
    order.splice(j, 0, moved);
    this._metaOp({ kind: "move", label: "Move sheet", from: i, to: j, activeTo: order.indexOf(this._si()) });
  }

  _setHidden(i, hidden) {
    const m = this._model();
    const s = m.sheets[i];
    if (!s) return;
    if (hidden && m.sheets.filter((x) => !x.hidden).length <= 1) { this.onStatus("At least one sheet must stay visible"); return; }
    let activeTo = this._si();
    if (hidden && i === activeTo) activeTo = m.sheets.findIndex((x, k) => k !== i && !x.hidden);
    if (!hidden) activeTo = i;
    this._metaOp({ kind: "meta", label: hidden ? "Hide sheet" : "Unhide sheet", si: i, key: "hidden", before: !!s.hidden, after: hidden, activeTo });
  }

  // =====================================================================
  // Find
  // =====================================================================

  /** Open the find panel (replace=true shows the Replace row, like Ctrl+H). */
  openFind(replace = false) {
    this._commitAll();
    this.elFind.hidden = false;
    this._setFindMode(replace);
    this.elFindInput.focus();
    this.elFindInput.select();
    if (this.elFindInput.value) this._runFind(false);
  }

  closeFind() {
    this.elFind.hidden = true;
    this._matches = [];
    this._focusGrid();
  }

  _runFind(jump) {
    const query = this.elFindInput.value;
    const opts = this._findOpts();
    this._matches = [];
    this._matchIdx = -1;
    if (!query.trim()) { this.elFindCount.textContent = ""; return; }
    const m = this._model();
    const sheets = opts.within === "book" ? m.sheets.map((_, i) => i).filter((i) => !m.sheets[i].hidden || i === this._si()) : [this._si()];
    for (const si of sheets) {
      const found = [];
      for (const [addr, cell] of Object.entries(m.sheets[si].cells || {})) {
        if (!this._isMatch(cell, query, opts)) continue;
        const rc = addrToRC(addr);
        if (rc && !(si === this._si() && this._isHidden(rc.r))) found.push({ si, r: rc.r, c: rc.c });
        if (found.length >= 100000) break;
      }
      found.sort((a, b) => a.r - b.r || a.c - b.c);
      this._matches.push(...found);
      if (this._matches.length >= 100000) break;
    }
    if (!this._matches.length) { this.elFindCount.textContent = "No matches"; return; }
    // first match at/after the active cell
    const si0 = this._si();
    let idx = this._matches.findIndex((x) => x.si > si0 || (x.si === si0 && (x.r > this.cur.r || (x.r === this.cur.r && x.c >= this.cur.c))));
    if (idx < 0) idx = 0;
    this._matchIdx = idx - 1;
    if (jump) this._findStep(1, true);
    else this.elFindCount.textContent = `${this._matches.length} match${this._matches.length === 1 ? "" : "es"}`;
  }

  _findStep(d, keepFocus = true, quiet = false) {
    if (!this._matches || !this._matches.length) { this._runFind(false); if (!this._matches || !this._matches.length) return; }
    const n = this._matches.length;
    this._matchIdx = ((this._matchIdx ?? -1) + d + n) % n;
    const m = this._matches[this._matchIdx];
    if (m.si !== this._si()) this._switchSheet(m.si);
    this._setSelection(m, m, "cells");
    this.elFindCount.textContent = `${this._matchIdx + 1} of ${n}`;
    if (keepFocus && !quiet) this.elFindInput.focus();
  }

  // =====================================================================
  // Notices and custom undo steps
  // =====================================================================

  /** A short message over the grid (and in the status bar). */
  _notice(msg, kind = "info") {
    this.onStatus(msg);
    this._statsMsg = false;
    const el = this.elNotice;
    el.textContent = msg;
    el.dataset.kind = kind;
    el.hidden = false;
    clearTimeout(this._noticeTimer);
    this._noticeTimer = setTimeout(() => { if (!this._destroyed) el.hidden = true; }, kind === "warn" ? 7000 : 4000);
  }

  /**
   * An undoable step with its own redo/undo functions (view-state changes
   * such as filter criteria, possibly together with a model change).
   * entry: {label, redo(), undo(), dirty?:boolean}
   */
  _customOp(entry) {
    entry.kind = "custom";
    entry.sel = this._selState();
    entry.redo();
    this._pushUndo(entry);
    this._afterCustom(entry);
  }

  _afterCustom(entry) {
    this._computeHidden();
    if (this._isHidden(this.cur.r)) {
      // the active cell was filtered out: move to the nearest visible row
      const r = this._visibleRow(this.cur.r);
      this.cur = { r, c: this.cur.c };
      this.anchor = { ...this.cur };
      this.focusCell = { ...this.cur };
      this.selKind = "cells";
    }
    this._layoutCols();
    this._win = null;
    this._renderCharts();
    this._renderGrid(true);
    this._syncBar();
    this._updateToolbar();
    if (entry.dirty) this.onDirty(true);
    this.onStats();
  }

  // =====================================================================
  // Sort
  // =====================================================================

  /** The range a sort applies to: {rg, header} or null. */
  _sortTarget() {
    const sheet = this._sheet();
    const sel = this._selClipped();
    const multi = sel.r1 !== sel.r2 || sel.c1 !== sel.c2;
    const af = this._afRange();
    const inAf = af && this.cur.r >= af.r1 && this.cur.r <= af.r2 && this.cur.c >= af.c1 && this.cur.c <= af.c2;
    if (inAf && (!multi || (sel.r1 >= af.r1 && sel.r2 <= af.r2 && sel.c1 >= af.c1 && sel.c2 <= af.c2 && sel.r1 === sel.r2))) {
      return { rg: { ...af }, header: true };
    }
    if (multi) return { rg: sel, header: detectHeaderRow(sheet, sel) };
    const reg = currentRegion(sheet, this.cur.r, this.cur.c);
    if (reg.r1 === reg.r2 && reg.c1 === reg.c2 && cellIsEmpty(this._cellAt(reg.r1, reg.c1))) return null;
    return { rg: reg, header: detectHeaderRow(sheet, reg) };
  }

  _quickSort(desc) {
    this._commitAll();
    const t = this._sortTarget();
    if (!t) { this._notice("Click a cell inside your data first, then sort."); return; }
    const col = clamp(this.cur.c, t.rg.c1, t.rg.c2);
    this._doSort(t.rg, [{ col, desc }], t.header, desc ? "Sort Z to A" : "Sort A to Z");
  }

  _doSort(rg, keys, header, label) {
    const sheet = this._sheet();
    const res = sortRange(sheet, rg, keys, { header, translate: (f, dr, dc) => translateFormula(f, dr, dc) });
    if (res.error) { this._notice(res.error, "warn"); return false; }
    this._applyChanges(res.changes, label);
    if (this._filterState.get(sheet)) {
      this._computeHidden();
      this._layoutCols();
      this._renderGrid(true);
    }
    const n = rg.r2 - rg.r1 + (header ? 0 : 1);
    const by = keys.map((k) => this._colLabel(k.col, rg, header)).join(", then ");
    this._notice(`Sorted ${n} row${n === 1 ? "" : "s"} (${rangeToRef(rg)}) by ${by}${header ? "; the header row stayed on top" : ""}. Undo with Ctrl+Z.`);
    return true;
  }

  _colLabel(c, rg, header) {
    const t = header ? filterText(this._cellAt(rg.r1, c)) : "";
    return t ? `"${t}"` : `column ${this._col(c)}`;
  }

  _openSortDialog() {
    this._commitAll();
    const t = this._sortTarget();
    if (!t) { this._notice("Click a cell inside your data first, then sort."); return; }
    const rg = t.rg;
    let header = t.header;
    const levels = [{ col: clamp(this.cur.c, rg.c1, rg.c2), desc: false }];
    const d = this.doc;
    const back = d.createElement("div");
    back.className = "xv-dlg-back";
    back.innerHTML = `<div class="xv-dlg" role="dialog" aria-modal="true" aria-labelledby="xv-sort-title">
      <h3 id="xv-sort-title">Sort ${rangeToRef(rg)}</h3>
      <label class="xv-chk"><input type="checkbox" class="xv-sd-hdr"> My data has headers (keep the first row on top)</label>
      <div class="xv-sd-levels" role="list"></div>
      <button type="button" class="xv-btn xv-sd-add" title="Sort by one more column when values are equal">+ Add level</button>
      <p class="xv-dlg-note">Numbers come before text, blanks always go last. Formulas move with their rows.</p>
      <div class="xv-dlg-btns"><button type="button" class="xv-btn xv-sd-cancel">Cancel</button><button type="button" class="xv-btn xv-primary xv-sd-ok">Sort</button></div>
    </div>`;
    this.elWrap.appendChild(back);
    this._dialog = back;
    const hdr = back.querySelector(".xv-sd-hdr");
    hdr.checked = header;
    const list = back.querySelector(".xv-sd-levels");
    const colOpts = () => {
      let o = "";
      for (let c = rg.c1; c <= rg.c2; c++) {
        const name = header ? filterText(this._cellAt(rg.r1, c)) : "";
        o += `<option value="${c}">${escapeHtmlText(name ? `${name} (column ${this._col(c)})` : `Column ${this._col(c)}`)}</option>`;
      }
      return o;
    };
    const draw = () => {
      list.innerHTML = levels.map((lv, i) => `<div class="xv-sd-level" role="listitem">
        <span class="xv-sd-by">${i ? "Then by" : "Sort by"}</span>
        <select class="xv-sd-col" data-i="${i}" aria-label="${i ? "Then by column" : "Sort by column"}">${colOpts()}</select>
        <select class="xv-sd-ord" data-i="${i}" aria-label="Order"><option value="asc">A to Z (smallest first)</option><option value="desc">Z to A (largest first)</option></select>
        <button type="button" class="xv-btn xv-sd-del" data-i="${i}" title="Remove this level" aria-label="Remove level"${levels.length === 1 ? " disabled" : ""}>&#215;</button>
      </div>`).join("");
      list.querySelectorAll(".xv-sd-col").forEach((s) => { s.value = String(levels[+s.dataset.i].col); });
      list.querySelectorAll(".xv-sd-ord").forEach((s) => { s.value = levels[+s.dataset.i].desc ? "desc" : "asc"; });
    };
    draw();
    const close = () => { back.remove(); if (this._dialog === back) this._dialog = null; this._focusGrid(); };
    const ok = () => {
      close();
      this._doSort(rg, levels.map((l) => ({ ...l })), header, "Sort");
    };
    const sig = { signal: this.ac.signal };
    hdr.addEventListener("change", () => { header = hdr.checked; draw(); }, sig);
    list.addEventListener("change", (e) => {
      const i = +e.target.dataset.i;
      if (e.target.classList.contains("xv-sd-col")) levels[i].col = +e.target.value;
      if (e.target.classList.contains("xv-sd-ord")) levels[i].desc = e.target.value === "desc";
    }, sig);
    list.addEventListener("click", (e) => {
      const b = e.target.closest(".xv-sd-del");
      if (b && levels.length > 1) { levels.splice(+b.dataset.i, 1); draw(); }
    }, sig);
    back.querySelector(".xv-sd-add").addEventListener("click", () => {
      if (levels.length >= Math.min(8, rg.c2 - rg.c1 + 1)) return;
      const used = new Set(levels.map((l) => l.col));
      let c = rg.c1;
      while (used.has(c) && c < rg.c2) c++;
      levels.push({ col: c, desc: false });
      draw();
      const sels = list.querySelectorAll(".xv-sd-col");
      sels[sels.length - 1].focus();
    }, sig);
    back.querySelector(".xv-sd-cancel").addEventListener("click", close, sig);
    back.querySelector(".xv-sd-ok").addEventListener("click", ok, sig);
    back.addEventListener("keydown", (e) => {
      e.stopPropagation();
      if (e.key === "Escape") { e.preventDefault(); close(); }
      else if (e.key === "Enter" && e.target.tagName !== "BUTTON") { e.preventDefault(); ok(); }
    }, sig);
    back.addEventListener("mousedown", (e) => { if (e.target === back) { e.preventDefault(); close(); } }, sig);
    list.querySelector(".xv-sd-col").focus();
  }

  // =====================================================================
  // AutoFilter
  // =====================================================================

  _afRange() {
    const sheet = this._sheet();
    return sheet && sheet.autoFilter ? parseRange(sheet.autoFilter) : null;
  }

  _toggleFilter() {
    this._commitAll();
    const si = this._si();
    const sheet = this._sheet();
    const m = this._model();
    const beforeRef = sheet.autoFilter || null;
    const beforeCrit = clone(this._filterState.get(sheet) || null);
    let afterRef = null;
    if (!beforeRef) {
      const sel = this._selClipped();
      let rg = sel.r1 !== sel.r2 || sel.c1 !== sel.c2 ? sel : currentRegion(sheet, this.cur.r, this.cur.c);
      if (rg.r1 === rg.r2 && rg.c1 === rg.c2 && cellIsEmpty(this._cellAt(rg.r1, rg.c1))) {
        this._notice("Click a cell inside your table (with a header row) first, then turn on the filter.");
        return;
      }
      if (rg.r2 === rg.r1) rg = { ...rg, r2: rg.r1 + 1 };
      afterRef = rangeToRef(rg);
    }
    const apply = (ref, crit) => {
      const sh = m.sheets[si];
      if (!sh) return;
      if (ref) sh.autoFilter = ref; else delete sh.autoFilter;
      if (crit && Object.keys(crit).length) this._filterState.set(sh, clone(crit)); else this._filterState.delete(sh);
    };
    this._customOp({
      label: afterRef ? "Turn on filter" : "Remove filter",
      dirty: true,
      redo: () => apply(afterRef, null),
      undo: () => apply(beforeRef, beforeCrit),
    });
    if (afterRef) this._notice(`Filter on for ${afterRef}: use the arrow buttons in the header row (or Alt+Down) to sort and filter.`);
    else this._notice("Filter removed; all rows are visible again.");
  }

  /** Replace the filter criteria of the active sheet (undoable view change). */
  _setCriteria(next, label) {
    const si = this._si();
    const m = this._model();
    const sheet = this._sheet();
    const before = clone(this._filterState.get(sheet) || null);
    const after = next && Object.keys(next).length ? clone(next) : null;
    const apply = (crit) => {
      const sh = m.sheets[si];
      if (!sh) return;
      if (crit) this._filterState.set(sh, clone(crit)); else this._filterState.delete(sh);
    };
    this._customOp({ label, redo: () => apply(after), undo: () => apply(before) });
    const af = this._afRange();
    const shown = af ? af.r2 - af.r1 - this._hidden.length : 0;
    if (af) this._notice(`${shown} of ${af.r2 - af.r1} rows shown${this._hidden.length ? " (row numbers in blue are filtered results)" : ""}.`);
  }

  _closeFilterPanel(focus = false) {
    if (this.elFilterPanel) {
      this.elFilterPanel.remove();
      this.elFilterPanel = null;
      if (focus) this._focusGrid();
    }
  }

  /** Open the filter drop-down for column c of the AutoFilter range. */
  _openFilterPanel(c) {
    const af = this._afRange();
    if (!af || c < af.c1 || c > af.c2) return;
    if (this.elFilterPanel && +this.elFilterPanel.dataset.c === c) { this._closeFilterPanel(true); return; }
    this._closeFilterPanel();
    this._commitAll();
    const sheet = this._sheet();
    const state = this._filterState.get(sheet) || {};
    const crit = state[c] || {};
    const r1 = af.r1 + 1, r2 = af.r2;
    const values = filterValues(sheet, c, r1, r2, 5000);
    const shownValues = values.slice(0, 1000);
    const checked = crit.values ? new Set(crit.values) : null;
    let nums = 0, texts = 0;
    for (let r = r1; r <= Math.min(r2, r1 + 2000); r++) {
      const k = sortKeyOf(this._cellAt(r, c)).k;
      if (k === 0) nums++; else if (k === 1) texts++;
    }
    const numeric = nums >= texts;
    const ops = FILTER_OPS.filter(([op]) => (numeric
      ? ["eq", "ne", "gt", "ge", "lt", "le", "between", "top", "bottom", "above", "below"]
      : ["eq", "ne", "contains", "notcontains", "begins", "ends"]).includes(op));
    const name = filterText(this._cellAt(af.r1, c)) || `Column ${this._col(c)}`;
    const cond = crit.cond || {};
    const d = this.doc;
    const p = d.createElement("div");
    p.className = "xv-fpanel";
    p.dataset.c = String(c);
    p.setAttribute("role", "dialog");
    p.setAttribute("aria-label", `Filter ${name}`);
    const esc = escapeHtmlText;
    p.innerHTML = `
      <button type="button" class="xv-fp-act" data-fa="asc">${svgIcon(ICONS.sortAsc, "")}Sort A to Z${numeric ? " (smallest first)" : ""}</button>
      <button type="button" class="xv-fp-act" data-fa="desc">${svgIcon(ICONS.sortDesc, "")}Sort Z to A${numeric ? " (largest first)" : ""}</button>
      <button type="button" class="xv-fp-act" data-fa="clear"${state[c] ? "" : " disabled"}>Clear filter from "${esc(name)}"</button>
      <div class="xv-fp-sep"></div>
      <div class="xv-fp-cond">
        <label>${numeric ? "Number filter" : "Text filter"}
          <select class="xv-fp-op" aria-label="${numeric ? "Number filter" : "Text filter"}"><option value="">(none)</option>${ops.map(([op, label]) => `<option value="${op}">${esc(label)}</option>`).join("")}</select></label>
        <span class="xv-fp-args"><input class="xv-fp-a" aria-label="Value" spellcheck="false"><span class="xv-fp-and">and</span><input class="xv-fp-b" aria-label="Second value" spellcheck="false"></span>
      </div>
      <input type="search" class="xv-fp-search" placeholder="Search values" aria-label="Search values" spellcheck="false">
      <div class="xv-fp-list" role="group" aria-label="Values to show">
        <label class="xv-fp-all"><input type="checkbox" data-all="1"> <span>(Select all)</span></label>
        ${shownValues.map((v, i) => `<label class="xv-fp-item" data-t="${esc(v.toLowerCase())}"><input type="checkbox" data-i="${i}"${!checked || checked.has(v) ? " checked" : ""}> <span>${v === "" ? "<i>(Blanks)</i>" : esc(v)}</span></label>`).join("")}
        ${values.length > shownValues.length ? `<div class="xv-fp-more">Showing the first ${shownValues.length} of ${values.length} values; use Search or a ${numeric ? "number" : "text"} filter.</div>` : ""}
      </div>
      <div class="xv-fp-btns"><button type="button" class="xv-btn xv-fp-cancel">Cancel</button><button type="button" class="xv-btn xv-primary xv-fp-ok">OK</button></div>`;
    this.elWrap.appendChild(p);
    this.elFilterPanel = p;
    const opSel = p.querySelector(".xv-fp-op"), inA = p.querySelector(".xv-fp-a"), inB = p.querySelector(".xv-fp-b");
    const search = p.querySelector(".xv-fp-search");
    const all = p.querySelector("[data-all]");
    opSel.value = cond.op && ops.some(([o]) => o === cond.op) ? cond.op : "";
    inA.value = cond.a ?? "";
    inB.value = cond.b ?? "";
    const syncArgs = () => {
      const op = opSel.value;
      p.querySelector(".xv-fp-args").hidden = !op || op === "above" || op === "below";
      inB.hidden = op !== "between";
      p.querySelector(".xv-fp-and").hidden = op !== "between";
      inA.placeholder = op === "top" || op === "bottom" ? "How many (10)" : numeric ? "Number" : "Text";
    };
    syncArgs();
    const items = () => [...p.querySelectorAll(".xv-fp-item")].filter((el) => !el.hidden);
    const syncAll = () => {
      const vis = items();
      const n = vis.filter((el) => el.querySelector("input").checked).length;
      all.checked = n === vis.length && n > 0;
      all.indeterminate = n > 0 && n < vis.length;
    };
    syncAll();
    // position under the header cell's button
    const af1 = { r1: af.r1, r2: af.r1, c1: c, c2: c };
    const rect = this._viewRect(af1) || this._rectOf(af1);
    const sr = this.elScroll.getBoundingClientRect(), wr = this.elWrap.getBoundingClientRect();
    const left = sr.left - wr.left + rect.left - this.elScroll.scrollLeft + (rect.left < this.elScroll.scrollLeft ? this.elScroll.scrollLeft : 0);
    const top = sr.top - wr.top + rect.top + rect.height - this.elScroll.scrollTop + (this._fz.rows > af.r1 ? this.elScroll.scrollTop : 0);
    const pw = p.offsetWidth || 270, ph = p.offsetHeight || 400;
    p.style.left = `${clamp(left + rect.width - pw, 4, Math.max(4, wr.width - pw - 4))}px`;
    p.style.top = `${clamp(top + 2, 4, Math.max(4, wr.height - ph - 4))}px`;

    const sig = { signal: this.ac.signal };
    const apply = () => {
      const next = { ...(this._filterState.get(sheet) || {}) };
      const q = search.value.trim();
      const boxes = [...p.querySelectorAll(".xv-fp-item")];
      const allChecked = !q && boxes.every((el) => el.querySelector("input").checked) && values.length === shownValues.length;
      let vals;
      if (q) vals = boxes.filter((el) => !el.hidden && el.querySelector("input").checked).map((el) => shownValues[+el.querySelector("input").dataset.i]);
      else if (!allChecked) {
        vals = boxes.filter((el) => el.querySelector("input").checked).map((el) => shownValues[+el.querySelector("input").dataset.i]);
        if (values.length > shownValues.length) vals = vals.concat(values.slice(shownValues.length));
      }
      const op = opSel.value;
      const crit2 = {};
      if (vals) crit2.values = vals;
      if (op) {
        if ((op === "between" && (inA.value.trim() === "" || inB.value.trim() === "")) || (!["above", "below", "top", "bottom"].includes(op) && inA.value.trim() === "" && !["eq", "ne"].includes(op))) {
          inA.focus();
          inA.setCustomValidity("Enter a value");
          inA.reportValidity && inA.reportValidity();
          return;
        }
        crit2.cond = { op, a: inA.value.trim(), b: inB.value.trim() };
      }
      if (vals && !vals.length && !op) { this._notice("Select at least one value to show.", "warn"); return; }
      if (Object.keys(crit2).length) next[c] = crit2; else delete next[c];
      this._closeFilterPanel(true);
      this._setCriteria(next, `Filter "${name}"`);
    };
    p.addEventListener("click", (e) => {
      const act = e.target.closest("[data-fa]");
      if (act && !act.disabled) {
        this._closeFilterPanel(true);
        if (act.dataset.fa === "clear") {
          const next = { ...(this._filterState.get(sheet) || {}) };
          delete next[c];
          this._setCriteria(next, `Clear filter from "${name}"`);
        } else this._doSort(af, [{ col: c, desc: act.dataset.fa === "desc" }], true, act.dataset.fa === "desc" ? "Sort Z to A" : "Sort A to Z");
        return;
      }
      if (e.target.closest(".xv-fp-ok")) apply();
      else if (e.target.closest(".xv-fp-cancel")) this._closeFilterPanel(true);
    }, sig);
    p.addEventListener("change", (e) => {
      if (e.target === all) { for (const el of items()) el.querySelector("input").checked = all.checked; }
      else if (e.target === opSel) { syncArgs(); if (opSel.value) inA.focus(); }
      syncAll();
    }, sig);
    p.addEventListener("input", (e) => {
      if (e.target === search) {
        const q = search.value.trim().toLowerCase();
        for (const el of p.querySelectorAll(".xv-fp-item")) {
          el.hidden = !!q && !el.dataset.t.includes(q);
          if (q) el.querySelector("input").checked = !el.hidden;
        }
        syncAll();
      } else if (e.target === inA) inA.setCustomValidity("");
    }, sig);
    p.addEventListener("keydown", (e) => {
      e.stopPropagation();
      if (e.key === "Escape") { e.preventDefault(); this._closeFilterPanel(true); }
      else if (e.key === "Enter" && e.target.tagName === "INPUT" && e.target.type !== "checkbox") { e.preventDefault(); apply(); }
    }, sig);
    (opSel.value ? inA : search).focus();
  }

  // =====================================================================
  // Freeze panes
  // =====================================================================

  _showFreezeMenu(btn) {
    const r = btn.getBoundingClientRect();
    const frozen = this._fz.rows || this._fz.cols;
    const at = rcToAddr(this.cur.r, this.cur.c);
    this._showMenu(r.left, r.bottom + 2, [
      ["fz-row", "Freeze top row"],
      ["fz-col", "Freeze first column"],
      ["fz-sel", `Freeze panes above and left of ${at}`, "", this.cur.r === 0 && this.cur.c === 0],
      null,
      ["fz-off", "Unfreeze panes", "", !frozen],
    ], { freeze: true });
  }

  _setFreeze(rows, cols) {
    this._commitAll();
    const sheet = this._sheet();
    const after = rows || cols ? { rows, cols } : null;
    const before = clone(sheet.freeze) || null;
    if (JSON.stringify(before) === JSON.stringify(after)) return;
    const label = !after ? "Unfreeze panes" : rows && cols ? "Freeze panes" : rows ? `Freeze ${rows === 1 ? "top row" : `${rows} rows`}` : `Freeze ${cols === 1 ? "first column" : `${cols} columns`}`;
    // Frozen rows are shown from the top of the sheet, like Excel after scrolling back up.
    if (after) {
      if (rows) this.elScroll.scrollTop = 0;
      if (cols) this.elScroll.scrollLeft = 0;
    }
    this._metaOp({ kind: "meta", label, si: this._si(), key: "freeze", before, after });
    this._notice(after ? `${label}: ${rows ? `row${rows > 1 ? "s" : ""} 1${rows > 1 ? `-${rows}` : ""}` : ""}${rows && cols ? " and " : ""}${cols ? `column${cols > 1 ? "s" : ""} A${cols > 1 ? `-${this._col(cols - 1)}` : ""}` : ""} stay visible while you scroll.` : "Panes unfrozen.");
  }

  // =====================================================================
  // Fill handle
  // =====================================================================

  /** Fill from the selection rg in direction dir ("down"|"up"|"right"|"left") over n cells. */
  _fillTo(rg, dir, n, ctrl) {
    if (n <= 0) return;
    const sheet = this._sheet();
    const vertical = dir === "down" || dir === "up";
    const reverse = dir === "up" || dir === "left";
    if (vertical) n = Math.min(n, reverse ? rg.r1 : MAX_ROWS - 1 - rg.r2);
    else n = Math.min(n, reverse ? rg.c1 : MAX_COLS - 1 - rg.c2);
    if (n <= 0) return;
    const target = vertical
      ? (reverse ? { r1: rg.r1 - n, r2: rg.r1 - 1, c1: rg.c1, c2: rg.c2 } : { r1: rg.r2 + 1, r2: rg.r2 + n, c1: rg.c1, c2: rg.c2 })
      : (reverse ? { r1: rg.r1, r2: rg.r2, c1: rg.c1 - n, c2: rg.c1 - 1 } : { r1: rg.r1, r2: rg.r2, c1: rg.c2 + 1, c2: rg.c2 + n });
    if (mergesIn(sheet, { r1: Math.min(rg.r1, target.r1), r2: Math.max(rg.r2, target.r2), c1: Math.min(rg.c1, target.c1), c2: Math.max(rg.c2, target.c2) }).length) {
      this._notice("Cells cannot be filled across merged cells.", "warn");
      return;
    }
    if ((target.r2 - target.r1 + 1) * (target.c2 - target.c1 + 1) > 1000000) { this._notice("That area is too large to fill at once.", "warn"); return; }
    const changes = [];
    if (vertical) {
      for (let c = rg.c1; c <= rg.c2; c++) {
        const src = [];
        for (let r = rg.r1; r <= rg.r2; r++) src.push(this._cellAt(r, c) || null);
        const out = fillSeries(src, n, { reverse, ctrl, translate: (f, delta) => translateFormula(f, delta, 0) });
        out.forEach((cell, j) => changes.push({ addr: rcToAddr(reverse ? rg.r1 - 1 - j : rg.r2 + 1 + j, c), after: cell }));
      }
    } else {
      for (let r = rg.r1; r <= rg.r2; r++) {
        const src = [];
        for (let c = rg.c1; c <= rg.c2; c++) src.push(this._cellAt(r, c) || null);
        const out = fillSeries(src, n, { reverse, ctrl, translate: (f, delta) => translateFormula(f, 0, delta) });
        out.forEach((cell, j) => changes.push({ addr: rcToAddr(r, reverse ? rg.c1 - 1 - j : rg.c2 + 1 + j), after: cell }));
      }
    }
    this._applyChanges(changes, "Fill");
    const all = { r1: Math.min(rg.r1, target.r1), r2: Math.max(rg.r2, target.r2), c1: Math.min(rg.c1, target.c1), c2: Math.max(rg.c2, target.c2) };
    this._setSelection({ r: all.r1, c: all.c1 }, { r: all.r2, c: all.c2 }, "cells", { r: rg.r1, c: rg.c1 }, false);
    this.onStatus(`Filled ${rangeToRef(target)}${ctrl ? " (Ctrl: copy/series switched)" : ""}`);
    this._statsMsg = false;
  }

  /** Double-click on the fill handle: fill down as far as the neighbouring column has data. */
  _fillDownAuto() {
    const rg = this._selClipped();
    const probe = (c) => {
      if (c < 0 || c >= MAX_COLS) return 0;
      let r = rg.r2 + 1, n = 0;
      while (r < MAX_ROWS && n < 1000000 && this._hasValue(r, c) && !this._hasValue(r, rg.c1)) { r++; n++; }
      return n;
    };
    const n = Math.max(probe(rg.c1 - 1), probe(rg.c2 + 1));
    if (!n) { this.onStatus("Nothing to fill: double-click fills down next to a column of data"); this._statsMsg = false; return; }
    this._fillTo(rg, "down", n, false);
  }

  _fillDragTo(rc, ctrl) {
    const d = this.drag;
    const rg = d.src;
    const ext = [["down", rc.r - rg.r2], ["up", rg.r1 - rc.r], ["right", rc.c - rg.c2], ["left", rg.c1 - rc.c]]
      .filter(([, n]) => n > 0).sort((a, b) => b[1] - a[1])[0];
    d.ctrl = ctrl;
    if (!ext) { d.target = null; this.elFillPrev.hidden = true; return; }
    const [dir, n] = ext;
    d.target = { dir, n };
    const u = { ...rg };
    if (dir === "down") u.r2 = rg.r2 + n; else if (dir === "up") u.r1 = rg.r1 - n;
    else if (dir === "right") u.c2 = rg.c2 + n; else u.c1 = rg.c1 - n;
    this._ensureExtent(u.r2, u.c2);
    this.elFillPrev.hidden = false;
    this._place(this.elFillPrev, this._viewRect(u));
    this.onStatus(`Fill ${dir} ${n} cell${n === 1 ? "" : "s"} — release to fill${ctrl ? " (Ctrl held)" : "; hold Ctrl to copy instead of extending a series"}`);
    this._statsMsg = false;
  }

  // =====================================================================
  // Find & replace
  // =====================================================================

  _findOpts() {
    const q = (s) => this.elFind.querySelector(s);
    return {
      within: q(".xv-find-within").value, lookin: q(".xv-find-lookin").value,
      matchCase: q(".xv-find-case").checked, whole: q(".xv-find-whole").checked,
    };
  }

  _matchText(cell, opts) {
    if (cellIsEmpty(cell)) return null;
    return opts.lookin === "formulas" ? cellRaw(cell) : formatCell(cell).text;
  }

  _isMatch(cell, query, opts) {
    let t = this._matchText(cell, opts);
    if (t == null) return false;
    let q = query;
    if (!opts.matchCase) { t = t.toLowerCase(); q = q.toLowerCase(); }
    return opts.whole ? t === q : t.includes(q);
  }

  _setFindMode(replace) {
    this.elFind.classList.toggle("xv-find-replace", !!replace);
    this.elFind.querySelector(".xv-find-more").setAttribute("aria-expanded", replace ? "true" : "false");
  }

  /** Text replaced inside a cell's raw content; null when nothing changes. */
  _replaceIn(raw, query, rep, opts) {
    if (opts.whole) {
      const same = opts.matchCase ? raw === query : raw.toLowerCase() === query.toLowerCase();
      return same ? { text: rep, n: 1 } : null;
    }
    const re = new RegExp(query.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), opts.matchCase ? "g" : "gi");
    let n = 0;
    const text = raw.replace(re, () => { n++; return rep; });
    return n ? { text, n } : null;
  }

  _replaceOne() {
    const query = this.elFindInput.value;
    if (!query) { this.elFindInput.focus(); return; }
    const opts = this._findOpts();
    const cell = this._curCell();
    if (this._isMatch(cell, query, { ...opts, lookin: "formulas" }) || this._isMatch(cell, query, opts)) {
      const res = this._replaceIn(cellRaw(cell), query, this.elRepInput.value, opts);
      if (res) {
        const m = this._mergeAt(this.cur.r, this.cur.c);
        const r = m ? m.r1 : this.cur.r, c = m ? m.c1 : this.cur.c;
        const next = this._prepareEntry(res.text, cell);
        this._applyChanges([{ addr: rcToAddr(r, c), after: next }], "Replace");
      }
    }
    this._runFind(false);
    this._findStep(1, true, true);
  }

  _replaceAll() {
    const query = this.elFindInput.value;
    if (!query) { this.elFindInput.focus(); return; }
    this._commitAll();
    const opts = this._findOpts();
    const m = this._model();
    const sheets = opts.within === "book" ? m.sheets.map((_, i) => i) : [this._si()];
    const bySheet = new Map();
    let count = 0, cells = 0;
    for (const si of sheets) {
      const sh = m.sheets[si];
      for (const [addr, cell] of Object.entries(sh.cells || {})) {
        if (!this._isMatch(cell, query, opts) && !this._isMatch(cell, query, { ...opts, lookin: "formulas" })) continue;
        const res = this._replaceIn(cellRaw(cell), query, this.elRepInput.value, opts);
        if (!res) continue;
        const next = this._prepareEntry(res.text, cell);
        if (next === cell) continue;
        if (!bySheet.has(si)) bySheet.set(si, []);
        bySheet.get(si).push({ addr, after: next });
        count += res.n;
        cells++;
      }
    }
    if (!cells) { this.elFindCount.textContent = "No matches"; this._notice(`Nothing replaced: "${query}" was not found.`); return; }
    if (bySheet.size === 1) {
      const [si, changes] = [...bySheet.entries()][0];
      this._applyChanges(changes, "Replace all", { si });
    } else {
      this._bookOp("Replace all", () => {
        for (const [si, changes] of bySheet) for (const ch of changes) m.sheets[si].cells[ch.addr] = ch.after;
      });
    }
    this._notice(`Replaced ${count} occurrence${count === 1 ? "" : "s"} in ${cells} cell${cells === 1 ? "" : "s"}. Undo with Ctrl+Z.`);
    this._runFind(false);
  }

  // =====================================================================
  // Formula autocomplete and argument hints
  // =====================================================================

  _popup(cls) {
    const el = this.doc.createElement("div");
    el.className = cls;
    el.hidden = true;
    this.doc.body.appendChild(el);
    this._popups.push(el);
    return el;
  }

  _hideAutocomplete() {
    if (this.elAc) this.elAc.hidden = true;
    if (this.elSig) this.elSig.hidden = true;
    this._ac = null;
  }

  /** Update the function list / argument hint for the formula being typed in `input`. */
  _acUpdate(input) {
    if (this._destroyed) return;
    const v = input.value;
    const pos = input.selectionStart;
    if (!v.startsWith("=") || pos !== input.selectionEnd) { this._hideAutocomplete(); return; }
    const before = v.slice(0, pos);
    if ((before.match(/"/g) || []).length % 2) { this._hideAutocomplete(); return; }
    if (!this.elAc) { this.elAc = this._popup("xv-ac"); this.elAc.setAttribute("role", "listbox"); this.elSig = this._popup("xv-sig"); this.elSig.setAttribute("role", "tooltip"); this._bindAcPopup(); }
    const m = /(^|[=(,;+\-*/^&<>:\s{])([A-Za-z_][A-Za-z0-9_.]*)$/.exec(before);
    const nextCh = v.charAt(pos);
    let names = [];
    if (m && !/^\$/.test(m[2]) && !/[(A-Za-z0-9_.]/.test(nextCh)) {
      const pre = m[2].toUpperCase();
      names = Object.keys(FUNCTION_INFO).filter((n) => n.startsWith(pre)).sort((a, b) => a.length - b.length || (a < b ? -1 : 1)).slice(0, 12);
      if (names.length === 1 && names[0] === pre && nextCh === "(") names = [];
    }
    const anchor = this._acAnchor(input, before);
    if (names.length) {
      const prev = this._ac && this._ac.input === input ? this._ac.names[this._ac.idx] : null;
      const idx = Math.max(0, names.indexOf(prev));
      this._ac = { input, names, idx, start: pos - m[2].length, end: pos };
      this._drawAc(anchor);
      this.elSig.hidden = true;
      return;
    }
    this.elAc.hidden = true;
    this._ac = null;
    const info = functionArgIndex(before);
    const fi = info && FUNCTION_INFO[info.name];
    if (!fi) { this.elSig.hidden = true; return; }
    this.elSig.innerHTML = this._sigHtml(fi.sig, info.argIndex) + `<div class="xv-sig-desc">${escapeHtmlText(fi.desc)}</div>`;
    this.elSig.hidden = false;
    this._placePopup(this.elSig, anchor);
  }

  /** Bold the argIndex-th argument of a signature like "SUM(number1, [number2], ...)". */
  _sigHtml(sig, argIndex) {
    const m = /^([^(]*)\((.*)\)$/.exec(sig);
    if (!m) return escapeHtmlText(sig);
    const args = m[2] ? m[2].split(/,\s*/) : [];
    let k = argIndex;
    if (k >= args.length) {
      const dots = args.indexOf("...");
      // repeating arguments: highlight the repeating group's last named argument
      k = dots > 0 ? dots - 1 : args.length - 1;
    }
    return `<span class="xv-sig-fn">${escapeHtmlText(m[1])}</span>(${args.map((a, i) => (i === k ? `<b>${escapeHtmlText(a)}</b>` : escapeHtmlText(a))).join(", ")})`;
  }

  _acAnchor(input, before) {
    const r = input.getBoundingClientRect();
    if (input === this.elEditor) return { left: r.left, top: r.bottom + 2, above: r.top };
    const font = this.win.getComputedStyle ? this.win.getComputedStyle(input).font : "";
    const x = Math.min(r.width - 40, this._measure(before, font || `14px ${GRID_FONT_FAMILY}`) - input.scrollLeft + 6);
    return { left: r.left + Math.max(0, x), top: r.bottom + 2, above: r.top };
  }

  _placePopup(el, anchor) {
    const vw = this.win.innerWidth, vh = this.win.innerHeight;
    const w = el.offsetWidth || 300, h = el.offsetHeight || 100;
    el.style.left = `${clamp(anchor.left, 4, Math.max(4, vw - w - 4))}px`;
    el.style.top = `${anchor.top + h > vh - 4 ? Math.max(4, anchor.above - h - 2) : anchor.top}px`;
  }

  _drawAc(anchor) {
    const a = this._ac;
    const fi = FUNCTION_INFO[a.names[a.idx]];
    this.elAc.innerHTML = a.names.map((n, i) => `<div class="xv-ac-item${i === a.idx ? " xv-ac-on" : ""}" role="option" aria-selected="${i === a.idx}" data-name="${n}"><b>${n}</b><span>${escapeHtmlText(FUNCTION_INFO[n].desc)}</span></div>`).join("") +
      `<div class="xv-ac-sig">${escapeHtmlText(fi.sig)}<div class="xv-ac-hint">Tab or Enter to insert · Esc to close</div></div>`;
    this.elAc.hidden = false;
    this._placePopup(this.elAc, anchor || this._ac.anchor);
    this._ac.anchor = anchor || this._ac.anchor;
    const on = this.elAc.querySelector(".xv-ac-on");
    if (on && on.scrollIntoView) on.scrollIntoView({ block: "nearest" });
  }

  _bindAcPopup() {
    this.elAc.addEventListener("mousedown", (e) => {
      e.preventDefault(); // keep the caret in the editor
      const it = e.target.closest(".xv-ac-item");
      if (!it || !this._ac) return;
      this._ac.idx = this._ac.names.indexOf(it.dataset.name);
      this._acAccept();
    }, { signal: this.ac.signal });
  }

  /** Keys while the list is open: returns true when handled. */
  _acKey(e) {
    if (!this._ac || !this.elAc || this.elAc.hidden) {
      if (e.key === "Escape" && this.elSig && !this.elSig.hidden) { this.elSig.hidden = true; e.preventDefault(); return true; }
      return false;
    }
    const a = this._ac;
    switch (e.key) {
      case "ArrowDown": e.preventDefault(); a.idx = (a.idx + 1) % a.names.length; this._drawAc(); return true;
      case "ArrowUp": e.preventDefault(); a.idx = (a.idx - 1 + a.names.length) % a.names.length; this._drawAc(); return true;
      case "Tab": case "Enter":
        if (e.shiftKey || e.altKey || e.ctrlKey) return false;
        e.preventDefault(); this._acAccept(); return true;
      case "Escape": e.preventDefault(); this._hideAutocomplete(); return true;
      default: return false;
    }
  }

  _acAccept() {
    const a = this._ac;
    if (!a) return;
    const input = a.input;
    const name = a.names[a.idx];
    const v = input.value;
    const text = name + "(";
    input.value = v.slice(0, a.start) + text + v.slice(a.end);
    const caret = a.start + text.length;
    input.setSelectionRange(caret, caret);
    this._ac = null;
    this.elAc.hidden = true;
    if (input === this.elEditor) {
      this.elFormula.value = input.value;
      this._sizeEditor();
    } else this.fbDirty = true;
    input.focus();
    this._acUpdate(input);
  }

  // =====================================================================
  // Charts
  // =====================================================================

  _chartTop(ch) {
    const y = Math.max(0, Number(ch.y) || 0);
    const row = Math.floor(y / ROW_H);
    return this._ry(row) + (this._isHidden(row) ? 0 : y - row * ROW_H);
  }

  _renderCharts() {
    if (!this.elCharts) return;
    const sheet = this._sheet();
    const charts = (sheet && Array.isArray(sheet.charts)) ? sheet.charts : [];
    if (this._selChart >= charts.length) this._selChart = -1;
    if (!charts.length) { this.elCharts.innerHTML = ""; return; }
    const focusedIdx = this.doc.activeElement && this.doc.activeElement.closest && this.doc.activeElement.closest(".xv-chart") ? this._selChart : -1;
    this.elCharts.innerHTML = charts.map((ch, i) => {
      const w = Math.max(60, Math.round(Number(ch.w) || 480)), h = Math.max(40, Math.round(Number(ch.h) || 288));
      const label = `${(CHART_TYPES.find(([t]) => t === ch.type) || ["", "Column"])[1]} chart${ch.title ? ` "${ch.title}"` : ""} of ${ch.range || "?"}`;
      const handles = ["nw", "n", "ne", "e", "se", "s", "sw", "w"].map((hd) => `<span class="xv-chh xv-chh-${hd}" data-h="${hd}"></span>`).join("");
      return `<div class="xv-chart${i === this._selChart ? " xv-chart-sel" : ""}" data-i="${i}" tabindex="${i === this._selChart ? 0 : -1}" role="img" aria-label="${escapeHtmlText(label)}"` +
        ` title="${escapeHtmlText(label)}: drag to move, double-click to edit, Delete to remove" style="left:${Math.max(0, Number(ch.x) || 0)}px;top:${this._chartTop(ch)}px;width:${w}px;height:${h}px">` +
        `${chartSvg(sheet, ch, w, h)}${handles}</div>`;
    }).join("");
    if (focusedIdx >= 0) {
      const el = this.elCharts.querySelector(`.xv-chart[data-i="${focusedIdx}"]`);
      if (el) el.focus({ preventScroll: true });
    }
  }

  _scheduleCharts() {
    const sheet = this._sheet();
    if (!sheet || !Array.isArray(sheet.charts) || !sheet.charts.length || this._chartRaf) return;
    this._chartRaf = requestAnimationFrame(() => { this._chartRaf = 0; this._renderCharts(); });
  }

  _selectChart(i, focus = true) {
    this._selChart = i;
    for (const el of this.elCharts.children) {
      const on = +el.dataset.i === i;
      el.classList.toggle("xv-chart-sel", on);
      el.tabIndex = on ? 0 : -1;
      if (on && focus) el.focus({ preventScroll: true });
    }
    if (i >= 0) {
      const ch = this._sheet().charts[i];
      this.onStatus(`Chart selected (${ch.range}). Drag to move, drag a handle to resize, double-click or Enter to edit, Delete to remove, Esc to go back to the cells.`);
      this._statsMsg = false;
    }
  }

  _setCharts(charts, label, keepSel = -1) {
    const sheet = this._sheet();
    this._keepChartSel = keepSel;
    this._metaOp({ kind: "meta", label, si: this._si(), key: "charts", before: clone(sheet.charts) || null, after: charts.length ? clone(charts) : null });
  }

  _insertChart() {
    this._commitAll();
    const sheet = this._sheet();
    const sel = this._selClipped();
    let rg = sel.r1 !== sel.r2 || sel.c1 !== sel.c2 ? sel : currentRegion(sheet, this.cur.r, this.cur.c);
    if (rg.r2 - rg.r1 < 1) { this._notice("Select a table first: a header row with labels plus at least one row of numbers.", "warn"); return; }
    let numbers = 0;
    for (let r = rg.r1 + 1; r <= Math.min(rg.r2, rg.r1 + 500); r++) {
      for (let c = rg.c1; c <= rg.c2; c++) if (sortKeyOf(this._cellAt(r, c)).k === 0) numbers++;
    }
    if (!numbers) { this._notice("There are no numbers to chart in the selection. Select a header row plus rows of numbers.", "warn"); return; }
    if (rg.r2 - rg.r1 > 5000) rg = { ...rg, r2: rg.r1 + 5000 };
    const series = Math.max(1, rg.c2 - rg.c1);
    const title = series === 1 && rg.c2 > rg.c1 ? filterText(this._cellAt(rg.r1, rg.c2)) : "";
    const x = this._cx[Math.min(rg.c2 + 1, this.ncols)] + 16;
    // below frozen rows, so the chart does not start hidden behind them
    const y = Math.max(rg.r1, this._fz.rows) * ROW_H + (this._fz.rows > rg.r1 ? 8 : 0);
    const chart = { type: "column", title, range: rangeToRef(rg), x, y, w: 480, h: 288 };
    const charts = [...(sheet.charts || []), chart];
    this._setCharts(charts, "Insert chart", charts.length - 1);
    const sc = this.elScroll;
    if (x + 480 > sc.scrollLeft + sc.clientWidth) sc.scrollLeft = Math.max(0, x + 480 - sc.clientWidth + 20);
    this._notice("Chart added. Drag it to move, drag a corner to resize, double-click to change its type or title, press Delete to remove it.");
  }

  _deleteChart(i) {
    const sheet = this._sheet();
    const charts = (sheet.charts || []).slice();
    if (!charts[i]) return;
    charts.splice(i, 1);
    this._closeChartPanel();
    this._setCharts(charts, "Delete chart");
    this._focusGrid();
    this._notice("Chart deleted (Ctrl+Z brings it back).");
  }

  _onChartMouseDown(e, el) {
    if (e.button !== 0 && e.button !== 2) return;
    e.preventDefault();
    e.stopPropagation();
    this._commitAll();
    this._hideMenu();
    const i = +el.dataset.i;
    this._selectChart(i);
    if (e.button !== 0) return;
    const ch = this._sheet().charts[i];
    const h = e.target.closest(".xv-chh");
    this.drag = { kind: "chart", i, mode: h ? h.dataset.h : "move", sx: e.clientX, sy: e.clientY, orig: { ...ch }, el, top0: this._chartTop(ch) };
  }

  _chartDragMove(e) {
    const d = this.drag;
    const dx = e.clientX - d.sx, dy = e.clientY - d.sy;
    const o = d.orig;
    let x = Number(o.x) || 0, y = Number(o.y) || 0, w = Number(o.w) || 480, h = Number(o.h) || 288;
    if (d.mode === "move") { x += dx; y += dy; }
    else {
      if (d.mode.includes("e")) w += dx;
      if (d.mode.includes("s")) h += dy;
      if (d.mode.includes("w")) { x += dx; w -= dx; }
      if (d.mode.includes("n")) { y += dy; h -= dy; }
    }
    w = Math.max(120, Math.round(w)); h = Math.max(90, Math.round(h));
    x = Math.max(0, Math.round(x)); y = Math.max(0, Math.round(y));
    d.geom = { x, y, w, h };
    d.moved = d.moved || Math.abs(dx) > 2 || Math.abs(dy) > 2;
    const st = d.el.style;
    st.left = `${x}px`;
    st.top = `${d.top0 + (y - (Number(o.y) || 0))}px`;
    st.width = `${w}px`;
    st.height = `${h}px`;
    const svg = d.el.querySelector("svg");
    if (svg && d.mode !== "move") { svg.setAttribute("width", w); svg.setAttribute("height", h); svg.setAttribute("preserveAspectRatio", "none"); }
  }

  _chartDragEnd(d) {
    if (!d.moved || !d.geom) return;
    const sheet = this._sheet();
    const charts = clone(sheet.charts || []);
    if (!charts[d.i]) return;
    Object.assign(charts[d.i], d.geom);
    this._setCharts(charts, d.mode === "move" ? "Move chart" : "Resize chart", d.i);
    const el = this.elCharts.querySelector(`.xv-chart[data-i="${d.i}"]`);
    if (el) el.focus({ preventScroll: true });
  }

  _onChartKey(e) {
    const el = e.target.closest(".xv-chart");
    if (!el) return;
    const i = +el.dataset.i;
    if (e.key === "Delete" || e.key === "Backspace") { e.preventDefault(); this._deleteChart(i); }
    else if (e.key === "Escape") { e.preventDefault(); this._selectChart(-1, false); this._focusGrid(); }
    else if (e.key === "Enter" || e.key === "F2") { e.preventDefault(); this._openChartPanel(i); }
    else if ((e.ctrlKey || e.metaKey) && (e.key === "z" || e.key === "y")) { e.preventDefault(); if (e.key === "z" && !e.shiftKey) this.undo(); else this.redo(); }
    else if (e.key.startsWith("Arrow")) {
      e.preventDefault();
      const step = e.shiftKey ? 10 : 1;
      const charts = clone(this._sheet().charts || []);
      const ch = charts[i];
      if (e.key === "ArrowLeft") ch.x = Math.max(0, (Number(ch.x) || 0) - step);
      if (e.key === "ArrowRight") ch.x = (Number(ch.x) || 0) + step;
      if (e.key === "ArrowUp") ch.y = Math.max(0, (Number(ch.y) || 0) - step);
      if (e.key === "ArrowDown") ch.y = (Number(ch.y) || 0) + step;
      this._setCharts(charts, "Move chart", i);
      const n = this.elCharts.querySelector(`.xv-chart[data-i="${i}"]`);
      if (n) n.focus({ preventScroll: true });
    }
  }

  _closeChartPanel() {
    if (this.elChartPanel) { this.elChartPanel.remove(); this.elChartPanel = null; }
  }

  _openChartPanel(i) {
    this._closeChartPanel();
    const sheet = this._sheet();
    const ch = sheet.charts && sheet.charts[i];
    if (!ch) return;
    const p = this.doc.createElement("div");
    p.className = "xv-chartpanel";
    p.setAttribute("role", "dialog");
    p.setAttribute("aria-label", "Edit chart");
    p.innerHTML = `<h3>Edit chart</h3>
      <label>Title <input class="xv-cp-title" spellcheck="true" placeholder="(no title)"></label>
      <label>Type <select class="xv-cp-type">${CHART_TYPES.map(([t, l]) => `<option value="${t}">${escapeHtmlText(l)}</option>`).join("")}</select></label>
      <label>Data range <input class="xv-cp-range" spellcheck="false" title="Header row + first column as labels, e.g. A1:C10"></label>
      <p class="xv-dlg-note">First row: series names. First column: labels (x values for scatter).</p>
      <div class="xv-dlg-btns">
        <button type="button" class="xv-btn" data-cp="png" title="Download the chart as a PNG picture">PNG</button>
        <button type="button" class="xv-btn" data-cp="svg" title="Download the chart as an SVG drawing">SVG</button>
        <button type="button" class="xv-btn xv-danger" data-cp="del" title="Delete this chart (Delete)">Delete</button>
      </div>
      <div class="xv-dlg-btns"><button type="button" class="xv-btn" data-cp="close">Close</button><button type="button" class="xv-btn xv-primary" data-cp="apply">Apply</button></div>`;
    this.elWrap.appendChild(p);
    this.elChartPanel = p;
    const title = p.querySelector(".xv-cp-title"), type = p.querySelector(".xv-cp-type"), range = p.querySelector(".xv-cp-range");
    title.value = ch.title || "";
    type.value = CHART_TYPES.some(([t]) => t === ch.type) ? ch.type : "column";
    range.value = ch.range || "";
    const sig = { signal: this.ac.signal };
    const apply = () => {
      const rg = parseRange(range.value.replace(/\$/g, "").toUpperCase().replace(/^.*!/, ""));
      if (!rg || rg.r2 <= rg.r1) {
        range.setCustomValidity("Enter a range like A1:C10 (a header row plus data rows)");
        range.reportValidity && range.reportValidity();
        return false;
      }
      const charts = clone(this._sheet().charts || []);
      if (!charts[i]) return false;
      charts[i] = { ...charts[i], title: title.value.trim(), type: type.value, range: rangeToRef(rg) };
      if (!charts[i].title) delete charts[i].title;
      this._setCharts(charts, "Edit chart", i);
      return true;
    };
    p.addEventListener("click", (e) => {
      const b = e.target.closest("[data-cp]");
      if (!b) return;
      switch (b.dataset.cp) {
        case "apply": if (apply()) { this._closeChartPanel(); this._selectChart(i); } break;
        case "close": this._closeChartPanel(); this._selectChart(i); break;
        case "del": this._deleteChart(i); break;
        case "png": case "svg": this._downloadChart(i, b.dataset.cp); break;
        default: break;
      }
    }, sig);
    p.addEventListener("input", (e) => { if (e.target === range) range.setCustomValidity(""); }, sig);
    p.addEventListener("keydown", (e) => {
      e.stopPropagation();
      if (e.key === "Escape") { e.preventDefault(); this._closeChartPanel(); this._selectChart(i); }
      else if (e.key === "Enter" && e.target.tagName === "INPUT") { e.preventDefault(); if (apply()) { this._closeChartPanel(); this._selectChart(i); } }
    }, sig);
    title.focus();
    title.select();
  }

  /** Download chart i as "png" or "svg". */
  _downloadChart(i, fmt) {
    const sheet = this._sheet();
    const ch = sheet.charts && sheet.charts[i];
    if (!ch) return;
    const w = Math.round(Number(ch.w) || 480), h = Math.round(Number(ch.h) || 288);
    const svg = chartSvg(sheet, ch, w, h);
    const base = String(ch.title || "chart").replace(/[\\/:*?"<>|\u0000-\u001f]/g, "_").slice(0, 80) || "chart";
    const save = (blob, name) => {
      const a = this.doc.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = name;
      a.style.display = "none";
      this.doc.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(a.href), 30000);
      this.onStatus(`Downloaded ${name}`);
      this._statsMsg = false;
    };
    if (fmt === "svg") { save(new Blob([svg], { type: "image/svg+xml" }), `${base}.svg`); return; }
    const img = new Image();
    img.onload = () => {
      const cv = this.doc.createElement("canvas");
      cv.width = w * 2; cv.height = h * 2;
      const g = cv.getContext("2d");
      g.scale(2, 2);
      g.drawImage(img, 0, 0, w, h);
      cv.toBlob((blob) => { if (blob) save(blob, `${base}.png`); else this._notice("The picture could not be created.", "warn"); }, "image/png");
    };
    img.onerror = () => this._notice("The picture could not be created.", "warn");
    img.src = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(svg);
  }

  // =====================================================================
  // Printing
  // =====================================================================

  _preparePrint() {
    if (this._destroyed || this.root.hidden) return;
    const sheet = this._sheet();
    if (!sheet) return;
    const { html } = renderSheetHtml(sheet, { maxRows: 10000, maxCols: 200, cellBudget: 400000, minRows: 1, minCols: 1 });
    this.elPrint.innerHTML = html;
  }
}
