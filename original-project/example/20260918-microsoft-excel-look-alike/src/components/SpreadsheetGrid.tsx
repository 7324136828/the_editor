import { useEffect, useMemo } from 'react';
import type { KeyboardEvent } from 'react';
import { ChevronDown } from 'lucide-react';
import { columnLetters, rowCount } from '../data';
import { evaluateCell, formatCellValue } from '../formulas';
import type { CellFormat, CellFormatMap, CellMap, ToastFn } from '../types';

const colWidths = [122, 92, 104, 98, 118, 116, 32, 104, 116, 32, 120, 118, 100, 126];
const rowHeight = 24;

export type SpreadsheetGridProps = {
  cells: CellMap;
  formats: CellFormatMap;
  selectedCell: string;
  onSelectCell: (ref: string) => void;
  showFormulas: boolean;
  showGridlines: boolean;
  tableStyle: boolean;
  conditionalFormatting: boolean;
  dataBars: boolean;
  toast: ToastFn;
};

function colIndexOf(ref: string): number {
  const m = /^([A-Z]+)([0-9]+)$/.exec(ref);
  return m ? columnLetters.indexOf(m[1] as (typeof columnLetters)[number]) : 0;
}

function rowIndexOf(ref: string): number {
  const m = /^([A-Z]+)([0-9]+)$/.exec(ref);
  return m ? parseInt(m[2], 10) : 1;
}

function builtinNumberFormat(col: string, row: number): CellFormat['number'] {
  if ((col === 'C' || col === 'E') && row >= 2 && row <= 25) return 'currency';
  if (col === 'L' && ((row >= 3 && row <= 7) || row === 12)) return 'currency';
  if (col === 'M' && row >= 3 && row <= 7) return 'percent';
  if (row === 11 && (col === 'L' || col === 'N')) return 'percent';
  return 'general';
}

export function SpreadsheetGrid(props: SpreadsheetGridProps) {
  const {
    cells,
    formats,
    selectedCell,
    onSelectCell,
    showFormulas,
    showGridlines,
    tableStyle,
    conditionalFormatting,
    dataBars,
  } = props;

  const maxGross = useMemo(() => {
    let max = 0;
    for (let row = 2; row <= 25; row++) {
      const v = evaluateCell(`E${row}`, cells);
      if (typeof v === 'number' && v > max) max = v;
    }
    return max;
  }, [cells]);

  useEffect(() => {
    const el = document.getElementById(`cell-${selectedCell}`);
    el?.focus({ preventScroll: false });
  }, [selectedCell]);

  const move = (colIdx: number, row: number) => {
    const c = Math.min(Math.max(colIdx, 0), columnLetters.length - 1);
    const r = Math.min(Math.max(row, 1), rowCount);
    onSelectCell(`${columnLetters[c]}${r}`);
  };

  const handleKeyDown = (e: KeyboardEvent<HTMLDivElement>) => {
    const col = colIndexOf(selectedCell);
    const row = rowIndexOf(selectedCell);
    switch (e.key) {
      case 'ArrowUp':
        e.preventDefault();
        move(col, row - 1);
        break;
      case 'ArrowDown':
        e.preventDefault();
        move(col, row + 1);
        break;
      case 'ArrowLeft':
        e.preventDefault();
        move(col - 1, row);
        break;
      case 'ArrowRight':
        e.preventDefault();
        move(col + 1, row);
        break;
      case 'Enter':
        e.preventDefault();
        move(col, row + 1);
        break;
      case 'Home':
        e.preventDefault();
        if (e.ctrlKey || e.metaKey) move(0, 1);
        else move(0, row);
        break;
      default:
        break;
    }
  };

  const renderCell = (col: string, colIdx: number, row: number) => {
    const ref = `${col}${row}`;
    const entry = cells[ref];
    const raw = evaluateCell(ref, cells);
    const userFormat = formats[ref];
    const numberFormat = userFormat?.number ?? builtinNumberFormat(col, row);
    const display =
      showFormulas && entry?.formula !== undefined
        ? entry.formula
        : entry === undefined
          ? ''
          : formatCellValue(raw, numberFormat);

    const isTableHead = colIdx <= 5 && row === 1;
    const isTableBody = colIdx <= 5 && row >= 2 && row <= 25;
    const isLookupHead = (col === 'H' || col === 'I') && row === 1;
    const isLookupBody = (col === 'H' || col === 'I') && row >= 2 && row <= 10;
    const inSummary = colIdx >= 10 && colIdx <= 12 && row >= 1 && row <= 7;
    const inChecks = colIdx >= 10 && colIdx <= 13 && row >= 9 && row <= 12;

    const classes = ['cell'];
    let fillStyle: string | undefined;

    if (isTableHead) classes.push('table-head');
    if (isTableBody) {
      classes.push('table-cell');
      if (tableStyle && row % 2 === 0) classes.push('banded');
    }
    if (isLookupHead) classes.push('lookup-head');
    if (isLookupBody) classes.push('lookup-cell');
    if (inSummary || inChecks) classes.push('block-cell');
    if (colIdx >= 10 && colIdx <= 12 && row === 1) classes.push('block-title');
    if (colIdx >= 10 && colIdx <= 13 && row === 9) classes.push('block-title');
    if (colIdx >= 10 && colIdx <= 12 && row === 2) classes.push('block-head');
    if (colIdx >= 10 && colIdx <= 12 && row === 7) classes.push('block-total');
    if (inChecks && (col === 'K' || col === 'M') && row >= 10) classes.push('block-label');

    const isNumeric = typeof raw === 'number' && display !== '';
    if (isNumeric) classes.push('numeric');

    const conditionalHit =
      conditionalFormatting && col === 'E' && row >= 2 && row <= 25 && isNumeric && raw > 5000;
    if (conditionalHit) classes.push('cf-bad');
    if (userFormat?.bold) classes.push('fmt-bold');
    if (userFormat?.italic) classes.push('fmt-italic');
    if (userFormat?.align) classes.push(`align-${userFormat.align}`);
    if (userFormat?.fill && !conditionalHit) fillStyle = userFormat.fill;

    const isActive = ref === selectedCell;
    if (isActive) classes.push('active');

    const showBar =
      dataBars && col === 'E' && row >= 2 && row <= 25 && isNumeric && raw > 0 && maxGross > 0;

    return (
      <button
        key={ref}
        id={`cell-${ref}`}
        type="button"
        role="gridcell"
        aria-label={ref}
        aria-selected={isActive}
        className={classes.join(' ')}
        style={fillStyle ? { backgroundColor: fillStyle } : undefined}
        onClick={() => onSelectCell(ref)}
        tabIndex={isActive ? 0 : -1}
      >
        {showBar && (
          <span
            className="data-bar"
            style={{ width: `${Math.min(100, (raw as number) / maxGross * 100)}%` }}
          />
        )}
        <span className="cell-text">{display}</span>
        {isTableHead && <ChevronDown size={10} className="filter-chevron" />}
        {isActive && <span className="fill-handle" />}
      </button>
    );
  };

  const rows = [];
  for (let row = 1; row <= rowCount; row++) {
    rows.push(
      <div key={`rh-${row}`} className="row-header" role="rowheader">
        {row}
      </div>,
    );
    columnLetters.forEach((col, colIdx) => rows.push(renderCell(col, colIdx, row)));
  }

  return (
    <div
      className={`sheet-scroll${showGridlines ? '' : ' no-gridlines'}`}
      onKeyDown={handleKeyDown}
      role="grid"
      aria-label="Sales Data worksheet"
    >
      <div
        className="sheet-grid"
        style={{
          gridTemplateColumns: `44px ${colWidths.map((w) => `${w}px`).join(' ')}`,
          gridAutoRows: `${rowHeight}px`,
        }}
      >
        <div className="corner" />
        {columnLetters.map((col) => (
          <div key={`ch-${col}`} className="col-header" role="columnheader">
            {col}
          </div>
        ))}
        {rows}
        <div className="table-badge">Table1</div>
      </div>
    </div>
  );
}
