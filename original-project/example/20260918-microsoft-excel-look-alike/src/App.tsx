import { useMemo, useRef, useState } from 'react';
import type { CSSProperties } from 'react';
import {
  Check,
  ChevronFirst,
  ChevronLast,
  ChevronLeft,
  ChevronRight,
  FileSpreadsheet,
  Minus,
  Plus,
  Redo2,
  Save,
  Search,
  Share2,
  Square,
  Undo2,
  X,
} from 'lucide-react';
import { createInitialCells } from './data';
import { evaluateCell, getSalesPoints } from './formulas';
import { Ribbon } from './components/Ribbon';
import { SpreadsheetGrid } from './components/SpreadsheetGrid';
import { ScatterChart } from './components/ScatterChart';
import { PivotView } from './components/PivotView';
import type {
  CellFormatMap,
  CellMap,
  RibbonTab,
  WorksheetId,
} from './types';

function rawText(cells: CellMap, ref: string): string {
  const entry = cells[ref];
  if (!entry) return '';
  if (entry.formula !== undefined) return entry.formula;
  return entry.value === undefined ? '' : String(entry.value);
}

export default function App() {
  const [cells, setCells] = useState<CellMap>(createInitialCells);
  const [activeSheet, setActiveSheet] = useState<WorksheetId>('sales');
  const [ribbonTab, setRibbonTab] = useState<RibbonTab>('Home');
  const [selectedCell, setSelectedCell] = useState('E2');
  const [formats, setFormats] = useState<CellFormatMap>({});
  const [showFormulas, setShowFormulas] = useState(false);
  const [showGridlines, setShowGridlines] = useState(true);
  const [tableStyle, setTableStyle] = useState(true);
  const [conditionalFormatting, setConditionalFormatting] = useState(true);
  const [dataBars, setDataBars] = useState(true);
  const [formulaDraft, setFormulaDraft] = useState(() => rawText(createInitialCells(), 'E2'));
  const [zoom, setZoom] = useState(100);
  const [toastMessage, setToastMessage] = useState<string | null>(null);
  const [autoSave, setAutoSave] = useState(false);
  const toastTimer = useRef<ReturnType<typeof setTimeout> | null>(null);

  const salesPoints = useMemo(() => getSalesPoints(cells), [cells]);

  const toast = (message: string) => {
    setToastMessage(message);
    if (toastTimer.current) clearTimeout(toastTimer.current);
    toastTimer.current = setTimeout(() => setToastMessage(null), 2400);
  };

  const selectCell = (ref: string) => {
    setSelectedCell(ref);
    setFormulaDraft(rawText(cells, ref));
  };

  const commitFormula = () => {
    const text = formulaDraft.trim();
    setCells((prev) => {
      const next = { ...prev };
      if (text === '') {
        delete next[selectedCell];
      } else if (text.startsWith('=')) {
        next[selectedCell] = { formula: text };
      } else if (Number.isFinite(Number(text))) {
        next[selectedCell] = { value: Number(text) };
      } else {
        next[selectedCell] = { value: text };
      }
      return next;
    });
    setFormulaDraft(text);
  };

  const cancelFormula = () => {
    setFormulaDraft(rawText(cells, selectedCell));
  };

  const mergeFormat = (patch: Partial<(typeof formats)[string]>) => {
    setFormats((prev) => ({
      ...prev,
      [selectedCell]: { ...prev[selectedCell], ...patch },
    }));
  };

  const selectedFormat = formats[selectedCell] ?? {};

  const toggleFill = () => {
    mergeFormat({ fill: selectedFormat.fill === '#fff2cc' ? undefined : '#fff2cc' });
  };

  const selectedValue = evaluateCell(selectedCell, cells);
  const numericSelected = typeof selectedValue === 'number';
  const editing = formulaDraft !== rawText(cells, selectedCell);

  const sheetTabs: { id: WorksheetId; label: string }[] = [
    { id: 'sales', label: 'Sales Data' },
    { id: 'pivot', label: 'Regional Pivot' },
    { id: 'chart', label: 'Price × Volume' },
  ];

  const moveSheet = (delta: number) => {
    const current = sheetTabs.findIndex((sheet) => sheet.id === activeSheet);
    const next = Math.min(sheetTabs.length - 1, Math.max(0, current + delta));
    if (next === current) {
      toast(delta < 0 ? 'This is the first sheet' : 'This is the last sheet');
      return;
    }
    setActiveSheet(sheetTabs[next].id);
  };

  return (
    <div className="app">
      <header className="title-bar">
        <div className="title-left">
          <div className="excel-tile" aria-hidden="true">
            X
          </div>
          <button
            type="button"
            className={`autosave${autoSave ? ' on' : ''}`}
            onClick={() => {
              setAutoSave(!autoSave);
              toast(autoSave ? 'AutoSave off' : 'AutoSave on');
            }}
            aria-pressed={autoSave}
            title="AutoSave"
          >
            <span className="autosave-knob" />
            AutoSave {autoSave ? 'On' : 'Off'}
          </button>
          <button type="button" className="title-btn" title="Save" onClick={() => toast('Saved to OneDrive')}>
            <Save size={14} />
          </button>
          <button type="button" className="title-btn" title="Undo" onClick={() => toast('Nothing to undo')}>
            <Undo2 size={14} />
          </button>
          <button type="button" className="title-btn" title="Redo" onClick={() => toast('Nothing to redo')}>
            <Redo2 size={14} />
          </button>
          <span className="doc-title">Sales Performance Model</span>
        </div>
        <div className="title-search">
          <Search size={13} />
          <input type="search" placeholder="Search in workbook" aria-label="Search in workbook" />
        </div>
        <div className="title-right">
          <button type="button" className="share-btn" onClick={() => toast('Share link copied')}>
            <Share2 size={13} /> Share
          </button>
          <div className="avatar" title="Zach N">
            ZN
          </div>
          <div className="window-controls">
            <button
              type="button"
              className="title-btn"
              title="Minimize"
              onClick={() => {
                if (window.desktopWindow) window.desktopWindow.minimize();
                else toast('Minimize is available in the desktop app');
              }}
            >
              <Minus size={14} />
            </button>
            <button
              type="button"
              className="title-btn"
              title="Maximize or restore"
              onClick={() => {
                if (window.desktopWindow) window.desktopWindow.toggleMaximize();
                else toast('Maximize is available in the desktop app');
              }}
            >
              <Square size={12} />
            </button>
            <button
              type="button"
              className="title-btn close"
              title="Close"
              onClick={() => {
                if (window.desktopWindow) window.desktopWindow.close();
                else toast('Close is available in the desktop app');
              }}
            >
              <X size={14} />
            </button>
          </div>
        </div>
      </header>

      <Ribbon
        activeTab={ribbonTab}
        onTabChange={setRibbonTab}
        bold={!!selectedFormat.bold}
        italic={!!selectedFormat.italic}
        filled={selectedFormat.fill === '#fff2cc'}
        currency={selectedFormat.number === 'currency'}
        showFormulas={showFormulas}
        showGridlines={showGridlines}
        tableStyle={tableStyle}
        conditionalFormatting={conditionalFormatting}
        dataBars={dataBars}
        onToggleBold={() => mergeFormat({ bold: !selectedFormat.bold })}
        onToggleItalic={() => mergeFormat({ italic: !selectedFormat.italic })}
        onToggleFill={toggleFill}
        onCurrency={() => mergeFormat({ number: 'currency' })}
        onToggleConditionalFormatting={() => setConditionalFormatting(!conditionalFormatting)}
        onToggleTableStyle={() => setTableStyle(!tableStyle)}
        onToggleDataBars={() => setDataBars(!dataBars)}
        onToggleShowFormulas={() => setShowFormulas(!showFormulas)}
        onToggleGridlines={() => setShowGridlines(!showGridlines)}
        onInsertTable={() => {
          setActiveSheet('sales');
          toast('Table1 selected on Sales Data');
        }}
        onShowPivot={() => setActiveSheet('pivot')}
        onShowChart={() => setActiveSheet('chart')}
        toast={toast}
      />

      <div className="formula-bar">
        <div className="name-box" title="Name Box">
          {selectedCell}
        </div>
        <div className="formula-actions">
          <button type="button" title="Cancel" aria-label="Cancel formula edit" onClick={cancelFormula}>
            <X size={13} />
          </button>
          <button type="button" title="Enter" aria-label="Commit formula" onClick={commitFormula}>
            <Check size={13} />
          </button>
          <span className="fx" title="Insert Function">
            fx
          </span>
        </div>
        <input
          className="formula-input"
          aria-label="Formula bar"
          value={formulaDraft}
          onChange={(e) => setFormulaDraft(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === 'Enter') {
              e.preventDefault();
              commitFormula();
            } else if (e.key === 'Escape') {
              e.preventDefault();
              cancelFormula();
            }
          }}
          spellCheck={false}
        />
      </div>

      <main className="work-area" style={{ '--sheet-zoom': zoom / 100 } as CSSProperties}>
        {activeSheet === 'sales' && (
          <SpreadsheetGrid
            cells={cells}
            formats={formats}
            selectedCell={selectedCell}
            onSelectCell={selectCell}
            showFormulas={showFormulas}
            showGridlines={showGridlines}
            tableStyle={tableStyle}
            conditionalFormatting={conditionalFormatting}
            dataBars={dataBars}
            toast={toast}
          />
        )}
        {activeSheet === 'chart' && <ScatterChart points={salesPoints} toast={toast} />}
        {activeSheet === 'pivot' && <PivotView points={salesPoints} toast={toast} />}
      </main>

      <div className="sheet-strip">
        <div className="sheet-nav">
          <button type="button" title="First sheet" onClick={() => setActiveSheet('sales')}>
            <ChevronFirst size={13} />
          </button>
          <button type="button" title="Previous sheet" onClick={() => moveSheet(-1)}>
            <ChevronLeft size={13} />
          </button>
          <button type="button" title="Next sheet" onClick={() => moveSheet(1)}>
            <ChevronRight size={13} />
          </button>
          <button type="button" title="Last sheet" onClick={() => setActiveSheet('chart')}>
            <ChevronLast size={13} />
          </button>
          <button type="button" title="New sheet" onClick={() => toast('New sheet added')}>
            <Plus size={13} />
          </button>
        </div>
        <div className="sheet-tabs" role="tablist" aria-label="Worksheets">
          {sheetTabs.map((tab) => (
            <button
              key={tab.id}
              type="button"
              role="tab"
              aria-selected={activeSheet === tab.id}
              className={`sheet-tab${activeSheet === tab.id ? ' active' : ''}`}
              onClick={() => setActiveSheet(tab.id)}
            >
              {tab.id === 'sales' && <FileSpreadsheet size={12} />}
              {tab.label}
            </button>
          ))}
        </div>
      </div>

      <footer className="status-bar">
        <div className="status-left">
          <span>{editing ? 'Edit' : 'Ready'}</span>
          <span className="status-accessibility">Accessibility: Good to go</span>
        </div>
        <div className="status-right">
          {numericSelected && (
            <>
              <span>
                Average:{' '}
                {new Intl.NumberFormat('en-US', { maximumFractionDigits: 2 }).format(selectedValue)}
              </span>
              <span>Count: 1</span>
              <span>
                Sum:{' '}
                {new Intl.NumberFormat('en-US', { maximumFractionDigits: 2 }).format(selectedValue)}
              </span>
            </>
          )}
          <div className="view-buttons">
            <button type="button" title="Normal" aria-pressed="true" onClick={() => toast('Normal view')}>
              <FileSpreadsheet size={13} />
            </button>
          </div>
          <input
            type="range"
            className="zoom-slider"
            min={50}
            max={200}
            value={zoom}
            aria-label="Zoom"
            onChange={(event) => setZoom(Number(event.target.value))}
          />
          <span className="zoom-label">{zoom}%</span>
        </div>
      </footer>

      {toastMessage && (
        <div className="toast" role="status" onClick={() => setToastMessage(null)}>
          {toastMessage}
        </div>
      )}
    </div>
  );
}
