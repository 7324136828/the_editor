import type { ReactNode } from 'react';
import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  ArrowDownAZ,
  Baseline,
  Bold,
  Braces,
  ChartColumn,
  ChartLine,
  ChartPie,
  ChartScatter,
  Check,
  CircleHelp,
  ClipboardPaste,
  Code,
  Copy,
  DollarSign,
  Eraser,
  FileSpreadsheet,
  Filter,
  History,
  Italic,
  LayoutGrid,
  ListFilter,
  Lock,
  MessageCircleQuestion,
  PaintBucket,
  Percent,
  Plus,
  Printer,
  RefreshCw,
  Save,
  Scissors,
  Settings2,
  ShieldCheck,
  Sigma,
  Snowflake,
  Sparkles,
  SquareFunction,
  Table,
  Underline,
  ZoomIn,
} from 'lucide-react';
import type { RibbonTab, ToastFn } from '../types';

type RibbonButtonProps = {
  icon: ReactNode;
  label: string;
  onClick: () => void;
  pressed?: boolean;
  title?: string;
};

function RibbonButton({ icon, label, onClick, pressed, title }: RibbonButtonProps) {
  return (
    <button
      type="button"
      className="ribbon-btn"
      onClick={onClick}
      aria-pressed={pressed}
      title={title ?? label}
    >
      <span className="ribbon-btn-icon">{icon}</span>
      <span className="ribbon-btn-label">{label}</span>
    </button>
  );
}

function RibbonGroup({ label, children }: { label: string; children: ReactNode }) {
  return (
    <div className="ribbon-group">
      <div className="ribbon-group-items">{children}</div>
      <div className="ribbon-group-label">{label}</div>
    </div>
  );
}

export type RibbonProps = {
  activeTab: RibbonTab;
  onTabChange: (tab: RibbonTab) => void;
  bold: boolean;
  italic: boolean;
  filled: boolean;
  currency: boolean;
  showFormulas: boolean;
  showGridlines: boolean;
  tableStyle: boolean;
  conditionalFormatting: boolean;
  dataBars: boolean;
  onToggleBold: () => void;
  onToggleItalic: () => void;
  onToggleFill: () => void;
  onCurrency: () => void;
  onToggleConditionalFormatting: () => void;
  onToggleTableStyle: () => void;
  onToggleDataBars: () => void;
  onToggleShowFormulas: () => void;
  onToggleGridlines: () => void;
  onInsertTable: () => void;
  onShowPivot: () => void;
  onShowChart: () => void;
  toast: ToastFn;
};

const tabs: RibbonTab[] = [
  'File',
  'Home',
  'Insert',
  'Page Layout',
  'Formulas',
  'Data',
  'Review',
  'View',
  'Help',
];

const iconSize = 16;

export function Ribbon(props: RibbonProps) {
  const { activeTab, onTabChange, toast } = props;
  const size = iconSize;

  const content = () => {
    switch (activeTab) {
      case 'Home':
        return (
          <>
            <RibbonGroup label="Clipboard">
              <RibbonButton
                icon={<ClipboardPaste size={size} />}
                label="Paste"
                onClick={() => toast('Clipboard is empty')}
              />
              <RibbonButton
                icon={<Scissors size={size} />}
                label="Cut"
                onClick={() => toast('Cut is not available')}
              />
              <RibbonButton
                icon={<Copy size={size} />}
                label="Copy"
                onClick={() => toast('Copied selection')}
              />
            </RibbonGroup>
            <RibbonGroup label="Font">
              <RibbonButton
                icon={<Bold size={size} />}
                label="Bold"
                pressed={props.bold}
                onClick={props.onToggleBold}
                title="Bold (Ctrl+B)"
              />
              <RibbonButton
                icon={<Italic size={size} />}
                label="Italic"
                pressed={props.italic}
                onClick={props.onToggleItalic}
                title="Italic (Ctrl+I)"
              />
              <RibbonButton
                icon={<Underline size={size} />}
                label="Underline"
                onClick={() => toast('Underline applied')}
              />
              <RibbonButton
                icon={<PaintBucket size={size} />}
                label="Fill Color"
                pressed={props.filled}
                onClick={props.onToggleFill}
                title="Fill Color"
              />
              <RibbonButton
                icon={<Baseline size={size} />}
                label="Font Color"
                onClick={() => toast('Font color applied')}
              />
            </RibbonGroup>
            <RibbonGroup label="Alignment">
              <RibbonButton
                icon={<AlignLeft size={size} />}
                label="Left"
                onClick={() => toast('Aligned left')}
              />
              <RibbonButton
                icon={<AlignCenter size={size} />}
                label="Center"
                onClick={() => toast('Aligned center')}
              />
              <RibbonButton
                icon={<AlignRight size={size} />}
                label="Right"
                onClick={() => toast('Aligned right')}
              />
            </RibbonGroup>
            <RibbonGroup label="Number">
              <RibbonButton
                icon={<DollarSign size={size} />}
                label="Currency"
                pressed={props.currency}
                onClick={props.onCurrency}
                title="Currency format"
              />
              <RibbonButton
                icon={<Percent size={size} />}
                label="Percent"
                onClick={() => toast('Percent style applied')}
              />
            </RibbonGroup>
            <RibbonGroup label="Styles">
              <RibbonButton
                icon={<Sparkles size={size} />}
                label="Conditional Formatting"
                pressed={props.conditionalFormatting}
                onClick={props.onToggleConditionalFormatting}
              />
              <RibbonButton
                icon={<Table size={size} />}
                label="Format as Table"
                pressed={props.tableStyle}
                onClick={props.onToggleTableStyle}
              />
              <RibbonButton
                icon={<ChartColumn size={size} />}
                label="Data Bars"
                pressed={props.dataBars}
                onClick={props.onToggleDataBars}
              />
            </RibbonGroup>
            <RibbonGroup label="Cells">
              <RibbonButton
                icon={<Plus size={size} />}
                label="Insert"
                onClick={() => toast('Insert cells is unavailable in this view')}
              />
              <RibbonButton
                icon={<Eraser size={size} />}
                label="Delete"
                onClick={() => toast('Delete cells is unavailable in this view')}
              />
              <RibbonButton
                icon={<Settings2 size={size} />}
                label="Format"
                onClick={() => toast('Format options opened')}
              />
            </RibbonGroup>
            <RibbonGroup label="Editing">
              <RibbonButton
                icon={<Sigma size={size} />}
                label="AutoSum"
                onClick={() => toast('AutoSum inserted =SUM(E2:E25)')}
              />
              <RibbonButton
                icon={<ListFilter size={size} />}
                label="Sort & Filter"
                onClick={() => toast('Sort & Filter menu opened')}
              />
              <RibbonButton
                icon={<Braces size={size} />}
                label="Find & Select"
                onClick={() => toast('Find & Select opened')}
              />
            </RibbonGroup>
          </>
        );
      case 'Insert':
        return (
          <>
            <RibbonGroup label="Tables">
              <RibbonButton
                icon={<Table size={size} />}
                label="Table"
                onClick={props.onInsertTable}
              />
              <RibbonButton
                icon={<LayoutGrid size={size} />}
                label="PivotTable"
                onClick={props.onShowPivot}
              />
            </RibbonGroup>
            <RibbonGroup label="Illustrations">
              <RibbonButton
                icon={<FileSpreadsheet size={size} />}
                label="Pictures"
                onClick={() => toast('Insert Pictures opened')}
              />
            </RibbonGroup>
            <RibbonGroup label="Charts">
              <RibbonButton
                icon={<ChartScatter size={size} />}
                label="Scatter (X, Y)"
                onClick={props.onShowChart}
              />
              <RibbonButton
                icon={<ChartLine size={size} />}
                label="Recommended Charts"
                onClick={() => toast('Recommended Charts opened')}
              />
              <RibbonButton
                icon={<ChartPie size={size} />}
                label="Pie"
                onClick={() => toast('Pie chart inserted')}
              />
            </RibbonGroup>
            <RibbonGroup label="Links">
              <RibbonButton
                icon={<Braces size={size} />}
                label="Link"
                onClick={() => toast('Insert Link opened')}
              />
            </RibbonGroup>
          </>
        );
      case 'Formulas':
        return (
          <>
            <RibbonGroup label="Function Library">
              <RibbonButton
                icon={<SquareFunction size={size} />}
                label="Insert Function"
                onClick={() => toast('Insert Function dialog opened')}
              />
              <RibbonButton
                icon={<Sigma size={size} />}
                label="AutoSum"
                onClick={() => toast('AutoSum inserted =SUM(E2:E25)')}
              />
              <RibbonButton
                icon={<History size={size} />}
                label="Recently Used"
                onClick={() => toast('Recently used functions shown')}
              />
            </RibbonGroup>
            <RibbonGroup label="Defined Names">
              <RibbonButton
                icon={<Table size={size} />}
                label="Name Manager"
                onClick={() => toast('Name Manager opened (Table1)')}
              />
            </RibbonGroup>
            <RibbonGroup label="Formula Auditing">
              <RibbonButton
                icon={<Code size={size} />}
                label="Show Formulas"
                pressed={props.showFormulas}
                onClick={props.onToggleShowFormulas}
              />
              <RibbonButton
                icon={<Braces size={size} />}
                label="Error Checking"
                onClick={() => toast('No formula errors found')}
              />
            </RibbonGroup>
            <RibbonGroup label="Calculation">
              <RibbonButton
                icon={<RefreshCw size={size} />}
                label="Calculate Now"
                onClick={() => toast('Workbook recalculated')}
              />
            </RibbonGroup>
          </>
        );
      case 'Data':
        return (
          <>
            <RibbonGroup label="Get & Transform">
              <RibbonButton
                icon={<RefreshCw size={size} />}
                label="Refresh All"
                onClick={() => toast('All connections refreshed')}
              />
            </RibbonGroup>
            <RibbonGroup label="Sort & Filter">
              <RibbonButton
                icon={<ArrowDownAZ size={size} />}
                label="Sort"
                onClick={() => toast('Sorted by Gross Sales (largest to smallest)')}
              />
              <RibbonButton
                icon={<Filter size={size} />}
                label="Filter"
                onClick={() => toast('Filter toggled for Table1')}
              />
            </RibbonGroup>
            <RibbonGroup label="Data Tools">
              <RibbonButton
                icon={<ShieldCheck size={size} />}
                label="Data Validation"
                onClick={() => toast('Data Validation settings opened')}
              />
            </RibbonGroup>
            <RibbonGroup label="Outline">
              <RibbonButton
                icon={<LayoutGrid size={size} />}
                label="Group"
                onClick={() => toast('Rows grouped')}
              />
            </RibbonGroup>
          </>
        );
      case 'View':
        return (
          <>
            <RibbonGroup label="Workbook Views">
              <RibbonButton
                icon={<FileSpreadsheet size={size} />}
                label="Normal"
                pressed
                onClick={() => toast('Normal view active')}
              />
            </RibbonGroup>
            <RibbonGroup label="Show">
              <RibbonButton
                icon={<LayoutGrid size={size} />}
                label="Gridlines"
                pressed={props.showGridlines}
                onClick={props.onToggleGridlines}
              />
              <RibbonButton
                icon={<Code size={size} />}
                label="Formula Bar"
                pressed
                onClick={() => toast('Formula bar visibility toggled')}
              />
            </RibbonGroup>
            <RibbonGroup label="Window">
              <RibbonButton
                icon={<Snowflake size={size} />}
                label="Freeze Panes"
                onClick={() => toast('Panes frozen at B2')}
              />
            </RibbonGroup>
            <RibbonGroup label="Zoom">
              <RibbonButton
                icon={<ZoomIn size={size} />}
                label="Zoom"
                onClick={() => toast('Zoom set to 100%')}
              />
            </RibbonGroup>
          </>
        );
      case 'Page Layout':
        return (
          <>
            <RibbonGroup label="Page Setup">
              <RibbonButton
                icon={<FileSpreadsheet size={size} />}
                label="Margins"
                onClick={() => toast('Margins set to Normal')}
              />
              <RibbonButton
                icon={<LayoutGrid size={size} />}
                label="Orientation"
                onClick={() => toast('Orientation set to Landscape')}
              />
              <RibbonButton
                icon={<Printer size={size} />}
                label="Print Area"
                onClick={() => toast('Print area set to A1:N40')}
              />
            </RibbonGroup>
            <RibbonGroup label="Sheet Options">
              <RibbonButton
                icon={<LayoutGrid size={size} />}
                label="Gridlines"
                pressed={props.showGridlines}
                onClick={props.onToggleGridlines}
              />
            </RibbonGroup>
          </>
        );
      case 'Review':
        return (
          <>
            <RibbonGroup label="Proofing">
              <RibbonButton
                icon={<Check size={size} />}
                label="Spelling"
                onClick={() => toast('Spelling check complete')}
              />
            </RibbonGroup>
            <RibbonGroup label="Comments">
              <RibbonButton
                icon={<MessageCircleQuestion size={size} />}
                label="New Comment"
                onClick={() => toast('New comment added')}
              />
            </RibbonGroup>
            <RibbonGroup label="Protect">
              <RibbonButton
                icon={<Lock size={size} />}
                label="Protect Sheet"
                onClick={() => toast('Sheet protection toggled')}
              />
            </RibbonGroup>
          </>
        );
      case 'Help':
        return (
          <>
            <RibbonGroup label="Help">
              <RibbonButton
                icon={<CircleHelp size={size} />}
                label="Help"
                onClick={() => toast('Help pane opened')}
              />
              <RibbonButton
                icon={<MessageCircleQuestion size={size} />}
                label="Contact Support"
                onClick={() => toast('Support request drafted')}
              />
            </RibbonGroup>
            <RibbonGroup label="About">
              <RibbonButton
                icon={<FileSpreadsheet size={size} />}
                label="About"
                onClick={() => toast('Excel Look-Alike 1.0')}
              />
            </RibbonGroup>
          </>
        );
      case 'File':
        return (
          <>
            <RibbonGroup label="Workbook">
              <RibbonButton
                icon={<Save size={size} />}
                label="Save"
                onClick={() => toast('Workbook saved')}
              />
              <RibbonButton
                icon={<FileSpreadsheet size={size} />}
                label="Open"
                onClick={() => toast('Open dialog opened')}
              />
              <RibbonButton
                icon={<Printer size={size} />}
                label="Print"
                onClick={() => toast('Sent to printer')}
              />
            </RibbonGroup>
            <RibbonGroup label="Share">
              <RibbonButton
                icon={<Braces size={size} />}
                label="Export"
                onClick={() => toast('Exported as XLSX')}
              />
            </RibbonGroup>
          </>
        );
      default:
        return null;
    }
  };

  return (
    <div className="ribbon">
      <div className="ribbon-tabs" role="tablist" aria-label="Ribbon tabs">
        {tabs.map((tab) => (
          <button
            key={tab}
            type="button"
            role="tab"
            aria-selected={activeTab === tab}
            className={`ribbon-tab${tab === 'File' ? ' ribbon-tab-file' : ''}${
              activeTab === tab && tab !== 'File' ? ' active' : ''
            }${activeTab === tab && tab === 'File' ? ' file-active' : ''}`}
            onClick={() => onTabChange(tab)}
          >
            {tab}
          </button>
        ))}
      </div>
      <div className="ribbon-content" role="tabpanel">
        {content()}
      </div>
    </div>
  );
}
