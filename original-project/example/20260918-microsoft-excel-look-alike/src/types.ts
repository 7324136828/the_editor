export type CellValue = string | number;
export type CellEntry = { value?: CellValue; formula?: string };
export type CellMap = Record<string, CellEntry>;
export type WorksheetId = 'sales' | 'pivot' | 'chart';
export type RibbonTab =
  | 'File'
  | 'Home'
  | 'Insert'
  | 'Page Layout'
  | 'Formulas'
  | 'Data'
  | 'Review'
  | 'View'
  | 'Help';
export type CellFormat = {
  bold?: boolean;
  italic?: boolean;
  fill?: string;
  align?: 'left' | 'center' | 'right';
  number?: 'general' | 'currency' | 'percent';
};
export type CellFormatMap = Record<string, CellFormat>;
export type SalesPoint = {
  row: number;
  productId: string;
  region: string;
  unitPrice: number;
  unitsSold: number;
  grossSales: number;
  category: string;
};
export type ToastFn = (message: string) => void;
