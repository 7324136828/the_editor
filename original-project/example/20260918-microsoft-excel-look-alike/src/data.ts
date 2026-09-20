import type { CellMap } from './types';

export const columnLetters = [
  'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N',
] as const;

export const rowCount = 40;
export const salesFirstRow = 2;
export const salesLastRow = 25;
export const salesTotalRow = 7;
export const regions = ['North', 'East', 'South', 'West'] as const;

const salesRows: [string, string, number, number][] = [
  ['P-1001', 'North', 42.5, 120],
  ['P-1002', 'East', 89.99, 72],
  ['P-1003', 'West', 34.75, 145],
  ['P-1004', 'South', 120, 38],
  ['P-1005', 'North', 68.5, 94],
  ['P-1006', 'East', 24.99, 210],
  ['P-1007', 'West', 155, 31],
  ['P-1008', 'South', 49.5, 112],
  ['P-1009', 'North', 210, 28],
  ['P-1001', 'East', 42.5, 132],
  ['P-1002', 'South', 89.99, 55],
  ['P-1003', 'North', 34.75, 168],
  ['P-1004', 'West', 120, 44],
  ['P-1005', 'East', 68.5, 87],
  ['P-1006', 'South', 24.99, 196],
  ['P-1007', 'North', 155, 36],
  ['P-1008', 'West', 49.5, 124],
  ['P-1009', 'East', 210, 24],
  ['P-1001', 'South', 42.5, 108],
  ['P-1002', 'North', 89.99, 63],
  ['P-1003', 'East', 34.75, 154],
  ['P-1004', 'South', 120, 47],
  ['P-1005', 'West', 68.5, 81],
  ['P-1006', 'North', 24.99, 228],
];

const categories = [
  'Office', 'Technology', 'Furniture', 'Technology', 'Office',
  'Supplies', 'Furniture', 'Supplies', 'Technology',
];

export function createInitialCells(): CellMap {
  const cells: CellMap = {};
  const headers = ['Product ID', 'Region', 'Unit Price', 'Units Sold', 'Gross Sales', 'Category'];
  headers.forEach((h, i) => {
    cells[`${columnLetters[i]}1`] = { value: h };
  });
  salesRows.forEach(([id, region, price, units], i) => {
    const row = i + salesFirstRow;
    cells[`A${row}`] = { value: id };
    cells[`B${row}`] = { value: region };
    cells[`C${row}`] = { value: price };
    cells[`D${row}`] = { value: units };
    cells[`E${row}`] = { formula: `=C${row}*D${row}` };
    cells[`F${row}`] = { formula: `=VLOOKUP(A${row},H$2:I$10,2,FALSE)` };
  });

  cells['H1'] = { value: 'Product ID' };
  cells['I1'] = { value: 'Category' };
  for (let i = 0; i < 9; i++) {
    cells[`H${i + 2}`] = { value: `P-100${i + 1}` };
    cells[`I${i + 2}`] = { value: categories[i] };
  }

  cells['K1'] = { value: 'REGIONAL SUMMARY' };
  cells['K2'] = { value: 'Region' };
  cells['L2'] = { value: 'Gross Sales' };
  cells['M2'] = { value: '% Total' };
  regions.forEach((region, i) => {
    const row = i + 3;
    cells[`K${row}`] = { value: region };
    cells[`L${row}`] = { formula: `=SUMIFS(E2:E25,B2:B25,"${region}")` };
    cells[`M${row}`] = { formula: `=L${row}/$L$7` };
  });
  cells['K7'] = { value: 'Grand Total' };
  cells['L7'] = { formula: '=SUM(E2:E25)' };
  cells['M7'] = { formula: '=SUM(M3:M6)' };

  cells['K9'] = { value: 'FORMULA CHECKS' };
  cells['K10'] = { value: 'East avg. units' };
  cells['L10'] = { formula: '=AVERAGEIF(B2:B25,"East",D2:D25)' };
  cells['K11'] = { value: 'Tax rate' };
  cells['L11'] = { value: 0.0725 };
  cells['K12'] = { value: 'Taxed sales' };
  cells['L12'] = { formula: '=L7*(1+$L$11)' };
  cells['M10'] = { value: 'Joined key' };
  cells['N10'] = { formula: '=CONCAT(LEFT(A2,3),"-",B2)' };
  cells['M11'] = { value: 'Absolute ref' };
  cells['N11'] = { formula: '=$L$11' };
  cells['M12'] = { value: 'Model state' };
  cells['N12'] = { value: 'Ready' };

  return cells;
}
