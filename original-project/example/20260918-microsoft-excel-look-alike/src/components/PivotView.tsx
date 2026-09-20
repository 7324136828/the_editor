import { useMemo, useState } from 'react';
import { GripVertical, Search } from 'lucide-react';
import { regions } from '../data';
import type { SalesPoint, ToastFn } from '../types';

export type PivotViewProps = {
  points: SalesPoint[];
  toast: ToastFn;
};

const currency = new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' });
const fields = ['Product ID', 'Region', 'Unit Price', 'Units Sold', 'Gross Sales', 'Category'];

export function PivotView({ points, toast }: PivotViewProps) {
  const [query, setQuery] = useState('');
  const [checked, setChecked] = useState<Record<string, boolean>>({
    Region: true,
    'Gross Sales': true,
  });

  const totals = useMemo(() => {
    const map = new Map<string, number>();
    for (const p of points) map.set(p.region, (map.get(p.region) ?? 0) + p.grossSales);
    return regions.map((region) => ({ region, total: map.get(region) ?? 0 }));
  }, [points]);

  const grand = totals.reduce((a, t) => a + t.total, 0);
  const max = Math.max(1, ...totals.map((t) => t.total));
  const visibleFields = fields.filter((f) => f.toLowerCase().includes(query.toLowerCase()));

  const toggleField = (f: string) => {
    setChecked((prev) => ({ ...prev, [f]: !prev[f] }));
    toast(`${f} ${checked[f] ? 'removed from' : 'added to'} PivotTable`);
  };

  return (
    <div className="pivot-sheet excel-bg">
      <div className="pivot-main">
        <div className="pivot-area">
          <div className="pivot-selected-frame">
            <table className="pivot-table" aria-label="Regional Sales Summary pivot table">
              <caption className="pivot-title">Regional Sales Summary</caption>
              <thead>
                <tr>
                  <th>Row Labels</th>
                  <th>Sum of Gross Sales</th>
                </tr>
              </thead>
              <tbody>
                {totals.map((t) => (
                  <tr key={t.region}>
                    <td>{t.region}</td>
                    <td className="pivot-num">{currency.format(t.total)}</td>
                  </tr>
                ))}
                <tr className="pivot-grand">
                  <td>Grand Total</td>
                  <td className="pivot-num">{currency.format(grand)}</td>
                </tr>
              </tbody>
            </table>
            <span className="fill-handle" />
          </div>
          <div className="pivot-bars" aria-label="Gross sales by region">
            <div className="pivot-bars-title">Gross sales by region</div>
            {totals.map((t) => (
              <div key={t.region} className="pivot-bar-row">
                <span className="pivot-bar-label">{t.region}</span>
                <span className="pivot-bar-track">
                  <span className="pivot-bar" style={{ width: `${(t.total / max) * 100}%` }} />
                </span>
                <span className="pivot-bar-value">{currency.format(t.total)}</span>
              </div>
            ))}
          </div>
        </div>
      </div>
      <aside className="pivot-pane" aria-label="PivotTable Fields pane">
        <div className="pane-title">PivotTable Fields</div>
        <div className="pane-search">
          <Search size={13} />
          <input
            type="search"
            placeholder="Search fields"
            aria-label="Search PivotTable fields"
            value={query}
            onChange={(e) => setQuery(e.target.value)}
          />
        </div>
        <div className="pane-field-list">
          {visibleFields.map((f) => (
            <label key={f} className="pane-check">
              <input
                type="checkbox"
                checked={!!checked[f]}
                onChange={() => toggleField(f)}
              />
              {f}
            </label>
          ))}
          {visibleFields.length === 0 && <div className="pane-empty">No fields match</div>}
        </div>
        <div className="pane-hint">Drag fields between areas below</div>
        <div className="pane-zones">
          <div className="zone">
            <div className="zone-label">Filters</div>
            <div className="zone-drop" />
          </div>
          <div className="zone">
            <div className="zone-label">Columns</div>
            <div className="zone-drop" />
          </div>
          <div className="zone">
            <div className="zone-label">Rows</div>
            <div className="zone-drop">
              <span className="zone-chip">
                <GripVertical size={11} /> Region
              </span>
            </div>
          </div>
          <div className="zone">
            <div className="zone-label">Values</div>
            <div className="zone-drop">
              <span className="zone-chip">
                <GripVertical size={11} /> Sum of Gross Sales
              </span>
            </div>
          </div>
        </div>
      </aside>
    </div>
  );
}
