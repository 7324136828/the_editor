import { useMemo, useState } from 'react';
import type { SalesPoint, ToastFn } from '../types';

export type ScatterChartProps = {
  points: SalesPoint[];
  toast: ToastFn;
};

const swatches = ['#4472c4', '#ed7d31', '#70ad47', '#a5a5a5', '#5b9bd5'];

function pearson(points: SalesPoint[]): number {
  const n = points.length;
  if (n < 2) return 0;
  const mx = points.reduce((a, p) => a + p.unitPrice, 0) / n;
  const my = points.reduce((a, p) => a + p.unitsSold, 0) / n;
  let num = 0;
  let dx = 0;
  let dy = 0;
  for (const p of points) {
    const vx = p.unitPrice - mx;
    const vy = p.unitsSold - my;
    num += vx * vy;
    dx += vx * vx;
    dy += vy * vy;
  }
  const den = Math.sqrt(dx * dy);
  return den === 0 ? 0 : num / den;
}

export function ScatterChart({ points, toast }: ScatterChartProps) {
  const [paneTab, setPaneTab] = useState<'Chart' | 'Format'>('Chart');
  const [showTitle, setShowTitle] = useState(true);
  const [showLegend, setShowLegend] = useState(true);
  const [showGrid, setShowGrid] = useState(true);
  const [color, setColor] = useState(swatches[0]);

  const stats = useMemo(() => {
    const r = pearson(points);
    const avgGross = points.length
      ? points.reduce((a, p) => a + p.grossSales, 0) / points.length
      : 0;
    return { r, avgGross };
  }, [points]);

  const width = 720;
  const height = 400;
  const pad = { left: 58, right: 24, top: 20, bottom: 52 };
  const xMax = Math.max(1, ...points.map((p) => p.unitPrice)) * 1.08;
  const yMax = Math.max(1, ...points.map((p) => p.unitsSold)) * 1.1;
  const x = (v: number) => pad.left + (v / xMax) * (width - pad.left - pad.right);
  const y = (v: number) => height - pad.bottom - (v / yMax) * (height - pad.top - pad.bottom);
  const xTicks = Array.from({ length: 6 }, (_, i) => (xMax / 5) * i);
  const yTicks = Array.from({ length: 6 }, (_, i) => (yMax / 5) * i);

  return (
    <div className="chart-sheet excel-bg">
      <div className="chart-main">
        <div className="chart-object" role="figure" aria-label="Price vs. Volume scatter chart">
          <span className="resize-handle tl" />
          <span className="resize-handle tc" />
          <span className="resize-handle tr" />
          <span className="resize-handle ml" />
          <span className="resize-handle mr" />
          <span className="resize-handle bl" />
          <span className="resize-handle bc" />
          <span className="resize-handle br" />
          {showTitle && (
            <div className="chart-heading">
              <div className="chart-title">Price vs. Volume</div>
              <div className="chart-subtitle">
                {points.length} products • linked to Sales Data
              </div>
            </div>
          )}
          <svg
            viewBox={`0 0 ${width} ${height}`}
            className="chart-svg"
            role="img"
            aria-label="Scatter plot of unit price against units sold"
          >
            {showGrid &&
              yTicks.map((t) => (
                <line
                  key={`gy-${t}`}
                  x1={pad.left}
                  x2={width - pad.right}
                  y1={y(t)}
                  y2={y(t)}
                  className="chart-gridline"
                />
              ))}
            {showGrid &&
              xTicks.map((t) => (
                <line
                  key={`gx-${t}`}
                  x1={x(t)}
                  x2={x(t)}
                  y1={pad.top}
                  y2={height - pad.bottom}
                  className="chart-gridline"
                />
              ))}
            <line
              x1={pad.left}
              x2={width - pad.right}
              y1={height - pad.bottom}
              y2={height - pad.bottom}
              className="chart-axis"
            />
            <line
              x1={pad.left}
              x2={pad.left}
              y1={pad.top}
              y2={height - pad.bottom}
              className="chart-axis"
            />
            {xTicks.map((t) => (
              <text key={`xt-${t}`} x={x(t)} y={height - pad.bottom + 16} className="tick" textAnchor="middle">
                {Math.round(t)}
              </text>
            ))}
            {yTicks.map((t) => (
              <text key={`yt-${t}`} x={pad.left - 8} y={y(t) + 4} className="tick" textAnchor="end">
                {Math.round(t)}
              </text>
            ))}
            <text x={(pad.left + width - pad.right) / 2} y={height - 10} className="axis-title" textAnchor="middle">
              Unit Price ($)
            </text>
            <text
              x={14}
              y={(pad.top + height - pad.bottom) / 2}
              className="axis-title"
              textAnchor="middle"
              transform={`rotate(-90 14 ${(pad.top + height - pad.bottom) / 2})`}
            >
              Units Sold
            </text>
            {points.map((p) => (
              <circle key={p.row} cx={x(p.unitPrice)} cy={y(p.unitsSold)} r={5.5} className="chart-dot" fill={color}>
                <title>{`${p.productId} • ${p.region} • $${p.unitPrice} • ${p.unitsSold} units`}</title>
              </circle>
            ))}
          </svg>
          {showLegend && (
            <div className="chart-legend">
              <span className="legend-swatch" style={{ backgroundColor: color }} />
              Sales observations
            </div>
          )}
          <div className="insight-strip">
            <span>
              Correlation (r): <strong>{stats.r.toFixed(2)}</strong>
            </span>
            <span>
              Observations: <strong>{points.length}</strong>
            </span>
            <span>
              Avg gross sale:{' '}
              <strong>
                {new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' }).format(
                  stats.avgGross,
                )}
              </strong>
            </span>
          </div>
        </div>
      </div>
      <aside className="chart-pane" aria-label="Chart pane">
        <div className="pane-tabs" role="tablist">
          {(['Chart', 'Format'] as const).map((t) => (
            <button
              key={t}
              type="button"
              role="tab"
              aria-selected={paneTab === t}
              className={`pane-tab${paneTab === t ? ' active' : ''}`}
              onClick={() => setPaneTab(t)}
            >
              {t}
            </button>
          ))}
        </div>
        {paneTab === 'Chart' ? (
          <div className="pane-body">
            <div className="pane-section-title">Chart Options</div>
            <label className="pane-check">
              <input
                type="checkbox"
                checked={showTitle}
                onChange={(e) => setShowTitle(e.target.checked)}
              />
              Chart Title
            </label>
            <label className="pane-check">
              <input
                type="checkbox"
                checked={showLegend}
                onChange={(e) => setShowLegend(e.target.checked)}
              />
              Legend
            </label>
            <label className="pane-check">
              <input
                type="checkbox"
                checked={showGrid}
                onChange={(e) => setShowGrid(e.target.checked)}
              />
              Gridlines
            </label>
            <div className="pane-section-title">Chart Styles</div>
            <div className="swatch-row">
              {swatches.map((s) => (
                <button
                  key={s}
                  type="button"
                  className={`swatch${color === s ? ' active' : ''}`}
                  style={{ backgroundColor: s }}
                  aria-label={`Series color ${s}`}
                  aria-pressed={color === s}
                  onClick={() => setColor(s)}
                />
              ))}
            </div>
          </div>
        ) : (
          <div className="pane-body">
            <div className="pane-section-title">Shape Styles</div>
            <label className="pane-check">
              <input type="checkbox" defaultChecked onChange={() => toast('Glow effect applied')} />
              Glow
            </label>
            <label className="pane-check">
              <input type="checkbox" onChange={() => toast('Shadow effect applied')} />
              Shadow
            </label>
            <label className="pane-check">
              <input type="checkbox" defaultChecked onChange={() => toast('Border updated')} />
              Border
            </label>
            <div className="pane-section-title">Size</div>
            <button type="button" className="pane-btn" onClick={() => toast('Chart resized to fit')}>
              Fit to data
            </button>
          </div>
        )}
      </aside>
    </div>
  );
}
