import { Component, useMemo, useState, useSyncExternalStore } from 'react';
import { RotateCcw } from 'lucide-react';
import ContainerScene from './ContainerScene';
import { buildContainerLayout } from './services/containerLayout';
import { formatDecimal } from './utils/decimal';

function subscribeTheme(notify) {
  const observer = new MutationObserver(notify);
  observer.observe(document.documentElement, { attributes: true, attributeFilter: ['class'] });
  return () => observer.disconnect();
}
const isDark = () => document.documentElement.classList.contains('dark');

class SceneErrorBoundary extends Component {
  state = { failed: false };
  static getDerivedStateFromError() { return { failed: true }; }
  render() {
    return this.state.failed
      ? <p role="alert" className="flex h-full items-center justify-center p-6 text-center">3D graphics are unavailable in this browser. You can continue in List view.</p>
      : this.props.children;
  }
}

export default function ContainerVisualization({ container, maxCbm }) {
  const dark = useSyncExternalStore(subscribeTheme, isDark, () => false);
  const layout = useMemo(() => buildContainerLayout(container, maxCbm), [container, maxCbm]);
  const [resetKey, setResetKey] = useState(0);
  const reasons = new Map();
  layout.unplaced.forEach(item => reasons.set(item.reason, (reasons.get(item.reason) || 0) + 1));
  return (
    <section aria-label={`Container ${container.id} estimated 3D layout`} className="overflow-hidden rounded-xl border border-slate-200 bg-slate-50 text-slate-900 dark:border-slate-700 dark:bg-slate-950 dark:text-slate-100">
      <div className="flex flex-wrap items-center justify-between gap-3 border-b border-slate-200 px-4 py-3 dark:border-slate-700">
        <div>
          <h4 className="font-semibold">Estimated carton layout</h4>
          <p className="text-xs text-slate-600 dark:text-slate-400">Estimated layout — dimensions unavailable. Shapes inferred from CBM; physical fit is not verified.</p>
        </div>
        <button type="button" onClick={() => setResetKey(key => key + 1)} className="flex items-center gap-2 rounded-lg px-3 py-2 text-sm text-slate-700 transition-colors hover:bg-slate-200 dark:text-slate-200 dark:hover:bg-slate-800">
          <RotateCcw className="h-4 w-4" aria-hidden="true" />Reset camera
        </button>
      </div>
      <div className="grid lg:grid-cols-[minmax(0,1fr)_240px]">
        <div>
          <div className="relative h-72 sm:h-96" role="img" aria-label={`Estimated arrangement of ${layout.placements.length} cartons in container ${container.id}`}>
            {layout.capacityCbm === null ? <p role="status" className="p-6">Set a positive container CBM capacity to display the layout.</p> : (
              <SceneErrorBoundary key={resetKey}>
                <ContainerScene layout={layout} dark={dark} />
              </SceneErrorBoundary>
            )}
            {container.items.length === 0 && <p className="pointer-events-none absolute inset-x-0 bottom-4 text-center text-sm text-slate-500 dark:text-slate-400">This container is empty.</p>}
          </div>
          <p className="px-4 pb-4 text-xs text-slate-600 dark:text-slate-400">Drag to rotate · Scroll or pinch to zoom · Right-drag to pan</p>
        </div>
        <aside className="border-t border-slate-200 p-4 dark:border-slate-700 lg:border-l lg:border-t-0">
          <p className="text-xs font-medium uppercase tracking-wide text-slate-600 dark:text-slate-400">Volume used</p>
          <p className="mt-1 text-3xl font-bold tracking-tight">{layout.fillPercent === null ? '—' : `${layout.fillPercent.toFixed(1)}%`}</p>
          <p className="mt-1 text-sm text-slate-600 dark:text-slate-300">{layout.usedCbm === null ? '—' : formatDecimal(layout.usedCbm)} / {layout.capacityCbm === null ? '—' : formatDecimal(layout.capacityCbm)} m³</p>
          {layout.fillPercent > 100 && <p className="mt-2 text-sm font-medium text-red-700 dark:text-red-300">Assigned volume exceeds capacity.</p>}
          <p className="mt-4 text-sm">{layout.placements.length} / {container.items.length} cartons shown</p>
          {layout.unplaced.length > 0 && (
            <div role="status" className="mt-3 rounded-lg bg-amber-50 p-3 text-sm text-amber-900 dark:bg-amber-950/50 dark:text-amber-200">
              <p className="font-semibold">{layout.unplaced.length} cartons unplaced</p>
              <ul className="mt-1 space-y-1">{[...reasons].map(([reason, count]) => <li key={reason}>{count} · {reason}</li>)}</ul>
              <p className="mt-2 text-xs">Totals include all assigned cartons. List view shows every product.</p>
            </div>
          )}
          <h5 className="mb-2 mt-5 text-sm font-semibold">Products</h5>
          <ul aria-label="Product color legend" className="max-h-48 space-y-3 overflow-y-auto pr-1 text-sm">
            {layout.products.map(product => (
              <li key={product.productKey} className="flex items-start gap-2">
                <span aria-hidden="true" className="mt-1 h-3 w-3 shrink-0 rounded-sm" style={{ backgroundColor: product.color }} />
                <div className="min-w-0"><p className="break-words font-medium">{product.markNo || product.description}</p>
                  {product.markNo && <p className="break-words text-xs text-slate-600 dark:text-slate-400">{product.description}</p>}
                  <p className="text-xs text-slate-600 dark:text-slate-400">{product.placedCartons} / {product.totalCartons} cartons shown</p>
                </div>
              </li>
            ))}
          </ul>
        </aside>
      </div>
    </section>
  );
}
