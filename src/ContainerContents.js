import { Component, lazy, Suspense, useId, useRef } from 'react';
import { Box, List, Loader } from 'lucide-react';

const ContainerVisualization = lazy(() => import('./ContainerVisualization'));

class ViewErrorBoundary extends Component {
  state = { failed: false };
  static getDerivedStateFromError() { return { failed: true }; }
  render() {
    if (this.state.failed) return (
      <p role="alert" className="rounded-lg bg-slate-100 p-6 text-slate-700 dark:bg-slate-800 dark:text-slate-200">
        The 3D view could not load. Switch to List view to continue; your products are still available.
      </p>
    );
    return this.props.children;
  }
}

export default function ContainerContents({ container, maxCbm, active, onViewChange, children }) {
  const id = useId();
  const tabs = useRef([]);
  const selected = active ? 1 : 0;
  const select = index => {
    onViewChange(index === 1);
    tabs.current[index]?.focus();
  };
  return (
    <>
      <div role="tablist" aria-label={`Container ${container.id} view`} className="mb-3 inline-flex gap-1 rounded-lg bg-slate-100 p-1 dark:bg-slate-800">
        {['List view', '3D view'].map((label, index) => {
          const Icon = index === 0 ? List : Box;
          return (
            <button key={label} type="button" role="tab" id={`${id}-tab-${index}`} aria-controls={`${id}-panel-${index}`}
              aria-selected={selected === index} tabIndex={selected === index ? 0 : -1} ref={node => { tabs.current[index] = node; }}
              onClick={() => select(index)} onKeyDown={event => {
                if (!['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(event.key)) return;
                event.preventDefault();
                select(event.key === 'Home' ? 0 : event.key === 'End' ? 1 : 1 - selected);
              }}
              className={`flex items-center gap-2 rounded-md px-3 py-2 text-sm font-medium transition-colors ${selected === index
                ? 'bg-white text-slate-900 shadow-sm dark:bg-slate-700 dark:text-slate-100'
                : 'text-slate-600 hover:bg-slate-200 dark:text-slate-300 dark:hover:bg-slate-700'}`}>
              <Icon className="h-4 w-4" aria-hidden="true" />{label}
            </button>
          );
        })}
      </div>
      <div role="tabpanel" id={`${id}-panel-0`} aria-labelledby={`${id}-tab-0`} hidden={active}>
        {children}
      </div>
      <div role="tabpanel" id={`${id}-panel-1`} aria-labelledby={`${id}-tab-1`} hidden={!active}>
        {active && (
          <ViewErrorBoundary>
            <Suspense fallback={<p role="status" className="flex items-center gap-2 p-6 text-slate-600 dark:text-slate-300"><Loader className="h-4 w-4 animate-spin" />Loading 3D view…</p>}>
              <ContainerVisualization container={container} maxCbm={maxCbm} />
            </Suspense>
          </ViewErrorBoundary>
        )}
      </div>
    </>
  );
}
