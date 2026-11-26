import type { ActiveSheetView, HostEnvironment } from '../types/spss';

interface BottomViewTabsProps {
  activeView: ActiveSheetView;
  host: HostEnvironment;
  caseCount: number;
  variableCount: number;
  showValueLabels: boolean;
  lastBackupTimestamp: string | null;
  onChangeView: (view: ActiveSheetView) => void;
}

export function BottomViewTabs({
  activeView,
  host,
  caseCount,
  variableCount,
  showValueLabels,
  lastBackupTimestamp,
  onChangeView,
}: BottomViewTabsProps) {
  const backupText = lastBackupTimestamp
    ? `Backup saved ${new Date(lastBackupTimestamp).toLocaleTimeString([], {
        hour: '2-digit',
        minute: '2-digit',
      })}`
    : 'No backup snapshot yet';

  return (
    <div className="bg-slate-100 border-t border-slate-200 px-4 flex flex-wrap items-center justify-between gap-4 shrink-0 select-none">
      {/* Classic SPSS bottom sheet tabs */}
      <div className="flex items-end gap-1 pt-1" role="tablist" aria-label="SPSS Views">
        <button
          type="button"
          role="tab"
          aria-selected={activeView === 'data'}
          onClick={() => onChangeView('data')}
          className={`px-4 py-2 text-xs font-semibold border-x border-t rounded-t-md transition-colors whitespace-nowrap shrink-0 cursor-pointer focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-blue-600 ${
            activeView === 'data'
              ? 'bg-white text-blue-700 border-slate-300 -mb-px pb-2.5'
              : 'bg-slate-200/70 text-slate-600 border-transparent hover:bg-slate-200 hover:text-slate-900'
          }`}
        >
          Data View
        </button>
        <button
          type="button"
          role="tab"
          aria-selected={activeView === 'variable'}
          onClick={() => onChangeView('variable')}
          className={`px-4 py-2 text-xs font-semibold border-x border-t rounded-t-md transition-colors whitespace-nowrap shrink-0 cursor-pointer focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-blue-600 ${
            activeView === 'variable'
              ? 'bg-white text-blue-700 border-slate-300 -mb-px pb-2.5'
              : 'bg-slate-200/70 text-slate-600 border-transparent hover:bg-slate-200 hover:text-slate-900'
          }`}
        >
          Variable View
        </button>
      </div>

      {/* Unboxed quiet metadata with typographic separators */}
      <div className="flex items-center gap-2 text-xs text-slate-500 py-2 tabular-nums overflow-x-auto whitespace-nowrap">
        <span>{host === 'excel' ? 'Excel Connected' : 'Simulator Workbook'}</span>
        <span aria-hidden="true">·</span>
        <span>{caseCount} cases</span>
        <span aria-hidden="true">·</span>
        <span>{variableCount} variables</span>
        <span aria-hidden="true">·</span>
        <span>Display: {showValueLabels ? 'Value Labels' : 'Raw Codes'}</span>
        <span aria-hidden="true" className="hidden md:inline">
          ·
        </span>
        <span className="hidden md:inline">{backupText}</span>
      </div>
    </div>
  );
}
