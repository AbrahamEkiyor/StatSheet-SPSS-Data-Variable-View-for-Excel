import { Download, RefreshCw } from 'lucide-react';
import type { HostEnvironment } from '../types/spss';

interface TopBarProps {
  host: HostEnvironment;
  isSyncing: boolean;
  onFocusGrid: () => void;
  onOpenCodebook: () => void;
  onOpenStats: () => void;
  onOpenExcelAddin: () => void;
  onPrimaryAction: () => void;
}

export function TopBar({
  host,
  isSyncing,
  onFocusGrid,
  onOpenCodebook,
  onOpenStats,
  onOpenExcelAddin,
  onPrimaryAction,
}: TopBarProps) {
  return (
    <header className="flex items-center justify-between gap-8 px-5 h-13 bg-white border-b border-slate-200 shrink-0">
      {/* Zone 1: Brand wordmark only */}
      <a
        href="#top"
        onClick={(e) => {
          e.preventDefault();
          onFocusGrid();
        }}
        className="font-display text-lg font-bold tracking-tight text-slate-900 whitespace-nowrap shrink-0 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-blue-600 rounded"
      >
        StatSheet
      </a>

      {/* Zone 2: Clean single-line navigation links */}
      <nav
        aria-label="Workspace tools"
        className="hidden sm:flex items-center gap-6 text-sm font-medium text-slate-600"
      >
        <button
          type="button"
          onClick={onFocusGrid}
          className="py-1 text-slate-900 border-b-2 border-blue-700 whitespace-nowrap shrink-0 cursor-pointer focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-blue-600"
        >
          Spreadsheet
        </button>
        <button
          type="button"
          onClick={onOpenCodebook}
          className="py-1 border-b-2 border-transparent hover:text-slate-900 hover:border-slate-300 transition-colors whitespace-nowrap shrink-0 cursor-pointer focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-blue-600"
        >
          Codebook
        </button>
        <button
          type="button"
          onClick={onOpenStats}
          className="py-1 border-b-2 border-transparent hover:text-slate-900 hover:border-slate-300 transition-colors whitespace-nowrap shrink-0 cursor-pointer focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-blue-600"
        >
          Statistics
        </button>
        <button
          type="button"
          onClick={onOpenExcelAddin}
          className="py-1 border-b-2 border-transparent hover:text-slate-900 hover:border-slate-300 transition-colors whitespace-nowrap shrink-0 cursor-pointer focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-blue-600"
        >
          Excel Add-in
        </button>
      </nav>

      {/* Zone 3: Single primary CTA */}
      <div className="flex items-center shrink-0">
        <button
          type="button"
          onClick={onPrimaryAction}
          disabled={isSyncing}
          className="inline-flex items-center gap-2 px-4 py-2 text-xs font-semibold text-white bg-[#107C41] hover:bg-[#0D6635] disabled:opacity-60 rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-offset-1 focus-visible:ring-[#107C41]"
        >
          {host === 'excel' ? (
            <>
              <RefreshCw className={`w-3.5 h-3.5 ${isSyncing ? 'animate-spin' : ''}`} />
              <span>{isSyncing ? 'Syncing...' : 'Sync to Excel'}</span>
            </>
          ) : (
            <>
              <Download className="w-3.5 h-3.5" />
              <span>Export .xlsx</span>
            </>
          )}
        </button>
      </div>
    </header>
  );
}
