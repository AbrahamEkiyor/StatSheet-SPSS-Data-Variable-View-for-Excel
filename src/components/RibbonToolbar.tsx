import {
  AlertTriangle,
  BarChart3,
  BookOpen,
  Columns3,
  FileCode2,
  FileSpreadsheet,
  FolderOpen,
  PanelRight,
  Plus,
  Redo2,
  RotateCcw,
  Search,
  Tags,
  Undo2,
  Upload,
} from 'lucide-react';
import { useRef } from 'react';
import { SAMPLE_PRESETS } from '../data/sampleDatasets';
import type { ActiveSheetView, HostEnvironment } from '../types/spss';

interface RibbonToolbarProps {
  host: HostEnvironment;
  datasetName: string;
  activeView: ActiveSheetView;
  showValueLabels: boolean;
  canUndo: boolean;
  canRedo: boolean;
  issueCount: number;
  issuesOnly: boolean;
  searchQuery: string;
  taskPanePreview: boolean;
  onToggleValueLabels: () => void;
  onUndo: () => void;
  onRedo: () => void;
  onAddRow: () => void;
  onAddVariable: () => void;
  onToggleIssuesOnly: () => void;
  onSearchChange: (q: string) => void;
  onLoadPreset: (presetId: string) => void;
  onImportWorkbook: (file: File) => void;
  onToggleTaskPanePreview: () => void;
  onOpenStats: () => void;
  onOpenCodebook: () => void;
  onOpenExcelAddin: () => void;
  onRefreshFromExcel: () => void;
  onRestoreRawBackup: () => void;
}

export function RibbonToolbar({
  host,
  datasetName,
  activeView,
  showValueLabels,
  canUndo,
  canRedo,
  issueCount,
  issuesOnly,
  searchQuery,
  taskPanePreview,
  onToggleValueLabels,
  onUndo,
  onRedo,
  onAddRow,
  onAddVariable,
  onToggleIssuesOnly,
  onSearchChange,
  onLoadPreset,
  onImportWorkbook,
  onToggleTaskPanePreview,
  onOpenStats,
  onOpenCodebook,
  onOpenExcelAddin,
  onRefreshFromExcel,
  onRestoreRawBackup,
}: RibbonToolbarProps) {
  const fileInputRef = useRef<HTMLInputElement>(null);

  return (
    <div className="bg-slate-100/90 border-b border-slate-200 px-4 py-2 flex flex-wrap items-center justify-between gap-3 shrink-0">
      {/* Left group: Dataset selector, Undo/Redo, Value Labels Toggle */}
      <div className="flex flex-wrap items-center gap-2">
        {host === 'simulator' ? (
          <div className="flex items-center gap-1.5">
            <label htmlFor="dataset-preset" className="sr-only">
              Select Dataset
            </label>
            <div className="relative flex items-center">
              <FolderOpen className="w-3.5 h-3.5 text-slate-500 absolute left-2.5 pointer-events-none" />
              <select
                id="dataset-preset"
                value=""
                onChange={(e) => {
                  if (e.target.value) onLoadPreset(e.target.value);
                }}
                aria-label="Load sample dataset"
                className="pl-8 pr-7 py-1.5 text-xs font-medium text-slate-800 bg-white border border-slate-300 rounded-md hover:border-slate-400 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-blue-600 cursor-pointer max-w-[200px] truncate"
              >
                <option value="" disabled>
                  {datasetName}
                </option>
                {SAMPLE_PRESETS.map((p) => (
                  <option key={p.id} value={p.id}>
                    {p.title}
                  </option>
                ))}
              </select>
            </div>

            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx,.xls,.csv"
              className="hidden"
              onChange={(e) => {
                const file = e.target.files?.[0];
                if (file) {
                  onImportWorkbook(file);
                  e.target.value = '';
                }
              }}
            />
            <button
              type="button"
              onClick={() => fileInputRef.current?.click()}
              className="inline-flex items-center gap-1.5 px-2.5 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-300 rounded-md hover:bg-slate-50 transition-colors whitespace-nowrap shrink-0 cursor-pointer focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-blue-600"
              title="Import .xlsx or .csv workbook"
            >
              <Upload className="w-3.5 h-3.5 text-slate-500" />
              <span>Import</span>
            </button>
          </div>
        ) : (
          <div className="flex items-center gap-1.5">
            <button
              type="button"
              onClick={onRefreshFromExcel}
              className="inline-flex items-center gap-1.5 px-2.5 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-300 rounded-md hover:bg-slate-50 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
              title="Read active Excel worksheet into StatSheet"
            >
              <FileSpreadsheet className="w-3.5 h-3.5 text-[#107C41]" />
              <span>Read Sheet</span>
            </button>
            <button
              type="button"
              onClick={onRestoreRawBackup}
              className="inline-flex items-center gap-1.5 px-2.5 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-300 rounded-md hover:bg-slate-50 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
              title="Restore original numeric codes from _SPSS_RawBackup"
            >
              <RotateCcw className="w-3.5 h-3.5 text-slate-500" />
              <span>Restore Codes</span>
            </button>
          </div>
        )}

        <div className="h-5 w-px bg-slate-300 mx-0.5 hidden sm:block" aria-hidden="true" />

        {/* Undo / Redo */}
        <div className="flex items-center gap-1">
          <button
            type="button"
            onClick={onUndo}
            disabled={!canUndo}
            aria-label="Undo"
            title="Undo last edit"
            className="p-1.5 text-slate-700 bg-white border border-slate-300 rounded-md hover:bg-slate-50 disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-blue-600"
          >
            <Undo2 className="w-3.5 h-3.5" />
          </button>
          <button
            type="button"
            onClick={onRedo}
            disabled={!canRedo}
            aria-label="Redo"
            title="Redo edit"
            className="p-1.5 text-slate-700 bg-white border border-slate-300 rounded-md hover:bg-slate-50 disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-blue-600"
          >
            <Redo2 className="w-3.5 h-3.5" />
          </button>
        </div>

        <div className="h-5 w-px bg-slate-300 mx-0.5 hidden sm:block" aria-hidden="true" />

        {/* Signature SPSS Value Labels Toggle (1 ↔ A) */}
        <button
          type="button"
          onClick={onToggleValueLabels}
          aria-pressed={showValueLabels}
          className={`inline-flex items-center gap-2 px-3 py-1.5 text-xs font-semibold rounded-md border transition-colors whitespace-nowrap shrink-0 cursor-pointer focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-blue-600 ${
            showValueLabels
              ? 'bg-blue-700 text-white border-blue-700 shadow-xs'
              : 'bg-white text-slate-800 border-slate-300 hover:bg-slate-50'
          }`}
          title="Toggle between raw codes and defined Value Labels (reversible backup preserved)"
        >
          <span
            className={`font-mono text-[11px] px-1 py-0.5 rounded border leading-none ${
              showValueLabels
                ? 'bg-blue-600 border-blue-500 text-white'
                : 'bg-slate-100 border-slate-300 text-slate-700'
            }`}
          >
            1↔A
          </span>
          <Tags className="w-3.5 h-3.5" />
          <span>Value Labels: {showValueLabels ? 'On' : 'Off'}</span>
        </button>

        {/* Insert Case / Variable */}
        {activeView === 'data' ? (
          <button
            type="button"
            onClick={onAddRow}
            className="inline-flex items-center gap-1.5 px-2.5 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-300 rounded-md hover:bg-slate-50 transition-colors whitespace-nowrap shrink-0 cursor-pointer focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-blue-600"
          >
            <Plus className="w-3.5 h-3.5 text-slate-600" />
            <span>Insert Case</span>
          </button>
        ) : (
          <button
            type="button"
            onClick={onAddVariable}
            className="inline-flex items-center gap-1.5 px-2.5 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-300 rounded-md hover:bg-slate-50 transition-colors whitespace-nowrap shrink-0 cursor-pointer focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-blue-600"
          >
            <Columns3 className="w-3.5 h-3.5 text-slate-600" />
            <span>Add Variable</span>
          </button>
        )}
      </div>

      {/* Right group: Search, Validation Issue Filter, Stats, Task Pane Preview */}
      <div className="flex flex-wrap items-center gap-2">
        <div className="relative">
          <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-1/2 -translate-y-1/2 pointer-events-none" />
          <input
            type="search"
            value={searchQuery}
            onChange={(e) => onSearchChange(e.target.value)}
            placeholder={
              activeView === 'data'
                ? 'Filter cases or labels...'
                : 'Filter variables...'
            }
            aria-label="Filter grid"
            className="pl-8 pr-3 py-1.5 text-xs bg-white border border-slate-300 rounded-md w-44 sm:w-52 text-slate-800 placeholder:text-slate-400 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-blue-600"
          />
        </div>

        {activeView === 'data' && (
          <button
            type="button"
            onClick={onToggleIssuesOnly}
            aria-pressed={issuesOnly}
            className={`inline-flex items-center gap-1.5 px-2.5 py-1.5 text-xs font-medium rounded-md border transition-colors whitespace-nowrap shrink-0 cursor-pointer focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-blue-600 ${
              issuesOnly
                ? 'bg-amber-50 text-amber-900 border-amber-300'
                : 'bg-white text-slate-700 border-slate-300 hover:bg-slate-50'
            }`}
            title="Filter rows containing missing-value flags or validation errors"
          >
            <AlertTriangle
              className={`w-3.5 h-3.5 ${
                issueCount > 0 ? 'text-amber-600' : 'text-slate-400'
              }`}
            />
            <span className="tabular-nums">Issues ({issueCount})</span>
          </button>
        )}

        <button
          type="button"
          onClick={onOpenStats}
          className="inline-flex items-center gap-1.5 px-2.5 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-300 rounded-md hover:bg-slate-50 transition-colors whitespace-nowrap shrink-0 cursor-pointer focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-blue-600"
        >
          <BarChart3 className="w-3.5 h-3.5 text-blue-700" />
          <span>Descriptives</span>
        </button>

        <button
          type="button"
          onClick={onOpenCodebook}
          className={`${
            taskPanePreview ? 'inline-flex' : 'inline-flex sm:hidden'
          } items-center gap-1.5 px-2.5 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-300 rounded-md hover:bg-slate-50 transition-colors whitespace-nowrap shrink-0 cursor-pointer focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-blue-600`}
        >
          <BookOpen className="w-3.5 h-3.5 text-slate-600" />
          <span>Codebook</span>
        </button>

        <button
          type="button"
          onClick={onOpenExcelAddin}
          className={`${
            taskPanePreview ? 'inline-flex' : 'inline-flex sm:hidden'
          } items-center gap-1.5 px-2.5 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-300 rounded-md hover:bg-slate-50 transition-colors whitespace-nowrap shrink-0 cursor-pointer focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-blue-600`}
        >
          <FileCode2 className="w-3.5 h-3.5 text-[#107C41]" />
          <span>Add-in</span>
        </button>

        {host === 'simulator' && (
          <button
            type="button"
            onClick={onToggleTaskPanePreview}
            aria-pressed={taskPanePreview}
            className={`inline-flex items-center gap-1.5 px-2.5 py-1.5 text-xs font-medium rounded-md border transition-colors whitespace-nowrap shrink-0 cursor-pointer focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-blue-600 ${
              taskPanePreview
                ? 'bg-slate-900 text-white border-slate-900'
                : 'bg-white text-slate-700 border-slate-300 hover:bg-slate-50'
            }`}
            title="Preview how StatSheet fits inside Excel's 400px Task Pane"
          >
            <PanelRight className="w-3.5 h-3.5" />
            <span>Task Pane Preview</span>
          </button>
        )}
      </div>
    </div>
  );
}
