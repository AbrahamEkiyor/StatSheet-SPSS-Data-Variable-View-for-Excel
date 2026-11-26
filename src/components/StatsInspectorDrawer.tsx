import { ArrowRightLeft, SlidersHorizontal, X } from 'lucide-react';
import { useState } from 'react';
import {
  computeGroupedVariableStats,
  computeVariableStats,
  formatClassBinsSummary,
  isNumericType,
  sanitizeVariableName,
  validateVariableName,
} from '../engine/spssEngine';
import type { ClassBinSpec, DataRow, VariableDefinition } from '../types/spss';
import { MeasureIcon } from './MeasureIcon';

interface StatsInspectorDrawerProps {
  variables: VariableDefinition[];
  rows: DataRow[];
  selectedVariableId: string;
  onSelectVariable: (variableId: string) => void;
  onOpenClassBins: (variableId: string) => void;
  onRecodeVariable: (
    sourceVarId: string,
    targetName: string,
    spec: ClassBinSpec
  ) => void;
  onClose: () => void;
}

export function StatsInspectorDrawer({
  variables,
  rows,
  selectedVariableId,
  onSelectVariable,
  onOpenClassBins,
  onRecodeVariable,
  onClose,
}: StatsInspectorDrawerProps) {
  const variable =
    variables.find((v) => v.id === selectedVariableId) ?? variables[0];

  const [modeByVar, setModeByVar] = useState<
    Record<string, 'ungrouped' | 'grouped'>
  >({});
  const [showRecodeInput, setShowRecodeInput] = useState(false);
  const [recodeName, setRecodeName] = useState('');
  const [recodeError, setRecodeError] = useState<string | null>(null);

  if (!variable) return null;

  const numeric = isNumericType(variable.type);
  const hasBins = Boolean(
    variable.classBins?.enabled && variable.classBins.intervals.length > 0
  );
  const activeTab: 'ungrouped' | 'grouped' =
    numeric && (modeByVar[variable.id] ? modeByVar[variable.id] === 'grouped' : hasBins)
      ? 'grouped'
      : 'ungrouped';

  const stats =
    activeTab === 'grouped' && hasBins && variable.classBins
      ? computeGroupedVariableStats(variable, rows, variable.classBins)
      : computeVariableStats(variable, rows);

  const fmt = (n: number | null) =>
    n === null ? '—' : n.toFixed(Math.min(2, Math.max(1, variable.decimals + 1)));

  const maxCount = Math.max(1, ...stats.frequencies.map((f) => f.count));

  const openRecodePrompt = () => {
    setRecodeName(
      sanitizeVariableName(
        `${variable.name}_grp`,
        variables.map((v) => v.name)
      )
    );
    setRecodeError(null);
    setShowRecodeInput((s) => !s);
  };

  const submitRecode = () => {
    if (!variable.classBins) return;
    const err = validateVariableName(
      recodeName,
      variables.map((v) => v.name)
    );
    if (err) {
      setRecodeError(err);
      return;
    }
    onRecodeVariable(variable.id, recodeName.trim(), variable.classBins);
    setShowRecodeInput(false);
  };

  return (
    <div
      className="fixed inset-0 z-50 flex justify-end bg-slate-900/30"
      role="dialog"
      aria-modal="true"
      aria-labelledby="stats-drawer-title"
    >
      <div className="bg-white border-l border-slate-300 w-full max-w-2xl h-full flex flex-col shadow-2xl">
        {/* Header */}
        <div className="px-5 py-4 border-b border-slate-200 flex items-center justify-between gap-4">
          <div className="min-w-0">
            <h2
              id="stats-drawer-title"
              className="font-display text-base font-semibold text-slate-900"
            >
              Descriptives &amp; Frequencies
            </h2>
            <p className="text-xs text-slate-500 mt-0.5 truncate">
              Excludes system-missing and user-defined missing codes from valid statistics.
            </p>
          </div>
          <button
            type="button"
            onClick={onClose}
            aria-label="Close statistics drawer"
            className="p-1.5 text-slate-400 hover:text-slate-700 rounded cursor-pointer"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        {/* Variable Selector */}
        <div className="px-5 py-3 bg-slate-50 border-b border-slate-200 flex flex-wrap items-center justify-between gap-3">
          <label
            htmlFor="stats-var-select"
            className="text-xs font-semibold text-slate-700 shrink-0"
          >
            Analyze Variable
          </label>
          <div className="flex items-center gap-2 flex-1 max-w-md">
            <MeasureIcon measure={variable.measure} />
            <select
              id="stats-var-select"
              value={variable.id}
              onChange={(e) => {
                setShowRecodeInput(false);
                onSelectVariable(e.target.value);
              }}
              className="w-full px-2.5 py-1.5 text-xs font-medium bg-white border border-slate-300 rounded-md text-slate-900 focus:outline-none focus:ring-2 focus:ring-blue-600 cursor-pointer"
            >
              {variables.map((v) => (
                <option key={v.id} value={v.id}>
                  {v.name} {v.label ? `— ${v.label}` : ''}
                </option>
              ))}
            </select>
          </div>
        </div>

        {/* Grouping Mode Bar for Numeric / Scale Variables */}
        {numeric && (
          <div className="px-5 py-2.5 bg-white border-b border-slate-200 flex flex-wrap items-center justify-between gap-3 text-xs">
            <div className="flex items-center gap-1 p-0.5 bg-slate-100 rounded-md">
              <button
                type="button"
                onClick={() =>
                  setModeByVar((prev) => ({ ...prev, [variable.id]: 'ungrouped' }))
                }
                className={`px-2.5 py-1 text-[11px] font-medium rounded cursor-pointer ${
                  activeTab === 'ungrouped'
                    ? 'bg-white text-slate-900 shadow-xs'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                Ungrouped Values
              </button>
              <button
                type="button"
                onClick={() =>
                  setModeByVar((prev) => ({ ...prev, [variable.id]: 'grouped' }))
                }
                className={`px-2.5 py-1 text-[11px] font-medium rounded cursor-pointer ${
                  activeTab === 'grouped'
                    ? 'bg-white text-slate-900 shadow-xs'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                Class Intervals{hasBins ? ` (${variable.classBins!.intervals.length})` : ''}
              </button>
            </div>

            <div className="flex items-center gap-2">
              <button
                type="button"
                onClick={() => onOpenClassBins(variable.id)}
                className="inline-flex items-center gap-1.5 px-2.5 py-1 text-[11px] font-semibold text-blue-700 bg-blue-50/70 hover:bg-blue-100 border border-blue-200 rounded cursor-pointer"
              >
                <SlidersHorizontal className="w-3 h-3" />
                <span>{hasBins ? 'Edit Class Boundaries' : 'Define Class Boundaries'}</span>
              </button>
              {hasBins && (
                <button
                  type="button"
                  onClick={openRecodePrompt}
                  className="inline-flex items-center gap-1.5 px-2.5 py-1 text-[11px] font-medium text-slate-700 bg-white border border-slate-300 hover:bg-slate-50 rounded cursor-pointer"
                >
                  <ArrowRightLeft className="w-3 h-3 text-[#107C41]" />
                  <span>Recode Variable</span>
                </button>
              )}
            </div>
          </div>
        )}

        {/* Inline Recode Prompt inside Drawer */}
        {showRecodeInput && hasBins && (
          <div className="px-5 py-3 bg-blue-50/60 border-b border-blue-200 flex flex-wrap items-center justify-between gap-2 text-xs">
            <div className="flex items-center gap-2 flex-1 min-w-[200px]">
              <label
                htmlFor="drawer-recode-name"
                className="font-semibold text-slate-800 shrink-0"
              >
                New Variable:
              </label>
              <input
                id="drawer-recode-name"
                type="text"
                value={recodeName}
                onChange={(e) => {
                  setRecodeName(e.target.value);
                  setRecodeError(null);
                }}
                className="px-2 py-1 font-mono text-xs bg-white border border-slate-300 rounded focus:outline-none focus:ring-2 focus:ring-blue-600"
              />
              {recodeError && (
                <span className="text-[11px] text-red-700">{recodeError}</span>
              )}
            </div>
            <div className="flex items-center gap-2">
              <button
                type="button"
                onClick={() => setShowRecodeInput(false)}
                className="px-2.5 py-1 text-[11px] font-medium text-slate-700 bg-white border border-slate-300 rounded cursor-pointer"
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={submitRecode}
                className="px-2.5 py-1 text-[11px] font-semibold text-white bg-[#107C41] hover:bg-[#0D6635] rounded cursor-pointer"
              >
                Create Ordinal Column
              </button>
            </div>
          </div>
        )}

        {/* Body */}
        <div className="flex-1 overflow-y-auto p-5 space-y-6">
          {/* Case Processing Summary */}
          <div>
            <h3 className="font-display text-xs font-semibold text-slate-900 mb-2.5">
              Case Processing Summary
            </h3>
            <div className="grid grid-cols-4 border border-slate-200 rounded-md divide-x divide-slate-200 bg-slate-50/50">
              <div className="p-3">
                <div className="text-[11px] text-slate-500">Total N</div>
                <div className="mt-1 font-mono text-base font-semibold text-slate-900 tabular-nums">
                  {stats.totalN}
                </div>
              </div>
              <div className="p-3">
                <div className="text-[11px] text-slate-500">Valid</div>
                <div className="mt-1 font-mono text-base font-semibold text-emerald-700 tabular-nums">
                  {stats.validN}
                </div>
              </div>
              <div className="p-3">
                <div className="text-[11px] text-slate-500">User Missing</div>
                <div className="mt-1 font-mono text-base font-semibold text-amber-700 tabular-nums">
                  {stats.userMissingN}
                </div>
              </div>
              <div className="p-3">
                <div className="text-[11px] text-slate-500">Sys. Missing</div>
                <div className="mt-1 font-mono text-base font-semibold text-slate-600 tabular-nums">
                  {stats.systemMissingN}
                </div>
              </div>
            </div>
          </div>

          {/* Numeric Descriptives */}
          {stats.mean !== null && (
            <div>
              <h3 className="font-display text-xs font-semibold text-slate-900 mb-2.5">
                Descriptive Statistics (Valid N = {stats.validN})
              </h3>
              <div className="grid grid-cols-5 border border-slate-200 rounded-md divide-x divide-slate-200 bg-white">
                <div className="p-3">
                  <div className="text-[11px] text-slate-500">Mean</div>
                  <div className="mt-1 font-mono text-sm font-semibold text-slate-900 tabular-nums">
                    {fmt(stats.mean)}
                  </div>
                </div>
                <div className="p-3">
                  <div className="text-[11px] text-slate-500">Median</div>
                  <div className="mt-1 font-mono text-sm font-semibold text-slate-900 tabular-nums">
                    {fmt(stats.median)}
                  </div>
                </div>
                <div className="p-3">
                  <div className="text-[11px] text-slate-500">Std. Dev.</div>
                  <div className="mt-1 font-mono text-sm font-semibold text-slate-900 tabular-nums">
                    {fmt(stats.stdDev)}
                  </div>
                </div>
                <div className="p-3">
                  <div className="text-[11px] text-slate-500">Min</div>
                  <div className="mt-1 font-mono text-sm font-semibold text-slate-900 tabular-nums">
                    {stats.min}
                  </div>
                </div>
                <div className="p-3">
                  <div className="text-[11px] text-slate-500">Max</div>
                  <div className="mt-1 font-mono text-sm font-semibold text-slate-900 tabular-nums">
                    {stats.max}
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* Frequency Table */}
          <div>
            <div className="flex items-center justify-between gap-2 mb-2.5">
              <h3 className="font-display text-xs font-semibold text-slate-900">
                {activeTab === 'grouped' && hasBins
                  ? 'Grouped Frequency Distribution'
                  : 'Frequency Table'}
              </h3>
              {activeTab === 'grouped' && hasBins && (
                <span className="text-[11px] font-mono text-slate-500">
                  {formatClassBinsSummary(variable.classBins)}
                </span>
              )}
            </div>

            {activeTab === 'grouped' && !hasBins ? (
              <div className="p-6 border border-dashed border-slate-300 rounded-md text-center space-y-2.5 bg-slate-50/50">
                <p className="text-xs font-medium text-slate-800">
                  No class boundaries defined for <span className="font-mono">{variable.name}</span> yet.
                </p>
                <p className="text-[11px] text-slate-500 max-w-sm mx-auto">
                  Define equal-width bins or custom intervals to group continuous scale observations into a frequency distribution.
                </p>
                <button
                  type="button"
                  onClick={() => onOpenClassBins(variable.id)}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-white bg-blue-700 hover:bg-blue-800 rounded-md cursor-pointer"
                >
                  <SlidersHorizontal className="w-3.5 h-3.5" />
                  <span>Define Class Boundaries</span>
                </button>
              </div>
            ) : stats.frequencies.length === 0 ? (
              <div className="p-6 border border-slate-200 rounded-md text-center text-xs text-slate-500">
                No non-blank values recorded for this variable.
              </div>
            ) : (
              <div className="border border-slate-200 rounded-md overflow-hidden">
                <table className="w-full border-collapse text-xs">
                  <thead className="bg-slate-50 border-b border-slate-200 text-slate-600 font-semibold">
                    <tr>
                      <th className="px-3 py-2 text-left">
                        {activeTab === 'grouped' ? 'Class Interval' : 'Value / Label'}
                      </th>
                      {activeTab === 'grouped' && (
                        <th className="px-2.5 py-2 text-right">Midpoint</th>
                      )}
                      <th className="px-2.5 py-2 text-right">Freq</th>
                      <th className="px-2.5 py-2 text-right">Percent</th>
                      <th className="px-2.5 py-2 text-right">Valid %</th>
                      <th className="px-2.5 py-2 text-right">Cum. %</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-200">
                    {stats.frequencies.map((row) => (
                      <tr
                        key={`${row.intervalId ?? row.rawKey}_${row.isMissing ? 'm' : 'v'}`}
                        className={row.isMissing ? 'bg-amber-50/50' : 'bg-white'}
                      >
                        <td className="px-3 py-2">
                          <div className="flex items-center justify-between gap-2">
                            <span className="font-medium text-slate-800 truncate">
                              {activeTab === 'grouped' ? (
                                <span className="font-mono">{row.label}</span>
                              ) : row.label !== row.rawKey ? (
                                <>
                                  <span className="font-mono text-slate-500 mr-1.5">
                                    [{row.rawKey}]
                                  </span>
                                  {row.label}
                                </>
                              ) : (
                                <span className="font-mono">{row.rawKey}</span>
                              )}
                            </span>
                            {row.isMissing && (
                              <span className="text-[10px] font-semibold text-amber-800 shrink-0">
                                Missing
                              </span>
                            )}
                          </div>
                          <div className="mt-1 h-1.5 w-full bg-slate-100 rounded-full overflow-hidden">
                            <div
                              className={`h-full rounded-full ${
                                row.isMissing ? 'bg-amber-500' : 'bg-blue-600'
                              }`}
                              style={{
                                width: `${Math.round((row.count / maxCount) * 100)}%`,
                              }}
                            />
                          </div>
                        </td>
                        {activeTab === 'grouped' && (
                          <td className="px-2.5 py-2 text-right font-mono tabular-nums text-slate-600">
                            {row.midpoint !== null && row.midpoint !== undefined
                              ? row.midpoint.toLocaleString('en-US')
                              : '—'}
                          </td>
                        )}
                        <td className="px-2.5 py-2 text-right font-mono tabular-nums text-slate-800">
                          {row.count}
                        </td>
                        <td className="px-2.5 py-2 text-right font-mono tabular-nums text-slate-600">
                          {row.percent.toFixed(1)}%
                        </td>
                        <td className="px-2.5 py-2 text-right font-mono tabular-nums text-slate-800">
                          {row.validPercent !== null
                            ? `${row.validPercent.toFixed(1)}%`
                            : '—'}
                        </td>
                        <td className="px-2.5 py-2 text-right font-mono tabular-nums text-slate-600">
                          {row.cumulativePercent !== null
                            ? `${row.cumulativePercent.toFixed(1)}%`
                            : '—'}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}
