import { X } from 'lucide-react';
import { computeVariableStats } from '../engine/spssEngine';
import type { DataRow, VariableDefinition } from '../types/spss';
import { MeasureIcon } from './MeasureIcon';

interface StatsInspectorDrawerProps {
  variables: VariableDefinition[];
  rows: DataRow[];
  selectedVariableId: string;
  onSelectVariable: (variableId: string) => void;
  onClose: () => void;
}

export function StatsInspectorDrawer({
  variables,
  rows,
  selectedVariableId,
  onSelectVariable,
  onClose,
}: StatsInspectorDrawerProps) {
  const variable =
    variables.find((v) => v.id === selectedVariableId) ?? variables[0];

  if (!variable) return null;

  const stats = computeVariableStats(variable, rows);
  const fmt = (n: number | null) =>
    n === null ? '—' : n.toFixed(Math.min(2, Math.max(1, variable.decimals + 1)));

  const maxCount = Math.max(1, ...stats.frequencies.map((f) => f.count));

  return (
    <div
      className="fixed inset-0 z-50 flex justify-end bg-slate-900/30"
      role="dialog"
      aria-modal="true"
      aria-labelledby="stats-drawer-title"
    >
      <div className="bg-white border-l border-slate-300 w-full max-w-xl h-full flex flex-col shadow-2xl">
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
        <div className="px-5 py-3 bg-slate-50 border-b border-slate-200 flex items-center justify-between gap-3">
          <label
            htmlFor="stats-var-select"
            className="text-xs font-semibold text-slate-700 shrink-0"
          >
            Analyze Variable
          </label>
          <div className="flex items-center gap-2 flex-1 max-w-sm">
            <MeasureIcon measure={variable.measure} />
            <select
              id="stats-var-select"
              value={variable.id}
              onChange={(e) => onSelectVariable(e.target.value)}
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
            <h3 className="font-display text-xs font-semibold text-slate-900 mb-2.5">
              Frequency Table
            </h3>
            {stats.frequencies.length === 0 ? (
              <div className="p-6 border border-slate-200 rounded-md text-center text-xs text-slate-500">
                No non-blank values recorded for this variable.
              </div>
            ) : (
              <div className="border border-slate-200 rounded-md overflow-hidden">
                <table className="w-full border-collapse text-xs">
                  <thead className="bg-slate-50 border-b border-slate-200 text-slate-600 font-semibold">
                    <tr>
                      <th className="px-3 py-2 text-left">Value / Label</th>
                      <th className="px-2.5 py-2 text-right">Freq</th>
                      <th className="px-2.5 py-2 text-right">Percent</th>
                      <th className="px-2.5 py-2 text-right">Valid %</th>
                      <th className="px-2.5 py-2 text-right">Cum. %</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-200">
                    {stats.frequencies.map((row) => (
                      <tr
                        key={row.rawKey}
                        className={row.isMissing ? 'bg-amber-50/50' : 'bg-white'}
                      >
                        <td className="px-3 py-2">
                          <div className="flex items-center justify-between gap-2">
                            <span className="font-medium text-slate-800 truncate">
                              {row.label !== row.rawKey ? (
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
