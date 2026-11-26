import { AlertCircle, AlertTriangle, BarChart3, Plus, Trash2 } from 'lucide-react';
import { useEffect, useRef, useState } from 'react';
import {
  isSystemMissing,
  lookupValueLabel,
  parseUserInputToRaw,
  resolveDisplayValue,
} from '../engine/spssEngine';
import type {
  CellIssue,
  CellValue,
  DataRow,
  VariableDefinition,
} from '../types/spss';
import { MeasureIcon } from './MeasureIcon';

interface DataViewGridProps {
  variables: VariableDefinition[];
  rows: DataRow[];
  showValueLabels: boolean;
  issues: Map<string, CellIssue>;
  searchQuery: string;
  issuesOnly: boolean;
  selectedCell: { rowId: string; variableId: string } | null;
  onSelectCell: (cell: { rowId: string; variableId: string } | null) => void;
  onUpdateCell: (rowId: string, variableName: string, nextValue: CellValue) => void;
  onDeleteRow: (rowId: string) => void;
  onAddRow: () => void;
  onJumpToVariable: (variableId: string) => void;
  onInspectVariableStats: (variableId: string) => void;
}

export function DataViewGrid({
  variables,
  rows,
  showValueLabels,
  issues,
  searchQuery,
  issuesOnly,
  selectedCell,
  onSelectCell,
  onUpdateCell,
  onDeleteRow,
  onAddRow,
  onJumpToVariable,
  onInspectVariableStats,
}: DataViewGridProps) {
  const [editingCell, setEditingCell] = useState<{
    rowId: string;
    variableId: string;
  } | null>(null);
  const [draftValue, setDraftValue] = useState<string>('');
  const inputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    if (editingCell && inputRef.current) {
      inputRef.current.focus();
      inputRef.current.select();
    }
  }, [editingCell]);

  const query = searchQuery.trim().toLowerCase();

  const filteredRows = rows
    .map((row, originalIndex) => ({ row, originalIndex }))
    .filter(({ row }) => {
      if (issuesOnly) {
        const rowHasIssue = variables.some((v) =>
          issues.has(`${row.id}::${v.id}`)
        );
        if (!rowHasIssue) return false;
      }
      if (!query) return true;
      return variables.some((v) => {
        const raw = row.values[v.name];
        const resolved = resolveDisplayValue(raw, v, true);
        return (
          String(raw ?? '').toLowerCase().includes(query) ||
          resolved.text.toLowerCase().includes(query)
        );
      });
    });

  const activeVariable = selectedCell
    ? variables.find((v) => v.id === selectedCell.variableId) ?? null
    : null;
  const activeRowEntry = selectedCell
    ? rows.find((r) => r.id === selectedCell.rowId) ?? null
    : null;
  const activeRawValue =
    activeVariable && activeRowEntry
      ? activeRowEntry.values[activeVariable.name]
      : undefined;
  const activeIssue = selectedCell
    ? issues.get(`${selectedCell.rowId}::${selectedCell.variableId}`) ?? null
    : null;

  const startEditing = (row: DataRow, variable: VariableDefinition) => {
    const raw = row.values[variable.name];
    setEditingCell({ rowId: row.id, variableId: variable.id });
    setDraftValue(isSystemMissing(raw) ? '' : String(raw));
  };

  const commitTextEdit = (row: DataRow, variable: VariableDefinition) => {
    const parsed = parseUserInputToRaw(draftValue, variable);
    onUpdateCell(row.id, variable.name, parsed);
    setEditingCell(null);
  };

  if (variables.length === 0) {
    return (
      <div className="flex-1 flex flex-col items-center justify-center p-10 text-center bg-white">
        <h3 className="font-display text-base font-semibold text-slate-900">
          No variables defined yet
        </h3>
        <p className="mt-1.5 text-sm text-slate-500 max-w-md">
          Switch to Variable View to define columns, value labels, and missing value rules, or import an existing Excel workbook.
        </p>
      </div>
    );
  }

  return (
    <div className="flex-1 flex flex-col min-h-0 bg-white">
      {/* Formula / Active Cell Inspector Bar */}
      <div className="px-4 py-2 bg-slate-50 border-b border-slate-200 flex flex-wrap items-center justify-between gap-3 text-xs shrink-0">
        <div className="flex items-center gap-3 min-w-0">
          <span className="font-mono font-medium text-slate-700 bg-white border border-slate-300 px-2 py-0.5 rounded tabular-nums shrink-0">
            {activeVariable && selectedCell
              ? `${activeVariable.name} · Case ${
                  rows.findIndex((r) => r.id === selectedCell.rowId) + 1
                }`
              : 'No cell selected'}
          </span>

          {activeVariable && activeRowEntry ? (
            <div className="flex items-center gap-2 text-slate-600 truncate">
              <span className="truncate font-medium text-slate-800">
                {activeVariable.label || activeVariable.name}
              </span>
              <span aria-hidden="true">·</span>
              <span className="font-mono tabular-nums text-slate-700">
                Raw:{' '}
                {isSystemMissing(activeRawValue) ? '(sysmis)' : String(activeRawValue)}
              </span>
              {lookupValueLabel(activeRawValue, activeVariable) !== null && (
                <>
                  <span aria-hidden="true">·</span>
                  <span className="text-blue-700 font-medium">
                    Label: {lookupValueLabel(activeRawValue, activeVariable)}
                  </span>
                </>
              )}
            </div>
          ) : (
            <span className="text-slate-500">
              Click any cell to inspect raw code vs. label, or double-click to edit. Double-click a column header to edit its variable schema.
            </span>
          )}
        </div>

        {activeIssue && (
          <div
            className={`inline-flex items-center gap-1.5 font-medium ${
              activeIssue.kind === 'missing_user'
                ? 'text-amber-800'
                : 'text-red-700'
            }`}
          >
            {activeIssue.kind === 'missing_user' ? (
              <AlertTriangle className="w-3.5 h-3.5 shrink-0" />
            ) : (
              <AlertCircle className="w-3.5 h-3.5 shrink-0" />
            )}
            <span>{activeIssue.message}</span>
          </div>
        )}
      </div>

      {/* Main Spreadsheet Viewport */}
      <div className="flex-1 overflow-auto grid-scroll">
        {filteredRows.length === 0 ? (
          <div className="flex flex-col items-center justify-center py-16 px-6 text-center">
            <p className="text-sm font-medium text-slate-800">
              {rows.length === 0
                ? 'This worksheet has no cases yet.'
                : 'No cases match the current filter.'}
            </p>
            <p className="mt-1 text-xs text-slate-500">
              {rows.length === 0
                ? 'Insert a case row to begin entering coded or numeric values.'
                : 'Clear the search box or Issues filter to view all cases.'}
            </p>
            {rows.length === 0 && (
              <button
                type="button"
                onClick={onAddRow}
                className="mt-4 inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-white bg-blue-700 hover:bg-blue-800 rounded-md cursor-pointer"
              >
                <Plus className="w-3.5 h-3.5" />
                <span>Insert First Case</span>
              </button>
            )}
          </div>
        ) : (
          <table className="border-collapse w-max min-w-full text-[13px]">
            <thead className="sticky top-0 z-20 bg-slate-100 border-b border-slate-300">
              <tr>
                <th
                  scope="col"
                  className="sticky left-0 z-30 w-14 min-w-14 bg-slate-100 border-r border-slate-300 px-2 py-2 text-center text-[11px] font-semibold text-slate-500 font-mono tabular-nums"
                >
                  Case
                </th>
                {variables.map((variable) => {
                  const hasLabels = Object.keys(variable.valueLabels).length > 0;
                  return (
                    <th
                      key={variable.id}
                      scope="col"
                      style={{
                        width: `${Math.max(104, variable.columns)}px`,
                        minWidth: `${Math.max(104, variable.columns)}px`,
                      }}
                      onDoubleClick={() => onJumpToVariable(variable.id)}
                      title={`${variable.label || variable.name} (${variable.measure} · ${
                        hasLabels
                          ? `${Object.keys(variable.valueLabels).length} value labels`
                          : 'no value labels'
                      }) — Double-click to open in Variable View`}
                      className="group border-r border-slate-300 px-2.5 py-2 text-left font-semibold text-slate-800 select-none hover:bg-slate-200/60 transition-colors cursor-pointer"
                    >
                      <div className="flex items-center justify-between gap-1.5">
                        <div className="flex items-center gap-1.5 min-w-0">
                          <MeasureIcon measure={variable.measure} className="w-3.5 h-3.5" />
                          <span className="truncate font-display text-xs">
                            {variable.name}
                          </span>
                        </div>
                        <button
                          type="button"
                          onClick={(e) => {
                            e.stopPropagation();
                            onInspectVariableStats(variable.id);
                          }}
                          title={`Inspect statistics for ${variable.name}`}
                          className="opacity-0 group-hover:opacity-100 focus:opacity-100 p-0.5 text-slate-500 hover:text-blue-700 rounded cursor-pointer transition-opacity"
                        >
                          <BarChart3 className="w-3.5 h-3.5" />
                        </button>
                      </div>
                      {variable.label && (
                        <div className="text-[11px] font-normal text-slate-500 truncate mt-0.5">
                          {variable.label}
                        </div>
                      )}
                    </th>
                  );
                })}
                <th
                  scope="col"
                  className="w-10 min-w-10 bg-slate-100 border-b border-slate-300"
                  aria-label="Row actions"
                />
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200">
              {filteredRows.map(({ row, originalIndex }) => (
                <tr
                  key={row.id}
                  className="group h-9 hover:bg-slate-50/80 transition-colors"
                >
                  <td className="sticky left-0 z-10 bg-slate-50 group-hover:bg-slate-100 border-r border-slate-300 px-2 text-center font-mono text-xs text-slate-500 tabular-nums select-none">
                    {originalIndex + 1}
                  </td>
                  {variables.map((variable) => {
                    const raw = row.values[variable.name];
                    const cellKey = `${row.id}::${variable.id}`;
                    const issue = issues.get(cellKey);
                    const isSelected =
                      selectedCell?.rowId === row.id &&
                      selectedCell?.variableId === variable.id;
                    const isEditing =
                      editingCell?.rowId === row.id &&
                      editingCell?.variableId === variable.id;

                    const { text, isLabeled } = resolveDisplayValue(
                      raw,
                      variable,
                      showValueLabels
                    );
                    const hasValueLabels =
                      Object.keys(variable.valueLabels).length > 0;

                    const alignClass =
                      isLabeled || variable.align === 'left'
                        ? 'text-left'
                        : variable.align === 'center'
                          ? 'text-center'
                          : 'text-right';

                    const issueBg =
                      issue?.kind === 'missing_user'
                        ? 'bg-amber-50/80'
                        : issue
                          ? 'bg-red-50/80'
                          : '';

                    return (
                      <td
                        key={variable.id}
                        onClick={() =>
                          onSelectCell({ rowId: row.id, variableId: variable.id })
                        }
                        onDoubleClick={() => startEditing(row, variable)}
                        title={
                          issue
                            ? issue.message
                            : isLabeled
                              ? `Label "${text}" · Raw code: ${raw}`
                              : undefined
                        }
                        className={`relative border-r border-slate-200 px-2.5 py-0 h-9 cursor-cell ${alignClass} ${issueBg} ${
                          isSelected
                            ? 'ring-2 ring-inset ring-blue-700 bg-blue-50/40 z-10'
                            : ''
                        }`}
                      >
                        {isEditing ? (
                          showValueLabels && hasValueLabels ? (
                            <select
                              autoFocus
                              value={isSystemMissing(raw) ? '' : String(raw)}
                              onChange={(e) => {
                                const val = e.target.value;
                                const parsed =
                                  val === ''
                                    ? null
                                    : parseUserInputToRaw(val, variable);
                                onUpdateCell(row.id, variable.name, parsed);
                                setEditingCell(null);
                              }}
                              onBlur={() => setEditingCell(null)}
                              className="w-full h-7 text-xs bg-white border border-blue-600 rounded px-1.5 text-slate-900 focus:outline-none"
                            >
                              <option value="">(System Missing)</option>
                              {Object.entries(variable.valueLabels).map(
                                ([code, lbl]) => (
                                  <option key={code} value={code}>
                                    {code} — {lbl}
                                  </option>
                                )
                              )}
                            </select>
                          ) : (
                            <input
                              ref={inputRef}
                              type="text"
                              value={draftValue}
                              onChange={(e) => setDraftValue(e.target.value)}
                              onBlur={() => commitTextEdit(row, variable)}
                              onKeyDown={(e) => {
                                if (e.key === 'Enter') {
                                  commitTextEdit(row, variable);
                                } else if (e.key === 'Escape') {
                                  setEditingCell(null);
                                }
                              }}
                              className="w-full h-7 text-xs font-mono tabular-nums bg-white border border-blue-600 rounded px-1.5 text-slate-900 focus:outline-none"
                            />
                          )
                        ) : (
                          <div className="flex items-center justify-between gap-1.5 min-w-0">
                            <span
                              className={`flex-1 truncate ${
                                isLabeled
                                  ? 'font-sans text-slate-900'
                                  : variable.type !== 'string'
                                    ? 'font-mono tabular-nums text-slate-800'
                                    : 'font-sans text-slate-800'
                              } ${
                                isSystemMissing(raw)
                                  ? 'text-slate-300 font-mono text-xs'
                                  : ''
                              }`}
                            >
                              {isSystemMissing(raw) ? '.' : text}
                            </span>

                            {issue && (
                              <span
                                className={`shrink-0 inline-flex items-center gap-0.5 text-[10px] font-semibold ${
                                  issue.kind === 'missing_user'
                                    ? 'text-amber-700'
                                    : 'text-red-700'
                                }`}
                              >
                                {issue.kind === 'missing_user' ? (
                                  <AlertTriangle className="w-3 h-3" />
                                ) : (
                                  <AlertCircle className="w-3 h-3" />
                                )}
                                <span className="sr-only">{issue.message}</span>
                              </span>
                            )}
                          </div>
                        )}
                      </td>
                    );
                  })}
                  <td className="px-1.5 text-center border-b border-slate-200">
                    <button
                      type="button"
                      onClick={() => onDeleteRow(row.id)}
                      aria-label={`Delete case ${originalIndex + 1}`}
                      title={`Delete case ${originalIndex + 1}`}
                      className="opacity-0 group-hover:opacity-100 focus:opacity-100 p-1 text-slate-400 hover:text-red-600 rounded transition-opacity cursor-pointer"
                    >
                      <Trash2 className="w-3.5 h-3.5" />
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>
    </div>
  );
}
