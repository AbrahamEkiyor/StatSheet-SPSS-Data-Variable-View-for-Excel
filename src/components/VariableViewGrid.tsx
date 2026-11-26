import {
  ArrowDown,
  ArrowUp,
  BarChart3,
  Copy,
  MoreHorizontal,
  Plus,
  Trash2,
} from 'lucide-react';
import { useState } from 'react';
import {
  formatClassBinsSummary,
  formatMissingSummary,
  formatValueLabelsSummary,
  isNumericType,
  validateVariableName,
} from '../engine/spssEngine';
import type {
  VariableAlign,
  VariableDefinition,
  VariableMeasure,
  VariableRole,
  VariableType,
} from '../types/spss';
import { MeasureIcon } from './MeasureIcon';

interface VariableViewGridProps {
  variables: VariableDefinition[];
  searchQuery: string;
  selectedVariableId: string | null;
  onSelectVariable: (variableId: string) => void;
  onUpdateVariable: (
    variableId: string,
    patch: Partial<VariableDefinition>,
    renameFrom?: string
  ) => void;
  onAddVariable: () => void;
  onDuplicateVariable: (variableId: string) => void;
  onMoveVariable: (variableId: string, direction: 'up' | 'down') => void;
  onDeleteVariable: (variableId: string) => void;
  onOpenValueLabels: (variableId: string) => void;
  onOpenMissingModal: (variableId: string) => void;
  onOpenClassBins: (variableId: string) => void;
  onOpenStats: (variableId: string) => void;
}

const VARIABLE_TYPES: Array<{ value: VariableType; label: string }> = [
  { value: 'numeric', label: 'Numeric' },
  { value: 'comma', label: 'Comma' },
  { value: 'dot', label: 'Dot' },
  { value: 'scientific', label: 'Scientific' },
  { value: 'date', label: 'Date' },
  { value: 'dollar', label: 'Dollar' },
  { value: 'custom_currency', label: 'Custom Currency' },
  { value: 'string', label: 'String' },
];

const ALIGN_OPTIONS: Array<{ value: VariableAlign; label: string }> = [
  { value: 'left', label: 'Left' },
  { value: 'center', label: 'Center' },
  { value: 'right', label: 'Right' },
];

const MEASURE_OPTIONS: Array<{ value: VariableMeasure; label: string }> = [
  { value: 'scale', label: 'Scale' },
  { value: 'ordinal', label: 'Ordinal' },
  { value: 'nominal', label: 'Nominal' },
];

const ROLE_OPTIONS: Array<{ value: VariableRole; label: string }> = [
  { value: 'input', label: 'Input' },
  { value: 'target', label: 'Target' },
  { value: 'both', label: 'Both' },
  { value: 'none', label: 'None' },
  { value: 'partition', label: 'Partition' },
  { value: 'split', label: 'Split' },
];

export function VariableViewGrid({
  variables,
  searchQuery,
  selectedVariableId,
  onSelectVariable,
  onUpdateVariable,
  onAddVariable,
  onDuplicateVariable,
  onMoveVariable,
  onDeleteVariable,
  onOpenValueLabels,
  onOpenMissingModal,
  onOpenClassBins,
  onOpenStats,
}: VariableViewGridProps) {
  const [nameDrafts, setNameDrafts] = useState<Record<string, string>>({});
  const [nameErrors, setNameErrors] = useState<Record<string, string | null>>({});

  const query = searchQuery.trim().toLowerCase();
  const filteredVariables = variables
    .map((v, index) => ({ variable: v, index }))
    .filter(({ variable }) => {
      if (!query) return true;
      return (
        variable.name.toLowerCase().includes(query) ||
        variable.label.toLowerCase().includes(query)
      );
    });

  const handleNameCommit = (variable: VariableDefinition) => {
    const draft = nameDrafts[variable.id];
    if (draft === undefined) return;
    const trimmed = draft.trim();
    const err = validateVariableName(
      trimmed,
      variables.map((v) => v.name),
      variable.name
    );
    if (err) {
      setNameErrors((prev) => ({ ...prev, [variable.id]: err }));
      return;
    }
    setNameErrors((prev) => ({ ...prev, [variable.id]: null }));
    if (trimmed !== variable.name) {
      onUpdateVariable(variable.id, { name: trimmed }, variable.name);
    }
    setNameDrafts((prev) => {
      const next = { ...prev };
      delete next[variable.id];
      return next;
    });
  };

  return (
    <div className="flex-1 flex flex-col min-h-0 bg-white">
      {/* Subtle instructional bar */}
      <div className="px-4 py-2 bg-slate-50 border-b border-slate-200 flex flex-wrap items-center justify-between gap-3 text-xs text-slate-600 shrink-0">
        <span>
          Define the 11 SPSS variable properties below. Click <strong className="font-semibold text-slate-800">Values</strong> to map numeric codes to labels, or <strong className="font-semibold text-slate-800">Missing</strong> to configure missing codes &amp; validation bounds.
        </span>
        <button
          type="button"
          onClick={onAddVariable}
          className="inline-flex items-center gap-1.5 px-2.5 py-1 text-xs font-semibold text-blue-700 hover:bg-blue-50 rounded border border-blue-200 cursor-pointer whitespace-nowrap"
        >
          <Plus className="w-3.5 h-3.5" />
          <span>New Variable</span>
        </button>
      </div>

      {/* 11-Column SPSS Variable View Grid */}
      <div className="flex-1 overflow-auto grid-scroll">
        {filteredVariables.length === 0 ? (
          <div className="flex flex-col items-center justify-center py-16 px-6 text-center">
            <p className="text-sm font-medium text-slate-800">
              {variables.length === 0
                ? 'No variables defined.'
                : 'No variables match your search.'}
            </p>
            <button
              type="button"
              onClick={onAddVariable}
              className="mt-4 inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-white bg-blue-700 hover:bg-blue-800 rounded-md cursor-pointer"
            >
              <Plus className="w-3.5 h-3.5" />
              <span>Add Variable</span>
            </button>
          </div>
        ) : (
          <table className="border-collapse w-max min-w-full text-[13px]">
            <thead className="sticky top-0 z-20 bg-slate-100 border-b border-slate-300 text-xs font-semibold text-slate-700">
              <tr>
                <th className="sticky left-0 z-30 w-11 min-w-11 bg-slate-100 border-r border-slate-300 px-2 py-2.5 text-center font-mono">
                  #
                </th>
                <th className="w-36 min-w-36 border-r border-slate-300 px-2.5 py-2.5 text-left">
                  Name
                </th>
                <th className="w-32 min-w-32 border-r border-slate-300 px-2.5 py-2.5 text-left">
                  Type
                </th>
                <th className="w-20 min-w-20 border-r border-slate-300 px-2.5 py-2.5 text-right">
                  Width
                </th>
                <th className="w-22 min-w-22 border-r border-slate-300 px-2.5 py-2.5 text-right">
                  Decimals
                </th>
                <th className="w-56 min-w-56 border-r border-slate-300 px-2.5 py-2.5 text-left">
                  Label
                </th>
                <th className="w-48 min-w-48 border-r border-slate-300 px-2.5 py-2.5 text-left">
                  Values
                </th>
                <th className="w-40 min-w-40 border-r border-slate-300 px-2.5 py-2.5 text-left">
                  Missing &amp; Rules
                </th>
                <th className="w-36 min-w-36 border-r border-slate-300 px-2.5 py-2.5 text-left">
                  Class Bins
                </th>
                <th className="w-22 min-w-22 border-r border-slate-300 px-2.5 py-2.5 text-right">
                  Columns
                </th>
                <th className="w-26 min-w-26 border-r border-slate-300 px-2.5 py-2.5 text-left">
                  Align
                </th>
                <th className="w-32 min-w-32 border-r border-slate-300 px-2.5 py-2.5 text-left">
                  Measure
                </th>
                <th className="w-28 min-w-28 border-r border-slate-300 px-2.5 py-2.5 text-left">
                  Role
                </th>
                <th className="w-36 min-w-36 px-2.5 py-2.5 text-center">
                  Actions
                </th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200">
              {filteredVariables.map(({ variable, index }) => {
                const isSelected = selectedVariableId === variable.id;
                const nameValue =
                  nameDrafts[variable.id] !== undefined
                    ? nameDrafts[variable.id]
                    : variable.name;
                const nameError = nameErrors[variable.id];
                const hasValidationBounds =
                  variable.validation?.min !== undefined ||
                  variable.validation?.max !== undefined ||
                  variable.validation?.restrictToLabels ||
                  variable.validation?.required;

                return (
                  <tr
                    key={variable.id}
                    onClick={() => onSelectVariable(variable.id)}
                    className={`h-10 transition-colors ${
                      isSelected ? 'bg-blue-50/50' : 'hover:bg-slate-50'
                    }`}
                  >
                    <td className="sticky left-0 z-10 bg-slate-50 border-r border-slate-300 px-2 text-center font-mono text-xs text-slate-500 tabular-nums">
                      {index + 1}
                    </td>

                    {/* 1. Name */}
                    <td className="border-r border-slate-200 px-2 py-1">
                      <input
                        type="text"
                        value={nameValue}
                        aria-label={`Variable ${index + 1} name`}
                        onChange={(e) => {
                          const val = e.target.value;
                          setNameDrafts((prev) => ({
                            ...prev,
                            [variable.id]: val,
                          }));
                          setNameErrors((prev) => ({
                            ...prev,
                            [variable.id]: null,
                          }));
                        }}
                        onBlur={() => handleNameCommit(variable)}
                        onKeyDown={(e) => {
                          if (e.key === 'Enter') handleNameCommit(variable);
                        }}
                        className={`w-full px-1.5 py-1 text-xs font-mono rounded border bg-white focus:outline-none focus:ring-2 focus:ring-blue-600 ${
                          nameError
                            ? 'border-red-500 text-red-700'
                            : 'border-transparent hover:border-slate-300 text-slate-900'
                        }`}
                      />
                      {nameError && (
                        <div className="text-[10px] text-red-700 mt-0.5 leading-tight">
                          {nameError}
                        </div>
                      )}
                    </td>

                    {/* 2. Type */}
                    <td className="border-r border-slate-200 px-2 py-1">
                      <select
                        value={variable.type}
                        aria-label={`${variable.name} type`}
                        onChange={(e) => {
                          const nextType = e.target.value as VariableType;
                          onUpdateVariable(variable.id, {
                            type: nextType,
                            decimals:
                              nextType === 'string' || nextType === 'date'
                                ? 0
                                : variable.decimals,
                            align: nextType === 'string' ? 'left' : variable.align,
                          });
                        }}
                        className="w-full px-1.5 py-1 text-xs bg-white border border-transparent hover:border-slate-300 rounded text-slate-800 focus:outline-none focus:ring-2 focus:ring-blue-600 cursor-pointer"
                      >
                        {VARIABLE_TYPES.map((t) => (
                          <option key={t.value} value={t.value}>
                            {t.label}
                          </option>
                        ))}
                      </select>
                    </td>

                    {/* 3. Width */}
                    <td className="border-r border-slate-200 px-2 py-1">
                      <input
                        type="number"
                        min={1}
                        max={256}
                        value={variable.width}
                        aria-label={`${variable.name} width`}
                        onChange={(e) =>
                          onUpdateVariable(variable.id, {
                            width: Math.max(1, Math.min(256, Number(e.target.value) || 1)),
                          })
                        }
                        className="w-full px-1.5 py-1 text-xs font-mono tabular-nums text-right bg-white border border-transparent hover:border-slate-300 rounded text-slate-800 focus:outline-none focus:ring-2 focus:ring-blue-600"
                      />
                    </td>

                    {/* 4. Decimals */}
                    <td className="border-r border-slate-200 px-2 py-1">
                      <input
                        type="number"
                        min={0}
                        max={10}
                        disabled={variable.type === 'string' || variable.type === 'date'}
                        value={variable.decimals}
                        aria-label={`${variable.name} decimals`}
                        onChange={(e) =>
                          onUpdateVariable(variable.id, {
                            decimals: Math.max(0, Math.min(10, Number(e.target.value) || 0)),
                          })
                        }
                        className="w-full px-1.5 py-1 text-xs font-mono tabular-nums text-right bg-white border border-transparent hover:border-slate-300 disabled:bg-slate-100 disabled:text-slate-400 rounded text-slate-800 focus:outline-none focus:ring-2 focus:ring-blue-600"
                      />
                    </td>

                    {/* 5. Label */}
                    <td className="border-r border-slate-200 px-2 py-1">
                      <input
                        type="text"
                        value={variable.label}
                        placeholder="Add variable description..."
                        aria-label={`${variable.name} label`}
                        onChange={(e) =>
                          onUpdateVariable(variable.id, { label: e.target.value })
                        }
                        className="w-full px-1.5 py-1 text-xs bg-white border border-transparent hover:border-slate-300 rounded text-slate-800 placeholder:text-slate-400 focus:outline-none focus:ring-2 focus:ring-blue-600"
                      />
                    </td>

                    {/* 6. Values */}
                    <td className="border-r border-slate-200 px-2 py-1">
                      <button
                        type="button"
                        onClick={() => onOpenValueLabels(variable.id)}
                        className="w-full flex items-center justify-between gap-2 px-2 py-1 text-xs text-left rounded border border-transparent hover:border-slate-300 hover:bg-white transition-colors cursor-pointer"
                        title="Edit Value Labels"
                      >
                        <span
                          className={`truncate ${
                            Object.keys(variable.valueLabels).length > 0
                              ? 'text-blue-700 font-medium'
                              : 'text-slate-400'
                          }`}
                        >
                          {formatValueLabelsSummary(variable.valueLabels)}
                        </span>
                        <span className="inline-flex items-center justify-center w-5 h-5 bg-slate-100 border border-slate-300 rounded text-slate-600 shrink-0">
                          <MoreHorizontal className="w-3 h-3" />
                        </span>
                      </button>
                    </td>

                    {/* 7. Missing & Validation */}
                    <td className="border-r border-slate-200 px-2 py-1">
                      <button
                        type="button"
                        onClick={() => onOpenMissingModal(variable.id)}
                        className="w-full flex items-center justify-between gap-2 px-2 py-1 text-xs text-left rounded border border-transparent hover:border-slate-300 hover:bg-white transition-colors cursor-pointer"
                        title="Configure Missing Values & Data Validation Rules"
                      >
                        <span className="truncate font-mono tabular-nums text-slate-700">
                          {formatMissingSummary(variable.missingValues)}
                          {hasValidationBounds ? ' · Rules' : ''}
                        </span>
                        <span className="inline-flex items-center justify-center w-5 h-5 bg-slate-100 border border-slate-300 rounded text-slate-600 shrink-0">
                          <MoreHorizontal className="w-3 h-3" />
                        </span>
                      </button>
                    </td>

                    {/* Class Boundaries */}
                    <td className="border-r border-slate-200 px-2 py-1">
                      {isNumericType(variable.type) ? (
                        <button
                          type="button"
                          onClick={() => onOpenClassBins(variable.id)}
                          className="w-full flex items-center justify-between gap-2 px-2 py-1 text-xs text-left rounded border border-transparent hover:border-slate-300 hover:bg-white transition-colors cursor-pointer"
                          title="Define Class Boundaries for Grouped Frequency Distribution"
                        >
                          <span
                            className={`truncate font-mono tabular-nums ${
                              variable.classBins?.enabled
                                ? 'text-blue-700 font-medium'
                                : 'text-slate-400'
                            }`}
                          >
                            {formatClassBinsSummary(variable.classBins)}
                          </span>
                          <span className="inline-flex items-center justify-center w-5 h-5 bg-slate-100 border border-slate-300 rounded text-slate-600 shrink-0">
                            <MoreHorizontal className="w-3 h-3" />
                          </span>
                        </button>
                      ) : (
                        <span className="px-2 py-1 text-xs text-slate-400 block">
                          N/A
                        </span>
                      )}
                    </td>

                    {/* 8. Columns (px display width) */}
                    <td className="border-r border-slate-200 px-2 py-1">
                      <input
                        type="number"
                        min={80}
                        max={360}
                        step={8}
                        value={variable.columns}
                        aria-label={`${variable.name} column display width`}
                        onChange={(e) =>
                          onUpdateVariable(variable.id, {
                            columns: Math.max(80, Math.min(360, Number(e.target.value) || 120)),
                          })
                        }
                        className="w-full px-1.5 py-1 text-xs font-mono tabular-nums text-right bg-white border border-transparent hover:border-slate-300 rounded text-slate-800 focus:outline-none focus:ring-2 focus:ring-blue-600"
                      />
                    </td>

                    {/* 9. Align */}
                    <td className="border-r border-slate-200 px-2 py-1">
                      <select
                        value={variable.align}
                        aria-label={`${variable.name} alignment`}
                        onChange={(e) =>
                          onUpdateVariable(variable.id, {
                            align: e.target.value as VariableAlign,
                          })
                        }
                        className="w-full px-1.5 py-1 text-xs bg-white border border-transparent hover:border-slate-300 rounded text-slate-800 focus:outline-none focus:ring-2 focus:ring-blue-600 cursor-pointer"
                      >
                        {ALIGN_OPTIONS.map((a) => (
                          <option key={a.value} value={a.value}>
                            {a.label}
                          </option>
                        ))}
                      </select>
                    </td>

                    {/* 10. Measure */}
                    <td className="border-r border-slate-200 px-2 py-1">
                      <div className="flex items-center gap-1.5">
                        <MeasureIcon measure={variable.measure} className="w-3.5 h-3.5" />
                        <select
                          value={variable.measure}
                          aria-label={`${variable.name} measure`}
                          onChange={(e) =>
                            onUpdateVariable(variable.id, {
                              measure: e.target.value as VariableMeasure,
                            })
                          }
                          className="w-full px-1 py-1 text-xs bg-white border border-transparent hover:border-slate-300 rounded text-slate-800 focus:outline-none focus:ring-2 focus:ring-blue-600 cursor-pointer"
                        >
                          {MEASURE_OPTIONS.map((m) => (
                            <option key={m.value} value={m.value}>
                              {m.label}
                            </option>
                          ))}
                        </select>
                      </div>
                    </td>

                    {/* 11. Role */}
                    <td className="border-r border-slate-200 px-2 py-1">
                      <select
                        value={variable.role}
                        aria-label={`${variable.name} role`}
                        onChange={(e) =>
                          onUpdateVariable(variable.id, {
                            role: e.target.value as VariableRole,
                          })
                        }
                        className="w-full px-1.5 py-1 text-xs bg-white border border-transparent hover:border-slate-300 rounded text-slate-800 focus:outline-none focus:ring-2 focus:ring-blue-600 cursor-pointer"
                      >
                        {ROLE_OPTIONS.map((r) => (
                          <option key={r.value} value={r.value}>
                            {r.label}
                          </option>
                        ))}
                      </select>
                    </td>

                    {/* Row Actions */}
                    <td className="px-2 py-1">
                      <div className="flex items-center justify-center gap-1">
                        <button
                          type="button"
                          onClick={() => onOpenStats(variable.id)}
                          title={`Descriptives & Frequencies for ${variable.name}`}
                          className="p-1 text-slate-500 hover:text-blue-700 hover:bg-blue-50 rounded cursor-pointer"
                        >
                          <BarChart3 className="w-3.5 h-3.5" />
                        </button>
                        <button
                          type="button"
                          onClick={() => onMoveVariable(variable.id, 'up')}
                          disabled={index === 0}
                          title="Move variable up"
                          className="p-1 text-slate-500 hover:text-slate-900 disabled:opacity-30 rounded cursor-pointer"
                        >
                          <ArrowUp className="w-3.5 h-3.5" />
                        </button>
                        <button
                          type="button"
                          onClick={() => onMoveVariable(variable.id, 'down')}
                          disabled={index === variables.length - 1}
                          title="Move variable down"
                          className="p-1 text-slate-500 hover:text-slate-900 disabled:opacity-30 rounded cursor-pointer"
                        >
                          <ArrowDown className="w-3.5 h-3.5" />
                        </button>
                        <button
                          type="button"
                          onClick={() => onDuplicateVariable(variable.id)}
                          title="Duplicate variable"
                          className="p-1 text-slate-500 hover:text-slate-900 rounded cursor-pointer"
                        >
                          <Copy className="w-3.5 h-3.5" />
                        </button>
                        <button
                          type="button"
                          onClick={() => onDeleteVariable(variable.id)}
                          disabled={variables.length <= 1}
                          title="Delete variable"
                          className="p-1 text-slate-500 hover:text-red-600 disabled:opacity-30 rounded cursor-pointer"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                        </button>
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        )}
      </div>
    </div>
  );
}
