import { ClipboardPaste, Plus, Trash2, X } from 'lucide-react';
import { useState } from 'react';
import type { VariableDefinition } from '../types/spss';

interface ValueLabelsModalProps {
  variable: VariableDefinition;
  allVariables: VariableDefinition[];
  onSave: (variableId: string, nextLabels: Record<string, string>) => void;
  onClose: () => void;
}

const PRESET_SCALES: Array<{
  label: string;
  entries: Record<string, string>;
}> = [
  {
    label: '5-pt Likert (Agreement)',
    entries: {
      '1': 'Strongly Disagree',
      '2': 'Disagree',
      '3': 'Neutral',
      '4': 'Agree',
      '5': 'Strongly Agree',
      '-9': 'Skipped Question',
    },
  },
  {
    label: 'Yes / No (1 / 0)',
    entries: {
      '1': 'Yes',
      '0': 'No',
      '9': 'Unsure',
    },
  },
  {
    label: 'Frequency (1–5)',
    entries: {
      '1': 'Never',
      '2': 'Rarely',
      '3': 'Sometimes',
      '4': 'Often',
      '5': 'Always',
    },
  },
];

export function ValueLabelsModal({
  variable,
  allVariables,
  onSave,
  onClose,
}: ValueLabelsModalProps) {
  const [labels, setLabels] = useState<Record<string, string>>({
    ...variable.valueLabels,
  });
  const [valueInput, setValueInput] = useState('');
  const [labelInput, setLabelInput] = useState('');
  const [selectedKey, setSelectedKey] = useState<string | null>(null);
  const [showBulkPaste, setShowBulkPaste] = useState(false);
  const [bulkText, setBulkText] = useState('');

  const sortedEntries = Object.entries(labels).sort(([a], [b]) => {
    const na = Number(a);
    const nb = Number(b);
    if (!Number.isNaN(na) && !Number.isNaN(nb)) return na - nb;
    return a.localeCompare(b);
  });

  const handleAddOrUpdate = () => {
    const val = valueInput.trim();
    const lbl = labelInput.trim();
    if (!val || !lbl) return;
    setLabels((prev) => ({ ...prev, [val]: lbl }));
    setValueInput('');
    setLabelInput('');
    setSelectedKey(null);
  };

  const handleSelectEntry = (key: string, lbl: string) => {
    setSelectedKey(key);
    setValueInput(key);
    setLabelInput(lbl);
  };

  const handleRemove = (key: string) => {
    setLabels((prev) => {
      const next = { ...prev };
      delete next[key];
      return next;
    });
    if (selectedKey === key) {
      setSelectedKey(null);
      setValueInput('');
      setLabelInput('');
    }
  };

  const handleApplyBulk = () => {
    const next = { ...labels };
    bulkText.split('\n').forEach((line) => {
      const match = line.match(/^\s*([^=:\t]+?)\s*[=:\t]\s*(.+?)\s*$/);
      if (match) {
        next[match[1].trim()] = match[2].trim();
      }
    });
    setLabels(next);
    setBulkText('');
    setShowBulkPaste(false);
  };

  const otherCodedVars = allVariables.filter(
    (v) => v.id !== variable.id && Object.keys(v.valueLabels).length > 0
  );

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4"
      role="dialog"
      aria-modal="true"
      aria-labelledby="value-labels-title"
    >
      <div className="bg-white border border-slate-300 rounded-lg shadow-xl w-full max-w-xl flex flex-col max-h-[90vh]">
        {/* Header */}
        <div className="px-5 py-3.5 border-b border-slate-200 flex items-center justify-between">
          <div>
            <h2
              id="value-labels-title"
              className="font-display text-sm font-semibold text-slate-900"
            >
              Value Labels — <span className="font-mono">{variable.name}</span>
            </h2>
            {variable.label && (
              <p className="text-xs text-slate-500 mt-0.5">{variable.label}</p>
            )}
          </div>
          <button
            type="button"
            onClick={onClose}
            aria-label="Close modal"
            className="p-1.5 text-slate-400 hover:text-slate-700 rounded cursor-pointer"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        {/* Body */}
        <div className="p-5 overflow-y-auto space-y-4">
          {/* Code + Label Entry */}
          <div className="grid grid-cols-1 sm:grid-cols-12 gap-2.5 items-end">
            <div className="sm:col-span-3">
              <label
                htmlFor="vl-code"
                className="block text-xs font-medium text-slate-700 mb-1"
              >
                Value
              </label>
              <input
                id="vl-code"
                type="text"
                value={valueInput}
                onChange={(e) => setValueInput(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') handleAddOrUpdate();
                }}
                placeholder="e.g. 1"
                className="w-full px-2.5 py-1.5 text-xs font-mono border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-600"
              />
            </div>
            <div className="sm:col-span-6">
              <label
                htmlFor="vl-label"
                className="block text-xs font-medium text-slate-700 mb-1"
              >
                Label
              </label>
              <input
                id="vl-label"
                type="text"
                value={labelInput}
                onChange={(e) => setLabelInput(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') handleAddOrUpdate();
                }}
                placeholder="e.g. Strongly Agree"
                className="w-full px-2.5 py-1.5 text-xs border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-600"
              />
            </div>
            <div className="sm:col-span-3">
              <button
                type="button"
                onClick={handleAddOrUpdate}
                disabled={!valueInput.trim() || !labelInput.trim()}
                className="w-full inline-flex items-center justify-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-white bg-blue-700 hover:bg-blue-800 disabled:opacity-40 rounded-md cursor-pointer"
              >
                <Plus className="w-3.5 h-3.5" />
                <span>{selectedKey && labels[valueInput.trim()] ? 'Update' : 'Add'}</span>
              </button>
            </div>
          </div>

          {/* Quick Presets & Copy */}
          <div className="flex flex-wrap items-center justify-between gap-2 pt-1 border-t border-slate-100">
            <div className="flex flex-wrap items-center gap-1.5">
              <span className="text-xs text-slate-500 mr-1">Presets:</span>
              {PRESET_SCALES.map((preset) => (
                <button
                  key={preset.label}
                  type="button"
                  onClick={() => setLabels({ ...preset.entries })}
                  className="px-2 py-1 text-[11px] font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 rounded border border-slate-200 cursor-pointer"
                >
                  {preset.label}
                </button>
              ))}
            </div>

            <div className="flex items-center gap-2">
              {otherCodedVars.length > 0 && (
                <select
                  value=""
                  onChange={(e) => {
                    const source = otherCodedVars.find(
                      (v) => v.id === e.target.value
                    );
                    if (source) setLabels({ ...source.valueLabels });
                  }}
                  aria-label="Copy value labels from another variable"
                  className="px-2 py-1 text-[11px] bg-white border border-slate-300 rounded text-slate-700 cursor-pointer"
                >
                  <option value="" disabled>
                    Copy from variable...
                  </option>
                  {otherCodedVars.map((v) => (
                    <option key={v.id} value={v.id}>
                      {v.name} ({Object.keys(v.valueLabels).length})
                    </option>
                  ))}
                </select>
              )}

              <button
                type="button"
                onClick={() => setShowBulkPaste((s) => !s)}
                className="inline-flex items-center gap-1 px-2 py-1 text-[11px] font-medium text-slate-700 bg-white border border-slate-300 rounded hover:bg-slate-50 cursor-pointer"
              >
                <ClipboardPaste className="w-3 h-3" />
                <span>Bulk Paste</span>
              </button>
            </div>
          </div>

          {showBulkPaste && (
            <div className="p-3 bg-slate-50 border border-slate-200 rounded-md space-y-2">
              <label
                htmlFor="bulk-vl"
                className="block text-xs font-medium text-slate-700"
              >
                Paste one mapping per line (e.g. <code className="font-mono">1 = Low</code>)
              </label>
              <textarea
                id="bulk-vl"
                rows={4}
                value={bulkText}
                onChange={(e) => setBulkText(e.target.value)}
                placeholder={'1 = Strongly Disagree\n2 = Disagree\n3 = Neutral'}
                className="w-full p-2 text-xs font-mono bg-white border border-slate-300 rounded focus:outline-none focus:ring-2 focus:ring-blue-600"
              />
              <div className="flex justify-end gap-2">
                <button
                  type="button"
                  onClick={() => setShowBulkPaste(false)}
                  className="px-2.5 py-1 text-xs text-slate-600 hover:text-slate-900 cursor-pointer"
                >
                  Cancel
                </button>
                <button
                  type="button"
                  onClick={handleApplyBulk}
                  className="px-3 py-1 text-xs font-semibold text-white bg-slate-900 hover:bg-slate-800 rounded cursor-pointer"
                >
                  Apply Lines
                </button>
              </div>
            </div>
          )}

          {/* Existing Mappings Table */}
          <div className="border border-slate-200 rounded-md overflow-hidden">
            <div className="bg-slate-50 border-b border-slate-200 px-3 py-2 flex items-center justify-between text-xs font-semibold text-slate-600">
              <span>Defined Mappings ({sortedEntries.length})</span>
              {sortedEntries.length > 0 && (
                <button
                  type="button"
                  onClick={() => setLabels({})}
                  className="text-[11px] font-medium text-red-600 hover:text-red-700 cursor-pointer"
                >
                  Clear All
                </button>
              )}
            </div>
            {sortedEntries.length === 0 ? (
              <div className="p-6 text-center text-xs text-slate-500">
                No value labels defined yet. Enter a code and label above or choose a preset.
              </div>
            ) : (
              <div className="max-h-56 overflow-y-auto divide-y divide-slate-200">
                {sortedEntries.map(([code, lbl]) => (
                  <div
                    key={code}
                    onClick={() => handleSelectEntry(code, lbl)}
                    className={`px-3 py-2 flex items-center justify-between gap-3 text-xs cursor-pointer transition-colors ${
                      selectedKey === code
                        ? 'bg-blue-50 text-blue-900'
                        : 'hover:bg-slate-50 text-slate-800'
                    }`}
                  >
                    <div className="flex items-center gap-3 min-w-0">
                      <span className="font-mono font-semibold tabular-nums text-slate-700 w-14 shrink-0">
                        {code} =
                      </span>
                      <span className="truncate">"{lbl}"</span>
                    </div>
                    <button
                      type="button"
                      onClick={(e) => {
                        e.stopPropagation();
                        handleRemove(code);
                      }}
                      aria-label={`Remove ${code}`}
                      className="p-1 text-slate-400 hover:text-red-600 rounded cursor-pointer shrink-0"
                    >
                      <Trash2 className="w-3.5 h-3.5" />
                    </button>
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>

        {/* Footer */}
        <div className="px-5 py-3 bg-slate-50 border-t border-slate-200 flex items-center justify-end gap-2">
          <button
            type="button"
            onClick={onClose}
            className="px-3.5 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-300 rounded-md hover:bg-slate-50 cursor-pointer"
          >
            Cancel
          </button>
          <button
            type="button"
            onClick={() => {
              onSave(variable.id, labels);
              onClose();
            }}
            className="px-4 py-1.5 text-xs font-semibold text-white bg-blue-700 hover:bg-blue-800 rounded-md cursor-pointer"
          >
            Save Value Labels
          </button>
        </div>
      </div>
    </div>
  );
}
