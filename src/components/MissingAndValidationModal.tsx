import { X } from 'lucide-react';
import { useState } from 'react';
import { isNumericType } from '../engine/spssEngine';
import type {
  MissingValueMode,
  MissingValueSpec,
  ValidationRuleSpec,
  VariableDefinition,
} from '../types/spss';

interface MissingAndValidationModalProps {
  variable: VariableDefinition;
  onSave: (
    variableId: string,
    missingValues: MissingValueSpec,
    validation: ValidationRuleSpec
  ) => void;
  onClose: () => void;
}

export function MissingAndValidationModal({
  variable,
  onSave,
  onClose,
}: MissingAndValidationModalProps) {
  const [mode, setMode] = useState<MissingValueMode>(variable.missingValues.mode);
  const [d1, setD1] = useState<string>(
    variable.missingValues.discrete[0] !== undefined
      ? String(variable.missingValues.discrete[0])
      : ''
  );
  const [d2, setD2] = useState<string>(
    variable.missingValues.discrete[1] !== undefined
      ? String(variable.missingValues.discrete[1])
      : ''
  );
  const [d3, setD3] = useState<string>(
    variable.missingValues.discrete[2] !== undefined
      ? String(variable.missingValues.discrete[2])
      : ''
  );
  const [rangeLow, setRangeLow] = useState<string>(
    variable.missingValues.range?.low !== undefined
      ? String(variable.missingValues.range.low)
      : ''
  );
  const [rangeHigh, setRangeHigh] = useState<string>(
    variable.missingValues.range?.high !== undefined
      ? String(variable.missingValues.range.high)
      : ''
  );
  const [rangeDiscrete, setRangeDiscrete] = useState<string>(
    mode === 'range_discrete' && variable.missingValues.discrete[0] !== undefined
      ? String(variable.missingValues.discrete[0])
      : ''
  );

  const [minVal, setMinVal] = useState<string>(
    variable.validation?.min !== undefined ? String(variable.validation.min) : ''
  );
  const [maxVal, setMaxVal] = useState<string>(
    variable.validation?.max !== undefined ? String(variable.validation.max) : ''
  );
  const [restrictToLabels, setRestrictToLabels] = useState<boolean>(
    Boolean(variable.validation?.restrictToLabels)
  );
  const [required, setRequired] = useState<boolean>(
    Boolean(variable.validation?.required)
  );

  const numeric = isNumericType(variable.type);

  const parseCode = (val: string): string | number | null => {
    const trimmed = val.trim();
    if (!trimmed) return null;
    if (numeric && !Number.isNaN(Number(trimmed))) {
      return Number(trimmed);
    }
    return trimmed;
  };

  const handleSave = () => {
    let nextMissing: MissingValueSpec = { mode: 'none', discrete: [] };

    if (mode === 'discrete') {
      const discrete = [d1, d2, d3]
        .map(parseCode)
        .filter((x): x is string | number => x !== null);
      nextMissing = {
        mode: discrete.length > 0 ? 'discrete' : 'none',
        discrete,
      };
    } else if (mode === 'range_discrete') {
      const low = Number(rangeLow);
      const high = Number(rangeHigh);
      const disc = parseCode(rangeDiscrete);
      if (!Number.isNaN(low) && !Number.isNaN(high)) {
        nextMissing = {
          mode: 'range_discrete',
          range: { low: Math.min(low, high), high: Math.max(low, high) },
          discrete: disc !== null ? [disc] : [],
        };
      }
    }

    const nextValidation: ValidationRuleSpec = {
      min: minVal.trim() !== '' && !Number.isNaN(Number(minVal)) ? Number(minVal) : undefined,
      max: maxVal.trim() !== '' && !Number.isNaN(Number(maxVal)) ? Number(maxVal) : undefined,
      restrictToLabels,
      required,
    };

    onSave(variable.id, nextMissing, nextValidation);
    onClose();
  };

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4"
      role="dialog"
      aria-modal="true"
      aria-labelledby="missing-modal-title"
    >
      <div className="bg-white border border-slate-300 rounded-lg shadow-xl w-full max-w-lg flex flex-col max-h-[90vh]">
        <div className="px-5 py-3.5 border-b border-slate-200 flex items-center justify-between">
          <div>
            <h2
              id="missing-modal-title"
              className="font-display text-sm font-semibold text-slate-900"
            >
              Missing Values &amp; Validation —{' '}
              <span className="font-mono">{variable.name}</span>
            </h2>
            <p className="text-xs text-slate-500 mt-0.5">
              User-missing codes are excluded from valid percentages and scale statistics.
            </p>
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

        <div className="p-5 overflow-y-auto space-y-5 text-xs">
          {/* Missing Values Section */}
          <fieldset className="space-y-3">
            <legend className="font-display text-xs font-semibold text-slate-900 mb-1">
              SPSS Missing Values
            </legend>

            <label className="flex items-center gap-2 cursor-pointer">
              <input
                type="radio"
                name="missingMode"
                checked={mode === 'none'}
                onChange={() => setMode('none')}
                className="text-blue-700 focus:ring-blue-600"
              />
              <span className="text-slate-800">No missing values</span>
            </label>

            <div className="space-y-2">
              <label className="flex items-center gap-2 cursor-pointer">
                <input
                  type="radio"
                  name="missingMode"
                  checked={mode === 'discrete'}
                  onChange={() => setMode('discrete')}
                  className="text-blue-700 focus:ring-blue-600"
                />
                <span className="text-slate-800">
                  Discrete missing values (up to 3 codes)
                </span>
              </label>
              <div className="pl-6 grid grid-cols-3 gap-2">
                {[
                  { val: d1, set: setD1, placeholder: 'e.g. -9' },
                  { val: d2, set: setD2, placeholder: 'e.g. 99' },
                  { val: d3, set: setD3, placeholder: 'e.g. 999' },
                ].map((item, idx) => (
                  <input
                    key={idx}
                    type="text"
                    disabled={mode !== 'discrete'}
                    value={item.val}
                    onChange={(e) => item.set(e.target.value)}
                    placeholder={item.placeholder}
                    className="px-2.5 py-1.5 font-mono border border-slate-300 rounded disabled:bg-slate-100 disabled:text-slate-400 focus:outline-none focus:ring-2 focus:ring-blue-600"
                  />
                ))}
              </div>
            </div>

            {numeric && (
              <div className="space-y-2">
                <label className="flex items-center gap-2 cursor-pointer">
                  <input
                    type="radio"
                    name="missingMode"
                    checked={mode === 'range_discrete'}
                    onChange={() => setMode('range_discrete')}
                    className="text-blue-700 focus:ring-blue-600"
                  />
                  <span className="text-slate-800">
                    Range plus one optional discrete missing value
                  </span>
                </label>
                <div className="pl-6 grid grid-cols-3 gap-2">
                  <div>
                    <label className="block text-[11px] text-slate-500 mb-1">
                      Low
                    </label>
                    <input
                      type="number"
                      disabled={mode !== 'range_discrete'}
                      value={rangeLow}
                      onChange={(e) => setRangeLow(e.target.value)}
                      placeholder="-99"
                      className="w-full px-2.5 py-1.5 font-mono border border-slate-300 rounded disabled:bg-slate-100"
                    />
                  </div>
                  <div>
                    <label className="block text-[11px] text-slate-500 mb-1">
                      High
                    </label>
                    <input
                      type="number"
                      disabled={mode !== 'range_discrete'}
                      value={rangeHigh}
                      onChange={(e) => setRangeHigh(e.target.value)}
                      placeholder="-1"
                      className="w-full px-2.5 py-1.5 font-mono border border-slate-300 rounded disabled:bg-slate-100"
                    />
                  </div>
                  <div>
                    <label className="block text-[11px] text-slate-500 mb-1">
                      Discrete (opt.)
                    </label>
                    <input
                      type="text"
                      disabled={mode !== 'range_discrete'}
                      value={rangeDiscrete}
                      onChange={(e) => setRangeDiscrete(e.target.value)}
                      placeholder="999"
                      className="w-full px-2.5 py-1.5 font-mono border border-slate-300 rounded disabled:bg-slate-100"
                    />
                  </div>
                </div>
              </div>
            )}
          </fieldset>

          {/* Validation Rules Section */}
          <div className="pt-4 border-t border-slate-200 space-y-3">
            <h3 className="font-display text-xs font-semibold text-slate-900">
              Data Validation Rules
            </h3>

            {numeric && (
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label
                    htmlFor="val-min"
                    className="block text-slate-600 mb-1"
                  >
                    Minimum Valid Value
                  </label>
                  <input
                    id="val-min"
                    type="number"
                    value={minVal}
                    onChange={(e) => setMinVal(e.target.value)}
                    placeholder="No min"
                    className="w-full px-2.5 py-1.5 font-mono border border-slate-300 rounded focus:outline-none focus:ring-2 focus:ring-blue-600"
                  />
                </div>
                <div>
                  <label
                    htmlFor="val-max"
                    className="block text-slate-600 mb-1"
                  >
                    Maximum Valid Value
                  </label>
                  <input
                    id="val-max"
                    type="number"
                    value={maxVal}
                    onChange={(e) => setMaxVal(e.target.value)}
                    placeholder="No max"
                    className="w-full px-2.5 py-1.5 font-mono border border-slate-300 rounded focus:outline-none focus:ring-2 focus:ring-blue-600"
                  />
                </div>
              </div>
            )}

            <label className="flex items-center gap-2 cursor-pointer">
              <input
                type="checkbox"
                checked={restrictToLabels}
                onChange={(e) => setRestrictToLabels(e.target.checked)}
                className="rounded border-slate-300 text-blue-700 focus:ring-blue-600"
              />
              <span className="text-slate-800">
                Restrict entries to defined Value Labels (also creates Excel dropdowns on sync)
              </span>
            </label>

            <label className="flex items-center gap-2 cursor-pointer">
              <input
                type="checkbox"
                checked={required}
                onChange={(e) => setRequired(e.target.checked)}
                className="rounded border-slate-300 text-blue-700 focus:ring-blue-600"
              />
              <span className="text-slate-800">
                Flag blank (system-missing) cells as required
              </span>
            </label>
          </div>
        </div>

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
            onClick={handleSave}
            className="px-4 py-1.5 text-xs font-semibold text-white bg-blue-700 hover:bg-blue-800 rounded-md cursor-pointer"
          >
            Apply Rules
          </button>
        </div>
      </div>
    </div>
  );
}
