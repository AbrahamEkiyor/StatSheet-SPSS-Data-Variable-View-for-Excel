import { AlertCircle, ArrowRightLeft, BookmarkCheck, Plus, RefreshCw, Trash2, X } from 'lucide-react';
import { useMemo, useState } from 'react';
import {
  buildIntervalLabel,
  classifyValueIntoInterval,
  computeVariableStats,
  generateEqualWidthIntervals,
  getValidNumericValues,
  sanitizeVariableName,
  validateClassIntervals,
  validateVariableName,
} from '../engine/spssEngine';
import type {
  ClassBinMode,
  ClassBinSpec,
  ClassInterval,
  DataRow,
  EndpointInclusionRule,
  VariableDefinition,
} from '../types/spss';

const BINNING_TEMPLATE_STORAGE_KEY = 'statsheet_binning_template_v1';

interface SavedBinningTemplate {
  mode: ClassBinMode;
  endpointRule: EndpointInclusionRule;
  binCount?: number;
  binWidth?: number;
  startAt?: number;
  niceRound: boolean;
  intervals: ClassInterval[];
  sourceVariableName?: string;
  savedAt: string;
}

function loadBinningTemplate(): SavedBinningTemplate | null {
  if (typeof window === 'undefined') return null;
  try {
    const raw = window.localStorage.getItem(BINNING_TEMPLATE_STORAGE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw) as Partial<SavedBinningTemplate>;
    if (!parsed || !parsed.mode || !parsed.endpointRule || !Array.isArray(parsed.intervals)) {
      return null;
    }
    return {
      mode: parsed.mode,
      endpointRule: parsed.endpointRule,
      binCount: parsed.binCount,
      binWidth: parsed.binWidth,
      startAt: parsed.startAt,
      niceRound: parsed.niceRound ?? true,
      intervals: parsed.intervals,
      sourceVariableName: parsed.sourceVariableName,
      savedAt: parsed.savedAt ?? new Date().toISOString(),
    };
  } catch {
    return null;
  }
}

function saveBinningTemplate(template: SavedBinningTemplate): boolean {
  if (typeof window === 'undefined') return false;
  try {
    window.localStorage.setItem(
      BINNING_TEMPLATE_STORAGE_KEY,
      JSON.stringify(template)
    );
    return true;
  } catch {
    return false;
  }
}

function clearBinningTemplate(): void {
  if (typeof window === 'undefined') return;
  try {
    window.localStorage.removeItem(BINNING_TEMPLATE_STORAGE_KEY);
  } catch {
    // Ignore storage errors in sandboxed frames
  }
}

function describeTemplate(tpl: SavedBinningTemplate): string {
  const rule = tpl.endpointRule === 'lower_inclusive' ? '[a, b)' : '(a, b]';
  if (tpl.mode === 'auto_count') {
    return `${tpl.binCount ?? tpl.intervals.length} equal-width bins ${rule}`;
  }
  if (tpl.mode === 'auto_width') {
    return `Width ${tpl.binWidth ?? 'auto'}${tpl.startAt !== undefined ? ` from ${tpl.startAt}` : ''} ${rule}`;
  }
  return `${tpl.intervals.length} custom intervals ${rule}`;
}

interface ClassBoundariesModalProps {
  variable: VariableDefinition;
  rows: DataRow[];
  existingNames: string[];
  onSave: (variableId: string, spec: ClassBinSpec | undefined) => void;
  onRecodeVariable: (
    sourceVarId: string,
    targetName: string,
    spec: ClassBinSpec
  ) => void;
  onClose: () => void;
}

export function ClassBoundariesModal({
  variable,
  rows,
  existingNames,
  onSave,
  onRecodeVariable,
  onClose,
}: ClassBoundariesModalProps) {
  const validValues = useMemo(
    () => getValidNumericValues(variable, rows),
    [variable, rows]
  );
  const baseStats = useMemo(
    () => computeVariableStats(variable, rows),
    [variable, rows]
  );

  const initialSpec = variable.classBins;
  const initialMode: ClassBinMode = initialSpec?.mode ?? 'auto_count';
  const initialEndpoint: EndpointInclusionRule =
    initialSpec?.endpointRule ?? 'lower_inclusive';

  const [mode, setMode] = useState<ClassBinMode>(initialMode);
  const [endpointRule, setEndpointRule] =
    useState<EndpointInclusionRule>(initialEndpoint);
  const [binCount, setBinCount] = useState<string>(
    String(initialSpec?.binCount ?? 5)
  );
  const [binWidth, setBinWidth] = useState<string>(
    initialSpec?.binWidth !== undefined ? String(initialSpec.binWidth) : ''
  );
  const [startAt, setStartAt] = useState<string>(
    initialSpec?.startAt !== undefined ? String(initialSpec.startAt) : ''
  );
  const [niceRound, setNiceRound] = useState<boolean>(true);

  const [intervals, setIntervals] = useState<ClassInterval[]>(() => {
    if (initialSpec?.enabled && initialSpec.intervals.length > 0) {
      return initialSpec.intervals.map((b) => ({ ...b }));
    }
    const generated = generateEqualWidthIntervals(validValues, {
      mode: 'auto_count',
      binCount: 5,
      endpointRule: initialEndpoint,
      niceRound: true,
      variable,
    });
    return generated.intervals;
  });

  // Saved template & remember checkbox state
  const [savedTemplate, setSavedTemplate] = useState<SavedBinningTemplate | null>(() =>
    loadBinningTemplate()
  );
  const [rememberSettings, setRememberSettings] = useState<boolean>(false);

  // Recode inline prompt state
  const [showRecodePrompt, setShowRecodePrompt] = useState(false);
  const [recodeName, setRecodeName] = useState(() =>
    sanitizeVariableName(`${variable.name}_grp`, existingNames)
  );
  const [recodeError, setRecodeError] = useState<string | null>(null);

  const validationErrors = useMemo(
    () => validateClassIntervals(intervals),
    [intervals]
  );

  const liveCounts = useMemo(() => {
    const counts = new Array<number>(intervals.length).fill(0);
    let outside = 0;
    if (validationErrors.length > 0) {
      return { counts, outside };
    }
    for (const v of validValues) {
      const idx = classifyValueIntoInterval(v, intervals, endpointRule);
      if (idx >= 0) counts[idx] += 1;
      else outside += 1;
    }
    return { counts, outside };
  }, [validValues, intervals, endpointRule, validationErrors.length]);

  const applyAutoGeneration = (overrides?: {
    nextMode?: 'auto_count' | 'auto_width';
    nextCount?: string;
    nextWidth?: string;
    nextStart?: string;
    nextNice?: boolean;
    nextRule?: EndpointInclusionRule;
  }) => {
    const targetMode =
      overrides?.nextMode ?? (mode === 'auto_width' ? 'auto_width' : 'auto_count');
    const cVal = Number(overrides?.nextCount ?? binCount);
    const wVal = Number(overrides?.nextWidth ?? binWidth);
    const sStr = (overrides?.nextStart ?? startAt).trim();
    const sVal = sStr !== '' && !Number.isNaN(Number(sStr)) ? Number(sStr) : undefined;
    const nice = overrides?.nextNice ?? niceRound;
    const rule = overrides?.nextRule ?? endpointRule;

    const result = generateEqualWidthIntervals(validValues, {
      mode: targetMode,
      binCount: !Number.isNaN(cVal) && cVal >= 2 ? cVal : 5,
      binWidth: !Number.isNaN(wVal) && wVal > 0 ? wVal : undefined,
      startAt: sVal,
      endpointRule: rule,
      niceRound: nice,
      variable,
    });

    setIntervals(result.intervals);
    if (targetMode === 'auto_width' && (overrides?.nextWidth ?? binWidth).trim() === '') {
      setBinWidth(String(result.resolvedWidth));
    }
  };

  const handleModeSwitch = (next: ClassBinMode) => {
    setMode(next);
    if (next === 'auto_count' || next === 'auto_width') {
      applyAutoGeneration({ nextMode: next });
    }
  };

  const handleEndpointSwitch = (nextRule: EndpointInclusionRule) => {
    const prevRule = endpointRule;
    setEndpointRule(nextRule);
    setIntervals((prev) => {
      const hasOpenLower = prev[0]?.lower === null;
      const hasOpenUpper = prev[prev.length - 1]?.upper === null;
      return prev.map((bin, idx) => {
        const isFirstFinite = hasOpenLower ? idx === 1 : idx === 0;
        const isLastFinite = hasOpenUpper
          ? idx === prev.length - 2
          : idx === prev.length - 1;
        const oldAuto = buildIntervalLabel(
          bin.lower,
          bin.upper,
          prevRule,
          isFirstFinite,
          isLastFinite,
          variable
        );
        const newAuto = buildIntervalLabel(
          bin.lower,
          bin.upper,
          nextRule,
          isFirstFinite,
          isLastFinite,
          variable
        );
        return {
          ...bin,
          label: bin.label === oldAuto || bin.label.trim() === '' ? newAuto : bin.label,
        };
      });
    });
  };

  const regenerateAllLabels = () => {
    setIntervals((prev) => {
      const hasOpenLower = prev[0]?.lower === null;
      const hasOpenUpper = prev[prev.length - 1]?.upper === null;
      return prev.map((bin, idx) => {
        const isFirstFinite = hasOpenLower ? idx === 1 : idx === 0;
        const isLastFinite = hasOpenUpper
          ? idx === prev.length - 2
          : idx === prev.length - 1;
        return {
          ...bin,
          label: buildIntervalLabel(
            bin.lower,
            bin.upper,
            endpointRule,
            isFirstFinite,
            isLastFinite,
            variable
          ),
        };
      });
    });
  };

  const updateIntervalRow = (
    idx: number,
    patch: Partial<ClassInterval>,
    autoLabel = true
  ) => {
    setMode('custom');
    setIntervals((prev) => {
      const next = prev.map((b, i) => (i === idx ? { ...b, ...patch } : { ...b }));
      const hasOpenLower = next[0]?.lower === null;
      const hasOpenUpper = next[next.length - 1]?.upper === null;

      if (autoLabel) {
        const prevRow = prev[idx];
        const wasOpenLower = prev[0]?.lower === null;
        const wasOpenUpper = prev[prev.length - 1]?.upper === null;
        const oldAuto = buildIntervalLabel(
          prevRow.lower,
          prevRow.upper,
          endpointRule,
          wasOpenLower ? idx === 1 : idx === 0,
          wasOpenUpper ? idx === prev.length - 2 : idx === prev.length - 1,
          variable
        );
        const cur = next[idx];
        if (
          patch.label === undefined &&
          (cur.label === oldAuto || cur.label.trim() === '')
        ) {
          cur.label = buildIntervalLabel(
            cur.lower,
            cur.upper,
            endpointRule,
            hasOpenLower ? idx === 1 : idx === 0,
            hasOpenUpper ? idx === next.length - 2 : idx === next.length - 1,
            variable
          );
        }
      }
      return next;
    });
  };

  const handleAddInterval = () => {
    setMode('custom');
    setIntervals((prev) => {
      const last = prev[prev.length - 1];
      const prevWidth =
        prev.length >= 1 && last.lower !== null && last.upper !== null
          ? last.upper - last.lower
          : 10;

      // If previous last was open-ended (+∞), close it first before appending
      const updated = prev.map((b, idx) => {
        if (idx === prev.length - 1 && b.upper === null) {
          const closedUpper =
            b.lower !== null ? Math.round((b.lower + 10) * 1e6) / 1e6 : 10;
          return {
            ...b,
            upper: closedUpper,
            label: buildIntervalLabel(
              b.lower,
              closedUpper,
              endpointRule,
              idx === 0,
              false,
              variable
            ),
          };
        }
        return b;
      });
      const tail = updated[updated.length - 1];
      const nextLower = tail?.upper ?? 0;
      const nextUpper = Math.round((nextLower + Math.max(1, prevWidth)) * 1e6) / 1e6;
      const nextList = [
        ...updated,
        {
          id: `bin_${updated.length + 1}_${Date.now()}`,
          lower: nextLower,
          upper: nextUpper,
          label: buildIntervalLabel(
            nextLower,
            nextUpper,
            endpointRule,
            false,
            true,
            variable
          ),
        },
      ];
      return nextList;
    });
  };

  const handleDeleteInterval = (idx: number) => {
    if (intervals.length <= 1) return;
    setMode('custom');
    setIntervals((prev) => prev.filter((_, i) => i !== idx));
  };

  const buildCurrentSpec = (): ClassBinSpec => {
    const cVal = Number(binCount);
    const wVal = Number(binWidth);
    const sVal = startAt.trim() !== '' ? Number(startAt) : undefined;
    return {
      enabled: true,
      mode,
      endpointRule,
      binCount: mode === 'auto_count' && !Number.isNaN(cVal) ? cVal : undefined,
      binWidth: mode === 'auto_width' && !Number.isNaN(wVal) ? wVal : undefined,
      startAt:
        sVal !== undefined && !Number.isNaN(sVal) ? sVal : undefined,
      intervals,
    };
  };

  const persistCurrentAsTemplate = () => {
    const spec = buildCurrentSpec();
    const tpl: SavedBinningTemplate = {
      mode: spec.mode,
      endpointRule: spec.endpointRule,
      binCount: spec.binCount,
      binWidth: spec.binWidth,
      startAt: spec.startAt,
      niceRound,
      intervals: spec.intervals.map((b) => ({ ...b })),
      sourceVariableName: variable.name,
      savedAt: new Date().toISOString(),
    };
    if (saveBinningTemplate(tpl)) {
      setSavedTemplate(tpl);
    }
  };

  const handleApplySavedTemplate = () => {
    if (!savedTemplate) return;
    setMode(savedTemplate.mode);
    setEndpointRule(savedTemplate.endpointRule);
    setNiceRound(savedTemplate.niceRound);
    if (savedTemplate.binCount !== undefined) {
      setBinCount(String(savedTemplate.binCount));
    }
    setBinWidth(
      savedTemplate.binWidth !== undefined ? String(savedTemplate.binWidth) : ''
    );
    setStartAt(
      savedTemplate.startAt !== undefined ? String(savedTemplate.startAt) : ''
    );

    if (
      savedTemplate.mode === 'auto_count' ||
      savedTemplate.mode === 'auto_width'
    ) {
      const result = generateEqualWidthIntervals(validValues, {
        mode: savedTemplate.mode,
        binCount: savedTemplate.binCount,
        binWidth: savedTemplate.binWidth,
        startAt: savedTemplate.startAt,
        endpointRule: savedTemplate.endpointRule,
        niceRound: savedTemplate.niceRound,
        variable,
      });
      setIntervals(result.intervals);
    } else {
      setIntervals(
        savedTemplate.intervals.map((b, idx) => ({
          ...b,
          id: `bin_tpl_${idx + 1}_${Date.now()}`,
        }))
      );
    }
  };

  const handleForgetTemplate = () => {
    clearBinningTemplate();
    setSavedTemplate(null);
    setRememberSettings(false);
  };

  const handleSave = () => {
    if (validationErrors.length > 0) return;
    if (rememberSettings) {
      persistCurrentAsTemplate();
    }
    onSave(variable.id, buildCurrentSpec());
    onClose();
  };

  const handleClearBins = () => {
    onSave(variable.id, undefined);
    onClose();
  };

  const handleConfirmRecode = () => {
    if (validationErrors.length > 0) return;
    const err = validateVariableName(recodeName, existingNames);
    if (err) {
      setRecodeError(err);
      return;
    }
    if (rememberSettings) {
      persistCurrentAsTemplate();
    }
    const spec = buildCurrentSpec();
    onSave(variable.id, spec);
    onRecodeVariable(variable.id, recodeName.trim(), spec);
    onClose();
  };

  const validTotal = validValues.length;
  const maxBinCount = Math.max(1, ...liveCounts.counts);

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4"
      role="dialog"
      aria-modal="true"
      aria-labelledby="class-bins-modal-title"
    >
      <div className="bg-white border border-slate-300 rounded-lg shadow-xl w-full max-w-3xl flex flex-col max-h-[90vh]">
        {/* Header */}
        <div className="px-5 py-3.5 border-b border-slate-200 flex items-center justify-between gap-4">
          <div className="min-w-0">
            <h2
              id="class-bins-modal-title"
              className="font-display text-sm font-semibold text-slate-900"
            >
              Class Boundaries — <span className="font-mono">{variable.name}</span>
            </h2>
            <p className="text-xs text-slate-500 mt-0.5 truncate">
              Group continuous scale values into intervals for frequency distributions or recode into an ordinal variable.
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

        {/* Observed Valid Numeric Summary */}
        <div className="px-5 py-2.5 bg-slate-50 border-b border-slate-200 flex flex-wrap items-center justify-between gap-3 text-xs text-slate-600">
          <div className="flex flex-wrap items-center gap-3 font-mono tabular-nums">
            <span>
              Valid N: <strong className="text-slate-900">{baseStats.validN}</strong>
            </span>
            <span aria-hidden="true">·</span>
            <span>
              Min:{' '}
              <strong className="text-slate-900">
                {baseStats.min !== null ? baseStats.min : '—'}
              </strong>
            </span>
            <span aria-hidden="true">·</span>
            <span>
              Max:{' '}
              <strong className="text-slate-900">
                {baseStats.max !== null ? baseStats.max : '—'}
              </strong>
            </span>
            {(baseStats.userMissingN > 0 || baseStats.systemMissingN > 0) && (
              <>
                <span aria-hidden="true">·</span>
                <span className="text-amber-800 font-sans">
                  {baseStats.userMissingN + baseStats.systemMissingN} missing excluded
                </span>
              </>
            )}
          </div>
          {liveCounts.outside > 0 && validationErrors.length === 0 && (
            <span className="text-amber-800 font-medium">
              {liveCounts.outside} valid case(s) fall outside current bounds
            </span>
          )}
        </div>

        {/* Body */}
        <div className="p-5 overflow-y-auto space-y-4 text-xs">
          {/* Saved Binning Template Banner */}
          {savedTemplate && (
            <div className="px-3.5 py-2.5 border border-blue-200 bg-blue-50/60 rounded-md flex flex-wrap items-center justify-between gap-3">
              <div className="flex items-center gap-2 min-w-0 text-slate-800">
                <BookmarkCheck className="w-4 h-4 text-blue-700 shrink-0" />
                <div className="truncate">
                  <span className="font-semibold">Saved template:</span>{' '}
                  <span className="font-mono text-slate-700">
                    {describeTemplate(savedTemplate)}
                  </span>
                  {savedTemplate.sourceVariableName && (
                    <span className="text-slate-500">
                      {' '}
                      · from <span className="font-mono">{savedTemplate.sourceVariableName}</span>
                    </span>
                  )}
                </div>
              </div>
              <div className="flex items-center gap-2 shrink-0">
                <button
                  type="button"
                  onClick={handleApplySavedTemplate}
                  className="px-2.5 py-1 text-[11px] font-semibold text-white bg-blue-700 hover:bg-blue-800 rounded cursor-pointer"
                >
                  Apply Saved Template
                </button>
                <button
                  type="button"
                  onClick={handleForgetTemplate}
                  className="px-2 py-1 text-[11px] font-medium text-slate-600 hover:text-red-700 bg-white border border-slate-300 rounded cursor-pointer"
                >
                  Forget
                </button>
              </div>
            </div>
          )}

          {/* Mode + Endpoint Rule Controls */}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div className="p-3 border border-slate-200 rounded-md bg-slate-50/70 space-y-2.5">
              <div className="font-display font-semibold text-slate-900">
                Binning Method
              </div>
              <div className="flex items-center gap-1 p-1 bg-slate-200/70 rounded-md">
                {(
                  [
                    { id: 'auto_count', label: 'Equal Count' },
                    { id: 'auto_width', label: 'Fixed Width' },
                    { id: 'custom', label: 'Custom Intervals' },
                  ] as const
                ).map((item) => (
                  <button
                    key={item.id}
                    type="button"
                    onClick={() => handleModeSwitch(item.id)}
                    className={`flex-1 px-2.5 py-1.5 text-[11px] font-medium rounded transition-colors cursor-pointer whitespace-nowrap ${
                      mode === item.id
                        ? 'bg-white text-slate-900 shadow-xs'
                        : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    {item.label}
                  </button>
                ))}
              </div>

              {mode !== 'custom' && (
                <div className="grid grid-cols-2 gap-2.5 pt-1">
                  {mode === 'auto_count' ? (
                    <div>
                      <label
                        htmlFor="bin-count-input"
                        className="block text-[11px] text-slate-600 mb-1"
                      >
                        Number of Bins (2–30)
                      </label>
                      <input
                        id="bin-count-input"
                        type="number"
                        min={2}
                        max={30}
                        value={binCount}
                        onChange={(e) => {
                          setBinCount(e.target.value);
                          applyAutoGeneration({
                            nextMode: 'auto_count',
                            nextCount: e.target.value,
                          });
                        }}
                        className="w-full px-2.5 py-1.5 font-mono tabular-nums bg-white border border-slate-300 rounded focus:outline-none focus:ring-2 focus:ring-blue-600"
                      />
                    </div>
                  ) : (
                    <div>
                      <label
                        htmlFor="bin-width-input"
                        className="block text-[11px] text-slate-600 mb-1"
                      >
                        Interval Width
                      </label>
                      <input
                        id="bin-width-input"
                        type="number"
                        step="any"
                        min="0.0001"
                        value={binWidth}
                        placeholder="Auto"
                        onChange={(e) => {
                          setBinWidth(e.target.value);
                          applyAutoGeneration({
                            nextMode: 'auto_width',
                            nextWidth: e.target.value,
                          });
                        }}
                        className="w-full px-2.5 py-1.5 font-mono tabular-nums bg-white border border-slate-300 rounded focus:outline-none focus:ring-2 focus:ring-blue-600"
                      />
                    </div>
                  )}

                  <div>
                    <label
                      htmlFor="bin-start-input"
                      className="block text-[11px] text-slate-600 mb-1"
                    >
                      Start Bound (opt.)
                    </label>
                    <input
                      id="bin-start-input"
                      type="number"
                      step="any"
                      value={startAt}
                      placeholder="Auto from min"
                      onChange={(e) => {
                        setStartAt(e.target.value);
                        applyAutoGeneration({ nextStart: e.target.value });
                      }}
                      className="w-full px-2.5 py-1.5 font-mono tabular-nums bg-white border border-slate-300 rounded focus:outline-none focus:ring-2 focus:ring-blue-600"
                    />
                  </div>

                  <label className="col-span-2 flex items-center gap-2 cursor-pointer text-[11px] text-slate-700">
                    <input
                      type="checkbox"
                      checked={niceRound}
                      onChange={(e) => {
                        setNiceRound(e.target.checked);
                        applyAutoGeneration({ nextNice: e.target.checked });
                      }}
                      className="rounded border-slate-300 text-blue-700 focus:ring-blue-600"
                    />
                    <span>Snap cutpoints to nice rounded steps</span>
                  </label>
                </div>
              )}

              {mode === 'custom' && (
                <p className="text-[11px] text-slate-600 leading-relaxed pt-1">
                  Edit lower/upper bounds directly. Contiguous intervals ensure every observation maps cleanly without overlaps or gaps.
                </p>
              )}
            </div>

            {/* Endpoint Inclusion Rule */}
            <div className="p-3 border border-slate-200 rounded-md bg-slate-50/70 space-y-2.5">
              <div className="font-display font-semibold text-slate-900">
                Boundary Endpoint Rule
              </div>
              <div className="space-y-2">
                <label className="flex items-start gap-2 cursor-pointer">
                  <input
                    type="radio"
                    name="endpointRule"
                    checked={endpointRule === 'lower_inclusive'}
                    onChange={() => handleEndpointSwitch('lower_inclusive')}
                    className="mt-0.5 text-blue-700 focus:ring-blue-600"
                  />
                  <div>
                    <div className="font-medium text-slate-900">
                      Lower-inclusive <span className="font-mono">[a, b)</span>
                    </div>
                    <div className="text-[11px] text-slate-500">
                      Default. Values equal to a cutpoint belong to the upper bin; final finite bin is closed [a, b].
                    </div>
                  </div>
                </label>
                <label className="flex items-start gap-2 cursor-pointer">
                  <input
                    type="radio"
                    name="endpointRule"
                    checked={endpointRule === 'upper_inclusive'}
                    onChange={() => handleEndpointSwitch('upper_inclusive')}
                    className="mt-0.5 text-blue-700 focus:ring-blue-600"
                  />
                  <div>
                    <div className="font-medium text-slate-900">
                      Upper-inclusive <span className="font-mono">(a, b]</span>
                    </div>
                    <div className="text-[11px] text-slate-500">
                      Values equal to a cutpoint belong to the lower bin; first finite bin is closed [a, b].
                    </div>
                  </div>
                </label>
              </div>
            </div>
          </div>

          {/* Validation Errors */}
          {validationErrors.length > 0 && (
            <div className="p-3 border border-red-300 bg-red-50 rounded-md text-red-800 space-y-1">
              <div className="flex items-center gap-1.5 font-semibold">
                <AlertCircle className="w-4 h-4 shrink-0" />
                <span>Fix interval boundary issues before saving:</span>
              </div>
              <ul className="list-disc pl-5 text-[11px] space-y-0.5">
                {validationErrors.map((err, i) => (
                  <li key={i}>{err}</li>
                ))}
              </ul>
            </div>
          )}

          {/* Intervals Table */}
          <div className="border border-slate-200 rounded-md overflow-hidden">
            <div className="px-3.5 py-2 bg-slate-100 border-b border-slate-200 flex items-center justify-between gap-3">
              <span className="font-semibold text-slate-800">
                Class Intervals ({intervals.length})
              </span>
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={regenerateAllLabels}
                  className="inline-flex items-center gap-1 px-2 py-1 text-[11px] font-medium text-slate-700 bg-white border border-slate-300 hover:bg-slate-50 rounded cursor-pointer"
                >
                  <RefreshCw className="w-3 h-3" />
                  <span>Reset Labels</span>
                </button>
                <button
                  type="button"
                  onClick={handleAddInterval}
                  className="inline-flex items-center gap-1 px-2.5 py-1 text-[11px] font-semibold text-blue-700 bg-white border border-blue-200 hover:bg-blue-50 rounded cursor-pointer"
                >
                  <Plus className="w-3 h-3" />
                  <span>Add Interval</span>
                </button>
              </div>
            </div>

            <div className="max-h-64 overflow-y-auto">
              <table className="w-full border-collapse text-xs">
                <thead className="bg-slate-50 border-b border-slate-200 text-slate-600 font-semibold sticky top-0 z-10">
                  <tr>
                    <th className="w-10 px-2.5 py-2 text-center font-mono">#</th>
                    <th className="w-36 px-2.5 py-2 text-left">Lower Bound</th>
                    <th className="w-36 px-2.5 py-2 text-left">Upper Bound</th>
                    <th className="px-2.5 py-2 text-left">Interval Label</th>
                    <th className="w-32 px-2.5 py-2 text-right">Valid Cases</th>
                    <th className="w-10 px-2 py-2 text-center">Del</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200">
                  {intervals.map((bin, idx) => {
                    const isFirst = idx === 0;
                    const isLast = idx === intervals.length - 1;
                    const count = liveCounts.counts[idx] ?? 0;
                    const pct =
                      validTotal > 0 ? ((count / validTotal) * 100).toFixed(1) : '0.0';

                    return (
                      <tr key={bin.id} className="hover:bg-slate-50">
                        <td className="px-2.5 py-1.5 text-center font-mono text-slate-500 tabular-nums">
                          {idx + 1}
                        </td>
                        <td className="px-2.5 py-1.5">
                          <div className="flex items-center gap-1.5">
                            {bin.lower === null ? (
                              <span className="flex-1 px-2 py-1 font-mono text-slate-500 bg-slate-100 border border-slate-200 rounded">
                                −∞ (Open)
                              </span>
                            ) : (
                              <input
                                type="number"
                                step="any"
                                aria-label={`Interval ${idx + 1} lower bound`}
                                value={bin.lower}
                                onChange={(e) => {
                                  const num = Number(e.target.value);
                                  if (!Number.isNaN(num)) {
                                    updateIntervalRow(idx, { lower: num });
                                  }
                                }}
                                className="w-full px-2 py-1 font-mono tabular-nums bg-white border border-slate-300 rounded focus:outline-none focus:ring-2 focus:ring-blue-600"
                              />
                            )}
                            {isFirst && (
                              <label
                                className="inline-flex items-center gap-1 text-[10px] text-slate-600 cursor-pointer shrink-0"
                                title="Open lower bound (−∞)"
                              >
                                <input
                                  type="checkbox"
                                  checked={bin.lower === null}
                                  onChange={(e) =>
                                    updateIntervalRow(idx, {
                                      lower: e.target.checked
                                        ? null
                                        : (bin.upper ?? 10) - 10,
                                    })
                                  }
                                />
                                <span>−∞</span>
                              </label>
                            )}
                          </div>
                        </td>
                        <td className="px-2.5 py-1.5">
                          <div className="flex items-center gap-1.5">
                            {bin.upper === null ? (
                              <span className="flex-1 px-2 py-1 font-mono text-slate-500 bg-slate-100 border border-slate-200 rounded">
                                +∞ (Open)
                              </span>
                            ) : (
                              <input
                                type="number"
                                step="any"
                                aria-label={`Interval ${idx + 1} upper bound`}
                                value={bin.upper}
                                onChange={(e) => {
                                  const num = Number(e.target.value);
                                  if (!Number.isNaN(num)) {
                                    updateIntervalRow(idx, { upper: num });
                                  }
                                }}
                                className="w-full px-2 py-1 font-mono tabular-nums bg-white border border-slate-300 rounded focus:outline-none focus:ring-2 focus:ring-blue-600"
                              />
                            )}
                            {isLast && (
                              <label
                                className="inline-flex items-center gap-1 text-[10px] text-slate-600 cursor-pointer shrink-0"
                                title="Open upper bound (+∞)"
                              >
                                <input
                                  type="checkbox"
                                  checked={bin.upper === null}
                                  onChange={(e) =>
                                    updateIntervalRow(idx, {
                                      upper: e.target.checked
                                        ? null
                                        : (bin.lower ?? 0) + 10,
                                    })
                                  }
                                />
                                <span>+∞</span>
                              </label>
                            )}
                          </div>
                        </td>
                        <td className="px-2.5 py-1.5">
                          <input
                            type="text"
                            aria-label={`Interval ${idx + 1} label`}
                            value={bin.label}
                            onChange={(e) =>
                              updateIntervalRow(idx, { label: e.target.value }, false)
                            }
                            className="w-full px-2 py-1 bg-white border border-slate-300 rounded focus:outline-none focus:ring-2 focus:ring-blue-600"
                          />
                        </td>
                        <td className="px-2.5 py-1.5 text-right">
                          <div className="font-mono tabular-nums text-slate-800">
                            {count}{' '}
                            <span className="text-slate-400 text-[11px]">({pct}%)</span>
                          </div>
                          <div className="mt-0.5 h-1 w-full bg-slate-100 rounded-full overflow-hidden">
                            <div
                              className="h-full bg-blue-600 rounded-full"
                              style={{
                                width: `${Math.round((count / maxBinCount) * 100)}%`,
                              }}
                            />
                          </div>
                        </td>
                        <td className="px-2 py-1.5 text-center">
                          <button
                            type="button"
                            disabled={intervals.length <= 1}
                            onClick={() => handleDeleteInterval(idx)}
                            aria-label={`Delete interval ${idx + 1}`}
                            className="p-1 text-slate-400 hover:text-red-600 disabled:opacity-30 cursor-pointer"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                          </button>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>

          {/* Inline Recode into New Variable prompt */}
          {showRecodePrompt && (
            <div className="p-3.5 border border-blue-200 bg-blue-50/60 rounded-md flex flex-wrap items-center justify-between gap-3">
              <div className="flex-1 min-w-[220px]">
                <label
                  htmlFor="recode-target-name"
                  className="block text-[11px] font-semibold text-slate-800 mb-1"
                >
                  New Coded Ordinal Variable Name (codes 1–{intervals.length} + Value Labels)
                </label>
                <input
                  id="recode-target-name"
                  type="text"
                  value={recodeName}
                  onChange={(e) => {
                    setRecodeName(e.target.value);
                    setRecodeError(null);
                  }}
                  className="w-full max-w-xs px-2.5 py-1.5 font-mono text-xs bg-white border border-slate-300 rounded focus:outline-none focus:ring-2 focus:ring-blue-600"
                />
                {recodeError && (
                  <p className="text-[11px] text-red-700 mt-1">{recodeError}</p>
                )}
              </div>
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={() => setShowRecodePrompt(false)}
                  className="px-2.5 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-300 rounded hover:bg-slate-50 cursor-pointer"
                >
                  Cancel
                </button>
                <button
                  type="button"
                  disabled={validationErrors.length > 0}
                  onClick={handleConfirmRecode}
                  className="px-3 py-1.5 text-xs font-semibold text-white bg-[#107C41] hover:bg-[#0D6635] disabled:opacity-50 rounded cursor-pointer"
                >
                  Create Binned Variable
                </button>
              </div>
            </div>
          )}
        </div>

        {/* Remember Template Option */}
        <div className="px-5 py-2.5 bg-white border-t border-slate-200 flex flex-wrap items-center justify-between gap-2 text-xs">
          <label className="inline-flex items-center gap-2 cursor-pointer text-slate-700 select-none">
            <input
              type="checkbox"
              checked={rememberSettings}
              onChange={(e) => {
                const checked = e.target.checked;
                setRememberSettings(checked);
                if (checked && validationErrors.length === 0) {
                  persistCurrentAsTemplate();
                }
              }}
              className="rounded border-slate-300 text-blue-700 focus:ring-blue-600"
            />
            <span className="font-medium text-slate-800">
              Remember these binning settings
            </span>
          </label>
          <span className="text-[11px] text-slate-500">
            Saves this interval configuration in browser storage to reuse on other variables.
          </span>
        </div>

        {/* Footer */}
        <div className="px-5 py-3 bg-slate-50 border-t border-slate-200 flex flex-wrap items-center justify-between gap-2">
          <div className="flex items-center gap-2">
            {variable.classBins?.enabled && (
              <button
                type="button"
                onClick={handleClearBins}
                className="px-3 py-1.5 text-xs font-medium text-red-700 bg-white border border-red-200 rounded-md hover:bg-red-50 cursor-pointer"
              >
                Clear Bins
              </button>
            )}
            <button
              type="button"
              disabled={validationErrors.length > 0}
              onClick={() => setShowRecodePrompt((s) => !s)}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-slate-800 bg-white border border-slate-300 rounded-md hover:bg-slate-100 disabled:opacity-50 cursor-pointer"
            >
              <ArrowRightLeft className="w-3.5 h-3.5 text-[#107C41]" />
              <span>Recode into New Variable…</span>
            </button>
          </div>

          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={onClose}
              className="px-3.5 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-300 rounded-md hover:bg-slate-50 cursor-pointer"
            >
              Cancel
            </button>
            <button
              type="button"
              disabled={validationErrors.length > 0}
              onClick={handleSave}
              className="px-4 py-1.5 text-xs font-semibold text-white bg-blue-700 hover:bg-blue-800 disabled:opacity-50 rounded-md cursor-pointer"
            >
              Save Class Boundaries
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}
