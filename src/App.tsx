import { AlertCircle, CheckCircle2, Info, X } from 'lucide-react';
import { useCallback, useEffect, useMemo, useState } from 'react';
import {
  detectOfficeHost,
  exportDatasetToXlsx,
  importWorkbookFile,
  isOfficeHostContext,
  loadSimulatorState,
  readFromLiveExcel,
  saveSimulatorState,
  syncToLiveExcel,
} from './bridge/workbookBridge';
import { BottomViewTabs } from './components/BottomViewTabs';
import { CodebookModal } from './components/CodebookModal';
import { DataViewGrid } from './components/DataViewGrid';
import { ExcelAddinModal } from './components/ExcelAddinModal';
import { MissingAndValidationModal } from './components/MissingAndValidationModal';
import { RibbonToolbar } from './components/RibbonToolbar';
import { StatsInspectorDrawer } from './components/StatsInspectorDrawer';
import { TopBar } from './components/TopBar';
import { ValueLabelsModal } from './components/ValueLabelsModal';
import { VariableViewGrid } from './components/VariableViewGrid';
import { EMPLOYEE_SURVEY_DATASET, SAMPLE_PRESETS } from './data/sampleDatasets';
import { sanitizeVariableName, validateDataset } from './engine/spssEngine';
import type {
  ActiveSheetView,
  CellValue,
  DatasetState,
  HostEnvironment,
  MissingValueSpec,
  ToastMessage,
  ValidationRuleSpec,
  VariableDefinition,
} from './types/spss';

const MAX_HISTORY = 40;

export default function App() {
  const [host, setHost] = useState<HostEnvironment>('simulator');
  const [isInitializing, setIsInitializing] = useState(() => isOfficeHostContext());
  const [isSyncing, setIsSyncing] = useState(false);

  const [dataset, setDataset] = useState<DatasetState>(() => {
    return loadSimulatorState() ?? EMPLOYEE_SURVEY_DATASET;
  });

  const [past, setPast] = useState<DatasetState[]>([]);
  const [future, setFuture] = useState<DatasetState[]>([]);

  const [selectedCell, setSelectedCell] = useState<{
    rowId: string;
    variableId: string;
  } | null>(null);
  const [selectedVariableId, setSelectedVariableId] = useState<string | null>(
    () => dataset.variables[0]?.id ?? null
  );

  const [searchQuery, setSearchQuery] = useState('');
  const [issuesOnly, setIssuesOnly] = useState(false);
  const [taskPanePreview, setTaskPanePreview] = useState(false);

  // Modals & Drawers
  const [valueLabelsVarId, setValueLabelsVarId] = useState<string | null>(null);
  const [missingVarId, setMissingVarId] = useState<string | null>(null);
  const [statsVarId, setStatsVarId] = useState<string | null>(null);
  const [showCodebookModal, setShowCodebookModal] = useState(false);
  const [showAddinModal, setShowAddinModal] = useState(false);

  const [toasts, setToasts] = useState<ToastMessage[]>([]);

  const pushToast = useCallback(
    (
      tone: ToastMessage['tone'],
      title: string,
      detail?: string
    ) => {
      const id = `toast_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`;
      setToasts((prev) => [...prev.slice(-2), { id, tone, title, detail }]);
      window.setTimeout(() => {
        setToasts((prev) => prev.filter((t) => t.id !== id));
      }, 4200);
    },
    []
  );

  // Detect Office.js Excel host vs standalone simulator
  useEffect(() => {
    let active = true;
    detectOfficeHost()
      .then(async ({ host: detected }) => {
        if (!active) return;
        setHost(detected);
        if (detected === 'excel') {
          try {
            const excelData = await readFromLiveExcel();
            if (active && excelData.variables.length > 0) {
              setDataset(excelData);
              setSelectedVariableId(excelData.variables[0]?.id ?? null);
            }
          } catch {
            // Keep default state if sheet is blank or inaccessible initially
          }
        }
      })
      .finally(() => {
        if (active) setIsInitializing(false);
      });

    return () => {
      active = false;
    };
  }, []);

  // Persist simulator dataset in localStorage
  useEffect(() => {
    if (host === 'simulator' && !isInitializing) {
      saveSimulatorState(dataset);
    }
  }, [dataset, host, isInitializing]);

  const commitDataset = useCallback(
    (updater: (prev: DatasetState) => DatasetState, recordHistory = true) => {
      setDataset((prev) => {
        const next = updater(prev);
        if (recordHistory) {
          setPast((p) => [...p.slice(-(MAX_HISTORY - 1)), prev]);
          setFuture([]);
        }
        return next;
      });
    },
    []
  );

  const handleUndo = () => {
    if (past.length === 0) return;
    const previous = past[past.length - 1];
    setPast((p) => p.slice(0, -1));
    setFuture((f) => [dataset, ...f]);
    setDataset(previous);
  };

  const handleRedo = () => {
    if (future.length === 0) return;
    const next = future[0];
    setFuture((f) => f.slice(1));
    setPast((p) => [...p, dataset]);
    setDataset(next);
  };

  const issues = useMemo(
    () => validateDataset(dataset.variables, dataset.rows),
    [dataset.variables, dataset.rows]
  );

  const handleToggleValueLabels = async () => {
    const nextShow = !dataset.showValueLabels;
    commitDataset((prev) => ({
      ...prev,
      showValueLabels: nextShow,
      lastBackupTimestamp: new Date().toISOString(),
    }));

    if (host === 'excel') {
      try {
        setIsSyncing(true);
        const ts = await syncToLiveExcel(dataset, nextShow);
        setDataset((prev) => ({ ...prev, lastBackupTimestamp: ts }));
        pushToast(
          'success',
          nextShow
            ? 'Value Labels applied to Excel sheet'
            : 'Raw numeric codes restored in Excel',
          'Original raw values preserved in _SPSS_RawBackup.'
        );
      } catch (err) {
        pushToast(
          'error',
          'Excel sync failed',
          err instanceof Error ? err.message : 'Unable to write to worksheet.'
        );
      } finally {
        setIsSyncing(false);
      }
    }
  };

  const handlePrimaryAction = async () => {
    if (host === 'excel') {
      try {
        setIsSyncing(true);
        const ts = await syncToLiveExcel(dataset, dataset.showValueLabels);
        setDataset((prev) => ({ ...prev, lastBackupTimestamp: ts }));
        pushToast(
          'success',
          'Workbook synced with Excel',
          'Saved _SPSS_RawBackup, _SPSS_Metadata, and validation dropdowns.'
        );
      } catch (err) {
        pushToast(
          'error',
          'Failed to sync workbook',
          err instanceof Error ? err.message : 'Check Excel permissions.'
        );
      } finally {
        setIsSyncing(false);
      }
    } else {
      exportDatasetToXlsx(dataset, dataset.showValueLabels);
      setDataset((prev) => ({
        ...prev,
        lastBackupTimestamp: new Date().toISOString(),
      }));
      pushToast(
        'success',
        'Exported .xlsx workbook',
        `Includes Data (${
          dataset.showValueLabels ? 'Value Labels' : 'Raw Codes'
        }), hidden _SPSS_RawBackup, and _SPSS_Metadata.`
      );
    }
  };

  const handleRefreshFromExcel = async () => {
    try {
      setIsSyncing(true);
      const fresh = await readFromLiveExcel();
      commitDataset(() => fresh);
      setSelectedVariableId(fresh.variables[0]?.id ?? null);
      pushToast('info', 'Loaded active Excel worksheet', `${fresh.rows.length} cases read.`);
    } catch (err) {
      pushToast(
        'error',
        'Could not read worksheet',
        err instanceof Error ? err.message : 'Ensure Excel sheet has headers in row 1.'
      );
    } finally {
      setIsSyncing(false);
    }
  };

  const handleRestoreRawBackup = async () => {
    try {
      setIsSyncing(true);
      const ts = await syncToLiveExcel(dataset, false);
      commitDataset((prev) => ({
        ...prev,
        showValueLabels: false,
        lastBackupTimestamp: ts,
      }));
      pushToast(
        'success',
        'Raw codes restored to Excel',
        'Active worksheet now displays raw numeric/string codes.'
      );
    } catch (err) {
      pushToast(
        'error',
        'Restore failed',
        err instanceof Error ? err.message : 'Unable to restore backup.'
      );
    } finally {
      setIsSyncing(false);
    }
  };

  const handleUpdateCell = (
    rowId: string,
    variableName: string,
    nextValue: CellValue
  ) => {
    commitDataset((prev) => ({
      ...prev,
      rows: prev.rows.map((r) =>
        r.id === rowId
          ? { ...r, values: { ...r.values, [variableName]: nextValue } }
          : r
      ),
    }));
  };

  const handleAddRow = () => {
    const newId = `row_${Date.now()}`;
    const emptyValues: Record<string, CellValue> = {};
    dataset.variables.forEach((v) => {
      emptyValues[v.name] = null;
    });
    commitDataset((prev) => ({
      ...prev,
      rows: [...prev.rows, { id: newId, values: emptyValues }],
    }));
    if (dataset.variables[0]) {
      setSelectedCell({ rowId: newId, variableId: dataset.variables[0].id });
    }
  };

  const handleDeleteRow = (rowId: string) => {
    commitDataset((prev) => ({
      ...prev,
      rows: prev.rows.filter((r) => r.id !== rowId),
    }));
    if (selectedCell?.rowId === rowId) {
      setSelectedCell(null);
    }
  };

  const handleAddVariable = () => {
    const existingNames = dataset.variables.map((v) => v.name);
    const nextName = sanitizeVariableName(
      `var_${dataset.variables.length + 1}`,
      existingNames
    );
    const newVar: VariableDefinition = {
      id: `var_${Date.now()}`,
      name: nextName,
      type: 'numeric',
      width: 8,
      decimals: 0,
      label: '',
      valueLabels: {},
      missingValues: { mode: 'none', discrete: [] },
      columns: 128,
      align: 'right',
      measure: 'scale',
      role: 'input',
    };
    commitDataset((prev) => ({
      ...prev,
      variables: [...prev.variables, newVar],
      rows: prev.rows.map((r) => ({
        ...r,
        values: { ...r.values, [nextName]: null },
      })),
    }));
    setSelectedVariableId(newVar.id);
  };

  const handleUpdateVariable = (
    variableId: string,
    patch: Partial<VariableDefinition>,
    renameFrom?: string
  ) => {
    commitDataset((prev) => {
      const nextVariables = prev.variables.map((v) =>
        v.id === variableId ? { ...v, ...patch } : v
      );
      let nextRows = prev.rows;
      if (renameFrom && patch.name && renameFrom !== patch.name) {
        const newName = patch.name;
        nextRows = prev.rows.map((row) => {
          const nextValues = { ...row.values };
          nextValues[newName] = nextValues[renameFrom] ?? null;
          delete nextValues[renameFrom];
          return { ...row, values: nextValues };
        });
      }
      return {
        ...prev,
        variables: nextVariables,
        rows: nextRows,
      };
    });
  };

  const handleDuplicateVariable = (variableId: string) => {
    const source = dataset.variables.find((v) => v.id === variableId);
    if (!source) return;
    const nextName = sanitizeVariableName(
      `${source.name}_copy`,
      dataset.variables.map((v) => v.name)
    );
    const clone: VariableDefinition = {
      ...source,
      id: `var_${Date.now()}`,
      name: nextName,
      valueLabels: { ...source.valueLabels },
      missingValues: {
        ...source.missingValues,
        discrete: [...source.missingValues.discrete],
        range: source.missingValues.range
          ? { ...source.missingValues.range }
          : undefined,
      },
      validation: source.validation ? { ...source.validation } : undefined,
    };
    commitDataset((prev) => ({
      ...prev,
      variables: [...prev.variables, clone],
      rows: prev.rows.map((r) => ({
        ...r,
        values: { ...r.values, [nextName]: r.values[source.name] ?? null },
      })),
    }));
    setSelectedVariableId(clone.id);
  };

  const handleMoveVariable = (variableId: string, direction: 'up' | 'down') => {
    const idx = dataset.variables.findIndex((v) => v.id === variableId);
    if (idx === -1) return;
    const targetIdx = direction === 'up' ? idx - 1 : idx + 1;
    if (targetIdx < 0 || targetIdx >= dataset.variables.length) return;
    commitDataset((prev) => {
      const next = [...prev.variables];
      const [item] = next.splice(idx, 1);
      next.splice(targetIdx, 0, item);
      return { ...prev, variables: next };
    });
  };

  const handleDeleteVariable = (variableId: string) => {
    if (dataset.variables.length <= 1) return;
    const target = dataset.variables.find((v) => v.id === variableId);
    if (!target) return;
    commitDataset((prev) => ({
      ...prev,
      variables: prev.variables.filter((v) => v.id !== variableId),
      rows: prev.rows.map((r) => {
        const nextValues = { ...r.values };
        delete nextValues[target.name];
        return { ...r, values: nextValues };
      }),
    }));
  };

  const handleSaveValueLabels = (
    variableId: string,
    nextLabels: Record<string, string>
  ) => {
    handleUpdateVariable(variableId, { valueLabels: nextLabels });
    pushToast('success', 'Value Labels updated');
  };

  const handleSaveMissingAndValidation = (
    variableId: string,
    missingValues: MissingValueSpec,
    validation: ValidationRuleSpec
  ) => {
    handleUpdateVariable(variableId, { missingValues, validation });
    pushToast('success', 'Missing values & validation rules updated');
  };

  const handleLoadPreset = (presetId: string) => {
    const preset = SAMPLE_PRESETS.find((p) => p.id === presetId);
    if (!preset) return;
    commitDataset(() => structuredClone(preset.dataset));
    setSelectedCell(null);
    setSelectedVariableId(preset.dataset.variables[0]?.id ?? null);
    setSearchQuery('');
    setIssuesOnly(false);
    pushToast('info', `Loaded ${preset.title}`);
  };

  const handleImportWorkbook = async (file: File) => {
    try {
      const imported = await importWorkbookFile(file);
      commitDataset(() => imported);
      setSelectedCell(null);
      setSelectedVariableId(imported.variables[0]?.id ?? null);
      pushToast(
        'success',
        `Imported ${file.name}`,
        `${imported.rows.length} cases · ${imported.variables.length} variables`
      );
    } catch (err) {
      pushToast(
        'error',
        'Import failed',
        err instanceof Error ? err.message : 'Unsupported or malformed workbook.'
      );
    }
  };

  const handleImportCodebook = (
    importedVars: VariableDefinition[],
    mode: 'merge' | 'replace'
  ) => {
    commitDataset((prev) => {
      if (mode === 'replace') {
        const nextRows = prev.rows.map((row) => {
          const vals: Record<string, CellValue> = {};
          importedVars.forEach((v) => {
            vals[v.name] = row.values[v.name] ?? null;
          });
          return { ...row, values: vals };
        });
        return {
          ...prev,
          variables: importedVars,
          rows: nextRows,
        };
      }

      // Merge mode: enrich matching variable names, preserve existing columns
      const importedByName = new Map(
        importedVars.map((v) => [v.name.toLowerCase(), v])
      );
      const mergedVars = prev.variables.map((v) => {
        const match = importedByName.get(v.name.toLowerCase());
        if (!match) return v;
        return {
          ...v,
          label: match.label || v.label,
          valueLabels: { ...match.valueLabels },
          missingValues: match.missingValues,
          measure: match.measure,
          role: match.role,
          validation: match.validation ?? v.validation,
        };
      });
      return { ...prev, variables: mergedVars };
    });

    pushToast(
      'success',
      mode === 'replace' ? 'Variable View replaced from Codebook' : 'Codebook merged onto matching variables'
    );
  };

  const handleJumpToVariable = (variableId: string) => {
    setSelectedVariableId(variableId);
    setDataset((prev) => ({ ...prev, activeView: 'variable' }));
  };

  const handleChangeView = (view: ActiveSheetView) => {
    setDataset((prev) => ({ ...prev, activeView: view }));
  };

  const valueLabelsTarget = valueLabelsVarId
    ? dataset.variables.find((v) => v.id === valueLabelsVarId) ?? null
    : null;
  const missingTarget = missingVarId
    ? dataset.variables.find((v) => v.id === missingVarId) ?? null
    : null;

  return (
    <div className="h-screen w-screen flex flex-col bg-slate-100 text-slate-900 overflow-hidden">
      {/* Main Frame (switches cleanly to 410px Task Pane simulation when requested) */}
      <div
        className={`flex-1 flex flex-col min-h-0 transition-all duration-150 ${
          taskPanePreview
            ? 'w-full max-w-[420px] mx-auto my-3 border border-slate-300 rounded-lg shadow-lg overflow-hidden bg-white'
            : 'w-full'
        }`}
      >
        <TopBar
          host={host}
          isSyncing={isSyncing}
          onFocusGrid={() => {
            setShowCodebookModal(false);
            setShowAddinModal(false);
            setStatsVarId(null);
          }}
          onOpenCodebook={() => setShowCodebookModal(true)}
          onOpenStats={() =>
            setStatsVarId(selectedVariableId ?? dataset.variables[0]?.id ?? null)
          }
          onOpenExcelAddin={() => setShowAddinModal(true)}
          onPrimaryAction={handlePrimaryAction}
        />

        <RibbonToolbar
          host={host}
          datasetName={dataset.datasetName}
          activeView={dataset.activeView}
          showValueLabels={dataset.showValueLabels}
          canUndo={past.length > 0}
          canRedo={future.length > 0}
          issueCount={issues.size}
          issuesOnly={issuesOnly}
          searchQuery={searchQuery}
          taskPanePreview={taskPanePreview}
          onToggleValueLabels={handleToggleValueLabels}
          onUndo={handleUndo}
          onRedo={handleRedo}
          onAddRow={handleAddRow}
          onAddVariable={handleAddVariable}
          onToggleIssuesOnly={() => setIssuesOnly((v) => !v)}
          onSearchChange={setSearchQuery}
          onLoadPreset={handleLoadPreset}
          onImportWorkbook={handleImportWorkbook}
          onToggleTaskPanePreview={() => setTaskPanePreview((p) => !p)}
          onOpenStats={() =>
            setStatsVarId(selectedVariableId ?? dataset.variables[0]?.id ?? null)
          }
          onOpenCodebook={() => setShowCodebookModal(true)}
          onOpenExcelAddin={() => setShowAddinModal(true)}
          onRefreshFromExcel={handleRefreshFromExcel}
          onRestoreRawBackup={handleRestoreRawBackup}
        />

        {/* Active View Viewport */}
        <main className="flex-1 flex flex-col min-h-0">
          {isInitializing ? (
            <div className="flex-1 p-6 space-y-3 bg-white">
              <div className="h-8 bg-slate-100 rounded animate-pulse w-1/3" />
              <div className="h-9 bg-slate-100 rounded animate-pulse" />
              <div className="h-9 bg-slate-50 rounded animate-pulse" />
              <div className="h-9 bg-slate-100 rounded animate-pulse" />
            </div>
          ) : dataset.activeView === 'data' ? (
            <DataViewGrid
              variables={dataset.variables}
              rows={dataset.rows}
              showValueLabels={dataset.showValueLabels}
              issues={issues}
              searchQuery={searchQuery}
              issuesOnly={issuesOnly}
              selectedCell={selectedCell}
              onSelectCell={(cell) => {
                setSelectedCell(cell);
                if (cell) setSelectedVariableId(cell.variableId);
              }}
              onUpdateCell={handleUpdateCell}
              onDeleteRow={handleDeleteRow}
              onAddRow={handleAddRow}
              onJumpToVariable={handleJumpToVariable}
              onInspectVariableStats={(varId) => setStatsVarId(varId)}
            />
          ) : (
            <VariableViewGrid
              variables={dataset.variables}
              searchQuery={searchQuery}
              selectedVariableId={selectedVariableId}
              onSelectVariable={setSelectedVariableId}
              onUpdateVariable={handleUpdateVariable}
              onAddVariable={handleAddVariable}
              onDuplicateVariable={handleDuplicateVariable}
              onMoveVariable={handleMoveVariable}
              onDeleteVariable={handleDeleteVariable}
              onOpenValueLabels={(varId) => setValueLabelsVarId(varId)}
              onOpenMissingModal={(varId) => setMissingVarId(varId)}
              onOpenStats={(varId) => setStatsVarId(varId)}
            />
          )}
        </main>

        <BottomViewTabs
          activeView={dataset.activeView}
          host={host}
          caseCount={dataset.rows.length}
          variableCount={dataset.variables.length}
          showValueLabels={dataset.showValueLabels}
          lastBackupTimestamp={dataset.lastBackupTimestamp}
          onChangeView={handleChangeView}
        />
      </div>

      {/* Modals & Drawers */}
      {valueLabelsTarget && (
        <ValueLabelsModal
          variable={valueLabelsTarget}
          allVariables={dataset.variables}
          onSave={handleSaveValueLabels}
          onClose={() => setValueLabelsVarId(null)}
        />
      )}

      {missingTarget && (
        <MissingAndValidationModal
          variable={missingTarget}
          onSave={handleSaveMissingAndValidation}
          onClose={() => setMissingVarId(null)}
        />
      )}

      {statsVarId && (
        <StatsInspectorDrawer
          variables={dataset.variables}
          rows={dataset.rows}
          selectedVariableId={statsVarId}
          onSelectVariable={setStatsVarId}
          onClose={() => setStatsVarId(null)}
        />
      )}

      {showCodebookModal && (
        <CodebookModal
          dataset={dataset}
          onImportCodebook={handleImportCodebook}
          onClose={() => setShowCodebookModal(false)}
        />
      )}

      {showAddinModal && (
        <ExcelAddinModal
          host={host}
          lastBackupTimestamp={dataset.lastBackupTimestamp}
          onClose={() => setShowAddinModal(false)}
        />
      )}

      {/* Toast Stack */}
      {toasts.length > 0 && (
        <div
          aria-live="polite"
          className="fixed bottom-12 right-4 z-50 flex flex-col gap-2 max-w-sm"
        >
          {toasts.map((t) => (
            <div
              key={t.id}
              className="bg-white border border-slate-300 rounded-md shadow-lg px-3.5 py-2.5 flex items-start gap-2.5 text-xs"
            >
              {t.tone === 'success' ? (
                <CheckCircle2 className="w-4 h-4 text-[#16A34A] shrink-0 mt-0.5" />
              ) : t.tone === 'error' ? (
                <AlertCircle className="w-4 h-4 text-[#DC2626] shrink-0 mt-0.5" />
              ) : (
                <Info className="w-4 h-4 text-blue-700 shrink-0 mt-0.5" />
              )}
              <div className="flex-1 min-w-0">
                <div className="font-semibold text-slate-900">{t.title}</div>
                {t.detail && (
                  <div className="text-slate-600 mt-0.5">{t.detail}</div>
                )}
              </div>
              <button
                type="button"
                onClick={() =>
                  setToasts((prev) => prev.filter((item) => item.id !== t.id))
                }
                aria-label="Dismiss notification"
                className="text-slate-400 hover:text-slate-700 cursor-pointer"
              >
                <X className="w-3.5 h-3.5" />
              </button>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}
