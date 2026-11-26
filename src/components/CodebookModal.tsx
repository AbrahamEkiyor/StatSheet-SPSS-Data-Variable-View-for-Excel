import { Download, FileJson, FileSpreadsheet, Upload, X } from 'lucide-react';
import { useRef, useState } from 'react';
import { exportDatasetToCsv } from '../bridge/workbookBridge';
import {
  buildCodebookDocument,
  buildDataDictionaryCsv,
  formatMissingSummary,
} from '../engine/spssEngine';
import type { CodebookDocument, DatasetState, VariableDefinition } from '../types/spss';

interface CodebookModalProps {
  dataset: DatasetState;
  onImportCodebook: (variables: VariableDefinition[], mode: 'merge' | 'replace') => void;
  onClose: () => void;
}

function triggerDownload(filename: string, content: string, mime: string) {
  const blob = new Blob([content], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

export function CodebookModal({
  dataset,
  onImportCodebook,
  onClose,
}: CodebookModalProps) {
  const fileRef = useRef<HTMLInputElement>(null);
  const [importMode, setImportMode] = useState<'merge' | 'replace'>('merge');
  const [error, setError] = useState<string | null>(null);

  const safeName = (dataset.datasetName || 'dataset').replace(/[^A-Za-z0-9_-]/g, '_');

  const handleExportJson = () => {
    const doc = buildCodebookDocument(dataset);
    triggerDownload(
      `${safeName}_codebook.json`,
      JSON.stringify(doc, null, 2),
      'application/json'
    );
  };

  const handleExportDictionaryCsv = () => {
    const csv = buildDataDictionaryCsv(dataset.variables);
    triggerDownload(`${safeName}_dictionary.csv`, csv, 'text/csv;charset=utf-8');
  };

  const handleExportDataCsv = (withLabels: boolean) => {
    const csv = exportDatasetToCsv(dataset, withLabels);
    triggerDownload(
      `${safeName}_${withLabels ? 'labels' : 'raw'}.csv`,
      csv,
      'text/csv;charset=utf-8'
    );
  };

  const handleFileChange = async (file: File) => {
    try {
      setError(null);
      const text = await file.text();
      const parsed = JSON.parse(text) as Partial<CodebookDocument>;
      if (!parsed || !Array.isArray(parsed.variables)) {
        throw new Error('Invalid codebook JSON format (missing variables array).');
      }
      onImportCodebook(parsed.variables, importMode);
      onClose();
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to parse codebook JSON.');
    }
  };

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4"
      role="dialog"
      aria-modal="true"
      aria-labelledby="codebook-modal-title"
    >
      <div className="bg-white border border-slate-300 rounded-lg shadow-xl w-full max-w-3xl flex flex-col max-h-[90vh]">
        <div className="px-5 py-3.5 border-b border-slate-200 flex items-center justify-between">
          <div>
            <h2
              id="codebook-modal-title"
              className="font-display text-sm font-semibold text-slate-900"
            >
              Codebook &amp; Data Dictionary Manager
            </h2>
            <p className="text-xs text-slate-500 mt-0.5">
              Export portable JSON codebooks, CSV data dictionaries, or apply an existing schema onto this workbook.
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

        <div className="p-5 overflow-y-auto space-y-5">
          {/* Export / Import Actions */}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div className="p-4 border border-slate-200 rounded-md bg-slate-50/70 space-y-3">
              <h3 className="font-display text-xs font-semibold text-slate-900">
                Export Codebook &amp; Data
              </h3>
              <div className="flex flex-wrap gap-2">
                <button
                  type="button"
                  onClick={handleExportJson}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-white bg-blue-700 hover:bg-blue-800 rounded-md cursor-pointer"
                >
                  <FileJson className="w-3.5 h-3.5" />
                  <span>JSON Codebook</span>
                </button>
                <button
                  type="button"
                  onClick={handleExportDictionaryCsv}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-300 hover:bg-slate-50 rounded-md cursor-pointer"
                >
                  <Download className="w-3.5 h-3.5" />
                  <span>Dictionary (.csv)</span>
                </button>
              </div>
              <div className="pt-2 border-t border-slate-200 flex flex-wrap gap-2">
                <button
                  type="button"
                  onClick={() => handleExportDataCsv(false)}
                  className="inline-flex items-center gap-1.5 px-2.5 py-1 text-[11px] font-medium text-slate-700 bg-white border border-slate-300 hover:bg-slate-50 rounded cursor-pointer"
                >
                  <FileSpreadsheet className="w-3 h-3" />
                  <span>Export Raw Codes (.csv)</span>
                </button>
                <button
                  type="button"
                  onClick={() => handleExportDataCsv(true)}
                  className="inline-flex items-center gap-1.5 px-2.5 py-1 text-[11px] font-medium text-slate-700 bg-white border border-slate-300 hover:bg-slate-50 rounded cursor-pointer"
                >
                  <FileSpreadsheet className="w-3 h-3" />
                  <span>Export Value Labels (.csv)</span>
                </button>
              </div>
            </div>

            <div className="p-4 border border-slate-200 rounded-md bg-slate-50/70 space-y-3">
              <h3 className="font-display text-xs font-semibold text-slate-900">
                Import JSON Codebook
              </h3>
              <div className="flex items-center gap-4 text-xs text-slate-700">
                <label className="flex items-center gap-1.5 cursor-pointer">
                  <input
                    type="radio"
                    name="cbMode"
                    checked={importMode === 'merge'}
                    onChange={() => setImportMode('merge')}
                  />
                  <span>Match &amp; enrich columns</span>
                </label>
                <label className="flex items-center gap-1.5 cursor-pointer">
                  <input
                    type="radio"
                    name="cbMode"
                    checked={importMode === 'replace'}
                    onChange={() => setImportMode('replace')}
                  />
                  <span>Replace Variable View</span>
                </label>
              </div>
              <input
                ref={fileRef}
                type="file"
                accept=".json,application/json"
                className="hidden"
                onChange={(e) => {
                  const file = e.target.files?.[0];
                  if (file) {
                    handleFileChange(file);
                    e.target.value = '';
                  }
                }}
              />
              <button
                type="button"
                onClick={() => fileRef.current?.click()}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-slate-800 bg-white border border-slate-300 hover:bg-slate-100 rounded-md cursor-pointer"
              >
                <Upload className="w-3.5 h-3.5" />
                <span>Choose Codebook .json</span>
              </button>
              {error && <p className="text-xs text-red-700">{error}</p>}
            </div>
          </div>

          {/* Codebook Dictionary Preview */}
          <div className="border border-slate-200 rounded-md overflow-hidden">
            <div className="px-3.5 py-2 bg-slate-100 border-b border-slate-200 text-xs font-semibold text-slate-700">
              Current Schema Dictionary ({dataset.variables.length} variables)
            </div>
            <div className="max-h-72 overflow-auto">
              <table className="w-full border-collapse text-xs">
                <thead className="bg-slate-50 border-b border-slate-200 text-slate-600 font-semibold sticky top-0">
                  <tr>
                    <th className="px-3 py-2 text-left">Variable</th>
                    <th className="px-3 py-2 text-left">Label</th>
                    <th className="px-3 py-2 text-left">Measure</th>
                    <th className="px-3 py-2 text-left">Value Labels</th>
                    <th className="px-3 py-2 text-left">Missing</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200">
                  {dataset.variables.map((v) => (
                    <tr key={v.id} className="hover:bg-slate-50">
                      <td className="px-3 py-2 font-mono font-semibold text-slate-900 align-top">
                        {v.name}
                      </td>
                      <td className="px-3 py-2 text-slate-700 align-top">
                        {v.label || <span className="text-slate-400">—</span>}
                      </td>
                      <td className="px-3 py-2 capitalize text-slate-700 align-top">
                        {v.measure}
                      </td>
                      <td className="px-3 py-2 text-slate-700 align-top">
                        {Object.keys(v.valueLabels).length === 0 ? (
                          <span className="text-slate-400">None</span>
                        ) : (
                          <div className="space-y-0.5">
                            {Object.entries(v.valueLabels).map(([k, lbl]) => (
                              <div key={k} className="font-mono text-[11px]">
                                {k} = <span className="font-sans">{lbl}</span>
                              </div>
                            ))}
                          </div>
                        )}
                      </td>
                      <td className="px-3 py-2 font-mono text-slate-700 align-top">
                        {formatMissingSummary(v.missingValues)}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
