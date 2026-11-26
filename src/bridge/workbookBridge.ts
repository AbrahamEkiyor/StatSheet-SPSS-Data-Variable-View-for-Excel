/// <reference types="office-js" />
import * as XLSX from 'xlsx';

declare global {
  interface Window {
    Office?: typeof Office;
    Excel?: typeof Excel;
  }
}
import {
  isSystemMissing,
  lookupValueLabel,
  sanitizeVariableName,
} from '../engine/spssEngine';
import type {
  CellValue,
  DataRow,
  DatasetState,
  HostEnvironment,
  VariableDefinition,
} from '../types/spss';

const SIMULATOR_STORAGE_KEY = 'statsheet_spss_simulator_v1';
const METADATA_SHEET_NAME = '_SPSS_Metadata';
const BACKUP_SHEET_NAME = '_SPSS_RawBackup';

export interface HostDetectionResult {
  host: HostEnvironment;
  excelVersion?: string;
}

export function isOfficeHostContext(): boolean {
  if (typeof window === 'undefined') return false;
  if (window.Office?.context?.host) return true;
  const search = window.location.search || '';
  const hash = window.location.hash || '';
  if (search.includes('_host_Info=') || hash.includes('_host_Info=')) {
    return true;
  }
  try {
    if (window.sessionStorage?.getItem('hostInfoValue')) return true;
  } catch {
    // Ignore storage access errors in sandboxed iframes
  }
  return false;
}

function loadOfficeJsScript(): Promise<void> {
  return new Promise((resolve, reject) => {
    if (window.Office) {
      resolve();
      return;
    }
    const script = document.createElement('script');
    script.src = 'https://appsforoffice.microsoft.com/lib/1/hosted/office.js';
    script.type = 'text/javascript';
    script.onload = () => resolve();
    script.onerror = () => reject(new Error('Failed to load Office.js'));
    document.head.appendChild(script);
  });
}

export async function detectOfficeHost(): Promise<HostDetectionResult> {
  if (!isOfficeHostContext()) {
    return { host: 'simulator' };
  }

  try {
    await loadOfficeJsScript();
  } catch {
    return { host: 'simulator' };
  }

  const officeApi = typeof window !== 'undefined' ? window.Office : undefined;
  if (!officeApi) {
    return { host: 'simulator' };
  }

  return new Promise((resolve) => {
    const timer = window.setTimeout(() => {
      resolve({ host: 'simulator' });
    }, 1500);

    try {
      officeApi.onReady((info) => {
        window.clearTimeout(timer);
        if (info?.host === officeApi.HostType.Excel) {
          resolve({ host: 'excel' });
        } else {
          resolve({ host: 'simulator' });
        }
      });
    } catch {
      window.clearTimeout(timer);
      resolve({ host: 'simulator' });
    }
  });
}

export function loadSimulatorState(): DatasetState | null {
  try {
    const raw = localStorage.getItem(SIMULATOR_STORAGE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw) as DatasetState;
    if (
      parsed &&
      Array.isArray(parsed.variables) &&
      Array.isArray(parsed.rows)
    ) {
      return parsed;
    }
    return null;
  } catch {
    return null;
  }
}

export function saveSimulatorState(state: DatasetState): void {
  try {
    localStorage.setItem(SIMULATOR_STORAGE_KEY, JSON.stringify(state));
  } catch {
    // Ignore storage quota errors
  }
}

/**
 * Exports dataset to an .xlsx workbook including a hidden-friendly _SPSS_Metadata sheet
 * and applies Value Labels if requested while keeping _SPSS_RawBackup for reversibility.
 */
export function exportDatasetToXlsx(
  dataset: DatasetState,
  exportWithLabels: boolean
): void {
  const wb = XLSX.utils.book_new();

  const headers = dataset.variables.map((v) => v.name);
  const dataMatrix: Array<Array<string | number | null>> = [headers];

  dataset.rows.forEach((row) => {
    const line = dataset.variables.map((v) => {
      const raw = row.values[v.name];
      if (isSystemMissing(raw)) return null;
      if (exportWithLabels) {
        const lbl = lookupValueLabel(raw, v);
        if (lbl !== null) return lbl;
      }
      return raw ?? null;
    });
    dataMatrix.push(line);
  });

  const dataSheet = XLSX.utils.aoa_to_sheet(dataMatrix);
  XLSX.utils.book_append_sheet(wb, dataSheet, 'Data');

  // Include raw backup sheet so raw numeric codes are preserved inside the Excel file
  const rawMatrix: Array<Array<string | number | null>> = [headers];
  dataset.rows.forEach((row) => {
    rawMatrix.push(
      dataset.variables.map((v) => {
        const raw = row.values[v.name];
        return isSystemMissing(raw) ? null : raw ?? null;
      })
    );
  });
  const backupSheet = XLSX.utils.aoa_to_sheet(rawMatrix);
  XLSX.utils.book_append_sheet(wb, backupSheet, BACKUP_SHEET_NAME);

  // Include SPSS variable dictionary metadata sheet
  const metaRows = dataset.variables.map((v, idx) => ({
    Position: idx + 1,
    Name: v.name,
    Type: v.type,
    Width: v.width,
    Decimals: v.decimals,
    Label: v.label,
    ValueLabelsJSON: JSON.stringify(v.valueLabels),
    MissingJSON: JSON.stringify(v.missingValues),
    Columns: v.columns,
    Align: v.align,
    Measure: v.measure,
    Role: v.role,
    ValidationJSON: JSON.stringify(v.validation ?? {}),
    ClassBinsJSON: v.classBins ? JSON.stringify(v.classBins) : '',
  }));
  const metaSheet = XLSX.utils.json_to_sheet(metaRows);
  XLSX.utils.book_append_sheet(wb, metaSheet, METADATA_SHEET_NAME);

  // Mark backup & metadata sheets hidden in workbook structure when supported
  wb.Workbook = {
    Sheets: [
      { name: 'Data', Hidden: 0 },
      { name: BACKUP_SHEET_NAME, Hidden: 1 },
      { name: METADATA_SHEET_NAME, Hidden: 1 },
    ],
  };

  const safeName = (dataset.datasetName || 'StatSheet_Export').replace(
    /[^A-Za-z0-9_-]/g,
    '_'
  );
  XLSX.writeFile(wb, `${safeName}.xlsx`);
}

export function exportDatasetToCsv(
  dataset: DatasetState,
  exportWithLabels: boolean
): string {
  const escape = (val: CellValue) => {
    if (isSystemMissing(val)) return '';
    const str = String(val);
    if (/[",\n]/.test(str)) {
      return `"${str.replace(/"/g, '""')}"`;
    }
    return str;
  };
  const headerLine = dataset.variables.map((v) => escape(v.name)).join(',');
  const lines = [headerLine];
  dataset.rows.forEach((row) => {
    const cols = dataset.variables.map((v) => {
      const raw = row.values[v.name];
      if (isSystemMissing(raw)) return '';
      if (exportWithLabels) {
        const lbl = lookupValueLabel(raw, v);
        if (lbl !== null) return escape(lbl);
      }
      return escape(raw);
    });
    lines.push(cols.join(','));
  });
  return lines.join('\n');
}

export async function importWorkbookFile(file: File): Promise<DatasetState> {
  const buffer = await file.arrayBuffer();
  const wb = XLSX.read(buffer, { type: 'array' });
  const primarySheetName =
    wb.SheetNames.find(
      (n) => n !== METADATA_SHEET_NAME && n !== BACKUP_SHEET_NAME
    ) ?? wb.SheetNames[0];

  if (!primarySheetName) {
    throw new Error('No worksheets found in the uploaded file.');
  }

  // Prefer raw backup sheet if present so numeric codes are restored cleanly
  const sourceSheetName = wb.SheetNames.includes(BACKUP_SHEET_NAME)
    ? BACKUP_SHEET_NAME
    : primarySheetName;

  const sheet = wb.Sheets[sourceSheetName];
  const aoa = XLSX.utils.sheet_to_json<Array<string | number | null>>(sheet, {
    header: 1,
    defval: null,
  });

  if (aoa.length === 0) {
    throw new Error('The selected worksheet is empty.');
  }

  const rawHeaders = (aoa[0] ?? []).map((h, i) =>
    h !== null && h !== undefined && String(h).trim() !== ''
      ? String(h).trim()
      : `var_${i + 1}`
  );

  // Check if _SPSS_Metadata sheet exists for lossless round-trip
  let savedVarsMap = new Map<string, VariableDefinition>();
  if (wb.SheetNames.includes(METADATA_SHEET_NAME)) {
    const metaJson = XLSX.utils.sheet_to_json<Record<string, unknown>>(
      wb.Sheets[METADATA_SHEET_NAME]
    );
    metaJson.forEach((item, index) => {
      const name = String(item.Name ?? `var_${index + 1}`);
      try {
        savedVarsMap.set(name, {
          id: `imp_${index}_${Date.now()}`,
          name,
          type: (item.Type as VariableDefinition['type']) || 'numeric',
          width: Number(item.Width) || 8,
          decimals: Number(item.Decimals) || 0,
          label: String(item.Label ?? ''),
          valueLabels: item.ValueLabelsJSON
            ? JSON.parse(String(item.ValueLabelsJSON))
            : {},
          missingValues: item.MissingJSON
            ? JSON.parse(String(item.MissingJSON))
            : { mode: 'none', discrete: [] },
          columns: Number(item.Columns) || 128,
          align: (item.Align as VariableDefinition['align']) || 'right',
          measure: (item.Measure as VariableDefinition['measure']) || 'scale',
          role: (item.Role as VariableDefinition['role']) || 'input',
          validation: item.ValidationJSON
            ? JSON.parse(String(item.ValidationJSON))
            : undefined,
          classBins: item.ClassBinsJSON
            ? JSON.parse(String(item.ClassBinsJSON))
            : undefined,
        });
      } catch {
        // Fallback to inferred variable if metadata row is malformed
      }
    });
  }

  const usedNames: string[] = [];
  const variables: VariableDefinition[] = rawHeaders.map((hdr, colIdx) => {
    const existingMeta = savedVarsMap.get(hdr);
    if (existingMeta) {
      usedNames.push(existingMeta.name);
      return existingMeta;
    }

    const cleanName = sanitizeVariableName(hdr, usedNames);
    usedNames.push(cleanName);

    // Infer numeric vs string from column values
    let hasString = false;
    let maxDecimals = 0;
    for (let r = 1; r < aoa.length; r += 1) {
      const cell = aoa[r]?.[colIdx];
      if (cell === null || cell === undefined || cell === '') continue;
      const num = Number(cell);
      if (Number.isNaN(num)) {
        hasString = true;
        break;
      } else {
        const parts = String(cell).split('.');
        if (parts[1]) {
          maxDecimals = Math.max(maxDecimals, Math.min(4, parts[1].length));
        }
      }
    }

    return {
      id: `var_${colIdx}_${Date.now()}`,
      name: cleanName,
      type: hasString ? 'string' : 'numeric',
      width: hasString ? 24 : 8,
      decimals: hasString ? 0 : maxDecimals,
      label: hdr !== cleanName ? hdr : '',
      valueLabels: {},
      missingValues: { mode: 'none', discrete: [] },
      columns: 128,
      align: hasString ? 'left' : 'right',
      measure: hasString ? 'nominal' : 'scale',
      role: 'input',
    };
  });

  const rows: DataRow[] = aoa.slice(1).map((line, rIdx) => {
    const values: Record<string, CellValue> = {};
    variables.forEach((v, cIdx) => {
      const rawCell = line?.[cIdx];
      if (rawCell === undefined || rawCell === null || rawCell === '') {
        values[v.name] = null;
      } else if (v.type !== 'string' && v.type !== 'date') {
        const num = Number(rawCell);
        values[v.name] = Number.isNaN(num) ? String(rawCell) : num;
      } else {
        values[v.name] = String(rawCell);
      }
    });
    return {
      id: `imp_row_${rIdx}_${Date.now()}`,
      values,
    };
  });

  return {
    datasetName: file.name.replace(/\.(xlsx|xls|csv)$/i, '') || 'Imported_Sheet',
    variables,
    rows,
    showValueLabels: false,
    activeView: 'data',
    lastBackupTimestamp: new Date().toISOString(),
  };
}

/**
 * Live Excel Office.js Bridge:
 * Reads active worksheet into StatSheet, and performs reversible cell sync
 * with hidden _SPSS_RawBackup & _SPSS_Metadata worksheets.
 */
export async function readFromLiveExcel(): Promise<DatasetState> {
  const excelApi = window.Excel;
  if (!excelApi) {
    throw new Error('Excel Office.js runtime is not available.');
  }
  return excelApi.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load('name');

    // Prefer _SPSS_RawBackup if present so we always load true raw codes
    const worksheets = context.workbook.worksheets;
    worksheets.load('items/name');
    await context.sync();

    const hasMeta = worksheets.items.some((s) => s.name === METADATA_SHEET_NAME);
    const hasBackup = worksheets.items.some((s) => s.name === BACKUP_SHEET_NAME);

    const dataSourceSheet = hasBackup
      ? worksheets.getItem(BACKUP_SHEET_NAME)
      : sheet;
    const usedRange = dataSourceSheet.getUsedRangeOrNullObject();
    usedRange.load('values');

    let metaRange: Excel.Range | null = null;
    if (hasMeta) {
      const metaSheet = worksheets.getItem(METADATA_SHEET_NAME);
      metaRange = metaSheet.getUsedRangeOrNullObject();
      metaRange.load('values');
    }

    await context.sync();

    const rawValues = usedRange.isNullObject ? [] : (usedRange.values as unknown[][]);
    if (rawValues.length === 0) {
      return {
        datasetName: sheet.name,
        variables: [],
        rows: [],
        showValueLabels: false,
        activeView: 'variable',
        lastBackupTimestamp: null,
      };
    }

    let savedVars: VariableDefinition[] = [];
    if (metaRange && !metaRange.isNullObject && metaRange.values.length > 0) {
      try {
        const jsonPayload = String(metaRange.values[0][0] ?? '');
        savedVars = JSON.parse(jsonPayload) as VariableDefinition[];
      } catch {
        savedVars = [];
      }
    }

    const headers = rawValues[0].map((h, idx) =>
      String(h ?? '').trim() || `var_${idx + 1}`
    );
    const usedNames: string[] = [];

    const variables: VariableDefinition[] = headers.map((hdr, idx) => {
      const matched = savedVars.find((sv) => sv.name === hdr);
      if (matched) {
        usedNames.push(matched.name);
        return matched;
      }
      const cleanName = sanitizeVariableName(hdr, usedNames);
      usedNames.push(cleanName);
      return {
        id: `xl_var_${idx}_${Date.now()}`,
        name: cleanName,
        type: 'numeric',
        width: 8,
        decimals: 0,
        label: hdr !== cleanName ? hdr : '',
        valueLabels: {},
        missingValues: { mode: 'none', discrete: [] },
        columns: 128,
        align: 'right',
        measure: 'scale',
        role: 'input',
      };
    });

    const rows: DataRow[] = rawValues.slice(1).map((line, rIdx) => {
      const values: Record<string, CellValue> = {};
      variables.forEach((v, cIdx) => {
        const cell = line[cIdx];
        values[v.name] =
          cell === '' || cell === null || cell === undefined
            ? null
            : typeof cell === 'number' || typeof cell === 'string'
              ? cell
              : String(cell);
      });
      return { id: `xl_row_${rIdx}`, values };
    });

    return {
      datasetName: sheet.name,
      variables,
      rows,
      showValueLabels: false,
      activeView: 'data',
      lastBackupTimestamp: hasBackup ? new Date().toISOString() : null,
    };
  });
}

export async function syncToLiveExcel(
  dataset: DatasetState,
  applyLabels: boolean
): Promise<string> {
  const excelApi = window.Excel;
  if (!excelApi) {
    throw new Error('Excel Office.js runtime is not available.');
  }
  return excelApi.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    worksheets.load('items/name');
    const activeSheet = worksheets.getActiveWorksheet();
    activeSheet.load('name');
    await context.sync();

    // 1. Ensure hidden _SPSS_RawBackup exists and stores the exact raw numeric/text codes
    let backupSheet = worksheets.items.find((s) => s.name === BACKUP_SHEET_NAME);
    if (!backupSheet) {
      backupSheet = worksheets.add(BACKUP_SHEET_NAME);
    }
    backupSheet.visibility = excelApi.SheetVisibility.hidden;

    const headers = dataset.variables.map((v) => v.name);
    const rawMatrix = [
      headers,
      ...dataset.rows.map((row) =>
        dataset.variables.map((v) => {
          const val = row.values[v.name];
          return isSystemMissing(val) ? '' : val;
        })
      ),
    ];

    if (rawMatrix.length > 0 && headers.length > 0) {
      const backupRange = backupSheet.getRangeByIndexes(
        0,
        0,
        rawMatrix.length,
        headers.length
      );
      backupRange.values = rawMatrix;
    }

    // 2. Persist full 11-column Variable View metadata in hidden _SPSS_Metadata sheet
    let metaSheet = worksheets.items.find((s) => s.name === METADATA_SHEET_NAME);
    if (!metaSheet) {
      metaSheet = worksheets.add(METADATA_SHEET_NAME);
    }
    metaSheet.visibility = excelApi.SheetVisibility.hidden;
    metaSheet.getRange('A1').values = [[JSON.stringify(dataset.variables)]];

    // 3. Write either Value Labels or restored Raw Codes to the active worksheet
    const displayMatrix = [
      headers,
      ...dataset.rows.map((row) =>
        dataset.variables.map((v) => {
          const val = row.values[v.name];
          if (isSystemMissing(val)) return '';
          if (applyLabels) {
            const lbl = lookupValueLabel(val, v);
            if (lbl !== null) return lbl;
          }
          return val;
        })
      ),
    ];

    if (displayMatrix.length > 0 && headers.length > 0) {
      const targetRange = activeSheet.getRangeByIndexes(
        0,
        0,
        displayMatrix.length,
        headers.length
      );
      targetRange.values = displayMatrix;

      // Apply Excel Data Validation dropdowns for coded variables when labels are active
      dataset.variables.forEach((v, colIdx) => {
        const labelEntries = Object.entries(v.valueLabels);
        if (dataset.rows.length > 0 && labelEntries.length > 0 && v.validation?.restrictToLabels) {
          const colRange = activeSheet.getRangeByIndexes(
            1,
            colIdx,
            dataset.rows.length,
            1
          );
          const listValues = applyLabels
            ? labelEntries.map(([, label]) => label)
            : labelEntries.map(([code]) => code);
          colRange.dataValidation.rule = {
            list: {
              inCellDropDown: true,
              source: listValues.join(','),
            },
          };
        }
      });
    }

    await context.sync();
    return new Date().toISOString();
  });
}
