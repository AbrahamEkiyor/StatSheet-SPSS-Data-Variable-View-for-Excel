import type {
  CellIssue,
  CellValue,
  CodebookDocument,
  DataRow,
  DatasetState,
  FrequencyRow,
  MissingValueSpec,
  VariableDefinition,
  VariableStats,
} from '../types/spss';

export function isSystemMissing(value: CellValue | undefined): boolean {
  return value === null || value === undefined || value === '';
}

export function isNumericType(type: VariableDefinition['type']): boolean {
  return type !== 'string' && type !== 'date';
}

export function isUserMissing(
  value: CellValue | undefined,
  spec: MissingValueSpec
): boolean {
  if (isSystemMissing(value) || spec.mode === 'none') return false;
  const rawStr = String(value).trim();
  for (const d of spec.discrete) {
    if (String(d).trim() === rawStr) return true;
  }
  if (spec.mode === 'range_discrete' && spec.range) {
    const num = Number(value);
    if (!Number.isNaN(num) && num >= spec.range.low && num <= spec.range.high) {
      return true;
    }
  }
  return false;
}

export function formatRawCell(
  value: CellValue | undefined,
  variable: VariableDefinition
): string {
  if (isSystemMissing(value)) return '';
  if (variable.type === 'string' || variable.type === 'date') {
    return String(value);
  }
  const num = Number(value);
  if (Number.isNaN(num)) return String(value);

  const d = Math.max(0, Math.min(10, variable.decimals));

  switch (variable.type) {
    case 'comma':
      return num.toLocaleString('en-US', {
        minimumFractionDigits: d,
        maximumFractionDigits: d,
      });
    case 'dot':
      return num
        .toLocaleString('de-DE', {
          minimumFractionDigits: d,
          maximumFractionDigits: d,
        });
    case 'dollar':
      return `$${num.toLocaleString('en-US', {
        minimumFractionDigits: d,
        maximumFractionDigits: d,
      })}`;
    case 'custom_currency':
      return `€${num.toLocaleString('en-US', {
        minimumFractionDigits: d,
        maximumFractionDigits: d,
      })}`;
    case 'scientific':
      return num.toExponential(Math.max(1, d));
    case 'numeric':
    default:
      return d > 0 ? num.toFixed(d) : String(Math.round(num * 1e9) / 1e9);
  }
}

export function lookupValueLabel(
  value: CellValue | undefined,
  variable: VariableDefinition
): string | null {
  if (isSystemMissing(value)) return null;
  const key = String(value).trim();
  if (Object.prototype.hasOwnProperty.call(variable.valueLabels, key)) {
    return variable.valueLabels[key];
  }
  // Handle numeric equivalence like "1.0" -> "1"
  const num = Number(key);
  if (!Number.isNaN(num)) {
    const normalized = String(num);
    if (Object.prototype.hasOwnProperty.call(variable.valueLabels, normalized)) {
      return variable.valueLabels[normalized];
    }
  }
  return null;
}

export function resolveDisplayValue(
  value: CellValue | undefined,
  variable: VariableDefinition,
  showValueLabels: boolean
): { text: string; isLabeled: boolean } {
  if (isSystemMissing(value)) {
    return { text: '', isLabeled: false };
  }
  if (showValueLabels) {
    const label = lookupValueLabel(value, variable);
    if (label !== null) {
      return { text: label, isLabeled: true };
    }
  }
  return { text: formatRawCell(value, variable), isLabeled: false };
}

export function parseUserInputToRaw(
  input: string,
  variable: VariableDefinition
): CellValue {
  const trimmed = input.trim();
  if (trimmed === '') return null;

  // Allow typing a label name directly when labels exist
  for (const [code, label] of Object.entries(variable.valueLabels)) {
    if (label.toLowerCase() === trimmed.toLowerCase()) {
      if (isNumericType(variable.type) && !Number.isNaN(Number(code))) {
        return Number(code);
      }
      return code;
    }
  }

  if (isNumericType(variable.type)) {
    const cleaned = trimmed.replace(/[$,€\s]/g, '');
    const parsed = Number(cleaned);
    if (!Number.isNaN(parsed)) {
      return parsed;
    }
  }
  return trimmed;
}

export function validateVariableName(
  name: string,
  existingNames: string[],
  currentName?: string
): string | null {
  const trimmed = name.trim();
  if (!trimmed) return 'Variable name cannot be empty.';
  if (trimmed.length > 64) return 'Maximum 64 characters allowed.';
  if (!/^[A-Za-z@#][A-Za-z0-9_.$@#]*$/.test(trimmed)) {
    return 'Must start with a letter and contain no spaces or illegal symbols.';
  }
  if (trimmed.endsWith('.') || trimmed.endsWith('_')) {
    return 'Cannot end with a period or underscore.';
  }
  const lower = trimmed.toLowerCase();
  const duplicate = existingNames.some(
    (other) =>
      other.toLowerCase() === lower &&
      (!currentName || other.toLowerCase() !== currentName.toLowerCase())
  );
  if (duplicate) return 'Variable name must be unique.';
  return null;
}

export function sanitizeVariableName(raw: string, existingNames: string[]): string {
  let candidate = raw
    .trim()
    .replace(/\s+/g, '_')
    .replace(/[^A-Za-z0-9_.$@#]/g, '')
    .replace(/[._]+$/, '');
  if (!candidate || !/^[A-Za-z@#]/.test(candidate)) {
    candidate = `var_${candidate || 'col'}`;
  }
  candidate = candidate.slice(0, 60);
  let unique = candidate;
  let idx = 2;
  const lowerSet = new Set(existingNames.map((n) => n.toLowerCase()));
  while (lowerSet.has(unique.toLowerCase())) {
    unique = `${candidate}_${idx}`;
    idx += 1;
  }
  return unique;
}

export function validateDataset(
  variables: VariableDefinition[],
  rows: DataRow[]
): Map<string, CellIssue> {
  const issues = new Map<string, CellIssue>();

  rows.forEach((row, rowIndex) => {
    variables.forEach((variable) => {
      const val = row.values[variable.name];
      const cellKey = `${row.id}::${variable.id}`;

      if (isSystemMissing(val)) {
        if (variable.validation?.required) {
          issues.set(cellKey, {
            rowIndex,
            rowId: row.id,
            variableId: variable.id,
            variableName: variable.name,
            kind: 'validation_required',
            message: `Required value is missing for "${variable.name}".`,
          });
        }
        return;
      }

      if (isUserMissing(val, variable.missingValues)) {
        const label = lookupValueLabel(val, variable);
        issues.set(cellKey, {
          rowIndex,
          rowId: row.id,
          variableId: variable.id,
          variableName: variable.name,
          kind: 'missing_user',
          message: `User-missing code (${val}${label ? ` · ${label}` : ''}) excluded from valid statistics.`,
        });
        return;
      }

      if (isNumericType(variable.type)) {
        const num = Number(val);
        if (Number.isNaN(num)) {
          issues.set(cellKey, {
            rowIndex,
            rowId: row.id,
            variableId: variable.id,
            variableName: variable.name,
            kind: 'validation_type',
            message: `Expected numeric value for "${variable.name}", got "${val}".`,
          });
          return;
        }
        if (
          variable.validation?.min !== undefined &&
          num < variable.validation.min
        ) {
          issues.set(cellKey, {
            rowIndex,
            rowId: row.id,
            variableId: variable.id,
            variableName: variable.name,
            kind: 'validation_range',
            message: `Value ${num} is below minimum (${variable.validation.min}).`,
          });
          return;
        }
        if (
          variable.validation?.max !== undefined &&
          num > variable.validation.max
        ) {
          issues.set(cellKey, {
            rowIndex,
            rowId: row.id,
            variableId: variable.id,
            variableName: variable.name,
            kind: 'validation_range',
            message: `Value ${num} exceeds maximum (${variable.validation.max}).`,
          });
          return;
        }
      }

      const labelCount = Object.keys(variable.valueLabels).length;
      if (variable.validation?.restrictToLabels && labelCount > 0) {
        if (lookupValueLabel(val, variable) === null) {
          issues.set(cellKey, {
            rowIndex,
            rowId: row.id,
            variableId: variable.id,
            variableName: variable.name,
            kind: 'validation_uncoded',
            message: `Value "${val}" has no defined Value Label in "${variable.name}".`,
          });
        }
      }
    });
  });

  return issues;
}

export function formatMissingSummary(spec: MissingValueSpec): string {
  if (spec.mode === 'none') return 'None';
  if (spec.mode === 'discrete') {
    return spec.discrete.length > 0 ? spec.discrete.join(', ') : 'None';
  }
  const rangePart = spec.range ? `${spec.range.low} thru ${spec.range.high}` : '';
  const discPart = spec.discrete.length > 0 ? `, ${spec.discrete[0]}` : '';
  return `${rangePart}${discPart}` || 'None';
}

export function formatValueLabelsSummary(
  valueLabels: Record<string, string>
): string {
  const entries = Object.entries(valueLabels);
  if (entries.length === 0) return 'None';
  const [firstCode, firstLabel] = entries[0];
  if (entries.length === 1) return `{${firstCode}, ${firstLabel}}`;
  return `{${firstCode}, ${firstLabel}}… (+${entries.length - 1})`;
}

export function computeVariableStats(
  variable: VariableDefinition,
  rows: DataRow[]
): VariableStats {
  const totalN = rows.length;
  let validN = 0;
  let userMissingN = 0;
  let systemMissingN = 0;

  const numericValues: number[] = [];
  const freqMap = new Map<
    string,
    { rawKey: string; label: string; isMissing: boolean; count: number }
  >();

  for (const row of rows) {
    const val = row.values[variable.name];
    if (isSystemMissing(val)) {
      systemMissingN += 1;
      continue;
    }

    const userMissing = isUserMissing(val, variable.missingValues);
    const key = String(val).trim();
    const mappedLabel = lookupValueLabel(val, variable);
    const displayLabel = mappedLabel ?? key;

    if (userMissing) {
      userMissingN += 1;
    } else {
      validN += 1;
      if (isNumericType(variable.type)) {
        const num = Number(val);
        if (!Number.isNaN(num)) {
          numericValues.push(num);
        }
      }
    }

    const existing = freqMap.get(key);
    if (existing) {
      existing.count += 1;
    } else {
      freqMap.set(key, {
        rawKey: key,
        label: displayLabel,
        isMissing: userMissing,
        count: 1,
      });
    }
  }

  let mean: number | null = null;
  let median: number | null = null;
  let stdDev: number | null = null;
  let min: number | null = null;
  let max: number | null = null;

  if (numericValues.length > 0) {
    const sorted = [...numericValues].sort((a, b) => a - b);
    const n = sorted.length;
    min = sorted[0];
    max = sorted[n - 1];
    const sum = sorted.reduce((acc, v) => acc + v, 0);
    mean = sum / n;
    median =
      n % 2 === 1
        ? sorted[(n - 1) / 2]
        : (sorted[n / 2 - 1] + sorted[n / 2]) / 2;
    if (n > 1) {
      const variance =
        sorted.reduce((acc, v) => acc + (v - mean!) ** 2, 0) / (n - 1);
      stdDev = Math.sqrt(variance);
    } else {
      stdDev = 0;
    }
  }

  // Sort frequency rows: valid first (numeric ascending or alpha), then user-missing
  const entries = Array.from(freqMap.values()).sort((a, b) => {
    if (a.isMissing !== b.isMissing) return a.isMissing ? 1 : -1;
    const na = Number(a.rawKey);
    const nb = Number(b.rawKey);
    if (!Number.isNaN(na) && !Number.isNaN(nb)) return na - nb;
    return a.rawKey.localeCompare(b.rawKey);
  });

  let runningValidPct = 0;
  const frequencies: FrequencyRow[] = entries.map((entry) => {
    const percent = totalN > 0 ? (entry.count / totalN) * 100 : 0;
    let validPercent: number | null = null;
    let cumulativePercent: number | null = null;
    if (!entry.isMissing && validN > 0) {
      validPercent = (entry.count / validN) * 100;
      runningValidPct = Math.min(100, runningValidPct + validPercent);
      cumulativePercent = runningValidPct;
    }
    return {
      rawKey: entry.rawKey,
      label: entry.label,
      isMissing: entry.isMissing,
      count: entry.count,
      percent,
      validPercent,
      cumulativePercent,
    };
  });

  return {
    variableId: variable.id,
    totalN,
    validN,
    userMissingN,
    systemMissingN,
    mean,
    median,
    stdDev,
    min,
    max,
    frequencies,
  };
}

export function buildCodebookDocument(dataset: DatasetState): CodebookDocument {
  return {
    format: 'statsheet-codebook-v1',
    exportedAt: new Date().toISOString(),
    datasetName: dataset.datasetName,
    variables: dataset.variables,
  };
}

export function buildDataDictionaryCsv(variables: VariableDefinition[]): string {
  const escape = (s: string) => `"${s.replace(/"/g, '""')}"`;
  const headers = [
    'Position',
    'Name',
    'Type',
    'Width',
    'Decimals',
    'Variable Label',
    'Value Labels',
    'Missing Values',
    'Measure',
    'Role',
    'Validation Bounds',
  ];
  const lines = [headers.map(escape).join(',')];

  variables.forEach((v, idx) => {
    const valStr = Object.entries(v.valueLabels)
      .map(([k, lbl]) => `${k}=${lbl}`)
      .join('; ');
    const bounds: string[] = [];
    if (v.validation?.min !== undefined) bounds.push(`Min: ${v.validation.min}`);
    if (v.validation?.max !== undefined) bounds.push(`Max: ${v.validation.max}`);
    if (v.validation?.restrictToLabels) bounds.push('Restrict to Value Labels');

    const row = [
      String(idx + 1),
      v.name,
      v.type,
      String(v.width),
      String(v.decimals),
      v.label,
      valStr || 'None',
      formatMissingSummary(v.missingValues),
      v.measure,
      v.role,
      bounds.join(' | ') || 'None',
    ];
    lines.push(row.map(escape).join(','));
  });

  return lines.join('\n');
}

export function generateOfficeManifestXml(baseUrl: string): string {
  const cleanUrl = baseUrl.replace(/\/$/, '');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
  xsi:type="TaskPaneApp">
  <Id>c8b4f210-6a9c-4e82-9a14-5d037e9b2401</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>StatSheet</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="StatSheet — SPSS Data &amp; Variable View"/>
  <Description DefaultValue="Manage SPSS-style Variable View metadata, missing values, validation rules, and reversible Value Labels directly inside Excel."/>
  <IconUrl DefaultValue="${cleanUrl}/icon-32.png"/>
  <HighResolutionIconUrl DefaultValue="${cleanUrl}/icon-80.png"/>
  <SupportUrl DefaultValue="${cleanUrl}"/>
  <AppDomains>
    <AppDomain>${cleanUrl}</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Workbook"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="${cleanUrl}/"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Workbook">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title"/>
            <Description resid="GetStarted.Description"/>
            <LearnMoreUrl resid="GetStarted.LearnMoreUrl"/>
          </GetStarted>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="StatSheet.Group">
                <Label resid="Group.Label"/>
                <Icon>
                  <bt:Image size="16" resid="Icon.16"/>
                  <bt:Image size="32" resid="Icon.32"/>
                  <bt:Image size="80" resid="Icon.80"/>
                </Icon>
                <Control xsi:type="Button" id="StatSheet.OpenPaneButton">
                  <Label resid="Button.Label"/>
                  <Supertip>
                    <Title resid="Button.Label"/>
                    <Description resid="Button.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16"/>
                    <bt:Image size="32" resid="Icon.32"/>
                    <bt:Image size="80" resid="Icon.80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>StatSheetTaskpane</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url"/>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16" DefaultValue="${cleanUrl}/icon-16.png"/>
        <bt:Image id="Icon.32" DefaultValue="${cleanUrl}/icon-32.png"/>
        <bt:Image id="Icon.80" DefaultValue="${cleanUrl}/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="GetStarted.LearnMoreUrl" DefaultValue="${cleanUrl}"/>
        <bt:Url id="Taskpane.Url" DefaultValue="${cleanUrl}/"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="GetStarted.Title" DefaultValue="StatSheet for Excel"/>
        <bt:String id="Group.Label" DefaultValue="SPSS View"/>
        <bt:String id="Button.Label" DefaultValue="Data &amp; Variable View"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="GetStarted.Description" DefaultValue="Open SPSS Data View, Variable View, and Value Labels toggle inside Excel."/>
        <bt:String id="Button.Tooltip" DefaultValue="Open the StatSheet SPSS Data &amp; Variable View task pane."/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>`;
}
