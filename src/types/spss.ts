export type VariableType =
  | 'numeric'
  | 'comma'
  | 'dot'
  | 'scientific'
  | 'date'
  | 'dollar'
  | 'custom_currency'
  | 'string';

export type VariableAlign = 'left' | 'center' | 'right';

export type VariableMeasure = 'scale' | 'ordinal' | 'nominal';

export type VariableRole =
  | 'input'
  | 'target'
  | 'both'
  | 'none'
  | 'partition'
  | 'split';

export type MissingValueMode = 'none' | 'discrete' | 'range_discrete';

export interface MissingValueSpec {
  mode: MissingValueMode;
  discrete: Array<string | number>;
  range?: {
    low: number;
    high: number;
  };
}

export interface ValidationRuleSpec {
  min?: number;
  max?: number;
  restrictToLabels?: boolean;
  required?: boolean;
}

export type EndpointInclusionRule = 'lower_inclusive' | 'upper_inclusive';

export type ClassBinMode = 'auto_count' | 'auto_width' | 'custom';

export interface ClassInterval {
  id: string;
  lower: number | null;
  upper: number | null;
  label: string;
}

export interface ClassBinSpec {
  enabled: boolean;
  mode: ClassBinMode;
  endpointRule: EndpointInclusionRule;
  binCount?: number;
  binWidth?: number;
  startAt?: number;
  intervals: ClassInterval[];
}

export interface VariableDefinition {
  id: string;
  name: string;
  type: VariableType;
  width: number;
  decimals: number;
  label: string;
  valueLabels: Record<string, string>;
  missingValues: MissingValueSpec;
  columns: number;
  align: VariableAlign;
  measure: VariableMeasure;
  role: VariableRole;
  validation?: ValidationRuleSpec;
  classBins?: ClassBinSpec;
}

export type CellValue = string | number | null;

export interface DataRow {
  id: string;
  values: Record<string, CellValue>;
}

export type ActiveSheetView = 'data' | 'variable';

export interface DatasetState {
  datasetName: string;
  variables: VariableDefinition[];
  rows: DataRow[];
  showValueLabels: boolean;
  activeView: ActiveSheetView;
  lastBackupTimestamp: string | null;
}

export type CellIssueKind =
  | 'missing_user'
  | 'validation_range'
  | 'validation_uncoded'
  | 'validation_required'
  | 'validation_type';

export interface CellIssue {
  rowIndex: number;
  rowId: string;
  variableId: string;
  variableName: string;
  kind: CellIssueKind;
  message: string;
}

export interface FrequencyRow {
  rawKey: string;
  label: string;
  isMissing: boolean;
  count: number;
  percent: number;
  validPercent: number | null;
  cumulativePercent: number | null;
  midpoint?: number | null;
  intervalId?: string;
}

export interface VariableStats {
  variableId: string;
  totalN: number;
  validN: number;
  userMissingN: number;
  systemMissingN: number;
  mean: number | null;
  median: number | null;
  stdDev: number | null;
  min: number | null;
  max: number | null;
  frequencies: FrequencyRow[];
}

export type HostEnvironment = 'simulator' | 'excel';

export interface CodebookDocument {
  format: 'statsheet-codebook-v1';
  exportedAt: string;
  datasetName: string;
  variables: VariableDefinition[];
}

export interface ToastMessage {
  id: string;
  tone: 'success' | 'info' | 'warning' | 'error';
  title: string;
  detail?: string;
}
