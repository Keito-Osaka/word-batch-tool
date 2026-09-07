export type OutputFormat = "word" | "pdf";
export type OutputMethod = "folder" | "merged" | "zip";
export type ExcludeMode =
  | "any_empty_except_first"
  | "any_empty"
  | "all_empty_except_first"
  | "selected_column_number_empty";

export interface Settings {
  outputFormat: OutputFormat;
  outputMethod: OutputMethod;
  addSerialNumber: boolean;
  serialDigits: number;
  formatAmountWithComma: boolean;
  rowExcludeMode: ExcludeMode;
  targetColumnNumber: number;
  filenameKeys: string[];
  fastPdfSplitEnabled: boolean;
}

export interface ExcludedRow {
  source_row_number: number;
  row: Record<string, string>;
  reason: string;
}

export interface DataPreview {
  original_count: number;
  example_count: number;
  excluded_count: number;
  included_count: number;
  columns: string[];
  included_rows: Record<string, string>[];
  excluded_rows: ExcludedRow[];
}

export interface GenerateResult {
  success: boolean;
  output_path: string;
  generated_count: number;
  output_directory?: string;
  format?: OutputFormat;
  method?: OutputMethod;
}

export interface GenerationProgress {
  type: "progress";
  stage: string;
  label: string;
  current: number;
  total: number;
  percent: number | null;
  detail: string;
  state: "running" | "done" | "notice" | "error";
}

export interface DroppedPathClassification {
  template_path: string | null; data_path: string | null; output_path: string | null; unsupported_paths: string[];
}
