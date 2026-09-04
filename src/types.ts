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
}
