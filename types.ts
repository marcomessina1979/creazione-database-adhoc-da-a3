export interface SkippedRowInfo {
  row_excel: number;
  article: string;
  reason: string;
}

export interface UnreconciledRow {
  rowIndex: number; // 0-based index from AoA
  excelRow: number; // 1-based for display
  segments: (string | number)[];
  description: string;
}

export interface TextValueInfo {
  codice: string;
  campo: 'Q.ty' | 'Unit Price';
  valore_originale: any;
}

export interface MissingValueInfo {
  codice: string;
  campo: 'Q.ty' | 'Unit Price';
}

export interface DescriptionMismatchInfo {
  codice: string;
  db_description: string;
  a3_description: string;
}

export interface UnprocessedRowInfo {
  headers: string[];
  rows: any[][];
}

export interface FlaggedRowInfo {
  codice: string;
  cella: string;
}

export interface ProcessResult {
  summary: {
    updated_rows: number;
    found_and_updated: string[];
    not_found_in_db: string[];
    duplicates_in_db: string[];
    text_values_detected: TextValueInfo[];
    missing_values_replaced: MissingValueInfo[];
    description_mismatches: DescriptionMismatchInfo[];
    unprocessed_db_rows: UnprocessedRowInfo;
    skipped_strikethrough_rows: SkippedRowInfo[];
    lumpsum_rows: FlaggedRowInfo[];
    included_rows: FlaggedRowInfo[];
    fallback_mode: 'enabled' | 'disabled';
    assunzioni: string[];
    output_file: string;
  };
  updatedFileBuffer: Uint8Array;
}