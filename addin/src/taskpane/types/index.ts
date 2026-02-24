export interface CellStyle {
  sz?: number;
  color?: string;
  family?: string;
  fgColor?: string;
  bold?: boolean;
  italic?: boolean;
  numberFormat?: string;
}

export interface WorksheetContext {
  name: string;
  sheetId: number;
  dimension: string | null;
  selection: string;
  cells?: Record<string, any>;
  csv?: string;
  rowCount?: number;
  columnCount?: number;
  hasMore?: boolean;
  styles: Record<string, CellStyle>;
  borders: Record<string, any>;
}

export interface WorkbookContext {
  workbook_name?: string;
  worksheet: WorksheetContext;
  sheets: Array<{ name: string; sheetId: number }>;
  tables: Array<{ name: string; range: string }>;
  named_ranges: string[];
}

export interface ToolExecuteRequest {
  type: "tool_execute_request";
  tool_call_id: string;
  nonce: string;
  expires_at: number;
  tool_name: string;
  tool_input: Record<string, any>;
}

export interface ToolResultPayload {
  tool_call_id: string;
  nonce: string;
  result: Record<string, any> | null;
  error: { code: string; message: string; details?: any } | null;
}
