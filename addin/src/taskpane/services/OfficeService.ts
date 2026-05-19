/* global Excel */

import { CellStyle, WorkbookContext } from "../types";

export class OfficeService {
  private static createUndoSheetName(kind: string): string {
    const suffix = `${Date.now().toString(36)}${Math.random().toString(36).slice(2, 8)}`;
    return `__undo_${kind}_${suffix}`.slice(0, 31);
  }

  private static encodeRestoreToken(kind: string, payload: Record<string, any>): string {
    const json = JSON.stringify({ version: 1, kind, payload });
    const bytes = new TextEncoder().encode(json);
    let binary = "";
    bytes.forEach((byte) => {
      binary += String.fromCharCode(byte);
    });
    return btoa(binary);
  }

  private static decodeRestoreToken(restoreToken: string, expectedKind: string): Record<string, any> {
    try {
      const binary = atob(String(restoreToken || ""));
      const bytes = Uint8Array.from(binary, (char) => char.charCodeAt(0));
      const decoded = JSON.parse(new TextDecoder().decode(bytes));
      if (decoded?.version !== 1 || decoded?.kind !== expectedKind || typeof decoded.payload !== "object") {
        throw new Error("wrong token kind");
      }
      return decoded.payload;
    } catch {
      throw new Error(`Invalid restore_token for ${expectedKind}`);
    }
  }

  static async gatherContext(): Promise<WorkbookContext> {
    return Excel.run(async (ctx) => {
      try {
        console.log("[OfficeService] gatherContext starting...");
        const workbook = ctx.workbook;
        const activeSheet = workbook.worksheets.getActiveWorksheet();
        const usedRange = activeSheet.getUsedRangeOrNullObject();
        const tables = activeSheet.tables;
        const names = workbook.names;

        // getSelectedRange() throws on invalid selection — try separately
        let selection: Excel.Range | null = null;
        let hasSelection = false;
        try {
          selection = workbook.getSelectedRange();
          selection.load(["address", "values", "rowCount", "columnCount", "rowIndex", "columnIndex"]);
          hasSelection = true;
        } catch {
          console.log("[OfficeService] getSelectedRange not available at proxy time, will retry after sync");
        }

        workbook.load("name");
        workbook.worksheets.load("items/name");
        activeSheet.load("name");
        usedRange.load(["address", "rowCount", "columnCount", "isNullObject"]);
        tables.load("items/name");
        names.load("items/name");

        console.log("[OfficeService] First sync...");
        try {
          await ctx.sync();
        } catch (syncError: any) {
          // If sync fails due to InvalidSelection, retry without selection
          if (syncError?.code === "InvalidSelection" || syncError?.debugInfo?.code === "InvalidSelection") {
            console.log("[OfficeService] First sync failed (InvalidSelection), retrying without selection");
            selection = null;
            hasSelection = false;
            await ctx.sync();
          } else {
            throw syncError;
          }
        }
        console.log("[OfficeService] First sync complete. Sheet:", activeSheet.name);

      // Load table ranges separately after first sync
      const tableData: Array<{ name: string; range: string }> = [];
      const tableRangeObjects: Excel.Range[] = [];
      for (const table of tables.items) {
        const range = table.getRange();
        range.load("address");
        tableRangeObjects.push(range);
      }
      if (tableRangeObjects.length > 0) {
        console.log("[OfficeService] Loading table ranges...");
        await ctx.sync();
      }
      console.log("[OfficeService] Tables loaded:", tables.items.length);
      for (let i = 0; i < tables.items.length; i++) {
        tableData.push({
          name: tables.items[i].name,
          range: tableRangeObjects[i].address,
        });
      }

      const sheetHasData =
        !usedRange.isNullObject && usedRange.rowCount > 0 && usedRange.columnCount > 0;
      const sheetIndex = workbook.worksheets.items.findIndex((s) => s.name === activeSheet.name);

      const sheetsInfo = workbook.worksheets.items.map((sheet, index) => ({
        name: sheet.name,
        sheetId: index + 1,
      }));
      // Filter out add-in internal named ranges — FactSet (__FDS*), S&P Capital IQ
      // (IQ_*), data model cache (_bdm.*), and Excel internals (_xlfn*). These are
      // often 4,000+ entries adding ~28K tokens to every request with no value.
      const namedRanges = names.items
        .map((name) => name.name)
        .filter((n) => !n.startsWith("__FDS") && !n.startsWith("IQ_") &&
                       !n.startsWith("_bdm.") && !n.startsWith("_xlfn"));

      // If no valid selection, return context without cell data
      if (!hasSelection) {
        console.log("[OfficeService] No valid selection, returning context without cell data");
        return {
          workbook_name: workbook.name || undefined,
          worksheet: {
            name: activeSheet.name,
            sheetId: sheetIndex >= 0 ? sheetIndex + 1 : 1,
            dimension: sheetHasData ? usedRange.address.split("!").pop() || null : null,
            selection: "",
            cells: {},
            styles: {},
            borders: {},
          },
          sheets: sheetsInfo,
          tables: tableData,
          named_ranges: namedRanges,
        };
      }

      const totalCells = selection.rowCount * selection.columnCount;
      const useCsv = totalCells > 100; // Switch to CSV for large selections
      const maxRows = 500; // Limit for CSV

      // Load selection formatting
      console.log("[OfficeService] Loading formatting...");
      selection.format.font.load(["bold", "italic", "color", "name", "size"]);
      selection.format.fill.load("color");
      await ctx.sync();
      console.log("[OfficeService] Formatting loaded. Building context...");

      // Build styles object for the selection range
      const styles: Record<string, CellStyle> = {};
      const style: CellStyle = {};
      if (selection.format.font.size) style.sz = selection.format.font.size;
      if (selection.format.font.color) style.color = selection.format.font.color;
      if (selection.format.font.name) style.family = selection.format.font.name;
      if (selection.format.fill.color) style.fgColor = selection.format.fill.color;
      if (selection.format.font.bold === true) style.bold = true;
      if (selection.format.font.italic === true) style.italic = true;

      if (Object.keys(style).length > 0) {
        styles[selection.address.split("!").pop() || selection.address] = style;
      }

      const selectionAddress = selection.address.split("!").pop() || selection.address;

      const worksheetBase = {
        name: activeSheet.name,
        sheetId: sheetIndex >= 0 ? sheetIndex + 1 : 1,
        dimension: sheetHasData ? usedRange.address.split("!").pop() || null : null,
        selection: selectionAddress,  // e.g., "A1:C10"
        styles,
        borders: {},
      };

      if (useCsv) {
        // Large selection: use CSV format
        const rowsToProcess = Math.min(selection.rowCount, maxRows);
        const csvLines: string[] = [];

        for (let r = 0; r < rowsToProcess; r++) {
          const row = selection.values[r];
          const csvRow = row.map((cell: any) => {
            if (cell === null || cell === undefined) return "";
            const str = String(cell);
            if (str.includes(",") || str.includes('"') || str.includes("\n")) {
              return `"${str.replace(/"/g, '""')}"`;
            }
            return str;
          });
          csvLines.push(csvRow.join(","));
        }

        return {
          workbook_name: workbook.name || undefined,
          worksheet: {
            ...worksheetBase,
            csv: csvLines.join("\n"),
            rowCount: selection.rowCount,
            columnCount: selection.columnCount,
            hasMore: selection.rowCount > maxRows,
          },
          sheets: sheetsInfo,
          tables: tableData,
          named_ranges: namedRanges,
        };
      } else {
        // Small selection: use sparse cells format
        const cells: Record<string, any> = {};
        const values = selection.values;
        const startRow = selection.rowIndex;
        const startCol = selection.columnIndex;

        for (let r = 0; r < values.length; r++) {
          for (let c = 0; c < values[r].length; c++) {
            const val = values[r][c];
            if (val !== null && val !== undefined && val !== "") {
              const cellAddress = OfficeService.getCellAddress(startRow + r, startCol + c);
              cells[cellAddress] = val;
            }
          }
        }

        return {
          workbook_name: workbook.name || undefined,
          worksheet: {
            ...worksheetBase,
            cells,
          },
          sheets: sheetsInfo,
          tables: tableData,
          named_ranges: namedRanges,
        };
      }
      } catch (error) {
        console.error("[OfficeService] gatherContext error:", error);
        throw error;
      }
    });
  }

  protected static getCellAddress(row: number, col: number): string {
    return `${OfficeService.getColLetter(col)}${row + 1}`;
  }

  protected static getColLetter(col: number): string {
    let colStr = "";
    let c = col;
    while (c >= 0) {
      colStr = String.fromCharCode((c % 26) + 65) + colStr;
      c = Math.floor(c / 26) - 1;
    }
    return colStr;
  }

  protected static getColIndex(col: string): number {
    const normalized = String(col || "").trim().toUpperCase();
    if (!/^[A-Z]{1,3}$/.test(normalized)) {
      throw new Error(`Invalid column reference: ${col}`);
    }
    let index = 0;
    for (let i = 0; i < normalized.length; i++) {
      index = index * 26 + (normalized.charCodeAt(i) - 64);
    }
    return index - 1;
  }

  protected static normalizeColumnInput(column: string | number): number {
    if (typeof column === "number" && Number.isFinite(column)) {
      const oneBased = Math.trunc(column);
      if (oneBased < 1) {
        throw new Error("column must be >= 1");
      }
      return oneBased - 1;
    }
    if (typeof column === "string") {
      const trimmed = column.trim();
      if (!trimmed) {
        throw new Error("column is required");
      }
      if (/^\d+$/.test(trimmed)) {
        const oneBased = Number(trimmed);
        if (!Number.isFinite(oneBased) || oneBased < 1) {
          throw new Error("column must be >= 1");
        }
        return Math.trunc(oneBased) - 1;
      }
      return OfficeService.getColIndex(trimmed);
    }
    throw new Error("column must be a letter (e.g., 'C') or 1-based index");
  }

  private static buildColumnAddress(startColIndex: number, count: number): string {
    const startCol = OfficeService.getColLetter(startColIndex);
    if (count <= 1) {
      return `${startCol}:${startCol}`;
    }
    const endCol = OfficeService.getColLetter(startColIndex + count - 1);
    return `${startCol}:${endCol}`;
  }

  private static buildColumnLetters(startCol: number, count: number): string[] {
    const letters: string[] = [];
    for (let i = 0; i < count; i++) {
      letters.push(OfficeService.getColLetter(startCol + i));
    }
    return letters;
  }

  static async readCells(rangeAddress: string, includeFormulas = false): Promise<{ address: string; values: any[][] }> {
    console.log("[OfficeService] readCells:", rangeAddress);
    return Excel.run(async (ctx) => {
      try {
        const range = ctx.workbook.worksheets.getActiveWorksheet().getRange(rangeAddress);
        range.load(["address", includeFormulas ? "formulas" : "values"]);
        await ctx.sync();
        const values = includeFormulas ? range.formulas : range.values;
        console.log("[OfficeService] readCells complete:", range.address);
        return { address: range.address, values };
      } catch (error) {
        console.error("[OfficeService] readCells error:", error);
        throw error;
      }
    });
  }

  static async readRangeCsv(
    rangeAddress: string,
    sheetName?: string
  ): Promise<{
    csv: string;
    columns: string[];
    rowCount: number;
    columnCount: number;
    sheetName: string;
    hasMore: boolean;
  }> {
    return Excel.run(async (ctx) => {
      const sheet = sheetName
        ? ctx.workbook.worksheets.getItem(sheetName)
        : ctx.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(rangeAddress);

      sheet.load("name");
      range.load(["values", "rowCount", "columnCount", "columnIndex"]);
      await ctx.sync();

      const maxRows = 500; // Limit for performance
      const hasMore = range.rowCount > maxRows;
      const rowsToProcess = Math.min(range.rowCount, maxRows);

      // Convert to CSV
      const csvLines: string[] = [];
      for (let r = 0; r < rowsToProcess; r++) {
        const row = range.values[r];
        const csvRow = row.map((cell: any) => {
          if (cell === null || cell === undefined) return "";
          const str = String(cell);
          // Escape quotes and wrap in quotes if contains comma, quote, or newline
          if (str.includes(",") || str.includes('"') || str.includes("\n")) {
            return `"${str.replace(/"/g, '""')}"`;
          }
          return str;
        });
        csvLines.push(csvRow.join(","));
      }

      return {
        csv: csvLines.join("\n"),
        columns: OfficeService.buildColumnLetters(range.columnIndex, range.columnCount),
        rowCount: range.rowCount,
        columnCount: range.columnCount,
        sheetName: sheet.name,
        hasMore,
      };
    });
  }

  static async readSheetCsv(
    sheetName?: string
  ): Promise<{
    csv: string;
    columns: string[];
    rowCount: number;
    columnCount: number;
    sheetName: string;
    dimension: string | null;
    hasMore: boolean;
  }> {
    return Excel.run(async (ctx) => {
      const sheet = sheetName
        ? ctx.workbook.worksheets.getItem(sheetName)
        : ctx.workbook.worksheets.getActiveWorksheet();
      const usedRange = sheet.getUsedRangeOrNullObject();

      sheet.load("name");
      usedRange.load(["values", "address", "rowCount", "columnCount", "columnIndex", "isNullObject"]);
      await ctx.sync();

      if (usedRange.isNullObject) {
        return {
          csv: "",
          columns: [],
          rowCount: 0,
          columnCount: 0,
          sheetName: sheet.name,
          dimension: null,
          hasMore: false,
        };
      }

      const maxRows = 500;
      const maxChars = 100000;
      const rowsToProcess = Math.min(usedRange.rowCount, maxRows);
      const estimatedChars = rowsToProcess * usedRange.columnCount * 8; // ~8 chars per cell average
      if (estimatedChars > maxChars) {
        throw new Error(
          `Sheet too large to return as CSV (estimated ${Math.round(estimatedChars / 1000)}K chars, limit ${maxChars / 1000}K). ` +
          `Dimensions: ${usedRange.rowCount} rows × ${usedRange.columnCount} columns. ` +
          `Use read_range_csv for targeted reads.`
        );
      }
      const hasMore = usedRange.rowCount > maxRows;

      const csvLines: string[] = [];
      for (let r = 0; r < rowsToProcess; r++) {
        const row = usedRange.values[r];
        const csvRow = row.map((cell: any) => {
          if (cell === null || cell === undefined) return "";
          const str = String(cell);
          if (str.includes(",") || str.includes('"') || str.includes("\n")) {
            return `"${str.replace(/"/g, '""')}"`;
          }
          return str;
        });
        csvLines.push(csvRow.join(","));
      }

      return {
        csv: csvLines.join("\n"),
        columns: OfficeService.buildColumnLetters(usedRange.columnIndex, usedRange.columnCount),
        rowCount: usedRange.rowCount,
        columnCount: usedRange.columnCount,
        sheetName: sheet.name,
        dimension: usedRange.address.split("!").pop() || null,
        hasMore,
      };
    });
  }

  static async writeCells(
    rangeAddress: string,
    values: string | number | null | any[][],
    numberFormat?: string
  ): Promise<{ address: string; rows: number; cols: number }> {
    console.log("[OfficeService] writeCells:", rangeAddress, values);
    return Excel.run(async (ctx) => {
      try {
        const range = ctx.workbook.worksheets.getActiveWorksheet().getRange(rangeAddress);
        range.load(["address", "rowCount", "columnCount"]);
        await ctx.sync();

        const isScalar = typeof values === "string" || typeof values === "number" || values === null;
        if ((range.rowCount > 1 || range.columnCount > 1) && isScalar) {
          throw new Error("Scalar values require a single-cell range.");
        }

        if (Array.isArray(values)) {
          range.values = values as any[][];
        } else {
          range.values = [[values]];
        }

        if (numberFormat) {
          range.numberFormat = [[numberFormat]];
        }

        await ctx.sync();
        console.log("[OfficeService] writeCells complete:", range.address);
        return { address: range.address, rows: range.rowCount, cols: range.columnCount };
      } catch (error) {
        console.error("[OfficeService] writeCells error:", error);
        throw error;
      }
    });
  }

  static async getSelection(): Promise<{ address: string; values: any[][] }> {
    return Excel.run(async (ctx) => {
      const range = ctx.workbook.getSelectedRange();
      range.load(["address", "values"]);
      await ctx.sync();
      return { address: range.address, values: range.values };
    });
  }

  static async getUsedRange(sheetName?: string): Promise<{ address: string; rows: number; cols: number } | null> {
    return Excel.run(async (ctx) => {
      const sheet = sheetName
        ? ctx.workbook.worksheets.getItem(sheetName)
        : ctx.workbook.worksheets.getActiveWorksheet();
      const usedRange = sheet.getUsedRangeOrNullObject();
      usedRange.load(["address", "rowCount", "columnCount", "isNullObject"]);
      await ctx.sync();
      if (usedRange.isNullObject) {
        return null;
      }
      return { address: usedRange.address, rows: usedRange.rowCount, cols: usedRange.columnCount };
    });
  }

  static async listSheets(): Promise<{
    sheets: Array<{
      name: string;
      sheetId: number;
      visibility: string;
      isActive: boolean;
      dimension: string | null;
      rowCount: number;
      columnCount: number;
    }>;
    activeSheet: string;
    count: number;
  }> {
    return Excel.run(async (ctx) => {
      const workbook = ctx.workbook;
      const worksheets = workbook.worksheets;
      const activeSheet = worksheets.getActiveWorksheet();

      worksheets.load("items/name,items/visibility");
      activeSheet.load("name");
      await ctx.sync();

      const usedRanges = worksheets.items.map((sheet) => sheet.getUsedRangeOrNullObject());
      for (const usedRange of usedRanges) {
        usedRange.load(["address", "rowCount", "columnCount", "isNullObject"]);
      }
      await ctx.sync();

      const sheets = worksheets.items.map((sheet, index) => {
        const usedRange = usedRanges[index];
        const isEmpty = usedRange.isNullObject;
        return {
          name: sheet.name,
          sheetId: index + 1,
          visibility: String((sheet as any).visibility || "Visible"),
          isActive: sheet.name === activeSheet.name,
          dimension: isEmpty ? null : usedRange.address.split("!").pop() || null,
          rowCount: isEmpty ? 0 : usedRange.rowCount,
          columnCount: isEmpty ? 0 : usedRange.columnCount,
        };
      });

      return {
        sheets,
        activeSheet: activeSheet.name,
        count: sheets.length,
      };
    });
  }

  static async switchSheet(sheetName: string): Promise<{
    previousSheet: string;
    sheetName: string;
    dimension: string | null;
    rowCount: number;
    columnCount: number;
  }> {
    const normalizedSheetName = String(sheetName || "").trim();
    if (!normalizedSheetName) {
      throw new Error("sheet_name is required");
    }

    return Excel.run(async (ctx) => {
      const workbook = ctx.workbook;
      const previousSheet = workbook.worksheets.getActiveWorksheet();
      const targetSheet = workbook.worksheets.getItemOrNullObject(normalizedSheetName);

      previousSheet.load("name");
      targetSheet.load(["name", "isNullObject"]);
      await ctx.sync();

      if (targetSheet.isNullObject) {
        throw new Error(`Sheet not found: ${normalizedSheetName}`);
      }

      targetSheet.activate();
      const usedRange = targetSheet.getUsedRangeOrNullObject();
      usedRange.load(["address", "rowCount", "columnCount", "isNullObject"]);
      await ctx.sync();

      return {
        previousSheet: previousSheet.name,
        sheetName: targetSheet.name,
        dimension: usedRange.isNullObject ? null : usedRange.address.split("!").pop() || null,
        rowCount: usedRange.isNullObject ? 0 : usedRange.rowCount,
        columnCount: usedRange.isNullObject ? 0 : usedRange.columnCount,
      };
    });
  }

  static async renameSheet(
    newName: string,
    sheetName?: string
  ): Promise<{
    previousSheet: string;
    sheetName: string;
    activeSheet: string;
    renamed: boolean;
  }> {
    const normalizedNewName = String(newName || "").trim();
    if (!normalizedNewName) {
      throw new Error("new_name is required");
    }

    const normalizedSheetName = sheetName ? String(sheetName).trim() : "";

    return Excel.run(async (ctx) => {
      const workbook = ctx.workbook;
      const activeSheet = workbook.worksheets.getActiveWorksheet();
      const targetSheet = normalizedSheetName
        ? workbook.worksheets.getItemOrNullObject(normalizedSheetName)
        : activeSheet;
      const conflictingSheet = workbook.worksheets.getItemOrNullObject(normalizedNewName);

      activeSheet.load("name");
      targetSheet.load(["name", "isNullObject"]);
      conflictingSheet.load(["name", "isNullObject"]);
      await ctx.sync();

      if (targetSheet.isNullObject) {
        throw new Error(`Sheet not found: ${normalizedSheetName}`);
      }

      const previousSheet = targetSheet.name;
      if (previousSheet === normalizedNewName) {
        return {
          previousSheet,
          sheetName: targetSheet.name,
          activeSheet: activeSheet.name,
          renamed: false,
        };
      }

      if (!conflictingSheet.isNullObject && conflictingSheet.name !== previousSheet) {
        throw new Error(`A sheet named '${normalizedNewName}' already exists`);
      }

      targetSheet.name = normalizedNewName;
      await ctx.sync();

      const currentActive = workbook.worksheets.getActiveWorksheet();
      currentActive.load("name");
      await ctx.sync();

      return {
        previousSheet,
        sheetName: normalizedNewName,
        activeSheet: currentActive.name,
        renamed: true,
      };
    });
  }

  static async createSheet(
    sheetName?: string,
    activate = true
  ): Promise<{
    created: boolean;
    sheetName: string;
    activeSheet: string;
    summary: string;
  }> {
    const normalizedSheetName = sheetName ? String(sheetName).trim() : "";

    return Excel.run(async (ctx) => {
      const workbook = ctx.workbook;
      const worksheets = workbook.worksheets;

      if (normalizedSheetName) {
        const existing = worksheets.getItemOrNullObject(normalizedSheetName);
        existing.load(["name", "isNullObject"]);
        await ctx.sync();
        if (!existing.isNullObject) {
          throw new Error(`A sheet named '${normalizedSheetName}' already exists`);
        }
      }

      const newSheet = normalizedSheetName ? worksheets.add(normalizedSheetName) : worksheets.add();
      if (activate) {
        newSheet.activate();
      }

      newSheet.load("name");
      const activeSheet = worksheets.getActiveWorksheet();
      activeSheet.load("name");
      await ctx.sync();

      return {
        created: true,
        sheetName: newSheet.name,
        activeSheet: activeSheet.name,
        summary: `Created sheet '${newSheet.name}'.`,
      };
    });
  }

  static async deleteSheet(
    sheetName: string,
    force = false
  ): Promise<{
    deleted: boolean;
    requires_confirmation?: boolean;
    sheetName: string;
    activeSheet?: string;
    remainingSheets?: number;
    summary: string;
    restore_token?: string;
    undo_tool?: string;
  }> {
    const normalizedSheetName = String(sheetName || "").trim();
    if (!normalizedSheetName) {
      throw new Error("sheet_name is required");
    }

    if (!force) {
      return {
        deleted: false,
        requires_confirmation: true,
        sheetName: normalizedSheetName,
        summary: "Set force=true to confirm irreversible sheet deletion.",
      };
    }

    return Excel.run(async (ctx) => {
      const workbook = ctx.workbook;
      const worksheets = workbook.worksheets;
      const targetSheet = worksheets.getItemOrNullObject(normalizedSheetName);

      worksheets.load("items/name,items/visibility");
      targetSheet.load(["name", "isNullObject"]);
      await ctx.sync();

      if (targetSheet.isNullObject) {
        throw new Error(`Sheet not found: ${normalizedSheetName}`);
      }
      if (worksheets.items.length <= 1) {
        throw new Error("Cannot delete the last worksheet in a workbook");
      }

      const deletedName = targetSheet.name;
      const backupSheetName = OfficeService.createUndoSheetName("sheet");
      const visibleOtherSheet = worksheets.items.some(
        (sheet) => sheet.name !== deletedName && sheet.visibility === Excel.SheetVisibility.visible
      );
      const backupSheet = targetSheet.copy(Excel.WorksheetPositionType.after, targetSheet);
      backupSheet.name = backupSheetName;
      if (visibleOtherSheet) {
        backupSheet.visibility = Excel.SheetVisibility.veryHidden;
      }
      targetSheet.delete();
      await ctx.sync();

      worksheets.load("items/name");
      const activeSheet = worksheets.getActiveWorksheet();
      activeSheet.load("name");
      await ctx.sync();

      return {
        deleted: true,
        sheetName: deletedName,
        activeSheet: activeSheet.name,
        remainingSheets: worksheets.items.length,
        restore_token: OfficeService.encodeRestoreToken("delete_sheet", {
          sheetName: deletedName,
          backupSheetName,
        }),
        undo_tool: "restore_deleted_sheet",
        summary: `Deleted sheet '${deletedName}'.`,
      };
    });
  }

  static async deleteRow(
    row: number,
    count = 1,
    sheetName?: string,
    force = false
  ): Promise<{
    deleted: boolean;
    requires_confirmation?: boolean;
    sheetName: string;
    deletedRange: string;
    row: number;
    count: number;
    summary: string;
  }> {
    const startRow = Math.trunc(Number(row));
    const deleteCount = Math.trunc(Number(count));
    if (!Number.isFinite(startRow) || startRow < 1) {
      throw new Error("row must be >= 1");
    }
    if (!Number.isFinite(deleteCount) || deleteCount < 1) {
      throw new Error("count must be >= 1");
    }

    const rowAddress = `${startRow}:${startRow + deleteCount - 1}`;
    if (!force) {
      return {
        deleted: false,
        requires_confirmation: true,
        sheetName: String(sheetName || "").trim() || "(active)",
        deletedRange: rowAddress,
        row: startRow,
        count: deleteCount,
        summary: "Set force=true to confirm irreversible row deletion.",
      };
    }

    return Excel.run(async (ctx) => {
      const sheet = sheetName
        ? ctx.workbook.worksheets.getItem(sheetName)
        : ctx.workbook.worksheets.getActiveWorksheet();
      const rowRange = sheet.getRange(rowAddress);

      sheet.load("name");
      rowRange.load("address");
      rowRange.delete(Excel.DeleteShiftDirection.up);
      await ctx.sync();

      return {
        deleted: true,
        sheetName: sheet.name,
        deletedRange: rowRange.address,
        row: startRow,
        count: deleteCount,
        summary: `Deleted row(s) ${rowAddress} on '${sheet.name}'.`,
      };
    });
  }

  static async insertRow(
    row: number,
    count = 1,
    sheetName?: string
  ): Promise<{
    sheetName: string;
    insertedRange: string;
    row: number;
    count: number;
  }> {
    const startRow = Math.trunc(Number(row));
    const insertCount = Math.trunc(Number(count));
    if (!Number.isFinite(startRow) || startRow < 1) {
      throw new Error("row must be >= 1");
    }
    if (!Number.isFinite(insertCount) || insertCount < 1) {
      throw new Error("count must be >= 1");
    }

    return Excel.run(async (ctx) => {
      const sheet = sheetName
        ? ctx.workbook.worksheets.getItem(sheetName)
        : ctx.workbook.worksheets.getActiveWorksheet();
      const address = `${startRow}:${startRow + insertCount - 1}`;
      const rowRange = sheet.getRange(address);

      sheet.load("name");
      rowRange.load("address");
      rowRange.insert(Excel.InsertShiftDirection.down);
      await ctx.sync();

      return {
        sheetName: sheet.name,
        insertedRange: rowRange.address,
        row: startRow,
        count: insertCount,
      };
    });
  }

  static async deleteColumn(
    column: string | number,
    count = 1,
    sheetName?: string,
    force = false
  ): Promise<{
    deleted: boolean;
    requires_confirmation?: boolean;
    sheetName: string;
    deletedRange: string;
    column: string;
    count: number;
    summary: string;
    restore_token?: string;
    undo_tool?: string;
  }> {
    const startColIndex = OfficeService.normalizeColumnInput(column);
    const deleteCount = Math.trunc(Number(count));
    if (!Number.isFinite(deleteCount) || deleteCount < 1) {
      throw new Error("count must be >= 1");
    }

    const rangeAddress = OfficeService.buildColumnAddress(startColIndex, deleteCount);
    if (!force) {
      return {
        deleted: false,
        requires_confirmation: true,
        sheetName: String(sheetName || "").trim() || "(active)",
        deletedRange: rangeAddress,
        column: OfficeService.getColLetter(startColIndex),
        count: deleteCount,
        summary: "Set force=true to confirm irreversible column deletion.",
      };
    }

    return Excel.run(async (ctx) => {
      const worksheets = ctx.workbook.worksheets;
      const sheet = sheetName
        ? ctx.workbook.worksheets.getItem(sheetName)
        : ctx.workbook.worksheets.getActiveWorksheet();
      const colRange = sheet.getRange(rangeAddress);
      const usedRange = sheet.getUsedRangeOrNullObject();

      sheet.load("name");
      usedRange.load(["isNullObject", "rowCount"]);
      colRange.load("address");
      await ctx.sync();

      const rowCount = usedRange.isNullObject ? 1 : Math.max(1, usedRange.rowCount);
      const backupSheetName = OfficeService.createUndoSheetName("col");
      const backupSheet = worksheets.add(backupSheetName);
      const sourceRange = sheet.getRangeByIndexes(0, startColIndex, rowCount, deleteCount);
      const backupRange = backupSheet.getRangeByIndexes(0, 0, rowCount, deleteCount);
      backupRange.copyFrom(sourceRange, Excel.RangeCopyType.all);
      backupSheet.visibility = Excel.SheetVisibility.veryHidden;
      colRange.delete(Excel.DeleteShiftDirection.left);
      await ctx.sync();

      return {
        deleted: true,
        sheetName: sheet.name,
        deletedRange: colRange.address,
        column: OfficeService.getColLetter(startColIndex),
        count: deleteCount,
        restore_token: OfficeService.encodeRestoreToken("delete_column", {
          sheetName: sheet.name,
          backupSheetName,
          startColIndex,
          rowCount,
          count: deleteCount,
          rangeAddress,
        }),
        undo_tool: "restore_deleted_column",
        summary: `Deleted column(s) ${rangeAddress} on '${sheet.name}'.`,
      };
    });
  }

  static async restoreDeletedSheet(
    restoreToken: string
  ): Promise<{
    restored: boolean;
    sheetName: string;
    backupSheetName: string;
    summary: string;
  }> {
    const payload = OfficeService.decodeRestoreToken(restoreToken, "delete_sheet");
    const sheetName = String(payload.sheetName || "").trim();
    const backupSheetName = String(payload.backupSheetName || "").trim();
    if (!sheetName || !backupSheetName) {
      throw new Error("restore_token missing sheet snapshot");
    }

    return Excel.run(async (ctx) => {
      const worksheets = ctx.workbook.worksheets;
      const existing = worksheets.getItemOrNullObject(sheetName);
      const backup = worksheets.getItemOrNullObject(backupSheetName);
      existing.load("isNullObject");
      backup.load(["name", "isNullObject"]);
      await ctx.sync();

      if (!existing.isNullObject) {
        throw new Error(`Cannot restore '${sheetName}' because a sheet with that name already exists`);
      }
      if (backup.isNullObject) {
        throw new Error(`Restore backup sheet not found: ${backupSheetName}`);
      }

      backup.visibility = Excel.SheetVisibility.visible;
      backup.name = sheetName;
      backup.activate();
      await ctx.sync();

      return {
        restored: true,
        sheetName,
        backupSheetName,
        summary: `Restored sheet '${sheetName}'.`,
      };
    });
  }

  static async restoreDeletedColumn(
    restoreToken: string
  ): Promise<{
    restored: boolean;
    sheetName: string;
    restoredRange: string;
    backupSheetName: string;
    summary: string;
  }> {
    const payload = OfficeService.decodeRestoreToken(restoreToken, "delete_column");
    const sheetName = String(payload.sheetName || "").trim();
    const backupSheetName = String(payload.backupSheetName || "").trim();
    const startColIndex = Math.trunc(Number(payload.startColIndex));
    const rowCount = Math.trunc(Number(payload.rowCount));
    const count = Math.trunc(Number(payload.count));
    const rangeAddress = String(payload.rangeAddress || "");
    if (!sheetName || !backupSheetName || !Number.isFinite(startColIndex) || startColIndex < 0 || !Number.isFinite(rowCount) || rowCount < 1 || !Number.isFinite(count) || count < 1 || !rangeAddress) {
      throw new Error("restore_token missing column snapshot");
    }

    return Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItem(sheetName);
      const backup = ctx.workbook.worksheets.getItemOrNullObject(backupSheetName);
      const insertRange = sheet.getRange(rangeAddress);
      backup.load(["name", "isNullObject"]);
      await ctx.sync();

      if (backup.isNullObject) {
        throw new Error(`Restore backup sheet not found: ${backupSheetName}`);
      }

      insertRange.insert(Excel.InsertShiftDirection.right);
      const targetRange = sheet.getRangeByIndexes(0, startColIndex, rowCount, count);
      const backupRange = backup.getRangeByIndexes(0, 0, rowCount, count);
      targetRange.copyFrom(backupRange, Excel.RangeCopyType.all);
      backup.delete();
      await ctx.sync();

      return {
        restored: true,
        sheetName,
        restoredRange: rangeAddress,
        backupSheetName,
        summary: `Restored column(s) ${rangeAddress} on '${sheetName}'.`,
      };
    });
  }

  static async insertColumn(
    column: string | number,
    count = 1,
    sheetName?: string
  ): Promise<{
    sheetName: string;
    insertedRange: string;
    column: string;
    count: number;
  }> {
    const startColIndex = OfficeService.normalizeColumnInput(column);
    const insertCount = Math.trunc(Number(count));
    if (!Number.isFinite(insertCount) || insertCount < 1) {
      throw new Error("count must be >= 1");
    }
    const rangeAddress = OfficeService.buildColumnAddress(startColIndex, insertCount);

    return Excel.run(async (ctx) => {
      const sheet = sheetName
        ? ctx.workbook.worksheets.getItem(sheetName)
        : ctx.workbook.worksheets.getActiveWorksheet();
      const colRange = sheet.getRange(rangeAddress);

      sheet.load("name");
      colRange.load("address");
      colRange.insert(Excel.InsertShiftDirection.right);
      await ctx.sync();

      return {
        sheetName: sheet.name,
        insertedRange: colRange.address,
        column: OfficeService.getColLetter(startColIndex),
        count: insertCount,
      };
    });
  }

  static async createTable(
    rangeAddress: string,
    hasHeaders = true,
    tableName?: string
  ): Promise<{ name: string; range: string }> {
    return Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      const table = sheet.tables.add(rangeAddress, hasHeaders);
      if (tableName) {
        table.name = tableName;
      }
      table.load("name");
      const tableRange = table.getRange();
      tableRange.load("address");
      await ctx.sync();
      return { name: table.name, range: tableRange.address };
    });
  }

  static async formatCells(options: {
    range: string;
    bold?: boolean;
    italic?: boolean;
    fill_color?: string;
    font_color?: string;
    number_format?: string;
    horizontal_alignment?: string;
  }): Promise<{ address: string }> {
    return Excel.run(async (ctx) => {
      const range = ctx.workbook.worksheets.getActiveWorksheet().getRange(options.range);
      range.load("address");
      if (typeof options.bold === "boolean") {
        range.format.font.bold = options.bold;
      }
      if (typeof options.italic === "boolean") {
        range.format.font.italic = options.italic;
      }
      if (options.fill_color != null) {
        if (options.fill_color === "" || options.fill_color.toLowerCase() === "none") {
          range.format.fill.clear();
        } else {
          range.format.fill.color = options.fill_color;
        }
      }
      if (options.font_color != null) {
        if (options.font_color === "" || options.font_color.toLowerCase() === "auto") {
          range.format.font.color = "#000000";
        } else {
          range.format.font.color = options.font_color;
        }
      }
      if (options.number_format) {
        range.numberFormat = [[options.number_format]];
      }
      if (options.horizontal_alignment != null) {
        const raw = String(options.horizontal_alignment).trim().toLowerCase();
        const alignmentMap: Record<string, string> = {
          left: "Left",
          center: "Center",
          right: "Right",
          justify: "Justify",
          distributed: "Distributed",
        };
        const excelAlignment = alignmentMap[raw];
        if (!excelAlignment) {
          throw new Error(
            "horizontal_alignment must be one of: left, center, right, justify, distributed"
          );
        }
        range.format.horizontalAlignment = excelAlignment as Excel.HorizontalAlignment;
      }
      await ctx.sync();
      return { address: range.address };
    });
  }

  static async findCells(options: {
    search_type: "value" | "fill_color" | "font";
    query: string;
    range?: string;
    sheet_name?: string;
  }): Promise<{ matches: Array<{ address: string; value: any }>; count: number }> {
    const MAX_MATCHES = 100;
    const MAX_FORMAT_CELLS = 10_000;
    const COLOR_NAME_MAP: Record<string, string> = {
      yellow: "#FFFF00",
      red: "#FF0000",
      green: "#00FF00",
      blue: "#0000FF",
      orange: "#FFA500",
      purple: "#800080",
      white: "#FFFFFF",
      black: "#000000",
    };

    return Excel.run(async (ctx) => {
      const sheet = options.sheet_name
        ? ctx.workbook.worksheets.getItem(options.sheet_name)
        : ctx.workbook.worksheets.getActiveWorksheet();

      let range: Excel.Range;
      if (options.range) {
        range = sheet.getRange(options.range);
      } else {
        range = sheet.getUsedRangeOrNullObject();
      }

      range.load(["values", "rowIndex", "columnIndex", "rowCount", "columnCount", "isNullObject"]);
      await ctx.sync();

      if ((range as any).isNullObject) {
        return { matches: [], count: 0 };
      }

      const matches: Array<{ address: string; value: any }> = [];
      const rows = range.rowCount;
      const cols = range.columnCount;
      const totalCells = rows * cols;
      const startRow = range.rowIndex;
      const startCol = range.columnIndex;
      const values = range.values;

      if (options.search_type !== "value" && totalCells > MAX_FORMAT_CELLS) {
        throw new Error(
          `Range too large for ${options.search_type} search (${totalCells} cells, limit ${MAX_FORMAT_CELLS}). ` +
          `Specify a smaller range (e.g., 'A1:Z100').`
        );
      }

      if (options.search_type === "value") {
        const queryLower = options.query.toLowerCase();
        for (let r = 0; r < rows && matches.length < MAX_MATCHES; r++) {
          for (let c = 0; c < cols && matches.length < MAX_MATCHES; c++) {
            const cellVal = values[r][c];
            if (cellVal === null || cellVal === undefined || cellVal === "") continue;
            const strVal = String(cellVal).toLowerCase();
            if (strVal.includes(queryLower)) {
              matches.push({
                address: OfficeService.getCellAddress(startRow + r, startCol + c),
                value: cellVal,
              });
            }
          }
        }
      } else if (options.search_type === "fill_color") {
        const queryNorm = COLOR_NAME_MAP[options.query.toLowerCase()] || options.query.toUpperCase();

        // Batch-load fill colors: queue all getCell().format.fill loads, then single sync
        const cellRefs: Array<{ r: number; c: number; fill: Excel.RangeFill }> = [];
        for (let r = 0; r < rows; r++) {
          for (let c = 0; c < cols; c++) {
            const cell = range.getCell(r, c);
            const fill = cell.format.fill;
            fill.load("color");
            cellRefs.push({ r, c, fill });
          }
        }
        await ctx.sync();

        for (const ref of cellRefs) {
          if (matches.length >= MAX_MATCHES) break;
          const fillColor = (ref.fill.color || "").toUpperCase();
          if (fillColor === queryNorm) {
            matches.push({
              address: OfficeService.getCellAddress(startRow + ref.r, startCol + ref.c),
              value: values[ref.r][ref.c],
            });
          }
        }
      } else if (options.search_type === "font") {
        let fontQuery: Record<string, any>;
        try {
          fontQuery = JSON.parse(options.query);
        } catch {
          return { matches: [], count: 0 };
        }

        // Batch-load font properties for all cells
        const cellRefs: Array<{ r: number; c: number; font: Excel.RangeFont }> = [];
        const propsToLoad: string[] = [];
        if (fontQuery.bold !== undefined) propsToLoad.push("bold");
        if (fontQuery.italic !== undefined) propsToLoad.push("italic");
        if (fontQuery.color !== undefined) propsToLoad.push("color");
        if (propsToLoad.length === 0) {
          return { matches: [], count: 0 };
        }

        for (let r = 0; r < rows; r++) {
          for (let c = 0; c < cols; c++) {
            const cell = range.getCell(r, c);
            const font = cell.format.font;
            font.load(propsToLoad);
            cellRefs.push({ r, c, font });
          }
        }
        await ctx.sync();

        for (const ref of cellRefs) {
          if (matches.length >= MAX_MATCHES) break;
          let isMatch = true;
          if (fontQuery.bold !== undefined && ref.font.bold !== fontQuery.bold) isMatch = false;
          if (fontQuery.italic !== undefined && ref.font.italic !== fontQuery.italic) isMatch = false;
          if (fontQuery.color !== undefined) {
            const fontColor = (ref.font.color || "").toUpperCase();
            const queryColor = String(fontQuery.color).toUpperCase();
            if (fontColor !== queryColor) isMatch = false;
          }
          if (isMatch) {
            matches.push({
              address: OfficeService.getCellAddress(startRow + ref.r, startCol + ref.c),
              value: values[ref.r][ref.c],
            });
          }
        }
      }

      return { matches, count: matches.length };
    });
  }
}
