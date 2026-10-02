import * as XLSX from "xlsx";
import { BarDatum } from "./chartToSlide";

export type SheetGrid = string[][];

export function parseWorkbook(data: ArrayBuffer): XLSX.WorkBook {
  return XLSX.read(data, { type: "array" });
}

export function getSheetNames(workbook: XLSX.WorkBook): string[] {
  return workbook.SheetNames;
}

/** Returns the sheet as a grid of strings, one row per array entry. */
export function getSheetGrid(workbook: XLSX.WorkBook, sheetName: string): SheetGrid {
  const worksheet = workbook.Sheets[sheetName];
  if (!worksheet) return [];
  return XLSX.utils.sheet_to_json<string[]>(worksheet, { header: 1, raw: false, defval: "" });
}

export function columnLabel(grid: SheetGrid, columnIndex: number, hasHeaderRow: boolean): string {
  const headerValue = hasHeaderRow ? grid[0]?.[columnIndex] : undefined;
  if (headerValue && headerValue.trim().length > 0) return headerValue;
  return `Spalte ${columnIndex + 1}`;
}

export function getColumnCount(grid: SheetGrid): number {
  return grid.reduce((max, row) => Math.max(max, row.length), 0);
}

/** Returns the distinct, non-empty values of a column, alphabetically. */
export function getDistinctValues(
  grid: SheetGrid,
  columnIndex: number,
  hasHeaderRow: boolean
): string[] {
  const dataRows = hasHeaderRow ? grid.slice(1) : grid;
  const values = new Set<string>();
  for (const row of dataRows) {
    const value = (row[columnIndex] ?? "").trim();
    if (value.length > 0) values.add(value);
  }
  return Array.from(values).sort((a, b) => a.localeCompare(b, "de"));
}

export interface FilterRule {
  columnIndex: number;
  /** Rows are kept only if their value in this column is one of these. */
  selectedValues: string[];
}

/**
 * Keeps only the rows matching every filter rule (AND). A rule with no
 * selected values matches nothing for that column. An empty filter list
 * returns the grid unchanged.
 */
export function applyFilters(
  grid: SheetGrid,
  filters: FilterRule[],
  hasHeaderRow: boolean
): SheetGrid {
  if (filters.length === 0) return grid;
  const headerRows = hasHeaderRow ? grid.slice(0, 1) : [];
  const dataRows = hasHeaderRow ? grid.slice(1) : grid;
  const filteredRows = dataRows.filter((row) =>
    filters.every((f) => f.selectedValues.includes((row[f.columnIndex] ?? "").trim()))
  );
  return [...headerRows, ...filteredRows];
}

/**
 * Counts how often each distinct value occurs in the given column and
 * returns the counts as chart data. Empty cells are skipped. The chart
 * renderer displays categories in alphabetical order regardless of the
 * order returned here.
 */
export function buildFrequencyData(
  grid: SheetGrid,
  columnIndex: number,
  hasHeaderRow: boolean
): BarDatum[] {
  const dataRows = hasHeaderRow ? grid.slice(1) : grid;
  const counts = new Map<string, number>();
  for (const row of dataRows) {
    const value = (row[columnIndex] ?? "").trim();
    if (value.length === 0) continue;
    counts.set(value, (counts.get(value) ?? 0) + 1);
  }
  return Array.from(counts, ([category, value]) => ({ category, value }));
}
