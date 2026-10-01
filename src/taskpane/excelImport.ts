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
