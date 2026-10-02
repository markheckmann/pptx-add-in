import { useState } from "react";
import type * as XLSX from "xlsx";
import { getSheetGrid, getSheetNames, parseWorkbook, SheetGrid } from "./excelImport";

export interface ExcelDataState {
  sheetNames: string[];
  selectedSheet: string | null;
  grid: SheetGrid;
  hasHeaderRow: boolean;
  loadWorkbookFromFile: (file: File) => Promise<void>;
  selectSheet: (sheetName: string) => void;
  setHasHeaderRow: (checked: boolean) => void;
}

/* global FileReader File */

export function useExcelData(): ExcelDataState {
  const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [selectedSheet, setSelectedSheet] = useState<string | null>(null);
  const [grid, setGrid] = useState<SheetGrid>([]);
  const [hasHeaderRow, setHasHeaderRow] = useState(true);

  const applySheetSelection = (wb: XLSX.WorkBook, sheetName: string) => {
    setSelectedSheet(sheetName);
    setGrid(getSheetGrid(wb, sheetName));
  };

  const loadWorkbookFromFile = (file: File): Promise<void> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => {
        try {
          const wb = parseWorkbook(reader.result as ArrayBuffer);
          const names = getSheetNames(wb);
          setWorkbook(wb);
          setSheetNames(names);
          if (names.length > 0) {
            applySheetSelection(wb, names[0]);
          } else {
            setSelectedSheet(null);
            setGrid([]);
          }
          resolve();
        } catch {
          reject(new Error("Die Datei konnte nicht als Excel-Arbeitsmappe gelesen werden."));
        }
      };
      reader.onerror = () => reject(new Error("Die Datei konnte nicht gelesen werden."));
      reader.readAsArrayBuffer(file);
    });
  };

  return {
    sheetNames,
    selectedSheet,
    grid,
    hasHeaderRow,
    loadWorkbookFromFile,
    selectSheet: (sheetName: string) => workbook && applySheetSelection(workbook, sheetName),
    setHasHeaderRow,
  };
}
