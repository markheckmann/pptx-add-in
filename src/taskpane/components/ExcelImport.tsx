import * as React from "react";
import { useRef, useState } from "react";
import { Button, Checkbox, Combobox, Dropdown, Field, Option, tokens, makeStyles } from "@fluentui/react-components";
import type * as XLSX from "xlsx";
import { buildFrequencyData, columnLabel, getColumnCount, getSheetGrid, getSheetNames, parseWorkbook, SheetGrid } from "../excelImport";
import { insertBarChart, nextInsertIndex } from "../chartToSlide";

/* global HTMLInputElement FileReader */

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    marginTop: "24px",
    paddingTop: "16px",
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  instructions: {
    fontWeight: tokens.fontWeightSemibold,
    marginBottom: "10px",
  },
  field: {
    width: "80%",
    marginBottom: "12px",
  },
  error: {
    color: tokens.colorPaletteRedForeground1,
    marginTop: "10px",
  },
  hint: {
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
    marginTop: "4px",
  },
});

const ExcelImport: React.FC = () => {
  const styles = useStyles();
  const fileInputRef = useRef<HTMLInputElement>(null);

  const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [selectedSheet, setSelectedSheet] = useState<string | null>(null);
  const [grid, setGrid] = useState<SheetGrid>([]);
  const [hasHeaderRow, setHasHeaderRow] = useState(true);
  const [selectedColumn, setSelectedColumn] = useState<number | null>(null);
  const [columnQuery, setColumnQuery] = useState("");
  const [error, setError] = useState<string | null>(null);

  const loadSheet = (wb: XLSX.WorkBook, sheetName: string) => {
    const sheetGrid = getSheetGrid(wb, sheetName);
    setSelectedSheet(sheetName);
    setGrid(sheetGrid);
    setSelectedColumn(null);
    setColumnQuery("");
  };

  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;
    setError(null);

    const reader = new FileReader();
    reader.onload = () => {
      try {
        const wb = parseWorkbook(reader.result as ArrayBuffer);
        const names = getSheetNames(wb);
        setWorkbook(wb);
        setSheetNames(names);
        if (names.length > 0) loadSheet(wb, names[0]);
      } catch {
        setError("Die Datei konnte nicht als Excel-Arbeitsmappe gelesen werden.");
      }
    };
    reader.onerror = () => setError("Die Datei konnte nicht gelesen werden.");
    reader.readAsArrayBuffer(file);
  };

  const handleSheetSelect = (sheetName: string) => {
    if (!workbook) return;
    loadSheet(workbook, sheetName);
  };

  const columnCount = getColumnCount(grid);
  const allColumnOptions = Array.from({ length: columnCount }, (_, i) => ({
    index: i,
    label: columnLabel(grid, i, hasHeaderRow),
  }));
  const filteredColumnOptions = allColumnOptions.filter((col) =>
    col.label.toLowerCase().includes(columnQuery.toLowerCase())
  );
  const selectedColumnLabel = selectedColumn !== null ? columnLabel(grid, selectedColumn, hasHeaderRow) : "";

  const handleInsertClick = async () => {
    if (selectedColumn === null) return;
    try {
      setError(null);
      const data = buildFrequencyData(grid, selectedColumn, hasHeaderRow);
      if (data.length === 0) {
        setError("Keine Werte in der gewählten Spalte gefunden.");
        return;
      }
      await insertBarChart(data, nextInsertIndex());
    } catch (e) {
      setError(`Diagramm konnte nicht eingefügt werden: ${e instanceof Error ? e.message : String(e)}`);
    }
  };

  return (
    <div className={styles.container}>
      <Field className={styles.instructions}>Häufigkeit einer Spalte aus Excel als Diagramm einfügen.</Field>

      <input
        ref={fileInputRef}
        type="file"
        accept=".xlsx,.xls"
        style={{ display: "none" }}
        onChange={handleFileChange}
      />
      <Button appearance="secondary" onClick={() => fileInputRef.current?.click()}>
        Excel-Datei auswählen
      </Button>

      {sheetNames.length > 0 && (
        <>
          <Field className={styles.field} label="Arbeitsblatt">
            <Dropdown
              value={selectedSheet ?? ""}
              selectedOptions={selectedSheet ? [selectedSheet] : []}
              onOptionSelect={(_, data) => data.optionValue && handleSheetSelect(data.optionValue)}
            >
              {sheetNames.map((name) => (
                <Option key={name} value={name}>
                  {name}
                </Option>
              ))}
            </Dropdown>
          </Field>

          <Checkbox
            label="Erste Zeile ist Überschrift"
            checked={hasHeaderRow}
            onChange={(_, data) => setHasHeaderRow(!!data.checked)}
          />

          <Field className={styles.field} label="Spalte (Häufigkeit wird ausgezählt)">
            <Combobox
              placeholder="Spalte suchen..."
              value={columnQuery}
              selectedOptions={selectedColumn !== null ? [String(selectedColumn)] : []}
              onInput={(e) => setColumnQuery((e.target as HTMLInputElement).value)}
              onOptionSelect={(_, data) => {
                if (data.optionValue === undefined) return;
                setSelectedColumn(Number(data.optionValue));
                setColumnQuery(data.optionText ?? "");
              }}
            >
              {filteredColumnOptions.map((col) => (
                <Option key={col.index} value={String(col.index)} text={col.label}>
                  {col.label}
                </Option>
              ))}
            </Combobox>
          </Field>

          {selectedColumn !== null && (
            <div className={styles.hint}>Ausgewählt: {selectedColumnLabel}</div>
          )}

          <Button
            appearance="primary"
            size="large"
            disabled={selectedColumn === null}
            onClick={handleInsertClick}
          >
            Häufigkeits-Diagramm einfügen
          </Button>
        </>
      )}

      {error && <div className={styles.error}>{error}</div>}
    </div>
  );
};

export default ExcelImport;
