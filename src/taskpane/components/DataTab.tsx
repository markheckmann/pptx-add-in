import * as React from "react";
import { useRef, useState } from "react";
import { Button, Checkbox, Dropdown, Field, Option, tokens, makeStyles } from "@fluentui/react-components";
import { ExcelDataState } from "../useExcelData";
import { openDataDialog } from "../dialogController";

/* global HTMLInputElement localStorage */

const LAST_FILE_NAME_KEY = "surveyStudio.lastExcelFileName";

interface DataTabProps {
  excelData: ExcelDataState;
}

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    padding: "16px",
  },
  field: {
    width: "80%",
    marginTop: "12px",
  },
  fileName: {
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
    marginTop: "8px",
  },
  error: {
    color: tokens.colorPaletteRedForeground1,
    marginTop: "10px",
  },
});

const DataTab: React.FC<DataTabProps> = ({ excelData }) => {
  const styles = useStyles();
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [fileName, setFileName] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [lastFileName] = useState<string | null>(() => {
    try {
      return localStorage.getItem(LAST_FILE_NAME_KEY);
    } catch {
      return null;
    }
  });

  const handleFileChange = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;
    setError(null);
    try {
      await excelData.loadWorkbookFromFile(file);
      setFileName(file.name);
      try {
        localStorage.setItem(LAST_FILE_NAME_KEY, file.name);
      } catch {
        // ignore storage failures (private browsing, quota, ...) — it's just a convenience hint
      }
    } catch (e) {
      setFileName(null);
      setError(e instanceof Error ? e.message : String(e));
    }
  };

  return (
    <div className={styles.container}>
      <input
        ref={fileInputRef}
        type="file"
        accept=".xlsx,.xls"
        style={{ display: "none" }}
        onChange={handleFileChange}
      />
      <Button appearance="primary" onClick={() => fileInputRef.current?.click()}>
        Excel-Datei auswählen
      </Button>
      {fileName && <div className={styles.fileName}>Geladen: {fileName}</div>}
      {!fileName && lastFileName && (
        <div className={styles.fileName}>Zuletzt verwendet: {lastFileName} — bitte erneut auswählen.</div>
      )}

      {excelData.sheetNames.length > 0 && (
        <>
          <Field className={styles.field} label="Arbeitsblatt">
            <Dropdown
              value={excelData.selectedSheet ?? ""}
              selectedOptions={excelData.selectedSheet ? [excelData.selectedSheet] : []}
              onOptionSelect={(_, data) => data.optionValue && excelData.selectSheet(data.optionValue)}
            >
              {excelData.sheetNames.map((name) => (
                <Option key={name} value={name}>
                  {name}
                </Option>
              ))}
            </Dropdown>
          </Field>

          <Field className={styles.field}>
            <Checkbox
              label="Erste Zeile ist Überschrift"
              checked={excelData.hasHeaderRow}
              onChange={(_, data) => excelData.setHasHeaderRow(!!data.checked)}
            />
          </Field>

          <Button
            appearance="secondary"
            style={{ marginTop: "12px" }}
            onClick={() => openDataDialog(excelData.grid, excelData.hasHeaderRow, setError)}
          >
            Rohdaten in neuem Fenster anzeigen
          </Button>
        </>
      )}

      {error && <div className={styles.error}>{error}</div>}
    </div>
  );
};

export default DataTab;
