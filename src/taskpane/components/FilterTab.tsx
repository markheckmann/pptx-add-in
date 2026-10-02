import * as React from "react";
import { useEffect, useState } from "react";
import { Field, tokens, makeStyles } from "@fluentui/react-components";
import { ExcelDataState } from "../useExcelData";
import { FilterRule } from "../excelImport";
import { loadSlideFilters, saveSlideFilters } from "../slideFilters";
import FilterRuleEditor from "./FilterRuleEditor";

interface FilterTabProps {
  excelData: ExcelDataState;
  globalFilters: FilterRule[];
  onGlobalFiltersChange: (filters: FilterRule[]) => void;
  currentSlideId: string | null;
}

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    padding: "16px",
    gap: "24px",
  },
  hint: {
    color: tokens.colorNeutralForeground3,
    padding: "16px",
    textAlign: "center",
  },
  field: {
    width: "80%",
  },
});

const FilterTab: React.FC<FilterTabProps> = ({ excelData, globalFilters, onGlobalFiltersChange, currentSlideId }) => {
  const styles = useStyles();
  const { grid, hasHeaderRow } = excelData;

  const [slideFilters, setSlideFilters] = useState<FilterRule[]>([]);
  const [error, setError] = useState<string | null>(null);

  // Follows the currently active slide automatically — no manual slide picker.
  useEffect(() => {
    if (!currentSlideId) {
      setSlideFilters([]);
      return;
    }
    loadSlideFilters(currentSlideId)
      .then(setSlideFilters)
      .catch((e) => {
        console.error("Failed to load slide filters", e);
        setError(e instanceof Error ? e.message : String(e));
      });
  }, [currentSlideId]);

  const handleSlideFiltersChange = (next: FilterRule[]) => {
    setSlideFilters(next);
    if (!currentSlideId) return;
    saveSlideFilters(currentSlideId, next).catch((e) => {
      console.error("Failed to save slide filters", e);
      setError(e instanceof Error ? e.message : String(e));
    });
  };

  if (excelData.sheetNames.length === 0) {
    return <div className={styles.hint}>Bitte lade zuerst im Tab "Daten" eine Excel-Datei.</div>;
  }

  return (
    <div className={styles.container}>
      <Field className={styles.field} label="Globaler Filter (alle Grafiken)">
        <FilterRuleEditor
          grid={grid}
          hasHeaderRow={hasHeaderRow}
          filters={globalFilters}
          onFiltersChange={onGlobalFiltersChange}
        />
      </Field>

      <Field className={styles.field} label="Filter für diese Folie">
        {currentSlideId ? (
          <FilterRuleEditor
            grid={grid}
            hasHeaderRow={hasHeaderRow}
            filters={slideFilters}
            onFiltersChange={handleSlideFiltersChange}
          />
        ) : (
          <div className={styles.hint}>Keine Folie erkannt.</div>
        )}
      </Field>

      {error && <div className={styles.hint}>{error}</div>}
    </div>
  );
};

export default FilterTab;
