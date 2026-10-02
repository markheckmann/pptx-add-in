import * as React from "react";
import { Button, Field, tokens, makeStyles } from "@fluentui/react-components";
import { ExcelDataState } from "../useExcelData";
import { ChartTypeId } from "../chartRegistry";

interface GrafikenTabProps {
  excelData: ExcelDataState;
  onCreatePlaceholder: (chartType: ChartTypeId) => void;
}

const CHART_TYPES: { id: ChartTypeId; label: string }[] = [
  { id: "bar", label: "Balkendiagramm" },
  { id: "donut", label: "Donut-Diagramm" },
  { id: "treemap", label: "Treemap" },
];

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    padding: "16px",
  },
  hint: {
    color: tokens.colorNeutralForeground3,
    padding: "16px",
    textAlign: "center",
  },
  typeList: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    width: "80%",
  },
});

const GrafikenTab: React.FC<GrafikenTabProps> = ({ excelData, onCreatePlaceholder }) => {
  const styles = useStyles();

  if (excelData.sheetNames.length === 0) {
    return <div className={styles.hint}>Bitte lade zuerst im Tab "Daten" eine Excel-Datei.</div>;
  }

  return (
    <div className={styles.container}>
      <Field className={styles.typeList} label="Diagrammtyp">
        {CHART_TYPES.map((type) => (
          <Button key={type.id} appearance="secondary" onClick={() => onCreatePlaceholder(type.id)}>
            {type.label}
          </Button>
        ))}
      </Field>
    </div>
  );
};

export default GrafikenTab;
