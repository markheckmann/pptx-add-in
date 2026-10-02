import * as React from "react";
import { useEffect, useState } from "react";
import { Button, Combobox, Field, Input, Option, tokens, makeStyles } from "@fluentui/react-components";
import { ArrowDown16Regular, ArrowUp16Regular } from "@fluentui/react-icons";
import { ExcelDataState } from "../useExcelData";
import { applyFilters, buildFrequencyData, columnLabel, FilterRule, getColumnCount } from "../excelImport";
import { orderBarData, updateChart } from "../chartToSlide";
import { ChartFontSizes, ChartTypeId } from "../chartRegistry";
import { getSlideFiltersForShape } from "../slideFilters";
import { ensureUniqueChartId } from "../chartIds";
import FilterRuleEditor from "./FilterRuleEditor";

/* global HTMLInputElement */

interface ChartEditTabProps {
  excelData: ExcelDataState;
  globalFilters: FilterRule[];
  editingShapeName: string | null;
  onShapeRenamed: (newShapeName: string) => void;
  chartType: ChartTypeId;
  selectedColumn: number | null;
  onSelectColumn: (index: number) => void;
  categoryOrder: string[];
  onCategoryOrderChange: (order: string[]) => void;
  categoryColors: Record<string, string>;
  onCategoryColorsChange: (colors: Record<string, string>) => void;
  chartFilters: FilterRule[];
  onChartFiltersChange: (filters: FilterRule[]) => void;
  fontSizes: ChartFontSizes;
  onFontSizesChange: (fontSizes: ChartFontSizes) => void;
  color: string;
  onColorChange: (color: string) => void;
  name: string;
  onNameChange: (name: string) => void;
  title: string;
  onTitleChange: (title: string) => void;
}

function moveItem<T>(items: T[], from: number, to: number): T[] {
  const next = [...items];
  const [item] = next.splice(from, 1);
  next.splice(to, 0, item);
  return next;
}

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
  field: {
    width: "80%",
    marginTop: "16px",
  },
  error: {
    color: tokens.colorPaletteRedForeground1,
    marginTop: "10px",
  },
  orderList: {
    listStyle: "none",
    padding: 0,
    margin: 0,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium,
  },
  orderItem: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    padding: "4px 8px",
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    "&:last-child": {
      borderBottom: "none",
    },
  },
});

const ChartEditTab: React.FC<ChartEditTabProps> = ({
  excelData,
  globalFilters,
  editingShapeName,
  onShapeRenamed,
  chartType,
  selectedColumn,
  onSelectColumn,
  categoryOrder,
  onCategoryOrderChange,
  categoryColors,
  onCategoryColorsChange,
  chartFilters,
  onChartFiltersChange,
  fontSizes,
  onFontSizesChange,
  color,
  onColorChange,
  name,
  onNameChange,
  title,
  onTitleChange,
}) => {
  const styles = useStyles();
  const [columnQuery, setColumnQuery] = useState("");
  const [error, setError] = useState<string | null>(null);
  const [slideFilters, setSlideFilters] = useState<FilterRule[]>([]);

  const { grid, hasHeaderRow } = excelData;
  const effectiveFilters = [...globalFilters, ...slideFilters, ...chartFilters];

  // The slide a chart lives on can have its own filter, fetched live (it's
  // stored on the slide itself, not in local state shared with other tabs).
  useEffect(() => {
    if (!editingShapeName) {
      setSlideFilters([]);
      return;
    }
    getSlideFiltersForShape(editingShapeName)
      .then(setSlideFilters)
      .catch((e) => console.error("Failed to load slide filters", e));
  }, [editingShapeName]);

  useEffect(() => {
    if (selectedColumn !== null) {
      setColumnQuery(columnLabel(grid, selectedColumn, hasHeaderRow));
    }
  }, [selectedColumn, grid, hasHeaderRow]);

  // Reset the category order to an alphabetical default whenever the
  // selected variable's category set actually changes (new column, or data
  // changed). `categoryOrder` is deliberately not a dependency (it's also
  // the thing this effect writes); a restore sets selectedColumn and
  // categoryOrder together, so by the time this runs categoryOrder already
  // matches and sameSet is true, leaving the restored order untouched.
  useEffect(() => {
    if (selectedColumn === null) return;
    const filteredGrid = applyFilters(grid, effectiveFilters, hasHeaderRow);
    const categories = buildFrequencyData(filteredGrid, selectedColumn, hasHeaderRow).map((d) => d.category);
    const sameSet =
      categoryOrder.length === categories.length && categories.every((c) => categoryOrder.includes(c));
    if (!sameSet) {
      onCategoryOrderChange([...categories].sort((a, b) => a.localeCompare(b, "de")));
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedColumn, grid, hasHeaderRow, globalFilters, slideFilters, chartFilters]);

  if (!editingShapeName) {
    return <div className={styles.hint}>Noch keine Grafik ausgewählt. Erstelle eine im Tab "Grafiken".</div>;
  }

  const columnCount = getColumnCount(grid);
  const allColumnOptions = Array.from({ length: columnCount }, (_, i) => ({
    index: i,
    label: columnLabel(grid, i, hasHeaderRow),
  }));
  const filteredColumnOptions = allColumnOptions.filter((col) =>
    col.label.toLowerCase().includes(columnQuery.toLowerCase())
  );

  const handleUpdateClick = async () => {
    if (selectedColumn === null) return;
    try {
      setError(null);
      const filteredGrid = applyFilters(grid, effectiveFilters, hasHeaderRow);
      const data = orderBarData(buildFrequencyData(filteredGrid, selectedColumn, hasHeaderRow), categoryOrder);
      if (data.length === 0) {
        setError("Keine Werte in der gewählten Spalte gefunden.");
        return;
      }
      // Guard against the user typing an ID that another chart (or any
      // other shape) already uses — auto-append a suffix rather than
      // silently causing a name collision.
      const desiredId = name.trim() || editingShapeName;
      const finalId = await ensureUniqueChartId(desiredId, editingShapeName);

      const newShapeName = await updateChart(editingShapeName, data, {
        chartType,
        columnIndex: selectedColumn,
        color,
        name: finalId,
        title,
        categoryOrder,
        categoryColors,
        chartFilters,
        fontSizes,
      });
      if (newShapeName !== editingShapeName) {
        onShapeRenamed(newShapeName);
      }
      onNameChange(newShapeName);
    } catch (e) {
      console.error("Chart update failed", e);
      setError(`Diagramm konnte nicht aktualisiert werden: ${e instanceof Error ? e.message : String(e)}`);
    }
  };

  return (
    <div className={styles.container}>
      <Field className={styles.field} label="Variable">
        <Combobox
          placeholder="Spalte suchen..."
          value={columnQuery}
          selectedOptions={selectedColumn !== null ? [String(selectedColumn)] : []}
          onInput={(e) => setColumnQuery((e.target as HTMLInputElement).value)}
          onOptionSelect={(_, data) => {
            if (data.optionValue === undefined) return;
            onSelectColumn(Number(data.optionValue));
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

      <Field className={styles.field} label="Nur für diese Grafik">
        <FilterRuleEditor
          grid={grid}
          hasHeaderRow={hasHeaderRow}
          filters={chartFilters}
          onFiltersChange={onChartFiltersChange}
        />
      </Field>

      {selectedColumn !== null && categoryOrder.length > 0 && (
        <Field className={styles.field} label="Reihenfolge der Kategorien">
          <ul className={styles.orderList}>
            {categoryOrder.map((category, index) => (
              <li key={category} className={styles.orderItem}>
                <span>{category}</span>
                <span>
                  <input
                    type="color"
                    value={categoryColors[category] ?? color}
                    onChange={(e) => onCategoryColorsChange({ ...categoryColors, [category]: e.target.value })}
                  />
                  <Button
                    appearance="subtle"
                    icon={<ArrowUp16Regular />}
                    disabled={index === 0}
                    onClick={() => onCategoryOrderChange(moveItem(categoryOrder, index, index - 1))}
                  />
                  <Button
                    appearance="subtle"
                    icon={<ArrowDown16Regular />}
                    disabled={index === categoryOrder.length - 1}
                    onClick={() => onCategoryOrderChange(moveItem(categoryOrder, index, index + 1))}
                  />
                </span>
              </li>
            ))}
          </ul>
        </Field>
      )}

      <Field className={styles.field} label="ID (Name, wird nicht angezeigt)">
        <Input value={name} onChange={(_, data) => onNameChange(data.value)} />
      </Field>

      <Field className={styles.field} label="Titel (im Diagramm sichtbar)">
        <Input value={title} onChange={(_, data) => onTitleChange(data.value)} placeholder="Diagrammtitel" />
      </Field>

      <Field
        className={styles.field}
        label={chartType === "bar" ? "Standard-Balkenfarbe (ohne eigene Farbe oben)" : "Standardfarbe (ohne eigene Farbe oben)"}
      >
        <input type="color" value={color} onChange={(e) => onColorChange(e.target.value)} />
      </Field>

      <Field className={styles.field} label="Schriftgröße Überschrift">
        <Input
          type="number"
          min={6}
          max={72}
          value={String(fontSizes.title)}
          onChange={(_, data) => onFontSizesChange({ ...fontSizes, title: Number(data.value) || fontSizes.title })}
        />
      </Field>

      <Field
        className={styles.field}
        label={
          chartType === "bar"
            ? "Schriftgröße X-Achse"
            : chartType === "donut"
              ? "Schriftgröße Segment-Beschriftung"
              : "Schriftgröße Kachel-Beschriftung"
        }
      >
        <Input
          type="number"
          min={6}
          max={72}
          value={String(fontSizes.xLabels)}
          onChange={(_, data) =>
            onFontSizesChange({ ...fontSizes, xLabels: Number(data.value) || fontSizes.xLabels })
          }
        />
      </Field>

      {chartType === "bar" && (
        <Field className={styles.field} label="Schriftgröße Y-Achse">
          <Input
            type="number"
            min={6}
            max={72}
            value={String(fontSizes.yLabels)}
            onChange={(_, data) =>
              onFontSizesChange({ ...fontSizes, yLabels: Number(data.value) || fontSizes.yLabels })
            }
          />
        </Field>
      )}

      <Button
        appearance="primary"
        size="large"
        disabled={selectedColumn === null}
        onClick={handleUpdateClick}
        style={{ marginTop: "16px" }}
      >
        Aktualisieren
      </Button>
      {selectedColumn === null && <div className={styles.hint}>Wähle zuerst eine Variable aus.</div>}

      {error && <div className={styles.error}>{error}</div>}
    </div>
  );
};

export default ChartEditTab;
