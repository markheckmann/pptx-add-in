import * as React from "react";
import { Button, Checkbox, Dropdown, Option, tokens, makeStyles } from "@fluentui/react-components";
import { Delete16Regular } from "@fluentui/react-icons";
import { SheetGrid, columnLabel, getColumnCount, getDistinctValues, FilterRule } from "../excelImport";

interface FilterRuleEditorProps {
  grid: SheetGrid;
  hasHeaderRow: boolean;
  filters: FilterRule[];
  onFiltersChange: (filters: FilterRule[]) => void;
}

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    gap: "12px",
    width: "100%",
  },
  rule: {
    width: "100%",
    boxSizing: "border-box",
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium,
    padding: "12px",
  },
  ruleHeader: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: "8px",
  },
  valueList: {
    maxHeight: "160px",
    overflowY: "auto",
    marginTop: "8px",
    display: "flex",
    flexDirection: "column",
  },
});

/**
 * Reusable column + checkbox-list filter-rule editor. Used for the global
 * filter tab, a slide's filter, and a single chart's own filter — all three
 * share the exact same FilterRule[] shape and editing UI.
 */
const FilterRuleEditor: React.FC<FilterRuleEditorProps> = ({ grid, hasHeaderRow, filters, onFiltersChange }) => {
  const styles = useStyles();

  const columnCount = getColumnCount(grid);
  const columnOptions = Array.from({ length: columnCount }, (_, i) => ({
    index: i,
    label: columnLabel(grid, i, hasHeaderRow),
  }));

  const updateRule = (index: number, rule: FilterRule) => {
    const next = [...filters];
    next[index] = rule;
    onFiltersChange(next);
  };

  const removeRule = (index: number) => {
    onFiltersChange(filters.filter((_, i) => i !== index));
  };

  const handleAddFilter = () => {
    // All values start checked, so a freshly added filter excludes nothing
    // until the user unchecks something.
    onFiltersChange([...filters, { columnIndex: 0, selectedValues: getDistinctValues(grid, 0, hasHeaderRow) }]);
  };

  return (
    <div className={styles.container}>
      {filters.map((rule, index) => {
        const values = getDistinctValues(grid, rule.columnIndex, hasHeaderRow);
        return (
          // eslint-disable-next-line react/no-array-index-key
          <div key={index} className={styles.rule}>
            <div className={styles.ruleHeader}>
              <Dropdown
                value={columnLabel(grid, rule.columnIndex, hasHeaderRow)}
                selectedOptions={[String(rule.columnIndex)]}
                onOptionSelect={(_, data) => {
                  if (data.optionValue === undefined) return;
                  const columnIndex = Number(data.optionValue);
                  updateRule(index, {
                    columnIndex,
                    selectedValues: getDistinctValues(grid, columnIndex, hasHeaderRow),
                  });
                }}
              >
                {columnOptions.map((col) => (
                  <Option key={col.index} value={String(col.index)}>
                    {col.label}
                  </Option>
                ))}
              </Dropdown>
              <Button appearance="subtle" icon={<Delete16Regular />} onClick={() => removeRule(index)} />
            </div>

            <div className={styles.valueList}>
              {values.map((value) => (
                <Checkbox
                  key={value}
                  label={value}
                  checked={rule.selectedValues.includes(value)}
                  onChange={(_, data) => {
                    const selectedValues = data.checked
                      ? [...rule.selectedValues, value]
                      : rule.selectedValues.filter((v) => v !== value);
                    updateRule(index, { ...rule, selectedValues });
                  }}
                />
              ))}
            </div>
          </div>
        );
      })}

      <Button appearance="primary" onClick={handleAddFilter}>
        Filter hinzufügen
      </Button>
    </div>
  );
};

export default FilterRuleEditor;
