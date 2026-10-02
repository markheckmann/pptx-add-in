import * as React from "react";
import { useEffect, useState } from "react";
import {
  DataGrid,
  DataGridBody,
  DataGridRow,
  DataGridHeader,
  DataGridHeaderCell,
  DataGridCell,
  TableColumnDefinition,
  createTableColumn,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import { columnLabel, getColumnCount, SheetGrid } from "../taskpane/excelImport";

/* global Office */

interface DialogDataMessage {
  type: "data";
  grid: SheetGrid;
  hasHeaderRow: boolean;
}

interface RowItem {
  id: string;
  [columnIndex: string]: string;
}

const useStyles = makeStyles({
  container: {
    padding: "16px",
    height: "100vh",
    boxSizing: "border-box",
    display: "flex",
    flexDirection: "column",
  },
  title: {
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase500,
    marginBottom: "12px",
  },
  gridWrapper: {
    flex: 1,
    overflow: "auto",
  },
  hint: {
    color: tokens.colorNeutralForeground3,
  },
});

const DataDialog: React.FC = () => {
  const styles = useStyles();
  const [grid, setGrid] = useState<SheetGrid | null>(null);
  const [hasHeaderRow, setHasHeaderRow] = useState(true);

  useEffect(() => {
    Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (arg) => {
      try {
        const payload = (arg as { message: string }).message;
        const message = JSON.parse(payload) as DialogDataMessage;
        if (message.type === "data") {
          setGrid(message.grid);
          setHasHeaderRow(message.hasHeaderRow);
        }
      } catch (e) {
        console.error("Failed to parse dialog message", e);
      }
    });

    // Tell the task pane we're ready to receive the data.
    Office.context.ui.messageParent(JSON.stringify({ type: "ready" }));
  }, []);

  if (!grid) {
    return (
      <div className={styles.container}>
        <div className={styles.hint}>Lade Daten…</div>
      </div>
    );
  }

  const columnCount = getColumnCount(grid);
  const dataRows = hasHeaderRow ? grid.slice(1) : grid;

  const columns: TableColumnDefinition<RowItem>[] = Array.from({ length: columnCount }, (_, i) =>
    createTableColumn<RowItem>({
      columnId: String(i),
      renderHeaderCell: () => columnLabel(grid, i, hasHeaderRow),
      renderCell: (item) => item[String(i)] ?? "",
    })
  );

  const items: RowItem[] = dataRows.map((row, rowIndex) => {
    const item: RowItem = { id: String(rowIndex) };
    for (let i = 0; i < columnCount; i++) {
      item[String(i)] = row[i] ?? "";
    }
    return item;
  });

  return (
    <div className={styles.container}>
      <div className={styles.title}>Rohdaten ({items.length} Zeilen)</div>
      <div className={styles.gridWrapper}>
        <DataGrid items={items} columns={columns} getRowId={(item) => item.id}>
          <DataGridHeader>
            <DataGridRow>
              {({ renderHeaderCell }) => <DataGridHeaderCell>{renderHeaderCell()}</DataGridHeaderCell>}
            </DataGridRow>
          </DataGridHeader>
          <DataGridBody<RowItem>>
            {({ item, rowId }) => (
              <DataGridRow<RowItem> key={rowId}>
                {({ renderCell }) => <DataGridCell>{renderCell(item)}</DataGridCell>}
              </DataGridRow>
            )}
          </DataGridBody>
        </DataGrid>
      </div>
    </div>
  );
};

export default DataDialog;
