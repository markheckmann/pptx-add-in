/* global Office window console */

import { SheetGrid } from "./excelImport";

/**
 * Opens the raw-data table in a separate (larger) dialog window and sends it
 * the current grid once the dialog signals it's ready to receive it.
 */
export function openDataDialog(
  grid: SheetGrid,
  hasHeaderRow: boolean,
  onError: (message: string) => void
): void {
  const url = `${window.location.origin}/dialog.html`;

  Office.context.ui.displayDialogAsync(url, { height: 70, width: 70 }, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      onError(`Fenster konnte nicht geöffnet werden: ${asyncResult.error.message}`);
      return;
    }

    const dialog = asyncResult.value;
    dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
      try {
        const message = JSON.parse((arg as { message: string }).message) as { type: string };
        if (message.type === "ready") {
          dialog.messageChild(JSON.stringify({ type: "data", grid, hasHeaderRow }));
        }
      } catch (e) {
        console.error("Failed to handle dialog message", e);
      }
    });
  });
}
