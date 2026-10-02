import * as React from "react";
import { createRoot } from "react-dom/client";
import DataDialog from "./DataDialog";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";

/* global document, Office, HTMLElement */

const rootElement: HTMLElement | null = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

Office.onReady(() => {
  root?.render(
    <FluentProvider theme={webLightTheme}>
      <DataDialog />
    </FluentProvider>
  );
});
