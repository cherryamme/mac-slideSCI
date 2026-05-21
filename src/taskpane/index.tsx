import React from "react";
import { createRoot } from "react-dom/client";
import { App } from "./App";
import type { OfficeReadyInfo } from "../shared/types";
import "./styles.css";

const rootElement = document.getElementById("root");

if (!rootElement) {
  throw new Error("Root element was not found.");
}

const root = createRoot(rootElement);

function render(officeInfo?: OfficeReadyInfo): void {
  root.render(
    <React.StrictMode>
      <App officeInfo={officeInfo} />
    </React.StrictMode>,
  );
}

if (typeof Office === "undefined") {
  render();
} else {
  Office.onReady((info) => render(info)).catch(() => render());
}
