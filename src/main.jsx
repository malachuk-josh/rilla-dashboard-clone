import React from "react";
import { createRoot } from "react-dom/client";
import RillaUsageDashboardApp from "./RillaUsageDashboardApp.jsx";
import "./styles.css";

createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <RillaUsageDashboardApp />
  </React.StrictMode>,
);
