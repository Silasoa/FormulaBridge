// ============================================================
// FormulaBridge – Main Entry Point
// ============================================================

import { initFormulaWatcher } from "./excel/formulaWatcher.js";

// Wait for Office.js to be ready before initializing
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    console.log("🌉 FormulaBridge: Office.js ready, host is Excel");
    initUI();
    initFormulaWatcher();
  } else {
    console.warn("🌉 FormulaBridge: Not running in Excel, host:", info.host);
  }
});

function initUI(): void {
  const statusDot = document.querySelector(".fb-status-dot");
  const statusText = document.querySelector(".fb-status-text");

  if (statusDot && statusText) {
    statusDot.classList.remove("fb-status-dot--inactive");
    statusDot.classList.add("fb-status-dot--active");
    statusText.textContent = "Active";
  }

  console.log("🌉 FormulaBridge: UI initialized");
}
