// ============================================================
// FormulaBridge – Main Entry Point
// ============================================================

import { initFormulaWatcher, setTargetLanguage, setSourceLanguage, setEnabled, onFormulaTranslated, TranslationLog } from "./excel/formulaWatcher.js";
import { detectExcelLocale } from "./excel/localeDetector.js";
import { translateFormula } from "./engine/translator.js";
import { Dictionary, LanguageCode } from "./engine/types.js";
import dictionaryData from "./data/dictionary.json";

const dictionary = dictionaryData as unknown as Dictionary;

// Wait for Office.js to be ready before initializing
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    console.log("🌉 FormulaBridge: Office.js ready, host is Excel");
    initApp();
  } else {
    console.warn("🌉 FormulaBridge: Not running in Excel, host:", info.host);
    // Initialize anyway for browser testing
    initApp();
  }
});

let currentTargetLang: LanguageCode = "en";

function initApp(): void {
  // 1. Detect and set language
  const locale = detectExcelLocale();
  currentTargetLang = locale;
  setTargetLanguage(locale);
  console.log(`🌉 FormulaBridge: Detected Excel locale as '${locale}'`);

  // 2. Initialize UI & Event Listeners
  initUI();

  // 3. Start Excel Watcher
  initFormulaWatcher();
}

function initUI(): void {
  // Status Badge
  const statusDot = document.querySelector(".fb-status-dot");
  const statusText = document.querySelector(".fb-status-text");

  if (statusDot && statusText) {
    statusDot.classList.remove("fb-status-dot--inactive");
    statusDot.classList.add("fb-status-dot--active");
    statusText.textContent = "Active";
  }

  // Toggles
  const autoTranslateToggle = document.getElementById("auto-translate-toggle") as HTMLInputElement;
  if (autoTranslateToggle) {
    autoTranslateToggle.addEventListener("change", (e) => {
      const isChecked = (e.target as HTMLInputElement).checked;
      setEnabled(isChecked);
    });
  }

  const autoDetectToggle = document.getElementById("auto-detect-toggle") as HTMLInputElement;
  const sourceLangSelect = document.getElementById("source-language") as HTMLSelectElement;
  
  if (autoDetectToggle && sourceLangSelect) {
    // Populate source languages
    // Add English manually since it's the base key
    const enOption = document.createElement("fluent-option");
    enOption.setAttribute("value", "en");
    enOption.textContent = "English";
    sourceLangSelect.appendChild(enOption);

    dictionary.meta.languages.forEach((langCode) => {
      const localeInfo = dictionary.locales[langCode];
      if (localeInfo) {
        const option = document.createElement("fluent-option");
        option.setAttribute("value", langCode);
        option.textContent = localeInfo.name;
        sourceLangSelect.appendChild(option);
      }
    });

    autoDetectToggle.addEventListener("change", (e) => {
      const isChecked = (e.target as HTMLInputElement).checked;
      if (isChecked) {
        sourceLangSelect.setAttribute("disabled", "true");
        sourceLangSelect.value = "auto";
        setSourceLanguage("auto");
      } else {
        sourceLangSelect.removeAttribute("disabled");
        // default to English when disabled if no selection yet
        if (sourceLangSelect.value === "auto" || !sourceLangSelect.value) {
            sourceLangSelect.value = "en";
        }
        setSourceLanguage(sourceLangSelect.value as LanguageCode);
      }
      updateReferenceList(searchInput?.value);
    });

    sourceLangSelect.addEventListener("change", () => {
      if (!autoDetectToggle.checked) {
        setSourceLanguage(sourceLangSelect.value as LanguageCode);
      }
      updateReferenceList(searchInput?.value);
    });
  }

  // Copy-Paste Translator
  const cpInput = document.getElementById("copypaste-input") as HTMLInputElement;
  const cpInsertBtn = document.getElementById("copypaste-insert-btn") as HTMLButtonElement;
  if (cpInput && cpInsertBtn) {
    cpInsertBtn.addEventListener("click", async () => {
      let formula = cpInput.value.trim();
      if (!formula) return;
      
      // Auto-prefix with '=' if missing, for convenience
      if (!formula.startsWith("=")) {
        formula = "=" + formula;
      }

      const sourceLang = sourceLangSelect.value === "auto" ? undefined : sourceLangSelect.value as LanguageCode;
      const result = translateFormula(formula, dictionary, currentTargetLang, sourceLang);
      
      try {
        await Excel.run(async (context) => {
          const range = context.workbook.getSelectedRange();
          range.load("address");
          
          context.runtime.enableEvents = false;
          await context.sync();
          
          range.formulasLocal = [[result.translated]];
          await context.sync();
          
          context.runtime.enableEvents = true;
          await context.sync();
          
          console.log(`🌉 FormulaBridge: Inserted translated formula to ${range.address}`);
          cpInput.value = ""; // Clear on success
        });
      } catch (error) {
        console.error("🌉 FormulaBridge: Error inserting formula:", error);
        // Ensure events are re-enabled
        try {
          await Excel.run(async (ctx) => {
            ctx.runtime.enableEvents = true;
            await ctx.sync();
          });
        } catch(e) {}
      }
    });
  }

  // Formula Reference Search
  const searchInput = document.getElementById("reference-search") as HTMLInputElement;
  const referenceContainer = document.getElementById("formula-reference");

  function updateReferenceList(query: string = "") {
    if (!referenceContainer) return;
    
    referenceContainer.innerHTML = "";
    const q = query.toUpperCase();
    
    let displaySourceLang = sourceLangSelect?.value;
    if (displaySourceLang === "auto" || !displaySourceLang) {
      displaySourceLang = "en";
    }

    // Convert dictionary object to array
    const entries = Object.entries(dictionary.functions);
    
    // Filter
    const filtered = entries.filter(([enName, translations]) => {
      const sourceName = displaySourceLang === "en" ? enName : translations[displaySourceLang as LanguageCode];
      if (!sourceName) return false;

      const localName = currentTargetLang === "en" ? enName : translations[currentTargetLang];
      
      if (!q) return true;
      if (sourceName.toUpperCase().includes(q)) return true;
      if (localName && localName.toUpperCase().includes(q)) return true;
      
      return false;
    });

    // Take top 50 to avoid DOM overload
    const top = filtered.slice(0, 50);

    top.forEach(([enName, translations]) => {
      const sourceName = displaySourceLang === "en" ? enName : translations[displaySourceLang as LanguageCode];
      const localName = currentTargetLang === "en" ? enName : translations[currentTargetLang];
      
      const row = document.createElement("div");
      row.className = "fb-reference__row";
      row.innerHTML = `
        <div class="fb-reference__en">${sourceName}</div>
        <div class="fb-reference__local">${localName || "-"}</div>
      `;
      referenceContainer.appendChild(row);
    });

    if (filtered.length === 0) {
      referenceContainer.innerHTML = '<div style="padding: 10px; color: var(--fb-text-muted); font-size: var(--fb-font-size-xs);">No functions found.</div>';
    }
  }

  if (searchInput && referenceContainer) {
    searchInput.addEventListener("input", (e) => {
      updateReferenceList((e.target as HTMLInputElement).value);
    });
    // Initial populate
    updateReferenceList("");
  }

  // Translation Log
  const logContainer = document.getElementById("translation-log");
  const clearLogBtn = document.getElementById("clear-log-btn");

  if (clearLogBtn && logContainer) {
    clearLogBtn.addEventListener("click", () => {
      logContainer.innerHTML = '<div class="fb-log__empty">No translations yet. Start typing formulas!</div>';
    });
  }

  onFormulaTranslated((log: TranslationLog) => {
    if (!logContainer) return;
    
    const emptyMsg = logContainer.querySelector(".fb-log__empty");
    if (emptyMsg) {
      emptyMsg.remove();
    }

    const entry = document.createElement("div");
    entry.className = "fb-log__entry";
    
    entry.innerHTML = `
      <div class="fb-log__cell">Cell: ${log.address} | Source: ${log.sourceLanguage.toUpperCase()} | Confidence: ${(log.confidence * 100).toFixed(0)}%</div>
      <div class="fb-log__original">${log.original}</div>
      <div class="fb-log__translated">${log.translated}</div>
    `;

    logContainer.prepend(entry);

    if (logContainer.children.length > 50) {
      logContainer.lastElementChild?.remove();
    }
  });

  // Help Modal Logic
  const helpBtn = document.getElementById("help-btn") as HTMLButtonElement;
  const helpModal = document.getElementById("help-modal") as HTMLDialogElement;
  const closeHelpBtn = document.getElementById("close-help-btn") as HTMLButtonElement;

  if (helpBtn && helpModal && closeHelpBtn) {
    helpBtn.addEventListener("click", () => {
      helpModal.showModal();
    });
    
    closeHelpBtn.addEventListener("click", () => {
      helpModal.close();
    });

    // Close on backdrop click
    helpModal.addEventListener("click", (e) => {
      const dialogDimensions = helpModal.getBoundingClientRect();
      if (
        e.clientX < dialogDimensions.left ||
        e.clientX > dialogDimensions.right ||
        e.clientY < dialogDimensions.top ||
        e.clientY > dialogDimensions.bottom
      ) {
        helpModal.close();
      }
    });
  }

  console.log("🌉 FormulaBridge: UI initialized");
}
