// ============================================================
// FormulaBridge – Formula Watcher (Excel Event Handler)
// ============================================================
// Listens for cell changes and translates formulas in real-time.
// ============================================================

import { translateFormula } from "../engine/translator.js";
import { Dictionary, LanguageCode } from "../engine/types.js";
import dictionaryData from "../data/dictionary.json";

const dictionary = dictionaryData as unknown as Dictionary;

// Default target language – will be updated by locale detector
let targetLanguage: LanguageCode = "en";
let isEnabled = true;

/**
 * Initialize the formula watcher on the active worksheet.
 */
export async function initFormulaWatcher(): Promise<void> {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.onChanged.add(handleCellChanged);
      await context.sync();
      console.log("🌉 FormulaBridge: Formula watcher active");
    });
  } catch (error) {
    console.error("🌉 FormulaBridge: Failed to initialize watcher:", error);
  }
}

/**
 * Handle a cell change event.
 * This is the core reactive loop:
 *   1. Check if the change was made by us (avoid infinite loops)
 *   2. Read the formula from the changed cell
 *   3. Detect language and translate
 *   4. Write the translated formula back
 */
async function handleCellChanged(event: Excel.WorksheetChangedEventArgs): Promise<void> {
  if (!isEnabled) return;

  // Avoid infinite loops: ignore changes triggered by this add-in
  if (event.triggerSource === Excel.EventTriggerSource.thisLocalAddin) {
    return;
  }

  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(event.worksheetId);
      const range = sheet.getRange(event.address);

      range.load("formulas");
      await context.sync();

      const formula = range.formulas[0][0] as string;

      // Only process formulas (strings starting with =)
      if (typeof formula !== "string" || !formula.startsWith("=")) {
        return;
      }

      // Translate the formula
      const result = translateFormula(formula, dictionary, targetLanguage);

      if (!result.wasTranslated) {
        return; // Nothing to translate
      }

      // Disable events to prevent recursive triggering
      context.runtime.enableEvents = false;
      await context.sync();

      // Write the translated formula
      range.formulas = [[result.translated]];
      await context.sync();

      // Re-enable events
      context.runtime.enableEvents = true;
      await context.sync();

      console.log(
        `🌉 FormulaBridge: [${event.address}] ${result.original} → ${result.translated} ` +
        `(${result.sourceLanguage} → ${result.targetLanguage}, confidence: ${(result.confidence * 100).toFixed(0)}%)`
      );
    });
  } catch (error) {
    console.error("🌉 FormulaBridge: Translation error:", error);
    // Make sure events are re-enabled even if we error
    try {
      await Excel.run(async (context) => {
        context.runtime.enableEvents = true;
        await context.sync();
      });
    } catch {
      // Ignore – best effort
    }
  }
}

/**
 * Set the target language for translations.
 */
export function setTargetLanguage(lang: LanguageCode): void {
  targetLanguage = lang;
}

/**
 * Enable or disable the formula watcher.
 */
export function setEnabled(enabled: boolean): void {
  isEnabled = enabled;
}
