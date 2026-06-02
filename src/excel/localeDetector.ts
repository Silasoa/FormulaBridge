// ============================================================
// FormulaBridge – Locale Detector
// ============================================================
// Detects the user's current Excel display language and maps
// it to a supported FormulaBridge LanguageCode.
// ============================================================

import { LanguageCode } from "../engine/types.js";

const supportedLanguages: LanguageCode[] = [
  "en", "de", "fr", "es", "it", "pt-BR", "nl", "pl", 
  "cs", "ru", "tr", "sv", "nb", "da", "fi", "hu", "sk"
];

/**
 * Gets the current Office application display language and maps it.
 */
export function detectExcelLocale(): LanguageCode {
  // Fallback to English if not available
  if (typeof Office === "undefined" || !Office.context || !Office.context.displayLanguage) {
    return "en";
  }

  const rawLang = Office.context.displayLanguage; // e.g. "de-DE"
  return mapLocaleToSupported(rawLang);
}

/**
 * Maps a locale string to our supported languages list.
 */
export function mapLocaleToSupported(locale: string): LanguageCode {
  // Normalize
  const normalized = locale.trim();
  
  // Exact match (e.g., "pt-BR")
  if (supportedLanguages.includes(normalized as LanguageCode)) {
    return normalized as LanguageCode;
  }

  // Two-letter code match (e.g., "de-DE" -> "de")
  const twoLetter = normalized.split("-")[0].toLowerCase();
  
  if (supportedLanguages.includes(twoLetter as LanguageCode)) {
    return twoLetter as LanguageCode;
  }

  // Special cases for Portuguese
  if (twoLetter === "pt") {
    return "pt-BR"; // Map all Portuguese to pt-BR since it's our only PT dictionary
  }

  // Fallback
  console.warn(`🌉 FormulaBridge: Unsupported locale '${locale}', falling back to 'en'`);
  return "en";
}
