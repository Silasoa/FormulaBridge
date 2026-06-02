// ============================================================
// FormulaBridge – Translation Engine
// ============================================================
// Core translation logic that takes a formula string, detects
// (or accepts) a source language, and translates all function
// names to the target language. Also converts separators.
// ============================================================

import { tokenize, tokensToString } from "./parser.js";
import { detectLanguage } from "./autoDetect.js";
import { convertSeparators, getEnglishLocale } from "./separatorConverter.js";
import {
  Token,
  TokenType,
  LanguageCode,
  Dictionary,
  TranslationResult,
  TranslationChange,
  LocaleInfo,
} from "./types.js";

/**
 * Build a lookup map for translating function names between two languages.
 * Returns a Map<string (uppercase source name), string (target name)>.
 */
function buildTranslationMap(
  dictionary: Dictionary,
  sourceLang: LanguageCode,
  targetLang: LanguageCode
): Map<string, string> {
  const map = new Map<string, string>();

  for (const [englishName, translations] of Object.entries(dictionary.functions)) {
    // Get the source name (what the user typed)
    let sourceName: string;
    if (sourceLang === "en") {
      sourceName = englishName;
    } else {
      sourceName = translations[sourceLang];
      if (!sourceName) continue; // no translation available
    }

    // Get the target name (what we want to output)
    let targetName: string;
    if (targetLang === "en") {
      targetName = englishName;
    } else {
      targetName = translations[targetLang];
      if (!targetName) continue; // no translation available
    }

    // Only add if source and target differ
    if (sourceName.toUpperCase() !== targetName.toUpperCase()) {
      map.set(sourceName.toUpperCase(), targetName);
    }
  }

  return map;
}

/**
 * Translate a single formula string.
 *
 * @param formula - The raw formula string (e.g., "=SUMME(A1:A10)")
 * @param dictionary - The loaded dictionary
 * @param targetLang - The target language (usually the Excel installation language)
 * @param sourceLang - Optional source language. If not provided, auto-detection is used.
 * @returns TranslationResult with the translated formula and metadata
 */
export function translateFormula(
  formula: string,
  dictionary: Dictionary,
  targetLang: LanguageCode,
  sourceLang?: LanguageCode
): TranslationResult {
  // Don't process non-formula strings
  if (!formula.startsWith("=")) {
    return {
      original: formula,
      translated: formula,
      wasTranslated: false,
      sourceLanguage: sourceLang || "en",
      targetLanguage: targetLang,
      changes: [],
      confidence: 1,
    };
  }

  // 1. Tokenize the formula
  const tokens = tokenize(formula);

  // 2. Detect source language if not provided
  let detectedLang = sourceLang;
  let confidence = 1;

  if (!detectedLang) {
    const detection = detectLanguage(tokens, dictionary);
    if (!detection || detection.confidence < 0.3) {
      // Can't reliably detect – don't translate
      return {
        original: formula,
        translated: formula,
        wasTranslated: false,
        sourceLanguage: "en",
        targetLanguage: targetLang,
        changes: [],
        confidence: detection?.confidence ?? 0,
      };
    }
    detectedLang = detection.language;
    confidence = detection.confidence;
  }

  // 3. If source and target are the same, no translation needed
  if (detectedLang === targetLang) {
    return {
      original: formula,
      translated: formula,
      wasTranslated: false,
      sourceLanguage: detectedLang,
      targetLanguage: targetLang,
      changes: [],
      confidence,
    };
  }

  // 4. Build translation map
  const translationMap = buildTranslationMap(dictionary, detectedLang, targetLang);

  // 5. Translate function tokens
  const changes: TranslationChange[] = [];
  const translatedTokens: Token[] = tokens.map((token) => {
    if (
      token.type === TokenType.FUNCTION ||
      token.type === TokenType.BOOLEAN ||
      token.type === TokenType.ERROR_LITERAL
    ) {
      const key = token.value.toUpperCase();
      const translated = translationMap.get(key);
      if (translated) {
        changes.push({
          original: token.value,
          translated,
          position: token.position,
        });
        return { ...token, value: translated };
      }
    }
    return token;
  });

  // 6. Convert separators
  const sourceLocale = getLocaleInfo(dictionary, detectedLang);
  const targetLocale = getLocaleInfo(dictionary, targetLang);
  const finalTokens = convertSeparators(translatedTokens, sourceLocale, targetLocale);

  // 7. Reconstruct the formula
  const translated = tokensToString(finalTokens);

  return {
    original: formula,
    translated,
    wasTranslated: changes.length > 0 || sourceLocale.separator !== targetLocale.separator,
    sourceLanguage: detectedLang,
    targetLanguage: targetLang,
    changes,
    confidence,
  };
}

/**
 * Get locale info for a language, with English as fallback.
 */
function getLocaleInfo(dictionary: Dictionary, lang: LanguageCode): LocaleInfo {
  if (lang === "en") return getEnglishLocale();
  return dictionary.locales[lang] || getEnglishLocale();
}

/**
 * Translate all formulas in a 2D array (for batch translation).
 * This matches the shape of Excel's `range.formulas` property.
 */
export function translateFormulas(
  formulas: string[][],
  dictionary: Dictionary,
  targetLang: LanguageCode,
  sourceLang?: LanguageCode
): { results: TranslationResult[][]; totalTranslated: number } {
  let totalTranslated = 0;
  const results = formulas.map((row) =>
    row.map((formula) => {
      const result = translateFormula(formula, dictionary, targetLang, sourceLang);
      if (result.wasTranslated) totalTranslated++;
      return result;
    })
  );
  return { results, totalTranslated };
}
