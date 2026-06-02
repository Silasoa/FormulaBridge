// ============================================================
// FormulaBridge – Auto-Detect Engine
// ============================================================
// Automatically detects the source language of a formula by
// analyzing function names against the dictionary.
// ============================================================

import { Token, TokenType, LanguageCode, DetectionResult, Dictionary } from "./types.js";

/** Inverted lookup: local function name → { englishName, language } */
interface InvertedEntry {
  englishName: string;
  language: LanguageCode;
}

/** Cache for inverted lookup maps */
let invertedMap: Map<string, InvertedEntry[]> | null = null;

/**
 * Build an inverted map from the dictionary for fast language detection.
 * Maps each localized function name to its English equivalent + language code.
 */
export function buildInvertedMap(dictionary: Dictionary): Map<string, InvertedEntry[]> {
  if (invertedMap) return invertedMap;

  invertedMap = new Map();

  for (const [englishName, translations] of Object.entries(dictionary.functions)) {
    // Add English itself
    const enKey = englishName.toUpperCase();
    if (!invertedMap.has(enKey)) {
      invertedMap.set(enKey, []);
    }
    invertedMap.get(enKey)!.push({ englishName, language: "en" });

    // Add all translations
    for (const [langCode, localName] of Object.entries(translations)) {
      const key = localName.toUpperCase();
      if (!invertedMap.has(key)) {
        invertedMap.set(key, []);
      }
      invertedMap.get(key)!.push({
        englishName,
        language: langCode as LanguageCode,
      });
    }
  }

  return invertedMap;
}

/**
 * Clear the cached inverted map (useful for testing).
 */
export function clearInvertedMapCache(): void {
  invertedMap = null;
}

/**
 * Detect the source language of a formula based on its function name tokens.
 *
 * Strategy:
 * 1. Extract all FUNCTION tokens from the parsed formula
 * 2. Look up each function name in the inverted map
 * 3. Count matches per language
 * 4. Return the language with the most matches
 * 5. Confidence = matchCount / totalFunctions (0..1)
 *
 * Special cases:
 * - If all functions exist in English, return "en" (no translation needed)
 * - If functions match multiple languages equally, prefer the non-English one
 * - If no functions are found, return null
 */
export function detectLanguage(
  tokens: Token[],
  dictionary: Dictionary
): DetectionResult | null {
  const inverted = buildInvertedMap(dictionary);
  const functionTokens = tokens.filter((t) => t.type === TokenType.FUNCTION);

  if (functionTokens.length === 0) return null;

  // Count matches per language
  const langScores = new Map<LanguageCode, string[]>();

  for (const token of functionTokens) {
    const key = token.value.toUpperCase();
    const entries = inverted.get(key);

    if (entries) {
      for (const entry of entries) {
        if (!langScores.has(entry.language)) {
          langScores.set(entry.language, []);
        }
        langScores.get(entry.language)!.push(token.value);
      }
    }
  }

  if (langScores.size === 0) return null;

  // Find the language with the highest score
  let bestLang: LanguageCode = "en";
  let bestScore = 0;
  let bestMatches: string[] = [];

  for (const [lang, matches] of langScores) {
    // Count unique matches (same function appearing twice shouldn't count double)
    const uniqueMatches = [...new Set(matches)];
    const score = uniqueMatches.length;

    if (score > bestScore) {
      bestScore = score;
      bestLang = lang;
      bestMatches = uniqueMatches;
    } else if (score === bestScore && lang !== "en") {
      // When tied, prefer non-English (since many function names are the same in EN)
      // But only if this language has unique matches that EN doesn't
      const enMatches = langScores.get("en") || [];
      const uniqueToThisLang = uniqueMatches.filter(
        (m) => !enMatches.includes(m)
      );
      if (uniqueToThisLang.length > 0) {
        bestLang = lang;
        bestMatches = uniqueMatches;
      }
    }
  }

  const confidence = bestScore / functionTokens.length;

  return {
    language: bestLang,
    confidence,
    matchedFunctions: bestMatches,
  };
}
