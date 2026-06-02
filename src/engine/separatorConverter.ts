// ============================================================
// FormulaBridge – Separator Converter
// ============================================================
// Handles conversion between locale-specific separators:
//   - Argument separator: , (EN) ↔ ; (DE, FR, etc.)
//   - Decimal separator: . (EN) ↔ , (DE, FR, etc.)
// ============================================================

import { Token, TokenType, LocaleInfo } from "./types.js";

/**
 * Convert argument separators in a token array from one locale to another.
 * Only SEPARATOR tokens are affected. Numbers with decimal separators
 * are handled via the NUMBER token.
 */
export function convertSeparators(
  tokens: Token[],
  sourceLocale: LocaleInfo,
  targetLocale: LocaleInfo
): Token[] {
  if (sourceLocale.separator === targetLocale.separator) {
    return tokens; // nothing to do
  }

  return tokens.map((token) => {
    if (token.type === TokenType.SEPARATOR) {
      // Convert argument separator
      if (token.value === sourceLocale.separator) {
        return { ...token, value: targetLocale.separator };
      }
    }
    return token;
  });
}

/**
 * Get locale info for English (the internal Excel format).
 */
export function getEnglishLocale(): LocaleInfo {
  return {
    separator: ",",
    decimalSeparator: ".",
    name: "English",
  };
}
