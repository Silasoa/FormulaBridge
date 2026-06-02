// ============================================================
// FormulaBridge – Core Type Definitions
// ============================================================

/** Supported language codes */
export type LanguageCode =
  | "en"
  | "de"
  | "fr"
  | "es"
  | "it"
  | "pt-BR"
  | "pt-PT"
  | "nl"
  | "pl"
  | "cs"
  | "ru"
  | "tr"
  | "sv"
  | "nb"
  | "da"
  | "fi"
  | "hu"
  | "sk";

/** Token types produced by the formula parser */
export enum TokenType {
  EQUALS = "EQUALS",
  FUNCTION = "FUNCTION",
  CELL_REF = "CELL_REF",
  RANGE_REF = "RANGE_REF",
  NUMBER = "NUMBER",
  STRING = "STRING",
  BOOLEAN = "BOOLEAN",
  ERROR_LITERAL = "ERROR_LITERAL",
  OPERATOR = "OPERATOR",
  SEPARATOR = "SEPARATOR",
  PAREN_OPEN = "PAREN_OPEN",
  PAREN_CLOSE = "PAREN_CLOSE",
  COLON = "COLON",
  WHITESPACE = "WHITESPACE",
  UNKNOWN = "UNKNOWN",
}

/** A single token from the formula parser */
export interface Token {
  type: TokenType;
  value: string;
  /** Start position in the original string (0-indexed) */
  position: number;
}

/** Locale metadata for a language */
export interface LocaleInfo {
  separator: string;
  decimalSeparator: string;
  name: string;
}

/** Dictionary metadata */
export interface DictionaryMeta {
  version: string;
  totalFunctions: number;
  languages: LanguageCode[];
}

/** Complete dictionary structure */
export interface Dictionary {
  meta: DictionaryMeta;
  locales: Record<string, LocaleInfo>;
  functions: Record<string, Record<string, string>>;
}

/** Result of language auto-detection */
export interface DetectionResult {
  language: LanguageCode;
  confidence: number;
  matchedFunctions: string[];
}

/** A single translation change */
export interface TranslationChange {
  original: string;
  translated: string;
  position: number;
}

/** Result of a formula translation */
export interface TranslationResult {
  /** The original formula string */
  original: string;
  /** The translated formula string */
  translated: string;
  /** Whether any translation was actually performed */
  wasTranslated: boolean;
  /** Detected source language (or the one provided) */
  sourceLanguage: LanguageCode;
  /** Target language used */
  targetLanguage: LanguageCode;
  /** Individual changes made */
  changes: TranslationChange[];
  /** Confidence of the auto-detection (1.0 if manually specified) */
  confidence: number;
}

/** Translation log entry for the UI */
export interface LogEntry {
  id: string;
  timestamp: Date;
  cellAddress: string;
  result: TranslationResult;
}

/** Application state */
export interface AppState {
  isActive: boolean;
  autoDetect: boolean;
  autoTranslate: boolean;
  sourceLanguage: LanguageCode | "auto";
  excelLocale: LanguageCode;
  translationLog: LogEntry[];
}
