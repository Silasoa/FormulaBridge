// ============================================================
// FormulaBridge – Formula Parser (Tokenizer)
// ============================================================
// Breaks an Excel formula string into semantic tokens.
// Rules:
//   - Strings in "" are never translated
//   - Cell references (A1, $B$2, Sheet1!A1) stay unchanged
//   - Operators (+, -, *, /, ^, &, =, <, >, >=, <=, <>) stay unchanged
//   - Nested parentheses are correctly handled
//   - Function names are identified (word followed by open paren)
// ============================================================

import { Token, TokenType } from "./types.js";

/**
 * Excel formula tokenizer.
 * Converts a raw formula string into an array of typed tokens.
 */
export function tokenize(formula: string): Token[] {
  const tokens: Token[] = [];
  let pos = 0;
  const len = formula.length;

  while (pos < len) {
    const ch = formula[pos];

    // ── Equals sign (formula start) ──
    if (ch === "=" && pos === 0) {
      tokens.push({ type: TokenType.EQUALS, value: "=", position: pos });
      pos++;
      continue;
    }

    // ── Whitespace ──
    if (/\s/.test(ch)) {
      const start = pos;
      while (pos < len && /\s/.test(formula[pos])) pos++;
      tokens.push({
        type: TokenType.WHITESPACE,
        value: formula.slice(start, pos),
        position: start,
      });
      continue;
    }

    // ── String literal ("...") ──
    if (ch === '"') {
      const start = pos;
      pos++; // skip opening quote
      while (pos < len) {
        if (formula[pos] === '"') {
          pos++;
          // Excel uses "" for escaped quotes inside strings
          if (pos < len && formula[pos] === '"') {
            pos++;
            continue;
          }
          break;
        }
        pos++;
      }
      tokens.push({
        type: TokenType.STRING,
        value: formula.slice(start, pos),
        position: start,
      });
      continue;
    }

    // ── Parentheses ──
    if (ch === "(") {
      tokens.push({ type: TokenType.PAREN_OPEN, value: "(", position: pos });
      pos++;
      continue;
    }
    if (ch === ")") {
      tokens.push({ type: TokenType.PAREN_CLOSE, value: ")", position: pos });
      pos++;
      continue;
    }

    // ── Argument separators (, and ;) ──
    if (ch === "," || ch === ";") {
      tokens.push({ type: TokenType.SEPARATOR, value: ch, position: pos });
      pos++;
      continue;
    }

    // ── Colon (range separator A1:B2) ──
    if (ch === ":") {
      tokens.push({ type: TokenType.COLON, value: ":", position: pos });
      pos++;
      continue;
    }

    // ── Multi-char operators (<>, >=, <=) ──
    if (ch === "<" || ch === ">") {
      const start = pos;
      pos++;
      if (pos < len && (formula[pos] === "=" || (ch === "<" && formula[pos] === ">"))) {
        pos++;
      }
      tokens.push({
        type: TokenType.OPERATOR,
        value: formula.slice(start, pos),
        position: start,
      });
      continue;
    }

    // ── Single-char operators ──
    if ("+-*/^&=".includes(ch)) {
      tokens.push({ type: TokenType.OPERATOR, value: ch, position: pos });
      pos++;
      continue;
    }

    // ── Exclamation mark (sheet reference separator like Sheet1!A1) ──
    if (ch === "!") {
      tokens.push({ type: TokenType.OPERATOR, value: "!", position: pos });
      pos++;
      continue;
    }

    // ── Numbers (including decimals with . or ,) ──
    if (/[0-9]/.test(ch)) {
      const start = pos;
      while (pos < len && /[0-9]/.test(formula[pos])) pos++;
      // Handle decimal part (. or , depending on locale)
      if (pos < len && (formula[pos] === "." || formula[pos] === ",")) {
        const nextChar = pos + 1 < len ? formula[pos + 1] : "";
        if (/[0-9]/.test(nextChar)) {
          pos++; // skip decimal separator
          while (pos < len && /[0-9]/.test(formula[pos])) pos++;
        }
      }
      // Handle scientific notation
      if (pos < len && (formula[pos] === "e" || formula[pos] === "E")) {
        pos++;
        if (pos < len && (formula[pos] === "+" || formula[pos] === "-")) pos++;
        while (pos < len && /[0-9]/.test(formula[pos])) pos++;
      }
      // Check if followed by a letter → it's not a pure number but part of a cell ref
      // In that case, we roll back and let the identifier handler deal with it
      if (start === pos || (pos < len && /[A-Za-z_]/.test(formula[pos]))) {
        // Rollback – this might be something like '1A' which isn't valid,
        // but let's not mangle it. Actually pure numbers won't be followed by letters
        // unless in a weird context, so just output as number.
        // But 1:1 (row reference) is fine since : is not a letter.
      }
      tokens.push({
        type: TokenType.NUMBER,
        value: formula.slice(start, pos),
        position: start,
      });
      continue;
    }

    // ── Error literals (#REF!, #VALUE!, #N/A, etc.) ──
    if (ch === "#") {
      const start = pos;
      pos++;
      while (pos < len && /[A-Za-z0-9/]/.test(formula[pos])) pos++;
      if (pos < len && formula[pos] === "!") pos++; // trailing !
      if (pos < len && formula[pos] === "?") pos++; // #NULL? etc.
      tokens.push({
        type: TokenType.ERROR_LITERAL,
        value: formula.slice(start, pos),
        position: start,
      });
      continue;
    }

    // ── Identifiers: function names, cell references, named ranges, booleans ──
    // Identifiers can contain letters (unicode), digits, underscores, dots, and $
    if (/[A-Za-z_$\u00C0-\u024F\u0400-\u04FF]/.test(ch)) {
      const start = pos;
      // Consume the identifier including dots (for BEREICH.VERSCHIEBEN etc.)
      while (
        pos < len &&
        /[A-Za-z0-9_.$\u00C0-\u024F\u0400-\u04FF]/.test(formula[pos])
      ) {
        pos++;
      }
      const word = formula.slice(start, pos);
      const upperWord = word.toUpperCase();

      // Check for boolean literals
      if (upperWord === "TRUE" || upperWord === "FALSE" ||
          upperWord === "WAHR" || upperWord === "FALSCH" ||
          upperWord === "VRAI" || upperWord === "FAUX" ||
          upperWord === "VERDADERO" || upperWord === "FALSO" ||
          upperWord === "VERO") {
        tokens.push({
          type: TokenType.BOOLEAN,
          value: word,
          position: start,
        });
        continue;
      }

      // Determine if this is a function name or a cell/range reference
      // Look ahead: is the next non-whitespace char a '(' ?
      let lookAhead = pos;
      while (lookAhead < len && /\s/.test(formula[lookAhead])) lookAhead++;

      if (lookAhead < len && formula[lookAhead] === "(") {
        tokens.push({
          type: TokenType.FUNCTION,
          value: word,
          position: start,
        });
      } else if (isCellReference(word)) {
        tokens.push({
          type: TokenType.CELL_REF,
          value: word,
          position: start,
        });
      } else {
        // Could be a named range, sheet name, or unknown identifier
        // Treat as CELL_REF to avoid translating it
        tokens.push({
          type: TokenType.CELL_REF,
          value: word,
          position: start,
        });
      }
      continue;
    }

    // ── Percent sign ──
    if (ch === "%") {
      tokens.push({ type: TokenType.OPERATOR, value: "%", position: pos });
      pos++;
      continue;
    }

    // ── Curly braces for array formulas ──
    if (ch === "{" || ch === "}") {
      tokens.push({ type: TokenType.OPERATOR, value: ch, position: pos });
      pos++;
      continue;
    }

    // ── Single quotes (sheet name delimiter like 'Sheet Name'!A1) ──
    if (ch === "'") {
      const start = pos;
      pos++;
      while (pos < len && formula[pos] !== "'") pos++;
      if (pos < len) pos++; // closing quote
      tokens.push({
        type: TokenType.CELL_REF, // treat quoted sheet name as ref part
        value: formula.slice(start, pos),
        position: start,
      });
      continue;
    }

    // ── Anything else → UNKNOWN ──
    tokens.push({ type: TokenType.UNKNOWN, value: ch, position: pos });
    pos++;
  }

  return tokens;
}

/**
 * Check if a string looks like an Excel cell reference.
 * Matches: A1, $A$1, A$1, $A1, AA123, XFD1048576
 * Does NOT match: SUMME, WENN, etc.
 */
function isCellReference(s: string): boolean {
  // Remove leading $ signs for matching
  const cleaned = s.replace(/\$/g, "");
  // Pattern: 1-3 letters followed by 1-7 digits
  return /^[A-Za-z]{1,3}[0-9]{1,7}$/.test(cleaned);
}

/**
 * Reconstruct a formula string from tokens.
 */
export function tokensToString(tokens: Token[]): string {
  return tokens.map((t) => t.value).join("");
}
