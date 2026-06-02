// ============================================================
// FormulaBridge – Unit Tests: Parser
// ============================================================

import { describe, it, expect } from "vitest";
import { tokenize, tokensToString } from "../../src/engine/parser.js";
import { TokenType } from "../../src/engine/types.js";

describe("Formula Parser (Tokenizer)", () => {
  it("should tokenize a simple SUM formula", () => {
    const tokens = tokenize("=SUM(A1:A10)");
    expect(tokens).toHaveLength(7);
    expect(tokens[0]).toMatchObject({ type: TokenType.EQUALS, value: "=" });
    expect(tokens[1]).toMatchObject({ type: TokenType.FUNCTION, value: "SUM" });
    expect(tokens[2]).toMatchObject({ type: TokenType.PAREN_OPEN, value: "(" });
    expect(tokens[3]).toMatchObject({ type: TokenType.CELL_REF, value: "A1" });
    expect(tokens[4]).toMatchObject({ type: TokenType.COLON, value: ":" });
    expect(tokens[5]).toMatchObject({ type: TokenType.CELL_REF, value: "A10" });
    expect(tokens[6]).toMatchObject({ type: TokenType.PAREN_CLOSE, value: ")" });
  });

  it("should tokenize a German SUMME formula", () => {
    const tokens = tokenize("=SUMME(A1:A10)");
    expect(tokens[1]).toMatchObject({ type: TokenType.FUNCTION, value: "SUMME" });
  });

  it("should tokenize nested functions", () => {
    const tokens = tokenize("=WENN(SUMME(A1:A10)>0;MITTELWERT(B1:B10);0)");
    const functionTokens = tokens.filter((t) => t.type === TokenType.FUNCTION);
    expect(functionTokens).toHaveLength(3);
    expect(functionTokens[0].value).toBe("WENN");
    expect(functionTokens[1].value).toBe("SUMME");
    expect(functionTokens[2].value).toBe("MITTELWERT");
  });

  it("should protect string literals from being parsed as functions", () => {
    const tokens = tokenize('=VERKETTEN("SUMME";A1)');
    const strings = tokens.filter((t) => t.type === TokenType.STRING);
    expect(strings).toHaveLength(1);
    expect(strings[0].value).toBe('"SUMME"');
    // SUMME inside quotes should NOT be a FUNCTION token
    const functions = tokens.filter((t) => t.type === TokenType.FUNCTION);
    expect(functions).toHaveLength(1);
    expect(functions[0].value).toBe("VERKETTEN");
  });

  it("should handle cell references with $ signs", () => {
    const tokens = tokenize("=SUM($A$1:$B$10)");
    const refs = tokens.filter((t) => t.type === TokenType.CELL_REF);
    expect(refs).toHaveLength(2);
    expect(refs[0].value).toBe("$A$1");
    expect(refs[1].value).toBe("$B$10");
  });

  it("should handle comparison operators", () => {
    const tokens = tokenize("=IF(A1>=10;A1;0)");
    const ops = tokens.filter((t) => t.type === TokenType.OPERATOR);
    expect(ops.some((o) => o.value === ">=")).toBe(true);
  });

  it("should handle <> operator", () => {
    const tokens = tokenize("=IF(A1<>0;1;0)");
    const ops = tokens.filter((t) => t.type === TokenType.OPERATOR);
    expect(ops.some((o) => o.value === "<>")).toBe(true);
  });

  it("should tokenize semicolons as separators", () => {
    const tokens = tokenize("=IF(A1>0;A1;0)");
    const seps = tokens.filter((t) => t.type === TokenType.SEPARATOR);
    expect(seps).toHaveLength(2);
    expect(seps[0].value).toBe(";");
  });

  it("should tokenize commas as separators", () => {
    const tokens = tokenize("=IF(A1>0,A1,0)");
    const seps = tokens.filter((t) => t.type === TokenType.SEPARATOR);
    expect(seps).toHaveLength(2);
    expect(seps[0].value).toBe(",");
  });

  it("should handle error literals", () => {
    const tokens = tokenize("=IFERROR(A1/B1;#N/A)");
    const errors = tokens.filter((t) => t.type === TokenType.ERROR_LITERAL);
    expect(errors).toHaveLength(1);
    expect(errors[0].value).toBe("#N/A");
  });

  it("should handle dot-separated function names (BEREICH.VERSCHIEBEN)", () => {
    const tokens = tokenize("=BEREICH.VERSCHIEBEN(A1;0;0)");
    const fns = tokens.filter((t) => t.type === TokenType.FUNCTION);
    expect(fns).toHaveLength(1);
    expect(fns[0].value).toBe("BEREICH.VERSCHIEBEN");
  });

  it("should handle numbers", () => {
    const tokens = tokenize("=A1+100");
    const nums = tokens.filter((t) => t.type === TokenType.NUMBER);
    expect(nums).toHaveLength(1);
    expect(nums[0].value).toBe("100");
  });

  it("should roundtrip a formula through tokenize → tokensToString", () => {
    const formula = '=WENN(SUMME(A1:A10)>0;MITTELWERT(B1:B10);"Kein Wert")';
    const tokens = tokenize(formula);
    const reconstructed = tokensToString(tokens);
    expect(reconstructed).toBe(formula);
  });

  it("should handle sheet references with !", () => {
    const tokens = tokenize("=SUM(Sheet1!A1:A10)");
    expect(tokens[1]).toMatchObject({ type: TokenType.FUNCTION, value: "SUM" });
    // Sheet1 should be a CELL_REF, ! should be an operator, A1 a CELL_REF
    const refs = tokens.filter((t) => t.type === TokenType.CELL_REF);
    expect(refs.length).toBeGreaterThanOrEqual(2);
  });

  it("should handle escaped quotes in strings", () => {
    const tokens = tokenize('=CONCATENATE("He said ""hello""")');
    const strings = tokens.filter((t) => t.type === TokenType.STRING);
    expect(strings).toHaveLength(1);
    expect(strings[0].value).toBe('"He said ""hello"""');
  });

  it("should handle empty formula", () => {
    const tokens = tokenize("");
    expect(tokens).toHaveLength(0);
  });

  it("should handle formula with whitespace", () => {
    const tokens = tokenize("= SUM( A1 : A10 )");
    const fns = tokens.filter((t) => t.type === TokenType.FUNCTION);
    expect(fns).toHaveLength(1);
    expect(fns[0].value).toBe("SUM");
  });
});
